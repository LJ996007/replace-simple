"""
replace-simple GUI 入口。
只保留批量文本替换功能。
"""

import ctypes
import json
import os
import queue
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from tksheet import Sheet

from app_info import format_changelog, format_version_title

# string_replacer 会拖入 openpyxl(~290ms)+ docx(~100ms)，合计约 400ms。
# 改为按需延迟导入：仅在「导入 Excel」或「开始替换」时加载，让窗口瞬间弹出。


DEFAULT_BLANK_RULE_ROWS = 4
SUPPORTED_EXTENSIONS = (".docx", ".xlsx", ".xlsm", ".pptx")
APP_STATE_DIR_NAME = "replace-simple"
ERROR_LOG_NAME = "error.log"
SETTINGS_NAME = "settings.json"

COMPRESSED_SHEET_HEIGHT = 210      # 基准规则表像素高度（默认窗口下约露 6 行 + 表头，且底部按钮可见）
MIN_SHEET_HEIGHT = 120             # 窗口较矮或系统缩放较大时，优先保住底部操作区
TASK_POLL_INTERVAL_MS = 40
FOLDER_SCAN_PROGRESS_INTERVAL = 100

# 网格线配色（清晰可见版）：tksheet 的网格线宽度被硬编码为 1px，无法加粗，
# 只能靠加深颜色让线条在白底上明显可辨。
GRID_COLOR = "#9CA3AF"          # 数据区网格线（中等灰，白底上清晰）
HEADER_GRID_COLOR = "#8A92A0"   # 表头/序号列网格线（略深，浅灰底上清晰）
HEADER_BG = "#EEF1F5"           # 表头/序号列底色
ZEBRA_BG = "#F4F6F9"            # 斑马纹底色


def _app_data_dir(env_name):
    base = os.environ.get(env_name) or os.path.expanduser("~")
    return os.path.join(base, APP_STATE_DIR_NAME)


def error_log_path():
    return os.path.join(_app_data_dir("LOCALAPPDATA"), ERROR_LOG_NAME)


def settings_path():
    return os.path.join(_app_data_dir("APPDATA"), SETTINGS_NAME)


def write_error_log(exc_type, exc_value, exc_tb, context="未捕获异常"):
    path = error_log_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "\n" + "=" * 72 + "\n",
        f"时间：{timestamp}\n",
        f"位置：{context}\n",
        "".join(traceback.format_exception(exc_type, exc_value, exc_tb)),
    ]
    with open(path, "a", encoding="utf-8") as file:
        file.writelines(lines)
    return path


def install_exception_handlers(root):
    def _handle_exception(exc_type, exc_value, exc_tb, context):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        try:
            log_path = write_error_log(exc_type, exc_value, exc_tb, context=context)
        except Exception:
            log_path = error_log_path()
        try:
            messagebox.showerror(
                "程序出错",
                "程序遇到未处理错误，详细信息已写入日志：\n"
                f"{log_path}\n\n"
                "可以把这个日志文件发给维护人员排查。",
                parent=root if root and root.winfo_exists() else None,
            )
        except Exception:
            pass

    sys.excepthook = lambda exc_type, exc_value, exc_tb: _handle_exception(
        exc_type, exc_value, exc_tb, "主线程"
    )
    root.report_callback_exception = lambda exc_type, exc_value, exc_tb: _handle_exception(
        exc_type, exc_value, exc_tb, "Tk 回调"
    )


def load_settings():
    path = settings_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(data):
    path = settings_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)


def resource_path(name):
    """资源文件绝对路径：开发时取源码目录，PyInstaller 打包后取运行时目录(_MEIPASS)。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def normalize_rule_rows(rows):
    """将界面规则行整理为替换规则，跳过空原文并按精确规则对去重。"""
    rules = []
    seen = set()

    for row in rows:
        old_value = row[0] if len(row) > 0 else ""
        new_value = row[1] if len(row) > 1 else ""
        old_text = "" if old_value is None else str(old_value).strip()
        new_text = "" if new_value is None else str(new_value).strip()

        if not old_text:
            continue

        pair = (old_text, new_text)
        if pair in seen:
            continue

        seen.add(pair)
        rules.append(pair)

    return rules


def make_blank_rows(count):
    """生成 count 个相互独立的空规则行。

    不能用 [["", ""]] * count —— 那样得到的是同一个列表对象的多个引用，
    tksheet 直接持有这些行，编辑任一行时（data[r][c] = value）会连带改写
    所有共享引用的行，表现为「输入一行，其余行跟着一起变」。
    """
    return [["", ""] for _ in range(max(count, 0))]


def _file_identity(path):
    """生成用于判断同一文件的规范化路径。"""
    return os.path.normcase(os.path.abspath(os.path.realpath(os.fspath(path))))


def remove_matching_file_paths(file_paths, target_path):
    """从文件列表中移除与 target_path 指向同一文件的路径。"""
    target_identity = _file_identity(target_path)
    remaining = []
    removed = []

    for file_path in file_paths:
        if _file_identity(file_path) == target_identity:
            removed.append(file_path)
        else:
            remaining.append(file_path)

    return remaining, removed


def collect_supported_files(directory, extensions, progress_callback=None):
    """递归收集支持的文件，跳过 Office 创建的临时文件。"""
    collected = []
    scanned_count = 0

    for root_dir, _dirs, files in os.walk(directory):
        for filename in files:
            scanned_count += 1
            if filename.startswith("~$"):
                continue
            file_path = os.path.join(root_dir, filename)
            if os.path.splitext(filename)[1].lower() in extensions:
                collected.append(file_path)
            if progress_callback and scanned_count % FOLDER_SCAN_PROGRESS_INTERVAL == 0:
                progress_callback(scanned_count)

    if progress_callback:
        progress_callback(scanned_count)
    return collected


class BackgroundTaskRunner:
    """通过队列把后台任务结果安全地交回 Tk 主线程。"""

    def __init__(self, widget):
        self.widget = widget
        self._events = queue.Queue()
        self._active = False
        self._on_progress = None
        self._on_success = None
        self._on_error = None

    @property
    def active(self):
        return self._active

    def submit(self, work, on_success, on_error, on_progress=None):
        if self._active:
            return False

        self._active = True
        self._on_progress = on_progress
        self._on_success = on_success
        self._on_error = on_error

        def publish_progress(*args):
            self._events.put(("progress", args))

        def task():
            try:
                result = work(publish_progress)
            except Exception as exc:
                self._events.put(("error", (exc,)))
            else:
                self._events.put(("success", (result,)))

        threading.Thread(target=task, daemon=True).start()
        self._poll_events()
        return True

    def _poll_events(self):
        latest_progress = None
        completion = None

        while True:
            try:
                event_name, args = self._events.get_nowait()
            except queue.Empty:
                break
            if event_name == "progress":
                latest_progress = args
            else:
                completion = (event_name, args)

        if latest_progress and self._on_progress:
            self._on_progress(*latest_progress)

        if completion:
            event_name, args = completion
            on_success = self._on_success
            on_error = self._on_error
            self._active = False
            self._on_progress = None
            self._on_success = None
            self._on_error = None
            if event_name == "success" and on_success:
                on_success(*args)
            elif event_name == "error" and on_error:
                on_error(*args)
            return

        if self._active:
            self.widget.after(TASK_POLL_INTERVAL_MS, self._poll_events)


class ReplaceSimpleApp:
    def __init__(self, root, restore_session=True):
        self.root = root
        self.restore_session = restore_session
        self._task_runner = BackgroundTaskRunner(root)
        self._busy_widgets = []
        self.root.title(format_version_title())
        self.root.geometry("860x720")
        self.root.minsize(800, 660)

        # 窗口标题栏 / 任务栏图标
        self._apply_window_icon()

        self.replace_files = []
        self.recent_word_files = []
        self.output_dir = None
        self._last_output_dir_to_open = None
        self.table_export_window = None

        self._setup_fonts()
        self._create_widgets()
        if self.restore_session:
            self._restore_session()
            self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _apply_window_icon(self):
        """设置窗口图标：用 Tkinter iconphoto 加载大尺寸 PNG，
        让任务栏显示高清图标（iconbitmap 的小图标在任务栏放大后会模糊）。"""
        try:
            icon_png = resource_path("icon-256.png")
            self._icon_photo = tk.PhotoImage(file=icon_png)
            self.root.iconphoto(True, self._icon_photo)
        except Exception:
            pass

    def _setup_fonts(self):
        self.font_family = "Microsoft YaHei UI"
        self.body_font = tkfont.Font(family=self.font_family, size=10)
        self.title_font = tkfont.Font(family=self.font_family, size=12, weight="bold")
        self.section_font = tkfont.Font(family=self.font_family, size=10, weight="bold")
        self.small_font = tkfont.Font(family=self.font_family, size=9)
        self.app_bg = "#f6f8fb"
        self.surface_bg = "#ffffff"
        self.border_color = "#d0d7de"
        self.muted_fg = "#667085"
        self.accent_fg = "#0067c0"
        self.text_fg = "#1F2937"

        self.root.configure(bg=self.app_bg)
        style = ttk.Style()
        self._use_fast_native_theme(style)
        style.configure(".", font=self.body_font)
        style.configure("App.TFrame", background=self.app_bg)
        style.configure(
            "Surface.TFrame", background=self.surface_bg,
            relief="solid", borderwidth=1, bordercolor=self.border_color,
        )
        style.configure("Toolbar.TFrame", background=self.surface_bg)
        style.configure("TLabel", font=self.body_font)
        style.configure("TButton", font=self.body_font, padding=(10, 5))
        style.configure("TEntry", font=self.body_font)
        style.configure("Title.TLabel", font=self.title_font, background=self.app_bg)
        style.configure("Subtitle.TLabel", font=self.small_font, foreground=self.muted_fg, background=self.app_bg)
        style.configure("SectionTitle.TLabel", font=self.section_font, background=self.surface_bg)
        style.configure("Hint.TLabel", font=self.small_font, foreground=self.muted_fg, background=self.surface_bg)
        style.configure("Muted.TLabel", font=self.body_font, foreground=self.muted_fg, background=self.surface_bg)
        style.configure("Status.TLabel", font=self.body_font, foreground=self.accent_fg, background=self.surface_bg)
        style.configure(
            "Scan.Treeview",
            font=self.body_font,
            rowheight=32,
            background=self.surface_bg,
            fieldbackground=self.surface_bg,
            foreground=self.text_fg,
            bordercolor=self.border_color,
            borderwidth=1,
        )
        style.configure(
            "Scan.Treeview.Heading",
            font=self.section_font,
            padding=(8, 8),
            background=HEADER_BG,
            foreground=self.text_fg,
            relief="flat",
            bordercolor=HEADER_GRID_COLOR,
        )
        style.map(
            "Scan.Treeview.Heading",
            background=[("active", "#E2E8F0")],
            foreground=[("active", self.text_fg)],
        )
        style.configure(
            "Scan.Vertical.TScrollbar",
            gripcount=0,
            width=20,
            arrowsize=20,
            background="#DDE5EF",
            darkcolor="#DDE5EF",
            lightcolor="#DDE5EF",
            troughcolor="#F1F4F8",
            bordercolor="#C4CEDA",
            arrowcolor=self.text_fg,
            relief="flat",
        )
        style.map(
            "Scan.Vertical.TScrollbar",
            background=[("pressed", "#C8D3E0"), ("active", "#D3DCE8")],
            arrowcolor=[("disabled", "#9AA4B2")],
        )
        style.configure(
            "Scan.Horizontal.TScrollbar",
            gripcount=0,
            width=20,
            arrowsize=20,
            background="#DDE5EF",
            darkcolor="#DDE5EF",
            lightcolor="#DDE5EF",
            troughcolor="#F1F4F8",
            bordercolor="#C4CEDA",
            arrowcolor=self.text_fg,
            relief="flat",
        )
        style.map(
            "Scan.Horizontal.TScrollbar",
            background=[("pressed", "#C8D3E0"), ("active", "#D3DCE8")],
            arrowcolor=[("disabled", "#9AA4B2")],
        )

        # 扁平次级按钮（clam 主题下方可定制 background/relief）
        style.configure(
            "TButton", font=self.body_font, padding=(12, 7), relief="flat",
            background="#E7ECF2", foreground=self.text_fg,
            borderwidth=1, bordercolor="#C2CAD5", focuscolor=self.surface_bg,
        )
        style.map(
            "TButton",
            background=[("pressed", "#D7DEE6"), ("active", "#E2E8F0"), ("disabled", "#F3F5F8")],
            bordercolor=[("active", "#B8C0CC")],
            foreground=[("disabled", "#9AA4B2")],
        )
        # 主操作按钮（开始替换）：蓝底白字，一眼可见的主行动点
        style.configure(
            "Accent.TButton", font=self.body_font, padding=(14, 8), relief="flat",
            background=self.accent_fg, foreground="#FFFFFF",
            borderwidth=0, focuscolor=self.accent_fg,
        )
        style.map(
            "Accent.TButton",
            background=[("pressed", "#004C90"), ("active", "#0058A8"), ("disabled", "#9DBBD8")],
            foreground=[("disabled", "#EFF4FA")],
        )
        # 进度条配色（与强调色一致）
        style.configure(
            "Horizontal.TProgressbar", troughcolor="#E6EAF0",
            background=self.accent_fg, borderwidth=0, thickness=8,
        )

        self.root.option_add("*Font", self.body_font)

    def _use_fast_native_theme(self, style):
        # 选 clam 而非 winnative/vista：只有 clam（及 alt/default）允许自定义按钮的
        # background/relief，才能做扁平按钮与主按钮强调色；winnative/vista 会忽略
        # 按钮底色配置，永远是系统灰按钮。clam 同为内置主题，启动开销可忽略。
        for theme_name in ("clam", "alt", "default"):
            if theme_name in style.theme_names():
                try:
                    style.theme_use(theme_name)
                    return
                except tk.TclError:
                    continue

    def _create_busy_button(self, parent, **kwargs):
        button = ttk.Button(parent, **kwargs)
        self._busy_widgets.append(button)
        return button

    def _set_busy_controls(self, busy):
        state = "disabled" if busy else "normal"
        for widget in self._busy_widgets:
            try:
                if widget.winfo_exists():
                    widget.config(state=state)
            except tk.TclError:
                continue

    def _begin_indeterminate_task(self, status):
        self._set_busy_controls(True)
        self.progress.configure(mode="indeterminate", value=0)
        self.progress.grid()
        self.progress.start(12)
        self.status_var.set(status)

    def _begin_determinate_task(self, status, maximum):
        self._set_busy_controls(True)
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=maximum, value=0)
        self.progress.grid()
        self.status_var.set(status)

    def _handle_background_error(self, exc, context):
        log_path = write_error_log(type(exc), exc, exc.__traceback__, context=context)
        self._finish_with_error(str(exc), log_path)

    def _create_widgets(self):
        self.container = ttk.Frame(self.root, padding=(16, 12), style="App.TFrame")
        self.container.pack(fill="both", expand=True)
        self.container.columnconfigure(0, weight=1)

        self.header_frame = ttk.Frame(self.container, style="App.TFrame")
        header_frame = self.header_frame
        header_frame.pack(fill="x", pady=(0, 10))

        header_top = ttk.Frame(header_frame, style="App.TFrame")
        header_top.pack(fill="x")
        ttk.Label(header_top, text="批量文本替换", style="Title.TLabel").pack(side="left", anchor="w")
        self.table_export_button = self._create_busy_button(
            header_top,
            text="提取 Word 表格",
            command=self.open_word_table_exporter,
            width=14,
        )
        self.version_info_button = ttk.Button(
            header_top,
            text="版本",
            command=self.show_version_info,
            width=6,
        )
        self.version_info_button.pack(side="right")
        self.table_export_button.pack(side="right", padx=(0, 6))

        ttk.Label(
            header_frame,
            text="Word / Excel / PPT 文件内容与文件名同步替换",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        self.file_frame = ttk.Frame(self.container, padding=(12, 9), style="Surface.TFrame")
        file_frame = self.file_frame
        file_frame.pack(fill="x", pady=(0, 8))
        file_frame.columnconfigure(0, weight=1)
        ttk.Label(file_frame, text="1  选择待处理文件", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")

        file_toolbar = ttk.Frame(file_frame, style="Toolbar.TFrame")
        file_toolbar.grid(row=0, column=1, sticky="e")
        self._create_busy_button(file_toolbar, text="添加文件", command=self.select_files, width=10).pack(side="left", padx=(0, 6))
        self._create_busy_button(file_toolbar, text="添加文件夹", command=self.select_folder, width=11).pack(side="left", padx=(0, 6))
        self._create_busy_button(file_toolbar, text="移除选中", command=self.remove_selected_files, width=10).pack(side="left", padx=(0, 6))
        self._create_busy_button(file_toolbar, text="清空", command=self.clear_files, width=7).pack(side="left")

        file_list_frame = ttk.Frame(file_frame, style="Toolbar.TFrame")
        file_list_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(7, 0))
        file_list_frame.columnconfigure(0, weight=1)
        self.file_listbox = tk.Listbox(
            file_list_frame,
            height=4,
            selectmode="extended",
            exportselection=False,
            font=self.small_font,
            bg=self.surface_bg,
            fg=self.text_fg,
            highlightthickness=1,
            highlightbackground=self.border_color,
            relief="flat",
        )
        self.file_listbox.grid(row=0, column=0, sticky="ew")
        file_scrollbar = ttk.Scrollbar(file_list_frame, orient="vertical", command=self.file_listbox.yview)
        file_scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=file_scrollbar.set)

        self.files_label = ttk.Label(file_frame, text="已选择 0 个文件", style="Muted.TLabel")
        self.files_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(5, 0))

        self.rules_frame = ttk.Frame(self.container, padding=(12, 9), style="Surface.TFrame")
        rules_frame = self.rules_frame
        # 注意：这里不设 expand=True、不设 row weight —— 规则表高度由像素动态控制
        # （_resize_rules_sheet），避免在默认窗口高度下 tksheet 的最小请求高度把底部
        # 「输出目录 / 开始替换」按钮挤出可视区。
        rules_frame.pack(fill="x", pady=(0, 8))
        rules_frame.columnconfigure(0, weight=1)

        self.rules_header = ttk.Frame(rules_frame, style="Toolbar.TFrame")
        rules_header = self.rules_header
        rules_header.grid(row=0, column=0, sticky="ew")
        rules_header.columnconfigure(0, weight=1)
        ttk.Label(rules_header, text="2  编辑替换规则", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")

        rules_toolbar = ttk.Frame(rules_header, style="Toolbar.TFrame")
        rules_toolbar.grid(row=0, column=1, sticky="e")
        self._create_busy_button(rules_toolbar, text="新增一行", command=self.add_rule_row, width=9).pack(side="left", padx=(0, 6))
        self._create_busy_button(rules_toolbar, text="删除选中", command=self.delete_selected_rules, width=9).pack(side="left", padx=(0, 6))
        self._create_busy_button(rules_toolbar, text="清空规则", command=self.clear_rules, width=9).pack(side="left", padx=(0, 6))
        self._create_busy_button(rules_toolbar, text="导入 Excel", command=self.import_rules_from_excel, width=10).pack(side="left", padx=(0, 6))
        self._create_busy_button(rules_toolbar, text="导入招标文件", command=self.import_rules_from_tender_file, width=13).pack(side="left", padx=(0, 6))
        self._create_busy_button(rules_toolbar, text="导出 Excel", command=self.export_rules_to_excel, width=10).pack(side="left")

        self.rules_hint = ttk.Label(
            rules_frame,
            text="双击单元格可编辑；可从 Excel 或招标 Word 提取规则；执行时按原文长度长词优先。",
            style="Hint.TLabel",
        )
        self.rules_hint.grid(row=1, column=0, sticky="w", pady=(4, 6))

        # 规则表 —— 用 tksheet 取代 ttk.Treeview：
        # ttk.Treeview 在 Windows 上画不出单元格网格线（Tk/Ttk 的硬限制），
        # 而 tksheet 是纯 Python 的电子表格控件，原生支持网格线、单元格
        # 编辑、多选、列宽拖动，体验接近 Excel。
        self.rules_sheet = Sheet(
            rules_frame,
            data=make_blank_rows(DEFAULT_BLANK_RULE_ROWS),
            headers=["原文本", "替换后文本"],
            align="w",
            header_align="center",
            default_column_width=320,
            default_row_height="1",
            height=COMPRESSED_SHEET_HEIGHT,   # 初始压缩高度，由 _resize_rules_sheet 动态调整
            font=(self.font_family, 10, "normal"),
            header_font=(self.font_family, 10, "bold"),
            paste_can_expand_x=False,
            paste_can_expand_y=True,
        )
        self.rules_sheet.grid(row=2, column=0, sticky="new")
        # 规则表只有两列且始终自适应填满，不需要水平滚动；隐藏 tksheet 默认
        # 常显的横向滚动条，避免它占位、显得多余。
        self.rules_sheet.hide("x_scrollbar")

        # 只开启需要的交互：单选/多选、列宽拖动、键盘导航、复制粘贴撤销、单元格编辑。
        # 关闭序号/表头编辑、行拖拽、排序（避免打乱规则顺序）和右键菜单（用顶部按钮代替）。
        self.rules_sheet.enable_bindings(
            "single_select",
            "drag_select",
            "ctrl_click_select",
            "column_width_resize",
            "row_height_resize",
            "arrowkeys",
            "copy", "paste", "delete", "undo",
            "edit_cell",
            menu=False,
        )

        # 配色：网格线清晰 + 浅色斑马纹，与浅色主题协调。
        self.rules_sheet.set_options(
            table_grid_fg=GRID_COLOR,
            table_bg=self.surface_bg,
            table_fg=self.text_fg,
            alternate_color=ZEBRA_BG,
            header_bg=HEADER_BG,
            header_fg=self.text_fg,
            header_grid_fg=HEADER_GRID_COLOR,
            header_selected_columns_bg="#DCE6F5",
            index_bg=HEADER_BG,
            index_fg=self.muted_fg,
            index_grid_fg=HEADER_GRID_COLOR,
            table_selected_cells_bg="#CFE2F7",
            table_selected_cells_fg=self.text_fg,
            top_left_bg=HEADER_BG,
        )

        # 编辑、粘贴、删除、撤销后刷新底部状态计数
        for binding in ("end_edit_cell", "end_paste", "end_delete", "end_undo"):
            self.rules_sheet.extra_bindings(binding, func=self._on_rules_changed)

        # 列宽自动撑满：界面宽度变化时两列等分（减去左侧 index 列），让表格贴合面板宽度。
        self.rules_sheet.bind("<Configure>", self._fit_columns)

        self.bottom_frame = ttk.Frame(self.container, padding=(12, 9), style="Surface.TFrame")
        bottom_frame = self.bottom_frame
        bottom_frame.pack(fill="x")
        bottom_frame.columnconfigure(0, weight=1)

        # 标题行：左标题 + 右「选择目录」按钮
        output_header = ttk.Frame(bottom_frame, style="Toolbar.TFrame")
        output_header.grid(row=0, column=0, sticky="ew")
        output_header.columnconfigure(0, weight=1)
        ttk.Label(output_header, text="3  输出目录", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        self._create_busy_button(output_header, text="选择目录", command=self.select_output_dir, width=12).grid(row=0, column=1, sticky="e")

        # 输出路径/说明单独占整行，避免被右侧按钮挤掉；超长路径自动换行
        self.output_label = ttk.Label(
            bottom_frame, text="未选择则输出到原文件目录；文件名按规则同步替换",
            style="Muted.TLabel", wraplength=820, justify="left",
        )
        self.output_label.grid(row=1, column=0, sticky="w", pady=(5, 0))

        ttk.Separator(bottom_frame, orient="horizontal").grid(row=2, column=0, sticky="ew", pady=8)

        # 状态行：状态文字 + 进度条（空闲隐藏）+ 主按钮
        action_row = ttk.Frame(bottom_frame, style="Toolbar.TFrame")
        action_row.grid(row=3, column=0, sticky="ew")
        action_row.columnconfigure(1, weight=1)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(action_row, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(action_row, mode="determinate", length=220)
        self.progress.grid(row=0, column=1, sticky="ew", padx=(14, 12))
        self.progress.grid_remove()   # 空闲时隐藏，避免显示成一根灰槽；开始替换时再显示
        self.start_button = self._create_busy_button(
            action_row, text="开始替换", command=self.start_replace, width=14, style="Accent.TButton",
        )
        self.start_button.grid(row=0, column=2, sticky="e")

        ttk.Label(
            bottom_frame,
            text="规则表第一列为原文、第二列为替换文；招标文件导入会生成 {项目名称} 等占位符规则。",
            style="Hint.TLabel",
        ).grid(row=4, column=0, sticky="w", pady=(7, 0))

        # 监听窗口高度变化，动态调整规则表高度：规则表吃掉剩余高度，
        # 底部「输出目录 / 开始替换」区域始终完整保留在可视区。
        self._rules_resize_after = None
        self.root.bind("<Configure>", self._on_window_configure)

        # tksheet 在 Windows / 高 DPI 下偶尔会在首帧拿到不完整的内部画布尺寸，
        # 导致表头和网格暂时不绘制，用户点击后才恢复。窗口完成布局后主动刷新几次，
        # 让规则表打开软件时就稳定显示表头和默认空行。
        self.root.after_idle(self._refresh_rules_sheet_view)
        self.root.after(80, self._refresh_rules_sheet_view)
        self.root.after(250, self._refresh_rules_sheet_view)

    # ---------- 列宽自适应 ----------
    def _on_window_configure(self, event=None):
        """窗口尺寸变化时节流刷新规则表高度。"""
        if event is not None and event.widget is not self.root:
            return  # 只处理 root 自身的 Configure，忽略子组件的
        if self._rules_resize_after is not None:
            try:
                self.root.after_cancel(self._rules_resize_after)
            except Exception:
                pass
        self._rules_resize_after = self.root.after(16, self._resize_rules_sheet)

    def _resize_rules_sheet(self):
        """根据当前窗口实际剩余高度动态设定规则表高度。

        旧逻辑用「窗口高度 - 默认高度」简单放大规则表；在 Windows 缩放、
        恢复上次窗口尺寸或主题控件请求高度变化时，底部操作区可能被挤出窗口。
        这里改为先测量顶部、文件区、底部区和规则区固定控件，再把剩余高度给表格。
        """
        self._rules_resize_after = None
        try:
            self.root.update_idletasks()
        except Exception:
            pass

        win_h = self.root.winfo_height()
        if win_h <= 1:
            return

        try:
            container_pad_y = sum(self.container.pack_info().get("pady", (0, 0)))
        except Exception:
            container_pad_y = 0

        fixed_h = (
            self.header_frame.winfo_reqheight() + 10
            + self.file_frame.winfo_reqheight() + 8
            + self.bottom_frame.winfo_reqheight()
            + self._rules_fixed_height()
            + 24
            + container_pad_y
        )
        available = win_h - fixed_h
        new_h = max(MIN_SHEET_HEIGHT, available)

        if new_h != self._current_sheet_height():
            self.rules_sheet.config(height=new_h)
            self._refresh_rules_sheet_view()

    def _rules_fixed_height(self):
        """规则面板中除表格外的固定高度。"""
        try:
            padding_y = self._vertical_padding(self.rules_frame.cget("padding"))
        except Exception:
            padding_y = 18
        return (
            self.rules_header.winfo_reqheight()
            + self.rules_hint.winfo_reqheight()
            + 18
            + padding_y
        )

    def _vertical_padding(self, padding):
        if isinstance(padding, str):
            parts = [int(float(part)) for part in padding.split()]
        elif isinstance(padding, (tuple, list)):
            parts = [int(float(part)) for part in padding]
        else:
            return 0

        if len(parts) == 1:
            return parts[0] * 2
        if len(parts) == 2:
            return parts[1] * 2
        if len(parts) >= 4:
            return parts[1] + parts[3]
        return 0

    def _current_sheet_height(self):
        try:
            h = self.rules_sheet.cget("height")
            if isinstance(h, int):
                return h
            return int(str(h))
        except Exception:
            return COMPRESSED_SHEET_HEIGHT

    def _fit_columns(self, event=None, redraw=True):
        """两列等分撑满表格可视宽度，随界面宽度自动适配。

        以 MT（表格画布）的真实宽度为准来等分——它已不含左侧序号列和右侧
        垂直滚动条，省去对各部件宽度的脆弱估算；横向滚动条已隐藏，故两列之和
        只要不超过该宽度即可完全填满且不被裁剪。"""
        sheet = self.rules_sheet
        try:
            table_w = sheet.MT.winfo_width()
        except Exception:
            table_w = sheet.winfo_width()
        if table_w <= 1:
            return
        each = max((table_w - 4) // 2, 120)
        changed = False
        for c in (0, 1):
            try:
                current = int(sheet.column_width(column=c))
            except Exception:
                current = None
            if current != each:
                sheet.column_width(column=c, width=each, redraw=False)
                changed = True
        if changed and redraw:
            self._redraw_rules_sheet()

    def _redraw_rules_sheet(self):
        """兼容不同 tksheet 版本的主动重绘。"""
        try:
            self.rules_sheet.refresh(redraw_header=True, redraw_row_index=True)
        except Exception:
            try:
                self.rules_sheet.redraw(redraw_header=True, redraw_row_index=True)
            except Exception:
                pass

    def _refresh_rules_sheet_view(self):
        """在 Tk 布局稳定后刷新规则表尺寸、列宽和绘制状态。"""
        try:
            self.root.update_idletasks()
        except Exception:
            pass
        self._fit_columns(redraw=False)
        self._redraw_rules_sheet()

    def _on_rules_changed(self, event=None):
        self._refresh_status()
        self.root.after_idle(self._refresh_rules_sheet_view)

    # ---------- 会话状态 ----------
    def _restore_session(self):
        settings = load_settings()

        geometry = settings.get("geometry")
        if isinstance(geometry, str) and geometry:
            try:
                self.root.geometry(geometry)
            except Exception:
                pass

        rules = settings.get("rules")
        if isinstance(rules, list) and rules:
            normalized = normalize_rule_rows(rules)
            if normalized:
                self.rules_sheet.set_sheet_data([[old, new] for old, new in normalized])
                self._ensure_blank_rule_rows()
                self._refresh_rules_sheet_view()
                self._refresh_status()

        output_dir = settings.get("output_dir")
        if isinstance(output_dir, str) and output_dir:
            self.output_dir = output_dir
            self.output_label.config(text=output_dir, foreground="green")

    def _save_session(self):
        if not self.restore_session:
            return
        data = {
            "geometry": self.root.geometry(),
            "output_dir": self.output_dir,
            "rules": self.get_rules_from_table(),
        }
        save_settings(data)

    def on_close(self):
        table_exporter_busy = (
            self.table_export_window
            and self.table_export_window.exists()
            and self.table_export_window._task_runner.active
        )
        if self._task_runner.active or table_exporter_busy:
            messagebox.showinfo("任务进行中", "当前任务仍在读写文件，请等待完成后再关闭程序。")
            return
        try:
            try:
                self._save_session()
            except Exception:
                pass
        finally:
            self.root.destroy()

    # ---------- 规则表数据操作 ----------
    def _refresh_status(self):
        count = len(self.get_rules_from_table())
        self.status_var.set(f"当前规则 {count} 条" if count else "就绪")

    def _update_files_label(self):
        count = len(self.replace_files)
        foreground = "green" if count else self.muted_fg
        self.files_label.config(text=f"已选择 {count} 个文件", foreground=foreground)

    def _is_docx_file(self, file_path):
        return (
            os.path.splitext(file_path)[1].lower() == ".docx"
            and not os.path.basename(file_path).startswith("~$")
        )

    def _remember_word_files(self, file_paths):
        existing = {_file_identity(path) for path in self.recent_word_files}
        changed = False
        for file_path in file_paths:
            if not self._is_docx_file(file_path):
                continue
            identity = _file_identity(file_path)
            if identity in existing:
                continue
            self.recent_word_files.append(file_path)
            existing.add(identity)
            changed = True
        return changed

    def _selected_replace_word_files(self):
        selected = list(self.file_listbox.curselection())
        if not selected:
            return []
        return [
            self.replace_files[index]
            for index in selected
            if 0 <= index < len(self.replace_files) and self._is_docx_file(self.replace_files[index])
        ]

    def _table_export_word_files(self, selected_only=False):
        window = self.table_export_window
        if not window or not window.exists():
            return []
        if selected_only:
            selected = list(window.file_listbox.curselection())
            return [
                window.file_paths[index]
                for index in selected
                if 0 <= index < len(window.file_paths) and self._is_docx_file(window.file_paths[index])
            ]
        return [path for path in window.file_paths if self._is_docx_file(path)]

    def _dedupe_existing_word_files(self, file_paths):
        seen = set()
        result = []
        for file_path in file_paths:
            if not self._is_docx_file(file_path):
                continue
            identity = _file_identity(file_path)
            if identity in seen:
                continue
            seen.add(identity)
            result.append(file_path)
        return result

    def _word_files_for_reuse(self):
        return self._dedupe_existing_word_files([
            *self._selected_replace_word_files(),
            *self._table_export_word_files(selected_only=True),
            *self.recent_word_files,
            *self._table_export_word_files(selected_only=False),
        ])

    def _choose_tender_file_for_import(self):
        candidates = self._word_files_for_reuse()
        if len(candidates) == 1:
            return candidates[0]

        if len(candidates) > 1:
            messagebox.showinfo(
                "选择招标文件",
                "已检测到多个 Word 文件。\n\n"
                "请在主文件列表或「提取 Word 表格」窗口中选中一个招标文件，"
                "或者在接下来的窗口中重新选择。",
            )

        return filedialog.askopenfilename(
            title="从招标文件读取项目信息",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )

    def _remove_rules_file_from_replace_files(self, file_path):
        return self._remove_import_source_from_replace_files(file_path, "导入源文件")

    def _remove_import_source_from_replace_files(self, file_path, source_label):
        remaining, removed = remove_matching_file_paths(self.replace_files, file_path)
        if not removed:
            return False

        self.replace_files = remaining
        self._refresh_file_list()
        messagebox.showinfo(
            "提示",
            f"{source_label}已从待处理文件列表中自动移除，避免被一起替换：\n"
            f"{os.path.basename(file_path)}",
        )
        return True

    def _ensure_blank_rule_rows(self, minimum=DEFAULT_BLANK_RULE_ROWS):
        """保证至少有 minimum 行空白行可编辑。"""
        current = self.rules_sheet.total_rows()
        if current < minimum:
            new_rows = make_blank_rows(minimum - current)
            self.rules_sheet.insert_rows(rows=new_rows)
            self._refresh_rules_sheet_view()

    def get_rules_from_table(self):
        """从表格读取并归一化规则（跳过空原文、按精确规则对去重）。"""
        rows = self.rules_sheet.get_sheet_data()
        return normalize_rule_rows(rows)

    def add_rule_row(self):
        self.rules_sheet.insert_row(row=["", ""])
        last = max(self.rules_sheet.total_rows() - 1, 0)
        try:
            self.rules_sheet.see(row=last, column=0)
            self.rules_sheet.select_row(last)
        except Exception:
            pass
        self._refresh_rules_sheet_view()
        self._refresh_status()

    def delete_selected_rules(self):
        # get_selected_rows() 默认只返回「点行号选中的整行」；用户在单元格里
        # 单击或双击编辑属于 cell 选择，必须用 get_cells_as_rows=True 才能把
        # 选中/正在编辑单元格所在的行号一并取回，否则「删除选中」始终拿到空集。
        rows = self.rules_sheet.get_selected_rows(get_cells_as_rows=True)
        if not rows:
            return
        # 从大到小删，避免索引偏移
        for r in sorted(rows, reverse=True):
            self.rules_sheet.delete_row(r)
        self._ensure_blank_rule_rows()
        self._refresh_rules_sheet_view()
        self._refresh_status()

    def clear_rules(self):
        self.rules_sheet.set_sheet_data(make_blank_rows(DEFAULT_BLANK_RULE_ROWS))
        self._refresh_rules_sheet_view()
        self._refresh_status()
        self.status_var.set("规则已清空")

    def import_rules_from_excel(self):
        file_path = filedialog.askopenfilename(
            title="导入替换规则",
            filetypes=[("Excel 工作簿", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")],
        )
        if not file_path:
            return

        self._begin_indeterminate_task("正在读取 Excel 规则表...")

        def work(_publish_progress):
            from string_replacer import load_replacement_rules

            return load_replacement_rules(file_path)

        def on_success(rules):
            self._reset_busy_state()
            if not rules:
                messagebox.showwarning("提示", "规则表为空，请检查 Excel 文件。")
                return

            data = [
                ["" if old is None else str(old), "" if new is None else str(new)]
                for old, new in rules
            ]
            self.rules_sheet.set_sheet_data(data)
            self._ensure_blank_rule_rows()
            self._refresh_rules_sheet_view()
            removed_from_targets = self._remove_rules_file_from_replace_files(file_path)
            status = f"已导入 {len(rules)} 条规则：{os.path.basename(file_path)}"
            if removed_from_targets:
                status += "；已从待处理文件中移除规则表"
            self.status_var.set(status)

        def on_error(exc):
            self._reset_busy_state()
            log_path = write_error_log(type(exc), exc, exc.__traceback__, context="导入 Excel 规则")
            messagebox.showerror("错误", f"导入规则失败：{exc}\n\n错误日志：{log_path}")

        self._task_runner.submit(work, on_success, on_error)

    def import_rules_from_tender_file(self):
        file_path = self._choose_tender_file_for_import()
        if not file_path:
            return

        self._begin_indeterminate_task("正在读取招标文件项目信息...")

        def work(_publish_progress):
            from tender_info_extractor import extract_project_info_rules

            return extract_project_info_rules(file_path)

        def on_success(rules):
            self._reset_busy_state()
            if not rules:
                messagebox.showwarning(
                    "提示",
                    "未识别到可导入的项目信息。\n\n"
                    "目前支持常见写法，例如：项目名称、项目编号、采购人、预算金额等字段。",
                )
                return

            data = [[old, new] for old, new in rules]
            self.rules_sheet.set_sheet_data(data)
            self._ensure_blank_rule_rows()
            self._refresh_rules_sheet_view()
            self._remember_word_files([file_path])
            removed_from_targets = self._remove_import_source_from_replace_files(file_path, "招标文件")
            status = f"已从招标文件导入 {len(rules)} 条项目信息：{os.path.basename(file_path)}"
            if removed_from_targets:
                status += "；已从待处理文件中移除招标文件"
            self.status_var.set(status)

        def on_error(exc):
            self._reset_busy_state()
            log_path = write_error_log(type(exc), exc, exc.__traceback__, context="读取招标文件")
            messagebox.showerror("错误", f"读取招标文件失败：{exc}\n\n错误日志：{log_path}")

        self._task_runner.submit(work, on_success, on_error)

    def export_rules_to_excel(self):
        rules = self.get_rules_from_table()
        if not rules:
            messagebox.showwarning("提示", "当前没有可导出的规则。")
            return

        file_path = filedialog.asksaveasfilename(
            title="导出替换规则",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            initialfile="替换规则.xlsx",
        )
        if not file_path:
            return

        self._begin_indeterminate_task("正在导出 Excel 规则表...")

        def work(_publish_progress):
            from openpyxl import Workbook

            workbook = Workbook()
            try:
                worksheet = workbook.active
                worksheet.title = "替换规则"
                for row_index, (old_text, new_text) in enumerate(rules, start=1):
                    worksheet.cell(row=row_index, column=1, value=old_text)
                    worksheet.cell(row=row_index, column=2, value=new_text)
                worksheet.column_dimensions["A"].width = 32
                worksheet.column_dimensions["B"].width = 32
                workbook.save(file_path)
            finally:
                workbook.close()
            return len(rules)

        def on_success(rule_count):
            self._reset_busy_state()
            self.status_var.set(f"已导出 {rule_count} 条规则：{os.path.basename(file_path)}")

        def on_error(exc):
            self._reset_busy_state()
            log_path = write_error_log(type(exc), exc, exc.__traceback__, context="导出 Excel 规则")
            detail = "导出失败：请先关闭正在打开的规则表文件后重试。" if isinstance(exc, PermissionError) else f"导出失败：{exc}"
            messagebox.showerror("错误", f"{detail}\n\n错误日志：{log_path}")

        self._task_runner.submit(work, on_success, on_error)

    # ---------- 文件 / 输出目录 ----------
    def _is_supported_file(self, file_path):
        return (
            not os.path.basename(file_path).startswith("~$")
            and os.path.splitext(file_path)[1].lower() in SUPPORTED_EXTENSIONS
        )

    def _refresh_file_list(self):
        self.file_listbox.delete(0, "end")
        if self.replace_files:
            self.file_listbox.insert("end", *self.replace_files)
        self._update_files_label()

    def _append_replace_files(self, file_paths):
        existing = {_file_identity(path) for path in self.replace_files}
        added = 0
        skipped = 0
        added_paths = []

        for file_path in file_paths:
            if not self._is_supported_file(file_path):
                skipped += 1
                continue
            identity = _file_identity(file_path)
            if identity in existing:
                skipped += 1
                continue
            self.replace_files.append(file_path)
            existing.add(identity)
            added_paths.append(file_path)
            added += 1

        self._remember_word_files(added_paths)
        self._refresh_file_list()
        if added:
            message = f"已添加 {added} 个待处理文件"
            if skipped:
                message += f"；跳过 {skipped} 个重复或不支持的文件"
            self.status_var.set(message)
        elif skipped:
            self.status_var.set(f"未添加新文件；跳过 {skipped} 个重复或不支持的文件")

    def select_files(self):
        file_paths = filedialog.askopenfilenames(
            title="选择待处理文件",
            filetypes=[
                ("Office 文件", "*.docx;*.xlsx;*.xlsm;*.pptx"),
                ("Word 文档", "*.docx"),
                ("Excel 工作簿", "*.xlsx;*.xlsm"),
                ("PowerPoint 演示", "*.pptx"),
                ("所有文件", "*.*"),
            ],
        )
        if not file_paths:
            return

        self._append_replace_files(file_paths)

    def select_folder(self):
        directory = filedialog.askdirectory(title="选择包含待处理文件的文件夹")
        if not directory:
            return

        self._begin_indeterminate_task("正在扫描文件夹中的 Office 文件...")

        def work(publish_progress):
            return collect_supported_files(directory, SUPPORTED_EXTENSIONS, publish_progress)

        def on_progress(scanned_count):
            self.status_var.set(f"正在扫描文件夹：已检查 {scanned_count} 个文件...")

        def on_success(collected):
            self._reset_busy_state()
            if not collected:
                messagebox.showinfo("提示", "该文件夹中未找到支持的 Office 文件。")
                return
            self._append_replace_files(collected)

        def on_error(exc):
            self._handle_background_error(exc, "扫描待处理文件夹")

        self._task_runner.submit(work, on_success, on_error, on_progress)

    def remove_selected_files(self):
        selected = list(self.file_listbox.curselection())
        if not selected:
            return

        for index in reversed(selected):
            del self.replace_files[index]
        self._refresh_file_list()
        self.status_var.set(f"已移除 {len(selected)} 个文件")

    def clear_files(self):
        if not self.replace_files:
            return
        self.replace_files = []
        self._refresh_file_list()
        self.status_var.set("待处理文件已清空")

    def select_output_dir(self):
        directory = filedialog.askdirectory(title="选择输出目录")
        if not directory:
            return

        self.output_dir = directory
        self.output_label.config(text=directory, foreground="green")
        self.status_var.set(f"已选择输出目录：{directory}")

    # ---------- 替换执行 ----------
    def start_replace(self):
        if not self.replace_files:
            messagebox.showwarning("提示", "请先选择待处理文件。")
            return

        file_paths = list(self.replace_files)
        rules = list(self.get_rules_from_table())
        output_dir = self.output_dir
        if not rules:
            messagebox.showwarning("提示", "请先新增或导入至少一条替换规则。")
            return

        self._begin_determinate_task(f"正在使用 {len(rules)} 条规则替换...", len(file_paths))

        def work(publish_progress):
            from string_replacer import batch_replace

            return batch_replace(
                file_paths,
                rules,
                output_dir=output_dir,
                progress_callback=publish_progress,
            )

        def on_progress(current, total, filename):
            self.progress.configure(maximum=total, value=current)
            self.status_var.set(f"正在处理 ({current}/{total})：{filename}")

        def on_success(result):
            results, error = result
            self._show_result(results, error, file_paths=file_paths, output_dir=output_dir)

        def on_error(exc):
            self._handle_background_error(exc, "替换线程")

        self._task_runner.submit(work, on_success, on_error, on_progress)

    def _reset_busy_state(self):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.configure(value=0)
        self.progress.grid_remove()
        self._set_busy_controls(False)

    def _finish_with_warning(self, message):
        self._reset_busy_state()
        self.status_var.set("未执行替换")
        messagebox.showwarning("提示", message)

    def _finish_with_error(self, message, log_path=None):
        self._reset_busy_state()
        self.status_var.set("替换失败")
        detail = f"替换失败：{message}"
        if log_path:
            detail += f"\n\n错误日志：{log_path}"
        messagebox.showerror("错误", detail)

    def _show_result(self, results, error, file_paths=None, output_dir=None):
        self._reset_busy_state()
        total_count = sum(results.values())
        lines = [
            "替换完成。",
            "",
            f"共处理 {len(results)} 个文件，替换 {total_count} 处。",
            "",
        ]

        if results:
            lines.append("详细结果：")
            for filename, count in results.items():
                lines.append(f"  {filename}: {count} 处")

        if error:
            lines.extend(["", "部分文件处理失败：", error])
            self.status_var.set("替换完成（部分失败）")
        else:
            self.status_var.set("替换完成")

        self._last_output_dir_to_open = output_dir or self._default_output_dir_to_open(file_paths=file_paths)
        self._show_result_window("\n".join(lines))

    def _default_output_dir_to_open(self, file_paths=None):
        if self.output_dir:
            return self.output_dir
        source_files = file_paths if file_paths is not None else self.replace_files
        if source_files:
            return os.path.dirname(source_files[0])
        return None

    def _open_output_dir(self):
        directory = self._last_output_dir_to_open or self._default_output_dir_to_open()
        if not directory:
            return
        try:
            os.startfile(directory)
        except Exception as exc:
            messagebox.showerror("错误", f"无法打开输出目录：{exc}")

    def _show_result_window(self, message):
        window = tk.Toplevel(self.root)
        window.title("替换结果")
        window.geometry("680x420")
        window.minsize(520, 300)
        window.configure(bg=self.app_bg)
        window.transient(self.root)

        body = ttk.Frame(window, padding=12, style="App.TFrame")
        body.pack(fill="both", expand=True)
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        text = tk.Text(
            body,
            wrap="word",
            font=self.body_font,
            bg=self.surface_bg,
            fg=self.text_fg,
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=8,
        )
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)
        text.insert("1.0", message)
        text.configure(state="disabled")

        buttons = ttk.Frame(body, style="App.TFrame")
        buttons.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="打开输出目录", command=self._open_output_dir, width=14).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="关闭", command=window.destroy, width=10).pack(side="left")

    def show_version_info(self):
        window = tk.Toplevel(self.root)
        window.title("版本信息")
        window.geometry("560x360")
        window.minsize(480, 280)
        window.configure(bg=self.app_bg)
        window.transient(self.root)

        try:
            window.iconphoto(True, self._icon_photo)
        except Exception:
            pass

        body = ttk.Frame(window, padding=12, style="App.TFrame")
        body.pack(fill="both", expand=True)
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        text = tk.Text(
            body,
            wrap="word",
            font=self.body_font,
            bg=self.surface_bg,
            fg=self.text_fg,
            relief="solid",
            borderwidth=1,
            padx=10,
            pady=10,
        )
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)
        text.insert("1.0", format_changelog())
        text.configure(state="disabled")

        buttons = ttk.Frame(body, style="App.TFrame")
        buttons.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="关闭", command=window.destroy, width=10).pack()

    def open_word_table_exporter(self):
        initial_files = self._word_files_for_reuse()
        if self.table_export_window and self.table_export_window.exists():
            if initial_files:
                self.table_export_window._append_files(initial_files, show_status=False)
            self.table_export_window.focus()
            return
        self.table_export_window = WordTableExportWindow(self, initial_files=initial_files)


class WordTableExportWindow:
    def __init__(self, app, initial_files=None):
        self.app = app
        self.window = tk.Toplevel(app.root)
        self._task_runner = BackgroundTaskRunner(self.window)
        self._busy_widgets = []
        self.window.title("提取 Word 表格到 Excel")
        self.window.geometry("1180x720")
        self.window.minsize(900, 640)
        self.window.configure(bg=app.app_bg)
        self.window.transient(app.root)
        try:
            self.window.iconphoto(True, app._icon_photo)
        except Exception:
            pass

        self.file_paths = []
        self.table_items = []
        self.table_item_by_iid = {}
        self.selected_table_keys = set()
        self.output_dir = None
        self._last_output_dir_to_open = None
        self.status_var = tk.StringVar(value="就绪")

        self._create_widgets()
        if initial_files:
            self._append_files(initial_files, show_status=False)
            self.status_var.set(f"已自动带入 {len(self.file_paths)} 个 Word 文件")
        self.window.protocol("WM_DELETE_WINDOW", self.close)

    def exists(self):
        try:
            return bool(self.window.winfo_exists())
        except Exception:
            return False

    def focus(self):
        try:
            self.window.lift()
            self.window.focus_force()
        except Exception:
            pass

    def close(self):
        if self._task_runner.active:
            messagebox.showinfo("任务进行中", "当前任务仍在读写文件，请等待完成后再关闭窗口。", parent=self.window)
            return
        self.app.table_export_window = None
        self.window.destroy()

    def _create_busy_button(self, parent, **kwargs):
        button = ttk.Button(parent, **kwargs)
        self._busy_widgets.append(button)
        return button

    def _set_busy_controls(self, busy):
        state = "disabled" if busy else "normal"
        for widget in self._busy_widgets:
            try:
                if widget.winfo_exists():
                    widget.config(state=state)
            except tk.TclError:
                continue

    def _begin_indeterminate_task(self, status):
        self._set_busy_controls(True)
        self.progress.configure(mode="indeterminate", value=0)
        self.progress.grid()
        self.progress.start(12)
        self.status_var.set(status)

    def _begin_determinate_task(self, status, maximum):
        self._set_busy_controls(True)
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=maximum, value=0)
        self.progress.grid()
        self.status_var.set(status)

    def _create_widgets(self):
        body = ttk.Frame(self.window, padding=12, style="App.TFrame")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)

        header = ttk.Frame(body, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(header, text="提取 Word 表格到 Excel", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="选择 .docx 文件；每个 Word 生成一个 Excel，每张表格对应一个工作表。",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        file_frame = ttk.Frame(body, padding=(12, 9), style="Surface.TFrame")
        file_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        file_frame.columnconfigure(0, weight=1)

        file_header = ttk.Frame(file_frame, style="Toolbar.TFrame")
        file_header.grid(row=0, column=0, sticky="ew")
        file_header.columnconfigure(0, weight=1)
        ttk.Label(file_header, text="Word 文件", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")

        file_toolbar = ttk.Frame(file_header, style="Toolbar.TFrame")
        file_toolbar.grid(row=0, column=1, sticky="e")
        self._create_busy_button(file_toolbar, text="添加文件", command=self.select_files, width=10).pack(side="left", padx=(0, 6))
        self._create_busy_button(file_toolbar, text="添加文件夹", command=self.select_folder, width=11).pack(side="left", padx=(0, 6))
        self._create_busy_button(file_toolbar, text="移除选中", command=self.remove_selected_files, width=10).pack(side="left", padx=(0, 6))
        self._create_busy_button(file_toolbar, text="清空", command=self.clear_files, width=7).pack(side="left")

        list_frame = ttk.Frame(file_frame, style="Toolbar.TFrame")
        list_frame.grid(row=1, column=0, sticky="nsew", pady=(7, 0))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.file_listbox = tk.Listbox(
            list_frame,
            height=2,
            selectmode="extended",
            exportselection=False,
            font=self.app.small_font,
            bg=self.app.surface_bg,
            fg=self.app.text_fg,
            highlightthickness=1,
            highlightbackground=self.app.border_color,
            relief="flat",
        )
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=scrollbar.set)

        self.files_label = ttk.Label(file_frame, text="已选择 0 个 Word 文件", style="Muted.TLabel")
        self.files_label.grid(row=2, column=0, sticky="w", pady=(5, 0))

        scan_frame = ttk.Frame(body, padding=(12, 9), style="Surface.TFrame")
        scan_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        scan_frame.columnconfigure(0, weight=1)
        scan_frame.rowconfigure(1, weight=1)

        scan_header = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        scan_header.grid(row=0, column=0, sticky="ew")
        scan_header.columnconfigure(0, weight=1)
        ttk.Label(scan_header, text="扫描结果", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")

        scan_toolbar = ttk.Frame(scan_header, style="Toolbar.TFrame")
        scan_toolbar.grid(row=0, column=1, sticky="e")
        self.scan_button = self._create_busy_button(scan_toolbar, text="扫描表格", command=self.scan_tables, width=10)
        self.scan_button.pack(side="left", padx=(0, 6))
        self._create_busy_button(scan_toolbar, text="推荐选择", command=self.select_recommended_tables, width=10).pack(side="left", padx=(0, 6))
        self._create_busy_button(scan_toolbar, text="全选", command=self.select_all_tables, width=7).pack(side="left", padx=(0, 6))
        self._create_busy_button(scan_toolbar, text="全不选", command=self.clear_table_selection, width=8).pack(side="left")

        tree_frame = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        tree_frame.grid(row=1, column=0, sticky="nsew", pady=(7, 0))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.scan_tree = ttk.Treeview(
            tree_frame,
            columns=("selected", "file", "table", "size", "hint", "section"),
            show="headings",
            selectmode="extended",
            style="Scan.Treeview",
            height=7,
        )
        self.scan_tree.heading("selected", text="导出")
        self.scan_tree.heading("file", text="文件")
        self.scan_tree.heading("table", text="表格")
        self.scan_tree.heading("size", text="行列")
        self.scan_tree.heading("hint", text="提示")
        self.scan_tree.heading("section", text="所在章节")
        self.scan_tree.column("selected", width=68, minwidth=60, anchor="center", stretch=False)
        self.scan_tree.column("file", width=170, minwidth=120, stretch=False)
        self.scan_tree.column("table", width=70, minwidth=62, anchor="center", stretch=False)
        self.scan_tree.column("size", width=72, minwidth=62, anchor="center", stretch=False)
        self.scan_tree.column("hint", width=320, minwidth=220)
        self.scan_tree.column("section", width=420, minwidth=240)
        self.scan_tree.grid(row=0, column=0, sticky="nsew")
        y_scrollbar = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=self.scan_tree.yview,
            style="Scan.Vertical.TScrollbar",
        )
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar = ttk.Scrollbar(
            tree_frame,
            orient="horizontal",
            command=self.scan_tree.xview,
            style="Scan.Horizontal.TScrollbar",
        )
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.scan_tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        self.scan_tree.bind("<Button-1>", self._toggle_scan_checkbox_from_click)
        self.scan_tree.bind("<Double-1>", self._toggle_scan_row_from_event)
        self.scan_tree.bind("<space>", self._toggle_scan_rows_from_keyboard)
        self.scan_tree.bind("<<TreeviewSelect>>", self._update_scan_detail)

        detail_frame = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        detail_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        detail_frame.columnconfigure(0, weight=1)
        self.detail_text = tk.Text(
            detail_frame,
            height=3,
            wrap="word",
            font=self.app.small_font,
            bg="#F8FAFC",
            fg=self.app.text_fg,
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=6,
        )
        self.detail_text.grid(row=0, column=0, sticky="ew")
        self.detail_text.insert("1.0", "选择扫描结果中的一行，可在这里查看完整章节、表格前文和内容预览。")
        self.detail_text.configure(state="disabled")

        self.scan_label = ttk.Label(scan_frame, text="请先添加 Word 文件并扫描表格", style="Muted.TLabel")
        self.scan_label.grid(row=3, column=0, sticky="w", pady=(5, 0))

        output_frame = ttk.Frame(body, padding=(12, 9), style="Surface.TFrame")
        output_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        output_frame.columnconfigure(0, weight=1)

        output_header = ttk.Frame(output_frame, style="Toolbar.TFrame")
        output_header.grid(row=0, column=0, sticky="ew")
        output_header.columnconfigure(0, weight=1)
        ttk.Label(output_header, text="输出目录", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        self._create_busy_button(output_header, text="选择目录", command=self.select_output_dir, width=12).grid(row=0, column=1, sticky="e")

        self.output_label = ttk.Label(
            output_frame,
            text="未选择则输出到原 Word 文件所在目录",
            style="Muted.TLabel",
            wraplength=700,
            justify="left",
        )
        self.output_label.grid(row=1, column=0, sticky="w", pady=(5, 0))

        action_frame = ttk.Frame(body, padding=(12, 9), style="Surface.TFrame")
        action_frame.grid(row=4, column=0, sticky="ew")
        action_frame.columnconfigure(1, weight=1)

        ttk.Label(action_frame, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(action_frame, mode="determinate", length=220)
        self.progress.grid(row=0, column=1, sticky="ew", padx=(14, 12))
        self.progress.grid_remove()

        self.export_button = self._create_busy_button(
            action_frame,
            text="开始导出",
            command=self.start_export,
            width=14,
            style="Accent.TButton",
        )
        self.export_button.grid(row=0, column=2, sticky="e")

    def select_files(self):
        file_paths = filedialog.askopenfilenames(
            title="选择 Word 文件",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
            parent=self.window,
        )
        if file_paths:
            self._append_files(file_paths)

    def select_folder(self):
        directory = filedialog.askdirectory(title="选择包含 Word 文件的文件夹", parent=self.window)
        if not directory:
            return

        self._begin_indeterminate_task("正在扫描文件夹中的 Word 文件...")

        def work(publish_progress):
            return collect_supported_files(directory, (".docx",), publish_progress)

        def on_progress(scanned_count):
            self.status_var.set(f"正在扫描文件夹：已检查 {scanned_count} 个文件...")

        def on_success(collected):
            self._reset_busy_state()
            if not collected:
                messagebox.showinfo("提示", "该文件夹中未找到 .docx 文件。", parent=self.window)
                return
            self._append_files(collected)

        def on_error(exc):
            self._finish_background_error(exc, "扫描 Word 文件夹")

        self._task_runner.submit(work, on_success, on_error, on_progress)

    def _append_files(self, file_paths, show_status=True):
        existing = {_file_identity(path) for path in self.file_paths}
        added = 0
        skipped = 0
        added_paths = []

        for file_path in file_paths:
            if not self._is_docx(file_path):
                skipped += 1
                continue
            identity = _file_identity(file_path)
            if identity in existing:
                skipped += 1
                continue
            self.file_paths.append(file_path)
            existing.add(identity)
            added_paths.append(file_path)
            added += 1

        self.app._remember_word_files(added_paths)
        self._refresh_file_list()
        if added:
            self._clear_scan_results("文件列表已变化，请重新扫描表格")
        if not show_status:
            return
        if added:
            message = f"已添加 {added} 个 Word 文件"
            if skipped:
                message += f"；跳过 {skipped} 个重复或不支持的文件"
            self.status_var.set(message)
        elif skipped:
            self.status_var.set(f"未添加新文件；跳过 {skipped} 个重复或不支持的文件")

    def _is_docx(self, file_path):
        return (
            not os.path.basename(file_path).startswith("~$")
            and os.path.splitext(file_path)[1].lower() == ".docx"
        )

    def _refresh_file_list(self):
        self.file_listbox.delete(0, "end")
        if self.file_paths:
            self.file_listbox.insert("end", *self.file_paths)
        count = len(self.file_paths)
        foreground = "green" if count else self.app.muted_fg
        self.files_label.config(text=f"已选择 {count} 个 Word 文件", foreground=foreground)

    def remove_selected_files(self):
        selected = list(self.file_listbox.curselection())
        if not selected:
            return
        for index in reversed(selected):
            del self.file_paths[index]
        self._refresh_file_list()
        self._clear_scan_results("文件列表已变化，请重新扫描表格")
        self.status_var.set(f"已移除 {len(selected)} 个文件")

    def clear_files(self):
        if not self.file_paths:
            return
        self.file_paths = []
        self._refresh_file_list()
        self._clear_scan_results("请先添加 Word 文件并扫描表格")
        self.status_var.set("Word 文件列表已清空")

    def scan_tables(self):
        if not self.file_paths:
            messagebox.showwarning("提示", "请先选择 Word 文件。", parent=self.window)
            return

        file_paths = list(self.file_paths)
        self._begin_determinate_task(f"正在扫描 {len(file_paths)} 个 Word 文件...", len(file_paths))

        def work(publish_progress):
            from word_table_exporter import batch_scan_word_tables

            return batch_scan_word_tables(file_paths, progress_callback=publish_progress)

        def on_success(result):
            items, skipped, error = result
            self._show_scan_result(items, skipped, error)

        def on_error(exc):
            self._finish_background_error(exc, "Word 表格扫描线程")

        self._task_runner.submit(
            work,
            on_success,
            on_error,
            lambda current, total, filename: self._update_progress(current, total, filename, "正在扫描"),
        )

    def _show_scan_result(self, items, skipped, error):
        self.table_items = list(items)
        self.table_item_by_iid = {}
        self.selected_table_keys = set()
        self.scan_tree.delete(*self.scan_tree.get_children())

        for row_index, item in enumerate(self.table_items, start=1):
            iid = str(row_index)
            self.table_item_by_iid[iid] = item
            self.scan_tree.insert(
                "",
                "end",
                iid=iid,
                values=self._scan_tree_values(item, selected=False),
            )
        if self.table_items:
            self.scan_tree.selection_set("1")
            self.scan_tree.focus("1")
            self._update_scan_detail()
        else:
            self._set_detail_text("没有扫描到可导出的正文表格。")

        self._reset_busy_state()
        selected_count = len(self.selected_table_keys)
        skipped_count = len(skipped)
        message = f"已扫描到 {len(items)} 张表格，已选择 {selected_count} 张"
        if skipped_count:
            message += f"；跳过 {skipped_count} 个文件"
        if error:
            message += "；部分文件失败"
        self.scan_label.config(text=message, foreground="green" if items else self.app.muted_fg)
        self.status_var.set("扫描完成" if not error else "扫描完成（部分失败）")

        if skipped or error:
            lines = ["扫描完成。", "", message]
            if skipped:
                lines.extend(["", "跳过文件："])
                for filename, reason in skipped.items():
                    lines.append(f"  {filename}: {reason}")
            if error:
                lines.extend(["", "失败文件：", error])
            self._show_result_window("\n".join(lines), title="扫描结果")

    def _clear_scan_results(self, label_text):
        self.table_items = []
        self.table_item_by_iid = {}
        self.selected_table_keys = set()
        if hasattr(self, "scan_tree"):
            self.scan_tree.delete(*self.scan_tree.get_children())
        if hasattr(self, "scan_label"):
            self.scan_label.config(text=label_text, foreground=self.app.muted_fg)
        if hasattr(self, "detail_text"):
            self._set_detail_text("选择扫描结果中的一行，可在这里查看完整章节、表格前文和内容预览。")

    def _scan_tree_values(self, item, selected):
        return (
            "☑" if selected else "☐",
            item.filename,
            f"表格{item.table_index}",
            f"{item.row_count}x{item.column_count}",
            item.hint,
            item.section,
        )

    def _table_key(self, item):
        return (_file_identity(item.file_path), item.table_index)

    def _refresh_scan_row(self, iid, refresh_label=True):
        item = self.table_item_by_iid.get(iid)
        if item is None:
            return
        selected = self._table_key(item) in self.selected_table_keys
        self.scan_tree.item(iid, values=self._scan_tree_values(item, selected))
        if refresh_label:
            self._refresh_scan_label()

    def _refresh_scan_label(self):
        total = len(self.table_items)
        selected = len(self.selected_table_keys)
        text = f"已扫描到 {total} 张表格，已选择 {selected} 张"
        self.scan_label.config(text=text, foreground="green" if selected else self.app.muted_fg)

    def _toggle_scan_row_from_event(self, event):
        row_id = self.scan_tree.identify_row(event.y)
        if row_id:
            self._toggle_scan_iids([row_id])

    def _toggle_scan_checkbox_from_click(self, event):
        if self.scan_tree.identify_region(event.x, event.y) != "cell":
            return None
        if self.scan_tree.identify_column(event.x) != "#1":
            return None
        row_id = self.scan_tree.identify_row(event.y)
        if not row_id:
            return None
        self.scan_tree.selection_set(row_id)
        self.scan_tree.focus(row_id)
        self._toggle_scan_iids([row_id])
        self._update_scan_detail()
        return "break"

    def _toggle_scan_rows_from_keyboard(self, event=None):
        self.toggle_selected_scan_rows()
        return "break"

    def toggle_selected_scan_rows(self):
        selected_rows = self.scan_tree.selection()
        if not selected_rows:
            return
        self._toggle_scan_iids(selected_rows)
        self._update_scan_detail()

    def _toggle_scan_iids(self, iids):
        for iid in iids:
            item = self.table_item_by_iid.get(iid)
            if item is None:
                continue
            key = self._table_key(item)
            if key in self.selected_table_keys:
                self.selected_table_keys.remove(key)
            else:
                self.selected_table_keys.add(key)
            self._refresh_scan_row(iid)

    def select_recommended_tables(self):
        self.selected_table_keys = {
            self._table_key(item)
            for item in self.table_items
            if item.hint.startswith("建议关注")
        }
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set(f"已按提示选择 {len(self.selected_table_keys)} 张建议关注的表格")

    def select_all_tables(self):
        self.selected_table_keys = {self._table_key(item) for item in self.table_items}
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set(f"已选择全部 {len(self.selected_table_keys)} 张表格")

    def clear_table_selection(self):
        self.selected_table_keys = set()
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set("已取消所有表格选择")

    def _refresh_all_scan_rows(self):
        for iid in self.table_item_by_iid:
            self._refresh_scan_row(iid, refresh_label=False)
        self._refresh_scan_label()

    def _update_scan_detail(self, event=None):
        selection = self.scan_tree.selection()
        iid = selection[0] if selection else self.scan_tree.focus()
        item = self.table_item_by_iid.get(iid)
        if item is None:
            self._set_detail_text("选择扫描结果中的一行，可在这里查看完整章节、表格前文和内容预览。")
            return

        selected = "是" if self._table_key(item) in self.selected_table_keys else "否"
        detail = (
            f"导出：{selected}    文件：{item.filename}    表格：{item.table_index}    行列：{item.row_count}x{item.column_count}\n"
            f"所在章节：{item.section}\n"
            f"提示：{item.hint}\n"
            f"表格前文：{item.context}\n"
            f"内容预览：{item.preview}"
        )
        self._set_detail_text(detail)

    def _set_detail_text(self, text):
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", text)
        self.detail_text.configure(state="disabled")

    def select_output_dir(self):
        directory = filedialog.askdirectory(title="选择输出目录", parent=self.window)
        if not directory:
            return
        self.output_dir = directory
        self.output_label.config(text=directory, foreground="green")
        self.status_var.set(f"已选择输出目录：{directory}")

    def start_export(self):
        if not self.file_paths:
            messagebox.showwarning("提示", "请先选择 Word 文件。", parent=self.window)
            return
        if not self.table_items:
            messagebox.showwarning("提示", "请先点击“扫描表格”，再选择需要导出的表格。", parent=self.window)
            return
        if not self.selected_table_keys:
            messagebox.showwarning("提示", "请先在扫描结果中选择至少一张表格。", parent=self.window)
            return

        selected_tables = {}
        for item in self.table_items:
            key = self._table_key(item)
            if key not in self.selected_table_keys:
                continue
            selected_tables.setdefault(key[0], []).append(item.table_index)

        file_paths = [
            file_path
            for file_path in self.file_paths
            if _file_identity(file_path) in selected_tables
        ]
        output_dir = self.output_dir
        self._begin_determinate_task(
            f"正在导出 {len(self.selected_table_keys)} 张已选表格...",
            len(file_paths),
        )

        def work(publish_progress):
            from word_table_exporter import batch_export_word_tables

            return batch_export_word_tables(
                file_paths,
                output_dir=output_dir,
                progress_callback=publish_progress,
                selected_tables=selected_tables,
            )

        def on_success(result):
            results, skipped, error = result
            self._show_export_result(results, skipped, error, file_paths)

        def on_error(exc):
            self._finish_background_error(exc, "Word 表格导出线程")

        self._task_runner.submit(work, on_success, on_error, self._update_progress)

    def _update_progress(self, current, total, filename, action="正在导出"):
        self.progress.configure(maximum=total, value=current)
        self.status_var.set(f"{action} ({current}/{total})：{filename}")

    def _reset_busy_state(self):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.configure(value=0)
        self.progress.grid_remove()
        self._set_busy_controls(False)

    def _finish_background_error(self, exc, context):
        log_path = write_error_log(type(exc), exc, exc.__traceback__, context=context)
        self._finish_with_error(str(exc), log_path)

    def _finish_with_error(self, message, log_path=None):
        self._reset_busy_state()
        self.status_var.set("导出失败")
        detail = f"导出失败：{message}"
        if log_path:
            detail += f"\n\n错误日志：{log_path}"
        messagebox.showerror("错误", detail, parent=self.window)

    def _show_export_result(self, results, skipped, error, source_files):
        self._reset_busy_state()
        total_tables = sum(int(item["tables"]) for item in results.values())
        lines = [
            "导出完成。",
            "",
            f"成功导出 {len(results)} 个文件，共 {total_tables} 张表格。",
            "",
        ]

        if results:
            lines.append("成功文件：")
            for filename, info in results.items():
                lines.append(f"  {filename}: {info['tables']} 张表格 -> {info['output_path']}")

        if skipped:
            lines.extend(["", "跳过文件："])
            for filename, reason in skipped.items():
                lines.append(f"  {filename}: {reason}")

        if error:
            lines.extend(["", "失败文件：", error])
            self.status_var.set("导出完成（部分失败）")
        else:
            self.status_var.set("导出完成")

        self._last_output_dir_to_open = self._default_output_dir_to_open(results, source_files)
        self._show_result_window("\n".join(lines))

    def _default_output_dir_to_open(self, results=None, source_files=None):
        if self.output_dir:
            return self.output_dir
        if results:
            first = next(iter(results.values()), None)
            if first and first.get("output_path"):
                return os.path.dirname(first["output_path"])
        if source_files:
            return os.path.dirname(source_files[0])
        return None

    def _open_output_dir(self):
        directory = self._last_output_dir_to_open or self._default_output_dir_to_open()
        if not directory:
            return
        try:
            os.startfile(directory)
        except Exception as exc:
            messagebox.showerror("错误", f"无法打开输出目录：{exc}", parent=self.window)

    def _show_result_window(self, message, title="导出结果"):
        window = tk.Toplevel(self.window)
        window.title(title)
        window.geometry("720x440")
        window.minsize(560, 320)
        window.configure(bg=self.app.app_bg)
        window.transient(self.window)

        body = ttk.Frame(window, padding=12, style="App.TFrame")
        body.pack(fill="both", expand=True)
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        text = tk.Text(
            body,
            wrap="word",
            font=self.app.body_font,
            bg=self.app.surface_bg,
            fg=self.app.text_fg,
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=8,
        )
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)
        text.insert("1.0", message)
        text.configure(state="disabled")

        buttons = ttk.Frame(body, style="App.TFrame")
        buttons.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="打开输出目录", command=self._open_output_dir, width=14).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="关闭", command=window.destroy, width=10).pack(side="left")


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    root = tk.Tk()
    install_exception_handlers(root)
    ReplaceSimpleApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
