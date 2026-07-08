"""
replace-simple GUI 入口。
只保留批量文本替换功能。
"""

import ctypes
import json
import os
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from tksheet import Sheet

# string_replacer 会拖入 openpyxl(~290ms)+ docx(~100ms)，合计约 400ms。
# 改为按需延迟导入：仅在「导入 Excel」或「开始替换」时加载，让窗口瞬间弹出。


DEFAULT_BLANK_RULE_ROWS = 4
SUPPORTED_EXTENSIONS = (".docx", ".xlsx", ".xlsm", ".pptx")
APP_STATE_DIR_NAME = "replace-simple"
ERROR_LOG_NAME = "error.log"
SETTINGS_NAME = "settings.json"

COMPRESSED_SHEET_HEIGHT = 210      # 基准规则表像素高度（默认窗口下约露 6 行 + 表头，且底部按钮可见）
MIN_SHEET_HEIGHT = 120             # 窗口较矮或系统缩放较大时，优先保住底部操作区

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


class ReplaceSimpleApp:
    def __init__(self, root, restore_session=True):
        self.root = root
        self.restore_session = restore_session
        self.root.title("replace-simple")
        self.root.geometry("860x720")
        self.root.minsize(800, 660)

        # 窗口标题栏 / 任务栏图标
        self._apply_window_icon()

        self.replace_files = []
        self.output_dir = None
        self._last_output_dir_to_open = None

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

    def _create_widgets(self):
        self.container = ttk.Frame(self.root, padding=(16, 12), style="App.TFrame")
        self.container.pack(fill="both", expand=True)
        self.container.columnconfigure(0, weight=1)

        self.header_frame = ttk.Frame(self.container, style="App.TFrame")
        header_frame = self.header_frame
        header_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(header_frame, text="批量文本替换", style="Title.TLabel").pack(anchor="w")
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
        ttk.Button(file_toolbar, text="添加文件", command=self.select_files, width=10).pack(side="left", padx=(0, 6))
        ttk.Button(file_toolbar, text="添加文件夹", command=self.select_folder, width=11).pack(side="left", padx=(0, 6))
        ttk.Button(file_toolbar, text="移除选中", command=self.remove_selected_files, width=10).pack(side="left", padx=(0, 6))
        ttk.Button(file_toolbar, text="清空", command=self.clear_files, width=7).pack(side="left")

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
        ttk.Button(rules_toolbar, text="新增一行", command=self.add_rule_row, width=9).pack(side="left", padx=(0, 6))
        ttk.Button(rules_toolbar, text="删除选中", command=self.delete_selected_rules, width=9).pack(side="left", padx=(0, 6))
        ttk.Button(rules_toolbar, text="清空规则", command=self.clear_rules, width=9).pack(side="left", padx=(0, 6))
        ttk.Button(rules_toolbar, text="导入 Excel", command=self.import_rules_from_excel, width=10).pack(side="left", padx=(0, 6))
        ttk.Button(rules_toolbar, text="导出 Excel", command=self.export_rules_to_excel, width=10).pack(side="left")

        self.rules_hint = ttk.Label(
            rules_frame,
            text="双击单元格可编辑；粘贴可自动向下扩行；执行时按原文长度长词优先。",
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
        ttk.Button(output_header, text="选择目录", command=self.select_output_dir, width=12).grid(row=0, column=1, sticky="e")

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
        self.start_button = ttk.Button(
            action_row, text="开始替换", command=self.start_replace, width=14, style="Accent.TButton",
        )
        self.start_button.grid(row=0, column=2, sticky="e")

        ttk.Label(
            bottom_frame,
            text="规则表第一列为原文、第二列为替换文；处理文件支持 .docx / .xlsx / .xlsm / .pptx。",
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

    def _remove_rules_file_from_replace_files(self, file_path):
        remaining, removed = remove_matching_file_paths(self.replace_files, file_path)
        if not removed:
            return False

        self.replace_files = remaining
        self._refresh_file_list()
        messagebox.showinfo(
            "提示",
            "导入的规则表已从待处理文件列表中自动移除，避免规则表被一起替换：\n"
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

        try:
            from string_replacer import load_replacement_rules
            rules = load_replacement_rules(file_path)
        except Exception as exc:
            messagebox.showerror("错误", f"导入规则失败：{exc}")
            return

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

        wb = None
        try:
            from openpyxl import Workbook

            wb = Workbook()
            ws = wb.active
            ws.title = "替换规则"
            for row_index, (old_text, new_text) in enumerate(rules, start=1):
                ws.cell(row=row_index, column=1, value=old_text)
                ws.cell(row=row_index, column=2, value=new_text)
            ws.column_dimensions["A"].width = 32
            ws.column_dimensions["B"].width = 32
            wb.save(file_path)
        except PermissionError:
            messagebox.showerror("错误", "导出失败：请先关闭正在打开的规则表文件后重试。")
            return
        except Exception as exc:
            messagebox.showerror("错误", f"导出失败：{exc}")
            return
        finally:
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass

        self.status_var.set(f"已导出 {len(rules)} 条规则：{os.path.basename(file_path)}")

    # ---------- 文件 / 输出目录 ----------
    def _is_supported_file(self, file_path):
        return os.path.splitext(file_path)[1].lower() in SUPPORTED_EXTENSIONS

    def _refresh_file_list(self):
        self.file_listbox.delete(0, "end")
        for file_path in self.replace_files:
            self.file_listbox.insert("end", file_path)
        self._update_files_label()

    def _append_replace_files(self, file_paths):
        existing = {_file_identity(path) for path in self.replace_files}
        added = 0
        skipped = 0

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
            added += 1

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

        collected = []
        for root_dir, _dirs, files in os.walk(directory):
            for filename in files:
                file_path = os.path.join(root_dir, filename)
                if self._is_supported_file(file_path):
                    collected.append(file_path)

        if not collected:
            messagebox.showinfo("提示", "该文件夹中未找到支持的 Office 文件。")
            return

        self._append_replace_files(collected)

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

        rules = self.get_rules_from_table()
        if not rules:
            messagebox.showwarning("提示", "请先新增或导入至少一条替换规则。")
            return

        self.start_button.config(state="disabled")
        self.progress.grid()
        self.progress.configure(maximum=len(self.replace_files), value=0)
        self.status_var.set(f"正在使用 {len(rules)} 条规则替换...")

        def task():
            try:
                def progress_callback(current, total, filename):
                    def update_progress():
                        self.progress.configure(maximum=total, value=current)
                        self.status_var.set(f"正在处理 ({current}/{total})：{filename}")

                    self.root.after(0, update_progress)

                from string_replacer import batch_replace
                results, error = batch_replace(
                    self.replace_files,
                    rules,
                    output_dir=self.output_dir,
                    progress_callback=progress_callback,
                )

                self.root.after(0, lambda: self._show_result(results, error))
            except Exception as exc:
                log_path = write_error_log(type(exc), exc, exc.__traceback__, context="替换线程")
                self.root.after(0, lambda: self._finish_with_error(str(exc), log_path))

        threading.Thread(target=task, daemon=True).start()

    def _reset_busy_state(self):
        self.progress.configure(value=0)
        self.progress.grid_remove()
        self.start_button.config(state="normal")

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

    def _show_result(self, results, error):
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

        self._last_output_dir_to_open = self._default_output_dir_to_open()
        self._show_result_window("\n".join(lines))

    def _default_output_dir_to_open(self):
        if self.output_dir:
            return self.output_dir
        if self.replace_files:
            return os.path.dirname(self.replace_files[0])
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
