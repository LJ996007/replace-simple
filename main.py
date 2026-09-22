"""批量替换主窗口与程序入口。"""

import ctypes
import os
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk
from tksheet import Sheet
from app_info import format_changelog, format_version_title
from app_settings import load_settings, save_settings
from replacement_rules import DELETE_MARKER, normalize_rule_rows, validate_rules
from ui_common import (
    DEFAULT_BLANK_RULE_ROWS,
    DEFAULT_PRESETS,
    SUPPORTED_EXTENSIONS,
    WORD_EXTENSIONS,
    DEFAULT_SYMBOL_CHARS,
    COMPRESSED_SHEET_HEIGHT,
    MIN_SHEET_HEIGHT,
    TASK_POLL_INTERVAL_MS,
    FOLDER_SCAN_PROGRESS_INTERVAL,
    OUTPUT_DIR_HINT,
    OUTPUT_DIR_MAX_LINES,
    GRID_COLOR,
    HEADER_GRID_COLOR,
    HEADER_BG,
    ZEBRA_BG,
    elide_middle,
    HoverTooltip,
    ElidedTextController,
    CanvasCheckbox,
    center_window_on_parent,
    write_error_log,
    install_exception_handlers,
    normalize_presets,
    presets_from_settings,
    symbol_chars_from_settings,
    keep_clause_symbols_from_settings,
    append_presets_to_rule_rows,
    resource_path,
    make_blank_rows,
    _file_identity,
    remove_matching_file_paths,
    collect_supported_files,
    BackgroundTaskRunner,
    cancel_widget_callbacks,
    CappedScrollbarModel,
    FlatScrollbar,
    cap_existing_scrollbar,
)
from word_export_window import WordTableExportWindow


class ReplaceSimpleApp:
    def __init__(self, root, restore_session=True):
        self.root = root
        root.bind("<Destroy>", cancel_widget_callbacks, add="+")
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
        self.same_dir_replace = False
        self._last_output_dir_to_open = None
        self._rules_resize_after = None
        self.table_export_window = None
        self._project_info_details = []
        if restore_session:
            settings = load_settings()
            self.presets = presets_from_settings(settings)
            self.symbol_chars = symbol_chars_from_settings(settings)
            self.keep_clause_symbols = keep_clause_symbols_from_settings(settings)
        else:
            self.presets = list(DEFAULT_PRESETS)
            self.symbol_chars = DEFAULT_SYMBOL_CHARS
            self.keep_clause_symbols = False
        self.preset_popup = None

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
        # 方案 A：柔和护眼浅色（去纯白、略降对比与冷蓝刺眼感）
        self.app_bg = "#F0F2F5"
        self.surface_bg = "#F7F8FA"
        self.border_color = "#C8CED6"
        self.muted_fg = "#5F6B7A"
        self.accent_fg = "#1A6BB5"
        self.text_fg = "#2B2F36"

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
        style.configure("TLabel", font=self.body_font, background=self.surface_bg, foreground=self.text_fg)
        style.configure("TButton", font=self.body_font, padding=(10, 5))
        for control in ("TEntry", "TCombobox"):
            style.configure(control, font=self.body_font, fieldbackground=self.surface_bg,
                            background=HEADER_BG, foreground=self.text_fg,
                            bordercolor=self.border_color, lightcolor=self.border_color,
                            darkcolor=self.border_color, padding=5,
                            selectbackground="#C6D8EB", selectforeground=self.text_fg,
                            arrowcolor=self.muted_fg)
            style.map(control,
                      fieldbackground=[("disabled", self.app_bg), ("readonly", self.surface_bg)],
                      foreground=[("disabled", self.muted_fg), ("readonly", self.text_fg)],
                      selectbackground=[("readonly", "#C6D8EB")],
                      selectforeground=[("readonly", self.text_fg)],
                      bordercolor=[("focus", self.accent_fg)])
        style.configure("TCheckbutton", font=self.body_font, background=self.surface_bg,
                        foreground=self.text_fg, indicatorbackground=self.surface_bg,
                        indicatorforeground=self.accent_fg)
        style.map("TCheckbutton", background=[("active", self.surface_bg)],
                  indicatorbackground=[("active", "#E3EDF8"), ("selected", self.surface_bg)],
                  foreground=[("disabled", self.muted_fg)])
        self.root.option_add("*TCombobox*Listbox.background", self.surface_bg)
        self.root.option_add("*TCombobox*Listbox.foreground", self.text_fg)
        self.root.option_add("*TCombobox*Listbox.selectBackground", "#C6D8EB")
        self.root.option_add("*TCombobox*Listbox.selectForeground", self.text_fg)
        style.configure("Title.TLabel", font=self.title_font, background=self.app_bg)
        style.configure("Subtitle.TLabel", font=self.small_font, foreground=self.muted_fg, background=self.app_bg)
        style.configure("SectionTitle.TLabel", font=self.section_font, background=self.surface_bg)
        style.configure("Hint.TLabel", font=self.small_font, foreground=self.muted_fg, background=self.surface_bg)
        style.configure("Muted.TLabel", font=self.body_font, foreground=self.muted_fg, background=self.surface_bg)
        style.configure("Status.TLabel", font=self.body_font, foreground=self.accent_fg, background=self.surface_bg)
        # 统一使用扁平淡蓝滚动条；两端保留小箭头作为方向提示，也避免内容不足
        # 一页时滑块从视觉上铺满整条轨道。
        flat_vertical_layout = [
            ("Vertical.Scrollbar.trough", {
                "sticky": "ns",
                "children": [
                    ("Vertical.Scrollbar.uparrow", {"side": "top", "sticky": "ew"}),
                    ("Vertical.Scrollbar.downarrow", {"side": "bottom", "sticky": "ew"}),
                    ("Vertical.Scrollbar.thumb", {"sticky": "nswe"}),
                ],
            }),
        ]
        flat_horizontal_layout = [
            ("Horizontal.Scrollbar.trough", {
                "sticky": "we",
                "children": [
                    ("Horizontal.Scrollbar.leftarrow", {"side": "left", "sticky": "ns"}),
                    ("Horizontal.Scrollbar.rightarrow", {"side": "right", "sticky": "ns"}),
                    ("Horizontal.Scrollbar.thumb", {"sticky": "nswe"}),
                ],
            }),
        ]
        style.layout("Flat.Vertical.TScrollbar", flat_vertical_layout)
        style.layout("Flat.Horizontal.TScrollbar", flat_horizontal_layout)
        style.configure(
            "Flat.Vertical.TScrollbar",
            gripcount=0,
            width=15,
            arrowsize=13,
            background="#BFD3E7",
            troughcolor="#EDF3F8",
            bordercolor="#EDF3F8",
            arrowcolor="#5D7891",
            lightcolor="#BFD3E7",
            darkcolor="#BFD3E7",
            relief="flat",
            borderwidth=0,
        )
        style.map(
            "Flat.Vertical.TScrollbar",
            background=[("pressed", "#83AED3"), ("active", "#A5C4DF")],
            arrowcolor=[("pressed", "#315F86"), ("active", "#416F96")],
            lightcolor=[("pressed", "#83AED3"), ("active", "#A5C4DF")],
            darkcolor=[("pressed", "#83AED3"), ("active", "#A5C4DF")],
        )
        style.configure(
            "Flat.Horizontal.TScrollbar",
            gripcount=0,
            width=15,
            arrowsize=13,
            background="#BFD3E7",
            troughcolor="#EDF3F8",
            bordercolor="#EDF3F8",
            arrowcolor="#5D7891",
            lightcolor="#BFD3E7",
            darkcolor="#BFD3E7",
            relief="flat",
            borderwidth=0,
        )
        style.map(
            "Flat.Horizontal.TScrollbar",
            background=[("pressed", "#83AED3"), ("active", "#A5C4DF")],
            arrowcolor=[("pressed", "#315F86"), ("active", "#416F96")],
            lightcolor=[("pressed", "#83AED3"), ("active", "#A5C4DF")],
            darkcolor=[("pressed", "#83AED3"), ("active", "#A5C4DF")],
        )
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
            background=[("active", "#DDE2E8")],
            foreground=[("active", self.text_fg)],
        )
        style.layout("Scan.Vertical.TScrollbar", flat_vertical_layout)
        style.layout("Scan.Horizontal.TScrollbar", flat_horizontal_layout)
        for scan_style in ("Scan.Vertical.TScrollbar", "Scan.Horizontal.TScrollbar"):
            style.configure(
                scan_style,
                gripcount=0,
                width=15,
                arrowsize=13,
                background="#BFD3E7",
                troughcolor="#EDF3F8",
                bordercolor="#EDF3F8",
                arrowcolor="#5D7891",
                lightcolor="#BFD3E7",
                darkcolor="#BFD3E7",
                relief="flat",
                borderwidth=0,
            )
            style.map(
                scan_style,
                background=[("pressed", "#83AED3"), ("active", "#A5C4DF")],
                arrowcolor=[("pressed", "#315F86"), ("active", "#416F96")],
                lightcolor=[("pressed", "#83AED3"), ("active", "#A5C4DF")],
                darkcolor=[("pressed", "#83AED3"), ("active", "#A5C4DF")],
            )

        # 扁平次级按钮（clam 主题下方可定制 background/relief）
        style.configure(
            "TButton", font=self.body_font, padding=(12, 7), relief="flat",
            background="#E1E6ED", foreground=self.text_fg,
            borderwidth=1, bordercolor="#B8C0CC", focuscolor=self.surface_bg,
        )
        style.map(
            "TButton",
            background=[("pressed", "#D0D6DF"), ("active", "#D9DFE7"), ("disabled", "#EBEEF2")],
            bordercolor=[("active", "#AEB6C2")],
            foreground=[("disabled", "#9AA4B2")],
        )
        # 主操作按钮（开始替换）：蓝底白字，一眼可见的主行动点
        style.configure(
            "Accent.TButton", font=self.body_font, padding=(14, 8), relief="flat",
            background=self.accent_fg, foreground="#F7F8FA",
            borderwidth=0, focuscolor=self.accent_fg,
        )
        style.map(
            "Accent.TButton",
            background=[("pressed", "#14588F"), ("active", "#175FA0"), ("disabled", "#9DB6CF")],
            foreground=[("disabled", "#E8EEF5")],
        )
        # 进度条配色（与强调色一致）
        style.configure(
            "Horizontal.TProgressbar", troughcolor="#E0E5EB",
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

    def _bind_vertical_mousewheel(self, scrollable, *widgets):
        """让鼠标停在内容区或滚动条上时均可使用滚轮。"""
        def on_mousewheel(event):
            delta = getattr(event, "delta", 0)
            if not delta:
                return None
            steps = -int(delta / 120)
            if not steps:
                steps = -1 if delta > 0 else 1
            scrollable.yview_scroll(steps, "units")
            return "break"

        for widget in widgets:
            widget.bind("<MouseWheel>", on_mousewheel, add="+")
        return on_mousewheel

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
        self._refresh_output_and_status_layout()

    def _begin_determinate_task(self, status, maximum):
        self._set_busy_controls(True)
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=maximum, value=0)
        self.progress.grid()
        self.status_var.set(status)
        self._refresh_output_and_status_layout()

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
            text="提取信息",
            command=self.open_word_table_exporter,
            width=10,
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
        file_scrollbar = FlatScrollbar(
            file_list_frame,
            orient="vertical",
            command=self.file_listbox.yview,
            style="Flat.Vertical.TScrollbar",
        )
        file_scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=file_scrollbar.set)
        self._bind_vertical_mousewheel(self.file_listbox, self.file_listbox, file_scrollbar)

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
        self._create_busy_button(rules_toolbar, text="新增一行", command=self.add_rule_row, width=7).pack(side="left", padx=(0, 4))
        self._create_busy_button(rules_toolbar, text="删除选中", command=self.delete_selected_rules, width=7).pack(side="left", padx=(0, 4))
        self._create_busy_button(rules_toolbar, text="清空规则", command=self.clear_rules, width=7).pack(side="left", padx=(0, 4))
        more = ttk.Menubutton(rules_toolbar, text="更多 ▼", width=7, style="TButton")
        more.pack(side="left", padx=(0, 4))
        self._busy_widgets.append(more)
        menu = tk.Menu(more, tearoff=False)
        menu.add_command(label="将选中规则设为明确删除", command=self.mark_selected_rules_for_deletion)
        menu.add_command(label="保存规则为 Excel…", command=self.save_rules_to_excel)
        menu.add_command(label="执行前检查", command=self.preview_replacement)
        menu.add_command(label="查看项目信息来源", command=self.show_project_info_details)
        more.configure(menu=menu)
        self._create_busy_button(rules_toolbar, text="导入 Excel", command=self.import_rules_from_excel, width=9).pack(side="left", padx=(0, 4))
        self._create_busy_button(rules_toolbar, text="导入招标文件", command=self.import_rules_from_tender_file, width=11).pack(side="left", padx=(0, 4))
        self.preset_button = self._create_busy_button(
            rules_toolbar, text="预设 ▼", command=self.toggle_preset_popup, width=7
        )
        self.preset_button.pack(side="left")

        self.rules_hint = ttk.Label(
            rules_frame,
            text=f"空白替换值跳过；明确删除请用“更多”或填入 {DELETE_MARKER}；同原文冲突须先修正。",
            style="Hint.TLabel",
        )
        self.rules_hint.grid(row=1, column=0, sticky="ew", pady=(4, 6))

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
            scrollbar_show_arrows=True,
            vertical_scroll_background="#BFD3E7",
            horizontal_scroll_background="#BFD3E7",
            vertical_scroll_troughcolor="#EDF3F8",
            horizontal_scroll_troughcolor="#EDF3F8",
            vertical_scroll_lightcolor="#BFD3E7",
            horizontal_scroll_lightcolor="#BFD3E7",
            vertical_scroll_darkcolor="#BFD3E7",
            horizontal_scroll_darkcolor="#BFD3E7",
            vertical_scroll_bordercolor="#EDF3F8",
            horizontal_scroll_bordercolor="#EDF3F8",
            vertical_scroll_borderwidth=0,
            horizontal_scroll_borderwidth=0,
            vertical_scroll_relief="flat",
            horizontal_scroll_relief="flat",
            vertical_scroll_troughrelief="flat",
            horizontal_scroll_troughrelief="flat",
            vertical_scroll_not_active_bg="#BFD3E7",
            horizontal_scroll_not_active_bg="#BFD3E7",
            vertical_scroll_active_bg="#A5C4DF",
            horizontal_scroll_active_bg="#A5C4DF",
            vertical_scroll_pressed_bg="#83AED3",
            horizontal_scroll_pressed_bg="#83AED3",
            vertical_scroll_not_active_fg="#5D7891",
            horizontal_scroll_not_active_fg="#5D7891",
            vertical_scroll_active_fg="#416F96",
            horizontal_scroll_active_fg="#416F96",
            vertical_scroll_pressed_fg="#315F86",
            horizontal_scroll_pressed_fg="#315F86",
            vertical_scroll_arrowsize=13,
            horizontal_scroll_arrowsize=13,
            font=(self.font_family, 10, "normal"),
            header_font=(self.font_family, 10, "bold"),
            paste_can_expand_x=False,
            paste_can_expand_y=True,
        )
        self.rules_sheet.grid(row=2, column=0, sticky="new")
        # 规则表只有两列且始终自适应填满，不需要水平滚动；隐藏 tksheet 默认
        # 常显的横向滚动条，避免它占位、显得多余。
        self.rules_sheet.hide("x_scrollbar")
        # tksheet 自己创建滚动条，创建后再套用与应用其余位置相同的限长样式。
        cap_existing_scrollbar(
            self.rules_sheet.yscroll,
            self.rules_sheet.MT._yscrollbar,
            self.rules_sheet.MT,
            orientation="vertical",
        )

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
            header_selected_columns_bg="#D2DCE9",
            index_bg=HEADER_BG,
            index_fg=self.muted_fg,
            index_grid_fg=HEADER_GRID_COLOR,
            table_selected_cells_bg="#C6D8EB",
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

        # 标题行：左标题 + 右侧输出方式按钮
        output_header = ttk.Frame(bottom_frame, style="Toolbar.TFrame")
        output_header.grid(row=0, column=0, sticky="ew")
        output_header.columnconfigure(0, weight=1)
        ttk.Label(output_header, text="3  输出目录", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.same_dir_button = self._create_busy_button(
            output_header,
            text="同目录替换",
            command=self.select_same_dir_replace,
            width=13,
        )
        self.same_dir_button.grid(row=0, column=1, sticky="e", padx=(0, 6))
        self._create_busy_button(output_header, text="选择目录", command=self.select_output_dir, width=12).grid(row=0, column=2, sticky="e")

        # 输出路径/说明单独占整行；按面板宽度换行，且最多两行（中间省略），
        # 避免超长目录把下方「开始替换」顶出可视区。
        self.output_label = ttk.Label(
            bottom_frame, text=OUTPUT_DIR_HINT,
            style="Muted.TLabel", wraplength=800, justify="left",
        )
        self.output_label.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        self._output_path = ElidedTextController(
            self.output_label,
            self.body_font,
            self._bottom_wraplength,
            max_lines=OUTPUT_DIR_MAX_LINES,
            tooltip=True,
            tooltip_font=self.small_font,
            tooltip_fg=self.text_fg,
            fallback_width=800,
        )

        ttk.Separator(bottom_frame, orient="horizontal").grid(row=2, column=0, sticky="ew", pady=8)

        # 状态行：状态文字单行省略，右侧主按钮列宽固定，不会被长文案挤走
        self.action_row = ttk.Frame(bottom_frame, style="Toolbar.TFrame")
        action_row = self.action_row
        action_row.grid(row=3, column=0, sticky="ew")
        action_row.columnconfigure(0, weight=1)
        action_row.columnconfigure(2, minsize=118)
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(action_row, text="就绪", style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="ew")
        self.progress = ttk.Progressbar(action_row, mode="determinate", length=220)
        self.progress.grid(row=0, column=1, sticky="ew", padx=(14, 12))
        self.progress.grid_remove()   # 空闲时隐藏，避免显示成一根灰槽；开始替换时再显示
        self.start_button = self._create_busy_button(
            action_row, text="开始替换", command=self.start_replace, width=14, style="Accent.TButton",
        )
        self.start_button.grid(row=0, column=2, sticky="e")
        self._status_elide = ElidedTextController(
            self.status_label,
            self.body_font,
            self._status_wraplength,
            max_lines=1,
            fallback_width=360,
        ).attach_var(self.status_var)

        self.footer_hint = ttk.Label(
            bottom_frame,
            text="规则表第一列为原文、第二列为替换文；招标文件导入会生成 [项目名称] 等占位符规则。",
            style="Hint.TLabel",
        )
        self.footer_hint.grid(row=4, column=0, sticky="ew", pady=(7, 0))

        # 监听窗口高度变化，动态调整规则表高度：规则表吃掉剩余高度，
        # 底部「输出目录 / 开始替换」区域始终完整保留在可视区。
        self._rules_resize_after = None
        self.root.bind("<Configure>", self._on_window_configure)
        self.root.after_idle(self._refresh_output_and_status_layout)

        # tksheet 在 Windows / 高 DPI 下偶尔会在首帧拿到不完整的内部画布尺寸，
        # 导致表头和网格暂时不绘制，用户点击后才恢复。窗口完成布局后主动刷新几次，
        # 让规则表打开软件时就稳定显示表头和默认空行。
        self.root.after_idle(self._refresh_rules_sheet_view)
        self.root.after(80, self._refresh_rules_sheet_view)
        self.root.after(250, self._refresh_rules_sheet_view)

    # ---------- 列宽自适应 ----------
    def _on_window_configure(self, event=None):
        """窗口尺寸变化时节流刷新路径换行和规则表高度。"""
        if event is not None and event.widget is not self.root:
            return  # 只处理 root 自身的 Configure，忽略子组件的
        if self._rules_resize_after is not None:
            try:
                self.root.after_cancel(self._rules_resize_after)
            except Exception:
                pass
        self._rules_resize_after = self.root.after(16, self._after_window_configure)

    def _after_window_configure(self):
        self._rules_resize_after = None
        self._refresh_output_and_status_layout()
        self._resize_rules_sheet()

    def _resize_rules_sheet(self):
        """根据当前窗口实际剩余高度动态设定规则表高度。

        旧逻辑用「窗口高度 - 默认高度」简单放大规则表；在 Windows 缩放、
        恢复上次窗口尺寸或主题控件请求高度变化时，底部操作区可能被挤出窗口。
        这里改为先测量顶部、文件区、底部区和规则区固定控件，再把剩余高度给表格。
        """
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

    def _bottom_wraplength(self):
        try:
            width = self.bottom_frame.winfo_width()
        except Exception:
            width = 0
        if width <= 1:
            return 800
        return max(width - 28, 200)

    def _status_wraplength(self):
        try:
            row_width = self.action_row.winfo_width()
            button_width = self.start_button.winfo_reqwidth()
        except Exception:
            return 360
        extra = 20
        try:
            if self.progress.winfo_ismapped():
                extra += int(self.progress.winfo_reqwidth()) + 26
        except Exception:
            pass
        return max(row_width - button_width - extra, 120)

    def _refresh_output_and_status_layout(self):
        """按当前面板宽度限制路径、状态和提示文案，避免挤走操作按钮。"""
        if hasattr(self, "_output_path"):
            self._output_path.refresh()
        if hasattr(self, "_status_elide"):
            self._status_elide.refresh()
        wrap = self._bottom_wraplength()
        try:
            self.rules_hint.configure(wraplength=wrap)
            self.footer_hint.configure(wraplength=wrap)
        except (AttributeError, tk.TclError):
            pass

    def _apply_output_dir(self, directory, *, same_dir_replace=False):
        self.output_dir = directory
        self.same_dir_replace = bool(same_dir_replace and not directory)
        if hasattr(self, "same_dir_button"):
            self.same_dir_button.configure(
                text="✓ 同目录替换" if self.same_dir_replace else "同目录替换",
            )
        if directory:
            self._output_path.set_text(directory, foreground="green")
        elif self.same_dir_replace:
            self._output_path.set_text(
                "同目录替换：结果保存到各待处理文件所在目录；同名时直接覆盖原文件",
                foreground="green",
            )
        else:
            self._output_path.set_text(OUTPUT_DIR_HINT, foreground=self.muted_fg)
        self._refresh_output_and_status_layout()
        self._resize_rules_sheet()

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
            self._apply_output_dir(output_dir)
        elif settings.get("same_dir_replace") is True:
            self._apply_output_dir(None, same_dir_replace=True)

    def _save_session(self):
        if not self.restore_session:
            return
        data = {
            "geometry": self.root.geometry(),
            "output_dir": self.output_dir,
            "same_dir_replace": self.same_dir_replace,
            "rules": normalize_rule_rows(self.rules_sheet.get_sheet_data()),
            "presets": self.presets,
            "symbol_chars": self.symbol_chars,
            "keep_clause_symbols": bool(self.keep_clause_symbols),
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
            self._save_session()
        except Exception as exc:
            messagebox.showerror("保存失败", f"无法保存本次规则和设置：{exc}\n\n窗口将保持打开，请重试关闭，或先将规则保存为 Excel。", parent=self.root)
            return
        if hasattr(self, "_output_path"):
            self._output_path.hide_tooltip()
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
            os.path.splitext(file_path)[1].lower() in WORD_EXTENSIONS
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
                "请在主文件列表或「提取信息」窗口中选中一个招标文件，"
                "或者在接下来的窗口中重新选择。",
            )

        return filedialog.askopenfilename(
            title="从招标文件读取项目信息",
            filetypes=[("Word 文档", "*.doc;*.docx"), ("所有文件", "*.*")],
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
        """未填写的规则保留在界面，仅返回可执行的规则。"""
        rows = self.rules_sheet.get_sheet_data()
        return [(old, new) for old, new in normalize_rule_rows(rows) if new]

    def mark_selected_rules_for_deletion(self):
        for row in self.rules_sheet.get_selected_rows(get_cells_as_rows=True):
            self.rules_sheet.set_cell_data(row, 1, DELETE_MARKER)
        self._on_rules_changed()

    def save_rules_to_excel(self):
        path = filedialog.asksaveasfilename(title="保存替换规则", defaultextension=".xlsx", filetypes=[("Excel 规则表", "*.xlsx")])
        if not path:
            return
        rows = normalize_rule_rows(self.rules_sheet.get_sheet_data())
        self._begin_indeterminate_task("正在保存规则…")

        def work(_progress):
            from string_replacer import save_replacement_rules
            save_replacement_rules(path, rows)

        def on_success(_result):
            self._reset_busy_state()
            self._remove_rules_file_from_replace_files(path)
            self.status_var.set(f"已保存 {len(rows)} 条规则：{path}")

        self._task_runner.submit(work, on_success, lambda exc: self._handle_background_error(exc, "保存规则"))

    def preview_replacement(self):
        from string_replacer import get_output_path

        rows = self.rules_sheet.get_sheet_data()
        try:
            validate_rules(rows)
        except ValueError as exc:
            messagebox.showwarning("规则冲突", str(exc))
            return
        rules = self.get_rules_from_table()
        pending = [old for old, new in normalize_rule_rows(rows) if not new]
        lines = [f"可执行规则：{len(rules)} 条；空白跳过：{len(pending)} 条。"]
        if pending:
            lines.append("待填写：" + "、".join(pending))
        reserved = set()
        for path in self.replace_files:
            output = get_output_path(path, rules, self.output_dir, reserved)
            action = "覆盖原文件" if _file_identity(path) == _file_identity(output) else "生成文件"
            lines.append(f"\n{path}\n→ {output}（{action}）")
        lines.append("\n这是执行前的路径检查；最终文件名以结果明细为准。")
        self._show_result_window("\n".join(lines), title="执行前检查")

    def show_project_info_details(self):
        if not self._project_info_details:
            messagebox.showinfo("项目信息", "请先导入招标文件。")
            return
        lines = [f"来源文件：{self._project_info_source}", "冲突项已留空，请核对候选值后在规则表填写。"]
        for item in self._project_info_details:
            lines.append(f"\n{item['placeholder']}【{item['status']}】 {item['value'] or '待填写'}")
            for candidate in item["candidates"]:
                lines.append(f"  {candidate['source']}：{candidate['value']}")
        self._show_result_window("\n".join(lines), title="项目信息来源")

    def add_rule_row(self):
        selected_rows = self.rules_sheet.get_selected_rows(get_cells_as_rows=True)
        insert_at = max(selected_rows) + 1 if selected_rows else self.rules_sheet.total_rows()
        self.rules_sheet.insert_row(row=["", ""], idx=insert_at)
        try:
            self.rules_sheet.see(row=insert_at, column=0)
            self.rules_sheet.select_row(insert_at)
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

    # ---------- 原文本预设 ----------
    def toggle_preset_popup(self):
        if self.preset_popup and self.preset_popup.winfo_exists():
            self._close_preset_popup()
            return
        self._show_preset_popup()

    def _show_preset_popup(self):
        popup = tk.Toplevel(self.root)
        self.preset_popup = popup
        popup.withdraw()
        popup.overrideredirect(True)
        popup.transient(self.root)
        popup.configure(bg=self.border_color)
        popup.bind("<Escape>", lambda _event: self._close_preset_popup())

        body = ttk.Frame(popup, padding=12, style="Surface.TFrame")
        body.pack(fill="both", expand=True, padx=1, pady=1)
        header = ttk.Frame(body, style="Toolbar.TFrame")
        header.pack(fill="x")
        ttk.Label(header, text="选择原文本预设", style="SectionTitle.TLabel").pack(side="left")
        self._preset_selection_label = ttk.Label(header, text="已选 0 项", style="Hint.TLabel")
        self._preset_selection_label.pack(side="right")
        ttk.Label(
            body, text="可多选，已存在的原文本会自动跳过。", style="Hint.TLabel"
        ).pack(anchor="w", pady=(2, 7))

        choices_border = tk.Frame(body, bg=self.border_color, bd=0)
        choices_border.pack(fill="both", expand=True)
        choices_canvas = tk.Canvas(
            choices_border,
            width=300,
            height=min(max(len(self.presets), 1) * 32, 320),
            bg=self.surface_bg,
            highlightthickness=0,
            bd=0,
        )
        choices_canvas.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        choices_scrollbar = FlatScrollbar(
            choices_border,
            orient="vertical",
            command=choices_canvas.yview,
            style="Flat.Vertical.TScrollbar",
        )
        choices = tk.Frame(choices_canvas, bg=self.surface_bg, bd=0)
        choices_window = choices_canvas.create_window((0, 0), window=choices, anchor="nw")
        choices_canvas.configure(yscrollcommand=choices_scrollbar.set)
        choices.bind(
            "<Configure>",
            lambda _event: choices_canvas.configure(scrollregion=choices_canvas.bbox("all")),
        )
        choices_canvas.bind(
            "<Configure>",
            lambda event: choices_canvas.itemconfigure(choices_window, width=event.width),
        )
        self._preset_vars = []
        if self.presets:
            for preset in self.presets:
                variable = tk.BooleanVar(value=False)
                item = tk.Canvas(
                    choices,
                    height=32,
                    bg=self.surface_bg,
                    bd=0,
                    highlightthickness=0,
                    cursor="hand2",
                )

                def draw_item(canvas=item, value=preset, selected=None):
                    if selected is None:
                        selected = variable.get()
                    background = "#E3EDF8" if selected else self.surface_bg
                    foreground = self.accent_fg if selected else self.text_fg
                    canvas.configure(bg=background)
                    canvas.delete("all")
                    canvas.create_rectangle(
                        10, 9, 23, 22,
                        outline=self.accent_fg if selected else "#6F7B88",
                        width=1,
                        fill="#F7F8FA",
                    )
                    if selected:
                        canvas.create_line(
                            13, 15, 16.5, 19, 21, 12,
                            fill=self.accent_fg,
                            width=2,
                            capstyle="round",
                            joinstyle="round",
                        )
                    text_width = max(int(canvas.winfo_width()) - 40, 80)
                    canvas.create_text(
                        32, 16,
                        text=elide_middle(value, self.body_font, text_width),
                        anchor="w",
                        font=self.body_font,
                        fill=foreground,
                    )

                def toggle_item(_event=None, var=variable, canvas=item, value=preset):
                    selected = not var.get()
                    var.set(selected)
                    draw_item(canvas, value, selected)
                    selected_count = sum(var_.get() for _preset, var_ in self._preset_vars)
                    self._preset_selection_label.configure(text=f"已选 {selected_count} 项")

                draw_item()
                item.bind("<Button-1>", toggle_item)
                item.bind("<Configure>", lambda _event, canvas=item, value=preset: draw_item(canvas, value))
                self._bind_vertical_mousewheel(choices_canvas, item)
                item.pack(fill="x", anchor="w")
                self._preset_vars.append((preset, variable))
        else:
            tk.Label(
                choices,
                text="暂无预设，请先进入管理预设添加。",
                font=self.small_font,
                bg=self.surface_bg,
                fg=self.muted_fg,
                padx=9,
                pady=12,
            ).pack(anchor="w")
        choices.update_idletasks()
        self._bind_vertical_mousewheel(
            choices_canvas, choices_canvas, choices, choices_border, choices_scrollbar
        )
        if choices.winfo_reqheight() > int(choices_canvas.cget("height")):
            choices_scrollbar.pack(side="right", fill="y", pady=1, padx=(0, 1))

        actions = ttk.Frame(body, style="Toolbar.TFrame")
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="管理预设…", command=self.open_preset_manager).pack(side="left")
        ttk.Button(actions, text="取消", command=self._close_preset_popup).pack(side="right")
        ttk.Button(
            actions, text="添加所选", command=self.add_selected_presets, style="Accent.TButton"
        ).pack(side="right", padx=(0, 6))

        popup.update_idletasks()
        x = self.preset_button.winfo_rootx() + self.preset_button.winfo_width() - popup.winfo_reqwidth()
        y = self.preset_button.winfo_rooty() + self.preset_button.winfo_height() + 2
        screen_w = popup.winfo_screenwidth()
        screen_h = popup.winfo_screenheight()
        x = max(0, min(x, screen_w - popup.winfo_reqwidth()))
        y = max(0, min(y, screen_h - popup.winfo_reqheight()))
        popup.geometry(f"+{x}+{y}")
        popup.deiconify()
        popup.lift()
        popup.focus_force()

    def _close_preset_popup(self):
        popup = self.preset_popup
        self.preset_popup = None
        if popup and popup.winfo_exists():
            popup.destroy()

    def add_selected_presets(self):
        selected = [preset for preset, variable in self._preset_vars if variable.get()]
        if not selected:
            self.status_var.set("请先勾选要添加的预设")
            return
        rows, added = append_presets_to_rule_rows(self.rules_sheet.get_sheet_data(), selected)
        self.rules_sheet.set_sheet_data(rows)
        self._ensure_blank_rule_rows()
        self._refresh_rules_sheet_view()
        self._refresh_status()
        self._close_preset_popup()
        skipped = len(selected) - added
        message = f"已从预设添加 {added} 条原文本规则"
        if skipped:
            message += f"；跳过 {skipped} 条已存在规则"
        self.status_var.set(message)

    def open_preset_manager(self):
        self._close_preset_popup()
        dialog = tk.Toplevel(self.root)
        dialog.title("管理预设")
        dialog.withdraw()
        dialog.transient(self.root)
        dialog.resizable(True, True)
        dialog.configure(bg=self.app_bg)
        try:
            dialog.iconphoto(True, self._icon_photo)
        except Exception:
            pass

        outer = ttk.Frame(dialog, padding=14, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        body = ttk.Frame(outer, padding=14, style="Surface.TFrame")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)
        ttk.Label(body, text="管理原文本预设", style="SectionTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(
            body,
            text="选中一项后可修改内容；系统初始预设同样可以编辑或删除。",
            style="Hint.TLabel",
            wraplength=520,
        ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(3, 9))

        list_border = tk.Frame(body, bg=self.border_color, bd=0)
        list_border.grid(row=2, column=0, columnspan=2, sticky="nsew")
        preset_list = tk.Listbox(
            list_border,
            width=38,
            height=12,
            exportselection=False,
            font=self.body_font,
            bg=self.surface_bg,
            fg=self.text_fg,
            selectbackground=self.accent_fg,
            selectforeground="#FFFFFF",
            activestyle="none",
            relief="flat",
            bd=0,
            highlightthickness=0,
        )
        preset_list.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        scrollbar = FlatScrollbar(
            list_border,
            orient="vertical",
            command=preset_list.yview,
            style="Flat.Vertical.TScrollbar",
        )
        scrollbar.pack(side="right", fill="y", pady=1, padx=(0, 1))
        preset_list.configure(yscrollcommand=scrollbar.set)
        self._bind_vertical_mousewheel(preset_list, preset_list, list_border, scrollbar)
        preset_list.insert("end", *self.presets)

        ttk.Label(body, text="预设内容", style="SectionTitle.TLabel").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(12, 5)
        )
        editor = ttk.Frame(body, style="Toolbar.TFrame")
        editor.grid(row=4, column=0, columnspan=2, sticky="ew")
        editor.columnconfigure(0, weight=1)
        value_var = tk.StringVar()
        entry = ttk.Entry(editor, textvariable=value_var, width=25)
        entry.grid(row=0, column=0, sticky="ew")

        def load_selection(_event=None):
            selection = preset_list.curselection()
            if selection:
                value_var.set(preset_list.get(selection[0]))
                entry.select_range(0, "end")

        def add_item():
            value = value_var.get().strip()
            if not value:
                messagebox.showinfo("提示", "请先输入预设内容。", parent=dialog)
                return
            if value in preset_list.get(0, "end"):
                messagebox.showinfo("提示", "该预设已存在。", parent=dialog)
                return
            preset_list.insert("end", value)
            preset_list.selection_clear(0, "end")
            preset_list.selection_set("end")
            preset_list.see("end")
            value_var.set("")
            entry.focus_set()

        def rename_item():
            selection = preset_list.curselection()
            value = value_var.get().strip()
            if not selection:
                messagebox.showinfo("提示", "请先从列表中选择要修改的预设。", parent=dialog)
                return
            if not value:
                messagebox.showinfo("提示", "预设内容不能为空。", parent=dialog)
                return
            index = selection[0]
            existing = list(preset_list.get(0, "end"))
            if value in existing and existing[index] != value:
                messagebox.showinfo("提示", "该预设已存在。", parent=dialog)
                return
            preset_list.delete(index)
            preset_list.insert(index, value)
            preset_list.selection_set(index)

        def delete_item():
            selection = preset_list.curselection()
            if selection:
                index = selection[0]
                preset_list.delete(index)
                value_var.set("")
                if preset_list.size():
                    next_index = min(index, preset_list.size() - 1)
                    preset_list.selection_set(next_index)
                    load_selection()

        def save_items():
            self.presets = normalize_presets(list(preset_list.get(0, "end")))
            settings = load_settings()
            settings["presets"] = self.presets
            save_settings(settings)
            dialog.destroy()
            self.status_var.set(f"已保存 {len(self.presets)} 个预设")

        preset_list.bind("<<ListboxSelect>>", load_selection)
        preset_list.bind("<Double-Button-1>", load_selection)
        entry.bind("<Return>", lambda _event: rename_item() if preset_list.curselection() else add_item())
        ttk.Button(editor, text="新增", command=add_item, width=8).grid(row=0, column=1, padx=(7, 0))
        ttk.Button(editor, text="保存修改", command=rename_item, width=9).grid(row=0, column=2, padx=(7, 0))
        ttk.Button(editor, text="删除选中", command=delete_item, width=9).grid(row=0, column=3, padx=(7, 0))

        bottom = ttk.Frame(body, style="Toolbar.TFrame")
        bottom.grid(row=5, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(bottom, text="取消", command=dialog.destroy, width=8).pack(side="left")
        ttk.Button(
            bottom, text="保存并关闭", command=save_items, width=10, style="Accent.TButton"
        ).pack(side="left", padx=(7, 0))
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.update_idletasks()
        width = max(dialog.winfo_reqwidth(), 570)
        height = max(dialog.winfo_reqheight(), 590)
        dialog.minsize(520, 540)
        center_window_on_parent(dialog, self.root, width, height)
        dialog.deiconify()
        dialog.grab_set()
        dialog.lift()
        entry.focus_set()

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
            from tender_info_extractor import extract_project_info_details
            from legacy_office import LegacyOfficeSession, temporary_docx_source

            if os.path.splitext(file_path)[1].lower() != ".doc":
                return extract_project_info_details(file_path)
            with LegacyOfficeSession() as session:
                with temporary_docx_source(file_path, session) as readable:
                    return extract_project_info_details(readable)

        def on_success(details):
            self._reset_busy_state()
            self._project_info_details = details
            self._project_info_source = file_path
            recognized = sum(bool(item["value"]) for item in details)
            self.rules_sheet.set_sheet_data([[item["placeholder"], item["value"]] for item in details])
            self._ensure_blank_rule_rows()
            self._refresh_rules_sheet_view()
            self._remember_word_files([file_path])
            self._remove_import_source_from_replace_files(file_path, "招标文件")
            self.status_var.set(f"已识别 {recognized}/{len(details)} 项：{os.path.basename(file_path)}；空白项不执行，可在“更多”核对来源")
            if not recognized or any(item["status"] == "冲突" for item in details):
                self.show_project_info_details()

        def on_error(exc):
            self._reset_busy_state()
            log_path = write_error_log(type(exc), exc, exc.__traceback__, context="读取招标文件")
            messagebox.showerror("错误", f"读取招标文件失败：{exc}\n\n错误日志：{log_path}")

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
                ("Office 文件", "*.doc;*.docx;*.xls;*.xlsx;*.xlsm;*.ppt;*.pptx"),
                ("Word 文档", "*.doc;*.docx"),
                ("Excel 工作簿", "*.xls;*.xlsx;*.xlsm"),
                ("PowerPoint 演示", "*.ppt;*.pptx"),
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

        self._apply_output_dir(directory)
        self.status_var.set("已选择输出目录")
        self._refresh_output_and_status_layout()

    def select_same_dir_replace(self):
        self._apply_output_dir(None, same_dir_replace=True)
        self.status_var.set("已选择同目录替换")
        self._refresh_output_and_status_layout()

    # ---------- 替换执行 ----------
    def start_replace(self):
        if not self.replace_files:
            messagebox.showwarning("提示", "请先选择待处理文件。")
            return
        try:
            validate_rules(self.rules_sheet.get_sheet_data())
        except ValueError as exc:
            messagebox.showwarning("规则冲突", str(exc))
            return
        rules = self.get_rules_from_table()
        if not rules:
            messagebox.showwarning("提示", "没有可执行规则。空白替换值会跳过，请填写替换值或明确标记删除。")
            return
        self._run_replace(list(self.replace_files), rules, self.output_dir)

    def _run_replace(self, file_paths, rules, output_dir):
        if self._task_runner.active:
            return
        self._begin_determinate_task(f"正在使用 {len(rules)} 条规则替换...", len(file_paths))

        def work(publish_progress):
            from string_replacer import batch_replace
            details = []
            results, error = batch_replace(file_paths, rules, output_dir=output_dir,
                                          progress_callback=publish_progress,
                                          result_callback=details.append)
            return results, error, details

        def on_progress(current, total, filename):
            self.progress.configure(maximum=total, value=current)
            self.status_var.set(f"正在处理 ({current}/{total})：{filename}")

        def on_success(result):
            results, error, details = result
            failed = [item["source_path"] for item in details if item["status"] == "失败"]
            retry = (lambda: self._run_replace(failed, rules, output_dir)) if failed else None
            self._show_result(results, error, file_paths=file_paths, output_dir=output_dir,
                              details=details, retry=retry)

        self._task_runner.submit(work, on_success,
                                lambda exc: self._handle_background_error(exc, "替换线程"), on_progress)

    def _reset_busy_state(self):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.configure(value=0)
        self.progress.grid_remove()
        self._set_busy_controls(False)
        self._refresh_output_and_status_layout()

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

    def _show_result(self, results, error, file_paths=None, output_dir=None, details=None, retry=None):
        self._reset_busy_state()
        state = "全部失败" if error and not results else "部分失败" if error else "完成"
        lines = [f"替换{state}。", f"成功处理 {len(results)} 个文件，替换 {sum(results.values())} 处。"]
        if details is not None:
            for item in details:
                lines.append(f"\n【{item['status']}】{item['source_path']}：{item['count']} 处")
                if item["output_path"]:
                    lines.append("输出：" + item["output_path"])
                if item["error"]:
                    lines.append("原因：" + item["error"])
        else:
            lines.extend(f"{filename}：{count} 处" for filename, count in results.items())
            if error:
                lines.append(error)
        self.status_var.set("替换" + state)
        self._last_output_dir_to_open = output_dir or self._default_output_dir_to_open(file_paths=file_paths)
        self._show_result_window("\n".join(lines), retry=retry)

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

    def _show_result_window(self, message, title="替换结果", retry=None):
        window = tk.Toplevel(self.root)
        window.title(title)
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
        scrollbar = FlatScrollbar(
            body, orient="vertical", command=text.yview, style="Flat.Vertical.TScrollbar"
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)
        self._bind_vertical_mousewheel(text, text, scrollbar)
        text.insert("1.0", message)
        text.configure(state="disabled")

        buttons = ttk.Frame(body, style="App.TFrame")
        buttons.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        def save_report():
            from file_io import atomic_output_path
            path = filedialog.asksaveasfilename(parent=window, title="保存结果", defaultextension=".txt", filetypes=[("文本报告", "*.txt")])
            if path:
                try:
                    with atomic_output_path(path) as temporary:
                        with open(temporary, "w", encoding="utf-8-sig") as report:
                            report.write(message)
                except OSError as exc:
                    messagebox.showerror("保存失败", str(exc), parent=window)

        self._create_busy_button(buttons, text="保存结果", command=save_report, width=10).pack(side="left", padx=(0, 8))
        if retry:
            def retry_failed():
                if not self._task_runner.active:
                    window.destroy()
                    retry()
            self._create_busy_button(buttons, text="仅重试失败", command=retry_failed, width=12).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="打开输出目录", command=self._open_output_dir, width=14).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="关闭", command=window.destroy, width=10).pack(side="left")

    def show_version_info(self):
        window = tk.Toplevel(self.root)
        window.title("版本信息")
        window.minsize(480, 280)
        window.configure(bg=self.app_bg)
        window.transient(self.root)
        # 先藏起来，排完布局再居中显示，避免用户看到从左上角跳出来
        window.withdraw()
        window.overrideredirect(False)

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
        scrollbar = FlatScrollbar(
            body, orient="vertical", command=text.yview, style="Flat.Vertical.TScrollbar"
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)
        self._bind_vertical_mousewheel(text, text, scrollbar)
        text.insert("1.0", format_changelog())
        text.configure(state="disabled")

        buttons = ttk.Frame(body, style="App.TFrame")
        buttons.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="关闭", command=window.destroy, width=10).pack()

        geometry = center_window_on_parent(window, self.root, width=560, height=360)
        window.deiconify()
        # 映射到屏幕后再写一次，防止首次显示时被系统重置到默认位置
        if geometry:
            window.geometry(geometry)
        else:
            center_window_on_parent(window, self.root, width=560, height=360)
        window.lift()
        window.focus_force()

    def open_word_table_exporter(self):
        initial_files = self._word_files_for_reuse()
        if self.table_export_window and self.table_export_window.exists():
            if initial_files:
                self.table_export_window._append_files(initial_files, show_status=False)
            self.table_export_window.focus()
            return
        self.table_export_window = WordTableExportWindow(self, initial_files=initial_files)


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
