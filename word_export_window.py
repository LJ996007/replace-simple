"""Word 表格与指标条款提取窗口。"""

import os
from dataclasses import replace
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from app_settings import load_settings, save_settings
from ui_common import (
    WORD_EXTENSIONS,
    DEFAULT_SYMBOL_CHARS,
    HEADER_GRID_COLOR,
    ElidedTextController,
    CanvasCheckbox,
    center_window_on_parent,
    write_error_log,
    _file_identity,
    collect_supported_files,
    BackgroundTaskRunner,
    cancel_widget_callbacks,
    FlatScrollbar,
)


class WordTableExportWindow:
    WINDOW_WIDTH = 1280
    WINDOW_HEIGHT = 1100
    MIN_WIDTH = 1020
    MIN_HEIGHT = 820
    SCAN_COLUMN_TITLES = {
        "selected": "导出",
        "type": "类型",
        "section": "所在章节",
        "item": "项目",
        "quantity": "数量",
        "hint": "提示",
    }
    FILTER_TITLES = {"file": "文件", "section": "章节", "type": "类型"}

    def __init__(self, app, initial_files=None):
        self.app = app
        self.window = tk.Toplevel(app.root)
        self.window.bind("<Destroy>", cancel_widget_callbacks, add="+")
        self._task_runner = BackgroundTaskRunner(self.window)
        self._busy_widgets = []
        self.window.title("提取信息")
        # Give the scan list enough room on first open.  The result list is the
        # primary workspace in this window, so it should not start out cramped
        # by the fixed-height sections around it.
        self.window.configure(bg=app.app_bg)
        # 不设置 transient：Windows 会为普通可调整大小的 Toplevel 提供最大化按钮。
        self.window.resizable(True, True)
        # 先藏起来，排完布局再相对主窗口居中显示，避免闪到屏幕左上角
        self.window.withdraw()
        width, height, min_w, min_h = self._fitted_window_size()
        self.window.minsize(min_w, min_h)
        try:
            self.window.iconphoto(True, app._icon_photo)
        except Exception:
            pass

        self.file_paths = []
        self.table_items = []
        self.symbol_items = []
        self.scan_rows = []
        self.scan_stamps = {}
        self.scan_item_by_iid = {}
        self.selected_scan_keys = set()
        self.scan_filters = {}
        self.keyword_var = tk.StringVar(value="")
        self.only_selected_var = tk.BooleanVar(value=False)
        self.output_dir = None
        self.keep_symbols_var = tk.BooleanVar(value=bool(app.keep_clause_symbols))
        self._last_output_dir_to_open = None
        self._scan_separator_after = None
        self._check_anchor_iid = None
        self._filter_popup = None
        self._filter_popup_column = None
        self._filter_popup_rows = None
        self._filter_click_bind = None
        self._filter_escape_bind = None
        self._file_group_labels = {}
        self._reopening_groups = False
        self._empty_message = "加上 Word 文件后，点「扫描表格和符号」"
        self._choice_widgets = []
        self.status_var = tk.StringVar(value="就绪")

        self._create_widgets()
        self._refresh_symbols_label()
        if initial_files:
            self._append_files(initial_files, show_status=False)
            self.status_var.set(f"已自动带入 {len(self.file_paths)} 个 Word 文件")
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        geometry = center_window_on_parent(
            self.window,
            app.root,
            width=width,
            height=height,
        )
        self.window.deiconify()
        if geometry:
            self.window.geometry(geometry)
            # Windows 上 withdraw 后再显示，偶发丢掉刚才写入的尺寸，idle 后再钉一次
            self.window.after_idle(lambda g=geometry: self._reapply_geometry(g))
        else:
            center_window_on_parent(
                self.window,
                app.root,
                width=width,
                height=height,
            )
        self.window.lift()
        self.window.focus_force()

    def _fitted_window_size(self):
        """按屏幕可用区域收敛默认尺寸，避免超出任务栏或小屏显示器。"""
        width, height = self.WINDOW_WIDTH, self.WINDOW_HEIGHT
        min_w, min_h = self.MIN_WIDTH, self.MIN_HEIGHT
        try:
            screen_w = int(self.window.winfo_screenwidth())
            screen_h = int(self.window.winfo_screenheight())
        except Exception:
            return width, height, min_w, min_h
        if screen_w > 1:
            width = min(width, max(screen_w - 48, 800))
            min_w = min(min_w, width)
        if screen_h > 1:
            height = min(height, max(screen_h - 88, 600))
            min_h = min(min_h, height)
        return width, height, min_w, min_h

    def _reapply_geometry(self, geometry):
        try:
            if self.window.winfo_exists() and geometry:
                self.window.geometry(geometry)
        except tk.TclError:
            pass

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
        if hasattr(self, "_output_path"):
            self._output_path.hide_tooltip()
        self.app.table_export_window = None
        self.window.destroy()

    def _create_busy_button(self, parent, **kwargs):
        button = ttk.Button(parent, **kwargs)
        self._busy_widgets.append(button)
        return button

    def _create_choice_button(self, parent, **kwargs):
        button = self._create_busy_button(parent, **kwargs)
        self._choice_widgets.append(button)
        return button

    def _track_choice_widget(self, widget):
        self._choice_widgets.append(widget)
        self._busy_widgets.append(widget)
        return widget

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
        self._refresh_constrained_texts()

    def _begin_determinate_task(self, status, maximum):
        self._set_busy_controls(True)
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=maximum, value=0)
        self.progress.grid()
        self.status_var.set(status)
        self._refresh_constrained_texts()

    def _create_widgets(self):
        body = ttk.Frame(self.window, padding=12, style="App.TFrame")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)

        header = ttk.Frame(body, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(header, text="提取信息", style="Title.TLabel").pack(anchor="w")
        self.header_hint = ttk.Label(
            header,
            text="选择 .doc / .docx 文件，扫描表格和符号后在列表里勾选，再导出。",
            style="Subtitle.TLabel",
        )
        self.header_hint.pack(anchor="w", pady=(2, 0))

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
        scrollbar = FlatScrollbar(
            list_frame,
            orient="vertical",
            command=self.file_listbox.yview,
            style="Flat.Vertical.TScrollbar",
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        self.app._bind_vertical_mousewheel(self.file_listbox, self.file_listbox, scrollbar)

        self.files_label = ttk.Label(file_frame, text="已选择 0 个 Word 文件", style="Muted.TLabel")
        self.files_label.grid(row=2, column=0, sticky="w", pady=(5, 0))

        scan_frame = ttk.Frame(body, padding=(12, 9), style="Surface.TFrame")
        scan_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        scan_frame.columnconfigure(0, weight=1)
        scan_frame.rowconfigure(3, weight=1)

        scan_header = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        scan_header.grid(row=0, column=0, sticky="ew")
        scan_header.columnconfigure(0, weight=1)
        ttk.Label(scan_header, text="选择要导出的内容", style="SectionTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )

        scan_toolbar = ttk.Frame(scan_header, style="Toolbar.TFrame")
        scan_toolbar.grid(row=0, column=1, sticky="e")
        self.symbols_label = ttk.Label(scan_toolbar, text="符号 ★ # △ ▲", style="Muted.TLabel")
        self.symbols_label.pack(side="left", padx=(0, 8))
        self.customize_symbols_button = self._create_busy_button(
            scan_toolbar,
            text="自定义符号…",
            command=self.open_symbol_settings,
            width=12,
        )
        self.customize_symbols_button.pack(side="left", padx=(0, 6))
        self.scan_button = self._create_busy_button(
            scan_toolbar,
            text="扫描表格和符号",
            command=self.scan_content,
            width=14,
            style="Accent.TButton",
        )
        self.scan_button.pack(side="left")

        self.scan_label = ttk.Label(scan_frame, text="尚未扫描", style="Muted.TLabel")
        self.scan_label.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        filters = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        filters.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        filters.columnconfigure(4, weight=1)
        self.filter_boxes = {}
        for column_index, column in enumerate(("file", "section", "type")):
            box = self._create_choice_button(
                filters,
                text=f"{self.FILTER_TITLES[column]} 全部 ▾",
                style="TButton",
                command=lambda c=column: self._open_multi_filter(c),
            )
            box.grid(row=0, column=column_index, sticky="w", padx=(0, 6))
            self.filter_boxes[column] = box
        ttk.Label(filters, text="关键词").grid(row=0, column=3, sticky="w", padx=(8, 4))
        self.keyword_entry = self._track_choice_widget(
            ttk.Entry(filters, textvariable=self.keyword_var, width=24)
        )
        self.keyword_entry.grid(row=0, column=4, sticky="ew")
        self.keyword_var.trace_add("write", lambda *_: self._apply_scan_filters())
        self.only_selected_check = self._track_choice_widget(
            ttk.Checkbutton(
                filters,
                text="只看已选",
                variable=self.only_selected_var,
                command=self._apply_scan_filters,
            )
        )
        self.only_selected_check.grid(row=0, column=5, sticky="w", padx=(8, 6))
        self.reset_filters_button = self._create_choice_button(
            filters, text="清除筛选", command=self.reset_scan_filters, style="TButton"
        )
        self.reset_filters_button.grid(row=0, column=6, sticky="e")

        content = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        content.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, minsize=380, weight=0)
        content.rowconfigure(0, weight=1)

        tree_frame = ttk.Frame(content, style="Toolbar.TFrame")
        tree_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.scan_tree = ttk.Treeview(
            tree_frame,
            columns=("selected", "type", "section", "item", "quantity", "hint"),
            show="tree headings",
            selectmode="extended",
            style="Scan.Treeview",
            height=18,
        )
        for column, title in self.SCAN_COLUMN_TITLES.items():
            self.scan_tree.heading(column, text=title)
        self.scan_tree.column("#0", width=28, minwidth=24, stretch=False)
        self.scan_tree.column("selected", width=56, minwidth=52, anchor="center", stretch=False)
        self.scan_tree.column("type", width=76, minwidth=64, anchor="center", stretch=False)
        self.scan_tree.column("section", width=200, minwidth=140)
        self.scan_tree.column("item", width=96, minwidth=72, anchor="center", stretch=False)
        self.scan_tree.column("quantity", width=118, minwidth=104, anchor="center", stretch=False)
        self.scan_tree.column("hint", width=76, minwidth=64, anchor="center", stretch=False)
        self.scan_tree.tag_configure("filegroup", font=self.app.section_font)
        self.scan_tree.grid(row=0, column=0, sticky="nsew")
        y_scrollbar = FlatScrollbar(
            tree_frame,
            orient="vertical",
            command=self.scan_tree.yview,
            style="Scan.Vertical.TScrollbar",
        )
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        self.scan_x_scrollbar = FlatScrollbar(
            tree_frame,
            orient="horizontal",
            command=self._scan_tree_xview,
            style="Scan.Horizontal.TScrollbar",
        )
        # 横向滚动条只在列宽超出可见区域时再出现，避免空表也占掉一行高度
        self.scan_tree.configure(
            yscrollcommand=y_scrollbar.set,
            xscrollcommand=self._on_scan_tree_xscroll,
        )
        self.app._bind_vertical_mousewheel(self.scan_tree, self.scan_tree, y_scrollbar)
        # 挂到 tree_frame 上，才能盖住表头区域（Treeview 子控件会被表头层盖住）
        self._scan_column_separators = [
            tk.Frame(
                tree_frame,
                width=1,
                bg=HEADER_GRID_COLOR,
                borderwidth=0,
                highlightthickness=0,
                takefocus=False,
            )
            for _column in self.scan_tree.cget("columns")[:-1]
        ]
        # 展开列和导出列之间的竖线。没有它时，导出方框会看起来偏在整列一侧。
        self._scan_gutter_separator = tk.Frame(
            tree_frame,
            width=1,
            bg=HEADER_GRID_COLOR,
            borderwidth=0,
            highlightthickness=0,
            takefocus=False,
        )
        # 表头底部分隔横线，与竖线同色，补齐表头与数据区边界
        self._scan_header_bottom_line = tk.Frame(
            tree_frame,
            height=1,
            bg=HEADER_GRID_COLOR,
            borderwidth=0,
            highlightthickness=0,
            takefocus=False,
        )
        self.scan_tree.bind("<Button-1>", self._toggle_scan_checkbox_from_click)
        self.scan_tree.bind("<ButtonRelease-1>", self._schedule_scan_tree_column_separators, add="+")
        self.scan_tree.bind("<B1-Motion>", self._schedule_scan_tree_column_separators, add="+")
        self.scan_tree.bind("<Configure>", self._schedule_scan_tree_column_separators, add="+")
        self.scan_tree.bind("<space>", self._toggle_scan_rows_from_keyboard)
        self.scan_tree.bind("<<TreeviewSelect>>", self._update_scan_detail)
        self.scan_tree.bind("<<TreeviewClose>>", self._reopen_file_groups, add="+")
        self._schedule_scan_tree_column_separators()
        self.scan_placeholder = None

        preview = ttk.Frame(content, style="Toolbar.TFrame")
        preview.grid(row=0, column=1, sticky="nsew")
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(2, weight=1)
        ttk.Label(preview, text="预览", style="SectionTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.detail_caption = ttk.Label(
            preview,
            text="选择一行后，这里按表格显示内容。",
            style="Muted.TLabel",
            justify="left",
            anchor="w",
        )
        self.detail_caption.grid(row=1, column=0, sticky="ew", pady=(2, 4))
        self.detail_frame = ttk.Frame(preview, style="Toolbar.TFrame")
        self.detail_frame.grid(row=2, column=0, sticky="nsew")
        self.detail_frame.columnconfigure(0, weight=1)
        self.detail_frame.rowconfigure(0, weight=1)
        self.detail_canvas = tk.Canvas(
            self.detail_frame,
            bg="#FFFFFF",
            highlightthickness=1,
            highlightbackground=self.app.border_color,
            borderwidth=0,
        )
        self.detail_canvas.grid(row=0, column=0, sticky="nsew")
        self.detail_scrollbar = FlatScrollbar(
            self.detail_frame,
            orient="vertical",
            command=self.detail_canvas.yview,
            style="Flat.Vertical.TScrollbar",
        )
        self.detail_scrollbar.grid(row=0, column=1, sticky="ns")
        self.detail_x_scrollbar = FlatScrollbar(
            self.detail_frame,
            orient="horizontal",
            command=self.detail_canvas.xview,
            style="Flat.Horizontal.TScrollbar",
        )
        self.detail_canvas.configure(
            yscrollcommand=self.detail_scrollbar.set,
            xscrollcommand=self._on_preview_xscroll,
        )
        self.detail_table = tk.Frame(self.detail_canvas, bg="#D5DBE3")
        self._detail_window = self.detail_canvas.create_window(
            (0, 0), window=self.detail_table, anchor="nw"
        )
        self.detail_table.bind("<Configure>", self._refresh_preview_scroll)
        self.detail_canvas.bind("<Configure>", self._refresh_preview_scroll)
        self._bind_preview_wheel(self.detail_canvas)
        self._bind_preview_wheel(self.detail_table)

        self.output_frame = ttk.Frame(body, padding=(12, 6), style="Surface.TFrame")
        output_frame = self.output_frame
        output_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        output_frame.columnconfigure(1, weight=1)

        # 单行：标题 + 路径说明 + 选择按钮；长路径中间省略，保持一行高度
        ttk.Label(output_frame, text="输出目录", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.output_label = ttk.Label(
            output_frame,
            text="未选择则输出到原 Word 文件所在目录",
            style="Muted.TLabel",
            justify="left",
        )
        self.output_label.grid(row=0, column=1, sticky="ew", padx=(10, 10))
        self._create_busy_button(output_frame, text="选择目录", command=self.select_output_dir, width=12).grid(
            row=0, column=2, sticky="e"
        )
        self._output_path = ElidedTextController(
            self.output_label,
            self.app.body_font,
            self._output_path_width,
            max_lines=1,
            tooltip=True,
            tooltip_font=self.app.small_font,
            tooltip_fg=self.app.text_fg,
            fallback_width=700,
        )
        output_frame.bind("<Configure>", self._on_output_frame_configure)

        self.action_frame = ttk.Frame(body, padding=(12, 9), style="Surface.TFrame")
        action_frame = self.action_frame
        action_frame.grid(row=4, column=0, sticky="ew")
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(3, minsize=118)

        self.status_label = ttk.Label(action_frame, text="就绪", style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="ew")
        self.progress = ttk.Progressbar(action_frame, mode="determinate", length=180)
        self.progress.grid(row=0, column=1, sticky="ew", padx=(14, 12))
        self.progress.grid_remove()

        self.keep_symbols_check = CanvasCheckbox(
            action_frame,
            "条款内容中保留符号",
            self.keep_symbols_var,
            self.app,
            command=self._on_keep_symbols_toggled,
            padx=6,
            pady=2,
        )
        self.keep_symbols_check.canvas.grid(row=0, column=2, sticky="e", padx=(0, 8))

        self.export_button = self._create_busy_button(
            action_frame,
            text="导出已选 0 项",
            command=self.start_export,
            width=14,
            style="Accent.TButton",
        )
        self.export_button.grid(row=0, column=3, sticky="e")
        self._refresh_scan_filter_headings()
        self._sync_scan_choice_state()
        self._status_elide = ElidedTextController(
            self.status_label,
            self.app.body_font,
            self._status_width,
            max_lines=1,
            fallback_width=400,
        ).attach_var(self.status_var)
        self.window.bind("<Configure>", self._on_window_configure)
        self.window.after_idle(self._refresh_constrained_texts)

    def _scan_tree_xview(self, *args):
        self.scan_tree.xview(*args)
        self._schedule_scan_tree_column_separators()

    def _on_scan_tree_xscroll(self, first, last):
        self.scan_x_scrollbar.set(first, last)
        self._sync_scan_tree_x_scrollbar(first, last)
        self._schedule_scan_tree_column_separators()

    def _sync_scan_tree_x_scrollbar(self, first, last):
        """列宽超出可见区域才显示横向滚动条，把垂直空间留给表格行。"""
        try:
            overflow = float(first) > 0.001 or float(last) < 0.999
        except (TypeError, ValueError):
            overflow = True
        try:
            shown = bool(self.scan_x_scrollbar.grid_info())
        except tk.TclError:
            return
        if overflow and not shown:
            self.scan_x_scrollbar.grid(row=1, column=0, sticky="ew")
        elif not overflow and shown:
            self.scan_x_scrollbar.grid_remove()

    def _refresh_scan_filter_headings(self):
        if not hasattr(self, "scan_tree"):
            return
        for column, title in self.SCAN_COLUMN_TITLES.items():
            if column == "selected":
                self.scan_tree.heading(column, text=self._visible_export_mark(), anchor="center")
                continue
            marker = " [筛]" if column in self.scan_filters else ""
            self.scan_tree.heading(column, text=title + marker)
        for column, box in self.filter_boxes.items():
            box.configure(text=self._filter_button_text(column))

    def _filter_button_text(self, column):
        title = self.FILTER_TITLES[column]
        chosen = self.scan_filters.get(column)
        if not chosen:
            return f"{title} 全部 ▾"
        if len(chosen) == 1:
            value = next(iter(chosen))
            if column == "file":
                value = os.path.basename(value)
            return f"{title} {self._short_filter_label(value)} ▾"
        return f"{title} {len(chosen)} ▾"

    @staticmethod
    def _short_filter_label(text, limit=14):
        text = " ".join(str(text).split())
        if len(text) <= limit:
            return text
        return text[: limit - 1] + "…"

    def _filter_choice_pairs(self, column):
        values = list(dict.fromkeys(
            self._scan_filter_value(column, kind, item) for kind, item in self.scan_rows
        ))
        if column != "file":
            return values, values
        labels = [os.path.basename(path) for path in values]
        if len(labels) != len(set(labels)):
            labels = list(values)
        return values, labels

    def _open_multi_filter(self, column):
        if self._filter_popup_column == column and self._filter_popup_exists():
            self._close_filter_dropdown()
            return
        self._close_filter_dropdown()
        values, labels = self._filter_choice_pairs(column)
        rows = [(None, "全部"), *zip(values, labels)]
        button = self.filter_boxes[column]
        font = self.app.body_font
        text_width = max((font.measure(label) for _value, label in rows), default=80)
        menu_width = min(max(text_width + 44, button.winfo_width(), 220), 480)
        row_height = max(font.metrics("linespace") + 10, 28)
        visible_rows = min(10, len(rows))

        popup = tk.Toplevel(self.window)
        popup.withdraw()
        popup.overrideredirect(True)
        popup.transient(self.window)
        popup.configure(bg=self.app.border_color)
        outer = tk.Frame(popup, bg=self.app.border_color, padx=1, pady=1)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(
            outer,
            width=menu_width,
            height=visible_rows * row_height,
            bg=self.app.surface_bg,
            highlightthickness=0,
            borderwidth=0,
        )
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar = FlatScrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        inner = tk.Frame(canvas, bg=self.app.surface_bg)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        self._filter_popup = popup
        self._filter_popup_column = column
        self._filter_popup_rows = []
        available = list(values)
        for index, (value, label) in enumerate(rows):
            if index == 1:
                tk.Frame(inner, bg=self.app.border_color, height=1).pack(fill="x")
            row = tk.Frame(inner, bg=self.app.surface_bg, height=row_height, cursor="hand2")
            row.pack(fill="x")
            row.pack_propagate(False)
            mark = tk.Label(
                row,
                text=self._filter_choice_mark(column, value, available),
                bg=self.app.surface_bg,
                fg=self.app.text_fg,
                font=font,
                width=2,
                anchor="center",
            )
            mark.pack(side="left", padx=(6, 0))
            shown = self._fit_filter_menu_label(label, menu_width - 40)
            text = tk.Label(
                row,
                text=shown,
                bg=self.app.surface_bg,
                fg=self.app.text_fg,
                font=font,
                anchor="w",
            )
            text.pack(side="left", fill="x", expand=True, padx=(2, 8))
            for widget in (row, mark, text):
                widget.bind("<Button-1>", lambda _event, item=value: self._toggle_filter_choice(column, item, available))
                widget.bind("<Enter>", lambda _event, target=row: self._set_filter_row_hover(target, True))
                widget.bind("<Leave>", lambda _event, target=row: self._set_filter_row_hover(target, False))
            self._filter_popup_rows.append((value, mark))

        def sync_width(event):
            canvas.itemconfigure(window_id, width=event.width)

        def sync_scroll(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind("<Configure>", sync_width)
        inner.bind("<Configure>", sync_scroll)
        popup.bind("<MouseWheel>", lambda event: self._scroll_filter_dropdown(canvas, event))
        popup.bind("<Escape>", lambda _event: self._close_filter_dropdown())
        self._filter_escape_bind = self.window.bind(
            "<Escape>", lambda _event: self._close_filter_dropdown(), add="+"
        )
        if len(rows) > visible_rows:
            scrollbar.pack(side="right", fill="y")

        popup.update_idletasks()
        width = max(popup.winfo_reqwidth(), menu_width + 2)
        height = popup.winfo_reqheight()
        x = button.winfo_rootx()
        y = button.winfo_rooty() + button.winfo_height()
        screen_w = popup.winfo_screenwidth()
        screen_h = popup.winfo_screenheight()
        if y + height > screen_h - 8:
            y = max(8, button.winfo_rooty() - height)
        if x + width > screen_w - 8:
            x = max(8, screen_w - width - 8)
        popup.geometry(f"{width}x{height}+{x}+{y}")
        popup.deiconify()
        popup.lift()
        popup.update_idletasks()
        popup.focus_set()
        self._filter_click_bind = self.window.bind(
            "<Button-1>", self._on_filter_dropdown_click, add="+"
        )

    def _filter_popup_exists(self):
        popup = self._filter_popup
        if popup is None:
            return False
        try:
            return bool(popup.winfo_exists())
        except tk.TclError:
            return False

    def _fit_filter_menu_label(self, text, width):
        font = self.app.body_font
        if font.measure(text) <= width:
            return text
        ellipsis = "…"
        limit = width
        count = len(text)
        while count > 0 and font.measure(text[:count] + ellipsis) > limit:
            count -= 1
        return text[:count] + ellipsis

    def _filter_choice_mark(self, column, value, available):
        chosen = self.scan_filters.get(column)
        selected = set(available if chosen is None else chosen)
        if value is None:
            if available and selected >= set(available):
                return "☑"
            if selected:
                return "◩"
            return "☐"
        return "☑" if value in selected else "☐"

    def _set_filter_row_hover(self, row, hover):
        color = "#EDF1F6" if hover else self.app.surface_bg
        row.configure(bg=color)
        for child in row.winfo_children():
            try:
                child.configure(bg=color)
            except tk.TclError:
                pass

    def _scroll_filter_dropdown(self, canvas, event):
        canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def _toggle_filter_choice(self, column, value, available):
        chosen = self.scan_filters.get(column)
        selected = set(available if chosen is None else chosen)
        if column == "file":
            if value is None or (chosen is not None and value in selected):
                self.scan_filters.pop(column, None)
                self._apply_scan_filters(select_first=True)
            else:
                self._set_multi_filter(column, {value})
            self._close_filter_dropdown()
            return
        if value is None:
            selected = set() if selected >= set(available) else set(available)
        elif value in selected:
            selected.remove(value)
        else:
            selected.add(value)
        self._set_multi_filter(column, selected)
        self._refresh_filter_dropdown_marks(available)

    def _refresh_filter_dropdown_marks(self, available):
        if not self._filter_popup_exists():
            return
        column = self._filter_popup_column
        for value, mark in self._filter_popup_rows:
            mark.configure(text=self._filter_choice_mark(column, value, available))

    def _on_filter_dropdown_click(self, event):
        if not self._filter_popup_exists():
            return
        widget = event.widget
        try:
            if widget.winfo_toplevel() is self._filter_popup:
                return
        except tk.TclError:
            return
        button = self.filter_boxes.get(self._filter_popup_column)
        if widget is button:
            return
        self._close_filter_dropdown()

    def _close_filter_dropdown(self):
        bind_id = self._filter_click_bind
        self._filter_click_bind = None
        if bind_id:
            try:
                self.window.unbind("<Button-1>", bind_id)
            except tk.TclError:
                pass
        escape_bind = self._filter_escape_bind
        self._filter_escape_bind = None
        if escape_bind:
            try:
                self.window.unbind("<Escape>", escape_bind)
            except tk.TclError:
                pass
        popup = self._filter_popup
        self._filter_popup = None
        self._filter_popup_column = None
        self._filter_popup_rows = None
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass

    def _set_multi_filter(self, column, chosen):
        available = {self._scan_filter_value(column, kind, item) for kind, item in self.scan_rows}
        if chosen == available:
            self.scan_filters.pop(column, None)
        else:
            self.scan_filters[column] = set(chosen)
        self._apply_scan_filters(select_first=True)

    def reset_scan_filters(self):
        self._close_filter_dropdown()
        self.scan_filters.clear()
        self.only_selected_var.set(False)
        self.keyword_var.set("")
        self._apply_scan_filters(select_first=True)

    def _scan_filter_value(self, column, kind, item):
        if column == "file":
            return item.file_path
        if column == "type":
            return "表格" if kind == "table" else "符号条款"
        if column == "section":
            return item.section
        return ""

    def _scan_row_matches_filters(self, kind, item):
        if self.only_selected_var.get() and not self._row_keys(kind, item) & self.selected_scan_keys:
            return False
        keyword = self.keyword_var.get().strip().casefold()
        if keyword and keyword not in self._search_text(kind, item).casefold():
            return False
        for column, selected_values in self.scan_filters.items():
            if self._scan_filter_value(column, kind, item) not in selected_values:
                return False
        return True

    def _search_text(self, kind, item):
        return f"{item.file_path}\n{item.section}\n{self._preview_content(kind, item)}"

    def _preview_content(self, kind, item):
        if kind == "table" and item.table is not None:
            rows = {}
            for cell in item.table.cells:
                rows.setdefault(cell.row, []).append(f"[列{cell.column}] {cell.text}")
            return "\n".join(f"第{row}行：" + " | ".join(cells) for row, cells in rows.items())
        parts = []
        for clause in getattr(item, "clauses", ()):
            if getattr(clause, "columns", ()):
                parts.append("\n".join(value for value in clause.columns if value))
            elif clause.text:
                parts.append(clause.text)
        return "\n\n".join(parts) or item.preview

    def _clause_items(self, item):
        return [replace(item, clauses=(clause,), clause_count=1, symbols=(clause.symbol,),
                        sources=(clause.source,), preview=clause.text) for clause in item.clauses]

    def _row_keys(self, kind, item):
        if kind == "symbol" and item.clauses:
            return {self._scan_key("clause", child) for child in self._clause_items(item)}
        return {self._scan_key(kind, item)}

    def _schedule_scan_tree_column_separators(self, event=None):
        if self._scan_separator_after is not None:
            return
        try:
            self._scan_separator_after = self.window.after_idle(self._run_scan_tree_column_separator_update)
        except tk.TclError:
            pass

    def _run_scan_tree_column_separator_update(self):
        self._scan_separator_after = None
        self._update_scan_tree_column_separators()

    def _scan_tree_header_height(self):
        """测量 Treeview 表头高度，用于底部分隔横线贴齐灰底下沿。"""
        try:
            for iid in self.scan_tree.get_children():
                bbox = self.scan_tree.bbox(iid)
                if bbox:
                    return max(1, int(bbox[1]))
        except tk.TclError:
            pass

        # 空表无 bbox：用 identify_region 实测 heading 区域下沿（比字体估算准）
        try:
            tree_height = self.scan_tree.winfo_height()
            if tree_height > 1:
                probe_x = max(2, min(20, self.scan_tree.winfo_width() // 4 or 2))
                last_heading_y = -1
                for y in range(0, min(tree_height, 96)):
                    region = self.scan_tree.identify_region(probe_x, y)
                    if region == "heading":
                        last_heading_y = y
                    elif last_heading_y >= 0:
                        break
                if last_heading_y >= 0:
                    return last_heading_y + 1
        except tk.TclError:
            pass

        try:
            # 兜底：section 字体行高 + Heading padding(8,8) + 边框余量
            linespace = int(self.app.section_font.metrics("linespace"))
            return max(36, linespace + 22)
        except Exception:
            return 41

    def _update_scan_tree_column_separators(self, event=None):
        """Keep header separator overlays aligned while columns resize or scroll."""
        try:
            columns = tuple(self.scan_tree.cget("columns"))
            widths = [int(self.scan_tree.column(column, "width")) for column in columns]
            tree_width = int(self.scan_tree.column("#0", "width"))
            total_width = sum(widths) + tree_width
            if total_width <= 0:
                return

            scroll_offset = float(self.scan_tree.xview()[0]) * total_width
            visible_width = self.scan_tree.winfo_width()
            tree_height = self.scan_tree.winfo_height()
            if visible_width <= 1 or tree_height <= 1:
                return

            # tree_frame 坐标系：竖线贯穿表头和数据区；<B1-Motion> 会在
            # 拖动列宽期间实时重算位置，避免旧线残留在内容中间。
            origin_x = self.scan_tree.winfo_x()
            origin_y = self.scan_tree.winfo_y()
            header_height = self._scan_tree_header_height()
            boundary = tree_width
            gutter_x = round(boundary - scroll_offset)
            if 1 < gutter_x < visible_width - 1:
                self._scan_gutter_separator.place(
                    x=origin_x + gutter_x - 1,
                    y=origin_y,
                    width=1,
                    height=tree_height,
                )
                self._scan_gutter_separator.lift()
            else:
                self._scan_gutter_separator.place_forget()
            for separator, width in zip(self._scan_column_separators, widths[:-1]):
                boundary += width
                x = round(boundary - scroll_offset)
                if 1 < x < visible_width - 1:
                    separator.place(
                        x=origin_x + x - 1,
                        y=origin_y,
                        width=1,
                        height=tree_height,
                    )
                    separator.lift()
                else:
                    separator.place_forget()

            # 表头底部分隔横线：横跨可见表头宽度
            if header_height < tree_height:
                self._scan_header_bottom_line.place(
                    x=origin_x,
                    y=origin_y + header_height - 1,
                    width=visible_width,
                    height=1,
                )
                self._scan_header_bottom_line.lift()
            else:
                self._scan_header_bottom_line.place_forget()
        except (tk.TclError, ValueError, AttributeError):
            pass

    def select_files(self):
        file_paths = filedialog.askopenfilenames(
            title="选择 Word 文件",
            filetypes=[("Word 文档", "*.doc;*.docx"), ("所有文件", "*.*")],
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
            return collect_supported_files(directory, WORD_EXTENSIONS, publish_progress)

        def on_progress(scanned_count):
            self.status_var.set(f"正在扫描文件夹：已检查 {scanned_count} 个文件...")

        def on_success(collected):
            self._reset_busy_state()
            if not collected:
                messagebox.showinfo("提示", "该文件夹中未找到 .doc 或 .docx 文件。", parent=self.window)
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
            self._clear_scan_results("文件列表已变化，请重新扫描表格和符号")
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
            and os.path.splitext(file_path)[1].lower() in WORD_EXTENSIONS
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
        self._clear_scan_results("文件列表已变化，请重新扫描表格和符号")
        self.status_var.set(f"已移除 {len(selected)} 个文件")

    def clear_files(self):
        if not self.file_paths:
            return
        self.file_paths = []
        self._refresh_file_list()
        self._clear_scan_results("请先添加 Word 文件并扫描表格和符号")
        self.status_var.set("Word 文件列表已清空")

    def scan_content(self):
        if not self.file_paths:
            messagebox.showwarning("提示", "请先选择 Word 文件。", parent=self.window)
            return

        file_paths = list(self.file_paths)
        symbol_chars = self.app.symbol_chars
        self._begin_determinate_task(f"正在扫描 {len(file_paths)} 个 Word 文件...", len(file_paths))

        def work(publish_progress):
            from word_scan import scan_word_content
            return scan_word_content(file_paths, symbol_chars, publish_progress)

        def on_success(result):
            table_result, symbol_result, self.scan_stamps = result
            self._show_scan_result(*table_result, *symbol_result)

        def on_error(exc):
            self._finish_background_error(exc, "Word 表格和符号扫描线程")

        self._task_runner.submit(work, on_success, on_error, self._update_progress)

    def _show_scan_result(
        self,
        table_items,
        table_skipped,
        table_error,
        symbol_items,
        symbol_skipped,
        symbol_error,
    ):
        self.table_items = list(table_items)
        self.symbol_items = list(symbol_items)
        self.selected_scan_keys = set()
        self.scan_filters = {}
        self.only_selected_var.set(False)
        self.keyword_var.set("")

        tables_by_file = {}
        symbols_by_file = {}
        for item in self.table_items:
            tables_by_file.setdefault(_file_identity(item.file_path), []).append(item)
        for item in self.symbol_items:
            symbols_by_file.setdefault(_file_identity(item.file_path), []).append(item)

        ordered_rows = []
        for file_path in self.file_paths:
            identity = _file_identity(file_path)
            ordered_rows.extend(("table", item) for item in tables_by_file.get(identity, []))
            ordered_rows.extend(("symbol", item) for item in symbols_by_file.get(identity, []))

        self.scan_rows = ordered_rows
        if not ordered_rows:
            self._empty_message = "没有扫描到可导出的正文表格或带符号条款。"
        self._apply_scan_filters(select_first=True)

        self._reset_busy_state()
        message = (
            f"已扫描到 {len(self.table_items)} 张表格、{len(self.symbol_items)} 个符号章节，"
            "当前选择 0 项"
        )
        if table_error or symbol_error:
            message += "；部分文件失败"
        self.status_var.set("扫描完成" if not (table_error or symbol_error) else "扫描完成（部分失败）")

        if table_skipped or symbol_skipped or table_error or symbol_error:
            lines = ["扫描完成。", "", message]
            if table_skipped:
                lines.extend(["", "表格扫描提示："])
                for filename, reason in table_skipped.items():
                    lines.append(f"  {filename}: {reason}")
            if symbol_skipped:
                lines.extend(["", "符号扫描提示："])
                for filename, reason in symbol_skipped.items():
                    lines.append(f"  {filename}: {reason}")
            if table_error or symbol_error:
                lines.extend(["", "失败文件："])
                if table_error:
                    lines.append(table_error)
                if symbol_error:
                    lines.append(symbol_error)
            self._show_result_window("\n".join(lines), title="扫描结果")

    def _clear_scan_results(self, label_text):
        self._close_filter_dropdown()
        self._empty_message = label_text
        self._check_anchor_iid = None
        self.table_items = []
        self.symbol_items = []
        self.scan_rows = []
        self.scan_stamps = {}
        self.scan_item_by_iid = {}
        self._file_group_labels = {}
        self.selected_scan_keys = set()
        self.scan_filters = {}
        self.only_selected_var.set(False)
        self.keyword_var.set("")
        if hasattr(self, "scan_tree"):
            self.scan_tree.delete(*self.scan_tree.get_children())
            self._refresh_scan_filter_headings()
            self._schedule_scan_tree_column_separators()
        self._refresh_scan_label()
        if hasattr(self, "detail_caption"):
            self._set_detail_text("选择一行后，这里按表格显示内容。")

    def _hint_short(self, kind, item):
        if kind != "table":
            return ""
        hint = getattr(item, "hint", "") or ""
        if hint.startswith("建议关注"):
            return "建议关注"
        if hint.startswith("谨慎选择"):
            return "谨慎选择"
        if hint:
            return "未命中"
        return ""

    def _scan_tree_values(self, kind, item):
        keys = self._row_keys(kind, item)
        checked = keys & self.selected_scan_keys
        if kind == "table":
            item_text = f"表格{item.table_index}"
            quantity = f"{item.row_count}x{item.column_count}"
            type_text = "表格"
        else:
            item_text = " ".join(item.symbols)
            type_text = "指标条款" if kind == "clause" else "符号条款"
            if kind == "symbol" and item.clauses:
                quantity = f"{item.clause_count}条 · 已选 {len(checked)}"
            else:
                quantity = f"{item.clause_count}条"
        return (
            "☑" if checked == keys else "◩" if checked else "☐",
            type_text,
            item.section,
            item_text,
            quantity,
            self._hint_short(kind, item),
        )

    def _scan_key(self, kind, item):
        if kind == "table":
            return (kind, _file_identity(item.file_path), item.table_index)
        if kind == "clause":
            return (kind, _file_identity(item.file_path), item.section, id(item.clauses[0]))
        return (kind, _file_identity(item.file_path), item.section)

    def _apply_scan_filters(self, select_first=False):
        if not hasattr(self, "scan_tree"):
            return
        previous_selection = self.scan_tree.selection()
        opened = {iid for iid in self.scan_item_by_iid if self.scan_tree.item(iid, "open")}
        scroll = self.scan_tree.yview()
        focus_key = None
        current_iid = self.scan_tree.focus() if hasattr(self, "scan_tree") else ""
        current_row = self.scan_item_by_iid.get(current_iid)
        if current_row is not None:
            focus_key = self._scan_key(*current_row)

        self.scan_tree.delete(*self.scan_tree.get_children())
        self.scan_item_by_iid = {}
        self._file_group_labels = {}
        group_files = len({_file_identity(item.file_path) for _, item in self.scan_rows}) >= 2
        group_parent = {}
        group_counts = {}

        def parent_for(item):
            if not group_files:
                return ""
            identity = _file_identity(item.file_path)
            if identity not in group_parent:
                gid = f"file:{len(group_parent)}"
                group_parent[identity] = gid
                group_counts[gid] = 0
                self._file_group_labels[gid] = item.filename
                self.scan_tree.insert(
                    "",
                    "end",
                    iid=gid,
                    open=True,
                    tags=("filegroup",),
                    values=("", "", item.filename, "", "", ""),
                )
            return group_parent[identity]

        focus_iid = None
        keyword_open = bool(self.keyword_var.get().strip())
        for row_index, (kind, item) in enumerate(self.scan_rows, start=1):
            children = [(index, child) for index, child in enumerate(self._clause_items(item), 1)
                        if self._scan_row_matches_filters("clause", child)] if kind == "symbol" and item.clauses else []
            if kind == "symbol" and item.clauses:
                matches = bool(children)
            else:
                matches = self._scan_row_matches_filters(kind, item)
            if not matches:
                continue
            iid = str(row_index)
            parent = parent_for(item)
            if parent:
                group_counts[parent] += 1
            key = self._scan_key(kind, item)
            self.scan_item_by_iid[iid] = (kind, item)
            self.scan_tree.insert(
                parent,
                "end",
                iid=iid,
                open=iid in opened or keyword_open,
                values=self._scan_tree_values(kind, item),
            )
            if key == focus_key:
                focus_iid = iid
            for index, child in children:
                child_iid = f"{iid}.{index}"
                self.scan_item_by_iid[child_iid] = ("clause", child)
                self.scan_tree.insert(
                    iid,
                    "end",
                    iid=child_iid,
                    values=self._scan_tree_values("clause", child),
                )
                if self._scan_key("clause", child) == focus_key:
                    focus_iid = child_iid

        for gid, count in group_counts.items():
            values = list(self.scan_tree.item(gid, "values"))
            values[4] = f"{count} 项"
            self.scan_tree.item(gid, values=values)

        children = self.scan_tree.get_children()
        target_iid = focus_iid or (self._first_checkable_iid() if select_first else None)
        retained = [iid for iid in previous_selection if iid in self.scan_item_by_iid]
        if retained:
            self.scan_tree.selection_set(retained)
        elif target_iid:
            self.scan_tree.selection_set(target_iid)
        if target_iid:
            self.scan_tree.focus(target_iid)
        if scroll:
            self.scan_tree.yview_moveto(scroll[0])
        if children:
            self._update_scan_detail()
        elif self.scan_rows:
            self._set_detail_text("当前筛选条件下没有扫描结果，请调整上方筛选条件。")
        else:
            self._set_detail_text("没有扫描到可导出的正文表格或带符号条款。")

        self._refresh_scan_filter_headings()
        self._refresh_scan_label()
        self._schedule_scan_tree_column_separators()

    def _iter_visible_checkable_iids(self, parent=""):
        for iid in self.scan_tree.get_children(parent):
            if iid in self.scan_item_by_iid:
                yield iid
            yield from self._iter_visible_checkable_iids(iid)

    def _first_checkable_iid(self):
        return next(self._iter_visible_checkable_iids(), None)

    def _refresh_scan_placeholder(self):
        return

    def _highlighted_count(self):
        try:
            selected = self.scan_tree.selection()
        except tk.TclError:
            return 0
        return sum(1 for iid in selected if iid in self.scan_item_by_iid)

    def _visible_export_mark(self):
        visible = self._visible_scan_keys()
        if not visible or not (visible & self.selected_scan_keys):
            return "☐"
        if visible <= self.selected_scan_keys:
            return "☑"
        return "◩"

    def _refresh_highlight_label(self):
        self._refresh_scan_label()

    def _sync_scan_choice_state(self):
        if self._task_runner.active:
            return
        state = "normal" if self.scan_rows else "disabled"
        for widget in self._choice_widgets:
            try:
                if widget.winfo_exists():
                    widget.configure(state=state)
            except tk.TclError:
                continue

    def _reopen_file_groups(self, event=None):
        if self._reopening_groups:
            return
        self._reopening_groups = True
        try:
            for iid in list(self._file_group_labels):
                try:
                    if self.scan_tree.exists(iid) and not self.scan_tree.item(iid, "open"):
                        self.scan_tree.item(iid, open=True)
                except tk.TclError:
                    continue
        finally:
            self._reopening_groups = False

    def _refresh_scan_label(self):
        if not hasattr(self, "scan_label"):
            return
        if not self.scan_rows:
            self.scan_label.config(text="尚未扫描", foreground=self.app.muted_fg)
        else:
            selected = len(self.selected_scan_keys)
            visible_keys = self._visible_scan_keys()
            hidden = len(self.selected_scan_keys - visible_keys)
            text = (
                f"{len(self.table_items)} 张表 · {len(self.symbol_items)} 个符号章节"
                f" · 正在显示 {len(visible_keys)} 项 · 已选 {selected} 项"
            )
            if hidden:
                text += f" · 另有 {hidden} 项已选但被筛掉"
            highlighted = self._highlighted_count()
            if highlighted:
                text += f" · 已高亮 {highlighted} 行，空格切换"
            self.scan_label.config(
                text=text,
                foreground="green" if selected else self.app.muted_fg,
            )
        self._refresh_visible_selection_button()
        self._sync_scan_choice_state()
        self._refresh_scan_placeholder()

    def _toggle_visible_from_heading(self):
        visible = self._visible_scan_keys()
        if not visible:
            return
        if visible <= self.selected_scan_keys:
            self.clear_scan_selection()
        else:
            self.select_all_scan_items()

    def _toggle_scan_checkbox_from_click(self, event):
        region = self.scan_tree.identify_region(event.x, event.y)
        if region == "heading":
            if self.scan_tree.identify_column(event.x) == "#1":
                self._toggle_visible_from_heading()
                return "break"
            return None
        if region != "cell":
            return None
        if self.scan_tree.identify_column(event.x) != "#1":
            return None
        row_id = self.scan_tree.identify_row(event.y)
        if not row_id or row_id not in self.scan_item_by_iid:
            return "break"
        # 0x0001 是 Shift。点方框只改导出勾选，不冲掉已经圈好的高亮。
        shift = bool(getattr(event, "state", 0) & 0x0001)
        self.scan_tree.focus(row_id)
        if shift and self._check_anchor_iid in self.scan_item_by_iid:
            self._set_check_range(self._check_anchor_iid, row_id)
        else:
            self._toggle_scan_iids([row_id])
            self._check_anchor_iid = row_id
        if self.scan_tree.exists(row_id):
            self.scan_tree.focus(row_id)
            self._update_scan_detail(iid=row_id)
        return "break"

    def _row_is_fully_checked(self, iid):
        row = self.scan_item_by_iid.get(iid)
        if row is None:
            return False
        keys = self._row_keys(*row) & self._visible_scan_keys()
        return bool(keys) and keys <= self.selected_scan_keys

    def _set_check_range(self, anchor_iid, target_iid):
        order = list(self._iter_visible_checkable_iids())
        if anchor_iid not in order or target_iid not in order:
            self._toggle_scan_iids([target_iid])
            return
        start, end = sorted((order.index(anchor_iid), order.index(target_iid)))
        checked = self._row_is_fully_checked(anchor_iid)
        visible = self._visible_scan_keys()
        keys = set()
        for iid in order[start:end + 1]:
            row = self.scan_item_by_iid.get(iid)
            if row is None:
                continue
            keys.update(self._row_keys(*row) & visible)
        if checked:
            self.selected_scan_keys.update(keys)
        else:
            self.selected_scan_keys.difference_update(keys)
        self._apply_scan_filters(select_first=True)

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
        keys = set()
        visible_keys = self._visible_scan_keys()
        for iid in iids:
            row = self.scan_item_by_iid.get(iid)
            if row is None:
                continue
            kind, item = row
            # 章节只操作筛选后列出的条款，避免误选隐藏的搜索结果。
            keys.update(self._row_keys(kind, item) & visible_keys)
        if keys and keys.issubset(self.selected_scan_keys):
            self.selected_scan_keys.difference_update(keys)
        else:
            self.selected_scan_keys.update(keys)
        self._apply_scan_filters(select_first=True)

    def select_recommended_tables(self):
        recommended_tables = {
            self._scan_key(kind, item)
            for kind, item in self.scan_item_by_iid.values()
            if kind == "table" and item.hint.startswith("建议关注")
        }
        self.selected_scan_keys.update(recommended_tables)
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set(f"已加上当前结果里的 {len(recommended_tables)} 张建议关注的表，原有勾选保留")

    def _visible_scan_keys(self):
        return {
            self._scan_key(kind, item)
            for kind, item in self.scan_item_by_iid.values()
            if kind != "symbol" or not item.clauses
        }

    def _refresh_visible_selection_button(self):
        if hasattr(self, "export_button"):
            self.export_button.config(text=f"导出已选 {len(self.selected_scan_keys)} 项")

    def clear_all_scan_selection(self):
        self.selected_scan_keys.clear()
        self._refresh_all_scan_rows()

    def select_all_scan_items(self):
        visible_keys = self._visible_scan_keys()
        self.selected_scan_keys.update(visible_keys)
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set(f"已勾选当前筛出的 {len(visible_keys)} 项")

    def clear_scan_selection(self):
        visible_keys = self._visible_scan_keys()
        self.selected_scan_keys.difference_update(visible_keys)
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set(f"已取消当前筛出的 {len(visible_keys)} 项")

    def _refresh_all_scan_rows(self):
        self._apply_scan_filters(select_first=True)

    def _update_scan_detail(self, event=None, iid=None):
        self._refresh_highlight_label()
        if iid is None:
            selection = self.scan_tree.selection()
            iid = selection[0] if selection else self.scan_tree.focus()
        if iid in self._file_group_labels:
            self._set_detail_text(f"文件：{self._file_group_labels[iid]}。勾选下面的表格或章节。")
            return
        row = self.scan_item_by_iid.get(iid)
        if row is None:
            self._set_detail_text("选择一行后，这里按表格显示内容。")
            return

        kind, item = row
        if kind == "table":
            caption = f"{item.section}    表格{item.table_index}    {item.row_count}×{item.column_count}"
            if item.hint:
                caption = f"{caption}\n{item.hint}"
            self._show_preview_table(caption, self._word_table_preview_cells(item))
            return
        clauses = list(getattr(item, "clauses", ()) or ())
        caption = f"{item.section}    {len(clauses) or item.clause_count}条"
        symbols = " ".join(getattr(item, "symbols", ()) or ())
        if symbols:
            caption = f"{caption}    {symbols}"
        self._show_preview_table(caption, self._symbol_preview_cells(clauses, item))

    def _word_table_preview_cells(self, item):
        table = getattr(item, "table", None)
        cells = getattr(table, "cells", None) if table is not None else None
        if not cells:
            return [(0, 0, 1, 1, item.preview or "这一行没有可显示的表格内容。", False)]
        return [
            (cell.row - 1, cell.column - 1, max(1, cell.row_span), max(1, cell.column_span), cell.text, cell.row == 1)
            for cell in cells
        ]

    def _symbol_preview_cells(self, clauses, item):
        if clauses and any(getattr(clause, "columns", ()) for clause in clauses):
            sample = next(clause for clause in clauses if clause.columns)
            headers = sample.headers or tuple(f"列{index + 1}" for index in range(len(sample.columns)))
            cells = [(0, index, 1, 1, title, True) for index, title in enumerate(headers)]
            for row_index, clause in enumerate(clauses, start=1):
                for column_index, value in enumerate(clause.columns):
                    cells.append((row_index, column_index, 1, 1, value, False))
            return cells
        header = ("数量序号", "符号", "详细内容（带序号）")
        cells = [(0, index, 1, 1, title, True) for index, title in enumerate(header)]
        if not clauses:
            symbols = " ".join(getattr(item, "symbols", ()) or ())
            cells.append((1, 0, 1, 1, "1", False))
            cells.append((1, 1, 1, 1, symbols, False))
            cells.append((1, 2, 1, 1, getattr(item, "preview", "") or "", False))
            return cells
        for index, clause in enumerate(clauses, start=1):
            cells.append((index, 0, 1, 1, str(index), False))
            cells.append((index, 1, 1, 1, clause.symbol, False))
            cells.append((index, 2, 1, 1, clause.text, False))
        return cells

    def _show_preview_table(self, caption, cells):
        self.detail_caption.configure(text=caption)
        self._clear_preview_cells()
        wraps = {}
        for row, column, _row_span, col_span, text, _header in cells:
            wraps[(row, column)] = self._preview_wrap(text, col_span)
        for row, column, row_span, col_span, text, header in cells:
            bg = "#E8EEF4" if header else "#FFFFFF"
            label = tk.Label(
                self.detail_table,
                text=text,
                bg=bg,
                fg=self.app.text_fg,
                font=self.app.body_font,
                justify="left",
                anchor="nw",
                wraplength=wraps[(row, column)],
                padx=6,
                pady=4,
            )
            label.grid(
                row=row,
                column=column,
                rowspan=row_span,
                columnspan=col_span,
                sticky="nsew",
                padx=1,
                pady=1,
            )
            self._bind_preview_wheel(label)
        self._refresh_preview_scroll()

    def _preview_wrap(self, text, column_span):
        length = len(text or "")
        if length > 48:
            width = 240
        elif length > 18:
            width = 150
        else:
            width = 88
        return width * max(1, column_span)

    def _clear_preview_cells(self):
        for child in self.detail_table.winfo_children():
            child.destroy()

    def _set_detail_text(self, text):
        if not hasattr(self, "detail_caption"):
            return
        self.detail_caption.configure(text=text)
        self._clear_preview_cells()
        self._refresh_preview_scroll()

    def preview_plain_text(self):
        parts = [self.detail_caption.cget("text")]
        for child in self.detail_table.winfo_children():
            try:
                parts.append(child.cget("text"))
            except tk.TclError:
                continue
        return "\n".join(parts)

    def _bind_preview_wheel(self, widget):
        widget.bind("<MouseWheel>", self._on_preview_wheel)

    def _on_preview_wheel(self, event):
        if event.state & 0x0001:
            self.detail_canvas.xview_scroll(int(-event.delta / 120), "units")
        else:
            self.detail_canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def _on_preview_xscroll(self, first, last):
        self.detail_x_scrollbar.set(first, last)
        self._sync_preview_x_scrollbar(first, last)

    def _sync_preview_x_scrollbar(self, first, last):
        if float(last) - float(first) >= 0.999:
            self.detail_x_scrollbar.grid_remove()
            return
        if not self.detail_x_scrollbar.grid_info():
            self.detail_x_scrollbar.grid(row=1, column=0, sticky="ew")

    def _refresh_preview_scroll(self, event=None):
        if not hasattr(self, "detail_canvas"):
            return
        try:
            bbox = self.detail_canvas.bbox("all")
        except tk.TclError:
            return
        if not bbox:
            self.detail_canvas.configure(scrollregion=(0, 0, 0, 0))
            self.detail_x_scrollbar.grid_remove()
            return
        self.detail_canvas.configure(scrollregion=bbox)
        width = self.detail_frame.winfo_width()
        if width > 1:
            self.detail_caption.configure(wraplength=max(width - 8, 160))
        self._sync_preview_x_scrollbar(*self.detail_canvas.xview())

    def _on_window_configure(self, event=None):
        if event is not None and event.widget is not self.window:
            return
        self._refresh_constrained_texts()

    def _on_output_frame_configure(self, event=None):
        if event is not None and event.widget is not self.output_frame:
            return
        self._refresh_constrained_texts()

    def _output_path_width(self):
        try:
            frame_w = self.output_frame.winfo_width()
        except Exception:
            return 0
        if frame_w <= 1:
            return 0
        return max(frame_w - 80 - 110 - 40, 120)

    def _status_width(self):
        try:
            row_width = self.action_frame.winfo_width()
            button_width = self.export_button.winfo_reqwidth()
            keep_width = self.keep_symbols_check.canvas.winfo_reqwidth()
        except Exception:
            return 0
        extra = 36 + keep_width
        try:
            if self.progress.winfo_ismapped():
                extra += int(self.progress.winfo_reqwidth()) + 26
        except Exception:
            pass
        return max(row_width - button_width - extra, 120)

    def _refresh_constrained_texts(self):
        if hasattr(self, "_output_path"):
            self._output_path.refresh()
        if hasattr(self, "_status_elide"):
            self._status_elide.refresh()
        try:
            wrap = max(self.window.winfo_width() - 48, 200)
            self.header_hint.configure(wraplength=wrap)
            self.scan_label.configure(wraplength=wrap)
        except (AttributeError, tk.TclError):
            pass

    def select_output_dir(self):
        directory = filedialog.askdirectory(title="选择输出目录", parent=self.window)
        if not directory:
            return
        self.output_dir = directory
        self._output_path.set_text(directory, foreground="green")
        self.status_var.set("已选择输出目录")
        self._refresh_constrained_texts()

    def _on_keep_symbols_toggled(self):
        self._persist_symbol_extract_settings()
        if self.keep_symbols_var.get():
            self.status_var.set("条款正文将保留标记符号")
        else:
            self.status_var.set("条款正文不再保留标记符号，只写入符号列")
        self._refresh_constrained_texts()

    def _persist_symbol_extract_settings(self):
        self.app.keep_clause_symbols = bool(self.keep_symbols_var.get())
        if not self.app.restore_session:
            return
        settings = load_settings()
        settings["keep_clause_symbols"] = self.app.keep_clause_symbols
        save_settings(settings)

    def _refresh_symbols_label(self):
        chars = self.app.symbol_chars or DEFAULT_SYMBOL_CHARS
        self.symbols_label.config(text=f"符号 {' '.join(chars)}")

    def open_symbol_settings(self):
        """打开自定义提取符号对话框：勾选/取消常用符号，也可输入其他符号。"""
        from symbol_clause_extractor import COMMON_SYMBOL_CHOICES, normalize_symbol_chars

        current = normalize_symbol_chars(self.app.symbol_chars) or DEFAULT_SYMBOL_CHARS
        choices = list(COMMON_SYMBOL_CHOICES) + [
            char for char in current if char not in COMMON_SYMBOL_CHOICES
        ]

        dialog = tk.Toplevel(self.window)
        dialog.title("自定义提取符号")
        dialog.withdraw()
        dialog.transient(self.window)
        dialog.resizable(False, False)
        dialog.configure(bg=self.app.app_bg)
        try:
            dialog.iconphoto(True, self.app._icon_photo)
        except Exception:
            pass

        outer = ttk.Frame(dialog, padding=14, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        body = ttk.Frame(outer, padding=14, style="Surface.TFrame")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        ttk.Label(body, text="选择要提取的标记符号", style="SectionTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            body,
            text="勾选要参与提取的符号（默认 ★ # △ ▲，可取消）；也可以在下方输入其他符号。",
            style="Hint.TLabel",
            wraplength=440,
            justify="left",
        ).grid(row=1, column=0, sticky="ew", pady=(3, 9))

        grid = ttk.Frame(body, style="Toolbar.TFrame")
        grid.grid(row=2, column=0, sticky="ew")
        symbol_vars = {}
        chips = {}
        for index, symbol in enumerate(choices):
            var = tk.BooleanVar(value=symbol in current)
            symbol_vars[symbol] = var
            chip = CanvasCheckbox(
                grid,
                symbol,
                var,
                self.app,
                command=lambda: refresh_selected(),
                padx=7,
                pady=5,
            )
            chip.canvas.grid(row=index // 8, column=index % 8, padx=3, pady=2, sticky="w")
            chips[symbol] = chip
        dialog._symbol_chips = chips

        ttk.Label(body, text="其他符号（可直接输入多个，追加到勾选项）", style="SectionTitle.TLabel").grid(
            row=3, column=0, sticky="w", pady=(12, 5)
        )
        entry_var = tk.StringVar()
        entry = ttk.Entry(body, textvariable=entry_var, width=32)
        entry.grid(row=4, column=0, sticky="w")
        ttk.Label(
            body,
            text="数字、字母、汉字和空格会被自动忽略；符号须出现在条款开头或序号之后才会命中。",
            style="Hint.TLabel",
            wraplength=440,
            justify="left",
        ).grid(row=5, column=0, sticky="ew", pady=(4, 0))

        selected_label = ttk.Label(body, text="", style="Muted.TLabel")
        selected_label.grid(row=6, column=0, sticky="w", pady=(8, 0))

        def selected_chars():
            chosen = [symbol for symbol, var in symbol_vars.items() if var.get()]
            chosen += list(normalize_symbol_chars(entry_var.get()))
            return "".join(dict.fromkeys(chosen))

        def refresh_selected():
            chars = selected_chars()
            selected_label.config(
                text=f"已选 {len(chars)} 个：{' '.join(chars)}" if chars else "已选 0 个：请至少保留一个符号"
            )

        def reset_default():
            for symbol, var in symbol_vars.items():
                var.set(symbol in DEFAULT_SYMBOL_CHARS)
            entry_var.set("")
            refresh_selected()

        def save_selection():
            chars = selected_chars()
            if not chars:
                messagebox.showwarning("提示", "请至少选择或输入一个符号。", parent=dialog)
                return
            changed = chars != self.app.symbol_chars
            self.app.symbol_chars = chars
            settings = load_settings()
            settings["symbol_chars"] = chars
            save_settings(settings)
            self._refresh_symbols_label()
            dialog.destroy()
            if changed:
                self._clear_scan_results("符号设置已变化，请重新扫描表格和符号")
                self.status_var.set(f"扫描符号已更新为 {''.join(chars)}，请重新扫描")
            else:
                self.status_var.set(f"扫描符号保持为 {''.join(chars)}")

        buttons = ttk.Frame(body, style="Toolbar.TFrame")
        buttons.grid(row=7, column=0, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="恢复默认", command=reset_default, width=10).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="取消", command=dialog.destroy, width=8).pack(side="left", padx=(0, 8))
        ttk.Button(
            buttons, text="确定", style="Accent.TButton", command=save_selection, width=10
        ).pack(side="left")

        entry.bind("<KeyRelease>", lambda _event: refresh_selected())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        refresh_selected()

        # 高度必须按内容自适应：符号块较多时写死高度会把底部按钮裁掉
        dialog.update_idletasks()
        center_window_on_parent(
            dialog,
            self.window,
            width=max(520, dialog.winfo_reqwidth()),
            height=dialog.winfo_reqheight(),
        )
        dialog.deiconify()
        dialog.lift()
        dialog.focus_force()
        entry.focus_set()

    def start_export(self):
        if not self.file_paths:
            messagebox.showwarning("提示", "请先选择 Word 文件。", parent=self.window)
            return
        if not self.scan_rows:
            messagebox.showwarning(
                "提示", "请先点击“扫描表格和符号”。", parent=self.window
            )
            return
        if not self.selected_scan_keys:
            messagebox.showwarning(
                "提示", "请先在扫描结果中选择至少一项。", parent=self.window
            )
            return

        tables = [item for item in self.table_items if self._scan_key("table", item) in self.selected_scan_keys]
        sections = []
        for item in self.symbol_items:
            clauses = tuple(child.clauses[0] for child in self._clause_items(item)
                            if self._scan_key("clause", child) in self.selected_scan_keys)
            if clauses:
                sections.append(replace(item, clauses=clauses, clause_count=len(clauses)))
        table_file_paths = list(dict.fromkeys(item.file_path for item in tables))
        symbol_file_paths = list(dict.fromkeys(item.file_path for item in sections))
        if not tables and not sections:
            messagebox.showwarning("提示", "请先在扫描结果中选择至少一项。", parent=self.window)
            return

        output_dir = self.output_dir
        symbol_chars = self.app.symbol_chars
        keep_symbols_in_text = bool(self.keep_symbols_var.get())
        total_files = len(table_file_paths) + len(symbol_file_paths)
        self._begin_determinate_task(f"正在导出 {total_files} 个文件...", total_files)

        stamps = dict(self.scan_stamps)

        def work(publish_progress):
            from word_scan import export_scanned_content
            return export_scanned_content(tables, sections, stamps, output_dir, symbol_chars,
                                          keep_symbols_in_text, publish_progress)

        def on_success(result):
            table_results, table_skipped, table_error, symbol_results, symbol_skipped, symbol_error = result
            self._show_export_result(
                table_results,
                table_skipped,
                table_error,
                list(dict.fromkeys(table_file_paths + symbol_file_paths)),
                symbol_results=symbol_results,
                symbol_skipped=symbol_skipped,
                symbol_error=symbol_error,
            )

        def on_error(exc):
            self._finish_background_error(exc, "Word 表格和符号导出线程")

        self._task_runner.submit(work, on_success, on_error, self._update_progress)

    def _update_progress(self, current, total, filename, action="正在导出"):
        self.progress.configure(maximum=total, value=current)
        self.status_var.set(f"{action} ({current}/{total})：{filename}")
        self._refresh_constrained_texts()

    def _reset_busy_state(self):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.configure(value=0)
        self.progress.grid_remove()
        self._set_busy_controls(False)
        self._sync_scan_choice_state()
        self._refresh_constrained_texts()

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

    def _show_export_result(
        self,
        results,
        skipped,
        error,
        source_files,
        symbol_results=None,
        symbol_skipped=None,
        symbol_error=None,
    ):
        self._reset_busy_state()
        total_tables = sum(int(item["tables"]) for item in results.values())
        symbol_results = symbol_results or {}
        symbol_skipped = symbol_skipped or {}
        total_clauses = sum(int(item["count"]) for item in symbol_results.values())

        lines = ["导出完成。"]
        table_activity = bool(results or skipped or error)
        symbol_activity = bool(symbol_results or symbol_skipped or symbol_error)

        if table_activity:
            lines.extend(["", f"表格导出：成功 {len(results)} 个文件，共 {total_tables} 张表格。"])
            for filename, info in results.items():
                lines.append(f"  {filename}: {info['tables']} 张表格 -> {info['output_path']}")

        if symbol_activity:
            lines.extend([
                "",
                f"指标参数提取：成功 {len(symbol_results)} 个文件，共 {total_clauses} 条带符号条款。",
            ])
            for filename, info in symbol_results.items():
                lines.append(f"  {filename}: {info['count']} 条 -> {info['output_path']}")
            if not symbol_results:
                lines.append("  （未提取到带符号条款）")

        if skipped:
            lines.extend(["", "表格导出跳过文件："])
            for filename, reason in skipped.items():
                lines.append(f"  {filename}: {reason}")

        if symbol_skipped:
            lines.extend(["", "指标参数提取跳过文件："])
            for filename, reason in symbol_skipped.items():
                lines.append(f"  {filename}: {reason}")

        if error or symbol_error:
            lines.extend(["", "失败文件："])
            if error:
                lines.append(error)
            if symbol_error:
                lines.append(symbol_error)
            self.status_var.set("导出完成（部分失败）")
        else:
            self.status_var.set("导出完成")

        self._last_output_dir_to_open = self._default_output_dir_to_open(
            results, source_files, symbol_results=symbol_results,
        )
        self._show_result_window("\n".join(lines))

    def _default_output_dir_to_open(self, results=None, source_files=None, symbol_results=None):
        if self.output_dir:
            return self.output_dir
        outputs = [results, symbol_results]
        for collected in outputs:
            first = next(iter(collected.values()), None) if collected else None
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
        window.minsize(560, 320)
        window.configure(bg=self.app.app_bg)
        window.transient(self.window)
        window.withdraw()

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
        scrollbar = FlatScrollbar(
            body, orient="vertical", command=text.yview, style="Flat.Vertical.TScrollbar"
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)
        self.app._bind_vertical_mousewheel(text, text, scrollbar)
        text.insert("1.0", message)
        text.configure(state="disabled")

        buttons = ttk.Frame(body, style="App.TFrame")
        buttons.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="打开输出目录", command=self._open_output_dir, width=14).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="关闭", command=window.destroy, width=10).pack(side="left")

        geometry = center_window_on_parent(window, self.window, width=720, height=440)
        window.deiconify()
        if geometry:
            window.geometry(geometry)
        else:
            center_window_on_parent(window, self.window, width=720, height=440)
        window.lift()
        window.focus_force()


