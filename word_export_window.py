"""Word 表格与指标条款提取窗口。"""

import os
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
    FlatScrollbar,
)


class WordTableExportWindow:
    WINDOW_WIDTH = 1280
    WINDOW_HEIGHT = 1100
    MIN_WIDTH = 1020
    MIN_HEIGHT = 820
    SCAN_COLUMN_TITLES = {
        "selected": "导出",
        "file": "文件",
        "type": "类型",
        "section": "所在章节",
        "item": "项目",
        "quantity": "数量",
    }

    def __init__(self, app, initial_files=None):
        self.app = app
        self.window = tk.Toplevel(app.root)
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
        self._filter_popup = None
        self.output_dir = None
        self.keep_symbols_var = tk.BooleanVar(value=bool(app.keep_clause_symbols))
        self._last_output_dir_to_open = None
        self._scan_separator_after = None
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
        self._close_scan_filter_popup()
        if self._scan_separator_after is not None:
            try:
                self.window.after_cancel(self._scan_separator_after)
            except tk.TclError:
                pass
            self._scan_separator_after = None
        if hasattr(self, "_output_path"):
            self._output_path.hide_tooltip()
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
            text="选择 .doc / .docx 文件；同时扫描 Word 表格和当前符号集，按文件实际章节勾选后导出。",
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
        # 扫描表吃掉中间全部弹性空间；明细栏保持可读高度，窗口再拉高时表继续长
        scan_frame.rowconfigure(1, weight=1)
        scan_frame.rowconfigure(2, weight=0)

        scan_header = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        scan_header.grid(row=0, column=0, sticky="ew")
        scan_header.columnconfigure(0, weight=1)
        ttk.Label(scan_header, text="扫描结果", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")

        scan_toolbar = ttk.Frame(scan_header, style="Toolbar.TFrame")
        scan_toolbar.grid(row=0, column=1, sticky="e")
        self.scan_button = self._create_busy_button(
            scan_toolbar, text="扫描表格和符号", command=self.scan_content, width=14
        )
        self.scan_button.pack(side="left", padx=(0, 6))
        self._create_busy_button(scan_toolbar, text="推荐表格", command=self.select_recommended_tables, width=10).pack(side="left", padx=(0, 6))
        self.toggle_visible_selection_button = self._create_busy_button(
            scan_toolbar,
            text="筛选结果全选",
            command=self.toggle_visible_scan_selection,
            width=14,
        )
        self.toggle_visible_selection_button.pack(side="left")

        tree_frame = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        tree_frame.grid(row=1, column=0, sticky="nsew", pady=(7, 0))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.scan_tree = ttk.Treeview(
            tree_frame,
            columns=("selected", "file", "type", "section", "item", "quantity"),
            show="headings",
            selectmode="extended",
            style="Scan.Treeview",
            height=18,
        )
        for column, title in self.SCAN_COLUMN_TITLES.items():
            self.scan_tree.heading(column, text=f"{title} ▼")
        self.scan_tree.column("selected", width=64, minwidth=56, anchor="center", stretch=False)
        self.scan_tree.column("file", width=205, minwidth=130, stretch=False)
        self.scan_tree.column("type", width=90, minwidth=76, anchor="center", stretch=False)
        self.scan_tree.column("section", width=430, minwidth=220)
        self.scan_tree.column("item", width=150, minwidth=100, anchor="center", stretch=False)
        self.scan_tree.column("quantity", width=76, minwidth=64, anchor="center", stretch=False)
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
        self.scan_tree.bind("<ButtonRelease-1>", self._open_scan_filter_from_heading, add="+")
        self.scan_tree.bind("<ButtonRelease-1>", self._schedule_scan_tree_column_separators, add="+")
        self.scan_tree.bind("<B1-Motion>", self._schedule_scan_tree_column_separators, add="+")
        self.scan_tree.bind("<Configure>", self._schedule_scan_tree_column_separators, add="+")
        self.scan_tree.bind("<Double-1>", self._toggle_scan_row_from_event)
        self.scan_tree.bind("<space>", self._toggle_scan_rows_from_keyboard)
        self.scan_tree.bind("<<TreeviewSelect>>", self._update_scan_detail)
        self._schedule_scan_tree_column_separators()

        detail_frame = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        detail_frame.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(0, weight=1)
        self.detail_text = tk.Text(
            detail_frame,
            height=10,
            wrap="word",
            font=self.app.body_font,
            bg="#F1F3F6",
            fg=self.app.text_fg,
            relief="solid",
            borderwidth=1,
            padx=10,
            pady=6,
        )
        self.detail_text.grid(row=0, column=0, sticky="nsew")
        self.detail_scrollbar = FlatScrollbar(
            detail_frame,
            orient="vertical",
            command=self.detail_text.yview,
            style="Flat.Vertical.TScrollbar",
        )
        self.detail_scrollbar.grid(row=0, column=1, sticky="ns")
        self.detail_text.configure(yscrollcommand=self.detail_scrollbar.set)
        self.app._bind_vertical_mousewheel(
            self.detail_text, self.detail_text, self.detail_scrollbar
        )
        self.detail_text.insert("1.0", "选择扫描结果中的一行，可查看完整章节、表格或符号条款预览。")
        self.detail_text.configure(state="disabled")

        self.scan_label = ttk.Label(scan_frame, text="请先添加 Word 文件并扫描表格和符号", style="Muted.TLabel")
        self.scan_label.grid(row=3, column=0, sticky="ew", pady=(5, 0))

        # 符号扫描设置：符号范围在扫描前确定，是否保留符号只影响最终导出
        symbols_frame = ttk.Frame(body, padding=(12, 8), style="Surface.TFrame")
        symbols_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        symbols_frame.columnconfigure(0, weight=1)

        symbols_header = ttk.Frame(symbols_frame, style="Toolbar.TFrame")
        symbols_header.grid(row=0, column=0, sticky="ew")
        symbols_header.columnconfigure(0, weight=1)
        ttk.Label(symbols_header, text="符号扫描设置", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.customize_symbols_button = self._create_busy_button(
            symbols_header,
            text="自定义符号…",
            command=self.open_symbol_settings,
            width=12,
        )
        self.customize_symbols_button.grid(row=0, column=1, sticky="e")

        symbols_content = ttk.Frame(symbols_frame, style="Toolbar.TFrame")
        symbols_content.grid(row=1, column=0, sticky="ew", pady=(7, 0))
        symbols_content.columnconfigure(0, weight=1)

        symbols_row = ttk.Frame(symbols_content, style="Toolbar.TFrame")
        symbols_row.grid(row=0, column=0, sticky="ew")
        ttk.Label(symbols_row, text="扫描符号").pack(side="left")
        self.symbols_label = ttk.Label(symbols_row, text="当前：★ # △ ▲", style="Muted.TLabel")
        self.symbols_label.pack(side="left", padx=(10, 0))

        keep_row = ttk.Frame(symbols_content, style="Toolbar.TFrame")
        keep_row.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.keep_symbols_check = CanvasCheckbox(
            keep_row,
            "条款内容中保留符号",
            self.keep_symbols_var,
            self.app,
            command=self._on_keep_symbols_toggled,
        )
        self.keep_symbols_check.canvas.pack(side="left")
        ttk.Label(
            keep_row,
            text="不勾选则只在符号列保留，条款正文不再重复带符号",
            style="Muted.TLabel",
        ).pack(side="left", padx=(10, 0))

        self.output_frame = ttk.Frame(body, padding=(12, 6), style="Surface.TFrame")
        output_frame = self.output_frame
        output_frame.grid(row=4, column=0, sticky="ew", pady=(0, 8))
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
        action_frame.grid(row=5, column=0, sticky="ew")
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(2, minsize=118)

        self.status_label = ttk.Label(action_frame, text="就绪", style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="ew")
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

    def _close_scan_filter_popup(self):
        popup = self._filter_popup
        self._filter_popup = None
        if popup is None:
            return
        try:
            popup.grab_release()
        except tk.TclError:
            pass
        try:
            if popup.winfo_exists():
                popup.destroy()
        except tk.TclError:
            pass

    def _refresh_scan_filter_headings(self):
        if not hasattr(self, "scan_tree"):
            return
        for column, title in self.SCAN_COLUMN_TITLES.items():
            marker = " [筛] ▼" if column in self.scan_filters else " ▼"
            self.scan_tree.heading(column, text=title + marker)

    def _scan_filter_value(self, column, kind, item):
        if column == "selected":
            return "已选" if self._scan_key(kind, item) in self.selected_scan_keys else "未选"
        if column == "file":
            return item.filename
        if column == "type":
            return "表格" if kind == "table" else "符号条款"
        if column == "section":
            return item.section
        if column == "item":
            return f"表格{item.table_index}" if kind == "table" else " ".join(item.symbols)
        if column == "quantity":
            return (
                f"{item.row_count}x{item.column_count}"
                if kind == "table"
                else f"{item.clause_count}条"
            )
        return ""

    def _scan_row_matches_filters(self, kind, item, ignore_column=None):
        for column, selected_values in self.scan_filters.items():
            if column == ignore_column:
                continue
            if self._scan_filter_value(column, kind, item) not in selected_values:
                return False
        return True

    def _available_scan_filter_values(self, column):
        values = []
        seen = set()
        for kind, item in self.scan_rows:
            if not self._scan_row_matches_filters(kind, item, ignore_column=column):
                continue
            value = self._scan_filter_value(column, kind, item)
            if value in seen:
                continue
            seen.add(value)
            values.append(value)
        return values

    def _set_scan_filter(self, column, selected_values):
        available = set(self._available_scan_filter_values(column))
        chosen = set(selected_values)
        if chosen == available:
            self.scan_filters.pop(column, None)
        else:
            self.scan_filters[column] = chosen
        self._apply_scan_filters(select_first=True)

    def _open_scan_filter_from_heading(self, event):
        if self.scan_tree.identify_region(event.x, event.y) != "heading":
            return
        column_id = self.scan_tree.identify_column(event.x)
        try:
            column_index = int(column_id.removeprefix("#")) - 1
            column = self.scan_tree.cget("columns")[column_index]
        except (ValueError, IndexError, TypeError):
            return
        if column not in self.SCAN_COLUMN_TITLES:
            return

        self._close_scan_filter_popup()
        values = self._available_scan_filter_values(column)
        selected_values = set(self.scan_filters.get(column, values))

        popup = tk.Toplevel(self.window)
        self._filter_popup = popup
        popup.title(f"筛选：{self.SCAN_COLUMN_TITLES[column]}")
        popup.transient(self.window)
        popup.resizable(False, False)
        popup.configure(bg=self.app.app_bg)
        popup.protocol("WM_DELETE_WINDOW", self._close_scan_filter_popup)

        outer = ttk.Frame(popup, padding=10, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        body = ttk.Frame(outer, padding=10, style="Surface.TFrame")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)
        ttk.Label(
            body,
            text=f"筛选“{self.SCAN_COLUMN_TITLES[column]}”（点击可多选）",
            style="SectionTitle.TLabel",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 7))

        list_width = min(72, max(26, max((len(value) for value in values), default=12) + 3))
        listbox = tk.Listbox(
            body,
            selectmode="multiple",
            exportselection=False,
            height=min(14, max(4, len(values))),
            width=list_width,
            font=self.app.body_font,
            bg=self.app.surface_bg,
            fg=self.app.text_fg,
            selectbackground="#3F6D92",
            selectforeground="white",
            relief="solid",
            borderwidth=1,
        )
        listbox.grid(row=1, column=0, sticky="nsew")
        for index, value in enumerate(values):
            listbox.insert("end", value)
            if value in selected_values:
                listbox.selection_set(index)

        scrollbar = FlatScrollbar(
            body,
            orient="vertical",
            command=listbox.yview,
            style="Flat.Vertical.TScrollbar",
        )
        scrollbar.grid(row=1, column=1, sticky="ns")
        x_scrollbar = FlatScrollbar(
            body,
            orient="horizontal",
            command=listbox.xview,
            style="Flat.Horizontal.TScrollbar",
        )
        x_scrollbar.grid(row=2, column=0, sticky="ew")
        listbox.configure(
            yscrollcommand=scrollbar.set,
            xscrollcommand=x_scrollbar.set,
        )
        self.app._bind_vertical_mousewheel(listbox, listbox, scrollbar)

        selection_buttons = ttk.Frame(body, style="Toolbar.TFrame")
        selection_buttons.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(
            selection_buttons,
            text="全选",
            width=8,
            command=lambda: listbox.selection_set(0, "end"),
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            selection_buttons,
            text="全不选",
            width=8,
            command=lambda: listbox.selection_clear(0, "end"),
        ).pack(side="left")

        def apply_selection():
            chosen = [values[index] for index in listbox.curselection()]
            self._close_scan_filter_popup()
            self._set_scan_filter(column, chosen)

        def clear_filter():
            self.scan_filters.pop(column, None)
            self._close_scan_filter_popup()
            self._apply_scan_filters(select_first=True)

        buttons = ttk.Frame(body, style="Toolbar.TFrame")
        buttons.grid(row=4, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="清除筛选", command=clear_filter, width=10).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="取消", command=self._close_scan_filter_popup, width=8).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="确定", style="Accent.TButton", command=apply_selection, width=8).pack(side="left")

        popup.bind("<Escape>", lambda _event: self._close_scan_filter_popup())
        popup.update_idletasks()
        popup_width = popup.winfo_reqwidth()
        popup_height = popup.winfo_reqheight()
        screen_width = popup.winfo_screenwidth()
        screen_height = popup.winfo_screenheight()
        x = max(8, min(event.x_root, screen_width - popup_width - 8))
        y = max(8, min(event.y_root + 4, screen_height - popup_height - 48))
        popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
        popup.grab_set()
        popup.lift()
        listbox.focus_set()

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
            total_width = sum(widths)
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
            boundary = 0
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
        self._close_scan_filter_popup()
        self.selected_scan_keys = set()
        self.scan_filters = {}

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
        self._apply_scan_filters(select_first=True)

        self._reset_busy_state()
        message = (
            f"已扫描到 {len(self.table_items)} 张表格、{len(self.symbol_items)} 个符号章节，"
            "当前选择 0 项"
        )
        if table_error or symbol_error:
            message += "；部分文件失败"
        has_items = bool(ordered_rows)
        self.scan_label.config(text=message, foreground="green" if has_items else self.app.muted_fg)
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
        self.table_items = []
        self.symbol_items = []
        self.scan_rows = []
        self.scan_stamps = {}
        self.scan_item_by_iid = {}
        self.selected_scan_keys = set()
        self.scan_filters = {}
        self._close_scan_filter_popup()
        if hasattr(self, "scan_tree"):
            self.scan_tree.delete(*self.scan_tree.get_children())
            self._refresh_scan_filter_headings()
            self._schedule_scan_tree_column_separators()
        if hasattr(self, "scan_label"):
            self.scan_label.config(text=label_text, foreground=self.app.muted_fg)
        self._refresh_visible_selection_button()
        if hasattr(self, "detail_text"):
            self._set_detail_text("选择扫描结果中的一行，可查看完整章节、表格或符号条款预览。")

    def _scan_tree_values(self, kind, item, selected):
        if kind == "table":
            item_text = f"表格{item.table_index}"
            quantity = f"{item.row_count}x{item.column_count}"
            type_text = "表格"
        else:
            item_text = " ".join(item.symbols)
            quantity = f"{item.clause_count}条"
            type_text = "符号条款"
        return (
            "☑" if selected else "☐",
            item.filename,
            type_text,
            item.section,
            item_text,
            quantity,
        )

    def _scan_key(self, kind, item):
        if kind == "table":
            return (kind, _file_identity(item.file_path), item.table_index)
        return (kind, _file_identity(item.file_path), item.section)

    def _apply_scan_filters(self, select_first=False):
        focus_key = None
        current_iid = self.scan_tree.focus() if hasattr(self, "scan_tree") else ""
        current_row = self.scan_item_by_iid.get(current_iid)
        if current_row is not None:
            focus_key = self._scan_key(*current_row)

        self.scan_tree.delete(*self.scan_tree.get_children())
        self.scan_item_by_iid = {}
        focus_iid = None
        for row_index, (kind, item) in enumerate(self.scan_rows, start=1):
            if not self._scan_row_matches_filters(kind, item):
                continue
            iid = str(row_index)
            key = self._scan_key(kind, item)
            self.scan_item_by_iid[iid] = (kind, item)
            self.scan_tree.insert(
                "",
                "end",
                iid=iid,
                values=self._scan_tree_values(
                    kind, item, selected=key in self.selected_scan_keys
                ),
            )
            if key == focus_key:
                focus_iid = iid

        children = self.scan_tree.get_children()
        target_iid = focus_iid or (children[0] if children and select_first else None)
        if target_iid:
            self.scan_tree.selection_set(target_iid)
            self.scan_tree.focus(target_iid)
        if children:
            self._update_scan_detail()
        elif self.scan_rows:
            self._set_detail_text("当前筛选条件下没有扫描结果，请调整表头筛选。")
        else:
            self._set_detail_text("没有扫描到可导出的正文表格或带符号条款。")

        self._refresh_scan_filter_headings()
        self._refresh_scan_label()
        self._schedule_scan_tree_column_separators()

    def _refresh_scan_label(self):
        selected = len(self.selected_scan_keys)
        visible = len(self.scan_item_by_iid)
        total = len(self.scan_rows)
        filter_text = f"，筛选显示 {visible}/{total} 项" if self.scan_filters else ""
        text = (
            f"已扫描到 {len(self.table_items)} 张表格、{len(self.symbol_items)} 个符号章节"
            f"{filter_text}，当前选择 {selected} 项"
        )
        self.scan_label.config(text=text, foreground="green" if selected else self.app.muted_fg)
        self._refresh_visible_selection_button()

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
            row = self.scan_item_by_iid.get(iid)
            if row is None:
                continue
            kind, item = row
            key = self._scan_key(kind, item)
            if key in self.selected_scan_keys:
                self.selected_scan_keys.remove(key)
            else:
                self.selected_scan_keys.add(key)
        self._apply_scan_filters(select_first=True)

    def select_recommended_tables(self):
        visible_table_keys = {
            self._scan_key(kind, item)
            for kind, item in self.scan_item_by_iid.values()
            if kind == "table"
        }
        recommended_tables = {
            self._scan_key(kind, item)
            for kind, item in self.scan_item_by_iid.values()
            if kind == "table" and item.hint.startswith("建议关注")
        }
        self.selected_scan_keys.difference_update(visible_table_keys)
        self.selected_scan_keys.update(recommended_tables)
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set(f"已按提示选择 {len(recommended_tables)} 张建议关注的表格")

    def _visible_scan_keys(self):
        return {
            self._scan_key(kind, item)
            for kind, item in self.scan_item_by_iid.values()
        }

    def _refresh_visible_selection_button(self):
        if not hasattr(self, "toggle_visible_selection_button"):
            return
        visible_keys = self._visible_scan_keys()
        all_selected = bool(visible_keys) and visible_keys.issubset(self.selected_scan_keys)
        self.toggle_visible_selection_button.config(
            text="筛选结果全不选" if all_selected else "筛选结果全选"
        )

    def toggle_visible_scan_selection(self):
        visible_keys = self._visible_scan_keys()
        if not visible_keys:
            self.status_var.set("当前筛选条件下没有可选择内容")
            return
        if visible_keys.issubset(self.selected_scan_keys):
            self.selected_scan_keys.difference_update(visible_keys)
            action = "已取消当前筛选显示的全部项目"
        else:
            self.selected_scan_keys.update(visible_keys)
            action = "已选择当前筛选显示的全部项目"
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set(f"{action}（{len(visible_keys)} 项）")

    def select_all_scan_items(self):
        visible_keys = self._visible_scan_keys()
        self.selected_scan_keys.update(visible_keys)
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set(f"已选择当前显示的 {len(visible_keys)} 项扫描结果")

    def clear_scan_selection(self):
        visible_keys = self._visible_scan_keys()
        self.selected_scan_keys.difference_update(visible_keys)
        self._refresh_all_scan_rows()
        self._update_scan_detail()
        self.status_var.set(f"已取消当前显示的 {len(visible_keys)} 项选择")

    def _refresh_all_scan_rows(self):
        self._apply_scan_filters(select_first=True)

    def _update_scan_detail(self, event=None):
        selection = self.scan_tree.selection()
        iid = selection[0] if selection else self.scan_tree.focus()
        row = self.scan_item_by_iid.get(iid)
        if row is None:
            self._set_detail_text("选择扫描结果中的一行，可查看完整章节、表格或符号条款预览。")
            return

        kind, item = row
        selected = "是" if self._scan_key(kind, item) in self.selected_scan_keys else "否"
        if kind == "table":
            detail = (
                f"导出：{selected}    类型：表格    文件：{item.filename}    "
                f"表格：{item.table_index}    行列：{item.row_count}x{item.column_count}\n"
                f"所在章节：{item.section}\n"
                f"提示：{item.hint}\n"
                f"表格前文：{item.context}\n"
                f"内容预览：{item.preview}"
            )
        else:
            detail = (
                f"导出：{selected}    类型：符号条款    文件：{item.filename}    "
                f"条款：{item.clause_count}条\n"
                f"所在章节：{item.section}\n"
                f"发现符号：{' '.join(item.symbols)}\n"
                f"内容来源：{'、'.join(item.sources) or '未识别'}\n"
                f"内容预览：{item.preview or '无可读文本'}"
            )
        self._set_detail_text(detail)

    def _set_detail_text(self, text):
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", text)
        self.detail_text.configure(state="disabled")

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
        except Exception:
            return 0
        extra = 20
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
        self.symbols_label.config(text=f"当前：{' '.join(chars)}")

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
        sections = [item for item in self.symbol_items if self._scan_key("symbol", item) in self.selected_scan_keys]
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


