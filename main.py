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
DEFAULT_PRESETS = (
    "[项目名称]",
    "[项目编号]",
    "[标的名称]",
    "[采购人名称]",
    "[采购人联系人]",
    "[采购人电话]",
    "[开标时间]",
    "[开标日期]",
    "[招标公告日期]",
    "[开标地点]",
    "[报名人数]",
    "[采购人地址]",
)
SUPPORTED_EXTENSIONS = (".docx", ".xlsx", ".xlsm", ".pptx")
APP_STATE_DIR_NAME = "replace-simple"
ERROR_LOG_NAME = "error.log"
SETTINGS_NAME = "settings.json"

COMPRESSED_SHEET_HEIGHT = 210      # 基准规则表像素高度（默认窗口下约露 6 行 + 表头，且底部按钮可见）
MIN_SHEET_HEIGHT = 120             # 窗口较矮或系统缩放较大时，优先保住底部操作区
TASK_POLL_INTERVAL_MS = 40
FOLDER_SCAN_PROGRESS_INTERVAL = 100
OUTPUT_DIR_HINT = "未选择则输出到原文件目录；文件名按规则同步替换，同名时直接覆盖原文件"
OUTPUT_DIR_MAX_LINES = 2           # 长路径最多占两行，避免把「开始替换」顶出可视区

# 网格线配色（护眼浅色版）：略降亮度与冷蓝感，长时间观看更柔和。
# tksheet 网格线宽度硬编码 1px，靠颜色保持可辨。
GRID_COLOR = "#A0A7B2"          # 数据区网格线（柔和中灰）
HEADER_GRID_COLOR = "#8B929C"   # 表头/序号列网格线（略深于数据区）
HEADER_BG = "#E8EBEF"           # 表头/序号列底色（低于纯浅灰）
ZEBRA_BG = "#EEF0F3"            # 斑马纹底色

def _app_data_dir(env_name):
    base = os.environ.get(env_name) or os.path.expanduser("~")
    return os.path.join(base, APP_STATE_DIR_NAME)


def error_log_path():
    return os.path.join(_app_data_dir("LOCALAPPDATA"), ERROR_LOG_NAME)


def settings_path():
    return os.path.join(_app_data_dir("APPDATA"), SETTINGS_NAME)


def elide_middle(text, font, max_width, ellipsis="..."):
    """把文本中间省略，使像素宽度不超过 max_width。路径会多留尾部目录名。"""
    if not text or max_width <= 0:
        return text
    if font.measure(text) <= max_width:
        return text

    ellipsis_w = font.measure(ellipsis)
    if ellipsis_w >= max_width:
        for index in range(len(ellipsis), 0, -1):
            piece = ellipsis[:index]
            if font.measure(piece) <= max_width:
                return piece
        return ""

    low, high = 0, len(text)
    best = ellipsis
    while low <= high:
        keep = (low + high) // 2
        if keep <= 0:
            candidate = ellipsis
        else:
            head = max(1, keep * 2 // 5)
            tail = keep - head
            candidate = text[:head] + ellipsis + text[-tail:] if tail > 0 else text[:keep] + ellipsis
        if font.measure(candidate) <= max_width:
            best = candidate
            low = keep + 1
        else:
            high = keep - 1
    return best


class HoverTooltip:
    """鼠标悬停时显示完整文本，仅在控件展示内容被截断时出现。"""

    def __init__(self, widget, text_getter, font, *, wraplength=480, fg="#2B2F36", shown_getter=None):
        self.widget = widget
        self.text_getter = text_getter
        self.shown_getter = shown_getter
        self.font = font
        self.wraplength = wraplength
        self.fg = fg
        self._tip = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<Destroy>", lambda _event: self._hide(), add="+")

    def _shown_text(self):
        if self.shown_getter is not None:
            try:
                return str(self.shown_getter() or "")
            except Exception:
                return ""
        try:
            return str(self.widget.cget("text") or "")
        except Exception:
            return ""

    def _show(self, _event=None):
        self._hide()
        try:
            full = self.text_getter()
        except Exception:
            return
        if not full:
            return
        if self._shown_text() == full:
            return
        tip = tk.Toplevel(self.widget)
        tip.wm_overrideredirect(True)
        try:
            tip.wm_attributes("-topmost", True)
        except Exception:
            pass
        tk.Label(
            tip,
            text=full,
            justify="left",
            background="#FFF8DC",
            foreground=self.fg,
            relief="solid",
            borderwidth=1,
            font=self.font,
            wraplength=self.wraplength,
            padx=8,
            pady=5,
        ).pack()
        try:
            x = self.widget.winfo_rootx()
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
            tip.geometry(f"+{x}+{y}")
        except Exception:
            tip.destroy()
            return
        self._tip = tip

    def _hide(self, _event=None):
        tip = self._tip
        self._tip = None
        if tip is None:
            return
        try:
            tip.destroy()
        except Exception:
            pass

    def hide(self):
        self._hide()


class ElidedTextController:
    """把 Label 上的长文本限制在可用宽度内，避免把同行按钮挤出窗口。"""

    def __init__(
        self,
        label,
        font,
        width_getter,
        *,
        max_lines=1,
        tooltip=False,
        tooltip_font=None,
        tooltip_fg=None,
        fallback_width=360,
    ):
        self.label = label
        self.font = font
        self.width_getter = width_getter
        self.max_lines = max(1, int(max_lines))
        self.fallback_width = fallback_width
        self.full_text = str(label.cget("text") or "")
        self._tooltip = None
        if tooltip:
            self._tooltip = HoverTooltip(
                label,
                lambda: self.full_text,
                tooltip_font or font,
                fg=tooltip_fg or "#2B2F36",
            )

    def attach_var(self, var):
        """让 StringVar 的每次写入都自动按宽度省略显示。"""
        var.trace_add("write", lambda *_args: self.set_text(var.get()))
        self.set_text(var.get())
        return self

    def set_text(self, text, **label_kwargs):
        self.full_text = "" if text is None else str(text)
        if label_kwargs:
            try:
                self.label.configure(**label_kwargs)
            except tk.TclError:
                return
        self.refresh()

    def refresh(self):
        try:
            if not self.label.winfo_exists():
                return
        except tk.TclError:
            return
        try:
            width = int(self.width_getter() or 0)
        except Exception:
            width = 0
        if width <= 1:
            width = self.fallback_width
        if self.max_lines > 1:
            try:
                current = int(float(self.label.cget("wraplength") or 0))
            except (TypeError, ValueError, tk.TclError):
                current = 0
            if current != width:
                self.label.configure(wraplength=width)
        displayed = elide_middle(self.full_text, self.font, width * self.max_lines)
        try:
            if str(self.label.cget("text") or "") != displayed:
                self.label.configure(text=displayed)
        except tk.TclError:
            pass

    def hide_tooltip(self):
        if self._tooltip is not None:
            self._tooltip.hide()


def center_window_on_parent(window, parent, width=None, height=None):
    """把子窗口放到父窗口可视区域正中央，避免默认落在屏幕左上角。

    子窗口比父窗口更大时也会按中心对齐，并尽量夹在屏幕可见范围内。
    """
    try:
        parent.update_idletasks()
        window.update_idletasks()
    except Exception:
        pass

    if width is None or height is None:
        try:
            current_w = window.winfo_width()
            current_h = window.winfo_height()
        except Exception:
            current_w = current_h = 1
        if width is None:
            width = current_w if current_w > 1 else window.winfo_reqwidth()
        if height is None:
            height = current_h if current_h > 1 else window.winfo_reqheight()

    width = int(width)
    height = int(height)

    try:
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
    except Exception:
        window.geometry(f"{width}x{height}")
        return f"{width}x{height}"

    if parent_w <= 1 or parent_h <= 1:
        window.geometry(f"{width}x{height}")
        return f"{width}x{height}"

    # 允许负偏移：子窗口大于父窗口时仍相对父窗口中心对齐
    x = parent_x + (parent_w - width) // 2
    y = parent_y + (parent_h - height) // 2
    try:
        screen_w = int(parent.winfo_screenwidth())
        screen_h = int(parent.winfo_screenheight())
        if screen_w > 0 and screen_h > 0:
            x = max(0, min(x, max(screen_w - width, 0)))
            y = max(0, min(y, max(screen_h - height, 0)))
    except Exception:
        pass

    geometry = f"{width}x{height}+{x}+{y}"
    # 一次性写入宽高与坐标；Windows 上 withdraw 时 geometry() 读回值可能仍是默认值，
    # 以本次写入值为准，显示后再校正一次即可。
    window.geometry(geometry)
    return geometry


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


def normalize_presets(values):
    """清理预设名称，并在保留顺序的同时去重。"""
    if not isinstance(values, (list, tuple)):
        return []
    result = []
    seen = set()
    for value in values:
        if not isinstance(value, str):
            continue
        value = value.strip()
        if not value or value in seen:
            continue
        result.append(value)
        seen.add(value)
    return result


def presets_from_settings(settings):
    """首次运行使用系统预设；用户保存过空列表时也尊重该设置。"""
    if isinstance(settings, dict) and isinstance(settings.get("presets"), list):
        return normalize_presets(settings["presets"])
    return list(DEFAULT_PRESETS)


def append_presets_to_rule_rows(rows, selected_presets):
    """把尚不存在的预设追加为原文本规则，并返回新增数量。"""
    clean_rows = []
    existing_old_texts = set()
    for row in rows or []:
        values = list(row) if isinstance(row, (list, tuple)) else [row]
        old_text = "" if not values or values[0] is None else str(values[0])
        new_text = "" if len(values) < 2 or values[1] is None else str(values[1])
        if old_text or new_text:
            clean_rows.append([old_text, new_text])
        if old_text:
            existing_old_texts.add(old_text)

    added = 0
    for preset in normalize_presets(selected_presets):
        if preset in existing_old_texts:
            continue
        clean_rows.append([preset, ""])
        existing_old_texts.add(preset)
        added += 1
    return clean_rows, added


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


class CappedScrollbarModel:
    """限制滑块最大视觉长度，同时保持拖动位置与真实滚动范围一致。"""

    def __init__(self, command, display_setter, max_thumb_fraction=0.72):
        self.command = command
        self.display_setter = display_setter
        self.max_thumb_fraction = max_thumb_fraction
        self.actual_first = 0.0
        self.actual_last = 1.0
        self.display_span = max_thumb_fraction

    def set(self, first, last):
        first = float(first)
        last = float(last)
        self.actual_first = first
        self.actual_last = last
        actual_span = max(0.0, min(last - first, 1.0))
        self.display_span = min(actual_span, self.max_thumb_fraction)

        actual_range = max(1.0 - actual_span, 0.0)
        display_range = max(1.0 - self.display_span, 0.0)
        if actual_range <= 1e-9:
            display_first = display_range / 2
        else:
            display_first = (first / actual_range) * display_range
        display_first = max(0.0, min(display_first, display_range))
        self.display_setter(display_first, display_first + self.display_span)

    def dispatch(self, *args):
        if not args:
            return
        if args[0] == "moveto" and len(args) >= 2:
            actual_span = max(self.actual_last - self.actual_first, 0.0)
            actual_range = max(1.0 - actual_span, 0.0)
            display_range = max(1.0 - self.display_span, 0.0)
            if actual_range <= 1e-9 or display_range <= 1e-9:
                return
            display_first = max(0.0, min(float(args[1]), display_range))
            actual_first = (display_first / display_range) * actual_range
            self.command("moveto", actual_first)
            return
        self.command(*args)


class FlatScrollbar(ttk.Scrollbar):
    """带扁平箭头、淡蓝配色和限长滑块的统一滚动条。"""

    def __init__(self, parent, *, command, orient="vertical", **kwargs):
        style_name = (
            "Flat.Vertical.TScrollbar" if orient == "vertical"
            else "Flat.Horizontal.TScrollbar"
        )
        kwargs.setdefault("style", style_name)
        super().__init__(parent, orient=orient, **kwargs)
        self._capped_model = CappedScrollbarModel(command, super().set)
        self.configure(command=self._capped_model.dispatch)

    def set(self, first, last):
        self._capped_model.set(first, last)


def cap_existing_scrollbar(scrollbar, command, scrollable, orientation="vertical"):
    """给第三方控件内部已创建的 ttk 滚动条套用统一样式与限长模型。"""
    style_name = (
        "Flat.Vertical.TScrollbar" if orientation == "vertical"
        else "Flat.Horizontal.TScrollbar"
    )
    scrollbar.configure(style=style_name)
    display_setter = lambda first, last: ttk.Scrollbar.set(scrollbar, first, last)
    model = CappedScrollbarModel(command, display_setter)
    scrollbar.configure(command=model.dispatch)
    if orientation == "vertical":
        scrollable.configure(yscrollcommand=model.set)
    else:
        scrollable.configure(xscrollcommand=model.set)
    scrollbar._capped_model = model
    return model


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
        self._rules_resize_after = None
        self.table_export_window = None
        self.presets = presets_from_settings(load_settings()) if restore_session else list(DEFAULT_PRESETS)
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
        style.configure("TLabel", font=self.body_font)
        style.configure("TButton", font=self.body_font, padding=(10, 5))
        style.configure("TEntry", font=self.body_font)
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
        self._create_busy_button(rules_toolbar, text="新增一行", command=self.add_rule_row, width=9).pack(side="left", padx=(0, 6))
        self._create_busy_button(rules_toolbar, text="删除选中", command=self.delete_selected_rules, width=9).pack(side="left", padx=(0, 6))
        self._create_busy_button(rules_toolbar, text="清空规则", command=self.clear_rules, width=9).pack(side="left", padx=(0, 6))
        self._create_busy_button(rules_toolbar, text="导入 Excel", command=self.import_rules_from_excel, width=10).pack(side="left", padx=(0, 6))
        self._create_busy_button(rules_toolbar, text="导入招标文件", command=self.import_rules_from_tender_file, width=13).pack(side="left", padx=(0, 6))
        self.preset_button = self._create_busy_button(
            rules_toolbar, text="预设 ▼", command=self.toggle_preset_popup, width=8
        )
        self.preset_button.pack(side="left")

        self.rules_hint = ttk.Label(
            rules_frame,
            text="双击单元格可编辑；可从 Excel、招标 Word 或预设添加规则；执行时按原文长度长词优先。",
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

        # 标题行：左标题 + 右「选择目录」按钮
        output_header = ttk.Frame(bottom_frame, style="Toolbar.TFrame")
        output_header.grid(row=0, column=0, sticky="ew")
        output_header.columnconfigure(0, weight=1)
        ttk.Label(output_header, text="3  输出目录", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        self._create_busy_button(output_header, text="选择目录", command=self.select_output_dir, width=12).grid(row=0, column=1, sticky="e")

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

    def _apply_output_dir(self, directory):
        self.output_dir = directory
        if directory:
            self._output_path.set_text(directory, foreground="green")
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

    def _save_session(self):
        if not self.restore_session:
            return
        data = {
            "geometry": self.root.geometry(),
            "output_dir": self.output_dir,
            "rules": self.get_rules_from_table(),
            "presets": self.presets,
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
                if hasattr(self, "_output_path"):
                    self._output_path.hide_tooltip()
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
            from tender_info_extractor import extract_project_info_rules

            return extract_project_info_rules(file_path)

        def on_success(rules):
            self._reset_busy_state()
            if not rules:
                messagebox.showwarning(
                    "提示",
                    "未识别到可导入的项目信息。\n\n"
                    "目前支持项目名称、项目编号、标的名称、采购人信息、开标信息等字段。",
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

        self._apply_output_dir(directory)
        self.status_var.set("已选择输出目录")
        self._refresh_output_and_status_layout()

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


class WordTableExportWindow:
    WINDOW_WIDTH = 1180
    WINDOW_HEIGHT = 800

    def __init__(self, app, initial_files=None):
        self.app = app
        self.window = tk.Toplevel(app.root)
        self._task_runner = BackgroundTaskRunner(self.window)
        self._busy_widgets = []
        self.window.title("提取 Word 表格到 Excel")
        # Give the scan list enough room on first open.  The result list is the
        # primary workspace in this window, so it should not start out cramped
        # by the fixed-height sections around it.
        self.window.minsize(900, 640)
        self.window.configure(bg=app.app_bg)
        self.window.transient(app.root)
        # 先藏起来，排完布局再相对主窗口居中显示，避免闪到屏幕左上角
        self.window.withdraw()
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
        self._scan_separator_after = None
        self.status_var = tk.StringVar(value="就绪")

        self._create_widgets()
        if initial_files:
            self._append_files(initial_files, show_status=False)
            self.status_var.set(f"已自动带入 {len(self.file_paths)} 个 Word 文件")
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        geometry = center_window_on_parent(
            self.window,
            app.root,
            width=self.WINDOW_WIDTH,
            height=self.WINDOW_HEIGHT,
        )
        self.window.deiconify()
        if geometry:
            self.window.geometry(geometry)
        else:
            center_window_on_parent(
                self.window,
                app.root,
                width=self.WINDOW_WIDTH,
                height=self.WINDOW_HEIGHT,
            )
        self.window.lift()
        self.window.focus_force()

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
        ttk.Label(header, text="提取 Word 表格到 Excel", style="Title.TLabel").pack(anchor="w")
        self.header_hint = ttk.Label(
            header,
            text="选择 .docx 文件；每个 Word 生成一个 Excel，每张表格对应一个工作表。",
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
            columns=("selected", "file", "section", "table", "size"),
            show="headings",
            selectmode="extended",
            style="Scan.Treeview",
            height=10,
        )
        self.scan_tree.heading("selected", text="导出")
        self.scan_tree.heading("file", text="文件")
        self.scan_tree.heading("section", text="所在章节")
        self.scan_tree.heading("table", text="表格")
        self.scan_tree.heading("size", text="行列")
        self.scan_tree.column("selected", width=68, minwidth=60, anchor="center", stretch=False)
        self.scan_tree.column("file", width=250, minwidth=160, stretch=False)
        self.scan_tree.column("section", width=600, minwidth=280)
        self.scan_tree.column("table", width=70, minwidth=62, anchor="center", stretch=False)
        self.scan_tree.column("size", width=72, minwidth=62, anchor="center", stretch=False)
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
        self.scan_x_scrollbar.grid(row=1, column=0, sticky="ew")
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
        self.scan_tree.bind("<ButtonRelease-1>", self._schedule_scan_tree_column_separators, add="+")
        self.scan_tree.bind("<Configure>", self._schedule_scan_tree_column_separators, add="+")
        self.scan_tree.bind("<Double-1>", self._toggle_scan_row_from_event)
        self.scan_tree.bind("<space>", self._toggle_scan_rows_from_keyboard)
        self.scan_tree.bind("<<TreeviewSelect>>", self._update_scan_detail)
        self._schedule_scan_tree_column_separators()

        detail_frame = ttk.Frame(scan_frame, style="Toolbar.TFrame")
        detail_frame.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        detail_frame.columnconfigure(0, weight=1)
        self.detail_text = tk.Text(
            detail_frame,
            height=2,
            wrap="word",
            font=self.app.small_font,
            bg="#F1F3F6",
            fg=self.app.text_fg,
            relief="solid",
            borderwidth=1,
            padx=8,
            pady=4,
        )
        self.detail_text.grid(row=0, column=0, sticky="ew")
        self.detail_text.insert("1.0", "选择扫描结果中的一行，可在这里查看完整章节、表格前文和内容预览。")
        self.detail_text.configure(state="disabled")

        self.scan_label = ttk.Label(scan_frame, text="请先添加 Word 文件并扫描表格", style="Muted.TLabel")
        self.scan_label.grid(row=3, column=0, sticky="ew", pady=(5, 0))

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
        self._schedule_scan_tree_column_separators()

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

    def _update_scan_tree_column_separators(self):
        """Overlay continuous grid lines at column edges, including the header row."""
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

            # tree_frame 坐标系：对齐到 Treeview 左上角，整列贯穿表头+数据区
            origin_x = self.scan_tree.winfo_x()
            origin_y = self.scan_tree.winfo_y()
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
            header_height = self._scan_tree_header_height()
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
        self._schedule_scan_tree_column_separators()

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
            self._schedule_scan_tree_column_separators()
        if hasattr(self, "scan_label"):
            self.scan_label.config(text=label_text, foreground=self.app.muted_fg)
        if hasattr(self, "detail_text"):
            self._set_detail_text("选择扫描结果中的一行，可在这里查看完整章节、表格前文和内容预览。")

    def _scan_tree_values(self, item, selected):
        return (
            "☑" if selected else "☐",
            item.filename,
            item.section,
            f"表格{item.table_index}",
            f"{item.row_count}x{item.column_count}",
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
