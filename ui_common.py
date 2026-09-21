"""两个窗口共用的控件、布局和后台任务。"""

import os
import queue
import sys
import threading
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import messagebox, ttk
from app_settings import error_log_path

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
SUPPORTED_EXTENSIONS = (".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".ppt", ".pptx")
WORD_EXTENSIONS = (".doc", ".docx")
# 指标参数提取的默认符号，与 symbol_clause_extractor.DEFAULT_SYMBOL_CHARS 保持一致。
# 这里单独定义一份是为了避免启动时加载 docx 依赖（symbol_clause_extractor 会引入它）。
DEFAULT_SYMBOL_CHARS = "★#△▲"

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


class CanvasCheckbox:
    """手绘复选框：方框 + 蓝色对勾。

    ttk 的 clam 主题不支持自定义选中标记（画出来是叉 ✕），与「预设」
    下拉面板保持一致，全部用 Canvas 自绘对勾 ✓。variable 变化时自动
    重绘，外部直接改 var 也能同步显示。
    """

    def __init__(
        self,
        parent,
        text,
        variable,
        app,
        command=None,
        font=None,
        background=None,
        padx=10,
        pady=7,
        box_size=14,
    ):
        self.text = text
        self.variable = variable
        self.app = app
        self.command = command
        self.box_size = box_size
        self.padx = padx
        self.background = background or app.surface_bg
        self.font = font or app.body_font

        try:
            text_width = self.font.measure(text)
            line_space = self.font.metrics("linespace")
        except tk.TclError:
            text_width = 60
            line_space = 18
        self.height = max(box_size + 2 * pady, line_space + 2 * pady)

        self.canvas = tk.Canvas(
            parent,
            width=padx + box_size + 8 + text_width + padx,
            height=self.height,
            bg=self.background,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
        )
        self._hover = False
        variable.trace_add("write", lambda *_args: self.redraw())
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<Enter>", lambda _event: self._set_hover(True))
        self.canvas.bind("<Leave>", lambda _event: self._set_hover(False))
        self.redraw()

    def _set_hover(self, hover):
        if self._hover != hover:
            self._hover = hover
            self.redraw()

    def _on_click(self, _event=None):
        self.variable.set(not self.variable.get())
        if self.command:
            self.command()

    def toggle(self):
        self._on_click()

    def redraw(self):
        canvas = self.canvas
        selected = bool(self.variable.get())
        if selected:
            background = "#E3EDF8"
        elif self._hover:
            background = "#EDF1F6"
        else:
            background = self.background
        try:
            canvas.configure(bg=background)
        except tk.TclError:
            return

        box = self.box_size
        x0 = self.padx
        y0 = (self.height - box) // 2
        canvas.delete("all")
        canvas.create_rectangle(
            x0, y0, x0 + box, y0 + box,
            outline=self.app.accent_fg if selected else "#6F7B88",
            width=1,
            fill="#F7F8FA",
        )
        if selected:
            # 对勾三个点按 13px 方框的比例缩放，加圆角端点
            scale = box / 13.0
            canvas.create_line(
                x0 + 3 * scale, y0 + 6 * scale,
                x0 + 6.5 * scale, y0 + 10 * scale,
                x0 + 11 * scale, y0 + 3 * scale,
                fill=self.app.accent_fg,
                width=2,
                capstyle="round",
                joinstyle="round",
            )
        canvas.create_text(
            x0 + box + 8, self.height // 2,
            text=self.text,
            anchor="w",
            font=self.font,
            fill=self.app.text_fg,
        )


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


def symbol_chars_from_settings(settings):
    """读取自定义提取符号；保存时已清洗过，这里只做基本类型校验。"""
    value = settings.get("symbol_chars") if isinstance(settings, dict) else None
    if isinstance(value, str) and value.strip():
        return value
    return DEFAULT_SYMBOL_CHARS


def keep_clause_symbols_from_settings(settings):
    """条款正文是否保留标记符号；未保存过时默认不保留，只写入符号列。"""
    if isinstance(settings, dict) and "keep_clause_symbols" in settings:
        return bool(settings["keep_clause_symbols"])
    return False


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


def cancel_widget_callbacks(event):
    """销毁时取消控件的延迟回调；根窗口关闭时清空所属解释器的任务。"""
    widget = event.widget
    commands = set(widget._tclCommands or ())
    for task in widget.tk.splitlist(widget.tk.call("after", "info")):
        script, _kind = widget.tk.call("after", "info", task)
        if isinstance(widget, tk.Tk) or widget.tk.splitlist(script)[0] in commands:
            widget.after_cancel(task)


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

