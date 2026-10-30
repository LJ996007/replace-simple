"""与主界面共用颜色、字体和按钮样式的行高输入窗口。"""
import math
import tkinter as tk
from tkinter import ttk

from ui_common import center_window_on_parent


class RowHeightDialog:
    def __init__(self, parent, app, initialvalue, first_row, last_row):
        self.parent = parent
        self.result = None
        self.previous_grab = parent.grab_current()
        self.window = tk.Toplevel(parent)
        self.window.withdraw()
        scope = f'第 {first_row} 行' if first_row == last_row else f'第 {first_row}–{last_row} 行'
        self.window.title(f'设置行高 · {scope}')
        self.window.configure(bg=app.surface_bg)
        self.window.resizable(False, False)
        self.window.transient(parent)
        try:
            self.window.iconphoto(True, app._icon_photo)
        except (AttributeError, tk.TclError):
            pass
        body = ttk.Frame(self.window, padding=20, style='Toolbar.TFrame')
        body.pack(fill='both', expand=True)
        body.columnconfigure(1, weight=1)
        ttk.Label(body, text='行高').grid(row=0, column=0, sticky='w', padx=(0, 14))
        self.value = tk.StringVar(value=f'{initialvalue:g}')
        self.entry = ttk.Entry(body, textvariable=self.value, width=18, font=app.body_font)
        self.entry.grid(row=0, column=1, sticky='ew')
        ttk.Label(body, text='磅').grid(row=0, column=2, padx=(10, 0))
        self.error = tk.StringVar()
        tk.Label(body, textvariable=self.error, anchor='w', font=app.small_font,
                 background=app.surface_bg, foreground='#B42318').grid(
                     row=1, column=1, columnspan=2, sticky='ew', pady=(6, 12))
        actions = ttk.Frame(body, style='Toolbar.TFrame')
        actions.grid(row=2, column=0, columnspan=3, sticky='e')
        ttk.Button(actions, text='取消', width=7, command=self.close).pack(side='left', padx=(0, 8))
        self.apply_button = ttk.Button(actions, text='应用', width=7, style='Accent.TButton', command=self.apply)
        self.apply_button.pack(side='left')
        self.window.bind('<Return>', self.apply)
        self.window.bind('<Escape>', self.close)
        self.window.protocol('WM_DELETE_WINDOW', self.close)
        self.value.trace_add('write', lambda *_: self.error.set(''))
        self.window.update_idletasks()
        center_window_on_parent(self.window, parent, max(360, self.window.winfo_reqwidth()), self.window.winfo_reqheight())

    def show(self):
        self.window.deiconify()
        self.window.wait_visibility()
        self.window.grab_set()
        self.entry.focus_set()
        self.entry.selection_range(0, 'end')
        self.window.wait_window()
        return self.result

    def apply(self, _event=None):
        try:
            value = float(self.value.get().strip())
            if not math.isfinite(value) or not 1 <= value <= 409:
                raise ValueError()
        except ValueError:
            self.error.set('请输入 1–409 之间的数值')
            self.entry.focus_set()
            self.entry.selection_range(0, 'end')
            return 'break'
        self.result = value
        self.close()
        return 'break'

    def close(self, _event=None):
        self.window.destroy()
        if self.previous_grab is not None and self.previous_grab.winfo_exists():
            self.previous_grab.grab_set()
        if self.parent.winfo_exists():
            self.parent.focus_set()
        return 'break'


def ask_row_height(parent, app, initialvalue, first_row, last_row):
    return RowHeightDialog(parent, app, initialvalue, first_row, last_row).show()
