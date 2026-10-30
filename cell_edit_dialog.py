"""与主界面共用颜色、字体和按钮样式的单元格文字编辑窗口。"""
import tkinter as tk
from tkinter import ttk

from ui_common import center_window_on_parent


class CellEditDialog:
    def __init__(self, parent, app, ref, value):
        self.parent = parent
        self.result = None
        self.previous_grab = parent.grab_current()
        self.window = tk.Toplevel(parent)
        self.window.withdraw()
        self.window.title(f'编辑单元格 {ref}')
        self.window.configure(bg=app.surface_bg)
        self.window.resizable(True, True)
        self.window.minsize(360, 220)
        self.window.transient(parent)
        try:
            self.window.iconphoto(True, app._icon_photo)
        except (AttributeError, tk.TclError):
            pass
        body = ttk.Frame(self.window, padding=16, style='Toolbar.TFrame')
        body.pack(fill='both', expand=True)
        self.text = tk.Text(body, width=48, height=9, font=app.body_font, wrap='word',
                            bg=app.surface_bg, fg=app.text_fg, relief='solid', borderwidth=1,
                            highlightthickness=0, insertbackground=app.text_fg, padx=8, pady=6)
        self.text.pack(fill='both', expand=True)
        self.text.insert('1.0', value)
        ttk.Label(body, text='Ctrl+Enter 应用，Esc 取消；留空则清除该格内容', style='Subtitle.TLabel').pack(anchor='w', pady=(8, 0))
        actions = ttk.Frame(body, style='Toolbar.TFrame')
        actions.pack(anchor='e', pady=(10, 0))
        self.enlarge_button = ttk.Button(actions, text='放大', width=7, command=self.toggle_size)
        self.enlarge_button.pack(side='left', padx=(0, 8))
        ttk.Button(actions, text='取消', width=7, command=self.close).pack(side='left', padx=(0, 8))
        ttk.Button(actions, text='应用', width=7, style='Accent.TButton', command=self.apply).pack(side='left')
        self.window.bind('<Escape>', self.close)
        self.window.bind('<Control-Return>', self.apply)
        self.window.bind('<F11>', self.toggle_size)
        self.text.bind('<Control-Return>', self.apply)
        self._enlarged = False
        self._normal_geometry = None
        self.window.protocol('WM_DELETE_WINDOW', self.close)
        self.window.update_idletasks()
        center_window_on_parent(self.window, parent, max(440, self.window.winfo_reqwidth()), self.window.winfo_reqheight())

    def show(self):
        self.window.deiconify()
        self.window.wait_visibility()
        self.window.grab_set()
        self.text.focus_set()
        self.window.wait_window()
        return self.result

    def toggle_size(self, _event=None):
        """在大窗口编辑与原尺寸之间切换；窗口本身也可手动拖边调整。"""
        if self._enlarged:
            self.window.geometry(self._normal_geometry)
            self.enlarge_button.configure(text='放大')
            self._enlarged = False
        else:
            self._normal_geometry = self.window.geometry()
            screen_w, screen_h = self.window.winfo_screenwidth(), self.window.winfo_screenheight()
            width, height = int(screen_w * .8), int(screen_h * .8)
            self.window.geometry(f'{width}x{height}+{(screen_w - width) // 2}+{max(0,(screen_h - height) // 2)}')
            self.enlarge_button.configure(text='还原')
            self._enlarged = True
        self.text.focus_set()
        return 'break'

    def apply(self, _event=None):
        self.result = self.text.get('1.0', 'end-1c')
        self.close()
        return 'break'

    def close(self, _event=None):
        self.window.destroy()
        if self.previous_grab is not None and self.previous_grab.winfo_exists():
            self.previous_grab.grab_set()
        if self.parent.winfo_exists():
            self.parent.focus_set()
        return 'break'


def ask_cell_text(parent, app, ref, value):
    return CellEditDialog(parent, app, ref, value).show()
