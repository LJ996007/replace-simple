"""统一应用弹窗；保留常用 tkinter 调用接口，避免新增原生灰色窗口。"""
import math
import tkinter as tk
from tkinter import ttk


class AppDialog:
    def __init__(self, title, prompt, parent=None, kind=None, initialvalue=None, minvalue=None, maxvalue=None):
        root = parent or tk._default_root
        if root is None:raise RuntimeError('请先创建主窗口。')
        self.previous = root.grab_current()
        parent = parent or self.previous or root
        self.parent,self.kind,self.minimum,self.maximum = parent,kind,minvalue,maxvalue
        self.result = None
        self.window = tk.Toplevel(parent);self.window.withdraw()
        self.window.title(title);self.window.transient(parent)
        self.window.resizable(False,False)
        bg = ttk.Style(parent).lookup('TLabel','background') or '#F7F8FA'
        self.window.configure(bg=bg)
        frame = ttk.Frame(self.window,padding=20,style='Toolbar.TFrame');frame.pack(fill='both',expand=True)
        if kind:
            ttk.Label(frame,text=prompt,wraplength=420).pack(anchor='w',pady=(0,12))
            self.value = tk.StringVar(value='' if initialvalue is None else str(initialvalue))
            self.entry = ttk.Entry(frame,textvariable=self.value,width=32)
            self.entry.pack(fill='x')
            self.error = tk.StringVar()
            tk.Label(frame,textvariable=self.error,bg=bg,fg='#B42318',anchor='w').pack(fill='x',pady=(5,10))
            self.value.trace_add('write',lambda *_:self.error.set(''))
        else:
            content = ttk.Frame(frame,style='Toolbar.TFrame');content.pack(fill='both',expand=True,pady=(0,18))
            if len(prompt)>240 or prompt.count('\n')>5:
                text = tk.Text(content,width=54,height=12,wrap='word',relief='flat',bg=bg,
                               fg=ttk.Style(parent).lookup('TLabel','foreground'),font=ttk.Style(parent).lookup('TLabel','font'))
                scroll=ttk.Scrollbar(content,command=text.yview);scroll.pack(side='right',fill='y')
                text.configure(yscrollcommand=scroll.set);text.pack(fill='both',expand=True)
                text.insert('1.0',prompt);text.configure(state='disabled')
            else:ttk.Label(content,text=prompt,wraplength=440,justify='left').pack(anchor='w')
        actions=ttk.Frame(frame,style='Toolbar.TFrame');actions.pack(anchor='e')
        if kind:ttk.Button(actions,text='取消',command=self.close).pack(side='left',padx=(0,8))
        ttk.Button(actions,text='应用' if kind else '确定',style='Accent.TButton',command=self.apply).pack(side='left')
        self.window.bind('<Return>',self.apply);self.window.bind('<Escape>',self.close)
        self.window.protocol('WM_DELETE_WINDOW',self.close)
        from ui_common import center_window_on_parent
        self.window.update_idletasks()
        center_window_on_parent(self.window,parent,max(380,self.window.winfo_reqwidth()),self.window.winfo_reqheight())

    def apply(self,_event=None):
        if self.kind:
            try:
                raw=self.value.get().strip()
                value = int(raw) if self.kind=='integer' else float(raw) if self.kind=='float' else raw
                if self.kind!='string' and (not math.isfinite(value) or self.minimum is not None and value<self.minimum or self.maximum is not None and value>self.maximum):raise ValueError()
            except (ValueError,OverflowError):
                label='整数' if self.kind=='integer' else '数值'
                bounds=f'{self.minimum}–{self.maximum} 之间的' if self.minimum is not None and self.maximum is not None else '有效的'
                self.error.set(f'请输入{bounds}{label}');self.entry.focus_set();return 'break'
            self.result=value
        else:self.result='ok'
        self.close();return 'break'

    def close(self,_event=None):
        self.window.destroy()
        if self.previous is not None and self.previous.winfo_exists():self.previous.grab_set()
        if self.parent.winfo_exists():self.parent.focus_set()
        return 'break'

    def show(self):
        self.window.deiconify();self.window.wait_visibility();self.window.grab_set()
        if self.kind:self.entry.focus_set();self.entry.selection_range(0,'end')
        else:self.window.focus_set()
        self.window.wait_window();return self.result


class _Messages:
    def showinfo(self,title,message,**kwargs):return AppDialog(title,str(message),parent=kwargs.get('parent')).show()
    showwarning=showerror=showinfo


class _Inputs:
    def askinteger(self,title,prompt,**kwargs):return AppDialog(title,prompt,kind='integer',**kwargs).show()
    def askfloat(self,title,prompt,**kwargs):return AppDialog(title,prompt,kind='float',**kwargs).show()
    def askstring(self,title,prompt,**kwargs):return AppDialog(title,prompt,kind='string',**kwargs).show()


messagebox=_Messages()
simpledialog=_Inputs()
