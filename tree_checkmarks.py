"""在 Treeview 数据列中显示图标，保留原列坐标、选择语义和键盘操作。

Tk 8.6 仅支持树列图片，因此复用少量可见单元格覆盖层，不为每条数据创建控件。
原单元格的三态值仍供程序读取，覆盖层的鼠标事件交回 Treeview。
"""
import tkinter as tk
from tkinter import ttk
from PIL import ImageTk
from ui_icons import checkbox_icon


class TreeviewCheckmarks:
    STATES = {'☐': 'unchecked', '☑': 'checked', '◩': 'mixed'}

    def __init__(self, tree, column, header_height, on_painted=None):
        self.tree, self.column = tree, column
        self.header_height = header_height
        self.on_painted = on_painted
        self.header_mark = '☐'
        self.parent = tree.master
        self.style = ttk.Style(tree)
        self.layers, self.images = [], {}
        self.pending = None
        row_height = int(self.style.lookup(tree.cget('style') or 'Treeview', 'rowheight') or 32)
        self.size = min(max(16, round(18 * tree.winfo_fpixels('1i') / 96)), max(12, round(row_height*.7)))
        self.header_layer = self.make_layer()
        for event in ('<<TreeviewSelect>>', '<<TreeviewOpen>>', '<<TreeviewClose>>',
                      '<Configure>', '<FocusIn>', '<FocusOut>'):
            tree.bind(event, self.schedule, add='+')
        tree.bind('<Destroy>', self.destroy, add='+')
        previous = tree.cget('yscrollcommand')
        def scrolled(first, last):
            if previous:tree.tk.call(previous, first, last)
            self.schedule()
        tree.configure(yscrollcommand=scrolled)

    def icon(self, mark, color):
        key = (mark, color)
        if key not in self.images:
            self.images[key] = ImageTk.PhotoImage(
                checkbox_icon(self.size, color, self.STATES.get(mark, 'unchecked')), master=self.tree)
        return self.images[key]

    def schedule(self, event=None):
        if self.pending is None:self.pending = self.tree.after_idle(self.refresh)

    def forward(self, event, sequence):
        self.tree.event_generate(sequence, x=event.x_root-self.tree.winfo_rootx(),
                                 y=event.y_root-self.tree.winfo_rooty(), state=event.state,
                                 **({'delta': event.delta} if sequence=='<MouseWheel>' else {}))
        return 'break'

    def make_layer(self):
        layer = tk.Canvas(self.parent, highlightthickness=0, borderwidth=0, takefocus=False)
        layer._icon_item = layer.create_image(0,0)
        layer._paint_state = None
        for sequence in ('<ButtonPress-1>', '<ButtonRelease-1>', '<B1-Motion>', '<MouseWheel>'):
            layer.bind(sequence, lambda event, seq=sequence:self.forward(event,seq))
        return layer

    def paint(self, layer, box, mark, bg, fg, viewport):
        x,top,cell_width,cell_height = box
        width,height = viewport
        # 给两侧列边线留出空间，表头与行内使用同一列中心。
        left,right = max(1,x+1),min(width-1,x+cell_width-2)
        bottom = min(top+cell_height,height-1)
        if right<=left or bottom<=top:
            layer.place_forget();layer._paint_state=None
            return
        state=(box,mark,bg,fg,viewport,self.tree.winfo_x(),self.tree.winfo_y())
        if layer._paint_state == state:return
        layer._paint_state = state
        layer.configure(bg=bg)
        layer.itemconfigure(layer._icon_item,image=self.icon(mark,fg))
        layer.coords(layer._icon_item,x+cell_width/2-left,cell_height/2)
        layer.place(x=self.tree.winfo_x()+left,y=self.tree.winfo_y()+top,
                    width=right-left,height=bottom-top)
        tk.Misc.lift(layer)

    def refresh(self):
        if self.pending is not None:
            self.tree.after_cancel(self.pending)
            self.pending = None
        tree = self.tree
        height, width = tree.winfo_height(), tree.winfo_width()
        if height<=1 or width<=1:return
        selected = set(tree.selection())
        style = tree.cget('style') or 'Treeview'
        colors = {active: (self.style.lookup(style,'background',state) or 'white',
                          self.style.lookup(style,'foreground',state) or 'black')
                  for active,state in ((False,()),(True,('selected',)))}
        header_height = self.header_height()
        used, seen, y = 0, set(), header_height
        column_box = None
        # identify_row 只扫描视口；展开、折叠与滚动不遍历整棵树。
        while y < height:
            iid = tree.identify_row(y)
            if not iid or iid in seen:
                if not iid and (seen or y>header_height+2):break
                y += 1
                continue
            seen.add(iid)
            box = tree.bbox(iid, self.column)
            if not box:
                y += 1
                continue
            x, top, cell_width, cell_height = box
            column_box = (x,cell_width)
            y = max(y+1, top+cell_height)
            mark = tree.set(iid, self.column)
            if mark not in self.STATES:continue
            if x+cell_width<=1 or x>=width-1:continue
            bg,fg = colors[iid in selected]
            if used == len(self.layers):
                self.layers.append(self.make_layer())
            layer = self.layers[used];used += 1
            self.paint(layer,box,mark,bg,fg,(width,height))
        for layer in self.layers[used:]:
            layer.place_forget();layer._paint_state=None
        if column_box is None:
            columns=['#0',*tree.cget('columns')]
            sizes=[int(tree.column(c,'width')) for c in columns]
            index=columns.index(self.column)
            column_box=(1+sum(sizes[:index])-round(tree.xview()[0]*sum(sizes)),sizes[index])
        x,cell_width=column_box
        heading_style=style+'.Heading'
        self.paint(self.header_layer,(x,1,cell_width,header_height-2),self.header_mark,
                   self.style.lookup(heading_style,'background') or '#E8EBEF',
                   self.style.lookup(heading_style,'foreground') or 'black',(width,height))
        if self.on_painted:self.on_painted()

    def destroy(self, event):
        if event.widget is not self.tree:return
        if self.pending is not None:self.tree.after_cancel(self.pending)
        self.pending = None
        for layer in self.layers:layer.destroy()
        self.header_layer.destroy()
        self.layers.clear();self.images.clear()
