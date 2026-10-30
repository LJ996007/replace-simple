"""保留合并结构的只读表格画布；选择、拖放和预览共用格坐标。"""
import bisect
import time
from collections import OrderedDict
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

from openpyxl.utils.cell import get_column_letter, range_boundaries
from ui_common import FlatScrollbar
from visual_fill import expand_selection


class VisualTable(ttk.Frame):
    def __init__(self, parent, app, on_select=None, on_drag=None, on_drop=None, on_cancel=None, on_resize=None, on_edit=None):
        super().__init__(parent,style="Toolbar.TFrame")
        self.app,self.on_select,self.on_drag,self.on_drop = app,on_select,on_drag,on_drop
        self.on_cancel = on_cancel
        self.on_resize = on_resize
        self.on_edit = on_edit
        self._click = None
        self.info = None
        self.selection = None
        self.mark = None
        self.updates = {}
        self.zoom = 1.0
        self.display_scale=max(1,float(self.winfo_fpixels('1i'))/96)
        self.zoom_label = tk.StringVar(value='100%')
        self.header_width,self.header_height = 42,30
        self.columnconfigure(0,weight=1)
        self.rowconfigure(0,weight=1)
        self.canvas = tk.Canvas(self,bg=app.surface_bg,highlightthickness=1,
                                highlightbackground=app.border_color)
        self.canvas.grid(row=0,column=0,sticky="nsew")
        vertical = FlatScrollbar(self,orient="vertical",command=self.scroll_y,style="Flat.Vertical.TScrollbar")
        horizontal = FlatScrollbar(self,orient="horizontal",command=self.scroll_x,style="Flat.Horizontal.TScrollbar")
        vertical.grid(row=0,column=1,sticky="ns")
        horizontal.grid(row=1,column=0,sticky="ew")
        self.canvas.configure(yscrollcommand=vertical.set,xscrollcommand=horizontal.set)
        self.canvas.bind("<Configure>",lambda _:self.request_draw())
        self.canvas.bind("<MouseWheel>",self.wheel)
        self.canvas.bind("<Shift-MouseWheel>",lambda e:self.wheel(e,True))
        self.canvas.bind("<Control-MouseWheel>",self.zoom_wheel)
        self.canvas.bind("<ButtonPress-1>",self.press)
        self.canvas.bind("<B1-Motion>",self.motion)
        self.canvas.bind("<ButtonRelease-1>",self.release)
        self.canvas.bind('<Motion>',self.resize_cursor)
        self.canvas.bind("<Escape>",lambda _:self.cancel())
        self.canvas.bind("<Left>",lambda _:self.navigate(0,-1))
        self.canvas.bind("<Right>",lambda _:self.navigate(0,1))
        self.canvas.bind("<Up>",lambda _:self.navigate(-1,0))
        self.canvas.bind("<Down>",lambda _:self.navigate(1,0))
        self.gesture = None
        self.font = (app.font_family,10)
        self.text_fonts = {False:tkfont.Font(family=app.font_family,size=10),
                           True:tkfont.Font(family=app.font_family,size=10,weight='bold')}
        self.cell_fonts = {}
        self._frame = None
        self._font_zoom = None
        self._geometry_dirty = False
        self._text_cache = OrderedDict()
        self._glyph_cache = {}
        self.bind('<Destroy>',self._dispose,add='+')

    def _dispose(self,event):
        if event.widget is self and self._frame is not None:
            self.after_cancel(self._frame);self._frame=None

    def request_draw(self):
        if self._frame is None:self._frame=self.after(16,self._paint)

    def _paint(self):
        self._frame=None
        self.draw()

    def measure_glyph(self,font,ch):
        key=(str(font),ch)
        if key not in self._glyph_cache:
            if len(self._glyph_cache)>=4096:self._glyph_cache.clear()
            self._glyph_cache[key]=font.measure(ch)
        return self._glyph_cache[key]

    def visible_text(self,text,width,height,font,wrap):
        key=(text,round(width,2),round(height,2),str(font),wrap)
        if key in self._text_cache:
            self._text_cache.move_to_end(key);return self._text_cache[key]
        if wrap:result=self.clipped_text(text,width,height,font=font)
        else:
            line=text.split('\n',1)[0]
            lo,hi=0,len(line)
            while lo<hi:
                mid=(lo+hi+1)//2
                if font.measure(line[:mid])<=width:lo=mid
                else:hi=mid-1
            result=line[:lo]
        self._text_cache[key]=result
        if len(self._text_cache)>128:self._text_cache.popitem(last=False)
        return result

    def clipped_text(self,text,width,height,bold=False,font=None):
        font = font or self.text_fonts[bold]
        limit = max(1,int(height/font.metrics('linespace')))
        lines,line,used = [],'',0
        for ch in text:
            advance = self.measure_glyph(font,ch)
            if ch=='\n' or (line and used+advance>width):
                lines.append(line)
                if len(lines)>=limit:
                    last = lines[-1]
                    while last and font.measure(last+'…')>width:
                        last = last[:-1]
                    lines[-1] = last+'…'
                    return '\n'.join(lines)
                line,used = '',0
                if ch=='\n':continue
            line += ch;used += advance
        lines.append(line)
        return '\n'.join(lines[:limit])

    def load(self, info, updates=None):
        self.info = info
        self.updates = updates or {}
        self.selection = None
        self.mark = None
        # 初始提供模板末尾额外空行，可直接拖入；不改变文件行数。
        rows = max(1,info['max_row'])
        cols = max(1,info['max_col'])
        if rows*cols>300000:
            raise ValueError("表格区域过大，请先清理 Excel 中多余的空白格式区域。")
        self.merged = [range_boundaries(ref) for ref in info['merges']]
        # 每 64 行一组，避免一个跨数万行的合并格生成数万个列表。
        self._merge_bands={}
        for merged in self.merged:
            for band in range(merged[1]//64,merged[3]//64+1):
                self._merge_bands.setdefault(band,[]).append(merged)
        self.rebuild_geometry()

    def rebuild_geometry(self):
        if not self.info:return
        self._geometry_dirty = False
        info = self.info
        rows,cols = info['max_row'],info['max_col']
        scale=self.zoom*self.display_scale
        self.header_width,self.header_height = 42*scale,30*scale
        self.font = (self.app.font_family,max(7,round(10*self.zoom)))
        if self._font_zoom!=self.zoom:
            for font in self.text_fonts.values():font.configure(size=max(6,round(10*self.zoom)))
            self.cell_fonts.clear();self._text_cache.clear();self._glyph_cache.clear()
            self._font_zoom=self.zoom
        self.xs,self.ys = [self.header_width],[self.header_height]
        for c in range(1,cols+1):
            width = info.get('widths',{}).get(c,info.get('default_width',160))*scale
            self.xs.append(self.xs[-1]+width)
        for r in range(1,rows+1):
            height = info.get('heights',{}).get(r,info.get('default_height',64))*scale
            self.ys.append(self.ys[-1]+height)
        self.canvas.configure(scrollregion=(0,0,self.xs[-1],self.ys[-1]))
        self.request_draw()

    def set_zoom(self,value,x=None,y=None):
        value = max(.5,min(2.5,round(value,2)))
        if value==self.zoom:return
        cv = self.canvas
        x = self.header_width if x is None else x
        y = self.header_height if y is None else y
        content_x = (cv.canvasx(x)-self.header_width)/self.zoom
        content_y = (cv.canvasy(y)-self.header_height)/self.zoom
        self.zoom = value
        self.zoom_label.set(f'{round(value*100)}%')
        self.rebuild_geometry()
        # 以指针所在内容为中心缩放，保持用户当前查看的位置。
        if self.info:
            cv.xview_moveto(max(0,self.header_width+content_x*value-x)/max(1,self.xs[-1]))
            cv.yview_moveto(max(0,self.header_height+content_y*value-y)/max(1,self.ys[-1]))
        self.draw()

    def zoom_wheel(self,event):
        self.set_zoom(self.zoom+(.1 if event.delta>0 else -.1),event.x,event.y)
        return 'break'

    def add_zoom_controls(self,parent):
        """高分屏按实际 DPI 渲染，四倍采样消除圆弧和斜线锯齿。"""
        from PIL import ImageTk
        from ui_icons import zoom_icon
        bar = ttk.Frame(parent,style='Toolbar.TFrame')
        self.zoom_buttons = []
        self.zoom_tooltips = []
        self.zoom_images = []
        dpi = max(1,float(self.winfo_fpixels('1i'))/96)
        side,icon_size = round(32*dpi),round(24*dpi)
        for sign,delta in (('−',-.1),('+',.1)):
            button = tk.Canvas(bar,width=side,height=side,bg=self.app.surface_bg,
                               highlightthickness=1,highlightbackground=self.app.surface_bg,
                               highlightcolor=self.app.accent_fg,cursor='hand2',takefocus=True)
            icon = ImageTk.PhotoImage(zoom_icon(icon_size,self.app.accent_fg,delta>0),master=button)
            self.zoom_images.append(icon)
            button.create_image(side/2,side/2,image=icon)
            button.bind('<Enter>',lambda _,b=button:b.configure(bg='#E3EDF8'))
            button.bind('<Leave>',lambda _,b=button:b.configure(bg=self.app.surface_bg))
            command = lambda _,d=delta:self.set_zoom(self.zoom+d)
            button.bind('<Button-1>',command);button.bind('<Return>',command);button.bind('<space>',command)
            button.pack(side='left',padx=2)
            self.zoom_buttons.append(button)
            from ui_common import HoverTooltip
            label = '缩小预览（Ctrl + 滚轮）' if delta<0 else '放大预览（Ctrl + 滚轮）'
            self.zoom_tooltips.append(HoverTooltip(button,lambda text=label:text,self.app.body_font))
        ttk.Button(bar,textvariable=self.zoom_label,style='Header.TButton',width=6,
                   command=lambda:self.set_zoom(1)).pack(side='left',padx=(4,0))
        return bar

    def cell_at(self, x, y):
        if not self.info:
            return None
        if x < self.header_width or y < self.header_height:
            return None
        x,y = self.canvas.canvasx(x),self.canvas.canvasy(y)
        c,r = bisect.bisect_right(self.xs,x),bisect.bisect_right(self.ys,y)
        if not 1<=r<=self.info['max_row'] or not 1<=c<=self.info['max_col']:
            return None
        return r,c

    def cell_at_screen(self,x,y):
        x,y = x-self.canvas.winfo_rootx(),y-self.canvas.winfo_rooty()
        if not 0<=x<self.canvas.winfo_width() or not 0<=y<self.canvas.winfo_height():
            return None
        return self.cell_at(x,y)

    def bounds(self,r,c):
        return next((m for m in self._merge_bands.get(r//64,()) if m[0]<=c<=m[2] and m[1]<=r<=m[3]),(c,r,c,r))

    def select(self,r,c,extend=False):
        left,top,right,bottom = self.bounds(r,c)
        if extend and self.selection:
            l,t,rr,b = self.selection
            left,top,right,bottom = min(l,left),min(t,top),max(rr,right),max(b,bottom)
        self.selection = expand_selection(self.info,(left,top,right,bottom))
        self.request_draw()
        if self.on_select:
            self.on_select(self)

    def navigate(self,dr,dc):
        if not self.info:
            return
        c,r,_,_ = self.selection or (1,1,1,1)
        r,c = min(self.info['max_row'],max(1,r+dr)),min(self.info['max_col'],max(1,c+dc))
        self.select(r,c)
        return "break"

    def press(self,event):
        self.canvas.focus_set()
        # 手动识别双击：绑定 <Double-Button-1> 会吞掉第二次按下，破坏随后起拖。
        now = event.time if event.time > 0 else int(time.monotonic()*1000)
        recent,self._click = self._click,(now,event.x,event.y)
        if (recent and self.on_edit and 0 <= now-recent[0] < 450
                and abs(event.x-recent[1]) < 6 and abs(event.y-recent[2]) < 6):
            hit = self.cell_at(event.x,event.y)
            if hit:
                self.gesture = None
                self.on_edit(*hit)
                return
        edge=self.resize_edge(event.x,event.y)
        if edge:
            axis,index=edge
            sizes=self.ys if axis=='row' else self.xs
            self.gesture=('resize',axis,index,event.y if axis=='row' else event.x,(sizes[index]-sizes[index-1])/(self.zoom*self.display_scale))
            return
        if event.x < self.header_width and event.y >= self.header_height and self.info:
            row = bisect.bisect_right(self.ys,self.canvas.canvasy(event.y))
            if 1<=row<=self.info['max_row']:
                start = self.selection[1] if event.state & 1 and self.selection else row
                self.select_rows(start,row)
                self.gesture = ('rows',start,1,event.x_root,event.y_root)
            return
        if event.y < self.header_height and event.x >= self.header_width and self.info:
            column = bisect.bisect_right(self.xs,self.canvas.canvasx(event.x))
            if 1<=column<=self.info['max_col']:
                start = self.selection[0] if event.state & 1 and self.selection else column
                self.select_columns(start,column)
                self.gesture = ('cols',1,start,event.x_root,event.y_root)
            return
        hit = self.cell_at(event.x,event.y)
        if hit is None:
            return
        r,c = hit
        selected = self.selection and self.selection[0]<=c<=self.selection[2] and self.selection[1]<=r<=self.selection[3]
        if event.state & 1:
            self.select(r,c,True)
            self.gesture = None
        elif self.on_drag and selected:
            self.gesture = ('drag',r,c,event.x_root,event.y_root)
        else:
            self.select(r,c)
            self.gesture = ('select',r,c,event.x_root,event.y_root)

    def motion(self,event):
        if not self.gesture:
            return
        if self.gesture[0]=='resize':
            _,axis,index,start,original=self.gesture
            pixels=max(12,original+((event.y if axis=='row' else event.x)-start)/(self.zoom*self.display_scale))
            key='heights' if axis=='row' else 'widths'
            self.info.setdefault(key,{})[index]=pixels
            self._geometry_dirty = True
            self.request_draw()
            return
        mode,r,c,x,y = self.gesture
        if mode=='rows':
            row = min(self.info['max_row'],max(1,bisect.bisect_right(self.ys,self.canvas.canvasy(event.y))))
            self.select_rows(r,row)
        elif mode=='cols':
            column = min(self.info['max_col'],max(1,bisect.bisect_right(self.xs,self.canvas.canvasx(event.x))))
            self.select_columns(c,column)
        elif mode=='select':
            hit = self.cell_at(event.x,event.y)
            if hit:
                rr,cc = hit
                self.selection = expand_selection(self.info,(min(c,cc),min(r,rr),max(c,cc),max(r,rr)))
                self.request_draw()
                if self.on_select:
                    self.on_select(self)
        elif abs(event.x_root-x)+abs(event.y_root-y)>6:
            self.canvas.configure(cursor='hand2')
            self.on_drag(event.x_root,event.y_root)

    def select_rows(self,start,end):
        self.selection = expand_selection(self.info,(1,min(start,end),self.info['max_col'],max(start,end)))
        self.request_draw()
        if self.on_select:self.on_select(self)

    def select_columns(self,start,end):
        # 整列选择不做合并扩展：跨列的标题合并不把其他列带进选区。
        self.selection = (min(start,end),1,max(start,end),self.info['max_row'])
        self.request_draw()
        if self.on_select:self.on_select(self)

    def release(self,event):
        gesture,self.gesture = self.gesture,None
        self.canvas.configure(cursor='')
        if gesture and gesture[0]=='resize':
            _,axis,index,_,original=gesture
            if self._geometry_dirty:self.rebuild_geometry()
            self.on_resize(axis,index,self.info.get('heights' if axis=='row' else 'widths',{}).get(index,original))
            return
        if gesture and gesture[0]=='drag' and abs(event.x_root-gesture[3])+abs(event.y_root-gesture[4])>6:
            if self.on_drop:
                self.on_drop(event.x_root,event.y_root)

    def cancel(self):
        if self.gesture and self.gesture[0]=='resize':
            _,axis,index,_,original=self.gesture
            self.info.setdefault('heights' if axis=='row' else 'widths',{})[index]=original
            self.rebuild_geometry()
        self.gesture = None
        self.canvas.configure(cursor='')
        if self.on_cancel:self.on_cancel()

    def resize_edge(self,x,y):
        if not self.on_resize or not self.info:return None
        if x<self.header_width and y>=self.header_height:
            coordinate=self.canvas.canvasy(y);sizes=self.ys;axis='row'
        elif y<self.header_height and x>=self.header_width:
            coordinate=self.canvas.canvasx(x);sizes=self.xs;axis='column'
        else:return None
        index=bisect.bisect_left(sizes,coordinate)
        for i in (index,index-1):
            if 1<=i<len(sizes) and abs(sizes[i]-coordinate)<=5:return axis,i
        return None

    def resize_cursor(self,event):
        if self.gesture:return
        edge=self.resize_edge(event.x,event.y)
        self.canvas.configure(cursor=('sb_v_double_arrow' if edge[0]=='row' else 'sb_h_double_arrow') if edge else '')

    def scroll_y(self,*args):
        self.canvas.yview(*args);self.request_draw()

    def scroll_x(self,*args):
        self.canvas.xview(*args);self.request_draw()

    def wheel(self,event,horizontal=False):
        (self.canvas.xview_scroll if horizontal else self.canvas.yview_scroll)(-1 if event.delta>0 else 1,'units')
        self.request_draw()
        return 'break'

    def show_value(self,ref):
        return self.updates.get(ref,self.info['cells'].get(ref,''))

    def draw(self):
        if self._geometry_dirty:self.rebuild_geometry()
        if self._frame is not None:
            self.after_cancel(self._frame);self._frame=None
        if not self.info:
            return
        cv = self.canvas
        cv.delete('all')
        left,top = cv.canvasx(0),cv.canvasy(0)
        right,bottom = left+cv.winfo_width(),top+cv.winfo_height()
        first_c,max_c = max(1,bisect.bisect_right(self.xs,left)),min(self.info['max_col'],bisect.bisect_right(self.xs,right))
        first_r,max_r = max(1,bisect.bisect_right(self.ys,top)),min(self.info['max_row'],bisect.bisect_right(self.ys,bottom))
        drawn = set()
        for r in range(first_r,max_r+1):
            for c in range(first_c,max_c+1):
                l,t,rr,b = self.bounds(r,c)
                if (l,t) in drawn:
                    continue
                drawn.add((l,t))
                ref = f'{get_column_letter(l)}{t}'
                x1,y1,x2,y2 = self.xs[l-1],self.ys[t-1],self.xs[rr],self.ys[b]
                style = self.info.get('styles',{}).get(ref,{})
                # 完整落入选区的合并格才算选中；整列选择时只压住一半的标题合并不高亮。
                selected = self.selection and l >= self.selection[0] and rr <= self.selection[2] \
                    and t >= self.selection[1] and b <= self.selection[3]
                bg = '#DCEAF8' if ref in self.updates else style.get('bg') or self.app.surface_bg
                if selected:
                    bg = '#C6D8EB'
                cv.create_rectangle(x1,y1,x2,y2,fill=bg,outline=self.app.border_color)
                for side,edge in style.get('borders',{}).items():
                    coords={'left':(x1,y1,x1,y2),'right':(x2,y1,x2,y2),'top':(x1,y1,x2,y1),'bottom':(x1,y2,x2,y2)}[side]
                    weight=2 if edge['style'] in {'medium','thick','double'} else 1
                    cv.create_line(*coords,fill=edge['color'],width=max(1,weight*self.zoom*self.display_scale))
                text = str(self.show_value(ref))
                if style.get('size'):
                    key = (style.get('family',self.app.font_family),max(6,round(style['size']*4/3*self.zoom*self.display_scale)),bool(style.get('bold')),bool(style.get('italic')),bool(style.get('underline')))
                    if key not in self.cell_fonts:
                        self.cell_fonts[key] = tkfont.Font(family=key[0],size=-key[1],weight='bold' if key[2] else 'normal',slant='italic' if key[3] else 'roman',underline=key[4])
                    font = self.cell_fonts[key]
                else:
                    font = self.text_fonts[bool(style.get('bold'))]
                # 画布文本无单元格裁剪；超长文本在格内折行截断，完整内容在下方显示。
                pad=4*self.zoom*self.display_scale
                width=max(1,x2-x1-2*pad);height=max(1,y2-y1-2*pad)
                visible=self.visible_text(text,width,height,font,style.get('wrap',True))
                align=style.get('horizontal','left')
                if align=='general':
                    try:float(text);align='right'
                    except ValueError:align='left'
                x=(x1+x2)/2 if align=='center' else x2-pad if align=='right' else x1+pad
                vertical=style.get('vertical','top')
                y=(y1+y2)/2 if vertical=='center' else y2-pad if vertical=='bottom' else y1+pad
                anchor=('' if vertical=='center' else 's' if vertical=='bottom' else 'n')+('' if align=='center' else 'e' if align=='right' else 'w')
                cv.create_text(x,y,text=visible,anchor=anchor or 'center',justify=align if align in {'left','right','center'} else 'left',font=font,
                               fill=style.get('fg') or self.app.text_fg)
        if self.selection:
            l,t,r,b = self.selection
            cv.create_rectangle(self.xs[l-1],self.ys[t-1],self.xs[r],self.ys[b],outline=self.app.accent_fg,width=2)
        if self.mark:
            l,t,r,b = self.mark
            cv.create_rectangle(self.xs[l-1],self.ys[t-1],self.xs[r],self.ys[b],outline=self.app.accent_fg,width=3,dash=(5,3))
        # 行列标题固定在视口边缘；整行选择不再点亮全部列字母，整列选择同理。
        whole_width = self.selection and self.selection[0]==1 and self.selection[2]>=self.info['max_col']
        for c in range(first_c,max_c+1):
            x1,x2 = self.xs[c-1],self.xs[c]
            selected = not whole_width and self.selection and self.selection[0]<=c<=self.selection[2]
            cv.create_rectangle(x1,top,x2,top+self.header_height,fill='#C6D8EB' if selected else '#E8EBEF',outline=self.app.border_color)
            cv.create_text((x1+x2)/2,top+self.header_height/2,text=get_column_letter(c),font=self.font,fill=self.app.muted_fg)
        whole_height = self.selection and self.selection[1]==1 and self.selection[3]>=self.info['max_row']
        for r in range(first_r,max_r+1):
            y1,y2 = self.ys[r-1],self.ys[r]
            selected = not whole_height and self.selection and self.selection[1]<=r<=self.selection[3]
            cv.create_rectangle(left,y1,left+self.header_width,y2,fill='#C6D8EB' if selected else '#E8EBEF',outline=self.app.border_color)
            cv.create_text(left+self.header_width/2,(y1+y2)/2,text=str(r),font=self.font,fill=self.app.muted_fg)
        cv.create_rectangle(left,top,left+self.header_width,top+self.header_height,fill='#E8EBEF',outline=self.app.border_color)
