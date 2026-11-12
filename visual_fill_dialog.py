"""左右原表拖放填入；来源快照、右侧预览和最终单元格变更使用同一映射。"""
import hashlib
import json
import os
import tempfile
import tkinter as tk
from pathlib import Path
from copy import deepcopy
from tkinter import ttk
from app_dialogs import messagebox, simpledialog
from tkinter import font as tkfont

from openpyxl.utils.cell import get_column_letter, range_boundaries

from app_settings import load_settings, save_settings
from review_mapping import layout_signature, _anchor
from review_table_filler import review_output_path
from visual_fill import (source_grid, inspect_visual_workbook, drop_values,
                         plan_visual_fill, apply_visual_plan, _value)
from visual_table import VisualTable
from row_height_dialog import ask_row_height
from worksheet_layout import RowMap, InsertRowMap, InsertColumnMap, preview_layout, empty_rows, auto_heights
from ui_common import center_window_on_parent, FlatScrollbar


def source_signature(info):
    return hashlib.sha256(json.dumps([info['max_row'],info['max_col'],info['merges'],
        [(k,v) for k,v in sorted(info['cells'].items()) if range_boundaries(k)[1]==1]],ensure_ascii=False).encode()).hexdigest()


class VisualFillDialog:
    def __init__(self,owner,template,items,sheets=None):
        self.owner,self.app,self.template,self.items = owner,owner.app,template,items
        self.sheets = sheets if sheets is not None else inspect_visual_workbook(template)
        self.digest = hashlib.sha256(Path(template).read_bytes()).hexdigest()
        self.original_digest = self.digest
        stat = Path(template).stat()
        self._original_stamp = (stat.st_size, stat.st_mtime_ns)
        self.original_sheets = self.sheets
        self.working_template = template
        self._scratch = tempfile.TemporaryDirectory(prefix='replace-preview-')
        self._serial = 0
        self._saving = False
        self.applied = {'count':0,'layout_count':0,'structure_count':0,'sheets':[]}
        self.sources = [source_grid(item.table) for item in items]
        self.pending,self.operations,self.history = {},[],[]
        self.layouts = {}
        self.marked = {}
        self._measure_fonts = {}
        self.plan = None
        self.window = tk.Toplevel(self.app.root)
        self.window.title('拖拽填入开评标表')
        self.window.configure(bg=self.app.app_bg)
        self.window.minsize(900,600)
        self.window.resizable(True,True)
        try:
            self.window.iconphoto(True,self.app._icon_photo)
        except (AttributeError,tk.TclError):
            pass
        # 独立普通窗口保留 Windows 最大化按钮；transient 会隐藏标题栏最大化控件。
        body = ttk.Frame(self.window,padding=12,style='App.TFrame')
        body.pack(fill='both',expand=True)
        body.columnconfigure(0,weight=1)
        body.rowconfigure(2,weight=1)
        bar = ttk.Frame(body,style='App.TFrame')
        bar.grid(row=0,column=0,sticky='ew',pady=(0,10))
        ttk.Button(bar,text='撤销',command=self.undo).pack(side='left')
        ttk.Button(bar,text='重做',command=self.redo).pack(side='left',padx=6)
        ttk.Button(bar,text='清空预览',command=self.clear).pack(side='left')
        ttk.Button(bar,text='复用拖放位置…',command=self.restore_recipe).pack(side='left',padx=6)
        ttk.Button(bar,text='记住拖放位置…',command=self.save_recipe).pack(side='left')
        ttk.Button(bar,text='自动匹配…',command=self.auto_match).pack(side='right')
        self.maximize_button = ttk.Button(bar,text='最大化',command=self.toggle_maximize)
        self.maximize_button.pack(side='right',padx=6)
        ttk.Label(body,text='框选左表，再从选区拖到右表。蓝色内容为待填入；双击右表单元格可修改内容。',style='Subtitle.TLabel').grid(row=1,column=0,sticky='w',pady=(0,8))
        panes = ttk.Panedwindow(body,orient='horizontal')
        panes.grid(row=2,column=0,sticky='nsew')
        left = ttk.Frame(panes,padding=10,style='Surface.TFrame')
        right = ttk.Frame(panes,padding=10,style='Surface.TFrame')
        panes.add(left,weight=1);panes.add(right,weight=1)
        for pane in (left,right):
            pane.columnconfigure(0,weight=1)
            pane.rowconfigure(2,weight=1)
        ttk.Label(left,text='招标文件提取的表格',style='SectionTitle.TLabel').grid(row=0,column=0,sticky='w',pady=(0,6))
        choices = [f'{item.filename} · 表 {item.table_index}' for item in items]
        self.source = ttk.Combobox(left,state='readonly',values=choices)
        self.source.grid(row=1,column=0,sticky='ew',pady=(0,8))
        heading = ttk.Frame(right,style='Toolbar.TFrame')
        heading.grid(row=0,column=0,sticky='ew',pady=(0,6))
        ttk.Label(heading,text=os.path.basename(template),style='SectionTitle.TLabel').pack(side='left')
        ttk.Button(heading,text='更换文件…',style='Header.TButton',command=self.change_file).pack(side='right')
        self.sheet = tk.StringVar()
        self.destination = ttk.Combobox(right,textvariable=self.sheet,state='readonly',values=list(self.sheets))
        self.destination.grid(row=1,column=0,sticky='ew',pady=(0,8))
        self.left = VisualTable(left,self.app,on_select=self.selected,on_drag=self.dragging,on_drop=self.dropped,on_cancel=self.cancel_drag)
        # 右表用原生 Frame 承载，便于 wm manage 整体弹出为独立窗口（ttk 框架不可管理）。
        self.right_host = tk.Frame(right,bg=self.app.surface_bg)
        self.right_host.grid(row=2,column=0,sticky='nsew')
        self.right_host.rowconfigure(0,weight=1);self.right_host.columnconfigure(0,weight=1)
        self.right = VisualTable(self.right_host,self.app,on_select=self.selected,on_resize=self.resize_dimension,on_edit=self.edit_cell)
        self.left.grid(row=2,column=0,sticky='nsew');self.right.grid(row=0,column=0,sticky='nsew')
        self.left.add_zoom_controls(left).grid(row=3,column=0,sticky='e',pady=(4,0))
        editing = ttk.Frame(right,style='Toolbar.TFrame')
        editing.grid(row=3,column=0,sticky='ew',pady=(4,0))
        self.fit_button = ttk.Button(editing,text='自动行高',style='Header.TButton',command=lambda:self.adjust_rows('fit'))
        self.fit_button.pack(side='left')
        self.popout_button = ttk.Button(editing,text='放大表格',style='Header.TButton',command=self.toggle_table_popout)
        self.popout_button.pack(side='left',padx=(4,0))
        self.edit_button = ttk.Menubutton(editing,text='编辑表格',style='Header.TMenubutton')
        self.edit_button.pack(side='left',padx=(4,0))
        self.edit_menu = tk.Menu(self.edit_button,tearoff=False)
        self.edit_button.configure(menu=self.edit_menu)
        self.edit_menu.add_command(label='编辑单元格…',command=self.edit_selected_cell)
        self.edit_menu.add_command(label='清空所选内容',command=self.clear_selected_cells)
        self.edit_menu.add_command(label='设置行高…',command=lambda:self.adjust_rows('height'))
        self.edit_menu.add_command(label='删除所选行',command=lambda:self.edit_selection('delete_rows'))
        self.edit_menu.add_separator()
        for label,action in [('合并单元格','merge'),('拆分单元格','unmerge'),
                             ('上方插入行…','insert_above'),('下方插入行…','insert_below'),
                             ('左侧插入列…','insert_left'),('右侧插入列…','insert_right')]:
            self.edit_menu.add_command(label=label,command=lambda a=action:self.edit_selection(a))
        self.edit_menu.add_command(label='删除空行',command=lambda:self.adjust_rows('delete'))
        self.right.canvas.bind('<Button-3>',self.open_edit_menu)
        self.window.bind('<F2>',lambda _:self.edit_selected_cell())
        self.right.bind('<F2>',lambda _:self.edit_selected_cell())  # 表格弹出为独立窗口时仍可用
        self.right.add_zoom_controls(editing).pack(side='right')
        details = ttk.Frame(body,style='Surface.TFrame')
        details.grid(row=3,column=0,sticky='ew',pady=(8,0))
        self.detail = tk.Text(details,height=3,wrap='word',font=self.app.body_font,bg=self.app.surface_bg,
                              fg=self.app.text_fg,relief='flat',padx=10,pady=6,state='disabled')
        self.detail.pack(side='left',fill='both',expand=True)
        scroll = FlatScrollbar(details,orient='vertical',command=self.detail.yview,style='Flat.Vertical.TScrollbar')
        scroll.pack(side='right',fill='y');self.detail.configure(yscrollcommand=scroll.set)
        footer = ttk.Frame(body,style='App.TFrame')
        footer.grid(row=4,column=0,sticky='ew',pady=(10,0));footer.columnconfigure(0,weight=1)
        self.status = tk.StringVar(value='在左侧选择要填写的内容。')
        ttk.Label(footer,textvariable=self.status,style='Subtitle.TLabel',wraplength=620).grid(row=0,column=0,sticky='w')
        self.generate = ttk.Button(footer,text='填入并另存',style='Accent.TButton',command=self.write,state='disabled')
        self.generate.grid(row=0,column=1,sticky='e')
        self.save_progress = ttk.Progressbar(footer,mode='indeterminate',maximum=100)
        self.save_progress.grid(row=1,column=0,columnspan=2,sticky='ew',pady=(6,0))
        self.save_progress.grid_remove()
        self.source.bind('<<ComboboxSelected>>',self.source_changed)
        self.destination.bind('<<ComboboxSelected>>',self.target_changed)
        self.window.bind('<Control-z>',lambda _:self.undo())
        self.window.bind('<Control-y>',lambda _:self.redo())
        self.window.bind('<F11>',self.toggle_fullscreen)
        self.window.bind('<Escape>',self.leave_fullscreen,add='+')
        self.window.bind('<Configure>',self.update_window_state,add='+')
        self.window.bind('<Destroy>',self.release_preview,add='+')
        self.source.current(0)
        self.source_changed()
        preferred = next((name for name in self.sheets if '评分' in name or '评审' in name),next(iter(self.sheets)))
        self.sheet.set(preferred);self.target_changed()
        self.future = []
        center_window_on_parent(self.window,self.app.root,1200,760)
        self.window.grab_set()

    def toggle_maximize(self):
        self.window.attributes('-fullscreen',False)
        self.window.state('normal' if self.window.state()=='zoomed' else 'zoomed')
        self.update_window_state()

    def toggle_fullscreen(self,_event=None):
        self.window.attributes('-fullscreen',not bool(self.window.attributes('-fullscreen')))
        return 'break'

    def leave_fullscreen(self,_event=None):
        self.window.attributes('-fullscreen',False)

    def update_window_state(self,event=None):
        if event is None or event.widget is self.window:
            self.maximize_button.configure(text='还原' if self.window.state()=='zoomed' else '最大化')

    def source_changed(self,*_):
        self.left.load(self.sources[self.source.current()])

    def target_changed(self,*_):
        name = self.sheet.get()
        info, updates = preview_layout(self.sheets[name],self.pending.get(name,{}),self.layouts.get(name,{}),name,self.layouts)
        mapping = RowMap(self.layouts.get(name,{}).get('delete_rows',[]))
        painted = {moved:info['cells'].get(moved,'') for ref in self.marked.get(name,[]) if (moved:=mapping.ref(ref))}
        updates = painted | updates
        # 留出空白接收区域；这只是画布，不向文件中插入行。
        occupied = [range_boundaries(ref) for ref in updates]
        info['max_row'] = min(1048576,max(30,info['max_row']+10,max((b[1]+5 for b in occupied),default=1)))
        info['max_col'] = min(16384,max(10,info['max_col'],max((b[0] for b in occupied),default=1)))
        self.right.load(info,updates)

    def selected(self,grid):
        if not grid.selection:
            return
        key = (id(grid),grid.selection,id(grid.info),id(grid.updates))
        if getattr(self,'_detail_key',None)==key:return
        self._detail_key = key
        l,t,r,b = grid.selection
        lines = []
        chars = 0
        for row in range(t,b+1):
            for col in range(l,r+1):
                ref = f'{get_column_letter(col)}{row}'
                value = grid.show_value(ref)
                if value != '':
                    lines.append(f'{ref}：{value}')
                    chars += len(str(value))
                if len(lines)>=100 or chars>=100000:break
            if len(lines)>=100 or chars>=100000:
                lines.append('选区较大，仅显示部分内容；拖放仍包含全部所选单元格。')
                break
        self.detail.configure(state='normal');self.detail.delete('1.0','end')
        self.detail.insert('1.0','\n\n'.join(lines));self.detail.configure(state='disabled')
        if grid is self.left:
            self.status.set('已选来源区域，按住选区拖到右表；Shift + 点击可扩大选区。')
        elif t==1 and b==grid.info['max_row']:
            self.status.set(f'已选第 {get_column_letter(l)}–{get_column_letter(r)} 列，可插入或删除列。')
        else:
            self.status.set(f'已选第 {t}–{b} 行，可调整行高或删除其中的空行。')

    def dragging(self,x,y):
        hit = self.right.cell_at_screen(x,y)
        if hit:
            r,c = hit
            # 靠近右表边缘时滚动，便于拖到当前视口之外。
            local_y = y-self.right.canvas.winfo_rooty()
            if local_y>self.right.canvas.winfo_height()-24:
                self.right.scroll_y('scroll',1,'units')
            elif local_y<55:
                self.right.scroll_y('scroll',-1,'units')
            hit = self.right.cell_at_screen(x,y)
            if hit:
                r,c = hit
                self.right.mark = self.right.bounds(r,c)
        else:
            self.right.mark = None
        self.right.request_draw()

    def dropped(self,x,y):
        hit = self.right.cell_at_screen(x,y)
        self.right.mark = None;self.right.draw()
        if hit and self.left.selection:
            try:
                self.stage_drop(self.source.current(),self.left.selection,self.sheet.get(),*hit)
            except (ValueError,OSError) as exc:
                self.status.set(str(exc))
                messagebox.showwarning('无法拖入',str(exc),parent=self.window)

    def cancel_drag(self):
        self.right.mark = None
        self.right.draw()

    def checkpoint(self):
        self.history.append(self.snapshot())
        del self.history[:-100]
        self.future.clear()

    def snapshot(self):
        state=deepcopy({key:getattr(self,key) for key in ('pending','operations','layouts','working_template','digest','applied','marked')})
        state['sheets']=self.sheets  # 工作表快照只读，版本之间共享，不重复复制整本模板。
        return state

    def restore_snapshot(self,snapshot):
        for key,value in snapshot.items():setattr(self,key,value if key=='sheets' else deepcopy(value))

    def release_preview(self,event):
        if event.widget is self.window and not self._saving:
            self._scratch.cleanup()

    def verify_original(self, fast=False):
        if fast and hasattr(self, '_original_stamp'):
            stat = Path(self.template).stat()
            if (stat.st_size, stat.st_mtime_ns) == self._original_stamp:
                return
        if hashlib.sha256(Path(self.template).read_bytes()).hexdigest()!=self.original_digest:
            raise ValueError('原模板文件已变化，请重新打开填表窗口。')

    def preview_plan(self, pending):
        """Plain text edits validate against the loaded snapshot; saving revalidates bytes."""
        if self.layouts:
            return plan_visual_fill(self.working_template, pending, self.layouts)
        from workbook_session import WorkbookSession
        session = WorkbookSession(self.working_template, self.sheets, self.digest)
        changes, errors = {}, []
        for name, values in pending.items():
            try:
                session.edit(name, values)
                edits = [{'cell': ref, 'old': self.sheets[name]['cells'].get(ref, ''), 'new': value}
                         for ref, value in session.pending.get(name, {}).items()
                         if str(self.sheets[name]['cells'].get(ref, '')) != str(value)]
                if edits:
                    changes[name] = edits
            except ValueError as exc:
                errors.append(str(exc))
        return {'digest': self.digest, 'sheets': changes, 'errors': errors,
                'count': sum(map(len, changes.values())), 'layouts': {}, 'layout_count': 0}

    def preview_path(self):
        self._serial += 1
        return str(Path(self._scratch.name)/f'preview-{self._serial}{Path(self.template).suffix}')

    def commit(self,pending,operation):
        if self._saving:raise ValueError('正在保存，请稍候。')
        self.verify_original(fast=True)
        plan = self.preview_plan(pending)
        if plan['digest']!=self.digest:
            raise ValueError('模板文件已变化，请重新打开填表窗口。')
        if plan['errors']:
            raise ValueError('；'.join(plan['errors']))
        self.checkpoint()
        self.pending = pending
        self.operations.append(operation)
        self.refresh(plan)

    def stage_drop(self,index,bounds,sheet,row,column):
        values = self.mapped_drop(index,bounds,sheet,row,column,self.layouts)
        if not values:
            raise ValueError('所选区域没有可填写的内容。')
        pending = deepcopy(self.pending)
        pending.setdefault(sheet,{}).update(values)
        original_row = RowMap(self.layouts.get(sheet,{}).get('delete_rows',[])).original(row)
        self.commit(pending,{'kind':'drag','source':index,'bounds':list(bounds),'sheet':sheet,'row':original_row,'column':column})

    def mapped_drop(self,index,bounds,sheet,row,column,layouts):
        layout = layouts.get(sheet,{})
        info,_ = preview_layout(self.sheets[sheet],{},layout)
        values = drop_values(self.sources[index],bounds,info,row,column)
        mapping = RowMap(layout.get('delete_rows',[]))
        return {f'{get_column_letter(range_boundaries(ref)[0])}{mapping.original(range_boundaries(ref)[1])}':v
                for ref,v in values.items()}

    def measure_text(self,text,style):
        key = (style.get('family',self.app.font_family),style.get('size',11),bool(style.get('bold')))
        if not hasattr(self,'_measure_cache'):self._measure_cache={}
        cache_key=(key,text)
        if cache_key in self._measure_cache:return self._measure_cache[cache_key]
        if key not in self._measure_fonts:
            # 负字号直接指定 96dpi 像素，与 Excel 列宽换算共用坐标系。
            self._measure_fonts[key] = tkfont.Font(root=self.window,family=key[0],size=-round(key[1]*4/3),weight='bold' if key[2] else 'normal')
        width=self._measure_fonts[key].measure(text)
        if len(self._measure_cache)>=4096:self._measure_cache.clear()
        self._measure_cache[cache_key]=width
        return width

    def build_layout(self,action,name,rows,pending,layouts,height=None):
        result = deepcopy(layouts)
        layout = result.setdefault(name,{'heights':{},'delete_rows':[],'wrap_cells':[]})
        if action=='fit':
            info=deepcopy(self.sheets[name])
            for col,width in layout.get('widths',{}).items():info.setdefault('column_pixels',{})[col]=width*7+5
            heights,wraps = auto_heights(info,pending.get(name,{}),rows,layout['delete_rows'],self.measure_text)
            layout['heights'].update(heights)
            layout['wrap_cells'] = sorted(set(layout['wrap_cells']) | set(wraps))
        elif action=='height':
            layout['heights'].update({row:height for row in rows if row not in layout['delete_rows']})
        elif action=='width':
            layout.setdefault('widths',{}).update({col:height for col in rows})
        elif action=='delete':
            deleted = empty_rows(self.sheets[name],pending.get(name,{}),rows)
            layout['delete_rows'] = sorted(set(layout['delete_rows']) | set(deleted))
            for row in deleted:
                layout['heights'].pop(row,None)
        else:
            raise ValueError('未知的排版操作。')
        return result

    def resize_dimension(self,axis,index,pixels):
        if self._saving:return
        name=self.sheet.get()
        index=RowMap(self.layouts.get(name,{}).get('delete_rows',[])).original(index) if axis=='row' else index
        value=round(max(1,min(409,pixels*3/4)),2) if axis=='row' else round(max(1,min(255,(pixels-5)/7)),3)
        action='height' if axis=='row' else 'width'
        try:
            self.verify_original()
            layouts=self.build_layout(action,name,[index],self.pending,self.layouts,value)
            plan=plan_visual_fill(self.working_template,self.pending,layouts)
            self.checkpoint();self.layouts=layouts
            self.operations.append({'kind':'layout','action':action,'sheet':name,'rows':[index],'height':value})
            self.refresh(plan)
        except (ValueError,OSError) as exc:
            self.target_changed();messagebox.showwarning('无法调整',str(exc),parent=self.window)

    def adjust_rows(self,action):
        if self._saving:return
        name = self.sheet.get()
        mapping = RowMap(self.layouts.get(name,{}).get('delete_rows',[]))
        if self.right.selection:
            _,top,_,bottom = self.right.selection
            rows = [mapping.original(r) for r in range(top,bottom+1)]
        elif action=='fit':
            rows = sorted({range_boundaries(ref)[1] for ref in self.pending.get(name,{})})
        else:
            rows = []
        if not rows:
            self.status.set('请先在右表选择要调整的行（可拖动行号多选）。')
            return
        height = None
        if action=='height':
            current = self.layouts.get(name,{}).get('heights',{}).get(rows[0])
            if current is None:
                current = self.sheets[name].get('heights',{}).get(rows[0],self.sheets[name].get('default_height',40))*3/4
            height = ask_row_height(self.window,self.app,current,mapping.display(rows[0]),mapping.display(rows[-1]))
            if height is None:return
        try:
            layouts = self.build_layout(action,name,rows,self.pending,self.layouts,height)
            deleted = len(layouts[name]['delete_rows'])-len(self.layouts.get(name,{}).get('delete_rows',[]))
            if action=='delete' and not deleted:
                self.status.set('所选行中没有可删除的整行空白；有内容、公式或跨行合并的行会保留。')
                return
            self.verify_original()
            plan = plan_visual_fill(self.working_template,self.pending,layouts)
            if plan['digest']!=self.digest:
                raise ValueError('模板文件已变化，请重新打开填表窗口。')
            self.checkpoint();self.layouts = layouts
            self.operations.append({'kind':'layout','action':action,'sheet':name,'rows':rows,'height':height})
            self.refresh(plan)
            self.status.set(f'已删除 {deleted} 个空行，可撤销。' if action=='delete' else '行高已更新，可撤销。')
        except (ValueError,OSError) as exc:
            self.status.set(str(exc))
            messagebox.showwarning('无法调整',str(exc),parent=self.window)

    def refresh(self,plan=None):
        self.plan = plan or self.preview_plan(self.pending)
        selection = self.right.selection
        self.target_changed()
        if selection and selection[2]<=self.right.info['max_col'] and selection[3]<=self.right.info['max_row']:
            self.right.selection = selection
        self.right.draw()
        if self.plan['digest']!=self.digest:
            self.generate.configure(state='disabled')
            self.status.set('模板文件已变化，请重新打开填表窗口。')
            return
        self.generate.configure(state='normal' if (self.plan['count'] or self.plan['layout_count'] or self.applied['structure_count']) and not self.plan['errors'] else 'disabled')
        self.status.set(f"待保存：{self.plan['count']+self.applied['count']} 个单元格，{self.plan['layout_count']+self.applied['layout_count']} 处行调整，{self.applied['structure_count']} 次结构编辑。")

    def undo(self):
        if self.history and not self._saving:
            self.future.append(self.snapshot())
            self.restore_snapshot(self.history.pop())
            self.refresh()

    def redo(self):
        if self.future and not self._saving:
            self.history.append(self.snapshot())
            self.restore_snapshot(self.future.pop())
            self.refresh()

    def clear(self):
        if self._saving:return
        if self.pending or self.layouts or self.applied['structure_count']:
            self.checkpoint();self.reset_preview();self.refresh()

    def reset_preview(self):
        self.pending,self.operations,self.layouts = {},[],{}
        self.marked = {}
        self.working_template,self.digest = self.template,self.original_digest
        self.sheets = self.original_sheets
        self.applied = {'count':0,'layout_count':0,'structure_count':0,'sheets':[]}

    def open_edit_menu(self,event):
        hit = self.right.cell_at(event.x,event.y)
        if hit:
            r,c = hit
            bounds = self.right.selection
            if not bounds or not bounds[0]<=c<=bounds[2] or not bounds[1]<=r<=bounds[3]:
                self.right.select(r,c)
        try:self.edit_menu.tk_popup(event.x_root,event.y_root)
        finally:self.edit_menu.grab_release()

    def edit_selection(self,action):
        if self._saving:return
        bounds = self.right.selection
        if not bounds:
            self.status.set('请先在右表选择单元格或行。');return
        count = 1
        if action.startswith('insert_'):
            unit='列' if action in {'insert_left','insert_right'} else '行'
            count = simpledialog.askinteger(f'插入{unit}',f'插入{unit}数',initialvalue=1,minvalue=1,maxvalue=1000,parent=self.window)
            if count is None:return
        if self.owner._task_runner.active or self.owner.app._task_runner.active:
            return
        sheet = self.sheet.get()
        self._saving = True
        self.generate.configure(state='disabled')
        self.owner._begin_indeterminate_task('正在后台调整工作簿结构…')
        self.status.set('正在后台调整工作簿结构…')
        def done(result):
            self._saving = False
            self.owner._reset_busy_state()
            if self.window.winfo_exists():
                self._accept_structure(result)
            else:
                self._scratch.cleanup()
        def failed(exc):
            self._saving = False
            self.owner._reset_busy_state()
            if self.window.winfo_exists():
                self.refresh()
                self.status.set(str(exc))
                messagebox.showwarning('无法编辑',str(exc),parent=self.window)
            else:
                self._scratch.cleanup()
        self.owner._task_runner.submit(lambda publish: self._prepare_structure(action,sheet,bounds,count),done,failed)

    def toggle_table_popout(self):
        """把右侧表格整体弹出为独立窗口编辑；还原时嵌回原位，内容与操作不变。"""
        host = self.right_host
        if getattr(self, '_table_popped', False):
            self.window.tk.call('wm', 'forget', host._w)
            host.grid(row=2, column=0, sticky='nsew')
            self._table_popped = False
            self.popout_button.configure(text='放大表格')
            self.window.grab_set()
        else:
            host.grid_forget()
            self.window.tk.call('wm', 'manage', host._w)
            screen_w, screen_h = host.winfo_screenwidth(), host.winfo_screenheight()
            width, height = int(screen_w * .8), int(screen_h * .8)
            wm = host.tk
            wm.call('wm', 'title', host._w, '表格放大编辑')
            wm.call('wm', 'geometry', host._w, f'{width}x{height}+{(screen_w - width) // 2}+{max(0,(screen_h - height) // 2)}')
            wm.call('wm', 'minsize', host._w, 560, 360)
            wm.call('wm', 'resizable', host._w, 1, 1)
            wm.call('wm', 'protocol', host._w, 'WM_DELETE_WINDOW', self.window.register(self.toggle_table_popout))
            self._table_popped = True
            self.popout_button.configure(text='还原表格')
        self.right.request_draw()

    def edit_cell(self,row,column):
        self.right.select(row,column)
        self.edit_selected_cell()

    def edit_selected_cell(self):
        if self._saving:return
        if not self.right.selection:
            self.status.set('请先在右表选择要编辑的单元格，或直接双击单元格。');return
        left,top,_,_ = self.right.selection
        from cell_edit_dialog import ask_cell_text
        name = self.sheet.get()
        mapping = RowMap(self.layouts.get(name,{}).get('delete_rows',[]))
        column,row = self.right.bounds(top,left)[:2]
        display = f'{get_column_letter(column)}{row}'
        ref = f'{get_column_letter(column)}{mapping.original(row)}'
        text = ask_cell_text(self.window,self.app,display,str(self.right.show_value(display)))
        if text is None:return
        try:
            self.stage_edit(name,column,row,text)
            self.status.set(f'已修改 {name}!{ref}，可撤销。')
        except (ValueError,OSError) as exc:
            self.status.set(str(exc));messagebox.showwarning('无法编辑',str(exc),parent=self.window)

    def stage_edit(self,sheet,column,row,text):
        """把右表显示坐标 (column,row) 的单元格内容改为 text；合并格落到左上锚点。"""
        if self._saving:raise ValueError('正在保存，请稍候。')
        column,row = self.right.bounds(row,column)[:2]
        original = RowMap(self.layouts.get(sheet,{}).get('delete_rows',[])).original(row)
        if original is None:
            raise ValueError('该行已删除，请重新选择。')
        ref = f'{get_column_letter(column)}{original}'
        pending = deepcopy(self.pending)
        pending.setdefault(sheet,{})[ref] = _value(text)
        self.commit(pending,{'kind':'edit','sheet':sheet,'cell':ref,'value':text})

    def clear_selected_cells(self):
        if self._saving:return
        if not self.right.selection:
            self.status.set('请先在右表选择要清空的单元格。');return
        try:
            skipped = self.stage_clear(self.sheet.get(),self.right.selection)
            note = f'（跳过 {skipped} 个公式格）' if skipped else ''
            self.status.set(f'已清空所选内容，可撤销。{note}')
        except (ValueError,OSError) as exc:
            self.status.set(str(exc));messagebox.showwarning('无法清空',str(exc),parent=self.window)

    def stage_clear(self,sheet,bounds):
        """清空选区（右表显示坐标）内有内容的普通单元格；公式格跳过并返回跳过数。"""
        if self._saving:raise ValueError('正在保存，请稍后。')
        info = self.sheets[sheet]
        mapping = RowMap(self.layouts.get(sheet,{}).get('delete_rows',[]))
        formula_ranges = [range_boundaries(ref) for ref in info.get('formula_ranges',[])]
        left,top,right,bottom = bounds
        targets = {}
        for row in range(top,bottom+1):
            original = mapping.original(row)
            if original is None:continue
            for column in range(left,right+1):
                ref = _anchor(f'{get_column_letter(column)}{original}',info['merges'])
                if ref in targets:
                    continue
                if ref in info.get('formulas',{}) or any(
                        l<=range_boundaries(ref)[0]<=r and t<=range_boundaries(ref)[1]<=b for l,t,r,b in formula_ranges):
                    targets[ref] = None
                    continue
                if str(info['cells'].get(ref,'') or '')!='' or ref in self.pending.get(sheet,{}):
                    targets[ref] = ''
        skipped = sum(1 for value in targets.values() if value is None)
        values = {ref:value for ref,value in targets.items() if value is not None}
        if not values:
            raise ValueError('所选区域没有可清空的内容（公式格不会被覆盖）。')
        pending = deepcopy(self.pending)
        pending.setdefault(sheet,{}).update(values)
        self.commit(pending,{'kind':'clear','sheet':sheet,'cells':sorted(values)})
        return skipped

    def stage_structure(self,action,sheet,bounds,count=1):
        """Synchronous transaction API for scripts/tests; GUI dispatches preparation to a worker."""
        if self._saving:raise ValueError('正在保存，请稍候。')
        self._accept_structure(self._prepare_structure(action,sheet,bounds,count))

    def _prepare_structure(self,action,sheet,bounds,count=1):
        from workbook_structure import edit_structure
        self.verify_original()
        plan = plan_visual_fill(self.working_template,self.pending,self.layouts)
        if plan['errors']:raise ValueError('；'.join(plan['errors']))
        if plan['digest']!=self.digest:raise ValueError('预览文件已变化，请重新打开。')
        current = self.working_template
        if plan['count'] or plan['layout_count']:
            current = self.preview_path()
            apply_visual_plan(self.working_template,current,plan)
        output = self.preview_path()
        edit_structure(current,output,sheet,action,bounds,count)
        sheets = inspect_visual_workbook(output)
        marked = {}
        for name in self.sheets:
            mapping = RowMap(self.layouts.get(name,{}).get('delete_rows',[]))
            refs = set(self.marked.get(name,[])) | set(self.pending.get(name,{}))
            marked[name] = {moved for ref in refs if (moved:=mapping.ref(ref))}
        left,top,right,bottom = bounds
        if action in {'insert_above','insert_below','insert_left','insert_right','delete_rows'}:
            if action in {'insert_left','insert_right'}:mapping=InsertColumnMap(left if action=='insert_left' else right+1,count)
            else:mapping = RowMap(range(top,bottom+1)) if action=='delete_rows' else InsertRowMap(top if action=='insert_above' else bottom+1,count)
            marked[sheet] = {moved for ref in marked[sheet] if (moved:=mapping.ref(ref))}
        elif action=='merge':
            marked[sheet] = {ref for ref in marked[sheet] if not (left<=range_boundaries(ref)[0]<=right and top<=range_boundaries(ref)[1]<=bottom)}
            marked[sheet].add(f'{get_column_letter(left)}{top}')
        return {'output': output, 'sheets': sheets, 'marked': marked, 'plan': plan,
                'digest': hashlib.sha256(Path(output).read_bytes()).hexdigest(),
                'operation': {'kind':'structure','action':action,'sheet':sheet,'bounds':list(bounds),'count':count}}

    def _accept_structure(self, result):
        plan = result['plan']
        sheet = result['operation']['sheet']
        self.checkpoint()
        self.working_template,self.sheets = result['output'],result['sheets']
        self.digest = result['digest']
        self.applied['count'] += plan['count'];self.applied['layout_count'] += plan['layout_count']
        self.applied['structure_count'] += 1
        self.applied['sheets'] = list(dict.fromkeys([*self.applied['sheets'],*plan['sheets'],*plan['layouts'],sheet]))
        self.pending,self.layouts = {},{}
        self.marked = result['marked']
        self.operations.append(result['operation'])
        self.refresh()

    def auto_match(self):
        if self._saving:return
        from review_mapping_dialog import ReviewMappingDialog
        self._auto_dialog = ReviewMappingDialog(self.owner,self.working_template,self.items,self.sheets,preview_receiver=self.receive_auto)

    def receive_auto(self,plan,item,configs):
        try:
            if plan['errors']:
                raise ValueError('请先处理自动匹配中的未匹配项目。')
            pending = deepcopy(self.pending)
            pending.setdefault(plan['sheet'],{}).update({c['cell']:c['new'] for c in plan['changes']})
            self.commit(pending,{'kind':'auto','source':self.items.index(item),'configs':configs})
            self.window.after_idle(self.window.grab_set)
            return True
        except ValueError as exc:
            messagebox.showwarning('无法应用预览',str(exc),parent=self.window)
            return False

    def save_recipe(self):
        if not self.operations:
            return
        name = simpledialog.askstring('记住拖放位置','给这套填表方式起个名字：',parent=self.window)
        if not name:
            return
        recipes = load_settings().get('review_drag_profiles',{})
        if not isinstance(recipes,dict):
            recipes = {}
        recipes[name] = {'operations':deepcopy(self.operations),
                         'sources':[source_signature(s) for s in self.sources],
                         'targets':{n:layout_signature(v,v['guess']) for n,v in self.original_sheets.items()}}
        save_settings({'review_drag_profiles':recipes})
        self.status.set('已记住拖放位置，下次可复用并预览。')

    def restore_recipe(self):
        recipes = load_settings().get('review_drag_profiles',{})
        if not isinstance(recipes,dict) or not recipes:
            messagebox.showinfo('复用拖放位置','还没有保存过拖放位置。',parent=self.window);return
        popup = tk.Toplevel(self.window);popup.title('复用拖放位置');popup.configure(bg=self.app.app_bg)
        frame = ttk.Frame(popup,padding=14,style='Surface.TFrame');frame.pack(fill='both',expand=True)
        choice = ttk.Combobox(frame,state='readonly',values=list(recipes),width=36);choice.pack(fill='x');choice.current(0)
        def apply():
            try:
                self.replay_recipe(recipes[choice.get()])
                popup.destroy();self.window.grab_set()
            except (ValueError,KeyError,TypeError) as exc:
                messagebox.showwarning('无法复用',str(exc),parent=popup)
        ttk.Button(frame,text='应用到预览',style='Accent.TButton',command=apply).pack(pady=(12,0),anchor='e')
        popup.protocol('WM_DELETE_WINDOW',lambda:(popup.destroy(),self.window.grab_set()))
        center_window_on_parent(popup,self.window);popup.grab_set()

    def replay_recipe(self,recipe):
        if self._saving:raise ValueError('正在保存，请稍候。')
        self.verify_original()
        # 从原模板重放整套操作，失败时恢复完整状态；成功只占一次撤销。
        before = self.snapshot()
        history,future = list(self.history),list(self.future)
        try:
            for op in recipe['operations']:
                name = op['configs'][1]['sheet'] if op['kind']=='auto' else op['sheet']
                if name not in self.original_sheets or recipe['targets'][name]!=layout_signature(self.original_sheets[name],self.original_sheets[name]['guess']):
                    raise ValueError('目标工作表的表头或合并结构已变化，请重新拖放。')
            self.reset_preview()
            for saved in recipe['operations']:
                op = deepcopy(saved)
                if 'source' in op:
                    signature = recipe['sources'][op['source']]
                    matches = [i for i,s in enumerate(self.sources) if source_signature(s)==signature]
                    if len(matches)!=1:
                        raise ValueError('来源表结构已变化或存在多个同结构表，请重新拖放。')
                    op['source'] = matches[0]
                kind = op['kind']
                if kind=='structure':
                    self.stage_structure(op['action'],op['sheet'],op['bounds'],op.get('count',1))
                elif kind=='layout':
                    self.layouts = self.build_layout(op['action'],op['sheet'],op['rows'],self.pending,self.layouts,op.get('height'))
                    plan_visual_fill(self.working_template,self.pending,self.layouts)
                    self.operations.append(op)
                elif kind=='drag':
                    row = RowMap(self.layouts.get(op['sheet'],{}).get('delete_rows',[])).display(op['row'])
                    if row is None:raise ValueError('保存的拖入行已被删除，请重新拖放。')
                    self.stage_drop(op['source'],op['bounds'],op['sheet'],row,op['column'])
                elif kind=='edit':
                    cell = op['cell']
                    row = RowMap(self.layouts.get(op['sheet'],{}).get('delete_rows',[])).display(range_boundaries(cell)[1])
                    if row is None:raise ValueError('保存的编辑行已被删除，请重新编辑。')
                    self.stage_edit(op['sheet'],range_boundaries(cell)[0],row,op['value'])
                elif kind=='clear':
                    pending = deepcopy(self.pending)
                    pending.setdefault(op['sheet'],{}).update({ref:'' for ref in op['cells']})
                    self.commit(pending,op)
                elif kind=='auto':
                    from review_mapping import plan_fill
                    from review_table_filler import extract_review_rules
                    source,target = op['configs']
                    rules = extract_review_rules(self.items[op['source']].table,source,target.get('include_price',False))
                    auto = plan_fill(self.working_template,rules,target)
                    if auto['errors']:raise ValueError('自动匹配设置需要重新核对，请打开自动匹配。')
                    pending = deepcopy(self.pending)
                    pending.setdefault(target['sheet'],{}).update({c['cell']:c['new'] for c in auto['changes']})
                    self.commit(pending,op)
                else:
                    raise ValueError('保存的设置包含未知操作，请重新设置。')
            self.history,self.future = history+[before],[]
            self.refresh()
        except Exception:
            self.restore_snapshot(before);self.history,self.future = history,future
            self.refresh()
            raise

    def change_file(self):
        if self._saving:return
        self.window.destroy();self.owner.start_review_fill(choose_file=True)

    def write(self):
        if not self.plan or self.plan['errors'] or not (self.plan['count'] or self.plan['layout_count'] or self.applied['structure_count']) or self.plan['digest']!=self.digest or self.owner._task_runner.active or self.owner.app._task_runner.active:
            return
        from word_scan import validate_scan_files
        plan = deepcopy(self.plan)
        paths = list({self.items[op['source']].file_path for op in self.operations if 'source' in op})
        stamps = dict(self.owner.scan_stamps)
        output = review_output_path(self.template,self.owner.output_dir)
        working = self.working_template
        applied = deepcopy(self.applied)
        self.owner._begin_indeterminate_task('正在核验并准备保存…')
        self._saving = True
        self.generate.configure(state='disabled')
        self.save_progress.configure(mode='indeterminate',maximum=100,value=0)
        self.save_progress.grid();self.save_progress.start(30)
        self.status.set('正在核验并准备保存…')
        def progress(current,total,status):
            self.owner._update_progress(current,total,os.path.basename(output),status)
            if self.window.winfo_exists():
                self.save_progress.stop()
                self.save_progress.configure(mode='determinate',maximum=total,value=current)
                self.status.set(f'{status}（{current}/{total}）')
        def work(publish):
            self.verify_original()
            validate_scan_files(paths,stamps)
            if plan['count'] or plan['layout_count']:
                result = apply_visual_plan(working,output,plan,publish)
            else:
                from review_table_filler import _Package
                from workbook_structure import write_package
                if Path(output).exists():raise ValueError('输出文件已存在，请重新另存。')
                package = _Package(working)
                if hashlib.sha256(package.raw).hexdigest()!=plan['digest']:raise ValueError('预览文件已变化，请重新打开。')
                write_package(package,output,{},progress=publish)
                result = {'count':0,'layout_count':0,'sheets':[],'output_path':output}
            result['count'] += applied['count'];result['layout_count'] += applied['layout_count']
            result['structure_count'] = applied['structure_count']
            result['sheets'] = list(dict.fromkeys([*applied['sheets'],*result['sheets']]))
            return result
        def done(result):
            self._saving = False
            if self.window.winfo_exists():self.window.destroy()
            else:self._scratch.cleanup()
            self.owner._reset_busy_state()
            self.owner._last_output_dir_to_open = os.path.dirname(output)
            self.owner.status_var.set(f"已保存 {result['count']} 个单元格、{result['layout_count']} 处行调整")
            self.owner._show_result_window(f"已保存：{'、'.join(result['sheets'])}\n修改 {result['count']} 个单元格、{result['layout_count']} 处行调整、{result['structure_count']} 次结构编辑\n输出：{output}",title='填表结果')
        def failed(exc):
            self._saving = False
            self.owner._finish_background_error(exc,'可视化填表')
            if self.window.winfo_exists():
                self.save_progress.stop();self.save_progress.grid_remove();self.refresh()
            else:self._scratch.cleanup()
        self.owner._task_runner.submit(work,done,failed,progress)
