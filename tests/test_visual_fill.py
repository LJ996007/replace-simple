import os
import tempfile
import time
import tkinter as tk
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.worksheet.formula import ArrayFormula

from visual_fill import (source_grid, drop_values, plan_visual_fill, apply_visual_plan,
                         inspect_visual_workbook, expand_selection)
from word_table_exporter import ExportedTable, ExportedTableCell, TableScanItem


class VisualFillTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name)/'template.xlsm'
        self.output = Path(self.temp.name)/'result.xlsm'
        book = Workbook();sheet = book.active;sheet.title='商务评审'
        sheet['A1']='评分说明';sheet['B1']='分值'
        sheet.merge_cells('A2:A3');sheet.merge_cells('B2:B3')
        sheet['A2']='旧说明';sheet['B4']='=SUM(B2:B3)'
        sheet['A2'].font=Font(name='宋体',bold=True)
        sheet.row_dimensions[2].height=80
        sheet.column_dimensions['A'].width=40
        sheet.print_area='A1:C5'
        book.create_sheet('其他')["D4"]='原内容'
        book.save(self.path);book.close()
        self.table = ExportedTable([
            ExportedTableCell(1,1,1,1,'名称'),ExportedTableCell(1,2,1,1,'评分标准'),ExportedTableCell(1,3,1,1,'分值'),
            ExportedTableCell(2,1,2,1,'技术方案'),ExportedTableCell(2,2,1,1,'第一条原文'),
            ExportedTableCell(3,2,1,1,'第二条原文'),ExportedTableCell(2,3,2,1,'10')],3,3)
        self.source = source_grid(self.table)

    def test_block_drop_merges_descriptions_and_counts_shared_score_once(self):
        target=inspect_visual_workbook(self.path)['商务评审']
        values=drop_values(self.source,(2,2,3,3),target,2,1)
        self.assertEqual(values,{'A2':'第一条原文\n\n第二条原文','B2':Decimal(10)})
        plan=plan_visual_fill(self.path,{'商务评审':values,'其他':{'D4':'=作为原文'}})
        self.assertFalse(plan['errors'])
        original=self.path.read_bytes()
        apply_visual_plan(self.path,self.output,plan)
        self.assertEqual(self.path.read_bytes(),original)
        book=load_workbook(self.output)
        self.assertEqual(book['商务评审']['A2'].value,values['A2'])
        self.assertEqual(book['商务评审']['B2'].value,10)
        self.assertEqual(book['商务评审']['B4'].value,'=SUM(B2:B3)')
        self.assertTrue(book['商务评审']['A2'].font.bold)
        self.assertEqual(book['商务评审'].row_dimensions[2].height,80)
        self.assertEqual(book['商务评审'].column_dimensions['A'].width,40)
        self.assertEqual(str(book['商务评审'].merged_cells),'A2:A3 B2:B3')
        self.assertEqual(book['其他']['D4'].data_type,'s')
        book.close()
        with ZipFile(self.path) as before, ZipFile(self.output) as after:
            for name in before.namelist():
                if name not in {'xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml','xl/workbook.xml'}:
                    self.assertEqual(before.read(name),after.read(name),name)

    def test_merged_selection_expands_and_empty_cells_do_not_clear_target(self):
        self.assertEqual(expand_selection(self.source,(3,2,3,2)),(3,2,3,3))
        target=inspect_visual_workbook(self.path)['商务评审']
        values=drop_values(self.source,(3,2,3,2),target,2,2)
        self.assertEqual(values,{'B2':Decimal(10)})
        blank=dict(self.source,cells={})
        self.assertEqual(drop_values(blank,(2,2,2,2),target,2,1),{})

    def test_distinct_scores_do_not_silently_collapse_into_one_merge(self):
        source={'cells':{'A1':'10','A2':'10'},'merges':[],'max_row':2,'max_col':1}
        with self.assertRaisesRegex(ValueError,'多个独立分值'):
            drop_values(source,(1,1,1,2),inspect_visual_workbook(self.path)['商务评审'],2,2)

    def test_formula_protection_stale_preview_and_never_overwrite(self):
        plan=plan_visual_fill(self.path,{'商务评审':{'B4':'改掉公式'}})
        self.assertTrue(plan['errors'])
        with self.assertRaises(ValueError):apply_visual_plan(self.path,self.output,plan)
        plan=plan_visual_fill(self.path,{'商务评审':{'A2':'新内容'}})
        with self.assertRaises(ValueError):apply_visual_plan(self.path,self.path,plan)
        self.output.write_bytes(b'existing')
        with self.assertRaises(ValueError):apply_visual_plan(self.path,self.output,plan)
        self.assertEqual(self.output.read_bytes(),b'existing')
        book=load_workbook(self.path);book['商务评审']['C1']='changed';book.save(self.path);book.close()
        with self.assertRaisesRegex(ValueError,'发生变化'):
            apply_visual_plan(self.path,Path(self.temp.name)/'new.xlsx',plan)

    def test_protected_sheet_and_array_formula(self):
        book=load_workbook(self.path)
        book['其他']['A1']=ArrayFormula(ref='A1:A3',text='=1')
        book['商务评审'].protection.sheet=True
        book.save(self.path);book.close()
        with self.assertRaisesRegex(ValueError,'保护'):
            plan_visual_fill(self.path,{'商务评审':{'A2':'x'}})
        self.assertTrue(plan_visual_fill(self.path,{'其他':{'A3':'x'}})['errors'])

    def test_automatic_matching_can_feed_visual_preview_without_writing(self):
        import main
        from visual_fill_dialog import VisualFillDialog
        book=load_workbook(self.path)
        sheet=book['商务评审'];sheet['C1']='名称';sheet['C2']='技术方案'
        book.save(self.path);book.close()
        item=TableScanItem('snapshot.docx','snapshot.docx',1,'评分标准','','',3,3,'',self.table)
        root=tk.Tk()
        try:
            app=main.ReplaceSimpleApp(root,restore_session=False);app.show_workspace('extract')
            dialog=VisualFillDialog(app.table_export_window,str(self.path),[item])
            dialog.auto_match();root.update()
            helper=dialog._auto_dialog
            self.assertIsNotNone(helper.plan)
            self.assertFalse(helper.plan['errors'])
            helper.write();root.update()
            self.assertEqual(dialog.pending['商务评审']['B2'],10)
            self.assertEqual(dialog.pending['商务评审']['A2'],'第一条原文\n\n第二条原文')
            self.assertFalse(self.output.exists())
        finally:
            root.destroy()

    def test_real_canvas_drag_undo_redo_recipe_and_background_write(self):
        import main
        from visual_fill_dialog import VisualFillDialog
        from word_scan import file_stamp
        from ui_common import _file_identity
        source=Path(self.temp.name)/'source.docx';source.write_bytes(b'snapshot')
        item=TableScanItem(str(source),source.name,1,'评分标准','','',3,3,'',self.table)
        root=tk.Tk()
        try:
            app=main.ReplaceSimpleApp(root,restore_session=False);app.show_workspace('extract')
            owner=app.table_export_window
            owner.scan_stamps={_file_identity(str(source)):file_stamp(source)}
            dialog=VisualFillDialog(owner,str(self.path),[item]);root.update()
            # 使用真实画布事件完成一次跨窗口控件拖动，而非直接调用提交函数。
            left,right=dialog.left,dialog.right
            x=int((left.xs[1]+left.xs[2])/2);y=int((left.ys[1]+left.ys[2])/2)
            left.canvas.event_generate('<ButtonPress-1>',x=x,y=y)
            left.canvas.event_generate('<ButtonRelease-1>',x=x,y=y);root.update()
            self.assertEqual(left.selection,(2,2,2,2))
            tx=int((right.xs[0]+right.xs[1])/2);ty=int((right.ys[1]+right.ys[2])/2)
            target_x=right.canvas.winfo_rootx()+tx;target_y=right.canvas.winfo_rooty()+ty
            left.canvas.event_generate('<ButtonPress-1>',x=x,y=y)
            left.canvas.event_generate('<B1-Motion>',x=target_x-left.canvas.winfo_rootx(),y=target_y-left.canvas.winfo_rooty(),rootx=target_x,rooty=target_y)
            left.canvas.event_generate('<ButtonRelease-1>',x=target_x-left.canvas.winfo_rootx(),y=target_y-left.canvas.winfo_rooty(),rootx=target_x,rooty=target_y)
            root.update()
            self.assertEqual(dialog.pending['商务评审']['A2'],'第一条原文')
            dialog.undo();self.assertEqual(dialog.pending,{})
            dialog.redo();self.assertEqual(dialog.pending['商务评审']['A2'],'第一条原文')
            self.assertEqual(dialog.window.transient(),'')
            dialog.toggle_maximize();root.update()
            self.assertEqual(dialog.window.state(),'zoomed')
            dialog.toggle_maximize();root.update()
            dialog.toggle_fullscreen();root.update()
            self.assertTrue(dialog.window.attributes('-fullscreen'))
            dialog.leave_fullscreen();root.update()
            self.assertFalse(dialog.window.attributes('-fullscreen'))
            right.select(2,1)
            right.canvas.event_generate('<Control-MouseWheel>',delta=120,x=100,y=100)
            self.assertEqual(right.zoom,1.1)
            right.zoom_buttons[0].event_generate('<Button-1>')
            self.assertEqual(right.zoom,1)
            right.zoom_buttons[1].event_generate('<Button-1>')
            self.assertEqual(right.zoom_label.get(),'110%')
            right.set_zoom(.5);self.assertEqual(right.zoom,.5)
            right.set_zoom(9);self.assertEqual(right.zoom,2.5)
            right.set_zoom(.1);self.assertEqual(right.zoom,.5)
            right.set_zoom(1.5)
            root.update()
            right.canvas.xview_moveto(0);right.canvas.yview_moveto(0);right.draw()
            px=int((right.xs[0]+right.xs[1])/2);py=int((right.ys[1]+right.ys[2])/2)
            self.assertEqual(right.cell_at(px,py),(2,1))
            # 缩放后再完成一次真实跨画布拖放，不能因像素比例改变写错格。
            tx=right.canvas.winfo_rootx()+px;ty=right.canvas.winfo_rooty()+py
            left.canvas.event_generate('<ButtonPress-1>',x=x,y=y)
            left.canvas.event_generate('<B1-Motion>',x=tx-left.canvas.winfo_rootx(),y=ty-left.canvas.winfo_rooty(),rootx=tx,rooty=ty)
            left.canvas.event_generate('<ButtonRelease-1>',x=tx-left.canvas.winfo_rootx(),y=ty-left.canvas.winfo_rooty(),rootx=tx,rooty=ty)
            self.assertEqual(dialog.pending['商务评审']['A2'],'第一条原文')
            self.assertEqual(right.selection,(1,2,1,3))
            dialog.stage_drop(0,(2,3,2,3),'其他',4,4)
            with patch.dict(os.environ,{'APPDATA':self.temp.name}),patch('visual_fill_dialog.simpledialog.askstring',return_value='常用拖放'):
                dialog.save_recipe()
                from app_settings import load_settings
                recipe=load_settings()['review_drag_profiles']['常用拖放']
                dialog.clear();dialog.replay_recipe(recipe)
                self.assertEqual(dialog.pending['其他']['D4'],'第二条原文')
            dialog.window.geometry('900x600');root.update()
            self.assertLessEqual(dialog.generate.winfo_rooty()+dialog.generate.winfo_height(),dialog.window.winfo_rooty()+dialog.window.winfo_height())
            with patch.object(owner,'_show_result_window') as result:
                dialog.write()
                deadline=time.monotonic()+15
                while owner._task_runner.active and time.monotonic()<deadline:
                    root.update();time.sleep(.01)
                self.assertFalse(owner._task_runner.active);result.assert_called_once()
            output=Path(self.temp.name)/'template_商务技术评审.xlsm'
            book=load_workbook(output)
            self.assertEqual(book['商务评审']['A2'].value,'第一条原文')
            self.assertEqual(book['其他']['D4'].value,'第二条原文');book.close()
        finally:
            root.destroy()


if __name__=='__main__':unittest.main()
