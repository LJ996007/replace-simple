import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook,load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from review_table_filler import _xml,_tag,_serialize
from workbook_structure import edit_structure
from visual_fill import plan_visual_fill,apply_visual_plan


class InsertFormulaTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.source=Path(self.tmp.name)/'source.xlsx'
        book=Workbook();ws=book.active;ws.title='Data'
        for r in range(2,5):ws.cell(r,1,r);ws.cell(r,2,f'=A{r}*2')
        ws['C2']=ArrayFormula(ref='C2:C4',text='=A2:A4*3')
        ws.column_dimensions['B'].width=24
        book.create_sheet('Total')['A1']='=Data!B4+Data!C4'
        book.save(self.source);book.close()
        with ZipFile(self.source) as z:parts={e.filename:(e,z.read(e.filename)) for e in z.infolist()}
        root=_xml(parts['xl/worksheets/sheet1.xml'][1])
        for c in root.iter(_tag('c')):
            if c.get('r') in ('B2','B3','B4'):
                f=c.find(_tag('f'));f.set('t','shared');f.set('si','0')
                if c.get('r')=='B2':f.set('ref','B2:B4')
                else:f.text=None
        with ZipFile(self.source,'w') as z:
            for name,(e,data) in parts.items():z.writestr(e,_serialize(root) if name=='xl/worksheets/sheet1.xml' else data)

    def test_insert_row_inside_shared_and_array(self):
        output=Path(self.tmp.name)/'rows.xlsx'
        edit_structure(self.source,output,'Data','insert_above',(1,3,3,3),1)
        book=load_workbook(output)
        self.assertEqual(book['Data']['B4'].value,'=A4*2')
        self.assertEqual(book['Data']['C2'].value.ref,'C2:C5')
        self.assertEqual(book['Data']['C2'].value.text,'=A2:A5*3')
        self.assertEqual(book['Total']['A1'].value,'=Data!B5+Data!C5');book.close()

    def test_insert_column_updates_shared_array_and_width(self):
        output=Path(self.tmp.name)/'cols.xlsx'
        edit_structure(self.source,output,'Data','insert_left',(2,2,2,4),1)
        book=load_workbook(output);ws=book['Data']
        self.assertIsNone(ws['B2'].value)
        self.assertEqual(ws['C3'].value,'=A3*2')
        self.assertEqual(ws['D2'].value.ref,'D2:D4')
        self.assertEqual(ws.column_dimensions['B'].width,24)
        self.assertEqual(book['Total']['A1'].value,'=Data!C4+Data!D4');book.close()

    def test_dimensions_save_and_real_header_drag(self):
        import main
        from visual_table import VisualTable
        from visual_fill import inspect_visual_workbook
        root=tk.Tk()
        try:
            app=main.ReplaceSimpleApp(root,restore_session=False)
            events=[]
            grid=VisualTable(root,app,on_resize=lambda *args:events.append(args));grid.pack()
            grid.load(inspect_visual_workbook(self.source)['Data']);root.update()
            x=int(grid.xs[1]);y=10
            grid.canvas.event_generate('<ButtonPress-1>',x=x,y=y)
            grid.canvas.event_generate('<B1-Motion>',x=x+60,y=y)
            grid.canvas.event_generate('<ButtonRelease-1>',x=x+60,y=y)
            self.assertEqual(events[-1][:2],('column',1))
            self.assertGreater(events[-1][2],89)
            y=int(grid.ys[2])
            grid.canvas.event_generate('<ButtonPress-1>',x=10,y=y)
            grid.canvas.event_generate('<B1-Motion>',x=10,y=y+60)
            grid.canvas.event_generate('<ButtonRelease-1>',x=10,y=y+60)
            self.assertEqual(events[-1][:2],('row',2))
            self.assertGreater(events[-1][2],20)
            output=Path(self.tmp.name)/'sizes.xlsx'
            plan=plan_visual_fill(self.source,{}, {'Data':{'heights':{2:80},'widths':{2:40}}})
            apply_visual_plan(self.source,output,plan)
            book=load_workbook(output);self.assertEqual(book['Data'].column_dimensions['B'].width,40)
            self.assertEqual(book['Data'].row_dimensions[2].height,80);book.close()
        finally:root.destroy()

    def test_themed_input_validation_and_cancel(self):
        import main
        from app_dialogs import AppDialog
        root=tk.Tk()
        try:
            main.ReplaceSimpleApp(root,restore_session=False)
            dialog=AppDialog('插入列','插入列数',root,'integer',1,1,1000)
            dialog.value.set('1.5');dialog.apply()
            self.assertTrue(dialog.error.get());self.assertTrue(dialog.window.winfo_exists())
            dialog.value.set('3');dialog.apply();self.assertEqual(dialog.result,3)
            dialog=AppDialog('提示','测试提示',root);dialog.close();self.assertIsNone(dialog.result)
        finally:root.destroy()
