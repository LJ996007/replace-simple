import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

from visual_fill import inspect_visual_workbook, plan_visual_fill, apply_visual_plan
from worksheet_layout import RowMap, empty_rows, auto_heights, preview_layout, shift_formula


class WorksheetLayoutTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name)/'template.xlsx'
        self.out = Path(self.temp.name)/'result.xlsx'
        book = Workbook(); sheet = book.active; sheet.title = '评分'
        sheet.append(['项目', '说明', '分值'])
        sheet.append(['方案', '旧说明', 10])
        sheet['B2'].font = Font(name='宋体',size=12,bold=True)
        sheet['B2'].fill = PatternFill('solid',fgColor='FFEEAA')
        sheet.column_dimensions['B'].width = 28
        sheet.row_dimensions[2].height = 90
        sheet['A6'] = '服务'; sheet['B6'] = '服务原文'; sheet['C6'] = 20
        sheet['A7'] = '合计'; sheet['C7'] = '=SUM($C$2:C6)'
        sheet['A9'] = '评委签字'; sheet.merge_cells('A9:C9')
        sheet.print_area = 'A1:C9'; sheet.print_title_rows = '1:1'
        sheet.freeze_panes = 'B2'
        for row in (3,4,5):
            sheet.cell(row,2).fill = PatternFill('solid',fgColor='FFFFFF')
        rule = DataValidation(type='whole',operator='between',formula1='0',formula2='100')
        rule.add('C2:C6'); sheet.add_data_validation(rule)
        other = book.create_sheet('汇总'); other['A1'] = "='评分'!$C$7"
        book.defined_names.add(DefinedName('总分',attr_text="'评分'!$C$7"))
        book.save(self.path); book.close()

    def test_autofit_merged_rows_and_format_only_save(self):
        book = load_workbook(self.path); sheet = book['评分']
        sheet.merge_cells('B3:B4'); sheet['B3'] = '合并说明内容' * 35
        book.save(self.path); book.close()
        info = inspect_visual_workbook(self.path)['评分']
        heights, wraps = auto_heights(info,{},[2,3])
        self.assertLess(heights[2],90)
        self.assertIn(4,heights)
        self.assertGreater(heights[3]+heights[4],80)
        layout = {'评分':{'heights':heights,'wrap_cells':wraps,'delete_rows':[]}}
        plan = plan_visual_fill(self.path,{},layout)
        self.assertEqual(plan['count'],0); self.assertTrue(plan['layout_count'])
        apply_visual_plan(self.path,self.out,plan)
        book = load_workbook(self.out); sheet = book['评分']
        self.assertTrue(sheet['B2'].alignment.wrap_text)
        self.assertTrue(sheet['B2'].font.bold)
        self.assertEqual(sheet['B2'].font.name,'宋体')
        self.assertEqual(sheet['B2'].fill.fgColor.rgb,'00FFEEAA')
        self.assertEqual(sheet.row_dimensions[3].height,heights[3])
        self.assertEqual(str(sheet.merged_cells),'A9:C9 B3:B4')
        self.assertEqual(sheet['C7'].value,'=SUM($C$2:C6)'); book.close()
        with ZipFile(self.path) as before, ZipFile(self.out) as after:
            for name in before.namelist():
                if name not in {'xl/worksheets/sheet1.xml','xl/styles.xml','xl/workbook.xml'}:
                    self.assertEqual(before.read(name),after.read(name),name)

    def test_delete_rows_moves_formulas_names_print_ranges_and_preserves_values(self):
        info = inspect_visual_workbook(self.path)['评分']
        pending = {'B2':'新评分说明' * 30}
        self.assertEqual(empty_rows(info,pending,[2,3,4,5,6,7]),[3,4,5])
        layout = {'heights':{2:130},'delete_rows':[3,4],'wrap_cells':['B2']}
        preview,updates = preview_layout(info,pending,layout,'评分',{'评分':layout})
        self.assertEqual(preview['cells']['A4'],'服务')
        self.assertEqual(preview['merges'],['A7:C7'])
        self.assertEqual(updates['B2'],pending['B2'])
        self.assertEqual(preview['cells']['C5'],'=SUM($C$2:C4)')
        original = self.path.read_bytes()
        plan = plan_visual_fill(self.path,{'评分':pending},{'评分':layout})
        apply_visual_plan(self.path,self.out,plan)
        self.assertEqual(original,self.path.read_bytes())
        book = load_workbook(self.out); sheet = book['评分']
        self.assertEqual(sheet['B2'].value,pending['B2'])
        self.assertEqual(sheet['A4'].value,'服务')
        self.assertEqual(sheet['C5'].value,'=SUM($C$2:C4)')
        self.assertEqual(book['汇总']['A1'].value,"='评分'!$C$5")
        self.assertEqual(book.defined_names['总分'].attr_text,"'评分'!$C$5")
        self.assertIn('$A$1:$C$7',str(sheet.print_area))
        self.assertEqual(sheet.print_title_rows,'$1:$1')
        self.assertEqual(str(sheet.data_validations.dataValidation[0].sqref),'C2:C4')
        self.assertEqual(sheet.freeze_panes,'B2')
        self.assertEqual(sheet.row_dimensions[2].height,130)
        self.assertEqual(sheet['B2'].fill.fgColor.rgb,'00FFEEAA')
        book.close()

    def test_empty_detection_whole_row_and_vertical_merges(self):
        book = load_workbook(self.path); sheet = book['评分']
        sheet.merge_cells('A3:A4'); sheet['A3'] = '跨行组'
        sheet['Z5'] = '选区外文字'; book.save(self.path); book.close()
        info = inspect_visual_workbook(self.path)['评分']
        self.assertEqual(empty_rows(info,{},[3,4,5]),[])
        for row in (3,4,5,7):
            with self.assertRaisesRegex(ValueError,'内容、公式或跨行合并'):
                plan_visual_fill(self.path,{}, {'评分':{'delete_rows':[row]}})

    def test_formula_literals_lowercase_and_row_mapping(self):
        mapping = RowMap([3,5])
        self.assertEqual([mapping.original(r) for r in range(1,6)],[1,2,4,6,7])
        self.assertEqual(mapping.ref('$A$2:$C$7'),'$A$2:$C$5')
        formula = '=IF(a7="A7",SUM($C$2:c6),0)'
        self.assertEqual(shift_formula(formula,'评分',{'评分':mapping}),'=IF(a5="A7",SUM($C$2:c4),0)')
        self.assertEqual(shift_formula('=DATA!A7','Other',{'Data':mapping}),'=DATA!A5')
        for formula in ('=C3','=INDIRECT("C7")',"='评分:汇总'!C7"):
            with self.assertRaises(ValueError):shift_formula(formula,'评分',{'评分':mapping})

    def test_unsupported_objects_and_direct_references_fail_before_write(self):
        book = load_workbook(self.path)
        book['汇总']['B1'] = "='评分'!B3"
        book.save(self.path); book.close()
        with self.assertRaisesRegex(ValueError,'直接引用'):
            plan_visual_fill(self.path,{}, {'评分':{'delete_rows':[3]}})
        self.assertFalse(self.out.exists())
        book = load_workbook(self.path)
        book['汇总']['B1'] = None
        from openpyxl.comments import Comment
        book['评分']['A6'].comment = Comment('评审备注','测试')
        book.save(self.path); book.close()
        # 批注是浮动对象，锚点随删行移动后随单元格保持相对位置，不再拒绝。
        layout = {'评分': {'heights': {}, 'delete_rows': [3], 'wrap_cells': []}}
        plan = plan_visual_fill(self.path, {}, layout)
        apply_visual_plan(self.path, self.out, plan)
        moved = load_workbook(self.out)['评分']['A5'].comment
        self.assertIn('评审备注', moved.text if moved else '')
        from openpyxl.worksheet.table import Table
        book = load_workbook(self.path)
        book['评分'].add_table(Table(displayName='分项表', ref='A1:C6'))
        book.save(self.path); book.close()
        with self.assertRaisesRegex(ValueError,'特殊对象'):
            plan_visual_fill(self.path,{}, {'评分':{'delete_rows':[3]}})
        plan_visual_fill(self.path,{}, {'评分':{'heights':{2:120}}})

    def test_manual_height_protection_and_limits(self):
        for height in (0,410,float('nan'),float('inf')):
            with self.assertRaises(ValueError):
                plan_visual_fill(self.path,{}, {'评分':{'heights':{2:height}}})
        info = inspect_visual_workbook(self.path)['评分']
        with self.assertRaisesRegex(ValueError,'上限'):
            auto_heights(info,{'B2':'很多文字'*3000},[2])
        book = load_workbook(self.path); book['评分'].protection.sheet=True
        book.save(self.path); book.close()
        with self.assertRaisesRegex(ValueError,'保护'):
            plan_visual_fill(self.path,{}, {'评分':{'heights':{2:50}}})

    def test_ui_layout_undo_redo_drop_after_deletion_and_recipe(self):
        import main
        from visual_fill_dialog import VisualFillDialog
        from word_table_exporter import ExportedTable, ExportedTableCell, TableScanItem
        from app_settings import load_settings
        table = ExportedTable([ExportedTableCell(1,1,1,1,'说明'),ExportedTableCell(2,1,1,1,'新说明'*25)],2,1)
        item = TableScanItem('source.docx','source.docx',1,'评分','','',2,1,'',table)
        root = tk.Tk()
        try:
            app = main.ReplaceSimpleApp(root,restore_session=False); app.show_workspace('extract')
            dialog = VisualFillDialog(app.table_export_window,str(self.path),[item]); root.update()
            dialog.right.select(3,2); dialog.right.select(4,2,True)
            dialog.adjust_rows('delete'); root.update()
            self.assertEqual(dialog.layouts['评分']['delete_rows'],[3,4])
            self.assertEqual(dialog.right.show_value('A4'),'服务')
            dialog.stage_drop(0,(1,2,1,2),'评分',4,2)
            self.assertEqual(dialog.pending['评分']['B6'],'新说明'*25)
            dialog.right.select(4,2); dialog.fit_button.invoke(); root.update()
            height = dialog.layouts['评分']['heights'][6]
            self.assertGreater(height,18)
            self.assertAlmostEqual(dialog.right.info['heights'][4],height*4/3)
            dialog.undo(); self.assertNotIn(6,dialog.layouts['评分']['heights'])
            dialog.redo(); self.assertEqual(dialog.layouts['评分']['heights'][6],height)
            with patch.dict('os.environ',{'APPDATA':self.temp.name}),patch('visual_fill_dialog.simpledialog.askstring',return_value='排版'):
                dialog.save_recipe(); recipe = load_settings()['review_drag_profiles']['排版']
                dialog.clear(); self.assertFalse(dialog.layouts)
                # 复用后依据新内容计算，而非复用旧文案或旧行高。
                dialog.sources[0]['cells']['A2'] = '新的更短内容'
                dialog.replay_recipe(recipe)
                self.assertEqual(dialog.pending['评分']['B6'],'新的更短内容')
                self.assertLess(dialog.layouts['评分']['heights'][6],height)
            dialog.clear()
            dialog.right.select(2,2)
            with patch('visual_fill_dialog.ask_row_height',return_value=62):
                dialog.adjust_rows('height')
            self.assertEqual(dialog.layouts['评分']['heights'][2],62)
            self.assertEqual(str(dialog.generate['state']),'normal')
            # 行号拖动选择整行；鼠标坐标以当前实际行高为准。
            cv = dialog.right.canvas
            cv.yview_moveto(0); root.update()
            y1 = int((dialog.right.ys[2]+dialog.right.ys[3])/2)
            y2 = int((dialog.right.ys[3]+dialog.right.ys[4])/2)
            cv.event_generate('<ButtonPress-1>',x=15,y=y1)
            cv.event_generate('<B1-Motion>',x=15,y=y2)
            cv.event_generate('<ButtonRelease-1>',x=15,y=y2)
            self.assertEqual(dialog.right.selection,(1,3,dialog.right.info['max_col'],4))
            for size in ('900x600','1200x760'):
                dialog.window.geometry(size); root.update()
                self.assertLessEqual(dialog.edit_button.winfo_rootx()+dialog.edit_button.winfo_width(),dialog.right.zoom_buttons[0].winfo_rootx())
        finally:
            root.destroy()


if __name__=='__main__':unittest.main()
