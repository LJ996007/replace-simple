import os
import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import load_workbook

from workbook_structure import edit_structure
import test_worksheet_layout as fixtures


class WorkbookStructureTests(unittest.TestCase):
    def setUp(self):fixtures.WorksheetLayoutTests.setUp(self)

    def test_insert_inherits_format_and_updates_references(self):
        edit_structure(self.path,self.out,'评分','insert_above',(1,2,3,2),2)
        book = load_workbook(self.out); sheet = book['评分']
        self.assertIsNone(sheet['B2'].value);self.assertIsNone(sheet['B3'].value)
        self.assertEqual(sheet['B4'].value,'旧说明')
        self.assertTrue(sheet['B2'].font.bold)
        self.assertEqual(sheet.row_dimensions[2].height,90)
        self.assertEqual(sheet['C9'].value,'=SUM($C$4:C8)')
        self.assertEqual(book['汇总']['A1'].value,"='评分'!$C$9")
        self.assertIn('$A$1:$C$11',str(sheet.print_area))
        self.assertIn('A11:C11',str(sheet.merged_cells));book.close()

    def test_delete_content_row_and_undo_is_original_file(self):
        original = self.path.read_bytes()
        edit_structure(self.path,self.out,'评分','delete_rows',(1,2,3,2))
        self.assertEqual(self.path.read_bytes(),original)
        book=load_workbook(self.out);sheet=book['评分']
        self.assertIsNone(sheet['B2'].value)
        self.assertEqual(sheet['A5'].value,'服务')
        self.assertEqual(sheet['C6'].value,'=SUM($C$2:C5)')
        self.assertEqual(book['汇总']['A1'].value,"='评分'!$C$6")
        book.close()

    def test_merge_text_unmerge_and_numeric_value(self):
        edit_structure(self.path,self.out,'评分','merge',(2,2,2,6))
        book=load_workbook(self.out);sheet=book['评分']
        self.assertEqual(sheet['B2'].value,'旧说明\n服务原文')
        self.assertIn('B2:B6',str(sheet.merged_cells));book.close()
        unmerged=Path(self.temp.name)/'unmerged.xlsx'
        edit_structure(self.out,unmerged,'评分','unmerge',(2,2,2,6))
        book=load_workbook(unmerged);sheet=book['评分']
        self.assertNotIn('B2:B6',str(sheet.merged_cells))
        self.assertEqual(sheet['B2'].value,'旧说明\n服务原文')
        self.assertIsNone(sheet['B6'].value);book.close()
        numeric=Path(self.temp.name)/'numeric.xlsx'
        edit_structure(self.path,numeric,'评分','merge',(3,2,3,4))
        book=load_workbook(numeric)
        self.assertEqual(book['评分']['C2'].value,10)
        self.assertEqual(book['评分']['C2'].data_type,'n');book.close()

    def test_structural_guards_and_other_parts_preserved(self):
        for action,bounds in [('merge',(3,2,3,7)),('merge',(3,2,3,6)),('delete_rows',(1,7,3,7))]:
            with self.assertRaises(ValueError):edit_structure(self.path,self.out,'评分',action,bounds)
            self.assertFalse(self.out.exists())
        edit_structure(self.path,self.out,'评分','merge',(2,2,2,4))
        with ZipFile(self.path) as before,ZipFile(self.out) as after:
            for name in before.namelist():
                if name not in {'xl/workbook.xml','xl/worksheets/sheet1.xml'}:
                    self.assertEqual(before.read(name),after.read(name),name)
        with self.assertRaisesRegex(ValueError,'一部分'):
            edit_structure(self.out,Path(self.temp.name)/'bad.xlsx','评分','delete_rows',(1,3,3,3))

    def test_insert_inside_merge_and_horizontal_template(self):
        book=load_workbook(self.path);sheet=book['评分']
        sheet.merge_cells('B2:B4');book.save(self.path);book.close()
        edit_structure(self.path,self.out,'评分','insert_above',(1,3,1,3),1)
        book=load_workbook(self.out)
        self.assertIn('B2:B5',str(book['评分'].merged_cells));book.close()
        output=Path(self.temp.name)/'below.xlsx'
        edit_structure(self.path,output,'评分','insert_below',(1,9,3,9),2)
        book=load_workbook(output)
        self.assertIn('A10:C10',str(book['评分'].merged_cells))
        self.assertIn('A11:C11',str(book['评分'].merged_cells));book.close()

    def test_ui_composed_edit_recipe_rollback_and_save(self):
        import main
        from visual_fill_dialog import VisualFillDialog
        from word_table_exporter import ExportedTable,ExportedTableCell,TableScanItem
        from word_scan import file_stamp
        from ui_common import _file_identity
        from app_settings import load_settings
        source=Path(self.temp.name)/'source.docx';source.write_bytes(b'snapshot')
        table=ExportedTable([ExportedTableCell(1,1,1,1,'说明'),ExportedTableCell(2,1,1,1,'新项目正文')],2,1)
        item=TableScanItem(str(source),source.name,1,'评分','','',2,1,'',table)
        root=tk.Tk()
        try:
            app=main.ReplaceSimpleApp(root,restore_session=False);app.show_workspace('extract')
            owner=app.table_export_window
            owner.scan_stamps={_file_identity(str(source)):file_stamp(source)}
            dialog=VisualFillDialog(owner,str(self.path),[item]);root.update()
            dialog.stage_drop(0,(1,2,1,2),'评分',2,2)
            dialog.stage_structure('insert_below','评分',(1,2,3,2),2)
            dialog.stage_drop(0,(1,2,1,2),'评分',3,2)
            dialog.stage_structure('merge','评分',(2,3,2,4))
            self.assertEqual(dialog.right.show_value('B3'),'新项目正文')
            dialog.stage_structure('unmerge','评分',(2,3,2,4))
            self.assertNotIn('B3:B4',dialog.right.info['merges'])
            self.assertIn('B3',dialog.right.updates)
            dialog.undo();self.assertIn('B3:B4',dialog.right.info['merges'])
            dialog.redo();self.assertNotIn('B3:B4',dialog.right.info['merges'])
            with patch.dict(os.environ,{'APPDATA':self.temp.name}),patch('visual_fill_dialog.simpledialog.askstring',return_value='结构设置'):
                dialog.save_recipe();recipe=load_settings()['review_drag_profiles']['结构设置']
                dialog.clear()
                self.assertEqual(dialog.right.show_value('B2'),'旧说明')
                dialog.sources[0]['cells']['A2']='下一项目内容'
                dialog.replay_recipe(recipe)
                self.assertEqual(dialog.right.show_value('B3'),'下一项目内容')
                before=dialog.snapshot()
                broken=dict(recipe,operations=[*recipe['operations'],{'kind':'structure','sheet':'评分','action':'merge','bounds':[3,2,3,9]}])
                with self.assertRaises(ValueError):dialog.replay_recipe(broken)
                self.assertEqual(dialog.working_template,before['working_template'])
                self.assertEqual(dialog.right.show_value('B3'),'下一项目内容')
            dialog.stage_structure('delete_rows','评分',(1,3,3,3))
            self.assertIsNone(dialog.sheets['评分']['cells'].get('B3') or None)
            with patch.object(owner,'_show_result_window') as result:
                dialog.write()
                deadline=time.monotonic()+15
                while owner._task_runner.active and time.monotonic()<deadline:
                    root.update();time.sleep(.01)
                result.assert_called_once()
            output=Path(self.temp.name)/'template_商务技术评审.xlsx'
            book=load_workbook(output)
            self.assertEqual(book['评分']['B2'].value,'下一项目内容')
            self.assertIsNone(book['评分']['B3'].value)
            self.assertEqual(book['汇总']['A1'].value,"='评分'!$C$8")
            book.close()
        finally:root.destroy()


if __name__=='__main__':unittest.main()
