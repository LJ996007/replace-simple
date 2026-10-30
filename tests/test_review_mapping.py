import os
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment

from review_mapping import inspect_workbook, plan_fill, apply_plan, save_profile, load_profiles, layout_signature
from review_table_filler import ReviewCriterion, ReviewRules, extract_review_rules
from word_table_exporter import ExportedTable, ExportedTableCell


class MappingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name)/"template.xlsx"
        self.output = Path(self.temp.name)/"output.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.title = "任意表名"
        for ref, value in {"A1":"标准说明", "C1":"名称", "F1":"分值", "C2":"乙项", "C3":"甲项",
                           "A2":"旧说明", "D2":"=F2*2", "F4":"=SUM(F2:F3)", "B5":"签字"}.items():
            sheet[ref] = value
        sheet.merge_cells("A2:B2")
        sheet.merge_cells("A3:B3")
        sheet["A2"].font = Font(name="宋体",size=12,bold=True)
        sheet["A2"].alignment = Alignment(wrap_text=True)
        sheet.row_dimensions[2].height = 88
        sheet.column_dimensions['A'].width = 32
        sheet.print_area = "A1:G5"
        book.create_sheet("汇总")["A1"] = "='任意表名'!F4"
        book.save(self.path)
        self.rules = ReviewRules((ReviewCriterion("甲项","=原文",Decimal('2.5'),2,2),
                                  ReviewCriterion("乙项","乙项原文",Decimal('3'),3,3)), ())
        self.config = dict(sheet="任意表名",header=1,start=2,end=3,factor="C",description="A",points="F",mode="match")

    def test_reordered_rows_columns_preserve_layout_formulas_and_other_parts(self):
        plan = plan_fill(self.path,self.rules,self.config)
        self.assertFalse(plan['errors'])
        self.assertEqual([e['row'] for e in plan['entries']], [3,2])
        before = self.path.read_bytes()
        apply_plan(self.path,self.output,plan)
        self.assertEqual(before,self.path.read_bytes())
        with ZipFile(self.path) as a, ZipFile(self.output) as b:
            changed = {n for n in a.namelist() if a.read(n)!=b.read(n)}
            self.assertEqual(changed,{'xl/worksheets/sheet1.xml','xl/workbook.xml'})
        book=load_workbook(self.output);sheet=book['任意表名']
        self.assertEqual(sheet['A3'].value,'=原文')
        self.assertEqual(sheet['A3'].data_type,'s')
        self.assertEqual(sheet['F3'].value,2.5)
        self.assertEqual(sheet['D2'].value,'=F2*2')
        self.assertEqual(sheet['F4'].value,'=SUM(F2:F3)')
        self.assertEqual(sheet.row_dimensions[2].height,88)
        self.assertEqual(sheet['A2'].font.name,'宋体')
        self.assertTrue(sheet['A2'].font.bold)
        self.assertEqual(str(sheet.merged_cells),'A2:B2 A3:B3')
        self.assertEqual(sheet.print_area,"'任意表名'!$A$1:$G$5")
        book.close()

    def test_manual_source_columns_and_price_toggle(self):
        table = ExportedTable([ExportedTableCell(r,c,1,1,v) for r,row in enumerate(
            [['自定义分','自定义要求','自定义项目'],['30','价格说明','价格'],['70','技术说明','甲项']],1)
            for c,v in enumerate(row,1)],3,3)
        mapping=dict(header=1,factor=3,description=2,points=1)
        rules=extract_review_rules(table,mapping)
        self.assertEqual(rules.total,70)
        self.assertEqual(extract_review_rules(table,mapping,True).total,100)

    def test_missing_duplicate_and_manual_row_resolution(self):
        book=load_workbook(self.path);book.active['C2']='甲项';book.save(self.path);book.close()
        plan=plan_fill(self.path,self.rules,self.config)
        self.assertEqual(len(plan['errors']),2)
        fixed=plan_fill(self.path,self.rules,self.config,{0:3,1:2})
        self.assertFalse(fixed['errors'])
        self.assertTrue(plan_fill(self.path,self.rules,self.config,{0:2,1:2})['errors'])
        skipped=plan_fill(self.path,self.rules,self.config,{0:3,1:0})
        self.assertFalse(skipped['errors'])
        self.assertEqual(skipped['entries'][1]['status'],'已跳过')

    def test_formula_protection_and_stale_preview(self):
        plan=plan_fill(self.path,self.rules,self.config)
        book=load_workbook(self.path);book.active['F2']='=1+2';book.save(self.path);book.close()
        with self.assertRaisesRegex(ValueError,'发生变化'):
            apply_plan(self.path,self.output,plan)
        new=plan_fill(self.path,self.rules,self.config)
        self.assertTrue(any('公式' in e for e in new['errors']))
        with self.assertRaises(ValueError):
            apply_plan(self.path,self.output,new)
        self.assertFalse(self.output.exists())

    def test_sequence_capacity_and_merge_conflict(self):
        config=dict(self.config,mode='sequence',end=2)
        self.assertTrue(plan_fill(self.path,self.rules,config)['errors'])
        config=dict(self.config,mode='sequence')
        self.assertFalse(plan_fill(self.path,self.rules,config)['errors'])
        book=load_workbook(self.path);book.active.merge_cells('F2:F3');book.save(self.path);book.close()
        self.assertTrue(any('多个不同内容' in e for e in plan_fill(self.path,self.rules,config)['errors']))

    def test_profiles_persist_and_detect_template_structure_changes(self):
        info=inspect_workbook(self.path)['任意表名']
        with patch.dict(os.environ,{'APPDATA':self.temp.name}):
            save_profile('常用模板',self.config,dict(header=1,factor=1,description=2,points=3),info)
            profile=load_profiles()['常用模板']
            self.assertEqual(profile['config'],self.config)
            self.assertEqual(profile['signature'],layout_signature(info,self.config))
        book=load_workbook(self.path);book.active['C1']='项目名称';book.save(self.path);book.close()
        self.assertNotEqual(profile['signature'],layout_signature(inspect_workbook(self.path)['任意表名'],self.config))

    def test_array_formula_range_is_protected(self):
        from openpyxl.worksheet.formula import ArrayFormula
        book=load_workbook(self.path)
        book.active['F2']=ArrayFormula(ref='F2:F3',text='=C2:C3')
        book.save(self.path);book.close()
        plan=plan_fill(self.path,self.rules,self.config)
        self.assertTrue(any('F3' in error and '公式' in error for error in plan['errors']))

    def test_dialog_manual_mapping_reuse_and_background_write(self):
        import tkinter as tk
        import time
        import main
        from review_mapping_dialog import ReviewMappingDialog
        from word_table_exporter import TableScanItem
        from word_scan import file_stamp
        from ui_common import _file_identity
        source=Path(self.temp.name)/"source.docx"
        source.write_bytes(b'snapshot')
        table=ExportedTable([ExportedTableCell(r,c,1,1,v) for r,row in enumerate(
            [['名称','评分标准','分值'],['甲项','甲项要求','2.5'],['乙项','乙项要求','3']],1)
            for c,v in enumerate(row,1)],3,3)
        item=TableScanItem(str(source),source.name,1,'评分标准','','',3,3,'',table)
        root=tk.Tk()
        try:
            app=main.ReplaceSimpleApp(root,restore_session=False)
            app.show_workspace('extract')
            owner=app.table_export_window
            owner.scan_stamps={_file_identity(str(source)):file_stamp(source)}
            dialog=ReviewMappingDialog(owner,str(self.path),[item],inspect_workbook(self.path))
            dialog.target_vars['factor'].set('C')
            dialog.target_vars['description'].set('A')
            dialog.target_vars['points'].set('F')
            dialog.end.set('3')
            dialog.preview()
            root.update()
            self.assertFalse(dialog.plan['errors'])
            self.assertEqual([e['row'] for e in dialog.plan['entries']],[3,2])
            self.assertLessEqual(dialog.generate.winfo_rooty()+dialog.generate.winfo_height(),
                                 dialog.window.winfo_rooty()+dialog.window.winfo_height())
            dialog.window.geometry('800x600');root.update()
            self.assertLessEqual(dialog.generate.winfo_rooty()+dialog.generate.winfo_height(),
                                 dialog.window.winfo_rooty()+dialog.window.winfo_height())
            self.assertEqual(dialog.tree.cget('style'),'Scan.Treeview')
            self.assertEqual(dialog.generate.cget('style'),'Accent.TButton')
            dialog.edit_mapping();root.update()
            popup=next(w for w in dialog.window.winfo_children() if isinstance(w,tk.Toplevel))
            def descendants(widget):
                for child in widget.winfo_children():
                    yield child
                    yield from descendants(child)
            boxes=[w for w in descendants(popup) if isinstance(w,main.ttk.Combobox)]
            self.assertEqual(len(boxes),6)
            self.assertTrue(all(box.get() for box in boxes))
            self.assertTrue(any('名称' in box.get() for box in boxes))
            next(w for w in descendants(popup) if isinstance(w,main.ttk.Button) and w.cget('text')=='完成并预览').invoke()
            root.update()
            with patch.dict(os.environ,{'APPDATA':self.temp.name}), patch('review_mapping_dialog.simpledialog.askstring',return_value='测试配置'):
                dialog.store_profile()
                dialog.target_vars['description'].set('D')
                self.assertIsNone(dialog.plan)
                dialog.restore_profile()
                self.assertEqual(dialog.target_vars['description'].get(),'A')
                dialog.preview()
            with patch.object(owner,'_show_result_window') as result:
                dialog.write()
                deadline=time.monotonic()+15
                while owner._task_runner.active and time.monotonic()<deadline:
                    root.update();time.sleep(.01)
                self.assertFalse(owner._task_runner.active)
                result.assert_called_once()
            output=self.path.with_name('template_商务技术评审.xlsx')
            self.assertTrue(output.exists())
            book=load_workbook(output)
            self.assertEqual(book['任意表名']['A3'].value,'甲项要求')
            book.close()
        finally:
            root.destroy()

    def test_all_sheets_are_available_and_output_never_overwrites(self):
        self.assertEqual(list(inspect_workbook(self.path)),['任意表名','汇总'])
        plan=plan_fill(self.path,self.rules,self.config)
        with self.assertRaises(ValueError):
            apply_plan(self.path,self.path,plan)
        self.output.write_bytes(b'existing')
        with self.assertRaises(ValueError):
            apply_plan(self.path,self.output,plan)
        self.assertEqual(self.output.read_bytes(),b'existing')

    def test_empty_inline_string_in_template_is_read_as_blank(self):
        book=load_workbook(self.path)
        book.active['G2']=''
        book.save(self.path);book.close()
        self.assertEqual(inspect_workbook(self.path)['任意表名']['cells']['G2'],'')
        self.assertFalse(plan_fill(self.path,self.rules,self.config)['errors'])


if __name__=='__main__':
    unittest.main()
