import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from openpyxl import Workbook

import word_scan
from visual_fill import plan_visual_fill,apply_visual_plan


class ActualProgressTests(unittest.TestCase):
    def test_scan_advances_after_real_stages(self):
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'source.docx'
            doc=Document();doc.add_table(rows=1,cols=1).cell(0,0).text='评分';doc.save(source)
            events=[]
            original_tables=word_scan.scan_word_tables
            original_sections=word_scan.scan_symbol_clause_sections
            def tables(*args,**kwargs):
                self.assertEqual(events[-1][0],1)
                self.assertLess(events[-1][0],events[-1][1])
                return original_tables(*args,**kwargs)
            def sections(*args,**kwargs):
                self.assertEqual(events[-1][0],2)
                self.assertLess(events[-1][0],events[-1][1])
                return original_sections(*args,**kwargs)
            with patch.object(word_scan,'scan_word_tables',side_effect=tables),patch.object(word_scan,'scan_symbol_clause_sections',side_effect=sections):
                result=word_scan.scan_word_content([str(source)],'★',lambda *args:events.append(args))
            self.assertEqual([e[0] for e in events],[0,1,2,3,4,5])
            self.assertEqual(events[-1][0],events[-1][1])
            export_events=[]
            exporter=word_scan.export_word_tables_to_excel
            def export(*args,**kwargs):
                self.assertEqual(export_events[-1][0],0)
                self.assertLess(export_events[-1][0],export_events[-1][1])
                return exporter(*args,**kwargs)
            with patch.object(word_scan,'export_word_tables_to_excel',side_effect=export):
                exported=word_scan.export_scanned_content(result[0][0],[],result[2],folder,'★',False,lambda *args:export_events.append(args))
            self.assertIsNone(exported[2])
            self.assertEqual([e[0] for e in export_events],[0,1])

    def test_save_only_reaches_full_after_output_committed(self):
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'template.xlsx';output=Path(folder)/'result.xlsx'
            book=Workbook();book.active['A1']='原文';book.save(source);book.close()
            plan=plan_visual_fill(source,{'Sheet':{'A1':'新文'}})
            events=[]
            def progress(current,total,status):
                events.append((current,total))
                if current==total:self.assertTrue(output.exists())
                else:self.assertFalse(output.exists())
            apply_visual_plan(source,output,plan,progress)
            self.assertGreater(len(events),2)
            self.assertEqual(events[-1][0],events[-1][1])
            self.assertEqual(sorted(events),events)

    def test_wait_animation_resets_single_file_maximum(self):
        import main
        root=tk.Tk()
        try:
            app=main.ReplaceSimpleApp(root,restore_session=False);app.show_workspace('extract')
            owner=app.table_export_window
            owner._begin_determinate_task('扫描',1)
            owner._update_progress(1,1,'source.docx')
            owner._begin_indeterminate_task('读取模板')
            self.assertEqual(float(owner.progress['maximum']),100)
            self.assertEqual(str(owner.progress['mode']),'indeterminate')
            self.assertLess(float(owner.progress['value']),100)
            owner._reset_busy_state()
        finally:root.destroy()


if __name__=='__main__':unittest.main()
