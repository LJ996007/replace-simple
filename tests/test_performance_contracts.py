"""性能回归检查：验证工作量有界、状态正确，不依赖机器速度。"""
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from docx import Document
from openpyxl import Workbook

from ui_common import BackgroundTaskRunner
from visual_table import VisualTable
from visual_fill_dialog import VisualFillDialog
import visual_fill
import word_scan
from worksheet_layout import preview_layout


class PerformanceContracts(unittest.TestCase):
    def test_preview_does_not_mutate_shared_template_styles(self):
        style={'wrap':False}
        info={'cells':{'A1':'a','B1':'b'},'styles':{'A1':style,'B1':style},
              'merges':[],'max_row':1,'max_col':2}
        preview,_=preview_layout(info,{}, {'wrap_cells':['A1']})
        self.assertTrue(preview['styles']['A1']['wrap'])
        self.assertFalse(preview['styles']['B1']['wrap'])
        self.assertFalse(info['styles']['A1']['wrap'])

    def test_progress_burst_retains_latest_and_completion_without_backlog(self):
        runner=BackgroundTaskRunner(Mock())
        success,progress=Mock(),Mock()
        def work(publish):
            for i in range(10000):publish(i,9999)
            self.assertEqual(runner._events.qsize(),0)
            return 'done'
        with patch('ui_common.threading.Thread') as thread:
            thread.side_effect=lambda target,**_:SimpleNamespace(start=target)
            runner.submit(work,success,self.fail,progress)
        progress.assert_called_once_with(9999,9999)
        success.assert_called_once_with('done')
        self.assertFalse(runner.active)

    def test_history_shares_readonly_sheet_versions_but_copies_edits(self):
        dialog=VisualFillDialog.__new__(VisualFillDialog)
        for key in ('pending','operations','layouts','working_template','digest','applied','marked'):
            setattr(dialog,key,{'value':[]})
        dialog.sheets={'Data':{'cells':{'A1':'old'}}}
        snapshot=dialog.snapshot()
        self.assertIs(snapshot['sheets'],dialog.sheets)
        dialog.pending['value'].append('new')
        dialog.sheets={'Data':{'cells':{'A1':'new'}}}
        dialog.restore_snapshot(snapshot)
        self.assertEqual(dialog.pending,{'value':[]})
        self.assertEqual(dialog.sheets['Data']['cells']['A1'],'old')
        dialog.history=[];dialog.future=[]
        for _ in range(110):dialog.checkpoint()
        self.assertEqual(len(dialog.history),100)

    def test_scan_shares_one_numbering_snapshot_between_extractors(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'source.docx'
            doc=Document();doc.add_table(rows=1,cols=1).cell(0,0).text='★要求';doc.save(path)
            with patch.object(word_scan,'document_numbered_text',wraps=word_scan.document_numbered_text) as numbering, \
                 patch.object(word_scan,'scan_word_tables',wraps=word_scan.scan_word_tables) as tables, \
                 patch.object(word_scan,'scan_symbol_clause_sections',wraps=word_scan.scan_symbol_clause_sections) as sections:
                result=word_scan.scan_word_content([str(path)],'★')
            self.assertIsNone(result[0][2])
            self.assertEqual(numbering.call_count,1)
            self.assertIs(tables.call_args.kwargs['numbered'],sections.call_args.kwargs['numbered'])

    def test_layout_plan_opens_one_workbook_package(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'template.xlsx'
            book=Workbook();book.active['A1']='test';book.save(path);book.close()
            with patch.object(visual_fill,'_Package',wraps=visual_fill._Package) as package, \
                 patch('review_mapping._Package',side_effect=AssertionError('重复打开工作簿')):
                visual_fill.plan_visual_fill(path,{}, {'Sheet':{'heights':{1:30}}})
            self.assertEqual(package.call_count,1)


class CanvasPerformanceContracts(unittest.TestCase):
    def setUp(self):
        self.root=tk.Tk();self.addCleanup(self.root.destroy)
        app=SimpleNamespace(surface_bg='white',border_color='#cccccc',font_family='Arial',
                            text_fg='black',accent_fg='blue',muted_fg='gray')
        self.grid=VisualTable(self.root,app,on_resize=Mock());self.grid.pack()
        self.grid.load({'cells':{'A1':'text'},'merges':[],'max_row':100,'max_col':8})
        self.root.update();self.grid.draw()

    def test_event_burst_schedules_one_frame_and_renders_latest_selection(self):
        grid=self.grid
        for row in range(1,80):grid.select(row,1)
        frame=grid._frame
        grid.request_draw();self.assertEqual(grid._frame,frame)
        with patch.object(grid,'draw',wraps=grid.draw) as draw:
            self.root.after(40,self.root.quit);self.root.mainloop()
        self.assertEqual(draw.call_count,1)
        self.assertEqual(grid.selection,(1,79,1,79))

    def test_text_measurements_are_reused_and_cache_is_bounded(self):
        font=Mock();font.measure.side_effect=len
        grid=self.grid
        self.assertEqual(grid.visible_text('x'*10000,15,20,font,False),'x'*15)
        calls=font.measure.call_count
        self.assertLess(calls,20)
        grid.visible_text('x'*10000,15,20,font,False)
        self.assertEqual(font.measure.call_count,calls)
        for index in range(200):grid.visible_text(str(index),15,20,font,False)
        self.assertLessEqual(len(grid._text_cache),128)

    def test_resize_burst_rebuilds_once_and_commits_latest_size(self):
        grid=self.grid;scale=grid.zoom*grid.display_scale
        grid.gesture=('resize','column',1,0,160)
        with patch.object(grid,'rebuild_geometry',wraps=grid.rebuild_geometry) as rebuild:
            for x in range(1,100):grid.motion(SimpleNamespace(x=x,y=0))
            self.assertEqual(rebuild.call_count,0)
            grid.release(None)
            self.assertEqual(rebuild.call_count,1)
        grid.on_resize.assert_called_once_with('column',1,160+99/scale)
        self.assertAlmostEqual(grid.xs[1]-grid.xs[0],160*scale+99)

    def test_tall_merge_index_does_not_allocate_per_row(self):
        self.grid.load({'cells':{},'merges':['A2:A100000'],'max_row':100000,'max_col':1})
        self.assertLess(len(self.grid._merge_bands),1600)
        self.assertEqual(self.grid.bounds(1,1),(1,1,1,1))
        self.assertEqual(self.grid.bounds(100000,1),(1,2,1,100000))

    def test_prepared_workbook_is_reused_without_ui_thread_reload(self):
        import main
        from word_table_exporter import ExportedTable,ExportedTableCell,TableScanItem
        app=main.ReplaceSimpleApp(self.root,restore_session=False);app.show_workspace('extract')
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'template.xlsx'
            book=Workbook();book.active['A1']='原文';book.save(path);book.close()
            sheets=visual_fill.inspect_visual_workbook(path)
            table=ExportedTable([ExportedTableCell(1,1,1,1,'新文')],1,1)
            item=TableScanItem('source.docx','source.docx',1,'','','',1,1,'',table)
            with patch('visual_fill_dialog.inspect_visual_workbook',side_effect=AssertionError('界面重复读取')):
                dialog=VisualFillDialog(app.table_export_window,str(path),[item],sheets)
            try:
                self.assertIs(dialog.sheets,sheets)
                self.assertIs(dialog.original_sheets,sheets)
            finally:dialog.window.destroy()


if __name__=='__main__':unittest.main()
