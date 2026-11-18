"""Safety and interaction contracts for the native frontend (works with offscreen Qt)."""
import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from PySide6.QtCore import Qt, QSortFilterProxyModel, QThread, QPoint, QTimer
from PySide6.QtTest import QTest
from PySide6.QtGui import QFontDatabase
from PySide6.QtWidgets import QApplication, QDialog, QSplitter, QToolBar, QToolButton, QMessageBox
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from docx import Document

from qt_models import RowsModel, SheetModel, ScanModel
from qt_tasks import TaskController
from qt_main import MainWindow, STYLE, ProjectReview, format_result_summary
from qt_editor import VisualEditor, MappingDialog
from workbook_session import WorkbookSession
from word_table_exporter import ExportedTable, ExportedTableCell, TableScanItem

APP = QApplication.instance() or QApplication([])
# Windows offscreen plugin has no system-font discovery. Load local fonts for render QA.
if not QFontDatabase.families() and os.name == 'nt':
    for filename in ('msyh.ttc', 'msyhbd.ttc', 'simsun.ttc'):
        font = Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts' / filename
        if font.exists():
            QFontDatabase.addApplicationFont(str(font))
from qt_theme import apply_theme
apply_theme(APP)


def wait_task(controller, timeout=15):
    deadline = time.monotonic()+timeout
    while controller.busy and time.monotonic() < deadline:
        APP.processEvents()
        time.sleep(.005)
    APP.processEvents()
    if controller.busy:
        raise AssertionError('后台任务超时')


class NativeFrontendTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.path = self.root/'模板.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = '评审'
        sheet.append(['评审因素', '评分说明', '分值', '公式'])
        sheet.append(['技术方案', '原说明', 10, '=SUM(C2:C3)'])
        sheet.append(['服务方案', '原服务', 5])
        sheet['B2'].font = Font(name='宋体', size=12, bold=True)
        sheet.row_dimensions[2].height = 42
        book.create_sheet('原样保留')['A1'] = '第二张表'
        book.save(self.path)
        book.close()
        self.session = WorkbookSession.open(self.path)
        self.source = self.root/'来源.docx'
        doc = Document()
        table = doc.add_table(rows=3, cols=3)
        for cells, values in zip(table.rows, [('评审因素', '评分标准', '分值'), ('技术方案', '新说明', '10'), ('服务方案', '新服务', '5')]):
            for cell, text in zip(cells.cells, values):
                cell.text = text
        doc.add_paragraph('★必须提供支持')
        doc.save(self.source)

    def item(self):
        from word_scan import scan_word_content
        result = scan_word_content([str(self.source)], '★')
        return result[0][0][0], result[2]

    def library_window(self):
        from template_library import create_category, add_templates
        window = MainWindow(restore=False)
        self.addCleanup(window.deleteLater)
        self.addCleanup(window.close)
        window.library_root = str(self.root / '模板库')
        for category in ('公开招标', '竞争性磋商', '空分类'):
            create_category(window.library_root, category)
        second = self.root / '另一模板.xlsx'
        second.write_bytes(self.path.read_bytes())
        add_templates(window.library_root, '公开招标', [str(self.path), str(second)])
        add_templates(window.library_root, '竞争性磋商', [str(self.path)])
        window.show()
        window.tabs.setCurrentIndex(2)
        wait_task(window.tasks)
        window.category.setCurrentText('公开招标')
        APP.processEvents()
        return window

    def click_template_check(self, window, row):
        from PySide6.QtWidgets import QStyle, QStyleOptionViewItem
        index = window.templates.index(row, 0)
        option = QStyleOptionViewItem()
        option.initFrom(window.template_view)
        window.template_view.itemDelegate().initStyleOption(option, index)
        option.rect = window.template_view.visualRect(index)
        rect = window.template_view.style().subElementRect(
            QStyle.SE_ItemViewItemCheckIndicator, option, window.template_view)
        self.assertFalse(rect.isEmpty())
        QTest.mouseClick(window.template_view.viewport(), Qt.LeftButton, pos=rect.center())
        APP.processEvents()

    def test_template_mouse_checks_and_use_ignore_row_highlight(self):
        window = self.library_window()
        self.assertFalse(window.use_templates_button.isEnabled())
        self.click_template_check(window, 0)
        self.assertEqual(window.templates.checked, {0})
        window.template_view.selectRow(1)
        self.assertEqual(window.templates.checked, {0})
        self.assertIn('已勾选 1 / 2', window.template_selection.text())
        QTest.mouseClick(window.use_templates_button, Qt.LeftButton)
        self.assertEqual([row[1] for row in window.files.rows], [window.visible_templates[0].path])
        self.assertEqual(window.tabs.currentIndex(), 0)

    def test_template_mouse_header_clear_category_and_refresh(self):
        window = self.library_window()
        header = window.template_view.horizontalHeader()
        point = QPoint(header.sectionViewportPosition(0) + header.sectionSize(0)//2, header.height()//2)
        QTest.mouseClick(header.viewport(), Qt.LeftButton, pos=point)
        self.assertEqual(window.templates.checked, {0, 1})
        self.assertEqual(window.templates.headers[0], '选择（清空）')
        QTest.mouseClick(header.viewport(), Qt.LeftButton, pos=point)
        self.assertEqual(window.templates.checked, set())
        self.click_template_check(window, 0)
        self.click_template_check(window, 0)
        self.assertFalse(window.delete_templates_button.isEnabled())
        self.click_template_check(window, 1)
        window.category.setCurrentText('竞争性磋商')
        self.assertFalse(window.templates.checked)
        self.assertFalse(window.use_templates_button.isEnabled())
        window.category.setCurrentText('公开招标')
        self.assertFalse(window.templates.checked)
        self.click_template_check(window, 0)
        Path(window.visible_templates[0].path).unlink()
        window.refresh_library()
        wait_task(window.tasks)
        self.assertEqual(len(window.visible_templates), 1)
        self.assertFalse(window.templates.checked)
        window.category.setCurrentText('空分类')
        self.assertIn('暂无模板', window.template_selection.text())
        self.assertFalse(window.delete_templates_button.isEnabled())

    def test_template_directory_change_failure_cannot_reuse_old_checks(self):
        window = self.library_window()
        self.click_template_check(window, 0)
        old_paths = [x.path for x in window.visible_templates]
        with patch('qt_main.QFileDialog.getExistingDirectory', return_value=str(self.root / '新模板库')), \
             patch('template_library.list_categories', side_effect=OSError('读取失败')), \
             patch('qt_main.QMessageBox.warning'):
            window.change_library()
            self.assertFalse(window.templates.checked)
            self.assertFalse(window.visible_templates)
            wait_task(window.tasks)
        self.assertFalse(window.use_templates_button.isEnabled())
        self.assertFalse(window.delete_templates_button.isEnabled())
        self.assertEqual(window.checked_templates(), [])
        self.assertTrue(all(Path(path).exists() for path in old_paths))

    def test_template_delete_only_checked_current_category(self):
        window = self.library_window()
        self.click_template_check(window, 0)
        selected = window.visible_templates[0]
        untouched = window.visible_templates[1]
        window.template_view.selectRow(1)
        from template_library import list_templates
        hidden = [x for x in list_templates(window.library_root) if x.category == '竞争性磋商'][0]
        with patch('qt_main.QMessageBox.question', return_value=QMessageBox.Yes) as confirm:
            QTest.mouseClick(window.delete_templates_button, Qt.LeftButton)
            wait_task(window.tasks)
        self.assertIn(selected.name, confirm.call_args.args[2])
        self.assertFalse(Path(selected.path).exists())
        self.assertTrue(Path(untouched.path).exists())
        self.assertTrue(Path(hidden.path).exists())
        self.assertFalse(window.templates.checked)
        self.assertFalse(window.delete_templates_button.isEnabled())

    def test_horizontal_splitter_thin_and_mouse_draggable(self):
        window = MainWindow(restore=False)
        self.addCleanup(window.deleteLater)
        self.addCleanup(window.close)
        window.show()
        APP.processEvents()
        split = window.tabs.widget(0).findChild(QSplitter)
        self.assertEqual(split.orientation(), Qt.Vertical)
        self.assertGreaterEqual(split.handleWidth(), 6)
        self.assertLessEqual(split.handleWidth(), 10)
        handle = split.handle(1)
        self.assertGreaterEqual(handle.height(), 6)
        self.assertLessEqual(handle.height(), 10)
        before = split.sizes()
        start = handle.rect().center()
        QTest.mousePress(handle, Qt.LeftButton, pos=start)
        QTest.mouseMove(handle, start + QPoint(0, 45), 30)
        QTest.mouseRelease(handle, Qt.LeftButton, pos=start + QPoint(0, 45))
        APP.processEvents()
        self.assertGreater(split.sizes()[0], before[0] + 20)
        self.assertLess(split.sizes()[1], before[1] - 20)
        editor = VisualEditor(window, self.session, [], {})
        self.addCleanup(editor.deleteLater)
        editor_split = editor.findChild(QSplitter)
        self.assertEqual(editor_split.orientation(), Qt.Horizontal)
        self.assertEqual(editor_split.handleWidth(), split.handleWidth())

    def test_edits_and_undo_do_not_parse_or_read_workbook(self):
        with patch('pathlib.Path.read_bytes', side_effect=AssertionError('UI 线程 IO')), \
             patch('visual_fill._Package', side_effect=AssertionError('重复解析')):
            for n in range(110):
                self.session.edit('评审', {'B2': str(n)})
            self.assertEqual(len(self.session.history), 100)
            self.session.undo()
            self.assertEqual(self.session.pending['评审']['B2'], '108')
            self.session.redo()
            self.assertEqual(self.session.pending['评审']['B2'], '109')
            self.assertIs(self.session.history[-1][2], self.session.sheets)

    def test_formula_protection_and_transaction_rollback(self):
        with self.assertRaisesRegex(ValueError, '公式'):
            self.session.edit('评审', {'B2': '不应提交', 'D2': '破坏'})
        self.assertEqual(self.session.pending, {})
        self.assertEqual(self.session.history, [])
        self.session.sheets['评审']['protected'] = True
        with self.assertRaisesRegex(ValueError, '保护'):
            self.session.edit('评审', {'B2': '不能改'})

    def test_save_preserves_formula_styles_other_sheet_and_original(self):
        original = self.path.read_bytes()
        self.session.edit('评审', {'B2': '新说明\n第二行', 'C3': '=不是公式'})
        output = self.root/'输出.xlsx'
        self.session.save(output)
        self.assertEqual(self.path.read_bytes(), original)
        book = load_workbook(output)
        self.assertEqual(book['评审']['B2'].value, '新说明\n第二行')
        self.assertEqual(book['评审']['B2'].font.name, '宋体')
        self.assertTrue(book['评审']['B2'].font.bold)
        self.assertEqual(book['评审'].row_dimensions[2].height, 42)
        self.assertEqual(book['评审']['D2'].value, '=SUM(C2:C3)')
        self.assertEqual(book['评审']['C3'].data_type, 's')
        book.close()
        with ZipFile(self.path) as before, ZipFile(output) as after:
            self.assertEqual(before.read('xl/worksheets/sheet2.xml'), after.read('xl/worksheets/sheet2.xml'))
        with self.assertRaisesRegex(ValueError, '另存'):
            self.session.save(output)

    def test_external_change_refuses_save(self):
        self.session.edit('评审', {'B2': '新值'})
        self.path.write_bytes(self.path.read_bytes()+b'changed')
        with self.assertRaisesRegex(ValueError, '变化'):
            self.session.save(self.root/'禁止输出.xlsx')
        self.assertFalse((self.root/'禁止输出.xlsx').exists())

    def test_structure_versions_undo_redo_and_save_without_pending(self):
        self.session.edit('评审', {'B2': '待写入'})
        version, intermediate = self.root/'结构.xlsx', self.root/'中间.xlsx'
        self.session.structure(version, '评审', 'insert_above', (1, 2, 1, 2), 1, intermediate)
        self.assertEqual(self.session.sheets['评审']['cells']['B3'], '待写入')
        self.assertEqual(self.session.sheets['评审']['formulas']['D3'], 'SUM(C3:C4)')
        self.session.undo()
        self.assertEqual(self.session.path, str(self.path))
        self.assertEqual(self.session.pending['评审']['B2'], '待写入')
        self.session.redo()
        output = self.root/'仅结构另存.xlsx'
        self.session.save(output)
        self.assertTrue(output.exists())

    def test_model_edits_and_restricted_columns(self):
        model = SheetModel(self.session.sheets['评审'], self.session, '评审')
        self.assertTrue(model.setData(model.index(1, 1), '模型写入'))
        self.assertEqual(model.data(model.index(1, 1)), '模型写入')
        self.assertFalse(model.flags(model.index(1, 3)) & Qt.ItemIsEditable)
        rows = RowsModel(['固定', '值'], [['名字', '内容']], editable=True)
        rows.editable_columns = {1}
        self.assertFalse(rows.setData(rows.index(0, 0), '不允许'))
        self.assertTrue(rows.setData(rows.index(0, 1), '允许'))

    def test_proxy_filter_preserves_checks_without_model_reset(self):
        item, _ = self.item()
        model = ScanModel()
        model.load([item], [])
        model.setData(model.index(0, 0), Qt.Checked, Qt.CheckStateRole)
        proxy = QSortFilterProxyModel()
        proxy.setSourceModel(model)
        proxy.setFilterKeyColumn(-1)
        resets = []
        model.modelReset.connect(lambda: resets.append(True))
        proxy.setFilterFixedString('没有的关键词')
        self.assertEqual(proxy.rowCount(), 0)
        proxy.setFilterFixedString('评分')
        self.assertEqual(proxy.rowCount(), 1)
        self.assertEqual(model.checked, {0})
        self.assertEqual(resets, [])

    def test_worker_thread_and_coalesced_progress_completion(self):
        controller = TaskController()
        results, errors, threads, progresses = [], [], [], []
        controller.error.connect(errors.append)
        controller.progress.connect(progresses.append)
        def work(publish):
            threads.append(QThread.currentThread())
            for n in range(10000):
                publish(n, 9999, '进度')
            return '完成'
        controller.submit(work, results.append)
        wait_task(controller)
        self.assertEqual(results, ['完成'])
        self.assertFalse(errors)
        self.assertNotEqual(threads[0], APP.thread())
        self.assertEqual(progresses[-1], (9999, 9999, '进度'))
        self.assertLess(len(progresses), 100)
        controller.submit(lambda p: (_ for _ in ()).throw(ValueError('预期失败')), results.append)
        wait_task(controller)
        self.assertIn('预期失败', errors[-1])

    def test_visual_transfer_mapping_and_save_source_change_guard(self):
        item, stamps = self.item()
        window = MainWindow(restore=False)
        editor = VisualEditor(window, self.session, [item], stamps)
        editor.transfer_values((2, 2, 2, 2), 2, 2)
        self.assertEqual(self.session.pending['评审']['B2'], '新说明')
        editor.undo()
        self.assertEqual(self.session.pending, {})
        mapping = MappingDialog(editor, item, self.session, '评审')
        mapping.build()
        wait_task(mapping.tasks)
        self.assertIsNotNone(mapping.plan)
        self.assertEqual(mapping.plan['errors'], [])
        self.assertTrue(mapping.apply_button.isEnabled())
        mapping.apply()
        self.assertEqual(self.session.pending['评审']['B3'], '新服务')
        mapping.model.setData(mapping.model.index(0, 1), '0')
        self.assertIsNone(mapping.plan)
        self.source.write_bytes(self.source.read_bytes()+b'change')
        with self.assertRaisesRegex(ValueError, '变化'):
            editor.verify_sources()
        editor.scratch.cleanup()
        editor.hide()
        editor.deleteLater()
        window.hide()
        window.deleteLater()
        APP.processEvents()

    def test_workbench_batch_replace_writes_real_docx_xlsx_pptx(self):
        from pptx import Presentation
        doc = Document(); doc.add_paragraph('[项目名称]'); doc.save(self.root/'待替换.docx')
        book = Workbook(); book.active['A1'] = '[项目名称]'; book.save(self.root/'待替换.xlsx'); book.close()
        deck = Presentation(); slide = deck.slides.add_slide(deck.slide_layouts[6])
        from pptx.util import Inches
        slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text = '[项目名称]'
        deck.save(self.root/'待替换.pptx')
        output = self.root/'替换输出'; output.mkdir()
        window = MainWindow(restore=False)
        window.append_files([str(self.root/f'待替换{ext}') for ext in ('.docx', '.xlsx', '.pptx')])
        window.rules.replace([['[项目名称]', '已核对项目']])
        window.output.setText(str(output))
        window.replace()
        wait_task(window.tasks)
        self.assertEqual(Document(output/'待替换.docx').paragraphs[0].text, '已核对项目')
        book = load_workbook(output/'待替换.xlsx')
        self.assertEqual(book.active['A1'].value, '已核对项目'); book.close()
        self.assertEqual(Presentation(output/'待替换.pptx').slides[0].shapes[0].text, '已核对项目')
        self.assertIn('输出', window.log.toPlainText())
        window.close(); window.deleteLater()

    def test_workbench_scan_export_and_template_library(self):
        from template_library import create_category, add_templates
        window = MainWindow(restore=False)
        errors = []
        window.tasks.error.disconnect(window.error)
        window.tasks.error.connect(errors.append)
        window.scan_files.append([str(self.source)])
        window.symbols.setText('★')
        window.scan()
        wait_task(window.tasks)
        self.assertEqual(len(window.scans.items), 2)
        window.check_visible(True)
        output = self.root/'提取输出'; output.mkdir()
        window.output.setText(str(output))
        window.export_scan()
        wait_task(window.tasks)
        self.assertEqual(len(list(output.glob('*.xlsx'))), 2)
        window.library_root = str(self.root/'模板库')
        create_category(window.library_root, '公开招标')
        add_templates(window.library_root, '公开招标', [str(self.path)])
        window.refresh_library()
        wait_task(window.tasks)
        self.assertEqual(len(window.templates.rows), 1)
        window.templates.setData(window.templates.index(0, 0), Qt.Checked, Qt.CheckStateRole)
        window.use_templates()
        self.assertEqual(window.files.rows[0][0], '模板.xlsx')
        self.assertFalse(errors)
        window.close(); window.deleteLater()

    def test_project_review_fields_are_fixed_values_are_editable(self):
        details = [{'field': '项目名称', 'placeholder': '[项目名称]', 'value': '', 'status': '缺失', 'candidates': []}]
        review = ProjectReview(None, details)
        self.assertFalse(review.model.setData(review.model.index(0, 0), '错误字段'))
        review.model.setData(review.model.index(0, 1), '某项目')
        self.assertEqual(review.rules(), [['[项目名称]', '某项目']])
        review.deleteLater()


    def test_layout_only_save_undo_and_structure_preserve_dimensions(self):
        from worksheet_layout import auto_heights
        self.session.edit('评审', {'B2': '很多文字\n第二行\n第三行'})
        heights, wrapped = auto_heights(self.session.sheets['评审'], self.session.pending['评审'], [2])
        self.session.layout('评审', heights=heights, widths={2: 38}, wrap_cells=wrapped)
        self.assertEqual(self.session.view_info('评审')['column_pixels'][2], 271)
        self.session.undo()
        self.assertEqual(self.session.layouts, {})
        self.session.redo()
        output = self.root/'排版.xlsx'
        self.session.save(output)
        book = load_workbook(output)
        self.assertEqual(book['评审'].column_dimensions['B'].width, 38)
        self.assertEqual(book['评审'].row_dimensions[2].height, heights[2])
        self.assertTrue(book['评审']['B2'].alignment.wrap_text)
        self.assertEqual(book['评审']['D2'].value, '=SUM(C2:C3)')
        book.close()
        self.session.structure(self.root/'排版插行.xlsx', '评审', 'insert_above', (1, 2, 1, 2), 1, self.root/'排版中间.xlsx')
        self.assertEqual(self.session.sheets['评审']['cells']['B3'], '很多文字\n第二行\n第三行')
        self.session.undo()
        self.assertEqual(self.session.layouts['评审']['widths'][2], 38)
        clean = WorkbookSession.open(self.path)
        clean.layout('评审', widths={2: 28})
        clean.save(self.root/'仅列宽.xlsx')
        with self.assertRaises(ValueError):
            clean.layout('评审', heights={2: float('nan')})

    def test_batch_preview_failed_retry_and_csv_report(self):
        import csv
        output = self.root/'报告输出'; output.mkdir()
        broken = self.root/'损坏.docx'; broken.write_bytes(b'not a zip')
        window = MainWindow(restore=False)
        window.append_files([str(self.source), str(broken)])
        window.rules.replace([['新说明', '替换说明']])
        window.output.setText(str(output))
        previews = []
        window.show_output_rows = previews.append
        window.preview_outputs()
        wait_task(window.tasks)
        self.assertEqual(len(previews[0]), 2)
        self.assertFalse(list(output.iterdir()))
        window.replace(); wait_task(window.tasks)
        self.assertEqual([x['status'] for x in window.last_batch['明细']], ['成功', '失败'])
        original_output = Path(window.last_batch['明细'][0]['output_path'])
        before = original_output.read_bytes()
        doc = Document(); doc.add_paragraph('新说明'); doc.save(broken)
        window.rules.replace([['新说明', '不应采用的新规则']])
        window.retry_failed(); wait_task(window.tasks)
        self.assertEqual(Document(output/'损坏.docx').paragraphs[0].text, '替换说明')
        self.assertEqual(original_output.read_bytes(), before)
        self.assertEqual(len(list(output.glob('*.docx'))), 2)
        self.assertTrue(all(x['status'] == '成功' for x in window.last_batch['明细']))
        report = self.root/'报告.csv'
        with patch('qt_extras.QFileDialog.getSaveFileName', return_value=(str(report), 'CSV')):
            window.export_report(); wait_task(window.tasks)
        with report.open(encoding='utf-8-sig', newline='') as file:
            rows = list(csv.reader(file))
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0], ['来源', '输出', '状态', '命中数', '错误'])
        window.close(); window.deleteLater()

    def test_result_details_are_readable_with_raw_diagnostics_secondary(self):
        output = self.root/'详情输出'; output.mkdir()
        broken = self.root/'损坏.docx'; broken.write_bytes(b'not a zip')
        window = MainWindow(restore=False)
        window.append_files([str(self.source), str(broken)])
        window.rules.replace([['新说明', '替换说明']])
        window.output.setText(str(output))
        window.replace(); wait_task(window.tasks)
        text = window.log.toPlainText()
        self.assertIn('批量替换完成', text)
        self.assertIn('共 2 个文件（成功 1、失败 1）', text)
        self.assertIn('来源.docx — 成功：命中', text)
        self.assertIn('输出', text)
        self.assertIn('损坏.docx — 失败：', text)
        self.assertNotIn("{'source_path'", text)
        self.assertNotIn("'明细'", text)
        self.assertIn("'明细'", window.raw_log.toPlainText())
        self.assertFalse(window.raw_log.isVisibleTo(window))
        self.assertTrue(window.details_panel.isVisibleTo(window))
        window.copy_details()
        self.assertIn('批量替换完成', APP.clipboard().text())
        self.assertNotIn("'明细'", APP.clipboard().text())
        window.raw_toggle.setChecked(True)
        self.assertTrue(window.raw_log.isVisibleTo(window))
        window.copy_details()
        copied = APP.clipboard().text()
        self.assertIn('批量替换完成', copied)
        self.assertIn("'明细'", copied)
        window.raw_toggle.setChecked(False)
        window.close(); window.deleteLater()

    def test_format_result_summary_covers_scan_export_and_unknown_shapes(self):
        value = ({'来源.docx': {'tables': 2, 'output_path': 'C:/x/表格.xlsx'}}, {}, None, {}, {}, None)
        summary = format_result_summary(value)
        self.assertIn('导出完成：表格 1 个来源、条款 0 个来源', summary)
        self.assertIn('来源.docx → 表格.xlsx（2 张）', summary)
        self.assertIn('导出完成：表格 0 个来源、条款 1 个来源',
                      format_result_summary(({}, {}, None, {'来源.docx': {'count': 3, 'output_path': 'C:/x/条款.xlsx'}}, {}, None)))
        self.assertIn('失败：无法导出', format_result_summary(({}, {}, '无法导出', {}, {}, None)))
        self.assertEqual(format_result_summary({'foo': 'bar'}), '')
        self.assertIn('没有可记录的文件明细', format_result_summary({'结果': {}, '错误': None, '明细': []}))

    def test_clause_selection_and_combined_filters_export_only_checked_clause(self):
        from word_scan import scan_word_content
        from qt_models import ScanFilter
        doc = Document(self.source); doc.add_paragraph('★第二条单独选择'); doc.save(self.source)
        tables, sections, stamps = scan_word_content([str(self.source)], '★')
        window = MainWindow(restore=False)
        window.scans.load(tables[0], sections[0]); window.stamps = stamps; window.scan_symbols = '★'
        clauses = [i for i, (kind, _) in enumerate(window.scans.items) if kind == 'symbol']
        self.assertEqual(len(clauses), 2)
        row = clauses[-1]
        window.scans.setData(window.scans.index(row, 0), Qt.Checked, Qt.CheckStateRole)
        window.proxy.select_values(['symbol'], [window.scans.items[row][1].section])
        self.assertEqual(window.proxy.rowCount(), 2)
        window.proxy.setFilterFixedString('第二条')
        self.assertEqual(window.proxy.rowCount(), 1)
        window.proxy.select_values(['table'])
        self.assertEqual(window.proxy.rowCount(), 0)
        self.assertEqual(window.scans.checked, {row})
        output = self.root/'单条导出'; output.mkdir(); window.output.setText(str(output))
        window.export_scan(); wait_task(window.tasks)
        book = load_workbook(next(output.glob('*.xlsx')))
        text = '\n'.join(str(c.value) for sheet in book for row in sheet for c in row if c.value)
        self.assertIn('第二条单独选择', text)
        self.assertNotIn('必须提供支持', text)
        book.close(); window.close(); window.deleteLater()

    def test_recipe_roundtrip_validation_and_zoom_do_not_touch_template(self):
        from qt_editor import SheetPreview
        item, stamps = self.item()
        editor = VisualEditor(None, self.session, [item], stamps)
        editor.transfer_values((2, 2, 2, 2), 2, 2)
        recipe = dict(editor.last_recipe)
        editor.undo()
        editor.apply_recipe(recipe)
        self.assertEqual(self.session.pending['评审']['B2'], '新说明')
        before = self.session.snapshot()
        with self.assertRaises(ValueError):
            editor.apply_recipe({**recipe, 'bounds': [1, 1, 9999, 9999]})
        self.assertEqual(self.session.snapshot(), before)
        path = self.root/'配方.json'
        with patch('qt_editor.QFileDialog.getSaveFileName', return_value=(str(path), 'JSON')):
            editor.save_recipe(); wait_task(editor.tasks)
        editor.undo()
        with patch('qt_editor.QFileDialog.getOpenFileName', return_value=(str(path), 'JSON')):
            editor.load_recipe(); wait_task(editor.tasks)
        self.assertEqual(self.session.pending['评审']['B2'], '新说明')
        original = self.path.read_bytes()
        preview = SheetPreview(None, self.session.view_info('评审'))
        width = preview.view.columnWidth(1)
        preview.zoom.setValue(150)
        self.assertGreater(preview.view.columnWidth(1), width)
        self.assertEqual(self.path.read_bytes(), original)
        preview.close(); preview.deleteLater()
        editor.scratch.cleanup(); editor.hide(); editor.deleteLater()
        APP.processEvents()


    def layout_editor(self, items=None):
        item, stamps = self.item()
        # Display-only stress case: no filesystem path limit is involved.
        item.filename = '超长中文采购项目招标文件技术参数评分办法与需求核对终稿' * 8 + '.docx'
        editor = VisualEditor(None, self.session, [item] if items is None else items, stamps)

        def cleanup():
            wait_task(editor.tasks)
            editor.scratch.cleanup()
            editor.hide()
            editor.deleteLater()
            APP.processEvents()
        self.addCleanup(cleanup)
        return editor

    def test_editor_long_names_keep_equal_panels_at_supported_sizes(self):
        editor = self.layout_editor()
        # Exercise the target selector's long-text behavior as well.
        long_sheet = '中文目标工作表名称' * 3
        self.session.sheets[long_sheet] = self.session.sheets['评审']
        editor.sheet.addItem(long_sheet)
        for width, height in [(1024, 675), (860, 560), (1260, 800), (1024, 675)]:
            with self.subTest(size=(width, height)):
                editor.resize(width, height)
                editor.show()
                APP.processEvents()
                self.assertEqual(editor.size().toTuple(), (width, height))
                split = editor.splitter
                left, right = split.sizes()
                self.assertLessEqual(abs(left - right), 2)
                self.assertGreaterEqual(min(left, right), 260)
                self.assertGreaterEqual(editor.source_view.viewport().width(), 230)
                self.assertGreaterEqual(editor.target.viewport().width(), 230)
                self.assertLessEqual(editor.source_choice.width(), left)
                self.assertLessEqual(editor.sheet.width(), right)
                for a, b in [(editor.source_choice, editor.sheet),
                             (editor.source_view, editor.target)]:
                    self.assertEqual(a.mapTo(editor, QPoint()).y(), b.mapTo(editor, QPoint()).y())
                    self.assertEqual(a.height(), b.height())
                for toolbar in editor.findChildren(QToolBar):
                    self.assertTrue(editor.body.rect().contains(toolbar.geometry()))
                    # Every command is either visible or available from Qt's overflow menu.
                    extension = toolbar.findChild(QToolButton, 'qt_toolbar_ext_button')
                    for action in toolbar.actions():
                        widget = toolbar.widgetForAction(action)
                        if not widget.isVisible():
                            self.assertTrue(extension.isVisible())
                        else:
                            self.assertTrue(toolbar.rect().contains(widget.geometry()))
        combo = editor.source_choice
        self.assertEqual(combo.toolTip(), combo.currentText())
        self.assertEqual(combo.itemData(0, Qt.ToolTipRole), combo.currentText())
        self.assertGreater(combo.fontMetrics().horizontalAdvance(combo.currentText()), combo.width() * 3)
        # Switch the actual selectors via keyboard, keeping their full text and models intact.
        editor.sheet.setFocus()
        QTest.keyClick(editor.sheet, Qt.Key_End)
        APP.processEvents()
        self.assertEqual(editor.sheet.currentText(), long_sheet)
        self.assertEqual(editor.sheet.toolTip(), long_sheet)
        from copy import deepcopy
        editor.items.append(deepcopy(editor.items[0]))
        combo.addItem('另一份中文来源.docx')
        combo.showPopup()
        APP.processEvents()
        self.assertTrue(combo.view().isVisible())
        QTest.keyClick(combo.view(), Qt.Key_Down)
        QTest.keyClick(combo.view(), Qt.Key_Return)
        APP.processEvents()
        self.assertEqual(combo.currentText(), '另一份中文来源.docx')
        self.assertEqual(combo.toolTip(), '另一份中文来源.docx')

    def test_editor_splitter_real_drag_minima_and_resize(self):
        editor = self.layout_editor()
        editor.resize(1024, 675)
        editor.show()
        APP.processEvents()
        split = editor.splitter

        def drag(delta):
            handle = split.handle(1)
            center = handle.rect().center()
            QTest.mousePress(handle, Qt.LeftButton, Qt.NoModifier, center)
            QTest.mouseMove(handle, center + QPoint(delta, 0), 20)
            QTest.mouseRelease(handle, Qt.LeftButton, Qt.NoModifier,
                               handle.rect().center())
            APP.processEvents()

        before = split.sizes()
        drag(100)
        after = split.sizes()
        self.assertGreater(after[0], before[0] + 70)
        self.assertLess(after[1], before[1] - 70)
        # Resizing preserves a user's adjustment; it must not reset to 50/50.
        ratio = after[0] / sum(after)
        editor.resize(1260, 800)
        APP.processEvents()
        self.assertAlmostEqual(split.sizes()[0] / sum(split.sizes()), ratio, delta=.015)
        drag(-2000)
        self.assertGreaterEqual(split.sizes()[0], 260)
        self.assertLessEqual(split.sizes()[0], 262)
        drag(4000)
        self.assertGreaterEqual(split.sizes()[1], 260)
        self.assertLessEqual(split.sizes()[1], 262)
        self.assertTrue(editor.source_view.isVisible())
        self.assertTrue(editor.target.isVisible())

    def test_editor_empty_source_and_shortcuts_survive_layout_change(self):
        editor = self.layout_editor(items=[])
        editor.resize(860, 560)
        editor.show()
        editor.activateWindow()
        APP.processEvents()
        self.assertLessEqual(abs(editor.splitter.sizes()[0] - editor.splitter.sizes()[1]), 2)
        self.assertEqual(editor.source_choice.count(), 0)
        self.session.edit('评审', {'B2': '快捷键核验'})
        editor.target_changed()
        editor.target.setFocus()
        QTest.keyClick(editor.target, Qt.Key_Z, Qt.ControlModifier)
        APP.processEvents()
        self.assertEqual(self.session.pending, {})
        QTest.keyClick(editor.target, Qt.Key_Y, Qt.ControlModifier)
        APP.processEvents()
        self.assertEqual(self.session.pending['评审']['B2'], '快捷键核验')
        with patch('qt_editor.QFileDialog.getSaveFileName', return_value=('', '')) as save:
            QTest.keyClick(editor.target, Qt.Key_S, Qt.ControlModifier | Qt.ShiftModifier)
            APP.processEvents()
            save.assert_called_once()
        bars = editor.findChildren(QToolBar)
        self.assertEqual(len(bars[0].actions()), 6)
        self.assertEqual(len(bars[1].actions()), 8)
        self.assertEqual(len(bars[1].actions()[-1].menu().actions()), 7)

    def test_editor_toolbar_overflow_and_layout_menu_remain_clickable(self):
        editor = self.layout_editor()
        editor.resize(860, 560)
        editor.show()
        APP.processEvents()
        structures = editor.findChildren(QToolBar)[1]
        menu_action = structures.actions()[-1]
        observed = []

        def choose(text):
            # Native popup menus enter a nested event loop during mousePress.
            # Choose by mouse from that loop, with a timeout to avoid hanging on failure.
            popup = APP.activePopupWidget()
            if popup is not None:
                action = next((a for a in popup.actions() if a.text() == text), None)
                if action is not None:
                    observed.append(text)
                    QTest.mouseClick(popup, Qt.LeftButton,
                                     pos=popup.actionGeometry(action).center())
                popup.hide()

        timeout = QTimer(editor)
        timeout.setSingleShot(True)
        timeout.timeout.connect(lambda: APP.activePopupWidget().hide()
                                if APP.activePopupWidget() else None)
        with patch.object(editor, 'resize_cells') as resize_cells:
            QTimer.singleShot(30, lambda: choose('设置行高'))
            timeout.start(1000)
            QTest.mouseClick(structures.widgetForAction(menu_action), Qt.LeftButton)
            QTest.qWait(60)
            timeout.stop()
            APP.processEvents()
            self.assertIn('设置行高', observed)
            resize_cells.assert_called_once_with('height')
        # Force real overflow even on machines with smaller fonts.
        structures.setMaximumWidth(350)
        APP.processEvents()
        extension = structures.findChild(QToolButton, 'qt_toolbar_ext_button')
        self.assertTrue(extension.isVisible())
        with patch.object(editor, 'structure') as structure:
            QTimer.singleShot(30, lambda: choose('合并选区'))
            timeout.start(1000)
            QTest.mouseClick(extension, Qt.LeftButton)
            QTest.qWait(60)
            timeout.stop()
            APP.processEvents()
            self.assertIn('合并选区', observed)
            structure.assert_called_once_with('merge')
        self.assertEqual(editor.size().toTuple(), (860, 560))

    def test_main_splitters_and_long_category_do_not_squeeze_panels(self):
        window = MainWindow(restore=False)
        def cleanup():
            wait_task(window.tasks)
            window.close()
            window.deleteLater()
            APP.processEvents()
        self.addCleanup(cleanup)
        window.resize(1024, 675)
        window.show()
        window.tabs.setCurrentIndex(1)
        APP.processEvents()
        split = window.scan_view.parentWidget()
        self.assertIsInstance(split, QSplitter)
        self.assertEqual(split.orientation(), Qt.Horizontal)
        self.assertLessEqual(abs(split.sizes()[0] - split.sizes()[1]), 2)
        self.assertGreaterEqual(min(split.sizes()), 260)
        self.assertFalse(split.childrenCollapsible())
        window.tabs.setCurrentIndex(2)
        wait_task(window.tasks)
        window.category.addItem('中文模板库分类名称超长文本' * 15)
        window.category.setCurrentIndex(window.category.count() - 1)
        APP.processEvents()
        self.assertEqual(window.size().toTuple(), (1024, 675))
        self.assertTrue(window.category.parentWidget().rect().contains(window.category.geometry()))
        self.assertEqual(window.category.toolTip(), window.category.currentText())
        vertical = window.file_view.parentWidget().parentWidget()
        self.assertEqual(vertical.orientation(), Qt.Vertical)
        self.assertFalse(vertical.childrenCollapsible())
        review = ProjectReview(window, [])
        review.resize(860, 560)
        review.show()
        APP.processEvents()
        split = review.findChild(QSplitter)
        self.assertLessEqual(abs(split.sizes()[0] - split.sizes()[1]), 2)
        self.assertGreaterEqual(min(split.sizes()), 260)
        self.assertFalse(split.childrenCollapsible())
        review.close()

    def test_editor_layout_actions_and_save_use_current_selection(self):
        item, stamps = self.item()
        editor = VisualEditor(None, self.session, [item], stamps)
        editor.target.setCurrentIndex(editor.target.model().index(1, 1))
        with patch('qt_editor.QInputDialog.getDouble', return_value=(63.0, True)):
            editor.resize_cells('height')
        self.assertEqual(self.session.layouts['评审']['heights'], {2: 63.0})
        editor.target.setCurrentIndex(editor.target.model().index(1, 1))
        with patch('qt_editor.QInputDialog.getDouble', return_value=(35.0, True)):
            editor.resize_cells('width')
        self.session.edit('评审', {'B2': '第一行\n第二行\n第三行'})
        editor.target.setCurrentIndex(editor.target.model().index(1, 1))
        editor.auto_height(); wait_task(editor.tasks)
        self.assertIn('2 项排版', editor.status.text())
        self.assertIn('B2', self.session.layouts['评审']['wrap_cells'])
        output = self.root/'编辑器排版另存.xlsx'
        with patch('qt_editor.QFileDialog.getSaveFileName', return_value=(str(output), 'Excel')):
            editor.save(); wait_task(editor.tasks)
        book = load_workbook(output)
        self.assertEqual(book['评审'].column_dimensions['B'].width, 35)
        self.assertTrue(book['评审']['B2'].alignment.wrap_text)
        book.close()
        editor.scratch.cleanup(); editor.hide(); editor.deleteLater(); APP.processEvents()

    def test_recipe_rejects_changed_target_structure_without_partial_edits(self):
        item, stamps = self.item()
        editor = VisualEditor(None, self.session, [item], stamps)
        editor.transfer_values((2, 2, 2, 2), 2, 2)
        recipe = dict(editor.last_recipe)
        editor.undo()
        self.session.sheets['评审']['cells']['A1'] = '另一种表头'
        with self.assertRaisesRegex(ValueError, '表头'):
            editor.apply_recipe(recipe)
        self.assertEqual(self.session.pending, {})
        self.assertEqual(self.session.history, [])
        editor.scratch.cleanup(); editor.hide(); editor.deleteLater(); APP.processEvents()


if __name__ == '__main__':
    unittest.main()
