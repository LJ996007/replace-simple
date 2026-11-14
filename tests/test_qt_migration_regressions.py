"""Migration contracts driven through Qt events with isolated Office samples."""
import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
import json
import time
from pathlib import Path
from unittest.mock import patch

import pytest
from docx import Document
from openpyxl import Workbook
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QMessageBox
from qt_main import MainWindow
from qt_editor import MappingDialog, VisualEditor
from qt_theme import apply_theme
from workbook_session import WorkbookSession
from word_scan import scan_word_content

APP = QApplication.instance() or QApplication([])
apply_theme(APP)


def wait(controller):
    deadline = time.monotonic() + 15
    while controller.busy and time.monotonic() < deadline:
        QTest.qWait(10)
    assert not controller.busy
    APP.processEvents()


@pytest.fixture
def window(tmp_path, monkeypatch):
    monkeypatch.setenv('APPDATA', str(tmp_path))
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    host = MainWindow(restore=False)
    host.library_root = str(tmp_path / 'library')
    host.library_path.setText(host.library_root)
    host.show()
    APP.processEvents()
    yield host
    wait(host.tasks)
    host.close()
    host.deleteLater()
    APP.processEvents()


def source(path):
    doc = Document()
    table = doc.add_table(rows=2, cols=3)
    for cells, values in zip(table.rows, [('评审因素', '评分标准', '分值'), ('服务', '原始扫描内容', '10')]):
        for cell, value in zip(cells.cells, values):
            cell.text = value
    doc.add_paragraph('★提供服务')
    doc.save(path)
    return scan_word_content([str(path)], '★')


def test_failed_rule_import_preserves_files_and_rules(window, tmp_path):
    path = tmp_path / 'broken.xlsx'
    path.write_bytes(b'not a workbook')
    window.append_files([str(path)])
    window.rules.replace([['old', 'new']])
    with patch('qt_main.QFileDialog.getOpenFileName', return_value=(str(path), '')), patch('qt_main.QMessageBox.warning'):
        window.import_rules()
        wait(window.tasks)
    assert window.rules.rows == [['old', 'new']]
    assert [row[1] for row in window.files.rows] == [str(path)]


def test_scan_reset_and_filtered_selection_clear_preview(window, tmp_path):
    path = tmp_path / 'source.docx'
    result = source(path)
    window.scans.load(result[0][0], result[1][0])
    window.tabs.setCurrentIndex(1)
    window.scan_view.setCurrentIndex(window.proxy.index(0, 1))
    assert '原始扫描内容' in window.preview_text.toPlainText()
    window.proxy.setFilterFixedString('no matching rows')
    assert window.preview_text.toPlainText() == ''
    assert window.scan_preview.model().info['cells'] == {}
    window.proxy.setFilterFixedString('')
    window.scan_view.setCurrentIndex(window.proxy.index(0, 1))
    window.scans.load([], [])
    assert window.preview_text.toPlainText() == ''
    assert window.scan_preview.model().info['cells'] == {}


def test_rule_clipboard_delete_and_undo_events(window):
    view = window.rule_view
    window.rules.replace([['old', 'new']])
    view.setCurrentIndex(window.rules.index(0, 0))
    view.setFocus()
    APP.clipboard().setText('项目名称\t新项目\n项目编号\t2026-01')
    if APP.clipboard().text() != '项目名称\t新项目\n项目编号\t2026-01':
        pytest.skip('独立 Windows 桌面的系统剪贴板不可用；此用例另用 offscreen Qt 剪贴板执行')
    QTest.keyClick(view, Qt.Key_V, Qt.ControlModifier)
    assert window.rules.rows == [['项目名称', '新项目'], ['项目编号', '2026-01']]
    QTest.keyClick(view, Qt.Key_Z, Qt.ControlModifier)
    assert window.rules.rows == [['old', 'new']]
    QTest.keyClick(view, Qt.Key_Y, Qt.ControlModifier)
    assert window.rules.rows == [['项目名称', '新项目'], ['项目编号', '2026-01']]
    view.selectAll()
    QTest.keyClick(view, Qt.Key_C, Qt.ControlModifier)
    assert '项目编号\t2026-01' in APP.clipboard().text()
    QTest.keyClick(view, Qt.Key_Delete)
    assert window.rules.rows == [['', ''], ['', '']]
    QTest.keyClick(view, Qt.Key_Z, Qt.ControlModifier)
    assert window.rules.rows[1] == ['项目编号', '2026-01']


def test_rule_multiline_roundtrip_and_rejected_wide_paste(window):
    view = window.rule_view
    values = [['占位符', '第一行\n第二行\t内容'], ['其他', '=原文']]
    window.rules.replace(values)
    view.selectAll()
    view.copy()
    if not APP.clipboard().text():
        pytest.skip('独立 Windows 桌面的系统剪贴板不可用')
    window.rules.replace([])
    view.paste()
    assert window.rules.rows == values
    APP.clipboard().setText('a\tb\tc')
    with patch('qt_main.QMessageBox.warning') as warning:
        view.paste()
        assert warning.called
    assert window.rules.rows == values


def test_rule_import_success_and_cancel(window, tmp_path):
    path = tmp_path / 'rules.xlsm'
    book = Workbook()
    book.active.append(['before', 'after'])
    book.save(path)
    book.close()
    window.append_files([str(path)])
    with patch('qt_main.QFileDialog.getOpenFileName', return_value=('', '')):
        window.import_rules()
    assert window.files.rowCount() == 1
    with patch('qt_main.QFileDialog.getOpenFileName', return_value=(str(path), '')) as chooser:
        window.import_rules()
        wait(window.tasks)
    assert '*.xlsm' in chooser.call_args.args[-1]
    assert window.rules.rows == [['before', 'after']]
    assert window.files.rowCount() == 0
    assert window.rules.setData(window.rules.index(0, 1), 'edited')
    window.rules.undo()
    assert window.rules.rows == [['before', 'after']]


def test_scan_failure_is_visible_without_opening_log(window, tmp_path):
    window.scan_files.append([str(tmp_path / 'missing.docx')])
    window.scan()
    wait(window.tasks)
    assert '失败' in window.status.text()
    assert window.log.isVisible()


def test_library_completion_does_not_leave_busy_status(window):
    window.tabs.setCurrentIndex(2)
    wait(window.tasks)
    assert '后台处理' not in window.status.text()
    assert window.tabs.isEnabled()
    assert not window.progress.isVisible()


def test_same_directory_cancel_and_confirm_preserve_engine_semantics(window, tmp_path):
    path = tmp_path / 'original.docx'
    doc = Document()
    doc.add_paragraph('old text')
    doc.save(path)
    original = path.read_bytes()
    window.append_files([str(path)])
    window.rules.replace([['old text', 'new text']])
    window.output.setText(str(tmp_path))
    with patch('qt_main.QMessageBox.question', return_value=QMessageBox.No):
        window.replace()
    assert not window.tasks.busy
    assert path.read_bytes() == original
    with patch('qt_main.QMessageBox.question', return_value=QMessageBox.Yes):
        window.replace()
        wait(window.tasks)
    assert Document(path).paragraphs[0].text == 'new text'
    # The existing engine intentionally replaces in place after consent.
    # It makes no backup promise, and must not leave partial temp files.
    assert list(tmp_path.glob('*.docx')) == [path]
    assert not list(tmp_path.glob('.replace-simple-*'))


def test_editor_cancel_save_and_existing_output_remain_untouched(window, tmp_path):
    path = tmp_path / 'template.xlsx'
    output = tmp_path / 'existing.xlsx'
    book = Workbook()
    book.active['A1'] = 'original'
    book.save(path)
    book.close()
    original = path.read_bytes()
    output.write_bytes(original)
    editor = VisualEditor(window, WorkbookSession.open(path), [], {})
    editor.session.edit('Sheet', {'A1': 'edited'})
    try:
        with patch('qt_editor.QFileDialog.getSaveFileName', return_value=('', '')):
            editor.save()
        assert not editor.tasks.busy
        with patch('qt_editor.QFileDialog.getSaveFileName', return_value=(str(output), '')), patch('qt_editor.QMessageBox.warning') as warning:
            editor.save()
            wait(editor.tasks)
            assert warning.called
        assert path.read_bytes() == original
        assert output.read_bytes() == original
        assert editor.session.pending['Sheet']['A1'] == 'edited'
        with patch('qt_editor.QMessageBox.question', return_value=QMessageBox.No):
            editor.show()
            assert not editor.close()
            assert editor.isVisible()
    finally:
        with patch('qt_editor.QMessageBox.question', return_value=QMessageBox.Yes):
            editor.close()
        editor.deleteLater()


def test_busy_editor_save_shortcut_does_not_open_dialog(window, tmp_path):
    path = tmp_path / 'template.xlsx'
    book = Workbook()
    book.active['A1'] = 'text'
    book.save(path)
    book.close()
    editor = VisualEditor(window, WorkbookSession.open(path), [], {})
    try:
        editor.tasks.submit(lambda p: time.sleep(.15), lambda _: None)
        with patch('qt_editor.QFileDialog.getSaveFileName', return_value=('', '')) as chooser:
            editor.save()
        assert not chooser.called
    finally:
        wait(editor.tasks)
        editor.close()
        editor.deleteLater()


def test_settings_roundtrip_keeps_legacy_values_and_failed_close_open(window, tmp_path):
    from app_settings import load_settings, save_settings
    from PySide6.QtGui import QCloseEvent
    save_settings({'legacy_rule_path': 'keep-me', 'review_drag_profiles': {'saved': {}}})
    window.persist = True
    window.rules.replace([['原文', '替换值']])
    window.output.setText(str(tmp_path / 'output'))
    window.symbols.setText('★▲')
    window.keep_symbols.setChecked(True)
    failed = QCloseEvent()
    with patch('qt_main.save_settings', side_effect=OSError('写入失败')), patch('qt_main.QMessageBox.warning'):
        window.closeEvent(failed)
    assert not failed.isAccepted()
    assert load_settings()['legacy_rule_path'] == 'keep-me'
    event = QCloseEvent()
    window.closeEvent(event)
    assert event.isAccepted()
    restored = MainWindow()
    try:
        assert restored.rules.rows == [['原文', '替换值']]
        assert restored.output.text() == str(tmp_path / 'output')
        assert restored.symbols.text() == '★▲'
        assert restored.keep_symbols.isChecked()
        assert load_settings()['review_drag_profiles'] == {'saved': {}}
    finally:
        restored.persist = False
        restored.close()
        restored.deleteLater()
        window.persist = False


def test_editor_zoom_fullscreen_and_dimension_units(window, tmp_path):
    from PySide6.QtCore import QPoint, QPointF
    from PySide6.QtGui import QWheelEvent
    path = tmp_path / 'zoom.xlsx'
    book = Workbook()
    book.active['A1'] = '原文'
    book.save(path)
    book.close()
    editor = VisualEditor(window, WorkbookSession.open(path), [], {})
    editor.show()
    editor.activateWindow()
    APP.processEvents()
    try:
        editor.target.setCurrentIndex(editor.target.model().index(0, 0))
        editor.target.setFocus()
        initial_width = editor.target.columnWidth(0)
        wheel = QWheelEvent(QPointF(20, 20), QPointF(20, 20), QPoint(), QPoint(0, 120),
                            Qt.NoButton, Qt.ControlModifier, Qt.NoScrollPhase, False)
        APP.sendEvent(editor.target.viewport(), wheel)
        assert editor.target.columnWidth(0) > initial_width
        assert editor.session.layouts == {}
        assert editor.target.currentIndex().row() == 0
        assert editor.source_zoom.value() == 100
        assert editor.target_zoom.value() == 125
        editor._header_gesture = ('Sheet', {}, {})
        editor.header_resized('widths', 0, 250)
        editor.commit_header_resize()
        assert editor.session.layouts['Sheet']['widths'][1] == pytest.approx((200 - 5) / 7, abs=.001)
        QTest.keyClick(editor.target, Qt.Key_F11)
        APP.processEvents()
        assert editor.isFullScreen()
        QTest.keyClick(editor.target, Qt.Key_Escape)
        APP.processEvents()
        assert not editor.isFullScreen()
        assert editor.isVisible()
    finally:
        with patch('qt_editor.QMessageBox.question', return_value=QMessageBox.Yes):
            editor.close()
        editor.deleteLater()


def test_capture_layouts(window, tmp_path):
    """Optional rendered evidence; geometry assertions always run."""
    output = os.environ.get('REPLACE_QA_CAPTURE_DIR')
    captures = Path(output) if output else None
    if captures:
        captures.mkdir(parents=True, exist_ok=True)
    source_path = tmp_path / '2026年度采购服务项目-招标文件与评分要求.docx'
    result = source(source_path)
    window.append_files([str(source_path)])
    window.rules.replace([['[项目名称]', '2026年度采购服务项目'], ['[项目编号]', 'TEST-2026-001']] + [['', ''] for _ in range(6)])
    window.scan_files.append([str(source_path)])
    window.scans.load(result[0][0], result[1][0])
    window.scan_view.setCurrentIndex(window.proxy.index(0, 1))
    from template_library import create_category, add_templates
    create_category(window.library_root, '服务采购')
    add_templates(window.library_root, '服务采购', [str(source_path)])
    measurements = []
    for width, height in ((920, 640), (1260, 860)):
        window.resize(width, height)
        for tab in range(3):
            window.tabs.setCurrentIndex(tab)
            wait(window.tasks)
            QTest.qWait(60)
            measurements.append({'view': f'main-{tab}', 'requested': [width, height], 'actual': list(window.size().toTuple()), 'dpr': window.devicePixelRatioF()})
            if captures:
                window.grab().save(str(captures / f'main-{tab}-{width}.png'))
    path = tmp_path / 'template.xlsx'
    book = Workbook()
    for row in [('因素', '说明', '分值'), ('服务', '', 10)]:
        book.active.append(row)
    book.save(path)
    book.close()
    editor = VisualEditor(window, WorkbookSession.open(path), result[0][0], result[2])
    editor.show()
    editor.resize(860, 560)
    QTest.qWait(60)
    if captures:
        editor.grab().save(str(captures / 'editor.png'))
    mapping = MappingDialog(editor, result[0][0][0], editor.session, 'Sheet')
    mapping.show()
    mapping.build()
    wait(mapping.tasks)
    QTest.qWait(60)
    if captures:
        mapping.grab().save(str(captures / 'mapping.png'))
        measurements.append({'view': 'mapping', 'actual': list(mapping.size().toTuple()), 'preview_height': mapping.preview.height(), 'dpr': mapping.devicePixelRatioF()})
        (captures / 'geometry.json').write_text(json.dumps(measurements, indent=2), encoding='utf-8')
    try:
        assert mapping.preview.height() >= 140, mapping.preview.height()
    finally:
        mapping.close()
        mapping.deleteLater()
        editor.close()
        editor.deleteLater()
