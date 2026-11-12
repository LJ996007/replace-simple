"""Transfer permissions, transaction safety, and opt-in Windows OS mouse proof.
Run native proof: QT_QPA_PLATFORM=windows python -m unittest discover -s tests -p test_qt_drag_drop.py -v
"""
import os
import unittest
from copy import deepcopy
from unittest.mock import patch

from PySide6.QtCore import Qt, QPoint, QItemSelection, QItemSelectionModel, QEventLoop, QTimer
from PySide6.QtGui import QDragEnterEvent, QDragMoveEvent, QDropEvent
from PySide6.QtCore import QMimeData, QPointF
from PySide6.QtTest import QTest
from openpyxl import load_workbook

import test_qt_frontend as fixtures
from test_qt_frontend import APP
from qt_editor import VisualEditor, TransferTable, selected_bounds
from qt_models import SheetModel


class TransferTests(unittest.TestCase):
    setUp = fixtures.NativeFrontendTests.setUp
    item = fixtures.NativeFrontendTests.item

    def editor(self):
        item, stamps = self.item()
        editor = VisualEditor(None, self.session, [item], stamps)
        self.addCleanup(editor.scratch.cleanup)
        self.addCleanup(editor.deleteLater)
        self.addCleanup(editor.hide)
        editor.show()
        APP.processEvents()
        return editor

    def select(self, view, bounds):
        l, t, r, b = bounds
        model = view.model()
        view.selectionModel().select(QItemSelection(model.index(t-1, l-1), model.index(b-1, r-1)),
                                     QItemSelectionModel.ClearAndSelect)
        view.setCurrentIndex(model.index(t-1, l-1))
        # setCurrentIndex may update selection: set the rectangle last.
        view.selectionModel().select(QItemSelection(model.index(t-1, l-1), model.index(b-1, r-1)),
                                     QItemSelectionModel.ClearAndSelect)

    def test_roles_and_formula_ranges_enforced_in_flags_and_setData(self):
        editor = self.editor()
        source = editor.source_view.model()
        self.assertTrue(source.flags(source.index(0, 0)) & Qt.ItemIsDragEnabled)
        self.assertFalse(source.flags(source.index(0, 0)) & Qt.ItemIsEditable)
        self.assertFalse(source.setData(source.index(0, 0), 'changed'))
        preview = SheetModel(source.info)
        self.assertFalse(preview.flags(preview.index(0, 0)) & (Qt.ItemIsDragEnabled | Qt.ItemIsEditable))
        target = editor.target.model()
        self.assertFalse(target.flags(target.index(1, 3)) & (Qt.ItemIsDragEnabled | Qt.ItemIsEditable))
        self.assertFalse(target.setData(target.index(1, 3), 'overwrite'))
        info = deepcopy(target.info)
        info['formula_ranges'] = ['A2:B3']
        model = SheetModel(info, self.session, '评审')
        self.assertFalse(model.flags(model.index(2, 1)) & Qt.ItemIsEditable)
        self.assertFalse(model.setData(model.index(2, 1), 'overwrite'))
        self.assertEqual(self.session.pending, {})

    def test_fragmented_rectangle_and_merged_selection_but_reject_holes(self):
        editor = self.editor()
        view = editor.source_view
        model = view.model()
        selection = view.selectionModel()
        for row in (0, 1):
            selection.select(QItemSelection(model.index(row, 0), model.index(row, 1)), QItemSelectionModel.Select)
        self.assertEqual(selected_bounds(view), (1, 1, 2, 2))
        selection.clearSelection()
        for row, col in ((0, 0), (1, 1)):
            selection.select(model.index(row, col), QItemSelectionModel.Select)
        with self.assertRaises(ValueError):
            selected_bounds(view)
        model.info = {**model.info, 'merges': ['A1:B2']}
        selection.clearSelection()
        selection.select(model.index(0, 0), QItemSelectionModel.Select)
        self.assertEqual(selected_bounds(view), (1, 1, 2, 2))
        model.info = {**model.info, 'merges': ['B1:C2']}
        self.select(view, (1, 1, 2, 1))
        self.assertEqual(selected_bounds(view), (1, 1, 3, 2))

    def test_foreign_mime_rejected_at_every_stage(self):
        editor = self.editor()
        mime = QMimeData()
        mime.setData(TransferTable.mime_type, b'[1,1,2,2]')
        pos = editor.target.visualRect(editor.target.model().index(1, 1)).center()
        events = (QDragEnterEvent(pos, Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier),
                  QDragMoveEvent(pos, Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier),
                  QDropEvent(QPointF(pos), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier))
        for event, method in zip(events, (editor.target.dragEnterEvent, editor.target.dragMoveEvent, editor.target.dropEvent)):
            method(event)
            self.assertFalse(event.isAccepted())
        self.assertEqual(self.session.history, [])

    def test_multicell_one_undo_failures_atomic_and_save_preserves_sources(self):
        editor = self.editor()
        original = self.path.read_bytes(), self.source.read_bytes()
        self.assertTrue(editor.transfer_values((2, 2, 3, 3), 2, 2))
        self.assertEqual(len(self.session.history), 1)
        self.assertEqual(len(self.session.pending['评审']), 4)
        state = self.session.snapshot()
        history = list(self.session.history)
        with patch.object(editor, 'error') as error:
            self.assertFalse(editor.transfer_values((2, 2, 3, 3), 2, 3))  # formula after valid cell
            error.assert_called_once()
        self.assertEqual(self.session.snapshot(), state)
        self.assertEqual(self.session.history, history)
        for row, col, protected in ((1048576, 2, False), (2, 16384, False), (2, 2, True)):
            self.session.sheets['评审']['protected'] = protected
            before = self.session.snapshot()
            with patch.object(editor, 'error'):
                self.assertFalse(editor.transfer_values((2, 2, 3, 3), row, col))
            self.assertEqual(self.session.snapshot(), before)
            self.assertEqual(self.session.history, history)
        self.session.sheets['评审']['protected'] = False
        editor.undo()
        self.assertEqual(self.session.pending, {})
        editor.redo()
        output = self.root / '保存.xlsx'
        self.session.save(output)
        book = load_workbook(output)
        self.addCleanup(book.close)
        self.assertEqual(book['评审']['B2'].value, '新说明')
        self.assertEqual(book['评审']['B3'].value, '新服务')
        self.assertEqual(book['评审']['D2'].value, '=SUM(C2:C3)')
        self.assertEqual((self.path.read_bytes(), self.source.read_bytes()), original)

    def test_header_mouse_resize_persists_once_and_undoes(self):
        editor = self.editor()
        for header, axis in ((editor.target.horizontalHeader(), 'widths'),
                             (editor.target.verticalHeader(), 'heights')):
            start = (QPoint(header.sectionViewportPosition(0)+header.sectionSize(0)-1, header.height()//2)
                     if axis == 'widths' else QPoint(header.width()//2, header.sectionViewportPosition(0)+header.sectionSize(0)-1))
            end = start + (QPoint(35, 0) if axis == 'widths' else QPoint(0, 24))
            before = self.session.snapshot()
            count = len(self.session.history)
            QTest.mousePress(header.viewport(), Qt.LeftButton, Qt.NoModifier, start)
            QTest.mouseMove(header.viewport(), end, 40)
            QTest.mouseRelease(header.viewport(), Qt.LeftButton, Qt.NoModifier, end)
            APP.processEvents()
            self.assertEqual(len(self.session.history), count+1)
            value = self.session.layouts['评审'][axis][1]
            editor.target_changed()
            self.assertAlmostEqual(header.sectionSize(0), value*7+5 if axis == 'widths' else value*4/3, delta=1)
            output = self.root / (axis + '.xlsx')
            self.session.save(output)
            book = load_workbook(output)
            self.addCleanup(book.close)
            actual = book['评审'].column_dimensions['A'].width if axis == 'widths' else book['评审'].row_dimensions[1].height
            self.assertEqual(actual, value)
            editor.undo()
            self.assertEqual(self.session.snapshot(), before)

    @unittest.skipUnless(os.name == 'nt' and os.environ.get('QT_QPA_PLATFORM') == 'windows', 'requires native Windows desktop')
    def test_native_os_mouse_drag_start_to_drop_cancel_and_splitter(self):
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        class MouseInput(ctypes.Structure):
            _fields_ = [('dx', wintypes.LONG), ('dy', wintypes.LONG), ('mouseData', wintypes.DWORD),
                        ('dwFlags', wintypes.DWORD), ('time', wintypes.DWORD), ('dwExtraInfo', ctypes.c_size_t)]
        class InputUnion(ctypes.Union):
            _fields_ = [('mi', MouseInput), ('padding', ctypes.c_byte * (32 if ctypes.sizeof(ctypes.c_void_p) == 8 else 24))]
        class Input(ctypes.Structure):
            _fields_ = [('type', wintypes.DWORD), ('u', InputUnion)]
        def send(point, flags=0):
            # Qt global coordinates are logical pixels; Win32 uses physical pixels.
            ratio = APP.primaryScreen().devicePixelRatio()
            user32.SetCursorPos(round(point.x()*ratio), round(point.y()*ratio))
            if flags:
                data = Input(type=0, u=InputUnion(mi=MouseInput(dwFlags=flags)))
                if user32.SendInput(1, ctypes.byref(data), ctypes.sizeof(data)) != 1:
                    raise OSError('SendInput failed')
        editor = self.editor()
        editor.raise_()
        editor.activateWindow()
        user32.SetForegroundWindow(int(editor.winId()))
        QTest.qWait(200)
        screenshot = os.environ.get('DRAG_QA_SCREENSHOT')
        if screenshot:
            APP.primaryScreen().grabWindow(0).save(screenshot)
        trace = []
        def gesture(start, end):
            # Windows OLE's nested loop may not dispatch Qt timers. Deliver the
            # physical movement/release from a worker, independent of that loop.
            import threading
            import time
            loop = QEventLoop()
            failures = []
            def inject():
                try:
                    send(start)
                    time.sleep(.08)
                    send(start, 0x0002)
                    for n in range(1, 25):
                        time.sleep(.025)
                        send(start + (end-start)*n/24)
                    time.sleep(.1)
                except Exception as exc:
                    failures.append(exc)
                finally:
                    send(end, 0x0004)
            worker = threading.Thread(target=inject, daemon=True)
            worker.start()
            QTimer.singleShot(1300, loop.quit)
            loop.exec()
            worker.join(1)
            if failures:
                raise failures[0]
        original_start = TransferTable.startDrag
        original_drop = TransferTable.dropEvent
        original_enter = TransferTable.dragEnterEvent
        def start(view, actions):
            trace.append(('startDrag', selected_bounds(view)))
            return original_start(view, actions)
        def enter(view, event):
            result = original_enter(view, event)
            trace.append(('enter', event.isAccepted()))
            return result
        def drop(view, event):
            result = original_drop(view, event)
            trace.append(('drop', event.isAccepted()))
            return result
        self.select(editor.source_view, (2, 2, 3, 3))
        original_cells = deepcopy(editor.source_info)
        a = editor.source_view.viewport().mapToGlobal(editor.source_view.visualRect(editor.source_view.model().index(1, 1)).center())
        b = editor.target.viewport().mapToGlobal(editor.target.visualRect(editor.target.model().index(1, 1)).center())
        with patch.object(TransferTable, 'startDrag', start), patch.object(TransferTable, 'dragEnterEvent', enter), patch.object(TransferTable, 'dropEvent', drop):
            gesture(a, b)
            self.assertIn(('startDrag', (2, 2, 3, 3)), trace)
            self.assertIn(('enter', True), trace)
            self.assertIn(('drop', True), trace)
            self.assertEqual(len(self.session.history), 1)
            self.assertEqual(len(self.session.pending['评审']), 4)
            self.assertEqual(editor.source_info, original_cells)
            editor.undo()
            self.assertEqual(self.session.pending, {})
            editor.redo()
            before = self.session.snapshot()
            history = list(self.session.history)
            bad = editor.target.viewport().mapToGlobal(editor.target.visualRect(editor.target.model().index(1, 2)).center())
            with patch.object(editor, 'error') as error:
                gesture(a, bad)
                error.assert_called_once()
            self.assertIn(('drop', False), trace)
            self.assertEqual(self.session.snapshot(), before)
            self.assertEqual(self.session.history, history)
            gesture(a, editor.status.mapToGlobal(editor.status.rect().center()))
            self.assertEqual(self.session.snapshot(), before)
            self.assertEqual(self.session.history, history)
        handle = editor.splitter.handle(1)
        before_sizes = editor.splitter.sizes()
        origin = handle.mapToGlobal(handle.rect().center())
        gesture(origin, origin+QPoint(100, 0))
        self.assertGreater(editor.splitter.sizes()[0], before_sizes[0]+50)
        for header, axis in ((editor.target.horizontalHeader(), 'widths'),
                             (editor.target.verticalHeader(), 'heights')):
            point = (QPoint(header.sectionSize(0)-1, header.height()//2) if axis == 'widths'
                     else QPoint(header.width()//2, header.sectionSize(0)-1))
            origin = header.viewport().mapToGlobal(point)
            before = self.session.snapshot()
            history_count = len(self.session.history)
            gesture(origin, origin+(QPoint(35, 0) if axis == 'widths' else QPoint(0, 24)))
            self.assertEqual(len(self.session.history), history_count+1)
            self.assertIn(1, self.session.layouts['评审'][axis])
            editor.undo()
            self.assertEqual(self.session.snapshot(), before)
        # Repeat with a real merged Word source; the source stays immutable.
        editor.hide()
        from docx import Document
        doc = Document(self.source)
        doc.tables[0].cell(1, 1).merge(doc.tables[0].cell(1, 2))
        doc.save(self.source)
        editor = self.editor()
        editor.raise_()
        editor.activateWindow()
        user32.SetForegroundWindow(int(editor.winId()))
        QTest.qWait(150)
        self.assertTrue(editor.source_info['merges'])
        self.select(editor.source_view, (2, 2, 3, 3))
        a = editor.source_view.viewport().mapToGlobal(editor.source_view.visualRect(editor.source_view.model().index(1, 1)).center())
        b = editor.target.viewport().mapToGlobal(editor.target.visualRect(editor.target.model().index(1, 1)).center())
        before = self.session.snapshot()
        history_count = len(self.session.history)
        with patch.object(TransferTable, 'startDrag', start), patch.object(TransferTable, 'dragEnterEvent', enter), patch.object(TransferTable, 'dropEvent', drop):
            gesture(a, b)
        self.assertEqual(len(self.session.history), history_count+1)
        self.assertEqual(trace[-1], ('drop', True))
        self.assertIn('10', str(self.session.pending['评审']['B2']))
        editor.undo()
        self.assertEqual(self.session.snapshot(), before)
        print('WINDOWS_NATIVE_DRAG_PROOF', trace, 'splitter', before_sizes, '[+100, -100]', flush=True)
        screenshot = os.environ.get('DRAG_QA_SCREENSHOT')
        if screenshot:
            editor.grab().save(screenshot)
