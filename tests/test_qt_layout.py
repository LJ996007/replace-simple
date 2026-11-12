"""Geometry, mouse interaction and rendered splitter contracts."""
import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
import unittest
from PySide6.QtCore import QPoint, Qt
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QComboBox, QStyle
from PySide6.QtTest import QTest
from qt_layout import BoundedComboBox
from qt_main import MainWindow
from qt_theme import apply_theme

APP = QApplication.instance() or QApplication([])
apply_theme(APP)


def settle():
    APP.processEvents()
    # Windows can animate combo popups before mapping the final list window.
    QTest.qWait(200 if APP.platformName() == 'windows' else 40)


def drag(split, delta):
    handle = split.handle(1)
    start = handle.rect().center()
    offset = QPoint(delta, 0) if split.orientation() == Qt.Horizontal else QPoint(0, delta)
    QTest.mousePress(handle, Qt.LeftButton, pos=start)
    QTest.mouseMove(handle, start + offset, 30)
    QTest.mouseRelease(handle, Qt.LeftButton, pos=handle.rect().center())
    settle()


class LayoutInteractionTests(unittest.TestCase):
    def test_popup_geometry_mouse_keyboard_and_screen_bottom(self):
        host = QWidget()
        if APP.platformName() == 'windows':
            host.setWindowFlag(Qt.WindowStaysOnTopHint)
        self.addCleanup(host.deleteLater)
        self.addCleanup(host.close)
        layout = QVBoxLayout(host)
        for cls in (BoundedComboBox, QComboBox):
            combo = cls()
            combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            combo.setMinimumContentsLength(12)
            combo.setMinimumWidth(0)
            layout.addWidget(combo)
            combo.addItems(['中文分类名称非常长' * 20, '采购方式二', '采购方式三'])
            combo.setMaxVisibleItems(3)
        host.resize(360, 140)
        screen = APP.primaryScreen().availableGeometry()
        for bottom in (False, True):
            if bottom:
                # Put the bounded selector nearest the screen edge so its
                # own placement code must take the upward branch.
                layout.insertWidget(0, host.findChildren(QComboBox)[-1])
            host.move(screen.left() + 30, screen.bottom() - host.height() if bottom else screen.top() + 50)
            host.show()
            host.raise_()
            host.activateWindow()
            # Allow Windows' asynchronous activation before creating a popup;
            # a late activation event otherwise dismisses it during the test.
            QTest.qWait(300)
            settle()
            for combo in host.findChildren(QComboBox):
                combo.setCurrentIndex(1)
                self.assertEqual(combo.style().styleHint(QStyle.SH_ComboBox_Popup, None, combo), 0)
                combo.showPopup()
                settle()
                popup = combo.view().window()
                self.assertTrue(popup.isVisible(), (bottom, type(combo).__name__, host.geometry(), combo.geometry()))
                rect = popup.frameGeometry()
                if isinstance(combo, BoundedComboBox):
                    self.assertLessEqual(rect.width(), combo.width())
                top = combo.mapToGlobal(QPoint()).y()
                below = top + combo.height()
                self.assertTrue(rect.top() >= below + 4 or rect.bottom() <= top - 5,
                                (rect, top, below))
                self.assertTrue(screen.contains(rect), (screen, rect))
                if not bottom:
                    self.assertEqual(rect.top(), below + 4)
                elif combo is layout.itemAt(layout.count() - 1).widget():
                    self.assertEqual(rect.bottom(), top - 5)
                view = combo.view()
                target = view.model().index(2, 0)
                view.scrollTo(target)
                QTest.mouseMove(view.viewport(), view.visualRect(target).center(), 30)
                QTest.mouseClick(view.viewport(), Qt.LeftButton, pos=view.visualRect(target).center())
                settle()
                self.assertFalse(popup.isVisible())
                self.assertEqual(combo.currentText(), '采购方式三')
                combo.showPopup()
                settle()
                QTest.keyClick(view, Qt.Key_Up)
                QTest.keyClick(view, Qt.Key_Return)
                settle()
                self.assertFalse(popup.isVisible())
                self.assertEqual(combo.currentText(), '采购方式二')
                combo.showPopup()
                settle()
                QTest.keyClick(view, Qt.Key_Down)
                QTest.keyClick(view, Qt.Key_Escape)
                settle()
                self.assertFalse(popup.isVisible())
                self.assertEqual(combo.currentText(), '采购方式二')
                if isinstance(combo, BoundedComboBox):
                    self.assertEqual(combo.toolTip(), combo.currentText())
                    self.assertEqual(combo.itemData(0, Qt.ToolTipRole), combo.itemText(0))
                    self.assertEqual(view.textElideMode(), Qt.ElideMiddle)
                # Clicking the current selector again closes the open popup.
                combo.showPopup()
                settle()
                # Native popup mouse grabbing routes an outside selector click
                # to the popup, with coordinates outside its client rectangle.
                QTest.mouseClick(popup, Qt.LeftButton, pos=popup.mapFromGlobal(
                    combo.mapToGlobal(combo.rect().center())))
                settle()
                self.assertFalse(popup.isVisible())

    def assert_thin_handle(self, split):
        handle = split.handle(1)
        # The entire mouse target stays in the splitter's reserved layout slot.
        self.assertGreaterEqual(split.handleWidth(), 6)
        self.assertLessEqual(split.handleWidth(), 10)
        self.assertFalse(handle.geometry().intersects(split.widget(0).geometry()))
        self.assertFalse(handle.geometry().intersects(split.widget(1).geometry()))
        # Verify actual pixels, rather than equating the hit width with line width.
        image = handle.grab().toImage().scaled(handle.size())
        if split.orientation() == Qt.Horizontal:
            colors = [image.pixelColor(x, image.height() // 2).name() for x in range(image.width())]
        else:
            colors = [image.pixelColor(image.width() // 2, y).name() for y in range(image.height())]
        scan_surface = bool(split.property('scanSurface'))
        self.assertEqual(colors.count('#e0e4e7'), 1 if scan_surface else 2, colors)
        # Fractional native DPI can blend gutter/line boundary pixels.
        for y in colors:
            self.assertTrue(all(224 <= int(y[i:i+2], 16) <= (255 if scan_surface else 246)
                                for i in (1, 3, 5)), colors)
        center = handle.rect().center()
        QTest.mouseMove(handle, center, 20)
        settle()
        self.assertEqual(handle.grab().toImage().scaled(handle.size()).pixelColor(center).name(), '#cbd4dc')
        QTest.mousePress(handle, Qt.LeftButton, pos=center)
        settle()
        self.assertEqual(handle.grab().toImage().scaled(handle.size()).pixelColor(center).name(), '#b5c3cf')
        QTest.mouseRelease(handle, Qt.LeftButton, pos=center)
        QTest.mouseMove(split.widget(0), QPoint(20, 20), 20)
        settle()

    def test_scan_joined_boundary_pixels_and_preserved_content(self):
        from word_table_exporter import ExportedTable, ExportedTableCell, TableScanItem
        window = MainWindow(restore=False)
        if APP.platformName() == 'windows':
            window.setWindowFlag(Qt.WindowStaysOnTopHint)
        self.addCleanup(window.deleteLater)
        self.addCleanup(window.close)
        window.tabs.setCurrentIndex(1)
        window.show()
        split = window.scan_view.parentWidget()

        def assert_boundary(text=False):
            QTest.mouseMove(window.search, QPoint(10, 10))
            settle()
            preview = window.preview_text if text else window.scan_preview
            # Grab the WHOLE joined boundary, including both adjacent frames
            # and the row-header exterior, rather than only the splitter handle.
            # Sample inside a row; its horizontal rule is real content and
            # must not be counted as another vertical decoration at the edge.
            row_y = preview.viewport().height() // 2 if text else preview.rowHeight(0) // 2
            y = preview.viewport().mapTo(split, QPoint(0, row_y)).y()
            image = split.grab().toImage().scaled(split.size())
            left = split.handle(1).x() - 1
            right = preview.mapTo(split, QPoint()).x()
            colors = [image.pixelColor(x, y) for x in range(left, right + 1)]
            sample = [c.name() for c in colors]
            # All decorative frame/line colors are darker than the content
            # backgrounds (white, scrollbar track, alternating row, row header).
            ink = [i for i, c in enumerate(colors) if c.red() < 235]
            runs = sum(i == 0 or ink[i] != ink[i - 1] + 1 for i in range(len(ink)))
            self.assertEqual(runs, 1, sample)
            self.assertTrue(all(1 <= x <= 8 for x in ink), sample)
            self.assertLessEqual(len(ink), 2, sample)  # fractional DPI coverage
            self.assertGreaterEqual(colors[0].red(), 238, sample)
            self.assertGreaterEqual(colors[-1].red(), 238, sample)
            # Measure physical-pixel coverage: one logical pixel of line ink,
            # with tolerance for rasterization at native fractional scaling.
            handle_image = split.handle(1).grab().toImage()
            row = handle_image.height() // 2
            coverage = sum((255 - handle_image.pixelColor(x, row).red()) / 31
                           for x in range(handle_image.width()))
            self.assertAlmostEqual(coverage / handle_image.devicePixelRatio(), 1, delta=.15)

        assert_boundary()  # empty
        cells = [ExportedTableCell(r + 1, c + 1, 1, 1, f'内容 {r + 1}-{c + 1}')
                 for r in range(40) for c in range(5)]
        item = TableScanItem('采购需求.docx', '采购需求.docx', 0, '技术评分', '',
                             '项目理解与实施方案', 40, 5, '', ExportedTable(cells, 40, 5))
        window.scans.load([item] * 40, [])
        window.scan_view.setCurrentIndex(window.proxy.index(0, 1))
        settle()
        self.assertTrue(window.scan_view.verticalScrollBar().isVisible())
        self.assertTrue(window.scan_preview.verticalScrollBar().isVisible())
        for view in (window.scan_view, window.scan_preview):
            self.assertTrue(view.verticalHeader().isVisible())
            self.assertTrue(view.horizontalHeader().isVisible())
            self.assertTrue(view.showGrid())
            self.assertEqual(view.model().headerData(0, Qt.Vertical), '1')
        self.assertEqual(window.scan_preview.model().headerData(0, Qt.Horizontal), 'A')
        self.assertEqual(window.scan_preview.model().index(0, 0).data(), '内容 1-1')
        for width, height in ((1260, 860), (920, 640)):
            window.resize(width, height)
            for tab in (0, 1):
                window.preview_stack.setCurrentIndex(tab)
                for focus in (window.search, window.scan_view,
                              window.preview_text if tab else window.scan_preview):
                    focus.setFocus()
                    assert_boundary(text=bool(tab))
            self.assertIn('内容 1-1', window.preview_text.toPlainText())
            before = split.sizes()[0]
            drag(split, 45)
            self.assertGreater(split.sizes()[0], before + 20)
            assert_boundary(text=True)
            window.preview_stack.setCurrentIndex(0)
            assert_boundary()
        # Ordinary pages retain their independent framed tables.
        self.assertEqual(window.file_view.frameWidth(), 1)
        self.assertEqual(window.scan_file_view.frameWidth(), 1)

    def test_scan_source_and_result_drags_resize_and_neighbor_controls(self):
        window = MainWindow(restore=False)
        if APP.platformName() == 'windows':
            window.setWindowFlag(Qt.WindowStaysOnTopHint)
        self.addCleanup(window.deleteLater)
        self.addCleanup(window.close)
        window.tabs.setCurrentIndex(1)
        window.show()
        settle()
        vertical = window.scan_splitter
        horizontal = window.scan_view.parentWidget()
        self.assertGreater(window.scan_file_view.maximumHeight(), 10000)
        for split in (vertical, horizontal):
            self.assertFalse(split.childrenCollapsible())
            self.assert_thin_handle(split)
        self.assertLessEqual(abs(horizontal.sizes()[0] - horizontal.sizes()[1]), 2)
        initial = vertical.sizes()[0]
        filter_y = window.search.mapTo(window, QPoint()).y()
        drag(vertical, 75)
        self.assertGreater(vertical.sizes()[0], initial + 40)
        self.assertEqual(window.search.mapTo(window, QPoint()).y() - filter_y,
                         vertical.sizes()[0] - initial)
        drag(vertical, -45)
        self.assertLess(vertical.sizes()[0], initial + 60)
        initial_left = horizontal.sizes()[0]
        drag(horizontal, 65)
        self.assertGreater(horizontal.sizes()[0], initial_left + 35)
        ratio = horizontal.sizes()[0] / sum(horizontal.sizes())
        for width, height in ((920, 640), (1420, 920), (1260, 860)):
            window.resize(width, height)
            settle()
            self.assertEqual(window.size().toTuple(), (width, height))
            self.assertGreaterEqual(min(horizontal.sizes()), 260)
            self.assertAlmostEqual(horizontal.sizes()[0] / sum(horizontal.sizes()), ratio, delta=.02)
            self.assertGreaterEqual(vertical.sizes()[0], 85)
            self.assertGreaterEqual(vertical.sizes()[1], 210)
            self.assertTrue(window.search.isVisible())
            self.assertTrue(window.scan_view.isVisible())
            self.assertTrue(window.preview_stack.isVisible())
        drag(vertical, -2000)
        self.assertGreaterEqual(vertical.sizes()[0], 85)
        drag(vertical, 4000)
        self.assertGreaterEqual(vertical.sizes()[1], 210)
        # Real clicks next to the narrow divider still reach the table and filter.
        window.scan_files.append(['来源文件.docx'])
        rect = window.scan_file_view.visualRect(window.scan_files.index(0, 0))
        QTest.mouseClick(window.scan_file_view.viewport(), Qt.LeftButton, pos=rect.center())
        self.assertEqual(window.scan_file_view.currentIndex().row(), 0)
        QTest.mouseClick(window.search, Qt.LeftButton)
        QTest.keyClicks(window.search, 'filter')
        self.assertEqual(window.search.text(), 'filter')


if __name__ == '__main__':
    unittest.main()
