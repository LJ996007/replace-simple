"""Render regressions for Fusion fallback surfaces, independent of text antialiasing."""
import os
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QPalette, QColor, QFontDatabase
from PySide6.QtTest import QTest
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QComboBox, QSpinBox, QTableView, QLineEdit, QToolBar, QToolButton,
    QMenu, QProgressBar, QCheckBox, QTabWidget, QFileDialog, QStyle,
    QStyleOptionSlider,
)
from qt_theme import apply_theme, PENDING, SELECTION
from qt_models import RowsModel, SheetModel

APP = QApplication.instance() or QApplication([])
if os.name == 'nt':
    for name in ('msyh.ttc', 'msyhbd.ttc'):
        QFontDatabase.addApplicationFont(str(Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts' / name))


class FlatThemeTests(unittest.TestCase):
    def setUp(self):
        apply_theme(APP)
        self.widgets = []

    def tearDown(self):
        for widget in reversed(self.widgets):
            widget.hide()
            widget.deleteLater()
        APP.processEvents()

    def show(self, widget, width=460, height=300):
        self.widgets.append(widget)
        widget.resize(width, height)
        widget.show()
        APP.processEvents()
        return widget

    def assertFlat(self, image, x, y, width=20, height=6, expected=None):
        # Widget geometry is logical; native grabs contain physical DPI pixels.
        dpr = image.devicePixelRatio()
        if dpr != 1:
            image = image.scaled(round(image.width() / dpr), round(image.height() / dpr))
        # Sample an interior rectangle: same-row pixels AND successive rows.
        # Never use text/antialiased border edges to infer a surface gradient.
        colors = {image.pixelColor(xx, yy).name()
                  for xx in range(x, x+width) for yy in range(y, y+height)}
        self.assertEqual(len(colors), 1, (x, y, colors))
        if expected:
            self.assertEqual(colors, {expected})
        color = image.pixelColor(x, y)
        self.assertFalse(color.red() > color.blue()+8 and color.green() > color.blue()+8,
                         f'warm UI surface: {color.name()}')

    def test_palette_fallback_is_cool_in_all_groups_and_roles(self):
        palette = APP.palette()
        for group in (QPalette.Active, QPalette.Inactive, QPalette.Disabled):
            for role in QPalette.ColorRole:
                if role in (QPalette.NoRole, QPalette.NColorRoles):
                    continue
                color = palette.color(group, role)
                self.assertFalse(color.red() > color.blue()+8 and color.green() > color.blue()+8,
                                 (group, role, color.name()))
        self.assertEqual(palette.color(QPalette.Inactive, QPalette.Highlight).name(), SELECTION)
        self.assertFalse(APP.testAttribute(Qt.AA_DontUseNativeDialogs))

    def test_table_alternation_selection_editor_and_document_exception(self):
        view = QTableView()
        model = RowsModel(['字段', '值'], [['', ''] for _ in range(30)], editable=True)
        view.setModel(model)
        view.setAlternatingRowColors(True)
        view.setColumnWidth(0, 180)
        self.show(view)
        view.clearFocus()
        APP.processEvents()
        for row, color in [(2, '#ffffff'), (3, '#f1f3f5')]:
            rect = view.visualRect(model.index(row, 0))
            self.assertFlat(view.viewport().grab().toImage(), rect.x()+60, rect.y()+7, expected=color)
        index = model.index(2, 0)
        view.setCurrentIndex(index)
        other = self.show(QWidget(), 100, 100)
        for active in (True, False):
            (view if active else other).activateWindow()
            view.setFocus() if active else other.setFocus()
            APP.processEvents()
            self.assertEqual(view.isActiveWindow(), active)
            rect = view.visualRect(index)
            self.assertFlat(view.viewport().grab().toImage(), rect.x()+60, rect.y()+7, expected=SELECTION)
        view.edit(index)
        APP.processEvents()
        editor = view.findChild(QLineEdit)
        self.assertIsNotNone(editor)
        editor.setText('编辑项')
        self.assertFlat(editor.grab().toImage(), 60, 5, expected='#ffffff')
        # Formatting supplied by Office is exempt only inside document cells.
        info = {'max_row': 1, 'max_col': 1, 'cells': {'A1': '文档底色'},
                'merges': [], 'styles': {'A1': {'bg': '#fff2cc', 'fg': '#a06000'}}}
        document = SheetModel(info)
        self.assertEqual(document.data(document.index(0, 0), Qt.BackgroundRole), QColor('#fff2cc'))
        self.assertEqual(info['styles']['A1']['bg'], '#fff2cc')
        view.closePersistentEditor(index)
        view.setModel(document)
        view.clearSelection()
        view.setCurrentIndex(document.index(-1, -1))
        APP.processEvents()
        rect = view.visualRect(document.index(0, 0))
        self.assertFlatDocument(view.viewport().grab().toImage(), rect.x()+100, rect.y()+8)

    def assertFlatDocument(self, image, x, y):
        dpr = image.devicePixelRatio()
        if dpr != 1:
            image = image.scaled(round(image.width() / dpr), round(image.height() / dpr))
        self.assertEqual({image.pixelColor(xx, y).name() for xx in range(x, x+15)}, {'#fff2cc'})

    def test_buttons_tabs_progress_and_disabled_surfaces_are_solid(self):
        root = QWidget()
        layout = QVBoxLayout(root)
        button = QPushButton('操作')
        button.setCheckable(True)
        layout.addWidget(button)
        tabs = QTabWidget()
        tabs.addTab(QWidget(), '第一页')
        tabs.addTab(QWidget(), '第二页')
        layout.addWidget(tabs)
        progress = QProgressBar()
        progress.setValue(60)
        layout.addWidget(progress)
        self.show(root)
        for checked, enabled in [(False, True), (True, True), (False, False)]:
            button.setChecked(checked)
            button.setEnabled(enabled)
            APP.processEvents()
            self.assertFlat(button.grab().toImage(), 15, 8)
        for i in (0, 1):
            tabs.setCurrentIndex(i)
            APP.processEvents()
            rect = tabs.tabBar().tabRect(i)
            self.assertFlat(tabs.tabBar().grab().toImage(), rect.x()+5, rect.y()+5, width=10)
        self.assertFlat(progress.grab().toImage(), 10, 4, expected='#91b7d3')
        self.assertFlat(progress.grab().toImage(), progress.width()-30, 4, expected='#edf0f2')

    def test_combo_popup_spin_scrollbar_menu_header_and_corner(self):
        root = QWidget()
        layout = QVBoxLayout(root)
        combo = QComboBox()
        combo.addItems(['选项一', '选项二', '选项三'])
        spin = QSpinBox()
        spin.setRange(0, 100)
        spin.setValue(10)
        view = QTableView()
        view.setModel(RowsModel(['字段', '值'], [['', ''] for _ in range(40)]))
        view.horizontalHeader().setSectionsClickable(True)
        layout.addWidget(combo)
        layout.addWidget(spin)
        layout.addWidget(view)
        self.show(root, 500, 400)
        self.assertFlat(combo.grab().toImage(), 60, 6, expected='#ffffff')
        QTest.mouseClick(combo, Qt.LeftButton, pos=QPoint(combo.width()-12, combo.height()//2))
        APP.processEvents()
        popup = combo.view()
        rect = popup.visualRect(popup.model().index(2, 0))
        self.assertFlat(popup.viewport().grab().toImage(), 130, rect.y()+4, height=4, expected='#ffffff')
        combo.hidePopup()
        QTest.mouseClick(spin, Qt.LeftButton, pos=QPoint(spin.width()-10, spin.height()//4))
        self.assertEqual(spin.value(), 11)
        for control in (combo, spin):
            control.setEnabled(False)
            APP.processEvents()
            self.assertFlat(control.grab().toImage(), 100, 6, expected='#edf0f2')
        header = view.horizontalHeader()
        QTest.mousePress(header.viewport(), Qt.LeftButton, pos=QPoint(50, 12))
        APP.processEvents()
        self.assertFlat(header.viewport().grab().toImage(), 65, 5, width=20, height=4, expected=SELECTION)
        QTest.mouseRelease(header.viewport(), Qt.LeftButton, pos=QPoint(50, 12))
        corner = view.findChild(QWidget, 'qt_tablecornerbutton')
        if corner is None:
            corner = next(w for w in view.findChildren(QWidget) if w.metaObject().className() == 'QTableCornerButton')
        self.assertFlat(corner.grab().toImage(), 4, 4, width=5, height=5, expected='#eef0f2')
        bar = view.verticalScrollBar()
        option = QStyleOptionSlider()
        bar.initStyleOption(option)
        handle = bar.style().subControlRect(QStyle.CC_ScrollBar, option, QStyle.SC_ScrollBarSlider, bar)
        self.assertFlat(bar.grab().toImage(), handle.center().x()-1, handle.y()+6, width=3, height=8, expected='#a9b4be')
        menu = QMenu()
        menu.addAction('普通操作')
        disabled = menu.addAction('不可用操作')
        disabled.setEnabled(False)
        checked = menu.addAction('勾选操作')
        checked.setCheckable(True)
        checked.setChecked(True)
        self.show(menu, 240, 150)
        menu.setActiveAction(menu.actions()[0])
        APP.processEvents()
        rect = menu.actionGeometry(menu.actions()[0])
        self.assertFlat(menu.grab().toImage(), rect.right()-20, rect.y()+5, width=8, height=5, expected=SELECTION)

    def test_combo_container_and_keyboard_current_are_quiet(self):
        root = QWidget()
        layout = QVBoxLayout(root)
        combo = QComboBox()
        combo.addItems(['选项一', '选项二', '选项三'])
        layout.addWidget(combo)
        self.show(root, 380, 160)
        combo.showPopup()
        QTest.qWait(250 if APP.platformName() == 'windows' else 40)
        APP.processEvents()
        view = combo.view()
        container = view.window()
        self.assertEqual(container.windowType(), Qt.Popup)
        self.assertTrue(container.windowFlags() & Qt.NoDropShadowWindowHint)
        self.assertTrue(container.windowFlags() & Qt.FramelessWindowHint)
        image = container.grab().toImage().scaled(container.size())
        self.assertEqual(image.pixelColor(0, image.height() // 2).name(), '#d5dbe0')
        self.assertEqual(image.pixelColor(image.width() // 2, 0).name(), '#d5dbe0')
        QTest.keyClick(view, Qt.Key_Down)
        APP.processEvents()
        row = view.visualRect(view.currentIndex())
        self.assertGreaterEqual(row.height(), 30)
        self.assertLessEqual(row.height(), 36)
        image = view.viewport().grab().toImage().scaled(view.viewport().size())
        image.setDevicePixelRatio(1)
        self.assertFlat(image, row.right()-40, row.y()+5, expected='#f0f3f6')
        self.assertEqual(image.pixelColor(row.left()+4, row.center().y()).name(), '#738493')
        plain = view.visualRect(view.model().index(2, 0))
        self.assertFlat(image, plain.right()-40, plain.y()+5, expected='#ffffff')
        # A focused item must not bring the global blue table outline back.
        self.assertEqual(image.pixelColor(row.right()-8, row.top()).name(), '#f0f3f6')
        QTest.keyClick(view, Qt.Key_Escape)
        self.assertFalse(container.isVisible())

    def test_document_colors_survive_edit_and_save_without_theme_storage(self):
        from openpyxl import Workbook, load_workbook
        from openpyxl.styles import PatternFill, Font
        from workbook_session import WorkbookSession
        with TemporaryDirectory() as directory:
            source = Path(directory, 'source.xlsx')
            output = Path(directory, 'output.xlsx')
            book = Workbook()
            book.active['A1'] = '内容格式'
            book.active['A1'].fill = PatternFill('solid', fgColor='FFF2CC')
            book.active['A1'].font = Font(color='A06000')
            book.save(source)
            book.close()
            original = source.read_bytes()
            session = WorkbookSession.open(source)
            sheet = next(iter(session.sheets))
            session.edit(sheet, {'A1': '修改文字'})
            model = SheetModel(session.view_info(sheet), session, sheet)
            self.assertEqual(model.data(model.index(0, 0), Qt.BackgroundRole).name(), PENDING)
            session.save(output)
            self.assertEqual(source.read_bytes(), original)
            saved = load_workbook(output)
            self.assertEqual(saved.active['A1'].fill.fgColor.rgb[-6:], 'FFF2CC')
            self.assertEqual(saved.active['A1'].font.color.rgb[-6:], 'A06000')
            self.assertEqual(saved.active['A1'].value, '修改文字')
            saved.close()

    def test_standard_icons_have_no_platform_yellow(self):
        from PySide6.QtWidgets import QMessageBox
        message = self.show(QMessageBox(QMessageBox.Warning, '提示', '核对输入后重试。'))
        icons = [APP.style().standardIcon(QStyle.SP_MessageBoxWarning),
                 APP.style().standardIcon(QStyle.SP_DirIcon)]
        for icon in icons:
            image = icon.pixmap(32, 32).toImage()
            colored = [image.pixelColor(x, y) for x in range(image.width())
                       for y in range(image.height()) if image.pixelColor(x, y).alpha() > 30]
            self.assertTrue(colored)
            self.assertFalse(any(c.red() > c.blue()+8 and c.green() > c.blue()+8 for c in colored))
        # Application icons stay flat; Windows file-dialog icons use system colors.
        image = message.iconPixmap().toImage()
        self.assertFalse(any((c := image.pixelColor(x, y)).red() > c.blue()+8
                             and c.green() > c.blue()+8
                             for x in range(image.width()) for y in range(image.height())))
        self.assertFlat(message.grab().toImage(), 4, 4, width=4, height=4,
                        expected='#f4f5f6')

    def test_file_dialogs_allow_native_open_save_and_folder_modes(self):
        self.assertFalse(APP.testAttribute(Qt.AA_DontUseNativeDialogs))
        for mode, accept in ((QFileDialog.ExistingFile, QFileDialog.AcceptOpen),
                             (QFileDialog.ExistingFiles, QFileDialog.AcceptOpen),
                             (QFileDialog.Directory, QFileDialog.AcceptOpen),
                             (QFileDialog.AnyFile, QFileDialog.AcceptSave)):
            dialog = QFileDialog()
            self.widgets.append(dialog)
            dialog.setFileMode(mode)
            dialog.setAcceptMode(accept)
            self.assertFalse(dialog.testOption(QFileDialog.DontUseNativeDialog))
            self.assertEqual(dialog.fileMode(), mode)
            self.assertEqual(dialog.acceptMode(), accept)

    def test_toolbar_overflow_and_popup_actions_remain_usable(self):
        root = QWidget()
        layout = QVBoxLayout(root)
        toolbar = QToolBar()
        toolbar.setMovable(False)
        toolbar.setMaximumWidth(180)
        layout.addWidget(toolbar)
        triggered = []
        for number in range(10):
            toolbar.addAction(f'操作 {number}', lambda checked=False, n=number: triggered.append(n))
        self.show(root, 240, 150)
        extension = toolbar.findChild(QToolButton, 'qt_toolbar_ext_button')
        self.assertTrue(extension.isVisible())
        self.assertFlat(extension.grab().toImage(), 3, 3, width=4, height=4)
        menu = extension.menu()
        self.assertIsNotNone(menu)
        self.assertGreaterEqual(extension.width(), 26)
        menu.popup(extension.mapToGlobal(QPoint(0, extension.height())))
        APP.processEvents()
        QTest.mouseClick(menu, Qt.LeftButton, pos=menu.actionGeometry(menu.actions()[-1]).center())
        self.assertEqual(triggered, [9])

    def test_keyboard_focus_is_visible_on_primary_action(self):
        root = QWidget()
        layout = QVBoxLayout(root)
        first = QPushButton('先选中此项')
        primary = QPushButton('主要操作')
        primary.setProperty('primary', True)
        layout.addWidget(first)
        layout.addWidget(primary)
        self.show(root, 300, 160)
        first.setFocus()
        APP.processEvents()
        before = primary.grab().toImage()
        QTest.keyClick(first, Qt.Key_Tab)
        APP.processEvents()
        self.assertTrue(primary.hasFocus())
        after = primary.grab().toImage()
        y = after.height() // 2
        self.assertEqual(before.pixelColor(1, y).name(), '#245d85')
        self.assertEqual(after.pixelColor(1, y).name(), '#142f43')

    def test_qt_file_dialog_fallback_navigation_is_available(self):
        with TemporaryDirectory() as directory:
            Path(directory, 'sample.xlsx').touch()
            dialog = QFileDialog(None, '选择工作簿', directory)
            # Explicitly test the fallback for platforms without native dialogs;
            # application file pickers leave this option disabled.
            dialog.setOption(QFileDialog.DontUseNativeDialog, True)
            dialog.setFileMode(QFileDialog.ExistingFile)
            dialog.setNameFilter('Excel (*.xlsx)')
            dialog.selectFile(str(Path(directory, 'sample.xlsx')))
            self.show(dialog, 760, 480)
            APP.processEvents()
            self.assertTrue(dialog.findChildren(QLineEdit))
            self.assertEqual(Path(dialog.selectedFiles()[0]).name, 'sample.xlsx')
            self.assertEqual(dialog.palette().color(QPalette.Window).name(), '#f4f5f6')


if __name__ == '__main__':
    unittest.main()
