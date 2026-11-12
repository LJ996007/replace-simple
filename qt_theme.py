"""Application-wide flat Qt theme; document cell formatting belongs to SheetModel.

All colors here describe application chrome. Office foreground/fill colors are
content and are deliberately not normalized, either in preview or on disk.
"""
from pathlib import Path

from PySide6.QtCore import Qt, QEvent, QObject, QSize, QTimer
from PySide6.QtGui import QColor, QFont, QPalette, QIcon, QPixmap, QPainter, QBrush
from PySide6.QtWidgets import (QApplication, QProxyStyle, QStyle, QComboBox,
                               QFrame, QStyledItemDelegate, QStyleOptionViewItem)

from qt_layout import BoundedComboBox, position_combo_popup

SURFACE = '#f4f5f6'
TEXT = '#28343f'
BLUE = '#245d85'
SELECTION = '#dceaf6'
PENDING = '#e1edf7'


def theme_palette():
    """Set every Fusion role, including inactive/disabled fallback painting."""
    palette = QPalette()
    roles = {
        'Window': SURFACE, 'WindowText': TEXT, 'Base': '#ffffff',
        'AlternateBase': '#f1f3f5', 'Text': TEXT, 'Button': '#ffffff',
        'ButtonText': TEXT, 'BrightText': '#ffffff', 'Light': '#ffffff',
        'Midlight': '#eef0f2', 'Mid': '#b9c2ca', 'Dark': '#667582',
        'Shadow': '#667582', 'Highlight': SELECTION,
        'HighlightedText': '#173f5e', 'ToolTipBase': '#ffffff',
        'ToolTipText': TEXT, 'Link': BLUE, 'LinkVisited': '#405e7a',
        'PlaceholderText': '#697581', 'Accent': BLUE,
    }
    for group in (QPalette.Active, QPalette.Inactive, QPalette.Disabled):
        for role, color in roles.items():
            palette.setColor(group, getattr(QPalette.ColorRole, role), QColor(color))
    for role in (QPalette.Text, QPalette.WindowText, QPalette.ButtonText,
                 QPalette.PlaceholderText):
        palette.setColor(QPalette.Disabled, role, QColor('#77828c'))
    for role in (QPalette.Base, QPalette.Button):
        palette.setColor(QPalette.Disabled, role, QColor('#edf0f2'))
    palette.setColor(QPalette.Disabled, QPalette.Highlight, QColor('#e2e6ea'))
    palette.setColor(QPalette.Disabled, QPalette.HighlightedText, QColor('#697581'))
    return palette


_ICON_ROOT = (Path(__file__).resolve().parent / 'theme_icons').as_posix()
STYLE = r'''
QWidget { font-family: "Microsoft YaHei UI"; font-size: 10pt; }
QMainWindow, QDialog, QTabWidget, QTabWidget > QWidget { background: #f4f5f6; }
QWidget:disabled { color: #77828c; }
QLabel#title { font-size: 19pt; font-weight: 600; color: #28343f; }
QLabel#hint { color: #596875; padding: 3px 0; }
QPushButton, QToolButton {
    color: #28343f; background: #ffffff; border: 1px solid #bdc6ce;
    border-radius: 3px; padding: 6px 12px; min-height: 20px;
}
QPushButton:hover, QToolButton:hover { background: #edf3f8; border-color: #829db1; }
QPushButton:pressed, QToolButton:pressed, QPushButton:checked, QToolButton:checked {
    background: #dceaf6; border-color: #245d85;
}
QPushButton:focus, QToolButton:focus { border: 1px solid #245d85; }
QPushButton:default { border: 1px solid #245d85; }
QPushButton:disabled, QToolButton:disabled {
    color: #77828c; background: #edf0f2; border-color: #d5dbe0;
}
QPushButton[primary="true"] { color: #ffffff; background: #245d85; border-color: #245d85; }
QPushButton[primary="true"]:focus { border: 2px solid #142f43; padding: 5px 11px; }
QPushButton[primary="true"]:hover { background: #1e5277; }
QPushButton[primary="true"]:pressed { background: #183f5d; }
QPushButton[primary="true"]:disabled { color: #f4f5f6; background: #879aaa; border-color: #879aaa; }
QLineEdit, QComboBox, QAbstractSpinBox {
    color: #28343f; background: #ffffff; border: 1px solid #bdc6ce;
    border-radius: 3px; padding: 5px 7px; min-height: 22px;
    selection-background-color: #dceaf6; selection-color: #173f5e;
}
QLineEdit:focus, QComboBox:focus, QAbstractSpinBox:focus,
QAbstractItemView:focus, QPlainTextEdit:focus, QTextEdit:focus { border: 1px solid #245d85; }
QLineEdit:disabled, QComboBox:disabled, QAbstractSpinBox:disabled,
QPlainTextEdit:disabled, QTextEdit:disabled { color: #77828c; background: #edf0f2; border-color: #d5dbe0; }
QLineEdit[readOnly="true"] { background: #f1f3f5; }
QLineEdit QToolButton { min-height: 0px; min-width: 0px; padding: 0px; border: none; background: transparent; }
QComboBox { padding-right: 29px; }
QComboBox::drop-down { subcontrol-origin: border; subcontrol-position: top right;
    width: 26px; border: none; border-left: 1px solid #d5dbe0; background: #f1f3f5; }
QComboBox::drop-down:hover { background: #e5edf4; }
QComboBox::down-arrow, QToolButton::menu-indicator { image: url("@ICONS@/down.svg"); width: 10px; height: 10px; }
QComboBox::down-arrow:disabled { image: url("@ICONS@/down-disabled.svg"); }
QAbstractSpinBox { padding-right: 25px; }
QAbstractSpinBox::up-button, QAbstractSpinBox::down-button {
    subcontrol-origin: border; width: 23px; border: none;
    border-left: 1px solid #d5dbe0; background: #f1f3f5;
}
QAbstractSpinBox::up-button { subcontrol-position: top right; border-bottom: 1px solid #d5dbe0; }
QAbstractSpinBox::down-button { subcontrol-position: bottom right; }
QAbstractSpinBox::up-button:hover, QAbstractSpinBox::down-button:hover { background: #dceaf6; }
QAbstractSpinBox::up-button:pressed, QAbstractSpinBox::down-button:pressed { background: #c9ddec; }
QAbstractSpinBox::up-arrow { image: url("@ICONS@/up.svg"); width: 9px; height: 9px; }
QAbstractSpinBox::down-arrow { image: url("@ICONS@/down.svg"); width: 9px; height: 9px; }
QAbstractSpinBox::up-arrow:disabled, QAbstractSpinBox::up-arrow:off { image: url("@ICONS@/up-disabled.svg"); }
QAbstractSpinBox::down-arrow:disabled, QAbstractSpinBox::down-arrow:off { image: url("@ICONS@/down-disabled.svg"); }
QAbstractItemView, QPlainTextEdit, QTextEdit {
    color: #28343f; background: #ffffff; alternate-background-color: #f1f3f5;
    border: 1px solid #cbd2d8; border-radius: 0px; gridline-color: #e2e6ea;
    selection-background-color: #dceaf6; selection-color: #173f5e;
}
QAbstractItemView { outline: 0; }
QAbstractItemView::item:hover { background: #edf3f8; }
QAbstractItemView::item:selected { background: #dceaf6; color: #173f5e; }
QAbstractItemView::item:selected:!active { background: #dceaf6; color: #173f5e; }
QAbstractItemView::item:selected:disabled { background: #e2e6ea; color: #697581; }
QAbstractItemView::item:focus { border: 1px solid #245d85; }
QAbstractItemView QLineEdit { padding: 0px 3px; min-height: 0px; border-radius: 0px; }
QComboBox QAbstractItemView, QComboBox QAbstractItemView:focus {
    background: #ffffff; color: #28343f; border: none; outline: 0;
    selection-background-color: #f0f3f6; selection-color: #28343f;
}
QComboBox QAbstractItemView::item { border: none; }
QComboBox QAbstractItemView::item:hover,
QComboBox QAbstractItemView::item:selected,
QComboBox QAbstractItemView::item:selected:!active,
QComboBox QAbstractItemView::item:focus {
    background: #f0f3f6; color: #28343f; border: none;
}
QHeaderView { background: #eef0f2; }
QHeaderView::section, QTableCornerButton::section {
    color: #465461; background: #eef0f2; border: none;
    border-right: 1px solid #cbd2d8; border-bottom: 1px solid #cbd2d8;
    padding: 7px; font-weight: 600;
}
QHeaderView::section:hover { background: #e4eaf0; }
QHeaderView::section:pressed, QHeaderView::section:checked, QTableCornerButton::section:pressed { background: #dceaf6; color: #173f5e; }
QHeaderView::up-arrow { image: url("@ICONS@/up.svg"); width: 10px; height: 10px; }
QHeaderView::down-arrow { image: url("@ICONS@/down.svg"); width: 10px; height: 10px; }
QScrollBar:vertical { background: #f1f3f5; width: 14px; margin: 0px; border: none; }
QScrollBar:horizontal { background: #f1f3f5; height: 14px; margin: 0px; border: none; }
QScrollBar::handle { background: #a9b4be; border: 3px solid #f1f3f5; border-radius: 5px; }
QScrollBar::handle:vertical { min-height: 28px; }
QScrollBar::handle:horizontal { min-width: 28px; }
QScrollBar::handle:hover { background: #7d95a7; }
QScrollBar::handle:pressed { background: #245d85; }
QScrollBar::handle:disabled { background: #d5dbe0; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0px; height: 0px; border: none; background: #f1f3f5; }
QScrollBar::add-page, QScrollBar::sub-page { background: #f1f3f5; }
QAbstractScrollArea::corner { background: #eef0f2; border: none; }
QToolBar { background: #f4f5f6; border: none; spacing: 6px; padding: 2px 0px; }
QToolBar::separator { background: #cbd2d8; width: 1px; height: 1px; margin: 4px; }
QToolButton { padding: 5px 9px; }
QToolButton#qt_toolbar_ext_button { image: url("@ICONS@/more.svg"); background: #eef0f2;
    border: 1px solid #bdc6ce; padding: 2px; min-width: 16px; min-height: 20px; }
QToolButton#qt_toolbar_ext_button:hover { background: #dceaf6; }
QToolButton::menu-button { background: #eef0f2; border-left: 1px solid #cbd2d8; width: 16px; }
QMenuBar { background: #f4f5f6; color: #28343f; }
QMenuBar::item { background: transparent; padding: 6px 10px; }
QMenuBar::item:selected, QMenuBar::item:pressed { background: #dceaf6; }
QMenu { color: #28343f; background: #ffffff; border: 1px solid #bdc6ce; padding: 4px; }
QMenu::item { padding: 7px 28px; background: transparent; }
QMenu::item:selected { background: #dceaf6; color: #173f5e; }
QMenu::item:disabled { color: #77828c; }
QMenu::separator { height: 1px; background: #d5dbe0; margin: 4px 8px; }
QMenu::right-arrow { image: url("@ICONS@/right.svg"); width: 10px; height: 10px; }
QMenu::indicator { width: 14px; height: 14px; }
QMenu::indicator:checked { image: url("@ICONS@/check-blue.svg"); }
QTabWidget::pane { border: none; background: #f4f5f6; }
QTabBar { background: #f4f5f6; }
QTabBar::tab { color: #596875; padding: 10px 20px; background: #e9ecef;
    border: none; border-bottom: 2px solid #d5dbe0; margin-right: 2px; }
QTabBar::tab:hover { background: #e1e9f0; }
QTabBar::tab:selected { color: #173f5e; background: #f4f5f6; border-bottom: 2px solid #245d85; }
QTabBar::tab:focus { border-top: 1px solid #245d85; }
QTabBar::tab:disabled { color: #77828c; background: #edf0f2; border-color: #d5dbe0; }
QTabBar QToolButton { background: #eef0f2; }
QSplitter::handle { background: #e0e4e7; border: 0px solid #f4f5f6; }
QSplitter::handle:horizontal { border-left-width: 3px; border-right-width: 3px; }
QSplitter::handle:vertical { border-top-width: 3px; border-bottom-width: 3px; }
QSplitter::handle:hover { background: #cbd4dc; }
QSplitter::handle:pressed { background: #b5c3cf; }
/* The scan boundary owns one line; adjacent view frames must not add more.
   Keep focus on the other three edges and keep actual table headers/grid. */
QAbstractItemView[scanEdge="results"] { border-right: 0px; }
QAbstractItemView[scanEdge="preview"], QPlainTextEdit[scanEdge="preview"] { border-left: 0px; }
QSplitter[scanSurface="true"], QTabWidget[scanSurface="true"],
QTabWidget[scanSurface="true"]::pane { background: #ffffff; }
QSplitter[scanSurface="true"] QHeaderView { border: none; }
QTabWidget[scanSurface="true"] QTabBar { background: #eef0f2; }
QTabWidget[scanSurface="true"] QTabBar::tab { padding-top: 8px; padding-bottom: 8px; }
QSplitter[scanSurface="true"]::handle:horizontal {
    background: #e0e4e7; border-left: 3px solid #ffffff; border-right: 4px solid #ffffff;
}
QSplitter[scanSurface="true"]::handle:horizontal:hover { background: #cbd4dc; }
QSplitter[scanSurface="true"]::handle:horizontal:pressed { background: #b5c3cf; }
QProgressBar { color: #28343f; border: 1px solid #bdc6ce; background: #edf0f2;
    border-radius: 2px; text-align: center; min-height: 19px; }
QProgressBar::chunk { background: #91b7d3; border: none; }
QProgressBar::chunk:disabled { background: #c4cdd4; }
QCheckBox, QRadioButton { spacing: 7px; }
QCheckBox::indicator, QAbstractItemView::indicator { width: 15px; height: 15px;
    border: 1px solid #95a3ae; border-radius: 2px; background: #ffffff; }
QCheckBox::indicator:hover, QAbstractItemView::indicator:hover { border-color: #245d85; background: #edf3f8; }
QCheckBox::indicator:checked, QAbstractItemView::indicator:checked { background: #245d85;
    border-color: #245d85; image: url("@ICONS@/check.svg"); }
QCheckBox::indicator:indeterminate, QAbstractItemView::indicator:indeterminate { background: #245d85; image: url("@ICONS@/minus.svg"); }
QCheckBox::indicator:disabled, QAbstractItemView::indicator:disabled { background: #edf0f2; border-color: #cbd2d8; }
QCheckBox::indicator:checked:disabled, QAbstractItemView::indicator:checked:disabled { image: url("@ICONS@/check-disabled.svg"); }
QCheckBox:focus, QRadioButton:focus { outline: 1px solid #245d85; }
QRadioButton::indicator { width: 16px; height: 16px; border: 1px solid #95a3ae; border-radius: 9px; background: #ffffff; }
QRadioButton::indicator:checked { image: url("@ICONS@/radio.svg"); border-color: #245d85; }
QRadioButton::indicator:disabled { background: #edf0f2; border-color: #cbd2d8; }
QGroupBox { border: 1px solid #cbd2d8; margin-top: 12px; padding-top: 10px; }
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0px 4px; color: #465461; }
QToolTip { color: #28343f; background: #ffffff; border: 1px solid #bdc6ce; padding: 6px; }
QStatusBar { background: #f4f5f6; color: #596875; }
QSizeGrip { background: #f4f5f6; }
'''.replace('@ICONS@', _ICON_ROOT)


def _flat_pixmap(source, color):
    if source.isNull():
        return QPixmap()
    result = QPixmap(source.size())
    result.setDevicePixelRatio(source.devicePixelRatio())
    result.fill(Qt.transparent)
    painter = QPainter(result)
    painter.drawPixmap(0, 0, source)
    painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
    painter.fillRect(result.rect(), QColor(color))
    painter.end()
    return result


def _flat_icon(source):
    result = QIcon()
    sizes = source.availableSizes() or []
    # Some platform engines only expose scalable icons without availableSizes.
    sizes = sizes or [QSize(n, n) for n in (16, 24, 32, 48)]
    for size in sizes:
        for mode, color in ((QIcon.Normal, '#465461'), (QIcon.Active, BLUE),
                            (QIcon.Selected, BLUE), (QIcon.Disabled, '#89949e')):
            for state in (QIcon.Off, QIcon.On):
                pixmap = source.pixmap(size, mode, state)
                result.addPixmap(_flat_pixmap(pixmap, color), mode, state)
    return result


def _combo_container(widget):
    return (widget is not None
            and widget.metaObject().className() == 'QComboBoxPrivateContainer'
            and isinstance(widget.parentWidget(), QComboBox))


class _ComboDelegate(QStyledItemDelegate):
    """Only combo rows: quiet current/hover without QSS's blue focus rectangle."""
    def sizeHint(self, option, index):
        size = super().sizeHint(option, index)
        size.setHeight(max(30, option.fontMetrics.height() + 12))
        size.setWidth(size.width() + 24)
        return size

    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        current = bool(opt.state & (QStyle.State_Selected | QStyle.State_HasFocus))
        hover = bool(opt.state & QStyle.State_MouseOver)
        painter.save()
        painter.fillRect(opt.rect, QColor('#f0f3f6' if current or hover else '#ffffff'))
        if current:
            painter.fillRect(opt.rect.left() + 4, opt.rect.top() + 8,
                             1, max(1, opt.rect.height() - 16), QColor('#738493'))
        opt.state &= ~(QStyle.State_Selected | QStyle.State_HasFocus | QStyle.State_MouseOver)
        opt.backgroundBrush = QBrush(Qt.NoBrush)
        opt.rect = opt.rect.adjusted(12, 0, -12, 0)
        # The style still handles icons, check states, disabled text and elision.
        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.CE_ItemViewItem, opt, painter, opt.widget)
        painter.restore()


class _ComboTheme(QObject):
    def eventFilter(self, watched, event):
        if event.type() == QEvent.Polish and isinstance(watched, QComboBox):
            view = watched.view()
            if not isinstance(view.itemDelegate(), _ComboDelegate):
                view.setItemDelegate(_ComboDelegate(view))
            view.setMouseTracking(True)
            self._prepare_container(view.window())
        if _combo_container(watched):
            if event.type() == QEvent.Polish:
                self._prepare_container(watched)
            elif event.type() == QEvent.Show:
                combo = watched.parentWidget()
                # Qt finishes sizing/placing after Show. BoundedComboBox also
                # applies this immediately after its native showPopup returns.
                QTimer.singleShot(0, watched, lambda: position_combo_popup(
                    combo, isinstance(combo, BoundedComboBox)))
        return False

    @staticmethod
    def _prepare_container(container):
        if not _combo_container(container) or container.property('_flatComboPopup'):
            return
        container.setProperty('_flatComboPopup', True)
        # The container can receive Polish before Qt attaches it to the combo.
        # Prepare it explicitly from the combo too, before the native window maps.
        container.setFrameStyle(QFrame.NoFrame)
        container.setStyleSheet('QComboBoxPrivateContainer { '
                               'background: #ffffff; border: 1px solid #d5dbe0; }')
        container.setContentsMargins(1, 1, 1, 1)
        palette = container.palette()
        for group in (QPalette.Active, QPalette.Inactive, QPalette.Disabled):
            for role in (QPalette.Window, QPalette.Base):
                palette.setColor(group, role, QColor('#ffffff'))
        container.setPalette(palette)
        container.setWindowFlags(container.windowFlags() | Qt.NoDropShadowWindowHint
                                 | Qt.FramelessWindowHint)


class FlatStyle(QProxyStyle):
    """Keep native glyph shapes while removing platform yellow/color shading."""
    def __init__(self):
        super().__init__('Fusion')

    def styleHint(self, hint, option=None, widget=None, returnData=None):
        if hint == QStyle.SH_ComboBox_Popup:
            # Fusion otherwise aligns the selected menu row over the edit field.
            # Let QComboBox place a normal list below (or above at the screen edge).
            return 0
        return super().styleHint(hint, option, widget, returnData)

    def pixelMetric(self, metric, option=None, widget=None):
        if metric == QStyle.PM_ToolBarExtensionExtent:
            # Fusion's narrow default clips a padded extension toolbutton.
            return 28
        return super().pixelMetric(metric, option, widget)

    def drawPrimitive(self, element, option, painter, widget=None):
        if element == QStyle.PE_PanelMenu and _combo_container(widget):
            painter.save()
            painter.fillRect(option.rect, QColor('#ffffff'))
            painter.setPen(QColor('#d5dbe0'))
            painter.drawRect(option.rect.adjusted(0, 0, -1, -1))
            painter.restore()
            return
        super().drawPrimitive(element, option, painter, widget)

    def standardIcon(self, standard_icon, option=None, widget=None):
        return _flat_icon(super().standardIcon(standard_icon, option, widget))

    def standardPixmap(self, standard_pixmap, option=None, widget=None):
        disabled = option is not None and not option.state & QStyle.State_Enabled
        return _flat_pixmap(super().standardPixmap(standard_pixmap, option, widget),
                            '#89949e' if disabled else '#465461')


def apply_theme(app=None):
    """Theme application widgets while leaving file pickers to the operating system."""
    app = app or QApplication.instance()
    if app is None:
        raise RuntimeError('Create QApplication before applying the theme')
    # Use Windows' own open/save/folder dialogs; do not recolor their file icons.
    app.setAttribute(Qt.AA_DontUseNativeDialogs, False)
    app.setStyle(FlatStyle())
    app.setPalette(theme_palette())
    if not hasattr(app, '_combo_theme'):
        app._combo_theme = _ComboTheme(app)
        app.installEventFilter(app._combo_theme)
    app.setFont(QFont('Microsoft YaHei UI', 10))
    app.setStyleSheet(STYLE)
