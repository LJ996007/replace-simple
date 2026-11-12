"""Layout primitives that remain usable with long Office filenames."""
from PySide6.QtCore import Qt, QSize
from PySide6.QtWidgets import (
    QComboBox, QSizePolicy, QStyle, QStyleOptionComboBox, QStylePainter,
    QSplitter, QSplitterHandle,
)


COMBO_POPUP_GAP = 4


def position_combo_popup(combo, bounded=False):
    """Retain Qt's popup/input handling, adding a gap on either screen edge."""
    popup = combo.view().window()
    if not popup.isVisible():
        return
    screen = combo.screen().availableGeometry()
    origin = combo.mapToGlobal(combo.rect().topLeft())
    below = origin.y() + combo.height() + COMBO_POPUP_GAP
    room_below = max(0, screen.bottom() - below + 1)
    room_above = max(0, origin.y() - COMBO_POPUP_GAP - screen.top())
    height = min(popup.height(), max(room_above, room_below))
    width = min(combo.width() if bounded else popup.width(), screen.width())
    x = max(screen.left(), min(origin.x(), screen.right() - width + 1))
    y = below if height <= room_below else origin.y() - COMBO_POPUP_GAP - height
    popup.setGeometry(x, y, width, height)
    combo.view().scrollTo(combo.view().currentIndex())


class BoundedComboBox(QComboBox):
    """Keep full model text, but never size a panel from the longest item."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.setMinimumContentsLength(12)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setMinimumWidth(0)
        self.view().setTextElideMode(Qt.ElideMiddle)
        self.view().setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.currentTextChanged.connect(self.setToolTip)
        self.model().rowsInserted.connect(self._update_tooltips)
        self.model().dataChanged.connect(self._update_tooltips)
        self._updating_tooltips = False

    def showPopup(self):
        super().showPopup()
        if not self.count():
            return
        # QComboBox still measures the longest model string for its popup even
        # with an eliding view. Bound that window too, preserving native input,
        # current-index scrolling and popup mouse grabbing.
        position_combo_popup(self, bounded=True)

    def _update_tooltips(self, *_):
        if self._updating_tooltips:
            return
        self._updating_tooltips = True
        try:
            for row in range(self.count()):
                text = self.itemText(row)
                if self.itemData(row, Qt.ToolTipRole) != text:
                    self.setItemData(row, text, Qt.ToolTipRole)
            self.setToolTip(self.currentText())
        finally:
            self._updating_tooltips = False

    def paintEvent(self, event):
        painter = QStylePainter(self)
        option = QStyleOptionComboBox()
        self.initStyleOption(option)
        field = self.style().subControlRect(QStyle.CC_ComboBox, option,
                                           QStyle.SC_ComboBoxEditField, self)
        option.currentText = self.fontMetrics().elidedText(
            option.currentText, Qt.ElideMiddle, max(0, field.width() - 4))
        painter.drawComplexControl(QStyle.CC_ComboBox, option)
        painter.drawControl(QStyle.CE_ComboBoxLabel, option)


class _LayoutHandle(QSplitterHandle):
    def sizeHint(self):
        # QStyleSheetStyle adds border widths to the default vertical sizeHint,
        # while horizontal splitters use handleWidth directly. Keep both equal.
        width = self.splitter().handleWidth()
        return QSize(width, width)


class LayoutSplitter(QSplitter):
    def createHandle(self):
        return _LayoutHandle(self.orientation(), self)


def configure_splitter(splitter, minimum=260, sizes=None):
    """Call after adding children; equal stretch preserves the user's split on resize."""
    # An 8px layout slot paints a 2px line with 3px surface gutters in QSS.
    # Unlike Qt's automatic expansion for <5px handles, this hit area stays
    # inside its own slot and cannot steal table-edge or selector clicks.
    splitter.setHandleWidth(8)
    splitter.setChildrenCollapsible(False)
    for index in range(splitter.count()):
        panel = splitter.widget(index)
        if splitter.orientation() == Qt.Horizontal:
            panel.setMinimumWidth(minimum)
            cursor = Qt.SplitHCursor
        else:
            panel.setMinimumHeight(minimum)
            cursor = Qt.SplitVCursor
        splitter.setStretchFactor(index, 1)
        if index:
            splitter.handle(index).setCursor(cursor)
            splitter.handle(index).setToolTip('拖动调整两侧大小')
    splitter.setSizes(sizes or [1] * splitter.count())
