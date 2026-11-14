"""Two-column rules with the spreadsheet editing shortcuts of the legacy UI."""
import csv
import io
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeySequence
from PySide6.QtWidgets import QApplication, QTableView, QMenu
from qt_models import RowsModel


class RuleModel(RowsModel):
    def __init__(self, headers, rows=(), parent=None):
        super().__init__(headers, [list(row) for row in rows], editable=True, parent=parent)
        self.history, self.future = [], []

    def checkpoint(self):
        self.history.append([list(row) for row in self.rows])
        del self.history[:-100]
        self.future.clear()

    def replace(self, rows):
        rows = [list(row) for row in rows]
        if rows != self.rows:
            self.checkpoint()
            super().replace(rows)

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid() or role != Qt.EditRole:
            return False
        if self.rows[index.row()][index.column()] != str(value):
            self.checkpoint()
        return super().setData(index, value, role)

    def append(self, row):
        self.checkpoint()
        super().append(row)

    def remove(self, numbers):
        numbers = list(numbers)
        if numbers:
            self.checkpoint()
            super().remove(numbers)

    def undo(self):
        if self.history:
            self.future.append([list(row) for row in self.rows])
            super().replace(self.history.pop())

    def redo(self):
        if self.future:
            self.history.append([list(row) for row in self.rows])
            super().replace(self.future.pop())


class RuleTable(QTableView):
    error = Signal(str)

    def copy(self):
        selected = self.selectionModel().selectedIndexes()
        if not selected:
            return
        cells = {(i.row(), i.column()): i.data() or '' for i in selected}
        rows, cols = [i.row() for i in selected], [i.column() for i in selected]
        stream = io.StringIO(newline='')
        writer = csv.writer(stream, delimiter='\t', lineterminator='\n')
        writer.writerows([[cells.get((r, c), '') for c in range(min(cols), max(cols)+1)]
                         for r in range(min(rows), max(rows)+1)])
        QApplication.clipboard().setText(stream.getvalue())

    def paste(self):
        text = QApplication.clipboard().text()
        if not text:
            return
        try:
            values = list(csv.reader(io.StringIO(text, newline=''), delimiter='\t'))
            index = self.currentIndex()
            row, col = (index.row(), index.column()) if index.isValid() else (0, 0)
            if any(len(value) + col > 2 for value in values):
                raise ValueError('规则只有两列，请从第一列粘贴两列内容，或在当前列粘贴一列。')
            rows = [list(value) for value in self.model().rows]
            while len(rows) < row + len(values):
                rows.append(['', ''])
            for offset, value in enumerate(values):
                for column, cell in enumerate(value):
                    rows[row+offset][col+column] = cell
            self.model().replace(rows)
            self.setCurrentIndex(self.model().index(row, col))
        except (csv.Error, ValueError) as exc:
            self.error.emit(str(exc))

    def clear_cells(self):
        rows = [list(row) for row in self.model().rows]
        for index in self.selectionModel().selectedIndexes():
            rows[index.row()][index.column()] = ''
        self.model().replace(rows)

    def keyPressEvent(self, event):
        # Inline QLineEdit editors keep their own text/undo shortcuts. These
        # events belong to the table only, not a window-wide QShortcut.
        for key, callback in ((QKeySequence.Copy, self.copy), (QKeySequence.Paste, self.paste),
                              (QKeySequence.Undo, self.model().undo), (QKeySequence.Redo, self.model().redo)):
            if event.matches(key):
                callback()
                event.accept()
                return
        if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            self.clear_cells()
            event.accept()
            return
        super().keyPressEvent(event)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        for text, callback in [('复制\tCtrl+C', self.copy), ('粘贴\tCtrl+V', self.paste),
                               ('清空选区\tDelete', self.clear_cells), ('撤销\tCtrl+Z', self.model().undo),
                               ('重做\tCtrl+Y', self.model().redo)]:
            menu.addAction(text, callback)
        menu.exec(event.globalPos())
