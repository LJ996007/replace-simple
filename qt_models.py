"""Qt models: virtual cell rendering and proxy filtering, no per-cell widgets."""
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, QSortFilterProxyModel
from PySide6.QtGui import QColor, QFont
from openpyxl.utils.cell import get_column_letter
from review_mapping import _anchor
from visual_fill import _value
from qt_theme import PENDING


class RowsModel(QAbstractTableModel):
    def __init__(self, headers, rows=(), editable=False, parent=None):
        super().__init__(parent)
        self.headers, self.rows, self.editable = list(headers), list(rows), editable
        self.editable_columns = None

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.headers)

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and role in (Qt.DisplayRole, Qt.EditRole, Qt.ToolTipRole):
            value = self.rows[index.row()][index.column()]
            return '' if value is None else str(value)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return self.headers[section] if orientation == Qt.Horizontal else str(section + 1)

    def flags(self, index):
        flags = super().flags(index)
        return flags | Qt.ItemIsEditable if self.editable and (self.editable_columns is None or index.column() in self.editable_columns) else flags

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid() or not self.editable or role != Qt.EditRole or (self.editable_columns is not None and index.column() not in self.editable_columns):
            return False
        self.rows[index.row()][index.column()] = str(value)
        self.dataChanged.emit(index, index)
        return True

    def replace(self, rows):
        self.beginResetModel()
        self.rows = [list(row) for row in rows]
        self.endResetModel()

    def append(self, row):
        self.beginInsertRows(QModelIndex(), len(self.rows), len(self.rows))
        self.rows.append(list(row))
        self.endInsertRows()

    def remove(self, numbers):
        for row in sorted(set(numbers), reverse=True):
            self.beginRemoveRows(QModelIndex(), row, row)
            del self.rows[row]
            self.endRemoveRows()


class TemplateModel(RowsModel):
    """Checks belong only to the currently displayed category; resets clear them."""
    def __init__(self):
        super().__init__(['选择（全选）', '模板名称', '分类', '路径'])
        self.checked = set()

    def replace(self, rows):
        self.checked.clear()
        super().replace(rows)

    def flags(self, index):
        flags = super().flags(index)
        return flags | Qt.ItemIsUserCheckable if index.isValid() and index.column() == 0 else flags

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and index.column() == 0 and role == Qt.CheckStateRole:
            return Qt.Checked if index.row() in self.checked else Qt.Unchecked
        return super().data(index, role)

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid() or index.column() != 0 or role != Qt.CheckStateRole:
            return False
        if value in (Qt.Checked, Qt.Checked.value):
            self.checked.add(index.row())
        else:
            self.checked.discard(index.row())
        self.dataChanged.emit(index, index, [Qt.CheckStateRole])
        return True

    def toggle_all(self):
        self.checked = set() if len(self.checked) == len(self.rows) else set(range(len(self.rows)))
        if self.rows:
            self.dataChanged.emit(self.index(0, 0), self.index(len(self.rows) - 1, 0), [Qt.CheckStateRole])


class ScanModel(RowsModel):
    def __init__(self):
        super().__init__(['选用', '类型', '文件', '章节', '内容摘要'])
        self.items, self.checked = [], set()

    def load(self, tables, sections):
        self.beginResetModel()
        from dataclasses import replace
        self.items = [('table', x) for x in tables] + [('symbol', replace(x, clauses=[clause])) for x in sections for clause in x.clauses]
        self.checked = set()
        self.rows = [[ '', '表格' if kind == 'table' else '条款', item.filename,
                       item.section, item.preview if kind == 'table' else '\n'.join(c.text for c in item.clauses)]
                     for kind, item in self.items]
        self.endResetModel()

    def flags(self, index):
        result = super().flags(index)
        return result | Qt.ItemIsUserCheckable if index.column() == 0 else result

    def data(self, index, role=Qt.DisplayRole):
        if index.column() == 0 and role == Qt.CheckStateRole:
            return Qt.Checked if index.row() in self.checked else Qt.Unchecked
        return super().data(index, role)

    def setData(self, index, value, role=Qt.EditRole):
        if index.column() == 0 and role == Qt.CheckStateRole:
            if value == Qt.Checked or value == Qt.Checked.value:
                self.checked.add(index.row())
            else:
                self.checked.discard(index.row())
            self.dataChanged.emit(index, index, [Qt.CheckStateRole])
            return True
        return False


class ScanFilter(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.kinds, self.sections = set(), set()

    def filterAcceptsRow(self, row, parent):
        model = self.sourceModel()
        kind, item = model.items[row]
        return (not self.kinds or kind in self.kinds) and (not self.sections or item.section in self.sections) and super().filterAcceptsRow(row, parent)

    def select_values(self, kinds=(), sections=()):
        if hasattr(self, 'endFilterChange'):
            self.beginFilterChange()
            self.kinds, self.sections = set(kinds), set(sections)
            self.endFilterChange(QSortFilterProxyModel.Direction.Rows)
        else:  # PySide6 6.8 / 6.9
            self.kinds, self.sections = set(kinds), set(sections)
            self.invalidateFilter()


class SheetModel(QAbstractTableModel):
    def __init__(self, info=None, session=None, sheet=None, on_error=None, *, role=None):
        super().__init__()
        self.role = role or ('target' if session is not None else 'preview')
        if self.role not in ('source', 'target', 'preview'):
            raise ValueError('Unknown sheet role')
        self.info = info or {'cells': {}, 'merges': [], 'max_row': 1, 'max_col': 1}
        self.session, self.sheet, self.on_error = session, sheet, on_error

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else self.info['max_row']

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else self.info['max_col']

    def ref(self, index):
        return _anchor(f'{get_column_letter(index.column()+1)}{index.row()+1}', self.info['merges'])

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        ref = self.ref(index)
        style = self.info.get('styles', {}).get(ref, {})
        pending = self.session.pending.get(self.sheet, {}) if self.session else {}
        if role in (Qt.DisplayRole, Qt.EditRole, Qt.ToolTipRole):
            if ref in self.info.get('formulas', {}):
                return '=' + self.info['formulas'][ref]
            return str(pending.get(ref, self.info['cells'].get(ref, '')))
        if role == Qt.BackgroundRole:
            return QColor(PENDING) if ref in pending else QColor(style['bg']) if style.get('bg') else None
        if role == Qt.ForegroundRole and style.get('fg'):
            return QColor(style['fg'])
        if role == Qt.FontRole and style:
            font = QFont(style.get('family', 'Microsoft YaHei UI'), int(style.get('size', 11)))
            font.setBold(style.get('bold', False))
            font.setItalic(style.get('italic', False))
            return font

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return get_column_letter(section+1) if orientation == Qt.Horizontal else str(section+1)

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        result = super().flags(index)
        if self.role == 'source':
            result |= Qt.ItemIsDragEnabled
        if self._can_edit(index):
            result |= Qt.ItemIsEditable
        return result

    def _can_edit(self, index):
        if not index.isValid() or self.role != 'target' or not self.session or self.info.get('protected'):
            return False
        ref = self.ref(index)
        from openpyxl.utils.cell import range_boundaries
        col, row, _, _ = range_boundaries(ref)
        return ref not in self.info.get('formulas', {}) and not any(
            l <= col <= r and t <= row <= b
            for l, t, r, b in map(range_boundaries, self.info.get('formula_ranges', [])))

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or not self._can_edit(index):
            return False
        try:
            self.session.edit(self.sheet, {self.ref(index): _value(str(value))})
            self.dataChanged.emit(index, index)
            return True
        except ValueError as exc:
            if self.on_error:
                self.on_error(str(exc))
            return False
