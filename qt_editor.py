"""Native model/view visual editor, with transactional background structure operations."""
import tempfile
from pathlib import Path
from copy import deepcopy
from PySide6.QtCore import Qt, QMimeData, Signal, QEvent, QTimer
from PySide6.QtGui import QKeySequence, QShortcut, QDrag
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QComboBox,
    QTableView, QFileDialog, QMessageBox, QInputDialog, QAbstractItemView,
    QDialogButtonBox, QFormLayout, QLineEdit, QMenu, QSpinBox, QToolBar, QToolButton,
)
from openpyxl.utils.cell import range_boundaries
from qt_models import SheetModel, RowsModel
from qt_tasks import TaskController
from qt_layout import BoundedComboBox, configure_splitter, LayoutSplitter as QSplitter
from visual_fill import source_grid, drop_values
from workbook_session import WorkbookSession


def configure_sheet(view, model):
    view.setModel(model)
    view.clearSpans()
    view.setWordWrap(False)
    view.horizontalHeader().setDefaultSectionSize(150)
    view.verticalHeader().setDefaultSectionSize(32)
    # Only explicit sizes are applied: never allocate one widget per cell.
    for ref in model.info.get('merges', []):
        l, t, r, b = range_boundaries(ref)
        if b > t or r > l:
            view.setSpan(t-1, l-1, b-t+1, r-l+1)
    for col, width in model.info.get('column_pixels', {}).items():
        if col <= model.columnCount():
            view.setColumnWidth(col-1, max(30, min(1790, int(width))))
    for row, height in model.info.get('heights', {}).items():
        if row <= model.rowCount():
            view.setRowHeight(row-1, max(24, min(546, int(height))))


def selected_bounds(view):
    ranges = view.selectionModel().selection()
    if ranges.isEmpty():
        raise ValueError('请先选择表格区域。')
    # Qt may represent one rectangle as several ranges, particularly at spans.
    rects = [(x.left()+1, x.top()+1, x.right()+1, x.bottom()+1) for x in ranges]
    from visual_fill import expand_selection
    bounds = (
        min(x[0] for x in rects), min(x[1] for x in rects),
        max(x[2] for x in rects), max(x[3] for x in rects))
    for ref in view.model().info.get('merges', []):
        l, t, r, b = range_boundaries(ref)
        if any(a <= r and c >= l and d >= t and e <= b for a, e, c, d in rects):
            rects.append((l, t, r, b))
    l, t, r, b = bounds
    # Check interval coverage by horizontal strips, without enumerating cells.
    edges = sorted({t, b+1} | {y for a, top, c, bottom in rects for y in (top, bottom+1)})
    for top, bottom in zip(edges, edges[1:]):
        if top < t or top > b:
            continue
        end = l
        for left, right in sorted((a, c) for a, y, c, z in rects if y <= top <= z):
            if left > end:
                break
            end = max(end, right+1)
        if end <= r:
            raise ValueError('请选择一个连续矩形区域。')
    return expand_selection(view.model().info, bounds)


class SheetPreview(QDialog):
    def __init__(self, parent, info, title='表格预览'):
        super().__init__(parent)
        self.info = info
        self.setWindowTitle(title)
        self.resize(1000, 700)
        layout = QVBoxLayout(self)
        row = QHBoxLayout()
        row.addWidget(QLabel('缩放（仅影响预览）'))
        self.zoom = QSpinBox()
        self.zoom.setRange(25, 250)
        self.zoom.setSuffix('%')
        self.zoom.setValue(100)
        row.addWidget(self.zoom)
        row.addStretch()
        layout.addLayout(row)
        self.view = QTableView()
        layout.addWidget(self.view, 1)
        self.zoom.valueChanged.connect(self.rescale)
        self.rescale(100)

    def rescale(self, percent):
        scale = percent / 100
        info = dict(self.info)
        info['styles'] = {ref: {**style, 'size': max(1, style.get('size', 11) * scale)}
                          for ref, style in self.info.get('styles', {}).items()}
        configure_sheet(self.view, SheetModel(info, role='preview'))
        self.view.horizontalHeader().setDefaultSectionSize(max(15, int(150 * scale)))
        self.view.verticalHeader().setDefaultSectionSize(max(10, int(32 * scale)))
        for col, width in info.get('column_pixels', {}).items():
            self.view.setColumnWidth(col-1, max(15, int(width * scale)))
        for row, height in info.get('heights', {}).items():
            self.view.setRowHeight(row-1, max(10, int(height * scale)))
        font = self.view.font()
        font.setPointSizeF(max(1, 10 * scale))
        self.view.setFont(font)


class TransferTable(QTableView):
    regionDropped = Signal(object, int, int)
    mime_type = 'application/x-replace-simple-region'

    def __init__(self, parent=None):
        super().__init__(parent)
        self.allowed_source = None
        self.drop_handler = None
        self._active_drag = None
        self.setDefaultDropAction(Qt.CopyAction)

    def startDrag(self, actions):
        import json
        if self.model().role != 'source':
            return
        try:
            bounds = selected_bounds(self)
        except ValueError:
            return
        mime = QMimeData()
        mime.setData(self.mime_type, json.dumps(bounds).encode('utf-8'))
        drag = QDrag(self)
        drag.setMimeData(mime)
        self._active_drag = (mime, self.model(), bounds)
        try:
            drag.exec(Qt.CopyAction, Qt.CopyAction)
        finally:
            self._active_drag = None
            drag.deleteLater()

    def _drop_bounds(self, event):
        import json
        source = event.source()
        if (source is None or source is not self.allowed_source or source is self
                or not self.isEnabled() or not source.isEnabled()
                or not event.possibleActions() & Qt.CopyAction):
            return None
        active = source._active_drag
        if (not active or active[0] is not event.mimeData() or active[1] is not source.model()
                or source.model().role != 'source' or not event.mimeData().hasFormat(self.mime_type)):
            return None
        try:
            bounds = json.loads(bytes(event.mimeData().data(self.mime_type)))
            if (not isinstance(bounds, list) or len(bounds) != 4
                    or not all(type(x) is int for x in bounds) or tuple(bounds) != tuple(active[2])):
                return None
            from visual_fill import expand_selection
            return expand_selection(source.model().info, bounds)
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _accept_copy(event):
        # DragOnly may propose MoveAction; the protocol always copies.
        event.setDropAction(Qt.CopyAction)
        event.accept()

    def dragEnterEvent(self, event):
        if self._drop_bounds(event) is not None:
            self._accept_copy(event)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self._drop_bounds(event) is not None and self.indexAt(event.position().toPoint()).isValid():
            self._accept_copy(event)
        else:
            event.ignore()

    def dropEvent(self, event):
        bounds = self._drop_bounds(event)
        index = self.indexAt(event.position().toPoint())
        if bounds is not None and index.isValid() and self.drop_handler:
            row, column = index.row()+1, index.column()+1
            if self.drop_handler(bounds, row, column):
                self.regionDropped.emit(bounds, row, column)
                self._accept_copy(event)
                return
        event.ignore()


class VisualEditor(QDialog):
    def __init__(self, parent, session, items, stamps):
        super().__init__(parent)
        self.session, self.items, self.stamps = session, items, dict(stamps)
        self.scratch = tempfile.TemporaryDirectory(prefix='replace-qt-')
        self.serial = 0
        self.last_recipe = None
        self.setWindowTitle('可视化编辑 · 另存为保护模板')
        self.resize(1260, 800)
        self.setMinimumSize(860, 560)
        self.tasks = TaskController(self)
        self.tasks.error.connect(self.error)
        self.tasks.busyChanged.connect(self.set_busy)
        layout = QVBoxLayout(self)
        self.body = QWidget()
        body = QVBoxLayout(self.body)
        # Native toolbars expose overflow actions instead of widening the dialog.
        bar = QToolBar('编辑操作')
        bar.setMovable(False)
        for text, callback in [('撤销', self.undo), ('重做', self.redo), ('编辑文字', self.edit_text),
                                ('清空选区', self.clear_cells), ('映射填表…', self.mapping), ('另存为…', self.save)]:
            bar.addAction(text, callback)
        body.addWidget(bar)
        structures = QToolBar('结构与排版')
        structures.setMovable(False)
        for text, action in [('上方插行', 'insert_above'), ('下方插行', 'insert_below'),
                             ('左侧插列', 'insert_left'), ('右侧插列', 'insert_right'),
                             ('删除选中行', 'delete_rows'), ('合并选区', 'merge'), ('取消合并', 'unmerge')]:
            structures.addAction(text, lambda checked=False, a=action: self.structure(a))
        menu = QMenu('排版与配方…', structures)
        for text, callback in [('设置行高', lambda: self.resize_cells('height')),
                               ('设置列宽', lambda: self.resize_cells('width')),
                               ('自动行高', self.auto_height), ('保存整套操作配方…', self.save_recipe),
                               ('载入并应用配方…', self.load_recipe), ('复用旧版已存配方…', self.load_saved_recipe), ('弹出目标表格…', self.popup_target)]:
            menu.addAction(text, callback)
        structures.addAction(menu.menuAction())
        structures.widgetForAction(menu.menuAction()).setPopupMode(QToolButton.InstantPopup)
        body.addWidget(structures)
        self.splitter = split = QSplitter(Qt.Horizontal)
        source = QWidget()
        left = QVBoxLayout(source)
        left.setContentsMargins(0, 0, 0, 0)
        left.addWidget(QLabel('来源表格'))
        self.source_choice = BoundedComboBox()
        self.source_choice.setAccessibleName('来源表格')
        self.source_choice.addItems([f'{x.filename} · 表{x.table_index} · {x.section}' for x in items])
        self.source_choice.currentIndexChanged.connect(self.source_changed)
        left.addWidget(self.source_choice)
        self.source_view = TransferTable()
        self.source_view.setToolTip('框选来源内容后填入目标位置，也可拖动选区到目标表格。')
        self.source_view.setDragEnabled(True)
        self.source_view.setDragDropMode(QAbstractItemView.DragOnly)
        self.source_view.setSelectionMode(QAbstractItemView.ContiguousSelection)
        left.addWidget(self.source_view, 1)
        transfer = QPushButton('将来源选区填入目标选中位置')
        transfer.setEnabled(bool(items))
        transfer.clicked.connect(self.transfer)
        left.addWidget(transfer)
        split.addWidget(source)
        target = QWidget()
        right = QVBoxLayout(target)
        right.setContentsMargins(0, 0, 0, 0)
        right.addWidget(QLabel('目标工作表'))
        self.sheet = BoundedComboBox()
        self.sheet.setAccessibleName('目标工作表')
        self.sheet.addItems(session.sheets)
        self.sheet.currentTextChanged.connect(self.target_changed)
        right.addWidget(self.sheet)
        self.target = TransferTable()
        self.target.setAcceptDrops(True)
        self.target.setDragDropMode(QAbstractItemView.DropOnly)
        self.target.allowed_source = self.source_view
        self.target.drop_handler = self.transfer_values
        self.target.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self._header_gesture = None
        for header, axis in ((self.target.horizontalHeader(), 'widths'),
                             (self.target.verticalHeader(), 'heights')):
            header.viewport().installEventFilter(self)
            header.sectionResized.connect(lambda index, old, new, a=axis: self.header_resized(a, index, new))
        right.addWidget(self.target, 1)
        popup = QPushButton('弹出目标表格…')
        popup.clicked.connect(self.popup_target)
        right.addWidget(popup)
        split.addWidget(target)
        configure_splitter(split)
        body.addWidget(split, 1)
        layout.addWidget(self.body, 1)
        self.status = QLabel('双击或按 F2 编辑。浅蓝色表示待写入；原文单元格颜色保留。公式格受保护。')
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        self.source_changed()
        self.target_changed()
        self.shortcuts = []
        for key, callback in [(QKeySequence.Undo, self.undo), (QKeySequence.Redo, self.redo),
                              (QKeySequence.SaveAs, self.save)]:
            shortcut = QShortcut(key, self)
            shortcut.activated.connect(callback)
            self.shortcuts.append(shortcut)

    def eventFilter(self, watched, event):
        if hasattr(self, '_header_gesture'):
            headers = (self.target.horizontalHeader().viewport(), self.target.verticalHeader().viewport())
            if watched in headers:
                if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                    self._header_gesture = (self.sheet.currentText(), {}, {})
                elif event.type() == QEvent.MouseButtonRelease and self._header_gesture is not None:
                    QTimer.singleShot(0, self.commit_header_resize)
        return super().eventFilter(watched, event)

    def header_resized(self, axis, index, pixels):
        if self._header_gesture is None:
            return
        values = self._header_gesture[1 if axis == 'heights' else 2]
        values[index+1] = (round(max(1, min(409, pixels * 3/4)), 2) if axis == 'heights'
                           else round(max(1, min(255, (pixels-5)/7)), 3))

    def commit_header_resize(self):
        gesture, self._header_gesture = self._header_gesture, None
        if gesture is None:
            return
        name, heights, widths = gesture
        if not heights and not widths:
            return
        try:
            if self.tasks.busy or name != self.sheet.currentText():
                return
            self.session.layout(name, heights=heights, widths=widths)
        except ValueError as exc:
            self.error(exc)
        finally:
            self.target_changed()

    def set_busy(self, busy):
        self.body.setEnabled(not busy)
        if busy:
            self.status.setText('正在后台核验、解析和写入，请稍候…')

    def error(self, message):
        self.status.setText(str(message))
        QMessageBox.warning(self, '无法完成操作', str(message))

    def source_changed(self, *_):
        index = self.source_choice.currentIndex()
        self.source_info = source_grid(self.items[index].table) if index >= 0 else None
        self.source_view.setDragEnabled(self.source_info is not None)
        configure_sheet(self.source_view, SheetModel(self.source_info, role='source' if self.source_info else 'preview'))

    def target_changed(self, *_):
        name = self.sheet.currentText()
        if name not in self.session.sheets:
            return
        model = SheetModel(self.session.view_info(name), self.session, name, self.error, role='target')
        configure_sheet(self.target, model)
        model.dataChanged.connect(self.update_status)
        self.update_status()

    def update_status(self, *_):
        sizes = sum(len(x['heights']) + len(x['widths']) for x in self.session.layouts.values())
        self.status.setText(f'待保存 {sum(map(len, self.session.pending.values()))} 个单元格；{sizes} 项排版；{self.session.structure_count} 次结构编辑。')

    def undo(self):
        if not self.tasks.busy:
            self.session.undo()
            self.target_changed()

    def redo(self):
        if not self.tasks.busy:
            self.session.redo()
            self.target_changed()

    def edit_text(self):
        index = self.target.currentIndex()
        if not index.isValid():
            return self.error('请先选择目标单元格。')
        model = self.target.model()
        text, ok = QInputDialog.getMultiLineText(self, '编辑单元格', '文字（留空清除，公式格不能覆盖）', model.data(index))
        if ok:
            model.setData(index, text)

    def clear_cells(self):
        try:
            bounds = selected_bounds(self.target)
            model = self.target.model()
            l, t, r, b = bounds
            values = {}
            for row in range(t-1, b):
                for col in range(l-1, r):
                    ref = model.ref(model.index(row, col))
                    if ref not in model.info.get('formulas', {}):
                        values[ref] = ''
            self.session.edit(self.sheet.currentText(), values)
            self.target_changed()
        except ValueError as exc:
            self.error(exc)

    def transfer(self):
        try:
            bounds = selected_bounds(self.source_view)
            l, t, _, _ = selected_bounds(self.target)
            self.transfer_values(bounds, t, l)
        except ValueError as exc:
            self.error(exc)

    def transfer_values(self, bounds, row, column):
        if self.tasks.busy or not self.source_info:
            return False
        try:
            name = self.sheet.currentText()
            values = drop_values(self.source_info, bounds, self.session.sheets[name], row, column)
            if not values:
                raise ValueError('来源选区没有非空内容。')
            self.verify_sources()
            self.session.verify()
            self.session.edit(name, values, {'kind': 'drag', 'source': self.source_choice.currentIndex(),
                                            'sheet': name, 'bounds': list(bounds), 'row': row, 'column': column})
            item = self.items[self.source_choice.currentIndex()]
            from review_mapping import layout_signature
            self.last_recipe = {'version': 1, 'source_filename': item.filename, 'table_index': item.table_index,
                                'source_signature': layout_signature(self.source_info, {'header': 1}),
                                'target_signature': layout_signature(self.session.sheets[name], self.session.sheets[name]['guess']),
                                'bounds': list(bounds), 'sheet': name, 'row': row, 'column': column}
            self.target_changed()
            return True
        except ValueError as exc:
            self.error(exc)
            return False

    def resize_cells(self, dimension):
        if self.tasks.busy:
            return
        try:
            l, t, r, b = selected_bounds(self.target)
            height = dimension == 'height'
            value, ok = QInputDialog.getDouble(self, '设置行高' if height else '设置列宽',
                                               '磅（1–409）' if height else '字符宽度（1–255）',
                                               24 if height else 20, 1, 409 if height else 255, 2)
            if ok:
                values = {i: value for i in range(t if height else l, (b if height else r)+1)}
                self.session.layout(self.sheet.currentText(), **{'heights' if height else 'widths': values})
                self.target_changed()
        except ValueError as exc:
            self.error(exc)

    def auto_height(self):
        if self.tasks.busy:
            return
        try:
            _, t, _, b = selected_bounds(self.target)
            name = self.sheet.currentText()
            info, values = self.session.view_info(name), dict(self.session.pending.get(name, {}))
            from worksheet_layout import auto_heights
            def done(result):
                heights, wrapped = result
                self.session.layout(name, heights=heights, wrap_cells=wrapped)
                self.target_changed()
            self.tasks.submit(lambda p: auto_heights(info, values, range(t, b+1)), done)
        except ValueError as exc:
            self.error(exc)

    def popup_target(self):
        name = self.sheet.currentText()
        info = deepcopy(self.session.view_info(name))
        info['cells'].update(self.session.pending.get(name, {}))
        SheetPreview(self, info, f'目标预览 · {name}').exec()

    def save_recipe(self):
        if self.tasks.busy:
            return
        if not self.session.operations:
            return self.error('请先完成一次编辑、排版或填表。')
        path, _ = QFileDialog.getSaveFileName(self, '保存整套操作配方', '填表配方.json', 'JSON (*.json)')
        if not path:
            return
        from workbook_recipe import export_recipe, save_recipe_file
        recipe = export_recipe(self.session, self.items)
        def work(progress):
            self.session.verify()
            self.verify_sources()
            save_recipe_file(path, recipe)
            return path
        self.tasks.submit(work, lambda p: self.status.setText(f'配方已保存：{p}'))

    def load_recipe(self):
        if self.tasks.busy:
            return
        path, _ = QFileDialog.getOpenFileName(self, '载入整套操作或旧版拖放配方', '', 'JSON (*.json)')
        if path:
            from workbook_recipe import read_recipe_file
            self.tasks.submit(lambda p: self.prepare_recipe(read_recipe_file(path)), self.accept_recipe)

    def load_saved_recipe(self):
        if self.tasks.busy:
            return
        from app_settings import load_settings
        profiles = load_settings().get('review_drag_profiles', {})
        if not isinstance(profiles, dict) or not profiles:
            return self.error('旧版设置中没有已保存的整套配方。')
        name, ok = QInputDialog.getItem(self, '复用旧版已存配方', '选择配方（从原模板重放，成功后可整体撤销）', list(profiles), 0, False)
        if ok:
            recipe = deepcopy(profiles[name])
            self.tasks.submit(lambda p: self.prepare_recipe(recipe), self.accept_recipe)

    def prepare_recipe(self, recipe):
        from workbook_recipe import replay_recipe
        created = []
        def path():
            output = self.temp_path()
            created.append(output)
            return output
        try:
            return replay_recipe(self.session, self.items, recipe, path, self.verify_sources)
        except Exception:
            for output in created:
                Path(output).unlink(missing_ok=True)
            raise

    def accept_recipe(self, work):
        self.session.checkpoint()
        self.session.restore(work.snapshot())
        self.target_changed()
        self.status.setText(f'已从原模板重放 {len(work.operations)} 项操作，可整体撤销；请核对预览后另存。')

    def apply_recipe(self, recipe):
        self.accept_recipe(self.prepare_recipe(recipe))

    def temp_path(self):
        self.serial += 1
        return str(Path(self.scratch.name) / f'version-{self.serial}{Path(self.session.original).suffix}')

    def structure(self, action):
        try:
            bounds = selected_bounds(self.target)
            count = 1
            if action.startswith('insert'):
                count, ok = QInputDialog.getInt(self, '插入行列', '数量', 1, 1, 1000)
                if not ok:
                    return
            if action == 'delete_rows' and QMessageBox.question(self, '删除行', '删除所选整行？可以撤销。') != QMessageBox.Yes:
                return
            output, intermediate = self.temp_path(), self.temp_path()
            name = self.sheet.currentText()
            self.tasks.submit(lambda p: self.session.structure(output, name, action, bounds, count, intermediate),
                              lambda result: self.target_changed())
        except ValueError as exc:
            self.error(exc)

    def verify_sources(self):
        if self.items:
            from word_scan import validate_scan_files
            validate_scan_files(list(dict.fromkeys(x.file_path for x in self.items)), self.stamps)

    def save(self):
        output, _ = QFileDialog.getSaveFileName(self, '另存为新文件', str(Path(self.session.original).with_stem(Path(self.session.original).stem+'_已填')),
                                                'Excel 工作簿 (*.xlsx *.xlsm)')
        if not output:
            return
        if Path(output).suffix.lower() != Path(self.session.original).suffix.lower():
            return self.error('保存格式必须与模板一致，以保留宏和原有部件。')
        def work(progress):
            self.verify_sources()
            return self.session.save(output, progress)
        self.tasks.submit(work, lambda result: self.status.setText(f'已另存：{result["output_path"]}'))

    def mapping(self):
        index = self.source_choice.currentIndex()
        if index < 0:
            return self.error('请先扫描并选择一张评分表。')
        MappingDialog(self, self.items[index], self.session, self.sheet.currentText()).exec()
        self.target_changed()

    def closeEvent(self, event):
        if self.tasks.busy:
            event.ignore()
            self.status.setText('后台任务进行中，完成后可关闭。')
            return
        if (self.session.pending or self.session.layouts or self.session.structure_count) and QMessageBox.question(
                self, '关闭编辑器', '关闭会丢弃编辑会话。已另存的文件保留，是否关闭？') != QMessageBox.Yes:
            event.ignore()
            return
        self.scratch.cleanup()
        event.accept()

    def reject(self):
        self.close()


class MappingDialog(QDialog):
    def __init__(self, parent, item, session, sheet):
        super().__init__(parent)
        self.item, self.session, self.sheet = item, session, sheet
        self.plan = None
        self.setWindowTitle('映射填表 · 先预览再回填')
        self.resize(850, 640)
        layout = QVBoxLayout(self)
        self.body = QWidget()
        body = QVBoxLayout(self.body)
        form = QFormLayout()
        self.inputs = {}
        guess = session.sheets[sheet]['guess']
        for key, label, value in [('source_header', '来源表头行', 1), ('source_factor', '来源因素列（数字）', 1),
                                  ('source_description', '来源说明列（数字）', 2), ('source_points', '来源分值列（数字）', 3),
                                  ('header', '目标表头行', guess['header']), ('start', '目标起始行', guess['start']),
                                  ('end', '目标结束行', guess['end']), ('factor', '目标因素列', guess['factor']),
                                  ('description', '目标说明列', guess['description']), ('points', '目标分值列', guess['points'])]:
            field = QLineEdit(str(value))
            self.inputs[key] = field
            field.textChanged.connect(self.invalidate)
            form.addRow(label, field)
        self.mode = QComboBox()
        self.mode.addItems(['按因素名称匹配', '按顺序填入'])
        self.mode.currentIndexChanged.connect(self.invalidate)
        form.addRow('填入方式', self.mode)
        body.addLayout(form)
        self.preview = QTableView()
        self.model = RowsModel(['因素', '目标行（0 跳过）', '状态'], editable=True)
        self.model.editable_columns = {1}
        self.preview.setModel(self.model)
        self.model.dataChanged.connect(self.invalidate)
        body.addWidget(self.preview, 1)
        buttons = QHBoxLayout()
        preview = QPushButton('核验并预览')
        preview.clicked.connect(self.build)
        self.apply_button = QPushButton('确认回填到编辑器')
        self.apply_button.setEnabled(False)
        self.apply_button.clicked.connect(self.apply)
        buttons.addWidget(preview)
        buttons.addWidget(self.apply_button)
        body.addLayout(buttons)
        layout.addWidget(self.body)
        self.message = QLabel('仅修改目标行列；目标行可手动指定。保存仍通过编辑器另存为。')
        self.message.setWordWrap(True)
        layout.addWidget(self.message)
        self.tasks = TaskController(self)
        self.tasks.busyChanged.connect(lambda busy: self.body.setEnabled(not busy))
        self.tasks.error.connect(lambda text: self.message.setText(text))

    def invalidate(self, *_):
        self.plan = None
        if hasattr(self, 'apply_button'):
            self.apply_button.setEnabled(False)

    def build(self):
        try:
            from review_table_filler import extract_review_rules
            from review_mapping import plan_fill
            values = {key: field.text().strip() for key, field in self.inputs.items()}
            source = {key: int(values['source_'+key]) for key in ('header', 'factor', 'description', 'points')}
            config = {key: values[key] for key in ('header', 'start', 'end', 'factor', 'description', 'points')}
            config.update(sheet=self.sheet, mode='match' if self.mode.currentIndex() == 0 else 'sequence')
            overrides = {i: int(row[1]) for i, row in enumerate(self.model.rows) if str(row[1]).strip()}
            def work(progress):
                self.session.verify()
                rules = extract_review_rules(self.item.table, source)
                plan = plan_fill(self.session.path, rules, config, overrides)
                plan['recipe_configs'] = [source, config]
                plan['recipe_overrides'] = overrides
                return plan
            self.tasks.submit(work, self.show_plan)
        except ValueError as exc:
            self.message.setText(str(exc))

    def show_plan(self, plan):
        self.plan = plan
        self.model.replace([[e['name'], '' if e['row'] is None else str(e['row']), e['status']] for e in plan['entries']])
        self.preview.resizeColumnsToContents()
        self.message.setText('\n'.join(plan['errors']) if plan['errors'] else f'已核验 {len(plan["changes"])} 个变更。')
        self.apply_button.setEnabled(not plan['errors'] and bool(plan['changes']))

    def apply(self):
        if self.plan and not self.plan['errors']:
            try:
                self.parent().verify_sources()
                self.session.verify()
                self.session.edit(self.sheet, {c['cell']: c['new'] for c in self.plan['changes']},
                                  {'kind': 'auto', 'source': self.parent().items.index(self.item),
                                   'configs': self.plan['recipe_configs'], 'overrides': self.plan['recipe_overrides']})
                self.accept()
            except ValueError as exc:
                self.message.setText(str(exc))

    def reject(self):
        if not self.tasks.busy:
            super().reject()

    def closeEvent(self, event):
        event.ignore() if self.tasks.busy else event.accept()
