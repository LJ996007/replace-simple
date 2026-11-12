"""PySide6 desktop entry point. Run main.py --legacy for the Tk frontend."""
import os
import sys
from pathlib import Path
from PySide6.QtCore import Qt, QSortFilterProxyModel, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QTableView, QTabWidget, QFileDialog,
    QMessageBox, QPlainTextEdit, QCheckBox, QProgressBar,
    QAbstractItemView, QInputDialog, QDialog, QDialogButtonBox, QListWidget, QListWidgetItem,
)
from app_settings import load_settings, save_settings
from replacement_rules import normalize_rule_rows, validate_rules, DELETE_MARKER
from qt_models import RowsModel, ScanModel, SheetModel, ScanFilter, TemplateModel
from qt_tasks import TaskController
from qt_editor import VisualEditor, configure_sheet
from qt_layout import BoundedComboBox, configure_splitter, LayoutSplitter as QSplitter
from template_library import default_template_library_path
from symbol_clause_extractor import DEFAULT_SYMBOL_CHARS

SUPPORTED = {'.doc', '.docx', '.xls', '.xlsx', '.xlsm', '.ppt', '.pptx'}
from qt_theme import STYLE, apply_theme


def button(text, callback, primary=False):
    result = QPushButton(text)
    result.setProperty('primary', primary)
    result.clicked.connect(callback)
    return result


def toolbar(layout, actions):
    row = QHBoxLayout()
    row.setSpacing(6)
    for text, callback in actions:
        row.addWidget(button(text, callback))
    row.addStretch()
    layout.addLayout(row)
    return row


def table(model):
    view = QTableView()
    view.setModel(model)
    view.setAlternatingRowColors(True)
    view.setWordWrap(False)
    view.setSelectionBehavior(QAbstractItemView.SelectRows)
    view.setSelectionMode(QAbstractItemView.ExtendedSelection)
    view.verticalHeader().setDefaultSectionSize(34)
    view.horizontalHeader().setStretchLastSection(True)
    return view


from qt_extras import WorkbenchActions


class MainWindow(WorkbenchActions, QMainWindow):
    def __init__(self, restore=True):
        super().__init__()
        self.settings = load_settings() if restore else {}
        self.persist = restore
        from ui_common import presets_from_settings
        self.presets = presets_from_settings(self.settings)
        self.last_batch = None
        self.batch_context = None
        self.library_root = self.settings.get('template_library_root') or default_template_library_path()
        self.stamps, self.scan_symbols, self.project_details = {}, '', []
        self.setWindowTitle('replace-simple')
        self.resize(1260, 860)
        self.setMinimumSize(920, 640)
        self.setWindowIcon(QIcon(str(Path(__file__).with_name('icon-256.png'))))
        self.tasks = TaskController(self)
        self.tasks.busyChanged.connect(self.busy_changed)
        self.tasks.progress.connect(self.progress_changed)
        self.tasks.error.connect(self.error)
        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 10, 20, 14)
        layout.setSpacing(10)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, 1)
        self.build_replace()
        self.build_scan()
        self.build_library()
        self.output = QLineEdit(self.settings.get('qt_output_dir', ''))
        self.output.setPlaceholderText('请选择输出目录，批量替换将生成到此处')
        row = QHBoxLayout()
        row.addWidget(QLabel('输出目录'))
        row.addWidget(self.output, 1)
        self.choose_output_button = button('选择目录…', self.choose_output)
        row.addWidget(self.choose_output_button)
        layout.addLayout(row)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        layout.addWidget(self.progress)
        self.status = QLabel('就绪 · 添加文件和替换规则，或切换到扫描提取。')
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(3000)
        self.log.setMaximumHeight(100)
        self.log.setPlaceholderText('处理结果和错误将在这里显示。')
        layout.addWidget(self.log)
        self.log.setVisible(False)
        actions = QHBoxLayout()
        layout.addLayout(actions)
        self.log_toggle = QPushButton('处理详情')
        self.log_toggle.setCheckable(True)
        self.log_toggle.toggled.connect(self.log.setVisible)
        actions.addWidget(self.log_toggle)
        actions.addStretch()
        self.setCentralWidget(central)

    def build_replace(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 14, 0, 8)
        split = QSplitter(Qt.Vertical)
        files = QWidget()
        fl = QVBoxLayout(files)
        fl.setContentsMargins(0, 0, 0, 0)
        toolbar(fl, [('添加文件', self.add_files), ('添加文件夹', self.add_folder),
                     ('移除选中', self.remove_files), ('使用模板库', lambda: self.tabs.setCurrentIndex(2))])
        self.files = RowsModel(['文件名', '完整路径'])
        self.file_view = table(self.files)
        self.file_view.setColumnWidth(0, 280)
        fl.addWidget(self.file_view)
        split.addWidget(files)
        rules = QWidget()
        rl = QVBoxLayout(rules)
        rl.setContentsMargins(0, 8, 0, 0)
        toolbar(rl, [('新增规则', lambda: self.rules.append(['', ''])), ('删除选中规则', self.remove_rules),
                     ('标记删除原文', self.mark_delete), ('导入 Excel', self.import_rules),
                     ('导出规则', self.export_rules), ('提取项目信息…', self.extract_project)])
        self.rules = RowsModel(['原文本 / 模板占位符', '替换为（空白跳过）'], self.settings.get('qt_rules', [['', ''] for _ in range(8)]), editable=True)
        self.rule_view = table(self.rules)
        self.rule_view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.rule_view.setColumnWidth(0, 440)
        rl.addWidget(self.rule_view)
        split.addWidget(rules)
        configure_splitter(split, minimum=100, sizes=[230, 320])
        layout.addWidget(split, 1)
        row = QHBoxLayout()
        note = QLabel(f'空白替换值跳过；明确删除请使用“{DELETE_MARKER}”。')
        note.setObjectName('hint')
        row.addWidget(note, 1)
        row.addWidget(button('开始批量替换', self.replace, True))
        layout.addLayout(row)
        toolbar(layout, [('预览输出路径…', self.preview_outputs), ('常用占位符…', self.manage_presets),
                         ('仅重试失败', self.retry_failed), ('导出处理报告…', self.export_report)])
        self.tabs.addTab(page, '批量替换')

    def build_scan(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 14, 0, 8)
        toolbar(layout, [('添加 Word', self.add_scan_files), ('移除选中来源', self.remove_scan_files),
                         ('扫描表格与条款', self.scan), ('打开工作簿编辑…', self.open_editor)])
        self.scan_splitter = QSplitter(Qt.Vertical)
        self.scan_files = RowsModel(['Word 来源文件'])
        self.scan_file_view = table(self.scan_files)
        self.scan_splitter.addWidget(self.scan_file_view)
        results = QWidget()
        results_layout = QVBoxLayout(results)
        results_layout.setContentsMargins(0, 0, 0, 0)
        filters = QHBoxLayout()
        self.symbols = QLineEdit(self.settings.get('symbol_chars', DEFAULT_SYMBOL_CHARS))
        self.symbols.setMaximumWidth(130)
        self.symbols.setAccessibleName('扫描条款符号')
        filters.addWidget(QLabel('条款符号'))
        filters.addWidget(self.symbols)
        self.search = QLineEdit()
        self.search.setClearButtonEnabled(True)
        self.search.setPlaceholderText('筛选文件、章节或内容…')
        filters.addWidget(self.search, 1)
        self.keep_symbols = QCheckBox('导出保留符号')
        self.keep_symbols.setChecked(bool(self.settings.get('keep_clause_symbols', False)))
        filters.addWidget(self.keep_symbols)
        results_layout.addLayout(filters)
        split = QSplitter(Qt.Horizontal)
        # Local chrome contract: only this scan list / preview boundary is joined.
        split.setProperty('scanSurface', True)
        self.scans = ScanModel()
        self.proxy = ScanFilter(self)
        self.proxy.setSourceModel(self.scans)
        self.proxy.setFilterKeyColumn(-1)
        self.proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.scan_view = table(self.proxy)
        self.scan_view.setProperty('scanEdge', 'results')
        self.scan_view.setColumnWidth(0, 52)
        self.scan_view.setColumnWidth(1, 60)
        self.scan_view.setColumnWidth(2, 180)
        self.scan_view.setColumnWidth(3, 160)
        self.scan_view.selectionModel().currentChanged.connect(self.preview_scan)
        self.filter_timer = QTimer(self)
        self.filter_timer.setSingleShot(True)
        self.filter_timer.setInterval(150)
        self.filter_timer.timeout.connect(lambda: self.proxy.setFilterFixedString(self.search.text()))
        self.search.textChanged.connect(lambda: self.filter_timer.start())
        split.addWidget(self.scan_view)
        self.preview_stack = QTabWidget()
        self.preview_stack.setProperty('scanSurface', True)
        self.scan_preview = QTableView()
        self.scan_preview.setProperty('scanEdge', 'preview')
        self.scan_preview.setModel(SheetModel())
        self.preview_text = QPlainTextEdit()
        self.preview_text.setProperty('scanEdge', 'preview')
        self.preview_text.setReadOnly(True)
        self.preview_text.setPlaceholderText('选择扫描结果后显示完整文本。')
        self.preview_stack.addTab(self.scan_preview, '表格预览')
        self.preview_stack.addTab(self.preview_text, '完整文本')
        split.addWidget(self.preview_stack)
        configure_splitter(split)
        results_layout.addWidget(split, 1)
        toolbar(results_layout, [('勾选当前筛选结果', lambda: self.check_visible(True)),
                                 ('取消当前筛选勾选', lambda: self.check_visible(False)), ('筛选…', self.filter_scans),
                                 ('弹出表格…', self.popup_scan), ('导出勾选结果', self.export_scan)])
        self.scan_splitter.addWidget(results)
        configure_splitter(self.scan_splitter, minimum=85, sizes=[110, 360])
        results.setMinimumHeight(210)
        self.scan_splitter.setStretchFactor(0, 0)
        self.scan_splitter.setStretchFactor(1, 1)
        layout.addWidget(self.scan_splitter, 1)
        self.tabs.addTab(page, '扫描提取与编辑')

    def build_library(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 14, 0, 8)
        row = QHBoxLayout()
        self.library_path = QLineEdit(self.library_root)
        self.library_path.setReadOnly(True)
        row.addWidget(QLabel('模板库目录'))
        row.addWidget(self.library_path, 1)
        row.addWidget(button('更改目录…', self.change_library))
        layout.addLayout(row)
        self.category = BoundedComboBox()
        self.category.setAccessibleName('模板分类')
        self.category.currentTextChanged.connect(self.refresh_templates)
        row = toolbar(layout, [('刷新', self.refresh_library), ('新建分类', self.new_category),
                               ('重命名分类', self.rename_category), ('删除空分类', self.delete_category)])
        row.insertWidget(0, self.category)
        self.templates = TemplateModel()
        self.template_view = table(self.templates)
        self.template_view.setColumnWidth(0, 120)
        self.template_view.setColumnWidth(1, 340)
        header = self.template_view.horizontalHeader()
        header.setSectionsClickable(True)
        header.sectionClicked.connect(lambda section: self.templates.toggle_all() if section == 0 else None)
        header.setToolTip('点击“选择”表头，全选或清空当前分类的模板。')
        layout.addWidget(self.template_view, 1)
        self.template_selection = QLabel()
        self.template_selection.setObjectName('hint')
        layout.addWidget(self.template_selection)
        actions = toolbar(layout, [('添加模板…', self.add_templates)])
        self.delete_templates_button = button('删除勾选模板', self.delete_templates)
        self.use_templates_button = button('使用勾选模板', self.use_templates)
        actions.insertWidget(1, self.delete_templates_button)
        actions.insertWidget(2, self.use_templates_button)
        self.templates.dataChanged.connect(self.update_template_selection)
        self.templates.modelReset.connect(self.update_template_selection)
        self.update_template_selection()
        hint = QLabel('模板添加时复制入库；批量替换输出目录必须位于库外。')
        hint.setObjectName('hint')
        layout.addWidget(hint)
        self.library_items = []
        self.visible_templates = []
        self.tabs.addTab(page, '模板库')
        self.tabs.currentChanged.connect(lambda index: self.refresh_library() if index == 2 else None)

    def error(self, text):
        self.status.setText(str(text))
        self.log.appendPlainText(str(text))
        QMessageBox.warning(self, '操作未完成', str(text))

    def busy_changed(self, busy):
        self.tabs.setEnabled(not busy)
        self.output.setEnabled(not busy)
        self.choose_output_button.setEnabled(not busy)
        self.progress.setRange(0, 0 if busy else 100)
        if busy:
            self.status.setText('正在后台处理…')
        else:
            self.progress.setValue(100)

    def progress_changed(self, args):
        if len(args) >= 2:
            self.progress.setRange(0, max(1, int(args[1])))
            self.progress.setValue(int(args[0]))
        self.status.setText(' · '.join(str(x) for x in args))

    def result(self, value):
        import pprint
        self.log.appendPlainText(pprint.pformat(value, width=110, sort_dicts=False))
        self.log_toggle.setChecked(True)
        self.status.setText('处理结束，详细结果见下方。失败项不会被当作成功。')

    def choose_output(self):
        path = QFileDialog.getExistingDirectory(self, '选择输出目录', self.output.text())
        if path:
            self.output.setText(path)

    def append_files(self, paths):
        existing = {os.path.normcase(os.path.abspath(row[1])) for row in self.files.rows}
        for path in paths:
            path = os.path.abspath(path)
            key = os.path.normcase(path)
            if key not in existing and Path(path).suffix.lower() in SUPPORTED and not Path(path).name.startswith('~$'):
                self.files.append([Path(path).name, path])
                existing.add(key)

    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, '选择 Office 文件', '', 'Office 文件 (*.doc *.docx *.xls *.xlsx *.xlsm *.ppt *.pptx)')
        self.append_files(paths)

    def add_folder(self):
        path = QFileDialog.getExistingDirectory(self, '添加文件夹（含子目录）')
        if path:
            self.tasks.submit(lambda p: [str(x) for x in Path(path).rglob('*') if x.is_file() and x.suffix.lower() in SUPPORTED], self.append_files)

    def selected_rows(self, view):
        return sorted({index.row() for index in view.selectionModel().selectedIndexes()})

    def remove_files(self):
        self.files.remove(self.selected_rows(self.file_view))

    def remove_rules(self):
        self.rules.remove(self.selected_rows(self.rule_view))

    def mark_delete(self):
        for row in self.selected_rows(self.rule_view):
            self.rules.setData(self.rules.index(row, 1), DELETE_MARKER)

    def import_rules(self):
        path, _ = QFileDialog.getOpenFileName(self, '导入规则', '', 'Excel (*.xlsx *.xls)')
        if path:
            from string_replacer import load_replacement_rules
            self.tasks.submit(lambda p: load_replacement_rules(path), self.rules.replace)
            self.files.remove([i for i, row in enumerate(self.files.rows) if os.path.normcase(row[1]) == os.path.normcase(os.path.abspath(path))])

    def export_rules(self):
        path, _ = QFileDialog.getSaveFileName(self, '导出规则', '替换规则.xlsx', 'Excel (*.xlsx)')
        if path:
            from string_replacer import save_replacement_rules
            rows = deepcopy_rows(self.rules.rows)
            self.tasks.submit(lambda p: save_replacement_rules(path, rows), lambda result: self.status.setText('规则已导出。'))

    def replace(self):
        try:
            validate_rules(self.rules.rows)
            rules = [(a, b) for a, b in normalize_rule_rows(self.rules.rows) if b]
            files = [row[1] for row in self.files.rows]
            output = self.output.text().strip()
            if not files or not rules:
                raise ValueError('请添加文件并填写有效替换值。')
            if not output:
                raise ValueError('请选择输出目录。')
            root = Path(self.library_root).resolve()
            if Path(output).resolve().is_relative_to(root):
                raise ValueError('输出目录不能位于模板库内，请选择库外目录。')
            # The original engine overwrites an unchanged filename in-place. Make that explicit.
            if any(Path(path).resolve().parent == Path(output).resolve() for path in files):
                if QMessageBox.question(self, '输出到来源目录', '输出目录与部分原文件相同，同名结果可能覆盖原文件。是否继续？') != QMessageBox.Yes:
                    return
            from string_replacer import batch_replace
            def work(progress):
                details = []
                result, errors = batch_replace(files, rules, output, progress, details.append)
                return {'结果': result, '错误': errors, '明细': details}
            self.batch_context = (list(rules), output)
            self.tasks.submit(work, self.batch_done)
        except ValueError as exc:
            self.error(exc)

    def extract_project(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择招标文件 / 公告', '', 'Word (*.doc *.docx)')
        if not path:
            return
        def work(progress):
            from tender_info_extractor import extract_project_info_details
            if Path(path).suffix.lower() == '.doc':
                from legacy_office import LegacyOfficeSession, temporary_docx_source
                with LegacyOfficeSession() as session:
                    with temporary_docx_source(path, session) as readable:
                        return extract_project_info_details(readable)
            return extract_project_info_details(path)
        def done(details):
            self.project_details = details
            self.files.remove([i for i, row in enumerate(self.files.rows) if os.path.normcase(row[1]) == os.path.normcase(os.path.abspath(path))])
            dialog = ProjectReview(self, details)
            if dialog.exec() == QDialog.Accepted:
                self.rules.replace(dialog.rules())
                self.status.setText('已回填审核结果；空白字段跳过，请核对后开始替换。')
        self.tasks.submit(work, done)

    def add_scan_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, '选择 Word 来源', '', 'Word (*.doc *.docx)')
        existing = {row[0] for row in self.scan_files.rows}
        for path in paths:
            if path not in existing:
                self.scan_files.append([path])
                existing.add(path)

    def remove_scan_files(self):
        self.scan_files.remove(self.selected_rows(self.scan_file_view))
        self.scans.load([], [])
        self.stamps = {}

    def scan(self):
        paths = [row[0] for row in self.scan_files.rows]
        if not paths:
            return self.error('请先添加 Word 来源文件。')
        from word_scan import scan_word_content
        symbols = self.symbols.text().strip()
        if not symbols:
            return self.error('请至少填写一个条款符号。')
        def done(result):
            tables, sections, self.stamps = result
            self.scan_symbols = symbols
            self.proxy.select_values()
            self.scans.load(tables[0], sections[0])
            self.search.clear()
            self.status.setText(f'扫描完成：{len(tables[0])} 张表格、{len(sections[0])} 个条款章节。勾选后可导出。')
            if tables[2] or sections[2]:
                self.log.appendPlainText(tables[2] or sections[2])
        self.tasks.submit(lambda p: scan_word_content(paths, symbols, p), done)

    def preview_scan(self, current, previous=None):
        if not current.isValid():
            return
        kind, item = self.scans.items[self.proxy.mapToSource(current).row()]
        if kind == 'table':
            from visual_fill import source_grid
            configure_sheet(self.scan_preview, SheetModel(source_grid(item.table)))
            self.preview_text.setPlainText('\n'.join(c.text for c in item.table.cells))
            self.preview_stack.setCurrentIndex(0)
        else:
            self.preview_text.setPlainText('\n\n'.join(c.text for c in item.clauses))
            self.preview_stack.setCurrentIndex(1)

    def check_visible(self, checked):
        for row in range(self.proxy.rowCount()):
            index = self.proxy.mapToSource(self.proxy.index(row, 0))
            self.scans.setData(index, Qt.Checked if checked else Qt.Unchecked, Qt.CheckStateRole)

    def export_scan(self):
        selected = [self.scans.items[row] for row in sorted(self.scans.checked)]
        if not selected:
            return self.error('请勾选至少一项扫描结果。')
        output = self.output.text().strip()
        if not output or not Path(output).is_dir():
            return self.error('请先选择已存在的输出目录。')
        from word_scan import export_scanned_content
        tables = [item for kind, item in selected if kind == 'table']
        sections = [item for kind, item in selected if kind == 'symbol']
        stamps, symbols, keep = dict(self.stamps), self.scan_symbols, self.keep_symbols.isChecked()
        self.tasks.submit(lambda p: export_scanned_content(tables, sections, stamps, output, symbols, keep, p), self.result)

    def open_editor(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择 Excel 模板', '', 'Excel (*.xlsx *.xlsm)')
        if not path:
            return
        from workbook_session import WorkbookSession
        # Use checked tables when present; otherwise offer every scanned table as a source.
        checked = [self.scans.items[row] for row in sorted(self.scans.checked) if self.scans.items[row][0] == 'table']
        items = [item for _, item in checked] if checked else [item for kind, item in self.scans.items if kind == 'table']
        stamps = dict(self.stamps)
        def work(progress):
            from word_scan import validate_scan_files
            if items:
                validate_scan_files(list(dict.fromkeys(x.file_path for x in items)), stamps)
            return WorkbookSession.open(path)
        self.tasks.submit(work, lambda session: VisualEditor(self, session, items, stamps).exec())

    def refresh_library(self):
        if self.tasks.busy:
            return
        from template_library import list_categories, list_templates
        # Clear immediately, including when a refresh fails, so stale checks
        # cannot later target files that changed on disk.
        self.templates.replace(self.templates.rows)
        root = self.library_root
        def done(result):
            categories, items = result
            current = self.category.currentText()
            self.category.blockSignals(True)
            self.category.clear()
            self.category.addItems(categories)
            if current in categories:
                self.category.setCurrentText(current)
            self.category.blockSignals(False)
            self.library_items = items
            self.show_templates()
        self.tasks.submit(lambda p: (list_categories(root), list_templates(root)), done)

    def refresh_templates(self, *_):
        self.show_templates()

    def show_templates(self):
        category = self.category.currentText()
        self.visible_templates = [item for item in self.library_items if item.category == category]
        self.templates.replace([['', x.name, x.category, x.path] for x in self.visible_templates])

    def update_template_selection(self, *_):
        count, total = len(self.templates.checked), len(self.templates.rows)
        self.template_selection.setText(
            f'已勾选 {count} / {total} 个模板 · 切换分类或刷新后清空勾选。' if total else
            '当前分类暂无模板，请添加模板。' if self.category.currentText() else
            '暂无分类，请新建分类或更改模板库目录。')
        self.templates.headers[0] = '选择（清空）' if total and count == total else '选择（全选）'
        self.templates.headerDataChanged.emit(Qt.Horizontal, 0, 0)
        self.delete_templates_button.setEnabled(bool(count))
        self.use_templates_button.setEnabled(bool(count))

    def checked_templates(self):
        return [item for row, item in enumerate(self.visible_templates)
                if row in self.templates.checked and item.category == self.category.currentText()]

    def change_library(self):
        root = QFileDialog.getExistingDirectory(self, '选择模板库目录', self.library_root)
        if root:
            self.library_root = root
            self.library_path.setText(root)
            self.library_items = []
            self.show_templates()
            self.refresh_library()

    def library_action(self, work):
        self.tasks.submit(lambda p: work(), lambda result: self.refresh_library())

    def new_category(self):
        name, ok = QInputDialog.getText(self, '新建分类', '采购方式或采购人名称')
        if ok:
            from template_library import create_category
            self.library_action(lambda: create_category(self.library_root, name))

    def rename_category(self):
        old = self.category.currentText()
        name, ok = QInputDialog.getText(self, '重命名分类', '新分类名称', text=old)
        if ok and old:
            from template_library import rename_category
            self.library_action(lambda: rename_category(self.library_root, old, name))

    def delete_category(self):
        category = self.category.currentText()
        if category and QMessageBox.question(self, '删除空分类', f'删除分类“{category}”？仅空分类可删除。') == QMessageBox.Yes:
            from template_library import delete_category
            self.library_action(lambda: delete_category(self.library_root, category))

    def add_templates(self):
        category = self.category.currentText()
        if not category:
            return self.error('请先新建分类。')
        paths, _ = QFileDialog.getOpenFileNames(self, '复制模板入库', '', 'Office (*.doc *.docx *.xls *.xlsx *.xlsm *.ppt *.pptx)')
        if paths:
            from template_library import add_templates
            self.library_action(lambda: add_templates(self.library_root, category, paths))

    def delete_templates(self):
        selected = self.checked_templates()
        if selected and QMessageBox.question(self, '删除模板', f'删除当前分类勾选的 {len(selected)} 个模板？\n\n' + '\n'.join(x.name for x in selected)) == QMessageBox.Yes:
            from template_library import delete_template
            root = self.library_root
            self.library_action(lambda: [delete_template(root, x) for x in selected])

    def use_templates(self):
        selected = [item.path for item in self.checked_templates()]
        if not selected:
            return self.error('请先勾选当前分类中的模板。')
        self.append_files(selected)
        self.tabs.setCurrentIndex(0)
        self.status.setText(f'已加入 {len(selected)} 个模板，请填写规则并选择库外输出目录。')

    def closeEvent(self, event):
        if self.tasks.busy:
            event.ignore()
            self.status.setText('后台任务进行中，完成后可关闭。')
            return
        if not self.persist:
            event.accept()
            return
        save_settings({'qt_rules': self.rules.rows, 'qt_output_dir': self.output.text(),
                       'template_library_root': self.library_root, 'symbol_chars': self.symbols.text(),
                       'keep_clause_symbols': self.keep_symbols.isChecked(), 'presets': self.presets})
        event.accept()


def deepcopy_rows(rows):
    return [list(row) for row in rows]


class ProjectReview(QDialog):
    def __init__(self, parent, details):
        super().__init__(parent)
        self.details = details
        self.setWindowTitle('审核项目信息 · 候选值与原文来源')
        self.resize(1000, 680)
        layout = QVBoxLayout(self)
        label = QLabel('核对冲突、推断和缺失字段后回填规则。双击左表修改值，右侧显示来源。')
        label.setWordWrap(True)
        layout.addWidget(label)
        split = QSplitter(Qt.Horizontal)
        self.model = RowsModel(['字段', '替换值', '识别状态'], [[d['field'], d['value'], d['status']] for d in details], editable=True)
        self.model.editable_columns = {1}
        self.view = table(self.model)
        self.view.setColumnWidth(0, 160)
        self.view.setColumnWidth(1, 280)
        split.addWidget(self.view)
        self.sources = QPlainTextEdit()
        self.sources.setReadOnly(True)
        split.addWidget(self.sources)
        configure_splitter(split)
        layout.addWidget(split, 1)
        self.view.selectionModel().currentChanged.connect(self.show_sources)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText('确认并回填规则')
        buttons.button(QDialogButtonBox.Cancel).setText('取消')
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        if details:
            self.view.setCurrentIndex(self.model.index(0, 1))

    def show_sources(self, current, previous=None):
        if current.isValid():
            detail = self.details[current.row()]
            self.sources.setPlainText('\n\n'.join(f'{c["kind"]}：{c["value"]}\n{c["source"]}' for c in detail['candidates']) or '未识别到候选，请手动核对原文。')

    def rules(self):
        return [[d['placeholder'], row[1]] for d, row in zip(self.details, self.model.rows)]


def main(argv=None):
    app = QApplication.instance() or QApplication(sys.argv if argv is None else argv)
    def report_exception(kind, value, traceback):
        if issubclass(kind, KeyboardInterrupt):
            return sys.__excepthook__(kind, value, traceback)
        try:
            from ui_common import write_error_log
            log = write_error_log(kind, value, traceback, context='Qt 主线程')
            message = f'{value}\n错误日志：{log}'
        except Exception:
            message = str(value)
        QMessageBox.critical(None, '程序出错', message)
    sys.excepthook = report_exception
    apply_theme(app)
    smoke = '--smoke-test' in (sys.argv if argv is None else argv)
    window = MainWindow(restore=not smoke)
    window.show()
    if smoke:
        QTimer.singleShot(350, app.quit)
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
