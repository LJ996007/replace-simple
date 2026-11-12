"""Secondary workbench actions, backed by the same replacement and scan engines."""
from pathlib import Path
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QDialogButtonBox, QListWidget,
                              QListWidgetItem, QInputDialog, QFileDialog, QTableView)
from qt_models import RowsModel
from replacement_rules import normalize_rule_rows, validate_rules


class WorkbenchActions:
    def batch_done(self, value):
        self.last_batch = value
        self.result(value)

    def retry_failed(self):
        if self.tasks.busy:
            return
        details = (self.last_batch or {}).get('明细', [])
        failed = [x['source_path'] for x in details if x['status'] == '失败']
        if not failed or not self.batch_context:
            return self.error('没有可重试的失败文件。')
        rules, output = self.batch_context
        from string_replacer import batch_replace
        def work(progress):
            retried = []
            result, errors = batch_replace(failed, rules, output, progress, retried.append)
            combined = [x for x in details if x['status'] != '失败'] + retried
            return {'结果': result, '错误': errors, '明细': combined}
        self.tasks.submit(work, self.batch_done)

    def export_report(self):
        if self.tasks.busy:
            return
        if not self.last_batch:
            return self.error('请先完成一次批量替换。')
        path, _ = QFileDialog.getSaveFileName(self, '导出处理报告', '处理报告.csv', 'CSV (*.csv)')
        if not path:
            return
        details = [dict(x) for x in self.last_batch.get('明细', [])]
        def work(progress):
            import csv
            from file_io import atomic_output_path
            with atomic_output_path(path) as temporary:
                with open(temporary, 'w', encoding='utf-8-sig', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerow(['来源', '输出', '状态', '命中数', '错误'])
                    for x in details:
                        row = [x['source_path'], x['output_path'], x['status'], x['count'], x['error']]
                        writer.writerow(["'"+v if isinstance(v, str) and v.startswith(('=', '+', '-', '@')) else v for v in row])
            return path
        self.tasks.submit(work, lambda p: self.status.setText(f'处理报告已导出：{p}'))

    def preview_outputs(self):
        if self.tasks.busy:
            return
        try:
            validate_rules(self.rules.rows)
            rules = [(a, b) for a, b in normalize_rule_rows(self.rules.rows) if b]
            paths, output = [row[1] for row in self.files.rows], self.output.text().strip()
            if not paths or not output:
                raise ValueError('请先添加文件并选择输出目录。')
            def work(progress):
                from string_replacer import get_output_path
                reserved = set()
                return [[p, get_output_path(p, rules, output, reserved)] for p in paths]
            self.tasks.submit(work, self.show_output_rows)
        except ValueError as exc:
            self.error(exc)

    def show_output_rows(self, rows):
        dialog = QDialog(self)
        dialog.setWindowTitle('预览输出路径 · 执行时按磁盘占用重新核对')
        dialog.resize(1000, 560)
        layout = QVBoxLayout(dialog)
        view = QTableView()
        view.setModel(RowsModel(['来源', '预计输出'], rows, parent=dialog))
        view.setColumnWidth(0, 430)
        view.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(view)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec()

    def manage_presets(self):
        if self.tasks.busy:
            return
        from ui_common import normalize_presets, append_presets_to_rule_rows
        from qt_main import toolbar
        dialog = QDialog(self)
        dialog.setWindowTitle('常用占位符 · 勾选添加到规则，双击编辑内容')
        dialog.resize(600, 550)
        layout = QVBoxLayout(dialog)
        listing = QListWidget()
        def add(value):
            item = QListWidgetItem(value, listing)
            item.setFlags(item.flags() | Qt.ItemIsEditable | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
        for value in self.presets:
            add(value)
        layout.addWidget(listing)
        def new():
            value, ok = QInputDialog.getText(dialog, '新增预设', '原文本 / 占位符')
            if ok and value.strip():
                add(value.strip())
        def remove():
            for item in listing.selectedItems():
                listing.takeItem(listing.row(item))
        def apply():
            self.presets = normalize_presets([listing.item(i).text() for i in range(listing.count())])
            selected = [listing.item(i).text() for i in range(listing.count()) if listing.item(i).checkState() == Qt.Checked]
            rows, count = append_presets_to_rule_rows(self.rules.rows, normalize_presets(selected))
            self.rules.replace(rows)
            self.status.setText(f'已保存预设并添加 {count} 条规则。')
            dialog.accept()
        toolbar(layout, [('新增', new), ('删除选中', remove), ('保存并添加勾选项', apply)])
        dialog.exec()

    def filter_scans(self):
        dialog = QDialog(self)
        dialog.setWindowTitle('类型与章节筛选 · 留空表示全部')
        dialog.resize(650, 500)
        layout = QVBoxLayout(dialog)
        kinds, sections = QListWidget(), QListWidget()
        def add(view, label, value, checked):
            item = QListWidgetItem(label, view)
            item.setData(Qt.UserRole, value)
            item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        for kind, label in [('table', '表格'), ('symbol', '条款')]:
            add(kinds, label, kind, kind in self.proxy.kinds)
        kinds.setMaximumHeight(90)
        layout.addWidget(kinds)
        for section in sorted({x.section for _, x in self.scans.items}):
            add(sections, section or '（未分章节）', section, section in self.proxy.sections)
        layout.addWidget(sections)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() == QDialog.Accepted:
            selected = lambda view: [view.item(i).data(Qt.UserRole) for i in range(view.count()) if view.item(i).checkState() == Qt.Checked]
            self.proxy.select_values(selected(kinds), selected(sections))

    def popup_scan(self):
        current = self.scan_view.currentIndex()
        if not current.isValid():
            return self.error('请先选中一张扫描表格。')
        kind, item = self.scans.items[self.proxy.mapToSource(current).row()]
        if kind != 'table':
            return self.error('请选择表格；条款全文位于右侧。')
        from qt_editor import SheetPreview
        from visual_fill import source_grid
        SheetPreview(self, source_grid(item.table), item.filename).exec()
