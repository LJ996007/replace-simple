"""UI-independent workbook edit state. Versions share immutable worksheet snapshots."""
from copy import deepcopy
from pathlib import Path
import hashlib
import re

from openpyxl.utils.cell import range_boundaries
from review_mapping import _anchor
from visual_fill import inspect_visual_workbook, plan_visual_fill, apply_visual_plan, _value


class WorkbookSession:
    def __init__(self, path, sheets, digest):
        self.original = str(path)
        self.original_digest = digest
        self.path = str(path)
        self.digest = digest
        self.sheets = sheets
        self.original_sheets = sheets
        self.operations = []
        self.pending = {}
        self.layouts = {}
        self.history = []
        self.future = []
        self.structure_count = 0

    @classmethod
    def open(cls, path):
        before = hashlib.sha256(Path(path).read_bytes()).hexdigest()
        sheets = inspect_visual_workbook(path)
        if hashlib.sha256(Path(path).read_bytes()).hexdigest() != before:
            raise ValueError('读取期间模板已变化，请重新打开。')
        return cls(path, sheets, before)

    def snapshot(self):
        return self.path, self.digest, self.sheets, deepcopy(self.pending), self.structure_count, deepcopy(self.layouts), deepcopy(self.operations)

    def restore(self, state):
        self.path, self.digest, self.sheets, pending, self.structure_count, layouts, operations = state
        self.pending = deepcopy(pending)
        self.layouts = deepcopy(layouts)
        self.operations = deepcopy(operations)

    def checkpoint(self):
        self.history.append(self.snapshot())
        del self.history[:-100]
        self.future.clear()

    def undo(self):
        if self.history:
            self.future.append(self.snapshot())
            self.restore(self.history.pop())

    def redo(self):
        if self.future:
            self.history.append(self.snapshot())
            self.restore(self.future.pop())

    def edit(self, sheet, values, operation=None):
        info = self.sheets[sheet]
        if info.get('protected'):
            raise ValueError('工作表受保护，请先解除保护。')
        normalized = {}
        for ref, value in values.items():
            if not re.fullmatch(r'[A-Z]{1,3}[1-9]\d*', ref):
                raise ValueError('无效单元格地址。')
            col, row, _, _ = range_boundaries(ref)
            if col > 16384 or row > 1048576:
                raise ValueError('单元格超出 Excel 范围。')
            ref = _anchor(ref, info['merges'])
            col, row, _, _ = range_boundaries(ref)
            if ref in info.get('formulas', {}) or any(
                l <= col <= r and t <= row <= b
                for l, t, r, b in map(range_boundaries, info.get('formula_ranges', []))
            ):
                raise ValueError(f'{sheet}!{ref} 包含公式，不能覆盖。')
            if len(str(value)) > 32767:
                raise ValueError('单元格内容超过 Excel 长度上限。')
            if ref in normalized and normalized[ref] != value:
                raise ValueError(f'{ref} 合并区域对应多个不同值。')
            normalized[ref] = value
        if normalized:
            self.checkpoint()
            self.pending.setdefault(sheet, {}).update(normalized)
            self.operations.append(deepcopy(operation) if operation is not None else
                                   {'kind': 'cells', 'sheet': sheet, 'values': {r: str(v) for r, v in normalized.items()}})

    def layout(self, sheet, heights=None, widths=None, wrap_cells=()):
        import math
        if self.sheets[sheet].get('protected'):
            raise ValueError('工作表受保护，请先解除保护。')
        heights, widths = dict(heights or {}), dict(widths or {})
        for values, limit, maximum in ((heights, 409, 1048576), (widths, 255, 16384)):
            if any(not isinstance(k, int) or not 1 <= k <= maximum or
                   not math.isfinite(float(v)) or not 1 <= float(v) <= limit for k, v in values.items()):
                raise ValueError('行高须为 1–409 磅，列宽须为 1–255。')
        if not heights and not widths and not wrap_cells:
            return
        self.checkpoint()
        layout = self.layouts.setdefault(sheet, {'heights': {}, 'widths': {}, 'wrap_cells': [], 'delete_rows': []})
        layout['heights'].update(heights)
        layout['widths'].update(widths)
        layout['wrap_cells'] = sorted(set(layout['wrap_cells']) | set(wrap_cells))
        self.operations.append({'kind': 'format', 'sheet': sheet, 'heights': {str(k): v for k, v in heights.items()},
                                'widths': {str(k): v for k, v in widths.items()}, 'wrap_cells': list(wrap_cells)})

    def view_info(self, sheet):
        if sheet not in self.layouts:
            return self.sheets[sheet]
        from worksheet_layout import preview_layout
        return preview_layout(self.sheets[sheet], self.pending.get(sheet, {}), self.layouts[sheet])[0]

    def verify(self):
        for path, digest in ((self.original, self.original_digest), (self.path, self.digest)):
            if hashlib.sha256(Path(path).read_bytes()).hexdigest() != digest:
                raise ValueError('模板文件已变化，请重新打开；当前操作未保存。')

    def save(self, output, progress=None):
        self.verify()
        plan = plan_visual_fill(self.path, self.pending, self.layouts)
        if plan['digest'] != self.digest:
            raise ValueError('模板文件已变化，请重新打开。')
        if plan['errors']:
            raise ValueError('；'.join(plan['errors']))
        if not (plan['count'] or plan['layout_count']) and self.structure_count:
            from file_io import atomic_output_path
            if Path(output).exists() or Path(output).resolve() == Path(self.original).resolve():
                raise ValueError('请另存为不存在的新文件。')
            with atomic_output_path(output) as temporary:
                Path(temporary).write_bytes(Path(self.path).read_bytes())
            return {'output_path': str(output), 'count': 0}
        return apply_visual_plan(self.path, output, plan, progress)

    def structure(self, output, sheet, action, bounds, count=1, intermediate=None):
        from workbook_structure import edit_structure
        self.verify()
        current = self.path
        if any(self.pending.values()) or self.layouts:
            plan = plan_visual_fill(self.path, self.pending, self.layouts)
            if plan['digest'] != self.digest:
                raise ValueError('模板文件已变化，请重新打开。')
            if plan['errors']:
                raise ValueError('；'.join(plan['errors']))
            if plan['count'] or plan['layout_count']:
                if intermediate is None:
                    raise ValueError('结构操作需要临时文件路径。')
                self.save(intermediate)
                current = intermediate
        expected = self.digest if current == self.path else hashlib.sha256(Path(current).read_bytes()).hexdigest()
        edit_structure(current, output, sheet, action, bounds, count, expected_digest=expected)
        version = self.open(output)
        self.checkpoint()
        self.path, self.digest, self.sheets = version.path, version.digest, version.sheets
        self.pending = {}
        self.layouts = {}
        self.structure_count += 1
        self.operations.append({'kind': 'structure', 'sheet': sheet, 'action': action,
                                'bounds': list(bounds), 'count': count})
