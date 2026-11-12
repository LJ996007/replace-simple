"""Portable mixed-operation recipes; replay only through validated workbook services."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path

from openpyxl.utils.cell import range_boundaries, get_column_letter
from review_mapping import layout_signature
from visual_fill import source_grid, drop_values, plan_visual_fill, _value
from worksheet_layout import RowMap, preview_layout, auto_heights, empty_rows
from workbook_session import WorkbookSession


def source_signature(info):
    # Byte-for-byte compatible with VisualFillDialog's original signature.
    return hashlib.sha256(json.dumps([info['max_row'], info['max_col'], info['merges'],
        [(k, v) for k, v in sorted(info['cells'].items()) if range_boundaries(k)[1] == 1]],
        ensure_ascii=False).encode()).hexdigest()


def export_recipe(session, items):
    return {'version': 2, 'operations': deepcopy(session.operations),
            'sources': [source_signature(source_grid(item.table)) for item in items],
            'targets': {n: layout_signature(v, v['guess']) for n, v in session.original_sheets.items()}}


def save_recipe_file(path, recipe):
    from file_io import atomic_output_path
    with atomic_output_path(path) as temporary:
        Path(temporary).write_text(json.dumps(recipe, ensure_ascii=False, indent=2), encoding='utf-8')


def read_recipe_file(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def _bounds(value, max_col=16384, max_row=1048576):
    if (not isinstance(value, (list, tuple)) or len(value) != 4 or
            any(type(x) is not int or x < 1 for x in value) or
            not value[0] <= value[2] <= max_col or not value[1] <= value[3] <= max_row):
        raise ValueError('配方包含无效选区。')
    return value


def mapped_drop(session, source, bounds, name, row, column):
    _bounds(bounds, source['max_col'], source['max_row'])
    _bounds([column, row, column, row])
    mapping = RowMap(session.layouts.get(name, {}).get('delete_rows', []))
    display_row = mapping.display(row)
    if display_row is None:
        raise ValueError('保存的拖入行已被删除。')
    values = drop_values(source, bounds, session.view_info(name), display_row, column)
    if not values:
        raise ValueError('配方来源选区没有非空内容。')
    return {f'{get_column_letter(range_boundaries(ref)[0])}{mapping.original(range_boundaries(ref)[1])}': v
            for ref, v in values.items()}


def _check_plan(session):
    session.verify()
    plan = plan_visual_fill(session.path, session.pending, session.layouts)
    if plan['digest'] != session.digest:
        raise ValueError('模板文件已变化，请重新打开。')
    if plan['errors']:
        raise ValueError('；'.join(plan['errors']))


def replay_recipe(session, items, recipe, temp_path, verify_sources):
    """Prepare in a separate session. Caller commits only a fully verified result."""
    session.verify()
    verify_sources()
    if not isinstance(recipe, dict):
        raise ValueError('不支持的配方格式。')
    grids = [source_grid(item.table) for item in items]
    recipe = deepcopy(recipe)
    single_drop = recipe.get('version') == 1 and 'operations' not in recipe
    if single_drop:
        name = recipe.get('sheet')
        matches = [i for i, grid in enumerate(grids)
                   if layout_signature(grid, {'header': 1}) == recipe.get('source_signature')]
        if len(matches) > 1:
            matches = [i for i in matches if items[i].filename == recipe.get('source_filename')
                       and items[i].table_index == recipe.get('table_index')]
        if len(matches) != 1:
            raise ValueError('配方来源表格缺失、结构已变化或不唯一。')
        recipe = {'operations': [{'kind': 'drag', 'source': matches[0], 'sheet': name,
                   'bounds': recipe.get('bounds'), 'row': recipe.get('row'), 'column': recipe.get('column')}],
                  'sources': [source_signature(g) for g in grids],
                  'targets': {name: recipe.get('target_signature')}}
    elif recipe.get('version') not in (None, 2):
        raise ValueError('不支持的配方版本。')
    operations = recipe.get('operations')
    if not isinstance(operations, list) or not operations:
        raise ValueError('配方操作历史为空或无效。')
    work = WorkbookSession.open(session.original)
    if work.digest != session.original_digest:
        raise ValueError('原模板文件已变化。')
    if single_drop:
        work.restore(session.snapshot())
    target_sheets = session.sheets if single_drop else session.original_sheets
    original_operations = deepcopy(session.operations) if single_drop else []
    try:
        for saved in operations:
            op = deepcopy(saved)
            kind = op['kind']
            name = op['configs'][1]['sheet'] if kind == 'auto' else op['sheet']
            if (name not in target_sheets or recipe['targets'][name] !=
                    layout_signature(target_sheets[name], target_sheets[name]['guess'])):
                raise ValueError('目标工作表的表头或合并结构已变化。')
            if (name not in (work.sheets if single_drop else work.original_sheets) or recipe['targets'][name] !=
                    layout_signature((work.sheets if single_drop else work.original_sheets)[name],
                                     (work.sheets if single_drop else work.original_sheets)[name]['guess'])):
                raise ValueError('目标工作表的表头或合并结构已变化。')
            if 'source' in op:
                index = op['source']
                if type(index) is not int or index < 0:
                    raise ValueError('无效的来源索引。')
                signature = recipe['sources'][index]
                matches = [i for i, grid in enumerate(grids) if source_signature(grid) == signature]
                if single_drop:
                    matches = [index]  # v1 already matched signature plus filename/table index above.
                if len(matches) != 1:
                    raise ValueError('来源表结构已变化或存在多个同结构表。')
                op['source'] = matches[0]
            if kind == 'structure':
                _bounds(op['bounds'])
                work.structure(temp_path(), name, op['action'], op['bounds'], op.get('count', 1), temp_path())
            elif kind == 'layout':
                rows = op['rows']
                action = op['action']
                maximum = 16384 if action == 'width' else 1048576
                if not isinstance(rows, list) or not rows or any(type(r) is not int or not 1 <= r <= maximum for r in rows):
                    raise ValueError('无效的排版行列。')
                layout = work.layouts.setdefault(name, {'heights': {}, 'widths': {}, 'wrap_cells': [], 'delete_rows': []})
                if action == 'fit':
                    info = deepcopy(work.sheets[name])
                    for col, width in layout.get('widths', {}).items():
                        info.setdefault('column_pixels', {})[col] = width * 7 + 5
                    heights, wraps = auto_heights(info, work.pending.get(name, {}), rows, layout['delete_rows'])
                    layout['heights'].update(heights)
                    layout['wrap_cells'] = sorted(set(layout['wrap_cells']) | set(wraps))
                elif action in ('height', 'width'):
                    layout.setdefault('widths' if action == 'width' else 'heights', {}).update(
                        {r: op['height'] for r in rows if action == 'width' or r not in layout['delete_rows']})
                elif action == 'delete':
                    deleted = empty_rows(work.sheets[name], work.pending.get(name, {}), rows)
                    layout['delete_rows'] = sorted(set(layout['delete_rows']) | set(deleted))
                    for row in deleted:
                        layout['heights'].pop(row, None)
                else:
                    raise ValueError('未知的排版操作。')
            elif kind == 'drag':
                work.edit(name, mapped_drop(work, grids[op['source']], op['bounds'], name, op['row'], op['column']))
            elif kind == 'edit':
                work.edit(name, {op['cell']: _value(op['value'])})
            elif kind == 'clear':
                work.edit(name, {ref: '' for ref in op['cells']})
            elif kind == 'cells':
                work.edit(name, {ref: _value(value) for ref, value in op['values'].items()})
            elif kind == 'format':
                work.layout(name, {int(k): v for k, v in op['heights'].items()},
                            {int(k): v for k, v in op['widths'].items()}, op['wrap_cells'])
            elif kind == 'auto':
                from review_mapping import plan_fill
                from review_table_filler import extract_review_rules
                source, target = op['configs']
                rules = extract_review_rules(items[op['source']].table, source, target.get('include_price', False))
                overrides = {int(k): v for k, v in op.get('overrides', {}).items()}
                plan = plan_fill(work.path, rules, target, overrides)
                if plan['errors']:
                    raise ValueError('自动匹配设置需要重新核对：' + '；'.join(plan['errors']))
                work.edit(name, {c['cell']: c['new'] for c in plan['changes']})
            else:
                raise ValueError('配方包含未知操作：' + str(kind))
            _check_plan(work)
            # Services also log operations; replace that log with the exact imported history.
        if any(v.get('delete_rows') for v in work.layouts.values()):
            output = temp_path()
            work.save(output)
            version = WorkbookSession.open(output)
            work.path, work.digest, work.sheets = version.path, version.digest, version.sheets
            work.pending, work.layouts = {}, {}
            work.structure_count += 1
        work.operations = original_operations
        for saved in operations:
            op = deepcopy(saved)
            if 'source' in op:
                sig = recipe['sources'][op['source']]
                if not single_drop:
                    op['source'] = next(i for i, g in enumerate(grids) if source_signature(g) == sig)
            work.operations.append(op)
        session.verify()
        verify_sources()
        work.verify()
        return work
    except (KeyError, TypeError, IndexError, AttributeError, OverflowError) as exc:
        raise ValueError('配方字段缺失或格式无效：' + str(exc)) from exc
