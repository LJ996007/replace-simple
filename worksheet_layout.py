"""可预览的行排版：坐标始终存为模板坐标，删行后才转换显示/输出坐标。"""
import bisect
import math
import posixpath
import re
import unicodedata
from copy import deepcopy

from lxml import etree as ET
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter, range_boundaries, column_index_from_string
from openpyxl.formula.translate import Translator

from review_table_filler import _tag, _xml, _serialize


class RowMap:
    def __init__(self, deleted=()):
        self.deleted = sorted(set(deleted))

    def display(self, row):
        if row in self.deleted:
            return None
        return row - bisect.bisect_left(self.deleted, row)

    def original(self, row):
        for removed in self.deleted:
            if removed <= row:
                row += 1
            else:
                break
        return row

    def boundary(self, row):
        return row - bisect.bisect_right(self.deleted, row)

    def span(self, top, bottom):
        while top <= bottom and top in self.deleted:
            top += 1
        while bottom >= top and bottom in self.deleted:
            bottom -= 1
        return (self.display(top), self.display(bottom)) if top <= bottom else None

    def ref(self, ref):
        """A1 区域/整行/整列，保留绝对引用；全部删除返回 None。"""
        match = re.fullmatch(r'(\$?[A-Z]{1,3})?(\$?)([1-9]\d*)(?::(\$?[A-Z]{1,3})?(\$?)([1-9]\d*))?', ref, re.I)
        if not match:
            if re.fullmatch(r'\$?[A-Z]{1,3}:\$?[A-Z]{1,3}', ref, re.I):
                return ref
            raise ValueError(f'暂不支持调整引用 {ref}，请保留这些行。')
        c, absolute, first, c2, absolute2, last = match.groups()
        span = self.span(int(first), int(last or first))
        if not span:
            return None
        result = f'{c or ""}{absolute}{span[0]}'
        return result + (f':{c2 or ""}{absolute2}{span[1]}' if last else '')


class InsertRowMap(RowMap):
    """在 before 前插入 count 行；区域跨越插入点时随之扩展。"""
    def __init__(self, before, count):
        super().__init__()
        self.before, self.count = before, count

    def display(self, row):
        return row + (self.count if row >= self.before else 0)

    def original(self, row):
        if self.before <= row < self.before + self.count:
            return None
        return row - (self.count if row >= self.before + self.count else 0)

    def boundary(self, row):
        return self.display(row)


class InsertColumnMap(InsertRowMap):
    axis = 'column'

    def ref(self, ref):
        if re.fullmatch(r'\$?\d+:\$?\d+',ref):return ref
        def move(match):
            number=self.display(column_index_from_string(match[2]))
            if number>16384:raise ValueError('插入后超出 Excel 最大列数。')
            return match[1]+get_column_letter(number)
        return re.sub(r'(\$?)([A-Za-z]{1,3})(?=\$?\d|:|$)',move,ref)


def preview_layout(info, values, layout, sheet=None, layouts=None):
    """返回删行/行高后的画布数据和待填内容，不改调用方快照。"""
    result = deepcopy(info)
    mapping = RowMap(layout.get('delete_rows', []))
    for key in ('cells', 'styles'):
        result[key] = {moved: value for ref, value in result.get(key, {}).items()
                       if (moved := mapping.ref(ref))}
    if sheet and layouts:
        mappings = {name: RowMap(value.get('delete_rows', [])) for name, value in layouts.items() if value.get('delete_rows')}
        if mappings:
            for ref, formula in info.get('formulas', {}).items():
                if moved := mapping.ref(ref):
                    result['cells'][moved] = shift_formula('=' + formula, sheet, mappings)
    result['heights'] = {mapping.display(r): h for r, h in info.get('heights', {}).items()
                         if mapping.display(r) is not None}
    for row, height in layout.get('heights', {}).items():
        if mapping.display(row) is not None:
            result['heights'][mapping.display(row)] = height * 4 / 3
    for col,width in layout.get('widths',{}).items():
        result.setdefault('widths',{})[col]=width*7+5
        result.setdefault('column_pixels',{})[col]=width*7+5
    for ref in layout.get('wrap_cells',[]):
        if moved:=mapping.ref(ref):
            # 同一模板样式可能被多个格共用，只替换当前格的预览样式。
            styles=result.setdefault('styles',{})
            styles[moved]={**styles.get(moved,{}),'wrap':True}
    result['merges'] = [ref for old in info['merges'] if (ref := mapping.ref(old))]
    result['max_row'] = max(1, info['max_row'] - bisect.bisect_right(mapping.deleted, info['max_row']))
    updates = {moved: value for ref, value in values.items() if (moved := mapping.ref(ref))}
    return result, updates


def empty_rows(info, values, rows):
    """整行有值/公式，或属于纵向合并/数组公式的行均不能删除。"""
    occupied = {range_boundaries(ref)[1] for ref, value in (info['cells'] | values).items()
                if value is not None and str(value).strip()}
    protected = [range_boundaries(ref) for ref in info['merges'] if range_boundaries(ref)[1] != range_boundaries(ref)[3]]
    protected.extend(range_boundaries(ref) for ref in info.get('formula_ranges', []))
    return sorted(r for r in set(rows) if 1 <= r <= info['max_row'] and r not in occupied
                  and not any(t <= r <= b for _, t, _, b in protected))


def auto_heights(info, values, rows, deleted=(), measure=None):
    """依据模板字体、实际列宽和合并区域估算高度，单位为 point。

    measure(text, style) 可由界面注入本机字体测量，默认实现供无 GUI 调用。
    """
    rows = set(rows) - set(deleted)
    merges = [range_boundaries(ref) for ref in info['merges']]
    # 同一跨行合并区域一起计算，不把一格文字乘上行数。
    changed = True
    while changed:
        old = set(rows)
        for _, top, _, bottom in merges:
            if rows.intersection(range(top, bottom + 1)):
                rows.update(range(top, bottom + 1))
        changed = old != rows
    heights = {r: 18.0 for r in rows}
    wrapped, constraints = [], []
    for ref, value in (info['cells'] | values).items():
        col, row, _, _ = range_boundaries(ref)
        if row not in rows or value is None or str(value) == '':
            continue
        l, t, r, b = next((m for m in merges if m[0] <= col <= m[2] and m[1] <= row <= m[3]), (col, row, col, row))
        if (col, row) != (l, t):
            continue
        style = info.get('styles', {}).get(ref, {})
        size = style.get('size', 11)
        if style.get('rotation', 0):
            raise ValueError(f'{ref} 使用旋转文字，请改用手动行高。')
        width = sum(info.get('column_pixels', {}).get(c, 89) for c in range(l, r + 1)) - 12 - style.get('indent', 0) * size * 2
        width = max(8, width)
        text = str(value)
        if ref in info.get('formulas', {}) and ref not in values:
            text = info.get('formula_display', {}).get(ref, '') or '0000000000'
        def advance(ch):
            return measure(ch, style) if measure else size * 4 / 3 * (1 if unicodedata.east_asian_width(ch) in 'WF' else .58)
        lines, used = 1, 0
        for ch in text.replace('\r\n', '\n').replace('\r', '\n'):
            amount = advance(ch)
            if ch == '\n':
                lines += 1; used = 0
            else:
                if used and used + amount > width:
                    lines += 1; used = 0
                used += amount
        height = round(max(18, lines * size * 1.4 + 9), 2)
        wrapped.append(ref)
        if t == b:
            heights[t] = max(heights[t], height)
        else:
            constraints.append((t, b, height))
    for top, bottom, required in constraints:
        extra = required - sum(heights[r] for r in range(top, bottom + 1))
        if extra > 0:
            for row in range(top, bottom + 1):
                heights[row] = math.ceil((heights[row] + extra / (bottom - top + 1)) * 100) / 100
    if any(height > 409 for height in heights.values()):
        raise ValueError('内容超过 Excel 单行高度上限，请加宽说明列或将内容分到多行。')
    return heights, wrapped


def shift_formula(text, current_sheet, mappings):
    """按删行语义更新 A1 引用（包括绝对引用），不改公式中的文字常量。"""
    if not text:
        return text
    has_equal = text.startswith('=')
    try:
        tokens = Tokenizer(text if has_equal else '=' + text).items
    except Exception as exc:
        raise ValueError('存在无法解析的公式，暂不能增删行。') from exc
    result = []
    mapping_names = {name.casefold(): name for name in mappings}
    for token in tokens:
        value = token.value
        if token.type == 'FUNC' and token.subtype == 'OPEN' and value.upper().split('.')[-1] in {'INDIRECT(', 'OFFSET('}:
            raise ValueError('工作簿包含 INDIRECT/OFFSET 动态引用，暂不能增删行。')
        if token.type == 'OPERAND' and token.subtype == 'RANGE':
            prefix, sep, ref = value.rpartition('!')
            sheet = prefix.strip("'").replace("''", "'") if sep else current_sheet
            if sep and '[' in sheet:
                result.append(value)  # 外部引用指向其他工作簿的单元格，坐标不随本簿增删行变化
                continue
            if sep and ':' in sheet:
                raise ValueError('工作簿包含三维引用，暂不能增删行。')
            if not sep:
                ref = value
            sheet = mapping_names.get(sheet.casefold(), sheet) if sheet else sheet
            if sheet is None and re.fullmatch(r'\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?', ref, re.I):
                raise ValueError('全局名称含未指定工作表的引用，暂不能增删行。')
            if sheet in mappings:
                if re.fullmatch(r'\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?|\$?[1-9]\d*:\$?[1-9]\d*|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}', ref, re.I):
                    moved = mappings[sheet].ref(ref)
                    if moved is None:
                        raise ValueError(f'公式直接引用待删除区域 {value}，请保留这些行。')
                    value = (prefix + '!' if sep else '') + moved
                elif '[' in ref or '#' in ref:
                    raise ValueError('存在结构化或溢出引用，暂不能增删行。')
        result.append(value)
    return ('=' if has_equal else '') + ''.join(result)


def _shift_qualified(value, mappings):
    """按对应工作表映射改写 'Sheet!A1:B2'（图表、透视表缓存的坐标引用）。

    外部引用、名称引用或该表未被增删行时原样返回；区域被删空则拒绝。
    """
    prefix, sep, ref = value.rpartition('!')
    sheet = prefix.strip("'").replace("''", "'") if sep else ''
    if not sep or '[' in sheet or ':' in sheet:
        return value
    mapping = next((m for name, m in mappings.items() if name.casefold() == sheet.casefold()), None)
    if mapping is None:
        return value
    if not re.fullmatch(r'\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?|\$?[1-9]\d*:\$?[1-9]\d*|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}', ref, re.I):
        raise ValueError(f'{value} 的引用无法随增删行列自动调整，请先在 Excel 中处理该对象。')
    moved = mapping.ref(ref)
    if moved is None:
        raise ValueError(f'{value} 引用的区域已被删除，请保留这些行或先在 Excel 中调整该对象。')
    return f'{prefix}!{moved}'


def _object_intersection(mapping, bounds):
    """操作是否落在对象区域内部：插入点在区域内，或删除的行/列与区域重叠。"""
    left, top, right, bottom = bounds
    if getattr(mapping, 'axis', 'row') == 'column':
        if left < getattr(mapping, 'before', 0) <= right:
            return True
        return any(left <= c <= right for c in getattr(mapping, 'deleted', ()))
    if hasattr(mapping, 'before'):  # 插入行
        return top < mapping.before <= bottom
    return any(top <= r <= bottom for r in mapping.deleted)


def _local(node):
    return ET.QName(node).localname


def _shift_object_parts(package, mappings):
    """图表、透视表、浮动图形内也缓存坐标：能精确改写的改写，穿越区域内部的拒绝。"""
    extra = {}
    for part in package.zip.namelist():
        if not (part.startswith('xl/charts/') and part.endswith('.xml') and '_rels' not in part):
            continue
        root = _xml(package.zip.read(part))
        moved = False
        for node in root.iter():
            if _local(node) == 'f' and node.text and '!' in node.text:
                new = _shift_qualified(node.text.strip(), mappings)
                if new != node.text.strip():
                    node.text = new
                    moved = True
        if moved:
            extra[part] = _serialize(root)
    sheet_of = {part: name for name, part in package.sheets}
    cache_done = set()
    for rels in (p for p in package.zip.namelist()
                 if re.fullmatch(r'xl/worksheets/_rels/[^/]+\.rels', p)):
        sheet = sheet_of.get('xl/worksheets/' + rels.rsplit('/', 1)[-1][:-len('.rels')])
        for relation in _xml(package.zip.read(rels)):
            kind = relation.get('Type', '').rsplit('/', 1)[-1]
            target = relation.get('Target')
            if kind in {'drawing', 'vmlDrawing'}:
                if sheet not in mappings:
                    continue
                _shift_drawing_anchors(package, mappings[sheet], kind, target, sheet, extra)
            elif kind == 'comments':
                if sheet not in mappings:
                    continue
                _shift_comment_refs(package, mappings[sheet], target, sheet, extra)
            elif kind == 'pivotTable':
                _shift_pivot_table(package, mappings, sheet, target, rels, cache_done, extra)
    return extra


def _resolve(base, target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base, target))


def _shift_drawing_anchors(package, mapping, kind, target, sheet, extra):
    """编辑表上浮动图形（按钮、图片、批注等）的锚点随增删行列移动。"""
    part = _resolve('xl/worksheets', target)
    row_mode = getattr(mapping, 'axis', 'row') == 'row'
    root = _xml(package.zip.read(part))
    moved = False
    if kind == 'drawing':  # xdr 锚：from/to 内为 0 起始的行列号
        for anchor in root.iter():
            if _local(anchor) not in {'twoCellAnchor', 'oneCellAnchor'}:
                continue
            for edge in anchor:
                if _local(edge) not in {'from', 'to'}:
                    continue
                for child in edge:
                    axis = _local(child) == 'row'
                    if child.text is None or axis != row_mode or _local(child) not in {'row', 'col'}:
                        continue
                    new = mapping.display(int(child.text) + 1)
                    if new is None:
                        raise ValueError(f'{sheet} 的图形对象位于待删除的行内，请保留这些行。')
                    child.text = str(new - 1)
                    moved = True
    else:  # VML 锚（表单按钮、批注的兼容表示）：逗号分隔的 8 段行列坐标
        for node in root.iter():
            if _local(node) != 'Anchor' or not node.text:
                continue
            fields = node.text.split(',')
            if len(fields) != 8 or not all(f.strip().lstrip('-').isdigit() for f in fields):
                raise ValueError(f'{sheet} 的图形对象锚点无法识别，暂不能增删行列；仍可调整行高。')
            for index in ((2, 6) if row_mode else (0, 4)):
                new = mapping.display(int(fields[index]) + 1)
                if new is None:
                    raise ValueError(f'{sheet} 的图形对象位于待删除的行内，请保留这些行。')
                fields[index] = str(new - 1)
            node.text = ','.join(fields)
            moved = True
    if moved:
        extra[part] = _serialize(root)


def _shift_comment_refs(package, mapping, target, sheet, extra):
    """批注锚定在单元格上：comments 部件里的 ref 随增删行列移动。"""
    part = _resolve('xl/worksheets', target)
    root = _xml(package.zip.read(part))
    moved = False
    for node in root.iter():
        if _local(node) != 'comment' or not node.get('ref'):
            continue
        new = mapping.ref(node.get('ref'))
        if new is None:
            raise ValueError(f'{sheet} 待删除的行上带有批注，请先删除批注或保留该行。')
        node.set('ref', new)
        moved = True
    if moved:
        if root.get('count') is not None:
            root.set('count', str(sum(1 for n in root.iter() if _local(n) == 'comment')))
        extra[part] = _serialize(root)


def _shift_pivot_table(package, mappings, sheet, target, rels, cache_done, extra):
    """透视表的摆放区域随所在表移动；数据源区域随其工作表移动。"""
    part = _resolve('xl/worksheets', target)
    root = _xml(package.zip.read(part))
    location = next((n for n in root if _local(n) == 'location' and n.get('ref')), None)
    if location is not None and sheet in mappings:
        if _object_intersection(mappings[sheet], range_boundaries(location.get('ref'))):
            raise ValueError(f'{sheet} 的数据透视表区域包含增删行列的位置，请在透视区域外操作。')
        location.set('ref', mappings[sheet].ref(location.get('ref')))
        extra[part] = _serialize(root)
    cache_rels = f'xl/pivotTables/_rels/{part.rsplit("/", 1)[-1]}.rels'
    if cache_rels in package.zip.namelist() and part not in cache_done:
        cache_done.add(part)
        for relation in _xml(package.zip.read(cache_rels)):
            if not relation.get('Type', '').endswith('/pivotCacheDefinition'):
                continue
            cache = _resolve('xl/pivotTables', relation.get('Target'))
            cache_root = _xml(package.zip.read(cache))
            source = cache_root.find(f'{_tag("cacheSource")}/{_tag("worksheetSource")}')
            name = source.get('sheet') if source is not None else None
            ref = source.get('ref') if source is not None else None
            if not ref or name is None:
                continue
            mapping = next((m for n, m in mappings.items() if n.casefold() == name.casefold()), None)
            if mapping is None:
                continue
            moved = mapping.ref(ref)
            if moved is None:
                raise ValueError(f'数据透视表数据源 {name}!{ref} 的行已被删除，请保留这些行。')
            source.set('ref', moved)
            extra[cache] = _serialize(cache_root)


def validate_layout(package, infos, pending, layouts):
    checked = {}
    for name, layout in layouts.items():
        if name not in infos:
            raise ValueError(f'工作表不存在：{name}')
        heights = {int(r): float(h) for r, h in layout.get('heights', {}).items()}
        widths = {int(c): float(w) for c,w in layout.get('widths',{}).items()}
        deleted = sorted(set(int(r) for r in layout.get('delete_rows', [])))
        wraps = sorted(set(layout.get('wrap_cells', [])))
        if not heights and not deleted and not wraps and not widths:
            continue
        root = _xml(package.zip.read(dict(package.sheets)[name]))
        if root.find(_tag('sheetProtection')) is not None:
            raise ValueError(f'{name} 已保护，请先解除工作表保护。')
        if any(not 1 <= r <= 1048576 or not math.isfinite(h) or not 1 <= h <= 409 for r, h in heights.items()):
            raise ValueError('行高须在 1–409 磅之间。')
        if any(not 1<=c<=16384 or not math.isfinite(w) or not 1<=w<=255 for c,w in widths.items()):
            raise ValueError('列宽须在 1–255 之间。')
        if deleted != empty_rows(infos[name], pending.get(name, {}), deleted):
            raise ValueError('待删除行已有内容、公式或跨行合并，请重新选择空行。')
        if any(not re.fullmatch(r'[A-Z]{1,3}[1-9]\d*', ref) or range_boundaries(ref)[0] > 16384 or range_boundaries(ref)[1] > 1048576 for ref in wraps):
            raise ValueError('自动换行的单元格地址无效。')
        checked[name] = {'heights': heights, 'delete_rows': deleted, 'wrap_cells': wraps,'widths':widths}
    return checked


def apply_layout_xml(package, roots, layouts, row_maps=None):
    """修改已写值的 XML；不能可靠移动的独立对象明确拒绝，绝不丢弃。"""
    mappings = row_maps if row_maps is not None else {name: RowMap(v['delete_rows']) for name, v in layouts.items() if v['delete_rows']}
    changed = set(layouts)
    styles = _xml(package.zip.read('xl/styles.xml'))
    xfs = styles.find(_tag('cellXfs'))
    wrap_styles = {}
    for name, layout in layouts.items():
        root = roots[name]
        data = root.find(_tag('sheetData'))
        if layout.get('widths'):
            cols=root.find(_tag('cols'))
            if cols is None:
                cols=ET.Element(_tag('cols'));root.insert(list(root).index(data),cols)
            for column,width in layout['widths'].items():
                for dim in list(cols):
                    a,b=int(dim.get('min')),int(dim.get('max'))
                    if a<=column<=b:
                        cols.remove(dim)
                        if a<column:
                            before=deepcopy(dim);before.set('max',str(column-1));cols.append(before)
                        if b>column:
                            after=deepcopy(dim);after.set('min',str(column+1));cols.append(after)
                ET.SubElement(cols,_tag('col'),min=str(column),max=str(column),width=str(width),customWidth='1')
            cols[:]=sorted(cols,key=lambda c:int(c.get('min')))
        rows = {int(r.get('r')): r for r in data}
        for rownum, height in layout['heights'].items():
            if rownum not in rows:
                rows[rownum] = ET.SubElement(data, _tag('row'), r=str(rownum))
            rows[rownum].set('ht', str(height)); rows[rownum].set('customHeight', '1')
        wraps = set(layout['wrap_cells'])
        for row in data:
            for cell in row:
                if cell.get('r') not in wraps:
                    continue
                index = int(cell.get('s', '0'))
                if index not in wrap_styles:
                    xf = deepcopy(xfs[index])
                    align = xf.find(_tag('alignment'))
                    if align is None:
                        align = ET.SubElement(xf, _tag('alignment'))
                    align.set('wrapText', '1'); xf.set('applyAlignment', '1')
                    wrap_styles[index] = len(xfs); xfs.append(xf)
                cell.set('s', str(wrap_styles[index]))
        data[:] = sorted(data, key=lambda r: int(r.get('r')))
    extra = {}
    if mappings:
        # 图表、透视表缓存的坐标引用随之改写；穿越其区域内部的操作在函数内拒绝。
        extra.update(_shift_object_parts(package, mappings))
        unsupported = {'legacyDrawingHF', 'tableParts', 'oleObjects', 'controls', 'extLst', 'protectedRanges', 'scenarios', 'customSheetViews'}
        for name, root in roots.items():
            if name in mappings and any(ET.QName(n).localname in unsupported for n in root):
                raise ValueError(f'{name} 含列表、ActiveX 控件等特殊对象，暂不能增删行；仍可调整行高。')
            masters={f.get('si'):(f.getparent().get('r'),f.text) for f in root.iter(_tag('f')) if f.get('t')=='shared' and f.text}
            for f in root.iter(_tag('f')):
                if f.get('t')=='shared':
                    if f.get('si') not in masters:raise ValueError('共享公式缺少主公式，无法可靠移动。')
                    origin,text=masters[f.get('si')]
                    f.text=Translator('='+text,origin=origin).translate_formula(f.getparent().get('r'))[1:]
                    for attr in ('t','si','ref'):f.attrib.pop(attr,None)
                    changed.add(name)
                elif f.get('t')=='dataTable':
                    raise ValueError('该表包含模拟运算表，请先转换为普通公式。')
                elif f.get('t')=='array' and name in mappings and f.get('ref'):
                    _,top,_,bottom=range_boundaries(f.get('ref'))
                    removed=set(mappings[name].deleted)
                    if removed.intersection(range(top,bottom+1)) and not set(range(top,bottom+1)).issubset(removed):
                        raise ValueError('不能只删除数组公式区域的一部分，请选择完整区域。')
                    moved=mappings[name].ref(f.get('ref'))
                    if moved:f.set('ref',moved)
            for tag in ('f', 'formula', 'formula1', 'formula2'):
                for node in root.iter(_tag(tag)):
                    if tag=='f' and name in mappings and getattr(mappings[name],'axis','row')=='row' and mappings[name].display(range_boundaries(node.getparent().get('r'))[1]) is None:
                        continue  # 整行删除的公式无需再改写其引用。
                    shifted = shift_formula(node.text, name, mappings)
                    if shifted != node.text:
                        node.text = shifted; changed.add(name)
                    if tag == 'f':
                        cached = node.getparent().find(_tag('v'))
                        if cached is not None:
                            node.getparent().remove(cached); changed.add(name)
            for node in root.iter(_tag('hyperlink')):
                if node.get('location'):
                    old = node.get('location')
                    new = shift_formula(old, name, mappings)
                    if new != old:
                        node.set('location', new); changed.add(name)
            if name not in mappings:
                continue
            mapping = mappings[name]
            column_mode=getattr(mapping,'axis','row')=='column'
            data = root.find(_tag('sheetData'))
            for row in list(data):
                newrow = int(row.get('r')) if column_mode else mapping.display(int(row.get('r')))
                if newrow is None:
                    data.remove(row); continue
                row.set('r', str(newrow))
                for cell in row:
                    cell.set('r', mapping.ref(cell.get('r')))
            if column_mode:
                for row in data:row.attrib.pop('spans',None)
                for auto in root.findall(_tag('autoFilter')):
                    old_left=range_boundaries(auto.get('ref'))[0]
                    new_left=mapping.display(old_left)
                    for column in auto.findall(_tag('filterColumn')):
                        column.set('colId',str(mapping.display(old_left+int(column.get('colId')))-new_left))
                for dim in root.findall(f"{_tag('cols')}/{_tag('col')}"):
                    dim.set('min',str(mapping.display(int(dim.get('min')))))
                    dim.set('max',str(mapping.display(int(dim.get('max')))))
            for node in list(root.iter()):
                tag = ET.QName(node).localname
                if tag in {'mergeCell', 'dimension', 'autoFilter', 'hyperlink', 'sortState', 'sortCondition'} and node.get('ref'):
                    ref = mapping.ref(node.get('ref'))
                    if ref:
                        node.set('ref', ref)
                    else:
                        node.getparent().remove(node); continue
                for attr in ('sqref', 'activeCell', 'topLeftCell'):
                    if node.get(attr):
                        refs = [new for ref in node.get(attr).split() if (new := mapping.ref(ref))]
                        if refs:
                            node.set(attr, ' '.join(refs))
                        elif tag in {'conditionalFormatting', 'dataValidation'}:
                            node.getparent().remove(node)
                        else:
                            node.set(attr, 'A1')
                if tag == 'pane' and node.get('state') in {'frozen', 'frozenSplit'}:
                    attr='xSplit' if column_mode else 'ySplit'
                    split = int(float(node.get(attr, '0')))
                    node.set(attr, str(mapping.boundary(split)))
                if tag == 'brk' and ET.QName(node.getparent()).localname == ('colBreaks' if column_mode else 'rowBreaks'):
                    old = int(node.get('id'))
                    node.set('id', str(max(1, mapping.boundary(old))))
            for tag in ('mergeCells', 'dataValidations', 'rowBreaks'):
                node = root.find(_tag(tag))
                if node is not None:
                    node.set('count', str(len(node)))
        names = package.book.find(_tag('definedNames'))
        if names is not None:
            for node in names:
                local = node.get('localSheetId')
                current = package.sheets[int(local)][0] if local is not None else None
                node.text = shift_formula(node.text, current, mappings)
    parts = {dict(package.sheets)[name]: _serialize(roots[name]) for name in changed}
    parts.update(extra)
    if wrap_styles:
        xfs.set('count', str(len(xfs)))
        parts['xl/styles.xml'] = _serialize(styles)
    return parts
