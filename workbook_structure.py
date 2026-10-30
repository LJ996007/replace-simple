"""工作簿结构编辑：在独立预览副本上编辑，保留未涉及的压缩包部件。"""
import hashlib
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET
from openpyxl.utils.cell import get_column_letter, range_boundaries

from file_io import atomic_output_path
from review_table_filler import _Package, _tag, _xml, _serialize, _write
from worksheet_layout import RowMap, InsertRowMap, InsertColumnMap, apply_layout_xml


def write_package(package, output, changed, clear_chain=False, progress=None):
    """按真实写入部件数报告进度，最后一步为原子提交。"""
    removed = set()
    if clear_chain and 'xl/calcChain.xml' in package.zip.namelist():
        removed.add('xl/calcChain.xml')
        for part in ('[Content_Types].xml', 'xl/_rels/workbook.xml.rels'):
            root = _xml(package.zip.read(part))
            for child in list(root):
                if child.get('PartName') == '/xl/calcChain.xml' or child.get('Type', '').endswith('/calcChain'):
                    root.remove(child)
            changed[part] = _serialize(root)
    entries = [e for e in package.zip.infolist() if e.filename not in removed]
    total = len(entries) + 1
    with atomic_output_path(output) as temporary:
        with ZipFile(temporary,'w') as archive:
            archive.comment = package.zip.comment
            for index, entry in enumerate(entries,1):
                archive.writestr(entry,changed.get(entry.filename,package.zip.read(entry.filename)))
                if progress:progress(index,total,'正在写入工作簿')
        if progress:progress(len(entries),total,'正在完成保存')
    if progress:progress(total,total,'保存完成')


def _merges_node(root):
    node = root.find(_tag('mergeCells'))
    if node is None:
        node = ET.Element(_tag('mergeCells'))
        # ECMA 工作表顺序：mergeCells 在 autoFilter 后、条件格式及校验前。
        later = {'phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions',
                 'pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties',
                 'cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','extLst'}
        index = next((i for i,n in enumerate(root) if ET.QName(n).localname in later),len(root))
        root.insert(index,node)
    return node


def edit_structure(path, output, sheet, action, bounds, count=1, expected_digest=None):
    """bounds 使用当前预览坐标，输出成功后调用方才切换快照。"""
    if Path(output).exists() or Path(path).resolve()==Path(output).resolve():
        raise ValueError('结构编辑必须使用新的预览副本。')
    package = _Package(path)
    if expected_digest and hashlib.sha256(package.raw).hexdigest()!=expected_digest:
        raise ValueError('模板文件已变化，请重新打开填表窗口。')
    roots = {n:_xml(package.zip.read(p)) for n,p in package.sheets}
    roots = {n:r for n,r in roots.items() if r.find(_tag('sheetData')) is not None}
    if sheet not in roots:raise ValueError('工作表不存在。')
    root = roots[sheet]
    if root.find(_tag('sheetProtection')) is not None:
        raise ValueError('工作表已保护，请先解除保护。')
    left,top,right,bottom = map(int,bounds)
    if not 1<=left<=right<=16384 or not 1<=top<=bottom<=1048576:
        raise ValueError('选区超出 Excel 范围。')
    data = root.find(_tag('sheetData'))
    cells = {c.get('r'):c for row in data for c in row}
    merges = [range_boundaries(n.get('ref')) for n in root.findall(f"{_tag('mergeCells')}/{_tag('mergeCell')}")]
    changed = {}
    structural = action in {'insert_above','insert_below','insert_left','insert_right','delete_rows'}
    if structural:
        column_mode=action in {'insert_left','insert_right'}
        if column_mode:
            count=int(count)
            if not 1<=count<=1000:raise ValueError('一次可插入 1–1000 列。')
            before=left if action=='insert_left' else right+1
            last=max([range_boundaries(ref)[0] for ref in cells]+[r for _,_,r,_ in merges]+[1])
            if max(last,before)+count>16384:raise ValueError('插入后超出 Excel 最大列数。')
            mapping=InsertColumnMap(before,count);prototype=None
            sample=left if action=='insert_left' else right
            column_samples={int(row.get('r')):deepcopy(next((c for c in row if range_boundaries(c.get('r'))[0]==sample),None)) for row in data}
            dim_sample=next((deepcopy(c) for c in root.findall(f"{_tag('cols')}/{_tag('col')}") if int(c.get('min'))<=sample<=int(c.get('max'))),None)
        elif action=='delete_rows':
            if bottom-top+1>10000:raise ValueError('一次最多删除 10000 行。')
            for _,t,_,b in merges:
                if t<=bottom and b>=top and not top<=t<=b<=bottom:
                    raise ValueError('所选行只包含跨行合并的一部分，请选择完整的合并区域。')
            mapping = RowMap(range(top,bottom+1))
            prototype = None
        else:
            count = int(count)
            if not 1<=count<=1000:raise ValueError('一次可插入 1–1000 行。')
            before = top if action=='insert_above' else bottom+1
            last = max([int(r.get('r')) for r in data]+[b for _,_,_,b in merges]+[1])
            if max(last,before)+count>1048576:raise ValueError('插入后超出 Excel 最大行数。')
            mapping = InsertRowMap(before,count)
            sample = top if action=='insert_above' else bottom
            prototype = deepcopy(next((r for r in data if int(r.get('r'))==sample),ET.Element(_tag('row'))))
        layout = {sheet:{'heights':{},'delete_rows':[],'wrap_cells':[]}}
        changed.update(apply_layout_xml(package,roots,layout,{sheet:mapping}))
        if column_mode:
            for row in data:
                sample_cell=column_samples[int(row.get('r'))]
                if sample_cell is not None:
                    for column in range(before,before+count):
                        cell=deepcopy(sample_cell);cell.set('r',f'{get_column_letter(column)}{row.get("r")}');_write(cell,None);row.append(cell)
                row[:]=sorted(row,key=lambda c:range_boundaries(c.get('r'))[0])
            if dim_sample is not None:
                cols=root.find(_tag('cols'))
                # 若原来一段列定义跨越插入点，先分段，避免重叠定义。
                for dim in list(cols):
                    a,b=int(dim.get('min')),int(dim.get('max'))
                    if a<before<=b:
                        dim.set('max',str(before-1))
                        tail=deepcopy(dim);tail.set('min',str(before+count));tail.set('max',str(b));cols.append(tail)
                dim_sample.set('min',str(before));dim_sample.set('max',str(before+count-1));cols.append(dim_sample)
                cols[:]=sorted(cols,key=lambda c:int(c.get('min')))
            merge_node=_merges_node(root)
            for l,t,r,b in merges:
                if l==r==sample:
                    for column in range(before,before+count):
                        ET.SubElement(merge_node,_tag('mergeCell'),ref=f'{get_column_letter(column)}{t}:{get_column_letter(column)}{b}')
            merge_node.set('count',str(len(merge_node)))
        if prototype is not None:
            merge_node = _merges_node(root)
            for rownum in range(before,before+count):
                new = deepcopy(prototype); new.set('r',str(rownum))
                new.attrib.pop('hidden',None)
                for cell in new:
                    col = range_boundaries(cell.get('r'))[0]
                    cell.set('r',f'{get_column_letter(col)}{rownum}')
                    _write(cell,None)
                data.append(new)
                for l,t,r,b in merges:
                    if t==b==sample:
                        ET.SubElement(merge_node,_tag('mergeCell'),ref=f'{get_column_letter(l)}{rownum}:{get_column_letter(r)}{rownum}')
            data[:] = sorted(data,key=lambda r:int(r.get('r')))
            merge_node.set('count',str(len(merge_node)))
        dim = root.find(_tag('dimension'))
        refs = [range_boundaries(c.get('r')) for row in data for c in row]
        if dim is not None:
            dim.set('ref',f'A1:{get_column_letter(max((b[0] for b in refs),default=1))}{max((b[1] for b in refs),default=1)}')
    elif action in {'merge','unmerge'}:
        if (right-left+1)*(bottom-top+1)>10000:raise ValueError('一次最多处理 10000 个单元格。')
        for l,t,r,b in merges:
            if l<=right and r>=left and t<=bottom and b>=top and not (left<=l<=r<=right and top<=t<=b<=bottom):
                raise ValueError('请选择完整的合并单元格。')
        node = _merges_node(root)
        affected = [n for n in node if (lambda b:left<=b[0]<=b[2]<=right and top<=b[1]<=b[3]<=bottom)(range_boundaries(n.get('ref')))]
        if action=='unmerge':
            if not affected:raise ValueError('选区中没有合并单元格。')
            for n in affected:node.remove(n)
        else:
            if left==right and top==bottom:raise ValueError('请至少选择两个单元格。')
            selected = [c for ref,c in cells.items() if left<=range_boundaries(ref)[0]<=right and top<=range_boundaries(ref)[1]<=bottom]
            formula_ranges = [range_boundaries(f.get('ref')) for f in root.iter(_tag('f')) if f.get('ref')]
            if any(c.find(_tag('f')) is not None for c in selected) or any(l<=right and r>=left and t<=bottom and b>=top for l,t,r,b in formula_ranges):
                raise ValueError('选区包含公式，不能合并；请保留公式所在单元格。')
            selected.sort(key=lambda c:(range_boundaries(c.get('r'))[1],range_boundaries(c.get('r'))[0]))
            values = [package.value(c) for c in selected if package.value(c) not in ('',None)]
            if len(values)>1 and any(c.get('t') not in {'s','inlineStr','str'} and package.value(c) not in ('',None) for c in selected):
                raise ValueError('选区含多个值且有数值，合并可能改变分值含义，请分别保留。')
            value = values[0] if len(values)==1 else '\n'.join(str(v) for v in values)
            if len(str(value))>32767:raise ValueError('合并后内容超过 Excel 单元格长度上限。')
            anchor = f'{get_column_letter(left)}{top}'
            if anchor not in cells:
                row = next((r for r in data if int(r.get('r'))==top),None)
                if row is None:row = ET.SubElement(data,_tag('row'),r=str(top))
                cells[anchor] = ET.SubElement(row,_tag('c'),r=anchor)
                row[:] = sorted(row,key=lambda c:range_boundaries(c.get('r'))[0])
            # 单个数值保留原 XML 类型（避免将分值改为文本）。
            numeric = next((c for c in selected if package.value(c) not in ('',None)),None)
            if len(values)==1 and numeric is not None and numeric.get('t') not in {'s','inlineStr','str'}:
                from decimal import Decimal
                value = Decimal(str(value))
            for c in selected:_write(c,None)
            _write(cells[anchor],value)
            for n in affected:node.remove(n)
            ET.SubElement(node,_tag('mergeCell'),ref=f'{anchor}:{get_column_letter(right)}{bottom}')
            data[:] = sorted(data,key=lambda r:int(r.get('r')))
        node.set('count',str(len(node)))
    else:
        raise ValueError('未知的结构编辑操作。')
    changed[dict(package.sheets)[sheet]] = _serialize(root)
    calc = package.book.find(_tag('calcPr'))
    if calc is None:calc = ET.SubElement(package.book,_tag('calcPr'))
    calc.attrib.update({'fullCalcOnLoad':'1','forceFullCalc':'1','calcMode':'auto'})
    changed['xl/workbook.xml'] = _serialize(package.book)
    write_package(package,output,changed,structural)
    return str(output)
