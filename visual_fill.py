"""可视化填表的拖放计划；拖动只记录映射，提交时一次写入工作簿。"""
import hashlib
import re
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from lxml import etree as ET
from openpyxl.utils.cell import get_column_letter, range_boundaries

from review_mapping import inspect_workbook, _anchor
from review_table_filler import _Package, _xml, _tag, _serialize, _write
from worksheet_layout import validate_layout, apply_layout_xml


def source_grid(table):
    widths = {}
    for column in range(1, table.column_count+1):
        contents = [cell.text for cell in table.cells if cell.column==column]
        widths[column] = 300 if any(len(text)>30 for text in contents) else 150
        if any(text.strip() in {'分值','满分','分数','数量','序号'} for text in contents[:3]):
            widths[column] = 80
    merges = [f"{get_column_letter(c.column)}{c.row}:{get_column_letter(c.column+c.column_span-1)}{c.row+c.row_span-1}"
              for c in table.cells if c.row_span > 1 or c.column_span > 1]
    spans = [range_boundaries(ref) for ref in merges]
    # 合并区域可能延伸到表格行数之外，几何范围须覆盖完整合并。
    return {"cells": {f"{get_column_letter(c.column)}{c.row}": c.text for c in table.cells},
            "widths": widths,
            "merges": merges,
            "max_row": max(table.row_count, max((b[3] for b in spans), default=0)),
            "max_col": max(table.column_count, max((b[2] for b in spans), default=0))}


def inspect_visual_workbook(path, package=None, base_info=None):
    package = package if package is not None else _Package(path)
    result = base_info if base_info is not None else inspect_workbook(path,package)
    styles = _xml(package.zip.read("xl/styles.xml"))
    fonts, fills, xfs = (styles.find(_tag(n)) for n in ("fonts", "fills", "cellXfs"))
    borders=styles.find(_tag('borders'))
    palette = {}
    for i, xf in enumerate(xfs):
        font = fonts[int(xf.get("fontId", 0))]
        fill = fills[int(xf.get("fillId", 0))]
        fg = fill.find(f"{_tag('patternFill')}/{_tag('fgColor')}")
        color = font.find(_tag("color"))
        size, family = font.find(_tag('sz')), font.find(_tag('name'))
        align = xf.find(_tag('alignment'))
        border=borders[int(xf.get('borderId',0))] if borders is not None else None
        edges={}
        if border is not None:
            for side in ('left','right','top','bottom'):
                edge=border.find(_tag(side))
                if edge is not None and edge.get('style'):
                    edge_color=edge.find(_tag('color'))
                    edges[side]={'style':edge.get('style'),'color':'#'+edge_color.get('rgb')[-6:] if edge_color is not None and edge_color.get('rgb') else '#333333'}
        palette[str(i)] = {"bold": font.find(_tag("b")) is not None,
                           'italic':font.find(_tag('i')) is not None,'underline':font.find(_tag('u')) is not None,'borders':edges,
                           "size": float(size.get('val', 11)) if size is not None else 11,
                           "family": family.get('val', '宋体') if family is not None else '宋体',
                           "indent": int(align.get('indent', 0)) if align is not None else 0,
                           "rotation": int(align.get('textRotation', 0)) if align is not None else 0,
                           "horizontal": align.get('horizontal','general') if align is not None else 'general',
                           "vertical": align.get('vertical','bottom') if align is not None else 'bottom',
                           "wrap": align.get('wrapText','0')=='1' if align is not None else False,
                           "bg": "#"+fg.get("rgb")[-6:] if fg is not None and fg.get("rgb") else None,
                           "fg": "#"+color.get("rgb")[-6:] if color is not None and color.get("rgb") else None}
    for name, part in package.sheets:
        if name not in result:
            continue
        info = result[name]
        root = _xml(package.zip.read(part))
        info['protected'] = root.find(_tag('sheetProtection')) is not None
        fmt = root.find(_tag('sheetFormatPr'))
        default_width = float(fmt.get('defaultColWidth', 12)) if fmt is not None else 12
        default_height = float(fmt.get('defaultRowHeight', 15)) if fmt is not None else 15
        info["widths"] = {}
        info['column_pixels'] = {c: default_width * 7 + 5 for c in range(1, info['max_col'] + 1)}
        for column in root.findall(f"{_tag('cols')}/{_tag('col')}"):
            for c in range(int(column.get("min")), min(int(column.get("max")), 16384)+1):
                info['column_pixels'][c] = float(column.get('width', default_width)) * 7 + 5
                info["widths"][c] = max(12, info['column_pixels'][c])
        info['default_width'] = default_width * 7 + 5
        info['default_height'] = default_height * 4 / 3
        info["heights"] = {int(row.get("r")): float(row.get("ht", default_height))*4/3
                           for row in root.find(_tag("sheetData"))}
        info['formulas'] = {cell.get('r'): cell.find(_tag('f')).text or ''
                            for cell in root.iter(_tag('c')) if cell.find(_tag('f')) is not None}
        info['formula_ranges'] = [f.get('ref') for f in root.iter(_tag('f')) if f.get('ref')]
        info['formula_display'] = {cell.get('r'): cell.findtext(_tag('v')) or ''
                                  for cell in root.iter(_tag('c')) if cell.find(_tag('f')) is not None}
        info["styles"] = {cell.get("r"): palette.get(cell.get("s", "0"), {})
                          for row in root.find(_tag("sheetData")) for cell in row}
        for merged in info["merges"]:
            _, _, right, bottom = range_boundaries(merged)
            info["max_row"] = max(info["max_row"], bottom)
            info["max_col"] = max(info["max_col"], right)
    return result


def expand_selection(info, bounds):
    """框选碰到合并单元格时扩展到整个区域。坐标为 1 起始闭区间。"""
    left, top, right, bottom = bounds
    if not 1 <= left <= right <= info["max_col"] or not 1 <= top <= bottom <= info["max_row"]:
        raise ValueError("所选区域超出来源表格。")
    changed = True
    while changed:
        changed = False
        for ref in info["merges"]:
            l, t, r, b = range_boundaries(ref)
            if l <= right and r >= left and t <= bottom and b >= top:
                new = min(left,l), min(top,t), max(right,r), max(bottom,b)
                if new != (left,top,right,bottom):
                    left,top,right,bottom = new
                    changed = True
    return left,top,right,bottom


def _value(text):
    if re.fullmatch(r"-?(?:0|[1-9]\d*)(?:\.\d+)?", str(text)):
        return Decimal(text)
    return str(text)


def drop_values(source, bounds, target, row, column):
    left, top, right, bottom = expand_selection(source, bounds)
    if row < 1 or column < 1 or row+bottom-top > 1048576 or column+right-left > 16384:
        raise ValueError("拖入位置超出 Excel 行列范围。")
    grouped = defaultdict(list)
    # 合并来源只传递锚点一次，不复制覆盖格的重复分值。
    for r in range(top,bottom+1):
        for c in range(left,right+1):
            ref = f"{get_column_letter(c)}{r}"
            if _anchor(ref,source["merges"]) != ref:
                continue
            text = source["cells"].get(ref, "")
            if text == "":  # 空白不清除模板内容
                continue
            dest = _anchor(f"{get_column_letter(column+c-left)}{row+r-top}", target["merges"])
            grouped[dest].append(_value(text))
    result = {}
    for dest, values in grouped.items():
        if len(values)>1 and any(isinstance(v,Decimal) for v in values):
            raise ValueError(f"{dest} 是合并单元格，不能同时放入多个独立分值。请分别拖入。")
        result[dest] = values[0] if len(values)==1 else "\n\n".join(str(v) for v in values)
    return result


def plan_visual_fill(path, pending, layouts=None, *, package=None, info=None):
    """pending 为 {工作表名: {单元格: 值}}，可供界面及 Agent 共用。"""
    package = package if package is not None else _Package(path)
    info = info if info is not None else inspect_workbook(path,package)
    changes, errors = {}, []
    for name, values in pending.items():
        if not values:
            continue
        if name not in info:
            raise ValueError(f"工作表不存在：{name}")
        root = _xml(package.zip.read(dict(package.sheets)[name]))
        if root.find(_tag("sheetProtection")) is not None:
            raise ValueError(f"{name} 已保护，请先解除工作表保护。")
        nodes = {c.get("r"): c for row in root.find(_tag("sheetData")) for c in row}
        formula_ranges = [range_boundaries(f.get("ref")) for f in root.iter(_tag("f")) if f.get("ref")]
        sheet_changes = []
        normalized = {}
        for ref, value in values.items():
            if not re.fullmatch(r"[A-Z]{1,3}[1-9]\d*",ref):
                raise ValueError("无效单元格地址。")
            col,row,_,_ = range_boundaries(ref)
            if row>1048576 or col>16384:
                raise ValueError("单元格超出 Excel 范围。")
            ref = _anchor(ref,info[name]["merges"])
            if ref in normalized and normalized[ref] != value:
                raise ValueError(f"{name}!{ref} 存在相互冲突的拖入内容。")
            normalized[ref] = value
        for ref,value in normalized.items():
            col,row,_,_ = range_boundaries(ref)
            node = nodes.get(ref)
            if (node is not None and node.find(_tag("f")) is not None) or any(
                    l<=col<=r and t<=row<=b for l,t,r,b in formula_ranges):
                errors.append(f"{name}!{ref} 包含公式，不能覆盖")
                continue
            if isinstance(value,str) and len(value)>32767:
                errors.append(f"{name}!{ref} 内容超过 Excel 单元格长度上限")
                continue
            before = info[name]["cells"].get(ref,"")
            if str(before)!=str(value):
                sheet_changes.append({"cell":ref,"old":before,"new":value})
        if sheet_changes:
            changes[name] = sheet_changes
    layouts = validate_layout(package, inspect_visual_workbook(path,package,info), pending, layouts) if layouts else {}
    if layouts:
        # 在预览阶段执行同一结构校验，保存时不会首次发现不支持的对象。
        roots = {n: _xml(package.zip.read(p)) for n, p in package.sheets if n in info}
        apply_layout_xml(package, roots, layouts)
    return {"digest":hashlib.sha256(package.raw).hexdigest(),"sheets":changes,"errors":errors,
            "count":sum(map(len,changes.values())), 'layouts': layouts,
            'layout_count': sum(len(v['heights']) + len(v['delete_rows']) + len(v.get('widths',{})) for v in layouts.values())}


def apply_visual_plan(path, output_path, plan, progress=None):
    if Path(path).resolve()==Path(output_path).resolve() or Path(output_path).exists():
        raise ValueError("请另存为新文件，不能覆盖模板或已有文件。")
    package = _Package(path)
    if hashlib.sha256(package.raw).hexdigest()!=plan["digest"]:
        raise ValueError("模板在预览后发生变化，请重新打开文件。")
    if plan["errors"] or not (plan["count"] or plan.get('layout_count')):
        raise ValueError("请先处理预览中的问题，或拖入需要填写的内容。")
    # 再验证一次，防止调用方绕开保护/公式检查。
    checked = plan_visual_fill(path,{name:{c["cell"]:c["new"] for c in values} for name,values in plan["sheets"].items()}, plan.get('layouts'), package=None if plan.get('layouts') else package)
    if checked["errors"]:
        raise ValueError("；".join(checked["errors"]))
    changed = {}
    roots = {n: _xml(package.zip.read(p)) for n, p in package.sheets
             if _xml(package.zip.read(p)).find(_tag('sheetData')) is not None}
    for name, values in checked["sheets"].items():
        part = dict(package.sheets)[name]
        root = roots[name]
        data = root.find(_tag("sheetData"))
        rows = {int(r.get("r")):r for r in data}
        for change in values:
            ref = change["cell"]
            _,r,_,_ = range_boundaries(ref)
            if r not in rows:
                rows[r] = ET.SubElement(data,_tag("row"),r=str(r))
            row = rows[r]
            cell = next((c for c in row if c.get("r")==ref),None)
            if cell is None:
                cell = ET.SubElement(row,_tag("c"),r=ref)
            _write(cell,change["new"])
            row[:] = sorted(row,key=lambda c:range_boundaries(c.get("r"))[0])
        data[:] = sorted(data,key=lambda r:int(r.get("r")))
        dimension = root.find(_tag("dimension"))
        if dimension is not None:
            bounds = [range_boundaries(c.get("r")) for r in data for c in r]
            dimension.set("ref",f"{get_column_letter(min(b[0] for b in bounds))}{min(b[1] for b in bounds)}:"
                          f"{get_column_letter(max(b[0] for b in bounds))}{max(b[1] for b in bounds)}")
        changed[part] = _serialize(root)
    changed.update(apply_layout_xml(package, roots, checked['layouts']))
    calc = package.book.find(_tag("calcPr"))
    if calc is None:
        calc = ET.SubElement(package.book,_tag("calcPr"))
    calc.attrib.update({"fullCalcOnLoad":"1","forceFullCalc":"1","calcMode":"auto"})
    changed["xl/workbook.xml"] = _serialize(package.book)
    from workbook_structure import write_package
    write_package(package,output_path,changed,any(v['delete_rows'] for v in checked['layouts'].values()),progress)
    return {"output_path":str(output_path),"count":checked["count"],
            'layout_count': checked['layout_count'],
            "sheets":list(dict.fromkeys([*checked["sheets"], *checked['layouts']]))}
