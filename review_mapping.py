"""配置化填表：先生成可审查的单元格变更计划，再原子提交。"""
import hashlib
import json
from copy import deepcopy
from zipfile import ZipFile
from pathlib import Path

from lxml import etree as ET
from openpyxl.utils.cell import get_column_letter, column_index_from_string, range_boundaries

from app_settings import load_settings, save_settings
from file_io import atomic_output_path
from review_table_filler import _Package, _xml, _tag, _norm, _write, _serialize, HEADERS

FIELDS = ("factor", "description", "points")


def inspect_workbook(path, package=None):
    package = package if package is not None else _Package(path)
    result = {}
    for name, part in package.sheets:
        root = _xml(package.zip.read(part))
        data = root.find(_tag("sheetData"))
        if data is None:  # 图表工作表不作为可写单元格工作表
            continue
        cells = {c.get("r"): package.value(c) for row in data for c in row}
        merges = [m.get("ref") for m in root.findall(f"{_tag('mergeCells')}/{_tag('mergeCell')}")]
        max_row = max((range_boundaries(ref)[1] for ref in cells), default=1)
        max_col = max((range_boundaries(ref)[0] for ref in cells), default=1)
        guess = {"header": 1, "start": 2, "end": max(2, max_row),
                 "factor": "A", "description": "B", "points": "C"}
        for row in range(1, min(max_row, 30) + 1):
            found = [[get_column_letter(col) for col in range(1, max_col + 1)
                      if _norm(cells.get(f"{get_column_letter(col)}{row}")) in aliases]
                     for aliases in HEADERS]
            if all(len(columns) == 1 for columns in found):
                guess.update(dict(zip(FIELDS, [columns[0] for columns in found])))
                start = row + 1
                for ref in merges:
                    left, top, right, bottom = range_boundaries(ref)
                    if top <= row <= bottom:
                        start = max(start, bottom + 1)
                totals = [r for r in range(start, max_row + 1) if any(
                    _norm(cells.get(f"{get_column_letter(c)}{r}")) in {"合计", "总计", "总分"}
                    for c in range(1, max_col + 1))]
                guess.update(header=row, start=start, end=(totals[0]-1 if totals else max_row))
                break
        result[name] = {"cells": cells, "merges": merges, "max_row": max_row,
                        "max_col": max_col, "guess": guess}
    return result


def layout_signature(info, config):
    """模板结构标识不包含项目正文，可跨项目复用；表头/合并布局变动则提示。"""
    header = int(config["header"])
    payload = [info["max_col"], info["merges"],
               [(ref, value) for ref, value in sorted(info["cells"].items())
                if range_boundaries(ref)[1] == header]]
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False).encode()).hexdigest()


def save_profile(name, config, source_mapping, info):
    if not name.strip():
        raise ValueError("请填写配置名称。")
    profiles = load_profiles()
    profiles[name.strip()] = {"config": deepcopy(config), "source": deepcopy(source_mapping),
                              "signature": layout_signature(info, config)}
    save_settings({"review_mapping_profiles": profiles})


def load_profiles():
    value = load_settings().get("review_mapping_profiles", {})
    return value if isinstance(value, dict) else {}


def _anchor(ref, merges):
    col, row, _, _ = range_boundaries(ref)
    for merged in merges:
        left, top, right, bottom = range_boundaries(merged)
        if left <= col <= right and top <= row <= bottom:
            return f"{get_column_letter(left)}{top}"
    return ref


def plan_fill(path, rules, config, overrides=None):
    """不模糊猜测行。重复名称/缺失名称必须手动指定；0 表示明确跳过。"""
    package = _Package(path)
    sheet = config["sheet"]
    info = inspect_workbook(path, package).get(sheet)
    if info is None:
        raise ValueError("目标工作表不存在，请重新选择。")
    start, end, header = (int(config[k]) for k in ("start", "end", "header"))
    if not 1 <= header < start <= end <= 1048576:
        raise ValueError("请指定有效的表头行和填入范围（起始行须在表头之后）。")
    if end - start > 10000:
        raise ValueError("填入范围超过一万行，请缩小到实际表格区域。")
    columns = [str(config[k]).strip().upper() for k in FIELDS]
    try:
        if len(set(columns)) != 3 or any(not 1 <= column_index_from_string(c) <= 16384 for c in columns):
            raise ValueError()
    except ValueError:
        raise ValueError("三个目标列须为不同的有效 Excel 列字母。") from None
    mode = config.get("mode", "match")
    if mode not in {"match", "sequence"}:
        raise ValueError("未知填入方式。")
    root = _xml(package.zip.read(dict(package.sheets)[sheet]))
    if root.find(_tag("sheetProtection")) is not None:
        raise ValueError("目标工作表受保护，请先解除保护。")
    nodes = {c.get("r"): c for row in root.find(_tag("sheetData")) for c in row}
    formula_ranges = [range_boundaries(f.get("ref")) for f in root.iter(_tag("f")) if f.get("ref")]
    matches = {}
    for row in range(start, end + 1):
        ref = f"{columns[0]}{row}"
        key = _norm(info["cells"].get(_anchor(ref, info["merges"]), ""))
        if key:
            matches.setdefault(key, []).append(row)
    changes, entries, errors, used_rows = {}, [], [], set()
    overrides = overrides or {}
    for index, criterion in enumerate(rules.criteria):
        candidates = matches.get(_norm(criterion.factor), [])
        row = overrides.get(index, start + index if mode == "sequence" else
                            candidates[0] if len(candidates) == 1 else None)
        entry = {"index": index, "name": criterion.factor, "description": criterion.description,
                 "points": str(criterion.points), "row": row, "status": "待确认"}
        entries.append(entry)
        if row == 0:
            entry["status"] = "已跳过"
            continue
        if row is None:
            reason = "名称重复，请指定行" if candidates else "未匹配，请指定行或跳过"
            entry["status"] = reason
            errors.append(f"{criterion.factor}：{reason}")
            continue
        if not isinstance(row, int) or not start <= row <= end or row in used_rows:
            errors.append(f"{criterion.factor}：目标行超出范围或重复使用")
            entry["status"] = "目标行冲突"
            continue
        used_rows.add(row)
        # 按名称匹配时保留模板项目名称；顺序填入时写入名称。
        pairs = list(zip(columns, [criterion.factor, criterion.description, criterion.points]))
        if mode == "match":
            pairs = pairs[1:]
        entry["status"] = "可填入"
        for column, value in pairs:
            requested = f"{column}{row}"
            ref = _anchor(requested, info["merges"])
            col, anchor_row, _, _ = range_boundaries(ref)
            if not start <= anchor_row <= end or get_column_letter(col) != column:
                errors.append(f"{requested} 的合并区域跨越所选字段或填入范围")
                entry["status"] = "合并区域冲突"
                continue
            node = nodes.get(ref)
            if (node is not None and node.find(_tag("f")) is not None) or any(
                    left <= col <= right and top <= anchor_row <= bottom
                    for left, top, right, bottom in formula_ranges):
                errors.append(f"{ref} 包含公式，不能覆盖")
                entry["status"] = "公式冲突"
                continue
            if ref in changes and changes[ref]["new"] != value:
                errors.append(f"{ref} 合并单元格对应多个不同内容")
                entry["status"] = "合并区域冲突"
                continue
            before = info["cells"].get(ref, "")
            if str(before) != str(value):
                changes[ref] = {"cell": ref, "old": before, "new": value}
                if before != "":
                    entry["status"] = "将更新已有内容"
    return {"sheet": sheet, "start": start, "end": end, "digest": hashlib.sha256(package.raw).hexdigest(),
            "entries": entries, "changes": list(changes.values()), "errors": errors}


def apply_plan(path, output_path, plan):
    if Path(path).resolve() == Path(output_path).resolve() or Path(output_path).exists():
        raise ValueError("请另存到不存在的新文件，不能覆盖模板或已有文件。")
    if plan["errors"]:
        raise ValueError("请先解决预览中的未匹配项或冲突。")
    package = _Package(path)
    if hashlib.sha256(package.raw).hexdigest() != plan["digest"]:
        raise ValueError("模板在预览后发生变化，请重新预览。")
    if not plan["changes"]:
        raise ValueError("没有需要写入的变更。")
    part = dict(package.sheets)[plan["sheet"]]
    root = _xml(package.zip.read(part))
    data = root.find(_tag("sheetData"))
    rows = {int(row.get("r")): row for row in data}
    for change in plan["changes"]:
        ref = change["cell"]
        col, rownum, _, _ = range_boundaries(ref)
        row = rows.get(rownum)
        if row is None:
            row = ET.SubElement(data, _tag("row"), r=str(rownum))
            rows[rownum] = row
        cell = next((c for c in row if c.get("r") == ref), None)
        if cell is None:
            cell = ET.SubElement(row, _tag("c"), r=ref)
            # 只补缺失节点，已有单元格样式不变。
            for other in sorted(rows, key=lambda r: abs(r-rownum)):
                if not plan["start"] <= other <= plan["end"]:
                    continue
                prototype = next((c for c in rows[other] if c.get("r") == f"{get_column_letter(col)}{other}"), None)
                if prototype is not None and prototype is not cell and prototype.get("s"):
                    cell.set("s", prototype.get("s"))
                    break
        _write(cell, change["new"])
        row[:] = sorted(row, key=lambda c: range_boundaries(c.get("r"))[0])
    data[:] = sorted(data, key=lambda r: int(r.get("r")))
    dimension = root.find(_tag("dimension"))
    if dimension is not None:
        bounds = [range_boundaries(c.get("r")) for row in data for c in row]
        dimension.set("ref", f"{get_column_letter(min(b[0] for b in bounds))}{min(b[1] for b in bounds)}:"
                      f"{get_column_letter(max(b[0] for b in bounds))}{max(b[1] for b in bounds)}")
    calc = package.book.find(_tag("calcPr"))
    if calc is None:
        calc = ET.SubElement(package.book, _tag("calcPr"))
    calc.set("fullCalcOnLoad", "1")
    calc.set("forceFullCalc", "1")
    calc.set("calcMode", "auto")
    changed = {part: _serialize(root), "xl/workbook.xml": _serialize(package.book)}
    with atomic_output_path(output_path) as temporary:
        with ZipFile(temporary, "w") as output:
            output.comment = package.zip.comment
            for info in package.zip.infolist():
                output.writestr(info, changed.get(info.filename, package.zip.read(info.filename)))
    return {"output_path": str(output_path), "sheet_name": plan["sheet"], "count": len(plan["changes"])}
