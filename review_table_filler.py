"""将扫描快照中的评分标准填入开评标工作簿，其他 ZIP 部件原样保留。

首版支持「序号/评审因素/评分标准说明/分值/投标人…」模板。
结构不明确或已有评分时拒绝写入，不能靠猜测覆盖评审结果。
"""

import io
import math
import os
import posixpath
import re
import unicodedata
from copy import deepcopy
from dataclasses import dataclass
from decimal import Decimal
from zipfile import ZipFile

from lxml import etree as ET
from openpyxl.utils.cell import get_column_letter, range_boundaries

from file_io import atomic_output_path

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PRICE_LABELS = {"投标报价", "投标价格", "报价", "价格", "报价得分", "价格得分", "价格分", "价格部分"}
HEADERS = (
    {"评审因素", "评分因素", "评分项目", "评审项目", "名称", "项目名称", "审查项目"},
    {"评分说明", "评分标准", "评分标准说明", "评分细则"},
    {"分值", "满分", "分数", "标准分"},
)


def _norm(value):
    return re.sub(r"[\s、，,：:]", "", unicodedata.normalize("NFKC", str(value or "")))


@dataclass(frozen=True)
class ReviewCriterion:
    factor: str
    description: str
    points: Decimal
    source_row: int
    factor_group: int


@dataclass(frozen=True)
class ReviewRules:
    criteria: tuple
    excluded_prices: tuple

    @property
    def total(self):
        return sum((c.points for c in self.criteria), Decimal(0))


def extract_review_rules(table, mapping=None, include_price=False):
    """按表头定位列；继承合并因素，但不猜测/重复计算合并分值。"""
    if table is None:
        raise ValueError("缺少表格扫描快照，请重新扫描。")
    grid = {}
    for cell in table.cells:
        for row in range(cell.row, cell.row + cell.row_span):
            for col in range(cell.column, cell.column + cell.column_span):
                grid[row, col] = cell
    headers = []
    for row in range(1, min(table.row_count, 8) + 1):
        columns = [[col for col in range(1, table.column_count + 1)
                    if (row, col) in grid and _norm(grid[row, col].text) in aliases]
                   for aliases in HEADERS]
        if all(len(cols) == 1 for cols in columns):
            headers.append((row, *(cols[0] for cols in columns)))
    if mapping is not None:
        header = int(mapping["header"])
        cols = [int(mapping[key]) for key in ("factor", "description", "points")]
        if not 1 <= header < table.row_count or len(set(cols)) != 3 or any(not 1 <= c <= table.column_count for c in cols):
            raise ValueError("来源表头行或对应列无效。")
        headers = [(header, *cols)]
    if len(headers) != 1:
        raise ValueError("无法唯一识别评分表的“评审因素、评分说明、分值”表头。")
    header, factor_col, description_col, points_col = headers[0]
    criteria, prices, totals = [], [], []
    processed_scores = set()
    for row in range(header + 1, table.row_count + 1):
        cells = [grid.get((row, col)) for col in (factor_col, description_col, points_col)]
        factor, description, points = [c.text if c else "" for c in cells]
        if not any(s.strip() for s in (factor, description, points)):
            continue
        if _norm(factor) in HEADERS[0] and _norm(points) in HEADERS[2]:
            continue
        if _norm(factor) in {"合计", "总计", "总分"}:
            if re.fullmatch(r"\d+(?:\.\d+)?(?:分)?", points.strip()):
                totals.append(Decimal(points.strip().removesuffix("分")))
            continue
        score_cell = cells[2]
        score_key = (score_cell.row, score_cell.column) if score_cell else None
        if score_key in processed_scores:
            continue
        if not re.fullmatch(r"\d+(?:\.\d+)?(?:分)?", points.strip()):
            raise ValueError(f"第 {row} 行分值不是明确数值：{points!r}。")
        bottom = score_cell.row + score_cell.row_span
        if score_cell.row <= header or bottom > table.row_count + 1:
            raise ValueError(f"第 {row} 行分值跨行合并范围超出评分区域，请核对来源表格。")
        names, descriptions, segments = [], [], []
        seen_descriptions = set()
        for member in range(score_cell.row, bottom):
            name_cell = grid.get((member, factor_col))
            desc_cell = grid.get((member, description_col))
            name = name_cell.text if name_cell else ""
            if _norm(name) in {"合计", "总计", "总分"} or _norm(name) in HEADERS[0]:
                raise ValueError(f"第 {row} 行合并分值跨越表头或合计，请核对来源表格。")
            if name.strip() and name not in names:
                names.append(name)
            if desc_cell is not None:
                key = (desc_cell.row, desc_cell.column)
                if key not in seen_descriptions and desc_cell.text.strip():
                    descriptions.append(desc_cell.text)
                    segments.append((name, desc_cell.text))
                    seen_descriptions.add(key)
        if not names or not descriptions:
            raise ValueError(f"第 {row} 行评分组缺少项目名称或评分说明，请核对源表。")
        is_price = [_norm(name) in PRICE_LABELS for name in names]
        if any(is_price) and not all(is_price):
            raise ValueError(f"第 {row} 行合并分值同时包含价格和非价格项目，请核对评分组边界。")
        # 同一说明单元格只读一次；不同单元格即使文字相同也保留，避免删掉原文。
        grouped_description = "\n\n".join(
            (name + "\n" if len(names) > 1 and name.strip() else "") + text
            for name, text in segments)
        criterion = ReviewCriterion(" / ".join(names), grouped_description,
                                    Decimal(points.strip().removesuffix("分")), row,
                                    cells[0].row if cells[0] else row)
        processed_scores.add(score_key)
        (prices if all(is_price) and not include_price else criteria).append(criterion)
    if not criteria:
        raise ValueError("没有可填入的商务技术评分项。")
    full_total = sum((c.points for c in criteria + prices), Decimal(0))
    if any(total != full_total for total in totals):
        raise ValueError("源表合计与各项分值之和不一致，请先核对。")
    return ReviewRules(tuple(criteria), tuple(prices))


def _tag(name):
    return f"{{{NS}}}{name}"


def _xml(data):
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))


def _serialize(root):
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


class _Package:
    def __init__(self, path):
        with open(path, "rb") as file:
            self.raw = file.read()
        self.zip = ZipFile(io.BytesIO(self.raw))
        self.book = _xml(self.zip.read("xl/workbook.xml"))
        rels = _xml(self.zip.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): posixpath.normpath(posixpath.join("xl", r.get("Target")))
                   if not r.get("Target").startswith("/") else r.get("Target").lstrip("/")
                   for r in rels if r.get("TargetMode") != "External"}
        self.sheets = [(s.get("name"), targets[s.get(f"{{{REL}}}id")])
                       for s in self.book.find(_tag("sheets"))]
        self.strings = []
        if "xl/sharedStrings.xml" in self.zip.namelist():
            self.strings = ["".join(si.itertext()) for si in
                            _xml(self.zip.read("xl/sharedStrings.xml"))]

    def value(self, cell):
        if cell is None:
            return ""
        if cell.find(_tag("f")) is not None:
            return "=" + (cell.findtext(_tag("f")) or "")
        if cell.get("t") == "inlineStr":
            inline = cell.find(_tag("is"))
            return "".join(inline.itertext()) if inline is not None else ""
        value = cell.findtext(_tag("v"), "")
        return self.strings[int(value)] if cell.get("t") == "s" else value


@dataclass(frozen=True)
class ReviewTarget:
    sheet_name: str
    header_row: int
    start_row: int
    total_row: int
    last_column: int
    blank: bool


def _target(package, name, root):
    rows = root.find(_tag("sheetData"))
    cells = {c.get("r"): c for row in rows for c in row}
    headers = [int(r.get("r")) for r in rows
               if [_norm(package.value(cells.get(f"{col}{r.get('r')}"))) for col in "ABCD"]
               in [["序号", factor, desc, points] for factor in HEADERS[0]
                   for desc in HEADERS[1] for points in HEADERS[2]]]
    if len(headers) != 1:
        raise ValueError("模板需包含唯一的 A:D 列“序号、评审因素、评分标准说明、分值”表头。")
    header = headers[0]
    start = header + 1
    merges = root.find(_tag("mergeCells"))
    for merge in merges if merges is not None else []:
        left, top, right, bottom = range_boundaries(merge.get("ref"))
        if top <= header <= bottom and left <= 4:
            start = max(start, bottom + 1)
    totals = [int(r.get("r")) for r in rows if int(r.get("r")) >= start
              and _norm(package.value(cells.get(f"A{r.get('r')}"))) in {"合计", "总计"}]
    if len(totals) != 1 or totals[0] <= start:
        raise ValueError("模板需在评分区后保留一行“合计”。")
    last_col = max((range_boundaries(cell.get("r"))[0] for row in rows
                    if int(row.get("r")) == header for cell in row), default=4)
    if last_col < 5:
        raise ValueError("模板缺少投标人评分列。")
    blank = not any(package.value(cells.get(f"{col}{row}")) for row in range(start, totals[0]) for col in "BC")
    return ReviewTarget(name, header, start, totals[0], last_col, blank)


def inspect_review_targets(path):
    package = _Package(path)
    targets = []
    for name, part in package.sheets:
        if "评审" not in name and "评分" not in name:
            continue
        try:
            targets.append(_target(package, name, _xml(package.zip.read(part))))
        except ValueError:
            continue
    return sorted(targets, key=lambda t: (not t.blank, "商务" not in t.sheet_name, t.sheet_name))


def review_output_path(template_path, output_dir=None):
    stem, extension = os.path.splitext(os.path.basename(template_path))
    directory = output_dir or os.path.dirname(template_path)
    base = os.path.join(directory, stem + "_商务技术评审")
    path, suffix = base + extension, 1
    while os.path.exists(path):
        path = f"{base}_{suffix}{extension}"
        suffix += 1
    return path


def _write(cell, value, formula=None):
    for child in list(cell):
        if child.tag in {_tag("f"), _tag("v"), _tag("is")}:
            cell.remove(child)
    cell.attrib.pop("t", None)
    if formula:
        ET.SubElement(cell, _tag("f")).text = formula
        ET.SubElement(cell, _tag("v")).text = str(value)
    elif isinstance(value, (Decimal, int)):
        ET.SubElement(cell, _tag("v")).text = str(value)
    elif value is not None:
        cell.set("t", "inlineStr")
        node = ET.SubElement(ET.SubElement(cell, _tag("is")), _tag("t"))
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        node.text = str(value)


def _move_ref(ref, boundary, delta):
    return re.sub(r"(\$?[A-Z]{1,3}\$?)(\d+)",
                  lambda m: m[1] + str(int(m[2]) + (delta if int(m[2]) >= boundary else 0)), ref)


def _height(package, root, row, criterion):
    """按模板列宽/字体估计换行高度，过长单元格显式报错而非截断原文。"""
    styles = _xml(package.zip.read("xl/styles.xml"))
    xfs, fonts = styles.find(_tag("cellXfs")), styles.find(_tag("fonts"))
    height = 30.0
    for col, text in ((2, criterion.factor), (3, criterion.description)):
        cell = next(c for c in row if c.get("r").startswith(get_column_letter(col)))
        xf = xfs[int(cell.get("s", "0"))]
        font = fonts[int(xf.get("fontId", "0"))]
        size = font.find(_tag("sz"))
        points = float(size.get("val")) if size is not None else 11
        width = 9.0
        for dim in root.findall(f"{_tag('cols')}/{_tag('col')}"):
            if int(dim.get("min")) <= col <= int(dim.get("max")):
                width = float(dim.get("width", 9))
        available = max(12, width * 7 - 10)
        lines = sum(max(1, math.ceil(sum((1 if unicodedata.east_asian_width(ch) in "WF" else .55)
                                        * points * 4 / 3 for ch in line) / available))
                    for line in text.splitlines() or [""])
        height = max(height, lines * points * 1.35 + 10)
    if height > 409:
        raise ValueError(f"“{criterion.factor.strip()}”内容过长，请先加宽模板评分说明列。")
    return round(height, 2)


def fill_review_workbook(template_path, output_path, sheet_name, rules):
    """只改目标 sheet XML 与该表打印范围；不经过整本工作簿重新序列化。"""
    if os.path.normcase(os.path.abspath(template_path)) == os.path.normcase(os.path.abspath(output_path)):
        raise ValueError("请另存结果，不能覆盖模板。")
    if os.path.exists(output_path):
        raise ValueError("输出文件已存在，请选择新的文件名。")
    if not rules.criteria:
        raise ValueError("没有可填入的评分项。")
    package = _Package(template_path)
    paths = dict(package.sheets)
    if sheet_name not in paths:
        raise ValueError("目标工作表不存在。")
    part = paths[sheet_name]
    root = _xml(package.zip.read(part))
    target = _target(package, sheet_name, root)
    start, end = target.start_row, target.total_row
    delta = len(rules.criteria) - (end - start)
    rows = root.find(_tag("sheetData"))
    by_row = {int(row.get("r")): row for row in rows}
    if root.find(_tag("sheetProtection")) is not None:
        raise ValueError("目标工作表受保护，请先解除保护。")
    # 这些对象拥有独立坐标/公式，不能只移动单元格后声称已完整保留。
    unsupported = {"drawing", "legacyDrawing", "legacyDrawingHF", "tableParts", "oleObjects",
                   "controls", "dataValidations", "conditionalFormatting", "hyperlinks", "extLst"}
    if any(ET.QName(child).localname in unsupported for child in root):
        raise ValueError("目标表含图形、批注、校验或其他带坐标对象，暂不支持自动填入该表。")
    if "xl/calcChain.xml" in package.zip.namelist():
        raise ValueError("工作簿包含计算链，请用 Excel/WPS 另存后重试。")
    for rownum, row in by_row.items():
        for c in row:
            col = range_boundaries(c.get("r"))[0]
            value = package.value(c)
            if start <= rownum < end and col >= 5 and value:
                raise ValueError(f"{sheet_name}!{c.get('r')} 已有评分或公式，请使用未评分模板。")
            if c.find(_tag("f")) is not None and rownum >= start:
                expected = f"=SUM(D{start}:D{end - 1})"
                if c.get("r") != f"D{end}" or value.upper().replace("$", "") != expected:
                    raise ValueError(f"目标区域包含非标准公式 {c.get('r')}，暂不支持自动调整。")
    # 调整行数或分值后，其他表对目标表的引用缓存可能失效；不静默留下旧值。
    for name, other_part in package.sheets:
        if other_part == part:
            continue
        other = _xml(package.zip.read(other_part))
        if any(sheet_name in (f.text or "") for f in other.iter(_tag("f"))):
            raise ValueError(f"“{name}”引用了目标表，请先使用独立的评分模板。")
    names = package.book.find(_tag("definedNames"))
    index = [n for n, _ in package.sheets].index(sheet_name)
    if names is not None:
        for name in names:
            if sheet_name in (name.text or "") and name.get("name") not in {"_xlnm.Print_Area", "_xlnm.Print_Titles"}:
                raise ValueError("存在引用目标表的自定义名称，暂不支持自动调整。")
    proto = by_row.get(start)
    if proto is None or any(not any(c.get("r") == f"{col}{start}" for c in proto) for col in "ABCD"):
        raise ValueError("首行评分模板缺少 A:D 单元格或样式。")
    if not any(c.get("r") == f"D{end}" for c in by_row[end]):
        raise ValueError("合计行缺少 D 列分值单元格。")
    row_templates = [deepcopy(by_row.get(rownum, proto)) for rownum in range(start, end)]
    # 移动尾部（合计、签字区等），评分区按模板第一行扩展。
    for rownum, row in by_row.items():
        if start <= rownum < end:
            rows.remove(row)
        elif rownum >= end:
            row.set("r", str(rownum + delta))
            for c in row:
                c.set("r", _move_ref(c.get("r"), end, delta))
    insertion = list(rows).index(by_row[end])
    for offset, criterion in enumerate(rules.criteria):
        rownum = start + offset
        row = deepcopy(row_templates[min(offset, len(row_templates) - 1)])
        # 合并因素的续行有时不存储 B 单元格，沿用首行样式补齐。
        for col in "ABCD":
            if not any(re.match(r"[A-Z]+", c.get("r"))[0] == col for c in row):
                row.append(deepcopy(next(c for c in proto if c.get("r") == f"{col}{start}")))
        row[:] = sorted(row, key=lambda c: range_boundaries(c.get("r"))[0])
        row.set("r", str(rownum))
        row.set("ht", str(_height(package, root, row, criterion)))
        row.set("customHeight", "1")
        values = (offset + 1, criterion.factor, criterion.description, criterion.points)
        for cell in row:
            col = range_boundaries(cell.get("r"))[0]
            cell.set("r", f"{get_column_letter(col)}{rownum}")
            _write(cell, values[col - 1] if col <= 4 else None)
        rows.insert(insertion + offset, row)
    total_cell = next(c for c in by_row[end] if c.get("r") == f"D{end + delta}")
    _write(total_cell, rules.total, f"SUM(D{start}:D{end + delta - 1})")
    merges = root.find(_tag("mergeCells"))
    if merges is None:
        merges = ET.Element(_tag("mergeCells"))
        root.insert(list(root).index(rows) + 1, merges)
    for merge in list(merges):
        left, top, right, bottom = range_boundaries(merge.get("ref"))
        if top < end and bottom >= start:
            if top < start or bottom >= end or left != 2 or right != 2:
                raise ValueError("评分区域存在非因素列合并或跨区合并，请先调整模板。")
            merges.remove(merge)
        else:
            merge.set("ref", _move_ref(merge.get("ref"), end, delta))
    # 同一源合并单元格才合并；相同文字不自动视为同一因素。
    offset = 0
    while offset < len(rules.criteria):
        tail = offset + 1
        while tail < len(rules.criteria) and rules.criteria[tail].factor_group == rules.criteria[offset].factor_group:
            tail += 1
        if tail - offset > 1:
            ET.SubElement(merges, _tag("mergeCell"), ref=f"B{start + offset}:B{start + tail - 1}")
            for row in rows:
                if start + offset < int(row.get("r")) < start + tail:
                    _write(next(c for c in row if c.get("r").startswith("B")), None)
        offset = tail
    merges.set("count", str(len(merges)))
    for node in root:
        if ET.QName(node).localname in {"dimension", "autoFilter"} and node.get("ref"):
            node.set("ref", _move_ref(node.get("ref"), end, delta))
        if ET.QName(node).localname == "rowBreaks":
            for brk in node:
                if int(brk.get("id")) >= start:
                    node.remove(brk)
            node.set("count", str(len(node)))
            node.set("manualBreakCount", str(len(node)))
    changed = {part: _serialize(root)}
    if names is not None and delta:
        for name in names:
            if name.get("localSheetId") == str(index) and name.get("name") == "_xlnm.Print_Area":
                prefix, coords = name.text.rsplit("!", 1)
                name.text = prefix + "!" + _move_ref(coords, end, delta)
                changed["xl/workbook.xml"] = _serialize(package.book)
    with atomic_output_path(output_path) as temporary:
        with ZipFile(temporary, "w") as output:
            output.comment = package.zip.comment
            for info in package.zip.infolist():
                output.writestr(info, changed.get(info.filename, package.zip.read(info.filename)))
    return {"output_path": output_path, "sheet_name": sheet_name,
            "count": len(rules.criteria), "total": str(rules.total),
            "excluded": [c.factor for c in rules.excluded_prices]}
