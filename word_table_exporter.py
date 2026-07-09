"""
Word table to Excel exporter.

Reads .docx files directly with python-docx and writes .xlsx workbooks with
openpyxl. It does not require Word, WPS, or Excel to be installed.
"""

import os
import re
import unicodedata
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter


TABLE_OUTPUT_SUFFIX = "_表格"
MAX_COLUMN_WIDTH = 80
MIN_COLUMN_WIDTH = 8
MAX_HINT_TEXT_LENGTH = 80
MAX_PREVIEW_LENGTH = 120


@dataclass
class ExportedTableCell:
    row: int
    column: int
    row_span: int
    column_span: int
    text: str


@dataclass
class ExportedTable:
    cells: List[ExportedTableCell]
    row_count: int
    column_count: int


@dataclass
class TableScanItem:
    file_path: str
    filename: str
    table_index: int
    section: str
    context: str
    preview: str
    row_count: int
    column_count: int
    hint: str


def get_table_export_output_path(docx_path: str, output_dir: Optional[str] = None) -> str:
    """Return a non-overwriting xlsx path for a Word table export."""
    directory = output_dir or os.path.dirname(docx_path)
    stem = os.path.splitext(os.path.basename(docx_path))[0]
    base_name = f"{stem}{TABLE_OUTPUT_SUFFIX}.xlsx"
    candidate = os.path.join(directory, base_name)
    index = 1

    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{stem}{TABLE_OUTPUT_SUFFIX}_{index}.xlsx")
        index += 1

    return candidate


def scan_word_tables(docx_path: str) -> List[TableScanItem]:
    """Scan body tables and return context-rich metadata for user selection."""
    document = Document(docx_path)
    filename = os.path.basename(docx_path)
    table_items: List[TableScanItem] = []
    headings: Dict[int, str] = {}
    recent_paragraphs: List[str] = []
    table_index = 0

    for block in _iter_body_blocks(document):
        if isinstance(block, Paragraph):
            text = _normalize_space(block.text)
            if not text:
                continue

            heading_level = _detect_heading_level(block, text)
            if heading_level:
                headings = {
                    level: heading
                    for level, heading in headings.items()
                    if level < heading_level
                }
                headings[heading_level] = _shorten(text, MAX_HINT_TEXT_LENGTH)

            recent_paragraphs.append(text)
            recent_paragraphs = recent_paragraphs[-4:]
            continue

        table_index += 1
        exported_table = _extract_table(block)
        section = _section_text(headings)
        context = _context_text(recent_paragraphs)
        preview = _table_preview(exported_table)
        hint = _table_hint(section, context, preview)
        table_items.append(TableScanItem(
            file_path=docx_path,
            filename=filename,
            table_index=table_index,
            section=section,
            context=context,
            preview=preview,
            row_count=exported_table.row_count,
            column_count=exported_table.column_count,
            hint=hint,
        ))

    return table_items


def batch_scan_word_tables(
    file_paths: List[str],
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> Tuple[List[TableScanItem], Dict[str, str], Optional[str]]:
    items: List[TableScanItem] = []
    skipped: Dict[str, str] = {}
    errors: List[str] = []

    for index, file_path in enumerate(file_paths):
        filename = os.path.basename(file_path)
        if progress_callback:
            progress_callback(index + 1, len(file_paths), filename)

        if os.path.splitext(file_path)[1].lower() != ".docx":
            skipped[filename] = "不支持的文件格式"
            continue

        try:
            file_items = scan_word_tables(file_path)
            if not file_items:
                skipped[filename] = "未找到表格"
                continue
            items.extend(file_items)
        except Exception as exc:
            errors.append(f"{filename}: {_format_export_error(exc)}")

    return items, skipped, "\n".join(errors) if errors else None


def export_word_tables_to_excel(
    docx_path: str,
    output_path: str,
    table_indexes: Optional[List[int]] = None,
) -> int:
    """Export all body tables from one .docx file to one .xlsx workbook.

    Returns the number of exported tables. If the document has no body tables,
    no workbook is written and 0 is returned.
    """
    document = Document(docx_path)
    tables = _body_tables(document)
    if table_indexes is not None:
        selected = set(table_indexes)
        tables = [
            table
            for index, table in enumerate(tables, start=1)
            if index in selected
        ]

    if not tables:
        return 0

    workbook = Workbook()
    try:
        for sheet_index, table in enumerate(tables, start=1):
            worksheet = workbook.active if sheet_index == 1 else workbook.create_sheet()
            worksheet.title = f"表格{sheet_index}"
            _write_table_to_worksheet(worksheet, _extract_table(table))

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        workbook.save(output_path)
        return len(tables)
    finally:
        workbook.close()


def batch_export_word_tables(
    file_paths: List[str],
    output_dir: Optional[str] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    selected_tables: Optional[Dict[str, List[int]]] = None,
) -> Tuple[Dict[str, Dict[str, object]], Dict[str, str], Optional[str]]:
    """Batch export Word tables.

    Returns (results, skipped, error_text). results maps original filenames to
    {"tables": int, "output_path": str}; skipped maps filenames to a reason.
    """
    results: Dict[str, Dict[str, object]] = {}
    skipped: Dict[str, str] = {}
    errors: List[str] = []

    for index, file_path in enumerate(file_paths):
        filename = os.path.basename(file_path)
        if progress_callback:
            progress_callback(index + 1, len(file_paths), filename)

        if os.path.splitext(file_path)[1].lower() != ".docx":
            skipped[filename] = "不支持的文件格式"
            continue

        try:
            table_indexes = None
            if selected_tables is not None:
                table_indexes = selected_tables.get(_file_identity(file_path), [])
                if not table_indexes:
                    skipped[filename] = "未选择表格"
                    continue

            output_path = get_table_export_output_path(file_path, output_dir)
            table_count = export_word_tables_to_excel(file_path, output_path, table_indexes=table_indexes)
            if table_count == 0:
                skipped[filename] = "未找到表格"
                continue

            results[filename] = {
                "tables": table_count,
                "output_path": output_path,
            }
        except Exception as exc:
            errors.append(f"{filename}: {_format_export_error(exc)}")

    return results, skipped, "\n".join(errors) if errors else None


def _file_identity(path: str) -> str:
    return os.path.normcase(os.path.abspath(os.path.realpath(os.fspath(path))))


def _iter_body_blocks(document):
    for child in document.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def _body_tables(document) -> List[Table]:
    return [block for block in _iter_body_blocks(document) if isinstance(block, Table)]


def _detect_heading_level(paragraph: Paragraph, text: str) -> Optional[int]:
    style_name = ""
    try:
        style_name = paragraph.style.name or ""
    except Exception:
        pass

    style_match = re.search(r"(?:Heading|标题)\s*([1-9])", style_name, flags=re.IGNORECASE)
    if style_match:
        return int(style_match.group(1))

    if re.match(r"^第[一二三四五六七八九十百千万\d]+[章节篇部分]\s*", text):
        return 1
    if re.match(r"^[一二三四五六七八九十]+[、.．]\s*\S+", text):
        return 2
    if re.match(r"^[（(][一二三四五六七八九十\d]+[）)]\s*\S+", text):
        return 3

    number_match = re.match(r"^(\d+(?:[.．]\d+){0,3})[、.．\s]+\S+", text)
    if number_match:
        return min(4, 2 + number_match.group(1).count(".") + number_match.group(1).count("．"))

    return None


def _section_text(headings: Dict[int, str]) -> str:
    if not headings:
        return "未识别章节"
    return " / ".join(headings[level] for level in sorted(headings))


def _context_text(paragraphs: List[str]) -> str:
    if not paragraphs:
        return "表格前无明显文字"
    return _shorten(" / ".join(paragraphs[-3:]), MAX_PREVIEW_LENGTH)


def _table_preview(table: ExportedTable) -> str:
    parts = []
    for cell in sorted(table.cells, key=lambda item: (item.row, item.column)):
        if cell.row > 3 and parts:
            break
        text = _normalize_space(cell.text)
        if text:
            parts.append(text)
        if len(parts) >= 8:
            break

    if not parts:
        return "空表格或无可读文本"
    return _shorten(" | ".join(parts), MAX_PREVIEW_LENGTH)


def _table_hint(section: str, context: str, preview: str) -> str:
    source = f"{section} {context} {preview}"
    positive_groups = [
        ("资格审查", ("资格审查", "资格性审查", "资格条件", "资格要求")),
        ("符合性审查", ("符合性审查", "符合审查", "实质性响应")),
        ("评分/评审表", ("评分", "评审因素", "评审标准", "分值", "商务技术", "技术评审", "综合评分")),
        ("采购需求/技术参数", ("采购需求", "技术要求", "技术参数", "服务要求", "货物需求", "参数要求")),
    ]
    negative_groups = [
        ("投标文件格式", ("投标函", "授权委托", "法定代表人", "承诺函", "声明函", "格式")),
        ("合同/报价类", ("合同条款", "报价", "开标一览表", "分项报价", "价格表")),
    ]

    for label, keywords in positive_groups:
        matched = [keyword for keyword in keywords if keyword in source]
        if matched:
            return f"建议关注：{label}（命中：{'、'.join(matched[:3])}）"

    for label, keywords in negative_groups:
        matched = [keyword for keyword in keywords if keyword in source]
        if matched:
            return f"谨慎选择：可能是{label}（命中：{'、'.join(matched[:3])}）"

    return "未命中常见评审关键词，请看章节和预览确认"


def _normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _shorten(text: str, max_length: int) -> str:
    text = _normalize_space(text)
    if len(text) <= max_length:
        return text
    return text[: max_length - 1] + "…"


def _extract_table(table) -> ExportedTable:
    cells: List[ExportedTableCell] = []
    active_vertical: Dict[int, ExportedTableCell] = {}
    max_column = 0
    rows = list(table._tbl.tr_lst)

    for row_index, tr in enumerate(rows, start=1):
        column_index = _row_grid_before(tr) + 1

        for tc in tr.tc_lst:
            column_span = _grid_span(tc)
            v_merge = _v_merge(tc)
            max_column = max(max_column, column_index + column_span - 1)

            if v_merge == "continue":
                start_cell = _find_active_vertical_cell(active_vertical, column_index, column_span)
                if start_cell is not None:
                    start_cell.row_span = max(start_cell.row_span, row_index - start_cell.row + 1)
                    for offset in range(column_span):
                        active_vertical[column_index + offset] = start_cell
                else:
                    cells.append(ExportedTableCell(
                        row=row_index,
                        column=column_index,
                        row_span=1,
                        column_span=column_span,
                        text=_cell_text(tc, table),
                    ))
            else:
                for offset in range(column_span):
                    active_vertical.pop(column_index + offset, None)

                cell = ExportedTableCell(
                    row=row_index,
                    column=column_index,
                    row_span=1,
                    column_span=column_span,
                    text=_cell_text(tc, table),
                )
                cells.append(cell)

                if v_merge == "restart":
                    for offset in range(column_span):
                        active_vertical[column_index + offset] = cell

            column_index += column_span

    return ExportedTable(cells=cells, row_count=len(rows), column_count=max_column)


def _row_grid_before(tr) -> int:
    tr_pr = getattr(tr, "trPr", None)
    if tr_pr is None:
        return 0
    grid_before = tr_pr.find(qn("w:gridBefore"))
    if grid_before is None:
        return 0
    try:
        return int(grid_before.get(qn("w:val")) or 0)
    except ValueError:
        return 0


def _grid_span(tc) -> int:
    tc_pr = getattr(tc, "tcPr", None)
    if tc_pr is None:
        return 1
    grid = tc_pr.find(qn("w:gridSpan"))
    if grid is None:
        return 1
    try:
        return max(1, int(grid.get(qn("w:val")) or 1))
    except ValueError:
        return 1


def _v_merge(tc) -> Optional[str]:
    tc_pr = getattr(tc, "tcPr", None)
    if tc_pr is None:
        return None
    merge = tc_pr.find(qn("w:vMerge"))
    if merge is None:
        return None
    value = merge.get(qn("w:val"))
    return value or "continue"


def _find_active_vertical_cell(
    active_vertical: Dict[int, ExportedTableCell],
    column_index: int,
    column_span: int,
) -> Optional[ExportedTableCell]:
    for offset in range(column_span):
        cell = active_vertical.get(column_index + offset)
        if cell is not None:
            return cell
    return None


def _cell_text(tc, table) -> str:
    try:
        text = _Cell(tc, table).text
    except Exception:
        text_parts = []
        for text_node in tc.iter(qn("w:t")):
            if text_node.text:
                text_parts.append(text_node.text)
        text = "".join(text_parts)
    return text.replace("\r\a", "").replace("\a", "").replace("\x07", "")


def _write_table_to_worksheet(worksheet, table: ExportedTable) -> None:
    border = _thin_border()
    alignment = Alignment(wrap_text=True, vertical="top")

    for row in range(1, max(table.row_count, 1) + 1):
        for column in range(1, max(table.column_count, 1) + 1):
            cell = worksheet.cell(row=row, column=column)
            cell.number_format = "@"
            cell.alignment = alignment
            cell.border = border

    for exported_cell in table.cells:
        cell = worksheet.cell(row=exported_cell.row, column=exported_cell.column)
        cell.value = exported_cell.text

    for exported_cell in table.cells:
        end_row = exported_cell.row + exported_cell.row_span - 1
        end_column = exported_cell.column + exported_cell.column_span - 1
        if end_row == exported_cell.row and end_column == exported_cell.column:
            continue
        try:
            worksheet.merge_cells(
                start_row=exported_cell.row,
                start_column=exported_cell.column,
                end_row=end_row,
                end_column=end_column,
            )
        except ValueError:
            # Extremely irregular Word tables can produce overlapping ranges;
            # keep the cell text export rather than failing the whole document.
            pass

    _fit_columns(worksheet, table)


def _thin_border() -> Border:
    side = Side(style="thin", color="808080")
    return Border(left=side, right=side, top=side, bottom=side)


def _fit_columns(worksheet, table: ExportedTable) -> None:
    widths = {column: MIN_COLUMN_WIDTH for column in range(1, max(table.column_count, 1) + 1)}

    for exported_cell in table.cells:
        if not exported_cell.text:
            continue
        widest_line = max(_display_width(line) for line in exported_cell.text.splitlines() or [""])
        width = min(MAX_COLUMN_WIDTH, max(MIN_COLUMN_WIDTH, widest_line + 2))
        per_column_width = min(MAX_COLUMN_WIDTH, max(MIN_COLUMN_WIDTH, width // exported_cell.column_span + 2))
        for offset in range(exported_cell.column_span):
            column = exported_cell.column + offset
            widths[column] = max(widths.get(column, MIN_COLUMN_WIDTH), per_column_width)

    for column, width in widths.items():
        worksheet.column_dimensions[get_column_letter(column)].width = width


def _display_width(text: str) -> int:
    width = 0
    for char in text:
        width += 2 if unicodedata.east_asian_width(char) in {"F", "W"} else 1
    return width


def _format_export_error(exc: Exception) -> str:
    if isinstance(exc, PermissionError):
        return "文件无法读写，请先关闭正在打开的 Word/Excel 文件后重试"
    return str(exc)
