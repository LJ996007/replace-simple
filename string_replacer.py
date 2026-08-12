"""
简化版批量文本替换模块。
支持 Word (.docx)、Excel (.xlsx/.xlsm)、PowerPoint (.pptx) 文件。
"""

import os
import shutil
import warnings
from typing import Callable, Dict, List, Optional, Tuple

warnings.filterwarnings("ignore", message="Data Validation extension is not supported")


_WINDOWS_INVALID_FILENAME_REPLACEMENTS = str.maketrans(
    {
        "<": "＜",
        ">": "＞",
        ":": "：",
        '"': "＂",
        "/": "／",
        "\\": "＼",
        "|": "｜",
        "?": "？",
        "*": "＊",
    }
)
_WINDOWS_RESERVED_FILENAME_STEMS = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{index}" for index in range(1, 10)),
    *(f"LPT{index}" for index in range(1, 10)),
}

def _clean_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()
    return str(value).strip()


def load_replacement_rules(excel_path: str) -> List[Tuple[str, str]]:
    """从 Excel 第一张工作表读取两列替换规则。"""
    if os.path.splitext(excel_path)[1].lower() == ".xls":
        return _load_xls_replacement_rules(excel_path)

    from openpyxl import load_workbook

    rules = []
    seen = set()
    wb = load_workbook(excel_path, data_only=False)
    ws = wb.active

    try:
        for old_cell, new_cell in ws.iter_rows(min_row=1, max_col=2):
            old_text = _clean_text(old_cell.value)
            new_text = _clean_text(new_cell.value)
            if not old_text:
                continue

            pair = (old_text, new_text)
            if pair in seen:
                continue

            seen.add(pair)
            rules.append(pair)
    finally:
        wb.close()

    return rules


def _load_xls_replacement_rules(excel_path: str) -> List[Tuple[str, str]]:
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError("读取 .xls 规则表需要安装 xlrd：pip install xlrd") from exc

    rules = []
    seen = set()
    wb = xlrd.open_workbook(excel_path)
    ws = wb.sheet_by_index(0)

    for row_idx in range(ws.nrows):
        old_value = ws.cell_value(row_idx, 0) if ws.ncols >= 1 else ""
        new_value = ws.cell_value(row_idx, 1) if ws.ncols >= 2 else ""
        old_text = _clean_text(old_value)
        new_text = _clean_text(new_value)
        if not old_text:
            continue

        pair = (old_text, new_text)
        if pair in seen:
            continue

        seen.add(pair)
        rules.append(pair)

    return rules


def _prepare_rules(rules: List[Tuple[str, str]]) -> List[Dict[str, str]]:
    prepared = []
    for order, (old_text, new_text) in enumerate(rules):
        if not old_text:
            continue
        prepared.append({
            "old": old_text,
            "new": new_text,
            "first_char": old_text[0],
            "order": order,
        })
    return sorted(prepared, key=lambda rule: (-len(rule["old"]), rule["order"]))


def _replace_text_with_rules(text: str, prepared_rules: List[Dict[str, str]]) -> Tuple[str, int]:
    if not text:
        return text, 0

    replaced = text
    total_count = 0
    present_chars = set(text)

    for rule in prepared_rules:
        if rule["first_char"] not in present_chars:
            continue

        old_text = rule["old"]
        if old_text not in replaced:
            continue

        replaced, replaced_count = _replace_text_with_rule(replaced, old_text, rule["new"])
        total_count += replaced_count
        present_chars = set(replaced)

    return replaced, total_count


def _build_run_spans(runs) -> List[Dict[str, object]]:
    spans = []
    position = 0

    for index, run in enumerate(runs):
        text = run.text or ""
        end = position + len(text)
        spans.append({
            "index": index,
            "run": run,
            "start": position,
            "end": end,
        })
        position = end

    return spans


def _find_span_for_position(spans: List[Dict[str, object]], position: int) -> Optional[Dict[str, object]]:
    for span in spans:
        if span["start"] <= position < span["end"]:
            return span
    return None


def _find_occurrences(text: str, needle: str) -> List[int]:
    positions = []
    start = 0

    while True:
        index = text.find(needle, start)
        if index == -1:
            return positions
        positions.append(index)
        start = index + len(needle)


def _is_occurrence_inside_replacement(text: str, start: int, old_text: str, new_text: str) -> bool:
    """判断当前位置的 old_text 是否已经属于完整 new_text，避免重复追加。

    例如规则 A -> A采购 时，第二次处理 A采购 不应再把前缀 A 替换成 A采购。
    """
    if not old_text or old_text not in new_text:
        return False

    offset = new_text.find(old_text)
    while offset != -1:
        replacement_start = start - offset
        replacement_end = replacement_start + len(new_text)
        if (
            replacement_start >= 0
            and replacement_end <= len(text)
            and text[replacement_start:replacement_end] == new_text
        ):
            return True
        offset = new_text.find(old_text, offset + 1)

    return False


def _replace_text_with_rule(text: str, old_text: str, new_text: str) -> Tuple[str, int]:
    if not text or not old_text or old_text not in text:
        return text, 0

    parts = []
    count = 0
    cursor = 0
    search_start = 0

    while True:
        index = text.find(old_text, search_start)
        if index == -1:
            parts.append(text[cursor:])
            break

        end = index + len(old_text)
        if _is_occurrence_inside_replacement(text, index, old_text, new_text):
            parts.append(text[cursor:end])
        else:
            parts.append(text[cursor:index])
            parts.append(new_text)
            count += 1

        cursor = end
        search_start = end

    return "".join(parts), count


def _replace_range_in_runs(runs, spans: List[Dict[str, object]], start: int, end: int, new_text: str) -> None:
    start_span = _find_span_for_position(spans, start)
    end_span = _find_span_for_position(spans, end - 1)
    if start_span is None or end_span is None:
        return

    start_index = start_span["index"]
    end_index = end_span["index"]
    start_run = start_span["run"]
    end_run = end_span["run"]

    start_text = start_run.text or ""
    prefix = start_text[:start - start_span["start"]]

    if start_index == end_index:
        suffix = start_text[end - start_span["start"]:]
        start_run.text = prefix + new_text + suffix
        return

    end_text = end_run.text or ""
    suffix = end_text[end - end_span["start"]:]
    start_run.text = prefix + new_text

    for index in range(start_index + 1, end_index):
        runs[index].text = ""

    end_run.text = suffix


def _apply_rules_to_paragraph(paragraph, prepared_rules: List[Dict[str, str]]) -> int:
    runs = paragraph.runs
    if not runs:
        return 0

    total_count = 0

    for rule in prepared_rules:
        old_text = rule["old"]
        full_text = "".join(run.text or "" for run in runs)
        if not full_text or rule["first_char"] not in set(full_text) or old_text not in full_text:
            continue

        occurrences = [
            start
            for start in _find_occurrences(full_text, old_text)
            if not _is_occurrence_inside_replacement(full_text, start, old_text, rule["new"])
        ]
        if not occurrences:
            continue

        spans = _build_run_spans(runs)

        for start in reversed(occurrences):
            _replace_range_in_runs(
                runs,
                spans,
                start,
                start + len(old_text),
                rule["new"],
            )
            total_count += 1

    return total_count


def replace_in_docx(file_path: str, rules: List[Tuple[str, str]], output_path: str) -> int:
    from docx import Document

    doc = Document(file_path)
    prepared_rules = _prepare_rules(rules)
    total_count = 0

    for paragraph in doc.paragraphs:
        total_count += _apply_rules_to_paragraph(paragraph, prepared_rules)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    total_count += _apply_rules_to_paragraph(paragraph, prepared_rules)

    for section in doc.sections:
        for container in (section.header, section.footer):
            for paragraph in container.paragraphs:
                total_count += _apply_rules_to_paragraph(paragraph, prepared_rules)
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            total_count += _apply_rules_to_paragraph(paragraph, prepared_rules)

    doc.save(output_path)
    return total_count


def replace_in_workbook(workbook, rules: List[Tuple[str, str]]) -> int:
    prepared_rules = _prepare_rules(rules)
    total_count = 0

    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value:
                    continue
                if cell.value.startswith("="):
                    continue

                new_value, replaced_count = _replace_text_with_rules(cell.value, prepared_rules)
                if replaced_count:
                    cell.value = new_value
                    total_count += replaced_count

    return total_count


def replace_in_xlsx(file_path: str, rules: List[Tuple[str, str]], output_path: str) -> int:
    from openpyxl import load_workbook

    if os.path.abspath(file_path) != os.path.abspath(output_path):
        shutil.copy2(file_path, output_path)

    keep_vba = os.path.splitext(output_path)[1].lower() == ".xlsm"
    workbook = load_workbook(output_path, keep_vba=keep_vba)
    try:
        total_count = replace_in_workbook(workbook, rules)
        workbook.save(output_path)
    finally:
        workbook.close()

    return total_count


def _is_group_shape(shape) -> bool:
    return hasattr(shape, "shapes") and "GROUP" in str(getattr(shape, "shape_type", ""))


def _replace_in_ppt_text_frame(text_frame, prepared_rules: List[Dict[str, str]]) -> int:
    total_count = 0
    for paragraph in text_frame.paragraphs:
        total_count += _apply_rules_to_paragraph(paragraph, prepared_rules)
    return total_count


def _visit_ppt_shape(shape, text_frame_handler: Callable) -> int:
    total_count = 0

    if getattr(shape, "has_text_frame", False):
        total_count += text_frame_handler(shape.text_frame)

    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    total_count += text_frame_handler(cell.text_frame)

    if _is_group_shape(shape):
        for child_shape in shape.shapes:
            total_count += _visit_ppt_shape(child_shape, text_frame_handler)

    return total_count


def replace_in_pptx(file_path: str, rules: List[Tuple[str, str]], output_path: str) -> int:
    from pptx import Presentation

    presentation = Presentation(file_path)
    prepared_rules = _prepare_rules(rules)
    total_count = 0

    for slide in presentation.slides:
        for shape in slide.shapes:
            total_count += _visit_ppt_shape(
                shape,
                lambda text_frame: _replace_in_ppt_text_frame(text_frame, prepared_rules),
            )

    presentation.save(output_path)
    return total_count


def sanitize_windows_filename_stem(stem: str) -> str:
    """把文件名主体转换为可在 Windows 上安全保存的形式。"""
    sanitized = stem.translate(_WINDOWS_INVALID_FILENAME_REPLACEMENTS)
    sanitized = "".join("＿" if ord(char) < 32 else char for char in sanitized)
    sanitized = sanitized.rstrip(" .")
    if not sanitized:
        return "_"

    device_stem = sanitized.split(".", 1)[0].upper()
    if device_stem in _WINDOWS_RESERVED_FILENAME_STEMS:
        sanitized = f"_{sanitized}"
    return sanitized


def apply_rules_to_filename(filename: str, rules: List[Tuple[str, str]]) -> str:
    """替换文件名主体并清理 Windows 禁用字符，保留原扩展名不变。"""
    stem, ext = os.path.splitext(filename)
    replaced_stem, _count = _replace_text_with_rules(stem, _prepare_rules(rules))
    return sanitize_windows_filename_stem(replaced_stem) + ext


def _with_output_index(path: str, index: int) -> str:
    directory, filename = os.path.split(path)
    stem, ext = os.path.splitext(filename)
    return os.path.join(directory, f"{stem}_{index}{ext}")


def _same_file_path(first: str, second: str) -> bool:
    return os.path.normcase(os.path.abspath(os.path.realpath(first))) == os.path.normcase(
        os.path.abspath(os.path.realpath(second))
    )


def _file_identity(path: str) -> str:
    return os.path.normcase(os.path.abspath(os.path.realpath(os.fspath(path))))


def get_output_path(
    file_path: str,
    rules: Optional[List[Tuple[str, str]]] = None,
    output_dir: Optional[str] = None,
    reserved_output_paths: Optional[set] = None,
) -> str:
    """生成输出路径。

    - 文件名按规则替换后若与源文件相同，直接覆盖源文件（不再加「_已替换」）。
    - 若目标路径已被其他已有文件占用，或本批任务中其他文件已占用，则追加 _1、_2…
    """
    dir_path = output_dir or os.path.dirname(file_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    filename = os.path.basename(file_path)
    if rules:
        filename = apply_rules_to_filename(filename, rules)

    output_path = os.path.join(dir_path, filename)
    reserved = reserved_output_paths if reserved_output_paths is not None else set()
    candidate = output_path
    index = 1
    while _file_identity(candidate) in reserved or (
        os.path.exists(candidate) and not _same_file_path(file_path, candidate)
    ):
        candidate = _with_output_index(output_path, index)
        index += 1

    reserved.add(_file_identity(candidate))
    return candidate


def _result_key(results: Dict[str, int], filename: str, file_path: str) -> str:
    if filename not in results:
        return filename

    parent = os.path.dirname(file_path)
    candidate = f"{filename} ({parent})"
    index = 2
    while candidate in results:
        candidate = f"{filename} ({parent}, {index})"
        index += 1
    return candidate


def _format_processing_error(exc: Exception) -> str:
    if isinstance(exc, PermissionError):
        return "文件无法读写，请先关闭正在打开的 Word/Excel/PPT 文件后重试"
    return str(exc)


def batch_replace(
    file_paths: List[str],
    rules: List[Tuple[str, str]],
    output_dir: Optional[str] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> Tuple[Dict[str, int], Optional[str]]:
    """批量处理 Office 文件，返回替换结果和错误信息。"""
    results = {}
    errors = []
    reserved_output_paths = set()

    for index, file_path in enumerate(file_paths):
        filename = os.path.basename(file_path)
        if progress_callback:
            progress_callback(index + 1, len(file_paths), filename)

        try:
            output_path = get_output_path(
                file_path,
                rules,
                output_dir,
                reserved_output_paths=reserved_output_paths,
            )
            ext = os.path.splitext(file_path)[1].lower()

            if ext == ".docx":
                count = replace_in_docx(file_path, rules, output_path)
            elif ext in (".xlsx", ".xlsm"):
                count = replace_in_xlsx(file_path, rules, output_path)
            elif ext == ".pptx":
                count = replace_in_pptx(file_path, rules, output_path)
            else:
                errors.append(f"{filename}: 不支持的文件格式")
                continue

            results[_result_key(results, filename, file_path)] = count
        except Exception as exc:
            errors.append(f"{filename}: {_format_processing_error(exc)}")

    return results, "\n".join(errors) if errors else None
