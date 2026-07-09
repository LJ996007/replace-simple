"""
Extract common project fields from tender .docx files as replacement rules.
"""

import re
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from docx import Document


FIELD_DEFINITIONS: Sequence[Tuple[str, str, Sequence[str]]] = (
    ("项目名称", "{项目名称}", ("项目名称", "采购项目名称", "招标项目名称")),
    ("项目编号", "{项目编号}", ("项目编号", "采购项目编号", "招标编号", "采购编号")),
    ("采购人", "{采购人}", ("采购人", "采购人名称", "采购单位", "采购单位名称", "招标人", "招标人名称", "建设单位", "建设单位名称")),
    ("采购代理机构", "{采购代理机构}", ("采购代理机构", "采购代理机构名称", "代理机构", "代理机构名称", "招标代理机构", "招标代理机构名称", "采购代理")),
    ("预算金额", "{预算金额}", ("预算金额", "项目预算", "项目预算金额", "采购预算", "采购预算金额", "预算价")),
    ("最高限价", "{最高限价}", ("最高限价", "最高投标限价", "最高限价金额")),
    ("采购方式", "{采购方式}", ("采购方式", "招标方式")),
    ("合同履行期限", "{合同履行期限}", ("合同履行期限", "履约期限", "服务期限", "交付期限", "工期")),
    ("提交投标文件截止时间", "{提交投标文件截止时间}", ("提交投标文件截止时间", "投标截止时间", "递交投标文件截止时间")),
    ("开标时间", "{开标时间}", ("开标时间", "开启时间")),
    ("开标地点", "{开标地点}", ("开标地点", "开启地点")),
)

MAX_VALUE_LENGTH = 160
CONTEXT_NAME_ALIASES = ("名称", "名 称", "单位名称", "机构名称")


def _clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\u3000", " ")
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip(" \t\r\n:：;；")


def _normalize_key(value: str) -> str:
    text = _clean_text(value)
    text = re.sub(r"\s+", "", text)
    text = text.strip(":：;；")
    return text


def _key_matches_alias(normalized_key: str, alias: str) -> bool:
    normalized_alias = _normalize_key(alias)
    if normalized_key == normalized_alias:
        return True

    suffix = normalized_key.removeprefix(normalized_alias)
    if not suffix or suffix == normalized_key:
        return False

    return bool(re.fullmatch(r"(?:[（(][^）)]{1,12}[）)])+", suffix))


def _trim_value(value: str) -> str:
    text = _clean_text(value)
    text = re.sub(r"^[：:；;\s]+", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:MAX_VALUE_LENGTH].strip()


def _looks_like_key(text: str, aliases: Iterable[str]) -> bool:
    normalized = _normalize_key(text)
    return any(_key_matches_alias(normalized, alias) for alias in aliases)


def _looks_like_any_key(text: str) -> bool:
    normalized = _normalize_key(text)
    if not normalized:
        return False
    for _field_name, _placeholder, aliases in FIELD_DEFINITIONS:
        if any(_key_matches_alias(normalized, alias) for alias in aliases):
            return True
    return False


def _value_from_labeled_text(text: str, aliases: Iterable[str]) -> Optional[str]:
    line = _clean_text(text)
    if not line:
        return None

    prefix = r"^(?:[（(]?[一二三四五六七八九十\d]+[）)、.．]\s*)?"
    for alias in aliases:
        alias_pattern = re.escape(alias)
        key_pattern = rf"{alias_pattern}(?:[（(][^）)]{{1,12}}[）)])*"
        patterns = (
            rf"{prefix}{key_pattern}\s*[:：]\s*(.+)$",
            rf"{prefix}{key_pattern}\s+(.+)$",
        )
        for pattern in patterns:
            match = re.search(pattern, line)
            if match:
                value = _trim_value(match.group(1))
                if value and not _looks_like_key(value, aliases):
                    return value
    return None


def _set_if_missing(found: Dict[str, str], field_name: str, value: Optional[str]) -> None:
    if field_name in found:
        return
    value = _trim_value(value or "")
    if value:
        found[field_name] = value


def _field_name_from_context(text: str) -> Optional[str]:
    normalized = _normalize_key(text)
    if not normalized:
        return None

    procurement_contexts = (
        "采购人信息",
        "采购人联系方式",
        "采购单位信息",
        "招标人信息",
    )
    agency_contexts = (
        "采购代理机构信息",
        "代理机构信息",
        "招标代理机构信息",
        "采购代理信息",
    )
    if any(item in normalized for item in agency_contexts):
        return "采购代理机构"
    if any(item in normalized for item in procurement_contexts):
        return "采购人"
    return None


def _value_from_context_labeled_text(text: str) -> Optional[str]:
    return _value_from_labeled_text(text, CONTEXT_NAME_ALIASES)


def _extract_from_table(table, found: Dict[str, str]) -> None:
    active_context: Optional[str] = None
    for row in table.rows:
        cells = [_clean_text(cell.text) for cell in row.cells]
        for index, cell_text in enumerate(cells):
            if not cell_text:
                continue

            context_field = _field_name_from_context(cell_text)
            if context_field:
                active_context = context_field
                continue

            if active_context:
                value = _value_from_context_labeled_text(cell_text)
                if value:
                    _set_if_missing(found, active_context, value)
                elif _looks_like_key(cell_text, CONTEXT_NAME_ALIASES):
                    for next_text in cells[index + 1:]:
                        if not next_text:
                            continue
                        if _looks_like_any_key(next_text):
                            break
                        _set_if_missing(found, active_context, next_text)
                        break

            for field_name, _placeholder, aliases in FIELD_DEFINITIONS:
                value = _value_from_labeled_text(cell_text, aliases)
                if value:
                    _set_if_missing(found, field_name, value)
                    continue

                if not _looks_like_key(cell_text, aliases):
                    continue

                for next_text in cells[index + 1:]:
                    if not next_text:
                        continue
                    if _looks_like_any_key(next_text):
                        break
                    if active_context and _value_from_context_labeled_text(next_text):
                        break
                    else:
                        _set_if_missing(found, field_name, next_text)
                        break

        for cell in row.cells:
            for nested_table in cell.tables:
                _extract_from_table(nested_table, found)


def _extract_from_paragraphs(paragraphs, found: Dict[str, str]) -> None:
    active_context: Optional[str] = None
    for paragraph in paragraphs:
        text = _clean_text(paragraph.text)
        if not text:
            continue

        context_field = _field_name_from_context(text)
        if context_field:
            active_context = context_field
            value = _value_from_context_labeled_text(text)
            if value:
                _set_if_missing(found, active_context, value)
            continue

        if active_context:
            value = _value_from_context_labeled_text(text)
            if value:
                _set_if_missing(found, active_context, value)
                active_context = None

        for field_name, _placeholder, aliases in FIELD_DEFINITIONS:
            _set_if_missing(found, field_name, _value_from_labeled_text(text, aliases))


def _extract_from_document_part(part, found: Dict[str, str]) -> None:
    _extract_from_paragraphs(part.paragraphs, found)
    for table in part.tables:
        _extract_from_table(table, found)


def extract_project_info(docx_path: str) -> Dict[str, str]:
    """Return common tender project fields extracted from a .docx file."""
    document = Document(docx_path)
    found: Dict[str, str] = {}

    _extract_from_document_part(document, found)
    for section in document.sections:
        _extract_from_document_part(section.header, found)
        _extract_from_document_part(section.footer, found)

    return found


def extract_project_info_rules(docx_path: str) -> List[Tuple[str, str]]:
    """Return replacement rules using stable placeholders as the source text."""
    info = extract_project_info(docx_path)
    rules: List[Tuple[str, str]] = []
    for field_name, placeholder, _aliases in FIELD_DEFINITIONS:
        value = info.get(field_name)
        if value:
            rules.append((placeholder, value))
    return rules
