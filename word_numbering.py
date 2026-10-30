"""Word 自动编号的共享解析器，供表格和条款扫描复用。"""
import re
from typing import Dict, List, Optional, Tuple
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

_CN_NUMERAL = "零一二三四五六七八九"

class _NumberingTracker:
    """重建 Word 自动编号（w:numPr），把编号文本拼回段落文本前。

    支持 decimal / 中文编号 / 字母 / 罗马数字 / ①②③ / bullet 等常见
    numFmt，按 (numId, ilvl) 维护计数器，上级层级出现时重置下级计数。
    """

    _MAX_STYLE_CHAIN = 8

    def __init__(self, document):
        self._counters: Dict[Tuple[str, int], int] = {}
        self._nums: Dict[str, Dict[str, object]] = {}
        self._abstracts: Dict[str, Dict[int, object]] = {}
        self._load_numbering(document)

    def numbered_text(self, paragraph: Paragraph) -> str:
        text = paragraph.text or ""
        try:
            numpr = self._effective_numpr(paragraph)
        except Exception:
            numpr = None
        if numpr is None:
            return text

        num_id = self._child_val(numpr, "w:numId")
        if num_id is None or num_id == "0":
            return text
        ilvl = self._child_val(numpr, "w:ilvl") or "0"

        prefix = self._next_number_text(num_id, ilvl)
        if not prefix:
            return text
        return prefix + text

    def _load_numbering(self, document) -> None:
        element = self._numbering_element(document)
        if element is None:
            return

        for abstract in element.findall(qn("w:abstractNum")):
            abstract_id = abstract.get(qn("w:abstractNumId"))
            levels: Dict[int, object] = {}
            for lvl in abstract.findall(qn("w:lvl")):
                try:
                    levels[int(lvl.get(qn("w:ilvl")))] = lvl
                except (TypeError, ValueError):
                    continue
            self._abstracts[abstract_id] = levels

        for num in element.findall(qn("w:num")):
            num_id = num.get(qn("w:numId"))
            abstract_ref = num.find(qn("w:abstractNumId"))
            overrides: Dict[int, int] = {}
            level_overrides = {}
            for lvl_override in num.findall(qn("w:lvlOverride")):
                level = lvl_override.find(qn('w:lvl'))
                if level is not None:
                    try:
                        level_overrides[int(lvl_override.get(qn('w:ilvl')))] = level
                    except (TypeError,ValueError):
                        pass
                start_override = lvl_override.find(qn("w:startOverride"))
                if start_override is None:
                    continue
                try:
                    overrides[int(lvl_override.get(qn("w:ilvl")))] = int(
                        start_override.get(qn("w:val"))
                    )
                except (TypeError, ValueError):
                    continue
            self._nums[num_id] = {
                "abstract": abstract_ref.get(qn("w:val")) if abstract_ref is not None else None,
                "overrides": overrides,
                "levels": level_overrides,
            }

    @staticmethod
    def _numbering_element(document):
        part = None
        try:
            part = document.part.part_related_by(RT.NUMBERING)
        except Exception:
            part = None
        if part is None:
            try:
                for rel in document.part.rels.values():
                    if rel.reltype == RT.NUMBERING:
                        part = rel.target_part
                        break
            except Exception:
                return None
        if part is None:
            return None
        return getattr(part, "element", None)

    def _effective_numpr(self, paragraph: Paragraph):
        values = {}
        def collect(numpr):
            if numpr is not None:
                for tag in ('w:numId','w:ilvl'):
                    value = self._child_val(numpr,tag)
                    if value is not None:values.setdefault(tag,value)
        ppr = paragraph._p.find(qn("w:pPr"))
        if ppr is not None:
            collect(ppr.find(qn("w:numPr")))

        # 编号也可能挂在段落样式上：沿样式基类链向上找。
        style = None
        try:
            style = paragraph.style
        except Exception:
            style = None
        for _ in range(self._MAX_STYLE_CHAIN):
            if style is None:
                break
            element = getattr(style, "element", None)
            if element is None:
                break
            style_ppr = element.find(qn("w:pPr"))
            if style_ppr is not None:
                style_numpr = style_ppr.find(qn("w:numPr"))
                collect(style_numpr)
            try:
                base = style.base_style
            except Exception:
                break
            if base is style:
                break
            style = base
        if 'w:numId' not in values:return None
        merged = OxmlElement('w:numPr')
        for tag,value in values.items():
            child = OxmlElement(tag);child.set(qn('w:val'),value);merged.append(child)
        return merged

    def _next_number_text(self, num_id: str, ilvl: str) -> str:
        info = self._nums.get(num_id)
        if not info or info.get("abstract") is None:
            return ""
        levels = dict(self._abstracts.get(info["abstract"]) or {})
        levels.update(info.get('levels',{}))
        try:
            level_index = int(ilvl)
        except (TypeError, ValueError):
            level_index = 0
        if level_index not in levels and 0 in levels:
            level_index = 0
        lvl = levels.get(level_index)
        if lvl is None:
            return ""

        num_fmt = self._child_val(lvl, "w:numFmt") or "decimal"
        lvl_text = self._child_val(lvl, "w:lvlText")
        if lvl_text is None:
            lvl_text = "%1."

        self._advance_counter(num_id, info, levels, level_index)

        if num_fmt == "none":
            return ""
        if num_fmt == "bullet":
            return re.sub(r"%[1-9]", "", lvl_text)

        def render(level_digit: int) -> str:
            index = level_digit - 1
            value = self._counters.get((num_id, index))
            if value is None:
                value = self._level_start(info, levels, index)
            sub_fmt = self._child_val(levels.get(index), "w:numFmt") or num_fmt
            return _format_number(value, sub_fmt or "decimal")

        rendered = re.sub(r"%([1-9])", lambda m: render(int(m.group(1))), lvl_text)
        suffix = self._child_val(lvl,'w:suff')
        return rendered + (' ' if suffix=='space' else '\t' if suffix=='tab' else '')

    def _advance_counter(self, num_id: str, info: Dict[str, object], levels: Dict[int, object], level_index: int) -> None:
        # 上级层级重新出现时，Word 会重置所有更深层级的计数。
        for key in [key for key in self._counters if key[0] == num_id and key[1] > level_index]:
            restart = self._child_val(levels.get(key[1]),'w:lvlRestart')
            if restart == '0':continue
            if restart is not None:
                try:
                    if level_index >= int(restart):continue
                except ValueError:
                    pass
            del self._counters[key]

        key = (num_id, level_index)
        current = self._counters.get(key)
        if current is None:
            self._counters[key] = self._level_start(info, levels, level_index)
        else:
            self._counters[key] = current + 1

    def _level_start(self, info: Dict[str, object], levels: Dict[int, object], level_index: int) -> int:
        override = info.get("overrides", {}).get(level_index)
        if override is not None:
            return override
        start = self._child_val(levels.get(level_index), "w:start")
        try:
            return int(start)
        except (TypeError, ValueError):
            return 1

    @staticmethod
    def _child_val(element, tag: str) -> Optional[str]:
        if element is None:
            return None
        child = element.find(qn(tag))
        if child is None:
            return None
        return child.get(qn("w:val"))


def _format_number(value: int, num_fmt: str) -> str:
    if num_fmt == "zeroDecimal":
        return f"{max(value, 0):02d}"
    if num_fmt == "decimalEnclosedCircle":
        return _circled_number(value)
    if num_fmt in {
        "chineseCounting",
        "chineseCountingThousand",
        "chineseLegalSimplified",
        "ideographDigital",
        "japaneseCounting",
    }:
        return _int_to_chinese(value)
    if num_fmt == "ideographTraditional":
        return _heavenly_stem(value)
    if num_fmt == "upperLetter":
        return _letter_number(value).upper()
    if num_fmt == "lowerLetter":
        return _letter_number(value)
    if num_fmt == "upperRoman":
        return _roman_number(value).upper()
    if num_fmt == "lowerRoman":
        return _roman_number(value)
    return str(value)


def _int_to_chinese(value: int) -> str:
    if value <= 0 or value > 9999:
        return str(value)
    if value < 10:
        return _CN_NUMERAL[value]

    units = ["", "十", "百", "千"]
    digits = [int(digit) for digit in str(value)]
    length = len(digits)
    parts: List[str] = []
    zero_pending = False

    for position, digit in enumerate(digits):
        unit = units[length - 1 - position]
        if digit == 0:
            zero_pending = True
            continue
        if zero_pending and parts:
            parts.append("零")
        zero_pending = False
        if digit == 1 and unit == "十" and not parts:
            parts.append("十")
        else:
            parts.append(_CN_NUMERAL[digit] + unit)

    return "".join(parts) or "零"


def _heavenly_stem(value: int) -> str:
    stems = "甲乙丙丁戊己庚辛壬癸"
    if value <= 0:
        return str(value)
    if value <= 10:
        return stems[value - 1]
    return str(value)


def _letter_number(value: int) -> str:
    if value <= 0:
        return str(value)
    letters: List[str] = []
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(ord("a") + remainder))
    return "".join(reversed(letters))


def _roman_number(value: int) -> str:
    if value <= 0 or value >= 4000:
        return str(value)
    table = (
        (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
        (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
        (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
    )
    parts: List[str] = []
    for threshold, symbol in table:
        while value >= threshold:
            parts.append(symbol)
            value -= threshold
    return "".join(parts)


def _circled_number(value: int) -> str:
    if 1 <= value <= 20:
        return chr(0x2460 + value - 1)
    if 21 <= value <= 35:
        return chr(0x3251 + value - 21)
    if 36 <= value <= 50:
        return chr(0x32B1 + value - 36)
    return str(value)


def document_numbered_text(document):
    """按 XML 文档顺序计算一次，正文、嵌套表格和合并格共用计数。"""
    tracker = _NumberingTracker(document)
    return {element: tracker.numbered_text(Paragraph(element, document))
            for element in document.element.body.iter(qn("w:p"))}


class NumberingSnapshot:
    def __init__(self, texts):self.texts=texts

    def numbered_text(self, paragraph):
        return self.texts[paragraph._p] if paragraph._p in self.texts else paragraph.text
