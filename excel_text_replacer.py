"""Replace Excel cell text without resaving unrelated package contents."""

import html
import io
import re
from collections import Counter
from zipfile import ZipFile
from xml.etree import ElementTree


def _elements(name):
    return re.compile(rf'<(?:[\w.-]+:)?{name}(?=[\s>])[^>]*(?<!/)>.*?</(?:[\w.-]+:)?{name}\s*>', re.S)


_CELLS = _elements('c')
_ITEMS = re.compile(r'<(?:[\w.-]+:)?si(?=[\s/>])[^>]*?/>|<(?:[\w.-]+:)?si(?=[\s>])[^>]*>.*?</(?:[\w.-]+:)?si\s*>', re.S)
_TEXT = re.compile(r'<(?:[\w.-]+:)?t(?=[\s>])[^>]*(?<!/)>(.*?)</(?:[\w.-]+:)?t\s*>', re.S)
_TOKENS = re.compile(r'&(?:#x[\da-fA-F]+|#\d+|amp|lt|gt|quot|apos);|_x[\da-fA-F]{4}_|.', re.S)


def _decode_text(raw):
    characters, positions = [], []
    for token in _TOKENS.finditer(raw):
        value = token.group()
        decoded = chr(int(value[2:6], 16)) if value.startswith('_x') else html.unescape(value)
        characters.append(decoded)
        positions.extend([(token.start(), token.end())] * len(decoded))
    return ''.join(characters), positions


def _escape_text(text):
    text = re.sub(r'_x[\da-fA-F]{4}_', lambda match: '_x005F_' + match.group()[1:], text)
    text = ''.join(f'_x{ord(character):04X}_' if ord(character) < 32 and character not in '\n\t\r' else character
                   for character in text)
    return html.escape(text, quote=False).replace('\r', '&#13;')


def _patch_text(container, prepared):
    from string_replacer import _find_rule_regions

    phonetics = [match.span() for match in _elements('rPh').finditer(container)]
    nodes, full_text = [], ''
    for match in _TEXT.finditer(container):
        if any(start <= match.start() < end for start, end in phonetics):
            continue
        text, positions = _decode_text(match.group(1))
        nodes.append((match, len(full_text), text, positions))
        full_text += text
    regions = _find_rule_regions(full_text, prepared)
    if not regions:
        return container, 0
    edits = []
    for match, offset, text, positions in nodes:
        changed = False
        for start, end, replacement in regions:
            local_start, local_end = max(0, start - offset), min(len(text), end - offset)
            if local_start >= local_end:
                continue
            edits.append((match.start(1) + positions[local_start][0],
                          match.start(1) + positions[local_end - 1][1],
                          _escape_text(replacement) if offset <= start else ''))
            changed = True
        if changed and 'xml:space=' not in container[match.start():match.start(1)]:
            edits.append((match.start(1) - 1, match.start(1) - 1, ' xml:space="preserve"'))
    for start, end, replacement in sorted(edits, reverse=True):
        container = container[:start] + replacement + container[end:]
    return container, len(regions)


def replace_excel_package(path, rules):
    from string_replacer import _prepare_rules, _replace_text_with_rules, _workbook_cell_text

    prepared = _prepare_rules(rules)
    changes, shared_usage = {}, Counter()
    total = 0
    with ZipFile(path) as source:
        date_styles = set()
        if 'xl/styles.xml' in source.namelist():
            from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format, is_timedelta_format

            styles = ElementTree.fromstring(source.read('xl/styles.xml'))
            custom_formats = {
                int(node.get('numFmtId')): node.get('formatCode', '')
                for node in styles.findall('{*}numFmts/{*}numFmt')
            }
            for index, node in enumerate(styles.findall('{*}cellXfs/{*}xf')):
                format_id = int(node.get('numFmtId', '0'))
                code = custom_formats.get(format_id, BUILTIN_FORMATS.get(format_id, ''))
                if is_date_format(code) or is_timedelta_format(code):
                    date_styles.add(index)
        sheets = [name for name in source.namelist()
                  if name.startswith('xl/worksheets/') and name.endswith('.xml')]
        for name in sheets:
            original = source.read(name)
            document = original.decode('utf-8')

            def patch_cell(match):
                nonlocal total
                cell = match.group()
                if re.search(r'<(?:[\w.-]+:)?f(?=[\s/>])', cell):
                    return cell
                opening_end = cell.index('>') + 1
                cell_type = re.search(r'\bt=[\'"]([^\'"]+)[\'"]', cell[:opening_end])
                cell_type = cell_type.group(1) if cell_type else 'n'
                value = re.search(r'<(?:[\w.-]+:)?v(?=[\s>])[^>]*>(.*?)</(?:[\w.-]+:)?v\s*>', cell, re.S)
                if cell_type == 's' and value:
                    shared_usage[int(value.group(1))] += 1
                    return cell
                if cell_type in ('inlineStr', 'str'):
                    if cell_type == 'inlineStr':
                        updated, count = _patch_text(cell, prepared)
                    elif value:
                        text, _positions = _decode_text(value.group(1))
                        replacement, count = _replace_text_with_rules(text, prepared)
                        updated = cell[:value.start(1)] + _escape_text(replacement) + cell[value.end(1):] if count else cell
                    else:
                        return cell
                    total += count
                    return updated
                if cell_type == 'n' and value:
                    cell_style = re.search(r'\bs=[\'"](\d+)[\'"]', cell[:opening_end])
                    if int(cell_style.group(1) if cell_style else 0) in date_styles:
                        return cell
                    try:
                        text = _workbook_cell_text(float(value.group(1)))
                    except ValueError:
                        return cell
                    replacement, count = _replace_text_with_rules(text or '', prepared)
                    if count:
                        opening = re.sub(r'\s+t=[\'"][^\'"]+[\'"]', '', cell[:opening_end])
                        opening = opening[:-1] + ' t="inlineStr">'
                        prefix = re.match(r'<([\w.-]+:)?c', cell).group(1) or ''
                        content = f'<{prefix}is><{prefix}t xml:space="preserve">{_escape_text(replacement)}</{prefix}t></{prefix}is>'
                        total += count
                        return opening + cell[opening_end:value.start()] + content + cell[value.end():]
                return cell

            updated = _CELLS.sub(patch_cell, document)
            if updated != document:
                changes[name] = updated.encode('utf-8')
        shared_path = 'xl/sharedStrings.xml'
        if shared_path in source.namelist():
            document = source.read(shared_path).decode('utf-8')
            item_index = -1

            def patch_shared(match):
                nonlocal item_index, total
                item_index += 1
                if not shared_usage[item_index]:
                    return match.group()
                updated, count = _patch_text(match.group(), prepared)
                total += count * shared_usage[item_index]
                return updated

            updated = _ITEMS.sub(patch_shared, document)
            if updated != document:
                changes[shared_path] = updated.encode('utf-8')
        if not changes:
            return total
        buffer = io.BytesIO()
        with ZipFile(buffer, 'w') as output:
            output.comment = source.comment
            for entry in source.infolist():
                output.writestr(entry, changes.get(entry.filename, source.read(entry)))
    with open(path, 'wb') as destination:
        destination.write(buffer.getvalue())
    return total
