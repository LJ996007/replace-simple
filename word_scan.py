"""一次解析生成表格、条款快照；导出前检查源文件是否变化。"""

import os
from contextlib import ExitStack, nullcontext
from dataclasses import replace

from docx import Document

from legacy_office import LegacyOfficeSession, temporary_docx_source
from symbol_clause_extractor import (
    export_symbol_clauses_to_excel, get_symbol_output_path,
    scan_symbol_clause_sections, strip_clause_symbols,
)
from word_table_exporter import (
    _file_identity, _format_export_error, _result_key,
    export_word_tables_to_excel, get_table_export_output_path, scan_word_tables,
)


def file_stamp(path):
    stat = os.stat(path)
    # ponytail: 用大小和纳秒修改时间检测日常编辑；需要防篡改时改用内容哈希。
    return stat.st_size, stat.st_mtime_ns


def validate_scan_files(file_paths, stamps):
    changed = []
    for path in file_paths:
        try:
            unchanged = stamps.get(_file_identity(path)) == file_stamp(path)
        except OSError:
            unchanged = False
        if not unchanged:
            changed.append(path)
    if changed:
        raise ValueError("文件已变化、移除或尚未扫描，请重新扫描：\n" + "\n".join(changed))


def scan_word_content(file_paths, symbols, progress_callback=None):
    tables, sections, stamps = [], [], {}
    table_skipped, symbol_skipped, errors = {}, {}, []
    with ExitStack() as stack:
        session = None
        for index, path in enumerate(file_paths, 1):
            if progress_callback:
                progress_callback(index, len(file_paths), os.path.basename(path), "正在扫描")
            try:
                stamp = file_stamp(path)
                extension = os.path.splitext(path)[1].lower()
                if extension not in (".doc", ".docx"):
                    raise ValueError("不支持的文件格式")
                context = nullcontext(path)
                if extension == ".doc":
                    if session is None:
                        session = stack.enter_context(LegacyOfficeSession())
                    context = temporary_docx_source(path, session)
                with context as readable:
                    document = Document(readable)
                    file_tables = scan_word_tables(path, document=document)
                    file_sections = scan_symbol_clause_sections(path, symbols=symbols, document=document)
                if file_stamp(path) != stamp:
                    raise ValueError("文件在扫描过程中发生变化，请重新扫描")
                stamps[_file_identity(path)] = stamp
                tables.extend(file_tables)
                sections.extend(file_sections)
                if not file_tables:
                    table_skipped[_result_key(table_skipped, os.path.basename(path), path)] = "未找到表格"
                if not file_sections:
                    symbol_skipped[_result_key(symbol_skipped, os.path.basename(path), path)] = "未找到带符号条款"
            except Exception as exc:
                errors.append(f"{path}: {_format_export_error(exc)}")
    error = "\n".join(errors) or None
    return (tables, table_skipped, error), (sections, symbol_skipped, error), stamps


def export_scanned_content(tables, sections, stamps, output_dir, symbols,
                           keep_symbols_in_text, progress_callback=None):
    paths = list(dict.fromkeys(item.file_path for item in [*tables, *sections]))
    validate_scan_files(paths, stamps)
    results = [{}, {}]
    errors = [[], []]
    jobs = []
    for kind, items in enumerate((tables, sections)):
        grouped = {}
        for item in items:
            grouped.setdefault(item.file_path, []).append(item)
        jobs.extend((kind, path, selected) for path, selected in grouped.items())
    for index, (kind, path, items) in enumerate(jobs, 1):
        if progress_callback:
            progress_callback(index, len(jobs), os.path.basename(path))
        try:
            validate_scan_files([path], stamps)
            if kind == 0:
                output = get_table_export_output_path(path, output_dir)
                count = export_word_tables_to_excel(path, output, exported_tables=[item.table for item in items])
                result = {"tables": count, "output_path": output}
            else:
                clauses = [clause for item in items for clause in item.clauses]
                if not keep_symbols_in_text:
                    clauses = [replace(clause, text=strip_clause_symbols(clause.text, symbols)) for clause in clauses]
                output = get_symbol_output_path(path, output_dir)
                count = export_symbol_clauses_to_excel(path, output, clauses=clauses)
                result = {"count": count, "output_path": output}
            results[kind][_result_key(results[kind], os.path.basename(path), path)] = result
        except Exception as exc:
            errors[kind].append(f"{path}: {_format_export_error(exc)}")
    return results[0], {}, "\n".join(errors[0]) or None, results[1], {}, "\n".join(errors[1]) or None
