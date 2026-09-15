"""通过 Microsoft Office 或 WPS COM 处理旧版 Office 二进制格式。

本模块只在实际处理 .doc/.xls/.ppt 时加载 pywin32。新格式文件仍由
python-docx/openpyxl/python-pptx 直接处理，因此没有安装 Office/WPS 时不会影响
现有功能。
"""

from __future__ import annotations

import contextlib
import datetime as _datetime
import gc
import os
import shutil
import tempfile
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Tuple
from file_io import atomic_output_path


LEGACY_EXTENSIONS = (".doc", ".xls", ".ppt")

_APPLICATION_CANDIDATES = {
    "word": (
        ("Microsoft Word", "Word.Application"),
        ("WPS Writer", "KWPS.Application"),
        ("WPS Writer", "wps.Application"),
    ),
    "excel": (
        ("Microsoft Excel", "Excel.Application"),
        ("WPS Spreadsheets", "KET.Application"),
        ("WPS Spreadsheets", "et.Application"),
    ),
    "powerpoint": (
        ("Microsoft PowerPoint", "PowerPoint.Application"),
        ("WPS Presentation", "KWPP.Application"),
        ("WPS Presentation", "wpp.Application"),
    ),
}

_APPLICATION_LABELS = {
    "word": "Word/WPS Writer",
    "excel": "Excel/WPS Spreadsheets",
    "powerpoint": "PowerPoint/WPS Presentation",
}


class LegacyOfficeError(RuntimeError):
    """旧格式 Office 自动化的可读错误。"""


class OfficeApplicationUnavailableError(LegacyOfficeError):
    """电脑上没有可驱动的对应 Office/WPS 组件。"""


class OfficeFileLockedError(LegacyOfficeError):
    """文件被占用或以只读方式打开。"""


class OfficePasswordProtectedError(LegacyOfficeError):
    """文件需要密码。"""


class OfficeProtectedDocumentError(LegacyOfficeError):
    """文档或工作表受到保护，不能修改。"""


class OfficeSaveError(LegacyOfficeError):
    """Office 无法保存处理结果。"""


def _load_com_modules():
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise OfficeApplicationUnavailableError(
            "程序缺少旧格式处理组件 pywin32，请重新安装完整版程序"
        ) from exc
    return pythoncom, win32com.client


def _safe_setattr(obj, name: str, value) -> None:
    try:
        setattr(obj, name, value)
    except Exception:
        # WPS 和不同 Office 版本暴露的可选属性不完全一致。
        pass


def _safe_call(obj, method_name: str, *args, **kwargs):
    method = getattr(obj, method_name, None)
    if method is None:
        return None
    try:
        return method(*args, **kwargs)
    except Exception:
        return None


def _exception_text(exc: Exception) -> str:
    parts = [str(exc)]
    for arg in getattr(exc, "args", ()):
        if isinstance(arg, str):
            parts.append(arg)
        elif isinstance(arg, tuple):
            parts.extend(str(value) for value in arg if value)
    return " ".join(parts).lower()


def _translate_com_error(exc: Exception, filename: str, action: str) -> LegacyOfficeError:
    text = _exception_text(exc)
    if any(token in text for token in ("password", "密码", "口令")):
        return OfficePasswordProtectedError("文件需要密码，暂不支持处理")
    if any(token in text for token in ("protected", "protection", "保护")):
        return OfficeProtectedDocumentError("文件受到保护，无法修改")
    if any(
        token in text
        for token in (
            "read-only",
            "readonly",
            "permission denied",
            "sharing violation",
            "被占用",
            "只读",
        )
    ):
        return OfficeFileLockedError("文件被占用或只读，请关闭 Office/WPS 后重试")
    return LegacyOfficeError(f"{action}失败：{exc}")


def _actual_backend_name(kind: str, candidate_name: str, application) -> str:
    details = []
    for attribute in ("Path", "Caption", "Name"):
        try:
            details.append(str(getattr(application, attribute)))
        except Exception:
            pass
    signature = " ".join(details).lower()
    if "kingsoft" in signature or "wps" in signature:
        return {
            "word": "WPS Writer",
            "excel": "WPS Spreadsheets",
            "powerpoint": "WPS Presentation",
        }[kind]
    return candidate_name


class LegacyOfficeSession:
    """在同一个后台线程中复用并可靠关闭 Office/WPS 应用实例。"""

    def __init__(self, pythoncom_module=None, win32_client=None):
        self._pythoncom = pythoncom_module
        self._client = win32_client
        self._entered = False
        self._applications: Dict[str, object] = {}
        self._backend_names: Dict[str, str] = {}

    def __enter__(self) -> "LegacyOfficeSession":
        if self._pythoncom is None or self._client is None:
            self._pythoncom, self._client = _load_com_modules()
        self._pythoncom.CoInitialize()
        self._entered = True
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        self.close()

    def close(self) -> None:
        for kind in reversed(tuple(self._applications)):
            self.release(kind)
        gc.collect()
        if self._entered:
            self._entered = False
            self._pythoncom.CoUninitialize()

    def backend_name(self, kind: str) -> Optional[str]:
        return self._backend_names.get(kind)

    def release(self, kind: str) -> None:
        application = self._applications.pop(kind, None)
        self._backend_names.pop(kind, None)
        if application is not None:
            _safe_call(application, "Quit")
        gc.collect()

    def application(self, kind: str):
        if not self._entered:
            raise RuntimeError("LegacyOfficeSession 必须在 with 块中使用")
        if kind in self._applications:
            return self._applications[kind]

        # WPS 三个应用常共享同一套后台进程；切换组件前释放已有 WPS 根对象，
        # 避免批次结束时同时退出多个失效 COM 对象。
        for existing_kind, existing_backend in tuple(self._backend_names.items()):
            if existing_kind != kind and existing_backend.startswith("WPS"):
                self.release(existing_kind)

        failures = []
        for backend_name, prog_id in _APPLICATION_CANDIDATES[kind]:
            try:
                application = self._client.DispatchEx(prog_id)
            except Exception as exc:
                failures.append(f"{prog_id}: {exc}")
                continue

            self._configure_application(kind, application)
            actual_backend = _actual_backend_name(kind, backend_name, application)
            self._applications[kind] = application
            self._backend_names[kind] = actual_backend
            return application

        label = _APPLICATION_LABELS[kind]
        raise OfficeApplicationUnavailableError(
            f"处理该旧格式文件需要安装 {label}；也可以先另存为新版 Office 格式"
        )

    @staticmethod
    def _configure_application(kind: str, application) -> None:
        _safe_setattr(application, "Visible", False)
        _safe_setattr(application, "DisplayAlerts", 0 if kind != "excel" else False)
        _safe_setattr(application, "AutomationSecurity", 3)
        _safe_setattr(application, "EnableEvents", False)
        _safe_setattr(application, "AskToUpdateLinks", False)
        if kind == "word":
            options = getattr(application, "Options", None)
            if options is not None:
                _safe_setattr(options, "ConfirmConversions", False)
                _safe_setattr(options, "SaveNormalPrompt", False)


@contextlib.contextmanager
def _atomic_editable_copy(file_path: str, output_path: str) -> Iterator[str]:
    """在同目录同扩展名临时副本中修改，成功后原子替换最终文件。"""
    with atomic_output_path(output_path) as temporary_path:
        shutil.copy2(file_path, temporary_path)
        yield temporary_path
        if not os.path.isfile(temporary_path) or os.path.getsize(temporary_path) == 0:
            raise OfficeSaveError("Office 未生成有效的输出文件")


def _prepared_rules(rules: Sequence[Tuple[str, str]]):
    from string_replacer import _prepare_rules

    return _prepare_rules(list(rules))


def _replacement_regions(text: str, prepared_rules):
    from string_replacer import _find_rule_regions

    return _find_rule_regions(text, prepared_rules)


def _replace_in_word_range(word_range, prepared_rules) -> int:
    try:
        text = word_range.Text or ""
        base_start = int(word_range.Start)
    except Exception:
        return 0
    regions = _replacement_regions(text, prepared_rules)
    for start, end, new_text in reversed(regions):
        target = word_range.Duplicate
        target.SetRange(base_start + start, base_start + end)
        target.Text = new_text
    return len(regions)


def _iter_word_story_ranges(document) -> Iterator[object]:
    try:
        initial_ranges = list(document.StoryRanges)
    except Exception:
        initial_ranges = []

    for initial_range in initial_ranges:
        current = initial_range
        seen = set()
        while current is not None:
            try:
                identity = (int(current.StoryType), int(current.Start), int(current.End))
            except Exception:
                identity = id(current)
            if identity in seen:
                break
            seen.add(identity)
            yield current
            try:
                current = current.NextStoryRange
            except Exception:
                current = None


def _replace_in_word_document(document, rules: Sequence[Tuple[str, str]]) -> int:
    prepared = _prepared_rules(rules)
    total_count = 0
    for story_range in _iter_word_story_ranges(document):
        try:
            paragraphs = list(story_range.Paragraphs)
        except Exception:
            paragraphs = []
        if paragraphs:
            for paragraph in reversed(paragraphs):
                total_count += _replace_in_word_range(paragraph.Range, prepared)
        else:
            total_count += _replace_in_word_range(story_range, prepared)
    return total_count


def _open_word_document(application, path: str, *, read_only: bool = False):
    try:
        return application.Documents.Open(
            FileName=os.path.abspath(path),
            ConfirmConversions=False,
            ReadOnly=read_only,
            AddToRecentFiles=False,
            Revert=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
    except TypeError:
        # 部分 WPS 版本不接受完整的 Word 命名参数。
        return application.Documents.Open(os.path.abspath(path), False, read_only)


def replace_in_doc(
    file_path: str,
    rules: Sequence[Tuple[str, str]],
    output_path: str,
    session: LegacyOfficeSession,
) -> int:
    filename = os.path.basename(file_path)
    application = session.application("word")
    with _atomic_editable_copy(file_path, output_path) as temporary_path:
        document = None
        try:
            document = _open_word_document(application, temporary_path)
            original_track_revisions = None
            try:
                original_track_revisions = bool(document.TrackRevisions)
                if original_track_revisions:
                    document.TrackRevisions = False
            except Exception:
                original_track_revisions = None

            count = _replace_in_word_document(document, rules)

            if original_track_revisions is not None:
                try:
                    document.TrackRevisions = original_track_revisions
                except Exception:
                    pass
            document.Save()
            document.Close(SaveChanges=False)
            document = None
            return count
        except LegacyOfficeError:
            raise
        except Exception as exc:
            raise _translate_com_error(exc, filename, "Word 处理") from exc
        finally:
            if document is not None:
                _safe_call(document, "Close", SaveChanges=False)


def convert_doc_to_docx(
    file_path: str,
    output_path: str,
    session: LegacyOfficeSession,
) -> str:
    """把 .doc 转为指定临时 .docx，供现有 python-docx 提取器读取。"""
    filename = os.path.basename(file_path)
    application = session.application("word")
    document = None
    try:
        document = _open_word_document(application, file_path, read_only=True)
        # wdFormatXMLDocument = 12
        document.SaveAs2(FileName=os.path.abspath(output_path), FileFormat=12, AddToRecentFiles=False)
        document.Close(SaveChanges=False)
        document = None
    except Exception as exc:
        raise _translate_com_error(exc, filename, "Word 格式转换") from exc
    finally:
        if document is not None:
            _safe_call(document, "Close", SaveChanges=False)
    if not os.path.isfile(output_path) or os.path.getsize(output_path) == 0:
        raise OfficeSaveError("Word 未生成有效的 .docx 临时文件")
    return output_path


@contextlib.contextmanager
def temporary_docx_source(
    file_path: str,
    session: LegacyOfficeSession,
) -> Iterator[str]:
    """为 .doc 提供短生命周期 .docx 视图，供现有只读提取器复用。"""
    if os.path.splitext(file_path)[1].lower() == ".docx":
        yield file_path
        return
    with tempfile.TemporaryDirectory(prefix="replace-simple-doc-") as temporary_dir:
        stem = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(temporary_dir, f"{stem}.docx")
        yield convert_doc_to_docx(file_path, output_path, session)


def _replace_in_excel_workbook(workbook, rules: Sequence[Tuple[str, str]]) -> int:
    from string_replacer import _replace_text_with_rules, _workbook_cell_text

    prepared = _prepared_rules(rules)
    total_count = 0
    for worksheet in workbook.Worksheets:
        try:
            # xlCellTypeConstants = 2；没有常量单元格时会抛出 COM 异常。
            cells = worksheet.UsedRange.SpecialCells(2).Cells
        except Exception:
            continue
        for cell in cells:
            try:
                if bool(cell.HasFormula):
                    continue
                value = cell.Value
                if isinstance(value, (_datetime.date, _datetime.datetime, _datetime.time)):
                    continue
                text = _workbook_cell_text(value)
                if not text:
                    continue
                new_value, replaced_count = _replace_text_with_rules(text, prepared)
                if replaced_count:
                    cell.NumberFormat = "@"
                    cell.Value = new_value
                    total_count += replaced_count
            except LegacyOfficeError:
                raise
            except Exception as exc:
                raise OfficeProtectedDocumentError(
                    f"工作表 {getattr(worksheet, 'Name', '')} 无法修改：{exc}"
                ) from exc
    return total_count


def _open_excel_workbook(application, path: str):
    try:
        return application.Workbooks.Open(
            Filename=os.path.abspath(path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Notify=False,
        )
    except TypeError:
        return application.Workbooks.Open(os.path.abspath(path), 0, False)


def replace_in_xls(
    file_path: str,
    rules: Sequence[Tuple[str, str]],
    output_path: str,
    session: LegacyOfficeSession,
) -> int:
    filename = os.path.basename(file_path)
    application = session.application("excel")
    with _atomic_editable_copy(file_path, output_path) as temporary_path:
        workbook = None
        try:
            workbook = _open_excel_workbook(application, temporary_path)
            count = _replace_in_excel_workbook(workbook, rules)
            workbook.Save()
            workbook.Close(SaveChanges=False)
            workbook = None
            return count
        except LegacyOfficeError:
            raise
        except Exception as exc:
            raise _translate_com_error(exc, filename, "Excel 处理") from exc
        finally:
            if workbook is not None:
                _safe_call(workbook, "Close", SaveChanges=False)


def _replace_in_powerpoint_text_range_raw(text_range, prepared_rules) -> int:
    try:
        text = text_range.Text or ""
    except Exception:
        return 0
    regions = _replacement_regions(text, prepared_rules)
    for start, end, new_text in reversed(regions):
        target = text_range.Characters(start + 1, end - start)
        target.Text = new_text
    return len(regions)


def _replace_in_powerpoint_text_range(text_range, prepared_rules) -> int:
    """按段落替换，保持与 python-pptx 现有处理边界一致。"""
    try:
        paragraphs = text_range.Paragraphs()
        count = int(paragraphs.Count)
    except Exception:
        return _replace_in_powerpoint_text_range_raw(text_range, prepared_rules)
    if count <= 1:
        return _replace_in_powerpoint_text_range_raw(text_range, prepared_rules)

    total_count = 0
    for index in range(count, 0, -1):
        try:
            paragraph_range = text_range.Paragraphs(index, 1)
        except Exception:
            return _replace_in_powerpoint_text_range_raw(text_range, prepared_rules)
        total_count += _replace_in_powerpoint_text_range_raw(paragraph_range, prepared_rules)
    return total_count


def _replace_in_powerpoint_shape(shape, prepared_rules) -> int:
    total_count = 0
    try:
        has_table = bool(shape.HasTable)
    except Exception:
        has_table = False
    if has_table:
        table = shape.Table
        for row_index in range(1, int(table.Rows.Count) + 1):
            for column_index in range(1, int(table.Columns.Count) + 1):
                text_range = table.Cell(row_index, column_index).Shape.TextFrame.TextRange
                total_count += _replace_in_powerpoint_text_range(text_range, prepared_rules)
        return total_count

    try:
        has_text = bool(shape.HasTextFrame) and bool(shape.TextFrame.HasText)
    except Exception:
        has_text = False
    if has_text:
        total_count += _replace_in_powerpoint_text_range(shape.TextFrame.TextRange, prepared_rules)

    try:
        # msoGroup = 6
        is_group = int(shape.Type) == 6
    except Exception:
        is_group = False
    if is_group:
        group_items = shape.GroupItems
        for index in range(1, int(group_items.Count) + 1):
            total_count += _replace_in_powerpoint_shape(group_items.Item(index), prepared_rules)
    return total_count


def _replace_in_powerpoint_presentation(presentation, rules: Sequence[Tuple[str, str]]) -> int:
    prepared = _prepared_rules(rules)
    total_count = 0
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            total_count += _replace_in_powerpoint_shape(shape, prepared)
        try:
            notes_shapes = slide.NotesPage.Shapes
        except Exception:
            notes_shapes = ()
        for shape in notes_shapes:
            total_count += _replace_in_powerpoint_shape(shape, prepared)
    return total_count


def _open_powerpoint_presentation(application, path: str):
    try:
        return application.Presentations.Open(
            FileName=os.path.abspath(path),
            ReadOnly=False,
            Untitled=False,
            WithWindow=False,
        )
    except TypeError:
        return application.Presentations.Open(os.path.abspath(path), False, False, False)


def replace_in_ppt(
    file_path: str,
    rules: Sequence[Tuple[str, str]],
    output_path: str,
    session: LegacyOfficeSession,
) -> int:
    filename = os.path.basename(file_path)
    application = session.application("powerpoint")
    with _atomic_editable_copy(file_path, output_path) as temporary_path:
        presentation = None
        try:
            presentation = _open_powerpoint_presentation(application, temporary_path)
            count = _replace_in_powerpoint_presentation(presentation, rules)
            presentation.Save()
            presentation.Close()
            presentation = None
            return count
        except LegacyOfficeError:
            raise
        except Exception as exc:
            raise _translate_com_error(exc, filename, "PowerPoint 处理") from exc
        finally:
            if presentation is not None:
                _safe_call(presentation, "Close")
