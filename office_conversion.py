"""旧格式转 OOXML：优先本机 WPS/Office，不可用时回退到 LibreOffice。"""

import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import zipfile
from contextlib import contextmanager

from legacy_office import (
    LegacyOfficeSession, LegacyOfficeError, OfficeApplicationUnavailableError,
    _open_excel_workbook, _open_powerpoint_presentation, _safe_call,
    convert_doc_to_docx,
)

MODERN_EXTENSIONS = {".doc": ".docx", ".xls": ".xlsx", ".ppt": ".pptx"}


def _find_libreoffice():
    # 分发包优先：不要求接收者安装软件，也不依赖 PATH 或注册表。
    app_directory = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
    bundled = app_directory / "converter" / "LibreOffice" / "program"
    for name in ("soffice.com", "soffice.exe"):
        if (bundled / name).is_file():
            return str(bundled / name)
    executable = shutil.which("soffice")
    if executable:
        return executable
    for variable in ("PROGRAMFILES", "PROGRAMFILES(X86)"):
        root = os.environ.get(variable)
        if root:
            candidate = Path(root) / "LibreOffice/program/soffice.exe"
            if candidate.is_file():
                return str(candidate)
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidate = Path(local) / "Programs/LibreOffice/program/soffice.exe"
        if candidate.is_file():
            return str(candidate)
    return None


def _convert_with_office(source, destination, extension):
    with LegacyOfficeSession() as session:
        if extension == ".doc":
            convert_doc_to_docx(source, destination, session)
            return
        document = None
        try:
            if extension == ".xls":
                document = _open_excel_workbook(session.application("excel"), source)
                document.SaveAs(os.path.abspath(destination), 51)
            else:
                document = _open_powerpoint_presentation(session.application("powerpoint"), source)
                document.SaveAs(os.path.abspath(destination), 24)
        finally:
            if document is not None:
                if extension == ".xls":
                    _safe_call(document, "Close", SaveChanges=False)
                else:
                    _safe_call(document, "Close")
                # 在 COM 线程退出前释放文档代理，避免 WPS 关闭后的迟延释放。
                document = None


@contextmanager
def converted_office_source(file_path):
    """只转换隔离副本；校验成功后交给替换器，退出时清理全部临时文件。"""
    extension = Path(file_path).suffix.lower()
    target_extension = MODERN_EXTENSIONS[extension]
    if os.path.getsize(file_path) == 0:
        raise LegacyOfficeError("旧格式文件为空，无法转换")
    with tempfile.TemporaryDirectory(prefix="replace-simple-convert-") as directory:
        root = Path(directory)
        source = root / ("source" + extension)
        shutil.copy2(file_path, source)
        destination = root / ("source" + target_extension)
        try:
            _convert_with_office(str(source), str(destination), extension)
            executable = None
        except OfficeApplicationUnavailableError as exc:
            executable = _find_libreoffice()
            if not executable:
                raise OfficeApplicationUnavailableError(
                    "无法调用 WPS/Office 转换组件。请修复 WPS 的自动化组件，"
                    "或手动另存为新版格式后重试；也可使用本机 LibreOffice。"
                ) from exc
        if executable:
            try:
                completed = subprocess.run(
                    [executable, "-env:UserInstallation=" + (root / "profile").as_uri(),
                     "--headless", "--convert-to", target_extension[1:],
                     "--outdir", str(root), str(source)],
                    capture_output=True, timeout=120,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            except subprocess.TimeoutExpired as exc:
                raise LegacyOfficeError("格式转换超时，请检查文件后重试") from exc
            if completed.returncode:
                raise LegacyOfficeError("LibreOffice 格式转换失败，请检查文件是否损坏或需要密码")
        member = {".docx": "word/document.xml", ".xlsx": "xl/workbook.xml",
                  ".pptx": "ppt/presentation.xml"}[target_extension]
        try:
            with zipfile.ZipFile(destination) as archive:
                if member not in archive.namelist() or archive.testzip() is not None:
                    raise ValueError("invalid OOXML")
        except (OSError, zipfile.BadZipFile, ValueError) as exc:
            raise LegacyOfficeError("格式转换未生成有效文件，原文件已保留") from exc
        yield str(destination)
