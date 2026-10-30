import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from unittest.mock import MagicMock

from docx import Document
from openpyxl import Workbook, load_workbook
from pptx import Presentation

import office_conversion as conversion
from string_replacer import batch_replace, get_output_path


def make_document(source, destination, extension):
    if extension == ".doc":
        document = Document()
        document.add_paragraph("OLD")
    elif extension == ".xls":
        document = Workbook()
        document.active["A1"] = "OLD"
        document.active["B1"] = "=1+2"
    else:
        document = Presentation()
        document.slides.add_slide(document.slide_layouts[1]).shapes.title.text = "OLD"
    document.save(destination)


class ConversionTests(unittest.TestCase):
    def test_preview_uses_same_modern_filename_as_batch(self):
        for old, new in conversion.MODERN_EXTENSIONS.items():
            self.assertTrue(get_output_path("OLD" + old, [("OLD", "NEW")]).endswith("NEW" + new))

    def test_office_fallback_saves_modern_format_and_closes_documents(self):
        for extension, opener, format_code in (
            (".xls", "_open_excel_workbook", 51),
            (".ppt", "_open_powerpoint_presentation", 24),
        ):
            with self.subTest(extension=extension):
                document = MagicMock()
                destination = str(Path("converted" + conversion.MODERN_EXTENSIONS[extension]).resolve())
                with patch.object(conversion, "LegacyOfficeSession"), patch.object(
                    conversion, opener, return_value=document
                ):
                    conversion._convert_with_office("source" + extension, destination, extension)
                document.SaveAs.assert_called_once_with(destination, format_code)
                document.Close.assert_called_once()

    def test_packaged_engine_wins_without_system_installation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            executable = root / "converter/LibreOffice/program/soffice.exe"
            executable.parent.mkdir(parents=True)
            executable.touch()
            with patch.object(conversion.sys, "frozen", True, create=True), patch.object(
                conversion.sys, "executable", str(root / "replace-simple.exe")
            ), patch.object(conversion.shutil, "which", side_effect=AssertionError("No system lookup needed")):
                self.assertEqual(conversion._find_libreoffice(), str(executable))

    def test_all_formats_convert_then_replace_without_overwriting_existing_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for old, new in conversion.MODERN_EXTENSIONS.items():
                with self.subTest(extension=old):
                    source = root / ("OLD" + old.upper())
                    source.write_bytes(b"legacy original")
                    existing = root / ("NEW" + new)
                    existing.write_bytes(b"existing output")
                    details, progress = [], []
                    with patch.object(conversion, "_find_libreoffice", side_effect=AssertionError("WPS takes priority")), patch.object(
                        conversion, "_convert_with_office", side_effect=make_document
                    ) as convert:
                        results, error = batch_replace([str(source)], [("OLD", "NEW")],
                            result_callback=details.append, progress_callback=lambda *args: progress.append(args))
                    self.assertIsNone(error)
                    self.assertEqual(list(results.values()), [1])
                    output = root / ("NEW_1" + new)
                    self.assertEqual(details[0]["output_path"], str(output))
                    self.assertEqual(details[0]["converted_from"], old)
                    self.assertTrue(any("正在转换" in item[2] for item in progress))
                    self.assertEqual(source.read_bytes(), b"legacy original")
                    self.assertEqual(existing.read_bytes(), b"existing output")
                    self.assertFalse(Path(convert.call_args.args[0]).exists())
                    if old == ".doc":
                        self.assertEqual(Document(output).paragraphs[0].text, "NEW")
                    elif old == ".xls":
                        workbook = load_workbook(output)
                        self.assertEqual(workbook.active["A1"].value, "NEW")
                        self.assertEqual(workbook.active["B1"].value, "=1+2")
                        workbook.close()
                    else:
                        self.assertEqual(Presentation(output).slides[0].shapes.title.text, "NEW")

    def test_invalid_conversion_and_empty_files_fail_but_batch_continues(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            broken = root / "broken.doc"
            broken.write_bytes(b"legacy")
            empty = root / "empty.xls"
            empty.touch()
            modern = root / "valid.docx"
            make_document(None, modern, ".doc")
            with patch.object(conversion, "_find_libreoffice", return_value=None), patch.object(
                conversion, "_convert_with_office", side_effect=lambda s, d, e: Path(d).write_bytes(b"bad")
            ):
                results, error = batch_replace([str(broken), str(empty), str(modern)], [("OLD", "NEW")])
            self.assertEqual(results, {"valid.docx": 1})
            self.assertIn("格式转换未生成有效文件", error)
            self.assertIn("旧格式文件为空", error)
            self.assertFalse((root / "broken.docx").exists())
            self.assertFalse((root / "empty.xlsx").exists())
            self.assertEqual(broken.read_bytes(), b"legacy")

    def test_libreoffice_fallback_uses_isolated_profile(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.ppt"
            source.write_bytes(b"legacy")
            def run(command, **kwargs):
                input_path = Path(command[-1])
                make_document(None, input_path.with_suffix(".pptx"), ".ppt")
                self.assertIn("--headless", command)
                self.assertTrue(command[1].startswith("-env:UserInstallation=file:"))
                self.assertEqual(kwargs["timeout"], 120)
                from subprocess import CompletedProcess
                return CompletedProcess(command, 0)
            with patch.object(conversion, "_find_libreoffice", return_value="soffice"), patch.object(
                conversion.subprocess, "run", side_effect=run
            ), patch.object(conversion, "_convert_with_office", side_effect=conversion.OfficeApplicationUnavailableError("missing")) as com:
                with conversion.converted_office_source(str(source)) as output:
                    self.assertTrue(Path(output).exists())
                self.assertFalse(Path(output).exists())
                com.assert_called_once()

    def test_missing_engine_has_actionable_error(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.doc"
            source.write_bytes(b"legacy")
            with patch.object(conversion, "_find_libreoffice", return_value=None), patch.object(
                conversion, "_convert_with_office", side_effect=conversion.OfficeApplicationUnavailableError("missing")
            ):
                with self.assertRaisesRegex(conversion.OfficeApplicationUnavailableError, "LibreOffice"):
                    with conversion.converted_office_source(str(source)):
                        self.fail("Conversion must fail")
