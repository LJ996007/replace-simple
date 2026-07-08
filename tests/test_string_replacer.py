import sys
import tkinter as tk
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import ttk

import xlwt
from docx import Document
from openpyxl import Workbook, load_workbook
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import main as simple_main
from main import ReplaceSimpleApp, normalize_rule_rows, remove_matching_file_paths
from string_replacer import (
    apply_rules_to_filename,
    batch_replace,
    get_output_path,
    load_replacement_rules,
    replace_in_docx,
    replace_in_pptx,
)


class SimpleReplacementTests(unittest.TestCase):
    def test_load_replacement_rules_reads_two_columns_from_first_row(self):
        with TemporaryDirectory() as tmp_dir:
            rules_path = Path(tmp_dir) / "rules.xlsx"
            wb = Workbook()
            ws = wb.active
            ws["A1"] = "采购人"
            ws["B1"] = "建设单位"
            ws["A2"] = "供应商"
            ws["B2"] = "投标人"
            wb.save(rules_path)

            rules = load_replacement_rules(str(rules_path))

        self.assertEqual(rules, [("采购人", "建设单位"), ("供应商", "投标人")])

    def test_load_replacement_rules_supports_xlsm_files(self):
        with TemporaryDirectory() as tmp_dir:
            rules_path = Path(tmp_dir) / "rules.xlsm"
            wb = Workbook()
            ws = wb.active
            ws["A1"] = "采购人"
            ws["B1"] = "建设单位"
            wb.save(rules_path)

            rules = load_replacement_rules(str(rules_path))

        self.assertEqual(rules, [("采购人", "建设单位")])

    def test_load_replacement_rules_supports_xls_files(self):
        with TemporaryDirectory() as tmp_dir:
            rules_path = Path(tmp_dir) / "rules.xls"
            wb = xlwt.Workbook()
            ws = wb.add_sheet("Sheet1")
            ws.write(0, 0, "采购人")
            ws.write(0, 1, "建设单位")
            ws.write(1, 0, "供应商")
            ws.write(1, 1, "投标人")
            wb.save(str(rules_path))

            rules = load_replacement_rules(str(rules_path))

        self.assertEqual(rules, [("采购人", "建设单位"), ("供应商", "投标人")])

    def test_normalize_rule_rows_skips_empty_old_text_and_dedupes_exact_pairs(self):
        rules = normalize_rule_rows([
            (" 采购人 ", " 建设单位 "),
            ("", "忽略"),
            ("供应商", ""),
            ("采购人", "建设单位"),
            ("采购人", "采购单位"),
        ])

        self.assertEqual(rules, [
            ("采购人", "建设单位"),
            ("供应商", ""),
            ("采购人", "采购单位"),
        ])

    def test_app_starts_with_compact_fonts_and_blank_editable_rows(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = ReplaceSimpleApp(root, restore_session=False)
            root.update_idletasks()
            data = app.rules_sheet.get_sheet_data()

            self.assertLessEqual(app.body_font.cget("size"), 10)
            self.assertLessEqual(app.title_font.cget("size"), 13)
            self.assertGreaterEqual(len(data), 3)
            self.assertEqual(list(data[0]), ["", ""])
        finally:
            root.destroy()

    def test_app_uses_flat_sections_without_label_frames(self):
        root = tk.Tk()
        root.withdraw()
        try:
            ReplaceSimpleApp(root, restore_session=False)

            def walk(widget):
                children = widget.winfo_children()
                for child in children:
                    yield child
                    yield from walk(child)

            label_frames = [child for child in walk(root) if isinstance(child, ttk.LabelFrame)]

            self.assertEqual(label_frames, [])
        finally:
            root.destroy()

    def test_app_does_not_enable_sv_ttk_theme_on_startup(self):
        calls = []

        class ThemeProbe:
            def set_theme(self, theme):
                calls.append(theme)

        had_sv_ttk = hasattr(simple_main, "sv_ttk")
        previous_sv_ttk = getattr(simple_main, "sv_ttk", None)
        simple_main.sv_ttk = ThemeProbe()

        root = tk.Tk()
        root.withdraw()
        try:
            ReplaceSimpleApp(root, restore_session=False)
            self.assertEqual(calls, [])
        finally:
            root.destroy()
            if had_sv_ttk:
                simple_main.sv_ttk = previous_sv_ttk
            else:
                delattr(simple_main, "sv_ttk")

    def test_make_blank_rows_returns_independent_rows(self):
        from main import make_blank_rows

        rows = make_blank_rows(4)
        self.assertEqual(rows, [["", ""], ["", ""], ["", ""], ["", ""]])
        # 各行必须是独立对象：改第 0 行不应连带改写其它行
        rows[0][0] = "X"
        self.assertEqual([r[0] for r in rows], ["X", "", "", ""])
        self.assertEqual(make_blank_rows(0), [])

    def test_remove_matching_file_paths_removes_selected_rules_workbook(self):
        with TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            target = tmp_path / "rules.xlsx"
            target.touch()
            other = tmp_path / "document.xlsx"
            other.touch()

            remaining, removed = remove_matching_file_paths(
                [str(other), str(target)],
                str(target),
            )

        self.assertEqual(remaining, [str(other)])
        self.assertEqual(removed, [str(target)])

    def test_remove_matching_file_paths_handles_equivalent_relative_paths(self):
        with TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            target = tmp_path / "rules.xlsx"
            target.touch()
            other = tmp_path / "document.xlsx"
            other.touch()
            relative_target = target.parent / "." / target.name

            remaining, removed = remove_matching_file_paths(
                [str(other), str(relative_target)],
                str(target),
            )

        self.assertEqual(remaining, [str(other)])
        self.assertEqual(removed, [str(relative_target)])

    def test_get_output_path_replaces_filename_without_suffix(self):
        output_path = get_output_path(
            r"C:\work\采购人.xlsx",
            [("采购人", "建设单位")],
            None,
        )

        self.assertEqual(Path(output_path).name, "建设单位.xlsx")

    def test_get_output_path_adds_suffix_when_filename_would_not_change(self):
        output_path = get_output_path(
            r"C:\work\sample.xlsx",
            [("采购人", "建设单位")],
            None,
        )

        self.assertEqual(Path(output_path).name, "sample_已替换.xlsx")

    def test_get_output_path_adds_suffix_when_selected_output_dir_is_source_dir(self):
        output_path = get_output_path(
            r"C:\work\sample.xlsx",
            [("采购人", "建设单位")],
            r"C:\work",
        )

        self.assertEqual(Path(output_path).name, "sample_已替换.xlsx")

    def test_filename_replacement_preserves_extension(self):
        filename = apply_rules_to_filename(
            "投标文件.docx",
            [("doc", "文档"), ("投标", "响应")],
        )

        self.assertEqual(filename, "响应文件.docx")

    def test_rules_are_applied_longest_old_text_first(self):
        filename = apply_rules_to_filename(
            "采购人名单.xlsx",
            [("采购", "购买"), ("采购人", "建设单位")],
        )

        self.assertEqual(filename, "建设单位名单.xlsx")

    def test_filename_replacement_does_not_repeat_when_new_text_contains_old_text(self):
        filename = apply_rules_to_filename(
            "北京航空航天大学安全保卫部学院路校区中控室大屏更换采购.xlsx",
            [("北京航空航天大学安全保卫部学院路校区中控室大屏更换", "北京航空航天大学安全保卫部学院路校区中控室大屏更换采购")],
        )

        self.assertEqual(filename, "北京航空航天大学安全保卫部学院路校区中控室大屏更换采购.xlsx")

    def test_batch_replace_writes_replaced_xlsx_in_original_dir(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "采购人.xlsx"
            wb = Workbook()
            ws = wb.active
            ws["A1"] = "采购人名单"
            wb.save(source_path)

            results, error = batch_replace(
                [str(source_path)],
                [("采购人", "建设单位")],
                output_dir=None,
            )

            output_path = Path(tmp_dir) / "建设单位.xlsx"
            output_exists = output_path.exists()
            output_wb = load_workbook(output_path)
            try:
                value = output_wb.active["A1"].value
            finally:
                output_wb.close()

        self.assertIsNone(error)
        self.assertEqual(results, {"采购人.xlsx": 1})
        self.assertTrue(output_exists)
        self.assertEqual(value, "建设单位名单")

    def test_batch_replace_does_not_overwrite_source_when_filename_does_not_change(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "sample.xlsx"
            wb = Workbook()
            ws = wb.active
            ws["A1"] = "采购人名单"
            wb.save(source_path)

            results, error = batch_replace([str(source_path)], [("采购人", "建设单位")], output_dir=None)

            safe_output_path = Path(tmp_dir) / "sample_已替换.xlsx"
            source_wb = load_workbook(source_path)
            output_wb = load_workbook(safe_output_path)
            try:
                source_value = source_wb.active["A1"].value
                output_value = output_wb.active["A1"].value
            finally:
                source_wb.close()
                output_wb.close()

        self.assertIsNone(error)
        self.assertEqual(results, {"sample.xlsx": 1})
        self.assertEqual(source_value, "采购人名单")
        self.assertEqual(output_value, "建设单位名单")

    def test_docx_cross_run_replacement_preserves_unmatched_run_formatting(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "source.docx"
            output_path = Path(tmp_dir) / "output.docx"

            doc = Document()
            paragraph = doc.add_paragraph()
            first = paragraph.add_run("采购")
            first.bold = True
            second = paragraph.add_run("人")
            second.italic = True
            tail = paragraph.add_run("名单")
            tail.underline = True
            doc.save(source_path)

            count = replace_in_docx(str(source_path), [("采购人", "建设单位")], str(output_path))

            output_doc = Document(output_path)
            runs = output_doc.paragraphs[0].runs

        self.assertEqual(count, 1)
        self.assertEqual([run.text for run in runs], ["建设单位", "", "名单"])
        self.assertTrue(runs[0].bold)
        self.assertTrue(runs[2].underline)

    def test_docx_does_not_repeat_when_new_text_contains_old_text(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "source.docx"
            output_path = Path(tmp_dir) / "output.docx"

            doc = Document()
            doc.add_paragraph("北京航空航天大学安全保卫部学院路校区中控室大屏更换采购")
            doc.save(source_path)

            count = replace_in_docx(
                str(source_path),
                [
                    (
                        "北京航空航天大学安全保卫部学院路校区中控室大屏更换",
                        "北京航空航天大学安全保卫部学院路校区中控室大屏更换采购",
                    )
                ],
                str(output_path),
            )

            output_doc = Document(output_path)
            text = output_doc.paragraphs[0].text

        self.assertEqual(count, 0)
        self.assertEqual(text, "北京航空航天大学安全保卫部学院路校区中控室大屏更换采购")

    def test_pptx_cross_run_replacement_preserves_unmatched_run_formatting(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "source.pptx"
            output_path = Path(tmp_dir) / "output.pptx"

            presentation = Presentation()
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])
            textbox = slide.shapes.add_textbox(0, 0, 3000000, 1000000)
            paragraph = textbox.text_frame.paragraphs[0]
            first = paragraph.add_run()
            first.text = "供"
            first.font.bold = True
            second = paragraph.add_run()
            second.text = "应商"
            second.font.italic = True
            tail = paragraph.add_run()
            tail.text = "名单"
            tail.font.underline = True
            presentation.save(source_path)

            count = replace_in_pptx(str(source_path), [("供应商", "投标人")], str(output_path))

            output_presentation = Presentation(output_path)
            output_shape = output_presentation.slides[0].shapes[0]
            runs = output_shape.text_frame.paragraphs[0].runs

        self.assertEqual(count, 1)
        self.assertEqual([run.text for run in runs], ["投标人", "", "名单"])
        self.assertTrue(runs[0].font.bold)
        self.assertTrue(runs[2].font.underline)


if __name__ == "__main__":
    unittest.main()
