import sys
import time
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
from tender_info_extractor import extract_project_info, extract_project_info_rules
from word_table_exporter import (
    batch_export_word_tables,
    export_word_tables_to_excel,
    get_table_export_output_path,
    scan_word_tables,
)


def create_hidden_root():
    last_error = None
    for _attempt in range(3):
        try:
            root = tk.Tk()
            root.withdraw()
            return root
        except tk.TclError as exc:
            last_error = exc
            time.sleep(0.2)
    raise last_error


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
        root = create_hidden_root()
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
        root = create_hidden_root()
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

        root = create_hidden_root()
        try:
            ReplaceSimpleApp(root, restore_session=False)
            self.assertEqual(calls, [])
        finally:
            root.destroy()
            if had_sv_ttk:
                simple_main.sv_ttk = previous_sv_ttk
            else:
                delattr(simple_main, "sv_ttk")

    def test_app_has_word_table_export_entry_and_opens_window(self):
        root = create_hidden_root()
        app = None
        try:
            app = ReplaceSimpleApp(root, restore_session=False)
            root.update_idletasks()

            self.assertEqual(app.table_export_button.cget("text"), "提取 Word 表格")

            app.open_word_table_exporter()
            root.update_idletasks()

            self.assertIsNotNone(app.table_export_window)
            self.assertTrue(app.table_export_window.exists())
            self.assertEqual(app.table_export_window.window.title(), "提取 Word 表格到 Excel")
            self.assertTrue(hasattr(app.table_export_window, "scan_tree"))
            self.assertEqual(app.table_export_window.scan_button.cget("text"), "扫描表格")

            app.table_export_window.close()
            root.update_idletasks()
            self.assertIsNone(app.table_export_window)
        finally:
            if app is not None and app.table_export_window is not None:
                app.table_export_window.close()
            root.destroy()

    def test_app_has_tender_file_import_button(self):
        root = create_hidden_root()
        try:
            ReplaceSimpleApp(root, restore_session=False)

            def walk(widget):
                children = widget.winfo_children()
                for child in children:
                    yield child
                    yield from walk(child)

            button_texts = [
                child.cget("text")
                for child in walk(root)
                if isinstance(child, ttk.Button)
            ]

            self.assertIn("导入招标文件", button_texts)
        finally:
            root.destroy()

    def test_tender_import_reuses_remembered_word_file_without_dialog(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "tender.docx"
            doc = Document()
            doc.add_paragraph("项目名称：一次选择项目")
            doc.save(source_path)

            root = create_hidden_root()
            previous_dialog = simple_main.filedialog.askopenfilename
            try:
                app = ReplaceSimpleApp(root, restore_session=False)
                app._remember_word_files([str(source_path)])

                def fail_dialog(*_args, **_kwargs):
                    raise AssertionError("不应重复弹出文件选择窗口")

                simple_main.filedialog.askopenfilename = fail_dialog
                app.import_rules_from_tender_file()

                self.assertIn(("{项目名称}", "一次选择项目"), app.get_rules_from_table())
            finally:
                simple_main.filedialog.askopenfilename = previous_dialog
                root.destroy()

    def test_word_table_exporter_reuses_remembered_word_file(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "tender.docx"
            doc = Document()
            doc.add_paragraph("项目名称：表格复用项目")
            doc.save(source_path)

            root = create_hidden_root()
            app = None
            try:
                app = ReplaceSimpleApp(root, restore_session=False)
                app._remember_word_files([str(source_path)])

                app.open_word_table_exporter()
                root.update_idletasks()

                self.assertIsNotNone(app.table_export_window)
                self.assertEqual(app.table_export_window.file_paths, [str(source_path)])
            finally:
                if app is not None and app.table_export_window is not None:
                    app.table_export_window.close()
                root.destroy()

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

    def test_extract_project_info_rules_from_tender_docx(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "tender.docx"

            doc = Document()
            doc.add_paragraph("项目名称：智慧校园设备采购项目")
            doc.add_paragraph("项目编号：ABC-2026-001")
            table = doc.add_table(rows=3, cols=2)
            table.cell(0, 0).text = "采购人"
            table.cell(0, 1).text = "示例大学"
            table.cell(1, 0).text = "预算金额"
            table.cell(1, 1).text = "120万元"
            table.cell(2, 0).text = "开标地点"
            table.cell(2, 1).text = "北京市海淀区1号会议室"
            doc.save(source_path)

            info = extract_project_info(str(source_path))
            rules = extract_project_info_rules(str(source_path))

        self.assertEqual(info["项目名称"], "智慧校园设备采购项目")
        self.assertEqual(info["采购人"], "示例大学")
        self.assertIn(("{项目名称}", "智慧校园设备采购项目"), rules)
        self.assertIn(("{项目编号}", "ABC-2026-001"), rules)
        self.assertIn(("{预算金额}", "120万元"), rules)

    def test_extract_project_info_handles_procurement_contact_context(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "notice.docx"

            doc = Document()
            doc.add_paragraph("1.采购人信息")
            doc.add_paragraph("名 称：示例采购单位")
            doc.add_paragraph("2.采购代理机构信息")
            doc.add_paragraph("名称：示例代理机构")
            doc.save(source_path)

            info = extract_project_info(str(source_path))

        self.assertEqual(info["采购人"], "示例采购单位")
        self.assertEqual(info["采购代理机构"], "示例代理机构")

    def test_extract_project_info_handles_table_context_and_parenthesized_labels(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "table-context.docx"

            doc = Document()
            table = doc.add_table(rows=4, cols=4)
            table.cell(0, 0).text = "采购人信息"
            table.cell(1, 0).text = "名称"
            table.cell(1, 1).text = "表格采购单位"
            table.cell(2, 0).text = "项目名称"
            table.cell(2, 1).text = "设备采购"
            table.cell(2, 2).text = "项目编号"
            table.cell(2, 3).text = "XYZ-2026"
            table.cell(3, 0).text = "预算金额（万元）"
            table.cell(3, 1).text = "88"
            doc.save(source_path)

            info = extract_project_info(str(source_path))

        self.assertEqual(info["采购人"], "表格采购单位")
        self.assertEqual(info["项目名称"], "设备采购")
        self.assertEqual(info["项目编号"], "XYZ-2026")
        self.assertEqual(info["预算金额"], "88")

    def test_extract_project_info_handles_parenthesized_paragraph_labels(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "parenthesized-paragraph.docx"

            doc = Document()
            doc.add_paragraph("（一）预算金额（万元）：88")
            doc.add_paragraph("2.最高限价（如有）：90万元")
            doc.save(source_path)

            info = extract_project_info(str(source_path))

        self.assertEqual(info["预算金额"], "88")
        self.assertEqual(info["最高限价"], "90万元")

    def test_extract_project_info_does_not_cross_into_next_field_value(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "missing-value.docx"

            doc = Document()
            table = doc.add_table(rows=1, cols=3)
            table.cell(0, 0).text = "项目名称"
            table.cell(0, 1).text = "项目编号"
            table.cell(0, 2).text = "ABC-001"
            doc.save(source_path)

            info = extract_project_info(str(source_path))

        self.assertNotIn("项目名称", info)
        self.assertEqual(info["项目编号"], "ABC-001")

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

    def test_export_word_tables_writes_each_table_to_one_sheet(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "tables.docx"
            output_path = Path(tmp_dir) / "tables.xlsx"

            doc = Document()
            table1 = doc.add_table(rows=1, cols=2)
            table1.cell(0, 0).text = "项目"
            table1.cell(0, 1).text = "预算"
            doc.add_paragraph("between")
            table2 = doc.add_table(rows=1, cols=1)
            table2.cell(0, 0).text = "第二张表"
            doc.save(source_path)

            count = export_word_tables_to_excel(str(source_path), str(output_path))
            workbook = load_workbook(output_path)
            try:
                sheet_names = workbook.sheetnames
                first_value = workbook["表格1"]["A1"].value
                second_value = workbook["表格2"]["A1"].value
            finally:
                workbook.close()

        self.assertEqual(count, 2)
        self.assertEqual(sheet_names, ["表格1", "表格2"])
        self.assertEqual(first_value, "项目")
        self.assertEqual(second_value, "第二张表")

    def test_scan_word_tables_reports_section_context_preview_and_hint(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "tender.docx"

            doc = Document()
            heading = doc.add_paragraph("第五章 评审办法")
            heading.style = "Heading 1"
            doc.add_paragraph("三、商务技术评审")
            doc.add_paragraph("下表为评分标准。")
            table = doc.add_table(rows=2, cols=3)
            table.cell(0, 0).text = "评审因素"
            table.cell(0, 1).text = "分值"
            table.cell(0, 2).text = "评分标准"
            table.cell(1, 0).text = "技术方案"
            table.cell(1, 1).text = "30分"
            table.cell(1, 2).text = "按方案完整性评分"
            doc.save(source_path)

            items = scan_word_tables(str(source_path))

        self.assertEqual(len(items), 1)
        self.assertEqual(items[0].table_index, 1)
        self.assertIn("第五章 评审办法", items[0].section)
        self.assertIn("商务技术评审", items[0].context)
        self.assertIn("评审因素", items[0].preview)
        self.assertIn("评分/评审表", items[0].hint)
        self.assertEqual(items[0].row_count, 2)
        self.assertEqual(items[0].column_count, 3)

    def test_export_word_tables_can_export_selected_table_indexes_only(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "tables.docx"
            output_path = Path(tmp_dir) / "selected.xlsx"

            doc = Document()
            first = doc.add_table(rows=1, cols=1)
            first.cell(0, 0).text = "不要导出"
            second = doc.add_table(rows=1, cols=1)
            second.cell(0, 0).text = "需要导出"
            doc.save(source_path)

            count = export_word_tables_to_excel(
                str(source_path),
                str(output_path),
                table_indexes=[2],
            )
            workbook = load_workbook(output_path)
            try:
                sheet_names = workbook.sheetnames
                value = workbook["表格1"]["A1"].value
            finally:
                workbook.close()

        self.assertEqual(count, 1)
        self.assertEqual(sheet_names, ["表格1"])
        self.assertEqual(value, "需要导出")

    def test_export_word_tables_preserves_horizontal_merge(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "horizontal.docx"
            output_path = Path(tmp_dir) / "horizontal.xlsx"

            doc = Document()
            table = doc.add_table(rows=2, cols=3)
            merged = table.cell(0, 0).merge(table.cell(0, 1))
            merged.text = "横向合并"
            table.cell(0, 2).text = "普通"
            table.cell(1, 0).text = "A"
            table.cell(1, 1).text = "B"
            table.cell(1, 2).text = "C"
            doc.save(source_path)

            export_word_tables_to_excel(str(source_path), str(output_path))
            workbook = load_workbook(output_path)
            try:
                worksheet = workbook["表格1"]
                ranges = {str(cell_range) for cell_range in worksheet.merged_cells.ranges}
                value = worksheet["A1"].value
                tail = worksheet["C1"].value
            finally:
                workbook.close()

        self.assertIn("A1:B1", ranges)
        self.assertEqual(value, "横向合并")
        self.assertEqual(tail, "普通")

    def test_export_word_tables_preserves_vertical_merge(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "vertical.docx"
            output_path = Path(tmp_dir) / "vertical.xlsx"

            doc = Document()
            table = doc.add_table(rows=3, cols=2)
            merged = table.cell(0, 0).merge(table.cell(1, 0))
            merged.text = "纵向合并"
            table.cell(0, 1).text = "右上"
            table.cell(1, 1).text = "右下"
            table.cell(2, 0).text = "末行"
            doc.save(source_path)

            export_word_tables_to_excel(str(source_path), str(output_path))
            workbook = load_workbook(output_path)
            try:
                worksheet = workbook["表格1"]
                ranges = {str(cell_range) for cell_range in worksheet.merged_cells.ranges}
                value = worksheet["A1"].value
                after_merge = worksheet["A3"].value
            finally:
                workbook.close()

        self.assertIn("A1:A2", ranges)
        self.assertEqual(value, "纵向合并")
        self.assertEqual(after_merge, "末行")

    def test_export_word_tables_keeps_paragraph_breaks_in_cells(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "paragraphs.docx"
            output_path = Path(tmp_dir) / "paragraphs.xlsx"

            doc = Document()
            table = doc.add_table(rows=1, cols=1)
            cell = table.cell(0, 0)
            cell.text = "第一段"
            cell.add_paragraph("第二段")
            doc.save(source_path)

            export_word_tables_to_excel(str(source_path), str(output_path))
            workbook = load_workbook(output_path)
            try:
                value = workbook["表格1"]["A1"].value
            finally:
                workbook.close()

        self.assertEqual(value, "第一段\n第二段")

    def test_export_word_tables_skips_document_without_tables(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "plain.docx"
            output_path = Path(tmp_dir) / "plain.xlsx"

            doc = Document()
            doc.add_paragraph("没有表格")
            doc.save(source_path)

            count = export_word_tables_to_excel(str(source_path), str(output_path))

        self.assertEqual(count, 0)
        self.assertFalse(output_path.exists())

    def test_batch_export_word_tables_uses_unique_non_overwriting_output_name(self):
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "sample.docx"
            existing_path = Path(tmp_dir) / "sample_表格.xlsx"
            existing_path.touch()

            doc = Document()
            table = doc.add_table(rows=1, cols=1)
            table.cell(0, 0).text = "内容"
            doc.save(source_path)

            results, skipped, error = batch_export_word_tables([str(source_path)], output_dir=tmp_dir)

        self.assertIsNone(error)
        self.assertEqual(skipped, {})
        self.assertIn("sample.docx", results)
        self.assertEqual(Path(results["sample.docx"]["output_path"]).name, "sample_表格_1.xlsx")
        self.assertEqual(results["sample.docx"]["tables"], 1)

    def test_table_export_output_path_defaults_to_source_directory(self):
        output_path = get_table_export_output_path(r"C:\work\sample.docx")

        self.assertEqual(Path(output_path).name, "sample_表格.xlsx")


if __name__ == "__main__":
    unittest.main()
