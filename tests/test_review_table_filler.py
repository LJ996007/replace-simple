import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, Side

from review_table_filler import (
    extract_review_rules, fill_review_workbook, inspect_review_targets, review_output_path,
)
from word_table_exporter import ExportedTable, ExportedTableCell


def scoring_table():
    cells = []
    for r, values in enumerate([
        ["评审因素", "评分说明", "分值"],
        ["投标报价", "报价评分原文", "30"],
        ["技术响应", "=原文应保持为文本\n#重要参数", "69"],
        ["政策性得分", "节能产品原文", "0.5"],
        [None, "环境标志产品原文", "0.5"],
    ], 1):
        for c, value in enumerate(values, 1):
            if value is not None:
                cells.append(ExportedTableCell(r, c, 2 if (r, c) == (4, 1) else 1, 1, value))
    return ExportedTable(cells, 5, 3)


def make_template(path):
    wb = Workbook()
    main = wb.active
    main.title = "概要"
    main["A2"] = "项目名称：测试"
    ws = wb.create_sheet("9-商务评审表")
    for cell, value in {"A1": "商务技术评审表", "A2": "=概要!A2", "A4": "序号", "B4": "评审因素",
                        "C4": "评分标准说明", "D4": "分值", "E4": 1, "F4": 2,
                        "E5": "甲公司", "F5": "乙公司", "A6": 1, "A7": 2,
                        "A8": "合计", "D8": "=SUM(D6:D7)", "A9": "评标委员会成员："}.items():
        ws[cell] = value
    for merge in ["A1:F1", "A2:F2", "A4:A5", "B4:B5", "C4:C5", "D4:D5", "A8:C8", "A9:F9"]:
        ws.merge_cells(merge)
    for row in ws.iter_rows(min_row=6, max_row=7, max_col=6):
        for cell in row:
            cell.font = Font(name="宋体", size=11)
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"))
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 45
    ws.print_area = "A1:F9"
    ws.print_title_rows = "1:5"
    wb.create_sheet("10-商务技术评审表")["A1"] = "既有内容不修改"
    wb.save(path)


class ReviewTableFillerTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.template = Path(self.tmp.name) / "模板.xlsx"
        self.output = Path(self.tmp.name) / "结果.xlsx"
        make_template(self.template)
        self.rules = extract_review_rules(scoring_table())

    def fill(self):
        return fill_review_workbook(self.template, self.output, "9-商务评审表", self.rules)

    def test_price_excluded_raw_text_and_merged_factor_retained(self):
        self.assertEqual(len(self.rules.criteria), 3)
        self.assertEqual(self.rules.total, Decimal("70"))
        self.assertEqual(self.rules.excluded_prices[0].points, Decimal("30"))
        self.assertEqual(self.rules.criteria[1].factor_group, self.rules.criteria[2].factor_group)
        self.assertTrue(self.rules.criteria[0].description.startswith("="))

    def test_expand_template_preserves_every_other_part_and_cached_total(self):
        before = self.template.read_bytes()
        result = self.fill()
        self.assertEqual(result["count"], 3)
        self.assertEqual(self.template.read_bytes(), before)
        with ZipFile(self.template) as original, ZipFile(self.output) as output:
            self.assertEqual(original.namelist(), output.namelist())
            changed = {name for name in original.namelist() if original.read(name) != output.read(name)}
            self.assertEqual(changed, {"xl/worksheets/sheet2.xml", "xl/workbook.xml"})
        wb = load_workbook(self.output)
        ws = wb["9-商务评审表"]
        self.assertEqual(ws["C6"].value, self.rules.criteria[0].description)
        self.assertEqual(ws["C6"].data_type, "s")
        self.assertEqual(ws["D7"].value, .5)
        self.assertEqual(ws["D9"].value, "=SUM(D6:D8)")
        self.assertEqual(ws["A10"].value, "评标委员会成员：")
        self.assertIn("B7:B8", str(ws.merged_cells))
        self.assertEqual(ws["E5"].value, "甲公司")
        self.assertTrue(all(ws.cell(r, c).value is None for r in range(6, 9) for c in (5, 6)))
        self.assertEqual(ws.print_area, "'9-商务评审表'!$A$1:$F$10")
        self.assertEqual(ws.print_title_rows, "$1:$5")
        cached = load_workbook(self.output, data_only=True)
        self.assertEqual(cached["9-商务评审表"]["D9"].value, 70)

    def test_second_fill_does_not_accumulate_rows_or_merges(self):
        self.fill()
        second = Path(self.tmp.name) / "再次生成.xlsx"
        fill_review_workbook(self.output, second, "9-商务评审表", self.rules)
        a, b = load_workbook(self.output), load_workbook(second)
        self.assertEqual(a["9-商务评审表"].max_row, b["9-商务评审表"].max_row)
        self.assertEqual(str(a["9-商务评审表"].merged_cells), str(b["9-商务评审表"].merged_cells))

    def test_reject_ambiguous_points_and_total_mismatch(self):
        table = scoring_table()
        table.cells[-1].text = "0~0.5"
        with self.assertRaisesRegex(ValueError, "明确数值"):
            extract_review_rules(table)
        table = scoring_table()
        table.cells[-1].row_span = 2
        with self.assertRaisesRegex(ValueError, "跨行合并"):
            extract_review_rules(table)
        table = scoring_table()
        table.row_count = 6
        table.cells.extend([ExportedTableCell(6, 1, 1, 1, "合计"),
                            ExportedTableCell(6, 3, 1, 1, "99")])
        with self.assertRaisesRegex(ValueError, "不一致"):
            extract_review_rules(table)

    def test_merged_score_group_keeps_all_descriptions_and_counts_once(self):
        table = scoring_table()
        table.cells = [c for c in table.cells if (c.row, c.column) != (5, 3)]
        next(c for c in table.cells if (c.row, c.column) == (4, 3)).row_span = 2
        rules = extract_review_rules(table)
        self.assertEqual(len(rules.criteria), 2)
        group = rules.criteria[1]
        self.assertEqual(group.points, Decimal('0.5'))
        self.assertEqual(group.description, '节能产品原文\n\n环境标志产品原文')
        self.assertEqual(rules.total, Decimal('69.5'))
        output = Path(self.tmp.name) / '分组结果.xlsx'
        fill_review_workbook(self.template, output, '9-商务评审表', rules)
        book = load_workbook(output)
        self.assertEqual(book['9-商务评审表']['C7'].value, group.description)
        self.assertEqual(book['9-商务评审表']['D7'].value, .5)
        cached = load_workbook(output, data_only=True)
        self.assertEqual(cached['9-商务评审表']['D8'].value, 69.5)
        book.close(); cached.close()

    def test_merged_score_different_subitems_and_identical_text_preserved(self):
        cells = [ExportedTableCell(1,c,1,1,t) for c,t in enumerate(['名称','评分标准','分值'],1)]
        cells += [ExportedTableCell(2,1,1,1,'子项一'), ExportedTableCell(3,1,1,1,'子项二'),
                  ExportedTableCell(2,2,1,1,'相同原文'), ExportedTableCell(3,2,1,1,'相同原文'),
                  ExportedTableCell(2,3,2,1,'10'), ExportedTableCell(4,1,1,1,'合计'),
                  ExportedTableCell(4,3,1,1,'10')]
        rules = extract_review_rules(ExportedTable(cells,4,3))
        self.assertEqual(rules.total, 10)
        self.assertEqual(len(rules.criteria), 1)
        self.assertEqual(rules.criteria[0].factor, '子项一 / 子项二')
        self.assertEqual(rules.criteria[0].description, '子项一\n相同原文\n\n子项二\n相同原文')

    def test_merged_description_read_once_and_price_group_excluded(self):
        cells = [ExportedTableCell(1,c,1,1,t) for c,t in enumerate(['名称','评分标准','分值'],1)]
        cells += [ExportedTableCell(2,1,2,1,'价格'), ExportedTableCell(2,2,2,1,'价格组原文'),
                  ExportedTableCell(2,3,2,1,'30'), ExportedTableCell(4,1,2,1,'技术响应'),
                  ExportedTableCell(4,2,2,1,'技术组原文'),ExportedTableCell(4,3,2,1,'70'),
                  ExportedTableCell(6,1,1,1,'合计'),ExportedTableCell(6,3,1,1,'100')]
        table = ExportedTable(cells,6,3)
        rules = extract_review_rules(table)
        self.assertEqual(rules.total,70)
        self.assertEqual(rules.excluded_prices[0].points,30)
        self.assertEqual(rules.criteria[0].description,'技术组原文')
        self.assertEqual(extract_review_rules(table, include_price=True).total,100)

    def test_shrink_moves_signature_and_print_area_back(self):
        from dataclasses import replace
        self.rules = replace(self.rules, criteria=self.rules.criteria[:1])
        self.fill()
        ws = load_workbook(self.output)["9-商务评审表"]
        self.assertEqual(ws["D7"].value, "=SUM(D6:D6)")
        self.assertEqual(ws["A8"].value, "评标委员会成员：")
        self.assertEqual(ws.print_area, "'9-商务评审表'!$A$1:$F$8")

    def test_refuse_existing_scores_without_touching_files(self):
        wb = load_workbook(self.template)
        wb["9-商务评审表"]["E6"] = 0
        wb.save(self.template)
        original = self.template.read_bytes()
        with self.assertRaisesRegex(ValueError, "已有评分"):
            self.fill()
        self.assertEqual(self.template.read_bytes(), original)
        self.assertFalse(self.output.exists())

    def test_refuse_cross_sheet_dependencies_and_unknown_formulas(self):
        wb = load_workbook(self.template)
        wb["概要"]["B1"] = "='9-商务评审表'!D8"
        wb.save(self.template)
        with self.assertRaisesRegex(ValueError, "引用了目标表"):
            self.fill()
        wb["概要"]["B1"] = None
        wb["9-商务评审表"]["E8"] = "=SUM(E6:E7)"
        wb.save(self.template)
        with self.assertRaisesRegex(ValueError, "非标准公式"):
            self.fill()

    def test_never_overwrite_and_unique_output_name(self):
        with self.assertRaisesRegex(ValueError, "不能覆盖模板"):
            fill_review_workbook(self.template, self.template, "9-商务评审表", self.rules)
        self.output.write_bytes(b"existing")
        with self.assertRaisesRegex(ValueError, "已存在"):
            self.fill()
        self.assertEqual(self.output.read_bytes(), b"existing")
        first = review_output_path(str(self.template))
        Path(first).touch()
        self.assertNotEqual(first, review_output_path(str(self.template)))

    def test_target_detects_blank_review_table(self):
        targets = inspect_review_targets(self.template)
        self.assertEqual(len(targets), 1)
        self.assertEqual((targets[0].start_row, targets[0].total_row), (6, 8))
        self.assertTrue(targets[0].blank)


if __name__ == "__main__":
    unittest.main()
