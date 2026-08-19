import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from symbol_clause_extractor import (
    DEFAULT_SYMBOL_CHARS,
    SYMBOL_CLAUSE_HEADERS,
    SYMBOL_SHEET_TITLE,
    batch_export_symbol_clauses,
    export_symbol_clauses_to_excel,
    extract_symbol_clauses,
    get_symbol_output_path,
    _leading_symbols,
    normalize_symbol_chars,
)


DECIMAL_NUMBERING = (
    '<w:numbering ' + nsdecls("w") + ">"
    '<w:abstractNum w:abstractNumId="100">'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
    '<w:lvlText w:val="%1、"/></w:lvl>'
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/>'
    '<w:lvlText w:val="%1.%2"/></w:lvl>'
    "</w:abstractNum>"
    '<w:num w:numId="9"><w:abstractNumId w:val="100"/></w:num>'
    "</w:numbering>"
)


def _apply_numbering(paragraph, num_id=9, ilvl=0):
    ppr = paragraph._p.get_or_add_pPr()
    ppr.append(parse_xml(
        f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{ilvl}"/>'
        f'<w:numId w:val="{num_id}"/></w:numPr>'
    ))


def _replace_numbering(document, numbering_xml):
    part = document.part.part_related_by(RT.NUMBERING)
    root = part.element
    for child in list(root):
        root.remove(child)
    for child in parse_xml(numbering_xml):
        root.append(child)


def _make_docx(path, paragraphs=(), table_rows=None, numbered_cell_rows=()):
    """Build a test docx.

    paragraphs: list of (text, num_id_or_None, ilvl) tuples for body paragraphs.
    table_rows: rows of plain cell texts, e.g. [("序号", "名称", "要求"), ...].
    numbered_cell_rows: row indexes whose first-column paragraph uses w:numPr.
    """
    document = Document()
    for text, num_id, ilvl in paragraphs:
        paragraph = document.add_paragraph(text)
        if num_id is not None:
            _apply_numbering(paragraph, num_id=num_id, ilvl=ilvl)

    if table_rows:
        table = document.add_table(rows=len(table_rows), cols=len(table_rows[0]))
        for row_index, row in enumerate(table_rows):
            for column_index, value in enumerate(row):
                cell = table.cell(row_index, column_index)
                cell.paragraphs[0].text = value
                if row_index in numbered_cell_rows and column_index == 0:
                    _apply_numbering(cell.paragraphs[0])

    if paragraphs and any(num_id is not None for _, num_id, _ in paragraphs) or numbered_cell_rows:
        _replace_numbering(document, DECIMAL_NUMBERING)

    document.save(path)
    return path


class LeadingSymbolRuleTests(unittest.TestCase):
    def test_symbol_at_clause_start(self):
        self.assertEqual(_leading_symbols("★分辨率不低于400万像素"), ["★"])
        self.assertEqual(_leading_symbols("  △支持宽动态"), ["△"])

    def test_symbol_after_literal_number(self):
        self.assertEqual(_leading_symbols("3.2、▲支持H.265"), ["▲"])
        self.assertEqual(_leading_symbols("1.★支持夜视"), ["★"])
        self.assertEqual(_leading_symbols("（一）★实质性要求"), ["★"])
        self.assertEqual(_leading_symbols("(3)#重要参数"), ["#"])
        self.assertEqual(_leading_symbols("一、△检测精度"), ["△"])
        self.assertEqual(_leading_symbols("第2条▲必须响应"), ["▲"])
        self.assertEqual(_leading_symbols("１２３★全角序号"), ["★"])

    def test_symbol_after_bracket(self):
        self.assertEqual(_leading_symbols("（★）实质性要求"), ["★"])
        self.assertEqual(_leading_symbols("【▲】关键指标"), ["▲"])

    def test_symbol_only_line_matches(self):
        self.assertEqual(_leading_symbols("★"), ["★"])

    def test_symbol_in_middle_is_ignored(self):
        self.assertEqual(_leading_symbols("该参数为★级，仅供参考"), [])
        self.assertEqual(_leading_symbols("分辨率≥400万像素 ★重要"), [])
        self.assertEqual(_leading_symbols("详见第三章★要求"), [])

    def test_symbol_after_clause_separator_counts(self):
        self.assertEqual(_leading_symbols("★支持A功能；▲支持B功能"), ["★", "▲"])
        self.assertEqual(_leading_symbols("★分辨率≥400万，△支持宽动态"), ["★", "△"])

    def test_multiple_symbols_expand_to_distinct(self):
        self.assertEqual(_leading_symbols("★▲双重要求"), ["★", "▲"])
        self.assertEqual(_leading_symbols("★★强调"), ["★"])


class ExtractBodyParagraphTests(unittest.TestCase):
    def test_body_symbols_extracted_with_full_text(self):
        with TemporaryDirectory() as directory:
            path = _make_docx(Path(directory) / "正文.docx", paragraphs=[
                ("三、采购需求", None, None),
                ("★摄像机分辨率不低于400万像素，具备OSD叠加功能", None, None),
                ("3.2、▲支持H.265编码，码率可调", None, None),
                ("普通条款不提取", None, None),
                ("该参数为★级，仅供参考", None, None),
            ])
            clauses = extract_symbol_clauses(str(path))

        self.assertEqual([clause.symbol for clause in clauses], ["★", "▲"])
        self.assertEqual(clauses[0].text, "★摄像机分辨率不低于400万像素，具备OSD叠加功能")
        self.assertEqual(clauses[1].text, "3.2、▲支持H.265编码，码率可调")
        self.assertIn("采购需求", clauses[0].section)

    def test_auto_numbering_reconstructed(self):
        with TemporaryDirectory() as directory:
            path = _make_docx(Path(directory) / "自动编号.docx", paragraphs=[
                ("采购需求", None, None),
                ("★分辨率不低于400万像素", 9, 0),
                ("△支持OSD叠加", 9, 1),
                ("▲支持宽动态", 9, 1),
                ("★防护等级IP66", 9, 0),
            ])
            clauses = extract_symbol_clauses(str(path))

        self.assertEqual(
            [clause.text for clause in clauses],
            [
                "1、★分辨率不低于400万像素",
                "1.一△支持OSD叠加",
                "1.二▲支持宽动态",
                "2、★防护等级IP66",
            ],
        )

    def test_multiple_symbols_produce_one_row_each(self):
        with TemporaryDirectory() as directory:
            path = _make_docx(Path(directory) / "多符号.docx", paragraphs=[
                ("★支持A功能；▲支持B功能", None, None),
            ])
            clauses = extract_symbol_clauses(str(path))

        self.assertEqual([(c.symbol, c.text) for c in clauses], [
            ("★", "★支持A功能；▲支持B功能"),
            ("▲", "★支持A功能；▲支持B功能"),
        ])


class ExtractTableTests(unittest.TestCase):
    TABLE_ROWS = [
        ("序号", "设备名称", "技术参数要求"),
        ("1", "网络摄像机", "★分辨率不低于400万像素\n▲支持H.265编码"),
        ("2", "交换机", "端口数不少于48个"),
        ("3", "★", "服务器CPU不低于32核"),
    ]

    def test_table_rows_extracted_whole_with_number_column(self):
        with TemporaryDirectory() as directory:
            path = _make_docx(Path(directory) / "表格.docx", table_rows=self.TABLE_ROWS)
            clauses = extract_symbol_clauses(str(path))

        self.assertEqual([clause.symbol for clause in clauses], ["★", "▲", "★"])
        camera_row = clauses[0].text
        self.assertIn("1", camera_row.splitlines())
        self.assertIn("网络摄像机", camera_row.splitlines())
        self.assertIn("★分辨率不低于400万像素", camera_row.splitlines())
        # 无符号行（交换机）不提取
        self.assertNotIn("交换机", "\n".join(clause.text for clause in clauses))
        # 纯符号单元格：整行仍完整提取
        self.assertIn("服务器CPU不低于32核", clauses[2].text)

    def test_table_cell_auto_numbering_reconstructed(self):
        with TemporaryDirectory() as directory:
            path = _make_docx(
                Path(directory) / "表格编号.docx",
                table_rows=[
                    ("序号", "设备", "要求"),
                    ("", "摄像机", "支持夜间成像"),
                    ("", "摄像机", "★自动编号条款"),
                ],
                numbered_cell_rows=(1, 2),
            )
            clauses = extract_symbol_clauses(str(path))

        self.assertEqual(len(clauses), 1)
        self.assertEqual(clauses[0].symbol, "★")
        self.assertEqual(clauses[0].text, "2、\n摄像机\n★自动编号条款")


class ExcelExportTests(unittest.TestCase):
    def test_export_writes_three_column_sheet(self):
        with TemporaryDirectory() as directory:
            directory = Path(directory)
            path = _make_docx(directory / "导出.docx", paragraphs=[
                ("★指标一", None, None),
                ("3.2、▲指标二", None, None),
            ])
            output = get_symbol_output_path(str(path), str(directory))
            count = export_symbol_clauses_to_excel(str(path), output)
            self.assertEqual(count, 2)
            self.assertTrue(output.endswith("_指标参数.xlsx"))

            workbook = load_workbook(output)
            worksheet = workbook.active
            self.assertEqual(worksheet.title, SYMBOL_SHEET_TITLE)
            self.assertEqual(
                [worksheet.cell(1, column).value for column in range(1, 4)],
                list(SYMBOL_CLAUSE_HEADERS),
            )
            self.assertEqual([worksheet.cell(2, column).value for column in range(1, 4)],
                             ["1", "★", "★指标一"])
            self.assertEqual([worksheet.cell(3, column).value for column in range(1, 4)],
                             ["2", "▲", "3.2、▲指标二"])

    def test_output_path_does_not_overwrite(self):
        with TemporaryDirectory() as directory:
            directory = Path(directory)
            path = _make_docx(directory / "重名.docx", paragraphs=[("★条款", None, None)])
            first = get_symbol_output_path(str(path), str(directory))
            Path(first).write_bytes(b"placeholder")
            second = get_symbol_output_path(str(path), str(directory))
            self.assertNotEqual(first, second)
            self.assertIn("_指标参数_1", second)

    def test_batch_skips_file_without_symbols(self):
        with TemporaryDirectory() as directory:
            directory = Path(directory)
            docx_with = _make_docx(directory / "有符号.docx", paragraphs=[("★条款", None, None)])
            docx_without = _make_docx(directory / "无符号.docx", paragraphs=[("普通条款", None, None)])
            (directory / "说明.txt").write_text("占位", encoding="utf-8")

            progress = []
            results, skipped, error = batch_export_symbol_clauses(
                [str(docx_with), str(docx_without), str(directory / "说明.txt")],
                output_dir=str(directory),
                progress_callback=lambda current, total, name: progress.append(name),
            )

            self.assertIsNone(error)
            self.assertEqual(len(progress), 3)
            self.assertEqual(list(results), ["有符号.docx"])
            self.assertEqual(results["有符号.docx"]["count"], 1)
            self.assertTrue(results["有符号.docx"]["output_path"].endswith("有符号_指标参数.xlsx"))
            self.assertEqual(
                skipped,
                {"无符号.docx": "未找到带符号条款", "说明.txt": "不支持的文件格式"},
            )


class CustomSymbolSetTests(unittest.TestCase):
    def test_default_symbols_are_the_four_common_markers(self):
        self.assertEqual(DEFAULT_SYMBOL_CHARS, "★#△▲")

    def test_normalize_symbol_chars_filters_and_dedupes(self):
        self.assertEqual(normalize_symbol_chars("★ ▲"), "★▲")
        self.assertEqual(normalize_symbol_chars("★★▲"), "★▲")
        self.assertEqual(normalize_symbol_chars("★1a中 空格▲"), "★▲")
        self.assertEqual(normalize_symbol_chars("◆◇＃＊"), "◆◇＃＊")
        self.assertEqual(normalize_symbol_chars(""), "")
        self.assertEqual(normalize_symbol_chars("12ab一二三"), "")

    def test_default_set_excludes_other_markers(self):
        self.assertEqual(_leading_symbols("☆普通建议条款"), [])
        self.assertEqual(_leading_symbols("◆重要参数"), [])

    def test_extract_with_custom_symbols(self):
        with TemporaryDirectory() as directory:
            path = _make_docx(Path(directory) / "自定义.docx", paragraphs=[
                ("★默认符号条款", None, None),
                ("◆自定义符号条款", None, None),
                ("☆未选符号条款", None, None),
            ])

            only_star = extract_symbol_clauses(str(path), symbols="★")
            self.assertEqual([(c.symbol, c.text) for c in only_star],
                             [("★", "★默认符号条款")])

            diamond_and_star = extract_symbol_clauses(str(path), symbols="★◆☆")
            self.assertEqual([c.symbol for c in diamond_and_star], ["★", "◆", "☆"])

    def test_custom_symbols_work_in_tables_and_numbering(self):
        with TemporaryDirectory() as directory:
            path = _make_docx(
                Path(directory) / "自定义表格.docx",
                paragraphs=[("◆正文条款", 9, 0)],
                table_rows=[("1", "设备", "◆表格条款")],
            )
            clauses = extract_symbol_clauses(str(path), symbols="◆")
            self.assertEqual([c.symbol for c in clauses], ["◆", "◆"])
            self.assertEqual(clauses[0].text, "1、◆正文条款")
            self.assertIn("◆表格条款", clauses[1].text)

    def test_unclean_custom_input_is_sanitized(self):
        with TemporaryDirectory() as directory:
            path = _make_docx(Path(directory) / "清洗.docx", paragraphs=[
                ("★条款一", None, None),
                ("◆条款二", None, None),
            ])
            # 空格、汉字会被清理掉，只剩 ◆
            clauses = extract_symbol_clauses(str(path), symbols=" ◆条 ")
            self.assertEqual([(c.symbol, c.text) for c in clauses],
                             [("◆", "◆条款二")])

    def test_main_default_matches_extractor_default(self):
        from main import DEFAULT_SYMBOL_CHARS as main_default
        self.assertEqual(main_default, DEFAULT_SYMBOL_CHARS)


if __name__ == "__main__":
    unittest.main()
