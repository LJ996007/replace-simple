import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from string_replacer import replace_in_xlsx


class ExcelTextPreservationTests(unittest.TestCase):
    def test_real_workbook_keeps_wrap_alignment_row_height_and_rich_text(self):
        from datetime import datetime
        from openpyxl import Workbook, load_workbook
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        from openpyxl.styles import Alignment
        from openpyxl.utils.datetime import to_excel

        with TemporaryDirectory() as directory:
            source, output = Path(directory) / 'source.xlsx', Path(directory) / 'output.xlsx'
            book = Workbook()
            sheet = book.active
            sheet['A1'] = CellRichText(TextBlock(InlineFont(b=True), '\n before OLD\n'),
                                      TextBlock(InlineFont(i=True), 'after\n  '))
            sheet['A1'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
            sheet.merge_cells('A1:C4')
            sheet.row_dimensions[1].height = 100
            sheet['D1'] = 'unchanged\ntext'
            sheet['D2'] = '=1+1'
            original_date = datetime(2026, 11, 4)
            sheet['D3'] = original_date
            book.save(source)
            book.close()
            self.assertEqual(replace_in_xlsx(str(source), [('OLD', 'NEW'),
                             (str(int(to_excel(original_date))), 'not a date')], str(output)), 1)
            result = load_workbook(output, rich_text=True)
            try:
                sheet = result.active
                self.assertEqual(str(sheet['A1'].value), '\n before NEW\nafter\n  ')
                self.assertTrue(sheet['A1'].value[0].font.b)
                self.assertTrue(sheet['A1'].value[1].font.i)
                self.assertTrue(sheet['A1'].alignment.wrap_text)
                self.assertEqual(sheet['A1'].alignment.vertical, 'center')
                self.assertEqual(sheet.row_dimensions[1].height, 100)
                self.assertIn('A1:C4', {str(area) for area in sheet.merged_cells.ranges})
                self.assertEqual(sheet['D1'].value, 'unchanged\ntext')
                self.assertEqual(sheet['D2'].value, '=1+1')
                self.assertEqual(sheet['D3'].value, original_date)
            finally:
                result.close()
            with ZipFile(source) as original, ZipFile(output) as replaced:
                for name in original.namelist():
                    if name != 'xl/worksheets/sheet1.xml':
                        self.assertEqual(replaced.read(name), original.read(name))

    def test_replacement_preserves_raw_line_breaks_rich_runs_and_other_parts(self):
        sheet = b'''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1" ht="120" customHeight="1"><c r="Z9" s="1"/><c r="A1" s="7" t="inlineStr"><is><r><t/></r><r><rPr><b/></rPr><t xml:space="preserve">&#13;&#10;before\nOL</t></r><r><rPr><color rgb="FFFF0000"/></rPr><t xml:space="preserve">D\n after_x000D_\n  </t></r></is></c>
<c r="B1" s="2" t="inlineStr"><is><t>unchanged&#13;&#10;line\nnext</t></is></c>
<c r="C1"><f>OLD()</f><v>123</v></c></row></sheetData><mergeCells><mergeCell ref="A1:A3"/></mergeCells><extLst><custom/></extLst></worksheet>'''
        with TemporaryDirectory() as directory:
            source, output = Path(directory) / 'source.xlsm', Path(directory) / 'output.xlsm'
            parts = {'xl/worksheets/sheet1.xml': sheet, 'xl/styles.xml': b'<styleSheet><cellXfs><xf numFmtId="0"/></cellXfs></styleSheet>',
                     'xl/vbaProject.bin': b'macros', 'xl/drawings/drawing1.xml': b'original drawings',
                     'xl/worksheets/sheet2.xml': b'<worksheet><c>unrelated</c></worksheet>'}
            with ZipFile(source, 'w') as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)
            original = source.read_bytes()
            self.assertEqual(replace_in_xlsx(str(source), [('OLD', 'NEW')], str(output)), 1)
            self.assertEqual(source.read_bytes(), original)
            with ZipFile(output) as archive:
                expected = sheet.replace(b'before\nOL</t>', b'before\nNEW</t>').replace(b'>D\n after', b'>\n after')
                self.assertEqual(archive.read('xl/worksheets/sheet1.xml'), expected)
                for name, content in parts.items():
                    if name != 'xl/worksheets/sheet1.xml':
                        self.assertEqual(archive.read(name), content)

    def test_shared_strings_keep_escaped_breaks_and_count_each_cell(self):
        sheet = b'''<worksheet><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></worksheet>'''
        shared = b'''<sst count="3" uniqueCount="3"><si/><si><t xml:space="preserve">  OLD_x000D_\nnext&#13;&#10;line  </t></si><si><t>unmatched_x000D_\ntext</t></si></sst>'''
        with TemporaryDirectory() as directory:
            source, output = Path(directory) / 'source.xlsx', Path(directory) / 'result.xlsx'
            with ZipFile(source, 'w') as archive:
                archive.writestr('xl/worksheets/sheet1.xml', sheet)
                archive.writestr('xl/sharedStrings.xml', shared)
            self.assertEqual(replace_in_xlsx(str(source), [('OLD', '=NEW\nMORE')], str(output)), 2)
            with ZipFile(output) as archive:
                self.assertEqual(archive.read('xl/worksheets/sheet1.xml'), sheet)
                self.assertEqual(archive.read('xl/sharedStrings.xml'), shared.replace(b'OLD', b'=NEW\nMORE'))
