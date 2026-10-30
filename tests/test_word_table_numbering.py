import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from openpyxl import load_workbook

from word_table_exporter import scan_word_tables, export_word_tables_to_excel
from word_scan import scan_word_content, export_scanned_content
from test_symbol_clause_extractor import _apply_numbering, _replace_numbering


def numbering(fmt='decimal',text='（%1）',start=1):
    return ('<w:numbering '+nsdecls('w')+'>'
            '<w:abstractNum w:abstractNumId="100">'
            f'<w:lvl w:ilvl="0"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/></w:lvl>'
            '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:lvlText w:val="%1.%2"/></w:lvl></w:abstractNum>'
            '<w:num w:numId="9"><w:abstractNumId w:val="100"/></w:num></w:numbering>')


class TableNumberingTests(unittest.TestCase):
    def test_partial_paragraph_numbering_inherits_style_id_and_level_override(self):
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'source.docx';doc=Document()
            xml=numbering().replace('</w:num>','<w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="4"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%2)"/><w:suff w:val="space"/></w:lvl></w:lvlOverride></w:num>')
            _replace_numbering(doc,xml)
            style=doc.styles.add_style('编号基类',WD_STYLE_TYPE.PARAGRAPH)
            style.element.get_or_add_pPr().append(parse_xml('<w:numPr '+nsdecls('w')+'><w:numId w:val="9"/></w:numPr>'))
            cell=doc.add_table(rows=1,cols=1).cell(0,0)
            p=cell.paragraphs[0];p.style=style;p.text='保留格式'
            p._p.get_or_add_pPr().append(parse_xml('<w:numPr '+nsdecls('w')+'><w:ilvl w:val="1"/></w:numPr>'))
            doc.save(source)
            self.assertEqual(scan_word_tables(source)[0].table.cells[0].text,'D) 保留格式')

    def test_body_table_and_selected_export_keep_continuous_numbers(self):
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'source.docx';output=Path(folder)/'out.xlsx'
            doc=Document();_replace_numbering(doc,numbering())
            _apply_numbering(doc.add_paragraph('正文第一项'))
            table=doc.add_table(rows=1,cols=1)
            _apply_numbering(table.cell(0,0).paragraphs[0]);table.cell(0,0).paragraphs[0].add_run('第一条要求')
            _apply_numbering(table.cell(0,0).add_paragraph('第二条要求'))
            other=doc.add_table(rows=1,cols=1)
            _apply_numbering(other.cell(0,0).paragraphs[0]);other.cell(0,0).paragraphs[0].add_run('后表要求')
            doc.save(source)
            items=scan_word_tables(source)
            self.assertEqual(items[0].table.cells[0].text,'（2）第一条要求\n（3）第二条要求')
            self.assertEqual(items[1].table.cells[0].text,'（4）后表要求')
            export_word_tables_to_excel(source,output,table_indexes=[2])
            book=load_workbook(output);self.assertEqual(book.active['A1'].value,'（4）后表要求');book.close()

    def test_style_numbering_nested_table_and_merged_cell_count_once(self):
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'source.docx'
            doc=Document();_replace_numbering(doc,numbering())
            style=doc.styles.add_style('自动条款',WD_STYLE_TYPE.PARAGRAPH)
            style.element.get_or_add_pPr().append(parse_xml('<w:numPr '+nsdecls('w')+'><w:numId w:val="9"/></w:numPr>'))
            table=doc.add_table(rows=2,cols=2)
            merged=table.cell(0,0).merge(table.cell(1,0));merged.paragraphs[0].text='合并单元格';merged.paragraphs[0].style=style
            cell=table.cell(0,1);cell.paragraphs[0].text='外层';cell.paragraphs[0].style=style
            nested=cell.add_table(rows=1,cols=1);nested.cell(0,0).paragraphs[0].text='嵌套条款';nested.cell(0,0).paragraphs[0].style=style
            end=table.cell(1,1).paragraphs[0];end.text='结束';end.style=style
            doc.save(source)
            texts=[c.text for c in scan_word_tables(source)[0].table.cells]
            self.assertIn('（1）合并单元格',texts)
            self.assertTrue(any('（2）外层' in text and '（3）嵌套条款' in text for text in texts))
            self.assertIn('（4）结束',texts)

    def test_common_number_formats_and_manual_numbers_are_preserved(self):
        for fmt,text,expected in [('decimal','（%1）','（1）'),('decimalEnclosedCircle','%1','①'),
                                  ('chineseCounting','%1、','一、'),('upperLetter','%1.','A.'),
                                  ('lowerRoman','%1.','i.'),('bullet','•','•')]:
            with self.subTest(fmt=fmt),tempfile.TemporaryDirectory() as folder:
                source=Path(folder)/'source.docx';doc=Document();_replace_numbering(doc,numbering(fmt,text))
                cell=doc.add_table(rows=1,cols=1).cell(0,0)
                p=cell.paragraphs[0];p.text='自动条款';_apply_numbering(p)
                cell.add_paragraph('（2）手工编号条款')
                doc.save(source)
                self.assertEqual(scan_word_tables(source)[0].table.cells[0].text,expected+'自动条款\n（2）手工编号条款')

    def test_start_override_multilevel_and_scanned_snapshot_export(self):
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'source.docx';doc=Document()
            xml=numbering().replace('</w:num>','<w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride></w:num>')
            _replace_numbering(doc,xml)
            cell=doc.add_table(rows=1,cols=1).cell(0,0)
            cell.paragraphs[0].text='一级';_apply_numbering(cell.paragraphs[0])
            child=cell.add_paragraph('子项');_apply_numbering(child,ilvl=1)
            doc.save(source)
            tables,sections,stamps=scan_word_content([str(source)],'★')
            self.assertIsNone(tables[2])
            self.assertEqual(tables[0][0].table.cells[0].text,'（3）一级\n3.1子项')
            result=export_scanned_content(tables[0],sections[0],stamps,folder,'★',False)
            book=load_workbook(next(iter(result[0].values()))['output_path'])
            self.assertEqual(book.active['A1'].value,'（3）一级\n3.1子项');book.close()
