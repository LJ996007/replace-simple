import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET

from workbook_structure import edit_structure
from visual_fill import plan_visual_fill, apply_visual_plan
from review_table_filler import _tag
import test_worksheet_layout as fixtures

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = {'chart': 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      'pivotTable': 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
      'pivotCacheDefinition': 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
      'externalLink': 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'}


def rewrite(path, parts, overrides=(), defaults=()):
    """把部件原样写入工作簿并登记内容类型，保持包结构有效。"""
    with ZipFile(path) as src:
        entries = {e.filename: src.read(e.filename) for e in src.infolist()}
    entries.update(parts)
    types = entries['[Content_Types].xml'].decode('utf-8')
    for part, ctype in overrides:
        if f'PartName="/{part}"' not in types:
            types = types.replace('</Types>', f'<Override PartName="/{part}" ContentType="{ctype}"/></Types>')
    for extension, ctype in defaults:
        if f'Extension="{extension}"' not in types:
            types = types.replace('</Types>', f'<Default Extension="{extension}" ContentType="{ctype}"/></Types>')
    entries['[Content_Types].xml'] = types.encode('utf-8')
    with ZipFile(path, 'w') as out:
        for name, data in entries.items():
            out.writestr(name, data)


def serialize(root):
    return ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def chart_parts(refs):
    formulas = ''.join(f'<c:f>{ref}</c:f>' for ref in refs)
    chart = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
             f'<c:chart><c:plotArea><c:barChart><c:ser>{formulas}</c:ser></c:barChart>'
             '</c:plotArea></c:chart></c:chartSpace>')
    return {'xl/charts/chart1.xml': chart.encode('utf-8')}


def add_external_link(path, formula_ref, formula):
    """注入 externalLinks 部件、workbook 引用和一条使用外部函数的公式。"""
    with ZipFile(path) as src:
        entries = {e.filename: src.read(e.filename) for e in src.infolist()}
    entries['xl/externalLinks/externalLink1.xml'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<externalLink xmlns="{NS}" xmlns:r="{REL}"><externalBook r:id="rId1">'
        '<sheetNames><sheetName val="Sheet1"/></sheetNames></externalBook></externalLink>').encode('utf-8')
    entries['xl/externalLinks/_rels/externalLink1.xml.rels'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" '
        'Type="http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing" '
        'Target="file:///C:\\模板库\\代理费.xlam" TargetMode="External"/></Relationships>').encode('utf-8')
    rels = ET.fromstring(entries['xl/_rels/workbook.xml.rels'])
    used = {r.get('Id') for r in rels}
    rid = next(f'rId{i}' for i in range(1, 100) if f'rId{i}' not in used)
    ET.SubElement(rels, 'Relationship', Id=rid, Type=f'{REL}/externalLink',
                  Target='externalLinks/externalLink1.xml', TargetMode='External')
    entries['xl/_rels/workbook.xml.rels'] = serialize(rels)
    book = ET.fromstring(entries['xl/workbook.xml'])
    references = ET.Element(_tag('externalReferences'))
    ET.SubElement(references, _tag('externalReference'), {f'{{{REL}}}id': rid})
    book.insert(list(book).index(book.find(_tag('sheets'))) + 1, references)
    entries['xl/workbook.xml'] = serialize(book)
    sheet = 'xl/worksheets/sheet1.xml'
    root = ET.fromstring(entries[sheet])
    rownum = formula_ref[len(formula_ref.rstrip('0123456789')):]
    row = next(r for r in root.find(_tag('sheetData')) if r.get('r') == rownum)
    cell = ET.SubElement(row, _tag('c'), r=formula_ref)
    cell.append(ET.fromstring(f'<f xmlns="{NS}">{formula}</f>'))
    entries[sheet] = serialize(root)
    rewrite(path, entries, [('xl/externalLinks/externalLink1.xml', CT['externalLink'])])


def pivot_parts(location, cache_sheet, cache_ref):
    return {
        'xl/pivotTables/pivotTable1.xml': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<pivotTableDefinition xmlns="{NS}" name="透视" cacheId="1" dataCaption="值">'
            f'<location ref="{location}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>'
            '</pivotTableDefinition>').encode('utf-8'),
        'xl/pivotTables/_rels/pivotTable1.xml.rels': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{REL}/pivotCacheDefinition" '
            'Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>').encode('utf-8'),
        'xl/pivotCache/pivotCacheDefinition1.xml': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<pivotCacheDefinition xmlns="{NS}" refreshedBy="t" recordCount="0" cacheId="1">'
            f'<cacheSource type="worksheet"><worksheetSource sheet="{cache_sheet}" ref="{cache_ref}"/></cacheSource>'
            '</pivotCacheDefinition>').encode('utf-8'),
        'xl/worksheets/_rels/sheet1.xml.rels': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{REL}/pivotTable" '
            'Target="../pivotTables/pivotTable1.xml"/></Relationships>').encode('utf-8'),
    }


def drawing_parts(from_row, to_row, vml_rows):
    """隐藏宏按钮样式的一对锚点：xdr twoCellAnchor 与 VML Anchor（0 起始行列）。"""
    drawing = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<xdr:twoCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff>'
        f'<xdr:row>{from_row}</xdr:row><xdr:rowOff>15240</xdr:rowOff></xdr:from>'
        '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>30480</xdr:colOff>'
        f'<xdr:row>{to_row}</xdr:row><xdr:rowOff>426720</xdr:rowOff></xdr:to>'
        '<xdr:sp macro="[0]!zbinfo"><xdr:nvSpPr><xdr:cNvPr id="2049" name="Button 1" hidden="1"/>'
        '<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/></xdr:txBody></xdr:sp>'
        '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>').encode('utf-8')
    col1, row1, col2, row2 = vml_rows
    vml = (
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:x="urn:schemas-microsoft-com:office:excel">'
        f'<v:shape id="_x0000_s2049" type="#button"><x:Anchor>{col1},15,{row1},0,{col2},131,{row2},0</x:Anchor></v:shape>'
        '</xml>').encode('utf-8')
    return drawing, vml


def attach_drawing(path, from_row, to_row, vml_rows):
    drawing, vml = drawing_parts(from_row, to_row, vml_rows)
    with ZipFile(path) as src:
        entries = {e.filename: src.read(e.filename) for e in src.infolist()}
    entries['xl/drawings/drawing1.xml'] = drawing
    entries['xl/drawings/vmlDrawing1.vml'] = vml
    entries['xl/worksheets/_rels/sheet1.xml.rels'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL}">'
        f'<Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing1.xml"/>'
        f'<Relationship Id="rId2" Type="{REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
        '</Relationships>').encode('utf-8')
    sheet = 'xl/worksheets/sheet1.xml'
    root = ET.fromstring(entries[sheet])
    ET.SubElement(root, _tag('drawing'), {f'{{{REL}}}id': 'rId1'})
    ET.SubElement(root, _tag('legacyDrawing'), {f'{{{REL}}}id': 'rId2'})
    entries[sheet] = serialize(root)
    rewrite(path, entries,
            [('xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing')],
            [('vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing')])


def part_text(path, part):
    with ZipFile(path) as book:
        return book.read(part).decode('utf-8')


class StructureObjectTests(unittest.TestCase):
    def setUp(self):
        fixtures.WorksheetLayoutTests.setUp(self)

    def test_external_link_formula_moves_local_refs_only(self):
        add_external_link(self.path, 'E2', '[1]!dlf(D4,B3)*C6')
        edit_structure(self.path, self.out, '评分', 'insert_above', (1, 2, 3, 2), 1)
        sheet = part_text(self.out, 'xl/worksheets/sheet1.xml')
        self.assertEqual(sheet.count('[1]!dlf'), 1)  # 原型行复制时不带公式
        self.assertIn('<f>[1]!dlf(D5,B4)*C7</f>', sheet)
        self.assertIn('r="E3"', sheet)
        with ZipFile(self.path) as before, ZipFile(self.out) as after:
            self.assertEqual(before.read('xl/externalLinks/externalLink1.xml'),
                             after.read('xl/externalLinks/externalLink1.xml'))

    def test_chart_references_follow_insert_and_delete(self):
        rewrite(self.path, chart_parts(['评分!$A$1:$C$7', '[1]Sheet1!$A$1:$A$9']),
                [('xl/charts/chart1.xml', CT['chart'])])
        edit_structure(self.path, self.out, '评分', 'insert_above', (1, 2, 3, 2), 1)
        chart = part_text(self.out, 'xl/charts/chart1.xml')
        self.assertIn('评分!$A$1:$C$8', chart)
        self.assertIn('[1]Sheet1!$A$1:$A$9', chart)  # 外部引用不随本簿移动
        deleted = Path(self.temp.name) / 'deleted.xlsx'
        edit_structure(self.path, deleted, '评分', 'delete_rows', (1, 3, 3, 3))
        self.assertIn('评分!$A$1:$C$6', part_text(deleted, 'xl/charts/chart1.xml'))

    def test_chart_range_fully_deleted_refused(self):
        rewrite(self.path, chart_parts(['评分!$A$4:$A$4']), [('xl/charts/chart1.xml', CT['chart'])])
        with self.assertRaisesRegex(ValueError, '已被删除'):
            edit_structure(self.path, self.out, '评分', 'delete_rows', (1, 4, 1, 4))
        self.assertFalse(self.out.exists())

    def test_pivot_location_and_cache_follow_insert(self):
        rewrite(self.path, pivot_parts('A3:E7', '评分', 'A1:C7'),
                [('xl/pivotTables/pivotTable1.xml', CT['pivotTable']),
                 ('xl/pivotCache/pivotCacheDefinition1.xml', CT['pivotCacheDefinition'])])
        edit_structure(self.path, self.out, '评分', 'insert_above', (1, 2, 3, 2), 1)
        self.assertIn('ref="A4:E8"', part_text(self.out, 'xl/pivotTables/pivotTable1.xml'))
        self.assertIn('ref="A1:C8"', part_text(self.out, 'xl/pivotCache/pivotCacheDefinition1.xml'))

    def test_pivot_inside_location_or_source_refused(self):
        rewrite(self.path, pivot_parts('A3:E7', '评分', 'A1:C7'),
                [('xl/pivotTables/pivotTable1.xml', CT['pivotTable']),
                 ('xl/pivotCache/pivotCacheDefinition1.xml', CT['pivotCacheDefinition'])])
        for action, bounds in [('insert_above', (1, 5, 1, 5)), ('delete_rows', (1, 4, 1, 4))]:
            with self.assertRaisesRegex(ValueError, '透视'):
                edit_structure(self.path, self.out, '评分', action, bounds)
            self.assertFalse(self.out.exists())

    def test_layout_preflight_accepts_objects_and_plan_applies(self):
        rewrite(self.path, chart_parts(['评分!$A$1:$C$7']), [('xl/charts/chart1.xml', CT['chart'])])
        layout = {'评分': {'heights': {2: 60}, 'delete_rows': [3], 'wrap_cells': []}}
        plan = plan_visual_fill(self.path, {}, layout)  # 旧版在此直接拒绝图表工作簿
        apply_visual_plan(self.path, self.out, plan)
        self.assertIn('评分!$A$1:$C$6', part_text(self.out, 'xl/charts/chart1.xml'))

    def test_hidden_buttons_move_with_inserted_rows(self):
        attach_drawing(self.path, from_row=1, to_row=3, vml_rows=(2, 1, 3, 3))
        edit_structure(self.path, self.out, '评分', 'insert_above', (1, 2, 3, 2), 1)
        drawing = part_text(self.out, 'xl/drawings/drawing1.xml')
        self.assertIn('<xdr:row>2</xdr:row>', drawing)  # from 锚行 1→2
        self.assertIn('<xdr:row>4</xdr:row>', drawing)  # to 锚行 3→4
        self.assertNotIn('<xdr:row>1</xdr:row>', drawing)
        vml = part_text(self.out, 'xl/drawings/vmlDrawing1.vml')
        self.assertIn('<x:Anchor>2,15,2,0,3,131,4,0</x:Anchor>', vml)

    def test_drawing_on_deleted_row_refused(self):
        attach_drawing(self.path, from_row=2, to_row=3, vml_rows=(2, 3, 3, 4))
        with self.assertRaisesRegex(ValueError, '待删除的行'):
            edit_structure(self.path, self.out, '评分', 'delete_rows', (1, 3, 3, 3))
        self.assertFalse(self.out.exists())


if __name__ == '__main__':
    unittest.main()
