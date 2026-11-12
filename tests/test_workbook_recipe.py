"""Real legacy settings/JSON/XLSX compatibility and transactional rejection tests."""
from copy import deepcopy
import os
from pathlib import Path
import tkinter as tk
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

import test_qt_frontend as qt_tests
from test_qt_frontend import APP, wait_task
from qt_editor import VisualEditor
from workbook_session import WorkbookSession
from workbook_recipe import export_recipe, save_recipe_file, read_recipe_file, source_signature
from visual_fill import source_grid, apply_visual_plan
from review_mapping import layout_signature


class MixedRecipeTests(unittest.TestCase):
    setUp = qt_tests.NativeFrontendTests.setUp
    item = qt_tests.NativeFrontendTests.item

    def editor(self, items=None, stamps=None):
        if items is None:
            item, stamps = self.item()
            items = [item]
        editor = VisualEditor(None, self.session, items, stamps)
        def cleanup():
            editor.scratch.cleanup()
            editor.hide()
            editor.deleteLater()
            APP.processEvents()
        self.addCleanup(cleanup)
        return editor

    def recipe(self, operations):
        item, _ = self.item()
        return {'operations': operations, 'sources': [source_signature(source_grid(item.table))],
                'targets': {n: layout_signature(v, v['guess']) for n, v in self.session.sheets.items()}}

    def assert_rejected(self, editor, recipe, pattern=None):
        before = self.session.snapshot()
        history, future = list(self.session.history), list(self.session.future)
        files = set(Path(editor.scratch.name).iterdir())
        with self.assertRaisesRegex(ValueError, pattern or '.'):
            editor.apply_recipe(recipe)
        self.assertEqual(self.session.snapshot(), before)
        self.assertEqual(self.session.history, history)
        self.assertEqual(self.session.future, future)
        self.assertEqual(set(Path(editor.scratch.name).iterdir()), files)

    def test_real_legacy_saved_settings_mixed_replay_export_and_save(self):
        import main
        from visual_fill_dialog import VisualFillDialog
        from app_settings import load_settings
        item, stamps = self.item()
        original = self.path.read_bytes()
        root = tk.Tk()
        try:
            app = main.ReplaceSimpleApp(root, restore_session=False)
            app.show_workspace('extract')
            legacy = VisualFillDialog(app.table_export_window, str(self.path), [item])
            root.update()
            # These are original business/UI methods, not a hand-crafted Qt-format fixture.
            configs = [{'header': 1, 'factor': 1, 'description': 2, 'points': 3},
                       {'sheet': '评审', 'header': 1, 'start': 2, 'end': 3,
                        'factor': 'A', 'description': 'B', 'points': 'C', 'mode': 'match'}]
            from review_mapping import plan_fill
            from review_table_filler import extract_review_rules
            plan = plan_fill(str(self.path), extract_review_rules(item.table, configs[0]), configs[1])
            self.assertTrue(legacy.receive_auto(plan, item, configs))
            legacy.stage_drop(0, (2, 3, 2, 3), '评审', 3, 2)
            legacy.stage_edit('评审', 1, 3, '手工因素')
            legacy.stage_clear('评审', (3, 3, 3, 3))
            for action, rows, height in [('height', [2], 60), ('width', [2], 35), ('fit', [3], None),
                                          ('delete', [4], None)]:
                legacy.layouts = legacy.build_layout(action, '评审', rows, legacy.pending, legacy.layouts, height)
                legacy.operations.append({'kind': 'layout', 'action': action, 'sheet': '评审', 'rows': rows, 'height': height})
            legacy.stage_structure('insert_below', '评审', (1, 3, 3, 3), 1)
            legacy.stage_structure('merge', '评审', (1, 4, 2, 4))
            legacy.stage_edit('评审', 1, 4, '新增合并内容')
            expected = self.root/'legacy-output.xlsx'
            apply_visual_plan(legacy.working_template, expected, legacy.preview_plan(legacy.pending))
            with patch.dict(os.environ, {'APPDATA': str(self.root)}), \
                 patch('visual_fill_dialog.simpledialog.askstring', return_value='整套业务配方'):
                legacy.save_recipe()
                saved = load_settings()['review_drag_profiles']['整套业务配方']
                self.assertEqual(len(saved['operations']), 11)
                editor = self.editor([item], stamps)
                self.session.edit('评审', {'B2': '重放前的内容'})
                before = self.session.snapshot()
                with patch('qt_editor.QInputDialog.getItem', return_value=('整套业务配方', True)):
                    editor.load_saved_recipe()
                    wait_task(editor.tasks)
                self.assertEqual(self.session.operations, saved['operations'])
                output = self.root/'qt-output.xlsx'
                self.session.save(output)
                self.assertEqual(self.path.read_bytes(), original)
                old, new = load_workbook(expected), load_workbook(output)
                try:
                    for name in old.sheetnames:
                        self.assertEqual(list(old[name].values), list(new[name].values))
                        self.assertEqual(str(old[name].merged_cells), str(new[name].merged_cells))
                    self.assertEqual(new['评审']['D2'].value, '=SUM(C2:C3)')
                    self.assertEqual(new['评审'].row_dimensions[2].height, 60)
                    self.assertEqual(new['评审'].column_dimensions['B'].width, 35)
                    self.assertTrue(new['评审']['B3'].alignment.wrap_text)
                    self.assertTrue(new['评审']['B2'].font.bold)
                finally:
                    old.close(); new.close()
                editor.undo()
                self.assertEqual(self.session.snapshot(), before)
                editor.redo()
                recipe_path = self.root/'完整历史.json'
                with patch('qt_editor.QFileDialog.getSaveFileName', return_value=(str(recipe_path), 'JSON')):
                    editor.save_recipe(); wait_task(editor.tasks)
                portable = read_recipe_file(recipe_path)
                self.assertEqual(portable['operations'], saved['operations'])
                with patch('qt_editor.QFileDialog.getOpenFileName', return_value=(str(recipe_path), 'JSON')):
                    editor.load_recipe(); wait_task(editor.tasks)
                self.assertEqual(self.session.operations, saved['operations'])
        finally:
            root.destroy()

    def test_native_history_survives_undo_redo_json_and_structure(self):
        editor = self.editor()
        self.session.edit('评审', {'B2': '多行\n正文', 'C3': '5'})
        editor.transfer_values((2, 3, 2, 3), 3, 2)
        self.session.layout('评审', heights={2: 65}, widths={2: 44}, wrap_cells=['B2'])
        self.session.structure(editor.temp_path(), '评审', 'insert_below', (1, 3, 3, 3), 1, editor.temp_path())
        self.session.edit('评审', {'A4': '=作为文字'})
        self.session.undo()
        operations = deepcopy(self.session.operations)
        self.session.redo()
        self.assertEqual(len(self.session.operations), len(operations)+1)
        recipe = export_recipe(self.session, editor.items)
        path = self.root/'native.json'
        save_recipe_file(path, recipe)
        # A recipe starts from the original template, even after prior structure edits.
        editor.apply_recipe(read_recipe_file(path))
        self.assertEqual(self.session.operations, recipe['operations'])
        output = self.root/'native.xlsx'; self.session.save(output)
        book = load_workbook(output)
        try:
            self.assertEqual(book['评审']['B2'].value, '多行\n正文')
            self.assertEqual(book['评审']['B3'].value, '新服务')
            self.assertEqual(book['评审']['A4'].value, '=作为文字')
            self.assertEqual(book['评审']['A4'].data_type, 's')
            self.assertEqual(book['评审'].row_dimensions[2].height, 65)
            self.assertEqual(book['评审'].column_dimensions['B'].width, 44)
        finally:
            book.close()

    def test_partial_structure_then_formula_failure_cleans_and_keeps_redo(self):
        editor = self.editor()
        self.session.edit('评审', {'B2': '已有编辑'})
        self.session.undo()
        recipe = self.recipe([{'kind': 'structure', 'sheet': '评审', 'action': 'insert_below', 'bounds': [1, 3, 3, 3]},
                              {'kind': 'edit', 'sheet': '评审', 'cell': 'D2', 'value': '不能覆盖'}])
        self.assert_rejected(editor, recipe, '公式')
        self.session.redo()
        self.assertEqual(self.session.pending['评审']['B2'], '已有编辑')

    def test_sources_are_matched_by_signature_and_ambiguous_or_changed_rejected(self):
        editor = self.editor()
        recipe = self.recipe([{'kind': 'drag', 'source': 0, 'sheet': '评审', 'bounds': [2, 2, 2, 2], 'row': 2, 'column': 2}])
        editor.items.append(editor.items[0])
        self.assert_rejected(editor, recipe, '多个同结构')
        editor.items.pop()
        recipe['sources'][0] = 'not-the-original-structure'
        self.assert_rejected(editor, recipe, '结构已变化')

    def test_source_change_and_template_change_are_rejected_before_replay(self):
        editor = self.editor()
        recipe = self.recipe([{'kind': 'edit', 'sheet': '评审', 'cell': 'B2', 'value': '新'}])
        source = self.source.read_bytes()
        self.source.write_bytes(source+b'changed')
        self.assert_rejected(editor, recipe, '重新扫描')
        self.source.write_bytes(source)
        _, stamps = self.item(); editor.stamps = stamps
        self.path.write_bytes(self.path.read_bytes()+b'changed')
        self.assert_rejected(editor, recipe, '变化')

    def test_target_signature_missing_sheet_invalid_bounds_and_unknown_operation(self):
        editor = self.editor()
        base = self.recipe([{'kind': 'edit', 'sheet': '评审', 'cell': 'B2', 'value': '新'}])
        cases = []
        recipe = deepcopy(base); recipe['targets']['评审'] = 'changed'; cases.append(recipe)
        recipe = deepcopy(base); recipe['operations'][0]['sheet'] = '不存在'; cases.append(recipe)
        recipe = deepcopy(base); recipe['operations'][0]['kind'] = 'execute'; cases.append(recipe)
        recipe = deepcopy(base); recipe['operations'][0]['cell'] = 'A0'; cases.append(recipe)
        recipe = self.recipe([{'kind': 'drag', 'source': -1, 'sheet': '评审', 'bounds': [1, 1, 2, 2], 'row': 1, 'column': 1}]); cases.append(recipe)
        recipe = self.recipe([{'kind': 'drag', 'source': 0, 'sheet': '评审', 'bounds': [1, 1, 9999, 9999], 'row': 1, 'column': 1}]); cases.append(recipe)
        for recipe in cases:
            with self.subTest(recipe=recipe):
                self.assert_rejected(editor, recipe)

    def test_protected_array_formula_and_invalid_layout_are_rejected(self):
        editor = self.editor()
        for kind, update in [('array', lambda book: setattr(book['评审']['B3'], 'value', ArrayFormula(ref='B3:B4', text='=1'))),
                             ('protected', lambda book: setattr(book['评审'].protection, 'sheet', True))]:
            with self.subTest(kind=kind):
                book = load_workbook(self.path); update(book); book.save(self.path); book.close()
                self.session = WorkbookSession.open(self.path); editor.session = self.session
                recipe = self.recipe([{'kind': 'edit', 'sheet': '评审', 'cell': 'B3', 'value': '拒绝'}])
                self.assert_rejected(editor, recipe, '公式|保护')
        # Use the other, unprotected sheet for layout validation.
        recipe = self.recipe([{'kind': 'layout', 'sheet': '原样保留', 'action': 'height', 'rows': [1], 'height': float('nan')}])
        self.assert_rejected(editor, recipe, '行高')

    def test_deleted_empty_row_coordinates_are_kept_and_materialized(self):
        book = load_workbook(self.path)
        book['评审']['A7'] = '尾行'
        book.save(self.path); book.close()
        self.session = WorkbookSession.open(self.path)
        editor = self.editor()
        recipe = self.recipe([{'kind': 'layout', 'sheet': '评审', 'action': 'delete', 'rows': [4], 'height': None},
                              {'kind': 'edit', 'sheet': '评审', 'cell': 'B5', 'value': '删行后的第五行'},
                              {'kind': 'drag', 'source': 0, 'sheet': '评审', 'bounds': [2, 3, 2, 3], 'row': 6, 'column': 2}])
        editor.apply_recipe(recipe)
        self.assertEqual(self.session.layouts, {})
        self.assertEqual(self.session.sheets['评审']['cells']['B4'], '删行后的第五行')
        output = self.root/'deleted.xlsx'; self.session.save(output)
        book = load_workbook(output)
        try:
            self.assertEqual(book['评审']['B4'].value, '删行后的第五行')
            self.assertEqual(book['评审']['B5'].value, '新服务')
        finally:
            book.close()

    def test_change_during_replay_rolls_back_and_deletes_new_versions(self):
        editor = self.editor()
        recipe = self.recipe([{'kind': 'structure', 'sheet': '评审', 'action': 'insert_below', 'bounds': [1, 3, 3, 3]}])
        structure = WorkbookSession.structure
        def change_source(work, *args, **kwargs):
            structure(work, *args, **kwargs)
            self.source.write_bytes(self.source.read_bytes()+b'changed during replay')
        with patch.object(WorkbookSession, 'structure', change_source):
            self.assert_rejected(editor, recipe, '重新扫描')

    def test_same_filename_new_source_header_is_not_accepted(self):
        from docx import Document
        editor = self.editor()
        recipe = self.recipe([{'kind': 'drag', 'source': 0, 'sheet': '评审', 'bounds': [2, 2, 2, 2], 'row': 2, 'column': 2}])
        doc = Document(self.source)
        doc.tables[0].cell(0, 1).text = '另一种结构'
        doc.save(self.source)
        item, stamps = self.item()
        editor.items, editor.stamps = [item], stamps
        self.assert_rejected(editor, recipe, '结构已变化')

    def test_v1_json_keeps_existing_edits_and_saves_combined_history(self):
        editor = self.editor()
        editor.transfer_values((2, 2, 2, 2), 2, 2)
        recipe = dict(editor.last_recipe)
        editor.undo()
        self.session.edit('原样保留', {'A1': '此前编辑'})
        editor.apply_recipe(recipe)
        self.assertEqual(self.session.pending['原样保留']['A1'], '此前编辑')
        self.assertEqual(self.session.pending['评审']['B2'], '新说明')
        self.assertEqual([op['kind'] for op in self.session.operations], ['cells', 'drag'])
        portable = export_recipe(self.session, editor.items)
        editor.apply_recipe(portable)
        self.assertEqual(self.session.pending['原样保留']['A1'], '此前编辑')

    def test_noop_legacy_empty_row_layout_before_structure_is_supported(self):
        editor = self.editor()
        recipe = self.recipe([{'kind': 'layout', 'sheet': '评审', 'action': 'delete', 'rows': [2], 'height': None},
                              {'kind': 'structure', 'sheet': '评审', 'action': 'insert_below', 'bounds': [1, 3, 3, 3]},
                              {'kind': 'edit', 'sheet': '评审', 'cell': 'A4', 'value': '插入行'}])
        editor.apply_recipe(recipe)
        output = self.root/'noop-layout.xlsx'; self.session.save(output)
        book = load_workbook(output)
        try:
            self.assertEqual(book['评审']['A4'].value, '插入行')
            self.assertEqual(book['评审']['A2'].value, '技术方案')
        finally:
            book.close()
