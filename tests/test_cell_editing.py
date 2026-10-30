import tempfile
import time
import tkinter as tk
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from visual_fill import source_grid
from word_table_exporter import ExportedTable, ExportedTableCell, TableScanItem


class CellEditingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'template.xlsm'
        book = Workbook(); sheet = book.active; sheet.title = '商务评审'
        sheet['A1'] = '评分说明'; sheet['B1'] = '分值'
        sheet.merge_cells('A1:C1')  # 跨列标题行：整列选择不得被它扩展
        sheet.merge_cells('A2:A3'); sheet.merge_cells('B2:B3')
        sheet['A2'] = '旧说明'; sheet['B4'] = '=SUM(B2:B3)'
        book.create_sheet('其他')['D4'] = '原内容'
        book.save(self.path); book.close()
        self.table = ExportedTable([
            ExportedTableCell(1, 1, 1, 1, '名称'), ExportedTableCell(1, 2, 1, 1, '评分标准'),
            ExportedTableCell(2, 1, 2, 1, '技术方案'), ExportedTableCell(2, 2, 1, 1, '第一条原文'),
            ExportedTableCell(2, 3, 2, 1, '10')], 2, 3)
        self.item = TableScanItem('snapshot.docx', 'snapshot.docx', 1, '评分标准', '', '', 2, 3, '', self.table)

    def make_dialog(self):
        import main
        from visual_fill_dialog import VisualFillDialog
        self.root = tk.Tk()
        self.addCleanup(self.root.destroy)
        app = main.ReplaceSimpleApp(self.root, restore_session=False); app.show_workspace('extract')
        owner = app.table_export_window
        from word_scan import file_stamp
        from ui_common import _file_identity
        source = Path(self.temp.name) / 'source.docx'; source.write_bytes(b'snapshot')
        owner.scan_stamps = {_file_identity(str(source)): file_stamp(source)}
        self.item = TableScanItem(str(source), source.name, 1, '评分标准', '', '', 2, 3, '', self.table)
        dialog = VisualFillDialog(owner, str(self.path), [self.item]); self.root.update()
        return dialog

    def test_column_header_click_drag_and_status(self):
        dialog = self.make_dialog(); right = dialog.right
        x1 = int((right.xs[1] + right.xs[2]) / 2)
        x2 = int((right.xs[2] + right.xs[3]) / 2)
        y = int(right.header_height / 2)
        right.canvas.event_generate('<ButtonPress-1>', x=x1, y=y)
        right.canvas.event_generate('<ButtonRelease-1>', x=x1, y=y); self.root.update()
        # A1:C1 标题合并被跳过：点击 B 列不会扩展成 A–C 三列。
        self.assertEqual(right.selection, (2, 1, 2, right.info['max_row']))
        self.assertIn('列', dialog.status.get())
        right.canvas.event_generate('<ButtonPress-1>', x=x1, y=y)
        right.canvas.event_generate('<B1-Motion>', x=x2, y=y)
        right.canvas.event_generate('<ButtonRelease-1>', x=x2, y=y); self.root.update()
        self.assertEqual(right.selection, (2, 1, 3, right.info['max_row']))
        # Shift + 点击继续扩展到 E 列。
        right.canvas.event_generate('<ButtonPress-1>', x=int((right.xs[4] + right.xs[5]) / 2), y=y, state=0x1)
        self.root.update()
        self.assertEqual(right.selection, (2, 1, 5, right.info['max_row']))

    def test_edit_menu_order(self):
        dialog = self.make_dialog()
        labels = [dialog.edit_menu.entrycget(i, 'label') for i in range(dialog.edit_menu.index('end') + 1)
                  if dialog.edit_menu.type(i) == 'command']
        self.assertEqual(labels, ['编辑单元格…', '清空所选内容', '设置行高…', '删除所选行',
                                  '合并单元格', '拆分单元格', '上方插入行…', '下方插入行…',
                                  '左侧插入列…', '右侧插入列…', '删除空行'])

    def test_stage_edit_anchor_undo_formula_guard_and_write(self):
        import main
        dialog = self.make_dialog()
        dialog.stage_edit('商务评审', 1, 2, '新说明文字')
        self.assertEqual(dialog.pending['商务评审']['A2'], '新说明文字')
        # 合并格下半部分编辑落到左上锚点 A2。
        dialog.stage_edit('商务评审', 1, 3, '锚点改写')
        self.assertEqual(dialog.pending['商务评审']['A2'], '锚点改写')
        self.assertNotIn('A3', dialog.pending['商务评审'])
        # 数字文本按数值写入。
        dialog.sheet.set('其他'); dialog.target_changed(); self.root.update()
        dialog.stage_edit('其他', 4, 4, '123.5')
        self.assertEqual(dialog.pending['其他']['D4'], Decimal('123.5'))
        # 公式格拒绝覆盖。
        dialog.sheet.set('商务评审'); dialog.target_changed(); self.root.update()
        with self.assertRaisesRegex(ValueError, '公式'):
            dialog.stage_edit('商务评审', 2, 4, '覆盖公式')
        dialog.undo(); dialog.undo(); dialog.undo()
        self.assertEqual(dialog.pending, {})
        dialog.stage_edit('商务评审', 1, 2, '终稿说明')
        dialog.stage_edit('商务评审', 1, 2, '')  # 留空清除
        with patch.object(dialog.owner, '_show_result_window') as result:
            dialog.write()
            deadline = time.monotonic() + 15
            while dialog.owner._task_runner.active and time.monotonic() < deadline:
                self.root.update(); time.sleep(.01)
            result.assert_called_once()
        output = Path(self.temp.name) / 'template_商务技术评审.xlsm'
        book = load_workbook(output)
        self.assertTrue(book['商务评审']['A2'].value in ('', None))
        self.assertEqual(book['商务评审']['B4'].value, '=SUM(B2:B3)')
        self.assertEqual(book['其他']['D4'].value, '原内容')
        book.close()

    def test_double_click_f2_and_dialog_widget(self):
        dialog = self.make_dialog(); right = dialog.right
        x = int((right.xs[0] + right.xs[1]) / 2); y = int((right.ys[1] + right.ys[2]) / 2)
        with patch('cell_edit_dialog.ask_cell_text', return_value='双击修改') as ask:
            # 真实连点两下（第二次按下触发编辑），而非 Tk 的双击绑定。
            for _ in range(2):
                right.canvas.event_generate('<ButtonPress-1>', x=x, y=y); self.root.update()
                right.canvas.event_generate('<ButtonRelease-1>', x=x, y=y); self.root.update()
            self.assertEqual(dialog.pending['商务评审']['A2'], '双击修改')
            dialog.undo()
            right.select(2, 1)
            dialog.window.focus_force(); self.root.update()
            dialog.window.event_generate('<F2>'); self.root.update()
            self.assertEqual(dialog.pending['商务评审']['A2'], '双击修改')
            self.assertEqual(ask.call_count, 2)
            self.assertEqual(ask.call_args[0][3], '旧说明')  # 初值为当前显示内容
        from cell_edit_dialog import CellEditDialog
        editor = CellEditDialog(dialog.window, dialog.app, 'B2', '旧')
        editor.text.insert('end', '追加'); editor.apply()
        self.assertEqual(editor.result, '旧追加')
        editor = CellEditDialog(dialog.window, dialog.app, 'B2', '旧')
        editor.close()
        self.assertIsNone(editor.result)

    def test_recipe_replays_edits(self):
        import main
        from app_settings import load_settings
        dialog = self.make_dialog()
        dialog.stage_edit('商务评审', 1, 2, '复用文字')
        with patch.dict('os.environ', {'APPDATA': self.temp.name}), \
             patch('visual_fill_dialog.simpledialog.askstring', return_value='编辑设置'):
            dialog.save_recipe()
            recipe = load_settings()['review_drag_profiles']['编辑设置']
            dialog.clear()
            self.assertEqual(dialog.pending, {})
            dialog.replay_recipe(recipe)
            self.assertEqual(dialog.pending['商务评审']['A2'], '复用文字')

    def test_clear_selected_cells_skips_formulas_and_anchors(self):
        dialog = self.make_dialog()
        dialog.stage_drop(0, (2, 2, 2, 2), '商务评审', 3, 1)  # 拖到 A3，合并格落锚点 A2
        self.assertEqual(dialog.pending['商务评审'], {'A2': '第一条原文'})
        skipped = dialog.stage_clear('商务评审', (1, 2, 3, 3))
        self.assertEqual(skipped, 0)
        self.assertEqual(dialog.pending['商务评审'], {'A2': ''})  # 合并落锚点、空格跳过
        dialog.undo()
        self.assertEqual(dialog.pending['商务评审'], {'A2': '第一条原文'})
        # 只选公式格时无可清空内容并明确提示。
        with self.assertRaisesRegex(ValueError, '公式'):
            dialog.stage_clear('商务评审', (2, 4, 2, 4))
        # 含公式与其他内容的选区：公式跳过，其余清空。
        skipped = dialog.stage_clear('商务评审', (1, 2, 2, 4))
        self.assertEqual(skipped, 1)
        self.assertEqual(dialog.pending['商务评审'], {'A2': ''})
        self.assertNotIn('B4', dialog.pending['商务评审'])
        with patch.object(dialog.owner, '_show_result_window') as result:
            dialog.write()
            deadline = time.monotonic() + 15
            while dialog.owner._task_runner.active and time.monotonic() < deadline:
                self.root.update(); time.sleep(.01)
            result.assert_called_once()
        output = Path(self.temp.name) / 'template_商务技术评审.xlsm'
        book = load_workbook(output)
        self.assertTrue(book['商务评审']['A2'].value in ('', None))
        self.assertEqual(book['商务评审']['B4'].value, '=SUM(B2:B3)')
        self.assertEqual(book['其他']['D4'].value, '原内容')
        book.close()

    def test_table_popout_edit_and_restore(self):
        dialog = self.make_dialog(); right = dialog.right; host = dialog.right_host
        host.update_idletasks()
        embedded = (right.winfo_width(), right.winfo_height())
        dialog.popout_button.invoke(); self.root.update()
        self.assertTrue(dialog._table_popped)
        self.assertIs(host.winfo_toplevel(), host)  # 宿主已成为独立窗口
        self.assertGreater(right.winfo_width(), embedded[0])
        self.assertGreater(right.winfo_height(), embedded[1])
        # 弹出状态下画布交互照常。
        x = int((right.xs[0] + right.xs[1]) / 2); y = int((right.ys[1] + right.ys[2]) / 2)
        right.canvas.event_generate('<ButtonPress-1>', x=x, y=y)
        right.canvas.event_generate('<ButtonRelease-1>', x=x, y=y); self.root.update()
        self.assertEqual(right.selection, (1, 2, 1, 3))
        dialog.popout_button.invoke(); self.root.update()
        self.assertFalse(dialog._table_popped)
        self.assertIs(host.winfo_toplevel(), dialog.window)
        self.assertEqual(int(host.grid_info()['row']), 2)  # 嵌回原位
        # 还原后改点 B 列（与上一次不同位置，避免构成双击编辑）。
        xb = int((right.xs[1] + right.xs[2]) / 2)
        right.canvas.event_generate('<ButtonPress-1>', x=xb, y=y)
        right.canvas.event_generate('<ButtonRelease-1>', x=xb, y=y); self.root.update()
        self.assertEqual(right.selection, (2, 2, 2, 3))

    def test_cell_edit_dialog_enlarge_toggle(self):
        dialog = self.make_dialog()
        from cell_edit_dialog import CellEditDialog
        editor = CellEditDialog(dialog.window, dialog.app, 'B2', '旧')
        editor.window.deiconify(); editor.window.update()
        small = (editor.window.winfo_width(), editor.window.winfo_height())
        editor.toggle_size(); editor.window.update_idletasks()
        large = (editor.window.winfo_width(), editor.window.winfo_height())
        self.assertGreater(large[0], small[0]); self.assertGreater(large[1], small[1])
        self.assertEqual(editor.enlarge_button.cget('text'), '还原')
        editor.toggle_size(); editor.window.update_idletasks()
        self.assertEqual((editor.window.winfo_width(), editor.window.winfo_height()), small)
        self.assertEqual(editor.enlarge_button.cget('text'), '放大')
        editor.close()


if __name__ == '__main__':
    unittest.main()
