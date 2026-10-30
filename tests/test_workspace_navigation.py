"""双工作区的状态保留、任务生命周期与可见布局。"""

import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import main
from word_table_exporter import TableScanItem


class WorkspaceNavigationTests(unittest.TestCase):
    def setUp(self):
        self.root = main.tk.Tk()
        self.app = main.ReplaceSimpleApp(self.root, restore_session=False)
        self.root.update()

    def tearDown(self):
        self.root.destroy()

    def test_switch_preserves_rules_files_scan_selection_filters_and_directories(self):
        app = self.app
        source = str(Path("docs/_sample_files/招标文件示例.docx").resolve())
        app.replace_files = [source]
        app._remember_word_files([source])
        app._refresh_file_list()
        app.rules_sheet.set_sheet_data([["旧文本", "新文本"]])
        app._apply_output_dir("C:/replacement-output")
        self.assertEqual(app.active_workspace, "replace")
        app.workspace_buttons["extract"].invoke()
        self.root.update()
        exporter = app.table_export_window
        self.assertEqual(exporter.file_paths, [source])
        item = TableScanItem(source, Path(source).name, 1, "采购需求", "", "参数", 1, 1, "")
        exporter._show_scan_result([item], {}, None, [], {}, None)
        exporter.select_all_scan_items()
        exporter.keyword_var.set("参数")
        exporter.output_dir = "C:/extraction-output"
        selected = set(exporter.selected_scan_keys)
        scan_rows = exporter.scan_rows
        exporter.filter_boxes["section"].invoke()
        self.assertTrue(exporter._filter_popup_exists())

        app.workspace_buttons["replace"].invoke()
        self.root.update()
        self.assertFalse(exporter._filter_popup_exists())
        self.assertTrue(app.container.winfo_ismapped())
        self.assertFalse(exporter.window.winfo_ismapped())
        self.assertEqual(app.get_rules_from_table(), [("旧文本", "新文本")])
        self.assertEqual(app.replace_files, [source])
        self.assertEqual(app.output_dir, "C:/replacement-output")
        app.workspace_buttons["extract"].invoke()
        self.root.update()
        self.assertIs(app.table_export_window, exporter)
        self.assertIs(exporter.scan_rows, scan_rows)
        self.assertEqual(exporter.selected_scan_keys, selected)
        self.assertEqual(exporter.keyword_var.get(), "参数")
        self.assertEqual(exporter.output_dir, "C:/extraction-output")
        self.assertIs(exporter.window.winfo_toplevel(), self.root)
        exporter.clear_files()
        app.show_workspace("replace")
        app.show_workspace("extract")
        self.assertEqual(exporter.file_paths, [])

    def test_hidden_extraction_task_completes_and_close_remains_guarded(self):
        app = self.app
        app.show_workspace("extract")
        exporter = app.table_export_window
        release = threading.Event()
        completed, errors = [], []

        def finish(value):
            completed.append(value)
            exporter._set_busy_controls(False)
            exporter.status_var.set("扫描完成")

        exporter._set_busy_controls(True)
        exporter._task_runner.submit(lambda _publish: release.wait(5), finish, errors.append)
        try:
            app.workspace_buttons["replace"].invoke()
            self.assertEqual(app.active_workspace, "replace")
            with patch("main.messagebox.showinfo") as info:
                app.on_close()
            info.assert_called_once()
            self.assertTrue(self.root.winfo_exists())
            app._set_busy_controls(True)
            app.workspace_buttons["extract"].invoke()
            self.assertEqual(app.active_workspace, "extract")
            app.show_workspace("replace")
        finally:
            release.set()
        deadline = time.monotonic() + 5
        while exporter._task_runner.active and time.monotonic() < deadline:
            self.root.update()
            time.sleep(.01)
        self.assertEqual(completed, [True])
        self.assertEqual(errors, [])
        app.show_workspace("extract")
        self.assertEqual(exporter.status_var.get(), "扫描完成")
        self.assertFalse(exporter.scan_button.instate(["disabled"]))

    def test_navigation_and_actions_fit_minimum_window_and_keyboard_switches(self):
        app = self.app
        self.root.geometry("1020x720")
        self.root.update()
        for mode in ("replace", "extract", "replace"):
            app._navigate_workspace(mode)
            self.root.update()
            buttons = list(app.workspace_buttons.values())
            self.assertLessEqual(abs(buttons[0].winfo_width() - buttons[1].winfo_width()), 1)
            self.assertTrue(app.workspace_buttons[mode].instate(["selected"]))
            action = app.start_button if mode == "replace" else app.table_export_window.export_button
            self.assertTrue(action.winfo_ismapped())
            self.assertLessEqual(action.winfo_rooty() + action.winfo_height(),
                                 self.root.winfo_rooty() + self.root.winfo_height())
            self.assertLessEqual(action.winfo_rootx() + action.winfo_width(),
                                 self.root.winfo_rootx() + self.root.winfo_width())
        app.workspace_buttons["replace"].focus_force()
        app.workspace_buttons["replace"].event_generate("<Right>")
        self.root.update()
        self.assertEqual(app.active_workspace, "extract")


if __name__ == "__main__":
    unittest.main()
