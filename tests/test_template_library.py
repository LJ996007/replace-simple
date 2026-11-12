import os
import unittest
from tempfile import TemporaryDirectory

from openpyxl import Workbook

import template_library as library
from project_info_review_dialog import ProjectInfoReviewDialog


class TemplateLibraryTests(unittest.TestCase):
    def test_category_and_template_lifecycle(self):
        with TemporaryDirectory() as directory:
            root = os.path.join(directory, "templates")
            source = os.path.join(directory, "开评标表.xlsx")
            Workbook().save(source)

            library.create_category(root, "公开招标")
            library.add_templates(root, "公开招标", [source])
            library.add_templates(root, "公开招标", [source])
            templates = library.list_templates(root)

            self.assertEqual(templates[0].category, "公开招标")
            self.assertEqual(len(templates), 2)
            self.assertIn("开评标表 (2).xlsx", {item.name for item in templates})
            with self.assertRaises(ValueError):
                library.delete_category(root, "公开招标")
            for template in templates:
                library.delete_template(root, template)
            library.delete_category(root, "公开招标")
            self.assertEqual(library.list_categories(root), [])

    def test_names_and_paths_are_constrained(self):
        with TemporaryDirectory() as directory:
            self.assertEqual(library.normalize_category_name("../公开招标"), ".. 公开招标")
            with self.assertRaises(ValueError):
                library.normalize_category_name("CON")
            with self.assertRaises(ValueError):
                library.category_path(directory, "..")


class ProjectInfoReviewDialogTests(unittest.TestCase):
    def test_review_preserves_field_order_and_candidate_apply(self):
        import tkinter as tk

        details = [
            {
                "placeholder": "[项目名称]",
                "value": "",
                "status": "冲突",
                "candidates": [{"value": "甲项目", "kind": "直接识别", "source": "正文 第1段"}],
            },
            {"placeholder": "[项目编号]", "value": "ZB-001", "status": "直接识别", "candidates": []},
        ]
        root = tk.Tk()
        root.withdraw()
        dialog = None
        try:
            dialog = ProjectInfoReviewDialog(root, details, "招标文件.docx", on_apply=lambda _action, rules: rules)
            child = dialog.candidate_tree.get_children()[0]
            dialog.candidate_tree.selection_set(child)
            dialog.active_index = 0
            dialog.use_selected_candidate()
            dialog.value_vars[1].set("ZB-002")
            self.assertEqual(dialog.reviewed_rules(), [("[项目名称]", "甲项目"), ("[项目编号]", "ZB-002")])
        finally:
            if dialog is not None:
                dialog.destroy()
            root.destroy()


class TemplateUseSafetyTests(unittest.TestCase):
    def test_template_use_requires_external_output_directory(self):
        import tkinter as tk
        from unittest.mock import patch

        import main

        root = tk.Tk()
        root.withdraw()
        try:
            app = main.ReplaceSimpleApp(root, restore_session=False)
            app.template_library_root = os.path.join(os.getcwd(), "templates")
            app.output_dir = None
            with patch("main.filedialog.askdirectory", return_value=app.template_library_root), \
                    patch("main.messagebox.showwarning") as warning:
                app.use_library_templates(["template.xlsx"])
            warning.assert_called_once()
            self.assertEqual(app.replace_files, [])

            app.output_dir = os.getcwd()
            app.use_library_templates(["template.xlsx"])
            self.assertEqual(app.replace_files, ["template.xlsx"])
        finally:
            root.destroy()


class TemplateLibraryDialogTests(unittest.TestCase):
    def test_dialog_opens_with_empty_library(self):
        import tkinter as tk

        from template_library_dialog import TemplateLibraryDialog

        with TemporaryDirectory() as directory:
            root = tk.Tk()
            root.withdraw()
            dialog = None
            try:
                dialog = TemplateLibraryDialog(
                    root,
                    directory,
                    on_root_changed=None,
                    on_use=None,
                )
                root.update_idletasks()
                self.assertEqual(dialog.category_list.get(0, "end"), ())
                self.assertEqual(dialog.template_list.get_children(), ())
                self.assertEqual(dialog.selected_templates, [])
            finally:
                if dialog is not None:
                    dialog.destroy()
                root.destroy()

    def test_templates_default_to_checked_and_support_individual_and_all_toggle(self):
        import tkinter as tk
        from types import SimpleNamespace
        from template_library_dialog import TemplateLibraryDialog

        with TemporaryDirectory() as directory:
            category = library.create_category(directory, "公开招标")
            for name in ("开标表.xlsx", "评标报告.docx"):
                with open(os.path.join(category, name), "wb"):
                    pass
            root = tk.Tk()
            root.withdraw()
            dialog = None
            try:
                dialog = TemplateLibraryDialog(root, directory)
                self.assertEqual(len(dialog.selected_templates), 2)
                self.assertTrue(dialog.select_all_var.get())
                dialog.template_list.focus("0")
                dialog.toggle_template(SimpleNamespace(keysym="space"))
                self.assertEqual(len(dialog.selected_templates), 1)
                self.assertTrue(dialog.select_all_button.instate(["alternate"]))
                dialog.select_all_button.invoke()
                self.assertEqual(len(dialog.selected_templates), 2)
                dialog.select_all_button.invoke()
                self.assertEqual(dialog.selected_templates, [])
                self.assertTrue(dialog.use_button.instate(["disabled"]))
                dialog.select_all_templates()
                self.assertEqual(len(dialog.selected_templates), 2)
            finally:
                if dialog is not None:
                    dialog.destroy()
                root.destroy()
