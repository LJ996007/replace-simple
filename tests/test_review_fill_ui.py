"""验证从扫描选择到后台填表的完整 UI 路径，不依赖用户项目文件。"""
import time
import tkinter as tk
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import load_workbook

import main
from review_table_filler import extract_review_rules, inspect_review_targets
from test_review_table_filler import make_template, scoring_table
from word_scan import file_stamp
from word_table_exporter import TableScanItem
from ui_common import _file_identity


def test_preview_generate_from_snapshot_and_fit_small_window():
    with TemporaryDirectory() as directory:
        folder = Path(directory)
        template = folder / "开评标.xlsx"
        make_template(template)
        source = folder / "招标.docx"
        source.write_bytes(b"snapshot source")
        table = scoring_table()
        item = TableScanItem(str(source), source.name, 1, "评分标准", "", "", 5, 3, "", table)
        root = tk.Tk()
        try:
            app = main.ReplaceSimpleApp(root, restore_session=False)
            root.geometry("1020x720")
            app.show_workspace("extract")
            exporter = app.table_export_window
            exporter.scan_stamps = {_file_identity(str(source)): file_stamp(source)}
            root.update()
            button = exporter.review_button
            assert button.winfo_ismapped()
            assert button.winfo_rooty() + button.winfo_height() <= root.winfo_rooty() + root.winfo_height()
            exporter._show_review_fill_dialog(str(template), [(item, extract_review_rules(table))],
                                             inspect_review_targets(template))
            root.update()
            dialog = next(w for w in root.winfo_children() if isinstance(w, tk.Toplevel))

            def children(widget):
                for child in widget.winfo_children():
                    yield child
                    yield from children(child)

            generate = next(w for w in children(dialog) if isinstance(w, main.ttk.Button)
                            and w.cget("text") == "生成并另存")
            with patch.object(exporter, "_show_result_window") as result:
                generate.invoke()
                deadline = time.monotonic() + 10
                while exporter._task_runner.active and time.monotonic() < deadline:
                    root.update()
                    time.sleep(.01)
                assert not exporter._task_runner.active
                result.assert_called_once()
            output = folder / "开评标_商务技术评审.xlsx"
            assert output.exists()
            assert load_workbook(output)["9-商务评审表"]["D9"].value == "=SUM(D6:D8)"
        finally:
            root.destroy()
