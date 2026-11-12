import tkinter as tk
from tkinter import ttk
import unittest

from ui_common import FlatScrollbar, cap_existing_scrollbar


class ScrollbarBehaviorTests(unittest.TestCase):
    def test_native_ranges_and_disabled_state_follow_content_on_both_axes(self):
        root = tk.Tk()
        root.geometry("360x240")
        try:
            canvas = tk.Canvas(root, highlightthickness=0)
            canvas.pack(fill="both", expand=True)
            vertical = FlatScrollbar(root, command=canvas.yview)
            horizontal = ttk.Scrollbar(root, orient="horizontal")
            horizontal_model = cap_existing_scrollbar(
                horizontal, canvas.xview, canvas, orientation="horizontal"
            )
            canvas.configure(yscrollcommand=vertical.set, scrollregion=(0, 0, 100, 100))
            root.update()
            self.assertEqual(vertical.get(), (0.0, 1.0))
            self.assertEqual(horizontal.get(), (0.0, 1.0))
            self.assertTrue(vertical.instate(["disabled"]))
            self.assertTrue(horizontal.instate(["disabled"]))

            canvas.configure(scrollregion=(0, 0, 1200, 1000))
            root.update()
            self.assertFalse(vertical.instate(["disabled"]))
            self.assertFalse(horizontal.instate(["disabled"]))
            self.assertEqual(vertical.get(), canvas.yview())
            self.assertEqual(horizontal.get(), canvas.xview())
            self.assertAlmostEqual(vertical.get()[1], canvas.winfo_height() / 1000)

            vertical._capped_model.dispatch("moveto", 1)
            horizontal_model.dispatch("moveto", 1)
            root.update()
            self.assertEqual(vertical.get(), canvas.yview())
            self.assertEqual(horizontal.get(), canvas.xview())
            self.assertEqual(vertical.get()[1], 1.0)
            self.assertEqual(horizontal.get()[1], 1.0)

            canvas.configure(scrollregion=(0, 0, 100, 100))
            root.update()
            self.assertEqual(vertical.get(), (0.0, 1.0))
            self.assertEqual(horizontal.get(), (0.0, 1.0))
            self.assertTrue(vertical.instate(["disabled"]))
            self.assertTrue(horizontal.instate(["disabled"]))
        finally:
            root.destroy()
