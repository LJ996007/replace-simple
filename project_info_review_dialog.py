"""Editable review dialog for extracted project replacement fields."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class ProjectInfoReviewDialog(tk.Toplevel):
    def __init__(self, parent, details, source_path, on_apply):
        super().__init__(parent)
        self.title("审核项目信息")
        self.details = details
        self.source_path = source_path
        self.on_apply = on_apply
        self.value_vars = []
        self.entries = []
        self.active_index = 0
        self.transient(parent)
        self.geometry("760x590")
        self.minsize(680, 520)
        self.configure(background=ttk.Style(self).lookup("Toolbar.TFrame", "background") or "#F7F8FA")

        header = ttk.Frame(self, padding=(16, 12, 16, 0), style="Toolbar.TFrame")
        header.pack(fill="x")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text=f"来源：{source_path}", style="Hint.TLabel", wraplength=680).grid(row=1, column=0, sticky="w", pady=(3, 0))

        fields = ttk.Frame(self, padding=(16, 8, 16, 0), style="Toolbar.TFrame")
        fields.pack(fill="x")
        fields.columnconfigure(1, weight=1)
        for index, item in enumerate(self.details):
            ttk.Label(fields, text=item["placeholder"]).grid(row=index, column=0, sticky="w", padx=(0, 8), pady=2)
            variable = tk.StringVar(value=item.get("value", ""))
            entry = ttk.Entry(fields, textvariable=variable)
            entry.grid(row=index, column=1, sticky="ew", pady=2)
            status = item.get("status", "缺失")
            ttk.Label(fields, text=status, style="Hint.TLabel", width=7).grid(row=index, column=2, padx=(8, 0))
            entry.bind("<FocusIn>", lambda _event, row=index: self.select_field(row))
            self.value_vars.append(variable)
            self.entries.append(entry)

        candidates_panel = ttk.Frame(self, padding=(16, 12, 16, 0), style="Toolbar.TFrame")
        candidates_panel.pack(fill="both", expand=True)
        candidates_panel.columnconfigure(0, weight=1)
        candidates_panel.rowconfigure(1, weight=1)
        ttk.Label(candidates_panel, text="候选原文（双击填入左侧对应字段）", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        columns = ("field", "kind", "value", "source")
        self.candidate_tree = ttk.Treeview(candidates_panel, columns=columns, show="headings", selectmode="browse")
        self.candidate_tree.heading("field", text="字段")
        self.candidate_tree.heading("kind", text="类型")
        self.candidate_tree.heading("value", text="候选值")
        self.candidate_tree.heading("source", text="来源")
        self.candidate_tree.column("field", width=120, minwidth=100, stretch=False)
        self.candidate_tree.column("kind", width=72, minwidth=60, stretch=False)
        self.candidate_tree.column("value", width=220, minwidth=140)
        self.candidate_tree.column("source", width=280, minwidth=160)
        self.candidate_tree.grid(row=1, column=0, sticky="nsew", pady=(5, 8))
        scrollbar = ttk.Scrollbar(self.candidate_tree, orient="vertical", command=self.candidate_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.candidate_tree.configure(yscrollcommand=scrollbar.set)
        self.candidate_tree.bind("<Double-Button-1>", self.use_selected_candidate)

        footer = ttk.Frame(self, padding=(16, 0, 16, 16), style="Toolbar.TFrame")
        footer.pack(fill="x")
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, text="空白值不会执行替换；确认替换会使用当前待处理文件。", style="Hint.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="仅回填规则", command=lambda: self.apply_review("fill")).grid(row=0, column=1, sticky="e", padx=(0, 7))
        ttk.Button(footer, text="确认并替换", style="Accent.TButton", command=lambda: self.apply_review("replace")).grid(row=0, column=2, sticky="e")

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _event: self.destroy())
        self.populate_candidates()
        if self.entries:
            self.entries[0].focus_set()

    def populate_candidates(self):
        for item in self.candidate_tree.get_children():
            self.candidate_tree.delete(item)
        for field_index, item in enumerate(self.details):
            for candidate in item.get("candidates", []):
                self.candidate_tree.insert("", "end", values=(
                    item["placeholder"], candidate.get("kind", ""), candidate.get("value", ""), candidate.get("source", "")
                ), tags=(field_index,))
        self.candidate_tree.tag_configure("conflict", foreground="#a33d00")

    def select_field(self, index):
        self.active_index = index

    def use_selected_candidate(self, _event=None):
        selection = self.candidate_tree.selection()
        if not selection or not 0 <= self.active_index < len(self.value_vars):
            return
        values = self.candidate_tree.item(selection[0], "values")
        self.value_vars[self.active_index].set(values[2] if len(values) > 2 else "")

    def reviewed_rules(self):
        return [
            (item["placeholder"], variable.get())
            for item, variable in zip(self.details, self.value_vars)
        ]

    def apply_review(self, action):
        rules = self.reviewed_rules()
        if self.on_apply and self.on_apply(action, rules) is False:
            return
        self.destroy()
