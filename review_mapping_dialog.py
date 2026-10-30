"""开评标模板的列对应、行匹配和持久配置界面。"""
import os
import tkinter as tk
from tkinter import ttk
from app_dialogs import messagebox, simpledialog

from review_mapping import (FIELDS, inspect_workbook, plan_fill, apply_plan,
                            load_profiles, save_profile, layout_signature)
from review_table_filler import HEADERS, _norm, extract_review_rules, review_output_path
from openpyxl.utils.cell import get_column_letter, range_boundaries


class ReviewMappingDialog:
    def __init__(self, owner, template, items, sheets, preview_receiver=None):
        self.preview_receiver = preview_receiver
        self.owner, self.template, self.items, self.sheets = owner, template, items, sheets
        self.plan = None
        self.overrides = {}
        self.profiles = load_profiles()
        self.app = owner.app
        self.window = tk.Toplevel(self.app.root)
        self.window.title("填入开评标表")
        self.window.configure(bg=self.app.app_bg)
        self.window.geometry("1000x740")
        self.window.minsize(800, 600)
        self.window.transient(self.app.root)
        try:
            self.window.iconphoto(True, self.app._icon_photo)
        except (AttributeError, tk.TclError):
            pass
        style = ttk.Style(self.window)
        style.configure("Review.TNotebook", background=self.app.surface_bg, borderwidth=0)
        style.configure("Review.TNotebook.Tab", padding=(12, 6), background="#E1E6ED", foreground=self.app.muted_fg)
        style.map("Review.TNotebook.Tab", background=[("selected", self.app.surface_bg)], foreground=[("selected", self.app.accent_fg)])
        body = ttk.Frame(self.window, padding=14, style="App.TFrame")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)
        settings = ttk.Frame(body, padding=12, style="Surface.TFrame")
        settings.grid(row=0, column=0, sticky="ew")
        settings.columnconfigure(1, weight=1)
        ttk.Label(settings, text="开评标文件", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.file_label = ttk.Label(settings, text=os.path.basename(template), width=1, anchor="w")
        self.file_label.grid(row=0, column=1, sticky="ew")
        ttk.Button(settings, text="更换文件…", style="Header.TButton", command=self.change_file).grid(row=0, column=2, padx=(12, 0))
        self.profile = ttk.Combobox(settings, state="readonly", values=list(self.profiles), width=25)
        ttk.Label(settings, text="复用设置").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.profile.grid(row=1, column=1, sticky="ew", pady=(10, 0))
        self.profile.bind("<<ComboboxSelected>>", self.restore_profile)
        ttk.Button(settings, text="记住这套设置…", style="Header.TButton", command=self.store_profile).grid(row=1, column=2, padx=(12, 0), pady=(10, 0))
        self.profile.set("选择已保存的设置")
        self.source = ttk.Combobox(settings, state="readonly", values=[
            f"表 {item.table_index} · {item.section or '未标注章节'}" for item in items])
        ttk.Label(settings, text="招标来源表").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.source.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(10, 0))
        self.sheet = tk.StringVar()
        ttk.Label(settings, text="填入工作表").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.sheet_box = ttk.Combobox(settings, textvariable=self.sheet, state="readonly", values=list(sheets))
        self.sheet_box.grid(row=3, column=1, columnspan=2, sticky="ew", pady=(10, 0))
        options = ttk.Frame(settings, style="Toolbar.TFrame")
        options.grid(row=4, column=1, columnspan=2, sticky="ew", pady=(10, 0))
        self.mode = tk.StringVar(value="match")
        self.mode_label = tk.StringVar(value="按名称找到对应项目")
        self.mode_box = ttk.Combobox(options, textvariable=self.mode_label, state="readonly",
                                    values=["按名称找到对应项目", "按原表顺序依次填入"], width=24)
        self.mode_box.pack(side="left")
        self.mode_box.bind("<<ComboboxSelected>>", lambda _: self.mode.set("match" if self.mode_box.current() == 0 else "sequence"))
        self.mode.trace_add("write", lambda *_: self.mode_label.set("按名称找到对应项目" if self.mode.get() == "match" else "按原表顺序依次填入"))
        ttk.Label(settings, text="填入方式").grid(row=4, column=0, sticky="w", pady=(10, 0))
        self.include_price = tk.BooleanVar(value=False)
        from ui_common import CanvasCheckbox
        self.price_check = CanvasCheckbox(options, "含价格评分", self.include_price, self.app, pady=3)
        self.price_check.canvas.pack(side="left", padx=(12, 0))
        self.source_vars = {key: tk.StringVar(value=str(i+1)) for i, key in enumerate(FIELDS)}
        self.source_header = tk.StringVar(value="1")
        self.target_vars = {key: tk.StringVar(value=get_column_letter(i+1)) for i, key in enumerate(FIELDS)}
        self.header, self.start, self.end = (tk.StringVar(value=v) for v in ("1", "2", "10"))
        self.mapping_summary = tk.StringVar()
        ttk.Label(settings, textvariable=self.mapping_summary, style="Hint.TLabel").grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))
        self.adjust_button = ttk.Button(settings, text="调整列和行…", style="Header.TButton", command=self.edit_mapping)
        self.adjust_button.grid(row=5, column=2, sticky="e", pady=(10, 0))
        actions = ttk.Frame(body, style="App.TFrame")
        actions.grid(row=1, column=0, sticky="ew", pady=(12, 6))
        ttk.Button(actions, text="预览填入结果", command=self.preview).pack(side="left")
        self.assign_button = ttk.Button(actions, text="选择填入行…", command=self.assign_row, state="disabled")
        self.assign_button.pack(side="left", padx=6)
        self.skip_button = ttk.Button(actions, text="跳过此项", command=lambda: self.assign_row(skip=True), state="disabled")
        self.skip_button.pack(side="left")
        ttk.Button(actions, text="重新匹配", command=self.reset_rows).pack(side="right")
        notebook = ttk.Notebook(body, style="Review.TNotebook")
        notebook.grid(row=2, column=0, sticky="nsew")
        self.tree = self.make_tree(notebook, "项目预览", ("项目", "填入行", "分值", "处理结果"), (280, 70, 60, 220))
        self.changes = self.make_tree(notebook, "修改前后", ("单元格", "原内容", "填入内容"), (80, 300, 360))
        self.target_preview = self.make_tree(notebook, "Excel 预览", ("行", "内容"), (50, 750))
        source_frame = ttk.Frame(notebook, style="Toolbar.TFrame")
        notebook.add(source_frame, text="招标原文")
        self.raw = self.make_text(source_frame)
        self.raw.pack(fill="both", expand=True)
        detail_frame = ttk.Frame(self.tree.master, style="Toolbar.TFrame")
        detail_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.detail = self.make_text(detail_frame, height=4)
        self.detail.pack(side="left", fill="both", expand=True)
        from ui_common import FlatScrollbar
        detail_scroll = FlatScrollbar(detail_frame, orient="vertical", command=self.detail.yview,
                                      style="Flat.Vertical.TScrollbar")
        detail_scroll.pack(side="right", fill="y")
        self.detail.configure(yscrollcommand=detail_scroll.set)
        self.tree.bind("<<TreeviewSelect>>", self.show_criterion)
        self.tree.bind("<Double-1>", lambda event: self.assign_row())
        self.changes.bind("<Double-1>", self.show_change)
        footer = ttk.Frame(body, style="App.TFrame")
        footer.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        footer.columnconfigure(0, weight=1)
        self.status = tk.StringVar(value="")
        self.status_label = ttk.Label(footer, textvariable=self.status, style="Subtitle.TLabel", wraplength=560)
        self.status_label.grid(row=0, column=0, sticky="w")
        self.generate = ttk.Button(footer, text="确认填入并另存", style="Accent.TButton", command=self.write, state="disabled")
        if preview_receiver is not None:
            self.generate.configure(text="应用到右侧预览")
        self.generate.grid(row=0, column=1, sticky="e", padx=(12, 0))
        self.window.bind("<Configure>", lambda e: self.status_label.configure(wraplength=max(300, self.window.winfo_width()-260)) if e.widget == self.window else None)
        for var in [*self.source_vars.values(), *self.target_vars.values(), self.source_header,
                    self.header, self.start, self.end, self.sheet, self.mode, self.include_price]:
            var.trace_add("write", self.invalidate)
        self.source.bind("<<ComboboxSelected>>", self.source_changed)
        self.sheet_box.bind("<<ComboboxSelected>>", self.target_changed)
        self.source.current(0)
        self.source_changed()
        # 优先选择识别出评分列的工作表，避免默认停在“概要”。
        def rank(name):
            info = sheets[name]
            header = info["guess"]["header"]
            hits = sum(any(_norm(value) in aliases for ref, value in info["cells"].items()
                           if range_boundaries(ref)[1] == header) for aliases in HEADERS)
            return hits, "评分" in name or "评审" in name
        self.sheet.set(max(sheets, key=rank))
        self.target_changed()
        self.window.after_idle(lambda: self.preview(silent=True))
        from ui_common import center_window_on_parent
        center_window_on_parent(self.window, self.app.root, 1000, 740)
        self.window.grab_set()

    def make_text(self, parent, height=8):
        return tk.Text(parent, wrap="word", height=height, font=self.app.body_font,
                       bg=self.app.surface_bg, fg=self.app.text_fg, relief="flat", padx=12, pady=8,
                       selectbackground="#C6D8EB", insertbackground=self.app.text_fg)

    def make_tree(self, notebook, title, columns, widths):
        frame = ttk.Frame(notebook, style="Toolbar.TFrame")
        notebook.add(frame, text=title)
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse", height=5, style="Scan.Treeview")
        for column, width in zip(columns, widths):
            tree.heading(column, text=column)
            tree.column(column, width=width, minwidth=40)
        from ui_common import FlatScrollbar
        scrollbar = FlatScrollbar(frame, orient="vertical", command=tree.yview, style="Flat.Vertical.TScrollbar")
        horizontal = FlatScrollbar(frame, orient="horizontal", command=tree.xview, style="Flat.Horizontal.TScrollbar")
        tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=horizontal.set)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        return tree

    def show_criterion(self, _event=None):
        selected = self.tree.selection()
        self.assign_button.configure(state="normal" if selected and self.plan else "disabled")
        self.skip_button.configure(state="normal" if selected and self.plan else "disabled")
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        if selected and self.plan:
            entry = self.plan["entries"][int(selected[0])]
            self.detail.insert("1.0", entry["description"])
        self.detail.configure(state="disabled")

    def update_summary(self):
        self.mapping_summary.set(f"填写第 {self.start.get()}–{self.end.get()} 行 · 名称 {self.target_vars['factor'].get()} 列 / 说明 {self.target_vars['description'].get()} 列 / 分值 {self.target_vars['points'].get()} 列")

    def edit_mapping(self):
        popup = tk.Toplevel(self.window)
        popup.title("调整列和填写范围")
        popup.configure(bg=self.app.app_bg)
        popup.transient(self.window)
        popup.resizable(False, False)
        outer = ttk.Frame(popup, padding=14, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        body = ttk.Frame(outer, padding=14, style="Surface.TFrame")
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="每项内容从哪里取，填到哪里", style="SectionTitle.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))
        ttk.Label(body, text="招标表中的列").grid(row=1, column=1, sticky="w")
        ttk.Label(body, text="Excel 中的列").grid(row=1, column=2, sticky="w")
        table = self.items[self.source.current()].table
        info = self.sheets[self.sheet.get()]
        for index, (key, label) in enumerate(zip(FIELDS, ("项目名称", "评分说明", "分值")), 2):
            ttk.Label(body, text=label).grid(row=index, column=0, sticky="w", padx=(0, 14), pady=6)
            for source, var, column in ((True, self.source_vars[key], 1), (False, self.target_vars[key], 2)):
                names = {}
                count = table.column_count if source else max(info["max_col"], 10)
                try:
                    header = int(self.source_header.get() if source else self.header.get())
                except ValueError:
                    header = 1
                for c in range(1, count+1):
                    title = next((cell.text for cell in table.cells if cell.row <= header < cell.row+cell.row_span and cell.column == c), "") if source else info["cells"].get(f"{get_column_letter(c)}{header}", "")
                    number = str(c) if source else get_column_letter(c)
                    names[number] = f"{number} · {str(title).strip()[:20]}" if title else number
                box = ttk.Combobox(body, values=list(names.values()), state="readonly", width=25)
                box.set(names.get(var.get(), var.get()))
                box.grid(row=index, column=column, padx=6)
                box.bind("<<ComboboxSelected>>", lambda _, v=var, b=box: v.set(b.get().split(" · ")[0]))
        for r, (label, variable) in enumerate((("招标表的标题在第几行", self.source_header),
                                             ("Excel 的标题在第几行", self.header),
                                             ("从第几行开始填写", self.start), ("填到第几行结束", self.end)), 5):
            ttk.Label(body, text=label).grid(row=r, column=0, columnspan=2, sticky="w", pady=5)
            ttk.Entry(body, textvariable=variable, width=10).grid(row=r, column=2, sticky="w", padx=6)
        def close():
            popup.destroy()
            self.window.grab_set()
            self.preview()
        buttons = ttk.Frame(outer, style="App.TFrame")
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="自动增加标准模板行数…", command=lambda: (popup.destroy(), self.standard_fill())).pack(side="left")
        ttk.Button(buttons, text="完成并预览", style="Accent.TButton", command=close).pack(side="right")
        popup.protocol("WM_DELETE_WINDOW", lambda: (popup.destroy(), self.window.grab_set()))
        from ui_common import center_window_on_parent
        center_window_on_parent(popup, self.window)
        popup.grab_set()

    def invalidate(self, *_):
        self.plan = None
        self.overrides.clear()
        self.generate.configure(state="disabled")
        self.status.set("设置已更新，请点击“预览填入结果”。")
        self.update_summary()
        self.assign_button.configure(state="disabled")
        self.skip_button.configure(state="disabled")

    def source_changed(self, *_):
        self.invalidate()
        self.source_header.set("1")
        for index, key in enumerate(FIELDS, 1):
            self.source_vars[key].set(str(index))
        table = self.items[self.source.current()].table
        self.raw.configure(state="normal")
        self.raw.delete("1.0", "end")
        for cell in table.cells:
            self.raw.insert("end", f"第{cell.row}行 第{cell.column}列：{cell.text}\n")
        self.raw.configure(state="disabled")
        for row in range(1, min(table.row_count, 8)+1):
            found = [[cell.column for cell in table.cells if cell.row == row and _norm(cell.text) in aliases] for aliases in HEADERS]
            if all(len(cols) == 1 for cols in found):
                self.source_header.set(str(row))
                for key, cols in zip(FIELDS, found):
                    self.source_vars[key].set(str(cols[0]))
                break

    def target_changed(self, *_):
        info = self.sheets[self.sheet.get()]
        for key, var in self.target_vars.items():
            var.set(info["guess"][key])
        for key in ("header", "start", "end"):
            getattr(self, key).set(str(info["guess"][key]))
        self.render_target(info)

    def render_target(self, info, updates=None, start=1, end=None):
        updates = updates or {}
        self.target_preview.delete(*self.target_preview.get_children())
        count = min(info["max_col"], 100)
        columns = ["row"] + [get_column_letter(c) for c in range(1, count+1)]
        self.target_preview.configure(columns=columns)
        self.target_preview.heading("row", text="行")
        self.target_preview.column("row", width=45, stretch=False)
        for column in columns[1:]:
            title = info["cells"].get(f"{column}{self.header.get()}", "")
            self.target_preview.heading(column, text=f"{column} · {title}" if title else column)
            self.target_preview.column(column, width=180, minwidth=70, stretch=False)
        for row in range(start, min(end if end is not None else info["max_row"], start+999)+1):
            values = [row]
            for column in columns[1:]:
                ref = f"{column}{row}"
                value = updates.get(ref, info["cells"].get(ref, ""))
                values.append(("＊" if ref in updates else "") + str(value).replace("\n", " "))
            self.target_preview.insert("", "end", values=values)

    def configs(self):
        source = {k: int(v.get()) for k, v in self.source_vars.items()}
        source["header"] = int(self.source_header.get())
        target = {k: v.get().strip().upper() for k, v in self.target_vars.items()}
        target.update(sheet=self.sheet.get(), mode=self.mode.get(), include_price=self.include_price.get(),
                      **{k: int(getattr(self, k).get()) for k in ("header", "start", "end")})
        return source, target

    def preview(self, silent=False):
        try:
            source, target = self.configs()
            self.sheets = inspect_workbook(self.template)
            rules = extract_review_rules(self.items[self.source.current()].table, source, self.include_price.get())
            self.plan = plan_fill(self.template, rules, target, self.overrides)
            self.tree.delete(*self.tree.get_children())
            for entry in self.plan["entries"]:
                self.tree.insert("", "end", iid=str(entry["index"]), values=(entry["name"], entry["row"] if entry["row"] is not None else "", entry["points"], entry["status"]))
            self.changes.delete(*self.changes.get_children())
            for index, change in enumerate(self.plan["changes"]):
                self.changes.insert("", "end", iid=str(index), values=(change["cell"], change["old"], change["new"]))
            info = self.sheets[target["sheet"]]
            updates = {c["cell"]: c["new"] for c in self.plan["changes"]}
            self.render_target(info, updates, target["header"], target["end"] + 2)
            if self.tree.get_children():
                self.tree.selection_set(self.tree.get_children()[0])
                self.show_criterion()
            errors = self.plan["errors"]
            ready = sum(e["row"] not in (None, 0) for e in self.plan["entries"])
            self.status.set(f"{len(rules.criteria)} 个评分组，共 {rules.total} 分；" +
                            (f"{len(errors)} 项需要指定行或跳过。" if errors else
                             f"将填入 {ready} 项，修改 {len(self.plan['changes'])} 个单元格。"))
            self.generate.configure(state="normal" if not errors and self.plan["changes"] else "disabled")
        except (ValueError, OSError, KeyError) as exc:
            self.plan = None
            self.generate.configure(state="disabled")
            self.status.set(str(exc) + " 请点击“调整列和行”。")
            if not silent:
                messagebox.showwarning("请调整列和行", str(exc), parent=self.window)

    def assign_row(self, skip=False):
        selected = self.tree.selection()
        if not selected or self.plan is None:
            return
        index = int(selected[0])
        row = 0 if skip else simpledialog.askinteger("指定目标行", "填写 Excel 行号（0 表示跳过）：", parent=self.window, minvalue=0, maxvalue=1048576)
        if row is not None:
            self.overrides[index] = row
            self.preview()
            self.tree.selection_set(str(index))

    def reset_rows(self):
        self.overrides.clear()
        self.preview()

    def show_change(self, _event=None):
        selected = self.changes.selection()
        if selected and self.plan:
            change = self.plan["changes"][int(selected[0])]
            popup = tk.Toplevel(self.window)
            popup.configure(bg=self.app.app_bg)
            popup.title(change["cell"])
            text = self.make_text(popup, height=25)
            text.pack(fill="both", expand=True)
            text.insert("1.0", f"原内容：\n{change['old']}\n\n填入内容：\n{change['new']}")
            text.configure(state="disabled")

    def store_profile(self):
        try:
            source, target = self.configs()
            rules = extract_review_rules(self.items[self.source.current()].table, source, self.include_price.get())
            plan_fill(self.template, rules, target)
            initial_name = self.profile.get() if self.profile.get() in self.profiles else os.path.splitext(os.path.basename(self.template))[0]
            name = simpledialog.askstring("记住填表设置", "给这套填表方式起个名字：", initialvalue=initial_name, parent=self.window)
            if name:
                save_profile(name, target, source, self.sheets[target["sheet"]])
                self.profiles = load_profiles()
                self.profile.configure(values=list(self.profiles))
                self.profile.set(name.strip())
                self.status.set("配置已保存，下次可从配置列表选择；项目行对应每次重新匹配。")
        except (ValueError, OSError, KeyError) as exc:
            messagebox.showwarning("配置未保存", str(exc), parent=self.window)

    def restore_profile(self, *_):
        try:
            profile = self.profiles[self.profile.get()]
            config = profile["config"]
            if config["sheet"] not in self.sheets:
                raise ValueError("本文件没有配置中的工作表，请重新选择并保存配置。")
            self.sheet.set(config["sheet"])
            self.target_changed()
            for key, var in self.target_vars.items():
                var.set(config[key])
            for key in ("header", "start", "end", "mode", "include_price"):
                getattr(self, key).set(config[key])
            self.source_header.set(profile["source"]["header"])
            for key, var in self.source_vars.items():
                var.set(profile["source"][key])
            if layout_signature(self.sheets[config["sheet"]], config) != profile["signature"]:
                messagebox.showwarning("模板布局已变化", "已载入列对应；表头或合并布局已变化，请核对范围后重新预览。", parent=self.window)
            self.status.set("已载入配置，请核对来源列并匹配预览。")
        except (KeyError, ValueError, TypeError) as exc:
            messagebox.showwarning("配置不适用", str(exc), parent=self.window)

    def change_file(self):
        self.window.destroy()
        self.owner.start_review_fill(choose_file=True)

    def standard_fill(self):
        from review_table_filler import inspect_review_targets
        try:
            source, _ = self.configs()
            item = self.items[self.source.current()]
            rules = extract_review_rules(item.table, source, self.include_price.get())
            targets = inspect_review_targets(self.template)
            if not targets:
                raise ValueError("没有可自动扩行的标准模板。请使用通用填入，或先在 Excel 中扩充模板行。")
            self.window.destroy()
            self.owner._show_review_fill_dialog(self.template, [(item, rules)], targets)
        except ValueError as exc:
            messagebox.showwarning("无法自动扩行", str(exc), parent=self.window)

    def write(self):
        if self.plan is None or self.owner._task_runner.active or self.owner.app._task_runner.active:
            return
        if self.preview_receiver is not None:
            if self.preview_receiver(self.plan, self.items[self.source.current()], self.configs()):
                self.window.destroy()
            return
        from word_scan import validate_scan_files
        plan = self.plan
        item = self.items[self.source.current()]
        stamps = dict(self.owner.scan_stamps)
        output = review_output_path(self.template, self.owner.output_dir)
        self.window.destroy()
        self.owner._begin_indeterminate_task("正在填入开评标表…")
        def work(_publish):
            validate_scan_files([item.file_path], stamps)
            return apply_plan(self.template, output, plan)
        def done(result):
            self.owner._reset_busy_state()
            self.owner._last_output_dir_to_open = os.path.dirname(output)
            self.owner.status_var.set(f"已填入 {result['count']} 个单元格")
            self.owner._show_result_window(f"工作表：{result['sheet_name']}\n修改 {result['count']} 个单元格\n输出：{output}", title="填表结果")
        self.owner._task_runner.submit(work, done, lambda exc: self.owner._finish_background_error(exc, "填入开评标表"))
