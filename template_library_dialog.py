"""Dialog for maintaining and selecting a categorized template library."""

from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, ttk

from app_dialogs import messagebox
from PIL import ImageTk
from ui_icons import checkbox_icon
from ui_common import ElidedTextController, FlatScrollbar
import template_library as library


class TemplateLibraryDialog(tk.Toplevel):
    def __init__(self, parent, root_path, on_root_changed=None, on_use=None):
        super().__init__(parent)
        self.title("模板库")
        self.root_path = root_path
        self.on_root_changed = on_root_changed
        self.on_use = on_use
        self.selected_templates = []
        self._templates = []
        self._checked_templates = set()
        self.transient(parent)
        self.geometry("860x540")
        self.minsize(760, 460)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        style = ttk.Style(self)
        surface = style.lookup("Toolbar.TFrame", "background") or "#F7F8FA"
        foreground = style.lookup("TLabel", "foreground") or "#2B2F36"
        font = tkfont.Font(self, font=style.lookup("TLabel", "font") or ("Microsoft YaHei UI", 10))
        self._list_font = font
        self.configure(background=surface)
        header = ttk.Frame(self, padding=(16, 12), style="Toolbar.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        self.root_label = ttk.Label(header, style="Muted.TLabel", width=1)
        self.root_label.grid(row=0, column=0, sticky="ew", padx=(0, 16))
        self._root_display = ElidedTextController(
            self.root_label, font, lambda: self.root_label.winfo_width(), tooltip=True
        )
        self.root_label.bind("<Configure>", lambda _event: self._root_display.refresh())
        ttk.Button(header, text="更改目录…", command=self.change_root).grid(row=0, column=1, sticky="e")

        body = ttk.Frame(self, padding=(16, 0, 16, 12), style="Toolbar.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1, minsize=240)
        body.columnconfigure(1, weight=3, minsize=360)
        body.rowconfigure(1, weight=1)

        category_panel = ttk.Frame(body, style="Toolbar.TFrame")
        category_panel.grid(row=0, rowspan=2, column=0, sticky="nsew", padx=(0, 10))
        category_panel.columnconfigure(0, weight=1)
        category_panel.rowconfigure(1, weight=1, uniform="content")
        ttk.Label(category_panel, text="分类", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 5))
        def make_list(parent, selectmode):
            border = tk.Frame(parent, bg="#C8CED6", bd=0)
            border.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
            border.columnconfigure(0, weight=1)
            border.rowconfigure(0, weight=1)
            listing = tk.Listbox(
                border, activestyle="none", exportselection=False,
                selectmode=selectmode, font=font, bg=surface, fg=foreground,
                selectbackground="#1A6BB5", selectforeground="#FFFFFF",
                relief="flat", bd=0, highlightthickness=0, width=1,
            )
            listing.grid(row=0, column=0, sticky="nsew", padx=(1, 0), pady=(1, 0))
            vertical = FlatScrollbar(border, command=listing.yview, style="Flat.Vertical.TScrollbar")
            vertical.grid(row=0, column=1, sticky="ns", pady=1, padx=(0, 1))
            horizontal = FlatScrollbar(border, command=listing.xview, orient="horizontal", style="Flat.Horizontal.TScrollbar")
            horizontal.grid(row=1, column=0, sticky="ew", padx=1, pady=(0, 1))
            listing.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
            return listing

        self.category_list = make_list(category_panel, "browse")

        category_actions = ttk.Frame(category_panel, style="Toolbar.TFrame")
        category_actions.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 9))
        category_actions.columnconfigure(0, weight=1)
        self.category_entry = ttk.Entry(category_actions, width=12)
        self.category_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.category_entry.bind("<Return>", lambda _event: self.create_category())
        ttk.Button(category_actions, text="新增", command=self.create_category, width=7).grid(row=0, column=1)
        category_management = ttk.Frame(category_actions, style="Toolbar.TFrame")
        category_management.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.rename_button = ttk.Button(category_management, text="重命名…", command=self.rename_category)
        self.rename_button.pack(side="left", padx=(0, 8))
        self.delete_category_button = ttk.Button(category_management, text="删除空分类", command=self.delete_category)
        self.delete_category_button.pack(side="left")

        template_panel = ttk.Frame(body, style="Toolbar.TFrame")
        template_panel.grid(row=0, rowspan=2, column=1, sticky="nsew")
        template_panel.columnconfigure(0, weight=1)
        template_panel.rowconfigure(1, weight=1)
        template_header = ttk.Frame(template_panel, style="Toolbar.TFrame")
        template_header.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 5))
        ttk.Label(template_header, text="模板文件", style="SectionTitle.TLabel").pack(side="left")
        self.select_all_var = tk.BooleanVar(self, value=True)
        self.select_all_button = ttk.Checkbutton(
            template_header, text="全选", variable=self.select_all_var,
            command=self.toggle_all_templates,
        )
        self.select_all_button.pack(side="right")
        template_border = ttk.Frame(template_panel, style="Toolbar.TFrame")
        template_border.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        template_border.columnconfigure(0, weight=1)
        template_border.rowconfigure(0, weight=1)
        style.configure("Template.Treeview", font=font, rowheight=max(28, font.metrics("linespace") + 10),
                        background=surface, fieldbackground=surface, foreground=foreground)
        self.template_list = ttk.Treeview(template_border, show="tree", selectmode="browse", style="Template.Treeview")
        self.template_list.column("#0", width=420, minwidth=120, stretch=True)
        self.template_list.grid(row=0, column=0, sticky="nsew")
        vertical = FlatScrollbar(template_border, command=self.template_list.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = FlatScrollbar(template_border, command=self.template_list.xview, orient="horizontal")
        horizontal.grid(row=1, column=0, sticky="ew")
        self.template_list.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        icon_size = max(16, round(18 * self.winfo_fpixels("1i") / 96))
        self._check_images = {
            checked: ImageTk.PhotoImage(checkbox_icon(icon_size, "#1A6BB5", state), master=self)
            for checked, state in ((True, "checked"), (False, "unchecked"))
        }
        self._mixed_check_image = ImageTk.PhotoImage(
            checkbox_icon(icon_size, "#1A6BB5", "mixed"), master=self
        )
        indicator_name = f"TemplateSelectAll{id(self)}.indicator"
        check_style = f"TemplateSelectAll{id(self)}.TCheckbutton"
        style.element_create(
            indicator_name, "image", self._check_images[False],
            ("alternate", self._mixed_check_image), ("selected", self._check_images[True]),
            sticky="w",
        )
        style.layout(check_style, [("Checkbutton.padding", {"sticky": "nswe", "children": [
            (indicator_name, {"side": "left", "sticky": ""}),
            ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [
                ("Checkbutton.label", {"sticky": "nswe"}),
            ]}),
        ]})])
        self.select_all_button.configure(style=check_style)
        self.template_list.bind("<Button-1>", self.toggle_template)
        self.template_list.bind("<space>", self.toggle_template)
        self.template_list.bind("<Control-a>", self.select_all_templates)

        template_actions = ttk.Frame(template_panel, style="Toolbar.TFrame")
        template_actions.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 9))
        self.add_button = ttk.Button(template_actions, text="添加模板…", command=self.add_templates)
        self.add_button.pack(side="left", padx=(0, 8))
        self.delete_button = ttk.Button(template_actions, text="删除所选", command=self.delete_template)
        self.delete_button.pack(side="left")
        category_panel.bind("<Configure>", lambda _event: template_actions.configure(height=category_actions.winfo_reqheight()))
        template_actions.pack_propagate(False)

        footer = ttk.Frame(self, padding=(16, 0, 16, 16), style="Toolbar.TFrame")
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        self.selection_label = ttk.Label(footer, style="Hint.TLabel")
        self.selection_label.grid(row=0, column=0, sticky="w")
        self.use_button = ttk.Button(footer, text="使用所选模板", style="Accent.TButton", command=self.use_selected)
        self.use_button.grid(row=0, column=1, sticky="e")

        self.category_list.bind("<<ListboxSelect>>", lambda _event: self.refresh_templates())
        self.bind("<Escape>", lambda _event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.refresh_categories(select_first=True)

    def _refresh_root_label(self):
        self._root_display.set_text(self.root_path)

    def _selected_category(self):
        selection = self.category_list.curselection()
        return self.category_list.get(selection[0]) if selection else ""

    def refresh_categories(self, select_first=False, selected=None):
        try:
            categories = library.list_categories(self.root_path)
        except OSError as exc:
            categories = []
            messagebox.showerror("模板库", f"读取分类失败：{exc}", parent=self)
        self.category_list.delete(0, "end")
        self.category_list.insert("end", *categories)
        index = 0
        if selected in categories:
            index = categories.index(selected)
        elif not categories:
            index = None
        if index is not None:
            self.category_list.selection_clear(0, "end")
            self.category_list.selection_set(index)
            self.category_list.see(index)
        self.refresh_templates()
        self._refresh_root_label()

    def refresh_templates(self):
        self.template_list.delete(*self.template_list.get_children())
        self._templates = []
        self._checked_templates.clear()
        category = self._selected_category()
        if not category:
            self.selected_templates = []
            self._update_actions()
            return
        self._templates = library.list_templates(self.root_path, [category])
        for index, item in enumerate(self._templates):
            row = str(index)
            self._checked_templates.add(row)
            self.template_list.insert("", "end", iid=row, text=item.name, image=self._check_images[True])
        self.template_list.column("#0", width=max(
            [420, *(self._list_font.measure(item.name) + 60 for item in self._templates)]
        ))
        self._on_template_selection()

    def toggle_template(self, event):
        row = self.template_list.focus() if event.keysym == "space" else self.template_list.identify_row(event.y)
        if row:
            self.template_list.focus_set()
            self.template_list.focus(row)
            if row in self._checked_templates:
                self._checked_templates.remove(row)
            else:
                self._checked_templates.add(row)
            self.template_list.item(row, image=self._check_images[row in self._checked_templates])
            self._on_template_selection()
        return "break"

    def toggle_all_templates(self):
        self._checked_templates = set(self.template_list.get_children()) if self.select_all_var.get() else set()
        for row in self.template_list.get_children():
            self.template_list.item(row, image=self._check_images[row in self._checked_templates])
        self._on_template_selection()

    def select_all_templates(self, _event=None):
        self.select_all_var.set(True)
        self.toggle_all_templates()
        return "break"

    def create_category(self):
        try:
            category = library.normalize_category_name(self.category_entry.get())
            library.create_category(self.root_path, category)
        except (OSError, ValueError) as exc:
            messagebox.showwarning("新增分类", str(exc), parent=self)
            return
        self.category_entry.delete(0, "end")
        self.refresh_categories(selected=category)

    def rename_category(self):
        old_category = self._selected_category()
        if not old_category:
            messagebox.showinfo("重命名分类", "请先选择分类。", parent=self)
            return
        dialog = tk.Toplevel(self)
        dialog.title("重命名分类")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.configure(background=ttk.Style(self).lookup("Toolbar.TFrame", "background"))
        frame = ttk.Frame(dialog, padding=16, style="Toolbar.TFrame")
        frame.pack(fill="both")
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text="新分类名称").grid(row=0, column=0, sticky="w")
        entry = ttk.Entry(frame)
        entry.grid(row=1, column=0, sticky="ew", pady=(5, 12))
        entry.insert(0, old_category)
        entry.focus_set()
        entry.selection_range(0, "end")

        def confirm():
            try:
                new_category = library.rename_category(self.root_path, old_category, entry.get())
            except (OSError, ValueError) as exc:
                messagebox.showwarning("重命名分类", str(exc), parent=dialog)
                return
            dialog.destroy()
            self.refresh_categories(selected=library.normalize_category_name(new_category))

        ttk.Button(frame, text="确定", command=confirm).grid(row=2, column=0, sticky="e")
        entry.bind("<Return>", lambda _event: confirm())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def delete_category(self):
        category = self._selected_category()
        if not category:
            messagebox.showinfo("删除分类", "请先选择分类。", parent=self)
            return
        if not messagebox.askyesno("删除分类", f"确定删除空分类“{category}”？", parent=self):
            return
        try:
            library.delete_category(self.root_path, category)
        except (OSError, ValueError) as exc:
            messagebox.showwarning("删除分类", str(exc), parent=self)
            return
        self.refresh_categories(select_first=True)

    def add_templates(self):
        category = self._selected_category()
        if not category:
            messagebox.showinfo("添加模板", "请先选择分类。", parent=self)
            return
        paths = filedialog.askopenfilenames(
            parent=self,
            title=f"添加模板到“{category}”",
            filetypes=[
                ("Office 模板", "*.doc;*.docx;*.xls;*.xlsx;*.xlsm;*.ppt;*.pptx"),
                ("所有文件", "*.*"),
            ],
        )
        if not paths:
            return
        try:
            added = library.add_templates(self.root_path, category, paths)
        except OSError as exc:
            messagebox.showerror("添加模板", f"复制模板失败：{exc}", parent=self)
            return
        if not added:
            messagebox.showwarning("添加模板", "未添加支持的 Office 模板。", parent=self)
            return
        self.refresh_templates()
        messagebox.showinfo("添加模板", f"已添加 {len(added)} 个模板。", parent=self)

    def delete_template(self):
        category = self._selected_category()
        selections = [index for index in range(len(self._templates)) if str(index) in self._checked_templates]
        if not category or not selections:
            messagebox.showinfo("删除模板", "请先选择模板。", parent=self)
            return
        names = [self._templates[index].name for index in selections]
        if not messagebox.askyesno("删除模板", "确定删除所选模板？\n\n" + "\n".join(names), parent=self):
            return
        templates = {item.name: item for item in library.list_templates(self.root_path, [category])}
        for name in names:
            try:
                library.delete_template(self.root_path, templates[name])
            except (KeyError, OSError, ValueError) as exc:
                messagebox.showerror("删除模板", f"删除失败：{name}\n{exc}", parent=self)
                self.refresh_templates()
                return
        self.refresh_templates()

    def change_root(self):
        directory = filedialog.askdirectory(parent=self, title="选择模板库目录", mustexist=True)
        if not directory:
            return
        try:
            self.root_path = library.ensure_library(directory)
        except OSError as exc:
            messagebox.showerror("模板库", f"打开目录失败：{exc}", parent=self)
            return
        if self.on_root_changed:
            self.on_root_changed(self.root_path)
        self.refresh_categories(select_first=True)

    def _on_template_selection(self, _event=None):
        self.selected_templates = [item.path for index, item in enumerate(self._templates) if str(index) in self._checked_templates]
        self._update_actions()

    def _update_actions(self):
        category = self._selected_category()
        self.select_all_var.set(bool(self._templates) and len(self.selected_templates) == len(self._templates))
        self.select_all_button.state(["!disabled"] if self._templates else ["disabled"])
        self.select_all_button.state(["alternate"] if 0 < len(self.selected_templates) < len(self._templates) else ["!alternate"])
        for button in (self.rename_button, self.delete_category_button, self.add_button):
            button.state(["!disabled"] if category else ["disabled"])
        for button in (self.delete_button, self.use_button):
            button.state(["!disabled"] if self.selected_templates else ["disabled"])
        self.selection_label.configure(text=(
            "先新增采购方式或采购人分类" if not category else
            "此分类还没有模板" if not self._templates else
            f"已勾选 {len(self.selected_templates)} / {len(self._templates)} 个模板"
        ))

    def use_selected(self):
        if not self.selected_templates:
            messagebox.showinfo("使用模板", "请先选择模板。", parent=self)
            return
        if self.on_use:
            self.on_use(self.selected_templates)
        self.destroy()
