"""图标覆盖层不得改变勾选、Shift 范围选择或滚动后的命中行。"""
import tkinter as tk
import unittest
from unittest.mock import patch

import main
from word_table_exporter import TableScanItem


class CheckmarkInteractionTests(unittest.TestCase):
    def test_click_shift_scroll_and_highlight_preserve_selection_semantics(self):
        root=tk.Tk()
        self.addCleanup(root.destroy)
        errors=[]
        root.report_callback_exception=lambda *args:errors.append(args)
        app=main.ReplaceSimpleApp(root,restore_session=False)
        app.show_workspace('extract');root.geometry('1200x900')
        owner=app.table_export_window;owner.file_paths=['source.docx']
        items=[TableScanItem('source.docx','source.docx',i,'采购需求','','预览',1,1,'') for i in range(1,101)]
        owner._show_scan_result(items,{},None,[],{},None)
        root.update()
        tree=owner.scan_tree;marks=owner._scan_checkmarks
        tree.selection_set('2');root.update()

        def click(iid, state=0):
            tree.see(iid);root.update()
            box=tree.bbox(iid,'selected')
            layer=next(layer for layer in marks.layers if layer.winfo_ismapped()
                       and layer.winfo_y()==tree.winfo_y()+box[1])
            layer.event_generate('<ButtonPress-1>',x=layer.winfo_width()//2,y=layer.winfo_height()//2,state=state)
            layer.event_generate('<ButtonRelease-1>',x=layer.winfo_width()//2,y=layer.winfo_height()//2,state=state)
            root.update()

        with patch.object(tree,'delete',wraps=tree.delete) as deleted:
            click('1')
            deleted.assert_not_called()
        self.assertEqual(owner.selected_scan_keys,{owner._scan_key('table',items[0])})
        self.assertEqual(tree.selection(),('2',))
        click('3',state=1)
        self.assertEqual(owner.selected_scan_keys,{owner._scan_key('table',item) for item in items[:3]})
        click('100')
        self.assertIn(owner._scan_key('table',items[-1]),owner.selected_scan_keys)
        self.assertLess(len(marks.layers),len(items))
        owner._toggle_visible_from_heading();root.update()
        self.assertEqual(len(owner.selected_scan_keys),100)
        owner._toggle_visible_from_heading();root.update()
        self.assertFalse(owner.selected_scan_keys)
        # 放宽勾选列后，表头与行内图标仍在同一中心，边线不被图标覆盖。
        tree.column('selected',width=150)
        owner._schedule_scan_tree_column_separators();root.update()
        row_layer=next(layer for layer in marks.layers if layer.winfo_ismapped())
        def center(layer):return layer.winfo_rootx()+layer.coords(layer._icon_item)[0]
        self.assertAlmostEqual(center(row_layer),center(marks.header_layer))
        line=owner._scan_column_separators[0]
        stacking=list(tree.master.winfo_children())
        self.assertGreater(stacking.index(line),stacking.index(row_layer))
        # 重复勾选同一表格不重建右侧预览。
        owner._update_scan_detail(iid='100')
        with patch.object(owner,'_show_preview_table',wraps=owner._show_preview_table) as preview:
            for _ in range(5):click('100')
            root.after(80,root.quit);root.mainloop()
            preview.assert_not_called()
        # “只看已选”仍能在取消勾选后移除对应行。
        owner.only_selected_var.set(True);owner._apply_scan_filters();root.update()
        self.assertEqual(tree.get_children(),('100',))
        click('100')
        self.assertEqual(tree.get_children(),())
        self.assertEqual(errors,[])

    def test_filter_dropdown_marks_use_same_drawn_icons_as_tree(self):
        root=tk.Tk()
        self.addCleanup(root.destroy)
        app=main.ReplaceSimpleApp(root,restore_session=False)
        app.show_workspace('extract')
        owner=app.table_export_window;owner.file_paths=['source.docx']
        items=[TableScanItem('source.docx','source.docx',1,'第一章','采购需求','预览',1,1,''),
               TableScanItem('source.docx','source.docx',2,'第二章','采购需求','预览',1,1,'')]
        owner._show_scan_result(items,{},None,[],{},None);root.update()
        owner.filter_boxes['section'].invoke();root.update()
        rows=dict(owner._filter_popup_rows)
        # 勾选框用绘制图标（与表格列表同源），不再是 ☑/☐ 字体字符。
        for value,mark in rows.items():
            self.assertTrue(mark.cget('image'),value)
            self.assertEqual(str(mark.cget('text')),'',value)
        values,_labels=owner._filter_choice_pairs('section');available=list(values)
        self.assertEqual(str(rows[values[0]].cget('image')),str(owner._filter_mark_image('☑')))
        owner._toggle_filter_choice('section',values[0],available);root.update()
        self.assertEqual(str(rows[values[0]].cget('image')),str(owner._filter_mark_image('☐')))
        self.assertEqual(str(rows[None].cget('image')),str(owner._filter_mark_image('◩')))
        owner._close_filter_dropdown()

    def test_section_filter_strips_file_name_prefix(self):
        root=tk.Tk()
        self.addCleanup(root.destroy)
        app=main.ReplaceSimpleApp(root,restore_session=False)
        app.show_workspace('extract')
        owner=app.table_export_window;owner.file_paths=['C:/a.docx','C:/b.docx']
        prefix='2641STC74317_招标文件（第2册）.docx / '
        items=[TableScanItem('C:/a.docx','a.docx',1,prefix+'第一章 采购公告','','预览',1,1,''),
               TableScanItem('C:/a.docx','a.docx',2,'第一章 采购公告','','预览',1,1,''),
               TableScanItem('C:/b.docx','b.docx',1,prefix+'第一章 采购公告','','预览',1,1,''),
               TableScanItem('C:/b.docx','b.docx',2,'未识别章节','','预览',1,1,'')]
        owner._show_scan_result(items,{},None,[],{},None);root.update()
        # 树内章节列与筛选值都去掉文件名前缀。
        self.assertEqual(owner.scan_tree.set('1','section'),'第一章 采购公告')
        self.assertEqual(owner._scan_filter_value('section','table',items[0]),'第一章 采购公告')
        self.assertEqual(owner._scan_filter_value('section','table',items[3]),'未识别章节')
        owner.filter_boxes['section'].invoke();root.update()
        labels=[mark.master.winfo_children()[1].cget('text') for _v,mark in owner._filter_popup_rows]
        # 前缀去重后每个章节只出现一次，不再有 .docx 记号。
        self.assertEqual(labels,['全部','第一章 采购公告','未识别章节'])
        owner._close_filter_dropdown()
        # 按章节筛选仍能命中两个文件里的同一章。
        owner._set_multi_filter('section',{'第一章 采购公告'});root.update()
        self.assertEqual(len(owner.scan_tree.get_children()),2)


if __name__=='__main__':unittest.main()
