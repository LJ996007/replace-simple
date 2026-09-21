# 新增「带符号指标参数提取」模块（融合进 提取Word表格 窗口）

未确认的三项设计决策按推荐方案执行：**整篇文档扫描**（★▲△# 几乎只用于指标条款，避免章节识别失败漏提）、**每个符号单独一行**、**每个 Word 单独生成 `{文件名}_指标参数.xlsx`**。

## 1. 新建核心模块 `symbol_clause_extractor.py`（无 UI，风格同 word_table_exporter.py）

**符号识别**
- 符号集常量 `SYMBOL_CHARS = "★☆▲△▽▼◆◇●○■□※◎＃#＊*"`（后续可扩展）。
- 位置规则（按需求）：符号必须在条款最开始，或紧跟条款序号之后（序号后可带 （【( 等括号）。用 `^` 锚定的正则匹配"可选序号前缀 + 可选括号 + 符号"，因此"该参数为★级"这类句中符号不会误报。
- 序号前缀正则覆盖纯文本编号：`（一）`、`(1)`、`1.`、`3.2、`、`一、`、`第1条` 等。

**Word 自动编号重建（关键难点，现有代码完全没有处理 w:numPr）**
- `_NumberingTracker` 类：从 docx 包中取 numbering 部件（优先 `document.part.part_related_by(RT.NUMBERING)`，异常时遍历 rels 找 numbering 关系，再无则视为无自动编号），解析 `w:num`/`w:abstractNum`/`w:lvl` 的 `numFmt`、`lvlText`、`start`、`startOverride`。
- `numbered_text(paragraph)`：查段落自身 `w:pPr/w:numPr`，无则沿样式链（style/base_style）查；按 (numId, ilvl) 维护计数器（遇上级层级重置下级，首次用 start/startOverride）；支持 decimal、chineseCounting 等（数字转中文）、字母、罗马数字、①②③、bullet 等格式；按 lvlText 模板（如 `%1、`、`%1.%2`）填级号后拼在段落文本前 → 输出"3.2、★支持XXX"，满足"保留序号"要求。

**条款提取 `extract_symbol_clauses(docx_path)`**
- 复用 `_iter_body_blocks` 按文档顺序遍历正文段落和表格，同时用 `_detect_heading_level` 记录章节（内部留存，不导出）。
- 纯文本段落：用重建编号后的完整文本，检查首行符号位置；命中则整段完整提取。
- 表格：先把表格内所有段落（含嵌套表格）过一遍编号跟踪器，得到"段落→带序号文本"映射；复用 `_extract_table` 的合并单元格处理逻辑按行取各单元格文本（嵌套表格已展开其中），逐单元格逐段检查符号位置；任一命中则**整行**（各非空单元格按列序换行拼接）完整提取——序号列、设备名称、参数全文都保留。同一行同符号去重，不同符号各生成一条。

**Excel 输出**
- `get_symbol_output_path`：`{stem}_指标参数.xlsx`，重名追加 `_1/_2`（同现有规则）。
- `export_symbol_clauses_to_excel`：openpyxl 单工作表「指标参数」，表头 3 列：**数量序号（1..N）、符号、详细内容（带序号）**；样式沿用现有导出（文本格式 @、自动换行、灰色细边框、CJK 宽度自适应列宽、内容列上限 80、行高不设由 Excel 自适应）。
- `batch_export_symbol_clauses`：批量 + 进度回调 + skipped/errors 结构完全镜像 `batch_export_word_tables`（如"未找到带符号条款"）。

## 2. `word_table_exporter.py` 最小重构（不改变现有行为）
- `_extract_table` / `_cell_text` / `_nested_table_lines` 增加可选参数 `paragraph_text_fn`（段落文本转换函数，默认 None 走原逻辑），供新模块传入"带重建序号的段落文本"。

## 3. `main.py` 融合进 WordTableExportWindow
- 输出目录区新增一行复选框：**「同时提取带符号的指标参数（★ ▲ △ # 等）」**（默认不勾选，保持现有行为）；在 `_setup_fonts` 补一个 `Surface.TCheckbutton` 样式使其与护眼底色融合。
- `start_export`：
  - 勾选时不强制要求先扫描/勾选表格（可只提取指标参数）；未勾选时保持现有校验。
  - 后台任务按顺序执行表格导出（有选中表格时）+ 指标参数提取（勾选时，作用于**全部**文件列表），进度条总量合并。
  - `_show_export_result` 扩展汇总：分「表格导出」「指标参数提取」两节显示条数与输出路径，跳过原因与失败信息合并展示。

## 4. 版本与文档（项目惯例）
- `app_info.py` 升至 v1.0.10 并写 changelog（build.py 会校验）。
- `docs/使用教程.md` 第七章补该选项说明与输出格式；README 功能列表加一行。

## 5. 测试与验证
- 新增 `tests/test_symbol_clause_extractor.py`：用 python-docx 生成文档；自动编号用例通过 zipfile 注入 numbering.xml（含 Content_Types 与 rels）再重开验证。覆盖：句首符号、序号后符号（纯文本编号）、自动编号重建（含中文编号"一、"）、表格整行提取（含纯符号单元格列）、句中符号不误报、多符号多行、xlsx 三列表头回读。
- 运行 `python -m pytest tests/ -q` 全量通过；再用 `docs/_sample_files/招标文件示例.docx` 做一次手工冒烟验证输出。