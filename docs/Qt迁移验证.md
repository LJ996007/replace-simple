# Qt 工作台交付验证

验证日期：2026-11-12。本轮按要求直接在当前工作区完成，未转交、未启动外部编码会话、未提交 Git commit。下列结果对应整套配方兼容修复后的最终源码。

## 实现与格式核对

旧版 `VisualFillDialog.save_recipe()` 将配方保存到用户设置的 `review_drag_profiles`：每项为无版本号的 `{operations, sources, targets}`。`operations` 有 `drag / edit / clear / layout / structure / auto` 六类；`layout` 包括 `height / width / fit / delete`。来源签名的负载为行数、列数、合并区域和首行单元格；目标签名沿用 `review_mapping.layout_signature()` 的列数、合并区域和猜测表头。旧配方不包含来源正文，也没有完整历史文件快照。

本轮增加独立的 `workbook_recipe.py`，不依赖 Tk 窗口进行重放。Qt 编辑器“排版与配方”菜单现在提供：

- “复用旧版已存配方”：读取原设置中的真实配方，选择后后台核验并应用。
- “载入并应用配方”：支持上述旧混合配方对象、Qt v2 整套操作 JSON 及原 Qt v1 单次拖放 JSON。
- “保存整套操作配方”：导出当前有效操作序列；旧操作及其参数保留，来源索引按结构匹配重定位；普通 Qt 编辑/清空以 `cells` 值记录，Qt 排版以 `format` 记录，映射配置和手动指定行随 `auto` 保存。操作日志纳入会话快照，撤销、重做保持一致。v2 是 Qt 可读格式，不承诺旧 Tk 能读取新增的 `cells / format` 操作。

整套混合配方从原模板创建独立会话，使用 `drop_values`、`WorkbookSession.edit/layout/structure`、`plan_fill`、`plan_visual_fill` 等现有业务层逐项验证。目标核对原模板签名；来源核对原算法结构签名，同结构不唯一时拒绝。原 Qt v1 保留“追加到当前编辑状态”的语义，仍要求结构匹配，文件名和表号只用于已匹配结构后的消歧。

重放前后校验原模板/当前版本的 SHA-256，以及现有扫描文件的大小和纳秒修改时间。结构编辑还传递 `expected_digest` 给业务层。重放成功才一次性提交，可整体撤销；失败保持当前状态及撤销/重做分支，并删除此次生成的临时版本。旧式空行删除在重放尾部物化，避免后续 Qt 编辑混淆显示行和原始行；不产生实际变更的旧排版操作也能继续结构重放。

## 最终回归

执行环境：Python 3.14.6 / PySide6 6.12.0 / Windows。

```text
PYTHONUTF8=1 python tests/hidden_pytest.py -q -s
296 passed, 44 subtests passed in 31.65s
```

完整回归包含原 Qt 的 18 项测试和新增混合配方的 12 项测试。最终专项复跑 `PYTHONUTF8=1 python tests/hidden_pytest.py tests/test_qt_frontend.py tests/test_workbook_recipe.py -q -s`：30 passed，8 subtests passed（4.36s）。`hidden_pytest.py` 使用独立 Windows 桌面，避免 Tk 测试窗口抢占焦点；Qt 测试使用 offscreen。

新增测试实际读取 DOCX、设置 JSON、配方 JSON，写入 XLSX，并用 openpyxl 重新打开核验：

- 由真实旧版窗口的业务方法生成 11 步配方并通过原保存方法写入设置；Qt 从设置菜单读取后重放，另存内容与旧版输出逐表比较，检查公式、合并区域、字体、换行、手动行高列宽和原模板字节不变。再导出 v2 JSON 并重新载入，验证完整旧操作序列保留及整体撤销/重做。
- 原生 Qt 文字、拖放、排版、结构编辑跨 JSON 往返；撤销/重做后的操作日志一致；以等号开头的正文安全保存。
- 结构操作已生成版本后发生公式冲突，事务回滚且清理临时版本，原有重做分支仍可使用。
- 重复来源、同文件名但新表头、缺失/不符的目标签名、无效选区/单元格/来源索引、未知操作、数组公式、受保护工作表及非有限行高拒绝。
- 来源和模板在重放前变化，以及来源在结构重放过程中变化时拒绝；删除空行的原始行坐标正确物化；无实际删行的布局记录后仍可插行。
- 原 Qt v1 配方追加现有编辑，随后导出 v2 能保留两部分历史。

此前已有的批量替换、失败重试、CSV 报告、扫描筛选、条款导出、模板库、Office 内容保留及工作簿结构/对象安全测试也包含在本次完整回归中。上一轮已有布局截图和机械检查记录，本轮没有将这些旧检查冒充新运行结果。

## 最终源码 PyInstaller 与启动冒烟

使用现有 `replace-simple.spec`，在项目内新建的 `.qt-final-verification-20261112` 下实际构建。最后业务修复后重新构建一次，最终退出码为 0：

```text
PYTHONUTF8=1 python -m PyInstaller --noconfirm --clean \
  --distpath .qt-final-verification-20261112/dist \
  --workpath .qt-final-verification-20261112/build replace-simple.spec
Build complete!
```

未运行会清理用户默认 `dist` 的 `build.py` 入口，也未改写用户已有 spec 或版本资源。核对冻结包包含 `workbook_recipe`、`workbook_session`、`qt_editor`、`qt_main` 及 `PySide6/plugins/platforms/qwindows.dll`。

通过 Win32 独立桌面运行下列最终入口，未设置 offscreen，使用原生 Windows Qt 平台插件；三者均正常显示/事件循环后退出，退出码为 0：

```text
python main.py --smoke-test
python qt_main.py --smoke-test
<隔离目录>/dist/replace-simple/replace-simple.exe --smoke-test
```

冻结包从新建的空工作目录运行，并清除 `PYTHONPATH/PYTHONHOME`。冒烟设置目录隔离，未产生应用 `error.log`。源码冒烟必须保留解释器原有用户 site-packages 搜索路径：本机 PySide6/PyInstaller 安装在 APPDATA 下，最初直接隔离 APPDATA 的辅助脚本曾报缺少 PySide6；修正辅助脚本路径后重测通过。该问题来自验证脚本的环境处理，不是通过修改应用或跳过依赖绕过。

打包前后核对用户已有 `dist` 的全部 2053 个文件，文件集合及每个文件 SHA-256 完全一致。验证完成后删除本轮创建的隔离目录，包含构建目录、冻结程序、日志、冒烟脚本、设置和空工作目录；未保留本轮发行包。

## 能力边界

- 旧配方自身没有正文、文件旧版本或已撤销分支，无法恢复未保存的信息；Qt 保存当前有效操作日志，不导出整个撤销树。重放须重新扫描来源；文件变化校验继续使用既有业务层策略。
- 自动行高保留操作与参数，但旧 Tk 使用字体实测，Qt 复用业务层字符宽度估算，结果可能有差异；本轮验证正文、换行和手动行高列宽，不声称自动行高逐像素相同。保存前需核对预览。
- 未知操作和现有业务层不支持的工作簿对象明确拒绝；保留公式、保护及文件不覆盖约束，没有降级为按名称应用。
- 本机原生启动/冻结包冒烟不等于干净 Windows 机器、混合 DPI、打印渲染或 WPS/Office COM 旧格式转换验证。
- Windows 默认 pytest 描述符捕获方式此前有间歇 Tcl 初始化问题，完整回归使用 `-s`；日志编码使用 `PYTHONUTF8=1`，避免 GBK 无法输出错误文本。
