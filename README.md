# replace-simple

只保留批量文本替换功能的简化版桌面工具。

## 功能

- 支持 Word `.docx`
- 支持 Excel `.xlsx` / `.xlsm`
- 支持 PowerPoint `.pptx`
- 可在界面中直接录入规则：左侧原文本，右侧替换后文本
- 可导入 Excel 规则表并继续编辑，支持 `.xlsx` / `.xlsm` / `.xls`
- Excel 规则表读取两列：第一列原文，第二列替换文
- 可从招标 Word `.docx` 中读取常见项目信息，自动生成 `{项目名称}`、`{项目编号}`、`{采购人}`、`{预算金额}` 等占位符替换规则
- 文件内容和文件名都会按规则同步替换
- 当文件名替换后会覆盖源文件时，自动添加 `_已替换` 后缀保护源文件
- 可独立提取 Word `.docx` 正文表格到 Excel：每个 Word 生成一个 `_表格.xlsx`，每张表格对应一个工作表，并尽量保留合并单元格
- 可通过主界面右上角「版本」按钮查看当前版本号、发布日期和历次更新说明

## 运行

```bash
pip install -r requirements.txt
python main.py
```

## 打包

```bash
python build.py
```

生成目录位于 `dist/replace-simple/`，可直接运行 `dist/replace-simple/replace-simple.exe`。
脚本会同时生成带版本号、便于分发的 `dist/replace-simple-v<版本号>-portable.zip`。

## Word 表格导出

点击主界面右上角「提取 Word 表格」打开独立工具窗口：

1. 添加一个或多个 `.docx` 文件，或选择包含 Word 文件的文件夹。
2. 点击「扫描表格」，根据所在章节、提示和内容预览勾选需要导出的表格。
3. 可选择输出目录；未选择时输出到各 Word 文件所在目录。
4. 点击「开始导出」。

输出文件默认命名为 `原文件名_表格.xlsx`；如果同名文件已存在，会自动追加 `_1`、`_2` 等后缀，避免覆盖。

## 从招标文件导入项目信息

点击「导入招标文件」选择招标 `.docx`，程序会识别常见字段并填入规则表。
模板文件中可使用 `{项目名称}`、`{项目编号}`、`{采购人}`、`{采购代理机构}`、`{预算金额}`、`{最高限价}`、`{采购方式}`、`{合同履行期限}`、`{提交投标文件截止时间}`、`{开标时间}`、`{开标地点}` 等占位符。

## 分发说明

- 当前打包产物为 Windows 64 位程序，基于 Python 3.11，建议在 Windows 10/11 64 位系统运行。
- 未签名 exe 在新电脑首次运行时可能触发 SmartScreen 提示；这是未签名程序的常见现象，代码签名证书可作为后续根治方案。
- 如果通过网络下载 zip，Windows 可能给文件加 Mark-of-the-Web。遇到拦截时，可先右键 zip → 属性 → 勾选“解除锁定”，再解压运行。
- 建议解压到普通本地目录，例如桌面或文档目录，避免路径过深、网络盘、受保护系统目录。
- 程序遇到未捕获异常时，会弹窗提示并写入日志：`%LOCALAPPDATA%\replace-simple\error.log`。

## 开发提示

- 测试：`python -m pytest tests`
- 打包前建议清理旧的 `build/`、`dist/`，`build.py --clean` 已会重新生成 PyInstaller 分析结果。
- 每次发布前在 `app_info.py` 中更新 `APP_VERSION`、`APP_RELEASE_DATE`，并把本次更新说明添加到 `APP_CHANGELOG` 首位。
- `python build.py` 会校验当前版本是否有更新说明，并同步更新窗口标题、软件内版本信息、Windows EXE 文件属性和压缩包文件名。
