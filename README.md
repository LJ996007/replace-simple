# replace-simple

只保留批量文本替换功能的简化版桌面工具。

## 功能

- 支持 Word `.docx`
- 支持 Excel `.xlsx` / `.xlsm`
- 支持 PowerPoint `.pptx`
- 可在界面中直接录入规则：左侧原文本，右侧替换后文本
- 可导入 Excel 规则表并继续编辑，支持 `.xlsx` / `.xlsm` / `.xls`
- Excel 规则表读取两列：第一列原文，第二列替换文
- 文件内容和文件名都会按规则同步替换
- 当文件名替换后会覆盖源文件时，自动添加 `_已替换` 后缀保护源文件

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
脚本会同时生成便于分发的 `dist/replace-simple-portable.zip`。

## 分发说明

- 当前打包产物为 Windows 64 位程序，基于 Python 3.11，建议在 Windows 10/11 64 位系统运行。
- 未签名 exe 在新电脑首次运行时可能触发 SmartScreen 提示；这是未签名程序的常见现象，代码签名证书可作为后续根治方案。
- 如果通过网络下载 zip，Windows 可能给文件加 Mark-of-the-Web。遇到拦截时，可先右键 zip → 属性 → 勾选“解除锁定”，再解压运行。
- 建议解压到普通本地目录，例如桌面或文档目录，避免路径过深、网络盘、受保护系统目录。
- 程序遇到未捕获异常时，会弹窗提示并写入日志：`%LOCALAPPDATA%\replace-simple\error.log`。

## 开发提示

- 测试：`python -m pytest tests`
- 打包前建议清理旧的 `build/`、`dist/`，`build.py --clean` 已会重新生成 PyInstaller 分析结果。
