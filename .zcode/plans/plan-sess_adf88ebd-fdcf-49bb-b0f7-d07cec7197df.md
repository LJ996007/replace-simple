## 已确认的实施方案

1. 修改 `string_replacer.py` 的 `batch_replace`：在进入 Office 解析器前识别零字节的受支持文件；目标与源不同时原样复制，目标等于源时保留原文件，并将替换计数记为 `0`。非空损坏文件仍正常报错。
2. 在 `tests/test_string_replacer.py` 增加回归测试：
   - 有效 Office 文件没有匹配内容时，仍输出到独立目录且内容不变；
   - `.doc/.docx/.xls/.xlsx/.xlsm/.ppt/.pptx` 七种零字节文件均原样输出并记为 `0`。
3. 更新 `README.md` 和 `docs/使用教程.md`，明确已选文件即使没有匹配内容或为零字节也会输出。
4. 运行目标测试及完整测试套件，保留工作区其他未提交修改。