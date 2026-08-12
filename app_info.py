"""Application version and release notes."""

APP_NAME = "replace-simple"
APP_VERSION = "1.0.7"
APP_RELEASE_DATE = "2026-08-12"

APP_CHANGELOG = [
    {
        "version": APP_VERSION,
        "date": APP_RELEASE_DATE,
        "items": [
            "修复预设管理窗口内容被裁切、无法正常操作的问题，并统一窗口视觉风格。",
            "预设多选改为清晰的蓝色对号，支持鼠标滚轮浏览预设和各类列表。",
            "统一所有滚动条为扁平淡蓝色样式，缩短空列表滑块并加入方向箭头。",
        ],
    },
    {
        "version": "1.0.6",
        "date": "2026-08-12",
        "items": [
            "移除替换规则的 Excel 导出按钮和功能。",
            "新增可多选的原文本预设下拉面板，内置 12 个常用项目字段。",
            "支持新增、重命名、删除并持久保存预设，内置预设同样可以编辑。",
        ],
    },
    {
        "version": "1.0.5",
        "date": "2026-08-07",
        "items": [
            "修复项目编号等替换内容包含 /、\\、* 等 Windows 禁用字符时，输出文件路径错误并导致部分文件处理失败的问题。",
            "仅对输出文件名转换禁用字符，文档正文中的项目编号等内容保持原样。",
        ],
    },
    {
        "version": "1.0.4",
        "date": "2026-07-15",
        "items": [
            "界面改为柔和护眼浅色：去掉大面积纯白，略降对比与冷蓝感，长时间查看更舒适。",
            "扫描结果表头增加列间竖线与底部分隔横线，并与表头灰底对齐。",
            "「提取 Word 表格」窗口与结果弹窗相对父窗口居中；输出目录区压缩为单行。",
        ],
    },
    {
        "version": "1.0.3",
        "date": "2026-07-14",
        "items": [
            "点击「版本」后弹窗显示在主窗口正中央，不再出现在屏幕左上角。",
        ],
    },
    {
        "version": "1.0.2",
        "date": "2026-07-14",
        "items": [
            "替换后若文件名与源文件相同，直接覆盖原文件，不再生成「_已替换」副本。",
            "目标路径被其他已有文件占用时，仍自动追加 _1、_2 等后缀，避免误覆盖无关文件。",
        ],
    },
    {
        "version": "1.0.1",
        "date": "2026-07-10",
        "items": [
            "文件夹扫描、规则导入和规则导出改为后台执行，减少界面卡顿。",
            "批处理开始时固定文件、规则和输出目录，避免处理中修改界面导致结果不一致。",
            "输出文件自动避开已有文件和同名任务，防止覆盖；Office 依赖按文件类型延迟加载。",
        ],
    },
    {
        "version": "1.0.0",
        "date": "2026-07-10",
        "items": [
            "新增版本信息按钮，可在软件内查看当前版本号和更新说明。",
            "打包时自动读取统一版本号，并生成带版本号的绿色压缩包文件名。",
            "为 Windows 可执行文件写入版本属性，便于在文件属性中核对版本。",
        ],
    },
]


def format_version_title():
    return f"{APP_NAME} v{APP_VERSION}"


def format_changelog():
    lines = [
        f"软件名称：{APP_NAME}",
        f"当前版本：v{APP_VERSION}",
        f"发布日期：{APP_RELEASE_DATE}",
        "",
        "更新说明：",
    ]

    for entry in APP_CHANGELOG:
        lines.append("")
        lines.append(f"v{entry['version']}（{entry['date']}）")
        for item in entry.get("items", []):
            lines.append(f"• {item}")

    return "\n".join(lines)
