"""Application version and release notes."""

APP_NAME = "replace-simple"
APP_VERSION = "1.0.1"
APP_RELEASE_DATE = "2026-07-10"

APP_CHANGELOG = [
    {
        "version": APP_VERSION,
        "date": APP_RELEASE_DATE,
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
