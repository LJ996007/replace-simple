"""
replace-simple 打包脚本。

打包策略：--onedir（目录形态）而非 --onefile。
原因：onefile 每次启动都要把整个运行时解压到临时目录，开销可达数百毫秒甚至
秒级，与「保证启动速度」直接冲突；onedir 无需解压、启动快，且整个目录复制
即可运行、不写注册表、不依赖系统 Python —— 本身就是绿色版。

打包完成后，额外把 dist/<APP_NAME> 压缩成 zip，方便作为单文件下载分发。
"""

import os
import shutil
import subprocess
import sys
import zipfile

from app_info import APP_CHANGELOG, APP_NAME, APP_RELEASE_DATE, APP_VERSION

MAIN_SCRIPT = "main.py"
VERSION_INFO_FILE = "version_info.txt"

PYINSTALLER_ARGS = [
    f"--name={APP_NAME}",
    "--onedir",          # 目录形态：无需每次解压，启动更快（绿色版）
    "--windowed",        # 无控制台窗口
    "--clean",
    "--noconfirm",
    "--noupx",           # DLL 压缩会拖慢加载并增加杀软误报概率，分发版明确关闭
    "--icon=icon.ico",           # exe 文件图标
    f"--version-file={VERSION_INFO_FILE}",  # Windows 文件属性中的版本信息
    "--add-data=icon-256.png;.",   # 运行时窗口图标用大尺寸 PNG（Windows 用 ; 分隔）
    "--hidden-import=openpyxl",
    "--hidden-import=xlrd",
    "--hidden-import=docx",
    "--hidden-import=pptx",
    "--hidden-import=tksheet",
    "--collect-data=tksheet",   # tksheet 带有内置主题/图标等数据文件，打包时需一并收集
]

# 排除程序完全用不到的重型依赖。它们会被传递性拉入，且 pywin32 的
# run-time hook（pyi_rth_pywintypes / pyi_rth_pythoncom）会在 exe 启动时
# 强制加载 COM dll，是启动开销与体积的大头。文本替换核心路径
# （openpyxl 读写单元格、docx/pptx 读写文本）不依赖以下任何模块。
EXCLUDE_MODULES = [
    "numpy", "pandas", "scipy", "matplotlib",
    "win32com", "pythoncom", "pywintypes", "pywin32",
    "win32evtlog", "win32evtlogutil", "win32api",
    "bs4", "charset_normalizer", "soupsieve",
    "pyreadline3",
    "lxml.isoschematron", "lxml.html", "lxml.objectify", "lxml.sax",
    "pythonnet", "clr_loader", "clr",
    "jinja2", "yaml", "PyYAML",
    # 注意：不要排除 PIL / Pillow —— python-pptx 在 pptx/parts/image.py
    # 硬编码 from PIL import Image，排除后替换 .pptx 会直接崩溃。
]


def _version_tuple() -> tuple[int, int, int, int]:
    """把 1.2.3 形式的版本号转换为 Windows 版本信息需要的四段整数。"""
    try:
        parts = [int(part) for part in APP_VERSION.split(".")]
    except ValueError as exc:
        raise ValueError(f"版本号必须是数字点分格式，例如 1.2.3：{APP_VERSION}") from exc
    if not 1 <= len(parts) <= 4:
        raise ValueError(f"版本号应包含 1 至 4 段数字：{APP_VERSION}")
    if any(part < 0 or part > 65535 for part in parts):
        raise ValueError(f"版本号每一段必须在 0 至 65535 之间：{APP_VERSION}")
    return tuple((parts + [0, 0, 0, 0])[:4])


def _validate_release_info() -> None:
    if not APP_CHANGELOG:
        raise ValueError("更新说明不能为空，请先在 app_info.py 中填写 APP_CHANGELOG")
    latest = APP_CHANGELOG[0]
    if latest.get("version") != APP_VERSION:
        raise ValueError("APP_CHANGELOG 第一条记录的版本号必须与 APP_VERSION 一致")
    if latest.get("date") != APP_RELEASE_DATE:
        raise ValueError("APP_CHANGELOG 第一条记录的日期必须与 APP_RELEASE_DATE 一致")
    if not latest.get("items"):
        raise ValueError("当前版本必须至少填写一条更新说明")


def _write_version_info_file(path: str) -> None:
    version = _version_tuple()
    content = f'''# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={version},
    prodvers={version},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '080404B0',
        [StringStruct('CompanyName', '{APP_NAME}'),
         StringStruct('FileDescription', '{APP_NAME}'),
         StringStruct('FileVersion', '{APP_VERSION}'),
         StringStruct('InternalName', '{APP_NAME}'),
         StringStruct('OriginalFilename', '{APP_NAME}.exe'),
         StringStruct('ProductName', '{APP_NAME}'),
         StringStruct('ProductVersion', '{APP_VERSION}')]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [2052, 1200])])
  ]
)
'''
    with open(path, "w", encoding="utf-8") as file:
        file.write(content)


def _clean_old_artifacts(dist_dir: str) -> None:
    """删除上一次 onefile 模式遗留的 dist/<APP_NAME>.exe，避免与新目录共存。"""
    old_exe = os.path.join("dist", f"{APP_NAME}.exe")
    if os.path.isfile(old_exe):
        os.remove(old_exe)
    if os.path.isdir(dist_dir):
        shutil.rmtree(dist_dir)


def _make_portable_zip(dist_dir: str, zip_path: str) -> None:
    """把整个 onedir 产物压缩成 zip，便于作为单文件分发。"""
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for root, _dirs, files in os.walk(dist_dir):
            for name in files:
                full = os.path.join(root, name)
                arc = os.path.relpath(full, "dist")
                zf.write(full, arc)


def build() -> bool:
    if not os.path.exists(MAIN_SCRIPT):
        print(f"错误：找不到主脚本 {MAIN_SCRIPT}")
        return False

    try:
        _validate_release_info()
        _write_version_info_file(VERSION_INFO_FILE)
    except (OSError, ValueError) as exc:
        print(f"版本信息错误：{exc}")
        return False

    print(f"当前版本：v{APP_VERSION}")
    command = ["pyinstaller", *PYINSTALLER_ARGS]
    command += [f"--exclude-module={m}" for m in EXCLUDE_MODULES]
    command.append(MAIN_SCRIPT)
    print("执行命令：")
    print(" ".join(command))

    try:
        subprocess.run(command, check=True)
    except FileNotFoundError:
        print("错误：找不到 pyinstaller，请先安装：pip install pyinstaller")
        return False
    except subprocess.CalledProcessError as exc:
        print(f"打包失败：{exc}")
        return False

    dist_dir = os.path.join("dist", APP_NAME)
    zip_path = os.path.join("dist", f"{APP_NAME}-v{APP_VERSION}-portable.zip")
    print(f"打包成功：{dist_dir}/{APP_NAME}.exe")

    try:
        _make_portable_zip(dist_dir, zip_path)
        print(f"绿色压缩包：{zip_path}")
    except Exception as exc:  # 压缩失败不影响主产物
        print(f"警告：生成压缩包失败：{exc}")

    return True


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    _clean_old_artifacts(os.path.join("dist", APP_NAME))
    sys.exit(0 if build() else 1)
