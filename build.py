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


APP_NAME = "replace-simple"
MAIN_SCRIPT = "main.py"

PYINSTALLER_ARGS = [
    f"--name={APP_NAME}",
    "--onedir",          # 目录形态：无需每次解压，启动更快（绿色版）
    "--windowed",        # 无控制台窗口
    "--clean",
    "--noconfirm",
    "--noupx",           # DLL 压缩会拖慢加载并增加杀软误报概率，分发版明确关闭
    "--icon=icon.ico",           # exe 文件图标
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
    zip_path = os.path.join("dist", f"{APP_NAME}-portable.zip")
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
