"""User-level template library organised by one directory per category."""

from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass
from typing import Iterable, List, Sequence

from app_settings import settings_directory
from ui_common import SUPPORTED_EXTENSIONS


DEFAULT_TEMPLATE_LIBRARY_NAME = "templates"
_INVALID_NAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}


@dataclass(frozen=True)
class TemplateItem:
    category: str
    name: str
    path: str


def default_template_library_path() -> str:
    return os.path.join(settings_directory(), DEFAULT_TEMPLATE_LIBRARY_NAME)


def normalize_category_name(value: str) -> str:
    name = _INVALID_NAME_CHARS.sub(" ", str(value or "")).strip()
    name = re.sub(r"\s+", " ", name)
    if not name:
        raise ValueError("分类名称不能为空")
    if name in {".", ".."} or name.upper() in _WINDOWS_RESERVED_NAMES:
        raise ValueError(f"分类名称不能使用：{name}")
    if len(name) > 60:
        raise ValueError("分类名称不能超过 60 个字符")
    return name


def _is_supported_template(path: str) -> bool:
    return (
        not os.path.basename(path).startswith("~$")
        and os.path.splitext(path)[1].lower() in SUPPORTED_EXTENSIONS
    )


def ensure_library(root: str) -> str:
    os.makedirs(root, exist_ok=True)
    return root


def category_path(root: str, category: str) -> str:
    normalized = normalize_category_name(category)
    result = os.path.abspath(os.path.join(root, normalized))
    base = os.path.abspath(root)
    if os.path.commonpath((base, result)) != base:
        raise ValueError("分类路径超出模板库范围")
    return result


def list_categories(root: str) -> List[str]:
    if not os.path.isdir(root):
        return []
    categories = []
    for entry in os.scandir(root):
        if entry.is_dir(follow_symlinks=False):
            try:
                categories.append(normalize_category_name(entry.name))
            except ValueError:
                continue
    return sorted(categories, key=str.casefold)


def create_category(root: str, category: str) -> str:
    path = category_path(root, category)
    os.makedirs(path, exist_ok=True)
    return path


def rename_category(root: str, old_category: str, new_category: str) -> str:
    old_path = category_path(root, old_category)
    new_path = category_path(root, new_category)
    if not os.path.isdir(old_path):
        raise FileNotFoundError("原分类不存在")
    if os.path.exists(new_path):
        raise FileExistsError("新分类已存在")
    os.rename(old_path, new_path)
    return new_path


def delete_category(root: str, category: str) -> None:
    path = category_path(root, category)
    if not os.path.isdir(path):
        raise FileNotFoundError("分类不存在")
    if any(entry.is_file(follow_symlinks=False) for entry in os.scandir(path)):
        raise ValueError("分类内还有模板，请先删除模板")
    resolved_root = os.path.realpath(os.path.abspath(root))
    resolved_path = os.path.realpath(path)
    if os.path.commonpath((resolved_root, resolved_path)) != resolved_root:
        raise ValueError("分类路径超出模板库范围")
    os.rmdir(path)


def list_templates(root: str, categories: Sequence[str] | None = None) -> List[TemplateItem]:
    selected = list_categories(root) if categories is None else list(categories)
    templates: List[TemplateItem] = []
    for category in selected:
        directory = category_path(root, category)
        if not os.path.isdir(directory):
            continue
        for entry in os.scandir(directory):
            if entry.is_file(follow_symlinks=False) and _is_supported_template(entry.name):
                templates.append(TemplateItem(category, entry.name, entry.path))
    return sorted(templates, key=lambda item: (item.category.casefold(), item.name.casefold()))


def _unique_destination(directory: str, filename: str) -> str:
    stem, extension = os.path.splitext(filename)
    destination = os.path.join(directory, filename)
    counter = 2
    while os.path.exists(destination):
        destination = os.path.join(directory, f"{stem} ({counter}){extension}")
        counter += 1
    return destination


def add_templates(root: str, category: str, file_paths: Iterable[str]) -> List[str]:
    directory = create_category(root, category)
    added = []
    for source in file_paths:
        source = os.path.abspath(source)
        if not os.path.isfile(source) or not _is_supported_template(source):
            continue
        destination = _unique_destination(directory, os.path.basename(source))
        shutil.copy2(source, destination)
        added.append(destination)
    return added


def delete_template(root: str, template: TemplateItem) -> None:
    expected = category_path(root, template.category)
    resolved_template = os.path.realpath(template.path)
    resolved_category = os.path.realpath(expected)
    if os.path.commonpath((resolved_category, resolved_template)) != resolved_category:
        raise ValueError("模板路径超出分类范围")
    if not os.path.isfile(resolved_template):
        raise FileNotFoundError("模板不存在")
    os.remove(resolved_template)
