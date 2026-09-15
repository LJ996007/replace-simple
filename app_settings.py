"""当前用户设置的路径、读取和原子保存。"""

import json
import os

from file_io import atomic_output_path

APP_STATE_DIR_NAME = "replace-simple"
ERROR_LOG_NAME = "error.log"
SETTINGS_NAME = "settings.json"


def _app_data_dir(env_name):
    return os.path.join(os.environ.get(env_name) or os.path.expanduser("~"), APP_STATE_DIR_NAME)


def error_log_path():
    return os.path.join(_app_data_dir("LOCALAPPDATA"), ERROR_LOG_NAME)


def settings_path():
    return os.path.join(_app_data_dir("APPDATA"), SETTINGS_NAME)


def load_settings():
    try:
        with open(settings_path(), encoding="utf-8") as file:
            data = json.load(file)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(data):
    current = load_settings()
    current.update(data)
    with atomic_output_path(settings_path()) as temporary:
        with open(temporary, "w", encoding="utf-8") as file:
            json.dump(current, file, ensure_ascii=False, indent=2)
