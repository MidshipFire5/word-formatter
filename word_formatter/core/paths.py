"""Пути, которыми пользуется приложение.

Встроенные профили лежат рядом с программой, пользовательские — в личной
папке настроек. Разделение нужно, чтобы обновление проекта из GitHub
не затирало профили, добавленные пользователем.
"""

from __future__ import annotations

import os
from pathlib import Path

APP_DIR_NAME = "WordFormatter"
ENV_PROFILES_DIR = "WORD_FORMATTER_PROFILES_DIR"


def project_root() -> Path:
    """Корень проекта: папка, в которой лежат main.py и word_formatter/."""
    return Path(__file__).resolve().parents[2]


def builtin_profiles_dir() -> Path:
    """Папка встроенных профилей.

    Переменная окружения WORD_FORMATTER_PROFILES_DIR позволяет подменить
    папку профилей — это нужно тестам и переносимым сборкам.
    """
    override = os.environ.get(ENV_PROFILES_DIR)
    if override:
        return Path(override)
    return project_root() / "profiles"


def user_config_dir() -> Path:
    """Личная папка настроек приложения для текущего пользователя.

    Windows: %APPDATA%\\WordFormatter
    Linux и macOS: $XDG_CONFIG_HOME/WordFormatter или ~/.config/WordFormatter
    """
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA")
        if not base:
            base = str(Path.home() / "AppData" / "Roaming")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
    return Path(base) / APP_DIR_NAME


def user_profiles_dir() -> Path:
    """Папка личных профилей пользователя."""
    return user_config_dir() / "profiles"


def user_log_dir() -> Path:
    """Папка журналов работы приложения."""
    return user_config_dir() / "logs"
