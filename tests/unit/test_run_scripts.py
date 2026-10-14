"""Скрипты запуска и файл зависимостей.

Проверки простые, но ловят ошибки, из-за которых приложение не запустится
у пользователя: русские буквы в пакетном файле Windows, переносы строк в
Linux-скрипте, метка порядка байтов, незафиксированные версии зависимостей.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from tools import bootstrap

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
WINDOWS_SCRIPT = PROJECT_ROOT / "run_windows.bat"
LINUX_SCRIPT = PROJECT_ROOT / "run_linux.sh"
REQUIREMENTS = PROJECT_ROOT / "requirements.txt"
REQUIREMENTS_DEV = PROJECT_ROOT / "requirements-dev.txt"
GITIGNORE = PROJECT_ROOT / ".gitignore"
MESSAGES_DIR = PROJECT_ROOT / "tools" / "messages"
MESSAGE_FILES = ("no_python.txt", "incomplete_copy.txt", "failed.txt")


@pytest.mark.parametrize("path", [WINDOWS_SCRIPT, LINUX_SCRIPT])
def test_script_exists_and_is_not_empty(path: Path) -> None:
    assert path.is_file()
    assert len(path.read_text(encoding="utf-8")) > 200


# --- Windows -----------------------------------------------------------------


def test_windows_script_is_ascii_only() -> None:
    """cmd.exe разбирает пакетные файлы по текущей кодовой странице.

    Русские буквы в самом файле приводят к тому, что части слов выполняются
    как команды. Поэтому в пакетном файле допустимы только латинские символы,
    а русский текст печатается из отдельных файлов как данные.
    """
    WINDOWS_SCRIPT.read_bytes().decode("ascii")


def test_windows_script_starts_without_bom() -> None:
    data = WINDOWS_SCRIPT.read_bytes()

    assert not data.startswith(b"\xef\xbb\xbf")
    assert data.startswith(b"@echo off")


def test_windows_script_switches_console_to_utf8() -> None:
    text = WINDOWS_SCRIPT.read_text(encoding="ascii")

    assert "chcp 65001" in text, "без этого русские сообщения превратятся в мусор"
    assert text.count("chcp 65001") >= 3, "переключение нужно перед каждым выводом файла"


def test_windows_script_prints_russian_from_files() -> None:
    text = WINDOWS_SCRIPT.read_text(encoding="ascii")

    for name in MESSAGE_FILES:
        assert f"tools\\messages\\{name}" in text


def test_windows_script_detects_store_stub() -> None:
    """Заглушка Microsoft Store не печатает ничего, её нужно распознать."""
    text = WINDOWS_SCRIPT.read_text(encoding="ascii")

    assert "tools\\bootstrap.py" in text
    assert 'py -3 -c "print(\'ok\')"' in text
    assert 'python -c "print(\'ok\')"' in text


def test_windows_script_reports_exit_code() -> None:
    text = WINDOWS_SCRIPT.read_text(encoding="ascii")

    assert "set \"EXIT_CODE=%ERRORLEVEL%\"" in text
    assert "exit /b %EXIT_CODE%" in text
    assert "pause" in text, "без паузы окно закроется и пользователь не прочитает ошибку"


def test_windows_script_does_not_install_globally() -> None:
    """Установкой занимается bootstrap.py и только внутри .venv."""
    text = WINDOWS_SCRIPT.read_text(encoding="ascii").lower()

    assert "pip install" not in text
    assert "--user" not in text


# --- Linux -------------------------------------------------------------------


def test_linux_script_has_unix_line_endings() -> None:
    """Переносы строк CRLF делают скрипт нерабочим в Linux."""
    data = LINUX_SCRIPT.read_bytes()

    assert b"\r\n" not in data
    assert data.startswith(b"#!/usr/bin/env bash\n")


def test_linux_script_hands_over_to_bootstrap() -> None:
    text = LINUX_SCRIPT.read_text(encoding="utf-8")

    assert 'exec "$PYTHON" tools/bootstrap.py "$@"' in text
    for candidate in ("python3", "python"):
        assert candidate in text


def test_linux_script_explains_fallback_launch() -> None:
    text = LINUX_SCRIPT.read_text(encoding="utf-8")

    assert "bash run_linux.sh" in text
    assert "apt install" in text
    assert "dnf install" in text


def test_linux_script_checks_bootstrap_presence() -> None:
    text = LINUX_SCRIPT.read_text(encoding="utf-8")

    assert "tools/bootstrap.py" in text
    assert "распакован не полностью" in text


# --- общее -------------------------------------------------------------------


@pytest.mark.parametrize("path", [WINDOWS_SCRIPT, LINUX_SCRIPT])
def test_scripts_do_not_request_admin_rights(path: Path) -> None:
    text = path.read_text(encoding="utf-8").lower()

    assert "runas" not in text
    assert "start-process" not in text


@pytest.mark.parametrize("name", MESSAGE_FILES)
def test_message_file_is_russian_utf8(name: str) -> None:
    path = MESSAGES_DIR / name
    assert path.is_file(), f"скрипт запуска печатает этот файл: {name}"

    text = path.read_text(encoding="utf-8")
    assert re.search(r"[А-Яа-яЁё]", text), "сообщение должно быть на русском"
    assert len(text.splitlines()) >= 4


def test_no_python_message_is_shared_with_bootstrap() -> None:
    """Инструкция по установке Python должна совпадать в двух местах.

    Скрипт запуска печатает файл, а bootstrap берёт из него же текст ошибки:
    так пользователь не получит двух разных инструкций.
    """
    file_text = (MESSAGES_DIR / "no_python.txt").read_text(encoding="utf-8")
    first_line = file_text.splitlines()[0].strip()

    message = bootstrap.unsupported_version_message((3, 9, 18), "nt")

    assert first_line in message
    assert "python.org/downloads/windows" in message


def test_missing_message_file_falls_back() -> None:
    fallback = bootstrap.message_file("нет-такого-файла.txt", "запасной текст")

    assert fallback == "запасной текст"


def test_bootstrap_reports_its_own_messages_in_russian() -> None:
    for message in (
        bootstrap.VENV_MODULE_HINT,
        bootstrap.QT_HINT,
        bootstrap.WINDOWS_QT_HINT,
        bootstrap.LINUX_INSTRUCTIONS,
    ):
        assert re.search(r"[А-Яа-яЁё]", message)


@pytest.mark.parametrize("path", [REQUIREMENTS, REQUIREMENTS_DEV])
def test_requirements_are_pinned(path: Path) -> None:
    """Незафиксированная версия однажды сломает установку у пользователя."""
    pattern = re.compile(r"^[A-Za-z0-9._-]+==[0-9][^\s]*$")

    lines = [
        line.strip()
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.strip().startswith("#") and not line.strip().startswith("-r ")
    ]
    assert lines, "файл зависимостей не должен быть пустым"
    for line in lines:
        assert pattern.match(line), f"версия не зафиксирована: {line}"


def test_venv_is_excluded_from_repository() -> None:
    text = GITIGNORE.read_text(encoding="utf-8")

    assert ".venv/" in text
    assert "__pycache__/" in text
    assert ".pytest_cache/" in text
