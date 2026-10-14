"""Сборка переносимой версии приложения (PyInstaller).

Собирается папка: исполняемый файл, библиотеки и встроенные профили рядом.
Одним файлом делать не будем — такой вариант распаковывается во временную
папку при каждом запуске, дольше стартует и чаще вызывает подозрение
антивирусов.

Запуск:
    python tools/build_exe.py

Результат: dist/WordFormatter/ (Windows) или dist/WordFormatter/ с двоичным
файлом без расширения (Linux).
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
NAME = "WordFormatter"
DIST_DIR = PROJECT_ROOT / "dist"
WORK_DIR = PROJECT_ROOT / ".tmp" / "pyinstaller"
MAIN_SCRIPT = PROJECT_ROOT / "main.py"
PROFILES_DIR = PROJECT_ROOT / "profiles"


def executable_name() -> str:
    return f"{NAME}.exe" if os.name == "nt" else NAME


def build_command() -> list[str]:
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--name",
        NAME,
        # Служебные файлы сборки держим в .tmp, чтобы не мусорить в проекте.
        "--specpath",
        str(WORK_DIR),
        "--workpath",
        str(WORK_DIR / "build"),
        "--distpath",
        str(DIST_DIR),
        # Встроенные профили обязаны попасть внутрь сборки: приложение ищет
        # их рядом с собой.
        "--add-data",
        f"{PROFILES_DIR}{os.pathsep}profiles",
        # pypdf подключается внутри функции, поэтому указываем явно.
        "--hidden-import",
        "pypdf",
    ]
    if os.name == "nt":
        # В Windows без этого при запуске открывалось бы чёрное окно консоли.
        command.append("--windowed")
    command.append(str(MAIN_SCRIPT))
    return command


def main() -> int:
    if not MAIN_SCRIPT.is_file():
        print(f"Не найден {MAIN_SCRIPT}")
        return 1
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("Не установлен PyInstaller. Установите: python -m pip install -r requirements-build.txt")
        return 1

    if DIST_DIR.exists():
        shutil.rmtree(DIST_DIR)

    print("Собираю приложение…")
    completed = subprocess.run(build_command(), cwd=PROJECT_ROOT)
    if completed.returncode != 0:
        print("Сборка не удалась.")
        return completed.returncode

    binary = DIST_DIR / NAME / executable_name()
    if not binary.is_file():
        print(f"Сборка завершилась, но файла {binary} нет.")
        return 1

    size_mb = sum(item.stat().st_size for item in (DIST_DIR / NAME).rglob("*") if item.is_file())
    print()
    print(f"Готово: {binary}")
    print(f"Размер папки: {size_mb / 1024 / 1024:.0f} МБ")
    print()
    print("Проверка:")
    print(f'  "{binary}" --version')
    print(f'  "{binary}" --list-profiles')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
