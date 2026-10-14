"""Сборка приложения: разбор аргументов командной строки и запуск интерфейса."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence

from word_formatter import APP_NAME, APP_VERSION
from word_formatter.core.errors import FormatterError
from word_formatter.core.profiles import ProfileCatalog

EXIT_OK = 0
EXIT_ERROR = 1
EXIT_GUI_MISSING = 3


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="main.py",
        description=f"{APP_NAME} — приведение документов Word к требованиям профиля оформления.",
    )
    parser.add_argument("--version", action="store_true", help="показать версию приложения и выйти")
    parser.add_argument(
        "--list-profiles",
        action="store_true",
        help="показать доступные профили оформления и выйти",
    )
    parser.add_argument(
        "--profiles-dir",
        metavar="ПАПКА",
        default=None,
        help="проверить профили в указанной папке вместо встроенной",
    )
    return parser


def _print_profiles(builtin_dir: str | None) -> int:
    catalog = ProfileCatalog(builtin_dir=builtin_dir) if builtin_dir else ProfileCatalog()
    builtin, user = catalog.directories
    print(f"{APP_NAME} {APP_VERSION}")
    print(f"Встроенные профили: {builtin}")
    print(f"Личные профили:     {user}")
    print()

    infos = catalog.list_infos()
    if not infos:
        print("Профили не найдены. Проверьте папку встроенных профилей.")
        return EXIT_ERROR

    for info in infos:
        origin = "личный" if info.origin == "user" else "встроенный"
        demo = ", демонстрационный" if info.demo else ""
        print(f"{info.id}  —  {info.name}")
        print(f"    версия {info.version or 'не указана'}, {origin}{demo}")
        if info.description:
            print(f"    {info.description}")
        if info.error:
            print("    ОШИБКА: " + info.error.replace("\n", "\n    "))
        print()
    return EXIT_OK


GUI_MISSING_MESSAGE = (
    "Не удалось запустить графический интерфейс: не установлена библиотека PySide6.\n"
    "Установите зависимости одной из команд:\n"
    "  Windows:  run_windows.bat\n"
    "  Linux:    bash run_linux.sh\n"
    "  вручную:  python -m pip install -r requirements.txt"
)


def _run_gui() -> int:
    try:
        from word_formatter.ui.application import run_application
    except ImportError:
        print(GUI_MISSING_MESSAGE, file=sys.stderr)
        return EXIT_GUI_MISSING
    try:
        return run_application()
    except ImportError:
        print(GUI_MISSING_MESSAGE, file=sys.stderr)
        return EXIT_GUI_MISSING


def main(argv: Sequence[str] | None = None) -> int:
    """Точка входа приложения. Возвращает код завершения процесса."""
    args = build_parser().parse_args(argv)
    try:
        if args.version:
            print(f"{APP_NAME} {APP_VERSION}")
            return EXIT_OK
        if args.list_profiles:
            return _print_profiles(args.profiles_dir)
        return _run_gui()
    except FormatterError as error:
        print("Ошибка: " + error.user_message, file=sys.stderr)
        return EXIT_ERROR
