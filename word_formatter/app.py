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
    parser.add_argument(
        "--self-check",
        action="store_true",
        help="проверить, что приложение обрабатывает документ: создать документ, "
        "оформить его встроенным профилем и показать итог",
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


def _run_self_check() -> int:
    """Проверяет, что приложение умеет обрабатывать документ.

    Документ создаётся тут же, во временной папке, и оформляется первым
    встроенным профилем: так проверяются библиотеки, встроенные профили и
    весь проход обработки — без графического интерфейса и без файлов
    пользователя. Полезно после сборки переносимой версии.
    """
    import tempfile
    from pathlib import Path

    from docx import Document

    from word_formatter.core.processing import process_document

    catalog = ProfileCatalog()
    infos = [info for info in catalog.list_infos() if info.is_usable]
    if not infos:
        print("Проверка не пройдена: встроенные профили не найдены.")
        return EXIT_ERROR

    selected = infos[0]
    profile = catalog.load(selected.id)

    with tempfile.TemporaryDirectory(prefix="word-formatter-check-") as directory:
        folder = Path(directory)
        source = folder / "проверка.docx"
        target = folder / "проверка-оформлено.docx"

        document = Document()
        document.add_paragraph("ВВЕДЕНИЕ")
        document.add_paragraph(
            "Проверочный абзац: приложение должно оформить его по профилю, "
            "не меняя текст."
        )
        document.save(source)

        result = process_document(source, target, profile)

    print(f"{APP_NAME} {APP_VERSION}")
    print(f"Профиль: {profile.name}")
    print(f"Абзацев в проверочном документе: {len(result.inventory.paragraphs)}")
    print(f"Изменений в отчёте: {len(result.changes.items)}")

    if not result.verification.passed:
        print("Проверка не пройдена: содержимое документа изменилось.")
        for problem in result.verification.problems:
            print("  • " + problem)
        return EXIT_ERROR

    print("Проверка пройдена: документ обработан, содержимое сохранено.")
    return EXIT_OK


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
        if args.self_check:
            return _run_self_check()
        return _run_gui()
    except FormatterError as error:
        print("Ошибка: " + error.user_message, file=sys.stderr)
        return EXIT_ERROR
