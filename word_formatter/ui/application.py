"""Запуск графического интерфейса."""

from __future__ import annotations

import logging
import sys
from collections.abc import Sequence
from pathlib import Path

from word_formatter import APP_NAME, APP_VERSION
from word_formatter.core.paths import user_log_dir

LOG_FILE_NAME = "word-formatter.log"


def setup_logging(log_directory: Path | None = None) -> None:
    """Настраивает запись журнала работы приложения в файл.

    Ошибка записи журнала не должна мешать работе: если папка недоступна,
    приложение продолжает работать без файла журнала.
    """
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stderr)]
    try:
        directory = log_directory or user_log_dir()
        directory.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(directory / LOG_FILE_NAME, encoding="utf-8"))
    except OSError:
        pass

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=handlers,
        force=True,
    )


def run_application(argv: Sequence[str] | None = None) -> int:
    """Создаёт приложение Qt, показывает окно и запускает цикл обработки событий."""
    from PySide6.QtWidgets import QApplication

    setup_logging()

    application = QApplication(list(argv) if argv is not None else sys.argv)
    application.setApplicationName(APP_NAME)
    application.setApplicationDisplayName(APP_NAME)
    application.setApplicationVersion(APP_VERSION)
    application.setOrganizationName(APP_NAME)

    from word_formatter.ui.main_window import MainWindow
    from word_formatter.ui.theme import apply_theme

    apply_theme(application)

    window = MainWindow()
    window.show()
    return application.exec()
