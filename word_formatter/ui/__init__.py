"""Графический интерфейс приложения на PySide6."""

from __future__ import annotations

__all__ = ["MainWindow", "ProcessingTask", "run_application"]


def __getattr__(name: str):
    """Загружает части интерфейса по требованию.

    Это нужно, чтобы команды без графики (--version, --list-profiles)
    работали даже там, где PySide6 не установлен.
    """
    if name == "MainWindow":
        from word_formatter.ui.main_window import MainWindow

        return MainWindow
    if name == "ProcessingTask":
        from word_formatter.ui.worker import ProcessingTask

        return ProcessingTask
    if name == "run_application":
        from word_formatter.ui.application import run_application

        return run_application
    raise AttributeError(name)
