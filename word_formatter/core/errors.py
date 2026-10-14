"""Ошибки приложения.

Правило проекта: у каждой ошибки есть текст, понятный человеку без опыта
программирования. Технические подробности хранятся отдельно и попадают
в журнал, а не в лицо пользователю.
"""

from __future__ import annotations


class FormatterError(Exception):
    """Ошибка, которую можно показать пользователю как есть."""

    def __init__(self, user_message: str, *, details: str = "") -> None:
        super().__init__(user_message)
        self.user_message = user_message
        self.details = details

    def __str__(self) -> str:
        return self.user_message


class ProfileError(FormatterError):
    """Профиль оформления не найден, повреждён или не соответствует формату."""


class DocumentError(FormatterError):
    """Документ нельзя открыть или обработать."""


class UnsupportedFormatError(DocumentError):
    """Формат файла не поддерживается первой версией приложения."""


class CorruptedDocumentError(DocumentError):
    """Файл повреждён или не является документом Word."""


class PasswordProtectedError(DocumentError):
    """Документ защищён паролем."""


class DocumentUnreadableError(DocumentError):
    """Файл занят другой программой или недоступен для чтения."""


class SaveError(FormatterError):
    """Не удалось сохранить результат."""

