"""Открытие документов Word с понятными сообщениями об ошибках.

Правило проекта: приложение не пытается «как-нибудь» открыть файл.
Если формат не тот или файл повреждён, пользователь получает объяснение
и следующий шаг, а не трассировку стека.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.opc.exceptions import PackageNotFoundError

from word_formatter.core.errors import (
    CorruptedDocumentError,
    DocumentUnreadableError,
    PasswordProtectedError,
    UnsupportedFormatError,
)

SUPPORTED_SUFFIX = ".docx"
LEGACY_SUFFIX = ".doc"
MACRO_SUFFIXES = {".docm", ".dotm"}

# Сигнатура контейнера OLE: так выглядят старые .doc и зашифрованные .docx.
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_SIGNATURE = b"PK\x03\x04"

LEGACY_MESSAGE = (
    "Файл сохранён в старом формате Word (.doc), который первая версия приложения не поддерживает. "
    "Откройте документ в Word, выберите «Файл → Сохранить как» и сохраните его в формате "
    "«Документ Word (*.docx)», затем попробуйте снова."
)

MACRO_MESSAGE = (
    "Файл содержит макросы (.docm) — такие документы приложение не обрабатывает, "
    "чтобы не затронуть программный код внутри файла. Сохраните документ как обычный «.docx» "
    "(«Файл → Сохранить как → Документ Word (*.docx)») и попробуйте снова."
)


def _check_suffix(path: Path) -> None:
    suffix = path.suffix.lower()
    if suffix == SUPPORTED_SUFFIX:
        return
    if suffix == LEGACY_SUFFIX:
        raise UnsupportedFormatError(LEGACY_MESSAGE, details=f"файл: {path}")
    if suffix in MACRO_SUFFIXES:
        raise UnsupportedFormatError(MACRO_MESSAGE, details=f"файл: {path}")
    shown = path.suffix or "без расширения"
    raise UnsupportedFormatError(
        f"Приложение работает только с документами Word в формате «.docx», "
        f"а выбран файл «{shown}». Если это документ Word, сохраните его как «.docx» "
        "и попробуйте снова.",
        details=f"файл: {path}",
    )


def _read_header(path: Path) -> bytes:
    try:
        with open(path, "rb") as handle:
            return handle.read(8)
    except PermissionError as error:
        raise DocumentUnreadableError(
            f"Не удалось прочитать файл «{path.name}»: он занят другой программой или закрыт для чтения. "
            "Закройте документ в Word и попробуйте снова.",
            details=str(error),
        ) from error
    except OSError as error:
        raise DocumentUnreadableError(
            f"Не удалось прочитать файл «{path.name}»: {error.strerror or error}. "
            "Проверьте, что файл доступен и не находится на отключённом диске.",
            details=str(error),
        ) from error


def ensure_supported_format(path: str | Path) -> None:
    """Проверяет расширение файла, не открывая его.

    Нужна интерфейсу: пользователь должен узнать о неподдерживаемом формате
    сразу после выбора файла, а не после нажатия кнопки.
    """
    _check_suffix(Path(path))


def open_document(path: str | Path) -> Document:
    """Открывает документ Word для чтения.

    Исходный файл не изменяется: python-docx читает его целиком в память.
    """
    document_path = Path(path)
    _check_suffix(document_path)

    if not document_path.exists():
        raise CorruptedDocumentError(
            f"Файл «{document_path.name}» не найден. Возможно, он был перемещён или удалён.",
            details=f"путь: {document_path}",
        )
    if not document_path.is_file():
        raise CorruptedDocumentError(
            f"«{document_path.name}» — это папка, а не документ Word.",
            details=f"путь: {document_path}",
        )
    if document_path.stat().st_size == 0:
        raise CorruptedDocumentError(
            f"Файл «{document_path.name}» пустой (0 байт), поэтому открыть его нельзя.",
            details=f"путь: {document_path}",
        )

    header = _read_header(document_path)

    if header.startswith(OLE_SIGNATURE):
        raise PasswordProtectedError(
            f"Файл «{document_path.name}» защищён паролем или сохранён в старом формате. "
            "Снимите защиту в Word («Рецензирование → Ограничить редактирование») или сохраните "
            "документ как «.docx» без пароля и попробуйте снова.",
            details=f"путь: {document_path}, сигнатура OLE",
        )

    if not header.startswith(ZIP_SIGNATURE):
        raise CorruptedDocumentError(
            f"Файл «{document_path.name}» не является документом Word: его содержимое не похоже "
            "на формат DOCX. Возможно, файл повреждён или переименован из другого формата.",
            details=f"путь: {document_path}, первые байты: {header!r}",
        )

    try:
        return Document(str(document_path))
    except PackageNotFoundError as error:
        raise CorruptedDocumentError(
            f"Документ «{document_path.name}» повреждён: внутри файла нет основной части документа. "
            "Попробуйте открыть его в Word и сохранить заново.",
            details=str(error),
        ) from error
    except Exception as error:  # библиотека сообщает о разных проблемах разными типами
        raise CorruptedDocumentError(
            f"Не удалось открыть документ «{document_path.name}». Возможно, файл повреждён "
            "или сохранён другой программой в несовместимом виде. Попробуйте открыть его в Word "
            "и сохранить заново в формате «.docx».",
            details=f"{type(error).__name__}: {error}",
        ) from error
