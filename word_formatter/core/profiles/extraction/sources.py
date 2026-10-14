"""Чтение текста методички.

Методичка может быть в Word, в PDF или простым текстом. Приложение умеет
прочитать все три случая и честно сказать, если файл прочитать нельзя:
например, это скан без текстового слоя или файл защищён паролем.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from word_formatter.core.docx.opening import open_document
from word_formatter.core.errors import MethodologyError, UnsupportedFormatError

TEXT_SUFFIXES = (".pdf", ".docx", ".txt", ".md", ".text")
SAMPLE_SUFFIXES = (".docx",)

# Ниже этого числа знаков текста считаем, что текстового слоя нет.
MINIMUM_TEXT_LENGTH = 200

KIND_PDF = "pdf"
KIND_DOCX = "docx"
KIND_TEXT = "text"

FILE_FILTER = (
    "Методичка (Word, PDF, текст) (*.docx *.pdf *.txt *.md);;"
    "Документ Word (*.docx);;"
    "Документ PDF (*.pdf);;"
    "Текстовый файл (*.txt *.md)"
)


@dataclass(frozen=True)
class SourceText:
    """Текст методички по страницам.

    Страницы нужны, чтобы в отчёте сослаться на место: «стр. 4».
    Для Word и текстовых файлов страница одна — весь файл.
    """

    pages: tuple[str, ...]
    kind: str
    path: Path

    @property
    def text(self) -> str:
        return "\n\n".join(self.pages)

    @property
    def page_count(self) -> int:
        return len(self.pages)

    def page_of(self, needle: str) -> int | None:
        """Номер страницы, где встретился фрагмент."""
        for number, page in enumerate(self.pages, start=1):
            if needle in page:
                return number
        return None


def read_source(path: str | Path) -> SourceText:
    """Читает текст методички. Формат определяется по расширению."""
    source_path = Path(path)
    if not source_path.is_file():
        raise MethodologyError(
            f"Файл «{source_path.name}» не найден. Возможно, он был перемещён или удалён."
        )

    suffix = source_path.suffix.lower()
    if suffix == ".pdf":
        return _read_pdf(source_path)
    if suffix == ".docx":
        return _read_docx(source_path)
    if suffix in (".txt", ".md", ".text"):
        return _read_text_file(source_path)

    if suffix == ".doc":
        raise UnsupportedFormatError(
            "Формат «.doc» не поддерживается. Откройте файл в Word и сохраните "
            "его как «.docx» или «.pdf» — тогда приложение сможет его прочитать."
        )
    raise UnsupportedFormatError(
        f"Формат «{suffix or 'без расширения'}» не поддерживается. "
        "Подходят Word (.docx), PDF (.pdf) и текстовые файлы (.txt, .md)."
    )


def _read_text_file(path: Path) -> SourceText:
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            text = path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise MethodologyError(
            f"Не удалось прочитать кодировку файла «{path.name}». "
            "Сохраните его в кодировке UTF-8 и попробуйте снова."
        )

    if not text.strip():
        raise MethodologyError(f"Файл «{path.name}» пустой — требований в нём нет.")
    return SourceText(pages=(text,), kind=KIND_TEXT, path=path)


def _read_docx(path: Path) -> SourceText:
    document = open_document(path)
    parts: list[str] = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            parts.append(" | ".join(cell for cell in cells))
    text = "\n".join(parts)
    if not text.strip():
        raise MethodologyError(
            f"В файле «{path.name}» нет текста. Похоже, это пустой документ "
            "или документ только с картинками."
        )
    return SourceText(pages=(text,), kind=KIND_DOCX, path=path)


PDF_LINE_BREAK = re.compile(r"-\s*\n\s*")
PDF_SOFT_BREAK = re.compile(r"\s*\n\s*")


def _join_pdf_lines(text: str) -> str:
    """Склеивает строки PDF в абзацы.

    PDF разбит на строки по ширине страницы, поэтому слова и фразы рвутся
    посередине: «над таблице» + «й». Для разбора требований строки нужно
    соединить. Перенос с дефисом склеивается без дефиса.
    """
    without_hyphens = PDF_LINE_BREAK.sub("", text)
    return PDF_SOFT_BREAK.sub(" ", without_hyphens)


def _read_pdf(path: Path) -> SourceText:
    try:
        from pypdf import PdfReader
    except ImportError as error:  # pragma: no cover - зависимость обязательна
        raise MethodologyError(
            "Для чтения PDF нужна библиотека pypdf, но она не установлена. "
            "Запустите подготовку окружения заново или сохраните методичку "
            "в формате Word или текста.",
            details=str(error),
        ) from error

    try:
        reader = PdfReader(str(path))
    except Exception as error:  # pypdf бросает разные типы исключений
        raise MethodologyError(
            f"Не удалось открыть PDF «{path.name}». Возможно, файл повреждён "
            "или защищён паролем. Попробуйте сохранить его заново.",
            details=str(error),
        ) from error

    if reader.is_encrypted:
        try:
            if reader.decrypt("") == 0:
                raise MethodologyError(
                    f"PDF «{path.name}» защищён паролем. Снимите защиту в программе "
                    "для чтения PDF и попробуйте снова."
                )
        except MethodologyError:
            raise
        except Exception as error:
            raise MethodologyError(
                f"PDF «{path.name}» защищён паролем — прочитать его нельзя.",
                details=str(error),
            ) from error

    pages: list[str] = []
    for page in reader.pages:
        try:
            pages.append(_join_pdf_lines(page.extract_text() or ""))
        except Exception:  # одна битая страница не должна ломать весь разбор
            pages.append("")

    text = "\n".join(pages)
    if len(text.strip()) < MINIMUM_TEXT_LENGTH:
        raise MethodologyError(
            f"В PDF «{path.name}» нет текстового слоя: похоже, это скан или фотографии "
            "страниц. Приложение читает только текст. Если методичка есть в Word "
            "или текстом — используйте её."
        )
    return SourceText(pages=tuple(pages), kind=KIND_PDF, path=path)
