"""Работа с документами Word: открытие, разбор структуры, сохранение."""

from __future__ import annotations

from word_formatter.core.docx.elements import (
    BodyItem,
    DocumentInventory,
    ParagraphInfo,
    Role,
    TableInfo,
)
from word_formatter.core.docx.inspection import classify_paragraph, describe_table, inspect_document
from word_formatter.core.docx.opening import open_document

__all__ = [
    "BodyItem",
    "DocumentInventory",
    "ParagraphInfo",
    "Role",
    "TableInfo",
    "classify_paragraph",
    "describe_table",
    "inspect_document",
    "open_document",
]
