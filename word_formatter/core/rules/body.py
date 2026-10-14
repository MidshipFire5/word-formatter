"""Элементы тела документа в порядке следования.

Общая часть для правил, которым документ нужен как последовательность
абзацев и таблиц: правила свободных строк и правила положения подписей.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from docx.oxml.ns import qn

from word_formatter.core.docx import Role
from word_formatter.core.rules.context import RuleContext

# Роли, вокруг которых методичка требует свободную строку.
SPECIAL_ROLES = frozenset(
    {
        Role.STRUCTURAL_HEADING,
        Role.HEADING,
        Role.FIGURE,
        Role.FIGURE_CAPTION,
        Role.TABLE_CAPTION,
        Role.TABLE_CONTINUATION,
        Role.FORMULA,
    }
)

ROLE_TITLES = {
    Role.BODY: "основной текст",
    Role.BLANK: "пустая строка",
    Role.STRUCTURAL_HEADING: "заголовок структурной части",
    Role.HEADING: "заголовок раздела или подраздела",
    Role.LIST_ITEM: "элемент перечисления",
    Role.FIGURE_CAPTION: "подпись рисунка",
    Role.TABLE_CAPTION: "название таблицы",
    Role.TABLE_CONTINUATION: "продолжение таблицы",
    Role.FORMULA: "формула",
    Role.FIGURE: "рисунок",
    Role.TABLE_CELL: "текст в таблице",
    Role.AMBIGUOUS: "спорный абзац",
}


@dataclass(frozen=True)
class BodyElement:
    """Один элемент тела документа: абзац или таблица."""

    kind: str  # "paragraph" или "table"
    index: int
    role: Role | None
    element: Any
    text: str

    @property
    def is_paragraph(self) -> bool:
        return self.kind == "paragraph"

    @property
    def is_table(self) -> bool:
        return self.kind == "table"

    @property
    def is_blank(self) -> bool:
        return self.role is Role.BLANK

    @property
    def is_special(self) -> bool:
        return self.kind == "table" or (self.role is not None and self.role in SPECIAL_ROLES)

    @property
    def title(self) -> str:
        if self.kind == "table":
            return "таблица"
        return ROLE_TITLES.get(self.role, "абзац") if self.role else "абзац"


def collect_body(context: RuleContext) -> list[BodyElement]:
    """Собирает элементы тела документа в порядке следования."""
    items: list[BodyElement] = []
    for body_item in context.inventory.body:
        if body_item.kind == "paragraph":
            info = context.inventory.paragraphs[body_item.index]
            element = context.document.paragraphs[body_item.index]._p
            items.append(
                BodyElement(
                    kind="paragraph",
                    index=body_item.index,
                    role=info.role,
                    element=element,
                    text=info.short_text,
                )
            )
        else:
            table = context.document.tables[body_item.index]
            items.append(
                BodyElement(
                    kind="table",
                    index=body_item.index,
                    role=None,
                    element=table._tbl,
                    text=f"таблица {body_item.index + 1}",
                )
            )
    return items


def starts_new_page(paragraph_element) -> bool:
    """Начинается ли абзац с новой страницы."""
    pPr = paragraph_element.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:pageBreakBefore")) is not None


def ends_with_page_break(paragraph_element) -> bool:
    """Заканчивается ли абзац разрывом страницы."""
    for node in paragraph_element.iter(qn("w:br")):
        if node.get(qn("w:type")) == "page":
            return True
    return False
