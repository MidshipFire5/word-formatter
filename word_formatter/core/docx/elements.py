"""Описание структуры документа: роли абзацев и инвентарь элементов.

Эти данные — единственный источник сведений для правил оформления.
Правила не разбирают XML документа повторно, а работают с готовой картиной.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class Role(str, Enum):
    """Роль абзаца в документе."""

    BODY = "body"
    BLANK = "blank"
    STRUCTURAL_HEADING = "structural_heading"
    HEADING = "heading"
    LIST_ITEM = "list_item"
    FIGURE_CAPTION = "figure_caption"
    TABLE_CAPTION = "table_caption"
    TABLE_CONTINUATION = "table_continuation"
    FORMULA = "formula"
    FIGURE = "figure"
    TABLE_CELL = "table_cell"
    AMBIGUOUS = "ambiguous"

    @property
    def title(self) -> str:
        """Название роли для отчёта."""
        return ROLE_TITLES[self]


ROLE_TITLES: dict[Role, str] = {
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
class ParagraphInfo:
    """Сведения об одном абзаце."""

    index: int
    text: str
    role: Role
    level: int | None = None
    style_name: str | None = None
    reason: str = ""
    has_drawing: bool = False
    has_formula: bool = False
    in_table: bool = False
    table_index: int | None = None

    @property
    def is_heading(self) -> bool:
        return self.role in (Role.HEADING, Role.STRUCTURAL_HEADING)

    @property
    def is_caption(self) -> bool:
        return self.role in (Role.FIGURE_CAPTION, Role.TABLE_CAPTION, Role.TABLE_CONTINUATION)

    @property
    def short_text(self) -> str:
        """Текст, укороченный для отчёта."""
        text = " ".join(self.text.split())
        return text if len(text) <= 80 else text[:77] + "…"


@dataclass(frozen=True)
class TableInfo:
    """Сведения об одной таблице."""

    index: int
    rows: int
    columns: int
    merged_cells: int
    paragraphs: tuple[ParagraphInfo, ...] = ()


@dataclass(frozen=True)
class BodyItem:
    """Элемент тела документа в порядке следования."""

    kind: str  # "paragraph" или "table"
    index: int


@dataclass(frozen=True)
class DocumentInventory:
    """Полная картина документа: что в нём есть и как оно устроено."""

    paragraphs: tuple[ParagraphInfo, ...] = ()
    tables: tuple[TableInfo, ...] = ()
    body: tuple[BodyItem, ...] = ()
    counts: dict[Role, int] = field(default_factory=dict)
    images: int = 0
    hyperlinks: int = 0
    footnotes: int = 0
    endnotes: int = 0
    sections: int = 0
    fields: int = 0
    toc_fields: int = 0
    ole_objects: int = 0
    content_controls: int = 0
    page_breaks: int = 0
    page_break_before: int = 0

    def role_count(self, role: Role) -> int:
        return self.counts.get(role, 0)

    @property
    def headings(self) -> tuple[ParagraphInfo, ...]:
        return tuple(item for item in self.paragraphs if item.is_heading)

    @property
    def ambiguous(self) -> tuple[ParagraphInfo, ...]:
        return tuple(item for item in self.paragraphs if item.role is Role.AMBIGUOUS)

    @property
    def has_toc(self) -> bool:
        return self.toc_fields > 0
