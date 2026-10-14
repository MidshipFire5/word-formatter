"""Модель профиля оформления.

Профиль — это данные, а не код: добавление нового профиля не требует
изменения алгоритма форматирования. Модель читается из уже проверенного
словаря (проверкой занимается модуль validation).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

# Размеры листа в миллиметрах: ширины и высоты для книжной ориентации.
PAGE_FORMATS_MM: dict[str, tuple[float, float]] = {
    "A4": (210.0, 297.0),
    "A5": (148.0, 210.0),
    "A3": (297.0, 420.0),
    "Letter": (215.9, 279.4),
    "Legal": (215.9, 355.6),
}

ORIGIN_BUILTIN = "builtin"
ORIGIN_USER = "user"


@dataclass(frozen=True)
class Margins:
    """Поля страницы в миллиметрах."""

    left: float
    right: float
    top: float
    bottom: float

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "Margins":
        return cls(
            left=float(raw["left"]),
            right=float(raw["right"]),
            top=float(raw["top"]),
            bottom=float(raw["bottom"]),
        )


@dataclass(frozen=True)
class PageSettings:
    format: str
    orientation: str
    margins: Margins
    header_mm: float
    footer_mm: float

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "PageSettings":
        return cls(
            format=str(raw["format"]),
            orientation=str(raw["orientation"]),
            margins=Margins.from_dict(raw["marginsMm"]),
            header_mm=float(raw["headerMm"]),
            footer_mm=float(raw["footerMm"]),
        )

    @property
    def size_mm(self) -> tuple[float, float]:
        """Размер листа с учётом ориентации."""
        width, height = PAGE_FORMATS_MM[self.format]
        if self.orientation == "landscape":
            return height, width
        return width, height


@dataclass(frozen=True)
class TextStyle:
    """Оформление абзаца: применяется к тексту, заголовкам и подписям."""

    font: str | None
    size_pt: float
    bold: bool
    italic: bool
    color: str | None
    alignment: str
    first_line_indent_cm: float
    line_spacing: float
    space_before_pt: float
    space_after_pt: float

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "TextStyle":
        return cls(
            font=raw.get("font"),
            size_pt=float(raw["sizePt"]),
            bold=bool(raw.get("bold", False)),
            italic=bool(raw.get("italic", False)),
            color=raw.get("color"),
            alignment=str(raw["alignment"]),
            first_line_indent_cm=float(raw["firstLineIndentCm"]),
            line_spacing=float(raw["lineSpacing"]),
            space_before_pt=float(raw["spaceBeforePt"]),
            space_after_pt=float(raw["spaceAfterPt"]),
        )


@dataclass(frozen=True)
class BodySettings(TextStyle):
    """Оформление основного текста."""

    remove_italic: bool = True
    remove_underline: bool = True

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "BodySettings":
        base = TextStyle.from_dict(raw)
        return cls(
            font=base.font,
            size_pt=base.size_pt,
            bold=base.bold,
            italic=base.italic,
            color=base.color,
            alignment=base.alignment,
            first_line_indent_cm=base.first_line_indent_cm,
            line_spacing=base.line_spacing,
            space_before_pt=base.space_before_pt,
            space_after_pt=base.space_after_pt,
            remove_italic=bool(raw.get("removeItalic", True)),
            remove_underline=bool(raw.get("removeUnderline", True)),
        )


@dataclass(frozen=True)
class StructuralHeading(TextStyle):
    """Заголовки структурных частей: ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и подобные."""

    uppercase: bool = True
    titles: tuple[str, ...] = ()

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "StructuralHeading":
        base = TextStyle.from_dict(raw)
        return cls(
            font=base.font,
            size_pt=base.size_pt,
            bold=base.bold,
            italic=base.italic,
            color=base.color,
            alignment=base.alignment,
            first_line_indent_cm=base.first_line_indent_cm,
            line_spacing=base.line_spacing,
            space_before_pt=base.space_before_pt,
            space_after_pt=base.space_after_pt,
            uppercase=bool(raw.get("uppercase", True)),
            titles=tuple(str(title) for title in raw.get("titles", ())),
        )


@dataclass(frozen=True)
class HeadingsSettings:
    max_level: int
    no_hyphenation: bool
    structural: StructuralHeading
    levels: dict[int, TextStyle] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "HeadingsSettings":
        levels = {int(key): TextStyle.from_dict(value) for key, value in raw["levels"].items()}
        return cls(
            max_level=int(raw["maxLevel"]),
            no_hyphenation=bool(raw.get("noHyphenation", True)),
            structural=StructuralHeading.from_dict(raw["structural"]),
            levels=levels,
        )

    def level_style(self, level: int) -> TextStyle | None:
        """Оформление заголовка указанного уровня, если уровень описан профилем."""
        if level in self.levels:
            return self.levels[level]
        return self.levels.get(self.max_level)


@dataclass(frozen=True)
class CaptionStyle(TextStyle):
    """Оформление подписи к рисунку или названия таблицы."""

    no_hyphenation: bool = True
    position: str = "keep"

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "CaptionStyle":
        base = TextStyle.from_dict(raw)
        return cls(
            font=base.font,
            size_pt=base.size_pt,
            bold=base.bold,
            italic=base.italic,
            color=base.color,
            alignment=base.alignment,
            first_line_indent_cm=base.first_line_indent_cm,
            line_spacing=base.line_spacing,
            space_before_pt=base.space_before_pt,
            space_after_pt=base.space_after_pt,
            no_hyphenation=bool(raw.get("noHyphenation", True)),
            position=str(raw.get("position", "keep")),
        )


@dataclass(frozen=True)
class CaptionsSettings:
    figure: CaptionStyle
    table: CaptionStyle

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "CaptionsSettings":
        return cls(
            figure=CaptionStyle.from_dict(raw["figure"]),
            table=CaptionStyle.from_dict(raw["table"]),
        )


@dataclass(frozen=True)
class TableSettings:
    font: str | None
    font_size_pt: float
    line_spacing: float
    first_line_indent_cm: float
    header_row_align: str
    numeric_cell_align: str
    text_cell_align: str
    marker_cells: tuple[str, ...]
    preferred_width: str = "keep"
    indent_cm: float | None = None

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "TableSettings":
        indent = raw.get("indentCm")
        return cls(
            font=raw.get("font"),
            font_size_pt=float(raw["fontSizePt"]),
            line_spacing=float(raw["lineSpacing"]),
            first_line_indent_cm=float(raw["firstLineIndentCm"]),
            header_row_align=str(raw["headerRowAlign"]),
            numeric_cell_align=str(raw["numericCellAlign"]),
            text_cell_align=str(raw["textCellAlign"]),
            marker_cells=tuple(str(item) for item in raw.get("markerCells", ())),
            preferred_width=str(raw.get("preferredWidth", "keep")),
            indent_cm=None if indent is None else float(indent),
        )


@dataclass(frozen=True)
class FigureSettings:
    """Оформление абзаца с рисунком.

    Рисунок — это абзац с картинкой. Одного выравнивания мало: если абзац
    наследует абзацный отступ (например, из стиля «Обычный», куда правило
    записывает отступ первой строки для основного текста), картинка
    смещается вправо на величину отступа. Поэтому отступ у абзаца
    с рисунком снимается, а выравнивание задаётся явно.
    """

    alignment: str = "center"
    line_spacing: float = 1.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "FigureSettings":
        return cls(
            alignment=str(raw.get("alignment", "center")),
            line_spacing=float(raw.get("lineSpacing", 1.0)),
            space_before_pt=float(raw.get("spaceBeforePt", 0.0)),
            space_after_pt=float(raw.get("spaceAfterPt", 0.0)),
        )


@dataclass(frozen=True)
class FootnotesSettings:
    font: str | None
    font_size_pt: float
    line_spacing: float

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "FootnotesSettings":
        return cls(
            font=raw.get("font"),
            font_size_pt=float(raw["fontSizePt"]),
            line_spacing=float(raw["lineSpacing"]),
        )


@dataclass(frozen=True)
class ListsSettings:
    apply_paragraph_formatting: bool
    replace_markers: bool
    first_line_indent_cm: float

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "ListsSettings":
        return cls(
            apply_paragraph_formatting=bool(raw["applyParagraphFormatting"]),
            replace_markers=bool(raw["replaceMarkers"]),
            first_line_indent_cm=float(raw["firstLineIndentCm"]),
        )


@dataclass(frozen=True)
class PageNumbersSettings:
    enabled: bool
    position: str
    font: str | None
    font_size_pt: float
    skip_first_page: bool

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "PageNumbersSettings":
        return cls(
            enabled=bool(raw["enabled"]),
            position=str(raw["position"]),
            font=raw.get("font"),
            font_size_pt=float(raw["fontSizePt"]),
            skip_first_page=bool(raw["skipFirstPage"]),
        )


@dataclass(frozen=True)
class HyphenationSettings:
    enabled: bool
    exclude_headings: bool
    exclude_captions: bool

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "HyphenationSettings":
        return cls(
            enabled=bool(raw["enabled"]),
            exclude_headings=bool(raw["excludeHeadings"]),
            exclude_captions=bool(raw["excludeCaptions"]),
        )


@dataclass(frozen=True)
class BlankLinesSettings:
    insert: bool
    style: str

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "BlankLinesSettings":
        return cls(insert=bool(raw["insert"]), style=str(raw["style"]))


@dataclass(frozen=True)
class Profile:
    """Профиль оформления целиком."""

    id: str
    name: str
    description: str
    version: str
    demo: bool
    source: dict[str, Any]
    applies_to: tuple[str, ...]
    page: PageSettings
    body: BodySettings
    headings: HeadingsSettings
    captions: CaptionsSettings
    figures: FigureSettings
    tables: TableSettings
    footnotes: FootnotesSettings
    lists: ListsSettings
    page_numbers: PageNumbersSettings
    hyphenation: HyphenationSettings
    blank_lines: BlankLinesSettings
    limitations: tuple[str, ...]
    path: str | None = None
    origin: str = ORIGIN_BUILTIN

    @classmethod
    def from_dict(cls, raw: dict[str, Any], *, path: str | None = None, origin: str = ORIGIN_BUILTIN) -> "Profile":
        return cls(
            id=str(raw["id"]),
            name=str(raw["name"]),
            description=str(raw.get("description", "")),
            version=str(raw["version"]),
            demo=bool(raw.get("demo", False)),
            source=dict(raw.get("source", {})),
            applies_to=tuple(str(item) for item in raw.get("appliesTo", ())),
            page=PageSettings.from_dict(raw["page"]),
            body=BodySettings.from_dict(raw["body"]),
            headings=HeadingsSettings.from_dict(raw["headings"]),
            captions=CaptionsSettings.from_dict(raw["captions"]),
            figures=FigureSettings.from_dict(raw.get("figures", {})),
            tables=TableSettings.from_dict(raw["tables"]),
            footnotes=FootnotesSettings.from_dict(raw["footnotes"]),
            lists=ListsSettings.from_dict(raw["lists"]),
            page_numbers=PageNumbersSettings.from_dict(raw["pageNumbers"]),
            hyphenation=HyphenationSettings.from_dict(raw["hyphenation"]),
            blank_lines=BlankLinesSettings.from_dict(raw["blankLines"]),
            limitations=tuple(str(item) for item in raw.get("limitations", ())),
            path=path,
            origin=origin,
        )

    @property
    def source_title(self) -> str:
        """Короткое описание источника требований — для интерфейса и отчёта."""
        source = self.source
        if not source:
            return "источник не указан"
        parts = [str(source.get("title", ""))]
        if source.get("year"):
            parts.append(str(source["year"]))
        return ", ".join(part for part in parts if part)
