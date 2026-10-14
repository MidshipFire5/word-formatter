"""Правила оформления таблиц.

Выравнивание внутри ячейки зависит от содержимого: шапка — по центру,
числовые ячейки — по центру, текстовые — по левому краю. Ячейки-прочерки
(«–», «х», пустые) не трогаются: это знаки отсутствия данных, и их
выравнивание задаёт автор.
"""

from __future__ import annotations

import re

from word_formatter.core.profiles import TableSettings
from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.text import alignment_title
from word_formatter.core.rules.typography import (
    ALIGNMENT,
    apply_run_style,
    paragraph_runs,
    run_has_text,
    set_first_line_indent,
    set_spacing,
)

CATEGORY_TABLES = "tables"
CATEGORY_TABLES_MANUAL = "tables-manual"
CATEGORY_ITALIC = "italic-removed"
CATEGORY_UNDERLINE = "underline-removed"

NUMBER = re.compile(r"^[+-]?(?:\d+(?:[.,]\d+)?|[.,]\d+)$")
TRAILING_UNITS = re.compile(r"(?i)\s*(?:руб\.?|р\.|₽|%|шт\.?|ед\.?|тыс\.?|млн\.?|млрд\.?)$")
LEADING_WORDS = re.compile(r"(?i)^(?:от|до|с|по)\s+")
DIGIT_SPACE = re.compile(r"(?<=\d)[\s\u00a0\u202f](?=\d)")
HAS_DIGIT = re.compile(r"\d")
HAS_LETTER = re.compile(r"[A-Za-zА-Яа-яЁё]")


def normalize_cell_text(text: str) -> str:
    """Убирает лишние пробелы и неразрывные пробелы внутри числа."""
    return " ".join(text.split())


def is_numeric_cell(text: str) -> bool:
    """Ячейка содержит только число (возможно, с единицей измерения)."""
    cleaned = normalize_cell_text(text)
    if not cleaned:
        return False
    cleaned = TRAILING_UNITS.sub("", cleaned)
    cleaned = LEADING_WORDS.sub("", cleaned).strip()
    cleaned = DIGIT_SPACE.sub("", cleaned)
    if not cleaned:
        return False
    return NUMBER.match(cleaned) is not None


def is_mixed_cell(text: str) -> bool:
    """В ячейке есть и текст, и числа: выравнивание по типу данных невозможно."""
    return bool(HAS_DIGIT.search(text)) and bool(HAS_LETTER.search(text))


def _short(text: str, limit: int = 80) -> str:
    collapsed = " ".join(text.split())
    return collapsed if len(collapsed) <= limit else collapsed[: limit - 1] + "…"


def _format_cell_paragraph(context: RuleContext, paragraph, settings: TableSettings) -> None:
    italic_allowed_to_remove = context.profile.body.remove_italic
    underline_allowed_to_remove = context.profile.body.remove_underline

    for run in paragraph_runs(paragraph):
        if not run_has_text(run):
            continue
        was_italic = run.italic is True
        was_underline = run.underline is not None and bool(run.underline)

        apply_run_style(
            run,
            font=settings.font,
            size_pt=settings.font_size_pt,
            italic=False if italic_allowed_to_remove else None,
            underline=False if underline_allowed_to_remove else None,
        )

        if italic_allowed_to_remove and was_italic:
            context.changes.add(
                CATEGORY_ITALIC,
                "снят курсив",
                text=_short(paragraph.text),
            )
        if underline_allowed_to_remove and was_underline:
            context.changes.add(
                CATEGORY_UNDERLINE,
                "снято подчёркивание",
                text=_short(paragraph.text),
            )

    set_first_line_indent(paragraph, settings.first_line_indent_cm)
    set_spacing(paragraph, before_pt=0, after_pt=0, line_spacing=settings.line_spacing)


def apply_tables(context: RuleContext) -> None:
    """Оформляет текст внутри таблиц и выравнивает ячейки по типу данных."""
    settings = context.profile.tables
    markers = {normalize_cell_text(marker).casefold() for marker in settings.marker_cells}

    filled = 0
    header = 0
    numeric = 0
    textual = 0
    marker_cells = 0
    mixed = 0
    merged = 0

    for index, table_info in enumerate(context.inventory.tables):
        table = context.table(index)
        merged += table_info.merged_cells
        # Объединённые по горизонтали ячейки возвращаются в row.cells несколько раз.
        # Сравнивать их по id() нельзя: lxml создаёт объекты-обёртки на лету,
        # поэтому ссылки на уже обработанные ячейки сохраняются и сравниваются
        # по тождеству объектов.
        processed_cells: list = []

        for row_index, row in enumerate(table.rows):
            for cell in row.cells:
                tc = cell._tc
                if any(tc is existing for existing in processed_cells):
                    continue
                processed_cells.append(tc)

                for paragraph in cell.paragraphs:
                    _format_cell_paragraph(context, paragraph, settings)
                    text = normalize_cell_text(paragraph.text)
                    if not text:
                        continue
                    filled += 1

                    if row_index == 0:
                        paragraph.paragraph_format.alignment = ALIGNMENT[settings.header_row_align]
                        header += 1
                    elif text.casefold() in markers:
                        marker_cells += 1
                    elif is_numeric_cell(text):
                        paragraph.paragraph_format.alignment = ALIGNMENT[settings.numeric_cell_align]
                        numeric += 1
                    else:
                        paragraph.paragraph_format.alignment = ALIGNMENT[settings.text_cell_align]
                        textual += 1
                        if is_mixed_cell(text):
                            mixed += 1

    if not context.inventory.tables:
        return

    context.changes.add(
        CATEGORY_TABLES,
        f"оформление таблиц: {settings.font or 'шрифт не задан'} {settings.font_size_pt:g} пт, "
        f"одинарный интервал, без абзацного отступа; шапка {alignment_title(settings.header_row_align)}, "
        f"числа {alignment_title(settings.numeric_cell_align)}, "
        f"текст {alignment_title(settings.text_cell_align)}",
        detail=f"таблиц: {len(context.inventory.tables)}; заполненных ячеек: {filled}; "
        f"шапка: {header}; числовых: {numeric}; текстовых: {textual}; "
        f"ячеек-прочерков без изменений: {marker_cells}",
    )

    if mixed:
        context.changes.add(
            CATEGORY_TABLES_MANUAL,
            "в ячейках есть и текст, и числа — выравнивание выполнено по левому краю, "
            "проверьте такие ячейки вручную",
            detail=f"ячеек: {mixed}",
        )
    if merged:
        context.changes.add(
            CATEGORY_TABLES_MANUAL,
            "в таблицах есть объединённые ячейки — их границы и объединение не изменялись, "
            "проверьте вручную",
            detail=f"объединённых ячеек: {merged}",
        )
