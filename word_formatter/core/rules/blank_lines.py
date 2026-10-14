"""Правила вставки свободных строк.

Методичка требует оставлять по одной свободной строке между заголовками и
текстом, выше и ниже таблиц, рисунков и формул. Поскольку интервалы между
абзацами заданы нулевыми, свободная строка — это именно пустой абзац.

Вставка делается осторожно: пустая строка не появляется второй раз, не
появляется в начале и конце документа, и не вставляется там, где часть
начинается с новой страницы. Каждая вставка попадает в отчёт.
"""

from __future__ import annotations

from word_formatter.core.docx import Role
from word_formatter.core.rules.body import (
    BodyElement,
    collect_body,
    ends_with_page_break,
    starts_new_page,
)
from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.typography import set_first_line_indent, set_spacing

CATEGORY_BLANK_LINES = "blank-lines"


def needs_blank_line(previous: BodyElement, following: BodyElement) -> bool:
    """Нужна ли свободная строка между двумя соседними элементами."""
    if previous.is_blank or following.is_blank:
        return False
    if not (previous.is_special or following.is_special):
        return False
    return not belongs_together(previous, following)


def belongs_together(previous: BodyElement, following: BodyElement) -> bool:
    """Части одного целого: рисунок с подписью, название таблицы с таблицей."""
    if previous.role is Role.FIGURE and following.role is Role.FIGURE_CAPTION:
        return True
    if previous.role in (Role.TABLE_CAPTION, Role.TABLE_CONTINUATION) and following.is_table:
        return True
    # Части разорванной таблицы не разделяются пустой строкой.
    if previous.is_table and following.is_table:
        return True
    if previous.is_table and following.role is Role.TABLE_CONTINUATION:
        return True
    return False


def _create_blank_paragraph(context: RuleContext):
    """Создаёт пустой абзац, оформленный как основной текст."""
    paragraph = context.document.add_paragraph()
    body = context.profile.body
    set_first_line_indent(paragraph, 0)
    if context.profile.blank_lines.style == "single":
        set_spacing(paragraph, before_pt=0, after_pt=0, line_spacing=1.0)
    else:
        set_spacing(
            paragraph,
            before_pt=body.space_before_pt,
            after_pt=body.space_after_pt,
            line_spacing=body.line_spacing,
        )
    return paragraph


def apply_blank_lines(context: RuleContext) -> None:
    """Вставляет свободные строки там, где требует профиль."""
    if not context.profile.blank_lines.insert:
        context.changes.add(
            CATEGORY_BLANK_LINES,
            "вставка свободных строк отключена профилем — проверьте интервалы вручную",
        )
        return

    items = collect_body(context)
    insertions: list[tuple[BodyElement, BodyElement]] = []

    for position in range(len(items) - 1):
        previous, following = items[position], items[position + 1]
        if not needs_blank_line(previous, following):
            continue
        if following.is_paragraph and starts_new_page(following.element):
            continue
        if previous.is_paragraph and ends_with_page_break(previous.element):
            continue
        insertions.append((previous, following))

    for previous, following in insertions:
        paragraph = _create_blank_paragraph(context)
        following.element.addprevious(paragraph._p)
        context.changes.add(
            CATEGORY_BLANK_LINES,
            "вставлена свободная строка",
            paragraph_index=following.index if following.is_paragraph else None,
            text=following.text,
            detail=f"между «{previous.title}» и «{following.title}»",
        )

    context.changes.add(
        CATEGORY_BLANK_LINES,
        f"свободные строки: вставлено {len(insertions)} — по одной между заголовками "
        "и текстом, выше и ниже таблиц, рисунков и формул",
        detail="повторный запуск на обработанном документе вторых пустых строк не добавит",
    )
