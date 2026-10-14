"""Правила оформления элементов перечислений.

Маркеры и нумерация списков не заменяются: приложение меняет только
оформление абзацев. Причина — автонумерация связана с текстом, и её замена
может поехать вместе со ссылками на пункты.
"""

from __future__ import annotations

from dataclasses import replace

from word_formatter.core.docx import Role
from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.text import apply_paragraph_and_runs, alignment_title

CATEGORY_LISTS = "lists"
CATEGORY_MARKERS = "list-markers"


def apply_lists(context: RuleContext) -> None:
    """Оформляет абзацы элементов перечислений, не трогая маркеры."""
    settings = context.profile.lists
    body = context.profile.body
    style = replace(body, first_line_indent_cm=settings.first_line_indent_cm)

    processed = 0
    for info in context.inventory.paragraphs:
        if info.role is not Role.LIST_ITEM:
            continue
        apply_paragraph_and_runs(
            context,
            context.paragraph(info.index),
            info,
            style,
            remove_italic=body.remove_italic,
            remove_underline=body.remove_underline,
        )
        processed += 1

    if processed:
        context.changes.add(
            CATEGORY_LISTS,
            f"оформление элементов перечислений: {style.font or 'шрифт не задан'} "
            f"{style.size_pt:g} пт, {alignment_title(style.alignment)}, "
            f"абзацный отступ {settings.first_line_indent_cm:g} см, "
            f"межстрочный интервал {style.line_spacing:g}",
            detail=f"элементов: {processed}",
        )

    if not settings.replace_markers:
        context.changes.add(
            CATEGORY_MARKERS,
            "маркеры и нумерация списков не изменялись — приложение оформляет "
            "только абзацы перечислений",
            detail="если методичка требует тире вместо маркеров, проверьте списки вручную",
        )
