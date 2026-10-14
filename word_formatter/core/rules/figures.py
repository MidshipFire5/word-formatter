"""Правила оформления рисунков.

Рисунок — это абзац с картинкой. Требование методички: рисунок стоит
по центру текстового поля, без абзацного отступа и без лишних интервалов.

Одного выравнивания недостаточно. Абзац с рисунком обычно не имеет
собственных свойств и наследует их от стиля «Обычный», которому правило
основного текста задаёт отступ первой строки (1,25 см). Из-за этого
картинка уезжает вправо даже при выравнивании по центру, поэтому отступ
у абзаца с рисунком снимается явно.
"""

from __future__ import annotations

from word_formatter.core.docx import Role
from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.text import alignment_title, title_page_end
from word_formatter.core.rules.typography import (
    ALIGNMENT,
    clear_paragraph_indent,
    set_spacing,
)

CATEGORY_FIGURES = "figures"


def apply_figures(context: RuleContext) -> None:
    """Ставит рисунки по центру и убирает у их абзацев абзацный отступ."""
    settings = context.profile.figures
    boundary = title_page_end(context)

    processed = 0
    title_page_paragraphs = 0

    for info in context.inventory.paragraphs:
        if info.role is not Role.FIGURE:
            continue
        if boundary is not None and info.index < boundary:
            title_page_paragraphs += 1
            continue

        paragraph = context.paragraph(info.index)
        paragraph.paragraph_format.alignment = ALIGNMENT.get(settings.alignment)
        clear_paragraph_indent(paragraph)
        set_spacing(
            paragraph,
            before_pt=settings.space_before_pt,
            after_pt=settings.space_after_pt,
            line_spacing=settings.line_spacing,
        )
        processed += 1

    if processed:
        context.changes.add(
            CATEGORY_FIGURES,
            f"рисунки: {alignment_title(settings.alignment)}, без абзацного отступа, "
            f"интервал {settings.line_spacing:g}",
            detail=f"рисунков: {processed}; отступ снят, иначе рисунок смещается вправо",
        )
    if title_page_paragraphs:
        context.changes.add(
            CATEGORY_FIGURES,
            "рисунки на титульном листе оставлены без изменений — проверьте вручную",
            detail=f"рисунков до первой структурной части: {title_page_paragraphs}",
        )
