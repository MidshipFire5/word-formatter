"""Правила оформления подписей к рисункам и названий таблиц."""

from __future__ import annotations

from word_formatter.core.docx import Role
from word_formatter.core.profiles import CaptionStyle
from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.typography import (
    apply_paragraph_style,
    apply_run_style,
    paragraph_runs,
    run_has_text,
    suppress_hyphenation,
)
from word_formatter.core.rules.text import alignment_title

CATEGORY_FIGURE_CAPTIONS = "figure-captions"
CATEGORY_TABLE_CAPTIONS = "table-captions"


def _format_caption(context: RuleContext, paragraph, style: CaptionStyle) -> None:
    apply_paragraph_style(paragraph, style)
    for run in paragraph_runs(paragraph):
        if not run_has_text(run):
            continue
        apply_run_style(
            run,
            font=style.font,
            size_pt=style.size_pt,
            bold=style.bold,
            italic=False,
            color=style.color,
        )
    if style.no_hyphenation:
        suppress_hyphenation(paragraph)


def apply_captions(context: RuleContext) -> None:
    """Оформляет подписи рисунков и названия таблиц."""
    profile = context.profile
    figures = 0
    tables = 0

    for info in context.inventory.paragraphs:
        if info.role is Role.FIGURE_CAPTION:
            style = profile.captions.figure
            _format_caption(context, context.paragraph(info.index), style)
            figures += 1
        elif info.role in (Role.TABLE_CAPTION, Role.TABLE_CONTINUATION):
            style = profile.captions.table
            _format_caption(context, context.paragraph(info.index), style)
            tables += 1

    if figures:
        style = profile.captions.figure
        context.changes.add(
            CATEGORY_FIGURE_CAPTIONS,
            f"оформление подписей рисунков: {style.font or 'шрифт не задан'} {style.size_pt:g} пт, "
            f"{alignment_title(style.alignment)}, без абзацного отступа, "
            "запрет переносов, одинарный интервал",
            detail=f"подписей: {figures}",
        )
    if tables:
        style = profile.captions.table
        context.changes.add(
            CATEGORY_TABLE_CAPTIONS,
            f"оформление названий таблиц: {style.font or 'шрифт не задан'} {style.size_pt:g} пт, "
            f"{alignment_title(style.alignment)}, без абзацного отступа, "
            "запрет переносов, одинарный интервал",
            detail=f"названий: {tables}",
        )
