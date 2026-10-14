"""Правила оформления заголовков.

Заголовки оформляются двумя способами сразу: приводится оформление самого
абзаца и обновляются определения стилей Word. Так документ выглядит правильно
и при этом сохраняет структуру — от неё зависит автоматическое содержание.
"""

from __future__ import annotations

import re

from docx.enum.style import WD_STYLE_TYPE

from word_formatter.core.docx import Role
from word_formatter.core.profiles import TextStyle
from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.typography import (
    apply_paragraph_style,
    apply_run_style,
    apply_style_text_style,
    paragraph_runs,
    run_has_text,
    suppress_hyphenation,
)
from word_formatter.core.rules.text import alignment_title

CATEGORY_STRUCTURAL = "structural-headings"
CATEGORY_HEADINGS = "headings"
CATEGORY_HEADING_STYLES = "heading-styles"

HEADING_STYLE_PATTERN = re.compile(r"^(?:heading|заголовок)\s*(\d)$", re.IGNORECASE)


def find_heading_style(document, level: int):
    """Находит стиль заголовка нужного уровня в документе.

    Имена стилей зависят от языка Word: «Heading 1» или «Заголовок 1».
    """
    for style in document.styles:
        if style.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        for candidate in (style.style_id or "", style.name or ""):
            match = HEADING_STYLE_PATTERN.match(candidate.strip())
            if match and int(match.group(1)) == level:
                return style
    return None


def apply_heading_styles(context: RuleContext) -> None:
    """Приводит определения стилей заголовков к профилю."""
    updated = 0
    missing: list[int] = []

    for level in sorted(context.profile.headings.levels):
        style = find_heading_style(context.document, level)
        if style is None:
            missing.append(level)
            continue
        apply_style_text_style(style, context.profile.headings.levels[level])
        updated += 1

    detail = f"обновлено стилей заголовков: {updated}"
    if missing:
        detail += f"; стили не найдены для уровней: {', '.join(str(level) for level in missing)}"
    context.changes.add(
        CATEGORY_HEADING_STYLES,
        "оформление стилей заголовков: полужирный, одинарный интервал, "
        "размер и отступы по профилю",
        detail=detail,
    )


def _format_heading(
    context: RuleContext,
    paragraph,
    info,
    style: TextStyle,
    *,
    uppercase: bool,
) -> None:
    """Оформляет абзац заголовка и его текст."""
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
            all_caps=True if uppercase else None,
            color=style.color,
        )

    if context.profile.headings.no_hyphenation:
        suppress_hyphenation(paragraph)


def _assign_heading_style(context: RuleContext, paragraph, level: int) -> bool:
    """Назначает абзацу стиль заголовка, если он ещё не назначен."""
    style = find_heading_style(context.document, level)
    if style is None:
        return False
    current = paragraph.style
    if current is not None and current.style_id == style.style_id:
        return False
    paragraph.style = style
    return True


def apply_headings(context: RuleContext) -> None:
    """Оформляет заголовки структурных частей, разделов и подразделов."""
    profile = context.profile
    structural = profile.headings.structural

    structural_count = 0
    assigned = 0
    level_counts: dict[int, int] = {}

    for info in context.inventory.paragraphs:
        if info.role is Role.STRUCTURAL_HEADING:
            paragraph = context.paragraph(info.index)
            _format_heading(context, paragraph, info, structural, uppercase=structural.uppercase)
            if _assign_heading_style(context, paragraph, 1):
                assigned += 1
            structural_count += 1
            context.changes.add(
                CATEGORY_STRUCTURAL,
                "оформление заголовка структурной части: "
                f"{structural.font or 'шрифт не задан'} {structural.size_pt:g} пт, полужирный, "
                f"{alignment_title(structural.alignment)}, без абзацного отступа, "
                + ("прописными буквами, " if structural.uppercase else "")
                + "одинарный интервал",
                paragraph_index=info.index,
                text=info.short_text,
            )
        elif info.role is Role.HEADING:
            level = info.level or 1
            style = profile.headings.level_style(level)
            if style is None:
                continue
            paragraph = context.paragraph(info.index)
            _format_heading(context, paragraph, info, style, uppercase=False)
            if _assign_heading_style(context, paragraph, level):
                assigned += 1
            level_counts[level] = level_counts.get(level, 0) + 1

    for level in sorted(level_counts):
        style = profile.headings.level_style(level)
        context.changes.add(
            CATEGORY_HEADINGS,
            f"оформление заголовков уровня {level}: {style.font or 'шрифт не задан'} "
            f"{style.size_pt:g} пт, полужирный, {alignment_title(style.alignment)}, "
            f"абзацный отступ {style.first_line_indent_cm:g} см, одинарный интервал",
            detail=f"заголовков уровня {level}: {level_counts[level]}",
        )

    if assigned:
        context.changes.add(
            CATEGORY_HEADINGS,
            "заголовкам назначен стиль Word — от этого зависит структура документа "
            "и автоматическое содержание",
            detail=f"абзацев со стилем заголовка: {assigned}",
        )

    if not structural_count and not level_counts:
        context.changes.add(
            CATEGORY_HEADINGS,
            "заголовки в документе не найдены",
            detail="проверьте, что заголовки оформлены стилями Word или начинаются с номера",
        )
