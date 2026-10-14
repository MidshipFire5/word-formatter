"""Правила оформления подстрочных сносок.

python-docx не умеет работать со сносками как с объектами, поэтому
оформление задаётся определением стиля сносок — это надёжно и не требует
трогать саму часть файла со сносками.
"""

from __future__ import annotations

from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt

from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.xml_helpers import set_rfonts

CATEGORY_FOOTNOTES = "footnotes"

FOOTNOTE_STYLE_IDS = {"footnotetext"}
FOOTNOTE_STYLE_NAMES = {"footnote text"}
ENDNOTE_STYLE_IDS = {"endnotetext"}
ENDNOTE_STYLE_NAMES = {"endnote text"}


def find_note_style(document, style_ids: set[str], style_names: set[str]):
    """Находит стиль сносок или концевых сносок в документе."""
    for style in document.styles:
        if style.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        if (style.style_id or "").casefold() in style_ids:
            return style
        if (style.name or "").strip().casefold() in style_names:
            return style
    return None


def apply_note_styles(context: RuleContext) -> None:
    """Задаёт оформление сносок и концевых сносок через их стили."""
    settings = context.profile.footnotes
    applied: list[str] = []

    for title, style_ids, style_names in (
        ("подстрочные сноски", FOOTNOTE_STYLE_IDS, FOOTNOTE_STYLE_NAMES),
        ("концевые сноски", ENDNOTE_STYLE_IDS, ENDNOTE_STYLE_NAMES),
    ):
        style = find_note_style(context.document, style_ids, style_names)
        if style is None:
            continue
        if settings.font:
            set_rfonts(style.element.get_or_add_rPr(), settings.font)
        style.font.size = Pt(settings.font_size_pt)
        paragraph_format = style.paragraph_format
        if paragraph_format is not None:
            paragraph_format.line_spacing = settings.line_spacing
        applied.append(title)

    if applied:
        context.changes.add(
            CATEGORY_FOOTNOTES,
            f"оформление сносок: {settings.font or 'шрифт не задан'} "
            f"{settings.font_size_pt:g} пт, межстрочный интервал {settings.line_spacing:g}",
            detail="изменены стили: " + ", ".join(applied),
        )
    elif context.inventory.footnotes:
        context.changes.add(
            CATEGORY_FOOTNOTES,
            "в документе есть сноски, но стиль сносок не найден — оформление сносок "
            "не изменялось, проверьте их вручную",
            detail=f"ссылок на сноски в тексте: {context.inventory.footnotes}",
        )
