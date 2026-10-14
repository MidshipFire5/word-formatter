"""Правила оформления страницы: формат листа, ориентация, поля, колонтитулы."""

from __future__ import annotations

from docx.enum.section import WD_ORIENT
from docx.shared import Mm

from word_formatter.core.rules.context import RuleContext

CATEGORY = "page"


def apply_page_settings(context: RuleContext) -> None:
    """Приводит параметры страницы к профилю во всех разделах документа."""
    page = context.profile.page
    width_mm, height_mm = page.size_mm
    orientation = WD_ORIENT.LANDSCAPE if page.orientation == "landscape" else WD_ORIENT.PORTRAIT

    sections = list(context.document.sections)
    for section in sections:
        section.orientation = orientation
        section.page_width = Mm(width_mm)
        section.page_height = Mm(height_mm)
        section.left_margin = Mm(page.margins.left)
        section.right_margin = Mm(page.margins.right)
        section.top_margin = Mm(page.margins.top)
        section.bottom_margin = Mm(page.margins.bottom)
        section.header_distance = Mm(page.header_mm)
        section.footer_distance = Mm(page.footer_mm)

    margins = page.margins
    context.changes.add(
        CATEGORY,
        f"параметры страницы: {page.format}, ориентация "
        f"{'альбомная' if page.orientation == 'landscape' else 'книжная'}, "
        f"поля {margins.left:g}/{margins.right:g}/{margins.top:g}/{margins.bottom:g} мм "
        "(лево/право/верх/низ)",
        detail=f"разделов документа: {len(sections)}",
    )
