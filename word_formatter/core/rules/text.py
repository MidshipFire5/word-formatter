"""Правила оформления основного текста.

Дополнительно задаётся оформление по умолчанию — то, что получает текст,
у которого нет прямого оформления. Без этого шага Word продолжил бы
показывать шрифт темы оформления (Calibri или Aptos) там, где оформление
не задано вручную.
"""

from __future__ import annotations

from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from word_formatter.core.docx import Role
from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.typography import (
    ALIGNMENT,
    apply_paragraph_style,
    apply_run_style,
    effective_italic,
    effective_underline,
    paragraph_runs,
    run_has_text,
)
from word_formatter.core.rules.xml_helpers import (
    PPR_ORDER,
    RPR_ORDER,
    ensure_child,
    ensure_ordered_child,
    set_rfonts,
)

CATEGORY_BODY = "body-text"
CATEGORY_DEFAULTS = "document-defaults"
CATEGORY_ITALIC = "italic-removed"
CATEGORY_UNDERLINE = "underline-removed"
CATEGORY_TITLE_PAGE = "title-page"

DOC_DEFAULTS_ORDER = ("w:rPrDefault", "w:pPrDefault")


def title_page_end(context: RuleContext) -> int | None:
    """Номер абзаца, с которого начинается текст работы.

    Всё, что стоит до первой структурной части (СОДЕРЖАНИЕ, ВВЕДЕНИЕ,
    ЗАКЛЮЧЕНИЕ), — это титульный лист. Методички обычно задают его отдельным
    образцом, которого у приложения нет, поэтому титульный лист остаётся
    таким, каким его сделал автор, и попадает в отчёт.
    """
    for item in context.inventory.paragraphs:
        if item.role is Role.STRUCTURAL_HEADING:
            return item.index
    return None


def _set_half_point_size(rpr, size_pt: float) -> None:
    half_points = str(int(round(size_pt * 2)))
    ensure_ordered_child(rpr, "w:sz", RPR_ORDER).set(qn("w:val"), half_points)
    ensure_ordered_child(rpr, "w:szCs", RPR_ORDER).set(qn("w:val"), half_points)


def apply_document_defaults(context: RuleContext) -> None:
    """Задаёт оформление по умолчанию: стиль «Обычный» и значения документа."""
    body = context.profile.body
    document = context.document

    normal = document.styles["Normal"]
    if body.font:
        set_rfonts(normal.element.get_or_add_rPr(), body.font)
    normal.font.size = Pt(body.size_pt)
    normal.font.italic = False
    normal.paragraph_format.alignment = ALIGNMENT.get(body.alignment)
    normal.paragraph_format.first_line_indent = Cm(body.first_line_indent_cm)
    normal.paragraph_format.line_spacing = body.line_spacing
    normal.paragraph_format.space_before = Pt(body.space_before_pt)
    normal.paragraph_format.space_after = Pt(body.space_after_pt)

    styles_element = document.styles.element
    doc_defaults = ensure_child(styles_element, "w:docDefaults", prepend=True)
    rpr_default = ensure_ordered_child(doc_defaults, "w:rPrDefault", DOC_DEFAULTS_ORDER)
    ppr_default = ensure_ordered_child(doc_defaults, "w:pPrDefault", DOC_DEFAULTS_ORDER)

    rpr = ensure_ordered_child(rpr_default, "w:rPr", RPR_ORDER)
    if body.font:
        set_rfonts(rpr, body.font)
    _set_half_point_size(rpr, body.size_pt)

    ppr = ensure_ordered_child(ppr_default, "w:pPr", PPR_ORDER)
    spacing = ensure_ordered_child(ppr, "w:spacing", PPR_ORDER)
    spacing.set(qn("w:line"), str(int(round(body.line_spacing * 240))))
    spacing.set(qn("w:lineRule"), "auto")
    spacing.set(qn("w:before"), str(int(round(body.space_before_pt * 20))))
    spacing.set(qn("w:after"), str(int(round(body.space_after_pt * 20))))

    context.changes.add(
        CATEGORY_DEFAULTS,
        f"оформление по умолчанию: {body.font or 'шрифт не задан'} {body.size_pt:g} пт, "
        f"межстрочный интервал {body.line_spacing:g}, интервалы до и после абзаца "
        f"{body.space_before_pt:g} и {body.space_after_pt:g} пт",
        detail="изменены стиль «Обычный» и значения документа",
    )


def apply_paragraph_and_runs(
    context: RuleContext,
    paragraph,
    info,
    style,
    *,
    remove_italic: bool,
    remove_underline: bool,
    italic_category: str = CATEGORY_ITALIC,
    underline_category: str = CATEGORY_UNDERLINE,
) -> tuple[int, int]:
    """Оформляет абзац и его прогоны.

    Возвращает количество снятых курсивов и подчёркиваний. Общая часть для
    основного текста и элементов перечислений: правила у них одинаковые,
    отличается только отступ и раздел отчёта.
    """
    apply_paragraph_style(paragraph, style)
    italic_removed = 0
    underline_removed = 0

    for run in paragraph_runs(paragraph):
        if not run_has_text(run):
            continue
        was_italic = effective_italic(run, paragraph)
        was_underline = effective_underline(run, paragraph)

        apply_run_style(
            run,
            font=style.font,
            size_pt=style.size_pt,
            italic=False if remove_italic else None,
            underline=False if remove_underline else None,
            color=style.color,
        )

        if remove_italic and was_italic:
            italic_removed += 1
            context.changes.add(
                italic_category,
                "снят курсив",
                paragraph_index=info.index,
                text=info.short_text,
            )
        if remove_underline and was_underline:
            underline_removed += 1
            context.changes.add(
                underline_category,
                "снято подчёркивание",
                paragraph_index=info.index,
                text=info.short_text,
            )

    return italic_removed, underline_removed


def apply_body_text(context: RuleContext) -> None:
    """Оформляет абзацы основного текста.

    Абзацы титульного листа (до первой структурной части) не трогаются:
    их оформление методичка не задаёт, и портить готовую вёрстку нельзя.
    """
    body = context.profile.body
    boundary = title_page_end(context)
    processed = 0
    italic_removed = 0
    underline_removed = 0
    title_page_paragraphs = 0

    for info in context.inventory.paragraphs:
        if info.role is not Role.BODY:
            continue
        if boundary is not None and info.index < boundary:
            if info.text.strip():
                title_page_paragraphs += 1
            continue

        paragraph = context.paragraph(info.index)
        italics, underlines = apply_paragraph_and_runs(
            context,
            paragraph,
            info,
            body,
            remove_italic=body.remove_italic,
            remove_underline=body.remove_underline,
        )
        italic_removed += italics
        underline_removed += underlines
        processed += 1

    if title_page_paragraphs:
        context.changes.add(
            CATEGORY_TITLE_PAGE,
            "титульный лист оставлен без изменений: методичка не задаёт его оформление, "
            "проверьте его вручную",
            detail=f"абзацев до первой структурной части: {title_page_paragraphs}",
        )

    context.changes.add(
        CATEGORY_BODY,
        f"оформление основного текста: {body.font or 'шрифт не задан'} {body.size_pt:g} пт, "
        f"выравнивание {alignment_title(body.alignment)}, "
        f"отступ первой строки {body.first_line_indent_cm:g} см, "
        f"межстрочный интервал {body.line_spacing:g}",
        detail=f"обработано абзацев: {processed}; снято курсива: {italic_removed}; "
        f"снято подчёркиваний: {underline_removed}",
    )


def alignment_title(alignment: str) -> str:
    """Название выравнивания по-русски — для отчёта."""
    return {
        "left": "по левому краю",
        "center": "по центру",
        "right": "по правому краю",
        "justify": "по ширине",
    }.get(alignment, alignment)
