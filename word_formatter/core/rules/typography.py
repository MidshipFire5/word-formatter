"""Применение оформления к абзацам и прогонам текста.

Здесь собраны все «мелочи», из-за которых форматирование обычно не срабатывает:
ссылки на шрифт темы, отступы в символах, автоматические интервалы.
Всё это снимается явно, иначе Word продолжит показывать старое оформление.
"""

from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from word_formatter.core.profiles import TextStyle
from word_formatter.core.rules.xml_helpers import (
    PPR_ORDER,
    RPR_ORDER,
    ensure_ordered_child,
    set_rfonts,
)
ALIGNMENT = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# Атрибуты, которые перебивают заданные значения и потому удаляются.
INDENT_CONFLICTS = ("w:firstLineChars", "w:hangingChars")
SPACING_CONFLICTS = (
    "w:beforeLines",
    "w:afterLines",
    "w:beforeAutospacing",
    "w:afterAutospacing",
)


def paragraph_runs(paragraph: Paragraph) -> list[Run]:
    """Все текстовые прогоны абзаца, включая прогоны внутри гиперссылок."""
    return [Run(element, paragraph) for element in paragraph._p.iter(qn("w:r"))]


def run_has_text(run: Run) -> bool:
    """Есть ли в прогоне обычный текст.

    Прогоны со сносками, полями и разрывами текста не содержат: менять им
    шрифт нельзя — сноска потеряет свой вид, а поле перестанет работать.
    """
    return run._element.find(qn("w:t")) is not None


def set_first_line_indent(paragraph: Paragraph, indent_cm: float) -> None:
    """Задаёт отступ первой строки, снимая конкурирующие настройки."""
    pPr = paragraph._p.get_or_add_pPr()
    indent = ensure_ordered_child(pPr, "w:ind", PPR_ORDER)
    for attribute in INDENT_CONFLICTS:
        key = qn(attribute)
        if key in indent.attrib:
            del indent.attrib[key]
    indent.set(qn("w:firstLine"), str(Cm(indent_cm).twips))
    hanging = qn("w:hanging")
    if hanging in indent.attrib:
        del indent.attrib[hanging]


def set_spacing(paragraph: Paragraph, *, before_pt: float, after_pt: float, line_spacing: float) -> None:
    """Задаёт интервалы абзаца, снимая автоматические интервалы Word."""
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(before_pt)
    paragraph_format.space_after = Pt(after_pt)
    paragraph_format.line_spacing = line_spacing

    pPr = paragraph._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is not None:
        for attribute in SPACING_CONFLICTS:
            key = qn(attribute)
            if key in spacing.attrib:
                del spacing.attrib[key]


def apply_paragraph_style(paragraph: Paragraph, style: TextStyle) -> None:
    """Применяет оформление абзаца."""
    paragraph.paragraph_format.alignment = ALIGNMENT.get(style.alignment)
    set_first_line_indent(paragraph, style.first_line_indent_cm)
    set_spacing(
        paragraph,
        before_pt=style.space_before_pt,
        after_pt=style.space_after_pt,
        line_spacing=style.line_spacing,
    )


def apply_run_style(
    run: Run,
    *,
    font: str | None,
    size_pt: float | None,
    italic: bool | None = None,
    underline: bool | None = None,
    bold: bool | None = None,
    all_caps: bool | None = None,
    color: str | None = None,
) -> None:
    """Применяет оформление прогона текста."""
    rPr = run._element.get_or_add_rPr()

    if font:
        set_rfonts(rPr, font)
    if size_pt is not None:
        half_points = str(int(round(size_pt * 2)))
        ensure_ordered_child(rPr, "w:sz", RPR_ORDER).set(qn("w:val"), half_points)
        ensure_ordered_child(rPr, "w:szCs", RPR_ORDER).set(qn("w:val"), half_points)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if underline is not None:
        run.font.underline = underline
    if all_caps is not None:
        # Прописные буквы включаются оформлением, а не правкой текста:
        # содержание документа остаётся прежним.
        run.font.all_caps = all_caps
    if color:
        from docx.shared import RGBColor

        run.font.color.rgb = RGBColor.from_string(color.lstrip("#"))


def suppress_hyphenation(paragraph: Paragraph) -> None:
    """Запрещает автоматический перенос слов в абзаце (заголовки, подписи)."""
    pPr = paragraph._p.get_or_add_pPr()
    element = ensure_ordered_child(pPr, "w:suppressAutoHyphenation", PPR_ORDER)
    element.set(qn("w:val"), "true")


def apply_style_text_style(style, text_style: TextStyle) -> None:
    """Задаёт оформление определению стиля (а не отдельному абзацу).

    Цвет задаётся всегда: если профиль его не указал, ставится «авто»
    (обычный цвет текста). Это важно, потому что встроенные стили Word
    заголовков окрашены в синий — а мы эти стили назначаем абзацам, и без
    явного цвета документ получил бы синие заголовки, которых у автора
    не было.
    """
    rpr = style.element.get_or_add_rPr()
    if text_style.font:
        set_rfonts(rpr, text_style.font)

    style.font.size = Pt(text_style.size_pt)
    style.font.bold = text_style.bold
    style.font.italic = text_style.italic

    color = ensure_ordered_child(rpr, "w:color", RPR_ORDER)
    color.set(qn("w:val"), text_style.color.lstrip("#") if text_style.color else "auto")

    paragraph_format = style.paragraph_format
    if paragraph_format is None:
        return
    paragraph_format.alignment = ALIGNMENT.get(text_style.alignment)
    paragraph_format.first_line_indent = Cm(text_style.first_line_indent_cm)
    paragraph_format.line_spacing = text_style.line_spacing
    paragraph_format.space_before = Pt(text_style.space_before_pt)
    paragraph_format.space_after = Pt(text_style.space_after_pt)


def effective_italic(run: Run, paragraph: Paragraph) -> bool:
    """Был ли текст курсивом — с учётом стиля абзаца, а не только прогона."""
    if run.italic is not None:
        return bool(run.italic)
    style = paragraph.style
    if style is not None and style.font is not None and style.font.italic:
        return True
    return False


def effective_underline(run: Run, paragraph: Paragraph) -> bool:
    """Было ли подчёркивание — с учётом стиля абзаца."""
    if run.underline is not None:
        return bool(run.underline)
    style = paragraph.style
    if style is not None and style.font is not None and style.font.underline:
        return True
    return False
