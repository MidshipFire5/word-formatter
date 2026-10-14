"""Правила нумерации страниц.

Номер страницы — это поле Word, а не текст: иначе при вставке или удалении
страницы номер «застынет». Поле ставится в колонтитул, а на первой странице
(титульном листе) номер скрывается настройкой «особый колонтитул первой
страницы».
"""

from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.typography import apply_run_style, paragraph_runs
from word_formatter.core.rules.xml_helpers import (
    SETTINGS_ORDER,
    XML_SPACE,
    ensure_ordered_child,
)

CATEGORY_PAGE_NUMBERS = "page-numbers"

POSITION_ALIGNMENT = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}


def split_position(position: str) -> tuple[str, str]:
    """«bottom-center» → («bottom», «center»)."""
    where, _, alignment = position.partition("-")
    return where, alignment or "center"


def has_page_field(paragraph) -> bool:
    """Есть ли в абзаце поле номера страницы."""
    for node in paragraph._p.iter(qn("w:instrText")):
        text = (node.text or "").upper()
        if "PAGE" in text and "NUMPAGES" not in text:
            return True
    return False


def append_page_field(paragraph) -> None:
    """Добавляет в абзац поле PAGE с готовым результатом."""
    def run_with(child) -> OxmlElement:
        run = OxmlElement("w:r")
        run.append(child)
        return run

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")

    instruction = OxmlElement("w:instrText")
    instruction.set(XML_SPACE, "preserve")
    instruction.text = " PAGE "

    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")

    result = OxmlElement("w:t")
    result.text = "2"  # Word пересчитает значение при открытии документа

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    for child in (begin, instruction, separate, result, end):
        paragraph._p.append(run_with(child))


def _target_paragraph(container) -> tuple[object, bool]:
    """Возвращает абзац для поля номера страницы и признак «создан заново»."""
    paragraphs = container.paragraphs
    for paragraph in paragraphs:
        if has_page_field(paragraph):
            return paragraph, False
    for paragraph in paragraphs:
        if not paragraph.text.strip() and not paragraph._p.findall(qn("w:r")):
            return paragraph, False
    return container.add_paragraph(), True


def _enable_field_update(context: RuleContext) -> None:
    """Просит Word обновить поля при открытии документа.

    Без этого номера страниц и автоматическое содержание останутся такими,
    какими их сохранила библиотека.
    """
    settings_element = context.document.settings.element
    node = ensure_ordered_child(settings_element, "w:updateFields", SETTINGS_ORDER)
    node.set(qn("w:val"), "true")


def _make_numbering_continuous(context: RuleContext) -> int:
    """Убирает сбросы нумерации: методичка требует сквозную нумерацию."""
    fixed = 0
    for section in context.document.sections:
        pgNumType = section._sectPr.find(qn("w:pgNumType"))
        if pgNumType is None:
            continue
        if pgNumType.get(qn("w:start")) is not None:
            del pgNumType.attrib[qn("w:start")]
            fixed += 1
        pgNumType.set(qn("w:fmt"), "decimal")
    return fixed


def apply_page_numbers(context: RuleContext) -> None:
    """Ставит или приводит к профилю номер страницы в колонтитуле."""
    settings = context.profile.page_numbers
    if not settings.enabled:
        context.changes.add(
            CATEGORY_PAGE_NUMBERS,
            "нумерация страниц отключена профилем — номера не проставлялись",
        )
        return

    where, alignment_key = split_position(settings.position)
    alignment = POSITION_ALIGNMENT.get(alignment_key, WD_ALIGN_PARAGRAPH.CENTER)

    added = 0
    updated = 0
    inherited = 0

    for index, section in enumerate(context.document.sections):
        container = section.header if where == "top" else section.footer
        # Раздел без своего колонтитула наследует предыдущий: добавлять туда
        # второй номер нельзя — получится дубль в общем колонтитуле.
        if container.is_linked_to_previous and index > 0:
            inherited += 1
            continue

        paragraph, created = _target_paragraph(container)
        if created:
            append_page_field(paragraph)
            added += 1
        elif not has_page_field(paragraph):
            append_page_field(paragraph)
            added += 1
        else:
            updated += 1

        paragraph.paragraph_format.alignment = alignment
        for run in paragraph_runs(paragraph):
            apply_run_style(run, font=settings.font, size_pt=settings.font_size_pt)

        if settings.skip_first_page and index == 0:
            section.different_first_page_header_footer = True

    _enable_field_update(context)
    restarts = _make_numbering_continuous(context)

    place = "внизу" if where == "bottom" else "вверху"
    where_text = {"center": "по центру", "left": "слева", "right": "справа"}.get(
        alignment_key, "по центру"
    )
    description = (
        f"нумерация страниц: {place} {where_text}, {settings.font or 'шрифт не задан'} "
        f"{settings.font_size_pt:g} пт, сквозная нумерация арабскими цифрами"
    )
    if settings.skip_first_page:
        description += ", на первой странице (титульном листе) номер не проставляется"

    details = []
    if added:
        details.append(f"номер добавлен в разделах: {added}")
    if updated:
        details.append(f"оформление номера обновлено: {updated}")
    if inherited:
        details.append(f"разделов наследуют колонтитул: {inherited}")
    if restarts:
        details.append(f"снято сбросов нумерации: {restarts}")
    details.append("поля помечены для обновления при открытии документа")

    context.changes.add(CATEGORY_PAGE_NUMBERS, description, detail="; ".join(details))
