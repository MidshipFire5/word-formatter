"""Разбор структуры документа: роли абзацев и инвентарь элементов.

Правила оформления ничего не «угадывают» повторно: всё, что можно понять о
документе, определяется один раз здесь. Если признаков не хватает, абзац
получает роль «спорный» и попадает в отчёт, а не оформляется наугад.
"""

from __future__ import annotations

import re
from collections import Counter

from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from word_formatter.core.docx.elements import (
    BodyItem,
    DocumentInventory,
    ParagraphInfo,
    Role,
    TableInfo,
)
from word_formatter.core.profiles import Profile

# Подпись рисунка: «Рисунок 1 – Схема», «Рис. 2.1 — Схема».
FIGURE_CAPTION = re.compile(r"^(Рисунок|Рис\.)\s*\d", re.IGNORECASE)
# Название таблицы: «Таблица 1 – Данные», «Таблица 2.4 – Расходы».
TABLE_CAPTION = re.compile(r"^Таблица\s*\d", re.IGNORECASE)
# Продолжение таблицы: «Продолжение таблицы 2.4».
TABLE_CONTINUATION = re.compile(r"^Продолжение\s+таблицы\b", re.IGNORECASE)
# Номер раздела или подраздела в начале абзаца.
NUMBERED_HEADING = re.compile(r"^(\d+(?:\.\d+){0,5})[.)]?\s+(\S.*)$", re.DOTALL)
# Ручной маркер перечисления в начале абзаца.
MANUAL_MARKER = re.compile(r"^(?:[–—•·]|-(?!-))\s+")
# Отточие: строка содержания, набранная вручную, а не заголовок.
LEADER_DOTS = re.compile(r"\.{4,}|…")
# Встроенный стиль заголовка: «Heading 1», «Заголовок 2».
HEADING_STYLE = re.compile(r"^(?:heading|заголовок)\s*(\d)$", re.IGNORECASE)
# Встроенный стиль списка.
LIST_STYLE = re.compile(r"^(?:list|список)", re.IGNORECASE)
# Граница предложения внутри строки — признак обычного текста, а не заголовка.
SENTENCE_INSIDE = re.compile(r"[.!?]\s")

MAX_HEADING_LENGTH = 150
LONG_BODY_LENGTH = 200

# Абзац заголовка, который на самом деле выглядит как обычный текст.
BODY_LOOKING_REASON = (
    "оформлен стилем заголовка, но по содержанию похож на обычный текст "
    "(длинный, с точкой в конце) — оформление не изменялось"
)
EMBEDDED_OBJECT_REASON = (
    "абзац содержит встроенный объект (формула или рисунок старого формата) — "
    "такие объекты приложение не переоформляет"
)
FIGURE_WITH_TEXT_REASON = (
    "в одном абзаце находятся рисунок и текст — разобрать их врозь приложение не может"
)
NUMBERED_CONFLICT_REASON = (
    "начинается с номера и оформлен как заголовок, но содержит признаки обычного текста — "
    "оформление не изменялось"
)


def _has_drawing(paragraph: Paragraph) -> bool:
    for tag in ("w:drawing", "w:pict"):
        if paragraph._p.find(".//" + qn(tag)) is not None:
            return True
    return False


def _has_formula(paragraph: Paragraph) -> bool:
    return paragraph._p.find(".//" + qn("m:oMath")) is not None


def _has_embedded_object(paragraph: Paragraph) -> bool:
    return paragraph._p.find(".//" + qn("w:object")) is not None


def _has_numbering(paragraph: Paragraph) -> bool:
    pPr = paragraph._p.pPr
    return pPr is not None and pPr.find(qn("w:numPr")) is not None


def _is_list_style(paragraph: Paragraph) -> bool:
    style = paragraph.style
    if style is None:
        return False
    for candidate in (style.style_id or "", style.name or ""):
        if LIST_STYLE.match(candidate.strip()):
            return True
    return False


def _style_candidates(paragraph: Paragraph) -> list[str]:
    style = paragraph.style
    if style is None:
        return []
    return [style.style_id or "", style.name or ""]


def _outline_level(paragraph: Paragraph) -> int | None:
    """Уровень заголовка по свойству outlineLvl — так его понимает сам Word.

    Сначала смотрится сам абзац, затем цепочка стилей, на которых он основан.
    """
    pPr = paragraph._p.pPr
    if pPr is not None:
        node = pPr.find(qn("w:outlineLvl"))
        if node is not None:
            value = node.get(qn("w:val"))
            if value is not None and value.lstrip("-").isdigit() and int(value) >= 0:
                return int(value) + 1

    style = paragraph.style
    seen: set[str] = set()
    while style is not None:
        style_id = style.style_id or ""
        if style_id in seen:
            break
        seen.add(style_id)
        style_pPr = style.element.find(qn("w:pPr"))
        if style_pPr is not None:
            node = style_pPr.find(qn("w:outlineLvl"))
            if node is not None:
                value = node.get(qn("w:val"))
                if value is not None and value.lstrip("-").isdigit() and int(value) >= 0:
                    return int(value) + 1
        style = style.base_style
    return None


def _style_heading_level(paragraph: Paragraph) -> int | None:
    for candidate in _style_candidates(paragraph):
        match = HEADING_STYLE.match(candidate.strip())
        if match:
            return int(match.group(1))
    return None


def _structural_match(text: str, titles: tuple[str, ...]) -> bool:
    """Совпадает ли абзац с названием структурной части из профиля."""
    normalized = " ".join(text.split()).rstrip(":").upper()
    if len(normalized) > 60:
        return False
    for title in titles:
        title_normalized = " ".join(title.split()).rstrip(":").upper()
        if not title_normalized:
            continue
        if normalized == title_normalized:
            return True
        # «ПРИЛОЖЕНИЕ А», «ПРИЛОЖЕНИЕ 1» — короткое уточнение после названия.
        if normalized.startswith(title_normalized + " ") and len(normalized) - len(title_normalized) <= 5:
            return True
    return False


def _split_number(text: str) -> tuple[tuple[str, ...], str] | None:
    """Разбирает номер в начале абзаца, отсекая числа, не похожие на номер раздела."""
    match = NUMBERED_HEADING.match(text)
    if not match:
        return None
    number, rest = match.group(1), match.group(2).strip()
    parts = tuple(number.split("."))
    if any(not part.isdigit() for part in parts):
        return None
    # «2020 год был трудным» — это не заголовок раздела, а начало предложения.
    if len(parts[0]) > 2 or int(parts[0]) < 1:
        return None
    if any(len(part) > 2 or int(part) < 1 for part in parts[1:]):
        return None
    return parts, rest


def _rest_looks_like_heading(rest: str, text: str) -> bool:
    if not rest or not rest[:1].isupper():
        return False
    if len(text) > MAX_HEADING_LENGTH:
        return False
    if LEADER_DOTS.search(text):
        # «1 Теоретические основы ………… 4» — это строка содержания,
        # набранная вручную, а не заголовок раздела.
        return False
    if rest.endswith((".", ";", ",")):
        return False
    if SENTENCE_INSIDE.search(rest):
        return False
    return True


def _looks_like_body_text(text: str) -> bool:
    """Признаки того, что абзац со стилем заголовка на самом деле обычный текст."""
    stripped = text.strip()
    return len(stripped) > LONG_BODY_LENGTH and stripped.endswith(".")


def classify_paragraph(
    paragraph: Paragraph,
    profile: Profile,
    *,
    index: int = 0,
    in_table: bool = False,
    table_index: int | None = None,
) -> ParagraphInfo:
    """Определяет роль абзаца."""
    text = paragraph.text.strip()
    drawing = _has_drawing(paragraph)
    formula = _has_formula(paragraph)
    embedded = _has_embedded_object(paragraph)
    style_name = paragraph.style.name if paragraph.style is not None else None

    def result(role: Role, level: int | None = None, reason: str = "") -> ParagraphInfo:
        return ParagraphInfo(
            index=index,
            text=paragraph.text,
            role=role,
            level=level,
            style_name=style_name,
            reason=reason,
            has_drawing=drawing,
            has_formula=formula,
            in_table=in_table,
            table_index=table_index,
        )

    if not text and not drawing and not formula and not embedded:
        return result(Role.BLANK)

    if in_table:
        return result(Role.TABLE_CELL)

    if embedded and not text:
        return result(Role.AMBIGUOUS, reason=EMBEDDED_OBJECT_REASON)

    if drawing:
        if text:
            return result(Role.AMBIGUOUS, reason=FIGURE_WITH_TEXT_REASON)
        return result(Role.FIGURE)

    if TABLE_CONTINUATION.match(text):
        return result(Role.TABLE_CONTINUATION)

    if TABLE_CAPTION.match(text):
        return result(Role.TABLE_CAPTION)

    if FIGURE_CAPTION.match(text):
        return result(Role.FIGURE_CAPTION)

    if _structural_match(text, profile.headings.structural.titles):
        return result(Role.STRUCTURAL_HEADING)

    outline_level = _outline_level(paragraph)
    style_level = _style_heading_level(paragraph)
    heading_level = outline_level or style_level

    if heading_level is not None:
        if text and _looks_like_body_text(text):
            return result(Role.AMBIGUOUS, reason=BODY_LOOKING_REASON)
        return result(Role.HEADING, level=min(heading_level, profile.headings.max_level))

    numbered = _split_number(text)
    if numbered is not None:
        parts, rest = numbered
        if _rest_looks_like_heading(rest, text):
            level = min(len(parts), profile.headings.max_level)
            return result(Role.HEADING, level=level)
        if paragraph.paragraph_format.alignment is not None or _is_bold(paragraph):
            return result(Role.AMBIGUOUS, reason=NUMBERED_CONFLICT_REASON)

    if _has_numbering(paragraph) or _is_list_style(paragraph) or MANUAL_MARKER.match(text):
        return result(Role.LIST_ITEM)

    if formula:
        return result(Role.FORMULA)

    return result(Role.BODY)


def _is_bold(paragraph: Paragraph) -> bool:
    """Полужирный ли абзац целиком.

    Смотрится и оформление абзаца, и свойства стиля: студенты оформляют
    заголовки по-разному, и оба варианта должны распознаваться.
    """
    runs = [run for run in paragraph.runs if run.text.strip()]
    if runs and all(run.bold for run in runs):
        return True
    style = paragraph.style
    if style is not None and style.font is not None and style.font.bold:
        return True
    return False


def describe_table(table: Table, profile: Profile, index: int) -> TableInfo:
    """Собирает сведения о таблице и её абзацах."""
    paragraphs: list[ParagraphInfo] = []
    merged = 0
    paragraph_index = 0

    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.tcPr
            if tcPr is not None and (
                tcPr.find(qn("w:gridSpan")) is not None or tcPr.find(qn("w:vMerge")) is not None
            ):
                merged += 1
            for paragraph in cell.paragraphs:
                paragraphs.append(
                    classify_paragraph(
                        paragraph,
                        profile,
                        index=paragraph_index,
                        in_table=True,
                        table_index=index,
                    )
                )
                paragraph_index += 1

    try:
        columns = len(table.columns)
    except (IndexError, ValueError):  # у повреждённых таблиц сетка может не читаться
        columns = max((len(row.cells) for row in table.rows), default=0)

    return TableInfo(
        index=index,
        rows=len(table.rows),
        columns=columns,
        merged_cells=merged,
        paragraphs=tuple(paragraphs),
    )


def _body_sequence(document) -> tuple[BodyItem, ...]:
    """Порядок элементов тела документа: абзацы и таблицы вперемешку."""
    items: list[BodyItem] = []
    paragraph_index = 0
    table_index = 0
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            items.append(BodyItem(kind="paragraph", index=paragraph_index))
            paragraph_index += 1
        elif child.tag == qn("w:tbl"):
            items.append(BodyItem(kind="table", index=table_index))
            table_index += 1
    return tuple(items)


def _count(root, tag: str) -> int:
    return sum(1 for _ in root.iter(qn(tag)))


def inspect_document(document, profile: Profile) -> DocumentInventory:
    """Полный разбор документа: роли абзацев, таблицы и состав элементов."""
    paragraphs = tuple(
        classify_paragraph(paragraph, profile, index=index)
        for index, paragraph in enumerate(document.paragraphs)
    )
    tables = tuple(
        describe_table(table, profile, index) for index, table in enumerate(document.tables)
    )

    body = document.element.body
    field_begin = sum(
        1
        for node in body.iter(qn("w:fldChar"))
        if node.get(qn("w:fldCharType")) == "begin"
    )
    toc_fields = sum(
        1 for node in body.iter(qn("w:instrText")) if "TOC" in (node.text or "").upper()
    )
    page_breaks = sum(
        1 for node in body.iter(qn("w:br")) if node.get(qn("w:type")) == "page"
    )

    counts = Counter(item.role for item in paragraphs)

    return DocumentInventory(
        paragraphs=paragraphs,
        tables=tables,
        body=_body_sequence(document),
        counts=dict(counts),
        images=_count(body, "w:drawing") + _count(body, "w:pict"),
        hyperlinks=_count(body, "w:hyperlink"),
        footnotes=_count(body, "w:footnoteReference"),
        endnotes=_count(body, "w:endnoteReference"),
        sections=len(document.sections),
        fields=field_begin,
        toc_fields=toc_fields,
        ole_objects=_count(body, "w:object"),
        content_controls=_count(body, "w:sdt"),
        page_breaks=page_breaks,
        page_break_before=_count(body, "w:pageBreakBefore"),
    )
