"""Измерение оформления в готовом документе.

Если методичка сама оформлена по своим требованиям, оформление можно
не вычитывать из текста, а просто измерить: поля страницы, шрифт и размер
основного текста, интервал, абзацный отступ, стили заголовков, наличие
нумерации страниц и автоматических переносов.

Измеренные значения считаются вспомогательными: они описывают то, как файл
оформлен сейчас, а не то, что в нём написано. Если текст и оформление
расходятся, приложение берёт текст и сообщает о расхождении.
"""

from __future__ import annotations

from docx.oxml.ns import qn

from word_formatter.core.profiles.extraction.draft import ORIGIN_SAMPLE, DraftValue

# Стандартные имена стилей в русской версии Word.
HEADING_STYLE_NAMES = (
    "Heading 1",
    "Heading 2",
    "Heading 3",
    "Заголовок 1",
    "Заголовок 2",
    "Заголовок 3",
)
FOOTNOTE_STYLE_NAMES = ("Footnote Text", "Текст сноски")

ALIGNMENT_VALUES = {
    "left": "left",
    "center": "center",
    "right": "right",
    "both": "justify",
    "justify": "justify",
    "distribute": "justify",
}


def _value(path: str, title: str, value, evidence: str) -> DraftValue:
    return DraftValue(
        path=path, value=value, origin=ORIGIN_SAMPLE, evidence=evidence, title=title
    )


def _alignment_of(paragraph_format, paragraph_element) -> str | None:
    """Выравнивание абзаца: сначала прямое, потом из свойств XML."""
    value = paragraph_format.alignment
    if value is not None:
        name = str(value).split(" ")[0].lower()
        if name in ALIGNMENT_VALUES:
            return ALIGNMENT_VALUES[name]

    pPr = paragraph_element.find(qn("w:pPr")) if paragraph_element is not None else None
    jc = pPr.find(qn("w:jc")) if pPr is not None else None
    if jc is not None:
        return ALIGNMENT_VALUES.get(jc.get(qn("w:val")), None)
    return None


def _effective_rpr(style):
    """Свойства шрифта стиля: сам стиль, затем стиль-основа."""
    return style.element.get_or_add_rPr()


def extract_from_sample(document) -> list[DraftValue]:
    """Измеряет оформление документа-образца."""
    values: list[DraftValue] = []

    # --- страница ---
    section = document.sections[0]
    evidence = "раздел 1 документа"
    values.extend(
        [
            _value("page.marginsMm.left", "левое поле", round(section.left_margin.mm), evidence),
            _value("page.marginsMm.right", "правое поле", round(section.right_margin.mm), evidence),
            _value("page.marginsMm.top", "верхнее поле", round(section.top_margin.mm), evidence),
            _value(
                "page.marginsMm.bottom", "нижнее поле", round(section.bottom_margin.mm), evidence
            ),
            _value(
                "page.headerMm",
                "расстояние до верхнего колонтитула",
                round(section.header_distance.mm),
                evidence,
            ),
            _value(
                "page.footerMm",
                "расстояние до нижнего колонтитула",
                round(section.footer_distance.mm),
                evidence,
            ),
        ]
    )

    width_mm, height_mm = round(section.page_width.mm), round(section.page_height.mm)
    if (width_mm, height_mm) == (210, 297):
        values.append(_value("page.format", "формат листа", "A4", evidence))

    # --- основной текст ---
    normal = document.styles["Normal"]
    normal_format = normal.paragraph_format
    style_evidence = "стиль «Обычный»"
    if normal.font.name:
        values.append(_value("body.font", "шрифт основного текста", normal.font.name, style_evidence))
    if normal.font.size is not None:
        values.append(
            _value("body.sizePt", "размер шрифта основного текста", normal.font.size.pt, style_evidence)
        )

    normal_element = document.styles["Normal"].element
    alignment = _alignment_of(normal_format, normal_element)
    if alignment:
        values.append(_value("body.alignment", "выравнивание основного текста", alignment, style_evidence))
    if normal_format.first_line_indent is not None:
        values.append(
            _value(
                "body.firstLineIndentCm",
                "абзацный отступ",
                round(normal_format.first_line_indent.cm, 2),
                style_evidence,
            )
        )
    if normal_format.line_spacing is not None and isinstance(normal_format.line_spacing, float):
        values.append(
            _value(
                "body.lineSpacing",
                "межстрочный интервал",
                round(float(normal_format.line_spacing), 2),
                style_evidence,
            )
        )

    # --- заголовки ---
    for level, style_name in enumerate(HEADING_STYLE_NAMES[:3], start=1):
        style = document.styles[style_name] if style_name in [s.name for s in document.styles] else None
        if style is None:
            continue
        prefix = f"headings.levels.{level}"
        if style.font.name:
            values.append(_value(f"{prefix}.font", f"шрифт заголовка {level} уровня", style.font.name, f"стиль «{style_name}»"))
        if style.font.size is not None:
            values.append(
                _value(
                    f"{prefix}.sizePt",
                    f"размер шрифта заголовка {level} уровня",
                    style.font.size.pt,
                    f"стиль «{style_name}»",
                )
            )
        if style.font.bold is not None:
            values.append(
                _value(
                    f"{prefix}.bold",
                    f"полужирный заголовок {level} уровня",
                    bool(style.font.bold),
                    f"стиль «{style_name}»",
                )
            )
        alignment = _alignment_of(style.paragraph_format, style.element)
        if alignment:
            values.append(
                _value(
                    f"{prefix}.alignment",
                    f"выравнивание заголовка {level} уровня",
                    alignment,
                    f"стиль «{style_name}»",
                )
            )
        indent = style.paragraph_format.first_line_indent
        if indent is not None:
            values.append(
                _value(
                    f"{prefix}.firstLineIndentCm",
                    f"абзацный отступ заголовка {level} уровня",
                    round(indent.cm, 2),
                    f"стиль «{style_name}»",
                )
            )

    present_levels = [
        level
        for level, style_name in enumerate(HEADING_STYLE_NAMES[:3], start=1)
        if style_name in [s.name for s in document.styles]
    ]
    if present_levels:
        values.append(
            _value("headings.maxLevel", "число уровней заголовков", max(present_levels), "стили заголовков документа")
        )

    # --- структурные заголовки: ищем по тексту документа ---
    structural = _extract_structural_heading(document)
    values.extend(structural)

    # --- сноски ---
    for style_name in FOOTNOTE_STYLE_NAMES:
        names = [style.name for style in document.styles]
        if style_name not in names:
            continue
        style = document.styles[style_name]
        if style.font.size is not None:
            values.append(
                _value(
                    "footnotes.fontSizePt",
                    "размер шрифта сносок",
                    style.font.size.pt,
                    f"стиль «{style_name}»",
                )
            )
        break

    # --- таблицы ---
    for table in document.tables:
        size = _first_cell_font_size(table)
        if size is not None:
            values.append(
                _value("tables.fontSizePt", "размер шрифта в таблицах", size, "первая ячейка таблицы")
            )
            break

    # --- переносы ---
    settings = document.settings.element
    values.append(
        _value(
            "hyphenation.enabled",
            "автоматические переносы",
            settings.find(qn("w:autoHyphenation")) is not None,
            "настройки документа",
        )
    )

    # --- нумерация страниц ---
    values.extend(_extract_page_numbers(document))

    # --- свободные строки ---
    values.append(
        _value(
            "blankLines.insert",
            "свободные строки",
            _has_blank_lines(document),
            "пустые абзацы вокруг заголовков и таблиц",
        )
    )

    return values


def _extract_structural_heading(document) -> list[DraftValue]:
    """Структурный заголовок: ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и подобные."""
    titles = ("ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СОДЕРЖАНИЕ", "ПРИЛОЖЕНИЕ")
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        if not any(text.upper().startswith(title) for title in titles):
            continue

        evidence = f"абзац «{text[:40]}»"
        values: list[DraftValue] = []
        alignment = _alignment_of(paragraph.paragraph_format, paragraph._p)
        if alignment:
            values.append(
                _value("headings.structural.alignment", "выравнивание структурного заголовка", alignment, evidence)
            )

        caps = False
        bold = False
        size = None
        for run in paragraph.runs:
            rpr = run._element.find(qn("w:rPr"))
            if rpr is not None and rpr.find(qn("w:caps")) is not None:
                caps = True
            if run.font.bold:
                bold = True
            if run.font.size is not None:
                size = run.font.size.pt

        if text.isupper() or caps:
            values.append(
                _value("headings.structural.uppercase", "структурные заголовки прописными", True, evidence)
            )
        if bold:
            values.append(
                _value("headings.structural.bold", "структурные заголовки полужирным", True, evidence)
            )
        if size is not None:
            values.append(
                _value("headings.structural.sizePt", "размер шрифта структурного заголовка", size, evidence)
            )
        return values
    return []


def _extract_page_numbers(document) -> list[DraftValue]:
    section = document.sections[0]
    values: list[DraftValue] = []
    evidence = "нижний колонтитул"

    texts = []
    for footer in (section.footer, section.first_page_footer):
        try:
            texts.append(footer.paragraphs[0] if footer.paragraphs else None)
        except Exception:
            texts.append(None)

    has_page_field = False
    position = "bottom-center"
    size = None
    for paragraph in texts:
        if paragraph is None:
            continue
        xml = paragraph._p.xml
        if "PAGE" in xml:
            has_page_field = True
            alignment = _alignment_of(paragraph.paragraph_format, paragraph._p)
            position = {
                "center": "bottom-center",
                "right": "bottom-right",
                "left": "bottom-left",
            }.get(alignment or "center", "bottom-center")
            for run in paragraph.runs:
                if run.font.size is not None:
                    size = run.font.size.pt

    values.append(
        _value("pageNumbers.enabled", "нумерация страниц", has_page_field, evidence)
    )
    if has_page_field:
        values.append(_value("pageNumbers.position", "положение номера страницы", position, evidence))
        if size is not None:
            values.append(
                _value("pageNumbers.fontSizePt", "размер шрифта номера страницы", size, evidence)
            )
    values.append(
        _value(
            "pageNumbers.skipFirstPage",
            "номер со второй страницы",
            bool(section.different_first_page_header_footer),
            "свойства раздела (особый колонтитул первой страницы)",
        )
    )
    return values


def _first_cell_font_size(table) -> float | None:
    try:
        cell = table.cell(0, 0)
    except Exception:
        return None
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            if run.font.size is not None:
                return run.font.size.pt
    return None


def _has_blank_lines(document) -> bool:
    """Есть ли в образце пустые абзацы: значит, их требуют оставлять."""
    body = list(document.element.body)
    empty = 0
    for node in body:
        if node.tag != qn("w:p"):
            continue
        if not "".join(node.itertext()).strip():
            empty += 1
    return empty >= 3
