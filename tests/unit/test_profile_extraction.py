"""Создание профиля из методички: разбор требований и измерение оформления."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from word_formatter.core.errors import MethodologyError, UnsupportedFormatError
from word_formatter.core.profiles import load_profile_from_path, validate_profile
from word_formatter.core.profiles.extraction import (
    ORIGIN_EDIT,
    ORIGIN_SAMPLE,
    ORIGIN_TEXT,
    build_profile_draft,
    make_profile_id,
    read_source,
    save_draft_profile,
)
from word_formatter.core.profiles.extraction.draft import coerce_value
from word_formatter.core.profiles.extraction.from_text import (
    Chunk,
    canon,
    chunks,
    extract_captions,
    extract_emphasis,
    extract_headings,
    extract_hyphenation,
    extract_margins,
    extract_page_numbers,
    extract_sizes,
    extract_spacing,
)

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
PDF_METHODOLOGY = FIXTURES / "требования-методички.pdf"

REQUIREMENTS = """МЕТОДИЧЕСКИЕ УКАЗАНИЯ

1. Общие требования
Текст работы набирается шрифтом Times New Roman, размер 14 пт, межстрочный интервал 1,5.
Абзацный отступ должен быть одинаковым и равным 1,25 см.
Поля: левое — 30 мм, правое — 15 мм, верхнее и нижнее — 20 мм.
Выравнивание основного текста — по ширине.
Курсив и подчёркивание в тексте не допускаются.

2. Заголовки
Заголовки структурных элементов (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ) печатают прописными буквами, полужирным шрифтом, без абзацного отступа.

3. Таблицы и рисунки
Шрифт в таблицах — 12 пт, интервал одинарный.
Подпись рисунка размещают под рисунком, название таблицы — над таблицей.
Между таблицей и текстом оставляют одну свободную строку.

4. Нумерация
Нумерацию страниц проставляют арабскими цифрами внизу по центру, начиная со второй страницы.
В тексте включают автоматическую расстановку переносов.
"""


def piece(text: str, previous: str = "") -> Chunk:
    """Кусок текста так, как его видит разбор."""
    return Chunk(canon=canon(text), text=text, previous=canon(previous) if previous else "")


@pytest.fixture()
def methodology(tmp_path: Path) -> Path:
    path = tmp_path / "методичка.txt"
    path.write_text(REQUIREMENTS, encoding="utf-8")
    return path


# --- разбор текста требований -------------------------------------------------


def test_chunks_split_and_keep_quotes() -> None:
    parts = chunks(REQUIREMENTS)

    assert any(part.text.startswith("Поля: левое") for part in parts)
    assert all(part.text.strip() == part.text for part in parts)


def test_font_and_size(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    assert draft.value_for("body.font").value == "Times New Roman"
    assert draft.value_for("body.sizePt").value == 14


def test_spacing_and_indent(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    assert draft.value_for("body.lineSpacing").value == 1.5
    assert draft.value_for("body.firstLineIndentCm").value == 1.25
    assert draft.value_for("body.alignment").value == "justify"


def test_margins_from_one_sentence(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    assert draft.value_for("page.marginsMm.left").value == 30
    assert draft.value_for("page.marginsMm.right").value == 15
    assert draft.value_for("page.marginsMm.top").value == 20
    assert draft.value_for("page.marginsMm.bottom").value == 20


def test_caption_positions(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    assert draft.value_for("captions.figure.position").value == "below"
    assert draft.value_for("captions.table.position").value == "above"


def test_page_numbers(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    assert draft.value_for("pageNumbers.enabled").value is True
    assert draft.value_for("pageNumbers.position").value == "bottom-center"
    assert draft.value_for("pageNumbers.skipFirstPage").value is True


def test_emphasis_and_hyphenation(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    assert draft.value_for("body.removeItalic").value is True
    assert draft.value_for("body.removeUnderline").value is True
    assert draft.value_for("hyphenation.enabled").value is True


# --- разбор методички, у которой в PDF теряются пробелы ------------------------


def test_size_list_is_parsed_in_order() -> None:
    """«Основного текста, в таблицах, в сносках: 14, 12, 10 пунктов»."""
    text = (
        "Размершрифта: основноготекста втаблицах вподстрочныхсносках "
        "14пунктов 12пунктов 10пунктов"
    )
    found = {item.path: item.value for item in extract_sizes(piece(text))}

    assert found == {
        "body.sizePt": 14.0,
        "tables.fontSizePt": 12.0,
        "footnotes.fontSizePt": 10.0,
    }


def test_size_of_tables_is_not_taken_for_the_body() -> None:
    found = {item.path: item.value for item in extract_sizes(piece("Шрифт в таблицах — 12 пт."))}

    assert found == {"tables.fontSizePt": 12.0}


def test_body_spacing_wins_over_heading_spacing() -> None:
    """Межстрочный интервал — это основной текст, даже если рядом таблицы."""
    text = (
        "Размершрифта: основноготекста втаблицах вподстрочныхсносках 14пунктов "
        "12пунктов 10пунктов Межстрочныйинтервал1 ,5интервала"
    )
    found = {item.path: item.value for item in extract_spacing(piece(text))}

    assert found["body.lineSpacing"] == 1.5


def test_heading_spacing_is_not_body_spacing() -> None:
    """Одиночный интервал в заголовках и сносках — не требование к тексту."""
    text = (
        "Заголовки,подзаголовки и подстрочные сноски(состоящие из нескольких строк)"
        "печатаются черезодинарный интервал."
    )
    found = {item.path: item.value for item in extract_spacing(piece(text))}

    assert "body.lineSpacing" not in found
    assert found["headings.levels.1.lineSpacing"] == 1.0
    assert found["footnotes.lineSpacing"] == 1.0


def test_spacing_takes_its_subject_from_the_previous_sentence() -> None:
    """«Размер шрифта в таблице 12. Интервал — одинарный.» — это про таблицу."""
    found = {
        item.path: item.value
        for item in extract_spacing(
            piece("Интервал–одинарный.", previous="Размершрифтавтаблице12.")
        )
    }

    assert found == {"tables.lineSpacing": 1.0}


def test_paragraph_gaps_are_not_line_spacing() -> None:
    """«Интервалы между абзацами 0 пт» — это расстояния до и после абзаца."""
    found = {
        item.path: item.value
        for item in extract_spacing(piece("Интервалымеждуабзацами0пт."))
    }

    assert found == {"body.spaceBeforePt": 0.0, "body.spaceAfterPt": 0.0}


def test_margins_are_found_in_a_list_with_semicolons() -> None:
    """Поля перечисляют списком, и слова «поля» рядом с каждым нет."""
    text = "Поля –правое– 15мм; –левое– 30мм; –верхнее–20мм; –нижнее–20мм"
    found = {item.path: item.value for item in extract_margins(piece(text))}

    assert found == {
        "page.marginsMm.right": 15.0,
        "page.marginsMm.left": 30.0,
        "page.marginsMm.top": 20.0,
        "page.marginsMm.bottom": 20.0,
    }


def test_margin_units_are_converted() -> None:
    values = extract_margins(piece("Поля: левое — 3 см, правое — 1,5 см."))
    found = {item.path: item.value for item in values}

    assert found["page.marginsMm.left"] == 30
    assert found["page.marginsMm.right"] == 15


def test_margins_without_numbers_are_not_guessed() -> None:
    assert extract_margins(piece("Поля стандартные, левое поле больше правого.")) == []


def test_header_and_footer_distances() -> None:
    from word_formatter.core.profiles.extraction.from_text import extract_header_footer

    found = {
        item.path: item.value
        for item in extract_header_footer(piece("Колонтитулы –верхний–20мм;–нижний– 12,5мм"))
    }

    assert found == {"page.headerMm": 20.0, "page.footerMm": 12.5}


def test_page_number_in_the_bottom_centre_of_the_sheet() -> None:
    text = "Номер страницы проставляют,начиная со второй,в центре нижней части листа без точки."
    found = {item.path: item.value for item in extract_page_numbers(piece(text))}

    assert found["pageNumbers.position"] == "bottom-center"
    assert found["pageNumbers.skipFirstPage"] is True


# --- отрицания -----------------------------------------------------------------


def test_uppercase_is_not_broken_by_unrelated_negation() -> None:
    """«без абзацного отступа» относится к отступу, а не к прописным буквам."""
    values = extract_headings(
        piece(
            "Заголовки структурных элементов печатают прописными буквами, "
            "полужирным шрифтом, без абзацного отступа."
        )
    )
    found = {item.path: item.value for item in values}

    assert found["headings.structural.uppercase"] is True
    assert found["headings.structural.bold"] is True


def test_negation_after_the_word_is_understood() -> None:
    assert extract_hyphenation(piece("Автоматическая расстановка переносов не допускается."))[
        0
    ].value is False


def test_negation_before_the_word_is_understood() -> None:
    assert extract_headings(piece("Заголовки не выделяют полужирным шрифтом."))[0].value is False


def test_unrelated_emphasis_is_not_removed() -> None:
    assert extract_emphasis(piece("Допускается использовать курсив для терминов.")) == []


def test_heading_of_numbering_is_not_a_requirement() -> None:
    """Заголовок «Нумерация» без слов про страницы ничего не значит."""
    assert extract_page_numbers(piece("4. Нумерация")) == []


def test_captions_are_not_swapped() -> None:
    assert extract_captions(piece("Подпись рисунка размещают над рисунком."))[0].value == "above"


# --- оформление документа-образца ---------------------------------------------


def test_sample_formatting_is_measured(tmp_path: Path) -> None:
    """Методичка, оформленная по своим требованиям, читается из файла."""
    document = Document()
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(12)
    document.styles["Normal"].paragraph_format.line_spacing = 1.5
    document.styles["Normal"].paragraph_format.first_line_indent = Cm(1.25)
    document.styles["Normal"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    section = document.sections[0]
    section.left_margin = Cm(3)
    section.right_margin = Cm(1.5)
    document.add_paragraph("Требования к оформлению изложены ниже.")
    path = tmp_path / "образец.docx"
    document.save(path)

    draft = build_profile_draft(path)

    assert draft.sample_measured is True
    sample = {item.path: item for item in draft.values if item.origin == ORIGIN_SAMPLE}
    assert sample["body.sizePt"].value == 12
    assert sample["body.lineSpacing"].value == 1.5
    assert sample["page.marginsMm.left"].value == 30
    assert sample["body.firstLineIndentCm"].value == 1.25


def test_text_wins_over_sample_and_conflict_is_reported(tmp_path: Path) -> None:
    """Текст важнее оформления, а расхождение не скрывается."""
    document = Document()
    document.styles["Normal"].font.size = Pt(12)
    document.styles["Normal"].paragraph_format.line_spacing = 2.0
    for line in REQUIREMENTS.split("\n"):
        document.add_paragraph(line)
    path = tmp_path / "методичка.docx"
    document.save(path)

    draft = build_profile_draft(path)

    assert draft.value_for("body.sizePt").value == 14
    assert draft.value_for("body.lineSpacing").value == 1.5
    paths = {conflict.path for conflict in draft.conflicts}
    assert "body.sizePt" in paths
    assert "body.lineSpacing" in paths


# --- PDF ----------------------------------------------------------------------


@pytest.mark.skipif(not PDF_METHODOLOGY.is_file(), reason="нет файла требований в PDF")
def test_pdf_requirements_are_read() -> None:
    """PDF читается, а искажения букв «ё» и «й» разбору не мешают."""
    draft = build_profile_draft(PDF_METHODOLOGY)

    assert draft.value_for("body.font").value == "Times New Roman"
    assert draft.value_for("body.sizePt").value == 14
    assert draft.value_for("page.marginsMm.left").value == 30
    assert draft.value_for("captions.table.position").value == "above"
    assert draft.value_for("body.removeUnderline").value is True
    assert draft.value_for("pageNumbers.skipFirstPage").value is True


# --- ошибки -------------------------------------------------------------------


def test_missing_file_is_explained(tmp_path: Path) -> None:
    with pytest.raises(MethodologyError) as error:
        read_source(tmp_path / "нет-такого.pdf")

    assert "не найден" in error.value.user_message


def test_old_doc_format_is_explained(tmp_path: Path) -> None:
    path = tmp_path / "методичка.doc"
    path.write_bytes(b"\xd0\xcf\x11\xe0")

    with pytest.raises(UnsupportedFormatError) as error:
        read_source(path)

    assert "сохраните" in error.value.user_message.lower()


def test_methodology_without_requirements_is_explained(tmp_path: Path) -> None:
    path = tmp_path / "без-требований.txt"
    path.write_text("Здесь просто текст без единого требования к оформлению.", encoding="utf-8")

    with pytest.raises(MethodologyError) as error:
        build_profile_draft(path)

    assert "не нашлось" in error.value.user_message


# --- правка значений человеком -------------------------------------------------


def test_editable_rows_include_missing_with_empty_values(methodology: Path) -> None:
    """Ненайденные требования показываются пустыми — их можно вписать."""
    draft = build_profile_draft(methodology)
    rows = draft.editable_rows()

    missing = [row for row in rows if not row.found]
    assert missing, "в методичке нет всех требований, значит строки должны быть"
    assert all(row.value is None for row in missing)
    assert all(row.placeholder for row in missing)
    assert len(rows) == len(draft.values) + len(draft.missing)


def test_edit_replaces_found_value(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    changed = draft.apply_edits({"body.sizePt": 12.0})

    assert changed == 1
    assert draft.value_for("body.sizePt").value == 12.0
    assert draft.value_for("body.sizePt").origin == ORIGIN_EDIT
    profile = draft.to_profile(profile_id="draft", name="Черновик")
    assert profile["body"]["sizePt"] == 12.0


def test_edit_fills_missing_requirement(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    draft.apply_edits({"figures.alignment": "left"})
    profile = draft.to_profile(profile_id="draft", name="Черновик")

    assert profile["figures"]["alignment"] == "left"
    assert draft.value_for("figures.alignment").origin == ORIGIN_EDIT


def test_untouched_missing_value_stays_typical_and_is_listed(methodology: Path) -> None:
    draft = build_profile_draft(methodology)
    profile = draft.to_profile(profile_id="draft", name="Черновик")

    assert any("не нашлось" in item for item in profile["limitations"])


def test_edits_do_not_change_anything_when_empty(methodology: Path) -> None:
    draft = build_profile_draft(methodology)
    before = draft.value_for("body.sizePt").value

    assert draft.apply_edits({"body.sizePt": None}) == 0
    assert draft.value_for("body.sizePt").value == before


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("12", 12.0),
        ("12,5", 12.5),
        ("14 пт", 14.0),
        ("да", True),
        ("нет", False),
        ("по центру", "center"),
        ("под рисунком", "below"),
        ("внизу по центру", "bottom-center"),
        ("Times New Roman", "Times New Roman"),
        ("   ", None),
    ],
)
def test_typed_text_becomes_value(text: str, expected) -> None:
    # У числового поля и у поля-строки разбор текста разный, поэтому
    # проверяем на подходящем поле.
    path = "body.font" if isinstance(expected, str) and expected == "Times New Roman" else "body.sizePt"

    assert coerce_value(path, text) == expected


# --- сохранение профиля -------------------------------------------------------


def test_draft_becomes_valid_profile(methodology: Path, tmp_path: Path) -> None:
    draft = build_profile_draft(methodology)
    directory = tmp_path / "профили"

    path = save_draft_profile(draft, name="Моя методичка 2026", directory=directory)

    assert path.name == "moya-metodichka-2026.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    assert validate_profile(data) == []

    profile = load_profile_from_path(path)
    assert profile.name == "Моя методичка 2026"
    assert profile.body.size_pt == 14
    assert profile.page.margins.left == 30
    assert profile.captions.table.position == "above"


def test_edited_profile_saves_the_edited_value(methodology: Path, tmp_path: Path) -> None:
    draft = build_profile_draft(methodology)
    draft.apply_edits({"footnotes.fontSizePt": 9.0})

    path = save_draft_profile(draft, name="С правками", directory=tmp_path / "профили")

    assert load_profile_from_path(path).footnotes.font_size_pt == 9.0


def test_profile_id_is_transliterated() -> None:
    assert make_profile_id("Моя методичка 2026", taken=set()) == "moya-metodichka-2026"
    assert make_profile_id("Моя методичка 2026", taken={"moya-metodichka-2026"}) == (
        "moya-metodichka-2026-2"
    )


def test_origin_of_values_is_recorded(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    assert all(item.origin == ORIGIN_TEXT for item in draft.values)
    assert all(item.evidence for item in draft.values)
