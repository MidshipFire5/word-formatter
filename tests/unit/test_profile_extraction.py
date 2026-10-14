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
    ORIGIN_SAMPLE,
    ORIGIN_TEXT,
    build_profile_draft,
    make_profile_id,
    read_source,
    save_draft_profile,
)
from word_formatter.core.profiles.extraction.from_text import (
    extract_captions,
    extract_headings,
    extract_hyphenation,
    extract_margins,
    extract_page_numbers,
    extract_sizes,
    sentences,
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


@pytest.fixture()
def methodology(tmp_path: Path) -> Path:
    path = tmp_path / "методичка.txt"
    path.write_text(REQUIREMENTS, encoding="utf-8")
    return path


# --- разбор текста требований -------------------------------------------------


def test_sentences_split(methodology: Path) -> None:
    parts = sentences(REQUIREMENTS)

    assert "Поля: левое — 30 мм, правое — 15 мм, верхнее и нижнее — 20 мм." in parts
    assert all(part.strip() == part for part in parts)


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


def test_table_and_footnote_sizes_do_not_mix() -> None:
    assert extract_sizes("Шрифт в таблицах — 12 пт.")[0].path == "tables.fontSizePt"
    assert extract_sizes("Размер шрифта сносок — 10 пт.")[0].path == "footnotes.fontSizePt"
    assert extract_sizes("Размер шрифта — 14 пт.")[0].path == "body.sizePt"


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


def test_margin_units_are_converted() -> None:
    values = extract_margins("Поля: левое — 3 см, правое — 1,5 см.")

    assert values[0].value == 30
    assert values[1].value == 15


def test_margins_without_numbers_are_not_guessed() -> None:
    assert extract_margins("Поля стандартные, левое поле больше правого.") == []


# --- отрицания -----------------------------------------------------------------


def test_uppercase_is_not_broken_by_unrelated_negation() -> None:
    """«без абзацного отступа» относится к отступу, а не к прописным буквам."""
    values = extract_headings(
        "Заголовки структурных элементов печатают прописными буквами, "
        "полужирным шрифтом, без абзацного отступа."
    )
    found = {item.path: item.value for item in values}

    assert found["headings.structural.uppercase"] is True
    assert found["headings.structural.bold"] is True


def test_negation_after_the_word_is_understood() -> None:
    assert extract_hyphenation("Автоматическая расстановка переносов не допускается.")[0].value is False


def test_negation_before_the_word_is_understood() -> None:
    assert extract_headings("Заголовки не выделяют полужирным шрифтом.")[0].value is False


def test_unrelated_emphasis_is_not_removed() -> None:
    from word_formatter.core.profiles.extraction.from_text import extract_emphasis

    assert extract_emphasis("Допускается использовать курсив для терминов.") == []


def test_heading_of_numbering_is_not_a_requirement() -> None:
    """Заголовок «Нумерация» без слов про страницы ничего не значит."""
    assert extract_page_numbers("4. Нумерация") == []


def test_captions_are_not_swapped() -> None:
    assert extract_captions("Подпись рисунка размещают над рисунком.")[0].value == "above"


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
    conflict = next(item for item in draft.conflicts if item.path == "body.sizePt")
    assert conflict.from_text == 14
    assert conflict.from_sample == 12
    assert "взято из текста" in conflict.describe()


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


def test_pdf_source_is_read_by_pages() -> None:
    source = read_source(PDF_METHODOLOGY)

    assert source.kind == "pdf"
    assert source.page_count >= 1
    assert "Times New Roman" in source.text
    assert source.page_of("Times New Roman") == 1


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


def test_empty_file_is_explained(tmp_path: Path) -> None:
    path = tmp_path / "пусто.txt"
    path.write_text("   \n  ", encoding="utf-8")

    with pytest.raises(MethodologyError):
        read_source(path)


def test_methodology_without_requirements_is_explained(tmp_path: Path) -> None:
    path = tmp_path / "без-требований.txt"
    path.write_text("Здесь просто текст без единого требования к оформлению.", encoding="utf-8")

    with pytest.raises(MethodologyError) as error:
        build_profile_draft(path)

    assert "не нашлось" in error.value.user_message


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


def test_unfound_values_are_listed_as_limitations(methodology: Path) -> None:
    """Что не нашлось — то перечислено в ограничениях, а не выдано за требование."""
    draft = build_profile_draft(methodology)
    profile = draft.to_profile(profile_id="draft", name="Черновик")

    assert any("не нашлось" in item for item in profile["limitations"])
    assert draft.missing


def test_profile_id_is_transliterated() -> None:
    assert make_profile_id("Моя методичка 2026", taken=set()) == "moya-metodichka-2026"
    assert make_profile_id("Моя методичка 2026", taken={"moya-metodichka-2026"}) == (
        "moya-metodichka-2026-2"
    )


def test_saved_profile_uses_different_ids_for_same_name(methodology: Path, tmp_path: Path) -> None:
    draft = build_profile_draft(methodology)
    directory = tmp_path / "профили"

    first = save_draft_profile(draft, name="Методичка", directory=directory)
    second = save_draft_profile(draft, name="Методичка", directory=directory)

    assert first != second
    assert second.name == "metodichka-2.json"


def test_origin_of_values_is_recorded(methodology: Path) -> None:
    draft = build_profile_draft(methodology)

    assert all(item.origin == ORIGIN_TEXT for item in draft.values)
    assert all(item.evidence for item in draft.values)
