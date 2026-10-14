"""Инспекция настоящего документа со «сложными» элементами."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

from word_formatter.core.docx import DocumentInventory, Role, inspect_document, open_document
from word_formatter.core.profiles import load_profile_from_path


@pytest.fixture()
def profile(profile_file: Path):
    return load_profile_from_path(profile_file)


@pytest.fixture()
def inventory(rich_docx: Path, profile) -> DocumentInventory:
    return inspect_document(open_document(rich_docx), profile)


def item_with_text(inventory: DocumentInventory, text: str):
    for item in inventory.paragraphs:
        if item.text.strip() == text:
            return item
    raise AssertionError(f"в документе нет абзаца с текстом: {text!r}")


def test_element_counts(inventory: DocumentInventory) -> None:
    assert inventory.images >= 1
    assert inventory.hyperlinks == 1
    assert inventory.footnotes == 1
    assert inventory.sections == 2
    assert inventory.toc_fields == 1
    assert inventory.ole_objects == 1
    assert inventory.content_controls == 1
    assert inventory.page_breaks == 1
    assert inventory.has_toc is True


def test_table_is_described(inventory: DocumentInventory) -> None:
    assert len(inventory.tables) == 1
    table = inventory.tables[0]
    assert (table.rows, table.columns) == (3, 3)
    assert len(table.paragraphs) == 9
    assert all(item.role is Role.TABLE_CELL for item in table.paragraphs)
    assert "55000" in [item.text for item in table.paragraphs]


def test_body_sequence_matches_paragraph_order(inventory: DocumentInventory) -> None:
    kinds = [item.kind for item in inventory.body]
    assert kinds.count("table") == 1
    assert kinds.count("paragraph") == len(inventory.paragraphs)
    indices = [item.index for item in inventory.body if item.kind == "paragraph"]
    assert indices == list(range(len(inventory.paragraphs)))


def test_headings_are_recognised(inventory: DocumentInventory) -> None:
    assertive = item_with_text(inventory, "ВВЕДЕНИЕ")
    assert assertive.role is Role.STRUCTURAL_HEADING

    conclusion = item_with_text(inventory, "ЗАКЛЮЧЕНИЕ")
    assert conclusion.role is Role.STRUCTURAL_HEADING

    subsection = item_with_text(inventory, "1.1 Понятие и сущность процесса")
    assert subsection.role is Role.HEADING
    assert subsection.level == 2


def test_captions_are_recognised(inventory: DocumentInventory) -> None:
    table_caption = item_with_text(inventory, "Таблица 1 – Показатели деятельности")
    assert table_caption.role is Role.TABLE_CAPTION

    figure_caption = item_with_text(inventory, "Рисунок 1 – Схема процесса")
    assert figure_caption.role is Role.FIGURE_CAPTION


def test_figure_paragraph_is_recognised(inventory: DocumentInventory) -> None:
    assert inventory.role_count(Role.FIGURE) == 1


def test_body_text_is_recognised(inventory: DocumentInventory) -> None:
    paragraph = item_with_text(inventory, "Текст заключения во втором разделе документа.")
    assert paragraph.role is Role.BODY

    emphasized = item_with_text(inventory, "Обычный текст с ручным шрифтом, курсивом и подчёркиванием.")
    assert emphasized.role is Role.BODY
    assert emphasized.has_drawing is False


def test_embedded_object_is_reported_as_ambiguous(inventory: DocumentInventory) -> None:
    assert inventory.role_count(Role.AMBIGUOUS) >= 1
    reasons = " ".join(item.reason for item in inventory.ambiguous)
    assert "встроенный объект" in reasons


def test_section_break_paragraph_is_blank(inventory: DocumentInventory) -> None:
    assert inventory.role_count(Role.BLANK) >= 1


def test_heading_count_is_reported(inventory: DocumentInventory) -> None:
    # Два структурных заголовка (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ) и один заголовок подраздела
    assert inventory.role_count(Role.STRUCTURAL_HEADING) == 2
    assert inventory.role_count(Role.HEADING) == 1
    assert len(inventory.headings) == 3


def test_inspection_does_not_modify_source(rich_docx: Path, profile) -> None:
    before = hashlib.sha256(rich_docx.read_bytes()).hexdigest()
    inspect_document(open_document(rich_docx), profile)
    assert hashlib.sha256(rich_docx.read_bytes()).hexdigest() == before
