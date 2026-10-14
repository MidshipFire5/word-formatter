"""Значения профиля: сверка с таблицей 1 методических указаний."""

from __future__ import annotations

import copy
import json
from pathlib import Path

from word_formatter.core.profiles import load_profile_from_path


def test_profile_matches_methodology_table(profile_file: Path) -> None:
    """Каждый пункт таблицы 1 методички должен быть в профиле явно."""
    profile = load_profile_from_path(profile_file)

    # Формат листа бумаги и поля
    assert profile.page.format == "A4"
    assert profile.page.orientation == "portrait"
    assert profile.page.margins.left == 30
    assert profile.page.margins.right == 15
    assert profile.page.margins.top == 20
    assert profile.page.margins.bottom == 20
    assert profile.page.header_mm == 20
    assert profile.page.footer_mm == 12.5

    # Гарнитура, размер и параметры основного текста
    assert profile.body.font == "Times New Roman"
    assert profile.body.size_pt == 14
    assert profile.body.alignment == "justify"
    assert profile.body.first_line_indent_cm == 1.25
    assert profile.body.line_spacing == 1.5
    assert profile.body.space_before_pt == 0
    assert profile.body.space_after_pt == 0
    assert profile.body.remove_italic is True
    assert profile.body.remove_underline is True

    # Таблицы и сноски
    assert profile.tables.font == "Times New Roman"
    assert profile.tables.font_size_pt == 12
    assert profile.tables.line_spacing == 1.0
    assert profile.footnotes.font_size_pt == 10

    # Нумерация страниц: внизу по центру, 12 пунктов, на первой странице номера нет
    assert profile.page_numbers.enabled is True
    assert profile.page_numbers.position == "bottom-center"
    assert profile.page_numbers.font_size_pt == 12
    assert profile.page_numbers.font == "Times New Roman"
    assert profile.page_numbers.skip_first_page is True

    # Структурные заголовки: прописными, по центру, без отступа, полужирные
    structural = profile.headings.structural
    assert structural.uppercase is True
    assert structural.alignment == "center"
    assert structural.first_line_indent_cm == 0
    assert structural.size_pt == 14
    assert structural.bold is True
    assert structural.line_spacing == 1.0
    assert "ВВЕДЕНИЕ" in structural.titles
    assert "ЗАКЛЮЧЕНИЕ" in structural.titles
    assert "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" in structural.titles

    # Разделы и подразделы: с абзацного отступа, по ширине, 14, полужирные
    for level in (1, 2, 3):
        style = profile.headings.levels[level]
        assert style.alignment == "justify"
        assert style.first_line_indent_cm == 1.25
        assert style.size_pt == 14
        assert style.bold is True
        assert style.line_spacing == 1.0

    # Подписи: рисунок по центру, название таблицы без абзацного отступа
    assert profile.captions.figure.alignment == "center"
    assert profile.captions.figure.first_line_indent_cm == 0
    assert profile.captions.figure.no_hyphenation is True
    assert profile.captions.table.first_line_indent_cm == 0
    assert profile.captions.table.no_hyphenation is True

    # Переносы и свободные строки
    assert profile.hyphenation.enabled is True
    assert profile.hyphenation.exclude_headings is True
    assert profile.blank_lines.insert is True

    # Списки: маркеры не заменяются, оформляются только абзацы
    assert profile.lists.apply_paragraph_formatting is True
    assert profile.lists.replace_markers is False


def test_profile_identifies_itself(profile_file: Path) -> None:
    profile = load_profile_from_path(profile_file)
    assert profile.id == "fu-krasnodar-2026"
    assert profile.version == "1.0.0"
    assert profile.demo is False
    assert "методические" in profile.source_title.lower()
    assert "2026" in profile.source_title
    assert profile.limitations, "ограничения профиля должны быть перечислены"
    assert str(profile_file) == profile.path


def test_page_size_follows_orientation(profile_file: Path, profile_dict: dict, tmp_path: Path) -> None:
    landscape = copy.deepcopy(profile_dict)
    landscape["page"]["orientation"] = "landscape"
    path = tmp_path / "landscape.json"
    path.write_text(json.dumps(landscape, ensure_ascii=False), encoding="utf-8")

    portrait_profile = load_profile_from_path(profile_file)
    landscape_profile = load_profile_from_path(path)

    assert portrait_profile.page.size_mm == (210.0, 297.0)
    assert landscape_profile.page.size_mm == (297.0, 210.0)


def test_unknown_level_falls_back_to_last_described(profile_file: Path) -> None:
    profile = load_profile_from_path(profile_file)
    assert profile.headings.level_style(2) is profile.headings.levels[2]
    assert profile.headings.level_style(7) is profile.headings.levels[profile.headings.max_level]


def test_profile_is_immutable(profile_file: Path) -> None:
    """Модель только для чтения: правила не могут случайно испортить профиль."""
    profile = load_profile_from_path(profile_file)
    try:
        profile.body.size_pt = 12  # type: ignore[misc]
    except AttributeError:
        return
    raise AssertionError("модель профиля должна быть неизменяемой")
