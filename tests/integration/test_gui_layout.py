"""Размеры окна и списка профилей.

Проверки появились после замечания: при уменьшении окна список профилей
сжимался до одной строки, и выбирать профиль становилось неудобно.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from PySide6.QtWidgets import QApplication

from word_formatter.core.profiles import ProfileCatalog
from word_formatter.ui import theme
from word_formatter.ui.main_window import (
    PROFILE_ROWS_VISIBLE,
    WINDOW_MINIMUM,
    MainWindow,
)

# Сколько профилей должно быть видно в списке при самом маленьком окне.
REQUIRED_VISIBLE_ROWS = 5


@pytest.fixture()
def window(qtbot, builtin_profile_dir: Path, user_profile_dir: Path) -> MainWindow:
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    main_window = MainWindow(catalog=catalog)
    qtbot.addWidget(main_window)
    return main_window


def test_visible_rows_constant_matches_requirement() -> None:
    assert PROFILE_ROWS_VISIBLE == REQUIRED_VISIBLE_ROWS


def test_profile_list_has_room_for_five_rows(window: MainWindow) -> None:
    """Минимальная высота списка рассчитана на пять названий."""
    row_height = window.profile_list.sizeHintForRow(0)

    assert row_height > 0
    assert window.profile_list.minimumHeight() >= REQUIRED_VISIBLE_ROWS * row_height


def test_profile_list_keeps_five_rows_in_a_small_window(
    qtbot, window: MainWindow
) -> None:
    """Даже в самом маленьком окне видно пять названий, а не одно."""
    window.resize(*WINDOW_MINIMUM)
    window.show()
    qtbot.waitExposed(window, timeout=2000)
    QApplication.processEvents()

    row_height = window.profile_list.sizeHintForRow(0)
    visible_height = window.profile_list.viewport().height()

    assert visible_height >= REQUIRED_VISIBLE_ROWS * row_height, (
        f"в списке помещается {visible_height // row_height} названий "
        f"при требуемых {REQUIRED_VISIBLE_ROWS}"
    )


def test_small_window_scrolls_instead_of_squeezing(qtbot, window: MainWindow) -> None:
    """Маленькое окно прокручивается: содержимое сохраняет свои размеры."""
    window.resize(*WINDOW_MINIMUM)
    window.show()
    qtbot.waitExposed(window, timeout=2000)
    QApplication.processEvents()

    scroll = window.centralWidget()
    assert scroll.objectName() == "contentScroll"
    assert scroll.widgetResizable() is True

    report_card = window.report_view.parentWidget()
    assert report_card.minimumHeight() > 0


def test_list_height_is_recomputed_after_new_profile(
    qtbot, window: MainWindow, user_profile_dir: Path, builtin_profile_dir: Path
) -> None:
    """Высота списка не растёт с числом профилей: она рассчитана на пять."""
    before = window.profile_list.minimumHeight()

    user_profile_dir.mkdir(parents=True, exist_ok=True)
    source = next(builtin_profile_dir.glob("*.json"))
    second = source.read_text(encoding="utf-8").replace(
        '"fu-krasnodar-2026"', '"second-profile"'
    )
    (user_profile_dir / "second-profile.json").write_text(second, encoding="utf-8")
    window.refresh_profiles()

    assert window.profile_list.count() == 2
    assert window.profile_list.minimumHeight() == before


def test_scroll_area_is_styled() -> None:
    """Прокрутка содержимого не должна менять фон окна."""
    assert "QScrollArea#contentScroll" in theme.stylesheet()
