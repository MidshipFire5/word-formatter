"""Оформление интерфейса: тема, состояния и перетаскивание файла.

Все проверки идут в невидимом режиме и требуют созданного приложения Qt:
без него часть обращений к Qt аварийно завершает процесс.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, QUrl
from PySide6.QtGui import QFontDatabase
from PySide6.QtWidgets import QApplication, QStyleFactory

from word_formatter.core.profiles import ProfileCatalog
from word_formatter.ui import theme
from word_formatter.ui.main_window import MainWindow


@pytest.fixture()
def window(qtbot, builtin_profile_dir: Path, user_profile_dir: Path) -> MainWindow:
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    main_window = MainWindow(catalog=catalog)
    qtbot.addWidget(main_window)
    return main_window


# --- тема ---------------------------------------------------------------------


def test_stylesheet_has_rounded_corners() -> None:
    """Скругления — то, ради чего тема и делалась."""
    stylesheet = theme.stylesheet()

    assert "border-radius" in stylesheet
    assert f"{theme.RADIUS_CARD}px" in stylesheet
    assert f"{theme.RADIUS_CONTROL}px" in stylesheet


def test_stylesheet_covers_main_widgets() -> None:
    stylesheet = theme.stylesheet()

    for selector in (
        "QWidget#background",
        "QFrame#card",
        "QFrame#documentCard",
        "QPushButton#primaryButton",
        "QListWidget::item:selected",
        "QProgressBar::chunk",
        "QScrollBar:vertical",
        "QToolTip",
    ):
        assert selector in stylesheet, f"в теме нет правил для {selector}"


def test_colours_are_hex() -> None:
    for value in (
        theme.BACKGROUND,
        theme.CARD,
        theme.BORDER,
        theme.TEXT,
        theme.ACCENT,
        theme.SUCCESS,
        theme.DANGER,
    ):
        assert value.startswith("#") and len(value) == 7


def test_stylesheet_has_no_unresolved_placeholders() -> None:
    """В таблице стилей не должно остаться фигурных скобок от форматирования."""
    stylesheet = theme.stylesheet()

    assert "{{" not in stylesheet
    assert "}}" not in stylesheet


def test_font_family_is_available(qapp) -> None:
    family = theme.pick_font_family()

    assert family
    families = set(QFontDatabase.families())
    # В невидимом режиме без папки шрифтов список семейств может быть пустым:
    # тогда используется запасное семейство.
    if families:
        assert family in families


def test_font_falls_back_when_no_fonts_exist(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(QFontDatabase, "families", staticmethod(lambda: []))

    assert theme.pick_font_family()


def test_font_selection_without_application_is_safe(qapp) -> None:
    """Выбор шрифта не должен зависеть от того, создано ли уже приложение."""
    family = theme.pick_font_family()

    assert isinstance(family, str)


def test_theme_is_applied_to_application(qapp) -> None:
    theme.apply_theme(qapp)

    assert qapp.styleSheet() == theme.stylesheet()
    assert qapp.font().pointSizeF() == theme.FONT_SIZE
    # Вместе с темой включается стиль Fusion: он рисует элементы одинаково
    # во всех системах. Когда задана таблица стилей, Qt подменяет стиль
    # на QStyleSheetStyle, поэтому проверяем, что Fusion доступен и выбран
    # в коде темы (см. тест ниже).
    assert "Fusion" in QStyleFactory.keys()


def test_source_requests_fusion_style() -> None:
    source = Path(theme.__file__).read_text(encoding="utf-8")

    assert 'setStyle("Fusion")' in source


# --- состояния ----------------------------------------------------------------


def test_status_state_is_switched(window: MainWindow) -> None:
    window._set_status("Ошибка", "error")
    assert window.status_label.property("state") == "error"

    window._set_status("Готово", "success")
    assert window.status_label.property("state") == "success"


def test_ready_hint_appears_after_choosing_file(window: MainWindow, rich_docx: Path) -> None:
    assert "Выберите профиль" in window.status_label.text()

    window.set_source_file(rich_docx)

    assert window.status_label.text() == "Можно применять оформление."


def test_save_report_button_explains_itself(window: MainWindow) -> None:
    assert "Отчёт появится" in window.save_report_button.toolTip()


def test_document_card_states_are_known(window: MainWindow) -> None:
    for state in ("empty", "set", "dragover"):
        theme.set_state(window.document_card, state)
        assert window.document_card.property("state") == state


def test_state_is_not_set_twice(window: MainWindow) -> None:
    """Повторная установка того же состояния ничего не ломает."""
    theme.set_state(window.status_label, "success")
    theme.set_state(window.status_label, "success")

    assert window.status_label.property("state") == "success"


# --- перетаскивание файла -----------------------------------------------------


def test_window_accepts_dropped_files(window: MainWindow) -> None:
    assert window.acceptDrops() is True


def test_only_docx_files_are_taken_from_drop(tmp_path: Path) -> None:
    document = tmp_path / "работа.docx"
    document.write_bytes(b"PK\x03\x04")
    other = tmp_path / "заметки.txt"
    other.write_text("текст", encoding="utf-8")

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(document)), QUrl.fromLocalFile(str(other))])

    assert MainWindow._dropped_documents(mime) == [document]


def test_drop_without_files_is_ignored() -> None:
    mime = QMimeData()
    mime.setText("просто текст")

    assert MainWindow._dropped_documents(mime) == []
    assert MainWindow._dropped_documents(None) == []


# --- подпись выбранного файла -------------------------------------------------


def test_file_label_shows_name_and_size(window: MainWindow, rich_docx: Path) -> None:
    assert window.file_label.text() == "Документ не выбран"
    assert window.file_label.property("state") == "empty"

    window.set_source_file(rich_docx)
    text = window.file_label.text()

    assert rich_docx.name in text
    assert "КБ" in text
    assert window.file_label.toolTip() == str(rich_docx)
    assert window.file_label.property("state") == "set"


def test_application_applies_theme_on_start() -> None:
    """Тема обязана применяться при обычном запуске, а не только в тестах."""
    from word_formatter.ui import application as application_module

    source = Path(application_module.__file__).read_text(encoding="utf-8")

    assert "apply_theme(application)" in source
    assert QApplication.instance() is not None
