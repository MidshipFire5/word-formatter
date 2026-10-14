"""Создание профиля из методички в интерфейсе."""

from __future__ import annotations

from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, Qt, QUrl
from PySide6.QtWidgets import QDialog

from word_formatter.core.profiles import ProfileCatalog, load_profile_from_path
from word_formatter.core.profiles.extraction import build_profile_draft
from word_formatter.core.profiles.extraction.draft import value_hint
from word_formatter.ui import main_window as window_module
from word_formatter.ui.main_window import MainWindow
from word_formatter.ui.methodology import (
    DropFileButton,
    MethodologyDraftDialog,
    human_value,
)

REQUIREMENTS = """МЕТОДИЧЕСКИЕ УКАЗАНИЯ

Текст работы набирается шрифтом Times New Roman, размер 12 пт, межстрочный интервал 1,5.
Абзацный отступ — 1,25 см. Поля: левое — 30 мм, правое — 10 мм, верхнее и нижнее — 20 мм.
Выравнивание основного текста — по ширине.
Подпись рисунка размещают под рисунком, название таблицы — над таблицей.
Нумерацию страниц проставляют арабскими цифрами внизу по центру, начиная со второй страницы.
"""


@pytest.fixture(autouse=True)
def silent_message_boxes(monkeypatch: pytest.MonkeyPatch) -> list[tuple[str, str]]:
    """Модальные окна в тестах заблокировали бы прогон."""
    shown: list[tuple[str, str]] = []

    def make(kind: str):
        def show(*args, **kwargs):
            title = args[1] if len(args) > 1 else ""
            shown.append((kind, str(title)))
            return None

        return show

    for name in ("information", "warning", "critical"):
        monkeypatch.setattr(window_module.QMessageBox, name, make(name))
    return shown


@pytest.fixture()
def methodology(tmp_path: Path) -> Path:
    path = tmp_path / "методичка.txt"
    path.write_text(REQUIREMENTS, encoding="utf-8")
    return path


@pytest.fixture()
def window(qtbot, builtin_profile_dir: Path, user_profile_dir: Path) -> MainWindow:
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    main_window = MainWindow(catalog=catalog)
    qtbot.addWidget(main_window)
    return main_window


# --- кнопка с перетаскиванием --------------------------------------------------


def test_dropped_file_is_taken_from_urls(tmp_path: Path) -> None:
    document = tmp_path / "методичка.docx"
    document.write_bytes(b"PK\x03\x04")
    other = tmp_path / "картинка.png"
    other.write_bytes(b"\x89PNG")

    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(other)), QUrl.fromLocalFile(str(document))])

    assert Path(DropFileButton.dropped_file(mime)) == document


def test_drop_without_files_is_ignored() -> None:
    mime = QMimeData()
    mime.setText("просто текст")

    assert DropFileButton.dropped_file(mime) is None
    assert DropFileButton.dropped_file(None) is None


def test_button_accepts_drops(window: MainWindow) -> None:
    assert window.create_profile_button.acceptDrops() is True


# --- окно проверки ------------------------------------------------------------


def test_dialog_lists_found_values(qtbot, methodology: Path) -> None:
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(
        draft, source_name="методичка", suggested_name="Моя методичка"
    )
    qtbot.addWidget(dialog)

    assert dialog.profile_name() == "Моя методичка"
    assert len(draft.values) > 10
    assert "Найдено требований" in dialog._summary_text()


def test_dialog_shows_values_in_plain_language(qtbot, methodology: Path) -> None:
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(
        draft, source_name="методичка", suggested_name="Моя методичка"
    )
    qtbot.addWidget(dialog)

    assert human_value(True) == "да"
    assert human_value(False) == "нет"
    assert human_value("bottom-center") == "внизу по центру"
    assert human_value(12) == "12"

    assert "Методичка" in dialog.windowTitle() or dialog.windowTitle()


def test_dialog_shows_origin_of_each_value(qtbot, methodology: Path) -> None:
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(
        draft, source_name="методичка", suggested_name="Моя методичка"
    )
    qtbot.addWidget(dialog)

    row = dialog._rows[0]
    text = dialog._source_text(row)

    assert "из текста методички" in text
    assert row.evidence[:20] in text


def test_dialog_values_can_be_edited(qtbot, methodology: Path) -> None:
    """Столбец «Значение» редактируется: правку можно внести сразу."""
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(
        draft, source_name="методичка", suggested_name="Моя методичка"
    )
    qtbot.addWidget(dialog)

    assert dialog.edited_values() == {}

    index = next(
        position for position, row in enumerate(dialog._rows) if row.path == "body.sizePt"
    )
    # В методичке указано 12 пт — исправляем на другое значение.
    dialog._fields[index].setText("13")

    assert dialog.edited_values() == {"body.sizePt": 13.0}


def test_dialog_shows_empty_rows_for_missing_requirements(qtbot, methodology: Path) -> None:
    """Чего приложение не нашло — то пустое, и его можно вписать."""
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(
        draft, source_name="методичка", suggested_name="Моя методичка"
    )
    qtbot.addWidget(dialog)

    empty_rows = [
        (position, row)
        for position, row in enumerate(dialog._rows)
        if not row.found
    ]
    assert empty_rows, "в методичке требований меньше, чем умеет профиль"

    position, row = empty_rows[0]
    field = dialog._fields[position]
    assert field.text() == ""
    assert field.placeholderText(), "в пустой ячейке должна быть серая подсказка"
    assert "не найдено" in dialog._table.item(position, 2).text()

    field.setText("1,5")
    assert dialog.edited_values().get(row.path) == 1.5


def test_every_value_cell_has_a_hint(qtbot, methodology: Path) -> None:
    """В каждой ячейке значения видно, что можно вписать."""
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(
        draft, source_name="методичка", suggested_name="Моя методичка"
    )
    qtbot.addWidget(dialog)

    for row, field in zip(dialog._rows, dialog._fields):
        assert field.placeholderText(), f"нет подсказки для {row.path}"


@pytest.mark.parametrize(
    ("path", "expected"),
    [
        ("pageNumbers.enabled", "да или нет"),
        ("headings.structural.uppercase", "да или нет"),
        ("body.sizePt", "число в пунктах"),
        ("body.firstLineIndentCm", "число в сантиметрах"),
        ("page.marginsMm.left", "число в миллиметрах"),
        ("headings.maxLevel", "целое число"),
        ("body.alignment", "по ширине"),
        ("captions.figure.position", "над элементом"),
        ("pageNumbers.position", "внизу по центру"),
        ("body.font", "название шрифта"),
        ("body.lineSpacing", "1,5 — полуторный"),
    ],
)
def test_hint_explains_what_to_type(path: str, expected: str) -> None:
    assert expected in value_hint(path)


def test_unclear_value_is_explained_and_blocks_saving(qtbot, methodology: Path) -> None:
    """Непонятное значение не теряется молча — окно объясняет, что вписать."""
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(
        draft, source_name="методичка", suggested_name="Моя методичка"
    )
    qtbot.addWidget(dialog)

    index = next(
        position for position, row in enumerate(dialog._rows) if row.path == "body.sizePt"
    )
    dialog._fields[index].setText("четырнадцать")

    assert dialog.unclear_rows()[0][0].path == "body.sizePt"

    dialog._on_save()

    assert dialog.result() != QDialog.Accepted
    assert "четырнадцать" in dialog.message_label.text()
    assert "число в пунктах" in dialog.message_label.text()
    assert dialog.edited_values() == {}


def test_clear_value_passes_and_is_reported(qtbot, methodology: Path) -> None:
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(
        draft, source_name="методичка", suggested_name="Моя методичка"
    )
    qtbot.addWidget(dialog)

    index = next(
        position for position, row in enumerate(dialog._rows) if row.path == "body.sizePt"
    )
    dialog._fields[index].setText("13")

    assert dialog.unclear_rows() == []
    dialog._on_save()

    assert dialog.result() == QDialog.Accepted
    assert "Исправлено вами: 1" in dialog.message_label.text()


def test_dialog_ignores_unchanged_rows(qtbot, methodology: Path) -> None:
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(
        draft, source_name="методичка", suggested_name="Моя методичка"
    )
    qtbot.addWidget(dialog)

    assert dialog.edited_values() == {}


def test_dialog_requires_a_name(qtbot, methodology: Path) -> None:
    draft = build_profile_draft(methodology)
    dialog = MethodologyDraftDialog(draft, source_name="методичка", suggested_name="")
    qtbot.addWidget(dialog)

    dialog._on_save()

    assert dialog.result() != QDialog.Accepted


# --- создание профиля целиком --------------------------------------------------


def test_edited_value_lands_in_the_saved_profile(
    qtbot, window: MainWindow, methodology: Path, user_profile_dir: Path, monkeypatch
) -> None:
    """Правка в таблице сохраняется в профиль, а не теряется."""
    index_holder: dict[str, int] = {}

    class EditingDialog(MethodologyDraftDialog):
        def __init__(self, *args, **kwargs) -> None:
            super().__init__(*args, **kwargs)
            position = next(
                place for place, row in enumerate(self._rows) if row.path == "body.sizePt"
            )
            self._fields[position].setText("13")

        def exec(self) -> int:
            return QDialog.Accepted

    monkeypatch.setattr(window_module_methodology(), "MethodologyDraftDialog", EditingDialog)

    profile_id = window.create_profile_from_methodology(methodology)

    assert profile_id is not None
    profile = load_profile_from_path(user_profile_dir / f"{profile_id}.json")
    assert profile.body.size_pt == 13


def window_module_methodology():
    from word_formatter.ui import methodology as module

    return module


def test_profile_is_created_from_methodology(
    qtbot, window: MainWindow, methodology: Path, user_profile_dir: Path, monkeypatch
) -> None:
    monkeypatch.setattr(MethodologyDraftDialog, "exec", lambda self: QDialog.Accepted)

    profile_id = window.create_profile_from_methodology(methodology)

    assert profile_id == "moya-metodichka-metodichka"
    saved = user_profile_dir / f"{profile_id}.json"
    assert saved.is_file()

    profile = load_profile_from_path(saved)
    assert profile.body.size_pt == 12
    assert profile.page.margins.right == 10
    assert profile.captions.table.position == "above"

    # профиль появился в списке и сразу выбран
    listed = [
        window.profile_list.item(row).data(Qt.UserRole)
        for row in range(window.profile_list.count())
    ]
    assert window.current_profile_id() == profile_id
    assert profile_id in listed


def test_created_profile_is_ready_to_use(
    qtbot, window: MainWindow, methodology: Path, rich_docx: Path, monkeypatch
) -> None:
    """После создания профиля кнопка «Применить оформление» доступна."""
    monkeypatch.setattr(MethodologyDraftDialog, "exec", lambda self: QDialog.Accepted)

    window.create_profile_from_methodology(methodology)
    window.set_source_file(rich_docx)

    assert window.apply_button.isEnabled() is True


def test_cancelled_dialog_saves_nothing(
    qtbot, window: MainWindow, methodology: Path, user_profile_dir: Path, monkeypatch
) -> None:
    monkeypatch.setattr(MethodologyDraftDialog, "exec", lambda self: QDialog.Rejected)

    result = window.create_profile_from_methodology(methodology)

    assert result is None
    assert list(user_profile_dir.glob("*.json")) == []


def test_unreadable_methodology_is_explained(
    qtbot, window: MainWindow, tmp_path: Path, silent_message_boxes
) -> None:
    empty = tmp_path / "пустая методичка.txt"
    empty.write_text("   ", encoding="utf-8")

    result = window.create_profile_from_methodology(empty)

    assert result is None
    assert window.status_label.property("state") == "error"
    assert any("методичку" in title.lower() for _kind, title in silent_message_boxes)


def test_unsupported_format_is_explained(
    qtbot, window: MainWindow, tmp_path: Path, silent_message_boxes
) -> None:
    old = tmp_path / "методичка.doc"
    old.write_bytes(b"\xd0\xcf\x11\xe0")

    assert window.create_profile_from_methodology(old) is None
    assert window.status_label.property("state") == "error"
