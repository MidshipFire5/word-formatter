"""Графический интерфейс: окно, выбор профиля, обработка, отчёт.

Окно проверяется в невидимом режиме (offscreen), поэтому тесты работают
и на машине без экрана, и в непрерывной интеграции.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

from fixtures.rich_docx import build_rich_docx
from word_formatter.core.profiles import ProfileCatalog
from word_formatter.ui import main_window as window_module
from word_formatter.ui.main_window import MainWindow

PROFILE_ID = "fu-krasnodar-2026"
PROCESSING_TIMEOUT = 60_000


@pytest.fixture(autouse=True)
def silent_dialogs(monkeypatch: pytest.MonkeyPatch) -> list[tuple[str, str]]:
    """Заменяет модальные окна Qt: в тестах они заблокировали бы прогон."""
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
def window(qtbot, builtin_profile_dir: Path, user_profile_dir: Path) -> MainWindow:
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    main_window = MainWindow(catalog=catalog)
    qtbot.addWidget(main_window)
    return main_window


def test_profiles_are_listed(window: MainWindow) -> None:
    assert window.profile_list.count() == 1
    assert "Финансовый университет" in window.profile_list.item(0).text()


def test_profile_is_selected_by_default(window: MainWindow) -> None:
    assert window.current_profile_id() == PROFILE_ID


def test_profile_description_is_shown(window: MainWindow) -> None:
    text = window.profile_description.toPlainText()

    assert "Версия: 1.0.0" in text
    assert "Источник требований" in text
    assert "Что профиль не делает" in text
    assert "основной текст Times New Roman 14 пт" in text


def test_selecting_profile_by_id(window: MainWindow) -> None:
    assert window.select_profile(PROFILE_ID) is True
    assert window.current_profile_id() == PROFILE_ID
    assert window.select_profile("нет-такого-профиля") is False


def test_apply_is_disabled_without_document(window: MainWindow) -> None:
    assert window.apply_button.isEnabled() is False


def test_document_selection_enables_apply(window: MainWindow, rich_docx: Path) -> None:
    assert window.set_source_file(rich_docx) is True
    assert rich_docx.name in window.file_label.text()
    assert window.apply_button.isEnabled() is True


def test_default_output_name_marks_result(window: MainWindow, rich_docx: Path) -> None:
    window.set_source_file(rich_docx)
    output = window.default_output_path()

    assert output is not None
    assert output.name.startswith("source")
    assert "_оформлено" in output.name
    assert output.parent == rich_docx.parent


def test_legacy_format_is_refused(window: MainWindow, tmp_path: Path, silent_dialogs) -> None:
    legacy = tmp_path / "старый.doc"
    legacy.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 32)

    assert window.set_source_file(legacy) is False
    assert any(kind == "warning" for kind, _title in silent_dialogs)
    assert window.apply_button.isEnabled() is False


def test_missing_file_is_refused(window: MainWindow, tmp_path: Path, silent_dialogs) -> None:
    assert window.set_source_file(tmp_path / "нет-такого.docx") is False
    assert silent_dialogs


def test_broken_profile_is_shown_without_crashing(
    qtbot, builtin_profile_dir: Path, user_profile_dir: Path
) -> None:
    (builtin_profile_dir / "broken.json").write_text("{ это не JSON", encoding="utf-8")
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    main_window = MainWindow(catalog=catalog)
    qtbot.addWidget(main_window)

    assert main_window.profile_list.count() == 2
    assert main_window.select_profile("broken") is True
    assert "не удалось прочитать" in main_window.profile_description.toPlainText().lower()
    assert main_window.apply_button.isEnabled() is False


def test_processing_end_to_end(window: MainWindow, rich_docx: Path, tmp_path: Path, qtbot) -> None:
    output = tmp_path / "результат.docx"
    assert window.set_source_file(rich_docx) is True

    with qtbot.waitSignal(window.processing_finished, timeout=PROCESSING_TIMEOUT):
        window.start_processing(output)

    assert output.exists()
    report_text = window.report_view.toPlainText()
    assert "Отчёт об оформлении документа" in report_text
    assert "Что сделано" in report_text
    assert "Проверка сохранности содержимого пройдена" in report_text
    assert window.save_report_button.isEnabled() is True
    assert window.apply_button.isEnabled() is True


def test_result_gets_profile_formatting(window: MainWindow, rich_docx: Path, tmp_path: Path, qtbot) -> None:
    from docx.shared import Mm, Pt

    output = tmp_path / "результат.docx"
    window.set_source_file(rich_docx)

    with qtbot.waitSignal(window.processing_finished, timeout=PROCESSING_TIMEOUT):
        window.start_processing(output)

    document = Document(str(output))
    paragraph = next(p for p in document.paragraphs if p.text.strip().startswith("Текст заключения"))
    assert paragraph.runs[0].font.size == Pt(14)
    assert document.sections[0].left_margin.twips == Mm(30).twips
    assert document.sections[0].page_width.twips == Mm(210).twips


def test_report_can_be_saved(window: MainWindow, rich_docx: Path, tmp_path: Path, qtbot) -> None:
    output = tmp_path / "результат.docx"
    window.set_source_file(rich_docx)
    with qtbot.waitSignal(window.processing_finished, timeout=PROCESSING_TIMEOUT):
        window.start_processing(output)

    report_path = tmp_path / "отчёт.txt"
    assert window.save_report(report_path) is True

    text = report_path.read_text(encoding="utf-8")
    assert "Отчёт об оформлении документа" in text
    assert "Как проверить результат" in text


def test_failure_is_reported(window: MainWindow, tmp_path: Path, qtbot, silent_dialogs) -> None:
    broken = tmp_path / "повреждённый.docx"
    broken.write_bytes(b"PK\x03\x04" + b"\x00" * 64)
    assert window.set_source_file(broken) is True

    with qtbot.waitSignal(window.processing_failed, timeout=PROCESSING_TIMEOUT) as signal:
        window.start_processing(tmp_path / "результат.docx")

    assert "повреждён" in signal.args[0] or "не удалось открыть" in signal.args[0].lower()
    assert any(kind == "critical" for kind, _title in silent_dialogs)
    assert window.apply_button.isEnabled() is True


def test_saving_over_source_is_refused(window: MainWindow, rich_docx: Path, qtbot, silent_dialogs) -> None:
    window.set_source_file(rich_docx)

    with qtbot.waitSignal(window.processing_failed, timeout=PROCESSING_TIMEOUT) as signal:
        window.start_processing(rich_docx)

    assert "поверх исходного" in signal.args[0]
