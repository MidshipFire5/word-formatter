"""Сообщение о проблеме: сбор отчёта и отправка по почте."""

from __future__ import annotations

import os
from pathlib import Path
from urllib.parse import unquote

import pytest
from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QDialog

from word_formatter import APP_VERSION
from word_formatter.core.profiles import ProfileCatalog
from word_formatter.ui import problem_report
from word_formatter.ui.main_window import MainWindow
from word_formatter.ui.problem_report import (
    MAILTO_BODY_LIMIT,
    SUPPORT_EMAIL,
    ProblemContext,
    ProblemReportDialog,
    build_message,
    collect_report,
    log_tail,
    mailto_url,
    mask_private,
)


@pytest.fixture()
def window(qtbot, builtin_profile_dir: Path, user_profile_dir: Path) -> MainWindow:
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    main_window = MainWindow(catalog=catalog)
    qtbot.addWidget(main_window)
    return main_window


# --- сбор отчёта ---------------------------------------------------------------


def test_home_directory_and_user_name_are_masked() -> None:
    """Личные данные в отчёт не попадают."""
    home = Path.home()
    name = os.environ.get("USERNAME") or os.environ.get("USER") or ""

    masked = mask_private(f"Файл: {home}\\работа.docx")

    assert str(home) not in masked
    assert "<домашняя папка>" in masked
    if name:
        assert name not in mask_private(f"Пользователь {name} запустил программу")


def test_report_contains_versions_and_context() -> None:
    context = ProblemContext(
        action="Обработка документа…",
        error="Не удалось сохранить результат",
        details="PermissionError",
        document="Doc1.docx",
        document_size="38 КБ",
        profile="ГОСТ 7.32-2017 (gost-7-32-2017)",
        result="Doc1_оформлено.docx, проверка сохранности пройдена",
    )

    report = collect_report(context, log_lines=0)

    assert APP_VERSION in report
    assert "Система:" in report and "Python:" in report and "Qt:" in report
    assert "Не удалось сохранить результат" in report
    assert "PermissionError" in report
    assert "Doc1.docx" in report
    assert "gost-7-32-2017" in report


def test_report_without_errors_says_so() -> None:
    report = collect_report(ProblemContext(), log_lines=0)

    assert "Последнее действие: не указано" in report
    assert "Ошибка: не указано" in report


def test_user_text_goes_first() -> None:
    """Слова человека — самое важное, они идут перед техникой."""
    message = build_message("Программа испортила отступы в таблице", ProblemContext(), log_lines=0)

    assert message.startswith("Программа испортила отступы в таблице")
    assert "Что приложила программа" in message


def test_empty_description_is_marked() -> None:
    message = build_message("   ", ProblemContext(), log_lines=0)

    assert message.startswith("Описание не заполнено.")


def test_log_tail_reads_last_lines(tmp_path: Path, monkeypatch) -> None:
    log_dir = tmp_path / "logs"
    log_dir.mkdir()
    (log_dir / problem_report.LOG_FILE_NAME).write_text(
        "\n".join(f"строка {number}" for number in range(1, 51)), encoding="utf-8"
    )
    monkeypatch.setattr(problem_report, "user_log_dir", lambda: log_dir)

    tail = log_tail(5)

    assert "строка 50" in tail
    assert "строка 46" in tail
    assert "строка 45" not in tail


def test_log_tail_is_empty_without_file(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(problem_report, "user_log_dir", lambda: tmp_path / "нет-папки")

    assert log_tail(10) == ""


# --- письмо --------------------------------------------------------------------


def test_mailto_points_to_support_address() -> None:
    url = mailto_url("текст отчёта")

    assert url.startswith(f"mailto:{SUPPORT_EMAIL}?")
    assert "subject=" in url and "body=" in url
    assert "%D1%82%D0%B5%D0%BA%D1%81%D1%82" in url  # «текст» в кодировке URL


def test_mailto_body_is_limited() -> None:
    """Почтовые программы не принимают бесконечно длинные ссылки."""
    url = mailto_url("я" * (MAILTO_BODY_LIMIT * 2))
    body = unquote(url.split("body=", 1)[1])

    assert len(body) == MAILTO_BODY_LIMIT


def test_mailto_keeps_user_text() -> None:
    url = mailto_url("Не сохраняется файл")

    assert "Не сохраняется файл" in unquote(url)


# --- окно ----------------------------------------------------------------------


def test_dialog_shows_report_and_requires_description(qtbot) -> None:
    dialog = ProblemReportDialog(ProblemContext(error="Сбой"))
    qtbot.addWidget(dialog)

    assert SUPPORT_EMAIL in dialog.status_label.text()
    assert "Сбой" in dialog.report_view.toPlainText()
    assert dialog.user_text() == ""

    assert dialog.send_by_mail() is False
    assert "Опишите" in dialog.status_label.text()


def test_dialog_sends_by_mail(qtbot, monkeypatch) -> None:
    opened: list[str] = []
    monkeypatch.setattr(
        QDesktopServices,
        "openUrl",
        staticmethod(lambda url: opened.append(url.toString()) or True),
    )
    dialog = ProblemReportDialog(ProblemContext(error="Сбой"))
    qtbot.addWidget(dialog)
    dialog.description_edit.setPlainText("Рисунок уехал вправо")

    assert dialog.send_by_mail() is True

    assert opened and opened[0].startswith(f"mailto:{SUPPORT_EMAIL}")
    body = unquote(opened[0].split("body=", 1)[1])
    assert "Рисунок уехал вправо" in body
    assert SUPPORT_EMAIL in dialog.status_label.text()


def test_dialog_reports_when_mail_program_missing(qtbot, monkeypatch) -> None:
    monkeypatch.setattr(QDesktopServices, "openUrl", staticmethod(lambda url: False))
    dialog = ProblemReportDialog(ProblemContext())
    qtbot.addWidget(dialog)
    dialog.description_edit.setPlainText("Ничего не работает")

    assert dialog.send_by_mail() is False
    assert "не удалось" in dialog.status_label.text().lower()


def test_dialog_saves_report_to_file(qtbot, tmp_path: Path, monkeypatch) -> None:
    target = tmp_path / "отчёт.txt"
    monkeypatch.setattr(
        problem_report.QFileDialog,
        "getSaveFileName",
        staticmethod(lambda *args, **kwargs: (str(target), "")),
    )
    dialog = ProblemReportDialog(ProblemContext(error="Сбой"))
    qtbot.addWidget(dialog)
    dialog.description_edit.setPlainText("Текст пользователя")

    assert dialog.save_report() is True

    saved = target.read_text(encoding="utf-8")
    assert "Текст пользователя" in saved
    assert "Сбой" in saved
    assert str(target) in dialog.status_label.text()


def test_dialog_copies_report_to_clipboard(qtbot) -> None:
    dialog = ProblemReportDialog(ProblemContext(error="Сбой"))
    qtbot.addWidget(dialog)
    dialog.description_edit.setPlainText("Текст для буфера")

    dialog.copy_report()

    from PySide6.QtGui import QGuiApplication

    assert "Текст для буфера" in QGuiApplication.clipboard().text()
    assert "скопирован" in dialog.status_label.text().lower()


def test_closed_dialog_does_not_send_anything(qtbot, monkeypatch) -> None:
    opened: list[str] = []
    monkeypatch.setattr(
        QDesktopServices, "openUrl", staticmethod(lambda url: opened.append(url.toString()) or True)
    )
    dialog = ProblemReportDialog(ProblemContext())
    qtbot.addWidget(dialog)
    dialog.reject()

    assert dialog.result() == QDialog.Rejected
    assert opened == []


# --- кнопка в окне -------------------------------------------------------------


def test_window_has_problem_button(window: MainWindow) -> None:
    assert window.problem_button.text() == "Сообщить об ошибке"
    assert window.problem_button.toolTip()


def test_window_opens_report_dialog_with_context(
    qtbot, window: MainWindow, rich_docx: Path, monkeypatch
) -> None:
    shown: list[ProblemContext] = []

    class FakeDialog:
        def __init__(self, context, **kwargs) -> None:
            shown.append(context)

        def exec(self) -> int:
            return QDialog.Rejected

    monkeypatch.setattr(problem_report, "ProblemReportDialog", FakeDialog)
    window.set_source_file(rich_docx)

    window.report_problem()

    assert len(shown) == 1
    context = shown[0]
    assert context.document == rich_docx.name
    assert "КБ" in context.document_size
    assert context.profile  # профиль выбран по умолчанию
    assert context.action  # состояние окна


def test_last_error_lands_in_the_report(window: MainWindow, monkeypatch) -> None:
    """Последняя ошибка попадает в отчёт без отдельного учёта."""
    shown: list[ProblemContext] = []

    class FakeDialog:
        def __init__(self, context, **kwargs) -> None:
            shown.append(context)

        def exec(self) -> int:
            return QDialog.Rejected

    monkeypatch.setattr(problem_report, "ProblemReportDialog", FakeDialog)
    window._set_status("Не удалось сохранить результат", "error")

    window.report_problem()

    assert shown[0].error == "Не удалось сохранить результат"


def test_report_does_not_contain_the_full_path(window: MainWindow, rich_docx: Path) -> None:
    """В отчёт идёт имя файла, а не путь с именем пользователя."""
    window.set_source_file(rich_docx)

    context = ProblemContext(document=window._source_file.name)

    assert str(Path.home()) not in context.document
