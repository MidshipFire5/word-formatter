"""Отчёт об обработке на настоящем документе."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from word_formatter.core.docx import inspect_document, open_document
from word_formatter.core.profiles import load_profile_from_path
from word_formatter.core.report import build_report
from word_formatter.core.rules import RuleContext, apply_all_rules
from word_formatter.core.verification import verify_files


@pytest.fixture()
def processed(rich_docx: Path, profile_file: Path, tmp_path: Path):
    document = open_document(rich_docx)
    profile = load_profile_from_path(profile_file)
    context = RuleContext(
        document=document,
        profile=profile,
        inventory=inspect_document(document, profile),
    )
    changes = apply_all_rules(context)

    target = tmp_path / "оформлено.docx"
    document.save(str(target))
    verification = verify_files(rich_docx, target)
    report = build_report(
        profile=profile,
        source_name=rich_docx.name,
        changes=changes,
        inventory=context.inventory,
        verification=verification,
        created_at=datetime(2026, 10, 14, 12, 0),
    )
    return report, changes, context.inventory, verification


def test_report_header(processed) -> None:
    report, _changes, _inventory, _verification = processed
    text = report.as_text()

    assert "Отчёт об оформлении документа" in text
    assert "Финансовый университет, Краснодарский филиал" in text
    assert "source.docx" in text
    assert "14.10.2026 12:00" in text
    assert "версия 1.0.0" in text


def test_report_lists_what_was_done(processed) -> None:
    report, _changes, _inventory, _verification = processed
    text = report.as_text()

    for expected in (
        "Что сделано",
        "Параметры страницы:",
        "Оформление по умолчанию:",
        "Нумерация страниц:",
        "Основной текст:",
        "Заголовки структурных частей:",
        "Подписи рисунков:",
        "Таблицы:",
        "Переносы:",
        "Свободные строки:",
    ):
        assert expected in text, f"в отчёте нет строки «{expected}»"


def test_report_mentions_removed_emphasis(processed) -> None:
    report, _changes, _inventory, _verification = processed
    assert "Снятый курсив:" in report.as_text()
    assert "Снятое подчёркивание:" in report.as_text()


def test_report_lists_manual_checks(processed) -> None:
    report, _changes, _inventory, verification = processed
    manual = "\n".join(report.manual_checks)

    assert "Спорные абзацы" in manual
    assert "встроенный объект" in manual
    assert "титульный лист" in manual.lower()
    assert "содержание" in manual.lower()
    assert verification.passed is True


def test_report_lists_profile_limitations(processed) -> None:
    report, _changes, _inventory, _verification = processed
    text = report.as_text()

    assert "Ограничения профиля" in text
    assert "Автоматическое содержание не пересчитывается" in text


def test_report_confirms_content_preservation(processed) -> None:
    report, _changes, _inventory, _verification = processed
    text = report.as_text()

    assert "Проверка сохранности содержимого" in text
    assert "пройдена" in text


def test_report_summary_counts(processed) -> None:
    report, _changes, _inventory, _verification = processed
    summary = "\n".join(report.summary)

    assert "Абзацев основного текста:" in summary
    assert "таблиц: 1" in summary
    assert "сносок: 1" in summary
    assert "Проверка сохранности содержимого: пройдена" in summary


def test_report_explains_how_to_check(processed) -> None:
    report, _changes, _inventory, _verification = processed
    text = report.as_text()

    assert "Как проверить результат" in text
    assert "Ctrl+A" in text
    assert "Исходный документ не изменялся" in text


def test_report_is_plain_text(processed) -> None:
    report, _changes, _inventory, _verification = processed
    text = report.as_text()

    assert isinstance(text, str)
    assert text.endswith("\n")
    assert len(text.splitlines()) > 30


def test_long_change_lists_are_truncated(profile_file: Path, tmp_path: Path) -> None:
    """В отчёте не может быть тысячи строк: длинные списки сворачиваются."""
    from docx import Document

    from word_formatter.core.rules import apply_all_rules

    document = Document()
    for number in range(40):
        paragraph = document.add_paragraph(f"Абзац номер {number} с курсивом.")
        paragraph.runs[0].italic = True

    profile = load_profile_from_path(profile_file)
    context = RuleContext(
        document=document,
        profile=profile,
        inventory=inspect_document(document, profile),
    )
    changes = apply_all_rules(context)
    report = build_report(
        profile=profile,
        source_name="большой.docx",
        changes=changes,
        inventory=context.inventory,
    )

    assert "… и ещё" in report.as_text()
