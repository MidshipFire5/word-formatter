"""Обработка документа: полный проход от исходного файла до отчёта.

Порядок действий важен и повторяет обещания приложения:
исходник открывается только для чтения, результат пишется в отдельный файл
через временный файл, после сохранения содержимое проверяется сравнением
с исходником, и только затем собирается отчёт.
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from word_formatter.core.docx import DocumentInventory, inspect_document, open_document
from word_formatter.core.errors import SaveError
from word_formatter.core.profiles import Profile
from word_formatter.core.report import Report, build_report
from word_formatter.core.rules import RuleContext, apply_all_rules
from word_formatter.core.rules.changes import ChangeLog
from word_formatter.core.verification import VerificationResult, verify_files

OUTPUT_SUFFIX = "_оформлено"


@dataclass(frozen=True)
class ProcessingResult:
    """Результат обработки документа."""

    source_path: Path
    output_path: Path
    report: Report
    verification: VerificationResult
    changes: ChangeLog
    inventory: DocumentInventory


def default_output_name(source_path: str | Path) -> str:
    """Имя файла результата по умолчанию."""
    path = Path(source_path)
    suffix = path.suffix or ".docx"
    return f"{path.stem}{OUTPUT_SUFFIX}{suffix}"


def ensure_output_is_safe(source_path: str | Path, output_path: str | Path) -> Path:
    """Проверяет, что результат не затрёт исходный документ."""
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if source == output:
        raise SaveError(
            "Результат нельзя сохранять поверх исходного документа: "
            "исходный файл должен остаться неизменным. Выберите другое имя файла.",
            details=f"путь: {output}",
        )
    return Path(output_path)


def _save_atomically(document, target: Path) -> None:
    """Сохраняет документ через временный файл рядом с целью.

    Так неудачное сохранение не оставит обрезанный документ на месте готового.
    """
    directory = target.parent
    if not directory.is_dir():
        raise SaveError(
            f"Папка «{directory}» не найдена. Выберите другую папку для сохранения.",
            details=f"путь: {directory}",
        )

    handle, temporary_name = tempfile.mkstemp(prefix=".word-formatter-", suffix=".docx", dir=str(directory))
    os.close(handle)
    temporary_path = Path(temporary_name)

    try:
        document.save(str(temporary_path))
        os.replace(temporary_path, target)
    except PermissionError as error:
        temporary_path.unlink(missing_ok=True)
        raise SaveError(
            f"Не удалось сохранить файл «{target.name}»: он открыт в другой программе. "
            "Закройте документ в Word и попробуйте снова.",
            details=str(error),
        ) from error
    except OSError as error:
        temporary_path.unlink(missing_ok=True)
        raise SaveError(
            f"Не удалось сохранить файл «{target.name}»: {error.strerror or error}. "
            "Проверьте, что в папке есть место и разрешение на запись.",
            details=str(error),
        ) from error


def process_document(
    source_path: str | Path,
    output_path: str | Path,
    profile: Profile,
    *,
    progress: Callable[[str], None] | None = None,
) -> ProcessingResult:
    """Применяет профиль к документу и сохраняет результат в отдельный файл.

    Через progress можно получать названия этапов работы — это использует
    графический интерфейс, чтобы показать пользователю, что происходит.
    """
    def step(message: str) -> None:
        if progress is not None:
            progress(message)

    source = Path(source_path)
    target = ensure_output_is_safe(source, output_path)

    step("Чтение документа")
    document = open_document(source)

    step("Разбор структуры документа")
    inventory = inspect_document(document, profile)

    step("Применение правил оформления")
    context = RuleContext(document=document, profile=profile, inventory=inventory)
    changes = apply_all_rules(context)

    step("Сохранение результата")
    _save_atomically(document, target)

    step("Проверка сохранности содержимого")
    verification = verify_files(source, target, expected_moves=context.order_changes)

    step("Подготовка отчёта")
    report = build_report(
        profile=profile,
        source_name=source.name,
        changes=changes,
        inventory=context.inventory,
        verification=verification,
    )

    return ProcessingResult(
        source_path=source,
        output_path=target,
        report=report,
        verification=verification,
        changes=changes,
        inventory=context.inventory,
    )
