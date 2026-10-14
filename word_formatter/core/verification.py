"""Проверка сохранности содержимого документа.

Приложение обещает менять оформление, а не содержание. Это обещание нужно
проверять на каждом документе, а не только в тестах: сравнение исходного
файла и результата выполняется после обработки, до сохранения отчёта.

Проверка намеренно не пользуется библиотекой python-docx там, где дело
касается «сложных» элементов: подсчёт идёт по XML, чтобы нельзя было
случайно проверить библиотекой саму себя.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from docx.oxml.ns import qn

from word_formatter.core.docx import open_document

MAX_REPORTED_VALUES = 3


@dataclass(frozen=True)
class Difference:
    """Найденное расхождение между исходником и результатом."""

    kind: str
    description: str
    before: str
    after: str


@dataclass(frozen=True)
class VerificationResult:
    """Итог проверки сохранности."""

    differences: tuple[Difference, ...] = ()
    notes: tuple[str, ...] = ()

    @property
    def passed(self) -> bool:
        return not self.differences

    @property
    def problems(self) -> tuple[str, ...]:
        return tuple(
            f"{difference.description}: было {difference.before}, стало {difference.after}"
            for difference in self.differences
        )


def _count(document, tag: str) -> int:
    return sum(1 for _ in document.element.body.iter(qn(tag)))


def _visible_paragraph_texts(document) -> list[str]:
    """Текст непустых абзацев.

    Пустые абзацы не сравниваются: правило свободных строк намеренно
    добавляет их, и это не потеря содержания.
    """
    return [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]


def _table_cell_texts(document) -> list[str]:
    texts: list[str] = []
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text.strip())
    return texts


def _shorten(values: list[str]) -> str:
    shown = values[:MAX_REPORTED_VALUES]
    text = "; ".join(shown)
    if len(values) > len(shown):
        text += f" … и ещё {len(values) - len(shown)}"
    return text or "(пусто)"


def _compare_lists(kind: str, description: str, before: list[str], after: list[str]) -> Difference | None:
    if before == after:
        return None
    return Difference(
        kind=kind,
        description=description,
        before=_shorten(before),
        after=_shorten(after),
    )


def _compare_counts(kind: str, description: str, before: int, after: int) -> Difference | None:
    if before == after:
        return None
    return Difference(kind=kind, description=description, before=str(before), after=str(after))


def _compare_text_order(
    before: list[str], after: list[str], expected_moves: int
) -> tuple[Difference | None, str | None]:
    """Сравнивает текст документа с учётом разрешённых переносов подписей.

    Возвращает расхождение (если текст потерян или изменён) и пояснение
    (если изменился только порядок абзацев). Порядок может измениться
    ровно по одной причине — профиль требует другое положение подписи,
    — и такие переносы приложение перечисляет в отчёте.
    """
    if before == after:
        return None, None

    if sorted(before) == sorted(after):
        if expected_moves > 0:
            return None, (
                "порядок абзацев изменён: подписи перенесены на место, "
                f"которое требует профиль (переносов: {expected_moves})"
            )
        return (
            Difference(
                kind="order",
                description="изменился порядок абзацев",
                before=_shorten(before),
                after=_shorten(after),
            ),
            None,
        )

    return (
        Difference(
            kind="text",
            description="изменился текст документа",
            before=_shorten(before),
            after=_shorten(after),
        ),
        None,
    )


def verify_documents(source, result, *, expected_moves: int = 0) -> VerificationResult:
    """Сравнивает исходный документ и результат обработки.

    expected_moves — сколько переносов подписей разрешено: приложение знает
    об этом из журнала изменений и сообщает проверке.
    """
    differences: list[Difference] = []
    notes: list[str] = []

    text_difference, order_note = _compare_text_order(
        _visible_paragraph_texts(source), _visible_paragraph_texts(result), expected_moves
    )
    if order_note:
        notes.append(order_note)

    checks: list[Difference | None] = [
        text_difference,
        _compare_lists(
            "tables",
            "изменился текст в таблицах",
            _table_cell_texts(source),
            _table_cell_texts(result),
        ),
        _compare_counts(
            "images",
            "изменилось число изображений",
            _count(source, "w:drawing") + _count(source, "w:pict"),
            _count(result, "w:drawing") + _count(result, "w:pict"),
        ),
        _compare_counts(
            "hyperlinks",
            "изменилось число гиперссылок",
            _count(source, "w:hyperlink"),
            _count(result, "w:hyperlink"),
        ),
        _compare_counts(
            "footnotes",
            "изменилось число сносок",
            _count(source, "w:footnoteReference"),
            _count(result, "w:footnoteReference"),
        ),
        _compare_counts(
            "endnotes",
            "изменилось число концевых сносок",
            _count(source, "w:endnoteReference"),
            _count(result, "w:endnoteReference"),
        ),
        _compare_counts(
            "sections",
            "изменилось число разделов документа",
            len(source.sections),
            len(result.sections),
        ),
        _compare_counts(
            "fields",
            "изменилось число полей Word (номеров страниц, содержания и других)",
            _count(source, "w:fldChar"),
            _count(result, "w:fldChar"),
        ),
        _compare_counts(
            "ole",
            "изменилось число встроенных объектов",
            _count(source, "w:object"),
            _count(result, "w:object"),
        ),
        _compare_counts(
            "content-controls",
            "изменилось число элементов управления содержимым",
            _count(source, "w:sdt"),
            _count(result, "w:sdt"),
        ),
        _compare_counts(
            "tables-count",
            "изменилось число таблиц",
            len(source.tables),
            len(result.tables),
        ),
    ]

    for difference in checks:
        if difference is not None:
            differences.append(difference)

    added_paragraphs = len(result.paragraphs) - len(source.paragraphs)
    if added_paragraphs > 0:
        notes.append(f"добавлено пустых абзацев для свободных строк: {added_paragraphs}")

    return VerificationResult(differences=tuple(differences), notes=tuple(notes))


def verify_files(
    source_path: str | Path, result_path: str | Path, *, expected_moves: int = 0
) -> VerificationResult:
    """Проверяет сохранность, открывая оба файла заново с диска."""
    source = open_document(source_path)
    result = open_document(result_path)
    return verify_documents(source, result, expected_moves=expected_moves)


def summary_lines(result: VerificationResult) -> tuple[str, ...]:
    """Строки для отчёта: итог проверки сохранности."""
    if result.passed:
        lines = ["Проверка сохранности содержимого пройдена: текст, таблицы, изображения, "
                 "гиперссылки, сноски, поля и объекты не изменились."]
    else:
        lines = ["Проверка сохранности содержимого обнаружила расхождения:"]
        lines.extend(f"  • {problem}" for problem in result.problems)
        lines.append("Сохраните результат под другим именем и проверьте эти места вручную.")
    lines.extend(f"  {note}" for note in result.notes)
    return tuple(lines)
