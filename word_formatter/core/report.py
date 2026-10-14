"""Отчёт об изменениях — то, что пользователь читает после обработки.

Отчёт отвечает на три вопроса: что именно приложение изменило, что оно
не смогло сделать и что нужно проверить вручную. Отчёт строится из журнала
изменений, картины документа и результата проверки сохранности.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

from word_formatter.core.docx import DocumentInventory, Role
from word_formatter.core.profiles import Profile
from word_formatter.core.rules.changes import ChangeLog
from word_formatter.core.verification import VerificationResult, summary_lines

MAX_ITEMS_PER_SECTION = 30

CATEGORY_TITLES: dict[str, str] = {
    "page": "Параметры страницы",
    "document-defaults": "Оформление по умолчанию",
    "body-text": "Основной текст",
    "title-page": "Титульный лист",
    "italic-removed": "Снятый курсив",
    "underline-removed": "Снятое подчёркивание",
    "structural-headings": "Заголовки структурных частей",
    "headings": "Заголовки разделов и подразделов",
    "heading-styles": "Стили заголовков Word",
    "figure-captions": "Подписи рисунков",
    "table-captions": "Названия таблиц",
    "caption-positions": "Положение подписей",
    "caption-positions-manual": "Подписи, требующие проверки",
    "tables": "Таблицы",
    "tables-manual": "Таблицы, требующие проверки",
    "lists": "Перечисления",
    "list-markers": "Маркеры перечислений",
    "footnotes": "Сноски",
    "hyphenation": "Переносы",
    "page-numbers": "Нумерация страниц",
    "blank-lines": "Свободные строки",
}

# Порядок разделов в отчёте: сначала общее оформление, затем частности.
CATEGORY_ORDER: tuple[str, ...] = (
    "page",
    "document-defaults",
    "page-numbers",
    "hyphenation",
    "title-page",
    "body-text",
    "italic-removed",
    "underline-removed",
    "structural-headings",
    "headings",
    "heading-styles",
    "figure-captions",
    "table-captions",
    "caption-positions",
    "tables",
    "lists",
    "list-markers",
    "footnotes",
    "blank-lines",
    "tables-manual",
    "caption-positions-manual",
)

MANUAL_CATEGORIES = frozenset({"tables-manual", "caption-positions-manual"})


@dataclass(frozen=True)
class ReportSection:
    """Раздел отчёта."""

    title: str
    lines: tuple[str, ...]


@dataclass(frozen=True)
class Report:
    """Готовый отчёт: заголовок, краткая сводка и разделы."""

    title: str
    summary: tuple[str, ...]
    sections: tuple[ReportSection, ...]

    def as_text(self) -> str:
        """Отчёт одним текстом — для показа в окне и сохранения в файл."""
        lines = [self.title, "=" * len(self.title), ""]
        lines.extend(self.summary)
        for section in self.sections:
            lines.append("")
            lines.append(section.title)
            lines.append("-" * len(section.title))
            lines.extend(section.lines)
        lines.append("")
        return "\n".join(lines)

    @property
    def manual_checks(self) -> tuple[str, ...]:
        """Пункты, которые пользователю нужно проверить руками."""
        for section in self.sections:
            if section.title == "Требует ручной проверки":
                return section.lines
        return ()


def _category_title(category: str) -> str:
    return CATEGORY_TITLES.get(category, category)


def _change_lines(changes: list) -> tuple[str, ...]:
    lines: list[str] = []
    for change in changes[:MAX_ITEMS_PER_SECTION]:
        text = f"  • {change.description}"
        if change.detail:
            text += f" — {change.detail}"
        if change.text:
            text += f": «{change.text}»"
        elif change.paragraph_index is not None:
            text += f": абзац №{change.paragraph_index + 1}"
        lines.append(text)
    if len(changes) > MAX_ITEMS_PER_SECTION:
        lines.append(f"  … и ещё {len(changes) - MAX_ITEMS_PER_SECTION}")
    return tuple(lines)


def _title_page_note(inventory: DocumentInventory) -> str | None:
    """Предупреждение о титульном листе.

    Методичка не задаёт оформление титульного листа, поэтому абзацы до первой
    структурной части остаются такими, какими их сделал автор, и об этом
    нужно сказать: пользователь должен проверить их сам.
    """
    first_structural = next(
        (
            item.index
            for item in inventory.paragraphs
            if item.role is Role.STRUCTURAL_HEADING
        ),
        None,
    )
    if first_structural is None or first_structural == 0:
        return None
    body_before = [
        item
        for item in inventory.paragraphs
        if item.index < first_structural and item.role is Role.BODY and item.text.strip()
    ]
    if not body_before:
        return None
    return (
        f"Титульный лист ({len(body_before)} абзацев) оставлен без изменений: "
        "методичка не задаёт его оформление — проверьте его вручную"
    )


def _toc_note(inventory: DocumentInventory) -> str | None:
    has_headings = any(
        item.role in (Role.STRUCTURAL_HEADING, Role.HEADING) for item in inventory.paragraphs
    )
    if inventory.has_toc:
        return (
            "Автоматическое содержание: поле помечено для обновления, Word пересчитает его "
            "при открытии документа"
        )
    if has_headings:
        return (
            "Автоматическое содержание в документе не найдено: при необходимости добавьте его "
            "в Word (Ссылки → Оглавление) — структура заголовков для этого уже подготовлена"
        )
    return None


def _unsupported_notes(inventory: DocumentInventory) -> tuple[str, ...]:
    notes: list[str] = []
    if inventory.role_count(Role.FORMULA) or inventory.ole_objects:
        notes.append(
            "Формулы и встроенные объекты: приложение их не переоформляет — проверьте вручную"
        )
    if inventory.content_controls:
        notes.append(
            "Элементы управления содержимым (готовые блоки Word) не изменялись — проверьте вручную"
        )
    if inventory.footnotes:
        notes.append(
            "Если оформление сноски задано вручную внутри текста сноски, оно осталось прежним"
        )
    if inventory.endnotes:
        notes.append("Концевые сноски: проверьте их оформление вручную")
    return tuple(notes)


def _manual_check_lines(changes: ChangeLog, inventory: DocumentInventory) -> tuple[str, ...]:
    lines: list[str] = []

    ambiguous = inventory.ambiguous
    if ambiguous:
        lines.append(f"  • Спорные абзацы: {len(ambiguous)} — оформление не изменялось")
        for item in ambiguous[:MAX_ITEMS_PER_SECTION]:
            reason = item.reason or "признаки противоречат друг другу"
            label = item.short_text or "без текста"
            lines.append(f"      – абзац №{item.index + 1}: «{label}» — {reason}")

    for category in MANUAL_CATEGORIES:
        for change in changes.by_category().get(category, []):
            lines.append(f"  • {change.description} — {change.detail}")

    toc_note = _toc_note(inventory)
    if toc_note:
        lines.append(f"  • {toc_note}")

    title_note = _title_page_note(inventory)
    if title_note:
        lines.append(f"  • {title_note}")

    for note in _unsupported_notes(inventory):
        lines.append(f"  • {note}")

    return tuple(lines)


def build_report(
    *,
    profile: Profile,
    source_name: str,
    changes: ChangeLog,
    inventory: DocumentInventory,
    verification: VerificationResult | None = None,
    created_at: datetime | None = None,
) -> Report:
    """Собирает отчёт об обработке документа."""
    moment = created_at or datetime.now()
    sections: list[ReportSection] = []

    grouped = changes.by_category()
    known = [category for category in CATEGORY_ORDER if category in grouped]
    unknown = sorted(set(grouped) - set(CATEGORY_ORDER) - MANUAL_CATEGORIES)
    done_lines: list[str] = []
    for category in (*known, *unknown):
        if category in MANUAL_CATEGORIES:
            continue
        done_lines.append(f"{_category_title(category)}:")
        done_lines.extend(_change_lines(grouped[category]))
        done_lines.append("")
    while done_lines and not done_lines[-1]:
        done_lines.pop()
    if not done_lines:
        done_lines = ["  • профиль не потребовал ни одного изменения в этом документе"]
    sections.append(ReportSection("Что сделано", tuple(done_lines)))

    manual = _manual_check_lines(changes, inventory)
    sections.append(
        ReportSection(
            "Требует ручной проверки",
            manual or ("Проверять вручную нечего: все правила профиля применимы к этому документу.",),
        )
    )

    if profile.limitations:
        sections.append(
            ReportSection(
                "Ограничения профиля",
                tuple(f"  • {item}" for item in profile.limitations),
            )
        )

    if verification is not None:
        sections.append(
            ReportSection("Проверка сохранности содержимого", summary_lines(verification))
        )

    sections.append(
        ReportSection(
            "Как проверить результат",
            (
                "  1. Откройте сохранённый файл в Word.",
                "  2. Нажмите Ctrl+A, затем F9 — Word обновит номера страниц и содержание.",
                "  3. Просмотрите места, перечисленные в разделе «Требует ручной проверки».",
                "  4. Исходный документ не изменялся: он остался там, где и был.",
            ),
        )
    )

    summary = (
        f"Исходный файл: {source_name}",
        f"Профиль: {profile.name} (версия {profile.version})",
        f"Источник требований: {profile.source_title}",
        f"Дата обработки: {moment.strftime('%d.%m.%Y %H:%M')}",
        "",
        f"Абзацев основного текста: {inventory.role_count(Role.BODY)}; "
        f"заголовков: {len(inventory.headings)}; "
        f"элементов перечислений: {inventory.role_count(Role.LIST_ITEM)}; "
        f"таблиц: {len(inventory.tables)}; "
        f"рисунков: {inventory.role_count(Role.FIGURE)}; "
        f"сносок: {inventory.footnotes}",
        f"Всего записей об изменениях: {len(changes)}",
    )
    if verification is not None:
        summary += (
            "Проверка сохранности содержимого: "
            + ("пройдена" if verification.passed else "обнаружены расхождения"),
        )

    return Report(
        title=f"Отчёт об оформлении документа — {profile.name}",
        summary=summary,
        sections=tuple(sections),
    )
