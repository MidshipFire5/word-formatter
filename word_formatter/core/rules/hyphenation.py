"""Правила расстановки переносов.

Автоматическая расстановка переносов — настройка всего документа.
Методичка требует исключить из переносов заголовки, названия таблиц и
подписи рисунков: это делается запретом переносов в самих абзацах
(модули заголовков и подписей).
"""

from __future__ import annotations

from docx.oxml.ns import qn

from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.xml_helpers import SETTINGS_ORDER, ensure_ordered_child

CATEGORY_HYPHENATION = "hyphenation"


def apply_hyphenation(context: RuleContext) -> None:
    """Включает автоматическую расстановку переносов по профилю."""
    settings = context.profile.hyphenation
    settings_element = context.document.settings.element

    auto = ensure_ordered_child(settings_element, "w:autoHyphenation", SETTINGS_ORDER)
    auto.set(qn("w:val"), "true" if settings.enabled else "false")

    if settings.enabled:
        # Аббревиатуры и сокращения пишутся прописными: их переносить нельзя.
        caps = ensure_ordered_child(settings_element, "w:doNotHyphenateCaps", SETTINGS_ORDER)
        caps.set(qn("w:val"), "true")

    if settings.enabled:
        context.changes.add(
            CATEGORY_HYPHENATION,
            "включена автоматическая расстановка переносов; слова, написанные "
            "прописными буквами, не переносятся",
            detail="заголовки, названия таблиц и подписи рисунков исключены из переносов "
            "отдельно, в своих правилах",
        )
    else:
        context.changes.add(
            CATEGORY_HYPHENATION,
            "автоматическая расстановка переносов отключена профилем",
        )
