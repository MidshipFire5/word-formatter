"""Применение всех правил профиля к документу.

Порядок важен: сначала общие настройки документа, затем определения стилей,
и только потом оформление конкретных абзацев. Иначе назначение стиля
заголовку перебило бы уже применённое оформление.
"""

from __future__ import annotations

from word_formatter.core.docx import inspect_document
from word_formatter.core.rules.blank_lines import apply_blank_lines
from word_formatter.core.rules.caption_positions import apply_caption_positions
from word_formatter.core.rules.captions import apply_captions
from word_formatter.core.rules.changes import ChangeLog
from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.footnotes import apply_note_styles
from word_formatter.core.rules.headings import apply_heading_styles, apply_headings
from word_formatter.core.rules.hyphenation import apply_hyphenation
from word_formatter.core.rules.lists import apply_lists
from word_formatter.core.rules.page import apply_page_settings
from word_formatter.core.rules.page_numbers import apply_page_numbers
from word_formatter.core.rules.tables import apply_tables
from word_formatter.core.rules.text import apply_body_text, apply_document_defaults


def apply_all_rules(context: RuleContext) -> ChangeLog:
    """Применяет правила профиля к документу и возвращает журнал изменений."""
    apply_page_settings(context)
    apply_document_defaults(context)
    apply_heading_styles(context)
    apply_hyphenation(context)

    apply_body_text(context)
    apply_lists(context)
    apply_headings(context)
    apply_captions(context)
    apply_tables(context)
    apply_note_styles(context)
    apply_page_numbers(context)

    # Перенос подписей меняет порядок элементов тела документа. После него
    # картину документа нужно собрать заново: иначе правило свободных строк
    # будет работать со старыми номерами абзацев.
    context.order_changes += apply_caption_positions(context)
    if context.order_changes:
        context.inventory = inspect_document(context.document, context.profile)

    # Свободные строки добавляются последними: они сдвигают нумерацию абзацев,
    # поэтому все остальные правила должны отработать до них.
    apply_blank_lines(context)

    return context.changes
