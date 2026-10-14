"""Правила оформления: применение профиля к документу."""

from __future__ import annotations

from word_formatter.core.rules.blank_lines import apply_blank_lines
from word_formatter.core.rules.body import BodyElement, collect_body
from word_formatter.core.rules.caption_positions import apply_caption_positions
from word_formatter.core.rules.captions import apply_captions
from word_formatter.core.rules.changes import Change, ChangeLog
from word_formatter.core.rules.context import RuleContext
from word_formatter.core.rules.footnotes import apply_note_styles
from word_formatter.core.rules.headings import apply_heading_styles, apply_headings
from word_formatter.core.rules.hyphenation import apply_hyphenation
from word_formatter.core.rules.lists import apply_lists
from word_formatter.core.rules.page import apply_page_settings
from word_formatter.core.rules.page_numbers import apply_page_numbers
from word_formatter.core.rules.runner import apply_all_rules
from word_formatter.core.rules.tables import apply_tables
from word_formatter.core.rules.text import apply_body_text, apply_document_defaults

__all__ = [
    "BodyElement",
    "Change",
    "ChangeLog",
    "RuleContext",
    "apply_all_rules",
    "apply_blank_lines",
    "apply_body_text",
    "apply_caption_positions",
    "apply_captions",
    "apply_document_defaults",
    "apply_heading_styles",
    "apply_headings",
    "apply_hyphenation",
    "apply_lists",
    "apply_note_styles",
    "apply_page_numbers",
    "apply_page_settings",
    "apply_tables",
    "collect_body",
]
