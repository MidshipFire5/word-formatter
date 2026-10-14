"""Создание профиля оформления из методички.

Приложение читает методичку — Word, PDF или текст — и собирает черновик
профиля: требования из текста плюс измеренное оформление файла. Черновик
показывается человеку со ссылками на источник каждого значения, и только
после подтверждения становится обычным профилем.
"""

from __future__ import annotations

from word_formatter.core.profiles.extraction.builder import (
    build_profile_draft,
    make_profile_id,
    merge_values,
    save_draft_profile,
)
from word_formatter.core.profiles.extraction.draft import (
    BASE_PROFILE,
    DESCRIBED_FIELDS,
    ORIGIN_EDIT,
    ORIGIN_SAMPLE,
    ORIGIN_TEXT,
    VALUE_CODES,
    DraftConflict,
    DraftRow,
    DraftValue,
    ProfileDraft,
    coerce_value,
)
from word_formatter.core.profiles.extraction.sources import (
    FILE_FILTER,
    SAMPLE_SUFFIXES,
    TEXT_SUFFIXES,
    SourceText,
    read_source,
)

__all__ = [
    "BASE_PROFILE",
    "DESCRIBED_FIELDS",
    "FILE_FILTER",
    "ORIGIN_EDIT",
    "ORIGIN_SAMPLE",
    "ORIGIN_TEXT",
    "SAMPLE_SUFFIXES",
    "TEXT_SUFFIXES",
    "VALUE_CODES",
    "DraftConflict",
    "DraftRow",
    "DraftValue",
    "ProfileDraft",
    "SourceText",
    "build_profile_draft",
    "coerce_value",
    "make_profile_id",
    "merge_values",
    "read_source",
    "save_draft_profile",
]
