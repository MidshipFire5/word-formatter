"""Профили оформления: модель, проверка, каталог и загрузка."""

from __future__ import annotations

from word_formatter.core.profiles.loader import (
    ProfileCatalog,
    ProfileInfo,
    describe_profile_file,
    load_profile_from_path,
    read_profile_file,
)
from word_formatter.core.profiles.model import (
    ORIGIN_BUILTIN,
    ORIGIN_USER,
    PAGE_FORMATS_MM,
    BlankLinesSettings,
    BodySettings,
    CaptionStyle,
    CaptionsSettings,
    FootnotesSettings,
    HeadingsSettings,
    HyphenationSettings,
    ListsSettings,
    Margins,
    PageNumbersSettings,
    PageSettings,
    Profile,
    StructuralHeading,
    TableSettings,
    TextStyle,
)
from word_formatter.core.profiles.validation import Issue, validate_profile

__all__ = [
    "ORIGIN_BUILTIN",
    "ORIGIN_USER",
    "PAGE_FORMATS_MM",
    "BlankLinesSettings",
    "BodySettings",
    "CaptionStyle",
    "CaptionsSettings",
    "FootnotesSettings",
    "HeadingsSettings",
    "HyphenationSettings",
    "Issue",
    "ListsSettings",
    "Margins",
    "PageNumbersSettings",
    "PageSettings",
    "Profile",
    "ProfileCatalog",
    "ProfileInfo",
    "StructuralHeading",
    "TableSettings",
    "TextStyle",
    "describe_profile_file",
    "load_profile_from_path",
    "read_profile_file",
    "validate_profile",
]
