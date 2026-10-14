"""Сборка черновика профиля из методички.

Методичку читают двумя способами сразу:

1. **текст требований** — то, что в ней написано («шрифт 12, Times New Roman»);
2. **оформление файла** — то, как она свёрстана (поля, стили, интервалы).

Если источники расходятся, побеждает текст: именно он выражает требование.
Расхождение не скрывается, а показывается человеку — возможно, методичку
сверстали не по её же требованиям, и это стоит проверить глазами.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from word_formatter.core.docx.opening import open_document
from word_formatter.core.errors import FormatterError, MethodologyError, ProfileError
from word_formatter.core.profiles.extraction.draft import (
    DraftConflict,
    DraftValue,
    ProfileDraft,
)
from word_formatter.core.profiles.extraction.from_sample import extract_from_sample
from word_formatter.core.profiles.extraction.from_text import extract_from_text
from word_formatter.core.profiles.extraction.sources import (
    SAMPLE_SUFFIXES,
    SourceText,
    read_source,
)
from word_formatter.core.profiles.validation import validate_profile

TRANSLITERATION = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e", "ж": "zh",
    "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o",
    "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f", "х": "h", "ц": "c",
    "ч": "ch", "ш": "sh", "щ": "sch", "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu",
    "я": "ya",
}
IDENTIFIER_CLEANUP = re.compile(r"[^a-z0-9]+")
NO_REQUIREMENTS = (
    "В методичке не нашлось ни одного требования к оформлению. "
    "Приложение ищет в тексте слова про шрифт, размер, интервалы, поля, отступ, "
    "выравнивание, нумерацию страниц, подписи, таблицы и сноски. "
    "Проверьте, что выбран нужный файл и что в нём есть текст требований."
)


def build_profile_draft(path: str | Path) -> ProfileDraft:
    """Читает методичку и собирает черновик профиля."""
    source_path = Path(path)
    source = read_source(source_path)

    text_values = extract_from_text(source)
    sample_values: list[DraftValue] = []

    if source_path.suffix.lower() in SAMPLE_SUFFIXES:
        try:
            document = open_document(source_path)
        except FormatterError:
            document = None
        if document is not None:
            sample_values = extract_from_sample(document)

    draft = merge_values(text_values, sample_values, text=source.text)
    if not draft.values:
        raise MethodologyError(NO_REQUIREMENTS, details=f"файл: {source_path}")
    return draft


def merge_values(
    text_values: list[DraftValue],
    sample_values: list[DraftValue],
    *,
    text: str = "",
) -> ProfileDraft:
    """Объединяет два источника: текст важнее оформления."""
    by_path = {item.path: item for item in sample_values}
    conflicts: list[DraftConflict] = []

    for item in text_values:
        sample = by_path.get(item.path)
        if sample is not None and not _same(sample.value, item.value):
            conflicts.append(
                DraftConflict(
                    path=item.path,
                    title=item.title or sample.title,
                    from_text=item.value,
                    from_sample=sample.value,
                )
            )
        by_path[item.path] = item

    return ProfileDraft(
        values=list(by_path.values()),
        conflicts=conflicts,
        text_excerpt=text[:6000],
        sample_measured=bool(sample_values),
    )


def _same(first, second) -> bool:
    """Числа сравниваем с допуском: 20 и 20.0 — одно и то же."""
    if isinstance(first, bool) or isinstance(second, bool):
        return first == second
    if isinstance(first, (int, float)) and isinstance(second, (int, float)):
        return abs(float(first) - float(second)) < 0.05
    return first == second


def make_profile_id(name: str, *, taken: set[str]) -> str:
    """Идентификатор профиля из его названия: латиница, цифры и дефисы."""
    transliterated = "".join(TRANSLITERATION.get(char, char) for char in name.lower())
    identifier = IDENTIFIER_CLEANUP.sub("-", transliterated).strip("-") or "profile"
    if identifier not in taken:
        return identifier

    number = 2
    while f"{identifier}-{number}" in taken:
        number += 1
    return f"{identifier}-{number}"


def save_draft_profile(
    draft: ProfileDraft,
    *,
    name: str,
    directory: str | Path,
) -> Path:
    """Сохраняет собранный профиль в личную папку профилей.

    Перед записью профиль проверяется тем же строгим правилом, что и любой
    другой: в личную папку не должен попасть файл с ошибкой.
    """
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)

    taken = {item.stem for item in directory.glob("*.json")}
    profile_id = make_profile_id(name, taken=taken)
    data = draft.to_profile(profile_id=profile_id, name=name)

    issues = validate_profile(data)
    if issues:
        raise ProfileError(
            "Собранный профиль не прошёл проверку, поэтому не сохранён: "
            + "; ".join(str(issue) for issue in issues)
        )

    path = directory / f"{profile_id}.json"
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return path


__all__ = [
    "SourceText",
    "build_profile_draft",
    "make_profile_id",
    "merge_values",
    "read_source",
    "save_draft_profile",
]
