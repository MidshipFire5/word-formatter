"""Разбор требований из текста методички.

Методички пишут по-разному, но требования в них повторяются: шрифт, размер,
интервал, поля, отступ, выравнивание, нумерация страниц, подписи, таблицы,
сноски. Разбор идёт по предложениям: предложение — это и источник значения,
и цитата для отчёта, чтобы человек видел, откуда взято каждое число.

Приложение разбирает только то, что написано явно. Формулировки вроде
«стандартные поля» или «оформление по ГОСТ» остаются неразобранными и
попадают в список «проверьте»: угадывать здесь опаснее, чем спросить.
"""

from __future__ import annotations

import re
from dataclasses import replace
from typing import Callable

from word_formatter.core.profiles.extraction.draft import ORIGIN_TEXT, DraftValue
from word_formatter.core.profiles.extraction.sources import SourceText

SENTENCE_SPLIT = re.compile(r"(?<=[.;!?])\s+|\n+")
SPACES = re.compile(r"[\s\u00a0\u202f]+")

FONT_NAMES = (
    ("times new roman", "Times New Roman"),
    ("таймс нью роман", "Times New Roman"),
    ("arial", "Arial"),
    ("calibri", "Calibri"),
    ("courier new", "Courier New"),
    ("georgia", "Georgia"),
    ("verdana", "Verdana"),
    ("gost type a", "GOST type A"),
    ("gost type b", "GOST type B"),
)

SIZE_TOKEN = re.compile(r"(\d{2})\s*(?:пт\.?|пункт(?:ов|а|ы)?|pt)\b", re.IGNORECASE)
SIZE_WORDS = ("шрифт", "размер", "кегл", "пт", "пункт")
MEASURE_TOKEN = re.compile(r"(\d+(?:[.,]\d+)?)\s*(см|мм|знак(?:ов|а|и)?)\b", re.IGNORECASE)
MARGIN_TOKEN = re.compile(
    # Составные названия идут первыми: иначе «верхнее и нижнее — 20 мм»
    # разберётся как одно верхнее поле, а нижнее потеряется.
    r"(верхнее и нижнее|верхний и нижний|левое|правого|правое|правый|верхнее|нижнее)"
    r"[^0-9\n]{0,25}(\d+(?:[.,]\d+)?)\s*(мм|см)",
    re.IGNORECASE,
)
SIDE_PATHS = {
    "левое": "page.marginsMm.left",
    "правое": "page.marginsMm.right",
    "правого": "page.marginsMm.right",
    "правый": "page.marginsMm.right",
    "верхнее": "page.marginsMm.top",
    "нижнее": "page.marginsMm.bottom",
    "верхнее и нижнее": "page.marginsMm.top_bottom",
    "верхний и нижний": "page.marginsMm.top_bottom",
}
MARGIN_TITLES = {
    "page.marginsMm.left": "левое поле",
    "page.marginsMm.right": "правое поле",
    "page.marginsMm.top": "верхнее поле",
    "page.marginsMm.bottom": "нижнее поле",
}

SPACING_VALUES = (
    ("полтора", 1.5),
    ("полуторн", 1.5),
    ("одинарн", 1.0),
    ("двойн", 2.0),
)
SPACING_NUMBER = re.compile(r"\b(1[,.]0|1[,.]5|2[,.]0|2)\b")

NEGATION_BEFORE = re.compile(
    r"(?:^|\W)(?:не|без|запрещ\w*|исключ\w*)\W+(?:\w+\W+){0,3}$", re.IGNORECASE
)
NEGATION_AFTER = re.compile(
    r"^.{0,60}?\b(?:не\s+(?:допуска|разреша|использ|применя|печата)|запрещ)",
    re.IGNORECASE | re.DOTALL,
)

ALIGNMENT_WORDS = (
    ("по ширине", "justify"),
    ("по центру", "center"),
    ("по левому краю", "left"),
    ("по правому краю", "right"),
    ("по левому полю", "left"),
)


def sentences(text: str) -> list[str]:
    """Делит текст на предложения, убирая лишние пробелы."""
    parts = [SPACES.sub(" ", part).strip() for part in SENTENCE_SPLIT.split(text)]
    return [part for part in parts if part]


def _normalize(text: str) -> str:
    """Приводит текст к виду, удобному для сравнения.

    В PDF буквы «ё» и «й» нередко извлекаются искажённо: «й» превращается
    в «и» с мусорным знаком («второи 9» вместо «второй»), «ё» — в латинскую
    букву («подчеHркивание»). Поэтому сравнение идёт без «ё» и без «й»:
    так требование находится и в тексте, и в PDF.
    """
    return (
        text.replace("ё", "е")
        .replace("Ё", "Е")
        .replace("й", "и")
        .replace("Й", "И")
    )


# Слова короче этой длины сравниваются строго: иначе «пол» совпадёт
# с чем угодно.
FUZZY_MINIMUM = 6
_FUZZY_CACHE: dict[str, re.Pattern] = {}


def _fuzzy(word: str) -> re.Pattern:
    """Шаблон, допускающий мусор между буквами слова.

    В PDF таблица символов бывает испорчена: «подчёркивание» извлекается как
    «подчеHркивание», а «второй» — как «второи 9». Такой шаблон находит слово
    и в этих случаях, а на обычном тексте ведёт себя как поиск подстроки
    (строгий поиск пробуется раньше).
    """
    pattern = _FUZZY_CACHE.get(word)
    if pattern is None:
        expression = "".join(re.escape(char) + r"[\w\s]{0,2}?" for char in word)
        pattern = re.compile(expression, re.IGNORECASE)
        _FUZZY_CACHE[word] = pattern
    return pattern


def _has(sentence: str, *words: str) -> bool:
    lowered = _normalize(sentence.lower())
    for word in words:
        normalized = _normalize(word)
        if normalized in lowered:
            return True
        if len(normalized) >= FUZZY_MINIMUM and _fuzzy(normalized).search(lowered):
            return True
    return False


def _negated(sentence: str, *keywords: str) -> bool:
    """Есть ли отрицание при самом слове-признаке.

    Смотреть отрицание по всему предложению нельзя: во фразе «печатают
    прописными буквами, полужирным шрифтом, без абзацного отступа» слово
    «без» относится к отступу, а не к прописным буквам. Поэтому отрицание
    ищется рядом с ключевым словом — перед ним или сразу после него.
    """
    sentence = _normalize(sentence)
    for keyword in keywords:
        for match in _fuzzy(_normalize(keyword)).finditer(sentence):
            before = sentence[: match.start()]
            after = sentence[match.end() :]
            if NEGATION_BEFORE.search(before) or NEGATION_AFTER.match(after):
                return True
    return False


def _short(sentence: str, limit: int = 160) -> str:
    return sentence if len(sentence) <= limit else sentence[: limit - 1] + "…"


def _value(path: str, title: str, value, sentence: str) -> DraftValue:
    return DraftValue(
        path=path, value=value, origin=ORIGIN_TEXT, evidence=_short(sentence), title=title
    )


# --- отдельные требования -----------------------------------------------------


def extract_font(sentence: str) -> list[DraftValue]:
    lowered = sentence.lower()
    for needle, name in FONT_NAMES:
        if needle in lowered and _has(sentence, "шрифт", "набира", "гарнитур", "оформля", "текст"):
            return [_value("body.font", "шрифт основного текста", name, sentence)]
    return []


def extract_sizes(sentence: str) -> list[DraftValue]:
    if not _has(sentence, *SIZE_WORDS):
        return []
    match = SIZE_TOKEN.search(sentence)
    if match is None:
        return []
    size = float(match.group(1))

    if _has(sentence, "таблиц", "ячейк"):
        return [_value("tables.fontSizePt", "размер шрифта в таблицах", size, sentence)]
    if _has(sentence, "снос"):
        return [_value("footnotes.fontSizePt", "размер шрифта сносок", size, sentence)]
    if _has(sentence, "заголовк"):
        return [_value("headings.levels.1.sizePt", "размер шрифта заголовков", size, sentence)]
    return [_value("body.sizePt", "размер шрифта основного текста", size, sentence)]


def extract_spacing(sentence: str) -> list[DraftValue]:
    if not _has(sentence, "интервал"):
        return []
    for needle, value in SPACING_VALUES:
        if _has(sentence, needle):
            return [_value("body.lineSpacing", "межстрочный интервал", value, sentence)]
    match = SPACING_NUMBER.search(sentence)
    if match is not None:
        value = float(match.group(1).replace(",", "."))
        if value in (1.0, 1.5, 2.0):
            return [_value("body.lineSpacing", "межстрочный интервал", value, sentence)]
    return []


def extract_indent(sentence: str) -> list[DraftValue]:
    if not _has(sentence, "отступ"):
        return []
    if _has(sentence, "отступ от", "отступ таблицы", "отступ от левого"):
        return []
    match = MEASURE_TOKEN.search(sentence)
    if match is None:
        return []

    number = float(match.group(1).replace(",", "."))
    unit = match.group(2).lower()
    if unit.startswith("мм"):
        centimetres = round(number / 10, 2)
    elif unit.startswith("знак"):
        # Пять знаков при 14 пт — это те же 1,25 см.
        centimetres = round(number * 0.25, 2)
    else:
        centimetres = number

    if not 0 <= centimetres <= 5:
        return []
    return [_value("body.firstLineIndentCm", "абзацный отступ", centimetres, sentence)]


def extract_margins(sentence: str) -> list[DraftValue]:
    if not _has(sentence, "пол"):
        return []
    values: list[DraftValue] = []
    for match in MARGIN_TOKEN.finditer(sentence):
        side = match.group(1).lower()
        number = float(match.group(2).replace(",", "."))
        unit = match.group(3).lower()
        millimetres = round(number * 10) if unit.startswith("см") else round(number)
        if not 5 <= millimetres <= 60:
            continue
        path = SIDE_PATHS.get(side)
        if path is None:
            continue
        if path.endswith("top_bottom"):
            values.append(_value("page.marginsMm.top", MARGIN_TITLES["page.marginsMm.top"], millimetres, sentence))
            values.append(
                _value("page.marginsMm.bottom", MARGIN_TITLES["page.marginsMm.bottom"], millimetres, sentence)
            )
        else:
            values.append(_value(path, MARGIN_TITLES[path], millimetres, sentence))
    return values


def extract_alignment(sentence: str) -> list[DraftValue]:
    if not _has(sentence, "выравнива", "выравнивают", "располагают"):
        return []
    for needle, value in ALIGNMENT_WORDS:
        if _has(sentence, needle):
            return [_value("body.alignment", "выравнивание основного текста", value, sentence)]
    return []


def extract_page_numbers(sentence: str) -> list[DraftValue]:
    if not _has(sentence, "нумерац", "номер страниц", "нумеруются"):
        return []
    # Заголовок «Нумерация» сам по себе требованием не является: нужны слова
    # про страницы, цифры или сквозной порядок.
    if not _has(sentence, "страниц", "цифр", "сквозн"):
        return []
    values = [
        _value("pageNumbers.enabled", "нумерация страниц", True, sentence),
    ]
    lowered = sentence.lower()

    if _has(sentence, "со второй страницы", "начиная со второй", "второй страницы"):
        values.append(
            _value("pageNumbers.skipFirstPage", "номер со второй страницы", True, sentence)
        )
    if _has(sentence, "титульный лист не нумеру", "первая страница не нумеру", "не проставля"):
        values.append(
            _value("pageNumbers.skipFirstPage", "номер со второй страницы", True, sentence)
        )

    if "внизу" in lowered or "снизу" in lowered:
        position = "bottom-center" if "центр" in lowered else "bottom-right"
        if "слева" in lowered:
            position = "bottom-left"
        values.append(_value("pageNumbers.position", "положение номера страницы", position, sentence))
    elif "вверху" in lowered or "сверху" in lowered:
        position = "top-center" if "центр" in lowered else "top-right"
        if "слева" in lowered:
            position = "top-left"
        values.append(_value("pageNumbers.position", "положение номера страницы", position, sentence))
    return values


def extract_hyphenation(sentence: str) -> list[DraftValue]:
    if not _has(sentence, "перенос"):
        return []
    if _has(sentence, "автоматическ", "переносы слов", "расстановк"):
        enabled = not _negated(sentence, "перенос")
        return [_value("hyphenation.enabled", "автоматические переносы", enabled, sentence)]
    return []


def extract_headings(sentence: str) -> list[DraftValue]:
    values: list[DraftValue] = []
    if _has(sentence, "заголовк", "структурн", "раздел"):
        if _has(sentence, "прописн", "заглавн"):
            values.append(
                _value(
                    "headings.structural.uppercase",
                    "структурные заголовки прописными",
                    not _negated(sentence, "прописн", "заглавн"),
                    sentence,
                )
            )
        if _has(sentence, "полужирн", "жирн"):
            values.append(
                _value(
                    "headings.structural.bold",
                    "структурные заголовки полужирным",
                    not _negated(sentence, "полужирн", "жирн"),
                    sentence,
                )
            )
    return values


def extract_captions(sentence: str) -> list[DraftValue]:
    values: list[DraftValue] = []
    if _has(sentence, "подпис", "рисуно", "рисунк", "иллюстрац"):
        if _has(sentence, "под рисунком", "под иллюстрацией", "под рисуно"):
            values.append(
                _value("captions.figure.position", "подпись рисунка", "below", sentence)
            )
        elif _has(sentence, "над рисунком", "над иллюстрацией"):
            values.append(
                _value("captions.figure.position", "подпись рисунка", "above", sentence)
            )
    if _has(sentence, "таблиц"):
        if _has(sentence, "над таблиц", "перед таблиц"):
            values.append(_value("captions.table.position", "название таблицы", "above", sentence))
        elif _has(sentence, "под таблиц"):
            values.append(_value("captions.table.position", "название таблицы", "below", sentence))
    return values


def extract_blank_lines(sentence: str) -> list[DraftValue]:
    if _has(sentence, "свободн", "пуст") and _has(sentence, "строк", "интервал"):
        if _has(sentence, "таблиц", "рисунк", "заголовк", "формул"):
            return [
                _value("blankLines.insert", "свободные строки", True, sentence),
            ]
    return []


def extract_emphasis(sentence: str) -> list[DraftValue]:
    values: list[DraftValue] = []
    if _has(sentence, "курсив") and _negated(sentence, "курсив"):
        values.append(_value("body.removeItalic", "снимать курсив", True, sentence))
    if _has(sentence, "подчёркиван", "подчеркиван") and _negated(
        sentence, "подчёркиван", "подчеркиван"
    ):
        values.append(_value("body.removeUnderline", "снимать подчёркивание", True, sentence))
    return values


EXTRACTORS: tuple[Callable[[str], list[DraftValue]], ...] = (
    extract_font,
    extract_sizes,
    extract_spacing,
    extract_indent,
    extract_margins,
    extract_alignment,
    extract_page_numbers,
    extract_hyphenation,
    extract_headings,
    extract_captions,
    extract_blank_lines,
    extract_emphasis,
)


def extract_from_text(source: SourceText) -> list[DraftValue]:
    """Разбирает требования из текста методички.

    Первое найденное значение по каждому полю считается основным: обычно
    требование формулируют один раз, а дальше лишь повторяют.
    """
    found: dict[str, DraftValue] = {}
    multiple_pages = source.page_count > 1

    for page_number, page in enumerate(source.pages, start=1):
        for sentence in sentences(page):
            for extractor in EXTRACTORS:
                for item in extractor(sentence):
                    if item.path in found:
                        continue
                    if multiple_pages:
                        item = replace(item, evidence=f"{item.evidence} — стр. {page_number}")
                    found[item.path] = item

    return list(found.values())
