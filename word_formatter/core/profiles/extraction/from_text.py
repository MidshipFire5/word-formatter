"""Разбор требований из текста методички.

Методички пишут по-разному, но требования в них повторяются: шрифт, размер,
интервал, поля, отступ, выравнивание, нумерация страниц, подписи, таблицы,
сноски. Разбор идёт по кускам текста: кусок — это и источник значения,
и цитата для отчёта, чтобы человек видел, откуда взято каждое число.

Отдельная сложность — текст из PDF. Там теряются пробелы и портятся буквы:
«Размер шрифта: основного текста в таблицах» превращается в
«Размершрифта:основноготекста втаблицах», «1,5 интервала» — в «1 ,5
интервала», а «й» извлекается как «и» с мусорным знаком. Поэтому сравнение
идёт по «спрессованному» виду строки: без пробелов, без «ё» и без «й».
Цитата при этом остаётся настоящей, как в файле.

Разбирается только написанное явно. Формулировки вроде «стандартные поля»
или «оформление по ГОСТ» остаются неразобранными и попадают в список
«проверьте»: угадывать здесь опаснее, чем спросить.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, replace
from typing import Callable

from word_formatter.core.profiles.extraction.draft import ORIGIN_TEXT, DraftValue
from word_formatter.core.profiles.extraction.sources import SourceText

# Точка в PDF часто стоит без пробела после неё («прочерк(тире).При»),
# поэтому конец предложения определяется и по заглавной букве следом.
# После цифры точка концом предложения не считается: это десятичная запятая
# или точка в нумерации.
SENTENCE_SPLIT = re.compile(r"(?<=[.!?])(?=\s|$|[«\"(А-ЯЁA-Z])|\n+")
SPACES = re.compile(r"\s+")

FONT_NAMES = (
    ("timesnewroman", "Times New Roman"),
    ("таймсньюроман", "Times New Roman"),
    ("arial", "Arial"),
    ("calibri", "Calibri"),
    ("couriernew", "Courier New"),
    ("georgia", "Georgia"),
    ("verdana", "Verdana"),
    ("gosttypea", "GOST type A"),
    ("gosttypeb", "GOST type B"),
)
FONT_WORDS = ("шрифт", "гарнитур", "набира", "оформля")

SIZE_TOKEN = re.compile(r"(\d{1,2})(?:пт|пункт|pt)")
SIZE_WORDS = ("шрифт", "размер", "кегл", "пункт")

# Категории размера шрифта: в методичках их перечисляют списком, а числа
# стоят следом в том же порядке — «основного текста, в таблицах, в сносках:
# 14, 12, 10 пунктов».
SIZE_CATEGORIES: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    ("body.sizePt", "размер шрифта основного текста", ("основноготекста", "основнойтекст", "текстаработы")),
    ("tables.fontSizePt", "размер шрифта в таблицах", ("втаблицах", "ячейкахтаблиц")),
    ("footnotes.fontSizePt", "размер шрифта сносок", ("вподстрочныхсносках", "всносках", "подстрочныхсносках")),
    ("captions.figure.sizePt", "размер шрифта подписей", ("вподписях", "подписяхрисунков")),
)

MEASURE_TOKEN = re.compile(r"(\d+(?:[.,]\d+)?)(см|мм|знак(?:ов|а|и)?)")
# Шаблоны сравниваются со «спрессованным» текстом, поэтому записаны без
# пробелов и без «й»: «верхний» там выглядит как «верхнии».
MARGIN_TOKEN = re.compile(
    r"(левое|правого|правое|правыи|верхнееинижнее|верхнииинижнии|верхнее|нижнее)"
    r"[^0-9]{0,25}(\d+(?:[.,]\d+)?)(мм|см)"
)
HEADER_FOOTER_TOKEN = re.compile(
    r"(верхнии|нижнии|верхнего|нижнего)[^0-9]{0,25}(\d+(?:[.,]\d+)?)(мм|см)"
)
# Ключи — в «спрессованном» виде: так их находит шаблон выше.
SIDE_PATHS = {
    "левое": "page.marginsMm.left",
    "правое": "page.marginsMm.right",
    "правого": "page.marginsMm.right",
    "правыи": "page.marginsMm.right",
    "верхнее": "page.marginsMm.top",
    "нижнее": "page.marginsMm.bottom",
    "верхнееинижнее": "page.marginsMm.top_bottom",
    "верхнииинижнии": "page.marginsMm.top_bottom",
}
SIDE_TITLES = {
    "page.marginsMm.left": "левое поле",
    "page.marginsMm.right": "правое поле",
    "page.marginsMm.top": "верхнее поле",
    "page.marginsMm.bottom": "нижнее поле",
}

SPACING_WORDS = (
    ("полтора", 1.5),
    ("полуторн", 1.5),
    ("одинарн", 1.0),
    ("двойн", 2.0),
)
SPACING_NUMBER = re.compile(r"(\d(?:[.,]\d)?)интервал")

# В «спрессованном» тексте слова склеены, поэтому «не» может стоять прямо
# перед словом-признаком без пробела: «заголовкиневыделяют».
NEGATION_BEFORE = re.compile(r"(не|без|запрещ\w*|исключ\w*)\w{0,15}$")
NEGATION_AFTER = re.compile(r"^.{0,60}?(недопуска|неразреша|неиспольз|неприменя|непечата|запрещ)")
# Те же проверки для настоящего текста, где пробелы на месте.
SPACED_NEGATION_BEFORE = re.compile(
    r"(?:^|\W)(?:не|без|запрещ\w*|исключ\w*)\W+(?:\w+\W+){0,3}$"
)
SPACED_NEGATION_AFTER = re.compile(
    r"^.{0,60}?\b(?:не\s+(?:допуска|разреша|использ|применя|печата)|запрещ)", re.IGNORECASE | re.DOTALL
)

ALIGNMENT_WORDS = (
    ("поширине", "justify"),
    ("поцентру", "center"),
    ("полевомукраю", "left"),
    ("полевомуполю", "left"),
    ("правомукраю", "right"),
)

# Слова, по которым понятно, что размер или интервал относится не к основному
# тексту, а к заголовкам или сноскам. Без этого одиночный интервал в сносках
# был бы принят за требование ко всему тексту.
NOT_BODY_WORDS = ("заголовк", "подзаголов", "подстрочн", "сноск", "таблиц", "подпис")


@dataclass(frozen=True)
class Chunk:
    """Кусок текста методички: спрессованный вид и настоящая цитата."""

    canon: str
    text: str
    previous: str = ""

    @property
    def evidence(self) -> str:
        return _short(self.text)

    @property
    def trailing_context(self) -> str:
        """Конец предыдущего куска: в нём бывает название того, о чём речь.

        В методичке пишут «Размер шрифта в таблице 12. Интервал — одинарный»:
        во втором предложении слово «таблица» уже не повторяется.
        """
        return self.previous[-60:]


def canon(text: str) -> str:
    """«Спрессованный» вид строки для сравнения.

    Убираются пробелы, «ё» и «й»: в PDF пробелы теряются, а «ё» и «й»
    извлекаются искажённо. Регистр не важен, поэтому всё приводится
    к нижнему.
    """
    lowered = text.lower().replace("ё", "е").replace("й", "и")
    return SPACES.sub("", lowered)


def chunks(text: str) -> list[Chunk]:
    """Делит текст на куски, сохраняя цитаты и связь с предыдущим куском."""
    parts = [part.strip() for part in SENTENCE_SPLIT.split(text) if part.strip()]
    result: list[Chunk] = []
    previous = ""
    for part in parts:
        compressed = canon(part)
        result.append(Chunk(canon=compressed, text=part, previous=previous))
        previous = compressed
    return result


def _short(text: str, limit: int = 160) -> str:
    collapsed = SPACES.sub(" ", text).strip()
    return collapsed if len(collapsed) <= limit else collapsed[: limit - 1] + "…"


def _has(chunk: Chunk, *words: str) -> bool:
    """Есть ли в куске хотя бы одно из слов.

    Сначала ищется точное совпадение, затем — устойчивое: в PDF внутри слова
    бывают лишние знаки («подчеHркивание», «одинарныи 9»).
    """
    for word in words:
        compressed = canon(word)
        if compressed in chunk.canon:
            return True
        if len(compressed) >= 6 and re.search(_loose(compressed), chunk.canon):
            return True
    return False


def _negated(chunk: Chunk, *keywords: str) -> bool:
    """Есть ли отрицание при самом слове-признаке.

    Смотреть отрицание по всему куску нельзя: во фразе «печатают прописными
    буквами, полужирным шрифтом, без абзацного отступа» слово «без» относится
    к отступу, а не к прописным буквам. Поэтому отрицание ищется рядом
    с ключевым словом — перед ним или сразу после него.

    Проверяются оба вида текста: «спрессованный» (в нём слова склеены,
    и «не выделяют» выглядит как «невыделяют») и настоящий, где пробелы
    сохранились.
    """
    for keyword in keywords:
        compressed = canon(keyword)
        for match in re.finditer(_loose(compressed), chunk.canon):
            before = chunk.canon[: match.start()]
            after = chunk.canon[match.end() :]
            if NEGATION_BEFORE.search(before) or NEGATION_AFTER.match(after):
                return True

        for match in re.finditer(keyword, chunk.text, re.IGNORECASE):
            before = chunk.text[: match.start()]
            after = chunk.text[match.end() :]
            if SPACED_NEGATION_BEFORE.search(before) or SPACED_NEGATION_AFTER.match(after):
                return True
    return False


_FUZZY_CACHE: dict[str, str] = {}


def _loose(word: str) -> str:
    """Шаблон, допускающий мусор между буквами слова.

    Нужен для PDF, где внутри слова появляются лишние знаки: «одинарныи 9»
    вместо «одинарный», «подчеHркивание» вместо «подчёркивание».
    """
    pattern = _FUZZY_CACHE.get(word)
    if pattern is None:
        pattern = "".join(re.escape(char) + r"[\w\s]{0,2}?" for char in word)
        _FUZZY_CACHE[word] = pattern
    return pattern


def _value(
    path: str,
    title: str,
    value,
    chunk: Chunk,
    evidence: str | None = None,
) -> DraftValue:
    """Значение с цитатой.

    Если кусок текста — это целая строка таблицы, цитата получилась бы
    нечитаемой. Тогда правило передаёт короткое пояснение само, например
    «левое поле — 30 мм».
    """
    return DraftValue(
        path=path,
        value=value,
        origin=ORIGIN_TEXT,
        evidence=evidence or chunk.evidence,
        title=title,
    )


def _millimetres(number: str, unit: str) -> float:
    value = float(number.replace(",", "."))
    return round(value * 10) if unit.startswith("см") else round(value, 1)


# --- отдельные требования -----------------------------------------------------


def extract_font(chunk: Chunk) -> list[DraftValue]:
    if not _has(chunk, *FONT_WORDS):
        return []
    for needle, name in FONT_NAMES:
        if needle in chunk.canon:
            return [_value("body.font", "шрифт основного текста", name, chunk)]
    return []


def extract_sizes(chunk: Chunk) -> list[DraftValue]:
    """Размеры шрифта: категории и числа идут в одном порядке."""
    numbers = [float(match.group(1)) for match in SIZE_TOKEN.finditer(chunk.canon)]
    if not numbers:
        return []

    mentioned: list[tuple[int, str, str]] = []
    for path, title, markers in SIZE_CATEGORIES:
        for marker in markers:
            position = chunk.canon.find(marker)
            if position >= 0:
                mentioned.append((position, path, title))
                break
    mentioned.sort()

    if not mentioned:
        if not _has(chunk, *SIZE_WORDS):
            return []
        return [_value("body.sizePt", "размер шрифта основного текста", numbers[0], chunk)]

    values: list[DraftValue] = []
    for index, (position, path, title) in enumerate(mentioned):
        # Числа стоят после своего названия: берём ближайшее следующее,
        # иначе — следующее по порядку.
        number = numbers[index] if index < len(numbers) else numbers[-1]
        values.append(_value(path, title, number, chunk))
    return values


def extract_spacing(chunk: Chunk) -> list[DraftValue]:
    values: list[DraftValue] = []

    # «Интервалы между абзацами 0 пт» — это расстояния до и после абзаца,
    # а не межстрочный интервал. В том же предложении методичка обычно
    # задаёт и сам межстрочный интервал, поэтому проверяем и его.
    gap = re.search(r"интервал\w*междуабзацами(\d{1,2})", chunk.canon)
    if gap is not None:
        points = float(gap.group(1))
        values.append(_value("body.spaceBeforePt", "интервал перед абзацем", points, chunk))
        values.append(_value("body.spaceAfterPt", "интервал после абзаца", points, chunk))

    if not _has(chunk, "интервал"):
        return values

    # Сначала ищем число рядом со словом «интервал»: «межстрочный интервал
    # 1,5» или «1,5 интервала». Это надёжнее словесных форм, потому что
    # в одном куске текста могут говорить и про заголовки, и про сноски.
    value = _spacing_number(chunk.canon)
    if value is None:
        for needle, candidate in SPACING_WORDS:
            if _has(chunk, needle):
                value = candidate
                break
    if value is None:
        return values

    for target in _spacing_targets(chunk):
        if target == "headings":
            values.extend(
                _value(
                    f"headings.levels.{level}.lineSpacing",
                    "интервал в заголовках",
                    value,
                    chunk,
                )
                for level in (1, 2, 3)
            )
        elif target == "tables":
            values.append(_value("tables.lineSpacing", "интервал в таблицах", value, chunk))
        elif target == "footnotes":
            values.append(_value("footnotes.lineSpacing", "интервал в сносках", value, chunk))
        else:
            values.append(_value("body.lineSpacing", "межстрочный интервал", value, chunk))
    return values


def _spacing_targets(chunk: Chunk) -> tuple[str, ...]:
    """К чему относится интервал: к тексту, заголовкам, таблицам или сноскам.

    Название обычно стоит перед словом «интервал», а в коротком предложении
    вроде «Интервал — одинарный» его приходится искать в предыдущем: там
    было «Размер шрифта в таблице 12».
    """
    if canon("межстрочный интервал") in chunk.canon:
        return ("body",)

    position = chunk.canon.find("интервал")
    before = chunk.canon[:position] if position > 0 else ""
    haystack = before if len(before) >= 20 else before + chunk.trailing_context
    if len(haystack) < 20:
        haystack = chunk.canon + chunk.trailing_context

    targets: list[str] = []
    if "заголовк" in haystack or "подзаголов" in haystack:
        targets.append("headings")
    if "таблиц" in haystack:
        targets.append("tables")
    if "подстрочн" in haystack or "сноск" in haystack:
        targets.append("footnotes")
    return tuple(targets) or ("body",)


def _spacing_number(compressed: str) -> float | None:
    """Число межстрочного интервала рядом со словом «интервал»."""
    for pattern in (r"интервал(\d(?:[.,]\d)?)", r"(\d(?:[.,]\d)?)интервал"):
        match = re.search(pattern, compressed)
        if match is None:
            continue
        candidate = float(match.group(1).replace(",", "."))
        if candidate in (1.0, 1.5, 2.0, 2.5, 3.0):
            return candidate
    return None


def extract_indent(chunk: Chunk) -> list[DraftValue]:
    if not _has(chunk, "отступ", "абзац"):
        return []
    if _has(chunk, "отступот", "отступтаблицы"):
        return []

    match = MEASURE_TOKEN.search(chunk.canon)
    if match is None:
        return []

    number, unit = match.group(1), match.group(2)
    if unit.startswith("мм"):
        centimetres = round(float(number.replace(",", ".")) / 10, 2)
    elif unit.startswith("знак"):
        # Пять знаков при 14 пт — это те же 1,25 см.
        centimetres = round(float(number.replace(",", ".")) * 0.25, 2)
    else:
        centimetres = float(number.replace(",", "."))

    if not 0 <= centimetres <= 5:
        return []
    return [_value("body.firstLineIndentCm", "абзацный отступ", centimetres, chunk)]


def extract_margins(chunk: Chunk) -> list[DraftValue]:
    values: list[DraftValue] = []
    for match in MARGIN_TOKEN.finditer(chunk.canon):
        side = match.group(1)
        millimetres = _millimetres(match.group(2), match.group(3))
        if not 5 <= millimetres <= 60:
            continue
        path = SIDE_PATHS.get(side)
        if path is None:
            continue

        if path.endswith("top_bottom"):
            for each in ("page.marginsMm.top", "page.marginsMm.bottom"):
                values.append(
                    _value(
                        each,
                        SIDE_TITLES[each],
                        millimetres,
                        chunk,
                        evidence=f"{SIDE_TITLES[each]} — {millimetres:g} мм",
                    )
                )
        else:
            values.append(
                _value(
                    path,
                    SIDE_TITLES[path],
                    millimetres,
                    chunk,
                    evidence=f"{SIDE_TITLES[path]} — {millimetres:g} мм",
                )
            )
    return values


def extract_header_footer(chunk: Chunk) -> list[DraftValue]:
    """Расстояния до колонтитулов: «верхний — 20 мм, нижний — 12,5 мм»."""
    if not _has(chunk, "колонтитул"):
        return []
    values: list[DraftValue] = []
    for match in HEADER_FOOTER_TOKEN.finditer(chunk.canon):
        side = match.group(1)
        millimetres = _millimetres(match.group(2), match.group(3))
        if not 5 <= millimetres <= 60:
            continue
        if side.startswith("верхн"):
            values.append(
                _value(
                    "page.headerMm",
                    "расстояние до верхнего колонтитула",
                    millimetres,
                    chunk,
                    evidence=f"верхний колонтитул — {millimetres:g} мм",
                )
            )
        else:
            values.append(
                _value(
                    "page.footerMm",
                    "расстояние до нижнего колонтитула",
                    millimetres,
                    chunk,
                    evidence=f"нижний колонтитул — {millimetres:g} мм",
                )
            )
    return values


def extract_alignment(chunk: Chunk) -> list[DraftValue]:
    if not _has(chunk, "выравнива", "выравнивают", "располагают"):
        return []
    for needle, value in ALIGNMENT_WORDS:
        if _has(chunk, needle):
            return [_value("body.alignment", "выравнивание основного текста", value, chunk)]
    return []


def extract_page_numbers(chunk: Chunk) -> list[DraftValue]:
    if not _has(chunk, "нумерац", "номерстраниц", "нумеруются"):
        return []
    # Заголовок «Нумерация» сам по себе требованием не является.
    if not _has(chunk, "страниц", "цифр", "сквозн"):
        return []

    values = [_value("pageNumbers.enabled", "нумерация страниц", True, chunk)]
    # Положение номера называют по-разному: «внизу по центру», «в центре
    # нижней части листа», «в правом верхнем углу».
    if _has(chunk, "внизу", "снизу", "нижнейчасти", "внижнейчасти"):
        centered = _has(chunk, "центр", "посередине")
        position = "bottom-center" if centered else "bottom-right"
        if _has(chunk, "слева"):
            position = "bottom-left"
        values.append(_value("pageNumbers.position", "положение номера страницы", position, chunk))
    elif _has(chunk, "вверху", "сверху", "верхнейчасти", "вверхнейчасти"):
        centered = _has(chunk, "центр", "посередине")
        position = "top-center" if centered else "top-right"
        if _has(chunk, "слева"):
            position = "top-left"
        values.append(_value("pageNumbers.position", "положение номера страницы", position, chunk))

    if _has(
        chunk,
        "второйстраницы",
        "второистраницы",
        "второйистраницы",
        "начинаясовторой",
        "начинаясовторои",
    ):
        values.append(_value("pageNumbers.skipFirstPage", "номер со второй страницы", True, chunk))
    if _has(chunk, "титульныйлист", "перваястраница") and _has(chunk, "ненумеру", "безномера"):
        values.append(_value("pageNumbers.skipFirstPage", "номер со второй страницы", True, chunk))
    return values


def extract_hyphenation(chunk: Chunk) -> list[DraftValue]:
    if not _has(chunk, "перенос"):
        return []
    if not _has(chunk, "автоматическ", "переносыслов", "расстановк"):
        return []

    values = [
        _value("hyphenation.enabled", "автоматические переносы", not _negated(chunk, "перенос"), chunk)
    ]
    if _has(chunk, "заисключением", "кроменазваний"):
        if _has(chunk, "разделов", "заголовк"):
            values.append(_value("hyphenation.excludeHeadings", "не переносить в заголовках", True, chunk))
        if _has(chunk, "таблиц", "рисунк", "подпис"):
            values.append(_value("hyphenation.excludeCaptions", "не переносить в подписях", True, chunk))
    return values


def extract_headings(chunk: Chunk) -> list[DraftValue]:
    values: list[DraftValue] = []
    if _has(chunk, "заголовк", "структурн", "раздел"):
        if _has(chunk, "прописн", "заглавн"):
            values.append(
                _value(
                    "headings.structural.uppercase",
                    "структурные заголовки прописными",
                    not _negated(chunk, "прописн", "заглавн"),
                    chunk,
                )
            )
        if _has(chunk, "полужирн", "жирн"):
            values.append(
                _value(
                    "headings.structural.bold",
                    "структурные заголовки полужирным",
                    not _negated(chunk, "полужирн", "жирн"),
                    chunk,
                )
            )
    return values


def extract_captions(chunk: Chunk) -> list[DraftValue]:
    values: list[DraftValue] = []
    if _has(chunk, "подпис", "рисуно", "рисунк", "иллюстрац"):
        if _has(chunk, "подрисунком", "подиллюстрацией", "подрисунками"):
            values.append(_value("captions.figure.position", "подпись рисунка", "below", chunk))
        elif _has(chunk, "надрисунком", "надиллюстрацией"):
            values.append(_value("captions.figure.position", "подпись рисунка", "above", chunk))
    if _has(chunk, "таблиц"):
        if _has(chunk, "надтаблиц", "передтаблиц"):
            values.append(_value("captions.table.position", "название таблицы", "above", chunk))
        elif _has(chunk, "подтаблиц"):
            values.append(_value("captions.table.position", "название таблицы", "below", chunk))
    return values


def extract_blank_lines(chunk: Chunk) -> list[DraftValue]:
    if _has(chunk, "свободн", "пуст") and _has(chunk, "строк", "интервал"):
        if _has(chunk, "таблиц", "рисунк", "заголовк", "формул"):
            return [_value("blankLines.insert", "свободные строки", True, chunk)]
    return []


def extract_emphasis(chunk: Chunk) -> list[DraftValue]:
    values: list[DraftValue] = []
    if _has(chunk, "курсив") and _negated(chunk, "курсив"):
        values.append(_value("body.removeItalic", "снимать курсив", True, chunk))
    if _has(chunk, "подчеркиван", "подчеркивание") and _negated(chunk, "подчеркиван"):
        values.append(_value("body.removeUnderline", "снимать подчёркивание", True, chunk))
    return values


EXTRACTORS: tuple[Callable[[Chunk], list[DraftValue]], ...] = (
    extract_font,
    extract_sizes,
    extract_spacing,
    extract_indent,
    extract_margins,
    extract_header_footer,
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
        for chunk in chunks(page):
            for extractor in EXTRACTORS:
                for item in extractor(chunk):
                    if item.path in found:
                        continue
                    if multiple_pages:
                        item = replace(item, evidence=f"{item.evidence} — стр. {page_number}")
                    found[item.path] = item

    return list(found.values())
