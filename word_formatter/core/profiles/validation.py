"""Проверка файла профиля.

Задача модуля — объяснить по-русски, что именно не так в файле профиля,
и назвать место ошибки. Проверка строгая: неизвестное поле считается
ошибкой, потому что опечатка в имени поля молча привела бы к неверному
оформлению документа.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

ALIGNMENTS = ("left", "center", "right", "justify")
ORIENTATIONS = ("portrait", "landscape")
PAGE_FORMATS = ("A4", "A5", "A3", "Letter", "Legal")
PAGE_NUMBER_POSITIONS = (
    "bottom-center",
    "bottom-left",
    "bottom-right",
    "top-center",
    "top-left",
    "top-right",
)
BLANK_LINE_STYLES = ("body", "single")

ID_PATTERN = re.compile(r"^[a-z0-9][a-z0-9._-]{1,63}$")
VERSION_PATTERN = re.compile(r"^\d+\.\d+(\.\d+)?$")
COLOR_PATTERN = re.compile(r"^#[0-9A-Fa-f]{6}$")

ROOT_KEYS = {
    "id", "name", "description", "version", "demo", "source", "appliesTo",
    "page", "body", "headings", "captions", "tables", "footnotes", "lists",
    "pageNumbers", "hyphenation", "blankLines", "limitations",
}
SOURCE_KEYS = {"title", "authors", "place", "year", "pages", "note"}
PAGE_KEYS = {"format", "orientation", "marginsMm", "headerMm", "footerMm"}
MARGIN_KEYS = ("left", "right", "top", "bottom")
TEXT_STYLE_KEYS = {
    "font", "sizePt", "bold", "italic", "color", "alignment",
    "firstLineIndentCm", "lineSpacing", "spaceBeforePt", "spaceAfterPt",
}
BODY_KEYS = TEXT_STYLE_KEYS | {"removeItalic", "removeUnderline"}
HEADINGS_KEYS = {"maxLevel", "noHyphenation", "structural", "levels"}
STRUCTURAL_KEYS = TEXT_STYLE_KEYS | {"uppercase", "titles"}
CAPTIONS_KEYS = {"figure", "table"}
CAPTION_KEYS = TEXT_STYLE_KEYS | {"noHyphenation", "position"}
CAPTION_POSITIONS = ("keep", "above", "below")
TABLES_KEYS = {
    "font", "fontSizePt", "lineSpacing", "firstLineIndentCm",
    "headerRowAlign", "numericCellAlign", "textCellAlign", "markerCells",
}
FOOTNOTES_KEYS = {"font", "fontSizePt", "lineSpacing"}
LISTS_KEYS = {"applyParagraphFormatting", "replaceMarkers", "firstLineIndentCm"}
PAGE_NUMBERS_KEYS = {"enabled", "position", "font", "fontSizePt", "skipFirstPage"}
HYPHENATION_KEYS = {"enabled", "excludeHeadings", "excludeCaptions"}
BLANK_LINES_KEYS = {"insert", "style"}


@dataclass(frozen=True)
class Issue:
    """Одна найденная проблема: где и что не так."""

    path: str
    message: str

    def __str__(self) -> str:
        return f"{self.path}: {self.message}" if self.path else self.message


def type_name(value: Any) -> str:
    if value is None:
        return "пустое значение (null)"
    if isinstance(value, bool):
        return "логическое значение true/false"
    if isinstance(value, (int, float)):
        return "число"
    if isinstance(value, str):
        return "строка"
    if isinstance(value, list):
        return "список"
    if isinstance(value, dict):
        return "объект"
    return type(value).__name__


class Checker:
    """Набор проверок с накоплением найденных проблем."""

    def __init__(self) -> None:
        self.issues: list[Issue] = []

    def add(self, path: str, message: str) -> None:
        self.issues.append(Issue(path, message))

    @staticmethod
    def _join(path: str, key: str) -> str:
        return f"{path}.{key}" if path else key

    def object(self, value: Any, path: str, *, required: bool = True) -> dict[str, Any] | None:
        if value is None:
            if required:
                self.add(path, "не заполнено обязательное поле")
            return None
        if not isinstance(value, dict):
            self.add(path, f"должно быть объектом, а получено: {type_name(value)}")
            return None
        return value

    def keys(self, obj: dict[str, Any], path: str, allowed: set[str]) -> None:
        unknown = sorted(set(obj) - allowed)
        for key in unknown:
            self.add(
                self._join(path, key),
                "неизвестное поле — возможно, опечатка. Допустимые поля: "
                + ", ".join(sorted(allowed)),
            )

    def text(
        self,
        obj: dict[str, Any],
        key: str,
        path: str,
        *,
        required: bool = True,
        choices: tuple[str, ...] | None = None,
        pattern: re.Pattern[str] | None = None,
        pattern_hint: str = "",
        allow_null: bool = False,
    ) -> str | None:
        full = self._join(path, key)
        if key not in obj:
            if required:
                self.add(full, "не заполнено обязательное поле")
            return None
        value = obj[key]
        if value is None:
            if allow_null:
                return None
            self.add(full, "не может быть пустым")
            return None
        if not isinstance(value, str):
            self.add(full, f"должно быть строкой, а получено: {type_name(value)}")
            return None
        if choices is not None and value not in choices:
            self.add(full, f"может быть только одним из: {', '.join(choices)}. Получено: «{value}»")
            return None
        if pattern is not None and not pattern.match(value):
            hint = f" (образец: {pattern_hint})" if pattern_hint else ""
            self.add(full, f"записано неверно: «{value}»{hint}")
            return None
        return value

    def number(
        self,
        obj: dict[str, Any],
        key: str,
        path: str,
        *,
        required: bool = True,
        minimum: float | None = None,
        maximum: float | None = None,
        unit_text: str = "",
        integer: bool = False,
    ) -> float | None:
        full = self._join(path, key)
        if key not in obj:
            if required:
                self.add(full, "не заполнено обязательное поле")
            return None
        value = obj[key]
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            unit = f" {unit_text}" if unit_text else ""
            self.add(full, f"должно быть числом{unit}, а получено: {type_name(value)}")
            return None
        if integer and float(value) != int(value):
            self.add(full, f"должно быть целым числом, а получено: {value}")
            return None
        if minimum is not None and value < minimum:
            self.add(full, f"должно быть не меньше {minimum}. Получено: {value}")
            return None
        if maximum is not None and value > maximum:
            self.add(full, f"должно быть не больше {maximum}. Получено: {value}")
            return None
        return float(value)

    def boolean(self, obj: dict[str, Any], key: str, path: str, *, required: bool = False) -> bool | None:
        full = self._join(path, key)
        if key not in obj:
            if required:
                self.add(full, "не заполнено обязательное поле")
            return None
        value = obj[key]
        if not isinstance(value, bool):
            self.add(full, f"должно быть true или false, а получено: {type_name(value)}")
            return None
        return value

    def string_list(
        self,
        obj: dict[str, Any],
        key: str,
        path: str,
        *,
        required: bool = True,
        allow_empty_items: bool = False,
        min_items: int = 0,
    ) -> list[str] | None:
        full = self._join(path, key)
        if key not in obj:
            if required:
                self.add(full, "не заполнено обязательное поле")
            return None
        value = obj[key]
        if not isinstance(value, list):
            self.add(full, f"должно быть списком строк, а получено: {type_name(value)}")
            return None
        problems = [
            index
            for index, item in enumerate(value)
            if not isinstance(item, str) or (not allow_empty_items and not item.strip())
        ]
        if problems:
            self.add(full, f"элементы с номерами {problems} должны быть непустыми строками")
            return None
        if len(value) < min_items:
            self.add(full, f"должно содержать не меньше {min_items} элементов")
            return None
        return list(value)

    def number_map(
        self,
        obj: dict[str, Any],
        key: str,
        path: str,
        *,
        expected_keys: tuple[str, ...],
        minimum: float | None = None,
        maximum: float | None = None,
        unit_text: str = "",
    ) -> dict[str, float] | None:
        full = self._join(path, key)
        if key not in obj:
            self.add(full, "не заполнено обязательное поле")
            return None
        value = obj[key]
        if not isinstance(value, dict):
            self.add(full, f"должно быть объектом с полями {', '.join(expected_keys)}, а получено: {type_name(value)}")
            return None
        self.keys(value, full, set(expected_keys))
        result: dict[str, float] = {}
        for item_key in expected_keys:
            number = self.number(value, item_key, full, minimum=minimum, maximum=maximum, unit_text=unit_text)
            if number is not None:
                result[item_key] = number
        return result


def _validate_text_style(checker: Checker, obj: dict[str, Any], path: str, *, allowed: set[str]) -> None:
    checker.keys(obj, path, allowed)
    checker.text(obj, "font", path, required=False, allow_null=True)
    checker.number(obj, "sizePt", path, minimum=6, maximum=72, unit_text="в пунктах")
    checker.boolean(obj, "bold", path)
    checker.boolean(obj, "italic", path)
    checker.text(obj, "color", path, required=False, allow_null=True, pattern=COLOR_PATTERN, pattern_hint="#1A2B3C")
    checker.text(obj, "alignment", path, choices=ALIGNMENTS)
    checker.number(obj, "firstLineIndentCm", path, minimum=0, maximum=10, unit_text="в сантиметрах")
    checker.number(obj, "lineSpacing", path, minimum=0.5, maximum=3, unit_text="(множитель интервала)")
    checker.number(obj, "spaceBeforePt", path, minimum=0, maximum=200, unit_text="в пунктах")
    checker.number(obj, "spaceAfterPt", path, minimum=0, maximum=200, unit_text="в пунктах")


def _validate_source(checker: Checker, raw: dict[str, Any]) -> None:
    source = checker.object(raw.get("source"), "source", required=False)
    if source is None:
        return
    checker.keys(source, "source", SOURCE_KEYS)
    checker.text(source, "title", "source", required=True)
    checker.text(source, "authors", "source", required=False)
    checker.text(source, "place", "source", required=False)
    checker.number(source, "year", "source", required=False, minimum=1900, maximum=2100, integer=True)
    checker.number(source, "pages", "source", required=False, minimum=1, integer=True)
    checker.text(source, "note", "source", required=False)


def _validate_page(checker: Checker, raw: dict[str, Any]) -> None:
    page = checker.object(raw.get("page"), "page")
    if page is None:
        return
    checker.keys(page, "page", PAGE_KEYS)
    checker.text(page, "format", "page", choices=PAGE_FORMATS)
    checker.text(page, "orientation", "page", choices=ORIENTATIONS)
    checker.number_map(page, "marginsMm", "page", expected_keys=MARGIN_KEYS, minimum=0, maximum=100, unit_text="в миллиметрах")
    checker.number(page, "headerMm", "page", minimum=0, maximum=100, unit_text="в миллиметрах")
    checker.number(page, "footerMm", "page", minimum=0, maximum=100, unit_text="в миллиметрах")


def _validate_body(checker: Checker, raw: dict[str, Any]) -> None:
    body = checker.object(raw.get("body"), "body")
    if body is None:
        return
    _validate_text_style(checker, body, "body", allowed=BODY_KEYS)
    checker.boolean(body, "removeItalic", "body")
    checker.boolean(body, "removeUnderline", "body")


def _validate_headings(checker: Checker, raw: dict[str, Any]) -> None:
    headings = checker.object(raw.get("headings"), "headings")
    if headings is None:
        return
    checker.keys(headings, "headings", HEADINGS_KEYS)
    max_level = checker.number(headings, "maxLevel", "headings", minimum=1, maximum=9, integer=True)
    checker.boolean(headings, "noHyphenation", "headings")

    structural = checker.object(headings.get("structural"), "headings.structural")
    if structural is not None:
        _validate_text_style(checker, structural, "headings.structural", allowed=STRUCTURAL_KEYS)
        checker.boolean(structural, "uppercase", "headings.structural")
        checker.string_list(structural, "titles", "headings.structural", min_items=1)

    levels = checker.object(headings.get("levels"), "headings.levels")
    if levels is None:
        return
    expected = tuple(str(level) for level in range(1, int(max_level or 1) + 1))
    checker.keys(levels, "headings.levels", set(expected))
    for key in expected:
        level_obj = checker.object(levels.get(key), f"headings.levels.{key}")
        if level_obj is not None:
            _validate_text_style(checker, level_obj, f"headings.levels.{key}", allowed=TEXT_STYLE_KEYS)
    for key in sorted(set(levels) - set(expected)):
        checker.add(
            f"headings.levels.{key}",
            f"уровень не описан профилем: допустимы уровни {', '.join(expected)}",
        )


def _validate_captions(checker: Checker, raw: dict[str, Any]) -> None:
    captions = checker.object(raw.get("captions"), "captions")
    if captions is None:
        return
    checker.keys(captions, "captions", CAPTIONS_KEYS)
    for key in sorted(CAPTIONS_KEYS):
        caption = checker.object(captions.get(key), f"captions.{key}")
        if caption is not None:
            _validate_text_style(checker, caption, f"captions.{key}", allowed=CAPTION_KEYS)
            checker.boolean(caption, "noHyphenation", f"captions.{key}")
            checker.text(caption, "position", f"captions.{key}", required=False, choices=CAPTION_POSITIONS)


def _validate_tables(checker: Checker, raw: dict[str, Any]) -> None:
    tables = checker.object(raw.get("tables"), "tables")
    if tables is None:
        return
    checker.keys(tables, "tables", TABLES_KEYS)
    checker.text(tables, "font", "tables", required=False, allow_null=True)
    checker.number(tables, "fontSizePt", "tables", minimum=6, maximum=72, unit_text="в пунктах")
    checker.number(tables, "lineSpacing", "tables", minimum=0.5, maximum=3, unit_text="(множитель интервала)")
    checker.number(tables, "firstLineIndentCm", "tables", minimum=0, maximum=10, unit_text="в сантиметрах")
    checker.text(tables, "headerRowAlign", "tables", choices=ALIGNMENTS)
    checker.text(tables, "numericCellAlign", "tables", choices=ALIGNMENTS)
    checker.text(tables, "textCellAlign", "tables", choices=ALIGNMENTS)
    checker.string_list(tables, "markerCells", "tables", required=False, allow_empty_items=True)


def _validate_footnotes(checker: Checker, raw: dict[str, Any]) -> None:
    footnotes = checker.object(raw.get("footnotes"), "footnotes")
    if footnotes is None:
        return
    checker.keys(footnotes, "footnotes", FOOTNOTES_KEYS)
    checker.text(footnotes, "font", "footnotes", required=False, allow_null=True)
    checker.number(footnotes, "fontSizePt", "footnotes", minimum=6, maximum=72, unit_text="в пунктах")
    checker.number(footnotes, "lineSpacing", "footnotes", minimum=0.5, maximum=3, unit_text="(множитель интервала)")


def _validate_lists(checker: Checker, raw: dict[str, Any]) -> None:
    lists = checker.object(raw.get("lists"), "lists")
    if lists is None:
        return
    checker.keys(lists, "lists", LISTS_KEYS)
    checker.boolean(lists, "applyParagraphFormatting", "lists", required=True)
    checker.boolean(lists, "replaceMarkers", "lists", required=True)
    checker.number(lists, "firstLineIndentCm", "lists", minimum=0, maximum=10, unit_text="в сантиметрах")


def _validate_page_numbers(checker: Checker, raw: dict[str, Any]) -> None:
    numbers = checker.object(raw.get("pageNumbers"), "pageNumbers")
    if numbers is None:
        return
    checker.keys(numbers, "pageNumbers", PAGE_NUMBERS_KEYS)
    checker.boolean(numbers, "enabled", "pageNumbers", required=True)
    checker.text(numbers, "position", "pageNumbers", choices=PAGE_NUMBER_POSITIONS)
    checker.text(numbers, "font", "pageNumbers", required=False, allow_null=True)
    checker.number(numbers, "fontSizePt", "pageNumbers", minimum=6, maximum=72, unit_text="в пунктах")
    checker.boolean(numbers, "skipFirstPage", "pageNumbers", required=True)


def _validate_hyphenation(checker: Checker, raw: dict[str, Any]) -> None:
    hyphenation = checker.object(raw.get("hyphenation"), "hyphenation")
    if hyphenation is None:
        return
    checker.keys(hyphenation, "hyphenation", HYPHENATION_KEYS)
    checker.boolean(hyphenation, "enabled", "hyphenation", required=True)
    checker.boolean(hyphenation, "excludeHeadings", "hyphenation", required=True)
    checker.boolean(hyphenation, "excludeCaptions", "hyphenation", required=True)


def _validate_blank_lines(checker: Checker, raw: dict[str, Any]) -> None:
    blank_lines = checker.object(raw.get("blankLines"), "blankLines")
    if blank_lines is None:
        return
    checker.keys(blank_lines, "blankLines", BLANK_LINES_KEYS)
    checker.boolean(blank_lines, "insert", "blankLines", required=True)
    checker.text(blank_lines, "style", "blankLines", choices=BLANK_LINE_STYLES)


def validate_profile(raw: Any) -> list[Issue]:
    """Проверяет содержимое профиля и возвращает список найденных проблем.

    Пустой список означает, что профиль пригоден к использованию.
    """
    checker = Checker()
    profile = checker.object(raw, "")
    if profile is None:
        return checker.issues

    checker.keys(profile, "", ROOT_KEYS)
    checker.text(profile, "id", "", pattern=ID_PATTERN, pattern_hint="fu-krasnodar-2026")
    checker.text(profile, "name", "")
    checker.text(profile, "description", "", required=False)
    checker.text(profile, "version", "", pattern=VERSION_PATTERN, pattern_hint="1.0.0")
    checker.boolean(profile, "demo", "")
    _validate_source(checker, profile)
    checker.string_list(profile, "appliesTo", "", min_items=1)
    _validate_page(checker, profile)
    _validate_body(checker, profile)
    _validate_headings(checker, profile)
    _validate_captions(checker, profile)
    _validate_tables(checker, profile)
    _validate_footnotes(checker, profile)
    _validate_lists(checker, profile)
    _validate_page_numbers(checker, profile)
    _validate_hyphenation(checker, profile)
    _validate_blank_lines(checker, profile)
    checker.string_list(profile, "limitations", "", required=False)

    return checker.issues
