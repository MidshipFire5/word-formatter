"""Черновик профиля оформления.

Разные методички устроены по-разному: в одних требования написаны текстом
(«шрифт 12, Times New Roman»), другие сами оформлены по своим требованиям,
и оформление можно просто измерить. Приложение читает оба источника и
складывает найденное в черновик профиля — со ссылкой на то, откуда взято
каждое значение.

Черновик — это обычный профиль плюс сведения об источниках. Движок
форматирования работает с готовым профилем и о черновике ничего не знает.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import Any

ORIGIN_TEXT = "текст"
ORIGIN_SAMPLE = "образец"
ORIGIN_EDIT = "правка"

# Человеческие названия значений-перечислений: их видно в окне проверки,
# их же можно вписать руками вместо кода.
VALUE_CODES: dict[str, str] = {
    "по левому краю": "left",
    "по центру": "center",
    "по правому краю": "right",
    "по ширине": "justify",
    "под элементом": "below",
    "над элементом": "above",
    "под рисунком": "below",
    "над рисунком": "above",
    "не переносить": "keep",
    "внизу по центру": "bottom-center",
    "внизу слева": "bottom-left",
    "внизу справа": "bottom-right",
    "вверху по центру": "top-center",
    "вверху слева": "top-left",
    "вверху справа": "top-right",
}
TRUE_WORDS = ("да", "yes", "true", "1", "истина", "+")
FALSE_WORDS = ("нет", "no", "false", "0", "ложь", "-")

# Названия кодов для показа человеку. Одно значение — одно название:
# «above» это «над элементом», потому что для подписи рисунка и для названия
# таблицы это выглядит по-разному, а в таблице столбец общий.
CODE_TITLES: dict[str, str] = {
    "left": "по левому краю",
    "center": "по центру",
    "right": "по правому краю",
    "justify": "по ширине",
    "above": "над элементом",
    "below": "под элементом",
    "keep": "не переносить",
    "bottom-center": "внизу по центру",
    "bottom-left": "внизу слева",
    "bottom-right": "внизу справа",
    "top-center": "вверху по центру",
    "top-left": "вверху слева",
    "top-right": "вверху справа",
}

# Поля, которые приложение умеет распознавать. Всё, что не нашлось,
# попадает в список «проверьте»: молча подставлять типовые значения нельзя.
DESCRIBED_FIELDS: tuple[tuple[str, str, str], ...] = (
    ("page.marginsMm.left", "страница", "левое поле"),
    ("page.marginsMm.right", "страница", "правое поле"),
    ("page.marginsMm.top", "страница", "верхнее поле"),
    ("page.marginsMm.bottom", "страница", "нижнее поле"),
    ("page.headerMm", "страница", "расстояние до верхнего колонтитула"),
    ("page.footerMm", "страница", "расстояние до нижнего колонтитула"),
    ("body.font", "основной текст", "шрифт"),
    ("body.sizePt", "основной текст", "размер шрифта"),
    ("body.alignment", "основной текст", "выравнивание"),
    ("body.firstLineIndentCm", "основной текст", "абзацный отступ"),
    ("body.lineSpacing", "основной текст", "межстрочный интервал"),
    ("body.removeItalic", "основной текст", "снимать курсив"),
    ("body.removeUnderline", "основной текст", "снимать подчёркивание"),
    ("headings.maxLevel", "заголовки", "число уровней заголовков"),
    ("headings.structural.uppercase", "заголовки", "структурные заголовки прописными"),
    ("headings.structural.bold", "заголовки", "структурные заголовки полужирным"),
    ("captions.figure.position", "подписи", "положение подписи рисунка"),
    ("captions.table.position", "подписи", "положение названия таблицы"),
    ("tables.fontSizePt", "таблицы", "размер шрифта в таблицах"),
    ("footnotes.fontSizePt", "сноски", "размер шрифта сносок"),
    ("pageNumbers.enabled", "нумерация страниц", "нужна ли нумерация"),
    ("pageNumbers.position", "нумерация страниц", "где стоит номер"),
    ("pageNumbers.skipFirstPage", "нумерация страниц", "номер со второй страницы"),
    ("hyphenation.enabled", "переносы", "автоматические переносы"),
    ("blankLines.insert", "свободные строки", "вставлять свободные строки"),
)

# Основа черновика: полный профиль с самыми распространёнными значениями.
# Всё, что не подтвердилось методичкой, перечисляется в списке «проверьте».
BASE_PROFILE: dict[str, Any] = {
    "id": "draft",
    "name": "Черновик профиля",
    "description": "Профиль, собранный приложением из методички.",
    "version": "1.0.0",
    "demo": False,
    "source": {"title": "Черновик профиля из методички"},
    "appliesTo": ["черновик"],
    "page": {
        "format": "A4",
        "orientation": "portrait",
        "marginsMm": {"left": 30, "right": 15, "top": 20, "bottom": 20},
        "headerMm": 20,
        "footerMm": 12.5,
    },
    "body": {
        "font": "Times New Roman",
        "sizePt": 14,
        "alignment": "justify",
        "firstLineIndentCm": 1.25,
        "lineSpacing": 1.5,
        "spaceBeforePt": 0,
        "spaceAfterPt": 0,
        "bold": False,
        "italic": False,
        "color": None,
        "removeItalic": False,
        "removeUnderline": False,
    },
    "headings": {
        "maxLevel": 3,
        "noHyphenation": True,
        "structural": {
            "titles": [
                "ВВЕДЕНИЕ",
                "ЗАКЛЮЧЕНИЕ",
                "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                "СОДЕРЖАНИЕ",
                "ПРИЛОЖЕНИЕ",
            ],
            "uppercase": True,
            "font": "Times New Roman",
            "sizePt": 14,
            "bold": True,
            "italic": False,
            "color": None,
            "alignment": "center",
            "firstLineIndentCm": 0,
            "lineSpacing": 1.0,
            "spaceBeforePt": 0,
            "spaceAfterPt": 0,
        },
        "levels": {
            str(level): {
                "font": "Times New Roman",
                "sizePt": 14,
                "bold": True,
                "italic": False,
                "color": None,
                "alignment": "justify",
                "firstLineIndentCm": 1.25,
                "lineSpacing": 1.0,
                "spaceBeforePt": 0,
                "spaceAfterPt": 0,
            }
            for level in (1, 2, 3)
        },
    },
    "captions": {
        "figure": {
            "font": "Times New Roman",
            "sizePt": 14,
            "alignment": "center",
            "firstLineIndentCm": 0,
            "lineSpacing": 1.0,
            "spaceBeforePt": 0,
            "spaceAfterPt": 0,
            "noHyphenation": True,
            "position": "keep",
        },
        "table": {
            "font": "Times New Roman",
            "sizePt": 14,
            "alignment": "left",
            "firstLineIndentCm": 0,
            "lineSpacing": 1.0,
            "spaceBeforePt": 0,
            "spaceAfterPt": 0,
            "noHyphenation": True,
            "position": "keep",
        },
    },
    "figures": {
        "alignment": "center",
        "lineSpacing": 1.0,
        "spaceBeforePt": 0,
        "spaceAfterPt": 0,
    },
    "tables": {
        "font": "Times New Roman",
        "fontSizePt": 12,
        "lineSpacing": 1.0,
        "firstLineIndentCm": 0,
        "headerRowAlign": "center",
        "numericCellAlign": "center",
        "textCellAlign": "left",
        "markerCells": ["", "–", "-", "—", "х", "x", "Х", "X"],
        "preferredWidth": "auto",
        "indentCm": 0,
    },
    "footnotes": {"font": "Times New Roman", "fontSizePt": 10, "lineSpacing": 1.0},
    "lists": {
        "applyParagraphFormatting": True,
        "replaceMarkers": False,
        "firstLineIndentCm": 1.25,
    },
    # Нумерация, переносы и свободные строки по умолчанию выключены:
    # лучше не добавить лишнего, чем добавить то, чего методичка не просила.
    "pageNumbers": {
        "enabled": False,
        "position": "bottom-center",
        "font": "Times New Roman",
        "fontSizePt": 12,
        "skipFirstPage": True,
    },
    "hyphenation": {"enabled": False, "excludeHeadings": True, "excludeCaptions": True},
    "blankLines": {"insert": False, "style": "body"},
    "limitations": [],
}


@dataclass(frozen=True)
class DraftValue:
    """Одно найденное значение и его источник."""

    path: str
    value: Any
    origin: str
    evidence: str
    title: str = ""

    def describe(self) -> str:
        return f"{self.title or self.path}: {self.value}"


@dataclass(frozen=True)
class DraftConflict:
    """Текст методички и её оформление говорят разное."""

    path: str
    title: str
    from_text: Any
    from_sample: Any

    def describe(self) -> str:
        return (
            f"{self.title}: в тексте {self.from_text}, "
            f"в оформлении файла {self.from_sample} — взято из текста"
        )


@dataclass
class ProfileDraft:
    """Черновик профиля: что найдено, что не найдено и где расхождения."""

    values: list[DraftValue] = field(default_factory=list)
    conflicts: list[DraftConflict] = field(default_factory=list)
    text_excerpt: str = ""
    sample_measured: bool = False

    def value_for(self, path: str) -> DraftValue | None:
        for item in self.values:
            if item.path == path:
                return item
        return None

    @property
    def found_paths(self) -> set[str]:
        return {item.path for item in self.values}

    @property
    def missing(self) -> list[tuple[str, str]]:
        """Что приложение не нашло — эти значения остались типовыми."""
        return [
            (section, title)
            for path, section, title in DESCRIBED_FIELDS
            if path not in self.found_paths
        ]

    def value_list(self) -> list[DraftValue]:
        return sorted(self.values, key=lambda item: item.path)

    def editable_rows(self) -> list["DraftRow"]:
        """Строки для окна проверки: найденное и то, что осталось найти человеку.

        Ненайденные требования показываются с пустым значением: типовое
        значение подставляется только при сохранении, если человек ничего
        не вписал.
        """
        rows = [
            DraftRow(
                path=item.path,
                title=item.title or item.path,
                value=item.value,
                origin=item.origin,
                evidence=item.evidence,
                found=True,
            )
            for item in self.value_list()
        ]

        found = self.found_paths
        rows.extend(
            DraftRow(
                path=path,
                title=title,
                value=None,
                origin="",
                evidence="",
                found=False,
            )
            for path, _section, title in DESCRIBED_FIELDS
            if path not in found
        )
        return rows

    def apply_edits(self, edits: dict[str, Any]) -> int:
        """Применяет значения, исправленные человеком в окне проверки.

        Возвращает число изменённых требований. Пустое значение означает
        «оставить как было»: типовое значение из основы профиля.
        """
        changed = 0
        for path, value in edits.items():
            if value is None:
                continue
            existing = self.value_for(path)
            if existing is not None:
                if existing.value == value:
                    continue
                self.values = [item for item in self.values if item.path != path]
            self.values.append(
                DraftValue(
                    path=path,
                    value=value,
                    origin=ORIGIN_EDIT,
                    evidence="исправлено вами при проверке профиля",
                    title=existing.title if existing is not None else _title_for(path),
                )
            )
            changed += 1
        return changed

    def to_profile(self, *, profile_id: str, name: str, description: str = "") -> dict[str, Any]:
        """Собирает из черновика полноценный профиль."""
        profile = copy.deepcopy(BASE_PROFILE)
        profile["id"] = profile_id
        profile["name"] = name
        profile["description"] = description or f"Профиль, собранный из методички: {name}"
        profile["source"] = {
            "title": name,
            "note": "Профиль собран приложением из методички. Проверьте значения перед использованием.",
        }
        profile["appliesTo"] = ["работа, оформленная по своей методичке"]

        for item in self.values:
            _set(profile, item.path, item.value)

        limitations = list(profile.get("limitations", ()))
        for section, title in self.missing:
            limitations.append(
                f"В методичке не нашлось: {title} ({section}) — оставлено типовое значение."
            )
        profile["limitations"] = limitations
        return profile

    def summary(self) -> str:
        """Короткий отчёт о разборе методички — для окна и журнала."""
        lines = [
            f"Найдено значений: {len(self.values)}",
        ]
        if self.conflicts:
            lines.append(f"Расхождений текста и оформления: {len(self.conflicts)}")
        if self.missing:
            lines.append(f"Не найдено (осталось типовым): {len(self.missing)}")
        return "\n".join(lines)


@dataclass(frozen=True)
class DraftRow:
    """Строка окна проверки: требование, значение и его источник."""

    path: str
    title: str
    value: Any
    origin: str
    evidence: str
    found: bool

    @property
    def placeholder(self) -> str:
        """Подсказка для пустого поля: что приложение поставит по умолчанию."""
        if self.found:
            return ""
        return f"не найдено — будет типовое значение ({_base_value(self.path)})"


def _title_for(path: str) -> str:
    for known_path, _section, title in DESCRIBED_FIELDS:
        if known_path == path:
            return title
    return path


def _base_value(path: str) -> str:
    node: Any = BASE_PROFILE
    for part in path.split("."):
        if not isinstance(node, dict) or part not in node:
            return "—"
        node = node[part]
    return str(node)


def coerce_value(path: str, text: str) -> Any:
    """Превращает вписанный человеком текст в значение нужного типа.

    Пустая строка означает «оставить типовое значение» и возвращает None.
    """
    cleaned = text.strip()
    if not cleaned:
        return None

    lowered = cleaned.lower()
    if lowered in TRUE_WORDS:
        return True
    if lowered in FALSE_WORDS:
        return False
    if lowered in VALUE_CODES:
        return VALUE_CODES[lowered]

    base = _base_value_object(path)
    if isinstance(base, str):
        return cleaned

    number = cleaned.replace(",", ".").replace(" ", "").replace("пт", "").replace("мм", "")
    try:
        return float(number)
    except ValueError:
        return None


def _base_value_object(path: str) -> Any:
    node: Any = BASE_PROFILE
    for part in path.split("."):
        if not isinstance(node, dict) or part not in node:
            return None
        node = node[part]
    return node


# Подсказки для пустых полей: серым текстом в самой ячейке. Человек должен
# видеть, что именно здесь можно вписать, иначе он напишет «полтора», а код
# ждёт число.
HINT_YES_NO = "да или нет"
HINT_SPACING = "число: 1 — одинарный, 1,5 — полуторный, 2 — двойной"
HINT_ALIGNMENT = "по ширине, по центру, по левому краю или по правому краю"
HINT_POSITION = "над элементом или под элементом"
HINT_PAGE_NUMBER_POSITION = (
    "внизу по центру, внизу слева, внизу справа, "
    "вверху по центру, вверху слева или вверху справа"
)
HINT_FONT = "название шрифта, например Times New Roman"
HINT_TEXT = "текст"


def value_hint(path: str) -> str:
    """Что можно вписать в это поле — серая подсказка в пустой ячейке."""
    base = _base_value_object(path)

    if isinstance(base, bool):
        return HINT_YES_NO
    if path.endswith("lineSpacing"):
        return HINT_SPACING
    if path.endswith("alignment"):
        return HINT_ALIGNMENT
    if path == "pageNumbers.position":
        return HINT_PAGE_NUMBER_POSITION
    if path.endswith("position"):
        return HINT_POSITION
    if path.endswith("font"):
        return HINT_FONT
    if isinstance(base, (int, float)):
        # Единица измерения видна в самом имени поля: sizePt, marginsMm,
        # firstLineIndentCm, spaceBeforePt.
        if "Pt" in path:
            return "число в пунктах, например 14"
        if "Cm" in path:
            return "число в сантиметрах, например 1,25"
        if "Mm" in path:
            return "число в миллиметрах, например 20"
        return "целое число, например 3"
    if isinstance(base, str):
        return HINT_TEXT
    return HINT_TEXT


def _set(profile: dict[str, Any], path: str, value: Any) -> None:
    """Записывает значение по пути вида «body.sizePt»."""
    parts = path.split(".")
    node: Any = profile
    for part in parts[:-1]:
        if not isinstance(node, dict) or part not in node:
            return
        node = node[part]
    if isinstance(node, dict):
        node[parts[-1]] = value
