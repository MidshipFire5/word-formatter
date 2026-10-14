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
