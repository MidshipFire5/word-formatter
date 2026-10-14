"""Вспомогательные функции для тестов: разбор DOCX-пакета и сравнение содержимого.

Здесь нет логики приложения: это независимый «измерительный прибор»,
которым тесты проверяют, что обработка документа ничего не потеряла.
"""

from __future__ import annotations

import hashlib
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

W = f"{{{NS_W}}}"
R = f"{{{NS_R}}}"


def read_parts(path: str | Path) -> dict[str, bytes]:
    """Возвращает содержимое всех частей пакета DOCX."""
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def xml_bytes(data: bytes) -> etree._Element:
    return etree.fromstring(data)


def texts(data: bytes) -> list[str]:
    """Весь текст документа: значения всех элементов w:t в порядке следования."""
    root = xml_bytes(data)
    return [node.text or "" for node in root.iter(f"{W}t")]


def count(data: bytes, tag: str) -> int:
    """Число элементов с указанным локальным именем (namespace WordprocessingML)."""
    root = xml_bytes(data)
    return len(list(root.iter(f"{W}{tag}")))


def count_attr(data: bytes, tag: str, attr: str, value: str) -> int:
    root = xml_bytes(data)
    return sum(1 for node in root.iter(f"{W}{tag}") if node.get(f"{W}{attr}") == value)


def relationships(data: bytes) -> dict[str, tuple[str, str]]:
    """Разрешает таблицу связей части: id -> (тип, цель)."""
    root = xml_bytes(data)
    result: dict[str, tuple[str, str]] = {}
    for node in root.iter(f"{{{NS_REL}}}Relationship"):
        rid = node.get("Id") or ""
        result[rid] = (node.get("Type") or "", node.get("Target") or "")
    return result


def stylesheet_ids(data: bytes) -> list[str]:
    root = xml_bytes(data)
    return [node.get(f"{W}styleId") or "" for node in root.iter(f"{W}style")]


def structure(data: bytes) -> list[tuple[str, tuple[tuple[str, str], ...]]]:
    """Состав XML: последовательность «элемент + его атрибуты», без текста и отступов.

    Сравнение таких последовательностей показывает, потерян или добавлен
    хоть один элемент, независимо от форматирования файла.
    """
    root = xml_bytes(data)
    return [(node.tag, tuple(sorted(node.attrib.items()))) for node in root.iter()]


def structure_diff(before: bytes, after: bytes) -> tuple[list, list]:
    """Возвращает (потеряно, добавлено) с учётом количества повторов элементов."""
    old = Counter(structure(before))
    new = Counter(structure(after))
    return list((old - new).elements()), list((new - old).elements())


def content_type_maps(data: bytes) -> tuple[dict[str, str], dict[str, str]]:
    """Разбирает [Content_Types].xml на правила по расширению и по части."""
    root = xml_bytes(data)
    ns = "{http://schemas.openxmlformats.org/package/2006/content-types}"
    defaults = {node.get("Extension") or "": node.get("ContentType") or "" for node in root.iter(f"{ns}Default")}
    overrides = {node.get("PartName") or "": node.get("ContentType") or "" for node in root.iter(f"{ns}Override")}
    return defaults, overrides


def effective_content_types(data: bytes, part_names) -> dict[str, str]:
    """Действующий тип содержимого каждой части — то, что реально видит Word."""
    defaults, overrides = content_type_maps(data)
    result: dict[str, str] = {}
    for name in part_names:
        override = overrides.get("/" + name)
        if override:
            result[name] = override
        else:
            extension = name.rsplit(".", 1)[-1]
            result[name] = defaults.get(extension, "")
    return result


# Элементы OOXML, которые описывают оформление: абзац, прогон текста и раздел.
# Приложение вправе их менять. Всё остальное обязано остаться нетронутым —
# это главный контракт приложения «меняем оформление, а не содержание».
FORMATTING_ELEMENTS = frozenset(
    {
        # абзац
        "pPr", "pStyle", "jc", "ind", "spacing", "contextualSpacing", "keepNext",
        "keepLines", "widowControl", "suppressAutoHyphenation", "outlineLvl", "tabs",
        "tab", "snapToGrid", "adjustRightInd", "autoSpaceDE", "autoSpaceDN",
        "textAlignment", "pBdr", "bdr", "shd", "divId", "numPr", "pPrChange",
        # прогон текста
        "rPr", "rStyle", "sz", "szCs", "rFonts", "b", "bCs", "i", "iCs", "u", "color",
        "kern", "position", "vertAlign", "smallCaps", "caps", "strike", "dstrike",
        "lang", "noProof", "rPrChange",
        # раздел и параметры страницы
        "sectPr", "pgSz", "pgMar", "cols", "docGrid", "headerReference",
        "footerReference", "titlePg", "pgNumType", "lnNumType", "textDirection",
        "rtlGutter", "type", "header", "footer", "gutter", "footnotePr", "endnotePr",
        # таблица: её положение и ширина, но не содержимое ячеек
        "tblPr", "tblW", "tblInd", "tblLayout",
    }
)


def non_formatting_difference(before: bytes, after: bytes, *, allowed_additions: frozenset = frozenset()) -> list:
    """Элементы, которые изменились, хотя описывают не оформление.

    Пустой список означает, что правка не вышла за пределы оформления.
    Через allowed_additions можно разрешить добавление отдельных элементов,
    например пустых абзацев для свободных строк; содержательность таких
    добавлений отдельно проверяется сравнением текста.
    """
    lost, added = structure_diff(before, after)

    def local_name(item: tuple[str, tuple[tuple[str, str], ...]]) -> str:
        return item[0].rsplit("}", 1)[-1]

    problems = [item for item in lost if local_name(item) not in FORMATTING_ELEMENTS]
    problems += [
        item
        for item in added
        if local_name(item) not in FORMATTING_ELEMENTS and local_name(item) not in allowed_additions
    ]
    return problems
