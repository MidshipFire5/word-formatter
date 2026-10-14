"""Положение и ширина таблицы.

Таблица «уезжает вправо» по двум причинам:

- у неё задана ширина (в Word это галочка «Ширина» на вкладке «Таблица»
  в свойствах таблицы) — тогда таблица может вылезать за правое поле;
- задан отступ от левого поля (`w:tblInd`) — тогда таблица сдвинута вправо
  на всю величину отступа.

Оба значения приходят из исходного документа и оформления к требованиям
методички не имеют, поэтому профиль может их снять.
"""

from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

from word_formatter.core.profiles import TableSettings
from word_formatter.core.rules.xml_helpers import TBLPR_ORDER, ensure_single_ordered_child


def set_table_geometry(table, settings: TableSettings) -> list[str]:
    """Приводит ширину и отступ таблицы к требованиям профиля.

    Возвращает список выполненных действий — для отчёта.
    """
    # Свойства ищем напрямую в XML: обращение через python-docx требует
    # обязательного w:tblPr и падает на документе, где его нет.
    tblPr = table._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        # Свойства таблицы идут первыми в w:tbl, перед сеткой и строками.
        tblPr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tblPr)
    applied: list[str] = []

    if settings.preferred_width == "auto":
        # Значение auto и ноль означает «ширина не задана»: ровно то, что
        # получается, если снять галочку «Ширина» в свойствах таблицы.
        width = ensure_single_ordered_child(tblPr, "w:tblW", TBLPR_ORDER)
        width.set(qn("w:w"), "0")
        width.set(qn("w:type"), "auto")
        applied.append("ширина не задана")

    if settings.indent_cm is not None:
        indent = ensure_single_ordered_child(tblPr, "w:tblInd", TBLPR_ORDER)
        indent.set(qn("w:w"), str(Cm(settings.indent_cm).twips))
        indent.set(qn("w:type"), "dxa")
        applied.append(f"отступ от левого поля {settings.indent_cm:g} см")

    return applied
