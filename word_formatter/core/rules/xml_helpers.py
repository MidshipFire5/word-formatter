"""Аккуратная правка XML документа.

Порядок элементов внутри w:rPr и w:pPr задан схемой OOXML. Word строго
следует этому порядку: если вставить элемент не на своё место, документ
может открыться с ошибкой. Поэтому вставка выполняется по таблице порядка.
"""

from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

RPR_ORDER = (
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
    "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint",
    "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing",
    "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
    "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath",
)

PPR_ORDER = (
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
    "w:suppressAutoHyphenation", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
)

# Порядок элементов в w:settings тоже задан схемой OOXML: элемент не на своём
# месте может привести к тому, что Word откажется открывать документ.
SETTINGS_ORDER = (
    "w:writeProtection",
    "w:view",
    "w:zoom",
    "w:removePersonalInformation",
    "w:removeDateAndTime",
    "w:doNotDisplayPageBoundaries",
    "w:displayBackgroundShape",
    "w:proofState",
    "w:attachedTemplate",
    "w:documentType",
    "w:documentProtection",
    "w:defaultTabStop",
    "w:autoHyphenation",
    "w:consecutiveHyphenLimit",
    "w:hyphenationZone",
    "w:doNotHyphenateCaps",
    "w:characterSpacingControl",
    "w:updateFields",
    "w:hdrShapeDefaults",
    "w:footnotePr",
    "w:endnotePr",
    "w:compat",
    "w:rsids",
    "w:mathPr",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:shapeDefaults",
    "w:decimalSymbol",
    "w:listSeparator",
)

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Порядок элементов в w:tblPr: свойства таблицы тоже подчиняются схеме.
TBLPR_ORDER = (
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
    "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
    "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
    "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
    "w:tblPrChange",
)


def _position(order: tuple[str, ...], tag: str) -> int:
    try:
        return order.index(tag)
    except ValueError:
        return len(order)


def ensure_child(parent, tag: str, *, prepend: bool = False):
    """Возвращает дочерний элемент, создавая его при необходимости."""
    child = parent.find(qn(tag))
    if child is not None:
        return child
    child = OxmlElement(tag)
    if prepend:
        parent.insert(0, child)
    else:
        parent.append(child)
    return child


def ensure_ordered_child(parent, tag: str, order: tuple[str, ...]):
    """Возвращает дочерний элемент, создавая его на месте, предписанном схемой."""
    child = parent.find(qn(tag))
    if child is not None:
        return child
    child = OxmlElement(tag)
    position = _position(order, tag)
    inserted = False
    for index, existing in enumerate(parent):
        existing_tag = existing.tag
        # сравниваем по имени с пространством имён WordprocessingML
        name = existing_tag.rsplit("}", 1)[-1]
        if _position(order, f"w:{name}") > position:
            parent.insert(index, child)
            inserted = True
            break
    if not inserted:
        parent.append(child)
    return child


def ensure_single_ordered_child(parent, tag: str, order: tuple[str, ...]):
    """Оставляет один элемент с таким именем на месте, предписанном схемой.

    Word иногда сохраняет два одинаковых элемента — например, `w:tblW`:
    один от стиля таблицы, другой от автора. Тогда значение ширины читается
    неоднозначно. Лишние элементы удаляются, остаётся первый.
    """
    children = parent.findall(qn(tag))
    if children:
        keep = children[0]
        for extra in children[1:]:
            parent.remove(extra)
        return keep
    return ensure_ordered_child(parent, tag, order)


def set_rfonts(rpr, font_name: str) -> None:
    """Задаёт явную гарнитуру шрифта, убирая ссылку на тему оформления.

    Атрибуты вида asciiTheme имеют приоритет над ascii, поэтому их нужно
    удалять: иначе Word продолжит показывать шрифт темы.
    """
    rfonts = ensure_ordered_child(rpr, "w:rFonts", RPR_ORDER)
    for attribute in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        key = qn(attribute)
        if key in rfonts.attrib:
            del rfonts.attrib[key]
    for attribute in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attribute), font_name)
