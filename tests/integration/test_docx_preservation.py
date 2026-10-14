"""Шаг 0. Проверка сохранности DOCX при работе через python-docx.

Смысл проверки: прежде чем писать правила форматирования, нужно убедиться,
что библиотека не теряет «сложные» элементы документа — сноски, комментарии,
объекты OLE, элементы управления содержимым, изображения, поля и связи.

Измерения выполняются отдельным «прибором» (tests/support.py), а не средствами
самой python-docx: иначе проверка сохранности проверяла бы саму себя.

Что установлено фактически:
  * части, которых библиотека не понимает (сноски, изображения, объекты OLE,
    пользовательский XML), переносятся байт в байт;
  * части, которые она разбирает (текст, стили, настройки, комментарии,
    колонтитулы, свойства), записываются заново с другими отступами, но состав
    элементов и тексты совпадают;
  * действующие типы содержимого всех частей не меняются.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.shared import Pt

from fixtures.rich_docx import BYTE_IDENTICAL_PARTS, RESERIALIZED_PARTS, build_rich_docx
from support import (
    count,
    count_attr,
    effective_content_types,
    non_formatting_difference,
    read_parts,
    relationships,
    sha256,
    structure_diff,
    stylesheet_ids,
    texts,
)

TEXT_PARTS = [
    "word/document.xml",
    "word/footnotes.xml",
    "word/comments.xml",
    "word/header1.xml",
    "word/footer1.xml",
]

# Элементы, которые обязаны пережить любую обработку без изменения количества.
STRUCTURAL_TAGS = [
    "drawing",
    "object",
    "sdt",
    "hyperlink",
    "footnoteReference",
    "commentReference",
    "tbl",
    "fldChar",
    "sectPr",
    "instrText",
]


def _roundtrip(source: Path, target: Path) -> None:
    """Открывает документ и сохраняет его без изменений."""
    document = Document(str(source))
    document.save(str(target))


def _roundtrip_with_edit(source: Path, target: Path) -> None:
    """Открывает документ, меняет оформление первого абзаца и сохраняет."""
    document = Document(str(source))
    document.paragraphs[0].runs[0].font.size = Pt(20)
    document.paragraphs[0].runs[0].font.name = "Times New Roman"
    document.save(str(target))


def test_fixture_is_readable(rich_docx: Path) -> None:
    document = Document(str(rich_docx))
    assert len(document.paragraphs) > 5
    assert len(document.tables) == 1
    assert len(document.sections) == 2
    assert len(document.inline_shapes) == 1
    assert document.paragraphs[0].text.startswith("КРАСНОДАРСКИЙ")


def test_all_parts_survive_roundtrip(rich_docx: Path, tmp_path: Path) -> None:
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)
    after = read_parts(target)

    missing = sorted(set(before) - set(after))
    assert missing == [], f"части пакета потеряны: {missing}"


def test_byte_identical_parts_survive_roundtrip(rich_docx: Path, tmp_path: Path) -> None:
    """Сноски, изображения, объекты OLE и пользовательский XML не должны меняться."""
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)
    after = read_parts(target)

    changed = [name for name in BYTE_IDENTICAL_PARTS if sha256(before[name]) != sha256(after[name])]
    assert changed == [], f"части изменены при обработке: {changed}"


def test_reserialized_parts_keep_structure(rich_docx: Path, tmp_path: Path) -> None:
    """Без правок документа пересобранные части обязаны совпасть целиком."""
    target = tmp_path / "out.docx"
    _roundtrip(rich_docx, target)

    before = read_parts(rich_docx)
    after = read_parts(target)

    problems = {}
    for name in RESERIALIZED_PARTS:
        lost, added = structure_diff(before[name], after[name])
        if lost or added:
            problems[name] = {"потеряно": lost[:3], "добавлено": added[:3]}
    assert problems == {}, f"изменился состав элементов: {problems}"


def test_editing_does_not_touch_other_parts(rich_docx: Path, tmp_path: Path) -> None:
    """Правка оформления не должна задевать ничего, кроме текста документа."""
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)
    after = read_parts(target)

    problems = {}
    for name in RESERIALIZED_PARTS:
        if name == "word/document.xml":
            continue
        lost, added = structure_diff(before[name], after[name])
        if lost or added:
            problems[name] = {"потеряно": lost[:3], "добавлено": added[:3]}
    assert problems == {}, f"правка задела посторонние части: {problems}"


def test_editing_changes_only_formatting_properties(rich_docx: Path, tmp_path: Path) -> None:
    """Главный контракт приложения: меняются только свойства оформления.

    Если обработка изменит текстовый прогон, картинку, поле, таблицу или
    сноску, этот тест это покажет.
    """
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)["word/document.xml"]
    after = read_parts(target)["word/document.xml"]

    outside_formatting = non_formatting_difference(before, after)
    assert outside_formatting == [], f"затронуто не только оформление: {outside_formatting[:3]}"


def test_effective_content_types_unchanged(rich_docx: Path, tmp_path: Path) -> None:
    """Действующий тип содержимого каждой части должен остаться прежним.

    Правила в [Content_Types].xml библиотека переписывает по-своему
    (правило по расширению заменяет правилом для конкретной части),
    поэтому сравниваются действующие типы, а не сам файл правил.
    """
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)
    after = read_parts(target)
    names = sorted(set(before) | set(after))

    old = effective_content_types(before["[Content_Types].xml"], names)
    new = effective_content_types(after["[Content_Types].xml"], names)

    changed = {name: (old[name], new[name]) for name in names if old[name] != new[name]}
    assert changed == {}, f"изменились типы содержимого: {changed}"


def test_text_is_untouched(rich_docx: Path, tmp_path: Path) -> None:
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)
    after = read_parts(target)

    for name in TEXT_PARTS:
        assert texts(before[name]) == texts(after[name]), f"изменился текст в {name}"


def test_footnotes_and_comments_are_untouched(rich_docx: Path, tmp_path: Path) -> None:
    """Комментарии библиотека разбирает, сноски — нет; проверяем оба случая."""
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)
    after = read_parts(target)

    assert count(before["word/footnotes.xml"], "footnote") == count(after["word/footnotes.xml"], "footnote")
    assert count(before["word/comments.xml"], "comment") == count(after["word/comments.xml"], "comment")
    assert "Пояснение в подстрочной сноске" in "".join(texts(after["word/footnotes.xml"]))
    assert "Проверить формулировку абзаца" in "".join(texts(after["word/comments.xml"]))


def test_structural_elements_are_untouched(rich_docx: Path, tmp_path: Path) -> None:
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)["word/document.xml"]
    after = read_parts(target)["word/document.xml"]

    for tag in STRUCTURAL_TAGS:
        assert count(before, tag) == count(after, tag), f"изменилось число элементов {tag}"

    assert count_attr(before, "br", "type", "page") == count_attr(after, "br", "type", "page")


def test_relationships_are_untouched(rich_docx: Path, tmp_path: Path) -> None:
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)["word/_rels/document.xml.rels"]
    after = read_parts(target)["word/_rels/document.xml.rels"]

    assert relationships(before) == relationships(after)


def test_styles_are_untouched(rich_docx: Path, tmp_path: Path) -> None:
    target = tmp_path / "out.docx"
    _roundtrip_with_edit(rich_docx, target)

    before = read_parts(rich_docx)["word/styles.xml"]
    after = read_parts(target)["word/styles.xml"]

    assert stylesheet_ids(before) == stylesheet_ids(after)


def test_source_file_is_not_modified(rich_docx: Path, tmp_path: Path) -> None:
    digest_before = sha256(rich_docx.read_bytes())
    _roundtrip_with_edit(rich_docx, tmp_path / "out.docx")
    assert sha256(rich_docx.read_bytes()) == digest_before


def test_fixture_can_be_rebuilt(tmp_path: Path) -> None:
    """Сборщик эталона должен давать одинаковый результат при каждом запуске."""
    first = build_rich_docx(tmp_path / "a.docx")
    second = build_rich_docx(tmp_path / "b.docx")
    assert sha256(first.read_bytes()) == sha256(second.read_bytes())
