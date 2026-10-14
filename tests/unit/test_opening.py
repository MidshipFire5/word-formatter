"""Открытие документов: форматы, повреждённые файлы, защита паролем."""

from __future__ import annotations

from pathlib import Path

import pytest

from word_formatter.core.docx import open_document
from word_formatter.core.errors import (
    CorruptedDocumentError,
    PasswordProtectedError,
    UnsupportedFormatError,
)

OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def test_docx_opens(rich_docx: Path) -> None:
    document = open_document(rich_docx)
    assert len(document.paragraphs) > 5


def test_legacy_doc_is_refused_with_instruction(tmp_path: Path) -> None:
    path = tmp_path / "старый.doc"
    path.write_bytes(OLE_SIGNATURE + b"\x00" * 32)

    with pytest.raises(UnsupportedFormatError) as error:
        open_document(path)

    message = error.value.user_message
    assert "старом формате" in message
    assert "Сохранить как" in message
    assert ".docx" in message


def test_macro_document_is_refused(tmp_path: Path) -> None:
    path = tmp_path / "с_макросами.docm"
    path.write_bytes(b"PK\x03\x04" + b"\x00" * 32)

    with pytest.raises(UnsupportedFormatError) as error:
        open_document(path)

    assert "макросы" in error.value.user_message
    assert ".docx" in error.value.user_message


def test_other_extension_is_refused(tmp_path: Path) -> None:
    path = tmp_path / "заметки.txt"
    path.write_text("просто текст", encoding="utf-8")

    with pytest.raises(UnsupportedFormatError) as error:
        open_document(path)

    assert ".docx" in error.value.user_message


def test_missing_file_is_reported(tmp_path: Path) -> None:
    with pytest.raises(CorruptedDocumentError) as error:
        open_document(tmp_path / "нет-такого.docx")

    assert "не найден" in error.value.user_message


def test_empty_file_is_reported(tmp_path: Path) -> None:
    path = tmp_path / "пустой.docx"
    path.write_bytes(b"")

    with pytest.raises(CorruptedDocumentError) as error:
        open_document(path)

    assert "пустой" in error.value.user_message


def test_folder_instead_of_file_is_reported(tmp_path: Path) -> None:
    folder = tmp_path / "папка.docx"
    folder.mkdir()

    with pytest.raises(CorruptedDocumentError) as error:
        open_document(folder)

    assert "папка" in error.value.user_message


def test_not_a_word_document_is_reported(tmp_path: Path) -> None:
    path = tmp_path / "картинка.docx"
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 64)

    with pytest.raises(CorruptedDocumentError) as error:
        open_document(path)

    assert "не является документом Word" in error.value.user_message


def test_encrypted_document_is_reported(tmp_path: Path) -> None:
    path = tmp_path / "защищённый.docx"
    path.write_bytes(OLE_SIGNATURE + b"\x00" * 64)

    with pytest.raises(PasswordProtectedError) as error:
        open_document(path)

    assert "парол" in error.value.user_message.lower()


def test_zip_without_document_part_is_reported(tmp_path: Path) -> None:
    import zipfile

    path = tmp_path / "чужой.docx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("hello.txt", "не документ Word")

    with pytest.raises(CorruptedDocumentError) as error:
        open_document(path)

    assert error.value.user_message
    assert "повреждён" in error.value.user_message or "не удалось открыть" in error.value.user_message.lower()


def test_original_file_is_not_modified(rich_docx: Path) -> None:
    import hashlib

    before = hashlib.sha256(rich_docx.read_bytes()).hexdigest()
    open_document(rich_docx)
    assert hashlib.sha256(rich_docx.read_bytes()).hexdigest() == before
