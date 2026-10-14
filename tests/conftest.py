"""Общие приспособления pytest."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pytest

# Графический интерфейс в тестах работает без экрана: это позволяет
# проверять окно и в CI, и на сервере без дисплея.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TESTS_DIR = PROJECT_ROOT / "tests"
for path in (PROJECT_ROOT, TESTS_DIR):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

BUILTIN_PROFILE = PROJECT_ROOT / "profiles" / "fu-krasnodar-2026.json"


@pytest.fixture()
def rich_docx(tmp_path: Path) -> Path:
    """Путь к эталонному документу со «сложными» элементами."""
    from fixtures.rich_docx import build_rich_docx

    return build_rich_docx(tmp_path / "source.docx")


@pytest.fixture()
def profile_dict() -> dict:
    """Свежая копия встроенного профиля: тесты вправе её портить."""
    return json.loads(BUILTIN_PROFILE.read_text(encoding="utf-8"))


@pytest.fixture()
def profile_file(tmp_path: Path, profile_dict: dict) -> Path:
    """Файл профиля на диске, собранный из встроенного."""
    path = tmp_path / "fu-krasnodar-2026.json"
    path.write_text(json.dumps(profile_dict, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


@pytest.fixture()
def builtin_profile_dir(tmp_path: Path, profile_dict: dict) -> Path:
    """Папка встроенных профилей с одной копией реального профиля."""
    directory = tmp_path / "builtin"
    directory.mkdir()
    (directory / "fu-krasnodar-2026.json").write_text(
        json.dumps(profile_dict, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return directory


@pytest.fixture()
def user_profile_dir(tmp_path: Path) -> Path:
    """Пустая папка личных профилей пользователя."""
    directory = tmp_path / "user"
    directory.mkdir()
    return directory


@pytest.fixture()
def make_context():
    """Готовит контекст применения правил для документа."""
    from word_formatter.core.docx import inspect_document
    from word_formatter.core.rules import RuleContext

    def build(document, profile) -> RuleContext:
        return RuleContext(
            document=document,
            profile=profile,
            inventory=inspect_document(document, profile),
        )

    return build

