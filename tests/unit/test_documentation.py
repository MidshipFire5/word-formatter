"""Документация и настройка непрерывной интеграции.

Проверки ловят то, что легко забыть: битые ссылки в документации, пример
профиля, который на самом деле не проходит проверку, и настройку CI, которая
не запускает тесты.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from word_formatter.core.profiles import validate_profile

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
README = PROJECT_ROOT / "README.md"
AGENTS = PROJECT_ROOT / "AGENTS.md"
STRUCTURE = PROJECT_ROOT / "PROJECT_STRUCTURE.md"
LICENSE = PROJECT_ROOT / "LICENSE"
DOCS = PROJECT_ROOT / "docs"
WORKFLOW = PROJECT_ROOT / ".github" / "workflows" / "ci.yml"

DOCUMENTED_FILES = [
    README,
    STRUCTURE,
    DOCS / "SPEC.md",
    DOCS / "ARCHITECTURE.md",
    DOCS / "PROFILES.md",
]

# Документы для работы над проектом вместе с ИИ-агентом. Они остаются
# у разработчика, но в публичный репозиторий не попадают, поэтому
# проверяем их только тогда, когда они есть в этой копии проекта.
LOCAL_ONLY_FILES = [
    AGENTS,
    DOCS / "ACCEPTANCE.md",
]

DOCUMENTED_FILES = DOCUMENTED_FILES + [path for path in LOCAL_ONLY_FILES if path.is_file()]
LINK_PATTERN = re.compile(r"\]\(([^)#\s]+)\)")


@pytest.mark.parametrize("path", DOCUMENTED_FILES)
def test_document_exists_and_is_not_empty(path: Path) -> None:
    assert path.is_file(), f"нет документа {path.relative_to(PROJECT_ROOT)}"
    assert len(path.read_text(encoding="utf-8").splitlines()) > 20


@pytest.mark.parametrize("path", DOCUMENTED_FILES)
def test_relative_links_point_to_existing_files(path: Path) -> None:
    """Битая ссылка в документации — это тупик для читателя."""
    text = path.read_text(encoding="utf-8")
    broken: list[str] = []
    for target in LINK_PATTERN.findall(text):
        if target.startswith(("http://", "https://", "mailto:")):
            continue
        resolved = (path.parent / target).resolve()
        if not resolved.exists():
            broken.append(target)
    assert broken == [], f"в {path.name} ссылки ведут в никуда: {broken}"


def test_readme_covers_user_questions() -> None:
    text = README.read_text(encoding="utf-8")

    for section in (
        "Системные требования",
        "Запуск в Windows",
        "Запуск в Linux",
        "Ручной запуск через main.py",
        "Профили оформления",
        "Отчёт об обработке",
        "Если что-то не работает",
        "Для разработчика",
    ):
        assert section in text, f"в README нет раздела «{section}»"


def test_readme_explains_what_application_does_not_do() -> None:
    text = README.read_text(encoding="utf-8")

    assert "Чего приложение не делает" in text
    assert "не переписывает текст" in text
    assert "не создаёт титульный лист" in text
    assert "не пересчитывает автоматическое содержание" in text


def test_readme_mentions_linux_fallback_launch() -> None:
    text = README.read_text(encoding="utf-8")

    assert "bash run_linux.sh" in text
    assert "chmod +x run_linux.sh" in text


def test_readme_mentions_shipped_profile() -> None:
    text = README.read_text(encoding="utf-8")

    assert "fu-krasnodar-2026" in text
    assert (PROJECT_ROOT / "profiles" / "fu-krasnodar-2026.json").is_file()


def test_spec_lists_acceptance_criteria() -> None:
    text = (DOCS / "SPEC.md").read_text(encoding="utf-8")

    assert "Критерии готовности" in text
    # Двенадцать критериев из задания
    criteria = re.findall(r"^\d+\. ", text, flags=re.MULTILINE)
    assert len(criteria) >= 12


def test_profile_example_in_documentation_is_valid() -> None:
    """Пример в документации обязан проходить настоящую проверку профиля."""
    text = (DOCS / "PROFILES.md").read_text(encoding="utf-8")
    blocks = re.findall(r"```json\n(.*?)```", text, flags=re.DOTALL)
    assert blocks, "в docs/PROFILES.md нет примера профиля"

    example = json.loads(blocks[0])
    issues = validate_profile(example)

    assert issues == [], "пример профиля не проходит проверку: " + "; ".join(str(i) for i in issues)


def test_profiles_document_lists_every_required_section() -> None:
    text = (DOCS / "PROFILES.md").read_text(encoding="utf-8")

    for section in (
        "`page`", "`body`", "`headings`", "`captions`", "`tables`",
        "`footnotes`", "`lists`", "`pageNumbers`", "`hyphenation`", "`blankLines`",
    ):
        assert section in text, f"в описании профиля нет раздела {section}"


def test_architecture_document_explains_key_decisions() -> None:
    text = (DOCS / "ARCHITECTURE.md").read_text(encoding="utf-8")

    assert "Порядок элементов OOXML" in text
    assert "asciiTheme" in text
    assert "Проверка сохранности" in text
    assert "cmd.exe разбирает пакетные файлы" in text
    assert "не импортирует" in text


def test_agents_rules_are_project_specific() -> None:
    if not AGENTS.is_file():
        pytest.skip("AGENTS.md не публикуется в репозитории")
    text = AGENTS.read_text(encoding="utf-8")

    assert "Содержание важнее оформления" in text
    assert "python -m pytest" in text
    assert "run_windows.bat" in text
    assert "update-index --chmod=+x" in text


def test_agent_documents_are_not_published() -> None:
    """Документы для работы ИИ-агента остаются у разработчика.

    Публичный репозиторий содержит приложение и документацию к нему,
    а служебные документы агента — нет.
    """
    ignored = (PROJECT_ROOT / ".gitignore").read_text(encoding="utf-8")

    for name in ("AGENTS.md", "docs/ACCEPTANCE.md"):
        assert name in ignored, f"{name} должен быть исключён из репозитория"

    # Документы, которые публикуются, наоборот не должны быть исключены.
    for name in ("README.md", "docs/SPEC.md", "docs/ARCHITECTURE.md", "docs/PROFILES.md"):
        assert name not in ignored, f"{name} должен публиковаться"


def test_project_structure_matches_reality() -> None:
    """Документ о структуре должен упоминать существующие каталоги."""
    text = STRUCTURE.read_text(encoding="utf-8")

    for path in (
        "word_formatter/core/rules/",
        "word_formatter/ui/",
        "profiles/",
        "messages/",
        "tests/unit/",
        "tests/integration/",
        "docs/",
    ):
        assert path in text, f"в PROJECT_STRUCTURE.md нет {path}"

    for directory in ("word_formatter", "profiles", "tools", "tests", "docs"):
        assert (PROJECT_ROOT / directory).is_dir()


def test_license_is_mit_and_mentioned_in_readme() -> None:
    """Публичный проект без лицензии нельзя законно использовать."""
    assert LICENSE.is_file(), "в проекте нет файла LICENSE"

    text = LICENSE.read_text(encoding="utf-8")
    assert text.startswith("MIT License")
    assert "Permission is hereby granted" in text
    assert "WITHOUT WARRANTY" in text

    readme = README.read_text(encoding="utf-8")
    assert "MIT" in readme
    assert "LICENSE" in readme


def test_readme_mentions_permission_problem() -> None:
    """Частая жалоба: «нет разрешения на сохранение файлов в этом месте»."""
    text = README.read_text(encoding="utf-8")

    assert "нет разрешения на сохранение" in text
    assert "папку самой программы" in text


# --- непрерывная интеграция ---------------------------------------------------


def test_workflow_exists() -> None:
    assert WORKFLOW.is_file()
    assert WORKFLOW.read_text(encoding="utf-8").startswith("name:")


def test_workflow_runs_tests_on_both_systems() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")

    assert "windows-latest" in text
    assert "ubuntu-latest" in text
    assert "python -m pytest" in text


def test_workflow_installs_development_requirements() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")

    assert "requirements-dev.txt" in text


def test_workflow_installs_qt_libraries_on_linux() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")

    assert "libxcb-cursor0" in text
    assert "apt-get install" in text


def test_workflow_checks_first_run_on_linux() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")

    assert "bash run_linux.sh --setup-only" in text
    assert "chmod +x run_linux.sh" in text


def test_workflow_does_not_use_tabs_for_indentation() -> None:
    """YAML не допускает табуляцию для отступов."""
    assert "\t" not in WORKFLOW.read_text(encoding="utf-8")


def test_workflow_checks_that_tests_leave_no_files() -> None:
    """Тесты не должны оставлять файлов в репозитории.

    Однажды временные файлы pytest попали в коммит: проверка в CI ловит это
    до того, как мусор уедет в репозиторий.
    """
    text = WORKFLOW.read_text(encoding="utf-8")

    assert "git status --porcelain" in text


def test_pytest_config_does_not_point_to_missing_directory() -> None:
    """Папка для временных файлов не должна задаваться: в клоне её нет."""
    config = (PROJECT_ROOT / "pytest.ini").read_text(encoding="utf-8")

    assert "--basetemp" not in config


def test_test_artifacts_are_ignored() -> None:
    ignored = (PROJECT_ROOT / ".gitignore").read_text(encoding="utf-8")

    for pattern in ("pytest-of-*/", "tmp*.json"):
        assert pattern in ignored, f"артефакты тестов не исключены: {pattern}"
