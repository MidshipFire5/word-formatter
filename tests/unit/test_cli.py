"""Проверочные команды запуска: --version и --list-profiles."""

from __future__ import annotations

from pathlib import Path

import pytest

from word_formatter import APP_VERSION
from word_formatter.app import EXIT_ERROR, EXIT_OK, main


@pytest.fixture(autouse=True)
def isolated_user_dir(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Личные профили и журналы не должны попадать в настоящий профиль пользователя."""
    monkeypatch.setenv("APPDATA", str(tmp_path / "appdata"))
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "config"))


def test_version_command(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--version"]) == EXIT_OK
    out = capsys.readouterr().out
    assert "Word Formatter" in out
    assert APP_VERSION in out


def test_list_profiles_shows_shipped_profile(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["--list-profiles"]) == EXIT_OK
    out = capsys.readouterr().out
    assert "fu-krasnodar-2026" in out
    assert "Финансовый университет" in out
    assert "встроенный" in out


def test_list_profiles_prints_directories(capsys: pytest.CaptureFixture[str]) -> None:
    main(["--list-profiles"])
    out = capsys.readouterr().out
    assert "Встроенные профили:" in out
    assert "Личные профили:" in out


def test_profiles_dir_argument(capsys: pytest.CaptureFixture[str], builtin_profile_dir: Path) -> None:
    assert main(["--list-profiles", "--profiles-dir", str(builtin_profile_dir)]) == EXIT_OK
    out = capsys.readouterr().out
    assert str(builtin_profile_dir) in out
    assert "fu-krasnodar-2026" in out


def test_empty_profiles_dir_is_reported(
    capsys: pytest.CaptureFixture[str], tmp_path: Path
) -> None:
    empty = tmp_path / "пусто"
    empty.mkdir()
    assert main(["--list-profiles", "--profiles-dir", str(empty)]) == EXIT_ERROR
    out = capsys.readouterr().out
    assert "Профили не найдены" in out


def test_broken_profile_is_reported_in_list(
    capsys: pytest.CaptureFixture[str], builtin_profile_dir: Path
) -> None:
    (builtin_profile_dir / "broken.json").write_text("{ это не JSON", encoding="utf-8")
    main(["--list-profiles", "--profiles-dir", str(builtin_profile_dir)])
    out = capsys.readouterr().out
    assert "ОШИБКА" in out
    assert "JSON" in out
    # Исправный профиль рядом остаётся в списке
    assert "fu-krasnodar-2026" in out
