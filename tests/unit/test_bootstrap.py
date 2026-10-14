"""Подготовка окружения: проверка версии Python и работа с зависимостями."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

from tools import bootstrap


def test_supported_versions() -> None:
    assert bootstrap.python_is_supported((3, 10, 0)) is True
    assert bootstrap.python_is_supported((3, 11, 5)) is True
    assert bootstrap.python_is_supported((3, 13, 1)) is True
    assert bootstrap.python_is_supported((4, 0, 0)) is True
    assert bootstrap.python_is_supported((3, 9, 18)) is False
    assert bootstrap.python_is_supported((2, 7, 18)) is False


def test_unsupported_version_message_windows() -> None:
    message = bootstrap.unsupported_version_message((3, 9, 18), "nt")

    assert "3.9.18" in message
    assert "python.org" in message
    assert "Add python.exe to PATH" in message


def test_unsupported_version_message_linux() -> None:
    message = bootstrap.unsupported_version_message((2, 7, 18), "posix")

    assert "2.7.18" in message
    assert "apt install" in message
    assert "bash run_linux.sh" in message


def test_venv_python_paths(tmp_path: Path) -> None:
    assert bootstrap.venv_python(tmp_path, "nt") == tmp_path / ".venv" / "Scripts" / "python.exe"
    assert bootstrap.venv_python(tmp_path, "posix") == tmp_path / ".venv" / "bin" / "python"


def test_requirements_digest_requires_file(tmp_path: Path) -> None:
    with pytest.raises(bootstrap.BootstrapError) as error:
        bootstrap.requirements_digest(tmp_path / "requirements.txt")

    assert "requirements.txt" in error.value.args[0]


def test_install_is_needed_when_stamp_is_missing(tmp_path: Path) -> None:
    requirements = tmp_path / "requirements.txt"
    requirements.write_text("python-docx==1.2.0\n", encoding="utf-8")
    venv_dir = tmp_path / ".venv"
    venv_dir.mkdir()

    assert bootstrap.needs_install(requirements, venv_dir) is True


def test_install_is_skipped_when_stamp_matches(tmp_path: Path) -> None:
    requirements = tmp_path / "requirements.txt"
    requirements.write_text("python-docx==1.2.0\n", encoding="utf-8")
    venv_dir = tmp_path / ".venv"
    venv_dir.mkdir()
    (venv_dir / bootstrap.STAMP_FILE).write_text(
        bootstrap.requirements_digest(requirements), encoding="utf-8"
    )

    assert bootstrap.needs_install(requirements, venv_dir) is False


def test_install_is_needed_when_requirements_changed(tmp_path: Path) -> None:
    requirements = tmp_path / "requirements.txt"
    requirements.write_text("python-docx==1.2.0\n", encoding="utf-8")
    venv_dir = tmp_path / ".venv"
    venv_dir.mkdir()
    (venv_dir / bootstrap.STAMP_FILE).write_text("старый отпечаток", encoding="utf-8")

    assert bootstrap.needs_install(requirements, venv_dir) is True


def test_empty_stamp_means_not_installed(tmp_path: Path) -> None:
    venv_dir = tmp_path / ".venv"
    venv_dir.mkdir()
    (venv_dir / bootstrap.STAMP_FILE).write_text("   \n", encoding="utf-8")

    assert bootstrap.installed_digest(venv_dir) is None


def test_missing_environment_is_reported(tmp_path: Path) -> None:
    (tmp_path / "requirements.txt").write_text("python-docx==1.2.0\n", encoding="utf-8")

    with pytest.raises(bootstrap.BootstrapError) as error:
        bootstrap.ensure_environment(tmp_path, allow_install=False, log=lambda _message: None)

    assert ".venv" in error.value.args[0]


def test_existing_environment_is_reused(tmp_path: Path) -> None:
    requirements = tmp_path / "requirements.txt"
    requirements.write_text("python-docx==1.2.0\n", encoding="utf-8")
    interpreter = bootstrap.venv_python(tmp_path)
    interpreter.parent.mkdir(parents=True)
    interpreter.write_text("", encoding="utf-8")
    (bootstrap.venv_directory(tmp_path) / bootstrap.STAMP_FILE).write_text(
        bootstrap.requirements_digest(requirements), encoding="utf-8"
    )

    result = bootstrap.ensure_environment(tmp_path, allow_install=False, log=lambda _message: None)

    assert result == interpreter


def test_dependencies_importable_uses_stderr() -> None:
    class Result:
        returncode = 1
        stderr = "ModuleNotFoundError: No module named 'docx'"
        stdout = ""

    ok, details = bootstrap.dependencies_importable(Path("python"), run=lambda *a, **k: Result())

    assert ok is False
    assert "No module named" in details


def test_qt_problem_message_mentions_libraries() -> None:
    linux = bootstrap.qt_problem_message("ошибка загрузки", "posix")
    windows = bootstrap.qt_problem_message("ошибка загрузки", "nt")

    assert "libxcb" in linux
    assert "apt install" in linux
    assert "libxcb" not in windows
    assert ".venv" in windows
    assert "ошибка загрузки" in linux


def test_launch_requires_entry_point(tmp_path: Path) -> None:
    with pytest.raises(bootstrap.BootstrapError) as error:
        bootstrap.launch(Path("python"), tmp_path, [], run=lambda *a, **k: 0)

    assert "main.py" in error.value.args[0]


def test_main_refuses_old_python(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    monkeypatch.setattr(sys, "version_info", (3, 9, 18))

    code = bootstrap.main(["--setup-only"])

    assert code == bootstrap.EXIT_BAD_PYTHON
    assert "3.9.18" in capsys.readouterr().err


def test_main_reports_missing_environment(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    (tmp_path / "requirements.txt").write_text("python-docx==1.2.0\n", encoding="utf-8")

    code = bootstrap.main(["--root", str(tmp_path), "--setup-only", "--no-install"])

    assert code == bootstrap.EXIT_ENVIRONMENT
    assert ".venv" in capsys.readouterr().err


def test_main_setup_only_with_ready_environment(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    requirements = tmp_path / "requirements.txt"
    requirements.write_text("python-docx==1.2.0\n", encoding="utf-8")
    interpreter = bootstrap.venv_python(tmp_path)
    interpreter.parent.mkdir(parents=True)
    interpreter.write_text("", encoding="utf-8")
    venv_dir = bootstrap.venv_directory(tmp_path)
    (venv_dir / bootstrap.STAMP_FILE).write_text(
        bootstrap.requirements_digest(requirements), encoding="utf-8"
    )

    class Ok:
        returncode = 0
        stderr = ""
        stdout = ""

    messages: list[str] = []
    code = bootstrap.main(
        ["--root", str(tmp_path), "--setup-only", "--no-install"],
        log=messages.append,
        run_capture=lambda *a, **k: Ok(),
    )

    assert code == bootstrap.EXIT_OK
    assert any(str(interpreter) in message for message in messages)
    assert (venv_dir / bootstrap.QT_MARKER_FILE).is_file()


def test_main_reports_qt_problem(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    requirements = tmp_path / "requirements.txt"
    requirements.write_text("python-docx==1.2.0\n", encoding="utf-8")
    interpreter = bootstrap.venv_python(tmp_path)
    interpreter.parent.mkdir(parents=True)
    interpreter.write_text("", encoding="utf-8")
    (bootstrap.venv_directory(tmp_path) / bootstrap.STAMP_FILE).write_text(
        bootstrap.requirements_digest(requirements), encoding="utf-8"
    )

    class Failed:
        returncode = 1
        stderr = "could not load the Qt platform plugin"
        stdout = ""

    code = bootstrap.main(
        ["--root", str(tmp_path), "--setup-only", "--no-install"],
        log=lambda _message: None,
        run_capture=lambda *a, **k: Failed(),
    )

    assert code == bootstrap.EXIT_QT_PROBLEM
    assert "Qt" in capsys.readouterr().err
