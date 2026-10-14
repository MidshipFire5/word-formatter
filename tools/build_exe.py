"""Сборка переносимой версии приложения (PyInstaller).

Два варианта:

* **папка** (по умолчанию) — `dist/WordFormatter/`: исполняемый файл,
  библиотеки и встроенные профили рядом. Это тот вариант, который
  публикуется в релизах: он проверен полностью, включая открытие окна;
* **один файл** (`--one-file`) — `dist/WordFormatter.exe`: ничего
  распаковывать не нужно. Сборка собирается и проходит проверки без графики
  (`--version`, `--list-profiles`, `--self-check`), но запуск **окна** в этом
  варианте не подтверждён: у оконной однофайловой сборки процесс завершается
  ещё до запуска Python. Поэтому вариант помечен как необязательный.

Запуск:
    python tools/build_exe.py              # папка — то, что идёт в релиз
    python tools/build_exe.py --one-file   # один файл, не подтверждён
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
NAME = "WordFormatter"
DIST_DIR = PROJECT_ROOT / "dist"
WORK_DIR = PROJECT_ROOT / ".tmp" / "pyinstaller"
MAIN_SCRIPT = PROJECT_ROOT / "main.py"
PROFILES_DIR = PROJECT_ROOT / "profiles"


def executable_name() -> str:
    return f"{NAME}.exe" if os.name == "nt" else NAME


def build_command(*, one_file: bool) -> list[str]:
    variant = "onefile" if one_file else "onedir"
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--name",
        NAME,
        # Служебные файлы сборки держим в .tmp, чтобы не мусорить в проекте.
        "--specpath",
        str(WORK_DIR / variant),
        "--workpath",
        str(WORK_DIR / variant / "build"),
        "--distpath",
        str(DIST_DIR),
        # Встроенные профили обязаны попасть внутрь сборки: приложение ищет
        # их рядом с собой.
        "--add-data",
        f"{PROFILES_DIR}{os.pathsep}profiles",
        # pypdf подключается внутри функции, поэтому указываем явно.
        "--hidden-import",
        "pypdf",
    ]
    if one_file:
        command.append("--onefile")
    if os.name == "nt":
        # В Windows без этого при запуске открывалось бы чёрное окно консоли.
        command.append("--windowed")
    command.append(str(MAIN_SCRIPT))
    return command


def build(*, one_file: bool) -> Path | None:
    """Собирает выбранный вариант и возвращает путь к исполняемому файлу."""
    target = DIST_DIR / executable_name() if one_file else DIST_DIR / NAME / executable_name()
    if target.exists():
        target.unlink() if target.is_file() else shutil.rmtree(target, ignore_errors=True)

    print(f"Собираю {'один файл' if one_file else 'папку'}…")
    completed = subprocess.run(build_command(one_file=one_file), cwd=PROJECT_ROOT)
    if completed.returncode != 0:
        print("Сборка не удалась.")
        return None
    if not target.is_file():
        print(f"Сборка завершилась, но файла {target} нет.")
        return None
    return target


def describe(target: Path) -> None:
    if target.parent == DIST_DIR:
        size_mb = target.stat().st_size / 1024 / 1024
        print(f"Готово: {target}")
        print(f"Размер файла: {size_mb:.0f} МБ")
    else:
        size_mb = sum(item.stat().st_size for item in target.parent.rglob("*") if item.is_file())
        print(f"Готово: {target}")
        print(f"Размер папки: {size_mb / 1024 / 1024:.0f} МБ")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Сборка переносимой версии приложения")
    parser.add_argument(
        "--one-file",
        action="store_true",
        help="собрать одним файлом (необязательный вариант: открытие окна "
        "в такой сборке не подтверждено, для релиза используется папка)",
    )
    args = parser.parse_args(argv)

    if not MAIN_SCRIPT.is_file():
        print(f"Не найден {MAIN_SCRIPT}")
        return 1
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("Не установлен PyInstaller. Установите: python -m pip install -r requirements-build.txt")
        return 1

    target = build(one_file=args.one_file)
    if target is None:
        return 1

    describe(target)
    print()
    print("Проверка:")
    print(f'  "{target}" --version')
    print(f'  "{target}" --list-profiles')
    print(f'  "{target}" --self-check')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
