"""Подготовка окружения и запуск приложения Word Formatter.

Этот файл запускается тем Python, который уже установлен у пользователя,
и делает всё остальное: проверяет версию, создаёт локальное окружение .venv,
устанавливает зависимости (только если они изменились) и запускает main.py.

Почему логика написана на Python, а не на языке командной строки: один и тот
же код работает и в Windows, и в Linux, а проверять его можно обычными
тестами. Скрипты run_windows.bat и run_linux.sh только находят Python и
передают управление сюда.

Ничего не устанавливается глобально, права администратора не нужны,
системные настройки не меняются, пользовательские файлы не удаляются.
"""

from __future__ import annotations

import argparse
import hashlib
import os
import subprocess
import sys
import venv
from collections.abc import Callable, Sequence
from pathlib import Path

MINIMUM_PYTHON = (3, 10)
RECOMMENDED_PYTHON = "3.11 или новее"
REQUIREMENTS_FILE = "requirements.txt"
STAMP_FILE = ".installed-requirements"
QT_MARKER_FILE = ".qt-platform-ok"
VENV_DIR_NAME = ".venv"
ENTRY_POINT = "main.py"
MESSAGES_DIR_NAME = "messages"
NO_PYTHON_MESSAGE_FILE = "no_python.txt"

EXIT_OK = 0
EXIT_BAD_PYTHON = 2
EXIT_ENVIRONMENT = 3
EXIT_QT_PROBLEM = 4

WINDOWS_INSTRUCTIONS = """\
Как установить Python в Windows:
  1. Откройте страницу https://www.python.org/downloads/windows/
  2. Скачайте «Windows installer (64-bit)» последней версии.
  3. При установке обязательно отметьте галочку «Add python.exe to PATH».
  4. Закройте это окно и запустите run_windows.bat снова.
Права администратора не нужны: подойдёт установка для текущего пользователя."""

LINUX_INSTRUCTIONS = """\
Как установить Python в Linux:
  Ubuntu и Debian:  sudo apt update && sudo apt install -y python3 python3-venv python3-pip
  Fedora:           sudo dnf install -y python3 python3-pip
  Arch:             sudo pacman -S python
После установки запустите приложение снова: bash run_linux.sh"""

VENV_MODULE_HINT = """\
Не удалось создать локальное окружение .venv: в системе нет модуля venv.
  Ubuntu и Debian:  sudo apt install -y python3-venv
  Fedora:           sudo dnf install -y python3
После установки запустите приложение снова."""

QT_HINT = """\
Qt не запускается: не хватает системных библиотек, которые нужны графическому интерфейсу.
  Ubuntu и Debian:
    sudo apt update && sudo apt install -y libgl1 libegl1 libxkbcommon-x11-0 \\
      libxcb-cursor0 libxcb-icccm4 libxcb-keysyms1 libxcb-shape0 libxcb-xinerama0 libdbus-1-3
После установки запустите приложение снова."""

WINDOWS_QT_HINT = """\
Не удалось запустить графическую библиотеку Qt.
Удалите папку .venv рядом с программой и запустите run_windows.bat снова:
зависимости установятся заново."""

QT_CHECK_CODE = (
    "import os; os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen');"
    "from PySide6.QtWidgets import QApplication; QApplication([])"
)

LogFunction = Callable[[str], None]


class BootstrapError(Exception):
    """Ошибка подготовки окружения с текстом, понятным пользователю."""


def log_to_console(message: str) -> None:
    print(message, flush=True)


def setup_console_encoding() -> None:
    """Печатает русский текст одинаково и в консоли, и при перенаправлении вывода.

    Скрипт запуска переключает консоль Windows на UTF-8, поэтому вывод тоже
    должен быть в UTF-8: иначе при перенаправлении в файл получится мусор.
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            pass


def message_file(name: str, fallback: str) -> str:
    """Читает текст сообщения из файла: он же печатается скриптом запуска.

    Так инструкция для пользователя существует в одном месте: скрипт запуска
    печатает файл, а bootstrap берёт из него же текст для своей ошибки.
    """
    path = Path(__file__).resolve().parent / MESSAGES_DIR_NAME / name
    try:
        content = path.read_text(encoding="utf-8").strip()
    except OSError:
        return fallback
    return content or fallback


def project_root() -> Path:
    """Папка проекта: родитель каталога tools."""
    return Path(__file__).resolve().parent.parent


def python_is_supported(version_info: Sequence[int]) -> bool:
    """Подходит ли версия Python для работы приложения."""
    return tuple(version_info[:2]) >= MINIMUM_PYTHON


def unsupported_version_message(version_info: Sequence[int], platform: str | None = None) -> str:
    """Объясняет, какая версия найдена и как поставить подходящую."""
    platform = platform or os.name
    found = ".".join(str(part) for part in version_info[:3])
    if platform == "nt":
        instructions = message_file(NO_PYTHON_MESSAGE_FILE, WINDOWS_INSTRUCTIONS)
    else:
        instructions = LINUX_INSTRUCTIONS
    return (
        "Не найдена подходящая версия Python.\n"
        f"Обнаружена версия: {found}\n"
        f"Нужен Python {MINIMUM_PYTHON[0]}.{MINIMUM_PYTHON[1]} или новее "
        f"(рекомендуется {RECOMMENDED_PYTHON}).\n\n"
        f"{instructions}"
    )


def venv_directory(root: Path) -> Path:
    return root / VENV_DIR_NAME


def venv_python(root: Path, platform: str | None = None) -> Path:
    """Путь к Python внутри локального окружения."""
    platform = platform or os.name
    if platform == "nt":
        return venv_directory(root) / "Scripts" / "python.exe"
    return venv_directory(root) / "bin" / "python"


def requirements_digest(requirements: Path) -> str:
    """Отпечаток файла зависимостей: по нему видно, менялся ли он."""
    if not requirements.is_file():
        raise BootstrapError(
            f"Не найден файл {requirements.name} с перечнем зависимостей. "
            "Похоже, архив с программой распакован не полностью — распакуйте его заново целиком."
        )
    return hashlib.sha256(requirements.read_bytes()).hexdigest()


def installed_digest(venv_dir: Path) -> str | None:
    """Отпечаток зависимостей, которые уже установлены."""
    stamp = venv_dir / STAMP_FILE
    if not stamp.is_file():
        return None
    try:
        value = stamp.read_text(encoding="utf-8").strip()
    except OSError:
        return None
    return value or None


def needs_install(requirements: Path, venv_dir: Path) -> bool:
    """Нужно ли устанавливать зависимости."""
    return installed_digest(venv_dir) != requirements_digest(requirements)


def create_venv(root: Path, *, log: LogFunction = log_to_console) -> None:
    """Создаёт локальное виртуальное окружение рядом с программой."""
    log("Создаю локальное окружение .venv — это займёт немного времени…")
    builder = venv.EnvBuilder(with_pip=True, clear=False, upgrade_deps=False)
    try:
        builder.create(str(venv_directory(root)))
    except Exception as error:  # окружение может не создаться по многим причинам
        hint = VENV_MODULE_HINT if os.name != "nt" else (
            "Не удалось создать локальное окружение .venv.\n"
            "Проверьте, что в папке с программой есть место и разрешение на запись."
        )
        raise BootstrapError(f"{hint}\n\nПодробность: {error}") from error


def install_requirements(
    python: Path,
    requirements: Path,
    *,
    log: LogFunction = log_to_console,
    run: Callable[..., int] = subprocess.call,
) -> None:
    """Устанавливает зависимости в локальное окружение."""
    log(
        "Устанавливаю зависимости. Сейчас нужен доступ в интернет — "
        "это происходит только при первом запуске и после обновления программы."
    )
    command = [
        str(python),
        "-m",
        "pip",
        "install",
        "--disable-pip-version-check",
        "--no-warn-script-location",
        "-r",
        str(requirements),
    ]
    code = run(command)
    if code != 0:
        raise BootstrapError(
            "Не удалось установить зависимости.\n"
            "Проверьте подключение к интернету и запустите приложение снова.\n\n"
            "Если интернета нет, установка невозможна: приложение использует библиотеки "
            "python-docx и PySide6, которые нужно скачать один раз.\n"
            "Подробный текст ошибки — выше в этом окне."
        )


def dependencies_importable(
    python: Path,
    *,
    run: Callable[..., subprocess.CompletedProcess] = subprocess.run,
) -> tuple[bool, str]:
    """Проверяет, что нужные библиотеки действительно можно загрузить."""
    result = run(
        [str(python), "-c", "import docx, PySide6.QtWidgets"],
        capture_output=True,
        text=True,
    )
    if result.returncode == 0:
        return True, ""
    return False, (result.stderr or result.stdout or "").strip()


def qt_problem_message(details: str, platform: str | None = None) -> str:
    """Объясняет, как починить незапускающийся Qt в этой системе."""
    platform = platform or os.name
    hint = QT_HINT if platform != "nt" else WINDOWS_QT_HINT
    return f"{hint}\n\nПодробность: {details}"


def qt_platform_works(
    python: Path,
    venv_dir: Path,
    *,
    run: Callable[..., subprocess.CompletedProcess] = subprocess.run,
) -> tuple[bool, str]:
    """Проверяет, что графическая библиотека запускается в этой системе.

    Проверка выполняется один раз: её результат сохраняется рядом с окружением,
    чтобы не тратить время при каждом запуске.
    """
    marker = venv_dir / QT_MARKER_FILE
    if marker.is_file():
        return True, ""

    result = run([str(python), "-c", QT_CHECK_CODE], capture_output=True, text=True)
    if result.returncode == 0:
        try:
            marker.write_text("проверено", encoding="utf-8")
        except OSError:
            pass
        return True, ""

    details = (result.stderr or result.stdout or "").strip()
    return False, qt_problem_message(details)


def ensure_environment(
    root: Path,
    *,
    log: LogFunction = log_to_console,
    allow_install: bool = True,
    run: Callable[..., int] = subprocess.call,
    run_capture: Callable[..., subprocess.CompletedProcess] = subprocess.run,
) -> Path:
    """Готовит окружение и возвращает путь к Python внутри него."""
    requirements = root / REQUIREMENTS_FILE
    interpreter = venv_python(root)

    if not interpreter.is_file():
        if not allow_install:
            raise BootstrapError(
                "Локальное окружение .venv ещё не создано. "
                "Запустите приложение без параметра --no-install, чтобы его создать."
            )
        create_venv(root, log=log)
        if not interpreter.is_file():
            raise BootstrapError(
                "Локальное окружение создано, но в нём нет Python. "
                "Удалите папку .venv и запустите приложение снова."
            )

    if needs_install(requirements, venv_directory(root)):
        if not allow_install:
            raise BootstrapError(
                "Зависимости не установлены или изменились, а установка запрещена "
                "параметром --no-install."
            )
        install_requirements(interpreter, requirements, log=log, run=run)
        ok, details = dependencies_importable(interpreter, run=run_capture)
        if not ok:
            raise BootstrapError(
                "Зависимости установлены, но не загружаются.\n"
                "Удалите папку .venv и запустите приложение снова.\n\n"
                f"Подробность: {details}"
            )
        try:
            (venv_directory(root) / STAMP_FILE).write_text(
                requirements_digest(requirements), encoding="utf-8"
            )
        except OSError:
            pass
        log("Окружение готово.")

    return interpreter


def launch(python: Path, root: Path, arguments: Sequence[str], *, run: Callable[..., int] = subprocess.call) -> int:
    """Запускает приложение из локального окружения."""
    entry_point = root / ENTRY_POINT
    if not entry_point.is_file():
        raise BootstrapError(
            f"Не найден файл {ENTRY_POINT}. "
            "Похоже, архив с программой распакован не полностью — распакуйте его заново целиком."
        )
    return run([str(python), str(entry_point), *arguments])


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="bootstrap.py",
        description="Подготовка окружения и запуск Word Formatter.",
    )
    parser.add_argument(
        "--setup-only",
        action="store_true",
        help="подготовить окружение и выйти, не запуская приложение",
    )
    parser.add_argument(
        "--no-install",
        action="store_true",
        help="не устанавливать зависимости, только проверить готовое окружение",
    )
    parser.add_argument(
        "--root",
        default=None,
        help="папка программы (по умолчанию — папка этого файла)",
    )
    return parser


def main(
    argv: Sequence[str] | None = None,
    *,
    log: LogFunction = log_to_console,
    run: Callable[..., int] = subprocess.call,
    run_capture: Callable[..., subprocess.CompletedProcess] = subprocess.run,
) -> int:
    """Точка входа подготовки окружения. Возвращает код завершения процесса."""
    parser = build_parser()
    arguments, extra = parser.parse_known_args(argv)

    setup_console_encoding()

    if not python_is_supported(sys.version_info):
        print(unsupported_version_message(sys.version_info), file=sys.stderr)
        return EXIT_BAD_PYTHON

    root = Path(arguments.root).resolve() if arguments.root else project_root()

    try:
        interpreter = ensure_environment(
            root,
            log=log,
            allow_install=not arguments.no_install,
            run=run,
            run_capture=run_capture,
        )

        works, details = qt_platform_works(interpreter, venv_directory(root), run=run_capture)
        if not works:
            print(details, file=sys.stderr)
            return EXIT_QT_PROBLEM

        if arguments.setup_only:
            log(f"Окружение готово: {interpreter}")
            return EXIT_OK

        return launch(interpreter, root, extra, run=run)
    except BootstrapError as error:
        print(f"Ошибка: {error}", file=sys.stderr)
        return EXIT_ENVIRONMENT


if __name__ == "__main__":
    sys.exit(main())
