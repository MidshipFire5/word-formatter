"""Загрузка профилей оформления.

Каталог профилей собирается из двух мест: встроенные профили рядом с
программой и личные профили пользователя. Профиль пользователя с тем же
идентификатором заменяет встроенный — так обновление проекта из GitHub
не затирает свои настройки.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from word_formatter.core.errors import ProfileError
from word_formatter.core.paths import builtin_profiles_dir, user_profiles_dir
from word_formatter.core.profiles.model import ORIGIN_BUILTIN, ORIGIN_USER, Profile
from word_formatter.core.profiles.validation import Issue, validate_profile

MAX_REPORTED_ISSUES = 10


@dataclass(frozen=True)
class ProfileInfo:
    """Сведения о профиле для списка в интерфейсе.

    Профиль с ошибкой тоже попадает в список — с заполненным полем error,
    чтобы пользователь видел проблему, а не пустой список.
    """

    id: str
    name: str
    version: str
    demo: bool
    description: str
    origin: str
    path: Path
    error: str | None = None

    @property
    def title(self) -> str:
        """Название для показа, с пометкой о демонстрационном профиле."""
        return f"{self.name} (демонстрационный)" if self.demo else self.name

    @property
    def is_usable(self) -> bool:
        return self.error is None


def format_issues(path: Path, issues: list[Issue]) -> str:
    """Собирает понятное сообщение о проблемах в файле профиля."""
    shown = issues[:MAX_REPORTED_ISSUES]
    lines = "\n".join(f"  • {issue}" for issue in shown)
    tail = ""
    if len(issues) > len(shown):
        tail = f"\n  … и ещё {len(issues) - len(shown)}"
    return f"Файл профиля «{path.name}» заполнен с ошибками:\n{lines}{tail}"


def read_profile_file(path: Path) -> object:
    """Читает файл профиля. При проблеме бросает ProfileError с понятным текстом."""
    try:
        text = path.read_text(encoding="utf-8-sig")
    except OSError as error:
        raise ProfileError(
            f"Не удалось прочитать файл профиля «{path.name}». "
            "Проверьте, что файл не открыт другой программой и доступен для чтения.",
            details=str(error),
        ) from error
    try:
        return json.loads(text)
    except json.JSONDecodeError as error:
        raise ProfileError(
            f"Файл профиля «{path.name}» не является корректным JSON. "
            f"Ошибка в строке {error.lineno}, позиция {error.colno}: {error.msg}.",
            details=str(error),
        ) from error


def load_profile_from_path(path: Path, *, origin: str = ORIGIN_BUILTIN) -> Profile:
    """Читает и проверяет профиль, возвращая готовую модель."""
    raw = read_profile_file(path)
    issues = validate_profile(raw)
    if issues:
        raise ProfileError(
            format_issues(path, issues),
            details="\n".join(str(issue) for issue in issues),
        )
    assert isinstance(raw, dict)  # гарантировано проверкой
    return Profile.from_dict(raw, path=str(path), origin=origin)


def describe_profile_file(path: Path, origin: str) -> ProfileInfo:
    """Собирает сведения о профиле, не прерываясь на ошибках."""
    fallback_id = path.stem
    try:
        raw = read_profile_file(path)
    except ProfileError as error:
        return ProfileInfo(
            id=fallback_id,
            name=path.stem,
            version="",
            demo=False,
            description="",
            origin=origin,
            path=path,
            error=error.user_message,
        )

    if not isinstance(raw, dict):
        return ProfileInfo(
            id=fallback_id,
            name=path.stem,
            version="",
            demo=False,
            description="",
            origin=origin,
            path=path,
            error=f"Файл профиля «{path.name}» должен содержать объект JSON, а не {type(raw).__name__}.",
        )

    issues = validate_profile(raw)
    return ProfileInfo(
        id=str(raw.get("id") or fallback_id),
        name=str(raw.get("name") or path.stem),
        version=str(raw.get("version") or ""),
        demo=bool(raw.get("demo", False)),
        description=str(raw.get("description") or ""),
        origin=origin,
        path=path,
        error=format_issues(path, issues) if issues else None,
    )


class ProfileCatalog:
    """Список доступных профилей и их загрузка."""

    def __init__(self, builtin_dir: Path | None = None, user_dir: Path | None = None) -> None:
        self.builtin_dir = Path(builtin_dir) if builtin_dir is not None else builtin_profiles_dir()
        self.user_dir = Path(user_dir) if user_dir is not None else user_profiles_dir()

    @property
    def directories(self) -> tuple[Path, Path]:
        return self.builtin_dir, self.user_dir

    def _profile_files(self) -> list[tuple[Path, str]]:
        found: list[tuple[Path, str]] = []
        for directory, origin in ((self.builtin_dir, ORIGIN_BUILTIN), (self.user_dir, ORIGIN_USER)):
            if not directory.is_dir():
                continue
            for path in sorted(directory.glob("*.json")):
                found.append((path, origin))
        return found

    def list_infos(self) -> list[ProfileInfo]:
        """Список профилей: личные профили заменяют встроенные с тем же id."""
        merged: dict[str, ProfileInfo] = {}
        for path, origin in self._profile_files():
            info = describe_profile_file(path, origin)
            existing = merged.get(info.id)
            if existing is None or (existing.origin == ORIGIN_BUILTIN and info.origin == ORIGIN_USER):
                merged[info.id] = info
        return sorted(merged.values(), key=lambda info: info.name.casefold())

    def find(self, profile_id: str) -> ProfileInfo:
        infos = self.list_infos()
        for info in infos:
            if info.id == profile_id:
                return info
        available = ", ".join(info.id for info in infos) or "нет ни одного профиля"
        raise ProfileError(f"Профиль «{profile_id}» не найден. Доступные профили: {available}.")

    def load(self, profile_id: str) -> Profile:
        """Загружает профиль по идентификатору."""
        info = self.find(profile_id)
        if info.error:
            raise ProfileError(info.error, details=f"файл: {info.path}")
        return load_profile_from_path(info.path, origin=info.origin)

    def load_default(self) -> Profile:
        """Первый пригодный профиль — используется, когда выбор не сделан."""
        for info in self.list_infos():
            if info.is_usable:
                return load_profile_from_path(info.path, origin=info.origin)
        raise ProfileError(
            "Не найдено ни одного исправного профиля оформления. "
            "Проверьте папку профилей: " + str(self.builtin_dir),
        )
