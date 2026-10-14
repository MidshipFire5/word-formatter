"""Каталог профилей: встроенные, личные, повреждённые."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from word_formatter.core.errors import ProfileError
from word_formatter.core.profiles import ORIGIN_BUILTIN, ORIGIN_USER, ProfileCatalog


def write_profile(directory: Path, name: str, data: object, *, encoding: str = "utf-8") -> Path:
    path = directory / name
    if isinstance(data, str):
        path.write_text(data, encoding=encoding)
    else:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding=encoding)
    return path


def test_catalog_lists_builtin_profile(builtin_profile_dir: Path, user_profile_dir: Path) -> None:
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    infos = catalog.list_infos()
    assert [info.id for info in infos] == ["fu-krasnodar-2026"]
    assert infos[0].origin == ORIGIN_BUILTIN
    assert infos[0].is_usable
    assert infos[0].error is None


def test_catalog_loads_profile_by_id(builtin_profile_dir: Path, user_profile_dir: Path) -> None:
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    profile = catalog.load("fu-krasnodar-2026")
    assert profile.id == "fu-krasnodar-2026"
    assert profile.origin == ORIGIN_BUILTIN


def test_user_profile_replaces_builtin(
    builtin_profile_dir: Path, user_profile_dir: Path, profile_dict: dict
) -> None:
    own = dict(profile_dict)
    own["name"] = "Моя кафедра — свои требования"
    own["demo"] = False
    write_profile(user_profile_dir, "my-department.json", own)

    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    infos = catalog.list_infos()

    assert len(infos) == 1
    assert infos[0].name == "Моя кафедра — свои требования"
    assert infos[0].origin == ORIGIN_USER
    assert catalog.load("fu-krasnodar-2026").name == "Моя кафедра — свои требования"


def test_second_profile_is_added_not_replaced(
    builtin_profile_dir: Path, user_profile_dir: Path, profile_dict: dict
) -> None:
    other = dict(profile_dict)
    other["id"] = "other-department-2026"
    other["name"] = "Другая кафедра"
    write_profile(user_profile_dir, "other.json", other)

    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    ids = sorted(info.id for info in catalog.list_infos())
    assert ids == ["fu-krasnodar-2026", "other-department-2026"]


def test_broken_json_is_listed_with_error(builtin_profile_dir: Path, user_profile_dir: Path) -> None:
    write_profile(builtin_profile_dir, "broken.json", "{ это не JSON")

    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    infos = {info.id: info for info in catalog.list_infos()}

    assert "broken" in infos
    assert infos["broken"].error is not None
    assert "не является корректным JSON" in infos["broken"].error
    assert infos["broken"].is_usable is False


def test_invalid_profile_is_listed_with_error(
    builtin_profile_dir: Path, user_profile_dir: Path, profile_dict: dict
) -> None:
    bad = dict(profile_dict)
    bad["id"] = "bad-profile"
    bad["body"]["alignment"] = "по ширине"
    write_profile(builtin_profile_dir, "bad.json", bad)

    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    infos = {info.id: info for info in catalog.list_infos()}

    assert "body.alignment" in (infos["bad-profile"].error or "")
    # Исправный профиль рядом с повреждённым продолжает работать
    assert infos["fu-krasnodar-2026"].is_usable


def test_loading_broken_profile_explains_problem(
    builtin_profile_dir: Path, user_profile_dir: Path
) -> None:
    write_profile(builtin_profile_dir, "broken.json", "{ это не JSON")
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)

    with pytest.raises(ProfileError) as error:
        catalog.load("broken")

    assert "broken.json" in error.value.user_message
    assert "JSON" in error.value.user_message


def test_unknown_profile_lists_available_ids(
    builtin_profile_dir: Path, user_profile_dir: Path
) -> None:
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)

    with pytest.raises(ProfileError) as error:
        catalog.load("нет-такого")

    assert "не найден" in error.value.user_message
    assert "fu-krasnodar-2026" in error.value.user_message


def test_missing_directories_are_not_an_error(tmp_path: Path) -> None:
    catalog = ProfileCatalog(builtin_dir=tmp_path / "нет", user_dir=tmp_path / "тоже-нет")
    assert catalog.list_infos() == []

    with pytest.raises(ProfileError):
        catalog.load_default()


def test_load_default_skips_broken_profile(
    builtin_profile_dir: Path, user_profile_dir: Path
) -> None:
    write_profile(builtin_profile_dir, "broken.json", "{ это не JSON")
    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)

    assert catalog.load_default().id == "fu-krasnodar-2026"


def test_file_with_bom_is_accepted(
    builtin_profile_dir: Path, user_profile_dir: Path, profile_dict: dict
) -> None:
    """Файл, сохранённый Блокнотом с меткой порядка байтов, должен читаться."""
    write_profile(builtin_profile_dir, "bom.json", profile_dict, encoding="utf-8-sig")

    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    infos = {info.id: info for info in catalog.list_infos()}

    assert infos["fu-krasnodar-2026"].is_usable
    assert catalog.load("fu-krasnodar-2026").body.size_pt == 14


def test_demo_profile_is_marked(builtin_profile_dir: Path, user_profile_dir: Path, profile_dict: dict) -> None:
    demo = dict(profile_dict)
    demo["id"] = "demo-2026"
    demo["name"] = "Демонстрационный профиль"
    demo["demo"] = True
    write_profile(builtin_profile_dir, "demo.json", demo)

    catalog = ProfileCatalog(builtin_dir=builtin_profile_dir, user_dir=user_profile_dir)
    infos = {info.id: info for info in catalog.list_infos()}

    assert infos["demo-2026"].demo is True
    assert "демонстрационный" in infos["demo-2026"].title
