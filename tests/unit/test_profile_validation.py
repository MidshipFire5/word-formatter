"""Проверка файла профиля: что приложение сообщает о неверных данных."""

from __future__ import annotations

import copy
import re

from word_formatter.core.profiles import validate_profile


def messages(profile: dict) -> str:
    return "\n".join(str(issue) for issue in validate_profile(profile))


def paths(profile: dict) -> list[str]:
    return [issue.path for issue in validate_profile(profile)]


def test_builtin_profile_is_valid(profile_dict: dict) -> None:
    assert validate_profile(profile_dict) == []


def test_profile_must_be_object() -> None:
    issues = validate_profile(["не объект"])
    assert len(issues) == 1
    assert "объектом" in issues[0].message


def test_missing_required_field_is_reported(profile_dict: dict) -> None:
    profile_dict.pop("name")
    assert "name" in paths(profile_dict)
    assert "не заполнено обязательное поле" in messages(profile_dict)


def test_unknown_field_is_reported_as_typo(profile_dict: dict) -> None:
    profile_dict["page"]["marginsMM"] = {"left": 30}
    text = messages(profile_dict)
    assert "page.marginsMM" in text
    assert "опечатка" in text
    # Профиль не должен считаться исправным: иначе опечатка тихо изменит вёрстку.
    assert validate_profile(profile_dict) != []


def test_wrong_type_is_explained(profile_dict: dict) -> None:
    profile_dict["body"]["sizePt"] = "14"
    text = messages(profile_dict)
    assert "body.sizePt" in text
    assert "должно быть числом в пунктах" in text
    assert "строка" in text


def test_boolean_field_rejects_text(profile_dict: dict) -> None:
    profile_dict["blankLines"]["insert"] = "да"
    text = messages(profile_dict)
    assert "blankLines.insert" in text
    assert "true или false" in text


def test_unknown_choice_is_rejected(profile_dict: dict) -> None:
    profile_dict["body"]["alignment"] = "по ширине"
    text = messages(profile_dict)
    assert "body.alignment" in text
    assert "left, center, right, justify" in text


def test_page_orientation_choices(profile_dict: dict) -> None:
    profile_dict["page"]["orientation"] = "sideways"
    text = messages(profile_dict)
    assert "page.orientation" in text
    assert "landscape" in text


def test_numeric_range_is_checked(profile_dict: dict) -> None:
    profile_dict["page"]["marginsMm"]["left"] = -5
    text = messages(profile_dict)
    assert "page.marginsMm.left" in text
    assert "не меньше 0" in text


def test_margins_require_all_sides(profile_dict: dict) -> None:
    del profile_dict["page"]["marginsMm"]["bottom"]
    text = messages(profile_dict)
    assert "page.marginsMm.bottom" in text


def test_line_spacing_range_is_checked(profile_dict: dict) -> None:
    profile_dict["body"]["lineSpacing"] = 5
    text = messages(profile_dict)
    assert "body.lineSpacing" in text
    assert "не больше 3" in text


def test_color_must_be_hex(profile_dict: dict) -> None:
    profile_dict["body"]["color"] = "красный"
    text = messages(profile_dict)
    assert "body.color" in text
    assert "#1A2B3C" in text


def test_identifier_pattern_is_checked(profile_dict: dict) -> None:
    profile_dict["id"] = "Профиль Кафедры"
    text = messages(profile_dict)
    assert "id" in text
    assert "fu-krasnodar-2026" in text


def test_version_pattern_is_checked(profile_dict: dict) -> None:
    profile_dict["version"] = "первая"
    text = messages(profile_dict)
    assert "version" in text


def test_heading_levels_must_cover_max_level(profile_dict: dict) -> None:
    profile_dict["headings"]["maxLevel"] = 4
    text = messages(profile_dict)
    assert "headings.levels.4" in text


def test_heading_level_above_max_is_rejected(profile_dict: dict) -> None:
    profile_dict["headings"]["levels"]["5"] = copy.deepcopy(profile_dict["headings"]["levels"]["1"])
    text = messages(profile_dict)
    assert "headings.levels.5" in text
    assert "уровень не описан профилем" in text


def test_structural_titles_must_not_be_empty(profile_dict: dict) -> None:
    profile_dict["headings"]["structural"]["titles"] = []
    text = messages(profile_dict)
    assert "headings.structural.titles" in text


def test_applies_to_must_have_at_least_one_item(profile_dict: dict) -> None:
    profile_dict["appliesTo"] = []
    text = messages(profile_dict)
    assert "appliesTo" in text


def test_issue_message_contains_path_and_russian_text(profile_dict: dict) -> None:
    profile_dict["tables"]["fontSizePt"] = {}
    text = messages(profile_dict)
    assert re.search(r"tables\.fontSizePt: ", text)
    assert re.search(r"[А-Яа-я]", text)


def test_several_errors_are_reported_together(profile_dict: dict) -> None:
    profile_dict["page"]["orientation"] = "sideways"
    profile_dict["body"]["alignment"] = "по ширине"
    profile_dict["tables"]["headerRowAlign"] = "середина"
    assert len(validate_profile(profile_dict)) >= 3
