"""Additional branches retain independent identity, bases, budgets and choices."""

import json

import pytest
from fastapi.testclient import TestClient

from app import app, get_catalog
from coc7_card.rules import RuleEngine
from coc7_card.skill_specializations import (
    GROUP_BASE_VALUES, is_branch_slot, matching_branch_slot, validate_branch_identity,
)
from coc7_card.web import DraftPayloadError, build_draft, catalog_payload


def branch(index=1, name="射击", specialization="冲锋枪", **points):
    slot = f"branch-{index:032x}"
    return {"key": slot, "template_slot": slot, "name": name, "specialization": specialization, **points}


@pytest.fixture(scope="module")
def catalog():
    return get_catalog()


@pytest.fixture
def payload(catalog):
    occupation = next(item for item in catalog.occupations if item.credit_min == 0)
    return {"identity": {"name": "多分支测试", "age": 30}, "occupation_id": occupation.occupation_id,
            "skills": []}


def test_group_catalog_exposes_independent_options_and_fallbacks(catalog):
    result = catalog_payload(catalog)
    groups = {item["name"]: item for item in result["skill_groups"]}
    assert set(groups) == set(GROUP_BASE_VALUES)
    assert {name: item["base_value"] for name, item in groups.items()} == GROUP_BASE_VALUES
    assert all(item["group"] == name for name, item in groups.items())
    assert {"name": "冲锋枪", "base_value": 15, "aliases": []} in groups["射击"]["specialization_options"]
    assert len(result["skills"]) == len(catalog.skills) == 67


def test_all_firearm_branches_can_coexist_beyond_template_slots(catalog, payload):
    names = ["步枪/霰弹枪", "冲锋枪", "弓术", "喷射器", "机枪", "重武器"]
    payload["skills"] = [branch(index, specialization=name, interest_points=index)
                         for index, name in enumerate(names, 1)]
    draft = build_draft(payload, catalog)
    additional = [skill for skill in draft.skills if is_branch_slot(skill.template_slot)]
    assert len(draft.skills) == len(catalog.skills) + 6
    assert [skill.specialization for skill in additional] == names
    assert [skill.base_value for skill in additional] == [25, 15, 15, 10, 10, 10]
    assert [skill.final_value for skill in additional] == [26, 17, 18, 14, 15, 16]
    assert [skill.key for skill in additional] == [row["key"] for row in payload["skills"]]
    assert RuleEngine.validate(draft).can_export


@pytest.mark.parametrize("group,expected", [("技艺", 5), ("生存", 10), ("外语", 1), ("科学", 1)])
def test_custom_branch_uses_group_fallback_without_losing_points(catalog, payload, group, expected):
    payload["skills"] = [branch(name=group, specialization="自定义分支", interest_points=9, extra_final=4)]
    skill = build_draft(payload, catalog).skills[-1]
    assert (skill.name, skill.specialization, skill.base_value, skill.interest_points,
            skill.extra_final, skill.final_value) == (group, "自定义分支", expected, 9, 4, expected + 13)


@pytest.mark.parametrize("override", [0, 7, 99])
def test_branch_base_override_has_priority(catalog, payload, override):
    payload["skills"] = [branch(base_override=override)]
    assert build_draft(payload, catalog).skills[-1].base_value == override


@pytest.mark.parametrize("override", [True, -1, 1000, 1.5, "15", float("inf")])
def test_branch_rejects_invalid_base_override(catalog, payload, override):
    payload["skills"] = [branch(base_override=override)]
    with pytest.raises(DraftPayloadError, match="0—999"):
        build_draft(payload, catalog)


@pytest.mark.parametrize("changes,message", [
    ({"template_slot": "unknown", "key": "unknown"}, "标识无效"),
    ({"template_slot": "branch-" + "A" * 32, "key": "branch-" + "A" * 32}, "标识无效"),
    ({"key": "branch-" + "2" * 32}, "标识不一致"),
    ({"name": "射击①"}, "所属大项无效"),
    ({"name": "不存在的大项"}, "所属大项无效"),
    ({"specialization": " "}, "填写分支"),
    ({"specialization": None}, "填写分支"),
    ({"specialization": "甲" * 81}, "80"),
])
def test_invalid_branch_is_reported_instead_of_discarded(catalog, payload, changes, message):
    payload["skills"] = [{**branch(), **changes}]
    with pytest.raises(DraftPayloadError, match=message):
        build_draft(payload, catalog)


def test_branch_identity_keeps_eighty_character_names():
    slot = branch()["template_slot"]
    assert validate_branch_identity(slot, " 科学 ", "甲" * 80) == (slot, "科学", "甲" * 80)


def test_duplicate_branch_identifier_is_not_overwritten(catalog, payload):
    payload["skills"] = [branch(), branch(specialization="机枪")]
    with pytest.raises(DraftPayloadError, match="标识重复"):
        build_draft(payload, catalog)


@pytest.mark.parametrize("rows", [
    [branch(specialization="手枪")],
    [{"template_slot": "F39", "specialization": "步/霰"}, branch(specialization="步枪/霰弹枪")],
    [branch(1, specialization="来复"), branch(2, specialization="步枪／霰弹枪")],
    [branch(1, name="科学", specialization="Ｆｏｏ"), branch(2, name="科学", specialization="foo")],
])
def test_duplicate_branch_names_include_aliases_and_existing_slots(catalog, payload, rows):
    payload["skills"] = rows
    with pytest.raises(DraftPayloadError, match="已存在"):
        build_draft(payload, catalog)


def test_legacy_slot_duplicates_still_build(catalog, payload):
    payload["skills"] = [{"template_slot": slot, "specialization": "冲锋枪"} for slot in ("F39", "F40")]
    draft = build_draft(payload, catalog)
    assert len(draft.skills) == len(catalog.skills)


def test_each_branch_counts_as_one_custom_occupation_skill(catalog, payload):
    payload["occupation"] = {"is_custom": True, "name": "多分支职业", "credit_min": 0,
                             "credit_max": 99, "point_formula": "EDU*4"}
    payload["skills"] = [branch(index, name="科学", specialization=f"领域{index}", occupation_points=2)
                         for index in range(1, 10)]
    payload["custom_skill_slots"] = [row["template_slot"] for row in payload["skills"][:8]]
    draft = build_draft(payload, catalog)
    assert [skill.selected_occupation for skill in draft.skills[-9:]] == [True] * 8 + [False]
    assert any(issue.code == "SKILL_NON_OCCUPATION" and issue.field == payload["skills"][-1]["template_slot"]
               for issue in RuleEngine.validate(draft).errors)
    payload["custom_skill_slots"].append(payload["skills"][-1]["template_slot"])
    with pytest.raises(DraftPayloadError, match="最多选择 8 项"):
        build_draft(payload, catalog)


def test_free_choice_selects_only_specific_branch_and_counts_each(catalog, payload):
    occupation = next(item for item in catalog.occupations if item.free_choices == 1)
    payload["occupation_id"] = occupation.occupation_id
    payload["skills"] = [branch(1, name="科学", specialization="自定义领域甲", occupation_points=2),
                         branch(2, name="科学", specialization="自定义领域乙", occupation_points=3)]
    payload["free_skill_choices"] = [payload["skills"][0]["template_slot"]]
    draft = build_draft(payload, catalog)
    assert [skill.selected_occupation for skill in draft.skills[-2:]] == [True, False]
    assert not any(issue.code == "OCCUPATION_FREE_COUNT" for issue in RuleEngine.validate(draft).errors)
    payload["free_skill_choices"].append(payload["skills"][1]["template_slot"])
    draft = build_draft(payload, catalog)
    assert any(issue.code == "OCCUPATION_FREE_COUNT" for issue in RuleEngine.validate(draft).errors)


def test_existing_occupation_group_does_not_authorize_every_new_branch(catalog, payload):
    payload["occupation_id"] = next(item.occupation_id for item in catalog.occupations if item.name == "士兵")
    payload["skills"] = [branch(specialization="机枪", occupation_points=4)]
    draft = build_draft(payload, catalog)
    assert not draft.skills[-1].selected_occupation
    assert any(issue.code == "SKILL_NON_OCCUPATION" and issue.field == draft.skills[-1].template_slot
               for issue in RuleEngine.validate(draft).errors)


def test_architect_reuses_existing_mathematics_branch_without_a_duplicate(catalog, payload):
    payload["occupation_id"] = next(item.occupation_id for item in catalog.occupations if item.name == "建筑师")
    mathematics = branch(1, name="科学", specialization="数学", occupation_points=15, interest_points=3)
    other = branch(2, name="科学", specialization="密码学")
    payload["skills"] = [{"template_slot": "AB31", "specialization": "物理学"}, mathematics, other]
    draft = build_draft(payload, catalog)
    skills = {skill.template_slot: skill for skill in draft.skills}
    assert skills[mathematics["template_slot"]].selected_occupation
    assert skills[mathematics["template_slot"]].final_value == 28
    assert not skills["AB31"].selected_occupation
    assert skills["AB31"].specialization == "物理学"
    assert not skills[other["template_slot"]].selected_occupation
    assert sum(skill.specialization == "数学" for skill in draft.skills) == 1
    assert not any(issue.code == "SKILL_NON_OCCUPATION" for issue in RuleEngine.validate(draft).errors)


def test_occupation_alias_reuses_only_matching_branch(catalog, payload):
    payload["occupation_id"] = next(item.occupation_id for item in catalog.occupations if item.name == "士兵")
    rifle = branch(1, specialization="步枪／霰弹枪", occupation_points=10)
    machine_gun = branch(2, specialization="机枪")
    payload["skills"] = [rifle, machine_gun]
    draft = build_draft(payload, catalog)
    skills = {skill.template_slot: skill for skill in draft.skills}
    assert skills[rifle["template_slot"]].selected_occupation
    assert skills[rifle["template_slot"]].final_value == 35
    assert not skills["F39"].selected_occupation
    assert not skills[machine_gun["template_slot"]].selected_occupation


@pytest.mark.parametrize("name,specialization", [("科学①", ""), ("科学①", "任一"),
                                                ("技艺①", "数学"), ("科学①", "化学")])
def test_branch_matching_requires_explicit_same_group_and_name(name, specialization):
    rows = [branch(name="科学", specialization="数学"), branch(2, name="科学", specialization="任一")]
    assert matching_branch_slot(name, specialization, rows) is None


@pytest.mark.parametrize("selection", ["custom_skill_slots", "free_skill_choices"])
def test_deleted_branch_cannot_remain_in_occupation_choices(catalog, payload, selection):
    if selection == "custom_skill_slots":
        payload["occupation"] = {"is_custom": True, "name": "多分支职业", "credit_min": 0,
                                 "credit_max": 99, "point_formula": "EDU*4"}
    payload[selection] = [branch()["template_slot"]]
    with pytest.raises(DraftPayloadError, match="不存在的技能"):
        build_draft(payload, catalog)


def test_normalized_known_name_gets_its_known_base(catalog, payload):
    payload["skills"] = [branch(specialization="步枪／霰弹枪")]
    skill = build_draft(payload, catalog).skills[-1]
    assert skill.specialization == "步枪／霰弹枪"
    assert skill.base_value == 25


@pytest.mark.parametrize("field,code", [("interest_points", "INTEREST_POINTS_TOTAL"),
                                       ("experience_points", "EXPERIENCE_BUDGET")])
def test_extra_branches_are_included_in_budgets(catalog, payload, field, code):
    payload["skills"] = [branch(index, name="科学", specialization=f"领域{index}", **{field: 20})
                         for index in range(1, 7)]
    report = RuleEngine.validate(build_draft(payload, catalog))
    assert any(issue.code == code for issue in report.errors)


def test_branch_occupation_points_use_shared_budget(catalog, payload):
    payload["occupation"] = {"is_custom": True, "name": "多分支职业", "credit_min": 0,
                             "credit_max": 99, "point_formula": "EDU*4"}
    payload["skills"] = [branch(index, name="科学", specialization=f"领域{index}", occupation_points=80)
                         for index in range(1, 4)]
    payload["custom_skill_slots"] = [row["template_slot"] for row in payload["skills"]]
    report = RuleEngine.validate(build_draft(payload, catalog))
    assert any(issue.code == "OCCUPATION_POINTS_TOTAL" for issue in report.errors)
    assert not any(issue.code == "SKILL_NON_OCCUPATION" for issue in report.errors)


def test_growth_does_not_use_experience_budget(catalog, payload):
    payload["skills"] = [branch(extra_final=12)]
    report = RuleEngine.validate(build_draft(payload, catalog))
    assert not any(issue.code == "EXPERIENCE_BUDGET" for issue in report.errors)


@pytest.mark.parametrize("endpoint", ["/api/validate", "/api/export/excel", "/api/export/pdf"])
def test_invalid_branch_is_rejected_at_api_boundaries(payload, endpoint):
    payload["skills"] = [branch(specialization="")]
    with TestClient(app) as client:
        response = (client.post(endpoint, json=payload) if endpoint == "/api/validate"
                    else client.post(endpoint, data={"draft_json": json.dumps(payload)}))
    assert response.status_code == 422
    assert "填写分支" in response.json()["detail"]
