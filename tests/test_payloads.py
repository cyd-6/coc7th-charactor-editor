"""Draft numeric errors stay visible at validation and export boundaries."""

import json

import pytest
from fastapi.testclient import TestClient

from app import app, get_catalog
from coc7_card.web import build_draft


@pytest.fixture(scope="module")
def client():
    with TestClient(app) as connection:
        yield connection


@pytest.fixture
def payload():
    occupation = next(item for item in get_catalog().occupations if item.credit_min == 0)
    return {
        "identity": {"name": "数值边界测试", "age": 30},
        "occupation_id": occupation.occupation_id,
        "skills": [{"template_slot": "F16", "interest_points": 0}],
        "experience": {"selection": "custom", "name": "测试经历", "skill_points": 0, "san_loss": 0},
    }


@pytest.mark.parametrize("endpoint", ["/api/validate", "/api/export/excel", "/api/export/pdf"])
@pytest.mark.parametrize("section,field", [("identity", "age"), ("skills", "interest_points"), ("experience", "san_loss")])
@pytest.mark.parametrize("value", ["abc", "NaN", "Infinity", {}])
def test_invalid_numeric_draft_is_rejected(client, payload, endpoint, section, field, value):
    target = payload["skills"][0] if section == "skills" else payload[section]
    target[field] = value
    if endpoint == "/api/validate":
        response = client.post(endpoint, json=payload)
    else:
        response = client.post(endpoint, data={"draft_json": json.dumps(payload)})
    assert response.status_code == 422
    assert response.json() == {"detail": "数值格式无效，请检查年龄和点数等数字字段。"}


@pytest.mark.parametrize("value,age,points", [(None, 30, 0), ("", 30, 0), (32, 32, 32), ("32", 32, 32), ("32.9", 32, 32)])
def test_numeric_defaults_and_conversion_are_preserved(payload, value, age, points):
    payload["identity"]["age"] = value
    payload["skills"][0]["interest_points"] = value
    draft = build_draft(payload, get_catalog())
    assert draft.identity.age == age
    assert next(skill for skill in draft.skills if skill.template_slot == "F16").interest_points == points
