"""The 200-branch boundary never truncates a draft or an imported workbook."""

import copy
from dataclasses import replace
import hashlib
import json
import shutil

import pytest
from fastapi.testclient import TestClient

from app import FONT_PATH, app, get_catalog
from coc7_card.exporters.excel import ExcelExporter, ExcelExportError
from coc7_card.exporters.excel_linux import LinuxExcelExporter
from coc7_card.exporters.pdf import PdfExporter, PdfExportError
from coc7_card.importers.excel import ExcelImportError, import_investigator
from coc7_card.rules import RuleEngine
from coc7_card.skill_specializations import is_branch_slot
from coc7_card.web import DraftPayloadError, build_draft, catalog_payload
from test_branch_exports import api_payload, assert_branch_values, branch, replace_cell, write_fixture


@pytest.fixture(scope="module")
def catalog():
    return get_catalog()


@pytest.fixture(scope="module")
def maximum_payload(catalog):
    return {
        "identity": {"name": "分支数量边界验证", "age": 30},
        "occupation": {"is_custom": True, "name": "边界验证职业", "credit_min": 0,
                       "credit_max": 99, "point_formula": "EDU*4"},
        "skills": [{"key": skill.template_slot, "template_slot": skill.template_slot}
                   for skill in catalog.skills]
        + [branch(index, "科学", f"测试领域{index}", interest_points=1 if index <= 5 else 0,
                  extra_final=index % 3) for index in range(1, 201)],
    }


@pytest.fixture(scope="module")
def maximum_draft(catalog, maximum_payload):
    return build_draft(maximum_payload, catalog)


@pytest.fixture
def oversized_draft(maximum_draft):
    draft = copy.deepcopy(maximum_draft)
    slot = branch(201)["template_slot"]
    draft.skills.append(replace(draft.skills[-1], key=slot, template_slot=slot,
                                specialization="测试领域201"))
    return draft


@pytest.fixture(scope="module")
def maximum_workbook(catalog, maximum_payload, tmp_path_factory):
    return write_fixture(catalog, maximum_payload, tmp_path_factory.mktemp("branch-limit"))


def test_two_hundred_branches_do_not_count_original_catalog_skills(catalog, maximum_payload, maximum_draft):
    assert catalog_payload(catalog)["meta"]["max_skill_branches"] == 200
    assert len(maximum_payload["skills"]) == len(maximum_draft.skills) == 67 + 200
    assert sum(is_branch_slot(skill.template_slot) for skill in maximum_draft.skills) == 200
    assert RuleEngine.validate(maximum_draft).can_export
    assert [skill.extra_final for skill in maximum_draft.skills[-200:]] == [index % 3 for index in range(1, 201)]


def test_extra_branch_rejected_without_mutating_payload(catalog, maximum_payload):
    payload = copy.deepcopy(maximum_payload)
    payload["skills"].append(branch(201, "科学", "测试领域201", interest_points=7, extra_final=9))
    before = copy.deepcopy(payload)
    with pytest.raises(DraftPayloadError, match="新增技能分支不能超过 200 个"):
        build_draft(payload, catalog)
    assert payload == before
    assert len(payload["skills"]) == 67 + 201


def test_direct_rule_validation_rejects_extra_branch_even_with_override(oversized_draft):
    oversized_draft.nonstandard_override = True
    report = RuleEngine.validate(oversized_draft)
    assert not report.can_export
    assert [issue.code for issue in report.errors] == ["SKILL_BRANCH_LIMIT"]
    assert len(oversized_draft.skills) == 67 + 201


@pytest.mark.parametrize("kind", ["windows", "linux", "pdf"])
def test_public_exporters_reject_direct_oversized_draft(catalog, oversized_draft, kind):
    exporter = (PdfExporter(FONT_PATH) if kind == "pdf"
                else LinuxExcelExporter(catalog) if kind == "linux" else ExcelExporter(catalog))
    with pytest.raises((ExcelExportError, PdfExportError), match="新增技能分支不能超过 200 个"):
        exporter.export(oversized_draft)
    assert len(oversized_draft.skills) == 67 + 201


@pytest.mark.parametrize("endpoint", ["/api/validate", "/api/export/excel", "/api/export/pdf"])
def test_api_reports_branch_limit(maximum_payload, endpoint):
    payload = copy.deepcopy(maximum_payload)
    payload["skills"].append(branch(201, "科学", "测试领域201"))
    with TestClient(app) as client:
        response = (client.post(endpoint, json=payload) if endpoint == "/api/validate"
                    else client.post(endpoint, data={"draft_json": json.dumps(payload)}))
    assert response.status_code == 422
    assert "新增技能分支不能超过 200 个" in response.json()["detail"]


def test_two_hundred_branch_workbook_import_preserves_all_values(catalog, maximum_draft, maximum_workbook):
    state = import_investigator(maximum_workbook, catalog)["draft"]
    assert len(state["skills"]) == 67 + 200
    assert_branch_values(maximum_draft, build_draft(api_payload(state), catalog))


def test_workbook_declaring_extra_branch_is_rejected_without_changing_source(catalog, maximum_workbook):
    data = replace_cell(maximum_workbook, "AP153", value=201)
    before = hashlib.sha256(data).digest()
    with pytest.raises(ExcelImportError, match="新增技能分支不能超过 200 个"):
        import_investigator(data, catalog)
    assert hashlib.sha256(data).digest() == before


@pytest.mark.skipif(not shutil.which("libreoffice") and not shutil.which("soffice"), reason="LibreOffice Calc is required")
def test_real_export_at_two_hundred_branches_roundtrips(catalog, maximum_draft):
    exported = LinuxExcelExporter(catalog).export(maximum_draft)
    state = import_investigator(exported.data, catalog)["draft"]
    assert len(state["skills"]) == 67 + 200
    assert_branch_values(maximum_draft, build_draft(api_payload(state), catalog))
