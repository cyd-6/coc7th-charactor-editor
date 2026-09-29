"""Import real template packages without modifying the user's source workbook."""

from __future__ import annotations

import base64
import copy
import hashlib
import io
import shutil
import zipfile

import pytest
from fastapi.testclient import TestClient
from lxml import etree as ET
from PIL import Image

from app import app, get_catalog
from coc7_card.exporters.excel import ExcelExporter
from coc7_card.exporters.excel_linux import LinuxExcelExporter
from coc7_card.exporters.xlsx_template import TemplateWorkbook, tag, worksheet_paths
from coc7_card.importers.excel import ExcelImportError, import_investigator
from coc7_card.rules import RuleEngine
from coc7_card.web import DraftPayloadError, build_draft


@pytest.fixture(scope="module")
def catalog():
    return get_catalog()


@pytest.fixture(scope="module")
def import_client():
    with TestClient(app) as client:
        yield client


@pytest.fixture(scope="module")
def investigator_payload(catalog):
    occupation = next(item for item in catalog.occupations if item.credit_min == 0)
    return {
        "version": 3,
        "identity": {
            "name": "林知秋", "player": "读表测试", "age": 32, "story_year": 1925,
            "era": "1920s", "current_date": "1925-04-12", "gender": "女",
            "birthplace": "上海", "residence": "阿卡姆",
        },
        "attributes": {"STR": 60, "CON": 65, "SIZ": 55, "DEX": 70, "APP": 50,
                       "INT": 75, "POW": 60, "EDU": 70, "Luck": 55},
        "occupation_id": occupation.occupation_id,
        "skills": [
            {"template_slot": "F16", "interest_points": 2, "extra_final": 12, "experience_points": 17},
            {"template_slot": "AB16", "interest_points": 3, "extra_final": 7, "experience_points": 11},
        ],
        "experience": {"selection": "custom", "name": "档案调查", "skill_points": 40,
                       "san_loss": 3, "notes": "曾在古档案馆任职。"},
        "background": {
            "appearance": "短发，灰色风衣", "beliefs": "寻找事实", "significant_people": "旧友",
            "meaningful_places": "图书馆", "treasured_possessions": "旧信", "traits": "谨慎",
            "scars": "左手浅疤", "phobias_manias": "惧高", "personal_story": "调查一封来自故友的信。",
            "key_connection": "思想与信念、重要之人",
        },
        "assets": {"currency_code": "GBP", "usd_cash": "123.45", "usd_assets": "0", "usd_spending": "",
                   "living_standard": "朴素", "asset_description": "一笔积蓄", "vehicles": "自行车",
                   "residence": "租住公寓", "luxuries": "怀表", "securities": "无", "other": "古书"},
        "weapons": [{"name": "小刀", "category": "近战", "skill": "斗殴", "damage": "1D4+DB",
                     "range": "接触", "attacks": "1", "ammo": "无", "malfunction": "无"}],
        "inventory": [{"name": "调查笔记", "status": "携带", "location": "外套口袋", "backpack_slot": "1"}],
    }


def _write_character(catalog, payload, directory):
    """Use the production writer, including formulas and native workbook features."""
    book = TemplateWorkbook(catalog.template_path)
    output = directory / "fixture.xlsx"
    try:
        writer = ExcelExporter(catalog)
        writer._patch_template(book)
        writer._write_character(book, build_draft(payload, catalog), None, directory)
        book.save(output)
    finally:
        book.close()
    return output.read_bytes()


@pytest.fixture(scope="module")
def workbook_bytes(catalog, investigator_payload, tmp_path_factory):
    return _write_character(catalog, investigator_payload, tmp_path_factory.mktemp("import-fixture"))


def _replace_parts(data, replacements=None, edit=None):
    with zipfile.ZipFile(io.BytesIO(data)) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    parts.update(replacements or {})
    if edit:
        edit(parts)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as destination:
        for name, content in parts.items():
            destination.writestr(name, content)
    return output.getvalue()


def _replace_cell(parts, sheet, address, *, value=None, formula=None, cached=None):
    path = worksheet_paths(parts)[sheet]
    root = ET.fromstring(parts[path])
    cell = next(node for node in root.iter(tag("c")) if node.get("r") == address)
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)
    if formula is not None:
        ET.SubElement(cell, tag("f")).text = formula
        if cached is not None:
            ET.SubElement(cell, tag("v")).text = str(cached)
    elif isinstance(value, str):
        cell.set("t", "inlineStr")
        ET.SubElement(ET.SubElement(cell, tag("is")), tag("t")).text = value
    elif value is not None:
        ET.SubElement(cell, tag("v")).text = str(value)
    parts[path] = ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def _api_payload(state):
    """Apply the UI's occupation selection before backend validation/re-export."""
    payload = copy.deepcopy(state)
    if state["occupation_mode"] == "custom":
        custom = state["custom_occupation"]
        formula = (f"EDU*2+{custom['secondary']}*2"
                   if custom["point_formula_kind"] == "secondary" else "EDU*4")
        payload["occupation"] = {**custom, "is_custom": True, "point_formula": formula}
    return payload


def test_template_roundtrip_preserves_inputs_and_source(catalog, investigator_payload, workbook_bytes):
    original_hash = hashlib.sha256(catalog.template_path.read_bytes()).hexdigest()
    uploaded_hash = hashlib.sha256(workbook_bytes).hexdigest()
    result = import_investigator(workbook_bytes, catalog)
    state = result["draft"]
    assert state["version"] == 3
    for key, value in investigator_payload["identity"].items():
        assert state["identity"][key] == value
    assert state["attributes"] == investigator_payload["attributes"]
    assert state["occupation_mode"] == "catalog"
    assert state["occupation_id"] == investigator_payload["occupation_id"]
    assert state["background"] == {**investigator_payload["background"], "contacts_notes": ""}
    for section in ("weapons", "inventory"):
        assert len(state[section]) == len(investigator_payload[section])
        for actual, expected in zip(state[section], investigator_payload[section]):
            assert all(actual[key] == value for key, value in expected.items())
    assert state["experience"]["selection"] == "custom"
    for key in ("name", "skill_points", "san_loss", "notes"):
        assert state["experience"][key] == investigator_payload["experience"][key]
    rows = {row["template_slot"]: row for row in state["skills"]}
    for expected in investigator_payload["skills"]:
        for key, value in expected.items():
            assert rows[expected["template_slot"]][key] == value
    assert state["assets"]["currency_code"] == "GBP"
    assert state["assets"]["exchange_year"] == 1925
    assert state["assets"]["usd_spending"] == ""  # Auto input is not a frozen formula cache.
    assert float(state["assets"]["usd_cash"]) == pytest.approx(123.45)
    assert state["assets"]["usd_assets"] != "" and float(state["assets"]["usd_assets"]) == 0  # Explicit zero is not auto.
    for key in ("living_standard", "asset_description", "vehicles", "residence", "luxuries", "securities", "other"):
        assert state["assets"][key] == investigator_payload["assets"][key]
    rebuilt = build_draft(_api_payload(state), catalog)
    expected = build_draft(investigator_payload, catalog)
    assert rebuilt.assets.cash == expected.assets.cash
    assert rebuilt.assets.other_assets == expected.assets.other_assets == "0.00"
    assert [(s.template_slot, s.final_value) for s in rebuilt.skills] == [(s.template_slot, s.final_value) for s in expected.skills]
    assert result["summary"]["name"] == "林知秋"
    assert result["summary"]["weapon_count"] == result["summary"]["inventory_count"] == 1
    assert all(isinstance(warning, str) for warning in result["warnings"])
    assert result["portrait"] is None
    assert hashlib.sha256(catalog.template_path.read_bytes()).hexdigest() == original_hash
    assert hashlib.sha256(workbook_bytes).hexdigest() == uploaded_hash


def test_custom_occupation_and_literal_text_roundtrip(catalog, investigator_payload, tmp_path):
    payload = copy.deepcopy(investigator_payload)
    payload["identity"]["name"] = '=HYPERLINK("https://example.invalid","literal")'
    payload["background"]["personal_story"] = '<img src=x onerror="alert(1)">'
    payload["occupation"] = {"is_custom": True, "name": "民俗研究员", "credit_min": 0,
                             "credit_max": 99, "point_formula": "EDU*2+INT*2"}
    payload["custom_skill_slots"] = [skill.template_slot for skill in catalog.skills[:8]]
    state = import_investigator(_write_character(catalog, payload, tmp_path), catalog)["draft"]
    assert state["occupation_mode"] == "custom"
    assert state["custom_occupation"]["name"] == "民俗研究员"
    assert state["custom_occupation"]["point_formula_kind"] == "secondary"
    assert state["custom_occupation"]["secondary"] == "INT"
    assert state["identity"]["name"] == payload["identity"]["name"]
    assert state["background"]["personal_story"] == payload["background"]["personal_story"]
    assert set(payload["custom_skill_slots"]).issubset(state["custom_skill_slots"])
    draft = build_draft(_api_payload(state), catalog)
    assert draft.occupation.is_custom
    assert draft.occupation.point_formula == "EDU*2+INT*2"


@pytest.mark.parametrize("legacy_credit_slot", [False, True])
def test_custom_occupation_keeps_eight_unallocated_skills(catalog, investigator_payload, tmp_path, legacy_credit_slot):
    # All chosen rows follow credit in the card order, exposing the historical
    # writer bug where automatic credit occupied one of the eight custom slots.
    slots = ["F28", "AB28", "F29", "AB29", "F30", "AB30", "F31", "AB31"]
    payload = copy.deepcopy(investigator_payload)
    payload["occupation"] = {"is_custom": True, "name": "自定义调查职业", "credit_min": 0,
                             "credit_max": 99, "point_formula": "EDU*4"}
    payload["custom_skill_slots"] = slots
    payload["skills"] = []
    data = _write_character(catalog, payload, tmp_path)
    if legacy_credit_slot:
        names = {skill.template_slot: skill.name.rstrip("：:") for skill in catalog.skills}

        def restore_old_bug(parts):
            for row, name in enumerate(["信用评级", *(names[slot] for slot in slots[:7])], start=3):
                _replace_cell(parts, "职业列表", f"I{row}", value=name)

        data = _replace_parts(data, edit=restore_old_bug)
    result = import_investigator(data, catalog)
    state = result["draft"]
    assert state["occupation_mode"] == "custom"
    assert "F26" not in state["custom_skill_slots"]
    assert set(state["custom_skill_slots"]) == set(slots[:7] if legacy_credit_slot else slots)
    assert all(not any(row[key] for key in ("occupation_points", "interest_points", "extra_final", "experience_points"))
               for row in state["skills"])
    if legacy_credit_slot:
        assert any("信用评级" in warning and "遗漏" in warning for warning in result["warnings"])
    else:
        rebuilt = build_draft(_api_payload(state), catalog)
        assert {skill.template_slot for skill in rebuilt.skills if skill.selected_occupation} == set(slots) | {"F26"}


def test_formula_without_cache_is_not_executed_or_silently_imported(catalog, workbook_bytes):
    changed = _replace_parts(workbook_bytes, edit=lambda parts: _replace_cell(parts, "人物卡", "E3", formula='WEBSERVICE("https://example.invalid")'))
    result = import_investigator(changed, catalog)
    assert result["draft"]["identity"]["name"] == ""
    assert any("E3" in warning and ("缓存" in warning or "重算" in warning) for warning in result["warnings"])


def test_literal_formula_cache_is_read_as_saved_value(catalog, workbook_bytes):
    changed = _replace_parts(workbook_bytes, edit=lambda parts: _replace_cell(parts, "人物卡", "U3", formula="50+15", cached=65))
    result = import_investigator(changed, catalog)
    assert result["draft"]["attributes"]["STR"] == 65


def test_free_text_era_is_preserved_and_invalid_date_is_flagged(catalog, workbook_bytes):
    def edit(parts):
        _replace_cell(parts, "人物卡", "M4", value="蒸汽朋克架空时代")
        _replace_cell(parts, "人物卡", "J8", value="2月")
        _replace_cell(parts, "人物卡", "L8", value="30日")
    result = import_investigator(_replace_parts(workbook_bytes, edit=edit), catalog)
    assert result["draft"]["identity"]["era"] == "蒸汽朋克架空时代"
    assert result["summary"]["era"] == "蒸汽朋克架空时代"
    assert result["draft"]["identity"]["current_date"] == ""
    assert any("日期无效" in warning for warning in result["warnings"])


def test_extreme_numeric_cells_are_flagged_without_crashing(catalog, workbook_bytes):
    def edit(parts):
        _replace_cell(parts, "人物卡", "U3", formula="1E999999999", cached="1e999999999")
        _replace_cell(parts, "人物卡", "G8", value=999999999999999)
    result = import_investigator(_replace_parts(workbook_bytes, edit=edit), catalog)
    assert result["draft"]["attributes"]["STR"] == 50
    assert result["draft"]["identity"]["current_date"] == ""
    assert any("U3" in warning for warning in result["warnings"])
    assert any("日期无效" in warning for warning in result["warnings"])


def test_summary_counts_all_imported_skills_including_unallocated_rows(catalog, workbook_bytes):
    result = import_investigator(workbook_bytes, catalog)
    assert result["summary"]["skill_count"] == len(result["draft"]["skills"]) == len(catalog.skills) == 67
    assert any(not any(row[key] for key in ("occupation_points", "interest_points", "extra_final", "experience_points"))
               for row in result["draft"]["skills"])


def test_played_hp_mp_and_san_are_reported_before_build_recalculation(catalog, workbook_bytes):
    def edit(parts):
        for cell, value in (("E10", 5), ("W10", 2), ("N10", 40)):
            _replace_cell(parts, "人物卡", cell, value=value)
    result = import_investigator(_replace_parts(workbook_bytes, edit=edit), catalog)
    warning = next(item for item in result["warnings"] if "当前游玩状态" in item)
    assert "HP 原表 5／建卡值 12" in warning
    assert "MP 原表 2／建卡值 12" in warning
    assert "SAN 原表 40／建卡值 57" in warning
    assert "保留原表" in warning


def test_import_preserves_manual_skill_base(catalog, workbook_bytes):
    cell = catalog.skill_cell("F16").base_cell
    changed = _replace_parts(workbook_bytes, edit=lambda parts: _replace_cell(parts, "人物卡", cell, value=23))
    state = import_investigator(changed, catalog)["draft"]
    row = next(row for row in state["skills"] if row["template_slot"] == "F16")
    assert row["base_override"] == 23
    draft = build_draft(_api_payload(state), catalog)
    assert next(skill for skill in draft.skills if skill.template_slot == "F16").final_value == 23 + 2 + 12 + 17


@pytest.mark.parametrize("field,cell_field", [("occupation_points", "occupation_cell"), ("interest_points", "interest_cell"), ("extra_final", "extra_cell")])
def test_negative_imported_points_require_correction(catalog, workbook_bytes, field, cell_field):
    cell = getattr(catalog.skill_cell("F16"), cell_field)
    changed = _replace_parts(workbook_bytes, edit=lambda parts: _replace_cell(parts, "人物卡", cell, value=-3))
    state = import_investigator(changed, catalog)["draft"]
    assert next(row for row in state["skills"] if row["template_slot"] == "F16")[field] == -3
    draft = build_draft(_api_payload(state), catalog)
    assert getattr(next(skill for skill in draft.skills if skill.template_slot == "F16"), field) == -3
    validation = RuleEngine.validate(draft)
    assert not validation.can_export
    assert any(issue.code == "SKILL_NEGATIVE" and issue.field == "F16" for issue in validation.errors)


def test_legacy_growth_is_not_counted_twice(catalog, workbook_bytes):
    def make_legacy(parts):
        _replace_cell(parts, "人物卡", "M15", value=None)
        _replace_cell(parts, "人物卡", "AI15", value=None)
        _replace_cell(parts, "人物卡", "L16", value=29)
        _replace_cell(parts, "人物卡", "AH16", value=18)
        _replace_cell(parts, "人物卡", "M16", value=None)
        _replace_cell(parts, "人物卡", "AI16", value=None)
    result = import_investigator(_replace_parts(workbook_bytes, edit=make_legacy), catalog)
    rows = {row["template_slot"]: row for row in result["draft"]["skills"]}
    assert rows["F16"]["extra_final"] == 29
    assert rows["AB16"]["extra_final"] == 18
    assert all(row["experience_points"] == 0 for row in rows.values())
    assert any("旧版" in warning and "成长点数" in warning for warning in result["warnings"])


@pytest.mark.parametrize("mark_choices", [False, True])
def test_catalog_choice_groups_are_recovered_or_flagged(catalog, workbook_bytes, mark_choices):
    occupation = next(item for item in catalog.occupations if item.choice_groups)
    group = occupation.choice_groups[0]
    choices = group.candidates[:group.required_count]

    def select_occupation(parts):
        _replace_cell(parts, "人物卡", "M5", value=occupation.occupation_id)
        _replace_cell(parts, "人物卡", "E5", value=occupation.name)
        if mark_choices:
            for token in choices:
                definition = next(item for item in catalog.skills if item.name.rstrip("：:") == token)
                column = "D" if definition.template_slot.startswith("F") else "Z"
                address = column + "".join(filter(str.isdigit, definition.template_slot))
                _replace_cell(parts, "人物卡", address, value="★")

    result = import_investigator(_replace_parts(workbook_bytes, edit=select_occupation), catalog)
    assert result["draft"]["occupation_mode"] == "catalog"
    assert result["draft"]["occupation_id"] == occupation.occupation_id
    if mark_choices:
        assert result["draft"]["group_choices"][group.marker] == list(choices)
    else:
        assert result["draft"]["group_choices"][group.marker] == []
        assert any(group.label in warning and "补选" in warning for warning in result["warnings"])


def test_missing_currency_is_not_silently_replaced(catalog, workbook_bytes):
    changed = _replace_parts(workbook_bytes, edit=lambda parts: _replace_cell(parts, "货币汇率", "K2", value="不存在的历史货币"))
    result = import_investigator(changed, catalog)
    assert result["draft"]["assets"]["currency_code"] == ""
    assert float(result["draft"]["assets"]["usd_cash"]) == pytest.approx(123.45)
    assert any("重新选择币种" in warning for warning in result["warnings"])
    with pytest.raises(DraftPayloadError):
        build_draft(_api_payload(result["draft"]), catalog)


@pytest.mark.parametrize("value", ["NaN", -1, 1000, 4.5, float("inf")])
def test_skill_base_override_rejects_invalid_values(catalog, investigator_payload, value):
    payload = copy.deepcopy(investigator_payload)
    payload["skills"][0]["base_override"] = value
    with pytest.raises(DraftPayloadError):
        build_draft(payload, catalog)


@pytest.mark.parametrize("content", [b"", b"not an Excel workbook", b"PK\x03\x04truncated archive"])
def test_non_workbook_files_are_rejected(catalog, content):
    with pytest.raises(ExcelImportError):
        import_investigator(content, catalog)


@pytest.mark.parametrize("part", ["xl/workbook.xml", "xl/worksheets/sheet1.xml"])
def test_malformed_xml_has_an_import_error(catalog, workbook_bytes, part):
    with pytest.raises(ExcelImportError):
        import_investigator(_replace_parts(workbook_bytes, {part: b"<broken"}), catalog)


def test_unrelated_excel_workbook_is_rejected(catalog, workbook_bytes):
    def rename_card(parts):
        root = ET.fromstring(parts["xl/workbook.xml"])
        for sheet in root.find(tag("sheets")):
            if sheet.get("name") == "人物卡":
                sheet.set("name", "销售数据")
        parts["xl/workbook.xml"] = ET.tostring(root)
    with pytest.raises(ExcelImportError):
        import_investigator(_replace_parts(workbook_bytes, edit=rename_card), catalog)


def test_compressed_expansion_limit_is_checked(catalog):
    # A tiny upload must not be allowed to expand into an arbitrarily large package.
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("huge.xml", b"0" * (80 * 1024 * 1024 + 1))
    assert len(buffer.getvalue()) < 1024 * 1024
    with pytest.raises(ExcelImportError, match="解压"):
        import_investigator(buffer.getvalue(), catalog)


def test_archive_member_limit_is_checked(catalog):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for index in range(2001):
            archive.writestr(f"entry-{index}", b"")
    with pytest.raises(ExcelImportError, match="解压"):
        import_investigator(buffer.getvalue(), catalog)


@pytest.mark.parametrize("extra_part", ["../outside.txt", "xl/vbaProject.bin"])
def test_unsafe_package_parts_are_rejected(catalog, workbook_bytes, extra_part):
    with pytest.raises(ExcelImportError):
        import_investigator(_replace_parts(workbook_bytes, {extra_part: b"unsupported"}), catalog)


def test_entity_declaration_cannot_read_local_files(catalog, workbook_bytes, tmp_path):
    secret = tmp_path / "unrelated.txt"
    secret.write_text("PRIVATE_SENTINEL")
    xml = ('<?xml version="1.0"?><!DOCTYPE workbook [<!ENTITY secret SYSTEM "'
           + secret.as_uri() + '">]><workbook>&secret;</workbook>').encode()
    with pytest.raises(ExcelImportError) as error:
        import_investigator(_replace_parts(workbook_bytes, {"xl/workbook.xml": xml}), catalog)
    assert "PRIVATE_SENTINEL" not in str(error.value)


def test_api_import_is_read_only_and_returns_preview(import_client, workbook_bytes):
    response = import_client.post("/api/import/excel", files={"workbook": ("调查员.XLSX", workbook_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert response.status_code == 200, response.text
    result = response.json()
    assert result["draft"]["identity"]["name"] == result["summary"]["name"] == "林知秋"
    assert {"draft", "warnings", "summary", "portrait"} <= result.keys()
    assert import_client.get("/api/health").json()["status"] == "ok"


@pytest.mark.parametrize("filename", ["card.xls", "card.xlsm", "card.csv", "card"])
def test_api_only_accepts_xlsx(import_client, workbook_bytes, filename):
    response = import_client.post("/api/import/excel", files={"workbook": (filename, workbook_bytes)})
    assert response.status_code == 422
    assert "xlsx" in response.json()["detail"].lower()


def test_api_rejects_oversize_before_parsing(import_client):
    response = import_client.post("/api/import/excel", files={"workbook": ("large.xlsx", b"x" * (12 * 1024 * 1024 + 1))})
    assert response.status_code == 413
    assert "12" in response.json()["detail"]


@pytest.mark.skipif(not (shutil.which("libreoffice") or shutil.which("soffice")), reason="LibreOffice Calc is required")
def test_linux_export_import_roundtrip_with_portrait(catalog, investigator_payload):
    portrait = io.BytesIO()
    Image.new("RGB", (80, 120), "#4c7084").save(portrait, format="PNG")
    original = build_draft(investigator_payload, catalog)
    result = LinuxExcelExporter(catalog).export(original, portrait.getvalue())
    imported = import_investigator(result.data, catalog)
    rebuilt = build_draft(_api_payload(imported["draft"]), catalog)
    assert rebuilt.identity == original.identity
    assert rebuilt.attributes == original.attributes
    assert [(row.template_slot, row.final_value) for row in rebuilt.skills] == [(row.template_slot, row.final_value) for row in original.skills]
    assert rebuilt.assets.currency_code == original.assets.currency_code
    assert rebuilt.assets.cash == original.assets.cash
    assert rebuilt.assets.auto_spending and not rebuilt.assets.auto_assets
    assert rebuilt.experience.name == original.experience.name
    assert rebuilt.experience.skill_points == original.experience.skill_points
    assert imported["portrait"]["mime_type"] == "image/png"
    with Image.open(io.BytesIO(base64.b64decode(imported["portrait"]["data_base64"]))) as image:
        assert image.size == (80, 120)
        assert image.convert("RGB").getpixel((40, 60)) == (76, 112, 132)
