"""Attribute-only imports reject ambiguity and never produce a full draft."""

from __future__ import annotations

import hashlib
import io
import zipfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from lxml import etree as ET

from app import app


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
HEADERS = ["力量 STR", "体质 CON", "体型 SIZ", "敏捷 DEX", "外貌 APP", "智力 INT", "意志 POW", "教育 EDU", "幸运 Luck"]
EXPECTED = dict(zip(["STR", "CON", "SIZ", "DEX", "APP", "INT", "POW", "EDU", "Luck"], [60, 65, 55, 70, 50, 75, 60, 70, 55]))


def simple_xlsx(*sheets):
    """Small OOXML fixtures permit exercising exact formula-cache/error states."""
    workbook = ET.Element(f"{{{MAIN}}}workbook", nsmap={None: MAIN, "r": REL})
    definitions = ET.SubElement(workbook, f"{{{MAIN}}}sheets")
    relationships = ET.Element(f"{{{PKG}}}Relationships", nsmap={None: PKG})
    parts = {"[Content_Types].xml": b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'}
    for index, rows in enumerate(sheets, 1):
        ET.SubElement(definitions, f"{{{MAIN}}}sheet", name=f"属性{index}", sheetId=str(index), attrib={f"{{{REL}}}id": f"rId{index}"})
        ET.SubElement(relationships, f"{{{PKG}}}Relationship", Id=f"rId{index}", Target=f"worksheets/sheet{index}.xml", Type=f"{REL}/worksheet")
        sheet = ET.Element(f"{{{MAIN}}}worksheet", nsmap={None: MAIN})
        sheet_data = ET.SubElement(sheet, f"{{{MAIN}}}sheetData")
        for row_number, values in enumerate(rows, 1):
            row = ET.SubElement(sheet_data, f"{{{MAIN}}}row", r=str(row_number))
            for column, value in enumerate(values):
                if value is None:
                    continue
                assert column < 26
                cell = ET.SubElement(row, f"{{{MAIN}}}c", r=f"{chr(65 + column)}{row_number}")
                if isinstance(value, dict):
                    if "error" in value:
                        cell.set("t", "e")
                        ET.SubElement(cell, f"{{{MAIN}}}v").text = value["error"]
                    else:
                        ET.SubElement(cell, f"{{{MAIN}}}f").text = value["formula"]
                        if "cache" in value:
                            ET.SubElement(cell, f"{{{MAIN}}}v").text = str(value["cache"])
                elif isinstance(value, str):
                    cell.set("t", "inlineStr")
                    ET.SubElement(ET.SubElement(cell, f"{{{MAIN}}}is"), f"{{{MAIN}}}t").text = value
                elif isinstance(value, bool):
                    cell.set("t", "b")
                    ET.SubElement(cell, f"{{{MAIN}}}v").text = "1" if value else "0"
                else:
                    ET.SubElement(cell, f"{{{MAIN}}}v").text = str(value)
        parts[f"xl/worksheets/sheet{index}.xml"] = ET.tostring(sheet)
    parts["xl/workbook.xml"] = ET.tostring(workbook)
    parts["xl/_rels/workbook.xml.rels"] = ET.tostring(relationships)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output.getvalue()


@pytest.fixture(scope="module")
def client():
    with TestClient(app) as value:
        yield value


def upload(client, data, name="attributes.xlsx"):
    return client.post("/api/import/attributes", files={"workbook": (name, data, "application/octet-stream")})


def test_horizontal_xlsx_only_returns_attributes_and_keeps_source(client):
    data = simple_xlsx([["简化属性表"], ["仅导入属性"], [], HEADERS + ["姓名", "年龄", "HP"], list(EXPECTED.values()) + ["不应导入的姓名", 45, 13]])
    before = hashlib.sha256(data).digest()
    response = upload(client, data)
    assert response.status_code == 200, response.text
    result = response.json()
    assert result["attributes"] == EXPECTED
    assert set(result) == {"attributes", "warnings", "summary"}
    assert result["summary"]["attribute_count"] == 9
    assert result["summary"]["layout"] == "horizontal"
    assert hashlib.sha256(data).digest() == before


def test_vertical_partial_blanks_and_zero(client):
    data = simple_xlsx([["属性", "数值"], ["力量（str）", 0], ["体质", None], ["DEX", 300], ["luck", 75], ["姓名", "保留网页姓名"]])
    response = upload(client, data)
    assert response.status_code == 200, response.text
    assert response.json()["attributes"] == {"STR": 0, "DEX": 300, "Luck": 75}
    assert response.json()["summary"]["layout"] == "vertical"


@pytest.mark.parametrize("rows", [
    [["STR", "姓名"], [60, "某人"]],
    [["属性", "数值"], ["STR", 60], ["姓名", "某人"]],
])
def test_single_attribute_with_ignored_identity_fields(client, rows):
    response = upload(client, simple_xlsx(rows))
    assert response.status_code == 200, response.text
    assert response.json()["attributes"] == {"STR": 60}


@pytest.mark.parametrize("name,body,encoding", [
    ("attributes.CSV", "STR,敏捷,幸运,姓名\n55,70,0,原名字不变\n", "utf-8-sig"),
    ("attributes.csv", "属性,数值\n力量,55\n敏捷,70\n幸运,0\n", "gb18030"),
    ("attributes.tsv", "力量\t敏捷\t幸运\n55\t70\t0\n", "utf-8"),
])
def test_delimited_formats(client, name, body, encoding):
    response = upload(client, body.encode(encoding), name)
    assert response.status_code == 200, response.text
    assert response.json()["attributes"] == {"STR": 55, "DEX": 70, "Luck": 0}


@pytest.mark.parametrize("value", [-1, 301, 55.5, True, "NaN", "Infinity", "1e999999999", "五十", "=50+10", "60/30/12"])
def test_invalid_values_fail_atomically(client, value):
    response = upload(client, simple_xlsx([["STR", "DEX"], [value, 70]]))
    assert response.status_code == 422, response.text
    assert "attributes" not in response.json()


@pytest.mark.parametrize("rows", [
    [["STR", "力量"], [50, 60]],
    [["属性", "数值"], ["STR", 50], ["力量", 50]],
    [["STR", "DEX"], [50, 60], [70, 80]],
    [["属性", "甲", "乙"], ["STR", 50, 60], ["DEX", 70, 80]],
    [["姓名", "年龄"], ["某人", 40]],
    [HEADERS, [None] * 9],
])
def test_ambiguous_or_empty_tables_rejected(client, rows):
    response = upload(client, simple_xlsx(rows))
    assert response.status_code == 422, response.text


def test_multiple_attribute_sheets_rejected(client):
    response = upload(client, simple_xlsx([["STR", "DEX"], [50, 60]], [["STR", "DEX"], [70, 80]]))
    assert response.status_code == 422, response.text


def test_non_attribute_sheet_does_not_block_single_attribute_sheet(client):
    response = upload(client, simple_xlsx([["使用说明"], ["填写属性后导入"]], [["STR", "DEX"], [50, 60]]))
    assert response.status_code == 200, response.text
    assert response.json()["attributes"] == {"STR": 50, "DEX": 60}


def test_cached_formula_read_without_evaluation(client):
    response = upload(client, simple_xlsx([["STR", "DEX"], [{"formula": 'WEBSERVICE("https://example.invalid")', "cache": 65}, 70]]))
    assert response.status_code == 200, response.text
    assert response.json()["attributes"] == {"STR": 65, "DEX": 70}
    assert response.json()["warnings"]


@pytest.mark.parametrize("value", [{"formula": "50+10"}, {"error": "#VALUE!"}])
def test_missing_cache_and_formula_error_rejected(client, value):
    response = upload(client, simple_xlsx([["STR", "DEX"], [value, 70]]))
    assert response.status_code == 422, response.text


@pytest.mark.parametrize("name,body", [("old.xls", b"not-xlsx"), ("file.xlsx", b"not-zip"), ("empty.csv", b""), ("macro.xlsm", b"no"), ("image.png", b"no")])
def test_invalid_uploads_rejected(client, name, body):
    response = upload(client, body, name)
    assert response.status_code == 422, response.text


def test_upload_size_limit(client):
    response = upload(client, b"x" * (12 * 1024 * 1024 + 1), "large.csv")
    assert response.status_code == 413, response.text


def test_deliverable_template_has_no_example_values(client):
    path = Path(__file__).resolve().parents[1] / "assets/templates/COC7属性简化表.xlsx"
    response = upload(client, path.read_bytes(), path.name)
    assert response.status_code == 422, response.text


def test_filled_deliverable_template_imports_all_nine_attributes(client):
    path = Path(__file__).resolve().parents[1] / "assets/templates/COC7属性简化表.xlsx"
    source_hash = hashlib.sha256(path.read_bytes()).digest()
    with zipfile.ZipFile(path) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    sheet_path = next(name for name in parts if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"))
    root = ET.fromstring(parts[sheet_path])
    for column, value in enumerate(EXPECTED.values()):
        cell = root.find(f".//{{{MAIN}}}c[@r='{chr(65 + column)}5']")
        assert cell is not None
        cell.attrib.pop("t", None)
        for child in list(cell):
            cell.remove(child)
        ET.SubElement(cell, f"{{{MAIN}}}v").text = str(value)
    parts[sheet_path] = ET.tostring(root)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in parts.items():
            target.writestr(name, data)
    response = upload(client, output.getvalue())
    assert response.status_code == 200, response.text
    assert response.json()["attributes"] == EXPECTED
    assert hashlib.sha256(path.read_bytes()).digest() == source_hash
