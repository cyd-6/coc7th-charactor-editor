"""Read the editable cells of a CY26.3 or compatible CY26.2 investigator card.

This deliberately does not open Office or recalculate an uploaded workbook.
XML values and saved formula caches are read from a bounded in-memory package;
relationships may only address other members of that package.
"""

from __future__ import annotations

import base64
import io
import posixpath
import re
import zipfile
import zlib
from datetime import date
from decimal import Decimal, InvalidOperation, localcontext

from lxml import etree as ET

from ..branch_workbook import BRANCH_FIRST_ROW, BRANCH_SCHEMA, BRANCH_SCHEMA_ROW, BRANCH_TITLE, branch_row, selection_title_row
from ..catalog import TemplateCatalog, normalize_skill_name, split_occupation_skill_token
from ..models import Attributes
from ..portraits import prepare_portrait
from ..rules import RuleEngine
from ..skill_specializations import (GROUP_BASE_VALUES, MAX_SKILL_BRANCHES, canonical_specialization, is_branch_slot, matching_branch_slot, skill_group,
                                      specialization_base_value, validate_branch_identity)


MAX_UPLOAD_BYTES = 12 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 80 * 1024 * 1024
MAX_MEMBERS = 2000
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
ART = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"s": MAIN, "r": REL, "d": DRAW, "a": ART}
ATTRIBUTE_CELLS = {"STR": "U3", "CON": "U5", "SIZ": "U7", "DEX": "AA3", "APP": "AA5", "INT": "AA7", "POW": "AG3", "EDU": "AG5", "Luck": "AG7"}
BACKGROUND_CELLS = {"appearance": "AA61", "beliefs": "AA63", "significant_people": "AA65", "meaningful_places": "AA67", "treasured_possessions": "AA69", "traits": "AA71", "scars": "AA73", "phobias_manias": "AA75", "personal_story": "W77"}
ASSET_CELLS = {"living_standard": "F62", "asset_description": "L63", "vehicles": "B70", "residence": "F70", "luxuries": "J70", "securities": "N70", "other": "R70"}


class ExcelImportError(ValueError):
    """An uploaded file cannot be safely interpreted as a supported card."""


def _tag(name: str) -> str:
    return f"{{{MAIN}}}{name}"


def _text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return format(value, "f")
    return str(value).strip()


def _number(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool) or not _text(value):
        return None
    try:
        number = Decimal(_text(value).replace(",", ""))
    except InvalidOperation:
        return None
    return number if number.is_finite() and number.copy_abs() <= Decimal("1e15") else None


class _Workbook:
    def __init__(self, data: bytes, *, require_card: bool = True):
        if not isinstance(data, bytes) or not data:
            raise ExcelImportError("请选择非空的 XLSX 调查员表格。")
        if len(data) > MAX_UPLOAD_BYTES:
            raise ExcelImportError("表格文件不能超过 12 MB。")
        self.archive = zipfile.ZipFile(io.BytesIO(data))
        members = self.archive.infolist()
        names = [member.filename for member in members]
        if len(members) > MAX_MEMBERS or sum(member.file_size for member in members) > MAX_UNCOMPRESSED_BYTES:
            raise ExcelImportError("表格解压后的内容过大，请使用原始调查员卡或精简后的 XLSX 文件。")
        if len(names) != len(set(names)) or any(member.flag_bits & 1 for member in members):
            raise ExcelImportError("不支持加密或包含重复内部文件的表格，请另存为普通 XLSX 后重试。")
        if any(name.startswith("/") or "\\" in name or ".." in name.split("/") for name in names):
            raise ExcelImportError("表格内部文件路径无效。")
        self.names = set(names)
        if not {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"} <= self.names:
            raise ExcelImportError("文件不是有效的 XLSX 工作簿；旧版 XLS 请先另存为 XLSX。")
        if any(name.lower().endswith("vbaproject.bin") for name in names):
            raise ExcelImportError("请将含宏的工作簿另存为不含宏的 XLSX 文件后导入。")
        self.warnings: list[str] = []
        self.missing_formulas: set[str] = set()
        self.failed_values: set[str] = set()
        self.roots: dict[str, ET._Element] = {}
        self.sheets: dict[str, dict[str, ET._Element]] = {}
        self.strings: list[str] = []
        if "xl/sharedStrings.xml" in self.names:
            for item in self.xml("xl/sharedStrings.xml"):
                self.strings.append("".join(node.text or "" for node in item.iter(_tag("t"))))
        workbook = self.xml("xl/workbook.xml")
        relations = self.relations("xl/workbook.xml")
        self.paths = {}
        for sheet in workbook.findall("s:sheets/s:sheet", NS):
            relation = relations.get(sheet.get(f"{{{REL}}}id"))
            if relation:
                self.paths[sheet.get("name")] = relation
        if require_card and "人物卡" not in self.paths:
            raise ExcelImportError("未找到受支持的“人物卡”工作表。请导入 CY26.3／CY26.2 模板或本工具导出的 XLSX。")

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)

    def xml(self, path: str) -> ET._Element:
        if path not in self.roots:
            raw = self.archive.read(path)
            if b"<!DOCTYPE" in raw.upper() or b"<!ENTITY" in raw.upper():
                raise ExcelImportError("表格包含不支持的 XML 声明，请重新另存为 XLSX。")
            parser = ET.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)
            self.roots[path] = ET.fromstring(raw, parser)
        return self.roots[path]

    def relations(self, path: str) -> dict[str, str]:
        directory, filename = posixpath.split(path)
        relation_path = f"{directory}/_rels/{filename}.rels"
        if relation_path not in self.names:
            return {}
        result = {}
        for node in self.xml(relation_path):
            if node.get("TargetMode", "").lower() == "external":
                continue
            target = node.get("Target", "")
            if not target or ":" in target or "\\" in target:
                continue
            resolved = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(directory, target))
            if resolved in self.names and not resolved.startswith("../"):
                result[node.get("Id")] = resolved
        return result

    def cells(self, sheet: str) -> dict[str, ET._Element]:
        if sheet not in self.paths:
            return {}
        if sheet not in self.sheets:
            self.sheets[sheet] = {cell.get("r"): cell for cell in self.xml(self.paths[sheet]).iter(_tag("c"))}
        return self.sheets[sheet]

    def formula(self, address: str, sheet: str = "人物卡") -> str | None:
        cell = self.cells(sheet).get(address)
        node = cell.find(_tag("f")) if cell is not None else None
        return None if node is None else node.text or ""

    def value(self, address: str, sheet: str = "人物卡", *, report: bool = True):
        cell = self.cells(sheet).get(address)
        if cell is None:
            return None
        return self.cell_value(cell, address, sheet, report=report)

    def cell_value(self, cell: ET._Element, address: str, sheet: str, *, report: bool = True):
        kind = cell.get("t", "n")
        value = cell.find(_tag("v"))
        if kind == "inlineStr":
            return "".join(node.text or "" for node in cell.findall("s:is//s:t", NS))
        if kind == "e":
            if report:
                self.failed_values.add(f"{sheet}!{address}")
            return None
        if value is None:
            if report and cell.find(_tag("f")) is not None:
                self.missing_formulas.add(f"{sheet}!{address}")
            return None
        raw = value.text or ""
        if kind == "s":
            index = int(raw)
            if index < 0 or index >= len(self.strings):
                raise ExcelImportError("表格的文本索引无效，文件可能已损坏。")
            return self.strings[index]
        if kind in {"str", "d"}:
            return raw
        if kind == "b":
            return raw == "1"
        if not raw:
            return None
        number = _number(raw)
        return number if number is not None else raw

    def text(self, address: str, sheet: str = "人物卡", *, report: bool = True) -> str:
        return _text(self.value(address, sheet, report=report))

    def integer(self, address: str, default: int = 0, sheet: str = "人物卡", *, report: bool = True) -> int:
        raw = self.value(address, sheet, report=report)
        number = _number(raw)
        if number is None:
            if raw not in (None, ""):
                self.warn(f"{sheet}!{address} 不是有效整数，已使用 {default}，请核对。")
            return default
        if number != int(number):
            self.warn(f"{sheet}!{address} 含小数，点数已取整数部分，请核对。")
        return int(number)

    def finish_warnings(self) -> None:
        for cells, message in ((self.missing_formulas, "部分公式没有已保存的计算结果，未执行这些公式"), (self.failed_values, "部分单元格保存了公式错误，未导入错误值")):
            if cells:
                sample = "、".join(sorted(cells)[:6])
                self.warn(f"{message}（{sample}{'等' if len(cells) > 6 else ''}）。请在 Excel 或 LibreOffice 重算并保存后重试，或手动补全。")


def _recognize(book: _Workbook, catalog: TemplateCatalog) -> None:
    anchors = {"B3": "姓名", "B4": "玩家", "B5": "职业", "B6": "年龄", "F15": "技能名称", "AB15": "技能名称", "B52": "武器名称"}
    score = sum(book.text(cell, report=False).replace("\n", "") == expected for cell, expected in anchors.items())
    skills = sum(normalize_skill_name(book.text(item.template_slot, report=False)) == normalize_skill_name(item.name) for item in catalog.skills)
    if score < 5 or skills < 12 or "STR" not in book.text("S3", report=False):
        raise ExcelImportError("表格布局与 CY26.3／CY26.2 调查员卡不一致。请使用原始模板或本工具导出的 XLSX。")


def _identity(book: _Workbook) -> tuple[dict, dict]:
    identity = {key: book.text(cell) for key, cell in {"name": "E3", "player": "E4", "gender": "M6", "residence": "E7", "birthplace": "M7", "occupation_name": "E5"}.items()}
    identity["age"] = book.integer("E6", 30)
    era = book.text("M4")
    identity["era"] = era or "1920s"
    attributes = {key: book.integer(cell, 50) for key, cell in ATTRIBUTE_CELLS.items()}
    missing = [key for key, cell in ATTRIBUTE_CELLS.items() if book.value(cell, report=False) is None]
    if missing:
        book.warn(f"属性 {', '.join(missing)} 没有可读取的值，暂填 50，请补全。")
    identity["current_date"] = ""
    year = book.integer("G8", 0)
    month = re.fullmatch(r"(\d{1,2})月?", book.text("J8"))
    day = re.fullmatch(r"(\d{1,2})日?", book.text("L8"))
    if year and month and day:
        try:
            identity["current_date"] = date(year, int(month[1]), int(day[1])).isoformat()
        except (ValueError, OverflowError):
            book.warn("表格中的当前日期无效，已留空，请重新填写。")
    elif any(book.value(cell, report=False) not in (None, "") for cell in ("G8", "J8", "L8")):
        book.warn("表格中的当前日期不完整，已留空，请重新填写。")
    return identity, attributes


def _skills(book: _Workbook, catalog: TemplateCatalog, attributes: dict) -> list[dict]:
    separate = all("经历" in book.text(cell, report=False).replace("\n", "") for cell in ("M15", "AI15"))
    if not separate:
        book.warn("旧版表格未区分经历包点与成长点数：合并的增长值已保留在“成长点数”，经历包点暂填 0，请按原记录拆分。")
    result = []
    overridden = []
    attribute_values = Attributes.from_mapping(attributes)
    for definition in catalog.skills:
        mapping = catalog.skill_cell(definition.template_slot)
        name = normalize_skill_name(book.text(mapping.name_cell)) or definition.name
        specialization = book.text(mapping.specialization_cell)
        expected = RuleEngine.skill_base_value(definition.base_formula, attribute_values, definition.base_value)
        expected = specialization_base_value(name, specialization or definition.default_specialization, expected)
        base = book.integer(mapping.base_cell, expected)
        formula = re.sub(r"[\s$]", "", book.formula(mapping.base_cell) or "").upper()
        # Known standard attribute references are recalculated by the existing
        # rules, without evaluating spreadsheet expressions or trusting a cache
        # left behind after the attributes were edited outside Excel.
        standard_formula = (
            definition.base_formula == "DEX/2" and formula in {"INT(DEX/2)", "INT(AA3/2)", "DEX/2", "AA3/2"}
            or definition.base_formula == "EDU" and formula in {"EDU", "AG5", "INT(EDU)", "INT(AG5)"}
        )
        if standard_formula:
            base = expected
        row = {
            "key": definition.template_slot, "template_slot": definition.template_slot, "name": name,
            "specialization": specialization, "base_value": base,
            "occupation_points": 0 if "克苏鲁神话" in name else book.integer(mapping.occupation_cell),
            "interest_points": 0 if "克苏鲁神话" in name else book.integer(mapping.interest_cell),
            "extra_final": book.integer(mapping.extra_cell),
            "experience_points": book.integer(mapping.experience_cell) if separate else 0,
            "selected_occupation": False,
        }
        if base != expected:
            if 0 <= base <= 999:
                row["base_override"] = base
                overridden.append(name)
            else:
                row["base_value"] = expected
                book.warn(f"“{name}”初始值 {base} 超出可导入范围，暂用规则值 {expected}，请核对。")
        if any(row[key] < 0 for key in ("occupation_points", "interest_points", "extra_final", "experience_points")):
            book.warn(f"“{name}”包含负数点数，已保留供核对，请修正后再导出。")
        if name != definition.name:
            book.warn(f"技能位置 {definition.template_slot} 已改名为“{name}”，已保留，请核对职业技能选择。")
        result.append(row)
    if overridden:
        book.warn(f"已保留表格的自定义技能初始值（{'、'.join(overridden)}），请确认。")
    return result



def _branch_integer(book: _Workbook, address: str, *, maximum: int | None = None) -> int:
    value = _number(book.value(address))
    if value is None or value != value.to_integral_value() or value < 0 or maximum is not None and value > maximum:
        raise ExcelImportError(f"新增技能补充区“人物卡!{address}”必须有已保存的非负整数，请核对并重算保存后导入。")
    return int(value)


def _branch_section(book: _Workbook) -> tuple[list[dict], list[tuple[str, str, str]] | None]:
    schema = book.text(f"B{BRANCH_SCHEMA_ROW}", report=False)
    if not schema and book.text("B151", report=False) != BRANCH_TITLE:
        return [], None
    if schema != BRANCH_SCHEMA:
        raise ExcelImportError("新增技能补充区的版本标记缺失或不受支持，请保留补充区及导入恢复信息后重试。")
    count = _branch_integer(book, f"AP{BRANCH_SCHEMA_ROW}", maximum=len(book.cells("人物卡")))
    if count > MAX_SKILL_BRANCHES:
        raise ExcelImportError(f"新增技能分支不能超过 {MAX_SKILL_BRANCHES} 个，请在原文件中调整后重新导入；当前人物未修改。")
    if count < 1 or selection_title_row(count) > 1048576:
        raise ExcelImportError("新增技能补充区的分支数量无效，请使用本工具导出的原始 XLSX。")
    expected_identifiers = {f"F{branch_row(index) + 1}" for index in range(count)}
    for address in book.cells("人物卡"):
        match = re.fullmatch(r"F(\d+)", address)
        if match and int(match[1]) >= BRANCH_FIRST_ROW and address not in expected_identifiers:
            if book.text(address, report=False).startswith("branch-"):
                raise ExcelImportError("新增技能补充区数量与分支标识记录不一致，请保留全部分支后重新导入。")
    branches = []
    identifiers = set()
    for index in range(count):
        row = branch_row(index)
        try:
            slot, name, specialization = validate_branch_identity(
                book.text(f"F{row + 1}"), book.text(f"B{row}"), book.text(f"F{row}"),
            )
        except ValueError as exc:
            raise ExcelImportError(f"新增技能补充区第 {index + 1} 项无法导入：{exc}") from exc
        if slot in identifiers:
            raise ExcelImportError("新增技能补充区包含重复分支标识，请恢复每个分支的独立标识后导入。")
        identifiers.add(slot)
        base = _branch_integer(book, f"R{row}", maximum=999)
        flag = book.text(f"AP{row}")
        if flag not in {"", "★"}:
            raise ExcelImportError(f"新增技能补充区“人物卡!AP{row}”本职标志应为空白或 ★。")
        skill = {
            "key": slot, "template_slot": slot, "name": name, "specialization": specialization,
            "base_value": base,
            "occupation_points": _branch_integer(book, f"U{row}"),
            "interest_points": _branch_integer(book, f"X{row}"),
            "experience_points": _branch_integer(book, f"AA{row}"),
            "extra_final": _branch_integer(book, f"AD{row}"),
            "selected_occupation": flag == "★",
        }
        expected = specialization_base_value(name, specialization, GROUP_BASE_VALUES[name])
        if base != expected:
            skill["base_override"] = base
        branches.append(skill)
    metadata_row = selection_title_row(count)
    if book.text(f"B{metadata_row}") != "导入恢复信息（请保留）":
        raise ExcelImportError("新增技能的职业选择恢复信息缺失，请保留导出文件中的补充区域后重试。")
    record_count = _branch_integer(book, f"AP{metadata_row}", maximum=len(book.cells("人物卡")))
    if metadata_row + record_count > 1048576:
        raise ExcelImportError("新增技能的职业选择恢复记录数量无效。")
    for address in book.cells("人物卡"):
        match = re.fullmatch(r"B(\d+)", address)
        if match and int(match[1]) > metadata_row + record_count:
            if book.text(address, report=False) in {"custom", "free", "group"}:
                raise ExcelImportError("新增技能的职业选择记录数量不完整，请保留全部恢复信息后重新导入。")
    records = [(book.text(f"B{row}"), book.text(f"F{row}"), book.text(f"J{row}"))
               for row in range(metadata_row + 1, metadata_row + record_count + 1)]
    if len(records) != len(set(records)):
        raise ExcelImportError("新增技能的职业选择恢复信息包含重复记录，请核对后重新导入。")
    return branches, records



def _validate_imported_branch_names(skills: list[dict]) -> None:
    seen = {}
    for skill in skills:
        group = skill_group(skill["name"])
        specialization = skill["specialization"].strip()
        if not group or not specialization:
            continue
        identity = group, canonical_specialization(group, specialization)
        branch = is_branch_slot(skill["template_slot"])
        if identity in seen and (branch or seen[identity]):
            raise ExcelImportError(f"新增技能“{group}（{specialization}）”与已有分项重名，请使用独立名称后重新导入。")
        seen[identity] = branch


def _restore_branch_selections(draft: dict, catalog: TemplateCatalog,
                               records: list[tuple[str, str, str]], branches: list[dict]) -> None:
    skills = draft["skills"]
    known = {skill["template_slot"] for skill in skills}
    custom, free, groups = [], [], {}
    for kind, marker, value in records:
        if kind in {"custom", "free"} and not marker and value in known and (kind == "free" or value != "F26"):
            (custom if kind == "custom" else free).append(value)
        elif kind == "group" and marker and value:
            groups.setdefault(marker, []).append(value)
        else:
            raise ExcelImportError("新增技能的职业选择恢复记录无效，技能标识必须对应原技能或补充分支。")
    if draft["occupation_mode"] == "custom":
        if free or groups or len(custom) > 8:
            raise ExcelImportError("自定义职业的技能恢复记录无效：最多八项本职技能，不能同时包含目录职业选择。")
        draft["custom_skill_slots"] = custom
        selected = set(custom) | {"F26"}
    else:
        occupation = catalog.occupation_by_id(draft["occupation_id"])
        definitions = {group.marker: group for group in occupation.choice_groups}
        if custom or len(free) > occupation.free_choices or any(
            marker not in definitions or len(choices) > definitions[marker].required_count
            or any(token not in definitions[marker].candidates for token in choices)
            for marker, choices in groups.items()
        ):
            raise ExcelImportError("目录职业的技能恢复记录与所选职业不一致，请核对职业与补充区后重新导入。")
        draft["free_skill_choices"] = free
        draft["group_choices"] = groups
        selected = {"F26", *free}
        for token in [*occupation.fixed_skills, *(token for choices in groups.values() for token in choices)]:
            selected.update(_matching_slots(token, skills)[:1])
    if any(skill["selected_occupation"] != (skill["template_slot"] in selected) for skill in branches):
        raise ExcelImportError("新增技能的本职标志与职业选择恢复信息不一致，请同步核对后重新导入。")
    for skill in skills:
        skill["selected_occupation"] = skill["template_slot"] in selected


def _matching_slots(token: str, skills: list[dict]) -> list[str]:
    name, specialization = split_occupation_skill_token(token)
    branch = matching_branch_slot(name, specialization, skills)
    if branch is not None:
        return [branch]
    return [skill["template_slot"] for skill in skills if normalize_skill_name(skill["name"]) == name and (not specialization or not skill["specialization"] or skill["specialization"] == specialization)]


def _occupation(book: _Workbook, catalog: TemplateCatalog, draft: dict) -> None:
    skills = draft["skills"]
    name = draft["identity"]["occupation_name"]
    identifier = book.integer("M5")
    by_name = catalog.occupation_by_name(name) if name else None
    occupation = by_name if name else catalog.occupation_by_id(identifier)
    draft.update(occupation_mode="catalog", occupation_id=occupation.occupation_id if occupation else catalog.occupations[0].occupation_id,
                 custom_occupation={"name": name or book.text("C3", "职业列表") or "自定义职业", "credit_min": 0, "credit_max": 99, "point_formula_kind": "EDU*4", "secondary": "DEX"},
                 custom_skill_slots=[], group_choices={}, free_skill_choices=[])
    # Explicit flags in the card take precedence over inference from spending.
    marked = set()
    for skill in skills:
        slot = skill["template_slot"]
        if is_branch_slot(slot):
            if skill["selected_occupation"] or skill["occupation_points"] > 0:
                marked.add(slot)
            continue
        flag = book.text(("D" if slot.startswith("F") else "Z") + re.search(r"\d+", slot)[0], report=False)
        if flag in {"★", "√", "☒", "1"} or skill["occupation_points"] > 0:
            marked.add(slot)
    if occupation is None or identifier == 1:
        draft["occupation_mode"] = "custom"
        if identifier != 1:
            book.warn("职业名称未能对应网页职业库，已保留为自定义职业，请核对职业点公式和本职技能。")
        custom = draft["custom_occupation"]
        credit = re.fullmatch(r"\s*(\d+)\s*[-—~～至]\s*(\d+)\s*", book.text("D3", "职业列表"))
        if credit:
            custom["credit_min"], custom["credit_max"] = sorted((int(credit[1]), int(credit[2])))
        formula = re.sub(r"\s+", "", book.formula("F3", "职业列表") or "").upper()
        secondary = re.fullmatch(r"EDU\*2\+(STR|CON|SIZ|DEX|APP|INT|POW|EDU|LUCK)\*2", formula)
        if secondary:
            custom["point_formula_kind"] = "secondary"
            custom["secondary"] = "Luck" if secondary[1] == "LUCK" else secondary[1]
        elif formula == "SUM(附表!AF10:AF12)":
            selected = [key for row, key in ((15,"STR"),(16,"DEX"),(17,"CON"),(18,"SIZ"),(19,"APP"),(20,"INT"),(21,"EDU"),(22,"POW"),(23,"Luck")) if book.text(f"I{row}", "职业列表") == "√"]
            if "EDU" in selected and len(selected) == 2:
                custom["point_formula_kind"] = "secondary"
                custom["secondary"] = next(key for key in selected if key != "EDU")
            elif selected != ["EDU"]:
                book.warn("自定义职业点公式无法完整对应网页选项，暂用 EDU×4，请重新确认。")
        elif formula != "EDU*4":
            book.warn("自定义职业点公式无法完整对应网页选项，暂用 EDU×4，请重新确认。")
        if custom["secondary"] in {"EDU", "Luck"}:
            custom["point_formula_kind"] = "EDU*4"
            custom["secondary"] = "DEX"
            book.warn("自定义职业使用的第二属性不在网页选项中，暂用 EDU×4，请重新确认。")
        named_slots = []
        for row in range(3, 11):
            token = book.text(f"I{row}", "职业列表")
            candidates = _matching_slots(token, skills) if token else []
            if len(candidates) == 1:
                named_slots.extend(candidates)
        if "F26" in named_slots:
            book.warn("原表本职技能列表包含自动计入的信用评级，请核对是否遗漏了其他本职技能。")
        chosen = list(dict.fromkeys(slot for slot in named_slots + [s["template_slot"] for s in skills if s["template_slot"] in marked] if slot != "F26"))
        draft["custom_skill_slots"] = chosen[:8]
        if len(chosen) > 8:
            book.warn("表格标记的自定义职业技能超过 8 项，仅选中前 8 项；各技能点数均已保留，请核对。")
        selected_slots = set(draft["custom_skill_slots"]) | {"F26"}
    else:
        if by_name and identifier not in (0, by_name.occupation_id):
            book.warn("表格职业名称与序号不一致，已按可识别的职业名称导入。")
        draft["identity"]["occupation_name"] = occupation.name
        selected_slots = {"F26"}
        for token in occupation.fixed_skills:
            selected_slots.update(_matching_slots(token, skills)[:1])
        for group in occupation.choice_groups:
            choices = [token for token in group.candidates if any(slot in marked for slot in _matching_slots(token, skills))]
            if len(choices) > group.required_count:
                choices = []
                book.warn(f"职业的“{group.label}”无法唯一恢复，请重新选择；技能点数已保留。")
            elif len(choices) < group.required_count:
                book.warn(f"职业的“{group.label}”未完整记录，请补选；技能点数已保留。")
            draft["group_choices"][group.marker] = choices
            for token in choices:
                selected_slots.update(_matching_slots(token, skills)[:1])
        other = [skill["template_slot"] for skill in skills if skill["template_slot"] in marked - selected_slots]
        if occupation.free_choices and len(other) <= occupation.free_choices:
            draft["free_skill_choices"] = other
            selected_slots.update(other)
        if occupation.free_choices and len(draft["free_skill_choices"]) != occupation.free_choices:
            book.warn("任意职业特长未能完整恢复，请重新确认选择；技能点数已保留。")
        if marked - selected_slots:
            book.warn("部分已填职业点的技能未能对应本职选择，请在职业配置中核对。")
    for skill in skills:
        skill["selected_occupation"] = skill["template_slot"] in selected_slots


def _experience(book: _Workbook, catalog: TemplateCatalog) -> dict:
    name = book.text("F113")
    result = {"selection": "", "name": "", "san_loss": 0, "skill_points": 0, "notes": ""}
    if not name or name == "无":
        return result
    definition = next((item for item in catalog.experience_packages if item["name"] == name), None)
    result.update(selection=name if definition else "custom", name=name,
                  skill_points=definition["skill_points"] if definition else book.integer("BG28"),
                  notes=definition["notes"] if definition else book.text("BC30"))
    if name == "自定义经历包":
        result["name"] = book.text("BC26") or name
    # BC28 is explicitly written by both exporters and is not inferred from a
    # played character's current SAN, which may have changed for other reasons.
    if book.value("BC28") is not None:
        result["san_loss"] = book.integer("BC28")
    else:
        book.warn("经历包的实际 SAN 减少值未记录，暂填 0，请按原记录补全；当前 SAN 不用于反推。")
    if not definition and book.value("BG28") is None:
        book.warn("自定义经历包额度未记录，暂填 0，请补全。")
    return result


def _assets(book: _Workbook, catalog: TemplateCatalog, identity: dict) -> dict:
    result = {key: book.text(cell) for key, cell in ASSET_CELLS.items()}
    if result["asset_description"] == "请在这里详述你的资产":
        result["asset_description"] = ""
    annual = "汇率年份" in book.text("J1", "货币汇率", report=False) and "目标币种" in book.text("J2", "货币汇率", report=False)
    if annual:
        year = book.integer("K1", 0, "货币汇率")
        currency_link = re.sub(r"[\s$'=]", "", book.formula("K2", "货币汇率") or "")
        # Read the authoritative editable cell, not an old saved formula cache.
        currency = book.text("S62") if currency_link == "人物卡!S62" else book.text("K2", "货币汇率")
    else:
        year = int(identity["current_date"][:4]) if identity["current_date"] else (2026 if identity["era"] == "现代" else 1920)
        currency = book.text("S62")
        book.warn("旧表未记录独立汇率年份，已按当前日期／时代选择年份，请核对。")
    result.update(exchange_year=year, currency=currency, currency_code="", usd_spending="", usd_cash="", usd_assets="")
    identity["story_year"] = year
    quotes = [quote for quote in catalog.currency_quotes if quote["year"] == year and quote["available"]]
    # Only match the actual contemporary currency, never a legacy proxy code.
    normalize_currency = lambda value: re.sub(r"[\s（）()]", "", value)
    candidates = [quote for quote in quotes if normalize_currency(currency) in {normalize_currency(quote["name"]), quote["code"], quote["name"].split("（")[0]}]
    quote = candidates[0] if len(candidates) == 1 else None
    if quote:
        result["currency_code"] = quote["code"]
        result["currency"] = quote["name"]
    else:
        book.warn(f"汇率年份／币种“{year}／{currency or '空白'}”无法对应可用年度报价，已保留美元输入，请重新选择币种。")
    for key, cell, card_cell in (("usd_spending", "K9", "I62"), ("usd_cash", "K10", "O62"), ("usd_assets", "K11", "L62")):
        if annual:
            formula = book.formula(cell, "货币汇率")
            # These exact input formulas mean “use credit reference”. Arbitrary
            # formulas retain their saved result instead of being executed.
            credit = "人物卡!R26"
            reference_formulas = {
                "K9": f"LOOKUP({credit},{{0,1,10,50,90,99}},{{0.5,2,10,50,250,5000}})",
                "K10": f"IF({credit}=0,0.5,IF({credit}=99,50000,{credit}*LOOKUP({credit},{{1,10,50,90}},{{1,2,5,20}})))",
                "K11": f"IF({credit}=0,0,IF({credit}=99,5000000,{credit}*LOOKUP({credit},{{1,10,50,90}},{{10,50,500,2000}})))",
            }
            normalized_formula = re.sub(r"[\s$']", "", formula or "").upper()
            if normalized_formula == reference_formulas[cell] + "*IF(K1=2026,20,1)":
                continue
            value = book.value(cell, "货币汇率")
            if value not in (None, ""):
                result[key] = book.text(cell, "货币汇率")
        elif quote:
            value = _number(book.value(card_cell))
            if value is not None:
                with localcontext() as context:
                    context.prec = 40
                    result[key] = format(value / Decimal(quote["rate"]), "f")
    if not annual:
        book.warn("旧版表格的资产金额按已识别的年度报价换回美元；原有报价可能不同，请核对美元输入。" if quote else "旧版表格只有本币资产金额，无法可靠恢复美元输入，金额未换算，请手动填写。")
    details = []
    for row, label, cell in ((13,"交通工具","B75"),(14,"住所","F75"),(15,"奢侈品","J75"),(16,"股票／证券","N75"),(17,"其他","R75")):
        amount = book.value(f"K{row}", "货币汇率") if annual else book.value(cell)
        number = _number(amount)
        if number is not None and number != 0:
            details.append(f"{label}：{_text(amount)}")
    if details:
        note = f"表格资产明细（{'美元' if annual else currency or '原表币种'}）：" + "；".join(details)
        result["asset_description"] = "\n".join(filter(None, (result["asset_description"], note)))
        book.warn("资产明细金额已放入“资产说明”，可继续查看；网页暂无对应的独立金额输入。")
    return result


def _equipment(book: _Workbook) -> tuple[list[dict], list[dict]]:
    weapons, inventory = [], []
    for row in range(53, 59):
        name, category = book.text(f"B{row}"), book.text(f"G{row}")
        if name == "无" and category == "肉搏":
            continue  # Blank template's example row.
        if not name and not category:
            continue
        weapon = {key: book.text(f"{col}{row}") for key, col in {"name": "B", "category": "G", "skill": "M", "damage": "W", "range": "AA", "attacks": "AE", "ammo": "AG", "malfunction": "AJ"}.items()}
        weapon["name"] = name or category
        weapon.update(notes="", catalog_id="custom")
        if weapon["skill"] == "←请选择类型":
            weapon["skill"] = ""
        weapons.append(weapon)
    for row in range(79, 94):
        if not book.text(f"F{row}"):
            continue
        item = {key: book.text(f"{col}{row}") for key, col in {"name": "F", "status": "B", "location": "D", "backpack_slot": "N"}.items()}
        item["catalog_id"] = "custom"
        inventory.append(item)
    return weapons, inventory


def _portrait(book: _Workbook) -> dict | None:
    path = book.paths["人物卡"]
    drawings = book.xml(path).findall("s:drawing", NS)
    candidates = []
    pictures = 0
    relations = book.relations(path)
    for ref in drawings:
        drawing_path = relations.get(ref.get(f"{{{REL}}}id"))
        if not drawing_path:
            continue
        image_relations = book.relations(drawing_path)
        for anchor in book.xml(drawing_path):
            image = anchor.find("d:pic/d:blipFill/a:blip", NS)
            if image is None:
                continue
            pictures += 1
            row, col = anchor.findtext("d:from/d:row", namespaces=NS), anchor.findtext("d:from/d:col", namespaces=NS)
            if row is None or col is None or not (2 <= int(row) <= 8 and 37 <= int(col) <= 44):
                continue
            target = image_relations.get(image.get(f"{{{REL}}}embed"))
            if target:
                candidates.append(target)
    if len(candidates) != 1:
        if pictures:
            book.warn("未能唯一识别肖像区域的内嵌图片，头像未导入，请手动选择。")
        return None
    raw = book.archive.read(candidates[0])
    if len(raw) > 8 * 1024 * 1024:
        book.warn("表格头像超过 8 MB，未导入，请手动上传较小图片。")
        return None
    try:
        with prepare_portrait(raw, allow_legacy=True) as source:
            output = io.BytesIO()
            source.save(output, format="PNG")
        image_data = output.getvalue()
        if len(image_data) > 8 * 1024 * 1024:
            raise ValueError("Portrait too large")
        return {"filename": "调查员头像.png", "mime_type": "image/png", "data_base64": base64.b64encode(image_data).decode("ascii")}
    except (OSError, ValueError):
        book.warn("表格头像格式或尺寸不受支持，头像未导入，请手动上传。")
        return None


def import_investigator(data: bytes, catalog: TemplateCatalog) -> dict:
    """Return an import preview; neither the draft nor any file is persisted."""
    book = None
    try:
        book = _Workbook(data)
        _recognize(book, catalog)
        identity, attributes = _identity(book)
        draft = {"version": 3, "current_step": 0, "identity": identity, "attributes": attributes,
                 "nonstandard_override": False, "skills": _skills(book, catalog, attributes)}
        branches, selections = _branch_section(book)
        draft["skills"].extend(dict(skill) for skill in branches)
        _validate_imported_branch_names(draft["skills"])
        _occupation(book, catalog, draft)
        if selections is not None:
            _restore_branch_selections(draft, catalog, selections, branches)
        draft["experience"] = _experience(book, catalog)
        derived = RuleEngine.calculate(Attributes.from_mapping(attributes), identity["age"], "EDU*4", draft["experience"]["san_loss"])
        changed_status = []
        for label, cell, expected in (("HP", "E10", derived.hp), ("MP", "W10", derived.mp), ("SAN", "N10", derived.san)):
            saved = _number(book.value(cell))
            if saved is not None and saved != expected:
                changed_status.append(f"{label} 原表 {_text(saved)}／建卡值 {expected}")
        if changed_status:
            book.warn("网页按属性与经历包重新计算派生值，未恢复原表当前游玩状态（" + "；".join(changed_status) + "）。请保留原表中的当前状态记录。")
        draft["background"] = {key: book.text(cell) for key, cell in BACKGROUND_CELLS.items()}
        if draft["background"]["personal_story"] == "请务必在此填写背景故事！\n使用Alt+Enter换行":
            draft["background"]["personal_story"] = ""
        labels = ["形象描述", "思想与信念", "重要之人", "意义非凡之地", "宝贵之物", "特质", "伤口和疤痕", "恐惧症和躁狂症"]
        draft["background"]["key_connection"] = "、".join(label for row, label in zip(range(61, 77, 2), labels) if book.text(f"AR{row}") in {"☒", "√", "✓", "1"})
        draft["background"]["contacts_notes"] = ""
        draft["assets"] = _assets(book, catalog, identity)
        draft["weapons"], draft["inventory"] = _equipment(book)
        portrait = _portrait(book)
        book.finish_warnings()
        return {"draft": draft, "warnings": book.warnings, "portrait": portrait,
                "summary": {"name": identity["name"], "occupation": draft["custom_occupation"]["name"] if draft["occupation_mode"] == "custom" else identity["occupation_name"],
                            "era": identity["era"], "skill_count": len(draft["skills"]),
                            "weapon_count": len(draft["weapons"]), "inventory_count": len(draft["inventory"])}}
    except ExcelImportError:
        raise
    except (zipfile.BadZipFile, zlib.error, KeyError, ET.XMLSyntaxError, ValueError, TypeError, OverflowError, OSError, RuntimeError) as exc:
        raise ExcelImportError("无法读取此表格。请确认文件未损坏、未加密，并使用 CY26.3／CY26.2 模板或本工具导出的 XLSX。") from exc
    finally:
        if book is not None:
            book.archive.close()
