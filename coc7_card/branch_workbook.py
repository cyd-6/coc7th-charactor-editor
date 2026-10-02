"""Visible XLSX branch rows and compact metadata for lossless imports."""

from __future__ import annotations

from collections.abc import Callable

from lxml import etree as ET
from openpyxl.utils import get_column_letter, range_boundaries

from .models import CharacterDraft
from .skill_specializations import is_branch_slot


BRANCH_TITLE = "新增技能分支"
BRANCH_SCHEMA = "COC7_BRANCHES_V1"
BRANCH_SCHEMA_ROW = 153
BRANCH_FIRST_ROW = 155
BRANCH_COLUMNS = (
    ("B:E", "技能大项"), ("F:Q", "分项名称"), ("R:T", "基础"),
    ("U:W", "职业点"), ("X:Z", "兴趣点"), ("AA:AC", "经历包点"),
    ("AD:AF", "成长点数"), ("AG:AI", "普通"), ("AJ:AL", "困难"),
    ("AM:AO", "极难"), ("AP:AS", "本职"),
)
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def branch_row(index: int) -> int:
    return BRANCH_FIRST_ROW + index * 2


def selection_title_row(branch_count: int) -> int:
    return branch_row(branch_count) + 1


def _tag(name: str) -> str:
    return f"{{{MAIN}}}{name}"


def _row_layout(sheet, row: int, *, hidden: bool = False, height: int = 26) -> None:
    if hasattr(sheet, "root"):
        sheet.cell(2, row)
        element = sheet.rows[row]
        element.set("ht", str(height))
        element.set("customHeight", "1")
        if hidden:
            element.set("hidden", "1")
    else:
        sheet.Rows(row).RowHeight = height
        sheet.Rows(row).Hidden = hidden


def _merged_cell(sheet, columns: str, row: int, style: str = "F16"):
    left, right = columns.split(":")
    address = f"{left}{row}:{right}{row}"
    target = sheet.Range(address)
    if hasattr(sheet, "root"):
        source_style = sheet.cells[style].get("s", "0")
        start, _, end, _ = range_boundaries(address)
        for column in range(start, end + 1):
            sheet.cell(column, row).set("s", source_style)
        merges = sheet.root.find(_tag("mergeCells"))
        if merges is None:
            merges = ET.Element(_tag("mergeCells"), count="0")
            sheet.root.insert(list(sheet.root).index(sheet.data) + 1, merges)
        ET.SubElement(merges, _tag("mergeCell"), ref=address)
        merges.set("count", str(len(merges)))
    else:
        source = sheet.Range(style)
        target.Merge()
        for name in ("Name", "Size", "Bold", "Color"):
            setattr(target.Font, name, getattr(source.Font, name))
        target.Interior.Color = source.Interior.Color
        target.HorizontalAlignment = source.HorizontalAlignment
        target.VerticalAlignment = source.VerticalAlignment
        target.NumberFormat = source.NumberFormat
    target.WrapText = True
    return sheet.Range(f"{left}{row}")


def _selection_records(draft: CharacterDraft) -> list[tuple[str, str, str]]:
    if draft.occupation.is_custom:
        return [("custom", "", skill.template_slot) for skill in draft.skills
                if skill.selected_occupation and skill.template_slot != "F26"]
    return ([("free", "", slot) for slot in draft.free_skill_choices]
            + [("group", marker, token) for marker, choices in draft.group_choices.items()
               for token in choices])


def write_branch_section(workbook, draft: CharacterDraft, set_value: Callable) -> None:
    """Add only new rows; original template cells, merges and print settings stay."""
    branches = [skill for skill in draft.skills if is_branch_slot(skill.template_slot)]
    if not branches:
        return
    sheet = workbook.Worksheets("人物卡")
    records = _selection_records(draft)
    metadata_row = selection_title_row(len(branches))
    last_row = metadata_row + len(records)
    if last_row > 1048576:
        raise ValueError("新增技能分支超出 XLSX 工作表行数，请减少分支后导出。")
    # Export is always based on the catalog's source template. Refuse to cover
    # future template content if this reserved extension area becomes occupied.
    if hasattr(sheet, "root"):
        for address, cell in sheet.cells.items():
            column, row, _, _ = range_boundaries(address)
            if 2 <= column <= 45 and 151 <= row <= last_row and any(
                cell.find(_tag(kind)) is not None for kind in ("f", "v", "is")
            ):
                raise ValueError(f"模板新增技能补充区域已被占用：人物卡!{address}")
    else:
        existing = sheet.Range(f"B151:AS{last_row}").Value2
        if any(value is not None for row in existing for value in row):
            raise ValueError("模板新增技能补充区域已被占用，请更新补充区位置后导出。")
    set_value(_merged_cell(sheet, "B:AS", 151, "B14"), BRANCH_TITLE)
    set_value(_merged_cell(sheet, "B:AS", 152), "原技能表以外的分支在此独立列出；各类点数已计入上方剩余预算。")
    _row_layout(sheet, 151, height=28)
    _row_layout(sheet, 152, height=28)
    set_value(sheet.Range("B153"), BRANCH_SCHEMA)
    set_value(sheet.Range("AP153"), len(branches))
    _row_layout(sheet, 153, hidden=True)
    for columns, title in BRANCH_COLUMNS:
        set_value(_merged_cell(sheet, columns, 154, "F15"), title)
    _row_layout(sheet, 154, height=28)
    for index, skill in enumerate(branches):
        row = branch_row(index)
        values = [skill.name, skill.specialization, skill.base_value, skill.occupation_points,
                  skill.interest_points, skill.experience_points, skill.extra_final,
                  None, None, None, "★" if skill.selected_occupation else ""]
        for (columns, _), value in zip(BRANCH_COLUMNS, values, strict=True):
            cell = _merged_cell(sheet, columns, row)
            if value is not None:
                set_value(cell, value)
        sheet.Range(f"AG{row}").Formula = f"=SUM(R{row}:AF{row})"
        sheet.Range(f"AJ{row}").Formula = f"=INT(AG{row}/2)"
        sheet.Range(f"AM{row}").Formula = f"=INT(AG{row}/5)"
        _row_layout(sheet, row, height=max(28, ((len(skill.specialization) + 20) // 21) * 14 + 8))
        set_value(sheet.Range(f"B{row + 1}"), "分支标识")
        set_value(sheet.Range(f"F{row + 1}"), skill.template_slot)
        _row_layout(sheet, row + 1, hidden=True)
    set_value(sheet.Range(f"B{metadata_row}"), "导入恢复信息（请保留）")
    set_value(sheet.Range(f"AP{metadata_row}"), len(records))
    _row_layout(sheet, metadata_row, hidden=True)
    for index, (kind, marker, value) in enumerate(records, metadata_row + 1):
        for column, content in (("B", kind), ("F", marker), ("J", value)):
            set_value(sheet.Range(f"{column}{index}"), content)
        _row_layout(sheet, index, hidden=True)
    end = branch_row(len(branches) - 1)
    formula = str(sheet.Range("J50").Formula)
    replacements = {
        "SUM(N16:O49,AJ16:AK49)": f"SUM(N16:O49,AJ16:AK49,U155:W{end})",
        "SUM(P16:Q49,AL16:AM49)": f"SUM(P16:Q49,AL16:AM49,X155:Z{end})",
        "SUM(M16:M49,AI16:AI49)": f"SUM(M16:M49,AI16:AI49,AA155:AC{end})",
    }
    if not all(original in formula for original in replacements):
        raise ValueError("模板技能预算公式与补充分支区域不兼容，请更新模板后导出。")
    for original, replacement in replacements.items():
        formula = formula.replace(original, replacement)
    sheet.Range("J50").Formula = formula
    # There is no print area on the original character sheet. Extending its
    # used range includes visible supplemental rows without shrinking the card
    # or excluding any existing annotations. Metadata rows alone are hidden.
    if hasattr(sheet, "root"):
        dimension = sheet.root.find(_tag("dimension"))
        if dimension is not None:
            left, top, right, bottom = range_boundaries(dimension.get("ref"))
            dimension.set("ref", f"{get_column_letter(left)}{top}:{get_column_letter(max(right, 45))}{max(bottom, last_row)}")
