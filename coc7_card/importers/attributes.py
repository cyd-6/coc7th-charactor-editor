"""Read a small attribute table without evaluating formulas or other content."""

from __future__ import annotations

import csv
import io
import re
import unicodedata
import zipfile
import zlib
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path

from lxml import etree as ET

from ..models import ATTRIBUTE_KEYS
from .excel import MAX_UPLOAD_BYTES, ExcelImportError, _Workbook, _tag


LABELS = dict(zip(ATTRIBUTE_KEYS, ("力量", "体质", "体型", "敏捷", "外貌", "智力", "意志", "教育", "幸运")))
MAX_TABLE_CELLS = 100_000
MAX_TEXT_CELL_CHARS = 100_000
_LABEL_ALIASES = {
    alias.casefold(): key
    for key, label in LABELS.items()
    for alias in (key, label, f"{label}{key}", f"{key}{label}")
}


@dataclass(frozen=True)
class _Cell:
    value: object = None
    address: str = ""
    formula: bool = False
    error: bool = False

    @property
    def present(self) -> bool:
        return self.formula or self.error or (self.value is not None and str(self.value).strip() != "")


def _attribute_name(value: object) -> str | None:
    if not isinstance(value, str):
        return None
    normalized = unicodedata.normalize("NFKC", value)
    normalized = re.sub(r"[\s()\[\]【】/:：]", "", normalized).casefold()
    return _LABEL_ALIASES.get(normalized)


def _coordinate(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]{0,6})", address)
    if not match:
        raise ExcelImportError("表格含无效单元格地址，请重新另存为 XLSX。")
    column = 0
    for char in match[1]:
        column = column * 26 + ord(char) - 64
    row = int(match[2])
    if row > 1_048_576 or column > 16_384:
        raise ExcelImportError("表格含超出 XLSX 范围的单元格地址。")
    return row, column


def _xlsx_table(book: _Workbook, sheet: str) -> dict[tuple[int, int], _Cell]:
    parser = ET.XMLPullParser(events=("start", "end"), resolve_entities=False, no_network=True, huge_tree=False)
    addresses = set()
    count = 0
    table = {}
    tail = b""
    with book.archive.open(book.paths[sheet]) as source:
        while chunk := source.read(64 * 1024):
            declaration = tail + chunk.upper()
            if b"<!DOCTYPE" in declaration or b"<!ENTITY" in declaration:
                raise ExcelImportError("表格包含不支持的 XML 声明，请重新另存为 XLSX。")
            tail = declaration[-8:]
            parser.feed(chunk)
            for event, cell in parser.read_events():
                if event == "start":
                    if cell.tag == _tag("c"):
                        count += 1
                        if count > MAX_TABLE_CELLS:
                            raise ExcelImportError("简化表内容过多，请仅保留要导入的一名调查员属性。")
                    continue
                if cell.tag == _tag("c"):
                    address = cell.get("r")
                    if address in addresses:
                        raise ExcelImportError("表格包含重复的单元格地址，请重新另存为普通 XLSX。")
                    addresses.add(address)
                    point = _coordinate(address or "")
                    item = _Cell(
                        book.cell_value(cell, address, sheet, report=False),
                        f"{sheet}!{address}",
                        cell.find(_tag("f")) is not None,
                        cell.get("t") == "e",
                    )
                    if item.present:
                        table[point] = item
                if cell.tag in {_tag("c"), _tag("row")}:
                    cell.clear()
                    parent = cell.getparent()
                    if parent is not None:
                        parent.remove(cell)
        parser.close()
    return table


def _text_table(data: bytes, extension: str) -> dict[tuple[int, int], _Cell]:
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        try:
            text = data.decode("gb18030")
        except UnicodeDecodeError as exc:
            raise ExcelImportError("无法读取文本编码，请另存为 UTF-8 的 CSV 或 TSV。") from exc
    if "\x00" in text:
        raise ExcelImportError("文本表格含无效字符，请另存为 UTF-8 的 CSV 或 TSV。")
    delimiter = "\t" if extension == ".tsv" else ","
    table = {}
    count = 0
    try:
        reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter, strict=True)
        for row_number, row in enumerate(reader, 1):
            count += len(row)
            if count > MAX_TABLE_CELLS:
                raise ExcelImportError("简化表内容过多，请仅保留要导入的一名调查员属性。")
            for column, value in enumerate(row, 1):
                if len(value) > MAX_TEXT_CELL_CHARS:
                    raise ExcelImportError("表格单元格内容过长，请精简后重试。")
                if value.strip():
                    table[(row_number, column)] = _Cell(value.strip(), f"第 {row_number} 行第 {column} 列")
    except csv.Error as exc:
        raise ExcelImportError("文本表格格式无效：CSV 请用逗号分隔，TSV 请用制表符分隔。") from exc
    return table


def _value(key: str, cell: _Cell, warnings: list[str]) -> int | None:
    if cell.error or (cell.formula and (cell.value is None or str(cell.value).strip() == "")):
        raise ExcelImportError(f"{LABELS[key]}（{cell.address}）的公式无有效缓存结果，请在 Excel 或 LibreOffice 重算并保存后重试。")
    if not cell.present:
        return None
    number = None
    if not isinstance(cell.value, bool):
        try:
            number = Decimal(str(cell.value).strip())
        except (InvalidOperation, ValueError):
            pass
    if number is None or not number.is_finite() or not 0 <= number <= 300 or number != number.to_integral_value():
        raise ExcelImportError(f"{LABELS[key]}（{cell.address}）必须是 0—300 的整数，请修改后重新导入。")
    if cell.formula:
        message = "属性公式使用文件中已保存的结果，未执行公式；请确认文件已重算保存。"
        if message not in warnings:
            warnings.append(message)
    return int(number)


def _looks_numeric(cell: _Cell) -> bool:
    if cell.formula or cell.error or isinstance(cell.value, (bool, int, float, Decimal)):
        return True
    try:
        Decimal(str(cell.value).strip())
        return True
    except InvalidOperation:
        return False


def _parse_table(table: dict[tuple[int, int], _Cell], sheet: str) -> dict | None:
    labels = [(point, name) for point, cell in table.items() if (name := _attribute_name(cell.value))]
    if not labels:
        return None
    seen = set()
    for _, key in labels:
        if key in seen:
            raise ExcelImportError(f"“{sheet}”中属性“{LABELS[key]}”重复，请每项属性只保留一次。")
        seen.add(key)
    rows = {point[0] for point, _ in labels}
    columns = {point[1] for point, _ in labels}
    warnings = []
    used = {point for point, _ in labels}
    values = {}
    horizontal = len(rows) == 1
    vertical = len(columns) == 1
    if horizontal and vertical:
        (row, column), _ = labels[0]
        right = table.get((row, column + 1), _Cell())
        below = [cell for (r, c), cell in table.items() if r > row and c == column]
        numeric_below = any(_looks_numeric(cell) for cell in below)
        if _looks_numeric(right) and numeric_below:
            raise ExcelImportError(f"“{sheet}”的属性排列不明确，请使用一行标题加一行数值，或“属性、数值”两列。")
        horizontal = bool(below) and (not right.present or (numeric_below and not _looks_numeric(right)))
        vertical = not horizontal
    if not horizontal and not vertical:
        raise ExcelImportError(f"“{sheet}”的属性排列不明确，请使用一行标题加一行数值，或“属性、数值”两列。")
    if horizontal:
        header = next(iter(rows))
        header_columns = {column for row, column in table if row == header}
        value_rows = sorted({row for row, col in table if row > header and col in header_columns})
        if len(value_rows) > 1:
            raise ExcelImportError(f"“{sheet}”含多行人物属性，请每次只导入一名调查员。")
        value_row = value_rows[0] if value_rows else header + 1
        for (row, column), key in labels:
            point = value_row, column
            used.add(point)
            value = _value(key, table.get(point, _Cell(address=f"{sheet} 第 {value_row} 行")), warnings)
            if value is not None:
                values[key] = value
    else:
        label_column = next(iter(columns))
        for (row, column), key in labels:
            if any(r == row and c > label_column + 1 for r, c in table):
                raise ExcelImportError(f"“{sheet}”的属性旁含多列数值，请只保留一名调查员的“属性、数值”两列。")
            point = row, column + 1
            used.add(point)
            value = _value(key, table.get(point, _Cell(address=f"{sheet} 第 {row} 行")), warnings)
            if value is not None:
                values[key] = value
    if not values:
        raise ExcelImportError(f"“{sheet}”未填写可导入的属性值；空白属性会保留网页当前值。")
    ignored = [cell for point, cell in table.items() if point not in used]
    generic_headers = {"属性", "数值", "属性值", "值", "attribute", "value", "attributes", "values"}
    if any(str(cell.value).strip().casefold() not in generic_headers for cell in ignored):
        warnings.append("已忽略姓名、年龄、HP 等非属性字段。")
    if len(values) < len(ATTRIBUTE_KEYS):
        warnings.append("未填写的属性保留网页当前值。")
    return {
        "attributes": {key: values[key] for key in ATTRIBUTE_KEYS if key in values},
        "warnings": warnings,
        "summary": {"attribute_count": len(values), "layout": "horizontal" if horizontal else "vertical", "sheet": sheet},
    }


def import_attributes(data: bytes, filename: str) -> dict:
    """Return only supplied attributes, never a draft or other investigator data."""
    if not isinstance(data, bytes) or not data:
        raise ExcelImportError("请选择非空的简化属性表。")
    if len(data) > MAX_UPLOAD_BYTES:
        raise ExcelImportError("表格文件不能超过 12 MB。")
    extension = Path(filename).suffix.lower()
    if extension not in {".xlsx", ".csv", ".tsv"}:
        raise ExcelImportError("请选择 .xlsx、.csv 或 .tsv 简化属性表。")
    book = None
    try:
        if extension == ".xlsx":
            book = _Workbook(data, require_card=False)
            results = []
            for sheet in book.paths:
                result = _parse_table(_xlsx_table(book, sheet), sheet)
                if result is not None:
                    results.append(result)
        else:
            result = _parse_table(_text_table(data, extension), extension[1:].upper())
            results = [result] if result is not None else []
        if len(results) > 1:
            raise ExcelImportError("工作簿含多个可读取属性的工作表，请仅保留一名调查员的属性工作表。")
        if not results:
            raise ExcelImportError("未找到属性标题。请使用力量、体质、体型、敏捷、外貌、智力、意志、教育、幸运，或 STR、CON、SIZ、DEX、APP、INT、POW、EDU、Luck。")
        return results[0]
    except ExcelImportError:
        raise
    except (zipfile.BadZipFile, zlib.error, KeyError, ET.XMLSyntaxError, ValueError, TypeError, OverflowError, OSError, RuntimeError) as exc:
        raise ExcelImportError("无法读取此简化表，请确认文件未损坏、未加密，且使用 XLSX、CSV 或 TSV 格式。") from exc
    finally:
        if book is not None:
            book.archive.close()
