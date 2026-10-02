"""Edit the original XLSX package without dropping Excel-only template features.

The small Excel-style interface deliberately matches the existing field writer.
Layout is read from the same XML that is edited, without a second workbook
object. Only openpyxl's formula translator is used; original x14 validations,
charts and VML stay intact.
"""

from __future__ import annotations

import copy
import io
import posixpath
import zipfile
from pathlib import Path

from lxml import etree as ET
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, range_boundaries

from ..portraits import prepare_portrait


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
ART = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"s": MAIN, "r": REL}


def tag(name: str) -> str:
    return f"{{{MAIN}}}{name}"


def worksheet_paths(parts: dict[str, bytes]) -> dict[str, str]:
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    relations = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    targets = {r.get("Id"): r.get("Target") for r in relations}
    paths = {}
    for sheet in workbook.find(tag("sheets")):
        target = targets[sheet.get(f"{{{REL}}}id")]
        paths[sheet.get("name")] = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
    return paths


class TemplateWorkbook:
    def __init__(self, path: Path):
        with zipfile.ZipFile(path) as archive:
            self.parts = {name: archive.read(name) for name in archive.namelist()}
        self.trees: dict[str, ET._Element] = {}
        self.paths = worksheet_paths(self.parts)
        self.shared_strings = []
        if "xl/sharedStrings.xml" in self.parts:
            self.shared_strings = [
                "".join(node.itertext())
                for node in ET.fromstring(self.parts["xl/sharedStrings.xml"])
            ]
        self.sheets = {
            name: TemplateSheet(self, name, self.tree(part))
            for name, part in self.paths.items()
        }
        # Expand shared formulas before overwriting any member or anchor.
        for name, sheet in self.sheets.items():
            shared = {
                formula.get("si"): Translator("=" + formula.text, origin=cell.get("r"))
                for cell in sheet.cells.values()
                if (formula := cell.find(tag("f"))) is not None
                and formula.get("t") == "shared" and formula.text
            }
            for cell in sheet.cells.values():
                formula = cell.find(tag("f"))
                if formula is not None and formula.get("t") == "shared":
                    translator = shared.get(formula.get("si"))
                    if translator is None:
                        raise ValueError(f"无法展开共享公式：{name}!{cell.get('r')}")
                    expression = translator.translate_formula(cell.get("r"))
                    formula.attrib.clear()
                    formula.text = expression[1:]
        self.style_cache: dict[tuple, str] = {}
        self._remove_calculation_chain()

    def tree(self, part: str) -> ET._Element:
        if part not in self.trees:
            self.trees[part] = ET.fromstring(self.parts[part])
        return self.trees[part]

    def Worksheets(self, name: str) -> TemplateSheet:
        return self.sheets[name]

    def _remove_calculation_chain(self) -> None:
        relations = self.tree("xl/_rels/workbook.xml.rels")
        for relation in list(relations):
            if relation.get("Type", "").endswith("/calcChain"):
                target = relation.get("Target", "")
                path = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
                self.parts.pop(path, None)
                relations.remove(relation)
        types = self.tree("[Content_Types].xml")
        for entry in list(types):
            if entry.get("ContentType", "").endswith("calcChain+xml"):
                types.remove(entry)
        book = self.tree("xl/workbook.xml")
        calculation = book.find(tag("calcPr"))
        if calculation is None:
            calculation = ET.SubElement(book, tag("calcPr"))
        calculation.set("calcMode", "auto")
        calculation.set("fullCalcOnLoad", "1")
        calculation.set("forceFullCalc", "1")

    def set_style(self, cell: ET._Element, property_name: str, value) -> None:
        original = cell.get("s", "0")
        key = (original, property_name, value)
        if key not in self.style_cache:
            styles = self.tree("xl/styles.xml")
            formats = styles.find(tag("cellXfs"))
            style = copy.deepcopy(formats[int(original)])
            if property_name == "numFmt":
                numbers = styles.find(tag("numFmts"))
                if numbers is None:
                    numbers = ET.Element(tag("numFmts"), count="0")
                    styles.insert(0, numbers)
                number = next((f for f in numbers if f.get("formatCode") == value), None)
                if number is None:
                    identifier = max([163, *(int(f.get("numFmtId")) for f in numbers)]) + 1
                    number = ET.SubElement(numbers, tag("numFmt"), numFmtId=str(identifier), formatCode=value)
                    numbers.set("count", str(len(numbers)))
                style.set("numFmtId", number.get("numFmtId"))
                style.set("applyNumberFormat", "1")
            else:
                alignment = style.find(tag("alignment"))
                if alignment is None:
                    alignment = ET.SubElement(style, tag("alignment"))
                alignment.set(property_name, "1" if value else "0")
                style.set("applyAlignment", "1")
            self.style_cache[key] = str(len(formats))
            formats.append(style)
            formats.set("count", str(len(formats)))
        cell.set("s", self.style_cache[key])

    def save(self, path: Path) -> None:
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, data in self.parts.items():
                if name in self.trees:
                    data = ET.tostring(self.trees[name], encoding="UTF-8", xml_declaration=True, standalone=True)
                archive.writestr(name, data)

    def close(self) -> None:
        for sheet in self.sheets.values():
            sheet.Parent = None
            sheet.root = sheet.data = None
            sheet.rows.clear()
            sheet.cells.clear()
        self.sheets.clear()
        self.trees.clear()
        self.parts.clear()
        self.paths.clear()
        self.shared_strings.clear()
        self.style_cache.clear()

    def add_portrait(self, sheet: TemplateSheet, data: bytes) -> None:
        with prepare_portrait(data) as portrait:
            output = io.BytesIO()
            portrait.save(output, format="PNG")
            width, height = portrait.size
        # Fit the same AL3:AS9 portrait box used by the Windows exporter.
        box_width = sum(sheet.column_pixels(column) for column in range(38, 46))
        box_height = sum(sheet.row_pixels(row) for row in range(3, 10))
        scale = min(box_width / width, box_height / height)
        width, height = width * scale, height * scale
        media = "xl/media/coc7-investigator-portrait.png"
        self.parts[media] = output.getvalue()
        sheet_part = self.paths[sheet.name]
        relation_part = posixpath.dirname(sheet_part) + "/_rels/" + posixpath.basename(sheet_part) + ".rels"
        if relation_part not in self.parts:
            self.parts[relation_part] = ET.tostring(ET.Element(f"{{{PACKAGE}}}Relationships", nsmap={None: PACKAGE}))
        relations = self.tree(relation_part)
        drawing_ref = sheet.root.find(tag("drawing"))
        if drawing_ref is not None:
            relation = next(r for r in relations if r.get("Id") == drawing_ref.get(f"{{{REL}}}id"))
            target = relation.get("Target")
            drawing_part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.dirname(sheet_part) + "/" + target)
        else:
            drawing_part = "xl/drawings/coc7-portrait.xml"
            relation_id = "rIdCoc7Portrait"
            ET.SubElement(relations, f"{{{PACKAGE}}}Relationship", Id=relation_id, Type=REL + "/drawing", Target="../drawings/coc7-portrait.xml")
            drawing_ref = ET.Element(tag("drawing"), {f"{{{REL}}}id": relation_id})
            later = sheet.root.find(tag("legacyDrawing"))
            if later is None:
                later = sheet.root.find(tag("extLst"))
            sheet.root.insert(list(sheet.root).index(later) if later is not None else len(sheet.root), drawing_ref)
            self.parts[drawing_part] = ET.tostring(ET.Element(f"{{{DRAWING}}}wsDr", nsmap={"xdr": DRAWING, "a": ART}))
        drawing = self.tree(drawing_part)
        drawing_rel_part = posixpath.dirname(drawing_part) + "/_rels/" + posixpath.basename(drawing_part) + ".rels"
        if drawing_rel_part not in self.parts:
            self.parts[drawing_rel_part] = ET.tostring(ET.Element(f"{{{PACKAGE}}}Relationships", nsmap={None: PACKAGE}))
        drawing_relations = self.tree(drawing_rel_part)
        picture_rel = "rIdCoc7Portrait"
        ET.SubElement(drawing_relations, f"{{{PACKAGE}}}Relationship", Id=picture_rel, Type=REL + "/image", Target="../media/coc7-investigator-portrait.png")
        anchor = ET.SubElement(drawing, f"{{{DRAWING}}}oneCellAnchor")
        origin = ET.SubElement(anchor, f"{{{DRAWING}}}from")
        for key, value in [("col", 37), ("colOff", round((box_width - width) / 2 * 9525)), ("row", 2), ("rowOff", round((box_height - height) / 2 * 9525))]:
            ET.SubElement(origin, f"{{{DRAWING}}}{key}").text = str(value)
        ET.SubElement(anchor, f"{{{DRAWING}}}ext", cx=str(round(width * 9525)), cy=str(round(height * 9525)))
        picture = ET.SubElement(anchor, f"{{{DRAWING}}}pic")
        properties = ET.SubElement(picture, f"{{{DRAWING}}}nvPicPr")
        identifier = max([0, *(int(n.get("id", "0")) for n in drawing.iter(f"{{{DRAWING}}}cNvPr"))]) + 1
        ET.SubElement(properties, f"{{{DRAWING}}}cNvPr", id=str(identifier), name="调查员头像")
        ET.SubElement(properties, f"{{{DRAWING}}}cNvPicPr")
        fill = ET.SubElement(picture, f"{{{DRAWING}}}blipFill")
        ET.SubElement(fill, f"{{{ART}}}blip", {f"{{{REL}}}embed": picture_rel})
        ET.SubElement(ET.SubElement(fill, f"{{{ART}}}stretch"), f"{{{ART}}}fillRect")
        shape = ET.SubElement(picture, f"{{{DRAWING}}}spPr")
        ET.SubElement(ET.SubElement(shape, f"{{{ART}}}prstGeom", prst="rect"), f"{{{ART}}}avLst")
        ET.SubElement(anchor, f"{{{DRAWING}}}clientData")
        types = self.tree("[Content_Types].xml")
        if not any(e.get("Extension") == "png" for e in types):
            ET.SubElement(types, f"{{{CONTENT}}}Default", Extension="png", ContentType="image/png")
        if not any(e.get("PartName") == "/" + drawing_part for e in types):
            ET.SubElement(types, f"{{{CONTENT}}}Override", PartName="/" + drawing_part, ContentType="application/vnd.openxmlformats-officedocument.drawing+xml")


class TemplateSheet:
    def __init__(self, workbook: TemplateWorkbook, name: str, root: ET._Element):
        self.Parent = workbook
        self.name = name
        self.root = root
        self.data = root.find(tag("sheetData"))
        self.rows = {int(row.get("r")): row for row in self.data}
        self.cells = {cell.get("r"): cell for row in self.data for cell in row if cell.tag == tag("c")}

    @property
    def PageSetup(self) -> TemplateSheet:
        return self

    def Range(self, address: str) -> TemplateRange:
        return TemplateRange(self, address)

    def Columns(self, column: str) -> TemplateColumns:
        return TemplateColumns(self, column)

    def cell(self, column: int, row: int) -> ET._Element:
        address = f"{get_column_letter(column)}{row}"
        if address not in self.cells:
            if row not in self.rows:
                element = ET.Element(tag("row"), r=str(row))
                following = next((r for r in self.data if int(r.get("r")) > row), None)
                self.data.insert(list(self.data).index(following) if following is not None else len(self.data), element)
                self.rows[row] = element
            element = ET.Element(tag("c"), r=address)
            row_element = self.rows[row]
            following = next((c for c in row_element if c.tag == tag("c") and range_boundaries(c.get("r"))[0] > column), None)
            row_element.insert(list(row_element).index(following) if following is not None else len(row_element), element)
            self.cells[address] = element
        return self.cells[address]

    def column_pixels(self, column: int) -> float:
        for dimension in self.root.iterfind(f"{tag('cols')}/{tag('col')}"):
            if int(dimension.get("min")) <= column <= int(dimension.get("max")):
                return float(dimension.get("width", "13")) * 7 + 5
        settings = self.root.find(tag("sheetFormatPr"))
        width = float(settings.get("defaultColWidth", "0")) if settings is not None else 0
        return (width or 8.43) * 7 + 5

    def row_pixels(self, row: int) -> float:
        element = self.rows.get(row)
        height = float(element.get("ht", "0")) if element is not None else 0
        if not height:
            settings = self.root.find(tag("sheetFormatPr"))
            height = float(settings.get("defaultRowHeight", "0")) if settings is not None else 0
        return (height or 15) * 4 / 3

    @property
    def PrintArea(self) -> str:
        index = str(list(self.Parent.paths).index(self.name))
        names = self.Parent.tree("xl/workbook.xml").iterfind(f"{tag('definedNames')}/{tag('definedName')}")
        return next((entry.text or "" for entry in names
                     if entry.get("name") == "_xlnm.Print_Area" and entry.get("localSheetId") == index), "")

    @PrintArea.setter
    def PrintArea(self, value: str) -> None:
        book = self.Parent.tree("xl/workbook.xml")
        names = book.find(tag("definedNames"))
        if names is None:
            names = ET.Element(tag("definedNames"))
            book.insert(list(book).index(book.find(tag("sheets"))) + 1, names)
        index = str(list(self.Parent.paths).index(self.name))
        entry = next((n for n in names if n.get("name") == "_xlnm.Print_Area" and n.get("localSheetId") == index), None)
        if entry is None:
            entry = ET.SubElement(names, tag("definedName"), name="_xlnm.Print_Area", localSheetId=index)
        entry.text = "'" + self.name.replace("'", "''") + "'!" + value


class TemplateRange:
    def __init__(self, sheet: TemplateSheet, address: str):
        self.sheet = sheet
        self.address = address
        self.bounds = range_boundaries(address)

    def _cells(self):
        left, top, right, bottom = self.bounds
        for row in range(top, bottom + 1):
            for column in range(left, right + 1):
                yield self.sheet.cell(column, row)

    @property
    def MergeArea(self) -> TemplateRange:
        left, top, right, bottom = self.bounds
        for merged in self.sheet.root.iterfind(f"{tag('mergeCells')}/{tag('mergeCell')}"):
            address = merged.get("ref")
            min_col, min_row, max_col, max_row = range_boundaries(address)
            if min_col <= left <= right <= max_col and min_row <= top <= bottom <= max_row:
                return TemplateRange(self.sheet, address)
        return self

    def ClearContents(self) -> None:
        for cell in self._cells():
            for child in list(cell):
                if child.tag in {tag("f"), tag("v"), tag("is")}:
                    cell.remove(child)
            cell.attrib.pop("t", None)

    @property
    def Value2(self):
        cell = next(self._cells())
        if cell.get("t") == "inlineStr":
            return "".join(cell.find(tag("is")).itertext())
        value = cell.find(tag("v"))
        if value is None or value.text is None:
            return None
        if cell.get("t") == "s":
            return self.sheet.Parent.shared_strings[int(value.text)]
        if cell.get("t") in {"str", "e"}:
            return value.text
        return float(value.text)

    @Value2.setter
    def Value2(self, value) -> None:
        self.ClearContents()
        cell = next(self._cells())
        if value is None:
            return
        if isinstance(value, str):
            cell.set("t", "inlineStr")
            text = ET.SubElement(ET.SubElement(cell, tag("is")), tag("t"))
            text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            text.text = value
        else:
            ET.SubElement(cell, tag("v")).text = str(int(value) if isinstance(value, bool) else value)

    @property
    def Formula(self):
        formula = next(self._cells()).find(tag("f"))
        return "=" + (formula.text or "") if formula is not None else self.Value2

    @Formula.setter
    def Formula(self, value: str) -> None:
        self.ClearContents()
        ET.SubElement(next(self._cells()), tag("f")).text = value.removeprefix("=")

    def _set_style(self, name: str, value) -> None:
        for cell in self._cells():
            self.sheet.Parent.set_style(cell, name, value)

    NumberFormat = property(fset=lambda self, value: self._set_style("numFmt", value))
    WrapText = property(fset=lambda self, value: self._set_style("wrapText", value))
    ShrinkToFit = property(fset=lambda self, value: self._set_style("shrinkToFit", value))


class TemplateColumns:
    def __init__(self, sheet: TemplateSheet, column: str):
        self.sheet = sheet
        self.column = range_boundaries(column + "1")[0]

    def _set_width(self, value: float) -> None:
        columns = self.sheet.root.find(tag("cols"))
        if columns is None:
            columns = ET.Element(tag("cols"))
            self.sheet.root.insert(list(self.sheet.root).index(self.sheet.data), columns)
        replacement = ET.Element(tag("col"), min=str(self.column), max=str(self.column))
        for column in list(columns):
            start, end = int(column.get("min")), int(column.get("max"))
            if start <= self.column <= end:
                replacement.attrib.update(column.attrib)
                replacement.set("min", str(self.column))
                replacement.set("max", str(self.column))
                columns.remove(column)
                if start < self.column:
                    before = copy.deepcopy(column)
                    before.set("max", str(self.column - 1))
                    columns.append(before)
                if self.column < end:
                    after = copy.deepcopy(column)
                    after.set("min", str(self.column + 1))
                    columns.append(after)
        replacement.set("width", str(value))
        replacement.set("customWidth", "1")
        columns.append(replacement)
        columns[:] = sorted(columns, key=lambda c: int(c.get("min")))

    ColumnWidth = property(fset=_set_width)
