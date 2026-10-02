"""Attribute imports bound worksheet parsing and discard each sheet's XML."""

import io
import zipfile

import pytest

from coc7_card.importers import attributes
from coc7_card.importers.excel import ExcelImportError, MAIN, REL, _Workbook


def workbook(*sheets, shared_strings=None):
    definitions = "".join(f'<sheet name="表{i}" sheetId="{i}" r:id="rId{i}"/>' for i in range(1, len(sheets) + 1))
    relations = "".join(f'<Relationship Id="rId{i}" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets) + 1))
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("xl/workbook.xml", f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>{definitions}</sheets></workbook>')
        archive.writestr("xl/_rels/workbook.xml.rels", f'<Relationships>{relations}</Relationships>')
        if shared_strings is not None:
            archive.writestr("xl/sharedStrings.xml", shared_strings)
        for index, sheet in enumerate(sheets, 1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", sheet)
    return output.getvalue()


def sheet(cells):
    return f'<worksheet xmlns="{MAIN}"><sheetData><row r="1">{cells}</row></sheetData></worksheet>'


ATTRIBUTE_CELLS = '<c r="A1" t="inlineStr"><is><t>STR</t></is></c><c r="B1"><v>60</v></c>'


def test_multiple_large_sheets_do_not_accumulate_xml_or_cell_caches(monkeypatch):
    unrelated = sheet("".join(f'<c r="A{i}" t="inlineStr"><is><t>说明</t></is></c>' for i in range(1, 5001)))
    original_xml = _Workbook.xml
    books = []

    def make_book(*args, **kwargs):
        book = _Workbook(*args, **kwargs)
        books.append(book)
        return book

    def xml(book, path):
        assert not path.startswith("xl/worksheets/"), "Attribute sheets must not become complete XML trees."
        return original_xml(book, path)

    monkeypatch.setattr(attributes, "_Workbook", make_book)
    monkeypatch.setattr(_Workbook, "xml", xml)
    result = attributes.import_attributes(workbook(*([unrelated] * 8), sheet(ATTRIBUTE_CELLS)), "many.xlsx")
    assert result["attributes"] == {"STR": 60}
    assert result["summary"]["sheet"] == "表9"
    assert books[0].sheets == {}
    assert set(books[0].roots).isdisjoint(books[0].paths.values())
    assert books[0].archive.fp is None


@pytest.mark.parametrize("count", [100, 101])
def test_cell_limit_counts_blank_cells_and_accepts_exact_boundary(monkeypatch, count):
    monkeypatch.setattr(attributes, "MAX_TABLE_CELLS", 100)
    cells = ATTRIBUTE_CELLS + "".join(f'<c r="A{i}"/>' for i in range(2, count))
    data = workbook(sheet(cells))
    if count == 100:
        assert attributes.import_attributes(data, "boundary.xlsx")["attributes"] == {"STR": 60}
    else:
        with pytest.raises(ExcelImportError, match="简化表内容过多"):
            attributes.import_attributes(data, "boundary.xlsx")


def test_limit_stops_reading_oversized_xml_and_closes_member(monkeypatch):
    monkeypatch.setattr(attributes, "MAX_TABLE_CELLS", 100)
    large_sheet = sheet("".join(f'<c r="A{i}"/>' for i in range(1, 20001)))
    data = workbook(large_sheet)
    original_read = zipfile.ZipExtFile.read
    observed = []

    def read(source, size=-1):
        result = original_read(source, size)
        if source.name == "xl/worksheets/sheet1.xml":
            observed.append((source, len(result)))
        return result

    monkeypatch.setattr(zipfile.ZipExtFile, "read", read)
    with pytest.raises(ExcelImportError, match="简化表内容过多"):
        attributes.import_attributes(data, "too-many.xlsx")
    assert sum(size for _, size in observed) < len(large_sheet.encode())
    assert all(source.closed for source, _ in observed)


@pytest.mark.parametrize("declaration", ['<!DOCTYPE worksheet>', '<!DOCTYPE worksheet [<!ENTITY x "STR">]>'])
def test_xml_declarations_split_across_read_boundary_rejected(declaration):
    # The first five bytes of the declaration are in the preceding 64 KiB read.
    raw = (" " * (64 * 1024 - 5) + declaration + sheet(ATTRIBUTE_CELLS)).encode()
    with pytest.raises(ExcelImportError, match="不支持的 XML 声明"):
        attributes.import_attributes(workbook(raw), "declaration.xlsx")


def test_duplicate_blank_cells_still_rejected():
    with pytest.raises(ExcelImportError, match="重复的单元格地址"):
        attributes.import_attributes(workbook(sheet(ATTRIBUTE_CELLS + '<c r="A2"/><c r="A2"/>')), "duplicate.xlsx")


def test_shared_rich_strings_and_cached_formula_survive_streaming():
    strings = f'<sst xmlns="{MAIN}"><si><r><t>S</t></r><r><t>TR</t></r></si></sst>'
    cells = '<c r="A1" t="s"><v>0</v></c><c r="B1"><f>1-1</f><v>0</v></c>'
    result = attributes.import_attributes(workbook(sheet(cells), shared_strings=strings), "formula.xlsx")
    assert result["attributes"] == {"STR": 0}
    assert "未执行公式" in result["warnings"][0]


def test_truncated_xml_rejected_after_complete_attribute_cells():
    raw = sheet(ATTRIBUTE_CELLS).removesuffix("</worksheet>")
    with pytest.raises(ExcelImportError, match="无法读取此简化表"):
        attributes.import_attributes(workbook(raw), "truncated.xlsx")
