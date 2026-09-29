"""Splice the authored point columns into the original, preserving native OOXML."""
from __future__ import annotations

import argparse
import copy
import json
import sys
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from lxml import etree as ET
from openpyxl.utils.cell import coordinate_to_tuple
from coc7_card.exporters.xlsx_template import TemplateWorkbook, tag, worksheet_paths


def merge(template: Path, output: Path, folder: Path) -> None:
    manifest = json.loads((folder / 'skill-points-manifest.json').read_text())
    with ZipFile(folder / 'skill-points-donor.xlsx') as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    paths = worksheet_paths(parts)
    donor_styles = ET.fromstring(parts['xl/styles.xml'])
    donor_fonts = donor_styles.find(tag('fonts'))
    donor_formats = donor_styles.find(tag('cellXfs'))
    strings = list(ET.fromstring(parts['xl/sharedStrings.xml'])) if 'xl/sharedStrings.xml' in parts else []
    book = TemplateWorkbook(template)
    try:
        styles = book.tree('xl/styles.xml')
        formats, fonts, borders = (styles.find(tag(name)) for name in ('cellXfs', 'fonts', 'borders'))
        cache = {}

        def split_style(anchor, end, donor, header):
            key = (anchor.get('s', '0'), end.get('s', '0'), donor.get('s', '0'), header)
            if key in cache:
                return cache[key]
            # Keep the original shading, number format, protection and borders.
            # Only use the donor's point-column font size and header wrapping.
            style = copy.deepcopy(formats[int(key[0])])
            incoming = donor_formats[int(donor.get('s', '0'))]
            font = copy.deepcopy(fonts[int(style.get('fontId', '0'))])
            new_font = donor_fonts[int(incoming.get('fontId', '0'))]
            for name in ('sz', 'name'):
                current = font.find(tag(name))
                if current is not None:
                    font.remove(current)
                wanted = new_font.find(tag(name))
                if wanted is not None:
                    font.append(copy.deepcopy(wanted))
            style.set('fontId', str(len(fonts)))
            style.set('applyFont', '1')
            fonts.append(font)
            border = copy.deepcopy(borders[int(style.get('borderId', '0'))])
            end_style = formats[int(end.get('s', '0'))]
            end_border = borders[int(end_style.get('borderId', '0'))]
            # The right edge of an original merged pair can live on its end
            # cell. Carry it to each now-independent cell, making the divider.
            for edge in ('left', 'right', 'top', 'bottom'):
                current, wanted = border.find(tag(edge)), end_border.find(tag(edge))
                if wanted is not None and wanted.get('style') and (current is None or not current.get('style')):
                    if current is not None:
                        border.replace(current, copy.deepcopy(wanted))
                    else:
                        border.append(copy.deepcopy(wanted))
            style.set('borderId', str(len(borders)))
            style.set('applyBorder', '1')
            borders.append(border)
            align = style.find(tag('alignment'))
            if align is None:
                align = ET.SubElement(style, tag('alignment'))
            align.set('horizontal', 'center')
            align.set('vertical', 'center')
            align.set('wrapText', '1' if header else '0')
            if not header:
                align.set('shrinkToFit', '1')
            style.set('applyAlignment', '1')
            cache[key] = str(len(formats))
            formats.append(style)
            return cache[key]

        card = book.Worksheets('人物卡')
        merges = card.root.find(tag('mergeCells'))
        for merged in list(merges):
            ref = merged.get('ref')
            if ref in {f'{left}{r}:{right}{r}' for left, right in [('L', 'M'), ('AH', 'AI')] for r in range(15, 50)}:
                merges.remove(merged)
        merges.set('count', str(len(merges)))

        for name, addresses in manifest['cells'].items():
            donor = ET.fromstring(parts[paths[name]])
            donor_cells = {cell.get('r'): cell for cell in donor.iter(tag('c'))}
            sheet = book.Worksheets(name)
            old_cells = {key: copy.deepcopy(value) for key, value in sheet.cells.items()}
            for address in addresses:
                incoming = donor_cells[address]
                row, col = coordinate_to_tuple(address)
                target = sheet.cell(col, row)
                if name == '人物卡' and col in (12, 13, 34, 35) and 15 <= row <= 49:
                    left, right = ('L', 'M') if col < 20 else ('AH', 'AI')
                    target.set('s', split_style(old_cells[f'{left}{row}'], old_cells[f'{right}{row}'], incoming, row == 15))
                    if row > 15:
                        continue  # Retain any existing, explicitly separated input.
                # Transfer authored values/formulas only; other cells keep all
                # of their original styling, comments and native relationships.
                for child in list(target):
                    target.remove(child)
                target.attrib.pop('t', None)
                if incoming.get('t') == 's':
                    inline = copy.deepcopy(strings[int(incoming.find(tag('v')).text)])
                    inline.tag = tag('is')
                    target.set('t', 'inlineStr')
                    target.append(inline)
                else:
                    if incoming.get('t'):
                        target.set('t', incoming.get('t'))
                    for child in incoming:
                        target.append(copy.deepcopy(child))
            if name == '人物卡':
                header = donor.find(tag('sheetData')).find(f"{tag('row')}[@r='15']")
                sheet.rows[15].set('ht', header.get('ht'))
                sheet.rows[15].set('customHeight', '1')
        for collection in (formats, fonts, borders):
            collection.set('count', str(len(collection)))
        names = book.tree('xl/workbook.xml').find(tag('definedNames'))
        for existing in list(names):
            if existing.get('name') == 'COC7_SKILL_POINTS_SCHEMA':
                names.remove(existing)
        ET.SubElement(names, tag('definedName'), name='COC7_SKILL_POINTS_SCHEMA').text = '"separate-v1"'
        output.parent.mkdir(parents=True, exist_ok=True)
        book.save(output)
    finally:
        book.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--template', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--folder', type=Path, default=ROOT / 'test-output/skill-points-upgrade')
    args = parser.parse_args()
    merge(args.template, args.output, args.folder)
    print(args.output)
