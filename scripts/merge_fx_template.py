"""Preserve native OOXML while inserting the Artifact Tool authored ranges."""
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
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter
from coc7_card.exporters.xlsx_template import TemplateWorkbook, tag, worksheet_paths


def merge(template, donor_path, manifest_path, output):
    manifest = json.loads(manifest_path.read_text())
    with ZipFile(donor_path) as z:
        donor = {name:z.read(name) for name in z.namelist()}
    paths = worksheet_paths(donor)
    strings = []
    if 'xl/sharedStrings.xml' in donor:
        strings = list(ET.fromstring(donor['xl/sharedStrings.xml']))
    book = TemplateWorkbook(template)
    try:
        original = book.tree('xl/styles.xml')
        incoming = ET.fromstring(donor['xl/styles.xml'])
        maps = {}
        for element in ('fonts','fills','borders','dxfs'):
            source = incoming.find(tag(element))
            if source is None:
                maps[element] = {}; continue
            target = original.find(tag(element))
            if target is None:
                target = ET.Element(tag(element),count='0')
                before = original.find(tag('tableStyles'))
                original.insert(list(original).index(before) if before is not None else len(original),target)
            maps[element] = {}
            for i,node in enumerate(source):
                maps[element][i] = len(target)
                target.append(copy.deepcopy(node))
            target.set('count',str(len(target)))
        number_map = {}
        numbers = original.find(tag('numFmts'))
        if numbers is None:
            numbers = ET.Element(tag('numFmts'),count='0');original.insert(0,numbers)
        source_numbers = incoming.find(tag('numFmts'))
        if source_numbers is not None:
            for node in source_numbers:
                match = next((n for n in numbers if n.get('formatCode')==node.get('formatCode')),None)
                if match is None:
                    new_id = max([163, *(int(n.get('numFmtId')) for n in numbers)])+1
                    match = copy.deepcopy(node);match.set('numFmtId',str(new_id));numbers.append(match)
                number_map[int(node.get('numFmtId'))] = int(match.get('numFmtId'))
        numbers.set('count',str(len(numbers)))
        formats = original.find(tag('cellXfs'))
        style_map = {}
        for i,node in enumerate(incoming.find(tag('cellXfs'))):
            cloned = copy.deepcopy(node)
            for attr, group in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]:
                if cloned.get(attr) is not None:
                    cloned.set(attr,str(maps[group][int(cloned.get(attr))]))
            num = int(cloned.get('numFmtId','0'))
            cloned.set('numFmtId',str(number_map.get(num,num)))
            cloned.set('xfId','0')
            style_map[i] = len(formats);formats.append(cloned)
        formats.set('count',str(len(formats)))

        for name in ('货币汇率','人物卡','附表'):
            source = ET.fromstring(donor[paths[name]])
            sheet = book.Worksheets(name)
            for cell in source.iter(tag('c')):
                address = cell.get('r')
                if name == '人物卡' and address not in manifest['cardCells']:
                    continue
                if name == '附表' and address not in manifest['appendixCells']:
                    continue
                row_index, column_index = coordinate_to_tuple(address)
                target = sheet.cell(column_index, row_index)
                style = target.get('s','0')
                for child in list(target): target.remove(child)
                target.attrib.clear();target.set('r',address)
                for key,value in cell.attrib.items(): target.set(key,value)
                # Keep the card's existing appearance; style only new FX areas.
                target.set('s',str(style_map[int(cell.get('s','0'))]) if name=='货币汇率' and address!='A1' else style)
                if cell.get('t')=='s':
                    si = strings[int(cell.find(tag('v')).text)]
                    inline = copy.deepcopy(si);inline.tag=tag('is')
                    target.set('t','inlineStr');target.append(inline)
                else:
                    for child in cell: target.append(copy.deepcopy(child))
            if name != '货币汇率': continue
            for row in source.find(tag('sheetData')):
                n = int(row.get('r'))
                if n not in sheet.rows: continue
                target = sheet.rows[n]
                if row.get('ht'):
                    target.set('ht',str(max(float(target.get('ht','15')),float(row.get('ht')))))
                    target.set('customHeight','1')
                if n==2: target.attrib.pop('hidden',None)
            # The currency selector is on row 2, hidden in the legacy sheet.
            sheet.rows[2].attrib.pop('hidden',None)
            view=sheet.root.find(tag('sheetViews'))
            if view is not None and len(view):
                view[0].set('topLeftCell','J1')
                selection=view[0].find(tag('selection'))
                if selection is not None:
                    selection.set('activeCell','K1');selection.set('sqref','K1')
            for col in source.find(tag('cols')):
                start,end = int(col.get('min')),int(col.get('max'))
                for n in range(max(11,start),end+1):
                    sheet.Columns(get_column_letter(n)).ColumnWidth = float(col.get('width','16'))
            for kind in ('dataValidations','conditionalFormatting'):
                # Re-running the import replaces our controls, retaining any
                # other validations and conditional rules in the template.
                for existing in list(sheet.root.findall(tag(kind))):
                    if kind=='conditionalFormatting' and existing.get('sqref')=='K8':
                        sheet.root.remove(existing)
                    elif kind=='dataValidations':
                        for rule in list(existing):
                            if rule.get('sqref') in ('K1','K2'): existing.remove(rule)
                        if not len(existing): sheet.root.remove(existing)
                        else: existing.set('count',str(len(existing)))
                for item in source.findall(tag(kind)):
                    cloned=copy.deepcopy(item)
                    for rule in cloned.iter(tag('cfRule')):
                        if rule.get('dxfId') is not None: rule.set('dxfId',str(maps['dxfs'][int(rule.get('dxfId'))]))
                    for validation in cloned.iter(tag('dataValidation')):
                        validation.set('showErrorMessage','1');validation.set('errorStyle','stop')
                        validation.set('errorTitle','请选择有效的年份和币种')
                        validation.set('error','年份为1920—2026；币种须从当前年份下拉列表选择。')
                        validation.set('showDropDown','0')
                    # These come before page/print setup and after sheet data.
                    previous=sheet.root.find(tag('sheetData'))
                    if kind=='dataValidations':
                        candidates=sheet.root.findall(tag('conditionalFormatting'))
                        if candidates: previous=candidates[-1]
                    sheet.root.insert(list(sheet.root).index(previous)+1,cloned)
            dimension=sheet.root.find(tag('dimension'))
            if dimension is not None: dimension.set('ref',f'A1:V{manifest["last"]}')
        workbook=book.tree('xl/workbook.xml')
        names=workbook.find(tag('definedNames'))
        if names is None:
            names=ET.Element(tag('definedNames'))
            workbook.insert(list(workbook).index(workbook.find(tag('sheets')))+1,names)
        for record in manifest['names']:
            for old in list(names):
                if old.get('name')==record['name']: names.remove(old)
            ET.SubElement(names,tag('definedName'),name=record['name']).text=record['formula']
        output.parent.mkdir(parents=True,exist_ok=True)
        book.save(output)
    finally:
        book.close()


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--template',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--donor',type=Path,default=ROOT/'test-output/fx-upgrade/fx-donor.xlsx')
    parser.add_argument('--manifest',type=Path,default=ROOT/'test-output/fx-upgrade/fx-manifest.json')
    args=parser.parse_args()
    merge(args.template,args.donor,args.manifest,args.output)
    print(args.output)
