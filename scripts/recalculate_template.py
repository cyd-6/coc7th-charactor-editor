"""Recalculate with Calc, merging caches into the feature-preserving workbook."""
import argparse
import shutil
import sys
import tempfile
import copy
from zipfile import ZipFile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from coc7_card.exporters.xlsx_template import TemplateWorkbook, tag, worksheet_paths
from coc7_card.exporters.excel_linux import LinuxExcelExporter
from lxml import etree as ET


def errors_in(path):
    with ZipFile(path) as z:
        parts={name:z.read(name) for name in z.namelist()}
    return {(name,c.get('r')):c.find(tag('v')).text for name,part in worksheet_paths(parts).items()
            for c in ET.fromstring(parts[part]).iter(tag('c')) if c.get('t')=='e'}


def recalculate(source, output, baseline=None):
    book = TemplateWorkbook(source)
    try:
        with tempfile.TemporaryDirectory(prefix='coc7-template-calc-') as folder:
            root=Path(folder)
            previous={}
            for name,sheet in book.sheets.items():
                for cell in sheet.cells.values():
                    if cell.find(tag('f')) is not None:
                        previous[(name,cell.get('r'))]=copy.deepcopy(cell)
                        cached=cell.find(tag('v'))
                        if cached is not None: cell.remove(cached)
                        cell.attrib.pop('t',None)
            incoming=root/'template.xlsx'
            book.save(incoming)
            calculated=root/'calculated';calculated.mkdir()
            LinuxExcelExporter._recalculate(shutil.which('libreoffice') or shutil.which('soffice'),incoming,calculated,root/'profile',120)
            errors=errors_in(calculated/incoming.name)
            ignored={}
            if errors and baseline:
                original=root/'baseline.xlsx'
                base=TemplateWorkbook(baseline)
                base_formulas={}
                try:
                    for name,sheet in base.sheets.items():
                        for cell in sheet.cells.values():
                            if cell.find(tag('f')) is not None:
                                base_formulas[(name,cell.get('r'))]=cell.find(tag('f')).text
                                cached=cell.find(tag('v'))
                                if cached is not None: cell.remove(cached)
                                cell.attrib.pop('t',None)
                    base.save(original)
                finally: base.close()
                LinuxExcelExporter._recalculate(shutil.which('libreoffice') or shutil.which('soffice'),original,calculated,root/'baseline-profile',120)
                before=errors_in(calculated/original.name)
                # Empty original cards have pre-existing attribute-ratio errors.
                # Preserve only those exact unchanged formulas and their old
                # caches. Any new error, especially in FX, remains fatal.
                for key,error in errors.items():
                    name,address=key
                    if before.get(key)==error and name!='货币汇率':
                        cell=book.sheets[name].cells[address]
                        prior=previous[key]
                        if cell.find(tag('f')).text!=base_formulas.get(key): continue
                        for child in list(cell): cell.remove(child)
                        cell.attrib.clear();cell.attrib.update(prior.attrib)
                        for child in prior: cell.append(copy.deepcopy(child))
                        ignored[key]=book.sheets[name].cells.pop(address)
                print('Preserved existing blank-card formula states:', sorted(f'{s}!{a}: {errors[(s,a)]}' for s,a in ignored))
            LinuxExcelExporter._merge_calculated_values(book,calculated/incoming.name)
            for (name,address),cell in ignored.items(): book.sheets[name].cells[address]=cell
            book.save(output)
    finally:
        book.close()


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source',type=Path)
    parser.add_argument('output',type=Path)
    parser.add_argument('--baseline',type=Path)
    args=parser.parse_args()
    recalculate(args.source,args.output,args.baseline)
    print(args.output)
