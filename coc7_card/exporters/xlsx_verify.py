"""Verify formulas and their saved results in one bounded-memory XML pass."""
from zipfile import ZipFile

from lxml import etree as ET
from openpyxl.formula.translate import Translator

from .excel import FORMULA_ERRORS, ExcelExportError
from .xlsx_template import tag, worksheet_paths


def _elements(archive, path, name):
    """Release parsed siblings as we advance, including empty worksheet rows."""
    tags = (tag(name), tag('row')) if name == 'c' else tag(name)
    with archive.open(path) as stream:
        for _, node in ET.iterparse(stream, events=('end',), tag=tags,
                                    resolve_entities=False, no_network=True):
            if node.tag == tag(name):
                yield node
            node.clear()
            parent = node.getparent()
            while node.getprevious() is not None:
                del parent[0]


def _value(cell, strings):
    if cell.get('t') == 'inlineStr':
        return ''.join(cell.itertext())
    value = cell.find(tag('v'))
    if value is None:
        return None
    if cell.get('t') == 's' and value.text is not None:
        return strings[int(value.text)]
    return value.text


def verify_workbook(path, sheet_names, template_revision_note=None):
    expected = {
        ('简化卡 骰娘导入', 'T29'): 'IF(W29=0,"",人物卡!AB42)',
        ('简化卡 骰娘导入', 'T34'): 'IF(W34=0,"",人物卡!AB47)',
        **{('附表', address): None for address in ('AT15', 'AT16', 'AT17')},
    }
    found = set()
    revision_matches = template_revision_note in (None, '')
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in ('xl/workbook.xml', 'xl/_rels/workbook.xml.rels')}
        paths = worksheet_paths(parts)
        if tuple(paths) != tuple(sheet_names):
            raise ExcelExportError('导出工作簿的工作表数量或顺序与原模板不一致。')
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            strings = [''.join(t.text or '' for t in item.iter(tag('t')))
                       for item in _elements(archive, 'xl/sharedStrings.xml', 'si')]
        for sheet, part in paths.items():
            shared = {}
            for cell in _elements(archive, part, 'c'):
                address = cell.get('r')
                key = sheet, address
                value = _value(cell, strings)
                if key == ('更新说明', 'P3'):
                    # User-edited templates may have a new legitimate revision
                    # note. Preserve and compare it, not a fixed empty baseline.
                    if (value or '') != (template_revision_note or '') or cell.find(tag('f')) is not None:
                        raise ExcelExportError('导出工作簿仍含模板版本说明或版本说明与当前模板不一致：更新说明!P3')
                    revision_matches = True
                node = cell.find(tag('f'))
                formula = (node.text or '') if node is not None else ''
                if node is not None and node.get('t') == 'shared':
                    index = node.get('si')
                    if formula:
                        shared[index] = Translator('=' + formula, origin=address)
                    elif key in expected and index in shared:
                        formula = shared[index].translate_formula(address)[1:]
                if key in expected:
                    actual = formula.replace("'", '')
                    valid = (actual == expected[key] if expected[key] is not None
                             else '附表!$V$226' in actual and '#REF!' not in actual)
                    if not valid:
                        raise ExcelExportError(f'模板兼容性修复未写入：{sheet}!{address}')
                    found.add(key)
                if '#REF!' in formula:
                    raise ExcelExportError(f'导出结果仍含断裂引用：{sheet}!{address}')
                if cell.get('t') == 'e' or node is not None and value in FORMULA_ERRORS:
                    raise ExcelExportError(f'导出结果含公式错误 {value}：{sheet}!{address}')
        if not revision_matches:
            raise ExcelExportError('导出工作簿缺少当前模板的版本说明：更新说明!P3')
        for sheet, address in expected.keys() - found:
            raise ExcelExportError(f'模板兼容性修复未写入：{sheet}!{address}')
        return len(paths)
