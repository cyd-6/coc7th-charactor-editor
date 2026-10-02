"""Excel enforces worksheet child order even when XML is well formed."""
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET

from coc7_card.exporters.xlsx_template import tag, worksheet_paths
from scripts.merge_fx_template import (
    discard_formula_cache, insert_worksheet_control, mask_top_asset_helpers,
    reveal_annual_rows_and_columns,
)
from coc7_card.template_config import TEMPLATE_PATH


ROOT = Path(__file__).resolve().parents[1]


class FxTemplateXmlTests(unittest.TestCase):
    def test_revealing_annual_data_does_not_change_cell_contents(self):
        with ZipFile(TEMPLATE_PATH) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        root = ET.fromstring(parts[worksheet_paths(parts)['货币汇率']])
        styles = ET.fromstring(parts['xl/styles.xml'])
        before = {cell.get('r'): [ET.tostring(child) for child in cell]
                  for cell in root.iter(tag('c'))}
        for row in root.find(tag('sheetData')):
            if int(row.get('r')) >= 35:
                row.set('hidden', '1')
                row.set('collapsed', '1')
                row.set('outlineLevel', '1')
        for col in root.find(tag('cols')):
            col.set('hidden', '1')
        reveal_annual_rows_and_columns(root)
        mask_top_asset_helpers(root, styles)
        after = {cell.get('r'): [ET.tostring(child) for child in cell]
                 for cell in root.iter(tag('c'))}
        self.assertEqual(before, after)
        for row in root.find(tag('sheetData')):
            if int(row.get('r')) >= 35:
                self.assertFalse(any(row.get(key) for key in ('hidden', 'collapsed', 'outlineLevel')))
        for col in root.find(tag('cols')):
            self.assertEqual(col.get('hidden') == '1', int(col.get('min')) > 20)
        size = len(styles.find(tag('cellXfs')))
        mask_top_asset_helpers(root, styles)
        self.assertEqual(len(styles.find(tag('cellXfs'))), size)

    def test_donor_error_cache_is_not_written_as_invalid_excel_error(self):
        cell = ET.Element(tag('c'),r='A25',t='e',s='12')
        ET.SubElement(cell,tag('f')).text = 'HYPERLINK("#A1","返回")'
        ET.SubElement(cell,tag('v')).text = 'HYPERLINK is not implemented.'
        discard_formula_cache(cell)
        self.assertIsNone(cell.find(tag('v')))
        self.assertIsNone(cell.get('t'))
        self.assertEqual(cell.find(tag('f')).text,'HYPERLINK("#A1","返回")')
        self.assertEqual(cell.get('s'),'12')

    def test_template_controls_follow_merges_and_phonetics(self):
        with ZipFile(TEMPLATE_PATH) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        path = worksheet_paths(parts)['货币汇率']
        root = ET.fromstring(parts[path])
        order = [ET.QName(child).localname for child in root]
        expected = ['sheetData', 'mergeCells', 'phoneticPr',
                    'conditionalFormatting', 'dataValidations', 'pageMargins', 'pageSetup']
        self.assertEqual(list(dict.fromkeys(name for name in order if name in expected)), expected)
        validations = root.findall(tag('dataValidations'))
        self.assertEqual(len(validations), 1)
        self.assertEqual({child.get('sqref') for child in validations[0]}, {'K1', 'C3', 'C5', 'C6'})
        self.assertEqual(validations[0].get('count'), '4')

    def test_inserting_controls_preserves_excel_order(self):
        for include_merges in (True, False):
            for reverse in (True, False):
                with self.subTest(merges=include_merges, reverse=reverse):
                    root = ET.Element(tag('worksheet'))
                    names = ['sheetData', 'sheetProtection', 'autoFilter']
                    if include_merges:
                        names += ['mergeCells', 'phoneticPr']
                    names += ['pageMargins', 'pageSetup', 'drawing', 'extLst']
                    original = [ET.SubElement(root, tag(name)) for name in names]
                    controls = ['conditionalFormatting', 'dataValidations']
                    for name in reversed(controls) if reverse else controls:
                        insert_worksheet_control(root, ET.Element(tag(name)))
                    expected = names[:names.index('pageMargins')] + controls + names[names.index('pageMargins'):]
                    self.assertEqual([ET.QName(child).localname for child in root], expected)
                    self.assertEqual([child for child in root if child in original], original)

    def test_multiple_conditional_formats_stay_before_validation(self):
        root = ET.Element(tag('worksheet'))
        for name in ('sheetData', 'mergeCells', 'dataValidations', 'pageMargins'):
            ET.SubElement(root, tag(name))
        for address in ('K8', 'K9'):
            insert_worksheet_control(root, ET.Element(tag('conditionalFormatting'), sqref=address))
        self.assertEqual([ET.QName(child).localname for child in root],
                         ['sheetData', 'mergeCells', 'conditionalFormatting',
                          'conditionalFormatting', 'dataValidations', 'pageMargins'])


if __name__ == '__main__':
    unittest.main()
