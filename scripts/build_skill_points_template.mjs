// Author only the point columns and affected formulas. A local OOXML merge
// preserves the original workbook's Excel-only features and unrelated styles.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';
import { templatePath } from './template_config.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test-output/skill-points-upgrade');
await fs.mkdir(out, { recursive: true });
const mode = process.argv[2] ?? 'build';
if (mode === 'preview' || mode === 'verify') {
  const source = process.argv[3] ?? templatePath;
  const book = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
  console.log((await book.inspect({ kind:'region', sheetId:'人物卡', range:'J15:Q18', maxChars:1500, tableMaxRows:4, tableMaxCols:8 })).ndjson);
  const picture = await book.render({ sheetName:'人物卡', range:'B14:W22', scale:2, format:'png' });
  await fs.writeFile(path.join(out, `${mode}-skills.png`), new Uint8Array(await picture.arrayBuffer()));
  process.exit(0);
}

const book = Workbook.create();
const card = book.worksheets.add('人物卡');
const appendix = book.worksheets.add('附表');
const cells = { '人物卡': [], '附表': [] };
const set = (sheet, address, value, formula=false) => {
  sheet.getRange(address)[formula ? 'formulas' : 'values'] = [[value]];
  cells[sheet.name].push(address);
};
for (const [growth, experience, base, interest, total] of [
  ['L', 'M', 'J', 'Q', 'R'], ['AH', 'AI', 'AF', 'AM', 'AN'],
]) {
  const header = card.getRange(`${growth}15:${experience}15`);
  header.values = [['成长\n点数', '经历\n包点']];
  cells['人物卡'].push(`${growth}15`, `${experience}15`);
  header.format.font = { name:'微软雅黑', size:8 };
  header.format.wrapText = true;
  header.format.rowHeight = 30;
  const input = card.getRange(`${growth}16:${experience}49`);
  input.format.font = { name:'微软雅黑', size:9 };
  input.format.horizontalAlignment = 'center';
  input.format.verticalAlignment = 'center';
  // Blank input cells are copied only for styling; existing growth values are
  // retained by the merge. Never infer how an old combined value was split.
  for (let row=16; row<=49; row++) {
    cells['人物卡'].push(`${growth}${row}`, `${experience}${row}`);
    set(card, `${total}${row}`, `=SUM(${base}${row}:${interest}${row})`, true);
  }
}
set(card, 'B14', '技能表（成长点数与经历包点分别填写；经历包额度不含成长）');
set(card, 'J50', '=IF(M5=0,"","本职属性："&附表!J52&"   剩余职业点="&(附表!J53-SUM(N16:O49,AJ16:AK49))&"   剩余兴趣点="&(AA7*2-SUM(P16:Q49,AL16:AM49)))&IF(F113=附表!B26,"","   剩余经历包点="&IF(ISNUMBER(附表!D34),附表!D34-SUM(M16:M49,AI16:AI49),附表!D34))', true);
set(appendix, 'AC27', '=IF(ISNUMBER(D34),D34-SUM(人物卡!M16:M49,人物卡!AI16:AI49),D34)', true);
book.recalculate();
await (await SpreadsheetFile.exportXlsx(book)).save(path.join(out, 'skill-points-donor.xlsx'));
await fs.writeFile(path.join(out, 'skill-points-manifest.json'), JSON.stringify({cells}, null, 2));
console.log('Authored separate growth and experience inputs, 68 skill totals and two package-budget formulas.');
