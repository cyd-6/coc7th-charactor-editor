// Author the changed ranges with Artifact Tool; merge only those OOXML ranges
// into the original template so its native Excel features survive intact.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';
import { templatePath } from './template_config.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'test-output/fx-upgrade');
await fs.mkdir(out, { recursive: true });
const mode = process.argv[2] ?? 'build';
if (mode === 'calculator-compact') {
  const source = process.argv[3] ?? templatePath;
  const destination = path.join(root,'test-output/fx-compact');
  await fs.mkdir(destination,{recursive:true});
  const reference = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
  const original = reference.worksheets.getItem('货币汇率');
  const book = Workbook.create();
  const fx = book.worksheets.add('货币汇率');
  for (const [cell,value] of [['A3','当前年份'],['A4','原始金额'],['A7','目标金额'],['E3','']]) fx.getRange(cell).values = [[value]];
  fx.getRange('E3').format.fill = '#FFFFFF';
  // Calculation fixtures are not merged; retain the user's actual inputs.
  for (const cell of ['C4','C5','C6','C9','C12','C13']) fx.getRange(cell).values = original.getRange(cell).values;
  fx.getRange('C4').setNumberFormat('General');
  // With no separate status field visible, report invalid selections in the
  // target amount itself instead of silently leaving an unexplained blank.
  fx.getRange('C7').formulas = [['=IF(C9="可用",IF(C5=C6,ROUND(C4,2),ROUND(C4/C12*C13,2)),C9)']];
  fx.getRange('C7').setNumberFormat('#,##0.00');
  fx.getRange('C7:F7').conditionalFormats.addCustom('$C$9<>"可用"',{fill:'#FCE6DF',font:{color:'#963B28'}});
  book.recalculate();
  console.log((await book.inspect({kind:'region',sheetId:'货币汇率',range:'C7',maxChars:500})).ndjson);
  await (await SpreadsheetFile.exportXlsx(book)).save(path.join(destination,'fx-donor.xlsx'));
  await fs.writeFile(path.join(destination,'fx-manifest.json'),JSON.stringify({
    upgrade:'compact-v1',cells:['A3','A4','A7','E3','C4','C7'],names:[],cardCells:[],appendixCells:[],
  },null,2));
  process.exit(0);
}
if (mode === 'calculator-preview') {
  const reference = await SpreadsheetFile.importXlsx(await FileBlob.load(process.argv[3]));
  // Rendering snapshot only. The delivered file keeps all formulas, and its
  // caches are verified in native Excel (this renderer's MATCH differs).
  const snapshot = JSON.parse(await fs.readFile(process.argv[4], 'utf8'));
  for (const [name,cells] of Object.entries(snapshot)) {
    const sheet = reference.worksheets.getItem(name);
    for (const [address,value] of Object.entries(cells)) {
      if (sheet.getRange(address).formulas[0][0]) sheet.getRange(address).values = [[value]];
    }
  }
  const compact = process.argv[5] === 'compact';
  for (const [sheetName,range,name] of [['货币汇率',compact?'A3:F7':'A1:F27','final-calculator'],['人物卡','B60:U62','final-assets']]) {
    const picture = await reference.render({sheetName,range,scale:1.5,format:'png'});
    await fs.writeFile(path.join(root,`test-output/${compact?'fx-compact':'fx-calculator'}/${name}.png`),new Uint8Array(await picture.arrayBuffer()));
  }
  console.log((await reference.inspect({kind:'region',sheetId:'货币汇率',range:'C3:F9',maxChars:1000,tableMaxRows:7,tableMaxCols:4})).ndjson);
  process.exit(0);
}
if (mode === 'calculator') {
  // A scoped upgrade of the user's latest saved workbook, not a data rebuild.
  const source = process.argv[3] ?? templatePath;
  const destination = path.join(root, 'test-output/fx-calculator');
  await fs.mkdir(destination, { recursive: true });
  const reference = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
  const oldFx = reference.worksheets.getItem('货币汇率');
  const book = Workbook.create();
  const fx = book.worksheets.add('货币汇率');
  const card = book.worksheets.add('人物卡');
  const data = JSON.parse(await fs.readFile(path.join(root, 'assets/rates/annual.json'), 'utf8'));
  const last = 35 + data.quotes.length;
  // These values are a calculation fixture only; the merger leaves the user's
  // complete annual table, sources, named lists and USD inputs untouched.
  fx.getRange(`A35:V${last}`).values = oldFx.getRange(`A35:V${last}`).values;
  const updated = oldFx.getRange('A1').values[0][0] === '货币换算';
  const year = oldFx.getRange('K1').values[0][0];
  const currency = reference.worksheets.getItem('人物卡').getRange('S62').formulas[0][0]
    ? oldFx.getRange('K2').values[0][0]
    : reference.worksheets.getItem('人物卡').getRange('S62').values[0][0];
  const merges = ['A1:F1','A2:F2','A10:F10','A11:F11','A24:F24','A25:F25','A26:F26','A27:F27'];
  for (let row=3;row<=9;row++) merges.push(`A${row}:B${row}`,`C${row}:${row===3?'D':'F'}${row}`);
  merges.push('E3:F3');
  for (let row=12;row<=23;row++) merges.push(`A${row}:B${row}`,`C${row}:F${row}`);
  for (const range of merges) fx.mergeCells(range);
  fx.getRange('A1:F31').format.font = {name:'Microsoft YaHei',size:11,color:'#202C30'};
  fx.getRange('A1:F31').format.verticalAlignment = 'center';
  fx.getRange('A1:F31').format.wrapText = true;
  [10,14,28,12,25.1796875,18].forEach((width,index) => {
    fx.getRangeByIndexes(0,index,31,1).format.columnWidth = width;
  });
  const value = (cell,text) => {fx.getRange(cell).values = [[text]];};
  const formula = (cell,text) => {fx.getRange(cell).formulas = [[text]];};
  value('A1','货币换算');
  fx.getRange('A1:F1').format.font = {name:'Microsoft YaHei',size:16,bold:true};
  value('A2','黄色格可编辑。先选年份，再选择原币种和目标币种。');
  for (const [row,label] of Object.entries({3:'换算年份',4:'输入金额',5:'原币种',6:'目标币种',7:'换算结果',8:'1 原币兑换目标币',9:'换算状态',12:'原币 / 1 美元',13:'目标币 / 1 美元',14:'原币来源',15:'原币期间及口径',16:'原币资料链接',17:'原币报价提示',18:'目标币来源',19:'目标币期间及口径',20:'目标币资料链接',21:'目标币报价提示',22:'主表资产币种',23:'主表资产年份'})) value(`A${row}`,label);
  value('C3',updated ? oldFx.getRange('C3').values[0][0] : year);
  value('C4',updated ? oldFx.getRange('C4').values[0][0] : 1);
  value('C5',updated ? oldFx.getRange('C5').values[0][0] : data.quotes.find(q=>q.year===year && q.code==='USD').name);
  value('C6',updated ? oldFx.getRange('C6').values[0][0] : currency);
  value('E3','1920—2026 年');
  value('A10','算法：金额 ÷ 原币兑美元报价 × 目标币兑美元报价；结果保留两位小数。');
  value('A11','报价与来源');
  value('A24','换算器与主表币种独立。主表资产栏的黄色币种格可直接切换。');
  formula('A25','=HYPERLINK("#\'货币汇率\'!K1","主表资产汇率年份和美元基准：右侧 K 列")');
  formula('A26','=HYPERLINK("#\'货币汇率\'!A35","查看年度数据与完整说明（第 35 行起）")');
  formula('A27','=HYPERLINK("#\'货币汇率\'!X36","查看保留的原始历史参考（右侧 X36）")');
  for (const range of ['C3:D3','C4:F4','C5:F5','C6:F6']) {
    fx.getRange(range).format.fill = '#FFF0BF';
    fx.getRange(range).format.font.color = '#174E83';
  }
  fx.getRange('C7:F7').format.fill = '#E9F0ED';
  fx.getRange('C7:F7').format.font.bold = true;
  fx.getRange('C4').setNumberFormat('#,##0.############');
  fx.getRange('C7').setNumberFormat('#,##0.00');
  for (const cell of ['C8','C12','C13']) fx.getRange(cell).setNumberFormat('General');
  for (const row of [1,11]) fx.getRange(`A${row}:F${row}`).format.borders = {bottom:{style:'thin',color:'#A8B8B0'}};
  const index = (input,helper) => formula(helper,`=IF(COUNTIF($M$36:$M$${last},$C$3&"|"&${input})=0,0,MATCH($C$3&"|"&${input},$M$36:$M$${last},0))`);
  value('U3','换算器原币匹配行'); value('U4','换算器目标匹配行');
  index('$C$5','V3'); index('$C$6','V4');
  const lookup = (column,helper) => `INDEX($${column}$36:$${column}$${last},$${helper})`;
  for (const [cell,helper] of [['C12','V$3'],['C13','V$4']]) {
    formula(cell,`=IF($${helper}=0,"",IF(${lookup('D',helper)}="可用",${lookup('E',helper)},""))`);
  }
  formula('C9','=IF(NOT(ISNUMBER(C3)),"年份须为 1920—2026 的整数",IF(OR(C3<1920,C3>2026,C3<>INT(C3)),"年份须为 1920—2026 的整数",IF(NOT(ISNUMBER(C4)),"请输入数值金额",IF(V3=0,"请重新选择原币种",IF(V4=0,"请重新选择目标币种",IF(OR(NOT(ISNUMBER(C12)),NOT(ISNUMBER(C13))),"缺少报价，请重新选择币种",IF(OR(C12<=0,C13<=0),"报价须大于零","可用")))))))');
  formula('C8','=IF(C9="可用",C13/C12,"")');
  formula('C7','=IF(C9="可用",IF(C5=C6,ROUND(C4,2),ROUND(C4/C12*C13,2)),"")');
  for (const [row,col,helper] of [[14,'G','V$3'],[16,'H','V$3'],[17,'T','V$3'],[18,'G','V$4'],[20,'H','V$4'],[21,'T','V$4']]) {
    formula(`C${row}`,`=IF($${helper}=0,"",IF(${lookup(col,helper)}="","",${lookup(col,helper)}))`);
  }
  for (const [row,helper] of [[15,'V$3'],[19,'V$4']]) formula(`C${row}`,`=IF($${helper}=0,"",${lookup('I',helper)}&"；"&${lookup('F',helper)})`);
  formula('C22',"='人物卡'!$S$62");
  fx.getRange('K1').values = [[year]]; // Not merged; retains current asset year.
  formula('C23','=K1');
  formula('K2',"='人物卡'!$S$62");
  value('K24','在人物卡资产栏 S62 选择币种，K1 修改资产汇率年份；改年后币种失效时重新选择。左上换算器可独立选择两种货币。美元基准不代表逐年购买力。');
  fx.getRange('K2').format.fill = '#E9F0ED';
  fx.getRange('C3').dataValidation = {rule:{type:'whole',operator:'between',formula1:1920,formula2:2026}};
  for (const cell of ['C5','C6']) fx.getRange(cell).dataValidation = {rule:{type:'list',formula1:'INDIRECT("FX_Y"&$C$3)'}};
  fx.getRange('C9:F9').conditionalFormats.addCustom('$C$9<>"可用"',{fill:'#FCE6DF',font:{color:'#963B28'}});
  card.getRange('S61').values = [['币种可选']];
  card.getRange('B60').formulas = [['=IF(S62="","资产","资产（"&S62&"）")']];
  card.getRange('S62').values = [[currency]];
  card.getRange('S62:U62').format.fill = '#FFF0BF';
  card.getRange('S62').dataValidation = {rule:{type:'list',formula1:'COC7_FX_ASSET_OPTIONS'}};
  const heights = [30,32,28,28,42,42,40,32,38,40,28,28,28,44,38,66,64,44,38,66,64,42,28,40,30,30,30];
  heights.forEach((height,index)=> {fx.getRange(`A${index+1}`).format.rowHeight = height;});
  book.recalculate();
  console.log((await book.inspect({kind:'region',sheetId:'货币汇率',range:'A3:F9',maxChars:2000,tableMaxRows:7,tableMaxCols:6})).ndjson);
  const picture = await book.render({sheetName:'货币汇率',range:'A1:F10',scale:1.5,format:'png'});
  await fs.writeFile(path.join(destination,'calculator.png'),new Uint8Array(await picture.arrayBuffer()));
  await (await SpreadsheetFile.exportXlsx(book)).save(path.join(destination,'fx-donor.xlsx'));
  await fs.writeFile(path.join(destination,'fx-manifest.json'),JSON.stringify({
    upgrade:'cross-v1',last,merges,cardCells:['B60','S61','S62'],appendixCells:[],
    names:[{name:'COC7_FX_CALCULATOR_SCHEMA',formula:'"cross-v1"'},
      {name:'COC7_FX_ASSET_OPTIONS',formula:'INDIRECT("FX_Y"&\'货币汇率\'!$K$1)'}],
  },null,2));
  process.exit(0);
}
if (mode === 'preview' || mode === 'verify') {
  const template = process.argv[3] ?? templatePath;
  const book = await SpreadsheetFile.importXlsx(await FileBlob.load(template));
  console.log(book.help('workbook.render', { include: 'index,examples,notes', maxChars: 2800 }).ndjson);
  console.log((await book.inspect({kind:'region', sheetId:'货币汇率', range:'J1:K17', maxChars:2200, tableMaxRows:17, tableMaxCols:2})).ndjson);
  const preview = await book.render({sheetName:'货币汇率', range:'J1:K17', scale:1.4, format:'png'});
  await fs.writeFile(path.join(out, `${mode}-controls.png`), new Uint8Array(await preview.arrayBuffer()));
  if (mode==='verify') {
    // The legacy card uses functions this renderer does not fully support;
    // verify its appearance with LibreOffice's native PDF rendering instead.
    for (const [sheet,range,name] of [['货币汇率','A35:F41','annual-table']]) {
      const picture=await book.render({sheetName:sheet,range,scale:1.5,format:'png'});
      await fs.writeFile(path.join(out,`${mode}-${name}.png`),new Uint8Array(await picture.arrayBuffer()));
    }
  }
  process.exit(0);
}

const data = JSON.parse(await fs.readFile(path.join(root, 'assets/rates/annual.json'), 'utf8'));
const book = Workbook.create();
const fx = book.worksheets.add('货币汇率');
const card = book.worksheets.add('人物卡');
const appendix = book.worksheets.add('附表');
const reference = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
const first = 36, last = first + data.quotes.length - 1;
const label = q => q.name;
const headers = ['年份','币种标识','当年实际币名','报价状态','1 美元兑换本币','统计口径','机构或文献','原始资料链接','观测期间','引用文件','说明与限制','换算说明','匹配键','选择项','原始报价','原始方向','转换倍率','原表来源评级','原表位置','报价提示'];
const rows = data.quotes.map(q => [q.year,q.code,q.name,q.available?'可用':'缺失',q.available?Number(q.rate):null,q.basis,q.source,q.source_url,q.date,data.source_file,q.notes||q.missing_reason,q.conversion,`${q.year}|${label(q)}`,label(q),q.raw_quote?Number(q.raw_quote):null,q.raw_direction,q.scale?Number(q.scale):null,q.source_grade,q.source_cells.join('、')+(q.raw_cell?'；'+q.raw_cell:''),q.warning||q.missing_reason]);
fx.getRange('A35:T35').values = [headers];
fx.getRange(`A${first}:T${last}`).values = rows;
fx.getRange(`A35:T${last}`).format.font = { name:'Noto Sans SC', size:10 };
fx.getRange(`A35:T${last}`).format.verticalAlignment = 'center';
fx.getRange(`A35:T${last}`).format.wrapText = true;
fx.getRange('A35:T35').format.fill = '#344D53';
fx.getRange('A35:T35').format.font = {name:'Noto Sans SC',size:10,bold:true,color:'#FFFFFF'};
fx.getRange('A35:T35').format.horizontalAlignment = 'center';
const widths = [10,14,28,12,25.21875,18,82.88671875,105.109375,31.33203125,32.77734375,68,60,40,52,22,18,18,32,56,58];
widths.forEach((w,i) => {fx.getRangeByIndexes(0,i,last,1).format.columnWidth=w;});
fx.getRange(`E${first}:E${last}`).setNumberFormat('General');
fx.getRange(`O${first}:O${last}`).setNumberFormat('0.###############');
fx.getRange(`Q${first}:Q${last}`).setNumberFormat('0.###############');
fx.getRange(`A35:T${last}`).format.autofitRows();
fx.getRange('A33').values = [['年度汇率 1920—2026（仅使用以下数据换算；缺失未插补）']];
fx.getRange('A33').format.font = {name:'Noto Sans SC',size:12,bold:true};
fx.getRange('A1').values = [['原始历史参考（保留原表；当前换算使用第 35 行起的年度数据）']];

const labels = {1:'汇率年份',2:'目标币种',3:'当年实际币名',4:'统计口径',5:'1 美元兑换本币',6:'观测期间',7:'数据来源',8:'选择状态',9:'消费水平（美元）',10:'现金（美元）',11:'其他资产（美元）',13:'交通工具（美元）',14:'住所（美元）',15:'奢侈品（美元）',16:'股票 / 证券（美元）',17:'其他（美元）',18:'报价提示',19:'原始资料链接',20:'换算说明',21:'说明与限制',22:'引用数据文件',23:'原表位置',24:'使用方法'};
for (const [row,text] of Object.entries(labels)) fx.getRange(`J${row}`).values = [[text]];
fx.getRange('J1:K24').format.font = {name:'Noto Sans SC',size:11};
fx.getRange('J1:K24').format.wrapText = true;
fx.getRange('J1:K24').format.verticalAlignment = 'center';
fx.getRange('J1:J24').format.font = {name:'Noto Sans SC',size:11,bold:true};
fx.getRange('K1').values = [[1920]];
fx.getRange('K2').values = [[label(data.quotes[0])]];
fx.getRange('K1:K2').format.fill = '#FFF0BF';
fx.getRange('K13:K17').values = [[0]];
fx.getRange('K13:K17').format.fill = '#FFF0BF';
fx.getRange('K9:K17').setNumberFormat('#,##0.00');
fx.getRange('K5').setNumberFormat('General');
fx.getRange('U1').values = [['匹配记录（公式）']];
fx.getRange('V1').formulas = [[`=IF(COUNTIF($M$${first}:$M$${last},$K$1&"|"&$K$2)=0,0,MATCH($K$1&"|"&$K$2,$M$${first}:$M$${last},0))`]];
const lookup = col => `INDEX($${col}$${first}:$${col}$${last},$V$1)`;
for (const [row,col] of [[3,'C'],[4,'F'],[6,'I'],[7,'G'],[18,'T'],[19,'H'],[20,'L'],[21,'K'],[23,'S']]) {
  fx.getRange(`K${row}`).formulas = [[`=IF($V$1=0,"",IF(${lookup(col)}="","",${lookup(col)}))`]];
}
fx.getRange('K8').formulas = [[`=IF(NOT(ISNUMBER(K1)),"年份须为 1920—2026 的整数",IF(OR(K1<1920,K1>2026,K1<>INT(K1)),"年份须为 1920—2026 的整数",IF($V$1=0,"请重新选择该年份的币种",IF(${lookup('D')}="可用","可用","缺少报价，请重新选择"))))`]];
fx.getRange('K5').formulas = [[`=IF(K8="可用",${lookup('E')},"")`]];
fx.getRange('K22').values = [[data.source_file]];
fx.getRange('K24').values = [['修改黄色年份和币种。改年后若选择失效，请重新选择币种。美元基准沿用信用评级规则，不代表逐年购买力；下方数据可查看来源。']];
fx.getRange('J25').values = [['年度数据']];
fx.getRange('K25').formulas = [['=HYPERLINK("#\'货币汇率\'!A35","查看全部年度汇率与来源")']];
fx.getRange('J25:K25').format.font = {name:'Noto Sans SC',size:11};
fx.getRange('K8').conditionalFormats.addCustom('K8<>"可用"',{fill:'#FCE6DF',font:{color:'#963B28'}});
fx.getRange('K1').dataValidation = {rule:{type:'whole',operator:'between',formula1:1920,formula2:2026}};
fx.getRange('K2').dataValidation = {rule:{type:'list',formula1:'INDIRECT("FX_Y"&$K$1)'}};
card.getRange('R26').values = [[0]]; // Verification fixture; never merged into the real card.
const credit = "'人物卡'!$R$26";
const base = [
  `LOOKUP(${credit},{0,1,10,50,90,99},{0.5,2,10,50,250,5000})`,
  `IF(${credit}=0,0.5,IF(${credit}=99,50000,${credit}*LOOKUP(${credit},{1,10,50,90},{1,2,5,20})))`,
  `IF(${credit}=0,0,IF(${credit}=99,5000000,${credit}*LOOKUP(${credit},{1,10,50,90},{10,50,500,2000})))`,
];
for (let i=0;i<3;i++) fx.getRange(`K${i+9}`).formulas = [[`=${base[i]}*IF($K$1=2026,20,1)`]];
fx.getRange('J1:K24').format.autofitRows();
fx.getRange('K2').format.rowHeight = 40;
fx.getRange('K19:K24').format.autofitRows();
// Formula changes do not automatically resize rows in Excel. Reserve enough
// height for the longest source note in the supplied dataset.
for (const [row,field] of [[3,'name'],[4,'basis'],[6,'date'],[7,'source'],[18,'warning'],[19,'source_url'],[20,'conversion'],[21,'notes'],[23,'raw_cell']]) {
  const lines = Math.max(1,...data.quotes.map(q => String(q[field]||q.missing_reason||'').split('\n').reduce((n,line)=>n+Math.max(1,Math.ceil([...line].reduce((s,c)=>s+(c.charCodeAt(0)>255?1:0.52),0)/32)),0)));
  fx.getRange(`K${row}`).format.rowHeight = 15 * lines + 6;
}

const cardCells = ['BK37','BO37','S62','I62','O62','L62','B75','F75','J75','N75','R75','N76'];
card.getRange('BK37').formulas = [['=IF(\'货币汇率\'!$K$8="可用",\'货币汇率\'!$K$3,"请选择可用币种")']];
card.getRange('S62').formulas = [['=BK37']];
card.getRange('BO37').formulas = [['=\'货币汇率\'!$K$5']];
for (const [cell,row] of [['I62',9],['O62',10],['L62',11],['B75',13],['F75',14],['J75',15],['N75',16],['R75',17]]) {
  card.getRange(cell).formulas = [[`=IF(ISNUMBER('货币汇率'!$K$5),ROUND('货币汇率'!$K$${row}*'货币汇率'!$K$5,2),"")`]];
}
card.getRange('N76').formulas = [['=IF(ISNUMBER(\'货币汇率\'!$K$5),"剩余资产值："&(L62-SUM(B75:U75)),"剩余资产值：暂无汇率")']];
for (const [cell,main] of [['AE230','I62'],['AF230','O62'],['AG230','L62']]) appendix.getRange(cell).formulas = [[`='人物卡'!${main}`]];
// The legacy appendix multiplies the custom rate and abbreviates asset sums.
// Guard its dependent outputs too, so an unavailable quote stays blank all the
// way through the workbook instead of producing VALUE errors or false zeros.
const moneyGuards = ['AF223','AG223','AF224','AG224','AF225','AG225',
  'AE231','AF231','AG231','AE234','AF234','AG234','AE237','AF237','AG237',
  'AA235','AB235','AC235','AA236','AB236','AC236','AE240','AF240','AG240',
  'AA263','AB263','AC263','AF263','AG263','AH263'];
for (const cell of moneyGuards) {
  const formula = reference.worksheets.getItem('附表').getRange(cell).formulas[0][0];
  if (!formula?.startsWith('=')) throw new Error(`Missing appendix formula: ${cell}`);
  appendix.getRange(cell).formulas = [[formula.replaceAll("'",'').includes("ISNUMBER(货币汇率!$K$5)") ? formula : `=IF(ISNUMBER('货币汇率'!$K$5),${formula.slice(1)},"")`]];
}

const names = [];
names.push({name:'COC7_FX_DATA_SHA256',formula:'"'+createHash('sha256').update(await fs.readFile(path.join(root,'assets/rates/annual.json'))).digest('hex')+'"'});
const listRows = [];
for (let year=1920;year<=2026;year++) {
  const options = data.quotes.filter(q=>q.year===year&&q.available);
  const start = 36 + listRows.length;
  for (const q of options) listRows.push([year,label(q)]);
  names.push({name:`FX_Y${year}`,formula:`'货币汇率'!$V$${start}:$V$${start+options.length-1}`});
}
fx.getRange('U35:V35').values = [['下拉年份','可选币种（供下拉选择）']];
fx.getRange(`U36:V${35+listRows.length}`).values = listRows;
fx.getRange('U1:U35').format.columnWidth=28;
fx.getRange('V1:V35').format.columnWidth=60;
book.recalculate();
console.log((await book.inspect({kind:'region',sheetId:'货币汇率',range:'J1:K11',maxChars:1800,tableMaxRows:11,tableMaxCols:2})).ndjson);
const rendered = await book.render({sheetName:'货币汇率',range:'J1:K24',scale:1.2,format:'png'});
await fs.writeFile(path.join(out,'new-controls.png'),new Uint8Array(await rendered.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(book)).save(path.join(out,'fx-donor.xlsx'));
await fs.writeFile(path.join(out,'fx-manifest.json'),JSON.stringify({names,last,cardCells,appendixCells:['AE230','AF230','AG230',...moneyGuards],source_sha256:data.source_sha256},null,2));
console.log(`Authored ${data.quotes.length} annual records and 107 dropdown lists.`);
