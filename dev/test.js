/* core.js 校验：解析 → 归一化 → 比对 → 导出 */
'use strict';
const fs = require('fs');
const path = require('path');
const root = __dirname;

global.window = global;
const CORE = require(path.join(root, 'core.js'));
const XLSX = require(path.join(root, 'xlsx.full.min.js'));
const ExcelJS = require(path.join(root, 'exceljs.min.js'));

let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; }
  else { fail++; console.log('  ✗ ' + name + (extra !== undefined ? '  → ' + JSON.stringify(extra) : '')); }
}
function eq(name, a, b) { ok(name + ' (' + JSON.stringify(a) + ' === ' + JSON.stringify(b) + ')', JSON.stringify(a) === JSON.stringify(b)); }

/* ---------- 1. 文本解码 ---------- */
const utf8 = new TextEncoder().encode('A-001\nB002\n');
eq('decode utf8', CORE.decodeText(utf8).text.trim(), 'A-001\nB002');
eq('decode utf8 enc', CORE.decodeText(utf8).encoding, 'UTF-8');

const bom = new Uint8Array([0xEF, 0xBB, 0xBF, ...new TextEncoder().encode('X1')]);
eq('decode bom', CORE.decodeText(bom).text, 'X1');

const gbk = new Uint8Array([0xCE, 0xEF, 0xC1, 0xCF, 0xB1, 0xE0, 0xC2, 0xEB, 0x0A, 0x41, 0x31]);
const dec = CORE.decodeText(gbk);
eq('decode gbk text', dec.text, '物料编码\nA1');
ok('decode gbk encoding 标记', /GB|gb/.test(dec.encoding), dec.encoding);

/* ---------- 2. 清单解析 ---------- */
const raw = [
  '# 这是注释',
  'A-001',
  ' A-002 ',
  '　B003　',
  'C004,D005;E006',
  'F007\tG008',
  '',
  '"H009"',
  '// 另一种注释',
  'A-001',
  'B002'
].join('\r\n');
const codes = CORE.parseCodeList(raw);
eq('parse count', codes.length, 11);
eq('parse first', codes[0], 'A-001');
eq('parse csv split', codes.slice(3, 6), ['C004', 'D005', 'E006']);
eq('parse tab split', codes.slice(6, 8), ['F007', 'G008']);
eq('parse quote strip', codes[8], 'H009');
eq('parse 全角空格 trim', codes[2], 'B003');
eq('parse 跳过注释', codes.indexOf('# 这是注释'), -1);

/* ---------- 3. 归一化 ---------- */
const opt0 = { ignoreCase: false, ignoreZero: false, ignoreSpace: false };
const optAll = { ignoreCase: true, ignoreZero: true, ignoreSpace: true };
eq('norm 基础 trim', CORE.normKey('  Ab-01  ', opt0), 'Ab-01');
eq('norm 忽略大小写', CORE.normKey('Ab-01', optAll), 'AB01');
eq('norm 忽略空格与连字符', CORE.normKey('AB 01', optAll), 'AB01');
eq('norm 忽略前导零', CORE.normKey('00123', optAll), '123');
eq('norm 全零保留', CORE.normKey('000', optAll), '0');
eq('norm 数字同文本', CORE.normKey('123', optAll), CORE.normKey(123, optAll));
eq('norm 全角转半角', CORE.normKey('ＡＢ１', optAll), 'AB1');
eq('norm 零宽字符', CORE.normKey('A\u200BB\uFEFF', opt0), 'AB');
eq('norm 斜杠点号', CORE.normKey('A/B.C', optAll), 'ABC');
eq('norm 空值', CORE.normKey(null, opt0), '');

/* ---------- 4. cellText ---------- */
eq('cellText 小数', CORE.cellText(1.5), '1.5');
eq('cellText 整数', CORE.cellText(100), '100');
eq('cellText 无格式数字当数字', CORE.cellText(45352), '45352');
eq('cellText 日期格式数字', CORE.cellText(45352, 'yyyy-mm-dd'), '2024-03-01');
eq('cellText 日期时间', CORE.cellText(45352.5, 'yyyy-mm-dd hh:mm'), '2024-03-01 12:00');
eq('cellText Date', CORE.cellText(new Date(2024, 0, 5)), '2024-01-05');
eq('cellText 对象 text', CORE.cellText({ text: 'X9' }), 'X9');
eq('cellText richText', CORE.cellText({ richText: [{ text: 'AB' }, { text: 'C' }] }), 'ABC');
eq('cellText 公式结果', CORE.cellText({ result: 42 }), '42');
eq('cellText null', CORE.cellText(null), '');
eq('cellText undefined', CORE.cellText(undefined), '');
eq('cellText 布尔', CORE.cellText(true), 'TRUE');

/* ---------- 5. 日期格式识别 ---------- */
ok('日期格式 yyyy-mm-dd', CORE.isDateFormat('yyyy-mm-dd'));
ok('日期格式 m/d/yy', CORE.isDateFormat('m/d/yy'));
ok('日期格式 中文', CORE.isDateFormat('yyyy"年"m"月"d"日"'));
ok('非日期 0.00', !CORE.isDateFormat('0.00'));
ok('非日期 General', !CORE.isDateFormat('General'));
ok('非日期 @文本', !CORE.isDateFormat('@'));
ok('非日期 #,##0', !CORE.isDateFormat('#,##0'));

/* ---------- 6. tableFromAOA ---------- */
const aoa = [
  ['物料编码', '名称', '数量', '备注'],
  ['A-001', '螺丝', 12, ''],
  ['B002', '螺母', 8, '急'],
  ['C004', '垫片', 30, ''],
  ['D005', '弹簧', 5, ''],
  ['', '空行', 0, ''],
  ['E006', '插销', 2, '']
];
const listKeys = CORE.listStats(codes, optAll).keys;
const hOpt = { listKeys: listKeys, ignoreCase: true, ignoreZero: true, ignoreSpace: true };
ok('表头识别为真', CORE.looksLikeHeader(aoa, hOpt) === true);
ok('无表头识别为假', CORE.looksLikeHeader(aoa.slice(1), hOpt) === false);
// 首行是料号、第二行不是 → 一定没有表头
ok('首行命中即无表头', CORE.looksLikeHeader([['A-001', 'x'], ['ZZ9', 'y']], hOpt) === false);
// 首行不像料号、第二行命中 → 首行是表头
ok('次行命中即有表头', CORE.looksLikeHeader([['编码', '名称'], ['A-001', 'x']], hOpt) === true);

const t1 = CORE.tableFromAOA(aoa, { header: true });
eq('table 行数(去表头)', t1.rows.length, 6);
eq('table 宽度', t1.width, 4);
eq('table 列名', t1.columns[1].title, '名称');
eq('table 首页是编码列', t1.columns[0].isCode, true);
eq('table 首行编码', CORE.cellText(t1.rows[0][0]), 'A-001');

const t2 = CORE.tableFromAOA(aoa, { header: false });
eq('无表头行数', t2.rows.length, 7);
eq('无表头列名', t2.columns[0].title, '第1列');

/* ---------- 7. runMatch ---------- */
const res = CORE.runMatch({ codes, table: t1, opts: optAll, outMode: 'hit' });
eq('命中行数', res.hitRows, 5);
eq('扫描行数', res.scanned, 6);
eq('命中编码数', res.hitCodeCount, 5);
eq('未命中数', res.unusedCount, res.uniqueCodes - 5);
eq('空编码行', res.emptyCodeRows, 1);
eq('结果行数=命中行数', res.rows.length, 5);
eq('结果首行编码', CORE.cellText(res.rows[0].cells[0]), 'A-001');

const codes2 = ['A001', 'b-002', 'C_004'];
const res2 = CORE.runMatch({ codes: codes2, table: t1, opts: optAll, outMode: 'hit' });
eq('宽松模式命中 3 行', res2.hitRows, 3);
const res3 = CORE.runMatch({ codes: codes2, table: t1, opts: opt0, outMode: 'hit' });
eq('严格模式命中 0 行', res3.hitRows, 0);

const resAll = CORE.runMatch({ codes, table: t1, opts: optAll, outMode: 'all' });
eq('all 模式行数=扫描行数', resAll.rows.length, res.scanned);
eq('all 模式标记正确', resAll.rows.map(r => !!r.hit), [true, true, true, true, false, true]);

const resMiss = CORE.runMatch({ codes, table: t1, opts: optAll, outMode: 'miss' });
eq('miss 模式行数=扫描-命中', resMiss.rows.length, res.scanned - res.hitRows);
eq('miss 模式全未命中', resMiss.rows.map(r => !!r.hit), [false]);
eq('miss 模式保留原行号', resMiss.rows[0].index, resAll.rows[4].index);
eq('miss 模式不改统计', [resMiss.hitRows, resMiss.hitCodeCount, resMiss.scanned], [res.hitRows, res.hitCodeCount, res.scanned]);
const resMissNone = CORE.runMatch({ codes, table: t1, opts: optAll, outMode: 'hit' });
eq('hit 模式与 miss 模式互补', resMissNone.rows.length + resMiss.rows.length, resAll.rows.length);

const aoaNum = [['编码', '值'], [1001, 'x'], ['1001', 'y'], [1002, 'z']];
const tNum = CORE.tableFromAOA(aoaNum, { header: true });
const resNum = CORE.runMatch({ codes: ['1001'], table: tNum, opts: optAll, outMode: 'hit' });
eq('数字与文本料号同键', resNum.hitRows, 2);

/* ---------- 8. 统计 ---------- */
const st = CORE.listStats(codes, optAll);
eq('unique', st.unique, codes.length - 1);
eq('dup', st.dup, 1);
eq('total', st.total, codes.length);

/* ---------- 9. 真实 xlsx 往返（含日期列） ---------- */
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([
  ['物料编码', '品名', '规格', '单价', '入库日期'],
  ['A-001', '六角螺丝', 'M6×20', 0.35, null],
  ['B002', '法兰螺母', 'M6', 0.28, null],
  ['Z999', '不在清单里', 'X', 9.9, null],
  ['C004', '平垫圈', 'M6', 0.05, null],
  [1005, '弹簧垫圈', 'M6', 0.08, null]
]);
// 用 Excel 序列号写日期，绕开库的 Date 转换
[45352, 45354, 45355, 45356, 45357].forEach((s, i) => {
  const addr = XLSX.utils.encode_cell({ r: i + 1, c: 4 });
  ws[addr] = { t: 'n', v: s, z: 'yyyy-mm-dd' };
});
ws['!ref'] = 'A1:E6';
XLSX.utils.book_append_sheet(wb, ws, '入库明细');
const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });

const wb2 = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false, cellNF: true });
const ws2 = wb2.Sheets['入库明细'];
const dateCols = CORE.scanDateFormats(ws2, XLSX, 80);
eq('嗅出日期列', JSON.stringify(dateCols), JSON.stringify({ 4: 'yyyy-mm-dd' }));

const aoa2 = XLSX.utils.sheet_to_json(ws2, { header: 1, raw: true, defval: '', blankrows: false });
const t3 = CORE.tableFromAOA(aoa2, { header: true, dateCols: dateCols });
eq('往返 行数', t3.rows.length, 5);
eq('往返 日期列', CORE.cellText(t3.rows[0][4]), '2024-03-01');
eq('往返 日期列2', CORE.cellText(t3.rows[4][4]), '2024-03-06');
eq('往返 日期是 Date', t3.rows[0][4] instanceof Date, true);
eq('往返 单价仍为数字', CORE.cellText(t3.rows[0][3]), '0.35');

const res4 = CORE.runMatch({ codes: ['A-001', 'B002', 'C004', '1005', 'Q1'], table: t3, opts: optAll, outMode: 'hit' });
eq('往返 命中 4 行', res4.hitRows, 4);
eq('未命中料号', res4.unusedRaw, ['Q1']);
eq('未命中数', res4.unusedCount, 1);

/* ---------- 10. 导出 ---------- */
(async () => {
  const outBuf = await CORE.buildExport({ ExcelJS, XLSX }, {
    table: t3, result: res4, opts: optAll,
    meta: { listName: '清单.txt（5 行）', tableName: '入库.xlsx', sheetName: '入库明细' }
  });
  ok('导出 buffer 非空', outBuf && outBuf.byteLength > 2000, outBuf && outBuf.byteLength);
  fs.writeFileSync(path.join(root, '_out_test.xlsx'), Buffer.from(outBuf));

  const back = XLSX.read(new Uint8Array(outBuf), { type: 'array', cellDates: false, cellNF: true });
  console.log('  导出工作表：', JSON.stringify(back.SheetNames));
  ok('导出含结果表', back.SheetNames.length >= 1);
  const sheet = back.Sheets[back.SheetNames[0]];
  const head = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
  console.log('  导出前两行：', JSON.stringify(head.slice(0, 2)));
  eq('导出数据行数(含表头)', head.filter(r => r.length).length, 5);

  // 关键：导出的日期单元格必须还是 2024-03-01 这个序列号，且带日期格式
  const dc = sheet.F2 || sheet['F2'];
  ok('导出日期序列号正确', dc && Math.abs(dc.v - 45352) < 1e-9, dc && dc.v);
  ok('导出日期带格式', dc && CORE.isDateFormat(dc.z), dc && dc.z);
  ok('导出日期显示 2024-03-01', CORE.cellText(dc.v, dc.z) === '2024-03-01', CORE.cellText(dc.v, dc.z));

  // 降级路径
  const outBuf2 = await CORE.buildExport({ ExcelJS: null, XLSX }, {
    table: t3, result: res4, opts: optAll, meta: { listName: 'a', tableName: 'b', sheetName: 'c' }
  });
  ok('降级导出可用', outBuf2 && outBuf2.byteLength > 2000, outBuf2 && outBuf2.byteLength);
  const back2 = XLSX.read(new Uint8Array(outBuf2), { type: 'array', cellDates: false, cellNF: true });
  const s2 = back2.Sheets[back2.SheetNames[0]];
  const dc2 = s2.F2;
  ok('降级日期序列号正确', dc2 && Math.abs(dc2.v - 45352) < 1e-9, dc2 && dc2.v);
  ok('降级日期带格式', dc2 && CORE.isDateFormat(dc2.z), dc2 && dc2.z);

  /* ---------- 11. 未命中行模式导出 ---------- */
  const resM = CORE.runMatch({ codes: ['A-001', 'B002', 'C004', '1005', 'Q1'], table: t3, opts: optAll, outMode: 'miss' });
  eq('miss 命中 1 行', resM.rows.length, 1);
  const bufM = await CORE.buildExport({ ExcelJS, XLSX }, {
    table: t3, result: resM, opts: optAll, meta: { listName: 'a', tableName: 'b', sheetName: 'c' }
  });
  const backM = XLSX.read(new Uint8Array(bufM), { type: 'array', cellDates: false, cellNF: true });
  const rowsM = XLSX.utils.sheet_to_json(backM.Sheets[backM.SheetNames[0]], { header: 1, raw: true, defval: '' });
  eq('miss 导出表头无是否命中', rowsM[0].indexOf('是否命中'), -1);
  eq('miss 导出数据行数(含表头)', rowsM.length, 2);
  eq('miss 导出的是没命中的那行', rowsM[1][1], 'Z999');
  const infoM = XLSX.utils.sheet_to_json(backM.Sheets[backM.SheetNames[2]], { header: 1, raw: true, defval: '' });
  const modeRow = infoM.filter(r => r[0] === '输出范围')[0];
  eq('miss 导出说明写明输出范围', modeRow && modeRow[1], '只要未命中行');
  ok('miss 导出仍带未命中料号表', backM.SheetNames.indexOf('未命中料号') >= 0, JSON.stringify(backM.SheetNames));

  /* 未命中行模式但一行都没列出来（全命中了）：仍应能导出「未命中料号」表 */
  const resM0 = CORE.runMatch({ codes: ['A-001', 'B002', 'Z999', 'C004', '1005', 'Q1'], table: t3, opts: optAll, outMode: 'miss' });
  eq('miss 全命中时行数为 0', resM0.rows.length, 0);
  const bufM0 = await CORE.buildExport({ ExcelJS, XLSX }, {
    table: t3, result: resM0, opts: optAll, meta: { listName: 'a', tableName: 'b', sheetName: 'c' }
  });
  const backM0 = XLSX.read(new Uint8Array(bufM0), { type: 'array', cellDates: false, cellNF: true });
  const rowsM0 = XLSX.utils.sheet_to_json(backM0.Sheets[backM0.SheetNames[0]], { header: 1, raw: true, defval: '' });
  eq('miss 零行导出只有表头', rowsM0.length, 1);
  const unusedM0 = XLSX.utils.sheet_to_json(backM0.Sheets['未命中料号'], { header: 1, raw: true, defval: '' });
  eq('miss 零行导出仍有未命中料号', unusedM0.length, 2);

  console.log('\n  ' + pass + ' 通过 / ' + fail + ' 失败');
  process.exit(fail ? 1 : 0);
})();
