/* 生成端到端测试用的样例文件 */
'use strict';
const fs = require('fs');
const path = require('path');
const root = __dirname;
const XLSX = require(path.join(root, 'xlsx.full.min.js'));

const dir = path.join(root, 'testdata');
fs.mkdirSync(dir, { recursive: true });

/* 清单：故意混入重复、前后空格、全角空格、大小写差异、逗号分隔 */
const list = [
  '# 本月要盘点的料号',
  'A-001',
  'A-002',
  '   A-003   ',
  'a-004',
  'B002',
  'C004,D005',
  'Ｆ007',
  '',
  'Z999',
  'A-001'
].join('\r\n');
fs.writeFileSync(path.join(dir, '清单.txt'), list, 'utf8');

/* 明细：第一列料号，后面随意；含日期列、数字列、空格行 */
const aoa = [
  ['物料编码', '品名', '规格', '库存', '单价', '入库日期', '备注'],
  ['A-001', '六角螺栓', 'M6×20', 1200, 0.35, 45352, '常备'],
  ['A-003', '平垫圈', 'M6', 5000, 0.05, 45353, ''],
  ['A-004', '弹簧垫圈', 'M6', 800, 0.08, 45354, ''],
  ['B002', '法兰螺母', 'M6', 3000, 0.28, 45355, ''],
  ['C004', '内六角螺钉', 'M8×30', 150, 0.62, 45356, '慢动'],
  ['D005', '开口销', '3×20', 9000, 0.02, 45357, ''],
  ['E006', '圆柱销', '5×20', 400, 0.12, 45358, ''],
  ['F007', '挡圈', 'M8', 700, 0.18, 45359, ''],
  ['Z999', '不在清单', 'X', 1, 9.9, 45360, '应被排除'],
  ['A-002', '十字螺钉', 'M4×12', 2600, 0.21, 45361, ''],
  [1005, '数字料号', 'PH1', 33, 1.1, 45362, '测试纯数字料号']
];
const ws = XLSX.utils.aoa_to_sheet(aoa);
// 日期列写成序列号 + 日期格式，绕开库的 Date 时区问题
for (let r = 1; r < aoa.length; r++) {
  const addr = XLSX.utils.encode_cell({ r, c: 5 });
  ws[addr] = { t: 'n', v: aoa[r][5], z: 'yyyy-mm-dd' };
}
ws['!ref'] = 'A1:G' + aoa.length;
ws['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 10 }, { wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 10 }];
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, '入库明细');
// 第二个 sheet，用来测切换
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['其他'], ['A-001']]), '其他表');
// 这份 sheetjs 的 fs 被剔掉了，writeFile 不可用，自己写字节
const bytes = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
fs.writeFileSync(path.join(dir, '入库明细.xlsx'), Buffer.from(bytes));

console.log('样例已生成：');
for (const f of fs.readdirSync(dir)) {
  console.log('  ' + f + '  ' + fs.statSync(path.join(dir, f)).size + ' B');
}
