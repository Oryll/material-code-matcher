/* 压力测试：大清单 + 大明细表下的耗时与导出可用性 */
'use strict';
const fs = require('fs');
const path = require('path');
const root = __dirname;
global.window = global;
const CORE = require(path.join(root, 'core.js'));
const XLSX = require(path.join(root, 'xlsx.full.min.js'));
const ExcelJS = require(path.join(root, 'exceljs.min.js'));

const N_LIST = 50000;
const N_ROWS = 120000;
const OPTS = { ignoreCase: true, ignoreZero: false, ignoreSpace: false };

// 造清单：M000001 …
const codes = [];
for (let i = 0; i < N_LIST; i++) codes.push('M' + String(i).padStart(6, '0'));

// 造明细：交错有命中/未命中
const aoa = [['物料编码', '品名', '规格', '库存', '单价', '入库日期', '备注']];
for (let i = 0; i < N_ROWS; i++) {
  const n = i % 3 === 0 ? i : i + 1000000; // 每 3 行有 2 行不在清单里
  aoa.push(['M' + String(n).padStart(6, '0'), '品名' + i, 'M' + (i % 20), i % 999, (i % 500) / 100, 45000 + (i % 400), '']);
}

let t = Date.now();
const stats = CORE.listStats(codes, OPTS);
console.log('清单条目 ' + codes.length + '，唯一 ' + stats.unique + '，耗时 ' + (Date.now() - t) + ' ms');

t = Date.now();
const table = CORE.tableFromAOA(aoa, { header: true, dateCols: {} });
console.log('明细表 ' + table.rows.length + ' 行 × ' + table.width + ' 列，耗时 ' + (Date.now() - t) + ' ms');

t = Date.now();
const res = CORE.runMatch({ codes, table, opts: OPTS, outMode: 'hit' });
console.log('比对：命中 ' + res.rows.length + ' 行 / 命中料号 ' + res.hitCodeCount + '，耗时 ' + (Date.now() - t) + ' ms');

t = Date.now();
const resAll = CORE.runMatch({ codes, table, opts: OPTS, outMode: 'all' });
console.log('全部行模式：' + resAll.rows.length + ' 行，耗时 ' + (Date.now() - t) + ' ms');

(async () => {
  t = Date.now();
  const buf = await CORE.buildExport({ ExcelJS, XLSX }, {
    table, result: res, opts: OPTS,
    meta: { listName: '压测清单.txt', tableName: '压测明细.xlsx', sheetName: 'Sheet1' }
  });
  console.log('导出（命中模式）：' + Math.round(buf.length / 1024) + ' KB，耗时 ' + (Date.now() - t) + ' ms');

  t = Date.now();
  const buf2 = await CORE.buildExport({ ExcelJS, XLSX }, {
    table, result: resAll, opts: OPTS,
    meta: { listName: '压测清单.txt', tableName: '压测明细.xlsx', sheetName: 'Sheet1' }
  });
  console.log('导出（全部模式）：' + Math.round(buf2.length / 1024) + ' KB，耗时 ' + (Date.now() - t) + ' ms');

  // 回读校验：行数、列数、日期
  const back = XLSX.read(buf2, { type: 'buffer', cellDates: false, cellNF: true });
  const a1 = XLSX.utils.sheet_to_json(back.Sheets['比对结果'], { header: 1 });
  console.log('回读表头：' + JSON.stringify(a1[0]));
  console.log('回读行数：' + (a1.length - 1) + '（应为 ' + resAll.rows.length + '）');
  const last = a1[a1.length - 1];
  console.log('最后一行前 3 列：' + JSON.stringify(last.slice(0, 3)) + '，末列：' + JSON.stringify(last[last.length - 1]));
  const ws = back.Sheets['比对结果'];
  const f2 = ws['F2'];
  console.log('日期单元格 F2：v=' + f2.v + ' z=' + f2.z + ' → ' + CORE.dateText(f2.v));
  fs.writeFileSync(path.join(root, '_out_stress.xlsx'), Buffer.from(buf2));
})();
