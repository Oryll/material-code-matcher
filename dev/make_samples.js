/* 生成给用户试用的两份样例文件：编码清单 + 明细表
   同时用 core.js 把「正确答案」算出来回报，方便核对。 */
'use strict';
const fs = require('fs');
const path = require('path');
const root = __dirname;
global.window = global;
const CORE = require(path.join(root, 'core.js'));
const XLSX = require(path.join(root, 'xlsx.full.min.js'));

const outDir = path.join(root, '..', 'samples');
fs.mkdirSync(outDir, { recursive: true });

/* 固定种子，保证每次生成同一份数据 */
let seed = 20261104;
function rnd() {
  seed = (seed * 1103515245 + 12345) & 0x7fffffff;
  return seed / 0x7fffffff;
}
function pick(arr) { return arr[Math.floor(rnd() * arr.length) % arr.length]; }

/* ---------- 料号池：清单与明细共用（命中率约一半） ---------- */
const pool = [];
for (let i = 1; i <= 40; i++) pool.push('A-' + String(1000 + i));   // A-1001…A-1040
for (let i = 1; i <= 40; i++) pool.push('B-' + String(2000 + i));   // B-2001…B-2040
for (let i = 1; i <= 25; i++) pool.push(String(3000 + i));          // 3001…3025
pool.push(
  'GL-M8X20', 'GL-M10X30', 'GL-M12X40', 'GL-M6X16', 'GL-M16X50',
  'ZC-6205', 'ZC-6206', 'ZC-6308', 'ZC-30206',
  'DJ-Y2-4KW', 'DJ-Y3-7.5KW', 'MF-BS-01', 'MF-BS-02', 'MF-GS-11',
  'JG-LX-100', 'JG-LX-150', 'JT-001', 'JT-002'
);
for (let i = 1; i <= 20; i++) pool.push('C-' + String(4000 + i));   // C-4001…C-4020

/* ---------- 1. 物料编码清单 ---------- */
const listEntries = [
  '# 本月盘点用的料号清单（示例）',
  '# 以 # 开头的行会被当成注释跳过',
  '# 清单里有 3 个料号是「写歪了」的，打开不同开关命中数会变，可以对着看'
];
listEntries.push(...pool);

// 三个「写歪了」的料号，只存在于下面这几条里
listEntries.push('a-1050');      // 小写 → 明细里写 A-1050，需勾「忽略大小写」
listEntries.push('X-5002');      // 明细里写 "X 5002"，需勾「忽略空格与分隔符」
listEntries.push('0004567');     // 明细里写数字 4567，需勾「忽略前导零」

listEntries.push('  A-1006  ');  // 前后带空格，默认就能对上
listEntries.push('Z9999');       // 清单里有、明细里没有 → 会进「未命中料号」表

// 重复条目 + 空行
listEntries.push('A-1002', 'ZC-6205', '3008');
listEntries.splice(20, 0, '');
listEntries.splice(90, 0, '');

const listText = listEntries.join('\r\n') + '\r\n';
fs.writeFileSync(path.join(outDir, '物料编码清单.txt'), listText, 'utf8');

/* ---------- 2. 入库明细表 ---------- */
const names = ['六角螺栓', '平垫圈', '弹簧垫圈', '深沟球轴承', '圆锥滚子轴承', '三相异步电机',
  '法兰螺母', '内六角螺钉', '开口销', '挡圈', '同步带', '链轮', '联轴器', '油封',
  '气缸', '电磁阀', '接近开关', '继电器', '变频器', '接触器'];
const specs = ['M6×16', 'M8×20', 'M10×30', 'M12×40', 'M16×50', '6205-2RS', '6206-2Z',
  '6308', '30206', 'Y2-4KW', 'Y3-7.5KW', 'BS-01', 'GS-11', 'LX-100'];
const units = ['个', '件', '套', '米', '公斤'];
const houses = ['一号库', '二号库', '三号库', '车间线边库'];
const suppliers = ['华泰紧固件', '恒力轴承', '中远机电', '天成密封', '瑞达电气'];

const aoa = [['物料编码', '品名', '规格型号', '单位', '库存数量', '单价', '入库日期', '仓库', '供应商']];
const offList = [];
for (let i = 0; i < 150; i++) offList.push('D-' + String(7000 + i));   // 一定不在清单里

const special = [
  ['A-1050', '小写料号（需忽略大小写）'],
  ['X 5002', '空格代替横杠（需忽略空格与分隔符）'],
  [4567, '前导零（需忽略前导零）'],
  ['   B-2015   ', '前后带空格'],
  ['A-1006', '清单里带空格，这里干净']
];

const N = 600;
// 日期从 2024-11-07 往前铺，跨度约半年，看着像真账
let serial = 45600;
const rowsOut = [];
for (let i = 0; i < N; i++) {
  let code;
  if (i < special.length) code = special[i][0];
  else if (rnd() < 0.55) code = pick(pool);
  else code = pick(offList);

  serial -= rnd() < 0.5 ? 0 : 1;
  aoa.push([
    code,
    pick(names),
    pick(specs),
    pick(units),
    Math.floor(rnd() * 5000),
    Math.round(rnd() * 50000) / 100,
    serial,
    pick(houses),
    pick(suppliers)
  ]);
  rowsOut.push(code);
}

const ws = XLSX.utils.aoa_to_sheet(aoa);
// 日期列写成序列号 + 日期格式（绕开两个库对 Date 的时区处理）
for (let r = 1; r < aoa.length; r++) {
  const addr = XLSX.utils.encode_cell({ r, c: 6 });
  ws[addr] = { t: 'n', v: aoa[r][6], z: 'yyyy-mm-dd' };
}
ws['!ref'] = 'A1:I' + aoa.length;
ws['!cols'] = [{ wch: 14 }, { wch: 16 }, { wch: 12 }, { wch: 6 }, { wch: 10 }, { wch: 8 },
  { wch: 12 }, { wch: 12 }, { wch: 14 }];
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, '入库明细');
const bytes = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
fs.writeFileSync(path.join(outDir, '入库明细表.xlsx'), Buffer.from(bytes));

/* ---------- 3. 算正确答案 ---------- */
const parsed = CORE.parseCodeList(listText);
const table = CORE.tableFromAOA(aoa, { header: true, dateCols: { 6: 'yyyy-mm-dd' } });

function report(label, opts) {
  const st = CORE.listStats(parsed, opts);
  const res = CORE.runMatch({ codes: parsed, table, opts, outMode: 'hit' });
  const resAll = CORE.runMatch({ codes: parsed, table, opts, outMode: 'all' });
  console.log('\n【' + label + '】');
  console.log('  清单 ' + st.total + ' 条 → 去重后 ' + st.unique + ' 个，重复 ' + st.dup + ' 条');
  console.log('  明细 ' + table.rows.length + ' 行 × ' + table.width + ' 列');
  console.log('  命中 ' + res.rows.length + ' 行 / 命中的料号 ' + res.hitCodeCount + ' 个 / 未命中 ' + res.unusedCount + ' 个');
  console.log('  全部行模式写出 ' + resAll.rows.length + ' 行，其中标「否」' + (resAll.rows.length - res.rows.length) + ' 行');
  return res;
}

console.log('已生成到：' + outDir);
console.log('  物料编码清单.txt   ' + fs.statSync(path.join(outDir, '物料编码清单.txt')).size + ' B');
console.log('  入库明细表.xlsx    ' + fs.statSync(path.join(outDir, '入库明细表.xlsx')).size + ' B');

const a = report('默认（只勾「忽略大小写」）', { ignoreCase: true, ignoreZero: false, ignoreSpace: false });
const b = report('三个开关全打开', { ignoreCase: true, ignoreZero: true, ignoreSpace: true });
const c = report('三个开关全关闭', { ignoreCase: false, ignoreZero: false, ignoreSpace: false });

console.log('\n三个「写歪了」的料号在明细表里：');
console.log('  第 2 行 A-1050    → 忽略大小写后才命中');
console.log('  第 3 行 X 5002    → 忽略空格与分隔符后才命中');
console.log('  第 4 行 4567      → 忽略前导零后才命中（清单里写的是 0004567）');
console.log('\n清单里没出现的料号（会进「未命中料号」表）：');
console.log('  ' + (a.unusedRaw.length > 10 ? a.unusedRaw.slice(0, 10).join('、') + ' 等 ' + a.unusedCount + ' 个' : a.unusedRaw.join('、')));
