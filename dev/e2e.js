/* 用 Edge + CDP 跑真实端到端：上传、比对、导出、截图。
   期望值不写死，先在 Node 里用 CORE 算出真值，再和界面上的数字对账。 */
'use strict';
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const root = __dirname;
const appUrl = 'file:///' + path.join(root, '..', 'index.html').replace(/\\/g, '/');
const outDir = path.join(root, 'shots');
fs.mkdirSync(outDir, { recursive: true });

global.window = global;
const CORE = require(path.join(root, 'core.js'));
const XLSX = require(path.join(root, 'xlsx.full.min.js'));

/* ---------- 先在 Node 里算真值 ---------- */
const listText = fs.readFileSync(path.join(root, 'testdata', '清单.txt'), 'utf8');
const codes = CORE.parseCodeList(listText);
const OPTS = { ignoreCase: true, ignoreZero: true, ignoreSpace: true };
const stats = CORE.listStats(codes, OPTS);

const wbIn = XLSX.read(fs.readFileSync(path.join(root, 'testdata', '入库明细.xlsx')), { type: 'buffer', cellDates: false, cellNF: true });
const wsIn = wbIn.Sheets['入库明细'];
const dateCols = CORE.scanDateFormats(wsIn, XLSX, 80);
const aoaIn = XLSX.utils.sheet_to_json(wsIn, { header: 1, raw: true, defval: '', blankrows: false });
const table = CORE.tableFromAOA(aoaIn, { header: true, dateCols });
const truth = CORE.runMatch({ codes, table, opts: OPTS, outMode: 'hit' });
const truthAll = CORE.runMatch({ codes, table, opts: OPTS, outMode: 'all' });
const truthMiss = CORE.runMatch({ codes, table, opts: OPTS, outMode: 'miss' });

/* 「其他表」只有两行：['其他'] 是标题，['A-001'] 是唯一数据行。
   A-001 在小清单里 → 表头会被判定为 true，数据行应为 1。这里不写死，
   直接让 CORE 算一遍，界面必须与之一致。 */
const aoaOther = XLSX.utils.sheet_to_json(wbIn.Sheets['其他表'], { header: 1, raw: true, defval: '', blankrows: false });
const tableOther = CORE.tableFromAOA(aoaOther, { header: true, dateCols: {} });

const EXPECT = {
  listTotal: stats.total,
  listUnique: stats.unique,
  listDup: stats.dup,
  tableRows: table.rows.length,
  width: table.width,
  hitRows: truth.hitRows,
  hitCodes: truth.hitCodeCount,
  unused: truth.unusedCount,
  scanned: truth.scanned,
  allRows: truthAll.rows.length,
  missRows: truthMiss.rows.length,
  missCodes: truthMiss.rows.map(r => CORE.cellText(r.cells[0])),
  otherRows: tableOther.rows.length,
  pct: Math.round(truth.hitCodeCount / truth.uniqueCodes * 100)
};
console.log('Node 侧真值：', JSON.stringify(EXPECT));

const listB64 = fs.readFileSync(path.join(root, 'testdata', '清单.txt')).toString('base64');
const xlsxB64 = fs.readFileSync(path.join(root, 'testdata', '入库明细.xlsx')).toString('base64');

const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const PORT = 9333;
const profile = path.join(root, '_edgeprofile');
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

class CDP {
  constructor(ws) { this.ws = ws; this.id = 0; this.pending = new Map(); }
  static async connect(url) {
    const ws = new WebSocket(url);
    const c = new CDP(ws);
    ws.addEventListener('message', (ev) => {
      const msg = JSON.parse(ev.data);
      if (msg.id && c.pending.has(msg.id)) {
        const { res, rej } = c.pending.get(msg.id);
        c.pending.delete(msg.id);
        msg.error ? rej(new Error(JSON.stringify(msg.error))) : res(msg.result);
      } else if (msg.method && c.onEvent) c.onEvent(msg);
    });
    await new Promise((res, rej) => {
      ws.addEventListener('open', res, { once: true });
      ws.addEventListener('error', rej, { once: true });
    });
    return c;
  }
  send(method, params) {
    const id = ++this.id;
    return new Promise((res, rej) => {
      this.pending.set(id, { res, rej });
      this.ws.send(JSON.stringify({ id, method, params: params || {} }));
      setTimeout(() => {
        if (this.pending.has(id)) { this.pending.delete(id); rej(new Error('CDP 超时: ' + method)); }
      }, 30000);
    });
  }
  async eval(expr, awaitPromise) {
    const r = await this.send('Runtime.evaluate', {
      expression: expr, returnByValue: true, awaitPromise: !!awaitPromise, userGesture: true
    });
    if (r.exceptionDetails) {
      const d = r.exceptionDetails;
      throw new Error('页面异常: ' + ((d.exception && d.exception.description) || d.text));
    }
    return r.result && r.result.value;
  }
  async shot(name) {
    const r = await this.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true });
    fs.writeFileSync(path.join(outDir, name), Buffer.from(r.data, 'base64'));
  }
}

let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + '  实得 ' + JSON.stringify(extra)); }
}
/* 勾选某个输出范围单选，并派发 change */
function pickMode(cdp, mode) {
  return cdp.eval('(function(){var e=document.querySelector(\'input[name="outmode"][value="' + mode +
    '"]\');e.checked=true;e.dispatchEvent(new Event("change",{bubbles:true}));return e.value;})()');
}

(async () => {
  try { fs.rmSync(profile, { recursive: true, force: true }); } catch (e) {}
  const edge = spawn(EDGE, [
    '--headless=new', '--remote-debugging-port=' + PORT,
    '--user-data-dir=' + profile, '--no-first-run', '--no-default-browser-check',
    '--disable-gpu', '--hide-scrollbars', '--window-size=1440,1000',
    '--force-device-scale-factor=1', '--allow-file-access-from-files', 'about:blank'
  ], { stdio: 'ignore' });

  let wsUrl = null;
  for (let i = 0; i < 60; i++) {
    await sleep(500);
    try {
      const tabs = await (await fetch('http://127.0.0.1:' + PORT + '/json/list')).json();
      const page = tabs.find(t => t.type === 'page');
      if (page && page.webSocketDebuggerUrl) { wsUrl = page.webSocketDebuggerUrl; break; }
    } catch (e) {}
  }
  if (!wsUrl) { console.log('无法连接 Edge'); process.exit(1); }

  const cdp = await CDP.connect(wsUrl);
  await cdp.send('Page.enable');
  await cdp.send('Runtime.enable');
  await cdp.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: outDir });
  cdp.onEvent = (msg) => {
    if (msg.method === 'Runtime.consoleAPICalled' && msg.params.type === 'error') {
      console.log('  [console.error] ' + msg.params.args.map(a => a.value || a.description).join(' '));
    }
    if (msg.method === 'Runtime.exceptionThrown') {
      console.log('  [页面未捕获异常] ' + JSON.stringify(msg.params.exceptionDetails.text));
    }
  };

  await cdp.send('Page.navigate', { url: appUrl });
  await sleep(2500);

  console.log('\n— 加载 —');
  ok('页面标题存在', !!(await cdp.eval('document.title')));
  ok('库已就绪', await cdp.eval('typeof XLSX === "object" && typeof ExcelJS === "object"'));
  ok('core 已就绪', await cdp.eval('typeof CORE === "object" && typeof CORE.runMatch === "function"'));
  ok('内联字体可用', await cdp.eval('document.fonts.size > 0'), await cdp.eval('document.fonts.size'));
  ok('无外部资源依赖', await cdp.eval('performance.getEntriesByType("resource").filter(r=>/^https?:/.test(r.name)).length') === 0,
    await cdp.eval('performance.getEntriesByType("resource").filter(r=>/^https?:/.test(r.name)).length'));
  await cdp.shot('01-空态.png');

  const inject = (inputId, b64, name, mime) => cdp.eval(`(function(){
    const bin = atob(${JSON.stringify(b64)});
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const f = new File([bytes], ${JSON.stringify(name)}, { type: ${JSON.stringify(mime)} });
    const dt = new DataTransfer(); dt.items.add(f);
    const el = document.getElementById(${JSON.stringify(inputId)});
    el.files = dt.files;
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  })()`);

  /* 用假的文件句柄接管 showSaveFilePicker：既不弹窗，又能把页面写进去的字节扣下来核对 */
  const installFakePicker = () => cdp.eval(`(function(){
    window.__saved = null;
    if (!window.__origPicker) window.__origPicker = window.showSaveFilePicker;
    window.showSaveFilePicker = function(opt){
      window.__pickerOpt = opt;
      return Promise.resolve({
        name: (opt && opt.suggestedName) || 'x.xlsx',
        createWritable: function(){
          var chunks = [];
          return Promise.resolve({
            write: function(d){ chunks.push(d); return Promise.resolve(); },
            close: function(){
              var all = chunks[0];
              var bin = '';
              for (var i = 0; i < all.length; i++) bin += String.fromCharCode(all[i]);
              window.__saved = { name: (opt && opt.suggestedName) || '', size: all.length, b64: btoa(bin) };
              return Promise.resolve();
            }
          });
        }
      });
    };
    return true;
  })()`);

  console.log('\n— 上传清单 txt —');
  await inject('input-list', listB64, '清单.txt', 'text/plain');
  await sleep(700);
  ok('清单区已加载', await cdp.eval('document.getElementById("zone-list").classList.contains("is-loaded")'));
  ok('清单条数 ' + EXPECT.listTotal,
    (await cdp.eval('document.getElementById("list-count").textContent')) === String(EXPECT.listTotal),
    await cdp.eval('document.getElementById("list-count").textContent'));
  ok('重复数 ' + EXPECT.listDup,
    (await cdp.eval('document.getElementById("list-dups").textContent')) === String(EXPECT.listDup),
    await cdp.eval('document.getElementById("list-dups").textContent'));
  ok('编码识别为 UTF-8', /UTF/i.test(await cdp.eval('document.getElementById("list-enc").textContent')));
  ok('预览有内容', (await cdp.eval('document.getElementById("list-peek").textContent')).length > 5);

  console.log('\n— 上传明细 xlsx —');
  await inject('input-table', xlsxB64, '入库明细.xlsx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await sleep(1200);
  ok('明细区已加载', await cdp.eval('document.getElementById("zone-table").classList.contains("is-loaded")'));
  ok('数据行数 ' + EXPECT.tableRows,
    (await cdp.eval('document.getElementById("table-count").textContent')) === String(EXPECT.tableRows),
    await cdp.eval('document.getElementById("table-count").textContent'));
  ok('列数 ' + EXPECT.width,
    (await cdp.eval('document.getElementById("table-cols").textContent')) === String(EXPECT.width),
    await cdp.eval('document.getElementById("table-cols").textContent'));
  ok('多 sheet 时出现下拉', !(await cdp.eval('document.getElementById("sheet-wrap").hidden')));
  ok('sheet 选项数 2', await cdp.eval('document.getElementById("sheet-sel").options.length') === 2);
  ok('表头自动勾上', await cdp.eval('document.getElementById("opt-header").checked'));
  await cdp.shot('02-已上传.png');

  console.log('\n— 比对 —');
  await cdp.eval('document.getElementById("run").click()');
  await sleep(1500);
  ok('结果区可见', !(await cdp.eval('document.getElementById("result").hidden')));
  const got = {
    hitRows: await cdp.eval('document.getElementById("n-rows").textContent'),
    hitCodes: await cdp.eval('document.getElementById("n-codes").textContent'),
    unused: await cdp.eval('document.getElementById("n-unused").textContent'),
    scanned: await cdp.eval('document.getElementById("n-scan").textContent')
  };
  console.log('    界面：命中行 ' + got.hitRows + ' / 命中编码 ' + got.hitCodes +
    ' / 未命中 ' + got.unused + ' / 扫描 ' + got.scanned);
  ok('命中行数 ' + EXPECT.hitRows, got.hitRows === String(EXPECT.hitRows), got.hitRows);
  ok('命中编码数 ' + EXPECT.hitCodes, got.hitCodes === String(EXPECT.hitCodes), got.hitCodes);
  ok('未命中数 ' + EXPECT.unused, got.unused === String(EXPECT.unused), got.unused);
  ok('扫描行数 ' + EXPECT.scanned, got.scanned === String(EXPECT.scanned), got.scanned);
  ok('百分比 ' + EXPECT.pct + '%',
    (await cdp.eval('document.getElementById("rpct").textContent')).includes(String(EXPECT.pct) + '%'),
    await cdp.eval('document.getElementById("rpct").textContent'));
  ok('预览表行数 ' + EXPECT.hitRows,
    await cdp.eval('document.querySelectorAll("#tbl-body tr").length') === EXPECT.hitRows);
  ok('序号从 1 开始', (await cdp.eval('document.querySelector("#tbl-body tr td").textContent')) === '1');

  const tblText = await cdp.eval('document.getElementById("tbl").innerText');
  ok('结果含日期 2024-03-01', tblText.includes('2024-03-01'));
  ok('日期没有漂到前一天', !tblText.includes('2024-02-29'));
  // 未命中的料号不应出现在结果里
  const missed = truth.unusedRaw.slice(0, 3);
  ok('结果里没有未命中料号', missed.every(m => !tblText.includes(m)), missed);
  // 全角 Ｆ007 应该对上了
  ok('全角料号已归一命中', tblText.includes('F007'));
  // 大小写不同的 a-004 应该对上了
  ok('大小写已归一命中', tblText.includes('A-004'));
  await cdp.shot('03-结果.png');

  console.log('\n— 改设置：按钮不能消失 —');
  // 用户反馈：比对完换个方式，「开始比对」按钮就没了，点不到重新比对
  const btnOn = '(function(){var b=document.getElementById("run");return !b.hidden && getComputedStyle(b).display!=="none" && b.offsetWidth>0 && b.offsetHeight>0;})()';
  ok('比对完按钮仍然可见', await cdp.eval(btnOn));
  await cdp.eval('(function(){var e=document.getElementById("opt-zero");e.checked=!e.checked;e.dispatchEvent(new Event("change",{bubbles:true}));})()');
  await sleep(900);
  ok('切换开关后按钮仍然可见', await cdp.eval(btnOn));
  ok('切换开关后结果没有消失', !(await cdp.eval('document.getElementById("result").hidden')));
  ok('切换开关后按钮文案是「重新比对」',
    (await cdp.eval('document.getElementById("run-label").textContent')) === '重新比对',
    await cdp.eval('document.getElementById("run-label").textContent'));
  await cdp.eval('(function(){var b=document.getElementById("run");b.click();})()');
  await sleep(1200);
  ok('手动点「重新比对」能再次出结果', !(await cdp.eval('document.getElementById("result").hidden')));
  // 复原开关，后面的断言仍按原设置算
  await cdp.eval('(function(){var e=document.getElementById("opt-zero");e.checked=!e.checked;e.dispatchEvent(new Event("change",{bubbles:true}));})()');
  await sleep(900);
  ok('复原开关后命中行数回到 ' + EXPECT.hitRows,
    (await cdp.eval('document.getElementById("n-rows").textContent')) === String(EXPECT.hitRows),
    await cdp.eval('document.getElementById("n-rows").textContent'));

  console.log('\n— 全部行模式 —');
  await cdp.eval('(function(){const e=document.querySelector(\'input[name="outmode"][value="all"]\');e.checked=true;e.dispatchEvent(new Event("change",{bubbles:true}));})()');
  await sleep(900);
  // 换模式不再把结果收起来，而是按新设置自动重算
  ok('换模式后结果仍在', !(await cdp.eval('document.getElementById("result").hidden')));
  ok('换模式后自动重算出全部行 ' + EXPECT.allRows,
    (await cdp.eval('document.getElementById("n-rows").textContent')) === String(EXPECT.allRows),
    await cdp.eval('document.getElementById("n-rows").textContent'));
  ok('换模式后「重新比对」按钮还在',
    await cdp.eval('(function(){var b=document.getElementById("run");return !b.hidden && getComputedStyle(b).display!=="none" && b.offsetWidth>0;})()'));
  ok('全部行模式 ' + EXPECT.allRows + ' 行',
    (await cdp.eval('document.getElementById("n-rows").textContent')) === String(EXPECT.allRows),
    await cdp.eval('document.getElementById("n-rows").textContent'));
  const tblAll = await cdp.eval('document.getElementById("tbl").innerText');
  ok('全部行模式含未命中料号', truthAll.rows.some(r => !r.hit) ? missed.every(m => tblAll.includes(m)) : true);
  ok('全部行模式有命中标记列', tblAll.includes('是否命中'));
  const noticeAll = await cdp.eval('document.getElementById("notice").innerText');
  ok('全部行模式文案不再说「已经剔除」', !/剔除/.test(noticeAll), noticeAll.slice(0, 120));
  ok('全部行模式文案提到「否」', /否/.test(noticeAll));
  await cdp.shot('04-全部行.png');

  console.log('\n— 未命中行模式 —');
  await pickMode(cdp, 'miss');
  await sleep(900);
  ok('切到未命中行后自动重算', !(await cdp.eval('document.getElementById("result").hidden')));
  ok('未命中行有说明文案', /未命中|对不上/.test(await cdp.eval('document.getElementById("seg-cap").textContent')));
  await cdp.eval('document.getElementById("run").click()');
  await sleep(1200);
  const nMiss = await cdp.eval('document.getElementById("n-rows").textContent');
  ok('未命中行模式 ' + EXPECT.missRows + ' 行', nMiss === String(EXPECT.missRows), nMiss);
  const lblMiss = await cdp.eval('document.getElementById("n-rows-l").textContent');
  ok('未命中行列名已改', /没命中/.test(lblMiss), lblMiss);
  const tblMiss = await cdp.eval('document.getElementById("tbl").innerText');
  ok('未命中行预览条数一致',
    (await cdp.eval('document.querySelectorAll("#tbl-body tr").length')) === Math.min(100, EXPECT.missRows));
  ok('未命中行全是没命中的', EXPECT.missCodes.every(c => tblMiss.includes(c)), EXPECT.missCodes.slice(0, 3));
  ok('未命中行里没有命中的料号', truth.rows.every(r => !tblMiss.includes(CORE.cellText(r.cells[0]))),
    truth.rows.slice(0, 3).map(r => CORE.cellText(r.cells[0])));
  ok('未命中行没有是否命中列', !/是否命中/.test(tblMiss));
  const noticeMiss = await cdp.eval('document.getElementById("notice").innerText');
  ok('未命中行文案说明只列出对不上的', /对不上|没有命中/.test(noticeMiss), noticeMiss.slice(0, 120));
  await cdp.shot('05-未命中行.png');

  const beforeMiss = new Set(fs.readdirSync(outDir));
  await installFakePicker();
  await cdp.eval('document.getElementById("export").click()');
  let dlMiss = null;
  for (let i = 0; i < 40; i++) {
    await sleep(300);
    dlMiss = await cdp.eval('window.__saved ? window.__saved.name : ""');
    if (dlMiss) break;
  }
  ok('未命中行可导出', !!dlMiss, dlMiss);
  if (dlMiss) {
    const b = Buffer.from(await cdp.eval('window.__saved.b64'), 'base64');
    const backMiss = XLSX.read(b, { type: 'buffer', cellDates: false, cellNF: true });
    const rowsMiss = XLSX.utils.sheet_to_json(backMiss.Sheets[backMiss.SheetNames[0]], { header: 1, raw: true, defval: '' });
    ok('未命中行导出数据行数 ' + EXPECT.missRows, rowsMiss.length === EXPECT.missRows + 1, rowsMiss.length - 1);
    ok('未命中行导出文件名带标记', /未命中行/.test(dlMiss), dlMiss);
  }

  await cdp.eval('(function(){const e=document.querySelector(\'input[name="outmode"][value="hit"]\');e.checked=true;e.dispatchEvent(new Event("change",{bubbles:true}));})()');
  await sleep(900);
  ok('切回命中行后自动重算 ' + EXPECT.hitRows,
    (await cdp.eval('document.getElementById("n-rows").textContent')) === String(EXPECT.hitRows),
    await cdp.eval('document.getElementById("n-rows").textContent'));

  console.log('\n— 导出：另存为直接写盘 —');
  await installFakePicker();
  await cdp.eval('document.getElementById("export").click()');
  let saved = null;
  for (let i = 0; i < 40; i++) {
    await sleep(300);
    saved = await cdp.eval('window.__saved ? window.__saved.name : ""');
    if (saved) break;
  }
  ok('走另存为写盘而不是下载', !!saved, saved);
  const pickerOpt = await cdp.eval('JSON.stringify(window.__pickerOpt)');
  ok('另存为对话框给了 .xlsx 文件名', /\.xlsx/.test(pickerOpt), pickerOpt);
  ok('另存为对话框声明了 xlsx 类型', /spreadsheetml/.test(pickerOpt), pickerOpt);

  const b64 = await cdp.eval('window.__saved ? window.__saved.b64 : ""');
  ok('写盘的字节不为空', !!b64 && b64.length > 5000, b64 ? b64.length : 0);
  if (b64) {
    const buf = Buffer.from(b64, 'base64');
    const back = XLSX.read(buf, { type: 'buffer', cellDates: false, cellNF: true });
    const ws = back.Sheets[back.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    ok('写盘的导出行数 ' + (EXPECT.hitRows + 1), rows.length === EXPECT.hitRows + 1, rows.length);
    ok('写盘的导出保留料号', JSON.stringify(rows).includes('A-001'));
    /* 找到日期列（表头为「入库日期」），确认单元格带日期格式且序列号是整数天 */
    const headRow = rows[0];
    const dateCol = headRow.findIndex((h) => String(h) === '入库日期');
    ok('写盘的导出有入库日期列', dateCol >= 0, JSON.stringify(headRow));
    if (dateCol >= 0) {
      const addr = XLSX.utils.encode_cell({ r: 1, c: dateCol });
      const cell = ws[addr];
      ok('写盘的日期单元格带日期格式', !!(cell && /y/.test(String(cell.z || ''))), cell && cell.z);
      ok('写盘的日期序列号是整天数', !!(cell && Math.abs(cell.v - Math.round(cell.v)) < 1e-9), cell && cell.v);
    }
  }
  ok('导出成功的提示提到保存',
    /保存/.test(await cdp.eval('document.getElementById("toast").innerHTML')),
    await cdp.eval('document.getElementById("toast").innerHTML'));

  console.log('\n— 导出：取消另存为 —');
  await cdp.eval(`(function(){
    window.showSaveFilePicker = function(){
      var e = new Error('aborted'); e.name = 'AbortError'; return Promise.reject(e);
    };
  })()`);
  await cdp.eval('document.getElementById("toast").innerHTML = ""');
  await cdp.eval('document.getElementById("export").click()');
  await sleep(2500);
  ok('取消另存为时不落盘也不下载',
    /取消/.test(await cdp.eval('document.getElementById("toast").innerHTML')),
    await cdp.eval('document.getElementById("toast").innerHTML'));

  console.log('\n— 导出：浏览器不支持时退回下载 —');
  await cdp.eval('(function(){ window.showSaveFilePicker = undefined; })()');
  const before = new Set(fs.readdirSync(outDir));
  await cdp.eval('document.getElementById("export").click()');
  let downloaded = null;
  for (let i = 0; i < 40; i++) {
    await sleep(300);
    const now = fs.readdirSync(outDir).filter(f => !before.has(f) && /\.xlsx$/i.test(f));
    if (now.length) { downloaded = now[0]; break; }
  }
  ok('不支持另存为时退回下载', !!downloaded, downloaded);
  if (downloaded) {
    const p = path.join(outDir, downloaded);
    const size = fs.statSync(p).size;
    ok('下载文件非空', size > 5000, size);
    console.log('    ' + downloaded + '  ' + size + ' B');
    const back = XLSX.read(fs.readFileSync(p), { type: 'buffer', cellDates: false, cellNF: true });
    ok('导出行数 ' + (EXPECT.hitRows + 1),
      XLSX.utils.sheet_to_json(back.Sheets[back.SheetNames[0]], { header: 1 }).length === EXPECT.hitRows + 1,
      XLSX.utils.sheet_to_json(back.Sheets[back.SheetNames[0]], { header: 1 }).length);
    ok('导出保留了料号', JSON.stringify(XLSX.utils.sheet_to_json(back.Sheets[back.SheetNames[0]], { header: 1 })).includes('A-001'));
  }
  await cdp.eval('(function(){ window.showSaveFilePicker = window.__origPicker; })()');

  console.log('\n— 边界：0 命中 —');
  await cdp.eval(`(function(){
    const f = new File([new Blob(['AAA111\\nBBB222\\n'],{type:'text/plain'})], '无关清单.txt', {type:'text/plain'});
    const dt = new DataTransfer(); dt.items.add(f);
    const el = document.getElementById('input-list');
    el.files = dt.files;
    el.dispatchEvent(new Event('change', {bubbles:true}));
  })()`);
  await sleep(800);
  ok('换清单后结果被清空', await cdp.eval('document.getElementById("result").hidden'));
  await cdp.eval('document.getElementById("run").click()');
  await sleep(1200);
  const notice = await cdp.eval('document.getElementById("notice").innerText');
  ok('0 命中给出可操作提示', /建议|换|勾/.test(notice), notice.slice(0, 100));
  ok('0 命中时命中数为 0', (await cdp.eval('document.getElementById("n-rows").textContent')) === '0');
  ok('0 命中时导出按钮禁用', await cdp.eval('document.getElementById("export").disabled'));
  ok('0 命中时导出文案已改',
    /没有可导出/.test(await cdp.eval('document.getElementById("export-label").textContent')),
    await cdp.eval('document.getElementById("export-label").textContent'));
  ok('0 命中时提示区不是全部命中',
    !/全部命中/.test(await cdp.eval('document.getElementById("notice").innerText')));
  await cdp.shot('05-零命中.png');

  console.log('\n— 换 sheet —');
  await cdp.eval(`(function(){
    const f = new File([new Blob([String.fromCharCode(65,45,48,48,49,10)],{type:'text/plain'})], '小清单.txt', {type:'text/plain'});
    const dt = new DataTransfer(); dt.items.add(f);
    const el = document.getElementById('input-list');
    el.files = dt.files;
    el.dispatchEvent(new Event('change', {bubbles:true}));
    const s = document.getElementById('sheet-sel');
    s.value = '其他表';
    s.dispatchEvent(new Event('change', {bubbles:true}));
  })()`);
  await sleep(900);
  ok('切到第二个 sheet 后行数变化',
    (await cdp.eval('document.getElementById("table-count").textContent')) === String(EXPECT.otherRows),
    await cdp.eval('document.getElementById("table-count").textContent'));

  console.log('\n— 响应式 —');
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 420, height: 900, deviceScaleFactor: 1, mobile: true });
  await sleep(700);
  await cdp.shot('06-窄屏.png');
  const overflow = await cdp.eval('document.documentElement.scrollWidth - document.documentElement.clientWidth');
  ok('窄屏无横向溢出', overflow <= 1, overflow);
  await cdp.send('Emulation.clearDeviceMetricsOverride');

  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1000, deviceScaleFactor: 1, mobile: false });
  await sleep(600);
  await cdp.shot('07-桌面.png');
  await cdp.send('Emulation.clearDeviceMetricsOverride');

  console.log('\n  ' + pass + ' 通过 / ' + fail + ' 失败');
  cdp.ws.close();
  edge.kill();
  await sleep(800);
  try { fs.rmSync(profile, { recursive: true, force: true }); } catch (e) { /* edge 还占着，无妨 */ }
  process.exit(fail ? 1 : 0);
})().catch((e) => {
  console.error('测试崩溃:', (e && e.stack) || e);
  process.exit(1);
});
