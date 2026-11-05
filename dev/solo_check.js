/* 把一个 index.html 单独拷到空目录，在隔离环境里跑完整流程，证明它不依赖任何同目录文件。
   用法: node solo_check.js <index.html 路径> */
'use strict';
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');

const src = process.argv[2] || path.join(__dirname, '..', 'index.html');
/* 纯 ASCII 隔离目录：中文路径过 PowerShell/浏览器命令行会出问题 */
const soloDir = path.join(os.tmpdir(), 'xtest', 'solo_' + Date.now());
fs.mkdirSync(soloDir, { recursive: true });
fs.copyFileSync(src, path.join(soloDir, 'index.html'));
const files = fs.readdirSync(soloDir);
console.log('隔离目录:', soloDir);
console.log('目录内容:', JSON.stringify(files), '(只应有 index.html)');

const appUrl = 'file:///' + path.join(soloDir, 'index.html').replace(/\\/g, '/');
const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const PORT = 9344;
const profile = path.join(soloDir, '_prof');

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
      setTimeout(() => { if (this.pending.has(id)) { this.pending.delete(id); rej(new Error('CDP 超时: ' + method)); } }, 30000);
    });
  }
  async eval(expr, awaitPromise) {
    const r = await this.send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: !!awaitPromise, userGesture: true });
    if (r.exceptionDetails) throw new Error('页面异常: ' + ((r.exceptionDetails.exception && r.exceptionDetails.exception.description) || r.exceptionDetails.text));
    return r.result && r.result.value;
  }
}

let pass = 0, fail = 0;
const ok = (name, cond, extra) => {
  if (cond) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + '  实得 ' + JSON.stringify(extra)); }
};

const samples = path.join(__dirname, '..', 'samples');
const listB64 = fs.readFileSync(path.join(samples, '物料编码清单.txt')).toString('base64');
const xlsxB64 = fs.readFileSync(path.join(samples, '入库明细表.xlsx')).toString('base64');

(async () => {
  const edge = spawn(EDGE, ['--headless=new', '--remote-debugging-port=' + PORT,
    '--user-data-dir=' + profile, '--no-first-run', '--no-default-browser-check',
    '--disable-gpu', '--hide-scrollbars', '--window-size=1440,1000',
    '--force-device-scale-factor=1', 'about:blank'], { stdio: 'ignore' });

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
  const netCalls = [];
  cdp.onEvent = (msg) => {
    if (msg.method === 'Network.requestWillBeSent') netCalls.push(msg.params.request.url);
    if (msg.method === 'Runtime.exceptionThrown') console.log('  [页面异常] ' + msg.params.exceptionDetails.text);
    if (msg.method === 'Runtime.consoleAPICalled' && msg.params.type === 'error') {
      console.log('  [console.error] ' + msg.params.args.map(a => a.value || a.description).join(' '));
    }
  };
  await cdp.send('Network.enable');

  await cdp.send('Page.navigate', { url: appUrl });
  await sleep(3000);

  console.log('\n— 单文件自包含 —');
  ok('页面标题存在', !!(await cdp.eval('document.title')));
  ok('SheetJS 已内嵌', await cdp.eval('typeof XLSX === "object" && typeof XLSX.read === "function"'));
  ok('ExcelJS 已内嵌', await cdp.eval('typeof ExcelJS === "object" && typeof ExcelJS.Workbook === "function"'));
  ok('核心逻辑已内嵌', await cdp.eval('typeof CORE === "object" && typeof CORE.runMatch === "function"'));
  ok('字体已内嵌', (await cdp.eval('document.fonts.size')) > 0, await cdp.eval('document.fonts.size'));
  const fontsLoaded = await cdp.eval(`(function(){
    var ok = [];
    document.fonts.forEach(function(f){ if (f.family.indexOf('Space Grotesk') >= 0) ok.push(f.status); });
    return ok.join(',');
  })()`);
  ok('Space Grotesk 可用', /loaded/.test(fontsLoaded || ''), fontsLoaded);
  const ext = await cdp.eval('performance.getEntriesByType("resource").filter(function(r){return /^https?:/.test(r.name);}).map(function(r){return r.name;}).join("|")');
  ok('零外部网络请求', !ext, ext);
  const httpReqs = netCalls.filter(u => /^https?:/.test(u));
  ok('网络面板也没有外部请求', httpReqs.length === 0, httpReqs);
  ok('只有 1 个文档请求', netCalls.filter(u => u.indexOf('file://') === 0).length === 1, netCalls.length);

  console.log('\n— 完整流程（上传 → 比对 → 导出）—');
  const inject = (id, b64, name, mime) => cdp.eval(`(function(){
    var bin = atob(${JSON.stringify(b64)}), bytes = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    var f = new File([bytes], ${JSON.stringify(name)}, { type: ${JSON.stringify(mime)} });
    var dt = new DataTransfer(); dt.items.add(f);
    var el = document.getElementById(${JSON.stringify(id)});
    el.files = dt.files;
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  })()`);
  await inject('input-list', listB64, '物料编码清单.txt', 'text/plain');
  await sleep(600);
  await inject('input-table', xlsxB64, '入库明细表.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await sleep(1200);
  ok('清单已读入', await cdp.eval('document.getElementById("list-count").textContent') !== '—');
  ok('明细表已读入', await cdp.eval('document.getElementById("table-count").textContent') !== '—');

  await cdp.eval('document.getElementById("run").click()');
  await sleep(2500);
  const rows = await cdp.eval('document.getElementById("n-rows").textContent');
  ok('比对出 338 行（示例文件的正确答案）', rows === '338', rows);
  const pct = await cdp.eval('document.getElementById("rpct").textContent');
  ok('命中率 88%', /88/.test(pct || ''), pct);

  /* 用假句柄接管另存为，把页面写出的字节扣下来 */
  await cdp.eval(`(function(){
    window.__saved = null;
    window.showSaveFilePicker = function(opt){
      return Promise.resolve({ name: (opt && opt.suggestedName) || 'x.xlsx', createWritable: function(){
        var chunks = [];
        return Promise.resolve({
          write: function(d){ chunks.push(d); return Promise.resolve(); },
          close: function(){
            var all = chunks[0], bin = '';
            for (var i = 0; i < all.length; i++) bin += String.fromCharCode(all[i]);
            window.__saved = { name: opt.suggestedName, size: all.length, b64: btoa(bin) };
            return Promise.resolve();
          }
        });
      }});
    };
    return true;
  })()`);
  await cdp.eval('document.getElementById("export").click()');
  await sleep(3500);
  const saved = await cdp.eval('window.__saved ? {name: window.__saved.name, size: window.__saved.size, b64: window.__saved.b64.slice(0,0)} : null');
  ok('导出成功并写盘', !!(saved && saved.size > 1000), saved);
  if (saved && saved.size) {
    const full = await cdp.eval('window.__saved.b64');
    const buf = Buffer.from(full, 'base64');
    const outPath = path.join(soloDir, '导出校验.xlsx');
    fs.writeFileSync(outPath, buf);
    const head = buf.slice(0, 2).toString('latin1');
    ok('导出的是合法 zip（PK 开头）', head === 'PK', head);
    console.log('  导出字节数: ' + buf.length + '  文件名: ' + saved.name);
    ok('文件名带日期戳', /料号比对结果_\d{8}_\d{4}\.xlsx/.test(saved.name || ''), saved.name);
  }

  console.log('\n结果: ' + pass + ' 通过 / ' + fail + ' 失败');
  edge.kill();
  await sleep(900);
  try { fs.rmSync(profile, { recursive: true, force: true }); } catch (e) {}
  console.log('隔离目录保留供检查: ' + soloDir);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
