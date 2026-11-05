/* 对最终产物 index.html 做一次冒烟：控制台无错误、关键节点存在、字体与库都内联 */
'use strict';
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const BUILD = __dirname;
const ROOT = path.join(BUILD, '..');
const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const PORT = 9341;
const PROFILE = path.join(BUILD, '_smokeprofile');
const FILE = path.join(ROOT, 'index.html');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

class CDP {
  constructor(ws) { this.ws = ws; this.id = 0; this.waiting = new Map(); }
  static connect(url) {
    return new Promise((resolve, reject) => {
      const ws = new WebSocket(url);
      ws.onopen = () => resolve(new CDP(ws));
      ws.onerror = (e) => reject(new Error('ws error'));
    }).then((c) => c);
  }
  static async open(wsUrl) {
    const c = await CDP.connect(wsUrl);
    c.ws.onmessage = (ev) => {
      const msg = JSON.parse(ev.data);
      if (msg.id && c.waiting.has(msg.id)) {
        const { resolve, reject } = c.waiting.get(msg.id);
        c.waiting.delete(msg.id);
        if (msg.error) reject(new Error(JSON.stringify(msg.error)));
        else resolve(msg.result);
      } else if (msg.method === 'Runtime.exceptionThrown') {
        c.errors.push(msg.params.exceptionDetails);
      } else if (msg.method === 'Runtime.consoleAPICalled' && msg.params.type === 'error') {
        c.errors.push(msg.params.args);
      }
    };
    c.errors = [];
    return c;
  }
  send(method, params) {
    const id = ++this.id;
    this.ws.send(JSON.stringify({ id, method, params: params || {} }));
    return new Promise((resolve, reject) => this.waiting.set(id, { resolve, reject }));
  }
  async eval(expr, awaitPromise) {
    const r = await this.send('Runtime.evaluate', {
      expression: expr, returnByValue: true, awaitPromise: !!awaitPromise, allowUnsafeEvalBlockedByCSP: true
    });
    if (r.exceptionDetails) throw new Error('eval failed: ' + JSON.stringify(r.exceptionDetails).slice(0, 300));
    return r.result.value;
  }
}

(async () => {
  try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  const edge = spawn(EDGE, [
    '--headless=new', '--remote-debugging-port=' + PORT, '--user-data-dir=' + PROFILE,
    '--allow-file-access-from-files', '--window-size=1440,1000', '--no-first-run',
    '--disable-extensions', '--hide-scrollbars', 'about:blank'
  ], { stdio: 'ignore' });

  const ok = [], bad = [];
  const check = (name, cond, extra) => {
    if (cond) ok.push(name); else bad.push(name + (extra === undefined ? '' : ' 实得 ' + JSON.stringify(extra)));
  };

  try {
    let target = null;
    for (let i = 0; i < 60; i++) {
      await sleep(400);
      try {
        const list = await (await fetch('http://127.0.0.1:' + PORT + '/json/list')).json();
        target = list.find((t) => t.type === 'page');
        if (target) break;
      } catch (e) {}
    }
    if (!target) throw new Error('找不到调试目标');
    const cdp = await CDP.open(target.webSocketDebuggerUrl);
    await cdp.send('Page.enable');
    await cdp.send('Runtime.enable');
    await cdp.send('Page.navigate', { url: 'file:///' + FILE.replace(/\\/g, '/') });
    await sleep(2500);

    const info = await cdp.eval(`(function(){
      var out = {};
      out.title = document.title;
      out.hasCORE = typeof window.CORE === 'object' && !!(window.CORE && window.CORE.runMatch);
      out.hasXLSX = typeof window.XLSX === 'object';
      out.hasExcelJS = typeof window.ExcelJS === 'object';
      out.ids = ['zone-list','zone-table','drop-list','drop-table','input-list','input-table',
                 'opt-header','opt-case','opt-zero','opt-space','run','status','result',
                 'n-rows','n-codes','n-unused','n-scan','export','export-label','rfill','rhead','rpct',
                 'notice','tablewrap','tbl','tbl-body','toast','sheet-sel','bench'].filter(function(id){
        return !document.getElementById(id);
      });
      out.fonts = performance.getEntriesByType('resource').filter(function(r){ return /^https?:/.test(r.name); }).length;
      var loaded = 0, total = 0;
      document.fonts.forEach(function(f){ total++; if (/Space Grotesk/.test(f.family) && f.status === 'loaded') loaded++; });
      out.fontFaceCount = loaded;
      out.fontTotal = total;
      out.fontCheck = document.fonts.check('400 16px "Space Grotesk"') && document.fonts.check('700 16px "Space Grotesk"');
      out.bodyH = document.body.scrollHeight;
      out.overflowX = document.documentElement.scrollWidth - document.documentElement.clientWidth;
      out.bench = document.getElementById('bench').getBoundingClientRect().width;
      return out;
    })()`);

    check('标题正确', /料号/.test(info.title), info.title);
    check('CORE 已就绪', info.hasCORE === true, info.hasCORE);
    check('SheetJS 已就绪', info.hasXLSX === true, info.hasXLSX);
    check('ExcelJS 已就绪', info.hasExcelJS === true, info.hasExcelJS);
    check('关键 DOM 节点齐全', info.ids.length === 0, info.ids);
    check('无外部网络请求', info.fonts === 0, info.fonts);
    check('字体已加载', info.fontCheck === true && info.fontFaceCount > 0, [info.fontFaceCount, info.fontTotal, info.fontCheck]);
    check('页面已渲染出内容', info.bodyH > 900, info.bodyH);
    check('无横向溢出', info.overflowX <= 0, info.overflowX);
    check('无脚本错误', cdp.errors.length === 0, cdp.errors.length);

    // 跑一次真实比对，确认打包后的页面链路可用
    await cdp.eval(`(async function(){
      var txt = 'A-001\\nA-002\\nZ999\\n';
      var dt1 = new DataTransfer();
      dt1.items.add(new File([new TextEncoder().encode(txt)], 's.txt', {type:'text/plain'}));
      var el = document.getElementById('input-list');
      el.files = dt1.files;
      el.dispatchEvent(new Event('change', {bubbles:true}));
      var wb = XLSX.utils.book_new();
      var ws = XLSX.utils.aoa_to_sheet([['物料编码','品名'],['A-001','螺丝'],['A-003','垫片'],['A-002','螺母']]);
      XLSX.utils.book_append_sheet(wb, ws, 'S1');
      var bytes = XLSX.write(wb, {type:'array', bookType:'xlsx'});
      var dt2 = new DataTransfer();
      dt2.items.add(new File([bytes], 'd.xlsx', {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      var el2 = document.getElementById('input-table');
      el2.files = dt2.files;
      el2.dispatchEvent(new Event('change', {bubbles:true}));
      await new Promise(function(r){ setTimeout(r, 900); });
      document.getElementById('run').click();
      await new Promise(function(r){ setTimeout(r, 700); });
      return true;
    })()`, true);

    const res = await cdp.eval(`(function(){
      return {
        rows: document.getElementById('n-rows').textContent,
        codes: document.getElementById('n-codes').textContent,
        unused: document.getElementById('n-unused').textContent,
        scan: document.getElementById('n-scan').textContent,
        resultVisible: !document.getElementById('result').hidden,
        bodyRows: document.querySelectorAll('#tbl-body tr').length,
        exportDisabled: document.getElementById('export').disabled,
        pct: document.getElementById('rpct').textContent
      };
    })()`);
    check('比对得出 2 行', res.rows === '2', res.rows);
    check('命中 2 个料号', res.codes === '2', res.codes);
    check('1 个料号没出现', res.unused === '1', res.unused);
    check('扫描 3 行', res.scan === '3', res.scan);
    check('结果区已显示', res.resultVisible === true, res.resultVisible);
    check('预览表 2 行', res.bodyRows === 2, res.bodyRows);
    check('导出可用', res.exportDisabled === false, res.exportDisabled);
    check('百分比 67%', /67/.test(res.pct), res.pct);
    check('比对过程无报错', cdp.errors.length === 0, cdp.errors.length);

    console.log(ok.map((n) => '  ✓ ' + n).join('\n'));
    if (bad.length) console.log('\n失败：\n' + bad.map((n) => '  ✗ ' + n).join('\n'));
    console.log('\n' + ok.length + ' 通过 / ' + bad.length + ' 失败');
    process.exitCode = bad.length ? 1 : 0;
  } finally {
    try { edge.kill(); } catch (e) {}
    await sleep(900);
    try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  }
})().catch((e) => { console.error('冒烟测试异常：', e && e.message); process.exitCode = 1; });
