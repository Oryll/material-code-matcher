/* 验证 GitHub Pages 上的线上版本能真正跑起来：加载、内联资源、比对、导出 */
'use strict';
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');

const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const PORT = 9361;
const PROFILE = path.join(os.tmpdir(), 'xtest', 'liveprofile');
const URL_ = process.argv[2] || 'https://oryll.github.io/material-code-matcher/';
const samples = path.join(__dirname, '..', 'samples');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let pass = 0, fail = 0;
const ok = (n, c, extra) => { if (c) { pass++; console.log('  \u2713 ' + n); } else { fail++; console.log('  \u2717 ' + n + (extra !== undefined ? '  实得: ' + JSON.stringify(extra) : '')); } };

(async () => {
  try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  fs.mkdirSync(PROFILE, { recursive: true });
  const edge = spawn(EDGE, ['--headless=new', '--remote-debugging-port=' + PORT, '--user-data-dir=' + PROFILE,
    '--window-size=1440,1080', '--no-first-run', 'about:blank'], { stdio: 'ignore' });
  try {
    let t = null;
    for (let i = 0; i < 60; i++) {
      await sleep(400);
      try { const l = await (await fetch('http://127.0.0.1:' + PORT + '/json/list')).json(); t = l.find((x) => x.type === 'page'); if (t) break; } catch (e) {}
    }
    if (!t) throw new Error('Edge 未就绪');
    const ws = new WebSocket(t.webSocketDebuggerUrl);
    await new Promise((r) => { ws.onopen = r; });
    let id = 0; const wait = new Map();
    const events = [];
    ws.onmessage = (ev) => {
      const m = JSON.parse(ev.data);
      if (m.id && wait.has(m.id)) { wait.get(m.id)(m); wait.delete(m.id); }
      else if (m.method === 'Network.responseReceived') events.push(m.params.response);
    };
    const send = (method, params) => { const n = ++id; ws.send(JSON.stringify({ id: n, method, params: params || {} })); return new Promise((r) => wait.set(n, r)); };
    const evalx = async (expr) => {
      const m = await send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true, userGesture: true });
      if (m.result && m.result.exceptionDetails) throw new Error(JSON.stringify(m.result.exceptionDetails));
      return m.result && m.result.result ? m.result.result.value : undefined;
    };

    await send('Page.enable'); await send('Runtime.enable'); await send('Network.enable');

    console.log('\n— 加载线上页面 —');
    const t0 = Date.now();
    await send('Page.navigate', { url: URL_ });
    await sleep(1500);
    let ready = false;
    for (let i = 0; i < 100; i++) {
      const s = await evalx('document.readyState');
      const hasCore = await evalx('typeof window.CORE');
      if (s === 'complete' && hasCore === 'object') { ready = true; break; }
      await sleep(500);
    }
    ok('页面加载完成且 CORE 就绪', ready, await evalx('document.readyState'));
    console.log('    耗时 ' + ((Date.now() - t0) / 1000).toFixed(1) + ' 秒');

    ok('标题正确', /料号比对/.test(await evalx('document.title')), await evalx('document.title'));
    ok('SheetJS 内联可用', await evalx('typeof window.XLSX.read') === 'function');
    ok('ExcelJS 内联可用', await evalx('typeof window.ExcelJS.Workbook') === 'function');
    ok('字体已加载', await evalx('document.fonts.check("500 16px \'Space Grotesk\'")'));

    /* 页面自己就是一个 http 文档，其余任何 http 请求都算外部依赖 */
    const others = events.filter((r) => /^https?:/.test(r.url) && r.url !== URL_ && r.type !== 'Document');
    ok('除文档本身外无其它 http 请求', others.length === 0, others.map((r) => r.url).slice(0, 5));
    const docReqs = events.filter((r) => r.type === 'Document' || r.url === URL_);
    ok('只请求了这一个文档', docReqs.length <= 1, docReqs.length);

    console.log('\n— 跑一遍比对 —');
    const listB64 = fs.readFileSync(path.join(samples, '物料编码清单.txt')).toString('base64');
    const xlsxB64 = fs.readFileSync(path.join(samples, '入库明细表.xlsx')).toString('base64');
    const res = await evalx(`(async function(){
      function b2a(b64){ var s=atob(b64), a=new Uint8Array(s.length); for(var i=0;i<s.length;i++) a[i]=s.charCodeAt(i); return a; }
      window.showSaveFilePicker = function(){
        var chunks = [];
        return Promise.resolve({
          name: 't.xlsx',
          createWritable: function(){ return Promise.resolve({
            write: function(d){ chunks.push(d); return Promise.resolve(); },
            close: function(){ var n=0; chunks.forEach(function(c){n+=c.length;});
              var all=new Uint8Array(n), o=0; chunks.forEach(function(c){all.set(c,o);o+=c.length;});
              window.__saved={size:n,b64:btoa(Array.from(all.slice(0,4)).map(function(b){return String.fromCharCode(b);}).join(''))}; return Promise.resolve(); }
          }); }
        });
      };
      var d1=new DataTransfer();
      d1.items.add(new File([b2a(${JSON.stringify(listB64)})],'物料编码清单.txt',{type:'text/plain'}));
      var e1=document.getElementById('input-list'); e1.files=d1.files; e1.dispatchEvent(new Event('change',{bubbles:true}));
      await new Promise(function(r){setTimeout(r,600);});
      var d2=new DataTransfer();
      d2.items.add(new File([b2a(${JSON.stringify(xlsxB64)})],'入库明细表.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      var e2=document.getElementById('input-table'); e2.files=d2.files; e2.dispatchEvent(new Event('change',{bubbles:true}));
      await new Promise(function(r){setTimeout(r,1200);});
      document.getElementById('run').click();
      await new Promise(function(r){setTimeout(r,1500);});
      return {
        rows: document.getElementById('n-rows').textContent,
        codes: document.getElementById('n-codes').textContent,
        unused: document.getElementById('n-unused').textContent,
        scan: document.getElementById('n-scan').textContent,
        pct: document.getElementById('rpct').textContent
      };
    })()`);
    ok('命中行 338', res && res.rows === '338', res && res.rows);
    ok('命中料号 130', res && res.codes === '130', res && res.codes);
    ok('清单里没出现的 17', res && res.unused === '17', res && res.unused);
    ok('扫描行数 600', res && res.scan === '600', res && res.scan);
    ok('命中率 88%', res && /88%/.test(res.pct), res && res.pct);

    console.log('\n— 导出 —');
    const saved = await evalx(`(async function(){
      document.getElementById('export').click();
      await new Promise(function(r){setTimeout(r,4000);});
      return window.__saved || null;
    })()`);
    ok('导出写盘成功', saved && saved.size > 5000, saved && saved.size);
    ok('导出开头是 PK', saved && /^UEsDBA/.test(saved.b64), saved && saved.b64);
    ok('导出提示提到保存', /保存/.test(await evalx('document.getElementById("toast").innerHTML')));

    const errs = await evalx('window.__errs ? window.__errs.length : 0');
    ok('无未捕获脚本错误', !errs, errs);

  } finally {
    edge.kill();
    await sleep(900);
    try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  }
  console.log('\n结果: ' + pass + ' 通过 / ' + fail + ' 失败');
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('脚本错误: ' + e.message); process.exit(2); });
