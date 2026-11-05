/* 用生成好的示例文件在真实浏览器里跑一遍，确认页面显示的数字与逻辑层一致 */
'use strict';
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const BUILD = __dirname;
const ROOT = path.join(BUILD, '..');
const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const PORT = 9347;
const PROFILE = path.join(BUILD, '_sampleprofile');
const DOWN = path.join(BUILD, '_sampledl');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const listB64 = fs.readFileSync(path.join(ROOT, 'samples', '物料编码清单.txt')).toString('base64');
const bookB64 = fs.readFileSync(path.join(ROOT, 'samples', '入库明细表.xlsx')).toString('base64');

(async () => {
  fs.rmSync(DOWN, { recursive: true, force: true });
  fs.mkdirSync(DOWN, { recursive: true });
  try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}

  const edge = spawn(EDGE, ['--headless=new', '--remote-debugging-port=' + PORT, '--user-data-dir=' + PROFILE,
    '--allow-file-access-from-files', '--window-size=1440,1080', '--no-first-run', '--hide-scrollbars', 'about:blank'],
    { stdio: 'ignore' });
  try {
    let t = null;
    for (let i = 0; i < 60; i++) {
      await sleep(400);
      try { const l = await (await fetch('http://127.0.0.1:' + PORT + '/json/list')).json(); t = l.find((x) => x.type === 'page'); if (t) break; } catch (e) {}
    }
    const ws = new WebSocket(t.webSocketDebuggerUrl);
    await new Promise((r) => { ws.onopen = r; });
    let id = 0; const wait = new Map(); const errs = [];
    ws.onmessage = (ev) => {
      const m = JSON.parse(ev.data);
      if (m.id && wait.has(m.id)) { wait.get(m.id)(m.result); wait.delete(m.id); }
      else if (m.method === 'Runtime.exceptionThrown' || (m.method === 'Runtime.consoleAPICalled' && m.params.type === 'error')) errs.push(m);
    };
    const send = (method, params) => { const n = ++id; ws.send(JSON.stringify({ id: n, method, params: params || {} })); return new Promise((r) => wait.set(n, r)); };
    const ev = async (expr, awaitPromise) => {
      const r = await send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: !!awaitPromise });
      if (r.exceptionDetails) throw new Error(JSON.stringify(r.exceptionDetails).slice(0, 400));
      return r.result.value;
    };

    await send('Page.enable'); await send('Runtime.enable');
    await send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: DOWN });
    await send('Page.navigate', { url: 'file:///' + path.join(ROOT, 'index.html').replace(/\\/g, '/') });
    await sleep(2400);

    const upload = `(async function(){
      function b2a(b64){ var s=atob(b64), a=new Uint8Array(s.length); for(var i=0;i<s.length;i++) a[i]=s.charCodeAt(i); return a; }
      var d1=new DataTransfer();
      d1.items.add(new File([b2a(${JSON.stringify(listB64)})],'物料编码清单.txt',{type:'text/plain'}));
      var e1=document.getElementById('input-list'); e1.files=d1.files; e1.dispatchEvent(new Event('change',{bubbles:true}));
      await new Promise(function(r){setTimeout(r,600);});
      var d2=new DataTransfer();
      d2.items.add(new File([b2a(${JSON.stringify(bookB64)})],'入库明细表.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      var e2=document.getElementById('input-table'); e2.files=d2.files; e2.dispatchEvent(new Event('change',{bubbles:true}));
      await new Promise(function(r){setTimeout(r,1200);});
      return true;
    })()`;
    await ev(upload, true);
    await sleep(600);

    const pre = await ev(`(function(){
      return {
        listCount: document.getElementById('list-count').textContent,
        listDups: document.getElementById('list-dups').textContent,
        listEnc: document.getElementById('list-enc').textContent,
        tableCount: document.getElementById('table-count').textContent,
        tableCols: document.getElementById('table-cols').textContent,
        tableCodes: document.getElementById('table-codes').textContent,
        headerChecked: document.getElementById('opt-header').checked
      };
    })()`);
    console.log('上传后：' + JSON.stringify(pre, null, 2));

    const snap = (label) => `(async function(){
      document.getElementById('run').click();
      await new Promise(function(r){setTimeout(r,1800);});
      return {
        rows: document.getElementById('n-rows').textContent,
        codes: document.getElementById('n-codes').textContent,
        unused: document.getElementById('n-unused').textContent,
        scan: document.getElementById('n-scan').textContent,
        pct: document.getElementById('rpct').textContent,
        cap: document.getElementById('rcap').textContent,
        notice: document.getElementById('notice').textContent.replace(/\\s+/g,' ').trim().slice(0,60),
        firstDate: (document.querySelector('#tbl-body tr td:nth-child(7)')||{}).textContent,
        exportDisabled: document.getElementById('export').disabled
      };
    })()`;

    const def = await ev(snap(), true);
    console.log('\n【默认】' + JSON.stringify(def, null, 2));

    // 全开
    await ev(`(function(){
      ['opt-zero','opt-space'].forEach(function(id){ document.getElementById(id).checked = true; });
      document.getElementById('opt-case').checked = true;
      document.getElementById('opt-case').dispatchEvent(new Event('change',{bubbles:true}));
      return true;
    })()`);
    await sleep(700);
    const all = await ev(snap(), true);
    console.log('\n【全开】' + JSON.stringify(all, null, 2));

    // 全关
    await ev(`(function(){
      ['opt-case','opt-zero','opt-space'].forEach(function(id){ document.getElementById(id).checked = false; });
      document.getElementById('opt-case').dispatchEvent(new Event('change',{bubbles:true}));
      return true;
    })()`);
    await sleep(700);
    const none = await ev(snap(), true);
    console.log('\n【全关】' + JSON.stringify(none, null, 2));

    // 导出（回到默认）
    await ev(`(function(){
      document.getElementById('opt-case').checked = true;
      ['opt-zero','opt-space'].forEach(function(id){ document.getElementById(id).checked = false; });
      document.getElementById('opt-case').dispatchEvent(new Event('change',{bubbles:true}));
      return true;
    })()`);
    await sleep(600);
    await ev(`(async function(){
      document.getElementById('run').click();
      await new Promise(function(r){setTimeout(r,1500);});
      document.getElementById('export').click();
      await new Promise(function(r){setTimeout(r,3500);});
      return true;
    })()`, true);
    await sleep(1500);
    const files = fs.readdirSync(DOWN);
    console.log('\n下载目录：' + JSON.stringify(files));
    if (files.length) {
      const f = files[0];
      const back = require(path.join(BUILD, 'xlsx.full.min.js')).read(fs.readFileSync(path.join(DOWN, f)),
        { type: 'buffer', cellDates: false, cellNF: true });
      const a1 = require(path.join(BUILD, 'xlsx.full.min.js')).utils.sheet_to_json(back.Sheets['比对结果'], { header: 1 });
      console.log('  工作表：' + JSON.stringify(back.SheetNames));
      console.log('  表头：' + JSON.stringify(a1[0]));
      console.log('  数据行：' + (a1.length - 1));
      const ws2 = back.Sheets['比对结果'];
      console.log('  日期 G2 v=' + ws2.G2.v + ' z=' + ws2.G2.z);
    }
    console.log('\n脚本错误数：' + errs.length);
  } finally {
    try { edge.kill(); } catch (e) {}
    await sleep(900);
    try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  }
})();
