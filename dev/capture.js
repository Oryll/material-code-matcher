/* 给最终产物拍一张结果态整页图 */
'use strict';
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const BUILD = __dirname;
const ROOT = path.join(BUILD, '..');
const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const PORT = 9343;
const PROFILE = path.join(BUILD, '_shotprofile');
const OUT = path.join(BUILD, 'shots');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  const edge = spawn(EDGE, ['--headless=new', '--remote-debugging-port=' + PORT, '--user-data-dir=' + PROFILE,
    '--allow-file-access-from-files', '--window-size=1440,1080', '--no-first-run', '--hide-scrollbars', 'about:blank'], { stdio: 'ignore' });
  try {
    let t = null;
    for (let i = 0; i < 60; i++) {
      await sleep(400);
      try { const l = await (await fetch('http://127.0.0.1:' + PORT + '/json/list')).json(); t = l.find((x) => x.type === 'page'); if (t) break; } catch (e) {}
    }
    const ws = new WebSocket(t.webSocketDebuggerUrl);
    await new Promise((r) => { ws.onopen = r; });
    let id = 0; const wait = new Map();
    ws.onmessage = (ev) => { const m = JSON.parse(ev.data); if (m.id && wait.has(m.id)) { wait.get(m.id)(m.result); wait.delete(m.id); } };
    const send = (method, params) => { const n = ++id; ws.send(JSON.stringify({ id: n, method, params: params || {} })); return new Promise((r) => wait.set(n, r)); };
    await send('Page.enable'); await send('Runtime.enable');
    await send('Page.navigate', { url: 'file:///' + path.join(ROOT, 'index.html').replace(/\\/g, '/') });
    await sleep(2200);

    const listB64 = fs.readFileSync(path.join(BUILD, 'testdata', '清单.txt')).toString('base64');
    const bookB64 = fs.readFileSync(path.join(BUILD, 'testdata', '入库明细.xlsx')).toString('base64');
    await send('Runtime.evaluate', {
      expression: `(async function(){
        function b2a(b64){ var s = atob(b64), a = new Uint8Array(s.length); for (var i=0;i<s.length;i++) a[i]=s.charCodeAt(i); return a; }
        var d1 = new DataTransfer();
        d1.items.add(new File([b2a('${listB64}')], '清单.txt', {type:'text/plain'}));
        var e1 = document.getElementById('input-list'); e1.files = d1.files; e1.dispatchEvent(new Event('change', {bubbles:true}));
        await new Promise(function(r){ setTimeout(r, 500); });
        var d2 = new DataTransfer();
        d2.items.add(new File([b2a('${bookB64}')], '入库明细.xlsx', {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
        var e2 = document.getElementById('input-table'); e2.files = d2.files; e2.dispatchEvent(new Event('change', {bubbles:true}));
        await new Promise(function(r){ setTimeout(r, 700); });
        document.getElementById('run').click();
        await new Promise(function(r){ setTimeout(r, 1600); });
        window.scrollTo(0, 0);
        return document.getElementById('n-rows').textContent;
      })()`, returnByValue: true, awaitPromise: true
    });
    await sleep(1200);
    const metrics = await send('Page.getLayoutMetrics');
    const h = Math.ceil(metrics.cssContentSize.height);
    const w = Math.ceil(metrics.cssContentSize.width);
    const shot = await send('Page.captureScreenshot', {
      format: 'png', captureBeyondViewport: true,
      clip: { x: 0, y: 0, width: w, height: h, scale: 1 }
    });
    fs.writeFileSync(path.join(OUT, 'final-结果.png'), Buffer.from(shot.data, 'base64'));
    console.log('saved final-结果.png ' + w + 'x' + h);
  } finally {
    try { edge.kill(); } catch (e) {}
    await sleep(900);
    try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  }
})();
