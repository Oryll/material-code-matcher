/* 给示例文件出一张结果态截图，确认视觉与数字 */
'use strict';
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const BUILD = __dirname;
const ROOT = path.join(BUILD, '..');
const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const PORT = 9348;
const PROFILE = path.join(BUILD, '_sp2');
const OUT = path.join(BUILD, 'shots');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const listB64 = fs.readFileSync(path.join(ROOT, 'samples', '物料编码清单.txt')).toString('base64');
const bookB64 = fs.readFileSync(path.join(ROOT, 'samples', '入库明细表.xlsx')).toString('base64');

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
    await sleep(2400);

    await send('Runtime.evaluate', {
      expression: `(async function(){
        function b2a(b64){ var s=atob(b64), a=new Uint8Array(s.length); for(var i=0;i<s.length;i++) a[i]=s.charCodeAt(i); return a; }
        var d1=new DataTransfer();
        d1.items.add(new File([b2a(${JSON.stringify(listB64)})],'物料编码清单.txt',{type:'text/plain'}));
        var e1=document.getElementById('input-list'); e1.files=d1.files; e1.dispatchEvent(new Event('change',{bubbles:true}));
        await new Promise(function(r){setTimeout(r,700);});
        var d2=new DataTransfer();
        d2.items.add(new File([b2a(${JSON.stringify(bookB64)})],'入库明细表.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
        var e2=document.getElementById('input-table'); e2.files=d2.files; e2.dispatchEvent(new Event('change',{bubbles:true}));
        await new Promise(function(r){setTimeout(r,1400);});
        document.getElementById('run').click();
        await new Promise(function(r){setTimeout(r,1800);});
        window.scrollTo(0,0);
        return true;
      })()`, returnByValue: true, awaitPromise: true
    });
    await sleep(1200);
    const shoot = async (name) => {
      const m = await send('Page.getLayoutMetrics');
      const w = Math.ceil(m.cssContentSize.width), h = Math.ceil(m.cssContentSize.height);
      const shot = await send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true, clip: { x: 0, y: 0, width: w, height: h, scale: 1 } });
      fs.writeFileSync(path.join(OUT, name), Buffer.from(shot.data, 'base64'));
      console.log('saved ' + name + ' ' + w + 'x' + h);
    };
    await shoot('sample-结果.png');

    // 未命中行模式：换到 miss 后重新比对，出一张图给用户看差异
    await send('Runtime.evaluate', {
      expression: `(async function(){
        var e=document.querySelector('input[name="outmode"][value="miss"]');
        e.checked=true; e.dispatchEvent(new Event('change',{bubbles:true}));
        await new Promise(function(r){setTimeout(r,400);});
        document.getElementById('run').click();
        await new Promise(function(r){setTimeout(r,2200);});
        window.scrollTo(0,0);
        return document.getElementById('n-rows').textContent;
      })()`, returnByValue: true, awaitPromise: true
    });
    await sleep(900);
    await shoot('sample-未命中行.png');
  } finally {
    try { edge.kill(); } catch (e) {}
    await sleep(900);
    try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  }
})();
