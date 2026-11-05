/* 拍一张「比对完换模式，按钮还在」的截图，只裁设置面板 */
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const EDGE = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
const ROOT = path.resolve(__dirname, '..');
const INDEX = path.join(ROOT, 'index.html');
const PROFILE = path.join(__dirname, '_edgeshot2');
const OUT = path.join(__dirname, 'shots');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

class CDP {
  constructor(ws) { this.ws = ws; this.id = 0; this.pending = new Map(); this.ws.onmessage = (ev) => this._on(JSON.parse(ev.data)); }
  static async connect(url) {
    const ws = new WebSocket(url);
    await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
    return new CDP(ws);
  }
  _on(msg) {
    if (msg.id && this.pending.has(msg.id)) {
      const { resolve, reject } = this.pending.get(msg.id);
      this.pending.delete(msg.id);
      msg.error ? reject(new Error(msg.error.message)) : resolve(msg.result);
    }
  }
  send(method, params) {
    const id = ++this.id;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.ws.send(JSON.stringify({ id, method, params: params || {} }));
    });
  }
  async eval(expr) {
    const r = await this.send('Runtime.evaluate', { expression: expr, awaitPromise: true, returnByValue: true });
    if (r.exceptionDetails) throw new Error(r.exceptionDetails.text);
    return r.result.value;
  }
  async shot(name, clip) {
    const p = { format: 'png', captureBeyondViewport: false };
    if (clip) p.clip = Object.assign({ scale: 1 }, clip);
    const r = await this.send('Page.captureScreenshot', p);
    fs.writeFileSync(path.join(OUT, name), Buffer.from(r.data, 'base64'));
  }
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  const edge = spawn(EDGE, [
    '--headless=new', '--remote-debugging-port=9445', '--user-data-dir=' + PROFILE,
    '--allow-file-access-from-files', '--window-size=1440,1000', '--disable-gpu', 'about:blank'
  ], { stdio: 'ignore' });

  let target = null;
  for (let i = 0; i < 60; i++) {
    await sleep(300);
    try {
      const list = await (await fetch('http://127.0.0.1:9445/json/list')).json();
      target = list.find((t) => t.type === 'page');
      if (target) break;
    } catch (e) {}
  }
  if (!target) { edge.kill(); throw new Error('Edge 没起来'); }
  const cdp = await CDP.connect(target.webSocketDebuggerUrl);
  await cdp.send('Page.enable');
  await cdp.send('Runtime.enable');
  await cdp.send('Page.navigate', { url: 'file:///' + INDEX.replace(/\\/g, '/') });
  await sleep(1500);

  const txt = fs.readFileSync(path.join(ROOT, 'samples', '物料编码清单.txt'));
  const xls = fs.readFileSync(path.join(ROOT, 'samples', '入库明细表.xlsx'));
  await cdp.eval(`(function(){
    const dt = new DataTransfer();
    dt.items.add(new File([new Uint8Array(${JSON.stringify(Array.from(txt))})], '物料编码清单.txt', {type:'text/plain'}));
    const el = document.getElementById('input-list');
    el.files = dt.files; el.dispatchEvent(new Event('change', {bubbles:true}));
  })()`);
  await sleep(700);
  await cdp.eval(`(function(){
    const dt = new DataTransfer();
    dt.items.add(new File([new Uint8Array(${JSON.stringify(Array.from(xls))})], '入库明细表.xlsx', {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
    const el = document.getElementById('input-table');
    el.files = dt.files; el.dispatchEvent(new Event('change', {bubbles:true}));
  })()`);
  await sleep(1200);

  await cdp.eval('document.getElementById("run").click()');
  await sleep(2500);

  // 比对完，切到「未命中行」
  await cdp.eval('(function(){const e=document.querySelector(\'input[name="outmode"][value="miss"]\');e.checked=true;e.dispatchEvent(new Event("change",{bubbles:true}));})()');
  await sleep(1800);

  await sleep(900);
  const box = await cdp.eval(`(function(){
    window.scrollTo(0, 0);
    const c = document.querySelector('.console').getBoundingClientRect();
    return { x: Math.round(c.x) - 8, y: Math.round(c.y) - 8, width: Math.round(c.width) + 16, height: Math.round(c.height) + 16 };
  })()`);
  await sleep(300);
  await cdp.shot('console-换模式后.png', box);
  const full = await cdp.eval(`(function(){
    const r = document.getElementById('result').getBoundingClientRect();
    window.scrollTo(0, Math.max(0, Math.round(r.top - 40)));
    return true;
  })()`);
  await sleep(700);
  await cdp.shot('result-改设置后.png');

  const state = await cdp.eval(`(function(){
    const b = document.getElementById('run');
    const r = b.getBoundingClientRect();
    const c = document.querySelector('.console');
    const st = document.querySelector('.console__status');
    const sr = st.getBoundingClientRect(), cr = c.getBoundingClientRect();
    return {
      按钮可见: !b.hidden && getComputedStyle(b).display !== 'none' && r.width > 0,
      按钮文案: document.getElementById('run-label').textContent,
      结果区可见: !document.getElementById('result').hidden,
      行数: document.getElementById('n-rows').textContent,
      行标签: document.getElementById('n-rows-l').textContent,
      控制台高: Math.round(cr.height),
      控制台内容高: c.scrollHeight,
      状态行底: Math.round(sr.bottom),
      面板底: Math.round(cr.bottom),
      溢出行数: Math.round(sr.bottom - (cr.bottom - 16))
    };
  })()`);
  console.log(JSON.stringify(state, null, 1));
  console.log('截图：' + path.join(OUT, 'console-换模式后.png'));

  edge.kill();
  await sleep(900);
  try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {}
  process.exit(0);
})().catch((e) => { console.error(e); process.exit(1); });
