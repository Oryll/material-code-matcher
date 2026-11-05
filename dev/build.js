/* 打包：把模板 + 字体 + 库 + 逻辑 拼成单文件 index.html */
'use strict';
const fs = require('fs');
const path = require('path');

const root = __dirname;
const out = path.resolve(root, '..', 'index.html');
const read = (p) => fs.readFileSync(path.join(root, p));
const readText = (p) => fs.readFileSync(path.join(root, p), 'utf8');

const fontFace = (family, weight, file) => {
  const b64 = read(file).toString('base64');
  return `@font-face{font-family:"${family}";font-style:normal;font-weight:${weight};` +
    `font-display:swap;src:url(data:font/woff2;base64,${b64}) format("woff2")}`;
};

const fonts = [
  fontFace('Space Grotesk', 400, 'sg400.woff2'),
  fontFace('Space Grotesk', 500, 'sg500.woff2'),
  fontFace('Space Grotesk', 700, 'sg700.woff2')
].join('\n');

const safe = (js) => js.replace(/<\/script/gi, '<\\/script');

const libs = [
  '<script>' + safe(readText('xlsx.full.min.js')) + '</script>',
  '<script>' + safe(readText('exceljs.min.js')) + '</script>'
].join('\n');

const app = [
  '<script>',
  safe(readText('core.js')),
  safe(readText('app.js')),
  '</script>'
].join('\n');

let html = readText('template.html');
const marks = ['/*@FONTS@*/', '<!--@LIBS@-->', '<!--@APP@-->'];
for (const m of marks) {
  if (html.indexOf(m) === -1) {
    console.error('缺少占位符：' + m);
    process.exit(1);
  }
}
html = html.replace('/*@FONTS@*/', () => fonts)
  .replace('<!--@LIBS@-->', () => libs)
  .replace('<!--@APP@-->', () => app);

fs.writeFileSync(out, html, 'utf8');
const kb = (fs.statSync(out).size / 1024).toFixed(0);
console.log('已生成 ' + out + '  ' + kb + ' KB');
