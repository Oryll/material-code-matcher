/* 料号比对 — 核心逻辑（不依赖 DOM，浏览器与 Node 均可运行） */
(function (root) {
  'use strict';

  var SEP_RE = /[\t,;，；、]+/;
  var ZERO_RE = /^0+(?=[0-9])/;
  var INVIS_RE = /[\u200B-\u200F\uFEFF]/g;
  var TRIM_PAIRS_RE = /^["'\u201C\u201D\u2018\u2019]+|["'\u201C\u201D\u2018\u2019]+$/g;

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  /* 三种输出范围的中文名，界面与导出共用 */
  var OUT_MODE_LABEL = { hit: '只要命中行', all: '全部行并标注', miss: '只要未命中行' };

  function dateToText(d) {
    try {
      return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
    } catch (e) { return String(d); }
  }

  /* ---------- 日期：Excel 序列号 ↔ 文本（不经 Date 时区，避免差一天） ---------- */
  function serialToDate(serial) {
    var days = Math.floor(serial);
    var frac = serial - days;
    // 1900-02-29 这个不存在的闰日让前 60 天偏移一天
    var eff = days < 61 ? days + 1 : days;
    var u = new Date(Date.UTC(1899, 11, 30) + eff * 86400000 + Math.round(frac * 86400000));
    return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate(),
      u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds());
  }

  function dateToSerial(d) {
    var ms = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(),
      d.getHours(), d.getMinutes(), d.getSeconds()) - Date.UTC(1899, 11, 30);
    var serial = ms / 86400000;
    if (serial < 61) serial -= 1;
    return serial;
  }

  /* 序列号 → 给人看的日期文本 */
  function dateText(serial) {
    if (typeof serial !== 'number' || !isFinite(serial) || serial <= 0) return cellText(serial);
    var d = serialToDate(serial);
    var s = dateToText(d);
    if (Math.abs(serial - Math.floor(serial)) > 1 / 86400) {
      s += ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
    }
    return s;
  }

  /* 判断数字格式是不是日期格式 */
  function isDateFormat(z) {
    if (!z) return false;
    var s = String(z);
    if (/\[[^\]]*\]/.test(s) && !/\[h+\]|\[m+\]|\[s+\]/i.test(s)) return false;
    var t = s.replace(/"[^"]*"/g, '').replace(/\\[^\s]/g, '').replace(/\[[^\]]*\]/g, '');
    if (/[ym]/i.test(t) && /[dmh]/i.test(t)) return true;
    if (/^[ym]{2,4}-[dm]{2}/i.test(t)) return true;
    if (/[dD]{2,4}\s*[月年]/.test(t)) return true;
    return false;
  }

  /* 从工作表里读出「哪几列是日期列」，返回 { 列号: 格式串 } */
  function scanDateFormats(ws, XL, maxRows) {
    var out = Object.create(null);
    if (!ws || !ws['!ref']) return out;
    var range = XL.utils.decode_range(ws['!ref']);
    var limit = Math.min(range.e.r, range.s.r + (maxRows || 60));
    for (var r = range.s.r; r <= limit; r++) {
      for (var c = range.s.c; c <= range.e.c; c++) {
        var cell = ws[XL.utils.encode_cell({ r: r, c: c })];
        if (!cell) continue;
        var z = cell.z;
        if (!z) continue;
        if (cell.t === 'd' || (cell.t === 'n' && isDateFormat(z))) {
          var idx = c - range.s.c;
          if (!(idx in out)) out[idx] = z;
        }
      }
    }
    return out;
  }

  /* 数字 → 料号文本：避免 1e+21 / 10023.0 这类变形 */
  function numToText(n) {
    if (!isFinite(n)) return String(n);
    if (Number.isInteger(n)) {
      if (Math.abs(n) < 1e21) return n.toFixed(0);
      return String(n);
    }
    var s = String(n);
    if (s.indexOf('e') >= 0 || s.indexOf('E') >= 0) return n.toFixed(20).replace(/0+$/, '').replace(/\.$/, '');
    return s;
  }

  /* 单元格原始值 → 可比较的字符串（不做业务归一化）
     z 可选：该单元格的数字格式串。带日期格式的数字按日期文本输出。 */
  function cellText(v, z) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') {
      if (z && isDateFormat(z) && v > 0) return dateText(v);
      return numToText(v);
    }
    if (v instanceof Date) return dateToText(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'object') {
      if (v instanceof Date) return dateToText(v);
      if (v.z && typeof v.v === 'number' && isDateFormat(v.z)) return dateText(v.v);
      if (typeof v.text === 'string') return v.text;          // 富文本
      if (v.result !== undefined) return cellText(v.result, z); // 公式结果
      if (Array.isArray(v.richText)) return v.richText.map(function (r) { return r.text || ''; }).join('');
      if (v.richText) return v.richText;
    }
    return String(v);
  }

  /* 归一化：用于“是否同一个料号”的判断，只在比对时使用，不改动原值 */
  function normKey(v, o) {
    var s = cellText(v);
    if (!s) return '';
    s = s.normalize ? s.normalize('NFKC') : s;
    s = s.replace(INVIS_RE, '').replace(/\u00A0/g, ' ');
    if (o && o.ignoreSpace) s = s.replace(/[\s\-_\uFF0F/\\.。·•]/g, '');
    else s = s.replace(/^\s+|\s+$/g, '');
    if (o && o.ignoreCase) s = s.toUpperCase();
    if (o && o.ignoreZero) s = s.replace(ZERO_RE, '');
    if (s === '' && cellText(v) !== '') {
      // 全是 0 的料号：保留 0，避免与空值混淆
      s = '0';
    }
    return s;
  }

  /* ---------- 文本解码：UTF-8 优先，失败回退 GB18030 ---------- */
  function decodeText(bytes) {
    var u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    // BOM
    if (u8.length >= 3 && u8[0] === 0xEF && u8[1] === 0xBB && u8[2] === 0xBF) {
      return { text: decode(u8.subarray(3), 'utf-8'), encoding: 'UTF-8' };
    }
    if (u8.length >= 2 && u8[0] === 0xFF && u8[1] === 0xFE) {
      return { text: decode(u8.subarray(2), 'utf-16le'), encoding: 'UTF-16LE' };
    }
    if (u8.length >= 2 && u8[0] === 0xFE && u8[1] === 0xFF) {
      return { text: decode(u8.subarray(2), 'utf-16be'), encoding: 'UTF-16BE' };
    }
    try {
      return { text: decode(u8, 'utf-8', true), encoding: 'UTF-8' };
    } catch (e) {
      try {
        return { text: decode(u8, 'gb18030'), encoding: 'GB18030' };
      } catch (e2) {
        return { text: decode(u8, 'utf-8'), encoding: 'UTF-8' };
      }
    }
  }

  function decode(u8, enc, fatal) {
    var td = new TextDecoder(enc, { fatal: !!fatal });
    return td.decode(u8);
  }

  /* ---------- 编码清单（txt）---------- */
  function parseCodeList(text) {
    var raw = String(text).replace(/^\uFEFF/, '');
    var lines = raw.split(/\r\n|\r|\n/);
    var out = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (!line) continue;
      var t = line.trim();
      if (!t) continue;
      if (t.charAt(0) === '#' || t.slice(0, 2) === '//') continue;
      var cells = t.split(SEP_RE);
      for (var j = 0; j < cells.length; j++) {
        var v = cells[j].replace(TRIM_PAIRS_RE, '').trim();
        if (v) out.push(v);
      }
    }
    return out;
  }

  /* 统计清单（按当前归一化选项） */
  function listStats(codes, opts) {
    var map = Object.create(null);
    var unique = 0, dup = 0;
    var keys = new Array(codes.length);
    for (var i = 0; i < codes.length; i++) {
      var k = normKey(codes[i], opts);
      keys[i] = k;
      if (map[k]) { dup++; map[k]++; }
      else { map[k] = 1; unique++; }
    }
    return { keys: keys, map: map, unique: unique, dup: dup, total: codes.length };
  }

  /* ---------- 明细表（AOA）----------
     opts.header     首行是表头
     opts.dateCols   { 列号: 格式串 }，这些列里的数字按 Excel 序列号解成日期
     列 0 是料号列，永不按日期解。 */
  function tableFromAOA(aoa, opts) {
    var rows = Array.isArray(aoa) ? aoa : [];
    var hasHeader = !!(opts && opts.header);
    var dateCols = (opts && opts.dateCols) || null;
    var headerRow = hasHeader ? (rows[0] || []) : [];
    var body = hasHeader ? rows.slice(1) : rows.slice(0);

    var width = 0;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r && r.length > width) width = r.length;
    }
    if (!width) width = 1;

    var columns = [];
    for (var c = 0; c < width; c++) {
      var title = cellText(headerRow[c]).trim();
      columns.push({
        index: c,
        title: title || ('第' + (c + 1) + '列'),
        isCode: c === 0
      });
    }

    var data = [];
    for (var r2 = 0; r2 < body.length; r2++) {
      var row = body[r2] || [];
      var empty = true;
      for (var k = 0; k < width; k++) {
        if (cellText(row[k]) !== '') { empty = false; break; }
      }
      if (empty) continue;
      var cells = new Array(width);
      for (var m = 0; m < width; m++) {
        var val = row[m] === undefined ? '' : row[m];
        // 日期列里的序列号转成本地日期对象：预览与导出口径一致
        if (m > 0 && dateCols && dateCols[m] && typeof val === 'number' && val > 0) {
          val = serialToDate(val);
        }
        cells[m] = val;
      }
      data.push(cells);
    }

    return { columns: columns, rows: data, hasHeader: hasHeader, width: width, dateCols: dateCols };
  }

  /* 判断首行像不像表头 */
  function looksLikeHeader(aoa, opts) {
    if (!aoa || !aoa.length) return false;
    var first = aoa[0] || [];
    var second = aoa[1] || [];
    if (cellText(first[0]).trim() === '') return false;
    if (cellText(second[0]).trim() === '' && aoa.length > 1) return false;
    if (aoa.length < 2) return false;
    var keys = Object.create(null);
    if (opts && opts.listKeys) {
      for (var i = 0; i < opts.listKeys.length; i++) keys[opts.listKeys[i]] = 1;
    }
    var k1 = normKey(first[0], opts);
    var k2 = normKey(second[0], opts);
    var inList1 = !!keys[k1], inList2 = !!keys[k2];
    if (inList1 && !inList2) return false;  // 首行就是清单里的料号 → 没有表头
    if (inList2 && !inList1) return true;   // 第二行才命中 → 首行是表头
    if (inList1 && inList2) return false;   // 两行都是料号 → 没有表头
    // 都不在清单里：看首行是否是“名字型”文本
    var t = cellText(first[0]).trim();
    if (/[\u4e00-\u9fa5]/.test(t)) return true;
    if (/code|编码|料号|物料|编号|sku|part/i.test(t)) return true;
    var numeric = /^[0-9A-Za-z.\-_]+$/.test(t);
    return !numeric;
  }

  /* ---------- 比对 ---------- */
  function runMatch(input) {
    var codes = input.codes || [];
    var table = input.table;
    var opts = input.opts || {};
    var outMode = input.outMode || 'hit';

    var list = listStats(codes, opts);
    var hitSet = Object.create(null);
    var usedCount = Object.create(null);
    var rows = [];
    var hitRows = 0, warnRows = 0;

    for (var i = 0; i < table.rows.length; i++) {
      var cells = table.rows[i];
      var key = normKey(cells[0], opts);
      var hit = !!(key && list.map[key]);
      if (hit) { hitRows++; hitSet[key] = 1; usedCount[key] = (usedCount[key] || 0) + 1; }
      if (outMode === 'all') rows.push({ index: i, cells: cells, hit: hit });
      else if (outMode === 'miss') { if (!hit) rows.push({ index: i, cells: cells, hit: hit }); }
      else if (hit) rows.push({ index: i, cells: cells, hit: hit });
      if (!hit && cellText(cells[0]) === '') warnRows++;
    }

    var hitCodes = [];
    var unused = [];
    for (var u = 0; u < list.keys.length; u++) {
      var k = list.keys[u];
      if (!k) continue;
      if (hitSet[k]) { if (!hitCodes.some(function (x) { return x === k; })) {} }
    }
    // 去重统计
    var seen = Object.create(null);
    for (var a = 0; a < list.keys.length; a++) {
      var kk = list.keys[a];
      if (!kk || seen[kk]) continue;
      seen[kk] = 1;
      if (hitSet[kk]) hitCodes.push(kk); else unused.push(kk);
    }

    // 未命中的原始料号（保留首次出现的原始写法）
    var unusedRaw = [];
    var rawSeen = Object.create(null);
    for (var b = 0; b < codes.length; b++) {
      var k2 = list.keys[b];
      if (!k2 || hitSet[k2]) continue;
      if (rawSeen[k2]) continue;
      rawSeen[k2] = 1;
      unusedRaw.push(String(codes[b]));
    }

    return {
      rows: rows,
      hitRows: hitRows,
      scanned: table.rows.length,
      uniqueCodes: list.unique,
      listedTotal: list.total,
      listDups: list.dup,
      hitCodeCount: hitCodes.length,
      unusedCount: unused.length,
      unusedRaw: unusedRaw,
      emptyCodeRows: warnRows,
      outMode: outMode
    };
  }

  /* ---------- 导出 ---------- */
  function cjkWidth(s) {
    var w = 0;
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      w += (c > 0x2E7F && c < 0x9FFF) || (c >= 0x3000 && c <= 0x303F) || (c >= 0xFF00 && c <= 0xFF60) ? 2 : 1;
    }
    return w;
  }

  function guessWidth(title, rows, colIndex, maxLen) {
    var w = cjkWidth(String(title || ''));
    for (var i = 0; i < Math.min(rows.length, 200); i++) {
      var v = cellText(rows[i][colIndex]);
      if (v.length > maxLen) v = v.slice(0, maxLen);
      var l = cjkWidth(v);
      if (l > w) w = l;
    }
    return Math.max(6, Math.min(44, w + 3));
  }

  function buildExport(libs, payload) {
    var table = payload.table;
    var result = payload.result;
    var meta = payload.meta || {};
    var opts = payload.opts || {};
    var ExcelJS = libs && libs.ExcelJS;
    var XLSX = libs && libs.XLSX;

    var codeHeader = table.columns[0] ? table.columns[0].title : '物料编码';
    var headers = ['序号', codeHeader].concat(table.columns.slice(1).map(function (c) { return c.title; }));
    if (result.outMode === 'all') headers.push('是否命中');

    var aoa = [];
    aoa.push(headers);
    for (var i = 0; i < result.rows.length; i++) {
      var r = result.rows[i];
      var line = [i + 1, r.cells[0]];
      for (var c = 1; c < table.width; c++) line.push(r.cells[c] === undefined ? '' : r.cells[c]);
      if (result.outMode === 'all') line.push(r.hit ? '是' : '否');
      aoa.push(line);
    }

    // 日期列：Date 换成 Excel 序列号再写，避免 ExcelJS/SheetJS 的时区偏移
    var dateCols = Object.create(null);
    var srcDate = table.dateCols || null;
    for (var d = 1; d < table.width; d++) {
      if (srcDate && srcDate[d]) dateCols[d + 1] = srcDate[d];   // 导出的第 1 列是序号
    }
    for (var ri = 1; ri < aoa.length; ri++) {
      for (var key in dateCols) {
        var ci = parseInt(key, 10);
        if (aoa[ri][ci] instanceof Date) aoa[ri][ci] = dateToSerial(aoa[ri][ci]);
      }
    }

    if (ExcelJS) return buildWithExcelJS(ExcelJS, aoa, headers, table, result, meta, opts, dateCols);
    if (XLSX) return buildWithXLSX(XLSX, aoa, headers, table, result, meta, dateCols);
    return Promise.reject(new Error('缺少导出组件'));
  }

  /* Excel 的日期格式串：原表带时间就保留时间 */
  function dateNumFmt(z) {
    if (z && /h|s/i.test(String(z).replace(/"[^"]*"/g, ''))) return 'yyyy-mm-dd hh:mm';
    return 'yyyy-mm-dd';
  }

  function buildWithExcelJS(ExcelJS, aoa, headers, table, result, meta, opts, dateCols) {
    var wb = new ExcelJS.Workbook();
    wb.creator = '料号比对';
    wb.created = new Date();

    var ws = wb.addWorksheet('比对结果', {
      views: [{ state: 'frozen', xSplit: 2, ySplit: 1 }]
    });

    ws.columns = headers.map(function (h, i) {
      return { header: h, key: 'c' + i, width: guessWidth(h, table.rows, i === 0 ? 0 : i - 1, 40) };
    });

    for (var i = 1; i < aoa.length; i++) ws.addRow(aoa[i]);

    // 日期列的显示格式
    if (dateCols && aoa.length > 1) {
      for (var dk in dateCols) {
        var ci = parseInt(dk, 10) + 1;          // ExcelJS 列号从 1 开始
        if (ci < 2 || ci > headers.length) continue;
        var fmt = dateNumFmt(dateCols[dk]);
        for (var r = 2; r < aoa.length + 1; r++) {
          var cell = ws.getRow(r).getCell(ci);
          cell.numFmt = fmt;
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
      }
    }

    var lastCol = headers.length;
    var lastRow = aoa.length;

    // 表头
    var head = ws.getRow(1);
    head.height = 22;
    head.font = { bold: true, size: 10.5, color: { argb: 'FFFFFFFF' }, name: '微软雅黑' };
    head.alignment = { vertical: 'middle', horizontal: 'center' };
    head.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1C3F35' } };
    head.eachCell(function (cell, col) {
      cell.border = { bottom: { style: 'thin', color: { argb: 'FF39DDA9' } } };
      if (col === 2) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF26604F' } };
    });

    // 正文基础样式
    ws.eachRow({ includeEmpty: false }, function (row, n) {
      if (n === 1) return;
      row.font = { size: 10.5, color: { argb: 'FF1B2B26' }, name: '微软雅黑' };
      row.alignment = { vertical: 'middle' };
      var c1 = row.getCell(1);
      c1.alignment = { vertical: 'middle', horizontal: 'center' };
      c1.font = { size: 10, color: { argb: 'FF7C8C86' }, name: '微软雅黑' };
      row.getCell(2).font = { size: 10.5, bold: true, color: { argb: 'FF11483B' }, name: '微软雅黑' };
    });

    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: lastRow, column: lastCol } };

    if (lastRow > 1) {
      ws.addConditionalFormatting({
        ref: 'A2:' + colName(lastCol) + lastRow,
        rules: [{
          type: 'expression',
          priority: 1,
          formulae: ['MOD(ROW(),2)=0'],
          style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F7F5' } } }
        }]
      });
    }

    // 未命中
    if (result.unusedRaw && result.unusedRaw.length) {
      var ws2 = wb.addWorksheet('未命中料号', { views: [{ state: 'frozen', ySplit: 1 }] });
      ws2.columns = [{ header: '序号', key: 'i', width: 8 }, { header: '清单里的料号', key: 'c', width: 28 }];
      for (var u = 0; u < result.unusedRaw.length; u++) ws2.addRow([u + 1, result.unusedRaw[u]]);
      var h2 = ws2.getRow(1);
      h2.font = { bold: true, size: 10.5, color: { argb: 'FFFFFFFF' }, name: '微软雅黑' };
      h2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7A5C22' } };
      h2.alignment = { vertical: 'middle', horizontal: 'center' };
      h2.height = 22;
      ws2.eachRow({ includeEmpty: false }, function (row, n) {
        if (n === 1) return;
        row.font = { size: 10.5, color: { argb: 'FF1B2B26' }, name: '微软雅黑' };
        row.getCell(1).alignment = { horizontal: 'center' };
        row.getCell(1).font = { size: 10, color: { argb: 'FF7C8C86' }, name: '微软雅黑' };
      });
    }

    // 说明
    var ws3 = wb.addWorksheet('比对说明');
    ws3.columns = [{ width: 15 }, { width: 52 }];
    var info = [
      ['生成时间', fmtTime(new Date())],
      ['编码清单', meta.listName || '—'],
      ['清单料号数', String(result.listedTotal) + ' 条（去重后 ' + result.uniqueCodes + ' 个）'],
      ['明细表', meta.tableName || '—'],
      ['工作表', meta.sheetName || '—'],
      ['扫描行数', String(result.scanned)],
      ['命中行数', String(result.hitRows)],
      ['命中料号', result.hitCodeCount + ' 个 / 共 ' + result.uniqueCodes + ' 个'],
      ['忽略大小写', opts.ignoreCase ? '是' : '否'],
      ['忽略前导零', opts.ignoreZero ? '是' : '否'],
      ['忽略空格与分隔符', opts.ignoreSpace ? '是' : '否'],
      ['输出范围', OUT_MODE_LABEL[result.outMode] || OUT_MODE_LABEL.hit]
    ];
    info.forEach(function (pair, i) {
      var row = ws3.addRow(pair);
      row.height = 20;
      row.getCell(1).font = { size: 10.5, color: { argb: 'FF5D726B' }, name: '微软雅黑' };
      row.getCell(2).font = { size: 10.5, color: { argb: 'FF15241F' }, name: '微软雅黑' };
      if (i === 0) {
        row.getCell(1).font = { size: 10.5, bold: true, color: { argb: 'FF1C3F35' }, name: '微软雅黑' };
      }
    });

    return wb.xlsx.writeBuffer();
  }

  function buildWithXLSX(XLSX, aoa, headers, table, result, meta, dateCols) {
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = headers.map(function (h, i) { return { wch: guessWidth(h, table.rows, i === 0 ? 0 : i - 1, 40) }; });
    if (dateCols) {
      for (var dk in dateCols) {
        var ci = parseInt(dk, 10);
        var fmt = dateNumFmt(dateCols[dk]);
        for (var r = 1; r < aoa.length; r++) {
          var addr = XLSX.utils.encode_cell({ r: r, c: ci });
          if (ws[addr] && ws[addr].t === 'n') ws[addr].z = fmt;
        }
      }
    }
    if (aoa.length > 1) ws['!autofilter'] = { ref: 'A1:' + colName(headers.length) + aoa.length };
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '比对结果');
    if (result.unusedRaw && result.unusedRaw.length) {
      var a2 = [['序号', '清单里的料号']];
      for (var i = 0; i < result.unusedRaw.length; i++) a2.push([i + 1, result.unusedRaw[i]]);
      var ws2 = XLSX.utils.aoa_to_sheet(a2);
      ws2['!cols'] = [{ wch: 8 }, { wch: 28 }];
      XLSX.utils.book_append_sheet(wb, ws2, '未命中料号');
    }
    var a3 = [['生成时间', fmtTime(new Date())], ['清单料号数', String(result.listedTotal)], ['扫描行数', String(result.scanned)], ['命中行数', String(result.hitRows)]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(a3), '比对说明');
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
  }

  function colName(n) {
    var s = '';
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function fmtTime(d) {
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
  }

  var CORE = {
    cellText: cellText,
    normKey: normKey,
    parseCodeList: parseCodeList,
    listStats: listStats,
    tableFromAOA: tableFromAOA,
    looksLikeHeader: looksLikeHeader,
    runMatch: runMatch,
    buildExport: buildExport,
    decodeText: decodeText,
    fmtTime: fmtTime,
    colName: colName,
    serialToDate: serialToDate,
    dateToSerial: dateToSerial,
    dateText: dateText,
    isDateFormat: isDateFormat,
    scanDateFormats: scanDateFormats
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = CORE;
  else root.CORE = CORE;
})(typeof globalThis !== 'undefined' ? globalThis : this);
