/* 料号比对 — 界面层 */
(function () {
  'use strict';

  var C = window.CORE;
  var $ = function (id) { return document.getElementById(id); };

  var state = {
    list: null,       // { name, codes, encoding }
    table: null,      // { name, sheets, sheetName, raw:{aoa}, parsed }
    headerAuto: true,
    result: null,
    previewCount: 100,
    busy: false
  };

  var E = {
    bench: $('bench'),
    run: $('run'),
    runLabel: $('run-label'),
    status: $('status'),
    result: $('result'),
    toast: $('toast')
  };

  /* ---------------- 工具 ---------------- */
  var toastTimer = null;
  function toast(msg) {
    E.toast.innerHTML = msg;
    E.toast.classList.add('is-on');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { E.toast.classList.remove('is-on'); }, 4200);
  }

  function fmtInt(n) {
    return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }

  function setBusy(on) {
    state.busy = on;
    E.bench.classList.toggle('is-busy', on);
    E.run.disabled = on || !ready();
    E.runLabel.textContent = on ? '正在比对…' : (state.result ? '重新比对' : '开始比对');
  }

  function ready() {
    return !!(state.list && state.table && state.table.parsed);
  }

  function refreshStatus() {
    var missing = [];
    if (!state.list) missing.push('编码清单');
    if (!state.table) missing.push('明细表');
    if (missing.length === 2) E.status.textContent = '还差 2 份文件';
    else if (missing.length === 1) E.status.innerHTML = '还差<b>' + missing[0] + '</b>';
    else if (state.result) E.status.innerHTML = '已经比对过，改设置会<b>自动重算</b>';
    else E.status.innerHTML = '两份文件都读进来了，<b>可以开始</b>';
    E.run.disabled = state.busy || !ready();
  }

  function readOpts() {
    return {
      ignoreCase: $('opt-case').checked,
      ignoreZero: $('opt-zero').checked,
      ignoreSpace: $('opt-space').checked
    };
  }

  function outMode() {
    var el = document.querySelector('input[name="outmode"]:checked');
    return el ? el.value : 'hit';
  }

  /* 导出文件名前缀，跟着输出范围走 */
  var FILE_TAG = { hit: '料号比对结果', all: '料号比对结果_全部行', miss: '料号比对结果_未命中行' };

  /* 输出范围下方的一句说明 */
  var SEG_CAP = {
    hit: '只有料号对得上的行会写进结果表。',
    all: '每一行都会写进结果表，最后一列标注命中与否。',
    miss: '只列出料号对不上的行。日常核对差异用这个。'
  };

  /* ---------------- 文件读取 ---------------- */
  function readBuffer(file) {
    return new Promise(function (res, rej) {
      var fr = new FileReader();
      fr.onload = function () { res(new Uint8Array(fr.result)); };
      fr.onerror = function () { rej(fr.error || new Error('读取失败')); };
      fr.readAsArrayBuffer(file);
    });
  }

  function isExcel(name) {
    return /\.(xlsx|xlsm|xlsb|xls)$/i.test(name);
  }

  /* ---------------- 编码清单 ---------------- */
  function acceptList(file) {
    if (!file) return;
    if (/\.(xlsx|xlsm|xlsb|xls)$/i.test(file.name)) {
      showErr('list', '这是 Excel 文件。编码清单请用 txt（一行一个料号），明细表才放 Excel。');
      return;
    }
    readBuffer(file).then(function (bytes) {
      var dec = C.decodeText(bytes);
      var codes = C.parseCodeList(dec.text);
      if (!codes.length) {
        showErr('list', '文件里没读到任何编码。请确认是一行一个料号，且不是空文件。');
        return;
      }
      state.list = { name: file.name, codes: codes, encoding: dec.encoding, size: bytes.length };
      showErr('list', '');
      renderList();
      refreshStatus();
      invalidate();
    }).catch(function (err) {
      showErr('list', '读不出来：' + (err && err.message ? err.message : '未知错误'));
    });
  }

  function renderList() {
    var L = state.list;
    if (!L) return;
    var stats = C.listStats(L.codes, readOpts());
    $('drop-list').hidden = true;
    $('filed-list').hidden = false;
    $('zone-list').classList.add('is-loaded');
    $('list-name').textContent = L.name;
    $('list-count').textContent = fmtInt(stats.total);
    $('list-dups').textContent = fmtInt(stats.dup);
    $('list-enc').textContent = L.encoding;

    var peek = $('list-peek');
    var seen = Object.create(null);
    var uniq = [];
    // 只铺得下两行的量：多出来的用 +N 收口，剩下的交给渐变淡出
    for (var i = 0; i < L.codes.length && uniq.length < 10; i++) {
      var s = String(L.codes[i]);
      var k = s.toUpperCase();
      if (seen[k]) continue;
      seen[k] = 1;
      uniq.push(s);
    }
    peek.innerHTML = uniq.map(function (s) {
      return '<span class="chip">' + esc(s) + '</span>';
    }).join('') + (stats.unique > uniq.length ? '<span class="chip chip--more">+' + fmtInt(stats.unique - uniq.length) + '</span>' : '');
  }

  /* ---------------- 明细表 ---------------- */
  function acceptTable(file) {
    if (!file) return;
    var name = file.name;
    if (/\.(xls)$/i.test(name)) {
      // 老格式：交给 SheetJS
    }
    readBuffer(file).then(function (bytes) {
      // cellDates 关掉：让日期留在原始序列号里，稍后按列格式自己换算，
      // 避免两个 Excel 库在时区上来回漂一天。
      var readOpt = { type: 'array', cellDates: false, cellNF: true, cellText: false };
      var wb;
      if (isExcel(name)) {
        wb = XLSX.read(bytes, readOpt);
      } else {
        var dec = C.decodeText(bytes);
        wb = XLSX.read(dec.text, { type: 'string', cellDates: false, raw: false });
      }
      var sheets = wb.SheetNames || [];
      if (!sheets.length) throw new Error('工作簿里没有工作表');
      state.table = {
        name: name,
        wb: wb,
        sheets: sheets,
        sheetName: sheets[0],
        dateCols: null,
        raw: null,
        parsed: null
      };
      $('table-swap').hidden = false;
      loadSheet(sheets[0], true);
    }).catch(function (err) {
      showErr('table', '读不出来：' + (err && err.message ? err.message : '未知错误'));
    });
  }

  function loadSheet(sheetName, guessHeader) {
    var T = state.table;
    var ws = T.wb.Sheets[sheetName];
    if (!ws) return;
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: false });
    T.sheetName = sheetName;
    T.raw = aoa;
    // 先嗅出哪些列是日期格式，后面的数字才能正确还原成日期
    T.dateCols = C.scanDateFormats(ws, XLSX, 80);

    if (guessHeader) {
      var guess = guessHeaderOn(aoa);
      $('opt-header').checked = guess;
      state.headerAuto = true;
    }
    rebuildTable();
  }

  function guessHeaderOn(aoa) {
    var listKeys = null;
    var opts = readOpts();
    if (state.list) {
      listKeys = C.listStats(state.list.codes, opts).keys;
    }
    // 归一化选项要和比对时完全一致，否则判断会自相矛盾
    return C.looksLikeHeader(aoa, {
      listKeys: listKeys,
      ignoreCase: opts.ignoreCase,
      ignoreZero: opts.ignoreZero,
      ignoreSpace: opts.ignoreSpace
    });
  }

  function rebuildTable() {
    var T = state.table;
    if (!T || !T.raw) return;
    T.parsed = C.tableFromAOA(T.raw, { header: $('opt-header').checked, dateCols: T.dateCols });
    showErr('table', '');
    renderTable();
    refreshStatus();
    refreshIfMatched();
  }

  function renderTable() {
    var T = state.table;
    if (!T || !T.parsed) return;
    var P = T.parsed;
    $('drop-table').hidden = true;
    $('filed-table').hidden = false;
    $('zone-table').classList.add('is-loaded');
    $('table-name').textContent = T.name;
    $('table-count').textContent = fmtInt(P.rows.length);
    $('table-cols').textContent = fmtInt(P.width);

    var nonEmpty = 0;
    for (var i = 0; i < P.rows.length; i++) {
      if (C.cellText(P.rows[i][0]) !== '') nonEmpty++;
    }
    $('table-codes').textContent = fmtInt(nonEmpty);

    // 工作表选择
    var sw = $('sheet-wrap'), sel = $('sheet-sel');
    if (T.sheets.length > 1) {
      sw.hidden = false;
      sel.innerHTML = T.sheets.map(function (s) {
        return '<option' + (s === T.sheetName ? ' selected' : '') + '>' + esc(s) + '</option>';
      }).join('');
    } else {
      sw.hidden = true;
    }

    // 首行预览
    var peek = $('table-peek');
    var first = P.rows[0] || [];
    var items = [];
    for (var c = 0; c < Math.min(P.width, 4); c++) {
      var v = C.cellText(first[c]);
      if (v.length > 13) v = v.slice(0, 13) + '…';
      items.push('<span class="chip">' + esc(v || '—') + '</span>');
    }
    if (P.width > 4) items.push('<span class="chip chip--more">+' + (P.width - 4) + ' 列</span>');
    peek.innerHTML = items.join('');
  }

  /* ---------------- 比对 ---------------- */
  /* 任何输入或设置变化：清掉旧结果、把「开始比对」按钮放回来。
     比对完成后按钮会被收起（结果区自己有导出按钮），
     如果这里不还原，换选项后就没地方点重新比对了。 */
  function invalidate() {
    var had = !!state.result;
    if (had) {
      state.result = null;
      E.result.hidden = true;
    }
    if (E.run.style.display === 'none') E.run.style.display = '';
    E.runLabel.textContent = '开始比对';
    refreshStatus();
  }

  function runMatch(quiet) {
    if (!ready() || state.busy) return;
    var t0 = performance.now();
    setBusy(true);
    if (!quiet) E.result.hidden = true;

    requestAnimationFrame(function () {
      setTimeout(function () {
        try {
          var res = C.runMatch({
            codes: state.list.codes,
            table: state.table.parsed,
            opts: readOpts(),
            outMode: outMode()
          });
          res.elapsed = Math.round(performance.now() - t0);
          state.result = res;
          state.previewCount = 100;
          renderResult(res);
          setBusy(false);
          if (!quiet) E.result.scrollIntoView({ behavior: 'smooth', block: 'start' });
          E.runLabel.textContent = '重新比对';
          if (state.rerunPending) { state.rerunPending = false; runMatch(true); }
        } catch (err) {
          setBusy(false);
          toast('比对出错：' + (err && err.message ? err.message : '未知错误'));
        }
      }, 30);
    });
  }

  /* 已经比对过、只是换了个开关或输出范围：直接按新设置重算，不跳回顶部。
     「开始比对」按钮始终留在原地，不会消失。 */
  function refreshIfMatched() {
    if (state.busy) { state.rerunPending = true; return; }
    if (state.result) runMatch(true);
    else invalidate();
  }

  function renderResult(res) {
    E.result.hidden = false;

    var allMode = res.outMode === 'all';
    var missMode = res.outMode === 'miss';
    $('n-rows').textContent = fmtInt(res.outMode === 'hit' ? res.hitRows : res.rows.length);
    $('n-rows-l').textContent = allMode ? '行写进结果' : (missMode ? '行没命中' : '行命中');
    $('n-codes').textContent = fmtInt(res.hitCodeCount);
    $('n-unused').textContent = fmtInt(res.unusedCount);
    $('n-scan').textContent = fmtInt(res.scanned);

    var pct = res.uniqueCodes ? (res.hitCodeCount / res.uniqueCodes) * 100 : 0;
    var pctText = pct === 0 ? '0' : (pct < 1 ? pct.toFixed(1) : Math.round(pct).toString());
    $('rfill').style.width = Math.max(pct, res.hitCodeCount ? 0.6 : 0) + '%';
    var headAt = Math.min(Math.max(pct, 0), 99.4);
    $('rhead').style.left = headAt + '%';
    // 靠右时把气泡翻到线左边，否则它会捅出容器把窄屏撑宽
    $('rhead').classList.toggle('is-end', headAt > 82);
    $('rpct').textContent = pctText + '%';
    $('rcap').textContent = '清单 ' + fmtInt(res.uniqueCodes) + ' 个料号里有 ' +
      fmtInt(res.hitCodeCount) + ' 个在明细表里出现过 · 用时 ' + res.elapsed + ' ms';

    renderNotice(res);
    renderPreview(res);
    // 未命中行模式即使一行都没列出来，清单里也可能有未出现的料号可导出
    var hasUnused = res.unusedRaw && res.unusedRaw.length;
    var none = res.rows.length === 0 && !(missMode && hasUnused);
    E.run.style.display = '';
    var ex = $('export');
    ex.disabled = none;
    $('export-label').textContent = none
      ? '没有可导出的行'
      : (res.rows.length === 0 ? '导出未命中料号' : '导出 Excel');
    refreshStatus();
  }

  function renderNotice(res) {
    var n = $('notice');
    var parts = [];
    if (res.outMode === 'all') {
      if (res.hitRows === res.scanned) {
        parts.push('<b>全部命中。</b>明细表 ' + fmtInt(res.scanned) + ' 行都没有被剔除，最后一列全部标了「是」。');
      } else {
        parts.push('<b>比对完成。</b>明细表 ' + fmtInt(res.scanned) + ' 行原样写进结果，其中 ' +
          fmtInt(res.hitRows) + ' 行的料号在清单里，另外 ' + fmtInt(res.scanned - res.hitRows) +
          ' 行标了「否」。最后一列就是命中标记。');
      }
      n.className = 'notice notice--ok';
      n.hidden = false;
      n.innerHTML = parts.join('');
      return;
    }
    if (res.outMode === 'miss') {
      var rest2 = res.scanned - res.hitRows;
      var hasUnused0 = !!(res.unusedRaw && res.unusedRaw.length);
      if (rest2 === 0) {
        parts.push('<b>没有漏网的。</b>明细表 ' + fmtInt(res.scanned) +
          ' 行的料号全部在清单里出现过，所以未命中行是空的。');
        if (hasUnused0) {
          parts.push('<ul><li>反过来，清单里有 ' + fmtInt(res.unusedCount) +
            ' 个料号在明细表里一条都没出现，这些记在导出的「未命中料号」表里</li></ul>');
        }
        n.className = hasUnused0 ? 'notice' : 'notice notice--ok';
      } else {
        parts.push('<b>比对完成。</b>下面就是明细表里对不上的 ' + fmtInt(rest2) +
          ' 行，其余 ' + fmtInt(res.hitRows) + ' 行没有列出来。');
        parts.push('<ul><li>这些行的料号没有在清单里出现，也可能是写法不一致</li>' +
          '<li>可以试试点亮「忽略前导零」「忽略空格与分隔符」再比对一次</li>' +
          '<li>清单里还有 ' + fmtInt(res.unusedCount) + ' 个料号在明细表里一条都没出现，见导出的「未命中料号」表</li></ul>');
        n.className = 'notice';
      }
      n.hidden = false;
      n.innerHTML = parts.join('');
      return;
    }
    if (res.hitRows === 0) {
      parts.push('<b>一行都没对上。</b>两份文件里的编码看起来不是同一套写法。下面这些可以先试试：');
      var tips = [];
      if (!$('opt-space').checked) tips.push('打开「忽略空格与分隔符」，AB-01 与 AB01 就能对上');
      if (!$('opt-zero').checked) tips.push('打开「忽略前导零」，00123 与 123 就能对上');
      tips.push('打开「首行是表头」或关掉它，确认第一列读的是料号而不是标题');
      if (state.table && state.table.sheets.length > 1) tips.push('换一个工作表，料号可能在别的表里');
      parts.push('<ul><li>' + tips.join('</li><li>') + '</li></ul>');
      n.className = 'notice';
    } else if (res.hitRows === res.scanned) {
      parts.push('<b>全部命中。</b>明细表里每一行的料号都出现在清单里，导出的表就是原表的完整复制。');
      n.className = 'notice notice--ok';
    } else {
      var rest = res.scanned - res.hitRows;
      parts.push('<b>比对完成。</b>明细表 ' + fmtInt(res.scanned) + ' 行里有 ' + fmtInt(res.hitRows) +
        ' 行对上，另外 ' + fmtInt(rest) + ' 行没有命中，已经剔除。');
      if (res.emptyCodeRows) parts.push('<ul><li>其中 ' + fmtInt(res.emptyCodeRows) + ' 行的第一列是空的</li></ul>');
      n.className = 'notice notice--ok';
    }
    n.innerHTML = parts.join('');
    n.hidden = false;
  }

  function renderPreview(res) {
    var P = state.table.parsed;
    var head = $('tbl-head'), body = $('tbl-body');
    var cols = [
      { t: '#', cls: 'is-idx', idx: -1 },
      { t: P.columns[0].title, cls: 'is-code', idx: 0 }
    ];
    for (var c = 1; c < P.width; c++) cols.push({ t: P.columns[c].title, cls: '', idx: c });
    if (res.outMode === 'all') cols.push({ t: '是否命中', cls: 'num flag', idx: -2 });

    head.innerHTML = '<tr>' + cols.map(function (col) {
      return '<th' + (col.cls ? ' class="' + col.cls + '"' : '') + '>' + esc(col.t) + '</th>';
    }).join('') + '</tr>';

    var limit = Math.min(state.previewCount, res.rows.length);
    var html = [];
    for (var i = 0; i < limit; i++) {
      var r = res.rows[i];
      var tds = '<td class="is-idx">' + (i + 1) + '</td>';
      tds += '<td class="is-code">' + esc(C.cellText(r.cells[0])) + '</td>';
      for (var k = 1; k < P.width; k++) {
        var v = C.cellText(r.cells[k]);
        tds += '<td>' + (v === '' ? '<span style="opacity:.32">—</span>' : esc(v)) + '</td>';
      }
      if (res.outMode === 'all') {
        tds += '<td class="num" style="color:' + (r.hit ? 'var(--jade)' : 'var(--fg-3)') + '">' + (r.hit ? '是' : '否') + '</td>';
      }
      html.push('<tr>' + tds + '</tr>');
    }
    body.innerHTML = html.join('');

    var wrap = $('tablewrap'), moreWrap = $('more-wrap');
    if (!res.rows.length) {
      wrap.hidden = true;
      moreWrap.hidden = true;
      return;
    }
    wrap.hidden = false;
    var rest = res.rows.length - limit;
    if (rest > 0) {
      moreWrap.hidden = false;
      $('more').hidden = false;
      $('more-note').textContent = '只预览前 ' + fmtInt(limit) + ' 行，还剩 ' + fmtInt(rest) + ' 行；导出的 Excel 是完整的 ' + fmtInt(res.rows.length) + ' 行。';
    } else {
      moreWrap.hidden = res.rows.length <= 100;
      $('more').hidden = true;
      $('more-note').textContent = '共 ' + fmtInt(res.rows.length) + ' 行，已全部在导出的 Excel 里。';
    }
  }

  /* ---------------- 导出 ---------------- */
  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  /* 优先走「另存为」，让用户自己挑位置（默认落在「文档」）。
     实测：浏览器写盘一律带 ZoneId=3 标记，file:// 和 http:// 都一样，
     所以真正决定能不能打开的是目标目录有没有被 Excel 信任，不是导出通道。
     必须在点击手势里同步调用，否则浏览器拒绝。 */
  function openSavePicker(name) {
    if (typeof window.showSaveFilePicker !== 'function') return null;
    try {
      return window.showSaveFilePicker({
        suggestedName: name,
        startIn: 'documents',
        types: [{ description: 'Excel 工作簿', accept: makeTypeMap() }]
      });
    } catch (e) {
      return null;
    }
  }

  function makeTypeMap() {
    var t = {};
    t[XLSX_MIME] = ['.xlsx'];
    return t;
  }

  function toBytes(buf) {
    if (buf instanceof Uint8Array) return buf;
    if (buf instanceof ArrayBuffer) return new Uint8Array(buf);
    if (buf && buf.buffer instanceof ArrayBuffer) {
      return new Uint8Array(buf.buffer, buf.byteOffset || 0, buf.byteLength || buf.buffer.byteLength);
    }
    return new Uint8Array(buf);
  }

  function downloadBytes(bytes, name) {
    var blob = new Blob([bytes], { type: XLSX_MIME });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }

  function writeToHandle(handle, bytes) {
    return handle.createWritable().then(function (w) {
      return w.write(bytes).then(function () { return w.close(); });
    }).then(function () { return handle.name; });
  }

  function exportXlsx() {
    var res = state.result;
    if (!res || !state.table) return;
    var btn = $('export');
    var lab = $('export-label');
    btn.disabled = true;
    var old = lab.textContent;
    lab.textContent = '正在生成…';

    var d = new Date();
    var stamp = '' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + '_' + pad2(d.getHours()) + pad2(d.getMinutes());
    var fname = FILE_TAG[res.outMode] + '_' + stamp + '.xlsx';

    var picker = openSavePicker(fname);   /* 必须在手势内同步发起 */
    var cancelled = false;
    if (picker) {
      picker.then(null, function (e) { if (e && e.name === 'AbortError') cancelled = true; });
    }

    C.buildExport({ ExcelJS: window.ExcelJS, XLSX: window.XLSX }, {
      table: state.table.parsed,
      result: res,
      opts: readOpts(),
      meta: {
        listName: state.list.name + '（' + fmtInt(state.list.codes.length) + ' 行）',
        tableName: state.table.name,
        sheetName: state.table.sheetName
      }
    }).then(function (buf) {
      var bytes = toBytes(buf);
      var done = res.rows.length
        ? '已保存 <b>' + fmtInt(res.rows.length) + '</b> 行'
        : '已保存 <b>' + fmtInt((res.unusedRaw || []).length) + '</b> 个未出现的料号';
      if (!picker) {
        downloadBytes(bytes, fname);
        return { how: 'download', msg: done };
      }
      return picker.then(function (handle) {
        return writeToHandle(handle, bytes).then(function (name) {
          return { how: 'saved', msg: done, name: name };
        });
      }, function (err) {
        if (err && err.name === 'AbortError') throw err;
        /* 直接保存失败就退回下载，别让用户白等 */
        downloadBytes(bytes, fname);
        return { how: 'download', msg: done };
      });
    }).then(function (r) {
      if (r.how === 'saved') toast(r.msg + '　' + esc(r.name));
      else toast(r.msg);
    }).catch(function (err) {
      if (cancelled || (err && err.name === 'AbortError')) { toast('已取消保存'); return; }
      toast('导出失败：' + (err && err.message ? err.message : '未知错误'));
    }).then(function () {
      btn.disabled = false;
      lab.textContent = old || '导出 Excel';
    });
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  /* ---------------- DOM 事件 ---------------- */
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (m) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
    });
  }

  function showErr(which, msg) {
    var el = $(which + '-err');
    if (!el) return;
    if (!msg) { el.hidden = true; el.textContent = ''; return; }
    el.hidden = false;
    el.textContent = msg;
  }

  function bindDrop(dropEl, inputEl, acceptFn) {
    ['dragenter', 'dragover'].forEach(function (ev) {
      dropEl.addEventListener(ev, function (e) {
        e.preventDefault(); e.stopPropagation();
        dropEl.parentElement.classList.add('is-over');
      });
    });
    ['dragleave', 'dragend'].forEach(function (ev) {
      dropEl.addEventListener(ev, function (e) {
        e.preventDefault();
        if (e.relatedTarget && dropEl.contains(e.relatedTarget)) return;
        dropEl.parentElement.classList.remove('is-over');
      });
    });
    dropEl.addEventListener('drop', function (e) {
      e.preventDefault(); e.stopPropagation();
      dropEl.parentElement.classList.remove('is-over');
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) acceptFn(f);
    });
    dropEl.addEventListener('click', function () { inputEl.click(); });
    dropEl.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); inputEl.click(); }
    });
    inputEl.addEventListener('change', function () {
      var f = inputEl.files && inputEl.files[0];
      if (f) acceptFn(f);
      inputEl.value = '';
    });
  }

  // 整个页面也能接住拖进来的文件
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) { e.preventDefault(); });

  bindDrop($('drop-list'), $('input-list'), acceptList);
  bindDrop($('drop-table'), $('input-table'), acceptTable);

  $('list-swap').addEventListener('click', function () { $('input-list').click(); });
  $('table-swap').addEventListener('click', function () { $('input-table').click(); });

  $('sheet-sel').addEventListener('change', function (e) {
    if (!state.table) return;
    loadSheet(e.target.value, true);
  });

  $('opt-header').addEventListener('change', function () {
    state.headerAuto = false;
    rebuildTable();
  });

  ['opt-case', 'opt-zero', 'opt-space'].forEach(function (id) {
    $(id).addEventListener('change', function () {
      if (state.list) renderList();
      refreshIfMatched();
    });
  });

  document.querySelectorAll('input[name="outmode"]').forEach(function (el) {
    el.addEventListener('change', function () {
      $('seg-cap').textContent = SEG_CAP[outMode()] || SEG_CAP.hit;
      refreshIfMatched();
    });
  });

  E.run.addEventListener('click', function () { runMatch(); });
  $('export').addEventListener('click', exportXlsx);
  $('more').addEventListener('click', function () {
    state.previewCount += 200;
    renderPreview(state.result);
  });

  document.addEventListener('keydown', function (e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && ready() && $('run').style.display !== 'none') {
      e.preventDefault();
      runMatch();
    }
  });

  refreshStatus();
})();
