/*
 * 内訳アシスト  taskpane.js（v2）
 * 作業ウィンドウの画面と、Excel（Office.js）への読み書き。
 *  - 入力：品目 → 室 → 仕様（→ 角ダクト・BOX の寸法）を選ぶと、選んでいる行に仕様の行を書く
 *  - 自動入力：内訳の E列にサイズを打つと、保温厚・単位・単価（参照元は Y〜AB列）を入れる
 *  - 選択行に単価：選んでいる行（離れた範囲も可）の単価を入れ直す
 *  - 計・経費：A/B/C列 集計・小計・工事費計・合計・出精値引き・経費
 *  - 単価ルール：直管×係数で出す単価（GV 50A以下＝×1.1 など）を画面で追加・編集
 */
(function () {
  'use strict';
  var U = UA.util;
  var VERSION = '2.0.5';
  var K_MASTER = 'uchiwake-assist.master.v1', K_SK = 'uchiwake-assist.sekisan.v1', K_SPECS = 'uchiwake-assist.speclocal.v1',
    K_UI = 'uchiwake-assist.ui.v2', K_BOX = 'uchiwake-assist.boxui.v1', K_RULES = 'uchiwake-assist.calcrules.v1', K_EXP = 'uchiwake-assist.expense.v1';
  var REF_HEAD = { y: '単価の参照元（保温積算資料）', z: '参照した仕様', aa: '参照したサイズ', ab: '参照した厚み' };
  var TAB = { input: '入力', total: '計・経費', rules: '単価ルール' };

  var st = {
    master: null, masterInfo: null, sk: null, skInfo: null, specLocal: { add: [], hide: {}, pin: {} }, siteType: '',
    sheets: [], target: '', model: null, sc: null, grid: null, headerOK: false, modelErr: '', last: 0,
    tab: 'input',
    wiz: freshWiz(),
    sel: { row: 0, rows: [], areas: 0, address: '', sheet: '' },
    auto: { on: true, log: [] },
    rules: null, rulesLocal: false,
    exp: { cfg: null, mode: 'auto', baseRow: '' },
    ui: { menu: false, busy: false, toast: null, composing: false, specEditMode: false, specEdit: null, showHidden: false, newSpec: '', newCat: '', newPlace: 'here',
      specIdx: [], showAll: false, full: true, discount: '', hover: '', ruleEdit: null, anim: '' },
    undo: null, api19: false, api17: false, autoOpen: false
  };
  function freshWiz() {
    return { open: 'item', item: '', place: '', spec: '', cont: false, d: '', done: null, insN: '',
      rect: { sub: '', t16: false, f: '', fx: '', rq: '' },
      box: { name: '', five: null, f: '', fx: '', show: { w: true, h: true, l: true }, sym: false, boxes: [{ w: '', h: '', l: '', n: '1' }] } };
  }

  /* ---------- 保存（このパソコンの中） ---------- */
  function pkey(k) { try { return (Office.context.partitionKey ? Office.context.partitionKey + '.' : '') + k; } catch (e) { return k; } }
  function load(k) { try { var v = localStorage.getItem(pkey(k)); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
  function save(k, v) { try { localStorage.setItem(pkey(k), JSON.stringify(v)); return true; } catch (e) { return false; } }
  function docGet(k) { try { return Office.context.document.settings.get(k); } catch (e) { return null; } }
  function docSet(k, v) { try { Office.context.document.settings.set(k, v); Office.context.document.settings.saveAsync(function () {}); } catch (e) { /* noop */ } }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function $(q) { return document.querySelector(q); }
  function msg(e) { return (e && (e.message || e.debugInfo && e.debugInfo.message)) || String(e); }
  function saveUi() { save(K_UI, { auto: st.auto.on, full: st.ui.full, tab: st.tab }); }

  /* ---------- 起動 ---------- */
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) { $('#app').innerHTML = '<div class="boot">Excel で開いてください。</div>'; return; }
    try {
      st.api19 = Office.context.requirements.isSetSupported('ExcelApi', '1.9');
      st.api17 = Office.context.requirements.isSetSupported('ExcelApi', '1.7');
    } catch (e) { /* noop */ }
    st.autoOpen = docGet('Office.AutoShowTaskpaneWithDocument') === true;
    var raw = load(K_MASTER); if (raw && raw.master) setMaster(raw.master, raw.info, false);
    var sl = load(K_SPECS); if (sl && sl.add) st.specLocal = { add: sl.add || [], hide: sl.hide || {}, pin: sl.pin || {}, order: sl.order || {} };
    var bu = load(K_BOX); if (bu && bu.show) { st.wiz.box.show = bu.show; st.wiz.box.sym = !!bu.sym; }
    st.siteType = docGet('uchiwake.siteType') || '';
    var skRaw = load(K_SK); if (skRaw && skRaw.sk && skRaw.sk.tables) { st.sk = skRaw.sk; st.skInfo = skRaw.info; }
    var ui = load(K_UI); if (ui) { if (ui.auto === false) st.auto.on = false; if (ui.full === false) st.ui.full = false; if (TAB[ui.tab]) st.tab = ui.tab; }
    var rl = load(K_RULES); if (rl && rl.rules) { st.rules = rl.rules; st.rulesLocal = true; }
    loadExpCfg();
    bindEvents();
    $('#masterFile').addEventListener('change', onMasterFile);
    $('#skFile').addEventListener('change', onSkFile);
    refreshSheets().then(function () { return readSelection(); }).then(render, function (e) { st.modelErr = msg(e); render(); });
    if (st.api17) {
      Excel.run(function (ctx) {
        ctx.workbook.worksheets.onActivated.add(function () { return refreshSheets().then(readSelection).then(render); });
        return ctx.sync();
      }).catch(function () { /* noop */ });
    }
  });

  /* ---------- 設定マスタ・保温積算資料 ---------- */
  function setMaster(m, info, persist) {
    st.master = UA.indexMaster(m);
    st.masterInfo = info;
    if (persist) {
      var copy = Object.assign({}, m); delete copy.priceIdx; delete copy.specCount; delete copy.thickIdx; delete copy.kindById; delete copy.skRules; delete copy.skAdj;
      if (!save(K_MASTER, { master: copy, info: info })) toast('マスタをこのパソコンに記憶できませんでした（容量不足の可能性）。表示中は使えます。', 'warn');
    }
  }
  function onMasterFile(ev) {
    var f = ev.target.files && ev.target.files[0];
    ev.target.value = '';
    if (!f) return;
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        var sheets = {};
        wb.SheetNames.forEach(function (n) { sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: '', raw: true }); });
        var m = UA.parseMaster(sheets);
        if (!m.kinds.length) { toast('設定マスタとして読めませんでした：' + m.errors.join('／'), 'err'); render(); return; }
        setMaster(m, { name: f.name, at: U.today() }, true);
        st.wiz.spec = '';
        toast('設定マスタを読み込みました（' + f.name + '）' + (m.errors.length ? '　注意：' + m.errors.join('／') : ''), m.errors.length ? 'warn' : 'ok');
        loadModel().then(render);
      } catch (err) { toast('読み込みに失敗しました：' + msg(err), 'err'); render(); }
    };
    reader.readAsArrayBuffer(f);
  }
  function onSkFile(ev) {
    var f = ev.target.files && ev.target.files[0];
    ev.target.value = '';
    if (!f) return;
    toast('保温積算資料を読み込んでいます…', 'info'); render();
    var reader = new FileReader();
    reader.onload = function (e) {
      setTimeout(function () {
        try {
          var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', bookFiles: true });
          var sk = UA.parseSekisan(wb, XLSX);
          if (!sk.n) { toast('単価表（テーブル）が見つかりませんでした。保温積算資料のファイルか確認してください。', 'err'); render(); return; }
          var lm = f.name.match(/令和\s*\d+\s*年度版?|R\d+年度版?/);
          sk.label = lm ? lm[0].replace(/\s/g, '') : '';
          st.sk = sk; st.skInfo = { name: f.name, at: U.today(), n: sk.n };
          if (!save(K_SK, { sk: sk, info: st.skInfo })) toast('保温積算資料をこのパソコンに記憶できませんでした（容量不足の可能性）。表示中は使えます。', 'warn');
          else toast('保温積算資料を読み込みました（' + f.name + '、' + sk.n + '表）', 'ok');
          render();
        } catch (err) { toast('読み込みに失敗しました：' + msg(err), 'err'); render(); }
      }, 30);
    };
    reader.readAsArrayBuffer(f);
  }
  function refColsOf() { return (st.master && st.master.output && st.master.output.refCols) || { y: 'Y', z: 'Z', aa: 'AA', ab: 'AB' }; }
  function rules() { return st.rules || (st.master && st.master.calcRulesRaw && st.master.calcRulesRaw.length ? st.master.calcRulesRaw : UA.defaultCalcRules()); }
  function saveRules(list) { st.rules = list; st.rulesLocal = true; if (!save(K_RULES, { rules: list, at: U.today() })) toast('ルールをこのパソコンに記憶できませんでした', 'warn'); }
  function loadExpCfg() {
    var base = JSON.parse(JSON.stringify(UA.EXPENSE_DEFAULT));
    var def = load(K_EXP); if (def && def.cfg) UA.EXP_KEYS.forEach(function (k) { if (def.cfg[k]) base[k] = Object.assign(base[k], def.cfg[k]); });
    var doc = docGet('uchiwake.expense');
    if (doc && doc.cfg) UA.EXP_KEYS.forEach(function (k) { if (doc.cfg[k]) base[k] = Object.assign(base[k], doc.cfg[k]); });
    st.exp.cfg = base; st.exp.mode = (doc && doc.mode) || 'auto';
  }

  /* ---------- シートの読み込み ---------- */
  function isUchiwakeName(n) { return /^内訳/.test(n) && n.indexOf('_積算') < 0 && n.indexOf('FMT') < 0; }
  function refreshSheets() {
    return Excel.run(function (ctx) {
      var wss = ctx.workbook.worksheets; wss.load('items/name,items/visibility');
      var act = wss.getActiveWorksheet(); act.load('name');
      return ctx.sync().then(function () {
        st.sheets = wss.items.filter(function (w) { return w.visibility === 'Visible' && isUchiwakeName(w.name); }).map(function (w) { return w.name; });
        if (st.sheets.indexOf(act.name) >= 0) st.target = act.name;
        else if (st.sheets.indexOf(st.target) < 0) st.target = st.sheets[0] || '';
      });
    }).then(function () { return loadModel(); });
  }
  function colsOf() { return (st.master && st.master.output && st.master.output.cols) || { area: 'A', sys: 'B', place: 'C', d: 'D', spec: 'E', size: 'E', f: 'F', g: 'G', h: 'H', i: 'I', j: 'J', label: 'I' }; }
  function startRowOf() { return (st.master && st.master.output && st.master.output.startRow) || 5; }
  function loadModel() {
    if (!st.target) { st.model = null; st.sc = null; st.headerOK = false; return Promise.resolve(); }
    var target = st.target;
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(target);
      var ur = ws.getUsedRange(false); ur.load('rowIndex,rowCount');
      return ctx.sync().then(function () {
        var last = Math.min(Math.max(ur.rowIndex + ur.rowCount, startRowOf() + 1), 6000);
        var rg = ws.getRange('A1:M' + last); rg.load('values,formulas');
        return ctx.sync().then(function () { setGrid(rg.values, rg.formulas, last); attachEvents(target); });
      });
    }).catch(function (e) { st.model = null; st.sc = null; st.modelErr = msg(e); });
  }
  function setGrid(values, formulas, last) {
    st.grid = { v: values, f: formulas };
    st.last = last;
    var model = UA.parseSheet(values, formulas, colsOf(), startRowOf());
    st.headerOK = String((values[0] || [])[5] || '').replace(/[\s　]/g, '').indexOf('保温厚') >= 0;
    model.jTpl = detectTpl(model, formulas);
    st.model = model; st.modelErr = '';
    st.sc = st.master ? UA.scan(values, formulas, st.master, startRowOf()) : null;
  }
  // 既存の金額式（J列）のくせを、行番号を {行} にして覚える
  function detectTpl(model, formulas) {
    var cnt = { line: {}, blank: {} };
    model.recs.forEach(function (r) {
      if (r.t !== 'line' && r.t !== 'blank') return;
      var f = (formulas[r.row - 1] || [])[9];
      if (typeof f !== 'string' || f.charAt(0) !== '=') return;
      var tpl = f.replace(new RegExp('(\\$?[A-Z]{1,2}\\$?)' + r.row + '(?!\\d)', 'g'), '$1{行}');
      if (tpl.indexOf('{行}') < 0) return;
      var c = cnt[r.t]; c[tpl] = (c[tpl] || 0) + 1;
    });
    function best(o) { var b = null, n = 0; Object.keys(o).forEach(function (k) { if (o[k] > n) { n = o[k]; b = k; } }); return b; }
    var dflt = (st.master && st.master.output && st.master.output.amountTpl) || '=IF(AND(G{行}<>"",I{行}=""),0,IF(AND(G{行}="",I{行}=""),"",IF(AND(G{行}<>"",I{行}<>""),G{行}*I{行},"")))';
    var line = best(cnt.line) || dflt;
    return { line: line, blank: best(cnt.blank) || line };
  }
  function jTpl(r) { var t = st.model && st.model.jTpl ? st.model.jTpl.line : ''; return t ? t.replace(/\{行\}/g, r) : ''; }

  /* ---------- Excel のイベント（入力の監視・選択の追従） ---------- */
  var handlers = { sheet: null, changed: null, selection: null };
  var changedTimer = null, selTimer = null, pending = { rows: {}, struct: false };
  var own = { until: 0, allUntil: 0, struct: 0, rows: {}, addrs: {} };   // 自分が書いた変更のイベントは無視する
  function attachEvents(target) {
    if (!st.api17 || handlers.sheet === target) return;
    var old = [handlers.changed, handlers.selection].filter(Boolean);
    handlers.sheet = target;
    var rm = old.length ? Excel.run(old[0].context, function (ctx) { old.forEach(function (h) { h.remove(); }); return ctx.sync(); }).catch(function () { /* noop */ }) : Promise.resolve();
    rm.then(function () {
      return Excel.run(function (ctx) {
        var ws = ctx.workbook.worksheets.getItem(target);
        handlers.changed = ws.onChanged.add(onSheetChanged);
        handlers.selection = ws.onSelectionChanged.add(function () { clearTimeout(selTimer); selTimer = setTimeout(function () { readSelection().then(renderSoft); }, 140); return Promise.resolve(); });
        return ctx.sync();
      });
    }).catch(function () { handlers.sheet = null; });
  }
  // 「E26」「E26:E30」「A5:J5」→ { c1, c2, r1, r2 }
  function parseAddr(a) {
    var s = String(a || '').replace(/\$/g, '').split('!').pop().split(',')[0];
    var m = s.match(/^([A-Z]+)?(\d+)?(?::([A-Z]+)?(\d+)?)?$/i);
    if (!m) return null;
    var c1 = m[1] ? U.colNum(m[1]) : 1, r1 = m[2] ? +m[2] : 1, c2 = m[3] ? U.colNum(m[3]) : (m[1] ? c1 : 16384), r2 = m[4] ? +m[4] : (m[2] ? r1 : 1048576);
    if (m[1] && !m[3] && !m[4]) c2 = c1;
    return { c1: c1, c2: c2, r1: r1, r2: r2 };
  }
  function onSheetChanged(ev) {
    if (Date.now() < own.allUntil) return Promise.resolve();
    if (ev && ev.source && ev.source !== 'Local') return Promise.resolve();
    // 自分（アドイン）が書いたセルの変更は無視する（入力フォーム・計・経費で書いたセル）。
    // 自動入力で入れた保温厚（F列）は登録しない：前は、厚みを自動で入れた直後に操作者が厚みを
    // 打ち直すと自分の書き込みと取り違えて無視し、単価が前の厚みのまま残ることがあった
    var ak = String(ev && ev.address || '').replace(/\$/g, '').split('!').pop();
    if (own.addrs[ak] && own.addrs[ak] > Date.now()) { delete own.addrs[ak]; return Promise.resolve(); }
    var type = ev && ev.changeType;
    if (type && type !== 'RangeEdited') { pending.struct = true; st.auto.memo = {}; if (Date.now() > own.struct) st.undo = null; schedule(); return Promise.resolve(); }
    var a = parseAddr(ev && ev.address);
    if (a && st.auto.on && st.master) {
      var hitsDE = a.c1 <= 5 && a.c2 >= 4, hitsF = a.c1 <= 6 && a.c2 >= 6;
      var rows = a.r2 - a.r1 + 1;
      if ((hitsDE || hitsF) && rows <= 400) {
        for (var r = a.r1; r <= a.r2; r++) pending.rows[r] = true;
      } else if ((hitsDE || hitsF) && rows > 400) toast('大きな範囲の変更は自動入力しません。単価は「選択行に単価」で入れられます。', 'info');
    }
    schedule();
    return Promise.resolve();
  }
  function schedule() {
    clearTimeout(changedTimer);
    changedTimer = setTimeout(flushPending, 160);
  }
  function flushPending() {
    if (st.ui.busy) { schedule(); return; }
    var rows = Object.keys(pending.rows).map(Number).sort(function (a, b) { return a - b; });
    pending.rows = {}; pending.struct = false;
    if (rows.length && st.auto.on) fillRows(rows, 'auto').then(render);
    else loadModel().then(renderSoft);
  }
  function readSelection() {
    if (!st.target) return Promise.resolve();
    return Excel.run(function (ctx) {
      var one = ctx.workbook.getSelectedRange(); one.load('rowIndex,rowCount,address');
      var ws = one.worksheet; ws.load('name');
      var many = st.api19 ? ctx.workbook.getSelectedRanges() : null;
      if (many) { many.load('address'); many.areas.load('items/rowIndex,items/rowCount'); }
      return ctx.sync().then(function () {
        var areas = many ? many.areas.items.map(function (x) { return { r1: x.rowIndex + 1, n: x.rowCount }; }) : [{ r1: one.rowIndex + 1, n: one.rowCount }];
        var rows = {}, lim = st.last || 6000;
        areas.forEach(function (x) { for (var r = x.r1; r < x.r1 + x.n && r <= lim; r++) rows[r] = true; });
        st.sel = { row: one.rowIndex + 1, rows: Object.keys(rows).map(Number).sort(function (a, b) { return a - b; }), areas: areas.length, address: many ? many.address : one.address, sheet: ws.name };
      });
    }).catch(function () { /* noop */ });
  }

  /* ---------- 自動入力・選択行に単価 ---------- */
  function guessFor(sc, r, aa) {
    var o = sc.rows[r];
    if (!o || o.kind !== 'line') return { skip: '明細の行ではありません' };
    var x = UA.ctxAt(sc, r, o.D);
    if (!x.spec) return { skip: '上に仕様の行（括弧書き）がありません' };
    var item = o.item, sz = o.sz, fNow = U.isEmpty(o.F) ? '' : U.thickKey(o.F), g = null, five;
    // 入力フォームで書いた BOX（E列に一部の辺だけ出した行）：AA列に残した3辺と面の数で計算する
    var am = typeof aa === 'string' ? aa.match(/^\s*([^（(]+)/) : null, ab = am ? UA.parseSize(am[1].trim()) : null;
    if (ab && ab.box && (String(o.H).trim() === '台' || item === 'box')) { item = 'box'; sz = ab; five = /5面/.test(aa) ? true : /6面/.test(aa) ? false : undefined; }
    else if (item === 'rect' && sz.wh && String(o.H).trim() === '台') return { skip: 'BOX の寸法が2辺だけです（入力フォームで書いてください）' };
    if (!fNow) g = UA.guessThick(sc, st.master, x, item, sz, r);
    var f = fNow || (g ? g.f : '');
    var h = U.isEmpty(o.H) ? UA.guessUnit(sc, st.master, x, item, r) : null;
    var pr = UA.priceLine(st.sk, st.master, rules(), { item: item, d: x.d, five: five, place: x.place, spec: x.spec, sz: sz, f: f },
      { findBase: function (b) { return UA.findBaseLine(sc, x, r, b); } });
    return { o: o, x: x, item: item, sz: sz, f: f, fNew: fNow ? null : (g && g.f ? g.f : null), fFrom: g ? g.from : '', h: h, pr: pr };
  }
  // mode：'auto'（空欄と自動で入れた単価だけ）／'force'（選んだ行の単価を入れ直す）
  function fillRows(rows, mode) {
    if (!st.target || !st.master) return Promise.resolve();
    var target = st.target, rc = refColsOf(), rk = ['y', 'z', 'aa', 'ab'].filter(function (k) { return rc[k]; });
    var res = { done: [], skip: {}, manual: 0, manualRows: [], none: 0 };
    st.ui.busy = true;
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(target);
      var ur = ws.getUsedRange(false); ur.load('rowIndex,rowCount');
      return ctx.sync().then(function () {
        var last = Math.min(Math.max(ur.rowIndex + ur.rowCount, startRowOf() + 1, rows[rows.length - 1] || 0), 6000);
        var rg = ws.getRange('A1:M' + last); rg.load('values,formulas');
        var r1 = rows[0], r2 = rows[rows.length - 1];
        var ks = rk.map(function (k) { return U.colNum(rc[k]); });
        var refRg = rk.length ? ws.getRange(U.colLetter(Math.min.apply(null, ks)) + r1 + ':' + U.colLetter(Math.max.apply(null, ks)) + r2) : null; if (refRg) refRg.load('values,formulas');
        var hdr = rk.map(function (k) { var c = ws.getRange(rc[k] + '1'); c.load('values'); return { k: k, c: c }; });
        return ctx.sync().then(function () {
          setGrid(rg.values, rg.formulas, last);
          var sc = st.sc, undo = { sheet: target, cells: [], kind: 'auto' }, wrote = {}, anyRef = false;
          function put(addr, v, prev) { undo.cells.push({ addr: addr, v: prev == null ? '' : prev }); ws.getRange(addr).formulas = [[v]]; if (!/^[DEF]\d/.test(addr)) own.addrs[addr] = Date.now() + 4000; }
          var c0 = rk.length ? Math.min.apply(null, ks) : 0;
          function off(i) { return U.colNum(rc[rk[i]]) - c0; }
          function refPrev(r, i) { if (!refRg) return ''; var row = refRg.formulas[r - r1]; return row ? row[off(i)] : ''; }
          function refVal(r, i) { if (!refRg) return ''; var row = refRg.values[r - r1]; return row ? row[off(i)] : ''; }
          rows.forEach(function (r) {
            if (r < startRowOf() || r > last) return;
            var o = sc.rows[r], yNow = refVal(r, 0), owned = UA.isAutoRef(yNow);
            var fv = rg.formulas[r - 1] || [];
            // サイズを消した行：自動で入れた単価・参照元を消す
            if (mode === 'auto' && (!o || o.kind !== 'line')) {
              if (o && U.isEmpty(o.E) && owned) {
                var memo = st.auto.memo && st.auto.memo[r];
                if (!U.isEmpty(fv[8])) put('I' + r, '', fv[8]);
                rk.forEach(function (k, i) { put(rc[k] + r, '', refPrev(r, i)); });
                if (memo && memo.f && U.thickKey(o.F) === memo.f) put('F' + r, '', fv[5]);
                if (memo && memo.h && String(o.H).trim() === memo.h) put('H' + r, '', fv[7]);
                wrote[r] = true;
                res.done.push({ r: r, cleared: true });
              }
              return;
            }
            var gu = guessFor(sc, r, refVal(r, rk.indexOf('aa')));
            if (gu.skip) {
              res.skip[gu.skip] = (res.skip[gu.skip] || 0) + 1;
              // 自動で入れた単価が残っていて、いまは単価を決められない行：前の単価を残さず空白にする
              if (mode === 'auto' && owned && !U.isEmpty(fv[8])) {
                put('I' + r, '', fv[8]);
                rk.forEach(function (k, i) { var nv = k === 'y' ? '単価なし：' + gu.skip : ''; if (String(refPrev(r, i)) !== nv) put(rc[k] + r, nv, refPrev(r, i)); });
                wrote[r] = true;
                res.done.push({ r: r, size: '', d: '', item: '', f: '', h: '', p: null, kept: false, why: gu.skip, y: '', changed: true });
              }
              return;
            }
            var n0 = undo.cells.length;
            var iEmpty = U.isEmpty(fv[8]);
            var canI = mode === 'force' || iEmpty || owned;
            if (gu.fNew) put('F' + r, numOrText(gu.fNew), fv[5]);
            if (gu.h) put('H' + r, gu.h, fv[7]);
            if (canI) {
              var iv = gu.pr.p == null ? '' : (gu.pr.formula || gu.pr.p);
              if (!(iEmpty && iv === '') && String(fv[8]) !== String(iv)) put('I' + r, iv, fv[8]);
              rk.forEach(function (k, i) { var nv = gu.pr.ref && gu.pr.ref[k] != null ? gu.pr.ref[k] : ''; if (String(refPrev(r, i)) !== String(nv)) put(rc[k] + r, nv, refPrev(r, i)); });
              anyRef = true;
            } else { res.manual++; res.manualRows.push(r); }
            if (U.isEmpty(fv[9])) put('J' + r, jTpl(r), fv[9]);
            if (gu.pr.p == null) res.none++;
            // 自動で入れた保温厚・単位の記録（サイズを消したときに一緒に消すため）。操作者が直したら記録を外す
            st.auto.memo = st.auto.memo || {};
            var mm = st.auto.memo[r] || {};
            if (gu.fNew) mm.f = U.thickKey(gu.fNew); else if (mm.f && U.thickKey(gu.o.F) !== mm.f) mm.f = null;
            if (gu.h) mm.h = gu.h; else if (mm.h && String(gu.o.H).trim() !== mm.h) mm.h = null;
            st.auto.memo[r] = mm;
            wrote[r] = true;
            res.done.push({ r: r, size: UA.sizeLabel(gu.item, gu.sz), d: gu.x.d, item: gu.item, f: gu.f, fNew: !!gu.fNew, fFrom: gu.fFrom, h: gu.h || String(gu.o.H || ''), p: canI ? gu.pr.p : null, kept: !canI, why: gu.pr.why || '', y: gu.pr.ref ? gu.pr.ref.y : '', via: gu.pr.via, formula: gu.pr.formula, changed: undo.cells.length > n0 });
          });
          if (anyRef) hdr.forEach(function (h) { if (U.isEmpty(h.c.values[0][0])) { undo.cells.push({ addr: rc[h.k] + '1', v: '' }); h.c.values = [[REF_HEAD[h.k]]]; } });
          return ctx.sync().then(function () { if (undo.cells.length) st.undo = undo; });
        });
      });
    }).then(function () {
      st.ui.busy = false;
      // 自動入力では、実際に書き換えた行だけを記録する（自分の書き込みの通知で動いたときは何も出さない）
      var logs = res.done.filter(function (d) { return !d.cleared && (mode === 'force' || d.changed); });
      if (logs.length) { st.auto.log = logs.slice(-30).reverse().concat(st.auto.log).slice(0, 40); }
      if (mode === 'force') {
        var t = logs.length ? logs.length + '行に単価を入れました' : '単価を入れる行がありませんでした';
        if (res.none) t += '（協会単価なし ' + res.none + '行）';
        var sk = Object.keys(res.skip); if (sk.length) t += '。入れなかった行：' + sk.map(function (k) { return k + ' ' + res.skip[k] + '行'; }).join('、');
        toast(t, logs.length ? (res.none ? 'warn' : 'ok') : 'warn', logs.length > 0);
      } else if (logs.length === 1 && logs[0].p == null && !logs[0].kept) {
        toast(logs[0].r + '行目：' + (logs[0].why || '協会単価なし') + '（単価は空白にしました）', 'warn', true);
      } else if (res.manualRows.length && !logs.some(function (d) { return !d.kept; })) {
        // 手で入れた単価は変えない。厚み・サイズを変えたときに前の単価が残るので知らせる
        var mr = res.manualRows, rn = mr.slice(0, 3).join('・') + '行目' + (mr.length > 3 ? 'ほか' + (mr.length - 3) + '行' : '');
        toast(rn + 'は手で入れた単価なので変えていません。厚み・サイズに合っているか確かめてください（入れ直すときは「選択行に単価」）', 'warn');
      }
      return loadModel();
    }).catch(function (e) {
      st.ui.busy = false;
      toast('自動入力できませんでした：' + msg(e) + (/protect|保護/i.test(msg(e)) ? '（シートの保護を解除してください）' : ''), 'err');
    });
  }
  function numOrText(v) { var t = U.toHalf(v); return /^\d+(\.\d+)?$/.test(t) ? parseFloat(t) : v; }
  function priceSelected() {
    if (!st.sel.rows.length) { toast('内訳で単価を入れたい行を選んでください（Ctrl を押しながら離れた範囲も選べます）', 'warn'); render(); return; }
    if (st.sel.sheet && st.sel.sheet !== st.target) { toast('「' + st.target + '」のシートで行を選んでください', 'warn'); render(); return; }
    fillRows(st.sel.rows.slice(0, 1500), 'force').then(render);
  }

  /* ---------- 書き込み（入力フォーム・計・経費） ---------- */
  // 書き込む直前に、シートと選択を読み直してから計画を作る（古い内容で書かないように）
  function fresh(fn) {
    if (st.ui.busy) return;
    loadModel().then(readSelection).then(function () {
      if (st.sel.sheet && st.sel.sheet !== st.target) { toast('「' + st.target + '」のシートを開いて、行を選んでから押してください（今は「' + st.sel.sheet + '」）', 'warn'); render(); return; }
      fn();
    });
  }
  // plan = { ins:{at,n}|null, writes:[{r,col,v,diag,align,fmt}], select:'E27', label }
  var FMT_SHEET = '内訳(FMT)', FMT_ROW = 5, fmtExists = null;
  // 内訳(FMT) が無いブックのとき：このシートの下のほうの、何も入っていない行（初めの書式のまま）
  function blankTemplateRow() {
    if (!st.grid) return null;
    for (var r = st.grid.v.length; r > startRowOf(); r--) { var row = st.grid.v[r - 1]; if (row && row.slice(0, 9).every(U.isEmpty)) return r; }
    return null;
  }
  function fmtSrc(fr, ins) {
    if (fmtExists) return "'" + FMT_SHEET + "'!A" + FMT_ROW + ':J' + FMT_ROW;
    if (!fr) return null;
    var r = fr >= ins.at ? fr + ins.n : fr;   // 挿入でずれた位置
    return 'A' + r + ':J' + r;
  }
  function applyPlan(plan, doneText) {
    if (st.ui.busy) return Promise.resolve(false);
    if (!st.headerOK) { toast('書き込み先が内訳シートではないようです（1行目に「保温厚」の見出しがありません）', 'err'); render(); return Promise.resolve(false); }
    st.ui.busy = true; render();
    var target = st.target, ins = plan.ins && plan.ins.n > 0 ? plan.ins : null;
    function orig(r) { return !ins || r < ins.at ? r : (r < ins.at + ins.n ? null : r - ins.n); }
    var undo = { sheet: target, cells: [], diags: [], ins: ins, kind: 'plan' };
    var until = Date.now() + 4000;
    plan.writes.forEach(function (w) { own.addrs[w.col + w.r] = until; });
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(target);
      var before = plan.writes.map(function (w) {
        var o = orig(w.r); if (o == null) return null;
        var c = ws.getRange(w.col + o); c.load('formulas');
        var b = w.diag ? c.format.borders.getItem('DiagonalUp') : null; if (b) b.load('style');
        return { w: w, o: o, c: c, b: b };
      });
      // 挿入した行の書式の元：内訳(FMT)（ひな型の初めの書式）。無ければこのシートの下の空き行
      var fmt = ins ? ctx.workbook.worksheets.getItemOrNullObject(FMT_SHEET) : null;
      if (fmt) fmt.load('name');
      return ctx.sync().then(function () {
        var srcRg = null;
        if (fmt) fmtExists = !fmt.isNullObject;
        if (ins && st.api19) {
          if (fmtExists) srcRg = fmt.getRange('A' + FMT_ROW + ':J' + FMT_ROW);
          else { var fr = blankTemplateRow(); if (fr) srcRg = ws.getRange('A' + fr + ':J' + fr); }
          if (srcRg) srcRg.format.load('rowHeight');
        }
        return ctx.sync().then(function () { return srcRg ? srcRg.format.rowHeight : null; });
      }).then(function (rowH) {
        var src = ins && st.api19 ? fmtSrc(blankTemplateRow(), ins) : null;
        before.forEach(function (x) {
          if (!x) return;
          undo.cells.push({ addr: x.w.col + x.o, v: x.c.formulas[0][0] });
          if (x.b) undo.diags.push({ addr: x.w.col + x.o, style: x.b.style });
        });
        if (ins) {
          own.struct = until; ws.getRange(ins.at + ':' + (ins.at + ins.n - 1)).insert('Down');
          // 挿入した行：A〜J列を内訳の初めの書式にし、J列に金額式を入れる（このあと書く行は上書き）
          for (var ir = ins.at; ir < ins.at + ins.n; ir++) {
            var dst = ws.getRange('A' + ir + ':J' + ir);
            if (src) dst.copyFrom(src, 'Formats');
            ws.getRange('J' + ir).formulas = [[jTpl(ir)]];
            own.addrs['J' + ir] = until;
          }
          if (rowH) ws.getRange(ins.at + ':' + (ins.at + ins.n - 1)).format.rowHeight = rowH;
        }
        plan.writes.forEach(function (w) {
          var c = ws.getRange(w.col + w.r);
          c.formulas = [[w.v == null ? '' : w.v]];
          if (w.diag) { var b = c.format.borders.getItem('DiagonalUp'); b.style = 'Continuous'; b.weight = 'Thin'; b.color = '#000000'; }
          if (w.align) c.format.horizontalAlignment = w.align;
          if (w.fmt) c.numberFormat = [[w.fmt]];
        });
        // 書き込んだシートを表示して（ほかのシートを見ていても）、続けて打つセルを選ぶ
        ws.activate();
        if (plan.select) ws.getRange(plan.select).select();
        return ctx.sync();
      });
    }).then(function () {
      st.undo = undo;
      st.ui.busy = false;
      if (doneText) toast(doneText, 'ok', true);
      return loadModel().then(readSelection).then(function () { return true; });
    }).catch(function (e) {
      st.ui.busy = false;
      toast('書き込めませんでした：' + msg(e) + (/protect|保護/i.test(msg(e)) ? '（シートの保護を解除してください）' : ''), 'err');
      render();
      return false;
    });
  }
  function undoLast() {
    var u = st.undo;
    if (!u || st.ui.busy) return;
    st.ui.busy = true; render();
    own.allUntil = Date.now() + 60000; own.struct = Date.now() + 4000;
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(u.sheet);
      if (u.ins) ws.getRange(u.ins.at + ':' + (u.ins.at + u.ins.n - 1)).delete('Up');
      return ctx.sync().then(function () {
        u.cells.slice().reverse().forEach(function (c) { ws.getRange(c.addr).formulas = [[c.v == null ? '' : c.v]]; });
        (u.diags || []).forEach(function (d) { ws.getRange(d.addr).format.borders.getItem('DiagonalUp').style = d.style || 'None'; });
        ws.activate();
        return ctx.sync();
      });
    }).then(function () { st.undo = null; own.allUntil = Date.now() + 900; st.ui.busy = false; return loadModel(); })
      .then(function () { toast('元に戻しました', 'ok'); })
      .catch(function (e) { own.allUntil = 0; st.ui.busy = false; toast('元に戻せませんでした：' + msg(e), 'err'); })
      .then(render);
  }

  /* ---------- 入力フォーム ---------- */
  function kinds() { return st.master ? st.master.kinds.filter(function (k) { return k.use; }) : []; }
  function kindOf(id) { return st.master && st.master.kindById[id]; }
  function subsOf(item) { return st.master ? (st.master.subs[item] || []).filter(function (s) { return s.use; }) : []; }
  function famOfItem(item) { var k = kindOf(item); return k ? k.fam : (item === 'rect' || item === 'round' || item === 'box' ? 'duct' : 'pipe'); }
  function sheetPlaces() {
    var out = [];
    if (!st.sc) return out;
    st.sc.rows.forEach(function (o) { if (o && (o.kind === 'spec' || o.kind === 'sys') && o.C && /^[\s　]*[（(]/.test(String(o.C))) { var p = U.trimAll(U.stripParen(o.C)); if (p && out.indexOf(p) < 0) out.push(p); } });
    return out;
  }
  function placeList() {
    var master = st.master ? st.master.places.filter(function (p) { return p.use; }).map(function (p) { return p.name; }) : [];
    var sh = sheetPlaces();
    return { sheet: sh, all: master.filter(function (p) { return sh.indexOf(p) < 0; }) };
  }
  // 書き込む位置（Excel で選んでいる行）
  function targetRow() { return st.sel && st.sel.sheet === st.target && st.sel.row ? st.sel.row : 0; }
  function ctxNow() { var R = targetRow(); if (!st.sc || !R) return null; var P = st.sc.rows[R] && !st.sc.rows[R].blank ? R + 1 : R; return UA.ctxAt(st.sc, P); }
  function specGroups() {
    var w = st.wiz, cx = ctxNow();
    return UA.specList(st.master, { item: w.item, place: w.place, fam: famOfItem(w.item), sheet: st.model ? UA.sheetSpecs(st.model) : [], cur: cx && U.trimAll(cx.place) === U.trimAll(w.place) ? cx.spec : null,
      cat: st.siteType, local: st.specLocal, showHidden: st.ui.showHidden });
  }
  function wizPlan() {
    var w = st.wiz, R = targetRow();
    if (!st.sc || !R || !w.spec) return null;
    return UA.specWritePlan(st.sc, R, { place: w.place, spec: w.spec, lines: extraLines().lines, insN: w.insN });
  }
  // 角ダクト・BOX の保温厚：ボタン（25・50）か、操作者が入れた厚み
  function ductThick(o, guess) { return U.thickKey(o.fx) || o.f || guess || '25'; }
  // BOX の寸法の書き方（E列）：出す辺だけ、記号付き／なし
  function boxText(b, bx) {
    var parts = [];
    [['w', 'W'], ['h', 'H'], ['l', 'L']].forEach(function (k) { if (bx.show[k[0]] && U.num(b[k[0]]) > 0) parts.push((bx.sym ? k[1] : '') + U.num(b[k[0]])); });
    return parts.join('×');
  }
  // 角ダクト・BOX：書き込む明細
  function extraLines() {
    var w = st.wiz, out = { lines: [], head: '' };
    if (!st.master || !w.spec) return out;
    var R = targetRow(), x = st.sc && R ? UA.ctxAt(st.sc, R) : { blockId: -1, sysId: 0 };
    var cont = x.spec === U.normSpec(w.spec) && U.trimAll(x.place) === U.trimAll(w.place);
    var xx = { spec: U.normSpec(w.spec), place: w.place, blockId: cont ? x.blockId : -1, sysId: x.sysId, d: '' };
    if (w.item === 'rect') {
      var rt = w.rect, g = U.num(rt.rq);
      var d = rt.sub ? rt.sub + (rt.t16 ? '(1.6t)' : '') : '';
      xx.d = d;
      var ft = ductThick(rt, st.sc ? UA.guessThick(st.sc, st.master, xx, 'rect', { rect: true, text: '矩形' }, -1).f : '');
      var pr = UA.priceLine(st.sk, st.master, rules(), { item: 'rect', d: d, place: w.place, spec: xx.spec, sz: { rect: true, text: '矩形' }, f: ft }, null);
      out.f = ft; out.pr = pr; out.g = g; out.d = d;
      out.lines.push({ d: d, e: '矩形', f: ft, g: g > 0 ? g : '', h: '㎡', pr: pr });
    } else if (w.item === 'box') {
      var bx = w.box, name = String(bx.name || '').trim(), five = bx.five == null ? UA.isFiveSided(name) : bx.five;
      xx.d = name;
      var fb = ductThick(bx, st.sc ? UA.guessThick(st.sc, st.master, xx, 'box', { box: { w: 1, d: 1, h: 1 }, text: '' }, -1).f : '');
      out.f = fb; out.five = five;
      // D列の名前は1行目に単独で書き、その下の行の E列に寸法（上の行と同じ名前なら書かない）
      out.head = name && !(cont && x.prevKind && x.prevKind !== 'blank' && U.trimAll(x.d) === U.trimAll(name)) ? name : '';
      if (out.head) out.lines.push({ head: true, d: name });
      bx.boxes.forEach(function (b) {
        var ww = U.num(b.w), hh = U.num(b.h), ll = U.num(b.l), n = U.num(b.n) || 1;
        if (!(ww > 0 && hh > 0 && ll > 0)) return;
        var sz = { box: { w: ww, d: ll, h: hh }, text: 'W' + ww + '×H' + hh + '×L' + ll };
        var p = UA.priceLine(st.sk, st.master, rules(), { item: 'box', d: name, five: five, place: w.place, spec: xx.spec, sz: sz, f: fb }, null);
        var area = UA.boxArea(sz.box, five);
        // AA列には3辺と面積を残す（E列に一部の辺しか出さなくても、単価を入れ直せるように）
        if (p.ref) p.ref = Object.assign({}, p.ref, { aa: sz.text + '（' + area + '㎡・' + (five ? '5面' : '6面') + '）' });
        out.lines.push({ d: '', e: boxText(b, bx) || sz.text, f: fb, g: n, h: '台', pr: p, area: area });
      });
      if (out.lines.length === (out.head ? 1 : 0)) out.lines = [];
    }
    return out;
  }
  function writeWizard() {
    var w = st.wiz, p = wizPlan(), ex = extraLines();
    if (!p) { toast(!targetRow() ? '内訳シートで書き込みたい行を選んでください' : '仕様を選んでください', 'warn'); render(); return; }
    if (!p.ok) { toast(p.message, 'warn'); render(); return; }
    if (w.item === 'box' && !ex.lines.length) { toast('BOX の寸法（W・H・L）を入れてください', 'warn'); render(); return; }
    var writes = [], rc = refColsOf(), rk = ['y', 'z', 'aa', 'ab'].filter(function (k) { return rc[k]; }), r = p.P;
    if (p.writeSpec) {
      writes.push({ r: r, col: 'C', v: p.writePlace ? '(' + w.place + ')' : '' });
      writes.push({ r: r, col: 'E', v: '(' + U.normSpec(w.spec) + ')', align: 'Left' });
      writes.push({ r: r, col: 'J', v: '' });
      r++;
    }
    var first = r;
    if (ex.lines.length) {
      var prevD = p.ctx && !p.writeSpec ? p.ctx.d : '';
      ex.lines.forEach(function (l) {
        if (l.head) { writes.push({ r: r, col: 'D', v: l.d, align: 'Center' }, { r: r, col: 'J', v: jTpl(r) }); prevD = l.d; r++; return; }
        var dShow = l.d && l.d !== prevD ? l.d : '';
        if (l.d) prevD = l.d;
        writes.push({ r: r, col: 'D', v: dShow, align: 'Center' }, { r: r, col: 'E', v: l.e, align: 'Right' }, { r: r, col: 'F', v: numOrText(l.f) }, { r: r, col: 'G', v: l.g }, { r: r, col: 'H', v: l.h });
        writes.push({ r: r, col: 'I', v: l.pr.p == null ? '' : (l.pr.formula || l.pr.p) }, { r: r, col: 'J', v: jTpl(r) });
        rk.forEach(function (k) { writes.push({ r: r, col: rc[k], v: l.pr.ref && l.pr.ref[k] != null ? l.pr.ref[k] : '' }); });
        r++;
      });
    } else {
      if (w.d) writes.push({ r: r, col: 'D', v: w.d, align: 'Center' });
      writes.push({ r: r, col: 'J', v: jTpl(r) });
    }
    // 角ダクトで数量（㎡）が空なら G列を選ぶ。BOX は次の行
    var sel = w.item === 'rect' && !(U.num(w.rect.rq) > 0) ? 'G' + first : ex.lines.length ? 'D' + r : 'E' + p.inputRow;
    var t = (p.ins ? p.ins + '行挿入して、' : '') + (p.writeSpec ? p.P + '行目に仕様の行を書きました。' : '');
    t += ex.lines.length ? first + (r - 1 > first ? '〜' + (r - 1) : '') + '行目に明細を書きました' + (w.item === 'rect' && !(U.num(w.rect.rq) > 0) ? '。数量（㎡）は G' + first + ' に入れてください' : '') : 'E' + p.inputRow + ' からサイズを打つと、保温厚・単位・単価が自動で入ります';
    applyPlan({ ins: p.ins ? { at: p.insAt, n: p.ins } : null, writes: writes, select: sel }, t).then(function (ok) {
      if (!ok) return;
      w.done = { row: p.P, input: p.inputRow, spec: p.writeSpec, lines: ex.lines.length };
      w.insN = '';
      if (w.item === 'rect') w.rect.rq = '';
      if (w.item === 'box') w.box.boxes = [{ w: '', h: '', l: '', n: '1' }];
      w.open = (w.item === 'rect' || w.item === 'box') ? 'extra' : null;
      render();
    });
  }

  /* ---------- 計・経費 ---------- */
  var TOTALS = [
    { id: 'rA', kind: 'range', col: 'A', grp: '集計', label: 'A列', tip: '大項目の見出しから下の明細を合計' },
    { id: 'rB', kind: 'range', col: 'B', grp: '集計', label: 'B列', tip: '系統の見出しから下の明細を合計' },
    { id: 'rC', kind: 'range', col: 'C', grp: '集計', label: 'C列', tip: '室（施工箇所）の見出しから下の明細を合計' },
    { id: 'sA', kind: 'sub', col: 'A', grp: '小計', label: 'A列', tip: '大項目の見出しの下にある計（集計）を合計' },
    { id: 'sB', kind: 'sub', col: 'B', grp: '小計', label: 'B列', tip: '系統の見出しの下にある計（集計）を合計' },
    { id: 'grand', kind: 'grand', grp: '計', label: '工事費 計', tip: '小計を入れた所は小計を、集計だけの所は集計を合計（小計に入っている集計は足さない）' },
    { id: 'total', kind: 'total', grp: '計', label: '合計', tip: '工事費計から下（経費まで）を合計' }
  ];
  function totalPlanOf(def) {
    if (!st.grid) return null;
    var R = targetRow();
    if (!R) return { ok: false, message: '内訳シートで行を選んでください' };
    return UA.totalPlan(st.grid, R, def.kind, { col: def.col, full: st.ui.full, amount: st.ui.discount });
  }
  function totalWrites(p) {
    return p.writes.map(function (w) { var o = { r: w.r, col: w.col, v: w.v }; if (w.diag) o.diag = true; if (w.col === 'I' && w.v) o.align = 'Right'; return o; });
  }
  function runTotal(def) {
    var p = totalPlanOf(def);
    if (!p || !p.ok) { toast(p ? p.message : '準備ができていません', 'warn'); render(); return; }
    applyPlan({ ins: p.ins ? { at: p.insAt, n: p.ins } : null, writes: totalWrites(p), select: 'A' + p.select },
      p.totalRow + '行目に「' + p.label + '」' + p.formula + ' を入れました' + (p.warn ? '（' + p.warn + '）' : '')).then(render);
  }
  function suggestTotal() {
    var R = targetRow();
    if (!st.grid || !R) return null;
    var v = st.grid.v;
    function t(r, c) { var x = (v[r - 1] || [])[c]; return typeof x === 'string' && x.trim() && !/^\s*\d+(\.\d+)?\s*$/.test(x) ? x : ''; }
    function sum(r) { var f = (st.grid.f[r - 1] || [])[9]; return typeof f === 'string' && /^=SUM\(/i.test(f.replace(/\s/g, '')); }
    var info = UA.expenseInfo(st.grid);
    if (info.grand && R > info.grand) return info.exp && R > info.exp ? (info.total ? null : 'total') : 'exp';
    for (var i = R; i >= 5; i--) {
      var a = t(i, 0);
      if (a) return UA.isBoundA(a) ? null : 'rA';
      if (t(i, 1)) return 'rB';
      if (sum(i)) {
        // すぐ上の計が大項目の計なら次は工事費計、系統の計なら大項目の小計
        var lab = UA.headKey((v[i - 1] || [])[8]);
        for (var k = i - 1; k >= 5; k--) { var ah = t(k, 0); if (ah) return UA.headKey(ah) === lab || (ah.match(/^\s*(\d+)\s*[.．]/) || [])[1] === lab ? 'grand' : 'sA'; }
        return 'sA';
      }
    }
    return null;
  }
  function expView() {
    var info = st.grid ? UA.expenseInfo(st.grid) : null, cfg = st.exp.cfg;
    var zaiko = st.exp.mode === 'auto' ? !!(info && info.zaiko) : st.exp.mode === 'zaiko';
    var base = info && info.grand ? info.grand : null, baseV = base ? U.num((st.grid.v[base - 1] || [])[9]) : 0;
    var plan = st.grid ? UA.expensePlan(st.grid, cfg, { mode: zaiko ? 'zaiko' : 'kyokai' }) : null;
    var laborV = 0;
    if (zaiko && plan && plan.ok) {
      if (info.labor) laborV = U.num((st.grid.v[info.labor - 1] || [])[9]);
      else if (info.material) laborV = baseV - U.num((st.grid.v[info.material - 1] || [])[9]);
      else laborV = baseV;
    }
    var h = '<section class="card exp"><div class="card-h"><b>経費</b><small>資材運搬交通費・諸経費・法定福利費</small></div>';
    if (!base) { h += '<div class="hint">「工事費　計」の行が見つかりません。先に上の「工事費 計」を入れてください。</div></section>'; return h; }
    h += '<div class="kv"><span>工事金額</span><b>J' + base + '</b><em>¥' + U.yen(baseV) + '</em></div>';
    h += '<div class="kv"><span>見積の種類</span><div class="segs">' + [['auto', '自動'], ['kyokai', '協会単価'], ['zaiko', '材工']].map(function (x) {
      return '<button type="button" class="seg' + (st.exp.mode === x[0] ? ' on' : '') + '" data-act="expMode" data-v="' + x[0] + '">' + x[1] + '</button>'; }).join('') + '</div></div>';
    h += '<div class="hint">' + (st.exp.mode === 'auto' ? (info.zaiko ? '自動判定：A列に「資材費」「労務費」の大項目があるので<b>材工の見積</b>' : '自動判定：<b>協会単価の見積</b>（A列に「資材費」「労務費」の大項目がありません）') : (zaiko ? '材工の見積として計算します' : '協会単価の見積として計算します')) +
      (zaiko && plan && plan.laborNote ? '。' + esc(plan.laborNote) : '') + '</div>';
    h += '<div class="exrows">';
    UA.EXP_KEYS.forEach(function (k) {
      var c = cfg[k], est = UA.expenseEstimate(k, c, baseV, laborV, zaiko);
      var fields;
      if (k === 'welfare') fields = zaiko ? '<span class="exf">労務費 ×</span><input class="in num sm" data-f="exp.l" data-k="' + k + '" value="' + esc(c.l) + '" aria-label="法定福利費の率">' :
        '<span class="exf">×</span><input class="in num sm" data-f="exp.k" data-k="' + k + '" value="' + esc(c.k) + '" aria-label="K列"><span class="exf">×</span><input class="in num sm" data-f="exp.l" data-k="' + k + '" value="' + esc(c.l) + '" aria-label="L列">';
      else fields = '<span class="exf">×</span><input class="in num sm" data-f="exp.k" data-k="' + k + '" value="' + esc(c.k) + '" aria-label="K列"><span class="exf">切上</span><input class="in num sm xs" data-f="exp.l" data-k="' + k + '" value="' + esc(c.l) + '" aria-label="L列（切り上げの桁）">';
      h += '<div class="exr' + (c.on ? '' : ' off') + '"><label class="ck"><input type="checkbox" data-f="exp.on" data-k="' + k + '"' + (c.on ? ' checked' : '') + '><span>' + esc(c.label) + '</span></label>' +
        '<em class="exa">¥' + U.yen(est) + '</em><div class="exv">' + fields + '</div></div>';
    });
    h += '</div>';
    var where = plan && plan.ok ? (plan.mode === 'update' ? '「経費」（' + info.exp + '行目）の項目を書き直します' : plan.head + '行目に「経費」を作ります（' + (plan.ins.length ? '行を入れます' : '空き行を使います') + '）') : (plan ? plan.message : '');
    h += '<div class="hint">' + esc(where) + '。K・L・M列に率と式、J列は =M列 にします</div>';
    h += '<div class="row"><button type="button" class="ghost sm" data-act="expDefault" title="今の率を、ほかの見積書でも使う既定にします">この率を既定にする</button><button type="button" class="cta mid" data-act="expRun"' + (plan && plan.ok && !st.ui.busy ? '' : ' disabled') + '>経費を入れる</button></div>';
    return h + '</section>';
  }
  function runExpense() {
    var info = UA.expenseInfo(st.grid);
    var zaiko = st.exp.mode === 'auto' ? info.zaiko : st.exp.mode === 'zaiko';
    var p = UA.expensePlan(st.grid, st.exp.cfg, { mode: zaiko ? 'zaiko' : 'kyokai' });
    if (!p.ok) { toast(p.message, 'warn'); render(); return; }
    var writes = p.writes.map(function (w) { var o = { r: w.r, col: w.col, v: w.v }; if (w.col === 'M') o.fmt = '#,##0'; return o; });
    applyPlan({ ins: p.ins.length ? p.ins[0] : null, writes: writes, select: 'A' + p.select },
      '経費を入れました（' + (p.zaiko ? '材工の見積' : '協会単価の見積') + '、工事金額 J' + p.base + '）' + (p.warn ? '。' + p.warn : '')).then(render);
  }

  /* ---------- 画面 ---------- */
  var ICON = {
    pipe: 'M2 10h5 M2 14h5 M10 10h4 M10 14h4 M17 10h5 M17 14h5 M7 8h3v8H7z M14 8h3v8h-3z',
    valve: 'M3 8v10l9-5z M21 8v10l-9-5z M12 13V6 M8.5 5h7',
    flange: 'M2 10h5 M2 14h5 M17 10h5 M17 14h5 M7 5h3v14H7z M14 5h3v14h-3z',
    rect: 'M3 9h13v10H3z M3 9l4-4h13l-4 4 M16 19l4-4V5',
    round: 'M6 6h12 M6 18h12 M6 6a3 6 0 0 0 0 12 M6 6a3 6 0 0 1 0 12 M18 6a3 6 0 0 1 0 12',
    box: 'M3 7h18v10H3z M7 7v10 M17 7v10 M3 12h4 M17 12h4',
    check: 'M5 12.5l4.2 4.2L19 7', chev: 'M9 6l6 6-6 6', bolt: 'M13 3L5 13h6l-1 8 8-10h-6z', undo: 'M9 14L4 9l5-5 M4 9h10a6 6 0 0 1 0 12h-3',
    sum: 'M17 5H7l6 7-6 7h10', yen: 'M6 4l6 8 6-8 M12 12v8 M8 13h8 M8 17h8', rule: 'M4 7h10 M4 12h16 M4 17h7 M17 4v6 M14 7h6', more: 'M5 12h.01 M12 12h.01 M19 12h.01'
  };
  function svg(d, cls) { return '<svg class="' + (cls || 'ic') + '" viewBox="0 0 24 24" aria-hidden="true"><path d="' + d + '"></path></svg>'; }
  function chip(act, val, label, on, extra) {
    return '<button type="button" class="chip' + (on ? ' on' : '') + '" data-act="' + act + '" data-v="' + esc(val) + '"' + (extra || '') + ' aria-pressed="' + (on ? 'true' : 'false') + '">' + label + '</button>';
  }
  var softPending = false;
  function renderSoft() { if (st.ui.composing || isTyping()) { softPending = true; return; } render(); }
  function isTyping() { var a = document.activeElement; return a && (a.tagName === 'INPUT' || a.tagName === 'TEXTAREA') && a.type !== 'checkbox'; }
  function render() {
    var app = $('#app'); if (!app) return;
    softPending = false;
    var ae = document.activeElement, focus = null;
    if (ae && ae.dataset && ae.dataset.f) focus = { f: ae.dataset.f, k: ae.dataset.k || '', s: ae.selectionStart, e: ae.selectionEnd };
    var main = $('#main'), scroll = main ? main.scrollTop : 0;
    var atEnd = main ? main.scrollTop + main.clientHeight >= main.scrollHeight - 4 : false;
    var h0 = main ? main.clientHeight : 0, hadToast = !!$('.toast');
    app.innerHTML = view();
    var m2 = $('#main');
    if (m2) {
      // お知らせは一覧に重ねず、一覧の下に出す。出たぶん一覧が狭くなるので、
      // 下の端に見えていた入力欄などがそのまま見えるように、狭くなったぶんだけ送る
      if (atEnd) m2.scrollTop = m2.scrollHeight;
      else if (!hadToast && $('.toast')) m2.scrollTop = scroll + Math.max(0, h0 - m2.clientHeight);
      else m2.scrollTop = scroll;
    }
    if (focus) {
      var el = document.querySelector('[data-f="' + focus.f + '"]' + (focus.k ? '[data-k="' + CSS.escape(focus.k) + '"]' : ''));
      if (el) { el.focus(); try { if (focus.s != null) el.setSelectionRange(focus.s, focus.e); } catch (e) { /* noop */ } }
    }
    st.ui.anim = '';
  }
  function view() {
    var h = [];
    h.push('<header class="hd"><div class="brand"><span class="logo">' + svg('M12 9a3 3 0 1 0 0 6a3 3 0 1 0 0-6z M12 4a8 8 0 1 0 0 16a8 8 0 1 0 0-16z M12 4v2 M12 18v2 M4 12h2 M18 12h2', 'ic dark') + '</span>' +
      '<div><b>内訳アシスト</b><small>v' + VERSION + '</small></div></div>' +
      '<div class="hd-r"><button type="button" class="autosw' + (st.auto.on ? ' on' : '') + '" data-act="autoToggle" aria-pressed="' + st.auto.on + '" title="E列にサイズを打つと、保温厚・単位・単価を自動で入れます">' + svg(ICON.bolt) + '<span>自動入力</span><i>' + (st.auto.on ? 'ON' : 'OFF') + '</i></button>' +
      '<button type="button" class="iconbtn" data-act="menu" aria-label="メニュー" aria-expanded="' + st.ui.menu + '">' + svg(ICON.more) + '</button></div></header>');
    if (st.ui.menu) h.push(menuView());
    h.push(statusView());
    if (!st.master) { h.push(onboard()); h.push(toastView()); return h.join(''); }
    h.push('<nav class="tabs" role="tablist">' + Object.keys(TAB).map(function (k) {
      return '<button type="button" role="tab" class="tab' + (st.tab === k ? ' on' : '') + '" data-act="tab" data-v="' + k + '" aria-selected="' + (st.tab === k) + '">' + TAB[k] + '</button>'; }).join('') + '</nav>');
    h.push('<div id="main" class="main">');
    if (!st.model) h.push('<div class="empty">' + (st.modelErr ? esc(st.modelErr) : '「内訳」で始まるシートを開いてください') + '</div>');
    else if (st.tab === 'input') h.push(inputView());
    else if (st.tab === 'total') h.push(totalView());
    else h.push(rulesView());
    h.push('</div>');
    h.push(toastView());
    h.push(footView());
    return h.join('');
  }
  function statusView() {
    var bits = [];
    bits.push(st.master ? '<span class="st-ok" title="' + esc(st.masterInfo ? st.masterInfo.name + '（' + st.masterInfo.at + '）' : '') + '">設定マスタ</span>' : '<button type="button" class="st-ng" data-act="loadMaster">設定マスタを読み込む</button>');
    bits.push(st.sk ? '<span class="st-ok" title="' + esc(st.skInfo ? st.skInfo.name : '') + '">協会単価 ' + st.sk.n + '表</span>' : '<button type="button" class="st-ng" data-act="loadSk">保温積算資料を読み込む</button>');
    var sheet = st.sheets.length > 1 ? '<select class="st-sel" data-f="target" aria-label="書き込み先のシート">' + st.sheets.map(function (n) { return '<option' + (n === st.target ? ' selected' : '') + '>' + esc(n) + '</option>'; }).join('') + '</select>'
      : st.target ? '<span class="' + (st.headerOK ? 'st-ok' : 'st-ng') + '">' + esc(st.target) + (st.headerOK ? '' : '（見出しなし）') + '</span>' : '<span class="st-ng">内訳シートなし</span>';
    bits.push(sheet);
    return '<div class="status">' + bits.join('') + '<button type="button" class="iconbtn sm" data-act="reload" aria-label="シートを読み直す" title="シートを読み直す">' + svg('M20 11a8 8 0 1 0-2.3 5.7 M20 5v6h-6') + '</button></div>';
  }
  function menuView() {
    var nLocal = st.specLocal.add.length + Object.keys(st.specLocal.hide).length + Object.keys(st.specLocal.pin).length;
    return '<div class="menu">' +
      '<button type="button" data-act="loadMaster">設定マスタを読み込む</button>' +
      '<button type="button" data-act="loadSk">保温積算資料（単価表）を読み込む</button>' +
      '<button type="button" data-act="reload">シートを読み直す</button>' +
      '<button type="button" data-act="autoOpen">' + (st.autoOpen ? '✓ ' : '') + 'このブックを開いたら自動で表示</button>' +
      '<button type="button" data-act="exportSpecs">仕様候補の変更を書き出す（' + nLocal + '件）</button>' +
      '<button type="button" data-act="exportRules">単価ルールを書き出す</button>' +
      '<button type="button" data-act="oldVer">前の版（v1.3.2）の画面を開く</button>' +
      '<div class="ver">内訳アシスト v' + VERSION + '／Excel API ' + (st.api19 ? '1.9以上' : st.api17 ? '1.7以上' : '1.7未満') + '</div></div>';
  }
  function onboard() {
    return '<div class="onboard"><h2>はじめに設定マスタを読み込みます</h2>' +
      '<p>積算フォルダの <b>内訳アシスト＼内訳アシスト_設定マスタ.xlsx</b> を選んでください。内容はこのパソコンに記憶され、どの見積書でも使えます。</p>' +
      '<button type="button" class="cta" data-act="loadMaster">設定マスタを選ぶ</button>' + (st.modelErr ? '<p class="err">' + esc(st.modelErr) + '</p>' : '') + '</div>';
  }

  /* ----- 入力タブ ----- */
  function acc(id, no, title, value, body, state, hint) {
    // state: 'open' | 'done' | 'todo'
    var open = state === 'open';
    return '<section class="ac ' + state + (st.ui.anim === id ? ' anim' : '') + '">' +
      '<button type="button" class="ac-h" data-act="acOpen" data-v="' + id + '"' + (state === 'todo' ? ' disabled' : '') + ' aria-expanded="' + open + '">' +
      '<span class="ac-n">' + (state === 'done' ? svg(ICON.check, 'ic ck') : no) + '</span><span class="ac-t">' + title + '</span>' +
      (state === 'done' ? '<span class="ac-v">' + value + '</span><span class="ac-e">変更</span>' : (open || state === 'ready') && hint ? '<span class="ac-hint">' + hint + '</span>' : '') + '</button>' +
      (open ? '<div class="ac-b">' + body + '</div>' : '') + '</section>';
  }
  function inputView() {
    var w = st.wiz, h = [];
    var hasExtra = w.item === 'rect' || w.item === 'box' || w.item === 'valve' || w.item === 'flange';
    function stateOf(id) {
      // open＝開いている／done＝選び終えた／ready＝まだだが選べる／todo＝前の段を先に
      if (w.open === id) return 'open';
      if (id === 'item') return w.item ? 'done' : 'ready';
      if (id === 'place') return w.place ? 'done' : (w.item ? 'ready' : 'todo');
      if (id === 'spec') return w.spec ? 'done' : (w.place && w.item ? 'ready' : 'todo');
      if (id === 'extra') return w.spec ? 'done' : 'todo';
      return 'todo';
    }
    var k = kindOf(w.item);
    h.push(acc('item', '1', '品目', k ? svg(ICON[w.item] || ICON.pipe, 'ic sm') + esc(k.label) : '', itemBody(), stateOf('item'), '配管・バルブ・ダクトなど'));
    h.push(acc('place', '2', '室', w.place ? '(' + esc(w.place) + ')' : '', placeBody(), stateOf('place'), 'C列の(隠蔽)(機械室)など'));
    h.push(acc('spec', '3', '仕様', w.spec ? specInline(w.spec) + (w.cont ? '<em class="tagc">続き</em>' : '') : '', w.open === 'spec' ? specBody() : '', stateOf('spec'), 'E列の(仕様)'));
    if (hasExtra && w.spec) {
      var et = w.item === 'rect' ? '寸法' : w.item === 'box' ? 'BOXの寸法' : '種類（D列）';
      var ev = w.item === 'rect' ? extraSumRect() : w.item === 'box' ? extraSumBox() : (w.d ? esc(w.d) : '<span class="mute">書かない</span>');
      h.push(acc('extra', '4', et, ev, w.open === 'extra' ? extraBody() : '', stateOf('extra'), w.item === 'valve' || w.item === 'flange' ? '選ばなくても書き込めます' : ''));
    }
    if (w.spec) h.push(writeCard());
    h.push(logView());
    return h.join('');
  }
  function itemBody() {
    return '<div class="tiles">' + kinds().map(function (x) {
      return '<button type="button" class="tile' + (x.id === st.wiz.item ? ' on' : '') + '" data-act="wItem" data-v="' + esc(x.id) + '">' + svg(ICON[x.id] || ICON.pipe) + '<b>' + esc(x.label) + '</b></button>';
    }).join('') + '</div>';
  }
  function placeBody() {
    var pl = placeList(), cx = ctxNow(), cur = cx && cx.place;
    function c(p) { return chip('wPlace', p, esc(p) + (cur && U.trimAll(cur) === U.trimAll(p) ? '<i class="here">上と同じ</i>' : ''), p === st.wiz.place); }
    var h = '';
    if (pl.sheet.length) h += '<div class="lbl">この見積書で使用中</div><div class="chips">' + pl.sheet.map(c).join('') + '</div>';
    h += (pl.sheet.length ? '<div class="lbl">ほかの室</div>' : '') + '<div class="chips">' + pl.all.map(c).join('') + '</div>';
    return h;
  }
  function specCols(spec) {
    var p = UA.specParts(spec);
    var c1 = p.tags.filter(function (t) { return t !== '2層' && t !== '遮音'; }).map(function (t) { return t + '＋'; }).join('') + p.pre.map(function (t) { return t + '＋'; }).join('') + p.facing;
    return { lead: c1, mat: p.mat, rest: p.rest.join('＋') };
  }
  // 仕様を1行で（保温材は太字、外装・仕上げはすぐ後ろに続ける）
  function specText(spec) { var c = specCols(spec); return '<span class="c1">' + esc(c.lead) + '<b>' + esc(c.mat) + '</b></span>' + (c.rest ? '<span class="c2">＋' + esc(c.rest) + '</span>' : ''); }
  function specInline(spec) { return '<span class="spi">' + specText(spec) + '</span>'; }
  function specBody() {
    var w = st.wiz, gs = specGroups(), flat = [], h = '';
    st.ui.specIdx = flat;
    var cx = ctxNow();
    var seg = ['', '民間', '官庁'].map(function (c) { return '<button type="button" class="seg' + (st.siteType === c ? ' on' : '') + '" data-act="siteType" data-v="' + c + '">' + (c || 'すべて') + '</button>'; }).join('');
    h += '<div class="sl-top"><span class="segs" title="この見積書を民間・官庁のどちらとして候補を出すか">' + seg + '</span><button type="button" class="link sm" data-act="specEditMode">' + (st.ui.specEditMode ? '編集を終わる' : '候補を編集') + '</button></div>';
    // 上の仕様の続き（仕様の行は書かない）
    if (cx && cx.spec && U.trimAll(cx.place) === U.trimAll(w.place) && cx.prevKind && cx.prevKind !== 'blank' && cx.prevKind !== 'total') {
      var i0 = flat.push({ spec: cx.spec, cont: true }) - 1;
      h += '<button type="button" class="srow cont' + (w.cont ? ' on' : '') + '" data-act="specPick" data-v="' + i0 + '"><span class="c0">↳ 上の仕様の続き（仕様の行は書かない）</span>' + specText(cx.spec) + '</button>';
    }
    if (st.ui.specEditMode) h += '<div class="hint">↑↓ で並びを変えられます（このパソコンに記憶）。☆ でいちばん上に固定、✎ で直す、× で候補から外します</div>';
    var shown = 0, hiddenAll = 0;
    gs.forEach(function (g) {
      var folded = g.key === 'all' && !st.ui.showAll && shown >= 3;
      if (folded) { hiddenAll += g.items.length; return; }
      h += '<div class="slgt">' + esc(g.title) + '</div>';
      h += g.items.map(function (it) { var i = flat.push(Object.assign({ grp: g.key }, it)) - 1; shown++; return optView(it, i, g.key); }).join('');
    });
    if (hiddenAll) h += '<button type="button" class="more" data-act="showAll">ほかの候補を表示（' + hiddenAll + '件）</button>';
    if (!shown && !hiddenAll) h += '<div class="empty sm">この品目・室の候補がありません。下から仕様を入れてください</div>';
    if (gs.hiddenCount) h += '<button type="button" class="link sm" data-act="showHidden">' + (st.ui.showHidden ? '外した候補を隠す' : '外した候補を表示（' + gs.hiddenCount + '件）') + '</button>';
    h += '<div class="addspec"><div class="lbl">候補にない仕様（括弧なしで入力）</div>' +
      '<input class="in" data-f="specNew" value="' + esc(st.ui.newSpec) + '" placeholder="例）ALK付Gw筒＋亀甲金網16m/m" aria-label="仕様">' +
      '<div class="row"><select class="in" data-f="specNewPlace" aria-label="候補を出す室"><option value="here"' + (st.ui.newPlace !== 'all' ? ' selected' : '') + '>' + esc(w.place || '（全箇所）') + 'だけ</option><option value="all"' + (st.ui.newPlace === 'all' ? ' selected' : '') + '>全箇所</option></select>' +
      '<select class="in" data-f="specNewCat" aria-label="民間・官庁"><option value="">民間・官庁とも</option><option value="民間"' + (st.ui.newCat === '民間' ? ' selected' : '') + '>民間だけ</option><option value="官庁"' + (st.ui.newCat === '官庁' ? ' selected' : '') + '>官庁だけ</option></select></div>' +
      '<div class="row"><button type="button" class="ghost" data-act="specUse"' + (st.ui.newSpec ? '' : ' disabled') + '>今回だけ使う</button><button type="button" class="ghost acc" data-act="specAdd"' + (st.ui.newSpec ? '' : ' disabled') + '>候補に追加して使う</button></div></div>';
    return h;
  }
  function optView(it, i, gk) {
    var w = st.wiz, on = !w.cont && it.spec === w.spec && gk !== 'hidden';
    var ed = st.ui.specEdit;
    if (ed && ed.i === i) {
      return '<div class="sopt edit"><input class="in" data-f="specEditText" value="' + esc(ed.spec) + '" aria-label="仕様を編集">' +
        '<div class="row"><select class="in" data-f="specEditPlace"><option value="here"' + (ed.place !== '（全箇所）' ? ' selected' : '') + '>' + esc(w.place) + 'だけ</option><option value="all"' + (ed.place === '（全箇所）' ? ' selected' : '') + '>全箇所</option></select>' +
        '<select class="in" data-f="specEditCat"><option value="">民間・官庁とも</option><option value="民間"' + (ed.cat === '民間' ? ' selected' : '') + '>民間だけ</option><option value="官庁"' + (ed.cat === '官庁' ? ' selected' : '') + '>官庁だけ</option></select></div>' +
        '<div class="row"><button type="button" class="ghost" data-act="specCancel">取消</button><button type="button" class="ghost acc" data-act="specSave">保存</button></div></div>';
    }
    var acts = '';
    if (st.ui.specEditMode || gk === 'hidden') {
      if (gk === 'hidden') acts = '<button type="button" class="ic-b" data-act="specUnhide" data-v="' + i + '" title="候補に戻す">戻す</button>';
      else if (gk !== 'sheet') acts = '<span class="ud"><button type="button" class="ic-b" data-act="specMove" data-v="' + i + '|-1" title="上へ">↑</button><button type="button" class="ic-b" data-act="specMove" data-v="' + i + '|1" title="下へ">↓</button></span>' +
        '<button type="button" class="ic-b' + (it.pin ? ' on' : '') + '" data-act="specPin" data-v="' + i + '" title="' + (it.pin ? '固定をやめる' : 'いちばん上に固定') + '">' + (it.pin ? '★' : '☆') + '</button>' +
        '<button type="button" class="ic-b" data-act="specEdit" data-v="' + i + '" title="直す">✎</button>' +
        '<button type="button" class="ic-b" data-act="specHide" data-v="' + i + '" title="' + (it.src === 'local' ? '削除' : '候補から外す（設定マスタは変わりません）') + '">×</button>';
      else acts = '<span class="ic-n" title="この見積書で使っている仕様は、ここでは直せません">—</span>';
    }
    return '<div class="srw' + (on ? ' on' : '') + (gk === 'hidden' ? ' off' : '') + '">' +
      '<button type="button" class="srow" data-act="specPick" data-v="' + i + '" title="(' + esc(it.spec) + ')"' + (gk === 'hidden' ? ' disabled' : '') + '>' +
      (it.pin ? '<i class="pin">★</i>' : '') + specText(it.spec) + '</button>' +
      (acts ? '<span class="srow-a">' + acts + '</span>' : '') + '</div>';
  }
  // 角ダクト・BOX：25・50 のボタンと、ほかの厚みの入力欄
  function ductThickChips(o, act, ff) {
    var fx = U.thickKey(o.fx), cur = fx || o.f || '';
    return '<div class="thick"><span class="lbl">保温厚</span>' + ['25', '50'].map(function (t) { return '<button type="button" class="tk' + (!fx && String(cur) === t ? ' on' : '') + '" data-act="' + act + '" data-v="' + t + '">' + t + '</button>'; }).join('') +
      '<span class="tkx' + (fx ? ' on' : '') + '"><span>ほか</span><input class="in num" data-f="' + ff + '" value="' + esc(o.fx) + '" inputmode="decimal" placeholder="mm" aria-label="ほかの保温厚（mm）"></span></div>';
  }
  function thickChips(cur, act) {
    var tl = (st.master.output && st.master.output.thickList) || ['20', '25', '30', '40', '50', '65', '75'];
    return '<div class="thick"><span class="lbl">保温厚</span>' + tl.map(function (t) { return '<button type="button" class="tk' + (String(cur) === String(t) ? ' on' : '') + '" data-act="' + act + '" data-v="' + t + '">' + t + '</button>'; }).join('') + '</div>';
  }
  function priceCap(pr) {
    if (!pr) return '';
    if (pr.p == null) return '<span class="pc ng" title="' + esc(pr.why || '') + '">協会単価なし</span>';
    var t = pr.via === 'rule' ? '計算ルール' : '協会単価 §' + (pr.id || '');
    return '<span class="pc ok" title="' + esc(pr.ref ? pr.ref.y : '') + '">¥' + U.yen(pr.p) + '<small>' + esc(t) + '</small></span>';
  }
  function extraSumRect() { var ex = extraLines(), r = st.wiz.rect; return esc((r.sub ? r.sub + (r.t16 ? '(1.6t)' : '') + '　' : '') + (ex.g > 0 ? U.fmtQ(ex.g) + '㎡　' : '') + 't' + (ex.f || '')); }
  function extraSumBox() { var ex = extraLines(), n = ex.lines.filter(function (l) { return !l.head; }).length; return esc((st.wiz.box.name || '名前なし') + '　' + (n ? n + '件' : '寸法未入力') + '　t' + (ex.f || '')); }
  function sideBtn(k, label) { var on = st.wiz.box.show[k]; return '<button type="button" class="sbtn' + (on ? ' on' : '') + '" data-act="bShow" data-v="' + k + '" aria-pressed="' + on + '" title="' + label + 'を E列に' + (on ? '出す' : '出さない') + '">' + label + '</button>'; }
  function extraBody() {
    var w = st.wiz, h = '';
    if (w.item === 'valve' || w.item === 'flange') {
      var subs = subsOf(w.item);
      h += '<div class="chips">' + chip('wD', '', '書かない', !w.d) + subs.map(function (s) { return chip('wD', s.id, esc(s.label), w.d === s.id, s.tip ? ' title="' + esc(s.tip) + '"' : ''); }).join('') + '</div>';
      h += '<div class="hint">最初の明細の行の D列に書きます。サイズ・数量は Excel で直接入れてください</div>';
      return h;
    }
    var ex = extraLines();
    if (w.item === 'rect') {
      var rt = w.rect;
      h += '<div class="lbl">ダクトの種別（D列）</div><div class="chips">' + subsOf('rect').map(function (s) { return chip('rSub', s.id, esc(s.id ? s.label : '（なし）'), rt.sub === s.id); }).join('') + chip('rT16', '1', '(1.6t)', rt.t16) + '</div>';
      h += ductThickChips(rt, 'rF', 'rFx');
      h += '<div class="kv"><span>数量（㎡）</span><input class="in num" data-f="rRq" value="' + esc(rt.rq) + '" placeholder="空欄で可" inputmode="decimal" aria-label="数量（㎡）"><span class="hint">空欄なら G列に直接入れます</span></div>';
      h += '<div class="kv"><span>単価</span>' + priceCap(ex.pr) + (ex.pr && ex.pr.p != null && ex.g > 0 ? '<em>¥' + U.yen(ex.pr.p * ex.g) + '</em>' : '') + '</div>';
      return h;
    }
    if (w.item === 'box') {
      var bx = w.box, names = [];
      subsOf('rect').forEach(function (s) { if (s.id) names.push(s.id); });
      subsOf('box').forEach(function (s) { var n = s.id || s.label; if (n && names.indexOf(n) < 0) names.push(n); });
      ['BOX', 'チャンバー'].forEach(function (n) { if (names.indexOf(n) < 0) names.push(n); });
      h += '<div class="lbl">名前（D列に1行で書き、その下の行の E列に寸法）</div><div class="chips">' + names.map(function (n) { return chip('bName', n, esc(n), bx.name === n); }).join('') + '</div>';
      h += '<input class="in" data-f="bName" value="' + esc(bx.name) + '" placeholder="直接入力も可（例：SMダクト　BOX）" aria-label="名前">';
      h += '<div class="lbl">寸法（mm）と台数</div><div class="dims">' + bx.boxes.map(function (b, i) {
        var ok = U.num(b.w) > 0 && U.num(b.h) > 0 && U.num(b.l) > 0;
        var l = ok ? ex.lines.filter(function (x) { return !x.head; })[bx.boxes.slice(0, i).filter(function (y) { return U.num(y.w) > 0 && U.num(y.h) > 0 && U.num(y.l) > 0; }).length] : null;
        return '<div class="dr box"><input class="in num" data-f="bDim.w" data-k="' + i + '" value="' + esc(b.w) + '" placeholder="W" inputmode="numeric" aria-label="W（幅）"><span>×</span>' +
          '<input class="in num" data-f="bDim.h" data-k="' + i + '" value="' + esc(b.h) + '" placeholder="H" inputmode="numeric" aria-label="H（高さ）"><span>×</span>' +
          '<input class="in num" data-f="bDim.l" data-k="' + i + '" value="' + esc(b.l) + '" placeholder="L" inputmode="numeric" aria-label="L（長さ）">' +
          '<input class="in num n" data-f="bDim.n" data-k="' + i + '" value="' + esc(b.n) + '" inputmode="numeric" aria-label="台数"><span>台</span>' +
          '<button type="button" class="x" data-act="bDel" data-v="' + i + '" aria-label="この行を消す">×</button>' +
          (l ? '<div class="bres"><b>E列：' + esc(l.e) + '</b>　' + l.area + '㎡　' + priceCap(l.pr) + '</div>' : '') + '</div>';
      }).join('') + '<button type="button" class="ghost sm" data-act="bAdd">＋寸法を追加</button></div>';
      h += '<div class="kv"><span>E列に出す辺</span><div class="sides">' + sideBtn('w', 'W') + sideBtn('h', 'H') + sideBtn('l', 'L') +
        '<button type="button" class="sbtn sym' + (bx.sym ? ' on' : '') + '" data-act="bSym" aria-pressed="' + bx.sym + '" title="W500×H1000×L500 のように記号を付けます">記号</button></div></div>';
      h += '<div class="kv"><span>面の数</span><div class="segs">' + [[false, '6面'], [true, '5面']].map(function (x) {
        return '<button type="button" class="seg' + (ex.five === x[0] ? ' on' : '') + '" data-act="bFive" data-v="' + (x[0] ? '1' : '0') + '">' + x[1] + '</button>'; }).join('') + '</div><span class="hint">面積×1.1。チャンバー・ガラリは5面</span></div>';
      h += ductThickChips(bx, 'bF', 'bFx');
      h += '<div class="hint">単価は3辺すべてから計算します（E列に出さない辺も使います）</div>';
      return h;
    }
    return h;
  }
  function writeCard() {
    var w = st.wiz, R = targetRow(), p = wizPlan(), ex = extraLines(), h = '<section class="wcard">';
    if (!R) return h + '<div class="wc-h"><b>書き込む位置</b></div><div class="hint">内訳シートで、書き込みたい行を選んでください</div></section>';
    if (!p || !p.ok) return h + '<div class="wc-h"><b>書き込む位置</b></div><div class="hint warn">' + esc(p ? p.message : '準備中') + '</div></section>';
    h += '<div class="wc-h"><b>書き込む位置</b><span class="wc-r">' + p.P + '行目〜</span><small>Excel で選んでいる行</small></div>';
    if (p.needIns) {
      var n = p.ins;
      h += '<div class="insrow"><span class="ins-t">' + esc(p.busy.join('・')) + '行目にデータがあるので、<b>行を挿入</b>して書きます</span>' +
        '<div class="stepper"><button type="button" class="ic-b" data-act="insN" data-v="-1" aria-label="1行減らす"' + (n <= p.minIns ? ' disabled' : '') + '>−</button>' +
        '<input class="in num" data-f="insN" value="' + esc(w.insN === '' ? n : w.insN) + '" inputmode="numeric" aria-label="挿入する行数"><span>行</span>' +
        '<button type="button" class="ic-b" data-act="insN" data-v="1" aria-label="1行増やす">＋</button></div></div>';
    }
    h += '<div class="wrs">';
    var r = p.P;
    if (p.writeSpec) { h += '<div class="wr"><span class="rn">' + r + '</span><span class="wt">' + (p.writePlace ? '<i>(' + esc(w.place) + ')</i>' : '') + '(' + esc(U.normSpec(w.spec)) + ')</span></div>'; r++; }
    if (ex.lines.length) ex.lines.forEach(function (l) {
      if (l.head) { h += '<div class="wr ln"><span class="rn">' + r + '</span><span class="wt"><i class="dcol">D列</i>' + esc(l.d) + '</span></div>'; r++; return; }
      var gtx = l.g === '' ? '数量は G列に' : U.fmtQ(l.g) + l.h;
      h += '<div class="wr ln"><span class="rn">' + r + '</span><span class="wt">' + esc((l.d ? l.d + '　' : '') + l.e + '　t' + l.f + '　' + gtx) + '</span>' + priceCap(l.pr) + '</div>'; r++;
    });
    else h += '<div class="wr in"><span class="rn">' + r + '</span><span class="wt">' + (w.d ? '<i>' + esc(w.d) + '</i>　' : '') + '<span class="caret"></span>E列にサイズを入力</span></div>';
    if (p.ins > p.k) for (var e = 0; e < Math.min(p.ins - p.k, 3); e++) h += '<div class="wr add"><span class="rn">' + (p.P + p.k + e) + '</span><span class="wt">（挿入した空き行）</span></div>';
    if (p.ins > p.k + 3) h += '<div class="wr add"><span class="rn">…</span><span class="wt">ほか ' + (p.ins - p.k - 3) + '行</span></div>';
    h += '</div>';
    var notes = [];
    if (!p.writeSpec) notes.push('上の仕様の続きなので、仕様の行は書きません');
    else if (!p.writePlace) notes.push('室は上と同じなので C列は書きません');
    if (p.ins) notes.push('挿入した行は内訳の初めの書式（A〜J列）にします。' + p.insAt + '行目から下は ' + p.ins + '行下にずれます');
    if (notes.length) h += '<div class="hint">' + esc(notes.join('。')) + '</div>';
    h += '<button type="button" class="cta" data-act="wWrite"' + (st.ui.busy ? ' disabled' : '') + '>' + (st.ui.busy ? '書き込み中…' : p.ins ? p.ins + '行挿入して書き込む' : '書き込む') + '</button>';
    return h + '</section>';
  }
  function logView() {
    var L = st.auto.log;
    var canUndo = st.undo && st.undo.kind === 'auto';
    var h = '<section class="log"><div class="log-h"><b>' + svg(ICON.bolt, 'ic sm') + '自動入力の記録</b>' +
      (canUndo ? '<button type="button" class="undo" data-act="undo" title="最後に自動で入れた内容を元に戻します（Excel の Ctrl＋Z で戻らないときに）">' + svg(ICON.undo, 'ic sm') + '元に戻す</button>' : st.auto.on ? '<small>E列にサイズを打つと入ります</small>' : '<small class="warn">自動入力は OFF です</small>') + '</div>';
    if (!L.length) return h + '<div class="empty sm">まだありません。仕様の行の下の E列に 50 や φ200 と打ってみてください</div></section>';
    h += '<div class="lg">' + L.slice(0, 8).map(function (x) {
      var ng = x.p == null && !x.kept;
      return '<div class="lr' + (ng ? ' ng' : '') + '" title="' + esc((x.fFrom ? '保温厚：' + x.fFrom + '\n' : '') + (x.y || x.why || '')) + '"><span class="rn">' + x.r + '</span><span class="lz">' + esc((x.d ? x.d + ' ' : '') + x.size) + '</span>' +
        '<span class="lf' + (x.fNew ? ' new' : '') + '">' + esc(x.f ? 't' + x.f : '') + '</span><span class="lh">' + esc(x.h || '') + '</span>' +
        '<span class="lp">' + (x.kept ? '<small>手入力のまま</small>' : ng ? '<small>単価なし</small>' : '¥' + U.yen(x.p)) + '</span></div>';
    }).join('') + '</div>';
    return h + '</section>';
  }

  /* ----- 計・経費タブ ----- */
  function totalView() {
    var R = targetRow(), sug = suggestTotal(), h = '';
    var rowInfo = !R ? '内訳シートで行を選んでください' : R + '行目' + (st.grid && rowBlank(R) ? '（空き行）' : '');
    h += '<div class="selinfo"><span>選んでいる行</span><b>' + esc(rowInfo) + '</b><div class="segs sm">' + [[true, '名前すべて'], [false, '番号だけ']].map(function (x) {
      return '<button type="button" class="seg' + (st.ui.full === x[0] ? ' on' : '') + '" data-act="full" data-v="' + (x[0] ? '1' : '0') + '">' + x[1] + '</button>'; }).join('') + '</div></div>';
    if (sug && sug !== 'exp') {
      var sd = TOTALS.filter(function (t) { return t.id === sug; })[0], sp = sd ? totalPlanOf(sd) : null;
      if (sp && sp.ok) h += '<button type="button" class="sug" data-act="total" data-v="' + sd.id + '"><span class="sg-l">おすすめ</span><span class="sg-t">' + esc(sd.grp === '計' ? sd.label : sd.label + ' ' + sd.grp) + '</span><span class="sg-d">' + sp.totalRow + '行目「' + esc(sp.label) + '」' + esc(sp.formula) + '</span></button>';
    }
    h += '<section class="card"><div class="card-h"><b>計を入れる</b><small>選んでいる行の下に、計の行を作ります</small></div>';
    ['集計', '小計', '計'].forEach(function (g) {
      h += '<div class="tgrp"><span class="tg-l">' + (g === '集計' ? '集計<small>範囲を合計</small>' : g === '小計' ? '小計<small>計を合計</small>' : '') + '</span><div class="tbtns">';
      TOTALS.filter(function (t) { return t.grp === g; }).forEach(function (t) {
        var p = totalPlanOf(t), ok = p && p.ok;
        h += '<button type="button" class="tbtn' + (sug === t.id ? ' sg' : '') + '" data-act="total" data-v="' + t.id + '"' + (ok && !st.ui.busy ? '' : ' disabled') + ' title="' + esc(t.tip + (p ? '\n' + (ok ? p.totalRow + '行目「' + p.label + '」' + p.formula : p.message) : '')) + '">' +
          '<b>' + esc(t.label) + '</b><small>' + esc(ok ? p.totalRow + '行目' + ((t.kind === 'range' || t.kind === 'sub') && p.label.replace(/　計$/, '') ? '・' + p.label.replace(/　計$/, '') : '') : '—') + '</small></button>';
      });
      h += '</div></div>';
    });
    h += '<div class="hint">ボタンにマウスを置くと、入れる式が表示されます。押したあとは「元に戻す」で取り消せます</div></section>';
    h += expView();
    var dp = st.grid && R ? UA.totalPlan(st.grid, R, 'discount', { amount: st.ui.discount }) : null;
    var dHint = dp ? (dp.ok ? dp.at + '行目に出精値引、' + dp.totalRow + '行目に総合計（合計＝J' + dp.from + '）' : dp.message) : '「合計」の行のすぐ下を選んでください';
    if (dp && dp.ok) {
      var gv = (st.grid.v[dp.from - 1] || [])[9];
      if (typeof gv === 'number') dHint += gv >= dp.amount ? '。値引き ¥' + U.yen(gv - dp.amount) + '（合計 ¥' + U.yen(gv) + '）' : '。決定金額が合計（¥' + U.yen(gv) + '）より大きくなっています';
    }
    h += '<section class="card"><div class="card-h"><b>出精値引き</b><small>決定金額にそろえる</small></div><div class="disc"><span class="disc-l">決定金額</span>' +
      '<span class="disc-in"><input class="in num" data-f="discount" value="' + esc(commaDigits(st.ui.discount)) + '" placeholder="0" inputmode="numeric" maxlength="11" autocomplete="off" aria-label="決定金額（円）"><em>円</em></span>' +
      '<button type="button" class="ghost acc" data-act="discount"' + (dp && dp.ok && !st.ui.busy ? '' : ' disabled') + '>入れる</button></div>' +
      '<div class="hint">' + esc(dHint) + '</div></section>';
    return h;
  }
  function commaDigits(d) { d = String(d == null ? '' : d).replace(/[^0-9]/g, ''); return d.replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
  function toHalfDigits(t) { return String(t == null ? '' : t).replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).replace(/[，、]/g, ','); }
  function rowBlank(r) { var row = st.grid.v[r - 1]; if (!row) return true; for (var c = 0; c < 9; c++) if (!U.isEmpty(row[c])) return false; return true; }

  /* ----- 単価ルールタブ ----- */
  function rulesView() {
    var list = rules(), h = '';
    h += '<div class="intro">協会単価の表に無いもの（GV の小さいサイズ・フレックス継手など）を、同じ仕様の直管の単価から計算します。<b>上のルールから順に</b>当てはめます。自動入力・選択行に単価・入力フォームのすべてで使います。</div>';
    h += '<div class="rules">' + list.map(function (r, i) {
      if (st.ui.ruleEdit && st.ui.ruleEdit.i === i) return ruleForm(st.ui.ruleEdit);
      var d = UA.describeRule(r, st.master);
      return '<div class="rc' + (r.on ? '' : ' off') + '"><button type="button" class="sw' + (r.on ? ' on' : '') + '" data-act="ruleOn" data-v="' + i + '" aria-pressed="' + !!r.on + '" title="' + (r.on ? '使う' : '使わない') + '"><i></i></button>' +
        '<div class="rc-b"><b>' + esc(r.name || '（名前なし）') + '</b><div class="rc-s"><span class="rw">' + esc(d.when) + '</span><span class="ra">→ ' + esc(d.then) + '</span><span class="rr">' + esc(d.round) + '</span></div>' +
        (r.memo ? '<small>' + esc(r.memo) + '</small>' : '') + '</div>' +
        '<div class="rc-a"><button type="button" class="ic-b" data-act="ruleUp" data-v="' + i + '" title="上へ（先に当てはめる）"' + (i ? '' : ' disabled') + '>↑</button>' +
        '<button type="button" class="ic-b" data-act="ruleDown" data-v="' + i + '" title="下へ"' + (i < list.length - 1 ? '' : ' disabled') + '>↓</button>' +
        '<button type="button" class="ic-b" data-act="ruleEdit" data-v="' + i + '" title="直す">✎</button></div></div>';
    }).join('') + '</div>';
    if (st.ui.ruleEdit && st.ui.ruleEdit.i === -1) h += '<div class="rules">' + ruleForm(st.ui.ruleEdit) + '</div>';
    h += '<div class="row"><button type="button" class="ghost acc" data-act="ruleAdd"' + (st.ui.ruleEdit ? ' disabled' : '') + '>＋ ルールを追加</button><button type="button" class="ghost" data-act="ruleReset">初めの設定に戻す</button></div>';
    h += '<div class="hint">' + (st.rulesLocal ? 'このパソコンに記憶しています。メニューの「単価ルールを書き出す」で一覧にでき、設定マスタの M_単価計算 シートに貼るとほかのパソコンでも使えます。' : st.master && st.master.calcRulesRaw && st.master.calcRulesRaw.length ? '設定マスタの M_単価計算 シートのルールです。' : '初めの設定のルールです。') + '</div>';
    return h;
  }
  function ruleForm(e) {
    var r = e.r, its = ['pipe', 'valve', 'flange', 'round', 'rect', 'box'];
    var ex = '';
    var base = 4310, val = UA.roundVal(base * (U.num(r.factor) || 1), r.round);
    ex = '例）' + (r.base === 'kind' && r.baseKind ? r.baseKind : '直管') + ' ' + U.yen(base) + '円 × ' + (U.num(r.factor) || 1) + ' ＝ ' + (Math.round(base * (U.num(r.factor) || 1) * 100) / 100).toLocaleString('ja-JP') + ' → <b>' + U.yen(val) + '円</b>';
    return '<div class="rf"><div class="rf-r"><label>名前</label><input class="in" data-f="re.name" value="' + esc(r.name) + '" placeholder="例）GV（50A以下）"></div>' +
      '<div class="rf-r"><label>品目</label><div class="chips sm">' + its.map(function (id) { return chip('reItem', id, esc(UA.itemLabel(st.master, id)), (r.items || []).indexOf(id) >= 0); }).join('') + '</div></div>' +
      '<div class="rf-r"><label>D列（種別）に<br>含む語</label><input class="in" data-f="re.words" value="' + esc(r.words) + '" placeholder="例）GV　（、で区切ると複数）"></div>' +
      '<div class="rf-r"><label>サイズ</label><div class="row in2"><input class="in num" data-f="re.min" value="' + esc(r.min) + '" placeholder="下限なし" inputmode="numeric"><span>〜</span><input class="in num" data-f="re.max" value="' + esc(r.max) + '" placeholder="上限なし" inputmode="numeric"></div></div>' +
      '<div class="rf-r"><label>基準の単価</label><div class="segs">' + [['run', '同じ仕様の直管'], ['kind', '別の種別']].map(function (x) { return '<button type="button" class="seg' + (r.base === x[0] ? ' on' : '') + '" data-act="reBase" data-v="' + x[0] + '">' + x[1] + '</button>'; }).join('') + '</div>' +
      (r.base === 'kind' ? '<input class="in" data-f="re.baseKind" value="' + esc(r.baseKind) + '" placeholder="例）GV">' : '') + '</div>' +
      '<div class="rf-r"><label>係数</label><input class="in num" data-f="re.factor" value="' + esc(r.factor) + '" inputmode="decimal" placeholder="例）1.1"></div>' +
      '<div class="rf-r"><label>端数</label><select class="in" data-f="re.round">' + ['up10', 'up100', 'up1', 'none'].map(function (k) { return '<option value="' + k + '"' + (r.round === k ? ' selected' : '') + '>' + esc(UA.ROUND_HINT[k]) + '</option>'; }).join('') + '</select></div>' +
      '<div class="rf-r"><label>メモ</label><input class="in" data-f="re.memo" value="' + esc(r.memo) + '" placeholder="任意"></div>' +
      '<div class="rf-ex">' + ex + '</div>' +
      '<div class="row"><button type="button" class="ghost" data-act="reCancel">取消</button>' + (e.i >= 0 ? '<button type="button" class="ghost danger" data-act="reDel">削除</button>' : '') + '<button type="button" class="ghost acc" data-act="reSave">保存</button></div></div>';
  }

  /* ----- 下の帯・お知らせ ----- */
  function footView() {
    var s = st.sel, n = s.sheet === st.target ? s.rows.length : 0;
    var info = n ? (s.areas > 1 ? s.areas + 'か所・' : '') + n + '行を選択中' : '行を選ぶと単価を入れられます';
    return '<footer class="ft"><span class="ft-i">' + esc(info) + '</span><button type="button" class="ftbtn" data-act="priceSel"' + (n && !st.ui.busy ? '' : ' disabled') + ' title="選んでいる行（離れた範囲も可）の見積単価を入れ直します。参照元は Y〜AB列">' + svg(ICON.yen) + '選択行に単価</button></footer>';
  }
  function toastView() {
    var t = st.ui.toast;
    if (!t) return '';
    return '<div class="toast ' + t.kind + '" role="status"><span>' + esc(t.text) + '</span>' + (t.undo && st.undo ? '<button type="button" data-act="undo">' + svg(ICON.undo, 'ic sm') + '元に戻す</button>' : '') +
      '<button type="button" class="x" data-act="closeToast" aria-label="閉じる">×</button></div>';
  }
  var toastTimer = null;
  function toast(text, kind, withUndo) {
    st.ui.toast = { text: text, kind: kind || 'ok', undo: !!withUndo };
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { st.ui.toast = null; render(); }, withUndo ? 30000 : (kind === 'err' ? 15000 : 7000));
  }

  /* ---------- 書き出し ---------- */
  function exportSheet(name, head, rows, done) {
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject(name);
      return ctx.sync().then(function () {
        if (ws.isNullObject) ws = ctx.workbook.worksheets.add(name); else ws.getRange().clear();
        var lc = U.colLetter(head.length);
        ws.getRange('A1:' + lc + '1').values = [head];
        ws.getRange('A1:' + lc + '1').format.font.bold = true;
        if (rows.length) ws.getRange('A2:' + lc + (rows.length + 1)).values = rows;
        ws.activate();
        return ctx.sync();
      });
    }).then(function () { toast(done, 'ok'); render(); }).catch(function (e) { toast(msg(e), 'err'); render(); });
  }
  function exportSpecs() {
    var L = st.specLocal, rows = [];
    L.add.forEach(function (a) { rows.push(['追加', a.kind, a.place, a.spec, a.cat || '', a.from ? '「' + a.from + '」を直したもの' : '']); });
    Object.keys(L.hide).forEach(function (k) { var p = k.split('|'); rows.push(['候補から外す', p[0], p[1], p.slice(2).join('|'), '', '設定マスタの「使用する」を × に']); });
    Object.keys(L.pin).forEach(function (k) { var p = k.split('|'); rows.push(['固定', p[0], p[1], p.slice(2).join('|'), '', '設定マスタの「優先順位」に 1 など']); });
    if (!rows.length) { toast('仕様候補の変更はまだありません', 'warn'); render(); return; }
    exportSheet('内訳アシスト_仕様候補', ['操作', '仕様区分', '施工箇所', '仕様（括弧なし）', '区分（民間・官庁）', 'メモ'], rows,
      '「内訳アシスト_仕様候補」シートに書き出しました。「追加」の行は設定マスタの M_仕様候補 に貼り付けると、ほかのパソコンでも使えます。');
  }
  function exportRules() {
    var rows = rules().map(function (r) { return [r.on ? '○' : '×', r.name, (r.items || []).join(','), r.words, r.min, r.max, r.base === 'kind' ? r.baseKind : '直管', U.num(r.factor) || 1, UA.ROUND_LABEL[r.round] || '', r.memo || '']; });
    exportSheet('内訳アシスト_単価計算', ['使用する', '名前', '品目ID', 'D列に含む語', 'サイズ以上', 'サイズ以下', '基準', '係数', '端数', 'メモ'], rows,
      '「内訳アシスト_単価計算」シートに書き出しました。設定マスタに「M_単価計算」シートを作って貼り付けると、ほかのパソコンでも同じルールになります。');
  }
  function setAutoOpen(on) {
    try {
      Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', !!on);
      Office.context.document.settings.saveAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) { st.autoOpen = !!on; toast(on ? 'このブックを開いたとき、パネルを自動で表示します。ブックを保存すると設定が残ります（ひな型で設定すればコピーにも引き継がれます）。' : '自動表示をやめました。ブックを保存すると反映されます。', 'ok'); }
        else toast('設定を保存できませんでした', 'err');
        render();
      });
    } catch (e) { toast(msg(e), 'err'); render(); }
  }
  function saveBoxUi() { save(K_BOX, { show: st.wiz.box.show, sym: st.wiz.box.sym }); }
  function saveSpecLocal() { if (!save(K_SPECS, st.specLocal)) toast('候補の変更をこのパソコンに記憶できませんでした', 'warn'); }
  function specKindOf() { return UA.KIND_OF_ITEM[st.wiz.item] || st.wiz.item; }

  /* ---------- 入力フォームの進め方 ---------- */
  function nextOpen() {
    var w = st.wiz;
    if (!w.item) return 'item';
    if (!w.place) return 'place';
    if (!w.spec) return 'spec';
    if (w.item === 'rect' || w.item === 'box') return 'extra';
    return null;
  }
  function setOpen(id) { st.wiz.open = id; st.ui.anim = id || ''; }
  function pickSpec(spec, cont) {
    var w = st.wiz;
    w.spec = U.normSpec(spec); w.cont = !!cont; w.done = null;
    st.ui.specEdit = null;
    setOpen(w.item === 'rect' || w.item === 'box' || w.item === 'valve' || w.item === 'flange' ? 'extra' : null);
  }

  /* ---------- 操作 ---------- */
  function bindEvents() {
    document.addEventListener('click', function (ev) {
      var el = ev.target.closest('[data-act]');
      if (!el || el.disabled) return;
      var a = el.dataset.act, v = el.dataset.v, w = st.wiz;
      if (a !== 'menu' && st.ui.menu) st.ui.menu = false;
      switch (a) {
        case 'menu': st.ui.menu = !st.ui.menu; break;
        case 'loadMaster': st.ui.menu = false; $('#masterFile').click(); return;
        case 'loadSk': st.ui.menu = false; $('#skFile').click(); return;
        case 'reload': refreshSheets().then(readSelection).then(function () { toast('シートを読み直しました', 'ok'); render(); }); return;
        case 'autoOpen': setAutoOpen(!st.autoOpen); return;
        case 'exportSpecs': exportSpecs(); return;
        case 'exportRules': exportRules(); return;
        case 'oldVer': window.location.href = 'v1/taskpane.html'; return;
        case 'tab': st.tab = v; saveUi(); readSelection().then(render); return;
        case 'autoToggle': st.auto.on = !st.auto.on; saveUi(); toast(st.auto.on ? '自動入力を ON にしました。E列にサイズを打つと、保温厚・単位・単価が入ります' : '自動入力を OFF にしました', 'info'); break;
        // 入力フォーム
        case 'acOpen': setOpen(w.open === v ? null : v); break;
        case 'wItem': {
          var changed = w.item !== v;
          w.item = v; w.done = null;
          if (changed) { w.spec = ''; w.cont = false; w.d = ''; }
          setOpen(nextOpen());
          break;
        }
        case 'wPlace': {
          var ch2 = w.place !== v;
          w.place = v; w.done = null; if (ch2) { w.spec = ''; w.cont = false; }
          setOpen(nextOpen());
          break;
        }
        case 'specPick': { var o = st.ui.specIdx[+v]; if (o) pickSpec(o.spec, o.cont); break; }
        case 'showAll': st.ui.showAll = true; break;
        case 'wD': w.d = v || ''; setOpen(null); break;
        case 'rSub': w.rect.sub = v; break;
        case 'rT16': w.rect.t16 = !w.rect.t16; break;
        case 'rF': w.rect.f = v; w.rect.fx = ''; break;
        case 'bName': w.box.name = v; w.box.five = null; break;
        case 'bFive': w.box.five = v === '1'; break;
        case 'bF': w.box.f = v; w.box.fx = ''; break;
        case 'bShow': w.box.show[v] = !w.box.show[v]; if (!w.box.show.w && !w.box.show.h && !w.box.show.l) { w.box.show[v] = true; toast('少なくとも1つの辺は E列に出します', 'info'); } saveBoxUi(); break;
        case 'bSym': w.box.sym = !w.box.sym; saveBoxUi(); break;
        case 'bAdd': w.box.boxes.push({ w: '', h: '', l: '', n: '1' }); break;
        case 'bDel': w.box.boxes.splice(+v, 1); if (!w.box.boxes.length) w.box.boxes = [{ w: '', h: '', l: '', n: '1' }]; break;
        case 'insN': { var pn = wizPlan(), cur = U.num(w.insN) || (pn ? pn.ins : 0); w.insN = String(Math.max(pn ? pn.minIns : 1, cur + (+v))); break; }
        case 'specMove': {
          var pv = String(v).split('|'), om = st.ui.specIdx[+pv[0]];
          if (!om) break;
          var grpItems = st.ui.specIdx.filter(function (x) { return x.grp === om.grp && !x.cont; }).map(function (x) { return x.spec; });
          var ia = grpItems.indexOf(om.spec), ib = ia + (+pv[1]);
          if (ia < 0 || ib < 0 || ib >= grpItems.length) break;
          var tmp2 = grpItems[ia]; grpItems[ia] = grpItems[ib]; grpItems[ib] = tmp2;
          var okey = specKindOf() + '|' + w.place;
          st.specLocal.order = st.specLocal.order || {};
          var keep = (st.specLocal.order[okey] || []).filter(function (sp) { return grpItems.indexOf(sp) < 0; });
          st.specLocal.order[okey] = keep.concat(grpItems);
          saveSpecLocal();
          break;
        }
        case 'wWrite': fresh(writeWizard); return;
        // 仕様の候補の編集
        case 'specPin': { var op = st.ui.specIdx[+v]; if (op && op.key) { if (st.specLocal.pin[op.key]) delete st.specLocal.pin[op.key]; else st.specLocal.pin[op.key] = true; saveSpecLocal(); } break; }
        case 'specHide': {
          var oh = st.ui.specIdx[+v];
          if (oh && oh.src === 'local') st.specLocal.add = st.specLocal.add.filter(function (x) { return x.id !== oh.id; });
          else if (oh && oh.key) st.specLocal.hide[oh.key] = true;
          saveSpecLocal(); break;
        }
        case 'specUnhide': { var ou = st.ui.specIdx[+v]; if (ou && ou.key) { delete st.specLocal.hide[ou.key]; saveSpecLocal(); } break; }
        case 'showHidden': st.ui.showHidden = !st.ui.showHidden; break;
        case 'specEditMode': st.ui.specEditMode = !st.ui.specEditMode; st.ui.specEdit = null; break;
        case 'specEdit': { var oe = st.ui.specIdx[+v]; if (oe) st.ui.specEdit = { i: +v, spec: oe.spec, place: oe.place, cat: oe.cat || '', src: oe.src, id: oe.id, key: oe.key, kind: oe.kind }; break; }
        case 'specCancel': st.ui.specEdit = null; break;
        case 'specSave': {
          var ed = st.ui.specEdit, sp = ed ? U.normSpec(ed.spec) : '';
          if (ed && sp) {
            var plc = ed.place === '（全箇所）' ? '（全箇所）' : w.place;
            if (ed.src === 'local') st.specLocal.add = st.specLocal.add.map(function (x) { return x.id === ed.id ? Object.assign({}, x, { spec: sp, place: plc, cat: ed.cat }) : x; });
            else { st.specLocal.hide[ed.key] = true; st.specLocal.add.push({ id: 'L' + Date.now(), kind: ed.kind || specKindOf(), place: plc, spec: sp, cat: ed.cat, from: ed.spec }); }
            saveSpecLocal(); st.ui.specEdit = null; pickSpec(sp, false);
          }
          break;
        }
        case 'specUse': if (st.ui.newSpec) { pickSpec(st.ui.newSpec, false); st.ui.newSpec = ''; } break;
        case 'specAdd': {
          var ns = U.normSpec(st.ui.newSpec);
          if (ns) {
            st.specLocal.add.push({ id: 'L' + Date.now(), kind: specKindOf(), place: st.ui.newPlace === 'all' ? '（全箇所）' : w.place, spec: ns, cat: st.ui.newCat });
            saveSpecLocal(); st.ui.newSpec = ''; pickSpec(ns, false);
            toast('仕様の候補に追加しました（このパソコンに記憶）', 'ok');
          }
          break;
        }
        case 'siteType': st.siteType = v || ''; docSet('uchiwake.siteType', st.siteType); break;
        // 計・経費
        case 'full': st.ui.full = v === '1'; saveUi(); break;
        case 'total': { var def = TOTALS.filter(function (t) { return t.id === v; })[0]; if (def) fresh(function () { runTotal(def); }); return; }
        case 'discount': {
          fresh(function () {
            var dp = UA.totalPlan(st.grid, targetRow(), 'discount', { amount: st.ui.discount });
            if (!dp.ok) { toast(dp.message, 'warn'); render(); return; }
            applyPlan({ ins: dp.ins ? { at: dp.insAt, n: dp.ins } : null, writes: totalWrites(dp), select: 'A' + dp.select }, '出精値引きと総合計を入れました（決定金額 ¥' + U.yen(dp.amount) + '）').then(render);
          });
          return;
        }
        case 'expMode': st.exp.mode = v; docSet('uchiwake.expense', { mode: st.exp.mode, cfg: st.exp.cfg }); break;
        case 'expRun': fresh(runExpense); return;
        case 'expDefault': save(K_EXP, { cfg: st.exp.cfg }); toast('今の率を既定にしました（このパソコン）', 'ok'); break;
        // 単価ルール
        case 'ruleOn': { var l1 = rules().map(function (x) { return Object.assign({}, x); }); l1[+v].on = !l1[+v].on; saveRules(l1); break; }
        case 'ruleUp': case 'ruleDown': {
          var l2 = rules().slice(), i2 = +v, j2 = a === 'ruleUp' ? i2 - 1 : i2 + 1;
          if (j2 >= 0 && j2 < l2.length) { var tmp = l2[i2]; l2[i2] = l2[j2]; l2[j2] = tmp; saveRules(l2); }
          break;
        }
        case 'ruleEdit': st.ui.ruleEdit = { i: +v, r: Object.assign({}, rules()[+v]) }; break;
        case 'ruleAdd': st.ui.ruleEdit = { i: -1, r: { id: 'R' + Date.now(), on: true, name: '', items: ['valve'], words: '', min: '', max: '', base: 'run', baseKind: '', factor: 1, round: 'up10', memo: '' } }; break;
        case 'reItem': { var re = st.ui.ruleEdit.r, its = (re.items || []).slice(), ix = its.indexOf(v); if (ix >= 0) its.splice(ix, 1); else its.push(v); re.items = its; break; }
        case 'reBase': st.ui.ruleEdit.r.base = v; break;
        case 'reCancel': st.ui.ruleEdit = null; break;
        case 'reDel': { var l3 = rules().slice(); l3.splice(st.ui.ruleEdit.i, 1); saveRules(l3); st.ui.ruleEdit = null; toast('ルールを削除しました', 'ok'); break; }
        case 'reSave': {
          var rr = st.ui.ruleEdit.r;
          if (!rr.name) rr.name = (rr.words || '').split(/[、,，\s]+/)[0] || '新しいルール';
          rr.factor = U.num(rr.factor) || 1;
          var l4 = rules().slice();
          if (st.ui.ruleEdit.i >= 0) l4[st.ui.ruleEdit.i] = rr; else l4.push(rr);
          saveRules(l4); st.ui.ruleEdit = null; toast('ルールを保存しました', 'ok');
          break;
        }
        case 'ruleReset': saveRules(UA.defaultCalcRules()); toast('初めの設定に戻しました', 'ok'); break;
        // そのほか
        case 'priceSel': readSelection().then(priceSelected); return;
        case 'undo': undoLast(); return;
        case 'closeToast': st.ui.toast = null; break;
        default: return;
      }
      render();
    });
    document.addEventListener('compositionstart', function () { st.ui.composing = true; });
    document.addEventListener('compositionend', function (ev) { st.ui.composing = false; onInput(ev); });
    document.addEventListener('input', function (ev) { if (ev.isComposing || st.ui.composing) { onInput(ev, true); return; } onInput(ev); });
    document.addEventListener('change', function (ev) {
      var el = ev.target, ff = el.dataset && el.dataset.f;
      if (ff === 'target') { st.target = el.value; handlers.sheet = null; loadModel().then(readSelection).then(render); return; }
      if (ff === 'specNewPlace') st.ui.newPlace = el.value;
      if (ff === 'specNewCat') st.ui.newCat = el.value;
      if (ff === 'specEditPlace' && st.ui.specEdit) st.ui.specEdit.place = el.value === 'all' ? '（全箇所）' : st.wiz.place;
      if (ff === 'specEditCat' && st.ui.specEdit) st.ui.specEdit.cat = el.value;
      if (ff === 're.round' && st.ui.ruleEdit) { st.ui.ruleEdit.r.round = el.value; render(); }
      if (ff === 'exp.on') { st.exp.cfg[el.dataset.k].on = el.checked; docSet('uchiwake.expense', { mode: st.exp.mode, cfg: st.exp.cfg }); render(); }
    });
    document.addEventListener('focusout', function () { setTimeout(function () { if (softPending && !isTyping()) render(); }, 0); });
    document.addEventListener('keydown', function (ev) {
      if (ev.key !== 'Enter' || ev.isComposing || ev.keyCode === 229 || st.ui.composing) return;
      var el = ev.target;
      if (!el || el.tagName !== 'INPUT' || el.type === 'checkbox') return;
      ev.preventDefault();
      var f = el.dataset && el.dataset.f;
      if (f === 'specNew' && st.ui.newSpec) { pickSpec(st.ui.newSpec, false); st.ui.newSpec = ''; render(); return; }
      if (f === 'discount') { el.blur(); return; }
      // 寸法の欄：Enter で次の欄へ
      var list = Array.prototype.slice.call(document.querySelectorAll('.dims input, .rf input, .exrows input'));
      var i = list.indexOf(el);
      if (i >= 0 && list[i + 1]) { list[i + 1].focus(); try { list[i + 1].select(); } catch (e) { /* noop */ } return; }
      el.blur();
    });
  }
  function onInput(ev, silent) {
    var el = ev.target;
    if (!el || !el.dataset || !el.dataset.f) return;
    var f = el.dataset.f, k = el.dataset.k, v = el.value, w = st.wiz;
    if (f === 'target' || el.type === 'checkbox') return;
    if (f === 'specNew') { st.ui.newSpec = v; if (!silent) render(); return; }
    if (f === 'specEditText') { if (st.ui.specEdit) st.ui.specEdit.spec = v; return; }
    if (f === 'specNewPlace' || f === 'specNewCat' || f === 'specEditPlace' || f === 'specEditCat' || f === 're.round') return;
    if (f === 'discount') {
      if (silent) return; // 日本語入力の変換中は触らない（確定したときに整える）
      var hv = toHalfDigits(v), caret = el.selectionStart == null ? hv.length : el.selectionStart;
      var before = toHalfDigits(v.slice(0, caret)).replace(/[^0-9]/g, '').length;
      var dg = hv.replace(/[^0-9]/g, '').replace(/^0+(?=[0-9])/, '').slice(0, 9);
      if (before > dg.length) before = dg.length;
      st.ui.discount = dg;
      var nv = commaDigits(dg);
      if (el.value !== nv) {
        el.value = nv;
        var pos = 0, seen = 0; while (pos < nv.length && seen < before) { if (nv.charAt(pos) !== ',') seen++; pos++; }
        try { el.setSelectionRange(pos, pos); } catch (e) { /* noop */ }
      }
    }
    else if (f === 'rRq') w.rect.rq = v;
    else if (f === 'rFx') w.rect.fx = v;
    else if (f === 'bFx') w.box.fx = v;
    else if (f === 'insN') w.insN = v;
    else if (f === 'bName') { w.box.name = v; w.box.five = null; }
    else if (f.indexOf('bDim.') === 0) w.box.boxes[+k][f.slice(5)] = v;
    else if (f.indexOf('re.') === 0 && st.ui.ruleEdit) st.ui.ruleEdit.r[f.slice(3)] = v;
    else if (f.indexOf('exp.') === 0) { st.exp.cfg[k][f.slice(4)] = v; docSet('uchiwake.expense', { mode: st.exp.mode, cfg: st.exp.cfg }); }
    if (!silent) render();
  }
})();
