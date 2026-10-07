/*
 * 内訳アシスト  taskpane.js
 * 作業ウィンドウの画面と、Excel（Office.js）への読み書き。
 */
(function () {
  'use strict';
  var U = UA.util;
  var VERSION = '1.3.0';
  var K_MASTER = 'uchiwake-assist.master.v1', K_MEMO = 'uchiwake-assist.pricememo.v1', K_SK = 'uchiwake-assist.sekisan.v1', K_SPECS = 'uchiwake-assist.speclocal.v1', K_UI = 'uchiwake-assist.ui.v1';
  var HOW_LABEL = { add: '下に追加', merge: '数量を足す', row: '行を指定' };
  var REF_HEAD = { y: '単価の参照元（保温積算資料）', z: '参照した仕様', aa: '参照したサイズ', ab: '参照した厚み' };

  var st = {
    master: null, masterInfo: null, sk: null, skInfo: null, specLocal: { add: [], hide: {}, pin: {} }, siteType: '', AL: {}, sheets: [], target: '', model: null, headerOK: false, modelErr: '',
    sel: freshSel(), ui: { specOpen: false, menu: false, busy: false, toast: null, composing: false, specEdit: null, showHidden: false, newSpec: '', newCat: '', newPlace: 'here', specIdx: [], how: 'add', atRow: '', elbowOn: false },
    undo: null, memo: { idx: {}, list: [] }, autoOpen: false, api19: false, api17: false, onChanged: null
  };
  function freshSel() {
    return { areaRow: null, area: '', newArea: false, noArea: false, newAreaName: '', sys: '', place: '', item: 'pipe', sub: '', t16: false, tSel: '',
      spec: '', specAuto: true, custom: '', ent: {}, vent: {}, rq: '', rf: undefined, ri: undefined, dims: [{ w: '', h: '', l: '' }], cmd: '' };
  }

  /* ---------- 保存（このパソコンの中） ---------- */
  function pkey(k) { try { return (Office.context.partitionKey ? Office.context.partitionKey + '.' : '') + k; } catch (e) { return k; } }
  function load(k) { try { var v = localStorage.getItem(pkey(k)); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
  function save(k, v) { try { localStorage.setItem(pkey(k), JSON.stringify(v)); return true; } catch (e) { return false; } }

  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function $(sel) { return document.querySelector(sel); }

  /* ---------- 起動 ---------- */
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) { $('#app').innerHTML = '<div class="boot">Excel で開いてください。</div>'; return; }
    try {
      st.api19 = Office.context.requirements.isSetSupported('ExcelApi', '1.9');
      st.api17 = Office.context.requirements.isSetSupported('ExcelApi', '1.7');
    } catch (e) { /* noop */ }
    try { st.autoOpen = Office.context.document.settings.get('Office.AutoShowTaskpaneWithDocument') === true; } catch (e) { /* noop */ }
    var raw = load(K_MASTER);
    if (raw && raw.master) { setMaster(raw.master, raw.info, false); }
    var memo = load(K_MEMO); if (memo && memo.idx) st.memo = memo;
    var sl = load(K_SPECS); if (sl && sl.add) st.specLocal = { add: sl.add || [], hide: sl.hide || {}, pin: sl.pin || {} };
    try { st.siteType = Office.context.document.settings.get('uchiwake.siteType') || ''; } catch (e) { /* noop */ }
    var skRaw = load(K_SK); if (skRaw && skRaw.sk && skRaw.sk.tables) { st.sk = skRaw.sk; st.skInfo = skRaw.info; }
    var uiRaw = load(K_UI); if (uiRaw && (uiRaw.how === 'add' || uiRaw.how === 'merge')) st.ui.how = uiRaw.how;
    st.ui.howSaved = st.ui.how;
    if (uiRaw && uiRaw.elbowOn) st.ui.elbowOn = true;
    bindEvents();
    $('#masterFile').addEventListener('change', onMasterFile);
    $('#skFile').addEventListener('change', onSkFile);
    refreshSheets().then(render, function (e) { st.modelErr = msg(e); render(); });
    if (st.api17) {
      Excel.run(function (ctx) {
        ctx.workbook.worksheets.onActivated.add(function () { return refreshSheets().then(render); });
        return ctx.sync();
      }).catch(function () { /* noop */ });
    }
  });

  function msg(e) { return (e && (e.message || e.debugInfo && e.debugInfo.message)) || String(e); }

  /* ---------- 設定マスタ ---------- */
  function setMaster(m, info, persist) {
    st.master = UA.indexMaster(m);
    st.masterInfo = info;
    st.AL = UA.buildAliases(st.master);
    if (persist) {
      var copy = Object.assign({}, m); delete copy.priceIdx; delete copy.specCount; delete copy.thickIdx; delete copy.kindById; delete copy.skRules; delete copy.skAdj;
      if (!save(K_MASTER, { master: copy, info: info })) toast('マスタをこのパソコンに記憶できませんでした（容量不足の可能性）。表示中は使えます。', 'warn');
    }
    normalizeSel();
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
        setMaster(m, { name: f.name, at: U.today(), prices: m.prices.length }, true);
        st.sel.spec = ''; st.sel.specAuto = true; normalizeSel();
        toast('設定マスタを読み込みました（' + f.name + '、単価の記録 ' + m.prices.length + '行）' + (m.errors.length ? '　注意：' + m.errors.join('／') : ''), m.errors.length ? 'warn' : 'ok');
        render();
      } catch (err) { toast('読み込みに失敗しました：' + msg(err), 'err'); render(); }
    };
    reader.readAsArrayBuffer(f);
  }

  /* ---------- 保温積算資料（単価表） ---------- */
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
  // 単価は保温積算資料（協会単価）だけから取る。表に無いものは空欄（過去の見積・手入力の記録は使わない）
  function priceOrder() { return ['sk']; }
  function seriesOf() { var k = kind(); if (!k) return ''; return st.sel.item === 'pipe' && st.sel.sub === 'SUS' && st.master.sizes.Su ? 'Su' : k.series; }
  function sizeDef(n) { return (st.master.sizes[seriesOf()] || []).filter(function (z) { return z.n === n; })[0] || null; }
  function sizeText(size) { var ser = seriesOf(); return ser === 'A' ? size + 'A' : ser === 'Su' ? size + 'Su' : ser === 'φ' ? 'φ' + size : String(size); }
  // 単価と、その参照元（Y〜AB列に書く内容）
  function priceOf(d, spec, size, f) {
    var s = st.sel, m = st.master, look = size, conv = '';
    if (seriesOf() === 'Su') {
      var z = sizeDef(U.num(size));
      if (z && z.priceSize) { look = z.priceSize; conv = size + 'Su（' + z.priceSize + 'A相当）'; }
      else return { p: null, cap: ['Suの読み替えなし', 'c-err'], ref: { y: '単価なし：' + size + 'Su に当たる鋼管のサイズが M_サイズ にありません', z: '', aa: '', ab: '' }, tip: 'M_サイズ の「単価に使うサイズ」' };
    }
    if (!st.sk) return { p: null, cap: ['保温積算資料が未読込', 'c-err'], ref: { y: '単価なし：保温積算資料が未読込', z: '', aa: '', ab: '' }, tip: '保温積算資料が未読込' };
    var skr = UA.skLook(st.sk, m, { item: s.item, d: d, place: s.place, spec: spec, size: look, f: f });
    if (skr.p != null) {
      var ref = Object.assign({}, skr.ref); if (conv) ref.aa = conv;
      return { p: skr.p, iF: skr.formula, cap: ['積算資料 §' + skr.id + (skr.approx ? '（近い表）' : ''), skr.approx ? 'c-warn' : 'c-info'], ref: ref, tip: ref.y + (conv ? '　' + conv : '') };
    }
    var why = skr.why && skr.why !== 'nomatch' ? skr.why : '対応する表がありません（M_積算資料対応）';
    return { p: null, cap: ['協会単価なし', 'c-err'], ref: { y: '単価なし：' + why, z: '', aa: '', ab: '' }, tip: why };
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
  function maxColOf() { var c = colsOf(), mx = 1; Object.keys(c).forEach(function (k) { mx = Math.max(mx, U.colNum(c[k])); }); return U.colLetter(mx); }

  function loadModel() {
    if (!st.target) { st.model = null; st.headerOK = false; return Promise.resolve(); }
    var cols = colsOf(), maxCol = maxColOf(), target = st.target;
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(target);
      var ur = ws.getUsedRange(false); ur.load('rowIndex,rowCount');
      return ctx.sync().then(function () {
        var last = Math.min(Math.max(ur.rowIndex + ur.rowCount, startRowOf() + 1), 6000);
        var rg = ws.getRange('A1:' + maxCol + last); rg.load('values,formulas');
        return ctx.sync().then(function () {
          var model = UA.parseSheet(rg.values, rg.formulas, cols, startRowOf());
          var fIdx = U.colNum(cols.f) - 1;
          st.headerOK = String((rg.values[0] || [])[fIdx] || '').replace(/[\s　]/g, '').indexOf('保温厚') >= 0;
          model.jTpl = detectTpl(model, rg.formulas, cols);
          st.model = model; st.modelErr = '';
          attachChanged(target);
          normalizeSel();
        });
      });
    }).catch(function (e) { st.model = null; st.modelErr = msg(e); });
  }
  // 既存の金額式（J列）のくせを、行番号を {行} にして覚える
  function detectTpl(model, formulas, cols) {
    var jIdx = U.colNum(cols.j) - 1, cnt = { line: {}, blank: {} };
    model.recs.forEach(function (r) {
      if (r.t !== 'line' && r.t !== 'blank') return;
      var f = (formulas[r.row - 1] || [])[jIdx];
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
  var changedTimer = null;
  function attachChanged(target) {
    if (!st.api17 || st.onChangedFor === target) return;
    var old = st.onChanged;
    st.onChangedFor = target;
    var removeOld = old ? Excel.run(old.context, function (ctx) { old.remove(); return ctx.sync(); }).catch(function () { /* noop */ }) : Promise.resolve();
    removeOld.then(function () { return Excel.run(function (ctx) {
      st.onChanged = ctx.workbook.worksheets.getItem(target).onChanged.add(function () {
        if (st.ui.busy) return Promise.resolve();
        clearTimeout(changedTimer);
        changedTimer = setTimeout(function () { loadModel().then(render); }, 900);
        return Promise.resolve();
      });
      return ctx.sync();
    }); }).catch(function () { st.onChangedFor = null; });
  }

  /* ---------- 選択の整え ---------- */
  function kinds() { return st.master ? st.master.kinds.filter(function (k) { return k.use; }) : []; }
  function kind() { return st.master && st.master.kindById[st.sel.item]; }
  function subsOf(item) { return st.master ? (st.master.subs[item] || []).filter(function (s) { return s.use; }) : []; }
  function areas() { return st.model ? UA.areasOf(st.model) : []; }
  function curAreaName() { var s = st.sel; if (s.noArea && !s.newArea) return ''; if (s.newArea) return areaNumNext() + '.' + stripNum(s.newAreaName); var a = areas().filter(function (x) { return x.row === s.areaRow; })[0]; return a ? a.name : ''; }
  // 大項目の番号は自動（シート上の最大の番号＋1）
  function areaNumNext() { var n = 0; areas().forEach(function (a) { var mm = U.toHalf(a.name).match(/^\s*(\d+)\s*[.．、]/); if (mm) n = Math.max(n, +mm[1]); }); return n + 1; }
  function stripNum(t) { return String(t == null ? '' : t).replace(/^\s*[0-9０-９]+\s*[.．、]\s*/, ''); }
  function famOfSel() { var n = curAreaName(); if (n) return UA.famOfArea(n); var k = kind(); return k ? k.fam : 'pipe'; }
  function sysList() {
    var s = st.sel, fam = famOfSel(), out = [];
    if (!st.master) return out;
    if (s.noArea && !s.newArea) return [''];
    var hasNone = st.master.systems.some(function (x) { return x.name === '' && x.use && x.fam === fam; });
    if (hasNone) out.push('');
    if (!s.newArea && s.areaRow) UA.sysOf(st.model, s.areaRow).forEach(function (b) { if (out.indexOf(b) < 0) out.push(b); });
    st.master.systems.forEach(function (x) { if (x.use && x.fam === fam && out.indexOf(x.name) < 0) out.push(x.name); });
    return out;
  }
  function placeList() { return st.master ? st.master.places.filter(function (p) { return p.use; }).map(function (p) { return p.name; }) : []; }
  function nextAreaName(fam) { return fam === 'duct' ? 'ダクト設備' : '配管設備'; }
  function defaultSub(item) {
    var subs = subsOf(item);
    if (!subs.length) return '';
    if (item === 'rect') { var sa = subs.filter(function (x) { return x.id === 'SAダクト'; })[0]; if (sa) return sa.id; }
    return subs[0].id;
  }
  function normalizeSel(patch) {
    var s = st.sel, prev = Object.assign({}, s);
    if (patch) Object.assign(s, patch);
    if (!st.master) return;
    if (!kind() || !kind().use) s.item = (kinds()[0] || {}).id || 'pipe';
    if (s.item !== prev.item || (patch && 'item' in patch)) {
      if (!patch || !('sub' in patch)) s.sub = defaultSub(s.item);
      if (!patch || !('ent' in patch)) s.ent = {};
      if (!patch || !('rq' in patch)) s.rq = '';
      if (!patch || !('vent' in patch)) s.vent = {};
      s.dims = [{ w: '', h: '', l: '' }]; if (!patch || !('t16' in patch)) s.t16 = false; if (!patch || !('tSel' in patch)) s.tSel = '';
      var kf = kind() ? kind().fam : 'pipe';
      if (!s.noArea && (!patch || (!('areaRow' in patch) && !('newArea' in patch))) && famOfSel() !== kf) {
        var alt = areas().filter(function (a) { return UA.famOfArea(a.name) === kf; })[0];
        if (alt) { s.areaRow = alt.row; s.newArea = false; }
      }
    }
    var subs = subsOf(s.item);
    if (!subs.length) s.sub = '';
    else if (!subs.some(function (x) { return x.id === s.sub; })) s.sub = defaultSub(s.item);
    // 管の種類（鋼管／SUS）を変えたらサイズの系列が変わるので選び直し
    if (s.item === 'pipe' && s.sub !== prev.sub && prev.item === 'pipe' && (!patch || !('ent' in patch))) s.ent = {};
    var al = areas();
    if (!s.newArea && !s.noArea && !al.some(function (a) { return a.row === s.areaRow; })) {
      if (al.length) { var kf2 = kind() ? kind().fam : 'pipe'; var pick = al.filter(function (a) { return UA.famOfArea(a.name) === kf2; })[0] || al[0]; s.areaRow = pick.row; }
      else { s.newArea = true; }
    }
    if (s.newArea && !s.newAreaName) s.newAreaName = nextAreaName(kind() ? kind().fam : 'pipe');
    var sl = sysList();
    if (sl.indexOf(s.sys) < 0) s.sys = sl.length ? sl[0] : '';
    var pl = placeList();
    if (pl.indexOf(s.place) < 0) s.place = pl[0] || '';
    if (s.specAuto || !s.spec) { s.spec = autoSpec(); if (!s.custom) s.specAuto = true; }
    if (s.spec !== prev.spec || s.sub !== prev.sub || s.item !== prev.item || s.t16 !== prev.t16) {
      var e2 = {}; Object.keys(s.ent).forEach(function (k) { e2[k] = { g: s.ent[k].g }; if (s.tSel) e2[k].f = s.tSel; }); s.ent = e2; s.rf = s.tSel || undefined; s.ri = undefined;
      if (s.spec !== prev.spec) { var v2 = {}; Object.keys(s.vent || {}).forEach(function (k) { v2[k] = {}; }); s.vent = v2; }
    }
  }
  function sheetSpec() {
    var s = st.sel;
    if (!st.model || s.newArea) return null;
    var isV = s.item === 'valve' || s.item === 'flange';
    if (s.noArea) return tailSpec(s.item) || (isV ? tailSpec('pipe') : null);
    // バルブ・フランジは、同じ施工箇所の配管（直管）の仕様をそのまま使う
    return UA.specInSheet(st.model, st.master, curAreaName(), s.areaRow, s.sys, s.place, s.item, dOf()) ||
      (isV ? UA.specInSheet(st.model, st.master, curAreaName(), s.areaRow, s.sys, s.place, 'pipe') : null);
  }
  // 大項目なしのとき：最後の大項目より後ろで、同じ施工箇所に使っている仕様
  function tailSpec(item) {
    var recs = st.model.recs, s0 = 0, i, cur = '', spec = null, found = null;
    for (i = 0; i < recs.length; i++) if (recs[i].t === 'area') s0 = i + 1;
    for (i = s0; i < recs.length; i++) {
      var r = recs[i];
      if (r.t === 'spec' || (r.t === 'sys' && r.e)) { if (r.c) cur = r.c; spec = U.trimAll(cur) === U.trimAll(st.sel.place) ? r.e : null; continue; }
      if (spec && UA.lineItem(r, 'pipe', st.master) === item) found = spec;
    }
    return found;
  }
  function blockSel() { var s = st.sel; return { noArea: s.noArea, newArea: s.newArea, area: curAreaName(), areaRow: s.areaRow, sys: s.sys, place: s.place, spec: s.spec }; }
  // 配管（直管）の保温厚：この塊の同じサイズの直管 → 無ければ M_標準厚
  function pipeThickFor(size) {
    var hit = UA.blockLines(st.model, blockSel()).filter(function (r) { return !UA.isPieceLine(r) && String(r.e) === String(size) && r.f; })[0];
    return hit ? { t: hit.f, from: '内訳の' + size + (hit.d === 'SUS' ? 'Su' : 'A') + '（' + hit.row + '行目）' } : { t: UA.defThick(st.master, 'pipe', st.sel.spec, size), from: '配管の標準厚（M_標準厚）' };
  }
  function valveThickFor(size) {
    var pt = pipeThickFor(size);
    return { f: UA.valveThick(pt.t, st.master.output && st.master.output.valveThick), pipe: pt.t, from: pt.from };
  }
  function specGroups() {
    var s = st.sel;
    return UA.specList(st.master, { item: s.item, place: s.place, fam: kind() ? kind().fam : 'pipe', sheet: st.model ? UA.sheetSpecs(st.model) : [], cur: sheetSpec(), cat: st.siteType, local: st.specLocal, showHidden: st.ui.showHidden });
  }
  // 自動で選ぶ仕様：今の塊 → この見積の同じ施工箇所 → 候補（固定・この施工箇所） → その他
  function autoSpec() {
    var all = [];
    specGroups().forEach(function (g) { if (g.key !== 'hidden') all = all.concat(g.items); });
    var pick = all.filter(function (x) { return x.cur; })[0] || all.filter(function (x) { return x.sheet && x.here; })[0] ||
      all.filter(function (x) { return !x.sheet && (x.here || x.pin); })[0] || all.filter(function (x) { return !x.sheet; })[0] || all[0];
    return pick ? pick.spec : (st.sel.spec || '');
  }
  function dOf() {
    var s = st.sel, d = '';
    if (s.item === 'valve' || s.item === 'flange' || s.item === 'pipe') d = s.sub;
    else if (s.item === 'rect' || s.item === 'round') d = s.sub ? s.sub + (s.t16 ? '(1.6t)' : '') : '';
    return UA.normalizeOut(st.master, U.colNum(colsOf().d) ? colsOf().d : 'D', d);
  }
  function unitOf() { var k = kind(); return UA.normalizeOut(st.master, colsOf().h, k ? k.unit : ''); }
  function dimsTotal() { var t = 0; st.sel.dims.forEach(function (d) { var w = U.num(d.w), h = U.num(d.h), l = U.num(d.l); if (w > 0 && h > 0 && l > 0) t += 2 * (w + h) / 1000 * l; }); return Math.round(t * 10) / 10; }

  /* ---------- 明細の組み立てと計画 ---------- */
  function buildLines() {
    var s = st.sel, k = kind(), out = [];
    if (!k) return out;
    if (s.item === 'valve' || s.item === 'flange') return valveLines();
    var d = dOf(), unit = unitOf();
    if (k.series) {
      (st.master.sizes[seriesOf()] || []).forEach(function (z) {
        var key = String(z.n), en = s.ent[key];
        if (!en) return;
        var g = U.num(en.g); if (!(g > 0)) return;
        var f = (en.f !== undefined && en.f !== '') ? U.thickKey(en.f) : UA.defThick(st.master, s.item, s.spec, z.n);
        var pr = priceOf(d, s.spec, z.n, f);
        var manual = en.i !== undefined;
        var i = manual ? (en.i === '' ? null : U.num(en.i)) : pr.p;
        out.push({ d: d, e: z.n, f: f, g: g, h: unit, i: i, iF: manual ? null : pr.iF, ref: manual ? (en.i === '' ? null : { y: '手入力', z: '', aa: '', ab: '' }) : pr.ref, manual: manual && en.i !== '', pkey: pr.key, auto: pr.p });
      });
    } else {
      var eVal = s.item === 'rect' ? '矩形' : (s.sub || 'BOX');
      var g2 = s.rq !== '' ? U.num(s.rq) : dimsTotal();
      if (g2 > 0) {
        var f2 = (s.rf !== undefined && s.rf !== '') ? U.thickKey(s.rf) : UA.defThick(st.master, s.item, s.spec, eVal);
        var pr2 = priceOf(d, s.spec, eVal, f2);
        var man2 = s.ri !== undefined;
        out.push({ d: s.item === 'rect' ? d : '', e: eVal, f: f2, g: g2, h: unit, i: man2 ? (s.ri === '' ? null : U.num(s.ri)) : pr2.p, iF: man2 ? null : pr2.iF, ref: man2 ? (s.ri === '' ? null : { y: '手入力', z: '', aa: '', ab: '' }) : pr2.ref, manual: man2 && s.ri !== '', pkey: pr2.key, auto: pr2.p });
      }
    }
    return out;
  }
  // バルブ・フランジ：種類とサイズを選ぶだけ。数量は空欄（内訳で直接入力）。保温厚は配管から、単価は協会単価
  function ventKey(sub, n) { return sub + '|' + n; }
  function ventList() {
    var s = st.sel, out = [];
    subsOf(s.item).forEach(function (sb) {
      (st.master.sizes.A || []).forEach(function (z) { var key = ventKey(sb.id, z.n); if (s.vent[key]) out.push({ key: key, sub: sb, z: z, en: s.vent[key] }); });
    });
    return out;
  }
  function ventRow(v) {
    var s = st.sel, d = UA.normalizeOut(st.master, 'D', v.sub.id), vt = valveThickFor(v.z.n);
    var f = (v.en.f !== undefined && v.en.f !== '') ? U.thickKey(v.en.f) : vt.f;
    var pr = priceOf(d, s.spec, v.z.n, f), manual = v.en.i !== undefined;
    return { d: d, e: v.z.n, f: f, vt: vt, pr: pr, manual: manual, i: manual ? (v.en.i === '' ? null : U.num(v.en.i)) : pr.p };
  }
  function valveLines() {
    var unit = unitOf();
    return ventList().map(function (v) {
      var r = ventRow(v);
      return { d: r.d, e: r.e, f: r.f, g: '', h: unit, i: r.i, iF: r.manual ? null : r.pr.iF, ref: r.manual ? (r.i == null ? null : { y: '手入力', z: '', aa: '', ab: '' }) : r.pr.ref, manual: r.manual && r.i != null, auto: r.pr.p };
    });
  }
  function currentPlan() {
    if (!st.master || !st.model) return null;
    var s = st.sel, o = st.master.output || {};
    return UA.plan(st.model, { area: curAreaName(), areaRow: s.newArea ? null : s.areaRow, newArea: s.newArea, noArea: s.noArea && !s.newArea, sys: s.sys, place: s.place, spec: s.spec, item: s.item, lines: buildLines(),
      elbow: s.item === 'pipe' && st.ui.elbowOn, elbowRate: o.elbowRate || 0.35, elbowLabel: o.elbowLabel || 'エルボ', totals: !!o.makeTotals,
      how: st.ui.how, atRow: st.ui.atRow }, st.master);
  }
  function saveUi() { save(K_UI, { how: st.ui.how === 'row' ? (st.ui.howSaved || 'add') : st.ui.how, elbowOn: st.ui.elbowOn }); }

  /* ---------- 登録（Excel への書き込み） ---------- */
  function register() {
    if (st.ui.busy) return;
    if (!st.headerOK) { toast('書き込み先が内訳シートではないようです（1行目に「保温厚」の見出しがありません）', 'err'); render(); return; }
    st.ui.busy = true; render();
    var lines = buildLines();
    loadModel().then(function () {
      var p = currentPlan();
      if (!p || !p.ok) throw new Error(p ? p.message : '準備ができていません');
      var cols = colsOf(), maxCol = maxColOf(), jc = cols.j, tpl = st.model.jTpl;
      var undo = { sheet: st.target, cells: [], rows: [] };
      var rc = refColsOf(), rk = ['y', 'z', 'aa', 'ab'].filter(function (k) { return rc[k]; });
      var refLines = p.news.filter(function (x) { return x.t === 'line' && x.ref; });
      return Excel.run(function (ctx) {
        var ws = ctx.workbook.worksheets.getItem(st.target);
        // Y〜AB列の見出し（1行目が空のときだけ書く）
        var hdr = refLines.length ? rk.map(function (k) { var c = ws.getRange(rc[k] + '1'); c.load('values'); return { k: k, c: c }; }) : [];
        return ctx.sync().then(function () {
        hdr.forEach(function (h) {
          if (!U.isEmpty(h.c.values[0][0])) return;
          undo.cells.push({ addr: rc[h.k] + '1', v: '' });
          h.c.values = [[REF_HEAD[h.k]]];
        });
        p.totalOps.forEach(function (op) { undo.cells.push({ addr: jc + op.origRow, v: op.prev }); });
        p.cellOps.forEach(function (op) {
          var a = cols[op.col] + op.row;
          undo.cells.push({ addr: a, v: op.prev == null ? '' : op.prev });
          ws.getRange(a).values = [[op.value]];
        });
        p.inserts.forEach(function (g) { ws.getRange(g.at + ':' + (g.at + g.count - 1)).insert(Excel.InsertShiftDirection.down); });
        return ctx.sync().then(function () {
          // 書式：既存の明細行・計の行からコピー
          var lineSrc = null, totalSrc = null;
          p.W.forEach(function (x) { if (!x.isNew && x.t === 'line' && !lineSrc) lineSrc = x.finalRow; if (!x.isNew && x.t === 'total' && !totalSrc) totalSrc = x.finalRow; });
          p.news.forEach(function (x) {
            var r = x.finalRow, rg = ws.getRange('A' + r + ':' + maxCol + r);
            if (st.api19 && x.t === 'total' && totalSrc) rg.copyFrom('A' + totalSrc + ':' + maxCol + totalSrc, Excel.RangeCopyType.formats);
            else if (st.api19 && lineSrc) rg.copyFrom('A' + lineSrc + ':' + maxCol + lineSrc, Excel.RangeCopyType.formats);
            else rg.format.font.bold = false;
            if (x.t === 'total' && !(st.api19 && totalSrc)) ws.getRange(cols.label + r + ':' + jc + r).format.font.bold = true;
            rg.formulas = [rowArray(x, r, cols, maxCol, tpl)];
            // 単価の参照元（Y〜AB列）
            if (x.t === 'line' && x.ref) rk.forEach(function (k) { ws.getRange(rc[k] + r).values = [[x.ref[k] == null ? '' : x.ref[k]]]; });
            // 仕様（括弧書き）は左揃え、サイズ・エルボは右揃え、D列（種別）は中央（コピーした書式に左右されないように）
            if (x.t === 'spec') ws.getRange(cols.spec + r).format.horizontalAlignment = 'Left';
            if (x.t === 'line' || x.t === 'elbow') ws.getRange(cols.size + r).format.horizontalAlignment = 'Right';
            if (x.t === 'line') ws.getRange(cols.d + r).format.horizontalAlignment = 'Center';
          });
          p.totalOps.forEach(function (op) { if (op.type === 'set') ws.getRange(jc + op.finalRow).formulas = [[op.formula]]; });
          var appends = p.totalOps.filter(function (op) { return op.type === 'append'; }).map(function (op) {
            var c = ws.getRange(jc + op.finalRow); c.load('formulas'); return { op: op, c: c };
          });
          return ctx.sync().then(function () {
            appends.forEach(function (a) {
              var f = String(a.c.formulas[0][0] || '');
              if (/\)\s*$/.test(f)) a.c.formulas = [[f.replace(/\)\s*$/, ',' + a.op.ref + ')')]];
            });
            ws.getRange('A' + p.firstRow + ':' + jc + p.lastRow).select();
            return ctx.sync();
          });
        });
        });
      }).then(function () {
        undo.rows = p.inserts.map(function (g) { return { start: g.recs[0].finalRow, count: g.count }; });
        st.undo = undo;
        var s = st.sel;
        s.ent = {}; s.vent = {}; s.rq = ''; s.rf = undefined; s.ri = undefined; s.dims = [{ w: '', h: '', l: '' }]; s.cmd = '';
        if (s.newArea) { s.pendingArea = curAreaName(); s.newArea = false; s.areaRow = null; }
        // 行を指定して入れたときは、次はその続き（入れた明細のすぐ下）を指す
        if (p.how === 'row') {
          var lastLine = 0; p.news.forEach(function (x) { if ((x.t === 'line' || x.t === 'elbow') && x.finalRow > lastLine) lastLine = x.finalRow; });
          if (lastLine) st.ui.atRow = String(lastLine + 1);
        }
        var t = (p.addCount ? '明細' + p.addCount + '行を追加' : '') + (p.addCount && p.mergeCount ? '・' : '') + (p.mergeCount ? p.mergeCount + '行の数量に加算' : '') +
          'しました（' + (p.firstRow === p.lastRow ? p.firstRow + '行目' : p.firstRow + '〜' + p.lastRow + '行目') + '）' + (p.how === 'row' ? '。次は ' + st.ui.atRow + '行目に入れます' : '');
        return loadModel().then(function () {
          if (s.pendingArea) { var a = areas().filter(function (x) { return x.name === s.pendingArea; }).pop(); if (a) s.areaRow = a.row; delete s.pendingArea; normalizeSel(); }
          toast(t, 'ok', true);
        });
      });
    }).catch(function (e) {
      toast('登録できませんでした：' + msg(e) + (/protect|保護/i.test(msg(e)) ? '（シートの保護を解除してください）' : ''), 'err');
    }).then(function () { st.ui.busy = false; render(); });
  }
  function numOrText(v) { var t = U.toHalf(v); return /^\d+(\.\d+)?$/.test(t) ? parseFloat(t) : v; }
  function rowArray(x, r, cols, maxCol, tpl) {
    var n = U.colNum(maxCol), arr = []; for (var i = 0; i < n; i++) arr.push('');
    function put(k, v) { arr[U.colNum(cols[k]) - 1] = v; }
    var jf = function (t) { return t ? t.replace(/\{行\}/g, r) : ''; };
    // 仕様の行（括弧書き）・大項目・系統の行の J列は式を入れない
    if (x.t !== 'spec' && x.t !== 'area' && x.t !== 'sys') put('j', jf(tpl.blank));
    if (x.t === 'area') put('area', x.a);
    else if (x.t === 'sys') put('sys', x.b);
    else if (x.t === 'spec') { if (x.c) put('place', '(' + x.c + ')'); if (x.d) put('d', x.d); put('spec', '(' + x.e + ')'); }
    else if (x.t === 'line') {
      put('d', x.dShow || ''); put('size', x.e); put('f', numOrText(x.f)); put('g', x.g); put('h', x.h); put('i', x.iF || (x.i == null ? '' : x.i)); put('j', jf(tpl.line));
    } else if (x.t === 'elbow') { put('size', x.e || 'エルボ'); put('g', 1); put('h', '式'); put('j', x.formula || ''); }
    else if (x.t === 'total') { put('label', x.label); put('j', x.formula); }
    return arr;
  }
  function rememberPrices(lines) {
    var s = st.sel, k = kind(), changed = false;
    lines.forEach(function (l) {
      if (!l.manual || l.i == null || l.i === l.auto) return;
      st.memo.idx[l.pkey] = { p: l.i, date: U.today() };
      st.memo.list.push({ item: s.item, kind: UA.KIND_OF_ITEM[s.item], d: l.d, spec: s.spec, size: l.e, f: l.f, unit: l.h, price: l.i, date: U.today(), sys: s.sys, place: s.place, src: '内訳アシスト手入力' });
      changed = true;
    });
    if (changed) save(K_MEMO, st.memo);
    return k;
  }
  function undoLast() {
    var u = st.undo;
    if (!u || st.ui.busy) return;
    st.ui.busy = true; render();
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(u.sheet);
      u.rows.slice().sort(function (a, b) { return b.start - a.start; }).forEach(function (g) {
        ws.getRange(g.start + ':' + (g.start + g.count - 1)).delete(Excel.DeleteShiftDirection.up);
      });
      return ctx.sync().then(function () {
        u.cells.forEach(function (c) { ws.getRange(c.addr).formulas = [[c.v == null ? '' : c.v]]; });
        return ctx.sync();
      });
    }).then(function () { st.undo = null; return loadModel(); })
      .then(function () { toast('元に戻しました', 'ok'); })
      .catch(function (e) { toast('元に戻せませんでした：' + msg(e), 'err'); })
      .then(function () { st.ui.busy = false; render(); });
  }
  function locate() {
    var p = currentPlan();
    if (!p || !p.ok || !p.firstRow) return;
    var jc = colsOf().j, a = p.firstRow, b = p.lastRow || p.firstRow;
    Excel.run(function (ctx) { ctx.workbook.worksheets.getItem(st.target).getRange('A' + a + ':' + jc + b).select(); return ctx.sync(); })
      .then(function () {
        toast(planWords(p) + '。シートはまだ変わっていません。', 'info');
        render();
      })
      .catch(function (e) { toast(msg(e), 'err'); render(); });
  }
  // 登録で何が起きるかを文で（位置を確認・登録先の説明に使う）
  function planWords(p) {
    var newRows = p.news.map(function (x) { return x.finalRow; }).sort(function (x, y) { return x - y; });
    var mg = p.merges.map(function (x) { return x.finalRow; });
    var parts = [];
    if (newRows.length) {
      var top = newRows[0], bot = newRows[newRows.length - 1];
      var first = p.inserts.length ? Math.min.apply(null, p.inserts.map(function (g) { return g.at; })) : top;
      parts.push((top === bot ? top + '行目' : top + '〜' + bot + '行目') + 'に新しい行を入れます（今の' + first + '行目から下は下にずれます）');
    }
    if (mg.length) parts.push(mg.join('・') + '行目は数量を足します（今ある行の数量が変わります）');
    return parts.join('。');
  }
  // Excel で選択しているセルの行を「行を指定」に使う（above=その行に入れる、below=その行の下に入れる）
  function pickRow(below) {
    Excel.run(function (ctx) {
      var r = ctx.workbook.getSelectedRange(); r.load('rowIndex,rowCount');
      var ws = r.worksheet; ws.load('name');
      return ctx.sync().then(function () { return { top: r.rowIndex + 1, bottom: r.rowIndex + r.rowCount, sheet: ws.name }; });
    }).then(function (o) {
      if (o.sheet !== st.target) { toast('「' + st.target + '」のシートでセルを選んでから押してください（今は「' + o.sheet + '」）', 'warn'); render(); return; }
      st.ui.how = 'row';
      st.ui.atRow = String(below ? o.bottom + 1 : o.top);
      render();
    }).catch(function (e) { toast(msg(e), 'err'); render(); });
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
  function exportMemo() {
    if (!st.memo.list.length) { toast('手入力した単価の記録はまだありません', 'warn'); render(); return; }
    var name = '内訳アシスト_単価メモ';
    var head = ['品目ID', '仕様区分', '種別（D列）', '仕様（E列・括弧なし）', 'サイズ（E列）', '保温厚（F列）', '単位（H列）', '単価（I列）', '見積日', '系統（B列）', '施工箇所', '施工箇所（元の表記）', '出典（見積書ファイル）'];
    var rows = st.memo.list.map(function (m) { return [m.item, m.kind, m.d, m.spec, m.size, numOrText(m.f), m.unit, m.price, m.date, m.sys, m.place, m.place, m.src]; });
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject(name);
      return ctx.sync().then(function () {
        if (ws.isNullObject) ws = ctx.workbook.worksheets.add(name); else ws.getRange().clear();
        ws.getRange('A1:M1').values = [head];
        ws.getRange('A1:M1').format.font.bold = true;
        ws.getRange('A2:M' + (rows.length + 1)).values = rows;
        ws.activate();
        return ctx.sync();
      });
    }).then(function () { toast('「' + name + '」シートに書き出しました。2行目以降を設定マスタの M_単価履歴 の下に貼り付けてください。', 'ok'); render(); })
      .catch(function (e) { toast(msg(e), 'err'); render(); });
  }

  // このパソコンで変えた仕様候補を、設定マスタ（M_仕様候補）へ移すための一覧に書き出す
  function exportSpecs() {
    var L = st.specLocal, rows = [];
    L.add.forEach(function (a) { rows.push(['追加', a.kind, a.place, a.spec, a.cat || '', a.from ? '「' + a.from + '」を直したもの' : '']); });
    Object.keys(L.hide).forEach(function (k) { var p = k.split('|'); rows.push(['候補から外す', p[0], p[1], p.slice(2).join('|'), '', '設定マスタの「使用する」を × に']); });
    Object.keys(L.pin).forEach(function (k) { var p = k.split('|'); rows.push(['固定', p[0], p[1], p.slice(2).join('|'), '', '設定マスタの「優先順位」に 1 など']); });
    if (!rows.length) { toast('仕様候補の変更はまだありません', 'warn'); render(); return; }
    var name = '内訳アシスト_仕様候補';
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject(name);
      return ctx.sync().then(function () {
        if (ws.isNullObject) ws = ctx.workbook.worksheets.add(name); else ws.getRange().clear();
        ws.getRange('A1:F1').values = [['操作', '仕様区分', '施工箇所', '仕様（括弧なし）', '区分（民間・官庁）', 'メモ']];
        ws.getRange('A1:F1').format.font.bold = true;
        ws.getRange('A2:F' + (rows.length + 1)).values = rows;
        ws.activate();
        return ctx.sync();
      });
    }).then(function () { toast('「' + name + '」シートに書き出しました。「追加」の行は設定マスタの M_仕様候補 に貼り付けると、ほかのパソコンでも使えます。', 'ok'); render(); })
      .catch(function (e) { toast(msg(e), 'err'); render(); });
  }

  /* ---------- スマート入力 ---------- */
  function onCmd(text) {
    var s = st.sel, r = UA.parseCmd(text, st.AL), patch = { cmd: text };
    var item = r.item || s.item;
    var km = st.master.kindById;
    if (!r.item && r.duct && km[item] && km[item].fam !== 'duct') item = 'rect';
    if (!r.item && r.sizes.length) {
      var ser = r.sizes[0].s;
      if (ser === 'A' && (!km[item] || km[item].series !== 'A')) item = 'pipe';
      if (ser === 'φ' && item !== 'round') item = 'round';
    }
    if (!r.item && r.rq && km[item] && km[item].series) item = 'rect';
    if (item !== s.item) patch.item = item;
    if (r.sub !== undefined) patch.sub = r.sub; else if (r.duct) patch.sub = r.duct;
    if (r.t16) patch.t16 = true;
    if (r.place) patch.place = r.place;
    if (r.sys !== undefined) {
      patch.sys = r.sys;
      var sysDef = st.master.systems.filter(function (x) { return x.name === r.sys; })[0];
      var fam = sysDef ? sysDef.fam : 'pipe';
      var alt = areas().filter(function (a) { return UA.famOfArea(a.name) === fam; })[0];
      if (alt && famOfSel() !== fam) { patch.areaRow = alt.row; patch.newArea = false; }
    }
    var k = km[item];
    if (r.sizes.length && (item === 'valve' || item === 'flange')) {
      // バルブ・フランジは数量を使わない（種類とサイズだけ）
      var vsub = r.sub !== undefined ? r.sub : (item === s.item ? s.sub : defaultSub(item)), vent = item === s.item ? Object.assign({}, s.vent) : {};
      r.sizes.forEach(function (z) { if ((st.master.sizes.A || []).some(function (x) { return x.n === z.n; })) vent[ventKey(vsub, z.n)] = {}; });
      patch.vent = vent;
    } else if (r.sizes.length && k && k.series) {
      var ent = {};
      var subNow = patch.sub !== undefined ? patch.sub : (item === s.item ? s.sub : defaultSub(item));
      var ser = item === 'pipe' && subNow === 'SUS' && st.master.sizes.Su ? 'Su' : k.series;
      r.sizes.forEach(function (z) { if ((st.master.sizes[ser] || []).some(function (x) { return x.n === z.n; })) ent[String(z.n)] = { g: z.g }; });
      patch.ent = ent;
    }
    if (r.rq && k && !k.series) patch.rq = r.rq;
    if (patch.item || patch.place || patch.sys !== undefined || patch.sub !== undefined) { patch.specAuto = true; patch.custom = ''; }
    normalizeSel(patch);
    if (r.t) {
      Object.keys(s.ent).forEach(function (kk) { s.ent[kk] = Object.assign({}, s.ent[kk], { f: r.t }); });
      if (k && !k.series) s.rf = r.t;
    }
  }

  /* ---------- 画面 ---------- */
  var ICON = {
    pipe: 'M2 10h5 M2 14h5 M10 10h4 M10 14h4 M17 10h5 M17 14h5 M7 8h3v8H7z M14 8h3v8h-3z',
    valve: 'M3 8v10l9-5z M21 8v10l-9-5z M12 13V6 M8.5 5h7',
    flange: 'M2 10h5 M2 14h5 M17 10h5 M17 14h5 M7 5h3v14H7z M14 5h3v14h-3z',
    rect: 'M3 9h13v10H3z M3 9l4-4h13l-4 4 M16 19l4-4V5',
    round: 'M6 6h12 M6 18h12 M6 6a3 6 0 0 0 0 12 M6 6a3 6 0 0 1 0 12 M18 6a3 6 0 0 1 0 12',
    box: 'M3 7h18v10H3z M7 7v10 M17 7v10 M3 12h4 M17 12h4'
  };
  function svg(d, cls) { return '<svg class="' + (cls || 'ic') + '" viewBox="0 0 24 24" aria-hidden="true"><path d="' + d + '"></path></svg>'; }
  function chip(act, val, label, on, extra) {
    return '<button type="button" class="chip' + (on ? ' on' : '') + '" data-act="' + act + '" data-v="' + esc(val) + '"' + (extra || '') + ' aria-pressed="' + (on ? 'true' : 'false') + '">' + esc(label) + '</button>';
  }
  function stepHead(no, title, sub, done) {
    return '<div class="sh"><span class="no' + (done ? ' done' : '') + '">' + no + '</span><b>' + esc(title) + '</b><small>' + esc(sub) + '</small></div>';
  }

  function render() {
    var app = $('#app'); if (!app) return;
    var ae = document.activeElement, focus = null;
    if (ae && ae.dataset && ae.dataset.f) focus = { f: ae.dataset.f, k: ae.dataset.k || '', s: ae.selectionStart, e: ae.selectionEnd };
    var main = $('#main'), scroll = main ? main.scrollTop : 0;
    app.innerHTML = view();
    var m2 = $('#main'); if (m2) m2.scrollTop = scroll;
    if (st.ui.noRestore) { focus = null; st.ui.noRestore = false; }
    if (focus) {
      var el = document.querySelector('[data-f="' + focus.f + '"]' + (focus.k ? '[data-k="' + CSS.escape(focus.k) + '"]' : ''));
      if (el) { el.focus(); try { if (focus.s != null) el.setSelectionRange(focus.s, focus.e); } catch (e) { /* noop */ } }
    }
  }

  function view() {
    var h = [];
    h.push('<header class="hd"><div class="brand"><span class="logo">' + svg('M12 9a3 3 0 1 0 0 6a3 3 0 1 0 0-6z M12 4a8 8 0 1 0 0 16a8 8 0 1 0 0-16z M12 4v2 M12 18v2 M4 12h2 M18 12h2', 'ic dark') + '</span>' +
      '<div><b>内訳アシスト</b><small>v' + VERSION + '</small></div></div>' +
      '<button type="button" class="iconbtn" data-act="menu" aria-label="メニュー" aria-expanded="' + st.ui.menu + '">' + svg('M5 12h.01 M12 12h.01 M19 12h.01') + '</button></header>');
    if (st.ui.menu) h.push(menuView());
    h.push('<div class="bar">' + (st.master ? '設定マスタ：<b>' + esc(st.masterInfo ? st.masterInfo.name : '') + '</b>' + (st.masterInfo ? '（' + esc(st.masterInfo.at) + '）' : '') : '<span class="warn">設定マスタが未読込</span>') +
      '<button type="button" class="link" data-act="loadMaster">' + (st.master ? '読み直す' : '読み込む') + '</button></div>');
    if (!st.master) { h.push(onboard()); h.push(toastView()); return h.join(''); }
    if (priceOrder().indexOf('sk') >= 0) h.push('<div class="bar">' + (st.sk ? '単価表：<b>' + esc(st.skInfo ? st.skInfo.name.replace(/\.xlsx?m?$/, '') : '') + '</b>（' + st.sk.n + '表）' : '<span class="warn">保温積算資料（単価表）が未読込</span>') +
      '<button type="button" class="link" data-act="loadSk">' + (st.sk ? '読み直す' : '読み込む') + '</button></div>');
    h.push(targetView());
    h.push('<div id="main" class="main">');
    if (st.model) {
      h.push(cmdView());
      h.push(step1()); h.push(step2()); h.push(step3()); h.push(step4()); h.push(step5());
      h.push(planView());
    }
    h.push('</div>');
    if (st.model) h.push(footView());
    h.push(toastView());
    return h.join('');
  }
  function menuView() {
    return '<div class="menu">' +
      '<button type="button" data-act="loadMaster">設定マスタを読み込む</button>' +
      '<button type="button" data-act="loadSk">保温積算資料（単価表）を読み込む</button>' +
      '<button type="button" data-act="reload">シートを読み直す</button>' +
      '<button type="button" data-act="autoOpen">' + (st.autoOpen ? '✓ ' : '') + 'このブックを開いたら自動で表示</button>' +
      '<button type="button" data-act="exportSpecs">仕様候補の変更を書き出す（' + (st.specLocal.add.length + Object.keys(st.specLocal.hide).length + Object.keys(st.specLocal.pin).length) + '件）</button>' +
      '<div class="ver">内訳アシスト v' + VERSION + '／Excel API ' + (st.api19 ? '1.9以上' : st.api17 ? '1.7以上' : '1.7未満') + '</div></div>';
  }
  function onboard() {
    return '<div class="onboard"><h2>はじめに設定マスタを読み込みます</h2>' +
      '<p>積算フォルダの <b>内訳アシスト＼内訳アシスト_設定マスタ.xlsx</b> を選んでください。内容はこのパソコンに記憶され、どの見積書でも使えます。</p>' +
      '<button type="button" class="cta" data-act="loadMaster">設定マスタを選ぶ</button>' + (st.modelErr ? '<p class="err">' + esc(st.modelErr) + '</p>' : '') + '</div>';
  }
  function targetView() {
    var opts = st.sheets.map(function (n) { return '<option' + (n === st.target ? ' selected' : '') + '>' + esc(n) + '</option>'; }).join('');
    var status = !st.sheets.length ? '<span class="ng">「内訳」で始まるシートがありません</span>'
      : st.modelErr ? '<span class="ng">' + esc(st.modelErr) + '</span>'
      : st.headerOK ? '<span class="ok">✓ 内訳シート</span>' : '<span class="ng">1行目に「保温厚」がありません</span>';
    return '<div class="target"><label for="tg">書き込み先</label><select id="tg" data-f="target">' + opts + '</select>' +
      '<button type="button" class="iconbtn sm" data-act="reload" aria-label="シートを読み直す">' + svg('M20 11a8 8 0 1 0-2.3 5.7 M20 5v6h-6') + '</button>' + status + '</div>';
  }
  function cmdView() {
    var pres = st.master.presets.map(function (p, i) {
      var on = !st.sel.newArea && curAreaName() === p.area && st.sel.sys === p.sys && st.sel.place === p.place && st.sel.item === p.item && (!p.sub || st.sel.sub === p.sub);
      return '<button type="button" class="pill' + (on ? ' on' : '') + '" data-act="preset" data-v="' + i + '">' + esc(p.label) + '</button>';
    }).join('');
    return '<div class="cmd"><label class="cmdbox">' + svg('M4 17l6-5-6-5 M12 19h8', 'ic acc') +
      '<input data-f="cmd" value="' + esc(st.sel.cmd) + '" placeholder="例）冷温水 機械室 Yスト 50A 2 65A 3" aria-label="スマート入力" autocomplete="off"></label>' +
      (pres ? '<div class="pills">' + pres + '</div>' : '') + '</div>';
  }
  function step1() {
    var s = st.sel, al = areas();
    var noA = s.noArea && !s.newArea;
    var ah = al.map(function (a) { return chip('area', a.row, a.label, !s.newArea && !noA && a.row === s.areaRow); }).join('') +
      chip('noArea', '1', '大項目なし', noA, ' title="大項目・系統は書かずに、施工箇所と仕様の行から書きます"') +
      chip('newArea', '1', '＋新しい大項目', s.newArea);
    var na = s.newArea ? '<div class="row newarea"><span class="anum" title="番号は自動で付きます">' + areaNumNext() + '.</span><input class="in" data-f="newAreaName" value="' + esc(stripNum(s.newAreaName)) + '" placeholder="例）配管設備" aria-label="新しい大項目の名前（番号は自動）"></div>' : '';
    var sh = sysList().map(function (b) { return chip('sys', b, b || '（系統なし）', b === s.sys); }).join('');
    return '<section class="st">' + stepHead('01', '区分・系統', noA ? '大項目なし' : 'A列・B列', true) +
      '<div class="lbl">大項目</div><div class="chips">' + ah + '</div>' + na +
      (noA ? '<div class="hint">大項目・系統の行は書かず、施工箇所（C列）と仕様（E列）の行から書きます。同じ施工箇所・仕様が最後の大項目の後ろにあれば、そこに追加します。</div>'
        : '<div class="lbl">系統</div><div class="chips">' + sh + '</div>') + '</section>';
  }
  function step2() {
    var s = st.sel;
    return '<section class="st">' + stepHead('02', '施工箇所', 'C列', true) + '<div class="chips">' +
      placeList().map(function (p) { return chip('place', p, p, p === s.place); }).join('') + '</div></section>';
  }
  function step3() {
    var s = st.sel, k = kind();
    var tiles = kinds().map(function (x) {
      return '<button type="button" class="tile' + (x.id === s.item ? ' on' : '') + '" data-act="item" data-v="' + esc(x.id) + '" aria-pressed="' + (x.id === s.item) + '">' +
        svg(ICON[x.id] || ICON.pipe) + '<span><b>' + esc(x.label) + '</b><small>' + esc(x.unit) + '</small></span></button>';
    }).join('');
    var subs = subsOf(s.item), sub = '';
    if (subs.length) {
      var subT = s.item === 'pipe' ? '管の種類（D列）' : (k && k.subTitle || '種別');
      if (s.item === 'valve' || s.item === 'flange') subT += '　押してからサイズを選ぶ（複数の種類をまとめて登録できます）';
      sub = '<div class="subbox"><div class="subhd"><b>' + esc(subT) + '</b>' +
        ((s.item === 'rect' || s.item === 'round') ? chip('t16', '1', '(1.6t)', s.t16) : '') + '</div><div class="chips">' +
        subs.map(function (x) {
          var n = (s.item === 'valve' || s.item === 'flange') ? Object.keys(s.vent || {}).filter(function (kk) { return kk.split('|')[0] === x.id; }).length : 0;
          return chip('sub', x.id, x.label + (n ? '（' + n + '）' : ''), x.id === s.sub, x.tip ? ' title="' + esc(x.tip) + '"' : '');
        }).join('') + '</div></div>';
    }
    var rate = (st.master.output && st.master.output.elbowRate) || 0.35;
    var elb = s.item === 'pipe' ? '<div class="elb"><button type="button" class="chip elbtn' + (st.ui.elbowOn ? ' on' : '') + '" data-act="elbow" aria-pressed="' + st.ui.elbowOn + '">' + (st.ui.elbowOn ? '✓ ' : '＋ ') + 'エルボ 1式</button>' +
      '<span class="hint">J列＝直管の金額の合計×' + rate + '（10円単位で切り上げ）。同じ塊に既にあれば、範囲だけ直します</span></div>' : '';
    return '<section class="st">' + stepHead('03', '品目', 'D列・E列', true) + '<div class="tiles">' + tiles + '</div>' + sub + elb + '</section>';
  }
  // 仕様を「保温材」と「外装・仕上げ」の2つの列に分けて、縦にそろえて見せる
  function specCols(spec) {
    var p = UA.specParts(spec);
    var c1 = p.tags.filter(function (t) { return t !== '2層' && t !== '遮音'; }).map(function (t) { return t + '＋'; }).join('') +
      p.pre.map(function (t) { return t + '＋'; }).join('') + p.facing;
    return { lead: c1, mat: p.mat, rest: p.rest.join(' ＋ ') || '（なし）' };
  }
  function specView(spec) {
    if (!spec) return '<span class="sv"><b>未選択</b></span>';
    var c = specCols(spec);
    return '<span class="sv"><span class="sv1">' + esc(c.lead) + '<b>' + esc(c.mat) + '</b></span><span class="sv2">＋ ' + esc(c.rest) + '</span></span>';
  }
  // 積算資料の表が当たるか（代表のサイズで確認）
  var skHintCache = {};
  function skHint(spec, place) {
    if (!st.sk || !spec) return null;
    var s = st.sel, k = kind(), pl = place || s.place, keyC = s.item + '|' + pl + '|' + dOf() + '|' + spec + '|' + Object.keys(s.ent).join(',');
    if (keyC in skHintCache) return skHintCache[keyC];
    var sizes = Object.keys(s.ent).map(Number);
    if (!sizes.length) sizes = k && k.series === 'A' ? (s.item === 'pipe' ? [50, 25, 100] : [100, 65]) : k && k.series === 'φ' ? [200] : [s.item === 'rect' ? '矩形' : (s.sub || 'BOX')];
    var hit = null;
    sizes.some(function (z) {
      return [UA.defThick(st.master, s.item, spec, z), '25', '40', '50', '20', '30'].some(function (f) {
        var r = UA.skLook(st.sk, st.master, { item: s.item, d: dOf(), place: pl, spec: spec, size: z, f: f });
        if (r.p != null) { hit = r; return true; } return false;
      });
    });
    skHintCache[keyC] = hit ? { id: hit.id, approx: hit.approx, tip: hit.ref.y } : { none: true };
    return skHintCache[keyC];
  }
  function skTag(spec, place) {
    var h = skHint(spec, place);
    if (!h) return '';
    if (h.none) return '<span class="sk none" title="保温積算資料に当たる表がありません（M_積算資料対応）">表なし</span>';
    return '<span class="sk' + (h.approx ? ' near' : '') + '" title="' + esc(h.tip) + '">§' + esc(h.id) + '</span>';
  }
  function step4() {
    var s = st.sel, gs = specGroups(), flat = [];
    st.ui.specIdx = flat;
    var hit = null;
    gs.forEach(function (g) { g.items.forEach(function (x) { if (!hit && x.spec === s.spec && g.key !== 'hidden') hit = x; }); });
    var badge = s.custom && s.spec === U.normSpec(s.custom) ? ['直接入力', 'b-warn'] : hit && hit.sheet ? [(s.specAuto ? 'AUTO ' : '') + 'この見積で使用中', 'b-info']
      : hit && hit.here ? [(s.specAuto ? 'AUTO ' : '') + s.place + 'の候補', 'b-info'] : hit ? ['全箇所の候補', 'b-mute'] : ['候補にない仕様', 'b-warn'];
    var nAll = 0; gs.forEach(function (g) { if (g.key !== 'hidden') nAll += g.items.length; });
    var head = '<div class="speccard"><span class="spv" title="(' + esc(s.spec) + ')">' + specView(s.spec) + '</span>' +
      '<button type="button" class="ghost" data-act="specToggle" aria-expanded="' + st.ui.specOpen + '">' + (st.ui.specOpen ? '閉じる' : '候補' + nAll + '件') + '</button></div>';
    var list = '';
    if (st.ui.specOpen) {
      var seg = ['', '民間', '官庁'].map(function (c) { return '<button type="button" class="seg' + (st.siteType === c ? ' on' : '') + '" data-act="siteType" data-v="' + c + '">' + (c || 'すべて') + '</button>'; }).join('');
      list = '<div class="speclist"><div class="slhd"><span title="この見積書を民間・官庁のどちらとして候補を出すか（この見積書に記憶）">この見積の区分</span><span class="segs">' + seg + '</span></div>';
      list += '<div class="scolh"><span>保温材</span><span>外装・仕上げ</span></div>';
      gs.forEach(function (g) {
        list += '<div class="slgt">' + esc(g.title.replace(/（ほかの施工箇所）/, '（ほかの施工箇所）')) + '</div>';
        list += g.items.map(function (it) { var i = flat.push(Object.assign({ grp: g.key }, it)) - 1; return optView(it, i, g.key); }).join('');
      });
      list += '<div class="slft"><button type="button" class="link sm" data-act="specEditMode">' + (st.ui.specEditMode ? '編集を終わる' : '候補を編集（固定・直す・外す）') + '</button>' +
        (gs.hiddenCount ? '<button type="button" class="link sm" data-act="showHidden">' + (st.ui.showHidden ? '外した候補を隠す' : '外した候補を表示（' + gs.hiddenCount + '件）') + '</button>' : '') + '</div>';
      var pl = s.place || '（全箇所）';
      list += '<div class="addspec"><div class="lbl">候補を追加（このパソコンに記憶。メニューから書き出して設定マスタへ移せます）</div>' +
        '<input class="in" data-f="specNew" value="' + esc(st.ui.newSpec) + '" placeholder="例）ALK付Gw筒＋亀甲金網16m/m（括弧なし）" aria-label="追加する仕様">' +
        '<div class="row"><select class="in" data-f="specNewPlace" aria-label="候補を出す施工箇所"><option value="here"' + (st.ui.newPlace !== 'all' ? ' selected' : '') + '>' + esc(pl) + 'だけ</option><option value="all"' + (st.ui.newPlace === 'all' ? ' selected' : '') + '>全箇所</option></select>' +
        '<select class="in" data-f="specNewCat" aria-label="民間・官庁"><option value="">民間・官庁とも</option><option value="民間"' + (st.ui.newCat === '民間' ? ' selected' : '') + '>民間だけ</option><option value="官庁"' + (st.ui.newCat === '官庁' ? ' selected' : '') + '>官庁だけ</option></select></div>' +
        '<div class="row"><button type="button" class="ghost" data-act="specUse"' + (st.ui.newSpec ? '' : ' disabled') + '>今回だけ使う</button><button type="button" class="ghost acc" data-act="specAdd"' + (st.ui.newSpec ? '' : ' disabled') + '>候補に追加して使う</button></div></div>';
      list += '</div>';
    }
    return '<section class="st">' + stepHead('04', '仕様', 'E列（括弧書き）', !!s.spec) + head + list + '</section>';
  }
  function optView(it, i, gk) {
    var s = st.sel, on = it.spec === s.spec && gk !== 'hidden';
    var ed = st.ui.specEdit;
    if (ed && ed.i === i) {
      return '<div class="sopt edit"><input class="in" data-f="specEditText" value="' + esc(ed.spec) + '" aria-label="仕様を編集">' +
        '<div class="row"><select class="in" data-f="specEditPlace"><option value="here"' + (ed.place !== '（全箇所）' ? ' selected' : '') + '>' + esc(s.place) + 'だけ</option><option value="all"' + (ed.place === '（全箇所）' ? ' selected' : '') + '>全箇所</option></select>' +
        '<select class="in" data-f="specEditCat"><option value="">民間・官庁とも</option><option value="民間"' + (ed.cat === '民間' ? ' selected' : '') + '>民間だけ</option><option value="官庁"' + (ed.cat === '官庁' ? ' selected' : '') + '>官庁だけ</option></select></div>' +
        '<div class="row"><button type="button" class="ghost" data-act="specCancel">取消</button><button type="button" class="ghost acc" data-act="specSave">保存</button></div></div>';
    }
    var c = specCols(it.spec), acts = '';
    if (st.ui.specEditMode || gk === 'hidden') {
      if (gk === 'hidden') acts = '<button type="button" class="ic-b" data-act="specUnhide" data-v="' + i + '" title="候補に戻す">戻す</button>';
      else if (gk !== 'sheet') acts = '<button type="button" class="ic-b' + (it.pin ? ' on' : '') + '" data-act="specPin" data-v="' + i + '" title="' + (it.pin ? '固定をやめる' : 'いちばん上に固定') + '">' + (it.pin ? '★' : '☆') + '</button>' +
        '<button type="button" class="ic-b" data-act="specEdit" data-v="' + i + '" title="直す">✎</button>' +
        '<button type="button" class="ic-b" data-act="specHide" data-v="' + i + '" title="' + (it.src === 'local' ? '削除' : '候補から外す（設定マスタは変わりません）') + '">×</button>';
      else acts = '<span class="ic-n" title="この見積書で使っている仕様は、ここでは直せません">—</span>';
    }
    return '<div class="srow' + (on ? ' on' : '') + (gk === 'hidden' ? ' off' : '') + '">' +
      '<button type="button" class="srow-b" data-act="specPick" data-v="' + i + '" title="(' + esc(it.spec) + ')"' + (gk === 'hidden' ? ' disabled' : '') + '>' +
      '<span class="c1">' + (it.pin ? '<i class="pin">★</i>' : '') + esc(c.lead) + '<b>' + esc(c.mat) + '</b></span><span class="c2">' + esc(c.rest) + '</span></button>' +
      (acts ? '<span class="srow-a">' + acts + '</span>' : '') + '</div>';
  }
  function saveSpecLocal() { if (!save(K_SPECS, st.specLocal)) toast('候補の変更をこのパソコンに記憶できませんでした', 'warn'); }
  function specKindOf() { return UA.KIND_OF_ITEM[st.sel.item] || st.sel.item; }
  function step5() {
    var s = st.sel, k = kind(), h = [];
    if (!k) return '';
    if (s.item === 'valve' || s.item === 'flange') return step5Valve();
    var d = dOf();
    var tl = (st.master.output && st.master.output.thickList) || ['20', '25', '30', '40', '50', '65', '75'];
    h.push('<div class="thick"><span class="lbl">保温厚</span>' + ['（標準）'].concat(tl).map(function (t) {
      var v = t === '（標準）' ? '' : t, on = (s.tSel || '') === v;
      return '<button type="button" class="tk' + (on ? ' on' : '') + '" data-act="thick" data-v="' + v + '" aria-pressed="' + on + '" title="' + (v ? '選んでいるサイズの保温厚を ' + v + 'mm にします' : 'サイズごとの標準の保温厚（M_標準厚）') + '">' + esc(v ? v : '標準') + '</button>';
    }).join('') + '</div>');
    if (k.series) {
      var sizes = st.master.sizes[seriesOf()] || [];
      if (seriesOf() === 'Su') h.push('<div class="hint">ステンレス管（D列に SUS）。単価は外径が同じ鋼管のサイズで引きます（例 20Su→15A。M_サイズ で変更可）</div>');
      h.push('<div class="sizes">' + sizes.map(function (z) {
        var key = String(z.n), on = !!s.ent[key], g = on ? U.num(s.ent[key].g) : 0;
        return '<button type="button" class="size' + (on ? ' on' : '') + '" data-act="size" data-v="' + key + '" aria-pressed="' + on + '"><b>' + esc(z.n) + '</b>' + (g > 0 ? '<small>' + U.fmtQ(g) + '</small>' : '') + '</button>';
      }).join('') + '</div>');
      var sel = sizes.filter(function (z) { return s.ent[String(z.n)]; });
      if (sel.length) {
        h.push('<div class="qtab"><div class="qh"><span>サイズ</span><span>保温厚</span><span>数量(' + esc(unitOf()) + ')</span><span>単価</span><span>金額</span><span></span></div>');
        sel.forEach(function (z) {
          var key = String(z.n), en = s.ent[key];
          var fShown = en.f !== undefined ? en.f : UA.defThick(st.master, s.item, s.spec, z.n);
          var lp = priceOf(d, s.spec, z.n, U.thickKey(fShown));
          var manual = en.i !== undefined, price = manual ? (en.i === '' ? null : U.num(en.i)) : lp.p, g = U.num(en.g);
          var cap = manual ? ['手入力', 'c-info'] : lp.cap;
          h.push('<div class="qr"><span class="sz">' + esc(sizeText(z.n)) + '</span>' +
            '<input class="in num c" tabindex="-1" data-f="ent.f" data-k="' + key + '" value="' + esc(fShown) + '" aria-label="' + key + 'の保温厚">' +
            '<input class="in num" data-f="ent.g" data-k="' + key + '" value="' + esc(en.g) + '" placeholder="数量" aria-label="' + key + 'の数量">' +
            '<span class="pc"><input class="in num' + (price == null ? ' bad' : '') + '" tabindex="-1" data-f="ent.i" data-k="' + key + '" value="' + esc(manual ? en.i : (lp.p != null ? lp.p : '')) + '" placeholder="未登録" aria-label="' + key + 'の単価"><small class="' + cap[1] + '" title="' + esc(manual ? '' : lp.tip || '') + '">' + esc(cap[0]) + '</small></span>' +
            '<span class="amt">' + (price == null || !(g > 0) ? '—' : '¥' + U.yen(g * price)) + '</span>' +
            '<button type="button" class="x" tabindex="-1" data-act="size" data-v="' + key + '" aria-label="' + key + 'を外す">×</button></div>');
        });
        h.push('</div>');
      } else h.push('<div class="empty">サイズを押して追加（複数まとめて選べます）</div>');
    } else {
      var eVal = s.item === 'rect' ? '矩形' : (s.sub || 'BOX');
      var rf = s.rf !== undefined ? s.rf : UA.defThick(st.master, s.item, s.spec, eVal);
      var lp2 = priceOf(d, s.spec, eVal, U.thickKey(rf));
      var man = s.ri !== undefined, pr = man ? (s.ri === '' ? null : U.num(s.ri)) : lp2.p;
      var g2 = s.rq !== '' ? U.num(s.rq) : dimsTotal();
      var cap2 = man ? ['手入力', 'c-info'] : lp2.cap;
      h.push('<div class="qtab"><div class="qh q4"><span>E列</span><span>保温厚</span><span>数量(㎡)</span><span>単価</span><span>金額</span></div>' +
        '<div class="qr q4"><span class="sz">' + esc(eVal) + '</span><input class="in num c" tabindex="-1" data-f="rf" value="' + esc(rf) + '" aria-label="保温厚">' +
        '<input class="in num" data-f="rq" value="' + esc(s.rq) + '" placeholder="' + (dimsTotal() > 0 ? '寸法から ' + dimsTotal() : '数量') + '" aria-label="数量（平方メートル）">' +
        '<span class="pc"><input class="in num' + (pr == null ? ' bad' : '') + '" tabindex="-1" data-f="ri" value="' + esc(man ? s.ri : (lp2.p != null ? lp2.p : '')) + '" placeholder="未登録" aria-label="単価"><small class="' + cap2[1] + '" title="' + esc(man ? '' : lp2.tip || '') + '">' + esc(cap2[0]) + '</small></span>' +
        '<span class="amt">' + (pr == null || !(g2 > 0) ? '—' : '¥' + U.yen(g2 * pr)) + '</span></div></div>');
      h.push('<div class="dims"><div class="lbl">寸法から計算（数量が空欄のとき使用）　2×(W＋H)×長さ</div>' + s.dims.map(function (dd, i) {
        var w = U.num(dd.w), hh = U.num(dd.h), l = U.num(dd.l), a = (w > 0 && hh > 0 && l > 0) ? 2 * (w + hh) / 1000 * l : 0;
        return '<div class="dr"><input class="in num" data-f="dim.w" data-k="' + i + '" value="' + esc(dd.w) + '" placeholder="W mm" aria-label="幅"><span>×</span>' +
          '<input class="in num" data-f="dim.h" data-k="' + i + '" value="' + esc(dd.h) + '" placeholder="H mm" aria-label="高さ"><span>×</span>' +
          '<input class="in num" data-f="dim.l" data-k="' + i + '" value="' + esc(dd.l) + '" placeholder="L m" aria-label="長さ">' +
          '<span class="amt">' + (a > 0 ? U.fmtQ(Math.round(a * 10) / 10) + '㎡' : '—') + '</span><button type="button" class="x" data-act="rmDim" data-v="' + i + '" aria-label="削除">×</button></div>';
      }).join('') + '<button type="button" class="ghost sm" data-act="addDim">＋寸法を追加</button></div>');
    }
    var ok = buildLines().length > 0;
    return '<section class="st">' + stepHead('05', k.series ? '保温厚・サイズ・数量' : '保温厚・数量', 'E・F・G・I列', ok) + h.join('') + (k.series && sel && sel.length ? '<div class="hint">数量の欄で Enter＝次のサイズへ、Ctrl＋Enter＝前のサイズへ</div>' : '') + '</section>';
  }
  // バルブ・フランジ：種類とサイズを選ぶと、保温厚（配管から）・単位・単価（協会単価）が入る。数量は入れない
  function step5Valve() {
    var s = st.sel, h = [], sizes = st.master.sizes.A || [];
    var cur = subsOf(s.item).filter(function (x) { return x.id === s.sub; })[0];
    h.push('<div class="lbl">' + esc(cur ? cur.label : '種類') + ' のサイズ</div><div class="sizes">' + sizes.map(function (z) {
      var on = !!s.vent[ventKey(s.sub, z.n)];
      return '<button type="button" class="size' + (on ? ' on' : '') + '" data-act="size" data-v="' + z.n + '" aria-pressed="' + on + '"><b>' + esc(z.n) + '</b></button>';
    }).join('') + '</div>');
    var list = ventList();
    if (list.length) {
      h.push('<div class="qtab vt"><div class="qh"><span>種類</span><span>サイズ</span><span>保温厚</span><span>単位</span><span>単価</span><span></span></div>');
      list.forEach(function (v) {
        var r = ventRow(v), cap = r.manual ? ['手入力', 'c-info'] : r.pr.cap;
        var tTip = '配管 ' + r.vt.pipe + 'mm（' + r.vt.from + '）→ ' + r.vt.f + 'mm';
        h.push('<div class="qr"><span class="vd">' + esc(v.sub.label) + '</span><span class="sz">' + esc(v.z.n + 'A') + '</span>' +
          '<input class="in num c" data-f="vent.f" data-k="' + esc(v.key) + '" value="' + esc(r.f) + '" title="' + esc(tTip) + '" aria-label="保温厚">' +
          '<span class="un">' + esc(unitOf()) + '</span>' +
          '<span class="pc"><input class="in num' + (r.i == null ? ' bad' : '') + '" data-f="vent.i" data-k="' + esc(v.key) + '" value="' + esc(r.manual ? v.en.i : (r.pr.p != null ? r.pr.p : '')) + '" placeholder="協会単価なし" aria-label="単価"><small class="' + cap[1] + '" title="' + esc(r.manual ? '' : r.pr.tip || '') + '">' + esc(cap[0]) + '</small></span>' +
          '<button type="button" class="x" tabindex="-1" data-act="vdel" data-v="' + esc(v.key) + '" aria-label="外す">×</button></div>');
      });
      h.push('</div>');
      var lt = (st.master.output && st.master.output.valveThick) || [25, 50];
      h.push('<div class="hint">数量（G列）は空欄で登録します。内訳に直接入力してください。保温厚は同じ施工箇所・仕様の配管（直管）の保温厚から（' + lt.map(function (x, i) { return (i ? '' : '') + x + 'mm以下→' + x; }).join('、') + '）。</div>');
    } else h.push('<div class="empty">種類を押してからサイズを押して追加（数量は入れません）</div>');
    return '<section class="st">' + stepHead('05', '種類・サイズ', 'D・E・F・H・I列', list.length > 0) + h.join('') + '</section>';
  }
  function planTitle(p) {
    var s = st.sel, where = s.sys || curAreaName();
    if (p.mode === 'row') {
      var c = p.ctx || {}, cx = [c.area, c.sys, c.place ? '(' + c.place + ')' : ''].filter(Boolean).join('／');
      return p.atRow + '行目に入れる' + (cx ? '（' + cx + ' の中）' : '') + (p.headAdded ? '・仕様の行も入れます' : '');
    }
    if (p.mode === 'sub') return p.how === 'merge' ? '同じ区分・仕様の塊（同じサイズは数量を足す）' : '同じ区分・仕様の明細の下に追加';
    if (p.mode === 'noarea') return '大項目なしで、施工箇所・仕様の行から末尾に作成';
    return p.mode === 'newsub' ? where + ' に新しい仕様の塊を作成' : p.mode === 'newsys' ? curAreaName() + ' に系統「' + s.sys + '」を作成' : '大項目「' + curAreaName() + '」を新しく作成';
  }
  // 登録先の入れ方（操作者が選ぶ）
  function howView() {
    var how = st.ui.how;
    var h = '<div class="howrow"><span class="lbl">入れ方</span><div class="segs">' + ['add', 'merge', 'row'].map(function (k) {
      return '<button type="button" class="seg' + (how === k ? ' on' : '') + '" data-act="how" data-v="' + k + '">' + HOW_LABEL[k] + '</button>';
    }).join('') + '</div></div>';
    if (how === 'row') {
      h += '<div class="atrow"><input class="in num" data-f="atRow" inputmode="numeric" value="' + esc(st.ui.atRow) + '" aria-label="入れる行の番号"><span>行目に入れる</span>' +
        '<button type="button" class="ghost sm" data-act="atSel" data-v="top" title="Excel で選んでいる行の位置に入れます（その行から下は下にずれます）">選択行に</button>' +
        '<button type="button" class="ghost sm" data-act="atSel" data-v="below" title="Excel で選んでいる行のすぐ下に入れます">選択行の下に</button></div>';
    }
    var note = how === 'add' ? '今ある行は変えずに、同じ区分・仕様の明細の下へ新しい行で追加します。'
      : how === 'merge' ? '同じ種別・サイズ・厚みの行があれば、その行の数量に足します（今ある行の数量が変わります）。'
      : '指定した行に新しい行を入れます（その行から下は下にずれます）。大項目・系統はその位置のものになり、仕様が違うときは仕様の行も入れます。';
    return h + '<div class="hint">' + note + '</div>';
  }
  function planView() {
    var p = currentPlan();
    if (!p) return '';
    if (!p.ok) return '<section class="plan' + (st.ui.how === 'row' ? ' man' : '') + '"><div class="pt"><b>登録先</b></div>' + howView() + '<div class="pt mute">' + esc(p.message) + '</div></section>';
    var t = planTitle(p);
    var rows = [];
    p.W.forEach(function (x) {
      if (!(x.isNew || x.addG || x.addJ || x.elbowSet)) return;
      var desc, jc = colsOf().j;
      if (x.t === 'area') desc = x.a; else if (x.t === 'sys') desc = x.b;
      else if (x.t === 'spec') desc = (x.c ? '(' + x.c + ') ' : '') + '(' + x.e + ')';
      else if (x.t === 'line') desc = (x.isNew ? (x.dShow ? x.dShow + ' ' : '') : (x.d ? x.d + ' ' : '')) + x.e + '　' + (x.f ? 't' + x.f + '　' : '') +
        (x.g === '' ? '数量は手入力　' + x.h : (x.addG ? U.fmtQ(x.g) + '＋' + U.fmtQ(x.addG) : U.fmtQ(x.g)) + x.h) + '　' + (x.i == null ? '協会単価なし' : '@' + U.yen(x.i));
      else if (x.t === 'elbow') desc = (x.e || 'エルボ') + ' 1式　' + (x.runFrom ? jc + x.runFrom + '〜' + jc + x.runTo + '×' + x.rateUsed + '（切上げ10円）' + (x.jEst != null ? ' ≒¥' + U.yen(x.jEst) : '') : '');
      else if (x.t === 'total') desc = x.label;
      else desc = '';
      var tag = x.elbowSet && !x.isNew ? '<span class="tag">式の範囲を直す</span>' : !x.isNew ? '<span class="tag">数量を足す</span>' : x.restore ? '<span class="tag">元の仕様に戻す</span>' : '';
      rows.push('<div class="pr' + (x.isNew ? '' : ' mg') + '"><span class="rn">' + x.finalRow + '</span><span class="pd">' + esc(desc) + '</span>' + tag + '</div>');
    });
    var notes = [];
    if (p.elbowNote) notes.push(p.elbowNote);
    if (p.elbowSkipped) notes.push('エルボは、上に直管が無いので入れません');
    return '<section class="plan' + (st.ui.how === 'row' ? ' man' : '') + '"><div class="pt"><b>登録先</b>　' + esc(t) + '（' + (p.firstRow === p.lastRow ? p.firstRow + '行目' : p.firstRow + '〜' + p.lastRow + '行目') + '）</div>' +
      howView() + '<div class="prs">' + rows.join('') + '</div>' + (notes.length ? '<div class="hint warn">' + esc(notes.join('。')) + '</div>' : '') + '</section>';
  }
  function footView() {
    var p = currentPlan(), ok = p && p.ok && st.headerOK;
    return '<footer class="ft"><div class="sum"><span>追加 <b>' + (p && p.ok ? p.addCount : 0) + '</b>行' + (p && p.ok && p.mergeCount ? '・加算 <b>' + p.mergeCount + '</b>行' : '') + (p && p.ok && p.fixCount ? '・式の修正 <b>' + p.fixCount + '</b>' : '') + '</span>' +
      '<strong>＋¥' + U.yen(p && p.ok ? p.addAmount : 0) + '</strong></div>' +
      '<div class="btns"><button type="button" class="ghost" data-act="locate"' + (ok ? '' : ' disabled') + ' title="登録する前に、書き込む位置をシート上で選択して見せます（シートは変わりません）">位置を確認</button>' +
      '<button type="button" class="cta" data-act="register"' + (ok && !st.ui.busy ? '' : ' disabled') + '>' + (st.ui.busy ? '処理中…' : '内訳に登録') + '</button></div></footer>';
  }
  function toastView() {
    var t = st.ui.toast;
    if (!t) return '';
    return '<div class="toast ' + t.kind + '"><span>' + esc(t.text) + '</span>' + (t.undo && st.undo ? '<button type="button" data-act="undo">元に戻す</button>' : '') +
      '<button type="button" class="x" data-act="closeToast" aria-label="閉じる">×</button></div>';
  }
  var toastTimer = null;
  function toast(text, kind, withUndo) {
    st.ui.toast = { text: text, kind: kind || 'ok', undo: !!withUndo };
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { st.ui.toast = null; render(); }, withUndo ? 60000 : (kind === 'err' ? 15000 : 6000));
  }

  /* ---------- 操作 ---------- */
  function bindEvents() {
    document.addEventListener('click', function (ev) {
      var el = ev.target.closest('[data-act]');
      if (!el || el.disabled) return;
      var a = el.dataset.act, v = el.dataset.v, s = st.sel;
      if (a !== 'menu' && st.ui.menu) st.ui.menu = false;
      switch (a) {
        case 'menu': st.ui.menu = !st.ui.menu; break;
        case 'loadMaster': st.ui.menu = false; $('#masterFile').click(); return;
        case 'loadSk': st.ui.menu = false; $('#skFile').click(); return;
        case 'reload': refreshSheets().then(render); return;
        case 'autoOpen': setAutoOpen(!st.autoOpen); return;
        case 'exportMemo': exportMemo(); return;
        case 'area': normalizeSel({ areaRow: +v, newArea: false, noArea: false, specAuto: true, custom: '' }); break;
        case 'newArea': normalizeSel({ newArea: true, noArea: false, newAreaName: s.newAreaName || nextAreaName(kind() ? kind().fam : 'pipe'), specAuto: true, custom: '' }); break;
        case 'noArea': normalizeSel({ noArea: true, newArea: false, sys: '', specAuto: true, custom: '' }); break;
        case 'elbow': st.ui.elbowOn = !st.ui.elbowOn; saveUi(); break;
        case 'vdel': { var vv = Object.assign({}, s.vent); delete vv[v]; s.vent = vv; st.ui.noRestore = true; break; }
        case 'sys': normalizeSel({ sys: v, specAuto: true, custom: '' }); break;
        case 'place': normalizeSel({ place: v, specAuto: true, custom: '' }); break;
        case 'item': normalizeSel({ item: v, specAuto: true, custom: '' }); break;
        case 'sub': normalizeSel({ sub: v }); break;
        case 't16': normalizeSel({ t16: !s.t16 }); break;
        case 'specToggle': st.ui.specOpen = !st.ui.specOpen; st.ui.specEdit = null; break;
        case 'specPick': {
          var o = st.ui.specIdx[+v];
          if (o) { normalizeSel({ spec: o.spec, specAuto: false, custom: '' }); st.ui.specOpen = false; }
          break;
        }
        case 'specPin': {
          var op = st.ui.specIdx[+v];
          if (op && op.key) { if (st.specLocal.pin[op.key]) delete st.specLocal.pin[op.key]; else st.specLocal.pin[op.key] = true; saveSpecLocal(); }
          break;
        }
        case 'specHide': {
          var oh = st.ui.specIdx[+v];
          if (oh && oh.src === 'local') st.specLocal.add = st.specLocal.add.filter(function (a) { return a.id !== oh.id; });
          else if (oh && oh.key) st.specLocal.hide[oh.key] = true;
          saveSpecLocal(); if (oh && oh.spec === s.spec) normalizeSel({ specAuto: true });
          break;
        }
        case 'specUnhide': { var ou = st.ui.specIdx[+v]; if (ou && ou.key) { delete st.specLocal.hide[ou.key]; saveSpecLocal(); } break; }
        case 'showHidden': st.ui.showHidden = !st.ui.showHidden; break;
        case 'specEditMode': st.ui.specEditMode = !st.ui.specEditMode; st.ui.specEdit = null; break;
        case 'specEdit': { var oe = st.ui.specIdx[+v]; if (oe) st.ui.specEdit = { i: +v, spec: oe.spec, place: oe.place, cat: oe.cat || '', src: oe.src, id: oe.id, key: oe.key, kind: oe.kind }; break; }
        case 'specCancel': st.ui.specEdit = null; break;
        case 'specSave': {
          var ed = st.ui.specEdit, sp = ed ? U.normSpec(ed.spec) : '';
          if (ed && sp) {
            var plc = ed.place === '（全箇所）' ? '（全箇所）' : s.place;
            if (ed.src === 'local') st.specLocal.add = st.specLocal.add.map(function (a) { return a.id === ed.id ? Object.assign({}, a, { spec: sp, place: plc, cat: ed.cat }) : a; });
            else { st.specLocal.hide[ed.key] = true; st.specLocal.add.push({ id: 'L' + Date.now(), kind: ed.kind || specKindOf(), place: plc, spec: sp, cat: ed.cat, from: ed.spec }); }
            saveSpecLocal(); st.ui.specEdit = null; normalizeSel({ spec: sp, specAuto: false, custom: '' });
          }
          break;
        }
        case 'specUse': { if (st.ui.newSpec) { normalizeSel({ spec: U.normSpec(st.ui.newSpec), specAuto: false, custom: st.ui.newSpec }); st.ui.specOpen = false; } break; }
        case 'specAdd': {
          var ns = U.normSpec(st.ui.newSpec);
          if (ns) {
            st.specLocal.add.push({ id: 'L' + Date.now(), kind: specKindOf(), place: st.ui.newPlace === 'all' ? '（全箇所）' : s.place, spec: ns, cat: st.ui.newCat });
            saveSpecLocal(); st.ui.newSpec = ''; normalizeSel({ spec: ns, specAuto: false, custom: '' }); st.ui.specOpen = false;
            toast('仕様の候補に追加しました（このパソコンに記憶）', 'ok');
          }
          break;
        }
        case 'siteType': {
          st.siteType = v || '';
          try { Office.context.document.settings.set('uchiwake.siteType', st.siteType); Office.context.document.settings.saveAsync(function () {}); } catch (e) { /* noop */ }
          normalizeSel({ specAuto: s.specAuto });
          break;
        }
        case 'thick': {
          s.tSel = v || '';
          var et = {}; Object.keys(s.ent).forEach(function (kk) { et[kk] = Object.assign({}, s.ent[kk]); if (v) et[kk].f = v; else delete et[kk].f; delete et[kk].i; }); s.ent = et;
          s.rf = v || undefined; s.ri = undefined;
          st.ui.noRestore = true;
          break;
        }
        case 'exportSpecs': exportSpecs(); return;
        case 'size': {
          if (s.item === 'valve' || s.item === 'flange') {
            var ve = Object.assign({}, s.vent), vk = ventKey(s.sub, v);
            if (ve[vk]) delete ve[vk]; else ve[vk] = {};
            s.vent = ve; st.ui.noRestore = true; break;
          }
          var e = Object.assign({}, s.ent);
          if (e[v]) delete e[v]; else e[v] = { g: '' };
          if (e[v] && s.tSel) e[v].f = s.tSel;
          s.ent = e;
          st.ui.noRestore = true;   // サイズのボタンを押しても画面は動かさない（数量の欄は下に増えるだけ）
          break;
        }
        case 'addDim': s.dims = s.dims.concat([{ w: '', h: '', l: '' }]); break;
        case 'rmDim': s.dims = s.dims.filter(function (_, i) { return i !== +v; }); if (!s.dims.length) s.dims = [{ w: '', h: '', l: '' }]; break;
        case 'preset': {
          var p = st.master.presets[+v];
          if (p) {
            var ar = areas().filter(function (x) { return U.trimAll(x.name) === U.trimAll(p.area); })[0];
            var patch = { item: p.item, sub: p.sub, place: p.place, sys: p.sys, specAuto: true, custom: '' };
            if (ar) { patch.areaRow = ar.row; patch.newArea = false; patch.noArea = false; } else if (p.area) { patch.newArea = true; patch.noArea = false; patch.newAreaName = stripNum(p.area); } else { patch.noArea = true; patch.newArea = false; }
            normalizeSel(patch);
            if (p.sys !== undefined) st.sel.sys = sysList().indexOf(p.sys) >= 0 ? p.sys : st.sel.sys;
            normalizeSel();
          }
          break;
        }
        case 'how': {
          if (v === 'row' && !st.ui.atRow) {
            // はじめて「行を指定」にしたときは、自動で入れる位置を初期値にする
            var auto = st.ui.how; st.ui.how = 'add';
            var pa = currentPlan(); st.ui.how = auto;
            if (pa && pa.ok && pa.news.length) st.ui.atRow = String(Math.min.apply(null, pa.news.map(function (x) { return x.finalRow; })));
          }
          st.ui.how = v;
          if (v !== 'row') st.ui.howSaved = v;
          saveUi();
          break;
        }
        case 'atSel': pickRow(v === 'below'); return;
        case 'register': register(); return;
        case 'locate': locate(); return;
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
      var el = ev.target;
      if (el.dataset && el.dataset.f === 'target') { st.target = el.value; st.sel.areaRow = null; loadModel().then(function () { normalizeSel(); render(); }); }
      var ff = el.dataset && el.dataset.f;
      if (ff === 'specNewPlace') st.ui.newPlace = el.value;
      if (ff === 'specNewCat') st.ui.newCat = el.value;
      if (ff === 'specEditPlace' && st.ui.specEdit) st.ui.specEdit.place = el.value === 'all' ? '（全箇所）' : st.sel.place;
      if (ff === 'specEditCat' && st.ui.specEdit) st.ui.specEdit.cat = el.value;
    });
    // Enter は登録しない。数量の欄では Enter＝次のサイズの数量、Ctrl＋Enter＝前のサイズの数量
    document.addEventListener('keydown', function (ev) {
      if (ev.key !== 'Enter' || ev.isComposing || ev.keyCode === 229 || st.ui.composing) return;
      var el = ev.target;
      if (!el || (el.tagName !== 'INPUT' && el.tagName !== 'SELECT')) return;
      var f = el.dataset && el.dataset.f;
      ev.preventDefault();
      if (f === 'ent.g' || f === 'ent.f' || f === 'ent.i') {
        var list = Array.prototype.slice.call(document.querySelectorAll('[data-f="ent.g"]'));
        var i = list.map(function (x) { return x.dataset.k; }).indexOf(el.dataset.k);
        var nx = list[ev.ctrlKey ? i - 1 : (f === 'ent.g' ? i + 1 : i)];
        if (nx) { nx.focus(); try { nx.select(); } catch (e) { /* noop */ } }
        return;
      }
      if (f === 'specNew' && st.ui.newSpec) { el.blur(); return; }
      if (f === 'specEditText') return;
      if (el.blur) el.blur();
    });  }
  function onInput(ev, silent) {
    var el = ev.target;
    if (!el || !el.dataset || !el.dataset.f) return;
    var f = el.dataset.f, k = el.dataset.k, v = el.value, s = st.sel;
    if (f === 'target') return;
    if (f === 'cmd') { s.cmd = v; if (!silent) onCmd(v); }
    else if (f === 'newAreaName') { s.newAreaName = stripNum(v); if (!silent) normalizeSel({ specAuto: true }); }
    else if (f === 'atRow') st.ui.atRow = v;
    else if (f === 'custom') { s.custom = v; if (!silent) normalizeSel(v ? { spec: U.normSpec(v), specAuto: false } : { specAuto: true }); }
    else if (f === 'specNew') { st.ui.newSpec = v; }
    else if (f === 'specEditText') { if (st.ui.specEdit) st.ui.specEdit.spec = v; return; }
    else if (f === 'specNewPlace' || f === 'specNewCat' || f === 'specEditPlace' || f === 'specEditCat') return;
    else if (f === 'rq') s.rq = v;
    else if (f === 'rf') s.rf = v;
    else if (f === 'ri') s.ri = v;
    else if (f.indexOf('vent.') === 0) { var vf = f.slice(5), ve2 = Object.assign({}, s.vent); ve2[k] = Object.assign({}, ve2[k]); ve2[k][vf] = v; s.vent = ve2; }
    else if (f.indexOf('ent.') === 0) { var fld = f.slice(4); var e = Object.assign({}, s.ent); e[k] = Object.assign({}, e[k]); e[k][fld] = v; s.ent = e; }
    else if (f.indexOf('dim.') === 0) { var dk = f.slice(4); s.dims = s.dims.map(function (d, i) { return i === +k ? Object.assign({}, d, (function () { var o = {}; o[dk] = v; return o; })()) : d; }); }
    if (!silent) render();
  }
})();
