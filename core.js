/*
 * 内訳アシスト  core.js
 * Excel（Office.js）に依存しない中核ロジック。
 *  - 設定マスタの読み込みと索引づくり
 *  - 「内訳」シートの解析（大項目・系統・施工箇所/仕様・明細・計・境界）
 *  - 登録計画（どこに何行入れるか、数量加算、D列の重複整理、計行のSUM修正）
 * ブラウザでは window.UA、node では module.exports で使えます。
 */
(function (root) {
  'use strict';
  var UA = {};

  /* ========== 共通関数 ========== */
  function toHalf(t) {
    return String(t == null ? '' : t).replace(/[Ａ-Ｚａ-ｚ０-９．＝：＊＋－]/g, function (c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    });
  }
  function trimAll(s) { return String(s == null ? '' : s).replace(/[\s　]+/g, ''); }
  function isEmpty(v) { return v === null || v === undefined || (typeof v === 'string' && trimAll(v) === ''); }
  function num(v) {
    if (typeof v === 'number') return v;
    if (isEmpty(v)) return 0;
    var n = parseFloat(toHalf(v).replace(/[,，\s]/g, ''));
    return isFinite(n) ? n : 0;
  }
  function isNumLike(v) { return typeof v === 'number' || (typeof v === 'string' && /^\s*[0-9０-９]+(\.[0-9]+)?\s*$/.test(v)); }
  function stripParen(s) {
    var t = String(s == null ? '' : s).trim().replace(/^[（(]/, '').replace(/[）)]$/, '');
    return t.trim();
  }
  function isParen(s) { return typeof s === 'string' && /^[\s　]*[（(]/.test(s); }
  function normSpec(s) { return trimAll(stripParen(s)).replace(/\+/g, '＋').replace(/ｍ\/ｍ/g, 'm/m'); }
  function yen(v) { return Math.round(v).toLocaleString('ja-JP'); }
  function fmtQ(q) { return (Math.round(q * 100) / 100).toLocaleString('ja-JP', { maximumFractionDigits: 2 }); }
  function colNum(letter) {
    var s = String(letter).toUpperCase().replace(/[^A-Z]/g, ''), n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  }
  function colLetter(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function dateStr(v) {
    if (typeof v === 'number' && v > 20000 && v < 80000) {           // Excel のシリアル値
      var d = new Date(Math.round((v - 25569) * 86400000));
      return d.getUTCFullYear() + '/' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '/' + ('0' + d.getUTCDate()).slice(-2);
    }
    var s = toHalf(v).trim().replace(/[-.年月]/g, '/').replace(/日/, '');
    var m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    return m ? m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2) : s;
  }
  function today() { var d = new Date(); return d.getFullYear() + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + ('0' + d.getDate()).slice(-2); }
  function sizeKey(v) { if (typeof v === 'number') return String(v); var t = trimAll(toHalf(v)); return /^\d+(\.\d+)?$/.test(t) ? String(parseFloat(t)) : t; }
  function thickKey(v) { if (typeof v === 'number') return String(v); return trimAll(toHalf(v)); }
  UA.util = { toHalf: toHalf, trimAll: trimAll, isEmpty: isEmpty, num: num, stripParen: stripParen, normSpec: normSpec, yen: yen, fmtQ: fmtQ, colNum: colNum, colLetter: colLetter, dateStr: dateStr, today: today, sizeKey: sizeKey, thickKey: thickKey };

  var PIECE_UNITS = ['ヶ', 'ケ', '個', 'ヶ所', 'ケ所', '箇所', '台', '基'];
  UA.PIECE_UNITS = PIECE_UNITS;
  var KIND_OF_ITEM = { pipe: 'pipe', valve: 'valve', flange: 'valve', rect: 'rect', round: 'round', box: 'box' };
  UA.KIND_OF_ITEM = KIND_OF_ITEM;
  var DEFAULT_THICK = { pipe: '20', valve: '25', rect: '25', round: '25', box: '25' };

  /* ========== 設定マスタ ========== */
  // sheets: { シート名: [ {見出し: 値, ...}, ... ] }
  function pick(row, prefix) {
    if (!row) return '';
    var keys = Object.keys(row), i, k;
    for (i = 0; i < keys.length; i++) if (trimAll(keys[i]) === prefix) return row[keys[i]];
    for (i = 0; i < keys.length; i++) { k = trimAll(keys[i]); if (k.indexOf(prefix + '（') === 0 || k.indexOf(prefix + '(') === 0) return row[keys[i]]; }
    for (i = 0; i < keys.length; i++) if (trimAll(keys[i]).indexOf(prefix) === 0) return row[keys[i]];
    return '';
  }
  function useOK(v) { var t = trimAll(v); return t === '' || t === '○' || t === '〇' || t === 'o' || t === 'O'; }
  function str(v) { return v == null ? '' : String(v).trim(); }
  function bySort(a, b) { return (a.order || 9999) - (b.order || 9999); }

  UA.parseMaster = function (sheets) {
    var errors = [];
    function rows(name) {
      if (!sheets[name]) { errors.push('シート「' + name + '」が見つかりません'); return []; }
      return sheets[name];
    }
    var m = { kinds: [], subs: {}, places: [], systems: [], sizes: { A: [], 'φ': [] }, specs: [], prices: [], thick: [], output: null, alias: [], presets: [], normalize: [] };

    rows('M_品目').forEach(function (r) {
      var id = str(pick(r, '品目ID'));
      if (!id) return;
      var s = str(pick(r, 'サイズ系列'));
      m.kinds.push({
        id: id, label: str(pick(r, '表示名')) || id, fam: str(pick(r, '区分')).indexOf('ダクト') >= 0 ? 'duct' : 'pipe',
        unit: str(pick(r, '単位')), series: s === 'A' ? 'A' : (s === 'φ' || s === 'Φ' ? 'φ' : ''),
        eText: str(pick(r, 'E列の書き方')), subTitle: str(pick(r, '種別欄の見出し')), use: useOK(pick(r, '使用する')), order: num(pick(r, '表示順'))
      });
    });
    m.kinds.sort(bySort);

    rows('M_種別').forEach(function (r) {
      var k = str(pick(r, '品目ID')), d = str(pick(r, '種別'));
      if (!k || !d) return;
      (m.subs[k] = m.subs[k] || []).push({ id: d === '（なし）' ? '' : d, label: d, tip: str(pick(r, '説明')), use: useOK(pick(r, '使用する')), order: num(pick(r, '表示順')) });
    });
    Object.keys(m.subs).forEach(function (k) { m.subs[k].sort(bySort); });

    rows('M_施工箇所').forEach(function (r) {
      var p = trimAll(stripParen(pick(r, '施工箇所')));
      if (!p) return;
      m.places.push({ name: p, aliases: str(pick(r, 'スマート入力の別名')).split(/[、,，\s]+/).filter(Boolean), use: useOK(pick(r, '使用する')), order: num(pick(r, '表示順')) });
    });
    m.places.sort(bySort);

    rows('M_系統').forEach(function (r) {
      var n = str(pick(r, '系統'));
      if (!n) return;
      m.systems.push({ name: n === '（系統なし）' ? '' : n, label: n, fam: str(pick(r, '区分')).indexOf('ダクト') >= 0 ? 'duct' : 'pipe',
        aliases: str(pick(r, 'スマート入力の別名')).split(/[、,，\s]+/).filter(Boolean), use: useOK(pick(r, '使用する')), order: num(pick(r, '表示順')) });
    });
    m.systems.sort(bySort);

    rows('M_サイズ').forEach(function (r) {
      var s = str(pick(r, '系列')); s = (s === 'Φ') ? 'φ' : s;
      var n = pick(r, 'サイズ');
      if (!s || isEmpty(n) || !useOK(pick(r, '使用する'))) return;
      (m.sizes[s] = m.sizes[s] || []).push({ n: num(n), label: str(pick(r, 'ボタンの表示')) || (s === 'A' ? num(n) + 'A' : 'φ' + num(n)), order: num(pick(r, '表示順')) });
    });
    Object.keys(m.sizes).forEach(function (k) { m.sizes[k].sort(function (a, b) { return a.n - b.n; }); });

    rows('M_仕様候補').forEach(function (r) {
      var s = normSpec(pick(r, '仕様'));
      if (!s) return;
      var pl = trimAll(stripParen(pick(r, '施工箇所')));
      m.specs.push({ kind: str(pick(r, '仕様区分')) || '共通', place: pl || '（全箇所）', spec: s, prio: num(pick(r, '優先順位')) || 0, use: useOK(pick(r, '使用する')), src: str(pick(r, '由来')) });
    });

    rows('M_単価履歴').forEach(function (r) {
      var price = pick(r, '単価');
      var spec = normSpec(pick(r, '仕様'));
      if (isEmpty(price) || !spec) return;
      m.prices.push({
        item: str(pick(r, '品目ID')), kind: str(pick(r, '仕様区分')) || KIND_OF_ITEM[str(pick(r, '品目ID'))] || '',
        d: str(pick(r, '種別')), spec: spec, size: sizeKey(pick(r, 'サイズ')), f: thickKey(pick(r, '保温厚')),
        unit: str(pick(r, '単位')), price: num(price), date: dateStr(pick(r, '見積日')), sys: str(pick(r, '系統')), place: trimAll(pick(r, '施工箇所'))
      });
    });

    rows('M_標準厚').forEach(function (r) {
      var s = normSpec(pick(r, '仕様'));
      var f = thickKey(pick(r, '標準の保温厚'));
      if (!s || !f) return;
      m.thick.push({ kind: str(pick(r, '仕様区分')), spec: s, size: sizeKey(pick(r, 'サイズ')), f: f });
    });

    var cols = { area: 'A', sys: 'B', place: 'C', d: 'D', spec: 'E', size: 'E', f: 'F', g: 'G', h: 'H', i: 'I', j: 'J', label: 'I' };
    var out = { cols: cols, startRow: 5, amountTpl: '=IF(AND(G{行}<>"",I{行}=""),0,IF(AND(G{行}="",I{行}=""),"",IF(AND(G{行}<>"",I{行}<>""),G{行}*I{行},"")))' };
    var MAPKEY = { '大項目': 'area', '系統': 'sys', '施工箇所': 'place', '種別': 'd', '仕様': 'spec', 'サイズ': 'size', '保温厚': 'f', '数量': 'g', '単位': 'h', '単価': 'i', '金額': 'j', '系統計（ラベル）': 'label' };
    (sheets['M_出力設定'] || []).forEach(function (r) {
      var k = str(pick(r, '項目')), c = trimAll(pick(r, '列')), v = str(pick(r, '書き込む内容'));
      if (k === '明細の開始行' && num(v) > 0) out.startRow = num(v);
      if (k === '金額' && v.charAt(0) === '=') out.amountTpl = v;
      if (MAPKEY[k] && /^[A-Za-z]{1,2}$/.test(c)) cols[MAPKEY[k]] = c.toUpperCase();
    });
    m.output = out;

    (sheets['M_別名'] || []).forEach(function (r) {
      var w = str(pick(r, '入力語')), t = str(pick(r, '種類')), v = str(pick(r, '値'));
      if (w && t) m.alias.push({ word: w, type: t, value: v, memo: str(pick(r, '備考')) });
    });
    (sheets['M_よく使う'] || []).forEach(function (r) {
      var l = str(pick(r, '表示名'));
      if (!l || !useOK(pick(r, '使用する'))) return;
      var sys = str(pick(r, '系統'));
      m.presets.push({ label: l, area: str(pick(r, '大項目')), sys: sys === '（系統なし）' ? '' : sys, place: trimAll(pick(r, '施工箇所')),
        item: str(pick(r, '品目ID')), sub: str(pick(r, '種別')) === '（なし）' ? '' : str(pick(r, '種別')), order: num(pick(r, '表示順')) });
    });
    m.presets.sort(bySort);
    (sheets['M_表記統一'] || []).forEach(function (r) {
      var a = str(pick(r, '元の表記')), b = str(pick(r, '統一後の表記')), c = str(pick(r, '対象の列')), on = trimAll(pick(r, '統一する'));
      if (a && b) m.normalize.push({ from: a, to: b, col: c.charAt(0).toUpperCase(), on: on === '○' || on === '〇' });
    });

    if (!m.kinds.length) errors.push('M_品目 に品目がありません');
    m.errors = errors;
    return UA.indexMaster(m);
  };

  var VALVE_PRI = ['GV', 'BV', 'CV', 'BAV', 'FLG', 'フランジ', 'Yスト', 'フレキ', 'FJ', 'SUSFJ'];
  UA.indexMaster = function (m) {
    var price = {}, specCount = {}, thick = {};
    m.prices.forEach(function (p) {
      var dk = p.kind === 'valve' ? p.d : '';
      var key = p.kind + '|' + dk + '|' + p.spec + '|' + p.size + '|' + p.f;
      var cur = price[key];
      if (!cur || p.date >= cur.date) price[key] = { p: p.price, date: p.date, d: p.d, ref: null };
      if (p.kind === 'valve') {
        var k2 = 'valve|*|' + p.spec + '|' + p.size + '|' + p.f;
        var c2 = price[k2];
        var better = !c2 || p.date > c2.date || (p.date === c2.date && rank(p.d) < rank(c2.d));
        if (better) price[k2] = { p: p.price, date: p.date, d: p.d, ref: p.d };
      }
      var ck = p.kind + '|' + p.place + '|' + p.spec;
      specCount[ck] = (specCount[ck] || 0) + 1;
      var ak = p.kind + '|*|' + p.spec;
      specCount[ak] = (specCount[ak] || 0) + 1;
    });
    function rank(d) { var i = VALVE_PRI.indexOf(d); return i < 0 ? 99 : i; }
    m.thick.forEach(function (t) { thick[t.kind + '|' + t.spec + '|' + t.size] = t.f; });
    m.priceIdx = price; m.specCount = specCount; m.thickIdx = thick;
    m.kindById = {}; m.kinds.forEach(function (k) { m.kindById[k.id] = k; });
    return m;
  };

  // 単価を探す（local: 手入力の単価メモ { key: {p, date} }）
  UA.lookPrice = function (m, item, d, spec, size, f, local) {
    var kind = KIND_OF_ITEM[item] || item;
    var dk = kind === 'valve' ? d : '';
    var key = kind + '|' + dk + '|' + spec + '|' + sizeKey(size) + '|' + thickKey(f);
    var hit = m.priceIdx[key] || null;
    var memo = local && local[key];
    if (memo && (!hit || memo.date >= hit.date)) return { p: memo.p, date: memo.date, ref: null, memo: true, key: key };
    if (hit) return { p: hit.p, date: hit.date, ref: null, key: key };
    if (kind === 'valve') {
      var fb = m.priceIdx['valve|*|' + spec + '|' + sizeKey(size) + '|' + thickKey(f)];
      if (fb) return { p: fb.p, date: fb.date, ref: fb.ref, key: key };
    }
    return { p: null, key: key };
  };
  UA.priceKey = function (item, d, spec, size, f) {
    var kind = KIND_OF_ITEM[item] || item;
    return kind + '|' + (kind === 'valve' ? d : '') + '|' + spec + '|' + sizeKey(size) + '|' + thickKey(f);
  };

  UA.defThick = function (m, item, spec, size) {
    var kind = KIND_OF_ITEM[item] || item;
    return m.thickIdx[kind + '|' + spec + '|' + sizeKey(size)] || m.thickIdx[kind + '|' + spec + '|（全サイズ）'] || DEFAULT_THICK[kind] || '25';
  };

  // 仕様の候補（シートで使用中 → 優先順位 → 過去の使用回数）
  UA.specOptions = function (m, item, place, sheetSpec) {
    var kind = KIND_OF_ITEM[item] || item, out = [], seen = {};
    if (sheetSpec) { out.push({ spec: sheetSpec, sheet: true, n: m.specCount[kind + '|' + place + '|' + sheetSpec] || 0 }); seen[sheetSpec] = 1; }
    var list = m.specs.filter(function (s) {
      return s.use && (s.kind === kind || s.kind === '共通') && (s.place === place || s.place === '（全箇所）');
    }).map(function (s) {
      var here = s.place === place;
      return { spec: s.spec, prio: s.prio, here: here, n: here ? (m.specCount[kind + '|' + place + '|' + s.spec] || 0) : (m.specCount[kind + '|*|' + s.spec] || 0) };
    });
    list.sort(function (a, b) {
      var pa = a.prio > 0 ? a.prio : 9999, pb = b.prio > 0 ? b.prio : 9999;
      if (pa !== pb) return pa - pb;
      if (a.here !== b.here) return a.here ? -1 : 1;
      return b.n - a.n;
    });
    list.forEach(function (s) { if (!seen[s.spec]) { seen[s.spec] = 1; out.push(s); } });
    return out;
  };

  // 表記統一
  UA.normalizeOut = function (m, col, v) {
    if (typeof v !== 'string' || !m || !m.normalize) return v;
    for (var i = 0; i < m.normalize.length; i++) {
      var n = m.normalize[i];
      if (n.on && n.col === col && n.from === v) return n.to;
    }
    return v;
  };

  /* ========== 内訳シートの解析 ========== */
  // values/formulas: 1行目からの2次元配列（A列〜）。cols: 列の割当。
  UA.parseSheet = function (values, formulas, cols, startRow) {
    cols = cols || { area: 'A', sys: 'B', place: 'C', d: 'D', spec: 'E', size: 'E', f: 'F', g: 'G', h: 'H', i: 'I', j: 'J', label: 'I' };
    startRow = startRow || 5;
    var ci = {}; Object.keys(cols).forEach(function (k) { ci[k] = colNum(cols[k]) - 1; });
    var recs = [], boundary = null, lastContent = startRow - 1, lastD = null;
    function cell(r, k) { var row = values[r - 1]; return row ? row[ci[k]] : ''; }
    function fcell(r, k) { var row = formulas && formulas[r - 1]; return row ? row[ci[k]] : ''; }
    for (var r = startRow; r <= values.length; r++) {
      var A = cell(r, 'area'), B = cell(r, 'sys'), C = cell(r, 'place'), D = cell(r, 'd'), E = cell(r, 'spec'),
        F = cell(r, 'f'), G = cell(r, 'g'), H = cell(r, 'h'), I = cell(r, 'label'), J = cell(r, 'j');
      var rec = { row: r };
      var allEmpty = [A, B, C, D, E, F, G, H, I].every(isEmpty);
      if (allEmpty) { rec.t = 'blank'; recs.push(rec); lastD = null; continue; }
      lastContent = r;
      var aT = trimAll(A), labelT = typeof I === 'string' ? trimAll(I) : '';
      if (!isEmpty(A) && typeof A === 'string') {
        if (aT === '経費' || aT === '見積条件' || aT.charAt(0) === '一' || aT.charAt(0) === '※') { boundary = { row: r, text: String(A).trim() }; break; }
        rec.t = 'area'; rec.a = String(A).trim(); recs.push(rec); lastD = null; continue;
      }
      if (labelT && (/計$/.test(labelT) || labelT === '合計') && isEmpty(G) && isEmpty(H)) {
        rec.t = 'total'; rec.label = String(I); rec.name = labelT.replace(/合?計$/, ''); rec.jf = fcell(r, 'j');
        if (labelT.indexOf('工事費') === 0 || labelT === '合計' || labelT === '総合計') { rec.scope = 'grand'; boundary = { row: r, text: String(I).trim(), grand: true, jf: rec.jf }; break; }
        recs.push(rec); lastD = null; continue;
      }
      if (!isEmpty(B)) {
        rec.t = 'sys'; rec.b = String(B).trim();
        if (isParen(E)) { rec.e = normSpec(E); rec.c = isParen(C) ? trimAll(stripParen(C)) : ''; }
        recs.push(rec); lastD = null; continue;
      }
      if (isParen(C) || isParen(E)) {
        rec.t = 'spec'; rec.c = isParen(C) ? trimAll(stripParen(C)) : ''; rec.d = isEmpty(D) ? '' : String(D).trim(); rec.e = isParen(E) ? normSpec(E) : '';
        recs.push(rec); lastD = null; continue;
      }
      var hT = trimAll(H);
      if (typeof E === 'string' && /エルボ|チーズ|継手/.test(E) && hT === '式') {
        rec.t = 'elbow'; rec.e = String(E).trim(); rec.g = num(G); rec.h = hT; rec.j = num(J); recs.push(rec); continue;
      }
      if (!isEmpty(G) && isNumLike(G)) {
        rec.t = 'line';
        rec.dRaw = isEmpty(D) ? '' : String(D).trim();
        if (rec.dRaw) { rec.d = rec.dRaw; lastD = rec.dRaw; }
        else rec.d = (PIECE_UNITS.indexOf(hT) >= 0 && lastD) ? lastD : '';
        rec.e = typeof E === 'number' ? E : (isEmpty(E) ? '' : String(E).trim());
        rec.f = isEmpty(F) ? '' : thickKey(F); rec.g = num(G); rec.h = hT; rec.i = isEmpty(cell(r, 'i')) ? null : num(cell(r, 'i'));
        recs.push(rec); continue;
      }
      rec.t = 'other'; rec.text = [D, E, C].filter(function (x) { return !isEmpty(x); }).map(String).join(' ');
      recs.push(rec); lastD = null;
    }
    // 計の範囲（系統計 or 大項目計）
    var areaNames = {}, curArea = null, sysNames = {};
    recs.forEach(function (x) {
      if (x.t === 'area') { curArea = x; areaNames[trimAll(x.a)] = true; }
      if (x.t === 'sys') { x.area = curArea ? curArea.a : ''; sysNames[trimAll(x.b)] = true; }
      if (x.t === 'total') {
        x.area = curArea ? curArea.a : '';
        if (curArea && trimAll(x.name) === trimAll(curArea.a)) x.scope = 'area';
        else if (sysNames[trimAll(x.name)]) x.scope = 'sys';
        else if (areaNames[trimAll(x.name)]) x.scope = 'area';
        else x.scope = 'sys';
      }
    });
    // J列の式（R1C1）のくせ
    return { recs: recs, boundary: boundary, lastContent: lastContent, startRow: startRow, cols: cols };
  };

  // 大項目の一覧（シート上）
  // 大項目の一覧（同じ名前が複数あるときは直前の見出しを添える）
  UA.areasOf = function (model) {
    var list = model.recs.filter(function (r) { return r.t === 'area'; });
    var cnt = {}; list.forEach(function (r) { var k = trimAll(r.a); cnt[k] = (cnt[k] || 0) + 1; });
    var out = [], lastUnique = '';
    list.forEach(function (r) {
      var k = trimAll(r.a);
      if (cnt[k] === 1) lastUnique = r.a;
      out.push({ row: r.row, name: r.a, label: cnt[k] > 1 && lastUnique ? r.a + '（' + lastUnique + '）' : r.a });
    });
    return out;
  };
  UA.sysOf = function (model, areaRow) {
    var rs = model.recs, out = [], inA = false;
    for (var i = 0; i < rs.length; i++) {
      if (rs[i].t === 'area') inA = rs[i].row === areaRow;
      else if (inA && rs[i].t === 'sys' && out.indexOf(rs[i].b) < 0) out.push(rs[i].b);
    }
    return out;
  };

  // 品目の種類を明細行から推定
  function lineItem(r, fam, m) {
    if (r.t !== 'line') return null;
    if (r.e === '矩形') return 'rect';
    if (r.e === 'BOX' || r.e === 'チャンバー') return 'box';
    var subs = (m && m.subs) || {};
    function inSubs(k) { return (subs[k] || []).some(function (s) { return s.id && s.id === r.d; }); }
    if (inSubs('flange') || r.d === 'FLG' || r.d === 'フランジ') return 'flange';
    if (inSubs('valve') || (r.d && PIECE_UNITS.indexOf(r.h) >= 0)) return 'valve';
    if (typeof r.e === 'number') return fam === 'duct' ? 'round' : 'pipe';
    return null;
  }
  UA.lineItem = lineItem;
  function famOfArea(a) { return /ダクト/.test(a || '') ? 'duct' : 'pipe'; }
  UA.famOfArea = famOfArea;

  function findArea(recs, area, areaRow) {
    var i;
    if (areaRow) { for (i = 0; i < recs.length; i++) if (recs[i].t === 'area' && !recs[i].isNew && recs[i].row === areaRow) return i; }
    for (i = 0; i < recs.length; i++) if (recs[i].t === 'area' && trimAll(recs[i].a) === trimAll(area)) return i;
    return -1;
  }
  function areaEnd(recs, ai) { for (var i = ai + 1; i < recs.length; i++) if (recs[i].t === 'area' || recs[i].t === 'end') return i; return recs.length; }
  function isAreaTotal(r) { return r.t === 'total' && r.scope === 'area'; }
  function sysRange(recs, ai, ae, sys) {
    var i, j;
    if (!sys) {
      for (i = ai + 1; i < ae; i++) if (recs[i].t === 'sys' || isAreaTotal(recs[i])) return { s: ai, e: i };
      return { s: ai, e: ae };
    }
    for (i = ai + 1; i < ae; i++) {
      if (recs[i].t === 'sys' && trimAll(recs[i].b) === trimAll(sys)) {
        for (j = i + 1; j < ae; j++) if (recs[j].t === 'sys' || isAreaTotal(recs[j])) return { s: i, e: j };
        return { s: i, e: ae };
      }
    }
    return null;
  }
  function specOfRec(r) { return r.t === 'spec' ? r : (r.t === 'sys' && r.e ? r : null); }
  function findSub(recs, s, e, place, spec) {
    var cur = '';
    for (var i = s; i < e; i++) {
      var r = specOfRec(recs[i]);
      if (!r) continue;
      if (r.c) cur = r.c;
      if (i === s && recs[i].t === 'area') continue;
      if (trimAll(cur) === trimAll(place) && r.e === spec) return recs[i];
    }
    return null;
  }
  function placeBefore(recs, s, idx) { var cur = ''; for (var i = s; i < idx; i++) { var r = specOfRec(recs[i]); if (r && r.c) cur = r.c; } return cur; }

  // この見積書の同じ系統・施工箇所で、同じ品目に使っている仕様
  UA.specInSheet = function (model, m, area, areaRow, sys, place, item, d) {
    var recs = model.recs, ai = findArea(recs, area, areaRow);
    if (ai < 0) return null;
    var sr = sysRange(recs, ai, areaEnd(recs, ai), sys);
    if (!sr) return null;
    var fam = famOfArea(area), cur = '', spec = null, found = null, foundSame = null;
    for (var i = (sys ? sr.s : sr.s + 1); i < sr.e; i++) {
      var r = recs[i], sp = specOfRec(r);
      if (sp) { if (sp.c) cur = sp.c; spec = trimAll(cur) === trimAll(place) ? sp.e : null; continue; }
      if (spec && lineItem(r, fam, m) === item) {
        found = spec;
        if (d !== undefined && d !== '' && r.d === d && !foundSame) foundSame = spec;
      }
    }
    return foundSame || found;
  };

  /* ========== 登録計画 ==========
     sel = { area, newArea(bool), sys('' = 系統なし), place, spec, item, lines:[{d,e,f,g,h,i,src}], elbow:金額 }
     返り値 = { ok, mode, inserts:[{at, recs:[...]}], cellOps:[{row,col,value,prev}], totalOps:[...], final:[...] } */
  UA.plan = function (model, sel, m) {
    var res = { ok: false, mode: '', message: '' };
    if (!sel.lines.length && !(sel.elbow > 0)) { res.message = '数量を入れると、ここに登録先が表示されます'; return res; }
    if (!sel.spec) { res.message = '仕様を選んでください'; return res; }
    if (!sel.area) { res.message = '大項目を選ぶか、新しい大項目を入力してください'; return res; }
    var base = model.recs.map(function (r) { return Object.assign({}, r); });
    // 末尾の目印（この行の上に差し込む）
    var endRow;
    if (model.boundary) {
      endRow = model.boundary.row;
    } else {
      endRow = model.lastContent >= model.startRow ? model.lastContent + 1 : model.startRow;
    }
    // 境界の直前の空行はまとめて「末尾」とみなす
    while (base.length && base[base.length - 1].t === 'blank') { endRow = Math.min(endRow, base[base.length - 1].row); base.pop(); }
    base.push({ t: 'end', row: endRow });

    var W = base;
    var lines = sel.lines.map(function (l) { return Object.assign({ t: 'line', isNew: true }, l); });
    if (sel.item === 'pipe' && sel.elbow > 0) lines.push({ t: 'elbow', isNew: true, e: 'エルボ・チーズ', g: 1, h: '式', j: sel.elbow });
    var fam = famOfArea(sel.area);
    var ai = sel.newArea ? -1 : findArea(W, sel.area, sel.areaRow);
    var i, p;
    function splice(at, arr) { Array.prototype.splice.apply(W, [at, 0].concat(arr)); }
    function nb(t, o) { return Object.assign({ t: t, isNew: true }, o || {}); }

    if (ai < 0) {
      // 新しい大項目を末尾に作る
      p = W.length - 1; // end の前
      var hasPrev = W.some(function (r) { return r.t !== 'end' && r.t !== 'blank'; });
      var block = [];
      if (hasPrev) block.push(nb('blank'));
      block.push(nb('area', { a: sel.area }));
      var sysTotal = null;
      if (sel.sys) block.push(nb('sys', { b: sel.sys }));
      block.push(nb('spec', { c: sel.place, d: '', e: sel.spec }));
      block = block.concat(lines);
      block.push(nb('blank'));
      if (sel.sys) { sysTotal = nb('total', { scope: 'sys', label: sel.sys + '　計', name: sel.sys }); block.push(sysTotal); block.push(nb('blank')); }
      block.push(nb('total', { scope: 'area', label: sel.area + '　計', name: sel.area, sysTotalRef: sysTotal }));
      splice(p, block);
      res.mode = 'newarea';
    } else {
      var ae = areaEnd(W, ai);
      var sr = sysRange(W, ai, ae, sel.sys);
      if (sr) {
        var sub = findSub(W, sr.s + (sel.sys ? 0 : 1), sr.e, sel.place, sel.spec);
        if (sub) {
          insertIntoSub(W, sub, lines);
          res.mode = 'sub';
        } else {
          p = sr.e;
          for (i = sr.s + 1; i < sr.e; i++) if (W[i].t === 'total' && W[i].scope === 'sys') { p = i; break; }
          while (p - 1 > sr.s && W[p - 1].t === 'blank') p--;
          var cur = placeBefore(W, sr.s, p);
          var blk = [nb('spec', { c: trimAll(cur) === trimAll(sel.place) ? '' : sel.place, d: '', e: sel.spec })].concat(lines);
          if (p > sr.s + 1) blk.unshift(nb('blank'));
          splice(p, blk);
          res.mode = 'newsub';
        }
      } else {
        // 新しい系統
        var useSysTotal = false;
        for (i = ai + 1; i < ae; i++) if (W[i].t === 'total' && W[i].scope === 'sys') useSysTotal = true;
        p = ae;
        for (i = ai + 1; i < ae; i++) if (isAreaTotal(W[i])) { p = i; break; }
        while (p - 1 > ai && W[p - 1].t === 'blank') p--;
        var sb = [nb('sys', { b: sel.sys }), nb('spec', { c: sel.place, d: '', e: sel.spec })].concat(lines);
        if (p > ai + 1) sb.unshift(nb('blank'));
        if (useSysTotal) { sb.push(nb('blank')); sb.push(nb('total', { scope: 'sys', label: sel.sys + '　計', name: sel.sys, newSys: true })); }
        splice(p, sb);
        res.mode = 'newsys';
      }
    }

    // D列：同じ種別が続くときは最初の行だけ
    var cellOps = [];
    var prevLine = null;
    for (i = 0; i < W.length; i++) {
      var r = W[i];
      if (r.t !== 'line') { prevLine = null; continue; }
      var sameAsPrev = prevLine && prevLine.d === r.d;
      if (r.isNew) r.dShow = sameAsPrev ? '' : (r.d || '');
      else if (sameAsPrev && r.dRaw && prevLine.isNew) cellOps.push({ row: r.row, col: 'd', value: '', prev: r.dRaw, why: 'D列の重複を整理' });
      prevLine = r;
    }
    // 数量・金額の加算
    W.forEach(function (r) {
      if (!r.isNew && r.addG) cellOps.push({ row: r.row, col: 'g', value: r.g + r.addG, prev: r.g, why: '数量加算' });
      if (!r.isNew && r.addJ) cellOps.push({ row: r.row, col: 'j', value: r.j + r.addJ, prev: r.j, why: '金額加算' });
    });

    // 差し込みグループ
    var inserts = [], grp = null;
    for (i = 0; i < W.length; i++) {
      if (W[i].isNew) { if (!grp) { grp = { recs: [] }; inserts.push(grp); } grp.recs.push(W[i]); }
      else { if (grp) { grp.at = W[i].row; grp = null; } }
    }
    if (grp) grp.at = W[W.length - 1].row;
    // 最終行番号
    function shiftAt(row) { var s = 0; inserts.forEach(function (g) { if (g.at <= row) s += g.recs.length; }); return s; }
    inserts.forEach(function (g) {
      var before = 0; inserts.forEach(function (h) { if (h.at < g.at) before += h.recs.length; });
      g.recs.forEach(function (x, k) { x.finalRow = g.at + before + k; });
    });
    W.forEach(function (x) { if (!x.isNew) x.finalRow = x.row + shiftAt(x.row); });

    // 計の式の手当て
    var totalOps = [];
    var jc = model.cols.j;
    var reRange = new RegExp('^=SUM\\(\\s*\\$?' + jc + '\\$?(\\d+)\\s*:\\s*\\$?' + jc + '\\$?(\\d+)\\s*\\)$', 'i');
    W.forEach(function (x, idx) {
      if (x.t !== 'total') return;
      if (!x.isNew) {
        var mm = typeof x.jf === 'string' ? x.jf.replace(/\s/g, '').match(reRange) : null;
        if (mm) {
          var from = +mm[1], to = +mm[2];
          var inside = inserts.some(function (g) { return g.at > from && g.at <= x.row; });
          if (from < x.row && to >= x.row - 2 && inside) {
            var nf = from + shiftAt(from), nt = x.finalRow - 1;
            totalOps.push({ type: 'set', finalRow: x.finalRow, origRow: x.row, formula: '=SUM(' + jc + nf + ':' + jc + nt + ')', prev: x.jf });
          }
        }
      } else {
        // 新しく作る計
        var start = idx - 1;
        if (x.scope === 'sys') { while (start > 0 && W[start].t !== 'sys') start--; }
        else { while (start > 0 && W[start].t !== 'area') start--; }
        if (x.scope === 'area' && x.sysTotalRef) x.formula = '=SUM(' + jc + x.sysTotalRef.finalRow + ')';
        else x.formula = '=SUM(' + jc + (W[start].finalRow + 1) + ':' + jc + (x.finalRow - 1) + ')';
      }
    });
    // 新しい系統計 → 大項目計（=SUM(J..,J..) の形）に参照を足す
    W.forEach(function (x, idx) {
      if (x.t !== 'total' || !x.isNew || !x.newSys) return;
      for (var k = idx + 1; k < W.length; k++) {
        if (W[k].t === 'area' || W[k].t === 'end') break;
        if (isAreaTotal(W[k]) && !W[k].isNew) {
          var f = String(W[k].jf || '');
          if (/^=SUM\([^:]*\)$/i.test(f.replace(/\s/g, '')) && f.indexOf(',') >= 0 || /^=SUM\(\s*\$?[A-Z]+\$?\d+\s*\)$/i.test(f)) {
            totalOps.push({ type: 'append', finalRow: W[k].finalRow, origRow: W[k].row, ref: jc + x.finalRow, prev: W[k].jf });
          }
          break;
        }
      }
    });
    // 新しい大項目計 → 工事費計に参照を足す
    if (res.mode === 'newarea' && model.boundary && model.boundary.grand) {
      var at = W.filter(function (x) { return x.isNew && x.t === 'total' && x.scope === 'area'; })[0];
      var gf = String(model.boundary.jf || '');
      if (at && /^=SUM\(/i.test(gf) && gf.indexOf(':') < 0) {
        totalOps.push({ type: 'append', finalRow: model.boundary.row + shiftAt(model.boundary.row), origRow: model.boundary.row, ref: jc + at.finalRow, prev: gf });
      }
    }

    // 表示用の要約
    var news = W.filter(function (x) { return x.isNew; });
    var merges = W.filter(function (x) { return !x.isNew && (x.addG || x.addJ); });
    var rowsTouched = news.map(function (x) { return x.finalRow; }).concat(merges.map(function (x) { return x.finalRow; }));
    res.ok = true;
    res.inserts = inserts.map(function (g) { return { at: g.at, count: g.recs.length, recs: g.recs }; }).sort(function (a, b) { return b.at - a.at; });
    res.cellOps = cellOps;
    res.totalOps = totalOps;
    res.W = W;
    res.news = news; res.merges = merges;
    res.addCount = news.filter(function (x) { return x.t === 'line' || x.t === 'elbow'; }).length;
    res.mergeCount = merges.length;
    res.firstRow = rowsTouched.length ? Math.min.apply(null, rowsTouched) : null;
    res.lastRow = rowsTouched.length ? Math.max.apply(null, rowsTouched) : null;
    res.addAmount = 0;
    news.forEach(function (x) { if (x.t === 'line' && x.i != null) res.addAmount += Math.round(x.g * x.i); if (x.t === 'elbow') res.addAmount += x.j; });
    merges.forEach(function (x) { if (x.addG && x.i != null) res.addAmount += Math.round(x.addG * x.i); if (x.addJ) res.addAmount += x.addJ; });
    // 挿入の目印になる行（場所を見る用）
    res.anchorRow = res.inserts.length ? res.inserts[res.inserts.length - 1].at : (merges[0] ? merges[0].row : null);
    return res;
  };

  function insertIntoSub(W, specRec, lines) {
    lines.forEach(function (ln) {
      var st = W.indexOf(specRec) + 1, en = st, k;
      while (en < W.length && (W[en].t === 'line' || W[en].t === 'elbow')) en++;
      if (ln.t === 'elbow') {
        for (k = st; k < en; k++) if (W[k].t === 'elbow') { W[k].addJ = (W[k].addJ || 0) + ln.j; return; }
        var pe = en;
        for (k = st; k < en; k++) if (W[k].t === 'line' && W[k].d !== '') { pe = k; break; }
        W.splice(pe, 0, ln);
        return;
      }
      for (k = st; k < en; k++) {
        var r = W[k];
        if (r.t === 'line' && !r.isNew && r.d === ln.d && String(r.e) === String(ln.e) && String(r.f) === String(ln.f) && (r.h === ln.h || !r.h)) {
          r.addG = (r.addG || 0) + ln.g; return;
        }
      }
      var same = [];
      for (k = st; k < en; k++) if (W[k].t === 'line' && W[k].d === ln.d) same.push(k);
      var pos;
      if (same.length) {
        pos = same[same.length - 1] + 1;
        for (var q = 0; q < same.length; q++) {
          var ev = W[same[q]].e;
          if (typeof ev === 'number' && typeof ln.e === 'number' && ev > ln.e) { pos = same[q]; break; }
        }
      } else if (ln.d === '') {
        pos = en;
        for (k = st; k < en; k++) if (W[k].t === 'elbow' || (W[k].t === 'line' && W[k].d !== '')) { pos = k; break; }
      } else {
        pos = en;
      }
      W.splice(pos, 0, ln);
    });
  }

  /* ========== スマート入力 ========== */
  UA.buildAliases = function (m) {
    var AL = {};
    function add(w, v) { if (w) AL[toHalf(w).toUpperCase().replace(/[\s　]/g, '')] = v; }
    m.places.forEach(function (p) { add(p.name, { place: p.name }); p.aliases.forEach(function (a) { add(a, { place: p.name }); }); });
    m.systems.forEach(function (s) { if (s.name) { add(s.name, { sys: s.name }); s.aliases.forEach(function (a) { add(a, { sys: s.name }); }); } });
    m.kinds.forEach(function (k) { add(k.label, { item: k.id }); });
    Object.keys(m.subs).forEach(function (k) {
      m.subs[k].forEach(function (s) { if (s.id) { if (k === 'rect' || k === 'round') add(s.id, { duct: s.id }); else add(s.id, { item: k, sub: s.id }); } });
    });
    m.alias.forEach(function (a) {
      var t = a.type, v = a.value;
      if (t === '施工箇所') add(a.word, { place: v });
      else if (t === '系統') add(a.word, { sys: v });
      else if (t === '品目') add(a.word, { item: v });
      else if (t === '種別') {
        var k = null;
        Object.keys(m.subs).forEach(function (kk) { if (!k && (m.subs[kk] || []).some(function (s) { return s.id === v; })) k = kk; });
        if (k === 'rect' || k === 'round') add(a.word, { duct: v }); else if (k) add(a.word, { item: k, sub: v });
      } else if (t === 'ダクト名') add(a.word, { duct: v });
      else if (t === 'その他' && /1\.6/.test(v)) add(a.word, { t16: true });
    });
    return AL;
  };
  UA.parseCmd = function (text, AL) {
    var toks = toHalf(text).toUpperCase().split(/[\s　,、]+/).filter(Boolean);
    var r = { sizes: [] };
    var isN = function (x) { return /^\d+(\.\d+)?$/.test(x || ''); };
    for (var i = 0; i < toks.length; i++) {
      var k = toks[i], mm, q;
      if (AL[k]) { Object.assign(r, AL[k]); continue; }
      if ((mm = k.match(/^T=?(\d+(?:\+\d+)?)$/)) || (mm = k.match(/^(\d+)T$/))) { r.t = mm[1]; continue; }
      if ((mm = k.match(/^(\d+(?:\.\d+)?)(?:㎡|M2)$/))) { r.rq = mm[1]; continue; }
      if ((mm = k.match(/^(\d+)A(?:[:=X×*](\d+(?:\.\d+)?))?$/))) {
        q = mm[2] || ''; if (!q && isN(toks[i + 1])) { q = toks[i + 1]; i++; }
        r.sizes.push({ n: +mm[1], g: q, s: 'A' }); continue;
      }
      if ((mm = k.match(/^[ΦФφ](\d+)(?:[:=X×*](\d+(?:\.\d+)?))?$/))) {
        q = mm[2] || ''; if (!q && isN(toks[i + 1])) { q = toks[i + 1]; i++; }
        r.sizes.push({ n: +mm[1], g: q, s: 'φ' }); continue;
      }
      if ((mm = k.match(/^(\d+)[:=](\d+(?:\.\d+)?)$/))) { r.sizes.push({ n: +mm[1], g: mm[2], s: '' }); continue; }
    }
    return r;
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = UA;
  else root.UA = UA;
})(typeof window !== 'undefined' ? window : this);
