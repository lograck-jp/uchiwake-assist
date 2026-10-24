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
  // ステンレス管の呼び径（Su）→ 単価を引く鋼管の呼び径（A）。外径が同じ（20Su・25Su は一つ小さいA）
  var SU_DEFAULT = [[13, 10], [20, 15], [25, 20], [30, 25], [40, 32], [50, 40], [60, 50], [75, 65], [80, 80], [100, 100], [125, 125], [150, 150], [200, 200], [250, 250], [300, 300]];
  UA.SU_DEFAULT = SU_DEFAULT;
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
    if (!m.subs.pipe) m.subs.pipe = [{ id: '', label: '鋼管', tip: 'D列は空欄', use: true, order: 1 }, { id: 'SUS', label: 'SUS', tip: 'ステンレス管（D列に SUS）', use: true, order: 2 }];
    m.subs.pipe.forEach(function (x) { if (x.id === '' && (x.label === '（なし）' || !x.label)) x.label = '鋼管'; });
    Object.keys(m.subs).forEach(function (k) { m.subs[k].sort(bySort); });

    rows('M_施工箇所').forEach(function (r) {
      var p = trimAll(stripParen(pick(r, '施工箇所')));
      if (!p) return;
      m.places.push({ name: p, aliases: str(pick(r, 'スマート入力の別名')).split(/[、,，\s]+/).filter(Boolean), use: useOK(pick(r, '使用する')), order: num(pick(r, '表示順')), grp: trimAll(pick(r, '積算資料の場所')) });
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
      var ps = pick(r, '単価に使うサイズ');
      (m.sizes[s] = m.sizes[s] || []).push({ n: num(n), label: str(pick(r, 'ボタンの表示')) || (s === 'A' ? num(n) + 'A' : s === 'Su' ? num(n) + 'Su' : 'φ' + num(n)), order: num(pick(r, '表示順')), priceSize: isEmpty(ps) ? null : num(ps) });
    });
    // ステンレス管（Su）：マスタに無ければ既定。単価は外径が同じ鋼管（A）の行で引く
    if (!m.sizes.Su) m.sizes.Su = SU_DEFAULT.map(function (x) { return { n: x[0], label: x[0] + 'Su', order: 0, priceSize: x[1] }; });
    Object.keys(m.sizes).forEach(function (k) { m.sizes[k].sort(function (a, b) { return a.n - b.n; }); });

    rows('M_仕様候補').forEach(function (r) {
      var s = normSpec(pick(r, '仕様'));
      if (!s) return;
      var pl = trimAll(stripParen(pick(r, '施工箇所')));
      m.specs.push({ kind: str(pick(r, '仕様区分')) || '共通', place: pl || '（全箇所）', spec: s, prio: num(pick(r, '優先順位')) || 0, use: useOK(pick(r, '使用する')), src: str(pick(r, '由来')), cat: trimAll(pick(r, '区分')) });
    });

    rows('M_単価履歴').forEach(function (r) {
      var price = pick(r, '単価');
      var spec = normSpec(pick(r, '仕様'));
      if (isEmpty(price) || !spec) return;
      m.prices.push({
        item: str(pick(r, '品目ID')), kind: str(pick(r, '仕様区分')) || KIND_OF_ITEM[str(pick(r, '品目ID'))] || '',
        d: str(pick(r, '種別')), spec: spec, size: sizeKey(pick(r, 'サイズ')), f: thickKey(pick(r, '保温厚')),
        unit: str(pick(r, '単位')), price: num(price), date: dateStr(pick(r, '見積日')), sys: str(pick(r, '系統')), place: trimAll(pick(r, '施工箇所')), src: str(pick(r, '出典'))
      });
    });

    rows('M_標準厚').forEach(function (r) {
      var s = normSpec(pick(r, '仕様'));
      var f = thickKey(pick(r, '標準の保温厚'));
      if (!s || !f) return;
      m.thick.push({ kind: str(pick(r, '仕様区分')), spec: s, size: sizeKey(pick(r, 'サイズ')), f: f });
    });

    var cols = { area: 'A', sys: 'B', place: 'C', d: 'D', spec: 'E', size: 'E', f: 'F', g: 'G', h: 'H', i: 'I', j: 'J', label: 'I' };
    var refCols = { y: 'Y', z: 'Z', aa: 'AA', ab: 'AB' };
    var REFKEY = { '単価の参照元': 'y', '参照した仕様': 'z', '参照したサイズ': 'aa', '参照した厚み': 'ab' };
    var out = { cols: cols, refCols: refCols, priceSrc: '保温積算資料', skFactor: 1, makeTotals: false, elbowRate: 0.35, elbowLabel: 'エルボ', valveThick: [25, 50], thickList: ['20', '25', '30', '40', '50', '65', '75'], startRow: 5, amountTpl: '=IF(AND(G{行}<>"",I{行}=""),0,IF(AND(G{行}="",I{行}=""),"",IF(AND(G{行}<>"",I{行}<>""),G{行}*I{行},"")))' };
    var MAPKEY = { '大項目': 'area', '系統': 'sys', '施工箇所': 'place', '種別': 'd', '仕様': 'spec', 'サイズ': 'size', '保温厚': 'f', '数量': 'g', '単位': 'h', '単価': 'i', '金額': 'j', '系統計（ラベル）': 'label' };
    (sheets['M_出力設定'] || []).forEach(function (r) {
      var k = str(pick(r, '項目')), c = trimAll(pick(r, '列')), v = str(pick(r, '書き込む内容'));
      if (k === '明細の開始行' && num(v) > 0) out.startRow = num(v);
      if (k === '金額' && v.charAt(0) === '=') out.amountTpl = v;
      if (MAPKEY[k] && /^[A-Za-z]{1,2}$/.test(c)) cols[MAPKEY[k]] = c.toUpperCase();
      if (REFKEY[k]) refCols[REFKEY[k]] = /^[A-Za-z]{1,3}$/.test(c) ? c.toUpperCase() : '';
      if (k === '単価の取り方' && v) out.priceSrc = v;
      if (k === '計の行を作る') out.makeTotals = /する|○/.test(v) && !/しない/.test(v);
      if (k === '保温厚の候補' && v) out.thickList = String(v).split(/[、,，\s]+/).map(function (x) { return thickKey(x); }).filter(Boolean);
      if (k === '保温積算資料の掛率' && num(v) > 0) out.skFactor = num(v);
      if (k === 'エルボの率' && num(v) > 0) out.elbowRate = num(v);
      if (k === 'エルボの表記' && v) out.elbowLabel = v;
      if (k === 'バルブ・フランジの保温厚' && v) out.valveThick = String(v).split(/[、,，\s]+/).map(function (x) { return num(x); }).filter(function (x) { return x > 0; }).sort(function (a, b) { return a - b; });
    });
    // 保温積算資料の対応表・加算（シートがあればそれを使う。無ければ既定）
    if (sheets['M_積算資料対応']) {
      m.skRulesRaw = sheets['M_積算資料対応'].filter(function (r) { return useOK(pick(r, '使用する')); }).map(function (r) {
        return [pick(r, '品目'), pick(r, '種別'), pick(r, '場所'), pick(r, '仕様'), pick(r, '表ID'), pick(r, '列'), pick(r, '行'), pick(r, '上被'), pick(r, '掛率'), pick(r, '加算'), pick(r, 'サイズ下限'), pick(r, 'サイズ上限'), pick(r, 'メモ')];
      });
    }
    if (sheets['M_単価加算']) {
      m.skAdjRaw = sheets['M_単価加算'].filter(function (r) { return useOK(pick(r, '使用する')); }).map(function (r) {
        return [pick(r, '対象'), pick(r, '含む語'), pick(r, '加算'), pick(r, '掛率'), pick(r, 'メモ')];
      });
    }
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

    // v2：単価の計算ルール（シートがあればそれを既定にする。パネルで直したものはパソコンに記憶）
    if (sheets['M_単価計算']) {
      m.calcRulesRaw = sheets['M_単価計算'].map(function (r, i) {
        var base = str(pick(r, '基準'));
        return {
          id: 'M' + (i + 1), on: useOK(pick(r, '使用する')), name: str(pick(r, '名前')), items: listOf(pick(r, '品目ID')),
          words: str(pick(r, 'D列に含む語')), min: str(pick(r, 'サイズ以上')), max: str(pick(r, 'サイズ以下')),
          base: !base || base === '直管' ? 'run' : 'kind', baseKind: !base || base === '直管' ? '' : base,
          factor: num(pick(r, '係数')) || 1, round: roundKey(str(pick(r, '端数'))), memo: str(pick(r, 'メモ'))
        };
      }).filter(function (x) { return x.name || x.words; });
    }
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
      if (!cur || p.date >= cur.date) price[key] = { p: p.price, date: p.date, d: p.d, ref: null, src: p.src, spec: p.spec, size: p.size, f: p.f };
      if (p.kind === 'valve') {
        var k2 = 'valve|*|' + p.spec + '|' + p.size + '|' + p.f;
        var c2 = price[k2];
        var better = !c2 || p.date > c2.date || (p.date === c2.date && rank(p.d) < rank(c2.d));
        if (better) price[k2] = { p: p.price, date: p.date, d: p.d, ref: p.d, src: p.src, spec: p.spec, size: p.size, f: p.f };
      }
      var ck = p.kind + '|' + p.place + '|' + p.spec;
      specCount[ck] = (specCount[ck] || 0) + 1;
      var ak = p.kind + '|*|' + p.spec;
      specCount[ak] = (specCount[ak] || 0) + 1;
    });
    function rank(d) { var i = VALVE_PRI.indexOf(d); return i < 0 ? 99 : i; }
    m.thick.forEach(function (t) { thick[t.kind + '|' + t.spec + '|' + t.size] = t.f; });
    m.priceIdx = price; m.specCount = specCount; m.thickIdx = thick;
    m.skRules = m.skRulesRaw && m.skRulesRaw.length ? UA.compileSkRules(m.skRulesRaw) : null;
    m.skAdj = m.skAdjRaw ? UA.compileSkAdj(m.skAdjRaw) : null;
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
    if (hit) return { p: hit.p, date: hit.date, ref: null, key: key, src: hit.src, d: hit.d };
    if (kind === 'valve') {
      var fb = m.priceIdx['valve|*|' + spec + '|' + sizeKey(size) + '|' + thickKey(f)];
      if (fb) return { p: fb.p, date: fb.date, ref: fb.ref, key: key, src: fb.src, d: fb.d };
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

  // この見積書（内訳シート）で使っている仕様の一覧
  UA.sheetSpecs = function (model) {
    var out = [], idx = {}, fam = 'pipe', place = '';
    if (!model) return out;
    model.recs.forEach(function (r) {
      if (r.t === 'area') { fam = famOfArea(r.a); place = ''; return; }
      if (r.t === 'sys') place = '';
      if ((r.t === 'spec' || r.t === 'sys') && r.e) {
        if (r.c) place = r.c;
        var key = fam + '|' + place + '|' + r.e;
        if (!idx[key]) { idx[key] = { spec: r.e, place: place, fam: fam, n: 0, rows: [] }; out.push(idx[key]); }
        idx[key].n++; idx[key].rows.push(r.row);
      }
    });
    return out;
  };
  // 仕様を「保温材」「表面材（貼）」「仕上げ」「目印」に分ける（画面で見やすく並べるため）
  var MAT_RE = /(Gw|Rw|PF|FP|ＧＷ|ＲＷ|グラスウール|ロックウール|ポリスチレンフォーム|フェノールフォーム|アクリア)[^＋+]*?(筒|ロール|帯|板|フェルト|ブランケット|保温材)|エアロフレックス|アーマフレックス|ウェーブロール|波形[^＋+]*板/;
  function matFam(t) {
    if (/Gw|ＧＷ|グラスウール|アクリア|ウェーブ|波形/.test(t)) return 'GW';
    if (/Rw|ＲＷ|ロックウール/.test(t)) return 'RW';
    if (/PF|ポリスチレン/.test(t)) return 'PF';
    if (/FP|フェノール/.test(t)) return 'FP';
    if (/エアロ|アーマ/.test(t)) return 'ゴム';
    return '';
  }
  UA.specParts = function (spec) {
    var res = { pre: [], facing: '', mat: '', fam: '', rest: [], tags: [] };
    String(spec || '').split(/[＋+]/).map(function (x) { return x.trim(); }).filter(Boolean).forEach(function (p) {
      if (/^ヒーター/.test(p)) { res.tags.push('ヒーター'); return; }
      if (p === '断熱鋲') { res.tags.push('断熱鋲'); return; }
      if (p === '鋲') { res.pre.push('鋲'); return; }
      var m = p.match(MAT_RE);
      if (m && !res.mat) { res.facing = p.slice(0, m.index); res.mat = p.slice(m.index); res.fam = matFam(res.mat); return; }
      if (m && res.mat) { if (res.tags.indexOf('2層') < 0) res.tags.push('2層'); res.rest.push(p); return; }
      if (/サンダム|遮音/.test(p)) { res.tags.push('遮音'); res.rest.push(p); return; }
      res.rest.push(p);
    });
    if (!res.mat && res.rest.length) { res.mat = res.rest.shift(); res.fam = matFam(res.mat); }
    return res;
  };
  // 仕様の候補（グループ分け）
  //  opt = { item, place, fam, sheet:[sheetSpecs], cur:現在の塊の仕様, cat:'民間'|'官庁'|'', local:{add,hide,pin}, showHidden }
  UA.specKey = function (kind, place, spec) { return kind + '|' + place + '|' + spec; };
  UA.specList = function (m, opt) {
    var kind = KIND_OF_ITEM[opt.item] || opt.item, place = opt.place, local = opt.local || { add: [], hide: {}, pin: {} };
    var seen = {}, groups = [];
    function catOK(c) { return !c || !opt.cat || c === opt.cat; }
    function cnt(spec, here) { return here ? (m.specCount[kind + '|' + place + '|' + spec] || 0) : (m.specCount[kind + '|*|' + spec] || 0); }
    // 1) この見積書で使用中（同じ配管/ダクトの区分）。今の塊 → 同じ施工箇所 → ほかの施工箇所
    var sh = (opt.sheet || []).filter(function (x) { return x.fam === opt.fam; });
    sh.sort(function (a, b) {
      var ca = a.spec === opt.cur ? 0 : a.place === place ? 1 : 2, cb = b.spec === opt.cur ? 0 : b.place === place ? 1 : 2;
      return ca - cb || b.n - a.n;
    });
    var g1 = [], g1b = [];
    sh.forEach(function (x) {
      if (seen[x.spec]) return; seen[x.spec] = 1;
      var it = { spec: x.spec, sheet: true, cur: x.spec === opt.cur, place: x.place, n: x.n, here: x.place === place };
      (it.cur || it.here ? g1 : g1b).push(it);
    });
    if (g1.length) groups.push({ key: 'sheet', title: 'この見積書で使用中（' + place + '）', items: g1 });
    // 候補（マスタ＋このパソコンで追加）
    var cands = [];
    // BOX・チャンバーは角ダクトの仕様も候補にする
    var kinds = kind === 'box' ? ['box', 'rect'] : [kind];
    function kindOK(k) { return k === '共通' || kinds.indexOf(k) >= 0; }
    m.specs.forEach(function (s) {
      if (!s.use || !kindOK(s.kind)) return;
      if (!(s.place === place || s.place === '（全箇所）')) return;
      if (!catOK(s.cat)) return;
      cands.push({ spec: s.spec, kind: s.kind, place: s.place, cat: s.cat, prio: s.prio, src: 'master' });
    });
    (local.add || []).forEach(function (a) {
      if (!kindOK(a.kind)) return;
      if (!(a.place === place || a.place === '（全箇所）')) return;
      if (!catOK(a.cat)) return;
      cands.push({ spec: a.spec, kind: a.kind, place: a.place, cat: a.cat, prio: 0, src: 'local', id: a.id });
    });
    var hidden = [];
    cands = cands.filter(function (c) {
      c.key = UA.specKey(c.kind, c.place, c.spec);
      c.here = c.place === place; c.n = cnt(c.spec, c.here); c.pin = !!(local.pin || {})[c.key];
      if ((local.hide || {})[c.key]) { hidden.push(c); return false; }
      return true;
    });
    var ord = (local.order || {})[kind + '|' + place] || [], oi = {};
    ord.forEach(function (sp, i) { if (oi[sp] == null) oi[sp] = i; });
    cands.sort(function (a, b) {
      if (a.pin !== b.pin) return a.pin ? -1 : 1;
      var oa = oi[a.spec], ob = oi[b.spec];
      if (oa != null || ob != null) return (oa == null ? 1e6 : oa) - (ob == null ? 1e6 : ob);
      var pa = a.prio > 0 ? a.prio : 9999, pb = b.prio > 0 ? b.prio : 9999;
      if (pa !== pb) return pa - pb;
      if (a.here !== b.here) return a.here ? -1 : 1;
      if (a.src !== b.src) return a.src === 'local' ? -1 : 1;
      return b.n - a.n;
    });
    var g2 = [], g3 = [];
    cands.forEach(function (c) {
      if (seen[c.spec]) return;
      seen[c.spec] = 1;
      (c.here || c.pin ? g2 : g3).push(c);
    });
    if (g2.length) groups.push({ key: 'place', title: place + 'の候補', items: g2 });
    if (g1b.length) groups.push({ key: 'sheet', title: 'この見積書で使用中（ほかの施工箇所）', items: g1b });
    if (g3.length) groups.push({ key: 'all', title: '全箇所の候補', items: g3 });
    if (opt.showHidden && hidden.length) groups.push({ key: 'hidden', title: '非表示にした候補', items: hidden });
    groups.hiddenCount = hidden.length;
    return groups;
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

  /* ========== 保温積算資料（単価表） ==========
     読み込んだ「保温積算資料」の表（tbl_○_○）から、仕様・施工箇所・サイズ・保温厚で単価を引く。
     どの仕様をどの表に当てるかは「対応表」（設定マスタ M_積算資料対応、無ければ下の既定）で決める。 */

  // 照合用の正規化（全角半角・空白・＋・半角カナを統一。「ヒーター＋」は外す）
  function nk(s) {
    var t = String(s == null ? '' : s);
    try { t = t.normalize('NFKC'); } catch (e) { t = toHalf(t); }
    t = t.replace(/[\s　]+/g, '').replace(/\+/g, '＋');
    t = t.replace(/ヒーター＋/g, '').replace(/ポリフイルム/g, 'ポリフィルム').replace(/ガラスクロス巻/g, 'ガラスクロス');
    return t;
  }
  UA.nk = nk;
  // Excel の Like と同じ書き方（* と ?）のパターン
  function likeRe(p) {
    var s = nk(p).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
    return new RegExp('^' + s + '$');
  }
  UA.likeRe = likeRe;

  // ---- 保温積算資料の読み込み（SheetJS の bookFiles:true で読んだブック） ----
  UA.parseSekisan = function (wb, X) {
    var files = wb.files || {};
    function raw(p) { return files[p] || files['/' + p] || null; }
    function text(p) {
      var f = raw(p); if (!f) return '';
      var c = f.content != null ? f.content : f._data;
      if (typeof c === 'string') return c;
      try { return new TextDecoder('utf-8').decode(c instanceof Uint8Array ? c : new Uint8Array(c)); } catch (e) { return ''; }
    }
    function attr(x, name) { var m = x.match(new RegExp('\\s' + name + '="([^"]*)"')); return m ? m[1] : ''; }
    function unxml(s) { return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'); }
    function resolve(base, target) {
      if (target.charAt(0) === '/') return target.slice(1);
      var parts = base.split('/'); parts.pop();
      target.split('/').forEach(function (s) { if (s === '..') parts.pop(); else if (s !== '.') parts.push(s); });
      return parts.join('/');
    }
    function rels(p) {
      var out = {}, x = text(p), re = /<Relationship\s([^>]*)\/?>/g, m;
      while ((m = re.exec(x))) out[attr(' ' + m[1], 'Id')] = { target: attr(' ' + m[1], 'Target'), type: attr(' ' + m[1], 'Type') };
      return out;
    }
    var wbx = text('xl/workbook.xml'), wrel = rels('xl/_rels/workbook.xml.rels');
    var sheetFile = {}, re = /<sheet\s([^>]*)\/?>/g, mm;
    while ((mm = re.exec(wbx))) {
      var a = ' ' + mm[1], rid = attr(a, 'r:id'), nm = unxml(attr(a, 'name'));
      if (wrel[rid]) sheetFile[nm] = resolve('xl/workbook.xml', wrel[rid].target);
    }
    function cv(ws, r, c) { var cell = ws[X.utils.encode_cell({ r: r, c: c })]; return cell ? cell.v : null; }
    function cs(ws, r, c) { var v = cv(ws, r, c); return v == null ? '' : String(v).trim(); }
    var out = { tables: {}, secs: {}, n: 0 };
    wb.SheetNames.forEach(function (sn) {
      var ws = wb.Sheets[sn], sf = sheetFile[sn];
      if (!ws || !sf) return;
      var sec = (sn.match(/§\s*(\d+)/) || [])[1] || '';
      // 章の題（例「§2 給水・冷水・冷温水：GW/屋内露出」）
      for (var r0 = 0; r0 < 6 && sec; r0++) {
        var t0 = cs(ws, r0, 0);
        if (t0.indexOf('§') === 0) { out.secs[sec] = t0.replace(/^§\s*\d+\s*/, ''); break; }
      }
      var srel = rels(sf.replace(/([^/]+)$/, '_rels/$1.rels'));
      Object.keys(srel).forEach(function (k) {
        if (!/\/table$/.test(srel[k].type)) return;
        var tx = text(resolve(sf, srel[k].target));
        if (!tx) return;
        var dn = attr(tx, 'displayName') || attr(tx, 'name'), ref = attr(tx, 'ref');
        if (!ref) return;
        var rg = X.utils.decode_range(ref);
        var cols = [], cre = /<tableColumn\s([^>]*)>/g, cm;
        while ((cm = cre.exec(tx))) cols.push(unxml(attr(' ' + cm[1], 'name')));
        if (!cols.length) for (var c = rg.s.c; c <= rg.e.c; c++) cols.push(cs(ws, rg.s.r, c));
        var rows = [];
        for (var r = rg.s.r + 1; r <= rg.e.r; r++) {
          var row = [cs(ws, r, rg.s.c)];
          for (var c2 = rg.s.c + 1; c2 <= rg.e.c; c2++) { var v = cv(ws, r, c2); row.push(typeof v === 'number' ? v : (v != null && v !== '' && isFinite(+v) ? +v : null)); }
          rows.push(row);
        }
        var title = '';
        for (var r1 = rg.s.r - 1; r1 >= Math.max(0, rg.s.r - 6); r1--) { var tt = cs(ws, r1, 0); if (tt.charAt(0) === '【') { title = tt; break; } }
        var comp = '';
        for (var r2 = rg.e.r + 1; r2 <= rg.e.r + 2; r2++) { var tc = cs(ws, r2, 0); if (tc.indexOf('構成') === 0) { comp = tc; break; } }
        var id = (title.match(/^【([^】]+)】/) || [])[1] || String(dn).replace(/^tbl_/, '').replace('_', '-');
        var unit = (title.match(/（(円\/[^）]+)）/) || [])[1] || '';
        var name = title.replace(/^【[^】]+】/, '').replace(/（円\/[^）]+）/, '').trim();
        if (name.indexOf(id + ' ') === 0) name = name.slice(id.length + 1);
        out.tables[id] = { id: id, sec: sec, name: name, unit: unit, cols: cols, rows: rows, comp: comp };
        out.n++;
      });
    });
    return out;
  };

  // 列見出しを「仕上げ」と「厚み」に分ける（例「(ALK貼)/AL粘着テープ/カラー亀甲金網10mm 同 50mm」）
  function colDesc(cols) {
    var out = [], prev = '';
    for (var j = 0; j < cols.length; j++) {
      var h = String(cols[j] || '').trim(), m, fin = h, t = null;
      if ((m = h.match(/^(\d+(?:\.\d+)?)\s*mm$/))) { fin = ''; t = m[1]; }
      else if ((m = h.match(/^同\s*(\d+(?:\.\d+)?)\s*mm$/))) { fin = prev; t = m[1]; }
      else if ((m = h.match(/^(.*\S)\s+(?:同\s*)?(\d+(?:\.\d+)?)\s*mm$/))) { fin = m[1]; t = m[2]; }
      if (fin) prev = fin;
      out.push({ fin: fin, t: t, h: h });
    }
    return out;
  }
  // 表の保温厚が固定のとき（見出しの「×50mm」や構成の「厚み50mm」）
  function fixedT(tbl, col) {
    var m = String(col && col.h || '').match(/×\s*(\d+)\s*mm/) || String(tbl.cols[0] || '').match(/厚み\s*(\d+)\s*mm/) || String(tbl.comp || '').match(/厚み\s*(\d+)\s*mm/);
    if (!m && col && col.t == null) { for (var j = 1; j < tbl.cols.length && !m; j++) m = String(tbl.cols[j]).match(/×\s*(\d+)\s*mm/); }
    return m ? m[1] : null;
  }
  // 角ダクト等の行（保温材）を仕様から決める
  function matLabel(spec) {
    var t = nk(spec);
    if (/Rwフェルト|Rwロール/.test(t)) return 'ロックウールフェルト1号';
    if (/Rw板2号|Rw板(1[2-9]\d|[2-9]\d\d)K/.test(t)) return 'ロックウール板2号';
    if (/Rw板/.test(t)) return 'ロックウール板1号';
    if (/Rwブランケット|ロックウールブランケット/.test(t)) return 'ロックウールブランケット1号';
    if (/Gwロール32K/.test(t)) return 'グラスウールロール32K';
    if (/Gwロール/.test(t)) return 'グラスウールロール24K';
    if (/Gw板32K/.test(t)) return 'グラスウール板32K';
    if (/Gw板/.test(t)) return 'グラスウール板40K';
    if (/ポリスチレンフォーム/.test(t)) return 'ポリスチレンフォーム3号';
    return '';
  }
  UA.matLabel = matLabel;

  // ---- 対応表（既定）----
  // [品目, 種別, 場所, 仕様（* と ? が使える）, 表ID, 列（仕上げ）, 行, 上被の表ID, 掛率, 加算, サイズ下限, サイズ上限, メモ]
  //  品目：pipe / valve / flange / rect / round / box（カンマ区切り可、* はすべて）
  //  場所：屋内露出 / 隠蔽 / ピット / 屋外露出 / 多湿 / 内貼 / *（M_施工箇所の「積算資料の場所」で決まる）
  //  表ID：「配管」と書くと、同じ仕様の配管の単価を使う（掛率をかける）
  //  列：空欄＝保温厚の列。文字＝その仕上げの列（厚み付きの見出しは保温厚も合わせる）
  //  行：空欄＝サイズ。「自動」＝仕様の保温材（Gw板40K→グラスウール板40K など）＋保温厚。その他＝行の名前
  function skDefaultRules() {
    var R = [];
    function add(items, ds, place, pat, id, col, row, over, factor, plus, memo) { R.push([items, ds || '', place, pat, id, col || '', row || '', over || '', factor || '', plus || '', '', '', memo || '']); }
    function P(place, pat, id, over, memo) { add('pipe', '', place, pat, id, '', '', over, '', '', memo); }
    var M16 = '表は亀甲金網16mm';
    // --- 配管：グラスウール筒 ---
    P('屋内露出', 'Gw筒＋ポリ*＋整形原紙＋綿*', '2-1'); P('屋内露出', 'Gw筒＋ポリ*＋整形原紙＋*テープ', '2-2');
    P('屋内露出', 'Gw筒＋ポリ*＋ALGC原紙＋ALGC粘着テープ', '2-3'); P('屋内露出', 'Gw筒＋ポリ*＋ファインカバー', '2-4');
    P('屋内露出', 'Gw筒＋ポリ*＋着色ALGC原紙＋着色ALGC粘着テープ', '2-5');
    P('屋内露出', 'ALK-P付Gw筒＋カラー金網10m/m', '2-6'); P('屋内露出', 'ALGC-P付Gw筒＋*カラー金網10m/m', '2-7');
    P('屋内露出', 'Gw筒＋整形原紙＋綿*', '1-1'); P('屋内露出', 'Gw筒＋整形原紙＋*テープ', '1-2');
    P('屋内露出', 'Gw筒＋ALGC原紙＋ALGC粘着テープ', '1-3'); P('屋内露出', 'Gw筒＋ファインカバー', '1-4');
    P('屋内露出', 'Gw筒＋着色ALGC原紙＋着色ALGC粘着テープ', '1-5'); P('屋内露出', '着色ALGC付Gw筒＋着色ALGC粘着テープ', '1-5', '', '表は着色ALGC化粧原紙');
    P('屋内露出', 'ALK付Gw筒＋カラー金網10m/m', '1-6'); P('屋内露出', 'ALGC付Gw筒＋*カラー金網10m/m', '1-7');
    P('隠蔽', 'Gw筒＋ポリ*＋*テープ', '4-1'); P('隠蔽', 'ALK-P付Gw筒＋亀甲金網16m/m', '4-2'); P('隠蔽', 'ALK-P付Gw筒＋カラー金網16m/m', '4-2', '', M16);
    P('隠蔽', 'ALGC-P付Gw筒＋*亀甲金網16m/m', '4-3'); P('隠蔽', 'ALGC-P付Gw筒＋ALGC粘着テープ', '4-3', '', '表は亀甲金網16mm込み');
    P('隠蔽', 'Gw筒＋*テープ', '3-1'); P('隠蔽', 'ALGC付Gw筒＋ALGC粘着テープ', '3-2'); P('隠蔽', 'ALGC付Gw筒＋*亀甲金網16m/m', '3-3');
    P('隠蔽', 'ALK付Gw筒＋亀甲金網16m/m', '3-4'); P('隠蔽', 'ALK付Gw筒＋カラー金網16m/m', '3-4', '', M16);
    P('ピット', 'Gw筒＋ポリ*＋着色ALGC*', '5-1'); P('ピット', '着色ALGC-P付Gw筒＋着色ALGC粘着テープ', '5-1', '', '表はポリエチレンフィルム＋着色ALGCテープ');
    P('ピット', 'ALW裏貼Gw筒＋カラー金網16m/m', '5-2');
    P('*', 'Gw筒＋ポリ*＋カラー鉄板', '6-1', '19-3'); P('*', 'Gw筒＋ポリ*＋*ステンレス*', '6-1', '19-5'); P('*', 'Gw筒＋ポリ*＋ガルバ*', '6-1', '19-4'); P('*', 'Gw筒＋ポリ*＋亜鉛鉄板', '6-1', '19-2');
    P('*', 'Gw筒＋ポリテープ', '6-1'); P('*', 'Gw筒＋ポリフィルム', '6-1');
    P('*', 'ALW裏貼Gw筒＋カラー鉄板', '6-1', '19-3', 'ALW貼は§6で代用'); P('*', 'ALW裏貼Gw筒＋ガルバ*', '6-1', '19-4', 'ALW貼は§6で代用');
    P('*', 'Gw筒＋カラー鉄板', '7-1', '19-3'); P('*', 'Gw筒＋*ステンレス*', '7-1', '19-5'); P('*', 'Gw筒＋ガルバ*', '7-1', '19-4'); P('*', 'Gw筒＋亜鉛鉄板', '7-1', '19-2');
    P('*', 'ALK付Gw筒＋カラー鉄板', '7-1', '19-3', 'ALK貼の分は含まず');
    P('*', 'Gw筒＋ファインジャケット', '7-1', '19-1', 'ファインカバーJを上被で計上'); P('*', 'Gw筒', '7-1');
    // --- 配管：ロックウール筒 ---
    P('屋内露出', 'Rw筒＋ポリ*＋整形原紙＋綿*', '9-1'); P('屋内露出', 'Rw筒＋ポリ*＋整形原紙＋*テープ', '9-2');
    P('屋内露出', 'Rw筒＋ポリ*＋ALGC原紙＋ALGC粘着テープ', '9-3'); P('屋内露出', 'Rw筒＋ポリ*＋ファインカバー', '9-4');
    P('屋内露出', 'Rw筒＋ポリ*＋着色ALGC原紙＋着色ALGC粘着テープ', '9-5'); P('屋内露出', 'ALK-P付Rw筒＋カラー金網10m/m', '9-6');
    P('屋内露出', 'Rw筒＋整形原紙＋綿*', '8-1'); P('屋内露出', 'Rw筒＋整形原紙＋*テープ', '8-2'); P('屋内露出', 'Rw筒＋ALGC原紙＋ALGC粘着テープ', '8-3');
    P('屋内露出', 'Rw筒＋ファインカバー', '8-4'); P('屋内露出', 'Rw筒＋着色ALGC原紙＋着色ALGC粘着テープ', '8-5'); P('屋内露出', 'ALK付Rw筒＋カラー金網10m/m', '8-6');
    P('隠蔽', 'Rw筒＋ポリ*＋*テープ', '10-2'); P('隠蔽', 'Rw筒＋*テープ', '10-1'); P('隠蔽', 'ALGC付Rw筒＋ALGC粘着テープ', '10-3');
    P('隠蔽', 'ALGC-P付Rw筒＋ALGC粘着テープ', '10-3', '', 'ポリ付きの表なし（ALGC貼の表）'); P('隠蔽', 'ALK付Rw筒＋亀甲金網16m/m', '10-4'); P('隠蔽', 'ALK付Rw筒＋カラー金網*', '10-4', '', M16);
    P('ピット', 'Rw筒＋ポリ*＋着色ALGC*', '11-1'); P('ピット', '着色ALGC-P付Rw筒＋着色ALGC粘着テープ', '11-1', '', '表はポリエチレンフィルム＋着色ALGCテープ');
    P('ピット', 'ALW裏貼Rw筒＋カラー金網16m/m', '11-2');
    P('*', 'Rw筒＋ポリ*＋カラー鉄板', '12-1', '19-3'); P('*', 'Rw筒＋ポリ*＋*ステンレス*', '12-1', '19-5'); P('*', 'Rw筒＋ポリ*＋ガルバ*', '12-1', '19-4'); P('*', 'Rw筒＋ポリ*＋亜鉛鉄板', '12-1', '19-2');
    P('*', 'Rw筒＋ポリテープ', '12-1'); P('*', 'Rw筒＋ポリフィルム', '12-1');
    P('*', 'Rw筒＋カラー鉄板', '13-1', '19-3'); P('*', 'Rw筒＋*ステンレス*', '13-1', '19-5'); P('*', 'Rw筒＋ガルバ*', '13-1', '19-4'); P('*', 'Rw筒＋亜鉛鉄板', '13-1', '19-2');
    P('*', 'ALK付Rw筒＋カラー鉄板', '13-1', '19-3', 'ALK貼の分は含まず'); P('*', 'Rw筒', '13-1');
    // --- 配管：ポリスチレンフォーム筒（PF筒） ---
    P('屋内露出', 'PF筒＋ポリ*＋整形原紙＋綿*', '14-2'); P('屋内露出', 'PF筒＋整形原紙＋綿*', '14-1'); P('屋内露出', 'PF筒＋*ファインカバー', '14-3');
    P('屋内露出', 'PF筒＋ポリ*＋整形原紙＋*テープ', '14-5'); P('屋内露出', 'PF筒＋整形原紙＋*テープ', '14-4');
    P('屋内露出', 'PF筒＋ポリ*＋ALGC原紙＋ALGC粘着テープ', '14-7'); P('屋内露出', 'PF筒＋ALGC原紙＋ALGC粘着テープ', '14-6');
    P('屋内露出', 'PF筒＋ポリ*＋着色ALGC原紙＋着色ALGC粘着テープ', '14-8'); P('屋内露出', 'ALK-P付PF筒＋カラー金網10m/m', '14-9');
    P('隠蔽', 'PF筒＋ポリ*＋*テープ', '15-2'); P('隠蔽', 'PF筒＋ビニルテープ', '15-3'); P('隠蔽', 'PF筒＋*テープ', '15-1');
    P('隠蔽', 'ALGC付PF筒*＋ALGC粘着テープ', '15-4'); P('隠蔽', 'PF筒＋アルミクラフト*', '15-5'); P('隠蔽', 'ALK-P付PF筒＋亀甲金網16m/m', '15-6');
    P('ピット', 'PF筒＋ポリ*＋着色ALGC*', '16-1'); P('ピット', '着色ALGC-P付PF筒＋着色ALGC粘着テープ', '16-1', '', '表はポリエチレンフィルム＋着色ALGCテープ');
    P('ピット', 'PF筒＋着色ALGC*', '16-1', '', '表はポリエチレンフィルム付き'); P('ピット', 'ALW裏貼PF筒＋カラー金網16m/m', '16-2');
    P('*', 'PF筒＋ポリ*＋カラー鉄板', '17-1', '19-3'); P('*', 'PF筒＋ポリ*＋*ステンレス*', '17-1', '19-5'); P('*', 'PF筒＋ポリ*＋ガルバ*', '17-1', '19-4'); P('*', 'PF筒＋ポリ*＋亜鉛鉄板', '17-1', '19-2');
    P('*', 'PF筒＋ポリテープ', '17-1'); P('*', 'PF筒＋ポリフィルム', '17-1');
    P('*', 'PF筒＋カラー鉄板', '18-1', '19-3'); P('*', 'PF筒＋*ステンレス*', '18-1', '19-5'); P('*', 'PF筒＋ガルバ*', '18-1', '19-4'); P('*', 'PF筒＋亜鉛鉄板', '18-1', '19-2'); P('*', 'PF筒', '18-1');
    // --- 配管：上被だけ・パイプガード・防食 ---
    P('*', 'ファインカバー', '19-1'); P('*', '亜鉛鉄板', '19-2'); P('*', 'カラー鉄板', '19-3'); P('*', 'ガルバ*', '19-4'); P('*', 'ステンレス*', '19-5');
    add('pipe', '', '*', 'パイプガード*', '20-1', '直管部 20mm');
    add('pipe', '', '*', '防食ビニ*', '21-1', '防食ビニールテープ 1/2・3回巻'); add('pipe', '', '*', '防食テープ', '21-1', '防食ビニールテープ 1/2・3回巻');
    add('pipe', '', '*', '自己融着テープ*0.4*', '21-1', '自己融着テープ 0.4mm 1/2・1回巻'); add('pipe', '', '*', '自己融着テープ*1.0*', '21-1', '自己融着テープ 1.0mm 1/2・2回巻');
    add('pipe', '', '*', '*ペトロ*シート*', '21-1', '継手部 ペトロ系ペースト/ペトロラタムシート/プラスチックテープ');
    add('pipe', '', '*', '*ペトロ*', '21-1', '配管部 ペトロ系ペースト/ペトロラタムテープ/プラスチックテープ');
    // --- 弁・フランジ（65A以上。小さいGVで筒の仕様は配管×1.2） ---
    add('valve', 'GV', '*', '*筒＋*', '配管', '', '', '', 1.2, '', '配管の単価×1.2（GV）');
    function V(items, ds, gw, rw) {
      var c1 = 'GWロール(ALK)24K×50mm+カラー亀甲金網10mm', c2 = 'GW帯(ALGC)40K×50mm+カラー亀甲金網10mm';
      if (gw === '24-1') {
        add(items, ds, '*', '*Gw*＋*亜鉛鉄板', gw, 'グラスウール保温帯40K×50mm+ポリエチレンフィルム 亜鉛鉄板');
        add(items, ds, '*', '*Gw*＋*カラー鉄板', gw, 'カラー鉄板'); add(items, ds, '*', '*Gw*＋*ガルバ*', gw, 'ガルバニウム鋼板'); add(items, ds, '*', '*Gw*＋*ステンレス*', gw, 'ステンレス鋼板 SUS304');
      } else {
        add(items, ds, '*', '*Gw帯*＋*カラー鉄板', gw, 'グラスウール保温帯40K×50mm+ポリエチレンフィルム カラー鉄板');
        add(items, ds, '*', '*Gw帯*＋*ガルバ*', gw, 'ガルバニウム鋼板'); add(items, ds, '*', '*Gw帯*＋*ステンレス*', gw, 'ステンレス鋼板 SUS304');
        add(items, ds, '屋内露出', 'ALK*Gwロール*＋カラー金網10m/m', gw, c1);
        add(items, ds, '*', 'ALK*Gwロール*', gw, c1, '', '', '', '', '表はカラー亀甲金網10mm');
        add(items, ds, '屋内露出', 'ALGC*Gw帯*＋*カラー金網10m/m', gw, c2);
        add(items, ds, '*', '*ALGC*Gw帯*', gw, c2, '', '', '', '', '表はALGC貼＋カラー亀甲金網10mm');
        add(items, ds, '*', '*Gw帯*ALGC*', gw, c2, '', '', '', '', '表はALGC貼＋カラー亀甲金網10mm');
      }
      if (rw === '27-1') {
        add(items, ds, '*', '*Rw*＋*亜鉛鉄板', rw, 'ロックウール保温帯1号×50mm+ポリエチレンフィルム 亜鉛鉄板');
      } else {
        add(items, ds, '*', '*Rw*＋*亜鉛鉄板', rw, 'ロックウール保温帯1号×50mm+ポリエチレンフィルム 亜鉛鉄板');
        add(items, ds, '*', '*Rw*＋*カラー金網10m/m', rw, 'RW帯(ALGC)1号50mm+カラー亀甲金網10mm');
      }
      add(items, ds, '*', '*Rw*＋*カラー鉄板', rw, 'カラー鉄板'); add(items, ds, '*', '*Rw*＋*ガルバ*', rw, 'ガルバニウム鋼板'); add(items, ds, '*', '*Rw*＋*ステンレス*', rw, 'ステンレス鋼板 SUS304');
    }
    V('valve,flange', 'BV,BAV,バタフライ*,FLG,フランジ,FJ*,SUSFJ', '23-1', '26-1');
    V('valve', 'ラインポンプ,ポンプ', '24-1', '27-1');
    V('valve', '', '22-1', '25-1');
    V('flange', '', '23-1', '26-1');
    // --- 丸ダクト ---
    function D(place, pat, id, col, memo) { add('round', '', place, pat, id, col, '', '', '', '', memo); }
    var K10 = 'カラー亀甲金網10mm', K16 = '亀甲金網16mm';
    D('屋内露出', 'ALK付Gwロール32K＋カラー金網10m/m', '35-5', '(ALK貼)/AL粘着テープ/' + K10); D('屋内露出', 'ALW裏貼Gwロール32K＋カラー金網10m/m', '35-5', '(ALW貼)/ALW粘着テープ/' + K10);
    D('屋内露出', 'ALGC付Gwロール32K＋*カラー金網10m/m', '35-5', '(ALGC貼)/ALGC粘着テープ/' + K10);
    D('屋内露出', 'ALK付Gwロール*＋カラー金網10m/m', '35-3', '(ALK貼)/AL粘着テープ/' + K10); D('屋内露出', 'ALW裏貼Gwロール*＋カラー金網10m/m', '35-3', '(ALW貼)/ALW粘着テープ/' + K10);
    D('屋内露出', 'ALGC付Gwロール*＋*カラー金網10m/m', '35-3', '(ALGC貼)/ALGC粘着テープ/' + K10);
    D('屋内露出', '*Gwロール32K＋整形原紙＋綿*', '35-4', '整形原紙/綿布'); D('屋内露出', '*Gwロール32K＋整形原紙＋*', '35-4', '整形原紙/アルミガラスクロス');
    D('屋内露出', '*Gwロール*＋整形原紙＋綿*', '35-1', '整形原紙/綿布'); D('屋内露出', '*Gwロール*＋整形原紙＋*', '35-1', '整形原紙/アルミガラスクロス');
    D('屋内露出', '*Gwロール*＋亜鉛鉄板', '35-2', '亜鉛鉄板'); D('屋内露出', '*Gwロール*＋カラー鉄板', '35-2', 'カラー鉄板'); D('屋内露出', '*Gwロール*＋ガルバ*', '35-2', 'ガルバニウム鋼板'); D('屋内露出', '*Gwロール*＋*ステンレス*', '35-2', 'ステンレス鋼板');
    D('屋内露出', '*Gw帯*＋カラー鉄板', '35-7', 'カラー鉄板'); D('屋内露出', '*Gw帯*＋ガルバ*', '35-7', 'ガルバニウム鋼板'); D('屋内露出', '*Gw帯*＋*ステンレス*', '35-7', 'ステンレス鋼板');
    D('屋内露出', '*波形*＋*カラー金網10m/m', '35-8', 'GW波形保温板40K(ALGC)/ALGC粘着テープ/' + K10); D('屋内露出', '*ウェーブ*＋*カラー金網10m/m', '35-8', 'GW波形保温板40K(ALGC)/ALGC粘着テープ/' + K10);
    D('屋内露出', '*波形*＋*ALGC粘着テープ', '35-8', 'GW波形保温板40K(ALGC)/ALGC粘着テープ'); D('屋内露出', '*ウェーブ*＋*ALGC粘着テープ', '35-8', 'GW波形保温板40K(ALGC)/ALGC粘着テープ');
    D('屋内露出', 'ALGC付Gw帯*＋*カラー金網10m/m', '35-8', 'GW帯40K(ALGC)/ALGC粘着テープ/' + K10); D('屋内露出', 'ALGC付Gw帯*＋ALGC粘着テープ', '35-8', 'GW帯40K(ALGC)/ALGC粘着テープ');
    D('屋内露出', '*着色ALGC*Gw帯*', '35-6', '着色ALGC化粧原紙/着色ALGC粘着テープ'); D('屋内露出', '*Gw帯*＋整形原紙＋綿*', '35-6', '整形原紙/綿布'); D('屋内露出', '*Gw帯*＋整形原紙＋*', '35-6', '整形原紙/アルミガラスクロス');
    D('屋内露出', 'Rw帯*＋亜鉛鉄板', '38-1', '亜鉛鉄板'); D('屋内露出', 'Rw帯*＋カラー鉄板', '38-1', 'カラー鉄板'); D('屋内露出', 'Rw帯*＋ガルバ*', '38-1', 'ガルバリウム鋼板'); D('屋内露出', 'Rw帯*＋*ステンレス*', '38-1', 'ステンレス鋼板');
    D('屋内露出', 'ALK付Rw*＋カラー金網10m/m', '38-2', 'RWフェルト1号(ALK貼)/ALK粘着テープ/' + K10);
    D('屋内露出', 'ALGC付Rw帯*＋*カラー金網10m/m', '38-2', 'RW帯1号(ALGC)/ALGC粘着テープ/' + K10); D('屋内露出', 'ALGC付Rw帯*＋ALGC粘着テープ', '38-2', 'RW帯1号(ALGC)/ALGC粘着テープ');
    D('屋内露出', 'ALGC付Rwフェルト*＋カラー金網10m/m', '38-2', 'RW帯1号(ALGC)/ALGC粘着テープ/' + K10, '表はRW帯1号(ALGC)');
    D('隠蔽', 'ALK付Gwロール32K＋*金網16m/m', '36-2', '(ALK貼)/AL粘着テープ/' + K16); D('隠蔽', 'ALW裏貼Gwロール32K＋*金網16m/m', '36-2', '(ALW貼)/ALW粘着テープ/' + K16);
    D('隠蔽', 'ALGC付Gwロール32K＋*亀甲金網16m/m', '36-2', '(ALGC貼)/ALGC粘着テープ/' + K16);
    D('隠蔽', 'ALK付Gwロール*＋亀甲金網16m/m', '36-1', '(ALK貼)/AL粘着テープ/' + K16); D('隠蔽', 'ALK付Gwロール*＋カラー金網*', '36-1', '(ALK貼)/AL粘着テープ/' + K16, M16);
    D('隠蔽', 'ALW裏貼Gwロール*＋*金網16m/m', '36-1', '(ALW貼)/ALW粘着テープ/' + K16); D('隠蔽', 'ALGC付Gwロール*＋*亀甲金網16m/m', '36-1', '(ALGC貼)/ALGC粘着テープ/' + K16);
    D('隠蔽', '*波形*＋*亀甲金網16m/m', '36-3', 'GW波形保温板40K(ALGC)/ALGC粘着テープ/' + K16); D('隠蔽', '*ウェーブ*＋*亀甲金網16m/m', '36-3', 'GW波形保温板40K(ALGC)/ALGC粘着テープ/' + K16);
    D('隠蔽', '*波形*＋*ALGC粘着テープ', '36-3', 'GW波形保温板40K(ALGC)/ALGC粘着テープ'); D('隠蔽', '*ウェーブ*＋*ALGC粘着テープ', '36-3', 'GW波形保温板40K(ALGC)/ALGC粘着テープ');
    D('隠蔽', 'ALGC付Gw帯*＋*亀甲金網16m/m', '36-3', 'GW帯40K(ALGC)/ALGC粘着テープ/' + K16); D('隠蔽', 'ALGC付Gw帯*＋ALGC粘着テープ', '36-3', 'GW帯40K(ALGC)/ALGC粘着テープ');
    D('隠蔽', 'ALK付Rw*＋*金網16m/m', '39-1', 'RWフェルト1号(ALK貼)/ALK粘着テープ/' + K16); D('隠蔽', 'Rw帯*＋アルミガラスクロス*＋亀甲金網16m/m', '39-1', 'RW帯1号/アルミガラスクロス/' + K16);
    D('隠蔽', 'Rw帯*＋亀甲金網16m/m', '39-1', 'RW帯1号/' + K16); D('隠蔽', 'ALGC付Rw*＋*亀甲金網16m/m', '39-2', 'ALGC粘着テープ/' + K16);
    D('隠蔽', 'ALGC付Rw帯*＋ALGC粘着テープ', '39-2', 'ALGC粘着テープ'); D('隠蔽', 'ALGC付Rwフェルト*＋カラー金網*', '39-2', 'ALGC粘着テープ/' + K16, '表はRW帯1号(ALGC)＋亀甲金網16mm');
    ['屋外露出', '多湿'].forEach(function (pl) {
      D(pl, '*Gwロール*＋ポリ*＋カラー鉄板', '37-1', 'ポリエチレンフィルム/カラー鉄板'); D(pl, '*Gwロール*＋ポリ*＋ガルバ*', '37-1', 'ポリエチレンフィルム/ガルバリウム鋼板'); D(pl, '*Gwロール*＋ポリ*＋*ステンレス*', '37-1', 'ポリエチレンフィルム/ステンレス鋼板');
      D(pl, '*Gw帯*＋ポリ*＋カラー鉄板', '37-2', 'ポリエチレンフィルム/カラー鉄板'); D(pl, '*Gw帯*＋ポリ*＋ガルバ*', '37-2', 'ポリエチレンフィルム/ガルバリウム鋼板'); D(pl, '*Gw帯*＋ポリ*＋*ステンレス*', '37-2', 'ポリエチレンフィルム/ステンレス鋼板');
      D(pl, '*Rw*＋ポリ*＋カラー鉄板', '40-1', 'ポリエチレンフィルム/カラー鉄板'); D(pl, '*Rw*＋ポリ*＋ガルバ*', '40-1', 'ポリエチレンフィルム/ガルバリウム鋼板'); D(pl, '*Rw*＋ポリ*＋*ステンレス*', '40-1', 'ポリエチレンフィルム/ステンレス鋼板');
    });
    D('*', 'エアロフレックス*', '53-1', ''); D('*', 'アーマフレックス*', '53-2', '');
    // --- 角ダクト・チャンバー（行は仕様の保温材＋保温厚） ---
    function Q(place, pat, id, col, memo, row) { add('rect,box', '', place, pat, id, col, row || '自動', '', '', '', memo); }
    [['亜鉛鉄板', '亜鉛鉄板'], ['カラー鉄板', 'カラー鉄板'], ['ガルバ*', 'ガルバニウム鋼板'], ['*ステンレス*', 'ステンレス板'], ['石膏ボード*', '石膏ボード(9mm)/アルミコーナー']].forEach(function (x) {
      Q('屋内露出', '*板*＋鋼枠＋' + x[0], '31-1b', '鋲/保温板/鋼枠/' + x[1]);
    });
    [['亜鉛鉄板', '亜鉛鉄板'], ['カラー鉄板', 'カラー鉄板'], ['ガルバ*', 'ガルバニウム鋼板'], ['*ステンレス*', 'ステンレス板']].forEach(function (x) {
      Q('屋内露出', '*板*＋' + x[0], '31-1', '鋲/保温板/' + x[1]);
    });
    Q('屋内露出', '*鋲＋*板*＋*寒冷紗*', '31-1', '鋲/保温板/角あて/目貼り/接着剤/寒冷紗'); Q('屋内露出', '*鋲＋*板*＋*アルミガラスクロス', '31-1', '鋲/保温板/角あて/目貼り/接着剤/アルミガラスクロス');
    Q('屋内露出', '*鋲＋ALK*＋カラー金網10m/m*', '31-2', '鋲/保温材(ALK貼)/カラー亀甲金網10mm'); Q('屋内露出', '*鋲＋ALW*＋カラー金網10m/m*', '31-2', '鋲/保温材(ALW貼)/カラー亀甲金網10mm');
    Q('屋内露出', '*鋲＋着色ALGC*＋着色ALGC粘着テープ*', '31-2', '鋲/保温板(着色ALGC貼)/アルミガラスクロス粘着テープ');
    Q('屋内露出', '*鋲＋ALGC*＋ALGC粘着テープ＋*亀甲金網*', '31-2', '鋲/保温板(ALGC貼)/アルミガラスクロス粘着テープ/亀甲金網10mm');
    Q('屋内露出', '*鋲＋ALGC*＋カラー金網10m/m*', '31-2', '鋲/保温材(ALGC貼)/カラー亀甲金網10mm');
    Q('屋内露出', '*鋲＋ALGC*＋ALGC粘着テープ*', '31-2', '鋲/保温板(ALGC貼)/アルミガラスクロス粘着テープ');
    Q('屋内露出', '*鋲＋ALK*＋*金網16m/m*', '31-2', '鋲/保温材(ALK貼)/カラー亀甲金網10mm', '露出の表はカラー亀甲金網10mm');
    Q('隠蔽', '*鋲＋*接着剤＋アルミガラスクロス', '32-1', '鋲/保温材/接着剤/アルミガラスクロス');
    Q('隠蔽', '*鋲＋ALGC*＋ALGC粘着テープ＋*亀甲金網16m/m', '32-1', '鋲/保温材(ALGC貼)/アルミガラスクロス粘着テープ/亀甲金網16mm');
    Q('隠蔽', '*鋲＋ALGC*＋亀甲金網16m/m', '32-1', '鋲/保温材(ALGC貼)/アルミガラスクロス粘着テープ/亀甲金網16mm', '表はALGC粘着テープ＋亀甲金網16mm');
    Q('隠蔽', '*鋲＋ALGC*＋ALGC粘着テープ*', '32-1', '鋲/保温材(ALGC貼)/アルミガラスクロス粘着テープ');
    Q('隠蔽', '*鋲＋*アルミクラフト*＋亀甲金網16m/m', '32-1', '鋲/保温材/アルミクラフト紙/亀甲金網16mm');
    Q('隠蔽', '*鋲＋ALK*＋亀甲金網16m/m', '32-1', '鋲/保温板(ALK貼)/亀甲金網16mm'); Q('隠蔽', '*鋲＋ALK*＋*金網*', '32-1', '鋲/保温板(ALK貼)/亀甲金網16mm', M16);
    ['屋外露出', '多湿'].forEach(function (pl) {
      [['カラー鉄板', 'カラー鉄板'], ['ガルバ*', 'ガルバ鋼板'], ['*ステンレス*', 'ステンレス板']].forEach(function (x) { Q(pl, '*鋲＋*ポリ*＋鋼枠＋' + x[0], '33-1', '鋲/保温材/ポリエチレンフィルム/鋼枠/' + x[1]); });
      [['カラー鉄板', 'カラー鉄板'], ['ガルバ*', 'ガルバ鉄板'], ['*ステンレス*', 'ステンレス板']].forEach(function (x) { Q(pl, '*鋲＋*ポリ*＋' + x[0], '33-1', '鋲/保温材/ポリエチレンフィルム/' + x[1]); });
    });
    Q('内貼', '*GC*板*＋アルミパンチング*', '34-1', '保温板/ガラスクロス/アルミスパンシーティング0.6mm/絶縁座金付スポット鋲', '表はアルミスパンシーティング');
    Q('内貼', '*GC*板*＋*銅亀甲金網*', '34-1', '鋲/保温板/ガラスクロス/鋼製亀甲金網10mm', '表は鋼製亀甲金網10mm');
    Q('内貼', '*GC*板*＋*カラー金網10m/m', '34-1', '鋲/保温板/ガラスクロス/カラー亀甲金網10mm');
    Q('内貼', '*GC*板*＋亀甲金網16m/m', '34-1', '鋲/保温板/ガラスクロス/亀甲金網16mm');
    Q('内貼', '*GC*板*', '34-1', '鋲/保温板/ガラスクロス');
    Q('*', 'エアロフレックス*', '52-1', '', '', '角ダクト'); Q('*', 'アーマフレックス*', '52-3', '', '', '角ダクト');
    return R;
  }
  UA.skDefaultRules = skDefaultRules;
  // 単価の加算・掛率（既定）[対象, 含む語, 加算(円), 掛率, メモ]
  UA.skDefaultAdj = function () {
    return [['仕様', 'ヒーター', 2000, '', 'ヒーター＋2,000円（取付費の行を別に立てるときは外す）'], ['仕様', '断熱鋲', 1500, '', '断熱鋲＋1,500円'], ['サイズ', 'BOX', 1000, '', 'BOX＋1,000円']];
  };

  function listOf(v) { return String(v == null ? '' : v).split(/[,、，]/).map(function (x) { return x.trim(); }).filter(Boolean); }
  // 対応表の行（配列 or シートの行）を照合用に整える
  UA.compileSkRules = function (rows) {
    return rows.map(function (r, i) {
      var items = listOf(r[0]), ds = listOf(r[1]);
      return {
        n: i + 1, items: items.length ? items : ['*'], ds: ds.map(likeRe), dsText: ds.join(','), place: trimAll(r[2]) || '*', pat: likeRe(r[3]), patText: String(r[3]),
        id: trimAll(r[4]), col: String(r[5] == null ? '' : r[5]).trim(), row: String(r[6] == null ? '' : r[6]).trim(), over: trimAll(r[7]),
        factor: num(r[8]) || 1, plus: num(r[9]) || 0, min: isEmpty(r[10]) ? null : num(r[10]), max: isEmpty(r[11]) ? null : num(r[11]), memo: String(r[12] == null ? '' : r[12]).trim()
      };
    }).filter(function (r) { return r.id && r.patText; });
  };
  UA.compileSkAdj = function (rows) {
    return rows.map(function (r) { return { target: trimAll(r[0]) || '仕様', word: nk(r[1]), plus: num(r[2]) || 0, factor: num(r[3]) || 1, memo: String(r[4] == null ? '' : r[4]).trim() }; })
      .filter(function (a) { return a.word && (a.plus || a.factor !== 1); });
  };

  // 施工箇所 → 積算資料の場所
  UA.placeGroup = function (m, place) {
    var p = (m && m.places || []).filter(function (x) { return x.name === place; })[0];
    if (p && p.grp) return p.grp;
    var t = String(place || '');
    if (/ピット|床下|暗渠/.test(t)) return 'ピット';
    if (/内貼/.test(t)) return '内貼';
    if (/屋外|屋上/.test(t)) return '屋外露出';
    if (/多湿|浴室/.test(t)) return '多湿';
    if (/隠蔽/.test(t)) return '隠蔽';
    if (/露出/.test(t)) return '屋内露出';
    if (/天井|シャフト|PS|ライニング/.test(t)) return '隠蔽';
    return '屋内露出';
  };

  function findRowBySize(tbl, size) {
    var n = parseFloat(toHalf(size));
    if (!isFinite(n)) return { i: -1 };
    var maxN = 0, sqm = -1;
    for (var i = 0; i < tbl.rows.length; i++) {
      var lab = String(tbl.rows[i][0]).trim(), v = parseFloat(toHalf(lab));
      if (/^(㎡|m2)$/.test(nk(lab))) { sqm = i; continue; }
      if (isFinite(v) && Math.abs(v - n) < 1e-9 && /^[\d.]+$/.test(toHalf(lab))) return { i: i };
      if (isFinite(v)) maxN = Math.max(maxN, v);
    }
    if (sqm >= 0 && n > maxN) return { i: sqm, area: Math.PI * n / 1000, sqm: true };
    return { i: -1 };
  }
  function findRowByLabel(tbl, label, f) {
    var want = nk(label), wantT = nk(label + f), i, lab;
    for (i = 0; i < tbl.rows.length; i++) { lab = nk(tbl.rows[i][0]).replace(/mm$/, ''); if (lab === wantT) return { i: i, t: f }; }
    for (i = 0; i < tbl.rows.length; i++) { lab = nk(tbl.rows[i][0]); if (lab === want) return { i: i, t: null }; }
    return { i: -1 };
  }
  function findCol(tbl, colText, f) {
    var cd = colDesc(tbl.cols), want = nk(colText), j, hit = -1;
    if (want) for (j = 1; j < cd.length; j++) if (nk(cd[j].h) === want) return { j: j, c: cd[j] };
    for (j = 1; j < cd.length; j++) {
      if (nk(cd[j].fin) !== want) continue;
      if (cd[j].t == null) return { j: j, c: cd[j] };
      if (String(parseFloat(cd[j].t)) === String(parseFloat(f))) return { j: j, c: cd[j] };
      if (hit < 0) hit = j;
    }
    return { j: -1, near: hit >= 0 ? cd[hit] : null };
  }
  function compText(tbl) {
    return String(tbl.comp || '').replace(/^構成[:：]\s*/, '').split('/').map(function (x) { return x.trim().replace(/^\d+\./, ''); })
      .filter(function (x) { return x && x.charAt(0) !== '※' && !/^厚み/.test(x); }).join('/');
  }
  function ceil10(x) { return (Math.ceil(x / 10 - 1e-9) * 10) || 0; }
  function fmtN(x) { return Math.round(x * 1e6) / 1e6; }

  // 単価を引く。q = { item, d, place, spec, size, f }
  // 戻り値 { p, formula, ref:{y,z,aa,ab}, id, approx, memo } / { p:null, why }
  UA.skLook = function (sk, m, q, depth) {
    if (!sk || !sk.tables) return { p: null, why: 'nofile' };
    var rules = (m && m.skRules && m.skRules.length) ? m.skRules : (UA._skDef || (UA._skDef = UA.compileSkRules(skDefaultRules())));
    var spec = nk(q.spec), pg = UA.placeGroup(m, q.place), item = q.item, d = nk(String(q.d || '').replace(/\(1\.6t\)$/, ''));
    var f = thickKey(q.f), sizeN = parseFloat(toHalf(q.size)), firstWhy = null;
    for (var k = 0; k < rules.length; k++) {
      var r = rules[k];
      if (r.items.indexOf('*') < 0 && r.items.indexOf(item) < 0) continue;
      if (r.ds.length && !r.ds.some(function (re) { return re.test(d); })) continue;
      if (r.place !== '*' && r.place !== pg) continue;
      if (!r.pat.test(spec)) continue;
      if (r.min != null && !(sizeN >= r.min)) continue;
      if (r.max != null && !(sizeN <= r.max)) continue;
      var res;
      if (r.id === '配管') {
        if (depth) continue;
        res = UA.skLook(sk, m, { item: 'pipe', d: '', place: q.place, spec: q.spec, size: q.size, f: q.f }, 1);
        if (res.p == null) { firstWhy = firstWhy || res.why; continue; }
      } else {
        res = cellOf(sk, r, q, f);
        if (res.p == null) { firstWhy = firstWhy || res.why; continue; }
      }
      return finish(sk, m, q, r, res, depth);
    }
    return { p: null, why: firstWhy || 'nomatch' };
  };
  function cellOf(sk, r, q, f) {
    var tbl = sk.tables[r.id];
    if (!tbl) return { p: null, why: '表' + r.id + 'が資料にありません' };
    var ri, mat = '', rowT = null;
    if (r.row === '') ri = findRowBySize(tbl, q.size);
    else {
      mat = r.row === '自動' ? matLabel(q.spec) : r.row;
      if (!mat) return { p: null, why: '保温材を読み取れません' };
      ri = findRowByLabel(tbl, mat, f); rowT = ri.t;
    }
    if (ri.i < 0) return { p: null, why: '【' + r.id + '】に' + (mat ? mat + (f ? ' ' + f + 'mm' : '') : 'サイズ' + q.size) + 'の行がありません' };
    var cj, colInfo = null;
    if (r.col === '') {
      cj = findCol(tbl, '', f);
      if (cj.j < 0 && r.row !== '' && tbl.cols.length >= 2 && colDesc(tbl.cols).slice(1).every(function (c) { return c.t == null; })) cj = { j: 1, c: colDesc(tbl.cols)[1] };
    } else cj = findCol(tbl, r.col, f);
    if (cj.j < 0) return { p: null, why: '【' + r.id + '】に' + (r.col ? '「' + r.col + '」' : '') + (f ? f + 'mm' : '') + 'の列がありません' };
    colInfo = cj.c;
    var v = tbl.rows[ri.i][cj.j];
    if (typeof v !== 'number') return { p: null, why: '【' + r.id + '】の' + (tbl.rows[ri.i][0]) + '・' + tbl.cols[cj.j] + 'は空欄です' };
    var res = { p: v, base: v, tbl: tbl, rowLabel: String(tbl.rows[ri.i][0]), col: colInfo, area: ri.area || null, sqm: !!ri.sqm, rowT: rowT };
    if (r.over) {
      var t2 = sk.tables[r.over];
      if (!t2) return { p: null, why: '表' + r.over + 'が資料にありません' };
      var r2 = findRowBySize(t2, q.size), c2 = findCol(t2, '', f);
      var v2 = (r2.i >= 0 && c2.j >= 0) ? t2.rows[r2.i][c2.j] : null;
      if (typeof v2 !== 'number') return { p: null, why: '上被【' + r.over + '】に' + q.size + 'A・' + f + 'mmの単価がありません' };
      res.over = { v: v2, tbl: t2 };
    }
    return res;
  }
  function finish(sk, m, q, r, res, depth) {
    var notes = [], factor = 1, plus = 0;
    // 掛率・加算（対応表の行）
    if (r.factor !== 1) factor *= r.factor;
    if (r.plus) { plus += r.plus; notes.push('＋' + yen(r.plus) + '円'); }
    if (!depth) {
      var adj = (m && m.skAdj && m.skAdj.length) ? m.skAdj : (UA._skAdj || (UA._skAdj = UA.compileSkAdj(UA.skDefaultAdj())));
      var specRaw = nk(String(q.spec).replace(/ヒーター[＋+]/g, 'ヒーター　'));
      adj.forEach(function (a) {
        var tgt = a.target === 'サイズ' ? nk(q.size) : a.target === '種別' ? nk(q.d) : specRaw;
        if (tgt.indexOf(a.word) < 0) return;
        if (a.factor !== 1) { factor *= a.factor; notes.push('×' + a.factor + '（' + a.word + '）'); }
        if (a.plus) { plus += a.plus; notes.push('＋' + yen(a.plus) + '円（' + a.word + '）'); }
      });
      var gf = (m && m.output && m.output.skFactor) || 1;
      if (gf !== 1) { factor *= gf; notes.push('×' + gf + '（掛率）'); }
    }
    var base = res.base, expr, val;
    if (r.id === '配管') {
      // 配管の結果（数式または数値）に掛率をかける
      expr = res.expr || String(res.p); val = res.p;
      notes.unshift('×' + fmtN(r.factor) + '（' + (q.d || '弁') + '・配管の単価）');
    } else {
      expr = String(base); val = base;
      if (res.over) { expr += '+' + res.over.v; val += res.over.v; }
      if (res.area) { expr = '(' + expr + ')*PI()*' + fmtN(res.area / Math.PI); val = val * res.area; }
    }
    var needRound = !!res.area || factor !== 1;
    if (factor !== 1 && r.id !== '配管') expr = '(' + expr + ')*' + fmtN(factor);
    else if (r.id === '配管' && factor !== 1) expr = '(' + expr.replace(/^=/, '') + ')*' + fmtN(factor);
    if (factor !== 1) val = val * factor;
    if (needRound) { expr = 'ROUNDUP(' + expr + ',-1)'; val = ceil10(val); }
    if (plus) { expr += '+' + plus; val += plus; }
    var formula = /[+*()]/.test(expr) ? '=' + expr : null;
    // 参照元の表示
    var src = r.id === '配管' ? res.src : res;
    var tbl = src.tbl, secT = sk.secs[tbl.sec] || '';
    var memo = r.id !== '配管' ? r.memo : (res.memo || '');
    var tFixed = fixedT(tbl, src.col), f = thickKey(q.f), approx = !!memo;
    if (tFixed && f && String(parseFloat(f)) !== tFixed) { notes.push('※表は' + tFixed + 'mm'); approx = true; }
    var showSec = secT && tbl.name.indexOf(secT) !== 0 && secT.indexOf(tbl.name) !== 0;
    var y = '§' + tbl.sec + '【' + tbl.id + '】' + tbl.name + (showSec ? '　' + secT : '');
    if (src.over) y += '＋§' + src.over.tbl.id + ' ' + src.over.tbl.name;
    if (src.sqm) y += ' ㎡単価×π×φ';
    if (notes.length) y += ' ' + notes.join(' ');
    if (memo) y += ' ※' + memo;
    if (sk.label) y += '（' + sk.label + '）';
    var z = (src.col && src.col.fin) || compText(tbl) || tbl.name;
    if (src.col && src.col.fin && src.col.fin.indexOf('×') < 0 && src.col.t == null) {
      for (var jj = tbl.cols.indexOf(src.col.h) - 1; jj >= 1; jj--) { var mh = String(tbl.cols[jj]).match(/^(.*×\s*\d+\s*mm\S*)\s+\S+$/); if (mh) { z = mh[1] + ' ' + src.col.fin; break; } }
    }
    if (src.over) z += '＋' + compText(src.over.tbl);
    if (r.row !== '' && r.id !== '配管') z += '／' + src.rowLabel;
    var item = q.item, aa, ab;
    if (item === 'rect' || item === 'box') aa = '（㎡単価）';
    else if (item === 'round') aa = 'φ' + q.size + (src.sqm ? '（㎡行）' : '');
    else aa = q.size + 'A';
    ab = src.col && src.col.t ? src.col.t + 'mm' : src.rowT ? src.rowT + 'mm' : tFixed ? tFixed + 'mm' : (f ? f + 'mm' : '');
    return { p: val, formula: formula, expr: expr, ref: { y: y, z: z, aa: aa, ab: ab }, id: tbl.id, approx: approx, memo: memo, src: src, rule: r.n };
  }

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
        rec.t = 'elbow'; rec.e = String(E).trim(); rec.g = num(G); rec.h = hT; rec.j = num(J); rec.jf = fcell(r, 'j'); recs.push(rec); continue;
      }
      // 明細：数量あり。バルブ・フランジは数量を空欄で書くことがある（単位が入っていれば明細とみなす）
      if ((!isEmpty(G) && isNumLike(G)) || (isEmpty(G) && !isEmpty(H) && (typeof E === 'number' || isNumLike(E)))) {
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

  // 今の区分（大項目・系統・施工箇所・仕様）の塊にある明細行
  UA.blockLines = function (model, sel) {
    if (!model) return [];
    var W = model.recs, sub = null, i;
    if (sel.noArea && !sel.newArea) {
      var s0 = 0; for (i = 0; i < W.length; i++) if (W[i].t === 'area') s0 = i + 1;
      sub = findSub(W, s0, W.length, sel.place, sel.spec);
    } else {
      if (sel.newArea) return [];
      var ai = findArea(W, sel.area, sel.areaRow);
      if (ai < 0) return [];
      var sr = sysRange(W, ai, areaEnd(W, ai), sel.sys);
      if (!sr) return [];
      sub = findSub(W, sr.s + (sel.sys ? 0 : 1), sr.e, sel.place, sel.spec);
    }
    if (!sub) return [];
    var out = [], k = W.indexOf(sub) + 1;
    while (k < W.length && (W[k].t === 'line' || W[k].t === 'elbow')) { if (W[k].t === 'line') out.push(W[k]); k++; }
    return out;
  };
  // バルブ・フランジの保温厚：配管の保温厚以上で一番近い候補（既定 25・50）。候補より厚ければ配管と同じ
  UA.valveThick = function (pipeT, list) {
    var parts = String(pipeT == null ? '' : pipeT).split(/[+＋]/).map(function (x) { return parseFloat(toHalf(x)); }).filter(function (x) { return isFinite(x); });
    if (!parts.length) return '';
    var tot = parts.reduce(function (a, b) { return a + b; }, 0), L = list && list.length ? list : [25, 50];
    for (var i = 0; i < L.length; i++) if (tot <= L[i]) return String(L[i]);
    return thickKey(pipeT);
  };

  /* ========== 登録計画 ==========
     sel = { area, newArea(bool), sys('' = 系統なし), place, spec, item, lines:[{d,e,f,g,h,i,src}], elbow:金額,
             how: 'add'（今ある行は変えず、同じ塊の明細の下に追加。既定）| 'merge'（同じサイズの行は数量を足す）| 'row'（atRow 行目に入れる） }
     返り値 = { ok, mode, inserts:[{at, recs:[...]}], cellOps:[{row,col,value,prev}], totalOps:[...], final:[...] } */
  UA.plan = function (model, sel, m) {
    var res = { ok: false, mode: '', message: '' };
    var how = sel.how === 'merge' || sel.how === 'row' ? sel.how : 'add';
    res.how = how;
    if (how === 'row') {
      var ar = parseInt(toHalf(String(sel.atRow == null ? '' : sel.atRow)), 10);
      if (!(ar > 0)) { res.message = '入れる行の番号を入れてください'; return res; }
      if (ar < model.startRow) { res.message = model.startRow + '行目より上には入れられません'; return res; }
      if (model.boundary && ar > model.boundary.row) { res.message = '「' + model.boundary.text + '」（' + model.boundary.row + '行目）より下には入れられません'; return res; }
    }
    if (!sel.lines.length && !sel.elbow) { res.message = (sel.item === 'valve' || sel.item === 'flange') ? '種類とサイズを選ぶと、ここに登録先が表示されます' : '数量を入れると、ここに登録先が表示されます'; return res; }
    if (!sel.spec) { res.message = '仕様を選んでください'; return res; }
    if (!sel.area && !sel.noArea && how !== 'row') { res.message = '大項目を選ぶか、「大項目なし」を選んでください'; return res; }
    var base = model.recs.map(function (r) { return Object.assign({}, r); });
    if (how === 'row') { var WR = planAtRow(model, base, sel, ar, res); return WR ? finishPlan(res, model, WR) : res; }
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
    var lines = newLines(sel);
    var ai = sel.newArea ? -1 : findArea(W, sel.area, sel.areaRow);
    var i, p;
    function splice(at, arr) { Array.prototype.splice.apply(W, [at, 0].concat(arr)); }
    function nb(t, o) { return Object.assign({ t: t, isNew: true }, o || {}); }

    if (sel.noArea && !sel.newArea) {
      // 大項目なし：最後の大項目より後ろ（大項目が無いシートは全体）で、同じ施工箇所・仕様の塊を探す。無ければ末尾に作る
      var s0 = 0;
      for (i = 0; i < W.length; i++) if (W[i].t === 'area') s0 = i + 1;
      var e0 = W.length - 1;
      var sub0 = findSub(W, s0, e0, sel.place, sel.spec);
      if (sub0) { insertIntoSub(W, sub0, lines, how === 'merge'); res.mode = 'sub'; }
      else {
        p = e0;
        for (i = s0; i < e0; i++) if (W[i].t === 'total' && W[i].scope === 'grand') { p = i; break; }
        while (p - 1 >= s0 && W[p - 1].t === 'blank') p--;
        var blk0 = [nb('spec', { c: sel.place, d: '', e: sel.spec })].concat(lines);
        if (W.slice(0, p).some(function (r) { return r.t !== 'blank'; })) blk0.unshift(nb('blank'));
        splice(p, blk0);
        res.mode = 'noarea';
      }
    } else if (ai < 0) {
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
      if (sel.totals) {
        block.push(nb('blank'));
        if (sel.sys) { sysTotal = nb('total', { scope: 'sys', label: sel.sys + '　計', name: sel.sys }); block.push(sysTotal); block.push(nb('blank')); }
        block.push(nb('total', { scope: 'area', label: sel.area + '　計', name: sel.area, sysTotalRef: sysTotal }));
      }
      splice(p, block);
      res.mode = 'newarea';
    } else {
      var ae = areaEnd(W, ai);
      var sr = sysRange(W, ai, ae, sel.sys);
      if (sr) {
        var sub = findSub(W, sr.s + (sel.sys ? 0 : 1), sr.e, sel.place, sel.spec);
        if (sub) {
          insertIntoSub(W, sub, lines, how === 'merge');
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
        if (useSysTotal && sel.totals) { sb.push(nb('blank')); sb.push(nb('total', { scope: 'sys', label: sel.sys + '　計', name: sel.sys, newSys: true })); }
        splice(p, sb);
        res.mode = 'newsys';
      }
    }
    return finishPlan(res, model, W);
  };

  // 登録する明細（エルボは「直管の合計×率」の式で J列に入れる）
  function newLines(sel) {
    var lines = sel.lines.map(function (l) { return Object.assign({ t: 'line', isNew: true }, l); });
    if (sel.item === 'pipe' && sel.elbow) lines.push({ t: 'elbow', isNew: true, e: sel.elbowLabel || 'エルボ', g: 1, h: '式', rate: sel.elbowRate || 0.35 });
    return lines;
  }
  function isPiece(r) { return PIECE_UNITS.indexOf(trimAll(r.h)) >= 0; }
  function isRunLine(r) { return r.t === 'line' && !isPiece(r); }   // 直管・ダクトなど（バルブ・フランジ以外）
  UA.isPieceLine = isPiece;

  // 行を指定して入れる（how = 'row'）。大項目・系統はその行の位置のもの。仕様が違えば仕様の行（括弧書き）も入れる
  function planAtRow(model, base, sel, atRow, res) {
    var endRow = model.boundary ? model.boundary.row : Math.max(model.lastContent + 1, model.startRow, atRow);
    var W = base.concat([{ t: 'end', row: endRow }]);
    var idx = W.length - 1, k;
    for (k = 0; k < W.length; k++) if (W[k].row >= atRow) { idx = k; break; }
    if (W[idx].t === 'end' && !model.boundary) W[idx].row = Math.max(W[idx].row, atRow);
    // その位置の文脈（大項目・系統・施工箇所・仕様）
    var cx = { area: '', sys: '', place: '', spec: null };
    // 明細の最後より下に空きを空けて入れるときは、仕様の続きとはみなさない
    var gap = W[idx].t === 'end' && idx > 0 && atRow > W[idx - 1].row + 1;
    for (k = gap ? -1 : idx - 1; k >= 0; k--) {
      var r = W[k], sp = specOfRec(r);
      if (sp) { cx.spec = sp.e; break; }
      if (r.t === 'line' || r.t === 'elbow' || r.t === 'other') continue;
      break;
    }
    var placeFound = false;
    for (k = idx - 1; k >= 0; k--) {
      var r2 = W[k], s2 = specOfRec(r2);
      if (!placeFound && s2 && s2.c) { cx.place = s2.c; placeFound = true; }
      if (r2.t === 'sys' && !cx.sys) { cx.sys = r2.b; }
      if (r2.t === 'area') { cx.area = r2.a; break; }
    }
    res.ctx = cx;
    var lines = newLines(sel);
    var samePlace = trimAll(cx.place) === trimAll(sel.place);
    var needHead = cx.spec !== sel.spec || !samePlace;
    // 同じ仕様の明細の中に入れるとき、その塊にエルボが既にあれば新しく作らない（範囲だけ直す）
    if (!needHead) {
      var hasElb = false;
      for (k = idx; k < W.length && (W[k].t === 'line' || W[k].t === 'elbow'); k++) if (W[k].t === 'elbow') { hasElb = true; W[k].elbowRefresh = true; }
      for (k = idx - 1; k >= 0 && (W[k].t === 'line' || W[k].t === 'elbow'); k--) if (W[k].t === 'elbow') { hasElb = true; W[k].elbowRefresh = true; }
      if (hasElb) lines = lines.filter(function (x) { return x.t !== 'elbow'; });
    }
    var blk = lines.slice();
    if (needHead) {
      var head = { t: 'spec', isNew: true, c: samePlace ? '' : sel.place, d: '', e: sel.spec };
      blk.unshift(head);
      // 別の仕様の明細の途中に入れたときは、後ろの明細のために元の仕様の行を入れ直す
      var nx = W[idx];
      if (cx.spec && nx && (nx.t === 'line' || nx.t === 'elbow')) blk.push({ t: 'spec', isNew: true, c: head.c ? cx.place : '', d: '', e: cx.spec, restore: true });
    }
    res.headAdded = needHead;
    Array.prototype.splice.apply(W, [idx, 0].concat(blk));
    res.mode = 'row';
    res.atRow = atRow;
    return W;
  }

  var ELB_RE = /^=ROUNDUP\(SUM\(\$?([A-Z]+)\$?(\d+):\$?[A-Z]+\$?(\d+)\)\*([\d.]+),-1\)$/i;
  function runAbove(W, idx) {
    var k = idx - 1, out = [];
    while (k >= 0 && isRunLine(W[k])) { out.unshift(W[k]); k--; }
    return out;
  }
  function finishPlan(res, model, W) {
    var i;
    // 直管が上に無い新しいエルボは入れない
    for (i = W.length - 1; i >= 0; i--) if (W[i].t === 'elbow' && W[i].isNew && !runAbove(W, i).length) { W.splice(i, 1); res.elbowSkipped = true; }
    // D列：同じ種別が続くときは最初の行だけ
    var cellOps = [];
    var prevLine = null;
    for (i = 0; i < W.length; i++) {
      var r = W[i];
      if (r.t !== 'line') { prevLine = null; continue; }
      var sameAsPrev = prevLine && prevLine.d === r.d;
      if (r.isNew) r.dShow = sameAsPrev ? '' : (r.d || '');
      // 今ある行の D列を消すのは「数量を足す」のときだけ（ほかの入れ方では今ある行を一切変えない）
      else if (res.how === 'merge' && sameAsPrev && r.dRaw && prevLine.isNew) cellOps.push({ row: r.row, col: 'd', value: '', prev: r.dRaw, why: 'D列の重複を整理' });
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
    // エルボ：J列＝ROUNDUP(SUM(すぐ上の直管のJ列)×率,-1)。既にある式のエルボは、直管を足したら範囲を直す
    res.elbowNote = '';
    W.forEach(function (x, idx) {
      if (x.t !== 'elbow') return;
      var cur = typeof x.jf === 'string' ? x.jf.replace(/\s/g, '') : '';
      var mm = cur.match(ELB_RE);
      if (!x.isNew && !mm) { if (x.elbowRefresh) res.elbowNote = 'エルボ（' + x.finalRow + '行目）は金額の手入力のままです'; return; }
      var run = runAbove(W, idx);
      if (!run.length) return;
      var rate = x.rateNew || x.rate || (mm ? parseFloat(mm[4]) : 0.35);
      var f = '=ROUNDUP(SUM(' + jc + run[0].finalRow + ':' + jc + run[run.length - 1].finalRow + ')*' + fmtN(rate) + ',-1)';
      var sum = 0; run.forEach(function (r) { if (r.i != null && r.g > 0) sum += (r.g + (r.addG || 0)) * r.i; });
      x.jEst = ceil10(sum * rate); x.rateUsed = rate; x.runFrom = run[0].finalRow; x.runTo = run[run.length - 1].finalRow;
      if (x.isNew) { x.formula = f; return; }
      // 行を入れたあと Excel が自動でずらした式と同じなら、書き直さない
      var a0 = +mm[2], b0 = +mm[3];
      var shifted = '=ROUNDUP(SUM(' + jc + (a0 + shiftAt(a0)) + ':' + jc + (b0 + shiftAt(b0)) + ')*' + fmtN(parseFloat(mm[4])) + ',-1)';
      if (shifted.toUpperCase() !== f.toUpperCase()) {
        totalOps.push({ type: 'set', finalRow: x.finalRow, origRow: x.row, formula: f, prev: x.jf, why: 'エルボの範囲' });
        x.elbowSet = true;
      }
    });
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
    var merges = W.filter(function (x) { return !x.isNew && (x.addG || x.addJ || x.elbowSet); });
    var rowsTouched = news.map(function (x) { return x.finalRow; }).concat(merges.map(function (x) { return x.finalRow; }));
    if (!news.length && !merges.length) { res.ok = false; res.message = res.elbowNote || 'エルボは入っています。数量を入れると、ここに登録先が表示されます'; return res; }
    res.ok = true;
    res.inserts = inserts.map(function (g) { return { at: g.at, count: g.recs.length, recs: g.recs }; }).sort(function (a, b) { return b.at - a.at; });
    res.cellOps = cellOps;
    res.totalOps = totalOps;
    res.W = W;
    res.news = news; res.merges = merges;
    res.addCount = news.filter(function (x) { return x.t === 'line' || x.t === 'elbow'; }).length;
    res.mergeCount = merges.filter(function (x) { return x.addG || x.addJ; }).length;
    res.fixCount = merges.filter(function (x) { return x.elbowSet; }).length;
    res.firstRow = rowsTouched.length ? Math.min.apply(null, rowsTouched) : null;
    res.lastRow = rowsTouched.length ? Math.max.apply(null, rowsTouched) : null;
    res.addAmount = 0;
    news.forEach(function (x) { if (x.t === 'line' && x.i != null && x.g > 0) res.addAmount += Math.round(x.g * x.i); if (x.t === 'elbow') res.addAmount += x.jEst || 0; });
    merges.forEach(function (x) { if (x.addG && x.i != null) res.addAmount += Math.round(x.addG * x.i); if (x.elbowSet && x.jEst != null) res.addAmount += x.jEst - (x.j || 0); });
    // 挿入の目印になる行（場所を見る用）
    res.anchorRow = res.inserts.length ? res.inserts[res.inserts.length - 1].at : (merges[0] ? merges[0].row : null);
    return res;
  }

  // 同じ区分・仕様の塊に入れる。merge=true のときだけ、同じ種別・サイズ・厚みの行の数量（エルボは金額）に足す。
  // merge=false（既定）は今ある行を変えず、同じ種別の明細の一番下（エルボはその下）に新しい行として入れる。
  function insertIntoSub(W, specRec, lines, merge) {
    lines.forEach(function (ln) {
      var st = W.indexOf(specRec) + 1, en = st, k;
      while (en < W.length && (W[en].t === 'line' || W[en].t === 'elbow')) en++;
      // エルボ：塊に既にあれば作らず、その式の範囲を直す。無ければ直管の最後の下に入れる
      if (ln.t === 'elbow') {
        for (k = st; k < en; k++) if (W[k].t === 'elbow') { W[k].elbowRefresh = true; if (ln.rate) W[k].rateNew = ln.rate; return; }
        var lastRun = -1;
        for (k = st; k < en; k++) if (isRunLine(W[k])) lastRun = k;
        W.splice(lastRun >= 0 ? lastRun + 1 : en, 0, ln);
        return;
      }
      var hasQty = typeof ln.g === 'number' && ln.g > 0;
      if (merge && hasQty) {
        for (k = st; k < en; k++) {
          var r = W[k];
          if (r.t === 'line' && !r.isNew && r.d === ln.d && String(r.e) === String(ln.e) && String(r.f) === String(ln.f) && (r.h === ln.h || !r.h)) {
            r.addG = (r.addG || 0) + ln.g; return;
          }
        }
      }
      var pos = -1, piece = isPiece(ln);
      if (merge) {
        // 同じ種別の中でサイズ順に入れる
        var same = [];
        for (k = st; k < en; k++) if (W[k].t === 'line' && W[k].d === ln.d && isPiece(W[k]) === piece) same.push(k);
        if (same.length) {
          pos = same[same.length - 1] + 1;
          for (var q = 0; q < same.length; q++) {
            var ev = W[same[q]].e;
            if (typeof ev === 'number' && typeof ln.e === 'number' && ev > ln.e) { pos = same[q]; break; }
          }
        }
      }
      if (pos < 0) {
        // 今ある行は変えず、同じ種別の明細の一番下へ
        for (k = st; k < en; k++) if (W[k].t === 'line' && W[k].d === ln.d && isPiece(W[k]) === piece) pos = k + 1;
      }
      if (pos < 0 && !piece) {
        // 直管・ダクト：ほかの直管の下（エルボ・バルブより上）
        for (k = st; k < en; k++) if (isRunLine(W[k])) pos = k + 1;
        if (pos < 0) { pos = en; for (k = st; k < en; k++) if (W[k].t === 'elbow' || (W[k].t === 'line' && isPiece(W[k]))) { pos = k; break; } }
      }
      if (pos < 0) pos = en;   // バルブ・フランジ：塊の一番下
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

  /* =====================================================================
     v2.0
     - 行ごとの文脈（大項目・系統・室・仕様・種別）を1回の走査で求める（UA.scan）
     - Excel に直接打ったサイズから、品目・保温厚・単位・単価を決める（自動入力）
     - 単価の計算ルール（GV 50A以下＝直管×1.1 など。画面で追加・編集できる）
     - 計・小計・工事費計・合計・出精値引き（見積書の VBA「集計」を移植。不具合は直してある）
     - 経費（資材運搬交通費・諸経費・法定福利費）
     ===================================================================== */
  var C = { A: 0, B: 1, C: 2, D: 3, E: 4, F: 5, G: 6, H: 7, I: 8, J: 9 };
  UA.COL = C;
  function nfkc(s) { var t = String(s == null ? '' : s); try { return t.normalize('NFKC'); } catch (e) { return toHalf(t); } }
  function isBoundA(t) { var a = trimAll(t); return a === '経費' || a === '見積条件' || a.charAt(0) === '一' || a.charAt(0) === '※'; }
  UA.isBoundA = isBoundA;
  // 計の行のラベル（I列）か
  function isTotalLabel(t) { var s = trimAll(t); return !!s && (/計$/.test(s) || s === '出精値引' || s === '出精値引き'); }
  // ラベル・見出しの比較用（空白・括弧・末尾の「計」を除く）
  function headKey(t) { return trimAll(nfkc(t)).replace(/[()（）]/g, '').replace(/計$/, ''); }
  UA.headKey = headKey;

  /* ---------- サイズの書き方 ---------- */
  // 50 / 50A / 20Su / 30VP / φ200 / 200φ / 矩形 / 500×300 / 1000×3300×1500（W×D×H。W・D・H の文字があればそれを優先）
  UA.parseSize = function (e) {
    if (typeof e === 'number') return isFinite(e) && e > 0 ? { n: e, ser: '', text: String(e) } : null;
    var t = nfkc(e).replace(/\s+/g, '');
    if (!t) return null;
    if (t === '矩形') return { rect: true, text: '矩形' };
    var u = t.replace(/[xX＊*✕]/g, '×');
    var parts = u.split('×');
    if (parts.length === 3 || parts.length === 2) {
      var dims = [], ok = true;
      parts.forEach(function (p, i) {
        var mm = p.match(/^([WDHLwdhl]?)(\d+(?:\.\d+)?)([WDHLwdhl]?)$/);
        if (!mm) { ok = false; return; }
        dims.push({ v: +mm[2], k: (mm[1] || mm[3] || '').toUpperCase().replace('L', 'D'), i: i });
      });
      if (ok && parts.length === 3) {
        var o = {}, order = ['W', 'D', 'H'];
        dims.forEach(function (d) { if (d.k && o[d.k] == null) o[d.k] = d.v; else if (d.k) d.k = ''; });
        dims.forEach(function (d) { if (d.k) return; var k = order.filter(function (x) { return o[x] == null; })[0]; if (k) o[k] = d.v; });
        if (o.W > 0 && o.D > 0 && o.H > 0) return { box: { w: o.W, d: o.D, h: o.H }, text: u };
      }
      if (ok && parts.length === 2) return { wh: [dims[0].v, dims[1].v], text: u };
      return null;
    }
    var m = u.match(/^([φΦФ]?)(\d+(?:\.\d+)?)(A|SU|Su|su|VP|vp|Vp|φ|Φ|Ф|mm|MM)?$/);
    if (!m) return null;
    var n = +m[2], s = (m[3] || '').toUpperCase();
    if (!(n > 0)) return null;
    var ser = m[1] || s === 'Φ' || s === 'Ф' || m[3] === 'φ' ? 'φ' : s === 'A' ? 'A' : s === 'SU' ? 'Su' : s === 'VP' ? 'VP' : '';
    return { n: n, ser: ser, text: u };
  };
  // VP（硬質塩ビ管）の呼び → 単価に使う鋼管の呼び径（A）。30VP は 32A、ほかは同じ数字
  UA.VP_TO_A = { 30: 32 };
  // BOX・チャンバーの表面積（VBA「見積単価自動入力」と同じ。W×D×H（D＝奥行・長さ L）で
  //   6面＝2(WH+WD+HD)×1.1、5面＝(WH+2WD+2HD)×1.1（W×H の面が1面だけ）
  UA.boxArea = function (b, five) {
    var w = b.w / 1000, d = b.d / 1000, h = b.h / 1000;
    var a = five ? (w * h + 2 * w * d + 2 * h * d) * 1.1 : 2 * (w * h + w * d + h * d) * 1.1;
    return Math.round(a * 1e4) / 1e4;
  };
  // 5面か（チャンバー・ガラリ・名前の末尾が G）
  UA.isFiveSided = function (name) { var t = nfkc(name).trim(); return /チャンバー|ガラリ/.test(t) || /G$/i.test(t); };

  /* ---------- 品目の区分（配管系／ダクト系） ---------- */
  function famOfCtx(m, x) {
    var a = String(x.area || ''), s = String(x.sys || ''), sp = String(x.spec || '');
    if (/ダクト/.test(a)) return 'duct';
    if (/配管|衛生|給排水|給湯|冷暖房|機器/.test(a)) return 'pipe';
    var sd = m && m.systems ? m.systems.filter(function (y) { return y.name && trimAll(y.name) === trimAll(s); })[0] : null;
    if (sd) return sd.fam;
    if (/ダクト|排煙|給気|排気|外気|還気/.test(s)) return 'duct';
    if (/管$|管\s*\(|水|ドレン|冷媒|蒸気/.test(s)) return 'pipe';
    if (/筒/.test(sp)) return 'pipe';
    if (/ロール|帯|波形|ウェーブ|フェルト|ブランケット|板|鋲/.test(sp)) return 'duct';
    return 'pipe';
  }
  UA.famOfCtx = famOfCtx;
  var PIPE_D = /^(SUS|SGP|VP|VU|HIVP|HTVP|鋼管|白管|黒管|銅管|CU|PE|ポリ管|架橋|ライニング|LP|DVLP|SGP-VA|SGP-VB)/i;
  // 明細の品目を決める（D列の種別・E列のサイズ・文脈から）
  UA.guessItem = function (m, x) {
    var sz = x.sz; if (!sz) return null;
    if (sz.box) return 'box';
    if (sz.rect || sz.wh) return 'rect';
    var d = nk(x.d || ''), subs = (m && m.subs) || {};
    function inSubs(k) { return (subs[k] || []).some(function (s) { return s.id && nk(s.id) === d; }); }
    var duct = famOfCtx(m, x) === 'duct';
    if (sz.ser === 'Su' || sz.ser === 'VP') return 'pipe';
    if (d) {
      if (PIPE_D.test(d)) return 'pipe';
      if (inSubs('flange') || /^(FLG|フランジ)/i.test(d)) return 'flange';
      if (duct || /ダクト|フレックス/.test(d)) {
        if (sz.ser === 'A') return inSubs('valve') ? 'valve' : 'pipe';
        return 'round';
      }
      return 'valve';
    }
    if (sz.ser === 'A') return 'pipe';
    if (sz.ser === 'φ') return 'round';
    return duct ? 'round' : 'pipe';
  };
  function clsOf(item) { return item === 'valve' || item === 'flange' ? 'piece' : item; }
  UA.clsOf = clsOf;

  /* ---------- シートの走査 ---------- */
  // values/formulas：1行目からの2次元配列（A列〜、J列まで以上）。戻り値 { rows:[行番号→情報], lines:[明細], last }
  UA.scan = function (values, formulas, m, startRow) {
    startRow = startRow || 5;
    var out = { rows: [], lines: [], last: values.length, startRow: startRow, bound: null, grand: null };
    var cur = { area: '', areaRow: 0, sys: '', sysRow: 0, sysId: 0, place: '', placeRow: 0, placeSys: '', spec: null, specRow: 0, blockId: 0, d: '', dRow: 0, after: false };
    var blockSeq = 0, sysSeq = 0;
    function g(r, c) { var row = values[r - 1]; var v = row ? row[c] : ''; return v == null ? '' : v; }
    function gf(r) { var row = formulas && formulas[r - 1]; var v = row ? row[C.J] : ''; return typeof v === 'string' ? v : ''; }
    for (var r = 1; r <= values.length; r++) {
      var A = g(r, C.A), B = g(r, C.B), Cc = g(r, C.C), D = g(r, C.D), E = g(r, C.E), F = g(r, C.F), G = g(r, C.G), H = g(r, C.H), I = g(r, C.I), J = g(r, C.J);
      var o = { r: r, A: A, B: B, C: Cc, D: D, E: E, F: F, G: G, H: H, I: I, J: J, jf: gf(r) };
      o.blank = [A, B, Cc, D, E, F, G, H, I].every(isEmpty);
      var aT = typeof A === 'string' ? A.trim() : '';
      if (r < startRow) { o.kind = 'head'; }
      else if (cur.after) { o.kind = o.blank ? 'blank' : (aT ? 'bound' : (isTotalLabel(I) ? 'total' : 'after')); if (aT === '経費') o.exp = true; }
      else if (o.blank) { o.kind = 'blank'; cur.d = ''; cur.dRow = 0; }
      else if (aT && !isNumLike(aT)) {
        if (isBoundA(aT)) { o.kind = 'bound'; cur.after = true; if (!out.bound) out.bound = r; }
        else { o.kind = 'area'; cur.area = aT; cur.areaRow = r; cur.sys = ''; cur.sysRow = 0; cur.placeSys = ''; cur.spec = null; cur.specRow = 0; cur.d = ''; cur.blockId = ++blockSeq; cur.sysId = ++sysSeq; }
      }
      else if (isTotalLabel(I) && isEmpty(G) && isEmpty(H)) {
        o.kind = 'total'; o.label = String(I).trim();
        if (headKey(I) === '工事費' && !out.grand) out.grand = r;
        cur.spec = null; cur.specRow = 0; cur.d = ''; cur.blockId = ++blockSeq;
      }
      else if (!isEmpty(B)) {
        o.kind = 'sys'; cur.sys = String(B).trim(); cur.sysRow = r; cur.placeSys = ''; cur.spec = null; cur.specRow = 0; cur.d = ''; cur.blockId = ++blockSeq; cur.sysId = ++sysSeq;
        if (isParen(Cc)) { cur.place = trimAll(stripParen(Cc)); cur.placeRow = r; cur.placeSys = cur.place; }
        if (isParen(E)) { cur.spec = normSpec(E); cur.specRow = r; }
      }
      else if (isParen(Cc) || isParen(E)) {
        o.kind = 'spec';
        if (isParen(Cc)) { cur.place = trimAll(stripParen(Cc)); cur.placeRow = r; cur.placeSys = cur.place; }
        if (isParen(E)) { cur.spec = normSpec(E); cur.specRow = r; }
        cur.d = isEmpty(D) ? '' : String(D).trim(); cur.dRow = cur.d ? r : 0; cur.blockId = ++blockSeq;
      }
      else {
        var sz = UA.parseSize(E);
        if (sz) {
          o.kind = 'line'; o.sz = sz;
          if (!isEmpty(D)) { cur.d = String(D).trim(); cur.dRow = r; }
        } else if (!isEmpty(D) && isEmpty(E) && isEmpty(G)) {
          o.kind = 'dhead'; cur.d = String(D).trim(); cur.dRow = r;
        } else o.kind = 'other';
      }
      o.area = cur.area; o.areaRow = cur.areaRow; o.sys = cur.sys; o.sysRow = cur.sysRow; o.sysId = cur.sysId; o.place = cur.place; o.placeSys = cur.placeSys;
      o.spec = cur.spec; o.specRow = cur.specRow; o.blockId = cur.blockId; o.d = cur.d; o.after = cur.after && o.kind !== 'bound' ? true : cur.after;
      if (o.kind === 'line') {
        o.item = UA.guessItem(m, { sz: o.sz, d: o.d, area: o.area, sys: o.sys, spec: o.spec });
        o.cls = clsOf(o.item);
        out.lines.push(o);
      }
      out.rows[r] = o;
    }
    return out;
  };
  // 行 r に書いたときの文脈（r の行の D 列は d で渡す。r が表の外でも使える）
  UA.ctxAt = function (sc, r, d) {
    var prev = null;
    for (var k = Math.min(r - 1, sc.rows.length - 1); k >= 1; k--) { if (sc.rows[k]) { prev = sc.rows[k]; break; } }
    var me = sc.rows[r];
    var base = prev || { area: '', sys: '', place: '', placeSys: '', spec: null, specRow: 0, blockId: 0, sysId: 0, d: '' };
    var x = { row: r, area: base.area, sys: base.sys, sysId: base.sysId, place: base.place, placeSys: base.placeSys, spec: base.spec, specRow: base.specRow, blockId: base.blockId, after: !!base.after || (prev && prev.kind === 'bound') };
    // 1行上が空行なら、上の明細の種別は引き継がない
    var dUp = prev && prev.kind !== 'blank' ? base.d : '';
    x.d = !isEmpty(d) ? String(d).trim() : dUp;
    x.prevKind = prev ? prev.kind : '';
    if (me && me.kind === 'spec') x.isSpecRow = true;
    return x;
  };

  /* ---------- 保温厚・単位の推定 ---------- */
  function bestOf(cnt) { var b = null, n = 0; Object.keys(cnt).forEach(function (k) { if (cnt[k] > n) { n = cnt[k]; b = k; } }); return b; }
  function fStr(v) { return isEmpty(v) ? '' : thickKey(v); }
  // 近い明細（上を優先し、無ければ下）
  function nearest(L, r, pred) {
    var up = null, dn = null;
    for (var i = 0; i < L.length; i++) { var l = L[i]; if (l.r === r || !pred(l)) continue; if (l.r < r) up = l; else if (!dn) dn = l; }
    return up || dn;
  }
  function szKey(sz) { return sz ? (sz.n != null ? sizeKey(sz.n) : sizeKey(sz.text)) : ''; }
  function isDuctCls(c) { return c === 'rect' || c === 'box'; }
  // 戻り値 { f, from }
  UA.guessThick = function (sc, m, x, item, sz, r) {
    var cls = clsOf(item), sk = szKey(sz), L = sc.lines, l;
    var kind = KIND_OF_ITEM[item] || item, dk = nk(x.d || '');
    function hasF(q) { return !!fStr(q.F); }
    function sameCls(q) { return isDuctCls(cls) ? isDuctCls(q.cls) : q.cls === cls; }
    function sameSize(q) { return szKey(q.sz) === sk; }
    // 同じ塊の明細に保温厚が書かれていない（鉛・撤去など）ときは入れない
    var near0 = nearest(L, r, function (q) { return q.blockId === x.blockId && (!isEmpty(q.G) || !isEmpty(q.I)); });
    if ((near0 && !hasF(near0)) || /鉛|撤去|分別/.test(x.spec || '')) return { f: '', from: '保温厚を書かない仕様' };
    if (isDuctCls(cls)) {
      // 角ダクト・BOX：同じ仕様の塊で同じ種別 → すぐ上（下）の明細 → 見積書の同じ仕様
      l = nearest(L, r, function (q) { return q.blockId === x.blockId && sameCls(q) && hasF(q) && dk && nk(q.d) === dk; });
      if (l) return { f: fStr(l.F), from: '同じ種別（' + l.r + '行目）' };
      l = nearest(L, r, function (q) { return q.blockId === x.blockId && sameCls(q) && hasF(q); });
      if (l) return { f: fStr(l.F), from: '同じ仕様の明細（' + l.r + '行目）' };
      if (x.spec) { l = nearest(L, r, function (q) { return q.spec === x.spec && sameCls(q) && hasF(q); }); if (l) return { f: fStr(l.F), from: 'この見積書の同じ仕様（' + l.r + '行目）' }; }
    } else {
      // 1) 同じ塊（同じ仕様の行の下）の同じ種別・同じサイズ
      l = nearest(L, r, function (q) { return q.blockId === x.blockId && sameCls(q) && hasF(q) && sameSize(q) && (cls !== 'piece' || nk(q.d) === dk); });
      if (l) return { f: fStr(l.F), from: '同じ仕様の' + sizeLabel(item, sz) + '（' + l.r + '行目）' };
      if (cls === 'piece') {
        // 2) バルブ・フランジ：同じ塊の配管（直管）の同じサイズ → 同じ厚み
        l = nearest(L, r, function (q) { return q.blockId === x.blockId && q.cls === 'pipe' && hasF(q) && sameSize(q); });
        if (l) return { f: fStr(l.F), from: '同じ仕様の配管 ' + sk + 'A（' + l.r + '行目）と同じ' };
        //    別の仕様（ロールなど）の塊：同じ系統の配管の同じサイズから 25以下→25、50以下→50
        l = nearest(L, r, function (q) { return x.sysId && q.sysId === x.sysId && q.cls === 'pipe' && hasF(q) && sameSize(q); });
        if (l) {
          var vt = UA.valveThick(fStr(l.F), m && m.output && m.output.valveThick);
          return { f: vt, from: '配管 ' + sk + 'A の ' + fStr(l.F) + 'mm（' + l.r + '行目）から' };
        }
        l = nearest(L, r, function (q) { return q.blockId === x.blockId && q.cls === 'piece' && hasF(q); });
        if (l && x.spec && !/筒/.test(x.spec)) return { f: fStr(l.F), from: '同じ仕様の' + (l.d || 'バルブ') + '（' + l.r + '行目）と同じ' };
      }
      // 3) この見積書の同じ仕様・同じサイズ（いちばん多い厚み）
      if (x.spec) {
        var cnt = {}, at = {};
        L.forEach(function (q) { if (q.r !== r && q.spec === x.spec && sameCls(q) && hasF(q) && sameSize(q)) { var f = fStr(q.F); cnt[f] = (cnt[f] || 0) + 1; at[f] = at[f] || q.r; } });
        var b = bestOf(cnt);
        if (b) return { f: b, from: 'この見積書の同じ仕様・' + sizeLabel(item, sz) + '（' + at[b] + '行目など）' };
        // 4) 同じ仕様で、前後のサイズの厚みが同じならその厚み
        if (sz.n != null) {
          var lo = null, hi = null;
          L.forEach(function (q) {
            if (q.r === r || q.spec !== x.spec || !sameCls(q) || !hasF(q) || q.sz.n == null) return;
            if (q.sz.n < sz.n && (!lo || q.sz.n > lo.sz.n)) lo = q;
            if (q.sz.n > sz.n && (!hi || q.sz.n < hi.sz.n)) hi = q;
          });
          if (lo && hi && fStr(lo.F) === fStr(hi.F)) return { f: fStr(lo.F), from: '同じ仕様の' + sizeLabel(item, lo.sz) + '・' + sizeLabel(item, hi.sz) + 'と同じ' };
        }
      }
    }
    // 5) 設定マスタの標準厚（サイズ別 → 全サイズ）
    var t1 = m && m.thickIdx ? m.thickIdx[kind + '|' + (x.spec || '') + '|' + sk] : null;
    if (t1) return { f: thickKey(t1), from: '標準厚（M_標準厚・' + sizeLabel(item, sz) + '）' };
    var t2 = m && m.thickIdx ? m.thickIdx[kind + '|' + (x.spec || '') + '|（全サイズ）'] : null;
    if (t2) return { f: thickKey(t2), from: '標準厚（M_標準厚）' };
    // 6) 同じ塊のすぐ上の明細 → 既定
    l = nearest(L, r, function (q) { return q.blockId === x.blockId && sameCls(q) && hasF(q); });
    if (l) return { f: fStr(l.F), from: '近くの行（' + l.r + '行目）と同じ' };
    return { f: DEFAULT_THICK[kind] || '25', from: '既定の厚み' };
  };
  function sizeLabel(item, sz) {
    if (!sz) return '';
    if (sz.box || sz.wh || sz.rect) return sz.text;
    if (sz.ser === 'Su') return sz.n + 'Su';
    if (sz.ser === 'VP') return sz.n + 'VP';
    if (item === 'round' || sz.ser === 'φ') return 'φ' + sz.n;
    return sz.n + 'A';
  }
  UA.sizeLabel = sizeLabel;
  var PIECE_D = /継手|エルボ|チーズ|T管|フレックス|フレキ/;
  UA.guessUnit = function (sc, m, x, item, r) {
    var cls = clsOf(item), L = sc.lines, l, dk = nk(x.d || '');
    var meP = cls === 'piece' || (!!dk && PIECE_D.test(dk));
    function hasH(q) { return !isEmpty(q.H); }
    function pz(q) { return q.cls === 'piece' || (!!q.d && PIECE_D.test(nk(q.d))) || PIECE_UNITS.indexOf(String(q.H).trim()) >= 0; }
    function sameD(q) { return hasH(q) && q.cls === cls && !!dk && nk(q.d) === dk; }
    function sameK(q) { return hasH(q) && q.cls === cls && pz(q) === meP; }
    function inB(q) { return q.blockId === x.blockId; }
    function inS(q) { return !!x.sysId && q.sysId === x.sysId; }
    l = nearest(L, r, function (q) { return inB(q) && sameD(q); }) || nearest(L, r, function (q) { return inB(q) && sameK(q); }) ||
      nearest(L, r, function (q) { return inS(q) && sameD(q); }) || nearest(L, r, function (q) { return inS(q) && sameK(q); }) ||
      nearest(L, r, sameD);
    if (l) return String(l.H).trim();
    var cnt = {}; L.forEach(function (q) { if (q.r !== r && sameK(q)) { var h = String(q.H).trim(); cnt[h] = (cnt[h] || 0) + 1; } });
    var b = bestOf(cnt); if (b) return b;
    if (item === 'box') return '台';
    if (meP) { var pc = {}; L.forEach(function (q) { var h = hasH(q) ? String(q.H).trim() : ''; if (PIECE_UNITS.indexOf(h) >= 0) pc[h] = (pc[h] || 0) + 1; }); return bestOf(pc) || 'ヶ'; }
    var k = m && m.kindById ? m.kindById[item] : null;
    var u = k && k.unit ? k.unit : (item === 'rect' ? '㎡' : 'm');
    return UA.normalizeOut(m, 'H', u);
  };

  /* ---------- 単価の計算ルール ---------- */
  var ROUND_LABEL = { up10: '1の位切り上げ', up100: '10の位切り上げ', up1: '1円未満切り上げ', none: '端数そのまま' };
  var ROUND_HINT = { up10: '10円単位に切り上げ（例 4,741 → 4,750）', up100: '100円単位に切り上げ（例 4,741 → 4,800）', up1: '1円単位に切り上げ（例 4,741.2 → 4,742）', none: '計算したまま（切り上げない）' };
  UA.ROUND_LABEL = ROUND_LABEL; UA.ROUND_HINT = ROUND_HINT;
  function roundKey(t) { var s = trimAll(t); if (!s || /1の位|10円/.test(s)) return 'up10'; if (/10の位|100円/.test(s)) return 'up100'; if (/1円|小数/.test(s)) return 'up1'; if (/なし|そのまま/.test(s)) return 'none'; return 'up10'; }
  UA.roundKey = roundKey;
  UA.defaultCalcRules = function () {
    return [
      { id: 'flex', on: true, name: 'フレックス継手（補修）', items: ['round'], words: 'フレックス継手、フレキ継手、継手補修', min: '', max: '', base: 'run', baseKind: '', factor: 0.8, round: 'up10', memo: '既存マクロ「継手補修×0.8」と同じ' },
      { id: 'gv50', on: true, name: 'GV（50A以下）', items: ['valve'], words: 'GV', min: '', max: '50', base: 'run', baseKind: '', factor: 1.1, round: 'up10', memo: '' },
      { id: 'pump', on: false, name: 'ラインポンプ', items: ['valve'], words: 'ラインポンプ', min: '', max: '', base: 'kind', baseKind: 'GV', factor: 1.5, round: 'up10', memo: '既存マクロにある掛率（確認してから使ってください）' },
      { id: 'ystr', on: false, name: 'Yスト', items: ['valve'], words: 'Yスト', min: '', max: '', base: 'kind', baseKind: 'GV', factor: 1.2, round: 'up10', memo: '既存マクロにある掛率（確認してから使ってください）' }
    ];
  };
  function wordsOf(w) { return String(w == null ? '' : w).split(/[、,，\s]+/).map(function (x) { return nk(x); }).filter(Boolean); }
  UA.wordsOf = wordsOf;
  UA.matchCalcRule = function (rules, q) {
    if (!rules) return null;
    var d = nk(q.d || ''), n = q.sz && q.sz.n;
    for (var i = 0; i < rules.length; i++) {
      var r = rules[i];
      if (!r.on) continue;
      var its = r.items && r.items.length ? r.items : ['*'];
      if (its.indexOf('*') < 0 && its.indexOf(q.item) < 0) continue;
      var ws = wordsOf(r.words);
      if (ws.length && !ws.some(function (w) { return d.indexOf(w) >= 0; })) continue;
      if (!isEmpty(r.min) && !(n >= num(r.min))) continue;
      if (!isEmpty(r.max) && !(n <= num(r.max))) continue;
      if (r.base === 'kind' && r.baseKind && nk(r.baseKind) === d) continue;   // 自分自身を基準にしない
      return r;
    }
    return null;
  };
  function roundExpr(expr, mode) { return mode === 'up10' ? 'ROUNDUP(' + expr + ',-1)' : mode === 'up100' ? 'ROUNDUP(' + expr + ',-2)' : mode === 'up1' ? 'ROUNDUP(' + expr + ',0)' : expr; }
  function roundVal(v, mode) {
    function up(x, k) { return Math.ceil(x / k - 1e-9) * k; }
    return mode === 'up10' ? up(v, 10) : mode === 'up100' ? up(v, 100) : mode === 'up1' ? up(v, 1) : Math.round(v * 100) / 100;
  }
  UA.roundVal = roundVal;
  function itemLabel(m, id) { var k = m && m.kindById && m.kindById[id]; return k ? k.label : ({ pipe: '配管', valve: 'バルブ類', flange: 'フランジ類', rect: '角ダクト', round: '丸ダクト', box: 'BOX' })[id] || id; }
  UA.itemLabel = itemLabel;
  UA.describeRule = function (r, m) {
    var its = (r.items || []).filter(function (x) { return x && x !== '*'; });
    var who = its.length ? its.map(function (x) { return itemLabel(m, x); }).join('・') : 'すべての品目';
    var ws = String(r.words || '').split(/[、,，\s]+/).filter(Boolean);
    var cond = [];
    if (ws.length) cond.push('D列に「' + ws.join('」「') + '」');
    if (!isEmpty(r.min) && !isEmpty(r.max)) cond.push(num(r.min) + '〜' + num(r.max));
    else if (!isEmpty(r.min)) cond.push(num(r.min) + '以上');
    else if (!isEmpty(r.max)) cond.push(num(r.max) + '以下');
    var base = r.base === 'kind' && r.baseKind ? '同じ仕様の「' + r.baseKind + '」' : '同じ仕様の直管';
    return { when: who + (cond.length ? '（' + cond.join('・') + '）' : ''), then: base + ' × ' + fmtN(num(r.factor) || 1), round: ROUND_LABEL[r.round] || ROUND_LABEL.up10 };
  };

  /* ---------- 明細1行の単価 ----------
     q = { item, d, place, spec, sz, f }
     help.findBase({item, d, sz, f}) → { row, i(数値) } シート上の基準の行（直管など）
     戻り値 { p, formula, ref:{y,z,aa,ab}, via:'rule'|'sk', approx } / { p:null, why, ref } */
  UA.priceLine = function (sk, m, rules, q, help, depth) {
    depth = depth || 0;
    var sz = q.sz;
    if (!sz) return { p: null, why: 'サイズを読み取れません' };
    var rule = depth < 3 ? UA.matchCalcRule(rules, q) : null;
    if (rule) return ruleLine(sk, m, rules, q, help, depth, rule);
    if (!sk) return noPrice('保温積算資料が未読込');
    var look = sz.n, conv = '', size;
    if (q.item === 'rect') size = '矩形';
    else if (q.item === 'box') size = sz.box ? sz.text : 'BOX';   // 寸法で書いた BOX は「BOX＋1,000円」を足さない（VBA と同じ）
    else {
      if (sz.ser === 'Su' || (q.item === 'pipe' && /^SUS$/i.test(String(q.d || '').trim()))) {
        var su = (m && m.sizes && m.sizes.Su || []).filter(function (z) { return z.n === sz.n; })[0];
        var ps = su && su.priceSize != null ? su.priceSize : (SU_DEFAULT.filter(function (x) { return x[0] === sz.n; })[0] || [])[1];
        if (ps == null) return noPrice(sz.n + 'Su に当たる鋼管のサイズがありません（M_サイズ の「単価に使うサイズ」）');
        look = ps; conv = sz.n + 'Su（' + ps + 'A相当）';
      } else if (sz.ser === 'VP') { look = UA.VP_TO_A[sz.n] || sz.n; conv = sz.n + 'VP（' + look + 'A相当）'; }
      size = look;
    }
    var dd = /^SUS$/i.test(String(q.d || '').trim()) ? '' : q.d;
    var r = UA.skLook(sk, m, { item: q.item, d: dd, place: q.place, spec: q.spec, size: size, f: q.f });
    if (r.p == null) return noPrice(r.why && r.why !== 'nomatch' ? r.why : '対応する表がありません（M_積算資料対応）');
    var ref = Object.assign({}, r.ref), formula = r.formula, p = r.p;
    if (conv) ref.aa = conv;
    if (q.item === 'box' && sz.box) {
      var five = q.five != null ? !!q.five : UA.isFiveSided(q.d || '');
      var area = UA.boxArea(sz.box, five);
      var ex = formula ? '(' + formula.slice(1) + ')' : String(p);
      formula = '=ROUNDUP(' + ex + '*' + area + ',-1)';
      p = roundVal(p * area, 'up10');
      ref.y += ' ×面積' + area + '㎡（' + (five ? '5面' : '6面') + '×1.1）';
      ref.aa = sz.text + '（' + area + '㎡・' + (five ? '5面' : '6面') + '）';
    } else if (q.item === 'rect' && sz.wh) ref.aa = '（㎡単価）' + sz.text;
    return { p: p, formula: formula, ref: ref, via: 'sk', approx: r.approx, id: r.id };
    function noPrice(why) { return { p: null, why: why, ref: { y: '単価なし：' + why, z: '', aa: sizeLabel(q.item, sz), ab: q.f ? q.f + 'mm' : '' } }; }
  };
  function ruleLine(sk, m, rules, q, help, depth, rule) {
    var baseItem = rule.base === 'kind' ? q.item : (q.item === 'round' || q.item === 'rect' || q.item === 'box' ? 'round' : 'pipe');
    var baseD = rule.base === 'kind' ? rule.baseKind : '';
    var baseName = rule.base === 'kind' ? rule.baseKind : '直管';
    var fac = num(rule.factor) || 1, expr, val, from, z = '';
    var hit = help && help.findBase ? help.findBase({ item: baseItem, d: baseD, sz: q.sz, f: q.f }) : null;
    if (hit && hit.i != null) {
      expr = 'I' + hit.row; val = hit.i; from = baseName + '＝' + hit.row + '行目の単価';
    } else {
      var bq = Object.assign({}, q, { item: baseItem, d: baseD });
      var bp = UA.priceLine(sk, m, rules, bq, help, depth + 1);
      if (bp.p == null) {
        var why = '計算ルール「' + rule.name + '」の基準（' + baseName + ' ' + sizeLabel(baseItem, q.sz) + (q.f ? ' ' + q.f + 'mm' : '') + '）の単価がありません';
        return { p: null, why: why, ref: { y: '単価なし：' + why + (bp.why ? '（' + bp.why + '）' : ''), z: '', aa: sizeLabel(q.item, q.sz), ab: q.f ? q.f + 'mm' : '' } };
      }
      expr = bp.formula ? '(' + bp.formula.slice(1) + ')' : String(bp.p); val = bp.p;
      from = baseName + '＝' + (bp.ref && bp.ref.y ? bp.ref.y.replace(/^単価なし：/, '') : '') + '（' + yen(bp.p) + '円）';
      z = bp.ref ? bp.ref.z : '';
    }
    var formula = '=' + roundExpr(expr + '*' + fmtN(fac), rule.round);
    var p = roundVal(val * fac, rule.round);
    var y = '計算：' + rule.name + '　' + baseName + '×' + fmtN(fac) + '（' + (ROUND_LABEL[rule.round] || '') + '）｜' + from;
    return { p: p, formula: formula, ref: { y: y, z: z || ('同じ仕様の' + baseName), aa: sizeLabel(q.item, q.sz), ab: q.f ? q.f + 'mm' : '' }, via: 'rule', rule: rule };
  }
  // 計算ルールの基準になる行（同じ塊の直管・指定の種別で、同じサイズ。同じ厚みを優先）
  UA.findBaseLine = function (sc, x, r, b) {
    var sk = szKey(b.sz), dk = nk(b.d || ''), best = null;
    sc.lines.forEach(function (q) {
      if (q.r === r || q.blockId !== x.blockId || szKey(q.sz) !== sk || typeof q.I !== 'number') return;
      if (b.d ? nk(q.d) !== dk : (q.item !== b.item || (q.d && PIECE_D.test(nk(q.d))) || PIECE_UNITS.indexOf(String(q.H).trim()) >= 0)) return;
      var sc2 = (fStr(q.F) === fStr(b.f) ? 2 : 0) + (q.r < r ? 1 : 0);
      if (!best || sc2 > best.s) best = { s: sc2, row: q.r, i: q.I };
    });
    return best;
  };
  // 自動入力が書いた単価か（Y列の書き方で見分ける）
  UA.isAutoRef = function (y) { return typeof y === 'string' && /^(§|単価なし：|計算：)/.test(y.trim()); };

  /* ---------- 計・小計・工事費計・合計・出精値引き（VBA「集計」の移植） ----------
     g = { v: values, f: formulas }（1行目から、A〜J列）。R＝選んでいる行。
     kind：'range'（A/B/C列 集計＝見出しの下から範囲の合計）／'sub'（A/B列 小計＝計の行の合計）
           'grand'（工事費 計）／'total'（合計）／'discount'（出精値引き）
     戻り値 { ok, message, warn, at（書き始めの行）, k（使う行数）, ins（入れる行数）, insAt, writes:[{r,col,v}], diag:[行], label, formula, select } */
  var TOP = 4;
  UA.totalPlan = function (g, R, kind, opt) {
    opt = opt || {};
    var col = opt.col || 'A', ci = C[col];
    function v(r, c) { var row = g.v[r - 1]; var x = row ? row[c] : ''; return x == null ? '' : x; }
    function f(r) { var row = g.f && g.f[r - 1]; var x = row ? row[C.J] : ''; return typeof x === 'string' ? x.replace(/\s/g, '') : ''; }
    function blank(r) { for (var c = 0; c <= C.I; c++) if (!isEmpty(v(r, c))) return false; return true; }
    function txt(r, c) { var x = v(r, c); return typeof x === 'string' && !isEmpty(x) && !isNumLike(x) ? x : ''; }
    function isSum(r) { return /^=SUM\(/i.test(f(r)); }
    function label(r) { return trimAll(v(r, C.I)); }
    function disp(t) {
      var s = String(t).replace(/[()（）]/g, '').trim();
      if (opt.full === false) { var mm = s.match(/^(.*?)[.．]/); if (mm) s = mm[1]; }
      return s;
    }
    var res = { ok: false, kind: kind };
    if (!(R >= TOP + 1)) { res.message = (TOP + 1) + '行目より下の行を選んでください'; return res; }
    if (kind === 'range' || kind === 'sub' || kind === 'grand') {
      // 経費・見積条件より下では使わない
      for (var ib = R; ib >= TOP; ib--) { var ab = txt(ib, C.A); if (ab && isBoundA(ab)) { res.message = '「' + ab.trim().slice(0, 8) + '」より下では使えません'; return res; } if (headKey(v(ib, C.I)) === '工事費' && kind !== 'grand') { res.message = '工事費計より下では使えません'; return res; } }
    }
    // 書き込む位置：選んだ行が空き行ならそこから、明細などが入っていればその下から
    var need = kind === 'discount' ? 3 : 2;
    var P = blank(R) ? R : R + 1, b = 0;
    while (b < need && P + b <= g.v.length + need && blank(P + b)) b++;
    res.at = P; res.k = need; res.ins = need - b; res.insAt = P + b;
    var Dg = P, T = P + 1, jc = 'J';
    var writes = [], refs = [], lab = '', fm = '';
    if (kind === 'range') {
      var cnt3 = 0, head = null, stop = null;
      for (var i = R; i >= TOP; i--) {
        var t = txt(i, ci);
        if (!t) { if (isSum(i)) { stop = i; break; } cnt3 = i; }
        else { head = i; break; }
      }
      if (head === R) { res.message = R + '行目は見出しの行です。明細の下の空き行を選んでください'; return res; }
      if (!cnt3) { res.message = '合計する明細が見つかりません'; return res; }
      lab = (head ? disp(v(head, ci)) : '') + '　計';
      fm = '=SUM(' + jc + cnt3 + ':' + jc + Dg + ')';
      res.from = cnt3; res.head = head;
      if (!head) res.warn = col + '列の見出しが見つからないため、' + (stop ? stop + '行目の計' : TOP + '行目') + 'の下から合計します';
    } else if (kind === 'sub') {
      var s1 = [], s2 = [], head2 = null;
      for (var i2 = R; i2 >= TOP; i2--) {
        var t2 = txt(i2, ci);
        if (!t2) { var fx = f(i2); if (/^=SUM\(.*,.*\)$/i.test(fx)) s1.push(i2); else if (/^=SUM\(/i.test(fx)) s2.push(i2); }
        else { head2 = i2; break; }
      }
      refs = (s1.length ? s1 : s2).slice().sort(function (a, b2) { return a - b2; });
      if (!refs.length) { res.message = (head2 ? '「' + String(v(head2, ci)).trim() + '」の下に' : '') + '計の行（=SUM の式）が見つかりません'; return res; }
      lab = (head2 ? disp(v(head2, ci)) : '') + '　計';
      fm = '=SUM(' + refs.map(function (x) { return jc + x; }).join(',') + ')';
      res.head = head2; res.refs = refs;
      if (!head2) res.warn = col + '列の見出しが見つかりません';
    } else if (kind === 'grand') {
      var heads = [], curH = null, missing = [];
      for (var i3 = TOP; i3 <= R; i3++) {
        var a = txt(i3, C.A);
        if (a) {
          if (isBoundA(a)) break;
          curH = { r: i3, key: headKey(a), num: (nfkc(a).match(/^\s*(\d+)\s*[.．]/) || [])[1] || null, text: a.trim(), tot: null };
          heads.push(curH); continue;
        }
        var lb = v(i3, C.I);
        if (curH && !curH.tot && isTotalLabel(lb) && isSum(i3)) {
          var k = headKey(lb);
          if (k === curH.key || (curH.num && k === curH.num)) curH.tot = i3;
        }
      }
      heads.forEach(function (h) { if (h.tot) refs.push(h.tot); else missing.push(h.text); });
      if (!heads.length) {
        // 大項目が無い見積：系統の計をまとめる。系統の計も無ければ範囲で合計
        var sysT = [], curB = null;
        for (var i4 = TOP; i4 <= R; i4++) {
          var bt = txt(i4, C.B); if (bt) { curB = headKey(bt); continue; }
          var lb2 = v(i4, C.I);
          if (curB && isTotalLabel(lb2) && isSum(i4) && headKey(lb2) === curB) { sysT.push(i4); curB = null; }
        }
        if (sysT.length) { refs = sysT; res.warn = '大項目が無いため、系統の計を合計します'; }
        else { fm = '=SUM(' + jc + (TOP + 1) + ':' + jc + Dg + ')'; res.warn = '大項目・系統の計が無いため、明細全体を合計します'; }
      } else if (!refs.length) { res.message = '大項目の計（例「1.配管設備　計」）が見つかりません。先に A列の集計・小計を入れてください'; return res; }
      if (missing.length && heads.length) res.warn = '「' + missing.join('」「') + '」の計が無いため、工事費計に入っていません';
      if (!fm) fm = '=SUM(' + refs.map(function (x) { return jc + x; }).join(',') + ')';
      lab = '工事費　計'; res.refs = refs;
    } else if (kind === 'total') {
      var w = null;
      for (var i5 = R; i5 >= TOP; i5--) if (headKey(v(i5, C.I)) === '工事費' && isSum(i5)) { w = i5; break; }
      if (!w) for (var i6 = R; i6 >= TOP; i6--) if (isSum(i6)) { w = i6; break; }
      if (!w) { res.message = '工事費計（=SUM の式）が見つかりません'; return res; }
      lab = '合計'; fm = '=SUM(' + jc + w + ':' + jc + Dg + ')'; res.from = w;
    } else if (kind === 'discount') {
      var amt = num(opt.amount);
      if (!(amt > 0)) { res.message = '決定金額を入れてください'; return res; }
      var gr = null;
      for (var i7 = R; i7 >= TOP; i7--) if (label(i7) === '合計') { gr = i7; break; }
      if (!gr) { res.message = '「合計」の行が見つかりません。先に合計を入れてください'; return res; }
      writes.push({ r: P, col: 'I', v: '出精値引' }, { r: P, col: 'J', v: '=SUM(' + Math.round(amt) + '-' + jc + gr + ')' });
      Dg = P + 1; T = P + 2;
      lab = '総合計'; fm = '=SUM(' + jc + gr + ':' + jc + P + ')'; res.from = gr; res.amount = amt;
    }
    writes.push({ r: Dg, col: 'J', v: '', diag: true });
    writes.push({ r: T, col: 'I', v: lab }, { r: T, col: 'J', v: fm });
    res.ok = true; res.writes = writes; res.diag = [Dg]; res.label = lab; res.formula = fm; res.totalRow = T; res.select = T + 1;
    return res;
  };

  /* ---------- 経費 ---------- */
  UA.EXPENSE_DEFAULT = {
    transport: { on: true, label: '資材運搬交通費', k: 0.05, l: -3 },
    misc: { on: true, label: '諸経費', k: 0.12, l: -3 },
    welfare: { on: true, label: '法定福利費', k: 0.55, l: 0.15188 }
  };
  var EXP_KEYS = ['transport', 'misc', 'welfare'];
  UA.EXP_KEYS = EXP_KEYS;
  function expMatch(key, c) { var t = trimAll(c); return key === 'transport' ? /運搬|交通費|駐車場/.test(t) : key === 'misc' ? /諸経費/.test(t) : /法定福利/.test(t); }
  // 材工の見積か（A列に「資材費」「労務費」の大項目がある）
  UA.expenseInfo = function (g) {
    function v(r, c) { var row = g.v[r - 1]; var x = row ? row[c] : ''; return x == null ? '' : x; }
    function f(r) { var row = g.f && g.f[r - 1]; var x = row ? row[C.J] : ''; return typeof x === 'string' ? x.replace(/\s/g, '') : ''; }
    var info = { grand: null, zaiko: false, labor: null, material: null, laborHead: null, materialHead: null, exp: null, items: {}, total: null };
    var curH = null;
    for (var r = TOP; r <= g.v.length; r++) {
      var a = v(r, C.A), at = typeof a === 'string' ? a.trim() : '';
      if (at && !isNumLike(at)) {
        if (trimAll(at) === '経費') { info.exp = r; curH = null; continue; }
        if (isBoundA(at)) { if (info.exp) break; curH = null; continue; }
        curH = { r: r, key: headKey(at), num: (nfkc(at).match(/^\s*(\d+)\s*[.．]/) || [])[1] || null };
        if (/労務費/.test(at)) { info.zaiko = true; info.laborHead = curH; }
        if (/資材費/.test(at)) { info.zaiko = true; info.materialHead = curH; }
        continue;
      }
      var lb = v(r, C.I);
      if (isTotalLabel(lb) && /^=SUM\(/i.test(f(r))) {
        var k = headKey(lb);
        if (k === '工事費' && !info.grand) info.grand = r;
        if (curH && (k === curH.key || (curH.num && k === curH.num))) {
          if (info.laborHead === curH && !info.labor) info.labor = r;
          if (info.materialHead === curH && !info.material) info.material = r;
        }
      }
      if (trimAll(lb) === '合計' && !info.total) info.total = r;
      if (info.exp && r > info.exp && !info.total) {
        var c = v(r, C.C);
        EXP_KEYS.forEach(function (key) { if (!info.items[key] && !isEmpty(c) && expMatch(key, c)) info.items[key] = r; });
      }
    }
    return info;
  };
  // cfg = { transport:{on,label,k,l}, misc:{...}, welfare:{...} }, mode:'kyokai'|'zaiko', base: 工事金額の行（省略時は工事費計）, labor: 労務費の式
  UA.expensePlan = function (g, cfg, opt) {
    opt = opt || {};
    var info = UA.expenseInfo(g), res = { ok: false, info: info };
    var W = opt.base || info.grand;
    if (!W) { res.message = '「工事費　計」の行が見つかりません。先に工事費計を入れてください'; return res; }
    function v(r, c) { var row = g.v[r - 1]; var x = row ? row[c] : ''; return x == null ? '' : x; }
    // K〜M列（手計算・メモ）に何か入っている行も使わない（上書きしないように）
    function blankRow(r) { for (var c = 0; c <= 12; c++) if (c !== C.J && !isEmpty(v(r, c))) return false; return true; }
    var zaiko = opt.mode ? opt.mode === 'zaiko' : info.zaiko;
    var laborExpr = null, laborNote = '';
    if (zaiko) {
      if (opt.labor) { laborExpr = opt.labor; laborNote = '労務費＝' + opt.labor; }
      else if (info.labor) { laborExpr = 'J' + info.labor; laborNote = '労務費＝J' + info.labor + '（労務費の計）'; }
      else if (info.material) { laborExpr = '(J' + W + '-J' + info.material + ')'; laborNote = '労務費＝工事金額−資材費（J' + W + '−J' + info.material + '）'; }
      else { laborExpr = 'J' + W; laborNote = '労務費の計が見つからないため、工事金額（J' + W + '）を使います'; res.warn = laborNote; }
    }
    var keys = EXP_KEYS.filter(function (k) { return cfg[k] && cfg[k].on; });
    if (!keys.length) { res.message = '入れる経費の項目がありません'; return res; }
    var rows = {}, ins = [], P = null, head = null;
    if (info.exp && info.exp > W) {
      // 「経費」の見出しがある：同じ項目の行は書き直し、無い項目は最後の項目の下に入れる
      head = info.exp;
      var last = head;
      EXP_KEYS.forEach(function (k) { if (info.items[k]) last = Math.max(last, info.items[k]); });
      var add = keys.filter(function (k) { return !info.items[k]; });
      keys.forEach(function (k) { if (info.items[k]) rows[k] = info.items[k]; });
      if (add.length) {
        var free = 0; while (free < add.length && blankRow(last + 1 + free) && v(last + 1 + free, C.I) === '') free++;
        // 空き行は使うが、合計の前の空き行1行は残す
        var useFree = Math.max(0, Math.min(free, add.length) - (free >= add.length + 1 ? 0 : 1));
        if (free >= add.length + 1) useFree = add.length;
        var insN = add.length - useFree;
        if (insN > 0) ins.push({ at: last + 1 + useFree, n: insN });
        add.forEach(function (k, i) { rows[k] = last + 1 + i; });
      }
      res.mode = 'update';
    } else {
      // 工事費計の下に「空き行・経費・項目・空き行」を作る
      P = W + 1;
      var needN = 2 + keys.length + 1, b = 0;
      while (b < needN && blankRow(P + b)) b++;
      if (b < needN) ins.push({ at: P + b, n: needN - b });
      head = P + 1;
      keys.forEach(function (k, i) { rows[k] = head + 1 + i; });
      res.mode = 'new';
    }
    var writes = [];
    if (res.mode === 'new') writes.push({ r: head, col: 'A', v: '経費' });
    keys.forEach(function (k) {
      var r = rows[k], c = cfg[k], isNew = res.mode === 'new' || !info.items[k];
      if (isNew) writes.push({ r: r, col: 'C', v: c.label });
      writes.push({ r: r, col: 'G', v: 1 }, { r: r, col: 'H', v: '式' });
      if (k === 'welfare') {
        if (zaiko) {
          writes.push({ r: r, col: 'K', v: '' }, { r: r, col: 'L', v: num(c.l) }, { r: r, col: 'M', v: '=ROUNDUP(' + laborExpr + '*L' + r + ',-1)' });
        } else {
          writes.push({ r: r, col: 'K', v: num(c.k) }, { r: r, col: 'L', v: num(c.l) }, { r: r, col: 'M', v: '=ROUNDUP(J' + W + '*K' + r + '*L' + r + ',-1)' });
        }
      } else {
        writes.push({ r: r, col: 'K', v: num(c.k) }, { r: r, col: 'L', v: num(c.l) }, { r: r, col: 'M', v: '=ROUNDUP(J' + W + '*K' + r + ',L' + r + ')' });
      }
      writes.push({ r: r, col: 'J', v: '=M' + r });
    });
    res.ok = true; res.base = W; res.zaiko = zaiko; res.laborNote = laborNote; res.ins = ins; res.writes = writes; res.rows = rows; res.head = head;
    res.select = Math.max.apply(null, keys.map(function (k) { return rows[k]; })) + 1;
    return res;
  };
  // 経費の金額の見込み（画面の表示用）
  UA.expenseEstimate = function (k, c, base, labor, zaiko) {
    if (k === 'welfare') return zaiko ? roundVal(labor * num(c.l), 'up10') : roundVal(base * num(c.k) * num(c.l), 'up10');
    var d = -num(c.l), p = Math.pow(10, d > 0 ? d : 0);
    return Math.ceil(base * num(c.k) / p - 1e-9) * p;
  };

  /* ---------- 仕様の行を書き込む計画（入力フォーム） ----------
     sel = { place, spec, cont（上の仕様の続きなら true を許す）, lines:[{d,e,f,g,h,i,iF,ref}]（角ダクト・BOX）, d（バルブの種類） }
     戻り値 { ok, P, k, ins, insAt, specRow, inputRow, writeSpec, writePlace, rows:[{r, kind, ...}], ctx } */
  UA.specWritePlan = function (sc, R, sel) {
    var res = { ok: false };
    if (!(R >= sc.startRow)) { res.message = sc.startRow + '行目より下の行を選んでください'; return res; }
    var P = R;
    var x = UA.ctxAt(sc, P);
    if (x.after) { res.message = '「工事費　計」・経費・見積条件より上の行を選んでください'; return res; }
    var prev = sc.rows[P - 1];
    var inBlock = prev && (prev.kind === 'line' || prev.kind === 'spec' || prev.kind === 'dhead' || prev.kind === 'other') && prev.blockId === x.blockId;
    var same = x.spec && sel.spec && x.spec === normSpec(sel.spec) && trimAll(x.place) === trimAll(sel.place);
    var writeSpec = !(same && inBlock);
    var writePlace = writeSpec && !(x.placeSys && trimAll(x.placeSys) === trimAll(sel.place));
    var nLines = sel.lines && sel.lines.length ? sel.lines.length : 1;
    var k = (writeSpec ? 1 : 0) + nLines;
    // 書く行（少なくとも選んだ行と次の行）にデータがあれば、選んだ行に行を挿入して書く
    var busy = [];
    for (var i = 0; i < Math.max(2, k); i++) { var o = sc.rows[P + i]; if (o && !o.blank) busy.push(P + i); }
    var need = busy.length > 0;
    var n = need ? Math.max(k, Math.floor(num(sel.insN)) || k) : 0;
    res.ok = true; res.P = P; res.k = k; res.ins = n; res.insAt = P; res.needIns = need; res.busy = busy; res.minIns = k;
    res.ctx = x; res.same = same; res.inBlock = !!inBlock;
    res.writeSpec = writeSpec; res.writePlace = writePlace;
    res.specRow = writeSpec ? P : null;
    res.inputRow = P + (writeSpec ? 1 : 0);
    res.select = sel.lines && sel.lines.length ? res.inputRow + nLines : res.inputRow;
    return res;
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = UA;
  else root.UA = UA;
})(typeof window !== 'undefined' ? window : this);
