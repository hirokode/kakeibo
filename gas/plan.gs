/**
 * 家計簿アプリの「計画」タブ（旧お金計画）のデータを、このスプレッドシートに保存する。
 *
 * 使い方（くわしくは gas/README.md）
 *   1. このファイルを、家計簿の Apps Script プロジェクトに「plan.gs」として追加する
 *   2. api.gs の doGet / doPost の先頭に 1 行ずつ足す
 *        function doGet(e)  { var r = planRouteGet_(e);  if (r) return r;  …既存の処理… }
 *        function doPost(e) { var r = planRoutePost_(e); if (r) return r;  …既存の処理… }
 *   3. デプロイ → デプロイを管理 → 編集 → バージョン「新バージョン」で更新（URL は変わらない）
 *
 * 保存先のシート（なければ自動で作る。人が読んでも分かる形にしてある）
 *   計画       … 1 行 = 1 項目（月・区分・有効・項目・種別・金額・支払日・メモ・ID・更新日時）。生活費の金額はマイナス（差し引く項目）も入る
 *   計画_月    … 1 行 = 1 か月（変動費の入れ方・全体予算・作成/更新日時・手取りの実績・上振れの貯金割合・生活費を引くか・毎月の貯金）
 *   計画_設定  … カテゴリなど（キーと値）
 *   どのシートも、列のない項目（アプリに後から足した項目）は最後の「その他」に JSON で入る。
 *   なので、アプリのデータに項目を足すだけなら、このファイルを入れ替えなくてよい
 *
 * アプリとのやり取り
 *   GET  ?action=planget&token=…            → { ok:true, data:{ schemaVersion, updatedAt, categories, plans } }（未保存なら data:null）
 *   POST { action:'planset', token, data }  → 3 シートを data の内容で書き直す
 */

var PLAN_SHEETS = { items: '計画', months: '計画_月', meta: '計画_設定' };
var PLAN_KINDS = ['income', 'living', 'fixed', 'savings', 'variable'];
var PLAN_KIND_LABEL = { income: '収入', living: '生活費', fixed: '固定費', savings: '貯金', variable: '変動費' };
var PLAN_ITEM_HEAD = ['月', '区分', '有効', '項目', '種別', '金額', '支払日', 'メモ', 'ID', '更新日時', 'その他'];
var PLAN_MONTH_HEAD = ['月', '変動費の入れ方', '変動費（全体）', '作成日時', '更新日時', '手取り（実績）', '上振れの貯金割合（%）', '生活費', '毎月の貯金', 'その他'];
// 列がある項目（これ以外は「その他」に JSON で入れる）
var PLAN_MONTH_KEYS = ['month', 'variableMode', 'variableTotal', 'createdAt', 'updatedAt', 'actualIncome', 'surplusSaveRate', 'livingEnabled', 'savingsBase'];
var PLAN_TOP_KEYS = ['schemaVersion', 'updatedAt', 'categoriesUpdatedAt', 'categories', 'plans'];
var PLAN_CAT_KEYS = { income: '収入カテゴリ', variable: '変動費カテゴリ', savingsType: '貯金・投資の種別' };
/** 区分ごとに「項目」「種別」「支払日」の列へ入れるもの */
function planItemCols_(kind) {
  if (kind === 'income' || kind === 'variable') return { label: 'category' };
  if (kind === 'fixed') return { label: 'name', payDay: 'payDay' };
  if (kind === 'savings') return { label: 'name', type: 'type' };
  return { label: 'name', type: 'type', payDay: 'payDay' }; // 生活費と、後から足した区分
}
/** 列にない項目だけを JSON にする（なければ空） */
function planExtra_(obj, known) {
  var x = {}, n = 0;
  Object.keys(obj || {}).forEach(function (k) { if (known.indexOf(k) < 0 && obj[k] !== undefined) { x[k] = obj[k]; n++; } });
  return n ? planText_(JSON.stringify(x)) : '';
}
function planParseExtra_(v, into) {
  var t = String(v == null ? '' : v);
  if (!t) return;
  try { var x = JSON.parse(t); if (x && typeof x === 'object') Object.keys(x).forEach(function (k) { if (into[k] === undefined) into[k] = x[k]; }); } catch (e) { /* 壊れた JSON は読まない */ }
}
var PLAN_META_HEAD = ['キー', '値'];

/** doGet の先頭で呼ぶ。計画以外のリクエストなら null を返す */
function planRouteGet_(e) {
  var p = (e && e.parameter) || {};
  if (p.action !== 'planget') return null;
  return planRespond_(p.token, function () { return planGet_(); });
}

/** doPost の先頭で呼ぶ。計画以外のリクエストなら null を返す */
function planRoutePost_(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return null; }
  if (!body || body.action !== 'planset') return null;
  return planRespond_(body.token, function () { return planSet_(body.data); });
}

function planRespond_(token, fn) {
  var out;
  try {
    if (!planTokenOk_(token)) throw new Error('合言葉が違います');
    out = { ok: true, data: fn() };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** 合言葉は家計簿と同じもの（api.gs の apiToken_）を使う。名前が違う場合はここを直す */
function planTokenOk_(token) {
  var expected = '';
  try { if (typeof apiToken_ === 'function') expected = String(apiToken_()); } catch (e1) { /* 次を試す */ }
  try { if (!expected && typeof apiToken_ !== 'undefined' && typeof apiToken_ !== 'function') expected = String(apiToken_); } catch (e2) { /* 次を試す */ }
  if (!expected) {
    var props = PropertiesService.getScriptProperties();
    expected = props.getProperty('apiToken') || props.getProperty('API_TOKEN') || props.getProperty('token') || '';
  }
  if (!expected) throw new Error('GAS 側の合言葉が見つかりません（plan.gs の planTokenOk_ を確認）');
  return String(token || '') === expected;
}

function planSheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function planRows_(sh, width) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, width).getValues();
}

/** シートの日付・数値を文字列に戻す（シートが勝手に日付へ変えることがあるため） */
function planStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'UTC', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  return v == null ? '' : String(v);
}
function planMonth_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  return String(v || '').replace(/^'/, '');
}
/** 文字列として保存する（先頭に ' を付けると、日付・数式に変換されずそのまま入る） */
function planText_(v) {
  var t = v == null ? '' : String(v);
  return t ? "'" + t : '';
}
function planAmount_(v) {
  var n = Number(String(v == null ? '' : v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) && n > 0 ? Math.round(n) : 0;
}
/** 生活費の金額（マイナス可） */
function planSigned_(v) {
  var n = Number(String(v == null ? '' : v).replace(/−/g, '-').replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? Math.round(n) : 0;
}

function planGet_() {
  var meta = {};
  planRows_(planSheet_(PLAN_SHEETS.meta, PLAN_META_HEAD), 2).forEach(function (r) { if (r[0]) meta[String(r[0])] = r[1]; });
  var monthRows = planRows_(planSheet_(PLAN_SHEETS.months, PLAN_MONTH_HEAD), PLAN_MONTH_HEAD.length);
  var itemRows = planRows_(planSheet_(PLAN_SHEETS.items, PLAN_ITEM_HEAD), PLAN_ITEM_HEAD.length);
  if (!monthRows.length && !itemRows.length && !meta.updatedAt) return null;

  var parseList = function (v) { try { var a = JSON.parse(String(v || '')); return Array.isArray(a) ? a : undefined; } catch (e) { return undefined; } };
  var data = {
    schemaVersion: Number(meta.schemaVersion) || 3,
    updatedAt: planStr_(meta.updatedAt),
    categoriesUpdatedAt: planStr_(meta.categoriesUpdatedAt),
    categories: {},
    plans: {}
  };
  Object.keys(PLAN_CAT_KEYS).forEach(function (k) { var a = parseList(meta[PLAN_CAT_KEYS[k]]); if (a) data.categories[k] = a; });
  var top = {};
  planParseExtra_(meta['その他'], top);
  if (top.categories && typeof top.categories === 'object') { planParseExtra_(JSON.stringify(top.categories), data.categories); delete top.categories; }
  planParseExtra_(JSON.stringify(top), data);

  monthRows.forEach(function (r) {
    var m = planMonth_(r[0]);
    if (!/^\d{4}-\d{2}$/.test(m)) return;
    data.plans[m] = {
      month: m,
      variableMode: r[1] === '全体' ? 'total' : 'items',
      variableTotal: planAmount_(r[2]),
      createdAt: planStr_(r[3]), updatedAt: planStr_(r[4]),
      income: [], living: [], fixed: [], savings: [], variable: []
    };
    // 手取りの実績と割合は、空欄なら送らない（アプリ側で端末の値か既定の 70% を使う）
    if (r[5] !== '' && r[5] != null) data.plans[m].actualIncome = planAmount_(r[5]);
    if (r[6] !== '' && r[6] != null && isFinite(Number(r[6]))) data.plans[m].surplusSaveRate = Number(r[6]);
    data.plans[m].livingEnabled = String(r[7]) !== 'オフ';
    if (r[8] !== '' && r[8] != null) data.plans[m].savingsBase = planAmount_(r[8]);
    planParseExtra_(r[9], data.plans[m]);
  });
  var kindOf = {};
  PLAN_KINDS.forEach(function (k) { kindOf[PLAN_KIND_LABEL[k]] = k; });
  itemRows.forEach(function (r) {
    var m = planMonth_(r[0]);
    var label = String(r[1] || '');
    var kind = kindOf[label] || label; // 知らない区分は、その名前のまま
    var p = data.plans[m];
    if (!p || !kind) return;
    var cols = planItemCols_(kind);
    var it = {
      id: planStr_(r[8]),
      enabled: !(r[2] === false || String(r[2]).toUpperCase() === 'FALSE' || r[2] === '×'),
      amount: kindOf[label] && kind !== 'living' ? planAmount_(r[5]) : planSigned_(r[5]),
      memo: planStr_(r[7]),
      updatedAt: planStr_(r[9])
    };
    it[cols.label] = planStr_(r[3]);
    // 生活費や後から足した区分では、空の「種別」「支払日」は読まない
    var loose = kind !== 'savings' && kind !== 'fixed';
    if (cols.type && !(loose && r[4] === '')) it[cols.type] = planStr_(r[4]);
    if (cols.payDay && !(loose && r[6] === '')) it[cols.payDay] = planStr_(r[6]);
    planParseExtra_(r[10], it);
    (p[kind] = p[kind] || []).push(it);
  });
  return data;
}

function planSet_(data) {
  if (!data || typeof data !== 'object' || !data.plans) throw new Error('計画のデータがありません');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var months = Object.keys(data.plans).sort();
    var monthRows = [], itemRows = [];
    months.forEach(function (m) {
      var p = data.plans[m] || {};
      // 月は先頭に ' を付けて文字列のまま保存する（日付に変換されないように）
      var rate = Number(p.surplusSaveRate);
      monthRows.push([planText_(m), p.variableMode === 'total' ? '全体' : '項目別', planAmount_(p.variableTotal), planText_(p.createdAt), planText_(p.updatedAt),
        planAmount_(p.actualIncome) || '', isFinite(rate) && p.surplusSaveRate !== '' && p.surplusSaveRate != null ? rate : '',
        p.livingEnabled === false ? 'オフ' : 'オン',
        p.savingsBase === undefined || p.savingsBase === null || p.savingsBase === '' ? '' : planAmount_(p.savingsBase),
        planExtra_(p, PLAN_MONTH_KEYS.concat(Object.keys(p).filter(function (k) { return Array.isArray(p[k]); })))]);
      // 区分は決まった順、そのあとに知らない区分（配列になっているもの）
      var kinds = PLAN_KINDS.concat(Object.keys(p).filter(function (k) { return Array.isArray(p[k]) && PLAN_KINDS.indexOf(k) < 0; }));
      kinds.forEach(function (k) {
        var cols = planItemCols_(k);
        var known = ['id', 'enabled', 'amount', 'memo', 'updatedAt', cols.label];
        if (cols.type) known.push(cols.type);
        if (cols.payDay) known.push(cols.payDay);
        (p[k] || []).forEach(function (it) {
          itemRows.push([planText_(m), PLAN_KIND_LABEL[k] || k, it.enabled !== false, planText_(it[cols.label]), cols.type ? planText_(it[cols.type]) : '',
            PLAN_KIND_LABEL[k] && k !== 'living' ? planAmount_(it.amount) : planSigned_(it.amount), cols.payDay ? planText_(it[cols.payDay]) : '',
            planText_(it.memo), planText_(it.id), planText_(it.updatedAt), planExtra_(it, known)]);
        });
      });
    });
    planWrite_(planSheet_(PLAN_SHEETS.items, PLAN_ITEM_HEAD), PLAN_ITEM_HEAD, itemRows);
    planWrite_(planSheet_(PLAN_SHEETS.months, PLAN_MONTH_HEAD), PLAN_MONTH_HEAD, monthRows);
    var cats = data.categories || {};
    planWrite_(planSheet_(PLAN_SHEETS.meta, PLAN_META_HEAD), PLAN_META_HEAD, [
      ['schemaVersion', Number(data.schemaVersion) || 3],
      ['updatedAt', planText_(data.updatedAt)],
      ['categoriesUpdatedAt', planText_(data.categoriesUpdatedAt)],
      ['収入カテゴリ', planText_(JSON.stringify(cats.income || []))],
      ['変動費カテゴリ', planText_(JSON.stringify(cats.variable || []))],
      ['貯金・投資の種別', planText_(JSON.stringify(cats.savingsType || []))],
      ['その他', (function () {
        var top = {}, n = 0;
        Object.keys(data).forEach(function (k) { if (PLAN_TOP_KEYS.indexOf(k) < 0) { top[k] = data[k]; n++; } });
        var extraCats = {}, c = 0;
        Object.keys(cats).forEach(function (k) { if (!PLAN_CAT_KEYS[k]) { extraCats[k] = cats[k]; c++; } });
        if (c) { top.categories = extraCats; n++; }
        return n ? planText_(JSON.stringify(top)) : '';
      })()]
    ]);
    return { months: months.length, items: itemRows.length, updatedAt: data.updatedAt || '' };
  } finally {
    lock.releaseLock();
  }
}

function planWrite_(sh, head, rows) {
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, Math.max(head.length, sh.getLastColumn())).clearContent();
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
}
