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
 *   計画_月    … 1 行 = 1 か月（変動費の入れ方・全体予算・作成/更新日時・手取りの実績・上振れの貯金割合・生活費を引くか）
 *   計画_設定  … カテゴリなど（キーと値）
 *
 * アプリとのやり取り
 *   GET  ?action=planget&token=…            → { ok:true, data:{ schemaVersion, updatedAt, categories, plans } }（未保存なら data:null）
 *   POST { action:'planset', token, data }  → 3 シートを data の内容で書き直す
 */

var PLAN_SHEETS = { items: '計画', months: '計画_月', meta: '計画_設定' };
var PLAN_KINDS = ['income', 'living', 'fixed', 'savings', 'variable'];
var PLAN_KIND_LABEL = { income: '収入', living: '生活費', fixed: '固定費', savings: '貯金', variable: '変動費' };
var PLAN_ITEM_HEAD = ['月', '区分', '有効', '項目', '種別', '金額', '支払日', 'メモ', 'ID', '更新日時'];
var PLAN_MONTH_HEAD = ['月', '変動費の入れ方', '変動費（全体）', '作成日時', '更新日時', '手取り（実績）', '上振れの貯金割合（%）', '生活費'];
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
  var catKeys = { income: '収入カテゴリ', variable: '変動費カテゴリ', savingsType: '貯金・投資の種別' };
  Object.keys(catKeys).forEach(function (k) { var a = parseList(meta[catKeys[k]]); if (a) data.categories[k] = a; });

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
  });
  var kindOf = {};
  PLAN_KINDS.forEach(function (k) { kindOf[PLAN_KIND_LABEL[k]] = k; });
  itemRows.forEach(function (r) {
    var m = planMonth_(r[0]);
    var kind = kindOf[String(r[1])];
    var p = data.plans[m];
    if (!p || !kind) return;
    var it = {
      id: planStr_(r[8]),
      enabled: !(r[2] === false || String(r[2]).toUpperCase() === 'FALSE' || r[2] === '×'),
      amount: kind === 'living' ? planSigned_(r[5]) : planAmount_(r[5]),
      memo: planStr_(r[7]),
      updatedAt: planStr_(r[9])
    };
    if (kind === 'income' || kind === 'variable') it.category = planStr_(r[3]);
    if (kind === 'fixed') { it.name = planStr_(r[3]); it.payDay = planStr_(r[6]); }
    if (kind === 'living') it.name = planStr_(r[3]);
    if (kind === 'savings') { it.name = planStr_(r[3]); it.type = planStr_(r[4]); }
    p[kind].push(it);
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
        p.livingEnabled === false ? 'オフ' : 'オン']);
      PLAN_KINDS.forEach(function (k) {
        (p[k] || []).forEach(function (it) {
          var label = (k === 'fixed' || k === 'savings' || k === 'living') ? (it.name || '') : (it.category || '');
          itemRows.push([planText_(m), PLAN_KIND_LABEL[k], it.enabled !== false, planText_(label), k === 'savings' ? planText_(it.type) : '',
            k === 'living' ? planSigned_(it.amount) : planAmount_(it.amount), k === 'fixed' ? planText_(it.payDay) : '', planText_(it.memo), planText_(it.id), planText_(it.updatedAt)]);
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
      ['貯金・投資の種別', planText_(JSON.stringify(cats.savingsType || []))]
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
