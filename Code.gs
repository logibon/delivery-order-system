/** 配送受注システム - Code.gs
 *  データの流れ: CSV取込 → 【予定データ】 → 締め日時に自動転記 → 【確定データ】 → 指示書印刷・ステータス入力・請求
 *  スプレッドシートの「拡張機能 > Apps Script」に貼り付けて使用(プロジェクトのタイムゾーンは Asia/Tokyo にしてください) */

/* ===== 設定 ===== */
const FORECAST_SS_ID = '16vmBfTANC7DKbzDe8YW-olR33CdqsJlmTOI-Rdl_lBA';   // 予定データ(依頼予定データ)
const FORECAST_SHEET = '依頼予定';
const CONFIRMED_SS_ID = '1H04N77fog2w4p8Rk_JKxpBPAWtNPg2_qYViT88V7KME';  // 確定データ(依頼予定確定データ)
const CONFIRMED_SHEET = '確定データ';
const CUTOFF_DAYS_BEFORE = 2;           // 締め日時: 配送日の何日前の
const CUTOFF_HOUR = 0;                  //          何時か(現在は「配送日の2日前の0:00」)

const USERS_SS_ID = '1n2q3mFtJJfJhOWvwYF8BQnRqW4LeL_mEx0FvBH16XTs';                 // ログイン情報(「ユーザー」シート)のスプレッドシートID(空欄=このスプレッドシート)
const PRICE_SS_ID = '1UExQK0ufrCQ1ECQpVsn7tHnaAYPZgdyMMEQ0cdUBMuM';                 // 料金表(「料金設定」シート)のスプレッドシートID(空欄=このスプレッドシート)
const BILLING_FOLDER_ID = '1D8_mRG0HmZ5-Bqk-IPsgOJSP_xwExoTw';            // 請求明細(スプレッドシート)の保存先 Googleドライブのフォルダ ID(空欄=マイドライブ直下)
const ISSUER = { name: 'ロジボンサービス株式会社', address: '', tel: '03-5856-1100', bank: '' }; // 請求書の発行者・振込先(空欄の項目は表示しません)

const PRICE = '料金設定', USERS = 'アカウント', DONE = '配送完了';
const BILLABLE = ['配送完了', '不在'];   // 請求の対象ステータス(配送完了=1件、不在=1件。再配送は別の行として残る)
const H = ['配送ステータス','配送完了時間','配送ID','配送日','時間カテゴリ','指定時間枠','配送カテゴリ','顧客名','住所','連絡先','品名','商品3辺(cm)','商品㎥数','サイズ','重量','組立設置','引取','個口数','備考','商品ID','取込日時','完了メモ'];


/* ===== シート操作(列名で読み書きするので、列の順序が違っても動きます) ===== */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function book_(id) { return id ? SpreadsheetApp.openById(id) : ss_(); }

function sheetOf_(id, name, extra) {
  const bk = book_(id); let sh = bk.getSheetByName(name);
  if (!sh) {
    const hd = H.concat(extra || []);
    sh = bk.insertSheet(name);
    sh.getRange(1, 1, 1, hd.length).setValues([hd]).setFontWeight('bold').setBackground('#e8eefc');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, sh.getMaxRows() - 1, hd.length).setNumberFormat('@'); // 文字列固定
  }
  return sh;
}
function fcSheet_() { return sheetOf_(FORECAST_SS_ID, FORECAST_SHEET, ['転記日時']); }
function cfSheet_() { return sheetOf_(CONFIRMED_SS_ID, CONFIRMED_SHEET); }

function hdr_(sh) {
  const n = sh.getLastColumn();
  return n ? sh.getRange(1, 1, 1, n).getValues()[0].map(h => String(h).replace(/[\s\u3000]/g, '').replace(/^配送ステータス.*/, '配送ステータス')) : [];
}
function colOf_(sh, name) { // 列番号(なければ末尾に列を追加)
  const hd = hdr_(sh); let i = hd.indexOf(name);
  if (i < 0) { i = hd.length; sh.getRange(1, i + 1).setValue(name).setFontWeight('bold'); }
  return i + 1;
}
function ensureCols_(sh, names) { names.forEach(n => colOf_(sh, n)); } // 無い列は末尾に追加
function readRows_(sh) {
  const last = sh.getLastRow(), hd = hdr_(sh);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, hd.length).getValues().map((r, i) => {
    const o = { _row: i + 2 };
    hd.forEach((h, c) => { if (h) o[h] = h === '配送日' ? normDate_(r[c]) : str_(r[c]); });
    return o;
  });
}
function readAll_() { return readRows_(cfSheet_()); }   // 確定データ
function key_(r) { return r['配送ID'] + '|' + r['商品ID'] + '|' + r['配送日']; }   // 再配送(同じ配送ID・別の配送日)は重複にせず別の行として残す
function rowsOf_(id, date) { // 配送ID(+配送日)の行。日付なしは、最新の配送日
  let rows = readAll_().filter(r => r['配送ID'] === String(id));
  if (date) rows = rows.filter(r => r['配送日'] === date);
  else if (rows.length) { const d = rows.map(r => r['配送日']).sort().pop(); rows = rows.filter(r => r['配送日'] === d); }
  return rows;
}
function appendRows_(sh, objs) {
  if (!objs.length) return;
  const hd = hdr_(sh), start = sh.getLastRow() + 1;
  sh.getRange(start, 1, objs.length, hd.length).setNumberFormat('@')
    .setValues(objs.map(o => hd.map(h => o[h] == null ? '' : o[h])));
}
function setCells_(sh, row, map) { // {列名: 値}
  Object.keys(map).forEach(k => sh.getRange(row, colOf_(sh, k)).setNumberFormat('@').setValue(map[k]));
}

/* ===== 認証・共通 ===== */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('index');
  t.initId = (e && e.parameter && e.parameter.id) || '';
  t.initD = (e && e.parameter && e.parameter.d) || '';   // 配送日(再配送の区別)
  t.initK = (e && e.parameter && e.parameter.k) || '';   // 報告用キー(指示書QR)   // 指示書のQR(?id=配送ID)から開いた場合
  return t.evaluate().setTitle('配送受注システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function usersSheet_() { // アカウント管理シート: 会社名 / 権限 / メールアドレス / パスワード (「有効」「氏名」列があれば併用)
  const bk = book_(USERS_SS_ID);
  let sh = bk.getSheetByName(USERS);
  if (!sh) {
    sh = bk.insertSheet(USERS);
    sh.getRange('A:D').setNumberFormat('@'); // パスワードの先頭ゼロ消え防止
    sh.getRange(1, 1, 1, 4).setValues([['会社名', '権限', 'メールアドレス', 'パスワード']]).setFontWeight('bold').setBackground('#e8eefc');
  } // 権限: 「管理者」以外はすべて「ユーザー」扱い
  return sh;
}

function login(email, pw) {
  email = String(email || '').trim().toLowerCase(); pw = String(pw || '').trim();
  const sh = usersSheet_(), hd = hdr_(sh), ci = n => hd.indexOf(n);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, hd.length).getValues() : [];
  const u = email && pw && rows.find(r =>
    String(r[ci('メールアドレス')]).trim().toLowerCase() === email &&
    String(r[ci('パスワード')]).trim() === pw &&
    (ci('有効') < 0 || String(r[ci('有効')]).trim() !== '0'));
  if (!u) { Utilities.sleep(800); return null; }
  const name = String(ci('氏名') >= 0 && u[ci('氏名')] ? u[ci('氏名')] : u[ci('会社名')]);
  const role = String(u[ci('権限')]).trim() === '管理者' ? '管理者' : 'ユーザー';
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('s:' + token, JSON.stringify({ email: email, name: name, role: role }), 21600); // 6時間
  return { token: token, name: name, role: role };
}

function auth_(token, adminOnly) {
  const v = token && CacheService.getScriptCache().get('s:' + token);
  if (!v) throw new Error('セッションが切れました。再ログインしてください');
  const u = JSON.parse(v);
  if (adminOnly && u.role !== '管理者') throw new Error('この操作は管理者のみ可能です');
  return u;
}

function priceSheet_() {
  let sh = book_(PRICE_SS_ID).getSheetByName(PRICE);
  if (!sh) {
    sh = book_(PRICE_SS_ID).insertSheet(PRICE);
    sh.getRange(1, 1, 5, 2).setValues([
      ['項目', '金額'], ['基本料金(1個口あたり)', 3000], ['組立設置加算(1配送あたり)', 2000],
      ['引取加算(1配送あたり)', 1000], ['消費税率', 0.1]]);
    sh.getRange('A1:B1').setFontWeight('bold');
  }
  return sh;
}

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  const m = String(v).match(/(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : String(v).trim();
}

function str_(v) { return v instanceof Date ? Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : String(v == null ? '' : v); }

function flag_(v) { v = String(v || '').trim(); return v !== '' && !/^(なし|無|無し|-|0|×)$/.test(v); }

function group_(rows) {
  const map = {}, order = [];
  rows.forEach(r => {
    const id = r['配送ID'] + '|' + r['配送日'];   // 配送ID+配送日で1件(再配送は別件)
    if (!map[id]) { map[id] = Object.assign({}, r, { 品名: [], 個口数: 0, items: [] }); order.push(id); }
    const g = map[id];
    if (r['品名']) g['品名'].push(r['品名']);
    g.items.push({ id: r['商品ID'], name: r['品名'], qty: Number(r['個口数']) || 1, setup: flag_(r['組立設置']), pick: flag_(r['引取']) });
    g['個口数'] += Number(r['個口数']) || 0;
    if (flag_(r['組立設置'])) g['組立設置'] = r['組立設置'];
    if (flag_(r['引取'])) g['引取'] = r['引取'];
    if (r['配送ステータス'] !== DONE) g['配送ステータス'] = r['配送ステータス'];
  });
  return order.map(id => { const g = map[id]; g['品名'] = g['品名'].join(' / '); return g; });
}

/* ===== 締め日時の自動確定(予定データ → 確定データ) ===== */
function cutoff_(dateStr) { // 配送日の締め日時
  const p = dateStr.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2] - CUTOFF_DAYS_BEFORE, CUTOFF_HOUR, 0, 0);
}

/** 締め日時を過ぎた予定データを確定データへ転記する(何度実行しても重複しません) */
function confirmDue() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const fs = fcSheet_(), cs = cfSheet_(), now = new Date();
    const ts = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    const due = readRows_(fs).filter(r => !r['転記日時'] && /^\d{4}-\d{2}-\d{2}$/.test(r['配送日']) && cutoff_(r['配送日']) <= now);
    if (!due.length) return 0;
    const have = {}, seen = {}; readAll_().forEach(r => have[key_(r)] = (have[key_(r)] || 0) + 1);
    const add = due.filter(r => { const k = key_(r); seen[k] = (seen[k] || 0) + 1; return seen[k] > (have[k] || 0); })
      .map(r => Object.assign({}, r, { '配送ステータス': '未配送', '配送完了時間': '', '完了メモ': '' }));
    ensureCols_(cs, ['取込日時', '完了メモ']);
    appendRows_(cs, add);
    const col = colOf_(fs, '転記日時'), rng = fs.getRange(2, col, fs.getLastRow() - 1, 1), v = rng.getValues();
    due.forEach(r => v[r._row - 2][0] = ts);
    rng.setNumberFormat('@').setValues(v);
    return add.length;
  } finally { lock.releaseLock(); }
}

/** 最初に1回だけ手動実行: 10分おきに confirmDue を動かすトリガーを作成 */
function setupTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'confirmDue').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('confirmDue').timeBased().everyMinutes(10).create();
}

function runConfirm(token) { auth_(token, true); return confirmDue(); } // 画面の「今すぐ実行」

/* ===== 1. CSV取込(予定データへ) ===== */
function importRows(token, rows) {
  auth_(token);
  const fs = fcSheet_(), have = {}, seen = {};
  ensureCols_(fs, ['取込日時', '転記日時', '完了メモ']);
  const cntOf = rs => rs.reduce((o, r) => (o[key_(r)] = (o[key_(r)] || 0) + 1, o), {}), fc = cntOf(readRows_(fs)), cf = cntOf(readAll_());
  Object.keys(fc).concat(Object.keys(cf)).forEach(k => have[k] = Math.max(fc[k] || 0, cf[k] || 0));   // 同じ行が複数ある場合は件数で比較
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const add = [], errors = []; let dup = 0;
  rows.forEach((r, i) => {
    if (!r['配送ID'] || !r['配送日'] || !r['顧客名']) { errors.push((i + 2) + '行目: 配送ID・配送日・顧客名は必須'); return; }
    r['配送日'] = normDate_(r['配送日']);
    const k = key_(r);
    seen[k] = (seen[k] || 0) + 1;
    if (seen[k] <= (have[k] || 0)) { dup++; return; }   // 取込済み(同じ行が2行あれば2行目も別の行として取り込む)
    r['配送日'] = normDate_(r['配送日']); r['取込日時'] = now;
    if (!r['配送ステータス']) r['配送ステータス'] = '予定';
    add.push(r);
  });
  appendRows_(fs, add);
  return { added: add.length, duplicated: dup, errors: errors, confirmed: confirmDue() };
}

/* ===== 画面用データ ===== */
function getCounts(token, ym) { // ym = 'yyyy-MM'  確定件数(total/done) + 予定件数(plan: 予定データの全件)
  auth_(token);
  const res = {}, slot = d => res[d] = res[d] || { total: 0, done: 0, plan: 0, absent: 0 };
  group_(readAll_().filter(r => r['配送日'].indexOf(ym) === 0)).forEach(g => {
    const d = slot(g['配送日']); d.total++; if (g['配送ステータス'] === DONE) d.done++; else if (g['配送ステータス'] === '不在') d.absent++;
  });
  const seen = {};
  readRows_(fcSheet_()).filter(r => r['配送日'].indexOf(ym) === 0).forEach(r => {
    const sk = r['配送ID'] + '|' + r['配送日']; if (seen[sk]) return; seen[sk] = 1; slot(r['配送日']).plan++;
  });
  return res;
}

function getForecast(token, from, to) { // 予定データ(転記日時があれば確定済み)
  auth_(token);
  return group_(readRows_(fcSheet_()).filter(o => o['配送日'] >= from && o['配送日'] <= to))   // 配送ID単位にまとめる
    .sort((a, b) => a['配送日'].localeCompare(b['配送日'])).slice(0, 1000);
}

function getConfirmed(token, from, to, status, kw) { // 確定データ
  auth_(token); kw = String(kw || '').trim();
  return group_(readAll_()).filter(g => g['配送日'] >= from && g['配送日'] <= to
      && (!status || g['配送ステータス'] === status)
      && (!kw || [g['顧客名'], g['住所'], g['配送ID']].join(' ').indexOf(kw) >= 0))
    .sort((a, b) => a['配送日'].localeCompare(b['配送日'])).slice(0, 1000);
}

/* ===== ステータス更新(確定データ) ===== */
function setStatus(token, id, date, status, memo) {
  auth_(token);
  if (['未配送', '配送完了', '不在', '持ち戻り', '搬入不可', 'その他'].indexOf(status) < 0) throw new Error('不正なステータスです');
  const sh = cfSheet_(), now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const rows = rowsOf_(id, date);
    if (!rows.length) throw new Error('確定データに配送IDが見つかりません(締め日時前の場合は確定後に入力できます)');
    rows.forEach(r => setCells_(sh, r._row, { '配送ステータス': status, '配送完了時間': status === '未配送' ? '' : now, '完了メモ': memo || '' }));
    return rows.length;
  } finally { lock.releaseLock(); }
}
function completeDelivery(token, id, date, memo, undo) { return setStatus(token, id, date, undo ? '未配送' : DONE, memo); }

/* ===== 指示書・検索・請求(確定データを使用) ===== */
function getDeliveries(token, date) {
  auth_(token);
  const rows = readAll_().filter(r => r['配送日'] === date);
  return group_(rows).map(g => Object.assign(g, { k: sign_(g['配送ID'], g['配送日']) })).sort((a, b) => (a['指定時間枠'] + a['時間カテゴリ']).localeCompare(b['指定時間枠'] + b['時間カテゴリ']));
}

function searchDeliveries(token, type, q) {
  auth_(token);
  q = String(q || '').trim();
  const digits = x => String(x).replace(/\D/g, '');
  if (!q || (type === 'tel' && !digits(q))) return [];
  const rows = readAll_().filter(r => type === 'tel'
    ? digits(r['連絡先']).indexOf(digits(q)) >= 0 : r['配送ID'].indexOf(q) >= 0);
  return group_(rows).sort((a, b) => b['配送日'].localeCompare(a['配送日'])).slice(0, 100);
}

/** 1配送あたりの料金計算。※料金ルール確定後は、この関数だけを差し替えます(現在は「料金設定」シートの仮の単価) */
function calcCharges_(g, price) {
  if (g['配送ステータス'] === '不在') return { '配送料': price.base * (g['個口数'] || 1), '組立設置': 0, '引取': 0 };   // 不在=1件(配送料のみ) ※金額ルールは確定後に調整
  return {
    '配送料': price.base * (g['個口数'] || 1),
    '組立設置': flag_(g['組立設置']) ? price.setup : 0,
    '引取': flag_(g['引取']) ? price.pickup : 0
  };
}

function makeInvoice(token, from, to, billTo) {
  auth_(token, true);
  billTo = String(billTo || '').trim();
  if (!from || !to || from > to) throw new Error('配送日の期間を正しく指定してください');
  if (!billTo) throw new Error('請求先を入力してください');
  const p = priceSheet_().getRange(2, 2, 4, 1).getValues().map(x => Number(x[0]));
  const price = { base: p[0], setup: p[1], pickup: p[2] }, taxRate = p[3];

  const list = group_(readAll_().filter(r => r['配送日'] >= from && r['配送日'] <= to))
    .filter(g => BILLABLE.indexOf(g['配送ステータス']) >= 0)
    .sort((a, b) => a['配送日'].localeCompare(b['配送日']) || a['配送ID'].localeCompare(b['配送ID']));
  if (!list.length) throw new Error('この期間に、請求対象(配送完了・不在)の確定データがありません');

  const head = ['配送日', '配送ID', 'お届け先', '配送カテゴリ', '品名', '個口数', '配送料', '組立設置', '引取', '小計', 'ステータス', '備考'];
  const view = [], rows = [];
  let sum = { '配送料': 0, '組立設置': 0, '引取': 0 };
  list.forEach((g, i) => {
    const c = calcCharges_(g, price), sub = c['配送料'] + c['組立設置'] + c['引取'];
    Object.keys(sum).forEach(k => sum[k] += c[k]);
    const base = [g['配送日'], g['配送ID'], g['顧客名'], g['配送カテゴリ'], g['品名'], g['個口数'], c['配送料'], c['組立設置'], c['引取']];
    view.push(base.concat([sub, g['配送ステータス'], '']));
    rows.push(base.concat(['=SUM(G' + (i + 2) + ':I' + (i + 2) + ')', g['配送ステータス'], '']));   // 小計は数式(手修正しても請求書に反映)
  });
  const subtotal = sum['配送料'] + sum['組立設置'] + sum['引取'], tax = Math.floor(subtotal * taxRate);

  let folder = null;
  if (BILLING_FOLDER_ID) {
    try { folder = DriveApp.getFolderById(BILLING_FOLDER_ID); }
    catch (e) { throw new Error('請求書の保存先フォルダが見つかりません(BILLING_FOLDER_ID を確認してください)'); }
  }
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm');
  const no = 'INV-' + stamp, name = '請求書_' + billTo + '_' + from.replace(/-/g, '') + '-' + to.replace(/-/g, '') + '_' + stamp;
  const file = SpreadsheetApp.create(name);
  if (folder) DriveApp.getFileById(file.getId()).moveTo(folder);
  const inv = file.getSheets()[0]; inv.setName('請求書');
  const det = file.insertSheet('明細');

  // --- 明細シート ---
  const n = rows.length, last = n + 1;
  det.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8eefc');
  det.getRange(2, 1, n, head.length).setValues(rows);
  det.getRange(2, 7, n, 4).setNumberFormat('#,##0');
  det.setFrozenRows(1); det.autoResizeColumns(1, head.length);

  // --- 請求書シート(金額は明細シートを集計する数式) ---
  const S = c => "=SUM('明細'!" + c + '2:' + c + last + ')';
  const money = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  inv.getRange('A1').setValue('請求書').setFontSize(22).setFontWeight('bold');
  inv.getRange('C2:D3').setValues([['請求番号', no], ['発行日', money]]);
  inv.getRange('A3').setValue(billTo + ' 御中').setFontSize(14).setFontWeight('bold');
  inv.getRange('A5:B6').setValues([['請求期間', from + ' 〜 ' + to], ['配送件数', n + '件']]);
  inv.getRange('A8:B8').setValues([['ご請求金額(税込)', '=B16']]).setFontSize(14).setFontWeight('bold');
  inv.getRange('A10:B10').setValues([['項目', '金額']]).setFontWeight('bold').setBackground('#e8eefc');
  inv.getRange('A11:B16').setFormulas([
    ['配送料', S('G')], ['組立設置', S('H')], ['引取', S('I')],
    ['小計(税抜)', '=SUM(B11:B13)'], ['消費税(' + Math.round(taxRate * 100) + '%)', '=ROUNDDOWN(B14*' + taxRate + ',0)'], ['合計(税込)', '=B14+B15']]);
  inv.getRange('A16:B16').setFontWeight('bold');
  inv.getRange('B8:B16').setNumberFormat('#,##0');
  inv.getRange('A10:B16').setBorder(true, true, true, true, true, true);
  const info = [[ISSUER.name], [ISSUER.address], ['TEL ' + ISSUER.tel], [ISSUER.bank ? '振込先: ' + ISSUER.bank : '']]
    .filter(r => r[0] && r[0] !== 'TEL ');
  if (info.length) inv.getRange(18, 1, info.length, 1).setValues(info);
  inv.setColumnWidth(1, 170); inv.setColumnWidth(2, 190); inv.setColumnWidth(3, 90); inv.setColumnWidth(4, 160);
  file.setActiveSheet(inv);

  return { head: head, rows: view, subtotal: subtotal, tax: tax, total: subtotal + tax, name: name, url: file.getUrl() };
}


function getAppUrl(token) { auth_(token); return ScriptApp.getService().getUrl(); }

function getDelivery(token, id, date) {
  auth_(token);
  return group_(rowsOf_(id, date))[0] || null;
}

/* ===== ドライバー報告(指示書のQR: ログイン不要・その配送のみ) ===== */

function secret_() { // QR用の署名キー(初回に自動生成してスクリプトのプロパティに保存)
  const pr = PropertiesService.getScriptProperties();
  let s = pr.getProperty('REPORT_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); pr.setProperty('REPORT_SECRET', s); }
  return s;
}
function sign_(id, date) { // 配送ID+配送日ごとの報告用キー(QRのURLに入る)
  return Utilities.computeHmacSha256Signature(String(id) + '|' + String(date), secret_()).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').slice(0, 20);
}
function chk_(id, date, k) {
  if (!id || !date || !k || sign_(id, date) !== String(k)) throw new Error('このQRでは報告できません。指示書のQRを読み取り直してください');
}
/** photos: JPEG画像のBase64文字列の配列(最大5枚) */

const RECEIPT_FOLDER_ID = '1mAAN-IafFCgZMXOtBk4RzGAKqVF6ZMLO';    // 受領書の写真(配送完了)の保存先フォルダ
const ABSENT_FOLDER_ID = '1gzCyO-f6xcZ8NIk5Spcl6FlZxF31ifIQ';   // 不在・持ち戻りの写真の保存先(サブフォルダ名: 不在連絡-配送日-運送会社-配送ID)
const SLACK_WEBHOOK_URL = '';    // 不在・持ち戻りの報告先Slackの Incoming Webhook URL(スクリプトのプロパティ SLACK_WEBHOOK_URL でも可)
const SLOT_LABEL = { receipt: '受領書', notice: '不在票投函', call: '発信履歴', other: '写真' };

function folder_(id, propKey, name) { // 指定フォルダ。未指定なら自動作成して再利用
  if (id) return DriveApp.getFolderById(id);
  const pr = PropertiesService.getScriptProperties(), fid = pr.getProperty(propKey);
  try { if (fid) return DriveApp.getFolderById(fid); } catch (e) { /* 作り直す */ }
  const f = DriveApp.createFolder(name); pr.setProperty(propKey, f.getId()); return f;
}
function subFolder_(parent, name) { // 配送IDごとのフォルダ
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** Slackへ投稿。Bot(SLACK_BOT_TOKEN + SLACK_CHANNEL_ID)を優先し、無ければWebhookを使用。戻り値: sent / error / skipped */
function slackPost_(text) {
  const pr = PropertiesService.getScriptProperties();
  const token = pr.getProperty('SLACK_BOT_TOKEN'), channel = pr.getProperty('SLACK_CHANNEL_ID');
  try {
    if (token && channel) {
      const res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
        method: 'post', contentType: 'application/json; charset=utf-8', headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ channel: channel, text: text }), muteHttpExceptions: true });
      const r = JSON.parse(res.getContentText());
      if (!r.ok) Logger.log('Slack error: ' + r.error);   // 例: not_in_channel(Botをチャンネルに招待) / channel_not_found / invalid_auth
      return r.ok ? 'sent' : 'error';
    }
    const url = SLACK_WEBHOOK_URL || pr.getProperty('SLACK_WEBHOOK_URL');
    if (!url) return 'skipped';
    const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: text }), muteHttpExceptions: true });
    return res.getResponseCode() === 200 ? 'sent' : 'error';
  } catch (e) { Logger.log(e); return 'error'; }
}

function notifySlack_(status, row, date, memo, lines, dir, carrier) {
  const text = ['*【' + status + '】* 配送ID: ' + row['配送ID'], '配送日: ' + date + ' ' + (row['指定時間枠'] || ''),
    'お届け先: ' + row['顧客名'] + ' 様', '運送会社: ' + (carrier || '-'), '住所: ' + row['住所'], '連絡先: ' + row['連絡先'],
    memo ? 'メモ: ' + memo : '', dir ? '写真フォルダ: ' + dir.getUrl() : ''].concat(lines).filter(Boolean).join('\n');
  return slackPost_(text);
}

function reportGet(id, date, k) {
  chk_(id, date, k);
  const g = group_(rowsOf_(id, date))[0];
  if (!g) throw new Error('配送データが見つかりません(締め日時前の配送は、確定後に報告できます)');
  return g;
}

/** photos: [{slot:'receipt'|'notice'|'call'|'other', data:JPEGのBase64}]
 *  配送完了=受領書の写真+受取人(本人/代理人/その他) / 不在・持ち戻り=不在票投函+発信履歴のスクショ → 配送IDフォルダに保存してSlackへ報告 */
function reportSubmit(id, date, k, status, recipient, memo, photos, carrier) {
  chk_(id, date, k);
  if (['配送完了', '不在', '持ち戻り', '搬入不可', 'その他'].indexOf(status) < 0) throw new Error('ステータスを選んでください');
  photos = photos || [];
  if (photos.length > 12) throw new Error('写真が多すぎます');
  const cnt = {};
  photos.forEach(p => { if (!SLOT_LABEL[p.slot]) throw new Error('不正な写真の種類です'); cnt[p.slot] = (cnt[p.slot] || 0) + 1; });
  const absent = status === '不在' || status === '持ち戻り';
  if (status === '配送完了') {
    if (['本人', '代理人', 'その他'].indexOf(recipient) < 0) throw new Error('受取人を選んでください');
    if (!cnt.receipt) throw new Error('受領書の写真を添付してください');
  }
  carrier = String(carrier || '').replace(/[\\\/:*?"<>|]/g, '').trim().slice(0, 40);   // フォルダ名に使えない文字を除く
  if (absent) {
    if (!carrier) throw new Error('運送会社を入力してください');
    if (!cnt.notice) throw new Error('不在票投函の写真を添付してください');
    if (!cnt.call) throw new Error('発信履歴のスクショを添付してください');
  }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = cfSheet_(), rows = rowsOf_(id, date);
    if (!rows.length) throw new Error('確定データに配送IDが見つかりません');
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    const lines = []; let dir = null;
    if (photos.length) {
      dir = absent ? subFolder_(folder_(ABSENT_FOLDER_ID, 'ABSENT_FOLDER', '配送報告_不在持ち戻り'), '不在連絡-' + date.replace(/-/g, '') + '-' + carrier + '-' + id)
                   : folder_(RECEIPT_FOLDER_ID, 'RECEIPT_FOLDER', '配送報告_受領書');
      const seq = {};
      photos.forEach(p => {
        seq[p.slot] = (seq[p.slot] || 0) + 1;
        const blob = Utilities.newBlob(Utilities.base64Decode(p.data), 'image/jpeg', id + '_' + date + '_' + SLOT_LABEL[p.slot] + '_' + seq[p.slot] + '.jpg');
        lines.push(SLOT_LABEL[p.slot] + ': ' + dir.createFile(blob).getUrl());
      });
    }
    const prev = rows[0]['報告写真'] || '', all = (prev ? prev.split('\n') : []).concat(lines).join('\n');
    rows.forEach(r => setCells_(sh, r._row, { '配送ステータス': status, '配送完了時間': now, '受取人名': status === '配送完了' ? recipient : '', '完了メモ': memo || '', '報告写真': all, '運送会社': absent ? carrier : '' }));
    return { rows: rows.length, slack: absent ? notifySlack_(status, rows[0], date, memo, lines, dir, carrier) : 'none' };
  } finally { lock.releaseLock(); }
}

/** Slack連携のテスト: エディタで1回実行すると、設定したチャンネルにテスト投稿が届きます(結果は「実行ログ」に出ます) */
function testSlack() {
  Logger.log(slackPost_('配送受注システムからのテスト投稿です'));
}
