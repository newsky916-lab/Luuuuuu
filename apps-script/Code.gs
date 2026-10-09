/**
 * 福岡行旅帳 — Google 試算表後端
 *
 * 貼到試算表的「擴充功能 → Apps Script」裡：
 *   1. 執行 setup() 一次（會建立「旅帳_」開頭的分頁、放入目前的資料、產生連線密碼）
 *   2. 部署 → 新增部署作業 → 網頁應用程式，執行身分「我」，存取權「所有人」
 *   3. 執行 shareLink()，把執行記錄裡的分享連結複製出來
 *
 * 只會讀寫「旅帳_」開頭的分頁，原本的 Money、list、Note 分頁不會動到。
 *
 * @OnlyCurrentDoc  只要求存取這一份試算表的權限
 */

const PAGE_URL = 'https://newsky916-lab.github.io/Luuuuuu/';

const TABS = {
  plans:    { name: '旅帳_行程', cols: ['id', 'day', 'time', 'title', 'place', 'note', 'sample', 'created'] },
  todos:    { name: '旅帳_清單', cols: ['id', 'group', 'text', 'done', 'sample', 'created'] },
  expenses: { name: '旅帳_記帳', cols: ['id', 'day', 'cat', 'note', 'amount', 'currency', 'payer', 'to', 'sample', 'created'] }
};
const SETTINGS_TAB = '旅帳_設定';
const NUM_FIELDS = ['day', 'amount', 'created', 'days', 'rate'];
const BOOL_FIELDS = ['done', 'sample'];

// ---------- one-time setup ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KEY')) props.setProperty('KEY', Utilities.getUuid().replace(/-/g, ''));

  let fresh = false;
  for (const coll of Object.keys(TABS)) {
    const t = TABS[coll];
    if (!ss.getSheetByName(t.name)) { makeSheet_(ss, t.name, t.cols); fresh = true; }
  }
  if (!ss.getSheetByName(SETTINGS_TAB)) { makeSheet_(ss, SETTINGS_TAB, ['key', 'value']); fresh = true; }
  if (fresh) seed_();

  Logger.log('設定完成。連線密碼：' + props.getProperty('KEY'));
  Logger.log('下一步：部署成網頁應用程式，然後執行 shareLink()。');
}

function shareLink() {
  const url = ScriptApp.getService().getUrl();
  if (!url) { Logger.log('還沒有部署。請先「部署 → 新增部署作業 → 網頁應用程式」。'); return; }
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  Logger.log('分享連結（傳給李先生，用手機打開就能一起編輯）：');
  Logger.log(PAGE_URL + '#api=' + encodeURIComponent(url) + '&k=' + key);
}

function makeSheet_(ss, name, cols) {
  const sh = ss.insertSheet(name);
  sh.getRange('A:Z').setNumberFormat('@'); // keep "09:30", "2026-10-19" as plain text
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

// Data copied from the Claude version of the trip notebook (2026-10-09)
function seed_() {
  writeSettings_({ name: '李氏夫妻福岡行', start: '2026-10-19', days: 6, rate: 0.24691358, travelers: ['李先生', '李太太'] });
  const todos = [
    ['t1', '出發前', '確認護照效期', true, 1],
    ['t2', '出發前', '填寫 Visit Japan Web', true, 2],
    ['t3', '行李', '行動電源放隨身行李', true, 3],
    ['t4', '當地', '儲值 IC 卡（SUGOCA / Suica 皆可）', true, 4],
    ['w1', '景點', '(klook)糸島一日遊: 白絲瀑布 & 宮地獄神社 & 一蘭之森 & 櫻井二見浦夫婦岩 & 天使之翼 & 椰子樹鞦韆 (福岡出發)', false, 11],
    ['w2', '景點', '(klook)熊本城 & 阿蘇火山 (含火山接駁車) & 草千里 & 阿蘇神社 & 黑川溫泉一日遊 (福岡/熊本出發)', false, 12],
    ['w3', '景點', '(klook)【福岡出發】湯布院童話小鎮＆別府一日遊｜如意輪寺・金麟湖水上鳥居・湯布院・別府地獄巡遊（含門票）', false, 13],
    ['w4', '景點', '(klook)九州一日遊: 高千穗小火車&高千穗峽&天安河原 (中英韓文導遊, 福岡出發)', false, 14],
    ['w5', '景點', '(klool)九州大分｜由布院＆九州自然動物公園＆海地獄＆灶地獄一日遊｜福岡出發', false, 15]
  ];
  todos.forEach(([id, group, text, sample, created]) => upsert_('todos', { id, group, text, done: false, sample, created }));
  upsert_('expenses', { id: 'x-flight', day: 0, cat: '機票', note: '機票', amount: 30614, currency: 'TWD', payer: '李先生', to: '', sample: false, created: 1 });
  upsert_('expenses', { id: 'x-hotel', day: 0, cat: '住宿', note: '住宿', amount: 11868, currency: 'TWD', payer: '李太太', to: '', sample: false, created: 2 });
  upsert_('expenses', { id: 'x-repay', day: 0, cat: '轉帳', note: '李太太已匯款給李先生（機票一半）', amount: 15307, currency: 'TWD', payer: '李太太', to: '李先生', sample: false, created: 3 });
}

// ---------- web app ----------
function doGet(e) {
  if (!authorized_(e.parameter.k)) return json_({ error: 'unauthorized' });
  return json_(readAll_());
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad_request' }); }
  if (!authorized_(body.k)) return json_({ error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.op === 'upsert' && TABS[body.coll] && body.item && body.item.id) upsert_(body.coll, body.item);
    else if (body.op === 'delete' && TABS[body.coll] && body.id) remove_(body.coll, String(body.id));
    else if (body.op === 'settings' && body.settings) writeSettings_(body.settings);
    else return json_({ error: 'bad_request' });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return json_(readAll_());
}

function authorized_(k) {
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  return !!key && k === key;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- storage ----------
function sheet_(coll) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS[coll].name);
}

function readAll_() {
  const out = { settings: readSettings_() };
  for (const coll of Object.keys(TABS)) out[coll] = readColl_(coll);
  return out;
}

function readColl_(coll) {
  const values = sheet_(coll).getDataRange().getValues();
  const head = values[0];
  return values.slice(1).filter(r => r[0] !== '').map(r => {
    const o = {};
    head.forEach((h, i) => { o[h] = norm_(h, r[i]); });
    return o;
  });
}

function norm_(field, v) {
  if (NUM_FIELDS.indexOf(field) >= 0) return v === '' ? 0 : Number(v);
  if (BOOL_FIELDS.indexOf(field) >= 0) return v === true || String(v).toUpperCase() === 'TRUE';
  return v === null ? '' : String(v);
}

function upsert_(coll, item) {
  const sh = sheet_(coll);
  const cols = TABS[coll].cols;
  const row = cols.map(c => {
    const v = item[c];
    if (v === undefined || v === null) return '';
    return typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
  });
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
  const at = ids.indexOf(String(item.id));
  if (at > 0) sh.getRange(at + 1, 1, 1, cols.length).setValues([row]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, cols.length).setNumberFormat('@').setValues([row]);
}

function remove_(coll, id) {
  const sh = sheet_(coll);
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
  const at = ids.indexOf(id);
  if (at > 0) sh.deleteRow(at + 1);
}

function readSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_TAB);
  const out = {};
  sh.getDataRange().getValues().slice(1).forEach(([k, v]) => {
    if (!k) return;
    if (k === 'travelers') { try { out[k] = JSON.parse(v); } catch (err) { out[k] = String(v).split(/[,，、]/).map(s => s.trim()).filter(String); } }
    else out[k] = NUM_FIELDS.indexOf(k) >= 0 ? Number(v) : String(v);
  });
  return out;
}

function writeSettings_(s) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_TAB);
  const keys = ['name', 'start', 'days', 'rate', 'travelers'];
  const rows = keys.map(k => [k, k === 'travelers' ? JSON.stringify(s[k] || []) : String(s[k] === undefined || s[k] === null ? '' : s[k])]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
}
