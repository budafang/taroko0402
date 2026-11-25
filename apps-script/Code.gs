/**
 * 太魯閣號 0402 溝通管理模擬｜Google 試算表後端
 *
 * 用法：在 Google 試算表選「擴充功能 → Apps Script」，把本檔全部貼上，
 * 再「部署 → 新增部署作業 → 網頁應用程式」，執行身分選「我」、存取權選「所有人」。
 * 部署後的網址貼到網站的 config.js。
 *
 * 寫入兩個工作表（第一次使用時自動建立）：
 *   加入紀錄：學生輸入學號開始遊戲
 *   提交紀錄：學生完成遊戲並交出學習卡（平時成績以這張為準）
 * doGet 只回傳全班匿名統計（不含學號），給老師投影頁使用。
 */

const SHEET_JOIN = '加入紀錄';
const SHEET_SUBMIT = '提交紀錄';
const TZ = 'Asia/Taipei';

const JOIN_HEADERS = ['時間', '日期', '學號', '裝置碼'];
const SUBMIT_HEADERS = [
  '提交時間', '日期', '學號', '角色',
  '決策1', '決策2', '決策3', '決策4', '決策5', '有效決策數(前4關)',
  '最意外的決策點', '意外的原因',
  '機制類型', '情境(當…時)', '誰通知', '通知誰', '管道', '時限', '確認方式', '機制整句',
  '以前以為', '現在知道', '心得字數', '裝置碼'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const sid = cleanSid(d.sid);
    if (!sid) return out({ ok: false, error: 'bad_sid' });
    const now = new Date();
    const time = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss');
    const date = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const device = clip(d.device, 8);

    if (d.type === 'join') {
      sheet(SHEET_JOIN, JOIN_HEADERS).appendRow([time, date, sid, device]);
      return out({ ok: true, time: time });
    }

    if (d.type === 'submit') {
      const choices = Array.isArray(d.choices) ? d.choices : [];
      const r = d.reflection || {};
      const m = r.mechanism || {};
      const text = [r.surpriseWhy, m.when, m.who, m.whom, m.channel, m.deadline, m.confirm, r.before, r.after].join('');
      sheet(SHEET_SUBMIT, SUBMIT_HEADERS).appendRow([
        time, date, sid, clip(d.role, 20),
        clip(choices[0], 2), clip(choices[1], 2), clip(choices[2], 2), clip(choices[3], 2), clip(choices[4], 2),
        Number(d.score) || 0,
        clip(r.surprise, 2), clip(r.surpriseWhy, 300),
        clip(m.type, 20), clip(m.when, 80), clip(m.who, 40), clip(m.whom, 40), clip(m.channel, 40), clip(m.deadline, 40), clip(m.confirm, 80), clip(m.sentence, 400),
        clip(r.before, 80), clip(r.after, 80), text.length, device
      ]);
      return out({ ok: true, time: time });
    }

    return out({ ok: false, error: 'bad_type' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 老師投影頁：?action=stats&date=yyyy-MM-dd（不給日期就是今天）。只回傳匿名統計。 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'stats') return out({ ok: true, service: 'taroko0402' });
  const date = /^\d{4}-\d{2}-\d{2}$/.test(p.date || '') ? p.date : Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  const joined = new Set();
  rows(SHEET_JOIN).forEach(r => { if (dateOf(r[1]) === date) joined.add(String(r[2])); });

  // 同一學號重複提交時，以最後一次為準
  const latest = {};
  rows(SHEET_SUBMIT).forEach(r => { if (dateOf(r[1]) === date) latest[String(r[2])] = r; });
  const subs = Object.keys(latest).map(k => latest[k]);
  subs.forEach(r => joined.add(String(r[2])));

  const choices = [0, 1, 2, 3, 4].map(i => count(subs.map(r => r[4 + i]), ['A', 'B', 'C']));
  const scores = count(subs.map(r => String(r[9])), ['0', '1', '2', '3', '4']);
  const surprise = count(subs.map(r => String(r[10])), ['1', '2', '3', '4', '5']);
  const types = count(subs.map(r => String(r[12])), ['規劃溝通', '管理溝通', '監視溝通']);
  // 文字牆打亂順序，避免從提交順序對回是誰寫的
  const wall = shuffle(subs.map(r => ({ type: String(r[12]), mechanism: String(r[19]), before: String(r[20]), after: String(r[21]) })));

  return out({ ok: true, date: date, joined: joined.size, submitted: subs.length, choices: choices, scores: scores, surprise: surprise, types: types, wall: wall });
}

/* ── helpers ── */
function sheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@'); // 全部當文字存，避免學號前導 0 或日期被自動轉換
  }
  return sh;
}
function rows(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function dateOf(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v).slice(0, 10);
}
function cleanSid(v) {
  const s = String(v || '').trim().toUpperCase();
  return /^[A-Z0-9]{5,12}$/.test(s) ? s : '';
}
function clip(v, n) { return String(v == null ? '' : v).replace(/^[=+\-@]/, "'$&").slice(0, n); }
function count(list, keys) {
  const o = {}; keys.forEach(k => { o[k] = 0; });
  list.forEach(v => { if (o[v] !== undefined) o[v]++; });
  return o;
}
function shuffle(a) {
  for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = a[i]; a[i] = a[j]; a[j] = t; }
  return a;
}
function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
