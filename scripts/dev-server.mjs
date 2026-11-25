// 本機預覽＋模擬 Google Apps Script 後端（不連 Google）。
// 用法：node scripts/dev-server.mjs [port]，開 http://localhost:8770/ 與 /teacher.html
// 直接執行 apps-script/Code.gs，以記憶體模擬試算表；/__sheets 可看目前寫入的列。
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const port = Number(process.argv[2]) || 8770;

/* ── 模擬 Apps Script 的服務 ── */
const sheets = {};
function makeSheet(name) {
  const data = [];
  const range = (r, c, nr = 1, nc = 1) => ({
    setValues(v) { v.forEach((row, i) => { data[r - 1 + i] = data[r - 1 + i] || []; row.forEach((x, j) => { data[r - 1 + i][c - 1 + j] = x; }); }); return this; },
    getValues() { return Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (data[r - 1 + i] || [])[c - 1 + j] ?? '')); },
    setFontWeight() { return this; }, setNumberFormat() { return this; }
  });
  return {
    data,
    getRange(a, b, c, d) { return typeof a === 'string' ? range(1, 1) : range(a, b, c, d); },
    appendRow(row) { data.push(row.map(String)); },
    setFrozenRows() {},
    getLastRow() { return data.length; },
    getLastColumn() { return Math.max(0, ...data.map(r => r.length)); }
  };
}
const ss = {
  getSheetByName: n => sheets[n] || null,
  insertSheet: n => (sheets[n] = makeSheet(n))
};
const pad = n => String(n).padStart(2, '0');
const ctx = vm.createContext({
  SpreadsheetApp: { getActiveSpreadsheet: () => ss },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Utilities: {
    formatDate(date, tz, fmt) {
      const t = new Date(date.getTime() + 8 * 3600e3); // Asia/Taipei
      const d = `${t.getUTCFullYear()}-${pad(t.getUTCMonth() + 1)}-${pad(t.getUTCDate())}`;
      return fmt === 'yyyy-MM-dd' ? d : `${d} ${pad(t.getUTCHours())}:${pad(t.getUTCMinutes())}:${pad(t.getUTCSeconds())}`;
    }
  },
  ContentService: { createTextOutput: s => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
  JSON, Date, Math, String, Number, Array, Object, Set
});
const gas = vm.runInContext(fs.readFileSync(path.join(root, 'apps-script', 'Code.gs'), 'utf8') + '\n;({ doGet, doPost })', ctx);

/* ── 靜態檔＋端點 ── */
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css' };
http.createServer((req, res) => {
  const url = new URL(req.url, `http://localhost:${port}`);
  const send = (code, body, type = 'application/json; charset=utf-8') => { res.writeHead(code, { 'Content-Type': type, 'Access-Control-Allow-Origin': '*' }); res.end(body); };

  if (url.pathname === '/exec') {
    if (req.method === 'GET') return send(200, gas.doGet({ parameter: Object.fromEntries(url.searchParams) }).s);
    let body = '';
    req.on('data', c => { body += c; });
    req.on('end', () => send(200, gas.doPost({ postData: { contents: body } }).s));
    return;
  }
  if (url.pathname === '/__sheets') return send(200, JSON.stringify(Object.fromEntries(Object.entries(sheets).map(([k, v]) => [k, v.data])), null, 1));
  if (url.pathname === '/config.js') return send(200, "window.TAROKO_CONFIG = { endpoint: '/exec' };", types['.js']);

  const file = path.join(root, url.pathname === '/' ? 'index.html' : decodeURIComponent(url.pathname));
  if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) return send(404, 'not found', 'text/plain');
  send(200, fs.readFileSync(file), types[path.extname(file)] || 'application/octet-stream');
}).listen(port, () => console.log(`taroko0402 dev server: http://localhost:${port}/  老師頁 /teacher.html  資料 /__sheets`));
