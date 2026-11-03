// JVM Foundation Attendance - shared storage
// Paste this whole file into Extensions > Apps Script in the Google Sheet, change SECRET, then deploy as a Web app.

const SECRET = 'jvm2026';   // the school code teachers type once on their phone
const SHEET_NAME = 'Records';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange('A:B').setNumberFormat('@');
    sh.appendRow(['key', 'value', 'clientTime', 'serverTime', 'deleted']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, service: 'JVM Foundation Attendance' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.token !== SECRET) return json_({ ok: false, error: 'bad_token' });
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try { out = handle_(req); } finally { lock.releaseLock(); }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return json_(out);
}

function handle_(req) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
  const index = {};
  data.forEach(function (r, i) { index[String(r[0])] = i; });
  const now = Date.now();
  const fresh = [];

  (req.changes || []).forEach(function (c) {
    const row = [String(c.k), JSON.stringify(c.v || []), Number(c.t), now, c.d ? 1 : 0];
    const i = index[row[0]];
    if (i === undefined) {
      fresh.push(row);
      index[row[0]] = data.length + fresh.length - 1;
      data.push(row);
    } else if (Number(c.t) > Number(data[i][2])) {
      data[i] = row;
      if (i + 2 <= last) sh.getRange(i + 2, 1, 1, 5).setValues([row]);
    }
  });
  if (fresh.length) {
    sh.getRange(last + 1, 1, fresh.length, 5).setValues(fresh);
  }

  const since = Number(req.since) || 0;
  const rows = [];
  data.forEach(function (r) {
    if (Number(r[3]) >= since) {
      rows.push({ k: String(r[0]), v: JSON.parse(r[1] || '[]'), t: Number(r[2]), s: Number(r[3]), d: Number(r[4]) ? 1 : 0 });
    }
  });
  return { ok: true, now: now, rows: rows };
}
