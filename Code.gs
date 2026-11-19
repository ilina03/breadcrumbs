/**
 * breadcrumbs : Google Sheet sync backend (Google Apps Script)
 *
 * Paste this into Extensions → Apps Script in your habits Google Sheet.
 * 1. Change SECRET below to a long random phrase. Don't commit your edited
 *    copy to GitHub: your repo is public.
 * 2. Run `setup` once (pick it in the function dropdown, then click Run).
 * 3. Deploy → New deployment → Web app → Execute as: Me,
 *    Who has access: Anyone. Copy the /exec URL into the app.
 *
 * Sheet layout:
 *   habits tab:  id | name | last 28 days (formula)
 *   log tab:      date | pomodoros | <one checkbox column per habit id>
 *   sessions tab: one row per focus block: when, how long, your breadcrumb
 * You can tick checkboxes or edit habits right in the sheet too.
 */

const SECRET = 'change-me-to-a-long-random-phrase';
const LOG = 'log';
const HABITS = 'habits';
const SESSIONS = 'sessions';
const SESSION_HEAD = ['id', 'date', 'start', 'focus min', 'break min', 'breadcrumb', 'finished'];
const SESSIONS_RETURNED = 150; // how many recent sessions the app keeps

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const l = ss.getSheetByName(LOG) || ss.insertSheet(LOG);
  if (l.getLastRow() === 0) l.getRange(1, 1, 1, 2).setValues([['date', 'pomodoros']]);
  l.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  l.setFrozenRows(1);

  const h = ss.getSheetByName(HABITS) || ss.insertSheet(HABITS);
  if (h.getLastRow() === 0) {
    h.getRange(1, 1, 4, 2).setValues([['id', 'name'], ['read', 'Read'], ['gym', 'Gym'], ['water', 'Water']]);
  }
  h.getRange(1, 3).setValue('last 28 days');
  for (let r = 2; r <= 20; r++) {
    h.getRange(r, 3).setFormula(
      `=IF(A${r}="","",IFERROR(COUNTIFS(${LOG}!$A:$A,">="&(TODAY()-27),` +
      `INDEX(${LOG}!$A:$Z,0,MATCH(A${r},${LOG}!$1:$1,0)),TRUE),0))`);
  }
  h.setFrozenRows(1);

  ensureColumns(l, readHabits(h));

  const se = ss.getSheetByName(SESSIONS) || ss.insertSheet(SESSIONS);
  if (se.getLastRow() <= 1) { // no sessions yet: (re)write the header so it matches this script
    se.getRange(1, 1, 1, Math.max(se.getLastColumn(), SESSION_HEAD.length)).clearContent();
    se.getRange(1, 1, 1, SESSION_HEAD.length).setValues([SESSION_HEAD]);
  }
  se.setFrozenRows(1);
  se.getRange('B2:C').setNumberFormat('@'); // keep dates/times as typed text
  se.setColumnWidth(6, 360);
  se.hideColumns(1); // the id is only for the app

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
}

function doGet() {
  return out({ok: true, note: 'this endpoint only answers POST requests from the app'});
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({error: 'bad request'}); }
  if (SECRET.indexOf('change-me') === 0) return out({error: 'set SECRET in the script first'});
  if (body.secret !== SECRET) return out({error: 'wrong secret'});

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.ops && body.ops.length) applyOps(body.ops);
    return out(readAll());
  } finally {
    lock.releaseLock();
  }
}

// ---------- helpers ----------
function sheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const l = ss.getSheetByName(LOG), h = ss.getSheetByName(HABITS);
  if (!l || !h) throw new Error('run setup() first');
  return {ss, l, h, tz: ss.getSpreadsheetTimeZone()};
}

function readHabits(h) {
  if (h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, 2).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({id: String(r[0]).trim(), name: String(r[1] || r[0]).trim()}));
}

// Make sure every habit id has a checkbox column in the log tab.
function ensureColumns(l, habits) {
  const width = Math.max(l.getLastColumn(), 2);
  const head = l.getRange(1, 1, 1, width).getValues()[0].map(String);
  habits.forEach(hb => {
    if (head.indexOf(hb.id) === -1) {
      const col = head.length + 1;
      l.getRange(1, col).setValue(hb.id);
      if (l.getLastRow() > 1) l.getRange(2, col, l.getLastRow() - 1, 1).insertCheckboxes();
      head.push(hb.id);
    }
  });
  return head;
}

function dateKey(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = String(v).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function applyOps(ops) {
  const sessionOps = ops.filter(op => op && op.session && op.session.id);
  if (sessionOps.length) applySessions(sessionOps.map(op => op.session));
  ops = ops.filter(op => op && !op.session);
  if (!ops.length) return;
  const {l, h, tz} = sheets();
  const head = ensureColumns(l, readHabits(h));
  const rowOf = {};
  if (l.getLastRow() > 1) {
    l.getRange(2, 1, l.getLastRow() - 1, 1).getValues()
      .forEach((r, i) => { const k = dateKey(r[0], tz); if (k) rowOf[k] = i + 2; });
  }
  ops.forEach(op => {
    if (!op || !/^\d{4}-\d{2}-\d{2}$/.test(op.date)) return;
    let row = rowOf[op.date];
    if (!row) {
      row = l.getLastRow() + 1;
      if (head.length > 2) l.getRange(row, 3, 1, head.length - 2).insertCheckboxes();
      l.getRange(row, 1, 1, 2).setValues([[op.date, 0]]);
      rowOf[op.date] = row;
    }
    if ('pomos' in op) {
      l.getRange(row, 2).setValue(Number(op.pomos) || 0);
    } else {
      const col = head.indexOf(String(op.id)) + 1;
      if (col > 2) l.getRange(row, col).setValue(!!op.done);
    }
  });
  if (l.getLastRow() > 2) l.getRange(2, 1, l.getLastRow() - 1, l.getLastColumn()).sort(1);
}

function sessionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let se = ss.getSheetByName(SESSIONS);
  if (!se) { se = ss.insertSheet(SESSIONS); se.getRange(1, 1, 1, SESSION_HEAD.length).setValues([SESSION_HEAD]); }
  return se;
}

// insert or update one row per session id
function applySessions(list) {
  const se = sessionSheet();
  const rowOf = {};
  if (se.getLastRow() > 1) {
    se.getRange(2, 1, se.getLastRow() - 1, 1).getValues().forEach((r, i) => { rowOf[String(r[0])] = i + 2; });
  }
  list.forEach(s => {
    const row = [String(s.id), String(s.date || ''), String(s.start || ''), Number(s.mins) || 0,
                 Number(s.brk) || 0, String(s.crumb || ''), s.finished ? 'yes' : 'stopped early'];
    let r = rowOf[s.id];
    if (!r) { r = se.getLastRow() + 1; rowOf[s.id] = r; }
    se.getRange(r, 2, 1, 2).setNumberFormat('@');
    se.getRange(r, 1, 1, row.length).setValues([row]);
  });
  if (se.getLastRow() > 2) se.getRange(2, 1, se.getLastRow() - 1, SESSION_HEAD.length).sort(3);
}

function readSessions(tz) {
  const se = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS);
  if (!se || se.getLastRow() < 2) return [];
  const n = Math.min(SESSIONS_RETURNED, se.getLastRow() - 1);
  const start = se.getLastRow() - n + 1;
  return se.getRange(start, 1, n, SESSION_HEAD.length).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      id: String(r[0]),
      date: dateKey(r[1], tz),
      start: r[2] instanceof Date ? Utilities.formatDate(r[2], tz, 'yyyy-MM-dd HH:mm') : String(r[2]),
      mins: Number(r[3]) || 0, brk: Number(r[4]) || 0,
      task: '', crumb: String(r[5]), finished: String(r[6]) === 'yes'
    }));
}

function readAll() {
  const {l, h, tz} = sheets();
  const habits = readHabits(h);
  ensureColumns(l, habits);
  const vals = l.getDataRange().getValues();
  const head = vals[0].map(String);
  const log = {}, pomos = {};
  for (let i = 1; i < vals.length; i++) {
    const k = dateKey(vals[i][0], tz);
    if (!k) continue;
    const n = Number(vals[i][1]) || 0;
    if (n) pomos[k] = n;
    const ids = [];
    for (let c = 2; c < head.length; c++) if (vals[i][c] === true) ids.push(head[c]);
    if (ids.length) log[k] = ids;
  }
  return {habits, log, pomos, sessions: readSessions(tz)};
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
