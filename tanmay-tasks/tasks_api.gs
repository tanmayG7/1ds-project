/**
 * Tanmay Task Tracker API — Google Apps Script
 * Deploy: Extensions > Apps Script (in the task sheet) > paste this file > Deploy > New deployment
 *   Type: Web app | Execute as: Me | Who has access: Anyone
 * After any code change: Deploy > Manage deployments > edit > new version > Deploy
 * Copy the resulting /exec URL into TASKS_API_URL in index.html.
 *
 * One-time setup for the daily digest email: after deploying, open this project,
 * select createDailyTrigger from the function dropdown at the top, and click Run.
 * That installs a trigger that calls dailyDigest() once a day and will prompt you
 * to authorize the Gmail send scope the first time.
 */

const SHEET_ID       = '1Qy1dVTOfSaMG4rQjKjd5-O1yImipQsDVUpiJ1TqWGuY';
const TASKS_TAB      = 'Sheet1';
const RECIPIENT_EMAIL = 'tanmay@aipply.io';
const DIGEST_HOUR    = 8; // 24-hour clock, in the script's timezone

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'update')          return json(updateTask(p));
    if (p.action === 'create')          return json(createTask(p));
    if (p.action === 'delete')          return json(deleteTask(p));
    if (p.action === 'ping')            return json({ ok: true, msg: 'Tasks API alive' });
    if (p.action === 'send_digest_now') { dailyDigest(); return json({ ok: true, msg: 'Digest sent' }); }
    return json({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json({ ok: false, error: err.toString() });
  }
}

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(TASKS_TAB);
}

// Maps normalized header name (spaces -> underscores, lowercase) to 0-based column index
function headerMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => h.toString().trim());
  const map = {};
  headers.forEach((h, i) => { map[h.toLowerCase().replace(/\s+/g, '_')] = i; });
  return { headers, map };
}

// Ensures a "Completed" column exists (appends it if missing), returns its 0-based index
function completedColIndex(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
  let idx = headers.findIndex(h => h.toLowerCase() === 'completed');
  if (idx === -1) {
    idx = headers.length;
    sheet.getRange(1, idx + 1).setValue('Completed');
  }
  return idx;
}

function findRowById(sheet, map, id) {
  const idCol = map['id'];
  if (idCol === undefined || sheet.getLastRow() < 2) return -1;
  const data = sheet.getRange(2, idCol + 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) return i + 2;
  }
  return -1;
}

function nextId(sheet, map) {
  const idCol = map['id'];
  if (idCol === undefined || sheet.getLastRow() < 2) return 1;
  const values = sheet.getRange(2, idCol + 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0]);
  let maxNum = 0, allNumeric = true;
  values.forEach(v => {
    const n = parseInt(v, 10);
    if (isNaN(n)) allNumeric = false; else if (n > maxNum) maxNum = n;
  });
  return allNumeric ? maxNum + 1 : Date.now();
}

function updateTask(p) {
  if (!p.id) return { ok: false, error: 'id is required' };
  const sheet = getSheet();
  const { map } = headerMap(sheet);
  const row = findRowById(sheet, map, p.id);
  if (row === -1) return { ok: false, error: 'Task not found: ' + p.id };

  const statusIdx = map['status'];
  const prevStatus = statusIdx !== undefined
    ? (sheet.getRange(row, statusIdx + 1).getValue() || '').toString().trim()
    : '';

  ['title', 'area', 'priority', 'status', 'note', 'due_date'].forEach(field => {
    if (p[field] !== undefined && map[field] !== undefined) {
      sheet.getRange(row, map[field] + 1).setValue(p[field]);
    }
  });

  // Stamp/clear Completed only on an actual transition into or out of Done,
  // so re-saving an already-Done task doesn't keep bumping its completed date.
  if (p.status !== undefined) {
    const completedIdx = completedColIndex(sheet);
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (p.status === 'Done' && prevStatus !== 'Done') {
      sheet.getRange(row, completedIdx + 1).setValue(today);
    } else if (p.status !== 'Done' && prevStatus === 'Done') {
      sheet.getRange(row, completedIdx + 1).setValue('');
    }
  }

  return { ok: true, id: p.id };
}

function createTask(p) {
  if (!p.title) return { ok: false, error: 'title is required' };
  const sheet = getSheet();
  const { headers, map } = headerMap(sheet);

  const id = nextId(sheet, map);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const rowArr = new Array(headers.length).fill('');

  if (map['id'] !== undefined) rowArr[map['id']] = id;
  if (map['title'] !== undefined) rowArr[map['title']] = p.title;
  if (map['area'] !== undefined) rowArr[map['area']] = p.area || '';
  if (map['priority'] !== undefined) rowArr[map['priority']] = p.priority || '';
  if (map['status'] !== undefined) rowArr[map['status']] = p.status || 'Pending';
  if (map['note'] !== undefined) rowArr[map['note']] = p.note || '';
  if (map['created'] !== undefined) rowArr[map['created']] = today;
  if (map['due_date'] !== undefined) rowArr[map['due_date']] = p.due_date || '';

  sheet.appendRow(rowArr);
  return { ok: true, id: id };
}

function deleteTask(p) {
  if (!p.id) return { ok: false, error: 'id is required' };
  const sheet = getSheet();
  const { map } = headerMap(sheet);
  const row = findRowById(sheet, map, p.id);
  if (row === -1) return { ok: false, error: 'Task not found: ' + p.id };
  sheet.deleteRow(row);
  return { ok: true, id: p.id };
}

// ── Daily digest ───────────────────────────────────────────────────────────────

function dailyDigest() {
  const sheet = getSheet();
  completedColIndex(sheet); // make sure the Completed column exists before reading
  const { map } = headerMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(yesterday, tz, 'yyyy-MM-dd');

  function fmtCell(v) {
    if (!v) return '';
    if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return v.toString().trim();
  }

  const idIdx = map['id'], titleIdx = map['title'], areaIdx = map['area'], priIdx = map['priority'],
        statusIdx = map['status'], dueIdx = map['due_date'], compIdx = map['completed'];

  const urgentOpen = [], dueToday = [], doneYesterday = [];

  data.forEach(row => {
    const title = titleIdx !== undefined ? (row[titleIdx] || '').toString().trim() : '';
    if (!title) return;
    const area     = areaIdx     !== undefined ? (row[areaIdx]     || '').toString().trim() : '';
    const priority = priIdx      !== undefined ? (row[priIdx]      || '').toString().trim() : '';
    const status   = statusIdx   !== undefined ? (row[statusIdx]  || '').toString().trim() : '';
    const due      = dueIdx      !== undefined ? fmtCell(row[dueIdx])  : '';
    const completed= compIdx     !== undefined ? fmtCell(row[compIdx]) : '';
    const isDone   = status.toLowerCase() === 'done';

    if (priority.toLowerCase() === 'urgent' && !isDone) urgentOpen.push({ title, area, status });
    if (due === todayStr && !isDone)                     dueToday.push({ title, area, priority });
    if (completed === yesterdayStr)                       doneYesterday.push({ title, area });
  });

  const lines = [];
  lines.push('Daily task digest for ' + todayStr);
  lines.push('');
  lines.push('URGENT AND OPEN (' + urgentOpen.length + ')');
  if (urgentOpen.length) {
    urgentOpen.forEach(t => lines.push('- [' + t.area + '] ' + t.title + ' (' + t.status + ')'));
  } else {
    lines.push('None.');
  }
  lines.push('');
  lines.push('DUE TODAY (' + dueToday.length + ')');
  if (dueToday.length) {
    dueToday.forEach(t => lines.push('- [' + t.area + '] ' + t.title + ' (' + t.priority + ')'));
  } else {
    lines.push('None.');
  }
  lines.push('');
  lines.push('DONE YESTERDAY (' + doneYesterday.length + ')');
  if (doneYesterday.length) {
    doneYesterday.forEach(t => lines.push('- [' + t.area + '] ' + t.title));
  } else {
    lines.push('None.');
  }

  const subject = 'Task digest: ' + urgentOpen.length + ' urgent, ' + dueToday.length + ' due today';
  MailApp.sendEmail(RECIPIENT_EMAIL, subject, lines.join('\n'));
}

// Run this once manually (select it in the function dropdown, click Run) to install
// the daily trigger. Safe to re-run — it clears any existing dailyDigest trigger first.
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyDigest')
    .timeBased()
    .everyDays(1)
    .atHour(DIGEST_HOUR)
    .create();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
