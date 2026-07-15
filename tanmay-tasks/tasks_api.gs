/**
 * Tanmay Task Tracker API — Google Apps Script
 * Deploy: Extensions > Apps Script (in the task sheet) > paste this file > Deploy > New deployment
 *   Type: Web app | Execute as: Me | Who has access: Anyone
 * After any code change: Deploy > Manage deployments > edit > new version > Deploy
 * Copy the resulting /exec URL into TASKS_API_URL in index.html.
 */

const SHEET_ID  = '1Qy1dVTOfSaMG4rQjKjd5-O1yImipQsDVUpiJ1TqWGuY';
const TASKS_TAB = 'Sheet1';

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'update') return json(updateTask(p));
    if (p.action === 'create') return json(createTask(p));
    if (p.action === 'delete') return json(deleteTask(p));
    if (p.action === 'ping')   return json({ ok: true, msg: 'Tasks API alive' });
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

  ['title', 'area', 'priority', 'status', 'note', 'due_date'].forEach(field => {
    if (p[field] !== undefined && map[field] !== undefined) {
      sheet.getRange(row, map[field] + 1).setValue(p[field]);
    }
  });
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

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
