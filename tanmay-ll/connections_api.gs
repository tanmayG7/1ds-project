/**
 * Tanmay's Connections Dashboard — Google Apps Script
 * Deploy: open the "Tanmay's LinkedIn Connections" sheet > Extensions > Apps Script >
 *   paste this file > Deploy > New deployment
 *   Type: Web app | Execute as: Me | Who has access: Anyone
 * After any code change: Deploy > Manage deployments > edit > new version > Deploy
 * Copy the resulting /exec URL into CONNECTIONS_API_URL in index.html.
 *
 * One-time data load: after deploying, run (from your machine, not in the script editor):
 *   curl --data-binary @connections_clean.csv "<exec-url>?action=bulk_import"
 * This POSTs the CSV body and replaces the sheet's contents with it.
 */

const SHEET_ID = '1cPrwxQ4-wuwh1WLd3yLlTkFTWNGfAvBPih-yqT00iJg';

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

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

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'update') return json(updateConnection(p));
    if (p.action === 'ping') return json({ ok: true, msg: 'Connections API alive' });
    return json({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json({ ok: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const p = (e.parameter && e.parameter.action) ? e.parameter : { action: 'bulk_import' };
    if (p.action === 'bulk_import') return json(bulkImport(e.postData.contents));
    return json({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json({ ok: false, error: err.toString() });
  }
}

function updateConnection(p) {
  if (!p.id) return { ok: false, error: 'id is required' };
  const sheet = getSheet();
  const { map } = headerMap(sheet);
  const row = findRowById(sheet, map, p.id);
  if (row === -1) return { ok: false, error: 'Connection not found: ' + p.id };

  ['status', 'notes', 'career_site', 'who_to_refer', 'next_steps'].forEach(field => {
    if (p[field] !== undefined && map[field] !== undefined) {
      sheet.getRange(row, map[field] + 1).setValue(p[field]);
    }
  });

  return { ok: true, id: p.id };
}

// Parses the posted CSV text and replaces the whole sheet with it in one batch write.
function bulkImport(csvText) {
  const rows = Utilities.parseCsv(csvText);
  if (!rows.length) return { ok: false, error: 'Empty CSV' };
  const sheet = getSheet();
  sheet.clear();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return { ok: true, rows: rows.length - 1 };
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
