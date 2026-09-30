/**
 * Sales Dashboard Uploader — Google Apps Script backend
 *
 * Receives JSON from https://uploadgooglesheet.netlify.app/ :
 *   { spreadsheetId, sheetName, data: [[row], [row], ...] }
 * Replaces the contents of the target tab and returns { status, rows }.
 *
 * Deploy (as edu@trbapp.com or any account with EDIT access to the sheet):
 *   Deploy > New deployment > Type: Web app
 *   Execute as: Me   |   Who has access: Anyone
 * Then paste the /exec URL into the uploader page (step 1).
 */

// Only these tabs may be written, so the public URL can't touch other tabs.
const ALLOWED_TABS = [
  'Data_uploadD1', 'Data_uploadD2', 'Data_uploadD3', 'Data_uploadD4', 'Data_uploadD5',
  'Data_uploadD6', 'Data_uploadD7', 'Data_uploadD8', 'Data_uploadD9'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(60 * 1000); // one upload at a time
    const body = JSON.parse(e.postData.contents);
    const { spreadsheetId, sheetName, data } = body;

    if (!spreadsheetId || !sheetName || !Array.isArray(data)) {
      return json_({ status: 'error', message: 'Missing spreadsheetId, sheetName or data' });
    }
    if (ALLOWED_TABS.indexOf(sheetName) === -1) {
      return json_({ status: 'error', message: 'Tab not allowed: ' + sheetName });
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // Pad rows so every row has the same number of columns (setValues requires it)
    const width = data.reduce((m, r) => Math.max(m, r.length), 0);
    const rows = data.map(r => r.length === width ? r : r.concat(new Array(width - r.length).fill('')));

    sheet.clearContents();
    if (rows.length && width) {
      sheet.getRange(1, 1, rows.length, width).setValues(rows);
    }
    SpreadsheetApp.flush();

    return json_({ status: 'success', rows: rows.length, tab: sheetName, at: new Date().toISOString() });
  } catch (err) {
    return json_({ status: 'error', message: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// Health check: open the /exec URL in a browser to confirm the deployment works
function doGet() {
  return json_({ status: 'ok', service: 'Sales Dashboard Uploader', account: Session.getEffectiveUser().getEmail() });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
