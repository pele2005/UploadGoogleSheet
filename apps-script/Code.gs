/**
 * Sales Dashboard Uploader — Google Apps Script backend
 *
 * Receives JSON from https://uploadgooglesheet.netlify.app/ (and the "Sync Dashboard" bookmark):
 *   { spreadsheetId, sheetName, data: [[row], [row], ...] }
 * Replaces the contents of the target tab and returns { status, rows }.
 *
 * NEW (2026-10-02): Total Year Progression (tab "TYP") is now refreshed here, automatically,
 * every time Data_uploadD4 (Sales Compare, Jan–Dec, by Brand) is uploaded. It replaces the
 * separate "Total year progression" scheduled task — no extra SIS export is needed, because
 * D4 already holds full-year Actual (cy_actual) and Plan (cy_target) for every brand.
 *
 * Deploy (as edu@trbapp.com or any account with EDIT access to the sheet):
 *   Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy
 *   (Execute as: Me | Who has access: Anyone)  — the /exec URL stays the same.
 */

// Only these tabs may be written, so the public URL can't touch other tabs.
const ALLOWED_TABS = [
  'Data_uploadD1', 'Data_uploadD2', 'Data_uploadD3', 'Data_uploadD4', 'Data_uploadD5',
  'Data_uploadD6', 'Data_uploadD7', 'Data_uploadD8', 'Data_uploadD9'
];

// ---------------- Total Year Progression (tab TYP) ----------------
const TYP_SPREADSHEET_ID = '1bQyqKpH7yxafv8Tg3ufVjCOsG8soCrV-PUJc65pCJ28';
const TYP_TAB = 'TYP';
const TYP_SOURCE_TAB = 'Data_uploadD4';   // G1 = Product Brand, full year (month1 = 1, month2 = 12)

// Label in TYP column A  ->  SIS Product Brand names (column G1 of Data_uploadD4) that add up to it.
// brands: null = every row (Grand Total).
// Note: SIS reports the discontinued Blephasteam items under the brand name "OPHTHAL".
const TYP_ROWS = [
  { label: 'Grand Total',                     brands: null },
  { label: 'Total Product Line -OPHTHAL',     brands: ['OPHTHAL', 'Blephasteam', 'Ocusoft', 'VISIOL', 'VISLUBE'] },
  { label: 'Total Product Brand-Blephasteam', brands: ['OPHTHAL', 'Blephasteam'] },
  { label: 'Total Product Brand-Ocusoft',     brands: ['Ocusoft'] },
  { label: 'Total Product Brand-Visiol',      brands: ['VISIOL'] },
  { label: 'Total Product Brand-Vislube',     brands: ['VISLUBE'] },
  { label: 'Total Product Line -ORTHO',       brands: ['ARTRODAR', 'Ostenil'] },
  { label: 'Total Product Brand-ARTRODAR',    brands: ['ARTRODAR'] },
  { label: 'Total Product Brand-Ostenil',     brands: ['Ostenil'] }
];

// Which of the rows above are actually written to the TYP tab.
// Only Grand Total for now: several Looker Studio dashboards sum the whole TYP column, so filling the
// Line/Brand rows would triple-count there. To fill every row later, set this to null.
const TYP_WRITE_LABELS = ['Grand Total'];

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

    const out = { status: 'success', rows: rows.length, tab: sheetName, at: new Date().toISOString() };

    // Refresh TYP right after the full-year file lands. A TYP problem never fails the upload itself.
    if (sheetName === TYP_SOURCE_TAB) {
      try { out.typ = updateTYP_(ss, rows); }
      catch (err) { out.typ = { status: 'error', message: String(err && err.message || err) }; }
    }
    return json_(out);
  } catch (err) {
    return json_({ status: 'error', message: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/**
 * Run this by hand from the Apps Script editor (Run > refreshTYP) to rebuild TYP
 * from whatever is in Data_uploadD4 right now. Check View > Logs for the result.
 */
function refreshTYP() {
  const ss = SpreadsheetApp.openById(TYP_SPREADSHEET_ID);
  const src = ss.getSheetByName(TYP_SOURCE_TAB);
  if (!src) throw new Error('Tab not found: ' + TYP_SOURCE_TAB);
  const result = updateTYP_(ss, src.getDataRange().getValues());
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Sum Actual / Plan by brand from the D4 rows and write TYP!B:D next to the matching labels in column A.
function updateTYP_(ss, rows) {
  const totals = computeTYP_(rows);

  const sheet = ss.getSheetByName(TYP_TAB);
  if (!sheet) throw new Error('Tab not found: ' + TYP_TAB);
  const last = Math.max(sheet.getLastRow(), 2);
  const labels = sheet.getRange(1, 1, last, 1).getValues().map(r => normLabel_(r[0]));

  const missing = [];
  const toWrite = totals.lines.filter(t => !TYP_WRITE_LABELS || TYP_WRITE_LABELS.indexOf(t.label) !== -1);
  toWrite.forEach(t => {
    const idx = labels.indexOf(normLabel_(t.label));
    if (idx === -1) { missing.push(t.label); return; }
    const row = idx + 1;
    sheet.getRange(row, 2, 1, 2).setValues([[t.actual, t.plan]]).setNumberFormat('#,##0.00');
    sheet.getRange(row, 4).setFormula('=IFERROR(B' + row + '/C' + row + ',0)').setNumberFormat('0.00%');
  });
  SpreadsheetApp.flush();

  return {
    status: missing.length || (!TYP_WRITE_LABELS && totals.unmapped.length) ? 'warning' : 'success',
    year: totals.year,
    grandActual: totals.lines[0].actual,
    grandPlan: totals.lines[0].plan,
    rowsWritten: toWrite.length - missing.length,
    missingLabels: missing,          // labels not found in TYP column A (nothing written for these)
    unmappedBrands: totals.unmapped  // brands in SIS that are in Grand Total but in no line/brand row
  };
}

// Pure calculation (no sheet access) — rows = Data_uploadD4 including its header row.
function computeTYP_(rows) {
  if (!rows || rows.length < 2) throw new Error(TYP_SOURCE_TAB + ' is empty');
  const h = rows[0].map(c => String(c).trim());
  const col = name => {
    const i = h.indexOf(name);
    if (i === -1) throw new Error('Column "' + name + '" not found in ' + TYP_SOURCE_TAB);
    return i;
  };
  const iBrand = col('G1'), iAct = col('cy_actual'), iPlan = col('cy_target');
  const iY = col('year1'), iM1 = col('month1'), iM2 = col('month2');

  const byBrand = {};
  let year = null;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const brand = String(row[iBrand]).trim();
    if (!brand && row[iAct] === '' && row[iPlan] === '') continue;   // blank trailing row
    // TYP is a whole-year figure: refuse anything that is not Jan–Dec so a wrong export can't corrupt it.
    if (Number(row[iM1]) !== 1 || Number(row[iM2]) !== 12) {
      throw new Error(TYP_SOURCE_TAB + ' is not a Jan–Dec export (month ' + row[iM1] + '–' + row[iM2] + '); TYP left unchanged');
    }
    if (year === null) year = Number(row[iY]);
    const key = brand.toUpperCase();
    if (!byBrand[key]) byBrand[key] = { name: brand, actual: 0, plan: 0 };
    byBrand[key].actual += num_(row[iAct]);
    byBrand[key].plan += num_(row[iPlan]);
  }

  const mapped = {};
  const lines = TYP_ROWS.map(def => {
    let actual = 0, plan = 0;
    const keys = def.brands ? def.brands.map(b => b.toUpperCase()) : Object.keys(byBrand);
    keys.forEach(k => {
      if (!byBrand[k]) return;
      actual += byBrand[k].actual; plan += byBrand[k].plan;
      if (def.brands) mapped[k] = true;
    });
    return { label: def.label, actual: round2_(actual), plan: round2_(plan) };
  });
  const unmapped = Object.keys(byBrand).filter(k => !mapped[k]).map(k => byBrand[k].name);
  return { year: year, lines: lines, unmapped: unmapped };
}

function num_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}
function round2_(n) { return Math.round(n * 100) / 100; }
function normLabel_(s) { return String(s).toLowerCase().replace(/\s+/g, ''); }

// Health check: open the /exec URL in a browser to confirm the deployment works
function doGet() {
  return json_({ status: 'ok', service: 'Sales Dashboard Uploader', typ: true, account: Session.getEffectiveUser().getEmail() });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
