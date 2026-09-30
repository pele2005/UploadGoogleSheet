/**
 * SIS4 → Google Sheets automatic sync (no browser, no PC needed)
 *
 * Runs inside the same Apps Script project as Code.gs (account edu@trbapp.com).
 * Logs in to SIS4 with a dedicated account and refreshes:
 *   • Sales Dashboard  — 9 reports D1…D9 → Data_uploadD# tabs            (morning)
 *   • Expense          — Expense Tracking YTD → 'allexpense' tab           (morning)
 *   • Daily Sales (DSD)— Report - Actual Sales, this month → 'DATA' tab   (morning + afternoon)
 *
 * Needs: Services → + → "Google Sheets API" (id: Sheets) for the Expense tab (writes RAW like upload2).
 *
 * One-time setup (see README in this folder / chat instructions):
 *   Project Settings → Time zone = (GMT+07:00) Bangkok
 *   Project Settings → Script Properties:
 *       SIS_USER      = the SIS login name of the bot account
 *       SIS_PASS      = its password
 *       NOTIFY_EMAIL  = where to send the result (e.g. kittipong.n@trbchemedica.co.th)
 *   Run checkSis()      → confirms Google can reach SIS and the login works
 *   Run syncNow()       → one full sync, check the log / email
 *   Run installTrigger() → Mon–Fri ~07:45 (all) and ~13:00 (DSD again)
 */

const SIS_BASE = 'https://app.trbchemedica.co.th/sis4/';
const SHEET_ID = '1bQyqKpH7yxafv8Tg3ufVjCOsG8soCrV-PUJc65pCJ28';
const LOGIN_PATHS = ['api/app/login', 'api/sis4/login'];          // first one that answers is used
const MENU_HEADER = {                                              // SIS checks which menu a request comes from
  SalesCompare: '/auth/a/report-param/SalesCompare?jlxpltunaq=pg30gn16',
  SalesMAT: '/auth/a/report-param/SalesMAT?jlxpltunaq=un05hb70',
  SalesActual: '/auth/a/report-param/SalesActual?jlxpltunaq=va28vk72',
  Expense: '/auth/a/expense-tran/detail?jlxpltunaq=vq89jt54'
};
const EXPENSE_SHEET_ID = '1iQ18yGtavcRAlD0Gu3Igr2qpCuFGT4dl4b32lWBTOdY';   // upload2 target
const DSD_SHEET_ID = '1WVwsQMBDA3JPcvGWoCFkr8zZ3dBQC4bWgDMgxpzVi3c';       // Daily Sales Dashboard
const AUTO_NAME = 'Auto (SIS)';                                              // shown as "uploader" in logs

// ---------------- report definitions (same as the Sync Dashboard button) ----------------
const JOBS = [
  { id: 1, report: 'SalesCompare', range: 'YTD',   groupBy: ['Product Line', 'Product Brand', 'Channel', 'Product Group', 'Product Item'] },
  { id: 3, report: 'SalesCompare', range: 'YTD',   groupBy: ['Product Line', 'Product Brand', 'Rep', 'Customer Group', 'Customer'] },
  { id: 2, report: 'SalesCompare', range: 'YTD',   groupBy: ['Product Brand', 'NSM', 'SM', 'Rep', 'Customer'] },
  { id: 4, report: 'SalesCompare', range: 'FULL',  groupBy: ['Product Brand', 'Product Group', 'NSM', 'SM', 'Rep'] },
  { id: 5, report: 'SalesCompare', range: 'HALF',  groupBy: ['Product Brand', 'Product Group', 'NSM', 'SM', 'Rep'] },
  { id: 6, report: 'SalesCompare', range: 'CYCLE', groupBy: ['Product Brand', 'Product Group', 'NSM', 'SM', 'Rep'] },
  { id: 7, report: 'SalesCompare', range: 'FULL',  groupBy: ['NSM', 'Product Brand', 'Product Item', 'Month'] },
  { id: 8, report: 'SalesCompare', range: 'FULL',  groupBy: ['SM', 'Rep', 'Product Brand', 'Product Item', 'Month'] },
  { id: 9, report: 'SalesMAT',     range: 'MAT',   groupBy: ['SM', 'Rep', 'Product Brand', 'Product Item', 'Customer'] }
];

const pad2 = n => String(n).padStart(2, '0');
const ym = (y, m) => `${y}-${pad2(m)}`;

// Reference month = month of "yesterday" (a run on the 1st still covers the month just closed)
function periods_(now) {
  now = now || new Date();
  const ref = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const Y = ref.getFullYear(), M = ref.getMonth() + 1;
  const halfStart = M <= 6 ? 1 : 7;
  const cycleStart = M % 2 === 1 ? M : M - 1;
  return {
    Y, M,
    YTD: [ym(Y, 1), ym(Y, M)],
    FULL: [ym(Y, 1), ym(Y, 12)],
    HALF: [ym(Y, halfStart), ym(Y, M)],
    CYCLE: [ym(Y, cycleStart), ym(Y, cycleStart + 1)],
    MAT: ym(Y, M)
  };
}

const monthIso_ = v => { const [y, m] = v.split('-').map(Number); return new Date(Date.UTC(y, m - 1, 15, 5)).toISOString(); };

function requestBody_(job, P) {
  const nulls = { customer_id: null, cust_channel_id: null, cust_type_id: null, cust_grp_id: null, province_id: null, region_id: null, product_id: null, prod_grp_id: null, org_id: null };
  const value = Object.assign({ REPORT_ID: job.report, GROUP_BY: job.groupBy }, {}, { EXPORT_FORMAT: 'EXCEL' }, nulls);
  const label = Object.assign({ REPORT_ID: 'undefined=' + job.report }, {}, { EXPORT_FORMAT: 'Export Type = EXCEL' }, nulls);
  if (job.range === 'MAT') {
    value.as_of_date = monthIso_(P.MAT);
    label.as_of_date = 'As of Month = ' + P.MAT;
  } else {
    const [a, b] = P[job.range];
    value.date_range = [monthIso_(a), monthIso_(b)];
    label.date_range = `Month Range=${a} - ${b}`;
  }
  return { value, label };
}

// ---------------- SIS login (credentials only from Script Properties) ----------------
function findJwt_(obj) {
  if (typeof obj === 'string') return /^eyJ[\w-]+\.[\w-]+\.[\w-]+$/.test(obj) ? obj : null;
  if (obj && typeof obj === 'object') {
    for (const k of Object.keys(obj)) { const t = findJwt_(obj[k]); if (t) return t; }
  }
  return null;
}

function login_() {
  const props = PropertiesService.getScriptProperties();
  const user = props.getProperty('SIS_USER'), pass = props.getProperty('SIS_PASS');
  if (!user || !pass) throw new Error('Script Properties SIS_USER / SIS_PASS are not set');
  const errors = [];
  for (const path of LOGIN_PATHS) {
    const r = UrlFetchApp.fetch(SIS_BASE + path, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ username: user, password: pass, remember: false })
    });
    const code = r.getResponseCode();
    if (code === 404 || code === 405) { errors.push(`${path}: HTTP ${code}`); continue; }
    let json = null; try { json = JSON.parse(r.getContentText()); } catch (e) {}
    const token = findJwt_(json);
    if (token) return token;
    const msg = json && (json.message || json.error || (json.result && json.result.message));
    throw new Error(`SIS login failed (${path}, HTTP ${code})${msg ? ': ' + msg : ''} — check SIS_USER / SIS_PASS`);
  }
  throw new Error('SIS login endpoint not found: ' + errors.join('; '));
}

function exportReport_(token, job, P) {
  const r = UrlFetchApp.fetch(SIS_BASE + 'api/sis4/ExportToExcel', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, 'X-Rds-Menu': MENU_HEADER[job.report], Accept: 'application/json, text/plain, */*' },
    payload: JSON.stringify(requestBody_(job, P))
  });
  const code = r.getResponseCode();
  if (code !== 200) throw new Error(`Export HTTP ${code}: ${r.getContentText().slice(0, 150)}`);
  return r.getBlob();
}

// ---------------- minimal .xlsx reader: DATA tab → rows (same values as SheetJS header:1, raw) ----------------
function xlsxRows_(blob, sheetName) {
  const files = {};
  Utilities.unzip(blob.setContentType('application/zip')).forEach(f => { files[f.getName()] = f; });
  const text = name => files[name] ? files[name].getDataAsString('UTF-8') : '';
  return parseXlsxParts_(text, sheetName);
}

function decodeXml_(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
          .replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16)))
          .replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(+d)).replace(/&amp;/g, '&');
}
function textOf_(xml) {   // concatenate all <t> runs
  let out = '', m; const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
  while ((m = re.exec(xml))) out += decodeXml_(m[1]);
  return out;
}
function colIndex_(ref) { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }

function parseXlsxParts_(text, sheetName) {
  // workbook: sheet name → r:id → target file
  const wb = text('xl/workbook.xml'), rels = text('xl/_rels/workbook.xml.rels');
  const sheetTag = (wb.match(/<sheet\b[^>]*>/g) || []).find(t => (t.match(/\bname="([^"]*)"/) || [])[1] === sheetName);
  if (!sheetTag) throw new Error(`No "${sheetName}" tab in exported file`);
  const rid = sheetTag.match(/\br:id="([^"]*)"/)[1];
  const relTag = (rels.match(/<Relationship\b[^>]*>/g) || []).find(t => t.indexOf(`Id="${rid}"`) >= 0);
  let target = relTag.match(/\bTarget="([^"]*)"/)[1];
  target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
  const sheetXml = text('xl/' + target);

  const shared = [];
  const sst = text('xl/sharedStrings.xml');
  (sst.match(/<si>[\s\S]*?<\/si>/g) || []).forEach(si => shared.push(textOf_(si)));

  const rows = [];
  let width = 0, firstRow = null;
  // SheetJS sizes the table from <dimension ref="A1:N123">; do the same
  const dim = (sheetXml.match(/<dimension\b[^>]*\bref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"/) || []);
  if (dim[3]) { width = colIndex_(dim[3]) + 1; firstRow = +dim[2]; }
  const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g;
  let rm;
  while ((rm = rowRe.exec(sheetXml))) {
    const attrs = rm[1] || rm[3] || '';
    const rNum = +(attrs.match(/\br="(\d+)"/) || [])[1];
    const body = rm[2] || '';
    const row = [];
    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cm, pos = 0;
    while ((cm = cellRe.exec(body))) {
      const ca = cm[1], inner = cm[2] || '';
      const ref = (ca.match(/\br="([A-Z]+\d+)"/) || [])[1];
      const ci = ref ? colIndex_(ref) : pos;
      pos = ci + 1;
      const t = (ca.match(/\bt="([^"]*)"/) || [])[1] || 'n';
      const v = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
      let val = '';
      if (t === 's') val = v != null ? shared[+v] : '';
      else if (t === 'inlineStr') val = textOf_(inner);
      else if (t === 'str') val = v != null ? decodeXml_(v) : '';
      else if (t === 'b') val = v === '1';
      else if (t === 'e') val = v != null ? decodeXml_(v) : '';
      else val = v != null && v !== '' ? Number(v) : '';
      row[ci] = val;
    }
    if (firstRow === null) firstRow = rNum || 1;
    if (rNum && rNum < firstRow) continue;
    rows[(rNum || rows.length + firstRow) - firstRow] = row;
    width = Math.max(width, row.length);
  }
  // fill gaps with "" (like SheetJS defval: "")
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    const full = new Array(width);
    for (let j = 0; j < width; j++) full[j] = r[j] === undefined ? '' : r[j];
    out.push(full);
  }
  return out;
}

// ---------------- write to the dashboard sheet ----------------
function writeTab_(tabName, rows) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sheet.clearContents();
  if (rows.length && rows[0].length) sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();
  return rows.length;
}

function checkPeriod_(rows, job, P) {
  if (job.range === 'MAT' || rows.length < 2) return;
  const h = rows[0], r = rows[1];
  const iy = h.indexOf('year1'), i1 = h.indexOf('month1'), i2 = h.indexOf('month2');
  if (iy < 0 || i1 < 0 || i2 < 0) return;
  const got = `${r[iy]}-${pad2(+r[i1])}→${r[iy]}-${pad2(+r[i2])}`;
  const want = P[job.range].join('→');
  if (got !== want) throw new Error(`Exported period ${got} ≠ expected ${want}`);
}

// ================= Daily Sales Dashboard (DSD) =================
// Same export the team did by hand: Report - Actual Sales, Group By Rep/Customer/Product Item/Invoice,
// Date Range = 1st of month → end of month. DATA tab → Date, Rep, Customer, Product, Invoice, Qty, Amount, F.O.C, F.O.C2, Borrow
const dayIso_ = d => new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), 5)).toISOString();
const ymd_ = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;

function dsdRange_(now) {
  now = now || new Date();
  const y = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);        // yesterday
  const start = new Date(y.getFullYear(), y.getMonth(), 1);                         // 1st of yesterday's month (keeps last month on the 1st)
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 0);                   // last day of this month
  return [start, end];
}

function syncDsd_(token, opts) {
  const [start, end] = dsdRange_();
  const res = { id: 'DSD', period: `${ymd_(start)}→${ymd_(end)}` };
  try {
    const nulls = { customer_id: null, cust_channel_id: null, cust_type_id: null, cust_grp_id: null, province_id: null, region_id: null, product_id: null, prod_grp_id: null, org_id: null };
    const body = {
      value: Object.assign({ REPORT_ID: 'SalesActual', GROUP_BY: ['Rep', 'Customer', 'Product Item', 'Invoice'], date_range: [dayIso_(start), dayIso_(end)], EXPORT_FORMAT: 'EXCEL' }, nulls),
      label: Object.assign({ REPORT_ID: 'undefined=SalesActual', date_range: `Date Range = ${ymd_(start)} - ${ymd_(end)}`, EXPORT_FORMAT: 'Export Type = EXCEL' }, nulls)
    };
    const raw = xlsxRows_(sisPost_(token, 'api/sis4/ExportToExcel', MENU_HEADER.SalesActual, body), 'DATA');
    const h = raw[0].map(String);
    const col = name => { const i = h.indexOf(name); if (i < 0) throw new Error(`SalesActual: column "${name}" missing`); return i; };
    const iQ = col('qty'), iA = col('amt'), iF = col('foc'), iF2 = col('foc2'), iB = col('borrow');
    const num = v => (v === '' || v == null) ? 0 : Number(v);
    const out = [['Date', 'Rep', 'Customer', 'Product', 'Invoice', 'Qty', 'Amount', 'F.O.C', 'F.O.C2', 'Borrow']];
    raw.slice(1).forEach(r => {
      if (r.every(v => v === '')) return;
      const inv = String(r[3]);
      const m = inv.match(/^(\d{4})-(\d{2})-(\d{2})/);
      out.push([
        m ? new Date(+m[1], +m[2] - 1, +m[3]) : '',
        String(r[0]).replace(/^Area\s+/, ''),
        r[1], r[2], inv,
        num(r[iQ]), num(r[iA]), num(r[iF]), num(r[iF2]), num(r[iB])
      ]);
    });
    if (out.length < 2) throw new Error('SIS returned no sales rows — sheet left unchanged');
    res.rows = out.length - 1;
    if (opts.write) {
      const ss = SpreadsheetApp.openById(DSD_SHEET_ID);
      const sh = ss.getSheetByName('DATA') || ss.insertSheet('DATA');
      sh.clearContents();
      sh.getRange(1, 1, out.length, out[0].length).setValues(out);
      sh.getRange(2, 1, out.length - 1, 1).setNumberFormat('yyyy-mm-dd');
      const log = ss.getSheetByName('UPLOAD_LOG') || ss.insertSheet('UPLOAD_LOG');
      if (log.getLastRow() === 0) log.appendRow(['ผู้อัพโหลด', 'ไฟล์', 'จำนวนรายการ', 'เวลา', 'สถานะ']);
      log.appendRow([AUTO_NAME, `SalesActual ${ymd_(start)}→${ymd_(end)}`, res.rows, new Date(), 'Success']);
      SpreadsheetApp.flush();
    }
    res.ok = true;
  } catch (e) {
    res.ok = false; res.error = String(e.message || e);
  }
  return res;
}

// ================= Expense (same result as upload2.netlify.app) =================
function expenseFormatDate_(raw, fmt) {       // port of upload2 formatDate()
  if (raw === '' || raw == null) return '';
  let d = null;
  if (raw instanceof Date && !isNaN(raw)) d = raw;
  else if (typeof raw === 'number' && raw > 1) {
    const t = new Date((raw - 25569) * 86400000);
    d = new Date(t.getTime() + t.getTimezoneOffset() * 60000);
  } else if (typeof raw === 'string') {
    const p = raw.trim().split(/[\/\-\.]/);
    if (p.length === 3) { let [dd, mm, yy] = p; if (yy.length === 2) yy = '20' + yy; d = new Date(`${yy}-${pad2(mm)}-${pad2(dd)}T00:00:00`); }
  }
  if (d && !isNaN(d)) {
    const dd = pad2(d.getDate()), mm = pad2(d.getMonth() + 1), yy = d.getFullYear();
    return fmt === 'MM/DD/YYYY' ? `${mm}/${dd}/${yy}` : `${dd}/${mm}/${yy}`;
  }
  return '';
}

function expenseRows_(data, uploader) {        // port of upload2 processWorkbook()
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  const ix = n => headers.indexOf(n);
  const hm = {
    date: ix('Date'), month: ix('Month'), year: ix('Year'), team: ix('Team'), costCenter: ix('Cost Center'), type: ix('Type'),
    accountGroup: ix('Account Group'), account: ix('Account'), hospital: ix('Hospital'), doctor: ix('Doctor'), event: ix('Event'),
    request: ix('Request'), requestAmount: ix('Request Amount'), payby: ix('Payby'), payee: ix('Payee'), status: ix('Status'),
    clearingDate: ix('Clearing Date'), clearingAmount: ix('Clearing Amount'), plan: ix('Plan'), createdAt: ix('Created At'),
    updatedBy1: ix('Updated By'), updatedAt1: ix('Updated At')
  };
  hm.updatedBy2 = headers.indexOf('Updated By', hm.updatedBy1 + 1);
  hm.updatedAt2 = headers.indexOf('Updated At', hm.updatedAt1 + 1);
  const v = (row, i) => (i < 0 || row[i] === undefined) ? undefined : row[i];
  return data.slice(1).filter(row => !row.every(c => c === '')).map(row => [
    expenseFormatDate_(v(row, hm.date), 'MM/DD/YYYY'), v(row, hm.month), v(row, hm.year), v(row, hm.team),
    v(row, hm.costCenter), v(row, hm.type), v(row, hm.accountGroup),
    v(row, hm.account), v(row, hm.hospital), '', v(row, hm.doctor),
    v(row, hm.event), '', v(row, hm.request), v(row, hm.requestAmount),
    v(row, hm.payby), v(row, hm.payee), v(row, hm.status),
    expenseFormatDate_(v(row, hm.clearingDate), 'DD/MM/YYYY'), v(row, hm.clearingAmount), v(row, hm.plan),
    uploader, expenseFormatDate_(v(row, hm.createdAt), 'MM/DD/YYYY'),
    hm.updatedBy1 !== -1 ? v(row, hm.updatedBy1) : '', expenseFormatDate_(hm.updatedAt1 !== -1 ? v(row, hm.updatedAt1) : '', 'DD/MM/YYYY'),
    hm.updatedBy2 !== -1 ? v(row, hm.updatedBy2) : '', expenseFormatDate_(hm.updatedAt2 !== -1 ? v(row, hm.updatedAt2) : '', 'DD/MM/YYYY'),
    ''
  ].map(x => x === undefined ? '' : x));
}

function syncExpense_(token, opts) {
  const now = new Date();
  const yest = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const start = new Date(yest.getFullYear(), 0, 1);                                  // 1 Jan (of yesterday's year) → today
  const res = { id: 'EXPENSE', period: `${ymd_(start)}→${ymd_(now)}` };
  try {
    const body = {
      PageIndex: 1, PageSize: 20, Filter: [],
      ExtrasParam: { report_date: [dayIso_(start), dayIso_(now)], eventm_id: null, accm_id: null, location_id: null, owner_id: null, payby_id: null, request_id: null, status_id: null, hospital: null, doctor: null, dsca: null, emp: null, area: null, position: null, data_slot: null, cost_center: null, payee_remark: null },
      ExportDataOnly: true
    };
    const data = xlsxRows_(sisPost_(token, 'api/expense/GetExpTranList', MENU_HEADER.Expense, body), 'Data');
    const rows = expenseRows_(data, AUTO_NAME);
    if (!rows.length) throw new Error('SIS returned no expense rows — sheet left unchanged');
    res.rows = rows.length;
    if (opts.write) {
      if (typeof Sheets === 'undefined') throw new Error('Enable the "Google Sheets API" service in this Apps Script project (Services → +)');
      Sheets.Spreadsheets.Values.clear({}, EXPENSE_SHEET_ID, "'allexpense'!A2:AC");
      Sheets.Spreadsheets.Values.update({ values: rows }, EXPENSE_SHEET_ID, "'allexpense'!A2", { valueInputOption: 'RAW' });
      Sheets.Spreadsheets.Values.append({ values: [[AUTO_NAME, `Expense ${ymd_(start)}→${ymd_(now)}`, now.toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' })]] },
        EXPENSE_SHEET_ID, "'UploadLog'!A1", { valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS' });
      const stamp = `${pad2(now.getDate())}/${pad2(now.getMonth() + 1)}/${String(now.getFullYear()).slice(-2)}`;
      Sheets.Spreadsheets.Values.update({ values: [[stamp]] }, EXPENSE_SHEET_ID, "'allexpense'!AB2", { valueInputOption: 'RAW' });
    }
    res.ok = true;
  } catch (e) {
    res.ok = false; res.error = String(e.message || e);
  }
  return res;
}

// generic SIS POST returning the exported file
function sisPost_(token, path, menu, body) {
  const r = UrlFetchApp.fetch(SIS_BASE + path, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, 'X-Rds-Menu': menu, Accept: 'application/json, text/plain, */*' },
    payload: JSON.stringify(body)
  });
  const code = r.getResponseCode();
  if (code !== 200) throw new Error(`${path} HTTP ${code}: ${r.getContentText().slice(0, 150)}`);
  return r.getBlob();
}

// ---------------- entry points ----------------
// tasks: 'DASH' (D1…D9), 'EXPENSE', 'DSD'
function syncNow() { return runSync_({ write: true }, ['DASH', 'EXPENSE', 'DSD']); }
function dryRun() { return runSync_({ write: false }, ['DASH', 'EXPENSE', 'DSD']); }
function syncDashboardNow() { return runSync_({ write: true }, ['DASH']); }
function syncExpenseNow() { return runSync_({ write: true }, ['EXPENSE']); }
function syncDsdNow() { return runSync_({ write: true }, ['DSD']); }

const isWeekend_ = () => { const d = new Date().getDay(); return d === 0 || d === 6; };   // project time zone must be Asia/Bangkok

function dailySync() {                         // morning: everything
  if (isWeekend_()) return;
  runSync_({ write: true, scheduled: true }, ['DASH', 'EXPENSE', 'DSD']);
}

function dsdAfternoon() {                      // afternoon: DSD only
  if (isWeekend_()) return;
  runSync_({ write: true, scheduled: true }, ['DSD']);
}

function runSync_(opts, tasks) {
  tasks = tasks || ['DASH'];
  const lock = LockService.getScriptLock();
  lock.waitLock(5 * 60 * 1000);               // don't overlap with a web-app upload or another run
  const P = periods_();
  const results = [];
  const plan = [];
  if (tasks.includes('DASH')) JOBS.forEach(job => plan.push(t => runJob_(t, job, P, opts)));
  if (tasks.includes('EXPENSE')) plan.push(t => syncExpense_(t, opts));
  if (tasks.includes('DSD')) plan.push(t => syncDsd_(t, opts));
  let token;
  try {
    token = login_();
    plan.forEach(fn => results.push(fn(token)));
    // one retry for failures (fresh login in case the token expired)
    if (results.some(r => !r.ok)) {
      token = login_();
      results.forEach((r, i) => { if (!r.ok) { const again = plan[i](token); again.retried = true; if (!again.ok) again.error += ` (first try: ${r.error})`; results[i] = again; } });
    }
  } catch (e) {
    results.push({ id: 'LOGIN', ok: false, error: String(e.message || e) });
  } finally {
    lock.releaseLock();
  }
  const ok = results.filter(r => r.ok).length;
  const lines = results.map(r => `${r.ok ? '✅' : '❌'} ${r.id}  ${r.period || ''}  ${r.ok ? r.rows + ' rows' : r.error}${r.retried ? ' (retried)' : ''}`);
  const summary = `SIS sync ${opts.write ? '' : '(DRY RUN) '}— ${ok}/${plan.length} OK — ${tasks.join(' + ')}\n\n` + lines.join('\n');
  console.log(summary);
  notify_(ok === plan.length, summary, opts);
  return summary;
}

function runJob_(token, job, P, opts) {
  const res = { id: 'D' + job.id, period: job.range === 'MAT' ? P.MAT : P[job.range].join('→') };
  try {
    const rows = xlsxRows_(exportReport_(token, job, P), 'DATA');
    checkPeriod_(rows, job, P);
    res.rows = opts.write ? writeTab_('Data_uploadD' + job.id, rows) : rows.length;
    res.ok = true;
  } catch (e) {
    res.ok = false; res.error = String(e.message || e);
  }
  return res;
}

function notify_(allOk, summary, opts) {
  const to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
  if (!to) return;
  if (opts.scheduled && allOk && PropertiesService.getScriptProperties().getProperty('NOTIFY_ON_SUCCESS') !== 'yes') return;
  MailApp.sendEmail(to, (allOk ? '✅' : '❌') + ' SIS auto sync', summary);
}

// Check that Google can reach SIS and the bot login works (no data is written)
function checkSis() {
  const r = UrlFetchApp.fetch(SIS_BASE, { muteHttpExceptions: true });
  console.log('SIS reachable from Google: HTTP ' + r.getResponseCode());
  login_();
  console.log('Login OK ✅');
}

// Mon–Fri ~07:45 (everything) and ~13:00 (DSD again). Apps Script runs within ±15 min of the chosen time.
function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['dailySync', 'dsdAfternoon'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailySync').timeBased().everyDays(1).atHour(7).nearMinute(45).create();
  ScriptApp.newTrigger('dsdAfternoon').timeBased().everyDays(1).atHour(13).nearMinute(0).create();
  console.log('Triggers installed: dailySync ~07:45 (Dashboard + Expense + DSD), dsdAfternoon ~13:00 (DSD). Sat/Sun are skipped.');
}
