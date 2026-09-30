/**
 * Sales Dashboard — automatic daily sync from SIS4 (no browser, no PC needed)
 *
 * Runs inside the same Apps Script project as Code.gs (account edu@trbapp.com).
 * Logs in to SIS4 with a dedicated read-only account, exports the 9 reports
 * (D1…D9), reads the "DATA" tab of each file and writes it to the matching
 * Data_uploadD# tab of the dashboard sheet.
 *
 * One-time setup (see README in this folder / chat instructions):
 *   Project Settings → Time zone = (GMT+07:00) Bangkok
 *   Project Settings → Script Properties:
 *       SIS_USER      = the SIS login name of the bot account
 *       SIS_PASS      = its password
 *       NOTIFY_EMAIL  = where to send the result (e.g. kittipong.n@trbchemedica.co.th)
 *   Run checkSis()      → confirms Google can reach SIS and the login works
 *   Run syncNow()       → one full sync, check the log / email
 *   Run installTrigger() → Mon–Fri around 07:45 every week
 */

const SIS_BASE = 'https://app.trbchemedica.co.th/sis4/';
const SHEET_ID = '1bQyqKpH7yxafv8Tg3ufVjCOsG8soCrV-PUJc65pCJ28';
const LOGIN_PATHS = ['api/app/login', 'api/sis4/login'];          // first one that answers is used
const MENU_HEADER = {                                              // SIS checks which menu a request comes from
  SalesCompare: '/auth/a/report-param/SalesCompare?jlxpltunaq=pg30gn16',
  SalesMAT: '/auth/a/report-param/SalesMAT?jlxpltunaq=un05hb70'
};

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

// ---------------- entry points ----------------
function syncNow() { return runSync_({ write: true }); }
function dryRun() { return runSync_({ write: false }); }

function dailySync() {
  const d = new Date().getDay();              // project time zone must be Asia/Bangkok
  if (d === 0 || d === 6) return;             // skip Sat/Sun
  runSync_({ write: true, scheduled: true });
}

function runSync_(opts) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5 * 60 * 1000);               // don't overlap with a web-app upload
  const P = periods_();
  const results = [];
  let token;
  try {
    token = login_();
    for (const job of JOBS) results.push(runJob_(token, job, P, opts));
    // one retry for failures (fresh login in case the token expired)
    if (results.some(r => !r.ok)) {
      token = login_();
      results.forEach((r, i) => { if (!r.ok) { const again = runJob_(token, JOBS[i], P, opts); again.retried = true; if (!again.ok) again.error += ` (first try: ${r.error})`; results[i] = again; } });
    }
  } catch (e) {
    results.push({ id: 'LOGIN', ok: false, error: String(e.message || e) });
  } finally {
    lock.releaseLock();
  }
  const ok = results.filter(r => r.ok).length;
  const lines = results.map(r => `${r.ok ? '✅' : '❌'} ${r.id}  ${r.period || ''}  ${r.ok ? r.rows + ' rows' : r.error}${r.retried ? ' (retried)' : ''}`);
  const summary = `Sales Dashboard sync ${opts.write ? '' : '(DRY RUN) '}— ${ok}/${JOBS.length} OK — ref month ${P.Y}-${pad2(P.M)}\n\n` + lines.join('\n');
  console.log(summary);
  notify_(ok === JOBS.length, summary, opts);
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
  MailApp.sendEmail(to, (allOk ? '✅' : '❌') + ' Sales Dashboard sync', summary);
}

// Check that Google can reach SIS and the bot login works (no data is written)
function checkSis() {
  const r = UrlFetchApp.fetch(SIS_BASE, { muteHttpExceptions: true });
  console.log('SIS reachable from Google: HTTP ' + r.getResponseCode());
  login_();
  console.log('Login OK ✅');
}

// Mon–Fri around 07:45 (Apps Script runs within ±15 min of the chosen time)
function installTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailySync').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailySync').timeBased().everyDays(1).atHour(7).nearMinute(45).create();
  console.log('Daily trigger installed (Sat/Sun are skipped inside dailySync).');
}
