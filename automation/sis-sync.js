/*
 * SIS → Google Sheet sync  (runs inside a logged-in SIS4 tab)
 *
 * Drives the SIS4 report forms exactly like a person would (Group By, Month
 * Range, Export Type = EXCEL, click Export), catches the exported .xlsx in
 * memory instead of downloading it, reads the "DATA" tab with SheetJS and sends
 * the rows to the Google Apps Script web app — the same payload the uploader
 * page (uploadgooglesheet.netlify.app) sends.
 *
 * Usage (in the SIS tab, after login):
 *   await SISSync.run({ appsScriptUrl: 'https://script.google.com/macros/s/.../exec' })
 *   await SISSync.run({ appsScriptUrl, only: [1,3], dryRun: true })   // test without writing
 *
 * No credentials are read or stored: every export request is made by SIS itself.
 */
(function () {
  const VERSION = '2026-09-30.1';
  const SPREADSHEET_ID = '1bQyqKpH7yxafv8Tg3ufVjCOsG8soCrV-PUJc65pCJ28';
  const XLSX_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const pad = n => String(n).padStart(2, '0');
  const ym = (y, m) => `${y}-${pad(m)}`;

  // ---------------- period logic ----------------
  // Reference month = month of "yesterday" so a run on the 1st still covers the month just closed.
  function periods(now = new Date()) {
    const ref = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
    const Y = ref.getFullYear(), M = ref.getMonth() + 1;
    const half = M <= 6 ? [1, 6] : [7, 12];
    const cycleStart = M % 2 === 1 ? M : M - 1;           // 2-month cycles: Jan-Feb, Mar-Apr, ...
    return {
      Y, M,
      YTD: [ym(Y, 1), ym(Y, M)],
      FULL: [ym(Y, 1), ym(Y, 12)],
      HALF: [ym(Y, half[0]), ym(Y, half[1])],
      CYCLE: [ym(Y, cycleStart), ym(Y, M)],
      MAT: ym(Y, M)
    };
  }

  // ---------------- report definitions (D1…D9) ----------------
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
  const MENU = { SalesCompare: 'Report - Sales Compar', SalesMAT: 'Report - Moving Annu' };

  // ---------------- helpers ----------------
  async function waitFor(fn, timeout = 15000, step = 200) {
    const t0 = Date.now();
    while (Date.now() - t0 < timeout) { const v = fn(); if (v) return v; await sleep(step); }
    throw new Error('Timeout waiting for page element');
  }

  function loadXLSX() {
    if (window.XLSX) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = XLSX_CDN; s.onload = res; s.onerror = () => rej(new Error('Cannot load SheetJS'));
      document.head.appendChild(s);
    });
  }

  function formItem(label) {
    return [...document.querySelectorAll('nz-form-item')]
      .find(fi => (fi.querySelector('nz-form-label, label')?.innerText || '').replace(/[*:]/g, '').trim() === label);
  }

  function hideDropdowns() {
    document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    document.querySelectorAll('.ant-select-dropdown').forEach(d => { d.classList.add('ant-select-dropdown-hidden'); d.style.display = 'none'; });
  }

  function visibleDropdown() {
    return [...document.querySelectorAll('.ant-select-dropdown')].find(d => getComputedStyle(d).display !== 'none' && !d.classList.contains('ant-select-dropdown-hidden'));
  }

  async function openSelect(sel) {
    hideDropdowns(); await sleep(150);
    const s = sel.querySelector('.ant-select-selector');
    s.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    s.click();
    return waitFor(visibleDropdown, 5000);
  }

  function clickOption(dd, text) {
    const opt = [...dd.querySelectorAll('nz-option-item, .ant-select-item-option')].find(o => o.innerText.trim() === text);
    if (!opt) throw new Error(`Option not found: ${text}`);
    opt.click();
  }

  // ---------------- navigation ----------------
  async function openReport(report) {
    if (location.hash.includes(`/report-param/${report}`) && formItem('Group By')) return;
    const label = MENU[report];
    const find = () => [...document.querySelectorAll('li, a, span')].find(e => e.children.length === 0 && e.innerText && e.innerText.trim().startsWith(label));
    let link = find();
    if (!link) {   // expand "Report Viewer" menu
      const parent = [...document.querySelectorAll('li, div, span')].find(e => e.children.length <= 3 && e.innerText && e.innerText.trim() === 'Report Viewer');
      if (parent) parent.click();
      link = await waitFor(find, 5000);
    }
    link.click();
    await waitFor(() => location.hash.includes(`/report-param/${report}`) && formItem('Group By'), 15000);
    await sleep(800);
  }

  // ---------------- form filling ----------------
  async function setGroupBy(labels) {
    const sel = formItem('Group By').querySelector('nz-select');
    // remove current tags
    for (let i = 0; i < 10; i++) {
      const rm = sel.querySelector('.ant-select-selection-item-remove');
      if (!rm) break;
      rm.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); rm.click(); await sleep(150);
    }
    const dd = await openSelect(sel);
    for (const l of labels) { clickOption(dd, l); await sleep(200); }
    hideDropdowns(); await sleep(200);
    const got = [...sel.querySelectorAll('.ant-select-selection-item-content')].map(e => e.innerText.trim());
    if (JSON.stringify(got) !== JSON.stringify(labels)) throw new Error(`Group By mismatch: ${got.join(', ')}`);
  }

  function setInput(input, value) {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    input.focus();
    setter.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    input.blur();
  }

  async function setMonthRange(start, end) {
    const fi = formItem('Month Range');
    const a = fi.querySelector('input[placeholder="Start month"]');
    const b = fi.querySelector('input[placeholder="End month"]');
    setInput(a, start); await sleep(250);
    setInput(b, end); await sleep(250);
    hideDropdowns(); await sleep(200);
    if (a.value !== start || b.value !== end) throw new Error(`Month range not set (${a.value} → ${b.value})`);
  }

  async function setMonth(value) {
    const fi = formItem('Month') || formItem('As of Month') || formItem('Select Month');
    const input = (fi || document).querySelector('input[placeholder="Select month"]');
    if (!input) throw new Error('Month input not found');
    setInput(input, value); await sleep(250); hideDropdowns(); await sleep(200);
    if (input.value !== value) throw new Error(`Month not set (${input.value})`);
  }

  async function setExportExcel() {
    const sel = formItem('Export Type').querySelector('nz-select');
    if ((sel.innerText || '').includes('EXCEL')) return;
    const dd = await openSelect(sel);
    clickOption(dd, 'EXCEL'); await sleep(200); hideDropdowns();
  }

  // ---------------- capture export blob instead of downloading ----------------
  let pending = null;
  function installCapture() {
    if (window.__sisSyncCapture) return;
    window.__sisSyncCapture = true;
    const origCreate = URL.createObjectURL.bind(URL);
    URL.createObjectURL = function (obj) {
      const url = origCreate(obj);
      if (pending && obj instanceof Blob && obj.size > 0) { pending.resolve(obj); pending = null; window.__sisSyncSuppress = url; }
      return url;
    };
    const origClick = HTMLAnchorElement.prototype.click;
    HTMLAnchorElement.prototype.click = function () {
      if (window.__sisSyncSuppress && this.href === window.__sisSyncSuppress) { window.__sisSyncSuppress = null; return; }
      return origClick.apply(this, arguments);
    };
  }

  async function exportAndCapture(timeoutMs = 180000) {
    const blobP = new Promise((resolve, reject) => {
      pending = { resolve };
      setTimeout(() => { if (pending) { pending = null; reject(new Error('Export timed out')); } }, timeoutMs);
    });
    const btn = [...document.querySelectorAll('button')].find(b => b.innerText.trim() === 'Export');
    if (!btn) throw new Error('Export button not found');
    btn.click();
    return blobP;
  }

  // ---------------- upload ----------------
  async function blobToRows(blob) {
    const wb = XLSX.read(await blob.arrayBuffer(), { type: 'array' });
    if (!wb.SheetNames.includes('DATA')) throw new Error('No "DATA" tab in exported file');
    return XLSX.utils.sheet_to_json(wb.Sheets.DATA, { header: 1, defval: '' });
  }

  async function postRows(appsScriptUrl, sheetName, data) {
    const r = await fetch(appsScriptUrl, {
      method: 'POST', mode: 'cors', cache: 'no-cache', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ spreadsheetId: SPREADSHEET_ID, sheetName, data })
    });
    const text = await r.text();
    let j; try { j = JSON.parse(text); } catch (e) { throw new Error('Apps Script did not return JSON: ' + text.slice(0, 120)); }
    if (j.status !== 'success') throw new Error(j.message || 'Apps Script error');
    return j.rows;
  }

  // ---------------- main ----------------
  async function run(opts = {}) {
    const { appsScriptUrl, only, dryRun = false, now } = opts;
    if (!appsScriptUrl && !dryRun) throw new Error('appsScriptUrl is required');
    if (location.hash.includes('/login')) throw new Error('NOT_LOGGED_IN: please log in to SIS first');
    await loadXLSX();
    installCapture();
    const P = periods(now ? new Date(now) : new Date());
    const results = [];
    for (const job of JOBS) {
      if (only && !only.includes(job.id)) continue;
      const res = { id: `D${job.id}`, tab: `Data_uploadD${job.id}`, period: job.range === 'MAT' ? P.MAT : P[job.range].join('→') };
      const t0 = Date.now();
      try {
        await openReport(job.report);
        await setGroupBy(job.groupBy);
        if (job.range === 'MAT') await setMonth(P.MAT); else await setMonthRange(...P[job.range]);
        await setExportExcel();
        const blob = await exportAndCapture();
        const rows = await blobToRows(blob);
        res.rows = rows.length;
        res.header = rows[0];
        if (!dryRun) res.written = await postRows(appsScriptUrl, res.tab, rows);
        res.ok = true;
      } catch (e) {
        res.ok = false; res.error = String(e.message || e);
      }
      res.seconds = Math.round((Date.now() - t0) / 1000);
      results.push(res);
      console.log('[SISSync]', JSON.stringify(res));
      await sleep(1500);
    }
    const summary = { version: VERSION, ref: `${P.Y}-${pad(P.M)}`, dryRun, ok: results.filter(r => r.ok).length, total: results.length, results };
    window.__sisSyncLast = summary;
    return summary;
  }

  window.SISSync = { run, periods, JOBS, VERSION };
})();
