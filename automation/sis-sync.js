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
  const VERSION = '2026-09-30.2';
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
    const halfStart = M <= 6 ? 1 : 7;
    const cycleStart = M % 2 === 1 ? M : M - 1;           // 2-month cycles: Jan-Feb, Mar-Apr, ...
    return {
      Y, M,
      YTD: [ym(Y, 1), ym(Y, M)],
      FULL: [ym(Y, 1), ym(Y, 12)],
      HALF: [ym(Y, halfStart), ym(Y, M)],                  // half year to date (e.g. Jul→Sep)
      CYCLE: [ym(Y, cycleStart), ym(Y, cycleStart + 1)],   // whole current cycle (e.g. Sep→Oct)
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

  const isOpen = sel => sel.classList.contains('ant-select-open');

  async function closeSelect(sel) {
    if (!isOpen(sel)) return;
    sel.querySelector('.ant-select-selector').click();
    await sleep(300);
    if (isOpen(sel)) {
      const inp = sel.querySelector('input');
      inp && inp.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true }));
      await sleep(300);
    }
  }

  // close every open select on the page
  async function hideDropdowns() {
    for (const s of document.querySelectorAll('nz-select.ant-select-open')) await closeSelect(s);
  }

  function visibleDropdown() {
    const all = [...document.querySelectorAll('.ant-select-dropdown')].filter(d => getComputedStyle(d).display !== 'none' && !d.classList.contains('ant-select-dropdown-hidden'));
    return all[all.length - 1];
  }

  async function openSelect(sel) {
    await hideDropdowns();
    const s = sel.querySelector('.ant-select-selector');
    s.click();
    await sleep(400);
    if (!isOpen(sel)) { s.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); await sleep(400); }
    if (!isOpen(sel)) throw new Error('Could not open dropdown');
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
    await hideDropdowns(); await sleep(200);
    const got = [...sel.querySelectorAll('.ant-select-selection-item-content')].map(e => e.innerText.trim());
    if (JSON.stringify(got) !== JSON.stringify(labels)) throw new Error(`Group By mismatch: ${got.join(', ')}`);
  }

  // ----- month pickers: click cells in the calendar panel like a person -----
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const openPicker = () => document.querySelector('.ant-picker-dropdown:not(.ant-picker-dropdown-hidden)');

  async function ensurePickerOpen(input) {
    if (openPicker()) return;
    input.click(); await sleep(400);
    if (!openPicker()) { input.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); input.focus(); input.click(); await sleep(400); }
    await waitFor(openPicker, 5000);
  }

  async function panelForYear(y) {
    for (let i = 0; i < 15; i++) {
      const dd = openPicker();
      if (!dd) throw new Error('Month picker closed unexpectedly');
      const panels = [...dd.querySelectorAll('.ant-picker-panel')];
      const p = panels.find(p => (p.querySelector('.ant-picker-header-view')?.innerText || '').trim() === String(y));
      if (p) return p;
      const first = +(panels[0].querySelector('.ant-picker-header-view').innerText.trim());
      dd.querySelector(first > y ? '.ant-picker-header-super-prev-btn' : '.ant-picker-header-super-next-btn').click();
      await sleep(250);
    }
    throw new Error(`Year ${y} not found in picker`);
  }

  async function clickMonth(value) {           // value = 'YYYY-MM'
    const [y, m] = value.split('-').map(Number);
    const p = await panelForYear(y);
    const td = [...p.querySelectorAll('td')].find(td => td.innerText.trim() === MONTHS[m - 1]);
    if (!td) throw new Error(`Month cell ${value} not found`);
    (td.querySelector('.ant-picker-cell-inner') || td).click();
    await sleep(400);
  }

  async function setMonthRange(start, end) {
    const fi = formItem('Month Range');
    const a = fi.querySelector('input[placeholder="Start month"]');
    const b = fi.querySelector('input[placeholder="End month"]');
    if (a.value === start && b.value === end) return;
    await ensurePickerOpen(a);
    await clickMonth(start);
    await clickMonth(end);
    await sleep(300);
    if (a.value !== start || b.value !== end) throw new Error(`Month range not set (${a.value} → ${b.value})`);
  }

  async function setMonth(value) {
    const input = document.querySelector('nz-form-item input[placeholder="Select month"]');
    if (!input) throw new Error('Month input not found');
    if (input.value === value) return;
    await ensurePickerOpen(input);
    await clickMonth(value);
    await sleep(300);
    if (input.value !== value) throw new Error(`Month not set (${input.value})`);
  }

  async function setExportExcel() {
    const sel = formItem('Export Type').querySelector('nz-select');
    if ((sel.innerText || '').includes('EXCEL')) return;
    const dd = await openSelect(sel);
    clickOption(dd, 'EXCEL'); await sleep(200); await hideDropdowns();
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
