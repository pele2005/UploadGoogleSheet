/*
 * SIS → Google Sheet sync  (runs inside a logged-in SIS4 tab)
 *
 * For each report D1…D9 it clicks SIS's own "Export" button. Just before SIS
 * sends the export request, the report parameters in the request body
 * (Group By, month range / as-of month) are replaced with the ones for that
 * report. SIS adds its own login to the request — this script never reads or
 * stores any credential. The exported .xlsx is caught in memory (no download),
 * its "DATA" tab is read with SheetJS and the rows are sent to the Google Apps
 * Script web app — the same payload the uploader page sends.
 *
 * Works in a background tab: waits are event-driven (DOM / network), not
 * timer polling, so Chrome's background-tab throttling doesn't stall it.
 *
 * Usage (in the SIS tab, after login):
 *   SISSync.start({ appsScriptUrl: 'https://script.google.com/macros/s/.../exec' })  // then poll SISSync.status()
 *   await SISSync.run({ dryRun: true, only: [1, 9] })                               // test without writing
 */
(function () {
  const VERSION = '2026-09-30.10';
  const SPREADSHEET_ID = '1bQyqKpH7yxafv8Tg3ufVjCOsG8soCrV-PUJc65pCJ28';
  const XLSX_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  const pad = n => String(n).padStart(2, '0');
  const ym = (y, m) => `${y}-${pad(m)}`;

  // ---------------- period logic ----------------
  // Reference month = month of "yesterday", so a run on the 1st still covers the month just closed.
  function periods(now = new Date()) {
    const ref = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
    const Y = ref.getFullYear(), M = ref.getMonth() + 1;
    const halfStart = M <= 6 ? 1 : 7;
    const cycleStart = M % 2 === 1 ? M : M - 1;             // 2-month cycles: Jan-Feb, Mar-Apr, ...
    return {
      Y, M,
      YTD: [ym(Y, 1), ym(Y, M)],
      FULL: [ym(Y, 1), ym(Y, 12)],
      HALF: [ym(Y, halfStart), ym(Y, M)],                    // half year to date (e.g. Jul→Sep)
      CYCLE: [ym(Y, cycleStart), ym(Y, cycleStart + 1)],     // whole current cycle (e.g. Sep→Oct)
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
  const PERIOD_FIELD = { SalesCompare: 'Month Range', SalesMAT: 'As of Month' };   // proves the right form is on screen

  // mid-month, midday local time → the ISO date always falls in the intended month
  const monthIso = v => { const [y, m] = v.split('-').map(Number); return new Date(y, m - 1, 15, 12).toISOString(); };

  function overrideFor(job, P) {
    if (job.range === 'MAT') {
      return { value: { GROUP_BY: job.groupBy, as_of_date: monthIso(P.MAT) }, label: { as_of_date: `As of Month = ${P.MAT}` } };
    }
    const [a, b] = P[job.range];
    return { value: { GROUP_BY: job.groupBy, date_range: [monthIso(a), monthIso(b)] }, label: { date_range: `Month Range=${a} - ${b}` } };
  }

  // ---------------- event-driven waiting (not throttled in background tabs) ----------------
  function waitFor(fn, timeout = 20000) {
    return new Promise((resolve, reject) => {
      const v0 = fn(); if (v0) return resolve(v0);
      const obs = new MutationObserver(() => { const v = fn(); if (v) { obs.disconnect(); clearTimeout(t); resolve(v); } });
      obs.observe(document.documentElement, { childList: true, subtree: true, attributes: true, characterData: true });
      const t = setTimeout(() => { obs.disconnect(); const v = fn(); v ? resolve(v) : reject(new Error('Timeout waiting for page element')); }, timeout);
    });
  }
  const nextFrame = () => new Promise(r => setTimeout(r, 50));

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

  // ---------------- navigation & the one form field we still set ----------------
  async function openReport(report) {
    const ready = () => location.hash.includes(`/report-param/${report}`) && formItem(PERIOD_FIELD[report]) && formItem('Export Type');
    if (ready()) return;
    const label = MENU[report];
    const find = () => [...document.querySelectorAll('li, a, span')].find(e => e.children.length === 0 && e.innerText && e.innerText.trim().startsWith(label));
    if (!find()) {
      const parent = [...document.querySelectorAll('li, div, span')].find(e => e.children.length <= 3 && e.innerText && e.innerText.trim() === 'Report Viewer');
      if (parent) parent.click();
    }
    (await waitFor(find, 10000)).click();
    await waitFor(ready, 20000);
  }

  async function setExportExcel() {
    const sel = formItem('Export Type').querySelector('nz-select');
    for (let attempt = 0; attempt < 3 && !(sel.innerText || '').includes('EXCEL'); attempt++) {
      sel.querySelector('.ant-select-selector').click();
      let opt;
      try {
        opt = await waitFor(() => [...document.querySelectorAll('.ant-select-dropdown nz-option-item, .ant-select-dropdown .ant-select-item-option')].find(o => o.innerText.trim() === 'EXCEL'), 5000);
      } catch (e) { sel.querySelector('.ant-select-selector').dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); continue; }
      opt.click();
      await waitFor(() => (sel.innerText || '').includes('EXCEL'), 3000).catch(() => {});
    }
    if (!(sel.innerText || '').includes('EXCEL')) throw new Error('Could not set Export Type = EXCEL');
  }

  // ---------------- hooks: replace report parameters in SIS's export request; catch the file ----------------
  function installHooks() {
    if (window.__sisSyncHooks === VERSION) return;
    window.__sisSyncHooks = VERSION;
    if (!window.__sisSyncXhrPatched) {
      window.__sisSyncXhrPatched = true;
      const origOpen = XMLHttpRequest.prototype.open, origSend = XMLHttpRequest.prototype.send;
      XMLHttpRequest.prototype.open = function (m, u) { this.__sisUrl = String(u); return origOpen.apply(this, arguments); };
      XMLHttpRequest.prototype.send = function (body) {
        const ov = window.__sisSyncOverride;
        if (ov && this.__sisUrl && this.__sisUrl.includes('ExportToExcel') && typeof body === 'string') {
          try {
            const b = JSON.parse(body);
            Object.assign(b.value, ov.value);
            if (b.label) Object.assign(b.label, ov.label);
            delete b.value[ov.value.as_of_date ? 'date_range' : 'as_of_date'];
            body = JSON.stringify(b);
            window.__sisSyncOverride = null;
            window.__sisSyncSent = b.value;
          } catch (e) { /* leave body unchanged */ }
        }
        return origSend.call(this, body);
      };
    }
    if (!window.__sisSyncBlobPatched) {
      window.__sisSyncBlobPatched = true;
      const origCreate = URL.createObjectURL.bind(URL);
      URL.createObjectURL = function (obj) {
        const url = origCreate(obj);
        const p = window.__sisSyncPending;
        if (p && obj instanceof Blob && obj.size > 0) { window.__sisSyncPending = null; window.__sisSyncSuppress = url; p(obj); }
        return url;
      };
      const origClick = HTMLAnchorElement.prototype.click;
      HTMLAnchorElement.prototype.click = function () {
        if (window.__sisSyncSuppress && this.href === window.__sisSyncSuppress) { window.__sisSyncSuppress = null; return; }
        return origClick.apply(this, arguments);
      };
    }
  }

  async function exportWith(override, timeoutMs = 240000) {
    window.__sisSyncSent = null;
    window.__sisSyncOverride = override;
    const blobP = new Promise((resolve, reject) => {
      window.__sisSyncPending = resolve;
      setTimeout(() => { if (window.__sisSyncPending === resolve) { window.__sisSyncPending = null; window.__sisSyncOverride = null; reject(new Error('Export timed out')); } }, timeoutMs);
    });
    const btn = [...document.querySelectorAll('button')].find(b => b.innerText.trim() === 'Export');
    if (!btn) throw new Error('Export button not found');
    btn.click();
    // SIS sends nothing if its form is invalid — fail fast instead of waiting for the full timeout
    const sent = await new Promise(res => {
      const t0 = Date.now();
      const tick = () => { if (window.__sisSyncSent) return res(true); if (Date.now() - t0 > 15000) return res(false); setTimeout(tick, 250); };
      tick();
    });
    if (!sent) {
      window.__sisSyncPending = null; window.__sisSyncOverride = null;
      const bad = [...document.querySelectorAll('.ant-form-item-has-error nz-form-label, .ant-form-item-has-error label')].map(e => e.innerText.trim()).join(', ');
      throw new Error('SIS did not send the export request' + (bad ? ` (form errors: ${bad})` : ''));
    }
    return blobP;
  }

  // ---------------- file → rows → Apps Script ----------------
  async function blobToRows(blob) {
    const wb = XLSX.read(await blob.arrayBuffer(), { type: 'array' });
    if (!wb.SheetNames.includes('DATA')) throw new Error('No "DATA" tab in exported file');
    return XLSX.utils.sheet_to_json(wb.Sheets.DATA, { header: 1, defval: '' });
  }

  // SalesCompare files carry year1/month1/month2 columns — confirm they match what we asked for
  function checkPeriod(rows, job, P) {
    if (job.range === 'MAT' || rows.length < 2) return;
    const h = rows[0], r = rows[1];
    const iy = h.indexOf('year1'), i1 = h.indexOf('month1'), i2 = h.indexOf('month2');
    if (iy < 0 || i1 < 0 || i2 < 0) return;
    const got = `${r[iy]}-${pad(+r[i1])}→${r[iy]}-${pad(+r[i2])}`;
    const want = P[job.range].join('→');
    if (got !== want) throw new Error(`Exported period ${got} ≠ expected ${want}`);
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
  async function runJob(job, P, opts, state) {
    const res = { id: `D${job.id}`, tab: `Data_uploadD${job.id}`, period: job.range === 'MAT' ? P.MAT : P[job.range].join('→') };
    const t0 = Date.now();
    state.current = res.id;
    try {
      await openReport(job.report);
      await setExportExcel();
      const blob = await exportWith(overrideFor(job, P));
      const rows = await blobToRows(blob);
      checkPeriod(rows, job, P);
      res.rows = rows.length;
      if (!opts.dryRun) res.written = await postRows(opts.appsScriptUrl, res.tab, rows);
      res.ok = true;
    } catch (e) {
      res.ok = false; res.error = String(e.message || e);
    }
    res.seconds = Math.round((Date.now() - t0) / 1000);
    console.log('[SISSync]', JSON.stringify(res));
    return res;
  }

  async function run(opts = {}) {
    const { appsScriptUrl, only, dryRun = false, now } = opts;
    if (!appsScriptUrl && !dryRun) throw new Error('appsScriptUrl is required');
    if (location.hash.includes('/login')) throw new Error('NOT_LOGGED_IN: please log in to SIS first');
    await loadXLSX();
    installHooks();
    const P = periods(now ? new Date(now) : new Date());
    const results = [];
    const state = window.__sisSyncState = { running: true, version: VERSION, dryRun, startedAt: new Date().toISOString(), current: null, results };
    const jobs = JOBS.filter(j => !only || only.includes(j.id));
    for (const job of jobs) results.push(await runJob(job, P, opts, state));
    // one retry for anything that failed
    for (let i = 0; i < results.length; i++) {
      if (results[i].ok) continue;
      const job = jobs[i];
      const again = await runJob(job, P, opts, state);
      again.retried = true;
      if (!again.ok) again.error = `${again.error} (first try: ${results[i].error})`;
      results[i] = again;
    }
    const summary = { version: VERSION, ref: `${P.Y}-${pad(P.M)}`, dryRun, ok: results.filter(r => r.ok).length, total: results.length, results };
    Object.assign(state, { running: false, current: null, finishedAt: new Date().toISOString(), summary });
    window.__sisSyncLast = summary;
    return summary;
  }

  // Fire-and-forget for tools with a short call timeout; poll SISSync.status()
  function start(opts) {
    if (window.__sisSyncState && window.__sisSyncState.running) return 'already running';
    window.__sisSyncState = { running: true, version: VERSION, results: [] };
    run(opts).catch(e => Object.assign(window.__sisSyncState, { running: false, fatal: String(e.message || e) }));
    return 'started';
  }

  function status() {
    const s = window.__sisSyncState;
    if (!s) return { running: false, note: 'not started' };
    return {
      running: s.running, current: s.current, fatal: s.fatal, version: s.version,
      done: (s.results || []).map(r => `${r.id} ${r.ok ? 'OK' : 'FAIL'} ${r.period} rows=${r.rows ?? '-'}${r.written != null ? ' written=' + r.written : ''}${r.retried ? ' (retried)' : ''}${r.error ? ' ERR=' + r.error : ''} (${r.seconds}s)`),
      ok: s.summary ? s.summary.ok : undefined, total: s.summary ? s.summary.total : undefined
    };
  }


  // ---------------- on-page progress panel (used by the bookmark button) ----------------
  function startWithPanel(opts) {
    const esc = t => String(t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    let box = document.getElementById('sis-sync-panel');
    if (!box) {
      box = document.createElement('div');
      box.id = 'sis-sync-panel';
      box.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;width:380px;max-width:calc(100vw - 32px);background:#111827;color:#f3f4f6;border-radius:10px;box-shadow:0 10px 30px rgba(0,0,0,.35);font:13px/1.45 system-ui,sans-serif;padding:14px 16px';
      document.body.appendChild(box);
    }
    if (location.hash.includes('/login')) {
      box.innerHTML = '<b style="color:#f87171">กรุณา login SIS ก่อน แล้วกดปุ่มอีกครั้ง</b>';
      setTimeout(() => box.remove(), 6000);
      return 'not logged in';
    }
    const started = start(opts);
    const render = () => {
      const st = status();
      const byId = {};
      (window.__sisSyncState?.results || []).forEach(r => { byId[r.id] = r; });
      const rows = JOBS.map(j => {
        const id = 'D' + j.id, r = byId[id];
        const icon = r ? (r.ok ? '✅' : '❌') : (st.current === id ? '⏳' : '·');
        const info = r ? (r.ok ? `${r.written ?? r.rows} แถว` : esc(r.error || '')) : (st.current === id ? 'กำลังดึงข้อมูล…' : '');
        return `<div style="display:flex;gap:8px;padding:2px 0"><span style="width:18px">${icon}</span><b style="width:28px">${id}</b><span style="color:${r && !r.ok ? '#fca5a5' : '#9ca3af'};flex:1;word-break:break-word">${info}</span></div>`;
      }).join('');
      const head = st.running
        ? '<b>กำลังอัพเดท Google Sheet…</b> <span style="color:#9ca3af">(ประมาณ 2 นาที อย่าปิดแท็บนี้)</span>'
        : (st.fatal ? `<b style="color:#f87171">หยุดทำงาน: ${esc(st.fatal)}</b>`
          : `<b style="color:${st.ok === st.total ? '#4ade80' : '#fbbf24'}">เสร็จแล้ว: สำเร็จ ${st.ok}/${st.total}</b>`);
      box.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px"><span style="font-weight:700;color:#60a5fa">Sales Dashboard Sync</span><button id="sis-sync-x" style="background:none;border:0;color:#9ca3af;font-size:18px;cursor:pointer">×</button></div><div style="margin-bottom:8px">${head}</div>${rows}`;
      box.querySelector('#sis-sync-x').onclick = () => { clearInterval(timer); box.remove(); };
      if (!st.running) clearInterval(timer);
    };
    const timer = setInterval(render, 1000);
    render();
    return started;
  }

  window.SISSync = { run, start, startWithPanel, status, periods, JOBS, VERSION };
})();
