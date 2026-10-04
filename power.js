/* ============================================================
   Surface Flow — Power Breakage v5
   power.js

   COLOR MAP (mirrors CSS :root):
   CLR.peach  #d28c64  S-Power · Source→AR41 · event tag · KPI1
   CLR.peri   #e8a040  S-Axis  · AR41→Brk   · Rx(−)     · KPI2
   CLR.teal   #d4c0a8  machines· Brk→Proc   · Rx(+)     · KPI3
   CLR.series [...]    material/chart multi-series
============================================================ */

const API = 'https://script.google.com/macros/s/AKfycbzpPE9mQznnF6oaIYW897SgqRVpsiMqZsTYj9HyMw0JADU31-6jN7slD_wJQs6f2njn/exec';


/* ============================================================
   POWER BREAKAGE PERFORMANCE AUDIT — LOGGING ONLY
   No API URL, fetch order, dashboard calculations, filters,
   charts, refresh cadence, or loader timing changes.
============================================================ */
const powerPerfState = {
  loadSeq: 0,
  initialLoadId: null,
  initialDataReadyAt: null,
  initialOverlayHideRequestedAt: null,
  initialOverlayHiddenAt: null,
  initialReadyLogged: false
};

function powerPerfNow_() {
  return performance.now();
}

function powerPerfLog_(message, details = {}) {
  console.log(`[PowerPerf] ${message}`, details);
}

function powerPerfStart_(label, details = {}) {
  const startedAt = powerPerfNow_();
  powerPerfLog_(`▶ ${label}`, details);
  return startedAt;
}

function powerPerfEnd_(label, startedAt, details = {}) {
  const elapsed = powerPerfNow_() - startedAt;
  powerPerfLog_(`✓ ${label}: ${elapsed.toFixed(1)} ms`, details);
  return elapsed;
}

powerPerfLog_(
  `JS executing at ${powerPerfNow_().toFixed(1)} ms after navigation start`,
  {}
);

window.addEventListener('load', () => {
  powerPerfLog_(
    `window.load at ${powerPerfNow_().toFixed(1)} ms after navigation start`,
    {}
  );
});

const CLR = {
  peach:  '#d28c64',
  peri:   '#e8a040',
  teal:   '#d4c0a8',
  series: ['#d28c64','#e8a040','#d4c0a8','#c8967a','#f0b870','#e8d0b8','#b87850','#d4a060'],
  gc: 'rgba(255,255,255,0.05)',
  tc: 'rgba(255,255,255,0.65)',
};

const charts = {};

// Active working sets (may be date-filtered)
let _brk=[], _reason=[], _research=[], _anom=[];

// Full loaded sets (always the complete live server response)
let _allBrk=[], _allReason=[], _allResearch=[], _allAnom=[];

// RX Flow has its own dataset (filters there never change Overview). It is
// filled from the live report, or from history when a header date is selected.
let _flowBrk=[], _flowResearch=[];

/* ============================================================
   FETCH HELPER
   Passes optional startDate / endDate (YYYY-MM-DD strings) to
   Apps Script. When supplied, script reads BREAKAGE_HISTORY
   instead of live sheets.
============================================================ */
async function fetchLiveBundle() {
  const totalStartedAt = powerPerfStart_(
    'API liveBundle',
    {}
  );

  const headersStartedAt = powerPerfNow_();
  const response = await fetch(`${API}?tab=liveBundle`);
  const headersMs = powerPerfNow_() - headersStartedAt;

  powerPerfLog_(
    `API liveBundle response headers: ${headersMs.toFixed(1)} ms`,
    {
      httpStatus: response.status,
      ok: response.ok
    }
  );

  const bodyStartedAt = powerPerfNow_();
  const responseText = await response.text();
  const bodyMs = powerPerfNow_() - bodyStartedAt;

  powerPerfLog_(
    `API liveBundle body read: ${bodyMs.toFixed(1)} ms`,
    { responseChars: responseText.length }
  );

  const parseStartedAt = powerPerfNow_();
  const d = JSON.parse(responseText);
  const parseMs = powerPerfNow_() - parseStartedAt;

  const data = d?.data || {};

  powerPerfLog_(
    `API liveBundle JSON parse: ${parseMs.toFixed(1)} ms`,
    {
      status: d?.status,
      breakageRows: Array.isArray(data.breakageSummary) ? data.breakageSummary.length : null,
      reasonRows: Array.isArray(data.reasonSummary) ? data.reasonSummary.length : null,
      researchRows: Array.isArray(data.powerResearch) ? data.powerResearch.length : null,
      anomalyRows: Array.isArray(data.anomalies) ? data.anomalies.length : null
    }
  );

  if (d.status !== 'ok') throw new Error(d.message);

  powerPerfEnd_(
    'API liveBundle',
    totalStartedAt,
    {
      responseHeadersMs: Number(headersMs.toFixed(1)),
      bodyReadMs: Number(bodyMs.toFixed(1)),
      jsonParseMs: Number(parseMs.toFixed(1)),
      responseChars: responseText.length
    }
  );

  return data;
}

async function fetchTab(tab, startDate, endDate) {
  let url = `${API}?tab=${tab}`;
  if (startDate) url += `&startDate=${encodeURIComponent(startDate)}`;
  if (endDate)   url += `&endDate=${encodeURIComponent(endDate)}`;

  const totalStartedAt = powerPerfStart_(
    `API ${tab}`,
    {
      startDate: startDate || null,
      endDate: endDate || null
    }
  );

  const headersStartedAt = powerPerfNow_();
  const response = await fetch(url);
  const headersMs = powerPerfNow_() - headersStartedAt;

  powerPerfLog_(
    `API ${tab} response headers: ${headersMs.toFixed(1)} ms`,
    {
      httpStatus: response.status,
      ok: response.ok
    }
  );

  const bodyStartedAt = powerPerfNow_();
  const responseText = await response.text();
  const bodyMs = powerPerfNow_() - bodyStartedAt;

  powerPerfLog_(
    `API ${tab} body read: ${bodyMs.toFixed(1)} ms`,
    { responseChars: responseText.length }
  );

  const parseStartedAt = powerPerfNow_();
  const d = JSON.parse(responseText);
  const parseMs = powerPerfNow_() - parseStartedAt;

  powerPerfLog_(
    `API ${tab} JSON parse: ${parseMs.toFixed(1)} ms`,
    {
      status: d?.status,
      rowCount: Array.isArray(d?.data) ? d.data.length : null
    }
  );

  if (d.status !== 'ok') throw new Error(d.message);

  powerPerfEnd_(
    `API ${tab}`,
    totalStartedAt,
    {
      responseHeadersMs: Number(headersMs.toFixed(1)),
      bodyReadMs: Number(bodyMs.toFixed(1)),
      jsonParseMs: Number(parseMs.toFixed(1)),
      responseChars: responseText.length,
      rowCount: Array.isArray(d.data) ? d.data.length : null
    }
  );

  return d.data;
}

/* ============================================================
   FORMAT HELPERS
============================================================ */
function fmtMin(m) {
  m = parseFloat(m);
  if (isNaN(m) || m <= 0) return '--';
  if (m < 60) return `${Math.round(m)}m`;
  const d  = Math.floor(m / 1440);
  const h  = Math.floor((m % 1440) / 60);
  const mn = Math.round(m % 60);
  return [d > 0 ? `${d}d` : '', h > 0 ? `${h}h` : '', mn > 0 ? `${mn}m` : ''].filter(Boolean).join(' ');
}

function fmtDate(s) {
  if (!s) return '--';
  const d = new Date(s);
  if (isNaN(d)) return s;
  return `${d.getMonth()+1}/${d.getDate()} ${d.getHours()}:${String(d.getMinutes()).padStart(2,'0')}`;
}

function diffMin(a, b) {
  // Missing, unparseable, or out-of-order timestamps return null (no timing), never 0.
  if (!a || !b) return null;
  const da = new Date(a), db = new Date(b);
  if (isNaN(da) || isNaN(db)) return null;
  const m = (db - da) / 60000;
  return m >= 0 ? m : null;
}

function avgArr(arr) {
  const n = arr.map(v => parseFloat(v)).filter(v => !isNaN(v) && v !== null);
  return n.length ? n.reduce((a, b) => a + b, 0) / n.length : 0;
}

function reasonColor(r) {
  if (!r) return CLR.teal;
  const s = r.toLowerCase();
  if (s.includes('power')) return CLR.peach;
  if (s.includes('axis'))  return CLR.peri;
  return CLR.teal;
}

function reasonPillClass(r) {
  if (!r) return 'pill-other';
  const s = r.toLowerCase();
  if (s.includes('power')) return 'pill-peach';
  if (s.includes('axis'))  return 'pill-peri';
  return 'pill-other';
}

/* ============================================================
   CHART HELPER
============================================================ */
function mkChart(id, type, labels, datasets, extraOpts = {}) {
  if (charts[id]) charts[id].destroy();
  function merge(t, s) {
    const o = Object.assign({}, t);
    for (const k in s) o[k] = (s[k] && typeof s[k] === 'object' && !Array.isArray(s[k])) ? merge(t[k] || {}, s[k]) : s[k];
    return o;
  }
  const base = {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
      x: { grid: { color: CLR.gc }, ticks: { color: CLR.tc, font: { size: 11, family: "'IBM Plex Mono'" } } },
      y: { grid: { color: CLR.gc }, ticks: { color: CLR.tc, font: { size: 11, family: "'IBM Plex Mono'" } } },
    },
  };
  charts[id] = new Chart(document.getElementById(id), { type, data: { labels, datasets }, options: merge(base, extraOpts) });
}

/* ============================================================
   NAVIGATION
============================================================ */
function sw(id, btn) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('t-' + id).classList.add('active');
  btn.classList.add('active');
}

function swInner(id, btn) {
  document.querySelectorAll('.inner-section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.inner-tab').forEach(b => b.classList.remove('active'));
  document.getElementById('is-' + id).classList.add('active');
  btn.classList.add('active');
  setTimeout(() => { if (charts[id + 'C']) charts[id + 'C'].resize(); }, 50);
}

/* ============================================================
   BAR LIST BUILDER
============================================================ */
function buildBarList(elId, entries) {
  const el = document.getElementById(elId);
  if (!el) return;
  const max = entries[0]?.[1] || 1;
  el.innerHTML = entries.map(([label, val, display, color]) => `
    <div class="bar-row">
      <div class="bar-label">${label}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.round(val / max * 100)}%;background:${color || CLR.teal}"></div></div>
      <div class="bar-val">${display || val}</div>
    </div>`).join('');
}

/* ============================================================
   OVERVIEW
============================================================ */
function buildOverview() {
  /* Overview v6 — see OVERVIEW HELPERS at the end of this file.
     Timing averages use only jobs that have a timing (untimed = excluded, not 0). */
  const totalJobs = _brk.length;
  const totalLens = _brk.reduce((s, r) => s + (parseInt(r.LensesBroken) || 0), 0);
  const ar41      = ovTimingStats_(_brk, 'AR41_to_Breakage_Min');
  const brkPro    = ovTimingStats_(_brk, 'Breakage_to_Processed_Min');
  const ar41Avg   = ar41.avg;
  const mode      = ovRenderMode_();

  if (mode === 'switch') ovFadeIn_();

  // ── Context title (follows the selected date) ──
  ovRenderContext_();

  // ── Data-quality note (sheet errors such as #N/A) ──
  ovRenderDataQuality_(_brk);

  // ── KPI: jobs / lenses ──
  ovSetCount_('m-jobs', totalJobs, mode);
  ovSetCount_('m-lens', totalLens, mode);
  ovJobsDelta_(totalJobs, mode);

  // ── KPI: reason split (jobs) ──
  const rsnAgg = {};
  _brk.forEach(r => {
    const k = r.BrkReason || 'Unknown';
    if (!rsnAgg[k]) rsnAgg[k] = { jobs: 0, lens: 0 };
    rsnAgg[k].jobs += 1;
    rsnAgg[k].lens += (parseInt(r.LensesBroken) || 0);
  });
  const rsnE = Object.entries(rsnAgg).sort((a, b) => b[1].jobs - a[1].jobs);
  ovRenderReasonSplit_(rsnE, totalJobs, mode);

  // ── KPI: timing (average of timed jobs + coverage + per-reason) ──
  ovRenderTiming_('ar', 'm-ar41brk', ar41, totalJobs, rsnE, 'AR41_to_Breakage_Min', mode);
  ovRenderTiming_('bp', 'm-brkpro',  brkPro, totalJobs, rsnE, 'Breakage_to_Processed_Min', mode);

  // ── Banner (event strip) — unchanged wording, now uses the corrected average ──
  const worst = [..._brk].sort((a, b) => (parseFloat(b.AR41_to_Breakage_Min) || 0) - (parseFloat(a.AR41_to_Breakage_Min) || 0))[0];
  if (worst) document.getElementById('bannerSub').textContent =
    `${worst.BrkSourceMachine || worst.AR41Machine} — avg AR41→Brk ${fmtMin(ar41Avg)} · ${totalJobs} jobs`;
  document.getElementById('bannerTime').textContent = fmtEtStamp_(new Date());  // ET, 12-hour

  // ── Where it broke: jobs by BrkSource machine ──
  const srcAgg = {}, srcRx = {};
  _brk.forEach(r => {
    const s = r.BrkSourceMachine || 'Unknown';
    srcAgg[s] = (srcAgg[s] || 0) + 1;
    if (!srcRx[s]) srcRx[s] = [];
    srcRx[s].push(r);
  });
  const srcE = Object.entries(srcAgg).sort((a, b) => b[1] - a[1]);
  ovRenderBars_('srcBarList', srcE, 'jobs', mode, label => {
    const jobs = srcRx[label] || [];
    const rc = {};
    const tl = jobs.reduce((s, j) => s + (parseInt(j.LensesBroken) || 0), 0);
    jobs.forEach(j => { const r = j.BrkReason || '?'; rc[r] = (rc[r] || 0) + 1; });
    return [...Object.entries(rc).map(([r, c]) => `${r}: ${c}`), `Lenses: ${tl}`].join(' · ');
  });
  const srcSum = srcE.reduce((s, e) => s + e[1], 0);
  const srcMax = srcE[0]?.[1] || 0;
  const srcSmall = srcE.length > 0 && srcMax < OV_SMALL_N;
  document.getElementById('ovSrcTags').innerHTML = '';   // badge removed by request; footer states the small sample
  document.getElementById('ovSrcFoot').innerHTML = srcE.length
    ? `Total <b>${srcSum}</b> jobs${srcSum === totalJobs ? '' : ` <span class="ov-bad">≠ ${totalJobs} in KPI</span>`}.` +
      (srcSmall ? ` Every machine has fewer than ${OV_SMALL_N} jobs, so differences of 1–2 are noise.` : '')
    : '';

  // ── What broke: lenses by material ──
  const matAgg = {};
  _brk.forEach(r => { const m = r.Material || 'Unknown'; matAgg[m] = (matAgg[m] || 0) + (parseInt(r.LensesBroken) || 0); });
  const matE = Object.entries(matAgg).sort((a, b) => b[1] - a[1]);
  ovRenderBars_('matList', matE, 'lenses', mode);
  const matSum = matE.reduce((s, e) => s + e[1], 0);
  document.getElementById('ovMatFoot').innerHTML = matE.length
    ? `Total <b>${matSum}</b> lenses${matSum === totalLens ? '' : ` <span class="ov-bad">≠ ${totalLens} in KPI</span>`}. Counted in lenses, not jobs.`
    : '';

  ovAfterRender_();
}

/* ============================================================
   RX FLOW SOURCE-RUN CLASSIFICATION
   The live breakage report is kept intact. Rows are separated by
   the ORB/OTB source-machine scan date: today vs previous days.
============================================================ */
function isoLocalDate(d = new Date()) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function safeIsoDate(raw) {
  if (!raw) return '';
  const d = new Date(raw);
  return isNaN(d) ? String(raw).slice(0, 10) : isoLocalDate(d);
}

function getSourceScanTime(r, res) {
  const src = (r.BrkSourceMachine || res?.BrkSourceMachine || res?.ORBMachine || '').toUpperCase();
  if (src.includes('OTB')) return res?.OTBScanTime || r.OTBScanTime || null;
  return res?.ORBScanTime || r.ORBScanTime || res?.OTBScanTime || r.OTBScanTime || null;
}

function sourceRunDate(r, res) {
  return safeIsoDate(getSourceScanTime(r, res));
}

/* ============================================================
   FLOW FILTERS
============================================================ */
function populateFlowFilters() {
  [
    ['filterReason', _flowBrk.map(r => r.BrkReason)],
    ['filterSource', _flowBrk.map(r => r.BrkSourceMachine)],
    ['filterAR41',   _flowBrk.map(r => r.AR41Machine)],
  ].forEach(([id, vals]) => {
    const el = document.getElementById(id);
    while (el.options.length > 1) el.remove(1);
    [...new Set(vals.filter(Boolean))].sort().forEach(v => {
      const o = document.createElement('option'); o.value = v; o.textContent = v; el.appendChild(o);
    });
  });
}

/* ── Tooltip ── */
const tip = document.getElementById('ganttTip');
function showTip(e, html) { if(!tip) return; tip.innerHTML = html; tip.style.display = 'block'; moveTip(e); }
function moveTip(e) {
  if(!tip) return;
  const x = e.clientX + 16;
  const y = e.clientY + 16;
  const w = tip.offsetWidth || 280;
  const h = tip.offsetHeight || 160;
  tip.style.left = (x + w > window.innerWidth  ? e.clientX - w - 10 : x) + 'px';
  tip.style.top  = (y + h > window.innerHeight ? e.clientY - h - 10 : y) + 'px';
}
function hideTip() { if(!tip) return; tip.style.display = 'none'; }
document.addEventListener('mousemove', e => { if (tip && tip.style.display === 'block') moveTip(e); });

/* ============================================================
   RENDER FLOW
============================================================ */
function renderFlow() {
  /* RX Flow v6 — Option A: process pipeline + every job on one shared clock.
     A stage with no recorded time is never drawn as a bar. It is labelled
     "skipped" (a later stage exists, or the viewed day is over) or
     "not scanned yet" (live view, after the job's last recorded scan). */
  const fR = document.getElementById('filterReason').value;
  const fS = document.getElementById('filterSource').value;
  const fA = document.getElementById('filterAR41').value;

  let data = [..._flowBrk];
  if (fR) data = data.filter(r => r.BrkReason === fR);
  if (fS) data = data.filter(r => r.BrkSourceMachine === fS);
  if (fA) data = data.filter(r => r.AR41Machine === fA);

  const filteredLenses = data.reduce((sum, r) => sum + (parseInt(r.LensesBroken, 10) || 0), 0);
  document.getElementById('flowCount').textContent = `${data.length} jobs · ${filteredLenses} lenses`;

  const resMap = {};
  _flowResearch.forEach(r => { resMap[String(r.RxNumber)] = r; });

  const histDate = document.getElementById('dateSingle')?.value || '';
  const anchor   = histDate || isoLocalDate();          // the day being viewed
  const dayLbl   = histDate ? ovPrettyDate_(histDate) : 'today';

  // ── Stage minutes per job (null = no recorded time) ──
  const jobs = data.map(r => {
    const res = resMap[String(r.RxNumber)];
    const srcTime = getSourceScanTime(r, res);
    const s1 = srcTime ? diffMin(srcTime, r.AR41ScanTime) : null;
    const s2 = parseFloat(r.AR41_to_Breakage_Min);
    const s3 = parseFloat(r.Breakage_to_Processed_Min);
    const st = [s1, s2, s3].map(v => Number.isFinite(v) ? v : null);
    const total = st.reduce((a, v) => a + (v || 0), 0);
    return { r, res, srcTime, st, total, runDate: sourceRunDate(r, res) };
  });

  flowRenderPipeline_(jobs);

  if (!jobs.length) {
    document.getElementById('flowList').innerHTML = '<div class="empty">No jobs match filters</div>';
    return;
  }

  // ── Shared time scale for every row ──
  const axisMax = flowNiceMax_(Math.max(60, ...jobs.map(j => j.total)));
  const pct = m => (m / axisMax) * 100;

  const runToday = [], runPrevious = [], runUnknown = [];
  jobs.forEach(j => {
    if (j.runDate === anchor) runToday.push(j);
    else if (j.runDate && j.runDate < anchor) runPrevious.push(j);
    else runUnknown.push(j);
  });

  const STAGE = [
    { cls: 'seg-1', short: 'Transit',    tipHead: '① JOB IN TRANSIT — SOURCE TO AR41',     dur: 'Transit Time' },
    { cls: 'seg-2', short: 'Identify',   tipHead: '② BREAKAGE IDENTIFIED — AR41 TO TABLE', dur: 'Identification Time' },
    { cls: 'seg-3', short: 'Processing', tipHead: '③ BREAKAGE PROCESSING — TABLE TO REORDER', dur: 'Processing Time' },
  ];

  function makeTip(segClass, headerLabel, rows, durationLabel, durationVal) {
    const rowsHtml = rows.map(([lbl, val]) =>
      `<div class="tip-row"><div class="tip-row-label">${lbl}</div><div class="tip-row-val">${val}</div></div>`
    ).join('<div class="tip-divider"></div>');
    return `<div class="tip-header ${segClass}">${headerLabel}</div>
            <div class="tip-body">${rowsHtml}</div>
            <div class="tip-duration">
              <span class="tip-duration-label">${durationLabel}</span>
              <span class="tip-duration-val ${segClass}">${durationVal}</span>
            </div>`;
  }

  // Missing stages → "skipped" / "not scanned yet", merged when consecutive
  function skipText(st) {
    const last = st.reduce((acc, v, i) => (v !== null ? i : acc), -1);
    const labels = [];
    st.forEach((v, i) => {
      if (v !== null) return;
      const trailing = i > last;
      labels.push({ i, kind: (!histDate && trailing) ? 'pending' : 'skipped' });
    });
    if (!labels.length) return '';
    if (labels.length === 3) {
      return `<span class="fj-skip" title="No scan time was recorded for any stage">${histDate ? 'All stages skipped' : 'No stage scanned yet'}</span>`;
    }
    const groups = [];
    labels.forEach(l => {
      const g = groups[groups.length - 1];
      if (g && g.kind === l.kind && g.last === l.i - 1) { g.names.push(STAGE[l.i].short); g.last = l.i; }
      else groups.push({ kind: l.kind, names: [STAGE[l.i].short], last: l.i });
    });
    return groups.map(g => {
      const what = g.names.join(' + ');
      return g.kind === 'pending'
        ? `<span class="fj-skip pending" title="No scan recorded yet — the job may still be on its way">${what} not scanned yet</span>`
        : `<span class="fj-skip" title="No scan time recorded for this stage (skipped, or scanned out of order)">${what} skipped</span>`;
    }).join('');
  }

  function makeRow(j) {
    const { r, res, srcTime, st, total } = j;
    const src = (r.BrkSourceMachine || '').toUpperCase();
    const srcMachine  = r.BrkSourceMachine || res?.ORBMachine || '?';
    const ar41Machine = r.AR41Machine || '?';
    const srcLabel = src.includes('ORB') ? `ORB — ${srcMachine}` : (src.includes('OTB') ? `OTB — ${srcMachine}` : srcMachine);
    const operatorScan = r.OperatorScan || res?.OperatorScan || '—';

    const tipRows = [
      [['Source Machine', srcLabel], ['Departed Source', fmtDate(srcTime)], ['Source Run Day', safeIsoDate(srcTime) || 'Unavailable'],
       ['Arrived AR41', fmtDate(r.AR41ScanTime)], ['AR41 Machine', ar41Machine]],
      [['AR41 Machine', ar41Machine], ['AR41 Scan Time', fmtDate(r.AR41ScanTime)], ['Operator Scan', operatorScan],
       ['Scan to Breakage Table', fmtDate(r.BrkTableScanTime)]],
      [['Operator Scan', operatorScan], ['Breakage Table Scan', fmtDate(r.BrkTableScanTime)],
       ['Breakage Processed', fmtDate(r.BreakageProcessedTime)], ['Breakage Reason', r.BrkReason || '—']],
    ];

    let segs = '';
    st.forEach((v, i) => {
      if (v === null) return;
      const tip = makeTip(STAGE[i].cls, STAGE[i].tipHead, tipRows[i], STAGE[i].dur, fmtMin(v)).replace(/'/g, '&#39;');
      segs += `<div class="fj-seg ${STAGE[i].cls}" style="width:${total ? v / total * 100 : 0}%" data-tip='${tip}' onmouseenter="showTip(event,this.dataset.tip)" onmouseleave="hideTip()"></div>`;
    });

    const dots = st.map((v, i) => `<i class="fj-dot ${STAGE[i].cls}${v !== null ? ' on' : ''}" title="${STAGE[i].short}: ${v !== null ? fmtMin(v) : 'no scan'}"></i>`).join('');

    return `
      <div class="fj-row" data-rx="${ovEsc_(r.RxNumber)}" data-w="${pct(total)}">
        <div class="fj-rx" title="${ovEsc_(r.RxNumber)}">${ovEsc_(r.RxNumber)}</div>
        <span class="fj-reason">${ovEsc_(r.BrkReason || '--')}</span>
        <div class="fj-lane">
          <div class="fj-bar" style="width:${pct(total)}%">${segs}</div>
          <div class="fj-after" style="left:${pct(total)}%">
            <b class="fj-total">${total ? fmtMin(total) : '--'}</b>${skipText(st)}
          </div>
        </div>
        <span class="fj-dots">${dots}</span>
      </div>`;
  }

  function makeGroup(title, subtitle, rows, groupClass) {
    const sorted = [...rows].sort((a, b) => b.total - a.total);
    const lenses = sorted.reduce((sum, j) => sum + (parseInt(j.r.LensesBroken, 10) || 0), 0);
    return `
      <section class="flow-run-group ${groupClass}">
        <div class="flow-run-head">
          <div>
            <div class="flow-run-title">${title}</div>
            <div class="flow-run-sub">${subtitle}</div>
          </div>
          <div class="flow-run-count"><strong>${sorted.length}</strong> jobs <span>·</span> <strong>${lenses}</strong> lenses</div>
        </div>
        <div class="flow-run-rows">${sorted.length ? sorted.map(makeRow).join('') : '<div class="empty">No breakages in this group</div>'}</div>
      </section>`;
  }

  // remember previous bar lengths so a refresh can extend bars instead of redrawing
  const prevW = {};
  document.querySelectorAll('#flowList .fj-row').forEach(el => { prevW[el.dataset.rx] = Number(el.dataset.w); });
  const prevAnchor = flowState_.anchor;
  flowState_.anchor = anchor;

  document.getElementById('flowList').innerHTML =
    flowAxisHtml_(axisMax) +
    (histDate
      ? makeGroup(`RUN ${dayLbl.toUpperCase()}`, `Source-machine scan date is ${dayLbl}`, runToday, 'run-today') +
        makeGroup(`RUN BEFORE ${dayLbl.toUpperCase()}`, `Source-machine scan date is before ${dayLbl}`, runPrevious, 'run-previous')
      : makeGroup('RUN TODAY', 'Source-machine scan date is today', runToday, 'run-today') +
        makeGroup('RUN PREVIOUS DAYS', 'Source-machine scan date is before today', runPrevious, 'run-previous')) +
    (runUnknown.length
      ? makeGroup('SOURCE DATE MISSING', `No ORB/OTB scan date${histDate ? `, or it is after ${dayLbl}` : ''} — check the source sheet`, runUnknown, 'run-unknown')
      : '');

  flowAnimate_(prevAnchor === anchor ? prevW : null);
}

/* ============================================================
   RESEARCH FILTERS
============================================================ */
function populateResearchFilters() {
  const isHistorical = (document.getElementById('dateSingle')?.value || '') !== '';
  const orbVals = isHistorical
    ? _research.map(r => r.BrkSourceMachine || r.ORBMachine)
    : _research.map(r => r.ORBMachine);

  // Enrich Material from _brk so the Material dropdown is populated
  const brkMatMap = {};
  _brk.forEach(r => { if (r.RxNumber && r.Material && r.Material !== 'Unknown') brkMatMap[String(r.RxNumber)] = r.Material; });
  const matVals = _research.map(r => {
    if (r.Material && r.Material !== 'Unknown') return r.Material;
    return brkMatMap[String(r.RxNumber)] || '';
  });

  [
    ['resORB',      orbVals],
    ['resReason',   _research.map(r => r.BrkReason)],
    ['resMaterial', matVals],
  ].forEach(([id, vals]) => {
    const el = document.getElementById(id);
    while (el.options.length > 1) el.remove(1);
    [...new Set(vals.filter(Boolean))].sort().forEach(v => {
      const o = document.createElement('option'); o.value = v; o.textContent = v; el.appendChild(o);
    });
  });
}

/* ── Sph / Cyl bucket helpers ── */
function sphBuckets(vals) {
  const b = { '≤-6': 0, '-6/-4': 0, '-4/-2': 0, '-2/0': 0, '0/+2': 0, '+2/+4': 0, '≥+4': 0 };
  vals.forEach(v => {
    if (isNaN(v)) return;
    if (v <= -6)      b['≤-6']++;
    else if (v <= -4) b['-6/-4']++;
    else if (v <= -2) b['-4/-2']++;
    else if (v < 0)   b['-2/0']++;
    else if (v < 2)   b['0/+2']++;
    else if (v < 4)   b['+2/+4']++;
    else              b['≥+4']++;
  });
  return b;
}

function cylBuckets(vals) {
  const b = { '≤-3': 0, '-3/-2': 0, '-2/-1': 0, '-1/0': 0, '0': 0, '0/+1': 0, '≥+1': 0 };
  vals.forEach(v => {
    if (isNaN(v)) return;
    if (v <= -3)      b['≤-3']++;
    else if (v <= -2) b['-3/-2']++;
    else if (v <= -1) b['-2/-1']++;
    else if (v < 0)   b['-1/0']++;
    else if (v === 0) b['0']++;
    else if (v < 1)   b['0/+1']++;
    else              b['≥+1']++;
  });
  return b;
}

function sphColors(keys) {
  return keys.map(k => (k.startsWith('+') || k === '0/+2' || k === '≥+4') ? '#d4c030' : '#4d94d4');
}

function cylColors(keys) {
  const neg = new Set(['≤-3', '-3/-2', '-2/-1', '-1/0']);
  const pos = new Set(['0/+1', '≥+1']);
  return keys.map(k => {
    if (k === '0')   return 'rgba(255,255,255,0.2)';
    if (pos.has(k))  return '#4d94d4';
    if (neg.has(k))  return '#d4c030';
    return 'rgba(255,255,255,0.2)';
  });
}

/* ============================================================
   RENDER RESEARCH
============================================================ */
function renderResearch() {
  /* Analysis v6 — plain-language Rx view.
     Each job is counted once, by its stronger eye. Cutoffs: RX_CUTOFFS (end of file). */
  const fO = document.getElementById('resORB').value;
  const fR = document.getElementById('resReason').value;
  const fM = document.getElementById('resMaterial').value;

  let data = [..._research];
  if (fO) data = data.filter(r => (r.BrkSourceMachine || r.ORBMachine) === fO);
  if (fR) data = data.filter(r => r.BrkReason === fR);

  // Build Material lookup from _brk — BreakageSummary always has Material populated
  const brkMaterialMap = {};
  _brk.forEach(r => {
    if (r.RxNumber && r.Material && r.Material !== 'Unknown')
      brkMaterialMap[String(r.RxNumber)] = r.Material;
  });

  // Enrich rows where Material is missing, then apply material filter
  data = data.map(r => {
    if (r.Material && r.Material !== 'Unknown') return r;
    const mat = brkMaterialMap[String(r.RxNumber)];
    return mat ? { ...r, Material: mat } : r;
  });
  if (fM) data = data.filter(r => r.Material === fM);

  const totalLens = data.reduce((s, r) => s + (parseInt(r.LensesBroken) || 0), 0);
  document.getElementById('res-jobs').textContent = data.length;
  document.getElementById('res-lens').textContent = totalLens;

  const isHistorical = (document.getElementById('dateSingle')?.value || '') !== '';
  const jobs  = data.map(r => ({ r, sph: rxStrongest_(r.R_Sph, r.L_Sph), cyl: rxStrongest_(r.R_Cyl, r.L_Cyl) }));
  const noRx  = jobs.every(j => j.sph === null && j.cyl === null);

  // ── Live-only note (history has no Rx) ──
  const rxNote = document.getElementById('rxHistoricalNote');
  if (rxNote) rxNote.style.display = (noRx && data.length) ? 'flex' : 'none';
  const cards = document.getElementById('anRxCards');
  if (cards) cards.hidden = noRx && data.length > 0;

  anRenderHeadline_(jobs, totalLens, noRx);

  // ── Sphere / cylinder cards ──
  anRenderScale_('anSphScale', 'anSphMiss', jobs, 'sph', rxSphCats_());
  anRenderScale_('anCylScale', 'anCylMiss', jobs, 'cyl', rxCylCats_());

  // ── Material / reason bars ──
  const agg = key => {
    const m = {};
    data.forEach(r => {
      const k = r[key] || 'Unknown';
      if (!m[k]) m[k] = { jobs: 0, lens: 0 };
      m[k].jobs += 1; m[k].lens += (parseInt(r.LensesBroken) || 0);
    });
    return Object.entries(m).sort((a, b) => b[1].jobs - a[1].jobs || b[1].lens - a[1].lens);
  };
  anRenderBars_('anMatBars', agg('Material'), data.length, () => CLR.teal, false);
  anRenderBars_('anRsnBars', agg('BrkReason'), data.length, k => reasonColor(k), true);

  // ── Every job ──
  const sCats = rxSphCats_(), cCats = rxCylCats_();
  const word = (v, cats) => v === null ? '—' : (cats.find(c => c.test(v))?.name || '—');
  const num  = v => {
    const n = parseFloat(v);
    if (v === '' || v === null || v === undefined || isNaN(n)) return '—';
    return n === 0 ? '0' : (n > 0 ? `+${n}` : `${n}`);
  };
  const tableEl = document.getElementById('resTable');
  if (!data.length) { tableEl.innerHTML = '<div class="empty">No jobs match these filters</div>'; anAfterRender_(); return; }

  tableEl.innerHTML = `
    <table class="an-table">
      <thead>
        <tr class="an-grp"><th colspan="3"></th>${noRx ? '' : '<th colspan="3">Sphere</th><th colspan="3">Cylinder</th><th></th><th colspan="2"></th>'}<th colspan="2"></th></tr>
        <tr>
          <th>RX</th><th>Reason</th><th>Material</th>
          ${noRx ? '' : '<th>Strength</th><th class="num">Right</th><th class="num">Left</th><th>Amount</th><th class="num">Right</th><th class="num">Left</th><th class="num">Add</th><th>Lens option</th><th>Curve</th>'}
          <th>${isHistorical ? 'Broke at' : 'ORB → broke at'}</th><th class="num">Lenses</th>
        </tr>
      </thead>
      <tbody>${jobs.map(({ r, sph, cyl }) => `
        <tr data-rx="${ovEsc_(r.RxNumber)}">
          <td class="an-rx">${ovEsc_(r.RxNumber || '--')}</td>
          <td><span class="an-pill"><i style="background:${reasonColor(r.BrkReason)}"></i>${ovEsc_(r.BrkReason || '--')}</span></td>
          <td>${ovEsc_(r.Material || '--')}</td>
          ${noRx ? '' : `
          <td><b>${word(sph, sCats)}</b></td><td class="num an-mono">${num(r.R_Sph)}</td><td class="num an-mono">${num(r.L_Sph)}</td>
          <td><b>${word(cyl, cCats)}</b></td><td class="num an-mono">${num(r.R_Cyl)}</td><td class="num an-mono">${num(r.L_Cyl)}</td>
          <td class="num an-mono">${num(r.R_AddPower)}</td><td>${ovEsc_(r.LensOption || '—')}</td><td class="an-mono">${ovEsc_(r.BaseCurve || '—')}</td>`}
          <td>${isHistorical ? ovEsc_(r.BrkSourceMachine || '--') : `${ovEsc_(r.ORBMachine || '--')} → ${ovEsc_(r.BrkSourceMachine || '--')}`}</td>
          <td class="num an-mono"><b>${parseInt(r.LensesBroken) || 0}</b></td>
        </tr>`).join('')}
      </tbody>
    </table>
    <div class="an-cap">${data.length} ${data.length === 1 ? 'job' : 'jobs'} · ${totalLens} ${totalLens === 1 ? 'lens' : 'lenses'}</div>`;

  anAfterRender_();
}

/* ============================================================
   ALERTS
============================================================ */
function buildAlerts() {
  const total  = _anom.length;
  const sPower = _anom.filter(r => (r.BrkReason || '').toLowerCase().includes('power')).length;
  const sAxis  = _anom.filter(r => (r.BrkReason || '').toLowerCase().includes('axis')).length;

  document.getElementById('al-total').textContent  = total;
  document.getElementById('al-spower').textContent = sPower;
  document.getElementById('al-saxis').textContent  = sAxis;

  const sorted = [..._anom]
    .map(r => ({ ...r, _v: parseFloat(r.AR41_to_Breakage_Min) || 0 }))
    .sort((a, b) => b._v - a._v);
  const maxV = sorted[0]?._v || 1;

  document.getElementById('alertList').innerHTML = sorted.map(r => {
    const color = reasonColor(r.BrkReason);
    const pct   = Math.max(2, Math.round(r._v / maxV * 100));
    return `
      <div class="alert-bar-row">
        <div>
          <div class="alert-bar-rx">${r.RxNumber}</div>
          <div class="alert-bar-desc">${r.AR41Machine || '--'} · ${r.BrkSourceMachine || '--'} · ${r.BrkReason || '--'}</div>
        </div>
        <div class="alert-bar-track"><div class="alert-bar-fill" style="width:${pct}%;background:${color}"></div></div>
        <div class="alert-bar-val" style="color:${color}">${fmtMin(r._v)}</div>
      </div>`;
  }).join('') || '<div class="empty">No alerts</div>';

  const worst   = sorted[0];
  const avgAR41 = avgArr(_anom.map(r => r.AR41_to_Breakage_Min));
  const avgBrk  = avgArr(_anom.map(r => r.Breakage_to_Processed_Min));

  document.getElementById('quickStats').innerHTML = [
    ['Total delays', total,                          ''],
    ['S-Power',      sPower,                         CLR.peach],
    ['S-Axis',       sAxis,                          CLR.peri],
    ['Avg AR41→Brk', fmtMin(avgAR41),                CLR.teal],
    ['Avg Brk→Proc', fmtMin(avgBrk),                 ''],
    ['Worst job',    worst ? fmtMin(worst._v) : '--', CLR.peach],
  ].map(([label, val, color]) => `
    <div class="stat-row">
      <span class="stat-label">${label}</span>
      <span class="stat-val" style="color:${color || 'inherit'}">${val}</span>
    </div>`).join('');
}

/* ============================================================
   POWER ANALYSIS SPLASH — OVERLAY DRIVERS
============================================================ */
const PA_STEPS = ['ls-brk','ls-reason','ls-research','ls-anom'];
const PA_LABELS = {
  'ls-brk':      'Loading breakage records…',
  'ls-reason':   'Aggregating reason data…',
  'ls-research': 'Pulling research prescriptions…',
  'ls-anom':     'Scanning for anomalies…',
};

function setLoadStep(id, state) {
  // state: 'active' | 'done' | ''
  PA_STEPS.forEach(s => {
    const el = document.getElementById(s);
    if (!el) return;
    el.classList.remove('active','done');
  });

  // Mark all steps before this one as done, this one as state
  let found = false;
  PA_STEPS.forEach(s => {
    const el = document.getElementById(s);
    if (!el) return;
    if (s === id) { found = true; if (state) el.classList.add(state); return; }
    if (!found) el.classList.add('done');
  });

  // Update progress bar
  const idx = PA_STEPS.indexOf(id);
  const pct = state === 'done'
    ? 100
    : state === 'active'
      ? Math.round((idx / PA_STEPS.length) * 100)
      : 0;
  const bar = document.getElementById('paProgressBar');
  if (bar) bar.style.width = pct + '%';

  // Update label
  const lbl = document.getElementById('paProgressLabel');
  if (lbl && state === 'active') lbl.textContent = PA_LABELS[id] || '';
  if (lbl && state === 'done')   lbl.textContent = 'Almost ready…';
}

function showOverlay() {
  powerPerfLog_('Loading overlay shown', {});
  const o = document.getElementById('loadingOverlay');
  if (o) o.classList.remove('hidden');
  // Reset all steps
  PA_STEPS.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('active','done');
  });
  const bar = document.getElementById('paProgressBar');
  if (bar) bar.style.width = '0%';
  const lbl = document.getElementById('paProgressLabel');
  if (lbl) lbl.textContent = 'Connecting to data source…';
}

function hideOverlay() {
  const hideRequestedAt = powerPerfNow_();

  if (powerPerfState.initialLoadId !== null && !powerPerfState.initialReadyLogged) {
    powerPerfState.initialOverlayHideRequestedAt = hideRequestedAt;
  }

  powerPerfLog_(
    'hideOverlay() called; existing 400 ms splash hide delay begins',
    {}
  );

  const bar = document.getElementById('paProgressBar');
  if (bar) bar.style.width = '100%';
  const lbl = document.getElementById('paProgressLabel');
  if (lbl) lbl.textContent = 'Ready';
  // Mark all done
  PA_STEPS.forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.classList.remove('active'); el.classList.add('done'); }
  });
  setTimeout(() => {
    const o = document.getElementById('loadingOverlay');
    if (o) o.classList.add('hidden');

    if (
      powerPerfState.initialLoadId !== null &&
      !powerPerfState.initialReadyLogged
    ) {
      powerPerfState.initialOverlayHiddenAt = powerPerfNow_();
      powerPerfState.initialReadyLogged = true;

      const totalReadyMs = powerPerfState.initialOverlayHiddenAt;
      const hideDelayMs = powerPerfState.initialOverlayHideRequestedAt == null
        ? null
        : powerPerfState.initialOverlayHiddenAt -
          powerPerfState.initialOverlayHideRequestedAt;

      powerPerfLog_('✅ INITIAL POWER BREAKAGE DASHBOARD READY', {
        fetchId: powerPerfState.initialLoadId,
        dataRenderReadyMs: powerPerfState.initialDataReadyAt == null
          ? null
          : powerPerfState.initialDataReadyAt.toFixed(1),
        preHideDelayMs: 300,
        overlayHideDelayMs: hideDelayMs == null
          ? null
          : Number(hideDelayMs.toFixed(1)),
        totalReadyMs: Number(totalReadyMs.toFixed(1))
      });

      console.log(
        `[PowerPerf] MAIN POWER BREAKAGE DASHBOARD TIME TO READY: ${totalReadyMs.toFixed(1)} ms (${(totalReadyMs / 1000).toFixed(2)} sec)`
      );
    }
  }, 400);
}

/* ============================================================
   DATE FILTER HELPERS  (client-side date objects for sub-filters)
============================================================ */
function parseDateInput(id) {
  const v = document.getElementById(id)?.value;
  return v ? new Date(v + 'T00:00:00') : null;
}

function inDateRange(dateStr, from, to) {
  if (!from && !to) return true;
  const d = new Date(dateStr);
  if (isNaN(d)) return true;
  if (from && d < from) return false;
  if (to   && d > new Date(to.getTime() + 86400000 - 1)) return false;
  return true;
}

/* ============================================================
   FIELD NAME HELPERS
   Maps both live sheet field names (camelCase, from BreakageSummary/
   PowerResearch) AND BREAKAGE_HISTORY raw column names (with spaces,
   exactly as seen in the sheet headers).

   BREAKAGE_HISTORY columns confirmed from sheet:
   SnapshotDate | RX Number | OTB Blocker Scan Point | OTB Blocker Scan Time
   ORB Scan Point | ORB Scan Time | AR41 Scan Point | AR41 Scan Time
   Operator Scan | Scan to Breakage Table | Brk Scan Point | Brk Reason
   Breakage Processed Time | Lenses Broken | [Material etc from RawData]

   Live BreakageSummary columns:
   RxNumber | AR41Machine | AR41ScanTime | BrkTableScanTime
   BrkSourceMachine | BrkReason | BreakageProcessedTime | LensesBroken
   Material | AR41_to_Breakage_Min | Breakage_to_Processed_Min | Status
   NOTE: BreakageSummary does NOT include Operator Scan in live mode.
         Operator Scan is available via PowerResearch (S-Power/S-Axis only)
         and via BREAKAGE_HISTORY in historical mode.

   Live PowerResearch columns:
   RxNumber | OTBMachine | OTBScanTime | ORBMachine | ORBScanTime
   AR41Machine | AR41ScanTime | OperatorScan | BrkSourceMachine | BrkReason | Material
   LensOption | BaseCurve | R_Sph | L_Sph | R_Cyl | L_Cyl | R_AddPower ...
============================================================ */
function fv(row, ...keys) {
  for (const k of keys) {
    const v = row[k];
    if (v !== undefined && v !== null && v !== '') return v;
  }
  return '';
}

// Calculate minutes between two date strings
function calcMin(fromStr, toStr) {
  // Missing, unparseable, or out-of-order timestamps return null (no timing),
  // never 0 — a fake 0 drags every average down. avgArr() skips null.
  if (!fromStr || !toStr) return null;
  const a = new Date(fromStr), b = new Date(toStr);
  if (isNaN(a) || isNaN(b)) return null;
  const m = (b - a) / 60000;
  return m >= 0 ? m : null;
}

function normalizeRow(r) {
  // ── Core time fields ──────────────────────────────────────────
  // Live sheet uses camelCase; BREAKAGE_HISTORY uses spaced names
  const ar41ScanTime          = fv(r, 'AR41ScanTime',          'AR41 Scan Time');
  const brkTableScanTime      = fv(r, 'BrkTableScanTime',      'Scan to Breakage Table');
  const breakageProcessedTime = fv(r, 'BreakageProcessedTime', 'Breakage Processed Time');
  const orbScanTime           = fv(r, 'ORBScanTime',           'ORB Scan Time');
  const otbScanTime           = fv(r, 'OTBScanTime',           'OTB Blocker Scan Time');

  // ── Timing mins — use pre-calculated if present, else compute ──
  const ar41ToBreakage = parseFloat(
    fv(r, 'AR41_to_Breakage_Min', 'AR41 to Breakage Min')
  ) || calcMin(ar41ScanTime, brkTableScanTime);

  const brkToProcessed = parseFloat(
    fv(r, 'Breakage_to_Processed_Min', 'Breakage to Processed Min')
  ) || calcMin(brkTableScanTime, breakageProcessedTime);

  return {
    // ── Identity ──────────────────────────────────────────────
    RxNumber:     fv(r, 'RxNumber', 'RX Number'),

    // ── Machines ──────────────────────────────────────────────
    // BREAKAGE_HISTORY: AR41 Scan Point = AR41Machine
    //                   Brk Scan Point  = BrkSourceMachine (the ORB that produced)
    //                   ORB Scan Point  = ORBMachine
    //                   OTB Blocker Scan Point = OTBMachine
    AR41Machine:      fv(r, 'AR41Machine',      'AR41 Scan Point'),
    BrkSourceMachine: fv(r, 'BrkSourceMachine', 'Brk Scan Point'),
    ORBMachine:       fv(r, 'ORBMachine',        'ORB Scan Point'),
    OTBMachine:       fv(r, 'OTBMachine',        'OTB Blocker Scan Point'),

    // ── Reason / Material ──────────────────────────────────────
    // Live sheets use 'Material'; BREAKAGE_HISTORY uses 'Lens Material'
    BrkReason:    fv(r, 'BrkReason',    'Brk Reason'),
    Material:     fv(r, 'Material', 'Lens Material', 'material') || 'Unknown',
    LensesBroken: parseFloat(fv(r, 'LensesBroken', 'Lenses Broken')) || 0,

    // ── Times ──────────────────────────────────────────────────
    AR41ScanTime:          ar41ScanTime,
    BrkTableScanTime:      brkTableScanTime,
    BreakageProcessedTime: breakageProcessedTime,
    ORBScanTime:           orbScanTime,
    OTBScanTime:           otbScanTime,

    // ── Calculated timing ─────────────────────────────────────
    AR41_to_Breakage_Min:      ar41ToBreakage,
    Breakage_to_Processed_Min: brkToProcessed,

    // ── Operator ──────────────────────────────────────────────
    // Sheet column "Operator Scan" = the AR41 operator who scanned this job
    OperatorScan: fv(r, 'OperatorScan', 'Operator Scan'),

    // ── Research / Rx fields ──────────────────────────────────
    // BREAKAGE_HISTORY: 'Lens Material'=Material, 'Lens Option'=LensOption, 'Base Curve'=BaseCurve
    LensOption: fv(r, 'LensOption', 'Lens Option'),
    BaseCurve:  fv(r, 'BaseCurve',  'Base Curve'),
    R_Sph:      fv(r, 'R_Sph', 'R Sph'),
    L_Sph:      fv(r, 'L_Sph', 'L Sph'),
    R_Cyl:      fv(r, 'R_Cyl', 'R Cyl'),
    L_Cyl:      fv(r, 'L_Cyl', 'L Cyl'),
    R_AddPower: fv(r, 'R_AddPower', 'R AddPower', 'R Add Power'),
  };
}

/* ============================================================
   APPLY DATE FILTER  (single date → sends same value as start+end)
   Re-fetches from Apps Script — reads BREAKAGE_HISTORY server-side
============================================================ */
async function applyDateFilter() {
  const dateVal = document.getElementById('dateSingle')?.value;  // "YYYY-MM-DD"
  if (!dateVal) return;

  ovBanner_('updating', `Loading ${ovPrettyDate_(dateVal)}…`);   // was showOverlay(): splash is first-load only
  document.getElementById('liveStatus').textContent   = 'Loading...';
  document.getElementById('liveDot').style.animation  = 'none';
  document.getElementById('liveDot').style.background = CLR.peri;

  // Show Historical badge
  const badge = document.getElementById('dateModeBadge');
  const label = document.getElementById('dateModeLabel');
  if (badge) badge.style.display = 'flex';
  if (label) label.textContent   = dateVal;

  try {
    setLoadStep('ls-brk', 'active');
    const brk = await fetchTab('breakageSummary', dateVal, dateVal);
    setLoadStep('ls-brk', 'done'); setLoadStep('ls-reason', 'active');

    const reason = await fetchTab('reasonSummary', dateVal, dateVal);
    setLoadStep('ls-reason', 'done'); setLoadStep('ls-research', 'active');

    const research = await fetchTab('powerResearch', dateVal, dateVal);
    setLoadStep('ls-research', 'done'); setLoadStep('ls-anom', 'active');

    const anom = await fetchTab('anomalies', dateVal, dateVal);
    setLoadStep('ls-anom', 'done');

    // Normalize all rows so field names are consistent regardless of source
    _brk      = (Array.isArray(brk)      ? brk      : []).map(normalizeRow);
    _reason   = Array.isArray(reason)   ? reason   : [];  // already aggregated by Apps Script
    _anom     = (Array.isArray(anom)     ? anom     : []).map(normalizeRow);

    // BREAKAGE_HISTORY has no Rx prescription fields (R_Sph, L_Sph, etc.)
    // For historical Research tab, reuse _brk rows — they have machine/reason/material
    // Rx charts will show empty (no prescription data in history) but table will work
    _research = _brk;

    // RX Flow follows the header date. History has no Rx detail, so there are
    // no research rows; source scan time comes from the history row itself.
    _flowBrk      = [..._brk];
    _flowResearch = [];

    document.getElementById('dateFilterCount').textContent = _brk.length;
    buildOverview();
    populateFlowFilters(); renderFlow();
    populateResearchFilters(); renderResearch();
    buildAlerts();
    buildWeekOptions();

    document.getElementById('liveStatus').textContent   = 'Historical';
    document.getElementById('liveDot').style.background = CLR.peri;
    document.getElementById('liveDot').style.animation  = 'none';
    ovBanner_('ok', `Loaded ${ovPrettyDate_(dateVal)}`);

  } catch (err) {
    console.error('Date filter fetch error:', err);
    document.getElementById('liveStatus').textContent   = 'Error';
    document.getElementById('liveDot').style.background = '#f87171';
    ovBanner_('fail', `Couldn't load ${ovPrettyDate_(dateVal)} — panels still show the previous view`);
  }
}

/* ============================================================
   CLEAR DATE FILTER — back to live
============================================================ */
async function clearDateFilter() {
  const inp = document.getElementById('dateSingle');
  if (inp) inp.value = '';
  const badge = document.getElementById('dateModeBadge');
  if (badge) badge.style.display = 'none';
  await loadAll();
}

/* ============================================================
   EXPORT HELPERS
============================================================ */
function getFilteredResearchData() {
  const fO   = document.getElementById('resORB').value;
  const fR   = document.getElementById('resReason').value;
  const fM   = document.getElementById('resMaterial').value;
  const isHistorical = (document.getElementById('dateSingle')?.value || '') !== '';
  const from = isHistorical ? null : parseDateInput('resDateFrom');
  const to   = isHistorical ? null : parseDateInput('resDateTo');

  let data = [..._research];
  if (fO) data = data.filter(r => (r.BrkSourceMachine || r.ORBMachine) === fO);
  if (fR) data = data.filter(r => r.BrkReason === fR);
  if (from || to) data = data.filter(r => inDateRange(r.ORBScanTime || r.OTBScanTime, from, to));

  // Enrich Material from _brk (BreakageSummary has Material; PowerResearch may not)
  const brkMatMap = {};
  _brk.forEach(r => { if (r.RxNumber && r.Material && r.Material !== 'Unknown') brkMatMap[String(r.RxNumber)] = r.Material; });
  data = data.map(r => {
    if (r.Material && r.Material !== 'Unknown') return r;
    const mat = brkMatMap[String(r.RxNumber)];
    return mat ? { ...r, Material: mat } : r;
  });

  // Apply material filter after enrichment so it works correctly
  if (fM) data = data.filter(r => r.Material === fM);
  return data;
}

const RES_COLS = [
  ['RxNumber', 'RX'],   ['ORBMachine', 'ORB'],    ['BrkSourceMachine', 'BrkSrc'],
  ['BrkReason', 'Reason'], ['Material', 'Material'], ['LensOption', 'Option'],
  ['BaseCurve', 'Curve'],  ['R_Sph', 'R Sph'],      ['L_Sph', 'L Sph'],
  ['R_Cyl', 'R Cyl'],      ['L_Cyl', 'L Cyl'],      ['R_AddPower', 'Add'],
  ['LensesBroken', 'Broken'],
];

function exportResearchCSV() {
  const data = getFilteredResearchData();
  if (!data.length) { alert('No data to export.'); return; }
  const header = RES_COLS.map(c => c[1]).join(',');
  const rows   = data.map(r =>
    RES_COLS.map(([key]) => {
      const v = r[key] ?? '';
      const s = String(v);
      return s.includes(',') || s.includes('"') ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(',')
  );
  const csv  = [header, ...rows].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = `power_research_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

function exportResearchXLSX() {
  const data = getFilteredResearchData();
  if (!data.length) { alert('No data to export.'); return; }

  const esc         = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const numericCols = new Set(['R_Sph', 'L_Sph', 'R_Cyl', 'L_Cyl', 'R_AddPower', 'LensesBroken']);

  const headerRow = '<Row>' + RES_COLS.map(([, label]) =>
    `<Cell><Data ss:Type="String">${esc(label)}</Data></Cell>`).join('') + '</Row>';

  const dataRows = data.map(r =>
    '<Row>' + RES_COLS.map(([key]) => {
      const v = r[key] ?? '';
      const n = parseFloat(v);
      if (numericCols.has(key) && !isNaN(n)) return `<Cell><Data ss:Type="Number">${n}</Data></Cell>`;
      return `<Cell><Data ss:Type="String">${esc(v)}</Data></Cell>`;
    }).join('') + '</Row>'
  ).join('');

  const xml = `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Worksheet ss:Name="Power Research">
  <Table>${headerRow}${dataRows}</Table>
 </Worksheet>
</Workbook>`;

  const blob = new Blob([xml], { type: 'application/vnd.ms-excel;charset=utf-8;' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = `power_research_${new Date().toISOString().slice(0, 10)}.xls`;
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ============================================================
   LOAD ALL  (live — no date params sent to Apps Script)
============================================================ */
async function loadAll() {
  const loadId = ++powerPerfState.loadSeq;
  const isInitialLoad = powerPerfState.initialLoadId === null;
  if (isInitialLoad) powerPerfState.initialLoadId = loadId;

  const totalStartedAt = powerPerfStart_(
    `${isInitialLoad ? 'INITIAL ' : ''}Power Breakage load #${loadId}`,
    {
      requests: ['liveBundle'],
      requestMode: 'single-live-bundle'
    }
  );

  if (isInitialLoad) showOverlay();               // splash on first load only
  else ovBanner_('updating', 'Updating…');         // refresh: non-blocking banner
  document.getElementById('liveStatus').textContent   = 'Loading...';
  document.getElementById('liveDot').style.background = CLR.peri;

  try {
    /*
      POWER BREAKAGE STAGE 3:
      One live browser request. The backend reads the four existing
      Stage 2 live cache keys and combines them only for this HTTP response.
      No combined CacheService value is created.
    */
    setLoadStep('ls-brk', 'active');
    setLoadStep('ls-reason', 'active');
    setLoadStep('ls-research', 'active');
    setLoadStep('ls-anom', 'active');

    const liveBundle = await fetchLiveBundle();

    const brk      = Array.isArray(liveBundle.breakageSummary) ? liveBundle.breakageSummary : [];
    const reason   = Array.isArray(liveBundle.reasonSummary)   ? liveBundle.reasonSummary   : [];
    const research = Array.isArray(liveBundle.powerResearch)   ? liveBundle.powerResearch   : [];
    const anom     = Array.isArray(liveBundle.anomalies)       ? liveBundle.anomalies       : [];

    setLoadStep('ls-brk', 'done');
    setLoadStep('ls-reason', 'done');
    setLoadStep('ls-research', 'done');
    setLoadStep('ls-anom', 'done');

    const normalizeStartedAt = powerPerfStart_(
      `Normalize datasets #${loadId}`,
      {}
    );

    _allBrk      = (Array.isArray(brk)      ? brk      : []).map(normalizeRow);
    _allReason   =  Array.isArray(reason)   ? reason   : [];
    _allResearch = (Array.isArray(research) ? research : []).map(normalizeRow);
    _allAnom     = (Array.isArray(anom)     ? anom     : []).map(normalizeRow);

    powerPerfEnd_(
      `Normalize datasets #${loadId}`,
      normalizeStartedAt,
      {
        breakageRows: _allBrk.length,
        reasonRows: _allReason.length,
        researchRows: _allResearch.length,
        anomalyRows: _allAnom.length
      }
    );

    const dateVal = document.getElementById('dateSingle')?.value;

    if (dateVal) {
      const historicalStartedAt = powerPerfStart_(
        `Existing date filter reload #${loadId}`,
        { dateVal }
      );
      await applyDateFilter();
      powerPerfEnd_(
        `Existing date filter reload #${loadId}`,
        historicalStartedAt,
        {}
      );
    } else {
      const renderStartedAt = powerPerfStart_(
        `Power Breakage build/render #${loadId}`,
        {}
      );

      _brk      = [..._allBrk];
      _reason   = [..._allReason];
      _research = [..._allResearch];
      _anom     = [..._allAnom];

      document.getElementById('dateFilterCount').textContent = _brk.length;
      buildOverview();
      _flowBrk = [..._allBrk];
      _flowResearch = [..._allResearch];
      populateFlowFilters();
      renderFlow();
      populateResearchFilters();
      renderResearch();
      buildAlerts();
      buildWeekOptions();

      document.getElementById('liveStatus').textContent   = 'Live';
      document.getElementById('liveDot').style.background = '#d4c0a8';
      document.getElementById('liveDot').style.animation  = 'pulse 1.5s infinite';
      if (!isInitialLoad) ovBanner_('ok', 'New data');

      powerPerfEnd_(
        `Power Breakage build/render #${loadId}`,
        renderStartedAt,
        {
          breakageRows: _brk.length,
          reasonRows: _reason.length,
          researchRows: _research.length,
          anomalyRows: _anom.length
        }
      );
    }

    const totalMs = powerPerfEnd_(
      `${isInitialLoad ? 'INITIAL ' : ''}Power Breakage load #${loadId}`,
      totalStartedAt,
      { success: true }
    );

    if (isInitialLoad) {
      powerPerfState.initialDataReadyAt = powerPerfNow_();
      powerPerfLog_(
        `Data/render ready for initial Power Breakage load #${loadId}; existing code waits 300 ms before hideOverlay(), then hideOverlay waits 400 ms`,
        {
          apiAndRenderMs: Number(totalMs.toFixed(1)),
          preHideDelayMs: 300,
          overlayHideDelayMs: 400
        }
      );

      powerPerfLog_(
        'Main dashboard complete; starting deferred keep-warm ping',
        {}
      );
      pingWarm();
    }

  } catch (err) {
    console.error(err);
    document.getElementById('liveStatus').textContent   = 'Error';
    document.getElementById('liveDot').style.background = '#f87171';
    document.getElementById('liveDot').style.animation  = 'none';
    if (!isInitialLoad) ovBanner_('fail', 'Update failed — still showing the last good data');

    powerPerfEnd_(
      `${isInitialLoad ? 'INITIAL ' : ''}Power Breakage load #${loadId}`,
      totalStartedAt,
      {
        success: false,
        message: err?.message || String(err)
      }
    );

    if (isInitialLoad) {
      powerPerfLog_(
        'Main dashboard failed; starting deferred keep-warm ping after main request settled',
        {}
      );
      pingWarm();
    }
  } finally {
    if (isInitialLoad) setTimeout(hideOverlay, 300);
  }
}

/* ============================================================
   KEEP-WARM PING
   Fires a cheap request to Apps Script on page load so the
   V8 runtime is already warm when loadAll() fetches real data.
   Also fires every 4 minutes in sync with the server-side trigger.
============================================================ */
function pingWarm() {
  const startedAt = powerPerfStart_('Keep-warm ping', {});
  fetch(`${API}?tab=keepwarm`)
    .then(r => {
      powerPerfEnd_('Keep-warm ping', startedAt, {
        httpStatus: r.status,
        ok: r.ok
      });
    })
    .catch(err => {
      powerPerfEnd_('Keep-warm ping', startedAt, {
        success: false,
        message: err?.message || String(err)
      });
    });
}

/*
  POWER BREAKAGE STAGE 1:
  Initial keep-warm is deferred until the main dashboard finishes.
  The recurring 4-minute keep-warm remains unchanged.
*/
setInterval(pingWarm, 4 * 60 * 1000);
loadAll();

/* ============================================================
   SUMMARY v6 — management brief (one page, sidebar = jump links)
   Data loading (approved):
   - Dates are sent to the API as ET calendar days (no UTC shift).
   - Past days come from history; today (if in range) from the live report.
   - The previous period is fetched for comparison. When the current
     period is still in progress, it is compared to the same point
     in the previous period, never to a full one.
============================================================ */
const SUM = {
  mode: 'week', weeks: [], weekKey: null, cur: null, prev: null,
  token: 0, entrancePlayed: false, cache: new Map(), text: ''
};
const SUM_CACHE_MS = 10 * 60 * 1000;

function sumTodayYmd_() { return etYmd_(new Date()); }
function sumDow_(ymd) { return new Date(`${ymd}T12:00:00Z`).getUTCDay(); }
function sumMondayOf_(ymd) { const d = sumDow_(ymd); return addDaysYmd_(ymd, d === 0 ? -6 : 1 - d); }
function sumFmt_(ymd, opts) { return new Intl.DateTimeFormat('en-US', { timeZone: 'UTC', ...opts }).format(new Date(`${ymd}T12:00:00Z`)); }
function sumWeekLabel_(from, to) {
  return `${sumFmt_(from, { month: 'short', day: 'numeric' })} – ${sumFmt_(to, { month: 'short', day: 'numeric' })}, ${sumFmt_(to, { year: 'numeric' })}`;
}
function sumVisible_() { return document.getElementById('t-summary')?.classList.contains('active'); }
function sumRowDate_(r) { return safeIsoDate(r.AR41ScanTime || r.BrkTableScanTime) || ''; }
function sumRowHour_(r) {
  const t = r.AR41ScanTime || r.BrkTableScanTime;
  const d = t ? new Date(t) : null;
  return d && !isNaN(d) ? d.getHours() : null;
}

/* ── Period definitions ───────────────────────────────────── */
function sumPeriods_() {
  const today = sumTodayYmd_();
  if (SUM.mode === 'day') {
    const day = document.getElementById('daySelect')?.value || today;
    const prevDay = addDaysYmd_(day, -1);
    return {
      cur:  { from: day, to: day, label: sumFmt_(day, { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' }) },
      prev: { from: prevDay, to: prevDay, label: sumFmt_(prevDay, { weekday: 'short', month: 'short', day: 'numeric' }) },
      inProgress: day === today, today
    };
  }
  const w = SUM.weeks[Number(document.getElementById('weekSelect')?.value) || 0] || SUM.weeks[0];
  const pf = addDaysYmd_(w.from, -7), pt = addDaysYmd_(w.to, -7);
  return {
    cur:  { from: w.from, to: w.to, label: sumWeekLabel_(w.from, w.to) },
    prev: { from: pf, to: pt, label: sumWeekLabel_(pf, pt) },
    inProgress: w.from <= today && today <= w.to, today
  };
}

/* ── Loading ──────────────────────────────────────────────── */
async function sumHistory_(from, to) {
  const key = `${from}|${to}`;
  const hit = SUM.cache.get(key);
  if (hit && Date.now() - hit.at < SUM_CACHE_MS) return hit.rows;
  const raw = await fetchTab('breakageSummary', from, to);   // ET calendar days, inclusive
  const rows = (Array.isArray(raw) ? raw : []).map(normalizeRow);
  SUM.cache.set(key, { at: Date.now(), rows });
  return rows;
}

async function sumRowsFor_(p, today) {
  const rows = [];
  const histTo = p.to < today ? p.to : addDaysYmd_(today, -1);
  if (p.from <= histTo) rows.push(...await sumHistory_(p.from, histTo));
  if (p.from <= today && today <= p.to) {
    // today comes from the live report (same rule the old Summary used)
    rows.push(..._allBrk.filter(r => sumRowDate_(r) === today));
  }
  return rows;
}

async function sumLoad_() {
  const token = ++SUM.token;
  const P = sumPeriods_();
  const main = document.getElementById('sumMain');
  main?.classList.add('is-loading');
  ovBanner_('updating', `Loading ${P.cur.label}…`);
  try {
    const [cur, prev] = await Promise.all([sumRowsFor_(P.cur, P.today), sumRowsFor_(P.prev, P.today)]);
    if (token !== SUM.token) return;   // a newer request won

    // Fair comparison while the current period is still running
    let prevCmp = prev, cmpNote = '';
    if (P.inProgress && SUM.mode === 'week') {
      const cutoff = addDaysYmd_(P.today, -7);
      prevCmp = prev.filter(r => { const d = sumRowDate_(r); return d && d <= cutoff; });
      cmpNote = 'the same days last week';
    } else if (P.inProgress && SUM.mode === 'day') {
      const hr = Number(new Intl.DateTimeFormat('en-US', { timeZone: 'America/New_York', hour: 'numeric', hour12: false }).format(new Date())) % 24;
      prevCmp = prev.filter(r => { const h = sumRowHour_(r); return h !== null && h <= hr; });
      cmpNote = 'the same hours yesterday';
    }

    const prevCur = SUM.cur;
    SUM.cur  = sumStats_(cur, P, false);
    SUM.prev = sumStats_(prevCmp, P, true);
    SUM.prevFull = sumStats_(prev, P, true);
    SUM.P = P; SUM.cmpNote = cmpNote || (SUM.mode === 'week' ? 'last week' : 'the day before');
    sumRender_(prevCur);
    ovBanner_('ok', `Loaded ${P.cur.label}`);
  } catch (err) {
    console.error('Summary load error:', err);
    if (token !== SUM.token) return;
    ovBanner_('fail', `Couldn't load ${P.cur.label} — showing the previous view`);
  } finally {
    if (token === SUM.token) main?.classList.remove('is-loading');
  }
}

/* ── Stats ────────────────────────────────────────────────── */
function sumStats_(rows, P, isPrev) {
  const lens = r => parseInt(r.LensesBroken) || 0;
  const agg = keyFn => {
    const m = {};
    rows.forEach(r => { const k = keyFn(r) || 'Unknown'; if (!m[k]) m[k] = { jobs: 0, lens: 0 }; m[k].jobs++; m[k].lens += lens(r); });
    return m;
  };
  const timing = f => {
    const v = rows.map(r => parseFloat(r[f])).filter(Number.isFinite);
    return { n: v.length, med: flowMedian_(v) };
  };
  // day-by-day (week) or hour-by-hour (day); future buckets stay null (no fake zeros)
  const buckets = [];
  if (SUM.mode === 'week') {
    const base = isPrev ? P.prev.from : P.cur.from;
    for (let i = 0; i < 7; i++) {
      const d = addDaysYmd_(base, i);
      const future = !isPrev && d > P.today;
      buckets.push({ key: d, label: sumFmt_(d, { weekday: 'short' }), v: future ? null : rows.filter(r => sumRowDate_(r) === d).length });
    }
  } else {
    const nowHr = Number(new Intl.DateTimeFormat('en-US', { timeZone: 'America/New_York', hour: 'numeric', hour12: false }).format(new Date())) % 24;
    for (let h = 0; h < 24; h++) {
      const future = !isPrev && P.inProgress && h > nowHr;
      buckets.push({ key: h, label: h === 0 ? '12a' : h < 12 ? `${h}a` : h === 12 ? '12p' : `${h - 12}p`, v: future ? null : rows.filter(r => sumRowHour_(r) === h).length });
    }
  }
  return {
    rows, jobs: rows.length, lens: rows.reduce((s, r) => s + lens(r), 0),
    reasons: agg(r => r.BrkReason), machines: agg(r => r.BrkSourceMachine), materials: agg(r => r.Material),
    s2: timing('AR41_to_Breakage_Min'), s3: timing('Breakage_to_Processed_Min'),
    buckets,
    sheetErr: rows.filter(r => [r.BrkSourceMachine, r.BrkReason, r.Material, r.RxNumber].some(v => OV_SHEET_ERR.test(String(v || '').trim()))).length,
    undated: rows.filter(r => !sumRowDate_(r)).length,
  };
}

const sumSorted_ = (m, by = 'jobs') => Object.entries(m).sort((a, b) => b[1][by] - a[1][by] || a[0].localeCompare(b[0]));
const sumPct_ = (a, b) => b ? Math.round(a / b * 100) : 0;

// ▲/▼ chip. higherIsWorse: more breakage or longer time = red
function sumDelta_(cur, prev, { unit = '', time = false } = {}) {
  if (prev === null || prev === undefined || cur === null || cur === undefined) return '<span class="sum-same">no comparison</span>';
  const d = cur - prev;
  if (Math.abs(d) < (time ? 1 : 0.5)) return '<span class="sum-same">same as before</span>';
  const txt = time ? fmtMin(Math.abs(d)) : `${Math.abs(Math.round(d))}${unit}`;
  return d > 0
    ? `<span class="sum-up">▲ ${txt}${time ? ' slower' : ''}</span>`
    : `<span class="sum-down">▼ ${txt}${time ? ' faster' : ''}</span>`;
}

/* ── Render ───────────────────────────────────────────────── */
function sumRender_(prevCur) {
  const C = SUM.cur, R = SUM.prev, P = SUM.P;
  const isWeek = SUM.mode === 'week';
  const per  = isWeek ? (P.inProgress ? 'so far this week' : 'this week') : (P.inProgress ? 'so far today' : `on ${P.cur.label}`);
  const anim = !ovReduced_() && sumVisible_();

  // sidebar + range text
  const ml = document.getElementById('sbModeLabel'); if (ml) ml.textContent = isWeek ? 'Weekly' : 'Daily';
  const dl = document.getElementById('sbDateLabel'); if (dl) dl.textContent = P.cur.label + (P.inProgress ? ' (in progress)' : '');
  const rg = document.getElementById('sumWeekRange'); if (rg) rg.textContent = `compared with ${P.prev.label}${P.inProgress ? ' · same point in time' : ''}`;

  // ── Hero ──
  const rs = sumSorted_(C.reasons), ms = sumSorted_(C.machines);
  const head = document.getElementById('sumHeadline');
  const sub  = document.getElementById('sumHeadSub');
  if (!C.jobs) {
    head.innerHTML = `No breakage recorded ${per}.`;
    sub.textContent = R.jobs ? `${SUM.cmpNote[0].toUpperCase() + SUM.cmpNote.slice(1)} had ${R.jobs} jobs.` : '';
  } else {
    let chg = '';
    if (R.jobs) {
      const pct = Math.round((C.jobs - R.jobs) / R.jobs * 100);
      chg = pct === 0 ? `, the same as ${SUM.cmpNote}` : `, <span class="${pct > 0 ? 'sum-up' : 'sum-down'}">${pct > 0 ? 'up' : 'down'} ${Math.abs(pct)}% from ${SUM.cmpNote}</span>`;
    } else if (R.rows) {
      chg = `, with none in ${SUM.cmpNote}`;
    }
    head.innerHTML = `${C.jobs} ${C.jobs === 1 ? 'job' : 'jobs'} broke ${per} (${C.lens} ${C.lens === 1 ? 'lens' : 'lenses'})${chg}. ` +
      `${ovEsc_(rs[0][0])} is ${sumPct_(rs[0][1].jobs, C.jobs)}% of it.`;
    let s = '';
    if (ms.length >= 2) s = `${ovEsc_(ms[0][0])} and ${ovEsc_(ms[1][0])} account for ${sumPct_(ms[0][1].jobs + ms[1][1].jobs, C.jobs)}% of jobs — the first places to look.`;
    else if (ms.length === 1) s = `All of it came from ${ovEsc_(ms[0][0])}.`;
    if (C.jobs < OV_SMALL_N) s += ` Only ${C.jobs} jobs, so read this as a hint, not a trend.`;
    sub.innerHTML = s;
  }

  // ── KPIs ──
  const setNum = (id, v) => { const el = document.getElementById(id); if (!el) return; const was = Number(el.dataset.v); ovCount_(el, v, anim && Number.isFinite(was) ? was : v); };
  setNum('sumKJobs', C.jobs);
  document.getElementById('sumKJobsD').innerHTML = sumDelta_(C.jobs, R.jobs) + ` <span class="sum-vs">vs ${ovEsc_(SUM.cmpNote)}</span>`;
  setNum('sumKLens', C.lens);
  document.getElementById('sumKLensD').innerHTML = `${C.jobs ? (C.lens / C.jobs).toFixed(1) : '0'} per job · ` + sumDelta_(C.lens, R.lens);
  document.getElementById('sumKReason').textContent = rs[0]?.[0] || '--';
  document.getElementById('sumKReasonD').textContent = rs[0] ? `${rs[0][1].jobs} of ${C.jobs} jobs · ${sumPct_(rs[0][1].jobs, C.jobs)}%` : 'no breakage';
  document.getElementById('sumKProc').textContent = C.s3.n ? fmtMin(C.s3.med) : '--';
  document.getElementById('sumKProcD').innerHTML = C.s3.n
    ? `median of ${C.s3.n} timed jobs · ${sumDelta_(C.s3.med, R.s3.n ? R.s3.med : null, { time: true })}`
    : 'no job has a processed time';

  // ── What to look at ──
  document.getElementById('sumLook').innerHTML = sumLookItems_(C, R, P).map((it, i) => `
    <div class="sum-act"><div class="sum-actn">${i + 1}</div><div>
      <b>${it.title}</b><p>${it.text}</p><div class="sum-ev">evidence: ${it.ev}</div></div></div>`).join('')
    || '<div class="sum-empty">Nothing stands out — no breakage to review.</div>';

  // ── Reasons ──
  sumBars_('sumReasonBars', rs.map(([k, v]) => ({ k, n: v.jobs, pct: sumPct_(v.jobs, C.jobs), color: reasonColor(k), dot: true })), anim);
  document.getElementById('sumReasonCheck').textContent = rs.length
    ? `${rs.map(e => e[1].jobs).join(' + ')} = ${C.jobs}${rs.reduce((s, e) => s + e[1].jobs, 0) === C.jobs ? ' ✓' : ' ✗'}` : '';

  // ── Machines (top 4 + others) ──
  const top = ms.slice(0, 4), rest = ms.slice(4);
  const mRows = top.map(([k, v]) => ({ k, n: v.jobs, delta: sumDelta_(v.jobs, R.machines[k]?.jobs ?? 0), mono: true, color: CLR.teal }));
  if (rest.length) {
    const n = rest.reduce((s, e) => s + e[1].jobs, 0);
    const pn = Object.entries(R.machines).filter(([k]) => !top.some(t => t[0] === k)).reduce((s, e) => s + e[1].jobs, 0);
    mRows.push({ k: `Others (${rest.length})`, n, delta: sumDelta_(n, pn), color: 'rgba(212,192,168,.4)' });
  }
  sumBars_('sumMachineBars', mRows, anim);

  // ── Materials (lenses) ──
  const mat = sumSorted_(C.materials, 'lens');
  sumBars_('sumMaterialBars', mat.map(([k, v]) => ({ k, n: v.lens, pct: sumPct_(v.lens, C.lens), color: CLR.teal })), anim);
  document.getElementById('sumMatCheck').textContent = mat.length ? `lenses by material · adds to ${mat.reduce((s, e) => s + e[1].lens, 0)}${mat.reduce((s, e) => s + e[1].lens, 0) === C.lens ? ' ✓' : ' ✗'}` : 'lenses by material';

  // ── Timing ──
  const tbox = (id, st, pst) => {
    document.getElementById(id).innerHTML = st.n
      ? `<div class="sum-tn">${fmtMin(st.med)}</div><div class="sum-ts">median · timed on ${st.n} of ${C.jobs}${st.n < OV_SMALL_N ? ' — too few to trust as a trend' : ''}</div><div class="sum-ts">${sumDelta_(st.med, pst.n ? pst.med : null, { time: true })}</div>`
      : `<div class="sum-tn">--</div><div class="sum-ts">no job has this timing (0 of ${C.jobs})</div>`;
  };
  tbox('sumT2', C.s2, R.s2);
  tbox('sumT3', C.s3, R.s3);

  // ── Chart ──
  sumChart_(C.buckets, (P.inProgress ? SUM.prevFull : R).buckets, anim && !SUM.entrancePlayed);

  // ── Data notes ──
  const notes = [];
  if (C.sheetErr) notes.push(`${C.sheetErr} ${C.sheetErr === 1 ? 'row has' : 'rows have'} a sheet error (#N/A) and ${C.sheetErr === 1 ? 'is' : 'are'} still counted`);
  if (C.undated) notes.push(`${C.undated} ${C.undated === 1 ? 'job has' : 'jobs have'} no scan date and ${C.undated === 1 ? "isn't" : "aren't"} on the ${isWeek ? 'day-by-day' : 'hour-by-hour'} chart`);
  notes.push(isWeek ? 'week runs Mon 12:00 AM → Sun 11:59 PM ET' : 'day runs 12:00 AM → 11:59 PM ET');
  if (P.inProgress) notes.push('period still in progress — compared to the same point last time');
  notes.push('▲ red = more breakage / slower · ▼ green = less / faster');
  document.getElementById('sumDataNotes').innerHTML = `<b>Data notes:</b> ${notes.join(' · ')}.`;

  // ── Report text (copy / export) ──
  SUM.text = sumPlainText_();
  document.getElementById('sumReportText').textContent = SUM.text;
  const ex = document.getElementById('sumExportBtn'); if (ex) ex.disabled = false;
  const cp = document.getElementById('sumCopyBtn');   if (cp) cp.disabled = false;

  if (sumVisible_() && !SUM.entrancePlayed) sumEntrance_();
}

function sumLookItems_(C, R, P) {
  const out = [];
  if (!C.jobs) return out;
  const small = C.jobs < OV_SMALL_N ? ' · small sample' : '';
  const prevWord = SUM.cmpNote;

  // 1) top machine
  const ms = sumSorted_(C.machines);
  const pms = sumSorted_(R.machines);
  if (ms.length) {
    const [m, v] = ms[0];
    const again = pms[0]?.[0] === m;
    out.push({
      title: `${ovEsc_(m)} had the most breakage`,
      text: again ? 'Highest machine again — it was also the top machine in the comparison period. Start the review here.' : 'Start the review here.',
      ev: `${v.jobs} of ${C.jobs} jobs (${sumPct_(v.jobs, C.jobs)}%) · ${ovEsc_(prevWord)} ${R.machines[m]?.jobs ?? 0}${small}`
    });
  }
  // 2) reason that rose the most
  const rise = sumSorted_(C.reasons).map(([k, v]) => ({ k, d: v.jobs - (R.reasons[k]?.jobs ?? 0), v: v.jobs, p: R.reasons[k]?.jobs ?? 0 }))
    .filter(x => x.d >= 3).sort((a, b) => b.d - a.d)[0];
  if (rise) out.push({ title: `${ovEsc_(rise.k)} rose the most`, text: `More ${ovEsc_(rise.k)} breakage than ${ovEsc_(prevWord)}.`, ev: `${rise.p} → ${rise.v} jobs (+${rise.d})${small}` });

  // 3) material whose share of lenses grew the most (≥ 3 points), else the top material
  const mt = sumSorted_(C.materials, 'lens').map(([k, v]) => {
    const share = sumPct_(v.lens, C.lens), pShare = R.lens ? sumPct_(R.materials[k]?.lens ?? 0, R.lens) : null;
    return { k, lens: v.lens, share, pShare, d: pShare === null ? null : share - pShare };
  });
  const grow = mt.filter(x => x.d !== null && x.d >= 3).sort((a, b) => b.d - a.d)[0];
  if (grow) out.push({ title: `${ovEsc_(grow.k)} lenses`, text: 'Share of broken lenses rose the most of any material.', ev: `${grow.lens} lenses (${grow.share}%) · ${ovEsc_(prevWord)} ${grow.pShare}%${small}` });
  else if (mt.length) out.push({ title: `${ovEsc_(mt[0].k)} lenses`, text: 'The most-broken material.', ev: `${mt[0].lens} of ${C.lens} lenses (${mt[0].share}%)${small}` });

  // 4) timing coverage
  if (C.s3.n < C.jobs * 0.5) out.push({
    title: 'Breakage-table scans are missing',
    text: 'Most jobs have no table or processed scan, so handling time is measured on a small share of jobs.',
    ev: `processed time on ${C.s3.n} of ${C.jobs} jobs (${sumPct_(C.s3.n, C.jobs)}%)`
  });
  return out.slice(0, 3);
}

function sumBars_(elId, rows, anim) {
  const box = document.getElementById(elId);
  if (!box) return;
  if (!rows.length) { box.innerHTML = '<div class="sum-empty">No breakage</div>'; return; }
  const max = Math.max(1, ...rows.map(r => r.n));
  const old = new Map([...box.querySelectorAll('.sum-bar')].map(el => [el.dataset.key, el.querySelector('.sum-fill').style.width]));
  box.innerHTML = rows.map(r => `
    <div class="sum-bar" data-key="${ovEsc_(r.k)}">
      <span class="sum-bl${r.mono ? ' mono' : ''}">${r.dot ? `<i style="background:${r.color}"></i>` : ''}${ovEsc_(r.k)}</span>
      <div class="sum-track"><div class="sum-fill" style="width:${r.n / max * 100}%;background:${r.color}"></div></div>
      <b>${r.n}</b>
      <small>${r.delta !== undefined ? r.delta : `${r.pct}%`}</small>
    </div>`).join('');
  if (!anim) return;
  box.querySelectorAll('.sum-bar').forEach(el => {
    const f = el.querySelector('.sum-fill'), to = f.style.width, from = old.get(el.dataset.key);
    if (from !== undefined && from !== to) f.animate?.([{ width: from }, { width: to }], { duration: 600, easing: 'ease-out' });
  });
}

function sumChart_(cur, prev, draw) {
  const svg = document.getElementById('sumChart');
  if (!svg) return;
  const W = 760, H = 240, L = 40, Rm = 14, T = 16, B = 34;
  const vals = [...cur, ...prev].map(b => b.v).filter(v => v !== null);
  const step = Math.max(1, Math.ceil(Math.max(1, ...vals) / 4));
  const ymax = step * 4;
  const n = cur.length;
  const x = i => L + (W - L - Rm) * (n === 1 ? 0.5 : i / (n - 1));
  const y = v => T + (H - T - B) * (1 - v / ymax);
  // straight segments between real points only; a null breaks the line
  const path = arr => {
    let d = '', pen = false;
    arr.forEach((b, i) => {
      if (b.v === null) { pen = false; return; }
      d += `${pen ? 'L' : 'M'}${x(i).toFixed(1)},${y(b.v).toFixed(1)} `; pen = true;
    });
    return d.trim();
  };
  let grid = '';
  for (let k = 0; k <= 4; k++) {
    const v = step * k;
    grid += `<line x1="${L}" x2="${W - Rm}" y1="${y(v)}" y2="${y(v)}" class="sum-g"/><text x="${L - 8}" y="${y(v) + 4}" text-anchor="end">${v}</text>`;
  }
  const every = n > 12 ? 3 : 1;
  const labels = cur.map((b, i) => i % every ? '' : `<text x="${x(i)}" y="${H - 10}" text-anchor="middle">${b.label}</text>`).join('');
  const dots = cur.map((b, i) => b.v === null ? '' : `<circle cx="${x(i)}" cy="${y(b.v)}" r="5" class="sum-pt"><title>${b.label}: ${b.v} jobs</title></circle>`).join('');
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.innerHTML = `${grid}${labels}
    <path d="${path(prev)}" class="sum-prev"/>
    <path d="${path(cur)}" class="sum-cur" id="sumCurLine"/>${dots}`;
  if (draw) {
    const line = document.getElementById('sumCurLine');
    const len = line.getTotalLength?.() || 0;
    if (len && line.animate) line.animate([{ strokeDasharray: `${len}`, strokeDashoffset: `${len}` }, { strokeDasharray: `${len}`, strokeDashoffset: '0' }], { duration: 1000, easing: 'ease-out' });
  }
}

function sumEntrance_() {
  SUM.entrancePlayed = true;
  if (ovReduced_()) return;
  const sec = document.getElementById('sumMain');
  if (!sec?.animate) return;
  sec.querySelectorAll('.sum-panel').forEach((p, i) => p.animate([{ opacity: 0, transform: 'translateY(10px)' }, { opacity: 1, transform: 'none' }], { duration: 420, delay: i * 60, easing: 'ease-out', fill: 'backwards' }));
  sec.querySelectorAll('.sum-fill').forEach((f, i) => { const w = f.style.width; f.animate([{ width: '0%' }, { width: w }], { duration: 700, delay: 200 + (i % 8) * 50, easing: 'cubic-bezier(.2,.7,.2,1)', fill: 'backwards' }); });
  ['sumKJobs', 'sumKLens'].forEach(id => { const el = document.getElementById(id); const v = Number(el?.dataset.v); if (Number.isFinite(v)) ovCount_(el, v, 0, 800); });
  const line = document.getElementById('sumCurLine');
  const len = line?.getTotalLength?.() || 0;
  if (len) line.animate([{ strokeDasharray: `${len}`, strokeDashoffset: `${len}` }, { strokeDasharray: `${len}`, strokeDashoffset: '0' }], { duration: 1000, delay: 250, easing: 'ease-out', fill: 'backwards' });
}

/* ── Plain text (copy for email / export) ─────────────────── */
function sumPlainText_() {
  const C = SUM.cur, R = SUM.prev, P = SUM.P;
  const strip = h => { const t = document.createElement('div'); t.innerHTML = h; return (t.textContent || '').replace(/\s+/g, ' ').trim(); };
  const lines = [
    `POWER BREAKAGE — ${SUM.mode === 'week' ? 'WEEKLY' : 'DAILY'} SUMMARY`,
    `${P.cur.label}${P.inProgress ? ' (in progress)' : ''} · compared with ${SUM.cmpNote}`,
    '',
    strip(document.getElementById('sumHeadline').innerHTML),
    strip(document.getElementById('sumHeadSub').innerHTML),
    '',
    `Jobs with breakage: ${C.jobs} (${SUM.cmpNote}: ${R.jobs})`,
    `Lenses broken:      ${C.lens} (${SUM.cmpNote}: ${R.lens})`,
    `Top reason:         ${document.getElementById('sumKReason').textContent} — ${document.getElementById('sumKReasonD').textContent}`,
    `Processing time:    ${C.s3.n ? `${fmtMin(C.s3.med)} median, timed on ${C.s3.n} of ${C.jobs}` : 'no job has a processed time'}`,
    '',
    'WHAT TO LOOK AT',
    ...[...document.querySelectorAll('#sumLook .sum-act')].map((a, i) => `${i + 1}. ${strip(a.querySelector('b').innerHTML)} — ${strip(a.querySelector('p').innerHTML)} (${strip(a.querySelector('.sum-ev').innerHTML)})`),
    '',
    'BY REASON (jobs)',
    ...sumSorted_(C.reasons).map(([k, v]) => `  ${k}: ${v.jobs}`),
    'TOP MACHINES (jobs)',
    ...sumSorted_(C.machines).slice(0, 5).map(([k, v]) => `  ${k}: ${v.jobs}`),
    'TOP MATERIALS (lenses)',
    ...sumSorted_(C.materials, 'lens').slice(0, 5).map(([k, v]) => `  ${k}: ${v.lens}`),
    '',
    strip(document.getElementById('sumDataNotes').innerHTML),
    `Generated ${fmtEtStamp_(new Date())} ET`,
  ];
  return lines.join('\n');
}

async function sumCopy_() {
  if (!SUM.text) return;
  try {
    await navigator.clipboard.writeText(SUM.text);
  } catch (e) {
    const ta = document.createElement('textarea');
    ta.value = SUM.text; document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); } finally { ta.remove(); }
  }
  ovBanner_('ok', 'Summary copied — paste it into your email');
}

function exportSummaryTXT() {
  if (!SUM.text || !SUM.P) return;
  const blob = new Blob([SUM.text], { type: 'text/plain;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `power_breakage_${SUM.mode === 'week' ? 'weekly' : 'daily'}_${SUM.P.cur.from}.txt`;   // ET date, no UTC shift
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ── Controls (names kept: called from power.html / loadAll) ── */
function setSumMode(mode) {
  SUM.mode = mode;
  document.getElementById('sumModeWeek').classList.toggle('active', mode === 'week');
  document.getElementById('sumModeDay').classList.toggle('active', mode === 'day');
  document.getElementById('sumWeekGroup').style.display = mode === 'week' ? '' : 'none';
  document.getElementById('sumDayGroup').style.display  = mode === 'day'  ? '' : 'none';
  const lk = document.getElementById('sumLookTitle');
  if (lk) lk.textContent = mode === 'week' ? 'What should we look at this week?' : 'What should we look at today?';
  const ct = document.getElementById('sumChartTitle');
  if (ct) ct.textContent = mode === 'week' ? 'Day by day' : 'Hour by hour';
  if (mode === 'day') {
    const inp = document.getElementById('daySelect');
    if (inp) { inp.max = sumTodayYmd_(); if (!inp.value) inp.value = sumTodayYmd_(); }
  }
  SUM.entrancePlayed = false;
  sumLoad_();
}

function onDayChange() { sumLoad_(); }
function onWeekChange() { sumLoad_(); }

// Called after every load (loadAll / applyDateFilter) and when the tab opens.
// Builds the 12 week options once per ET week, keeps the user's choice,
// and only fetches while the Summary tab is visible.
function buildWeekOptions() {
  const sel = document.getElementById('weekSelect');
  if (!sel) return;
  const thisMon = sumMondayOf_(sumTodayYmd_());
  if (SUM.weekKey !== thisMon) {
    const keep = sel.value ? SUM.weeks[Number(sel.value)]?.from : null;
    SUM.weekKey = thisMon;
    SUM.weeks = Array.from({ length: 12 }, (_, i) => {
      const from = addDaysYmd_(thisMon, -7 * i), to = addDaysYmd_(from, 6);
      return { from, to, label: sumWeekLabel_(from, to) + (i === 0 ? ' (in progress)' : '') };
    });
    while (sel.options.length) sel.remove(0);
    SUM.weeks.forEach((w, i) => { const o = document.createElement('option'); o.value = i; o.textContent = w.label; sel.appendChild(o); });
    const k = keep ? SUM.weeks.findIndex(w => w.from === keep) : -1;
    sel.value = k >= 0 ? k : 1;   // default: last complete week
  }
  if (sumVisible_()) sumLoad_();
}

// Sidebar: jump links + highlight follows scroll
function sumJump_(id, btn) {
  const el = document.getElementById(id);
  if (!el) return;
  el.scrollIntoView({ behavior: ovReduced_() ? 'auto' : 'smooth', block: 'start' });
  document.querySelectorAll('.sum-nav').forEach(b => b.classList.toggle('active', b === btn));
}

(function sumWatch_() {
  if (!('IntersectionObserver' in window)) return;
  const io = new IntersectionObserver(entries => {
    entries.forEach(e => {
      if (!e.isIntersecting) return;
      document.querySelectorAll('.sum-nav').forEach(b => b.classList.toggle('active', b.dataset.target === e.target.id));
    });
  }, { rootMargin: '-180px 0px -60% 0px' });
  document.querySelectorAll('#sumMain [data-nav]').forEach(s => io.observe(s));
})();

document.addEventListener('DOMContentLoaded', () => {
  const domStartedAt = powerPerfStart_('DOMContentLoaded → Power Breakage boot', {});
  powerPerfLog_(
    `DOMContentLoaded fired at ${powerPerfNow_().toFixed(1)} ms after navigation start`,
    {}
  );

  document.getElementById('summaryTabBtn')?.addEventListener('click', () => {
    buildWeekOptions();
  });

  powerPerfEnd_('DOMContentLoaded → Power Breakage boot', domStartedAt, {});
});


/* ============================================================
   HEADER (Coating-style) — UI only
   Added for the header redesign. Does not change the API URL,
   fetch order, calculations, filters, or refresh cadence.
   - ET clock (America/New_York)
   - Previous / next day stepping around the existing
     applyDateFilter() and clearDateFilter()
   - Live pill color + "Updated" chip, driven by watching the
     existing #liveStatus text (loadAll/applyDateFilter untouched)
============================================================ */
// Timezone is hard-coded per data rules: America/New_York

function etYmd_(d) {
  // "YYYY-MM-DD" in Eastern time (en-CA formats as ISO date)
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'America/New_York', year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
}

function etTime_(d) {
  return new Intl.DateTimeFormat('en-US', { timeZone: 'America/New_York', hour: 'numeric', minute: '2-digit', hour12: true }).format(d);
}

function fmtEtStamp_(d) {
  const md = new Intl.DateTimeFormat('en-US', { timeZone: 'America/New_York', month: 'numeric', day: 'numeric' }).format(d);
  return `${md} ${etTime_(d)}`;
}

function addDaysYmd_(ymd, delta) {
  // Noon UTC avoids any DST edge when stepping calendar days
  const dt = new Date(`${ymd}T12:00:00Z`);
  dt.setUTCDate(dt.getUTCDate() + delta);
  return dt.toISOString().slice(0, 10);
}

function updateEtClock_() {
  const now = new Date();
  const dEl = document.getElementById('hdrDate');
  const tEl = document.getElementById('hdrTime');
  if (dEl) dEl.textContent = new Intl.DateTimeFormat('en-US', { timeZone: 'America/New_York', weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' }).format(now);
  if (tEl) tEl.textContent = etTime_(now);
  const inp = document.getElementById('dateSingle');
  if (inp) inp.max = etYmd_(now);
}

function stepDate_(delta) {
  const inp = document.getElementById('dateSingle');
  if (!inp) return;
  const today  = etYmd_(new Date());
  const base   = inp.value || today;
  const target = addDaysYmd_(base, delta);

  if (target >= today) {
    // Stepping forward onto today = back to the live report
    if (inp.value) clearDateFilter();
    return;
  }
  inp.value = target;
  applyDateFilter();
}

function syncHeaderState_() {
  const status = (document.getElementById('liveStatus')?.textContent || '').trim().toLowerCase();
  const badge  = document.getElementById('liveBadge');
  const chip   = document.getElementById('lastUpdated');
  const nav    = document.querySelector('.nav');
  const inp    = document.getElementById('dateSingle');
  const next   = document.getElementById('dayNextBtn');

  let state = 'loading';
  if (status === 'live') state = 'live';
  else if (status === 'historical') state = 'historical';
  else if (status === 'error') state = 'error';
  if (badge) badge.dataset.state = state;

  if (chip) {
    if (state === 'live' || state === 'historical') {
      chip._lastOk = etTime_(new Date());
      chip.textContent = `Updated ${chip._lastOk}`;
      chip.classList.remove('is-failed');
    } else if (state === 'error') {
      chip.textContent = chip._lastOk ? `Update failed · last ${chip._lastOk}` : 'Update failed';
      chip.classList.add('is-failed');
    }
  }

  const isHist = !!(inp && inp.value);
  if (nav)  nav.classList.toggle('is-historical', isHist);
  if (next) next.disabled = !isHist;
}

(function initHeader_() {
  updateEtClock_();
  setInterval(updateEtClock_, 1000);

  const statusEl = document.getElementById('liveStatus');
  if (statusEl) {
    new MutationObserver(syncHeaderState_).observe(statusEl, { childList: true, characterData: true, subtree: true });
  }
  syncHeaderState_();
})();

/* ============================================================
   OVERVIEW HELPERS (v6) — rendering + animation, UI only
   Rules:
   - Entrance animation plays once, after the loading splash hides.
   - Refresh (same date): bars move old → new, numbers count old → new,
     changed values get a short white outline, jobs shows a +/- badge.
   - Date switch: quick fade-in, no "changed" highlights (different day,
     not new data).
   - prefers-reduced-motion: everything is instant.
============================================================ */
const OV_SMALL_N = 10;   // fewer than 10 timed jobs (or jobs per machine) = small sample
const _ov = { rendered: false, ctx: null, entrancePlayed: false, lastJobs: null };

function ovReduced_() {
  return window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

function ovCtxKey_() {
  return document.getElementById('dateSingle')?.value || 'live';
}

function ovRenderMode_() {
  // 'first'   = very first render (entrance plays later, after the splash)
  // 'refresh' = same date re-rendered with new data
  // 'switch'  = different date than last render
  const key = ovCtxKey_();
  const mode = !_ov.rendered ? 'first' : (_ov.ctx === key ? 'refresh' : 'switch');
  _ov.ctx = key;
  return mode;
}

function ovAfterRender_() {
  _ov.rendered = true;
  ovMaybePlayEntrance_();
}

function ovPrettyDate_(ymd) {
  if (!ymd) return '';
  const d = new Date(`${ymd}T12:00:00Z`);
  return new Intl.DateTimeFormat('en-US', { timeZone: 'UTC', weekday: 'short', month: 'short', day: 'numeric' }).format(d);
}

function ovEsc_(v) {
  return String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function ovTag_(kind, text) {
  return `<span class="ov-tag ov-tag-${kind}">${ovEsc_(text)}</span>`;
}

/* ── Stats ─────────────────────────────────────────────────── */
function ovTimingStats_(rows, field) {
  const vals = rows.map(r => parseFloat(r[field])).filter(v => Number.isFinite(v));
  return { n: vals.length, avg: vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null };
}

/* ── Numbers ───────────────────────────────────────────────── */
function ovCount_(el, to, from, ms = 700) {
  const gen = (el._ovGen = (el._ovGen || 0) + 1);   // a newer update cancels an older count
  el.dataset.v = to;
  if (ovReduced_() || from === to || !Number.isFinite(from)) { el.textContent = to; return; }
  const t0 = performance.now();
  const step = now => {
    if (el._ovGen !== gen) return;
    const p = Math.min(1, (now - t0) / ms), e = 1 - Math.pow(1 - p, 3);
    el.textContent = Math.round(from + (to - from) * e);
    if (p < 1) requestAnimationFrame(step);
  };
  requestAnimationFrame(step);
}

function ovFlash_(el) {
  if (!el) return;
  el.classList.remove('ov-changed', 'ov-changed-static');
  void el.offsetWidth;
  el.classList.add(ovReduced_() ? 'ov-changed-static' : 'ov-changed');
  clearTimeout(el._ovFlashT);
  el._ovFlashT = setTimeout(() => el.classList.remove('ov-changed', 'ov-changed-static'), 2500);
}

function ovSetCount_(id, value, mode) {
  const el = document.getElementById(id);
  if (!el) return;
  const prev = Number(el.dataset.v);
  if (mode === 'refresh') {
    ovCount_(el, value, Number.isFinite(prev) ? prev : value);
    if (Number.isFinite(prev) && prev !== value) ovFlash_(el);
  } else {
    ovCount_(el, value, value);   // instant; entrance count-up happens separately
  }
}

function ovJobsDelta_(jobs, mode) {
  const el = document.getElementById('ovJobsDelta');
  if (!el) return;
  if (mode === 'refresh' && _ov.lastJobs !== null && _ov.lastJobs !== jobs) {
    const d = jobs - _ov.lastJobs;
    el.textContent = (d > 0 ? '+' : '') + d;
    el.classList.add('show');
    clearTimeout(el._ovT);
    el._ovT = setTimeout(() => el.classList.remove('show'), 4000);
  } else if (mode !== 'refresh') {
    el.classList.remove('show');
  }
  _ov.lastJobs = jobs;
}

/* ── Widths (bars, split, coverage) ────────────────────────── */
function ovSetWidth_(el, pct, mode) {
  if (!el) return;
  const w = `${Math.max(0, Math.min(100, pct))}%`;
  if (mode === 'refresh' && !ovReduced_()) { void el.offsetWidth; el.style.width = w; return; }   // CSS transition
  el.style.transition = 'none';
  el.style.width = w;
  void el.offsetWidth;
  el.style.transition = '';
  el.dataset.w = w;
}

/* ── Context title ─────────────────────────────────────────── */
function ovRenderContext_() {
  const dateVal = document.getElementById('dateSingle')?.value || '';
  const t = document.getElementById('opsTitle');
  const s = document.getElementById('opsSubtitle');
  if (dateVal) {
    const long = new Intl.DateTimeFormat('en-US', { timeZone: 'UTC', weekday: 'long', month: 'short', day: 'numeric' })
      .format(new Date(`${dateVal}T12:00:00Z`));
    if (t) t.textContent = `Power breakage · ${long}`;
    if (s) s.textContent = 'Breakage-history record for one day. Prescription (Rx) detail is only available in the live view.';
  } else {
    if (t) t.textContent = 'Power breakage · Today (live)';
    if (s) s.textContent = 'Live report. Includes carryover jobs from earlier ORB dates.';
  }
}

/* ── Data quality ──────────────────────────────────────────── */
const OV_SHEET_ERR = /^#(N\/A|REF!|VALUE!|ERROR!|DIV\/0!|NAME\?|NUM!|NULL!)/i;
function ovRenderDataQuality_(rows) {
  const el = document.getElementById('ovDataQuality');
  if (!el) return;
  const bad = rows.filter(r => [r.BrkSourceMachine, r.BrkReason, r.Material, r.RxNumber]
    .some(v => OV_SHEET_ERR.test(String(v || '').trim())));
  if (!bad.length) { el.hidden = true; el.innerHTML = ''; return; }
  const sample = String([bad[0].BrkSourceMachine, bad[0].BrkReason, bad[0].Material, bad[0].RxNumber]
    .find(v => OV_SHEET_ERR.test(String(v || '').trim())) || '').trim();
  el.hidden = false;
  el.innerHTML = `<b>${bad.length} of ${rows.length} ${rows.length === 1 ? 'row has' : 'rows have'} a sheet error</b> (${ovEsc_(sample)}) ` +
    `in machine, reason, material, or RX. ${bad.length === 1 ? 'It is' : 'They are'} still counted in the totals below — fix at the source sheet.`;
}

/* ── Reason split ──────────────────────────────────────────── */
function ovRenderReasonSplit_(rsnE, totalJobs, mode) {
  const bar = document.getElementById('ovReasonSplit');
  const rowsEl = document.getElementById('ovReasonRows');
  if (!bar || !rowsEl) return;
  if (!totalJobs) {
    bar.innerHTML = '';
    rowsEl.innerHTML = '<div class="ov-empty">No breakage recorded</div>';
    return;
  }
  // keyed segments so widths can move on refresh
  const segs = new Map([...bar.children].map(c => [c.dataset.key, c]));
  const seen = new Set();
  rsnE.forEach(([reason, v]) => {
    let seg = segs.get(reason);
    if (!seg) {
      seg = document.createElement('div');
      seg.className = 'ov-seg';
      seg.dataset.key = reason;
      seg.style.width = '0%';
      seg.style.background = reasonColor(reason);
    }
    bar.appendChild(seg);
    seen.add(reason);
    ovSetWidth_(seg, v.jobs / totalJobs * 100, mode);
  });
  segs.forEach((seg, k) => { if (!seen.has(k)) seg.remove(); });

  const prevSig = {};
  rowsEl.querySelectorAll('.ov-rrow').forEach(r => { prevSig[r.dataset.key] = r.dataset.sig; });
  rowsEl.innerHTML = rsnE.map(([reason, v]) => `
    <div class="ov-rrow" data-key="${ovEsc_(reason)}" data-sig="${v.jobs}/${v.lens}">
      <span><span class="ov-sw" style="background:${reasonColor(reason)}"></span>${ovEsc_(reason)}</span>
      <span class="ov-mono"><b>${v.jobs}</b> ${v.jobs === 1 ? 'job' : 'jobs'} · ${v.lens} ${v.lens === 1 ? 'lens' : 'lenses'} · ${Math.round(v.jobs / totalJobs * 100)}%</span>
    </div>`).join('');
  if (mode === 'refresh') {
    rowsEl.querySelectorAll('.ov-rrow').forEach(r => {
      const was = prevSig[r.dataset.key];
      if (was !== undefined && was !== r.dataset.sig) ovFlash_(r.querySelector('.ov-mono'));
    });
  }
}

/* ── Timing tiles ──────────────────────────────────────────── */
function ovRenderTiming_(key, numId, stats, totalJobs, rsnE, field, mode) {
  const numEl  = document.getElementById(numId);
  const tagsEl = document.getElementById(key === 'ar' ? 'ovArTags' : 'ovBpTags');
  const covEl  = document.getElementById(key === 'ar' ? 'ovArCov' : 'ovBpCov');
  const barEl  = document.getElementById(key === 'ar' ? 'ovArCovBar' : 'ovBpCovBar');
  const byEl   = document.getElementById(key === 'ar' ? 'ovArByReason' : 'ovBpByReason');

  const text = stats.n ? fmtMin(stats.avg) : '--';
  if (numEl) {
    const changed = mode === 'refresh' && numEl.textContent !== text && numEl.textContent !== '--';
    numEl.textContent = text;
    if (changed) ovFlash_(numEl);
  }
  if (tagsEl) tagsEl.innerHTML = '';   // badge removed by request; small sample is stated in the coverage line
  if (covEl) covEl.innerHTML = totalJobs
    ? (!stats.n ? `No job has this timing yet (0 of ${totalJobs})`
       : stats.n < OV_SMALL_N ? `Based on only <b>${stats.n} of ${totalJobs}</b> jobs — too few to trust as a trend`
       : `Average of <b>${stats.n} of ${totalJobs}</b> jobs that have this timing`)
    : 'No jobs';
  ovSetWidth_(barEl, totalJobs ? stats.n / totalJobs * 100 : 0, mode);

  if (byEl) {
    byEl.innerHTML = rsnE.length > 1 ? rsnE.map(([reason]) => {
      const st = ovTimingStats_(_brk.filter(r => (r.BrkReason || 'Unknown') === reason), field);
      return `<span class="ov-by"><span class="ov-sw" style="background:${reasonColor(reason)}"></span>${ovEsc_(reason)} ` +
             `<b class="ov-mono">${st.n ? fmtMin(st.avg) : '--'}</b> <span class="ov-n">(${st.n} timed)</span></span>`;
    }).join('') : '';
  }
}

/* ── Keyed bar list (reuses rows so widths can move) ───────── */
function ovRenderBars_(elId, entries, unit, mode, tipFn) {
  const box = document.getElementById(elId);
  if (!box) return;
  if (!entries.length) { box.innerHTML = `<div class="ov-empty">No ${unit} recorded</div>`; return; }
  box.querySelector('.ov-empty')?.remove();

  const max = Math.max(1, ...entries.map(e => e[1]));
  const existing = new Map([...box.querySelectorAll('.ov-bar')].map(r => [r.dataset.key, r]));
  const seen = new Set();

  entries.forEach(([label, val]) => {
    let row = existing.get(label);
    const isNew = !row;
    if (isNew) {
      row = document.createElement('div');
      row.className = 'ov-bar';
      row.dataset.key = label;
      row.innerHTML = '<span class="ov-bar-lbl"></span><div class="ov-bar-track"><div class="ov-bar-fill"></div></div><span class="ov-bar-val" data-count></span>';
      row.querySelector('.ov-bar-lbl').textContent = label;
      row.querySelector('.ov-bar-fill').style.width = '0%';
    }
    box.appendChild(row);   // keeps sort order
    seen.add(label);
    if (tipFn) row.title = tipFn(label);

    const fill = row.querySelector('.ov-bar-fill');
    const valEl = row.querySelector('.ov-bar-val');
    const prev = Number(valEl.dataset.v);
    if (mode === 'refresh' && isNew && !ovReduced_()) void fill.offsetWidth;   // let a new row grow from 0
    ovSetWidth_(fill, val / max * 100, mode);
    if (mode === 'refresh') {
      ovCount_(valEl, val, Number.isFinite(prev) ? prev : 0);
      if (!isNew && prev !== val) ovFlash_(valEl);
    } else {
      ovCount_(valEl, val, val);
    }
  });
  existing.forEach((row, k) => { if (!seen.has(k)) row.remove(); });
}

/* ── Date switch fade ──────────────────────────────────────── */
function ovFadeIn_() {
  const sec = document.getElementById('t-overview');
  if (!sec || ovReduced_() || !sec.animate) return;
  sec.animate([{ opacity: 0.15 }, { opacity: 1 }], { duration: 260, easing: 'ease-out' });
}

/* ── Entrance (once, after the splash is gone) ─────────────── */
function ovSplashVisible_() {
  const o = document.getElementById('loadingOverlay');
  return !!o && !o.classList.contains('hidden');
}

function ovMaybePlayEntrance_() {
  if (_ov.entrancePlayed || !_ov.rendered) return;
  if (ovSplashVisible_()) return;   // the observer below calls again when it hides
  _ov.entrancePlayed = true;
  if (ovReduced_()) return;

  const sec = document.getElementById('t-overview');
  if (!sec) return;
  const cards = [...sec.querySelectorAll('.ov-enter')];
  if (sec.animate) {
    cards.forEach((c, i) => c.animate(
      [{ opacity: 0, transform: 'translateY(10px)' }, { opacity: 1, transform: 'none' }],
      { duration: 450, delay: i * 70, easing: 'ease-out', fill: 'backwards' }
    ));
  }
  sec.querySelectorAll('.ov-bar-fill, .ov-seg, .ov-covbar > div').forEach((f, i) => {
    const w = f.style.width;
    f.style.transition = 'none';
    f.style.width = '0%';
    void f.offsetWidth;
    f.style.transition = '';
    setTimeout(() => { f.style.width = w; }, 180 + (i % 8) * 60);
  });
  sec.querySelectorAll('[data-count]').forEach(el => {
    const to = Number(el.dataset.v);
    if (Number.isFinite(to)) { el.textContent = '0'; ovCount_(el, to, 0, 800); }
  });
}

(function ovWatchSplash_() {
  const o = document.getElementById('loadingOverlay');
  if (!o) return;
  new MutationObserver(() => { if (!ovSplashVisible_()) ovMaybePlayEntrance_(); })
    .observe(o, { attributes: true, attributeFilter: ['class'] });
})();

/* ── Non-blocking update banner ────────────────────────────── */
function ovBanner_(kind, msg) {
  const el = document.getElementById('updateBanner');
  const m  = document.getElementById('updateBannerMsg');
  if (!el || !m) return;
  m.textContent = msg;
  el.className = `update-banner show is-${kind}`;
  clearTimeout(el._ovT);
  if (kind === 'ok')   el._ovT = setTimeout(() => el.classList.remove('show'), 3000);
  if (kind === 'fail') el._ovT = setTimeout(() => el.classList.remove('show'), 8000);
}

/* ============================================================
   RX FLOW HELPERS (v6, Option A) — pipeline, shared axis, animation
   UI only. Stage times come from renderFlow(); missing = null.
============================================================ */
const flowState_ = { anchor: null, rendered: false, entrancePlayed: false };

function flowMedian_(vals) {
  if (!vals.length) return null;
  const s = [...vals].sort((a, b) => a - b), m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

function flowNiceMax_(m) {
  const steps = [60, 120, 180, 360, 720, 1440, 2160, 2880, 4320, 5760, 7200, 10080];
  return steps.find(s => s >= m) || Math.ceil(m / 1440) * 1440;
}

function flowTickStep_(max) {
  return [15, 30, 60, 120, 180, 360, 720, 1440, 2880].find(s => max / s <= 8) || 1440 * Math.ceil(max / 1440 / 8);
}

function flowAxisHtml_(max) {
  const step = flowTickStep_(max);
  const list = document.getElementById('flowList');
  if (list) list.style.setProperty('--fj-step', `${step / max * 100}%`);
  let ticks = '';
  for (let m = 0; m <= max + 0.5; m += step) {
    ticks += `<div class="fj-tick" style="left:${m / max * 100}%"><span>${m === 0 ? '0' : fmtMin(m)}</span></div>`;
  }
  return `
    <div class="fj-legend">
      <span><i class="fj-sw seg-1"></i>1 In transit</span>
      <span><i class="fj-sw seg-2"></i>2 Identified</span>
      <span><i class="fj-sw seg-3"></i>3 Processing</span>
      <span>Bar length = real time on one shared clock · dots: ● scanned ○ not scanned</span>
    </div>
    <div class="fj-axis"><div></div><div></div><div class="fj-axis-track">${ticks}</div><div></div></div>`;
}

function flowRenderPipeline_(jobs) {
  const box = document.getElementById('flowPipeline');
  const sub = document.getElementById('fpSub');
  const gap = document.getElementById('flowPipelineGap');
  if (!box) return;
  const N = jobs.length;
  const histDate = document.getElementById('dateSingle')?.value || '';
  if (sub) sub.textContent = `${N} ${N === 1 ? 'job' : 'jobs'} · ${histDate ? ovPrettyDate_(histDate) : 'live report'}`;

  const nodes = [['ORB / OTB', 'source scan'], ['AR41', 'inspection scan'], ['Breakage table', 'table scan'], ['Processed', 'reorder']];
  const stages = [
    { cls: 'seg-1', name: '1 · In transit',  desc: 'Source-machine scan → AR41 scan' },
    { cls: 'seg-2', name: '2 · Identified',  desc: 'AR41 scan → breakage table scan' },
    { cls: 'seg-3', name: '3 · Processing',  desc: 'Breakage table → processed / reorder' },
  ];

  let html = '';
  stages.forEach((sg, i) => {
    const vals = jobs.map(j => j.st[i]).filter(v => v !== null);
    const n = vals.length;
    const med = flowMedian_(vals);
    const max = n ? Math.max(...vals) : null;
    const big = !n ? '--' : (n === 1 ? fmtMin(vals[0]) : fmtMin(med));
    const line = !n ? 'No job has this timing'
      : n === 1 ? 'one job only — not a trend'
      : n < OV_SMALL_N ? `median of only ${n} jobs — not a trend`
      : `median · longest ${fmtMin(max)}`;
    const k = !n || !N ? 0 : Math.max(1, Math.round(n / N * 3));   // particles scale with timed jobs
    const parts = Array.from({ length: k }, (_, p) => `<i style="animation-delay:${(p * 3 / k).toFixed(2)}s"></i>`).join('');
    html += `<div class="fp-node" title="${nodes[i][1]}">${nodes[i][0]}<small>${nodes[i][1]}</small></div>
      <div class="fp-link">
        <div class="fp-line ${sg.cls}${n ? '' : ' empty'}">${parts}</div>
        <div class="fp-stage" title="${sg.desc}">
          <div class="fp-name"><i class="fj-sw ${sg.cls}"></i>${sg.name}</div>
          <div class="fp-big">${big}</div>
          <div class="fp-line2">${line}</div>
          <div class="fp-cov"><b>${n} of ${N}</b> jobs timed</div>
          <div class="fp-covbar"><div class="${sg.cls}" style="width:${N ? n / N * 100 : 0}%"></div></div>
        </div>
      </div>`;
  });
  html += `<div class="fp-node" title="${nodes[3][1]}">${nodes[3][0]}<small>${nodes[3][1]}</small></div>`;
  box.innerHTML = html;

  if (gap) {
    const k = jobs.filter(j => j.st[1] === null && j.st[2] === null).length;
    gap.hidden = !k;
    gap.textContent = !k ? '' : histDate
      ? `${k} of ${N} jobs have no breakage-table or processed time recorded for this day. They show as "Identify + Processing skipped".`
      : `${k} of ${N} jobs have no breakage-table or processed scan yet. They show as "not scanned yet".`;
  }
}

function flowTabVisible_() {
  return document.getElementById('t-flow')?.classList.contains('active');
}

function flowPlayEntrance_() {
  flowState_.entrancePlayed = true;
  if (ovReduced_()) return;
  document.querySelectorAll('#flowList .fj-bar').forEach((bar, i) => {
    if (!bar.animate) return;
    bar.animate([{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0 0 0 0)' }],
      { duration: 900, delay: Math.min(i, 20) * 30, easing: 'cubic-bezier(.2,.7,.2,1)', fill: 'backwards' });
  });
  document.querySelectorAll('#flowList .fj-after').forEach((el, i) => {
    if (!el.animate) return;
    el.animate([{ opacity: 0 }, { opacity: 1 }], { duration: 300, delay: 600 + Math.min(i, 20) * 30, fill: 'backwards' });
  });
}

function flowAnimate_(prevW) {
  flowState_.rendered = true;
  if (!flowTabVisible_()) return;                 // entrance waits until the tab is opened
  if (!flowState_.entrancePlayed) { flowPlayEntrance_(); return; }
  if (ovReduced_()) return;
  if (!prevW) {                                    // different day: quick fade, no "growth"
    const list = document.getElementById('flowList');
    list?.animate?.([{ opacity: 0.15 }, { opacity: 1 }], { duration: 260, easing: 'ease-out' });
    return;
  }
  document.querySelectorAll('#flowList .fj-row').forEach(row => {
    const was = prevW[row.dataset.rx], now = Number(row.dataset.w);
    const bar = row.querySelector('.fj-bar');
    if (!bar?.animate || !(now > 0)) return;
    if (was === undefined) {
      bar.animate([{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0 0 0 0)' }], { duration: 600, easing: 'ease-out' });
    } else if (Math.abs(was - now) > 0.2) {
      bar.style.transformOrigin = 'left center';
      bar.animate([{ transform: `scaleX(${Math.max(0.02, was / now)})` }, { transform: 'scaleX(1)' }], { duration: 600, easing: 'ease-out' });
    }
  });
}

(function flowWatchTab_() {
  const t = document.getElementById('t-flow');
  if (!t) return;
  new MutationObserver(() => {
    if (flowTabVisible_() && flowState_.rendered && !flowState_.entrancePlayed) flowPlayEntrance_();
  }).observe(t, { attributes: true, attributeFilter: ['class'] });
})();

/* ============================================================
   ANALYSIS HELPERS (v6) — plain-language Rx cards
   Cutoffs are PROPOSED defaults. Change them here; every card,
   label, and table word follows.
============================================================ */
const RX_CUTOFFS = {
  strongSph: 4.00,   // |sphere| ≥ this = strong
  cylMedium: 1.00,   // |cylinder| above this = medium
  cylHigh:   2.00,   // |cylinder| above this = high
};
const _an = { prevCounts: {}, prevChips: new Set(), entrancePlayed: false, rendered: false };

function rxNum_(v) {
  const n = parseFloat(v);
  return (v === '' || v === null || v === undefined || isNaN(n)) ? null : n;
}

// Stronger eye = larger absolute value; null when neither eye has a value
function rxStrongest_(a, b) {
  const x = rxNum_(a), y = rxNum_(b);
  if (x === null) return y;
  if (y === null) return x;
  return Math.abs(y) > Math.abs(x) ? y : x;
}

function rxD_(v) { return (v > 0 ? '+' : v < 0 ? '−' : '') + Math.abs(v).toFixed(2); }

function rxSphCats_() {
  const S = RX_CUTOFFS.strongSph;
  return [
    { key: 'sm', name: 'Strong minus', range: `${rxD_(-S)} or stronger`,       test: v => v <= -S },
    { key: 'mm', name: 'Mild minus',   range: `−0.25 to ${rxD_(-(S - 0.25))}`, test: v => v < 0 && v > -S },
    { key: 'np', name: 'No power',     range: '0.00',                          test: v => v === 0 },
    { key: 'mp', name: 'Mild plus',    range: `+0.25 to ${rxD_(S - 0.25)}`,    test: v => v > 0 && v < S },
    { key: 'sp', name: 'Strong plus',  range: `${rxD_(S)} or stronger`,        test: v => v >= S },
  ];
}

function rxCylCats_() {
  const M = RX_CUTOFFS.cylMedium, H = RX_CUTOFFS.cylHigh;
  return [
    { key: 'lo', name: 'Low',    range: `0 to ${rxD_(-M)}`,                         test: v => Math.abs(v) <= M },
    { key: 'md', name: 'Medium', range: `${rxD_(-(M + 0.25))} to ${rxD_(-H)}`,      test: v => Math.abs(v) > M && Math.abs(v) <= H },
    { key: 'hi', name: 'High',   range: `${rxD_(-(H + 0.25))} or stronger`,         test: v => Math.abs(v) > H },
  ];
}

function anRenderHeadline_(jobs, totalLens, noRx) {
  const h = document.getElementById('anHeadline');
  const sub = document.getElementById('anHeadSub');
  if (!h) return;
  const N = jobs.length;
  if (!N) { h.textContent = 'No jobs match these filters.'; if (sub) sub.textContent = ''; return; }

  const S = RX_CUTOFFS.strongSph;
  const sm = jobs.filter(j => j.sph !== null && j.sph <= -S).length;
  const sp = jobs.filter(j => j.sph !== null && j.sph >= S).length;
  const strong = sm + sp;
  const hi = jobs.filter(j => /high\s*index/i.test(j.r.Material || '')).length;
  const all = N === 1 ? 'The one broken job' : N === 2 ? 'Both broken jobs' : `All ${N} broken jobs`;
  const small = N < OV_SMALL_N ? ` · only ${N} ${N === 1 ? 'job' : 'jobs'}, so this is a hint to watch, not proof` : '';

  if (noRx) {
    h.textContent = hi === N ? `${all} ${N === 1 ? 'was a' : 'were'} high-index ${N === 1 ? 'lens' : 'lenses'}.` : `${hi} of ${N} broken jobs were high-index lenses.`;
    if (sub) sub.textContent = ` · prescription detail is live-only${small}`;
    return;
  }
  if (strong === N && hi === N) h.textContent = `${all} had a strong prescription and a high-index lens.`;
  else if (strong === N)        h.textContent = `${all} had a strong prescription.`;
  else h.textContent = `${strong} of ${N} broken jobs had a strong prescription · ${hi} of ${N} were high-index lenses.`;
  if (sub) sub.textContent = ` · ${sm} strong minus, ${sp} strong plus${small}`;
}

function anRenderScale_(elId, missId, jobs, field, cats) {
  const box = document.getElementById(elId);
  if (!box) return;
  const withVal = jobs.filter(j => j[field] !== null);
  const N = withVal.length || 1;
  box.innerHTML = cats.map(c => {
    const js = withVal.filter(j => c.test(j[field]));
    const chips = js.map(j => {
      const rx = String(j.r.RxNumber || '');
      const v = j[field];
      return `<span class="an-chip" data-rx="${ovEsc_(rx)}" title="${ovEsc_(rx)} · ${ovEsc_(j.r.BrkReason || '')}"><i style="background:${reasonColor(j.r.BrkReason)}"></i>${ovEsc_(rx.slice(-4))}<small>${v === 0 ? '0' : (v > 0 ? '+' + v : v)}</small></span>`;
    }).join('');
    return `<div class="an-cat${js.length ? '' : ' zero'}" data-key="${field}-${c.key}">
        <div class="an-cn" data-count>${js.length}</div>
        <div class="an-cl">${c.name}</div>
        <div class="an-cr">${c.range}</div>
        <div class="an-gauge"><div data-h="${js.length / N * 100}" style="height:${js.length / N * 100}%"></div></div>
        <div class="an-chips">${chips}</div>
      </div>`;
  }).join('');

  const miss = document.getElementById(missId);
  const m = jobs.length - withVal.length;
  if (miss) {
    miss.hidden = !m;
    miss.textContent = m ? `${m} of ${jobs.length} ${jobs.length === 1 ? 'job has' : 'jobs have'} no ${field === 'sph' ? 'sphere' : 'cylinder'} value and ${m === 1 ? "isn't" : "aren't"} on a card.` : '';
  }
}

function anRenderBars_(elId, entries, totalJobs, colorFn, dot) {
  const el = document.getElementById(elId);
  if (!el) return;
  if (!entries.length) { el.innerHTML = '<div class="empty">No jobs</div>'; return; }
  const max = Math.max(1, ...entries.map(e => e[1].jobs));
  el.innerHTML = entries.map(([k, v]) => `
    <div class="an-bar" data-key="${elId}-${ovEsc_(k)}">
      <span class="an-bar-lbl">${dot ? `<i style="background:${colorFn(k)}"></i>` : ''}${ovEsc_(k)}</span>
      <div class="an-bar-track"><div data-w="${v.jobs / max * 100}" style="width:${v.jobs / max * 100}%;background:${colorFn(k)}"></div></div>
      <b>${v.jobs} ${v.jobs === 1 ? 'job' : 'jobs'}</b>
      <small>${v.lens} ${v.lens === 1 ? 'lens' : 'lenses'}</small>
    </div>`).join('');
}

/* ── Animation ─────────────────────────────────────────────── */
function anTabVisible_() { return document.getElementById('t-research')?.classList.contains('active'); }

function anPlayEntrance_() {
  _an.entrancePlayed = true;
  if (ovReduced_()) return;
  const sec = document.getElementById('t-research');
  if (!sec?.animate) return;
  sec.querySelectorAll('.an-gauge > div, .an-bar-track > div').forEach((el, i) => {
    const prop = el.parentElement.classList.contains('an-gauge') ? 'height' : 'width';
    const to = el.style[prop];
    el.animate([{ [prop]: '0%' }, { [prop]: to }], { duration: 700, delay: 80 + (i % 8) * 50, easing: 'cubic-bezier(.2,.7,.2,1)', fill: 'backwards' });
  });
  sec.querySelectorAll('.an-chip').forEach((c, i) => {
    c.animate([{ opacity: 0, transform: 'scale(.6)' }, { opacity: 1, transform: 'none' }],
      { duration: 300, delay: 500 + Math.min(i, 20) * 45, easing: 'ease-out', fill: 'backwards' });
  });
  sec.querySelectorAll('.an-cn').forEach(el => {
    const to = Number(el.textContent);
    if (to > 0) { el.textContent = '0'; ovCount_(el, to, 0, 700); }
  });
}

function anAfterRender_() {
  const sec = document.getElementById('t-research');
  const counts = {}, chips = new Set();
  sec?.querySelectorAll('.an-cat').forEach(c => { counts[c.dataset.key] = Number(c.querySelector('.an-cn').textContent); });
  sec?.querySelectorAll('.an-chip').forEach(c => chips.add(c.dataset.rx));

  const first = !_an.rendered;
  _an.rendered = true;

  if (!anTabVisible_()) { _an.prevCounts = counts; _an.prevChips = chips; return; }
  if (!_an.entrancePlayed) { anPlayEntrance_(); }
  else if (!first && !ovReduced_() && sec?.animate) {
    // re-render: gauges move from old height, new chips pop in, changed counts flash
    sec.querySelectorAll('.an-cat').forEach(c => {
      const was = _an.prevCounts[c.dataset.key];
      const now = Number(c.querySelector('.an-cn').textContent);
      if (was !== undefined && was !== now) ovFlash_(c.querySelector('.an-cn'));
    });
    sec.querySelectorAll('.an-chip').forEach(c => {
      if (!_an.prevChips.has(c.dataset.rx)) c.animate([{ opacity: 0, transform: 'scale(.6)' }, { opacity: 1, transform: 'none' }], { duration: 300, easing: 'ease-out' });
    });
  }
  _an.prevCounts = counts; _an.prevChips = chips;
}

(function anWatchTab_() {
  const t = document.getElementById('t-research');
  if (!t) return;
  new MutationObserver(() => {
    if (anTabVisible_() && _an.rendered && !_an.entrancePlayed) anPlayEntrance_();
  }).observe(t, { attributes: true, attributeFilter: ['class'] });

  // Hover a chip or a table row: light up the same job everywhere on this tab
  const mark = (rx, on) => {
    if (!rx) return;
    t.querySelectorAll(`[data-rx="${CSS.escape(rx)}"]`).forEach(el => el.classList.toggle('an-hl', on));
  };
  t.addEventListener('mouseover', e => { const el = e.target.closest('[data-rx]'); if (el) mark(el.dataset.rx, true); });
  t.addEventListener('mouseout',  e => { const el = e.target.closest('[data-rx]'); if (el) mark(el.dataset.rx, false); });
})();