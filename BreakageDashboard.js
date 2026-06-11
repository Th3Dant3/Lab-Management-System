/************************************************************
 * QUALITY / BREAKAGE HUB
 * Productivity Dashboard style redesign.
 * Uses existing Apps Script actions:
 *   ?action=breakageBoot
 *   ?action=breakageWeek
 *   ?action=breakageAssociate
 ************************************************************/

const API_URL = 'https://script.google.com/macros/s/AKfycbySXKQSHTXQVlZNq2vubqp2D3W-_IgtmaiFr_GDxz5X4FO2cqcYeUkAo_A9LajOfj9f/exec';
const BREAKAGE_CACHE_PREFIX = 'QUALITY_BREAKAGE_WEEK_CACHE_V6_REASON_OPERATOR_';
const BREAKAGE_CACHE_TTL_MS = 60 * 60 * 1000;

const STATE = {
  boot: null,
  week: null,
  allWeeks: {},
  selectedWeek: null,
  activePage: 'dashboard',
  activeDept: 'ALL',
  activeShift: 'WEEKDAY',
  activeStationGroup: 'ALL',
  activeStatus: 'ALL',
  activeReason: 'ALL',
  selectedAssociate: '',
  search: '',
  isLoading: false,
  loadedWeekOrder: []
};

const $ = id => document.getElementById(id);

document.addEventListener('DOMContentLoaded', () => {
  startClock();
  bindEvents();
  startLoadingScreen();
  loadBoot();
});

/* CLOCK */
function startClock() {
  function tick() {
    const now = new Date();
    setText('topTime', now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
    setText('topDate', now.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' }));
  }
  tick();
  setInterval(tick, 1000);
}

/* EVENTS */
function bindEvents() {
  $('backToDashboardBtn')?.addEventListener('click', () => {
    window.location.href = 'index.html';
  });

  $('logoutBtn')?.addEventListener('click', () => {
    localStorage.removeItem('LMS_AUTH_USER');
    showRefreshStatus('Logged out locally');
  });

  document.querySelectorAll('.side-link').forEach(btn => {
    btn.addEventListener('click', () => setActivePage(btn.dataset.page || 'dashboard'));
  });

  document.querySelectorAll('.hub-shift-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      STATE.activeShift = btn.dataset.shift || 'WEEKDAY';
      document.querySelectorAll('.hub-shift-btn').forEach(b => b.classList.remove('active', 'wd', 'we', 'all'));
      btn.classList.add('active');
      if (STATE.activeShift === 'WEEKDAY') btn.classList.add('wd');
      else if (STATE.activeShift === 'WEEKEND') btn.classList.add('we');
      else btn.classList.add('all');

      const labels = { WEEKDAY: 'Weekday', WEEKEND: 'Weekend', ALL: 'All Shifts' };
      const label = $('shiftMetaLabel');
      if (label) label.innerHTML = `Shift: <strong>${labels[STATE.activeShift] || STATE.activeShift}</strong>`;
      renderEverything();
    });
  });

  document.querySelectorAll('.hub-dept-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      const jump = tab.dataset.pageJump;
      if (jump) {
        setActivePage(jump);
        return;
      }
      document.querySelectorAll('.hub-dept-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      STATE.activeDept = tab.dataset.dept || 'ALL';
      STATE.activeStationGroup = 'ALL';
      STATE.activeReason = 'ALL';
      STATE.selectedAssociate = '';
      const reasonSelect = $('reasonSelect');
      if (reasonSelect) reasonSelect.value = '';
      renderNavBar();
      renderEverything();
    });
  });

  [$('fiscalWeekSelect'), $('sideWeekSelect')].forEach(sel => {
    sel?.addEventListener('change', async e => {
      const enriched = getBreakageFiscalWeeksEnriched();
      const week = enriched.find(w => w.key === e.target.value)
        || (STATE.boot?.fiscalWeeks || []).find(w => w.key === e.target.value);
      if (!week) return;
      STATE.selectedWeek = week;
      STATE._userChangedFiscalWeek = true;  // user explicitly picked — don't override with current
      STATE.selectedAssociate = '';
      syncWeekSelects();
      await loadWeek(false);
    });
  });

  $('associateSelect')?.addEventListener('change', e => {
    STATE.selectedAssociate = e.target.value || '';
    if (STATE.selectedAssociate) setActivePage('associates');
    renderEverything();
  });

  $('reasonSelect')?.addEventListener('change', e => {
    STATE.activeReason = e.target.value || 'ALL';
    STATE.selectedAssociate = '';
    if (STATE.activeReason !== 'ALL') setActivePage('associates');
    renderEverything();
  });

  $('searchInput')?.addEventListener('input', e => {
    STATE.search = String(e.target.value || '').trim().toLowerCase();
    renderEverything();
  });

  $('refreshBtn')?.addEventListener('click', () => loadBoot(true));

  $('closeModalBtn')?.addEventListener('click', closeModal);
  $('associateModal')?.addEventListener('click', e => {
    if (e.target && e.target.id === 'associateModal') closeModal();
  });

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeModal();
  });
}

/* LOADER */
function startLoadingScreen() {
  updateLoader(12, 'Loading breakage command data...', 'boot');
  const loader = $('appLoader');
  if (loader) loader.classList.remove('hide');
}

function finishLoadingScreen() {
  updateLoader(100, 'Quality Hub ready.', 'render');
  setTimeout(() => $('appLoader')?.classList.add('hide'), 350);
}

function updateLoader(percent, message, step) {
  const pct = Math.max(0, Math.min(100, Number(percent || 0)));
  const bar = $('loaderBar');
  const pctEl = $('loaderPct');
  const msg = $('loaderMessage');
  if (bar) bar.style.width = `${pct}%`;
  if (pctEl) pctEl.textContent = `${Math.round(pct)}%`;
  if (msg && message) msg.textContent = message;

  const steps = ['Boot', 'Week', 'Render'];
  steps.forEach(name => {
    const el = $(`loadStep${name}`);
    if (!el) return;
    const dot = el.querySelector('.ldr-status-dot');
    el.classList.remove('active', 'done');
    if (dot) dot.className = 'ldr-status-dot pending';
  });

  const order = { boot: 0, week: 1, render: 2 };
  const idx = order[step || 'boot'] ?? 0;
  steps.forEach((name, i) => {
    const el = $(`loadStep${name}`);
    const dot = el?.querySelector('.ldr-status-dot');
    if (!el || !dot) return;
    if (i < idx) { el.classList.add('done'); dot.className = 'ldr-status-dot done'; }
    if (i === idx) { el.classList.add('active'); dot.className = 'ldr-status-dot running'; }
  });
}

/* LOAD */
async function loadBoot(force = false) {
  setLoading(true);
  updateLoader(22, force ? 'Refreshing boot data...' : 'Booting Quality Hub...', 'boot');

  try {
    const action = force ? 'breakageBoot&force=1' : 'breakageBoot';
    const data = await fetchApi(action);
    STATE.boot = data || {};

    // Always default to today's actual current FW, not API's defaultWeek.
    // The API may have an old WorkDate in SETTINGS (e.g. pointing to FW 23
    // when today is already FW 24). The client knows what day it is.
    if (!STATE._userChangedFiscalWeek) {
      STATE.selectedWeek = getCurrentBreakageFiscalWeek();
    }

    // If still null (no boot data at all), fall back to API default
    if (!STATE.selectedWeek) {
      STATE.selectedWeek = data.defaultWeek || (data.fiscalWeeks || [])[0] || null;
    }

    renderWeekDropdowns();

    if (STATE.selectedWeek) {
      await loadWeek(force);
    } else {
      STATE.week = null;
      renderEverything();
      finishLoadingScreen();
    }
  } catch (err) {
    console.error(err);
    showError(`Unable to load Breakage Hub. ${err.message || err}`);
  } finally {
    setLoading(false);
  }
}

async function loadWeek(force = false) {
  if (!STATE.selectedWeek) return;

  updateLoader(48, `Loading ${STATE.selectedWeek.fiscalWeek || 'selected week'}...`, 'week');
  showRefreshStatus(`Loading ${STATE.selectedWeek.fiscalWeek || 'selected week'}...`);

  const key = STATE.selectedWeek.key || getWeekKey(STATE.selectedWeek);

  if (!force && STATE.allWeeks[key]) {
    STATE.week = STATE.allWeeks[key];
    afterWeekLoaded('memory cache');
    return;
  }

  if (!force) {
    const cached = getCachedWeek(key);
    if (cached) {
      STATE.week = cached;
      STATE.allWeeks[key] = cached;
      afterWeekLoaded('browser cache');
      return;
    }
  }

  const params = new URLSearchParams({
    action: 'breakageWeek',
    fiscalWeek: STATE.selectedWeek.fiscalWeek || '',
    weekStartDate: STATE.selectedWeek.weekStartDate || '',
    weekEndDate: STATE.selectedWeek.weekEndDate || '',
    ts: String(Date.now())
  });

  try {
    const data = await fetchApiRaw(params.toString());
    STATE.week = normalizeWeekPayload(data || {});
    STATE.allWeeks[key] = STATE.week;
    saveCachedWeek(key, STATE.week);
    afterWeekLoaded('API');
  } catch (err) {
    console.error(err);
    showError(`Failed to load selected week. ${err.message || err}`);
  }
}

function afterWeekLoaded(source) {
  const key = STATE.selectedWeek?.key || getWeekKey(STATE.selectedWeek);
  if (key && !STATE.loadedWeekOrder.includes(key)) STATE.loadedWeekOrder.push(key);

  updateLoader(78, `Rendering ${STATE.selectedWeek?.fiscalWeek || 'week'} from ${source}...`, 'render');
  updateMeta();
  renderNavBar();
  renderReasonDropdown();
  renderEverything();
  showRefreshStatus(`Loaded ${STATE.selectedWeek?.fiscalWeek || 'week'} from ${source}`);
  finishLoadingScreen();
}

function normalizeWeekPayload(data) {
  const records = normalizeBreakageRecords([
    ...(Array.isArray(data.records) ? data.records : []),
    ...(Array.isArray(data.breakageDailySnapshot) ? data.breakageDailySnapshot : []),
    ...(Array.isArray(data.breakageMaster) ? data.breakageMaster : [])
  ]);

  let associates = normalizeAssociateRows(
    Array.isArray(data.associates) ? data.associates :
    Array.isArray(data.operators) ? data.operators :
    Array.isArray(data.breakageOperatorWeeklySummary) ? data.breakageOperatorWeeklySummary : []
  );

  // Some API versions return the weekly totals but not the associate list.
  // When raw records are included, rebuild the operator cards from those records.
  if (!associates.length && records.length) {
    associates = buildAssociatesFromRecords(records);
  }

  const reasons = normalizeReasonRows(
    Array.isArray(data.reasons) ? data.reasons : buildReasonsFromRecords(records)
  );

  const accessPoints = normalizeAccessPointRows(
    Array.isArray(data.accessPoints) ? data.accessPoints : buildAccessPointsFromRecords(records)
  );

  // Enrich associate cards with reasons/access points when records are available.
  if (records.length) {
    const byOperator = new Map();
    records.forEach(r => {
      const op = normalizeName(r.operator || r.OperatorName || 'No Operator');
      if (!byOperator.has(op)) byOperator.set(op, { reasons: new Set(), accessPoints: new Set() });
      const item = byOperator.get(op);
      if (r.reason) item.reasons.add(r.reason);
      if (r.accessPoint) item.accessPoints.add(r.accessPoint);
    });
    associates.forEach(a => {
      const op = normalizeName(a.operator || a.OperatorName || 'No Operator');
      const item = byOperator.get(op);
      if (item) {
        a.reasons = Array.from(item.reasons).sort();
        a.accessPoints = Array.from(new Set([...(a.accessPoints || []), ...Array.from(item.accessPoints)])).sort();
      }
    });
  }

  return {
    ...data,
    summary: data.summary || {},
    records,
    reasons,
    accessPoints,
    associates,
    departments: Array.isArray(data.departments) ? data.departments : []
  };
}

function normalizeAssociateRows(rows) {
  return (rows || []).map(row => {
    const lens = Number(row.lensBroken ?? row.LensBreakageCount ?? row['Lenses Broken'] ?? 0);
    const frame = Number(row.frameBroken ?? row.FrameBreakageCount ?? row['Frames Broken'] ?? 0);
    const total = Number(row.totalBreakage ?? row.TotalBreakageCount ?? row.Total ?? (lens + frame) ?? 0);
    const scans = Number(row.totalDailyScan ?? row.totalScans ?? row.OperatorWeeklyScans ?? row.OperatorDailyScans ?? row.DailyTotalScans ?? row.TotalDailyScan ?? 0);
    const accessPoints = Array.isArray(row.accessPoints) ? row.accessPoints : [row.accessPoint || row.AccessPoint || row.StationGroup || ''].filter(Boolean);
    const reasons = Array.isArray(row.reasons) ? row.reasons : [row.reason || row.BreakageReason || ''].filter(Boolean);
    return {
      ...row,
      operator: normalizeName(row.operator || row.OperatorName || row.operatorName || row.Associate || 'No Operator'),
      department: row.department || row.FinalDepartment || row.Department || 'Quality',
      lensBroken: lens,
      frameBroken: frame,
      totalBreakage: total,
      totalDailyScan: scans,
      lensBreakagePct: Number(row.lensBreakagePct ?? row.OperatorLensBreakagePercent ?? row.DailyLensBreakagePercent ?? (scans ? lens / scans : 0)),
      frameBreakagePct: Number(row.frameBreakagePct ?? row.OperatorFrameBreakagePercent ?? row.DailyFrameBreakagePercent ?? (scans ? frame / scans : 0)),
      records: Number(row.records ?? row.RecordCount ?? 0),
      accessPoints: Array.from(new Set(accessPoints.map(String).filter(Boolean))).sort(),
      reasons: Array.from(new Set(reasons.map(String).filter(Boolean))).sort()
    };
  }).filter(a => a.operator);
}

function normalizeBreakageRecords(rows) {
  return (rows || []).map(row => {
    const lens = Number(row.lensBroken ?? row.LensBreakageCount ?? row['Lenses Broken'] ?? row.LensesBroken ?? 0);
    const frame = Number(row.frameBroken ?? row.FrameBreakageCount ?? row['Frames Broken'] ?? row.FramesBroken ?? 0);
    const total = Number(row.totalBreakage ?? row.TotalBreakageCount ?? row.Total ?? (lens + frame) ?? 0);
    const scans = Number(row.totalDailyScan ?? row.OperatorDailyScans ?? row.DailyTotalScans ?? row.TotalDailyScan ?? row.TotalJobScan ?? 0);
    return {
      ...row,
      workDate: row.workDate || row.WorkDate || row.Date || row.BreakageDate || '',
      operator: normalizeName(row.operator || row.OperatorName || row.Operator || row.Associate || 'No Operator'),
      accessPoint: String(row.accessPoint || row.AccessPoint || row.StationGroup || 'No Access Point').trim() || 'No Access Point',
      reason: String(row.reason || row.BreakageReason || row.Reason || 'No Reason').trim() || 'No Reason',
      department: row.department || row.FinalDepartment || row.Department || 'Quality',
      lensBroken: lens,
      frameBroken: frame,
      totalBreakage: total,
      totalDailyScan: scans,
      lensBreakagePct: scans ? lens / scans : 0,
      frameBreakagePct: scans ? frame / scans : 0
    };
  }).filter(r => r.operator || r.totalBreakage > 0);
}

function buildAssociatesFromRecords(records) {
  const map = new Map();
  records.forEach(r => {
    const op = normalizeName(r.operator || 'No Operator');
    if (!map.has(op)) {
      map.set(op, { operator: op, department: r.department || 'Quality', lensBroken: 0, frameBroken: 0, totalBreakage: 0, totalDailyScan: 0, records: 0, accessPoints: new Set(), reasons: new Set() });
    }
    const a = map.get(op);
    a.lensBroken += Number(r.lensBroken || 0);
    a.frameBroken += Number(r.frameBroken || 0);
    a.totalBreakage += Number(r.totalBreakage || 0);
    a.totalDailyScan += Number(r.totalDailyScan || 0);
    a.records += 1;
    if (r.accessPoint) a.accessPoints.add(r.accessPoint);
    if (r.reason) a.reasons.add(r.reason);
  });
  return Array.from(map.values()).map(a => ({
    ...a,
    accessPoints: Array.from(a.accessPoints).sort(),
    reasons: Array.from(a.reasons).sort(),
    lensBreakagePct: a.totalDailyScan ? a.lensBroken / a.totalDailyScan : 0,
    frameBreakagePct: a.totalDailyScan ? a.frameBroken / a.totalDailyScan : 0
  })).sort((a,b) => b.totalBreakage - a.totalBreakage || a.operator.localeCompare(b.operator));
}

function normalizeReasonRows(rows) {
  return (rows || []).map(r => ({
    ...r,
    name: String(r.name || r.reason || r.BreakageReason || 'No Reason').trim() || 'No Reason',
    lensBroken: Number(r.lensBroken ?? r.LensBreakageCount ?? 0),
    frameBroken: Number(r.frameBroken ?? r.FrameBreakageCount ?? 0),
    totalBreakage: Number(r.totalBreakage ?? r.TotalBreakageCount ?? ((r.lensBroken || 0) + (r.frameBroken || 0)) ?? 0),
    totalDailyScan: Number(r.totalDailyScan ?? r.OperatorDailyScans ?? r.DailyTotalScans ?? 0),
    records: Number(r.records ?? r.RecordCount ?? 0),
    operators: Array.isArray(r.operators) ? r.operators.map(normalizeName) : []
  })).sort((a,b) => b.totalBreakage - a.totalBreakage || a.name.localeCompare(b.name));
}

function normalizeAccessPointRows(rows) {
  return (rows || []).map(ap => ({
    ...ap,
    name: String(ap.name || ap.accessPoint || ap.AccessPoint || 'No Access Point').trim() || 'No Access Point',
    lensBroken: Number(ap.lensBroken ?? ap.LensBreakageCount ?? 0),
    frameBroken: Number(ap.frameBroken ?? ap.FrameBreakageCount ?? 0),
    totalBreakage: Number(ap.totalBreakage ?? ap.TotalBreakageCount ?? ((ap.lensBroken || 0) + (ap.frameBroken || 0)) ?? 0),
    totalDailyScan: Number(ap.totalDailyScan ?? ap.OperatorDailyScans ?? ap.DailyTotalScans ?? 0),
    records: Number(ap.records ?? ap.RecordCount ?? 0)
  })).sort((a,b) => b.totalBreakage - a.totalBreakage || a.name.localeCompare(b.name));
}

function buildReasonsFromRecords(records) {
  const map = new Map();
  records.forEach(r => bumpLocalBreakageMap(map, r.reason, r));
  return Array.from(map.values()).map(x => ({...x, operators: Array.from(x.operators || [])}));
}

function buildAccessPointsFromRecords(records) {
  const map = new Map();
  records.forEach(r => bumpLocalBreakageMap(map, r.accessPoint, r));
  return Array.from(map.values()).map(x => ({...x, operators: Array.from(x.operators || [])}));
}

function bumpLocalBreakageMap(map, name, r) {
  const key = String(name || 'Unknown').trim() || 'Unknown';
  if (!map.has(key)) map.set(key, { name: key, lensBroken: 0, frameBroken: 0, totalBreakage: 0, totalDailyScan: 0, records: 0, operators: new Set() });
  const item = map.get(key);
  item.lensBroken += Number(r.lensBroken || 0);
  item.frameBroken += Number(r.frameBroken || 0);
  item.totalBreakage += Number(r.totalBreakage || 0);
  item.totalDailyScan += Number(r.totalDailyScan || 0);
  item.records += 1;
  if (r.operator) item.operators.add(r.operator);
  item.operators = item.operators instanceof Set ? item.operators : new Set(item.operators || []);
}

function normalizeName(name) {
  return String(name || 'No Operator').trim().replace(/\s+/g, ' ') || 'No Operator';
}

/* CACHE */
function getWeekKey(week) {
  if (!week) return '';
  return String(week.key || `${week.fiscalWeek || ''}|${week.weekStartDate || ''}|${week.weekEndDate || ''}`);
}

function getCacheStorageKey(key) {
  return BREAKAGE_CACHE_PREFIX + String(key || '').replace(/[^a-zA-Z0-9_-]/g, '_');
}

function getCachedWeek(key) {
  try {
    const raw = localStorage.getItem(getCacheStorageKey(key));
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || !parsed.savedAt || !parsed.data) return null;
    if (Date.now() - Number(parsed.savedAt) > BREAKAGE_CACHE_TTL_MS) {
      localStorage.removeItem(getCacheStorageKey(key));
      return null;
    }
    return parsed.data;
  } catch (err) {
    console.warn('Bad breakage week cache:', err);
    return null;
  }
}

function saveCachedWeek(key, data) {
  try {
    localStorage.setItem(getCacheStorageKey(key), JSON.stringify({ savedAt: Date.now(), data }));
  } catch (err) {
    console.warn('Breakage browser cache skipped:', err);
  }
}

/* ─────────────────────────────────────────────────────────────────
   FISCAL WEEK CALCULATION — CLIENT SIDE
   Same Zenni rule as Productivity Hub:
     Week starts Sunday. FW 1 = Sunday of week containing Jan 1.
     May 24 2026 = FW 22, May 31 = FW 23, Jun 7 = FW 24.
   Client calculates this so the dropdown always auto-selects the
   correct current week regardless of what the API returns, and
   isCurrent is always accurate using today's real date.
───────────────────────────────────────────────────────────────── */

function parseLocalDate(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value)) return value;
  const text = String(value).slice(0, 10);
  const parts = text.split('-').map(Number);
  if (parts.length === 3 && parts[0] && parts[1] && parts[2]) {
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }
  const fb = new Date(value);
  return isNaN(fb.getTime()) ? null : fb;
}

function formatDateKey(d) {
  const dt = d instanceof Date ? d : (parseLocalDate(d) || new Date());
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  const day = String(dt.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function getZenniSundayStart(dateValue) {
  const d = parseLocalDate(dateValue) || new Date();
  const local = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  local.setDate(local.getDate() - local.getDay()); // getDay()=0 on Sunday → subtracts 0
  return local;
}

function getZenniFiscalWeekNumber(dateValue) {
  const d = parseLocalDate(dateValue) || new Date();
  const weekStart = getZenniSundayStart(d);
  const jan1 = new Date(d.getFullYear(), 0, 1);
  const firstWeekStart = getZenniSundayStart(jan1);
  const diffDays = Math.floor((weekStart - firstWeekStart) / 86400000);
  return Math.floor(diffDays / 7) + 1;
}

function getZenniFWLabel(dateValue) {
  return 'FW ' + getZenniFiscalWeekNumber(dateValue);
}

function isCurrentFiscalWeek(weekStartDate, weekEndDate) {
  const today = new Date();
  const start = parseLocalDate(weekStartDate);
  const end   = parseLocalDate(weekEndDate);
  if (!start || !end) return false;
  const t = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return t >= start && t <= end;
}

// Re-derive FW label from weekStartDate + flag isCurrent accurately
function enrichBreakageWeek(w, isCurrent) {
  const fw = w.weekStartDate ? getZenniFWLabel(w.weekStartDate) : (w.fiscalWeek || '');
  return { ...w, fiscalWeek: fw, isCurrent: isCurrent === true };
}

// Build the synthetic current-week object when API has no records for it yet
function buildSyntheticCurrentWeek() {
  const weekStart = getZenniSundayStart(new Date());
  const weekEnd   = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);
  const startKey  = formatDateKey(weekStart);
  const endKey    = formatDateKey(weekEnd);
  const fw        = getZenniFWLabel(weekStart);
  return {
    key: `${fw}|${startKey}|${endKey}`,
    fiscalWeek: fw,
    weekStartDate: startKey,
    weekEndDate: endKey,
    rowCount: 0, workDateCount: 0, workDates: [], isCurrent: true
  };
}

// Returns all API weeks enriched with correct isCurrent + FW labels.
// Prepends a synthetic current week if the API doesn't include it yet.
function getBreakageFiscalWeeksEnriched() {
  const apiWeeks = STATE.boot?.fiscalWeeks || [];
  const enriched = apiWeeks.map(w => enrichBreakageWeek(w, isCurrentFiscalWeek(w.weekStartDate, w.weekEndDate)));

  if (!enriched.some(w => w.isCurrent)) {
    const synthetic = buildSyntheticCurrentWeek();
    // Check if the API already has this key under a different FW label
    const existingByDate = enriched.find(w => w.weekStartDate === synthetic.weekStartDate);
    if (existingByDate) {
      existingByDate.isCurrent = true;
    } else {
      enriched.unshift(synthetic);
    }
  }

  return enriched;
}

// Returns the current week object (isCurrent=true) or best fallback
function getCurrentBreakageFiscalWeek() {
  const weeks = getBreakageFiscalWeeksEnriched();
  return weeks.find(w => w.isCurrent) || weeks[0] || null;
}

/* RENDER MAIN */
function renderEverything() {
  renderTopSummary();
  renderReasonDropdown();
  renderAssociateDropdown();

  if (STATE.activePage === 'dashboard') return renderDashboardHub();
  if (STATE.activePage === 'associates') return renderAssociatesPage();
  if (STATE.activePage === 'summary') return renderSummaryPage();
  if (STATE.activePage === 'trends') return renderTrendsPage();
  renderDashboardHub();
}

function setActivePage(page) {
  STATE.activePage = page || 'dashboard';
  document.querySelectorAll('.side-link').forEach(btn => {
    btn.classList.toggle('active', (btn.dataset.page || '') === STATE.activePage);
  });
  renderEverything();
}

function renderTopSummary() {
  const s = STATE.week?.summary || {};
  const associates = getFilteredAssociates({ ignoreSearch: true, ignoreStatus: true });
  const watch = associates.filter(a => ['AMBER', 'RED'].includes(getAssociateStatus(a))).length;

  setText('kpiLens', formatNumber(s.lensBroken || 0));
  setText('kpiFrame', formatNumber(s.frameBroken || 0));
  setText('kpiTotal', formatNumber(s.totalBreakage || 0));
  setText('kpiWatch', formatNumber(watch));
  setText('kpiLensSub', `${formatPercent(s.lensBreakagePct || 0)} of scans`);
  setText('kpiFrameSub', `${formatPercent(s.frameBreakagePct || 0)} of scans`);
  setText('kpiRecordSub', `${formatNumber(s.recordCount || 0)} records`);
  setText('kpiAssociatesSub', `${formatNumber(s.associateCount || associates.length || 0)} associates`);
}

function renderDashboardHub() {
  const page = $('pageContent');
  if (!page) return;
  if (!STATE.week) return renderLoadingCard();

  const s = STATE.week.summary || {};
  const associates = getFilteredAssociates({ ignoreStatus: true });
  const green = associates.filter(a => getAssociateStatus(a) === 'GREEN').length;
  const amber = associates.filter(a => getAssociateStatus(a) === 'AMBER').length;
  const red = associates.filter(a => getAssociateStatus(a) === 'RED').length;
  const noData = associates.filter(a => getAssociateStatus(a) === 'NO_DATA').length;
  const riskPct = getRiskPercent(s);
  const ringStatus = getRiskStatusFromPercent(riskPct);
  const top = associates.slice().sort((a,b) => Number(b.totalBreakage || 0) - Number(a.totalBreakage || 0)).slice(0, 10);

  page.innerHTML = `
    <section class="hub-command-shell">
      <div class="hub-hero command-card">
        <div>          
          <h2>Breakage/Quality Overview</h2>
          <div class="hub-actions">
            <button class="primary-action" type="button" onclick="setActivePage('associates')">Open Operator Scorecards</button>
            <button class="ghost-action" type="button" onclick="loadBoot(true)">Refresh Live Data</button>
          </div>
        </div>
        <div class="quality-ring ${ringStatus}" style="--risk:${Math.min(100, Math.max(1, riskPct * 100))}">
          <div><strong>${formatPercent(riskPct)}</strong><span>Hub Breakage Risk</span></div>
        </div>
      </div>

      <div class="hub-status-grid">
        <article class="hub-status-card green"><span>Low Risk</span><strong>${green}</strong><small>Clean or controlled records</small></article>
        <article class="hub-status-card amber"><span>Watch</span><strong>${amber}</strong><small>Needs review</small></article>
        <article class="hub-status-card red"><span>Warning</span><strong>${red}</strong><small>High breakage count/rate</small></article>
        <article class="hub-status-card cyan"><span>No Data</span><strong>${noData}</strong><small>No meaningful scan base</small></article>
      </div>

      <section class="scoreboard-panel command-card">
        <div class="section-head">
          <div>
            <h3>Top Breakage Watchlist</h3>            
          </div>
          <span class="scoreboard-count">${top.length} shown</span>
        </div>
        <div class="scoreboard-grid">
          ${top.length ? top.map(renderAssociateCard).join('') : emptyState('No breakage records found for this view.')}
        </div>
      </section>
    </section>
  `;
}

function renderAssociatesPage() {
  const page = $('pageContent');
  if (!page) return;
  if (!STATE.week) return renderLoadingCard();

  const rows = getFilteredAssociates();
  const allRows = getFilteredAssociates({ ignoreStatus: true, ignoreSearch: true });
  const counts = {
    GREEN: allRows.filter(a => getAssociateStatus(a) === 'GREEN').length,
    AMBER: allRows.filter(a => getAssociateStatus(a) === 'AMBER').length,
    RED: allRows.filter(a => getAssociateStatus(a) === 'RED').length,
    NO_DATA: allRows.filter(a => getAssociateStatus(a) === 'NO_DATA').length
  };

  page.innerHTML = `
    <section class="scoreboard-panel command-card">
      <div class="section-head">
        <div>
          <h3>Operator Quality Scorecards</h3>          
        </div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;justify-content:flex-end">${activeFilterNote()}<span class="scoreboard-count">${rows.length} operators</span></div>
      </div>

      <div class="status-filter-row">
        <button class="status-pill ${STATE.activeStatus === 'ALL' ? 'active' : ''}" onclick="setStatusFilter('ALL')">All</button>
        <button class="status-pill green ${STATE.activeStatus === 'GREEN' ? 'active' : ''}" onclick="setStatusFilter('GREEN')">Low Risk ${counts.GREEN}</button>
        <button class="status-pill amber ${STATE.activeStatus === 'AMBER' ? 'active' : ''}" onclick="setStatusFilter('AMBER')">Watch ${counts.AMBER}</button>
        <button class="status-pill red ${STATE.activeStatus === 'RED' ? 'active' : ''}" onclick="setStatusFilter('RED')">Warning ${counts.RED}</button>
        <button class="status-pill ${STATE.activeStatus === 'NO_DATA' ? 'active' : ''}" onclick="setStatusFilter('NO_DATA')">No Data ${counts.NO_DATA}</button>
      </div>

      <div class="scoreboard-grid">
        ${rows.length ? rows.map(renderAssociateCard).join('') : emptyState('No operators found for this filter.')}
      </div>
    </section>
  `;
}

function renderAssociateCard(a) {
  const status = getAssociateStatus(a);
  const operator = a.operator || a.OperatorName || 'Unknown Operator';
  const dept = a.department || a.FinalDepartment || getAssociateDeptLabel(a);
  const total = Number(a.totalBreakage || 0);
  const lens = Number(a.lensBroken || 0);
  const frame = Number(a.frameBroken || 0);
  const pct = getAssociateRiskPercent(a);
  const access = Array.isArray(a.accessPoints) ? a.accessPoints.join(', ') : (a.accessPoint || a.AccessPoint || 'All Access Points');

  return `
    <button class="score-kpi ${status}" type="button" onclick="openModal('${escapeJsAttr(operator)}')">
      <span class="score-status-dot"></span>
      <span class="score-name">${escapeHtml(operator)}</span>
      <span class="score-dept">${escapeHtml(dept || 'Quality')} • ${escapeHtml(statusLabel(status))}</span>
      <strong>${formatNumber(total)}</strong>
      <span class="score-meta">${formatNumber(lens)} lens • ${formatNumber(frame)} frame • ${formatPercent(pct)} risk</span>
      <span class="score-meta" title="${escapeHtml(access)}">${escapeHtml(truncateText(access, 46))}</span>
    </button>
  `;
}

function renderSummaryPage() {
  const page = $('pageContent');
  if (!page) return;
  if (!STATE.week) return renderLoadingCard();

  const reasons = getFilteredReasons();
  const accessPoints = getFilteredAccessPoints();

  page.innerHTML = `
    <section class="summary-grid">
      <div class="command-card">
        <div class="section-head" style="padding:16px 16px 0">
          <div><h3>Top Reasons</h3><p>${reasons.length} unique reasons</p></div>
          <span class="scoreboard-count">${STATE.selectedWeek?.fiscalWeek || '—'}</span>
        </div>
        <div class="reason-list">
          ${reasons.length ? reasons.slice(0, 18).map(r => `
            <div class="reason-row" onclick="filterByReason('${escapeJsAttr(r.name)}')">
              <span class="reason-name" title="${escapeHtml(r.name)}">${escapeHtml(r.name || 'No Reason')}</span>
              <div class="reason-pills">
                <span class="pill l">${formatNumber(r.lensBroken || 0)} L</span>
                <span class="pill f">${formatNumber(r.frameBroken || 0)} F</span>
              </div>
            </div>`).join('') : emptyState('No reason data found.')}
        </div>
      </div>

      <div class="command-card">
        <div class="section-head" style="padding:16px 16px 0">
          <div><h3>Access Point Risk</h3><p>Click an access point to filter associate scorecards.</p></div>
        </div>
        <div class="ap-list">
          ${accessPoints.length ? accessPoints.slice(0, 24).map(ap => `
            <div class="ap-row-card" onclick="filterByAccessPoint('${escapeJsAttr(ap.name)}')">
              <span class="ap-name" title="${escapeHtml(ap.name)}">${escapeHtml(ap.name || 'No Access Point')}</span>
              <span class="ap-num l">${formatNumber(ap.lensBroken || 0)} L</span>
              <span class="ap-num f">${formatNumber(ap.frameBroken || 0)} F</span>
              <span class="ap-num">${formatNumber(ap.totalBreakage || ((ap.lensBroken || 0) + (ap.frameBroken || 0)))}</span>
            </div>`).join('') : emptyState('No access point data found.')}
        </div>
      </div>
    </section>
  `;
}

function renderTrendsPage() {
  const page = $('pageContent');
  if (!page) return;
  if (!STATE.week) return renderLoadingCard();

  const weekRows = buildLoadedWeekRows();
  const offenders = buildTopOffendersAcrossLoadedWeeks();
  const reasons = getFilteredReasons().slice(0, 8);
  const maxReason = reasons.reduce((m, r) => Math.max(m, Number(r.totalBreakage || r.lensBroken || 0)), 1);

  page.innerHTML = `
    <section class="trends-grid">
      <div class="command-card trends-wide">
        <div class="section-head" style="padding:16px 16px 0">
          <div><h3>Weekly Breakdown</h3><p>Switch fiscal weeks to load more history into this view.</p></div>
          <span class="scoreboard-count">${weekRows.length} loaded</span>
        </div>
        <div class="trends-table-wrap">
          <table class="trends-table">
            <thead>
              <tr>
                <th>Fiscal Week</th><th>Period</th><th>Lens</th><th>Lens %</th><th>Frame</th><th>Frame %</th><th>Total</th><th>Associates</th><th>Records</th>
              </tr>
            </thead>
            <tbody>
              ${weekRows.length ? weekRows.map(r => `
                <tr>
                  <td><span class="fw-badge">${escapeHtml(r.fiscalWeek)}</span></td>
                  <td>${formatPrettyDate(r.weekStartDate)} – ${formatPrettyDate(r.weekEndDate)}</td>
                  <td>${formatNumber(r.lens)}</td>
                  <td>${formatPercent(r.lensPct)}</td>
                  <td>${formatNumber(r.frame)}</td>
                  <td>${formatPercent(r.framePct)}</td>
                  <td>${formatNumber(r.total)}</td>
                  <td>${formatNumber(r.associates)}</td>
                  <td>${formatNumber(r.records)}</td>
                </tr>`).join('') : `<tr><td colspan="9">No loaded weeks yet.</td></tr>`}
            </tbody>
          </table>
        </div>
      </div>

      <div class="command-card">
        <div class="section-head" style="padding:16px 16px 0">
          <div><h3>Top Reasons</h3><p>Current selected week.</p></div>
        </div>
        <div style="padding:16px">
          ${reasons.length ? buildBarChart(reasons, maxReason) : emptyState('No reason data for this week.')}
        </div>
      </div>

      <div class="command-card">
        <div class="section-head" style="padding:16px 16px 0">
          <div><h3>Top Offenders</h3><p>Combined across loaded weeks.</p></div>
        </div>
        <div class="offender-list">
          ${offenders.length ? offenders.map((o, i) => `
            <div class="offender-row">
              <span class="offender-rank">${i + 1}</span>
              <span class="offender-name">${escapeHtml(o.operator)}</span>
              <span>${formatNumber(o.lens)} L</span>
              <span>${formatNumber(o.frame)} F</span>
              <span class="offender-count">${formatNumber(o.total)}</span>
            </div>`).join('') : emptyState('No offender data loaded yet.')}
        </div>
      </div>
    </section>
  `;
}

function renderLoadingCard() {
  const page = $('pageContent');
  if (!page) return;
  page.innerHTML = `<section class="loading-card command-loading-card"><div><div class="spinner"></div><strong>Loading Quality / Breakage records...</strong><br><small>Pulling selected week only.</small></div></section>`;
}

/* FILTERS */
function getFilteredAssociates(options = {}) {
  const search = options.ignoreSearch ? '' : STATE.search;
  const wantedReason = options.ignoreReason ? 'ALL' : (STATE.activeReason || 'ALL');
  const wantedAccess = options.ignoreAccess ? 'ALL' : (STATE.activeStationGroup || 'ALL');

  const allowedByReason = getOperatorsForReason(wantedReason);

  const rows = (STATE.week?.associates || []).filter(a => {
    const accessPoints = Array.isArray(a.accessPoints) ? a.accessPoints : [a.accessPoint || a.AccessPoint || ''].filter(Boolean);
    const reasons = Array.isArray(a.reasons) ? a.reasons : [a.reason || a.BreakageReason || ''].filter(Boolean);
    const operator = String(a.operator || a.OperatorName || '').trim();
    const combinedText = [operator, ...accessPoints, ...reasons].join(' ').toLowerCase();

    if (!options.ignoreSearch && search && !combinedText.includes(search)) return false;

    if (wantedReason !== 'ALL') {
      const directReasonMatch = reasons.some(r => String(r).toUpperCase() === String(wantedReason).toUpperCase());
      const mappedReasonMatch = allowedByReason.size ? allowedByReason.has(normalizeName(operator)) : false;
      // If API has no operator mapping for reasons, do not hide all operators.
      if ((allowedByReason.size || reasons.length) && !directReasonMatch && !mappedReasonMatch) return false;
    }

    if (wantedAccess !== 'ALL') {
      const wanted = String(wantedAccess).toUpperCase();
      const apMatch = accessPoints.some(ap => String(ap).toUpperCase() === wanted || String(ap).toUpperCase().includes(wanted));
      if (!apMatch) return false;
    }

    if (!options.ignoreStatus && STATE.activeStatus !== 'ALL' && getAssociateStatus(a) !== STATE.activeStatus) return false;
    if (STATE.selectedAssociate && normalizeName(operator) !== normalizeName(STATE.selectedAssociate)) return false;
    return true;
  });

  return rows.sort((a, b) => Number(b.totalBreakage || 0) - Number(a.totalBreakage || 0) || String(a.operator || '').localeCompare(String(b.operator || '')));
}

function getOperatorsForReason(reasonName) {
  const wanted = String(reasonName || 'ALL').trim();
  const set = new Set();
  if (!wanted || wanted === 'ALL') return set;

  const reasonRow = (STATE.week?.reasons || []).find(r => String(r.name || '').toUpperCase() === wanted.toUpperCase());
  if (reasonRow && Array.isArray(reasonRow.operators)) {
    reasonRow.operators.forEach(op => set.add(normalizeName(op)));
  }

  (STATE.week?.records || []).forEach(r => {
    if (String(r.reason || '').toUpperCase() === wanted.toUpperCase()) set.add(normalizeName(r.operator));
  });

  return set;
}

function getFilteredReasons() {
  const search = STATE.search;
  return (STATE.week?.reasons || []).filter(r => {
    const name = String(r.name || '').toLowerCase();
    if (search && !name.includes(search)) return false;
    return true;
  }).sort((a,b) => Number((b.totalBreakage ?? ((b.lensBroken||0)+(b.frameBroken||0)))) - Number((a.totalBreakage ?? ((a.lensBroken||0)+(a.frameBroken||0)))));
}

function getFilteredAccessPoints() {
  const search = STATE.search;
  return (STATE.week?.accessPoints || []).filter(ap => {
    const name = String(ap.name || '').toLowerCase();
    if (search && !name.includes(search)) return false;
    return true;
  }).sort((a,b) => Number((b.totalBreakage ?? ((b.lensBroken||0)+(b.frameBroken||0)))) - Number((a.totalBreakage ?? ((a.lensBroken||0)+(a.frameBroken||0)))));
}

function setStatusFilter(status) {
  STATE.activeStatus = status || 'ALL';
  renderEverything();
}

function filterByAccessPoint(ap) {
  STATE.activeStationGroup = ap || 'ALL';
  STATE.selectedAssociate = '';
  renderNavBar();
  setActivePage('associates');
}

function filterByReason(reason) {
  STATE.activeReason = reason || 'ALL';
  STATE.selectedAssociate = '';
  const sel = $('reasonSelect');
  if (sel) sel.value = STATE.activeReason === 'ALL' ? '' : STATE.activeReason;
  setActivePage('associates');
}

function clearQualityFilters() {
  STATE.activeReason = 'ALL';
  STATE.activeStationGroup = 'ALL';
  STATE.selectedAssociate = '';
  STATE.search = '';
  const reasonSelect = $('reasonSelect');
  const associateSelect = $('associateSelect');
  const searchInput = $('searchInput');
  if (reasonSelect) reasonSelect.value = '';
  if (associateSelect) associateSelect.value = '';
  if (searchInput) searchInput.value = '';
  renderNavBar();
  renderEverything();
}

/* NAV BAR */
function renderNavBar() {
  const bar = $('hubApBar');
  if (!bar) return;

  const accessPoints = (STATE.week?.accessPoints || [])
    .map(ap => String(ap.name || '').trim())
    .filter(Boolean)
    .sort();

  const current = STATE.activeStationGroup || 'ALL';
  const chips = accessPoints.slice(0, 18).map(ap => `
    <button class="hub-ap-chip ${current === ap ? 'active' : ''}" data-ap="${escapeHtml(ap)}" type="button">${escapeHtml(shortStationLabel(ap))}</button>
  `).join('');

  bar.innerHTML = `
    <span class="hub-ap-label">Access Point</span>
    <button class="hub-ap-chip ${current === 'ALL' ? 'active' : ''}" data-ap="ALL" type="button">All Access Points</button>
    ${chips}
  `;

  bar.querySelectorAll('.hub-ap-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      STATE.activeStationGroup = chip.dataset.ap || 'ALL';
      STATE.selectedAssociate = '';
      renderNavBar();
      renderEverything();
    });
  });
}

/* MODAL */
async function openModal(operator) {
  if (!operator || !STATE.selectedWeek) return;

  $('modalAssociateName').textContent = operator;
  $('modalAssociateSub').textContent = `${STATE.selectedWeek.fiscalWeek || ''} • ${formatPrettyDate(STATE.selectedWeek.weekStartDate)} – ${formatPrettyDate(STATE.selectedWeek.weekEndDate)}`;
  $('modalLens').textContent = '…';
  $('modalFrame').textContent = '…';
  $('modalTotal').textContent = '…';
  $('modalScan').textContent = '…';
  $('modalRecords').innerHTML = `<div class="command-loading-card"><div><div class="spinner"></div><strong>Loading associate records...</strong></div></div>`;
  $('associateModal')?.classList.remove('hidden');

  const params = new URLSearchParams({
    action: 'breakageAssociate',
    operator,
    fiscalWeek: STATE.selectedWeek.fiscalWeek || '',
    weekStartDate: STATE.selectedWeek.weekStartDate || '',
    weekEndDate: STATE.selectedWeek.weekEndDate || '',
    ts: String(Date.now())
  });

  try {
    const data = await fetchApiRaw(params.toString());
    const a = data.associate || {};
    const records = data.records || [];

    $('modalAssociateName').textContent = a.operator || operator;
    $('modalLens').textContent = formatNumber(a.lensBroken || 0);
    $('modalFrame').textContent = formatNumber(a.frameBroken || 0);
    $('modalTotal').textContent = formatNumber(a.totalBreakage || 0);
    $('modalScan').textContent = formatNumber(a.totalDailyScan || a.totalScans || 0);

    $('modalRecords').innerHTML = records.length ? records.map(rec => `
      <div class="record-row">
        <span class="rec-date">${escapeHtml(formatPrettyDate(rec.workDate || rec.WorkDate))}</span>
        <span class="rec-ap" title="${escapeHtml(rec.accessPoint || rec.AccessPoint || '')}">${escapeHtml(rec.accessPoint || rec.AccessPoint || 'No Access Point')}</span>
        <span class="rec-reason" title="${escapeHtml(rec.reason || rec.BreakageReason || '')}">${escapeHtml(rec.reason || rec.BreakageReason || 'No Reason')}</span>
        <span class="rec-num l">${formatNumber(rec.lensBroken || rec.LensBreakageCount || 0)} L</span>
        <span class="rec-num f">${formatNumber(rec.frameBroken || rec.FrameBreakageCount || 0)} F</span>
        <span class="rec-num">${formatNumber(rec.totalBreakage || rec.TotalBreakageCount || 0)}</span>
      </div>
    `).join('') : emptyState('No records found for this associate.');
  } catch (err) {
    console.error(err);
    $('modalRecords').innerHTML = emptyState(`Failed to load records. ${escapeHtml(err.message || err)}`);
  }
}

function closeModal() {
  $('associateModal')?.classList.add('hidden');
}

/* WEEK DROPDOWNS / META */
function renderWeekDropdowns() {
  // Use enriched weeks: FW labels recalculated client-side, isCurrent uses today's date
  const weeks = getBreakageFiscalWeeksEnriched();

  const opts = weeks.map(w => `
    <option value="${escapeHtml(w.key)}">
      ${escapeHtml(w.fiscalWeek || 'Week')} — ${formatPrettyDate(w.weekStartDate)} to ${formatPrettyDate(w.weekEndDate)}${w.rowCount ? ` • ${formatNumber(w.rowCount)} records` : ''}${w.isCurrent ? ' ★' : ''}
    </option>
  `).join('');

  [$('fiscalWeekSelect'), $('sideWeekSelect')].forEach(sel => {
    if (sel) sel.innerHTML = opts || '<option value="">No weeks found</option>';
  });

  // Auto-select: if STATE.selectedWeek is already set, keep it;
  // otherwise default to the current week (today's FW)
  if (!STATE.selectedWeek || !STATE._userChangedFiscalWeek) {
    const currentWeek = getCurrentBreakageFiscalWeek();
    if (currentWeek) {
      // Prefer an exact key match from enriched list so we pick up rowCount
      const match = weeks.find(w => w.key === currentWeek.key) || currentWeek;
      STATE.selectedWeek = match;
    }
  }

  syncWeekSelects();
}

function syncWeekSelects() {
  const key = STATE.selectedWeek?.key || '';
  [$('fiscalWeekSelect'), $('sideWeekSelect')].forEach(sel => {
    if (sel && sel.value !== key) sel.value = key;
  });
}

function updateMeta() {
  const week = STATE.selectedWeek || {};
  setText('sidebarWeekLabel', week.fiscalWeek || '—');
  setText('sidebarUpdated', `${formatPrettyDate(week.weekStartDate)} – ${formatPrettyDate(week.weekEndDate)}`);
  setText('weekStartLabel', formatPrettyDateLong(week.weekStartDate));
  setText('weekEndLabel', formatPrettyDateLong(week.weekEndDate));
  syncWeekSelects();
}

function renderReasonDropdown() {
  const select = $('reasonSelect');
  if (!select) return;
  const reasons = (STATE.week?.reasons || []).slice().sort((a,b) => Number(b.totalBreakage || 0) - Number(a.totalBreakage || 0));
  select.innerHTML = '<option value="">All Reasons</option>' + reasons.map(r => {
    const name = r.name || 'No Reason';
    const total = Number(r.totalBreakage || ((r.lensBroken || 0) + (r.frameBroken || 0)) || 0);
    return `<option value="${escapeHtml(name)}">${escapeHtml(name)} — ${formatNumber(total)} breakage</option>`;
  }).join('');
  select.value = STATE.activeReason === 'ALL' ? '' : STATE.activeReason;
}

function activeFilterNote() {
  const parts = [];
  if (STATE.activeReason && STATE.activeReason !== 'ALL') parts.push(`Reason: ${escapeHtml(STATE.activeReason)}`);
  if (STATE.activeStationGroup && STATE.activeStationGroup !== 'ALL') parts.push(`Access: ${escapeHtml(shortStationLabel(STATE.activeStationGroup))}`);
  if (!parts.length) return '';
  return `<span class="active-filter-note">${parts.join(' • ')} <button class="clear-filter-link" onclick="clearQualityFilters()" type="button">×</button></span>`;
}

function renderAssociateDropdown() {
  const select = $('associateSelect');
  if (!select) return;

  const rows = getFilteredAssociates({ ignoreStatus: true, ignoreSearch: true });
  if (!rows.length) {
    select.innerHTML = '<option value="">No operators found</option>'; 
    return;
  }

  select.innerHTML = `<option value="">All Operators</option>` + rows.map(a => {
    const name = a.operator || a.OperatorName || '';
    const total = Number(a.totalBreakage || 0);
    return `<option value="${escapeHtml(name)}">${escapeHtml(name)} — ${formatNumber(total)} total breakage</option>`;
  }).join('');

  if (STATE.selectedAssociate && rows.some(a => (a.operator || a.OperatorName) === STATE.selectedAssociate)) {
    select.value = STATE.selectedAssociate;
  } else {
    select.value = '';
  }
}

/* CHART / TREND HELPERS */
function buildLoadedWeekRows() {
  return STATE.loadedWeekOrder.map(key => {
    const w = STATE.allWeeks[key];
    const meta = (STATE.boot?.fiscalWeeks || []).find(f => f.key === key) || {};
    const s = w?.summary || {};
    return {
      fiscalWeek: meta.fiscalWeek || key.split('|')[0] || 'Week',
      weekStartDate: meta.weekStartDate || '',
      weekEndDate: meta.weekEndDate || '',
      lens: s.lensBroken || 0,
      frame: s.frameBroken || 0,
      total: s.totalBreakage || 0,
      associates: s.associateCount || 0,
      records: s.recordCount || 0,
      lensPct: s.lensBreakagePct || 0,
      framePct: s.frameBreakagePct || 0
    };
  }).sort((a,b) => String(a.weekStartDate).localeCompare(String(b.weekStartDate)));
}

function buildTopOffendersAcrossLoadedWeeks() {
  const map = new Map();
  Object.values(STATE.allWeeks || {}).forEach(w => {
    (w.associates || []).forEach(a => {
      const op = a.operator || a.OperatorName || 'Unknown Operator';
      if (!map.has(op)) map.set(op, { operator: op, lens: 0, frame: 0, total: 0 });
      const item = map.get(op);
      item.lens += Number(a.lensBroken || 0);
      item.frame += Number(a.frameBroken || 0);
      item.total += Number(a.totalBreakage || 0);
    });
  });
  return Array.from(map.values()).sort((a,b) => b.total - a.total).slice(0, 12);
}

function buildBarChart(reasons, maxVal) {
  const BAR_W = 420, BAR_H = 24, GAP = 12, PAD_LEFT = 150, PAD_RIGHT = 60;
  const totalH = Math.max(1, reasons.length) * (BAR_H + GAP);
  const bars = reasons.map((r, i) => {
    const value = Number(r.totalBreakage || r.lensBroken || 0);
    const fillW = Math.max(2, Math.round((value / Math.max(maxVal, 1)) * BAR_W));
    const y = i * (BAR_H + GAP);
    const name = truncateText(String(r.name || 'No Reason'), 20);
    return `
      <g>
        <text x="${PAD_LEFT - 8}" y="${y + BAR_H/2 + 4}" text-anchor="end" class="bar-label">${escapeHtml(name)}</text>
        <rect x="${PAD_LEFT}" y="${y}" width="${fillW}" height="${BAR_H}" rx="5" fill="url(#barGrad)" opacity=".9"/>
        <text x="${PAD_LEFT + fillW + 8}" y="${y + BAR_H/2 + 4}" class="bar-value">${formatNumber(value)}</text>
      </g>`;
  }).join('');

  return `
    <svg class="bar-chart" viewBox="0 0 ${PAD_LEFT + BAR_W + PAD_RIGHT} ${totalH}" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <linearGradient id="barGrad" x1="0" x2="1" y1="0" y2="0">
          <stop offset="0%" stop-color="#ef4444" stop-opacity=".9"/>
          <stop offset="100%" stop-color="#f59e0b" stop-opacity=".72"/>
        </linearGradient>
      </defs>
      ${bars}
    </svg>`;
}

/* STATUS LOGIC */
function getRiskPercent(summary) {
  const lens = Number(summary.lensBreakagePct || 0);
  const frame = Number(summary.frameBreakagePct || 0);
  return Math.max(lens, frame, 0);
}

function getRiskStatusFromPercent(pct) {
  const p = Number(pct || 0);
  if (p >= 0.02) return 'RED';
  if (p >= 0.01) return 'AMBER';
  return 'GREEN';
}

function getAssociateRiskPercent(a) {
  const lens = Number(a.lensBreakagePct || a.OperatorLensBreakagePercent || 0);
  const frame = Number(a.frameBreakagePct || a.OperatorFrameBreakagePercent || 0);
  return Math.max(lens, frame, 0);
}

function getAssociateStatus(a) {
  const total = Number(a.totalBreakage || 0);
  const scans = Number(a.totalDailyScan || a.totalScans || 0);
  const pct = getAssociateRiskPercent(a);

  if (scans <= 0 && total <= 0) return 'NO_DATA';
  if (total >= 8 || pct >= 0.02) return 'RED';
  if (total >= 3 || pct >= 0.01) return 'AMBER';
  return 'GREEN';
}

function statusLabel(status) {
  if (status === 'GREEN') return 'Low Risk';
  if (status === 'AMBER') return 'Watch';
  if (status === 'RED') return 'Warning';
  return 'No Data';
}

function getAssociateDeptLabel(a) {
  const aps = Array.isArray(a.accessPoints) ? a.accessPoints.join(' ') : String(a.accessPoint || '');
  const text = `${a.department || ''} ${aps}`.toUpperCase();
  if (text.includes('AR')) return 'AR';
  if (text.includes('SURFACE') || text.includes('54R') || text.includes('OTL') || text.includes('ODT')) return 'SURFACE';
  if (text.includes('FSV') || text.includes('SF SCAN') || text.includes('FRAME ONLY')) return 'INVENTORY';
  return a.department || 'FINISH';
}

/* API */
async function fetchApi(action) {
  return fetchApiRaw(`action=${action}`);
}

async function fetchApiRaw(query) {
  const resp = await fetch(`${API_URL}?${query}`, { cache: 'no-store' });
  if (!resp.ok) throw new Error(`API ${resp.status}`);
  const text = await resp.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (err) {
    throw new Error('API did not return JSON. Redeploy Apps Script and test the action URL.');
  }
  if (data?.ok === false) throw new Error(data.message || data.error || 'API ok=false');
  return data;
}

/* GENERAL HELPERS */
function setLoading(on) {
  STATE.isLoading = !!on;
  const btn = $('refreshBtn');
  if (btn) {
    btn.disabled = !!on;
    btn.textContent = on ? '↻ Loading…' : '⟳ Refresh';
  }
}

function showRefreshStatus(message) {
  const el = $('refreshStatus');
  if (!el) return;
  el.textContent = message || 'Ready';
}

function showError(message) {
  setLoading(false);
  const page = $('pageContent');
  if (page) page.innerHTML = `<section class="loading-card command-loading-card"><div><strong>Quality Hub Error</strong><br><small>${escapeHtml(message)}</small></div></section>`;
  showRefreshStatus('Error loading data');
  const loader = $('appLoader');
  if (loader) loader.classList.add('hide');
}

function emptyState(text) {
  return `<div class="empty-state">${escapeHtml(text || 'No data found.')}</div>`;
}

function setText(id, value) {
  const el = $(id);
  if (el) el.textContent = value;
}

function formatNumber(v) {
  return new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(Number(v || 0));
}

function formatPercent(v) {
  return `${(Number(v || 0) * 100).toFixed(2)}%`;
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const text = String(v).slice(0,10);
  const parts = text.split('-').map(Number);
  if (parts.length === 3 && parts.every(Boolean)) return new Date(parts[0], parts[1] - 1, parts[2]);
  const fb = new Date(v);
  return isNaN(fb.getTime()) ? null : fb;
}

function formatPrettyDate(v) {
  const d = parseDate(v);
  if (!d) return v || '';
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

function formatPrettyDateLong(v) {
  const d = parseDate(v);
  if (!d) return v || '--';
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

function escapeHtml(v) {
  return String(v ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function escapeJsAttr(v) {
  return String(v ?? '')
    .replaceAll('\\', '\\\\')
    .replaceAll("'", "\\'")
    .replaceAll('"', '&quot;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

function truncateText(text, max) {
  const t = String(text || '');
  return t.length > max ? t.slice(0, Math.max(0, max - 1)) + '…' : t;
}

function shortStationLabel(text) {
  const t = String(text || '').trim();
  if (t.length <= 18) return t;
  return t.slice(0, 17) + '…';
}

// Keep old function name working if anything still calls it.
function openAssociateModal(operator) {
  return openModal(operator);
}