/************************************************************
 * PRODUCTIVITY HUB - SCORECARD PROFILE JS
 * One operator profile at a time.
 * Normal webpage only. Apps Script is API only. Fast path uses ?action=boot and ?action=weekData.
 ************************************************************/

const API_URL = 'https://script.google.com/macros/s/AKfycbySXKQSHTXQVlZNq2vubqp2D3W-_IgtmaiFr_GDxz5X4FO2cqcYeUkAo_A9LajOfj9f/exec';
const CLIENT_CACHE_KEY = 'PRODUCTIVITY_HUB_SHEET_PAYLOAD_CACHE_V5';
const CLIENT_CACHE_TTL_MS = 15 * 60 * 1000;

const STATE = {
  meta: {},
  master: [],
  summary: [],
  dailySnapshot: [],
  editSnapshot: [],
  adjustments: [],
  breakageMaster: [],
  breakageSummary: [],
  breakageDailySnapshot: [],
  breakageEditSnapshot: [],
  breakageAdjustments: [],
  holidays: [],
  associateDayRules: [],
  associateSchedules: [],
  fiscalWeeks: [],
  weekDataCache: {},
  scorecardEmployeeType: 'FTE',
  activeDept: 'FINISH',
  activeShift: 'WEEKDAY',
  activeStationGroup: 'ALL',
  activePage: 'dashboard',
  selectedAssociate: '',
  selectedFiscalWeek: '',
  selectedWeekStartDate: '',
  selectedWeekEndDate: '',
  selectedBreakageDate: 'ALL',
  search: '',
  hasLoadedOnce: false,
  loadingTimer: null,
  currentUser: null,
  productivityLocks: [],
  productivityActivity: [],
  productivityActivityLastSeen: ''
};


/************************************************************
 * LOAD SPEED LOGGING
 * Shows client-side load seconds in console + Data updated status.
 ************************************************************/
const LOAD_TIMER = {
  bootStart: 0,
  weekStart: 0,
  breakageStart: 0
};

function startLoadTimer(label) {
  const key = `${label}Start`;
  LOAD_TIMER[key] = performance.now();
  console.info(`[Productivity Hub] ${label} load started`);
}

function endLoadTimer(label, meta = {}) {
  const key = `${label}Start`;
  const start = LOAD_TIMER[key] || performance.now();
  const seconds = ((performance.now() - start) / 1000).toFixed(2);

  const apiMs = Number(meta.apiMs || meta.serverMs || 0);
  const apiSeconds = apiMs ? (apiMs / 1000).toFixed(2) : '';
  const cacheHit = meta.cacheHit === true || (meta._cache && meta._cache.hit === true);

  const message = apiSeconds
    ? `${label} loaded in ${seconds}s client / ${apiSeconds}s API${cacheHit ? ' / cache hit' : ''}`
    : `${label} loaded in ${seconds}s${cacheHit ? ' / cache hit' : ''}`;

  console.info(`[Productivity Hub] ${message}`, meta);
  showRefreshStatus(message);

  return Number(seconds);
}

function getApiTimingMeta(data) {
  return {
    apiMs: Number(data && (data.apiMs || data.serverMs) || 0),
    cacheHit: !!(data && data._cache && data._cache.hit),
    cacheKey: data && data._cache ? data._cache.key : '',
    mode: data && data.mode ? data.mode : ''
  };
}


/************************************************************
 * BREAKAGE CLIENT CACHE
 * Makes Quality / Breakage Hub behave like Associate Scorecards:
 * once a week is loaded, opening it again is instant.
 ************************************************************/
const BREAKAGE_CLIENT_CACHE_PREFIX = 'PRODUCTIVITY_HUB_BREAKAGE_WEEK_V1_';
const BREAKAGE_CLIENT_CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour

function getBreakageWeekCacheKey(week) {
  const start = String((week && week.weekStartDate) || STATE.selectedWeekStartDate || STATE.meta.weekStartDate || '').trim();
  const end = String((week && week.weekEndDate) || STATE.selectedWeekEndDate || STATE.meta.weekEndDate || '').trim();
  return `${start}|${end}`;
}

function getBreakageLocalStorageKey(week) {
  return BREAKAGE_CLIENT_CACHE_PREFIX + getBreakageWeekCacheKey(week).replace(/[^a-zA-Z0-9_-]/g, '_');
}

function getCachedBreakageWeekData(week) {
  const weekKey = getBreakageWeekCacheKey(week);
  if (!weekKey || weekKey === '|') return null;

  const memory = STATE.weekDataCache && STATE.weekDataCache[weekKey];
  if (memory && (memory.breakageDailySnapshot || memory.breakageMaster || memory.breakageOperatorWeeklySummary)) {
    return {
      source: 'memory',
      breakageDailySnapshot: memory.breakageDailySnapshot || [],
      breakageMaster: memory.breakageMaster || [],
      breakageOperatorWeeklySummary: memory.breakageOperatorWeeklySummary || [],
      breakageEditSnapshot: memory.breakageEditSnapshot || [],
      breakageAdjustments: memory.breakageAdjustments || []
    };
  }

  try {
    const raw = localStorage.getItem(getBreakageLocalStorageKey(week));
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    if (!parsed || !parsed.savedAt || !parsed.data) return null;

    if ((Date.now() - Number(parsed.savedAt)) > BREAKAGE_CLIENT_CACHE_TTL_MS) {
      localStorage.removeItem(getBreakageLocalStorageKey(week));
      return null;
    }

    return {
      source: 'browser',
      ...(parsed.data || {})
    };
  } catch (err) {
    console.warn('[Productivity Hub] Bad breakage browser cache:', err);
    return null;
  }
}

function saveCachedBreakageWeekData(week, data) {
  const weekKey = getBreakageWeekCacheKey(week);
  if (!weekKey || weekKey === '|') return;

  const cacheData = {
    breakageDailySnapshot: data.breakageDailySnapshot || [],
    breakageMaster: data.breakageMaster || [],
    breakageOperatorWeeklySummary: data.breakageOperatorWeeklySummary || [],
    breakageEditSnapshot: data.breakageEditSnapshot || [],
    breakageAdjustments: data.breakageAdjustments || []
  };

  STATE.weekDataCache[weekKey] = {
    ...(STATE.weekDataCache[weekKey] || {}),
    dailySnapshot: STATE.dailySnapshot.slice(),
    editSnapshot: STATE.editSnapshot.slice(),
    adjustments: STATE.adjustments.slice(),
    ...cacheData
  };

  try {
    localStorage.setItem(getBreakageLocalStorageKey(week), JSON.stringify({
      savedAt: Date.now(),
      data: cacheData
    }));
  } catch (err) {
    console.warn('[Productivity Hub] Breakage browser cache skipped:', err);
  }
}

function applyCachedBreakageWeekData(week, cached) {
  if (!cached) return false;

  const combinedRows = normalizeBreakageRows([
    ...(cached.breakageOperatorWeeklySummary || []),
    ...(cached.breakageDailySnapshot || []),
    ...(cached.breakageMaster || [])
  ]);

  STATE.breakageDailySnapshot = dedupeBreakageRows([
    ...(STATE.breakageDailySnapshot || []),
    ...combinedRows
  ]);

  STATE.breakageEditSnapshot = normalizeBreakageRows(cached.breakageEditSnapshot || []);
  STATE.breakageAdjustments = normalizeBreakageRows(cached.breakageAdjustments || []);

  console.info(`[Productivity Hub] breakage loaded from ${cached.source || 'cache'} cache: ${week.fiscalWeek || getBreakageWeekCacheKey(week)}`);
  showRefreshStatus(`Breakage loaded from ${cached.source || 'cache'} cache`);
  renderEverything();
  return true;
}

document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  reloadData();
});

function bindEvents() {
  // Side nav page links
  document.querySelectorAll('.side-link').forEach(btn => {
    btn.addEventListener('click', () => {
      setActivePage(btn.dataset.page || 'dashboard');
    });
  });

  // Shift toggle buttons
  document.querySelectorAll('.hub-shift-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const shift = btn.dataset.shift || 'WEEKDAY';
      STATE.activeShift = shift;
      STATE.selectedAssociate = '';
      document.querySelectorAll('.hub-shift-btn').forEach(b => b.classList.remove('active','wd','we','all'));
      btn.classList.add('active');
      if (shift === 'WEEKDAY') btn.classList.add('wd');
      else if (shift === 'WEEKEND') btn.classList.add('we');
      else btn.classList.add('all');
      const label = document.getElementById('shiftMetaLabel');
      const shiftLabels = { WEEKDAY: 'Weekday', WEEKEND: 'Weekend', ALL: 'All Shifts' };
      if (label) label.innerHTML = `Shift: <strong>${shiftLabels[shift] || shift}</strong>`;
      STATE.activeStationGroup = 'ALL';
      renderNavBar();
      renderEverything();
    });
  });

  // Dept group tabs
  document.querySelectorAll('.hub-dept-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.hub-dept-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      STATE.activeDept = tab.dataset.dept || 'FINISH';
      STATE.selectedAssociate = '';
      STATE.activeStationGroup = 'ALL';
      const sideDept = document.getElementById('sideDeptSelect');
      if (sideDept) sideDept.value = STATE.activeDept;
      renderNavBar();
      renderEverything();
    });
  });

  // Sidebar dept select
  const sideDept = document.getElementById('sideDeptSelect');
  if (sideDept) {
    sideDept.addEventListener('change', e => {
      STATE.activeDept = e.target.value || 'FINISH';
      STATE.selectedAssociate = '';
      STATE.activeStationGroup = 'ALL';
      document.querySelectorAll('.hub-dept-tab').forEach(t => {
        t.classList.toggle('active', t.dataset.dept === STATE.activeDept);
      });
      renderNavBar();
      renderEverything();
    });
  }

  // Associate dropdown
  const associateSelect = document.getElementById('associateSelect');
  if (associateSelect) {
    associateSelect.addEventListener('change', e => {
      STATE.selectedAssociate = e.target.value || '';
      STATE.activePage = 'profile';
      setActiveSideNav('scorecards');
      renderEverything();
    });
  }

  // Search
  const searchInput = document.getElementById('searchInput');
  if (searchInput) {
    searchInput.addEventListener('input', e => {
      STATE.search = e.target.value.trim().toLowerCase();
      STATE.selectedAssociate = '';
      renderEverything();
    });
  }

  // Fiscal week dropdown
  const fiscalWeekSelect = document.getElementById('fiscalWeekSelect');
  if (fiscalWeekSelect) {
    fiscalWeekSelect.addEventListener('change', e => {
      const selectedKey = e.target.value || '';
      const week = getAvailableFiscalWeeks().find(w => w.key === selectedKey);

      if (!week) {
        STATE.selectedFiscalWeek = STATE.meta.fiscalWeek || '';
        STATE.selectedWeekStartDate = STATE.meta.weekStartDate || '';
        STATE.selectedWeekEndDate = STATE.meta.weekEndDate || '';
        STATE.selectedAssociate = '';
        updateMeta();
        renderNavBar();
        renderEverything();
        return;
      }

      loadFiscalWeekData(week);
    });
  }

  // Refresh button
  const refreshBtn = document.getElementById('refreshBtn');
  if (refreshBtn) refreshBtn.addEventListener('click', reloadData);
}

/************************************************************
 * HUB NAV: AP CHIP BAR
 * Builds StationGroup chips from rows matching current shift+dept.
 * Called after data loads and after shift/dept changes.
 ************************************************************/
function renderNavBar() {
  const bar = document.getElementById('hubApBar');
  if (!bar) return;

  // Build schedule map: operatorName -> scheduleType
  const scheduleMap = new Map();
  (STATE.associateSchedules || []).forEach(s => {
    scheduleMap.set(String(s.OperatorName || '').trim(), String(s.ScheduleType || 'WEEKDAY').trim().toUpperCase());
  });

  // Get all rows for active dept (no AP filter yet, shift filter applied)
  const startDate = STATE.selectedWeekStartDate || STATE.meta.weekStartDate;
  const endDate = STATE.selectedWeekEndDate || STATE.meta.weekEndDate;
  const start = parseLocalDate(startDate);
  const end = parseLocalDate(endDate);
  const activeDept = String(STATE.activeDept || 'FINISH').toUpperCase();
  const activeShift = String(STATE.activeShift || 'WEEKDAY').toUpperCase();

  const PROD_DEPTS_NAV = ['FINISH', 'SURFACE', 'AR'];
  const DIST_DEPTS_NAV = ['INVENTORY'];

  function navDeptMatch(r) {
    const rd = String(r.FinalDepartment || '').trim().toUpperCase();
    if (activeDept === 'PROD_ALL') return PROD_DEPTS_NAV.includes(rd);
    if (activeDept === 'DIST_ALL') return DIST_DEPTS_NAV.includes(rd);
    return rd === activeDept;
  }

  const allRows = dedupeDailyRows([...(STATE.dailySnapshot || [])]).filter(r => {
    const d = parseLocalDate(r.WorkDate);
    const dateMatch = d && start && end && d >= start && d <= end;
    const opName = String(r.OperatorName || '').trim();
    const opShift = scheduleMap.get(opName) || 'WEEKDAY';
    const rowShiftMatch = activeShift === 'ALL' || isMachineRow(r) || opShift === activeShift;
    return navDeptMatch(r) && dateMatch && rowShiftMatch;
  });

  // Collect unique StationGroups
  const groups = Array.from(new Set(allRows.map(r => String(r.StationGroup || '').trim()).filter(Boolean))).sort();

  const isInv = activeDept === 'INVENTORY';
  const chipClass = isInv ? 'hub-ap-chip inv' : 'hub-ap-chip';

  bar.innerHTML = `<span class="hub-ap-label">Station Group</span>
    <button class="${chipClass}${STATE.activeStationGroup === 'ALL' ? ' active' : ''}" data-ap="ALL" type="button">All Stations</button>
    ${groups.map(g => `<button class="${chipClass}${STATE.activeStationGroup === g ? ' active' : ''}" data-ap="${escapeHtml(g)}" type="button">${escapeHtml(g)}</button>`).join('')}`;

  // Wire AP chip clicks
  bar.querySelectorAll('.hub-ap-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      STATE.activeStationGroup = chip.dataset.ap || 'ALL';
      STATE.selectedAssociate = '';
      bar.querySelectorAll('.hub-ap-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      renderEverything();
    });
  });
}

function isMachineRow(r) {
  // Mirrors Apps Script isMachineOrPlaceholder_()
  // Machine/placeholder names are never in ASSOCIATE_SCHEDULES.
  // These rows show on BOTH shifts — they are equipment, not people.
  const n = String(r.OperatorName || '').toUpperCase();
  const ap = String(r.AccessPoint || '').toUpperCase();
  const sg = String(r.StationGroup || '').toUpperCase();

  // Named machine patterns (AccessPoint-based — matches Apps Script exactly)
  if (n.includes('UNKNOWN')) return true;
  if (n.includes('OTB'))        return true; // Blocking Line
  if (n.includes('ODT'))        return true; // Detaping Line
  if (n.includes('OTL'))        return true; // Engraving Line
  if (n.includes('ORB'))        return true; // Generating Line
  if (n.includes('FLEX'))       return true; // Polishing Line
  if (n.includes('MEI'))        return true; // MEI Line (Finish)
  if (n.includes('54R'))        return true; // Coating Line
  if (n.includes('AR41'))       return true; // Surface Inspection
  if (n.includes('IQ-STAR'))    return true; // Cooling Storage
  if (n.includes('MANUAL BLK')) return true; // Manual Block

  // Also catch by AccessPoint for rows where OperatorName may differ
  if (ap.includes('OTB'))        return true;
  if (ap.includes('ODT'))        return true;
  if (ap.includes('OTL'))        return true;
  if (ap.includes('ORB'))        return true;
  if (ap.includes('FLEX'))       return true;
  if (ap.includes('MEI'))        return true;
  if (ap.includes('54R'))        return true;
  if (ap.includes('AR41'))       return true;
  if (ap.includes('IQ-STAR'))    return true;
  if (ap.includes('MANUAL BLK')) return true;

  return false;
}

function getDeptLabel(dept) {
  const d = String(dept || '').toUpperCase();
  if (d === 'INVENTORY') return 'Picking';
  return d.charAt(0) + d.slice(1).toLowerCase();
}

function reloadData() {
  const cached = localStorage.getItem(CLIENT_CACHE_KEY) || localStorage.getItem('PRODUCTIVITY_HUB_CACHE_FAST_V2');

  if (cached && !STATE.hasLoadedOnce) {
    try {
      const parsed = JSON.parse(cached);
      const cachedData = parsed && parsed.data ? parsed.data : parsed;
      const savedAt = parsed && parsed.savedAt ? Number(parsed.savedAt) : 0;
      const cacheFresh = savedAt ? (Date.now() - savedAt) < CLIENT_CACHE_TTL_MS : true;

      if (cachedData && cacheFresh) {
        handleData(cachedData, true);
        console.info('[Productivity Hub] boot loaded from browser cache instantly');
      }
    } catch (err) {
      console.warn('Bad Productivity Hub cache:', err);
    }
  }

  if (!STATE.hasLoadedOnce) {
    startLoadingScreen();
    showLoading();
  } else {
    showRefreshStatus('Refreshing live data...');
  }

  startLoadTimer('boot');

  fetch(`${API_URL}?action=boot&ts=${Date.now()}`)
    .then(r => r.text())
    .then(text => {
      let data;
      try {
        data = JSON.parse(text);
      } catch (err) {
        throw new Error('API did not return JSON. Redeploy Apps Script and test /exec?action=boot.');
      }

      localStorage.setItem(CLIENT_CACHE_KEY, JSON.stringify({
        savedAt: Date.now(),
        data
      }));
      handleData(data, false);
      endLoadTimer('boot', getApiTimingMeta(data));

      // Quietly ask the API to warm the next cache paths.
      // This does not block the page.
      warmServerCacheQuietly();
    })
    .catch(showError);
}

function warmServerCacheQuietly() {
  fetch(`${API_URL}?action=warmCache&ts=${Date.now()}`)
    .catch(() => {});
}

function clearClientProductivityCache() {
  localStorage.removeItem(CLIENT_CACHE_KEY);
  localStorage.removeItem('PRODUCTIVITY_HUB_CACHE_FAST_V2');

  Object.keys(localStorage).forEach(key => {
    if (key.startsWith(BREAKAGE_CLIENT_CACHE_PREFIX)) {
      localStorage.removeItem(key);
    }
  });
}

function loadFiscalWeekData(week) {
  if (!week || !week.key) return;

  STATE.selectedFiscalWeek = week.fiscalWeek;
  STATE.selectedWeekStartDate = week.weekStartDate;
  STATE.selectedWeekEndDate = week.weekEndDate;
  STATE.selectedAssociate = '';

  const cached = STATE.weekDataCache[week.key];

  if (cached) {
    STATE.dailySnapshot = recalcClientProductivityRows(normalizeRows(cached.dailySnapshot || []));
    STATE.editSnapshot = normalizeRows(cached.editSnapshot || []);
    STATE.adjustments = normalizeRows(cached.adjustments || []);
    STATE.breakageDailySnapshot = normalizeBreakageRows([
      ...(cached.breakageOperatorWeeklySummary || []),
      ...(cached.breakageDailySnapshot || []),
      ...(cached.breakageMaster || [])
    ]);
    STATE.breakageEditSnapshot = normalizeBreakageRows(cached.breakageEditSnapshot || []);
    STATE.breakageAdjustments = normalizeBreakageRows(cached.breakageAdjustments || []);
    applyProductivityAdjustmentsToData();
    updateMeta();
    renderNavBar();
    renderEverything();
    console.info(`[Productivity Hub] week loaded from JS memory cache: ${week.fiscalWeek}`);
    showRefreshStatus(`Loaded ${week.fiscalWeek} from memory cache`);
    return;
  }

  showRefreshStatus(`Loading ${week.fiscalWeek}...`);
  startLoadTimer('week');

  const qs = new URLSearchParams({
    action: 'weekData',
    weekStartDate: week.weekStartDate,
    weekEndDate: week.weekEndDate,
    ts: String(Date.now())
  });

  fetch(`${API_URL}?${qs.toString()}`)
    .then(r => r.text())
    .then(text => {
      let data;
      try {
        data = JSON.parse(text);
      } catch (err) {
        throw new Error('API did not return JSON. Redeploy Apps Script and test /exec?action=weekData.');
      }

      if (!data || data.ok === false) {
        throw new Error(data && data.message ? data.message : 'Invalid weekData response.');
      }

      STATE.dailySnapshot = recalcClientProductivityRows(normalizeRows(data.dailySnapshot || []));
      STATE.editSnapshot = normalizeRows(data.editSnapshot || []);
      STATE.adjustments = normalizeRows(data.adjustments || []);
      STATE.breakageDailySnapshot = normalizeBreakageRows(data.breakageDailySnapshot || data.breakageSnapshot || []);
      STATE.breakageEditSnapshot = normalizeBreakageRows(data.breakageEditSnapshot || []);
      STATE.breakageAdjustments = normalizeBreakageRows(data.breakageAdjustments || []);

      STATE.weekDataCache[week.key] = {
        dailySnapshot: STATE.dailySnapshot.slice(),
        editSnapshot: STATE.editSnapshot.slice(),
        adjustments: STATE.adjustments.slice(),
        breakageOperatorWeeklySummary: normalizeBreakageRows(data.breakageOperatorWeeklySummary || []),
        breakageDailySnapshot: STATE.breakageDailySnapshot.slice(),
        breakageMaster: normalizeBreakageRows(data.breakageMaster || []),
        breakageEditSnapshot: STATE.breakageEditSnapshot.slice(),
        breakageAdjustments: STATE.breakageAdjustments.slice()
      };

      applyProductivityAdjustmentsToData();
      updateMeta();
      renderNavBar();
      renderEverything();
      endLoadTimer('week', getApiTimingMeta(data));
      if (STATE.activePage === 'breakage' && !getBreakageRowsForSelectedWeek().length) {
        fetchBreakageFallbackForWeek(week, 'weekData');
      }
    })
    .catch(showError);
}


function fetchBreakageFallbackForWeek(week, reason = 'fallback') {
  if (!week || !week.weekStartDate || !week.weekEndDate) return Promise.resolve();

  const existingRows = getBreakageRowsForSelectedWeek();
  if (existingRows.length) return Promise.resolve();

  const cachedBreakage = getCachedBreakageWeekData(week);
  if (cachedBreakage && applyCachedBreakageWeekData(week, cachedBreakage)) {
    return Promise.resolve(cachedBreakage);
  }

  showRefreshStatus(`Loading breakage records for ${week.fiscalWeek || 'selected week'}...`);
  startLoadTimer('breakage');

  if (STATE.activePage === 'breakage') {
    const page = document.getElementById('pageContent');
    if (page) {
      page.innerHTML = `
        <section class="quality-hub-shell">
          <div class="loading-card command-loading-card">
            <div class="spinner"></div>
            <strong>Loading Quality / Breakage records...</strong>
            <small>Pulling selected week only.</small>
          </div>
        </section>
      `;
    }
  }

  const qs = new URLSearchParams({
    action: 'breakageData',
    weekStartDate: week.weekStartDate,
    weekEndDate: week.weekEndDate,
    ts: String(Date.now())
  });

  return fetch(`${API_URL}?${qs.toString()}`)
    .then(r => r.text())
    .then(text => {
      let data;
      try {
        data = JSON.parse(text);
      } catch (err) {
        throw new Error('API did not return JSON. Redeploy Apps Script and test /exec?action=breakageData.');
      }

      if (!data || data.ok === false) {
        throw new Error(data && data.message ? data.message : 'Invalid breakageData response.');
      }

      const allBreakageRows = normalizeBreakageRows([
        ...(data.breakageOperatorWeeklySummary || []),
        ...(data.breakageDailySnapshot || []),
        ...(data.breakageMaster || [])
      ]);

      // Store in daily snapshot so getBreakageRowsForSelectedWeek can filter by WorkDate.
      STATE.breakageDailySnapshot = dedupeBreakageRows([
        ...(STATE.breakageDailySnapshot || []),
        ...allBreakageRows
      ]);

      saveCachedBreakageWeekData(week, {
        breakageOperatorWeeklySummary: data.breakageOperatorWeeklySummary || [],
        breakageDailySnapshot: data.breakageDailySnapshot || [],
        breakageMaster: data.breakageMaster || [],
        breakageEditSnapshot: data.breakageEditSnapshot || [],
        breakageAdjustments: data.breakageAdjustments || []
      });

      renderEverything();
      endLoadTimer('breakage', getApiTimingMeta(data));
    })
    .catch(showError);
}

function getCurrentFiscalWeekKeyFromMeta(meta) {
  return `${meta.fiscalWeek || ''}|${meta.weekStartDate || ''}|${meta.weekEndDate || ''}`;
}

function handleData(data, fromCache = false) {
  if (!data || data.ok === false) {
    throw new Error(data && data.message ? data.message : 'Invalid data response.');
  }

  STATE.hasLoadedOnce = true;

  STATE.meta = data.meta || {};
  STATE.master = recalcClientProductivityRows(normalizeRows(data.master || []));
  STATE.summary = normalizeRows(data.summary || []);
  STATE.dailySnapshot = recalcClientProductivityRows(normalizeRows(data.dailySnapshot || []));
  STATE.editSnapshot = normalizeRows(data.editSnapshot || []);
  STATE.adjustments = normalizeRows(data.adjustments || []);
  STATE.breakageMaster = normalizeBreakageRows(data.breakageMaster || data.breakage || []);
  STATE.breakageSummary = normalizeBreakageRows(data.breakageSummary || []);
  STATE.breakageDailySnapshot = normalizeBreakageRows(data.breakageDailySnapshot || data.breakageSnapshot || []);
  STATE.breakageEditSnapshot = normalizeBreakageRows(data.breakageEditSnapshot || []);
  STATE.breakageAdjustments = normalizeBreakageRows(data.breakageAdjustments || []);
  STATE.holidays = [];
  STATE.associateDayRules = [];
  STATE.associateSchedules = normalizeAssociateSchedules(data.associateSchedules || []);
  STATE.fiscalWeeks = normalizeFiscalWeeks(data.fiscalWeeks || []);

  const bootWeekKey = getCurrentFiscalWeekKeyFromMeta(data.meta || {});
  if (bootWeekKey) {
    STATE.weekDataCache[bootWeekKey] = {
      dailySnapshot: STATE.dailySnapshot.slice(),
      editSnapshot: STATE.editSnapshot.slice(),
      adjustments: STATE.adjustments.slice(),
      breakageDailySnapshot: STATE.breakageDailySnapshot.slice(),
      breakageEditSnapshot: STATE.breakageEditSnapshot.slice(),
      breakageAdjustments: STATE.breakageAdjustments.slice()
    };
  }

  applyProductivityAdjustmentsToData();

  updateMeta();
  renderNavBar();
  renderEverything();

  if (fromCache) {
    showRefreshStatus('Saved data loaded — refreshing live...');
  } else {
    showRefreshStatus(STATE.meta.generatedAt || 'Data updated');
    finishLoadingScreen();
    const currentWeek = {
      key: getCurrentFiscalWeekKey(),
      fiscalWeek: STATE.selectedFiscalWeek || STATE.meta.fiscalWeek,
      weekStartDate: STATE.selectedWeekStartDate || STATE.meta.weekStartDate,
      weekEndDate: STATE.selectedWeekEndDate || STATE.meta.weekEndDate
    };
    if (STATE.activePage === 'breakage' && !getBreakageRowsForSelectedWeek().length) {
      fetchBreakageFallbackForWeek(currentWeek, 'boot');
    }
  }
}

function normalizeRows(rows) {
  const numericFields = [
    'SessionTimeMins', 'SessionTimeHours', 'TargetAvgPerHour', 'AmberThreshold',
    'AveragePerMinute', 'AveragePerHour', 'ProductivityPercent', 'TotalJobScan',
    'H06','H07','H08','H09','H10','H11','H12','H13','H14','H15','H16','H17','H18','H19','H20',
    'Associate Count', 'Total Login Hours', 'Total Job Scan', 'Average JPH',
    'Average Productivity %', 'Green Count', 'Amber Count', 'Red Count', 'No Target Count',
    'OriginalSessionHours', 'AdjustedSessionHours', 'OriginalSessionMins', 'AdjustedSessionMins',
    'OriginalAveragePerHour', 'AdjustedAveragePerHour', 'OriginalProductivityPercent', 'AdjustedProductivityPercent',
    'WorkProduced',
    'LensBreakageCount', 'FrameBreakageCount', 'TotalBreakageCount',
    'OperatorDailyScans', 'OperatorLensBreakagePercent', 'OperatorFrameBreakagePercent',
    'DailyTotalScans', 'DailyLensBreakagePercent', 'DailyFrameBreakagePercent'
  ];

  return rows.map(row => {
    const copy = { ...row };
    numericFields.forEach(k => {
      if (copy[k] !== undefined) copy[k] = Number(copy[k] || 0);
    });
    return copy;
  });
}


function normalizeBreakageRows(rows) {
  return normalizeRows(rows || []).map(row => {
    const copy = { ...row };

    copy.WorkDate = copy.WorkDate || copy.workDate || copy.Date || copy.date || '';
    copy.FiscalWeek = copy.FiscalWeek || copy.fiscalWeek || '';
    copy.WeekStartDate = copy.WeekStartDate || copy.weekStartDate || '';
    copy.WeekEndDate = copy.WeekEndDate || copy.weekEndDate || '';
    copy.AccessPoint = String(copy.AccessPoint || copy.accessPoint || 'No Access Point').trim();
    copy.BreakageReason = String(copy.BreakageReason || copy.breakageReason || copy.Reason || 'NO BREAKAGE').trim();
    copy.OperatorName = String(copy.OperatorName || copy.operatorName || 'No Operator').trim();

    copy.LensBreakageCount = Number(copy.LensBreakageCount || copy['Lenses Broken'] || 0);
    copy.FrameBreakageCount = Number(copy.FrameBreakageCount || copy['Frames Broken'] || 0);
    copy.TotalBreakageCount = Number(copy.TotalBreakageCount || copy.Total || copy.LensBreakageCount + copy.FrameBreakageCount || 0);

    copy.OperatorDailyScans = Number(copy.OperatorDailyScans || copy.DailyTotalScans || 0);
    copy.OperatorLensBreakagePercent = Number(copy.OperatorLensBreakagePercent || copy.DailyLensBreakagePercent || 0);
    copy.OperatorFrameBreakagePercent = Number(copy.OperatorFrameBreakagePercent || copy.DailyFrameBreakagePercent || 0);
    copy.QualityStatus = String(copy.QualityStatus || copy.qualityStatus || getQualityStatus(copy)).trim();

    return copy;
  });
}


function normalizeHolidayRows(rows) {
  return rows
    .filter(row => row && row.Date)
    .map(row => ({
      Date: row.Date,
      Type: String(row.Type || 'HOLIDAY').trim().toUpperCase(),
      AppliesTo: String(row.AppliesTo || 'ALL').trim().toUpperCase(),
      Label: String(row.Label || row.Type || 'Excluded Day').trim(),
      ExcludeFromWeeklyWorkDays: String(row.ExcludeFromWeeklyWorkDays || 'YES').trim().toUpperCase(),
      CountIfWorked: String(row.CountIfWorked || 'YES').trim().toUpperCase(),
      Notes: String(row.Notes || '').trim()
    }));
}

function normalizeAssociateDayRules(rows) {
  return rows
    .filter(row => row && row.WorkDate && row.OperatorName)
    .filter(row => String(row.Active || 'TRUE').trim().toUpperCase() !== 'FALSE')
    .map(row => ({
      RuleID: String(row.RuleID || '').trim(),
      CreatedAt: row.CreatedAt || '',
      WorkDate: row.WorkDate,
      OperatorName: String(row.OperatorName || '').trim(),
      FinalDepartment: String(row.FinalDepartment || 'ALL').trim().toUpperCase(),
      Type: String(row.Type || 'NO_OT').trim().toUpperCase(),
      Label: String(row.Label || row.Type || 'Associate Excluded Day').trim(),
      ExcludeFromWeeklyWorkDays: String(row.ExcludeFromWeeklyWorkDays || 'YES').trim().toUpperCase(),
      CountIfWorked: String(row.CountIfWorked || 'YES').trim().toUpperCase(),
      Note: String(row.Note || '').trim(),
      CreatedBy: String(row.CreatedBy || '').trim(),
      Active: String(row.Active || 'TRUE').trim().toUpperCase()
    }));
}


function normalizeAssociateSchedules(rows) {
  return (rows || [])
    .filter(row => row && row.OperatorName)
    .filter(row => String(row.Active || 'TRUE').trim().toUpperCase() !== 'FALSE')
    .map(row => {
      const scheduleType = String(row.ScheduleType || 'WEEKDAY').trim().toUpperCase();
      const scheduledDays = String(row.ScheduledDays || defaultDaysForScheduleType(scheduleType)).trim().toUpperCase();

      return {
        OperatorName: String(row.OperatorName || '').trim(),
        ScheduleType: scheduleType,
        ScheduledDays: scheduledDays,
        Active: String(row.Active || 'TRUE').trim().toUpperCase(),
        Notes: String(row.Notes || '').trim()
      };
    });
}


function normalizeFiscalWeeks(rows) {
  return (rows || [])
    .filter(row => row && row.fiscalWeek && row.weekStartDate && row.weekEndDate)
    .map(row => ({
      key: String(row.key || `${row.fiscalWeek}|${row.weekStartDate}|${row.weekEndDate}`).trim(),
      fiscalWeek: String(row.fiscalWeek || '').trim(),
      weekStartDate: row.weekStartDate,
      weekEndDate: row.weekEndDate,
      rowCount: Number(row.rowCount || 0),
      latestSnapshotCreatedAt: row.latestSnapshotCreatedAt || ''
    }));
}

function defaultDaysForScheduleType(scheduleType) {
  const type = String(scheduleType || 'WEEKDAY').trim().toUpperCase();
  if (type === 'WEEKEND') return 'FRI,SAT,SUN';
  return 'MON,TUE,WED,THU';
}

function updateMeta() {
  if (!STATE.selectedFiscalWeek) {
    const weeks = getAvailableFiscalWeeks();
    const currentKey = getCurrentFiscalWeekKey();
    const currentWeek = weeks.find(w => w.key === currentKey) || weeks[0];

    if (currentWeek) {
      STATE.selectedFiscalWeek = currentWeek.fiscalWeek;
      STATE.selectedWeekStartDate = currentWeek.weekStartDate;
      STATE.selectedWeekEndDate = currentWeek.weekEndDate;
    } else {
      STATE.selectedFiscalWeek = STATE.meta.fiscalWeek || '';
      STATE.selectedWeekStartDate = STATE.meta.weekStartDate || '';
      STATE.selectedWeekEndDate = STATE.meta.weekEndDate || '';
    }
  }

  setText('workDateLabel', STATE.meta.workDate || '--');
  setText('fiscalWeekLabel', STATE.selectedFiscalWeek || STATE.meta.fiscalWeek || '--');
  setText('weekStartLabel', formatDatePretty(STATE.selectedWeekStartDate || STATE.meta.weekStartDate));
  setText('weekEndLabel', formatDatePretty(STATE.selectedWeekEndDate || STATE.meta.weekEndDate));
  setText('generatedAt', STATE.meta.generatedAt || '--');

  renderFiscalWeekDropdown();
}

function getCurrentFiscalWeekKey() {
  return `${STATE.meta.fiscalWeek || ''}|${STATE.meta.weekStartDate || ''}|${STATE.meta.weekEndDate || ''}`;
}

function getSelectedFiscalWeekKey() {
  return `${STATE.selectedFiscalWeek || ''}|${STATE.selectedWeekStartDate || ''}|${STATE.selectedWeekEndDate || ''}`;
}

function getAvailableFiscalWeeks() {
  const map = new Map();

  (STATE.fiscalWeeks || []).forEach(row => {
    const fiscalWeek = String(row.fiscalWeek || '').trim();
    const weekStartDate = row.weekStartDate || '';
    const weekEndDate = row.weekEndDate || '';
    if (!fiscalWeek || !weekStartDate || !weekEndDate) return;

    const key = row.key || `${fiscalWeek}|${weekStartDate}|${weekEndDate}`;
    map.set(key, { key, fiscalWeek, weekStartDate, weekEndDate, rowCount: Number(row.rowCount || 0) });
  });

  (STATE.dailySnapshot || []).forEach(row => {
    const fiscalWeek = String(row.FiscalWeek || '').trim();
    const weekStartDate = row.WeekStartDate || '';
    const weekEndDate = row.WeekEndDate || '';
    if (!fiscalWeek || !weekStartDate || !weekEndDate) return;

    const key = `${fiscalWeek}|${weekStartDate}|${weekEndDate}`;
    if (!map.has(key)) {
      map.set(key, { key, fiscalWeek, weekStartDate, weekEndDate, rowCount: 0 });
    }
  });

  if (STATE.meta.fiscalWeek && STATE.meta.weekStartDate && STATE.meta.weekEndDate) {
    const key = getCurrentFiscalWeekKey();
    if (!map.has(key)) {
      map.set(key, {
        key,
        fiscalWeek: STATE.meta.fiscalWeek,
        weekStartDate: STATE.meta.weekStartDate,
        weekEndDate: STATE.meta.weekEndDate,
        rowCount: STATE.master.length || 0
      });
    }
  }

  return Array.from(map.values()).sort((a, b) => {
    const diff = parseLocalDate(b.weekStartDate) - parseLocalDate(a.weekStartDate);
    if (diff !== 0) return diff;
    return String(b.fiscalWeek).localeCompare(String(a.fiscalWeek));
  });
}

function renderFiscalWeekDropdown() {
  const select = document.getElementById('fiscalWeekSelect');
  if (!select) return;

  const weeks = getAvailableFiscalWeeks();

  if (!weeks.length) {
    select.innerHTML = `<option value="">${escapeHtml(STATE.meta.fiscalWeek || 'Current Week')}</option>`;
    return;
  }

  const selectedKey = getSelectedFiscalWeekKey();

  select.innerHTML = weeks.map(w => {
    const rowText = w.rowCount ? ` • ${formatNumber(w.rowCount)} rows` : '';
    return `
      <option value="${escapeHtml(w.key)}" ${w.key === selectedKey ? 'selected' : ''}>
        ${escapeHtml(w.fiscalWeek)} — ${formatDatePretty(w.weekStartDate)} to ${formatDatePretty(w.weekEndDate)}${rowText}
      </option>
    `;
  }).join('');
}

function renderEverything() {
  renderTopSummary();
  renderAssociateDropdown();

  if (STATE.activePage === 'dashboard') {
    renderDashboardHub();
    return;
  }

  if (STATE.activePage === 'scorecards') {
    renderScorecardsPage();
    return;
  }

  if (STATE.activePage === 'profile') {
    renderProfile();
    return;
  }

  if (STATE.activePage === 'daily') {
    renderDailySummaryHub();
    return;
  }

  if (STATE.activePage === 'breakage') {
    renderBreakageHub();
    return;
  }

  if (STATE.activePage === 'trends') {
    renderTrendsHub();
    return;
  }

  if (STATE.activePage === 'targets') {
    renderTargetsHub();
    return;
  }

  renderPlaceholderPage(STATE.activePage);
}

function setActivePage(page) {
  STATE.activePage = page || 'dashboard';
  if (STATE.activePage !== 'profile') {
    STATE.selectedAssociate = STATE.selectedAssociate || '';
  }
  setActiveSideNav(STATE.activePage === 'profile' ? 'scorecards' : STATE.activePage);
  renderEverything();

  // Performance fix:
  // Do not load breakage data during normal Productivity boot/week switching.
  // Quality data loads only when the Quality / Breakage Hub is opened.
  if (STATE.activePage === 'breakage' && !getBreakageRowsForSelectedWeek().length) {
    fetchBreakageFallbackForWeek({
      key: getSelectedFiscalWeekKey() || getCurrentFiscalWeekKey(),
      fiscalWeek: STATE.selectedFiscalWeek || STATE.meta.fiscalWeek,
      weekStartDate: STATE.selectedWeekStartDate || STATE.meta.weekStartDate,
      weekEndDate: STATE.selectedWeekEndDate || STATE.meta.weekEndDate
    }, 'openBreakage');
  }
}

function setActiveSideNav(page) {
  document.querySelectorAll('.side-link').forEach(btn => {
    btn.classList.toggle('active', (btn.dataset.page || '') === page);
  });
}

function renderTopSummary() {
  const summaries = buildAssociateSummaries();

  const totalAssociates = summaries.filter(s => s.workProduced > 0).length;
  const totalJobs = summaries.reduce((sum, s) => sum + s.totalJobs, 0);
  const rated = summaries.filter(s => s.workProduced > 0);
  const avgProductivity = rated.length
    ? rated.reduce((sum, s) => sum + s.productivityPercent, 0) / rated.length
    : 0;
  const watchWarning = summaries.filter(s => s.status === 'AMBER' || s.status === 'RED').length;

  setText('totalAssociates', totalAssociates);
  setText('totalJobs', formatNumber(totalJobs));
  setText('avgProductivity', formatPercent(avgProductivity));
  setText('watchWarning', watchWarning);
}

function renderAssociateDropdown() {
  const select = document.getElementById('associateSelect');
  if (!select) return;

  const summaries = buildAssociateSummaries();

  if (!summaries.length) {
    select.innerHTML = '<option value="">No associates found</option>';
    STATE.selectedAssociate = '';
    return;
  }

  select.innerHTML = summaries.map(s => `
    <option value="${escapeHtml(s.associate)}">
      ${escapeHtml(s.associate)} — ${formatScheduleType(s.scheduleType)} — ${s.areas.length} area${s.areas.length === 1 ? '' : 's'} — ${s.workDays} day${s.workDays === 1 ? '' : 's'}
    </option>
  `).join('');

  if (!STATE.selectedAssociate || !summaries.some(s => s.associate === STATE.selectedAssociate)) {
    STATE.selectedAssociate = summaries[0].associate;
  }

  select.value = STATE.selectedAssociate;
}


function renderDashboardHub() {
  const page = document.getElementById('pageContent');
  if (!page) return;

  // Dashboard Hub is now department-top-scorecards only.
  // It does NOT depend on the selected sidebar department.
  // It uses the selected fiscal week + selected shift.
  const groups = buildDashboardTopScorecardGroups();
  const allTopRows = groups.flatMap(g => g.top || []);

  const green = allTopRows.filter(s => s.status === 'GREEN').length;
  const amber = allTopRows.filter(s => s.status === 'AMBER').length;
  const red = allTopRows.filter(s => s.status === 'RED').length;
  const noTarget = allTopRows.filter(s => s.status === 'NO TARGET').length;

  const totalHours = allTopRows.reduce((sum, s) => sum + Number(s.totalHours || 0), 0);
  const totalProduced = allTopRows.reduce((sum, s) => sum + Number(s.workProduced || 0), 0);
  const productivity = totalHours > 0 ? totalProduced / totalHours : 0;

  page.innerHTML = `
    <section class="hub-command-shell">
      <div class="hub-hero command-card">
        <div>
          <span class="profile-tag">Productivity Hub</span>
          <h2>Top Performance Scorecards</h2>          
          <div class="hub-actions">
            <button class="primary-action" type="button" onclick="setActivePage('scorecards')">Open Associate Scorecards</button>
            <button class="ghost-action" type="button" onclick="reloadData()">Refresh Live Data</button>
          </div>
        </div>
        <div class="hub-productivity-ring ${statusCss(getStatusFromPercent(productivity, totalProduced))}">
          <strong>${totalProduced > 0 ? formatPercent(productivity) : 'N/A'}</strong>
          <span>Hub Productivity</span>
        </div>
      </div>

      <div class="hub-status-grid">
        <article class="hub-status-card green"><span>On Target</span><strong>${green}</strong><small>100% or higher</small></article>
        <article class="hub-status-card amber"><span>Watch</span><strong>${amber}</strong><small>90% to 99%</small></article>
        <article class="hub-status-card red"><span>Warning</span><strong>${red}</strong><small>Below 90%</small></article>
        <article class="hub-status-card gray"><span>No Target</span><strong>${noTarget}</strong><small>Target missing or zero</small></article>
      </div>

      ${groups.map(group => `
        <section class="hub-panel command-card dept-top-score-section">
          <div class="section-head compact">
            <div>
              <h3>${escapeHtml(group.label)} Top Scorecards</h3>
              <p>${group.associateCount} associates • ${formatNumber(group.jobs)} jobs • ${group.workProduced > 0 ? formatPercent(group.productivityPercent) : 'N/A'} department productivity</p>
            </div>
            <button class="small-action" type="button" onclick="jumpToDeptScorecards('${escapeJs(group.department)}')">View ${escapeHtml(group.label)}</button>
          </div>
          <div class="top-score-strip">
            ${group.top.length ? group.top.map(s => `
              <button class="top-score-card ${statusCss(s.status)}" type="button" onclick="selectAssociateFromScoreboard('${escapeJs(s.associate)}')">
                <span class="score-status-dot"></span>
                <strong>${escapeHtml(s.associate)}</strong>
                <b>${s.workProduced > 0 ? formatPercent(s.productivityPercent) : 'N/A'}</b>
                <small>${formatDecimal(s.totalHours)} hrs • ${formatDecimal(s.workProduced)} produced • ${formatNumber(s.totalJobs)} jobs</small>
              </button>
            `).join('') : '<div class="empty-hub-note">No scorecards available for this department with the current week/shift/search filter.</div>'}
          </div>
        </section>
      `).join('')}
    </section>
  `;
}

function buildDashboardTopScorecardGroups() {
  const wanted = [
    { department: 'FINISH', label: 'Finish' },
    { department: 'AR', label: 'AR' },
    { department: 'INVENTORY', label: 'Picking' }
  ];

  return wanted.map(info => {
    const rows = getWeekRowsForDashboardDepartment(info.department);
    const summaries = buildAssociateSummariesFromRows(rows)
      .filter(s => Number(s.workProduced || 0) > 0)
      .sort((a, b) => Number(b.productivityPercent || 0) - Number(a.productivityPercent || 0));

    const totalHours = summaries.reduce((sum, s) => sum + Number(s.totalHours || 0), 0);
    const jobs = summaries.reduce((sum, s) => sum + Number(s.totalJobs || 0), 0);
    const workProduced = summaries.reduce((sum, s) => sum + Number(s.workProduced || 0), 0);
    const productivityPercent = totalHours > 0 ? workProduced / totalHours : 0;

    return {
      ...info,
      associateCount: summaries.length,
      hours: totalHours,
      jobs,
      workProduced,
      productivityPercent,
      status: getStatusFromPercent(productivityPercent, workProduced),
      top: summaries.slice(0, 6)
    };
  });
}

function getWeekRowsForDashboardDepartment(department) {
  const dept = String(department || '').trim().toUpperCase();
  const activeShift = String(STATE.activeShift || 'WEEKDAY').trim().toUpperCase();

  const startDate = STATE.selectedWeekStartDate || STATE.meta.weekStartDate;
  const endDate = STATE.selectedWeekEndDate || STATE.meta.weekEndDate;
  const start = parseLocalDate(startDate);
  const end = parseLocalDate(endDate);

  const scheduleMap = new Map();
  (STATE.associateSchedules || []).forEach(s => {
    scheduleMap.set(String(s.OperatorName || '').trim(), String(s.ScheduleType || 'WEEKDAY').trim().toUpperCase());
  });

  function shiftMatch(r) {
    if (activeShift === 'ALL') return true;
    if (isMachineRow(r)) return true;
    const opName = String(r.OperatorName || '').trim();
    const opShift = scheduleMap.get(opName) || 'WEEKDAY';
    return opShift === activeShift;
  }

  const rows = (STATE.dailySnapshot || []).filter(r => {
    const rowDept = String(r.FinalDepartment || '').trim().toUpperCase();
    const d = parseLocalDate(r.WorkDate);
    const dateMatch = d && start && end && d >= start && d <= end;
    const blob = [r.OperatorName, r.StationGroup, r.AccessPoint, r.FinalDepartment].join(' ').toLowerCase();
    const searchMatch = !STATE.search || blob.includes(STATE.search);
    return rowDept === dept && dateMatch && shiftMatch(r) && searchMatch;
  });

  return dedupeDailyRows(rows);
}

function buildAssociateSummariesFromRows(rows) {
  const map = new Map();

  (rows || []).forEach(row => {
    const name = String(row.OperatorName || 'Unknown Operator').trim();
    if (!name || isMachineRow(row)) return;
    if (!map.has(name)) map.set(name, []);
    map.get(name).push(row);
  });

  return Array.from(map.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([associate, associateRows]) => {
      const areas = buildAreaSummaries(associateRows);
      const totalHours = associateRows.reduce((sum, r) => sum + Number(r.SessionTimeHours || 0), 0);
      const totalJobs = associateRows.reduce((sum, r) => sum + Number(r.TotalJobScan || 0), 0);
      const avgJph = totalHours > 0 ? totalJobs / totalHours : 0;
      const workProduced = getWorkProducedHours(associateRows);
      const weightedTarget = getWeightedTarget(associateRows);
      const productivityPercent = totalHours > 0 && workProduced > 0 ? workProduced / totalHours : 0;
      const status = getStatusFromPercent(productivityPercent, workProduced);
      const workedDateKeys = new Set(associateRows.map(r => formatDateKey(r.WorkDate)).filter(Boolean));
      const workDays = workedDateKeys.size;
      const departments = Array.from(new Set(associateRows.map(r => r.FinalDepartment).filter(Boolean))).join(', ');
      const workDayInfo = getWeeklyWorkDayInfo(departments || STATE.activeDept, associate);
      const scheduledDaysWorked = (workDayInfo.eligibleDateKeys || [])
        .filter(dateKey => workedDateKeys.has(dateKey)).length;

      return {
        associate,
        departments,
        rows: associateRows,
        areas,
        totalHours,
        totalJobs,
        avgJph,
        weightedTarget,
        workProduced,
        productivityPercent,
        status,
        workDays,
        scheduledDaysWorked,
        scheduleType: workDayInfo.scheduleType,
        scheduledDays: workDayInfo.scheduledDays,
        scheduledDayGoal: workDayInfo.scheduledDayGoal,
        eligibleWorkDays: workDayInfo.eligibleWorkDays,
        excludedDays: workDayInfo.excludedDays,
        excludedDates: workDayInfo.excludedDates,
        eligibleDateKeys: workDayInfo.eligibleDateKeys,
        scheduledDateKeys: workDayInfo.scheduledDateKeys
      };
    });
}

function buildDepartmentHubRows(summaries) {
  const map = new Map();

  summaries.forEach(s => {
    const depts = String(s.departments || 'UNKNOWN').split(',').map(x => x.trim()).filter(Boolean);
    const primary = depts.length === 1 ? depts[0] : (depts[0] || 'MIXED');
    if (!map.has(primary)) {
      map.set(primary, { department: primary, associates: 0, hours: 0, jobs: 0, workProduced: 0 });
    }
    const d = map.get(primary);
    d.associates += 1;
    d.hours += Number(s.totalHours || 0);
    d.jobs += Number(s.totalJobs || 0);
    d.workProduced += Number(s.workProduced || 0);
  });

  return Array.from(map.values()).map(d => {
    const productivityPercent = d.hours > 0 ? d.workProduced / d.hours : 0;
    const status = getStatusFromPercent(productivityPercent, d.workProduced);
    return { ...d, productivityPercent, status };
  }).sort((a, b) => String(a.department).localeCompare(String(b.department)));
}

function jumpToDeptScorecards(department) {
  const dept = String(department || 'FINISH').trim().toUpperCase();
  STATE.activeDept = dept;
  STATE.activeStationGroup = 'ALL';
  STATE.selectedAssociate = '';

  const sideDept = document.getElementById('sideDeptSelect');
  if (sideDept) sideDept.value = dept;

  document.querySelectorAll('.hub-dept-tab').forEach(t => {
    t.classList.toggle('active', String(t.dataset.dept || '').toUpperCase() === dept);
  });

  renderNavBar();
  setActivePage('scorecards');
}


function getAssociateEmployeeType(summary) {
  const name = String(summary && summary.associate ? summary.associate : '').trim().toUpperCase();
  const departments = String(summary && summary.departments ? summary.departments : STATE.activeDept || '').trim().toUpperCase();
  const areasText = Array.isArray(summary && summary.areas) ? summary.areas.join(' ').toUpperCase() : '';
  const combined = `${name} ${areasText}`;

  // MACHINE TAB RULE:
  // Machine should be machine/operator placeholder rows only.
  // Do NOT send real associates or PROD temp names to Machine just because they touched Surface.
  const machineNamePatterns = [
    'UNKNOWN',
    'UNKNOWN OPERATOR',
    'MEI',
    '54R',
    'FLEX',
    'ODT',
    'OTL',
    'ORB',
    'OTB',
    'AR41',
    'IQ-STAR',
    'MANUAL BLK',
    'MANUAL BLOCK'
  ];

  if (machineNamePatterns.some(pattern => name.includes(pattern))) {
    return 'MACHINE';
  }

  // Some machine rows may have the machine pattern in the station/area text.
  // Only use this fallback when it is not a normal associate/temp name.
  const looksLikeTempOrPerson =
    name.startsWith('PROD') ||
    name.startsWith('OHIO') ||
    name.includes(',') ||
    /^[A-Z]+\s+[A-Z]+/.test(name);

  if (!looksLikeTempOrPerson && machineNamePatterns.some(pattern => combined.includes(pattern))) {
    return 'MACHINE';
  }

  const isProductionDept =
    departments.includes('FINISH') ||
    departments.includes('SURFACE') ||
    departments.includes('AR') ||
    String(STATE.activeDept || '').toUpperCase() === 'PROD_ALL';

  const isDistributionInventoryDept =
    departments.includes('INVENTORY') ||
    String(STATE.activeDept || '').toUpperCase() === 'INVENTORY' ||
    String(STATE.activeDept || '').toUpperCase() === 'DIST_ALL';

  // Production temporary labor naming rule:
  // Any production associate that starts with PROD is TEMP.
  if (isProductionDept && name.startsWith('PROD')) return 'TEMP';

  // Distribution & Inventory temporary labor naming rule:
  // Any picking/inventory associate that starts with OHIO is TEMP.
  if (isDistributionInventoryDept && name.startsWith('OHIO')) return 'TEMP';

  // Safety fallback in case the department is mixed or the active filter is ALL.
  if (name.startsWith('PROD') || name.startsWith('OHIO')) return 'TEMP';

  return 'FTE';
}

function getEmployeeTypeLabel(type) {
  const t = String(type || '').toUpperCase();
  if (t === 'TEMP') return 'Temps';
  if (t === 'MACHINE') return 'Machine';
  return 'Full Time';
}

function setScorecardEmployeeType(type) {
  const t = String(type || 'FTE').toUpperCase();
  STATE.scorecardEmployeeType = ['FTE', 'TEMP', 'MACHINE'].includes(t) ? t : 'FTE';
  STATE.selectedAssociate = '';
  renderEverything();

  const stage = document.getElementById('pageContent');
  if (stage) stage.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderScorecardsPage() {
  const page = document.getElementById('pageContent');
  if (!page) return;

  if (STATE.activeDept === 'SUMMARY') {
    renderSummaryPage();
    return;
  }

  const summaries = buildAssociateSummaries();

  if (!summaries.length) {
    page.innerHTML = `
      <section class="scoreboard-panel command-card">
        <div class="section-head compact">
          <div>
            <h3>Associate Scorecards</h3>
            <p>No associates found for this filter.</p>
          </div>
        </div>
      </section>
    `;
    return;
  }

  const sorted = summaries
    .slice()
    .sort((a, b) => {
      const statusRank = { GREEN: 1, AMBER: 2, RED: 3, 'NO TARGET': 4 };
      const sr = (statusRank[a.status] || 9) - (statusRank[b.status] || 9);
      if (sr !== 0) return sr;
      return Number(b.productivityPercent || 0) - Number(a.productivityPercent || 0);
    });

  const fte = sorted.filter(s => getAssociateEmployeeType(s) === 'FTE');
  const temps = sorted.filter(s => getAssociateEmployeeType(s) === 'TEMP');
  const machines = sorted.filter(s => getAssociateEmployeeType(s) === 'MACHINE');

  const activeTypeRaw = String(STATE.scorecardEmployeeType || 'FTE').toUpperCase();
  const activeType = ['FTE', 'TEMP', 'MACHINE'].includes(activeTypeRaw) ? activeTypeRaw : 'FTE';
  const activeRows = activeType === 'TEMP' ? temps : activeType === 'MACHINE' ? machines : fte;

  const fteJobs = fte.reduce((sum, s) => sum + Number(s.totalJobs || 0), 0);
  const tempJobs = temps.reduce((sum, s) => sum + Number(s.totalJobs || 0), 0);
  const machineJobs = machines.reduce((sum, s) => sum + Number(s.totalJobs || 0), 0);

  page.innerHTML = `
    <section class="scoreboard-panel command-card full-scoreboard-page">
      <div class="section-head compact">
        <div>
          <h3>Associate Scorecards</h3>
          <p>Separate Full Time, Temps, and Machine rows. Production temps start with PROD. Picking temps start with OHIO. Machine includes MEI, Unknown Operator, and machine/operator placeholder rows like 54R, FLEX, ODT, OTL, ORB, and OTB.</p>
        </div>
        <span class="scoreboard-count">${activeRows.length} ${escapeHtml(getEmployeeTypeLabel(activeType))}</span>
      </div>

      <div class="employee-type-kpis employee-type-kpis-three">
        <button class="employee-type-card ${activeType === 'FTE' ? 'active' : ''}" type="button" onclick="setScorecardEmployeeType('FTE')">
          <span>Full Time</span>
          <strong>${formatNumber(fte.length)}</strong>
          <small>${formatNumber(fteJobs)} jobs</small>
        </button>

        <button class="employee-type-card temp ${activeType === 'TEMP' ? 'active' : ''}" type="button" onclick="setScorecardEmployeeType('TEMP')">
          <span>Temps</span>
          <strong>${formatNumber(temps.length)}</strong>
          <small>${formatNumber(tempJobs)} jobs</small>
        </button>

        <button class="employee-type-card machine ${activeType === 'MACHINE' ? 'active' : ''}" type="button" onclick="setScorecardEmployeeType('MACHINE')">
          <span>Machine</span>
          <strong>${formatNumber(machines.length)}</strong>
          <small>${formatNumber(machineJobs)} jobs</small>
        </button>
      </div>

      <div class="scoreboard-toolbar">
        <div class="scoreboard-legend"><span class="legend-dot green"></span>On Target</div>
        <div class="scoreboard-legend"><span class="legend-dot amber"></span>Watch</div>
        <div class="scoreboard-legend"><span class="legend-dot red"></span>Warning</div>
        <div class="scoreboard-legend"><span class="legend-dot gray"></span>No Target</div>
      </div>

      ${activeRows.length ? `
        <div class="scoreboard-grid scoreboard-grid-expanded">
          ${activeRows.map(s => {
            const active = s.associate === STATE.selectedAssociate ? 'active' : '';
            const pct = s.workProduced > 0 ? formatPercent(s.productivityPercent) : 'N/A';
            const employeeType = getAssociateEmployeeType(s);
            return `
              <button class="score-kpi ${active} ${statusCss(s.status)} ${employeeType.toLowerCase()}-associate" type="button" onclick="selectAssociateFromScoreboard('${escapeJs(s.associate)}')">
                <span class="score-status-dot"></span>
                <span class="score-name">${escapeHtml(s.associate)}</span>
                <span class="score-dept">${escapeHtml(s.departments || 'No department')} • ${escapeHtml(getEmployeeTypeLabel(employeeType))}</span>
                <strong>${pct}</strong>
                <span class="score-meta">${formatDecimal(s.totalHours)} hrs • ${formatNumber(s.totalJobs)} jobs • ${formatDecimal(s.workProduced)} produced</span>
              </button>
            `;
          }).join('')}
        </div>
      ` : `
        <div class="empty-hub-note">
          No ${escapeHtml(getEmployeeTypeLabel(activeType))} rows found for this department, shift, week, station, or search filter.
        </div>
      `}
    </section>
  `;
}

function selectAssociateFromScoreboard(associate) {
  STATE.selectedAssociate = associate || '';
  STATE.activePage = 'profile';
  setActiveSideNav('scorecards');
  const select = document.getElementById('associateSelect');
  if (select) select.value = STATE.selectedAssociate;
  renderEverything();
  const stage = document.getElementById('pageContent');
  if (stage) stage.scrollIntoView({ behavior: 'smooth', block: 'start' });
}


function getInsightRows() {
  return getWeekRowsForActiveDept();
}

function summarizeInsightRows(rows, groupGetter) {
  const map = new Map();

  (rows || []).forEach(row => {
    const key = String(groupGetter(row) || 'Unmapped').trim() || 'Unmapped';
    if (!map.has(key)) {
      map.set(key, {
        label: key,
        associates: new Set(),
        hours: 0,
        jobs: 0,
        workProduced: 0,
        green: 0,
        amber: 0,
        red: 0,
        noTarget: 0
      });
    }

    const item = map.get(key);
    const name = String(row.OperatorName || '').trim();
    if (name && !isMachineRow(row)) item.associates.add(name);
    item.hours += Number(row.SessionTimeHours || 0);
    item.jobs += Number(row.TotalJobScan || 0);
    item.workProduced += Number(row.WorkProduced || 0);
  });

  return Array.from(map.values()).map(item => {
    item.associateCount = item.associates.size;
    item.productivityPercent = item.hours > 0 && item.workProduced > 0 ? item.workProduced / item.hours : 0;
    item.status = getStatusFromPercent(item.productivityPercent, item.workProduced);
    return item;
  }).sort((a, b) => b.jobs - a.jobs);
}

function getInsightTotals(rows, summaries) {
  const safeRows = rows || [];
  const safeSummaries = summaries || [];
  const hours = safeRows.reduce((sum, r) => sum + Number(r.SessionTimeHours || 0), 0);
  const jobs = safeRows.reduce((sum, r) => sum + Number(r.TotalJobScan || 0), 0);
  const workProduced = safeRows.reduce((sum, r) => sum + Number(r.WorkProduced || 0), 0);
  const productivityPercent = hours > 0 && workProduced > 0 ? workProduced / hours : 0;

  return {
    associates: safeSummaries.length,
    hours,
    jobs,
    workProduced,
    productivityPercent,
    green: safeSummaries.filter(s => s.status === 'GREEN').length,
    amber: safeSummaries.filter(s => s.status === 'AMBER').length,
    red: safeSummaries.filter(s => s.status === 'RED').length,
    noTarget: safeSummaries.filter(s => s.status === 'NO_TARGET').length
  };
}

function renderDailySummaryHub() {
  const page = document.getElementById('pageContent');
  if (!page) return;

  const rows = getInsightRows();
  const associateSummaries = buildAssociateSummariesFromRows(rows);
  const totals = getInsightTotals(rows, associateSummaries);

  const dayRows = summarizeInsightRows(rows, r => formatDateKey(r.WorkDate))
    .sort((a, b) => String(a.label).localeCompare(String(b.label)));

  const deptRows = summarizeInsightRows(rows, r => getDeptLabel(r.FinalDepartment));
  const stationRows = summarizeInsightRows(rows, r => r.StationGroup || r.AccessPoint || 'Unmapped Station');

  const topAssociates = associateSummaries
    .filter(s => Number(s.workProduced || 0) > 0)
    .sort((a, b) => Number(b.workProduced || 0) - Number(a.workProduced || 0))
    .slice(0, 8);

  const riskAssociates = associateSummaries
    .filter(s => s.status === 'AMBER' || s.status === 'RED')
    .sort((a, b) => Number(a.productivityPercent || 0) - Number(b.productivityPercent || 0))
    .slice(0, 6);

  page.innerHTML = `
    <section class="insight-shell">
      <article class="insight-hero">
        <div>
          <span class="profile-tag">Daily Summary</span>
          <h2>Summary</h2>
          <p>${escapeHtml(getDeptLabel(STATE.activeDept))} · ${escapeHtml(STATE.activeShift)} · ${escapeHtml(STATE.activeStationGroup || 'ALL')} · ${escapeHtml(STATE.selectedFiscalWeek || STATE.meta.fiscalWeek || '')}</p>
        </div>
        <div class="insight-hero-score ${totals.productivityPercent >= 1 ? 'GREEN' : totals.productivityPercent >= .9 ? 'AMBER' : 'RED'}">
          <span>Productivity</span>
          <strong>${formatPercent(totals.productivityPercent)}</strong>
        </div>
      </article>

      <div class="insight-kpi-grid">
        <div class="insight-kpi"><span>Associates</span><strong>${formatNumber(totals.associates)}</strong><small>Active people only</small></div>
        <div class="insight-kpi"><span>Total Jobs</span><strong>${formatNumber(totals.jobs)}</strong><small>Filtered week output</small></div>
        <div class="insight-kpi"><span>Total Hours</span><strong>${formatDecimal(totals.hours)}</strong><small>Login/session hours</small></div>
        <div class="insight-kpi"><span>Work Produced</span><strong>${formatDecimal(totals.workProduced)}</strong><small>Target-equivalent hours</small></div>
        <div class="insight-kpi green"><span>On Target</span><strong>${formatNumber(totals.green)}</strong><small>100%+</small></div>
        <div class="insight-kpi amber"><span>Watch</span><strong>${formatNumber(totals.amber)}</strong><small>90% - 99%</small></div>
        <div class="insight-kpi red"><span>Warning</span><strong>${formatNumber(totals.red)}</strong><small>Below 90%</small></div>
        <div class="insight-kpi"><span>No Target</span><strong>${formatNumber(totals.noTarget)}</strong><small>Missing JPH setup</small></div>
      </div>

      <section class="insight-grid two-col">
        <article class="insight-panel">
          <div class="section-head">
            <div>
              <h3>Daily Output by Date</h3>              
            </div>
          </div>
          <div class="daily-summary-list">
            ${dayRows.length ? dayRows.map(d => `
              <div class="daily-summary-row ${d.status}">
                <div>
                  <strong>${escapeHtml(formatDatePretty(d.label))}</strong>
                  <span>${formatNumber(d.associateCount)} associates · ${formatDecimal(d.hours)} hrs</span>
                </div>
                <div class="daily-summary-metrics">
                  <b>${formatNumber(d.jobs)}</b>
                  <small>jobs</small>
                </div>
                <div class="daily-summary-metrics">
                  <b>${formatPercent(d.productivityPercent)}</b>
                  <small>prod</small>
                </div>
              </div>
            `).join('') : `<div class="empty-card">No daily rows found for this filter.</div>`}
          </div>
        </article>

        <article class="insight-panel">
          <div class="section-head">
            <div>
              <h3>Top Associate Peformace</h3>              
            </div>
          </div>
          <div class="impact-list">
            ${topAssociates.length ? topAssociates.map((s, idx) => `
              <button class="impact-row ${s.status}" type="button" onclick="selectAssociateFromScoreboard('${escapeJs(s.associate)}')">
                <span class="rank-pill">#${idx + 1}</span>
                <div>
                  <strong>${escapeHtml(s.associate)}</strong>
                  <small>${escapeHtml(s.departments || getDeptLabel(STATE.activeDept))}</small>
                </div>
                <b>${formatPercent(s.productivityPercent)}</b>
              </button>
            `).join('') : `<div class="empty-card">No associate output found.</div>`}
          </div>
        </article>
      </section>

      <section class="insight-grid two-col">
        <article class="insight-panel">
          <div class="section-head">
            <div>
              <h3>Department Breakdown</h3>              
            </div>
          </div>
          ${renderInsightTable(deptRows, 'Department')}
        </article>

        <article class="insight-panel">
          <div class="section-head">
            <div>
              <h3>Station Group Breakdown</h3>              
            </div>
          </div>
          ${renderInsightTable(stationRows.slice(0, 12), 'Station')}
        </article>
      </section>

      ${riskAssociates.length ? `
        <section class="insight-panel">
          <div class="section-head">
            <div>
              <h3>Watch List</h3>
              <p>These associates need coaching, note review, or target validation.</p>
            </div>
          </div>
          <div class="watch-grid">
            ${riskAssociates.map(s => `
              <button class="watch-card ${s.status}" type="button" onclick="selectAssociateFromScoreboard('${escapeJs(s.associate)}')">
                <strong>${escapeHtml(s.associate)}</strong>
                <span>${formatPercent(s.productivityPercent)} · ${formatDecimal(s.totalHours)} hrs · ${formatNumber(s.totalJobs)} jobs</span>
              </button>
            `).join('')}
          </div>
        </section>
      ` : ''}
    </section>
  `;
}

function renderTrendsHub() {
  const page = document.getElementById('pageContent');
  if (!page) return;

  const rows = getInsightRows();
  const associateSummaries = buildAssociateSummariesFromRows(rows);
  const totals = getInsightTotals(rows, associateSummaries);

  const dayRows = summarizeInsightRows(rows, r => formatDateKey(r.WorkDate))
    .sort((a, b) => String(a.label).localeCompare(String(b.label)));

  const maxDayJobs = Math.max(1, ...dayRows.map(d => Number(d.jobs || 0)));
  const bestDay = dayRows.slice().sort((a, b) => b.productivityPercent - a.productivityPercent)[0];
  const weakDay = dayRows.slice().filter(d => d.workProduced > 0).sort((a, b) => a.productivityPercent - b.productivityPercent)[0];

  const stationRows = summarizeInsightRows(rows, r => r.StationGroup || r.AccessPoint || 'Unmapped Station');
  const strongestStations = stationRows
    .filter(s => s.workProduced > 0)
    .sort((a, b) => b.productivityPercent - a.productivityPercent)
    .slice(0, 6);
  const weakestStations = stationRows
    .filter(s => s.workProduced > 0)
    .sort((a, b) => a.productivityPercent - b.productivityPercent)
    .slice(0, 6);

  const topAssociates = associateSummaries
    .filter(s => Number(s.workProduced || 0) > 0)
    .sort((a, b) => Number(b.productivityPercent || 0) - Number(a.productivityPercent || 0))
    .slice(0, 8);

  const watchAssociates = associateSummaries
    .filter(s => s.status === 'AMBER' || s.status === 'RED')
    .sort((a, b) => Number(a.productivityPercent || 0) - Number(b.productivityPercent || 0))
    .slice(0, 8);

  page.innerHTML = `
    <section class="insight-shell">
      <article class="insight-hero trend">
        <div>
          <span class="profile-tag">Weekly Trends</span>
          <h2>Weekly Performance </h2>
          <p>${escapeHtml(getDeptLabel(STATE.activeDept))} · ${escapeHtml(STATE.activeShift)} · ${escapeHtml(STATE.selectedFiscalWeek || STATE.meta.fiscalWeek || '')}</p>
        </div>
        <div class="insight-hero-score ${totals.productivityPercent >= 1 ? 'GREEN' : totals.productivityPercent >= .9 ? 'AMBER' : 'RED'}">
          <span>Week Avg</span>
          <strong>${formatPercent(totals.productivityPercent)}</strong>
        </div>
      </article>

      <div class="trend-callout-grid">
        <div class="trend-callout">
          <span>Best Day</span>
          <strong>${bestDay ? escapeHtml(formatDatePretty(bestDay.label)) : '--'}</strong>
          <small>${bestDay ? formatPercent(bestDay.productivityPercent) : 'No data'}</small>
        </div>
        <div class="trend-callout">
          <span>Weak Day</span>
          <strong>${weakDay ? escapeHtml(formatDatePretty(weakDay.label)) : '--'}</strong>
          <small>${weakDay ? formatPercent(weakDay.productivityPercent) : 'No data'}</small>
        </div>
        <div class="trend-callout">
          <span>Strongest Station</span>
          <strong>${strongestStations[0] ? escapeHtml(strongestStations[0].label) : '--'}</strong>
          <small>${strongestStations[0] ? formatPercent(strongestStations[0].productivityPercent) : 'No data'}</small>
        </div>
        <div class="trend-callout danger">
          <span>Underperformance Associates</span>
          <strong>${formatNumber(watchAssociates.length)}</strong>
          <small>Amber / red only</small>
        </div>
      </div>

      <section class="insight-panel">
        <div class="section-head">
          <div>
            <h3>Daily Trend</h3>            
          </div>
        </div>
        <div class="trend-bars">
          ${dayRows.length ? dayRows.map(d => {
            const width = Math.max(4, Math.round((Number(d.jobs || 0) / maxDayJobs) * 100));
            return `
              <div class="trend-bar-row ${d.status}">
                <div class="trend-bar-label">
                  <strong>${escapeHtml(formatDatePretty(d.label))}</strong>
                  <span>${formatNumber(d.jobs)} jobs · ${formatDecimal(d.hours)} hrs</span>
                </div>
                <div class="trend-track">
                  <div class="trend-fill" style="width:${width}%"></div>
                </div>
                <b>${formatPercent(d.productivityPercent)}</b>
              </div>
            `;
          }).join('') : `<div class="empty-card">No daily trend rows found.</div>`}
        </div>
      </section>

      <section class="insight-grid two-col">
        <article class="insight-panel">
          <div class="section-head">
            <div>
              <h3>Strongest Stations</h3>              
            </div>
          </div>
          <div class="rank-list">
            ${strongestStations.length ? strongestStations.map((s, idx) => renderRankRow(s, idx)).join('') : `<div class="empty-card">No station productivity data.</div>`}
          </div>
        </article>

        <article class="insight-panel">
          <div class="section-head">
            <div>
              <h3>Low Stations</h3>              
            </div>
          </div>
          <div class="rank-list">
            ${weakestStations.length ? weakestStations.map((s, idx) => renderRankRow(s, idx)).join('') : `<div class="empty-card">No station risk data.</div>`}
          </div>
        </article>
      </section>

      <section class="insight-grid two-col">
        <article class="insight-panel">
          <div class="section-head">
            <div>
              <h3>Top Associate </h3>              
            </div>
          </div>
          <div class="impact-list">
            ${topAssociates.length ? topAssociates.map((s, idx) => `
              <button class="impact-row ${s.status}" type="button" onclick="selectAssociateFromScoreboard('${escapeJs(s.associate)}')">
                <span class="rank-pill">#${idx + 1}</span>
                <div>
                  <strong>${escapeHtml(s.associate)}</strong>
                  <small>${formatDecimal(s.totalHours)} hrs · ${formatNumber(s.totalJobs)} jobs</small>
                </div>
                <b>${formatPercent(s.productivityPercent)}</b>
              </button>
            `).join('') : `<div class="empty-card">No associate trend data.</div>`}
          </div>
        </article>

        <article class="insight-panel">
          <div class="section-head">
            <div>
              <h3>Associate Watch</h3>              
            </div>
          </div>
          <div class="impact-list">
            ${watchAssociates.length ? watchAssociates.map((s, idx) => `
              <button class="impact-row ${s.status}" type="button" onclick="selectAssociateFromScoreboard('${escapeJs(s.associate)}')">
                <span class="rank-pill">#${idx + 1}</span>
                <div>
                  <strong>${escapeHtml(s.associate)}</strong>
                  <small>${formatDecimal(s.totalHours)} hrs · ${formatNumber(s.totalJobs)} jobs</small>
                </div>
                <b>${formatPercent(s.productivityPercent)}</b>
              </button>
            `).join('') : `<div class="empty-card">No amber/red associates for this filter.</div>`}
          </div>
        </article>
      </section>
    </section>
  `;
}

function renderInsightTable(rows, labelName) {
  if (!rows || !rows.length) {
    return `<div class="empty-card">No rows found.</div>`;
  }

  return `
    <div class="insight-table-wrap">
      <table class="insight-table">
        <thead>
          <tr>
            <th>${escapeHtml(labelName || 'Group')}</th>
            <th>Associates</th>
            <th>Hours</th>
            <th>Jobs</th>
            <th>Work Produced</th>
            <th>Prod %</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(r => `
            <tr class="${r.status}">
              <td><strong>${escapeHtml(r.label)}</strong></td>
              <td>${formatNumber(r.associateCount || 0)}</td>
              <td>${formatDecimal(r.hours || 0)}</td>
              <td>${formatNumber(r.jobs || 0)}</td>
              <td>${formatDecimal(r.workProduced || 0)}</td>
              <td><b>${formatPercent(r.productivityPercent || 0)}</b></td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderRankRow(item, idx) {
  return `
    <div class="rank-row ${item.status}">
      <span class="rank-pill">#${idx + 1}</span>
      <div>
        <strong>${escapeHtml(item.label)}</strong>
        <small>${formatNumber(item.jobs)} jobs · ${formatDecimal(item.hours)} hrs · ${formatNumber(item.associateCount || 0)} associates</small>
      </div>
      <b>${formatPercent(item.productivityPercent)}</b>
    </div>
  `;
}

function escapeJs(value) {
  return String(value || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/\n/g, ' ');
}

function renderTargetsHub() {
  renderPlaceholderPage('Targets', 'Target JPH review and missing-target alerts will go here.');
}

function renderPlaceholderPage(title, message) {
  const page = document.getElementById('pageContent');
  if (!page) return;
  const niceTitle = String(title || 'Page').replace(/^./, c => c.toUpperCase());
  page.innerHTML = `
    <section class="hub-panel command-card placeholder-panel">
      <span class="profile-tag">Productivity Hub</span>
      <h2>${escapeHtml(niceTitle)}</h2>
      <p>${escapeHtml(message || 'This section is available for the next build phase.')}</p>
      <button class="primary-action" type="button" onclick="setActivePage('dashboard')">Back to Dashboard Hub</button>
    </section>
  `;
}


/************************************************************
 * BREAKAGE / QUALITY HUB
 ************************************************************/

function getBreakageRowsForSelectedWeek() {
  const startDate = STATE.selectedWeekStartDate || STATE.meta.weekStartDate;
  const endDate = STATE.selectedWeekEndDate || STATE.meta.weekEndDate;
  const start = parseLocalDate(startDate);
  const end = parseLocalDate(endDate);

  const selectedStartKey = formatDateKey(startDate);
  const selectedEndKey = formatDateKey(endDate);
  const currentStartKey = formatDateKey(STATE.meta.weekStartDate);
  const currentEndKey = formatDateKey(STATE.meta.weekEndDate);
  const useCurrentMaster = selectedStartKey === currentStartKey && selectedEndKey === currentEndKey;

  const rows = [
    ...(STATE.breakageDailySnapshot || []),
    ...(useCurrentMaster ? (STATE.breakageMaster || []) : [])
  ];

  return dedupeBreakageRows(rows).filter(r => {
    const d = parseLocalDate(r.WorkDate);
    const dateMatch = d && start && end && d >= start && d <= end;
    if (!dateMatch) return false;

    const activeAP = String(STATE.activeStationGroup || 'ALL').trim();
    const apMatch = activeAP === 'ALL' || normalizeQualityText(r.AccessPoint) === normalizeQualityText(activeAP);

    const blob = [r.OperatorName, r.AccessPoint, r.BreakageReason].join(' ').toLowerCase();
    const searchMatch = !STATE.search || blob.includes(STATE.search);

    return apMatch && searchMatch;
  });
}

function dedupeBreakageRows(rows) {
  const map = new Map();

  (rows || []).forEach(row => {
    const key = [
      formatDateKey(row.WorkDate),
      row.AccessPoint,
      row.BreakageReason,
      row.OperatorName,
      row.LensBreakageCount,
      row.FrameBreakageCount,
      row.TotalBreakageCount
    ].join('|');

    map.set(key, row);
  });

  return Array.from(map.values());
}

function normalizeQualityText(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

function getQualityStatus(rowOrPercent, framePercentMaybe, scansMaybe) {
  let lensPercent = 0;
  let framePercent = 0;
  let scans = 0;

  if (typeof rowOrPercent === 'object') {
    lensPercent = Number(rowOrPercent.OperatorLensBreakagePercent || rowOrPercent.DailyLensBreakagePercent || 0);
    framePercent = Number(rowOrPercent.OperatorFrameBreakagePercent || rowOrPercent.DailyFrameBreakagePercent || 0);
    scans = Number(rowOrPercent.OperatorDailyScans || rowOrPercent.DailyTotalScans || 0);
    if (scans <= 0 && Number(rowOrPercent.TotalBreakageCount || 0) > 0) return 'NO SCAN MATCH';
  } else {
    lensPercent = Number(rowOrPercent || 0);
    framePercent = Number(framePercentMaybe || 0);
    scans = Number(scansMaybe || 0);
    if (scans <= 0) return 'NO SCAN MATCH';
  }

  const worst = Math.max(lensPercent, framePercent);

  if (worst <= 0.005) return 'GREEN';
  if (worst <= 0.01) return 'AMBER';
  return 'RED';
}

function formatQualityPercent(value) {
  const n = Number(value || 0);
  return `${(n * 100).toFixed(2)}%`;
}

function buildBreakageSummary(rows) {
  const safeRows = rows || [];
  const lens = safeRows.reduce((sum, r) => sum + Number(r.LensBreakageCount || 0), 0);
  const frame = safeRows.reduce((sum, r) => sum + Number(r.FrameBreakageCount || 0), 0);
  const total = safeRows.reduce((sum, r) => sum + Number(r.TotalBreakageCount || 0), 0);

  const scanMap = new Map();
  safeRows.forEach(r => {
    const key = [formatDateKey(r.WorkDate), normalizeQualityText(r.OperatorName)].join('|');
    const scans = Number(r.OperatorDailyScans || r.DailyTotalScans || 0);
    if (!key || scans <= 0) return;
    scanMap.set(key, Math.max(scanMap.get(key) || 0, scans));
  });

  const scans = Array.from(scanMap.values()).reduce((sum, v) => sum + Number(v || 0), 0);
  const lensOpportunity = scans * 2;
  const lensPercent = lensOpportunity > 0 ? lens / lensOpportunity : 0;
  const framePercent = scans > 0 ? frame / scans : 0;

  return {
    rows: safeRows,
    lens,
    frame,
    total,
    scans,
    lensPercent,
    framePercent,
    status: total > 0 ? getQualityStatus(lensPercent, framePercent, scans) : 'GREEN',
    topReason: getTopBreakageDimension(safeRows, 'BreakageReason'),
    topAccessPoint: getTopBreakageDimension(safeRows, 'AccessPoint'),
    topOperator: getTopBreakageDimension(safeRows, 'OperatorName')
  };
}

function getTopBreakageDimension(rows, field) {
  const map = new Map();

  (rows || []).forEach(r => {
    const key = String(r[field] || '').trim() || 'Unknown';
    map.set(key, (map.get(key) || 0) + Number(r.TotalBreakageCount || 0));
  });

  const sorted = Array.from(map.entries()).sort((a, b) => b[1] - a[1]);
  return sorted.length ? { name: sorted[0][0], total: sorted[0][1] } : { name: 'None', total: 0 };
}

function buildBreakageDimensionRows(rows, field, limit = 10) {
  const map = new Map();

  (rows || []).forEach(r => {
    const key = String(r[field] || '').trim() || 'Unknown';
    if (!map.has(key)) {
      map.set(key, { name: key, lens: 0, frame: 0, total: 0, scans: 0, scanKeys: new Map() });
    }

    const item = map.get(key);
    item.lens += Number(r.LensBreakageCount || 0);
    item.frame += Number(r.FrameBreakageCount || 0);
    item.total += Number(r.TotalBreakageCount || 0);

    const scanKey = [formatDateKey(r.WorkDate), normalizeQualityText(r.OperatorName)].join('|');
    const scans = Number(r.OperatorDailyScans || r.DailyTotalScans || 0);
    if (scans > 0) item.scanKeys.set(scanKey, Math.max(item.scanKeys.get(scanKey) || 0, scans));
  });

  return Array.from(map.values()).map(item => {
    item.scans = Array.from(item.scanKeys.values()).reduce((sum, v) => sum + v, 0);
    item.lensPercent = item.scans > 0 ? item.lens / item.scans : 0;
    item.framePercent = item.scans > 0 ? item.frame / item.scans : 0;
    item.status = item.total > 0 ? getQualityStatus(item.lensPercent, item.framePercent, item.scans) : 'GREEN';
    delete item.scanKeys;
    return item;
  }).sort((a, b) => b.total - a.total).slice(0, limit);
}


function setBreakageDateFilter(dateKey) {
  STATE.selectedBreakageDate = String(dateKey || 'ALL');
  renderEverything();

  const panel = document.getElementById('breakageRecordsPanel');
  if (panel) panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function buildBreakageDateCards(rows) {
  const map = new Map();

  (rows || []).forEach(r => {
    const key = formatDateKey(r.WorkDate);
    if (!key) return;

    if (!map.has(key)) {
      map.set(key, {
        key,
        workDate: r.WorkDate,
        lens: 0,
        frame: 0,
        total: 0,
        records: 0
      });
    }

    const item = map.get(key);
    item.lens += Number(r.LensBreakageCount || 0);
    item.frame += Number(r.FrameBreakageCount || 0);
    item.total += Number(r.TotalBreakageCount || 0);
    item.records += 1;
  });

  return Array.from(map.values()).sort((a, b) => {
    return parseLocalDate(a.workDate) - parseLocalDate(b.workDate);
  });
}

function getSelectedBreakageRows(rows) {
  const selected = String(STATE.selectedBreakageDate || 'ALL');
  if (selected === 'ALL') return rows || [];

  return (rows || []).filter(r => formatDateKey(r.WorkDate) === selected);
}

function renderBreakageDateCards(rows) {
  const cards = buildBreakageDateCards(rows);
  const selected = String(STATE.selectedBreakageDate || 'ALL');

  const weekLens = (rows || []).reduce((sum, r) => sum + Number(r.LensBreakageCount || 0), 0);
  const weekFrame = (rows || []).reduce((sum, r) => sum + Number(r.FrameBreakageCount || 0), 0);
  const weekTotal = (rows || []).reduce((sum, r) => sum + Number(r.TotalBreakageCount || 0), 0);

  return `
    <section class="quality-panel command-card breakage-date-card-panel">
      <div class="section-head compact">
        <div>
          <h3>Breakage by Date</h3>
          <p>Select one date to review records. This keeps the page clean instead of showing the full week at once.</p>
        </div>
        <span class="scoreboard-count">${cards.length} day${cards.length === 1 ? '' : 's'}</span>
      </div>

      <div class="breakage-date-card-grid">
        <button class="breakage-date-card ${selected === 'ALL' ? 'active' : ''}" type="button" onclick="setBreakageDateFilter('ALL')">
          <span>Full Week</span>
          <strong>${formatNumber(weekTotal)}</strong>
          <small>${formatNumber(weekLens)} lens • ${formatNumber(weekFrame)} frame</small>
        </button>

        ${cards.map(card => `
          <button class="breakage-date-card ${selected === card.key ? 'active' : ''}" type="button" onclick="setBreakageDateFilter('${escapeJs(card.key)}')">
            <span>${formatDatePretty(card.workDate)}</span>
            <strong>${formatNumber(card.total)}</strong>
            <small>${formatNumber(card.lens)} lens • ${formatNumber(card.frame)} frame</small>
          </button>
        `).join('')}
      </div>
    </section>
  `;
}


function renderBreakageHub() {
  const page = document.getElementById('pageContent');
  if (!page) return;

  const rows = getBreakageRowsForSelectedWeek();
  const summary = buildBreakageSummary(rows);

  const selectedRows = getSelectedBreakageRows(rows);
  const selectedSummary = buildBreakageSummary(selectedRows);

  const byAccessPoint = buildBreakageDimensionRows(selectedRows, 'AccessPoint', 8);
  const byReason = buildBreakageDimensionRows(selectedRows, 'BreakageReason', 8);
  const byOperator = buildBreakageDimensionRows(selectedRows, 'OperatorName', 8);

  page.innerHTML = `
    <section class="quality-hub-shell clean-quality-hub">
      <div class="quality-hero command-card">
        <div>
          <span class="profile-tag">Quality / Breakage Hub</span>
          <h2>Breakage Control Center</h2>
          <p>Weekly breakage count view. Percent, scan count, and status were removed to keep this clean.</p>
        </div>
        <div class="quality-hero-metrics clean-breakage-totals">
          <div><span>Total Breakage</span><strong>${formatNumber(summary.total)}</strong></div>
          <div><span>Lens</span><strong>${formatNumber(summary.lens)}</strong></div>
          <div><span>Frame</span><strong>${formatNumber(summary.frame)}</strong></div>
        </div>
      </div>

      <div class="quality-kpi-grid clean-quality-kpis">
        <article class="quality-kpi"><span>Selected Total</span><strong>${formatNumber(selectedSummary.total)}</strong><small>${selectedRows.length} records</small></article>
        <article class="quality-kpi"><span>Selected Lens</span><strong>${formatNumber(selectedSummary.lens)}</strong><small>${escapeHtml(selectedSummary.topReason.name)}</small></article>
        <article class="quality-kpi"><span>Selected Frame</span><strong>${formatNumber(selectedSummary.frame)}</strong><small>${escapeHtml(selectedSummary.topAccessPoint.name)}</small></article>
        <article class="quality-kpi"><span>Top Operator</span><strong>${escapeHtml(selectedSummary.topOperator.name)}</strong><small>${formatNumber(selectedSummary.topOperator.total)} breaks</small></article>
      </div>

      ${renderBreakageDateCards(rows)}

      <div class="quality-grid-three">
        ${renderQualityListPanel('By Access Point', byAccessPoint)}
        ${renderQualityListPanel('By Reason', byReason)}
        ${renderQualityListPanel('By Operator', byOperator)}
      </div>

      ${renderBreakageRecordsTable(selectedRows)}
    </section>
  `;
}

function renderQualityListPanel(title, rows) {
  return `
    <section class="quality-panel command-card">
      <div class="section-head compact">
        <div>
          <h3>${escapeHtml(title)}</h3>
          <p>Ranked by total breakage count.</p>
        </div>
      </div>
      <div class="quality-list clean-quality-list">
        ${rows.length ? rows.map(r => `
          <div class="quality-row clean-quality-row">
            <div>
              <strong>${escapeHtml(r.name)}</strong>
              <span>${formatNumber(r.total)} total • ${formatNumber(r.lens)} lens • ${formatNumber(r.frame)} frame</span>
            </div>
            <div>
              <b>${formatNumber(r.total)}</b>
              <small>Total</small>
            </div>
          </div>
        `).join('') : '<div class="empty-hub-note">No breakage records for this filter.</div>'}
      </div>
    </section>
  `;
}

function renderBreakageRecordsTable(rows) {
  const selected = String(STATE.selectedBreakageDate || 'ALL');

  const sorted = rows.slice().sort((a, b) => {
    const d = parseLocalDate(a.WorkDate) - parseLocalDate(b.WorkDate);
    if (d !== 0) return d;
    return Number(b.TotalBreakageCount || 0) - Number(a.TotalBreakageCount || 0);
  });

  if (selected === 'ALL') {
    const dailyCards = buildBreakageDateCards(rows);

    return `
      <section class="quality-panel command-card organized-breakage-records" id="breakageRecordsPanel">
        <div class="section-head compact">
          <div>
            <h3>Breakage Daily Summary</h3>
            <p>Full week view is summarized by date. Click a date card above to open the detailed records.</p>
          </div>
          <span class="scoreboard-count">${dailyCards.length} day${dailyCards.length === 1 ? '' : 's'}</span>
        </div>

        <div class="breakage-daily-summary-grid">
          ${dailyCards.length ? dailyCards.map(card => `
            <button class="breakage-daily-summary-card" type="button" onclick="setBreakageDateFilter('${escapeJs(card.key)}')">
              <div>
                <span>${formatDatePretty(card.workDate)}</span>
                <strong>${formatNumber(card.total)} total</strong>
              </div>
              <div class="daily-breakage-mini">
                <b>${formatNumber(card.lens)}</b><small>Lens</small>
                <b>${formatNumber(card.frame)}</b><small>Frame</small>
                <b>${formatNumber(card.records)}</b><small>Records</small>
              </div>
            </button>
          `).join('') : '<div class="empty-hub-note">No breakage records found for this week.</div>'}
        </div>
      </section>
    `;
  }

  const title = sorted.length
    ? `Breakage Records - ${formatDatePretty(sorted[0].WorkDate)}`
    : 'Breakage Records';

  const totalLens = sorted.reduce((sum, r) => sum + Number(r.LensBreakageCount || 0), 0);
  const totalFrame = sorted.reduce((sum, r) => sum + Number(r.FrameBreakageCount || 0), 0);
  const totalBreakage = sorted.reduce((sum, r) => sum + Number(r.TotalBreakageCount || 0), 0);

  return `
    <section class="quality-panel command-card organized-breakage-records" id="breakageRecordsPanel">
      <div class="section-head compact">
        <div>
          <h3>${escapeHtml(title)}</h3>
          <p>Clean detail view for the selected date.</p>
        </div>
        <div class="record-summary-pills">
          <span>${formatNumber(totalBreakage)} total</span>
          <span>${formatNumber(totalLens)} lens</span>
          <span>${formatNumber(totalFrame)} frame</span>
          <span>${sorted.length} records</span>
        </div>
      </div>

      ${sorted.length ? `
        <div class="breakage-record-card-grid">
          ${sorted.map(r => `
            <article class="breakage-record-card">
              <div class="record-total-badge">
                <strong>${formatNumber(r.TotalBreakageCount)}</strong>
                <span>Total</span>
              </div>

              <div class="record-main">
                <div class="record-title-row">
                  <strong>${escapeHtml(r.AccessPoint)}</strong>
                  <span>${formatDatePretty(r.WorkDate)}</span>
                </div>

                <div class="record-reason">${escapeHtml(r.BreakageReason)}</div>
                <div class="record-operator">${escapeHtml(r.OperatorName)}</div>

                <div class="record-count-row">
                  <span><b>${formatNumber(r.LensBreakageCount)}</b> Lens</span>
                  <span><b>${formatNumber(r.FrameBreakageCount)}</b> Frame</span>
                </div>
              </div>
            </article>
          `).join('')}
        </div>
      ` : `
        <div class="empty-hub-note">No breakage records found for this date.</div>
      `}
    </section>
  `;
}

function buildAssociateBreakageSummary(summary, rows) {
  const productivityRows = Array.isArray(summary.rows) ? summary.rows : [];

  // The weekly denominator must come from Activity/Productivity, not breakage rows.
  // This includes days where the associate worked but had zero breakage.
  const scanByDate = new Map();
  productivityRows.forEach(row => {
    const dateKey = formatDateKey(row.WorkDate);
    if (!dateKey) return;
    const scans = Number(row.TotalJobScan || 0);
    scanByDate.set(dateKey, (scanByDate.get(dateKey) || 0) + scans);
  });

  const scans = Array.from(scanByDate.values())
    .reduce((sum, value) => sum + Number(value || 0), 0);

  const lens = (rows || []).reduce((sum, r) => sum + Number(r.LensBreakageCount || 0), 0);
  const frame = (rows || []).reduce((sum, r) => sum + Number(r.FrameBreakageCount || 0), 0);
  const total = (rows || []).reduce((sum, r) => sum + Number(r.TotalBreakageCount || 0), 0);

  const lensOpportunity = scans * 2;
  const lensPercent = lensOpportunity > 0 ? lens / lensOpportunity : 0;
  const framePercent = scans > 0 ? frame / scans : 0;

  return {
    rows: rows || [],
    scans,
    lens,
    frame,
    total,
    lensPercent,
    framePercent,
    status: total > 0 ? getQualityStatus(lensPercent, framePercent, scans) : 'GREEN',
    topReason: getTopBreakageDimension(rows || [], 'BreakageReason'),
    topAccessPoint: getTopBreakageDimension(rows || [], 'AccessPoint'),
    topOperator: getTopBreakageDimension(rows || [], 'OperatorName')
  };
}

function renderAssociateBreakageSection(summary) {
  const rows = getBreakageRowsForSelectedWeek().filter(r =>
    normalizeQualityText(r.OperatorName) === normalizeQualityText(summary.associate)
  );

  const q = buildAssociateBreakageSummary(summary, rows);
  const byReason = buildBreakageDimensionRows(rows, 'BreakageReason', 6);
  const byAccessPoint = buildBreakageDimensionRows(rows, 'AccessPoint', 6);

  return `
    <section class="associate-quality-panel command-card">
      <div class="section-head">
        <div>
          <h3>Quality / Breakage This Week</h3>
          <p>Matched to this associate by operator name from Activity and Breakage.</p>
        </div>
        ${statusBadge(q.status)}
      </div>

      <div class="quality-kpi-grid compact-quality">
        <article class="quality-kpi"><span>Operator Scans</span><strong>${formatNumber(q.scans)}</strong><small>From Activity</small></article>
        <article class="quality-kpi"><span>Lens Breakage</span><strong>${formatNumber(q.lens)}</strong><small>${formatQualityPercent(q.lensPercent)}</small></article>
        <article class="quality-kpi"><span>Frame Breakage</span><strong>${formatNumber(q.frame)}</strong><small>${formatQualityPercent(q.framePercent)}</small></article>
        <article class="quality-kpi"><span>Total Breakage</span><strong>${formatNumber(q.total)}</strong><small>${rows.length} records</small></article>
      </div>

      <div class="quality-grid-two">
        ${renderQualityListPanel('Top Reasons', byReason)}
        ${renderQualityListPanel('Access Points', byAccessPoint)}
      </div>
    </section>
  `;
}

function renderProfile() {
  const page = document.getElementById('pageContent');
  if (!page) return;

  if (STATE.activeDept === 'SUMMARY') {
    renderSummaryPage();
    return;
  }

  const summaries = buildAssociateSummaries();

  if (!summaries.length) {
    page.innerHTML = `
      <div class="loading-card">
        <strong>No records found for ${escapeHtml(STATE.activeDept)}.</strong>
        <span>Check DAILY_SNAPSHOT and PRODUCTIVITY_MASTER.</span>
      </div>
    `;
    return;
  }

  let summary = summaries.find(s => s.associate === STATE.selectedAssociate);
  if (!summary) {
    summary = summaries[0];
    STATE.selectedAssociate = summary.associate;
    const select = document.getElementById('associateSelect');
    if (select) select.value = summary.associate;
  }

  page.innerHTML = `
    <div class="profile-return-row">
      <button class="ghost-action" type="button" onclick="setActivePage('scorecards')">← Back to Associate Scorecards</button>
    </div>
    <div class="profile-shell profile-shell-clean">
      <div class="profile-main">
        ${renderHero(summary)}
        <div class="profile-body">
          ${renderAreaSection(summary)}
          ${renderDaySection(summary)}
          ${renderRecordsSection(summary)}
        </div>
      </div>
    </div>
  `;

  animateCounters();
}

function renderHero(summary) {
  const pct = summary.workProduced > 0 ? Math.min(Math.round(summary.productivityPercent * 100), 180) : 0;
  const breakageRows = getBreakageRowsForSelectedWeek().filter(r =>
    normalizeQualityText(r.OperatorName) === normalizeQualityText(summary.associate)
  );
  const quality = buildAssociateBreakageSummary(summary, breakageRows);

  return `
    <section class="profile-hero command-card">
      <div class="identity">
        <div class="avatar">👤</div>
        <div>
          <span class="profile-tag">Associate Profile</span>
          <h3>${escapeHtml(summary.associate)}</h3>
          <p>${escapeHtml(summary.departments || STATE.activeDept)} • ${formatScheduleType(summary.scheduleType)} Schedule • ${summary.scheduledDaysWorked} of ${summary.eligibleWorkDays} eligible day${summary.eligibleWorkDays === 1 ? '' : 's'} worked • ${summary.excludedDays} excluded • ${summary.areas.length} area${summary.areas.length === 1 ? '' : 's'}</p>
          ${statusBadge(summary.status)}
        </div>
      </div>

      <div class="ring-wrap">
        <div class="progress-ring" style="--pct:${pct}">
          <strong>${summary.workProduced > 0 ? pct + '%' : 'N/A'}</strong>
        </div>
        <small>Weekly Productivity</small>
      </div>

      <div class="hero-metric">
        <span class="metric-icon">▣</span>
        <div>
          <span>Total Jobs</span>
          <strong data-count="${summary.totalJobs}">${formatNumber(summary.totalJobs)}</strong>
        </div>
      </div>

      <div class="hero-metric">
        <span class="metric-icon">◴</span>
        <div>
          <span>Total Hours</span>
          <strong>${formatDecimal(summary.totalHours)}</strong>
        </div>
      </div>

      <div class="hero-metric">
        <span class="metric-icon">↗</span>
        <div>
          <span>Avg JPH</span>
          <strong>${formatDecimal(summary.avgJph)}</strong>
        </div>
      </div>

      <div class="hero-metric">
        <span class="metric-icon">◎</span>
        <div>
          <span>Work Produced</span>
          <strong>${formatDecimal(summary.workProduced)}</strong>
        </div>
      </div>

      <div class="hero-metric quality-hero-card ${statusCss(quality.status)}">
        <span class="metric-icon">◈</span>
        <div>
          <span>Lens Breakage</span>
          <strong>${formatNumber(quality.lens)}</strong>
          <small>${formatQualityPercent(quality.lensPercent)}</small>
        </div>
      </div>

      <div class="hero-metric quality-hero-card ${statusCss(quality.status)}">
        <span class="metric-icon">◇</span>
        <div>
          <span>Frame Breakage</span>
          <strong>${formatNumber(quality.frame)}</strong>
          <small>${formatQualityPercent(quality.framePercent)}</small>
        </div>
      </div>
    </section>
  `;
}

function renderAreaSection(summary) {
  return `
    <section>
      <div class="section-head">
        <div>
          <h3>Areas Worked This Week</h3>          
        </div>
      </div>

      <div class="area-grid">
        ${summary.areas.map(area => `
          <article class="area-card">
            <div class="area-top">
              <div>
                <h4>${escapeHtml(area.stationGroup)}</h4>
                <p>${escapeHtml(area.department)} • ${escapeHtml(area.accessPoint)}</p>
              </div>
              ${statusBadge(area.status)}
            </div>

            <div class="area-metrics area-metrics-five">
              <div><span>Days</span><strong>${area.workDays}</strong></div>
              <div><span>Hours</span><strong>${formatDecimal(area.totalHours)}</strong></div>
              <div><span>Jobs</span><strong>${formatNumber(area.totalJobs)}</strong></div>
              <div><span>Produced</span><strong>${formatDecimal(area.workProduced)}</strong></div>
              <div><span>Area %</span><strong>${formatPercent(area.productivityPercent)}</strong></div>
            </div>
          </article>
        `).join('')}
      </div>
    </section>
  `;
}

function renderDaySection(summary) {
  const dayRows = aggregateAssociateByDay(summary.rows);

  if (!dayRows.length) return '';

  return `
    <section>
      <div class="section-head">
        <div>
          <h3>Productivity Scorecards</h3>          
        </div>
      </div>

      <div class="day-grid day-grid-working-only">
        ${dayRows.map(dayRow => renderDayCard(getDayName(dayRow.workDate), dayRow, summary, dayRow.workDate)).join('')}
      </div>
    </section>
  `;
}

function renderDayCard(day, row, summary, dateValue) {
  const dateKey = formatDateKey(dateValue);
  const associate = summary.associate || STATE.selectedAssociate || '';

  if (!row) {
    return `
      <article class="day-card NO-TARGET day-card-clickable" onclick="openAssociateDayAdjustment('${escapeJs(dateKey)}','${escapeJs(associate)}')" title="Open this date">
        <div class="day-card-action-hint">Open</div>
        <h4>${day}</h4>
        <p>${formatDatePretty(dateValue)} • No record</p>
        <div class="day-stats day-stats-wide">
          <div><span>Hours</span><strong>0.00</strong></div>
          <div><span>Jobs</span><strong>0</strong></div>
          <div><span>Produced</span><strong>0.00</strong></div>
          <div><span>Daily %</span><strong>0%</strong></div>
          <div><span>Status</span><strong>—</strong></div>
        </div>
      </article>
    `;
  }

  const status = getStatusFromPercent(row.productivityPercent, row.workProduced);

  return `
    <article class="day-card ${statusCss(status)} day-card-clickable" onclick="openAssociateDayAdjustment('${escapeJs(dateKey)}','${escapeJs(associate)}')" title="Click to edit this date">
      <div class="day-card-action-hint">Edit Time</div>
      <h4>${day}</h4>
      <p>${formatDatePretty(row.workDate)}</p>
      <div class="day-stats day-stats-wide">
        <div><span>Hours</span><strong>${formatDecimal(row.totalHours)}</strong></div>
        <div><span>Jobs</span><strong>${formatNumber(row.totalJobs)}</strong></div>
        <div><span>Produced</span><strong>${formatDecimal(row.workProduced)}</strong></div>
        <div><span>Daily %</span><strong>${formatPercent(row.productivityPercent)}</strong></div>
        <div><span>Status</span><strong>${statusLabel(status)}</strong></div>
      </div>
    </article>
  `;
}

function renderRecordsSection(summary) {
  const rows = summary.rows.slice().sort((a, b) => {
    const dateCompare = parseLocalDate(a.WorkDate) - parseLocalDate(b.WorkDate);
    if (dateCompare !== 0) return dateCompare;

    return String(a.StationGroup || '').localeCompare(String(b.StationGroup || ''));
  });

  const totalHours = rows.reduce((sum, r) => sum + Number(r.SessionTimeHours || 0), 0);
  const totalJobs = rows.reduce((sum, r) => sum + Number(r.TotalJobScan || 0), 0);
  const totalProduced = rows.reduce((sum, r) => sum + Number(r.WorkProduced || 0), 0);
  const adjustedCount = rows.filter(r => r.AdjustmentID).length;

  return `
    <section class="adjustment-command-panel">
      <div class="section-head compact">
        <div>
          <h3>Adjustment Control</h3>
          <p>Use the date scorecards above to edit time. Detailed records stay collapsed to keep the profile clean.</p>
        </div>
        <button class="small-action" type="button" onclick="toggleAdjustmentRecords()">Show Records</button>
      </div>

      <div class="adjustment-summary-strip">
        <article><span>Records</span><strong>${rows.length}</strong></article>
        <article><span>Hours</span><strong>${formatDecimal(totalHours)}</strong></article>
        <article><span>Jobs</span><strong>${formatNumber(totalJobs)}</strong></article>
        <article><span>Produced</span><strong>${formatDecimal(totalProduced)}</strong></article>
        <article><span>Adjusted</span><strong>${formatNumber(adjustedCount)}</strong></article>
      </div>

      <div id="adjustmentRecordsDrawer" class="adjustment-records-drawer is-hidden">
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Day</th>
                <th>Date</th>
                <th>Department</th>
                <th>Area</th>
                <th>Access Point</th>
                <th>Hours</th>
                <th>Jobs</th>
                <th>JPH</th>
                <th>Target</th>
                <th>Work Produced</th>
                <th>Productivity %</th>
                <th>Status</th>
                <th>Note</th>
                <th>Edit</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map(r => {
                const key = getRecordKey(r);
                const adjusted = r.AdjustmentID ? true : false;
                const note = r.AdjustmentNote || r.Note || '';
                const reason = r.AdjustmentReason || r.Reason || '';

                return `
                  <tr class="${adjusted ? 'adjusted-row' : ''}">
                    <td>${getDayName(r.WorkDate)}</td>
                    <td>${formatDatePretty(r.WorkDate)}</td>
                    <td>${escapeHtml(r.FinalDepartment)}</td>
                    <td>${escapeHtml(r.StationGroup)}</td>
                    <td>${escapeHtml(r.AccessPoint)}</td>
                    <td>
                      <strong>${formatDecimal(r.SessionTimeHours)}</strong>
                      ${adjusted ? `<small class="adjustment-sub">was ${formatDecimal(r.OriginalSessionHours)}</small>` : ''}
                    </td>
                    <td>${formatNumber(r.TotalJobScan)}</td>
                    <td>${formatDecimal(r.AveragePerHour)}</td>
                    <td>${Number(r.TargetAvgPerHour || 0) > 0 ? formatDecimal(r.TargetAvgPerHour) : '—'}</td>
                    <td>${Number(r.WorkProduced || 0) > 0 ? formatDecimal(r.WorkProduced) : '—'}</td>
                    <td>${Number(r.TargetAvgPerHour || 0) > 0 ? formatPercent(r.ProductivityPercent) : '—'}</td>
                    <td>${statusBadge(r.PerformanceStatus)}</td>
                    <td>
                      ${note ? `<span class="note-pill" title="${escapeHtml(reason)}">${escapeHtml(note)}</span>` : '—'}
                    </td>
                    <td><button class="small-action" type="button" onclick="openHoursNoteEditor('${escapeJs(key)}')">Edit</button></td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  `;
}

function toggleAdjustmentRecords() {
  const drawer = document.getElementById('adjustmentRecordsDrawer');
  if (!drawer) return;

  drawer.classList.toggle('is-hidden');

  const btn = drawer.closest('.adjustment-command-panel')?.querySelector('.section-head .small-action');
  if (btn) btn.textContent = drawer.classList.contains('is-hidden') ? 'Show Records' : 'Hide Records';
}

function getWeekRowsForActiveDept() {
  const activeDept = String(STATE.activeDept || 'FINISH').trim().toUpperCase();
  const activeShift = String(STATE.activeShift || 'WEEKDAY').trim().toUpperCase();
  const activeAP = String(STATE.activeStationGroup || 'ALL').trim();

  const startDate = STATE.selectedWeekStartDate || STATE.meta.weekStartDate;
  const endDate = STATE.selectedWeekEndDate || STATE.meta.weekEndDate;
  const start = parseLocalDate(startDate);
  const end = parseLocalDate(endDate);

  // Determine which FinalDepartments to include
  const PROD_DEPTS = ['FINISH', 'SURFACE', 'AR'];
  const DIST_DEPTS = ['INVENTORY'];

  function deptMatch(r) {
    const rowDept = String(r.FinalDepartment || '').trim().toUpperCase();
    if (activeDept === 'PROD_ALL') return PROD_DEPTS.includes(rowDept);
    if (activeDept === 'DIST_ALL') return DIST_DEPTS.includes(rowDept);
    return rowDept === activeDept;
  }

  // Build schedule map for shift filtering
  const scheduleMap = new Map();
  (STATE.associateSchedules || []).forEach(s => {
    scheduleMap.set(String(s.OperatorName || '').trim(), String(s.ScheduleType || 'WEEKDAY').trim().toUpperCase());
  });

  function shiftMatch(r) {
    if (activeShift === 'ALL') return true;
    // Machine/equipment rows show on both shifts
    if (isMachineRow(r)) return true;
    const opName = String(r.OperatorName || '').trim();
    const opShift = scheduleMap.get(opName) || 'WEEKDAY';
    return opShift === activeShift;
  }

  const snapshotRows = STATE.dailySnapshot.filter(r => {
    const d = parseLocalDate(r.WorkDate);
    const dateMatch = d && start && end && d >= start && d <= end;
    return deptMatch(r) && dateMatch && shiftMatch(r);
  });

  const combined = dedupeDailyRows(snapshotRows);

  return combined.filter(r => {
    const apMatch = activeAP === 'ALL' || String(r.StationGroup || '').trim() === activeAP;
    const blob = [r.OperatorName, r.StationGroup, r.AccessPoint, r.FinalDepartment].join(' ').toLowerCase();
    const searchMatch = !STATE.search || blob.includes(STATE.search);
    return apMatch && searchMatch;
  });
}

function buildAssociateSummaries() {
  const rows = getWeekRowsForActiveDept();
  const map = new Map();

  rows.forEach(row => {
    const name = String(row.OperatorName || 'Unknown Operator').trim();
    if (!map.has(name)) map.set(name, []);
    map.get(name).push(row);
  });

  return Array.from(map.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([associate, associateRows]) => {
      const areas = buildAreaSummaries(associateRows);
      const totalHours = associateRows.reduce((sum, r) => sum + Number(r.SessionTimeHours || 0), 0);
      const totalJobs = associateRows.reduce((sum, r) => sum + Number(r.TotalJobScan || 0), 0);
      const avgJph = totalHours > 0 ? totalJobs / totalHours : 0;
      const workProduced = getWorkProducedHours(associateRows);
      const weightedTarget = getWeightedTarget(associateRows); // kept for reference only
      const productivityPercent = totalHours > 0 && workProduced > 0 ? workProduced / totalHours : 0;
      const status = getStatusFromPercent(productivityPercent, workProduced);
      const workedDateKeys = new Set(associateRows.map(r => formatDateKey(r.WorkDate)).filter(Boolean));
      const workDays = workedDateKeys.size;
      const departments = Array.from(new Set(associateRows.map(r => r.FinalDepartment).filter(Boolean))).join(', ');
      const workDayInfo = getWeeklyWorkDayInfo(departments || STATE.activeDept, associate);
      const scheduledDaysWorked = (workDayInfo.eligibleDateKeys || [])
        .filter(dateKey => workedDateKeys.has(dateKey)).length;

      return {
        associate,
        departments,
        rows: associateRows,
        areas,
        totalHours,
        totalJobs,
        avgJph,
        weightedTarget,
        workProduced,
        productivityPercent,
        status,
        workDays,
        scheduledDaysWorked,
        scheduleType: workDayInfo.scheduleType,
        scheduledDays: workDayInfo.scheduledDays,
        scheduledDayGoal: workDayInfo.scheduledDayGoal,
        eligibleWorkDays: workDayInfo.eligibleWorkDays,
        excludedDays: workDayInfo.excludedDays,
        excludedDates: workDayInfo.excludedDates,
        eligibleDateKeys: workDayInfo.eligibleDateKeys,
        scheduledDateKeys: workDayInfo.scheduledDateKeys
      };
    });
}

function buildAreaSummaries(rows) {
  const map = new Map();

  rows.forEach(row => {
    const key = [row.FinalDepartment || '', row.StationGroup || '', row.AccessPoint || ''].join('|');
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  });

  return Array.from(map.entries()).map(([key, areaRows]) => {
    const [department, stationGroup, accessPoint] = key.split('|');
    const totalHours = areaRows.reduce((sum, r) => sum + Number(r.SessionTimeHours || 0), 0);
    const totalJobs = areaRows.reduce((sum, r) => sum + Number(r.TotalJobScan || 0), 0);
    const avgJph = totalHours > 0 ? totalJobs / totalHours : 0;
    const workProduced = getWorkProducedHours(areaRows);
    const weightedTarget = getWeightedTarget(areaRows); // kept for reference only
    const productivityPercent = totalHours > 0 && workProduced > 0 ? workProduced / totalHours : 0;
    const status = getStatusFromPercent(productivityPercent, workProduced);
    const workDays = new Set(areaRows.map(r => formatDateKey(r.WorkDate)).filter(Boolean)).size;

    return {
      department,
      stationGroup,
      accessPoint,
      rows: areaRows,
      totalHours,
      totalJobs,
      avgJph,
      weightedTarget,
      workProduced,
      productivityPercent,
      status,
      workDays
    };
  }).sort((a, b) => {
    const stationCompare = String(a.stationGroup).localeCompare(String(b.stationGroup));
    if (stationCompare !== 0) return stationCompare;
    return String(a.accessPoint).localeCompare(String(b.accessPoint));
  });
}

function aggregateAssociateByDay(rows) {
  const map = new Map();

  rows.forEach(row => {
    const key = formatDateKey(row.WorkDate);
    if (!key) return;

    if (!map.has(key)) {
      map.set(key, {
        workDate: row.WorkDate,
        rows: [],
        totalHours: 0,
        totalJobs: 0
      });
    }

    const day = map.get(key);
    day.rows.push(row);
    day.totalHours += Number(row.SessionTimeHours || 0);
    day.totalJobs += Number(row.TotalJobScan || 0);
  });

  return Array.from(map.values()).map(day => {
    const target = getWeightedTarget(day.rows); // kept for display/reference
    const workProduced = getWorkProducedHours(day.rows);
    const avgJph = day.totalHours > 0 ? day.totalJobs / day.totalHours : 0;
    const productivityPercent = day.totalHours > 0 && workProduced > 0 ? workProduced / day.totalHours : 0;

    return {
      ...day,
      target,
      workProduced,
      productivityPercent,
      avgJph
    };
  }).sort((a, b) => parseLocalDate(a.workDate) - parseLocalDate(b.workDate));
}

function getWeightedTarget(rows) {
  const valid = rows.filter(r => Number(r.TargetAvgPerHour || 0) > 0 && Number(r.SessionTimeHours || 0) > 0);

  if (!valid.length) {
    const any = rows.find(r => Number(r.TargetAvgPerHour || 0) > 0);
    return any ? Number(any.TargetAvgPerHour || 0) : 0;
  }

  const hours = valid.reduce((sum, r) => sum + Number(r.SessionTimeHours || 0), 0);
  if (hours <= 0) return 0;

  return valid.reduce((sum, r) => {
    return sum + (Number(r.TargetAvgPerHour || 0) * Number(r.SessionTimeHours || 0));
  }, 0) / hours;
}

function getWorkProducedHours(rows) {
  return (rows || []).reduce((sum, r) => {
    const jobs = Number(r.TotalJobScan || 0);
    const target = Number(r.TargetAvgPerHour || 0);
    if (target <= 0 || jobs <= 0) return sum;
    return sum + (jobs / target);
  }, 0);
}

function getStatusFromPercent(percent, workProduced) {
  if (!workProduced || workProduced <= 0) return 'NO TARGET';

  const ratio = Number(percent || 0);

  // Higher-is-better model:
  // Work Produced / Hours Scanned
  // 100% or higher = GREEN, 90% to 99.99% = AMBER, below 90% = RED.
  if (ratio >= 1.00) return 'GREEN';
  if (ratio >= 0.90) return 'AMBER';
  return 'RED';
}




/************************************************************
 * CLIENT-SIDE PRODUCTIVITY RECALCULATION
 *
 * Formula used for display/status:
 * Work Produced = TotalJobScan / TargetAvgPerHour
 * Productivity % = Work Produced / Hours Scanned
 *
 * This keeps 100%+ GREEN, 90%-99.99% AMBER, below 90% RED.
 ************************************************************/

function recalcClientProductivityRows(rows) {
  return (rows || []).map(row => recalcClientProductivityRow(row));
}

function recalcClientProductivityRow(row) {
  const copy = { ...row };
  const hours = Number(copy.SessionTimeHours || 0);
  const mins = Number(copy.SessionTimeMins || hours * 60 || 0);
  const jobs = Number(copy.TotalJobScan || 0);
  const target = Number(copy.TargetAvgPerHour || 0);

  const workProduced = target > 0 && jobs > 0 ? jobs / target : 0;

  copy.WorkProduced = workProduced;
  copy.AveragePerHour = hours > 0 ? jobs / hours : 0;
  copy.AveragePerMinute = mins > 0 ? jobs / mins : 0;
  copy.ProductivityPercent = hours > 0 && workProduced > 0 ? workProduced / hours : 0;
  copy.PerformanceStatus = getStatusFromPercent(copy.ProductivityPercent, workProduced);

  return copy;
}

/************************************************************
 * HOURS + NOTES ADJUSTMENTS
 ************************************************************/

function applyProductivityAdjustmentsToData() {
  const latest = new Map();

  (STATE.adjustments || []).forEach(adj => {
    const key = getAdjustmentKeyFromAdjustment(adj);
    if (!key) return;

    const existing = latest.get(key);
    const existingTime = new Date(existing?.EditedAt || 0).getTime();
    const newTime = new Date(adj.EditedAt || 0).getTime();

    if (!existing || newTime >= existingTime) {
      latest.set(key, adj);
    }
  });

  STATE.master = STATE.master.map(row => applyAdjustmentToRow(row, latest));
  STATE.dailySnapshot = STATE.dailySnapshot.map(row => applyAdjustmentToRow(row, latest));
}

function applyAdjustmentToRow(row, adjustmentMap) {
  const key = getRecordKey(row);
  const adj = adjustmentMap.get(key);

  if (!adj) return row;

  const copy = { ...row };

  const originalHours = Number(adj.OriginalSessionHours || copy.SessionTimeHours || 0);
  const adjustedHours = Number(adj.AdjustedSessionHours || copy.SessionTimeHours || 0);
  const totalJobs = Number(adj.TotalJobScan || copy.TotalJobScan || 0);
  const target = Number(adj.TargetAvgPerHour || copy.TargetAvgPerHour || 0);

  copy.OriginalSessionHours = originalHours;
  copy.OriginalSessionMins = Number(adj.OriginalSessionMins || originalHours * 60);
  copy.SessionTimeHours = adjustedHours;
  copy.SessionTimeMins = Number(adj.AdjustedSessionMins || adjustedHours * 60);
  copy.AveragePerHour = adjustedHours > 0 ? totalJobs / adjustedHours : 0;
  copy.AveragePerMinute = copy.SessionTimeMins > 0 ? totalJobs / copy.SessionTimeMins : 0;
  const workProduced = target > 0 ? totalJobs / target : 0;
  copy.WorkProduced = workProduced;
  copy.ProductivityPercent = adjustedHours > 0 && workProduced > 0 ? workProduced / adjustedHours : 0;
  copy.PerformanceStatus = getStatusFromPercent(copy.ProductivityPercent, workProduced);
  copy.AdjustmentID = adj.AdjustmentID;
  copy.EditSnapshotID = adj.EditSnapshotID;
  copy.AdjustmentReason = adj.Reason || '';
  copy.AdjustmentNote = adj.Note || '';
  copy.EditedBy = adj.EditedBy || '';
  copy.EditedAt = adj.EditedAt || '';

  return copy;
}


function openAssociateDayAdjustment(dateKey, associateName) {
  const associate = String(associateName || STATE.selectedAssociate || '').trim();
  const dayRows = getWeekRowsForActiveDept().filter(row =>
    normalizeQualityText(row.OperatorName) === normalizeQualityText(associate) &&
    formatDateKey(row.WorkDate) === String(dateKey || '').trim()
  );

  if (!dayRows.length) {
    alert('No editable productivity record found for this date.');
    return;
  }

  // If there is only one row for the date, open the normal hours editor immediately.
  if (dayRows.length === 1) {
    openHoursNoteEditor(getRecordKey(dayRows[0]));
    return;
  }

  // If multiple areas/access points were worked on the same date, let the user choose.
  const existing = document.getElementById('dayAdjustmentPickerModal');
  if (existing) existing.remove();

  const totalHours = dayRows.reduce((sum, r) => sum + Number(r.SessionTimeHours || 0), 0);
  const totalJobs = dayRows.reduce((sum, r) => sum + Number(r.TotalJobScan || 0), 0);
  const totalProduced = dayRows.reduce((sum, r) => sum + Number(r.WorkProduced || 0), 0);

  const modal = document.createElement('div');
  modal.id = 'dayAdjustmentPickerModal';
  modal.className = 'modal-backdrop';
  modal.innerHTML = `
    <div class="edit-modal day-adjustment-modal">
      <div class="modal-head">
        <div>
          <h3>Edit Time by Date</h3>
          <p>${escapeHtml(associate)} • ${formatDatePretty(dateKey)}</p>
        </div>
        <button type="button" class="modal-x" onclick="closeAssociateDayAdjustment()">×</button>
      </div>

      <div class="modal-grid">
        <div class="readonly-box">
          <span>Total Hours</span>
          <strong>${formatDecimal(totalHours)}</strong>
        </div>
        <div class="readonly-box">
          <span>Total Jobs</span>
          <strong>${formatNumber(totalJobs)}</strong>
        </div>
        <div class="readonly-box">
          <span>Produced</span>
          <strong>${formatDecimal(totalProduced)}</strong>
        </div>
      </div>

      <div class="day-adjustment-list">
        ${dayRows.map(row => {
          const key = getRecordKey(row);
          const adjusted = row.AdjustmentID ? true : false;
          return `
            <button type="button" class="day-adjustment-option ${adjusted ? 'adjusted' : ''}" onclick="closeAssociateDayAdjustment(); openHoursNoteEditor('${escapeJs(key)}')">
              <div>
                <strong>${escapeHtml(row.StationGroup || 'No Area')}</strong>
                <span>${escapeHtml(row.AccessPoint || row.FinalDepartment || '')}</span>
              </div>
              <div class="day-adjustment-option-metrics">
                <span><b>${formatDecimal(row.SessionTimeHours)}</b> hrs</span>
                <span><b>${formatNumber(row.TotalJobScan)}</b> jobs</span>
                <span><b>${Number(row.TargetAvgPerHour || 0) > 0 ? formatPercent(row.ProductivityPercent) : 'N/A'}</b></span>
              </div>
              ${adjusted ? '<em>Adjusted</em>' : ''}
            </button>
          `;
        }).join('')}
      </div>
    </div>
  `;

  document.body.appendChild(modal);
}

function closeAssociateDayAdjustment() {
  const modal = document.getElementById('dayAdjustmentPickerModal');
  if (modal) modal.remove();
}


function openHoursNoteEditor(recordKey) {
  const row = findRecordByKey(recordKey);
  if (!row) {
    alert('Could not find this record. Refresh the page and try again.');
    return;
  }

  const existing = document.getElementById('hoursNoteModal');
  if (existing) existing.remove();

  const originalHours = Number(row.OriginalSessionHours || row.SessionTimeHours || 0);
  const currentHours = Number(row.SessionTimeHours || 0);
  const currentJph = Number(row.AveragePerHour || 0);
  const jobs = Number(row.TotalJobScan || 0);

  let reason = row.AdjustmentReason || '';
  if (reason === 'Indirect Work') reason = 'Non-Productive Time';

  const note = row.AdjustmentNote || '';

  const reasonOptions = [
    'Supervisor Adjustment',
    'Missed Login / Logout',
    'Non-Productive Time',
    'Floater Support',
    'Moved to Another Department',
    'Training / Support',
    'System Issue',
    'Correction',
    'Other'
  ];

  const modal = document.createElement('div');
  modal.id = 'hoursNoteModal';
  modal.className = 'modal-backdrop hours-adjustment-backdrop';
  modal.innerHTML = `
    <div class="edit-modal hours-adjustment-modal">
      <div class="modal-head hours-modal-head">
        <div>
          <span class="modal-eyebrow">Productivity Adjustment</span>
          <h3>Edit Hours + Supervisor Note</h3>
          <p>${escapeHtml(row.OperatorName)} • ${escapeHtml(row.StationGroup)} • ${formatDatePretty(row.WorkDate)}</p>
        </div>
        <button type="button" class="modal-x" onclick="closeHoursNoteEditor()" aria-label="Close adjustment modal">×</button>
      </div>

      <div class="hours-adjustment-body">
        <section class="hours-adjustment-summary">
          <div class="readonly-box">
            <span>Original Hours</span>
            <strong>${formatDecimal(originalHours)}</strong>
          </div>
          <div class="readonly-box">
            <span>Current Hours</span>
            <strong>${formatDecimal(currentHours)}</strong>
          </div>
          <div class="readonly-box">
            <span>Jobs</span>
            <strong>${formatNumber(jobs)}</strong>
          </div>
          <div class="readonly-box">
            <span>Current JPH</span>
            <strong>${formatDecimal(currentJph)}</strong>
          </div>
        </section>

        <section class="hours-adjustment-form">
          <label class="modal-field hours-field-primary">
            <span>Adjusted Hours</span>
            <input id="editAdjustedHours" type="number" min="0" step="0.01" value="${formatDecimal(currentHours)}" />
          </label>

          <div class="hours-quick-row">
            <button type="button" onclick="setAdjustedHoursQuick('${formatDecimal(originalHours)}')">Original</button>
            <button type="button" onclick="setAdjustedHoursQuick('${formatDecimal(Math.max(currentHours - 0.25, 0))}')">-0.25 hr</button>
            <button type="button" onclick="setAdjustedHoursQuick('${formatDecimal(Math.max(currentHours - 0.50, 0))}')">-0.50 hr</button>
            <button type="button" onclick="setAdjustedHoursQuick('${formatDecimal(Math.max(currentHours - 1.00, 0))}')">-1.00 hr</button>
            <button type="button" onclick="setAdjustedHoursQuick('0.00')">Set 0</button>
          </div>

          <label class="modal-field">
            <span>Adjustment Reason</span>
            <select id="editAdjustmentReason">
              ${reasonOptions.map(opt => `
                <option value="${escapeHtml(opt)}" ${opt === reason ? 'selected' : ''}>${escapeHtml(opt)}</option>
              `).join('')}
            </select>
          </label>

          <label class="modal-field note-field-wide">
            <span>Supervisor Note</span>
            <textarea id="editAdjustmentNote" rows="5" placeholder="Example: Associate moved to Final Inspection after lunch, supporting as floater, or non-productive time approved by supervisor.">${escapeHtml(note)}</textarea>
          </label>
        </section>
      </div>

      <div class="modal-actions hours-modal-actions">
        <button type="button" class="secondary-btn" onclick="closeHoursNoteEditor()">Cancel</button>
        <button type="button" class="primary-btn" onclick="saveHoursNoteAdjustment('${escapeJs(recordKey)}')">Save Adjustment</button>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  setTimeout(() => {
    const hoursInput = document.getElementById('editAdjustedHours');
    if (hoursInput) hoursInput.focus();
  }, 50);
}

function setAdjustedHoursQuick(value) {
  const input = document.getElementById('editAdjustedHours');
  if (!input) return;
  input.value = value;
  input.focus();
}

function closeHoursNoteEditor() {
  const modal = document.getElementById('hoursNoteModal');
  if (modal) modal.remove();
}

function saveHoursNoteAdjustment(recordKey) {
  const row = findRecordByKey(recordKey);
  if (!row) {
    alert('Could not find this record. Refresh the page and try again.');
    return;
  }

  const hoursEl = document.getElementById('editAdjustedHours');
  const reasonEl = document.getElementById('editAdjustmentReason');
  const noteEl = document.getElementById('editAdjustmentNote');

  const adjustedHours = Number(hoursEl ? hoursEl.value : row.SessionTimeHours);
  let reason = reasonEl ? reasonEl.value : 'Supervisor Adjustment';
  if (reason === 'Indirect Work') reason = 'Non-Productive Time';
  const note = noteEl ? noteEl.value : '';

  if (isNaN(adjustedHours) || adjustedHours < 0) {
    alert('Adjusted hours must be 0 or higher.');
    return;
  }

  const saveBtn = document.querySelector('#hoursNoteModal .primary-btn');
  if (saveBtn) {
    saveBtn.disabled = true;
    saveBtn.textContent = 'Saving...';
  }

  const params = new URLSearchParams({
    action: 'saveProductivityAdjustment',
    workDate: formatDateKey(row.WorkDate),
    finalDepartment: row.FinalDepartment || '',
    stationGroup: row.StationGroup || '',
    department: row.Department || row.FinalDepartment || '',
    accessPoint: row.AccessPoint || '',
    operatorName: row.OperatorName || '',
    originalSessionHours: String(Number(row.OriginalSessionHours || row.SessionTimeHours || 0)),
    adjustedSessionHours: String(adjustedHours),
    totalJobScan: String(Number(row.TotalJobScan || 0)),
    targetAvgPerHour: String(Number(row.TargetAvgPerHour || 0)),
    reason,
    note,
    editedBy: getCurrentProductivityUsername() || 'WEBPAGE'
  });

  fetch(`${API_URL}?${params.toString()}&ts=${Date.now()}`)
    .then(r => r.text())
    .then(text => {
      let response;
      try {
        response = JSON.parse(text);
      } catch (err) {
        throw new Error('Save adjustment API did not return JSON.');
      }

      if (!response.ok) {
        throw new Error(response.message || 'Save adjustment failed.');
      }

      closeHoursNoteEditor();

      if (typeof clearClientProductivityCache === 'function') {
        clearClientProductivityCache();
      }

      // Remove selected week from JS memory cache so the edited hours refresh immediately.
      const selectedWeekKey = getSelectedFiscalWeekKey();
      if (selectedWeekKey && STATE.weekDataCache) {
        delete STATE.weekDataCache[selectedWeekKey];
      }

      reloadData();
    })
    .catch(err => {
      const saveBtn = document.querySelector('#hoursNoteModal .primary-btn');
      if (saveBtn) {
        saveBtn.disabled = false;
        saveBtn.textContent = 'Save Adjustment';
      }
      alert('Could not save adjustment: ' + (err.message || err));
    });
}

function findRecordByKey(recordKey) {
  const rows = getWeekRowsForActiveDept();
  return rows.find(row => getRecordKey(row) === recordKey) || null;
}

function getRecordKey(row) {
  return [
    formatDateKey(row.WorkDate),
    row.FinalDepartment || '',
    row.StationGroup || '',
    row.AccessPoint || '',
    row.OperatorName || ''
  ].join('|');
}

function getAdjustmentKeyFromAdjustment(adj) {
  return [
    formatDateKey(adj.WorkDate),
    adj.FinalDepartment || '',
    adj.StationGroup || '',
    adj.AccessPoint || '',
    adj.OperatorName || ''
  ].join('|');
}

/************************************************************
 * ASSOCIATE DAY RULES - WEBPAGE SAVE
 ************************************************************/

function getDefaultRuleLabel(type) {
  const t = String(type || '').toUpperCase();
  if (t === 'HOLIDAY') return 'Holiday';
  if (t === 'EXCUSED') return 'Excused Day';
  return 'No OT Scheduled';
}


function renderAssociateRuleList(summary) {
  const rules = getAssociateRulesForSummary(summary);

  if (!rules.length) {
    return `<p class="rule-empty">No associate-specific excluded days saved.</p>`;
  }

  return `
    <div class="rule-list">
      ${rules.map(rule => `
        <div class="rule-chip">
          <div>
            <strong>${formatDatePretty(rule.WorkDate)}</strong>
            <span>${escapeHtml(rule.Label || statusLabel(rule.Type))} • ${escapeHtml(rule.Type)}</span>
          </div>
          <button type="button" onclick="deleteAssociateDayRule('${escapeJs(rule.RuleID)}')">Remove</button>
        </div>
      `).join('')}
    </div>
  `;
}

function getAssociateRulesForSummary(summary) {
  const weekKeys = new Set(getWeekDateList().map(formatDateKey));
  const departments = String(summary.departments || STATE.activeDept || '').toUpperCase();

  return STATE.associateDayRules
    .filter(rule => String(rule.OperatorName || '').trim() === String(summary.associate || '').trim())
    .filter(rule => weekKeys.has(formatDateKey(rule.WorkDate)))
    .filter(rule => associateRuleAppliesToDepartment(rule, departments))
    .sort((a, b) => parseLocalDate(a.WorkDate) - parseLocalDate(b.WorkDate));
}

function saveAssociateDayRuleFromProfile(associateName, departmentText) {
  const dateEl = document.getElementById('associateRuleDate');
  const typeEl = document.getElementById('associateRuleType');
  const labelEl = document.getElementById('associateRuleLabel');
  const noteEl = document.getElementById('associateRuleNote');

  const date = dateEl ? dateEl.value : '';
  const type = typeEl ? typeEl.value : 'NO_OT';
  const label = labelEl && labelEl.value ? labelEl.value : getDefaultRuleLabel(type);
  const note = noteEl ? noteEl.value : '';

  if (!date || !associateName) {
    alert('Missing date or associate.');
    return;
  }

  const params = new URLSearchParams({
    action: 'saveAssociateDayRule',
    date,
    operatorName: associateName,
    finalDepartment: departmentText || STATE.activeDept || 'ALL',
    type,
    label,
    excludeFromWeeklyWorkDays: 'YES',
    countIfWorked: 'YES',
    note,
    createdBy: 'WEBPAGE'
  });

  fetch(`${API_URL}?${params.toString()}&ts=${Date.now()}`)
    .then(r => r.text())
    .then(text => {
      let response;
      try {
        response = JSON.parse(text);
      } catch (err) {
        throw new Error('Save API did not return JSON.');
      }

      if (!response.ok) {
        throw new Error(response.message || 'Save failed.');
      }

      reloadData();
    })
    .catch(err => {
      alert('Could not save associate day rule: ' + (err.message || err));
    });
}

function deleteAssociateDayRule(ruleId) {
  if (!ruleId) return;

  const params = new URLSearchParams({
    action: 'deleteAssociateDayRule',
    ruleId
  });

  fetch(`${API_URL}?${params.toString()}&ts=${Date.now()}`)
    .then(r => r.text())
    .then(text => {
      let response;
      try {
        response = JSON.parse(text);
      } catch (err) {
        throw new Error('Delete API did not return JSON.');
      }

      if (!response.ok) {
        throw new Error(response.message || 'Delete failed.');
      }

      reloadData();
    })
    .catch(err => {
      alert('Could not remove associate day rule: ' + (err.message || err));
    });
}

function getAssociateRuleForDate(dateValue, associateName, departmentText) {
  const key = formatDateKey(dateValue);
  if (!key || !associateName) return null;

  const rules = STATE.associateDayRules
    .filter(rule => formatDateKey(rule.WorkDate) === key)
    .filter(rule => String(rule.OperatorName || '').trim() === String(associateName || '').trim())
    .filter(rule => String(rule.ExcludeFromWeeklyWorkDays || 'YES').toUpperCase() === 'YES')
    .filter(rule => associateRuleAppliesToDepartment(rule, departmentText));

  if (!rules.length) return null;

  // Latest saved active rule wins.
  return rules.slice().sort((a, b) => {
    return new Date(b.CreatedAt || 0).getTime() - new Date(a.CreatedAt || 0).getTime();
  })[0];
}

function associateRuleAppliesToDepartment(rule, departmentText) {
  const appliesTo = String(rule.FinalDepartment || 'ALL').trim().toUpperCase();
  if (appliesTo === 'ALL') return true;

  const department = String(departmentText || STATE.activeDept || '').toUpperCase();

  return department
    .split(',')
    .map(v => v.trim())
    .some(v => v === appliesTo);
}

/************************************************************
 * HOLIDAYS / NO OT DAYS
 ************************************************************/

function getWeekDateList() {
  const start = parseLocalDate(STATE.selectedWeekStartDate || STATE.meta.weekStartDate);
  if (!start) return [];

  return Array.from({ length: 7 }, (_, index) => {
    const d = new Date(start);
    d.setDate(start.getDate() + index);
    return d;
  });
}

function getHolidayForDate(dateValue, departmentText, associateName) {
  // Holiday / No-OT / associate day rules removed.
  // Fiscal week is the hard date range; shift is only view guidance.
  return null;
}

function holidayAppliesToDepartment(holiday, departmentText) {
  return false;
}

function getAssociateRuleForDate(dateValue, associateName, departmentText) {
  return null;
}

function associateRuleAppliesToDepartment(rule, departmentText) {
  return false;
}

function getAssociateSchedule(operatorName) {
  const name = String(operatorName || '').trim().toUpperCase();

  const found = (STATE.associateSchedules || []).find(row => {
    return String(row.OperatorName || '').trim().toUpperCase() === name;
  });

  if (found) {
    const scheduleType = String(found.ScheduleType || 'WEEKDAY').trim().toUpperCase();
    return {
      scheduleType,
      scheduledDays: parseScheduledDays(found.ScheduledDays || defaultDaysForScheduleType(scheduleType))
    };
  }

  return {
    scheduleType: 'WEEKDAY',
    scheduledDays: ['MON', 'TUE', 'WED', 'THU']
  };
}

function parseScheduledDays(value) {
  return String(value || '')
    .toUpperCase()
    .split(',')
    .map(day => day.trim())
    .filter(Boolean);
}

function getDayCode(dateValue) {
  const d = parseLocalDate(dateValue);
  if (!d) return '';

  const codes = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];
  return codes[d.getDay()];
}

function formatScheduleType(value) {
  const type = String(value || 'WEEKDAY').trim().toUpperCase();
  if (type === 'WEEKEND') return 'Weekend';
  if (type === 'CUSTOM') return 'Custom';
  return 'Weekday';
}

function getWeeklyWorkDayInfo(departmentText, associateName) {
  const schedule = getAssociateSchedule(associateName);
  const weekDates = getWeekDateList();

  const scheduledDates = weekDates.filter(dateValue => {
    const dayCode = getDayCode(dateValue);
    return schedule.scheduledDays.includes(dayCode);
  });

  return {
    scheduleType: schedule.scheduleType,
    scheduledDays: schedule.scheduledDays,
    scheduledDayGoal: scheduledDates.length,
    totalWeekDays: weekDates.length,
    excludedDays: 0,
    eligibleWorkDays: scheduledDates.length,
    excludedDates: [],
    eligibleDateKeys: scheduledDates.map(formatDateKey),
    scheduledDateKeys: scheduledDates.map(formatDateKey)
  };
}

/************************************************************
 * DATES / DEDUPE
 ************************************************************/

function parseLocalDate(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const text = String(value).trim();
  const match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  if (match) {
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  }

  const parsed = new Date(text);
  if (isNaN(parsed)) return null;

  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function formatDateKey(value) {
  const d = parseLocalDate(value);
  if (!d) return '';

  return [
    d.getFullYear(),
    String(d.getMonth() + 1).padStart(2, '0'),
    String(d.getDate()).padStart(2, '0')
  ].join('-');
}

function dedupeDailyRows(rows) {
  const map = new Map();

  rows.forEach(row => {
    const key = [
      formatDateKey(row.WorkDate),
      row.FinalDepartment,
      row.StationGroup,
      row.AccessPoint,
      row.OperatorName
    ].join('|');

    if (!map.has(key)) {
      map.set(key, row);
      return;
    }

    const existing = map.get(key);
    const existingTime = new Date(existing.SnapshotCreatedAt || 0).getTime();
    const newTime = new Date(row.SnapshotCreatedAt || 0).getTime();

    if (newTime >= existingTime) {
      map.set(key, row);
    }
  });

  return Array.from(map.values());
}

/************************************************************
 * UTILS
 ************************************************************/

function showLoading() {
  const page = document.getElementById('pageContent');
  if (!page) return;

  page.innerHTML = `
    <div class="loading-card command-loading-card">
      <div class="spinner"></div>
      <strong>Loading productivity profile...</strong>
      <span>Syncing scorecards, associate rules, and saved daily records.</span>
    </div>
  `;
}

function startLoadingScreen() {
  const overlay = document.getElementById('appLoader');
  const bar = document.getElementById('loaderProgressBar');
  const pct = document.getElementById('loaderPercent');
  const step = document.getElementById('loaderStepText');
  const clockEl = document.getElementById('ldrClock');

  if (!overlay) return;

  overlay.classList.remove('hide');
  overlay.classList.add('show');

  // Live clock
  function tickClock() {
    if (!clockEl) return;
    const now = new Date();
    const h = String(now.getHours()).padStart(2,'0');
    const m = String(now.getMinutes()).padStart(2,'0');
    const s = String(now.getSeconds()).padStart(2,'0');
    clockEl.textContent = h + ':' + m + ':' + s;
  }
  tickClock();
  const clockTimer = setInterval(tickClock, 1000);
  overlay._clockTimer = clockTimer;

  // Step status indicators
  const stepEls = [
    document.getElementById('ldrS1'),
    document.getElementById('ldrS2'),
    document.getElementById('ldrS3'),
    document.getElementById('ldrS4')
  ];

  function setStep(idx, state) {
    const el = stepEls[idx];
    if (!el) return;
    const dot = el.querySelector('.ldr-status-dot');
    el.classList.toggle('active', state === 'running');
    el.classList.toggle('done', state === 'done');
    if (dot) {
      dot.className = 'ldr-status-dot';
      if (state === 'running') dot.classList.add('running');
      else if (state === 'done') dot.classList.add('done');
      else dot.classList.add('pending');
    }
  }

  let progress = 0;
  const stepMessages = [
    { pct: 15,  text: 'Connecting to Productivity API...', step: 0, state: 'running' },
    { pct: 35,  text: 'Loading daily snapshots...', step: 0, state: 'done', next: [1, 'running'] },
    { pct: 55,  text: 'Building associate scorecards...', step: 1, state: 'done', next: [2, 'running'] },
    { pct: 72,  text: 'Applying day rules and adjustments...', step: 2, state: 'running' },
    { pct: 88,  text: 'Preparing command center view...', step: 2, state: 'done', next: [3, 'running'] }
  ];

  setStep(0, 'running');
  if (bar) bar.style.width = '0%';
  if (pct) pct.textContent = '0%';
  if (step) step.textContent = stepMessages[0].text;

  clearInterval(STATE.loadingTimer);
  STATE.loadingTimer = setInterval(() => {
    progress = Math.min(progress + Math.floor(Math.random() * 7) + 3, 91);
    if (bar) bar.style.width = progress + '%';
    if (pct) pct.textContent = progress + '%';

    const current = [...stepMessages].reverse().find(s => progress >= s.pct);
    if (current) {
      if (step) step.textContent = current.text;
      setStep(current.step, current.state);
      if (current.next) setStep(current.next[0], current.next[1]);
    }

    if (progress >= 91) clearInterval(STATE.loadingTimer);
  }, 220);
}

function finishLoadingScreen() {
  const overlay = document.getElementById('appLoader');
  const bar = document.getElementById('loaderProgressBar');
  const pct = document.getElementById('loaderPercent');
  const step = document.getElementById('loaderStepText');

  clearInterval(STATE.loadingTimer);
  if (overlay && overlay._clockTimer) clearInterval(overlay._clockTimer);

  if (bar) bar.style.width = '100%';
  if (pct) pct.textContent = '100%';
  if (step) step.textContent = 'Ready — opening Productivity Hub...';

  // Mark all steps done
  ['ldrS1','ldrS2','ldrS3','ldrS4'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.remove('active'); el.classList.add('done');
    const dot = el.querySelector('.ldr-status-dot');
    if (dot) { dot.className = 'ldr-status-dot done'; }
  });

  if (!overlay) return;
  setTimeout(() => {
    overlay.classList.add('hide');
    overlay.classList.remove('show');
  }, 100);
}

function showRefreshStatus(message) {
  const el = document.getElementById('generatedAt');
  if (el) el.textContent = message;
}

function showError(err) {
  clearInterval(STATE.loadingTimer);
  const overlay = document.getElementById('appLoader');
  if (overlay) overlay.classList.add('hide');

  const page = document.getElementById('pageContent');
  if (!page) return;

  page.innerHTML = `
    <div class="loading-card command-loading-card error-card">
      <strong>Failed to load productivity data</strong>
      <span>${escapeHtml(err.message || err)}</span>
    </div>
  `;
}

function animateCounters() {
  // Kept light. Actual values already render immediately.
}

function statusLabel(status) {
  const s = normalizeStatus(status);
  if (s === 'GREEN') return 'On Target';
  if (s === 'AMBER') return 'Watch';
  if (s === 'RED') return 'Warning';
  if (s === 'NO TARGET') return 'No Target';
  if (s === 'NO SCAN MATCH') return 'No Scan Match';
  if (s === 'HOLIDAY') return 'Holiday';
  if (s === 'NO_OT') return 'No OT';
  if (s === 'EXCUSED') return 'Excused';
  return s;
}

function normalizeStatus(status) {
  return String(status || 'NO TARGET').toUpperCase();
}

function statusCss(status) {
  return normalizeStatus(status).replaceAll(' ', '-');
}

function statusBadge(status) {
  const s = normalizeStatus(status);
  const css = statusCss(s);
  return `<span class="badge ${css}">${statusLabel(s)}</span>`;
}

function getDayName(dateValue) {
  const d = parseLocalDate(dateValue);
  if (!d) return '--';
  return d.toLocaleDateString(undefined, { weekday: 'long' });
}

function getDayShort(dateValue) {
  const d = parseLocalDate(dateValue);
  if (!d) return '--';
  return d.toLocaleDateString(undefined, { weekday: 'short' });
}

function formatDatePretty(dateValue) {
  const d = parseLocalDate(dateValue);
  if (!d) return '--';

  return d.toLocaleDateString(undefined, {
    month: 'short',
    day: 'numeric',
    year: 'numeric'
  });
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString();
}

function formatDecimal(value) {
  return Number(value || 0).toFixed(2);
}

function formatPercent(value) {
  return `${Math.round(Number(value || 0) * 100)}%`;
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}


function escapeJs(value) {
  return String(value ?? '')
    .replaceAll('\\', '\\\\')
    .replaceAll("'", "\\'")
    .replaceAll('"', '\\"')
    .replaceAll('\n', ' ')
    .replaceAll('\r', ' ');
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}


/************************************************************
 * SIDE PANEL
 ************************************************************/

function renderSidePanel(summary) {
  if (!summary) {
    return `
      <div class="side-title">
        <div>
          <h3>No associate selected</h3>
          <p>Select an associate to view details.</p>
        </div>
      </div>
    `;
  }

  const dayRows = typeof aggregateAssociateByDay === 'function'
    ? aggregateAssociateByDay(summary.rows || [])
    : [];

  const maxJph = Math.max(
    ...dayRows.map(d => Number(d.avgJph || 0)),
    Number(summary.weightedTarget || 0),
    1
  );

  const ruleDateOptions = getWeekDateList().map(d => {
    const key = formatDateKey(d);
    return `<option value="${escapeHtml(key)}">${escapeHtml(getDayShort(d))} — ${escapeHtml(formatDatePretty(d))}</option>`;
  }).join('');

  const departmentForRule = summary.departments || STATE.activeDept || 'ALL';

  return `
    <div class="side-title">
      <div>
        <h3>${escapeHtml(summary.associate || 'Unknown Operator')}</h3>
        <p>
          ${escapeHtml(summary.departments || STATE.activeDept || '')}
          • ${formatScheduleType(summary.scheduleType)} schedule
          • ${summary.scheduledDaysWorked || 0} of ${summary.eligibleWorkDays || 0} eligible days worked
          • ${(summary.areas || []).length} area${(summary.areas || []).length === 1 ? '' : 's'}
        </p>
      </div>
      ${statusBadge(summary.status || 'NO TARGET')}
    </div>

    <div class="side-metrics">
      <div class="side-metric">
        <span>Total Jobs</span>
        <strong>${formatNumber(summary.totalJobs || 0)}</strong>
      </div>

      <div class="side-metric">
        <span>Total Hours</span>
        <strong>${formatDecimal(summary.totalHours || 0)}</strong>
      </div>

      <div class="side-metric">
        <span>Avg JPH</span>
        <strong>${formatDecimal(summary.avgJph || 0)}</strong>
      </div>

      <div class="side-metric">
        <span>Work Produced</span>
        <strong>${Number(summary.workProduced || 0) > 0 ? formatDecimal(summary.workProduced) : 'N/A'}</strong>
      </div>

      <div class="side-metric">
        <span>Schedule</span>
        <strong>${formatScheduleType(summary.scheduleType)}</strong>
      </div>

      <div class="side-metric">
        <span>Eligible Days</span>
        <strong>${summary.scheduledDaysWorked || 0}/${summary.eligibleWorkDays || 0}</strong>
      </div>

      <div class="side-metric">
        <span>Scheduled</span>
        <strong>${(summary.scheduledDays || []).join(', ') || '—'}</strong>
      </div>

      <div class="side-metric">
        <span>Excluded Days</span>
        <strong>${summary.excludedDays || 0}</strong>
      </div>
    </div>

    <div class="associate-rule-panel">
      <div class="section-head">
        <div>
          <h3>Associate Day Rule</h3>
          <p>Mark this associate/date as Holiday, No OT, or Excused. This changes weekly eligible days only for this associate.</p>
        </div>
      </div>

      <div class="rule-form">
        <label>
          <span>Date</span>
          <select id="associateRuleDate">${ruleDateOptions}</select>
        </label>

        <label>
          <span>Type</span>
          <select id="associateRuleType" onchange="syncAssociateRuleLabel()">
            <option value="NO_OT">NO_OT</option>
            <option value="HOLIDAY">HOLIDAY</option>
            <option value="EXCUSED">EXCUSED</option>
          </select>
        </label>

        <label>
          <span>Label</span>
          <input id="associateRuleLabel" type="text" placeholder="No OT Scheduled" />
        </label>

        <label>
          <span>Supervisor Note</span>
          <textarea id="associateRuleNote" rows="3" placeholder="Optional note for this associate/date..."></textarea>
        </label>

        <button class="rule-save-btn" type="button" onclick="saveAssociateDayRuleFromProfile('${escapeJs(summary.associate)}', '${escapeJs(departmentForRule)}')">
          Save Day Rule
        </button>
      </div>

      ${renderAssociateRuleList(summary)}
    </div>

    <div class="chart-panel">
      <div class="section-head">
        <div>
          <h3>Weekly Trend</h3>
          <p>Actual JPH by day</p>
        </div>
      </div>

      <div class="bar-chart">
        ${
          dayRows.length
            ? dayRows.map(d => {
                const height = Math.max(4, (Number(d.avgJph || 0) / maxJph) * 150);
                return `
                  <div class="bar-item">
                    <div class="bar-value">${formatDecimal(d.avgJph || 0)}</div>
                    <div class="bar" style="height:${height}px"></div>
                    <div class="bar-label">${getDayShort(d.workDate)}</div>
                  </div>
                `;
              }).join('')
            : `<p style="color:#64748b;">No weekly trend available.</p>`
        }
      </div>
    </div>
  `;
}

function syncAssociateRuleLabel() {
  const typeEl = document.getElementById('associateRuleType');
  const labelEl = document.getElementById('associateRuleLabel');
  if (!typeEl || !labelEl) return;
  if (!labelEl.value || ['No OT Scheduled', 'Holiday', 'Excused Day'].includes(labelEl.value)) {
    labelEl.value = getDefaultRuleLabel(typeEl.value);
  }
}

/************************************************************
 * FINAL PATCH — BREAKAGE SUMMARY FIRST + DATE DETAILS ON CLICK
 *
 * Opening / switching the Breakage Hub uses ?action=breakageSummary.
 * A selected date uses ?action=breakageDateDetails&workDate=YYYY-MM-DD.
 ************************************************************/

if (!STATE.breakageDateDetailsCache) STATE.breakageDateDetailsCache = {};

function getBreakageWeekKeyForSelectedWeek() {
  const start = String(STATE.selectedWeekStartDate || STATE.meta.weekStartDate || '').trim();
  const end = String(STATE.selectedWeekEndDate || STATE.meta.weekEndDate || '').trim();
  return `${start}|${end}`;
}

function getBreakageDateDetailsCacheKey(dateKey) {
  return `${getBreakageWeekKeyForSelectedWeek()}|${String(dateKey || '').trim()}`;
}

function getBreakageSummaryRowsForSelectedWeek() {
  const startDate = STATE.selectedWeekStartDate || STATE.meta.weekStartDate;
  const endDate = STATE.selectedWeekEndDate || STATE.meta.weekEndDate;
  const start = parseLocalDate(startDate);
  const end = parseLocalDate(endDate);

  const rows = dedupeBreakageRows(STATE.breakageDailySnapshot || []);

  return rows.filter(r => {
    const d = parseLocalDate(r.WorkDate);
    const dateMatch = d && start && end && d >= start && d <= end;
    if (!dateMatch) return false;

    const blob = [r.OperatorName, r.AccessPoint, r.BreakageReason].join(' ').toLowerCase();
    const searchMatch = !STATE.search || blob.includes(STATE.search);

    return searchMatch;
  });
}

function getBreakageRowsForSelectedWeek() {
  // In summary-first mode, this returns weekly summary/date-card rows.
  return getBreakageSummaryRowsForSelectedWeek();
}

function fetchBreakageFallbackForWeek(week, reason = 'summary') {
  if (!week || !week.weekStartDate || !week.weekEndDate) return Promise.resolve();

  const existingRows = getBreakageSummaryRowsForSelectedWeek();
  if (existingRows.length) return Promise.resolve({ source: 'existing-summary' });

  showRefreshStatus(`Loading breakage summary for ${week.fiscalWeek || 'selected week'}...`);
  startLoadTimer('breakage');

  if (STATE.activePage === 'breakage') {
    const page = document.getElementById('pageContent');
    if (page) {
      page.innerHTML = `
        <section class="quality-hub-shell">
          <div class="loading-card command-loading-card">
            <div class="spinner"></div>
            <strong>Loading Quality / Breakage summary...</strong>
            <small>Summary first. Date details load only when selected.</small>
          </div>
        </section>
      `;
    }
  }

  const qs = new URLSearchParams({
    action: 'breakageSummary',
    weekStartDate: week.weekStartDate,
    weekEndDate: week.weekEndDate,
    ts: String(Date.now())
  });

  return fetch(`${API_URL}?${qs.toString()}`)
    .then(r => r.text())
    .then(text => {
      let data;
      try {
        data = JSON.parse(text);
      } catch (err) {
        throw new Error('API did not return JSON. Redeploy Apps Script and test /exec?action=breakageSummary.');
      }

      if (!data || data.ok === false) {
        throw new Error(data && data.message ? data.message : 'Invalid breakageSummary response.');
      }

      const summaryRows = normalizeBreakageRows(data.breakageSummaryRows || data.breakageDailySnapshot || []);

      STATE.breakageDailySnapshot = dedupeBreakageRows([
        ...(STATE.breakageDailySnapshot || []),
        ...summaryRows
      ]);

      const weekKey = week.key || `${week.weekStartDate}|${week.weekEndDate}`;
      STATE.weekDataCache[weekKey] = {
        ...(STATE.weekDataCache[weekKey] || {}),
        breakageSummaryRows: summaryRows,
        breakageDailySnapshot: summaryRows
      };

      renderEverything();
      endLoadTimer('breakage', getApiTimingMeta(data));
      return data;
    })
    .catch(showError);
}

function fetchBreakageDateDetails(dateKey) {
  const key = getBreakageDateDetailsCacheKey(dateKey);
  const cached = STATE.breakageDateDetailsCache[key];

  if (cached) {
    renderEverything();
    return Promise.resolve(cached);
  }

  showRefreshStatus(`Loading breakage records for ${formatDatePretty(dateKey)}...`);

  const panel = document.getElementById('breakageRecordsPanel');
  if (panel) {
    panel.innerHTML = `
      <div class="section-head compact">
        <div>
          <h3>Breakage Records - ${formatDatePretty(dateKey)}</h3>
          <p>Loading selected date records only...</p>
        </div>
      </div>
      <div class="loading-card command-loading-card">
        <div class="spinner"></div>
        <strong>Loading date details...</strong>
      </div>
    `;
  }

  const qs = new URLSearchParams({
    action: 'breakageDateDetails',
    workDate: dateKey,
    ts: String(Date.now())
  });

  return fetch(`${API_URL}?${qs.toString()}`)
    .then(r => r.text())
    .then(text => {
      let data;
      try {
        data = JSON.parse(text);
      } catch (err) {
        throw new Error('API did not return JSON. Redeploy Apps Script and test /exec?action=breakageDateDetails.');
      }

      if (!data || data.ok === false) {
        throw new Error(data && data.message ? data.message : 'Invalid breakageDateDetails response.');
      }

      const rows = normalizeBreakageRows(data.breakageDateDetails || data.breakageDailySnapshot || []);
      STATE.breakageDateDetailsCache[key] = rows;

      renderEverything();
      showRefreshStatus(`Loaded ${rows.length} breakage records for ${formatDatePretty(dateKey)}`);
      return rows;
    })
    .catch(showError);
}

function setBreakageDateFilter(dateKey) {
  STATE.selectedBreakageDate = String(dateKey || 'ALL');

  if (STATE.selectedBreakageDate !== 'ALL') {
    renderEverything();
    fetchBreakageDateDetails(STATE.selectedBreakageDate);
  } else {
    renderEverything();
  }

  setTimeout(() => {
    const panel = document.getElementById('breakageRecordsPanel');
    if (panel) panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 50);
}

function getSelectedBreakageRows(rows) {
  const selected = String(STATE.selectedBreakageDate || 'ALL');
  if (selected === 'ALL') return rows || [];

  const key = getBreakageDateDetailsCacheKey(selected);
  const details = STATE.breakageDateDetailsCache[key];

  if (details) {
    const activeAP = String(STATE.activeStationGroup || 'ALL').trim();
    return details.filter(r => {
      const apMatch = activeAP === 'ALL' || normalizeQualityText(r.AccessPoint) === normalizeQualityText(activeAP);
      const blob = [r.OperatorName, r.AccessPoint, r.BreakageReason].join(' ').toLowerCase();
      const searchMatch = !STATE.search || blob.includes(STATE.search);
      return apMatch && searchMatch;
    });
  }

  // Until details arrive, use the summary row for totals.
  return (rows || []).filter(r => formatDateKey(r.WorkDate) === selected);
}

function renderBreakageRecordsTable(rows) {
  const selected = String(STATE.selectedBreakageDate || 'ALL');

  if (selected === 'ALL') {
    const dailyCards = buildBreakageDateCards(getBreakageSummaryRowsForSelectedWeek());

    return `
      <section class="quality-panel command-card organized-breakage-records" id="breakageRecordsPanel">
        <div class="section-head compact">
          <div>
            <h3>Breakage Daily Summary</h3>
            <p>Full week view is summarized by date. Click a date card above to load only that date's records.</p>
          </div>
          <span class="scoreboard-count">${dailyCards.length} day${dailyCards.length === 1 ? '' : 's'}</span>
        </div>

        <div class="breakage-daily-summary-grid">
          ${dailyCards.length ? dailyCards.map(card => `
            <button class="breakage-daily-summary-card" type="button" onclick="setBreakageDateFilter('${escapeJs(card.key)}')">
              <div>
                <span>${formatDatePretty(card.workDate)}</span>
                <strong>${formatNumber(card.total)} total</strong>
              </div>
              <div class="daily-breakage-mini">
                <b>${formatNumber(card.lens)}</b><small>Lens</small>
                <b>${formatNumber(card.frame)}</b><small>Frame</small>
                <b>${formatNumber(card.records)}</b><small>Summary</small>
              </div>
            </button>
          `).join('') : '<div class="empty-hub-note">No breakage summary found for this week.</div>'}
        </div>
      </section>
    `;
  }

  const detailKey = getBreakageDateDetailsCacheKey(selected);
  const hasDetails = !!STATE.breakageDateDetailsCache[detailKey];

  if (!hasDetails) {
    const summaryRows = (getBreakageSummaryRowsForSelectedWeek() || []).filter(r => formatDateKey(r.WorkDate) === selected);
    const summary = buildBreakageSummary(summaryRows);

    return `
      <section class="quality-panel command-card organized-breakage-records" id="breakageRecordsPanel">
        <div class="section-head compact">
          <div>
            <h3>Breakage Records - ${formatDatePretty(selected)}</h3>
            <p>Summary loaded. Detail records are loading only for this date.</p>
          </div>
          <div class="record-summary-pills">
            <span>${formatNumber(summary.total)} total</span>
            <span>${formatNumber(summary.lens)} lens</span>
            <span>${formatNumber(summary.frame)} frame</span>
          </div>
        </div>
        <div class="loading-card command-loading-card">
          <div class="spinner"></div>
          <strong>Loading selected date details...</strong>
          <small>This avoids loading the full week of records.</small>
        </div>
      </section>
    `;
  }

  const sorted = (rows || []).slice().sort((a, b) => {
    const d = parseLocalDate(a.WorkDate) - parseLocalDate(b.WorkDate);
    if (d !== 0) return d;
    return Number(b.TotalBreakageCount || 0) - Number(a.TotalBreakageCount || 0);
  });

  const title = sorted.length
    ? `Breakage Records - ${formatDatePretty(sorted[0].WorkDate)}`
    : `Breakage Records - ${formatDatePretty(selected)}`;

  const totalLens = sorted.reduce((sum, r) => sum + Number(r.LensBreakageCount || 0), 0);
  const totalFrame = sorted.reduce((sum, r) => sum + Number(r.FrameBreakageCount || 0), 0);
  const totalBreakage = sorted.reduce((sum, r) => sum + Number(r.TotalBreakageCount || 0), 0);

  return `
    <section class="quality-panel command-card organized-breakage-records" id="breakageRecordsPanel">
      <div class="section-head compact">
        <div>
          <h3>${escapeHtml(title)}</h3>
          <p>Clean detail view for the selected date.</p>
        </div>
        <div class="record-summary-pills">
          <span>${formatNumber(totalBreakage)} total</span>
          <span>${formatNumber(totalLens)} lens</span>
          <span>${formatNumber(totalFrame)} frame</span>
          <span>${sorted.length} records</span>
        </div>
      </div>

      ${sorted.length ? `
        <div class="breakage-record-card-grid">
          ${sorted.map(r => `
            <article class="breakage-record-card">
              <div class="record-total-badge">
                <strong>${formatNumber(r.TotalBreakageCount)}</strong>
                <span>Total</span>
              </div>

              <div class="record-main">
                <div class="record-title-row">
                  <strong>${escapeHtml(r.AccessPoint)}</strong>
                  <span>${formatDatePretty(r.WorkDate)}</span>
                </div>

                <div class="record-reason">${escapeHtml(r.BreakageReason)}</div>
                <div class="record-operator">${escapeHtml(r.OperatorName)}</div>

                <div class="record-count-row">
                  <span><b>${formatNumber(r.LensBreakageCount)}</b> Lens</span>
                  <span><b>${formatNumber(r.FrameBreakageCount)}</b> Frame</span>
                </div>
              </div>
            </article>
          `).join('')}
        </div>
      ` : `
        <div class="empty-hub-note">No breakage records found for this date.</div>
      `}
    </section>
  `;
}

/************************************************************
 * FINAL PATCH — PRODUCTIVITY LOGIN, LOCKS, CONFIRM, NOTIFICATIONS
 ************************************************************/

const PRODUCTIVITY_ACCESS_USERS = {
  BLOPEZ: { username: 'BLOPEZ', role: 'LMS', canEdit: true, canConfirm: true, canRelease: true },
  JBOOMERSHINE: { username: 'JBOOMERSHINE', role: 'LMS', canEdit: true, canConfirm: true, canRelease: true },
  BKARR: { username: 'BKARR', role: 'Production Manager', canEdit: true, canConfirm: true, canRelease: false },
  RTATE: { username: 'RTATE', role: 'SR Director', canEdit: true, canConfirm: true, canRelease: false },
  SANDERSON: { username: 'SANDERSON', role: 'Sr Distribution & Inventory Manager', canEdit: true, canConfirm: true, canRelease: false },
  AIVANOVSKI: { username: 'AIVANOVSKI', role: 'Distribution & Inventory Supervisor', canEdit: true, canConfirm: true, canRelease: false },
  CBAYLIS: { username: 'CBAYLIS', role: 'Distribution & Inventory Supervisor', canEdit: true, canConfirm: true, canRelease: false },
  YKEEBEE: { username: 'YKEEBEE', role: 'Distribution & Inventory Supervisor', canEdit: true, canConfirm: true, canRelease: false },
  KMANACK: { username: 'KMANACK', role: 'Production Supervisor', canEdit: true, canConfirm: true, canRelease: false },
  BHECK: { username: 'BHECK', role: 'Production Supervisor', canEdit: true, canConfirm: true, canRelease: false },
  BHONICKER: { username: 'BHONICKER', role: 'Production Supervisor', canEdit: true, canConfirm: true, canRelease: false },
  NPOSTON: { username: 'NPOSTON', role: 'Production Supervisor', canEdit: true, canConfirm: true, canRelease: false },
  BDADE: { username: 'BDADE', role: 'Trainer Coordinator', canEdit: true, canConfirm: true, canRelease: false },
  PTOWNSEND: { username: 'PTOWNSEND', role: 'Trainer Coordinator', canEdit: true, canConfirm: true, canRelease: false }
};

const PRODUCTIVITY_USER_CACHE_KEY = 'PRODUCTIVITY_HUB_CURRENT_USER_V1';

function normalizeProductivityUsername(value) {
  return String(value || '').trim().toUpperCase();
}

function getCurrentProductivityUser() {
  if (STATE.currentUser && STATE.currentUser.username) return STATE.currentUser;

  try {
    const stored = localStorage.getItem(PRODUCTIVITY_USER_CACHE_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      const key = normalizeProductivityUsername(parsed.username);
      if (PRODUCTIVITY_ACCESS_USERS[key]) {
        STATE.currentUser = PRODUCTIVITY_ACCESS_USERS[key];
        return STATE.currentUser;
      }
    }
  } catch (err) {}

  return null;
}

function getCurrentProductivityUsername() {
  const user = getCurrentProductivityUser();
  return user ? user.username : '';
}

function canEditProductivity() {
  const user = getCurrentProductivityUser();
  return !!(user && user.canEdit);
}

function canConfirmProductivity() {
  const user = getCurrentProductivityUser();
  return !!(user && user.canConfirm);
}

function canReleaseProductivityLock() {
  const user = getCurrentProductivityUser();
  return !!(user && user.canRelease);
}

function showProductivityLoginModal(force = false) {
  if (!force && getCurrentProductivityUser()) {
    renderProductivityUserBadge();
    return;
  }

  const existing = document.getElementById('productivityLoginModal');
  if (existing) existing.remove();

  const modal = document.createElement('div');
  modal.id = 'productivityLoginModal';
  modal.className = 'modal-backdrop productivity-login-backdrop';
  modal.innerHTML = `
    <div class="edit-modal productivity-login-modal">
      <div class="modal-head">
        <div>
          <span class="modal-eyebrow">Authorized Productivity Access</span>
          <h3>Select Login</h3>
          <p>Only approved LMS, management, supervisor, and training logins can edit or confirm productivity.</p>
        </div>
      </div>

      <label class="modal-field">
        <span>Login</span>
        <select id="productivityLoginSelect">
          <option value="">Select your login...</option>
          ${Object.values(PRODUCTIVITY_ACCESS_USERS).map(user => `
            <option value="${escapeHtml(user.username)}">${escapeHtml(user.username)} — ${escapeHtml(user.role)}</option>
          `).join('')}
        </select>
      </label>

      <div class="modal-actions hours-modal-actions">
        <button type="button" class="secondary-btn" onclick="closeProductivityLoginModal()">View Only</button>
        <button type="button" class="primary-btn" onclick="saveProductivityLogin()">Continue</button>
      </div>
    </div>
  `;

  document.body.appendChild(modal);
}

function closeProductivityLoginModal() {
  const modal = document.getElementById('productivityLoginModal');
  if (modal) modal.remove();
  renderProductivityUserBadge();
  renderEverything();
}

function saveProductivityLogin() {
  const select = document.getElementById('productivityLoginSelect');
  const username = normalizeProductivityUsername(select ? select.value : '');
  const user = PRODUCTIVITY_ACCESS_USERS[username];

  if (!user) {
    alert('Select an authorized login.');
    return;
  }

  STATE.currentUser = user;
  localStorage.setItem(PRODUCTIVITY_USER_CACHE_KEY, JSON.stringify(user));

  closeProductivityLoginModal();
  loadProductivityLocksQuietly(true);
}

function switchProductivityLogin() {
  localStorage.removeItem(PRODUCTIVITY_USER_CACHE_KEY);
  STATE.currentUser = null;
  showProductivityLoginModal(true);
}

function renderProductivityUserBadge() {
  let badge = document.getElementById('productivityUserBadge');

  if (!badge) {
    const footer = document.querySelector('.sidebar-footer');
    badge = document.createElement('div');
    badge.id = 'productivityUserBadge';
    badge.className = 'productivity-user-badge';

    if (footer && footer.parentNode) {
      footer.parentNode.insertBefore(badge, footer);
    } else {
      document.body.appendChild(badge);
    }
  }

  const user = getCurrentProductivityUser();

  if (!user) {
    badge.innerHTML = `
      <span>Access</span>
      <strong>View Only</strong>
      <button type="button" onclick="showProductivityLoginModal(true)">Login</button>
    `;
    return;
  }

  badge.innerHTML = `
    <span>Logged In</span>
    <strong>${escapeHtml(user.username)}</strong>
    <small>${escapeHtml(user.role)}</small>
    <button type="button" onclick="switchProductivityLogin()">Switch</button>
  `;
}

function getProductivityLockKey(workDate, finalDepartment, operatorName) {
  return [
    formatDateKey(workDate),
    String(finalDepartment || '').trim().toUpperCase(),
    normalizeQualityText(operatorName)
  ].join('|');
}

function normalizeProductivityLocks(rows) {
  return (rows || []).map(row => ({
    ...row,
    WorkDate: row.WorkDate || row.workDate || row.Date || '',
    FinalDepartment: String(row.FinalDepartment || row.finalDepartment || '').trim().toUpperCase(),
    OperatorName: String(row.OperatorName || row.operatorName || '').trim(),
    LockStatus: String(row.LockStatus || row.lockStatus || '').trim().toUpperCase(),
    LockedBy: String(row.LockedBy || row.lockedBy || '').trim(),
    LockedAt: row.LockedAt || row.lockedAt || '',
    LastActionAt: row.LastActionAt || row.lastActionAt || ''
  }));
}

function getProductivityLockForDay(workDate, finalDepartment, operatorName) {
  const key = getProductivityLockKey(workDate, finalDepartment, operatorName);
  return (STATE.productivityLocks || []).find(lock => {
    return getProductivityLockKey(lock.WorkDate, lock.FinalDepartment, lock.OperatorName) === key &&
      String(lock.LockStatus || '').toUpperCase() === 'LOCKED';
  }) || null;
}

function loadProductivityLocksQuietly(renderAfter = false) {
  const qs = new URLSearchParams({
    action: 'productivityLocks',
    ts: String(Date.now())
  });

  return fetch(`${API_URL}?${qs.toString()}`)
    .then(r => r.text())
    .then(text => {
      const data = JSON.parse(text);
      if (!data || data.ok === false) throw new Error(data && data.message ? data.message : 'Invalid productivityLocks response.');

      STATE.productivityLocks = normalizeProductivityLocks(data.locks || []);
      STATE.productivityActivity = data.activity || [];

      renderProductivityUserBadge();
      showProductivityActivityToast(data.activity || []);

      if (renderAfter) renderEverything();

      return data;
    })
    .catch(err => console.warn('[Productivity Hub] Could not load locks:', err));
}

function showProductivityActivityToast(activityRows) {
  const rows = (activityRows || []).slice();
  if (!rows.length) return;

  const latest = rows[0];
  const latestId = String(latest.ActivityID || latest.CreatedAt || '');
  if (!latestId || STATE.productivityActivityLastSeen === latestId) return;

  STATE.productivityActivityLastSeen = latestId;

  const existing = document.getElementById('productivityActivityToast');
  if (existing) existing.remove();

  const toast = document.createElement('div');
  toast.id = 'productivityActivityToast';
  toast.className = 'productivity-activity-toast';
  toast.innerHTML = `
    <button type="button" onclick="this.parentElement.remove()">×</button>
    <span>Productivity Update</span>
    <strong>${escapeHtml(latest.Message || 'Productivity activity was updated.')}</strong>
    <small>${escapeHtml(latest.PerformedBy || '')} ${latest.CreatedAt ? '• ' + formatDatePretty(latest.CreatedAt) : ''}</small>
  `;

  document.body.appendChild(toast);

  setTimeout(() => {
    const el = document.getElementById('productivityActivityToast');
    if (el) el.remove();
  }, 8500);
}

function confirmProductivityDay(dateKey, associateName, finalDepartment) {
  const user = getCurrentProductivityUser();

  if (!user || !user.canConfirm) {
    alert('You are not authorized to confirm productivity days.');
    showProductivityLoginModal(true);
    return;
  }

  const associate = String(associateName || STATE.selectedAssociate || '').trim();
  const dept = String(finalDepartment || STATE.activeDept || '').trim().toUpperCase();

  if (!associate || !dateKey || !dept) {
    alert('Missing associate, date, or department for confirmation.');
    return;
  }

  const ok = confirm(`Confirm and lock productivity for ${associate} on ${formatDatePretty(dateKey)}?`);
  if (!ok) return;

  const qs = new URLSearchParams({
    action: 'confirmProductivityDay',
    workDate: dateKey,
    finalDepartment: dept,
    operatorName: associate,
    username: user.username,
    ts: String(Date.now())
  });

  fetch(`${API_URL}?${qs.toString()}`)
    .then(r => r.text())
    .then(text => {
      const data = JSON.parse(text);
      if (!data.ok) throw new Error(data.message || 'Could not confirm productivity day.');
      showRefreshStatus(`Confirmed ${associate} • ${formatDatePretty(dateKey)}`);
      return loadProductivityLocksQuietly(true);
    })
    .catch(err => alert(err.message || err));
}

function releaseProductivityDay(dateKey, associateName, finalDepartment) {
  const user = getCurrentProductivityUser();

  if (!user || !user.canRelease) {
    alert('Only BLOPEZ and JBOOMERSHINE can release confirmed productivity days.');
    return;
  }

  const associate = String(associateName || STATE.selectedAssociate || '').trim();
  const dept = String(finalDepartment || STATE.activeDept || '').trim().toUpperCase();
  const reason = prompt(`Release productivity lock for ${associate} on ${formatDatePretty(dateKey)}? Enter release reason:`, 'Correction required');

  if (reason === null) return;

  const qs = new URLSearchParams({
    action: 'releaseProductivityDay',
    workDate: dateKey,
    finalDepartment: dept,
    operatorName: associate,
    username: user.username,
    releaseReason: reason || 'Released by LMS',
    ts: String(Date.now())
  });

  fetch(`${API_URL}?${qs.toString()}`)
    .then(r => r.text())
    .then(text => {
      const data = JSON.parse(text);
      if (!data.ok) throw new Error(data.message || 'Could not release productivity day.');
      showRefreshStatus(`Released lock for ${associate} • ${formatDatePretty(dateKey)}`);
      return loadProductivityLocksQuietly(true);
    })
    .catch(err => alert(err.message || err));
}

// Protect the edit picker with access + confirmed-lock rules.
var __openAssociateDayAdjustmentBase = openAssociateDayAdjustment;
openAssociateDayAdjustment = function(dateKey, associateName) {
  const associate = String(associateName || STATE.selectedAssociate || '').trim();

  if (!canEditProductivity()) {
    showProductivityLoginModal(true);
    return;
  }

  const dayRows = getWeekRowsForActiveDept().filter(row =>
    normalizeQualityText(row.OperatorName) === normalizeQualityText(associate) &&
    formatDateKey(row.WorkDate) === String(dateKey || '').trim()
  );

  const firstRow = dayRows[0];
  if (firstRow) {
    const lock = getProductivityLockForDay(dateKey, firstRow.FinalDepartment, associate);
    if (lock) {
      alert(`This day is confirmed and locked by ${lock.LockedBy}. BLOPEZ or JBOOMERSHINE must release it before editing.`);
      return;
    }
  }

  return __openAssociateDayAdjustmentBase(dateKey, associateName);
};

// Replace day cards so confirmed days are grey/locked and action buttons are permission-aware.
renderDayCard = function(day, row, summary, dateValue) {
  const dateKey = formatDateKey(dateValue);
  const associate = summary.associate || STATE.selectedAssociate || '';
  const dept = (row && row.finalDepartment) || (row && row.FinalDepartment) || summary.departments || STATE.activeDept || '';
  const lock = getProductivityLockForDay(dateKey, dept, associate);
  const canEdit = canEditProductivity();
  const canConfirm = canConfirmProductivity();
  const canRelease = canReleaseProductivityLock();
  const locked = !!lock;

  if (!row) {
    return `
      <article class="day-card NO-TARGET ${locked ? 'day-card-locked' : ''}">
        <h4>${day}</h4>
        <p>${formatDatePretty(dateValue)} • No record</p>
        <div class="day-stats day-stats-wide">
          <div><span>Hours</span><strong>0.00</strong></div>
          <div><span>Jobs</span><strong>0</strong></div>
          <div><span>Produced</span><strong>0.00</strong></div>
          <div><span>Daily %</span><strong>0%</strong></div>
          <div><span>Status</span><strong>—</strong></div>
        </div>
      </article>
    `;
  }

  const status = getStatusFromPercent(row.productivityPercent, row.workProduced);

  const actionHtml = locked
    ? `
      <div class="day-lock-banner">Confirmed by ${escapeHtml(lock.LockedBy || 'Supervisor')}</div>
      ${canRelease ? `<button class="day-card-action-btn release" type="button" onclick="event.stopPropagation(); releaseProductivityDay('${escapeJs(dateKey)}','${escapeJs(associate)}','${escapeJs(dept)}')">Release Lock</button>` : ''}
    `
    : `
      <div class="day-card-action-row">
        ${canEdit ? `<button class="day-card-action-btn edit" type="button" onclick="event.stopPropagation(); openAssociateDayAdjustment('${escapeJs(dateKey)}','${escapeJs(associate)}')">Edit Time</button>` : `<span class="day-view-only-pill">View Only</span>`}
        ${canConfirm ? `<button class="day-card-action-btn confirm" type="button" onclick="event.stopPropagation(); confirmProductivityDay('${escapeJs(dateKey)}','${escapeJs(associate)}','${escapeJs(dept)}')">Confirm Day</button>` : ''}
      </div>
    `;

  return `
    <article class="day-card ${statusCss(status)} ${locked ? 'day-card-locked' : 'day-card-clickable'}" ${(!locked && canEdit) ? `onclick="openAssociateDayAdjustment('${escapeJs(dateKey)}','${escapeJs(associate)}')"` : ''} title="${locked ? 'Confirmed productivity day' : 'Productivity day'}">
      ${actionHtml}
      <h4>${day}</h4>
      <p>${formatDatePretty(row.workDate)}</p>
      <div class="day-stats day-stats-wide">
        <div><span>Hours</span><strong>${formatDecimal(row.totalHours)}</strong></div>
        <div><span>Jobs</span><strong>${formatNumber(row.totalJobs)}</strong></div>
        <div><span>Produced</span><strong>${formatDecimal(row.workProduced)}</strong></div>
        <div><span>Daily %</span><strong>${formatPercent(row.productivityPercent)}</strong></div>
        <div><span>Status</span><strong>${locked ? 'Confirmed' : statusLabel(status)}</strong></div>
      </div>
    </article>
  `;
};

// Load access context after app boot and keep light notification polling.
document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    renderProductivityUserBadge();
    if (!getCurrentProductivityUser()) showProductivityLoginModal(false);
    loadProductivityLocksQuietly(true);
  }, 800);

  setInterval(() => {
    loadProductivityLocksQuietly(false);
  }, 60000);
});

// Make sure locks reload after the normal data handlers render.
var __handleDataWithLocksBase = handleData;
handleData = function(data, fromCache = false) {
  __handleDataWithLocksBase(data, fromCache);
  loadProductivityLocksQuietly(true);
};

var __loadFiscalWeekDataWithLocksBase = loadFiscalWeekData;
loadFiscalWeekData = function(week) {
  const result = __loadFiscalWeekDataWithLocksBase(week);
  setTimeout(() => loadProductivityLocksQuietly(true), 1200);
  return result;
};

/************************************************************
 * TOP ACTIONS — BACK TO DASHBOARD + LIVE DATE/TIME
 ************************************************************/
function goBackToDashboard() {
  window.location.href = 'index.html';
}

function updateProductivityTopClock() {
  const timeEl = document.getElementById('productivityTopTime');
  const dateEl = document.getElementById('productivityTopDate');
  if (!timeEl && !dateEl) return;

  const now = new Date();

  if (timeEl) {
    timeEl.textContent = now.toLocaleTimeString([], {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });
  }

  if (dateEl) {
    dateEl.textContent = now.toLocaleDateString([], {
      weekday: 'short',
      month: 'short',
      day: '2-digit',
      year: 'numeric'
    });
  }
}

document.addEventListener('DOMContentLoaded', () => {
  updateProductivityTopClock();
  setInterval(updateProductivityTopClock, 1000);
});

/************************************************************
 * JSONP_LOCKS_CORS_FIX_V2_HEADER_HELPER
 ************************************************************/

/************************************************************
 * FINAL PATCH — JSONP LOCK ACTIONS FOR GITHUB CORS
 *
 * Fixes GitHub Pages CORS issue for:
 * - productivityLocks
 * - confirmProductivityDay
 * - releaseProductivityDay
 ************************************************************/

function productivityJsonpRequest(params) {
  return new Promise((resolve, reject) => {
    const callbackName = '__prodHubJsonp_' + Date.now() + '_' + Math.floor(Math.random() * 100000);
    const qs = new URLSearchParams(params || {});
    qs.set('callback', callbackName);
    qs.set('ts', String(Date.now()));

    const script = document.createElement('script');
    const cleanup = () => {
      try { delete window[callbackName]; } catch (err) { window[callbackName] = undefined; }
      if (script && script.parentNode) script.parentNode.removeChild(script);
    };

    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('Request timed out.'));
    }, 25000);

    window[callbackName] = data => {
      clearTimeout(timer);
      cleanup();
      if (!data || data.ok === false) {
        reject(new Error(data && data.message ? data.message : 'Invalid API response.'));
        return;
      }
      resolve(data);
    };

    script.onerror = () => {
      clearTimeout(timer);
      cleanup();
      reject(new Error('Could not load Apps Script JSONP endpoint.'));
    };

    script.src = `${API_URL}?${qs.toString()}`;
    document.head.appendChild(script);
  });
}

// Override lock loader to avoid CORS fetch.
loadProductivityLocksQuietly = function(renderAfter = false) {
  return productivityJsonpRequest({
    action: 'productivityLocks'
  })
    .then(data => {
      STATE.productivityLocks = normalizeProductivityLocks(data.locks || []);
      STATE.productivityActivity = data.activity || [];

      renderProductivityUserBadge();
      showProductivityActivityToast(data.activity || []);

      if (renderAfter) renderEverything();

      return data;
    })
    .catch(err => console.warn('[Productivity Hub] Could not load locks:', err));
};

// Override confirm to avoid CORS fetch.
confirmProductivityDay = function(dateKey, associateName, finalDepartment) {
  const user = getCurrentProductivityUser();

  if (!user || !user.canConfirm) {
    alert('You are not authorized to confirm productivity days.');
    showProductivityLoginModal(true);
    return;
  }

  const associate = String(associateName || STATE.selectedAssociate || '').trim();
  const dept = String(finalDepartment || STATE.activeDept || '').trim().toUpperCase();

  if (!associate || !dateKey || !dept) {
    alert('Missing associate, date, or department for confirmation.');
    return;
  }

  const ok = confirm(`Confirm and lock productivity for ${associate} on ${formatDatePretty(dateKey)}?`);
  if (!ok) return;

  showRefreshStatus(`Confirming ${associate} • ${formatDatePretty(dateKey)}...`);

  return productivityJsonpRequest({
    action: 'confirmProductivityDay',
    workDate: dateKey,
    finalDepartment: dept,
    operatorName: associate,
    username: user.username
  })
    .then(data => {
      showRefreshStatus(`Confirmed ${associate} • ${formatDatePretty(dateKey)}`);
      return loadProductivityLocksQuietly(true);
    })
    .catch(err => alert(err.message || err));
};

// Override release to avoid CORS fetch.
releaseProductivityDay = function(dateKey, associateName, finalDepartment) {
  const user = getCurrentProductivityUser();

  if (!user || !user.canRelease) {
    alert('Only BLOPEZ and JBOOMERSHINE can release confirmed productivity days.');
    return;
  }

  const associate = String(associateName || STATE.selectedAssociate || '').trim();
  const dept = String(finalDepartment || STATE.activeDept || '').trim().toUpperCase();
  const reason = prompt(`Release productivity lock for ${associate} on ${formatDatePretty(dateKey)}? Enter release reason:`, 'Correction required');

  if (reason === null) return;

  showRefreshStatus(`Releasing lock for ${associate} • ${formatDatePretty(dateKey)}...`);

  return productivityJsonpRequest({
    action: 'releaseProductivityDay',
    workDate: dateKey,
    finalDepartment: dept,
    operatorName: associate,
    username: user.username,
    releaseReason: reason || 'Released by LMS'
  })
    .then(data => {
      showRefreshStatus(`Released lock for ${associate} • ${formatDatePretty(dateKey)}`);
      return loadProductivityLocksQuietly(true);
    })
    .catch(err => alert(err.message || err));
};

/************************************************************
 * FINAL PATCH — USE DASHBOARD LOGIN SESSION, NO USER PICKER
 *
 * The Productivity page now follows the main dashboard login:
 * - sessionStorage.lms_user
 * - sessionStorage.lms_role
 * - sessionStorage.lms_subrole
 * - sessionStorage.lms_features
 *
 * This removes the manual user selector so supervisors cannot choose
 * somebody else's login inside Productivity Hub.
 ************************************************************/

function getDashboardSessionUser() {
  const username = normalizeProductivityUsername(sessionStorage.getItem('lms_user') || '');
  if (!username) return null;

  let features = {};
  try {
    features = JSON.parse(sessionStorage.getItem('lms_features') || '{}') || {};
  } catch (err) {
    features = {};
  }

  const base = PRODUCTIVITY_ACCESS_USERS[username] || {
    username,
    role: sessionStorage.getItem('lms_role') || 'View Only',
    canEdit: false,
    canConfirm: false,
    canRelease: false
  };

  const role = sessionStorage.getItem('lms_role') || base.role || '';
  const subRole = sessionStorage.getItem('lms_subrole') || '';
  const fullName = sessionStorage.getItem('lms_fullname') || username;

  return {
    ...base,
    username,
    role,
    subRole,
    fullName,
    features,
    canView: features.Productivity === true || features.Productivity_ViewHub === true || !!PRODUCTIVITY_ACCESS_USERS[username],
    canEdit: base.canEdit === true && features.Productivity_EditTime === true,
    canConfirm: base.canConfirm === true && features.Productivity_ConfirmDay === true,
    canRelease: base.canRelease === true && features.Productivity_ReleaseLock === true,
    canAdmin: features.Productivity_AdminSettings === true
  };
}

getCurrentProductivityUser = function() {
  const sessionUser = getDashboardSessionUser();
  STATE.currentUser = sessionUser;
  return STATE.currentUser;
};

getCurrentProductivityUsername = function() {
  const user = getCurrentProductivityUser();
  return user ? user.username : '';
};

canEditProductivity = function() {
  const user = getCurrentProductivityUser();
  return !!(user && user.canEdit);
};

canConfirmProductivity = function() {
  const user = getCurrentProductivityUser();
  return !!(user && user.canConfirm);
};

canReleaseProductivityLock = function() {
  const user = getCurrentProductivityUser();
  return !!(user && user.canRelease);
};

showProductivityLoginModal = function(force = false) {
  const user = getCurrentProductivityUser();

  if (!user) {
    alert('Your dashboard session is missing. Please log in again.');
    window.location.href = 'login.html';
    return;
  }

  renderProductivityUserBadge();
};

closeProductivityLoginModal = function() {
  const modal = document.getElementById('productivityLoginModal');
  if (modal) modal.remove();
  renderProductivityUserBadge();
};

saveProductivityLogin = function() {
  renderProductivityUserBadge();
};

switchProductivityLogin = function() {
  const ok = confirm('This will log you out and return to the Dashboard Login page. Continue?');
  if (!ok) return;

  sessionStorage.clear();
  localStorage.removeItem(PRODUCTIVITY_USER_CACHE_KEY);
  window.location.href = 'login.html';
};

renderProductivityUserBadge = function() {
  let badge = document.getElementById('productivityUserBadge');

  if (!badge) {
    const footer = document.querySelector('.sidebar-footer');
    badge = document.createElement('div');
    badge.id = 'productivityUserBadge';
    badge.className = 'productivity-user-badge';

    if (footer && footer.parentNode) {
      footer.parentNode.insertBefore(badge, footer);
    } else {
      document.body.appendChild(badge);
    }
  }

  const user = getCurrentProductivityUser();

  if (!user) {
    badge.innerHTML = `
      <span>Access</span>
      <strong>No Session</strong>
      <button type="button" onclick="window.location.href='login.html'">Login</button>
    `;
    return;
  }

  const mode = user.canEdit || user.canConfirm || user.canRelease ? 'Authorized' : 'View Only';

  badge.innerHTML = `
    <span>Logged In</span>
    <strong>${escapeHtml(user.username)}</strong>
    <small>${escapeHtml(user.role || '')}${user.subRole ? ' • ' + escapeHtml(user.subRole) : ''}</small>
    <em class="access-mode-pill">${escapeHtml(mode)}</em>
    <button type="button" onclick="switchProductivityLogin()">Logout</button>
  `;
};

openAssociateDayAdjustment = function(dateKey, associateName) {
  const associate = String(associateName || STATE.selectedAssociate || '').trim();

  if (!canEditProductivity()) {
    alert('Your login is view-only for productivity time edits.');
    return;
  }

  const dayRows = getWeekRowsForActiveDept().filter(row =>
    normalizeQualityText(row.OperatorName) === normalizeQualityText(associate) &&
    formatDateKey(row.WorkDate) === String(dateKey || '').trim()
  );

  if (!dayRows.length) {
    alert('No editable record found for this date.');
    return;
  }

  const firstRow = dayRows[0];
  const lock = getProductivityLockForDay(dateKey, firstRow.FinalDepartment, associate);
  if (lock) {
    alert(`This day is confirmed and locked by ${lock.LockedBy}. BLOPEZ or JBOOMERSHINE must release it before editing.`);
    return;
  }

  if (dayRows.length === 1) {
    openHoursNoteEditor(getRecordKey(dayRows[0]));
    return;
  }

  const existing = document.getElementById('dayAdjustmentPickerModal');
  if (existing) existing.remove();

  const modal = document.createElement('div');
  modal.id = 'dayAdjustmentPickerModal';
  modal.className = 'modal-backdrop day-adjustment-modal';
  modal.innerHTML = `
    <div class="edit-modal day-adjustment-picker">
      <div class="modal-head">
        <div>
          <h3>Select Record to Adjust</h3>
          <p>${escapeHtml(associate)} • ${formatDatePretty(dateKey)}</p>
        </div>
        <button type="button" class="modal-x" onclick="closeAssociateDayAdjustment()">×</button>
      </div>

      <div class="day-adjustment-list">
        ${dayRows.map(row => {
          const recordKey = getRecordKey(row);
          const status = getStatusFromPercent(row.ProductivityPercent, row.WorkProduced);
          return `
            <button type="button" class="day-adjustment-option ${statusCss(status)}" onclick="closeAssociateDayAdjustment(); openHoursNoteEditor('${escapeJs(recordKey)}')">
              <div>
                <strong>${escapeHtml(row.StationGroup || row.AccessPoint || row.FinalDepartment)}</strong>
                <span>${escapeHtml(row.FinalDepartment)} • ${escapeHtml(row.AccessPoint)}</span>
              </div>
              <div class="day-adjustment-option-metrics">
                <span>${formatDecimal(row.SessionTimeHours)} hrs</span>
                <span>${formatNumber(row.TotalJobScan)} jobs</span>
                <span>${formatPercent(row.ProductivityPercent)}</span>
              </div>
            </button>
          `;
        }).join('')}
      </div>
    </div>
  `;

  document.body.appendChild(modal);
};

document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    const picker = document.getElementById('productivityLoginModal');
    if (picker) picker.remove();

    const user = getCurrentProductivityUser();
    if (!user) {
      alert('Your dashboard session expired. Please log in again.');
      window.location.href = 'login.html';
      return;
    }

    renderProductivityUserBadge();
    renderEverything();
  }, 1200);
});

/************************************************************
 * FINAL PATCH — MOVE LOGIN BADGE UNDER DASHBOARD BUTTON
 ************************************************************/
renderProductivityUserBadge = function() {
  const mount = document.getElementById('productivityUserBadgeMount');
  let badge = document.getElementById('productivityUserBadge');

  if (!badge) {
    badge = document.createElement('div');
    badge.id = 'productivityUserBadge';
    badge.className = 'productivity-user-badge productivity-user-badge-nav';
  }

  if (mount && badge.parentNode !== mount) {
    mount.appendChild(badge);
  } else if (!mount && !badge.parentNode) {
    document.body.appendChild(badge);
  }

  const user = getCurrentProductivityUser();

  if (!user) {
    badge.innerHTML = `
      <span>Access</span>
      <strong>No Session</strong>
      <button type="button" onclick="window.location.href='login.html'">Login</button>
    `;
    return;
  }

  const mode = user.canEdit || user.canConfirm || user.canRelease ? 'Authorized' : 'View Only';

  badge.innerHTML = `
    <span>Logged In</span>
    <strong>${escapeHtml(user.username)}</strong>
    <small>${escapeHtml(user.role || '')}${user.subRole ? ' • ' + escapeHtml(user.subRole) : ''}</small>
    <em class="access-mode-pill">${escapeHtml(mode)}</em>
    <button type="button" onclick="switchProductivityLogin()">Logout</button>
  `;
};

/************************************************************
 * FINAL PATCH — MOVE LOGIN BADGE TO TOP RIGHT HEADER
 ************************************************************/
renderProductivityUserBadge = function() {
  const mount = document.getElementById('productivityUserBadgeMount');
  let badge = document.getElementById('productivityUserBadge');

  if (!badge) {
    badge = document.createElement('div');
    badge.id = 'productivityUserBadge';
    badge.className = 'productivity-user-badge productivity-user-badge-top';
  }

  badge.className = 'productivity-user-badge productivity-user-badge-top';

  if (mount && badge.parentNode !== mount) {
    mount.appendChild(badge);
  } else if (!mount && !badge.parentNode) {
    document.body.appendChild(badge);
  }

  const user = getCurrentProductivityUser();

  if (!user) {
    badge.innerHTML = `
      <div class="top-user-info">
        <span>Access</span>
        <strong>No Session</strong>
      </div>
      <button type="button" onclick="window.location.href='login.html'">Login</button>
    `;
    return;
  }

  const mode = user.canEdit || user.canConfirm || user.canRelease ? 'Authorized' : 'View Only';

  badge.innerHTML = `
    <div class="top-user-avatar">${escapeHtml((user.username || '?').slice(0, 1))}</div>
    <div class="top-user-info">
      <span>Logged In</span>
      <strong>${escapeHtml(user.username)}</strong>
      <small>${escapeHtml(user.role || '')}${user.subRole ? ' • ' + escapeHtml(user.subRole) : ''}</small>
    </div>
    <em class="access-mode-pill">${escapeHtml(mode)}</em>
    <button type="button" onclick="switchProductivityLogin()">Logout</button>
  `;
};

/************************************************************
 * FINAL PATCH — FAST ADJUSTMENT SAVE, NO FULL RELOAD
 *
 * Saves with JSONP to avoid CORS and avoids reloadData() after save.
 * The card/table update immediately from local state.
 ************************************************************/

function applyProductivityAdjustmentLocally(recordKey, adjustedHours, reason, note, response) {
  const rowsToPatch = [
    ...(STATE.master || []),
    ...(STATE.dailySnapshot || []),
    ...(STATE.editSnapshot || [])
  ];

  rowsToPatch.forEach(row => {
    if (getRecordKey(row) !== recordKey) return;

    const originalHours = Number(row.OriginalSessionHours || row.SessionTimeHours || 0);
    const jobs = Number(row.TotalJobScan || 0);
    const target = Number(row.TargetAvgPerHour || 0);
    const workProduced = target > 0 ? jobs / target : Number(row.WorkProduced || 0);
    const avgPerHour = adjustedHours > 0 ? jobs / adjustedHours : 0;
    const productivity = adjustedHours > 0 ? (workProduced / adjustedHours) * 100 : 0;

    row.OriginalSessionHours = originalHours || Number(row.SessionTimeHours || 0);
    row.SessionTimeHours = adjustedHours;
    row.TotalHours = adjustedHours;
    row.AveragePerHour = avgPerHour;
    row.WorkProduced = workProduced;
    row.ProductivityPercent = productivity;
    row.AdjustmentReason = reason;
    row.AdjustmentNote = note;
    row.HasAdjustment = true;
  });

  const saved = response && response.saved ? response.saved : null;
  if (saved) {
    STATE.adjustments = [
      ...(STATE.adjustments || []).filter(adj => getAdjustmentKeyFromAdjustment(adj) !== getAdjustmentKeyFromAdjustment(saved)),
      saved
    ];
  }

  if (response && response.locks) {
    STATE.productivityLocks = normalizeProductivityLocks(response.locks || []);
  }

  if (response && response.activity) {
    STATE.productivityActivity = response.activity || [];
    showProductivityActivityToast(response.activity || []);
  }

  const selectedWeekKey = getSelectedFiscalWeekKey();
  if (selectedWeekKey && STATE.weekDataCache) {
    delete STATE.weekDataCache[selectedWeekKey];
  }

  renderNavBar();
  renderEverything();
}

saveHoursNoteAdjustment = function(recordKey) {
  const row = findRecordByKey(recordKey);
  if (!row) {
    alert('Could not find this record. Refresh the page and try again.');
    return;
  }

  if (!canEditProductivity()) {
    alert('Your login is view-only for productivity time edits.');
    return;
  }

  const lock = getProductivityLockForDay(formatDateKey(row.WorkDate), row.FinalDepartment, row.OperatorName);
  if (lock) {
    alert(`This day is confirmed and locked by ${lock.LockedBy}. BLOPEZ or JBOOMERSHINE must release it before editing.`);
    return;
  }

  const hoursEl = document.getElementById('editAdjustedHours');
  const reasonEl = document.getElementById('editAdjustmentReason');
  const noteEl = document.getElementById('editAdjustmentNote');

  const adjustedHours = Number(hoursEl ? hoursEl.value : row.SessionTimeHours);
  let reason = reasonEl ? reasonEl.value : 'Supervisor Adjustment';
  if (reason === 'Indirect Work') reason = 'Non-Productive Time';
  const note = noteEl ? noteEl.value : '';

  if (isNaN(adjustedHours) || adjustedHours < 0) {
    alert('Adjusted hours must be 0 or higher.');
    return;
  }

  const saveBtn = document.querySelector('#hoursNoteModal .primary-btn');
  if (saveBtn) {
    saveBtn.disabled = true;
    saveBtn.textContent = 'Saving...';
  }

  showRefreshStatus('Saving productivity adjustment...');

  return productivityJsonpRequest({
    action: 'saveProductivityAdjustment',
    workDate: formatDateKey(row.WorkDate),
    finalDepartment: row.FinalDepartment || '',
    stationGroup: row.StationGroup || '',
    department: row.Department || row.FinalDepartment || '',
    accessPoint: row.AccessPoint || '',
    operatorName: row.OperatorName || '',
    originalSessionHours: String(Number(row.OriginalSessionHours || row.SessionTimeHours || 0)),
    adjustedSessionHours: String(adjustedHours),
    totalJobScan: String(Number(row.TotalJobScan || 0)),
    targetAvgPerHour: String(Number(row.TargetAvgPerHour || 0)),
    reason,
    note,
    editedBy: getCurrentProductivityUsername() || 'WEBPAGE'
  })
    .then(response => {
      applyProductivityAdjustmentLocally(recordKey, adjustedHours, reason, note, response);
      closeHoursNoteEditor();
      showRefreshStatus('Adjustment saved. View updated.');

      // Refresh locks/activity quietly, but do not block the UI.
      setTimeout(() => loadProductivityLocksQuietly(false), 800);

      return response;
    })
    .catch(err => {
      const saveBtn = document.querySelector('#hoursNoteModal .primary-btn');
      if (saveBtn) {
        saveBtn.disabled = false;
        saveBtn.textContent = 'Save Adjustment';
      }
      alert('Could not save adjustment: ' + (err.message || err));
    });
};

/************************************************************
 * FINAL PATCH — ASSOCIATE BREAKAGE DETAIL FIX
 *
 * The Quality Hub opens with breakageSummary rows for speed.
 * Those summary rows are NOT associate-level records.
 *
 * For an Associate Profile, load date details only for the days that
 * associate worked, then match by OperatorName. This prevents full
 * facility/date breakage totals from showing on one associate.
 ************************************************************/

function getAssociateWorkedDateKeys(summary) {
  const keys = new Set();
  (summary && summary.rows ? summary.rows : []).forEach(row => {
    const key = formatDateKey(row.WorkDate);
    if (key) keys.add(key);
  });
  return Array.from(keys).sort();
}

function getActualBreakageRowsForDate(dateKey) {
  const detailKey = getBreakageDateDetailsCacheKey(dateKey);
  const detailRows = STATE.breakageDateDetailsCache && STATE.breakageDateDetailsCache[detailKey]
    ? STATE.breakageDateDetailsCache[detailKey]
    : [];

  // Fallback only to rows that are not summary rows.
  const directRows = (STATE.breakageDailySnapshot || []).filter(row => {
    const rowType = String(row.RowType || '').toUpperCase();
    return rowType !== 'SUMMARY' && formatDateKey(row.WorkDate) === dateKey;
  });

  return dedupeBreakageRows([...(detailRows || []), ...directRows]);
}

function getAssociateBreakageRows(summary) {
  const associate = normalizeQualityText(summary && summary.associate);
  if (!associate) return [];

  const dateKeys = getAssociateWorkedDateKeys(summary);
  const rows = [];

  dateKeys.forEach(dateKey => {
    getActualBreakageRowsForDate(dateKey).forEach(row => {
      if (normalizeQualityText(row.OperatorName) === associate) rows.push(row);
    });
  });

  return dedupeBreakageRows(rows);
}

function ensureAssociateBreakageDetails(summary) {
  if (!summary || !summary.rows || !summary.rows.length) return;

  const dateKeys = getAssociateWorkedDateKeys(summary);
  const missing = dateKeys.filter(dateKey => {
    const detailKey = getBreakageDateDetailsCacheKey(dateKey);
    return !(STATE.breakageDateDetailsCache && STATE.breakageDateDetailsCache[detailKey]);
  });

  if (!missing.length) return;

  // Mark to prevent repeated parallel fetches.
  STATE.associateBreakageLoadingDates = STATE.associateBreakageLoadingDates || {};

  missing.forEach(dateKey => {
    if (STATE.associateBreakageLoadingDates[dateKey]) return;
    STATE.associateBreakageLoadingDates[dateKey] = true;

    fetchBreakageDateDetails(dateKey)
      .finally(() => {
        STATE.associateBreakageLoadingDates[dateKey] = false;
      });
  });
}

function hasAllAssociateBreakageDetails(summary) {
  const dateKeys = getAssociateWorkedDateKeys(summary);
  return dateKeys.every(dateKey => {
    const detailKey = getBreakageDateDetailsCacheKey(dateKey);
    return !!(STATE.breakageDateDetailsCache && STATE.breakageDateDetailsCache[detailKey]);
  });
}

// Override profile hero — shows breakage tiles immediately from any cached dates,
// updates live as each remaining date loads in parallel.
renderHero = function(summary) {
  // Fire parallel fetches for all worked dates immediately.
  ensureAssociateBreakageDetailsFast(summary);

  const pct = summary.workProduced > 0 ? Math.min(Math.round(summary.productivityPercent * 100), 180) : 0;

  // Use whatever breakage rows are cached right now — accurate partial counts.
  const breakageRows = getAssociateBreakageRows(summary);
  const quality = buildAssociateBreakageSummary(summary, breakageRows);

  // Build per-tile loading indicator: shows "2/5 days" until complete.
  const dateKeys = getAssociateWorkedDateKeys(summary);
  const loadedCount = dateKeys.filter(dk => {
    const ck = getBreakageDateDetailsCacheKey(dk);
    return !!(STATE.breakageDateDetailsCache && STATE.breakageDateDetailsCache[ck]);
  }).length;
  const totalDates = dateKeys.length;
  const allLoaded = loadedCount >= totalDates && totalDates > 0;

  const lensSmall = allLoaded
    ? formatQualityPercent(quality.lensPercent)
    : (loadedCount > 0 ? `${loadedCount}/${totalDates} days ↻` : '↻ loading');
  const frameSmall = allLoaded
    ? formatQualityPercent(quality.framePercent)
    : (loadedCount > 0 ? `${loadedCount}/${totalDates} days ↻` : '↻ loading');

  return `
    <section class="profile-hero command-card">
      <div class="identity">
        <div class="avatar">👤</div>
        <div>
          <span class="profile-tag">Associate Profile</span>
          <h3>${escapeHtml(summary.associate)}</h3>
          <p>${escapeHtml(summary.departments || STATE.activeDept)} • ${formatScheduleType(summary.scheduleType)} Schedule • ${summary.scheduledDaysWorked} of ${summary.eligibleWorkDays} eligible day${summary.eligibleWorkDays === 1 ? '' : 's'} worked • ${summary.areas.length} area${summary.areas.length === 1 ? '' : 's'}</p>
          ${statusBadge(summary.status)}
        </div>
      </div>

      <div class="ring-wrap">
        <div class="progress-ring" style="--pct:${pct}">
          <strong>${summary.workProduced > 0 ? pct + '%' : 'N/A'}</strong>
        </div>
        <small>Weekly Productivity</small>
      </div>

      <div class="hero-metric">
        <span class="metric-icon">▣</span>
        <div>
          <span>Total Jobs</span>
          <strong data-count="${summary.totalJobs}">${formatNumber(summary.totalJobs)}</strong>
        </div>
      </div>

      <div class="hero-metric">
        <span class="metric-icon">◴</span>
        <div>
          <span>Total Hours</span>
          <strong>${formatDecimal(summary.totalHours)}</strong>
        </div>
      </div>

      <div class="hero-metric">
        <span class="metric-icon">↗</span>
        <div>
          <span>Avg JPH</span>
          <strong>${formatDecimal(summary.avgJph)}</strong>
        </div>
      </div>

      <div class="hero-metric">
        <span class="metric-icon">◎</span>
        <div>
          <span>Work Produced</span>
          <strong>${formatDecimal(summary.workProduced)}</strong>
        </div>
      </div>

      <div class="hero-metric quality-hero-card ${statusCss(quality.status)}">
        <span class="metric-icon">◈</span>
        <div>
          <span>Lens Breakage</span>
          <strong>${formatNumber(quality.lens)}</strong>
          <small>${lensSmall}</small>
        </div>
      </div>

      <div class="hero-metric quality-hero-card ${statusCss(quality.status)}">
        <span class="metric-icon">◇</span>
        <div>
          <span>Frame Breakage</span>
          <strong>${formatNumber(quality.frame)}</strong>
          <small>${frameSmall}</small>
        </div>
      </div>
    </section>
  `;
};

// Override associate quality section — shows partial counts live as dates load.
renderAssociateBreakageSection = function(summary) {
  ensureAssociateBreakageDetailsFast(summary);

  const rows = getAssociateBreakageRows(summary);
  const q = buildAssociateBreakageSummary(summary, rows);
  const byReason = buildBreakageDimensionRows(rows, 'BreakageReason', 6);
  const byAccessPoint = buildBreakageDimensionRows(rows, 'AccessPoint', 6);

  const dateKeys = getAssociateWorkedDateKeys(summary);
  const loadedCount = dateKeys.filter(dk => {
    const ck = getBreakageDateDetailsCacheKey(dk);
    return !!(STATE.breakageDateDetailsCache && STATE.breakageDateDetailsCache[ck]);
  }).length;
  const allLoaded = loadedCount >= dateKeys.length && dateKeys.length > 0;
  const loadNote = allLoaded ? '' : ` • ${loadedCount}/${dateKeys.length} days ↻`;

  return `
    <section class="associate-quality-panel command-card">
      <div class="section-head">
        <div>
          <h3>Quality / Breakage This Week</h3>
          <p>Matched to this associate by operator name from selected date details.${loadNote}</p>
        </div>
        ${statusBadge(q.status)}
      </div>

      <div class="quality-kpi-grid compact-quality">
        <article class="quality-kpi"><span>Operator Scans</span><strong>${formatNumber(q.scans)}</strong><small>From Activity</small></article>
        <article class="quality-kpi"><span>Lens Breakage</span><strong>${formatNumber(q.lens)}</strong><small>${formatQualityPercent(q.lensPercent)}</small></article>
        <article class="quality-kpi"><span>Frame Breakage</span><strong>${formatNumber(q.frame)}</strong><small>${formatQualityPercent(q.framePercent)}</small></article>
        <article class="quality-kpi"><span>Total Breakage</span><strong>${formatNumber(q.total)}</strong><small>${rows.length} records</small></article>
      </div>

      <div class="quality-grid-two">
        ${renderQualityListPanel('Top Reasons', byReason)}
        ${renderQualityListPanel('Access Points', byAccessPoint)}
      </div>
    </section>
  `;
};

/************************************************************
 * FINAL PATCH — CLIENT FISCAL WEEK DROPDOWN WINDOW
 *
 * Extra front-end protection:
 * even if old cache sends FW21, only show:
 * previous week / current week / next week based on browser today.
 ************************************************************/

function getSundayStartFromDate(dateValue) {
  const d = parseLocalDate(dateValue) || new Date();
  const local = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  local.setDate(local.getDate() - local.getDay());
  return local;
}

function getFiscalWeekNumberFromDate(dateValue) {
  const d = parseLocalDate(dateValue) || new Date();
  const weekStart = getSundayStartFromDate(d);
  const jan1 = new Date(d.getFullYear(), 0, 1);
  const firstWeekStart = getSundayStartFromDate(jan1);
  const diffDays = Math.floor((weekStart - firstWeekStart) / 86400000);
  return Math.floor(diffDays / 7) + 1;
}

function buildClientWeekObjectFromStart(startDate, isCurrent) {
  const start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const end = new Date(start);
  end.setDate(start.getDate() + 6);

  const fw = 'FW ' + getFiscalWeekNumberFromDate(start);

  const existing = (STATE.fiscalWeeks || []).find(w =>
    String(w.fiscalWeek || '') === fw &&
    formatDateKey(w.weekStartDate) === formatDateKey(start)
  );

  return {
    key: `${fw}|${formatDateKey(start)}|${formatDateKey(end)}`,
    fiscalWeek: fw,
    weekStartDate: formatDateKey(start),
    weekEndDate: formatDateKey(end),
    rowCount: existing ? Number(existing.rowCount || 0) : 0,
    workDateCount: existing ? Number(existing.workDateCount || 0) : 0,
    workDates: existing && existing.workDates ? existing.workDates : [],
    isCurrent: isCurrent === true
  };
}

getAvailableFiscalWeeks = function() {
  const currentStart = getSundayStartFromDate(new Date());

  const previousStart = new Date(currentStart);
  previousStart.setDate(currentStart.getDate() - 7);

  const nextStart = new Date(currentStart);
  nextStart.setDate(currentStart.getDate() + 7);

  return [
    buildClientWeekObjectFromStart(previousStart, false),
    buildClientWeekObjectFromStart(currentStart, true),
    buildClientWeekObjectFromStart(nextStart, false)
  ];
};

renderFiscalWeekDropdown = function() {
  const select = document.getElementById('fiscalWeekSelect');
  if (!select) return;

  const weeks = getAvailableFiscalWeeks();

  select.innerHTML = weeks.map(w => `
    <option value="${escapeHtml(w.key)}">
      ${escapeHtml(w.fiscalWeek)} — ${formatDatePretty(w.weekStartDate)} to ${formatDatePretty(w.weekEndDate)} • ${formatNumber(w.rowCount || 0)} rows
    </option>
  `).join('');

  let selectedKey = getSelectedFiscalWeekKey();
  if (!weeks.some(w => w.key === selectedKey)) {
    const currentWeek = weeks.find(w => w.isCurrent) || weeks[1] || weeks[0];
    if (currentWeek) {
      STATE.selectedFiscalWeek = currentWeek.fiscalWeek;
      STATE.selectedWeekStartDate = currentWeek.weekStartDate;
      STATE.selectedWeekEndDate = currentWeek.weekEndDate;
      selectedKey = currentWeek.key;
    }
  }

  select.value = selectedKey;
};

/************************************************************
 * FINAL PATCH — DEFAULT TO CURRENT WEEK + BREAKAGE DETAIL FAILSAFE
 ************************************************************/

function getCurrentClientFiscalWeekObject() {
  const weeks = getAvailableFiscalWeeks();
  return weeks.find(w => w.isCurrent) || weeks[1] || weeks[0] || null;
}

function forceCurrentWeekIfNeeded() {
  const current = getCurrentClientFiscalWeekObject();
  if (!current) return;

  // If the page booted into the previous week, reset to current week.
  // User can still manually pick FW22 or FW24 afterwards.
  if (!STATE._userChangedFiscalWeek) {
    STATE.selectedFiscalWeek = current.fiscalWeek;
    STATE.selectedWeekStartDate = current.weekStartDate;
    STATE.selectedWeekEndDate = current.weekEndDate;
  }
}

// Mark manual fiscal week changes so the user's selection is respected after the first load.
document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    const select = document.getElementById('fiscalWeekSelect');
    if (select && !select._productivityWeekManualBound) {
      select._productivityWeekManualBound = true;
      select.addEventListener('change', () => {
        STATE._userChangedFiscalWeek = true;
      });
    }
  }, 500);
});

// Override dropdown one more time: show FW22/FW23/FW24 but default to FW23/current.
renderFiscalWeekDropdown = function() {
  const select = document.getElementById('fiscalWeekSelect');
  if (!select) return;

  const weeks = getAvailableFiscalWeeks();
  const currentWeek = weeks.find(w => w.isCurrent) || weeks[1] || weeks[0];

  if (!STATE._userChangedFiscalWeek && currentWeek) {
    STATE.selectedFiscalWeek = currentWeek.fiscalWeek;
    STATE.selectedWeekStartDate = currentWeek.weekStartDate;
    STATE.selectedWeekEndDate = currentWeek.weekEndDate;
  }

  select.innerHTML = weeks.map(w => `
    <option value="${escapeHtml(w.key)}">
      ${escapeHtml(w.fiscalWeek)} — ${formatDatePretty(w.weekStartDate)} to ${formatDatePretty(w.weekEndDate)} • ${formatNumber(w.rowCount || 0)} rows
    </option>
  `).join('');

  select.value = getSelectedFiscalWeekKey();
};

// Failsafe: if a date detail fetch fails, write empty array so profile never stays stuck.
var __fetchBreakageDateDetailsBaseSafe = fetchBreakageDateDetails;
fetchBreakageDateDetails = function(dateKey) {
  return __fetchBreakageDateDetailsBaseSafe(dateKey)
    .catch(err => {
      console.warn('[Productivity Hub] Breakage date details failed:', err);
      STATE.breakageDateDetailsCache = STATE.breakageDateDetailsCache || {};
      STATE.breakageDateDetailsCache[getBreakageDateDetailsCacheKey(dateKey)] = [];
      showRefreshStatus('Breakage date details unavailable; showing 0 matched records.');
      renderEverything();
      return [];
    });
};

/************************************************************
 * PARALLEL BREAKAGE DATE PREFETCH
 *
 * ensureAssociateBreakageDetailsFast fires ALL missing date
 * fetches simultaneously instead of waiting for each one.
 * Tiles show partial real counts immediately and update as
 * each date resolves. No more 70s wait on "...".
 ************************************************************/

function ensureAssociateBreakageDetailsFast(summary) {
  if (!summary || !summary.rows || !summary.rows.length) return;

  const dateKeys = getAssociateWorkedDateKeys(summary);
  if (!dateKeys.length) return;

  STATE.associateBreakageLoadingDates = STATE.associateBreakageLoadingDates || {};

  dateKeys.forEach(dateKey => {
    const cacheKey = getBreakageDateDetailsCacheKey(dateKey);
    const alreadyCached = !!(STATE.breakageDateDetailsCache && STATE.breakageDateDetailsCache[cacheKey]);
    const alreadyLoading = !!STATE.associateBreakageLoadingDates[dateKey];
    if (alreadyCached || alreadyLoading) return;

    STATE.associateBreakageLoadingDates[dateKey] = true;

    fetchBreakageDateDetails(dateKey)
      .finally(() => {
        STATE.associateBreakageLoadingDates[dateKey] = false;
      });
  });
}

/************************************************************
 * BACKGROUND WEEK PREFETCH
 *
 * When the user switches fiscal weeks, silently prefetch all
 * breakage date details for that week in the background with
 * a small stagger so they are ready before any profile opens.
 ************************************************************/

(function patchWeekChangeBreakagePrefetch() {
  const _orig = loadFiscalWeekData;
  loadFiscalWeekData = function(week) {
    const result = _orig(week);

    const workDates = (week && week.workDates) ? week.workDates.slice() : [];
    workDates.forEach((dateKey, idx) => {
      if (!dateKey) return;
      setTimeout(() => {
        const cacheKey = getBreakageDateDetailsCacheKey(dateKey);
        if (STATE.breakageDateDetailsCache && STATE.breakageDateDetailsCache[cacheKey]) return;

        const qs = new URLSearchParams({
          action: 'breakageDateDetails',
          workDate: dateKey,
          ts: String(Date.now())
        });

        fetch(`${API_URL}?${qs.toString()}`)
          .then(r => r.text())
          .then(text => {
            let data;
            try { data = JSON.parse(text); } catch(e) { return; }
            if (!data || data.ok === false) return;

            const rows = normalizeBreakageRows(
              data.breakageDateDetails || data.breakageDailySnapshot || []
            );
            STATE.breakageDateDetailsCache = STATE.breakageDateDetailsCache || {};
            STATE.breakageDateDetailsCache[cacheKey] = rows;

            if (STATE.activePage === 'profile') renderEverything();
          })
          .catch(() => {});
      }, idx * 600);
    });

    return result;
  };
})();

console.info('[Productivity Hub] Parallel breakage prefetch active. Tiles show real counts immediately.');

/************************************************************
 * FINAL OVERRIDE — WEBPAGE ZENNI FISCAL WEEK LABEL FIX
 *
 * Confirmed Zenni calendar:
 * May 24, 2026 - May 30, 2026 = FW 22
 * May 31, 2026 - Jun 06, 2026 = FW 23
 * Jun 07, 2026 - Jun 13, 2026 = FW 24
 *
 * This override forces the webpage dropdown labels to match the
 * corrected API/sheet fiscal week logic.
 ************************************************************/

function getZenniWebSundayStart(dateValue) {
  const d = parseLocalDate(dateValue) || new Date();
  const start = new Date(d.getFullYear(), d.getMonth(), d.getDate());

  // Sunday = 0. If Sunday, subtract 0.
  start.setDate(start.getDate() - start.getDay());

  return start;
}

function getZenniWebFiscalWeekNumber(dateValue) {
  const d = parseLocalDate(dateValue) || new Date();
  const weekStart = getZenniWebSundayStart(d);
  const jan1 = new Date(d.getFullYear(), 0, 1);
  const fiscalWeekOneStart = getZenniWebSundayStart(jan1);
  const diffDays = Math.floor((weekStart - fiscalWeekOneStart) / 86400000);
  return Math.floor(diffDays / 7) + 1;
}

function getZenniWebFiscalInfo(dateValue) {
  const weekStart = getZenniWebSundayStart(dateValue);
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);

  return {
    fiscalWeek: 'FW ' + getZenniWebFiscalWeekNumber(weekStart),
    weekStartDate: formatDateKey(weekStart),
    weekEndDate: formatDateKey(weekEnd)
  };
}

function buildClientWeekObjectFromStart(startDate, isCurrent) {
  const fiscal = getZenniWebFiscalInfo(startDate);

  // Keep row count from API snapshot when the date window matches.
  const existing = (STATE.fiscalWeeks || []).find(w =>
    formatDateKey(w.weekStartDate) === fiscal.weekStartDate &&
    formatDateKey(w.weekEndDate) === fiscal.weekEndDate
  );

  return {
    key: `${fiscal.fiscalWeek}|${fiscal.weekStartDate}|${fiscal.weekEndDate}`,
    fiscalWeek: fiscal.fiscalWeek,
    weekStartDate: fiscal.weekStartDate,
    weekEndDate: fiscal.weekEndDate,
    rowCount: existing ? Number(existing.rowCount || 0) : 0,
    workDateCount: existing ? Number(existing.workDateCount || 0) : 0,
    workDates: existing && existing.workDates ? existing.workDates : [],
    isCurrent: isCurrent === true
  };
}

getAvailableFiscalWeeks = function() {
  const currentStart = getZenniWebSundayStart(new Date());

  const previousStart = new Date(currentStart);
  previousStart.setDate(currentStart.getDate() - 7);

  const nextStart = new Date(currentStart);
  nextStart.setDate(currentStart.getDate() + 7);

  return [
    buildClientWeekObjectFromStart(previousStart, false),
    buildClientWeekObjectFromStart(currentStart, true),
    buildClientWeekObjectFromStart(nextStart, false)
  ];
};

getSelectedFiscalWeekKey = function() {
  const weeks = getAvailableFiscalWeeks();

  if (STATE.selectedFiscalWeek && STATE.selectedWeekStartDate && STATE.selectedWeekEndDate) {
    const selectedStart = formatDateKey(STATE.selectedWeekStartDate);
    const selectedEnd = formatDateKey(STATE.selectedWeekEndDate);
    const selected = weeks.find(w => w.weekStartDate === selectedStart && w.weekEndDate === selectedEnd);
    if (selected) return selected.key;
  }

  const current = weeks.find(w => w.isCurrent) || weeks[1] || weeks[0];
  return current ? current.key : '';
};

renderFiscalWeekDropdown = function() {
  const select = document.getElementById('fiscalWeekSelect');
  if (!select) return;

  const weeks = getAvailableFiscalWeeks();
  const currentWeek = weeks.find(w => w.isCurrent) || weeks[1] || weeks[0];

  if (!STATE._userChangedFiscalWeek && currentWeek) {
    STATE.selectedFiscalWeek = currentWeek.fiscalWeek;
    STATE.selectedWeekStartDate = currentWeek.weekStartDate;
    STATE.selectedWeekEndDate = currentWeek.weekEndDate;
  }

  select.innerHTML = weeks.map(w => `
    <option value="${escapeHtml(w.key)}">
      ${escapeHtml(w.fiscalWeek)} — ${formatDatePretty(w.weekStartDate)} to ${formatDatePretty(w.weekEndDate)} • ${formatNumber(w.rowCount || 0)} rows
    </option>
  `).join('');

  select.value = getSelectedFiscalWeekKey();

  const selected = weeks.find(w => w.key === select.value) || currentWeek;
  if (selected) {
    STATE.selectedFiscalWeek = selected.fiscalWeek;
    STATE.selectedWeekStartDate = selected.weekStartDate;
    STATE.selectedWeekEndDate = selected.weekEndDate;
    updateWeekLabels();
  }
};

function verifyZenniWebFiscalWeeks() {
  const tests = [
    ['2026-05-24', 'FW 22'],
    ['2026-05-31', 'FW 23'],
    ['2026-06-07', 'FW 24']
  ];

  const result = tests.map(([dateKey, expected]) => {
    const actual = getZenniWebFiscalInfo(dateKey).fiscalWeek;
    return `${dateKey}: ${actual} ${actual === expected ? '✅' : '❌ expected ' + expected}`;
  }).join('\n');

  console.info('[Productivity Hub] Zenni fiscal week check\n' + result);
  return result;
}

document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    verifyZenniWebFiscalWeeks();
    renderFiscalWeekDropdown();
  }, 1000);
});

/************************************************************
 * FINAL OVERRIDE — WEBPAGE READS ALL FISCAL WEEKS
 *
 * The dropdown now reads every week returned from the API/sheet.
 * It does NOT limit to previous/current/next.
 *
 * It still corrects the fiscal week label from WeekStartDate using
 * the Zenni rule:
 * 2026-05-24 = FW 22
 * 2026-05-31 = FW 23
 * 2026-06-07 = FW 24
 ************************************************************/

function normalizeFiscalWeekForWeb(week) {
  const startKey = formatDateKey(week.weekStartDate || week.WeekStartDate);
  const endKey = formatDateKey(week.weekEndDate || week.WeekEndDate);
  const fiscal = getZenniWebFiscalInfo(startKey);

  return {
    ...week,
    key: `${fiscal.fiscalWeek}|${fiscal.weekStartDate}|${fiscal.weekEndDate}`,
    fiscalWeek: fiscal.fiscalWeek,
    weekStartDate: fiscal.weekStartDate,
    weekEndDate: fiscal.weekEndDate,
    rowCount: Number(week.rowCount || week.RowCount || 0),
    workDateCount: Number(week.workDateCount || week.WorkDateCount || 0),
    workDates: week.workDates || week.WorkDates || [],
    isCurrent: isCurrentFiscalWeekWindow(fiscal.weekStartDate, fiscal.weekEndDate)
  };
}

function isCurrentFiscalWeekWindow(weekStartDate, weekEndDate) {
  const today = new Date();
  const start = parseLocalDate(weekStartDate);
  const end = parseLocalDate(weekEndDate);

  if (!start || !end) return false;

  return today >= start && today <= end;
}

getAvailableFiscalWeeks = function() {
  const map = new Map();

  (STATE.fiscalWeeks || []).forEach(rawWeek => {
    if (!rawWeek || !rawWeek.weekStartDate) return;

    const week = normalizeFiscalWeekForWeb(rawWeek);
    map.set(week.key, week);
  });

  // Include current week even if snapshot has no rows yet.
  const current = getZenniWebFiscalInfo(new Date());
  const currentKey = `${current.fiscalWeek}|${current.weekStartDate}|${current.weekEndDate}`;

  if (!map.has(currentKey)) {
    map.set(currentKey, {
      key: currentKey,
      fiscalWeek: current.fiscalWeek,
      weekStartDate: current.weekStartDate,
      weekEndDate: current.weekEndDate,
      rowCount: 0,
      workDateCount: 0,
      workDates: [],
      isCurrent: true
    });
  } else {
    map.get(currentKey).isCurrent = true;
  }

  return Array.from(map.values()).sort((a, b) => {
    const aDate = parseLocalDate(a.weekStartDate);
    const bDate = parseLocalDate(b.weekStartDate);
    return bDate - aDate; // newest first
  });
};

getSelectedFiscalWeekKey = function() {
  const weeks = getAvailableFiscalWeeks();

  if (STATE.selectedWeekStartDate && STATE.selectedWeekEndDate) {
    const selectedStart = formatDateKey(STATE.selectedWeekStartDate);
    const selectedEnd = formatDateKey(STATE.selectedWeekEndDate);
    const selected = weeks.find(w => w.weekStartDate === selectedStart && w.weekEndDate === selectedEnd);
    if (selected) return selected.key;
  }

  const current = weeks.find(w => w.isCurrent) || weeks[0];
  return current ? current.key : '';
};

renderFiscalWeekDropdown = function() {
  const select = document.getElementById('fiscalWeekSelect');
  if (!select) return;

  const weeks = getAvailableFiscalWeeks();
  const currentWeek = weeks.find(w => w.isCurrent) || weeks[0];

  if (!STATE._userChangedFiscalWeek && currentWeek) {
    STATE.selectedFiscalWeek = currentWeek.fiscalWeek;
    STATE.selectedWeekStartDate = currentWeek.weekStartDate;
    STATE.selectedWeekEndDate = currentWeek.weekEndDate;
  }

  select.innerHTML = weeks.map(w => `
    <option value="${escapeHtml(w.key)}">
      ${escapeHtml(w.fiscalWeek)} — ${formatDatePretty(w.weekStartDate)} to ${formatDatePretty(w.weekEndDate)} • ${formatNumber(w.rowCount || 0)} rows
    </option>
  `).join('');

  select.value = getSelectedFiscalWeekKey();

  const selected = weeks.find(w => w.key === select.value) || currentWeek;
  if (selected) {
    STATE.selectedFiscalWeek = selected.fiscalWeek;
    STATE.selectedWeekStartDate = selected.weekStartDate;
    STATE.selectedWeekEndDate = selected.weekEndDate;
    updateWeekLabels();
  }
};

document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    console.info('[Productivity Hub] Fiscal week dropdown now reads all available FW from DAILY_SNAPSHOT.');
    renderFiscalWeekDropdown();
  }, 1200);
});

/************************************************************
 * EMERGENCY FINAL OVERRIDE — FIX FW LABEL + updateWeekLabels ERROR
 *
 * Fixes:
 * 1. 2026-05-24 was still showing FW21 on webpage.
 * 2. updateWeekLabels is not defined error.
 * 3. Dropdown reads ALL fiscal weeks from API/DAILY_SNAPSHOT.
 *
 * Anchor confirmed by Zenni:
 * 2026-05-24 to 2026-05-30 = FW 22
 ************************************************************/

function getZenniWebFiscalInfo(dateValue) {
  const inputDate = parseLocalDate(dateValue) || new Date();
  const weekStart = new Date(inputDate.getFullYear(), inputDate.getMonth(), inputDate.getDate());
  weekStart.setDate(weekStart.getDate() - weekStart.getDay()); // Sunday start

  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);

  // Confirmed anchor: Sunday 2026-05-24 is FW 22
  const anchorStart = new Date(2026, 4, 24); // May 24, 2026
  const anchorFw = 22;

  const diffWeeks = Math.round((weekStart - anchorStart) / (7 * 86400000));
  const weekNum = anchorFw + diffWeeks;

  return {
    fiscalWeek: 'FW ' + weekNum,
    weekStartDate: formatDateKey(weekStart),
    weekEndDate: formatDateKey(weekEnd)
  };
}

function isCurrentFiscalWeekWindow(weekStartDate, weekEndDate) {
  const today = new Date();
  const start = parseLocalDate(weekStartDate);
  const end = parseLocalDate(weekEndDate);

  if (!start || !end) return false;

  const t = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return t >= start && t <= end;
}

function safeUpdateWeekLabelsFinal() {
  if (typeof updateWeekLabels === 'function') {
    updateWeekLabels();
    return;
  }

  // Fallback for older JS that does not have updateWeekLabels().
  const labels = [
    document.getElementById('selectedWeekLabel'),
    document.getElementById('currentWeekLabel'),
    document.querySelector('[data-selected-week-label]'),
    document.querySelector('[data-current-week-label]')
  ].filter(Boolean);

  const text = STATE.selectedFiscalWeek
    ? `${STATE.selectedFiscalWeek} • ${formatDatePretty(STATE.selectedWeekStartDate)} to ${formatDatePretty(STATE.selectedWeekEndDate)}`
    : '';

  labels.forEach(el => {
    el.textContent = text;
  });
}

function normalizeFiscalWeekForWeb(week) {
  const rawStart = week.weekStartDate || week.WeekStartDate || week.weekStart || week.WeekStart;
  const rawEnd = week.weekEndDate || week.WeekEndDate || week.weekEnd || week.WeekEnd;

  const fiscal = getZenniWebFiscalInfo(rawStart || new Date());

  return {
    ...week,
    key: `${fiscal.fiscalWeek}|${fiscal.weekStartDate}|${fiscal.weekEndDate}`,
    fiscalWeek: fiscal.fiscalWeek,
    weekStartDate: fiscal.weekStartDate,
    weekEndDate: fiscal.weekEndDate,
    rowCount: Number(week.rowCount || week.RowCount || 0),
    workDateCount: Number(week.workDateCount || week.WorkDateCount || 0),
    workDates: week.workDates || week.WorkDates || [],
    isCurrent: isCurrentFiscalWeekWindow(fiscal.weekStartDate, fiscal.weekEndDate)
  };
}

getAvailableFiscalWeeks = function() {
  const map = new Map();

  (STATE.fiscalWeeks || []).forEach(rawWeek => {
    if (!rawWeek) return;

    const start = rawWeek.weekStartDate || rawWeek.WeekStartDate || rawWeek.weekStart || rawWeek.WeekStart;
    if (!start) return;

    const week = normalizeFiscalWeekForWeb(rawWeek);
    map.set(week.key, week);
  });

  // Always include current week even if no snapshot rows exist yet.
  const current = getZenniWebFiscalInfo(new Date());
  const currentKey = `${current.fiscalWeek}|${current.weekStartDate}|${current.weekEndDate}`;

  if (!map.has(currentKey)) {
    map.set(currentKey, {
      key: currentKey,
      fiscalWeek: current.fiscalWeek,
      weekStartDate: current.weekStartDate,
      weekEndDate: current.weekEndDate,
      rowCount: 0,
      workDateCount: 0,
      workDates: [],
      isCurrent: true
    });
  } else {
    map.get(currentKey).isCurrent = true;
  }

  return Array.from(map.values()).sort((a, b) => {
    const aDate = parseLocalDate(a.weekStartDate);
    const bDate = parseLocalDate(b.weekStartDate);
    return bDate - aDate;
  });
};

getSelectedFiscalWeekKey = function() {
  const weeks = getAvailableFiscalWeeks();

  if (STATE.selectedWeekStartDate && STATE.selectedWeekEndDate) {
    const selectedStart = formatDateKey(STATE.selectedWeekStartDate);
    const selectedEnd = formatDateKey(STATE.selectedWeekEndDate);
    const selected = weeks.find(w => w.weekStartDate === selectedStart && w.weekEndDate === selectedEnd);
    if (selected) return selected.key;
  }

  const current = weeks.find(w => w.isCurrent) || weeks[0];
  return current ? current.key : '';
};

renderFiscalWeekDropdown = function() {
  const select = document.getElementById('fiscalWeekSelect');
  if (!select) return;

  const weeks = getAvailableFiscalWeeks();
  const currentWeek = weeks.find(w => w.isCurrent) || weeks[0];

  if (!STATE._userChangedFiscalWeek && currentWeek) {
    STATE.selectedFiscalWeek = currentWeek.fiscalWeek;
    STATE.selectedWeekStartDate = currentWeek.weekStartDate;
    STATE.selectedWeekEndDate = currentWeek.weekEndDate;
  }

  select.innerHTML = weeks.map(w => `
    <option value="${escapeHtml(w.key)}">
      ${escapeHtml(w.fiscalWeek)} — ${formatDatePretty(w.weekStartDate)} to ${formatDatePretty(w.weekEndDate)} • ${formatNumber(w.rowCount || 0)} rows
    </option>
  `).join('');

  select.value = getSelectedFiscalWeekKey();

  const selected = weeks.find(w => w.key === select.value) || currentWeek;
  if (selected) {
    STATE.selectedFiscalWeek = selected.fiscalWeek;
    STATE.selectedWeekStartDate = selected.weekStartDate;
    STATE.selectedWeekEndDate = selected.weekEndDate;
    safeUpdateWeekLabelsFinal();
  }
};

function verifyZenniWebFiscalWeeks_FINAL() {
  const tests = [
    ['2026-05-24', 'FW 22'],
    ['2026-05-31', 'FW 23'],
    ['2026-06-07', 'FW 24']
  ];

  const result = tests.map(([dateKey, expected]) => {
    const actual = getZenniWebFiscalInfo(dateKey).fiscalWeek;
    return `${dateKey}: ${actual} ${actual === expected ? '✅' : '❌ expected ' + expected}`;
  }).join('\n');

  console.info('[Productivity Hub] FINAL Zenni fiscal week check\n' + result);
  return result;
}

document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    verifyZenniWebFiscalWeeks_FINAL();
    renderFiscalWeekDropdown();
  }, 1500);
});

/************************************************************
 * FINAL OVERRIDE — PRODUCTIVITY ONLY + BREAKAGE COUNTS ONLY
 *
 * Productivity Hub should NOT open/load the old Quality / Breakage Hub tab.
 * Full breakage detail now belongs on BreakageDashboard.html.
 *
 * This page keeps only the lightweight associate scorecard breakage counts:
 * - Lens Breakage
 * - Frame Breakage
 ************************************************************/

function renderEverything() {
  renderTopSummary();
  renderAssociateDropdown();

  if (STATE.activePage === 'breakage') {
    STATE.activePage = 'scorecards';
  }

  if (STATE.activePage === 'dashboard') {
    renderDashboardHub();
    return;
  }

  if (STATE.activePage === 'scorecards') {
    renderScorecardsPage();
    return;
  }

  if (STATE.activePage === 'profile') {
    renderProfile();
    return;
  }

  if (STATE.activePage === 'daily') {
    renderDailySummaryHub();
    return;
  }

  if (STATE.activePage === 'trends') {
    renderTrendsHub();
    return;
  }

  if (STATE.activePage === 'targets') {
    renderTargetsHub();
    return;
  }

  renderPlaceholderPage(STATE.activePage);
}

function setActivePage(page) {
  const requestedPage = page || 'dashboard';

  // Old Quality / Breakage Hub tab removed from Productivity.
  // Full details are now on the separate BreakageDashboard.html page.
  STATE.activePage = requestedPage === 'breakage' ? 'scorecards' : requestedPage;

  if (STATE.activePage !== 'profile') {
    STATE.selectedAssociate = STATE.selectedAssociate || '';
  }

  setActiveSideNav(STATE.activePage === 'profile' ? 'scorecards' : STATE.activePage);
  renderEverything();
}

function openSeparateBreakageHub() {
  window.location.href = 'BreakageDashboard.html';
}

console.info('[Productivity Hub] Old Breakage Hub tab disabled. Associate scorecard still shows lens/frame breakage counts only.');