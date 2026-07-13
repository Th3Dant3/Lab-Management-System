const SURFACE_DASHBOARD_API_URL = 'https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec';
const SURFACE_DASHBOARD_REFRESH_MS = 5 * 60 * 1000;

const SURFACE_API_URL = 'https://script.google.com/macros/s/AKfycbzaropKAS7ujmUgz88cjchj9lFCJS3VX-TCBFmaE8x449QtEiWV-3cbTvVuRXxIpedv/exec';


// Productive hours use clock time minus unpaid/covered breaks.
// Weekday: Mon-Thu 7:00 AM - 5:30 PM = 9.5 productive hrs.
// 12-hour shifts: 12 clock hrs with standard planned non-productive time = 10.5 productive hrs.
const SHIFT_PRODUCTIVE_HOURS = {
  Weekday: 9.5,
  'Weekday OT 12': 10.5,
  Weekend: 10.5
};

// These are planning ranges, not exact promises.
// Surface usually receives extra incoming jobs during the day and remake/breakage work.
// The forecast uses low/high ranges so leadership sees risk instead of one fake number.
const SURFACE_EXTRA_WORK_DEFAULTS = {
  incomingLow: 300,
  incomingHigh: 500,
  remakeLow: 150,
  remakeHigh: 200
};

const SURFACE_BOS_WIP_STORAGE_KEY = 'productionGoalForecast.surface.bosByDate.v1';
const SURFACE_STARTUP_OUTPUT_DELAY_MIN = 90;

// Minimum estimated travel time from each last-scan WIP location to Surface OUT.
// These are planning assumptions and can be tuned after comparing forecast vs actual.
const SURFACE_STAGE_TIME_RANGES = Object.freeze({
  unbox: { low: 20, high: 30, label: '20–30 min' },
  autoblocker: { low: 5, high: 10, label: '5–10 min' },
  cooling: { low: 25, high: 50, label: '25–50 min dwell' },
  orb: { low: 20, high: 30, label: '20–30 min' },
  polisher: { low: 20, high: 30, label: '20–30 min' },
  engraving: { low: 10, high: 15, label: '10–15 min' },
  detaping: { low: 5, high: 10, label: '5–10 min' },
  coater: { low: 5, high: 10, label: '5–10 min' },
  inspection: { low: 0, high: 0, label: '0 min' }
});

// Cumulative midpoint travel time from each last-scan station to AR41 OUT.
// The card displays the direct station time-to-clear range above.
const SURFACE_MATURITY_MINUTES = Object.freeze({
  inspection: 0,
  coater: 8,
  detaping: 16,
  engraving: 29,
  polisher: 54,
  orb: 79,
  cooling: 117,
  autoblocker: 125,
  unbox: 150
});


const SURFACE_AREAS = [
 
  { key: 'unbox', label: 'SF Unbox', group: 'SF Unbox', type: 'operator', capacityMode: 'operator', jphPerUnit: 150, required: true, defaultUnits: 1.5, defaultAssociates: 1.5 },
  { key: 'autoblocker', label: 'Auto Blocker', group: 'Auto Blocker & Generator', type: 'machine', capacityMode: 'machine', jphPerUnit: 50, required: true, defaultUnits: 4, defaultAssociates: 1, machinesPerAssociate: 4 },
  { key: 'cooling', label: 'IQ Star / Cooling', group: 'Auto Blocker & Generator', type: 'buffer', capacityMode: 'buffer', jphPerUnit: 0, required: false, defaultUnits: 2, defaultAssociates: 0, bufferCapacity: 180, coolingLowMin: 25, coolingHighMin: 50 },
  { key: 'orb', label: 'ORB / Generator', group: 'Auto Blocker & Generator', type: 'machine', capacityMode: 'machine', jphPerUnit: 35, required: true, defaultUnits: 6, defaultAssociates: 1, machinesPerAssociate: 6 },
  { key: 'polisher', label: 'Polisher / Flex', group: 'Polisher / Engraving / Detaping', type: 'machine', capacityMode: 'machine', jphPerUnit: 35, required: true, defaultUnits: 6, defaultAssociates: 1, machinesPerAssociate: 6 },
  { key: 'engraving', label: 'Engraving / OTL', group: 'Polisher / Engraving / Detaping', type: 'machine', capacityMode: 'machine', jphPerUnit: 75, required: true, defaultUnits: 2, defaultAssociates: 1, machinesPerAssociate: 2 },
  { key: 'detaping', label: 'Detaping / ODT', group: 'Polisher / Engraving / Detaping', type: 'machine', capacityMode: 'machine', jphPerUnit: 100, required: true, defaultUnits: 2, defaultAssociates: 1, machinesPerAssociate: 2 },
  { key: 'coater', label: 'Coater / 54R', group: 'Coater & Surface Inspection', type: 'machine', capacityMode: 'machine', jphPerUnit: 55, required: true, defaultUnits: 4, defaultAssociates: 1, machinesPerAssociate: 4 },
  { key: 'inspection', label: 'Surface Inspection / AR41', group: 'Coater & Surface Inspection', type: 'machine', capacityMode: 'machine', jphPerUnit: 100, required: true, defaultUnits: 2, defaultAssociates: 2, machinesPerAssociate: 1.25 }
];


// Live WIP from the dashboard is treated as the LAST SCAN location.
// Bottleneck pressure is calculated against the NEXT process that must clear that WIP.

const SURFACE_FLOW_ORDER = [
  'unbox',
  'autoblocker',
  'cooling',
  'orb',
  'polisher',
  'engraving',
  'detaping',
  'coater',
  'inspection'
];

function getSurfaceFlowOrderIndex(key) {
  const index = SURFACE_FLOW_ORDER.indexOf(String(key || ''));
  return index >= 0 ? index : 999;
}

const SURFACE_PRESSURE_ROUTE = {
  unbox: {
    pressureAreaKey: 'autoblocker',
    pressureLabel: 'Taping / Auto Blocker',
    timeMode: 'sf_unbox_stage',
    note: 'SF Unbox scan is feeding Taping / Auto Blocker'
  },
  autoblocker: {
    pressureAreaKey: 'cooling',
    pressureLabel: 'Cooling',
    timeMode: 'cooling_transfer',
    bufferSignalLimit: 200,
    note: 'Auto Blocker scan is feeding Cooling'
  },
  cooling: {
    pressureAreaKey: 'orb',
    pressureLabel: 'ORB / Generator',
    timeMode: 'cooling_dwell',
    note: 'Cooling scan is feeding ORB / Generator'
  },
  orb: {
    pressureAreaKey: 'polisher',
    pressureLabel: 'Polisher / Flex',
    note: 'ORB scan is feeding Polisher / Flex'
  },
  polisher: {
    pressureAreaKey: 'engraving',
    pressureLabel: 'Engraving / OTL',
    note: 'Polisher scan is feeding Engraving / OTL'
  },
  engraving: {
    pressureAreaKey: 'detaping',
    pressureLabel: 'Manual Deblock / Detaping',
    capacityAreaKey: 'engraving',
    capacityLabel: 'Engraving / OTL',
    note: 'Engraving scan is feeding Manual Deblock / Detaping. Capacity stays tied to Engraving because Manual Deblock / Detaping has no reliable LMS scan/JPH feed.'
  },
  detaping: {
    pressureAreaKey: 'coater',
    pressureLabel: 'Coater / 54R',
    note: 'Detaping scan is feeding Coater / 54R'
  },
  coater: {
    pressureAreaKey: 'inspection',
    pressureLabel: 'AR41 / Surface OUT',
    note: 'Coater scan is feeding AR41 / Surface OUT'
  },
  inspection: {
    pressureAreaKey: 'inspection',
    pressureLabel: 'AR41 / Surface OUT',
    note: 'AR41 scan is final Surface OUT'
  }
};

const HOURLY_DEFAULTS = [
  '06:00', '07:00', '08:00', '09:00', '10:00', '11:00', '12:00',
  '13:00', '14:00', '15:00', '16:00', '17:00', '18:00', '19:00'
];

let shiftModeManualOverride = false;

let downtimeCounter = 0;
let latestForecastResult = null;

const SURFACE_STORAGE_KEY = 'surfaceForecastCommandCenter.v17-ar-goal-comparison';
let isRestoringState = false;

document.addEventListener('DOMContentLoaded', async function () {
  initTabs();
  buildAreaGroups();
  buildFloaterOptions();
  buildHourlyRows();
  initPerformanceControls();
  ensureAdditionalWorkControls();
  updateForecastRuleCopy();
  wireButtons();

  const initialDepartment =
    localStorage.getItem('productionGoalForecast.department') === 'AR'
      ? 'AR'
      : 'Surface';

  // Do not open the Surface Forecast storage API when the page starts in AR.
  // AR has its own local setup and reads live data from Production Flow.
  if (initialDepartment === 'Surface') {
    const restoredFromCloud = await loadCurrentStateFromApi();
    const restored = restoredFromCloud || restorePageState();

    if (!restored) {
      addDowntimeRow();
    }

    calculateForecast();
    savePageState();
    await syncSurfaceDashboardData(true);
  } else {
    if (!document.querySelector('.downtime-row')) {
      addDowntimeRow();
    }

    await syncARForecastData(true);
  }

  // One refresh timer. It refreshes only the department currently being viewed.
  window.setInterval(function () {
    const department =
      window.AR_FORECAST_APP?.getDepartment?.() ||
      localStorage.getItem('productionGoalForecast.department') ||
      'Surface';

    if (department === 'AR') {
      syncARForecastData(false);
    } else {
      syncSurfaceDashboardData(false);
    }
  }, SURFACE_DASHBOARD_REFRESH_MS);
});

const SURFACE_TAB_STORAGE_KEY = 'surfaceForecastCommandCenter.activeTab';

function initTabs() {
  const buttons = Array.from(document.querySelectorAll('.tab-btn'));
  const panels = Array.from(document.querySelectorAll('.tab-panel'));
  if (!buttons.length || !panels.length) return;

  const savedTab = localStorage.getItem(SURFACE_TAB_STORAGE_KEY);
  const validTab = buttons.some(btn => btn.dataset.tabTarget === savedTab) ? savedTab : 'status';

  activateTab(validTab, buttons, panels);

  buttons.forEach(function (btn) {
    btn.addEventListener('click', function () {
      activateTab(btn.dataset.tabTarget, buttons, panels);
      localStorage.setItem(SURFACE_TAB_STORAGE_KEY, btn.dataset.tabTarget);
    });
  });
}

function activateTab(target, buttons, panels) {
  buttons.forEach(function (btn) {
    btn.classList.toggle('active', btn.dataset.tabTarget === target);
  });

  const department = window.AR_FORECAST_APP ? window.AR_FORECAST_APP.getDepartment() : 'Surface';
  panels.forEach(function (panel) {
    const panelDepartment = panel.dataset.department || 'Surface';
    panel.hidden = panel.dataset.tabPanel !== target || panelDepartment !== department;
  });
}

function wireButtons() {
  const buttons = {
    calculateBtn: calculateForecast,
    addDowntimeBtn: addDowntimeRow,
    clearDowntimeBtn: clearDowntimeRows,
    saveBtn: saveForecast,
    saveDowntimeBtn: saveDowntime,
    clearHourlyBtn: clearHourlyRows,
    saveHourlyBtn: saveHourly,
    syncDashboardBtn: function () { syncSurfaceDashboardData(true); },
    saveSurfaceBosBtn: saveSurfaceBosSnapshot,
    resetSurfaceBosBtn: resetSurfaceBosSnapshot
  };

  Object.keys(buttons).forEach(function (id) {
    const el = document.getElementById(id);
    if (el) el.addEventListener('click', buttons[id]);
  });

  const bosTotalInput = document.getElementById('bosWip');
  if (bosTotalInput) {
    bosTotalInput.addEventListener('input', function () {
      setText('surfaceBosStartupMeta', 'Unsaved total BOS change — click Save BOS.');
    });
  }

  ['shiftMode', 'productionDate', 'perfShiftModeSelect', 'perfDateInput', 'shipGoal', 'incomingLow', 'incomingHigh', 'remakeLow', 'remakeHigh', 'floaterCount', 'floaterAssign'].forEach(function (id) {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('input', calculateForecast);
    el.addEventListener('change', calculateForecast);
  });

  document.addEventListener('input', autoSavePageState);
  document.addEventListener('change', autoSavePageState);
}

function updateForecastRuleCopy() {
  const oldTexts = [
    'Incoming and remake/breakage are not guessed. Forecast range uses current BOS WIP and capacity ceiling. Hourly input tells you if actual AR41 output is on track.',
    'Forecast range uses BOS WIP + average incoming Surface jobs + average breakage/remake jobs, capped by usable bottleneck capacity. Hourly input tells you if actual AR41 output is on track.'
  ];
  const newText = 'Forecast uses BOS SF WIP, route-specific capacity, 300–500 incoming jobs, 150–200 breakage/remake jobs, machine ceilings, associate coverage, downtime, and actual AR41 output.';

  Array.from(document.querySelectorAll('body *')).forEach(function (el) {
    if (!el || !el.childNodes || el.children.length > 0) return;
    const text = String(el.textContent || '').trim();
    if (oldTexts.includes(text)) el.textContent = newText;
  });
}

function ensureAdditionalWorkControls() {
  // Keep incoming/remake assumptions in the background.
  // The forecast engine still reads these values, but operators do not need to see or edit them in Daily Setup.
  const hiddenValues = {
    incomingLow: SURFACE_EXTRA_WORK_DEFAULTS.incomingLow,
    incomingHigh: SURFACE_EXTRA_WORK_DEFAULTS.incomingHigh,
    remakeLow: SURFACE_EXTRA_WORK_DEFAULTS.remakeLow,
    remakeHigh: SURFACE_EXTRA_WORK_DEFAULTS.remakeHigh
  };

  Object.keys(hiddenValues).forEach(function (id) {
    let input = document.getElementById(id);

    if (!input) {
      input = document.createElement('input');
      input.id = id;
      input.type = 'hidden';
      document.body.appendChild(input);
    }

    input.type = 'hidden';
    input.value = hiddenValues[id];
    input.setAttribute('aria-hidden', 'true');
    input.tabIndex = -1;
  });

  hideOldVisibleAdditionalWorkRows();
}

function hideOldVisibleAdditionalWorkRows() {
  // Safety cleanup in case an older cached V20 script already injected the visible labels.
  ['incomingLow', 'incomingHigh', 'remakeLow', 'remakeHigh'].forEach(function (id) {
    const input = document.getElementById(id);
    if (!input) return;

    input.type = 'hidden';

    const label = input.closest('label');
    if (label) {
      label.style.display = 'none';
      label.setAttribute('aria-hidden', 'true');
    }

    const wrapper = input.closest('.range-input-pair');
    if (wrapper) {
      wrapper.style.display = 'none';
      wrapper.setAttribute('aria-hidden', 'true');
    }
  });
}

function buildAreaGroups() {
  const holder = document.getElementById('areaGroups');
  if (!holder) return;

  holder.innerHTML = '';

  const groups = [
    'SF Unbox',
    'Auto Blocker & Generator',
    'Polisher / Engraving / Detaping',
    'Coater & Surface Inspection'
  ];

  groups.forEach(function (groupName) {
    const groupAreas = SURFACE_AREAS.filter(function (area) {
      return area.group === groupName;
    });

    const group = document.createElement('div');
    group.className = 'area-group';
    group.innerHTML = `<h3>${escapeHtml(groupName)}</h3>`;

    groupAreas.forEach(function (area) {
      const unitLabel = area.type === 'operator' ? 'Associate Count' : area.type === 'buffer' ? 'Towers / Buffer Running' : 'Machines Running';

      const card = document.createElement('div');
      card.className = 'area-card';

      card.innerHTML = `
        <div>
          <div class="area-card-title">${escapeHtml(area.label)}</div>
          <small>${area.type === 'buffer' ? 'Cooling range 25–50 min · Buffer capacity 180 jobs' : area.jphPerUnit + ' JPH each'}</small>
        </div>

        <label>
          Associates
          <input id="assoc_${area.key}" type="number" min="0" step="0.5" value="${area.defaultAssociates}">
        </label>

        <label>
          ${unitLabel}
          <input id="area_${area.key}" type="number" min="0" step="0.5" value="${area.defaultUnits}">
        </label>

        ${area.type === 'machine' ? `
        <label>
          Machines One Associate Can Cover
          <input id="ratio_${area.key}" type="number" min="0.25" step="0.25" value="${area.machinesPerAssociate}">
          <small class="label-note">Errors/jams stay cleared up to this many machines. Above it: downtime risk, not a hard stop.</small>
        </label>
        ` : ''}
      `;

      group.appendChild(card);

      card.querySelectorAll('input').forEach(function (input) {
        input.addEventListener('input', calculateForecast);
      });
    });

    holder.appendChild(group);
  });
}

function buildFloaterOptions() {
  const select = document.getElementById('floaterAssign');
  if (!select) return;

  SURFACE_AREAS.forEach(function (area) {
    const option = document.createElement('option');
    option.value = area.key;
    option.textContent = area.label;
    select.appendChild(option);
  });
}

function buildHourlyRows(shiftMode, preserveExisting = true) {
  const body = document.getElementById('hourlyRows');
  if (!body) return;

  const selectedShift = shiftMode || getValue('shiftMode', 'Weekday');
  const windowInfo = getShiftWindow(selectedShift);
  const startMinutes = parseTimeToMinutes(windowInfo.start) || 420;
  const endMinutes = parseTimeToMinutes(windowInfo.end) || 1050;
  const slots = buildPerformanceHourSlots(startMinutes, endMinutes);
  const previous = {};

  if (preserveExisting) {
    document.querySelectorAll('.hourly-row').forEach(function (row) {
      const key = row.dataset.apiHour || normalizeHourKey(getRowValue(row, '.hr-hour'));
      previous[key] = {
        unbox: getRowValue(row, '.hr-unbox'), autoblocker: getRowValue(row, '.hr-autoblocker'),
        cooling: getRowValue(row, '.hr-cooling'), orb: getRowValue(row, '.hr-orb'),
        polisher: getRowValue(row, '.hr-polisher'), engraving: getRowValue(row, '.hr-engraving'),
        detaping: getRowValue(row, '.hr-detaping'), coater: getRowValue(row, '.hr-coater'),
        inspection: getRowValue(row, '.hr-inspection'), notes: getRowValue(row, '.hr-notes'),
        active: getRowValue(row, '.hr-active') || 'Yes'
      };
    });
  }

  body.innerHTML = '';

  slots.forEach(function (slot) {
    const apiHour = minutesToHourKey(Math.floor(slot.start / 60) * 60);
    const saved = previous[apiHour] || {};
    const tr = document.createElement('tr');
    tr.className = 'hourly-row';
    tr.dataset.apiHour = apiHour;

    tr.innerHTML = `
      <td><input class="hour-input hr-hour" type="text" value="${escapeHtml(slot.label)}" readonly></td>
      <td><input class="hr-unbox" type="number" min="0" step="1" value="${escapeHtml(saved.unbox || '')}"></td>
      <td><input class="hr-autoblocker" type="number" min="0" step="1" value="${escapeHtml(saved.autoblocker || '')}"></td>
      <td><input class="hr-cooling" type="number" min="0" step="1" value="${escapeHtml(saved.cooling || '')}"></td>
      <td><input class="hr-orb" type="number" min="0" step="1" value="${escapeHtml(saved.orb || '')}"></td>
      <td><input class="hr-polisher" type="number" min="0" step="1" value="${escapeHtml(saved.polisher || '')}"></td>
      <td><input class="hr-engraving" type="number" min="0" step="1" value="${escapeHtml(saved.engraving || '')}"></td>
      <td><input class="hr-detaping" type="number" min="0" step="1" value="${escapeHtml(saved.detaping || '')}"></td>
      <td><input class="hr-coater" type="number" min="0" step="1" value="${escapeHtml(saved.coater || '')}"></td>
      <td><input class="hr-inspection" type="number" min="0" step="1" value="${escapeHtml(saved.inspection || '')}"></td>
      <td><input class="notes-input hr-notes" type="text" value="${escapeHtml(saved.notes || '')}"></td>
      <td><select class="hr-active"><option value="Yes">Yes</option><option value="No">No</option></select></td>`;

    body.appendChild(tr);
    const active = tr.querySelector('.hr-active');
    if (active) active.value = saved.active || 'Yes';
    tr.querySelectorAll('input, select').forEach(function (input) {
      input.addEventListener('input', calculateForecast);
      input.addEventListener('change', calculateForecast);
    });
  });
}

function addDowntimeRow() {
  downtimeCounter += 1;

  const holder = document.getElementById('downtimeRows');
  if (!holder) return;

  const row = document.createElement('div');
  row.className = 'downtime-row';
  row.dataset.id = String(downtimeCounter);

  const areaOptions = SURFACE_AREAS.map(function (area) {
    return `<option value="${area.key}">${escapeHtml(area.label)}</option>`;
  }).join('');

  row.innerHTML = `
    <label>
      Area
      <select class="dt-area">${areaOptions}</select>
    </label>

    <label>
      Issue
      <select class="dt-issue">
        <option value="Machine Down">Machine Down</option>
        <option value="Preventive Maintenance">Preventive Maintenance</option>
        <option value="Low Work">Low Work / Starved</option>
        <option value="Quality Issue">Quality Issue</option>
        <option value="Operator Short">Operator Short</option>
      </select>
    </label>

    <label>
      Units Down
      <input class="dt-units" type="number" min="0" step="0.5" value="1">
    </label>

    <label>
      Start
      <input class="dt-start" type="text" value="08:45">
    </label>

    <label>
      End
      <input class="dt-end" type="text" value="10:15">
    </label>

    <label>
      Full Day
      <select class="dt-fullday">
        <option value="No">No</option>
        <option value="Yes">Yes</option>
      </select>
    </label>

    <label>
      Notes
      <input class="dt-notes" type="text" placeholder="Reason / machine #">
    </label>

    <label>
      Active
      <select class="dt-active">
        <option value="No">No</option>
        <option value="Yes">Yes</option>
      </select>
    </label>
  `;

  holder.appendChild(row);

  row.querySelectorAll('input, select').forEach(function (el) {
    el.addEventListener('input', calculateForecast);
    el.addEventListener('change', calculateForecast);
  });

  if (!isRestoringState) {
    calculateForecast();
    savePageState();
  }

  return row;
}

function clearDowntimeRows() {
  const holder = document.getElementById('downtimeRows');
  if (!holder) return;

  holder.innerHTML = '';
  downtimeCounter = 0;
  addDowntimeRow();
  calculateForecast();
  savePageState();
}

function clearHourlyRows() {
  document.querySelectorAll('.hourly-row').forEach(function (row) {
    row.querySelectorAll('input').forEach(function (input) {
      if (input.classList.contains('hr-hour')) return;
      input.value = '';
    });
    const active = row.querySelector('.hr-active');
    if (active) active.value = 'Yes';
  });

  calculateForecast();
  savePageState();
}

function collectPayload() {
  const areas = {};
  const associates = {};
  const ratios = {};

  SURFACE_AREAS.forEach(function (area) {
    areas[area.key] = Number(getValue(`area_${area.key}`, 0)) || 0;
    associates[area.key] = Number(getValue(`assoc_${area.key}`, 0)) || 0;
    ratios[area.key] = area.type === 'machine'
      ? (Number(getValue(`ratio_${area.key}`, area.machinesPerAssociate)) || area.machinesPerAssociate || 1)
      : 0;
  });

  const downtime = [];

  document.querySelectorAll('.downtime-row').forEach(function (row) {
    downtime.push({
      areaKey: getRowValue(row, '.dt-area'),
      issue: getRowValue(row, '.dt-issue'),
      unitsDown: Number(getRowValue(row, '.dt-units')) || 0,
      startTime: getRowValue(row, '.dt-start'),
      endTime: getRowValue(row, '.dt-end'),
      fullDay: getRowValue(row, '.dt-fullday'),
      notes: getRowValue(row, '.dt-notes'),
      active: getRowValue(row, '.dt-active')
    });
  });

  const hourly = [];

  document.querySelectorAll('.hourly-row').forEach(function (row) {
    hourly.push({
      hour: row.dataset.apiHour || normalizeHourKey(getRowValue(row, '.hr-hour')),
      unbox: Number(getRowValue(row, '.hr-unbox')) || 0,
      autoblocker: Number(getRowValue(row, '.hr-autoblocker')) || 0,
      cooling: Number(getRowValue(row, '.hr-cooling')) || 0,
      orb: Number(getRowValue(row, '.hr-orb')) || 0,
      polisher: Number(getRowValue(row, '.hr-polisher')) || 0,
      engraving: Number(getRowValue(row, '.hr-engraving')) || 0,
      detaping: Number(getRowValue(row, '.hr-detaping')) || 0,
      coater: Number(getRowValue(row, '.hr-coater')) || 0,
      inspection: Number(getRowValue(row, '.hr-inspection')) || 0,
      notes: getRowValue(row, '.hr-notes'),
      active: getRowValue(row, '.hr-active')
    });
  });

  const currentWip = window.surfaceDashboardCurrentWip || {};
  const dashboardSync = window.surfaceDashboardSync || null;

  return {
    productionDate: getValue('productionDate', getLocalDateInputValue(new Date())),
    shiftMode: getValue('shiftMode', 'Weekday'),
    bosWip: Number(getValue('bosWip', 0)) || 0,
    shipGoal: Number(getValue('shipGoal', 0)) || 0,
    incomingLow: Number(getValue('incomingLow', SURFACE_EXTRA_WORK_DEFAULTS.incomingLow)) || 0,
    incomingHigh: Number(getValue('incomingHigh', SURFACE_EXTRA_WORK_DEFAULTS.incomingHigh)) || 0,
    remakeLow: Number(getValue('remakeLow', SURFACE_EXTRA_WORK_DEFAULTS.remakeLow)) || 0,
    remakeHigh: Number(getValue('remakeHigh', SURFACE_EXTRA_WORK_DEFAULTS.remakeHigh)) || 0,
    floaterCount: Number(getValue('floaterCount', 0)) || 0,
    floaterAssign: getValue('floaterAssign', 'none'),
    areas,
    associates,
    ratios,
    downtime,
    hourly,
    currentWip,
    dashboardSync
  };
}

function calculateForecast() {
  const payload = collectPayload();
  const result = calculateLocalForecast(payload);

  latestForecastResult = { payload, result };
  renderForecast(result);

  if (!isRestoringState) {
    savePageState();
  }
}


const SURFACE_CAPACITY_ROUTE_EXCLUSIONS = new Set([
  'cooling',
  'detaping'
]);

function getSurfaceCapacityRoute(sourceKey, includeSource) {
  const sourceIndex = SURFACE_FLOW_ORDER.indexOf(sourceKey);
  if (sourceIndex < 0) return [];

  let route = includeSource
    ? SURFACE_FLOW_ORDER.slice(sourceIndex)
    : SURFACE_FLOW_ORDER.slice(sourceIndex + 1);

  // Cooling is a dwell-time buffer and Detaping is not a reliable hard-cap
  // station because jobs may move through manual deblock without an ODT scan.
  // Both remain visible for WIP pressure and timing, but neither caps forecast OUT.
  route = route.filter(function (key) {
    return !SURFACE_CAPACITY_ROUTE_EXCLUSIONS.has(key);
  });

  // Surface Inspection WIP is already at the final output station.
  if (sourceKey === 'inspection' && !route.length) {
    route = ['inspection'];
  }

  return route;
}

function getSurfaceRouteCapacity(capacityMap, sourceKey, includeSource) {
  const route = getSurfaceCapacityRoute(sourceKey, includeSource);
  if (!route.length) return 0;

  return route.reduce(function (lowest, key) {
    const value = Math.max(0, Number(capacityMap[key]) || 0);
    return Math.min(lowest, value);
  }, Infinity);
}

function buildSurfaceRemainingCapacityMap(areas, remainingHours) {
  const map = {};

  (areas || []).forEach(function (area) {
    if (area.key === 'cooling') {
      map[area.key] = Infinity;
      return;
    }

    map[area.key] = Math.max(
      0,
      (Number(area.effectiveJph) || 0) * Math.max(0, remainingHours)
    );
  });

  return map;
}

function cloneSurfaceCapacityMap(source) {
  const copy = {};

  Object.keys(source || {}).forEach(function (key) {
    copy[key] = source[key] === Infinity
      ? Infinity
      : Math.max(0, Number(source[key]) || 0);
  });

  return copy;
}

function consumeCompletedSurfaceBos(bosByStation, completedOutput) {
  const remaining = normalizeSurfaceBosMap(bosByStation || {});
  let jobsToConsume = Math.max(0, Number(completedOutput) || 0);

  // Morning output normally comes from the most downstream BOS first.
  SURFACE_FLOW_ORDER.slice().reverse().forEach(function (key) {
    if (jobsToConsume <= 0) return;

    const available = Math.max(0, Number(remaining[key]) || 0);
    const consumed = Math.min(available, jobsToConsume);

    remaining[key] = available - consumed;
    jobsToConsume -= consumed;
  });

  return remaining;
}

function allocateSurfaceWorkPools(pools, startingCapacityMap, options) {
  const settings = options || {};
  const capacityMap = cloneSurfaceCapacityMap(startingCapacityMap);
  const allocations = {};
  let totalAllocated = 0;

  const orderedPools = (pools || []).slice().sort(function (a, b) {
    // Protect work that is already furthest downstream.
    return getSurfaceFlowOrderIndex(b.sourceKey) -
      getSurfaceFlowOrderIndex(a.sourceKey);
  });

  orderedPools.forEach(function (pool) {
    const sourceKey = String(pool.sourceKey || '');
    const requestedJobs = Math.max(0, Number(pool.jobs) || 0);
    const includeSource = Boolean(pool.includeSource);

    if (!requestedJobs) {
      allocations[sourceKey] = 0;
      return;
    }

    const maturityMinutes = includeSource
      ? Number(SURFACE_MATURITY_MINUTES.unbox || 0)
      : Number(SURFACE_MATURITY_MINUTES[sourceKey] || 0);

    if (
      Number.isFinite(Number(settings.availableMinutes)) &&
      Number(settings.availableMinutes) < maturityMinutes
    ) {
      allocations[sourceKey] = 0;
      return;
    }

    const route = getSurfaceCapacityRoute(sourceKey, includeSource);
    if (!route.length) {
      allocations[sourceKey] = 0;
      return;
    }

    const routeCapacity = route.reduce(function (lowest, key) {
      const available = capacityMap[key] === Infinity
        ? Infinity
        : Math.max(0, Number(capacityMap[key]) || 0);

      return Math.min(lowest, available);
    }, Infinity);

    const allocated = Math.max(
      0,
      Math.min(
        requestedJobs,
        Number.isFinite(routeCapacity) ? routeCapacity : requestedJobs
      )
    );

    allocations[sourceKey] = allocated;
    totalAllocated += allocated;

    route.forEach(function (key) {
      if (capacityMap[key] === Infinity) return;
      capacityMap[key] = Math.max(
        0,
        (Number(capacityMap[key]) || 0) - allocated
      );
    });
  });

  return {
    totalAllocated,
    allocations,
    remainingCapacityMap: capacityMap
  };
}

function buildSurfaceBosPools(bosByStation) {
  return SURFACE_FLOW_ORDER.map(function (key) {
    return {
      sourceKey: key,
      jobs: Math.max(0, Number(bosByStation && bosByStation[key]) || 0),
      includeSource: false
    };
  });
}

function calculateLocalForecast(payload) {
  const shiftMode = payload.shiftMode || 'Weekday';
  const shiftHours = getShiftHours(shiftMode);
  const bosByStation = captureSurfaceBosSnapshotIfMissing(
    payload.productionDate,
    payload.currentWip || {}
  );
  const bosWip = getSurfaceBosTotal(bosByStation);
  const shipGoal = Number(payload.shipGoal) || 0;
  const incomingRange = normalizeForecastRange(payload.incomingLow, payload.incomingHigh);
  const remakeRange = normalizeForecastRange(payload.remakeLow, payload.remakeHigh);

  // Daily work pool:
  // Low = BOS + conservative incoming + conservative breakage/remake.
  // High = BOS + upper incoming + upper breakage/remake.
  const availableWorkLow =
    bosWip + incomingRange.low + remakeRange.low;
  const availableWorkHigh =
    bosWip + incomingRange.high + remakeRange.high;
  const floaterCount = Number(payload.floaterCount) || 0;
  const floaterAssign = payload.floaterAssign || 'none';

  const areas = SURFACE_AREAS.map(function (area) {
    let unitsRunning = Number(payload.areas[area.key]) || 0;
    let associatesAssigned = Number(payload.associates[area.key]) || 0;

    // Floaters are people, not machines. Add them to the assigned associate side.
    if (floaterAssign === area.key && floaterCount > 0) {
      associatesAssigned += floaterCount;
    }

    const capacityInfo = getAreaCapacityInfo(area, unitsRunning, associatesAssigned, shiftHours);
    const downtimeLostJobs = getDowntimeLossForArea(area, payload.downtime || [], shiftHours);
    const normalCapacity = capacityInfo.normalCapacity;
    const adjustedCapacity = Math.max(0, normalCapacity - downtimeLostJobs);
    const currentWip = Number((payload.currentWip || {})[area.key]) || 0;
    const effectiveJph = shiftHours > 0 ? adjustedCapacity / shiftHours : 0;
    const timeToClear = effectiveJph > 0 ? currentWip / effectiveJph : 0;

    return {
      key: area.key,
      label: area.label,
      group: area.group,
      type: area.type,
      capacityMode: area.capacityMode || area.type,
      jphPerUnit: area.jphPerUnit,
      required: area.required,
      bufferCapacity: area.bufferCapacity || 0,
      coolingLowMin: area.coolingLowMin || 0,
      coolingHighMin: area.coolingHighMin || 0,
      unitsRunning,
      associatesAssigned,
      activeCapacityUnits: capacityInfo.activeCapacityUnits,
      capacityBasis: capacityInfo.capacityBasis,
      machineCeilingJph: Number(capacityInfo.machineCeilingJph || capacityInfo.totalJph || 0),
      requiredAssociates: Number(capacityInfo.requiredAssociates || 0),
      coverageFactor: capacityInfo.coverageFactor === undefined ? 1 : Number(capacityInfo.coverageFactor || 0),
      totalJph: capacityInfo.totalJph,
      normalCapacity,
      downtimeLostJobs,
      adjustedCapacity,
      currentWip,
      effectiveJph,
      timeToClear,
      forecastCapacityRole:
        area.key === 'detaping'
          ? 'wip-pressure-only'
          : area.key === 'cooling'
            ? 'dwell-time-only'
            : 'hard-cap'
    };
  });

  const requiredAreas = areas.filter(function (area) {
    return (
      area.required &&
      area.effectiveJph > 0 &&
      area.key !== 'detaping'
    );
  });

  let bottleneck = null;

  if (requiredAreas.length) {
    bottleneck = requiredAreas.reduce(function (lowest, current) {
      return current.adjustedCapacity < lowest.adjustedCapacity ? current : lowest;
    }, requiredAreas[0]);
  }

  const capacityCeiling = bottleneck ? bottleneck.adjustedCapacity : 0;
  const effectiveBottleneckJph = shiftHours > 0 ? capacityCeiling / shiftHours : 0;

  const downtimeLostTotal = areas.reduce(function (sum, area) {
    return sum + area.downtimeLostJobs;
  }, 0);

  const totalAssociates = areas.reduce(function (sum, area) {
    return sum + area.associatesAssigned;
  }, 0);

  const hourlySummary = calculateHourlySummary(
    payload.hourly || [],
    shiftHours,
    effectiveBottleneckJph,
    shipGoal,
    capacityCeiling,
    availableWorkLow,
    availableWorkHigh
  );
  const liveWipBottleneck = getLiveWipBottleneck(areas);

  const currentOut = Number(hourlySummary.actualAr41SoFar || 0);
  const progress = getSurfaceClockProgress(shiftMode);
  const maturity = getSurfaceBosMaturity(
    bosByStation,
    progress.elapsedMinutes,
    progress.totalMinutes
  );

  const remainingHours = progress.remainingMinutes / 60;
  const remainingCapacityMap = buildSurfaceRemainingCapacityMap(
    areas,
    remainingHours
  );

  // Remove completed output from BOS beginning with the most downstream work.
  // This prevents Current OUT from being added on top of the same BOS jobs twice.
  const remainingBosByStation = consumeCompletedSurfaceBos(
    bosByStation,
    currentOut
  );

  // Allocate remaining BOS from downstream to upstream. Each pool is limited
  // only by the stations it still must pass, not by an upstream station it
  // already cleared.
  const bosAllocation = allocateSurfaceWorkPools(
    buildSurfaceBosPools(remainingBosByStation),
    remainingCapacityMap,
    {
      // BOS was present at shift start, so use the full shift maturity window.
      availableMinutes: progress.totalMinutes
    }
  );

  const bosAdditionalOut = bosAllocation.totalAllocated;
  const knownWipForecast = Math.round(
    currentOut + bosAdditionalOut
  );

  const lowExtraWork = Math.max(
    0,
    incomingRange.low + remakeRange.low
  );
  const highExtraWork = Math.max(
    0,
    incomingRange.high + remakeRange.high
  );

  // New incoming and breakage/remake work enters the full Surface route.
  // Run low and high scenarios independently from the capacity remaining
  // after BOS so the range represents two real workload cases.
  const lowExtraAllocation = allocateSurfaceWorkPools(
    [{
      sourceKey: 'unbox',
      jobs: lowExtraWork,
      includeSource: true
    }],
    bosAllocation.remainingCapacityMap,
    {
      availableMinutes: progress.remainingMinutes
    }
  );

  const highExtraAllocation = allocateSurfaceWorkPools(
    [{
      sourceKey: 'unbox',
      jobs: highExtraWork,
      includeSource: true
    }],
    bosAllocation.remainingCapacityMap,
    {
      availableMinutes: progress.remainingMinutes
    }
  );

  // Low case uses conservative execution: 90% of the additional route-capacity
  // result after Current OUT. High case uses the full route-capacity result.
  // This keeps the range useful even when low/high workload both hit the same ceiling.
  const lowCaseFull =
    currentOut +
    bosAdditionalOut +
    lowExtraAllocation.totalAllocated;

  const highCaseFull =
    currentOut +
    bosAdditionalOut +
    highExtraAllocation.totalAllocated;

  const forecastRangeLow = Math.round(
    currentOut + Math.max(0, lowCaseFull - currentOut) * 0.90
  );

  const forecastRangeHigh = Math.round(
    Math.max(
      forecastRangeLow + 1,
      highCaseFull
    )
  );

  const paceProjection = Number(
    hourlySummary.projectedActualOut || 0
  );

  // Pace selects the working projection, but it can never fall below the
  // conservative route-capacity case or exceed the high workload case.
  const projectedOut = Math.round(
    Math.max(
      forecastRangeLow,
      Math.min(
        forecastRangeHigh,
        paceProjection > currentOut
          ? paceProjection
          : (forecastRangeLow + forecastRangeHigh) / 2
      )
    )
  );

  const capacityProjectedOut = forecastRangeHigh;
  const remainingCapacity = getSurfaceRouteCapacity(
    bosAllocation.remainingCapacityMap,
    'unbox',
    true
  );

  const bosRemainingAfterActual = Math.max(
    0,
    bosWip - currentOut
  );
  const availableWorkRemainingLow = Math.max(
    0,
    availableWorkLow - currentOut
  );
  const availableWorkRemainingHigh = Math.max(
    0,
    availableWorkHigh - currentOut
  );

  const extraCapacityForUnknownWork = Math.max(
    0,
    Number.isFinite(remainingCapacity)
      ? remainingCapacity
      : 0
  );

  const expectedEosWipIfNoExtraWork = Math.max(
    0,
    bosWip - knownWipForecast
  );

  // During the morning startup window, expected output begins only after the
  // first physically possible downstream release.
  const productiveElapsedAfterStartup = Math.max(
    0,
    progress.elapsedMinutes -
      (maturity.startupDelayApplied
        ? SURFACE_STARTUP_OUTPUT_DELAY_MIN
        : 0)
  );
  const expectedHoursAfterStartup =
    productiveElapsedAfterStartup / 60;
  const targetPaceForExpected = shipGoal > 0
    ? shipGoal / shiftHours
    : effectiveBottleneckJph;

  hourlySummary.expectedByNow =
    targetPaceForExpected * expectedHoursAfterStartup;
  hourlySummary.paceGap =
    currentOut - hourlySummary.expectedByNow;
  hourlySummary.remainingHours = remainingHours;
  hourlySummary.projectedActualOut = projectedOut;

  let requiredJph = 0;
  let goalRisk = 'NO GOAL';
  let goalGap = 0;

  if (shipGoal > 0) {
    requiredJph = shipGoal / shiftHours;
    goalGap = shipGoal - forecastRangeHigh;

    if (shipGoal <= forecastRangeLow) goalRisk = 'GREEN';
    else if (shipGoal <= forecastRangeHigh) goalRisk = 'YELLOW';
    else goalRisk = 'RED';
  }

  const recommendation = buildRecommendation({
    shiftMode,
    shiftHours,
    bosWip,
    bosByStation,
    bosMaturity: maturity,
    startupDelayRemaining: maturity.startupDelayRemaining,
    remainingCapacity,
    remainingCapacityMap,
    remainingBosByStation,
    bosRouteAdditionalOut: bosAdditionalOut,
    lowExtraProjected: lowExtraAllocation.totalAllocated,
    highExtraProjected: highExtraAllocation.totalAllocated,
    incomingLow: incomingRange.low,
    incomingHigh: incomingRange.high,
    remakeLow: remakeRange.low,
    remakeHigh: remakeRange.high,
    availableWorkLow,
    availableWorkHigh,
    bosRemainingAfterActual,
    availableWorkRemainingLow,
    availableWorkRemainingHigh,
    shipGoal,
    knownWipForecast,
    capacityCeiling,
    capacityProjectedOut,
    projectedOut,
    extraCapacityForUnknownWork,
    expectedEosWipIfNoExtraWork,
    bottleneck,
    effectiveBottleneckJph,
    requiredJph,
    goalRisk,
    downtimeLostTotal,
    totalAssociates,
    hourlySummary,
    liveWipBottleneck
  });

  const goalStaffingPlan = calculateGoalStaffingPlan(payload, areas, shiftHours, shipGoal);

  return {
    ok: true,
    shiftMode,
    shiftHours,
    bosWip,
    bosByStation,
    remainingBosByStation,
    bosRouteAdditionalOut: bosAdditionalOut,
    lowExtraProjected: lowExtraAllocation.totalAllocated,
    highExtraProjected: highExtraAllocation.totalAllocated,
    remainingCapacityMap,
    incomingLow: incomingRange.low,
    incomingHigh: incomingRange.high,
    remakeLow: remakeRange.low,
    remakeHigh: remakeRange.high,
    availableWorkLow,
    availableWorkHigh,
    bosRemainingAfterActual,
    availableWorkRemainingLow,
    availableWorkRemainingHigh,
    shipGoal,
    knownWipForecast,
    projectedOut,
    capacityProjectedOut,
    capacityCeiling,
    forecastRangeLow,
    forecastRangeHigh,
    extraCapacityForUnknownWork,
    expectedEosWipIfNoExtraWork,
    bottleneck,
    effectiveBottleneckJph,
    downtimeLostTotal,
    totalAssociates,
    requiredJph,
    goalRisk,
    goalGap,
    areas,
    hourlySummary,
    liveWipBottleneck,
    goalStaffingPlan,
    recommendation
  };
}

function roundUpToStep(value, step) {
  if (!(value > 0) || !(step > 0)) return 0;
  return Math.ceil(value / step) * step;
}

// Reverses the forecast: given the Ship Goal, what does EVERY required area need
// (machines and associates) to hold that pace for the full shift — not just the
// single current bottleneck. Machine areas are fully automatic; associates clear
// errors and keep WIP flowing rather than physically operating each machine. So
// an "AT RISK" area here means it's running above sustainable error-clearing
// coverage, not that a machine literally cannot run — treat it as a downtime-risk
// flag. (Follow-up: tie this into the real Downtime Loss / WIP mechanism instead
// of a flat gap number — not done yet.)
function calculateGoalStaffingPlan(payload, areas, shiftHours, shipGoal) {
  if (!(shipGoal > 0) || !(shiftHours > 0)) return null;

  const requiredJphOverall = shipGoal / shiftHours;

  const rows = SURFACE_AREAS.filter(function (area) {
    return area.required;
  }).map(function (area) {
    const liveArea = areas.find(function (a) { return a.key === area.key; }) || {};
    const requiredRawUnits = area.jphPerUnit > 0 ? requiredJphOverall / area.jphPerUnit : 0;

    if (area.type === 'operator') {
      const requiredAssociates = roundUpToStep(requiredRawUnits, 0.5);
      const currentAssociates = Math.max(Number(payload.associates[area.key]) || 0, Number(payload.areas[area.key]) || 0);
      const associateGap = Math.max(0, requiredAssociates - currentAssociates);

      return {
        key: area.key,
        label: area.label,
        type: area.type,
        jphPerUnit: area.jphPerUnit,
        requiredUnits: requiredAssociates,
        currentUnits: currentAssociates,
        unitGap: associateGap,
        requiredAssociates,
        currentAssociates,
        associateGap,
        onTrack: associateGap <= 0
      };
    }

    const requiredMachines = roundUpToStep(requiredRawUnits, 0.5);
    const currentMachines = Number(payload.areas[area.key]) || 0;
    const machineGap = Math.max(0, requiredMachines - currentMachines);
    const ratio = Number(payload.ratios[area.key]) || area.machinesPerAssociate || 1;
    const requiredAssociates = roundUpToStep(requiredMachines / ratio, 0.5);
    const currentAssociates = Number(payload.associates[area.key]) || 0;
    const associateGap = Math.max(0, requiredAssociates - currentAssociates);

    return {
      key: area.key,
      label: area.label,
      type: area.type,
      jphPerUnit: area.jphPerUnit,
      requiredUnits: requiredMachines,
      currentUnits: currentMachines,
      unitGap: machineGap,
      requiredAssociates,
      currentAssociates,
      associateGap,
      onTrack: machineGap <= 0 && associateGap <= 0
    };
  });

  const totalAdditionalAssociates = rows.reduce(function (sum, row) { return sum + row.associateGap; }, 0);
  const totalAdditionalMachines = rows.reduce(function (sum, row) {
    return sum + (row.type === 'machine' ? row.unitGap : 0);
  }, 0);
  const shortAreas = rows.filter(function (row) { return !row.onTrack; });
  const operatorShortAreas = shortAreas.filter(function (row) { return row.type === 'operator'; });
  const machineRiskAreas = shortAreas.filter(function (row) { return row.type === 'machine'; });

  return {
    requiredJphOverall,
    rows,
    totalAdditionalAssociates,
    totalAdditionalMachines,
    shortAreaCount: shortAreas.length,
    operatorShortCount: operatorShortAreas.length,
    machineRiskCount: machineRiskAreas.length,
    achievable: shortAreas.length === 0
  };
}

function calculateHourlySummary(hourly, shiftHours, effectiveBottleneckJph, shipGoal, capacityCeiling, availableWorkLow, availableWorkHigh) {
  const totals = {
    unbox: 0,
    autoblocker: 0,
    cooling: 0,
    orb: 0,
    polisher: 0,
    engraving: 0,
    detaping: 0,
    coater: 0,
    inspection: 0
  };

  let hoursEntered = 0;

  hourly.forEach(function (row) {
    if (!row || row.active === 'No') return;

    const anyValue =
      Number(row.unbox || 0) +
      Number(row.autoblocker || 0) +
      Number(row.cooling || 0) +
      Number(row.orb || 0) +
      Number(row.polisher || 0) +
      Number(row.engraving || 0) +
      Number(row.detaping || 0) +
      Number(row.coater || 0) +
      Number(row.inspection || 0);

    if (anyValue <= 0) return;

    totals.unbox += Number(row.unbox || 0);
    totals.autoblocker += Number(row.autoblocker || 0);
    totals.cooling += Number(row.cooling || 0);
    totals.orb += Number(row.orb || 0);
    totals.polisher += Number(row.polisher || 0);
    totals.engraving += Number(row.engraving || 0);
    totals.detaping += Number(row.detaping || 0);
    totals.coater += Number(row.coater || 0);
    totals.inspection += Number(row.inspection || 0);

    if (Number(row.inspection || 0) > 0) {
      hoursEntered += 1;
    }
  });

  const actualAr41SoFar = totals.inspection;
  const actualPace = hoursEntered > 0 ? actualAr41SoFar / hoursEntered : 0;
  const remainingHours = Math.max(0, shiftHours - hoursEntered);

  const targetPace = shipGoal > 0
    ? shipGoal / shiftHours
    : effectiveBottleneckJph;

  const expectedByNow = targetPace * hoursEntered;
  const paceGap = actualAr41SoFar - expectedByNow;

  let onTrackStatus = 'NO HOURLY';

  if (hoursEntered > 0) {
    if (actualAr41SoFar >= expectedByNow) onTrackStatus = 'GREEN';
    else if (actualAr41SoFar >= expectedByNow * 0.90) onTrackStatus = 'YELLOW';
    else onTrackStatus = 'RED';
  }

  // Available work = BOS WIP + average incoming Surface work + average remake/breakage work.
  // Current OUT is already part of today's available work pool.
  const availableRemainingLow = Math.max(0, Number(availableWorkLow || 0) - actualAr41SoFar);
  const availableRemainingHigh = Math.max(0, Number(availableWorkHigh || availableWorkLow || 0) - actualAr41SoFar);
  const capacityRemainingAfterActual = Math.max(0, Number(capacityCeiling || 0) - actualAr41SoFar);
  const capacityProjectedOutLow = actualAr41SoFar + Math.min(availableRemainingLow, capacityRemainingAfterActual);
  const capacityProjectedOutHigh = actualAr41SoFar + Math.min(availableRemainingHigh, capacityRemainingAfterActual);

  const currentPaceProjectedOut = hoursEntered > 0
    ? actualAr41SoFar + (actualPace * remainingHours)
    : capacityProjectedOutLow;

  const projectedActualOut = Math.min(capacityProjectedOutHigh, currentPaceProjectedOut);

  return {
    totals,
    hoursEntered,
    actualAr41SoFar,
    actualPace,
    targetPace,
    expectedByNow,
    paceGap,
    remainingHours,
    onTrackStatus,
    availableRemainingLow,
    availableRemainingHigh,
    capacityProjectedOutLow,
    capacityProjectedOutHigh,
    capacityProjectedOut: capacityProjectedOutHigh,
    projectedActualOut
  };
}

function renderForecast(result) {
  const hourly = result.hourlySummary || {};
  const status = hourly.onTrackStatus && hourly.onTrackStatus !== 'NO HOURLY' ? hourly.onTrackStatus : result.goalRisk;
  const nowText = `Last calculated: ${new Date().toLocaleTimeString()}`;
  const shipGoal = Number(result.shipGoal || 0);
  const projected = Number(result.projectedOut ?? hourly.projectedActualOut ?? 0);
  const expected = Number(hourly.expectedByNow || 0);
  const actual = Number(hourly.actualAr41SoFar || 0);
  const paceRatio = expected > 0 ? actual / expected : 0;
  const pacePercent = expected > 0 ? paceRatio * 100 : 0;
  const paceGap = actual - expected;
  const live = result.liveWipBottleneck || null;
  const real = result.realBottleneck || null;
  const structural = result.bottleneck || real || null;
  const downtimeArea = getTopDowntimeImpactArea(result.areas || []);
  const actionPlan = buildGoalActionPlan(result);

  renderCommandKpis(result, actionPlan);
  renderSurfaceBosSnapshot();
  updatePerformanceUI(result);
  renderGoalStaffingPlan(result);

  setText('forecastStatus', status || 'READY');
  const displayRangeLow = Math.max(
    actual,
    Number(result.forecastRangeLow ?? result.knownWipForecast ?? projected ?? 0)
  );
  const displayRangeHigh = Math.max(
    displayRangeLow,
    Number(result.forecastRangeHigh ?? result.capacityProjectedOut ?? projected ?? 0)
  );

  setText(
    'forecastRange',
    `${formatNumber(displayRangeLow)} - ${formatNumber(displayRangeHigh)}`
  );
  setText('shipGoalDisplay', shipGoal > 0 ? formatNumber(shipGoal) : '0');
  setText('projectedGap', buildProjectedGapText(shipGoal, projected));
  setText('pacePercent', expected > 0 ? `${round1(pacePercent)}%` : '0%');
  setText('paceGap', `${paceGap >= 0 ? '+' : ''}${formatNumber(paceGap)} jobs`);
  setText('lastUpdated', nowText);

  setText(
    'summaryRangeValue',
    `${formatNumber(displayRangeLow)} - ${formatNumber(displayRangeHigh)}`
  );
  setText('summaryProjectedOutValue', formatNumber(projected));
  setText('summaryTargetValue', shipGoal > 0 ? formatNumber(shipGoal) : '0');
  setText('summaryDowntimeValue', formatNumber(result.downtimeLostTotal));
  setText(
    'summaryFootnote',
    result.startupDelayRemaining > 0
      ? `${Math.round(result.startupDelayRemaining)} min startup delay remains before normal Surface OUT. ` +
        buildSummaryFootnote(result, actionPlan)
      : buildSummaryFootnote(result, actionPlan)
  );

  setText('topPaceValue', expected > 0 ? `${round1(pacePercent)}%` : '0%');
  setText('topPaceSub', hourly.onTrackStatus && hourly.onTrackStatus !== 'NO HOURLY' ? 'of expected' : 'waiting on hourly data');
  setText('topGoalValue', shipGoal > 0 ? `${actionPlan.extraJobsNeeded > 0 ? '+' : ''}${formatNumber(actionPlan.extraJobsNeeded)}` : 'NO GOAL');
  setText('topGoalSub', shipGoal > 0 ? (actionPlan.extraJobsNeeded > 0 ? 'jobs still needed' : 'goal covered') : 'enter ship goal');

  setText('focusStructuralName', structural ? structural.label : '-');
  setText('focusStructuralMeta', structural ? buildStructuralMeta(structural) : 'No structural bottleneck yet');
  setText('focusLiveWipName', live ? live.label : '-');
  setText('focusLiveWipMeta', live ? buildLivePressureMeta(live) : 'No live WIP pressure');
  setText('focusDowntimeAreaName', downtimeArea ? downtimeArea.label : '-');
  setText('focusDowntimeAreaMeta', downtimeArea ? `${formatNumber(downtimeArea.downtimeLostJobs || 0)} jobs lost` : 'No downtime recorded');
  setText('focusDowntimeLossValue', `${formatNumber(result.downtimeLostTotal)} jobs`);

  setText('actionExtraJobsNeeded', shipGoal > 0 ? `+${formatNumber(actionPlan.extraJobsNeeded)} jobs` : 'No target');
  setText('actionExtraJphNeeded', shipGoal > 0 ? `+${round1(actionPlan.extraJphNeeded)} JPH` : 'No target');
  setText('actionPrimaryFocusArea', actionPlan.primaryFocusArea || '-');
  setText('actionDowntimeImpact', `${formatNumber(result.downtimeLostTotal)} jobs`);
  setText('actionRecommendedAction', actionPlan.recommendedAction);
  setText('actionBottomLine', actionPlan.bottomLine);

  setText('knownWipForecast', formatNumber(result.knownWipForecast));
  setText('capacityCeiling', formatNumber(result.capacityCeiling));
  setText('actualAr41', formatNumber(actual));
  setText('expectedByNow', formatNumber(expected));
  setText('onTrackStatus', hourly.onTrackStatus || 'NO HOURLY');
  setText('projectedActualOut', formatNumber(projected));
  setText('actualPace', round1(hourly.actualPace || 0));
  setText('bottleneckArea', real ? real.label : (structural ? structural.label : '-'));
  setText('goalRisk', result.goalRisk || 'NO GOAL');
  setText('goalRiskSub', buildGoalRiskSub(result, projected));
  setText('shiftModeSetupText', getShiftModeDisplay(result.shiftMode));

  const realPressure = real || live;
  setText('resultLiveWipBottleneck', realPressure ? realPressure.label : '-');
  setText(
    'resultLiveWipHours',
    realPressure && realPressure.pressureScore !== undefined
      ? `${round1(realPressure.pressureScore)} hr pressure · ${round1(realPressure.wipHours || 0)} hr WIP · ${round1(realPressure.recoveryHours || 0)} hr recovery`
      : (realPressure ? `${round1(realPressure.timeToClear)} hr to clear` : '0 hr to clear')
  );
  setText('liveWipBottleneck', live ? live.label : '—');
  setText('liveWipBottleneckSub', live ? buildLivePressureMeta(live) : 'Waiting for WIP.');

  setText('effectiveJph', round1(result.effectiveBottleneckJph));
  setText('effectiveJphMirror', round1(result.effectiveBottleneckJph));
  setText('totalAssociates', round1(result.totalAssociates));
  setText('downtimeLost', formatNumber(result.downtimeLostTotal));
  setText('recommendation', result.recommendation);

  setRiskClass(document.getElementById('forecastStatus'), status);
  setRiskClass(document.getElementById('onTrackStatus'), hourly.onTrackStatus || 'NO HOURLY');
  setRiskClass(document.getElementById('goalRisk'), result.goalRisk || 'NO GOAL');
  setRiskClass(document.getElementById('paceGap'), paceGap >= 0 ? 'GREEN' : (paceRatio >= 0.90 ? 'YELLOW' : 'RED'));
  setRiskClass(document.getElementById('projectedGap'), shipGoal <= 0 ? 'NO GOAL' : (projected >= shipGoal ? 'GREEN' : 'RED'));
  setRiskClass(document.getElementById('topPaceValue'), paceRatio >= 1 ? 'GREEN' : (paceRatio >= 0.90 ? 'YELLOW' : 'RED'));
  setRiskClass(document.getElementById('topGoalValue'), shipGoal <= 0 ? 'NO GOAL' : (actionPlan.extraJobsNeeded <= 0 ? 'GREEN' : 'RED'));
  setRiskClass(document.getElementById('summaryDowntimeValue'), result.downtimeLostTotal > 0 ? 'RED' : 'GREEN');
  setRiskClass(document.getElementById('focusDowntimeLossValue'), result.downtimeLostTotal > 0 ? 'RED' : 'GREEN');

  setPaceVisuals(paceRatio, actual, expected);

  const hasHourly = hourly.onTrackStatus && hourly.onTrackStatus !== 'NO HOURLY';
  setText('forecastStatusBasis', hasHourly ? 'Based on hourly pace vs. target' : 'Based on ship goal vs. capacity');

  renderHourlyChart(latestForecastResult ? latestForecastResult.payload.hourly : [], hourly, result);
  renderBottleneckTable(result.areas);
}


function renderGoalStaffingPlan(result) {
  const summaryHolder = document.getElementById('goalPlanSummary');
  const tableWrap = document.getElementById('goalPlanTableWrap');
  const tableBody = document.getElementById('goalPlanTable');
  if (!summaryHolder || !tableWrap || !tableBody) return;

  const plan = result.goalStaffingPlan;

  if (!plan) {
    summaryHolder.innerHTML = '<p class="muted-cell">Set a Ship Goal above 0 to see the staffing plan.</p>';
    tableWrap.hidden = true;
    return;
  }

  tableWrap.hidden = false;

  if (plan.achievable) {
    summaryHolder.innerHTML = `<p class="goal-plan-ok">Every required area is staffed to hit ${formatNumber(result.shipGoal)} at ${round1(plan.requiredJphOverall)} JPH pace. No additional people or machines needed.</p>`;
  } else {
    const parts = [];
    if (plan.operatorShortCount > 0) {
      parts.push(`<strong>${plan.operatorShortCount}</strong> area(s) genuinely short-staffed to hit pace`);
    }
    if (plan.machineRiskCount > 0) {
      parts.push(`<strong>${plan.machineRiskCount}</strong> automatic area(s) running above sustainable error-clearing coverage (downtime risk, not a hard stop)`);
    }
    summaryHolder.innerHTML = `<p class="goal-plan-short">${parts.join(' and ')} — add <strong>${formatNumber(plan.totalAdditionalAssociates)}</strong> associate(s)${plan.totalAdditionalMachines > 0 ? ` and <strong>${formatNumber(plan.totalAdditionalMachines)}</strong> machine(s)` : ''} to comfortably hit ${formatNumber(result.shipGoal)} at ${round1(plan.requiredJphOverall)} JPH pace.</p>`;
  }

  tableBody.innerHTML = plan.rows.map(function (row) {
    const unitLabel = row.type === 'operator' ? 'associates' : 'machines';
    const statusText = row.onTrack ? 'OK' : (row.type === 'operator' ? 'SHORT' : 'AT RISK');
    return `
      <tr class="${row.onTrack ? 'row-ok' : (row.type === 'operator' ? 'row-short' : 'row-risk')}">
        <td><strong>${escapeHtml(row.label)}</strong></td>
        <td>${row.jphPerUnit}</td>
        <td>${formatNumber(row.requiredUnits)} ${unitLabel}</td>
        <td>${formatNumber(row.currentUnits)} ${unitLabel}</td>
        <td>${row.unitGap > 0 ? '+' + formatNumber(row.unitGap) : '—'}</td>
        <td>${formatNumber(row.requiredAssociates)}</td>
        <td>${formatNumber(row.currentAssociates)}</td>
        <td>${row.associateGap > 0 ? '+' + formatNumber(row.associateGap) : '—'}</td>
        <td class="${row.onTrack ? 'signal-ok' : (row.type === 'operator' ? 'signal-short' : 'signal-risk')}">${statusText}</td>
      </tr>
    `;
  }).join('');
}

function renderCommandKpis(result, actionPlan) {
  const hourly = result.hourlySummary || {};
  const shipGoal = Number(result.shipGoal || 0);
  const actual = Number(hourly.actualAr41SoFar || 0);
  const expected = Number(hourly.expectedByNow || 0);
  const projected = Number(result.projectedOut ?? hourly.projectedActualOut ?? 0);
  const calculatedLow = Number(
    result.forecastRangeLow ??
    result.knownWipForecast ??
    projected ??
    0
  );
  const calculatedHigh = Number(
    result.forecastRangeHigh ??
    result.capacityProjectedOut ??
    result.capacityCeiling ??
    projected ??
    0
  );
  const remainingHours = Math.max(0, Number(hourly.remainingHours || 0));
  const downtimeLost = Number(result.downtimeLostTotal || 0);

  // Display the route-specific conservative and high-workload cases.
  // Projected OUT is a working point inside this range; it must not replace the low.
  const projectedLow = Math.max(actual, Math.min(calculatedLow, calculatedHigh));
  const projectedHigh = Math.max(projectedLow, calculatedHigh);

  const goalGap = shipGoal > 0 ? projected - shipGoal : 0;
  const paceGap = actual - expected;
  const recoveryNeededJph = shipGoal > 0 && remainingHours > 0
    ? Math.ceil(Math.max(shipGoal - actual, 0) / remainingHours)
    : 0;

  const live = result.liveWipBottleneck || null;
  const downtimeArea = getTopDowntimeImpactArea(result.areas || []);
  const structural = result.bottleneck || result.realBottleneck || null;

  let focusArea = '-';
  let focusReason = 'Largest WIP / downtime pressure';

  if (live && Number(live.currentWip || 0) > 0) {
    focusArea = live.label;
    focusReason = buildLivePressureMeta(live);
  } else if (downtimeArea) {
    focusArea = downtimeArea.label;
    focusReason = `${formatNumber(downtimeArea.downtimeLostJobs || 0)} jobs lost from downtime`;
  } else if (structural) {
    focusArea = structural.label;
    focusReason = buildStructuralMeta(structural);
  }

  setText('kpiCurrentOut', formatNumber(actual));
  setText('kpiProjectedOut', formatNumber(projected));
  setText('kpiProjectedRange', `${formatNumber(projectedLow)} - ${formatNumber(projectedHigh)}`);
  setText('kpiShipGoal', shipGoal > 0 ? formatNumber(shipGoal) : '0');

  setText(
    'kpiGoalGap',
    shipGoal > 0
      ? (goalGap >= 0 ? `+${formatNumber(goalGap)} ahead of goal` : `${formatNumber(goalGap)} behind goal`)
      : 'No target entered'
  );

  setText('kpiExpectedNow', formatNumber(expected));
  setText(
    'kpiPaceGap',
    expected > 0
      ? (paceGap >= 0 ? `+${formatNumber(paceGap)} ahead pace` : `${formatNumber(paceGap)} behind pace`)
      : 'Waiting on hourly data'
  );

  setText('kpiRecoveryJph', `${formatNumber(recoveryNeededJph)} JPH`);
  setText('kpiDowntimeLoss', `${formatNumber(downtimeLost)} jobs`);
  setText('kpiFocusArea', focusArea);
  setText('kpiFocusReason', focusReason);

  setKpiCardStatus('kpiCardProjectedOut', shipGoal <= 0 ? 'neutral' : (projected >= shipGoal ? 'good' : (projected >= shipGoal * 0.90 ? 'warning' : 'bad')));
  setKpiCardStatus('kpiCardShipGoal', shipGoal <= 0 ? 'neutral' : (projected >= shipGoal ? 'good' : 'bad'));
  setKpiCardStatus('kpiCardExpectedNow', expected <= 0 ? 'neutral' : (actual >= expected ? 'good' : (actual >= expected * 0.90 ? 'warning' : 'bad')));
  setKpiCardStatus('kpiCardRecovery', recoveryNeededJph <= 0 ? 'good' : (recoveryNeededJph <= Number(hourly.targetPace || 0) ? 'warning' : 'bad'));
  setKpiCardStatus('kpiCardDowntime', downtimeLost > 0 ? 'bad' : 'good');
}

function setKpiCardStatus(cardId, status) {
  const card = document.getElementById(cardId);
  if (!card) return;
  card.classList.remove('good', 'warning', 'bad');
  if (status === 'good' || status === 'warning' || status === 'bad') {
    card.classList.add(status);
  }
}



function buildLivePressureMeta(live) {
  if (!live) return 'No live WIP pressure';
  const wip = Number(live.currentWip || live.sourceWip || 0);
  const source = live.sourceLabel || (live.sourceArea && live.sourceArea.label) || '';
  const pressure = live.pressureLabel || live.label || '';
  return `${formatNumber(wip)} WIP from ${source} → ${pressure} · ${formatPressureClearText(live)}`;
}

function buildSummaryFootnote(result, actionPlan) {
  const hourly = result.hourlySummary || {};
  const actual = Number(hourly.actualAr41SoFar || 0);
  const expected = Number(hourly.expectedByNow || 0);
  const parts = [
    `Actual AR41 is ${formatNumber(actual)} vs expected ${formatNumber(expected)}.`,
    'Forecast range uses BOS WIP, machines running, associates assigned, live WIP pressure, and downtime impacts.'
  ];
  if (actionPlan.primaryFocusArea && actionPlan.primaryFocusArea !== '-') {
    parts.push(`Primary focus right now is ${actionPlan.primaryFocusArea}.`);
  }
  return parts.join(' ');
}

function buildStructuralMeta(area) {
  if (!area) return 'No structural bottleneck yet';
  let msg = `at ${round1(area.effectiveJph || 0)} usable JPH`;
  if (Number(area.routeLoadFactor || 1) < 1) msg += ' · route-adjusted';
  if (area.supportCellLabel) msg += ` · ${area.supportCellLabel}`;
  return msg;
}

function getTopDowntimeImpactArea(areas) {
  const filtered = (areas || []).filter(function (area) {
    return Number(area.downtimeLostJobs || 0) > 0;
  });
  if (!filtered.length) return null;
  return filtered.sort(function (a, b) {
    return Number(b.downtimeLostJobs || 0) - Number(a.downtimeLostJobs || 0);
  })[0];
}

function buildGoalActionPlan(result) {
  const hourly = result.hourlySummary || {};
  const shipGoal = Number(result.shipGoal || 0);
  const projected = Number(hourly.projectedActualOut || 0);
  const remainingHours = Math.max(0, Number(hourly.remainingHours || 0));
  const live = result.liveWipBottleneck || null;
  const downtimeArea = getTopDowntimeImpactArea(result.areas || []);
  const structural = result.bottleneck || result.realBottleneck || null;
  const projectedGap = Math.max(0, shipGoal - projected);
  const extraJphNeeded = shipGoal > 0 && remainingHours > 0 ? projectedGap / remainingHours : 0;

  let primaryFocusArea = '-';
  if (live && Number(live.timeToClear || 0) >= 1) primaryFocusArea = live.label;
  else if (downtimeArea) primaryFocusArea = downtimeArea.label;
  else if (structural) primaryFocusArea = structural.label;

  let recommendedAction = 'Maintain current pace and hold the bottleneck stable.';
  if (shipGoal <= 0) {
    recommendedAction = primaryFocusArea !== '-' ? `Focus on ${primaryFocusArea} and keep downtime low.` : 'Enter a ship goal to unlock action guidance.';
  } else if (projectedGap <= 0) {
    recommendedAction = primaryFocusArea !== '-' ? `Goal is covered. Protect flow at ${primaryFocusArea} and prevent new downtime.` : 'Goal is covered. Maintain flow and protect machine uptime.';
  } else {
    const focusParts = [];
    if (live) focusParts.push(`clear live WIP in ${live.label}`);
    if (downtimeArea) focusParts.push(`cut downtime in ${downtimeArea.label}`);
    if (!focusParts.length && structural) focusParts.push(`protect capacity at ${structural.label}`);
    recommendedAction = focusParts.length ? `${capitalizeFirst(focusParts.join(', '))}.` : 'Raise usable JPH and remove local constraints.';
  }

  let bottomLine;
  if (shipGoal <= 0) {
    bottomLine = primaryFocusArea !== '-' ? `No ship goal is entered. Focus on ${primaryFocusArea} and keep downtime under control.` : 'No ship goal is entered. Add a target to get action guidance.';
  } else if (projectedGap <= 0) {
    bottomLine = `We are on pace to cover the ${formatNumber(shipGoal)} job goal. Protect uptime and keep the live WIP focus area stable.`;
  } else {
    const liveText = live ? `clear live WIP in ${live.label}` : 'clear the highest live WIP area';
    const downtimeText = downtimeArea ? `reduce downtime in ${downtimeArea.label}` : 'protect machine uptime';
    bottomLine = `We need ${formatNumber(projectedGap)} more jobs to hit the ${formatNumber(shipGoal)} goal. Focus first to ${liveText} and ${downtimeText}.`;
  }

  return {
    extraJobsNeeded: projectedGap,
    extraJphNeeded,
    primaryFocusArea,
    recommendedAction,
    bottomLine
  };
}

function capitalizeFirst(text) {
  const value = String(text || '');
  return value ? value.charAt(0).toUpperCase() + value.slice(1) : value;
}

function buildGoalRiskSub(result, projected) {
  const shipGoal = Number(result.shipGoal || 0);
  if (shipGoal <= 0) return 'Enter ship goal';
  if (result.goalRisk === 'GREEN') return 'Capacity can meet goal';
  if (result.goalRisk === 'YELLOW') return 'Goal is tight';
  if (result.goalRisk === 'RED') return 'Goal exceeds capacity';
  return projected >= shipGoal ? 'Projection can meet goal' : 'Projection below goal';
}

function buildProjectedGapText(shipGoal, projected) {
  if (Number(shipGoal || 0) <= 0) return 'No target entered';
  const gap = Number(projected || 0) - Number(shipGoal || 0);
  if (gap >= 0) return `+${formatNumber(gap)} ahead of goal`;
  return `${formatNumber(gap)} behind goal`;
}

function getShiftModeDisplay(value) {
  if (value === 'Weekday') return 'Weekday - 9.5 Hours';
  if (value === 'Weekday OT 12') return 'Weekday OT 12 - 10.5 Hours';
  if (value === 'Weekend') return 'Weekend - 10.5 Hours';
  return value || 'Weekday - 9.5 Hours';
}

function setPaceVisuals(paceRatio, actual, expected) {
  const safeRatio = Math.max(0, Number(paceRatio || 0));
  const gauge = document.querySelector('.gauge-ring');
  const statusDot = document.querySelector('.status-dot');
  const statusColor = safeRatio >= 1 ? 'var(--green)' : safeRatio >= 0.90 ? 'var(--yellow)' : 'var(--red)';
  const pct = Math.min(100, Math.round(safeRatio * 100));

  if (gauge) {
    gauge.style.background = `radial-gradient(circle at center, #06111f 0 58%, transparent 59%), conic-gradient(${statusColor} 0 ${pct}%, rgba(255,255,255,0.12) ${pct}% 100%)`;
  }

  if (statusDot) {
    statusDot.style.background = statusColor;
    statusDot.style.boxShadow = `0 0 22px ${statusColor}`;
  }

  const maxValue = Math.max(Number(actual || 0), Number(expected || 0), 1);
  setBarWidth('paceBarActual', (Number(actual || 0) / maxValue) * 100);
  setBarWidth('paceBarExpected', (Number(expected || 0) / maxValue) * 100);
}

function setBarWidth(id, percent) {
  const el = document.getElementById(id);
  if (!el) return;
  el.style.width = `${Math.max(0, Math.min(100, Number(percent || 0)))}%`;
}

function renderHourlyChart(hourlyRows, hourlySummary, result) {
  const svg = document.getElementById('hourlyChartSvg');
  if (!svg) return;

  const rows = Array.isArray(hourlyRows) ? hourlyRows : [];
  const shiftMode = result && result.shiftMode ? result.shiftMode : getValue('shiftMode', 'Weekday');
  const windowInfo = getShiftWindow(shiftMode);
  const startMinutes = parseTimeToMinutes(windowInfo.start) || 420;
  const endMinutes = parseTimeToMinutes(windowInfo.end) || 1050;
  const slots = buildPerformanceHourSlots(startMinutes, endMinutes);
  const totalPoints = Math.max(2, slots.length);
  const rowOffset = getHourlyRowOffsetForShift(shiftMode);
  const shiftHours = Number(result.shiftHours || 1);

  const displayRows = slots.map(function (slot, idx) {
    return rows[rowOffset + idx] || { hour: minutesToHourKey(slot.start), inspection: 0, active: 'Yes' };
  });

  const width = 720;
  const height = 230;
  const padLeft = 56;
  const padRight = 20;
  const padTop = 20;
  const padBottom = 36;
  const plotW = width - padLeft - padRight;
  const plotH = height - padTop - padBottom;
  const targetPace = Number(hourlySummary.targetPace || 0);
  const actualPace = Number(hourlySummary.actualPace || 0);
  const capacityCeiling = Number(result.capacityCeiling || 0);
  const hoursEntered = Number(hourlySummary.hoursEntered || 0);

  let cumulative = 0;
  const actualPoints = [];
  displayRows.forEach(function (row, index) {
    if (!row || row.active === 'No') {
      actualPoints.push({ xHour: index + 1, yValue: cumulative });
      return;
    }
    cumulative += Number(row.inspection || 0);
    actualPoints.push({ xHour: index + 1, yValue: cumulative });
  });

  const expectedPoints = [];
  const projectedPoints = [];
  for (let i = 1; i <= totalPoints; i += 1) {
    const slotEndHours = Math.max(0, (slots[i - 1].end - startMinutes) / 60);
    expectedPoints.push({ xHour: i, yValue: targetPace * slotEndHours });
    const actualSoFar = Number(hourlySummary.actualAr41SoFar || 0);
    const projectedValue = i <= hoursEntered
      ? (actualPoints[i - 1] ? actualPoints[i - 1].yValue : 0)
      : actualSoFar + (actualPace * Math.max(0, slotEndHours - hoursEntered));
    projectedPoints.push({ xHour: i, yValue: Math.min(capacityCeiling || projectedValue, projectedValue) });
  }

  const rawMax = Math.max(
    capacityCeiling,
    Number(hourlySummary.projectedActualOut || 0),
    Number(hourlySummary.expectedByNow || 0),
    Number(hourlySummary.actualAr41SoFar || 0),
    100
  );
  const maxY = getNiceChartMax(rawMax * 1.1);

  function x(hour) { return padLeft + ((hour - 1) / Math.max(1, totalPoints - 1)) * plotW; }
  function y(value) { return padTop + plotH - (Number(value || 0) / maxY) * plotH; }
  function path(points) {
    if (!points.length) return '';
    return points.map(function (p, idx) {
      return `${idx === 0 ? 'M' : 'L'} ${x(p.xHour).toFixed(1)} ${y(p.yValue).toFixed(1)}`;
    }).join(' ');
  }
  function areaPath(points) {
    if (!points.length) return '';
    const baseline = height - padBottom;
    const line = points.map(function (p, idx) {
      return `${idx === 0 ? 'M' : 'L'} ${x(p.xHour).toFixed(1)} ${y(p.yValue).toFixed(1)}`;
    }).join(' ');
    const lastX = x(points[points.length - 1].xHour).toFixed(1);
    const firstX = x(points[0].xHour).toFixed(1);
    return `${line} L ${lastX} ${baseline} L ${firstX} ${baseline} Z`;
  }

  let html = '';
  html += `<defs><linearGradient id="chartActualGrad" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0%" class="chart-grad-stop-top"></stop>
    <stop offset="100%" class="chart-grad-stop-bottom"></stop>
  </linearGradient></defs>`;
  for (let i = 0; i <= 4; i += 1) {
    const gridValue = (maxY / 4) * (4 - i);
    const yy = padTop + (plotH / 4) * i;
    html += `<line class="chart-grid" x1="${padLeft}" y1="${yy}" x2="${width - padRight}" y2="${yy}"></line>`;
    html += `<text class="chart-y-label" x="${padLeft - 8}" y="${yy + 4}" text-anchor="end">${formatCompactNumber(gridValue)}</text>`;
  }

  for (let i = 1; i <= totalPoints; i += 1) {
    const xx = x(i);
    if (i < totalPoints) {
      html += `<line class="chart-grid" x1="${xx}" y1="${padTop}" x2="${xx}" y2="${height - padBottom}"></line>`;
    }
    // Every point on this chart is a cumulative value computed through
    // slot.end (see expectedCum/slotEndHours above) — so the tick under it
    // must always say slot.end too. The table/setup grid intentionally keep
    // slots[].label start-based for readability; reusing that here caused a
    // partial edge slot's forced end-label to collide with the very next
    // slot's start-label at the same clock time (duplicate "7a" ticks) while
    // the true shift-start time never showed up anywhere.
    const label = slots[i - 1] ? formatHourLabel(formatMinutesAsClock(slots[i - 1].end)) : '';
    const anchor = i === 1 ? 'start' : (i === totalPoints ? 'end' : 'middle');
    html += `<text class="chart-label" x="${xx}" y="${height - 12}" text-anchor="${anchor}">${label}</text>`;
  }

  html += `<line class="chart-axis" x1="${padLeft}" y1="${height - padBottom}" x2="${width - padRight}" y2="${height - padBottom}"></line>`;
  html += `<line class="chart-axis" x1="${padLeft}" y1="${padTop}" x2="${padLeft}" y2="${height - padBottom}"></line>`;
  const realActualPoints = hoursEntered > 0 ? actualPoints.slice(0, Math.min(actualPoints.length, hoursEntered)) : [];
  if (realActualPoints.length > 1) {
    html += `<path class="chart-area-actual" d="${areaPath(realActualPoints)}"></path>`;
  }
  html += `<path class="chart-line-expected" d="${path(expectedPoints)}"></path>`;
  html += `<path class="chart-line-projected" d="${path(projectedPoints)}"></path>`;
  html += `<path class="chart-line-actual" d="${path(actualPoints)}"></path>`;

  if (hoursEntered > 0) {
    const dotPoint = actualPoints[Math.min(actualPoints.length, hoursEntered) - 1];
    if (dotPoint) {
      html += `<circle class="chart-actual-dot" cx="${x(dotPoint.xHour).toFixed(1)}" cy="${y(dotPoint.yValue).toFixed(1)}" r="4"></circle>`;
    }
  }

  if (hoursEntered > 0) {
    const nowIndex = Math.min(totalPoints, Math.max(1, hoursEntered + 1));
    const nowX = x(nowIndex);
    html += `<line class="chart-now" x1="${nowX}" y1="${padTop}" x2="${nowX}" y2="${height - padBottom}"></line>`;
    html += `<rect class="chart-now-tag" x="${nowX - 16}" y="4" width="32" height="14" rx="3"></rect>`;
    html += `<text class="chart-now-label" x="${nowX}" y="14" text-anchor="middle">NOW</text>`;
  }

  svg.innerHTML = html;
}

function getNiceChartMax(value) {
  const num = Math.max(100, Number(value || 0));
  if (num <= 250) return 250;
  if (num <= 500) return 500;
  if (num <= 1000) return 1000;
  if (num <= 1500) return 1500;
  if (num <= 2000) return 2000;
  return Math.ceil(num / 500) * 500;
}

function formatCompactNumber(value) {
  const num = Number(value || 0);
  if (num >= 1000) return `${(num / 1000).toFixed(num % 1000 === 0 ? 0 : 1)}k`;
  return String(Math.round(num));
}

function formatHourLabel(value) {
  const text = String(value || '');
  // Range labels (final partial slot, e.g. "5:00 PM–5:30 PM") — axis space is tight,
  // so show only the end time, which is what actually matters (when the shift closes).
  const parts = text.split(/[–-]/);
  const display = parts.length > 1 ? parts[parts.length - 1].trim() : text.trim();

  const match = display.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (!match) return display;

  const hour = match[1];
  const minute = match[2];
  const suffix = match[3].toUpperCase() === 'AM' ? 'a' : 'p';
  return minute === '00' ? `${hour}${suffix}` : `${hour}:${minute}${suffix}`;
}



function formatPressureClearText(row) {
  if (!row) return '0 min';

  const sourceKey = String(row.sourceKey || '');
  const sourceWip = Number(row.sourceWip || row.currentWip || 0);
  const pressureJph = Number(row.pressureJph || row.effectiveJph || 0);
  const pressureArea = row.pressureArea || {};
  const pressureAreaWip = Number(pressureArea.currentWip || 0);
  const timeMode = row.timeMode || '';

  if (sourceWip <= 0) return '0 min';

  // SF Unbox scan means jobs are staging/moving into Detaper / Auto Blocker.
  // Operators understand this better as operational minutes, not 0.3 hr.
  if (timeMode === 'sf_unbox_stage' || sourceKey === 'unbox') {
    if (sourceWip <= 150) return '10–20 min';
    if (sourceWip <= 250) return '20–30 min';
    return '30+ min';
  }

  // Auto Blocker scan means jobs are moving into Cooling.
  // It is usually a short transfer unless Cooling itself is overloaded.
  if (timeMode === 'cooling_transfer' || sourceKey === 'autoblocker') {
    const limit = Number(row.bufferSignalLimit || 200);
    if (pressureAreaWip > limit) return `Cooling over ${formatNumber(limit)} WIP`;
    return '5–10 min';
  }

  // Cooling scan means dwell time before ORB / Generator.
  if (timeMode === 'cooling_dwell' || sourceKey === 'cooling') {
    return '25–50 min dwell';
  }

  if (pressureJph <= 0) return 'BLOCKED';

  const mins = Math.max(0, Number(row.pressureHours || row.timeToClear || 0) * 60);
  return formatMinutesRange(mins);
}

function formatMinutesRange(minutes) {
  const mins = Number(minutes || 0);
  if (mins <= 0) return '0 min';
  if (mins < 5) return '<5 min';
  if (mins < 10) return '5–10 min';
  if (mins < 15) return '10–15 min';
  if (mins < 20) return '15–20 min';
  if (mins < 30) return '20–30 min';
  if (mins < 45) return '30–45 min';
  if (mins < 60) return '45–60 min';

  const rounded = Math.round(mins / 5) * 5;
  return `${formatNumber(rounded)} min`;
}

function renderBottleneckTable(areas) {
  const body = document.getElementById('bottleneckTable');
  if (!body) return;

  const pressureRows = buildLiveWipPressureRows(areas || []);

  const rows = pressureRows
    .slice()
    .sort(function (a, b) {
      const flowDiff = getSurfaceFlowOrderIndex(a.sourceKey) - getSurfaceFlowOrderIndex(b.sourceKey);
      if (flowDiff !== 0) return flowDiff;
      return getSurfaceFlowOrderIndex(a.pressureAreaKey) - getSurfaceFlowOrderIndex(b.pressureAreaKey);
    })
    .map(function (row) {
      const source = row.sourceArea || {};
      const capacity = row.capacityArea || row.pressureArea || source || {};
      const sourceIsBuffer = source.type === 'buffer';
      const sourceWip = Number(row.sourceWip || 0);
      const pressureJph = Number(row.pressureJph || 0);

      /*
        Important:
        The table row must match the setup card for the area on the left.
        Example: Detaping / ODT must show Detaping associates and Detaping machines,
        not Coater associates/machines just because Detaping scan feeds Coater.

        Route pressure is still used for the TIME TO CLEAR and SIGNAL because live WIP
        means "last scan -> next process pressure". The setup columns stay tied to
        the row area so the table is not confusing.
      */
      const associatesDisplay = round1(source.associatesAssigned || 0);
      const unitsDisplay = sourceIsBuffer
        ? `${round1(source.unitsRunning || 0)} tower(s)`
        : round1(source.unitsRunning || 0);
      const jphDisplay = sourceIsBuffer
        ? 'Cooling 25–50 min'
        : round1(source.effectiveJph || source.totalJph || 0);
      const capacityDisplay = sourceIsBuffer
        ? `Buffer ${formatNumber(source.bufferCapacity || 180)} jobs`
        : formatNumber(source.adjustedCapacity || 0);
      const downtimeDisplay = formatNumber(source.downtimeLostJobs || 0);

      const clearText = formatPressureClearText(row);

      return `
        <tr>
          <td>
            <strong>${escapeHtml(source.label || '-')}</strong>
            <div class="muted-cell">→ ${escapeHtml(row.pressureLabel || '-')}</div>
          </td>
          <td>${associatesDisplay}</td>
          <td>${unitsDisplay}</td>
          <td>${jphDisplay}</td>
          <td>${capacityDisplay}</td>
          <td>${downtimeDisplay}</td>
          <td>${formatNumber(sourceWip)}</td>
          <td>${clearText}</td>
          <td class="${row.signalClass}">${escapeHtml(row.signal)}</td>
        </tr>
      `;
    });

  body.innerHTML = rows.join('') || '<tr><td colspan="9">No live WIP pressure yet.</td></tr>';
}

async function saveForecast() {
  await postToApi('saveForecast', 'Forecast saved.');
}

async function saveDowntime() {
  await postToApi('saveDowntime', 'Downtime saved.');
}

async function saveHourly() {
  await postToApi('saveHourly', 'Hourly input saved.');
}

async function postToApi(action, successMessage) {
  if (!SURFACE_API_URL) {
    alert('API URL is missing.');
    return;
  }

  if (!latestForecastResult) calculateForecast();

  try {
    const response = await fetch(SURFACE_API_URL, {
      method: 'POST',
      body: JSON.stringify({
        action,
        payload: latestForecastResult.payload
      })
    });

    const data = await response.json();

    if (!data.ok) {
      throw new Error(data.error || 'Save failed.');
    }

    savePageState();
    alert(successMessage);

  } catch (err) {
    alert(`Save failed: ${err.message || err}`);
  }
}

function buildRecommendation(data) {
  if (!data.bottleneck) {
    return 'Enter running associates/machines to calculate the forecast.';
  }

  let msg =
    `Forecast range is ${formatNumber(data.projectedOut || data.knownWipForecast)} to ${formatNumber(data.capacityProjectedOut || data.knownWipForecast)}. ` +
    `Low = current pace projection. High = BOS WIP + average incoming (${formatNumber(data.incomingLow)}-${formatNumber(data.incomingHigh)}) + average remake/breakage (${formatNumber(data.remakeLow)}-${formatNumber(data.remakeHigh)}), capped by bottleneck capacity. ` +
    `Current bottleneck is ${data.bottleneck.label} at ${round1(data.effectiveBottleneckJph)} effective JPH. ` +
    `Total associates entered: ${round1(data.totalAssociates)}.`;

  if (data.liveWipBottleneck) {
    msg += ` Live WIP bottleneck is ${data.liveWipBottleneck.label}: ${formatNumber(data.liveWipBottleneck.currentWip)} WIP and ${round1(data.liveWipBottleneck.timeToClear)} hr to clear.`;
  }

  if (data.hourlySummary && data.hourlySummary.hoursEntered > 0) {
    msg += ` Hourly status is ${data.hourlySummary.onTrackStatus}. ` +
      `AR41 actual so far: ${formatNumber(data.hourlySummary.actualAr41SoFar)}. ` +
      `Expected by now: ${formatNumber(data.hourlySummary.expectedByNow)}. ` +
      `Projected actual OUT: ${formatNumber(data.hourlySummary.projectedActualOut)}.`;
  } else {
    msg += ' No hourly AR41 output entered yet.';
  }

  if (data.shipGoal > 0) {
    if (data.goalRisk === 'RED') msg += ` Ship goal is not realistic. It needs ${round1(data.requiredJph)} JPH.`;
    else if (data.goalRisk === 'YELLOW') msg += ' Ship goal is tight. Any downtime can break it.';
    else msg += ' Ship goal is feasible with current capacity.';
  } else {
    msg += ' No ship goal entered. Forecast is capacity-based only.';
  }

  if (data.downtimeLostTotal > 0) {
    msg += ` Downtime reduces capacity by ${formatNumber(data.downtimeLostTotal)} jobs.`;
  }

  return msg;
}



async function loadCurrentStateFromApi() {
  if (!SURFACE_API_URL) return false;

  try {
    const response = await fetch(`${SURFACE_API_URL}?action=getCurrentState&ts=${Date.now()}`);
    const data = await response.json();

    if (!data || !data.ok || !data.hasState || !data.payload) {
      return false;
    }

    isRestoringState = true;

    try {
      restorePayloadToPage(data.payload);
      latestForecastResult = {
        payload: data.payload,
        result: data.result || null
      };

      return true;
    } finally {
      isRestoringState = false;
    }

  } catch (err) {
    console.warn('Could not load current state from API. Local backup will be used.', err);
    return false;
  }
}

async function saveCurrentStateToApi(showAlert) {
  if (!SURFACE_API_URL) return;

  const payload = collectPayload();

  try {
    const response = await fetch(SURFACE_API_URL, {
      method: 'POST',
      body: JSON.stringify({
        action: 'saveCurrentState',
        payload: payload
      })
    });

    const data = await response.json();

    if (!data.ok) {
      throw new Error(data.error || 'Current state save failed.');
    }

    if (showAlert) {
      alert('Current page data saved. It can now be opened from another computer.');
    }

  } catch (err) {
    console.warn('Cloud autosave failed:', err);
    if (showAlert) {
      alert(`Cloud save failed: ${err.message || err}`);
    }
  }
}

function restorePayloadToPage(payload) {
  if (!payload) return false;

  setValue('productionDate', payload.productionDate || getLocalDateInputValue(new Date()));
    setValue('perfDateInput', payload.productionDate || getLocalDateInputValue(new Date()));
    setValue('shiftMode', payload.shiftMode || getAutomaticShiftMode(payload.productionDate));
  setValue('bosWip', payload.bosWip || 0);
  setValue('shipGoal', payload.shipGoal || 0);
  setValue('incomingLow', payload.incomingLow ?? SURFACE_EXTRA_WORK_DEFAULTS.incomingLow);
  setValue('incomingHigh', payload.incomingHigh ?? SURFACE_EXTRA_WORK_DEFAULTS.incomingHigh);
  setValue('remakeLow', payload.remakeLow ?? SURFACE_EXTRA_WORK_DEFAULTS.remakeLow);
  setValue('remakeHigh', payload.remakeHigh ?? SURFACE_EXTRA_WORK_DEFAULTS.remakeHigh);
  setValue('floaterCount', payload.floaterCount || 0);
  setValue('floaterAssign', payload.floaterAssign || 'none');

  if (payload.areas) {
    Object.keys(payload.areas).forEach(function (key) {
      setValue(`area_${key}`, payload.areas[key]);
    });
  }

  if (payload.associates) {
    Object.keys(payload.associates).forEach(function (key) {
      setValue(`assoc_${key}`, payload.associates[key]);
    });
  }

  if (payload.ratios) {
    Object.keys(payload.ratios).forEach(function (key) {
      if (document.getElementById(`ratio_${key}`)) setValue(`ratio_${key}`, payload.ratios[key]);
    });
  }

  buildHourlyRows(payload.shiftMode || getAutomaticShiftMode(payload.productionDate), false);
    restoreHourlyRows(payload.hourly || []);
  restoreDowntimeRows(payload.downtime || []);

  return true;
}


function autoSavePageState() {
  if (isRestoringState) return;

  window.clearTimeout(autoSavePageState.timer);
  autoSavePageState.timer = window.setTimeout(function () {
    savePageState();
    saveCurrentStateToApi(false);
  }, 700);
}

function savePageState() {
  try {
    const payload = collectPayload();

    const state = {
      savedAt: new Date().toISOString(),
      payload: payload
    };

    localStorage.setItem(SURFACE_STORAGE_KEY, JSON.stringify(state));
  } catch (err) {
    console.warn('Surface Forecast autosave failed:', err);
  }
}

function restorePageState() {
  const raw = localStorage.getItem(SURFACE_STORAGE_KEY);
  if (!raw) return false;

  let state = null;

  try {
    state = JSON.parse(raw);
  } catch (err) {
    console.warn('Surface Forecast restore failed:', err);
    return false;
  }

  if (!state || !state.payload) return false;

  isRestoringState = true;

  try {
    const payload = state.payload;

    setValue('productionDate', payload.productionDate || getLocalDateInputValue(new Date()));
    setValue('perfDateInput', payload.productionDate || getLocalDateInputValue(new Date()));
    setValue('shiftMode', payload.shiftMode || getAutomaticShiftMode(payload.productionDate));
    setValue('bosWip', payload.bosWip || 0);
    setValue('shipGoal', payload.shipGoal || 0);
    setValue('floaterCount', payload.floaterCount || 0);
    setValue('floaterAssign', payload.floaterAssign || 'none');

    if (payload.areas) {
      Object.keys(payload.areas).forEach(function (key) {
        setValue(`area_${key}`, payload.areas[key]);
      });
    }

    if (payload.associates) {
      Object.keys(payload.associates).forEach(function (key) {
        setValue(`assoc_${key}`, payload.associates[key]);
      });
    }

    if (payload.ratios) {
      Object.keys(payload.ratios).forEach(function (key) {
        if (document.getElementById(`ratio_${key}`)) setValue(`ratio_${key}`, payload.ratios[key]);
      });
    }

    buildHourlyRows(payload.shiftMode || getAutomaticShiftMode(payload.productionDate), false);
    restoreHourlyRows(payload.hourly || []);
    restoreDowntimeRows(payload.downtime || []);

    return true;

  } finally {
    isRestoringState = false;
  }
}

function restoreHourlyRows(hourlyRows) {
  const savedByHour = {};
  (hourlyRows || []).forEach(function (saved) {
    savedByHour[normalizeHourKey(saved.hour)] = saved;
  });

  document.querySelectorAll('.hourly-row').forEach(function (row) {
    const key = row.dataset.apiHour || normalizeHourKey(getRowValue(row, '.hr-hour'));
    const saved = savedByHour[key];
    if (!saved) return;

    setRowValue(row, '.hr-unbox', saved.unbox || '');
    setRowValue(row, '.hr-autoblocker', saved.autoblocker || '');
    setRowValue(row, '.hr-cooling', saved.cooling || '');
    setRowValue(row, '.hr-orb', saved.orb || '');
    setRowValue(row, '.hr-polisher', saved.polisher || '');
    setRowValue(row, '.hr-engraving', saved.engraving || '');
    setRowValue(row, '.hr-detaping', saved.detaping || '');
    setRowValue(row, '.hr-coater', saved.coater || '');
    setRowValue(row, '.hr-inspection', saved.inspection || '');
    setRowValue(row, '.hr-notes', saved.notes || '');
    setRowValue(row, '.hr-active', saved.active || 'Yes');
  });
}

function restoreDowntimeRows(downtimeRows) {
  const holder = document.getElementById('downtimeRows');
  if (!holder) return;

  holder.innerHTML = '';
  downtimeCounter = 0;

  if (!downtimeRows.length) {
    addDowntimeRow();
    return;
  }

  downtimeRows.forEach(function (saved) {
    const row = addDowntimeRow();

    if (!row) return;

    setRowValue(row, '.dt-area', saved.areaKey || 'unbox');
    setRowValue(row, '.dt-issue', saved.issue || 'Machine Down');
    setRowValue(row, '.dt-units', saved.unitsDown || 0);
    setRowValue(row, '.dt-start', saved.startTime || '08:45');
    setRowValue(row, '.dt-end', saved.endTime || '10:15');
    setRowValue(row, '.dt-fullday', saved.fullDay || 'No');
    setRowValue(row, '.dt-notes', saved.notes || '');
    setRowValue(row, '.dt-active', saved.active || 'No');
  });
}

function clearSavedSurfaceForecastState() {
  localStorage.removeItem(SURFACE_STORAGE_KEY);
  alert('Saved browser state cleared. Refresh the page to reload default values.');
}

function setValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value;
}

function setRowValue(row, selector, value) {
  const el = row.querySelector(selector);
  if (el) el.value = value;
}




function loadSurfaceBosSnapshots() {
  try {
    const parsed = JSON.parse(localStorage.getItem(SURFACE_BOS_WIP_STORAGE_KEY) || '{}');
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (err) {
    return {};
  }
}

function saveSurfaceBosSnapshots(snapshots) {
  localStorage.setItem(
    SURFACE_BOS_WIP_STORAGE_KEY,
    JSON.stringify(snapshots || {})
  );
}

function normalizeSurfaceBosMap(source) {
  const result = {};
  SURFACE_FLOW_ORDER.forEach(function (key) {
    result[key] = Math.max(0, Number(source && source[key]) || 0);
  });
  return result;
}

function getSurfaceBosTotal(snapshot) {
  return SURFACE_FLOW_ORDER.reduce(function (sum, key) {
    return sum + (Number(snapshot && snapshot[key]) || 0);
  }, 0);
}

function getSurfaceBosSnapshot(dateKey, fallbackCurrentWip) {
  const snapshots = loadSurfaceBosSnapshots();
  const saved = snapshots[String(dateKey || '')];

  if (saved && typeof saved === 'object') {
    return normalizeSurfaceBosMap(saved);
  }

  return normalizeSurfaceBosMap(fallbackCurrentWip || {});
}

function captureSurfaceBosSnapshotIfMissing(dateKey, currentWip) {
  const key = String(dateKey || '').trim();
  if (!key) return normalizeSurfaceBosMap(currentWip || {});

  const snapshots = loadSurfaceBosSnapshots();

  if (!snapshots[key]) {
    snapshots[key] = normalizeSurfaceBosMap(currentWip || {});
    saveSurfaceBosSnapshots(snapshots);
  }

  const snapshot = normalizeSurfaceBosMap(snapshots[key]);
  setValue('bosWip', Math.round(getSurfaceBosTotal(snapshot)));
  return snapshot;
}


function readSurfaceBosInputs() {
  const snapshot = {};

  SURFACE_FLOW_ORDER.forEach(function (key) {
    const input = document.querySelector(
      `.surface-bos-value-input[data-surface-bos-key="${key}"]`
    );
    snapshot[key] = Math.max(0, Number(input && input.value) || 0);
  });

  return normalizeSurfaceBosMap(snapshot);
}

function updateSurfaceBosTotalFromInputs() {
  const snapshot = readSurfaceBosInputs();
  const total = getSurfaceBosTotal(snapshot);
  setValue('bosWip', Math.round(total));
  setText('surfaceBosStartupMeta', 'Unsaved BOS changes — click Save BOS.');
  return snapshot;
}

function saveSurfaceBosSnapshot() {
  const dateKey = getValue('productionDate', getLocalDateInputValue(new Date()));
  let snapshot = readSurfaceBosInputs();
  const enteredTotal = Math.max(0, Number(getValue('bosWip', 0)) || 0);
  const stationTotal = getSurfaceBosTotal(snapshot);

  // When the user changes only the total BOS field, preserve the station mix
  // by scaling all station values proportionally. If every station is zero,
  // place the total in SF Unbox so the forecast does not invent downstream WIP.
  if (Math.round(enteredTotal) !== Math.round(stationTotal)) {
    if (stationTotal > 0) {
      const scale = enteredTotal / stationTotal;
      let allocated = 0;

      SURFACE_FLOW_ORDER.forEach(function (key, index) {
        if (index === SURFACE_FLOW_ORDER.length - 1) {
          snapshot[key] = Math.max(0, Math.round(enteredTotal - allocated));
        } else {
          snapshot[key] = Math.max(
            0,
            Math.round((Number(snapshot[key]) || 0) * scale)
          );
          allocated += snapshot[key];
        }
      });
    } else {
      snapshot = normalizeSurfaceBosMap({});
      snapshot.unbox = Math.round(enteredTotal);
    }
  }

  const snapshots = loadSurfaceBosSnapshots();
  snapshots[dateKey] = snapshot;
  saveSurfaceBosSnapshots(snapshots);

  setValue('bosWip', Math.round(getSurfaceBosTotal(snapshot)));
  renderSurfaceBosSnapshot();
  calculateForecast();

  alert(
    `Surface BOS WIP saved at ${formatNumber(getSurfaceBosTotal(snapshot))} jobs for ${dateKey}.`
  );
}

function resetSurfaceBosSnapshot() {
  const dateKey = getValue('productionDate', getLocalDateInputValue(new Date()));
  const currentWip = normalizeSurfaceBosMap(window.surfaceDashboardCurrentWip || {});
  const snapshots = loadSurfaceBosSnapshots();

  snapshots[dateKey] = currentWip;
  saveSurfaceBosSnapshots(snapshots);
  setValue('bosWip', Math.round(getSurfaceBosTotal(currentWip)));
  renderSurfaceBosSnapshot();
  calculateForecast();

  alert(
    `Surface BOS WIP reset to ${formatNumber(getSurfaceBosTotal(currentWip))} jobs for ${dateKey}.`
  );
}

function getSurfaceClockProgress(shiftMode) {
  const windowInfo = getShiftWindow(shiftMode || 'Weekday');
  const start = parseTimeToMinutes(windowInfo.start) || 420;
  const end = parseTimeToMinutes(windowInfo.end) || 1050;
  const now = new Date();
  const nowMinutes = now.getHours() * 60 + now.getMinutes();

  return {
    start,
    end,
    totalMinutes: Math.max(1, end - start),
    elapsedMinutes: Math.max(0, Math.min(end - start, nowMinutes - start)),
    remainingMinutes: Math.max(0, end - nowMinutes)
  };
}

function getSurfaceBosMaturity(snapshot, elapsedMinutes, totalShiftMinutes) {
  const byNow = {};
  const byEos = {};
  let maturedNow = 0;
  let maturedByEos = 0;

  SURFACE_FLOW_ORDER.forEach(function (key) {
    const jobs = Math.max(0, Number(snapshot && snapshot[key]) || 0);
    const travel = Number(SURFACE_MATURITY_MINUTES[key]) || 0;

    byNow[key] = elapsedMinutes >= travel ? jobs : 0;
    byEos[key] = totalShiftMinutes >= travel ? jobs : 0;
    maturedNow += byNow[key];
    maturedByEos += byEos[key];
  });

  const downstreamBos =
    (Number(snapshot.polisher) || 0) +
    (Number(snapshot.engraving) || 0) +
    (Number(snapshot.detaping) || 0) +
    (Number(snapshot.coater) || 0) +
    (Number(snapshot.inspection) || 0);

  return {
    byNow,
    byEos,
    maturedNow,
    maturedByEos,
    downstreamBos,
    startupDelayApplied: downstreamBos <= 0,
    startupDelayRemaining: downstreamBos <= 0
      ? Math.max(0, SURFACE_STARTUP_OUTPUT_DELAY_MIN - elapsedMinutes)
      : 0
  };
}

function renderSurfaceBosSnapshot() {
  const holder = document.getElementById('surfaceBosStationGrid');
  if (!holder) return;

  const dateKey = getValue('productionDate', getLocalDateInputValue(new Date()));
  const snapshot = getSurfaceBosSnapshot(
    dateKey,
    window.surfaceDashboardCurrentWip || {}
  );
  const progress = getSurfaceClockProgress(getValue('shiftMode', 'Weekday'));
  const maturity = getSurfaceBosMaturity(
    snapshot,
    progress.elapsedMinutes,
    progress.totalMinutes
  );

  const bosTotalInput = document.getElementById('bosWip');
  if (!bosTotalInput || document.activeElement !== bosTotalInput) {
    setValue('bosWip', Math.round(getSurfaceBosTotal(snapshot)));
  }
  setText(
    'surfaceBosStartupMeta',
    maturity.startupDelayRemaining > 0
      ? `${Math.round(maturity.startupDelayRemaining)} min until normal first output window`
      : `BOS saved for ${dateKey}. Edit station values and click Save BOS.`
  );

  holder.innerHTML = SURFACE_FLOW_ORDER.map(function (key) {
    const area = SURFACE_AREAS.find(function (row) { return row.key === key; });
    const label = area ? area.label : key;
    const jobs = Number(snapshot[key]) || 0;
    const range = SURFACE_STAGE_TIME_RANGES[key] || {
      label: '0 min'
    };
    const travel = Number(SURFACE_MATURITY_MINUTES[key]) || 0;
    const ready = progress.elapsedMinutes >= travel;

    return `
      <label class="surface-bos-station ${ready ? 'ready' : 'waiting'}">
        <span>${escapeHtml(label)}</span>
        <input
          class="surface-bos-value-input"
          data-surface-bos-key="${escapeHtml(key)}"
          type="number"
          min="0"
          step="1"
          value="${Math.round(jobs)}">
        <small>Time to clear: ${escapeHtml(range.label)}</small>
      </label>
    `;
  }).join('');

  holder.querySelectorAll('.surface-bos-value-input').forEach(function (input) {
    input.addEventListener('input', updateSurfaceBosTotalFromInputs);
  });
}

async function syncSurfaceDashboardData(showAlert) {
  setText('dashboardApiState', 'Syncing...');
  setText('dashboardSyncStatus', 'Pulling Surface Dashboard API...');
  setHeroSync('syncing', 'Syncing...', null);

  try {
    const [flowData, operatorData] = await Promise.all([
      fetchDashboardProductionFlow(),
      fetchDashboardOperatorActivity()
    ]);

    const wipMap = mapDashboardWip(flowData);
    window.surfaceDashboardCurrentWip = wipMap;
    captureSurfaceBosSnapshotIfMissing(
      getValue('productionDate', getLocalDateInputValue(new Date())),
      wipMap
    );
    renderSurfaceBosSnapshot();

    const hourlyRows = mapOperatorActivityToHourly(operatorData);
    applyDashboardHourlyRows(hourlyRows);

    window.surfaceDashboardSync = {
      updatedAt: new Date().toISOString(),
      source: 'Surface Dashboard API',
      flowRows: Array.isArray(flowData.rows) ? flowData.rows.length : 0,
      operatorRows: Array.isArray(operatorData.rows) ? operatorData.rows.length : 0
    };

    setText('dashboardApiState', 'Connected');
    setText('dashboardLastSync', new Date().toLocaleTimeString());
    setText('dashboardSyncStatus', 'Synced with Surface Dashboard API');
    setHeroSync('ok', 'Connected', new Date().toLocaleTimeString());

    calculateForecast();
    savePageState();
    saveCurrentStateToApi(false);

    if (showAlert) {
      console.log('Surface Dashboard sync complete.');
    }

  } catch (err) {
    console.error('Dashboard sync failed:', err);
    setText('dashboardApiState', 'Sync failed');
    setText('dashboardSyncStatus', err.message || 'Surface Dashboard API sync failed.');
    setHeroSync('error', 'Sync failed', null);
    if (showAlert) alert(`Dashboard sync failed: ${err.message || err}`);
  }
}

function setHeroSync(state, label, time) {
  const dot = document.getElementById('heroSyncDot');
  if (dot) {
    dot.classList.remove('dot-ok', 'dot-syncing', 'dot-error');
    if (state === 'ok') dot.classList.add('dot-ok');
    if (state === 'syncing') dot.classList.add('dot-syncing');
    if (state === 'error') dot.classList.add('dot-error');
  }

  setText('heroSyncState', label);
  setText('heroSyncTime', time ? `Synced ${time}` : 'Never synced');
}

async function fetchDashboardProductionFlow() {
  const url = `${SURFACE_DASHBOARD_API_URL}?action=productionFlow&area=Surface&t=${Date.now()}`;
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Production Flow API HTTP ${res.status}`);

  const json = await res.json();
  if (!json || json.status !== 'success') throw new Error(json?.message || 'Production Flow API returned error');

  const rows = Array.isArray(json.productionFlow)
    ? json.productionFlow
    : Array.isArray(json.surfaceFlow)
      ? json.surfaceFlow
      : [];

  return { raw: json, rows };
}

async function fetchDashboardOperatorActivity() {
  const url = `${SURFACE_DASHBOARD_API_URL}?action=operatorActivity&area=Surface&t=${Date.now()}`;
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Operator Activity API HTTP ${res.status}`);

  const json = await res.json();
  if (!json || json.status !== 'success') throw new Error(json?.message || 'Operator Activity API returned error');

  const rows = Array.isArray(json.operatorActivity) ? json.operatorActivity : [];
  return { raw: json, rows };
}

function mapDashboardWip(flowData) {
  const rows = Array.isArray(flowData.rows) ? flowData.rows : [];

  const map = {
    unbox: 0,
    autoblocker: 0,
    cooling: 0,
    orb: 0,
    polisher: 0,
    engraving: 0,
    detaping: 0,
    coater: 0,
    inspection: 0
  };

  rows.forEach(function (row) {
    const step = String(row.FlowStep || row.DisplayName || row.StationKey || '').trim();
    const wip = Number(String(row.CurrentWIP ?? row.CurrentJobTotal ?? row.CurrentJobTotal ?? 0).replace(/,/g, '')) || 0;
    const key = stationToForecastKey(step);

    if (key && map[key] !== undefined) {
      map[key] += wip;
    }
  });

  return map;
}

function mapOperatorActivityToHourly(operatorData) {
  const rows = Array.isArray(operatorData.rows) ? operatorData.rows : [];
  const hourly = {};

  HOURLY_DEFAULTS.forEach(function (hour) {
    hourly[hour] = {
      hour,
      unbox: 0,
      autoblocker: 0,
      cooling: 0,
      orb: 0,
      polisher: 0,
      engraving: 0,
      detaping: 0,
      coater: 0,
      inspection: 0,
      notes: 'Dashboard sync',
      active: 'Yes'
    };
  });

  rows.forEach(function (row) {
    const stationName = getDashboardStationName(row);
    const key = stationToForecastKey(stationName);

    // IMPORTANT:
    // Only use exact machine/station rows.
    // Do not use parent rows like "Surface", "Zenni Lab", "SF Scan & Verify", or generic scan rows.
    if (!key) return;

    const hourSource = getDashboardHourSource(row);

    Object.keys(hourSource).forEach(function (label) {
      const hourKey = dashboardHourLabelToInputHour(label);
      if (!hourKey || !hourly[hourKey]) return;

      const value = Number(String(hourSource[label]).replace(/,/g, '')) || 0;
      hourly[hourKey][key] += value;
    });
  });

  return Object.values(hourly);
}

function getDashboardStationName(row) {
  return String(
    row.DisplayName ||
    row.FlowStep ||
    row.FlowStation ||
    row.flowStation ||
    row.StationKey ||
    row.AccessPoint ||
    row['Access Point'] ||
    row.Station ||
    row.station ||
    ''
  ).trim();
}

function getDashboardHourSource(row) {
  // The dashboard station cards are built from hourly station totals.
  // Some payloads store them under Hours; others are flat H07/H08 or numeric 7/8/15 columns.
  if (row.Hours && typeof row.Hours === 'object') return row.Hours;
  if (row.hours && typeof row.hours === 'object') return row.hours;

  const out = {};

  Object.keys(row || {}).forEach(function (key) {
    if (/^H?\d{1,2}$/i.test(String(key))) {
      out[key] = row[key];
    }
  });

  return out;
}

function applyDashboardHourlyRows(hourlyRows) {
  const rowByHour = {};
  hourlyRows.forEach(function (row) {
    rowByHour[row.hour] = row;
  });

  document.querySelectorAll('.hourly-row').forEach(function (tr) {
    const hour = tr.dataset.apiHour || normalizeHourKey(getRowValue(tr, '.hr-hour'));
    const row = rowByHour[hour];

    if (!row) return;

    setRowValue(tr, '.hr-unbox', row.unbox || '');
    setRowValue(tr, '.hr-autoblocker', row.autoblocker || '');
    setRowValue(tr, '.hr-cooling', row.cooling || '');
    setRowValue(tr, '.hr-orb', row.orb || '');
    setRowValue(tr, '.hr-polisher', row.polisher || '');
    setRowValue(tr, '.hr-engraving', row.engraving || '');
    setRowValue(tr, '.hr-detaping', row.detaping || '');
    setRowValue(tr, '.hr-coater', row.coater || '');
    setRowValue(tr, '.hr-inspection', row.inspection || '');

    setRowValue(tr, '.hr-notes', 'Dashboard sync');
    setRowValue(tr, '.hr-active', 'Yes');
  });
}

function stationToForecastKey(station) {
  const raw = String(station || '').trim();
  const s = raw.toLowerCase();

  // Ignore parent/group rows. These are totals and will double-count the machine station rows.
  if (!s || s === 'surface' || s === 'zenni lab' || s === 'finish' || s === 'inventory' || s === 'breakage') return '';
  if (s.includes('scan & verify') || s === 'sf scan & verify') return '';

  // Exact dashboard station-card mapping.
  if (s === 'surface unbox' || s === 'sf unbox') return 'unbox';

  if (
    s === 'blocking line b' ||
    s === 'auto blockers' ||
    s === 'auto blocker' ||
    s.includes('auto blockers + manual blocker')
  ) return 'autoblocker';

  if (
    s === 'cooling storage' ||
    s === 'iq star' ||
    s === 'iq-star' ||
    s.includes('cooling storage')
  ) return 'cooling';

  if (
    s === 'generating line b' ||
    s === 'orbit generator' ||
    s === 'orb generator' ||
    s === 'orb'
  ) return 'orb';

  if (
    s === 'polishing line b' ||
    s === 'polisher' ||
    s === 'polisher / flex'
  ) return 'polisher';

  if (
    s === 'engraving line b' ||
    s === 'engraver' ||
    s === 'engraving / otl'
  ) return 'engraving';

  if (
    s === 'detaping line b' ||
    s === 'detaper' ||
    s === 'detaping / odt'
  ) return 'detaping';

  if (
    s === 'coating line b' ||
    s === '54r coater' ||
    s === 'coater' ||
    s === 'coater / 54r'
  ) return 'coater';

  if (
    s === 'surface inspection' ||
    s === 'surface inspection out' ||
    s === 'surface inspection / ar41' ||
    s === 'ar41'
  ) return 'inspection';

  return '';
}

function dashboardHourLabelToInputHour(label) {
  const text = String(label || '').trim();
  if (!text) return '';

  // Flat report columns: 0, 1, 2, ... 20
  if (/^\d{1,2}$/.test(text)) {
    const h = Number(text);
    if (h >= 0 && h <= 23) return `${String(h).padStart(2, '0')}:00`;
  }

  // API columns: H06, H07, H15
  const hMatch = text.match(/^H(\d{1,2})$/i);
  if (hMatch) {
    const h = Number(hMatch[1]);
    if (h >= 0 && h <= 23) return `${String(h).padStart(2, '0')}:00`;
  }

  // Station detail labels: 7 AM, 10 AM, 2 PM
  const shortMatch = text.match(/^(\d{1,2})\s*(AM|PM)$/i);
  if (shortMatch) {
    let h = Number(shortMatch[1]);
    const ampm = shortMatch[2].toUpperCase();
    if (ampm === 'PM' && h < 12) h += 12;
    if (ampm === 'AM' && h === 12) h = 0;
    return `${String(h).padStart(2, '0')}:00`;
  }

  const direct = {
    '6:00 AM': '06:00',
    '7:00 AM': '07:00',
    '8:00 AM': '08:00',
    '9:00 AM': '09:00',
    '10:00 AM': '10:00',
    '11:00 AM': '11:00',
    '12:00 PM': '12:00',
    '1:00 PM': '13:00',
    '2:00 PM': '14:00',
    '3:00 PM': '15:00',
    '4:00 PM': '16:00',
    '5:00 PM': '17:00',
    '6:00 PM': '18:00',
    '7:00 PM': '19:00',
    '8:00 PM': '20:00'
  };

  return direct[text] || '';
}


function getLiveWipBottleneck(areas) {
  const rows = buildLiveWipPressureRows(areas || [])
    .filter(function (row) {
      return Number(row.sourceWip || 0) > 0;
    })
    .sort(function (a, b) {
      if (a.signalRank !== b.signalRank) return b.signalRank - a.signalRank;
      return Number(b.pressureHours || 0) - Number(a.pressureHours || 0);
    });

  return rows.length ? rows[0] : null;
}

function buildLiveWipPressureRows(areas) {
  const areaByKey = {};
  (areas || []).forEach(function (area) {
    areaByKey[area.key] = area;
  });

  return (areas || []).map(function (sourceArea) {
    const route = SURFACE_PRESSURE_ROUTE[sourceArea.key] || {
      pressureAreaKey: sourceArea.key,
      pressureLabel: sourceArea.label,
      note: 'No pressure route configured'
    };

    const pressureArea = areaByKey[route.pressureAreaKey] || sourceArea;
    const capacityArea = areaByKey[route.capacityAreaKey] || pressureArea || sourceArea;
    const sourceWip = Number(sourceArea.currentWip || 0);
    const pressureJph = getPressureAreaJph(capacityArea);
    const isBuffer = capacityArea.type === 'buffer';
    const pressureHours = !isBuffer && pressureJph > 0 ? sourceWip / pressureJph : 0;

    let signal = 'OK';
    let signalClass = 'wip-ok';
    let signalRank = 0;

    if (isBuffer && sourceWip > Number(pressureArea.bufferCapacity || 180)) {
      signal = 'BUFFER OVER';
      signalClass = 'wip-hot';
      signalRank = 3;
    } else if (!isBuffer && sourceWip > 0 && pressureJph <= 0) {
      signal = 'BLOCKED';
      signalClass = 'wip-hot';
      signalRank = 4;
    } else if (!isBuffer && pressureHours >= 2) {
      signal = 'BOTTLENECK';
      signalClass = 'wip-hot';
      signalRank = 3;
    } else if (!isBuffer && pressureHours >= 1) {
      signal = 'WATCH';
      signalClass = 'wip-watch';
      signalRank = 2;
    }

    return {
      key: `${sourceArea.key}_to_${pressureArea.key}`,
      sourceKey: sourceArea.key,
      pressureAreaKey: pressureArea.key,
      label: route.pressureLabel || pressureArea.label,
      sourceLabel: sourceArea.label,
      pressureLabel: route.pressureLabel || pressureArea.label,
      note: route.note,
      sourceArea,
      pressureArea,
      capacityArea,
      capacityAreaKey: capacityArea.key,
      capacityLabel: route.capacityLabel || capacityArea.label,
      currentWip: sourceWip,
      sourceWip,
      pressureJph,
      effectiveJph: pressureJph,
      timeToClear: pressureHours,
      pressureHours,
      signal,
      signalClass,
      signalRank
    };
  });
}

function getPressureAreaJph(area) {
  if (!area || area.type === 'buffer') return 0;
  return Math.max(0, Number(area.effectiveJph || area.totalJph || 0));
}


function normalizeForecastRange(lowValue, highValue) {
  let low = Math.max(0, Number(lowValue) || 0);
  let high = Math.max(0, Number(highValue) || 0);

  if (high < low) {
    const temp = low;
    low = high;
    high = temp;
  }

  return { low, high };
}

function getAreaCapacityInfo(area, unitsRunning, associatesAssigned, shiftHours) {
  const mode = area.capacityMode || area.type;
  const jph = Number(area.jphPerUnit || 0);
  const machines = Math.max(0, Number(unitsRunning || 0));
  const associates = Math.max(0, Number(associatesAssigned || 0));

  if (mode === 'buffer') {
    const bufferCapacity = Number(area.bufferCapacity || 0);
    return {
      activeCapacityUnits: machines,
      capacityBasis: 'buffer',
      totalJph: 0,
      normalCapacity: bufferCapacity
    };
  }

  if (mode === 'operator') {
    // Some screens have both "Associates" and "Associate Count" for the same cell.
    // Use the higher value so the same person is not double-counted.
    const activeAssociates = Math.max(associates, machines);
    return {
      activeCapacityUnits: activeAssociates,
      capacityBasis: 'associates',
      totalJph: activeAssociates * jph,
      normalCapacity: activeAssociates * jph * shiftHours
    };
  }

  // Machine output is capped by the installed/running machines.
  // Associates provide coverage and uptime; extra associates above required coverage
  // do not create fake machine JPH.
  if (mode === 'machine') {
    const machineCeilingJph = machines * jph;
    const machinesPerAssociate = Math.max(0.25, Number(area.machinesPerAssociate || 1));
    const requiredAssociates = machines > 0 ? machines / machinesPerAssociate : 0;
    const coverageFactor = requiredAssociates > 0
      ? Math.min(1, associates / requiredAssociates)
      : 0;
    const effectiveMachineJph = machineCeilingJph * coverageFactor;

    return {
      activeCapacityUnits: machines,
      capacityBasis: coverageFactor >= 1
        ? 'machine ceiling fully covered'
        : `${round1(coverageFactor * 100)}% labor coverage`,
      machineCeilingJph,
      requiredAssociates,
      coverageFactor,
      totalJph: effectiveMachineJph,
      normalCapacity: effectiveMachineJph * shiftHours
    };
  }

  return {
    activeCapacityUnits: 0,
    capacityBasis: 'unknown',
    totalJph: 0,
    normalCapacity: 0
  };
}

function getDowntimeLossForArea(area, downtime, shiftHours) {
  let lost = 0;

  downtime.forEach(function (item) {
    if (!item || item.active !== 'Yes') return;
    if (item.areaKey !== area.key) return;

    const unitsDown = Number(item.unitsDown) || 0;
    const hours = item.fullDay === 'Yes'
      ? shiftHours
      : calculateHoursBetween(item.startTime, item.endTime);

    lost += unitsDown * area.jphPerUnit * hours;
  });

  return lost;
}

function calculateHoursBetween(startValue, endValue) {
  const start = parseTimeToMinutes(startValue);
  const end = parseTimeToMinutes(endValue);
  if (start === null || end === null) return 0;

  let diff = end - start;
  if (diff < 0) diff += 24 * 60;

  return diff / 60;
}

function parseTimeToMinutes(value) {
  if (value === null || value === undefined || value === '') return null;

  let text = String(value).trim();

  // Accept 13.30 or 12.45 as 13:30 / 12:45
  text = text.replace('.', ':');

  // Accept 7, 13 as full hour
  if (/^\d{1,2}$/.test(text)) {
    const hOnly = Number(text);
    if (hOnly < 0 || hOnly > 23) return null;
    return hOnly * 60;
  }

  const match = text.match(/^(\d{1,2}):(\d{1,2})(?:\s*(AM|PM))?$/i);
  if (!match) return null;

  let h = Number(match[1]);
  const m = Number(match[2]);
  const ampm = match[3] ? match[3].toUpperCase() : '';

  if (m < 0 || m > 59) return null;

  if (ampm === 'PM' && h < 12) h += 12;
  if (ampm === 'AM' && h === 12) h = 0;

  if (h < 0 || h > 23) return null;

  return h * 60 + m;
}



function initPerformanceControls() {
  const topDate = document.getElementById('productionDate');
  const perfDate = document.getElementById('perfDateInput');
  const perfShift = document.getElementById('perfShiftModeSelect');
  const mainShift = document.getElementById('shiftMode');
  const today = getLocalDateInputValue(new Date());

  if (topDate && !topDate.value) topDate.value = today;
  if (perfDate && !perfDate.value) perfDate.value = topDate ? topDate.value : today;

  function applyDate(dateValue) {
    const selectedDate = dateValue || today;
    if (topDate) topDate.value = selectedDate;
    if (perfDate) perfDate.value = selectedDate;
    shiftModeManualOverride = false;
    applyAutomaticShiftForDate(selectedDate, true);
  }

  if (topDate) topDate.addEventListener('change', function () { applyDate(topDate.value); });
  if (perfDate) perfDate.addEventListener('change', function () { applyDate(perfDate.value); });

  if (perfShift && mainShift) {
    perfShift.value = mainShift.value || getAutomaticShiftMode(topDate ? topDate.value : today);
    perfShift.addEventListener('change', function () {
      shiftModeManualOverride = true;
      mainShift.value = perfShift.value;
      buildHourlyRows(mainShift.value, true);
      updateShiftControlStatus();
      calculateForecast();
    });

    mainShift.addEventListener('change', function () {
      shiftModeManualOverride = true;
      perfShift.value = mainShift.value;
      buildHourlyRows(mainShift.value, true);
      updateShiftControlStatus();
      calculateForecast();
    });
  }

  applyAutomaticShiftForDate(topDate ? topDate.value : today, false);
}

function getAutomaticShiftMode(dateValue) {
  const date = new Date(`${dateValue || getLocalDateInputValue(new Date())}T12:00:00`);
  if (Number.isNaN(date.getTime())) return 'Weekday';
  const day = date.getDay();
  return day >= 1 && day <= 4 ? 'Weekday' : 'Weekend';
}

function applyAutomaticShiftForDate(dateValue, recalculate) {
  const shiftMode = getAutomaticShiftMode(dateValue);
  const mainShift = document.getElementById('shiftMode');
  const perfShift = document.getElementById('perfShiftModeSelect');
  if (mainShift) mainShift.value = shiftMode;
  if (perfShift) perfShift.value = shiftMode;
  buildHourlyRows(shiftMode, true);
  updateShiftControlStatus();
  updatePerformanceShiftChrome(latestForecastResult ? latestForecastResult.result : null);
  if (recalculate) calculateForecast();
}

function updateShiftControlStatus() {
  setText('shiftControlStatus', shiftModeManualOverride ? 'Manual Override' : 'Calendar Controlled');
}

function updatePerformanceUI(result) {
  updatePerformanceShiftChrome(result);
  renderPerformanceHourlyCards(result);
}

function getShiftWindow(shiftMode) {
  if (shiftMode === 'Weekday OT 12') {
    return { start: '07:00', end: '19:00', label: 'Weekday OT', dayBadge: '⚡ Weekday OT' };
  }

  if (shiftMode === 'Weekend') {
    return { start: '06:30', end: '18:30', label: 'Weekend', dayBadge: '◐ Weekend' };
  }

  return { start: '07:00', end: '17:30', label: 'Weekday', dayBadge: '☀ Weekday' };
}


function getHourlyRowOffsetForShift() {
  return 0;
}

function updatePerformanceShiftChrome(result) {
  const mainShift = document.getElementById('shiftMode');
  const perfShift = document.getElementById('perfShiftModeSelect');
  const shiftMode = (perfShift && perfShift.value) || (mainShift && mainShift.value) || (result && result.shiftMode) || 'Weekday';
  const shiftHours = result && result.shiftHours ? result.shiftHours : getShiftHours(shiftMode);
  const windowInfo = getShiftWindow(shiftMode);
  const startMinutes = parseTimeToMinutes(windowInfo.start) || 420;
  const endMinutes = parseTimeToMinutes(windowInfo.end) || 1050;
  const now = new Date();
  const selectedDate = getValue('productionDate', getLocalDateInputValue(now));
  const isToday = selectedDate === getLocalDateInputValue(now);
  const nowMinutes = isToday ? now.getHours() * 60 + now.getMinutes() : startMinutes;
  const totalMinutes = Math.max(1, endMinutes - startMinutes);
  const elapsedMinutes = Math.max(0, Math.min(totalMinutes, nowMinutes - startMinutes));
  const pct = Math.max(0, Math.min(100, (elapsedMinutes / totalMinutes) * 100));

  if (perfShift && perfShift.value !== shiftMode) perfShift.value = shiftMode;

  setText('perfStartTime', formatMinutesAsClock(startMinutes));
  setText('perfEndTime', formatMinutesAsClock(endMinutes));
  setText('perfProductiveHours', `${round1(shiftHours)} hr`);
  setText('perfShiftDayBadge', windowInfo.dayBadge);
  setText('perfShiftStartLabel', formatMinutesAsClock(startMinutes));
  setText('perfShiftEndLabel', formatMinutesAsClock(endMinutes));
  setText('perfNowTime', formatMinutesAsClock(nowMinutes));
  setText('perfTimelineMeta', `${formatMinutesAsClock(startMinutes)} → ${formatMinutesAsClock(endMinutes)} · ${round1(shiftHours)} productive hr`);

  const dateInput = document.getElementById('perfDateInput');
  if (dateInput && !dateInput.value) dateInput.value = getLocalDateInputValue(now);
  const dateValue = dateInput ? dateInput.value : selectedDate;
  setText('perfDayLabel', getPerformanceDateLabel(dateValue));
  setText('productionDateLabel', getFullProductionDateLabel(dateValue));
  setText('productionDateSetupText', getFullProductionDateLabel(dateValue));
  setText('shiftModeSetupText', getShiftModeDisplay(shiftMode));

  const progress = document.getElementById('perfShiftProgress');
  const remain = document.getElementById('perfShiftRemaining');
  const marker = document.getElementById('perfNowMarker');

  if (progress) progress.style.width = `${pct}%`;
  if (remain) remain.style.width = `${Math.max(0, 100 - pct)}%`;
  if (marker) { marker.style.left = `${pct}%`; marker.style.display = isToday ? '' : 'none'; }
}

function renderPerformanceHourlyCards(result) {
  const holder = document.getElementById('perfHourlyCards');
  if (!holder) return;

  const hourly = result && result.hourlySummary ? result.hourlySummary : {};
  const rows = latestForecastResult && latestForecastResult.payload && Array.isArray(latestForecastResult.payload.hourly)
    ? latestForecastResult.payload.hourly
    : [];

  const shiftMode = result && result.shiftMode ? result.shiftMode : getValue('shiftMode', 'Weekday');
  const windowInfo = getShiftWindow(shiftMode);
  const startMinutes = parseTimeToMinutes(windowInfo.start) || 420;
  const endMinutes = parseTimeToMinutes(windowInfo.end) || 1050;
  const now = new Date();
  const selectedDate = getValue('productionDate', getLocalDateInputValue(now));
  const todayValue = getLocalDateInputValue(now);
  const nowMinutes = selectedDate === todayValue ? now.getHours() * 60 + now.getMinutes() : (selectedDate < todayValue ? endMinutes : startMinutes);
  const targetPace = Number(hourly.targetPace || 0);
  const rowOffset = getHourlyRowOffsetForShift(shiftMode);

  const slots = buildPerformanceHourSlots(startMinutes, endMinutes);

  let cumulativeActual = 0;
  const cards = slots.map(function (slot, index) {
    const row = rows[rowOffset + index] || null;
    const actualThisHour = row && row.active !== 'No' ? Number(row.inspection || 0) || 0 : 0;
    if (actualThisHour > 0) cumulativeActual += actualThisHour;

    const elapsedSlotHours = Math.max(0, (slot.end - startMinutes) / 60);
    const expectedCum = targetPace > 0 ? targetPace * elapsedSlotHours : 0;
    const hasActual = actualThisHour > 0 || (row && row.active !== 'No' && row.inspection !== undefined && Number(row.inspection || 0) > 0);
    const isCurrent = nowMinutes >= slot.start && nowMinutes < slot.end;
    const started = nowMinutes >= slot.start;
    const ratio = expectedCum > 0 ? cumulativeActual / expectedCum : 0;

    let statusClass = 'not-started';
    let statusText = '—';

    if (hasActual || (started && cumulativeActual > 0)) {
      if (ratio >= 1) {
        statusClass = 'good';
        statusText = `${Math.round(ratio * 100)}%`;
      } else if (ratio >= 0.90) {
        statusClass = 'risk';
        statusText = `${Math.round(ratio * 100)}%`;
      } else {
        statusClass = 'behind';
        statusText = `${Math.round(ratio * 100)}%`;
      }
    }

    const actualText = hasActual || cumulativeActual > 0 ? formatNumber(cumulativeActual) : '—';
    const expectedText = expectedCum > 0 ? formatNumber(expectedCum) : '—';
    const fillPct = expectedCum > 0 ? Math.max(0, Math.min(100, ratio * 100)) : 0;

    return `
      <div class="sector-row ${statusClass}${isCurrent ? ' current' : ''}">
        <span class="sector-hour">${escapeHtml(slot.label)}${isCurrent ? '<i class="sector-now-tag">NOW</i>' : ''}</span>
        <div class="sector-track">
          <i class="sector-fill-actual" style="width:${fillPct}%"></i>
        </div>
        <span class="sector-values"><b>${actualText}</b> / ${expectedText}</span>
        <span class="sector-pct">${statusText}</span>
      </div>
    `;
  });

  holder.innerHTML = cards.join('');
}

function buildPerformanceHourSlots(startMinutes, endMinutes) {
  const slots = [];
  let start = startMinutes;

  while (start < endMinutes) {
    // Snap to the next clock-hour mark (7:00, 8:00, 9:00...) rather than just
    // adding 60 minutes to the previous slot's start. This makes the grid read
    // in normal clock hours (7:00-8:00, 8:00-9:00...) instead of drifting by
    // whatever minute the shift happens to start on (6:30-7:30, 7:30-8:30...).
    // Only the first and last slots end up partial, wherever the shift's
    // actual start/end don't land on the hour.
    const nextClockHour = Math.floor(start / 60) * 60 + 60;
    const nextHour = Math.min(endMinutes, nextClockHour);
    // Rows are labeled by their START (the "7:00 AM" row = the 7-8am hour) —
    // that's the existing convention and it's correct for every row except
    // one: the LAST slot of the shift. When shift length is an exact multiple
    // of 60 (Weekend 06:30-18:30, Weekday OT 12 07:00-19:00), that last slot
    // is also a full 60-min block, so start-labeling it silently drops the
    // true shift-end time from every row/tick. Force the final slot only to
    // show the full range so the real end time is always visible somewhere.
    const isFinalSlot = nextHour === endMinutes;
    const label = (nextHour - start < 60 || isFinalSlot)
      ? `${formatMinutesAsClock(start)}–${formatMinutesAsClock(nextHour)}`
      : formatMinutesAsClock(start);

    slots.push({
      start,
      end: nextHour,
      key: minutesToHourKey(start),
      label
    });

    start = nextHour;
  }

  return slots;
}

function normalizeHourKey(value) {
  const minutes = parseTimeToMinutes(value);
  if (minutes === null) return String(value || '');
  return minutesToHourKey(minutes);
}

function minutesToHourKey(minutes) {
  const h = Math.floor(Number(minutes || 0) / 60);
  return `${String(h).padStart(2, '0')}:00`;
}

function formatMinutesAsClock(minutes) {
  let total = Number(minutes || 0);
  total = ((total % 1440) + 1440) % 1440;
  let h = Math.floor(total / 60);
  const m = total % 60;
  const suffix = h >= 12 ? 'PM' : 'AM';
  if (h === 0) h = 12;
  else if (h > 12) h -= 12;
  return `${h}:${String(m).padStart(2, '0')} ${suffix}`;
}

function getLocalDateInputValue(date) {
  const d = date instanceof Date ? date : new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function getFullProductionDateLabel(value) {
  if (!value) return 'Today';
  const date = new Date(`${value}T12:00:00`);
  if (Number.isNaN(date.getTime())) return 'Selected day';
  return date.toLocaleDateString(undefined, {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
  });
}

function getPerformanceDateLabel(value) {
  if (!value) return 'Today';
  const date = new Date(`${value}T12:00:00`);
  if (Number.isNaN(date.getTime())) return 'Selected day';
  const today = getLocalDateInputValue(new Date());
  const label = date.toLocaleDateString(undefined, { weekday: 'long' });
  return value === today ? `${label} · Today` : label;
}

function getShiftHours(shiftMode) {
  if (Object.prototype.hasOwnProperty.call(SHIFT_PRODUCTIVE_HOURS, shiftMode)) {
    return SHIFT_PRODUCTIVE_HOURS[shiftMode];
  }
  return 9.5;
}

function getValue(id, fallback) {
  const el = document.getElementById(id);
  if (!el) return fallback;
  return el.value;
}

function getRowValue(row, selector) {
  const el = row.querySelector(selector);
  return el ? el.value : '';
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function setRiskClass(el, risk) {
  if (!el) return;
  el.classList.remove('risk-green', 'risk-yellow', 'risk-red');

  if (risk === 'GREEN') el.classList.add('risk-green');
  if (risk === 'YELLOW') el.classList.add('risk-yellow');
  if (risk === 'RED') el.classList.add('risk-red');
}

function formatNumber(value) {
  const n = Number(value) || 0;
  return Math.round(n).toLocaleString();
}

function round1(value) {
  const n = Number(value) || 0;
  return (Math.round(n * 10) / 10).toLocaleString();
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

/* =========================================================
   AR FORECAST MODULE V2
   Live Production WIP + RAW_ACTIVITY_CURRENT live staffing + AR Forecast storage.
========================================================= */
const AR_FORECAST_API_URL = 'https://script.google.com/macros/s/AKfycbz3ZhvArQTxy4y_LLmEaDXEYxaZsBDSP5JNGr_H8wvFl5qg5Aofe1zf9QQRdBvVUY3Y/exec';
const AR_FORECAST_STORAGE_KEY = 'productionGoalForecast.ar.v2';
const AR_STAFFING_CAPACITY_OVERRIDES_KEY = 'productionGoalForecast.ar.staffingCapacityOverrides.v1';
const AR_BOS_WIP_STORAGE_KEY = 'productionGoalForecast.ar.bosWipByDate.v1';
const AR_FORECAST_FRONTEND_VERSION = '2026-07-13-ar-bos-stable-v7';
const AR_DEFAULT_CONFIG = Object.freeze({goal:0,runSize:84,maintenanceDelay:50,arIn:40,basket:60,t40:45,oven:60,sectoring:40,chamber:45,deRing:20,arOut:30});
const AR_DEFAULT_JPH = Object.freeze({'AR-IN':120,'Basket':64,'Oven':0,'Sectoring':48,'DeRing':120,'AR-OUT':100});
const AR_DEFAULT_STAFFING_EXCLUSIONS = Object.freeze([
  'CALEB DAY'
]);
const AR_LOCAL_STAFFING_EXCLUSIONS_KEY = 'productionGoalForecast.ar.staffingExclusions.v1';
const AR_STATION_ORDER=['AR-IN','Basket','Oven','Sectoring','DeRing','AR-OUT'];
const AR_FORECAST_STATE={
  department:localStorage.getItem('productionGoalForecast.department')||'Surface',
  config:{...AR_DEFAULT_CONFIG},
  flow:null,
  operatorActivity:null,
  staffingExclusions:loadARLocalStaffingExclusions(),
  staffingCapacityOverrides:loadARStaffingCapacityOverrides(),
  bosWipByDate:loadARBosWipByDate(),
  liveStaff:[],
  hourly:{},
  manualHourly:{},
  calculated:null,
  lastSync:null,
  syncPromise:null,
  syncRequestId:0
};
window.AR_FORECAST_APP={getDepartment:()=>AR_FORECAST_STATE.department};

document.addEventListener('DOMContentLoaded',function initARForecastModuleV2(){restoreARLocalState();wireARForecastControls();setDepartmentMode(AR_FORECAST_STATE.department,false);buildARHourlyRows();calculateARForecast();});

function wireARForecastControls(){const dep=document.getElementById('departmentMode');if(dep){dep.value=AR_FORECAST_STATE.department;dep.addEventListener('change',()=>setDepartmentMode(dep.value,true));}document.getElementById('arStaffingCards')?.addEventListener('click',handleARStaffingClick);document.getElementById('arStaffingCards')?.addEventListener('change',handleARStaffingCapacityChange);document.getElementById('arExcludedAssociates')?.addEventListener('click',handleARStaffingClick);['arGoal','arRunSize','arMaintenanceDelay','arTimeArIn','arTimeBasket','arTimeT40','arTimeOven','arTimeSectoring','arTimeChamber','arTimeDeRing','arTimeArOut'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('input',()=>{readARConfigFromInputs();saveARLocalState();calculateARForecast();});});document.getElementById('arSaveBosBtn')?.addEventListener('click',saveARBosWip);document.getElementById('arResetBosBtn')?.addEventListener('click',resetARBosWip);document.getElementById('arBosWip')?.addEventListener('input',()=>{setTextSafe('arBosSetupMeta','Unsaved BOS change — click Save BOS.');});document.getElementById('arSaveSetupBtn')?.addEventListener('click',saveARSetup);document.getElementById('arResetSetupBtn')?.addEventListener('click',resetARSetup);document.getElementById('arSaveHourlyBtn')?.addEventListener('click',saveARHourly);document.getElementById('arClearHourlyBtn')?.addEventListener('click',clearARManualHourly);document.getElementById('arSaveForecastBtn')?.addEventListener('click',saveARForecast);const btn=document.getElementById('calculateBtn');btn?.addEventListener('click',e=>{if(AR_FORECAST_STATE.department==='AR'){e.stopImmediatePropagation();syncARForecastData(true);}},true);document.getElementById('shiftMode')?.addEventListener('change',()=>{if(AR_FORECAST_STATE.department==='AR'){buildARHourlyRows();calculateARForecast();}});document.getElementById('productionDate')?.addEventListener('change',()=>{if(AR_FORECAST_STATE.department==='AR')syncARForecastData(false);});}
function setDepartmentMode(mode,shouldSync){
  const d=mode==='AR'?'AR':'Surface';
  AR_FORECAST_STATE.department=d;
  localStorage.setItem('productionGoalForecast.department',d);
  document.body.dataset.department=d;

  const s=document.getElementById('departmentMode');
  if(s)s.value=d;

  setTextSafe('departmentBrandMark',d==='AR'?'AR':'SF');
  setTextSafe('departmentBrandName',d==='AR'?'AR':'SURFACE');
  setTextSafe('pageForecastTitle',d==='AR'?'AR Goal Forecast':'Surface Goal Forecast');
  setTextSafe('departmentModeStatus',d==='AR'?'AR WIP + live activity staffing':'Surface engine');

  const tab=document.querySelector('.tab-btn.active')?.dataset.tabTarget||'status';
  activateTab(
    tab,
    Array.from(document.querySelectorAll('.tab-btn')),
    Array.from(document.querySelectorAll('.tab-panel'))
  );

  if(!shouldSync)return;

  if(d==='AR'){
    syncARForecastData(true);
  }else{
    // Surface should refresh its own source only when Surface is selected.
    syncSurfaceDashboardData(true);
    calculateForecast();
  }
}

async function fetchARJson(url){
  const response=await fetch(url,{
    method:'GET',
    cache:'default',
    credentials:'omit'
  });

  if(!response.ok)throw new Error(`API ${response.status}`);

  const payload=await response.json();

  if(payload?.status==='error'){
    throw new Error(payload.message||'Production Flow API error');
  }

  if(payload?.ok===false){
    throw new Error(payload.error||payload.message||'API error');
  }

  return payload;
}

function getARShiftType(){
  return getValue('shiftMode','Weekday').includes('Weekend')?'Weekend':'Weekday';
}

async function syncARForecastData(showStatus){
  // Do not start another Google Apps Script request while one is already running.
  if(AR_FORECAST_STATE.syncPromise){
    if(showStatus)setARSyncState('syncing','AR refresh already running');
    return AR_FORECAST_STATE.syncPromise;
  }

  const requestId=++AR_FORECAST_STATE.syncRequestId;

  AR_FORECAST_STATE.syncPromise=(async()=>{
    if(showStatus)setARSyncState('syncing','Syncing AR WIP + live staffing');

    readARConfigFromInputs();

    try{
      // These two endpoints are independent, so request them together.
      // debug=true was intentionally removed so the Apps Script 30-second cache works.
      const [flowResult,operatorResult]=await Promise.allSettled([
        fetchARJson(`${SURFACE_DASHBOARD_API_URL}?action=productionFlow&area=AR`),
        fetchARJson(`${SURFACE_DASHBOARD_API_URL}?action=operatorActivity&area=AR`)
      ]);

      // Ignore a result only if a newer request somehow superseded this one.
      if(requestId!==AR_FORECAST_STATE.syncRequestId)return;

      if(flowResult.status==='fulfilled'){
        AR_FORECAST_STATE.flow=flowResult.value;
      }else{
        console.warn('[AR Forecast] Flow request failed:',flowResult.reason);
      }

      if(operatorResult.status==='fulfilled'){
        AR_FORECAST_STATE.operatorActivity=operatorResult.value;
      }else{
        console.warn('[AR Forecast] Operator request failed:',operatorResult.reason);
      }

      AR_FORECAST_STATE.staffingExclusions=loadARLocalStaffingExclusions();

      if(!AR_FORECAST_STATE.flow){
        throw new Error('Live AR WIP unavailable');
      }

      AR_FORECAST_STATE.liveStaff=buildARLiveStaffing(
        AR_FORECAST_STATE.operatorActivity
      ).filter(person=>!isARStaffingExcluded(person.name));

      mergeARHourlyFromLiveSources(
        AR_FORECAST_STATE.flow,
        AR_FORECAST_STATE.operatorActivity
      );

      AR_FORECAST_STATE.lastSync=new Date();

      // Render each expensive section once, after both API responses are resolved.
      buildARHourlyRows();
      renderARStaffingCards();
      renderARExcludedAssociates();
      calculateARForecast();

      setARSyncState(
        'ok',
        AR_FORECAST_STATE.liveStaff.length
          ? `AR live · ${AR_FORECAST_STATE.liveStaff.length} associates`
          : 'AR live · waiting for first scans'
      );
    }catch(err){
      console.error('[AR Forecast]',err);

      setARSyncState('error','AR sync failed');

      // Keep and display the last successful data instead of blanking the page.
      renderARStaffingCards();
      renderARExcludedAssociates();
      calculateARForecast();
    }finally{
      AR_FORECAST_STATE.syncPromise=null;
    }
  })();

  return AR_FORECAST_STATE.syncPromise;
}

function setARSyncState(state,text){if(AR_FORECAST_STATE.department!=='AR')return;setTextSafe('heroSyncState',text);setTextSafe('heroSyncTime',new Date().toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit'}));setTextSafe('arRosterSyncStatus',text);const dot=document.getElementById('heroSyncDot');if(dot)dot.className=`hero-sync-dot dot-${state}`;}


function loadARStaffingCapacityOverrides(){
  try{
    const raw=JSON.parse(localStorage.getItem(AR_STAFFING_CAPACITY_OVERRIDES_KEY)||'{}');
    const out={};
    AR_STATION_ORDER.forEach(station=>{
      if(raw[station]===undefined||raw[station]===null||raw[station]==='')return;
      const value=Number(raw[station]);
      if(Number.isFinite(value)&&value>=0)out[station]=Math.round(value*2)/2;
    });
    return out;
  }catch{
    return {};
  }
}

function saveARStaffingCapacityOverrides(){
  localStorage.setItem(
    AR_STAFFING_CAPACITY_OVERRIDES_KEY,
    JSON.stringify(AR_FORECAST_STATE.staffingCapacityOverrides||{})
  );
}

function getARCapacityAssociateCount(station,detectedCount){
  const value=AR_FORECAST_STATE.staffingCapacityOverrides?.[station];
  return Number.isFinite(Number(value))
    ? Math.max(0,Math.round(Number(value)*2)/2)
    : detectedCount;
}

function setARCapacityAssociateOverride(station,value){
  if(!AR_STATION_ORDER.includes(station)||station==='Oven')return;

  const numeric=Number(value);
  if(!Number.isFinite(numeric)||numeric<0)return;

  AR_FORECAST_STATE.staffingCapacityOverrides[station]=Math.round(numeric*2)/2;
  saveARStaffingCapacityOverrides();
}

function clearARCapacityAssociateOverride(station){
  if(!AR_FORECAST_STATE.staffingCapacityOverrides)return;
  delete AR_FORECAST_STATE.staffingCapacityOverrides[station];
  saveARStaffingCapacityOverrides();
}

function handleARStaffingCapacityChange(event){
  const input=event.target.closest('[data-ar-capacity-station]');
  if(!input)return;

  const station=String(input.dataset.arCapacityStation||'').trim();
  setARCapacityAssociateOverride(station,input.value);
  calculateARForecast();
  renderARStaffingCards();
  showARMessage(`${station} capacity count set to ${round1(input.value)} associate(s).`);
}

function normalizeAROperatorKey(name){return String(name||'').trim().replace(/\s+/g,' ').toUpperCase();}
function loadARLocalStaffingExclusions(){
  try{
    const saved=JSON.parse(localStorage.getItem(AR_LOCAL_STAFFING_EXCLUSIONS_KEY)||'[]');
    return Array.isArray(saved)?saved.map(normalizeAROperatorKey).filter(Boolean):[];
  }catch{return [];}
}
function saveARLocalStaffingExclusions(){
  const cleaned=Array.from(new Set((AR_FORECAST_STATE.staffingExclusions||[]).map(normalizeAROperatorKey).filter(Boolean)));
  AR_FORECAST_STATE.staffingExclusions=cleaned;
  localStorage.setItem(AR_LOCAL_STAFFING_EXCLUSIONS_KEY,JSON.stringify(cleaned));
}
function isARDefaultStaffingExclusion(name){
  const key=normalizeAROperatorKey(name);
  return AR_DEFAULT_STAFFING_EXCLUSIONS.some(item=>normalizeAROperatorKey(item)===key);
}
function isARStaffingExcluded(name){
  const key=normalizeAROperatorKey(name);
  return isARDefaultStaffingExclusion(key)||(AR_FORECAST_STATE.staffingExclusions||[]).some(item=>normalizeAROperatorKey(item)===key);
}
function refreshARStaffingAfterExclusionChange(){
  AR_FORECAST_STATE.liveStaff=buildARLiveStaffing(AR_FORECAST_STATE.operatorActivity).filter(person=>!isARStaffingExcluded(person.name));
  calculateARForecast();
  renderARStaffingCards();
  renderARExcludedAssociates();
  setARSyncState('ok',AR_FORECAST_STATE.liveStaff.length?`AR live · ${AR_FORECAST_STATE.liveStaff.length} associates`:'AR live · waiting for first scans');
}
function handleARStaffingClick(event){
  const button=event.target.closest('[data-ar-staff-action]');
  if(!button)return;
  const action=button.dataset.arStaffAction;

  if(action==='reset-capacity'){
    const station=String(button.dataset.station||'').trim();
    clearARCapacityAssociateOverride(station);
    calculateARForecast();
    renderARStaffingCards();
    showARMessage(`${station} capacity count returned to automatic live count.`);
    return;
  }

  const operator=String(button.dataset.operator||'').trim();
  if(!operator)return;
  const key=normalizeAROperatorKey(operator);
  if(action==='hide'){
    if(!AR_FORECAST_STATE.staffingExclusions.includes(key))AR_FORECAST_STATE.staffingExclusions.push(key);
    saveARLocalStaffingExclusions();
    showARMessage(`${operator} hidden from AR staffing on this browser.`);
  }else if(action==='restore'){
    if(isARDefaultStaffingExclusion(key)){
      showARMessage(`${operator} is a default JS exclusion and cannot be restored from the page.`);
      return;
    }
    AR_FORECAST_STATE.staffingExclusions=(AR_FORECAST_STATE.staffingExclusions||[]).filter(item=>normalizeAROperatorKey(item)!==key);
    saveARLocalStaffingExclusions();
    showARMessage(`${operator} restored to AR staffing on this browser.`);
  }
  refreshARStaffingAfterExclusionChange();
}
function renderARExcludedAssociates(){
  const box=document.getElementById('arExcludedAssociates');
  const count=document.getElementById('arExcludedCount');
  if(!box)return;
  const defaults=AR_DEFAULT_STAFFING_EXCLUSIONS.map(normalizeAROperatorKey).filter(Boolean);
  const local=(AR_FORECAST_STATE.staffingExclusions||[]).map(normalizeAROperatorKey).filter(Boolean).filter(name=>!defaults.includes(name));
  const excluded=[...defaults.map(name=>({name,source:'Default JS exclusion'})),...local.map(name=>({name,source:'Hidden on this browser'}))];
  if(count)count.textContent=String(excluded.length);
  box.innerHTML=excluded.length?excluded.map(item=>`<div class="ar-excluded-row"><div><strong>${escapeHtml(item.name)}</strong><small>${escapeHtml(item.source)}</small></div>${item.source==='Default JS exclusion'?'<span class="ar-default-exclusion-badge">DEFAULT</span>':`<button type="button" class="ar-restore-btn" data-ar-staff-action="restore" data-operator="${escapeHtml(item.name)}">Restore</button>`}</div>`).join(''):'<div class="ar-excluded-empty">No associates are hidden.</div>';
}
function getARRowsFromFlow(){const rows=Array.isArray(AR_FORECAST_STATE.flow?.productionFlow)?AR_FORECAST_STATE.flow.productionFlow:[];return rows.filter(r=>String(r.Area||'').trim().toUpperCase()==='AR');}
function findARFlowRow(names,bridgeRole){const rows=getARRowsFromFlow();if(bridgeRole){const x=rows.find(r=>String(r.BridgeRole||'').toUpperCase()===bridgeRole);if(x)return x;}const wanted=names.map(n=>n.toUpperCase());return rows.find(r=>wanted.includes(String(r.FlowStep||r.DisplayName||'').trim().toUpperCase()))||{};}

function loadARBosWipByDate(){
  try{
    const value=JSON.parse(localStorage.getItem(AR_BOS_WIP_STORAGE_KEY)||'{}');
    return value&&typeof value==='object'?value:{};
  }catch{
    return {};
  }
}

function saveARBosWipByDate(){
  localStorage.setItem(
    AR_BOS_WIP_STORAGE_KEY,
    JSON.stringify(AR_FORECAST_STATE.bosWipByDate||{})
  );
}

function getARProductionDateKey(){
  return getValue('productionDate',getLocalDateInputValue(new Date()));
}

function calculateCurrentARInsideWip(live){
  return toARNumber(live?.arIn)+
    toARNumber(live?.basket)+
    toARNumber(live?.oven)+
    toARNumber(live?.sectoring)+
    toARNumber(live?.deRing);
}

function captureARBosWipIfMissing(live){
  const dateKey=getARProductionDateKey();
  if(!dateKey)return 0;

  const saved=AR_FORECAST_STATE.bosWipByDate?.[dateKey];
  if(Number.isFinite(Number(saved))){
    return Math.max(0,toARNumber(saved));
  }

  const captured=calculateCurrentARInsideWip(live);
  AR_FORECAST_STATE.bosWipByDate[dateKey]=captured;
  saveARBosWipByDate();
  return captured;
}

function getARBosWip(live){
  return captureARBosWipIfMissing(live);
}


function saveARBosWip(){
  const dateKey=getARProductionDateKey();
  const bosInput=document.getElementById('arBosWip');
  const value=Math.max(0,toARNumber(bosInput?.value));

  AR_FORECAST_STATE.bosWipByDate[dateKey]=value;
  saveARBosWipByDate();
  calculateARForecast();

  showARMessage(
    `BOS AR WIP saved at ${formatARNumber(value)} jobs for ${dateKey}.`
  );

  setTextSafe(
    'arBosSetupMeta',
    `BOS saved for ${dateKey}. Reset Live pulls current AR WIP.`
  );
}

function resetARBosWip(){
  const live=getARLiveValues();
  const dateKey=getARProductionDateKey();
  const captured=calculateCurrentARInsideWip(live);

  AR_FORECAST_STATE.bosWipByDate[dateKey]=captured;
  saveARBosWipByDate();
  calculateARForecast();

  showARMessage(
    `BOS AR WIP reset from live WIP to ${formatARNumber(captured)} jobs for ${dateKey}.`
  );
}

function renderARBosSetup(live,bosWip){
  const bosInput=document.getElementById('arBosWip');
  const feedInput=document.getElementById('arSurfaceFeedLive');
  const dateKey=getARProductionDateKey();

  if(bosInput&&document.activeElement!==bosInput){
    bosInput.value=Math.round(bosWip);
  }
  if(feedInput)feedInput.value=Math.round(toARNumber(live?.surface));

  if(document.activeElement!==bosInput){
    setTextSafe(
      'arBosSetupMeta',
      `BOS saved for ${dateKey}. Reset Live pulls current AR WIP.`
    );
  }
}

function getARLiveValues(){const row=(n,r)=>findARFlowRow(n,r);return{surface:toARNumber(row(['Surface Inspection','Surface Inspection Input WIP'],'AR_INPUT').CurrentWIP),arIn:toARNumber(row(['AR-IN']).CurrentWIP),basket:toARNumber(row(['Basket']).CurrentWIP),oven:toARNumber(row(['Oven']).CurrentWIP),sectoring:toARNumber(row(['Sectoring']).CurrentWIP),deRing:toARNumber(row(['DeRing']).CurrentWIP),arOut:toARNumber(row(['AR-OUT'],'AR_OUTPUT').ActivityToday)};}
function toARBool(v){return v===true||String(v).toLowerCase()==='true'||String(v)==='1'||String(v).toLowerCase()==='yes';}
function normalizeARRole(value){const s=String(value||'').trim().toUpperCase().replace(/_/g,'-');if(s==='AR IN'||s.includes('AR-IN'))return'AR-IN';if(s.includes('BASKET'))return'Basket';if(s.includes('OVEN'))return'Oven';if(s.includes('SECTOR'))return'Sectoring';if(s.includes('DERING')||s.includes('DE-RING'))return'DeRing';if(s==='AR OUT'||s.includes('AR-OUT'))return'AR-OUT';return String(value||'').trim();}
function getARHourRank(value){const key=normalizeARHourKey(value);const parts=key.split(':').map(Number);return (parts[0]||0)*60+(parts[1]||0);}
function isValidAROperator(name){const key=String(name||'').trim().toUpperCase();return Boolean(key&&key!=='UNASSIGNED / NO OPERATOR'&&!key.includes('NO OPERATOR')&&!key.includes('UNASSIGNED'));}
function buildARLiveStaffing(payload){
  const rows=Array.isArray(payload?.operatorActivity)?payload.operatorActivity:[];
  const byOperator={};

  rows.forEach(row=>{
    const name=String(row.Operator??row.operator??'').trim();
    if(!isValidAROperator(name))return;

    const station=normalizeARRole(
      row.FlowStation??row.flowStation??row.AccessPoint??row.accessPoint??''
    );

    if(!AR_STATION_ORDER.includes(station))return;

    const hours=row.Hours??row.hours??{};

    Object.entries(hours).forEach(([hour,value])=>{
      const output=toARNumber(value);
      if(output<=0)return;

      const rank=getARHourRank(hour);
      const stationRank=AR_STATION_ORDER.indexOf(station);
      const key=name.toUpperCase();
      const current=byOperator[key];

      const candidate={
        name,
        station,
        lastHour:normalizeARHourKey(hour),
        lastHourLabel:formatARHourLabel(hour),
        lastHourOutput:output,
        dailyOutput:toARNumber(
          row.Total??row.total??row.HourlyTotal??row.hourlyTotal
        ),
        accessPoint:String(row.AccessPoint??row.accessPoint??station),
        rank,
        stationRank,
        otherStations:[]
      };

      if(!current){
        byOperator[key]=candidate;
        return;
      }

      // A newer hour always wins.
      if(rank>current.rank){
        byOperator[key]=candidate;
        return;
      }

      if(rank<current.rank)return;

      // During the same hour, the furthest downstream AR station wins.
      // This prevents a large Oven batch from keeping somebody assigned to Oven
      // after they have already scanned work at DeRing or AR-OUT.
      if(station!==current.station){
        const currentStationRank=AR_STATION_ORDER.indexOf(current.station);

        if(stationRank>currentStationRank){
          candidate.otherStations=[
            {station:current.station,output:current.lastHourOutput},
            ...(current.otherStations||[])
          ];
          byOperator[key]=candidate;
        }else{
          const exists=(current.otherStations||[]).some(
            item=>item.station===station
          );

          if(!exists){
            current.otherStations.push({station,output});
          }
        }

        return;
      }

      // Same hour and same station: retain the strongest row.
      if(output>current.lastHourOutput){
        current.lastHourOutput=output;
        current.dailyOutput=Math.max(
          current.dailyOutput,
          candidate.dailyOutput
        );
        current.accessPoint=candidate.accessPoint;
      }
    });
  });

  return Object.values(byOperator)
    .map(person=>({
      ...person,
      finalJph:
        person.station==='Oven'
          ? 0
          : (AR_DEFAULT_JPH[person.station]||0),
      status:person.otherStations.length
        ? 'Moved / Multi-area'
        : 'Active'
    }))
    .sort((a,b)=>
      AR_STATION_ORDER.indexOf(a.station)-
      AR_STATION_ORDER.indexOf(b.station)||
      a.name.localeCompare(b.name)
    );
}
function getARStationRoster(){const out={};AR_STATION_ORDER.forEach(s=>out[s]=[]);AR_FORECAST_STATE.liveStaff.forEach(person=>{if(out[person.station])out[person.station].push({...person,weight:1,effectiveJph:person.finalJph});});return out;}
function getARStationCapacities(remainingHours){
  const roster=getARStationRoster();

  return AR_STATION_ORDER.map(station=>{
    const associates=roster[station]||[];
    const detectedAssociates=associates.length;
    const currentAssociates=station==='Oven'
      ? detectedAssociates
      : getARCapacityAssociateCount(station,detectedAssociates);
    const stationJph=station==='Oven'?0:(AR_DEFAULT_JPH[station]||0);
    const usableJph=station==='Oven'?0:currentAssociates*stationJph;

    return{
      station,
      associates,
      detectedAssociates,
      currentAssociates,
      stationJph,
      usableJph,
      isManual:
        station!=='Oven'&&
        Number.isFinite(Number(AR_FORECAST_STATE.staffingCapacityOverrides?.[station])),
      remainingCapacity:station==='Oven'
        ? Infinity
        : Math.max(0,usableJph*remainingHours)
    };
  });
}
function mergeARHourlyFromLiveSources(flowPayload,operatorPayload){
  const combined={};
  const flowRows=Array.isArray(flowPayload?.hourlyActivity)?flowPayload.hourlyActivity:[];
  flowRows.forEach(row=>{
    const station=normalizeARRole(row.FlowStep??row.DisplayName??row.AccessPoint??'');
    if(station!=='AR-IN'&&station!=='AR-OUT')return;
    const hours=row.Hours??row.hours??{};
    Object.entries(hours).forEach(([hour,value])=>addARHour(combined,hour,station==='AR-IN'?toARNumber(value):0,station==='AR-OUT'?toARNumber(value):0));
  });
  if(!Object.keys(combined).length){
    const rows=Array.isArray(operatorPayload?.operatorActivity)?operatorPayload.operatorActivity:[];
    rows.forEach(row=>{
      const station=normalizeARRole(row.FlowStation??row.flowStation??row.AccessPoint??row.accessPoint??'');
      if(station!=='AR-IN'&&station!=='AR-OUT')return;
      Object.entries(row.Hours??row.hours??{}).forEach(([hour,value])=>addARHour(combined,hour,station==='AR-IN'?toARNumber(value):0,station==='AR-OUT'?toARNumber(value):0));
    });
  }
  AR_FORECAST_STATE.hourly=combined;
}
function addARHour(target,key,arin,arout){const h=normalizeARHourKey(key);if(!target[h])target[h]={arIn:0,arOut:0};target[h].arIn+=arin;target[h].arOut+=arout;}
function buildARHourlyRows(){
  const body=document.getElementById('arHourlyRows');
  if(!body)return;

  const w=getShiftWindow(getValue('shiftMode','Weekday'));
  const start=parseTimeToMinutes(w.start)||420;
  const end=parseTimeToMinutes(w.end)||1050;
  const slots=buildPerformanceHourSlots(start,end);

  body.innerHTML='';

  slots.forEach(slot=>{
    const key=minutesToHourKey(Math.floor(slot.start/60)*60);
    const api=AR_FORECAST_STATE.hourly[key]||{};
    const manual=AR_FORECAST_STATE.manualHourly[key]||{};

    // AR-IN and AR-OUT always come from RAW_ACTIVITY_CURRENT.
    // Old saved manual zeroes are intentionally ignored.
    const arIn=toARNumber(api.arIn);
    const arOut=toARNumber(api.arOut);

    const maint=Boolean(manual.maintenance);
    const delay=maint
      ? toARNumber(manual.delay||AR_FORECAST_STATE.config.maintenanceDelay)
      : 0;

    const tr=document.createElement('tr');
    tr.className='ar-hour-row';
    tr.dataset.hour=key;

    tr.innerHTML=`
      <td><input class="ar-hour-label" value="${escapeHtml(slot.label)}" readonly></td>
      <td>
        <input class="ar-hour-in ar-hour-auto" type="number" min="0" step="1"
          value="${arIn||''}" readonly title="Automatic from RAW_ACTIVITY_CURRENT">
      </td>
      <td>
        <input class="ar-hour-out ar-hour-auto" type="number" min="0" step="1"
          value="${arOut||''}" readonly title="Automatic from RAW_ACTIVITY_CURRENT">
      </td>
      <td class="ar-runs-in">${completeARRuns(arIn)}</td>
      <td class="ar-runs-out">${completeARRuns(arOut)}</td>
      <td><input class="ar-hour-maint" type="checkbox" ${maint?'checked':''}></td>
      <td>
        <input class="ar-delay-input" type="number" min="0" step="5"
          value="${delay||AR_FORECAST_STATE.config.maintenanceDelay}" ${maint?'':'disabled'}>
      </td>
      <td class="ar-projection-cell">--</td>
      <td><input class="ar-hour-note" type="text" value="${escapeHtml(manual.note||'')}"></td>`;

    body.appendChild(tr);

    tr.querySelector('.ar-hour-maint')?.addEventListener('change',()=>updateARHourlyRow(tr));
    tr.querySelector('.ar-delay-input')?.addEventListener('input',()=>updateARHourlyRow(tr));
    tr.querySelector('.ar-hour-note')?.addEventListener('input',()=>updateARHourlyRow(tr));

    updateARHourlyRow(tr,false);
  });
}

function updateARHourlyRow(row,save=true){
  const key=row.dataset.hour;
  const maintenance=row.querySelector('.ar-hour-maint')?.checked;
  const delayInput=row.querySelector('.ar-delay-input');

  if(delayInput)delayInput.disabled=!maintenance;

  const api=AR_FORECAST_STATE.hourly[key]||{};
  const arIn=toARNumber(api.arIn);
  const arOut=toARNumber(api.arOut);

  const data={
    maintenance,
    delay:maintenance
      ? toARNumber(delayInput?.value||AR_FORECAST_STATE.config.maintenanceDelay)
      : 0,
    note:String(row.querySelector('.ar-hour-note')?.value||'')
  };

  // Save only maintenance, delay, and note.
  // AR-IN / AR-OUT are always refreshed from the Production Flow API.
  AR_FORECAST_STATE.manualHourly[key]=data;

  const arInInput=row.querySelector('.ar-hour-in');
  const arOutInput=row.querySelector('.ar-hour-out');

  if(arInInput)arInInput.value=arIn||'';
  if(arOutInput)arOutInput.value=arOut||'';

  row.querySelector('.ar-runs-in').textContent=completeARRuns(arIn);
  row.querySelector('.ar-runs-out').textContent=completeARRuns(arOut);
  row.querySelector('.ar-projection-cell').textContent=getARProjectedCompletionLabel(
    key,
    {arIn,arOut,maintenance,delay:data.delay,note:data.note}
  );

  if(save){
    saveARLocalState();
    calculateARForecast();
  }
}

function getARProjectedCompletionLabel(key,data){if(!data.arIn)return'--';const[h,m]=key.split(':').map(Number),d=new Date();d.setHours(h,m||0,0,0);d.setMinutes(d.getMinutes()+getARTotalPipelineMinutes()+toARNumber(data.delay));return d.toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit'});}

function estimateARMaturedJobs(live,remainingMinutes,totalDelayMinutes){
  const c=AR_FORECAST_STATE.config;
  const usableMinutes=Math.max(
    0,
    toARNumber(remainingMinutes)-toARNumber(totalDelayMinutes)
  );

  // RAW WIP represents the last completed scan location.
  // Count a station's current WIP only when enough shift time remains
  // for that work to clear every downstream AR process and AR-OUT.
  const stages=[
    {
      jobs:toARNumber(live?.deRing),
      minutes:c.arOut
    },
    {
      jobs:toARNumber(live?.sectoring),
      minutes:c.chamber+c.deRing+c.arOut
    },
    {
      jobs:toARNumber(live?.oven),
      minutes:c.sectoring+c.chamber+c.deRing+c.arOut
    },
    {
      jobs:toARNumber(live?.basket),
      minutes:c.t40+c.oven+c.sectoring+c.chamber+c.deRing+c.arOut
    },
    {
      jobs:toARNumber(live?.arIn),
      minutes:c.basket+c.t40+c.oven+c.sectoring+c.chamber+c.deRing+c.arOut
    }
  ];

  return stages.reduce((sum,stage)=>{
    return sum+(usableMinutes>=stage.minutes?stage.jobs:0);
  },0);
}


function projectAROutFromHourly(currentOut,hourly,remainingHours){
  const rows=Array.isArray(hourly)?hourly:[];
  const now=new Date();
  const nowMinutes=now.getHours()*60+now.getMinutes();

  const completedRows=rows.filter(row=>{
    const hourMinutes=parseTimeToMinutes(row.hour);
    return hourMinutes!==null&&hourMinutes<=nowMinutes;
  });

  const usableRows=completedRows.length?completedRows:rows;
  const outputs=usableRows.map(row=>toARNumber(row.arOut));

  // Use elapsed scheduled hours, including zero-output hours, so the pace
  // projection does not become artificially high by ignoring downtime.
  const elapsedHours=Math.max(0,usableRows.length);
  const produced=outputs.reduce((sum,value)=>sum+value,0);

  if(elapsedHours<=0||produced<=0){
    return Math.max(0,toARNumber(currentOut));
  }

  const averageOutPerHour=produced/elapsedHours;
  const additionalProjection=Math.max(
    0,
    averageOutPerHour*Math.max(0,toARNumber(remainingHours))
  );

  return Math.max(
    toARNumber(currentOut),
    toARNumber(currentOut)+additionalProjection
  );
}


function getARWipFocus(live){
  return [
    ['AR-IN',toARNumber(live?.arIn)],
    ['Basket / T40',toARNumber(live?.basket)],
    ['Oven / Cooling',toARNumber(live?.oven)],
    ['Sector / Chamber',toARNumber(live?.sectoring)],
    ['DeRing',toARNumber(live?.deRing)]
  ].sort((a,b)=>b[1]-a[1])[0]||['No WIP',0];
}

function getARLaborFocus(caps,live){
  const wipByStation={
    'AR-IN':toARNumber(live?.arIn),
    'Basket':toARNumber(live?.basket),
    'Oven':toARNumber(live?.oven),
    'Sectoring':toARNumber(live?.sectoring),
    'DeRing':toARNumber(live?.deRing),
    'AR-OUT':toARNumber(live?.deRing)
  };

  const ranked=(Array.isArray(caps)?caps:[])
    .filter(row=>row.station!=='Oven')
    .map(row=>{
      const wip=wipByStation[row.station]||0;
      const usableJph=toARNumber(row.usableJph);

      return{
        ...row,
        wip,
        clearMin:usableJph>0?(wip/usableJph)*60:Infinity
      };
    })
    .sort((a,b)=>b.clearMin-a.clearMin);

  const focus=ranked[0];

  if(!focus){
    return{
      name:'No staffing',
      reason:'No AR station capacity available'
    };
  }

  return{
    name:focus.station,
    reason:Number.isFinite(focus.clearMin)
      ? `${Math.round(focus.clearMin)} min to clear ${formatARNumber(focus.wip)} WIP at ${round1(focus.usableJph)} JPH`
      : 'No usable JPH assigned'
  };
}

function getARExpectedByNow(goal){
  const numericGoal=toARNumber(goal);
  if(numericGoal<=0)return 0;

  const shift=getShiftWindow(getValue('shiftMode','Weekday'));
  const start=parseTimeToMinutes(shift.start)||420;
  const end=parseTimeToMinutes(shift.end)||1050;
  const now=new Date();
  const nowMinutes=now.getHours()*60+now.getMinutes();

  const ratio=Math.max(
    0,
    Math.min(1,(nowMinutes-start)/Math.max(1,end-start))
  );

  return numericGoal*ratio;
}

function calculateARForecast(){
  readARConfigFromInputs();

  const live=getARLiveValues();
  const runSize=Math.max(1,AR_FORECAST_STATE.config.runSize);
  const liveInside=calculateCurrentARInsideWip(live);
  const bosWip=getARBosWip(live);
  const surfaceFeed=toARNumber(live.surface);
  const totalWorkAvailable=Math.max(0,bosWip+surfaceFeed);

  const hourly=collectARHourlyData();
  const totalDelay=hourly.reduce(
    (sum,row)=>sum+(row.maintenance?row.delay:0),
    0
  );

  const currentOut=Math.max(
    toARNumber(live.arOut),
    hourly.reduce((sum,row)=>sum+toARNumber(row.arOut),0)
  );

  const windowInfo=getShiftWindow(getValue('shiftMode','Weekday'));
  const now=new Date();
  const endMinutes=parseTimeToMinutes(windowInfo.end)||1050;
  const endDate=new Date();
  endDate.setHours(Math.floor(endMinutes/60),endMinutes%60,0,0);

  const remainingMin=Math.max(0,(endDate-now)/60000);
  const remainingHours=remainingMin/60;

  const maturedJobs=estimateARMaturedJobs(live,remainingMin,totalDelay);
  const stationCaps=getARStationCapacities(remainingHours);
  const laborCaps=stationCaps
    .filter(row=>row.station!=='Oven')
    .map(row=>row.remainingCapacity);
  const laborCeiling=laborCaps.length?Math.min(...laborCaps):0;

  const remainingAvailableWork=Math.max(
    0,
    totalWorkAvailable-currentOut
  );

  const paceProjection=projectAROutFromHourly(
    currentOut,
    hourly,
    remainingHours
  );

  const pipelinePotential=Math.max(
    0,
    Math.min(
      remainingAvailableWork,
      maturedJobs+surfaceFeed
    )
  );

  const laborAdditional=laborCeiling>0
    ? Math.min(pipelinePotential,laborCeiling)
    : 0;

  const projectedAdditional=Math.max(
    0,
    Math.min(
      remainingAvailableWork,
      pipelinePotential,
      laborAdditional||pipelinePotential
    )
  );

  const capacityProjection=currentOut+projectedAdditional;
  const projectedOut=Math.max(
    currentOut,
    Math.round(
      Math.min(
        totalWorkAvailable,
        paceProjection>currentOut
          ? Math.max(capacityProjection,paceProjection)
          : capacityProjection
      )
    )
  );

  const low=Math.max(
    currentOut,
    Math.round(Math.min(projectedOut,paceProjection||projectedOut))
  );
  const high=Math.max(
    projectedOut,
    Math.round(Math.min(totalWorkAvailable,capacityProjection))
  );

  const runsCompleted=completeARRuns(currentOut);
  const runsProjected=completeARRuns(projectedOut);
  const bosRuns=bosWip/runSize;
  const feedRuns=surfaceFeed/runSize;
  const projectedRunsExact=projectedOut/runSize;
  const projectedEosWip=Math.max(0,totalWorkAvailable-projectedOut);

  const focus=getARLaborFocus(stationCaps,live);
  const wipFocus=getARWipFocus(live);
  const goal=AR_FORECAST_STATE.config.goal;
  const expectedNow=getARExpectedByNow(goal,currentOut);
  const pacePct=expectedNow>0?currentOut/expectedNow:0;

  AR_FORECAST_STATE.calculated={
    live,
    liveInside,
    bosWip,
    bosRuns,
    surfaceFeed,
    feedRuns,
    totalWorkAvailable,
    remainingAvailableWork,
    projectedEosWip,
    projectedRunsExact,
    currentOut,
    projectedOut,
    low,
    high,
    runsCompleted,
    runsProjected,
    totalDelay,
    focus,
    wipFocus,
    remainingMin,
    remainingHours,
    stationCaps,
    maturedJobs,
    laborCeiling,
    availableJobs:totalWorkAvailable,
    goal,
    expectedNow
  };

  renderARBosSetup(live,bosWip);
  renderARForecastAll(AR_FORECAST_STATE.calculated,hourly,pacePct);
}
function renderARForecastAll(r,hourly,pacePct){
  setTextSafe('arKpiCurrentOut',formatARNumber(r.currentOut));
  setTextSafe('arKpiProjectedOut',formatARNumber(r.projectedOut));

  const arGoal=Math.max(0,toARNumber(r.goal));
  const arExpectedNow=Math.max(0,toARNumber(r.expectedNow));
  const arProjectedGap=arGoal>0?r.projectedOut-arGoal:0;
  const arPaceGap=r.currentOut-arExpectedNow;
  const arRemainingHours=Math.max(0,toARNumber(r.remainingHours));
  const arJobsNeeded=Math.max(0,arGoal-r.currentOut);
  const arRecoveryJph=arGoal>0&&arRemainingHours>0
    ? arJobsNeeded/arRemainingHours
    : 0;

  setTextSafe('arKpiGoal',formatARNumber(arGoal));
  setTextSafe(
    'arKpiGoalMeta',
    arGoal>0
      ? `${Math.abs(Math.round(arProjectedGap))} jobs ${arProjectedGap>=0?'above':'below'} projected goal`
      : 'No AR goal entered'
  );

  setTextSafe('arKpiExpectedNow',formatARNumber(Math.round(arExpectedNow)));
  setTextSafe(
    'arKpiPaceMeta',
    arGoal>0
      ? `${Math.abs(Math.round(arPaceGap))} jobs ${arPaceGap>=0?'ahead of':'behind'} pace`
      : 'Enter an AR goal to compare pace'
  );

  setTextSafe('arKpiRecoveryJph',`${Math.round(arRecoveryJph)} JPH`);
  setTextSafe(
    'arKpiRecoveryMeta',
    arGoal<=0
      ? 'No AR goal entered'
      : arProjectedGap>=0
        ? 'No recovery needed at projected pace'
        : `${formatARNumber(Math.max(0,arGoal-r.projectedOut))} projected jobs still needed`
  );

  setTextSafe(
    'arKpiGoalGap',
    arGoal>0
      ? `${arProjectedGap>=0?'+':'-'}${formatARNumber(Math.abs(Math.round(arProjectedGap)))}`
      : '0'
  );
  setTextSafe(
    'arKpiGoalGapMeta',
    arGoal>0
      ? `Projected ${formatARNumber(r.projectedOut)} vs target ${formatARNumber(arGoal)}`
      : 'Projected vs target'
  );

  const arGoalCard=document.getElementById('arKpiGoalCard');
  const arExpectedCard=document.getElementById('arKpiExpectedCard');
  const arRecoveryCard=document.getElementById('arKpiRecoveryCard');
  const arGapCard=document.getElementById('arKpiGapCard');

  [arGoalCard,arExpectedCard,arRecoveryCard,arGapCard].forEach(card=>{
    if(!card)return;
    card.classList.remove('danger','success','warning');
  });

  if(arGoal>0){
    const paceRatio=arExpectedNow>0?r.currentOut/arExpectedNow:0;

    if(arProjectedGap>=0){
      arGoalCard?.classList.add('success');
      arGapCard?.classList.add('success');
      arRecoveryCard?.classList.add('success');
    }else{
      arGoalCard?.classList.add('danger');
      arGapCard?.classList.add('danger');
      arRecoveryCard?.classList.add('danger');
    }

    if(paceRatio>=1){
      arExpectedCard?.classList.add('success');
    }else if(paceRatio>=0.9){
      arExpectedCard?.classList.add('warning');
    }else{
      arExpectedCard?.classList.add('danger');
    }
  }
  setTextSafe('arKpiRunsCompleted',r.runsCompleted);
  setTextSafe('arKpiRunsProjected',round1(r.projectedRunsExact));
  setTextSafe('arKpiWip',formatARNumber(r.bosWip));
  setTextSafe(
    'arBosWipMeta',
    `${round1(r.bosRuns)} BOS runs · ${formatARNumber(r.liveInside)} live AR WIP now`
  );
  setTextSafe('arKpiSurfaceFeed',formatARNumber(r.surfaceFeed));
  setTextSafe('arKpiMaintenance',`${r.totalDelay} min`);
  setTextSafe(
    'arMaintenanceMeta',
    r.totalDelay
      ? `${hourly.filter(row=>row.maintenance).length} affected hour(s)`
      : 'No affected hours'
  );
  setTextSafe('arKpiFocus',r.focus.name);
  setTextSafe('arFocusMeta',r.focus.reason);

  setTextSafe('arSummaryRange',`${formatARNumber(r.low)} - ${formatARNumber(r.high)}`);
  setTextSafe('arSummaryProjected',formatARNumber(r.projectedOut));
  setTextSafe('arSummaryGoal',round1(r.bosRuns));
  setTextSafe(
    'arSummaryBosRunsMeta',
    `${formatARNumber(r.bosWip)} BOS jobs ÷ ${AR_FORECAST_STATE.config.runSize}`
  );
  setTextSafe(
    'arSummaryCapacity',
    formatARNumber(
      Math.min(
        r.totalWorkAvailable,
        r.currentOut+Math.max(0,r.laborCeiling)
      )
    )
  );

  setTextSafe(
    'arSummaryNarrative',
    `BOS AR WIP is ${formatARNumber(r.bosWip)} jobs (${round1(r.bosRuns)} runs). `+
    `Surface Inspection currently adds up to ${formatARNumber(r.surfaceFeed)} jobs (${round1(r.feedRuns)} runs) of incoming feed. `+
    `Projected AR-OUT is ${formatARNumber(r.projectedOut)} with ${formatARNumber(r.projectedEosWip)} jobs projected to remain.`
  );

  setTextSafe('arStructuralBottleneck',r.focus.name);
  setTextSafe('arStructuralMeta',r.focus.reason);
  setTextSafe('arLiveWipFocus',r.wipFocus[0]);
  setTextSafe('arLiveWipMeta',`${formatARNumber(r.wipFocus[1])} live WIP`);

  const next=getARNextRunCompletion(r.live,r.totalDelay);
  setTextSafe('arSummaryNextCompletion',next.label);
  setTextSafe('arSummaryNextCompletionMeta',next.meta);

  const gap=r.goal?r.projectedOut-r.goal:0;
  setTextSafe(
    'arSummaryGoalGap',
    r.goal
      ? `${Math.abs(gap)} jobs ${gap>=0?'above':'below'}`
      : `${formatARNumber(r.projectedEosWip)} jobs`
  );
  setTextSafe(
    'arSummaryGoalMeta',
    r.goal
      ? `Goal ${formatARNumber(r.goal)}`
      : 'Projected EOS AR WIP'
  );

  setTextSafe(
    'arProjectedOutMeta',
    `${round1(r.projectedRunsExact)} projected runs · ${Math.round(r.remainingMin)} min remain`
  );
  setTextSafe(
    'arRunsProjectedMeta',
    `${Math.max(0,round1(r.projectedRunsExact-r.currentOut/AR_FORECAST_STATE.config.runSize))} additional runs`
  );

  renderARStage('ArIn',r.live.arIn);
  renderARStage('Basket',r.live.basket);
  renderARStage('Oven',r.live.oven);
  renderARStage('Sectoring',r.live.sectoring);
  renderARStage('DeRing',r.live.deRing);
  renderARStage('ArOut',r.currentOut);

  highlightARStage(r.focus.name);
  renderARGoalPlan(r);
  renderARBottleneckTable(r);
  renderARRecommendations(r);
  renderARHourlyCards(hourly,r);
  renderARPace(r,pacePct);
  renderARStaffingCards();

  setTextSafe(
    'arPipelineStatus',
    `${getARTotalPipelineMinutes()} min normal flow · ${AR_FORECAST_STATE.config.runSize} jobs = 1 run`
  );

  if(AR_FORECAST_STATE.department==='AR'){
    setTextSafe(
      'lastUpdated',
      new Date().toLocaleString(
        'en-US',
        {month:'short',day:'numeric',hour:'numeric',minute:'2-digit'}
      )
    );
  }
}
function renderARGoalPlan(r){const body=document.getElementById('arGoalPlanTable'),summary=document.getElementById('arGoalPlanSummary');if(!body||!summary)return;const remainingGoal=Math.max(0,r.goal-r.currentOut),requiredJph=r.remainingHours>0?remainingGoal/r.remainingHours:0;const rows=r.stationCaps.filter(x=>x.station!=='Oven').map(x=>{const defaultJph=AR_DEFAULT_JPH[x.station]||1,requiredAssociates=requiredJph/defaultJph,gap=Math.max(0,requiredAssociates-x.currentAssociates),status=x.usableJph>=requiredJph?'OK':x.usableJph>=requiredJph*.9?'WATCH':'AT RISK';return{...x,requiredJph,requiredAssociates,gap,status};});body.innerHTML=rows.map(x=>`<tr><td><strong>${escapeHtml(x.station)}</strong></td><td>${round1(x.requiredJph)}</td><td>${round1(x.usableJph)}</td><td>${round1(x.requiredAssociates)}</td><td>${round1(x.currentAssociates)}</td><td>${x.gap>0?'+'+round1(x.gap):'—'}</td><td class="ar-status-pill ${x.status==='OK'?'ok':x.status==='WATCH'?'watch':'risk'}">${x.status}</td></tr>`).join('');const risks=rows.filter(x=>x.status!=='OK');summary.innerHTML=r.goal<=0?'<p class="muted-cell">Set an AR Goal above 0 to calculate required station JPH and associates.</p>':risks.length?`<p class="goal-plan-short"><strong>${risks.length}</strong> AR area(s) do not have comfortable Final JPH coverage for the remaining goal of <strong>${formatARNumber(remainingGoal)}</strong> jobs.</p>`:`<p class="goal-plan-ok">All labor-controlled AR stations have enough Final JPH to cover the remaining goal of ${formatARNumber(remainingGoal)} jobs.</p>`;}
function renderARBottleneckTable(r){const body=document.getElementById('arBottleneckTable');if(!body)return;const liveMap={'AR-IN':r.live.arIn,'Basket':r.live.basket,'Oven':r.live.oven,'Sectoring':r.live.sectoring,'DeRing':r.live.deRing,'AR-OUT':r.live.deRing};body.innerHTML=r.stationCaps.map(x=>{const w=liveMap[x.station]||0,clear=x.station==='Oven'?`${AR_FORECAST_STATE.config.oven} min process`:x.usableJph>0?`${Math.round(w/x.usableJph*60)} min`:'No JPH',signal=x.station==='Oven'?'WATCH':x.usableJph<=0?'AT RISK':w/x.usableJph*60>120?'AT RISK':w/x.usableJph*60>60?'WATCH':'OK';return`<tr><td><strong>${escapeHtml(x.station)}</strong></td><td>${round1(x.currentAssociates)}</td><td>${x.station==='Oven'?'Process':round1(x.usableJph)}</td><td>${x.station==='Oven'?'60 min cycle':formatARNumber(x.remainingCapacity)}</td><td>${formatARNumber(w)}</td><td>${completeARRuns(w)}</td><td>${clear}</td><td class="${signal==='OK'?'ar-signal-ok':signal==='WATCH'?'ar-signal-watch':'ar-signal-risk'}"><strong>${signal}</strong></td></tr>`;}).join('');}
function renderARStaffingCards(){
  const box=document.getElementById('arStaffingCards');
  if(!box)return;

  const roster=getARStationRoster();

  box.innerHTML=AR_STATION_ORDER.map(station=>{
    const people=roster[station]||[];
    const detectedCount=people.length;
    const capacityCount=station==='Oven'
      ? detectedCount
      : getARCapacityAssociateCount(station,detectedCount);
    const stationJph=station==='Oven'?0:(AR_DEFAULT_JPH[station]||0);
    const liveCapacity=station==='Oven'?0:capacityCount*stationJph;
    const isManual=
      station!=='Oven'&&
      Number.isFinite(Number(AR_FORECAST_STATE.staffingCapacityOverrides?.[station]));

    const rows=people.length
      ? people.map(p=>{
          const movement=p.otherStations?.length
            ? ` · Also ${p.otherStations.map(x=>`${x.station} ${formatARNumber(x.output)}`).join(', ')}`
            : '';

          return `<div class="ar-associate-row ${p.otherStations?.length?'floater':''}">
            <div>
              <strong>${escapeHtml(p.name)}</strong>
              <small>${escapeHtml(p.status)} · ${escapeHtml(p.lastHourLabel)} · ${formatARNumber(p.lastHourOutput)} jobs${escapeHtml(movement)}</small>
            </div>
            <div class="ar-associate-actions">
              <em>${station==='Oven'?'Process':round1(p.effectiveJph)+' JPH'}</em>
              <button type="button" class="ar-hide-btn"
                data-ar-staff-action="hide"
                data-operator="${escapeHtml(p.name)}"
                title="Hide from staffing capacity only">Hide</button>
            </div>
          </div>`;
        }).join('')
      : `<div class="ar-associate-row">
          <div>
            <strong>No current associate</strong>
            <small>No positive hourly activity at this station yet</small>
          </div>
          <em>${station==='Oven'?'Process':'0 JPH'}</em>
        </div>`;

    const countControl=station==='Oven'
      ? `<div>
          <small>Live Associates</small>
          <b>${detectedCount}</b>
        </div>`
      : `<div class="ar-capacity-count-box ${isManual?'manual':''}">
          <small>${isManual?'Manual Capacity Count':'Live Associates'}</small>
          <div class="ar-capacity-input-row">
            <input
              class="ar-capacity-associate-input"
              type="number"
              min="0"
              step="0.5"
              value="${round1(capacityCount)}"
              data-ar-capacity-station="${escapeHtml(station)}"
              title="Adjust staffing capacity in 0.5 associate increments">
            <button
              type="button"
              class="ar-capacity-auto-btn"
              data-ar-staff-action="reset-capacity"
              data-station="${escapeHtml(station)}"
              ${isManual?'':'disabled'}
              title="Return to detected live associate count">Auto</button>
          </div>
          <small class="ar-detected-count">Detected: ${detectedCount}</small>
        </div>`;

    return `<article class="ar-staff-card">
      <div class="ar-staff-card-head">
        <strong>${escapeHtml(station)}</strong>
        <span>${station==='Oven'?'PROCESS':round1(liveCapacity)+' JPH'}</span>
      </div>

      <div class="ar-staff-summary">
        ${countControl}
        <div>
          <small>Station JPH</small>
          <b>${station==='Oven'?'N/A':stationJph}</b>
        </div>
        <div>
          <small>Live Capacity</small>
          <b>${station==='Oven'?'60 min':round1(liveCapacity)}</b>
        </div>
      </div>

      <div class="ar-associate-list">${rows}</div>
    </article>`;
  }).join('');
}
function renderARRecommendations(r){const box=document.getElementById('arOperationalRecommendations');if(!box)return;const items=[];items.push(`<div class="recommendation-item"><strong>Primary focus: ${escapeHtml(r.focus.name)}</strong><span>${escapeHtml(r.focus.reason)}</span></div>`);if(r.live.surface>r.live.arIn&&r.live.surface>=AR_FORECAST_STATE.config.runSize)items.push(`<div class="recommendation-item"><strong>Protect AR-IN feed</strong><span>${formatARNumber(r.live.surface)} jobs are waiting from Surface while AR-IN has ${formatARNumber(r.live.arIn)} WIP.</span></div>`);if(r.totalDelay)items.push(`<div class="recommendation-item"><strong>Maintenance impact</strong><span>${r.totalDelay} total delay minutes are applied to affected pipeline work.</span></div>`);if(r.goal)items.push(`<div class="recommendation-item"><strong>${r.projectedOut>=r.goal?'Goal currently covered':'Recovery required'}</strong><span>Projected AR-OUT ${formatARNumber(r.projectedOut)} versus goal ${formatARNumber(r.goal)}.</span></div>`);box.innerHTML=items.join('');}
function renderARHourlyCards(hourly,r){const box=document.getElementById('arHourlyCards');if(!box)return;const goal=r.goal,productive=SHIFT_PRODUCTIVE_HOURS[getValue('shiftMode','Weekday')]||9.5,expectedPerHour=goal>0?goal/productive:0;box.innerHTML=hourly.map(x=>{const pct=expectedPerHour>0?Math.round(x.arOut/expectedPerHour*100):0,cls=x.maintenance?'risk':pct>=100?'on-track':pct>=90?'at-risk':'behind';return`<div class="sector-row ${cls}"><span>${escapeHtml(formatARHourLabel(x.hour))}</span><strong>OUT ${formatARNumber(x.arOut)} / ${expectedPerHour?formatARNumber(expectedPerHour):'—'}</strong><em>${completeARRuns(x.arOut)} Runs${expectedPerHour?` · ${pct}%`:''}${x.maintenance?` · +${x.delay}m`:''}</em></div>`;}).join('')||'<div class="sector-row not-started">No AR hourly data yet.</div>';}
function renderARPace(r,ratio){const pct=Math.max(0,Math.round(ratio*100)),gap=Math.round(r.currentOut-r.expectedNow);setTextSafe('arPacePercent',`${pct}%`);setTextSafe('arPaceGap',`${gap>=0?'+':''}${formatARNumber(gap)} jobs`);setTextSafe('arPaceStatus',ratio>=1?'ON TRACK':ratio>=.9?'AT RISK':'BEHIND');const actual=document.getElementById('arActualPaceBar'),expected=document.getElementById('arExpectedPaceBar'),gauge=document.getElementById('arPaceGauge');if(actual)actual.style.width=`${Math.min(100,pct)}%`;if(expected)expected.style.width='100%';if(gauge)gauge.style.background=`radial-gradient(circle at center,#06111f 0 58%,transparent 59%),conic-gradient(var(--${ratio>=1?'green':ratio>=.9?'yellow':'red'}) 0 ${Math.min(100,pct)}%,rgba(255,255,255,.12) ${Math.min(100,pct)}% 100%)`;}
function highlightARStage(name){document.querySelectorAll('.ar-stage').forEach(x=>x.classList.remove('pressure'));const map={'AR-IN':'arin','Basket':'basket','Sectoring':'sectoring','DeRing':'dering','AR-OUT':'arout'};const key=map[name];if(key)document.querySelector(`[data-ar-stage="${key}"]`)?.classList.add('pressure');}
function getARNextRunCompletion(live,totalDelay){const c=AR_FORECAST_STATE.config,stages=[['DeRing',live.deRing,c.arOut],['Sectoring',live.sectoring,c.deRing+c.arOut],['Oven',live.oven,c.sectoring+c.chamber+c.deRing+c.arOut],['Basket',live.basket,c.oven+c.sectoring+c.chamber+c.deRing+c.arOut],['AR-IN',live.arIn,c.basket+c.t40+c.oven+c.sectoring+c.chamber+c.deRing+c.arOut]],hit=stages.find(x=>completeARRuns(x[1])>0);if(!hit)return{label:'--',meta:'No complete run available'};const d=new Date(Date.now()+(hit[2]+totalDelay)*60000);return{label:d.toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit'}),meta:`From ${hit[0]} · ${hit[2]+totalDelay} min`};}
function renderARStage(name,value){setTextSafe(`arStage${name}`,formatARNumber(value));setTextSafe(`arStage${name}Runs`,`${completeARRuns(value)} Runs`);}
function collectARHourlyData(){
  return Array.from(document.querySelectorAll('.ar-hour-row')).map(row=>({
    hour:row.dataset.hour,
    arIn:toARNumber(row.querySelector('.ar-hour-in')?.value),
    arOut:toARNumber(row.querySelector('.ar-hour-out')?.value),
    maintenance:Boolean(row.querySelector('.ar-hour-maint')?.checked),
    delay:toARNumber(row.querySelector('.ar-delay-input')?.value),
    note:String(row.querySelector('.ar-hour-note')?.value||'')
  }));
}
function readARConfigFromInputs(){const map={goal:'arGoal',runSize:'arRunSize',maintenanceDelay:'arMaintenanceDelay',arIn:'arTimeArIn',basket:'arTimeBasket',t40:'arTimeT40',oven:'arTimeOven',sectoring:'arTimeSectoring',chamber:'arTimeChamber',deRing:'arTimeDeRing',arOut:'arTimeArOut'};Object.entries(map).forEach(([k,id])=>{const el=document.getElementById(id);if(!el)return;AR_FORECAST_STATE.config[k]=toARNumber(el.value);});}
function writeARConfigInputs(){const map={goal:'arGoal',runSize:'arRunSize',maintenanceDelay:'arMaintenanceDelay',arIn:'arTimeArIn',basket:'arTimeBasket',t40:'arTimeT40',oven:'arTimeOven',sectoring:'arTimeSectoring',chamber:'arTimeChamber',deRing:'arTimeDeRing',arOut:'arTimeArOut'};Object.entries(map).forEach(([k,id])=>{const el=document.getElementById(id);if(el)el.value=AR_FORECAST_STATE.config[k];});}
function getARTotalPipelineMinutes(){const c=AR_FORECAST_STATE.config;return c.arIn+c.basket+c.t40+c.oven+c.sectoring+c.chamber+c.deRing+c.arOut;}
function completeARRuns(jobs){return Math.floor(toARNumber(jobs)/Math.max(1,AR_FORECAST_STATE.config.runSize));}
function resetARSetup(){AR_FORECAST_STATE.config={...AR_DEFAULT_CONFIG};writeARConfigInputs();saveARLocalState();calculateARForecast();}
async function saveARSetup(){readARConfigFromInputs();saveARLocalState();await postARForecastApi('saveConfig',{config:AR_FORECAST_STATE.config});showARMessage('AR setup saved.');}
async function saveARHourly(){const payload={date:getValue('productionDate',''),shiftMode:getValue('shiftMode','Weekday'),config:AR_FORECAST_STATE.config,hourly:collectARHourlyData()};saveARLocalState();await postARForecastApi('saveHourly',payload);showARMessage('AR hourly saved.');}
async function saveARForecast(){if(!AR_FORECAST_STATE.calculated)return;const payload={date:getValue('productionDate',''),shiftMode:getValue('shiftMode','Weekday'),config:AR_FORECAST_STATE.config,result:AR_FORECAST_STATE.calculated,liveStaff:AR_FORECAST_STATE.liveStaff};await postARForecastApi('saveForecast',payload);showARMessage('AR forecast saved.');}
function clearARManualHourly(){AR_FORECAST_STATE.manualHourly={};saveARLocalState();buildARHourlyRows();calculateARForecast();}
async function postARForecastApi(action,payload){try{const r=await fetch(AR_FORECAST_API_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action,payload})});return await r.json();}catch(e){console.warn('[AR Forecast API]',e);return{ok:false,error:e.message};}}
function saveARLocalState(){
  localStorage.setItem(
    AR_FORECAST_STORAGE_KEY,
    JSON.stringify({
      config:AR_FORECAST_STATE.config,
      manualHourly:AR_FORECAST_STATE.manualHourly,
      staffingCapacityOverrides:AR_FORECAST_STATE.staffingCapacityOverrides
    })
  );
  saveARStaffingCapacityOverrides();
}
function restoreARLocalState(){
  try{
    const x=JSON.parse(localStorage.getItem(AR_FORECAST_STORAGE_KEY)||'{}');
    AR_FORECAST_STATE.config={...AR_DEFAULT_CONFIG,...(x.config||{})};
    AR_FORECAST_STATE.manualHourly=x.manualHourly||{};

    if(x.staffingCapacityOverrides&&typeof x.staffingCapacityOverrides==='object'){
      AR_FORECAST_STATE.staffingCapacityOverrides={
        ...AR_FORECAST_STATE.staffingCapacityOverrides,
        ...x.staffingCapacityOverrides
      };
    }
  }catch{}
  writeARConfigInputs();
}
function normalizeARHourKey(key){const s=String(key||'').trim().replace(/^H/i,''),m=s.match(/(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?/i);if(!m)return s;let h=Number(m[1]),min=Number(m[2]||0);if(m[3]){const ap=m[3].toUpperCase();if(ap==='PM'&&h<12)h+=12;if(ap==='AM'&&h===12)h=0;}return`${String(h).padStart(2,'0')}:${String(min).padStart(2,'0')}`;}
function formatARHourLabel(key){const[h,m]=normalizeARHourKey(key).split(':').map(Number),d=new Date();d.setHours(h,m||0,0,0);return d.toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit'});}
function toARNumber(v){const n=Number(String(v??'').replace(/,/g,'').trim());return Number.isFinite(n)?n:0;}
function formatARNumber(v){return Math.round(toARNumber(v)).toLocaleString('en-US');}
function setTextSafe(id,value){const el=document.getElementById(id);if(el)el.textContent=String(value);}
function showARMessage(text){if(typeof showToast==='function')showToast(text);else console.log(text);}

