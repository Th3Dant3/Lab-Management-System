/* Surface Forecast Command Center
 * File: surfaceforecast.js
 * Version: 30 - Futuristic Pace Tab + Minute Clear-Time Restore
 *
 * New V7:
 * - Forecast is driven by BOS WIP plus hidden average incoming/remake planning ranges, usable station capacity, machines running, associates assigned, and downtime.
 * - Incoming/remake assumptions stay in the background and are not displayed as setup inputs.
 * - Projected OUT no longer falls to 0 when hourly data is missing.
 * - Projected range is capped by available work and today's bottleneck capacity.
 * - Live WIP bottleneck now uses last-scan -> next-process pressure routing.
 * - Area Bottleneck Detail table is organized in Surface process flow order from SF Unbox to AR41 inspection.
 * - Engraving -> Manual Deblock / Detaping uses Engraving capacity because Manual Deblock / Detaping has no reliable LMS scan/JPH feed.
 * - Pace vs Expected and Hourly Outlook are now supported as their own HTML tab.
 * - Area Bottleneck Detail table stays in Surface process flow order.
 * - Time To Clear uses operational minute ranges instead of decimal hours.
 * - Futuristic Pace / Hourly tab enhancements remain intact.
 */

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

const SURFACE_AREAS = [
  // capacityMode:
  // operator = capacity is driven by associates assigned.
  // machine = capacity is driven by machines running, but requires at least 1 associate assigned to run/clear the area.
  // buffer = dwell/buffer area; shown for WIP pressure, not used as the main route bottleneck.
  //
  // machinesPerAssociate (machine areas only): these machines are fully automatic — the associate's
  // job is starting/feeding them, clearing errors/jams, and keeping WIP flowing to the next station.
  // This ratio is NOT a hard "cannot run more than N machines" cap. It's the number of machines one
  // associate can realistically keep clear of errors before backlog builds and machines start sitting
  // idle waiting on intervention. Going over it is a DOWNTIME RISK, not an immediate capacity loss —
  // treat any gap the Goal Plan shows as "watch this area," not "this machine physically cannot run."
  // Inferred starting point from your existing defaultUnits/defaultAssociates ratio — confirm/correct
  // per area in the Setup tab (editable).
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
    pressureLabel: 'Auto Blocker / Detaper',
    timeMode: 'sf_unbox_stage',
    note: 'SF Unbox scan is feeding Auto Blocker / Detaper'
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
  '06:00', '07:00', '08:00', '09:00', '10:00',
  '11:00', '12:00', '13:00', '14:00', '15:00',
  '16:00', '17:00', '18:00'
];

let downtimeCounter = 0;
let latestForecastResult = null;

const SURFACE_STORAGE_KEY = 'surfaceForecastCommandCenter.v7';
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

  const restoredFromCloud = await loadCurrentStateFromApi();
  const restored = restoredFromCloud || restorePageState();

  if (!restored) {
    addDowntimeRow();
  }

  calculateForecast();
  savePageState();

  await syncSurfaceDashboardData(true);
  window.setInterval(function () {
    syncSurfaceDashboardData(false);
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

  panels.forEach(function (panel) {
    panel.hidden = panel.dataset.tabPanel !== target;
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
    syncDashboardBtn: function () { syncSurfaceDashboardData(true); }
  };

  Object.keys(buttons).forEach(function (id) {
    const el = document.getElementById(id);
    if (el) el.addEventListener('click', buttons[id]);
  });

  ['shiftMode', 'perfShiftModeSelect', 'perfDateInput', 'bosWip', 'shipGoal', 'incomingLow', 'incomingHigh', 'remakeLow', 'remakeHigh', 'floaterCount', 'floaterAssign'].forEach(function (id) {
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
  const newText = 'Forecast range uses BOS WIP plus background incoming/remake planning averages, capped by usable bottleneck capacity. Hourly input tells you if actual AR41 output is on track.';

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

function buildHourlyRows() {
  const body = document.getElementById('hourlyRows');
  if (!body) return;

  body.innerHTML = '';

  HOURLY_DEFAULTS.forEach(function (hour) {
    const tr = document.createElement('tr');
    tr.className = 'hourly-row';

    tr.innerHTML = `
      <td><input class="hour-input hr-hour" type="text" value="${hour}"></td>
      <td><input class="hr-unbox" type="number" min="0" step="1"></td>
      <td><input class="hr-autoblocker" type="number" min="0" step="1"></td>
      <td><input class="hr-cooling" type="number" min="0" step="1"></td>
      <td><input class="hr-orb" type="number" min="0" step="1"></td>
      <td><input class="hr-polisher" type="number" min="0" step="1"></td>
      <td><input class="hr-engraving" type="number" min="0" step="1"></td>
      <td><input class="hr-detaping" type="number" min="0" step="1"></td>
      <td><input class="hr-coater" type="number" min="0" step="1"></td>
      <td><input class="hr-inspection" type="number" min="0" step="1"></td>
      <td><input class="notes-input hr-notes" type="text"></td>
      <td>
        <select class="hr-active">
          <option value="Yes">Yes</option>
          <option value="No">No</option>
        </select>
      </td>
    `;

    body.appendChild(tr);

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
        <option value="Machine Slow">Machine Slow</option>
        <option value="Quality Issue">Quality Issue</option>
        <option value="Material Issue">Material Issue</option>
        <option value="Operator Short">Operator Short</option>
        <option value="Other">Other</option>
      </select>
    </label>

    <label>
      Units Down
      <input class="dt-units" type="number" min="0" step="0.5" value="0">
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
      hour: getRowValue(row, '.hr-hour'),
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

function calculateLocalForecast(payload) {
  const shiftMode = payload.shiftMode || 'Weekday';
  const shiftHours = getShiftHours(shiftMode);
  const bosWip = Number(payload.bosWip) || 0;
  const shipGoal = Number(payload.shipGoal) || 0;
  const incomingRange = normalizeForecastRange(payload.incomingLow, payload.incomingHigh);
  const remakeRange = normalizeForecastRange(payload.remakeLow, payload.remakeHigh);
  const availableWorkLow = bosWip + incomingRange.low + remakeRange.low;
  const availableWorkHigh = bosWip + incomingRange.high + remakeRange.high;
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
      totalJph: capacityInfo.totalJph,
      normalCapacity,
      downtimeLostJobs,
      adjustedCapacity,
      currentWip,
      effectiveJph,
      timeToClear
    };
  });

  const requiredAreas = areas.filter(function (area) {
    return area.required && area.effectiveJph > 0;
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

  const hourlySummary = calculateHourlySummary(payload.hourly || [], shiftHours, effectiveBottleneckJph, shipGoal, capacityCeiling, availableWorkLow, availableWorkHigh);
  const liveWipBottleneck = getLiveWipBottleneck(areas);

  const currentOut = Number(hourlySummary.actualAr41SoFar || 0);
  const projectedOut = Number(hourlySummary.projectedActualOut || 0);
  const capacityProjectedOut = Number(hourlySummary.capacityProjectedOutHigh || hourlySummary.capacityProjectedOut || 0);
  const bosRemainingAfterActual = Math.max(0, bosWip - currentOut);
  const availableWorkRemainingLow = Math.max(0, availableWorkLow - currentOut);
  const availableWorkRemainingHigh = Math.max(0, availableWorkHigh - currentOut);

  // Low = current pace projection.
  // High = best case from BOS + average incoming + average remake/breakage, capped by bottleneck capacity.
  const knownWipForecast = Number(hourlySummary.capacityProjectedOutLow || 0);
  const forecastRangeLow = projectedOut;
  const forecastRangeHigh = capacityProjectedOut;
  const extraCapacityForUnknownWork = Math.max(0, capacityCeiling - availableWorkHigh);
  const expectedEosWipIfNoExtraWork = Math.max(0, availableWorkLow - knownWipForecast);

  let requiredJph = 0;
  let goalRisk = 'NO GOAL';
  let goalGap = 0;

  if (shipGoal > 0) {
    requiredJph = shipGoal / shiftHours;
    goalGap = shipGoal - capacityProjectedOut;

    if (shipGoal <= knownWipForecast) goalRisk = 'GREEN';
    else if (shipGoal <= capacityProjectedOut) goalRisk = 'YELLOW';
    else goalRisk = 'RED';
  }

  const recommendation = buildRecommendation({
    shiftMode,
    shiftHours,
    bosWip,
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
  updatePerformanceUI(result);
  renderGoalStaffingPlan(result);

  setText('forecastStatus', status || 'READY');
  setText('forecastRange', `${formatNumber(Math.max(actual, projected))} - ${formatNumber(Math.max(Math.max(actual, projected), Number(result.forecastRangeHigh || result.capacityCeiling || projected || 0)))}`);
  setText('shipGoalDisplay', shipGoal > 0 ? formatNumber(shipGoal) : '0');
  setText('projectedGap', buildProjectedGapText(shipGoal, projected));
  setText('pacePercent', expected > 0 ? `${round1(pacePercent)}%` : '0%');
  setText('paceGap', `${paceGap >= 0 ? '+' : ''}${formatNumber(paceGap)} jobs`);
  setText('lastUpdated', nowText);

  setText('summaryRangeValue', `${formatNumber(Math.max(actual, projected))} - ${formatNumber(Math.max(Math.max(actual, projected), Number(result.forecastRangeHigh || result.capacityCeiling || projected || 0)))}`);
  setText('summaryProjectedOutValue', formatNumber(projected));
  setText('summaryTargetValue', shipGoal > 0 ? formatNumber(shipGoal) : '0');
  setText('summaryDowntimeValue', formatNumber(result.downtimeLostTotal));
  setText('summaryFootnote', buildSummaryFootnote(result, actionPlan));

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
  const capacityHigh = Number(result.forecastRangeHigh ?? result.capacityProjectedOut ?? result.capacityCeiling ?? projected ?? 0);
  const remainingHours = Math.max(0, Number(hourly.remainingHours || 0));
  const downtimeLost = Number(result.downtimeLostTotal || 0);

  /*
    Do not use the old BOS/capacity "forecast range" for the KPI.
    That was causing 1,425 - 1,425 and giving no real decision value.
    KPI range now shows: current pace projection low -> best available capacity high.
  */
  const projectedLow = Math.max(actual, projected);
  const projectedHigh = Math.max(projectedLow, capacityHigh);

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
    const label = formatHourLabel(slots[i - 1] ? slots[i - 1].label : '');
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

  setValue('shiftMode', payload.shiftMode || 'Weekday');
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

    setValue('shiftMode', payload.shiftMode || 'Weekday');
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

    restoreHourlyRows(payload.hourly || []);
    restoreDowntimeRows(payload.downtime || []);

    return true;

  } finally {
    isRestoringState = false;
  }
}

function restoreHourlyRows(hourlyRows) {
  const rows = Array.from(document.querySelectorAll('.hourly-row'));

  rows.forEach(function (row, index) {
    const saved = hourlyRows[index];
    if (!saved) return;

    setRowValue(row, '.hr-hour', saved.hour || getRowValue(row, '.hr-hour'));
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
    const hour = getRowValue(tr, '.hr-hour');
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

  // Machine areas are driven by machines running, but an area with 0 associates should not produce.
  // One associate can run/clear multiple machines, so do NOT cap machines with associate count.
  if (mode === 'machine') {
    const activeMachines = associates > 0 ? machines : 0;
    return {
      activeCapacityUnits: activeMachines,
      capacityBasis: associates > 0 ? 'machines + associate coverage' : 'no associate assigned',
      totalJph: activeMachines * jph,
      normalCapacity: activeMachines * jph * shiftHours
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
  const dateInput = document.getElementById('perfDateInput');
  const perfShift = document.getElementById('perfShiftModeSelect');
  const mainShift = document.getElementById('shiftMode');

  if (dateInput && !dateInput.value) {
    dateInput.value = getLocalDateInputValue(new Date());
  }

  if (perfShift && mainShift) {
    perfShift.value = mainShift.value || 'Weekday';
  }

  if (perfShift && mainShift) {
    perfShift.addEventListener('change', function () {
      mainShift.value = perfShift.value;
      calculateForecast();
    });

    mainShift.addEventListener('change', function () {
      perfShift.value = mainShift.value;
      updatePerformanceShiftChrome(latestForecastResult ? latestForecastResult.result : null);
    });
  }

  if (dateInput) {
    dateInput.addEventListener('change', function () {
      updatePerformanceShiftChrome(latestForecastResult ? latestForecastResult.result : null);
    });
  }

  updatePerformanceShiftChrome(null);
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


function getHourlyRowOffsetForShift(shiftMode) {
  const windowInfo = getShiftWindow(shiftMode || 'Weekday');
  const startMinutes = parseTimeToMinutes(windowInfo.start);
  if (startMinutes === null) return 0;
  const floorHour = Math.floor(startMinutes / 60) * 60;
  const key = minutesToHourKey(floorHour);
  const idx = HOURLY_DEFAULTS.indexOf(key);
  return idx >= 0 ? idx : 0;
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
  const nowMinutes = now.getHours() * 60 + now.getMinutes();
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
  setText('perfDayLabel', getPerformanceDateLabel(dateInput ? dateInput.value : ''));

  const progress = document.getElementById('perfShiftProgress');
  const remain = document.getElementById('perfShiftRemaining');
  const marker = document.getElementById('perfNowMarker');

  if (progress) progress.style.width = `${pct}%`;
  if (remain) remain.style.width = `${Math.max(0, 100 - pct)}%`;
  if (marker) marker.style.left = `${pct}%`;
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
  const nowMinutes = now.getHours() * 60 + now.getMinutes();
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
    const nextHour = Math.min(endMinutes, start + 60);
    const label = nextHour - start < 60
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