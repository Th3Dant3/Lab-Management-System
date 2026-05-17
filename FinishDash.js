const FINISH_API_URL =
  "https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec?area=Finish";

let lastSyncTime = null;

const DEFAULT_CONFIG = {
  unboxRate: 150,
  unboxCount: 1,
  meiRate: 50,
  meiMachines: 5,
  mountRate: 25,
  mountCount: 10,
  drillRate: 5,
  drillCount: 13,
  finalRate: 75,
  finalCount: 1
};

const SHIFT_RULES = {
  weekday: {
    shiftStart: "7:00 AM",
    shiftEnd: "5:30 PM",
    breaks: [
      { name: "Morning Break", start: "9:30 AM", end: "9:45 AM" },
      { name: "Lunch Group 1", start: "11:30 AM", end: "12:00 PM" },
      { name: "Lunch Group 2", start: "12:00 PM", end: "12:30 PM" },
      { name: "Afternoon Break", start: "3:00 PM", end: "3:15 PM" }
    ]
  },

  weekend: {
    shiftStart: "6:30 AM",
    shiftEnd: "6:30 PM",
    breaks: [
      { name: "Morning Break", start: "9:30 AM", end: "9:50 AM" },
      { name: "Lunch Group 1", start: "11:30 AM", end: "12:00 PM" },
      { name: "Lunch Group 2", start: "12:00 PM", end: "12:30 PM" },
      { name: "Afternoon Break", start: "3:30 PM", end: "3:50 PM" }
    ]
  }
};

document.addEventListener("DOMContentLoaded", () => {
  initTabs();
  initClock();
  initConfigPanel();
  initRefreshButton();
  initHoverEffects();

  loadDashboard();

  setInterval(loadDashboard, 5 * 60 * 1000);
  setInterval(updateSyncAge, 1000);
});

/* ===============================
   MAIN LOAD
================================ */

async function loadDashboard() {
  showLoading(true);

  try {
    const response = await fetch(FINISH_API_URL, {
      method: "GET",
      cache: "no-store"
    });

    if (!response.ok) {
      throw new Error(`API failed with status ${response.status}`);
    }

    const data = await response.json();

    if (data.finishDashboard?.allStations) {
      renderEnrichedDashboard(data.finishDashboard);
    } else {
      renderLegacyDashboard(data);
    }

    lastSyncTime = Date.now();
    setText("syncAge", "just now");
  } catch (error) {
    console.error("Finish Dashboard Error:", error);
    renderError(error);
  } finally {
    showLoading(false);
  }
}

/* ===============================
   ENRICHED DASHBOARD
================================ */

function renderEnrichedDashboard(finishDashboard) {
  const config = loadConfig();

  const floorStations = (finishDashboard.allStations || []).map(station =>
    applyConfigToStation({ ...station }, config)
  );

  const healthStations = (finishDashboard.stationHealth || []).map(station =>
    applyConfigToStation({ ...station }, config)
  );

  const summary =
    finishDashboard.summary ||
    buildSummaryFromStations(floorStations);

  renderKpis(summary);
  renderStationNumbers(floorStations);
  renderFlowCore(summary, floorStations);

  renderStatusPanel(healthStations);
  renderStationDetail(healthStations);
  renderHourlyBreakdown(healthStations);

  applyStationStateClasses(floorStations);

  setText("summaryUpdated", formatTime(new Date()));
}

function renderKpis(summary) {
  setText("kpiIncomingWip", summary.incomingWip || 0);
  setText("kpiMeiFeedWip", summary.meiFeedWip || 0);
  setText("kpiMeiBankWip", summary.edgingWip || 0);
  setText("kpiMountDrillWip", summary.mountingDrillWip || 0);
  setText("kpiFinalWip", summary.finalInspectionWip || 0);

  setText("summaryTotalWip", summary.totalWip || 0);
  setText("summaryActivity", summary.totalActivityToday || 0);
  setText("summaryLargestValue", summary.largestWipTotal || 0);
  setText("summaryLargestName", summary.largestWipStation || "--");
  setText("summarySideWip", summary.sideStationWip || 0);
}

function renderStationNumbers(stations) {
  const get = name => stations.find(s => s.flowStep === name || s.displayName === name) || {};

  const fsv = get("FSV Scan & Verify");
  const unbox = get("Finish Unbox");
  const arout = get("AR-OUT");

  const meiB = get("MEI Line B");
  const meiC = get("MEI Line C");
  const easy = get("MEI Easy Fit");

  const mounting = get("Mounting");
  const drill = get("Drill");
  const bigs = get("Bigs");
  const sharps = get("Sharps");
  const final = get("Final Inspection");

  setText("fsvWip", fsv.currentWip || 0);

  setText("unboxWip", unbox.currentWip || 0);
  setText("unboxCnt", unbox.activityToday || 0);

  setText("aroutWip", arout.currentWip || 0);

  setText("meiBWip", meiB.currentWip || 0);
  setText("meiBCnt", meiB.activityToday || 0);

  setText("easyWip", easy.currentWip || 0);
  setText("easyCnt", easy.activityToday || 0);

  setText("meiCWip", meiC.currentWip || 0);
  setText("meiCCnt", meiC.activityToday || 0);

  setText("mountingWip", mounting.currentWip || 0);
  setText("mountingCnt", mounting.activityToday || 0);
  setText("mountingWipCenter", mounting.currentWip || 0);
  setText("mountingCntCenter", mounting.activityToday || 0);

  setText("drillWip", drill.currentWip || 0);
  setText("drillCnt", drill.activityToday || 0);
  setText("drillWipCenter", drill.currentWip || 0);
  setText("drillCntCenter", drill.activityToday || 0);

  setText("bigsWip", bigs.currentWip || 0);
  setText("bigsCnt", bigs.activityToday || 0);

  setText("sharpsWip", sharps.currentWip || 0);
  setText("sharpsCnt", sharps.activityToday || 0);

  setText("finalWip", final.currentWip || 0);
  setText("finalCnt", final.activityToday || 0);
}

function renderFlowCore(summary, stations) {
  const totalWip = Number(summary.totalWip || 0);

  const get = name => (stations || []).find(s => s.flowStep === name || s.displayName === name) || {};
  const finishWip =
    Number(get("Finish Unbox").currentWip || 0) +
    Number(get("AR-OUT").currentWip || 0) +
    Number(get("MEI Line B").currentWip || 0) +
    Number(get("MEI Line C").currentWip || 0) +
    Number(get("MEI Easy Fit").currentWip || 0) +
    Number(get("Mounting").currentWip || 0) +
    Number(get("Drill").currentWip || 0) +
    Number(get("Bigs").currentWip || 0) +
    Number(get("Sharps").currentWip || 0) +
    Number(get("Final Inspection").currentWip || 0);

  setText("totalWipValue", totalWip);
  setText("finishWipValue", finishWip);
}

/* ===============================
   LEGACY FALLBACK
================================ */

function renderLegacyDashboard(data) {
  const rows =
    data.productionFlow ||
    data.finishFlow ||
    data.areaSummary ||
    data.stations ||
    [];

  const stations = normalizeLegacyStations(rows);
  const summary = buildSummaryFromStations(stations);

  renderKpis(summary);
  renderStationNumbers(stations);
  renderFlowCore(summary, stations);
  renderStatusPanel(stations);
  renderStationDetail(stations);
  renderHourlyBreakdown(stations);
  applyStationStateClasses(stations);
}

function normalizeLegacyStations(rows) {
  if (!Array.isArray(rows)) return [];

  return rows.map(row => {
    const name =
      row.flowStep ||
      row.displayName ||
      row.station ||
      row.Station ||
      row.name ||
      "";

    const currentWip =
      row.currentWip ??
      row.wip ??
      row.WIP ??
      row.CurrentWip ??
      0;

    const activityToday =
      row.activityToday ??
      row.countToday ??
      row.cnt ??
      row.CNT ??
      row.totalToday ??
      0;

    const lastHourActivity =
      row.lastHourActivity ??
      row.lastHour ??
      row.LastHour ??
      0;

    const expectedNormalPerHour =
      row.expectedNormalPerHour ??
      row.capacityPerHour ??
      row.ratePerHour ??
      0;

    return {
      flowStep: normalizeStationName(name),
      displayName: normalizeStationName(name),
      currentWip: Number(currentWip) || 0,
      activityToday: Number(activityToday) || 0,
      lastHourActivity: Number(lastHourActivity) || 0,
      expectedNormalPerHour: Number(expectedNormalPerHour) || 0,
      metricMode: row.metricMode || "STANDARD",
      hourly: row.hourly || row.Hours || {}
    };
  }).map(station => {
    station.utilizationPct = calculateUtilization(station);
    station.status = calculateStatus(station);
    station.statusLabel = getStatusLabel(station.status);
    station.statusMessage = getStatusMessage(station);
    return station;
  });
}

function normalizeStationName(name) {
  const value = String(name || "").trim().toLowerCase();

  if (value.includes("fsv")) return "FSV Scan & Verify";
  if (value.includes("frame only")) return "FSV Scan & Verify";
  if (value.includes("unbox")) return "Finish Unbox";
  if (value.includes("ar-out") || value.includes("ar out")) return "AR-OUT";
  if (value.includes("line b")) return "MEI Line B";
  if (value.includes("line c")) return "MEI Line C";
  if (value.includes("easy")) return "MEI Easy Fit";
  if (value.includes("mount")) return "Mounting";
  if (value.includes("drill")) return "Drill";
  if (value.includes("big")) return "Bigs";
  if (value.includes("sharp")) return "Sharps";
  if (value.includes("final")) return "Final Inspection";

  return name || "--";
}

function buildSummaryFromStations(stations) {
  const get = name => stations.find(s => s.flowStep === name) || {};

  const incomingWip =
    Number(get("FSV Scan & Verify").currentWip || 0);

  const meiFeedWip =
    Number(get("Finish Unbox").currentWip || 0) +
    Number(get("AR-OUT").currentWip || 0);

  const edgingWip =
    Number(get("MEI Line B").currentWip || 0) +
    Number(get("MEI Line C").currentWip || 0) +
    Number(get("MEI Easy Fit").currentWip || 0);

  const mountingDrillWip =
    Number(get("Mounting").currentWip || 0) +
    Number(get("Drill").currentWip || 0);

  const finalInspectionWip =
    Number(get("Final Inspection").currentWip || 0);

  const sideStationWip =
    Number(get("Bigs").currentWip || 0) +
    Number(get("Sharps").currentWip || 0);

  const totalWip = stations.reduce((sum, s) => sum + Number(s.currentWip || 0), 0);
  const totalActivityToday = stations.reduce((sum, s) => sum + Number(s.activityToday || 0), 0);

  const largest = stations.reduce((best, s) => {
    return Number(s.currentWip || 0) > Number(best.currentWip || 0) ? s : best;
  }, { displayName: "--", currentWip: 0 });

  return {
    incomingWip,
    meiFeedWip,
    edgingWip,
    mountingDrillWip,
    finalInspectionWip,
    sideStationWip,
    totalWip,
    totalActivityToday,
    largestWipTotal: largest.currentWip || 0,
    largestWipStation: largest.displayName || largest.flowStep || "--"
  };
}

/* ===============================
   STATUS SYSTEM
================================ */

function applyConfigToStation(station, config) {
  switch (station.flowStep) {
    case "Finish Unbox":
      station.expectedNormalPerHour = config.unboxRate * config.unboxCount;
      break;

    case "MEI Line B":
    case "MEI Line C":
    case "MEI Easy Fit":
      station.expectedNormalPerHour = config.meiRate * config.meiMachines;
      break;

    case "Mounting":
      station.expectedNormalPerHour = config.mountRate * config.mountCount;
      break;

    case "Drill":
      station.expectedNormalPerHour = config.drillRate * config.drillCount;
      break;

case "Final Inspection":
  station.expectedNormalPerHour = config.finalRate * config.finalCount;
  break;

case "Bigs":
case "Sharps":
  station.expectedNormalPerHour = 0;
  break;

default:
  station.expectedNormalPerHour = station.expectedNormalPerHour || 0;
  }

  station.currentWip = Number(station.currentWip || 0);
  station.activityToday = Number(station.activityToday || 0);
  station.lastHourActivity = Number(station.lastHourActivity || 0);

  station.utilizationPct = calculateUtilization(station);
  station.status = calculateStatus(station);
  station.statusLabel = getStatusLabel(station.status);
  station.statusMessage = getStatusMessage(station);

  return station;
}

function calculateUtilization(station) {
  const capacity = Number(station.expectedNormalPerHour || 0);
  const lastHour = Number(station.lastHourActivity || 0);

  if (!capacity || capacity <= 0) return 0;

  return Math.round((lastHour / capacity) * 100);
}

function calculateStatus(station) {
  const wip = Number(station.currentWip || 0);
  const util = Number(station.utilizationPct || 0);
  const capacity = Number(station.expectedNormalPerHour || 0);

  if (station.metricMode === "WIP_ONLY") return "NORMAL";

  if (wip > 0 && util === 0 && capacity > 0) return "STARVED";
  if (util < 30 && capacity > 0) return "CRITICAL";
  if (util < 70 && capacity > 0) return "WARNING";
  if (util > 150) return "OVERLOAD";

  return "NORMAL";
}

function getStatusLabel(status) {
  switch (status) {
    case "CRITICAL": return "Critical";
    case "WARNING": return "Warning";
    case "STARVED": return "Starved";
    case "OVERLOAD": return "Overload";
    case "NORMAL": return "Normal";
    default: return "Normal";
  }
}

function getStatusMessage(station) {
  const name = station.displayName || station.flowStep || "Station";

  switch (station.status) {
    case "CRITICAL":
      return `${name} is running below expected threshold. Review staffing, feed, or station availability.`;

    case "WARNING":
      return `${name} is below target pace. Monitor before it becomes a bottleneck.`;

    case "STARVED":
      return `${name} has WIP but no recent throughput signal. Validate scan activity or station movement.`;

    case "OVERLOAD":
      return `${name} is above normal utilization. Watch for downstream buildup.`;

    default:
      return `${name} is operating within normal range.`;
  }
}

function renderStatusPanel(stations) {
  const container = document.getElementById("statusPanelContainer");
  if (!container) return;

  const groups = {
    CRITICAL: stations.filter(s => s.status === "CRITICAL"),
    WARNING: stations.filter(s => s.status === "WARNING"),
    STARVED: stations.filter(s => s.status === "STARVED"),
    OVERLOAD: stations.filter(s => s.status === "OVERLOAD"),
    NORMAL: stations.filter(s => s.status === "NORMAL")
  };

  container.innerHTML = `
    ${buildStatusGroup("Critical Alerts", groups.CRITICAL, "critical")}
    ${buildStatusGroup("Warnings", groups.WARNING, "warning")}
    ${buildStatusGroup("Starved Stations", groups.STARVED, "starved")}
    ${buildStatusGroup("Overload Watch", groups.OVERLOAD, "overload")}
    ${buildStatusGroup("Normal Operations", groups.NORMAL, "normal")}
  `;
}

function buildStatusGroup(title, stations, className) {
  if (!stations.length) return "";

  return `
    <div class="status-group ${className}">
      <div class="status-group-title">${title}</div>
      <div class="status-grid">
        ${stations.map(buildStatusCard).join("")}
      </div>
    </div>
  `;
}

function buildStatusCard(station) {
  const statusClass = String(station.status || "NORMAL").toLowerCase();
  const util = Number(station.utilizationPct || 0);
  const capped = Math.max(0, Math.min(util, 100));

  return `
    <article class="status-card status-${statusClass}">
      <div class="status-card-top">
        <h3>${escapeHtml(station.displayName || station.flowStep || "--")}</h3>
        <span>${escapeHtml(station.statusLabel || "Normal")}</span>
      </div>

      <div class="status-metrics">
        <div>
          <span>WIP</span>
          <strong>${formatNumber(station.currentWip || 0)}</strong>
        </div>
        <div>
          <span>Today</span>
          <strong>${formatNumber(station.activityToday || 0)}</strong>
        </div>
        <div>
          <span>Last Hr</span>
          <strong>${formatNumber(station.lastHourActivity || 0)}</strong>
        </div>
      </div>

      <div class="util-bar">
        <i style="width:${capped}%"></i>
      </div>

      <div class="util-copy">
        <span>${util}% utilization</span>
        <span>${formatNumber(station.expectedNormalPerHour || 0)}/hr target</span>
      </div>

      <p>${escapeHtml(station.statusMessage || "")}</p>
    </article>
  `;
}

function renderStationDetail(stations) {
  const body = document.getElementById("stationDetailBody");
  if (!body) return;

  if (!stations.length) {
    body.innerHTML = `<tr><td colspan="6">No station data available.</td></tr>`;
    return;
  }

  body.innerHTML = stations.map(station => `
    <tr>
      <td>${escapeHtml(station.displayName || station.flowStep || "--")}</td>
      <td>${formatNumber(station.currentWip || 0)}</td>
      <td>${formatNumber(station.activityToday || 0)}</td>
      <td>${formatNumber(station.lastHourActivity || 0)}</td>
      <td>${formatNumber(station.expectedNormalPerHour || 0)}</td>
      <td><span class="table-status ${String(station.status || "NORMAL").toLowerCase()}">${escapeHtml(station.statusLabel || "Normal")}</span></td>
    </tr>
  `).join("");
}

/* ===============================
   HOURLY BREAKDOWN
================================ */

function renderHourlyBreakdown(stations) {
  const container = document.getElementById("hourlyBreakdownGrid");
  if (!container) return;

  const safeStations = Array.isArray(stations) ? stations : [];

  const hourOrder = [
    "6:00 AM", "7:00 AM", "8:00 AM", "9:00 AM", "10:00 AM",
    "11:00 AM", "12:00 PM", "1:00 PM", "2:00 PM", "3:00 PM",
    "4:00 PM", "5:00 PM", "6:00 PM"
  ];

  container.innerHTML = safeStations.map(station => {
    const baseTarget = Number(station.expectedNormalPerHour || 0);
    const stationName = station.displayName || station.flowStep || "--";
    const isNoTarget = station.flowStep === "Bigs" || station.flowStep === "Sharps" || baseTarget <= 0;

    const cells = hourOrder.map(hour => {
      const count = Number(station.hourly?.[hour] || 0);
      const productiveMinutes = getProductiveMinutesForHour(hour);
      const adjustedTarget = isNoTarget ? 0 : Math.round(baseTarget * (productiveMinutes / 60));
      const pct = adjustedTarget > 0 ? Math.round((count / adjustedTarget) * 100) : 0;

      let statusClass = "hour-red";
      let statusText = "Below Rate";

      if (isNoTarget) {
        statusClass = "hour-neutral";
        statusText = "No Target";
      } else if (productiveMinutes <= 0) {
        statusClass = "hour-neutral";
        statusText = "Off Shift";
      } else if (pct >= 100) {
        statusClass = "hour-green";
        statusText = "Hit Rate";
      } else if (pct >= 90) {
        statusClass = "hour-amber";
        statusText = "Near Rate";
      }

      return `
        <div class="hour-cell ${statusClass}" title="${escapeHtml(stationName)} · ${escapeHtml(hour)}">
          <span>${escapeHtml(hour)}</span>
          <strong>${formatNumber(count)}</strong>
          <small>${isNoTarget ? "No Target" : `${pct}%`}</small>
          <em>${isNoTarget ? "Count Only" : `${formatNumber(adjustedTarget)} target`}</em>
        </div>
      `;
    }).join("");

    return `
      <article class="hourly-station-card">
        <div class="hourly-station-header">
          <div>
            <h3>${escapeHtml(stationName)}</h3>
            <small>
              Today: ${formatNumber(station.activityToday || station.hourlyTotal || 0)}
              · Last Hr: ${formatNumber(station.lastHourActivity || 0)}
            </small>
          </div>
          <span>${isNoTarget ? "No hourly target" : `${formatNumber(baseTarget)} / hr target`}</span>
        </div>

        <div class="hourly-cells">
          ${cells}
        </div>
      </article>
    `;
  }).join("");
}

function getCurrentShiftRule() {
  const day = new Date().getDay();
  const isWeekend = day === 0 || day === 6;
  return isWeekend ? SHIFT_RULES.weekend : SHIFT_RULES.weekday;
}

function getProductiveMinutesForHour(hourLabel) {
  const rule = getCurrentShiftRule();

  const hourStart = timeToMinutes(hourLabel);
  const hourEnd = hourStart + 60;

  const shiftStart = timeToMinutes(rule.shiftStart);
  const shiftEnd = timeToMinutes(rule.shiftEnd);

  let productiveStart = Math.max(hourStart, shiftStart);
  let productiveEnd = Math.min(hourEnd, shiftEnd);

  if (productiveEnd <= productiveStart) return 0;

  let productiveMinutes = productiveEnd - productiveStart;

  rule.breaks.forEach(breakItem => {
    const breakStart = timeToMinutes(breakItem.start);
    const breakEnd = timeToMinutes(breakItem.end);

    const overlapStart = Math.max(productiveStart, breakStart);
    const overlapEnd = Math.min(productiveEnd, breakEnd);

    if (overlapEnd > overlapStart) {
      productiveMinutes -= overlapEnd - overlapStart;
    }
  });

  return Math.max(0, productiveMinutes);
}

function timeToMinutes(label) {
  const text = String(label || "").trim().toUpperCase();
  const match = text.match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/);

  if (!match) return 0;

  let hour = Number(match[1]);
  const minute = Number(match[2] || 0);
  const suffix = match[3];

  if (suffix === "PM" && hour !== 12) hour += 12;
  if (suffix === "AM" && hour === 12) hour = 0;

  return hour * 60 + minute;
}

/* ===============================
   FLOOR STATE CLASSES
================================ */

function applyStationStateClasses(stations) {
  const map = {
    "FSV Scan & Verify": ".station-fsv",
    "Finish Unbox": ".station-unbox",
    "AR-OUT": ".station-arout",
    "MEI Line B": ".mei-b",
    "MEI Easy Fit": ".easy-fit",
    "MEI Line C": ".mei-c",
    "Mounting": ".process-card:nth-child(1), .rail-card.orange-card:nth-of-type(1)",
    "Drill": ".process-card:nth-child(2), .rail-card.orange-card:nth-of-type(2)",
    "Bigs": ".purple-card:nth-of-type(3)",
    "Sharps": ".purple-card:nth-of-type(4)",
    "Final Inspection": ".final-card"
  };

  document
    .querySelectorAll(".state-normal, .state-warning, .state-critical, .state-starved, .state-overload")
    .forEach(el => {
      el.classList.remove(
        "state-normal",
        "state-warning",
        "state-critical",
        "state-starved",
        "state-overload"
      );
    });

  stations.forEach(station => {
    const selector = map[station.flowStep];
    if (!selector) return;

    document.querySelectorAll(selector).forEach(el => {
      el.classList.add(`state-${String(station.status || "NORMAL").toLowerCase()}`);
    });
  });
}

/* ===============================
   CONFIG
================================ */

function loadConfig() {
  try {
    const saved = localStorage.getItem("finishCapacityConfig");
    return saved ? JSON.parse(saved) : { ...DEFAULT_CONFIG };
  } catch {
    return { ...DEFAULT_CONFIG };
  }
}

function saveConfig(config) {
  localStorage.setItem("finishCapacityConfig", JSON.stringify(config));
}

function initConfigPanel() {
  const config = loadConfig();

  setInputValue("cfgUnboxRate", config.unboxRate);
  setInputValue("cfgUnboxCount", config.unboxCount);
  setInputValue("cfgMeiRate", config.meiRate);
  setInputValue("cfgMeiMachines", config.meiMachines);
  setInputValue("cfgMountRate", config.mountRate);
  setInputValue("cfgMountCount", config.mountCount);
  setInputValue("cfgDrillRate", config.drillRate);
  setInputValue("cfgDrillCount", config.drillCount);
  setInputValue("cfgFinalRate", config.finalRate);
  setInputValue("cfgFinalCount", config.finalCount);

  updateConfigTotals();

  [
    "cfgUnboxRate",
    "cfgUnboxCount",
    "cfgMeiRate",
    "cfgMeiMachines",
    "cfgMountRate",
    "cfgMountCount",
    "cfgDrillRate",
    "cfgDrillCount",
    "cfgFinalRate",
    "cfgFinalCount"
  ].forEach(id => {
    document.getElementById(id)?.addEventListener("input", updateConfigTotals);
  });

  document.getElementById("configSaveBtn")?.addEventListener("click", () => {
    const newConfig = {
      unboxRate: getInputValue("cfgUnboxRate"),
      unboxCount: getInputValue("cfgUnboxCount"),
      meiRate: getInputValue("cfgMeiRate"),
      meiMachines: getInputValue("cfgMeiMachines"),
      mountRate: getInputValue("cfgMountRate"),
      mountCount: getInputValue("cfgMountCount"),
      drillRate: getInputValue("cfgDrillRate"),
      drillCount: getInputValue("cfgDrillCount"),
      finalRate: getInputValue("cfgFinalRate"),
      finalCount: getInputValue("cfgFinalCount")
    };

    saveConfig(newConfig);
    flashButton("configSaveBtn", "Saved");
    loadDashboard();
  });

  document.getElementById("configResetBtn")?.addEventListener("click", () => {
    saveConfig(DEFAULT_CONFIG);
    initConfigPanel();
    loadDashboard();
  });
}

function updateConfigTotals() {
  const unboxTotal = getInputValue("cfgUnboxRate") * getInputValue("cfgUnboxCount");
  const meiTotal = getInputValue("cfgMeiRate") * getInputValue("cfgMeiMachines");
  const mountTotal = getInputValue("cfgMountRate") * getInputValue("cfgMountCount");
  const drillTotal = getInputValue("cfgDrillRate") * getInputValue("cfgDrillCount");
  const finalTotal = getInputValue("cfgFinalRate") * getInputValue("cfgFinalCount");

  setText("cfgUnboxTotal", unboxTotal);
  setText("cfgMeiTotal", meiTotal);
  setText("cfgMountTotal", mountTotal);
  setText("cfgDrillTotal", drillTotal);
  setText("cfgFinalTotal", finalTotal);
}

/* ===============================
   TABS
================================ */

function initTabs() {
  const buttons = document.querySelectorAll(".tab-btn");
  const contents = document.querySelectorAll(".tab-content");

  buttons.forEach(button => {
    button.addEventListener("click", () => {
      const tab = button.dataset.tab;

      buttons.forEach(btn => btn.classList.remove("active"));
      contents.forEach(content => content.classList.remove("active"));

      button.classList.add("active");

      const target = document.querySelector(`[data-content="${tab}"]`);
      if (target) target.classList.add("active");
    });
  });
}

/* ===============================
   UI EFFECTS
================================ */

function initClock() {
  updateClock();
  setInterval(updateClock, 1000);
}

function updateClock() {
  setText("clock", new Date().toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  }));
}

function updateSyncAge() {
  if (!lastSyncTime) {
    setText("syncAge", "--");
    return;
  }

  const seconds = Math.floor((Date.now() - lastSyncTime) / 1000);

  if (seconds < 60) {
    setText("syncAge", `${seconds}s ago`);
  } else {
    setText("syncAge", `${Math.floor(seconds / 60)}m ago`);
  }
}

function initRefreshButton() {
  document.getElementById("refreshBtn")?.addEventListener("click", loadDashboard);
}

function initHoverEffects() {
  document.querySelectorAll(".station-card, .machine-card, .process-card, .rail-card, .kpi-card").forEach(card => {
    card.addEventListener("mousemove", event => {
      const rect = card.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;

      card.style.setProperty("--mx", `${x}px`);
      card.style.setProperty("--my", `${y}px`);
    });
  });
}

function showLoading(show) {
  const loader = document.getElementById("loadingScreen");
  if (!loader) return;

  if (show) {
    loader.classList.remove("hidden");
  } else {
    setTimeout(() => {
      loader.classList.add("hidden");
    }, 350);
  }
}

function renderError(error) {
  console.error(error);

  setText("syncAge", "error");

  const container = document.getElementById("statusPanelContainer");
  if (container) {
    container.innerHTML = `
      <article class="status-card status-critical">
        <div class="status-card-top">
          <h3>Dashboard API Error</h3>
          <span>Critical</span>
        </div>
        <p>${escapeHtml(error.message || "Unable to load Finish Dashboard data.")}</p>
      </article>
    `;
  }
}

function flashButton(id, label) {
  const btn = document.getElementById(id);
  if (!btn) return;

  const oldText = btn.textContent;
  btn.textContent = label;
  btn.classList.add("saved");

  setTimeout(() => {
    btn.textContent = oldText;
    btn.classList.remove("saved");
  }, 1500);
}

/* ===============================
   HELPERS
================================ */

function setText(id, value) {
  const el = document.getElementById(id);
  if (!el) return;

  if (typeof value === "number") {
    el.textContent = formatNumber(value);
  } else {
    el.textContent = value ?? "0";
  }
}

function getInputValue(id) {
  const el = document.getElementById(id);
  return el ? Number(el.value || 0) : 0;
}

function setInputValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString();
}

function formatTime(date) {
  return new Date(date).toLocaleString([], {
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}