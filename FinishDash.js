const FINISH_API_URL =
  "https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec?area=Finish";

const FINISH_OPERATOR_API_URL =
  "https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec?action=operatorActivity&area=Finish";

let lastSyncTime = null;

const DEFAULT_CONFIG = {
  unboxRate: 150,
  unboxCount: 1.5,

  meiARate: 50,
  meiACount: 0,
  meiBRate: 50,
  meiBCount: 5,
  meiCRate: 50,
  meiCCount: 5,

  mountRate: 25,
  mountCount: 10,

 mountTraineeW1: 0,
 mountTraineeW2: 0,
 mountTraineeW3: 0,
 mountTraineeW4: 0,
 mountTraineeW5: 0,
 mountTraineeW6: 0,
 mountTraineeW7: 0,
 mountTraineeW8: 0, 

  drillRate: 5,
  drillCount: 1,
 finalRate: 75,
finalCount: 3,

finalTraineeW1: 0,
finalTraineeW2: 0,
finalTraineeW3: 0,
finalTraineeW4: 0,
finalTraineeW5: 0,

  /* Operator-based capacity assignments.
     role: ignore | core | tq | training
     trainingWeek applies only to Mounting and Final Inspection. */
  operatorAssignments: {}

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
  initOperatorCommandUI();

  loadDashboard();
  loadFinishOperatorActivity();

  setInterval(loadDashboard, 5 * 60 * 1000);
  setInterval(loadFinishOperatorActivity, 5 * 60 * 1000);
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
  renderManagementSummary(summary, healthStations);
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
  setText("summaryLargestValue", summary.largestWipTotal || 0);
  setText("summaryLargestName", summary.largestWipStation || "--");
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


function renderManagementSummary(summary, stations) {
  const safeStations = Array.isArray(stations) ? stations : [];

  const riskStations = safeStations.filter(s =>
    ["CRITICAL", "WARNING", "STARVED", "OVERLOAD"].includes(String(s.status || "").toUpperCase())
  );

  const criticalStations = safeStations.filter(s =>
    String(s.status || "").toUpperCase() === "CRITICAL"
  );

  const bottleneck = safeStations.reduce((best, station) => {
    return Number(station.currentWip || 0) > Number(best.currentWip || 0) ? station : best;
  }, { displayName: "--", currentWip: 0 });

  const totalWip = Number(summary.totalWip || 0);
  const bottleneckName = bottleneck.displayName || bottleneck.flowStep || summary.largestWipStation || "--";
  const bottleneckWip = Number(bottleneck.currentWip || summary.largestWipTotal || 0);

  setText("summaryRiskCount", riskStations.length);

  if (criticalStations.length > 0) {
    setText("mgmtSummaryTitle", "Immediate attention required");
    setText(
      "mgmtSummaryText",
      `${criticalStations.length} station(s) are currently below expected performance. The largest WIP pressure is ${bottleneckName} with ${formatNumber(bottleneckWip)} jobs.`
    );
    setText(
      "mgmtRecommendation",
      `Prioritize support around ${bottleneckName}. Review staffing, machine availability, and upstream feed before WIP continues building.`
    );
    setText(
      "mgmtNextAction",
      `Check ${criticalStations.map(s => s.displayName || s.flowStep).join(", ")} and confirm whether the issue is labor, equipment, or scan timing.`
    );
    return;
  }

  if (riskStations.length > 0) {
    setText("mgmtSummaryTitle", "Operation stable but needs monitoring");
    setText(
      "mgmtSummaryText",
      `${riskStations.length} station(s) are trending below target. Total WIP is ${formatNumber(totalWip)}, with the highest load at ${bottleneckName}.`
    );
    setText(
      "mgmtRecommendation",
      `Monitor ${bottleneckName} and confirm the next hourly update improves before shifting more work downstream.`
    );
    setText(
      "mgmtNextAction",
      "Use the Hourly Breakdown tab to confirm whether the slowdown is isolated to one hour or becoming a trend."
    );
    return;
  }

  setText("mgmtSummaryTitle", "Operation running within control");
  setText(
    "mgmtSummaryText",
    `Finish is currently operating within expected range. Total WIP is ${formatNumber(totalWip)}, and the largest active WIP point is ${bottleneckName}.`
  );
  setText(
    "mgmtRecommendation",
    "Maintain current staffing plan and continue watching the bottleneck station during the next refresh cycle."
  );
  setText(
    "mgmtNextAction",
    "No immediate escalation needed. Continue monitoring WIP movement and hourly pace."
  );
}

/* ===============================
   STATUS SYSTEM
================================ */

function applyConfigToStation(station, config) {
  const cfg = config || loadConfig();
  const flowStep = normalizeFinishOperatorStation(station.flowStep || station.displayName || station.name);
  station.flowStep = flowStep;

  switch (flowStep) {
    case "Finish Unbox": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.unboxCount || 0);
      station.expectedNormalPerHour = Number(cfg.unboxRate || 0) * count;
      break;
    }

    case "MEI Line A": {
      const count = Number(cfg.meiACount || 0);
      station.expectedNormalPerHour = Number(cfg.meiARate || 0) * count;
      break;
    }

    case "MEI Line B": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.meiBCount || 0);
      station.expectedNormalPerHour = Number(cfg.meiBRate || 0) * count;
      break;
    }

    case "MEI Line C": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.meiCCount || 0);
      station.expectedNormalPerHour = Number(cfg.meiCRate || 0) * count;
      break;
    }

    case "MEI Easy Fit": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.meiBCount || 0);
      station.expectedNormalPerHour = Number(cfg.meiBRate || 0) * count;
      break;
    }

    case "Mounting": {
      const assignedCoreCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const coreCount = assignedCoreCount > 0 ? assignedCoreCount : Number(cfg.mountCount || 0);
      const coreTotal = Number(cfg.mountRate || 0) * coreCount;
      const assignedTrainingTotal = getConfiguredTrainingOperatorTotal(flowStep, cfg);
      const numericTrainingTotal = getNumericMountTrainingTotal(cfg);
      station.expectedNormalPerHour = coreTotal + (assignedTrainingTotal > 0 ? assignedTrainingTotal : numericTrainingTotal);
      break;
    }

    case "Drill": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.drillCount || 0);
      station.expectedNormalPerHour = Number(cfg.drillRate || 0) * count;
      break;
    }

    case "Final Inspection": {
      const assignedCoreCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const coreCount = assignedCoreCount > 0 ? assignedCoreCount : Number(cfg.finalCount || 0);
      const coreTotal = Number(cfg.finalRate || 0) * coreCount;
      const assignedTrainingTotal = getConfiguredTrainingOperatorTotal(flowStep, cfg);
      const numericTrainingTotal = getNumericFinalTrainingTotal(cfg);
      station.expectedNormalPerHour = coreTotal + (assignedTrainingTotal > 0 ? assignedTrainingTotal : numericTrainingTotal);
      break;
    }

    case "Bigs":
    case "Sharps":
      station.expectedNormalPerHour = 0;
      break;

    default:
      station.expectedNormalPerHour = Number(station.expectedNormalPerHour || 0);
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

  if (!safeStations.length) {
    container.innerHTML = `
      <article class="hourly-station-card">
        <div class="hourly-station-header">
          <h3>No hourly data available</h3>
          <span>--</span>
        </div>
      </article>
    `;
    return;
  }

  container.innerHTML = safeStations.map((station, index) => {
    const baseTarget = Number(station.expectedNormalPerHour || 0);
    const stationName = station.displayName || station.flowStep || "--";
    const isNoTarget = station.flowStep === "Bigs" || station.flowStep === "Sharps" || baseTarget <= 0;

    const cells = hourOrder.map(hour => {
      const count = Number(station.hourly?.[hour] || 0);
      const productiveMinutes = getProductiveMinutesForHour(hour);
      const adjustedTarget = isNoTarget ? 0 : Math.round(baseTarget * (productiveMinutes / 60));
      const pct = adjustedTarget > 0 ? Math.round((count / adjustedTarget) * 100) : 0;

      let statusClass = "hour-red";

      if (isNoTarget || productiveMinutes <= 0) {
        statusClass = "hour-neutral";
      } else if (pct >= 100) {
        statusClass = "hour-green";
      } else if (pct >= 90) {
        statusClass = "hour-amber";
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
      <details class="hourly-accordion" ${index === 0 ? "open" : ""}>
        <summary class="hourly-accordion-header">
          <div>
            <h3>${escapeHtml(stationName)}</h3>
            <small>
              Today: ${formatNumber(station.activityToday || station.hourlyTotal || 0)}
              · Last Hr: ${formatNumber(station.lastHourActivity || 0)}
            </small>
          </div>

          <div class="hourly-accordion-right">
            <span>${isNoTarget ? "No hourly target" : `${formatNumber(baseTarget)} / hr target`}</span>
            <b>Open / Close</b>
          </div>
        </summary>

        <div class="hourly-cells">
          ${cells}
        </div>
      </details>
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
    return saved ? { ...DEFAULT_CONFIG, ...JSON.parse(saved) } : { ...DEFAULT_CONFIG };
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

  setInputValue("cfgMeiARate", config.meiARate);
  setInputValue("cfgMeiACount", config.meiACount);
  setInputValue("cfgMeiBRate", config.meiBRate);
  setInputValue("cfgMeiBCount", config.meiBCount);
  setInputValue("cfgMeiCRate", config.meiCRate);
  setInputValue("cfgMeiCCount", config.meiCCount);

  setInputValue("cfgMountRate", config.mountRate);
  setInputValue("cfgMountCount", config.mountCount);
  setInputValue("cfgMountTraineeW1", config.mountTraineeW1);
 setInputValue("cfgMountTraineeW2", config.mountTraineeW2);
 setInputValue("cfgMountTraineeW3", config.mountTraineeW3);
 setInputValue("cfgMountTraineeW4", config.mountTraineeW4);
 setInputValue("cfgMountTraineeW5", config.mountTraineeW5);
 setInputValue("cfgMountTraineeW6", config.mountTraineeW6);
 setInputValue("cfgMountTraineeW7", config.mountTraineeW7);
 setInputValue("cfgMountTraineeW8", config.mountTraineeW8);

  setInputValue("cfgDrillRate", config.drillRate);
  setInputValue("cfgDrillCount", config.drillCount);
  setInputValue("cfgFinalRate", config.finalRate);
  setInputValue("cfgFinalCount", config.finalCount);
  setInputValue("cfgFinalTraineeW1", config.finalTraineeW1);
setInputValue("cfgFinalTraineeW2", config.finalTraineeW2);
setInputValue("cfgFinalTraineeW3", config.finalTraineeW3);
setInputValue("cfgFinalTraineeW4", config.finalTraineeW4);
setInputValue("cfgFinalTraineeW5", config.finalTraineeW5);

  ensureOperatorAssignmentConfigPanel();
  renderOperatorAssignmentConfigPanel();
  updateConfigTotals();

  [
    "cfgUnboxRate",
    "cfgUnboxCount",
    "cfgMeiARate",
    "cfgMeiACount",
    "cfgMeiBRate",
    "cfgMeiBCount",
    "cfgMeiCRate",
    "cfgMeiCCount",
    "cfgMountRate",
    "cfgMountTraineeW1",
    "cfgMountTraineeW2",
    "cfgMountTraineeW3",
    "cfgMountTraineeW4",
    "cfgMountTraineeW5",
    "cfgMountTraineeW6",
    "cfgMountTraineeW7",
    "cfgMountTraineeW8",
    "cfgMountCount",
    "cfgDrillRate",
    "cfgDrillCount",
    "cfgFinalRate",
"cfgFinalCount",
"cfgFinalTraineeW1",
"cfgFinalTraineeW2",
"cfgFinalTraineeW3",
"cfgFinalTraineeW4",
"cfgFinalTraineeW5"

  ].forEach(id => {
    document.getElementById(id)?.addEventListener("input", updateConfigTotals);
  });

  document.getElementById("configSaveBtn")?.addEventListener("click", () => {
    const newConfig = {
      unboxRate: getInputValue("cfgUnboxRate"),
      unboxCount: getInputValue("cfgUnboxCount"),

      meiARate: getInputValue("cfgMeiARate"),
      meiACount: getInputValue("cfgMeiACount"),
      meiBRate: getInputValue("cfgMeiBRate"),
      meiBCount: getInputValue("cfgMeiBCount"),
      meiCRate: getInputValue("cfgMeiCRate"),
      meiCCount: getInputValue("cfgMeiCCount"),

      mountRate: getInputValue("cfgMountRate"),
      mountCount: getInputValue("cfgMountCount"),
      mountTraineeW1: getInputValue("cfgMountTraineeW1"),
mountTraineeW2: getInputValue("cfgMountTraineeW2"),
mountTraineeW3: getInputValue("cfgMountTraineeW3"),
mountTraineeW4: getInputValue("cfgMountTraineeW4"),
mountTraineeW5: getInputValue("cfgMountTraineeW5"),
mountTraineeW6: getInputValue("cfgMountTraineeW6"),
mountTraineeW7: getInputValue("cfgMountTraineeW7"),
mountTraineeW8: getInputValue("cfgMountTraineeW8"),
      drillRate: getInputValue("cfgDrillRate"),
      drillCount: getInputValue("cfgDrillCount"),
      finalRate: getInputValue("cfgFinalRate"),
finalCount: getInputValue("cfgFinalCount"),

finalTraineeW1: getInputValue("cfgFinalTraineeW1"),
finalTraineeW2: getInputValue("cfgFinalTraineeW2"),
finalTraineeW3: getInputValue("cfgFinalTraineeW3"),
finalTraineeW4: getInputValue("cfgFinalTraineeW4"),
finalTraineeW5: getInputValue("cfgFinalTraineeW5"),
      operatorAssignments: loadConfig().operatorAssignments || {}

    };

    saveConfig(newConfig);
    flashButton("configSaveBtn", "Saved");
    loadDashboard();
  });

  document.getElementById("configResetBtn")?.addEventListener("click", () => {
    localStorage.removeItem("finishCapacityConfig");
    initConfigPanel();
    loadDashboard();
  });
}

function updateConfigTotals() {
  const saved = loadConfig();
  const config = {
    ...saved,
    unboxRate: getInputValue("cfgUnboxRate"),
    unboxCount: getInputValue("cfgUnboxCount"),
    meiARate: getInputValue("cfgMeiARate"),
    meiACount: getInputValue("cfgMeiACount"),
    meiBRate: getInputValue("cfgMeiBRate"),
    meiBCount: getInputValue("cfgMeiBCount"),
    meiCRate: getInputValue("cfgMeiCRate"),
    meiCCount: getInputValue("cfgMeiCCount"),
    mountRate: getInputValue("cfgMountRate"),
    mountCount: getInputValue("cfgMountCount"),
    mountTraineeW1: getInputValue("cfgMountTraineeW1"),
    mountTraineeW2: getInputValue("cfgMountTraineeW2"),
    mountTraineeW3: getInputValue("cfgMountTraineeW3"),
    mountTraineeW4: getInputValue("cfgMountTraineeW4"),
    mountTraineeW5: getInputValue("cfgMountTraineeW5"),
    mountTraineeW6: getInputValue("cfgMountTraineeW6"),
    mountTraineeW7: getInputValue("cfgMountTraineeW7"),
    mountTraineeW8: getInputValue("cfgMountTraineeW8"),
    drillRate: getInputValue("cfgDrillRate"),
    drillCount: getInputValue("cfgDrillCount"),
    finalRate: getInputValue("cfgFinalRate"),
    finalCount: getInputValue("cfgFinalCount"),
    finalTraineeW1: getInputValue("cfgFinalTraineeW1"),
    finalTraineeW2: getInputValue("cfgFinalTraineeW2"),
    finalTraineeW3: getInputValue("cfgFinalTraineeW3"),
    finalTraineeW4: getInputValue("cfgFinalTraineeW4"),
    finalTraineeW5: getInputValue("cfgFinalTraineeW5"),
    operatorAssignments: saved.operatorAssignments || {}
  };

  const unboxAssigned = getConfiguredCoreOperatorCount("Finish Unbox", config);
  const unboxCount = unboxAssigned > 0 ? unboxAssigned : config.unboxCount;
  const unboxTotal = Number(config.unboxRate || 0) * Number(unboxCount || 0);

  const meiATotal = Number(config.meiARate || 0) * Number(config.meiACount || 0);
  const meiBAssigned = getConfiguredCoreOperatorCount("MEI Line B", config);
  const meiBAssignedCount = meiBAssigned > 0 ? meiBAssigned : config.meiBCount;
  const meiBTotal = Number(config.meiBRate || 0) * Number(meiBAssignedCount || 0);

  const meiCAssigned = getConfiguredCoreOperatorCount("MEI Line C", config);
  const meiCAssignedCount = meiCAssigned > 0 ? meiCAssigned : config.meiCCount;
  const meiCTotal = Number(config.meiCRate || 0) * Number(meiCAssignedCount || 0);

  const mountAssignedCore = getConfiguredCoreOperatorCount("Mounting", config);
  const mountCount = mountAssignedCore > 0 ? mountAssignedCore : config.mountCount;
  const mountTotal = Number(config.mountRate || 0) * Number(mountCount || 0);
  const assignedMountTrainingTotal = getConfiguredTrainingOperatorTotal("Mounting", config);
  const trainingMountTotal = assignedMountTrainingTotal > 0
    ? assignedMountTrainingTotal
    : getNumericMountTrainingTotal(config);

  const drillAssigned = getConfiguredCoreOperatorCount("Drill", config);
  const drillCount = drillAssigned > 0 ? drillAssigned : config.drillCount;
  const drillTotal = Number(config.drillRate || 0) * Number(drillCount || 0);

  const finalAssignedCore = getConfiguredCoreOperatorCount("Final Inspection", config);
  const finalCount = finalAssignedCore > 0 ? finalAssignedCore : config.finalCount;
  const finalTotal = Number(config.finalRate || 0) * Number(finalCount || 0);
  const assignedFinalTrainingTotal = getConfiguredTrainingOperatorTotal("Final Inspection", config);
  const finalTrainingTotal = assignedFinalTrainingTotal > 0
    ? assignedFinalTrainingTotal
    : getNumericFinalTrainingTotal(config);

  setText("cfgUnboxTotal", unboxTotal);
  setText("cfgMeiATotal", meiATotal);
  setText("cfgMeiBTotal", meiBTotal);
  setText("cfgMeiCTotal", meiCTotal);
  setText("cfgMountTotal", mountTotal);
  setText("cfgMountTrainingTotal", trainingMountTotal);
  setText("cfgMountAdjustedTotal", mountTotal + trainingMountTotal);
  setText("cfgDrillTotal", drillTotal);
  setText("cfgFinalTotal", finalTotal);
  setText("cfgFinalTrainingTotal", finalTrainingTotal);
  setText("cfgFinalAdjustedTotal", finalTotal + finalTrainingTotal);

  setText("cfgAssignedMountCore", mountAssignedCore);
  setText("cfgAssignedMountTraining", getConfiguredTrainingOperatorCount("Mounting", config));
  setText("cfgAssignedFinalCore", finalAssignedCore);
  setText("cfgAssignedFinalTraining", getConfiguredTrainingOperatorCount("Final Inspection", config));
}



/* ===============================
   OPERATOR-BASED CAPACITY ASSIGNMENTS
   UI is injected into the existing Configuration tab.
   No backend change required. Saved in localStorage with finishCapacityConfig.
================================ */

const FINISH_ASSIGNMENT_STATIONS = [
  "Finish Unbox",
  "MEI Line B",
  "MEI Line C",
  "MEI Easy Fit",
  "Mounting",
  "Drill",
  "Final Inspection"
];

const MOUNTING_TRAINING_RATES = {
  1: 3,
  2: 6,
  3: 9,
  4: 12,
  5: 15,
  6: 18,
  7: 21,
  8: 25
};

const FINAL_TRAINING_RATES = {
  1: 15,
  2: 30,
  3: 45,
  4: 60,
  5: 75
};

function getNumericMountTrainingTotal(config = loadConfig()) {
  return (
    (Number(config.mountTraineeW1 || 0) * 3) +
    (Number(config.mountTraineeW2 || 0) * 6) +
    (Number(config.mountTraineeW3 || 0) * 9) +
    (Number(config.mountTraineeW4 || 0) * 12) +
    (Number(config.mountTraineeW5 || 0) * 15) +
    (Number(config.mountTraineeW6 || 0) * 18) +
    (Number(config.mountTraineeW7 || 0) * 21) +
    (Number(config.mountTraineeW8 || 0) * 25)
  );
}

function getNumericFinalTrainingTotal(config = loadConfig()) {
  return (
    (Number(config.finalTraineeW1 || 0) * 15) +
    (Number(config.finalTraineeW2 || 0) * 30) +
    (Number(config.finalTraineeW3 || 0) * 45) +
    (Number(config.finalTraineeW4 || 0) * 60) +
    (Number(config.finalTraineeW5 || 0) * 75)
  );
}

function normalizeAssignmentOperatorName(name) {
  return String(name || "").trim();
}

function makeAssignmentKey(stationName, operatorName) {
  return `${normalizeFinishOperatorStation(stationName)}||${normalizeAssignmentOperatorName(operatorName)}`;
}

function getOperatorAssignment(operatorName, stationName, config = loadConfig()) {
  const op = normalizeAssignmentOperatorName(operatorName);
  if (!op) return null;
  const assignments = config.operatorAssignments || {};
  return assignments[makeAssignmentKey(stationName, op)] || null;
}

function getTrainingRateForStationWeek(stationName, week) {
  const name = normalizeFinishOperatorStation(stationName);
  const safeWeek = Number(week || 0);

  if (name === "Mounting") return Number(MOUNTING_TRAINING_RATES[safeWeek] || 0);
  if (name === "Final Inspection") return Number(FINAL_TRAINING_RATES[safeWeek] || 0);

  return 0;
}

function getConfiguredCoreOperatorCount(stationName, config = loadConfig()) {
  const name = normalizeFinishOperatorStation(stationName);
  const assignments = Object.values(config.operatorAssignments || {});

  return assignments.filter(item =>
    normalizeFinishOperatorStation(item.station) === name &&
    ["core", "tq", "certified"].includes(String(item.role || "").toLowerCase())
  ).length;
}

function getConfiguredTrainingOperatorCount(stationName, config = loadConfig()) {
  const name = normalizeFinishOperatorStation(stationName);
  const assignments = Object.values(config.operatorAssignments || {});

  return assignments.filter(item =>
    normalizeFinishOperatorStation(item.station) === name &&
    String(item.role || "").toLowerCase() === "training"
  ).length;
}

function getConfiguredTrainingOperatorTotal(stationName, config = loadConfig()) {
  const name = normalizeFinishOperatorStation(stationName);
  const assignments = Object.values(config.operatorAssignments || {});

  return assignments.reduce((sum, item) => {
    if (normalizeFinishOperatorStation(item.station) !== name) return sum;
    if (String(item.role || "").toLowerCase() !== "training") return sum;
    return sum + getTrainingRateForStationWeek(name, item.trainingWeek);
  }, 0);
}

function getAssignmentRoleLabel(operatorName, stationName, config = loadConfig()) {
  const assignment = getOperatorAssignment(operatorName, stationName, config);
  if (!assignment || assignment.role === "ignore") return "Unassigned";
  if (assignment.role === "training") return `Training W${Number(assignment.trainingWeek || 1)}`;
  if (assignment.role === "tq") return "TQ";
  return "Certified";
}

function getAssignmentTargetLabel(operatorName, stationName, config = loadConfig()) {
  const base = getFinishOperatorBaseRate(stationName, operatorName);
  const role = getAssignmentRoleLabel(operatorName, stationName, config);
  return `${role} · ${numberFmt(base)}/hr`;
}

function getFinishOperatorRoster() {
  const stations = finishOperatorState?.stations || {};
  const roster = [];
  const seen = new Set();

  FINISH_ASSIGNMENT_STATIONS.forEach(stationName => {
    const station = stations[stationName];
    (station?.operatorList || []).forEach(operator => {
      const key = makeAssignmentKey(stationName, operator.name);
      if (seen.has(key)) return;
      seen.add(key);
      roster.push({
        station: stationName,
        operator: operator.name,
        total: Number(operator.total || 0),
        accessPoints: operator.accessPoints || []
      });
    });
  });

  return roster.sort((a, b) => {
    const s = FINISH_ASSIGNMENT_STATIONS.indexOf(a.station) - FINISH_ASSIGNMENT_STATIONS.indexOf(b.station);
    if (s !== 0) return s;
    return Number(b.total || 0) - Number(a.total || 0);
  });
}

function ensureOperatorAssignmentConfigPanel() {
  if (document.getElementById("operatorAssignmentConfigPanel")) return;

  injectOperatorAssignmentStyles();

  const configTab = document.querySelector('[data-content="config"]');
  const saveRow = document.querySelector(".config-actions") || document.getElementById("configSaveBtn")?.parentElement;
  const targetParent = saveRow?.parentElement || configTab;
  if (!targetParent) return;

  const panel = document.createElement("details");
  panel.id = "operatorAssignmentConfigPanel";
  panel.className = "config-group operator-assignment-config-group";
  panel.open = true;
  panel.innerHTML = `
    <summary>
      <div>
        <h3>Operator Capacity Assignment</h3>
        <p>Select who counts as Certified/TQ core capacity and who is on Training Metric.</p>
      </div>
      <span>Open / Close</span>
    </summary>

    <div class="operator-assignment-shell">
      <div class="assignment-warning-card">
        <strong>Important</strong>
        <span>When operators are assigned here, their station capacity and individual JPH target use the selected role. Trainees are graded against their ramp week, not the certified rate.</span>
      </div>

      <div class="assignment-summary-grid">
        <article><span>Mount Core/TQ</span><strong id="cfgAssignedMountCore">0</strong></article>
        <article><span>Mount Training</span><strong id="cfgAssignedMountTraining">0</strong></article>
        <article><span>Final Core/TQ</span><strong id="cfgAssignedFinalCore">0</strong></article>
        <article><span>Final Training</span><strong id="cfgAssignedFinalTraining">0</strong></article>
      </div>

      <div class="assignment-toolbar">
        <label>
          Station Filter
          <select id="assignmentStationFilter">
            <option value="all">All Finish stations</option>
            ${FINISH_ASSIGNMENT_STATIONS.map(station => `<option value="${escapeHtml(station)}">${escapeHtml(station)}</option>`).join("")}
          </select>
        </label>
        <button id="assignmentClearBtn" type="button">Clear Assignments</button>
      </div>

      <div id="operatorAssignmentRoster" class="assignment-roster"></div>
    </div>
  `;

  if (saveRow && saveRow.parentElement) {
    saveRow.parentElement.insertBefore(panel, saveRow);
  } else {
    targetParent.appendChild(panel);
  }

  document.getElementById("assignmentStationFilter")?.addEventListener("change", renderOperatorAssignmentConfigPanel);
  document.getElementById("assignmentClearBtn")?.addEventListener("click", () => {
    const config = loadConfig();
    config.operatorAssignments = {};
    saveConfig(config);
    renderOperatorAssignmentConfigPanel();
    updateConfigTotals();
    loadDashboard();
    if (finishOperatorState?.selectedStation) {
      const station = finishOperatorState.stations?.[finishOperatorState.selectedStation];
      if (station) renderDrawerOperators(station);
    }
  });
}

function renderOperatorAssignmentConfigPanel() {
  ensureOperatorAssignmentConfigPanel();

  const rosterTarget = document.getElementById("operatorAssignmentRoster");
  if (!rosterTarget) return;

  const config = loadConfig();
  const filter = document.getElementById("assignmentStationFilter")?.value || "all";
  const roster = getFinishOperatorRoster()
    .filter(item => filter === "all" || item.station === filter);

  if (!roster.length) {
    rosterTarget.innerHTML = `
      <article class="assignment-empty">
        No live Finish operator roster loaded yet. Click Refresh Operators on the Operator tab or wait for the API refresh.
      </article>`;
    updateConfigTotals();
    return;
  }

  rosterTarget.innerHTML = roster.map(item => {
    const assignment = getOperatorAssignment(item.operator, item.station, config) || {
      role: "ignore",
      trainingWeek: 1
    };
    const canTrain = item.station === "Mounting" || item.station === "Final Inspection";
    const role = String(assignment.role || "ignore");
    const trainingRate = getTrainingRateForStationWeek(item.station, assignment.trainingWeek || 1);
    const baseRate = getFinishOperatorBaseRate(item.station, item.operator);
    const effective = role === "training" ? trainingRate : (["core", "tq", "certified"].includes(role) ? baseRate : 0);

    return `
      <article class="assignment-row" data-assignment-row="${escapeHtml(makeAssignmentKey(item.station, item.operator))}">
        <div class="assignment-operator">
          <strong>${escapeHtml(item.operator)}</strong>
          <span>${escapeHtml(item.station)} · Today ${numberFmt(item.total)}</span>
        </div>

        <label>
          Role
          <select data-assignment-role data-station="${escapeHtml(item.station)}" data-operator="${escapeHtml(item.operator)}">
            <option value="ignore" ${role === "ignore" ? "selected" : ""}>Do not count</option>
            <option value="core" ${role === "core" || role === "certified" ? "selected" : ""}>Certified / Core Capacity</option>
            <option value="tq" ${role === "tq" ? "selected" : ""}>TQ / Core Capacity</option>
            ${canTrain ? `<option value="training" ${role === "training" ? "selected" : ""}>Training Metric</option>` : ""}
          </select>
        </label>

        <label class="assignment-week ${role === "training" && canTrain ? "active" : "disabled"}">
          Week
          <select data-assignment-week data-station="${escapeHtml(item.station)}" data-operator="${escapeHtml(item.operator)}" ${role === "training" && canTrain ? "" : "disabled"}>
            ${buildTrainingWeekOptions(item.station, assignment.trainingWeek || 1)}
          </select>
        </label>

        <div class="assignment-target">
          <span>Target</span>
          <strong>${effective > 0 ? numberFmt(effective) + "/hr" : "—"}</strong>
        </div>
      </article>`;
  }).join("");

  rosterTarget.querySelectorAll("[data-assignment-role]").forEach(select => {
    select.addEventListener("change", event => {
      saveOperatorAssignment(
        event.target.dataset.station,
        event.target.dataset.operator,
        event.target.value,
        Number(getOperatorAssignment(event.target.dataset.operator, event.target.dataset.station)?.trainingWeek || 1)
      );
    });
  });

  rosterTarget.querySelectorAll("[data-assignment-week]").forEach(select => {
    select.addEventListener("change", event => {
      const current = getOperatorAssignment(event.target.dataset.operator, event.target.dataset.station) || { role: "training" };
      saveOperatorAssignment(
        event.target.dataset.station,
        event.target.dataset.operator,
        current.role || "training",
        Number(event.target.value || 1)
      );
    });
  });

  updateConfigTotals();
}

function buildTrainingWeekOptions(stationName, selectedWeek) {
  const rates = normalizeFinishOperatorStation(stationName) === "Final Inspection"
    ? FINAL_TRAINING_RATES
    : MOUNTING_TRAINING_RATES;

  return Object.keys(rates).map(week => `
    <option value="${week}" ${Number(selectedWeek) === Number(week) ? "selected" : ""}>
      Week ${week} · ${rates[week]}/hr
    </option>
  `).join("");
}

function saveOperatorAssignment(stationName, operatorName, role, trainingWeek = 1) {
  const config = loadConfig();
  const station = normalizeFinishOperatorStation(stationName);
  const operator = normalizeAssignmentOperatorName(operatorName);
  if (!operator) return;

  const assignments = { ...(config.operatorAssignments || {}) };
  const key = makeAssignmentKey(station, operator);

  if (!role || role === "ignore") {
    delete assignments[key];
  } else {
    assignments[key] = {
      station,
      operator,
      role,
      trainingWeek: Number(trainingWeek || 1),
      updatedAt: new Date().toISOString()
    };
  }

  config.operatorAssignments = assignments;
  saveConfig(config);
  renderOperatorAssignmentConfigPanel();
  updateConfigTotals();
  loadDashboard();

  if (finishOperatorState?.selectedStation) {
    const activeStation = finishOperatorState.stations?.[finishOperatorState.selectedStation];
    if (activeStation) renderDrawerOperators(activeStation);
  }
}

function injectOperatorAssignmentStyles() {
  if (document.getElementById("operatorAssignmentStyles")) return;

  const style = document.createElement("style");
  style.id = "operatorAssignmentStyles";
  style.textContent = `
    .operator-assignment-shell {
      display: grid;
      gap: 16px;
    }

    .assignment-warning-card {
      display: grid;
      gap: 5px;
      padding: 16px;
      border: 1px solid rgba(251, 191, 36, .28);
      border-radius: 16px;
      background: rgba(251, 191, 36, .07);
      color: #eef8ff;
    }

    .assignment-warning-card strong {
      color: #fbbf24;
      letter-spacing: .08em;
      text-transform: uppercase;
      font-size: 12px;
    }

    .assignment-warning-card span {
      color: #bfdaf2;
      font-size: 13px;
      line-height: 1.45;
    }

    .assignment-summary-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
    }

    .assignment-summary-grid article {
      padding: 14px;
      border: 1px solid rgba(100, 221, 255, .16);
      border-radius: 16px;
      background: rgba(2, 12, 24, .64);
    }

    .assignment-summary-grid span,
    .assignment-toolbar label,
    .assignment-row label,
    .assignment-target span {
      color: #8fa6c3;
      font-family: "JetBrains Mono", monospace;
      font-size: 10px;
      font-weight: 800;
      letter-spacing: .1em;
      text-transform: uppercase;
    }

    .assignment-summary-grid strong {
      display: block;
      margin-top: 8px;
      color: #67e8f9;
      font-size: 28px;
      line-height: 1;
    }

    .assignment-toolbar {
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 14px;
      flex-wrap: wrap;
      padding: 14px;
      border: 1px solid rgba(100, 221, 255, .14);
      border-radius: 16px;
      background: rgba(3, 12, 24, .56);
    }

    .assignment-toolbar label,
    .assignment-row label {
      display: grid;
      gap: 8px;
    }

    .assignment-toolbar select,
    .assignment-row select {
      height: 42px;
      min-width: 230px;
      border-radius: 12px;
      border: 1px solid rgba(0, 217, 255, .28);
      background: #061225;
      color: #eef8ff;
      padding: 0 12px;
      font-weight: 800;
      outline: none;
    }

    .assignment-toolbar button {
      height: 42px;
      padding: 0 16px;
      border-radius: 12px;
      border: 1px solid rgba(251, 113, 133, .35);
      background: rgba(251, 113, 133, .08);
      color: #fb7185;
      font-family: "JetBrains Mono", monospace;
      font-size: 11px;
      font-weight: 900;
      letter-spacing: .08em;
      text-transform: uppercase;
      cursor: pointer;
    }

    .assignment-roster {
      display: grid;
      gap: 10px;
    }

    .assignment-row {
      display: grid;
      grid-template-columns: minmax(220px, 1.2fr) minmax(220px, .8fr) minmax(190px, .7fr) 110px;
      align-items: center;
      gap: 12px;
      padding: 14px;
      border: 1px solid rgba(100, 221, 255, .14);
      border-radius: 16px;
      background: linear-gradient(180deg, rgba(10, 28, 52, .72), rgba(2, 10, 22, .88));
    }

    .assignment-operator strong {
      display: block;
      color: #eef8ff;
      font-size: 14px;
      font-weight: 900;
    }

    .assignment-operator span {
      display: block;
      margin-top: 5px;
      color: #8fa6c3;
      font-size: 12px;
    }

    .assignment-week.disabled {
      opacity: .38;
    }

    .assignment-target {
      display: grid;
      justify-items: end;
      gap: 6px;
    }

    .assignment-target strong {
      color: #00f5a0;
      font-size: 18px;
    }

    .assignment-empty {
      padding: 18px;
      border: 1px dashed rgba(100, 221, 255, .25);
      border-radius: 16px;
      color: #8fa6c3;
      background: rgba(2, 12, 24, .52);
    }

    @media (max-width: 1100px) {
      .assignment-summary-grid,
      .assignment-row {
        grid-template-columns: 1fr;
      }

      .assignment-target {
        justify-items: start;
      }
    }
  `;

  document.head.appendChild(style);
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

let loadingInterval = null;
let loadingProgress = 0;

function startLoadingAnimation() {
  stopLoadingAnimation();

  const statuses = [
    "Connecting to Finish dashboard API...",
    "Reading live station signals...",
    "Calculating WIP and throughput metrics...",
    "Building hourly production performance...",
    "Rendering command center interface..."
  ];

  const stepIds = ["loadStep1", "loadStep2", "loadStep3", "loadStep4"];
  let statusIndex = 0;
  loadingProgress = 0;

  updateLoadingVisuals(loadingProgress, statuses[0], 0, stepIds);

  loadingInterval = setInterval(() => {
    if (loadingProgress < 92) {
      loadingProgress += Math.floor(Math.random() * 8) + 4;
      if (loadingProgress > 92) loadingProgress = 92;
    }

    const stepIndex =
      loadingProgress < 25 ? 0 :
      loadingProgress < 50 ? 1 :
      loadingProgress < 75 ? 2 : 3;

    statusIndex = stepIndex;

    updateLoadingVisuals(
      loadingProgress,
      statuses[statusIndex] || statuses[statuses.length - 1],
      stepIndex,
      stepIds
    );
  }, 350);
}

function finishLoadingAnimation() {
  updateLoadingVisuals(
    100,
    "Dashboard ready.",
    3,
    ["loadStep1", "loadStep2", "loadStep3", "loadStep4"]
  );

  stopLoadingAnimation();
}

function stopLoadingAnimation() {
  if (loadingInterval) {
    clearInterval(loadingInterval);
    loadingInterval = null;
  }
}

function updateLoadingVisuals(
  percent,
  text,
  activeStep,
  stepIds = ["loadStep1", "loadStep2", "loadStep3", "loadStep4"]
) {
  const fill = document.getElementById("loadingProgressFill");
  const percentEl = document.getElementById("loadingPercent");
  const statusText = document.getElementById("loadingStatusText");
  const footerStatus = document.getElementById("loadingFooterStatus");

  if (fill) fill.style.width = `${percent}%`;
  if (percentEl) percentEl.textContent = `${percent}%`;
  if (statusText) statusText.textContent = text;
  if (footerStatus) footerStatus.textContent = text;

  stepIds.forEach((id, index) => {
    const el = document.getElementById(id);
    if (!el) return;

    el.classList.remove("active", "done");

    if (index < activeStep) {
      el.classList.add("done");
    } else if (index === activeStep) {
      el.classList.add("active");
    }
  });
}

function showLoading(show) {
  const loader = document.getElementById("loadingScreen");
  if (!loader) return;

  if (show) {
    loader.classList.remove("hidden");
    startLoadingAnimation();
    return;
  }

  finishLoadingAnimation();

  setTimeout(() => {
    loader.classList.add("hidden");
  }, 350);
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

/* =====================================================
   FINISH OPERATOR COMMAND UI
   Live API: ?action=operatorActivity&area=Finish
===================================================== */

const FINISH_OPERATOR_STATION_ORDER = [
  "Finish Unbox",
  "MEI Line B",
  "MEI Line C",
  "MEI Easy Fit",
  "Mounting",
  "Drill",
  "Bigs",
  "Sharps",
  "Final Inspection"
];

const FINISH_OPERATOR_STATION_META = {
  "Finish Unbox": { zone: "MEI Feed", accent: "green", icon: "▣" },
  "MEI Line B": { zone: "Edging Bank", accent: "cyan", icon: "B" },
  "MEI Line C": { zone: "Edging Bank", accent: "purple", icon: "C" },
  "MEI Easy Fit": { zone: "Edging Bank", accent: "blue", icon: "EF" },
  "Mounting": { zone: "Mount / Assemble", accent: "orange", icon: "M" },
  "Drill": { zone: "Specialty", accent: "amber", icon: "D" },
  "Bigs": { zone: "Side Station", accent: "purple", icon: "BG" },
  "Sharps": { zone: "Side Station", accent: "purple", icon: "SH" },
  "Final Inspection": { zone: "Final QA", accent: "cyan", icon: "FI" }
};

function isFinishOperatorStationAllowed(stationName) {
  const name = normalizeFinishOperatorStation(stationName);

  return ![
    "FSV Scan & Verify",
    "FSV / Frame Only",
    "Frame Only",
    "FSV Scan & Verify / Frame Only",
    "AR-OUT"
  ].includes(name);
}

let finishOperatorState = {
  rows: [],
  stations: {},
  selectedStation: null,
  selectedOperator: "all",
  uiStatus: {}
};

function initOperatorCommandUI() {
  document.getElementById("operatorRefreshBtn")?.addEventListener("click", loadFinishOperatorActivity);
  hideOperatorApiConnectedPill();
  ensureExpandedOperatorViewer();

  document.querySelectorAll("[data-close-operator-drawer]").forEach(el => {
    el.addEventListener("click", closeOperatorDrawer);
  });

  document.querySelectorAll("[data-status-action]").forEach(btn => {
    btn.addEventListener("click", () => {
      const station = finishOperatorState.selectedStation;
      if (!station) return;

      const action = btn.dataset.statusAction || "clear";
      if (action === "clear") {
        delete finishOperatorState.uiStatus[station];
      } else {
        finishOperatorState.uiStatus[station] = action;
      }

      renderOperatorStationCards(finishOperatorState.stations);
      openOperatorDrawer(station);
    });
  });

  document.addEventListener("keydown", event => {
    if (event.key === "Escape") closeOperatorDrawer();
  });
}


function hideOperatorApiConnectedPill() {
  const apiStatus = document.getElementById("operatorApiStatus");
  if (!apiStatus) return;

  const apiPill = apiStatus.closest(".operator-api-pill, .hud-pill, .operator-status-pill");
  if (apiPill) {
    apiPill.style.display = "none";
  } else {
    apiStatus.style.display = "none";
  }
}

function getFinishOperatorBaseRate(stationName, operatorName = "") {
  const config = loadConfig();
  const name = normalizeFinishOperatorStation(stationName);
  const assignment = getOperatorAssignment(operatorName, name, config);

  if (assignment && assignment.role === "training") {
    const trainingRate = getTrainingRateForStationWeek(name, assignment.trainingWeek);
    if (trainingRate > 0) return trainingRate;
  }

  switch (name) {
    case "Finish Unbox":
      return Number(config.unboxRate || 0);
    case "MEI Line B":
      return Number(config.meiBRate || 0);
    case "MEI Line C":
      return Number(config.meiCRate || 0);
    case "MEI Easy Fit":
      return Number(config.meiBRate || 0);
    case "Mounting":
      return Number(config.mountRate || 0);
    case "Drill":
      return Number(config.drillRate || 0);
    case "Final Inspection":
      return Number(config.finalRate || 0);
    case "Bigs":
    case "Sharps":
      return 0;
    default:
      return 0;
  }
}

function getFinishOperatorStationTarget(stationName) {
  const station = applyConfigToStation({ flowStep: normalizeFinishOperatorStation(stationName) }, loadConfig());
  return Number(station.expectedNormalPerHour || 0);
}

function getOperatorHourTarget(stationName, hour, mode = "operator", operatorName = "") {
  const baseRate = mode === "station"
    ? getFinishOperatorStationTarget(stationName)
    : getFinishOperatorBaseRate(stationName, operatorName);

  if (baseRate <= 0) return 0;

  const productiveMinutes = getProductiveMinutesForHour(hour);
  if (productiveMinutes <= 0) return 0;

  return Math.round(baseRate * (productiveMinutes / 60));
}

function getOperatorPerformanceStatus(count, target) {
  const safeTarget = Number(target || 0);
  if (safeTarget <= 0) {
    return { pct: 0, className: "neutral", color: "#38bdf8", label: "No Target" };
  }

  const pct = Math.round((Number(count || 0) / safeTarget) * 100);

  if (pct >= 100) {
    return { pct, className: "green", color: "#00f5a0", label: `${pct}%` };
  }

  if (pct >= 90) {
    return { pct, className: "amber", color: "#fbbf24", label: `${pct}%` };
  }

  return { pct, className: "red", color: "#fb7185", label: `${pct}%` };
}

function getLatestOperatorPerformance(operator, stationName) {
  const hours = operator?.hourly || {};
  const orderedHours = getOperatorHourOrder(hours);
  let latestHour = "";
  let latestValue = 0;
  let latestTarget = 0;

  orderedHours.forEach(hour => {
    const target = getOperatorHourTarget(stationName, hour, "operator", operator?.name || "");
    const value = Number(hours[hour] || 0);

    if (target > 0 && value > 0) {
      latestHour = hour;
      latestValue = value;
      latestTarget = target;
    }
  });

  if (!latestHour) {
    const now = new Date();
    const currentHour = now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }).replace(/:\d{2}/, ":00");
    latestHour = orderedHours.includes(currentHour) ? currentHour : orderedHours[orderedHours.length - 1] || "";
    latestValue = Number(hours[latestHour] || 0);
    latestTarget = getOperatorHourTarget(stationName, latestHour, "operator", operator?.name || "");
  }

  const status = getOperatorPerformanceStatus(latestValue, latestTarget);
  return { hour: latestHour, value: latestValue, target: latestTarget, ...status };
}

function renderOperatorFilter(station, operators) {
  const target = document.getElementById("drawerOperatorList");
  if (!target) return;

  const existing = document.getElementById("drawerOperatorFilterWrap");
  if (existing) existing.remove();

  const wrap = document.createElement("div");
  wrap.id = "drawerOperatorFilterWrap";
  wrap.className = "operator-filter-wrap";
  wrap.style.cssText = "margin:0 0 16px;padding:14px 16px;border:1px solid rgba(100,221,255,.18);border-radius:16px;background:rgba(3,12,24,.72);display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;";

  const options = operators.map(operator =>
    `<option value="${escapeHtml(operator.name)}" ${finishOperatorState.selectedOperator === operator.name ? "selected" : ""}>${escapeHtml(operator.name)}</option>`
  ).join("");

  wrap.innerHTML = `
    <div>
      <strong style="display:block;font-size:13px;letter-spacing:.08em;text-transform:uppercase;color:#eef8ff;">Operator Filter</strong>
      <span style="display:block;margin-top:4px;font-size:11px;color:#8fa6c3;">Select one operator or view the full station crew.</span>
    </div>
    <select id="drawerOperatorFilter" style="min-width:260px;height:42px;border-radius:12px;border:1px solid rgba(0,217,255,.35);background:#061225;color:#eef8ff;padding:0 12px;font-weight:800;outline:none;">
      <option value="all" ${finishOperatorState.selectedOperator === "all" ? "selected" : ""}>All operators</option>
      ${options}
    </select>
  `;

  target.before(wrap);

  document.getElementById("drawerOperatorFilter")?.addEventListener("change", event => {
    finishOperatorState.selectedOperator = event.target.value || "all";
    renderDrawerOperators(station);
  });
}

async function loadFinishOperatorActivity() {
  const apiStatus = document.getElementById("operatorApiStatus");
  if (apiStatus) apiStatus.textContent = "Loading";

  try {
    const response = await fetch(FINISH_OPERATOR_API_URL, {
      method: "GET",
      cache: "no-store"
    });

    if (!response.ok) {
      throw new Error(`Operator API failed with status ${response.status}`);
    }

    const payload = await response.json();

    if (payload.status === "error" || payload.success === false) {
      throw new Error(payload.message || "Operator API returned an error.");
    }

    const rows = (Array.isArray(payload.operatorActivity)
      ? payload.operatorActivity
      : []
    ).filter(row => isFinishOperatorStationAllowed(
      row.FlowStation || row.flowStation || row.station || row.AccessPoint
    ));

    finishOperatorState.rows = rows;
    finishOperatorState.stations = buildFinishOperatorStations(rows);

    renderOperatorSummary(payload.summary || buildOperatorSummaryFromStations(finishOperatorState.stations));
    renderOperatorStationCards(finishOperatorState.stations);
    renderOperatorAssignmentConfigPanel();

    if (apiStatus) apiStatus.textContent = "Connected";
  } catch (error) {
    console.error("Finish Operator Activity Error:", error);
    if (apiStatus) apiStatus.textContent = "API Error";
    renderOperatorError(error);
  }
}

function buildFinishOperatorStations(rows) {
  const stations = {};

  rows.forEach(row => {
    const rawStation = row.FlowStation || row.flowStation || row.station || row.AccessPoint || "Unmapped";
    const stationName = normalizeFinishOperatorStation(rawStation);

    if (!stations[stationName]) {
      stations[stationName] = {
        name: stationName,
        total: 0,
        operators: {},
        hourly: {},
        accessPoints: new Set()
      };
    }

    const station = stations[stationName];
    const operator = String(row.Operator || row.operator || "Unassigned / No Operator").trim();
    const total = Number(row.Total ?? row.total ?? row.HourlyTotal ?? row.hourlyTotal ?? 0) || 0;
    const hours = row.Hours || row.hours || {};
    const accessPoint = String(row.AccessPoint || row.accessPoint || "").trim();

    station.total += total;
    if (accessPoint) station.accessPoints.add(accessPoint);

    if (!station.operators[operator]) {
      station.operators[operator] = {
        name: operator,
        total: 0,
        hourly: {},
        accessPoints: new Set()
      };
    }

    station.operators[operator].total += total;
    if (accessPoint) station.operators[operator].accessPoints.add(accessPoint);

    getOperatorHourOrder(hours).forEach(hour => {
      const value = Number(hours[hour] || 0) || 0;
      station.hourly[hour] = (station.hourly[hour] || 0) + value;
      station.operators[operator].hourly[hour] = (station.operators[operator].hourly[hour] || 0) + value;
    });
  });

  Object.values(stations).forEach(station => {
    station.accessPoints = Array.from(station.accessPoints);
    station.operatorList = Object.values(station.operators)
      .map(operator => ({
        ...operator,
        accessPoints: Array.from(operator.accessPoints)
      }))
      .sort((a, b) => Number(b.total || 0) - Number(a.total || 0));
  });

  return stations;
}

function normalizeFinishOperatorStation(value) {
  const text = String(value || "").trim();
  const key = text.toLowerCase();

  if (key.includes("unbox")) return "Finish Unbox";
  if (key.includes("line b")) return "MEI Line B";
  if (key.includes("line c")) return "MEI Line C";
  if (key.includes("easy")) return "MEI Easy Fit";
  if (key.includes("mount")) return "Mounting";
  if (key.includes("drill")) return "Drill";
  if (key.includes("big")) return "Bigs";
  if (key.includes("sharp")) return "Sharps";
  if (key.includes("final")) return "Final Inspection";
  if (key.includes("fsv") || key.includes("frame only")) return "FSV Scan & Verify";
  if (key.includes("ar-out") || key.includes("ar out")) return "AR-OUT";

  return text || "Unmapped";
}

function renderOperatorSummary(summary) {
  setText("operatorTotalJobs", summary.totalJobs || 0);
  setText("operatorTotalOperators", summary.totalOperators || 0);
  setText("operatorTopStationTotal", summary.topStationTotal || 0);
  setText("operatorTopStation", summary.topStation || "--");
  setText("operatorPeakHourTotal", summary.peakHourTotal || 0);
  setText("operatorPeakHour", summary.peakHour || "--");
}

function buildOperatorSummaryFromStations(stations) {
  const allStations = Object.values(stations || {});
  const operatorSet = {};
  let totalJobs = 0;
  let topStation = "";
  let topStationTotal = 0;
  let peakHour = "";
  let peakHourTotal = 0;
  const hourTotals = {};

  allStations.forEach(station => {
    totalJobs += Number(station.total || 0);

    if (Number(station.total || 0) > topStationTotal) {
      topStationTotal = Number(station.total || 0);
      topStation = station.name;
    }

    (station.operatorList || []).forEach(operator => {
      operatorSet[operator.name] = true;
    });

    Object.keys(station.hourly || {}).forEach(hour => {
      hourTotals[hour] = (hourTotals[hour] || 0) + Number(station.hourly[hour] || 0);
    });
  });

  Object.keys(hourTotals).forEach(hour => {
    if (hourTotals[hour] > peakHourTotal) {
      peakHourTotal = hourTotals[hour];
      peakHour = hour;
    }
  });

  return {
    totalJobs,
    totalOperators: Object.keys(operatorSet).length,
    topStation,
    topStationTotal,
    peakHour,
    peakHourTotal
  };
}

function renderOperatorStationCards(stations) {
  const grid = document.getElementById("operatorStationGrid");
  if (!grid) return;

  const ordered = getOrderedFinishOperatorStations(stations);

  if (!ordered.length) {
    grid.innerHTML = `<article class="operator-empty-card">No Finish operator rows returned yet. Confirm RAW_ACTIVITY_CURRENT has Area = Finish rows.</article>`;
    return;
  }

  grid.innerHTML = ordered.map(station => {
    const meta = FINISH_OPERATOR_STATION_META[station.name] || { zone: "Finish", accent: "cyan", icon: "•" };
    const topOperator = (station.operatorList || [])[0];
    const peak = getPeakHour(station.hourly);
    const uiStatus = finishOperatorState.uiStatus[station.name] || "online";

    return `
      <article class="operator-station-card ${escapeHtml(meta.accent)} ui-${escapeHtml(uiStatus)}" data-operator-station="${escapeHtml(station.name)}">
        <div class="operator-card-top">
          <div class="operator-station-icon">${escapeHtml(meta.icon)}</div>
          <div>
            <span>${escapeHtml(meta.zone)}</span>
            <h3>${escapeHtml(station.name)}</h3>
          </div>
          <strong class="operator-ui-status">${escapeHtml(formatUiStatus(uiStatus))}</strong>
        </div>

        <div class="operator-card-main-metric">
          <span>Output Today</span>
          <strong>${numberFmt(station.total)}</strong>
        </div>

        <div class="operator-card-metrics">
          <div><span>Operators</span><strong>${numberFmt((station.operatorList || []).length)}</strong></div>
          <div><span>Peak</span><strong>${escapeHtml(peak.hour || "--")}</strong></div>
          <div><span>Peak CNT</span><strong>${numberFmt(peak.value)}</strong></div>
        </div>

        <div class="operator-card-footer">
          <span>Top: ${escapeHtml(topOperator?.name || "--")}</span>
          <button type="button">Open Detail</button>
        </div>
      </article>`;
  }).join("");

  grid.querySelectorAll("[data-operator-station]").forEach(card => {
    card.addEventListener("click", () => openOperatorDrawer(card.dataset.operatorStation));
  });
}

function getOrderedFinishOperatorStations(stations) {
  const stationMap = stations || {};
  const ordered = [];

  FINISH_OPERATOR_STATION_ORDER.forEach(name => {
    if (stationMap[name]) ordered.push(stationMap[name]);
  });

  Object.keys(stationMap)
    .filter(name =>
      !FINISH_OPERATOR_STATION_ORDER.includes(name) &&
      isFinishOperatorStationAllowed(name)
    )
    .sort()
    .forEach(name => ordered.push(stationMap[name]));

  return ordered;
}

function openOperatorDrawer(stationName) {
  const station = finishOperatorState.stations?.[stationName];
  const drawer = document.getElementById("operatorDrawer");
  if (!station || !drawer) return;

  finishOperatorState.selectedStation = stationName;

  const meta = FINISH_OPERATOR_STATION_META[station.name] || { zone: "Finish" };
  const peak = getPeakHour(station.hourly);

  setText("operatorDrawerTitle", station.name);
  setText("operatorDrawerSub", `${meta.zone} • ${station.accessPoints?.join(" + ") || "Mapped station"}`);
  setText("drawerStationTotal", station.total || 0);
  setText("drawerOperatorCount", (station.operatorList || []).length);
  setText("drawerPeakHour", peak.hour ? `${peak.hour} (${peak.value})` : "--");

  if (!finishOperatorState.selectedOperator) finishOperatorState.selectedOperator = "all";
  const operatorNames = (station.operatorList || []).map(operator => operator.name);
  if (finishOperatorState.selectedOperator !== "all" && !operatorNames.includes(finishOperatorState.selectedOperator)) {
    finishOperatorState.selectedOperator = "all";
  }

  renderDrawerStationHourly(station);
  renderDrawerOperators(station);
  updateDrawerStatusButtons(finishOperatorState.uiStatus[station.name] || "online");

  drawer.classList.add("active");
  drawer.setAttribute("aria-hidden", "false");
  document.body.classList.add("operator-drawer-open");
}

function closeOperatorDrawer() {
  const drawer = document.getElementById("operatorDrawer");
  if (!drawer) return;

  drawer.classList.remove("active");
  drawer.setAttribute("aria-hidden", "true");
  document.body.classList.remove("operator-drawer-open");
}

function renderDrawerStationHourly(station) {
  /*
    Station hourly output is intentionally hidden in the Operator Command drawer.
    That full station-by-hour view already exists in Hourly Production Performance.
    The drawer stays focused on individual operator hourly output only.
  */
  const target = document.getElementById("drawerStationHourly");
  if (!target) return;

  const block = target.closest(".operator-drawer-block");
  if (block) block.style.display = "none";

  target.innerHTML = "";
}

function renderDrawerOperators(station) {
  const target = document.getElementById("drawerOperatorList");
  if (!target) return;

  const operators = station.operatorList || [];

  if (!operators.length) {
    const existing = document.getElementById("drawerOperatorFilterWrap");
    if (existing) existing.remove();
    target.innerHTML = `<article class="operator-empty-card">No individual operators returned for this station.</article>`;
    return;
  }

  renderOperatorFilter(station, operators);

  const filteredOperators = finishOperatorState.selectedOperator === "all"
    ? operators
    : operators.filter(operator => operator.name === finishOperatorState.selectedOperator);

  target.innerHTML = filteredOperators.map(operator => {
    const hours = operator.hourly || {};
    const currentPerf = getLatestOperatorPerformance(operator, station.name);

    return `
      <article class="operator-person-card perf-${escapeHtml(currentPerf.className)} operator-person-clickable" data-expanded-operator="${escapeHtml(operator.name)}" style="border-color:${currentPerf.color};box-shadow:0 0 18px ${currentPerf.color}1f;cursor:pointer;">
        <div class="operator-person-head">
          <div>
            <h4>${escapeHtml(operator.name)}</h4>
            <span>${escapeHtml(operator.accessPoints?.join(" + ") || station.name)} · ${escapeHtml(getAssignmentTargetLabel(operator.name, station.name))}</span>
            <small style="display:inline-flex;margin-top:8px;padding:5px 9px;border-radius:999px;border:1px solid ${currentPerf.color};color:${currentPerf.color};font-weight:900;letter-spacing:.06em;text-transform:uppercase;">
              Current JPH ${escapeHtml(currentPerf.label)} · ${escapeHtml(shortHour(currentPerf.hour)) || "--"}: ${numberFmt(currentPerf.value)} / ${numberFmt(currentPerf.target)}
            </small>
          </div>
          <div style="display:flex;flex-direction:column;align-items:flex-end;gap:8px;">
            <strong style="color:${currentPerf.color};">${numberFmt(operator.total)}</strong>
            <button type="button" class="operator-expand-btn" data-expanded-operator-btn="${escapeHtml(operator.name)}">Expand</button>
          </div>
        </div>

        <div class="operator-person-hours">
          ${getOperatorHourOrder(hours).map(hour => {
            const value = Number(hours[hour] || 0);
            const targetCount = getOperatorHourTarget(station.name, hour, "operator", operator.name);
            const status = getOperatorPerformanceStatus(value, targetCount);
            const pct = targetCount > 0
              ? Math.min(100, Math.round((value / targetCount) * 100))
              : 0;

            return `
              <div class="operator-mini-hour perf-${escapeHtml(status.className)}" title="${escapeHtml(operator.name)} · ${escapeHtml(hour)} · ${escapeHtml(status.label)}">
                <span>${escapeHtml(shortHour(hour))}</span>
                <b style="color:${status.color};">${numberFmt(value)}</b>
                <small style="color:${status.color};font-size:10px;font-weight:900;">${escapeHtml(status.label)}</small>
                <i style="height:${Math.max(4, pct)}%;background:${status.color};box-shadow:0 0 14px ${status.color};"></i>
              </div>`;
          }).join("")}
        </div>
      </article>`;
  }).join("");

  target.querySelectorAll("[data-expanded-operator]").forEach(card => {
    card.addEventListener("click", event => {
      event.stopPropagation();
      const operatorName = card.dataset.expandedOperator;
      openExpandedOperatorViewer(station.name, operatorName);
    });
  });

  target.querySelectorAll("[data-expanded-operator-btn]").forEach(btn => {
    btn.addEventListener("click", event => {
      event.preventDefault();
      event.stopPropagation();
      openExpandedOperatorViewer(station.name, btn.dataset.expandedOperatorBtn);
    });
  });
}

function ensureExpandedOperatorViewer() {
  if (document.getElementById("expandedOperatorViewer")) return;

  const style = document.createElement("style");
  style.id = "expandedOperatorViewerStyles";
  style.textContent = `
    .operator-expand-btn {
      height: 30px;
      padding: 0 12px;
      border-radius: 999px;
      border: 1px solid rgba(0, 217, 255, 0.42);
      background: rgba(0, 217, 255, 0.08);
      color: #67e8f9;
      font-family: "JetBrains Mono", monospace;
      font-size: 10px;
      font-weight: 900;
      letter-spacing: .08em;
      text-transform: uppercase;
      cursor: pointer;
    }

    .operator-person-clickable {
      transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease;
    }

    .operator-person-clickable:hover {
      transform: translateY(-2px);
      box-shadow: 0 0 28px rgba(0, 217, 255, .18), inset 0 1px 0 rgba(255,255,255,.08) !important;
    }

    .expanded-operator-backdrop {
      position: fixed;
      inset: 0;
      z-index: 9998;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 26px;
      background: rgba(0, 4, 10, .72);
      backdrop-filter: blur(12px);
    }

    .expanded-operator-backdrop.active {
      display: flex;
    }

    .expanded-operator-shell {
      width: min(1220px, 96vw);
      max-height: 92vh;
      overflow: hidden;
      border: 1px solid rgba(100, 221, 255, .24);
      border-radius: 26px;
      background:
        radial-gradient(circle at 8% 0%, rgba(0, 217, 255, .14), transparent 32%),
        radial-gradient(circle at 92% 8%, rgba(0, 245, 160, .10), transparent 34%),
        linear-gradient(180deg, rgba(10, 28, 52, .98), rgba(2, 7, 17, .99));
      box-shadow: 0 34px 90px rgba(0,0,0,.58), 0 0 38px rgba(0,217,255,.14);
    }

    .expanded-operator-header {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 18px;
      align-items: start;
      padding: 24px 26px 20px;
      border-bottom: 1px solid rgba(100, 221, 255, .18);
      background: linear-gradient(90deg, rgba(0, 217, 255, .08), transparent);
    }

    .expanded-operator-kicker {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      color: #5eead4;
      font-family: "JetBrains Mono", monospace;
      font-size: 11px;
      font-weight: 900;
      letter-spacing: .16em;
      text-transform: uppercase;
    }

    .expanded-operator-kicker::before {
      content: "";
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: #00f5a0;
      box-shadow: 0 0 14px #00f5a0;
    }

    .expanded-operator-title {
      margin: 10px 0 0;
      font-size: clamp(30px, 4vw, 54px);
      line-height: .95;
      font-weight: 900;
      letter-spacing: -.04em;
    }

    .expanded-operator-sub {
      margin-top: 8px;
      color: #9fb3cb;
      font-size: 14px;
      font-weight: 700;
    }

    .expanded-operator-close {
      width: 46px;
      height: 46px;
      display: grid;
      place-items: center;
      border-radius: 16px;
      border: 1px solid rgba(100, 221, 255, .22);
      background: rgba(3, 12, 24, .8);
      color: #eef8ff;
      font-size: 24px;
      cursor: pointer;
    }

    .expanded-operator-body {
      max-height: calc(92vh - 132px);
      overflow: auto;
      padding: 22px 26px 26px;
    }

    .expanded-operator-scoreboard {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 20px;
    }

    .expanded-score-card {
      min-height: 112px;
      padding: 18px;
      border-radius: 20px;
      border: 1px solid rgba(100, 221, 255, .18);
      background: rgba(3, 12, 24, .72);
      box-shadow: inset 0 1px 0 rgba(255,255,255,.06);
    }

    .expanded-score-card span {
      display: block;
      color: #8fa6c3;
      font-family: "JetBrains Mono", monospace;
      font-size: 10px;
      font-weight: 900;
      letter-spacing: .14em;
      text-transform: uppercase;
    }

    .expanded-score-card strong {
      display: block;
      margin-top: 10px;
      color: #eef8ff;
      font-size: 34px;
      line-height: 1;
      font-weight: 900;
    }

    .expanded-score-card small {
      display: block;
      margin-top: 8px;
      color: #9fb3cb;
      font-weight: 800;
    }

    .expanded-operator-timeline {
      display: grid;
      grid-template-columns: repeat(15, minmax(64px, 1fr));
      gap: 10px;
      align-items: end;
      min-height: 360px;
      padding: 22px;
      border-radius: 24px;
      border: 1px solid rgba(100, 221, 255, .18);
      background:
        linear-gradient(180deg, rgba(255,255,255,.045), transparent),
        rgba(3, 12, 24, .72);
      overflow-x: auto;
    }

    .expanded-hour-tower {
      min-width: 64px;
      height: 300px;
      display: grid;
      grid-template-rows: auto 1fr auto;
      gap: 8px;
      text-align: center;
    }

    .expanded-hour-top span,
    .expanded-hour-foot span {
      display: block;
      color: #8fa6c3;
      font-family: "JetBrains Mono", monospace;
      font-size: 10px;
      font-weight: 900;
    }

    .expanded-hour-top strong {
      display: block;
      margin-top: 5px;
      font-size: 22px;
      font-weight: 900;
    }

    .expanded-hour-bar-wrap {
      position: relative;
      align-self: stretch;
      min-height: 190px;
      border-radius: 16px;
      border: 1px solid rgba(100, 221, 255, .10);
      background: rgba(255,255,255,.035);
      overflow: hidden;
    }

    .expanded-hour-target-line {
      position: absolute;
      left: 8px;
      right: 8px;
      bottom: var(--target-line, 70%);
      height: 2px;
      background: rgba(255,255,255,.72);
      box-shadow: 0 0 12px rgba(255,255,255,.38);
      z-index: 2;
    }

    .expanded-hour-fill {
      position: absolute;
      left: 9px;
      right: 9px;
      bottom: 8px;
      height: var(--bar-height, 4%);
      min-height: 6px;
      border-radius: 12px 12px 4px 4px;
      background: var(--bar-color, #38bdf8);
      box-shadow: 0 0 18px var(--bar-color, #38bdf8);
    }

    .expanded-hour-foot strong {
      display: block;
      margin-top: 4px;
      font-size: 12px;
      font-weight: 900;
    }

    .expanded-operator-note {
      margin-top: 16px;
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid rgba(100, 221, 255, .14);
      background: rgba(255,255,255,.035);
      color: #bfdaf2;
      font-size: 13px;
      line-height: 1.5;
    }

    @media (max-width: 900px) {
      .expanded-operator-scoreboard { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .expanded-operator-header { padding: 20px; }
      .expanded-operator-body { padding: 18px; }
    }
  `;
  document.head.appendChild(style);

  const viewer = document.createElement("section");
  viewer.id = "expandedOperatorViewer";
  viewer.className = "expanded-operator-backdrop";
  viewer.setAttribute("aria-hidden", "true");
  viewer.innerHTML = `
    <div class="expanded-operator-shell" role="dialog" aria-modal="true" aria-label="Expanded operator performance">
      <header class="expanded-operator-header">
        <div>
          <div class="expanded-operator-kicker">Operator Drilldown</div>
          <h2 class="expanded-operator-title" id="expandedOperatorName">--</h2>
          <div class="expanded-operator-sub" id="expandedOperatorSub">--</div>
        </div>
        <button type="button" class="expanded-operator-close" data-close-expanded-operator aria-label="Close expanded operator view">×</button>
      </header>
      <div class="expanded-operator-body" id="expandedOperatorBody"></div>
    </div>
  `;

  document.body.appendChild(viewer);

  viewer.addEventListener("click", event => {
    if (event.target === viewer || event.target.closest("[data-close-expanded-operator]")) {
      closeExpandedOperatorViewer();
    }
  });

  document.addEventListener("keydown", event => {
    if (event.key === "Escape") closeExpandedOperatorViewer();
  });
}

function openExpandedOperatorViewer(stationName, operatorName) {
  ensureExpandedOperatorViewer();

  const station = finishOperatorState.stations?.[stationName];
  const operator = (station?.operatorList || []).find(item => item.name === operatorName);
  const viewer = document.getElementById("expandedOperatorViewer");
  const body = document.getElementById("expandedOperatorBody");

  if (!station || !operator || !viewer || !body) return;

  const currentPerf = getLatestOperatorPerformance(operator, station.name);
  const hours = operator.hourly || {};
  const orderedHours = getOperatorHourOrder(hours);
  const meta = FINISH_OPERATOR_STATION_META[station.name] || { zone: "Finish" };
  const best = getPeakHour(hours);
  const activeHours = orderedHours.filter(hour => Number(hours[hour] || 0) > 0).length;
  const avgPerActiveHour = activeHours > 0 ? Math.round(Number(operator.total || 0) / activeHours) : 0;

  setText("expandedOperatorName", operator.name);
  setText("expandedOperatorSub", `${station.name} • ${meta.zone} • ${getAssignmentTargetLabel(operator.name, station.name)} • ${operator.accessPoints?.join(" + ") || "Mapped station"}`);

  body.innerHTML = `
    <section class="expanded-operator-scoreboard">
      <article class="expanded-score-card">
        <span>Total Output</span>
        <strong>${numberFmt(operator.total)}</strong>
        <small>Jobs scanned today</small>
      </article>
      <article class="expanded-score-card">
        <span>Current JPH</span>
        <strong style="color:${currentPerf.color};">${escapeHtml(currentPerf.label)}</strong>
        <small>${escapeHtml(shortHour(currentPerf.hour)) || "--"}: ${numberFmt(currentPerf.value)} / ${numberFmt(currentPerf.target)}</small>
      </article>
      <article class="expanded-score-card">
        <span>Peak Hour</span>
        <strong>${escapeHtml(shortHour(best.hour) || "--")}</strong>
        <small>${numberFmt(best.value)} jobs</small>
      </article>
      <article class="expanded-score-card">
        <span>Active Hour Avg</span>
        <strong>${numberFmt(avgPerActiveHour)}</strong>
        <small>${numberFmt(activeHours)} active hours</small>
      </article>
    </section>

    <section class="expanded-operator-timeline">
      ${orderedHours.map(hour => {
        const value = Number(hours[hour] || 0);
        const target = getOperatorHourTarget(station.name, hour, "operator", operator.name);
        const status = getOperatorPerformanceStatus(value, target);
        const pctRaw = target > 0 ? Math.round((value / target) * 100) : 0;
        const barPct = target > 0 ? Math.min(100, Math.max(4, pctRaw)) : (value > 0 ? 18 : 4);
        const targetLine = target > 0 ? 70 : 4;

        return `
          <article class="expanded-hour-tower perf-${escapeHtml(status.className)}">
            <div class="expanded-hour-top">
              <span>${escapeHtml(shortHour(hour))}</span>
              <strong style="color:${status.color};">${numberFmt(value)}</strong>
            </div>
            <div class="expanded-hour-bar-wrap" style="--target-line:${targetLine}%;">
              <div class="expanded-hour-target-line"></div>
              <div class="expanded-hour-fill" style="--bar-height:${barPct}%;--bar-color:${status.color};"></div>
            </div>
            <div class="expanded-hour-foot">
              <span>${target > 0 ? numberFmt(target) + " target" : "No target"}</span>
              <strong style="color:${status.color};">${escapeHtml(status.label)}</strong>
            </div>
          </article>`;
      }).join("")}
    </section>

    <div class="expanded-operator-note">
      Green means the operator met or beat the configured JPH target. Amber means 90% to 99%. Red means below 90%. The white marker represents the target line for that hour after shift breaks are applied.
    </div>
  `;

  viewer.classList.add("active");
  viewer.setAttribute("aria-hidden", "false");
}

function closeExpandedOperatorViewer() {
  const viewer = document.getElementById("expandedOperatorViewer");
  if (!viewer) return;

  viewer.classList.remove("active");
  viewer.setAttribute("aria-hidden", "true");
}

function updateDrawerStatusButtons(activeStatus) {
  document.querySelectorAll("[data-status-action]").forEach(btn => {
    const action = btn.dataset.statusAction || "clear";
    btn.classList.toggle("active", action === activeStatus || (action === "clear" && !activeStatus));
  });
}

function getPeakHour(hours) {
  let bestHour = "";
  let bestValue = 0;

  Object.keys(hours || {}).forEach(hour => {
    const value = Number(hours[hour] || 0);
    if (value > bestValue) {
      bestValue = value;
      bestHour = hour;
    }
  });

  return { hour: bestHour, value: bestValue };
}

function getOperatorHourOrder(hours) {
  const preferred = [
    "6:00 AM", "7:00 AM", "8:00 AM", "9:00 AM", "10:00 AM", "11:00 AM",
    "12:00 PM", "1:00 PM", "2:00 PM", "3:00 PM", "4:00 PM", "5:00 PM",
    "6:00 PM", "7:00 PM", "8:00 PM"
  ];

  const keys = Object.keys(hours || {});
  const ordered = preferred.filter(hour => keys.includes(hour));
  const leftovers = keys.filter(hour => !preferred.includes(hour)).sort();

  return [...ordered, ...leftovers];
}

function shortHour(hour) {
  return String(hour || "")
    .replace(":00", "")
    .replace(" AM", "A")
    .replace(" PM", "P");
}

function formatUiStatus(status) {
  if (status === "down") return "Down";
  if (status === "issue") return "Issue";
  return "Online";
}

function numberFmt(value) {
  return Number(value || 0).toLocaleString();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function renderOperatorError(error) {
  const grid = document.getElementById("operatorStationGrid");
  if (!grid) return;

  grid.innerHTML = `
    <article class="operator-empty-card operator-error-card">
      Finish Operator API could not load.<br />
      <small>${escapeHtml(error.message || error)}</small>
    </article>`;
}
