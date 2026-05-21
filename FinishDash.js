const FINISH_API_URL =
  "https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec?area=Finish";

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
finalTraineeW5: 0

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
  switch (station.flowStep) {
    case "Finish Unbox":
      station.expectedNormalPerHour = config.unboxRate * config.unboxCount;
      break;

    case "MEI Line A":
      station.expectedNormalPerHour = config.meiARate * config.meiACount;
      break;

    case "MEI Line B":
      station.expectedNormalPerHour = config.meiBRate * config.meiBCount;
      break;

    case "MEI Line C":
      station.expectedNormalPerHour = config.meiCRate * config.meiCCount;
      break;

    case "MEI Easy Fit":
      station.expectedNormalPerHour = config.meiBRate * config.meiBCount;
      break;

   case "Mounting": {
  const mountTotal =
    Number(config.mountRate || 0) *
    Number(config.mountCount || 0);

  const trainingMountTotal =
    (Number(config.mountTraineeW1 || 0) * 3) +
    (Number(config.mountTraineeW2 || 0) * 6) +
    (Number(config.mountTraineeW3 || 0) * 9) +
    (Number(config.mountTraineeW4 || 0) * 12) +
    (Number(config.mountTraineeW5 || 0) * 15) +
    (Number(config.mountTraineeW6 || 0) * 18) +
    (Number(config.mountTraineeW7 || 0) * 21) +
    (Number(config.mountTraineeW8 || 0) * 25);

  station.expectedNormalPerHour = mountTotal + trainingMountTotal;
  break;
}

    case "Drill":
      station.expectedNormalPerHour = config.drillRate * config.drillCount;
      break;

    case "Final Inspection": {
  const finalCertifiedTotal =
    Number(config.finalRate || 0) *
    Number(config.finalCount || 0);

  const finalTrainingTotal =
    (Number(config.finalTraineeW1 || 0) * 15) +
    (Number(config.finalTraineeW2 || 0) * 30) +
    (Number(config.finalTraineeW3 || 0) * 45) +
    (Number(config.finalTraineeW4 || 0) * 60) +
    (Number(config.finalTraineeW5 || 0) * 75);

  station.expectedNormalPerHour = finalCertifiedTotal + finalTrainingTotal;
  break;
}

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
finalTraineeW5: getInputValue("cfgFinalTraineeW5")

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
  const unboxTotal = getInputValue("cfgUnboxRate") * getInputValue("cfgUnboxCount");

  const meiATotal = getInputValue("cfgMeiARate") * getInputValue("cfgMeiACount");
  const meiBTotal = getInputValue("cfgMeiBRate") * getInputValue("cfgMeiBCount");
  const meiCTotal = getInputValue("cfgMeiCRate") * getInputValue("cfgMeiCCount");

  const mountTotal =
  getInputValue("cfgMountRate") *
  getInputValue("cfgMountCount");

const trainingMountTotal =
  (getInputValue("cfgMountTraineeW1") * 3) +
  (getInputValue("cfgMountTraineeW2") * 6) +
  (getInputValue("cfgMountTraineeW3") * 9) +
  (getInputValue("cfgMountTraineeW4") * 12) +
  (getInputValue("cfgMountTraineeW5") * 15) +
  (getInputValue("cfgMountTraineeW6") * 18) +
  (getInputValue("cfgMountTraineeW7") * 21) +
  (getInputValue("cfgMountTraineeW8") * 25);

  const drillTotal = getInputValue("cfgDrillRate") * getInputValue("cfgDrillCount");

  const finalTotal =
  getInputValue("cfgFinalRate") *
  getInputValue("cfgFinalCount");

const finalTrainingTotal =
  (getInputValue("cfgFinalTraineeW1") * 15) +
  (getInputValue("cfgFinalTraineeW2") * 30) +
  (getInputValue("cfgFinalTraineeW3") * 45) +
  (getInputValue("cfgFinalTraineeW4") * 60) +
  (getInputValue("cfgFinalTraineeW5") * 75);

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