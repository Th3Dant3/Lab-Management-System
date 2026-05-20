/* =========================================================
   SurfaceWIP.js
   Production Surface WIP Dashboard
   Live Flow = Total Scan Today → Current WIP Up Now → Next Station
========================================================= */
"use strict";

let hourlyInOutChart = null;


/* =========================================================
   SECTION 01 — CONFIG
========================================================= */

const CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec",
  AREA: "Surface",
  REFRESH_MS: 30_000,

  WIP_CRITICAL: 200,
  WIP_HIGH: 100,
  WIP_ACTIVE: 25,

  LINE_A_OFFLINE: true,

  LINE_A_ONLY_STEPS: [
    "Blocking Line A",
    "Generating Line A",
    "Polishing Line A",
    "Engraving Line A"
  ],

  LINE_B_ONLY_STEPS: [
    "Blocking Line B",
    "Generating Line B",
    "Polishing Line B",
    "Engraving Line B",
    "Detaping Line B",
    "Coating Line B"
  ]
};


/* =========================================================
   SECTION 02 — GLOBAL STATE
========================================================= */

let state = {
  summary: null,
  surfaceFlow: [],
  surfaceTransfers: [],
  surfaceTransferDailyTotals: [],
  surfaceScanSummary: [],
  surfaceHourlyInOut: [],
  intervalMeta: null,

  lastFetch: null,
  lastSnapshotKey: null,
  lastDataSignature: null,

  lastWipSignature: null,
  lastActivitySignature: null,
  lastUpdateSource: "—",
  lastUpdateTime: null,

  activeComparisonIndex: 0,

  filter: "all",
  activeTab: "overview",
  hasRenderedOnce: false
};


/* =========================================================
   SECTION 03 — SVG ICONS
========================================================= */

const ICONS = {
  scan: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M4 7V4h3M17 4h3v3M20 17v3h-3M7 20H4v-3"/><path d="M8 8h3v3H8zM13 8h3M13 11h3M8 13h3M13 13h3v3h-3zM8 16h3"/></svg>`,

  box: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="m12 2 9 5-9 5-9-5 9-5Z"/><path d="M3 7v10l9 5 9-5V7"/><path d="M12 12v10"/></svg>`,

  blocking: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><rect x="6" y="4" width="12" height="16" rx="2"/><path d="M9 8h6M8 20v2M16 20v2"/></svg>`,

  snow: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 2v20M4.9 4.9l14.2 14.2M2 12h20M4.9 19.1 19.1 4.9"/><path d="m8 4 4 4 4-4M8 20l4-4 4 4M4 8l4 4-4 4M20 8l-4 4 4 4"/></svg>`,

  gear: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.2a1.7 1.7 0 0 0-1-1.5 1.7 1.7 0 0 0-1.8.3l-.1-.1a2 2 0 1 1-2.8-2.8l.1-.1A1.7 1.7 0 0 0 4.6 15a1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.2a1.7 1.7 0 0 0 1.5-1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1A1.7 1.7 0 0 0 9 4.6a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.2a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1A1.7 1.7 0 0 0 19.4 9c.2.6.8 1 1.5 1h.1a2 2 0 1 1 0 4h-.2c-.6 0-1.2.4-1.4 1Z"/></svg>`,

  sparkle: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="m12 2 2.2 6.8L21 11l-6.8 2.2L12 20l-2.2-6.8L3 11l6.8-2.2L12 2Z"/></svg>`,

  engrave: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 2v12"/><path d="m8 10 4 4 4-4"/><path d="M5 22h14M9 18h6M12 14v4"/></svg>`,

  tape: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><circle cx="9" cy="12" r="7"/><circle cx="9" cy="12" r="3"/><path d="M9 19h11l-3-4h-2"/></svg>`,

  droplet: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 2S5 10 5 15a7 7 0 0 0 14 0c0-5-7-13-7-13Z"/></svg>`,

  shield: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 2 20 5v6c0 5-3.4 9.3-8 11-4.6-1.7-8-6-8-11V5l8-3Z"/><path d="m8 12 3 3 5-6"/></svg>`,

  gearDefault: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"><circle cx="12" cy="12" r="3"/><path d="M12 2v4M12 18v4M4.9 4.9l2.8 2.8M16.3 16.3l2.8 2.8M2 12h4M18 12h4M4.9 19.1l2.8-2.8M16.3 7.7l2.8-2.8"/></svg>`
};

function iconForStep(step) {
  const s = String(step || "").toLowerCase();

  if (s.includes("scan")) return ICONS.scan;
  if (s.includes("unbox")) return ICONS.box;
  if (s.includes("blocking")) return ICONS.blocking;
  if (s.includes("cooling")) return ICONS.snow;
  if (s.includes("generating")) return ICONS.gear;
  if (s.includes("polish")) return ICONS.sparkle;
  if (s.includes("engrav")) return ICONS.engrave;
  if (s.includes("detap")) return ICONS.tape;
  if (s.includes("coating")) return ICONS.droplet;
  if (s.includes("inspect")) return ICONS.shield;

  return ICONS.gearDefault;
}


/* =========================================================
   SECTION 04 — BASIC HELPERS
========================================================= */

function num(value) {
  const n = Number(String(value ?? "").replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
}

function pct(value) {
  const n = parseFloat(value);
  return isNaN(n) ? 0 : Math.round(n * 100);
}

function safeText(value, fallback = "—") {
  const text = String(value ?? "").trim();
  return text || fallback;
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function isLineA(step) {
  return CONFIG.LINE_A_ONLY_STEPS.includes(String(step || ""));
}

function isLineB(step) {
  return CONFIG.LINE_B_ONLY_STEPS.includes(String(step || ""));
}

function isShared(step) {
  return !isLineA(step) && !isLineB(step);
}

function statusFromBackend(row) {
  const wip = num(row.CurrentJobTotal);

  if (wip >= CONFIG.WIP_CRITICAL) return "Critical";
  if (wip >= CONFIG.WIP_HIGH) return "Watch";
  if (wip >= CONFIG.WIP_ACTIVE) return "ACTIVE";
  if (wip > 0) return "LOW";
  return "EMPTY";
}

function stageClass(status, flowStep) {
  if (CONFIG.LINE_A_OFFLINE && isLineA(flowStep)) return "offline";
  if (status === "Critical") return "critical";
  if (status === "Watch") return "watch";
  return "";
}

function statusPillText(status, flowStep) {
  if (CONFIG.LINE_A_OFFLINE && isLineA(flowStep)) return "OFFLINE";
  return status;
}

function barWidth(wip, maxWip) {
  if (!maxWip) return 0;
  return Math.min(100, Math.round((wip / maxWip) * 100));
}

function filterRows(rows, filter) {
  if (!rows) return [];

  switch (filter) {
    case "lineA":
      return rows.filter(r => isShared(r.FlowStep) || isLineA(r.FlowStep));

    case "lineB":
      return rows.filter(r => isShared(r.FlowStep) || isLineB(r.FlowStep));

    case "critical":
      return rows.filter(r => {
        const status = statusFromBackend(r);
        return status === "Critical" || status === "Watch";
      });

    default:
      return rows;
  }
}

/* =========================================================
   SECTION 04B — FORMAT HELPERS
   Used by Hourly In / Out chart and table
========================================================= */

function formatSignedNumber(value) {
  const n = num(value);

  if (n > 0) return "+" + n.toLocaleString();
  if (n < 0) return n.toLocaleString();

  return "0";
}

function formatHourLabel(hour) {
  const h = num(hour);

  if (h === 0) return "12 AM";
  if (h < 12) return h + " AM";
  if (h === 12) return "12 PM";

  return (h - 12) + " PM";
}

/* =========================================================
   SECTION — UPDATE SOURCE HELPERS
========================================================= */

function buildWipSignature(payload) {
  const summary = payload.summary || {};
  const flow = Array.isArray(payload.surfaceFlow) ? payload.surfaceFlow : [];

  return [
    summary.ReportDate || "",
    summary.SourceFile || "",
    summary.FileUpdatedTime || "",
    summary.ImportedAt || "",
    flow.length,
    flow.map(row => [
      row.FlowStep || "",
      row.CurrentJobTotal || "",
      row.PercentOfSurfaceTotal || ""
    ].join(":")).join("|")
  ].join("||");
}

function buildActivitySignature(payload) {
  const scanRows = Array.isArray(payload.surfaceScanSummary)
    ? payload.surfaceScanSummary
    : [];

  return [
    scanRows.length,
    scanRows.map(row => [
      row.StationKey || "",
      row.TotalScansToday || "",
      row.PeakHour || "",
      row.PeakHourScans || "",
      row.LastUpdated || ""
    ].join(":")).join("|")
  ].join("||");
}

function getLatestActivityTimestamp(scanRows) {
  if (!Array.isArray(scanRows) || !scanRows.length) return null;

  const values = scanRows
    .map(row => String(row.LastUpdated || "").trim())
    .filter(Boolean);

  if (!values.length) return null;

  // if all scan rows share same timestamp, first is enough
  return values[0];
}

function determineUpdateSource(payload) {
  const newWipSignature = buildWipSignature(payload);
  const newActivitySignature = buildActivitySignature(payload);

  const wipChanged =
    state.lastWipSignature !== null &&
    newWipSignature !== state.lastWipSignature;

  const activityChanged =
    state.lastActivitySignature !== null &&
    newActivitySignature !== state.lastActivitySignature;

  let source = "—";
  let timeValue = state.lastFetch || new Date();

  if (wipChanged && activityChanged) {
    source = "WIP + Activity";
    timeValue =
      payload.summary?.ImportedAt ||
      getLatestActivityTimestamp(payload.surfaceScanSummary) ||
      new Date();
  } else if (wipChanged) {
    source = "WIP";
    timeValue =
      payload.summary?.ImportedAt ||
      new Date();
  } else if (activityChanged) {
    source = "Activity";
    timeValue =
      getLatestActivityTimestamp(payload.surfaceScanSummary) ||
      new Date();
  } else if (!state.hasRenderedOnce) {
    // first load
    source = "WIP + Activity";
    timeValue =
      payload.summary?.ImportedAt ||
      getLatestActivityTimestamp(payload.surfaceScanSummary) ||
      new Date();
  } else {
    source = state.lastUpdateSource || "—";
    timeValue = state.lastUpdateTime || state.lastFetch || new Date();
  }

  return {
    newWipSignature,
    newActivitySignature,
    source,
    timeValue
  };
}


/* =========================================================
   SECTION 05 — LOADER HELPERS
========================================================= */

function setSurfaceLoaderProgress(percent, message) {
  const fill = document.getElementById("surfaceLoaderFill");
  const status = document.getElementById("surfaceLoaderStatus");

  if (fill) {
    fill.style.width = Math.max(0, Math.min(100, percent)) + "%";
  }

  if (status && message) {
    status.textContent = message;
  }
}

function hideSurfaceLoader() {
  const loader = document.getElementById("surfaceLoader");

  if (!loader) return;

  setSurfaceLoaderProgress(100, "Surface dashboard ready");

  setTimeout(() => {
    loader.classList.add("is-hidden");
  }, 350);

  setTimeout(() => {
    if (loader && loader.parentNode) {
      loader.remove();
    }
  }, 1200);
}

function startSurfaceLoaderAnimation() {
  const loader = document.getElementById("surfaceLoader");

  if (!loader) return;

  const steps = [
    { pct: 18, msg: "Authenticating Surface access…" },
    { pct: 38, msg: "Loading WIP stations…" },
    { pct: 58, msg: "Reading scan summary…" },
    { pct: 76, msg: "Building live surface flow…" },
    { pct: 92, msg: "Rendering dashboard…" }
  ];

  let index = 0;

  window.surfaceLoaderTimer = setInterval(() => {
    if (index >= steps.length) {
      clearInterval(window.surfaceLoaderTimer);
      return;
    }

    const step = steps[index++];
    setSurfaceLoaderProgress(step.pct, step.msg);
  }, 420);
}


/* =========================================================
   SECTION 06 — API FETCH + DATA SIGNATURE
========================================================= */

function buildDataSignature(payload) {
  const summary = payload.summary || {};
  const meta = payload.intervalHistoryMeta || {};
  const scanRows = Array.isArray(payload.surfaceScanSummary)
    ? payload.surfaceScanSummary
    : [];

  const scanSignature = scanRows
    .map(row => [
      row.StationKey || "",
      row.TotalScansToday || "",
      row.LastUpdated || ""
    ].join(":"))
    .join("|");

  return [
    summary.ReportDate || "",
    summary.SourceFile || "",
    summary.FileUpdatedTime || "",
    summary.ImportedAt || "",
    meta.currentKey || "",
    meta.currentTime || "",
    (payload.surfaceFlow || []).length,
    (payload.surfaceTransfers || []).length,
    scanRows.length,
    scanSignature
  ].join("|");
}

function normalizeProductionFlowPayload(payload) {
  if (!payload || payload.status !== "success") return payload;

  const rows = Array.isArray(payload.productionFlow)
    ? payload.productionFlow
    : Array.isArray(payload.surfaceFlow)
      ? payload.surfaceFlow
      : [];

  const apiSummary = payload.summary || {};
  const totalWip = num(apiSummary.TotalWIP);
  const totalActivity = num(apiSummary.TotalActivityToday);
  const outputActivity = num(apiSummary.OutputActivity);
  const bridgeInputWip = num(apiSummary.BridgeInputWIP);

  const flowRows = rows.map(row => {
    const currentWip = num(row.CurrentWIP ?? row.CurrentJobTotal);
    const activityToday = num(row.ActivityToday ?? row.TotalScansToday);
    const flowStep = safeText(row.FlowStep, "");
    const displayName = safeText(row.DisplayName || row.FlowStep, "");
    const metricMode = safeText(row.MetricMode, "WIP_AND_ACTIVITY");
    const percent = totalWip > 0 ? currentWip / totalWip : 0;

    return {
      ...row,
      FlowOrder: num(row.FlowOrder),
      FlowStep: flowStep,
      DisplayName: displayName,
      CurrentWIP: currentWip,
      ActivityToday: activityToday,
      CurrentJobTotal: currentWip,
      PercentOfSurfaceTotal: percent,
      FlowGroup: safeText(row.RollupArea || row.Area || row.ReportDepartment, "Surface"),
      Status: metricMode === "OUTPUT_ONLY" ? "OUTPUT" : "",
      Notes: metricMode === "OUTPUT_ONLY"
        ? "Output/activity row. WIP is intentionally excluded."
        : metricMode === "WIP_ONLY"
          ? "Input queue row. Activity is intentionally excluded."
          : "",
      TotalScansToday: activityToday
    };
  });

  const intakeRow = flowRows.find(row => row.FlowStep === "SF Scan & Verify") || {};
  const activeSteps = flowRows.filter(row => num(row.CurrentJobTotal) > 0).length;
  const emptySteps = flowRows.filter(row => num(row.CurrentJobTotal) === 0).length;

  const normalizedSummary = {
    ...apiSummary,
    SurfaceTotalWIP: totalWip,
    SurfaceMainWIP: Math.max(0, totalWip - num(intakeRow.CurrentJobTotal)),
    SurfaceIntakeWIP: num(intakeRow.CurrentJobTotal),
    SurfaceTotalActivityToday: totalActivity,
    SurfaceOutputActivity: outputActivity,
    BridgeInputWIP: bridgeInputWip,
    LargestStep: apiSummary.LargestWIPStation || apiSummary.LargestStep || "—",
    LargestStepTotal: apiSummary.LargestWIPTotal || apiSummary.LargestStepTotal || 0,
    ActiveSteps: apiSummary.ActiveWIPStations || activeSteps,
    EmptySteps: emptySteps,
    ImportedAt: apiSummary.LastUpdated || payload.generatedAt || new Date().toISOString(),
    SourceFile: "PRODUCTION_FLOW_CURRENT",
    FileUpdatedTime: apiSummary.LastUpdated || payload.generatedAt || ""
  };

  const scanSummary = flowRows.map(row => ({
    StationKey: row.FlowStep,
    DisplayName: row.DisplayName,
    TotalScansToday: num(row.ActivityToday),
    PeakHour: "—",
    PeakHourScans: 0,
    LastUpdated: row.LastUpdated || normalizedSummary.ImportedAt
  }));

  return {
    ...payload,
    summary: normalizedSummary,
    surfaceCurrent: flowRows,
    surfaceFlow: flowRows,
    surfaceTransfers: Array.isArray(payload.productionTransfers) ? payload.productionTransfers : [],
    surfaceTransferDailyTotals: [],
    surfaceScanSummary: scanSummary,
    surfaceHourlyInOut: [],
    intervalHistoryMeta: {
      currentKey: `${payload.requestedArea || CONFIG.AREA}|${normalizedSummary.LastUpdated || normalizedSummary.ImportedAt}`,
      currentTime: normalizedSummary.LastUpdated || normalizedSummary.ImportedAt,
      intervalRows: flowRows.length,
      intervalSnapshots: 1
    }
  };
}

async function fetchData(forceRender = false) {
  setSystemStatus("loading");

  try {
    const url = `${CONFIG.API_URL}?action=productionFlow&area=${encodeURIComponent(CONFIG.AREA || "Surface")}&t=${Date.now()}`;

    const res = await fetch(url, {
      cache: "no-store"
    });

    if (!res.ok) {
      throw new Error("HTTP " + res.status);
    }

    const rawJson = await res.json();
    const json = normalizeProductionFlowPayload(rawJson);

    if (!json || json.status !== "success") {
      throw new Error(json?.message || "API returned error");
    }

    const incomingFlow = Array.isArray(json.surfaceFlow) ? json.surfaceFlow : [];
    const incomingTransfers = Array.isArray(json.surfaceTransfers) ? json.surfaceTransfers : [];
    const incomingScanSummary = Array.isArray(json.surfaceScanSummary)
      ? json.surfaceScanSummary
      : [];

    if (!incomingFlow.length && state.hasRenderedOnce) {
      console.warn("[SurfaceWIP] Refresh returned empty productionFlow. Keeping last good data.");
      setSystemStatus("ok");
      updateLatestUpdatePill();
      return;
    }

    if (!incomingFlow.length && !state.hasRenderedOnce) {
      throw new Error("No Production Flow rows returned from API.");
    }

    const newSignature = buildDataSignature(json);
    const isNewData = newSignature !== state.lastDataSignature;
    const updateInfo = determineUpdateSource(json);

    if (!forceRender && state.hasRenderedOnce && !isNewData) {
      state.lastFetch = new Date();
      setSystemStatus("ok");
      updateLatestUpdatePill();
      return;
    }

    state.summary = json.summary || state.summary || {};
    state.surfaceFlow = incomingFlow;
    state.surfaceTransfers = incomingTransfers.length
      ? incomingTransfers
      : state.surfaceTransfers || [];
    state.surfaceScanSummary = incomingScanSummary;

    state.surfaceHourlyInOut = Array.isArray(json.surfaceHourlyInOut)
      ? json.surfaceHourlyInOut
      : [];

    state.surfaceTransferDailyTotals = Array.isArray(json.surfaceTransferDailyTotals)
      ? json.surfaceTransferDailyTotals
      : [];

    state.intervalMeta = json.intervalHistoryMeta || state.intervalMeta || {};
    state.lastFetch = new Date();
    state.lastDataSignature = newSignature;

    state.lastWipSignature = updateInfo.newWipSignature;
    state.lastActivitySignature = updateInfo.newActivitySignature;
    state.lastUpdateSource = updateInfo.source;
    state.lastUpdateTime = updateInfo.timeValue;

    setSystemStatus("ok");
    updateReportMeta();
    updateLatestUpdatePill();

    renderAll();

    state.hasRenderedOnce = true;

    if (window.surfaceLoaderTimer) {
      clearInterval(window.surfaceLoaderTimer);
    }

    hideSurfaceLoader();

  } catch (err) {
    console.error("[SurfaceWIP] Fetch error:", err);

    if (state.hasRenderedOnce) {
      setSystemStatus("error");
      return;
    }

    setSystemStatus("error");
    renderError(err.message);

    if (window.surfaceLoaderTimer) {
      clearInterval(window.surfaceLoaderTimer);
    }

    setSurfaceLoaderProgress(100, "Unable to load Surface dashboard");
    setTimeout(hideSurfaceLoader, 700);
  }
}

/* =========================================================
   SECTION 07 — STATUS + META UI
========================================================= */

function setSystemStatus(status) {
  const dot = document.querySelector(".system-ok .dot");
  const text = document.getElementById("systemStatusText");

  if (!dot || !text) return;

  if (status === "ok") {
    dot.className = "dot green-dot";
    text.textContent = "All Systems OK";
  } else if (status === "loading") {
    dot.className = "dot yellow-dot";
    text.textContent = "Refreshing…";
  } else {
    dot.className = "dot red-dot";
    text.textContent = "API Error";
  }
}

function updateReportMeta() {
  const el = document.getElementById("reportMeta");
  if (!el) return;

  const s = state.summary || {};
  const date = safeText(s.ReportDate);
  const imported = safeText(s.ImportedAt);

  el.textContent = `Report: ${date} · Imported: ${imported}`;
}

function updateLatestUpdatePill() {
  const pill = document.getElementById("latestUpdatePill");
  const summaryPill = document.getElementById("summaryLatestUpdate");

  if (!pill && !summaryPill) return;

  const timeValue =
    state.lastUpdateTime ||
    state.summary?.ImportedAt ||
    state.lastFetch;

  const source =
    state.lastUpdateSource || "—";

  if (!timeValue) {
    if (pill) pill.textContent = "Latest Update: —";
    if (summaryPill) summaryPill.textContent = "Latest Update: —";
    return;
  }

  const label =
    "Latest Update: " +
    formatDisplayDateTime(timeValue) +
    " · " +
    source;

  if (pill) pill.textContent = label;
  if (summaryPill) summaryPill.textContent = label;

  const currentKey =
    state.intervalMeta?.currentKey ||
    state.summary?.SourceFile ||
    "";

  if (currentKey && state.lastSnapshotKey && currentKey !== state.lastSnapshotKey) {
    if (pill) {
      pill.classList.remove("fresh");
      void pill.offsetWidth;
      pill.classList.add("fresh");
    }

    if (summaryPill) {
      summaryPill.classList.remove("fresh");
      void summaryPill.offsetWidth;
      summaryPill.classList.add("fresh");
    }
  }

  if (currentKey) {
    state.lastSnapshotKey = currentKey;
  }
}

function formatDisplayDateTime(value) {
  if (!value) return "—";

  const raw = String(value).replace(/^'/, "").trim();
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})$/);

  if (match) {
    const d = new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4]),
      Number(match[5]),
      Number(match[6])
    );

    return d.toLocaleTimeString([], {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
  }

  const d = new Date(raw);

  if (!isNaN(d.getTime())) {
    return d.toLocaleTimeString([], {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
  }

  return raw;
}

function renderError(message) {
  const flow = document.getElementById("flowGrid");
  const transfer = document.getElementById("transferGrid");

  if (flow) {
    flow.innerHTML = `<div class="flow-loading" style="color:var(--red)">⚠ ${message}</div>`;
  }

  if (transfer) {
    transfer.innerHTML = `<div class="transfer-loading" style="color:var(--red)">⚠ ${message}</div>`;
  }
}


/* =========================================================
   SECTION 08 — MASTER RENDER
========================================================= */

function renderAll() {
  renderKPIs();
  renderFlowGrid();
  renderTransferTunnel();
  renderHourlyInOutComparison();
  renderFlowDetail();
  renderLines();
  renderAlerts();
  renderAnalytics();
  renderReports();
}


/* =========================================================
   SECTION 09 — KPI CARDS
========================================================= */

function renderKPIs() {
  const s = state.summary;
  if (!s) return;

  const total = num(s.SurfaceTotalWIP);
  const largestTotal = num(s.LargestStepTotal);
  const largestStep = safeText(s.LargestStep);
  const pctLargest = total > 0 ? Math.round((largestTotal / total) * 100) : 0;
  const inspectionTotal = getStationScanTotal("Surface Inspection");

  setText("kpiTotal", total.toLocaleString());
  setText("kpiMain", num(s.SurfaceMainWIP).toLocaleString());
  setText("kpiActiveSteps", `${num(s.ActiveSteps)} active · ${num(s.EmptySteps)} empty steps`);
  setText("kpiIntake", `Intake (SF Scan): ${num(s.SurfaceIntakeWIP).toLocaleString()}`);
  setText("kpiInspection", inspectionTotal.toLocaleString());

  const bottleneck = document.getElementById("kpiBottleneck");
  if (bottleneck) {
    bottleneck.innerHTML = `${largestTotal.toLocaleString()} <span style="font-size:14px;color:var(--red)">— ${largestStep}</span>`;
  }

  setText("kpiBottleneckSub", `${pctLargest}% of total Surface WIP`);
}


/* =========================================================
   SECTION 10 — TOP FLOW GRID
========================================================= */

function renderFlowGrid() {
  const grid = document.getElementById("flowGrid");
  if (!grid) return;

  let rows = filterRows(state.surfaceFlow, state.filter);

  if (CONFIG.LINE_A_OFFLINE) {
    rows = rows.filter(r => !isLineA(r.FlowStep));
  }

  if (!rows.length) {
    grid.innerHTML = `<div class="flow-loading">No stations match this filter.</div>`;
    updateBottleneckLabel([]);
    return;
  }

  const csv = [headers.join(","), ...csvRows].join("\n");

  grid.innerHTML = rows.map((row, index) => {
    const wip = num(row.CurrentJobTotal);
    const flowStep = safeText(row.FlowStep, "");
    const displayName = safeText(row.DisplayName || row.FlowStep, "");
    const group = safeText(row.FlowGroup, "");
    const order = num(row.FlowOrder);
    const status = statusFromBackend(row);
    const cls = stageClass(status, flowStep);
    const pill = statusPillText(status, flowStep);
    const bw = barWidth(wip, maxWip);
    const icon = iconForStep(flowStep);
    const isLast = index === rows.length - 1;

    return `
      <article class="stage ${cls}" data-order="${order}" data-flow-step="${flowStep}">
        <div class="stage-num">${index + 1}</div>
        <div class="stage-icon">${icon}</div>
        <div class="stage-name">${displayName}</div>
        <div class="stage-group">${group}</div>
        <div class="stage-count">${wip.toLocaleString()}</div>
        <div class="status-pill ${pill}">${pill}</div>

        <div class="health-row">
          <div class="bar-track">
            <div class="bar-fill" style="width:${bw}%"></div>
          </div>
          <div class="stage-percent">${pct(row.PercentOfSurfaceTotal)}%</div>
        </div>

        ${!isLast ? `<div class="stage-arrow">→</div>` : ""}
      </article>
    `;
  }).join("");

  updateBottleneckLabel(rows);
}

function updateBottleneckLabel(rows) {
  const label = document.getElementById("bottleneckLabel");
  if (!label) return;

  const hasCritical = rows.some(row => {
    const flowStep = String(row.FlowStep || "");
    return statusFromBackend(row) === "Critical" && !isLineA(flowStep);
  });

  label.classList.toggle("visible", hasCritical);
}


/* =========================================================
   SECTION 11 — LIVE SNAKE FLOW MAIN RENDER
   NEW STUFF: No duplicate station cards between rows.
========================================================= */

function renderTransferTunnel() {
  const grid = document.getElementById("transferGrid");
  const meta = document.getElementById("transferMeta");

  if (!grid || !meta) return;

  const rows = buildLiveTransferRowsFromSurfaceFlow_();

  if (!rows.length) {
    meta.textContent = "Waiting for live Surface WIP data…";

    grid.innerHTML = `
      <div class="transfer-loading">
        <div class="spinner"></div>
        <span>Waiting for current Surface WIP rows before the live flow can render.</span>
      </div>
    `;
    return;
  }

  meta.style.display = "none";
  grid.innerHTML = renderContinuousSurfaceFlow(rows);
}

function buildLiveTransferRowsFromSurfaceFlow_() {
  const rows = state.surfaceFlow || [];
  if (!rows.length) return [];

  const map = {};

  rows.forEach(row => {
    const step = String(row.FlowStep || "").trim();
    if (step) map[step] = row;
  });

  const transitions = [
    { from: "SF Scan & Verify", to: "Surface Unbox", transitionName: "SF Scan & Verify → SF Unbox" },
    { from: "Surface Unbox", to: "Blocking Line B", transitionName: "SF Unbox → Auto Blockers" },
    { from: "Blocking Line B", to: "Cooling Storage", transitionName: "Auto Blockers → IQ Star" },
    { from: "Cooling Storage", to: "Generating Line B", transitionName: "IQ Star → Orbit Generator" },
    { from: "Generating Line B", to: "Polishing Line B", transitionName: "Orbit Generator → Polisher" },
    { from: "Polishing Line B", to: "Engraving Line B", transitionName: "Polisher → Engraver" },
    { from: "Engraving Line B", to: "Detaping Line B", transitionName: "Engraver → Detaper" },
    { from: "Detaping Line B", to: "Coating Line B", transitionName: "Detaper → 54R Coater" },
    { from: "Coating Line B", to: "Surface Inspection", transitionName: "54R Coater → Inspection" }
  ];

  return transitions.map(pair => {
    const fromRow = map[pair.from] || {};
    const toRow = map[pair.to] || {};
    const fromWip = num(fromRow.CurrentJobTotal);
    const toWip = num(toRow.CurrentJobTotal);

    return {
      GeneratedAt: state.summary?.ImportedAt || state.lastFetch || "",
      TransitionName: pair.transitionName,

      FromOrder: num(fromRow.FlowOrder),
      FromStep: pair.from,
      FromDisplayName: safeText(fromRow.DisplayName || pair.from),
      FromCurrentWIP: fromWip,

      ToOrder: num(toRow.FlowOrder),
      ToStep: pair.to,
      ToDisplayName: safeText(toRow.DisplayName || pair.to),
      ToCurrentWIP: toWip,

      /*
        Current WIP Up Now:
        This is not interval movement.
        This is current WIP at the previous station feeding the next station.
      */
      EstimatedInTransfer: fromWip,

      Confidence: "Live",
      Notes: "Live WIP from current Facility WIP report."
    };
  });
}

function calcEstimatedMoving(row) {
  return num(row.EstimatedInTransfer);
}


/* =========================================================
   SECTION 12 — LIVE SNAKE FLOW ROW BUILDING
   NEW STUFF: Rows are grouped by transitions, not station count.
========================================================= */

function renderContinuousSurfaceFlow(rows) {
  const orderedRows = rows
    .slice()
    .sort((a, b) => num(a.FromOrder) - num(b.FromOrder));

  const totalMoving = orderedRows.reduce((sum, row) => {
    return sum + calcEstimatedMoving(row);
  }, 0);

  const totalScansToday = (state.surfaceScanSummary || []).reduce((sum, row) => {
    return sum + num(row.TotalScansToday);
  }, 0);

  const activeStations = (state.surfaceFlow || []).filter(row => {
    return num(row.CurrentJobTotal) > 0 && !(CONFIG.LINE_A_OFFLINE && isLineA(row.FlowStep));
  }).length;

  const bottleneckRow = (state.surfaceFlow || [])
    .filter(row => !(CONFIG.LINE_A_OFFLINE && isLineA(row.FlowStep)))
    .slice()
    .sort((a, b) => num(b.CurrentJobTotal) - num(a.CurrentJobTotal))[0] || {};

  const bottleneckName = safeText(bottleneckRow.DisplayName || bottleneckRow.FlowStep, "No bottleneck");
  const bottleneckWip = num(bottleneckRow.CurrentJobTotal);
  const bottleneckStatus = bottleneckWip >= CONFIG.WIP_CRITICAL ? "Critical" : bottleneckWip >= CONFIG.WIP_HIGH ? "Watch" : "Stable";

  const watchCount = (state.surfaceFlow || []).filter(row => {
    const wip = num(row.CurrentJobTotal);
    return wip >= CONFIG.WIP_HIGH && !(CONFIG.LINE_A_OFFLINE && isLineA(row.FlowStep));
  }).length;

  const flowCards = buildSerpentineFlowCards(orderedRows);
  const flowRows = chunkCards(flowCards, 5);

  const updateTime =
    state.lastUpdateTime ||
    state.summary?.ImportedAt ||
    state.lastFetch ||
    new Date();

  return `
    <section class="continuous-flow-shell surface-v3-flow option3-command-shell option3-clean-shell">

      <div class="surface-v3-bg-orb orb-one"></div>
      <div class="surface-v3-bg-orb orb-two"></div>
      <div class="surface-v3-bg-orb orb-three"></div>

      <div class="option3-clean-toolbar">
        <div>
          <span class="option3-section-label">Process Flow</span>
          
        </div>

        <div class="option3-clean-meta">
          <span class="option3-status-chip ${bottleneckStatus.toLowerCase()}">${bottleneckStatus}</span>
          <span class="surface-v3-live-pill"><span></span>Live</span>
          <span class="surface-v3-updated">${formatDisplayDateTime(updateTime)}</span>
        </div>
      </div>

      <div class="option3-layout option3-layout-clean">
<main class="option3-flow-stage option3-flow-stage-clean">
          <div class="serpentine-flow-wrap surface-v3-serpentine option3-serpentine">
            ${flowRows.map((rowCards, rowIndex) => renderSerpentineRow(rowCards, rowIndex, flowRows.length)).join("")}
          </div>
        </main>
      </div>

      <div class="surface-v3-bottom-strip option3-bottom-strip option3-bottom-strip-clean">
        <div class="surface-v3-status-card">
          <div class="surface-v3-mini-icon">🛡</div>
          <div><span>System Health</span><strong>Operational</strong></div>
        </div>

        <div class="surface-v3-mini-metric"><span>Stations Online</span><strong>${activeStations.toLocaleString()}/10</strong></div>
        <div class="surface-v3-mini-metric"><span>Watch Areas</span><strong>${watchCount.toLocaleString()}</strong></div>
        <div class="surface-v3-mini-metric"><span>Moving WIP</span><strong>${totalMoving.toLocaleString()}</strong></div>
        <div class="surface-v3-live-note"><span></span>Live data updates every 30 seconds</div>
      </div>
    </section>
  `;
}


function buildSerpentineFlowCards(orderedRows) {
  const cards = [];

  orderedRows.forEach((row, index) => {
    if (index === 0) {
      cards.push({
        type: "station",
        step: row.FromStep,
        display: row.FromDisplayName
      });
    }

    cards.push({
      type: "wip",
      row: row
    });

    cards.push({
      type: "station",
      step: row.ToStep,
      display: row.ToDisplayName
    });
  });

  return cards;
}

function chunkCards(cards, size) {
  const chunks = [];

  for (let i = 0; i < cards.length; i += size) {
    chunks.push(cards.slice(i, i + size));
  }

  return chunks;
}

function renderSerpentineRow(rowCards, rowIndex, totalRows) {
  const isReverseRow = rowIndex % 2 === 1;

  const visualCards = isReverseRow
    ? rowCards.slice().reverse()
    : rowCards;

  const inlineArrow = isReverseRow ? "←" : "→";
  const hasNextRow = rowIndex < totalRows - 1;

  const rowPieces = [];

  visualCards.forEach((card, index) => {
    rowPieces.push(renderSerpentineCard(card));

    if (index < visualCards.length - 1) {
      rowPieces.push(`
        <div class="serpentine-inline-arrow ${isReverseRow ? "reverse" : "forward"}" aria-hidden="true">
          ${inlineArrow}
        </div>
      `);
    }
  });

  return `
    <div class="serpentine-row-block ${isReverseRow ? "reverse-row" : "forward-row"}" data-row-index="${rowIndex}">
      <div class="serpentine-row ${isReverseRow ? "reverse-row" : ""}">
        ${rowPieces.join("")}
      </div>

      ${
        hasNextRow
          ? renderSerpentineDropArrow(isReverseRow)
          : ""
      }
    </div>
  `;
}

function renderSerpentineDropArrow(isReverseRow) {
  return `
    <div class="serpentine-row-drop ${isReverseRow ? "left" : "right"}" aria-hidden="true">
      <div class="serpentine-drop-line"></div>
      <div class="serpentine-drop-arrow">↓</div>
    </div>
  `;
}

function renderSerpentineCard(card) {
  if (card.type === "station") {
    return renderContinuousStationCard({
      step: card.step,
      display: card.display
    });
  }

  if (card.type === "wip") {
    return renderContinuousConveyor(card.row);
  }

  return "";
}



function renderContinuousStationCard(station) {
  const scanTotal = getStationScanTotal(station.step);
  const display = getStationScanDisplay(station.step, station.display);
  const peak = getStationPeakText(station.step);

  return `
    <article class="continuous-station-card">
      <div class="surface-v3-station-icon" aria-hidden="true">
        ${iconForStep(station.step)}
      </div>

      <div class="transfer-zone-label">Total Scan Today</div>
      <div class="transfer-zone-step">${display}</div>
      <div class="transfer-zone-value">${scanTotal.toLocaleString()}</div>
      ${peak ? `<div class="transfer-zone-sub">${peak}</div>` : ""}
    </article>
  `;
}

function renderContinuousConveyor(row) {
  const wip = calcEstimatedMoving(row);

  const severity =
    wip >= CONFIG.WIP_CRITICAL ? "critical" :
    wip >= CONFIG.WIP_HIGH ? "watch" :
    "normal";

  const packetsToShow = Math.max(
    3,
    Math.min(10, wip > 0 ? Math.ceil(wip / 30) : 3)
  );

  const packets = Array.from(
    { length: packetsToShow },
    (_, i) => `<span class="transfer-packet" style="--delay:${i * 0.34}s"></span>`
  ).join("");

  return `
    <article class="continuous-conveyor ${severity}">
      <div class="transfer-conveyor-head">
        <div class="transfer-conveyor-title">Current WIP Moving</div>
        <div class="transfer-conveyor-count">${wip.toLocaleString()}</div>
      </div>

      <div class="transfer-track">
        ${packets}
      </div>

      <div class="continuous-conveyor-route">
        ${row.FromDisplayName} → ${row.ToDisplayName}
      </div>
    </article>
  `;
}

/* =========================================================
   HOURLY IN / OUT COMPARISON TAB
========================================================= */

/* =========================================================
   HOURLY IN / OUT COMPARISON TAB
   TREND CHART VERSION
========================================================= */

function renderHourlyInOutComparison() {
  const meta = document.getElementById("transferMetaSecondary");
  const grid = document.getElementById("transferGridSecondary");

  if (!grid) return;

  const comparisons = state.surfaceHourlyInOut || [];

  if (!comparisons.length) {
    if (meta) meta.textContent = "No hourly activity data";

    grid.innerHTML = `
      <div class="transfer-loading">
        Waiting for SURFACE_SCAN_RAW hourly activity data.
      </div>
    `;
    return;
  }

  const selectedIndex = Math.min(
    Math.max(state.activeComparisonIndex || 0, 0),
    comparisons.length - 1
  );

  const selected = comparisons[selectedIndex];

  if (meta) {
    meta.textContent = selected.overallStatus || "Hourly Comparison";
  }

  const totalDeltaClass =
    selected.totalDelta >= 25 ? "building" :
    selected.totalDelta <= -25 ? "draining" :
    "even";

  grid.innerHTML = `
    <section class="hourly-compare-shell">

      <div class="hourly-compare-header">
        <div>
          <div class="hourly-compare-title">Hourly In / Out Comparison</div>
          <div class="hourly-compare-subtitle">
            Received this hour = current activity hour minus previous activity hour.
          </div>
        </div>

        <div class="hourly-compare-status ${totalDeltaClass}">
          ${selected.overallStatus}
        </div>
      </div>

      <div class="hourly-transition-tabs">
        ${comparisons.map((item, index) => `
          <button
            class="hourly-transition-btn ${index === selectedIndex ? "active" : ""}"
            data-comparison-index="${index}">
            ${item.toStep}
          </button>
        `).join("")}
      </div>

      <div class="hourly-compare-main">
        <article class="hourly-compare-card">
          <div class="hourly-card-label">Previous Step Received</div>
          <div class="hourly-card-step">${selected.fromStep}</div>
          <div class="hourly-card-value">${num(selected.totalFromReceived).toLocaleString()}</div>
        </article>

        <article class="hourly-compare-card">
          <div class="hourly-card-label">Current Step Received</div>
          <div class="hourly-card-step">${selected.toStep}</div>
          <div class="hourly-card-value">${num(selected.totalToReceived).toLocaleString()}</div>
        </article>

        <article class="hourly-compare-card ${totalDeltaClass}">
          <div class="hourly-card-label">Net Delta</div>
          <div class="hourly-card-step">${selected.transitionName}</div>
          <div class="hourly-card-value">${formatSignedNumber(selected.totalDelta)}</div>
        </article>
      </div>

      <div class="hourly-chart-shell">
        <div class="hourly-chart-head">
          <div class="hourly-chart-title">${selected.transitionName} Trend</div>
          <div class="hourly-chart-subtitle">
            Previous step vs current step with delta by hour
          </div>
        </div>

        <div class="hourly-chart-wrap">
          <canvas id="hourlyInOutTrendChart"></canvas>
        </div>
      </div>

      <div class="hourly-compare-table-wrap">
        <table class="hourly-compare-table">
          <thead>
            <tr>
              <th>Hour</th>
              <th>${selected.fromStep}</th>
              <th>${selected.toStep}</th>
              <th>Delta</th>
              <th>Status</th>
            </tr>
          </thead>

          <tbody>
            ${selected.hourly.map(row => {
              const rowClass =
                row.delta >= 25 ? "building" :
                row.delta <= -25 ? "draining" :
                "even";

              return `
                <tr class="${rowClass}">
                  <td>${formatHourLabel(row.hour)}</td>
                  <td>${num(row.fromReceived).toLocaleString()}</td>
                  <td>${num(row.toReceived).toLocaleString()}</td>
                  <td>${formatSignedNumber(row.delta)}</td>
                  <td><span class="hourly-status-pill ${rowClass}">${row.status}</span></td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      </div>

    </section>
  `;

  grid.querySelectorAll(".hourly-transition-btn").forEach(button => {
    button.addEventListener("click", () => {
      state.activeComparisonIndex = num(button.dataset.comparisonIndex);
      renderHourlyInOutComparison();
    });
  });

  renderHourlyInOutTrendChart(selected);
}

function renderHourlyInOutTrendChart(selected) {
  const canvas = document.getElementById("hourlyInOutTrendChart");
  if (!canvas) return;

  if (typeof Chart === "undefined") {
    console.error("Chart.js is not loaded.");
    return;
  }

  if (hourlyInOutChart) {
    hourlyInOutChart.destroy();
    hourlyInOutChart = null;
  }

  const labels = (selected.hourly || []).map(row => formatHourLabel(row.hour));
  const fromData = (selected.hourly || []).map(row => num(row.fromReceived));
  const toData = (selected.hourly || []).map(row => num(row.toReceived));
  const deltaData = (selected.hourly || []).map(row => num(row.delta));

  const ctx = canvas.getContext("2d");

  hourlyInOutChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          type: "line",
          label: selected.fromStep,
          data: fromData,
          borderColor: "#58d8ff",
          backgroundColor: "rgba(88, 216, 255, 0.15)",
          borderWidth: 3,
          tension: 0.35,
          fill: false,
          pointRadius: 4,
          pointHoverRadius: 6,
          yAxisID: "y"
        },
        {
          type: "line",
          label: selected.toStep,
          data: toData,
          borderColor: "#50dc96",
          backgroundColor: "rgba(80, 220, 150, 0.15)",
          borderWidth: 3,
          tension: 0.35,
          fill: false,
          pointRadius: 4,
          pointHoverRadius: 6,
          yAxisID: "y"
        },
        {
          type: "bar",
          label: "Delta",
          data: deltaData,
          backgroundColor: deltaData.map(value =>
            value >= 25
              ? "rgba(255, 71, 87, 0.55)"
              : value <= -25
              ? "rgba(34, 211, 238, 0.55)"
              : "rgba(80, 220, 150, 0.45)"
          ),
          borderColor: deltaData.map(value =>
            value >= 25
              ? "rgba(255, 71, 87, 1)"
              : value <= -25
              ? "rgba(34, 211, 238, 1)"
              : "rgba(80, 220, 150, 1)"
          ),
          borderWidth: 1,
          yAxisID: "y1"
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false
      },
      plugins: {
        legend: {
          labels: {
            color: "#d8ecff",
            font: {
              weight: "700"
            }
          }
        },
        tooltip: {
          backgroundColor: "rgba(5, 16, 34, 0.96)",
          titleColor: "#ffffff",
          bodyColor: "#dff4ff",
          borderColor: "rgba(88,216,255,0.3)",
          borderWidth: 1
        }
      },
      scales: {
        x: {
          ticks: {
            color: "#9fc7df",
            font: {
              weight: "700"
            }
          },
          grid: {
            color: "rgba(88,216,255,0.08)"
          }
        },
        y: {
          position: "left",
          beginAtZero: true,
          ticks: {
            color: "#9fc7df",
            font: {
              weight: "700"
            }
          },
          grid: {
            color: "rgba(88,216,255,0.08)"
          },
          title: {
            display: true,
            text: "Received",
            color: "#58d8ff",
            font: {
              weight: "800"
            }
          }
        },
        y1: {
          position: "right",
          beginAtZero: true,
          ticks: {
            color: "#ff8b99",
            font: {
              weight: "700"
            }
          },
          grid: {
            drawOnChartArea: false
          },
          title: {
            display: true,
            text: "Delta",
            color: "#ff5a6f",
            font: {
              weight: "800"
            }
          }
        }
      }
    }
  });
}


/* =========================================================
   SECTION 14 — SCAN SUMMARY HELPERS
========================================================= */

const SCAN_STEP_ALIASES = {
  "SF Scan & Verify": ["SF Scan & Verify", "Incoming SF Scan & Verify"],

  "Surface Unbox": ["Surface Unbox", "SF Unbox"],
  "SF Unbox": ["Surface Unbox", "SF Unbox"],

  "Blocking Line B": ["Blocking Line B", "Auto Blockers", "Blocking Line"],
  "Auto Blockers": ["Blocking Line B", "Auto Blockers", "Blocking Line"],

  "Cooling Storage": ["Cooling Storage", "IQ Star"],
  "IQ Star": ["Cooling Storage", "IQ Star"],

  "Generating Line B": ["Generating Line B", "Orbit Generator", "Generating Line"],
  "Orbit Generator": ["Generating Line B", "Orbit Generator", "Generating Line"],

  "Polishing Line B": ["Polishing Line B", "Polisher", "Polisher B", "Polishing Line"],
  "Polisher": ["Polishing Line B", "Polisher", "Polisher B", "Polishing Line"],

  "Engraving Line B": ["Engraving Line B", "Engraver", "Engraving Line"],
  "Engraver": ["Engraving Line B", "Engraver", "Engraving Line"],

  "Detaping Line B": ["Detaping Line B", "Detaper", "Detaper Line"],
  "Detaper": ["Detaping Line B", "Detaper", "Detaper Line"],

  "Coating Line B": ["Coating Line B", "54R Coater", "Coater Line"],
  "54R Coater": ["Coating Line B", "54R Coater", "Coater Line"],

  "Surface Inspection": ["Surface Inspection", "Inspection"],
  "Inspection": ["Surface Inspection", "Inspection"]
};

function getScanAliases(value) {
  const key = String(value || "").trim();
  if (!key) return [];

  const aliases = SCAN_STEP_ALIASES[key] || [key];
  return aliases.map(v => String(v || "").trim()).filter(Boolean);
}

function getScanRowForStep(flowStepOrDisplay) {
  const aliases = getScanAliases(flowStepOrDisplay);
  if (!aliases.length) return null;

  return (state.surfaceScanSummary || []).find(row => {
    const stationKey = String(row.StationKey || "").trim();
    const displayName = String(row.DisplayName || "").trim();

    return aliases.includes(stationKey) || aliases.includes(displayName);
  }) || null;
}

function getStationScanTotal(flowStepOrDisplay) {
  const row = getScanRowForStep(flowStepOrDisplay);
  return row ? num(row.TotalScansToday) : 0;
}

function getStationScanDisplay(flowStepOrDisplay, fallback = "") {
  const row = getScanRowForStep(flowStepOrDisplay);
  return row ? safeText(row.DisplayName, fallback) : safeText(fallback || flowStepOrDisplay);
}

function getStationPeakText(flowStepOrDisplay) {
  const row = getScanRowForStep(flowStepOrDisplay);
  if (!row) return "";

  const peakHour = safeText(row.PeakHour, "");
  const peakScans = num(row.PeakHourScans);

  if (!peakHour || peakHour === "—") return "";

  return `Peak: ${formatHourLabel(peakHour)} · ${peakScans.toLocaleString()} scans`;
}


/* =========================================================
   SECTION 15 — FLOW DETAIL TAB
========================================================= */

function renderFlowDetail() {
  const grid = document.getElementById("flowDetailGrid");
  if (!grid) return;

  const rows = (state.surfaceFlow || []).filter(row => {
    if (CONFIG.LINE_A_OFFLINE && isLineA(safeText(row.FlowStep))) return false;
    return true;
  });

  if (!rows.length) {
    grid.innerHTML = `<div class="flow-loading">No flow data available.</div>`;
    return;
  }

  const total = num((state.summary || {}).SurfaceTotalWIP) || 1;
  const maxWip = Math.max(...rows.map(r => num(r.CurrentJobTotal)), 1);

  grid.innerHTML = rows.map((row, index) => {
    const flowStep = safeText(row.FlowStep);
    const displayName = safeText(row.DisplayName || row.FlowStep);
    const group = safeText(row.FlowGroup);
    const wip = num(row.CurrentJobTotal);
    const status = statusFromBackend(row);
    const pctOfTotal = Math.round((wip / total) * 100);
    const barPct = Math.round((wip / maxWip) * 100);
    const icon = iconForStep(flowStep);
    const isLast = index === rows.length - 1;

    const statusColor =
      status === "Critical" ? "var(--red)" :
      status === "Watch"    ? "var(--orange)" :
      status === "ACTIVE"   ? "var(--green)" :
      status === "LOW"      ? "var(--blue)" : "var(--dim)";

    const cardCls =
      status === "Critical" ? "fd-card fd-critical" :
      status === "Watch"    ? "fd-card fd-watch" : "fd-card";

    return `
      <div class="fd-row">
        <article class="${cardCls}" style="animation-delay:${index * 60}ms">
          <div class="fd-card-inner">
            <div class="fd-step-num">${index + 1}</div>

            <div class="fd-icon-wrap" style="color:${statusColor}">
              ${icon}
            </div>

            <div class="fd-body">
              <div class="fd-name">${displayName}</div>
              <div class="fd-group">${group}</div>

              <div class="fd-wip-row">
                <span class="fd-wip-num" style="color:${statusColor}">${wip.toLocaleString()}</span>
                <span class="fd-wip-pct">${pctOfTotal}% of total</span>
                <div class="status-pill ${status}" style="margin-left:auto">${status}</div>
              </div>

              <div class="fd-bar-track">
                <div class="fd-bar-fill fd-bar-${status}" style="--fw:${barPct}%"></div>
              </div>
            </div>
          </div>
        </article>

        ${!isLast ? `<div class="fd-connector"><span class="fd-conn-line"></span><span class="fd-conn-arrow">▼</span></div>` : ""}
      </div>
    `;
  }).join("");

  requestAnimationFrame(() => {
    grid.querySelectorAll(".fd-bar-fill").forEach(el => {
      el.style.width = el.style.getPropertyValue("--fw") || "0%";
    });

    grid.querySelectorAll(".fd-card").forEach(el => {
      el.classList.add("fd-visible");
    });
  });
}


/* =========================================================
   SECTION 16 — LINES TAB
========================================================= */

function renderLines() {
  const grid = document.getElementById("linesGrid");
  if (!grid) return;

  const rows = state.surfaceFlow || [];
  const total = num((state.summary || {}).SurfaceTotalWIP) || 1;

  function isIQStar(row) {
    const step = String(row.FlowStep || "").trim();
    const display = String(row.DisplayName || "").trim();
    return step === "Cooling Storage" || step === "IQ Star" || display === "IQ Star";
  }

  function isSurfaceInspection(row) {
    const step = String(row.FlowStep || "").trim();
    const display = String(row.DisplayName || "").trim();
    return step === "Surface Inspection" || display === "Surface Inspection";
  }

  const lineB = rows.filter(r => isLineB(r.FlowStep) || isIQStar(r));

  const shared = rows.filter(r =>
    isShared(r.FlowStep) &&
    !isIQStar(r) &&
    !isSurfaceInspection(r)
  );

  const inspectionRows = rows.filter(r => isSurfaceInspection(r));

  const lineBTotal = lineB.reduce((sum, r) => sum + num(r.CurrentJobTotal), 0);
  const sharedTotal = shared.reduce((sum, r) => sum + num(r.CurrentJobTotal), 0);
  const inspectionWipTotal = inspectionRows.reduce((sum, r) => sum + num(r.CurrentJobTotal), 0);
  const inspectionOutTotal = getStationScanTotal("Surface Inspection");

  const maxWip = Math.max(lineBTotal, sharedTotal, inspectionWipTotal, 1);

  function laneCard(title, subtitle, cardRows, cardTotal, accentColor) {
    const icon = title.includes("Line B")
      ? ICONS.gear
      : title.includes("Inspection")
      ? ICONS.shield
      : title.includes("Shared")
      ? ICONS.scan
      : ICONS.box;

    const barRows = cardRows.map(row => {
      const wip = num(row.CurrentJobTotal);
      const displayName = safeText(row.DisplayName || row.FlowStep);
      const w = cardTotal > 0 ? Math.round((wip / cardTotal) * 100) : 0;
      const status = statusFromBackend(row);
      const dotColor =
        status === "Critical" ? "var(--red)" :
        status === "Watch"    ? "var(--orange)" :
        wip > 0               ? accentColor : "var(--dim)";

      return `
        <div class="ln-bar-row">
          <div class="ln-dot" style="background:${dotColor};box-shadow:0 0 8px ${dotColor}"></div>
          <div class="ln-step-name">${displayName}</div>
          <div class="ln-bar-track">
            <div class="ln-bar-fill" style="width:${w}%;background:${accentColor};box-shadow:0 0 12px ${accentColor}44"></div>
          </div>
          <div class="ln-val">${wip.toLocaleString()}</div>
        </div>`;
    }).join("");

    const shareOfFacility = Math.round((cardTotal / total) * 100);
    const shareOfMax = Math.round((cardTotal / maxWip) * 100);

    return `
      <article class="ln-card" style="--ln-accent:${accentColor}">
        <div class="ln-header">
          <div class="ln-icon" style="color:${accentColor}">${icon}</div>
          <div class="ln-header-text">
            <div class="ln-title">${title}</div>
            <div class="ln-subtitle">${subtitle}</div>
          </div>
          <div class="ln-total" style="color:${accentColor}">${cardTotal.toLocaleString()}</div>
        </div>

        <div class="ln-share-bar-track">
          <div class="ln-share-bar-fill" style="width:${shareOfMax}%;background:${accentColor};box-shadow:0 0 18px ${accentColor}55"></div>
          <span class="ln-share-label">${shareOfFacility}% of facility WIP</span>
        </div>

        <div class="ln-rows">${barRows}</div>
      </article>`;
  }

  function inspectionReportCard() {
    const shareOfFacility = Math.round((inspectionWipTotal / total) * 100);

    return `
      <article class="ln-card ln-inspection-report" style="--ln-accent:var(--orange)">
        <div class="ln-header">
          <div class="ln-icon" style="color:var(--orange)">${ICONS.shield}</div>
          <div class="ln-header-text">
            <div class="ln-title">Surface Inspection Report</div>
            <div class="ln-subtitle">Inspection output and remaining inspection WIP</div>
          </div>
          <div class="ln-total" style="color:var(--orange)">${inspectionOutTotal.toLocaleString()}</div>
        </div>

        <div class="ln-share-bar-track">
          <div class="ln-share-bar-fill" style="width:${Math.min(100, Math.round((inspectionOutTotal / Math.max(total, 1)) * 100))}%;background:var(--orange);box-shadow:0 0 18px rgba(249,115,22,0.45)"></div>
          <span class="ln-share-label">Inspection OUT today</span>
        </div>

        <div class="ln-rows">
          <div class="ln-bar-row">
            <div class="ln-dot" style="background:var(--orange);box-shadow:0 0 8px var(--orange)"></div>
            <div class="ln-step-name">Surface Inspection OUT</div>
            <div class="ln-bar-track">
              <div class="ln-bar-fill" style="width:${Math.min(100, Math.round((inspectionOutTotal / Math.max(total, 1)) * 100))}%;background:var(--orange);box-shadow:0 0 12px rgba(249,115,22,0.44)"></div>
            </div>
            <div class="ln-val">${inspectionOutTotal.toLocaleString()}</div>
          </div>

          <div class="ln-bar-row">
            <div class="ln-dot" style="background:var(--dim);box-shadow:0 0 8px var(--dim)"></div>
            <div class="ln-step-name">Inspection WIP Remaining</div>
            <div class="ln-bar-track">
              <div class="ln-bar-fill" style="width:${shareOfFacility}%;background:var(--dim);box-shadow:0 0 12px rgba(148,163,184,0.3)"></div>
            </div>
            <div class="ln-val">${inspectionWipTotal.toLocaleString()}</div>
          </div>
        </div>
      </article>`;
  }

  grid.innerHTML = `
    ${laneCard(
      "Shared / Intake",
      "SF Scan · Unbox",
      shared,
      sharedTotal,
      "var(--cyan)"
    )}

    ${laneCard(
      "Line B",
      "IQ Star · Auto Blockers · Generating · Polishing · Engraving · Detaping · Coating",
      lineB,
      lineBTotal,
      "var(--purple)"
    )}

    ${inspectionReportCard()}
  `;
}


/* =========================================================
   SECTION 17 — ALERTS TAB
========================================================= */

function renderAlerts() {
  const panel = document.getElementById("alertsPanel");
  const badge = document.getElementById("alertBadge");

  if (!panel) return;

  const alerts = generateAlerts(state.surfaceFlow || [], state.summary || {});

  if (badge) {
    badge.style.display = alerts.length ? "grid" : "none";
    badge.textContent = alerts.length;
  }

  if (!alerts.length) {
    panel.innerHTML = `
      <div class="alert-card">
        <div>
          <strong>No active WIP alerts</strong>
          <span>All visible stations are below alert thresholds.</span>
        </div>
      </div>
    `;
    return;
  }

  panel.innerHTML = alerts.map(alert => `
    <div class="alert-card">
      <div>
        <strong>${alert.title}</strong>
        <span>${alert.desc}</span>
      </div>
      <span class="sev-pill sev-${alert.severity}">${alert.severity}</span>
      <span class="alert-wip">WIP: ${alert.wip.toLocaleString()}</span>
    </div>
  `).join("");
}

function generateAlerts(rows, summary) {
  const alerts = [];
  const total = num(summary.SurfaceTotalWIP) || 1;

  rows.forEach(row => {
    const flowStep = safeText(row.FlowStep, "");
    const displayName = safeText(row.DisplayName || row.FlowStep);
    const group = safeText(row.FlowGroup);
    const wip = num(row.CurrentJobTotal);

    if (CONFIG.LINE_A_OFFLINE && isLineA(flowStep)) return;

    const pctOfTotal = Math.round((wip / total) * 100);

    if (wip >= CONFIG.WIP_CRITICAL) {
      alerts.push({
        severity: "Critical",
        title: `${displayName} WIP critical (${wip.toLocaleString()})`,
        desc: `${pctOfTotal}% of total Surface WIP — high risk in ${group}.`,
        wip
      });
    } else if (wip >= CONFIG.WIP_HIGH) {
      alerts.push({
        severity: "Watch",
        title: `${displayName} WIP elevated (${wip.toLocaleString()})`,
        desc: `${pctOfTotal}% of total Surface WIP — monitor ${group}.`,
        wip
      });
    }
  });

  rows.forEach(row => {
    const flowStep = safeText(row.FlowStep, "");
    const displayName = safeText(row.DisplayName || row.FlowStep);
    const wip = num(row.CurrentJobTotal);
    const note = String(row.Notes || "");

    if (CONFIG.LINE_A_OFFLINE && isLineA(flowStep)) return;

    if (wip === 0 && note.toLowerCase().includes("missing")) {
      alerts.push({
        severity: "Info",
        title: `${displayName} missing from report`,
        desc: "Forced to 0 for stable workflow order. Verify source data.",
        wip: 0
      });
    }
  });

  const order = { Critical: 0, Watch: 1, Info: 2 };
  alerts.sort((a, b) => (order[a.severity] ?? 9) - (order[b.severity] ?? 9));

  return alerts;
}




/* =========================================================
   SECTION 18 — ANALYTICS TAB
   Professional Optical Operations Analytics
========================================================= */

function renderAnalytics() {
  const grid = document.getElementById("analyticsGrid");
  if (!grid) return;

  const rows = state.surfaceFlow || [];
  const s = state.summary || {};

  const total = num(s.SurfaceTotalWIP);
  const totalSafe = total || 1;

  const mainWip = num(s.SurfaceMainWIP);
  const intakeWip = num(s.SurfaceIntakeWIP);
  const largestStep = safeText(s.LargestStep);
  const largestStepTotal = num(s.LargestStepTotal);
  const inspectionOut = getStationScanTotal("Surface Inspection");

  const bottleneckPct = Math.round((largestStepTotal / totalSafe) * 100);
  const intakePct = Math.round((intakeWip / totalSafe) * 100);
  const mainPct = Math.round((mainWip / totalSafe) * 100);
  const inspectionPct = Math.round((inspectionOut / totalSafe) * 100);

  const criticalRows = rows.filter(row => {
    const step = safeText(row.FlowStep, "");
    return !isLineA(step) && statusFromBackend(row) === "Critical";
  });

  const watchRows = rows.filter(row => {
    const step = safeText(row.FlowStep, "");
    return !isLineA(step) && statusFromBackend(row) === "Watch";
  });

  const topRows = rows
    .filter(row => !isLineA(row.FlowStep))
    .slice()
    .sort((a, b) => num(b.CurrentJobTotal) - num(a.CurrentJobTotal))
    .slice(0, 4);

  const healthStatus =
    bottleneckPct >= 40 || criticalRows.length >= 2 ? "At Risk" :
    bottleneckPct >= 25 || watchRows.length >= 2 ? "Needs Monitoring" :
    "Healthy";

  const healthClass =
    healthStatus === "At Risk" ? "high" :
    healthStatus === "Needs Monitoring" ? "moderate" :
    "controlled";

  const healthScore =
    healthStatus === "At Risk" ? 42 :
    healthStatus === "Needs Monitoring" ? 68 :
    88;

  const healthNarrative =
    healthStatus === "At Risk"
      ? `Surface is currently under pressure. ${largestStep} is holding ${largestStepTotal.toLocaleString()} jobs, which is ${bottleneckPct}% of total Surface WIP. Intake and downstream movement should be watched closely before EOS.`
      : healthStatus === "Needs Monitoring"
      ? `Surface is moving, but ${largestStep} is becoming the main pressure point. Current WIP is still manageable if inspection output continues improving.`
      : `Surface flow is in a healthy range. WIP is distributed well enough to continue normal monitoring.`;

  const recommendation =
    healthStatus === "At Risk"
      ? `${largestStep} needs attention first. Additional support near intake and the next downstream step may help reduce carryover before EOS.`
      : healthStatus === "Needs Monitoring"
      ? `Keep an eye on ${largestStep}. If the next refresh shows growth, consider shifting support before the issue becomes a hard bottleneck.`
      : `No major intervention is needed. Continue watching intake, inspection output, and station balance.`;

  grid.innerHTML = `
    <section class="health-report-shell">

      <article class="health-report-hero ${healthClass}">
        <div>
          <div class="report-eyebrow">Production Health Report</div>
          <h2>Surface Department Health</h2>
          <p>${healthNarrative}</p>
        </div>

        <div class="health-score-card">
          <span>Health Score</span>
          <strong>${healthScore}</strong>
          <small>${healthStatus}</small>
        </div>
      </article>

      <div class="health-report-grid">

        <article class="health-report-card">
          <span>Current WIP Load</span>
          <strong>${total.toLocaleString()}</strong>
          <p>${activeRowsText(rows)} active Surface areas are holding work right now.</p>
        </article>

        <article class="health-report-card">
          <span>Intake Pressure</span>
          <strong>${intakePct}%</strong>
          <p>${intakeWip.toLocaleString()} jobs remain at SF Scan & Verify.</p>
        </article>

        <article class="health-report-card">
          <span>Production Flow</span>
          <strong>${mainPct}%</strong>
          <p>${mainWip.toLocaleString()} jobs are already inside Surface production.</p>
        </article>

        <article class="health-report-card">
          <span>Surface Inspection OUT</span>
          <strong>${inspectionOut.toLocaleString()}</strong>
          <p>${inspectionPct}% output signal against current WIP volume.</p>
        </article>

      </div>

      <article class="health-diagnosis-card">
        <div>
          <div class="report-section-title">Health Diagnosis</div>
          <p>
            The primary constraint is <strong>${largestStep}</strong>, currently holding
            <strong>${largestStepTotal.toLocaleString()}</strong> jobs. This station represents
            <strong>${bottleneckPct}%</strong> of total Surface WIP.
          </p>
        </div>

        <div class="health-recommendation ${healthClass}">
          <span>Recommended Focus</span>
          <strong>${recommendation}</strong>
        </div>
      </article>

      <article class="health-table-card">
        <div class="executive-panel-head">
          <div>
            <div class="report-section-title">Stations Impacting Health</div>
            <p class="report-section-sub">Top areas affecting the current health status.</p>
          </div>
          <div class="report-chip">${topRows.length} stations</div>
        </div>

        <div class="executive-top-list">
          ${topRows.map((row, index) => {
            const displayName = safeText(row.DisplayName || row.FlowStep);
            const wip = num(row.CurrentJobTotal);
            const status = statusFromBackend(row);
            const pctOfTotal = total > 0 ? Math.round((wip / total) * 100) : 0;

            const statusClass =
              status === "Critical" ? "critical" :
              status === "Watch" ? "watch" :
              "";

            return `
              <div class="executive-step-row ${statusClass}">
                <div class="executive-rank">${index + 1}</div>
                <div class="executive-step-body">
                  <strong>${displayName}</strong>
                  <span>${status} · ${pctOfTotal}% of Surface WIP</span>
                </div>
                <div class="executive-step-value">${wip.toLocaleString()}</div>
              </div>
            `;
          }).join("")}
        </div>
      </article>

    </section>
  `;
}

function activeRowsText(rows) {
  return rows.filter(row => !isLineA(row.FlowStep) && num(row.CurrentJobTotal) > 0).length;
}


/* =========================================================
   SECTION 19 — SUMMARY TAB
   Executive Optical Operations Summary
========================================================= */

function renderReports() {
  const wrap = document.getElementById("reportTableWrap");
  if (!wrap) return;

  const rows = state.surfaceFlow || [];
  const s = state.summary || {};

  if (!rows.length) {
    wrap.innerHTML = `<div class="flow-loading">No Surface summary data available.</div>`;
    return;
  }

  const total = num(s.SurfaceTotalWIP);
  const totalSafe = total || 1;

  const mainWip = num(s.SurfaceMainWIP);
  const intakeWip = num(s.SurfaceIntakeWIP);
  const largestStep = safeText(s.LargestStep);
  const largestStepTotal = num(s.LargestStepTotal);
  const inspectionTotal = getStationScanTotal("Surface Inspection");

  const bottleneckPct = Math.round((largestStepTotal / totalSafe) * 100);
  const mainPct = Math.round((mainWip / totalSafe) * 100);
  const intakePct = Math.round((intakeWip / totalSafe) * 100);

  const activeRows = rows.filter(row => !isLineA(row.FlowStep) && num(row.CurrentJobTotal) > 0);

  const criticalRows = rows.filter(row => {
    const step = safeText(row.FlowStep, "");
    return !isLineA(step) && statusFromBackend(row) === "Critical";
  });

  const watchRows = rows.filter(row => {
    const step = safeText(row.FlowStep, "");
    return !isLineA(step) && statusFromBackend(row) === "Watch";
  });

  const topRows = rows
    .filter(row => !isLineA(row.FlowStep))
    .slice()
    .sort((a, b) => num(b.CurrentJobTotal) - num(a.CurrentJobTotal))
    .slice(0, 5);

  const secondStep = topRows[1];
  const secondStepName = secondStep ? safeText(secondStep.DisplayName || secondStep.FlowStep) : "—";
  const secondStepTotal = secondStep ? num(secondStep.CurrentJobTotal) : 0;

  const hourlyRows = Array.isArray(state.surfaceHourlyInOut)
    ? state.surfaceHourlyInOut
    : [];

  const hourlyOutputTotal = hourlyRows.reduce((sum, row) => {
    return sum + num(row.Output || row.Out || row.TotalOut || row.TotalOutput || row.ScanOut || 0);
  }, 0);

  const hourlyInputTotal = hourlyRows.reduce((sum, row) => {
    return sum + num(row.Input || row.In || row.TotalIn || row.TotalInput || row.ScanIn || 0);
  }, 0);

  const outputSignal = inspectionTotal || hourlyOutputTotal;
  const netBuild = Math.max(0, total - outputSignal);

  const releasePct = total > 0 ? Math.round((outputSignal / total) * 100) : 0;

  const eosRisk =
    bottleneckPct >= 40 || criticalRows.length >= 2 ? "High" :
    bottleneckPct >= 25 || watchRows.length >= 2 ? "Moderate" :
    "Controlled";

  const eosClass =
    eosRisk === "High" ? "high" :
    eosRisk === "Moderate" ? "moderate" :
    "controlled";

  const readiness =
    releasePct >= 45 ? "Strong" :
    releasePct >= 25 ? "Developing" :
    "Behind";

  const readinessClass =
    readiness === "Strong" ? "controlled" :
    readiness === "Developing" ? "moderate" :
    "high";

  const daySummary =
    `Surface is carrying ${total.toLocaleString()} jobs across ${activeRows.length} active WIP areas. ` +
    `${mainWip.toLocaleString()} jobs (${mainPct}%) are already inside production, while ` +
    `${intakeWip.toLocaleString()} jobs (${intakePct}%) remain at intake. ` +
    `The main constraint is ${largestStep} with ${largestStepTotal.toLocaleString()} jobs, representing ${bottleneckPct}% of total Surface WIP.`;

  const outputSummary =
    outputSignal > 0
      ? `Surface Inspection shows ${outputSignal.toLocaleString()} completed scan activity today, equal to ${releasePct}% of current WIP volume.`
      : `Hourly output data is not available yet, so EOS projection is based on current WIP and station concentration only.`;

  const eosSummary =
    eosRisk === "High"
      ? `EOS carryover risk is high. If the current constraint is not reduced, Surface will likely carry a heavy intake and production backlog into the next shift.`
      : eosRisk === "Moderate"
      ? `EOS carryover risk is moderate. The shift can recover if support is moved to the constraint and downstream stations keep releasing work.`
      : `EOS carryover risk is controlled. Current WIP distribution does not show a severe end-of-shift backlog pattern.`;

  const leadershipAction =
  eosRisk === "High"
    ? `${largestStep} is currently carrying most of the Surface WIP. Additional support might requiered in ${secondStepName} to stabilize flow.`
    : eosRisk === "Moderate"
    ? `${largestStep} should continue to be monitored through the next refresh cycle. If WIP continues building, consider shifting additional support downstream.`
    : `Surface flow is currently stable. Continue monitoring inspection output and intake movement through the remainder of the shift.`;

  wrap.innerHTML = `
    <section class="daily-summary-shell">

      <article class="daily-brief-card">
        <div class="daily-brief-head">
          <div>
            <div class="report-eyebrow">Daily Executive Summary</div>
            <h2>Surface Shift Performance Read</h2>
            <p>${daySummary}</p>
          </div>

          <div class="daily-risk-tile ${eosClass}">
            <span>EOS Risk</span>
            <strong>${eosRisk}</strong>
          </div>
        </div>

        <div class="daily-summary-grid">
          <div class="daily-summary-metric">
            <span>Current WIP Position</span>
            <strong>${total.toLocaleString()}</strong>
            <small>${activeRows.length} active areas · ${criticalRows.length} critical · ${watchRows.length} watch</small>
          </div>

          <div class="daily-summary-metric">
            <span>Production WIP</span>
            <strong>${mainWip.toLocaleString()}</strong>
            <small>${mainPct}% already inside Surface flow</small>
          </div>

          <div class="daily-summary-metric">
            <span>Intake Load</span>
            <strong>${intakeWip.toLocaleString()}</strong>
            <small>${intakePct}% still at SF Scan & Verify</small>
          </div>

          <div class="daily-summary-metric">
            <span>Output / Release Signal</span>
            <strong>${outputSignal.toLocaleString()}</strong>
            <small>${releasePct}% release signal against current WIP</small>
          </div>

          <div class="daily-summary-metric">
            <span>Constraint Impact</span>
            <strong>${bottleneckPct}%</strong>
            <small>${largestStep} · ${largestStepTotal.toLocaleString()} jobs</small>
          </div>

        <div class="daily-summary-metric">
  <span>Surface Inspection OUT</span>
  <strong>${outputSignal.toLocaleString()}</strong>
  <small>
    ${releasePct}% of current Surface WIP has reached inspection output today.
  </small>
</div>
        </div>
      </article>

      <article class="daily-narrative-card">
        <div>
          <div class="report-section-title">Professional Shift Readout</div>
          <p>${outputSummary}</p>
          <p>${eosSummary}</p>
        </div>

        <div class="daily-action-box ${eosClass}">
          <span>Leadership Action</span>
          <strong>${leadershipAction}</strong>
        </div>
      </article>

      <div class="daily-insight-grid">

        <article class="daily-panel">
          <div class="executive-panel-head">
            <div>
              <div class="report-section-title">EOS Carryover Watch</div>
              <p class="report-section-sub">Highest risk areas to reduce before end of shift.</p>
            </div>
            <div class="report-chip">${topRows.length} areas</div>
          </div>

          <div class="executive-top-list">
            ${topRows.map((row, index) => {
              const flowStep = safeText(row.FlowStep, "");
              const displayName = safeText(row.DisplayName || row.FlowStep);
              const wip = num(row.CurrentJobTotal);
              const status = statusFromBackend(row);
              const percent = total > 0 ? Math.round((wip / total) * 100) : 0;

              const statusClass =
                status === "Critical" ? "critical" :
                status === "Watch" ? "watch" :
                "";

              return `
                <div class="executive-step-row ${statusClass}">
                  <div class="executive-rank">${index + 1}</div>

                  <div class="executive-step-body">
                    <strong>${displayName}</strong>
                    <span>${flowStep} · ${status} · ${percent}% of total WIP</span>
                  </div>

                  <div class="executive-step-value">${wip.toLocaleString()}</div>
                </div>
              `;
            }).join("")}
          </div>
        </article>

        <article class="daily-panel">
          <div class="executive-panel-head">
            <div>
              <div class="report-section-title">Across-the-Board Read</div>
              <p class="report-section-sub">Simple leadership interpretation of the current shift.</p>
            </div>
            <div class="report-chip">${readiness}</div>
          </div>

          <div class="daily-read-list">
            <div>
              <span>Primary Bottleneck</span>
              <strong>${largestStep}</strong>
              <small>${largestStepTotal.toLocaleString()} jobs are concentrated here.</small>
            </div>

            <div>
              <span>Secondary Pressure</span>
              <strong>${secondStepName}</strong>
              <small>${secondStepTotal.toLocaleString()} jobs sitting behind the primary constraint.</small>
            </div>

            <div>
              <span>Hourly Output Signal</span>
              <strong>${outputSignal.toLocaleString()}</strong>
              <small>Uses Surface Inspection scan activity when hourly output is not available.</small>
            </div>

            <div>
              <span>Estimated Remaining Pressure</span>
              <strong>${netBuild.toLocaleString()}</strong>
              <small>Current WIP minus release signal. Use as a rough EOS pressure indicator.</small>
            </div>
          </div>
        </article>

      </div>

    </section>
  `;
}

/* =========================================================
   SECTION 20 — EXPORT CSV
========================================================= */

function exportCSV() {
  const rows = state.surfaceFlow || [];
  if (!rows.length) return;

  const total = num(state.summary?.SurfaceTotalWIP) || 1;

  const headers = [
    "FlowOrder",
    "FlowStep",
    "DisplayName",
    "FlowGroup",
    "CurrentJobTotal",
    "PercentOfTotal",
    "Status",
    "Notes"
  ];

  const csvRows = rows.map(row => {
    const flowStep = safeText(row.FlowStep, "");
    const displayName = safeText(row.DisplayName || row.FlowStep, "");
    const wip = num(row.CurrentJobTotal);
    const percent = Math.round((wip / total) * 100);
    const status = CONFIG.LINE_A_OFFLINE && isLineA(flowStep) ? "OFFLINE" : statusFromBackend(row);

    return [
      num(row.FlowOrder),
      `"${flowStep}"`,
      `"${displayName}"`,
      `"${safeText(row.FlowGroup, "")}"`,
      wip,
      `"${percent}%"`,
      `"${status}"`,
      `"${safeText(row.Notes, "")}"`
    ].join(",");
  });

  const csv = [headers.join(","), ...csvRows].join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const date = new Date().toISOString().slice(0, 10);

  link.href = url;
  link.download = `Surface_WIP_Flow_${date}.csv`;
  link.click();

  URL.revokeObjectURL(url);
}


/* =========================================================
   SECTION 21 — CLOCK
========================================================= */

function updateClock() {
  const now = new Date();

  const timeEl = document.getElementById("clockTime");
  const dateEl = document.getElementById("clockDate");

  if (timeEl) {
    timeEl.textContent = now.toLocaleTimeString([], {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
  }

  if (dateEl) {
    dateEl.textContent = now.toLocaleDateString([], {
      month: "short",
      day: "numeric",
      year: "numeric"
    });
  }
}


/* =========================================================
   SECTION 22 — EVENTS
========================================================= */

function initTabs() {
  const navItems = document.querySelectorAll(".nav-item[data-tab]");

  navItems.forEach(button => {
    button.addEventListener("click", () => {
      const tab = button.dataset.tab;

      navItems.forEach(btn => btn.classList.remove("active"));
      button.classList.add("active");

      document.querySelectorAll(".tab-panel").forEach(panel => {
        panel.classList.remove("active");
      });

      const target = document.getElementById("tab-" + tab);

      if (target) {
        target.classList.add("active");
      }

      state.activeTab = tab;
    });
  });
}

function initFilters() {
  const buttons = document.querySelectorAll(".filter-btn[data-filter]");

  buttons.forEach(button => {
    button.addEventListener("click", () => {
      buttons.forEach(btn => btn.classList.remove("active"));
      button.classList.add("active");

      state.filter = button.dataset.filter;
      renderFlowGrid();
      renderTransferTunnel();
    });
  });
}

function initExport() {
  const button = document.getElementById("btnExportCsv");

  if (button) {
    button.addEventListener("click", exportCSV);
  }
}


/* =========================================================
   SECTION 23 — BOOT
========================================================= */

function boot() {
  startSurfaceLoaderAnimation();

  initTabs();
  initFilters();
  initExport();

  updateClock();
  setInterval(updateClock, 1000);

  fetchData(true);

  setInterval(() => {
    fetchData(false);
  }, CONFIG.REFRESH_MS);
}

/* =================================================
   SIDEBAR TAB ACTIVE STATE
   Add only if your existing tab click logic does not already do this
================================================= */

document.querySelectorAll(".nav-item").forEach(button => {
  button.addEventListener("click", () => {
    const tab = button.dataset.tab;

    document.querySelectorAll(".nav-item").forEach(item => {
      item.classList.remove("active");
    });

    button.classList.add("active");

    document.querySelectorAll(".tab-panel").forEach(panel => {
      panel.classList.remove("active");
    });

    const targetPanel = document.getElementById(tab);
    if (targetPanel) {
      targetPanel.classList.add("active");
    }
  });
});

document.addEventListener("DOMContentLoaded", boot);

function goBackToDashboard() {
  window.location.href = "index.html";
}

