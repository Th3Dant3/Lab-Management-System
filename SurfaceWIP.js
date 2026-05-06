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
  API_URL: "https://script.google.com/macros/s/AKfycbzXB24Hq39ymhHnh5QYtHsGsRTdfUNtVq3gxx1IPDxmpV7uuzqXLLfRHP29ypwPIgl1ng/exec",
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

async function fetchData(forceRender = false) {
  setSystemStatus("loading");

  try {
    const res = await fetch(CONFIG.API_URL + "?debug=true&t=" + Date.now(), {
      cache: "no-store"
    });

    if (!res.ok) {
      throw new Error("HTTP " + res.status);
    }

    const json = await res.json();

    if (!json || json.status !== "success") {
      throw new Error(json?.message || "API returned error");
    }

    const incomingFlow = Array.isArray(json.surfaceFlow) ? json.surfaceFlow : [];
    const incomingTransfers = Array.isArray(json.surfaceTransfers) ? json.surfaceTransfers : [];
    const incomingScanSummary = Array.isArray(json.surfaceScanSummary)
      ? json.surfaceScanSummary
      : [];

    if (!incomingFlow.length && state.hasRenderedOnce) {
      console.warn("[SurfaceWIP] Refresh returned empty surfaceFlow. Keeping last good data.");
      setSystemStatus("ok");
      updateLatestUpdatePill();
      return;
    }

    if (!incomingFlow.length && !state.hasRenderedOnce) {
      throw new Error("No Surface WIP rows returned from API.");
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

  const maxWip = Math.max(...rows.map(r => num(r.CurrentJobTotal)), 1);

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

  const severity =
    totalMoving >= 300 ? "critical" :
    totalMoving >= 100 ? "watch" :
    "normal";

  /*
    2 transitions per row =:
    Station → Conveyor → Station → Conveyor → Station

    Next row does NOT repeat the previous ending station.
  */
  const transitionsPerRow = 2;
  const snakeRows = buildSnakeRowsByTransitions(orderedRows, transitionsPerRow);
  window.__surfaceSnakeRows = snakeRows;

  const stationCount = (() => {
    const seen = new Set();

    orderedRows.forEach((row, index) => {
      if (index === 0) seen.add(String(row.FromStep || "").trim());
      seen.add(String(row.ToStep || "").trim());
    });

    return seen.size;
  })();

  return `
    <section class="continuous-flow-shell ${severity}">
      <div class="continuous-flow-header">
        <div>
          <div class="continuous-flow-title">Surface Live Flow</div>
          <div class="continuous-flow-subtitle">
            Total Scan Today → Current WIP Up Now → Next Station
          </div>
        </div>

        <div class="continuous-flow-metrics">
          <div>
            <span>Total Current WIP Up Now</span>
            <strong>${totalMoving.toLocaleString()}</strong>
          </div>

          <div>
            <span>Stations</span>
            <strong>${stationCount}</strong>
          </div>
        </div>
      </div>

      <div class="snake-flow-wrap">
        ${snakeRows.map((row, rowIndex) => renderSnakeRow(row, rowIndex)).join("")}
      </div>
    </section>
  `;
}

function buildSnakeRowsByTransitions(orderedRows, transitionsPerRow) {
  const rows = [];

  for (let i = 0; i < orderedRows.length; i += transitionsPerRow) {
    rows.push({
      reverse: rows.length % 2 === 1,
      transitions: orderedRows.slice(i, i + transitionsPerRow)
    });
  }

  return rows;
}


/* =========================================================
   SECTION 13 — LIVE SNAKE FLOW ROW RENDERERS
   NEW STUFF: Reverse rows skip repeated starting station.
========================================================= */

function renderSnakeRow(row, rowIndex) {
  const transitions = row.transitions || [];
  const pieces = [];

  if (!transitions.length) return "";

  if (!row.reverse) {
    /*
      Normal row:
      Station → Conveyor → Station → Conveyor → Station
    */
    transitions.forEach((transition, index) => {
      if (index === 0) {
        pieces.push(renderContinuousStationCard({
          step: transition.FromStep,
          display: transition.FromDisplayName
        }));
      }

      pieces.push(renderContinuousConveyor(transition));

      pieces.push(renderContinuousStationCard({
        step: transition.ToStep,
        display: transition.ToDisplayName
      }));
    });
  } else {
    /*
      Reverse row:
      Do NOT repeat the previous row's ending station.

      Example:
      Previous row ended on Auto Blockers.
      Next row starts with conveyor Auto Blockers → IQ Star,
      then shows IQ Star, then next conveyor, then Orbit Generator.
    */
    const reversedTransitions = transitions.slice().reverse();

    reversedTransitions.forEach(transition => {
      pieces.push(renderContinuousConveyor(transition));

      pieces.push(renderContinuousStationCard({
        step: transition.FromStep,
        display: transition.FromDisplayName
      }));
    });
  }

  const nextFirstTransition = getNextRowFirstTransition(rowIndex);

  return `
    <div class="snake-row ${row.reverse ? "reverse no-repeat-start" : ""}" data-row-index="${rowIndex}">
      ${pieces.join("")}
    </div>

    ${
      nextFirstTransition
        ? renderSnakeBridge(nextFirstTransition, row.reverse)
        : ""
    }
  `;
}

function getNextRowFirstTransition(rowIndex) {
  const allRows = window.__surfaceSnakeRows || [];
  const nextRow = allRows[rowIndex + 1];

  if (!nextRow || !nextRow.transitions || !nextRow.transitions.length) {
    return null;
  }

  return nextRow.transitions[0];
}

function renderSnakeBridge(row, previousReverse) {
  const wip = calcEstimatedMoving(row);

  const severity =
    wip >= 200 ? "critical" :
    wip >= 75 ? "watch" :
    "normal";

  const nextDirectionArrow = previousReverse ? "→" : "←";

  return `
    <div class="snake-bridge ${previousReverse ? "left" : "right"} ${severity}">
      <div class="snake-bridge-path">
        <div class="snake-bridge-down">
          <span class="snake-bridge-arrow arrow-down">↓</span>
        </div>

        <div class="snake-bridge-turn">
          <span class="snake-bridge-arrow arrow-turn">${nextDirectionArrow}</span>
        </div>
      </div>

      <div class="snake-bridge-info compact">
        <div class="snake-bridge-label">Next Flow</div>
        <div class="snake-bridge-count">${wip.toLocaleString()}</div>
        <div class="snake-bridge-route">
          ${row.FromDisplayName} → ${row.ToDisplayName}
        </div>
      </div>
    </div>
  `;
}

function renderContinuousStationCard(station) {
  const scanTotal = getStationScanTotal(station.step);
  const display = getStationScanDisplay(station.step, station.display);
  const peak = getStationPeakText(station.step);

  return `
    <article class="continuous-station-card">
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
    wip >= 200 ? "critical" :
    wip >= 75 ? "watch" :
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
        <div class="transfer-conveyor-title">Current WIP Up Now</div>
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

  return `Peak: ${peakHour} · ${peakScans.toLocaleString()} scans`;
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

  const lineB = rows.filter(r => isLineB(r.FlowStep));
  const shared = rows.filter(r => isShared(r.FlowStep));

  const lineBTotal = lineB.reduce((sum, r) => sum + num(r.CurrentJobTotal), 0);
  const sharedTotal = shared.reduce((sum, r) => sum + num(r.CurrentJobTotal), 0);

  const maxWip = Math.max(lineBTotal, sharedTotal, 1);

  function laneCard(title, subtitle, cardRows, cardTotal, accentColor) {
    const icon = title.includes("Line B")
      ? ICONS.gear
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

  grid.innerHTML = `
    ${laneCard("Shared / Intake", "SF Scan · Unbox · Cooling · Inspection", shared, sharedTotal, "var(--cyan)")}
    ${laneCard("Line B", "Active production line", lineB, lineBTotal, "var(--purple)")}
  `;

  requestAnimationFrame(() => {
    grid.querySelectorAll(".ln-bar-fill").forEach(el => {
      const target = el.style.width;
      el.style.width = "0%";
      requestAnimationFrame(() => { el.style.width = target; });
    });

    grid.querySelectorAll(".ln-share-bar-fill").forEach(el => {
      const target = el.style.width;
      el.style.width = "0%";
      requestAnimationFrame(() => { el.style.width = target; });
    });
  });
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
========================================================= */

/* =========================================================
   SECTION 18 — ANALYTICS TAB
   Professional Optical Operations Analytics
========================================================= */

function renderAnalytics() {
  const grid = document.getElementById("analyticsGrid");
  if (!grid) return;

  const rows = state.surfaceFlow || [];
  const s = state.summary || {};
  const total = num(s.SurfaceTotalWIP) || 1;

  const groups = {};

  rows.forEach(row => {
    const group = safeText(row.FlowGroup, "Other");
    groups[group] = (groups[group] || 0) + num(row.CurrentJobTotal);
  });

  const entries = Object.entries(groups).sort((a, b) => b[1] - a[1]);
  const max = Math.max(...entries.map(([, value]) => value), 1);

  const largestStep = safeText(s.LargestStep);
  const largestStepTotal = num(s.LargestStepTotal);
  const largestPct = Math.round((largestStepTotal / total) * 100);

  const mainWip = num(s.SurfaceMainWIP);
  const intakeWip = num(s.SurfaceIntakeWIP);
  const inspectionTotal = getStationScanTotal("Surface Inspection");

  const criticalRows = rows.filter(row => statusFromBackend(row) === "Critical");
  const watchRows = rows.filter(row => statusFromBackend(row) === "Watch");

  const operationalRisk =
    criticalRows.length > 0 ? "High" :
    watchRows.length >= 3 ? "Moderate" :
    "Controlled";

  const recommendedAction =
    criticalRows.length > 0
      ? `Immediate support should be directed toward ${largestStep}. This area is carrying the highest load and may restrict downstream movement.`
      : watchRows.length > 0
      ? `Monitor the elevated stations and rebalance labor before the next report cycle.`
      : `Surface flow is currently stable. Continue monitoring intake, inspection, and downstream movement.`;

  grid.innerHTML = `
    <section class="optical-analytics-report">

      <article class="optical-report-hero">
        <div>
          <div class="report-eyebrow">Optical Operations Analytics</div>
          <h2>Surface Production Health Report</h2>
          <p>
            This report evaluates the current Surface workflow using live WIP, scan activity,
            bottleneck concentration, and station-level distribution. The goal is to identify
            where jobs are accumulating and where operational support should be focused.
          </p>
        </div>

        <div class="optical-risk-card">
          <span>Operational Risk</span>
          <strong class="${operationalRisk.toLowerCase()}">${operationalRisk}</strong>
        </div>
      </article>

      <div class="optical-insight-grid">

        <article class="optical-insight-card">
          <div class="insight-label">Total Surface WIP</div>
          <div class="insight-value">${total.toLocaleString()}</div>
          <p>
            Current jobs staged within the Surface workflow. This includes intake,
            active production areas, and downstream Surface stations.
          </p>
        </article>

        <article class="optical-insight-card">
          <div class="insight-label">Main Surface WIP</div>
          <div class="insight-value">${mainWip.toLocaleString()}</div>
          <p>
            Jobs beyond initial intake. This is the strongest indicator of workload
            already inside the Surface production path.
          </p>
        </article>

        <article class="optical-insight-card warning">
          <div class="insight-label">Primary Constraint</div>
          <div class="insight-value">${largestStep}</div>
          <p>
            ${largestStepTotal.toLocaleString()} jobs are currently concentrated here,
            representing ${largestPct}% of total Surface WIP.
          </p>
        </article>

        <article class="optical-insight-card">
          <div class="insight-label">Inspection Scan Activity</div>
          <div class="insight-value">${inspectionTotal.toLocaleString()}</div>
          <p>
            Total Surface Inspection scan activity. This helps indicate downstream
            release movement after jobs complete the Surface path.
          </p>
        </article>

      </div>

      <article class="optical-narrative-card">
        <div class="report-section-title">Operational Interpretation</div>
        <p>
          Surface is currently showing <strong>${total.toLocaleString()}</strong> jobs in WIP.
          The largest workload concentration is <strong>${largestStep}</strong> with
          <strong>${largestStepTotal.toLocaleString()}</strong> jobs. When one Surface station
          carries a large share of total WIP, it can slow the flow into downstream steps such as
          coating, inspection, and release.
        </p>

        <p>
          Intake currently shows <strong>${intakeWip.toLocaleString()}</strong> jobs at SF Scan.
          Comparing intake against main Surface WIP helps separate new arrivals from jobs already
          moving through blocking, generating, polishing, coating, and inspection.
        </p>

        <div class="recommendation-box">
          <span>Recommended Action</span>
          <strong>${recommendedAction}</strong>
        </div>
      </article>

      <article class="optical-distribution-card">
        <div class="report-section-title">WIP Distribution by Flow Group</div>
        <p class="report-section-sub">
          This view ranks Surface workload by process group to show where jobs are accumulating.
        </p>

        <div class="group-bars optical-bars">
          ${entries.map(([name, value]) => {
            const percent = Math.round((value / total) * 100);
            return `
              <div class="group-bar-row">
                <div class="gb-name">${name}</div>
                <div class="gb-track">
                  <div class="gb-fill" style="width:${Math.round((value / max) * 100)}%"></div>
                </div>
                <div class="gb-val">${value.toLocaleString()} <span>${percent}%</span></div>
              </div>
            `;
          }).join("")}
        </div>
      </article>

    </section>
  `;
}


/* =========================================================
   SECTION 19 — SUMMARY TAB
========================================================= */

function renderReports() {
  const wrap = document.getElementById("reportTableWrap");
  if (!wrap) return;

  const rows = state.surfaceFlow || [];
  const s = state.summary || {};

  if (!rows.length) {
    wrap.innerHTML = `<div class="flow-loading">No summary data available.</div>`;
    return;
  }

  const total = num(s.SurfaceTotalWIP);
  const main = num(s.SurfaceMainWIP);
  const intake = num(s.SurfaceIntakeWIP);
  const largestStep = safeText(s.LargestStep);
  const largestStepTotal = num(s.LargestStepTotal);

  const visibleRows = rows.filter(row => {
    const flowStep = safeText(row.FlowStep, "");
    return !(CONFIG.LINE_A_OFFLINE && isLineA(flowStep));
  });

  const criticalRows = visibleRows.filter(row => statusFromBackend(row) === "Critical");
  const watchRows = visibleRows.filter(row => statusFromBackend(row) === "Watch");
  const activeRows = visibleRows.filter(row => statusFromBackend(row) === "ACTIVE");
  const lowRows = visibleRows.filter(row => statusFromBackend(row) === "LOW");
  const emptyRows = visibleRows.filter(row => num(row.CurrentJobTotal) === 0);

  const topSteps = visibleRows
    .slice()
    .sort((a, b) => num(b.CurrentJobTotal) - num(a.CurrentJobTotal))
    .slice(0, 5);

  const groups = {};

  visibleRows.forEach(row => {
    const group = safeText(row.FlowGroup, "Other");
    groups[group] = (groups[group] || 0) + num(row.CurrentJobTotal);
  });

  const groupEntries = Object.entries(groups).sort((a, b) => b[1] - a[1]);
  const maxGroup = Math.max(...groupEntries.map(([, value]) => value), 1);

  const liveRows = buildLiveTransferRowsFromSurfaceFlow_();
  const totalCurrentUpNow = liveRows.reduce((sum, row) => sum + num(row.EstimatedInTransfer), 0);

  wrap.innerHTML = `
    <section class="report-summary-shell">

      <div class="report-hero-grid">

        <article class="report-hero-card">
          <div class="report-hero-content">
            <div class="report-eyebrow">Surface WIP Summary</div>
            <div class="report-hero-title">${total.toLocaleString()}</div>

            <p class="report-hero-sub">
           Executive summary of the current Surface operation.</p>

            <div class="report-hero-metrics">
              <div class="report-mini-metric">
                <div class="report-mini-label">Main WIP</div>
                <div class="report-mini-value">${main.toLocaleString()}</div>
                <div class="report-mini-note">Jobs already inside Surface production</div>
              </div>

              <div class="report-mini-metric">
                <div class="report-mini-label">Intake WIP</div>
                <div class="report-mini-value">${intake.toLocaleString()}</div>
                <div class="report-mini-note">New intake entering Surface</div>
              </div>

              <div class="report-mini-metric">
                <div class="report-mini-label">Current Up Now</div>
                <div class="report-mini-value">${totalCurrentUpNow.toLocaleString()}</div>
                <div class="report-mini-note">Estimated workload feeding next steps</div>
              </div>
            </div>
          </div>
        </article>

        <article class="report-status-panel">
          <div class="report-panel-title">Operational Status</div>

          <div class="report-status-list">
            <div class="report-status-row">
              <div>
                <div class="report-status-name">Critical Steps</div>
                <div class="report-status-sub">Visible stations above critical threshold</div>
              </div>
              <div class="report-status-count">${criticalRows.length}</div>
            </div>

            <div class="report-status-row">
              <div>
                <div class="report-status-name">Watch Steps</div>
                <div class="report-status-sub">Elevated WIP requiring monitoring</div>
              </div>
              <div class="report-status-count">${watchRows.length}</div>
            </div>

            <div class="report-status-row">
              <div>
                <div class="report-status-name">Active Steps</div>
                <div class="report-status-sub">Steps currently holding work</div>
              </div>
              <div class="report-status-count">${activeRows.length + lowRows.length}</div>
            </div>

            <div class="report-status-row">
              <div>
                <div class="report-status-name">Empty Visible Steps</div>
                <div class="report-status-sub">Visible non-Line-A steps at zero</div>
              </div>
              <div class="report-status-count">${emptyRows.length}</div>
            </div>
          </div>
        </article>

      </div>

      <div class="report-section-grid">

        <article class="report-section-card">
          <div class="report-section-head">
            <div>
              <div class="report-section-title">Top WIP Stations</div>
              <div class="report-section-sub">Highest current WIP by visible station.</div>
            </div>

            <div class="report-chip">
              ${largestStep}: ${largestStepTotal.toLocaleString()}
            </div>
          </div>

          <div class="report-list">
            ${topSteps.map((row, index) => {
              const wip = num(row.CurrentJobTotal);
              const displayName = safeText(row.DisplayName || row.FlowStep);
              const flowStep = safeText(row.FlowStep);
              const group = safeText(row.FlowGroup);
              const status = statusFromBackend(row);

              return `
                <div class="report-step-row">
                  <div class="report-rank">${index + 1}</div>

                  <div>
                    <div class="report-step-name">${displayName}</div>
                    <div class="report-step-meta">${flowStep} · ${group} · ${status}</div>
                  </div>

                  <div class="report-step-value">${wip.toLocaleString()}</div>
                </div>
              `;
            }).join("")}
          </div>
        </article>

        <article class="report-section-card">
          <div class="report-section-head">
            <div>
              <div class="report-section-title">WIP by Group</div>
              <div class="report-section-sub">Group-level distribution from visible Surface flow.</div>
            </div>

            <div class="report-chip">${groupEntries.length} groups</div>
          </div>

          <div class="report-list">
            ${groupEntries.map(([name, value]) => {
              const width = Math.round((value / maxGroup) * 100);

              return `
                <div class="report-bar-row">
                  <div class="report-bar-name">${name}</div>

                  <div class="report-bar-track">
                    <div class="report-bar-fill" style="width:${width}%"></div>
                  </div>

                  <div class="report-bar-value">${value.toLocaleString()}</div>
                </div>
              `;
            }).join("")}
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

document.addEventListener("DOMContentLoaded", boot);

function goBackToDashboard() {
  window.location.href = "index.html";
}