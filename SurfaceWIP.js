/* =================================================
   SurfaceWIP — Frontend JS
   Production Surface WIP + Transfer Tunnel + Daily Totals
================================================= */

"use strict";

/* =============================================
   SAFE LOADING SCREEN HELPERS
   These prevent JS from breaking if loader HTML is missing
============================================= */

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

  if (!loader) {
    return;
  }

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

  if (!loader) {
    return;
  }

  const steps = [
    { pct: 18, msg: "Authenticating Surface access…" },
    { pct: 38, msg: "Loading WIP stations…" },
    { pct: 58, msg: "Reading transfer snapshots…" },
    { pct: 76, msg: "Building movement tunnel…" },
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

/* =============================================
   CONFIG
============================================= */

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

/* =============================================
   ICONS
============================================= */

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

/* =============================================
   STATE
============================================= */

let state = {
  summary: null,
  surfaceFlow: [],
  surfaceTransfers: [],
  surfaceTransferDailyTotals: [],
  intervalMeta: null,
  lastFetch: null,
  lastSnapshotKey: null,
  lastDataSignature: null,
  filter: "all",
  activeTab: "overview",
  hasRenderedOnce: false
};

/* =============================================
   HELPERS
============================================= */

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

function findDailyTransferForTransition(transitionName) {
  const rows = state.surfaceTransferDailyTotals || [];

  return rows.find(row => {
    return String(row.TransitionName || "").trim() === String(transitionName || "").trim();
  }) || null;
}

function buildDataSignature(payload) {
  const summary = payload.summary || {};
  const meta = payload.intervalHistoryMeta || {};

  return [
    summary.ReportDate || "",
    summary.SourceFile || "",
    summary.FileUpdatedTime || "",
    summary.ImportedAt || "",
    meta.currentKey || "",
    meta.currentTime || "",
    (payload.surfaceFlow || []).length,
    (payload.surfaceTransfers || []).length,
    (payload.surfaceTransferDailyTotals || []).length
  ].join("|");
}

function hasUsablePayload(payload) {
  if (!payload || payload.status !== "success") return false;

  const flowRows = payload.surfaceFlow || [];

  if (!flowRows.length) return false;

  return true;
}

/* =============================================
   API FETCH
============================================= */

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
    const incomingDailyTotals = Array.isArray(json.surfaceTransferDailyTotals)
      ? json.surfaceTransferDailyTotals
      : [];

    /*
     * HARD SAFETY:
     * If the API temporarily returns no Surface rows,
     * DO NOT wipe the screen.
     */
    if (!incomingFlow.length && state.hasRenderedOnce) {
      console.warn("[SurfaceWIP] Refresh returned empty surfaceFlow. Keeping last good data.");
      setSystemStatus("ok");
      updateLatestUpdatePill();
      return;
    }

    /*
     * If this is first load and there is no data,
     * show the error/loading message.
     */
    if (!incomingFlow.length && !state.hasRenderedOnce) {
      throw new Error("No Surface WIP rows returned from API.");
    }

    const newSignature = buildDataSignature(json);
    const isNewData = newSignature !== state.lastDataSignature;

    /*
     * Same data = do nothing.
     * This prevents the page from constantly repainting.
     */
    if (!forceRender && state.hasRenderedOnce && !isNewData) {
      state.lastFetch = new Date();
      setSystemStatus("ok");
      updateLatestUpdatePill();
      return;
    }

    /*
     * Only now update state.
     * Never overwrite state before validating payload.
     */
    state.summary = json.summary || state.summary || {};
    state.surfaceFlow = incomingFlow;

    state.surfaceTransfers = incomingTransfers.length
      ? incomingTransfers
      : state.surfaceTransfers || [];

    state.surfaceTransferDailyTotals = incomingDailyTotals.length
      ? incomingDailyTotals
      : state.surfaceTransferDailyTotals || [];

    state.intervalMeta = json.intervalHistoryMeta || state.intervalMeta || {};
    state.lastFetch = new Date();
    state.lastDataSignature = newSignature;

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

    /*
     * After first successful render, NEVER clear the dashboard on refresh error.
     */
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

  const latestSnapshotTime =
    state.intervalMeta?.currentTime ||
    state.summary?.ImportedAt ||
    state.lastFetch;

  const currentKey =
    state.intervalMeta?.currentKey ||
    state.summary?.SourceFile ||
    "";

  if (!latestSnapshotTime) {
    if (pill) pill.textContent = "Latest Update: —";
    if (summaryPill) summaryPill.textContent = "Latest Update: —";
    return;
  }

  const label = "Latest Update: " + formatDisplayDateTime(latestSnapshotTime);

  if (pill) pill.textContent = label;
  if (summaryPill) summaryPill.textContent = label;

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

/* =============================================
   RENDER ALL
============================================= */

function renderAll() {
  renderKPIs();
  renderFlowGrid();
  renderTransferTunnel();
  renderFlowDetail();
  renderLines();
  renderAlerts();
  renderAnalytics();
  renderReports();
}

/* =============================================
   KPI
============================================= */

function renderKPIs() {
  const s = state.summary;

  if (!s) return;

  const total = num(s.SurfaceTotalWIP);
  const largestTotal = num(s.LargestStepTotal);
  const largestStep = safeText(s.LargestStep);
  const pctLargest = total > 0 ? Math.round((largestTotal / total) * 100) : 0;

  setText("kpiTotal", total.toLocaleString());
  setText("kpiMain", num(s.SurfaceMainWIP).toLocaleString());
  setText("kpiActiveSteps", `${num(s.ActiveSteps)} active · ${num(s.EmptySteps)} empty steps`);
  setText("kpiIntake", `Intake (SF Scan): ${num(s.SurfaceIntakeWIP).toLocaleString()}`);
  setText("kpiEmpty", num(s.EmptySteps).toString());

  const bottleneck = document.getElementById("kpiBottleneck");
  if (bottleneck) {
    bottleneck.innerHTML = `${largestTotal.toLocaleString()} <span style="font-size:14px;color:var(--red)">— ${largestStep}</span>`;
  }

  setText("kpiBottleneckSub", `${pctLargest}% of total Surface WIP`);
}

/* =============================================
   FLOW GRID
============================================= */

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

/* =============================================
   TRANSFER TUNNEL
============================================= */

function renderTransferTunnel() {
  const grid = document.getElementById("transferGrid");
  const meta = document.getElementById("transferMeta");

  if (!grid || !meta) return;

  const rows = (state.surfaceTransfers || [])
    .slice()
    .sort((a, b) => num(a.FromOrder) - num(b.FromOrder));

  const intervalSnapshots = num(state.intervalMeta?.intervalSnapshots);
  const minutesBetween = num(state.intervalMeta?.minutesBetween);

  if (!rows.length) {
    meta.textContent =
      intervalSnapshots >= 2
        ? "No transfer rows available."
        : "Waiting for at least 2 interval snapshots…";

    grid.innerHTML = `
      <div class="transfer-loading">
        <div class="spinner"></div>
        <span>Need at least 2 interval snapshots before movement estimates can render.</span>
      </div>
    `;
    return;
  }

  meta.style.display = "none";

  const converters = [
    {
      title: "Surface Transfer Flow",
      subtitle: "Incoming from SF S&V → SF Unbox → Blocking Line",
      rows: rows.filter(row => {
        const fromOrder = num(row.FromOrder);
        return fromOrder === 1 || fromOrder === 2;
      })
    },
    {
      title: "Surface Transfer Flow",
      subtitle: "Blocking Line → IQ Star → Generating Line → Polishing Line",
      rows: rows.filter(row => {
        const fromOrder = num(row.FromOrder);
        return fromOrder === 4 || fromOrder === 5 || fromOrder === 7;
      })
    },
    {
      title: "Surface Transfer Flow",
      subtitle: "Polishing Line → Engraving Line → Detaper Line → Coater Line (54R)",
      rows: rows.filter(row => {
        const fromOrder = num(row.FromOrder);
        return fromOrder === 9 || fromOrder === 11 || fromOrder === 12;
      })
    },
    {
      title: "Surface Transfer Flow",
      subtitle: "Coater Line (54R) → Surface Inspection",
      rows: rows.filter(row => {
        const fromOrder = num(row.FromOrder);
        return fromOrder === 13;
      })
    }
  ];

  const visibleGroups = converters.filter(group => group.rows.length > 0);

  if (!visibleGroups.length) {
    grid.innerHTML = `
      <div class="transfer-loading" style="color:var(--red)">
        Transfer rows exist, but no FromOrder groups matched. Check FromOrder values from the API.
      </div>
    `;
    return;
  }

  grid.innerHTML = visibleGroups
    .map(group => renderConverterGroup(group, minutesBetween))
    .join("");
}

function renderConverterGroup(group, minutesBetween) {
  let scanned = 0;
  let reached = 0;
  let moving = 0;

  group.rows.forEach(row => {
    const transitionName = safeText(
      row.TransitionName || `${row.FromDisplayName} → ${row.ToDisplayName}`
    );

    const daily = findDailyTransferForTransition(transitionName);

    const laneScanned = daily ? num(daily.TotalLeftFromStep) : 0;
    const laneReached = daily ? num(daily.TotalEnteredToStep) : 0;

    /*
     * Correct movement logic:
     * Backend already calculates the movement estimate.
     * Do NOT use scanned - reached here.
     */
    const laneMoving = daily
      ? num(daily.TotalEstimatedInTransfer)
      : num(row.EstimatedInTransfer);

    scanned += laneScanned;
    reached += laneReached;
    moving += laneMoving;
  });

  const severity =
    moving >= 100 ? "critical" :
    moving >= 35 ? "watch" :
    "normal";

  return `
    <section class="converter-tab ${severity}">
      <div class="converter-tab-header">
        <div>
          <div class="converter-tab-title">${group.title}</div>
          <div class="converter-tab-subtitle">${group.subtitle}</div>
        </div>

        <div class="converter-tab-metrics">
          <div>
            <span>Total Scanned</span>
            <strong>${scanned}</strong>
          </div>

          <div>
            <span>Estimated Count</span>
            <strong>${moving}</strong>
          </div>

          <div>
            <span>Total Reached</span>
            <strong>${reached}</strong>
          </div>
        </div>
      </div>

      <div class="converter-tab-body">
        ${group.rows.map(row => renderConverterMiniLane(row, minutesBetween)).join("")}
      </div>
    </section>
  `;
}

function renderConverterMiniLane(row, minutesBetween) {
  const transitionName = safeText(row.TransitionName || `${row.FromDisplayName} → ${row.ToDisplayName}`);
  const fromDisplay = safeText(row.FromDisplayName || row.FromStep, "From");
  const toDisplay = safeText(row.ToDisplayName || row.ToStep, "To");

  const fromCurrent = num(row.FromCurrentWIP);
  const toCurrent = num(row.ToCurrentWIP);

  const latestLeft = num(row.LeftFromStep);
  const latestReached = num(row.EnteredToStep);

  const daily = findDailyTransferForTransition(transitionName);

  const scannedToday = daily ? num(daily.TotalLeftFromStep) : 0;
  const reachedToday = daily ? num(daily.TotalEnteredToStep) : 0;

  /*
   * Correct movement logic:
   * Show backend daily estimated movement.
   * If daily totals are missing, fall back to latest live estimate.
   */
  const movingBetween = daily
    ? num(daily.TotalEstimatedInTransfer)
    : num(row.EstimatedInTransfer);

  const severity =
    movingBetween >= 100 || toCurrent >= CONFIG.WIP_CRITICAL ? "critical" :
    movingBetween >= 35 || toCurrent >= CONFIG.WIP_HIGH ? "watch" :
    "normal";

  const packetsToShow = Math.max(
    3,
    Math.min(10, movingBetween || latestLeft || latestReached || 3)
  );

  const packets = Array.from(
    { length: packetsToShow },
    (_, i) => `<span class="transfer-packet" style="--delay:${i * 0.34}s"></span>`
  ).join("");

  return `
    <article class="converter-mini-lane ${severity}">
      <div class="converter-mini-top">
        <div>
          <div class="converter-mini-title">${transitionName}</div>
          <div class="converter-mini-subtitle">${fromDisplay} moving into ${toDisplay}</div>
        </div>
      </div>

      <div class="converter-mini-flow">
        <div class="converter-mini-box">
          <div class="transfer-zone-label">Scanned Today</div>
          <div class="transfer-zone-step">${fromDisplay}</div>
          <div class="transfer-zone-value">${scannedToday}</div>
          <div class="transfer-zone-note">Current WIP: ${fromCurrent}</div>
        </div>

        <div class="transfer-flow-arrow">→</div>

        <div class="converter-mini-conveyor">
          <div class="transfer-conveyor-head">
            <div class="transfer-conveyor-title">Estimated Transfer Count</div>
            <div class="transfer-conveyor-count">${movingBetween}</div>
          </div>

          <div class="transfer-track">
            ${packets}
          </div>
        </div>

        <div class="transfer-flow-arrow">→</div>

        <div class="converter-mini-box">
          <div class="transfer-zone-label">Reached Today</div>
          <div class="transfer-zone-step">${toDisplay}</div>
          <div class="transfer-zone-value">${reachedToday}</div>
          <div class="transfer-zone-note">Current WIP: ${toCurrent}</div>
        </div>
      </div>
    </article>
  `;
}

function renderConverterLane(row, minutesBetween) {
  const transitionName = safeText(row.TransitionName || `${row.FromDisplayName} → ${row.ToDisplayName}`);
  const fromDisplay = safeText(row.FromDisplayName || row.FromStep, "From");
  const toDisplay = safeText(row.ToDisplayName || row.ToStep, "To");

  const fromCurrent = num(row.FromCurrentWIP);
  const toCurrent = num(row.ToCurrentWIP);

  const daily = findDailyTransferForTransition(transitionName);

  const totalScannedFromArea = daily ? num(daily.TotalLeftFromStep) : 0;
  const totalReachedNextStation = daily ? num(daily.TotalEnteredToStep) : 0;

  /*
   * Correct movement logic:
   * Backend daily estimate is the conveyor number.
   */
  const estimatedMovingBetween = daily
    ? num(daily.TotalEstimatedInTransfer)
    : num(row.EstimatedInTransfer);

  const latestLeft = num(row.LeftFromStep);
  const latestReached = num(row.EnteredToStep);

  let severity = "normal";

  if (estimatedMovingBetween >= 100 || toCurrent >= CONFIG.WIP_CRITICAL) {
    severity = "critical";
  } else if (estimatedMovingBetween >= 35 || toCurrent >= CONFIG.WIP_HIGH) {
    severity = "watch";
  }

  const packetsToShow = Math.max(
    3,
    Math.min(10, estimatedMovingBetween || latestLeft || latestReached || 3)
  );

  const packets = Array.from(
    { length: packetsToShow },
    (_, i) => `<span class="transfer-packet" style="--delay:${i * 0.34}s"></span>`
  ).join("");

  return `
    <article class="transfer-card converter-lane ${severity}">
      <div class="transfer-top">
        <div class="transfer-route">
          <div class="transfer-route-name">${transitionName}</div>
          <div class="transfer-route-steps">
            ${fromDisplay} scan flow moving into ${toDisplay}
          </div>
        </div>
      </div>

      <div class="transfer-single-lane">

        <div class="transfer-main-zone">
          <div class="transfer-zone-label">Total Scanned Today</div>
          <div class="transfer-zone-step">${fromDisplay}</div>
          <div class="transfer-zone-value">${totalScannedFromArea}</div>
          <div class="transfer-zone-note">Current ${fromDisplay} WIP: ${fromCurrent}</div>
        </div>

        <div class="transfer-flow-arrow">→</div>

        <div class="transfer-main-conveyor is-moving">
          <div class="transfer-conveyor-head">
            <div class="transfer-conveyor-title">Estimated Moving Count</div>
            <div class="transfer-conveyor-count">${estimatedMovingBetween}</div>
          </div>

          <div class="transfer-track">
            ${packets}
          </div>

          <div class="transfer-zone-note transfer-center-note">
            Backend estimated movement from interval snapshots
          </div>
        </div>

        <div class="transfer-flow-arrow">→</div>

        <div class="transfer-main-zone">
          <div class="transfer-zone-label">Total Reached Today</div>
          <div class="transfer-zone-step">${toDisplay}</div>
          <div class="transfer-zone-value">${totalReachedNextStation}</div>
          <div class="transfer-zone-note">Current ${toDisplay} WIP: ${toCurrent}</div>
        </div>

      </div>

      <div class="transfer-bottom">
        <span><strong>From:</strong> ${safeText(row.FromStep)}</span>
        <span><strong>To:</strong> ${safeText(row.ToStep)}</span>
        <span><strong>Logic:</strong> Moving Between = Backend Estimated Transfer</span>
      </div>
    </article>
  `;
}

function converterIcon(type) {
  const map = {
    scan: ICONS.scan,
    gear: ICONS.gear,
    engrave: ICONS.engrave,
    inspect: ICONS.shield
  };

  return map[type] || ICONS.gearDefault;
}

/* =============================================
   FLOW DETAIL TAB
============================================= */

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

/* =============================================
   LINES TAB
============================================= */

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

/* =============================================
   ALERTS TAB
============================================= */

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

/* =============================================
   ANALYTICS TAB
============================================= */

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

  grid.innerHTML = `
    <article class="analytics-card">
      <div class="kpi-label">Total Surface WIP</div>
      <div class="ac-value">${num(s.SurfaceTotalWIP).toLocaleString()}</div>
      <div class="ac-sub">Report date: ${safeText(s.ReportDate)}</div>
    </article>

    <article class="analytics-card">
      <div class="kpi-label">Main Surface WIP</div>
      <div class="ac-value">${num(s.SurfaceMainWIP).toLocaleString()}</div>
      <div class="ac-sub">Excludes SF Scan intake.</div>
    </article>

    <article class="analytics-card">
      <div class="kpi-label">Largest Step</div>
      <div class="ac-value">${safeText(s.LargestStep)}</div>
      <div class="ac-sub">${num(s.LargestStepTotal).toLocaleString()} jobs · ${Math.round((num(s.LargestStepTotal) / total) * 100)}%</div>
    </article>

    <article class="analytics-card">
      <div class="kpi-label">Transfer Snapshots</div>
      <div class="ac-value">${num(state.intervalMeta?.intervalSnapshots)}</div>
      <div class="ac-sub">Gap: ${num(state.intervalMeta?.minutesBetween)} minutes</div>
    </article>

    <article class="analytics-card" style="grid-column:1/-1">
      <div class="kpi-label">WIP by Flow Group</div>

      <div class="group-bars" style="margin-top:14px">
        ${entries.map(([name, value]) => `
          <div class="group-bar-row">
            <div class="gb-name">${name}</div>
            <div class="gb-track">
              <div class="gb-fill" style="width:${Math.round((value / max) * 100)}%"></div>
            </div>
            <div class="gb-val">${value.toLocaleString()}</div>
          </div>
        `).join("")}
      </div>
    </article>
  `;
}

/* =============================================
   SUMMARY TAB
============================================= */

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

  const topTransfers = (state.surfaceTransferDailyTotals || [])
    .slice()
    .sort((a, b) => num(b.TotalEstimatedInTransfer) - num(a.TotalEstimatedInTransfer))
    .slice(0, 5);

  const totalEstimatedTransfer = (state.surfaceTransferDailyTotals || []).reduce((sum, row) => {
    return sum + num(row.TotalEstimatedInTransfer);
  }, 0);

  const highTransferCount = (state.surfaceTransferDailyTotals || []).filter(row => {
    return num(row.TotalEstimatedInTransfer) >= 5;
  }).length;

  wrap.innerHTML = `
    <section class="report-summary-shell">

      <div class="report-hero-grid">

        <article class="report-hero-card">
          <div class="report-hero-content">
            <div class="report-eyebrow">Surface WIP Summary</div>
            <div class="report-hero-title">${total.toLocaleString()}</div>

            <p class="report-hero-sub">
              Current total Surface WIP from the latest Facility WIP report.
              This view focuses on decision-ready metrics instead of the raw table.
            </p>

            <div class="report-hero-metrics">
              <div class="report-mini-metric">
                <div class="report-mini-label">Main WIP</div>
                <div class="report-mini-value">${main.toLocaleString()}</div>
                <div class="report-mini-note">Excludes SF Scan intake</div>
              </div>

              <div class="report-mini-metric">
                <div class="report-mini-label">Intake WIP</div>
                <div class="report-mini-value">${intake.toLocaleString()}</div>
                <div class="report-mini-note">SF Scan & Verify</div>
              </div>

              <div class="report-mini-metric">
                <div class="report-mini-label">Daily Transfer Estimate</div>
                <div class="report-mini-value">${totalEstimatedTransfer.toLocaleString()}</div>
                <div class="report-mini-note">All snapshot pairs today</div>
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

            <div class="report-status-row">
              <div>
                <div class="report-status-name">High Daily Transfer Lanes</div>
                <div class="report-status-sub">Daily estimated lanes at 5+ jobs</div>
              </div>
              <div class="report-status-count">${highTransferCount}</div>
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

      <div class="report-section-grid">

        <article class="report-section-card">
          <div class="report-section-head">
            <div>
              <div class="report-section-title">Daily Transfer Summary</div>
              <div class="report-section-sub">
                Cumulative estimate across all snapshot pairs today. Not exact job tracking.
              </div>
            </div>

            <div class="report-chip">
              Pairs: ${num(state.surfaceTransferDailyTotals?.[0]?.SnapshotPairs)}
            </div>
          </div>

          <div class="report-transfer-list">
            ${
              topTransfers.length
                ? topTransfers.map(row => {
                    const value = num(row.TotalEstimatedInTransfer);
                    const cls = value >= 15 ? "critical" : value >= 5 ? "warn" : "";

                    return `
                      <div class="report-transfer-row">
                        <div>
                          <div class="report-transfer-name">${safeText(row.TransitionName)}</div>
                          <div class="report-transfer-sub">
                            Daily Left: ${num(row.TotalLeftFromStep)} · Daily Entered: ${num(row.TotalEnteredToStep)} · Max Gap: ${num(row.MaxEstimatedInTransfer)}
                          </div>
                        </div>

                        <div class="report-transfer-value ${cls}">${value}</div>
                      </div>
                    `;
                  }).join("")
                : `<div class="flow-loading">No daily transfer metrics available yet.</div>`
            }
          </div>
        </article>

      </div>

    </section>
  `;
}

/* =============================================
   EXPORT CSV
============================================= */

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

/* =============================================
   CLOCK
============================================= */

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

/* =============================================
   EVENTS
============================================= */

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

/* =============================================
   BOOT
============================================= */

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