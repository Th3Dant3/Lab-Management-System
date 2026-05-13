/*********************************************************
 * ARDash.js
 * AR Dashboard Frontend
 *********************************************************/

/* ─────────────────────────────────────────────────────
   LOADING SCREEN
───────────────────────────────────────────────────── */
(function initLoader() {
  const BOOT_LINES = [
    { tag: "tag-init", prefix: "INIT", text: "AR Dashboard v2.2" },
    { tag: "tag-sys", prefix: "SYS", text: "Connecting to Google Apps Script API..." },
    { tag: "tag-ok", prefix: "OK", text: "AR station map loaded" },
    { tag: "tag-sys", prefix: "SYS", text: "Fetching WIP, capacity, and output..." },
    { tag: "tag-ok", prefix: "OK", text: "Rendering dashboard components" }
  ];

  const log = document.getElementById("arlLog");
  const fill = document.getElementById("arlFill");
  const pct = document.getElementById("arlPct");
  let lineIdx = 0;

  function addLine(entry) {
    if (!log) return;

    const el = document.createElement("div");
    el.className = "arl-log-line";
    el.innerHTML = `<span class="${entry.tag}">${entry.prefix}</span><span>${entry.text}</span>`;
    log.appendChild(el);

    if (log.children.length > 4) {
      log.removeChild(log.firstChild);
    }
  }

  function tick() {
    if (lineIdx >= BOOT_LINES.length) return;

    addLine(BOOT_LINES[lineIdx++]);

    const progress = Math.min(92, Math.round((lineIdx / BOOT_LINES.length) * 92));

    if (fill) fill.style.width = progress + "%";
    if (pct) pct.textContent = progress + "%";
  }

  const canvas = document.getElementById("arl-canvas");

  if (canvas) {
    const ctx = canvas.getContext("2d");
    const colors = [
      "rgba(255,153,0,",
      "rgba(155,92,255,",
      "rgba(86,227,109,",
      "rgba(255,191,63,"
    ];
    const pts = [];

    function resizeCanvas() {
      canvas.width = window.innerWidth;
      canvas.height = window.innerHeight;
    }

    resizeCanvas();
    window.addEventListener("resize", resizeCanvas);

    for (let i = 0; i < 55; i++) {
      pts.push({
        x: Math.random() * window.innerWidth,
        y: Math.random() * window.innerHeight,
        r: Math.random() * 1.6 + 0.4,
        vx: (Math.random() - 0.5) * 0.24,
        vy: (Math.random() - 0.5) * 0.24,
        col: colors[Math.floor(Math.random() * colors.length)],
        a: Math.random() * 0.28 + 0.06
      });
    }

    let raf;

    function drawFrame() {
      if (!document.getElementById("ar-loader")) return;

      ctx.clearRect(0, 0, canvas.width, canvas.height);

      pts.forEach(p => {
        p.x += p.vx;
        p.y += p.vy;

        if (p.x < 0) p.x = canvas.width;
        if (p.x > canvas.width) p.x = 0;
        if (p.y < 0) p.y = canvas.height;
        if (p.y > canvas.height) p.y = 0;

        ctx.beginPath();
        ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
        ctx.fillStyle = p.col + p.a + ")";
        ctx.fill();
      });

      raf = requestAnimationFrame(drawFrame);
    }

    drawFrame();

    window._arlCancelRaf = () => cancelAnimationFrame(raf);
  }

  setTimeout(tick, 320);
  const logInterval = setInterval(tick, 430);

  window._arlDismiss = function(isLive) {
    clearInterval(logInterval);

    addLine(
      isLive
        ? { tag: "tag-ok", prefix: "OK", text: "Dashboard ready" }
        : { tag: "tag-warn", prefix: "WARN", text: "API unavailable" }
    );

    if (fill) fill.style.width = "100%";
    if (pct) pct.textContent = "100%";

    setTimeout(() => {
      const loader = document.getElementById("ar-loader");
      if (!loader) return;

      loader.classList.add("arl-exit");

      if (window._arlCancelRaf) {
        window._arlCancelRaf();
      }

      loader.addEventListener("transitionend", () => loader.remove(), { once: true });
    }, 400);
  };
})();

/* ─────────────────────────────────────────────────────
   CONFIG
───────────────────────────────────────────────────── */
const API_URL = "https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec";

const REFRESH_MS = 5 * 60 * 1000;
const TOTAL_VISIBLE_CHAMBERS = 6;

const AR_CAPACITY_RULES = {
  BASKET_LENS: 32,
  OVEN_BASKETS: 27,
  OVEN_LENS: 864,
  CHAMBER_LENS: 192
};

const STATION_COLORS = {
  "Surface Inspection WIP": "#4da3ff",
  "Surface Inspection": "#4da3ff",
  "AR-IN": "#ff4f7b",
  "Basket": "#ff9900",
  "Oven": "#ffd24a",
  "Sectoring": "#66d1cf",
  "DeRing": "#9b5cff",
  "AR-OUT": "#d9d9d9"
};

let LAST_FLOW_PAYLOAD = null;
let LAST_CAPACITY_PAYLOAD = null;
let LAST_AR_ROWS = [];
let LAST_VALUES = null;

let chartInstances = {
  outputTrend: null,
  wipTrend: null
};

/* ─────────────────────────────────────────────────────
   HELPERS
───────────────────────────────────────────────────── */
const $ = (id) => document.getElementById(id);

function toNumber(value) {
  const n = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

function fmt(value) {
  return toNumber(value).toLocaleString("en-US");
}

function pct(value, total) {
  const v = toNumber(value);
  const t = toNumber(total);

  if (!t) return 0;

  return Math.max(0, Math.min(100, Math.round((v / t) * 100)));
}

function ceilUnit(value, capacity) {
  const v = toNumber(value);
  const c = toNumber(capacity);

  if (v <= 0 || c <= 0) return 0;

  return Math.ceil(v / c);
}

function setText(id, value) {
  const el = $(id);
  if (el) el.textContent = value;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function fmtDateTime(value) {
  if (!value) return "--";

  const d = new Date(value);

  if (Number.isNaN(d.getTime())) return String(value);

  return d.toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit"
  });
}

function showToast(message) {
  const el = $("toast");
  if (!el) return;

  el.textContent = message;
  el.classList.add("show");

  window.clearTimeout(showToast._timer);

  showToast._timer = window.setTimeout(() => {
    el.classList.remove("show");
  }, 2800);
}

function makeDiv(className) {
  const div = document.createElement("div");
  div.className = className;
  return div;
}

function getStationColor(name, fallback = "#ff9900") {
  return STATION_COLORS[name] || fallback;
}

/* ─────────────────────────────────────────────────────
   API
───────────────────────────────────────────────────── */
async function fetchJson(url) {
  const res = await fetch(url, { cache: "no-store" });

  if (!res.ok) {
    throw new Error(`API failed: ${res.status}`);
  }

  const payload = await res.json();

  if (payload.status && String(payload.status).toLowerCase() !== "success") {
    throw new Error(payload.message || "API returned an error.");
  }

  return payload;
}

async function fetchProductionFlowAR() {
  return fetchJson(`${API_URL}?action=productionFlow&area=AR&debug=true`);
}

async function fetchARCapacity() {
  return fetchJson(`${API_URL}?action=arCapacity`);
}

async function loadARData() {
  try {
    const [flowPayload, capacityPayload] = await Promise.all([
      fetchProductionFlowAR(),
      fetchARCapacity().catch(() => null)
    ]);

    LAST_FLOW_PAYLOAD = flowPayload;
    LAST_CAPACITY_PAYLOAD = capacityPayload || null;

    console.log("REAL AR FLOW PAYLOAD:", flowPayload);
    console.table(flowPayload.productionFlow || []);

    setLiveState(true);
    renderDashboard(LAST_FLOW_PAYLOAD, LAST_CAPACITY_PAYLOAD);

    if (window._arlDismiss) {
      window._arlDismiss(true);
    }

  } catch (err) {
    console.error("AR API ERROR:", err);

    setLiveState(false);
    showToast("AR API failed. Real output data not loaded.");

    const emptyPayload = {
      status: "error",
      generatedAt: new Date().toISOString(),
      summary: {
        LastUpdated: new Date().toISOString(),
        TotalWIP: 0,
        ActiveWIPStations: 0,
        LargestWIPStation: "API Error"
      },
      productionFlow: []
    };

    LAST_FLOW_PAYLOAD = emptyPayload;
    LAST_CAPACITY_PAYLOAD = null;

    renderDashboard(emptyPayload, null);

    if (window._arlDismiss) {
      window._arlDismiss(false);
    }
  }
}

function setLiveState(isLive) {
  const live = $("liveState");

  if (live) {
    live.textContent = isLive ? "Live" : "API Error";
    live.classList.toggle("offline", !isLive);
  }

  setText("liveMiniStatus", isLive ? "Connected" : "API Failed");
}

/* ─────────────────────────────────────────────────────
   ROW / DATA HELPERS
───────────────────────────────────────────────────── */
function getARRows(flowPayload) {
  const rows = Array.isArray(flowPayload?.productionFlow)
    ? flowPayload.productionFlow
    : [];

  return rows
    .filter(row => String(row.Area || "").trim().toUpperCase() === "AR")
    .sort((a, b) => toNumber(a.FlowOrder) - toNumber(b.FlowOrder));
}

function getRowByStep(rows, stepName) {
  const target = String(stepName || "").trim().toUpperCase();

  return rows.find(row =>
    String(row.FlowStep || "").trim().toUpperCase() === target ||
    String(row.DisplayName || "").trim().toUpperCase() === target
  ) || {};
}

function getSurfaceInputRow(rows) {
  return rows.find(row =>
    String(row.BridgeRole || "").trim().toUpperCase() === "AR_INPUT"
  ) || getRowByStep(rows, "Surface Inspection") || {};
}

function getAROutRow(rows) {
  return rows.find(row =>
    String(row.BridgeRole || "").trim().toUpperCase() === "AR_OUTPUT"
  ) || getRowByStep(rows, "AR-OUT") || {};
}

function getWipRows(rows) {
  return rows.filter(row =>
    String(row.MetricMode || "").trim().toUpperCase() !== "OUTPUT_ONLY"
  );
}

function getOutputRows(rows) {
  return rows.filter(row =>
    String(row.MetricMode || "").trim().toUpperCase() !== "WIP_ONLY"
  );
}

function getCleanStationRows(rows) {
  return rows.filter(row => {
    const name = String(row.DisplayName || row.FlowStep || "").trim().toUpperCase();
    return name && name !== "TOTAL" && name !== "TOTAL AR ACTIVITY TODAY";
  });
}

function getTopRowName(rows, fieldName) {
  if (!rows || !rows.length) return "—";

  const top = rows.reduce((best, row) => {
    return toNumber(row[fieldName]) > toNumber(best[fieldName]) ? row : best;
  }, rows[0]);

  return top.DisplayName || top.FlowStep || "—";
}

function getLargestWipStation(rows) {
  const wipRows = getCleanStationRows(getWipRows(rows)).filter(row => {
    const name = String(row.DisplayName || row.FlowStep || "").trim().toUpperCase();
    return name && name !== "AR-OUT";
  });

  if (!wipRows.length) {
    return { name: "--", value: 0 };
  }

  const top = wipRows.reduce((best, row) => {
    return toNumber(row.CurrentWIP) > toNumber(best.CurrentWIP) ? row : best;
  }, wipRows[0]);

  return {
    name: top.DisplayName || top.FlowStep || "--",
    value: toNumber(top.CurrentWIP)
  };
}

function buildUnitsFromWip(total, capacity, unitType) {
  const output = [];
  const count = ceilUnit(total, capacity);

  for (let i = 0; i < count; i++) {
    const used = Math.max(0, Math.min(capacity, toNumber(total) - i * capacity));

    output.push({
      UnitType: unitType,
      UnitNumber: i + 1,
      Used: used,
      Capacity: capacity,
      PercentUsed: pct(used, capacity),
      Status: used >= capacity ? "FULL" : used > 0 ? "PARTIAL" : "EMPTY"
    });
  }

  return output;
}

/* ─────────────────────────────────────────────────────
   MAIN RENDER
───────────────────────────────────────────────────── */
function renderDashboard(flowPayload, capacityPayload) {
  const flowSummary = flowPayload?.summary || {};
  const arRows = getARRows(flowPayload);

  LAST_AR_ROWS = arRows;

  const surfaceInputRow = getSurfaceInputRow(arRows);
  const arInRow = getRowByStep(arRows, "AR-IN");
  const basketRow = getRowByStep(arRows, "Basket");
  const ovenRow = getRowByStep(arRows, "Oven");
  const sectoringRow = getRowByStep(arRows, "Sectoring");
  const deringRow = getRowByStep(arRows, "DeRing");
  const arOutRow = getAROutRow(arRows);

  const surfaceInputWip = toNumber(surfaceInputRow.CurrentWIP);
  const arInWip = toNumber(arInRow.CurrentWIP);
  const basketWip = toNumber(basketRow.CurrentWIP);
  const ovenWip = toNumber(ovenRow.CurrentWIP);
  const sectoringWip = toNumber(sectoringRow.CurrentWIP);
  const deringWip = toNumber(deringRow.CurrentWIP);
  const arOutActivity = toNumber(arOutRow.ActivityToday);

  const arInActivity = toNumber(arInRow.ActivityToday);
  const basketActivity = toNumber(basketRow.ActivityToday);
  const ovenActivity = toNumber(ovenRow.ActivityToday);
  const sectoringActivity = toNumber(sectoringRow.ActivityToday);
  const deringActivity = toNumber(deringRow.ActivityToday);

  const totalArWip = toNumber(flowSummary.TotalWIP) ||
    surfaceInputWip + arInWip + basketWip + ovenWip + sectoringWip + deringWip;

  const totalActivityToday =
    arInActivity + basketActivity + ovenActivity + sectoringActivity + deringActivity + arOutActivity;

  const activeWipStations = toNumber(flowSummary.ActiveWIPStations) ||
    [surfaceInputWip, arInWip, basketWip, ovenWip, sectoringWip, deringWip].filter(n => n > 0).length;

  const basketUnits = buildUnitsFromWip(basketWip, AR_CAPACITY_RULES.BASKET_LENS, "Basket");

  const ovenBasketLoad = ceilUnit(ovenWip, AR_CAPACITY_RULES.BASKET_LENS);
  const ovenUnits = buildUnitsFromWip(ovenWip, AR_CAPACITY_RULES.OVEN_LENS, "Oven");
  const activeOvens = ovenWip > 0 ? Math.max(1, ovenUnits.length) : 0;
  const ovenTotalLensCapacity = activeOvens > 0
    ? activeOvens * AR_CAPACITY_RULES.OVEN_LENS
    : AR_CAPACITY_RULES.OVEN_LENS;

  const chamberUnits = buildUnitsFromWip(sectoringWip, AR_CAPACITY_RULES.CHAMBER_LENS, "Chamber");
  const activeChambers = chamberUnits.length;
  const chamberTotalCapacity = Math.max(
    activeChambers * AR_CAPACITY_RULES.CHAMBER_LENS,
    AR_CAPACITY_RULES.CHAMBER_LENS
  );

  const lastBasket = basketUnits.length
    ? basketUnits[basketUnits.length - 1]
    : { Used: 0, Capacity: AR_CAPACITY_RULES.BASKET_LENS };

  const lastChamber = chamberUnits.length
    ? chamberUnits[chamberUnits.length - 1]
    : { Used: 0, Capacity: AR_CAPACITY_RULES.CHAMBER_LENS };

  const fullBasketCount = basketUnits.filter(u =>
    toNumber(u.Used) >= toNumber(u.Capacity)
  ).length;

  const partialBasketCount = basketUnits.filter(u =>
    toNumber(u.Used) > 0 && toNumber(u.Used) < toNumber(u.Capacity)
  ).length;

  const lastOvenBasketText = `${ovenBasketLoad} / ${AR_CAPACITY_RULES.OVEN_BASKETS}`;

  const ovenLensRemaining = Math.max(0, ovenTotalLensCapacity - ovenWip);
  const chamberLensRemaining = Math.max(0, chamberTotalCapacity - sectoringWip);

  const largestWip = getLargestWipStation(arRows);

  LAST_VALUES = {
    surfaceInputWip,
    arInWip,
    basketWip,
    ovenWip,
    sectoringWip,
    deringWip,
    arOutActivity,
    arInActivity,
    basketActivity,
    ovenActivity,
    sectoringActivity,
    deringActivity,
    totalArWip,
    totalActivityToday,
    activeWipStations,
    basketUnits,
    fullBasketCount,
    partialBasketCount,
    ovenBasketLoad,
    ovenUnits,
    activeOvens,
    ovenTotalLensCapacity,
    ovenLensRemaining,
    chamberUnits,
    activeChambers,
    chamberTotalCapacity,
    chamberLensRemaining,
    lastBasket,
    lastChamber,
    largestWip
  };

  setText("lastUpdated", fmtDateTime(flowSummary.LastUpdated || capacityPayload?.summary?.LastUpdated || flowPayload?.generatedAt));
  setText("largestWIP", largestWip.name);
  setText("largestPressure", largestWip.name);

  animateValue("totalArWip", totalArWip);
  animateValue("surfaceInputWip", surfaceInputWip);
  animateValue("activeBasketsKpi", basketUnits.length);
  animateValue("activeOvensKpi", activeOvens);
  animateValue("activeChambersKpi", activeChambers);
  animateValue("arOutActivity", arOutActivity);

  setText("activeStations", `${fmt(activeWipStations)} active WIP stations`);
  setText("basketLensText", `${fmt(basketWip)} lenses / ${AR_CAPACITY_RULES.BASKET_LENS} each`);
  setText("ovenBasketText", `${fmt(ovenBasketLoad)} baskets / ${AR_CAPACITY_RULES.OVEN_BASKETS} each (${fmt(AR_CAPACITY_RULES.OVEN_LENS)} lenses)`);
  setText("chamberLensText", `${fmt(sectoringWip)} lenses / ${AR_CAPACITY_RULES.CHAMBER_LENS} each`);

  setText("stationSurfaceInput", fmt(surfaceInputWip));
  setText("stationArIn", fmt(arInWip));
  setText("stationBaskets", fmt(basketWip));
  setText("stationOven", fmt(ovenWip));
  setText("stationChamberLoad", fmt(sectoringWip));
  setText("stationDeRing", fmt(deringWip));
  setText("stationArOut", fmt(arOutActivity));

  setText(
    "lastBasketLoad",
    `${fullBasketCount} full · ${partialBasketCount} partial | Last basket: ${fmt(lastBasket.Used)} / ${fmt(lastBasket.Capacity)}`
  );

  setText("lastOvenLoad", `Basket equivalent: ${lastOvenBasketText} | Lens cap: ${fmt(ovenTotalLensCapacity)}`);
  setText("ovenBasketBadge", lastOvenBasketText);

  setText("queueSurface", fmt(surfaceInputWip));
  setText("queueArIn", fmt(arInWip));
  setText("queueBasket", fmt(basketWip));
  setText("queueOven", fmt(ovenWip));
  setText("queueSectoring", fmt(sectoringWip));
  setText("queueDeRing", fmt(deringWip));

  setText("capBasketTotalWip", fmt(basketWip));
  setText("capBasketFull", fmt(fullBasketCount));
  setText("capBasketPartial", fmt(partialBasketCount));
  setText("capBasketCurrent", `${fmt(lastBasket.Used)} / ${fmt(lastBasket.Capacity)}`);

  setText("capOvenTotalWip", fmt(ovenWip));
  setText("capOvenBasketLoad", lastOvenBasketText);
  setText("capOvenLensCap", fmt(ovenTotalLensCapacity));
  setText("capOvenRemaining", fmt(ovenLensRemaining));

  setText("capChamberTotalWip", fmt(sectoringWip));
  setText("capChamberActive", fmt(activeChambers));
  setText("capChamberCurrent", `${fmt(lastChamber.Used)} / ${fmt(lastChamber.Capacity)}`);
  setText("capChamberRemaining", fmt(chamberLensRemaining));

  renderSplitDetailTabs(arRows, flowSummary);
  renderBaskets(basketUnits, AR_CAPACITY_RULES.BASKET_LENS);
  renderOvens(ovenWip, ovenBasketLoad, ovenTotalLensCapacity);
  renderChambers(chamberUnits, AR_CAPACITY_RULES.CHAMBER_LENS, sectoringWip, chamberTotalCapacity);
  renderUtilList(chamberUnits, AR_CAPACITY_RULES.CHAMBER_LENS);
  renderAlerts(LAST_VALUES);
  renderDailyBrief(LAST_VALUES);
  renderSummaryBreakdown(LAST_VALUES);
  renderCharts(flowPayload, arRows, LAST_VALUES);
  placeArrows();
}

/* ─────────────────────────────────────────────────────
   SPLIT DETAIL
───────────────────────────────────────────────────── */
function renderSplitDetailTabs(arRows, flowSummary) {
  const panel = $("arSplitDetailPanel");
  const body = $("arSplitBody");
  const summary = $("arSplitSummary");
  const subtitle = $("arSplitSubtitle");

  if (!panel || !body || !summary) return;

  const activeTab = panel.dataset.activeTab || "wip";

  if (activeTab === "output") {
    if (subtitle) {
      subtitle.textContent = "Completed output by each AR station.";
    }

    renderOutputTotals(body, summary, getCleanStationRows(getOutputRows(arRows)));
  } else {
    if (subtitle) {
      subtitle.textContent = "Current WIP sitting at each AR station.";
    }

    renderWipTotals(body, summary, getCleanStationRows(getWipRows(arRows)), flowSummary);
  }
}

function renderWipTotals(body, summaryBox, rows, flowSummary) {
  const totalWip = toNumber(flowSummary.TotalWIP) ||
    rows.reduce((sum, row) => sum + toNumber(row.CurrentWIP), 0);

  const bridgeInput = toNumber(flowSummary.BridgeInputWIP) ||
    toNumber(getSurfaceInputRow(rows).CurrentWIP);

  const largest = getLargestWipStation(LAST_AR_ROWS);
  const maxValue = Math.max(...rows.map(row => toNumber(row.CurrentWIP)), 1);

  summaryBox.innerHTML = `
    <div class="ar-summary-card">
      <span>Total AR WIP</span>
      <strong>${fmt(totalWip)}</strong>
    </div>

    <div class="ar-summary-card">
      <span>Surface Input WIP</span>
      <strong>${fmt(bridgeInput)}</strong>
    </div>

    <div class="ar-summary-card">
      <span>Largest WIP Station</span>
      <strong>${escapeHtml(largest.name)}</strong>
    </div>
  `;

  body.innerHTML = rows.map(row => {
    const value = toNumber(row.CurrentWIP);
    const width = Math.max(4, Math.round((value / maxValue) * 100));
    const isInput = String(row.BridgeRole || "").toUpperCase() === "AR_INPUT";

    return `
      <div class="ar-detail-row ${isInput ? "input-row" : ""}">
        <div class="ar-detail-main">
          <div class="ar-detail-title">
            <span>${escapeHtml(row.DisplayName || row.FlowStep || "")}</span>
            <em>${isInput ? "AR Input" : "Current WIP"}</em>
          </div>

          <div class="ar-detail-track">
            <div class="ar-detail-fill" style="width:${width}%"></div>
          </div>
        </div>

        <div class="ar-detail-value">${fmt(value)}</div>
      </div>
    `;
  }).join("");
}

function renderOutputTotals(body, summaryBox, rows) {
  const cleanRows = getCleanStationRows(rows);
  const maxValue = Math.max(...cleanRows.map(row => toNumber(row.ActivityToday)), 1);

  const topRows = [...cleanRows]
    .sort((a, b) => toNumber(b.ActivityToday) - toNumber(a.ActivityToday))
    .slice(0, 3);

  summaryBox.innerHTML = topRows.map(row => `
    <div class="ar-summary-card output">
      <span>${escapeHtml(row.DisplayName || row.FlowStep || "")}</span>
      <strong>${fmt(row.ActivityToday)}</strong>
    </div>
  `).join("");

  body.innerHTML = cleanRows.map(row => {
    const value = toNumber(row.ActivityToday);
    const width = Math.max(4, Math.round((value / maxValue) * 100));
    const isOutput = String(row.MetricMode || "").toUpperCase() === "OUTPUT_ONLY";
    const name = row.DisplayName || row.FlowStep || "";

    return `
      <div class="ar-detail-row ${isOutput ? "output-row" : ""}">
        <div class="ar-detail-main">
          <div class="ar-detail-title">
            <span>${escapeHtml(name)}</span>
            <em>${isOutput ? "Output to Finish" : "Station Activity"}</em>
          </div>

          <div class="ar-detail-track output-track">
            <div class="ar-detail-fill" style="width:${width}%; background:${getStationColor(name)};"></div>
          </div>
        </div>

        <div class="ar-detail-value output-value">${fmt(value)}</div>
      </div>
    `;
  }).join("");
}

/* ─────────────────────────────────────────────────────
   VISUAL RENDERERS
───────────────────────────────────────────────────── */
function renderBaskets(units, basketCap) {
  const fullGrid = $("basketFullGrid");
  const partGrid = $("basketPartialGrid");

  if (!fullGrid || !partGrid) return;

  const full = units.filter(u => toNumber(u.Used) >= toNumber(u.Capacity || basketCap));
  const partial = units.filter(u => toNumber(u.Used) > 0 && toNumber(u.Used) < toNumber(u.Capacity || basketCap));

  setText("fullBasketCount", full.length);
  setText("partialBasketCount", partial.length);

  fullGrid.innerHTML = "";
  partGrid.innerHTML = "";

  if (!full.length) {
    fullGrid.innerHTML = '<div class="basket-empty-msg">none</div>';
  } else {
    full.forEach((basket, index) => {
      const cell = document.createElement("div");
      cell.className = "basket-unit full-basket-unit";
      cell.innerHTML = `
        <strong>B${index + 1}</strong>
        <span>32 / 32</span>
      `;
      fullGrid.appendChild(cell);
    });
  }

  if (!partial.length) {
    partGrid.innerHTML = '<div class="basket-empty-msg">none</div>';
  } else {
    partial.forEach((basket, index) => {
      const used = toNumber(basket.Used);
      const cap = toNumber(basket.Capacity || basketCap);
      const fillPct = pct(used, cap);

      const cell = document.createElement("div");
      cell.className = "basket-unit partial-basket-unit";
      cell.innerHTML = `
        <div class="basket-fill" style="height:${fillPct}%"></div>
        <strong>P${index + 1}</strong>
        <span>${used} / ${cap}</span>
      `;
      partGrid.appendChild(cell);
    });
  }
}

function renderOvens(ovenWip, ovenBasketLoad, ovenTotalLensCapacity) {
  const fillPct = pct(ovenWip, ovenTotalLensCapacity);
  const ovenHeat = document.querySelector(".oven-door .heat");

  if (ovenHeat) {
    ovenHeat.style.opacity = String(Math.max(0.35, fillPct / 100));
  }
}

function renderChambers(units, chamberCap, sectoringWip, chamberTotalCapacity) {
  const grid = $("chamberGrid");
  if (!grid) return;

  grid.innerHTML = "";

  const visibleCount = Math.max(TOTAL_VISIBLE_CHAMBERS, units.length || 0);
  let activeCount = 0;
  let totalLenses = 0;

  for (let i = 0; i < visibleCount; i++) {
    const unit = units[i];
    const used = toNumber(unit?.Used);
    const cap = toNumber(unit?.Capacity) || chamberCap;

    const state = used <= 0
      ? "empty"
      : used >= cap
        ? "full active"
        : "partial";

    if (used > 0) activeCount++;
    totalLenses += used;

    const div = makeDiv(`chamber ${state}`);

    div.innerHTML = `
      <div class="chamber-num">C${i + 1}</div>
      <div class="chamber-count">${used > 0 ? `${used} / ${cap}` : "— / " + cap}</div>
      <div class="chamber-label">${used <= 0 ? "Empty" : used >= cap ? "Full" : "Partial"}</div>
    `;

    grid.appendChild(div);
  }

  const totalCap = Math.max(toNumber(chamberTotalCapacity), chamberCap);
  const utilPct = pct(sectoringWip || totalLenses, totalCap);

  setText("chambersActive", activeCount);
  setText("chamberUtilText", `Utilization: ${utilPct}%`);
  setText("bigDonutPct", `${utilPct}%`);
  setText("bigDonutFrac", `${fmt(sectoringWip || totalLenses)} / ${fmt(totalCap)}`);

  const bar = $("chamberUtilBar");
  if (bar) bar.style.width = `${utilPct}%`;

  const donut = $("bigDonutFill");
  if (donut) {
    donut.setAttribute("stroke-dashoffset", String(263.9 * (1 - utilPct / 100)));
  }
}

function renderUtilList(chambers, chamberCap) {
  const box = $("utilList");
  if (!box) return;

  box.innerHTML = "";

  const rows = Math.max(TOTAL_VISIBLE_CHAMBERS, chambers.length || 0);

  for (let i = 0; i < rows; i++) {
    const ch = chambers[i];
    const used = toNumber(ch?.Used);
    const cap = toNumber(ch?.Capacity) || chamberCap;
    const p = pct(used, cap);

    const row = makeDiv("util-row");

    row.innerHTML = `
      <span>C${i + 1}</span>
      <div class="util-track">
        <div class="util-fill" style="width:${p}%"></div>
      </div>
      <strong>${p}%</strong>
    `;

    box.appendChild(row);
  }
}

/* ─────────────────────────────────────────────────────
   SUMMARY / ALERTS
───────────────────────────────────────────────────── */
function renderAlerts(values) {
  if (!values) return;

  const alerts = [];
  const chamberUtil = pct(values.sectoringWip, Math.max(values.chamberTotalCapacity, 1));
  const ovenUtil = pct(values.ovenWip, Math.max(values.ovenTotalLensCapacity, 1));

  if (chamberUtil >= 85) {
    alerts.push({
      level: "red",
      title: "Sectoring Utilization Critical",
      desc: `Sectoring utilization is ${chamberUtil}%. Watch AR coating backlog.`
    });
  } else if (chamberUtil >= 70) {
    alerts.push({
      level: "amber",
      title: "Sectoring Utilization High",
      desc: `Sectoring utilization is ${chamberUtil}%. Keep load balanced.`
    });
  }

  if (ovenUtil >= 85) {
    alerts.push({
      level: "red",
      title: "Oven Load Critical",
      desc: `Oven lens utilization is ${ovenUtil}%. Degas can become the blocker.`
    });
  } else if (ovenUtil >= 70) {
    alerts.push({
      level: "amber",
      title: "Oven Load High",
      desc: `Oven lens utilization is ${ovenUtil}%. Monitor the next cycle.`
    });
  }

  if (values.surfaceInputWip > values.arInWip && values.surfaceInputWip > 100) {
    alerts.push({
      level: "amber",
      title: "Surface Input Pressure",
      desc: "Surface Inspection input is building faster than AR-IN movement."
    });
  }

  if (values.deringWip <= 0 && values.sectoringWip > 0) {
    alerts.push({
      level: "amber",
      title: "No DeRing WIP",
      desc: "Sectoring has WIP, but DeRing is empty. Confirm if unload movement is expected."
    });
  }

  if (!alerts.length) {
    alerts.push({
      level: "green",
      title: "AR Flow Stable",
      desc: "No active capacity pressure alerts detected."
    });
  }

  const list = $("alertList");

  if (list) {
    list.innerHTML = alerts.map(a => `
      <div class="alert">
        <div class="alert-ico ${a.level}">${a.level === "green" ? "✓" : "!"}</div>

        <div class="alert-body">
          <div class="alert-title">
            <strong>${escapeHtml(a.title)}</strong>
            <span class="t">Live</span>
          </div>

          <div class="alert-desc">${escapeHtml(a.desc)}</div>
        </div>
      </div>
    `).join("");
  }

  const active = alerts.filter(a => a.level !== "green").length;

  setText("alertBadge", active);
  setText("activeAlertCount", `${active} Active`);
  setText("alertPageBadge", `${active} Active`);
}

function renderDailyBrief(values) {
  const box = $("dailyBriefList");
  if (!box || !values) return;

  const largest = values.largestWip || getLargestWipStation(LAST_AR_ROWS);

  box.innerHTML = `
    <div class="brief-item station-brief">
      <strong>${fmt(values.totalArWip)}</strong>
      <span>Currently AR Work In Progress.</span>
    </div>

    <div class="brief-item station-brief">
      <strong>${escapeHtml(largest.name)}</strong>
      <span>Largest WIP station with ${fmt(largest.value)} lenses.</span>
    </div>

    <div class="brief-item station-brief">
      <strong>${fmt(values.arOutActivity)}</strong>
      <span>AR-OUT completed movement toward Finish.</span>
    </div>

    <div class="brief-item station-brief">
      <strong>${fmt(values.activeChambers)}</strong>
      <span>Active sectoring chambers.</span>
    </div>
  `;
}

function renderSummaryBreakdown(values) {
  if (!values || !LAST_AR_ROWS.length) return;

  const activityBox = $("summaryStationActivity");
  const wipBox = $("summaryWipPressure");

  const activityRows = getCleanStationRows(getOutputRows(LAST_AR_ROWS));
  const wipRows = getCleanStationRows(getWipRows(LAST_AR_ROWS));

  const maxActivity = Math.max(...activityRows.map(row => toNumber(row.ActivityToday)), 1);
  const maxWip = Math.max(...wipRows.map(row => toNumber(row.CurrentWIP)), 1);

  const topStation = activityRows.reduce((best, row) => {
    return toNumber(row.ActivityToday) > toNumber(best.ActivityToday) ? row : best;
  }, activityRows[0] || {});

  const lowestStation = activityRows.reduce((lowest, row) => {
    return toNumber(row.ActivityToday) < toNumber(lowest.ActivityToday) ? row : lowest;
  }, activityRows[0] || {});

  const largestWip = getLargestWipStation(LAST_AR_ROWS);

  if (activityBox) {
    activityBox.innerHTML = activityRows.map(row => {
      const name = row.DisplayName || row.FlowStep || "";
      const value = toNumber(row.ActivityToday);
      const width = Math.max(4, Math.round((value / maxActivity) * 100));

      return `
        <div class="summary-row">
          <div class="summary-name">${escapeHtml(name)}</div>
          <div class="summary-track">
            <div class="summary-fill" style="width:${width}%; background:${getStationColor(name)};"></div>
          </div>
          <div class="summary-value">${fmt(value)}</div>
        </div>
      `;
    }).join("");
  }

  if (wipBox) {
    wipBox.innerHTML = wipRows.map(row => {
      const name = row.DisplayName || row.FlowStep || "";
      const value = toNumber(row.CurrentWIP);
      const width = Math.max(4, Math.round((value / maxWip) * 100));

      return `
        <div class="summary-row">
          <div class="summary-name">${escapeHtml(name)}</div>
          <div class="summary-track">
            <div class="summary-fill wip" style="width:${width}%; background:${getStationColor(name)};"></div>
          </div>
          <div class="summary-value">${fmt(value)}</div>
        </div>
      `;
    }).join("");
  }

  setText("summaryTopStation", topStation.DisplayName || topStation.FlowStep || "--");
  setText("summaryTopStationText", `${fmt(topStation.ActivityToday)} completed scans today.`);

  setText("summaryLowestStation", lowestStation.DisplayName || lowestStation.FlowStep || "--");
  setText("summaryLowestStationText", `${fmt(lowestStation.ActivityToday)} completed scans today. Review if this station should be higher.`);

  setText("summaryArOut", fmt(values.arOutActivity));

  const balanceStatus =
    values.ovenWip > values.sectoringWip && values.ovenWip > values.basketWip
      ? "Oven Pressure"
      : values.sectoringWip > values.ovenWip
        ? "Sectoring WIP"
        : "Balanced";

  const balanceText =
    balanceStatus === "Balanced"
      ? "No single station is dominating the current WIP load."
      : `${largestWip.name} is currently carrying the largest WIP pressure.`;

  setText("summaryBalanceStatus", balanceStatus);
  setText("summaryBalanceText", balanceText);
}

/* ─────────────────────────────────────────────────────
   CHARTS
───────────────────────────────────────────────────── */
function renderCharts(flowPayload, arRows, values) {
  renderTrendCharts(flowPayload, arRows, values);
}

function renderTrendCharts(flowPayload, arRows, values) {
  if (typeof Chart === "undefined") return;

  const activityRows = getCleanStationRows(getOutputRows(arRows));
  const wipRows = getCleanStationRows(getWipRows(arRows));

  chartInstances.outputTrend = makeChart("outputTrendChart", chartInstances.outputTrend, {
    type: "bar",
    data: {
      labels: activityRows.map(row => row.DisplayName || row.FlowStep),
      datasets: [
        {
          label: "Station Activity",
          data: activityRows.map(row => toNumber(row.ActivityToday)),
          backgroundColor: activityRows.map(row => getStationColor(row.DisplayName || row.FlowStep)),
          borderColor: activityRows.map(row => getStationColor(row.DisplayName || row.FlowStep)),
          borderWidth: 1,
          borderRadius: 8,
          barThickness: 42,
          maxBarThickness: 50
        }
      ]
    },
    options: chartOptions("Station Activity Trend")
  });

  chartInstances.wipTrend = makeChart("wipTrendChart", chartInstances.wipTrend, {
    type: "bar",
    data: {
      labels: wipRows.map(row => row.DisplayName || row.FlowStep),
      datasets: [
        {
          label: "Current WIP",
          data: wipRows.map(row => toNumber(row.CurrentWIP)),
          backgroundColor: wipRows.map(row => getStationColor(row.DisplayName || row.FlowStep)),
          borderColor: wipRows.map(row => getStationColor(row.DisplayName || row.FlowStep)),
          borderWidth: 1,
          borderRadius: 8,
          barThickness: 42,
          maxBarThickness: 50
        }
      ]
    },
    options: chartOptions("Work In Progress Trend")
  });

  const largest = getLargestWipStation(wipRows);
  renderFlowRiskSummary(values, largest.name);

  const box = $("trendBriefList");

  if (box) {
    const topActivity = activityRows.reduce((best, row) => {
      return toNumber(row.ActivityToday) > toNumber(best.ActivityToday) ? row : best;
    }, activityRows[0] || {});

    box.innerHTML = `
      <div class="brief-item station-brief">
        <strong>${escapeHtml(largest.name)}</strong>
        <span>Current largest WIP station with ${fmt(largest.value)} lenses.</span>
      </div>

      <div class="brief-item station-brief">
        <strong>${escapeHtml(topActivity.DisplayName || topActivity.FlowStep || "--")}</strong>
        <span>Highest station activity today with ${fmt(topActivity.ActivityToday)} scans.</span>
      </div>

      <div class="brief-item station-brief">
        <strong>${fmt(values.arOutActivity)}</strong>
        <span>AR-OUT completed movement toward Finish.</span>
      </div>
    `;
  }
}

function renderFlowRiskSummary(values, largestStation) {
  const box = $("flowRiskList");

  if (!box || !values) return;

  const risks = [];

  const ovenBasketPct = pct(values.ovenBasketLoad, AR_CAPACITY_RULES.OVEN_BASKETS);
  const chamberPct = pct(
    values.sectoringWip,
    Math.max(values.chamberTotalCapacity, AR_CAPACITY_RULES.CHAMBER_LENS)
  );

  risks.push({
    title: largestStation,
    text: "Largest current WIP pressure point.",
    level: "watch"
  });

  risks.push({
    title: `${values.fullBasketCount} full · ${values.partialBasketCount} partial baskets`,
    text: `${fmt(values.basketWip)} total basket-stage lens WIP.`,
    level: values.partialBasketCount > 0 ? "active" : "stable"
  });

  risks.push({
    title: `${fmt(values.ovenBasketLoad)} / ${AR_CAPACITY_RULES.OVEN_BASKETS} oven basket load`,
    text: `Oven is ${ovenBasketPct}% loaded by basket equivalent.`,
    level: ovenBasketPct >= 70 ? "watch" : "stable"
  });

  risks.push({
    title: `${fmt(values.activeChambers)} active sectoring chambers`,
    text: `Sectoring utilization is ${chamberPct}%.`,
    level: chamberPct >= 70 ? "watch" : "stable"
  });

  if (values.deringWip <= 0 && values.sectoringWip > 0) {
    risks.push({
      title: "DeRing has no current WIP",
      text: "Confirm if chamber unload movement is expected or delayed.",
      level: "watch"
    });
  }

  box.innerHTML = risks.map(item => `
    <div class="flow-risk-item ${item.level}">
      <strong>${escapeHtml(item.title)}</strong>
      <span>${escapeHtml(item.text)}</span>
    </div>
  `).join("");

  setText(
    "flowRiskStatus",
    risks.some(r => r.level === "watch") ? "Needs Review" : "Stable"
  );
}

function makeChart(canvasId, existingChart, config) {
  const canvas = $(canvasId);

  if (!canvas) return existingChart;

  if (existingChart) {
    existingChart.destroy();
  }

  return new Chart(canvas, config);
}

function chartOptions(title) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: {
      duration: 700,
      easing: "easeOutQuart"
    },
    plugins: {
      legend: {
        labels: {
          color: "#8b8f9d",
          font: {
            size: 11,
            weight: "700"
          }
        }
      },
      title: {
        display: false,
        text: title,
        color: "#f5f5f7"
      }
    },
    scales: {
      x: {
        ticks: {
          color: "#8b8f9d",
          font: { size: 10 }
        },
        grid: {
          color: "rgba(255,255,255,.04)"
        }
      },
      y: {
        beginAtZero: true,
        ticks: {
          color: "#8b8f9d",
          font: { size: 10 }
        },
        grid: {
          color: "rgba(255,255,255,.05)"
        }
      }
    }
  };
}

function resizeCharts() {
  Object.values(chartInstances).forEach(chart => {
    if (chart && typeof chart.resize === "function") {
      chart.resize();
    }
  });
}

/* ─────────────────────────────────────────────────────
   NAVIGATION / CLOCK / EVENTS
───────────────────────────────────────────────────── */
function bindPageNavigation() {
  document.querySelectorAll(".nav-item").forEach(btn => {
    btn.addEventListener("click", () => {
      const section = btn.dataset.section;

      document.querySelectorAll(".nav-item").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");

      document.querySelectorAll(".page-section").forEach(panel => {
        panel.classList.remove("active");
        panel.hidden = true;
      });

      const target = document.querySelector(`.page-section[data-page="${section}"]`);

      if (target) {
        target.hidden = false;

        requestAnimationFrame(() => {
          target.classList.add("active");
        });
      }

      if (section === "flow") {
        setTimeout(placeArrows, 80);
      }

      if (section === "trends") {
        setTimeout(resizeCharts, 120);
      }

      if (section === "summary") {
        renderSummaryBreakdown(LAST_VALUES);
      }

      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  });
}

function bindSplitTabs() {
  document.querySelectorAll(".ar-split-tab").forEach(button => {
    button.addEventListener("click", () => {
      const panel = $("arSplitDetailPanel");

      if (!panel) return;

      panel.querySelectorAll(".ar-split-tab").forEach(btn => {
        btn.classList.remove("active");
      });

      button.classList.add("active");
      panel.dataset.activeTab = button.dataset.tab || "wip";

      renderSplitDetailTabs(LAST_AR_ROWS, LAST_FLOW_PAYLOAD?.summary || {});
    });
  });
}

function placeArrows() {
  const grid = $("stations");
  if (!grid) return;

  const activePage = document.querySelector(".page-section.active")?.dataset.page;
  if (activePage !== "flow") return;

  const stationCards = grid.querySelectorAll(".station");
  grid.querySelectorAll(".flow-arrow").forEach(a => a.remove());

  if (window.innerWidth < 1500) return;

  const containerRect = grid.getBoundingClientRect();

  const arrowSVG = `
    <svg viewBox="0 0 24 14" preserveAspectRatio="none">
      <defs>
        <linearGradient id="arrowGrad" x1="0" y1="0" x2="1" y2="0">
          <stop offset="0%" stop-color="#ff9900" stop-opacity=".25"/>
          <stop offset="100%" stop-color="#ff9900" stop-opacity="1"/>
        </linearGradient>
      </defs>
      <path d="M0 7 L18 7 M14 2 L20 7 L14 12" stroke="url(#arrowGrad)" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
      <path d="M4 7 L20 7" stroke="#ff9900" stroke-width="1" stroke-dasharray="3 3" opacity=".6" fill="none"/>
    </svg>
  `;

  for (let i = 0; i < stationCards.length - 1; i++) {
    const a = stationCards[i].getBoundingClientRect();
    const b = stationCards[i + 1].getBoundingClientRect();

    const div = document.createElement("div");
    div.className = "arrow flow-arrow";
    div.innerHTML = arrowSVG;
    div.style.left = `${a.right - containerRect.left - 5}px`;
    div.style.width = `${Math.max(12, b.left - a.right + 10)}px`;

    grid.appendChild(div);
  }
}

function updateClock() {
  const now = new Date();

  setText("todayDate", now.toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric"
  }));

  setText("clockTime", now.toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit"
  }));
}

function animateValue(id, target) {
  const el = $(id);

  if (!el) return;

  const start = toNumber(el.textContent);
  const end = toNumber(target);

  const duration = 600;
  const startTime = performance.now();

  function update(currentTime) {
    const elapsed = currentTime - startTime;
    const progress = Math.min(elapsed / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3);
    const value = Math.round(start + (end - start) * eased);

    el.textContent = fmt(value);

    if (progress < 1) {
      requestAnimationFrame(update);
    }
  }

  requestAnimationFrame(update);
}

function bindEvents() {
  const refresh = $("refreshBtn");

  if (refresh) {
    refresh.addEventListener("click", loadARData);
  }

  window.addEventListener("resize", () => {
    placeArrows();
    resizeCharts();
  });

  bindPageNavigation();
  bindSplitTabs();
}

window.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  updateClock();
  loadARData();

  setInterval(updateClock, 1000 * 30);
  setInterval(loadARData, REFRESH_MS);
});