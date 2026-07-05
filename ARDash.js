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

const USER_CONTROL_API_URL =
  "https://script.google.com/macros/s/AKfycbwhcxW8dfFw_gJW2s0aZoUP3nhilqEA0S6S-9W9lvPpOtQaWLjAwfFSmo5HrsheP5jR/exec";


const REFRESH_MS = 5 * 60 * 1000;
const TOTAL_VISIBLE_CHAMBERS = 6;

const AR_CAPACITY_RULES = {
  BASKET_LENS: 32,
  OVEN_BASKETS: 27,
  OVEN_LENS: 864,

  // Sectoring rules:
  // CurrentWIP is scanned JOB WIP, not lens count.
  // 1 chamber holds 84 jobs.
  // 1 job normally equals 2 lenses.
  CHAMBER_JOB: 84,
  CHAMBER_LENS: 168,
  LENSES_PER_JOB: 2
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

const AR_METRICS_STATE = {
  owner: "BLOPEZ",
  shift: "Weekday",
  rosterPayload: null,
  metricsPayload: null,
  isSaving: false,
  selectedOperator: null,
  selectedStation: "ALL"
};

const AR_ROLE_ORDER = ["AR-IN", "Basket", "Oven", "Sectoring", "DeRing", "AR-OUT"];


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

function floorUnit(value, capacity) {
  const v = toNumber(value);
  const c = toNumber(capacity);

  if (v <= 0 || c <= 0) return 0;

  return Math.floor(v / c);
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

function getAROwner() {
  const input = $("arMetricOwner");
  return String(input?.value || AR_METRICS_STATE.owner || "BLOPEZ").trim().toUpperCase() || "BLOPEZ";
}

function getARShift() {
  const select = $("arMetricShift");
  return String(select?.value || AR_METRICS_STATE.shift || "Weekday").trim() || "Weekday";
}


function parseARStoredUserValue(raw) {
  if (raw == null || raw === "") return null;
  if (typeof raw === "object") return raw;

  const text = String(raw).trim();
  if (!text) return null;

  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function readARNestedField(obj, keys) {
  if (!obj || typeof obj !== "object") return "";

  for (const key of keys) {
    if (obj[key] != null && String(obj[key]).trim() !== "") {
      return String(obj[key]).trim();
    }
  }

  const nestedObjects = [obj.user, obj.profile, obj.account, obj.auth, obj.data, obj.currentUser];
  for (const nested of nestedObjects) {
    if (!nested || typeof nested !== "object") continue;
    for (const key of keys) {
      if (nested[key] != null && String(nested[key]).trim() !== "") {
        return String(nested[key]).trim();
      }
    }
  }

  return "";
}

function getARCurrentUserProfile() {
  const objectKeys = [
    "lms_user",
    "lmsUser",
    "currentUser",
    "current_user",
    "loggedInUser",
    "logged_in_user",
    "userProfile",
    "authUser",
    "user"
  ];

  const stringRoleKeys = [
    "lms_role",
    "lmsRole",
    "LMS_ROLE",
    "currentUserRole",
    "loggedInRole",
    "userRole",
    "role",
    "Role"
  ];

  const stringUserKeys = [
    "lms_user",
    "lmsUser",
    "lms_username",
    "lmsUsername",
    "LMS_USERNAME",
    "currentUsername",
    "loggedInUsername",
    "username",
    "Username"
  ];

  let username = "";
  let role = "";
  let displayName = "";

  for (const key of objectKeys) {
    const raw = sessionStorage.getItem(key) || localStorage.getItem(key);
    const parsed = parseARStoredUserValue(raw);

    if (!parsed) continue;

    if (typeof parsed === "string") {
      if (!username && key.toLowerCase().includes("user")) username = parsed;
      continue;
    }

    username = username || readARNestedField(parsed, ["username", "Username", "userName", "UserName", "login", "Login", "email", "Email"]);
    role = role || readARNestedField(parsed, ["role", "Role", "userRole", "UserRole", "accessRole", "AccessRole"]);
    displayName = displayName || readARNestedField(parsed, ["displayName", "DisplayName", "name", "Name", "fullName", "FullName"]);
  }

  for (const key of stringRoleKeys) {
    const raw = sessionStorage.getItem(key) || localStorage.getItem(key);
    const parsed = parseARStoredUserValue(raw);
    if (typeof parsed === "string" && parsed.trim()) {
      role = role || parsed.trim();
    } else if (parsed && typeof parsed === "object") {
      role = role || readARNestedField(parsed, ["role", "Role", "userRole", "UserRole", "accessRole", "AccessRole"]);
    }
  }

  for (const key of stringUserKeys) {
    const raw = sessionStorage.getItem(key) || localStorage.getItem(key);
    const parsed = parseARStoredUserValue(raw);
    if (typeof parsed === "string" && parsed.trim()) {
      username = username || parsed.trim();
    } else if (parsed && typeof parsed === "object") {
      username = username || readARNestedField(parsed, ["username", "Username", "userName", "UserName", "login", "Login", "email", "Email"]);
    }
  }

  return {
    username: String(username || "").trim().toUpperCase(),
    role: String(role || "").trim(),
    displayName: String(displayName || "").trim()
  };
}

function isARLMSUser() {
  const profile = getARCurrentUserProfile();
  const role = String(profile.role || "").trim().toLowerCase();
  const username = String(profile.username || "").trim().toUpperCase();
  const selectedOwner = String(getAROwner ? getAROwner() : "").trim().toUpperCase();

  // BLOPEZ fallback keeps Brian's LMS setup visible even if the login script only saved username.
  // The selected Profile dropdown is also checked because AR Metrics is profile-driven.
  return (
    role === "lms" ||
    username === "BLOPEZ" ||
    username === "BRIAN LOPEZ CABRERA" ||
    selectedOwner === "BLOPEZ"
  );
}

function applyARLMSOnlyVisibility() {
  const canSee = isARLMSUser();
  const panel = $("arLmsRosterAdminPanel");
  const page = $("arLmsSetupPage");
  const nav = $("arLmsSetupNav");

  [panel, page, nav].forEach(el => {
    if (!el) return;
    el.hidden = !canSee;
    el.classList.toggle("is-hidden", !canSee);
    el.setAttribute("aria-hidden", canSee ? "false" : "true");
  });

  if (!canSee && document.querySelector('.page-section.active')?.dataset.page === "lmsSetup") {
    const overviewBtn = document.querySelector('.nav-item[data-section="overview"]');
    if (overviewBtn) overviewBtn.click();
  }

  const assignBtn = $("arAssignBtn");
  if (assignBtn) assignBtn.disabled = !canSee;

  document.body.classList.toggle("ar-is-lms", canSee);
  document.body.classList.toggle("ar-not-lms", !canSee);

  return canSee;
}

async function fetchARRosterControl() {
  const owner = encodeURIComponent(getAROwner());
  const shift = encodeURIComponent(getARShift());
  return fetchJson(`${API_URL}?action=getArRosterControl&owner=${owner}&shift=${shift}`);
}

async function fetchARCapacity() {
  const owner = encodeURIComponent(getAROwner());
  const shift = encodeURIComponent(getARShift());
  return fetchJson(`${API_URL}?action=getArCapacityMetrics&owner=${owner}&shift=${shift}`);
}

async function saveARRosterAssignment(payload) {
  const params = new URLSearchParams({
    action: "saveArRosterControl",
    owner: getAROwner(),
    shift: getARShift(),
    operatorName: payload.operatorName || "",
    role: payload.role || "AR-IN",
    defaultRole: payload.role || "AR-IN",
    certificationStatus: payload.certificationStatus || "Certified",
    trainingWeek: String(payload.trainingWeek || 1),
    individualJPH: String(payload.individualJPH || 0),
    updatedBy: getAROwner()
  });

  return fetchJson(`${API_URL}?${params.toString()}`);
}

async function clearARRosterAssignment(operatorName, role) {
  const params = new URLSearchParams({
    action: "clearArRosterAssignment",
    owner: getAROwner(),
    shift: getARShift(),
    operatorName: operatorName || "",
    role: role || "",
    defaultRole: role || "",
    updatedBy: getAROwner()
  });

  return fetchJson(`${API_URL}?${params.toString()}`);
}

async function loadARData() {
  try {
    const [flowPayload, capacityPayload, rosterPayload] = await Promise.all([
      fetchProductionFlowAR(),
      fetchARCapacity().catch(err => {
        console.warn("AR capacity metrics failed:", err);
        return null;
      }),
      fetchARRosterControl().catch(err => {
        console.warn("AR roster control failed:", err);
        return null;
      })
    ]);

    LAST_FLOW_PAYLOAD = flowPayload;
    LAST_CAPACITY_PAYLOAD = capacityPayload || null;
    AR_METRICS_STATE.metricsPayload = capacityPayload || null;
    AR_METRICS_STATE.rosterPayload = rosterPayload || null;
    AR_METRICS_STATE.owner = getAROwner();
    AR_METRICS_STATE.shift = getARShift();

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

  // CRITICAL FIX:
  // Sectoring CurrentWIP is scanned job WIP.
  // Chamber visuals must be calculated by 84 jobs per chamber.
  // Lens equivalent is only supporting text: jobs × 2.
  const sectoringWip = toNumber(sectoringRow.CurrentWIP);
  const sectoringLensLoad = sectoringWip * AR_CAPACITY_RULES.LENSES_PER_JOB;

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

  // Sectoring chamber build uses JOBS, not lenses.
  const chamberUnits = buildUnitsFromWip(
    sectoringWip,
    AR_CAPACITY_RULES.CHAMBER_JOB,
    "Chamber"
  );

  const fullChambers = chamberUnits.filter(u =>
    toNumber(u.Used) >= AR_CAPACITY_RULES.CHAMBER_JOB
  ).length;

  const partialChambers = chamberUnits.filter(u =>
    toNumber(u.Used) > 0 && toNumber(u.Used) < AR_CAPACITY_RULES.CHAMBER_JOB
  ).length;

  // Header should show active/full chambers only.
  // Example: 298 WIP = 3 active chambers + 1 partial chamber.
  const activeChambers = fullChambers;

  // Total loaded chamber cards include full + partial.
  const loadedChambers = chamberUnits.length;

  const chamberTotalCapacity = Math.max(
    loadedChambers * AR_CAPACITY_RULES.CHAMBER_JOB,
    AR_CAPACITY_RULES.CHAMBER_JOB
  );

  const lastBasket = basketUnits.length
    ? basketUnits[basketUnits.length - 1]
    : { Used: 0, Capacity: AR_CAPACITY_RULES.BASKET_LENS };

  const lastChamber = chamberUnits.length
    ? chamberUnits[chamberUnits.length - 1]
    : { Used: 0, Capacity: AR_CAPACITY_RULES.CHAMBER_JOB };

  const fullBasketCount = basketUnits.filter(u =>
    toNumber(u.Used) >= toNumber(u.Capacity)
  ).length;

  const partialBasketCount = basketUnits.filter(u =>
    toNumber(u.Used) > 0 && toNumber(u.Used) < toNumber(u.Capacity)
  ).length;

  const lastOvenBasketText = `${ovenBasketLoad} / ${AR_CAPACITY_RULES.OVEN_BASKETS}`;

  const ovenLensRemaining = Math.max(0, ovenTotalLensCapacity - ovenWip);
  const chamberJobRemaining = Math.max(0, chamberTotalCapacity - sectoringWip);
  const chamberLensRemaining = chamberJobRemaining * AR_CAPACITY_RULES.LENSES_PER_JOB;

  const largestWip = getLargestWipStation(arRows);

  LAST_VALUES = {
    surfaceInputWip,
    arInWip,
    basketWip,
    ovenWip,
    sectoringWip,
    sectoringLensLoad,
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
    fullChambers,
    partialChambers,
    loadedChambers,
    chamberTotalCapacity,
    chamberJobRemaining,
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

  setText(
    "chamberLensText",
    `${fmt(sectoringWip)} WIP jobs × 2 = ${fmt(sectoringLensLoad)} lenses | ${AR_CAPACITY_RULES.CHAMBER_JOB} jobs per chamber`
  );

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

  setText("capChamberTotalWip", `${fmt(sectoringWip)} jobs / ${fmt(sectoringLensLoad)} lenses`);
  setText("capChamberActive", `${fmt(fullChambers)} full · ${fmt(partialChambers)} partial`);
  setText("capChamberCurrent", `${fmt(lastChamber.Used)} / ${fmt(lastChamber.Capacity)} jobs`);
  setText("capChamberRemaining", `${fmt(chamberJobRemaining)} jobs / ${fmt(chamberLensRemaining)} lenses`);

  renderSplitDetailTabs(arRows, flowSummary);
  renderBaskets(basketUnits, AR_CAPACITY_RULES.BASKET_LENS);
  renderOvens(ovenWip, ovenBasketLoad, ovenTotalLensCapacity);
  renderChambers(chamberUnits, AR_CAPACITY_RULES.CHAMBER_JOB, sectoringWip, chamberTotalCapacity, fullChambers);
  renderUtilList(chamberUnits, AR_CAPACITY_RULES.CHAMBER_JOB);
  renderAlerts(LAST_VALUES);
  renderDailyBrief(LAST_VALUES);
  renderSummaryBreakdown(LAST_VALUES);
  renderARMetricsPanel(LAST_CAPACITY_PAYLOAD, AR_METRICS_STATE.rosterPayload);
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

function renderChambers(units, chamberCap, sectoringWip, chamberTotalCapacity, fullChambers) {
  const grid = $("chamberGrid");
  if (!grid) return;

  grid.innerHTML = "";

  const visibleCount = Math.max(TOTAL_VISIBLE_CHAMBERS, units.length || 0);
  let loadedCount = 0;
  let totalJobs = 0;

  for (let i = 0; i < visibleCount; i++) {
    const unit = units[i];
    const used = toNumber(unit?.Used);
    const cap = toNumber(unit?.Capacity) || chamberCap;

    const state = used <= 0
      ? "empty"
      : used >= cap
        ? "full active"
        : "partial";

    if (used > 0) loadedCount++;
    totalJobs += used;

    const lensUsed = used * AR_CAPACITY_RULES.LENSES_PER_JOB;
    const lensCap = cap * AR_CAPACITY_RULES.LENSES_PER_JOB;

    const div = makeDiv(`chamber ${state}`);

    div.innerHTML = `
      <div class="chamber-num">C${i + 1}</div>
      <div class="chamber-count">${used > 0 ? `${used} / ${cap}` : "— / " + cap}</div>
      <div class="chamber-label">${used <= 0 ? "Empty" : used >= cap ? "Full" : "Partial"}</div>
      <div class="chamber-lens">${used > 0 ? `${lensUsed} / ${lensCap} lenses` : ""}</div>
    `;

    grid.appendChild(div);
  }

  const totalCap = Math.max(toNumber(chamberTotalCapacity), chamberCap);
  const utilPct = pct(sectoringWip || totalJobs, totalCap);

  // Shows full active chambers only.
  // Example: 298 jobs = 3 active/full chambers and 1 partial chamber.
  setText("chambersActive", fullChambers);
  setText("chamberUtilText", `Utilization: ${utilPct}% | ${fmt(loadedCount)} loaded chambers`);
  setText("bigDonutPct", `${utilPct}%`);
  setText("bigDonutFrac", `${fmt(sectoringWip || totalJobs)} / ${fmt(totalCap)} jobs`);

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
      <span>Largest WIP station with ${fmt(largest.value)} WIP jobs.</span>
    </div>

    <div class="brief-item station-brief">
      <strong>${fmt(values.arOutActivity)}</strong>
      <span>AR-OUT completed movement toward Finish.</span>
    </div>

    <div class="brief-item station-brief">
      <strong>${fmt(values.fullChambers)}</strong>
      <span>Full active sectoring chambers. ${fmt(values.partialChambers)} partial chamber loaded.</span>
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


function normalizeARText(value, fallback = "--") {
  const text = String(value ?? "").trim();
  return text || fallback;
}

function getRosterOperatorName(row) {
  return normalizeARText(row.operatorName || row.OperatorName || row.name || row.Name, "");
}

function getRosterRole(row) {
  return normalizeARText(row.defaultRole || row.DefaultRole || row.role || row.Role, "AR-IN");
}

function getRosterFinalJph(row) {
  return toNumber(row.finalJph || row.FinalJPH || row.jph || row.JPH);
}

function getRosterCert(row) {
  return normalizeARText(row.certificationStatus || row.CertificationStatus || "Certified", "Certified");
}

function getOperatorMasterRows(rosterPayload, metricsPayload) {
  const fromMaster = Array.isArray(rosterPayload?.operatorMaster) ? rosterPayload.operatorMaster : [];
  const fromActivity = Array.isArray(metricsPayload?.operatorActivity) ? metricsPayload.operatorActivity : [];

  const map = new Map();

  fromMaster.forEach(row => {
    const name = getRosterOperatorName(row);
    if (!name) return;
    map.set(name.toUpperCase(), {
      operatorName: name,
      defaultRole: normalizeARText(row.defaultRole || row.DefaultRole || row.role || row.Role, "AR-IN")
    });
  });

  fromActivity.forEach(row => {
    const name = normalizeARText(row.Operator || row.operator || row.operatorName, "");
    if (!name) return;
    const station = normalizeARText(row.FlowStation || row.flowStation || row.AccessPoint || row.accessPoint, "AR-IN");
    if (!map.has(name.toUpperCase())) {
      map.set(name.toUpperCase(), {
        operatorName: name,
        defaultRole: station
      });
    }
  });

  return Array.from(map.values()).sort((a, b) => a.operatorName.localeCompare(b.operatorName));
}


function normalizeARStationFilter(value) {
  const clean = normalizeARText(value, "ALL");
  if (!clean || clean.toUpperCase() === "ALL") return "ALL";

  const upper = clean.toUpperCase();
  const match = AR_ROLE_ORDER.find(role => role.toUpperCase() === upper);
  if (match) return match;

  if (upper.includes("AR-IN") || upper.includes("AR IN")) return "AR-IN";
  if (upper.includes("BASKET")) return "Basket";
  if (upper.includes("OVEN")) return "Oven";
  if (upper.includes("SECTOR")) return "Sectoring";
  if (upper.includes("DERING") || upper.includes("DE RING")) return "DeRing";
  if (upper.includes("AR-OUT") || upper.includes("AR OUT")) return "AR-OUT";

  return clean;
}

function getSelectedARStation() {
  return normalizeARStationFilter(AR_METRICS_STATE.selectedStation || "ALL");
}

function isARStationMatch(value, selectedStation = getSelectedARStation()) {
  const selected = normalizeARStationFilter(selectedStation);
  if (selected === "ALL") return true;
  return normalizeARStationFilter(value).toUpperCase() === selected.toUpperCase();
}

function getARStationMetricRows(metricsPayload) {
  return Array.isArray(metricsPayload?.stationMetrics) ? metricsPayload.stationMetrics : [];
}

function getAvailableARStationFilters(metricsPayload) {
  const found = new Set();

  getARStationMetricRows(metricsPayload).forEach(row => {
    const role = normalizeARStationFilter(row.role || row.station);
    if (role && role !== "ALL") found.add(role);
  });

  (metricsPayload?.operatorActivity || []).forEach(row => {
    const station = normalizeARStationFilter(readStationFromActivity(row));
    if (station && station !== "ALL") found.add(station);
  });

  return AR_ROLE_ORDER.filter(role => found.has(role));
}

function renderARStationFilterTabs(metricsPayload) {
  const box = $("arStationFilterTabs");
  const label = $("arSelectedStationLabel");
  if (!box) return;

  const selected = getSelectedARStation();
  const stations = getAvailableARStationFilters(metricsPayload);
  const buttons = ["ALL", ...stations];

  if (selected !== "ALL" && !buttons.some(v => v.toUpperCase() === selected.toUpperCase())) {
    AR_METRICS_STATE.selectedStation = "ALL";
  }

  const activeStation = getSelectedARStation();
  if (label) label.textContent = activeStation === "ALL" ? "All AR" : activeStation;

  box.innerHTML = buttons.map(station => {
    const text = station === "ALL" ? "All AR" : station;
    const active = normalizeARStationFilter(station).toUpperCase() === activeStation.toUpperCase() ? "active" : "";
    return `<button class="${active}" type="button" data-ar-station-filter="${escapeHtml(station)}">${escapeHtml(text)}</button>`;
  }).join("");

  box.querySelectorAll("[data-ar-station-filter]").forEach(btn => {
    btn.addEventListener("click", () => {
      AR_METRICS_STATE.selectedStation = normalizeARStationFilter(btn.dataset.arStationFilter || "ALL");
      AR_METRICS_STATE.selectedOperator = null;
      closeARAssociateDrawer();
      renderARMetricsPanel(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {}, AR_METRICS_STATE.rosterPayload || {});
    });
  });
}

function calculateStationPeakFromOperators(operatorList) {
  const hourTotals = {};
  operatorList.forEach(item => {
    Object.entries(item.hours || {}).forEach(([hour, value]) => {
      hourTotals[hour] = toNumber(hourTotals[hour]) + toNumber(value);
    });
  });

  const entries = Object.entries(hourTotals);
  if (!entries.length) return { hour: "--", value: 0 };

  const top = entries.reduce((best, cur) => toNumber(cur[1]) > toNumber(best[1]) ? cur : best, entries[0]);
  return { hour: top[0], value: toNumber(top[1]) };
}

function renderARMetricsPanel(metricsPayload, rosterPayload) {
  renderARStationFilterTabs(metricsPayload);
  renderARMetricSummary(metricsPayload);
  renderARStationMetrics(metricsPayload);
  renderAROperatorPerformance(metricsPayload);
  renderARRosterControls(metricsPayload, rosterPayload);

  if (AR_METRICS_STATE.selectedOperator) {
    const stillExists = buildAROperatorRollups(metricsPayload || {}).some(item => item.name.toUpperCase() === AR_METRICS_STATE.selectedOperator.toUpperCase());
    if (stillExists && $("arAssociateDrawer")?.classList.contains("open")) {
      openARAssociateDrawer(AR_METRICS_STATE.selectedOperator);
    }
  }
}

function renderARMetricSummary(metricsPayload) {
  const summary = metricsPayload?.summary || {};
  const shiftHours = metricsPayload?.shiftHours || (getARShift() === "Weekend" ? 10.5 : 9.5);
  const selectedStation = getSelectedARStation();

  if (selectedStation === "ALL") {
    setText("arMetricTotalOutput", fmt(summary.totalActualOutput || 0));
    setText("arMetricAssigned", fmt(summary.assignedAssociates || 0));
    setText("arMetricCapacity", fmt(summary.totalCapacity || 0));
    setText("arMetricPace", `${fmt(summary.outputVsCapacityPercent || 0)}%`);
    setText("arMetricGap", `Gap: ${fmt(summary.gapToCapacity || 0)}`);
    setText("arMetricTopStation", summary.topStation || "--");
    setText("arMetricTopStationText", `${fmt(summary.topStationOutput || 0)} output today`);
    setText("arMetricPeakHour", summary.peakHour || "--");
    setText("arMetricPeakHourText", `${fmt(summary.peakHourValue || 0)} scans during peak hour`);
    setText("arMetricStatus", `${normalizeARText(metricsPayload?.shiftType, getARShift())} · ${shiftHours} hrs`);
    return;
  }

  const stationRow = getARStationMetricRows(metricsPayload).find(row => isARStationMatch(row.role || row.station, selectedStation)) || {};
  const stationOperators = buildAROperatorRollups(metricsPayload, selectedStation);
  const peak = calculateStationPeakFromOperators(stationOperators);
  const topOperator = stationOperators[0] || {};

  const actual = toNumber(stationRow.actualOutput) || stationOperators.reduce((sum, row) => sum + toNumber(row.total), 0);
  const assigned = toNumber(stationRow.assignedCount);
  const capacity = toNumber(stationRow.capacity);
  const pace = selectedStation === "Oven" ? 0 : toNumber(stationRow.pacePercent) || (capacity > 0 ? Math.round((actual / capacity) * 100) : 0);
  const gap = toNumber(stationRow.gap) || (capacity ? actual - capacity : 0);

  setText("arMetricTotalOutput", fmt(actual));
  setText("arMetricAssigned", fmt(assigned));
  setText("arMetricCapacity", selectedStation === "Oven" ? "N/A" : fmt(capacity));
  setText("arMetricPace", selectedStation === "Oven" ? "Process" : `${fmt(pace)}%`);
  setText("arMetricGap", selectedStation === "Oven" ? "1 hr cooldown / No JPH" : `Gap: ${fmt(gap)}`);
  setText("arMetricTopStation", selectedStation);
  setText("arMetricTopStationText", topOperator.name ? `Top: ${topOperator.name} · ${fmt(topOperator.total)}` : "No operator output found");
  setText("arMetricPeakHour", peak.hour || "--");
  setText("arMetricPeakHourText", `${fmt(peak.value)} scans during selected station peak`);
  setText("arMetricStatus", `${selectedStation} view · ${normalizeARText(metricsPayload?.shiftType, getARShift())} · ${shiftHours} hrs`);
}
function renderARStationMetrics(metricsPayload) {
  const box = $("arStationMetricGrid");
  if (!box) return;

  const selectedStation = getSelectedARStation();
  const rows = getARStationMetricRows(metricsPayload).filter(row => isARStationMatch(row.role || row.station, selectedStation));

  if (!rows.length) {
    box.innerHTML = `<div class="ar-empty-state">No AR capacity metrics returned yet.</div>`;
    return;
  }

  const ordered = rows.slice().sort((a, b) => AR_ROLE_ORDER.indexOf(a.role) - AR_ROLE_ORDER.indexOf(b.role));

  box.innerHTML = ordered.map(row => {
    const role = normalizeARText(row.role || row.station, "AR");
    const actual = toNumber(row.actualOutput);
    const capacity = toNumber(row.capacity);
    const assigned = toNumber(row.assignedCount);
    const pace = toNumber(row.pacePercent);
    const gap = toNumber(row.gap);
    const status = normalizeARText(row.status, "NO_ASSIGNED_CAPACITY");
    const note = normalizeARText(row.note, "Operator JPH capacity");
    const top = row.topOperator || {};
    const topName = normalizeARText(top.name || top.operatorName, "No operator");
    const topTotal = toNumber(top.total || top.Total);
    const width = role === "Oven" ? 100 : Math.max(2, Math.min(100, pace || 0));
    const statusClass = status.includes("AHEAD") ? "good" : status.includes("BEHIND") ? "bad" : status.includes("NO_ASSIGNED") ? "warn" : "neutral";

    return `
      <article class="ar-station-metric ${statusClass}" role="button" tabindex="0" data-ar-station-card="${escapeHtml(role)}">
        <div class="ar-station-metric-head">
          <strong>${escapeHtml(role)}</strong>
          <span>${escapeHtml(status.replaceAll("_", " "))}</span>
        </div>

        <div class="ar-station-numbers">
          <div>
            <em>Actual</em>
            <b>${fmt(actual)}</b>
          </div>
          <div>
            <em>Assigned</em>
            <b>${fmt(assigned)}</b>
          </div>
          <div>
            <em>Capacity</em>
            <b>${role === "Oven" ? "N/A" : fmt(capacity)}</b>
          </div>
        </div>

        <div class="ar-capacity-bar">
          <span style="width:${width}%"></span>
        </div>

        <div class="ar-station-meta">
          <span>${role === "Oven" ? escapeHtml(note) : `Pace ${fmt(pace)}% · Gap ${fmt(gap)}`}</span>
          <span>Top: ${escapeHtml(topName)}${topTotal ? ` · ${fmt(topTotal)}` : ""}</span>
        </div>
      </article>
    `;
  }).join("");

  box.querySelectorAll("[data-ar-station-card]").forEach(card => {
    card.addEventListener("click", () => {
      AR_METRICS_STATE.selectedStation = normalizeARStationFilter(card.dataset.arStationCard || "ALL");
      AR_METRICS_STATE.selectedOperator = null;
      closeARAssociateDrawer();
      renderARMetricsPanel(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {}, AR_METRICS_STATE.rosterPayload || {});
    });

    card.addEventListener("keydown", e => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        card.click();
      }
    });
  });
}


function readOperatorNameFromActivity(row) {
  return normalizeARText(row.Operator || row.operator || row.operatorName || row.name, "");
}

function readStationFromActivity(row) {
  return normalizeARText(row.FlowStation || row.flowStation || row.AccessPoint || row.accessPoint || row.station, "AR");
}

function readHoursObject(row) {
  const raw = row.Hours || row.hours || row.hourly || row.Hourly || {};
  return raw && typeof raw === "object" ? raw : {};
}

function getARHourLabels(metricsPayload) {
  const fromPayload = Array.isArray(metricsPayload?.hours) ? metricsPayload.hours : [];
  if (fromPayload.length) return fromPayload.map(h => String(h));

  const set = new Set();
  (metricsPayload?.operatorActivity || []).forEach(row => {
    Object.keys(readHoursObject(row)).forEach(h => set.add(h));
  });

  const fallback = ["6:00 AM", "7:00 AM", "8:00 AM", "9:00 AM", "10:00 AM", "11:00 AM", "12:00 PM", "1:00 PM", "2:00 PM", "3:00 PM", "4:00 PM", "5:00 PM", "6:00 PM", "7:00 PM", "8:00 PM"];
  return set.size ? fallback.filter(h => set.has(h)).concat(Array.from(set).filter(h => !fallback.includes(h))) : fallback;
}

function getActiveARRoster(metricsPayload) {
  const roster = Array.isArray(metricsPayload?.roster) ? metricsPayload.roster : [];
  return roster.filter(row => normalizeARText(row.activeStatus || row.ActiveStatus, "Active").toLowerCase() === "active");
}

function getRosterByOperator(metricsPayload, operatorName) {
  const target = String(operatorName || "").trim().toUpperCase();
  return getActiveARRoster(metricsPayload).find(row => getRosterOperatorName(row).toUpperCase() === target) || null;
}


function getRosterByOperatorAndRole(metricsPayload, operatorName, role) {
  const targetName = String(operatorName || "").trim().toUpperCase();
  const targetRole = normalizeARStationFilter(role || "").toUpperCase();

  return getActiveARRoster(metricsPayload).find(row => {
    const nameMatch = getRosterOperatorName(row).toUpperCase() === targetName;
    const roleMatch = normalizeARStationFilter(getRosterRole(row)).toUpperCase() === targetRole;
    return nameMatch && roleMatch;
  }) || null;
}

function getARCertifiedJphForRole(metricsPayload, role) {
  const target = normalizeARStationFilter(role || "");
  if (!target || target === "ALL" || target === "Oven") return 0;

  const config = metricsPayload?.capacityConfig || {};
  const direct = config[target] || config[target.toUpperCase()] || config[target.toLowerCase()];

  if (direct && typeof direct === "object") {
    return toNumber(
      direct.certifiedJph ||
      direct.CertifiedJPH ||
      direct.defaultJph ||
      direct.DefaultJPH ||
      direct.jph ||
      direct.JPH
    );
  }

  if (typeof direct === "number" || typeof direct === "string") {
    return toNumber(direct);
  }

  const roleDefaults = {
    "AR-IN": 120,
    "Basket": 64,
    "Sectoring": 48,
    "DeRing": 120,
    "AR-OUT": 100
  };

  return toNumber(roleDefaults[target]);
}

function getARLastActiveHourFromHours(hours, hourLabels) {
  let last = "--";
  hourLabels.forEach(hour => {
    if (toNumber(hours?.[hour]) > 0) last = hour;
  });
  return last;
}

function buildAROperatorStationBreakdown(metricsPayload, operatorName) {
  const target = String(operatorName || "").trim().toUpperCase();
  const hourLabels = getARHourLabels(metricsPayload);
  const byStation = new Map();

  (Array.isArray(metricsPayload?.operatorActivity) ? metricsPayload.operatorActivity : []).forEach(row => {
    const name = readOperatorNameFromActivity(row);
    if (!name || name.toUpperCase() !== target) return;

    const station = normalizeARStationFilter(readStationFromActivity(row));
    if (!station || station === "ALL") return;

    if (!byStation.has(station)) {
      byStation.set(station, {
        station,
        total: 0,
        rows: [],
        hours: Object.fromEntries(hourLabels.map(h => [h, 0])),
        bestHour: "--",
        bestValue: 0,
        lastActiveHour: "--",
        targetJph: 0,
        pace: 0,
        capacity: 0,
        cert: "Not Assigned"
      });
    }

    const item = byStation.get(station);
    const rowHours = readHoursObject(row);
    const rowTotal = toNumber(row.Total || row.total || row.ActivityToday || row.activityToday);

    item.total += rowTotal;
    item.rows.push(row);

    hourLabels.forEach(hour => {
      item.hours[hour] = toNumber(item.hours[hour]) + toNumber(rowHours[hour]);
    });
  });

  const shiftHours = toNumber(metricsPayload?.shiftHours || (getARShift() === "Weekend" ? 10.5 : 9.5));

  Array.from(byStation.values()).forEach(item => {
    const rosterRow = getRosterByOperatorAndRole(metricsPayload, operatorName, item.station);
    const exactJph = rosterRow ? getRosterFinalJph(rosterRow) : 0;
    const defaultJph = getARCertifiedJphForRole(metricsPayload, item.station);

    item.targetJph = item.station === "Oven" ? 0 : (exactJph || defaultJph);
    item.capacity = item.targetJph > 0 ? Math.round(item.targetJph * shiftHours) : 0;
    item.pace = item.capacity > 0 ? Math.round((item.total / item.capacity) * 100) : 0;
    item.cert = rosterRow ? getRosterCert(rosterRow) : (item.targetJph ? "Live role target" : "No JPH target");

    hourLabels.forEach(hour => {
      const value = toNumber(item.hours[hour]);
      if (value > item.bestValue) {
        item.bestValue = value;
        item.bestHour = hour;
      }
    });

    item.lastActiveHour = getARLastActiveHourFromHours(item.hours, hourLabels);
  });

  return Array.from(byStation.values()).sort((a, b) => {
    const ia = AR_ROLE_ORDER.indexOf(a.station);
    const ib = AR_ROLE_ORDER.indexOf(b.station);
    return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
  });
}

// Station-independent floater map. Scans ALL operatorActivity (ignores the
// station filter) so the floater badge is correct under any view.
// Floater = operator appears in >1 distinct FlowStation in what the API sent.
// Per-station hours = count of Hours buckets > 0; output = Total. The API owns
// any noise filtering (e.g. thin rows); the frontend trusts the payload.
function buildARFloaterMap(metricsPayload) {
  const rows = Array.isArray(metricsPayload?.operatorActivity) ? metricsPayload.operatorActivity : [];
  const byName = new Map();

  rows.forEach(row => {
    const name = readOperatorNameFromActivity(row);
    if (!name || /unassigned|no operator/i.test(name)) return;

    const station = readStationFromActivity(row);
    const output = toNumber(row.Total || row.total || row.ActivityToday || row.activityToday);
    const hours = readHoursObject(row);
    const activeHours = Object.keys(hours).reduce((n, h) => n + (toNumber(hours[h]) > 0 ? 1 : 0), 0);

    const key = name.toUpperCase();
    if (!byName.has(key)) byName.set(key, { name, stationMap: new Map() });

    const map = byName.get(key).stationMap;
    const prev = map.get(station) || { station, activeHours: 0, output: 0 };
    prev.activeHours += activeHours;
    prev.output += output;
    map.set(station, prev);
  });

  const out = {};
  byName.forEach((entry, key) => {
    const stations = Array.from(entry.stationMap.values()).sort((a, b) => b.output - a.output);
    out[key] = { name: entry.name, stations, isFloater: stations.length > 1 };
  });
  return out;
}

// Chips for the stations a floater worked OTHER than the one being rendered.
function renderARFloaterChips(floaterMap, operatorName, excludeStation) {
  const f = floaterMap[String(operatorName || "").toUpperCase()];
  if (!f || !f.isFloater) return "";
  const ex = String(excludeStation || "").trim().toUpperCase();
  const others = f.stations.filter(s => String(s.station || "").trim().toUpperCase() !== ex);
  if (!others.length) return "";
  return `<div class="ar-floated-chips">${others.map(s =>
    `<span class="ar-floated-chip"><span class="st">${escapeHtml(s.station)}</span> <b>${fmt(s.activeHours)}h</b> · <b>${fmt(s.output)}</b></span>`
  ).join("")}</div>`;
}

function buildAROperatorRollups(metricsPayload, stationFilter = getSelectedARStation()) {
  const selectedStation = normalizeARStationFilter(stationFilter || "ALL");
  const rows = (Array.isArray(metricsPayload?.operatorActivity) ? metricsPayload.operatorActivity : [])
    .filter(row => isARStationMatch(readStationFromActivity(row), selectedStation));
  const hourLabels = getARHourLabels(metricsPayload);
  const byName = new Map();

  rows.forEach(row => {
    const name = readOperatorNameFromActivity(row);
    if (!name) return;

    const key = name.toUpperCase();
    const total = toNumber(row.Total || row.total || row.ActivityToday || row.activityToday);
    const station = readStationFromActivity(row);
    const bestHour = normalizeARText(row.BestHour || row.bestHour || row.besthour, "--");
    const bestValue = toNumber(row.BestHourValue || row.bestHourValue || row.bestvalue);
    const lastActiveHour = normalizeARText(row.LastActiveHour || row.lastActiveHour || row.lastactivehour, "--");
    const hours = readHoursObject(row);

    if (!byName.has(key)) {
      byName.set(key, {
        name,
        total: 0,
        stations: new Set(),
        rows: [],
        hours: Object.fromEntries(hourLabels.map(h => [h, 0])),
        bestHour: "--",
        bestValue: 0,
        lastActiveHour: "--"
      });
    }

    const item = byName.get(key);
    item.total += total;
    item.stations.add(station);
    item.rows.push(row);

    hourLabels.forEach(hour => {
      item.hours[hour] = toNumber(item.hours[hour]) + toNumber(hours[hour]);
    });

    if (bestValue > item.bestValue) {
      item.bestHour = bestHour;
      item.bestValue = bestValue;
    }

    if (lastActiveHour && lastActiveHour !== "--") {
      item.lastActiveHour = lastActiveHour;
    }
  });

  const roster = getActiveARRoster(metricsPayload);
  const rosterByName = new Map(roster.map(row => [getRosterOperatorName(row).toUpperCase(), row]));

  return Array.from(byName.values()).map(item => {
    const rosterRow = rosterByName.get(item.name.toUpperCase()) || null;
    const role = rosterRow ? getRosterRole(rosterRow) : Array.from(item.stations)[0] || "AR";
    const cert = rosterRow ? getRosterCert(rosterRow) : "Not Assigned";
    const finalJph = rosterRow ? getRosterFinalJph(rosterRow) : 0;
    const shiftHours = toNumber(metricsPayload?.shiftHours || (getARShift() === "Weekend" ? 10.5 : 9.5));
    const capacity = role === "Oven" ? 0 : Math.round(finalJph * shiftHours);
    const pace = capacity > 0 ? Math.round((item.total / capacity) * 100) : 0;

    return {
      ...item,
      role,
      cert,
      finalJph,
      capacity,
      pace,
      assigned: !!rosterRow,
      rosterRow
    };
  }).sort((a, b) => b.total - a.total);
}

function openARAssociateDrawer(operatorName) {
  const metricsPayload = AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {};
  const target = String(operatorName || "").trim().toUpperCase();
  const item = buildAROperatorRollups(metricsPayload).find(row => row.name.toUpperCase() === target);
  const stationBreakdown = buildAROperatorStationBreakdown(metricsPayload, operatorName);

  if (!item && !stationBreakdown.length) {
    showToast("No hourly activity found for that associate.");
    return;
  }

  const activeItem = item || {
    name: operatorName,
    total: stationBreakdown.reduce((sum, row) => sum + toNumber(row.total), 0),
    stations: new Set(stationBreakdown.map(row => row.station)),
    bestHour: "--",
    bestValue: 0,
    lastActiveHour: "--",
    role: "Multiple areas",
    cert: "Live activity",
    finalJph: 0,
    capacity: 0,
    pace: 0
  };

  AR_METRICS_STATE.selectedOperator = activeItem.name;

  const drawer = $("arAssociateDrawer");
  const backdrop = $("arAssociateDrawerBackdrop");
  if (!drawer) return;

  const shiftHours = toNumber(metricsPayload?.shiftHours || (getARShift() === "Weekend" ? 10.5 : 9.5));
  const totalOutput = stationBreakdown.length
    ? stationBreakdown.reduce((sum, row) => sum + toNumber(row.total), 0)
    : toNumber(activeItem.total);
  const totalCapacity = stationBreakdown.reduce((sum, row) => sum + toNumber(row.capacity), 0);
  const totalPace = totalCapacity > 0 ? Math.round((totalOutput / totalCapacity) * 100) : 0;

  const bestStation = stationBreakdown.reduce((best, row) => {
    return toNumber(row.bestValue) > toNumber(best.bestValue || 0) ? row : best;
  }, {});

  const lastActive = stationBreakdown.reduce((last, row) => {
    return row.lastActiveHour && row.lastActiveHour !== "--" ? row.lastActiveHour : last;
  }, activeItem.lastActiveHour || "--");

  const roleText = stationBreakdown.length > 1
    ? `${stationBreakdown.length} areas`
    : (stationBreakdown[0]?.station || activeItem.role || "AR");

  const capText = totalCapacity > 0 ? `${fmt(totalCapacity)} total cap` : "By area / No JPH";
  const paceText = totalCapacity > 0 ? `${fmt(totalPace)}%` : "By area";

  setText("arDrawerName", activeItem.name);
  setText("arDrawerMeta", `${roleText} · ${capText} · ${shiftHours} hr shift`);
  setText("arDrawerRole", roleText);
  setText("arDrawerJph", stationBreakdown.length > 1 ? "By area" : (stationBreakdown[0]?.targetJph ? fmt(stationBreakdown[0].targetJph) : "--"));
  setText("arDrawerTotal", fmt(totalOutput));
  setText("arDrawerPace", paceText);
  setText("arDrawerBestHour", bestStation.station ? `${bestStation.station} · ${bestStation.bestHour} (${fmt(bestStation.bestValue)})` : `${activeItem.bestHour} (${fmt(activeItem.bestValue)})`);
  setText("arDrawerLastHour", lastActive || "--");
  setText("arDrawerStations", stationBreakdown.map(row => `${row.station}: ${fmt(row.total)}`).join(" · ") || Array.from(activeItem.stations || []).join(" · ") || "--");

  const hoursBox = $("arDrawerHours");
  if (hoursBox) {
    const labels = getARHourLabels(metricsPayload);

    if (!stationBreakdown.length) {
      const maxHour = Math.max(...Object.values(activeItem.hours || {}).map(toNumber), 1);
      const targetJph = getAROperatorHourlyTarget(activeItem);
      const hasTarget = targetJph > 0 && activeItem.role !== "Oven";

      hoursBox.innerHTML = labels.map(hour => {
        const value = toNumber(activeItem.hours?.[hour]);
        const width = Math.max(value > 0 ? 4 : 0, Math.round((value / maxHour) * 100));
        const pctVal = hasTarget ? Math.round((value / targetJph) * 100) : 0;
        const perfClass = value === 0 && !hasTarget ? "no-target" : getARPerformanceClass(pctVal, hasTarget);
        return `
          <div class="ar-hour-line ${value > 0 ? "active" : ""} ${perfClass}">
            <span>${escapeHtml(hour)}</span>
            <div><i style="width:${width}%"></i></div>
            <strong>${fmt(value)}${hasTarget ? ` · ${pctVal}%` : ""}</strong>
          </div>
        `;
      }).join("");
    } else {
      hoursBox.innerHTML = stationBreakdown.map(section => {
        const maxHour = Math.max(...Object.values(section.hours || {}).map(toNumber), 1);
        const hasTarget = section.targetJph > 0 && section.station !== "Oven";
        const sectionPerf = getARPerformanceClass(section.pace, hasTarget);
        const capacityText = section.station === "Oven"
          ? "Process constraint · 1 hr cooldown · No JPH"
          : hasTarget
            ? `${fmt(section.targetJph)} JPH · ${fmt(section.capacity)} cap · ${fmt(section.pace)}% pace`
            : "No JPH target";

        return `
          <article class="ar-drawer-station-block perf-${sectionPerf}">
            <div class="ar-drawer-station-head">
              <div>
                <strong>${escapeHtml(section.station)}</strong>
                <span>${escapeHtml(section.cert)} · ${escapeHtml(capacityText)}</span>
              </div>
              <div>
                <b>${fmt(section.total)}</b>
                <em>output</em>
              </div>
            </div>

            <div class="ar-drawer-station-meta">
              <span>Best: ${escapeHtml(section.bestHour)} (${fmt(section.bestValue)})</span>
              <span>Last active: ${escapeHtml(section.lastActiveHour || "--")}</span>
            </div>

            <div class="ar-drawer-station-hours">
              ${labels.map(hour => {
                const value = toNumber(section.hours?.[hour]);
                const width = Math.max(value > 0 ? 4 : 0, Math.round((value / maxHour) * 100));
                const pctVal = hasTarget ? Math.round((value / section.targetJph) * 100) : 0;
                const perfClass = value === 0 && !hasTarget ? "no-target" : getARPerformanceClass(pctVal, hasTarget);
                return `
                  <div class="ar-hour-line ${value > 0 ? "active" : ""} ${perfClass}">
                    <span>${escapeHtml(hour)}</span>
                    <div><i style="width:${width}%"></i></div>
                    <strong>${fmt(value)}${hasTarget ? ` · ${pctVal}%` : ""}</strong>
                  </div>
                `;
              }).join("")}
            </div>
          </article>
        `;
      }).join("");
    }
  }

  drawer.classList.add("open");
  drawer.setAttribute("aria-hidden", "false");
  if (backdrop) backdrop.classList.add("open");

  document.querySelectorAll(".ar-operator-row.active, .ar-roster-item.active").forEach(el => el.classList.remove("active"));
  document.querySelectorAll(`[data-operator-detail="${CSS.escape(activeItem.name)}"], [data-operator="${CSS.escape(activeItem.name)}"]`).forEach(el => el.classList.add("active"));
}

function closeARAssociateDrawer() {
  const drawer = $("arAssociateDrawer");
  const backdrop = $("arAssociateDrawerBackdrop");
  if (drawer) {
    drawer.classList.remove("open");
    drawer.setAttribute("aria-hidden", "true");
  }
  if (backdrop) backdrop.classList.remove("open");
}


function getARPerformanceClass(percent, hasTarget = true) {
  if (!hasTarget) return "no-target";
  const p = toNumber(percent);
  if (p >= 100) return "good";
  if (p >= 90) return "amber";
  return "bad";
}

function getAROperatorHourlyTarget(item) {
  const target = toNumber(item?.finalJph);
  return target > 0 ? target : 0;
}

function renderAROperatorHourStrip(item, metricsPayload) {
  const labels = getARHourLabels(metricsPayload);
  const target = getAROperatorHourlyTarget(item);
  const hasTarget = target > 0 && item.role !== "Oven";

  return `
    <div class="ar-hourly-strip" aria-label="Hourly operator output">
      ${labels.map(hour => {
        const value = toNumber(item.hours?.[hour]);
        const pctVal = hasTarget ? Math.round((value / target) * 100) : 0;
        const cls = value === 0 && !hasTarget ? "no-target" : getARPerformanceClass(pctVal, hasTarget);
        const shortHour = String(hour).replace(":00 ", "").replace(" AM", "A").replace(" PM", "P");
        return `
          <div class="ar-hour-cell ${cls}" title="${escapeHtml(hour)} · ${fmt(value)}${hasTarget ? ` / ${fmt(target)} (${pctVal}%)` : ""}">
            <span>${escapeHtml(shortHour)}</span>
            <strong>${fmt(value)}</strong>
            <em>${hasTarget ? `${fmt(pctVal)}%` : "No Target"}</em>
          </div>
        `;
      }).join("")}
    </div>
  `;
}

function renderAROperatorPerformance(metricsPayload) {
  const box = $("arOperatorPerformanceList");
  if (!box) return;

  const list = buildAROperatorRollups(metricsPayload);
  const floaterMap = buildARFloaterMap(metricsPayload);
  if (!list.length) {
    box.innerHTML = `<div class="ar-empty-state">No AR operator activity found.</div>`;
    return;
  }

  box.innerHTML = list.map((item, index) => {
    const target = getAROperatorHourlyTarget(item);
    const hasTarget = target > 0 && item.role !== "Oven";
    const currentPct = hasTarget ? Math.round((item.bestValue / target) * 100) : 0;
    const perfClass = getARPerformanceClass(currentPct, hasTarget);
    const assignClass = item.assigned ? "assigned" : "not-assigned";
    const capacityText = item.role === "Oven"
      ? "Process constraint · 1 hr cooldown · No JPH capacity"
      : item.capacity
        ? `Current JPH ${fmt(currentPct)}% · ${fmt(item.bestValue)} / ${fmt(target)}`
        : "No assigned capacity";

    return `
      <button class="ar-operator-row ${assignClass} perf-${perfClass}" type="button" data-operator-detail="${escapeHtml(item.name)}">
        <div class="ar-operator-topline">
          <div class="ar-operator-rank">${index + 1}</div>
          <div class="ar-operator-main">
            <strong>${escapeHtml(item.name)}${floaterMap[item.name.toUpperCase()]?.isFloater ? ` <span class="ar-floater-badge">⇄ Floater</span>` : ""}</strong>
            <span>${escapeHtml(item.role)} · ${escapeHtml(item.cert)} · Peak ${escapeHtml(item.bestHour)} (${fmt(item.bestValue)})</span>
          </div>
          <div class="ar-operator-score">
            <strong>${fmt(item.total)}</strong>
            <span>${hasTarget ? "JPH" : "Output"}</span>
          </div>
        </div>

        <div class="ar-operator-jph-pill ${perfClass}">
          ${escapeHtml(capacityText)}
        </div>

        ${renderARFloaterChips(floaterMap, item.name, item.role)}
        ${renderAROperatorHourStrip(item, metricsPayload)}
      </button>
    `;
  }).join("");

  box.querySelectorAll("[data-operator-detail]").forEach(btn => {
    btn.addEventListener("click", () => openARAssociateDrawer(btn.dataset.operatorDetail));
  });
}

function renderARRosterControls(metricsPayload, rosterPayload) {
  const ownerInput = $("arMetricOwner");
  const shiftSelect = $("arMetricShift");

  if (ownerInput && !ownerInput.value) ownerInput.value = AR_METRICS_STATE.owner || "BLOPEZ";
  if (shiftSelect) shiftSelect.value = AR_METRICS_STATE.shift || "Weekday";

  const canSeeRosterAdmin = applyARLMSOnlyVisibility();
  if (!canSeeRosterAdmin) {
    return;
  }

  const operatorSelect = $("arAssignOperator");
  if (operatorSelect) {
    const selected = operatorSelect.value;
    const options = getOperatorMasterRows(rosterPayload, metricsPayload);

    operatorSelect.innerHTML = `<option value="">Select associate</option>` + options.map(row => `
      <option value="${escapeHtml(row.operatorName)}" data-role="${escapeHtml(row.defaultRole)}">${escapeHtml(row.operatorName)}</option>
    `).join("");

    if (selected) operatorSelect.value = selected;
  }

  renderARRosterList(metricsPayload?.roster || rosterPayload?.roster || []);
}


function getARRosterDisplayKey(row) {
  return getRosterOperatorName(row).toUpperCase();
}

function compactARRosterForDisplay(roster) {
  const map = new Map();
  (Array.isArray(roster) ? roster : []).forEach(row => {
    const status = normalizeARText(row.activeStatus || row.ActiveStatus, "Active").toLowerCase();
    if (status !== "active") return;

    const key = getARRosterDisplayKey(row);
    const current = map.get(key);
    if (!current) {
      map.set(key, row);
      return;
    }

    const currentTime = new Date(current.updatedAt || current.UpdatedAt || 0).getTime() || 0;
    const rowTime = new Date(row.updatedAt || row.UpdatedAt || 0).getTime() || 0;
    if (rowTime >= currentTime) map.set(key, row);
  });

  return Array.from(map.values()).sort((a, b) => {
    const nameA = getRosterOperatorName(a).toUpperCase();
    const nameB = getRosterOperatorName(b).toUpperCase();
    return nameA.localeCompare(nameB);
  });
}

// Champ-select lane definitions. Order + colors map to the AR station tokens.
const AR_ROSTER_LANES = [
  { key: "AR-IN",     tag: "Intake",    color: "var(--cyan)",   glow: "rgba(0,217,255,.10)",  constraint: false },
  { key: "Basket",    tag: "Sorting",   color: "var(--purple)", glow: "rgba(155,92,255,.10)", constraint: false },
  { key: "Oven",      tag: "Cure",      color: "var(--red)",    glow: "rgba(255,64,64,.10)",  constraint: true  },
  { key: "Sectoring", tag: "Routing",   color: "var(--green)",  glow: "rgba(86,227,109,.10)", constraint: false },
  { key: "DeRing",    tag: "Finishing", color: "var(--orange)", glow: "rgba(255,153,0,.10)",  constraint: false },
  { key: "AR-OUT",    tag: "Dispatch",  color: "var(--amber)",  glow: "rgba(255,191,63,.10)", constraint: false },
];

function arRosterInitials(name) {
  return String(name || "")
    .split(/\s+/).filter(Boolean).map(w => w[0]).join("").slice(0, 2).toUpperCase() || "??";
}

function renderARRosterCard(row, lane, floaterMap) {
  floaterMap = floaterMap || {};
  const name = getRosterOperatorName(row);
  const role = getRosterRole(row);
  const cert = getRosterCert(row);
  const week = toNumber(row.trainingWeek || row.TrainingWeek) || 1;
  const custom = toNumber(row.individualJph || row.IndividualJPH);
  const finalJph = getRosterFinalJph(row);
  const training = cert === "Training";
  const isFloater = !!(floaterMap[name.toUpperCase()] && floaterMap[name.toUpperCase()].isFloater);

  const pips = training
    ? `<span class="ar-champ-pips">${[1, 2, 3, 4, 5]
        .map(i => `<span class="ar-champ-pip ${i <= week ? "on" : ""}"></span>`).join("")}</span>`
    : "";

  const stat = lane.constraint
    ? `<span class="ar-champ-jph process">PROCESS · NO JPH</span>`
    : `<span class="ar-champ-jph">${fmt(finalJph)}<small>JPH</small></span>${custom ? `<span class="ar-champ-custom">CUSTOM</span>` : ""}`;

  return `
    <div class="ar-champ${isFloater ? " is-floater" : ""}" role="button" tabindex="0" data-operator="${escapeHtml(name)}">
      <div class="ar-champ-top">
        <div class="ar-champ-av">${escapeHtml(arRosterInitials(name))}</div>
        <div class="ar-champ-id">
          <div class="nm">${escapeHtml(name)}</div>
          <div class="rl">${escapeHtml(role)}${isFloater ? " · home" : ""}</div>
        </div>
      </div>
      <div class="ar-champ-badges">
        <span class="ar-champ-cert ${training ? "training" : "certified"}">${escapeHtml(cert)}${training ? ` W${week}` : ""}${pips}</span>
        ${isFloater ? `<span class="ar-floater-badge">⇄ Floater</span>` : ""}
      </div>
      <div class="ar-champ-stat">${stat}</div>
      ${renderARFloaterChips(floaterMap, name, role)}
      <div class="ar-champ-actions">
        <button type="button" class="ar-mini-btn" data-edit-ar="${escapeHtml(name)}" data-edit-role="${escapeHtml(role)}">Edit</button>
        <button type="button" class="ar-mini-btn danger" data-clear-ar="${escapeHtml(name)}" data-clear-role="${escapeHtml(role)}">Clear</button>
      </div>
    </div>
  `;
}

function renderARRosterList(roster) {
  const box = $("arRosterList");
  const count = $("arRosterCount");
  const active = compactARRosterForDisplay(roster);
  // Floater map comes from live activity on the current metrics payload, so a
  // roster card badges automatically when that person gets scanned in 2+ areas.
  const floaterMap = buildARFloaterMap(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {});

  if (count) count.textContent = `${active.length} assigned`;

  if (!box) return;

  // Group active roster rows by role for the lane view.
  const byRole = new Map(AR_ROSTER_LANES.map(l => [l.key, []]));
  const overflowRole = [];
  active.forEach(row => {
    const role = getRosterRole(row);
    if (byRole.has(role)) byRole.get(role).push(row);
    else overflowRole.push(row); // unknown role still gets shown
  });

  box.innerHTML = AR_ROSTER_LANES.map(lane => {
    const rows = byRole.get(lane.key) || [];
    const laneJph = lane.constraint
      ? "—"
      : fmt(rows.reduce((sum, r) => sum + (getRosterFinalJph(r) || 0), 0));

    const cards = rows.map(r => renderARRosterCard(r, lane, floaterMap)).join("");
    const constraintChip = lane.constraint
      ? `<span class="ar-lane-constraint">⚠ Constraint · 1h Cooldown</span>` : "";

    return `
      <div class="ar-lane${lane.constraint ? " constraint" : ""}" style="--lane:${lane.color};--lane-glow:${lane.glow}">
        <div class="ar-lane-head">
          <div class="ar-lane-title">
            <span class="ar-lane-dot"></span>
            <span class="ar-lane-name">${escapeHtml(lane.key)}</span>
            <span class="ar-lane-tag">${escapeHtml(lane.tag)}</span>
            ${constraintChip}
          </div>
          <div class="ar-lane-stats">
            <div class="ar-lane-stat"><div class="v">${rows.length}</div><div class="l">Assigned</div></div>
            <div class="ar-lane-stat"><div class="v">${laneJph}</div><div class="l">Lane JPH</div></div>
          </div>
        </div>
        <div class="ar-lane-slots">
          ${cards}
          <div class="ar-slot-empty" data-assign-role="${escapeHtml(lane.key)}">+ Assign to ${escapeHtml(lane.key)}</div>
        </div>
      </div>
    `;
  }).join("");

  // Operator card / drawer open (click + keyboard) — unchanged behavior.
  box.querySelectorAll("[data-operator]").forEach(row => {
    row.addEventListener("click", () => openARAssociateDrawer(row.dataset.operator));
    row.addEventListener("keydown", event => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openARAssociateDrawer(row.dataset.operator);
      }
    });
  });

  // Clear — same handler as before.
  box.querySelectorAll("[data-clear-ar]").forEach(btn => {
    btn.addEventListener("click", event => {
      event.stopPropagation();
      handleClearARAssignment(btn.dataset.clearAr, btn.dataset.clearRole);
    });
  });

  // Edit — same handler as before.
  box.querySelectorAll("[data-edit-ar]").forEach(btn => {
    btn.addEventListener("click", event => {
      event.stopPropagation();
      fillARRosterForm(btn.dataset.editAr, active, btn.dataset.editRole);
    });
  });

  // Empty slot → preset the console role and focus the associate picker.
  box.querySelectorAll("[data-assign-role]").forEach(slot => {
    slot.addEventListener("click", () => {
      const roleSelect = $("arAssignRole");
      const opSelect = $("arAssignOperator");
      if (roleSelect) roleSelect.value = slot.dataset.assignRole;
      if (typeof syncARAssignWeekField === "function") syncARAssignWeekField();
      if (opSelect) opSelect.focus();
    });
  });
}

// Hide Training Week unless Status = Training (matches the card model).
function syncARAssignWeekField() {
  const cert = $("arAssignCert");
  const field = $("arAssignWeekField");
  if (!field) return;
  field.style.display = (cert && cert.value === "Training") ? "" : "none";
}

function fillARRosterForm(operatorName, roster, role) {
  const operatorKey = String(operatorName || "").toUpperCase();
  const roleKey = String(role || "").toUpperCase();
  const row = roster.find(item =>
    getRosterOperatorName(item).toUpperCase() === operatorKey &&
    (!roleKey || getRosterRole(item).toUpperCase() === roleKey)
  );
  if (!row) return;

  const operatorSelect = $("arAssignOperator");
  const roleSelect = $("arAssignRole");
  const certSelect = $("arAssignCert");
  const weekSelect = $("arAssignWeek");
  const customInput = $("arAssignCustomJph");

  if (operatorSelect) operatorSelect.value = getRosterOperatorName(row);
  if (roleSelect) roleSelect.value = getRosterRole(row);
  if (certSelect) certSelect.value = getRosterCert(row);
  if (weekSelect) weekSelect.value = String(toNumber(row.trainingWeek || row.TrainingWeek) || 1);
  if (customInput) customInput.value = toNumber(row.individualJph || row.IndividualJPH) || "";
}

async function handleSaveARAssignment() {
  if (!isARLMSUser()) {
    showToast("AR assignment control is LMS only.");
    applyARLMSOnlyVisibility();
    return;
  }

  if (AR_METRICS_STATE.isSaving) return;

  const operatorName = $("arAssignOperator")?.value || "";
  if (!operatorName) {
    showToast("Select an AR associate first.");
    return;
  }

  const payload = {
    operatorName,
    role: $("arAssignRole")?.value || "AR-IN",
    certificationStatus: $("arAssignCert")?.value || "Certified",
    trainingWeek: $("arAssignWeek")?.value || 1,
    individualJPH: $("arAssignCustomJph")?.value || 0
  };

  try {
    AR_METRICS_STATE.isSaving = true;
    const btn = $("arAssignBtn");
    if (btn) btn.textContent = "Saving...";

    await saveARRosterAssignment(payload);
    showToast("AR assignment saved.");
    await loadARData();
  } catch (err) {
    console.error("AR assignment save failed:", err);
    showToast("AR assignment save failed.");
  } finally {
    AR_METRICS_STATE.isSaving = false;
    const btn = $("arAssignBtn");
    if (btn) btn.textContent = "Assign / Update";
  }
}

async function handleClearARAssignment(operatorName, role) {
  if (!isARLMSUser()) {
    showToast("AR assignment clear is LMS only.");
    applyARLMSOnlyVisibility();
    return;
  }

  if (!operatorName || AR_METRICS_STATE.isSaving) return;

  try {
    AR_METRICS_STATE.isSaving = true;
    await clearARRosterAssignment(operatorName, role);

    const operatorKey = String(operatorName || "").toUpperCase();
    const roleKey = String(role || "").toUpperCase();
    const current = AR_METRICS_STATE.metricsPayload?.roster || [];
    renderARRosterList(current.filter(row => {
      const sameOperator = getRosterOperatorName(row).toUpperCase() === operatorKey;
      const sameRole = getRosterRole(row).toUpperCase() === roleKey;
      return !(sameOperator && sameRole);
    }));

    showToast("AR assignment cleared.");
    await loadARData();
  } catch (err) {
    console.error("AR assignment clear failed:", err);
    showToast("AR assignment clear failed.");
  } finally {
    AR_METRICS_STATE.isSaving = false;
  }
}

function bindARMetricEvents() {
  applyARLMSOnlyVisibility();

  const refresh = $("arMetricRefresh");
  const owner = $("arMetricOwner");
  const shift = $("arMetricShift");
  const assign = $("arAssignBtn");
  const operatorSelect = $("arAssignOperator");
  const drawerClose = $("arAssociateDrawerClose");
  const drawerBackdrop = $("arAssociateDrawerBackdrop");

  if (refresh && !refresh.dataset.bound) {
    refresh.addEventListener("click", loadARData);
    refresh.dataset.bound = "1";
  }

  if (owner && !owner.dataset.bound) {
    owner.addEventListener("change", loadARData);
    owner.dataset.bound = "1";
  }

  if (shift && !shift.dataset.bound) {
    shift.addEventListener("change", loadARData);
    shift.dataset.bound = "1";
  }

  if (assign && !assign.dataset.bound) {
    assign.addEventListener("click", handleSaveARAssignment);
    assign.dataset.bound = "1";
  }

  if (drawerClose && !drawerClose.dataset.bound) {
    drawerClose.addEventListener("click", closeARAssociateDrawer);
    drawerClose.dataset.bound = "1";
  }

  if (drawerBackdrop && !drawerBackdrop.dataset.bound) {
    drawerBackdrop.addEventListener("click", closeARAssociateDrawer);
    drawerBackdrop.dataset.bound = "1";
  }

  if (operatorSelect && !operatorSelect.dataset.bound) {
    operatorSelect.addEventListener("change", () => {
      const option = operatorSelect.options[operatorSelect.selectedIndex];
      const role = option?.dataset?.role;
      if (role && $("arAssignRole")) $("arAssignRole").value = role;
    });
    operatorSelect.dataset.bound = "1";
  }

  const certSelect = $("arAssignCert");
  if (certSelect && !certSelect.dataset.bound) {
    certSelect.addEventListener("change", syncARAssignWeekField);
    certSelect.dataset.bound = "1";
  }
  syncARAssignWeekField();

  if (!bindARMetricEvents._escapeBound) {
    document.addEventListener("keydown", event => {
      if (event.key === "Escape") closeARAssociateDrawer();
    });
    bindARMetricEvents._escapeBound = true;
  }
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
        <span>Current largest WIP station with ${fmt(largest.value)} WIP jobs.</span>
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
    Math.max(values.chamberTotalCapacity, AR_CAPACITY_RULES.CHAMBER_JOB)
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
    title: `${fmt(values.fullChambers)} full · ${fmt(values.partialChambers)} partial sectoring chambers`,
    text: `Sectoring utilization is ${chamberPct}% by job chamber load. Lens equivalent: ${fmt(values.sectoringLensLoad)} lenses.`,
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
        renderARMetricsPanel(LAST_CAPACITY_PAYLOAD, AR_METRICS_STATE.rosterPayload);
      }

      if (section === "lmsSetup") {
        if (!applyARLMSOnlyVisibility()) {
          showToast("LMS Setup is restricted to LMS users.");
          const overviewBtn = document.querySelector('.nav-item[data-section="overview"]');
          if (overviewBtn) overviewBtn.click();
          return;
        }
        renderARRosterControls(LAST_CAPACITY_PAYLOAD, AR_METRICS_STATE.rosterPayload);
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


function bindARLoadoutToggle() {
  const section = $("arLoadoutSection");
  const toggle = $("arLoadoutToggle");
  if (!section || !toggle || toggle.dataset.bound) return;

  toggle.addEventListener("click", () => {
    const isOpen = section.classList.toggle("open");
    section.classList.toggle("closed", !isOpen);
    toggle.setAttribute("aria-expanded", isOpen ? "true" : "false");
  });

  toggle.dataset.bound = "1";
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
  bindARMetricEvents();
  bindARLoadoutToggle();
}

window.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  updateClock();
  loadARData();

  setInterval(updateClock, 1000 * 30);
  setInterval(loadARData, REFRESH_MS);
});

/* =====================================================
   AR PRODUCTIVITY COMMAND CENTER — ADVANCED UI RENDERERS
   These override the summary tab renderers with the advanced layout.
===================================================== */

function arGetInitials(name) {
  const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "AR";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function arStatusFromPace(role, pace, statusText) {
  const roleText = normalizeARStationFilter(role || "");
  const status = String(statusText || "").toUpperCase();
  const p = toNumber(pace);

  if (roleText === "Oven" || status.includes("CONSTRAINT") || status.includes("NO_JPH")) {
    return { key: "constraint", label: "Constraint" };
  }
  if (p >= 100 || status.includes("AHEAD")) return { key: "good", label: "On Track" };
  if (p >= 75) return { key: "watch", label: "Watch" };
  return { key: "bad", label: "Behind" };
}

function arOperatorStatus(item, percent, hasTarget) {
  if (!hasTarget || item.role === "Oven") return { key: "constraint", label: "Constraint" };
  const p = toNumber(percent);
  if (p >= 100) return { key: "good", label: "Elite" };
  if (p >= 90) return { key: "watch", label: "On Target" };
  if (p >= 75) return { key: "watch", label: "Watch" };
  return { key: "bad", label: "Needs Support" };
}

function arShortHourLabel(hour) {
  return String(hour || "")
    .replace(":00 ", "")
    .replace(" AM", "A")
    .replace(" PM", "P");
}

function renderARMetricsPanel(metricsPayload, rosterPayload) {
  renderARMetricSummary(metricsPayload);
  renderARAdvancedProcessFlow(metricsPayload);
  renderARProductivityOverview(metricsPayload);
  renderAROperatorLeaderboard(metricsPayload);
  renderARHourlyTrend(metricsPayload);
  renderARCapacityOverview(metricsPayload);
  renderAROperatorPerformance(metricsPayload);
  renderARCommandTip(metricsPayload);
  renderARRosterControls(metricsPayload, rosterPayload);

  if (AR_METRICS_STATE.selectedOperator) {
    const stillExists = buildAROperatorRollups(metricsPayload || {}).some(item => item.name.toUpperCase() === AR_METRICS_STATE.selectedOperator.toUpperCase());
    if (stillExists && $("arAssociateDrawer")?.classList.contains("open")) {
      openARAssociateDrawer(AR_METRICS_STATE.selectedOperator);
    }
  }
}

function renderARAdvancedProcessFlow(metricsPayload) {
  const box = $("arAdvancedProcessFlow");
  if (!box) return;

  const rows = getARStationMetricRows(metricsPayload)
    .slice()
    .sort((a, b) => AR_ROLE_ORDER.indexOf(normalizeARStationFilter(a.role || a.station)) - AR_ROLE_ORDER.indexOf(normalizeARStationFilter(b.role || b.station)));

  if (!rows.length) {
    box.innerHTML = `<div class="ar-empty-state">No station flow metrics returned yet.</div>`;
    return;
  }

  const selectedStation = getSelectedARStation();
  const allBtn = $("arProcessAllBtn");
  if (allBtn) {
    allBtn.classList.toggle("active", selectedStation === "ALL");
    allBtn.onclick = () => {
      AR_METRICS_STATE.selectedStation = "ALL";
      AR_METRICS_STATE.selectedOperator = null;
      closeARAssociateDrawer();
      renderARMetricsPanel(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {}, AR_METRICS_STATE.rosterPayload || {});
    };
  }

  box.innerHTML = rows.map(row => {
    const role = normalizeARStationFilter(row.role || row.station || "AR");
    const activeClass = role.toUpperCase() === selectedStation.toUpperCase() ? "active" : "";
    const actual = toNumber(row.actualOutput);
    const capacity = toNumber(row.capacity);
    const assigned = toNumber(row.assignedCount);
    const pace = role === "Oven" ? 0 : toNumber(row.pacePercent) || (capacity ? Math.round((actual / capacity) * 100) : 0);
    const gap = toNumber(row.gap) || (capacity ? actual - capacity : 0);
    const status = arStatusFromPace(role, pace, row.status || row.note);
    const width = role === "Oven" ? 100 : Math.max(2, Math.min(100, pace || 0));

    return `
      <article class="ar-flow-node ${status.key} ${activeClass}" role="button" tabindex="0" data-ar-station-card="${escapeHtml(role)}">
        <div class="ar-flow-node-head">
          <strong>${escapeHtml(role)}</strong>
          <span>${escapeHtml(status.label)}</span>
        </div>
        <div class="ar-flow-node-main">
          <div><span>Output</span><strong>${fmt(actual)}</strong></div>
          <div><span>Capacity</span><strong>${role === "Oven" ? "N/A" : fmt(capacity)}</strong></div>
        </div>
        <div class="ar-flow-bar"><i style="width:${width}%"></i></div>
        <div class="ar-flow-node-foot">
          <div><span>Pace</span>${role === "Oven" ? "Process" : `${fmt(pace)}%`}</div>
          <div><span>Gap</span>${role === "Oven" ? "No JPH" : fmt(gap)}</div>
          <div><span>Assigned</span>${fmt(assigned)}</div>
        </div>
      </article>
    `;
  }).join("");

  box.querySelectorAll("[data-ar-station-card]").forEach(card => {
    card.addEventListener("click", () => {
      AR_METRICS_STATE.selectedStation = normalizeARStationFilter(card.dataset.arStationCard || "ALL");
      AR_METRICS_STATE.selectedOperator = null;
      closeARAssociateDrawer();
      renderARMetricsPanel(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {}, AR_METRICS_STATE.rosterPayload || {});
    });
    card.addEventListener("keydown", e => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        card.click();
      }
    });
  });
}

function renderARProductivityOverview(metricsPayload) {
  const list = buildAROperatorRollups(metricsPayload);
  const scored = list.map(item => {
    const target = getAROperatorHourlyTarget(item);
    const hasTarget = target > 0 && item.role !== "Oven";
    const percent = hasTarget ? Math.round((toNumber(item.bestValue) / target) * 100) : 0;
    return { item, target, hasTarget, percent };
  });

  const targetRows = scored.filter(row => row.hasTarget);
  const avg = targetRows.length
    ? Math.round(targetRows.reduce((sum, row) => sum + Math.min(row.percent, 150), 0) / targetRows.length)
    : 0;

  const full = targetRows.filter(row => row.percent >= 100).length;
  const near = targetRows.filter(row => row.percent >= 90 && row.percent < 100).length;
  const below = targetRows.filter(row => row.percent >= 50 && row.percent < 90).length;
  const low = targetRows.filter(row => row.percent < 50).length;

  const ring = $("arProductivityRing");
  if (ring) ring.style.setProperty("--score", Math.max(0, Math.min(100, avg)));

  setText("arProductivityScore", `${fmt(avg)}%`);
  setText("arTierFull", fmt(full));
  setText("arTierNear", fmt(near));
  setText("arTierBelow", fmt(below));
  setText("arTierLow", fmt(low));
  setText("arProductivitySubText", `${fmt(targetRows.length)} operators with active JPH targets`);

  const best = scored.filter(row => row.hasTarget).sort((a, b) => b.percent - a.percent)[0] || scored[0];
  if (best) {
    setText("arBestPerformer", best.item.name);
    setText("arBestPerformerText", `${fmt(best.percent)}% of target · ${escapeHtml(best.item.role)}`);
  } else {
    setText("arBestPerformer", "--");
    setText("arBestPerformerText", "No operator activity found");
  }

  const pressure = getARStationMetricRows(metricsPayload)
    .filter(row => normalizeARStationFilter(row.role || row.station) !== "Oven")
    .sort((a, b) => toNumber(a.gap) - toNumber(b.gap))[0];

  if (pressure) {
    setText("arPressureStation", normalizeARStationFilter(pressure.role || pressure.station));
    setText("arPressureStationText", `Gap ${fmt(pressure.gap)} · Pace ${fmt(pressure.pacePercent)}%`);
  } else {
    setText("arPressureStation", "--");
    setText("arPressureStationText", "No station capacity data");
  }
}

function renderAROperatorLeaderboard(metricsPayload) {
  const box = $("arOperatorLeaderboard");
  if (!box) return;

  const list = buildAROperatorRollups(metricsPayload).slice(0, 5);
  if (!list.length) {
    box.innerHTML = `<div class="ar-empty-state">No operator activity found.</div>`;
    return;
  }

  box.innerHTML = list.map((item, index) => {
    const target = getAROperatorHourlyTarget(item);
    const hasTarget = target > 0 && item.role !== "Oven";
    const percent = hasTarget ? Math.round((toNumber(item.bestValue) / target) * 100) : 0;
    const status = arOperatorStatus(item, percent, hasTarget);

    return `
      <button class="ar-leader-row" type="button" data-operator-detail="${escapeHtml(item.name)}">
        <div class="ar-leader-rank">${index + 1}</div>
        <div class="ar-leader-main">
          <strong>${escapeHtml(item.name)}</strong>
          <span>${escapeHtml(item.role)} · ${escapeHtml(item.cert)}</span>
        </div>
        <div class="ar-leader-cell"><span>Current JPH</span><strong>${fmt(item.total)}</strong></div>
        <div class="ar-leader-cell"><span>Target %</span><strong>${hasTarget ? `${fmt(percent)}%` : "N/A"}</strong></div>
        <div class="ar-leader-cell optional"><span>Peak Hour</span><strong>${escapeHtml(item.bestHour)}</strong></div>
        <span class="ar-status-badge ${status.key}">${escapeHtml(status.label)}</span>
      </button>
    `;
  }).join("");

  box.querySelectorAll("[data-operator-detail]").forEach(btn => {
    btn.addEventListener("click", () => openARAssociateDrawer(btn.dataset.operatorDetail));
  });

  renderARPerformanceInsights(metricsPayload);
}

function renderARPerformanceInsights(metricsPayload) {
  const box = $("arPerformanceInsights");
  if (!box) return;

  const summary = metricsPayload?.summary || {};
  const stationRows = getARStationMetricRows(metricsPayload);
  const list = buildAROperatorRollups(metricsPayload);
  const selected = getSelectedARStation();

  const worstStation = stationRows
    .filter(row => normalizeARStationFilter(row.role || row.station) !== "Oven")
    .sort((a, b) => toNumber(a.gap) - toNumber(b.gap))[0];

  const bestOperator = list[0];
  const peakHour = summary.peakHour || calculateStationPeakFromOperators(list).hour || "--";

  const insights = [
    {
      title: "Station Focus",
      text: worstStation
        ? `${normalizeARStationFilter(worstStation.role || worstStation.station)} has the largest gap (${fmt(worstStation.gap)}).`
        : "No station gap is available yet."
    },
    {
      title: "Labor Signal",
      text: bestOperator
        ? `${bestOperator.name} is leading output at ${fmt(bestOperator.total)} total scans.`
        : "No associate output is available yet."
    },
    {
      title: "Peak Demand",
      text: `${peakHour} is the current peak hour${summary.peakHourValue ? ` with ${fmt(summary.peakHourValue)} scans.` : "."}`
    }
  ];

  if (selected !== "ALL") {
    insights[0].text = `Focused station view: ${selected}. Use All AR to compare the full department.`;
  }

  box.innerHTML = insights.map(item => `
    <div class="ar-insight-item">
      <strong>${escapeHtml(item.title)}</strong>
      <span>${escapeHtml(item.text)}</span>
    </div>
  `).join("");
}

function renderARHourlyTrend(metricsPayload) {
  const box = $("arHourlyTrend");
  if (!box) return;

  const labels = getARHourLabels(metricsPayload);
  const operators = buildAROperatorRollups(metricsPayload);
  const totals = labels.map(hour => operators.reduce((sum, item) => sum + toNumber(item.hours?.[hour]), 0));
  const max = Math.max(...totals, 1);

  box.innerHTML = labels.map((hour, index) => {
    const value = totals[index];
    const height = Math.max(value > 0 ? 8 : 2, Math.round((value / max) * 160));
    return `
      <div class="ar-trend-bar" title="${escapeHtml(hour)} · ${fmt(value)} output">
        <i style="height:${height}px"></i>
        <strong>${fmt(value)}</strong>
        <span>${escapeHtml(arShortHourLabel(hour))}</span>
      </div>
    `;
  }).join("");
}

function renderARCapacityOverview(metricsPayload) {
  const summary = metricsPayload?.summary || {};
  const totalOutput = toNumber(summary.totalActualOutput);
  const totalCapacity = toNumber(summary.totalCapacity);
  const pace = toNumber(summary.outputVsCapacityPercent) || (totalCapacity ? Math.round((totalOutput / totalCapacity) * 100) : 0);
  const gap = toNumber(summary.gapToCapacity) || (totalCapacity ? totalOutput - totalCapacity : 0);

  const gauge = $("arCapacityGauge");
  if (gauge) gauge.style.setProperty("--score", Math.max(0, Math.min(100, pace)));

  setText("arCapacityGaugeValue", `${fmt(pace)}%`);
  setText("arCapacityTotalOutput", fmt(totalOutput));
  setText("arCapacityTotalCap", fmt(totalCapacity));
  setText("arCapacityGapValue", fmt(gap));
  setText("arCapacityPeakHour", summary.peakHour || "--");

  const statusText = pace >= 100 ? "Capacity target met" : pace >= 75 ? "Watch capacity gap" : "Behind daily capacity";
  setText("arCapacityStatusText", statusText);
}

function renderARStationMetrics(metricsPayload) {
  const box = $("arStationMetricGrid");
  if (!box) return;

  const selectedStation = getSelectedARStation();
  const rows = getARStationMetricRows(metricsPayload).filter(row => isARStationMatch(row.role || row.station, selectedStation));

  if (!rows.length) {
    box.innerHTML = `<div class="ar-empty-state">No AR capacity metrics returned yet.</div>`;
    return;
  }

  const ordered = rows.slice().sort((a, b) => AR_ROLE_ORDER.indexOf(normalizeARStationFilter(a.role || a.station)) - AR_ROLE_ORDER.indexOf(normalizeARStationFilter(b.role || b.station)));

  box.innerHTML = ordered.map(row => {
    const role = normalizeARStationFilter(row.role || row.station || "AR");
    const actual = toNumber(row.actualOutput);
    const capacity = toNumber(row.capacity);
    const assigned = toNumber(row.assignedCount);
    const pace = role === "Oven" ? 0 : toNumber(row.pacePercent) || (capacity ? Math.round((actual / capacity) * 100) : 0);
    const gap = toNumber(row.gap) || (capacity ? actual - capacity : 0);
    const top = row.topOperator || {};
    const topName = normalizeARText(top.name || top.operatorName, "No operator");
    const topTotal = toNumber(top.total || top.Total);
    const status = arStatusFromPace(role, pace, row.status || row.note);
    const width = role === "Oven" ? 100 : Math.max(2, Math.min(100, pace || 0));
    const action = role === "Oven" ? "Resolve constraint" : pace >= 100 ? "Maintain pace" : "Add support";

    return `
      <article class="ar-station-metric ${status.key}" role="button" tabindex="0" data-ar-station-card="${escapeHtml(role)}">
        <div class="ar-station-metric-head">
          <strong>${escapeHtml(role)}</strong>
          <span>${escapeHtml(status.label)}</span>
        </div>
        <div class="ar-station-numbers">
          <div><em>Output</em><b>${fmt(actual)}</b></div>
          <div><em>Assigned</em><b>${fmt(assigned)}</b></div>
          <div><em>Capacity</em><b>${role === "Oven" ? "N/A" : fmt(capacity)}</b></div>
        </div>
        <div class="ar-capacity-bar"><span style="width:${width}%"></span></div>
        <div class="ar-station-meta">
          <span>${role === "Oven" ? "Process constraint / No JPH capacity" : `Pace ${fmt(pace)}% · Gap ${fmt(gap)}`}</span>
          <span>Top: ${escapeHtml(topName)}${topTotal ? ` · ${fmt(topTotal)}` : ""}</span>
        </div>
        <div class="ar-station-action">Action: ${escapeHtml(action)}</div>
      </article>
    `;
  }).join("");

  box.querySelectorAll("[data-ar-station-card]").forEach(card => {
    card.addEventListener("click", () => {
      AR_METRICS_STATE.selectedStation = normalizeARStationFilter(card.dataset.arStationCard || "ALL");
      AR_METRICS_STATE.selectedOperator = null;
      closeARAssociateDrawer();
      renderARMetricsPanel(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {}, AR_METRICS_STATE.rosterPayload || {});
    });
    card.addEventListener("keydown", e => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        card.click();
      }
    });
  });
}

function renderAROperatorPerformance(metricsPayload) {
  const box = $("arOperatorPerformanceList");
  if (!box) return;

  const list = buildAROperatorRollups(metricsPayload);
  const floaterMap = buildARFloaterMap(metricsPayload);

  if (!list.length) {
    box.innerHTML = `<div class="ar-empty-state">No AR operator activity found.</div>`;
    return;
  }

  box.innerHTML = list.map((item, index) => {
    const target = getAROperatorHourlyTarget(item);
    const hasTarget = target > 0 && item.role !== "Oven";
    const currentPct = hasTarget ? Math.round((item.bestValue / target) * 100) : 0;
    const perfClass = getARPerformanceClass(currentPct, hasTarget);
    const status = arOperatorStatus(item, currentPct, hasTarget);
    const initials = arGetInitials(item.name);
    const capacityText = item.role === "Oven"
      ? "Process constraint · 1 hr cooldown · No JPH capacity"
      : hasTarget
        ? `${fmt(currentPct)}% target · ${fmt(item.bestValue)} / ${fmt(target)} best hour`
        : "No assigned capacity";

    return `
      <button class="ar-operator-loadout-card perf-${perfClass}" type="button" data-operator-detail="${escapeHtml(item.name)}">
        <div class="ar-loadout-top">
          <div class="ar-loadout-avatar">${escapeHtml(initials)}</div>
          <div class="ar-loadout-name">
            <strong>${escapeHtml(item.name)}${floaterMap[item.name.toUpperCase()]?.isFloater ? ` <span class="ar-floater-badge">⇄ Floater</span>` : ""}</strong>
            <span>#${index + 1} · ${escapeHtml(item.role)} · ${escapeHtml(item.cert)}</span>
          </div>
          <div class="ar-loadout-score">
            <strong>${fmt(item.total)}</strong>
            <span>${hasTarget ? "Output" : "Activity"}</span>
          </div>
        </div>

        <div class="ar-loadout-meta">
          <div><span>Target</span><strong>${hasTarget ? `${fmt(currentPct)}%` : "N/A"}</strong></div>
          <div><span>Peak Hour</span><strong>${escapeHtml(item.bestHour)}</strong></div>
          <div><span>Status</span><strong>${escapeHtml(status.label)}</strong></div>
        </div>

        <div class="ar-operator-jph-pill ${perfClass}">${escapeHtml(capacityText)}</div>
        ${renderARFloaterChips(floaterMap, item.name, item.role)}
        ${renderAROperatorHourStrip(item, metricsPayload)}
      </button>
    `;
  }).join("");

  box.querySelectorAll("[data-operator-detail]").forEach(btn => {
    btn.addEventListener("click", () => openARAssociateDrawer(btn.dataset.operatorDetail));
  });
}

function renderARCommandTip(metricsPayload) {
  const box = $("arCommandTip");
  if (!box) return;

  const summary = metricsPayload?.summary || {};
  const stationRows = getARStationMetricRows(metricsPayload);
  const worstStation = stationRows
    .filter(row => normalizeARStationFilter(row.role || row.station) !== "Oven")
    .sort((a, b) => toNumber(a.gap) - toNumber(b.gap))[0];

  let tip = "Monitor station pace and keep floaters ready for the largest capacity gap.";

  if (worstStation) {
    const station = normalizeARStationFilter(worstStation.role || worstStation.station);
    const gap = toNumber(worstStation.gap);
    const pace = toNumber(worstStation.pacePercent);
    tip = `Focus support on ${station}. Current pace is ${fmt(pace)}% with a gap of ${fmt(gap)} against capacity.`;
  }

  if (summary.outputVsCapacityPercent >= 100) {
    tip = "Daily capacity target is currently met. Keep the same station coverage and watch the next peak hour.";
  }

  box.innerHTML = `<strong>Command Tip</strong><span>${escapeHtml(tip)}</span>`;
}


/* ─────────────────────────────────────────────────────
   AR FLOATING SUPPORT OVERRIDES
   Adds AR-IN, AR-OUT Floater, Utility, and General Floater logic.
   LMS assignment = body/home role. Activity = production credit.
───────────────────────────────────────────────────── */
(function initARFloatingSupportOverride(){
  const floatingRoles = ["AR-IN, AR-OUT Floater", "Utility", "Floater"];
  if (Array.isArray(AR_ROLE_ORDER)) {
    floatingRoles.forEach(role => {
      if (!AR_ROLE_ORDER.some(r => String(r).toUpperCase() === role.toUpperCase())) AR_ROLE_ORDER.push(role);
    });
  }

  window.AR_FLOATING_ROLES = floatingRoles;
})();

function isARFloatingRole(role) {
  const r = String(role || "").trim().toUpperCase();
  return r === "AR + AR-OUT FLOATER" || r === "UTILITY" || r === "FLOATER" || r === "GENERAL FLOATER";
}

function isARProductionStation(role) {
  const r = normalizeARStationFilter(role);
  return ["AR-IN", "Basket", "Oven", "Sectoring", "DeRing", "AR-OUT"].some(x => x.toUpperCase() === r.toUpperCase());
}

function normalizeARStationFilter(value) {
  const clean = normalizeARText(value, "ALL");
  if (!clean || clean.toUpperCase() === "ALL") return "ALL";

  const upper = clean.toUpperCase().replace(/\s+/g, " ").trim();
  if (upper === "AR + AR-OUT FLOATER" || upper === "AR/AR-OUT FLOATER" || upper === "AR-IN, AR-OUT FLOATER") return "AR-IN, AR-OUT Floater";
  if (upper === "UTILITY" || upper.includes("UTILITY")) return "Utility";
  if (upper === "FLOATER" || upper === "GENERAL FLOATER") return "Floater";

  if (upper.includes("AR-IN") || upper.includes("AR IN")) return "AR-IN";
  if (upper.includes("BASKET")) return "Basket";
  if (upper.includes("OVEN")) return "Oven";
  if (upper.includes("SECTOR")) return "Sectoring";
  if (upper.includes("DERING") || upper.includes("DE RING")) return "DeRing";
  if (upper.includes("AR-OUT") || upper.includes("AR OUT")) return "AR-OUT";

  const match = AR_ROLE_ORDER.find(role => role.toUpperCase() === upper);
  return match || clean;
}

function getCurrentARRosterRows(metricsPayload) {
  const direct = Array.isArray(metricsPayload?.roster) ? metricsPayload.roster : [];
  const stateMetrics = Array.isArray(AR_METRICS_STATE?.metricsPayload?.roster) ? AR_METRICS_STATE.metricsPayload.roster : [];
  const stateRoster = Array.isArray(AR_METRICS_STATE?.rosterPayload?.roster) ? AR_METRICS_STATE.rosterPayload.roster : [];
  const rows = direct.length ? direct : (stateMetrics.length ? stateMetrics : stateRoster);
  return rows.filter(row => normalizeARText(row.activeStatus || row.ActiveStatus, "Active").toLowerCase() === "active");
}

function getARActivitySplitForOperator(metricsPayload, operatorName) {
  const target = String(operatorName || "").trim().toUpperCase();
  const rows = Array.isArray(metricsPayload?.operatorActivity) ? metricsPayload.operatorActivity : [];
  const stationTotals = {};
  rows.forEach(row => {
    const name = readOperatorNameFromActivity(row).toUpperCase();
    if (!target || name !== target) return;
    const station = normalizeARStationFilter(readStationFromActivity(row));
    if (!isARProductionStation(station)) return;
    const output = toNumber(row.Total || row.total || row.ActivityToday || row.activityToday);
    stationTotals[station] = (stationTotals[station] || 0) + output;
  });

  const total = Object.values(stationTotals).reduce((sum, v) => sum + toNumber(v), 0);
  const split = {};
  if (total > 0) {
    Object.keys(stationTotals).forEach(station => {
      split[station] = stationTotals[station] / total;
    });
  }
  return split;
}

function getARFloatingBodyImpact(metricsPayload) {
  const impact = {
    "AR-IN": 0,
    "Basket": 0,
    "Oven": 0,
    "Sectoring": 0,
    "DeRing": 0,
    "AR-OUT": 0,
    utilityBodies: 0,
    floaterBodies: 0,
    arArOutBodies: 0,
    totalBodies: 0
  };

  getCurrentARRosterRows(metricsPayload).forEach(row => {
    const role = normalizeARStationFilter(getRosterRole(row));
    const name = getRosterOperatorName(row);
    impact.totalBodies += 1;

    if (role === "AR-IN, AR-OUT Floater") {
      impact["AR-IN"] += 0.5;
      impact["AR-OUT"] += 0.5;
      impact.arArOutBodies += 1;
      return;
    }

    if (role === "Utility") {
      impact.utilityBodies += 1;
      return;
    }

    if (role === "Floater") {
      impact.floaterBodies += 1;
      const split = getARActivitySplitForOperator(metricsPayload, name);
      Object.keys(split).forEach(station => {
        impact[station] = (impact[station] || 0) + split[station];
      });
    }
  });

  return impact;
}

function fmtBodyCount(value) {
  const n = toNumber(value);
  if (Math.abs(n - Math.round(n)) < 0.01) return fmt(Math.round(n));
  return n.toFixed(1).replace(/\.0$/, "");
}

function adjustARAssignedCountForFloaters(metricsPayload, role, baseAssigned) {
  const station = normalizeARStationFilter(role);
  if (!isARProductionStation(station)) return toNumber(baseAssigned);
  const impact = getARFloatingBodyImpact(metricsPayload);
  return toNumber(baseAssigned) + toNumber(impact[station]);
}

function getARRoleTargetFromStationMetrics(metricsPayload, role) {
  const station = normalizeARStationFilter(role);
  if (!isARProductionStation(station) || station === "Oven") return 0;
  const row = getARStationMetricRows(metricsPayload).find(r => normalizeARStationFilter(r.role || r.station) === station);
  if (!row) return 0;

  const direct = toNumber(row.targetJph || row.TargetJPH || row.jph || row.JPH || row.laneJph || row.LaneJPH || row.capacityPerAssociate || row.CapacityPerAssociate);
  if (direct > 0) return direct;

  const assigned = toNumber(row.assignedCount);
  const cap = toNumber(row.capacity);
  return assigned > 0 && cap > 0 ? Math.round(cap / assigned) : 0;
}

function getAROperatorHourlyTarget(item) {
  const role = normalizeARStationFilter(item?.role || "");
  const custom = toNumber(item?.finalJph);
  if (custom > 0) return custom;
  if (role === "Oven" || role === "Utility") return 0;

  const metricsPayload = AR_METRICS_STATE?.metricsPayload || LAST_CAPACITY_PAYLOAD || {};
  if (role === "AR-IN, AR-OUT Floater") {
    const arInTarget = getARRoleTargetFromStationMetrics(metricsPayload, "AR-IN");
    const arOutTarget = getARRoleTargetFromStationMetrics(metricsPayload, "AR-OUT");
    return Math.round(((arInTarget || 0) + (arOutTarget || 0)) / ([arInTarget, arOutTarget].filter(Boolean).length || 1));
  }

  if (role === "Floater") {
    const rows = Array.isArray(item?.rows) ? item.rows : [];
    const stationTotals = {};
    rows.forEach(row => {
      const station = normalizeARStationFilter(readStationFromActivity(row));
      if (!isARProductionStation(station) || station === "Oven") return;
      stationTotals[station] = (stationTotals[station] || 0) + toNumber(row.Total || row.total || row.ActivityToday || row.activityToday);
    });
    const dominant = Object.entries(stationTotals).sort((a, b) => b[1] - a[1])[0]?.[0];
    return dominant ? getARRoleTargetFromStationMetrics(metricsPayload, dominant) : 0;
  }

  return getARRoleTargetFromStationMetrics(metricsPayload, role);
}

function renderARAdvancedProcessFlow(metricsPayload) {
  const box = $("arAdvancedProcessFlow");
  if (!box) return;

  const rows = getARStationMetricRows(metricsPayload)
    .filter(row => isARProductionStation(row.role || row.station))
    .slice()
    .sort((a, b) => AR_ROLE_ORDER.indexOf(normalizeARStationFilter(a.role || a.station)) - AR_ROLE_ORDER.indexOf(normalizeARStationFilter(b.role || b.station)));

  if (!rows.length) {
    box.innerHTML = `<div class="ar-empty-state">No station flow metrics returned yet.</div>`;
    return;
  }

  const selectedStation = getSelectedARStation();
  const allBtn = $("arProcessAllBtn");
  if (allBtn) {
    allBtn.classList.toggle("active", selectedStation === "ALL");
    allBtn.onclick = () => {
      AR_METRICS_STATE.selectedStation = "ALL";
      AR_METRICS_STATE.selectedOperator = null;
      closeARAssociateDrawer();
      renderARMetricsPanel(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {}, AR_METRICS_STATE.rosterPayload || {});
    };
  }

  box.innerHTML = rows.map(row => {
    const role = normalizeARStationFilter(row.role || row.station || "AR");
    const activeClass = role.toUpperCase() === selectedStation.toUpperCase() ? "active" : "";
    const actual = toNumber(row.actualOutput);
    const capacity = toNumber(row.capacity);
    const assigned = adjustARAssignedCountForFloaters(metricsPayload, role, row.assignedCount);
    const pace = role === "Oven" ? 0 : toNumber(row.pacePercent) || (capacity ? Math.round((actual / capacity) * 100) : 0);
    const gap = toNumber(row.gap) || (capacity ? actual - capacity : 0);
    const status = arStatusFromPace(role, pace, row.status || row.note);
    const width = role === "Oven" ? 100 : Math.max(2, Math.min(100, pace || 0));

    return `
      <article class="ar-flow-node ${status.key} ${activeClass}" role="button" tabindex="0" data-ar-station-card="${escapeHtml(role)}">
        <div class="ar-flow-node-head">
          <strong>${escapeHtml(role)}</strong>
          <span>${escapeHtml(status.label)}</span>
        </div>
        <div class="ar-flow-node-main">
          <div><span>Output</span><strong>${fmt(actual)}</strong></div>
          <div><span>Capacity</span><strong>${role === "Oven" ? "N/A" : fmt(capacity)}</strong></div>
        </div>
        <div class="ar-flow-bar"><i style="width:${width}%"></i></div>
        <div class="ar-flow-node-foot">
          <div><span>Pace</span>${role === "Oven" ? "Process" : `${fmt(pace)}%`}</div>
          <div><span>Gap</span>${role === "Oven" ? "No JPH" : fmt(gap)}</div>
          <div><span>Assigned</span>${fmtBodyCount(assigned)}</div>
        </div>
      </article>
    `;
  }).join("");

  box.querySelectorAll("[data-ar-station-card]").forEach(card => {
    card.addEventListener("click", () => {
      AR_METRICS_STATE.selectedStation = normalizeARStationFilter(card.dataset.arStationCard || "ALL");
      AR_METRICS_STATE.selectedOperator = null;
      closeARAssociateDrawer();
      renderARMetricsPanel(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {}, AR_METRICS_STATE.rosterPayload || {});
    });
    card.addEventListener("keydown", e => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        card.click();
      }
    });
  });
}

const AR_ROSTER_LANES_FLOATING = [
  { key: "AR-IN", tag: "Intake", color: "var(--cyan)", glow: "rgba(0,217,255,.10)", constraint: false },
  { key: "Basket", tag: "Sorting", color: "var(--purple)", glow: "rgba(155,92,255,.10)", constraint: false },
  { key: "Oven", tag: "Cure", color: "var(--red)", glow: "rgba(255,64,64,.10)", constraint: true },
  { key: "Sectoring", tag: "Routing", color: "var(--green)", glow: "rgba(86,227,109,.10)", constraint: false },
  { key: "DeRing", tag: "Finishing", color: "var(--orange)", glow: "rgba(255,153,0,.10)", constraint: false },
  { key: "AR-OUT", tag: "Dispatch", color: "var(--amber)", glow: "rgba(255,191,63,.10)", constraint: false },
  { key: "AR-IN, AR-OUT Floater", tag: "0.5 AR-IN / 0.5 AR-OUT", color: "#57d7ff", glow: "rgba(87,215,255,.10)", constraint: false, floating: true },
  { key: "Utility", tag: "Support Body", color: "#b7c0d8", glow: "rgba(183,192,216,.10)", constraint: false, support: true },
  { key: "Floater", tag: "Activity-Based Split", color: "#ffbf3f", glow: "rgba(255,191,63,.10)", constraint: false, floating: true }
];

function getARRosterLaneMeta(role) {
  const normalized = normalizeARStationFilter(role);
  return AR_ROSTER_LANES_FLOATING.find(l => l.key === normalized) || { key: normalized, tag: "Custom", color: "var(--orange)", glow: "rgba(255,153,0,.10)", constraint: false };
}

function renderARRosterCard(row, lane, floaterMap) {
  floaterMap = floaterMap || {};
  const name = getRosterOperatorName(row);
  const role = normalizeARStationFilter(getRosterRole(row));
  const cert = getRosterCert(row);
  const week = toNumber(row.trainingWeek || row.TrainingWeek) || 1;
  const custom = toNumber(row.individualJph || row.IndividualJPH);
  const finalJph = getRosterFinalJph(row);
  const training = cert === "Training";
  const floatingRole = isARFloatingRole(role);
  const autoFloater = !!(floaterMap[name.toUpperCase()] && floaterMap[name.toUpperCase()].isFloater);

  const pips = training
    ? `<span class="ar-champ-pips">${[1, 2, 3, 4, 5]
        .map(i => `<span class="ar-champ-pip ${i <= week ? "on" : ""}"></span>`).join("")}</span>`
    : "";

  let stat;
  if (lane.constraint || role === "Oven") {
    stat = `<span class="ar-champ-jph process">PROCESS · NO JPH</span>`;
  } else if (role === "Utility") {
    stat = `<span class="ar-champ-jph process">BODY COUNT · SUPPORT</span>`;
  } else if (role === "Floater") {
    stat = `<span class="ar-champ-jph">AUTO<small>JPH</small></span><span class="ar-champ-custom">ACTIVITY SPLIT</span>`;
  } else if (role === "AR-IN, AR-OUT Floater") {
    stat = `<span class="ar-champ-jph">0.5 / 0.5<small>BODY</small></span>${custom ? `<span class="ar-champ-custom">CUSTOM</span>` : ""}`;
  } else {
    stat = `<span class="ar-champ-jph">${fmt(finalJph)}<small>JPH</small></span>${custom ? `<span class="ar-champ-custom">CUSTOM</span>` : ""}`;
  }

  return `
    <div class="ar-champ${floatingRole || autoFloater ? " is-floater" : ""}" role="button" tabindex="0" data-operator="${escapeHtml(name)}">
      <div class="ar-champ-top">
        <div class="ar-champ-av">${escapeHtml(arRosterInitials(name))}</div>
        <div class="ar-champ-id">
          <div class="nm">${escapeHtml(name)}</div>
          <div class="rl">${escapeHtml(role)}${floatingRole ? " · home" : ""}</div>
        </div>
      </div>
      <div class="ar-champ-badges">
        <span class="ar-champ-cert ${training ? "training" : "certified"}">${escapeHtml(cert)}${training ? ` W${week}` : ""}${pips}</span>
        ${floatingRole ? `<span class="ar-floater-badge">⇄ ${role === "Utility" ? "Utility" : "Floater"}</span>` : ""}
        ${autoFloater && !floatingRole ? `<span class="ar-floater-badge">⇄ Multi-area</span>` : ""}
      </div>
      <div class="ar-champ-stat">${stat}</div>
      ${renderARFloaterChips(floaterMap, name, role)}
      <div class="ar-champ-actions">
        <button type="button" class="ar-mini-btn" data-edit-ar="${escapeHtml(name)}" data-edit-role="${escapeHtml(role)}">Edit</button>
        <button type="button" class="ar-mini-btn danger" data-clear-ar="${escapeHtml(name)}" data-clear-role="${escapeHtml(role)}">Clear</button>
      </div>
    </div>
  `;
}

function renderARRosterList(roster) {
  const box = $("arRosterList");
  const count = $("arRosterCount");
  const active = compactARRosterForDisplay(roster);
  const floaterMap = buildARFloaterMap(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {});
  const impact = getARFloatingBodyImpact(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {});

  if (count) {
    const support = impact.utilityBodies + impact.floaterBodies + impact.arArOutBodies;
    count.textContent = `${active.length} assigned · ${fmtBodyCount(support)} floating/support`;
  }

  if (!box) return;

  const byRole = new Map(AR_ROSTER_LANES_FLOATING.map(l => [l.key, []]));
  const overflowRole = [];
  active.forEach(row => {
    const role = normalizeARStationFilter(getRosterRole(row));
    if (byRole.has(role)) byRole.get(role).push(row);
    else overflowRole.push(row);
  });

  const laneHtml = AR_ROSTER_LANES_FLOATING.map(lane => {
    const rows = byRole.get(lane.key) || [];
    let laneJph = "—";
    let laneAssigned = rows.length;

    if (lane.key === "AR-IN, AR-OUT Floater") {
      laneJph = `${fmtBodyCount(rows.length * 0.5)} / ${fmtBodyCount(rows.length * 0.5)}`;
    } else if (lane.key === "Utility") {
      laneJph = "Support";
    } else if (lane.key === "Floater") {
      laneJph = "Auto";
    } else if (!lane.constraint) {
      laneAssigned = rows.length + toNumber(impact[lane.key]);
      laneJph = fmt(rows.reduce((sum, r) => sum + (getRosterFinalJph(r) || 0), 0));
    }

    const cards = rows.map(r => renderARRosterCard(r, lane, floaterMap)).join("");
    const constraintChip = lane.constraint ? `<span class="ar-lane-constraint">⚠ Constraint · 1h Cooldown</span>` : "";
    const supportChip = lane.floating || lane.support ? `<span class="ar-lane-constraint support">Support Role</span>` : "";

    return `
      <div class="ar-lane${lane.constraint ? " constraint" : ""}${lane.floating || lane.support ? " floating-lane" : ""}" style="--lane:${lane.color};--lane-glow:${lane.glow}">
        <div class="ar-lane-head">
          <div class="ar-lane-title">
            <span class="ar-lane-dot"></span>
            <span class="ar-lane-name">${escapeHtml(lane.key)}</span>
            <span class="ar-lane-tag">${escapeHtml(lane.tag)}</span>
            ${constraintChip}${supportChip}
          </div>
          <div class="ar-lane-stats">
            <div class="ar-lane-stat"><div class="v">${fmtBodyCount(laneAssigned)}</div><div class="l">Body</div></div>
            <div class="ar-lane-stat"><div class="v">${laneJph}</div><div class="l">${lane.floating ? "Split" : lane.support ? "Mode" : "Lane JPH"}</div></div>
          </div>
        </div>
        <div class="ar-lane-slots">
          ${cards || `<button type="button" class="ar-empty-slot" data-assign-role="${escapeHtml(lane.key)}">+ Assign to ${escapeHtml(lane.key)}</button>`}
        </div>
      </div>
    `;
  }).join("");

  const overflowHtml = overflowRole.length ? `
    <div class="ar-lane floating-lane" style="--lane:var(--orange);--lane-glow:rgba(255,153,0,.10)">
      <div class="ar-lane-head"><div class="ar-lane-title"><span class="ar-lane-dot"></span><span class="ar-lane-name">Other Roles</span><span class="ar-lane-tag">Custom</span></div></div>
      <div class="ar-lane-slots">${overflowRole.map(r => renderARRosterCard(r, getARRosterLaneMeta(getRosterRole(r)), floaterMap)).join("")}</div>
    </div>` : "";

  box.innerHTML = `
    <div class="ar-floating-summary-strip">
      <div><strong>${fmtBodyCount(impact.totalBodies)}</strong><span>Total bodies</span></div>
      <div><strong>${fmtBodyCount(impact.arArOutBodies)}</strong><span>AR + AR-OUT floaters</span></div>
      <div><strong>${fmtBodyCount(impact.utilityBodies)}</strong><span>Utility support</span></div>
      <div><strong>${fmtBodyCount(impact.floaterBodies)}</strong><span>Activity floaters</span></div>
      <div><strong>${fmtBodyCount(impact["AR-IN"])}</strong><span>Extra AR-IN body</span></div>
      <div><strong>${fmtBodyCount(impact["AR-OUT"])}</strong><span>Extra AR-OUT body</span></div>
    </div>
    ${laneHtml}${overflowHtml}
  `;

  box.querySelectorAll("[data-assign-role]").forEach(slot => {
    slot.addEventListener("click", () => {
      const roleSelect = $("arAssignRole");
      const opSelect = $("arAssignOperator");
      if (roleSelect) roleSelect.value = slot.dataset.assignRole;
      if (typeof syncARAssignWeekField === "function") syncARAssignWeekField();
      if (opSelect) opSelect.focus();
    });
  });

  box.querySelectorAll("[data-edit-ar]").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();
      const rosterRows = getCurrentARRosterRows(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {});
      fillARRosterForm(btn.dataset.editAr, rosterRows, btn.dataset.editRole);
      if (typeof syncARAssignWeekField === "function") syncARAssignWeekField();
    });
  });

  box.querySelectorAll("[data-clear-ar]").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();
      handleClearARAssignment(btn.dataset.clearAr, btn.dataset.clearRole);
    });
  });
}

function renderARStationMetrics(metricsPayload) {
  const box = $("arStationMetricGrid");
  if (!box) return;

  const selectedStation = getSelectedARStation();
  const rows = getARStationMetricRows(metricsPayload)
    .filter(row => isARProductionStation(row.role || row.station))
    .filter(row => isARStationMatch(row.role || row.station, selectedStation));

  if (!rows.length) {
    box.innerHTML = `<div class="ar-empty-state">No AR capacity metrics returned yet.</div>`;
    return;
  }

  const ordered = rows.slice().sort((a, b) => AR_ROLE_ORDER.indexOf(normalizeARStationFilter(a.role || a.station)) - AR_ROLE_ORDER.indexOf(normalizeARStationFilter(b.role || b.station)));

  box.innerHTML = ordered.map(row => {
    const role = normalizeARStationFilter(row.role || row.station || "AR");
    const actual = toNumber(row.actualOutput);
    const capacity = toNumber(row.capacity);
    const assigned = adjustARAssignedCountForFloaters(metricsPayload, role, row.assignedCount);
    const pace = role === "Oven" ? 0 : toNumber(row.pacePercent) || (capacity ? Math.round((actual / capacity) * 100) : 0);
    const gap = toNumber(row.gap) || (capacity ? actual - capacity : 0);
    const top = row.topOperator || {};
    const topName = normalizeARText(top.name || top.operatorName, "No operator");
    const topTotal = toNumber(top.total || top.Total);
    const status = arStatusFromPace(role, pace, row.status || row.note);
    const width = role === "Oven" ? 100 : Math.max(2, Math.min(100, pace || 0));
    const action = role === "Oven" ? "Resolve constraint" : pace >= 100 ? "Maintain pace" : "Add support";

    return `
      <article class="ar-station-metric ${status.key}" role="button" tabindex="0" data-ar-station-card="${escapeHtml(role)}">
        <div class="ar-station-metric-head">
          <strong>${escapeHtml(role)}</strong>
          <span>${escapeHtml(status.label)}</span>
        </div>
        <div class="ar-station-numbers">
          <div><em>Output</em><b>${fmt(actual)}</b></div>
          <div><em>Assigned</em><b>${fmtBodyCount(assigned)}</b></div>
          <div><em>Capacity</em><b>${role === "Oven" ? "N/A" : fmt(capacity)}</b></div>
        </div>
        <div class="ar-capacity-bar"><span style="width:${width}%"></span></div>
        <div class="ar-station-meta">
          <span>${role === "Oven" ? "Process constraint / No JPH capacity" : `Pace ${fmt(pace)}% · Gap ${fmt(gap)}`}</span>
          <span>Top: ${escapeHtml(topName)}${topTotal ? ` · ${fmt(topTotal)}` : ""}</span>
        </div>
        <div class="ar-station-action">Action: ${escapeHtml(action)}</div>
      </article>
    `;
  }).join("");

  box.querySelectorAll("[data-ar-station-card]").forEach(card => {
    card.addEventListener("click", () => {
      AR_METRICS_STATE.selectedStation = normalizeARStationFilter(card.dataset.arStationCard || "ALL");
      AR_METRICS_STATE.selectedOperator = null;
      closeARAssociateDrawer();
      renderARMetricsPanel(AR_METRICS_STATE.metricsPayload || LAST_CAPACITY_PAYLOAD || {}, AR_METRICS_STATE.rosterPayload || {});
    });
  });
}

function renderARRosterControls(metricsPayload, rosterPayload) {
  const ownerInput = $("arMetricOwner");
  const shiftSelect = $("arMetricShift");

  if (ownerInput && !ownerInput.value) ownerInput.value = AR_METRICS_STATE.owner || "BLOPEZ";
  if (shiftSelect) shiftSelect.value = AR_METRICS_STATE.shift || "Weekday";

  const canSeeRosterAdmin = applyARLMSOnlyVisibility();
  if (!canSeeRosterAdmin) return;

  const operatorSelect = $("arAssignOperator");
  if (operatorSelect) {
    const selected = operatorSelect.value;
    const options = getOperatorMasterRows(rosterPayload, metricsPayload);
    operatorSelect.innerHTML = `<option value="">Select associate</option>` + options.map(row => `
      <option value="${escapeHtml(row.operatorName)}" data-role="${escapeHtml(row.defaultRole)}">${escapeHtml(row.operatorName)}</option>
    `).join("");
    if (selected) operatorSelect.value = selected;
  }

  const roleSelect = $("arAssignRole");
  if (roleSelect && !roleSelect.dataset.floatingRolesReady) {
    const current = roleSelect.value;
    roleSelect.innerHTML = ["AR-IN", "Basket", "Oven", "Sectoring", "DeRing", "AR-OUT", "AR-IN, AR-OUT Floater", "Utility", "Floater"]
      .map(role => `<option value="${escapeHtml(role)}">${escapeHtml(role)}</option>`).join("");
    if (current) roleSelect.value = current;
    roleSelect.dataset.floatingRolesReady = "true";
  }

  renderARRosterList(metricsPayload?.roster || rosterPayload?.roster || []);
}

/*******************************************************
 * PRODUCTION USER CONTROL PATCH — AR DASHBOARD
 * Purpose:
 * - Keep API_URL for AR WIP / production flow / hourly performance.
 * - Use USER_CONTROL_API_URL for AR setup roster / JPH / training / audit log.
 * - Preserve AR floater rule.
 * - Add visible Weekday / Weekend selector inside AR Setup.
 * - Add button loading animation for AR setup actions.
 *******************************************************/

window.AR_USER_CONTROL_PATCH_VERSION = "2026-07-05-3API";
window.AR_USER_CONTROL_JPH_RULES = [];

function arUcApiUrl_() {
  if (typeof USER_CONTROL_API_URL !== "undefined" && USER_CONTROL_API_URL) return USER_CONTROL_API_URL;
  return "https://script.google.com/macros/s/AKfycbwhcxW8dfFw_gJW2s0aZoUP3nhilqEA0S6S-9W9lvPpOtQaWLjAwfFSmo5HrsheP5jR/exec";
}

function arUcClean_(value) {
  return String(value ?? "").trim();
}

function arUcUpper_(value) {
  return arUcClean_(value).toUpperCase();
}

function arUcNumber_(value) {
  if (typeof toNumber === "function") return toNumber(value);
  const n = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

function arUcNormalizeShift_(value) {
  const raw = arUcClean_(value);
  if (/weekend/i.test(raw)) return "Weekend";
  if (/weekday/i.test(raw)) return "Weekday";
  return "Weekday";
}

function arUcGetOwner_() {
  if (typeof getAROwner === "function") return arUcUpper_(getAROwner()) || "BLOPEZ";
  const el = document.getElementById("arMetricOwner");
  return arUcUpper_(el?.value) || "BLOPEZ";
}

function arUcGetShift_() {
  if (typeof getARShift === "function") return arUcNormalizeShift_(getARShift());
  const el = document.getElementById("arMetricShift") || document.getElementById("arSetupShiftSelect");
  return arUcNormalizeShift_(el?.value);
}

function arUcUpdatedBy_() {
  if (typeof getARCurrentUserProfile === "function") {
    const profile = getARCurrentUserProfile();
    if (profile?.username) return arUcUpper_(profile.username);
  }
  return arUcGetOwner_();
}

function arUcRoleToArea_(role) {
  const raw = arUcClean_(role);
  if (typeof normalizeARStationFilter === "function") return normalizeARStationFilter(raw);
  return raw || "AR-IN";
}

function arUcIsSupportRole_(role) {
  const r = arUcRoleToArea_(role).toUpperCase();
  return r === "AR-IN, AR-OUT FLOATER" || r === "UTILITY" || r === "FLOATER" || r === "GENERAL FLOATER";
}

function arUcStatusToRoleType_(status) {
  const s = arUcClean_(status);
  return /^training$/i.test(s) ? "Training" : "Certified";
}

function arUcRoleTypeToCert_(roleType) {
  const s = arUcClean_(roleType);
  return /^training$/i.test(s) ? "Training" : "Certified";
}

async function arUcFetchJson_(url, options = {}) {
  const response = await fetch(url, { cache: "no-store", ...options });
  const text = await response.text();
  let data;

  try {
    data = JSON.parse(text);
  } catch {
    data = { ok: response.ok, raw: text };
  }

  if (!response.ok || data.ok === false || data.status === "error") {
    throw new Error(data.error || data.message || `AR User Control API failed: ${response.status}`);
  }

  return data;
}

async function arUcGetJphRules_() {
  const url = `${arUcApiUrl_()}?action=getJphRules&department=AR`;
  const data = await arUcFetchJson_(url);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  window.AR_USER_CONTROL_JPH_RULES = rows;
  return rows;
}

function arUcFindJphRule_(area, roleType, trainingWeek) {
  const station = arUcRoleToArea_(area);
  const type = arUcRoleTypeToCert_(roleType);
  const week = arUcClean_(trainingWeek);
  const rows = Array.isArray(window.AR_USER_CONTROL_JPH_RULES) ? window.AR_USER_CONTROL_JPH_RULES : [];

  let match = rows.find(row => {
    const rowArea = arUcRoleToArea_(row.Area || row.area || row.DefaultArea || row.defaultArea);
    const rowType = arUcRoleTypeToCert_(row.RoleType || row.roleType);
    const rowWeek = arUcClean_(row.TrainingWeek || row.trainingWeek);
    return rowArea.toUpperCase() === station.toUpperCase() && rowType === type && rowWeek.toUpperCase() === week.toUpperCase();
  });

  if (!match && type === "Certified") {
    match = rows.find(row => {
      const rowArea = arUcRoleToArea_(row.Area || row.area || row.DefaultArea || row.defaultArea);
      const rowType = arUcRoleTypeToCert_(row.RoleType || row.roleType);
      return rowArea.toUpperCase() === station.toUpperCase() && rowType === "Certified";
    });
  }

  if (!match && type === "Training") {
    match = rows.find(row => {
      const rowArea = arUcRoleToArea_(row.Area || row.area || row.DefaultArea || row.defaultArea);
      const rowType = arUcRoleTypeToCert_(row.RoleType || row.roleType);
      return rowArea.toUpperCase() === station.toUpperCase() && rowType === "Certified";
    });
  }

  return match || null;
}

function arUcDefaultJphFor_(area, roleType, trainingWeek) {
  if (arUcIsSupportRole_(area)) return 0;
  const rule = arUcFindJphRule_(area, roleType, trainingWeek);
  if (!rule) return 0;
  return arUcNumber_(rule.DefaultJPH ?? rule.defaultJPH ?? rule.CertifiedJPH ?? rule.certifiedJPH ?? rule.TrainingJPH ?? rule.trainingJPH);
}

function arUcFinalJphFor_(area, roleType, trainingWeek, customJph) {
  const custom = arUcNumber_(customJph);
  if (custom > 0) return custom;
  if (arUcIsSupportRole_(area)) return 0;

  const type = arUcRoleTypeToCert_(roleType);
  const rule = arUcFindJphRule_(area, type, trainingWeek);
  if (!rule) return 0;

  if (type === "Training") {
    return arUcNumber_(rule.TrainingJPH ?? rule.trainingJPH ?? rule.DefaultJPH ?? rule.defaultJPH);
  }

  return arUcNumber_(rule.CertifiedJPH ?? rule.certifiedJPH ?? rule.DefaultJPH ?? rule.defaultJPH ?? rule.TrainingJPH ?? rule.trainingJPH);
}

function arUcRosterRowToLegacy_(row) {
  const area = arUcRoleToArea_(row.DefaultArea || row.defaultArea || row.Area || row.area || row.DefaultRole || row.defaultRole || row.Role || row.role || "AR-IN");
  const roleType = arUcRoleTypeToCert_(row.RoleType || row.roleType || row.CertificationStatus || row.certificationStatus || "Certified");
  const trainingWeek = arUcClean_(row.TrainingWeek || row.trainingWeek || "");
  const defaultJph = arUcNumber_(row.DefaultJPH ?? row.defaultJPH);
  const customJph = arUcNumber_(row.CustomJPH ?? row.customJPH ?? row.IndividualJPH ?? row.individualJPH);
  const finalJph = arUcNumber_(row.FinalJPH ?? row.finalJPH) || arUcFinalJphFor_(area, roleType, trainingWeek, customJph);

  return {
    ...row,
    Area: "AR",
    area: "AR",
    OwnerUsername: row.OwnerUsername || row.ownerUsername || arUcGetOwner_(),
    ownerUsername: row.OwnerUsername || row.ownerUsername || arUcGetOwner_(),
    ShiftType: row.ShiftType || row.shiftType || arUcGetShift_(),
    shiftType: row.ShiftType || row.shiftType || arUcGetShift_(),
    OperatorName: row.OperatorName || row.operatorName || "",
    operatorName: row.OperatorName || row.operatorName || "",
    ActiveStatus: row.ActiveStatus || row.activeStatus || "Active",
    activeStatus: row.ActiveStatus || row.activeStatus || "Active",
    DefaultRole: area,
    defaultRole: area,
    Role: area,
    role: area,
    DefaultArea: area,
    defaultArea: area,
    CertificationStatus: roleType,
    certificationStatus: roleType,
    RoleType: roleType,
    roleType: roleType,
    TrainingWeek: trainingWeek || (roleType === "Training" ? "W1" : ""),
    trainingWeek: trainingWeek || (roleType === "Training" ? "W1" : ""),
    IndividualJPH: customJph,
    individualJph: customJph,
    individualJPH: customJph,
    CustomJPH: customJph,
    customJPH: customJph,
    DefaultJPH: defaultJph,
    defaultJPH: defaultJph,
    FinalJPH: finalJph,
    finalJph: finalJph,
    finalJPH: finalJph,
    JPH: finalJph,
    jph: finalJph
  };
}

function arUcBuildCapacityConfig_(rules) {
  const out = {};
  (Array.isArray(rules) ? rules : []).forEach(row => {
    const area = arUcRoleToArea_(row.Area || row.area);
    if (!area) return;
    const certified = arUcNumber_(row.CertifiedJPH ?? row.certifiedJPH ?? row.DefaultJPH ?? row.defaultJPH);
    if (!out[area]) out[area] = {};
    if (certified > 0) {
      out[area].certifiedJph = certified;
      out[area].CertifiedJPH = certified;
      out[area].defaultJph = certified;
      out[area].DefaultJPH = certified;
      out[area].jph = certified;
      out[area].JPH = certified;
    }
  });
  return out;
}

function arUcBuildRosterSummary_(roster) {
  const active = (Array.isArray(roster) ? roster : []).filter(row => arUcClean_(row.ActiveStatus || row.activeStatus || "Active").toLowerCase() === "active");
  const summary = {
    assignedAssociates: active.length,
    totalBodies: active.length,
    arInOutFloaters: 0,
    utilitySupport: 0,
    activityFloaters: 0,
    extraArInBody: 0,
    extraArOutBody: 0
  };

  active.forEach(row => {
    const role = arUcRoleToArea_(row.DefaultRole || row.defaultRole || row.DefaultArea || row.defaultArea);
    if (role === "AR-IN, AR-OUT Floater") {
      summary.arInOutFloaters += 1;
      summary.extraArInBody += 0.5;
      summary.extraArOutBody += 0.5;
    } else if (role === "Utility") {
      summary.utilitySupport += 1;
    } else if (role === "Floater") {
      summary.activityFloaters += 1;
    }
  });

  return summary;
}

function arUcMergeRosterIntoMetrics_(metricsPayload, rosterPayload) {
  if (!metricsPayload || typeof metricsPayload !== "object") return metricsPayload;
  const roster = Array.isArray(rosterPayload?.roster) ? rosterPayload.roster : [];
  const rules = Array.isArray(rosterPayload?.jphRules) ? rosterPayload.jphRules : window.AR_USER_CONTROL_JPH_RULES;

  metricsPayload.roster = roster;
  metricsPayload.capacityConfig = {
    ...(metricsPayload.capacityConfig || {}),
    ...arUcBuildCapacityConfig_(rules)
  };

  metricsPayload.userControlRoster = true;
  return metricsPayload;
}

async function fetchARRosterControl() {
  const owner = arUcGetOwner_();
  const shift = arUcGetShift_();

  const [rosterData, jphRules] = await Promise.all([
    arUcFetchJson_(`${arUcApiUrl_()}?action=getRoster&department=AR&ownerUsername=${encodeURIComponent(owner)}&shiftType=${encodeURIComponent(shift)}`),
    arUcGetJphRules_().catch(err => {
      console.warn("AR User Control JPH rules failed:", err);
      return [];
    })
  ]);

  const roster = (Array.isArray(rosterData.rows) ? rosterData.rows : []).map(arUcRosterRowToLegacy_);
  const activeRoster = roster.filter(row => arUcClean_(row.ActiveStatus || row.activeStatus || "Active").toLowerCase() === "active");
  const operatorMaster = activeRoster.map(row => ({
    operatorName: row.OperatorName,
    OperatorName: row.OperatorName,
    defaultRole: row.DefaultRole,
    DefaultRole: row.DefaultRole,
    role: row.DefaultRole,
    Role: row.DefaultRole
  })).sort((a, b) => String(a.operatorName).localeCompare(String(b.operatorName)));

  return {
    ok: true,
    status: "success",
    success: true,
    action: "getRoster",
    department: "AR",
    ownerUsername: owner,
    shiftType: shift,
    shiftHours: shift === "Weekend" ? 10.5 : 9.5,
    roles: Array.isArray(AR_ROLE_ORDER) ? AR_ROLE_ORDER : ["AR-IN", "Basket", "Oven", "Sectoring", "DeRing", "AR-OUT"],
    operatorMaster,
    roster,
    jphRules,
    capacityConfig: arUcBuildCapacityConfig_(jphRules),
    summary: arUcBuildRosterSummary_(activeRoster),
    generatedAt: new Date().toISOString()
  };
}

async function fetchARCapacity() {
  const owner = encodeURIComponent(arUcGetOwner_());
  const shift = encodeURIComponent(arUcGetShift_());
  return fetchJson(`${API_URL}?action=getArCapacityMetrics&owner=${owner}&shift=${shift}&debug=true`);
}

async function saveARRosterAssignment(payload = {}) {
  const role = arUcRoleToArea_(payload.role || payload.defaultRole || "AR-IN");
  const roleType = arUcStatusToRoleType_(payload.certificationStatus || payload.certification || payload.status || "Certified");
  const rawWeek = arUcClean_(payload.trainingWeek || payload.week || "");
  const trainingWeek = roleType === "Training"
    ? (/^W/i.test(rawWeek) ? rawWeek.toUpperCase() : `W${rawWeek || 1}`)
    : "";
  const customJph = arUcNumber_(payload.individualJPH ?? payload.individualJph ?? payload.customJPH ?? payload.customJph ?? 0);
  const defaultJph = arUcDefaultJphFor_(role, roleType, trainingWeek);
  const finalJph = arUcFinalJphFor_(role, roleType, trainingWeek, customJph);

  const body = {
    action: "saveRosterRow",
    Department: "AR",
    OwnerUsername: arUcGetOwner_(),
    ShiftType: arUcGetShift_(),
    OperatorName: arUcClean_(payload.operatorName || payload.OperatorName),
    OperatorUsername: arUcClean_(payload.operatorUsername || payload.OperatorUsername || ""),
    ActiveStatus: "Active",
    DefaultArea: role,
    DefaultLine: arUcClean_(payload.line || payload.DefaultLine || ""),
    DefaultPosition: arUcClean_(payload.position || payload.DefaultPosition || ""),
    RoleType: roleType,
    TrainingWeek: trainingWeek,
    TrainingStatus: roleType === "Training" ? "In Training" : "",
    DefaultJPH: defaultJph,
    CustomJPH: customJph || "",
    UseCustomJPH: customJph > 0,
    FinalJPH: finalJph,
    Notes: arUcClean_(payload.notes || payload.Notes || ""),
    UpdatedBy: arUcUpdatedBy_()
  };

  if (!body.OperatorName) throw new Error("AR roster save failed: OperatorName is required.");

  return arUcFetchJson_(`${arUcApiUrl_()}?action=saveRosterRow`, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(body)
  });
}

async function clearARRosterAssignment(operatorName, role) {
  const body = {
    action: "saveRosterRow",
    Department: "AR",
    OwnerUsername: arUcGetOwner_(),
    ShiftType: arUcGetShift_(),
    OperatorName: arUcClean_(operatorName),
    ActiveStatus: "Removed",
    DefaultArea: "",
    DefaultLine: "",
    DefaultPosition: "",
    RoleType: "Certified",
    TrainingWeek: "",
    TrainingStatus: "",
    DefaultJPH: 0,
    CustomJPH: "",
    UseCustomJPH: false,
    FinalJPH: 0,
    Notes: role ? `Removed from ${role}` : "Removed from AR setup",
    UpdatedBy: arUcUpdatedBy_()
  };

  if (!body.OperatorName) throw new Error("AR roster clear failed: OperatorName is required.");

  return arUcFetchJson_(`${arUcApiUrl_()}?action=saveRosterRow`, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(body)
  });
}

async function loadARData() {
  try {
    const [flowPayload, capacityPayload, rosterPayload] = await Promise.all([
      fetchProductionFlowAR(),
      fetchARCapacity().catch(err => {
        console.warn("AR capacity metrics failed:", err);
        return null;
      }),
      fetchARRosterControl().catch(err => {
        console.warn("AR User Control roster failed:", err);
        return null;
      })
    ]);

    LAST_FLOW_PAYLOAD = flowPayload;
    LAST_CAPACITY_PAYLOAD = arUcMergeRosterIntoMetrics_(capacityPayload || null, rosterPayload || null);
    AR_METRICS_STATE.metricsPayload = LAST_CAPACITY_PAYLOAD;
    AR_METRICS_STATE.rosterPayload = rosterPayload || null;
    AR_METRICS_STATE.owner = arUcGetOwner_();
    AR_METRICS_STATE.shift = arUcGetShift_();

    setLiveState(true);
    renderDashboard(LAST_FLOW_PAYLOAD, LAST_CAPACITY_PAYLOAD);

    if (window._arlDismiss) window._arlDismiss(true);
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
    AR_METRICS_STATE.metricsPayload = null;
    renderDashboard(emptyPayload, null);

    if (window._arlDismiss) window._arlDismiss(false);
  }
}

function arUcSetButtonLoading_(button, isLoading, text) {
  if (!button) return;
  if (isLoading) {
    button.dataset.originalText = button.textContent || "";
    button.classList.add("ar-uc-loading");
    button.disabled = true;
    button.textContent = text || "Saving...";
  } else {
    button.classList.remove("ar-uc-loading");
    button.disabled = false;
    button.textContent = button.dataset.originalText || text || button.textContent || "Assign / Update";
  }
}

async function handleSaveARAssignment() {
  if (!isARLMSUser()) {
    showToast("AR assignment control is LMS only.");
    applyARLMSOnlyVisibility();
    return;
  }

  if (AR_METRICS_STATE.isSaving) return;

  const operatorName = document.getElementById("arAssignOperator")?.value || "";
  if (!operatorName) {
    showToast("Select an AR associate first.");
    return;
  }

  const payload = {
    operatorName,
    role: document.getElementById("arAssignRole")?.value || "AR-IN",
    certificationStatus: document.getElementById("arAssignCert")?.value || "Certified",
    trainingWeek: document.getElementById("arAssignWeek")?.value || 1,
    individualJPH: document.getElementById("arAssignCustomJph")?.value || 0
  };

  const btn = document.getElementById("arAssignBtn");

  try {
    AR_METRICS_STATE.isSaving = true;
    arUcSetButtonLoading_(btn, true, "Saving...");
    await saveARRosterAssignment(payload);
    showToast("AR assignment saved to User Control.");
    await loadARData();
  } catch (err) {
    console.error("AR assignment save failed:", err);
    showToast("AR assignment save failed.");
  } finally {
    AR_METRICS_STATE.isSaving = false;
    arUcSetButtonLoading_(btn, false, "Assign / Update");
  }
}

async function handleClearARAssignment(operatorName, role) {
  if (!isARLMSUser()) {
    showToast("AR assignment clear is LMS only.");
    applyARLMSOnlyVisibility();
    return;
  }

  if (!operatorName || AR_METRICS_STATE.isSaving) return;

  try {
    AR_METRICS_STATE.isSaving = true;
    await clearARRosterAssignment(operatorName, role);
    showToast("AR assignment cleared from User Control.");
    await loadARData();
  } catch (err) {
    console.error("AR assignment clear failed:", err);
    showToast("AR assignment clear failed.");
  } finally {
    AR_METRICS_STATE.isSaving = false;
  }
}

function ensureARSetupShiftToolbar_() {
  const setupPage = document.getElementById("arLmsSetupPage") || document.querySelector('[data-page="lmsSetup"]');
  if (!setupPage || document.getElementById("arSetupShiftSelect")) return;

  const target = setupPage.querySelector(".page-head") || setupPage.querySelector(".card-head") || setupPage.firstElementChild || setupPage;
  const toolbar = document.createElement("div");
  toolbar.className = "ar-setup-shift-toolbar";
  toolbar.innerHTML = `
    <label>
      Shift
      <select id="arSetupShiftSelect">
        <option value="Weekday">Weekday · 9.5 hrs</option>
        <option value="Weekend">Weekend · 10.5 hrs</option>
      </select>
    </label>
    <button id="arSetupRefreshBtn" type="button" class="page-pill ar-refresh-pill">Refresh Setup</button>
  `;

  target.appendChild(toolbar);

  const setupShift = document.getElementById("arSetupShiftSelect");
  const mainShift = document.getElementById("arMetricShift");

  if (setupShift) {
    setupShift.value = arUcGetShift_();
    setupShift.addEventListener("change", async () => {
      if (mainShift) mainShift.value = setupShift.value;
      AR_METRICS_STATE.shift = setupShift.value;
      await loadARData();
    });
  }

  const refreshBtn = document.getElementById("arSetupRefreshBtn");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", async () => {
      arUcSetButtonLoading_(refreshBtn, true, "Refreshing...");
      try {
        await loadARData();
      } finally {
        arUcSetButtonLoading_(refreshBtn, false, "Refresh Setup");
      }
    });
  }
}

function injectARUserControlStyles_() {
  if (document.getElementById("arUserControlPatchStyle")) return;
  const style = document.createElement("style");
  style.id = "arUserControlPatchStyle";
  style.textContent = `
    .ar-uc-loading {
      position: relative;
      pointer-events: none;
      opacity: .82;
      animation: arUcPulse .7s infinite alternate;
    }
    .ar-uc-loading::after {
      content: "";
      display: inline-block;
      width: 10px;
      height: 10px;
      margin-left: 8px;
      border-radius: 50%;
      border: 2px solid currentColor;
      border-top-color: transparent;
      vertical-align: -1px;
      animation: arUcSpin .7s linear infinite;
    }
    .ar-setup-shift-toolbar {
      display: flex;
      align-items: end;
      gap: 10px;
      justify-content: flex-end;
      flex-wrap: wrap;
      margin-top: 10px;
    }
    .ar-setup-shift-toolbar label {
      display: grid;
      gap: 6px;
      color: var(--text-dim, #9ca3af);
      font-size: 9px;
      letter-spacing: 1.2px;
      font-weight: 950;
      text-transform: uppercase;
    }
    .ar-setup-shift-toolbar select {
      min-height: 36px;
      border: 1px solid var(--panel-border-strong, rgba(255,255,255,.18));
      border-radius: 10px;
      background: rgba(10, 11, 17, 0.92);
      color: var(--text, #fff);
      padding: 0 11px;
      outline: 0;
      font-weight: 800;
      min-width: 160px;
    }
    @keyframes arUcPulse { from { transform: scale(1); filter: brightness(1); } to { transform: scale(1.025); filter: brightness(1.25); } }
    @keyframes arUcSpin { to { transform: rotate(360deg); } }
  `;
  document.head.appendChild(style);
}

document.addEventListener("DOMContentLoaded", () => {
  injectARUserControlStyles_();
  setTimeout(ensureARSetupShiftToolbar_, 300);
  setTimeout(ensureARSetupShiftToolbar_, 1200);
  setTimeout(ensureARSetupShiftToolbar_, 2500);
});

console.log("[AR User Control] 3-API patch loaded. WIP/hourly = production API. AR setup roster/JPH = User Control API.");
