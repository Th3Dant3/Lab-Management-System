/* =====================================================
   INDEX DASHBOARD — Command Center Redesign
===================================================== */

const API_URL =
  "https://script.google.com/a/macros/zennioptical.com/s/AKfycbzbQBjzoEEBpvukFkR-XMw8kG_gzCIuxZrTLodZZ_EnwqYAujOBqSzYslx-x9XTw7_UUA/exec";

const AUTH_API =
  "https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

const SCANNER_STORAGE_KEY = "lms_scanner_attention";

/* =====================================================
   PAGE LOAD
===================================================== */
document.addEventListener("DOMContentLoaded", () => {

  console.time("TOTAL_DASHBOARD_LOAD");

  initClock();
  initTabs();
  initUserDisplay();
  updateScannerFromStorage();

  // Show the UI immediately — don't wait for API
  unlockPage();

  const username = sessionStorage.getItem("lms_user");

  // Data loads in background and fills in when ready
  Promise.all([
    fetchDashboard(),
    fetchVisibility(username)
  ])
  .then(([dashboardData, visibilityData]) => {
    renderDashboard(dashboardData);
    applyVisibilityRules(visibilityData.visibility || {});
  })
  .catch(err => {
    console.error("Dashboard load error", err);
    showErrorState();
  })
  .finally(() => {
    console.timeEnd("TOTAL_DASHBOARD_LOAD");
  });

});

/* =====================================================
   LIVE CLOCK
===================================================== */
function initClock() {
  function tick() {
    // Command bar uses id="topbarTime"
    const el = document.getElementById("topbarTime");
    if (!el) return;
    const now = new Date();
    el.textContent = now.toLocaleTimeString([], {
      hour: "2-digit", minute: "2-digit", second: "2-digit"
    });
  }
  tick();
  setInterval(tick, 1000);
}

/* =====================================================
   USER DISPLAY
===================================================== */
function initUserDisplay() {
  const username  = sessionStorage.getItem("lms_user") || "";
  const avatarEl  = document.getElementById("userAvatar");
  const nameEl    = document.getElementById("userName");

  if (!username) return;

  const initials = username
    .split(/[\s._-]+/)
    .slice(0, 2)
    .map(w => w[0])
    .join("")
    .toUpperCase();

  if (avatarEl) avatarEl.textContent = initials || username[0]?.toUpperCase() || "U";
  if (nameEl)   nameEl.textContent   = username;
}

/* =====================================================
   TAB SYSTEM
===================================================== */
function initTabs() {
  const tabs     = document.querySelectorAll(".nav-item[data-tab]");
  const contents = document.querySelectorAll(".tab-content");

  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(t     => t.classList.remove("active"));
      contents.forEach(c => c.classList.remove("active"));

      tab.classList.add("active");
      const target = document.getElementById(tab.dataset.tab);
      if (target) target.classList.add("active");
    });
  });
}

/* =====================================================
   FETCH DASHBOARD
===================================================== */
async function fetchDashboard() {
  console.time("DASHBOARD_API");
  try {
    const res  = await fetch(API_URL + "?t=" + Date.now());
    const data = await res.json();
    console.timeEnd("DASHBOARD_API");
    return data || {};
  } catch (err) {
    console.error("Dashboard API failed", err);
    return {};
  }
}

/* =====================================================
   FETCH VISIBILITY
===================================================== */
async function fetchVisibility(username) {
  console.time("VISIBILITY_API");
  if (!username) return { status: "SUCCESS", visibility: {} };
  try {
    const res  = await fetch(`${AUTH_API}?action=visibility&username=${encodeURIComponent(username)}`);
    const data = await res.json();
    console.timeEnd("VISIBILITY_API");
    return data || { status: "SUCCESS", visibility: {} };
  } catch (err) {
    console.error("Visibility API failed", err);
    return { status: "SUCCESS", visibility: {} };
  }
}

/* =====================================================
   RENDER DASHBOARD
===================================================== */
function renderDashboard(data) {
  const active    = Number(data.active    ?? 0);
  const completed = Number(data.completed ?? 0);

  // Total Incoming — shown in command bar KPI
  const total = Number(data.totalIncoming ?? data.incoming ?? 0);

  let coverage = Number(data.coverage ?? 0);
  if (coverage <= 1) coverage *= 100;
  coverage = Math.max(0, Math.min(100, Math.round(coverage)));
  if (completed > 0 && coverage === 0) {
    coverage = Math.round((completed / (completed + active)) * 100);
  }

  // Command bar KPIs
  setText("totalIncoming", total || "—");
  setText("active",        active);

  // Card metrics (investigation card)
  setText("activeHolds",   active);
  setText("completed",     completed);
  setText("coverageDetail", active === 0 ? "Complete" : coverage + "%");
}

/* =====================================================
   VISIBILITY RULES
===================================================== */
function applyVisibilityRules(visibility) {
  const map = {
    "Scanner Map":        "#scanner-card",
    "Investigation Hold": "#investigation-card",
    "True Curve":         "#truecurve-card",
    "Tools":              "#tools-card",
    "Coating":            "#coating-card",
    "Incoming Jobs":      "#incoming-card",
    "Admin":              ".admin-only"
  };

  // Hide all gated cards first
  Object.values(map).forEach(sel => {
    document.querySelectorAll(sel).forEach(el => el.style.display = "none");
  });

  // Show permitted cards
  Object.keys(map).forEach(feature => {
    if (visibility[feature]) {
      document.querySelectorAll(map[feature]).forEach(el => el.style.display = "flex");
    }
  });

  // Power Analysis Tool — show when Coating OR Production permission granted
  if (visibility["Coating"] || visibility["Production"]) {
    const pw = document.getElementById("power-card");
    if (pw) pw.style.display = "flex";
  }
}

/* =====================================================
   SCANNER STATUS (localStorage)
===================================================== */
function updateScannerFromStorage() {
  let scanners = [];
  try {
    scanners = JSON.parse(localStorage.getItem(SCANNER_STORAGE_KEY)) || [];
  } catch {}

  const health = document.getElementById("scannerHealth");
  const sub    = document.getElementById("scannerSub");
  const detail = document.getElementById("scannerDetail");

  if (!scanners.length) {
    if (health) { health.textContent = "Online"; health.className = "cb-m-val cb-m-val--green"; }
    if (sub)    sub.textContent = "All ports";
    if (detail) detail.innerHTML = "";
  } else {
    if (health) { health.textContent = "Attention"; health.className = "cb-m-val cb-m-val--orange"; }
    if (sub)    sub.textContent = `${scanners.length} flagged`;
    if (detail) {
      detail.innerHTML = `<strong>Flagged Scanners:</strong><ul>${scanners.map(s => `<li>${s}</li>`).join("")}</ul>`;
    }
  }
}

/* =====================================================
   ERROR STATE
===================================================== */
function showErrorState() {
  ["totalIncoming","active","activeHolds","completed","coverageDetail","scannerHealth"]
    .forEach(id => setText(id, "ERR"));
}

/* =====================================================
   UNLOCK PAGE
===================================================== */
function unlockPage() {
  document.body.classList.remove("lms-hidden");
}

/* =====================================================
   LOGOUT
===================================================== */
function logout() {
  sessionStorage.clear();
  localStorage.removeItem(SCANNER_STORAGE_KEY);
  window.location.href = "login.html";
}

/* =====================================================
   HELPERS
===================================================== */
function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function openPage(page) {
  window.location.href = page;
}
