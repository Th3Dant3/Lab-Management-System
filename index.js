/* =====================================================
   INDEX DASHBOARD – PRODUCTION UPGRADE
===================================================== */

/* 🔗 DATA API */
const API_URL =
"https://script.google.com/a/macros/zennioptical.com/s/AKfycbzbQBjzoEEBpvukFkR-XMw8kG_gzCIuxZrTLodZZ_EnwqYAujOBqSzYslx-x9XTw7_UUA/exec";

/* 🔐 AUTH API */
const AUTH_API =
"https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

/* STORAGE */
const SCANNER_STORAGE_KEY = "lms_scanner_attention";

/* =====================================================
   AUTH GATE
===================================================== */
(function authGate() {
  if (sessionStorage.getItem("lms_logged_in") !== "true") {
    window.location.replace("login.html");
  }
})();

/* =====================================================
   PAGE LOAD
===================================================== */
document.addEventListener("DOMContentLoaded", () => {

  console.time("TOTAL_DASHBOARD_LOAD");

  initTabs(); // 🔥 NEW

  const username = sessionStorage.getItem("lms_user");

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

    updateScannerFromStorage();
    unlockPage();

    console.timeEnd("TOTAL_DASHBOARD_LOAD");

  });

});

/* =====================================================
   TAB SYSTEM (NEW)
===================================================== */
function initTabs() {

  const tabs = document.querySelectorAll(".tab");
  const contents = document.querySelectorAll(".tab-content");

  tabs.forEach(tab => {
    tab.addEventListener("click", () => {

      tabs.forEach(t => t.classList.remove("active"));
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

    const res = await fetch(API_URL + "?t=" + Date.now());
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

    const res = await fetch(
      `${AUTH_API}?action=visibility&username=${encodeURIComponent(username)}`
    );

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

  const active = Number(data.active ?? 0);
  const completed = Number(data.completed ?? 0);

  /* COVERAGE */
  let coverage = Number(data.coverage ?? 0);

  if (coverage <= 1) coverage *= 100;
  coverage = Math.max(0, Math.min(100, Math.round(coverage)));

  if (completed > 0 && coverage === 0) {
    coverage = Math.round((completed / (completed + active)) * 100);
  }

  /* AGING */
  const oldest = Number(data.oldestActiveMinutes ?? 0);
  const stuck = Number(data.stuckJobs ?? 0);

  updateAgingDisplay(active, oldest, stuck);

  /* UI UPDATE */
  setText("active", active);
  setText("activeHolds", active);
  setText("completed", completed);
  setText("coverage", coverage + "%");

  setText(
    "coverageDetail",
    active === 0 ? "Complete" : coverage + "%"
  );

  updateLabStatus(active, coverage);

}

/* =====================================================
   AGING DISPLAY (NEW CLEAN VERSION)
===================================================== */
function updateAgingDisplay(active, oldest, stuck) {

  const lastValEl = document.getElementById("lastUpdate");
  const sinceEl = document.getElementById("lastSince");

  if (!lastValEl || !sinceEl) return;

  lastValEl.classList.remove("ok", "warn", "bad");

  if (active === 0) {

    setText("lastUpdate", "No Active Jobs");
    sinceEl.textContent = "All clear";
    return;

  }

  const label =
    oldest >= 60
      ? `${Math.floor(oldest / 60)}h ${oldest % 60}m`
      : `${oldest} min`;

  setText("lastUpdate", label);
  sinceEl.textContent = `${stuck} stuck job${stuck === 1 ? "" : "s"}`;

  if (oldest < 30) lastValEl.classList.add("ok");
  else if (oldest < 90) lastValEl.classList.add("warn");
  else lastValEl.classList.add("bad");

}

/* =====================================================
   VISIBILITY (SAFE DEFAULT HIDE)
===================================================== */
function applyVisibilityRules(visibility) {

  const map = {
    "Scanner Map": "#scanner-card",
    "Investigation Hold": "#investigation-card",
    "True Curve": "#truecurve-card",
    "Tools": "#tools-card",
    "Coating": "#coating-card",
    "Incoming Jobs": "#incoming-card", // 🔥 ADD THIS
    "Admin": ".admin-only"
  };

  // Hide all first (safe default)
  Object.values(map).forEach(selector => {
    document.querySelectorAll(selector).forEach(el => {
      el.style.display = "none";
    });
  });

  // Enable allowed features
  Object.keys(map).forEach(feature => {
    if (visibility[feature]) {
      document.querySelectorAll(map[feature]).forEach(el => {
        el.style.display = "flex";
      });
    }
  });

}

/* =====================================================
   SCANNER STATUS
===================================================== */
function updateScannerFromStorage() {

  let scanners = [];

  try {
    scanners = JSON.parse(localStorage.getItem(SCANNER_STORAGE_KEY)) || [];
  } catch {}

  const health = document.getElementById("scannerHealth");
  const detail = document.getElementById("scannerDetail");

  if (!health || !detail) return;

  if (!scanners.length) {

    health.textContent = "Online";
    health.className = "value ok";
    detail.innerHTML = "<p class='sub'>No scanners currently flagged</p>";
    return;

  }

  health.textContent = "Attention";
  health.className = "value warn";

  detail.innerHTML = `
    <strong>Scanner Attention:</strong>
    <ul>${scanners.map(s => `<li>${s}</li>`).join("")}</ul>
  `;

}

/* =====================================================
   LAB STATUS
===================================================== */
function updateLabStatus(active, coverage) {

  const el = document.getElementById("labStatus");
  if (!el) return;

  el.classList.remove("ok","warn","bad");

  if (active === 0 && coverage >= 95) {
    el.textContent = "Normal";
    el.classList.add("ok");
  }
  else if (active <= 3) {
    el.textContent = "Attention";
    el.classList.add("warn");
  }
  else {
    el.textContent = "Issue";
    el.classList.add("bad");
  }

}

/* =====================================================
   ERROR STATE
===================================================== */
function showErrorState() {

  [
    "active",
    "coverage",
    "activeHolds",
    "completed",
    "coverageDetail",
    "scannerHealth",
    "lastUpdate",
    "lastSince"
  ].forEach(id => setText(id, "ERR"));

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