/* =====================================================
   INDEX DASHBOARD – FINAL PRODUCTION VERSION
   ===================================================== */

/* 🔗 DATA API */
const API_URL =
"https://script.google.com/a/macros/zennioptical.com/s/AKfycbzbQBjzoEEBpvukFkR-XMw8kG_gzCIuxZrTLodZZ_EnwqYAujOBqSzYslx-x9XTw7_UUA/exec";

/* 🔐 AUTH API */
const AUTH_API =
"https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

/* ===============================
   AUTH GATE
   =============================== */
(function authGate() {
  if (sessionStorage.getItem("lms_logged_in") !== "true") {
    window.location.replace("login.html");
  }
})();

/* ===============================
   STORAGE
   =============================== */
const SCANNER_STORAGE_KEY = "lms_scanner_attention";

/* ===============================
   PAGE LOAD
   =============================== */
document.addEventListener("DOMContentLoaded", () => {

  console.time("TOTAL_DASHBOARD_LOAD");

  const username = sessionStorage.getItem("lms_user");

  Promise.all([
    fetchDashboard(),
    fetchVisibility(username)
  ])
  .then(([dashboardData, visibilityData]) => {

    renderDashboard(dashboardData);

    if (visibilityData.status === "SUCCESS") {
      applyVisibilityRules(visibilityData.visibility);
    }

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

/* ===============================
   FETCH DASHBOARD
   =============================== */
function fetchDashboard() {

  console.time("DASHBOARD_API");

  return fetch(API_URL + "?t=" + Date.now())
    .then(res => res.json())
    .then(data => {

      console.timeEnd("DASHBOARD_API");

      return data;

    });

}

/* ===============================
   FETCH VISIBILITY
   =============================== */
function fetchVisibility(username) {

  console.time("VISIBILITY_API");

  if (!username) {
    return Promise.resolve({ status: "SUCCESS", visibility: {} });
  }

  return fetch(`${AUTH_API}?action=visibility&username=${encodeURIComponent(username)}`)
    .then(res => res.json())
    .then(data => {

      console.timeEnd("VISIBILITY_API");

      return data;

    });

}

/* ===============================
   RENDER DASHBOARD
   =============================== */
function renderDashboard(data) {

  const active = Number(data.active ?? 0);
  const completed = Number(data.completed ?? 0);

  /* ===============================
     FIX COVERAGE
     =============================== */
  let coverageRaw = Number(data.coverage);
  if (isNaN(coverageRaw)) coverageRaw = 0;

  let coveragePct =
    coverageRaw <= 1
      ? Math.floor(coverageRaw * 1000) / 10
      : Math.floor(coverageRaw * 10) / 10;

  coveragePct = Math.max(0, Math.min(100, coveragePct));

  /* Fallback if API bad */
  if (completed > 0 && coveragePct === 0) {
    coveragePct = Math.round((completed / (completed + active)) * 100);
  }

  /* ===============================
     LAST JOB COMPLETED
     =============================== */
  const lastCompletedStr = data.lastCompletedAt || "N/A";

  const sinceEl = document.getElementById("lastSince");
  const lastValEl = document.getElementById("lastUpdate");

  if (lastValEl) {
    lastValEl.classList.remove("ok", "warn", "bad");
  }

  if (lastCompletedStr !== "N/A") {

    const parsed = new Date(lastCompletedStr);

    if (!isNaN(parsed)) {

      /* 🔹 Format date */
      setText("lastUpdate",
        parsed.toLocaleDateString(undefined, {
          month: "short",
          day: "2-digit",
          year: "numeric"
        })
      );

      const now = new Date();
      const diffMs = now - parsed;

      const diffMin = Math.floor(diffMs / 60000);
      const diffHr = Math.floor(diffMin / 60);
      const diffDay = Math.floor(diffHr / 24);

      /* 🔹 Smart "time ago" */
      if (sinceEl) {
        if (diffMin < 1) sinceEl.textContent = "just now";
        else if (diffMin < 60) sinceEl.textContent = `${diffMin} min ago`;
        else if (diffHr < 24) sinceEl.textContent = `${diffHr} hr ago`;
        else sinceEl.textContent = `${diffDay} day ago`;
      }

      /* 🔹 Status color */
      if (lastValEl) {
        if (diffMin < 5) lastValEl.classList.add("ok");
        else if (diffHr < 8) lastValEl.classList.add("warn");
        else lastValEl.classList.add("bad");
      }

    } else {

      setText("lastUpdate", lastCompletedStr);
      if (sinceEl) sinceEl.textContent = "Invalid time";

    }

  } else {

    setText("lastUpdate", "—");
    if (sinceEl) sinceEl.textContent = "No completions yet";
    if (lastValEl) lastValEl.classList.add("warn");

  }

  /* ===============================
     UPDATE UI
     =============================== */
  setText("active", active);
  setText("activeHolds", active);
  setText("completed", completed);
  setText("coverage", coveragePct + "%");
  setText("coverageDetail", coveragePct + "%");

  updateLabStatus(active, coveragePct);

}

/* ===============================
   VISIBILITY
   =============================== */
function applyVisibilityRules(visibility) {

  const map = {
    "Scanner Map": "#scanner-card",
    "Investigation Hold": "#investigation-card",
    "True Curve": "#truecurve-card",
    "Tools": "#tools-card",
    "Coating": "#coating-card",
    "Admin": ".admin-only"
  };

  Object.keys(map).forEach(feature => {

    if (visibility[feature] === true) {

      document.querySelectorAll(map[feature]).forEach(el => {
        el.style.display = "flex";
      });

    }

  });

}

/* ===============================
   SCANNER STATUS
   =============================== */
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

/* ===============================
   LAB STATUS (FIXED LOGIC)
   =============================== */
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

/* ===============================
   ERROR STATE
   =============================== */
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

/* ===============================
   UNLOCK PAGE
   =============================== */
function unlockPage() {
  document.body.classList.remove("lms-hidden");
}

/* ===============================
   LOGOUT
   =============================== */
function logout() {

  sessionStorage.clear();
  localStorage.removeItem(SCANNER_STORAGE_KEY);

  window.location.href = "login.html";

}

/* ===============================
   HELPERS
   =============================== */
function setText(id,value) {

  const el = document.getElementById(id);
  if (el) el.textContent = value;

}

function openPage(page) {
  window.location.href = page;
}