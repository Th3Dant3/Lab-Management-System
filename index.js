/* =====================================================
   INDEX DASHBOARD – FINAL, STABLE, USERNAME-BASED
   ===================================================== */

/* 🔗 DATA API (dashboard metrics) */
const API_URL =
  "https://script.google.com/a/macros/zennioptical.com/s/AKfycbzbQBjzoEEBpvukFkR-XMw8kG_gzCIuxZrTLodZZ_EnwqYAujOBqSzYslx-x9XTw7_UUA/exec";

/* 🔐 AUTH / VISIBILITY API (USERNAME-BASED) */
const AUTH_API =
  "https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

/* ===============================
   AUTH GATE (HARD BLOCK)
   =============================== */
(function authGate() {
  if (sessionStorage.getItem("lms_logged_in") !== "true") {
    window.location.replace("login.html");
  }
})();

/* ===============================
   SCANNER STORAGE
   =============================== */
const SCANNER_STORAGE_KEY = "lms_scanner_attention";

/* ===============================
   PAGE LOAD
   =============================== */
document.addEventListener("DOMContentLoaded", () => {
  loadDashboard();
  updateScannerFromStorage();
  applyVisibility(); // 🔐 username-based gating
});

/* ===============================
   DASHBOARD DATA
   =============================== */
function loadDashboard() {
  fetch(API_URL)
    .then(res => res.json())
    .then(data => {
      const active = Number(data.active ?? 0);
      const completed = Number(data.completed ?? 0);

      let coveragePct = 0;
      if (typeof data.coverage === "string") {
        coveragePct = Number(data.coverage.replace("%", ""));
      } else if (typeof data.coverage === "number") {
        coveragePct =
          data.coverage <= 1
            ? Math.round(data.coverage * 100)
            : Math.round(data.coverage);
      }

      coveragePct = Math.max(0, Math.min(100, coveragePct));

      setText("active", active);
      setText("coverage", coveragePct + "%");
      setText("activeHolds", active);
      setText("completed", completed);
      setText("coverageDetail", coveragePct + "%");
      setText("lastUpdate", data.lastUpdated || "N/A");

      updateLabStatus(active, coveragePct);
    })
    .catch(() => showErrorState());
}

/* ===============================
   🔐 VISIBILITY (USERNAME-BASED)
   =============================== */
function applyVisibility() {
  const username = sessionStorage.getItem("lms_user");

  if (!username) {
    console.warn("No username in session");
    unlockPage();
    return;
  }

  fetch(
    `${AUTH_API}?action=visibility&username=${encodeURIComponent(username)}`
  )
    .then(res => res.json())
    .then(data => {
      if (data.status !== "SUCCESS") {
        console.warn("Visibility denied", data);
        return;
      }

      applyVisibilityRules(data.visibility);
    })
    .catch(err => {
      console.error("Visibility fetch failed", err);
    })
    .finally(() => {
      unlockPage(); // 🚫 NO FLICKER
    });
}

/* ===============================
   APPLY VISIBILITY RULES
   =============================== */
function applyVisibilityRules(visibility) {
  const map = {
    "Scanner Map": "#scanner-card",
    "Investigation Hold": "#investigation-card",
    "True Curve": "#truecurve-card",
    "Tools": "#tools-card",
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
   UNLOCK PAGE (NO FLICKER)
   =============================== */
function unlockPage() {
  document.body.classList.remove("lms-hidden");
}

/* ===============================
   SCANNER MAP STATUS
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
   STATUS LOGIC
   =============================== */
function updateLabStatus(active, coverage) {
  const el = document.getElementById("labStatus");
  if (!el) return;

  el.classList.remove("ok", "warn", "bad");

  if (active === 0 && coverage >= 98) {
    el.textContent = "Normal";
    el.classList.add("ok");
  } else if (active <= 3) {
    el.textContent = "Attention";
    el.classList.add("warn");
  } else {
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
    "lastUpdate"
  ].forEach(id => setText(id, "ERR"));

  const el = document.getElementById("labStatus");
  if (el) {
    el.textContent = "Offline";
    el.className = "value bad";
  }
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
function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function openPage(page) {
  window.location.href = page;
}
