/* =====================================================
   INDEX DASHBOARD – FULL JS (EXTENDED, NON-BREAKING)
   ===================================================== */

const API_URL =
  "https://script.google.com/a/macros/zennioptical.com/s/AKfycbzbQBjzoEEBpvukFkR-XMw8kG_gzCIuxZrTLodZZ_EnwqYAujOBqSzYslx-x9XTw7_UUA/exec";

/* ===============================
   🔹 NEW: SCANNER STORAGE
   =============================== */

const SCANNER_STORAGE_KEY = "lms_scanner_attention";

/* ===============================
   PAGE LOAD
   =============================== */

document.addEventListener("DOMContentLoaded", () => {
  loadDashboard();
  updateScannerFromStorage();
});

/* ===============================
   DASHBOARD (UNCHANGED CORE)
   =============================== */

function loadDashboard() {
  fetch(API_URL)
    .then(res => res.json())
    .then(data => {
      const active = Number(data.active ?? 0);
      const completed = Number(data.completed ?? 0);
      const coveragePct = Math.floor((data.coverage || 0) * 100);

      // Executive strip
      setText("active", active);
      setText("coverage", coveragePct + "%");

      // Detail card
      setText("activeHolds", active);
      setText("completed", completed);
      setText("coverageDetail", coveragePct + "%");

      // Investigation-driven status (baseline)
      updateLabStatus(active, coveragePct);

      // Timestamp
      setText("lastUpdate", data.lastUpdated || "N/A");
    })
    .catch(err => {
      console.error("Dashboard load failed", err);
      showErrorState();
    });
}

/* ===============================
   🔹 NEW: SCANNER INTEGRATION
   =============================== */

function updateScannerFromStorage() {
  let scanners = [];

  try {
    scanners = JSON.parse(localStorage.getItem(SCANNER_STORAGE_KEY)) || [];
  } catch (e) {
    console.warn("Scanner storage read failed", e);
  }

  const scannerHealthEl = document.getElementById("scannerHealth");
  const scannerDetailEl = document.getElementById("scannerDetail");

  if (!scannerHealthEl || !scannerDetailEl) return;

  // No scanners selected
  if (!scanners.length) {
    scannerHealthEl.textContent = "Online";
    scannerHealthEl.className = "value ok";
    scannerDetailEl.innerHTML =
      "<p class='sub'>No scanners currently flagged</p>";
    return;
  }

  // One or more scanners flagged
  scannerHealthEl.textContent = "Attention";
  scannerHealthEl.className = "value warn";

  scannerDetailEl.innerHTML = `
    <strong>Scanner Attention:</strong>
    <ul>
      ${scanners.map(s => `<li>${s}</li>`).join("")}
    </ul>
  `;
}

/* ===============================
   STATUS LOGIC (UNCHANGED)
   =============================== */

function updateLabStatus(active, coverage) {
  const statusEl = document.getElementById("labStatus");
  if (!statusEl) return;

  statusEl.classList.remove("ok", "warn", "bad");

  if (active === 0 && coverage >= 98) {
    statusEl.textContent = "Normal";
    statusEl.classList.add("ok");
  } else if (active <= 3) {
    statusEl.textContent = "Attention";
    statusEl.classList.add("warn");
  } else {
    statusEl.textContent = "Issue";
    statusEl.classList.add("bad");
  }
}

/* ===============================
   ERROR STATE (UNCHANGED)
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

  const statusEl = document.getElementById("labStatus");
  if (statusEl) {
    statusEl.textContent = "Offline";
    statusEl.classList.remove("ok", "warn");
    statusEl.classList.add("bad");
  }
}

/* ===============================
   HELPERS (UNCHANGED)
   =============================== */

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function openPage(page) {
  window.location.href = page;
}
