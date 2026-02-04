/* =====================================================
   INDEX DASHBOARD – FULL JS (EXTENDED, NON-BREAKING)
   ===================================================== */

const API_URL =
  "https://script.google.com/a/macros/zennioptical.com/s/AKfycbzbQBjzoEEBpvukFkR-XMw8kG_gzCIuxZrTLodZZ_EnwqYAujOBqSzYslx-x9XTw7_UUA/exec";

/* ===============================
   🔹 SCANNER STORAGE
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
   DASHBOARD LOAD
   =============================== */

function loadDashboard() {
  fetch(API_URL)
    .then(res => res.json())
    .then(data => {

      const active = Number(data.active ?? 0);
      const completed = Number(data.completed ?? 0);

      /* ===============================
         🔹 FIXED COVERAGE HANDLING
         =============================== */

      let coveragePct = 0;

      if (data.coverage !== undefined && data.coverage !== null) {
        if (typeof data.coverage === "string") {
          coveragePct = Number(data.coverage.replace("%", ""));
        } else if (typeof data.coverage === "number") {
          // If backend sends 0–1 ratio
          coveragePct = data.coverage <= 1
            ? Math.round(data.coverage * 100)
            : Math.round(data.coverage);
        }
      }

      // Safety clamp
      coveragePct = Math.max(0, Math.min(100, coveragePct));

      /* ===============================
         DASHBOARD UI UPDATES
         =============================== */

      // Executive strip
      setText("active", active);
      setText("coverage", coveragePct + "%");

      // Detail cards
      setText("activeHolds", active);
      setText("completed", completed);
      setText("coverageDetail", coveragePct + "%");

      // Status logic
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
   🔹 SCANNER MAP INTEGRATION
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

  // ✅ No scanners flagged
  if (!scanners.length) {
    scannerHealthEl.textContent = "Online";
    scannerHealthEl.className = "value ok";
    scannerDetailEl.innerHTML =
      "<p class='sub'>No scanners currently flagged</p>";
    return;
  }

  // ⚠️ One or more scanners flagged
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
   STATUS LOGIC (UNCHANGED CORE)
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

  const statusEl = document.getElementById("labStatus");
  if (statusEl) {
    statusEl.textContent = "Offline";
    statusEl.classList.remove("ok", "warn");
    statusEl.classList.add("bad");
  }
}

function logout() {
  sessionStorage.clear();
  localStorage.removeItem("lms_scanner_attention");
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
