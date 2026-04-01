/* =====================================================
   TRUE CURVE — UPGRADED JS v2
   Clean architecture, search, count-up KPIs,
   status badges, stagger animation
===================================================== */

const API = "https://script.google.com/macros/s/AKfycbzuUpwsa2VTGjxeSP3wVV-x2Z4yMkGH_eYS86VpFH0CzC2Ii5S5U2ag-S5d89RZ8oHJ/exec";

/* ─── STATE ─── */

const state = {
  year:    "2026",
  mode:    "production",
  view:    "material",
  query:   "",
  rows:    [],
  loading: false
};

/* ─── STATUS BADGE MAP ─── */
// Keys are always lowercased for case-insensitive matching

const STATUS_MAP = {
  "active":              { cls: "badge-active",    label: "Active"             },
  "shipped":             { cls: "badge-shipped",   label: "Shipped"            },
  "mailroom":            { cls: "badge-shipped",   label: "Mailroom"           },
  "escalated":           { cls: "badge-escalated", label: "Escalated"          },
  "breakage to picking": { cls: "badge-escalated", label: "Breakage to Picking"},
  "test":                { cls: "badge-test",       label: "Test"              },
  "cancel":              { cls: "badge-cancel",    label: "Cancel"             },
};

function statusBadge(raw) {
  if (!raw) return "";
  const key = raw.trim().toLowerCase();
  const map = STATUS_MAP[key] || null;
  if (map) return `<span class="badge ${map.cls}">${map.label}</span>`;
  // fallback: render as-is with a neutral style
  return `<span class="badge badge-cancel">${raw}</span>`;
}

/* ─── COUNT-UP ANIMATION ─── */

function countUp(el, target, duration = 800) {
  const start = performance.now();
  const from = parseInt(el.textContent) || 0;
  function step(now) {
    const progress = Math.min((now - start) / duration, 1);
    const ease = 1 - Math.pow(1 - progress, 3); // ease-out cubic
    el.textContent = Math.round(from + (target - from) * ease);
    if (progress < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

/* ─── STATUS HELPERS ─── */

// Case-insensitive status check
const is = (row, ...statuses) =>
  statuses.some(s => (row.status || "").trim().toLowerCase() === s.toLowerCase());

const DONE_STATUSES    = ["Shipped", "Mailroom"];
const TEST_STATUSES    = ["Test"];
const CANCEL_STATUSES  = ["Cancel"];
const SKIP_STATUSES    = [...DONE_STATUSES, ...TEST_STATUSES, ...CANCEL_STATUSES];

/* ─── KPI UPDATE ─── */

function updateKPIs() {
  const all = state.rows;
  const kpis = {
    "k-total":     all.length,
    "k-active":    all.filter(r => !is(r, ...SKIP_STATUSES)).length,
    "k-completed": all.filter(r => is(r, ...DONE_STATUSES)).length,
    "k-adjusted":  all.filter(r => r.adjustment && r.adjustment !== "No Adjustment").length,
    "k-escalated": all.filter(r => is(r, "Escalated", "Breakage to Picking")).length,
    "k-test":      all.filter(r => is(r, ...TEST_STATUSES)).length,
    "k-cancel":    all.filter(r => is(r, ...CANCEL_STATUSES)).length,
  };
  Object.entries(kpis).forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el) countUp(el, val);
  });
}

/* ─── FILTER ROWS ─── */

function getFilteredRows() {
  let rows = [...state.rows];

  // mode filter
  if (state.mode === "production") {
    rows = rows.filter(r => !is(r, ...TEST_STATUSES, ...CANCEL_STATUSES));
  } else if (state.mode === "completed") {
    rows = rows.filter(r => is(r, ...DONE_STATUSES));
  } else if (state.mode === "test") {
    rows = rows.filter(r => is(r, ...TEST_STATUSES, ...CANCEL_STATUSES));
  }

  // search filter
  if (state.query) {
    const q = state.query.toLowerCase();
    rows = rows.filter(r =>
      (r.rx       || "").toLowerCase().includes(q) ||
      (r.sku      || "").toLowerCase().includes(q) ||
      (r.material || "").toLowerCase().includes(q) ||
      (r.status   || "").toLowerCase().includes(q)
    );
  }

  return rows;
}

/* ─── ROW HTML ─── */

function rowHTML(r) {
  return `
    <td>${r.rx || ""}</td>
    <td>${r.sku || ""}</td>
    <td>${r.tcLms || ""}</td>
    <td>${r.tcAfter || ""}</td>
    <td>${r.material || ""}</td>
    <td>${statusBadge(r.status)}</td>
    <td>${r.dateShipped || ""}</td>
    <td>${r.breakage || ""}</td>
    <td class="adjustment">${r.adjustment || ""}</td>
  `;
}

/* ─── RENDER FLAT ─── */

function renderFlat(rows, tbody) {
  rows.forEach((r, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = rowHTML(r);
    tr.style.animationDelay = `${Math.min(i * 20, 600)}ms`;
    tbody.appendChild(tr);
  });
}

/* ─── RENDER BY MATERIAL ─── */

function renderByMaterial(rows, tbody) {
  // group
  const groups = rows.reduce((acc, r) => {
    const key = r.material || "Unknown";
    if (!acc[key]) acc[key] = [];
    acc[key].push(r);
    return acc;
  }, {});

  // sort groups by count desc
  const sorted = Object.entries(groups).sort((a, b) => b[1].length - a[1].length);

  sorted.forEach(([material, items]) => {
    // count escalated in this group
    const escalatedCount = items.filter(r => is(r, "Escalated", "Breakage to Picking")).length;

    const groupRow = document.createElement("tr");
    groupRow.className = "group-row";
    groupRow.dataset.open = "false";
    groupRow.style.animationDelay = "0ms";

    const critBadge = escalatedCount > 0
      ? `<span class="group-crit">${escalatedCount} escalated</span>`
      : "";

    groupRow.innerHTML = `
      <td colspan="9">
        <span class="arrow">▶</span>
        ${material}
        <span class="group-count">(${items.length})</span>
        ${critBadge}
      </td>
    `;

    tbody.appendChild(groupRow);

    const detailRows = items.map((r, i) => {
      const tr = document.createElement("tr");
      tr.classList.add("detail-row", "hidden");
      tr.innerHTML = rowHTML(r);
      tr.style.animationDelay = `${i * 18}ms`;
      tbody.appendChild(tr);
      return tr;
    });

    groupRow.addEventListener("click", () => {
      const isOpen = groupRow.dataset.open === "true";
      groupRow.dataset.open = String(!isOpen);
      groupRow.querySelector(".arrow").style.transform = isOpen ? "" : "rotate(90deg)";
      detailRows.forEach(tr => {
        tr.classList.toggle("hidden", isOpen);
        if (!isOpen) {
          // re-trigger stagger animation when opening
          tr.style.animation = "none";
          tr.offsetHeight; // reflow
          tr.style.animation = "";
        }
      });
    });
  });
}

/* ─── RENDER ─── */

function render() {
  const tbody = document.getElementById("table-body");
  tbody.innerHTML = "";

  const rows = getFilteredRows();

  // row count
  const countEl = document.getElementById("row-count");
  if (countEl) {
    countEl.innerHTML = rows.length > 0
      ? `<span>${rows.length}</span> record${rows.length !== 1 ? "s" : ""}`
      : `<span>0</span> records`;
  }

  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.className = "empty-row";
    tr.innerHTML = `<td colspan="9">No records match the current filters.</td>`;
    tbody.appendChild(tr);
    return;
  }

  if (state.view === "material") {
    renderByMaterial(rows, tbody);
  } else {
    renderFlat(rows, tbody);
  }
}

/* ─── LOAD DATA ─── */

async function loadData() {
  if (state.loading) return;
  state.loading = true;

  const tbody = document.getElementById("table-body");
  tbody.innerHTML = `<tr class="loading-row"><td colspan="9"><span class="spinner"></span>Loading data…</td></tr>`;

  // clear KPIs
  ["k-total","k-active","k-completed","k-adjusted","k-escalated","k-test","k-cancel"]
    .forEach(id => { const el = document.getElementById(id); if (el) el.textContent = "—"; });

  try {
    const res  = await fetch(`${API}?year=${state.year}`);
    const json = await res.json();
    state.rows = json.rows || [];
  } catch (err) {
    console.error("API error:", err);
    state.rows = [];
    tbody.innerHTML = `<tr class="empty-row"><td colspan="9">Failed to load data. Check your connection.</td></tr>`;
    state.loading = false;
    return;
  }

  state.loading = false;
  updateKPIs();
  render();
}

/* ─── TAB WIRING ─── */

function wireTab(selector, stateKey) {
  document.querySelectorAll(selector).forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(selector).forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      state[stateKey] = btn.dataset[stateKey];
      if (stateKey === "year") {
        loadData();
      } else {
        render();
      }
    });
  });
}

wireTab(".tabs-year .tab",  "year");
wireTab(".tabs-mode .tab",  "mode");
wireTab(".tabs-view .tab",  "view");

/* ─── SEARCH WIRING ─── */

const searchInput = document.getElementById("search-input");
let searchTimer = null;

if (searchInput) {
  searchInput.addEventListener("input", () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      state.query = searchInput.value.trim();
      render();
    }, 180);
  });
}

/* ─── INIT ─── */

loadData();