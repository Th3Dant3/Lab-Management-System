const API =
  "https://script.google.com/macros/s/AKfycbzuUpwsa2VTGjxeSP3wVV-x2Z4yMkGH_eYS86VpFH0CzC2Ii5S5U2ag-S5d89RZ8oHJ/exec";

let state = {
  year: "2025",
  mode: "production",
  view: "material",
  rows: []
};

/* =============================
   TAB HANDLERS
============================= */
document.querySelectorAll(".tab").forEach(btn => {
  btn.addEventListener("click", () => {
    const group = btn.parentElement;
    group.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
    btn.classList.add("active");

    if (btn.dataset.year) state.year = btn.dataset.year;
    if (btn.dataset.mode) state.mode = btn.dataset.mode;
    if (btn.dataset.view) state.view = btn.dataset.view;

    loadData();
  });
});

/* =============================
   LOAD DATA
============================= */
async function loadData() {
  const res = await fetch(`${API}?year=${state.year}`);
  const json = await res.json();
  state.rows = json.rows || [];
  render();
}

/* =============================
   RENDER
============================= */
function render() {
  const tbody = document.getElementById("table-body");
  tbody.innerHTML = "";

  let rows = [...state.rows];

  // Mode filter
  if (state.mode === "production") {
    rows = rows.filter(r => !["TEST", "CANCEL"].includes(r.status));
  }
  if (state.mode === "completed") {
    rows = rows.filter(r => r.status === "Shipped");
  }
  if (state.mode === "test") {
    rows = rows.filter(r => ["TEST", "CANCEL"].includes(r.status));
  }

  updateKPIs();

  if (state.view === "material") {
    renderByMaterial(rows, tbody);
  } else {
    rows.forEach(r => {
      const tr = document.createElement("tr");
      tr.innerHTML = rowHTML(r);
      tbody.appendChild(tr);
    });
  }
}

/* =============================
   BY MATERIAL (FIXED)
============================= */
function renderByMaterial(rows, tbody) {
  const groups = {};

  rows.forEach(r => {
    if (!groups[r.material]) groups[r.material] = [];
    groups[r.material].push(r);
  });

  Object.entries(groups).forEach(([material, items]) => {
    const groupRow = document.createElement("tr");
    groupRow.className = "group-row";
    groupRow.dataset.open = "false";

    groupRow.innerHTML = `
      <td colspan="9">
        <span class="arrow">▶</span> ${material} (${items.length})
      </td>
    `;

    tbody.appendChild(groupRow);

    const detailRows = items.map(r => {
      const tr = document.createElement("tr");
      tr.classList.add("detail-row", "hidden");
      tr.innerHTML = rowHTML(r);
      tbody.appendChild(tr);
      return tr;
    });

    // ✅ CLICK TO TOGGLE (THIS WAS MISSING / BROKEN)
    groupRow.addEventListener("click", () => {
      const isOpen = groupRow.dataset.open === "true";
      groupRow.dataset.open = String(!isOpen);

      groupRow.querySelector(".arrow").textContent = isOpen ? "▶" : "▼";

      detailRows.forEach(tr => {
        tr.classList.toggle("hidden", isOpen);
      });
    });
  });
}

/* =============================
   ROW TEMPLATE
============================= */
function rowHTML(r) {
  return `
    <td>${r.rx || ""}</td>
    <td>${r.sku || ""}</td>
    <td>${r.tcLms || ""}</td>
    <td>${r.tcAfter || ""}</td>
    <td>${r.material || ""}</td>
    <td>${r.status || ""}</td>
    <td>${r.dateShipped || ""}</td>
    <td>${r.breakage || ""}</td>
    <td class="adjustment">${r.adjustment || ""}</td>
  `;
}

/* =============================
   KPI LOGIC (UNCHANGED)
============================= */
function updateKPIs() {
  const all = state.rows;

  document.getElementById("k-total").textContent = all.length;
  document.getElementById("k-test").textContent =
    all.filter(r => r.status === "TEST").length;
  document.getElementById("k-cancel").textContent =
    all.filter(r => r.status === "CANCEL").length;
  document.getElementById("k-completed").textContent =
    all.filter(r => r.status === "Shipped").length;
  document.getElementById("k-active").textContent =
    all.filter(r => !["Shipped", "TEST", "CANCEL"].includes(r.status)).length;
  document.getElementById("k-adjusted").textContent =
    all.filter(r => r.adjustment).length;
  document.getElementById("k-escalated").textContent =
    all.filter(r => r.status === "Escalated").length;
}

/* =============================
   INIT
============================= */
loadData();
