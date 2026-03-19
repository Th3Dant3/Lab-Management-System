/* =====================================================
   SCANNER MAP – FINAL PRODUCTION VERSION
   ===================================================== */

const app = document.getElementById("app");
const selectedIssues = new Set();

const SCANNER_STORAGE_KEY = "lms_scanner_attention";

/* ===============================
   LOCAL STORAGE
=============================== */

function saveScannerState() {
  localStorage.setItem(
    SCANNER_STORAGE_KEY,
    JSON.stringify([...selectedIssues])
  );
}

function loadScannerState() {
  try {
    const saved = JSON.parse(localStorage.getItem(SCANNER_STORAGE_KEY));
    if (Array.isArray(saved)) {
      saved.forEach(k => selectedIssues.add(k));
    }
  } catch (e) {
    console.warn("Scanner state restore failed", e);
  }
}

loadScannerState();

/* ===============================
   DATA
=============================== */

const data = {
  surface: [
    { name: "Surface 1", iface: "mi", label: "MI 1", port: 43 },
    { name: "Surface 2", iface: "mi", label: "MI 1", port: 45 }
  ],

  finish: {
    "Line A": {
      mounting: [
        { name: "Station 1", iface: "mi", label: "MI 1", port: 16 },
        { name: "Station 2", iface: "mi", label: "MI 1", port: 8 },
        { name: "Station 3", iface: "mi", label: "MI 1", port: 11 },
        { name: "Station 4", iface: "mi", label: "MI 1", port: 10 },
        { name: "Station 5", iface: "mi", label: "MI 1", port: 15 },
        { name: "Station 6", iface: "mi", label: "MI 1", port: 9 },
        { name: "Station 7", iface: "mi", label: "MI 1", port: 14 },
        { name: "Station 8", iface: "mi", label: "MI 1", port: 12 },
        { name: "Station 9", iface: "mi", label: "MI 1", port: 17 },
        { name: "Station 10", iface: "mi", label: "MI 1", port: 13 }
      ],
      finalInspection: [
        { name: "Station 1", iface: "mi", label: "MI 1", port: 3 },
        { name: "Station 2", iface: "mi", label: "MI 1", port: 7 },
        { name: "Station 3", iface: "mi", label: "MI 1", port: 5 }
      ],
      finishUnbox: [
        { name: "Finish 1", iface: "mi", label: "MI 1", port: 41 },
        { name: "Finish 2", iface: "mi", label: "MI 1", port: 47 }
      ],
      handstone: [
        { name: "Handstone 1", iface: "mi", label: "MI 1", port: 44 }
      ]
    },

    "Line B": {
      mounting: [
        { name: "Station 1", iface: "si2", label: "SI 2", port: 5 },
        { name: "Station 2", iface: "si2", label: "SI 2", port: 6 },
        { name: "Station 3", iface: "si2", label: "SI 2", port: 15 },
        { name: "Station 4", iface: "si2", label: "SI 2", port: 16 },
        { name: "Station 5", iface: "si2", label: "SI 2", port: 10 },
        { name: "Station 6", iface: "si2", label: "SI 2", port: 9 },
        { name: "Station 7", iface: "si2", label: "SI 2", port: 12 },
        { name: "Station 8", iface: "si2", label: "SI 2", port: 11 },
        { name: "Station 9", iface: "si2", label: "SI 2", port: 14 },
        { name: "Station 10", iface: "si2", label: "SI 2", port: 13 }
      ],
      finalInspection: [
        { name: "Station 1", iface: "si2", label: "SI 2", port: 7 },
        { name: "Station 2", iface: "si2", label: "SI 2", port: 8 },
        { name: "Station 3", iface: "si2", label: "SI 2", port: 3 }
      ]
    },

    "Line C": {
      mounting: [
        { name: "Station 1", iface: "si3", label: "SI 3", port: 13 },
        { name: "Station 2", iface: "si3", label: "SI 3", port: 5 },
        { name: "Station 3", iface: "si3", label: "SI 3", port: 14 },
        { name: "Station 4", iface: "si3", label: "SI 3", port: 6 },
        { name: "Station 5", iface: "si3", label: "SI 3", port: 15 },
        { name: "Station 6", iface: "si3", label: "SI 3", port: 9 },
        { name: "Station 7", iface: "si3", label: "SI 3", port: 16 },
        { name: "Station 8", iface: "si3", label: "SI 3", port: 8 },
        { name: "Station 9", iface: "si3", label: "SI 3", port: 17 },
        { name: "Station 10", iface: "si3", label: "SI 3", port: 7 }
      ],
      finalInspection: [
        { name: "Station 1", iface: "si3", label: "SI 3", port: 11 },
        { name: "Station 2", iface: "si3", label: "SI 3", port: 3 },
        { name: "Station 3", iface: "si3", label: "SI 3", port: 12 }
      ]
    }
  },

  arInside: [
    { name: "Sectoring 1", iface: "mi", label: "MI 1", port: 25 },
    { name: "Sectoring 2", iface: "mi", label: "MI 1", port: 24 },
    { name: "Sectoring 3", iface: "mi", label: "MI 1", port: 26 },
    { name: "Oven", iface: "mi", label: "MI 1", port: 23 },
    { name: "De-Ring", iface: "mi", label: "MI 1", port: 28 },
    { name: "Inside Lab", iface: "mi", label: "MI 1", port: 31 },
    { name: "Inside Lab 2", iface: "mi", label: "MI 1", port: 32 }
  ],

  arOutside: {
    groupA: [
      { name: "Basket 1", iface: "si1", label: "SI 1", port: 6 },
      { name: "Basket 2", iface: "si1", label: "SI 1", port: 3 },
      { name: "Basket 3", iface: "si1", label: "SI 1", port: 5 },
      { name: "Basket 4", iface: "si1", label: "SI 1", port: 4 }
    ],
    groupB: [
      { name: "Basket 5", iface: "si1", label: "SI 1", port: 9 },
      { name: "Basket 6", iface: "si1", label: "SI 1", port: 7 },
      { name: "Basket 7", iface: "si1", label: "SI 1", port: 8 },
      { name: "Basket 8", iface: "si1", label: "SI 1", port: 10 }
    ]
  }
};

/* ===============================
   HELPERS
=============================== */

function createTextBlock(text, className = "") {
  const el = document.createElement("div");
  if (className) el.className = className;
  el.textContent = text;
  return el;
}

function getIssueKey(item) {
  return `${item.name} – ${item.label} – Port ${item.port}`;
}

/* ===============================
   CARD
=============================== */

function card(item) {
  const el = document.createElement("div");
  const key = getIssueKey(item);

  el.className = `card ${item.iface}`;

  if (selectedIssues.has(key)) {
    el.classList.add("selected", "error");
  } else {
    el.classList.add("active");
  }

  el.appendChild(createTextBlock(item.name, "card-title"));
  el.appendChild(createTextBlock(`Interface ${item.label}`, "card-interface"));
  el.appendChild(createTextBlock(`Port ${item.port}`, "card-port"));

  el.addEventListener("click", () => {

    el.style.transform = "scale(0.97)";
    setTimeout(() => el.style.transform = "", 120);

    const isSelected = el.classList.toggle("selected");

    el.classList.remove("active", "warning", "error");

    if (isSelected) {
      selectedIssues.add(key);
      el.classList.add("error");
    } else {
      selectedIssues.delete(key);
      el.classList.add("active");
    }

    saveScannerState();
    updateSummaryCounts();
  });

  return el;
}

/* ===============================
   SUMMARY COUNTS
=============================== */

function updateSummaryCounts() {

  const allCards = document.querySelectorAll(".card");
  if (!allCards.length) return;

  let healthy = 0;
  let warning = 0;
  let critical = 0;

  allCards.forEach(card => {

    if (card.classList.contains("error")) {
      critical++;
    }
    else if (card.classList.contains("warning")) {
      warning++;
    }
    else {
      healthy++;
    }

  });

  const h = document.getElementById("healthyCount");
  const w = document.getElementById("warnCount");
  const c = document.getElementById("criticalCount");

  if (h) h.textContent = healthy;
  if (w) w.textContent = warning;
  if (c) c.textContent = critical;
}

/* ===============================
   RENDER (same as yours)
=============================== */

function simpleGrid(items) {
  const g = document.createElement("div");
  g.className = "grid";
  items.forEach(item => g.appendChild(card(item)));
  return g;
}

function section(title) {
  const s = document.createElement("section");
  const h = document.createElement("div");
  const c = document.createElement("div");

  h.className = "section-header active";
  h.textContent = title;

  c.className = "section-content";

  h.addEventListener("click", () => {
    const open = c.style.display !== "none";
    c.style.display = open ? "none" : "block";
    h.classList.toggle("active", !open);
  });

  s.appendChild(h);
  s.appendChild(c);

  return { s, c };
}

function createLineSection(title) {
  const wrap = document.createElement("div");
  const h = document.createElement("div");
  const c = document.createElement("div");

  h.className = "line-header active";
  h.textContent = title;

  c.className = "line-content";

  h.addEventListener("click", () => {
    const open = c.style.display !== "none";
    c.style.display = open ? "none" : "block";
    h.classList.toggle("active", !open);
  });

  wrap.appendChild(h);
  wrap.appendChild(c);

  return { wrap, c };
}

/* ===============================
   BUILD UI
=============================== */

const surf = section("Surface");
surf.c.appendChild(simpleGrid(data.surface));
app.appendChild(surf.s);

const fin = section("Finish");

Object.entries(data.finish).forEach(([line, groups]) => {
  const lineSection = createLineSection(line);

  lineSection.c.appendChild(createTextBlock("Mounting", "sub-header"));
  lineSection.c.appendChild(simpleGrid(groups.mounting));

  lineSection.c.appendChild(createTextBlock("Final Inspection", "sub-header"));
  lineSection.c.appendChild(simpleGrid(groups.finalInspection));

  fin.c.appendChild(lineSection.wrap);
});

app.appendChild(fin.s);

const arIn = section("AR Inside");
arIn.c.appendChild(simpleGrid(data.arInside));
app.appendChild(arIn.s);

const arOut = section("AR Outside");
["groupA","groupB"].forEach(g=>{
  arOut.c.appendChild(simpleGrid(data.arOutside[g]));
});
app.appendChild(arOut.s);

/* 🔥 FIX: update counts AFTER render */
updateSummaryCounts();

/* ===============================
   FOOTER
=============================== */

const footer = document.createElement("div");
footer.className = "footer";

const sendBtn = document.createElement("button");
sendBtn.textContent = "Send Report";
sendBtn.onclick = () => {
  if (!selectedIssues.size) return alert("Select stations");
  window.open(`https://mail.google.com/mail/?view=cm&fs=1&su=Scanner Report&body=${encodeURIComponent([...selectedIssues].join("\n"))}`);
};

const backBtn = document.createElement("button");
backBtn.textContent = "Back";
backBtn.onclick = () => location.href="index.html";

footer.appendChild(sendBtn);
footer.appendChild(backBtn);
document.body.appendChild(footer);