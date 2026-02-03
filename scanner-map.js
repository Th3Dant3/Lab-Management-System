/* =====================================================
   SCANNER MAP – FULL FINAL JS (LOCKED & EXTENDED)
   ===================================================== */

const app = document.getElementById("app");
const selectedIssues = new Set();

/* ===============================
   🔹 NEW: LOCAL STORAGE (ADD ONLY)
   =============================== */

const SCANNER_STORAGE_KEY = "lms_scanner_attention";

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

/* ===============================
   DATA (UNCHANGED)
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
   CARD + SELECTION (UNCHANGED + SAVE)
   =============================== */

function card(item) {
  const el = document.createElement("div");
  el.className = `card ${item.iface}`;

  const key = `${item.name} – ${item.label} – Port ${item.port}`;

  el.innerHTML = `
    <div class="card-title">${item.name}</div>
    <div class="card-interface">Interface ${item.label}</div>
    <div class="card-port">Port ${item.port}</div>
  `;

  el.onclick = () => {
    el.classList.toggle("selected");
    el.classList.contains("selected")
      ? selectedIssues.add(key)
      : selectedIssues.delete(key);

    saveScannerState(); // 🔹 NEW (non-breaking)
  };

  return el;
}

/* ===============================
   LAYOUT HELPERS (UNCHANGED)
   =============================== */

function evenOddGrid(items) {
  const even = [], odd = [];
  items.forEach(i => {
    const n = parseInt(i.name.replace(/\D/g,""));
    (n % 2 === 0 ? even : odd).push(i);
  });

  const wrap = document.createElement("div");
  wrap.className = "two-row-grid";

  const top = document.createElement("div");
  top.className = "grid";
  even.forEach(i => top.appendChild(card(i)));

  const bottom = document.createElement("div");
  bottom.className = "grid";
  odd.forEach(i => bottom.appendChild(card(i)));

  wrap.append(top, bottom);
  return wrap;
}

function simpleGrid(items) {
  const g = document.createElement("div");
  g.className = "grid";
  items.forEach(i => g.appendChild(card(i)));
  return g;
}

/* ===============================
   SECTION BUILDER (UNCHANGED)
   =============================== */

function section(title) {
  const s = document.createElement("section");
  const h = document.createElement("div");
  const c = document.createElement("div");

  h.className = "section-header";
  h.textContent = `▼ ${title}`;
  c.className = "section-content";

  h.onclick = () => {
    const open = c.style.display !== "none";
    c.style.display = open ? "none" : "block";
    h.textContent = open ? `▶ ${title}` : `▼ ${title}`;
  };

  s.append(h, c);
  return { s, c };
}

/* ===============================
   RENDER (UNCHANGED)
   =============================== */

const surf = section("Surface Unbox");
surf.c.appendChild(simpleGrid(data.surface));
app.appendChild(surf.s);

const fin = section("Finish Department");

Object.entries(data.finish).forEach(([line, g]) => {
  const lh = document.createElement("div");
  lh.className = "line-header";
  lh.textContent = `▼ ${line}`;

  const lc = document.createElement("div");

  lc.append("Mounting", evenOddGrid(g.mounting));
  lc.append("Final Inspection", evenOddGrid(g.finalInspection));
  if (g.finishUnbox) lc.append("Finish Unbox", simpleGrid(g.finishUnbox));
  if (g.handstone) lc.append("Handstone", simpleGrid(g.handstone));

  fin.c.append(lh, lc);
});

app.appendChild(fin.s);

const arIn = section("AR Inside");
arIn.c.appendChild(simpleGrid(data.arInside));
app.appendChild(arIn.s);

const arOut = section("AR Outside");
["groupA","groupB"].forEach((g,i)=>{
  const h=document.createElement("div");
  h.className="sub-header";
  h.textContent=i===0?"Basket Group A":"Basket Group B";
  arOut.c.append(h, simpleGrid(data.arOutside[g]));
});
app.appendChild(arOut.s);

/* ===============================
   🔹 NEW: RESTORE SELECTIONS AFTER RENDER
   =============================== */

loadScannerState();

document.querySelectorAll(".card").forEach(cardEl => {
  const title = cardEl.querySelector(".card-title")?.textContent;
  const iface = cardEl.querySelector(".card-interface")?.textContent;
  const port = cardEl.querySelector(".card-port")?.textContent;

  if (!title || !iface || !port) return;

  const key = `${title} – ${iface.replace("Interface ", "")} – ${port}`;

  if (selectedIssues.has(key)) {
    cardEl.classList.add("selected");
  }
});

/* ===============================
   ACTION BUTTONS (UNCHANGED)
   =============================== */

const controls = document.createElement("div");
controls.className = "controls";

const sendBtn = document.createElement("button");
sendBtn.textContent = "Send Gmail Troubleshoot Report";
sendBtn.onclick = () => {
  if (!selectedIssues.size) return alert("Select at least one station.");
  const body = [...selectedIssues].join("\n");
  window.open(
    `https://mail.google.com/mail/?view=cm&fs=1&su=Scanner Troubleshooting&body=${encodeURIComponent(body)}`,
    "_blank"
  );
};

const backBtn = document.createElement("button");
backBtn.textContent = "Back";
backBtn.onclick = () => window.location.href = "index.html";

controls.append(sendBtn, backBtn);
document.body.appendChild(controls);
