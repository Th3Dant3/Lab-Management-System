/* =====================================================
   SCANNER MAP — GLASSMORPHISM DARK v2
   Features: pulse on critical, section count badges,
   legend bar, click-to-flag with note modal,
   stagger animation
===================================================== */

const app = document.getElementById("app");
const selectedIssues = new Set();
const issueNotes = {};
const STORAGE_KEY  = "lms_scanner_attention";
const NOTES_KEY    = "lms_scanner_notes";

/* ─── STORAGE ─── */

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify([...selectedIssues]));
  localStorage.setItem(NOTES_KEY, JSON.stringify(issueNotes));
}

function loadState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (Array.isArray(saved)) saved.forEach(k => selectedIssues.add(k));
  } catch(e) {}
  try {
    const notes = JSON.parse(localStorage.getItem(NOTES_KEY));
    if (notes && typeof notes === 'object') Object.assign(issueNotes, notes);
  } catch(e) {}
}

loadState();

/* ─── DATA ─── */

const data = {
  surface: [
    { name: "Surface 1", iface: "mi", label: "MI 1", port: 43 },
    { name: "Surface 2", iface: "mi", label: "MI 1", port: 45 }
  ],

  finish: {
    "Line A": {
      mounting: [
        { name: "Station 1",  iface: "mi", label: "MI 1", port: 16 },
        { name: "Station 2",  iface: "mi", label: "MI 1", port: 8  },
        { name: "Station 3",  iface: "mi", label: "MI 1", port: 11 },
        { name: "Station 4",  iface: "mi", label: "MI 1", port: 10 },
        { name: "Station 5",  iface: "mi", label: "MI 1", port: 15 },
        { name: "Station 6",  iface: "mi", label: "MI 1", port: 9  },
        { name: "Station 7",  iface: "mi", label: "MI 1", port: 14 },
        { name: "Station 8",  iface: "mi", label: "MI 1", port: 12 },
        { name: "Station 9",  iface: "mi", label: "MI 1", port: 17 },
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
        { name: "Station 1",  iface: "si2", label: "SI 2", port: 5  },
        { name: "Station 2",  iface: "si2", label: "SI 2", port: 6  },
        { name: "Station 3",  iface: "si2", label: "SI 2", port: 15 },
        { name: "Station 4",  iface: "si2", label: "SI 2", port: 16 },
        { name: "Station 5",  iface: "si2", label: "SI 2", port: 10 },
        { name: "Station 6",  iface: "si2", label: "SI 2", port: 9  },
        { name: "Station 7",  iface: "si2", label: "SI 2", port: 12 },
        { name: "Station 8",  iface: "si2", label: "SI 2", port: 11 },
        { name: "Station 9",  iface: "si2", label: "SI 2", port: 14 },
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
        { name: "Station 1",  iface: "si3", label: "SI 3", port: 13 },
        { name: "Station 2",  iface: "si3", label: "SI 3", port: 5  },
        { name: "Station 3",  iface: "si3", label: "SI 3", port: 14 },
        { name: "Station 4",  iface: "si3", label: "SI 3", port: 6  },
        { name: "Station 5",  iface: "si3", label: "SI 3", port: 15 },
        { name: "Station 6",  iface: "si3", label: "SI 3", port: 9  },
        { name: "Station 7",  iface: "si3", label: "SI 3", port: 16 },
        { name: "Station 8",  iface: "si3", label: "SI 3", port: 8  },
        { name: "Station 9",  iface: "si3", label: "SI 3", port: 17 },
        { name: "Station 10", iface: "si3", label: "SI 3", port: 7  }
      ],
      finalInspection: [
        { name: "Station 1", iface: "si3", label: "SI 3", port: 11 },
        { name: "Station 2", iface: "si3", label: "SI 3", port: 3  },
        { name: "Station 3", iface: "si3", label: "SI 3", port: 12 }
      ]
    }
  },

  arInside: [
    { name: "Sectoring 1", iface: "mi", label: "MI 1", port: 25 },
    { name: "Sectoring 2", iface: "mi", label: "MI 1", port: 24 },
    { name: "Sectoring 3", iface: "mi", label: "MI 1", port: 26 },
    { name: "Oven",        iface: "mi", label: "MI 1", port: 23 },
    { name: "De-Ring",     iface: "mi", label: "MI 1", port: 28 },
    { name: "Inside Lab",  iface: "mi", label: "MI 1", port: 31 },
    { name: "Inside Lab 2",iface: "mi", label: "MI 1", port: 32 }
  ],

  arOutside: {
    groupA: [
      { name: "Basket 1", iface: "si1", label: "SI 1", port: 6  },
      { name: "Basket 2", iface: "si1", label: "SI 1", port: 3  },
      { name: "Basket 3", iface: "si1", label: "SI 1", port: 5  },
      { name: "Basket 4", iface: "si1", label: "SI 1", port: 4  }
    ],
    groupB: [
      { name: "Basket 5", iface: "si1", label: "SI 1", port: 9  },
      { name: "Basket 6", iface: "si1", label: "SI 1", port: 7  },
      { name: "Basket 7", iface: "si1", label: "SI 1", port: 8  },
      { name: "Basket 8", iface: "si1", label: "SI 1", port: 10 }
    ]
  }
};

/* ─── HELPERS ─── */

let cardIndex = 0;

function getKey(item) {
  return `${item.name}__${item.label}__${item.port}`;
}

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

/* ─── NOTE MODAL ─── */

function openModal(item, cardEl) {
  const key = getKey(item);
  const existing = issueNotes[key] || "";
  const isFlagged = selectedIssues.has(key);

  const overlay = el("div", "modal-overlay");

  overlay.innerHTML = `
    <div class="modal">
      <div class="modal-header">
        <div>
          <div class="modal-title">${item.name}</div>
          <div class="modal-station">Interface ${item.label} · Port ${item.port}</div>
        </div>
        <button class="modal-close" id="modal-close-btn">✕</button>
      </div>
      <div class="modal-label">Flag note</div>
      <textarea class="modal-textarea" id="modal-note" placeholder="Describe the issue (e.g. cable loose, port unresponsive…)">${existing}</textarea>
      <div class="modal-actions">
        <button class="modal-btn flag" id="modal-flag-btn">${isFlagged ? "Update flag" : "Flag as critical"}</button>
        ${isFlagged ? `<button class="modal-btn clear" id="modal-clear-btn">Clear flag</button>` : ""}
        <button class="modal-btn cancel" id="modal-cancel-btn">Cancel</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);
  const textarea = overlay.querySelector("#modal-note");
  textarea.focus();

  function close() {
    overlay.style.opacity = "0";
    overlay.style.transition = "opacity 0.2s";
    setTimeout(() => overlay.remove(), 200);
  }

  overlay.querySelector("#modal-close-btn").onclick  = close;
  overlay.querySelector("#modal-cancel-btn").onclick = close;
  overlay.addEventListener("click", e => { if (e.target === overlay) close(); });

  overlay.querySelector("#modal-flag-btn").onclick = () => {
    const note = textarea.value.trim();
    selectedIssues.add(key);
    if (note) issueNotes[key] = note; else delete issueNotes[key];
    applyCardState(cardEl, item);
    saveState();
    updateSummary();
    updateSectionBadges();
    close();
  };

  const clearBtn = overlay.querySelector("#modal-clear-btn");
  if (clearBtn) clearBtn.onclick = () => {
    selectedIssues.delete(key);
    delete issueNotes[key];
    applyCardState(cardEl, item);
    saveState();
    updateSummary();
    updateSectionBadges();
    close();
  };
}

/* ─── CARD STATE ─── */

function applyCardState(cardEl, item) {
  const key = getKey(item);
  const isFlagged = selectedIssues.has(key);
  const note = issueNotes[key];

  cardEl.classList.remove("active", "error", "selected");
  cardEl.classList.add(isFlagged ? "error" : "active");

  let noteEl = cardEl.querySelector(".card-note-preview");
  if (isFlagged && note) {
    if (!noteEl) {
      noteEl = el("div", "card-note-preview");
      cardEl.appendChild(noteEl);
    }
    noteEl.textContent = "⚑ " + note;
  } else {
    if (noteEl) noteEl.remove();
  }
}

/* ─── CARD ─── */

function card(item) {
  const key = getKey(item);
  const idx = cardIndex++;

  const c = el("div", `card ${item.iface}`);
  c.style.animationDelay = `${idx * 28}ms`;

  const flag = el("div", "card-flag", "FLAGGED");
  const title = el("div", "card-title", item.name);
  const iface = el("div", "card-interface", `Interface ${item.label}`);
  const port  = el("div", "card-port", `Port ${item.port}`);

  c.appendChild(flag);
  c.appendChild(title);
  c.appendChild(iface);
  c.appendChild(port);

  applyCardState(c, item);

  c.addEventListener("click", e => {
    c.style.transform = "scale(0.96)";
    setTimeout(() => c.style.transform = "", 120);
    openModal(item, c);
  });

  return c;
}

/* ─── SUMMARY ─── */

function updateSummary() {
  const all = document.querySelectorAll(".card");
  let healthy = 0, critical = 0;
  all.forEach(c => {
    if (c.classList.contains("error")) critical++;
    else healthy++;
  });
  const h = document.getElementById("healthyCount");
  const w = document.getElementById("warnCount");
  const c = document.getElementById("criticalCount");
  if (h) h.textContent = healthy;
  if (w) w.textContent = 0;
  if (c) c.textContent = critical;
}

/* ─── SECTION BADGE REGISTRY ─── */

const sectionBadgeUpdaters = [];

function updateSectionBadges() {
  sectionBadgeUpdaters.forEach(fn => fn());
}

/* ─── GRID ─── */

function simpleGrid(items) {
  const g = el("div", "grid");
  items.forEach(item => g.appendChild(card(item)));
  return g;
}

/* ─── COUNT CRITICAL IN ITEM LIST ─── */

function countCritInItems(items) {
  return items.filter(item => selectedIssues.has(getKey(item))).length;
}

/* ─── SECTION ─── */

function section(title, allItems) {
  const s = document.createElement("section");
  const h = el("div", "section-header active");
  const c = el("div", "section-content");

  const left = el("div", "sh-left");
  const arrow = el("span", "sh-arrow", "▸");
  const titleSpan = el("span", "", title);
  left.appendChild(arrow);
  left.appendChild(titleSpan);

  const badges = el("div", "sh-badges");
  const totalBadge = el("span", "sh-badge total", `${allItems.length} stations`);
  const critBadge  = el("span", "sh-badge crit");
  critBadge.style.display = "none";
  badges.appendChild(totalBadge);
  badges.appendChild(critBadge);

  h.appendChild(left);
  h.appendChild(badges);

  function refreshBadge() {
    const cnt = countCritInItems(allItems);
    if (cnt > 0) {
      critBadge.textContent = `${cnt} critical`;
      critBadge.style.display = "";
    } else {
      critBadge.style.display = "none";
    }
  }
  sectionBadgeUpdaters.push(refreshBadge);
  refreshBadge();

  h.addEventListener("click", () => {
    const open = c.style.display !== "none";
    c.style.display = open ? "none" : "block";
    h.classList.toggle("active", !open);
  });

  s.appendChild(h);
  s.appendChild(c);
  return { s, c, refreshBadge };
}

/* ─── LINE SECTION ─── */

function lineSection(title, allItems) {
  const wrap = document.createElement("div");
  const h = el("div", "line-header active");
  const c = el("div", "line-content");

  const left = el("div", "sh-left");
  const arrow = el("span", "sh-arrow", "▸");
  const titleSpan = el("span", "", title);
  left.appendChild(arrow);
  left.appendChild(titleSpan);

  const badges = el("div", "sh-badges");
  const totalBadge = el("span", "sh-badge total", `${allItems.length} stations`);
  const critBadge  = el("span", "sh-badge crit");
  critBadge.style.display = "none";
  badges.appendChild(totalBadge);
  badges.appendChild(critBadge);

  h.appendChild(left);
  h.appendChild(badges);

  function refreshBadge() {
    const cnt = countCritInItems(allItems);
    if (cnt > 0) {
      critBadge.textContent = `${cnt} critical`;
      critBadge.style.display = "";
    } else {
      critBadge.style.display = "none";
    }
  }
  sectionBadgeUpdaters.push(refreshBadge);
  refreshBadge();

  h.addEventListener("click", () => {
    const open = c.style.display !== "none";
    c.style.display = open ? "none" : "block";
    h.classList.toggle("active", !open);
  });

  wrap.appendChild(h);
  wrap.appendChild(c);
  return { wrap, c, refreshBadge };
}

/* ─── BUILD UI ─── */

// Surface
const surfItems = data.surface;
const surf = section("Surface", surfItems);
surf.c.appendChild(simpleGrid(surfItems));
app.appendChild(surf.s);

// Finish
const allFinishItems = Object.values(data.finish).flatMap(g =>
  [...(g.mounting||[]), ...(g.finalInspection||[]), ...(g.finishUnbox||[]), ...(g.handstone||[])]
);
const fin = section("Finish", allFinishItems);

Object.entries(data.finish).forEach(([line, groups]) => {
  const lineAllItems = [
    ...(groups.mounting||[]),
    ...(groups.finalInspection||[]),
    ...(groups.finishUnbox||[]),
    ...(groups.handstone||[])
  ];
  const ls = lineSection(line, lineAllItems);

  if (groups.mounting) {
    ls.c.appendChild(el("div", "sub-header", "Mounting"));
    ls.c.appendChild(simpleGrid(groups.mounting));
  }
  if (groups.finalInspection) {
    ls.c.appendChild(el("div", "sub-header", "Final Inspection"));
    ls.c.appendChild(simpleGrid(groups.finalInspection));
  }
  if (groups.finishUnbox) {
    ls.c.appendChild(el("div", "sub-header", "Finish Unbox"));
    ls.c.appendChild(simpleGrid(groups.finishUnbox));
  }
  if (groups.handstone) {
    ls.c.appendChild(el("div", "sub-header", "Handstone"));
    ls.c.appendChild(simpleGrid(groups.handstone));
  }

  fin.c.appendChild(ls.wrap);
});
app.appendChild(fin.s);

// AR Inside
const arInItems = data.arInside;
const arIn = section("AR Inside", arInItems);
arIn.c.appendChild(simpleGrid(arInItems));
app.appendChild(arIn.s);

// AR Outside
const arOutItems = [...data.arOutside.groupA, ...data.arOutside.groupB];
const arOut = section("AR Outside", arOutItems);
arOut.c.appendChild(el("div", "sub-header", "Group A"));
arOut.c.appendChild(simpleGrid(data.arOutside.groupA));
arOut.c.appendChild(el("div", "sub-header", "Group B"));
arOut.c.appendChild(simpleGrid(data.arOutside.groupB));
app.appendChild(arOut.s);

/* ─── FOOTER ─── */

const footer = document.createElement("div");
footer.className = "footer";

const flagCountEl = el("div", "footer-flag-count");

function updateFooterCount() {
  const n = selectedIssues.size;
  flagCountEl.innerHTML = n > 0
    ? `<span>${n}</span> station${n > 1 ? "s" : ""} flagged`
    : "No stations flagged";
}

const sendBtn = document.createElement("button");
sendBtn.className = "send-btn";
sendBtn.textContent = "Send Report";
sendBtn.onclick = () => {
  if (!selectedIssues.size) { alert("No stations flagged."); return; }
  const lines = [...selectedIssues].map(k => {
    const note = issueNotes[k];
    return note ? `${k}\n  Note: ${note}` : k;
  });
  window.open(`https://mail.google.com/mail/?view=cm&fs=1&su=Scanner Report&body=${encodeURIComponent(lines.join("\n\n"))}`);
};

const backBtn = document.createElement("button");
backBtn.className = "back-footer-btn";
backBtn.textContent = "← Back";
backBtn.onclick = () => location.href = "index.html";

footer.appendChild(flagCountEl);
footer.appendChild(sendBtn);
footer.appendChild(backBtn);
document.body.appendChild(footer);

/* ─── INIT ─── */

updateSummary();
updateFooterCount();

// patch saveState to also update footer count
const _save = saveState;
window._updateAll = () => {
  updateSummary();
  updateSectionBadges();
  updateFooterCount();
};

// watch for changes
const origSave = saveState;
Object.defineProperty(window, '_savedOnce', { value: false, writable: true });

setInterval(() => {
  updateFooterCount();
}, 500);