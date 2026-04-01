/* =====================================================
   IN8 / IN9 — TERMINAL v2
   Features: typing log, terminal toggles,
   file info bar, live row count, blinking cursor
===================================================== */

/* ─── ELEMENTS ─── */
const dropzone    = document.getElementById("dropzone");
const fileInput   = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const previewBtn  = document.getElementById("previewBtn");
const downloadBtn = document.getElementById("downloadBtn");
const statusBar   = document.getElementById("statusBar");
const statusText  = document.getElementById("statusText");
const terminal    = document.getElementById("terminal");
const previewEl   = document.getElementById("preview");
const columnPicker= document.getElementById("columnPicker");
const fileInfo    = document.getElementById("fileInfo");
const cursor      = document.getElementById("cursor");
const logTS       = document.getElementById("logTimestamp");

let workbook = null, rows = [], headers = [];
let lastCsv = "", lastName = "output.csv";

/* ─── TIMESTAMP ─── */

function nowStr() {
  return new Date().toLocaleTimeString("en-US", { hour12: false });
}

function updateTS() {
  logTS.textContent = nowStr();
}
setInterval(updateTS, 1000);
updateTS();

/* ─── TERMINAL LOG ─── */

const logQueue = [];
let logBusy = false;

function log(msg, type = "sys", instant = false) {
  logQueue.push({ msg, type, instant });
  if (!logBusy) drainLog();
}

function drainLog() {
  if (!logQueue.length) { logBusy = false; return; }
  logBusy = true;
  const { msg, type, instant } = logQueue.shift();
  if (instant) {
    appendLogLine(msg, type);
    drainLog();
  } else {
    typeLogLine(msg, type, drainLog);
  }
}

function appendLogLine(msg, type) {
  // move cursor to new line
  const line = document.createElement("span");
  line.className = `log-line ${type}`;
  line.innerHTML = `<span class="log-prompt">▸</span> ${escapeHtml(msg)}`;
  terminal.insertBefore(line, cursor);
  terminal.scrollTop = terminal.scrollHeight;
}

function typeLogLine(msg, type, cb) {
  const line = document.createElement("span");
  line.className = `log-line ${type}`;
  line.innerHTML = `<span class="log-prompt">▸</span> `;
  terminal.insertBefore(line, cursor);

  let i = 0;
  const speed = Math.max(12, 300 / msg.length); // faster for longer msgs
  const iv = setInterval(() => {
    line.innerHTML = `<span class="log-prompt">▸</span> ${escapeHtml(msg.slice(0, i))}`;
    i++;
    terminal.scrollTop = terminal.scrollHeight;
    if (i > msg.length) {
      clearInterval(iv);
      cb();
    }
  }, speed);
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/* ─── BOOT SEQUENCE ─── */

function bootSequence() {
  const lines = [
    { msg: "system initializing…",          type: "dim"  },
    { msg: "xlsx parser loaded ✓",           type: "ok"   },
    { msg: "ready — drop a .xlsx or .xlsm file", type: "sys" },
  ];
  let delay = 400;
  lines.forEach(l => {
    setTimeout(() => log(l.msg, l.type), delay);
    delay += 600;
  });
}

bootSequence();

/* ─── STATUS BAR ─── */

function setStatus(msg, type = "ready") {
  statusBar.className = `status-bar ${type}`;
  statusText.textContent = msg;
}

/* ─── FILE INFO ─── */

function showFileInfo(file, wb, rowCount) {
  document.getElementById("fi-name").textContent   = file.name;
  document.getElementById("fi-size").textContent   = formatBytes(file.size);
  document.getElementById("fi-rows").textContent   = rowCount;
  document.getElementById("fi-cols").textContent   = headers.length;
  document.getElementById("fi-sheets").textContent = wb.SheetNames.length;
  fileInfo.classList.add("visible");
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/* ─── DRAG & DROP ─── */

dropzone.onclick = () => fileInput.click();

["dragenter", "dragover"].forEach(e =>
  dropzone.addEventListener(e, ev => { ev.preventDefault(); dropzone.classList.add("drag"); })
);
["dragleave", "drop"].forEach(e =>
  dropzone.addEventListener(e, ev => { ev.preventDefault(); dropzone.classList.remove("drag"); })
);
dropzone.addEventListener("drop", e => {
  fileInput.files = e.dataTransfer.files;
  loadFile();
});
fileInput.addEventListener("change", loadFile);

/* ─── LOAD FILE ─── */

async function loadFile() {
  const file = fileInput.files[0];
  if (!file) return;

  log(`loading file: ${file.name}`, "info");
  log(`size: ${formatBytes(file.size)}`, "dim", true);
  setStatus("reading file…", "ready");

  try {
    const buf = await file.arrayBuffer();
    workbook = XLSX.read(buf, { type: "array" });

    // populate sheet dropdown
    sheetSelect.innerHTML = "";
    workbook.SheetNames.forEach(s => {
      const o = document.createElement("option");
      o.value = s; o.textContent = s;
      sheetSelect.appendChild(o);
    });
    sheetSelect.disabled = false;
    previewBtn.disabled  = false;

    loadSheet(file);
    log(`parsed ${workbook.SheetNames.length} sheet(s) ✓`, "ok");
    setStatus(`loaded — ${file.name}`, "ok");

  } catch (e) {
    log(`error: failed to parse file`, "err");
    log(String(e.message), "err", true);
    setStatus("error reading file", "err");
  }
}

sheetSelect.onchange = () => loadSheet();

function loadSheet(file) {
  const ws = workbook.Sheets[sheetSelect.value];
  rows    = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
  headers = rows[0] || [];

  // render column toggles
  columnPicker.innerHTML = "";
  headers.forEach((h, i) => {
    const label = document.createElement("label");
    label.className = "col-toggle";
    label.innerHTML = `
      <input type="checkbox" checked value="${i}" />
      <span class="col-toggle-dot"></span>
      ${h || `(col ${i})`}
    `;
    columnPicker.appendChild(label);
  });

  const dataRows = rows.length - 1;
  log(`sheet "${sheetSelect.value}" → ${dataRows} rows, ${headers.length} cols`, "info");

  if (file) showFileInfo(file, workbook, dataRows);
  else {
    document.getElementById("fi-rows").textContent = dataRows;
    document.getElementById("fi-cols").textContent = headers.length;
  }
}

/* ─── CSV ─── */

function escapeCsv(v) {
  v = (v ?? "").toString();
  if (v.includes(",") || v.includes('"') || v.includes("\n"))
    return `"${v.replace(/"/g, '""')}"`;
  return v;
}

previewBtn.onclick = () => {
  const idxs = [...columnPicker.querySelectorAll("input:checked")].map(i => +i.value);
  const fmt  = document.querySelector("input[name=fmt]:checked").value;

  if (!idxs.length) {
    log("no columns selected — check at least one column", "warn");
    setStatus("no columns selected", "ready");
    return;
  }

  log(`generating ${fmt} CSV with ${idxs.length} columns…`, "sys");

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = idxs.map(j => escapeCsv(rows[i][j]));
    if (r.join("").trim()) out.push(r.join(","));
  }

  if (!out.length) {
    log("CSV empty — no data rows found", "warn");
    setStatus("empty result", "ready");
    return;
  }

  lastCsv  = out.join("\n");
  lastName = `${fmt}_output.csv`;

  // show preview
  const lines = lastCsv.split("\n");
  previewEl.textContent = lines.slice(0, 20).join("\n");
  if (lines.length > 20) {
    previewEl.textContent += `\n\n… and ${lines.length - 20} more rows`;
  }

  document.getElementById("previewMeta").textContent =
    `${lines.length} rows · ${fmt} format`;

  downloadBtn.disabled = false;
  log(`${fmt} CSV ready — ${out.length} rows generated ✓`, "ok");
  setStatus(`${fmt} ready — ${out.length} rows`, "ok");
};

downloadBtn.onclick = () => {
  const b = new Blob([lastCsv], { type: "text/csv" });
  const a = document.createElement("a");
  a.href     = URL.createObjectURL(b);
  a.download = lastName;
  a.click();
  log(`downloading: ${lastName}`, "ok");
};