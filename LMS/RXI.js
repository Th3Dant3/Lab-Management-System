const API_URL = "https://script.google.com/a/macros/zennioptical.com/s/AKfycbx4XnbRZE3iUIjSKdzgMgsn4ug-1dce_NFEBvQUtC3ZY-pqFAzYyc1F5bNbtBSV4rYv/exec";

let RXI_ROWS = [];

document.addEventListener("DOMContentLoaded", () => {
  bootGamingHud();
  loadTemplate();
});

async function loadTemplate() {
  setStatus("Loading...", "");
  setText("sideTemplateStatus", "Loading");
  showResult("Loading RXI template...");

  try {
    const res = await apiCall("getRXITemplateData");
    RXI_ROWS = Array.isArray(res.rows) ? res.rows : [];
    renderRows();
    syncQuickFields(true);
    updateAuditStrip();
    setStatus("Ready", "success");
    setText("sideTemplateStatus", "Ready");
    showResult("Template loaded.");
  } catch (err) {
    setStatus("Error", "danger");
    setText("sideTemplateStatus", "Error");
    showResult("Load failed: " + getErrorMessage(err), true);
  }
}

function renderRows() {
  const body = document.getElementById("fieldBody");
  if (!body) return;

  if (!RXI_ROWS.length) {
    body.innerHTML = `
      <tr>
        <td colspan="2" class="loading-cell">No RXI fields found.</td>
      </tr>
    `;
    return;
  }

  body.innerHTML = RXI_ROWS.map((row, index) => {
    return `
      <tr>
        <td class="field-name">${escapeHtml(row.field)}</td>
        <td>
          <input
            data-index="${index}"
            value="${escapeAttribute(row.value || "")}"
            oninput="updateRowValue(this)"
          >
        </td>
      </tr>
    `;
  }).join("");
}

function updateRowValue(input) {
  const index = Number(input.dataset.index);
  if (!RXI_ROWS[index]) return;

  RXI_ROWS[index].value = input.value;
  syncQuickFields(false);
  updateAuditStrip();
}

function syncQuickFields(overwrite) {
  const start = RXI_ROWS.find(row => row.field === "StartIDN");
  const count = RXI_ROWS.find(row => row.field === "JobCount");
  const cpo = RXI_ROWS.find(row => row.field === "CPO");

  if (overwrite) {
    const startInput = document.getElementById("quickStartIDN");
    const countInput = document.getElementById("quickJobCount");
    const cpoInput = document.getElementById("quickCPO");

    if (startInput) startInput.value = start ? start.value : "";
    if (countInput) countInput.value = count ? count.value : "";
    if (cpoInput) cpoInput.value = cpo ? cpo.value : "";
  }
}

function applyQuickValues() {
  const startInput = document.getElementById("quickStartIDN");
  const countInput = document.getElementById("quickJobCount");
  const cpoInput = document.getElementById("quickCPO");

  const startIDN = startInput ? startInput.value.trim() : "";
  const jobCount = countInput ? countInput.value.trim() : "";
  const cpo = cpoInput ? cpoInput.value.trim() : "";

  setValueByField("StartIDN", startIDN);
  setValueByField("JobCount", jobCount);
  setValueByField("IDN", startIDN);
  setValueByField("USR orderNumber", startIDN);
  setValueByField("RXN", startIDN);
  setValueByField("SPX", startIDN + " | 01-01");
  setValueByField("CPO", cpo);

  renderRows();
  updateAuditStrip();
  showResult("Quick values applied. Preview before running.");
  setText("sideOutputStatus", "Quick Set");
}

function setValueByField(field, value) {
  const row = RXI_ROWS.find(item => item.field === field);
  if (row) row.value = value;
}

function getValueByField(field) {
  const row = RXI_ROWS.find(item => item.field === field);
  return row ? String(row.value || "").trim() : "";
}

function getPayload() {
  return {
    rows: RXI_ROWS.map(row => ({
      field: row.field,
      value: row.value
    }))
  };
}

function validateBeforeRun() {
  const startIDN = getValueByField("StartIDN");
  const jobCount = Number(getValueByField("JobCount"));
  const cpo = getValueByField("CPO");

  if (!startIDN) {
    showResult("Missing StartIDN.", true);
    return false;
  }

  if (!String(startIDN).match(/^\d+.*$/)) {
    showResult("StartIDN must begin with numbers. Example: 87DOE8.", true);
    return false;
  }

  if (!jobCount || jobCount < 1) {
    showResult("JobCount must be 1 or higher.", true);
    return false;
  }

  if (jobCount > 250) {
    showResult("JobCount too high. Keep it 250 or less per run.", true);
    return false;
  }

  if (!cpo) {
    showResult("Missing CPO.", true);
    return false;
  }

  return true;
}

async function saveTemplate() {
  if (!RXI_ROWS.length) {
    showResult("Nothing to save. Template is empty.", true);
    return;
  }

  setStatus("Saving...", "");
  showResult("Saving template...");

  try {
    const res = await apiCall("saveRXITemplateData", getPayload());
    setStatus("Saved", "success");
    setText("sideTemplateStatus", "Saved");
    showResult((res.message || "Template saved successfully.") + " Rows saved: " + (res.totalRows || RXI_ROWS.length));
  } catch (err) {
    setStatus("Error", "danger");
    showResult("Save failed: " + getErrorMessage(err), true);
  }
}

async function previewRXI() {
  if (!validateBeforeRun()) return;

  setStatus("Previewing...", "");
  showResult("Building preview...");

  try {
    const res = await apiCall("previewRXIFromWeb", getPayload());
    setStatus("Preview Ready", "success");

    const previewBox = document.getElementById("previewBox");
    if (previewBox) previewBox.textContent = res.preview || "";

    showResult("Preview created. Review it before running.");
    setText("sideOutputStatus", "Preview Ready");
  } catch (err) {
    setStatus("Error", "danger");
    showResult("Preview failed: " + getErrorMessage(err), true);
  }
}

async function generateRXI() {
  if (!validateBeforeRun()) return;

  const jobCount = getValueByField("JobCount");
  const startIDN = getValueByField("StartIDN");
  const cpo = getValueByField("CPO");

  const confirmRun = confirm(
    "Generate " + jobCount + " RXI file(s)?\n\n" +
    "StartIDN: " + startIDN + "\n" +
    "CPO: " + cpo + "\n\n" +
    "Preview first if you have not checked the RX values."
  );

  if (!confirmRun) return;

  setStatus("Running...", "");
  showResult("Generating RXI files... Do not close this page.");

  try {
    const res = await apiCall("generateRXIFromWeb", getPayload());
    setStatus("Complete", "success");
    setText("sideOutputStatus", "Complete");

    let html = `
      <div class="success"><strong>${escapeHtml(res.message || "RXI files created successfully.")}</strong></div>
      <div>Total created: ${escapeHtml(res.totalCreated || 0)}</div>
      <br>
    `;

    if (res.files && res.files.length) {
      html += res.files.map(file => {
        return `
          <div>
            <strong>${escapeHtml(file.fileName)}</strong><br>
            IDN: ${escapeHtml(file.id)} | CPO: ${escapeHtml(file.cpo)}<br>
            <a href="${escapeAttribute(file.url)}" target="_blank">Open file</a>
          </div>
          <br>
        `;
      }).join("");
    }

    const resultBox = document.getElementById("resultBox");
    if (resultBox) resultBox.innerHTML = html;
  } catch (err) {
    setStatus("Error", "danger");
    showResult("Generate failed: " + getErrorMessage(err), true);
  }
}

function copyPreview() {
  const previewBox = document.getElementById("previewBox");
  if (!previewBox) return;

  const text = previewBox.textContent || "";

  if (!text || text === "Preview will show here.") {
    showResult("Nothing to copy. Generate a preview first.", true);
    return;
  }

  navigator.clipboard.writeText(text)
    .then(() => showResult("Preview copied."))
    .catch(() => showResult("Copy failed. Select the preview manually and copy.", true));
}

function apiCall(action, payload) {
  return new Promise((resolve, reject) => {
    const callbackName = "rxiCallback_" + Date.now() + "_" + Math.floor(Math.random() * 100000);
    const url = new URL(API_URL);

    url.searchParams.set("action", action);
    url.searchParams.set("callback", callbackName);

    if (payload !== undefined) {
      url.searchParams.set("payload", JSON.stringify(payload));
    }

    const script = document.createElement("script");
    script.src = url.toString();

    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error("API request timed out."));
    }, 30000);

    window[callbackName] = function (data) {
      cleanup();

      if (!data || data.ok === false) {
        reject(new Error(data && data.error ? data.error : "API returned an error."));
        return;
      }

      resolve(data);
    };

    script.onerror = function () {
      cleanup();
      reject(new Error("API script request failed. Confirm the Apps Script web app is deployed for access."));
    };

    function cleanup() {
      clearTimeout(timeout);
      if (script.parentNode) script.parentNode.removeChild(script);
      try {
        delete window[callbackName];
      } catch (err) {
        window[callbackName] = undefined;
      }
    }

    document.body.appendChild(script);
  });
}

function updateAuditStrip() {
  const startIDN = getValueByField("StartIDN");
  const jobCountRaw = getValueByField("JobCount");
  const jobCount = Number(jobCountRaw);
  const cpo = getValueByField("CPO");

  let lastId = "-";

  if (startIDN && String(startIDN).match(/^(\d+)(.*)$/) && jobCount && jobCount > 0) {
    const match = String(startIDN).match(/^(\d+)(.*)$/);
    const numberPart = parseInt(match[1], 10);
    const suffixPart = match[2] || "";
    lastId = String(numberPart + jobCount - 1) + suffixPart;
  }

  setText("auditStart", startIDN || "-");
  setText("auditJobs", jobCountRaw || "-");
  setText("auditCPO", cpo || "-");
  setText("auditLast", lastId);
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function setStatus(text, type) {
  const pill = document.getElementById("statusPill");
  if (!pill) return;

  pill.textContent = text;
  pill.classList.remove("success", "danger");

  if (type === "success") pill.classList.add("success");
  if (type === "danger") pill.classList.add("danger");
}

function showResult(message, isError) {
  const box = document.getElementById("resultBox");
  if (!box) return;

  box.innerHTML = isError
    ? `<span class="danger">${escapeHtml(message)}</span>`
    : escapeHtml(message);
}

function getErrorMessage(err) {
  if (!err) return "Unknown error.";
  if (typeof err === "string") return err;
  if (err.message) return err.message;

  try {
    return JSON.stringify(err);
  } catch (error) {
    return "Unknown error.";
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value);
}

function bootGamingHud() {
  startHudClock();
  startParticleField();
}

function startHudClock() {
  const clock = document.getElementById("hudClock");
  if (!clock) return;

  const tick = () => {
    const now = new Date();
    clock.textContent = now.toLocaleTimeString([], {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
  };

  tick();
  window.setInterval(tick, 1000);
}

function startParticleField() {
  const canvas = document.getElementById("particleCanvas");
  if (!canvas || !canvas.getContext) return;

  const ctx = canvas.getContext("2d");
  const particles = [];
  const maxParticles = 46;

  function resize() {
    canvas.width = window.innerWidth || document.documentElement.clientWidth || 1200;
    canvas.height = window.innerHeight || document.documentElement.clientHeight || 800;
  }

  function resetParticle(particle) {
    particle.x = Math.random() * canvas.width;
    particle.y = canvas.height + Math.random() * 120;
    particle.size = 1 + Math.random() * 2.4;
    particle.speed = 0.25 + Math.random() * 0.85;
    particle.drift = -0.35 + Math.random() * 0.7;
    particle.alpha = 0.12 + Math.random() * 0.34;
  }

  resize();
  window.addEventListener("resize", resize);

  for (let i = 0; i < maxParticles; i++) {
    const particle = {};
    resetParticle(particle);
    particle.y = Math.random() * canvas.height;
    particles.push(particle);
  }

  function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    particles.forEach(particle => {
      particle.y -= particle.speed;
      particle.x += particle.drift;

      if (particle.y < -20 || particle.x < -20 || particle.x > canvas.width + 20) {
        resetParticle(particle);
      }

      const gradient = ctx.createRadialGradient(
        particle.x,
        particle.y,
        0,
        particle.x,
        particle.y,
        particle.size * 5
      );
      gradient.addColorStop(0, "rgba(78, 208, 255, " + particle.alpha + ")");
      gradient.addColorStop(0.45, "rgba(175, 74, 255, " + (particle.alpha * 0.55) + ")");
      gradient.addColorStop(1, "rgba(78, 208, 255, 0)");

      ctx.fillStyle = gradient;
      ctx.beginPath();
      ctx.arc(particle.x, particle.y, particle.size * 5, 0, Math.PI * 2);
      ctx.fill();
    });

    window.requestAnimationFrame(draw);
  }

  draw();
}
