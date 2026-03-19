document.addEventListener("DOMContentLoaded", () => {

  const API_URL =
    "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

  const els = {
    dateFilter: document.getElementById("dateFilter"),
    btnAll: document.getElementById("btnAll"),

    tabReasons: document.getElementById("tabReasons"),
    tabSentBack: document.getElementById("tabSentBack"),
    tabDelay: document.getElementById("tabDelay"),

    reasonsView: document.getElementById("reasonsView"),
    sentBackView: document.getElementById("sentBackView"),
    delayView: document.getElementById("delayView"),

    reasonBubbles: document.getElementById("reasonBubbles"),
    sentBackBubbles: document.getElementById("sentBackBubbles"),

    delayTableBody: document.getElementById("delayTableBody"),
    delayByDept: document.getElementById("delayByDept"),

    bucket12: document.getElementById("bucket-1-2"),
    bucket35: document.getElementById("bucket-3-5"),
    bucket5p: document.getElementById("bucket-5plus"),

    shiftStatsWrap: document.getElementById("shiftStatsWrap")
  };

  let activeTab = "reasons";
  let delayFilter = "all";
  let lastData = null;
  let currentYear = "all";

  /***********************
   HELPERS
  ***********************/
  const safeNumber = v => Number.isFinite(Number(v)) ? Number(v) : 0;
  const formatInt = v => safeNumber(v).toLocaleString("en-US");
  const formatPct = v => `${safeNumber(v).toFixed(0)}%`;
  const formatPct1 = v => `${safeNumber(v).toFixed(1)}%`;
  const formatDec = (v, d = 2) => safeNumber(v).toFixed(d);

  const escapeHtml = str =>
    String(str ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");

  const getFilterValue = () =>
    els.dateFilter?.value?.trim() || "all";

  function set(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  }

  /***********************
   LOAD
  ***********************/
  async function load() {
    const url = `${API_URL}?date=${encodeURIComponent(getFilterValue())}&year=${currentYear}`;

    try {
      const res = await fetch(url, { cache: "no-store" });
      const data = await res.json();

      lastData = data;
      render(data);
    } catch (e) {
      console.error("LOAD ERROR:", e);
    }
  }

  /***********************
   RENDER
  ***********************/
  function render(d) {
    const total = safeNumber(d.total);
    const evaluatedBase = safeNumber(
      getFilterValue() === "all" ? d.evaluated : d.evaluatedToday
    );

    set("activeHolds", formatInt(d.active));
    set("evaluatedCount", formatInt(d.evaluated));
    set("coveragePct", formatPct(d.coverage));
    set("lastUpdated", d.lastUpdated || "--");

    set(
      "avgInvestigationTime",
      `${formatDec(d.avgInvestigationDays)} Days (${formatDec(d.avgInvestigationHours)} Hours)`
    );

    set("arrivedToday", formatInt(d.arrived));
    set("evaluatedToday", formatInt(d.evaluatedToday));
    set("priorDayIntake", formatInt(d.priorDayIntake));
    set("sameDayPercent", formatPct1(d.sameDayPercent));
	// 🔥 ADD THIS BLOCK RIGHT HERE
const sameDayEl = document.getElementById("sameDayPercent");

if (sameDayEl) {
  const val = parseFloat(d.sameDayPercent);

  sameDayEl.classList.remove("kpi-good", "kpi-warning", "kpi-bad");

  if (val >= 80) sameDayEl.classList.add("kpi-good");
  else if (val >= 50) sameDayEl.classList.add("kpi-warning");
  else sameDayEl.classList.add("kpi-bad");

  // 🔥 ADD THIS PART RIGHT HERE
  const parentCard = sameDayEl.closest(".stat");

  if (parentCard) {
    parentCard.classList.remove("kpi-good", "kpi-warning", "kpi-bad");

    if (val >= 80) parentCard.classList.add("kpi-good");
    else if (val >= 50) parentCard.classList.add("kpi-warning");
    else parentCard.classList.add("kpi-bad");
  }
}
	
    set("avgArrivalDelayDays", `${formatDec(d.avgArrivalDelayDays)} Days`);
    set("oldestInvestigation", d.oldestInvestigation || "--");
    set("latestInvestigation", d.latestInvestigation || "--");

    renderShiftStats(d.shiftStats || []);
    renderTop10("reasonBubbles", d.reasonsTop10, evaluatedBase || total);
    renderTop10("sentBackBubbles", d.sentBackTop10, evaluatedBase || total);

    renderDelaySummary(d);
    renderDelayTable(d.delayRows);
  }

  /***********************
   SHIFT STATS
  ***********************/
  function renderShiftStats(rows) {
    const wrap = els.shiftStatsWrap;
    if (!wrap) return;

    if (!rows?.length) {
      wrap.innerHTML = `<div class="empty-state">No shift data</div>`;
      return;
    }

    wrap.innerHTML = rows.map(r => `
      <div class="pill">
        <div class="pill-left">
          <div class="pill-title">${escapeHtml(r.shift)}</div>
          <div class="pill-bar">
            <div class="pill-fill" style="width:${Math.min(safeNumber(r.percent), 100)}%"></div>
          </div>
          <div class="pill-pct">
            ${formatPct1(r.percent)} of evaluated •
            Same-Day ${formatPct1(r.sameDayPercent)} •
            Prior Day ${formatPct1(r.priorDayPercent)} •
            Avg Delay ${formatDec(r.avgDelay, 2)} days •
            Avg Investigation ${formatDec(r.avgInvestigationDays, 2)} days
          </div>
        </div>
        <div class="pill-count">${formatInt(r.count)}</div>
      </div>
    `).join("");
  }

  /***********************
   TOP 10
  ***********************/
  function renderTop10(target, rows, total) {
    const wrap = document.getElementById(target);
    if (!wrap) return;

    if (!rows?.length) {
      wrap.innerHTML = `<div class="empty-state">No data</div>`;
      return;
    }

    wrap.innerHTML = rows.map(item => {
      const pct = total ? (safeNumber(item.count) / safeNumber(total)) * 100 : 0;

      return `
        <div class="pill">
          <div class="pill-left">
            <div class="pill-title">${escapeHtml(item.name)}</div>
            <div class="pill-bar">
              <div class="pill-fill" style="width:${pct.toFixed(1)}%"></div>
            </div>
            <div class="pill-pct">${pct.toFixed(1)}%</div>
          </div>
          <div class="pill-count">${formatInt(item.count)}</div>
        </div>
      `;
    }).join("");
  }

  /***********************
   FILTER FUNCTION
  ***********************/
  function applyDelayFilter(rows) {
    return (rows || []).filter(r => {
      const d = safeNumber(r.daysToArrive);

      if (delayFilter === "1-2") return d >= 1 && d <= 2;
      if (delayFilter === "3-5") return d > 2 && d <= 5;
      if (delayFilter === "5+") return d > 5;

      return true;
    });
  }

  /***********************
   DELAY SUMMARY
  ***********************/
  function renderDelaySummary(d) {
    let rows = applyDelayFilter(d.delayRows || []);

    let b12 = 0, b35 = 0, b5p = 0;

    rows.forEach(r => {
      const days = safeNumber(r.daysToArrive);

      if (days >= 1 && days <= 2) b12++;
      else if (days > 2 && days <= 5) b35++;
      else if (days > 5) b5p++;
    });

    if (els.bucket12) els.bucket12.textContent = b12;
    if (els.bucket35) els.bucket35.textContent = b35;
    if (els.bucket5p) els.bucket5p.textContent = b5p;

    const map = {};

    rows.forEach(r => {
      const key = r.sentBack || "Unknown";
      const days = safeNumber(r.daysToArrive);

      if (!map[key]) {
        map[key] = { count: 0, total: 0, max: 0 };
      }

      map[key].count++;
      map[key].total += days;
      if (days > map[key].max) map[key].max = days;
    });

    const grouped = Object.keys(map).map(k => ({
      name: k,
      count: map[k].count,
      avgDays: (map[k].total / map[k].count).toFixed(1),
      maxDays: map[k].max.toFixed(1)
    }));

    grouped.sort((a, b) => safeNumber(b.maxDays) - safeNumber(a.maxDays));

    const wrap = els.delayByDept;
    if (!wrap) return;

    if (!grouped.length) {
      wrap.innerHTML = `<div class="empty-state">No delay data</div>`;
      return;
    }

    wrap.innerHTML = grouped.map(r => `
      <div class="pill">
        <div class="pill-left">
          <div class="pill-title">${escapeHtml(r.name)}</div>
          <div class="pill-bar">
            <div class="pill-fill" style="width:${Math.min(r.count * 10, 100)}%"></div>
          </div>
          <div class="pill-pct">
            Avg: ${r.avgDays} days • Max: ${r.maxDays}
          </div>
        </div>
        <div class="pill-count">${formatInt(r.count)}</div>
      </div>
    `).join("");
  }

  /***********************
   DELAY TABLE
  ***********************/
  function renderDelayTable(rows) {
    rows = applyDelayFilter(rows);

    const body = els.delayTableBody;
    if (!body) return;

    if (!rows?.length) {
      body.innerHTML = `<tr><td colspan="5">No delay records</td></tr>`;
      return;
    }

    rows.sort((a, b) =>
      safeNumber(b.daysToArrive) - safeNumber(a.daysToArrive)
    );

    const groups = {
      "1-2": [],
      "3-5": [],
      "5+": []
    };

    rows.forEach(r => {
      const d = safeNumber(r.daysToArrive);

      if (d >= 1 && d <= 2) groups["1-2"].push(r);
      else if (d > 2 && d <= 5) groups["3-5"].push(r);
      else if (d > 5) groups["5+"].push(r);
    });

    let html = "";

    Object.entries(groups).forEach(([range, list]) => {
      if (!list.length) return;

      const id = `group-${range}`;

      html += `
        <tr class="group-header" data-target="${id}">
          <td colspan="5">
            <span class="toggle">▶</span>
            ${range} Days (${list.length})
          </td>
        </tr>
      `;

      html += list.map(r => {
        const d = safeNumber(r.daysToArrive);

        const cls =
          d >= 3 ? "delay-high" :
          d >= 1 ? "delay-medium" :
          "delay-low";

        return `
          <tr class="group-row ${id}" style="display:none;">
            <td>${escapeHtml(r.rx)}</td>
            <td>${escapeHtml(r.sentBack)}</td>
            <td>${escapeHtml(r.scanDate)}</td>
            <td>${escapeHtml(r.arrived)}</td>
            <td>
              <span class="delay-badge ${cls}">
                ${d}
              </span>
            </td>
          </tr>
        `;
      }).join("");
    });

    body.innerHTML = html;

    document.querySelectorAll(".group-header").forEach(header => {
      header.onclick = () => {
        const target = header.dataset.target;
        const groupRows = document.querySelectorAll(`.${target}`);
        const icon = header.querySelector(".toggle");

        const isOpen = groupRows[0]?.style.display !== "none";

        groupRows.forEach(r => {
          r.style.display = isOpen ? "none" : "table-row";
        });

        if (icon) icon.textContent = isOpen ? "▶" : "▼";
      };
    });
  }

  /***********************
   TABS
  ***********************/
  function switchTab(tab) {
    activeTab = tab;

    if (els.reasonsView)
      els.reasonsView.style.display = tab === "reasons" ? "block" : "none";

    if (els.sentBackView)
      els.sentBackView.style.display = tab === "sent" ? "block" : "none";

    if (els.delayView)
      els.delayView.style.display = tab === "delay" ? "block" : "none";

    els.tabReasons?.classList.toggle("active", tab === "reasons");
    els.tabSentBack?.classList.toggle("active", tab === "sent");
    els.tabDelay?.classList.toggle("active", tab === "delay");
  }

  /***********************
   DELAY FILTER CLICK
  ***********************/
  document.querySelectorAll(".delay-filter").forEach(el => {
    el.addEventListener("click", () => {
      const selected = el.dataset.range;

      delayFilter = delayFilter === selected ? "all" : selected;

      document.querySelectorAll(".delay-filter")
        .forEach(x => x.classList.remove("active"));

      if (delayFilter !== "all") {
        el.classList.add("active");
      }

      if (lastData) {
        renderDelaySummary(lastData);
        renderDelayTable(lastData.delayRows);
      }
    });
  });

  /***********************
   SUB TABS
  ***********************/
  function switchDelayTab(tab) {
    document.getElementById("delaySummaryView").style.display =
      tab === "summary" ? "block" : "none";

    document.getElementById("delayDeptView").style.display =
      tab === "dept" ? "block" : "none";

    document.getElementById("delayTableView").style.display =
      tab === "table" ? "block" : "none";

    document.getElementById("delayTabSummary")?.classList.toggle("active", tab === "summary");
    document.getElementById("delayTabDept")?.classList.toggle("active", tab === "dept");
    document.getElementById("delayTabTable")?.classList.toggle("active", tab === "table");
  }

  document.getElementById("delayTabSummary")?.addEventListener("click", () => switchDelayTab("summary"));
  document.getElementById("delayTabDept")?.addEventListener("click", () => switchDelayTab("dept"));
  document.getElementById("delayTabTable")?.addEventListener("click", () => switchDelayTab("table"));

  /***********************
   EVENTS
  ***********************/
  els.dateFilter?.addEventListener("change", load);

  els.btnAll?.addEventListener("click", () => {
    els.dateFilter.value = "";
    load();
  });

  els.tabReasons?.addEventListener("click", () => switchTab("reasons"));
  els.tabSentBack?.addEventListener("click", () => switchTab("sent"));
  els.tabDelay?.addEventListener("click", () => switchTab("delay"));
  
  // 🔥 YEAR TABS (ADD HERE)
  
document.querySelectorAll(".year-tab").forEach(btn => {
  btn.addEventListener("click", () => {

    currentYear = btn.dataset.year;

    // UI active state
    document.querySelectorAll(".year-tab")
      .forEach(b => b.classList.remove("active"));

    btn.classList.add("active");

    load();
  });
});

  /***********************
   INIT
  ***********************/
  switchTab(activeTab);
  switchDelayTab("summary");
  load();

});