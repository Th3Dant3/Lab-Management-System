document.addEventListener("DOMContentLoaded", () => {

  const API_URL =
    "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

  let lastData    = null;
  let currentYear = "all";
  let delayFilter = "all";
  let chartScan      = null;
  let chartArrival   = null;
  let chartEvaluated = null;

  /* ── HELPERS ── */
  const safe    = v => Number.isFinite(Number(v)) ? Number(v) : 0;
  const fmtInt  = v => safe(v).toLocaleString("en-US");
  const fmtPct  = v => `${safe(v).toFixed(0)}%`;
  const fmtDec  = (v, d = 2) => safe(v).toFixed(d);
  const esc     = s => String(s ?? "")
    .replaceAll("&","&amp;").replaceAll("<","&lt;")
    .replaceAll(">","&gt;").replaceAll('"',"&quot;")
    .replaceAll("'","&#39;");
  const getDate = () => document.getElementById("dateFilter")?.value?.trim() || "all";
  const set     = (id, v) => { const e = document.getElementById(id); if (e) e.textContent = v; };

  /* ── LOAD ── */
  async function load() {
    const url = `${API_URL}?date=${encodeURIComponent(getDate())}&year=${currentYear}`;
    try {
      const res  = await fetch(url, { cache: "no-store" });
      const data = await res.json();
      lastData = data;
      render(data);
    } catch (e) { console.error("LOAD ERROR:", e); }
  }

  /* ── RENDER ── */
  function render(d) {
    const total    = safe(d.total);
    const evalBase = safe(getDate() === "all" ? d.evaluated : d.evaluatedToday);

    /* header mirror */
    set("headerStatus",   safe(d.active) === 0 ? "All clear" : `${fmtInt(d.active)} active`);
    set("headerCoverage", fmtPct(d.coverage));
    set("headerSync",     d.lastUpdated || "—");

    /* primary KPIs */
    set("activeHolds",    fmtInt(d.active));
    set("evaluatedCount", fmtInt(d.evaluated));
    set("coveragePct",    fmtPct(d.coverage));
    set("lastUpdated",    d.lastUpdated || "—");

    /* active holds bar */
    const bar = document.getElementById("activeBar");
    if (bar) {
      const pct = safe(d.evaluated) > 0
        ? Math.min(safe(d.active) / safe(d.evaluated) * 100, 100) : 0;
      setTimeout(() => { bar.style.width = pct + "%"; }, 200);
    }

    /* secondary KPIs */
    set("avgInvestigationTime", `${fmtDec(d.avgInvestigationDays)}d / ${fmtDec(d.avgInvestigationHours)}h`);
    set("arrivedToday",         fmtInt(d.arrived));
    set("evaluatedToday",       fmtInt(d.evaluatedToday));
    set("priorDayIntake",       fmtInt(d.priorDayIntake));
    set("avgArrivalDelayDays",  `${fmtDec(d.avgArrivalDelayDays)} days`);
    set("oldestInvestigation",  d.oldestInvestigation || "—");
    set("latestInvestigation",  d.latestInvestigation || "—");

    renderShifts(d.shiftStats || []);
    renderPills("reasonBubbles",   d.reasonsTop10,  evalBase || total);
    renderPills("sentBackBubbles", d.sentBackTop10, evalBase || total);
    renderDelay(d);
    // Charts render on-demand when Trend tab is clicked
  }

  /* ── SHIFTS ── */
  function renderShifts(rows) {
    const wrap = document.getElementById("shiftStatsWrap");
    if (!wrap) return;
    if (!rows.length) {
      wrap.innerHTML = `<div class="empty-state">No shift data</div>`;
      return;
    }
    const max = Math.max(...rows.map(r => safe(r.count)));
    wrap.innerHTML = rows.map(r => `
      <div class="shift-row">
        <div class="sh-name">${esc(r.shift)}</div>
        <div class="sh-track">
          <div class="sh-fill" style="width:${max ? Math.round(safe(r.count)/max*100) : 0}%"></div>
        </div>
        <div class="sh-meta">
          <b>${fmtDec(r.percent,1)}%</b> of evaluated &nbsp;·&nbsp;
          Same-day <b>${fmtDec(r.sameDayPercent,1)}%</b> &nbsp;·&nbsp;
          Prior day <b>${fmtDec(r.priorDayPercent,1)}%</b> &nbsp;·&nbsp;
          Avg delay <b>${fmtDec(r.avgDelay,2)}d</b> &nbsp;·&nbsp;
          Avg inv <b>${fmtDec(r.avgInvestigationDays,2)}d</b>
        </div>
        <div class="sh-badge">${fmtInt(r.count)}</div>
      </div>`).join("");
  }

  /* ── PILLS ── */
  function renderPills(target, rows, total) {
    const wrap = document.getElementById(target);
    if (!wrap) return;
    if (!rows?.length) {
      wrap.innerHTML = `<div class="empty-state" style="padding:16px 22px;">No data available</div>`;
      return;
    }
    const max = rows[0]?.count || 1;
    wrap.innerHTML = rows.map((item, i) => {
      const pct = total ? (safe(item.count) / safe(total)) * 100 : 0;
      return `
        <div class="pill">
          <div class="pill-rank">${i + 1}</div>
          <div class="pill-name">${esc(item.name)}</div>
          <div class="pill-track">
            <div class="pill-fill" style="width:${(safe(item.count)/max*100).toFixed(1)}%"></div>
          </div>
          <div class="pill-pct">${pct.toFixed(1)}%</div>
          <div class="pill-cnt">${fmtInt(item.count)}</div>
        </div>`;
    }).join("");
  }

  /* ── DELAY ── */
  function applyDelayFilter(rows) {
    return (rows || []).filter(r => {
      const v = safe(r.daysToArrive);
      if (delayFilter === "1-2") return v >= 1 && v <= 2;
      if (delayFilter === "3-5") return v >  2 && v <= 5;
      if (delayFilter === "5+")  return v >  5;
      return true;
    });
  }

  function renderDelay(d) {
    const rows = applyDelayFilter(d.delayRows || []);
    let b12 = 0, b35 = 0, b5p = 0;
    rows.forEach(r => {
      const v = safe(r.daysToArrive);
      if      (v >= 1 && v <= 2) b12++;
      else if (v >  2 && v <= 5) b35++;
      else if (v >  5)           b5p++;
    });
    set("bucket-1-2",   b12);
    set("bucket-3-5",   b35);
    set("bucket-5plus", b5p);

    /* by dept */
    const map = {};
    rows.forEach(r => {
      const k = r.sentBack || "Unknown";
      const v = safe(r.daysToArrive);
      if (!map[k]) map[k] = { count: 0, total: 0, max: 0 };
      map[k].count++;
      map[k].total += v;
      if (v > map[k].max) map[k].max = v;
    });
    const dept = Object.entries(map)
      .map(([k, v]) => ({ name: k, count: v.count, avg: (v.total/v.count).toFixed(1), max: v.max.toFixed(1) }))
      .sort((a, b) => b.count - a.count);

    const dw = document.getElementById("delayByDept");
    if (dw) {
      if (!dept.length) {
        dw.innerHTML = `<div class="empty-state" style="padding:16px 22px;">No delay data</div>`;
      } else {
        const mx = dept[0].count;
        dw.innerHTML = dept.map((r, i) => `
          <div class="pill">
            <div class="pill-rank">${i + 1}</div>
            <div class="pill-name">${esc(r.name)}</div>
            <div class="pill-track">
              <div class="pill-fill" style="width:${Math.round(r.count/mx*100)}%"></div>
            </div>
            <div class="pill-pct" style="min-width:120px;text-align:right;font-size:10px;">
              avg ${r.avg}d &nbsp;·&nbsp; max ${r.max}d
            </div>
            <div class="pill-cnt">${fmtInt(r.count)}</div>
          </div>`).join("");
      }
    }

    /* timeline table */
    const tbody = document.getElementById("delayTableBody");
    if (tbody) {
      const sorted = [...rows].sort((a, b) => safe(b.daysToArrive) - safe(a.daysToArrive));
      const grps = { "1-2": [], "3-5": [], "5+": [] };
      sorted.forEach(r => {
        const v = safe(r.daysToArrive);
        if      (v >= 1 && v <= 2) grps["1-2"].push(r);
        else if (v >  2 && v <= 5) grps["3-5"].push(r);
        else if (v >  5)           grps["5+"].push(r);
      });

      let html = "";
      Object.entries(grps).forEach(([range, list]) => {
        if (!list.length) return;
        const id  = `grp-${range.replace("+","p").replace("-","")}`;
        const col = range === "5+" ? "#ff4d6a" : range === "3-5" ? "#f5a623" : "#00e87a";
        html += `
          <tr class="grp-hdr-row" data-target="${id}">
            <td colspan="5">
              <span class="tog">▶</span>${range} days (${list.length})
            </td>
          </tr>`;
        html += list.map(r => {
          const v   = safe(r.daysToArrive);
          const cls = v > 5 ? "hi" : v >= 3 ? "mid" : "low";
          return `
            <tr class="grp-row ${id}" style="display:none;">
              <td>${esc(r.rx)}</td>
              <td>${esc(r.sentBack)}</td>
              <td>${esc(r.scanDate)}</td>
              <td>${esc(r.arrived)}</td>
              <td><span class="dbadge ${cls}">${v}d</span></td>
            </tr>`;
        }).join("");
      });

      tbody.innerHTML = html || `
        <tr><td colspan="5" style="color:#1e1e1e;padding:14px 22px;font-size:12px;">
          No delay records
        </td></tr>`;

      document.querySelectorAll(".grp-hdr-row").forEach(h => {
        h.onclick = () => {
          const gRows = document.querySelectorAll("." + h.dataset.target);
          const icon  = h.querySelector(".tog");
          const open  = gRows[0]?.style.display !== "none";
          gRows.forEach(r => { r.style.display = open ? "none" : "table-row"; });
          if (icon) icon.textContent = open ? "▶" : "▼";
        };
      });
    }
  }

  /* ── CHARTS ── */
  function renderChart(canvasId, key, trendObj, color, bg) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    const entries = Object.entries(trendObj)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .slice(-45);
    if (!entries.length) return;

    const labels = entries.map(e => {
      const [, m, d] = e[0].split("-");
      return `${m}/${d}`;
    });
    const values = entries.map(e => e[1]);

    const prev = key === "arrival" ? chartArrival : key === "scan" ? chartScan : chartEvaluated;
    if (prev) { prev.destroy(); }

    const cfg = {
      type: "bar",
      data: {
        labels,
        datasets: [{
          data:            values,
          backgroundColor: bg,
          borderColor:     color,
          borderWidth:     1,
          borderRadius:    2,
          borderSkipped:   "bottom",
          barPercentage:   0.85,
          categoryPercentage: 0.9
        }]
      },
      options: {
        animation:           false,
        responsive:          true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: "#13151a",
            borderColor:     "#1f2330",
            borderWidth:     1,
            titleColor:      color,
            bodyColor:       "#8b96b0",
            titleFont:       { size: 10, weight: "700" },
            bodyFont:        { size: 11, weight: "700" },
            callbacks: {
              title: items => entries[items[0].dataIndex][0],
              label: item  => ` ${item.raw} records`
            }
          }
        },
        scales: {
          x: {
            ticks: { color: "#4a5368", font: { size: 9, weight: "600" }, maxTicksLimit: 12, maxRotation: 0, autoSkip: true },
            grid:  { color: "#1f2330" },
            border:{ color: "#1f2330" }
          },
          y: {
            ticks:       { color: "#4a5368", font: { size: 9, weight: "600" } },
            grid:        { color: "#1f2330" },
            border:      { color: "#1f2330" },
            beginAtZero: true
          }
        }
      }
    };

    const inst = new Chart(canvas, cfg);
    if (key === "arrival")   chartArrival   = inst;
    else if (key === "scan") chartScan      = inst;
    else                     chartEvaluated = inst;
  }

  /* ── TABS ── */
  function switchMainTab(tab) {
    const views = {
      reasons:  document.getElementById("reasonsView"),
      sent:     document.getElementById("sentBackView"),
      delay:    document.getElementById("delayView"),
      trend:    document.getElementById("trendView"),
      pipeline: document.getElementById("pipelineView")
    };
    Object.entries(views).forEach(([k, el]) => {
      if (el) el.style.display = k === tab ? "block" : "none";
    });
    document.getElementById("tabReasons")?.classList.toggle("active",   tab === "reasons");
    document.getElementById("tabSentBack")?.classList.toggle("active",  tab === "sent");
    document.getElementById("tabDelay")?.classList.toggle("active",     tab === "delay");
    document.getElementById("tabTrend")?.classList.toggle("active",     tab === "trend");
    document.getElementById("tabPipeline")?.classList.toggle("active",  tab === "pipeline");

    if (tab === "trend") {
      setTimeout(() => {
        if (!lastData) return;

        // Build scan trend from rows (col D) since API doesn't send it separately
        const scanTrend = {};
        (lastData.rows || []).forEach(r => {
          if (!r.scanDate) return;
          // scanDate comes as MM/dd/yyyy — convert to yyyy-MM-dd for sorting
          const parts = r.scanDate.split("/");
          if (parts.length === 3) {
            const key = `${parts[2]}-${parts[0].padStart(2,"0")}-${parts[1].padStart(2,"0")}`;
            scanTrend[key] = (scanTrend[key] || 0) + 1;
          }
        });

        renderChart("scanTrendChart",      "scan",     scanTrend,                    "#4a7ab0", "rgba(74,122,176,.2)");
        renderChart("arrivalTrendChart",   "arrival",  lastData.trend          || {}, "#00e87a", "rgba(0,232,122,.18)");
        renderChart("evaluatedTrendChart", "evaluated",lastData.evaluatedTrend || {}, "#f5a623", "rgba(245,166,35,.18)");
      }, 80);
    }
    if (tab === "pipeline" && lastData) renderPipeline(lastData);
  }

  function switchDelayTab(tab) {
    document.getElementById("delaySummaryView").style.display = tab === "summary" ? "block" : "none";
    document.getElementById("delayDeptView").style.display    = tab === "dept"    ? "block" : "none";
    document.getElementById("delayTableView").style.display   = tab === "table"   ? "block" : "none";
    document.getElementById("delayTabSummary")?.classList.toggle("active", tab === "summary");
    document.getElementById("delayTabDept")?.classList.toggle("active",    tab === "dept");
    document.getElementById("delayTabTable")?.classList.toggle("active",   tab === "table");
  }

  /* ── EVENTS ── */
  document.getElementById("tabReasons")?.addEventListener("click",   () => switchMainTab("reasons"));
  document.getElementById("tabSentBack")?.addEventListener("click",  () => switchMainTab("sent"));
  document.getElementById("tabDelay")?.addEventListener("click",     () => switchMainTab("delay"));
  document.getElementById("tabTrend")?.addEventListener("click",     () => switchMainTab("trend"));
  document.getElementById("tabPipeline")?.addEventListener("click",  () => switchMainTab("pipeline"));

  document.getElementById("delayTabSummary")?.addEventListener("click", () => switchDelayTab("summary"));
  document.getElementById("delayTabDept")?.addEventListener("click",    () => switchDelayTab("dept"));
  document.getElementById("delayTabTable")?.addEventListener("click",   () => switchDelayTab("table"));

  document.querySelectorAll(".delay-filter").forEach(el => {
    el.addEventListener("click", () => {
      const sel = el.dataset.range;
      delayFilter = delayFilter === sel ? "all" : sel;
      document.querySelectorAll(".delay-filter").forEach(x => x.classList.remove("active"));
      if (delayFilter !== "all") el.classList.add("active");
      if (lastData) renderDelay(lastData);
    });
  });

  document.getElementById("dateFilter")?.addEventListener("change", load);

  document.getElementById("clearDate")?.addEventListener("click", () => {
    const df = document.getElementById("dateFilter");
    if (df) df.value = "";
    load();
  });

  document.querySelectorAll(".fb-yr").forEach(btn => {
    btn.addEventListener("click", () => {
      currentYear = btn.dataset.year;
      document.querySelectorAll(".fb-yr").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      const df = document.getElementById("dateFilter");
      if (df) df.value = "";
      load();
    });
  });

  /* ── PIPELINE ── */
  function daysBetween(a, b) {
    if (!a || !b) return null;
    // Parse MM/dd/yyyy format from API
    function parseDate(s) {
      if (!s) return null;
      // Try MM/dd/yyyy first
      const p = s.split("/");
      if (p.length === 3) return new Date(`${p[2]}-${p[0].padStart(2,"0")}-${p[1].padStart(2,"0")}T00:00:00`);
      return new Date(s);
    }
    const da = parseDate(a), db = parseDate(b);
    if (!da || !db || isNaN(da) || isNaN(db)) return null;
    const diff = Math.round((db - da) / (1000 * 60 * 60 * 24));
    return diff;
  }

  function gapBadge(days) {
    if (days === null || days === undefined) return `<span class="dt-cell">—</span>`;
    const cls = days <= 1 ? "gbadge gbadge-ok" : days <= 3 ? "gbadge gbadge-warn" : "gbadge gbadge-bad";
    return `<span class="${cls}">${days}d</span>`;
  }

  function renderPipeline(d) {
    // Use rows that match current filter window (same rows shown in other tabs)
    const rows = (d.rows || []).filter(r => r.scanDate && r.arrived);
    if (!rows.length) {
      const tbody = document.getElementById("pipelineTableBody");
      if (tbody) tbody.innerHTML = `<tr><td colspan="8" class="pipe-empty">No journey data available</td></tr>`;
      return;
    }

    let totalScan = 0, countScan = 0, totalHold = 0, countHold = 0, maxTotal = 0;

    const processed = rows.map(r => {
      const scanGap = daysBetween(r.scanDate, r.arrived);

      // Compute hold from arrived→completed dates — most accurate
      // Fall back to col G if dates unavailable
      let holdDays = null;
      if (r.arrived && r.completed) {
        holdDays = daysBetween(r.arrived, r.completed);
      } else if (r.duration !== null && r.duration !== "" && r.duration !== undefined) {
        holdDays = Math.round(Number(r.duration));
      }

      if (scanGap  !== null) { totalScan += scanGap;  countScan++; }
      if (holdDays !== null) { totalHold += holdDays; countHold++; }

      const tot = (scanGap || 0) + (holdDays || 0);
      if (tot > maxTotal) maxTotal = tot;
      return { ...r, scanGap, holdDays, tot };
    });

    const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
    set("pipeTotal",     fmtInt(rows.length));
    set("pipeArrived",   fmtInt(rows.filter(r => r.arrived).length));
    set("pipeCompleted", fmtInt(rows.filter(r => r.completed).length));
    set("pipeGap1", countScan >= 0 ? `${(countScan ? totalScan/countScan : 0).toFixed(1)}d avg` : "—");
    set("pipeGap2", countHold >= 0 ? `${(countHold ? totalHold/countHold : 0).toFixed(1)}d avg` : "—");

    // Show most recent 60 rows sorted newest first
    const tbody = document.getElementById("pipelineTableBody");
    if (!tbody) return;

    const display = [...processed].sort((a, b) => {
      const da = new Date(a.arrived || a.scanDate || "");
      const db = new Date(b.arrived || b.scanDate || "");
      return db - da;
    }).slice(0, 60);
    const mx = maxTotal || 10;

    tbody.innerHTML = display.map(r => {
      const sg = r.scanGap, hd = r.holdDays;
      const scanW  = Math.max(sg  !== null ? Math.round(sg  / mx * 80) : 5, 5);
      const holdW  = Math.max(hd  !== null ? Math.round(hd  / mx * 80) : 5, 5);
      const restW  = Math.max(100 - scanW - holdW - 4, 0);

      return `
        <tr>
          <td><span class="rx-cell">${esc(r.rx)}</span></td>
          <td><span class="sent-cell">${esc(r.sentBack || "—")}</span></td>
          <td><span class="dt-cell">${esc(r.scanDate || "—")}</span></td>
          <td>${gapBadge(sg)}</td>
          <td><span class="dt-cell">${esc(r.arrived || "—")}</span></td>
          <td>${gapBadge(hd)}</td>
          <td><span class="dt-cell">${esc(r.completed || "—")}</span></td>
          <td>
            <div class="jbar">
              <div class="jbar-scan" style="width:${scanW}%">${sg !== null ? sg + "d" : ""}</div>
              <div class="jbar-marker"></div>
              <div class="jbar-hold" style="width:${holdW}%">${hd !== null ? hd + "d" : ""}</div>
            </div>
          </td>
        </tr>`;
    }).join("");
  }

  /* ── INIT ── */
  switchMainTab("reasons");
  switchDelayTab("summary");
  load();

});