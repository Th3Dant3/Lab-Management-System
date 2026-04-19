/* =====================================================
   INDEX DASHBOARD — Command Center Redesign
===================================================== */

const API_URL =
  "https://script.google.com/a/macros/zennioptical.com/s/AKfycbzbQBjzoEEBpvukFkR-XMw8kG_gzCIuxZrTLodZZ_EnwqYAujOBqSzYslx-x9XTw7_UUA/exec";

const AUTH_API =
  "https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

const SCANNER_STORAGE_KEY = "lms_scanner_attention";

/* =====================================================
   CUSTOM CURSOR (tracks mouse during loading)
===================================================== */
(function initCursor() {
  const cursor = document.getElementById("lms-cursor");
  if (!cursor) return;

  let mx = window.innerWidth / 2;
  let my = window.innerHeight / 2;
  let cx = mx, cy = my;
  let rafId;

  document.addEventListener("mousemove", e => {
    mx = e.clientX;
    my = e.clientY;
  });

  function animateCursor() {
    cx += (mx - cx) * 0.18;
    cy += (my - cy) * 0.18;
    cursor.style.transform = `translate(${Math.round(cx)}px, ${Math.round(cy)}px) translate(-50%, -50%)`;
    rafId = requestAnimationFrame(animateCursor);
  }
  animateCursor();

  window._stopCursorAnim = () => {
    cancelAnimationFrame(rafId);
  };
})();

/* =====================================================
   LOADING SCREEN
===================================================== */
(function initLoader() {
  const loader   = document.getElementById("lms-loader");
  const fill     = document.getElementById("loaderFill");
  const status   = document.getElementById("loaderStatus");
  const canvas   = document.getElementById("loaderCanvas");
  const cursor   = document.getElementById("lms-cursor");
  if (!loader) return;

  document.body.classList.add("lms-loading");

  /* Particle background on the loader canvas */
  const ctx = canvas.getContext("2d");
  function resizeCanvas() {
    canvas.width  = window.innerWidth;
    canvas.height = window.innerHeight;
  }
  resizeCanvas();
  window.addEventListener("resize", resizeCanvas);

  const GOLD  = "rgba(245,200,66,";
  const STEEL = "rgba(100,160,220,";
  const particles = Array.from({ length: 55 }, () => ({
    x:     Math.random() * window.innerWidth,
    y:     Math.random() * window.innerHeight,
    vx:    (Math.random() - 0.5) * 0.5,
    vy:    (Math.random() - 0.5) * 0.5,
    r:     1 + Math.random() * 2,
    alpha: 0.1 + Math.random() * 0.35,
    color: Math.random() > 0.5 ? GOLD : STEEL,
  }));

  let loaderRaf;
  function drawParticles() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    particles.forEach(p => {
      p.x += p.vx; p.y += p.vy;
      if (p.x < 0) p.x = canvas.width;
      if (p.x > canvas.width) p.x = 0;
      if (p.y < 0) p.y = canvas.height;
      if (p.y > canvas.height) p.y = 0;
      ctx.beginPath();
      ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
      ctx.fillStyle = p.color + p.alpha + ")";
      ctx.fill();
    });
    /* faint connecting lines between close particles */
    for (let i = 0; i < particles.length; i++) {
      for (let j = i + 1; j < particles.length; j++) {
        const dx = particles[i].x - particles[j].x;
        const dy = particles[i].y - particles[j].y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist < 120) {
          ctx.beginPath();
          ctx.moveTo(particles[i].x, particles[i].y);
          ctx.lineTo(particles[j].x, particles[j].y);
          ctx.strokeStyle = `rgba(100,160,220,${0.06 * (1 - dist / 120)})`;
          ctx.lineWidth = 0.5;
          ctx.stroke();
        }
      }
    }
    loaderRaf = requestAnimationFrame(drawParticles);
  }
  drawParticles();

  /* Progress bar steps */
  const steps = [
    { pct: 20, msg: "Authenticating session…" },
    { pct: 45, msg: "Loading modules…"        },
    { pct: 70, msg: "Fetching live data…"     },
    { pct: 90, msg: "Building dashboard…"     },
    { pct: 100, msg: "Ready"                  },
  ];
  let stepIdx = 0;

  function advanceLoader(targetPct, msg) {
    if (fill)   fill.style.width = targetPct + "%";
    if (status) status.textContent = msg;
  }

  /* Tick through fake progress while real data loads */
  const stepTimer = setInterval(() => {
    if (stepIdx < steps.length - 1) {
      const s = steps[stepIdx++];
      advanceLoader(s.pct, s.msg);
    }
  }, 480);

  /* Called by renderDashboard when real data is ready */
  window.dismissLoader = function () {
    clearInterval(stepTimer);
    advanceLoader(100, "Ready");
    setTimeout(() => {
      cancelAnimationFrame(loaderRaf);
      loader.classList.add("fade-out");
      document.body.classList.remove("lms-loading");
      if (cursor) cursor.classList.add("hidden");
      if (window._stopCursorAnim) window._stopCursorAnim();
    }, 400);
  };
})();

/* =====================================================
   NAV CANVAS ANIMATIONS
===================================================== */
function initNavAnimations() {
  animOps(document.getElementById("nav-canvas-ops"));
  animProd(document.getElementById("nav-canvas-prod"));
  animSys(document.getElementById("nav-canvas-sys"));
  animInv(document.getElementById("nav-canvas-inv"));
}

/* Inventory nav — stacked bar chart filling up like stock levels */
function animInv(canvas) {
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  const W = canvas.width, H = canvas.height;
  const GREEN = "#3fd189";
  const cols = 8;
  const bars = Array.from({ length: cols }, (_, i) => ({
    x: 6 + i * ((W - 12) / cols),
    fill: 0.3 + Math.random() * 0.7,
    target: 0.3 + Math.random() * 0.7,
    speed: 0.008 + Math.random() * 0.006,
  }));

  function draw() {
    ctx.clearRect(0, 0, W, H);
    const bw = ((W - 12) / cols) - 4;

    bars.forEach(b => {
      /* drift toward new target */
      b.fill += (b.target - b.fill) * b.speed;
      if (Math.abs(b.fill - b.target) < 0.01) {
        b.target = 0.2 + Math.random() * 0.75;
      }

      const bh = b.fill * (H - 8);
      const alpha = 0.35 + b.fill * 0.55;

      /* empty track */
      ctx.fillStyle = `rgba(63,209,137,0.08)`;
      ctx.beginPath();
      ctx.roundRect(b.x, 4, bw, H - 8, 2);
      ctx.fill();

      /* filled portion */
      ctx.fillStyle = `rgba(63,209,137,${alpha})`;
      ctx.beginPath();
      ctx.roundRect(b.x, H - 4 - bh, bw, bh, 2);
      ctx.fill();

      /* top cap glow */
      ctx.fillStyle = GREEN;
      ctx.globalAlpha = 0.9;
      ctx.beginPath();
      ctx.roundRect(b.x, H - 4 - bh, bw, 2, 1);
      ctx.fill();
      ctx.globalAlpha = 1;
    });

    requestAnimationFrame(draw);
  }
  draw();
}

/* Operations — LMS scanning beam animation */
function animOps(canvas) {
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  const W = canvas.width, H = canvas.height;
  const GOLD = "#f5c842";

  /* Horizontal scan lines */
  const scanLines = Array.from({ length: 5 }, (_, i) => ({
    y:     (i / 4) * H,
    speed: 0.6 + i * 0.25,
    x:     Math.random() * W,
    len:   20 + Math.random() * 40,
    alpha: 0.3 + Math.random() * 0.4,
  }));

  /* Small blip dots */
  const blips = Array.from({ length: 8 }, () => ({
    x: Math.random() * W,
    y: Math.random() * H,
    phase: Math.random() * Math.PI * 2,
    speed: 0.04 + Math.random() * 0.03,
  }));

  /* Vertical sweep beam */
  let beamX = 0;
  const beamSpeed = 0.8;

  function draw() {
    ctx.clearRect(0, 0, W, H);

    /* Sweep beam */
    beamX = (beamX + beamSpeed) % W;
    ctx.fillStyle = `rgba(245,200,66,0.06)`;
    ctx.fillRect(beamX - 12, 0, 24, H);
    ctx.fillStyle = `rgba(245,200,66,0.25)`;
    ctx.fillRect(beamX - 1, 0, 2, H);

    /* Horizontal scan lines scrolling right */
    scanLines.forEach(l => {
      l.x = (l.x + l.speed) % (W + l.len);
      ctx.beginPath();
      ctx.moveTo(l.x - l.len, l.y);
      ctx.lineTo(l.x, l.y);
      ctx.strokeStyle = GOLD;
      ctx.globalAlpha = l.alpha;
      ctx.lineWidth = 0.8;
      ctx.stroke();
      ctx.globalAlpha = 1;
    });

    /* Blips */
    blips.forEach(b => {
      b.phase += b.speed;
      const a = 0.2 + 0.8 * Math.abs(Math.sin(b.phase));
      ctx.beginPath();
      ctx.arc(b.x, b.y, 1.5, 0, Math.PI * 2);
      ctx.fillStyle = GOLD;
      ctx.globalAlpha = a;
      ctx.fill();
      ctx.globalAlpha = 1;
    });

    requestAnimationFrame(draw);
  }
  draw();
}

/* Production — orbiting rings */
function animProd(canvas) {
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  const W = canvas.width, H = canvas.height;
  const cx = W / 2, cy = H / 2;
  const rings = [
    { r: 26, speed: 0.014,  dots: 5, color: "#9b6bff", size: 3 },
    { r: 16, speed: -0.022, dots: 3, color: "#c084fc", size: 2 },
    { r:  8, speed: 0.038,  dots: 2, color: "#7a4fd4", size: 1.5 },
  ];
  let t = 0;
  function draw() {
    ctx.clearRect(0, 0, W, H);
    t++;
    rings.forEach(ring => {
      ctx.strokeStyle = ring.color + "22";
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.arc(cx, cy, ring.r, 0, Math.PI * 2);
      ctx.stroke();
      for (let i = 0; i < ring.dots; i++) {
        const angle = t * ring.speed + (i / ring.dots) * Math.PI * 2;
        ctx.fillStyle = ring.color;
        ctx.beginPath();
        ctx.arc(cx + Math.cos(angle) * ring.r, cy + Math.sin(angle) * ring.r, ring.size, 0, Math.PI * 2);
        ctx.fill();
      }
    });
    const pulse = 0.5 + 0.5 * Math.sin(t * 0.07);
    ctx.fillStyle = `rgba(155,107,255,${0.5 + pulse * 0.5})`;
    ctx.beginPath();
    ctx.arc(cx, cy, 4.5, 0, Math.PI * 2);
    ctx.fill();
    ctx.strokeStyle = `rgba(155,107,255,${0.15 + pulse * 0.25})`;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.arc(cx, cy, 4.5 + 5 + pulse * 4, 0, Math.PI * 2);
    ctx.stroke();
    requestAnimationFrame(draw);
  }
  draw();
}

/* System — network mesh with traveling packets */
function animSys(canvas) {
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  const W = canvas.width, H = canvas.height;
  const cx = W / 2, cy = H / 2;
  const GREEN = "#3fd189";
  const nodes = [
    { x: cx,      y: cy,      r: 5,   main: true  },
    { x: cx - 44, y: cy - 18, r: 3.5, main: false },
    { x: cx + 44, y: cy - 18, r: 3.5, main: false },
    { x: cx - 44, y: cy + 18, r: 3.5, main: false },
    { x: cx + 44, y: cy + 18, r: 3.5, main: false },
    { x: cx,      y: cy - 32, r: 2.5, main: false },
    { x: cx,      y: cy + 32, r: 2.5, main: false },
  ];
  const edges = [[0,1],[0,2],[0,3],[0,4],[1,5],[2,5],[3,6],[4,6]];
  const packets = edges.map(e => ({
    edge: e,
    progress: Math.random(),
    speed: 0.009 + Math.random() * 0.007,
  }));
  let t = 0;
  function draw() {
    ctx.clearRect(0, 0, W, H);
    t++;
    edges.forEach(([a, b]) => {
      ctx.strokeStyle = "rgba(63,209,137,0.18)";
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(nodes[a].x, nodes[a].y);
      ctx.lineTo(nodes[b].x, nodes[b].y);
      ctx.stroke();
    });
    packets.forEach(p => {
      p.progress += p.speed;
      if (p.progress > 1) p.progress = 0;
      const [a, b] = p.edge;
      ctx.fillStyle = GREEN;
      ctx.beginPath();
      ctx.arc(
        nodes[a].x + (nodes[b].x - nodes[a].x) * p.progress,
        nodes[a].y + (nodes[b].y - nodes[a].y) * p.progress,
        2, 0, Math.PI * 2
      );
      ctx.fill();
    });
    nodes.forEach((n, i) => {
      const pulse = i === 0 ? 0.5 + 0.5 * Math.sin(t * 0.065) : 0.75;
      ctx.fillStyle = n.main ? `rgba(63,209,137,${pulse})` : "rgba(63,209,137,0.55)";
      ctx.beginPath();
      ctx.arc(n.x, n.y, n.r, 0, Math.PI * 2);
      ctx.fill();
      if (n.main) {
        ctx.strokeStyle = `rgba(63,209,137,${0.18 + pulse * 0.22})`;
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        ctx.arc(n.x, n.y, n.r + 4 + pulse * 4, 0, Math.PI * 2);
        ctx.stroke();
      }
    });
    requestAnimationFrame(draw);
  }
  draw();
}

/* =====================================================
   CARD BACKGROUND CANVAS ANIMATIONS
===================================================== */
function initCardCanvases() {
  document.querySelectorAll(".card-bg-canvas").forEach(canvas => {
    const anim  = canvas.dataset.anim;
    const color = canvas.dataset.color || "#64a0dc";
    if (anim === "pulse-dots")    cardAnimDots(canvas, color);
    if (anim === "flow-lines")    cardAnimLines(canvas, color);
    if (anim === "inventory-flow") cardAnimInventory(canvas, color);
  });
}

/* Floating particle dots */
function cardAnimDots(canvas, color) {
  const card = canvas.parentElement;
  function resize() {
    canvas.width  = card.offsetWidth;
    canvas.height = card.offsetHeight;
  }
  resize();
  new ResizeObserver(resize).observe(card);

  const ctx = canvas.getContext("2d");
  const count = 18;
  const dots = Array.from({ length: count }, () => ({
    x:    Math.random(),
    y:    Math.random(),
    r:    1.2 + Math.random() * 2.2,
    vx:   (Math.random() - 0.5) * 0.0006,
    vy:   (Math.random() - 0.5) * 0.0006,
    phase: Math.random() * Math.PI * 2,
    speed: 0.012 + Math.random() * 0.01,
  }));

  function tick() {
    const W = canvas.width, H = canvas.height;
    ctx.clearRect(0, 0, W, H);
    dots.forEach(d => {
      d.x += d.vx; d.y += d.vy; d.phase += d.speed;
      if (d.x < 0) d.x = 1; if (d.x > 1) d.x = 0;
      if (d.y < 0) d.y = 1; if (d.y > 1) d.y = 0;
      const alpha = 0.3 + 0.7 * Math.abs(Math.sin(d.phase));
      ctx.beginPath();
      ctx.arc(d.x * W, d.y * H, d.r, 0, Math.PI * 2);
      ctx.fillStyle = color;
      ctx.globalAlpha = alpha;
      ctx.fill();
      ctx.globalAlpha = 1;
    });
    /* connect nearby dots */
    for (let i = 0; i < dots.length; i++) {
      for (let j = i + 1; j < dots.length; j++) {
        const dx = (dots[i].x - dots[j].x) * W;
        const dy = (dots[i].y - dots[j].y) * H;
        const dist = Math.sqrt(dx*dx + dy*dy);
        if (dist < 90) {
          ctx.beginPath();
          ctx.moveTo(dots[i].x * W, dots[i].y * H);
          ctx.lineTo(dots[j].x * W, dots[j].y * H);
          ctx.strokeStyle = color;
          ctx.globalAlpha = 0.12 * (1 - dist / 90);
          ctx.lineWidth = 0.5;
          ctx.stroke();
          ctx.globalAlpha = 1;
        }
      }
    }
    requestAnimationFrame(tick);
  }
  tick();
}

/* Flowing diagonal lines */
function cardAnimLines(canvas, color) {
  const card = canvas.parentElement;
  function resize() {
    canvas.width  = card.offsetWidth;
    canvas.height = card.offsetHeight;
  }
  resize();
  new ResizeObserver(resize).observe(card);

  const ctx = canvas.getContext("2d");
  const lines = Array.from({ length: 6 }, (_, i) => ({
    offset: (i / 6),
    speed:  0.0004 + Math.random() * 0.0003,
    width:  0.5 + Math.random() * 0.5,
    alpha:  0.15 + Math.random() * 0.2,
  }));

  function tick() {
    const W = canvas.width, H = canvas.height;
    ctx.clearRect(0, 0, W, H);
    lines.forEach(l => {
      l.offset = (l.offset + l.speed) % 1;
      const x = l.offset * (W + H) - H;
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x + H, H);
      ctx.strokeStyle = color;
      ctx.globalAlpha = l.alpha;
      ctx.lineWidth   = l.width;
      ctx.stroke();
      ctx.globalAlpha = 1;
    });
    requestAnimationFrame(tick);
  }
  tick();
}

/* Inventory card — falling particles like stock dropping into bins */
function cardAnimInventory(canvas, color) {
  const card = canvas.parentElement;
  function resize() {
    canvas.width  = card.offsetWidth;
    canvas.height = card.offsetHeight;
  }
  resize();
  new ResizeObserver(resize).observe(card);

  const ctx = canvas.getContext("2d");
  const cols = 6;
  const particles = Array.from({ length: 24 }, (_, i) => ({
    col:   i % cols,
    y:     Math.random(),
    speed: 0.0008 + Math.random() * 0.001,
    size:  2 + Math.random() * 3,
    alpha: 0.15 + Math.random() * 0.3,
  }));

  function tick() {
    const W = canvas.width, H = canvas.height;
    ctx.clearRect(0, 0, W, H);
    const colW = W / cols;

    /* faint column guides */
    for (let i = 1; i < cols; i++) {
      ctx.beginPath();
      ctx.moveTo(i * colW, 0);
      ctx.lineTo(i * colW, H);
      ctx.strokeStyle = color;
      ctx.globalAlpha = 0.04;
      ctx.lineWidth = 1;
      ctx.stroke();
      ctx.globalAlpha = 1;
    }

    particles.forEach(p => {
      const px = (p.col + 0.5) * colW;
      p.y += p.speed;
      if (p.y > 1.05) { p.y = -0.05; p.alpha = 0.1 + Math.random() * 0.3; }
      ctx.beginPath();
      ctx.arc(px, p.y * H, p.size, 0, Math.PI * 2);
      ctx.fillStyle = color;
      ctx.globalAlpha = p.alpha;
      ctx.fill();
      ctx.globalAlpha = 1;
    });

    requestAnimationFrame(tick);
  }
  tick();
}

function countUp(id, target, duration, suffix) {
  const el = document.getElementById(id);
  if (!el || !target) return;
  const start = performance.now();
  function easeOut(t) { return 1 - Math.pow(1 - t, 3); }
  function tick(now) {
    const p = Math.min((now - start) / duration, 1);
    const val = Math.round(easeOut(p) * target);
    el.textContent = val.toLocaleString() + (suffix || "");
    if (p < 1) requestAnimationFrame(tick);
  }
  requestAnimationFrame(tick);
}

/* =====================================================
   PAGE LOAD
===================================================== */
document.addEventListener("DOMContentLoaded", () => {

  console.time("TOTAL_DASHBOARD_LOAD");

  initClock();
  initTabs();
  initUserDisplay();
  updateScannerFromStorage();
  initNavAnimations();
  initBrandCanvas();
  initSparklines();
  initCardCanvases();
  animateScannerHealth("Online", "cb-m-val--green");

  // Show the UI immediately — don't wait for API
  unlockPage();

  const username = sessionStorage.getItem("lms_user");

  // Data loads in background and fills in when ready
  Promise.all([
    fetchDashboard(),
    fetchVisibility(username)
  ])
  .then(([dashboardData, visibilityData]) => {
    renderDashboard(dashboardData);
    applyVisibilityRules(visibilityData.visibility || {});
    if (window.dismissLoader) window.dismissLoader();
  })
  .catch(err => {
    console.error("Dashboard load error", err);
    showErrorState();
    if (window.dismissLoader) window.dismissLoader();
  })
  .finally(() => {
    console.timeEnd("TOTAL_DASHBOARD_LOAD");
  });

});

/* =====================================================
   LIVE CLOCK + DATE
===================================================== */
function initClock() {
  function tick() {
    const now = new Date();
    const timeEl = document.getElementById("topbarTime");
    const dateEl = document.getElementById("topbarDate");
    if (timeEl) timeEl.textContent = now.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
    if (dateEl) dateEl.textContent = now.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric", year: "numeric" });
  }
  tick();
  setInterval(tick, 1000);
}

/* =====================================================
   BRAND CANVAS — animated bar chart behind Zenni LMS
===================================================== */
function initBrandCanvas() {
  const canvas = document.getElementById("cb-brand-canvas");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  const W = canvas.width, H = canvas.height;
  const bars = Array.from({ length: 18 }, (_, i) => ({
    x: 4 + i * 12,
    phase: Math.random() * Math.PI * 2,
    speed: 0.025 + Math.random() * 0.02,
    base: 8 + Math.random() * 10,
  }));
  function draw() {
    ctx.clearRect(0, 0, W, H);
    bars.forEach(b => {
      b.phase += b.speed;
      const h = b.base + Math.abs(Math.sin(b.phase)) * (H * 0.55);
      const alpha = 0.05 + 0.09 * Math.abs(Math.sin(b.phase));
      ctx.fillStyle = `rgba(245,200,66,${alpha})`;
      ctx.beginPath();
      ctx.roundRect(b.x, H - h, 7, h, 2);
      ctx.fill();
    });
    requestAnimationFrame(draw);
  }
  draw();
}

/* =====================================================
   SPARKLINE ANIMATIONS — live flowing lines
===================================================== */
function initSparklines() {
  setTimeout(() => {
    liveSparkline("spark-incoming", "#f5c842", 60, 100, 1.2);
    liveSparkline("spark-holds",    "#64a0dc", 0,  20,  0.8);
    liveSparkline("spark-scanner",  "#3fd189", 88, 100, 0.5);
  }, 120);
}

function liveSparkline(id, color, minVal, maxVal, speed) {
  const canvas = document.getElementById(id);
  if (!canvas) return;

  const W = canvas.parentElement.offsetWidth || 200;
  const H = 28;
  canvas.width  = W;
  canvas.height = H;
  canvas.dataset.color = color; /* allow runtime recolor */

  const ctx    = canvas.getContext("2d");
  const points = 40;
  const range  = maxVal - minVal || 1;
  const pad    = 3;

  let last = minVal + Math.random() * range;
  const data = Array.from({ length: points }, () => {
    last = Math.min(maxVal, Math.max(minVal, last + (Math.random() - 0.48) * range * 0.18));
    return last;
  });

  let offset = 0;

  function getY(v) {
    return H - pad - ((v - minVal) / range) * (H - pad * 2);
  }

  function tick() {
    const c = canvas.dataset.color || color; /* live color swap */

    offset += speed;
    if (offset >= W / (points - 1)) {
      offset = 0;
      last = Math.min(maxVal, Math.max(minVal,
        last + (Math.random() - 0.48) * range * 0.22));
      data.shift();
      data.push(last);
    }

    ctx.clearRect(0, 0, W, H);
    const step = W / (points - 1);

    ctx.beginPath();
    ctx.moveTo(-offset, getY(data[0]));
    for (let i = 1; i < points; i++) ctx.lineTo(i * step - offset, getY(data[i]));
    ctx.lineTo((points - 1) * step - offset, H);
    ctx.lineTo(-offset, H);
    ctx.closePath();
    ctx.fillStyle = c + "18";
    ctx.fill();

    ctx.beginPath();
    ctx.moveTo(-offset, getY(data[0]));
    for (let i = 1; i < points; i++) {
      const x  = i * step - offset;
      const xp = (i - 1) * step - offset;
      const cp = xp + step * 0.5;
      ctx.bezierCurveTo(cp, getY(data[i-1]), cp, getY(data[i]), x, getY(data[i]));
    }
    ctx.strokeStyle = c;
    ctx.lineWidth   = 1.5;
    ctx.lineJoin    = "round";
    ctx.stroke();

    const tipX = (points - 1) * step - offset;
    const tipY = getY(data[points - 1]);
    const pulse = 0.5 + 0.5 * Math.sin(Date.now() * 0.005);
    ctx.beginPath();
    ctx.arc(tipX, tipY, 2 + pulse * 1.5, 0, Math.PI * 2);
    ctx.fillStyle   = c;
    ctx.globalAlpha = 0.4 + 0.6 * pulse;
    ctx.fill();
    ctx.globalAlpha = 1;

    requestAnimationFrame(tick);
  }

  tick();
}

/* =====================================================
   SCANNER HEALTH TYPE-ON ANIMATION
===================================================== */
function animateScannerHealth(text, colorClass) {
  const el = document.getElementById("scannerHealth");
  if (!el) return;
  el.textContent = "";
  el.className = "cb-m-val " + colorClass;
  let i = 0;
  const iv = setInterval(() => {
    el.textContent = text.slice(0, ++i);
    if (i >= text.length) clearInterval(iv);
  }, 80);
}

/* =====================================================
   USER DISPLAY
===================================================== */
function initUserDisplay() {
  const username  = sessionStorage.getItem("lms_user") || "";
  const avatarEl  = document.getElementById("userAvatar");
  const nameEl    = document.getElementById("userName");

  if (!username) return;

  const initials = username
    .split(/[\s._-]+/)
    .slice(0, 2)
    .map(w => w[0])
    .join("")
    .toUpperCase();

  if (avatarEl) avatarEl.textContent = initials || username[0]?.toUpperCase() || "U";
  if (nameEl)   nameEl.textContent   = username;
}

/* =====================================================
   TAB SYSTEM
===================================================== */
const TAB_ANIMS = {
  operations: "anim-slide-up",
  production: "anim-slide-left",
  system:     "anim-scale-up",
  inventory:  "anim-slide-up",
};

function animateTab(tabId) {
  const section = document.getElementById(tabId);
  if (!section) return;

  const animClass = TAB_ANIMS[tabId] || "anim-slide-up";

  /* Heading */
  const head = section.querySelector(".module-head");
  if (head) {
    head.classList.remove("anim-head");
    void head.offsetWidth;
    head.classList.add("anim-head");
  }

  /* Cards — only animate visible ones, staggered */
  const cards = [...section.querySelectorAll(".card")].filter(
    c => c.style.display !== "none" && getComputedStyle(c).display !== "none"
  );

  cards.forEach((card, i) => {
    /* strip old anim classes and force reflow */
    card.classList.remove("anim-slide-up", "anim-slide-left", "anim-scale-up");
    void card.offsetWidth;
    card.style.animationDelay = (i * 0.10) + "s";
    card.classList.add(animClass);
  });

  /* Tags — staggered pop */
  section.querySelectorAll(".card-tags .tag").forEach((tag, i) => {
    tag.classList.remove("tag-pop");
    void tag.offsetWidth;
    tag.style.animationDelay = (0.28 + i * 0.06) + "s";
    tag.classList.add("tag-pop");
  });

  /* Metric values */
  section.querySelectorAll(".metric-val").forEach((el, i) => {
    el.classList.remove("metric-flash");
    void el.offsetWidth;
    el.style.animationDelay = (0.32 + i * 0.07) + "s";
    el.classList.add("metric-flash");
  });
}

function initTabs() {
  const tabs     = document.querySelectorAll(".nav-item[data-tab]");
  const contents = document.querySelectorAll(".tab-content");

  /* Animate the default active tab on load */
  const defaultActive = document.querySelector(".tab-content.active");
  if (defaultActive) setTimeout(() => animateTab(defaultActive.id), 80);

  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(t     => t.classList.remove("active"));
      contents.forEach(c => c.classList.remove("active"));

      tab.classList.add("active");
      const target = document.getElementById(tab.dataset.tab);
      if (!target) return;
      target.classList.add("active");
      animateTab(tab.dataset.tab);
    });
  });
}

/* =====================================================
   FETCH DASHBOARD
===================================================== */
async function fetchDashboard() {
  console.time("DASHBOARD_API");
  try {
    const res  = await fetch(API_URL + "?t=" + Date.now());
    const data = await res.json();
    console.timeEnd("DASHBOARD_API");
    return data || {};
  } catch (err) {
    console.error("Dashboard API failed", err);
    return {};
  }
}

/* =====================================================
   FETCH VISIBILITY
===================================================== */
async function fetchVisibility(username) {
  console.time("VISIBILITY_API");
  if (!username) return { status: "SUCCESS", visibility: {} };
  try {
    const res  = await fetch(`${AUTH_API}?action=visibility&username=${encodeURIComponent(username)}`);
    const data = await res.json();
    console.timeEnd("VISIBILITY_API");
    return data || { status: "SUCCESS", visibility: {} };
  } catch (err) {
    console.error("Visibility API failed", err);
    return { status: "SUCCESS", visibility: {} };
  }
}

/* =====================================================
   THRESHOLD COLOR HELPERS
===================================================== */
function setHoldsColor(active) {
  const el = document.getElementById("active");
  if (!el) return;
  if (active >= 15) {
    el.className = "cb-m-val cb-m-val--red";
  } else if (active >= 7) {
    el.className = "cb-m-val cb-m-val--orange";
  } else {
    el.className = "cb-m-val cb-m-val--steel";
  }
  /* also recolor the sparkline */
  const color = active >= 15 ? "#e03e3e" : active >= 7 ? "#f07030" : "#64a0dc";
  updateSparklineColor("spark-holds", color);
}

function updateSparklineColor(id, color) {
  /* store on canvas element so the live loop picks it up */
  const c = document.getElementById(id);
  if (c) c.dataset.color = color;
}
function renderDashboard(data) {
  const active    = Number(data.active    ?? 0);
  const completed = Number(data.completed ?? 0);

  // Total Incoming — shown in command bar KPI
  const total = Number(data.totalIncoming ?? data.incoming ?? 0);

  let coverage = Number(data.coverage ?? 0);
  if (coverage <= 1) coverage *= 100;
  coverage = Math.max(0, Math.min(100, Math.round(coverage)));
  if (completed > 0 && coverage === 0) {
    coverage = Math.round((completed / (completed + active)) * 100);
  }

  // Command bar KPIs — animated count-up
  if (total)  countUp("totalIncoming", total,  1800, "");
  else        setText("totalIncoming", "—");
  if (active) countUp("active",        active, 1400, "");
  else        setText("active",        active);

  setHoldsColor(active);

  // Card metrics (investigation card)
  setText("activeHolds",   active);
  setText("completed",     completed);
  setText("coverageDetail", active === 0 ? "Complete" : coverage + "%");
}

/* =====================================================
   VISIBILITY RULES
===================================================== */
function applyVisibilityRules(visibility) {
  const map = {
    "Scanner Map":        "#scanner-card",
    "Investigation Hold": "#investigation-card",
    "True Curve":         "#truecurve-card",
    "Tools":              "#tools-card",
    "Coating":            "#coating-card",
    "Incoming Jobs":      "#incoming-card",
    "Admin":              ".admin-only"
  };

  // Hide all gated cards first
  Object.values(map).forEach(sel => {
    document.querySelectorAll(sel).forEach(el => el.style.display = "none");
  });

  // Show permitted cards
  Object.keys(map).forEach(feature => {
    if (visibility[feature]) {
      document.querySelectorAll(map[feature]).forEach(el => el.style.display = "flex");
    }
  });

  // Power Analysis Tool
  if (visibility["Coating"] || visibility["Production"]) {
    const pw = document.getElementById("power-card");
    if (pw) pw.style.display = "flex";
  }

  // Re-run tab animations NOW that cards are visible
  const activeSection = document.querySelector(".tab-content.active");
  if (activeSection) {
    setTimeout(() => animateTab(activeSection.id), 60);
  }
}

/* =====================================================
   SCANNER STATUS (localStorage)
===================================================== */
function updateScannerFromStorage() {
  let scanners = [];
  try {
    scanners = JSON.parse(localStorage.getItem(SCANNER_STORAGE_KEY)) || [];
  } catch {}

  const sub    = document.getElementById("scannerSub");
  const detail = document.getElementById("scannerDetail");

  if (!scanners.length) {
    animateScannerHealth("Online", "cb-m-val--green");
    if (sub)    sub.textContent = "All ports";
    if (detail) detail.innerHTML = "";
    updateSparklineColor("spark-scanner", "#3fd189");
  } else if (scanners.length <= 2) {
    animateScannerHealth("Attention", "cb-m-val--orange");
    if (sub)    sub.textContent = `${scanners.length} flagged`;
    if (detail) detail.innerHTML = `<strong>Flagged Scanners:</strong><ul>${scanners.map(s => `<li>${s}</li>`).join("")}</ul>`;
    updateSparklineColor("spark-scanner", "#f07030");
  } else {
    animateScannerHealth("Critical", "cb-m-val--red");
    if (sub)    sub.textContent = `${scanners.length} flagged`;
    if (detail) detail.innerHTML = `<strong>Flagged Scanners:</strong><ul>${scanners.map(s => `<li>${s}</li>`).join("")}</ul>`;
    updateSparklineColor("spark-scanner", "#e03e3e");
  }
}

/* =====================================================
   ERROR STATE
===================================================== */
function showErrorState() {
  ["totalIncoming","active","activeHolds","completed","coverageDetail","scannerHealth"]
    .forEach(id => setText(id, "ERR"));
}

/* =====================================================
   UNLOCK PAGE
===================================================== */
function unlockPage() {
  document.body.classList.remove("lms-hidden");
}

/* =====================================================
   LOGOUT
===================================================== */
function logout() {
  sessionStorage.clear();
  localStorage.removeItem(SCANNER_STORAGE_KEY);
  window.location.href = "login.html";
}

/* =====================================================
   HELPERS
===================================================== */
function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function openPage(page) {
  window.location.href = page;
}