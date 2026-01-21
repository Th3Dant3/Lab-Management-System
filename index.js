/* =====================================================
   LMS OPERATIONAL CONTROL CENTER — INDEX.JS
   ===================================================== */

const REFRESH_INTERVAL = 60000; // 2 minutes

document.addEventListener("DOMContentLoaded", () => {
  wireNavigation();
  updateLastRefresh();
  setInterval(updateLastRefresh, REFRESH_INTERVAL);
});

/* =====================================================
   NAVIGATION BUTTONS
   ===================================================== */

function wireNavigation() {
  const nav = {
    "btn-lens": "in8-in9.html",
    "btn-investigations": "page3.html",
    "btn-power": "true-curve.html"
  };

  Object.entries(nav).forEach(([id, target]) => {
    const btn = document.getElementById(id);
    if (!btn) return;

    btn.addEventListener("click", () => {
      window.location.href = target;
    });
  });
}

/* =====================================================
   LAST REFRESH DISPLAY
   ===================================================== */

function updateLastRefresh() {
  const el = document.getElementById("last-refresh");
  if (!el) return;

  el.textContent =
    "Last refresh: " + new Date().toLocaleTimeString();
}
