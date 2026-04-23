const AUTH_API =
  "https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

/**************************************************
 * LOADER MAP — users with a personal loader go there first
 * Anyone not listed falls back to index.html
 **************************************************/
const LOADER_MAP = {
  "BLOPEZ":        "loader_BLOPEZ.html",
  "MLITTLE":       "loader_MLITTLE.html",
  "JBOOMERSHINE":  "loader_JBOOMERSHINE.html",
  "BKARR":         "loader_BKARR.html",
  "RTATE":         "loader_RTATE.html",
  "AIVANOVSKI":    "loader_AIVANOVSKI.html",
  "SANDERSON":     "loader_SANDERSON.html"
};
document.addEventListener("DOMContentLoaded", () => {
  const userEl     = document.getElementById("username");
  const passwordEl = document.getElementById("password");
  const toggleBtn  = document.getElementById("togglePassword");

  if (userEl) {
    userEl.addEventListener("input", () => {
      userEl.value = userEl.value.toUpperCase();
    });
  }

  if (toggleBtn && passwordEl) {
    toggleBtn.addEventListener("click", () => {
      const isPassword = passwordEl.type === "password";
      passwordEl.type  = isPassword ? "text" : "password";
      toggleBtn.textContent = isPassword ? "Hide" : "Show";
    });
  }

  document.addEventListener("keydown", e => {
    if (e.key === "Enter") login();
  });
});

/**************************************************
 * UI HELPERS
 **************************************************/
function setMessage(text, type = "warning") {
  const wrap = document.getElementById("messageWrap");
  const el   = document.getElementById("message");
  if (!wrap || !el) return;
  wrap.classList.remove("warning", "success", "error");
  wrap.classList.add(type);
  el.textContent = text;
}

function setLoading(isLoading) {
  const btn     = document.getElementById("loginBtn");
  const btnText = document.getElementById("loginBtnText");
  if (!btn || !btnText) return;
  btn.disabled        = isLoading;
  btnText.textContent = isLoading ? "Signing In..." : "Sign In";
}

function setProgress(pct) {
  const prog = document.getElementById("progressBar");
  if (prog) prog.style.width = `${pct}%`;
}

/**************************************************
 * LOGIN
 **************************************************/
async function login() {
  const usernameEl = document.getElementById("username");
  const passwordEl = document.getElementById("password");

  const username = usernameEl?.value.trim().toUpperCase();
  const password = passwordEl?.value.trim();

  if (!username || !password) {
    setMessage("Enter username and password.", "error");
    setProgress(0);
    return;
  }

  setLoading(true);
  setMessage("Authenticating secure session...", "warning");
  setProgress(25);

  const url =
    `${AUTH_API}?action=loginFull` +
    `&username=${encodeURIComponent(username)}` +
    `&password=${encodeURIComponent(password)}`;

  try {
    const res  = await fetch(url);
    const data = await res.json();
    setProgress(70);
    await handleLoginResponse(data, password);
  } catch (err) {
    console.error("Login error:", err);
    setMessage("Connection error. Please try again.", "error");
    setProgress(0);
  } finally {
    setLoading(false);
  }
}

/**************************************************
 * HANDLE RESPONSE
 **************************************************/
async function handleLoginResponse(data, originalPassword) {
  if (!data || !data.status) {
    setMessage("Invalid server response.", "error");
    setProgress(0);
    return;
  }

  if (data.status === "ERROR") {
    setMessage(data.message || "Login failed.", "error");
    setProgress(0);
    return;
  }

  if (data.status === "SET_PASSWORD_REQUIRED") {
    setMessage("First login detected. Setting password...", "warning");
    setProgress(55);
    await setPassword(data.username, originalPassword);
    return;
  }

  if (data.status === "SUCCESS") {
    const fullName = ((data.firstName || "") + " " + (data.lastName || "")).trim()
                     || data.username;

    // ── Visibility: new structure has departments + features ──
    const visibility   = data.visibility || {};
    const departments  = visibility.departments || {};
    const features     = visibility.features    || {};

    // Core session
    sessionStorage.setItem("lms_logged_in",   "true");
    sessionStorage.setItem("lms_user",         data.username);
    sessionStorage.setItem("lms_role",         data.role    || "");
    sessionStorage.setItem("lms_subrole",      data.subRole || "");
    sessionStorage.setItem("lms_fullname",     fullName);
    sessionStorage.setItem("lms_firstname",    data.firstName || "");
    sessionStorage.setItem("lms_lastname",     data.lastName  || "");

    // Visibility — stored separately for easy access on dashboard
    sessionStorage.setItem("lms_visibility",   JSON.stringify(visibility));
    sessionStorage.setItem("lms_departments",  JSON.stringify(departments));
    sessionStorage.setItem("lms_features",     JSON.stringify(features));

    setProgress(100);
    setMessage("Access granted. Redirecting...", "success");

    // Route to personal loader if one exists, otherwise index.html
    const loader = LOADER_MAP[data.username] || "index.html";
    setTimeout(() => {
      window.location.replace(loader);
    }, 400);
  }
}

/**************************************************
 * SET PASSWORD (first login)
 **************************************************/
async function setPassword(username, password) {
  const url =
    `${AUTH_API}?action=setPassword` +
    `&username=${encodeURIComponent(username)}` +
    `&password=${encodeURIComponent(password)}`;

  try {
    const res  = await fetch(url);
    const data = await res.json();

    if (data.status !== "PASSWORD_SET") {
      setMessage("Password setup failed.", "error");
      setProgress(0);
      return;
    }

    setMessage("Password set. Signing in...", "warning");
    setProgress(75);

    const fullUrl =
      `${AUTH_API}?action=loginFull` +
      `&username=${encodeURIComponent(username)}` +
      `&password=${encodeURIComponent(password)}`;

    const res2  = await fetch(fullUrl);
    const data2 = await res2.json();
    await handleLoginResponse(data2, password);

  } catch (err) {
    console.error("Password setup error:", err);
    setMessage("Password setup failed.", "error");
    setProgress(0);
  }
}

/**************************************************
 * VISIBILITY HELPERS
 * Use these anywhere in your dashboard JS
 *
 * canAccessDept("Production")
 * canAccessFeature("Production_SurfaceWorkflow")
 **************************************************/
function canAccessDept(deptKey) {
  const depts = JSON.parse(sessionStorage.getItem("lms_departments") || "{}");
  return depts[deptKey] === true;
}

function canAccessFeature(featureKey) {
  const feats = JSON.parse(sessionStorage.getItem("lms_features") || "{}");
  return feats[featureKey] === true;
}

function applyVisibility() {
  // Hide sidebar sections where dept gate = false
  // Requires: data-dept="Production" on sidebar elements
  document.querySelectorAll("[data-dept]").forEach(el => {
    const dept = el.getAttribute("data-dept");
    el.style.display = canAccessDept(dept) ? "" : "none";
  });

  // Hide tabs/features where feature = false
  // Requires: data-feature="Production_SurfaceWorkflow" on tab elements
  document.querySelectorAll("[data-feature]").forEach(el => {
    const key = el.getAttribute("data-feature");
    el.style.display = canAccessFeature(key) ? "" : "none";
  });
}

function requireAuth() {
  if (sessionStorage.getItem("lms_logged_in") !== "true") {
    window.location.replace("login.html");
    return false;
  }
  return true;
}