/**************************************************
 * CONFIG
 **************************************************/
const AUTH_API =
"https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

/**************************************************
 * INIT
 **************************************************/
document.addEventListener("DOMContentLoaded", () => {
  const userEl = document.getElementById("username");
  if (userEl) {
    userEl.addEventListener("input", () => {
      userEl.value = userEl.value.toUpperCase();
    });
  }
});

/**************************************************
 * UI HELPERS
 **************************************************/
function setMessage(text, type = "error") {
  const el = document.getElementById("message");
  el.textContent = text;
  el.className = "login-message " + (type || "");
}

function setLoading(isLoading) {
  const btn    = document.getElementById("loginBtn");
  const prog   = document.getElementById("progressBar");
  const progWrap = document.getElementById("progressWrap");

  if (!btn) return;

  if (isLoading) {
    btn.disabled = true;
    btn.innerHTML = '<span class="btn-spinner"></span> Authenticating...';
    if (progWrap) progWrap.classList.add("active");
    if (prog)     prog.style.width = "40%";
  } else {
    btn.disabled = false;
    btn.innerHTML = 'Sign In';
    if (progWrap) progWrap.classList.remove("active");
    if (prog)     prog.style.width = "0%";
  }
}

function setProgress(pct) {
  const prog = document.getElementById("progressBar");
  if (prog) prog.style.width = pct + "%";
}

/**************************************************
 * LOGIN  (single API call — login + visibility merged)
 **************************************************/
async function login() {

  console.time("LOGIN_TOTAL");

  const usernameEl = document.getElementById("username");
  const passwordEl = document.getElementById("password");

  const username = usernameEl.value.trim().toUpperCase();
  const password = passwordEl.value.trim();

  setMessage("");
  setLoading(true);

  if (!username || !password) {
    setMessage("Enter username and password", "error");
    setLoading(false);
    return;
  }

  // ── Single call: action=loginFull returns auth + visibility together ──
  const url =
    `${AUTH_API}?action=loginFull` +
    `&username=${encodeURIComponent(username)}` +
    `&password=${encodeURIComponent(password)}`;

  console.time("LOGIN_API");

  try {
    const res  = await fetch(url);
    const data = await res.json();

    console.timeEnd("LOGIN_API");

    setProgress(80);
    await handleLoginResponse(data, password);

  } catch (err) {
    setMessage("Connection error. Try again.", "error");
  }

  setLoading(false);
  console.timeEnd("LOGIN_TOTAL");
}

/**************************************************
 * HANDLE RESPONSE
 **************************************************/
async function handleLoginResponse(data, originalPassword) {

  if (!data || !data.status) {
    setMessage("Invalid server response", "error");
    return;
  }

  if (data.status === "ERROR") {
    setMessage(data.message || "Login failed", "error");
    return;
  }

  if (data.status === "SET_PASSWORD_REQUIRED") {
    setMessage("First login — setting password…", "warning");
    await setPassword(data.username, originalPassword);
    return;
  }

  if (data.status === "SUCCESS") {
    setMessage("Access Granted", "success");
    setProgress(100);

    sessionStorage.setItem("lms_logged_in",  "true");
    sessionStorage.setItem("lms_user",        data.username);
    sessionStorage.setItem("lms_role",        data.role    || "");
    sessionStorage.setItem("lms_subrole",     data.subRole || "");

    // visibility already bundled in the same response — no second call needed
    sessionStorage.setItem(
      "lms_visibility",
      JSON.stringify(data.visibility || {})
    );

    setTimeout(() => {
      window.location.replace("index.html");
    }, 350);
  }
}

/**************************************************
 * SET PASSWORD  (first-login flow — still separate call)
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
      setMessage("Password setup failed", "error");
      return;
    }

    // After setting password call loginFull to get everything in one shot
    setMessage("Password set — signing in…", "warning");
    setProgress(60);

    const fullUrl =
      `${AUTH_API}?action=loginFull` +
      `&username=${encodeURIComponent(username)}` +
      `&password=${encodeURIComponent(password)}`;

    const res2  = await fetch(fullUrl);
    const data2 = await res2.json();

    await handleLoginResponse(data2, password);

  } catch (err) {
    setMessage("Password setup failed", "error");
  }
}

/**************************************************
 * ENTER KEY
 **************************************************/
document.addEventListener("keydown", e => {
  if (e.key === "Enter") login();
});