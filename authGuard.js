(() => {
  "use strict";

  /*
   * Change this value whenever every existing login session
   * must be invalidated after a deployment.
   */
  const AUTH_VERSION = "2026-07-12-main-rebuild";

  /*
   * Resolve the repository root from the physical location of
   * authGuard.js. This works from root pages and nested folders.
   */
  const guardScript =
    document.currentScript ||
    Array.from(document.scripts).find(script =>
      String(script.src || "").includes("authGuard.js")
    );

  const ROOT_URL =
    guardScript && guardScript.src
      ? new URL(".", guardScript.src)
      : new URL("./", window.location.href);

  /*
   * Keep the original public login address working.
   * The root login.html loads /auth/login.css and /auth/login.js.
   */
  const LOGIN_URL = new URL("login.html", ROOT_URL).href;

  function getUsername() {
    return (
      sessionStorage.getItem("lms_user") ||
      sessionStorage.getItem("lms_username") ||
      ""
    ).trim();
  }

  function isLoginPage() {
    try {
      const currentPath = new URL(window.location.href).pathname;
      const rootLoginPath = new URL("login.html", ROOT_URL).pathname;
      const authLoginPath = new URL("auth/login.html", ROOT_URL).pathname;

      return (
        currentPath === rootLoginPath ||
        currentPath === authLoginPath
      );
    } catch (error) {
      return false;
    }
  }

  function isLoggedIn() {
    const loggedIn =
      sessionStorage.getItem("lms_logged_in") === "true";

    const username = getUsername();

    const visibility =
      sessionStorage.getItem("lms_visibility") || "";

    const savedVersion =
      sessionStorage.getItem("lms_auth_version") || "";

    return Boolean(
      loggedIn &&
      username &&
      visibility &&
      savedVersion === AUTH_VERSION
    );
  }

  function saveRequestedPage() {
    if (isLoginPage()) return;

    const currentPath =
      window.location.pathname +
      window.location.search +
      window.location.hash;

    sessionStorage.setItem(
      "lms_redirect_after_login",
      currentPath
    );
  }

  function clearAuthSession() {
    [
      "lms_logged_in",
      "lms_user",
      "lms_username",
      "lms_role",
      "lms_subrole",
      "lms_fullname",
      "lms_firstname",
      "lms_lastname",
      "lms_visibility",
      "lms_departments",
      "lms_features",
      "lms_auth_version",
      "lms_redirect_after_login"
    ].forEach(key => sessionStorage.removeItem(key));

    [
      "LMS_AUTH_USER",
      "LMS_USER",
      "LMS_USERNAME"
    ].forEach(key => localStorage.removeItem(key));
  }

  function redirectToLogin({ rememberPage = true } = {}) {
    if (isLoginPage()) return;

    if (rememberPage) {
      saveRequestedPage();
    }

    window.location.replace(LOGIN_URL);
  }

  /*
   * Anyone already logged in with the previous deployment will
   * not have the new matching lms_auth_version. On refresh or
   * navigation, their session is cleared and they return to login.
   */
  if (!isLoggedIn() && !isLoginPage()) {
    clearAuthSession();
    redirectToLogin({ rememberPage: false });
  }

  window.requireAuth = function requireAuth() {
    if (!isLoggedIn()) {
      clearAuthSession();
      redirectToLogin({ rememberPage: false });
      return false;
    }

    return true;
  };

  window.logout = function logout() {
    clearAuthSession();
    window.location.replace(LOGIN_URL);
  };

  window.logOut = window.logout;
})();