/**************************************************
 * LMS GLOBAL AUTH GUARD
 * File name in project: authGuard.js
 *
 * Purpose:
 * - Forces login before any protected dashboard page opens.
 * - Put this script FIRST inside <head> on every dashboard page.
 * - Do NOT add this file to login.html.
 **************************************************/

(function () {
  const path = String(location.pathname || "").toLowerCase();

  // Do not block the login page itself.
  if (path.includes("login")) return;

  const loggedIn =
    sessionStorage.getItem("lms_logged_in") === "true";

  const username =
    sessionStorage.getItem("lms_user") ||
    sessionStorage.getItem("lms_username") ||
    "";

  const visibility =
    sessionStorage.getItem("lms_visibility") || "";

  if (!loggedIn || !username || !visibility) {
    const currentPage =
      (location.pathname.split("/").pop() || "index.html") + location.search;

    sessionStorage.clear();
    sessionStorage.setItem("lms_redirect_after_login", currentPage);

    window.location.replace("login.html");
  }
})();
