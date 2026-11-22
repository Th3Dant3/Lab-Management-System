ZENNI LMS V6 — GITHUB INDEX LAUNCHER TEST
===========================================
GOAL: Keep browser URL at /Lab-Management-System/v6-test/index.html while protected LMS runs inside Apps Script iframe.

UPLOAD THESE FILES TO GitHub repository Lab-Management-System, folder v6-test/:
  index.html                 NEW full-screen iframe launcher
  login.html                 optional branded login launcher
  login.css
  login-v6.js               UPDATED login button -> ./index.html
  LmsV6App.html             DISTINCT authenticated dashboard source (MUST NOT REPLACE WITH LAUNCHER)
  index.js                  Dashboard script loaded by secure Apps Script page
  style-index.css
  v6-auth-visibility.js

COPY the other actual dashboard HTML/CSS/JS assets to v6-test/ preserving exact relative paths, especially FacilityWIP.html and its dependencies. Do NOT replace the separate secure dashboard HTML (LmsV6App.html) with the small index.html launcher.

NO GOOGLE APPS SCRIPT CHANGES REQUIRED if the already-working /exec?page=lmsV6 uses:
 HTML_URL: https://raw.githubusercontent.com/th3dant3/Lab-Management-System/main/v6-test/LmsV6App.html
 ASSET_BASE: https://th3dant3.github.io/Lab-Management-System/v6-test/

TEST:
  1. Commit the files to GitHub, root files unchanged.
  2. Open https://th3dant3.github.io/Lab-Management-System/v6-test/login.html
  3. Click OPEN SECURE LMS V6; URL should become .../v6-test/index.html
  4. The iframe should display the Google-secured dashboard. Verify card navigation and Facility WIP.
  5. If iframe blank or sign-in blocked, click direct secure page link. This indicates Google browser framing/login constraints; do NOT change the deployment to Anyone.
  6. Test as a second Zenni employee and confirm that permission-denied cards do not show. Separate apps must enforce own server-side permission checks.

IMPORTANT SECURITY / SCOPE:
- This is STAGING only. Root production index.html and login.html remain unchanged.
- The GitHub launcher is public static HTML; confidential data must only be served by Google-restricted Apps Script.
- The live header KPI data remains disabled in V6 staging.
- Do not put Code.gs or identity test HTML into GitHub. Those belong to Apps Script.
- No local mock/DEMO login is used in this package.
- GitHub Pages cannot detect whether authenticated Google iframe succeeded: direct link is the fallback.
- For moving to root after successful test, update Code.gs HTML_URL and ASSET_BASE to final production paths, and keep source/launcher filenames separate. Do not just copy index.html to root without updating those references.
