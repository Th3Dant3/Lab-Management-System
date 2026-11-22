/* Zenni LMS V6.3 — direct Google Apps Script popup (no iframe bridge).
 * This UI return signal NEVER grants access; secure index is checked by Apps Script.
 */
'use strict';
const QA_V6_POPUP_URL = 'https://script.google.com/a/macros/zennioptical.com/s/AKfycbzqOAt8pc3gHT3vcnk6G-P3HuxkTHiV4ZrB5u7Gmr9oLPoI-8S2aixMfBWEZa4fk00y/exec?page=authPopup';
const QA_V6_RETURN_PATH = './index.html';
const QA_V6_ORIGIN = window.location.origin;
const QA_V6_BUTTON = document.getElementById('loginBtn');
const QA_V6_LABEL = document.getElementById('loginBtnText');
let qaV6Popup = null;
const qaV6Status = document.createElement('div');
qaV6Status.setAttribute('role', 'status');
qaV6Status.setAttribute('aria-live', 'polite');
qaV6Status.style.cssText = 'margin-top:14px;color:#b7d3e8;font:13px/1.5 system-ui;';
QA_V6_BUTTON?.insertAdjacentElement('afterend', qaV6Status);
function qaV6SetStatus(message) { qaV6Status.textContent = message; }
function qaV6ShowLoginLabel() { if (QA_V6_LABEL) QA_V6_LABEL.textContent = 'SIGN IN WITH GOOGLE'; }
function qaV6Open() {
  if (qaV6Popup && !qaV6Popup.closed) {
    qaV6Popup.focus();
    return;
  }
  const mobile = window.matchMedia('(max-width: 720px)').matches;
  const width = Math.min(560, screen.availWidth || 560);
  const height = Math.min(740, screen.availHeight || 740);
  const left = Math.max(0, Math.floor(((screen.availWidth || width) - width) / 2));
  const top = Math.max(0, Math.floor(((screen.availHeight || height) - height) / 2));
  // Called synchronously from user click, so normal popup-blocking rules apply.
  qaV6Popup = window.open(
    QA_V6_POPUP_URL, 'ZenniLMSGoogleSignIn',
    mobile ? undefined : `popup=yes,width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes`
  );
  if (!qaV6Popup) {
    qaV6SetStatus('Allow pop-ups for this page, then select Sign in with Google again.');
    qaV6ShowLoginLabel();
    return;
  }
  if (QA_V6_LABEL) QA_V6_LABEL.textContent = 'GOOGLE SIGN-IN OPEN';
  qaV6SetStatus('Finish Google verification in the new window.');
  try { qaV6Popup.focus(); } catch (_) {}
}
QA_V6_BUTTON?.addEventListener('click', qaV6Open);
window.addEventListener('message', (event) => {
  // Only the same-origin return page in the popup may signal navigation.
  if (event.origin !== QA_V6_ORIGIN || event.source !== qaV6Popup) return;
  const data = event.data;
  if (!data || data.product !== 'ZENNI_LMS_V6' || data.type !== 'RETURN_TO_LMS') return;
  qaV6SetStatus('Opening LMS. Google will check access again.');
  try { qaV6Popup.close(); } catch (_) {}
  // A message from any public page is NOT authentication. The Apps Script iframe
  // in index.html independently checks the visitor before displaying content.
  window.location.assign(QA_V6_RETURN_PATH);
});
window.addEventListener('pageshow', qaV6ShowLoginLabel);
