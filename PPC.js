/* PPC V21 hosting/embedding compatibility TEST ONLY — no spreadsheet data handled here. */
'use strict';
const PPC_SCRIPT_URL = 'https://script.google.com/a/macros/zennioptical.com/s/AKfycbybKUJ9OP708EbnplRJVrGPy9aenOEz07ZOWUEPJ8PrVCQprd6X2dFTffOxxpkXsqIiFg/exec';
const $ = id => document.getElementById(id);
function readJSON(key){try{return JSON.parse(sessionStorage.getItem(key)||'{}')||{};}catch(_){return {};}}
function deny(message,login){$('gate').hidden=true;$('app').hidden=true;$('denied').hidden=false;$('deniedText').textContent=message;$('returnLink').href=login?'login.html':'index.html';$('returnLink').textContent=login?'Go to LMS Login':'Return to LMS';}
function main(){
  if(sessionStorage.getItem('lms_logged_in')!=='true'){deny('Sign into LMS in this browser tab before opening PPC.',true);return;}
  let departments=readJSON('lms_departments'),features=readJSON('lms_features');
  if(!Object.keys(departments).length&&!Object.keys(features).length){const v=readJSON('lms_visibility');departments=v.departments||{};features=v.features||{};}
  if(departments.PPC_Document_Control!==true||features.PPC_Document_Control_ViewHub!==true){deny('Your LMS account does not have PPC Document Control / ViewHub permissions.',false);return;}
  // Browser gate for navigation only; actual records remain restricted by Google Workspace + PPC_USERS in Code.gs.
  $('account').textContent=sessionStorage.getItem('lms_user')||'LMS employee';
  $('gate').hidden=true;$('app').hidden=false;
  $('external').href=PPC_SCRIPT_URL;$('fallbackLink').href=PPC_SCRIPT_URL;
  $('ppcFrame').src=PPC_SCRIPT_URL;
  // Cross-origin iframe errors cannot reliably be detected in this page. Display a manual recovery hint.
  setTimeout(()=>{$('fallback').hidden=false;},12000);
  $('logout').addEventListener('click',()=>{['lms_logged_in','lms_user','lms_role','lms_visibility','lms_departments','lms_features'].forEach(k=>sessionStorage.removeItem(k));location.replace('login.html');});
}
document.addEventListener('DOMContentLoaded',main);
