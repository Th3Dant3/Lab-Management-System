/** V6 STAGING: authoritative profile returned by google.script.run from the serving project.
 * Never read localStorage or sessionStorage for authorization decisions.
 * All independent backend data APIs still require separate server-side enforcement.
 */
'use strict';
const REAL_PAGE_MAP={
 PAGE_FACILITY_WIP:['facility',null],
 PAGE_INVESTIGATION_HOLD:['operations','investigation-card'],
 PAGE_LENS_INTAKE:['operations','tools-card'],
 PAGE_TRUE_CURVE:['operations','truecurve-card'],
 PAGE_SURFACE_DASHBOARD:['production','surfaceflow-card'],
 PAGE_AR_DASHBOARD:['production','ar-card'],
 PAGE_AR_RUN_TRACKER:['production','arruntracker-card'],
 PAGE_PRODUCTION_FORECAST:['production','productionforecast-card'],
 PAGE_FINISH_DASHBOARD:['production','finishdash-card'],
 PAGE_COATING_BREAKAGE:['production','coating-card'],
 PAGE_POWER_ANALYSIS:['production','power-card'],
 PAGE_INCOMING_JOBS:['inventory','incoming-card'],
 PAGE_PICKING_DASHBOARD:['inventory','pickwip-card'],
 PAGE_PRODUCTIVITY_HUB:['productivity','productivity-card'],
 PAGE_BREAKAGE_QUALITY_HUB:['productivity','breakage-quality-card'],
 PAGE_FACILITY_BREAKAGE:['training','facility-breakage-card'],
 PAGE_PPC_DOCUMENT_CONTROL:['training','ppc-document-control-card'],
 PAGE_SYSTEM:['system','scanner-card']};

const qaV6Ready=new Promise(resolve=>{
 document.addEventListener('DOMContentLoaded',()=>{
   if(!window.google?.script?.run){
     qaV6Deny('Apps Script hosting required. Open the secured V6 deployment, not LmsV6App.html directly.');
     resolve(false);return;
   }
   google.script.run.withSuccessHandler(profile=>{
     if(profile?.state!=='AUTHORIZED'||!profile.pages ||Object.keys(profile.pages).length!==18){
       qaV6Deny('Access denied: '+String(profile?.state||'INCOMPLETE_PERMISSIONS'));resolve(false);return;
     }
     try {
       window.__ZENNI_V6_USER__=profile;
       qaApplyV6Visibility();
       if(profile.canManageRequests===true)qaV6MountAdminAlerts();
       document.getElementById('v6-preflight')?.remove();
       resolve(true);
     } catch (err) {
       console.error('[V6] visibility failed',err);
       qaV6Deny('The LMS permission mapping failed. Contact the LMS administrator.');
       resolve(false);
     }
   }).withFailureHandler(err=>{
     qaV6Deny('Verification error. Contact your LMS administrator.');
     console.error('[V6] identity error',err);resolve(false);
   }).qaV6Bootstrap();
 },{once:true});
});
function qaV6Deny(message){
 document.getElementById('v6-preflight')?.remove();
 document.body.style.visibility='visible';
 document.body.innerHTML='<main style="padding:42px;color:#eef;background:#091321;min-height:100vh;font:16px system-ui"><h2>Zenni LMS — Access unavailable</h2><p id="qa-deny"></p></main>';
 document.getElementById('qa-deny').textContent=message;
}
function qaApplyV6Visibility(){
 const p=window.__ZENNI_V6_USER__;
 if(!p||!p.pages)throw new Error('V6 authorization has not completed');
 const pages=p.pages;
 for(const [key,[tab,card]] of Object.entries(REAL_PAGE_MAP)){
  const el=document.getElementById(card||tab);
  if(el){
   const enabled=pages[key]===true;
   // Most cards carry [data-gated] { display:none } in style-index.css.
   // An empty inline value does NOT override that rule, so explicitly use flex.
   el.style.display=enabled?(card?'flex':''):'none';
   el.dataset.qaAllowed=String(enabled);
  }
 }
 const permitted=[];
 // Diagnostic: distinguish real FALSE values from a CSS/DOM display failure.
 const statusRows=Object.entries(REAL_PAGE_MAP).map(([key,[tab,card]])=>({
   permission:key, allowed:pages[key]===true,
   target:card||tab, exists:!!document.getElementById(card||tab)
 }));
 console.info('[V6] Real Google Sheet PAGE_* flags:',
   statusRows.filter(x=>x.allowed).length,'of',statusRows.length,'TRUE');
 console.table(statusRows);
 const unknown=Object.keys(pages).filter(key=>!Object.hasOwn(REAL_PAGE_MAP,key));
 if(unknown.length)console.warn('[V6] Unexpected PAGE_* keys:',unknown);
 const missing=statusRows.filter(x=>!x.exists);
 if(missing.length)console.warn('[V6] Cards missing from index.html:',missing);
 for(const tab of ['facility','operations','production','inventory','productivity','training','system']){
  const visible=Object.entries(REAL_PAGE_MAP).some(([key,[group]])=>group===tab&&pages[key]===true);
  const nav=document.querySelector('.nav-item[data-tab="'+tab+'"]'),section=document.getElementById(tab);
  if(nav)nav.style.display=visible?'':'none';
  if(section){section.style.display=visible?'':'none';section.classList.remove('active');}
  if(nav)nav.classList.remove('active');
  if(visible)permitted.push(tab);
 }
 const first=permitted[0];
 if(first){document.querySelector('.nav-item[data-tab="'+first+'"]')?.classList.add('active');document.getElementById(first)?.classList.add('active');}
 else{document.querySelector('main.main')?.insertAdjacentHTML('afterbegin','<div style="padding:22px;color:#ffb758">No LMS pages enabled for this employee in the real Access Control Sheet.</div>');}
 // Load Facility WIP only when the signed-in account has its PAGE_* flag.
 // This checks local UI visibility, NOT authorization for Facility WIP's own backend.
 const f=document.getElementById('facilityWipFrame');
 if(f){
  if(pages.PAGE_FACILITY_WIP===true){
   f.src=f.dataset.src||'./FacilityWIP.html?embedded=1';
   console.info('[V6] Facility WIP iframe enabled:', f.getAttribute('src'));
  }else{
   f.removeAttribute('src');
   console.info('[V6] Facility WIP disabled by PAGE_FACILITY_WIP');
  }
 }
 const user=document.getElementById('userName');if(user)user.textContent=(`${p.firstName} ${p.lastName}`).trim()||p.email;
 const role=document.querySelector('.user-role');if(role)role.textContent=p.role;
 const avatar=document.getElementById('userAvatar');if(avatar)avatar.textContent=(p.firstName||p.email)[0].toUpperCase();
 // Event capture prevents local opening of disabled page card navigation links.
 document.querySelectorAll('.card').forEach(card=>{
  const denied=card.dataset.qaAllowed==='false';
  card.querySelectorAll('a,button').forEach(b=>{
   if(denied){b.disabled=true;b.setAttribute('aria-disabled','true');}
  });
 });
}


// Admin-only, server-authenticated access request notifications.
// Nothing here grants access or writes profile permissions.
function qaV6MountAdminAlerts(){
 if(!window.google?.script?.run || !window.__ZENNI_V6_USER__?.canManageRequests)return;
 const host=document.createElement('div');host.id='qaV6Alerts';
 host.style.cssText='position:fixed;z-index:2000;right:14px;top:76px;font:13px system-ui;color:#eaf5ff;max-width:min(380px,calc(100vw - 28px));';
 const bell=document.createElement('button');bell.type='button';bell.textContent='🔔 Access requests';
 bell.style.cssText='cursor:pointer;border:1px solid #406b83;border-radius:12px;background:#0d2138;color:#fff;padding:11px 15px;font-weight:700;';
 const panel=document.createElement('div');panel.hidden=true;panel.style.cssText='max-height:60vh;overflow:auto;background:#0b182c;border:1px solid #38607e;border-radius:12px;margin-top:7px;padding:12px;box-shadow:0 18px 45px #0009;';
 bell.addEventListener('click',()=>{panel.hidden=!panel.hidden;});
 host.append(bell,panel);document.body.append(host);
 function render(data){
  const requests=Array.isArray(data?.requests)?data.requests:[];
  const count=Number(data?.count||0);
  bell.textContent='🔔 Access requests'+(count?' ('+count+')':'');
  panel.replaceChildren();
  const title=document.createElement('strong');title.textContent=count?'Pending registration / access requests':'No pending access requests';panel.append(title);
  requests.forEach(r=>{
   const row=document.createElement('div');row.style.cssText='border-top:1px solid #29435e;margin-top:9px;padding-top:9px;';
   const who=document.createElement('div');who.style.fontWeight='700';who.textContent=r.email;
   const info=document.createElement('div');info.textContent=r.reason+' · '+r.at+' · Attempts '+r.attempts;
   info.style.cssText='color:#9cb7cf;font-size:11px;margin:5px 0;';
   const review=document.createElement('button');review.textContent='Mark reviewed';
   review.style.cssText='background:#2c5d6c;color:#fff;border:0;border-radius:6px;padding:7px 10px;cursor:pointer';
   review.onclick=()=>{review.disabled=true;google.script.run.withSuccessHandler(poll).withFailureHandler(err=>{review.disabled=false;console.error(err);}).qaV6ResolveRequest(r.id);};
   row.append(who,info,review);panel.append(row);
  });
 }
 function poll(){
  if(document.visibilityState==='hidden')return;
  google.script.run.withSuccessHandler(render).withFailureHandler(err=>console.warn('[V6] Admin notification polling failed',err)).qaV6AdminNotifications();
 }
 poll();window.setInterval(poll,30000);
}
