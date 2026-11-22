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


// Admin notifications: explicit APPROVE / DENY / REVIEW actions.
// Every action is reauthorized by Apps Script; UI buttons are not authority.
function qaV6MountAdminAlerts(){
 if(!window.google?.script?.run || !window.__ZENNI_V6_USER__?.canManageRequests)return;
 const host=document.createElement('div');host.id='qaV6Alerts';
 host.style.cssText='position:fixed;z-index:2000;right:14px;top:76px;font:13px system-ui;color:#eaf5ff;max-width:min(405px,calc(100vw - 28px));';
 const bell=document.createElement('button');bell.type='button';bell.textContent='🔔 Access requests';
 bell.style.cssText='cursor:pointer;border:1px solid #406b83;border-radius:12px;background:#0d2138;color:#fff;padding:11px 15px;font-weight:700;';
 const panel=document.createElement('div');panel.hidden=true;panel.style.cssText='max-height:70vh;overflow:auto;background:#0b182c;border:1px solid #38607e;border-radius:12px;margin-top:7px;padding:12px;box-shadow:0 18px 45px #0009;';
 const status=document.createElement('div');status.style.cssText='color:#a6dfef;font-size:12px;padding-top:7px;';status.setAttribute('role','status');
 bell.addEventListener('click',()=>{panel.hidden=!panel.hidden;});
 host.append(bell,panel);document.body.append(host);
 const pageOptions=[
 ['PAGE_FACILITY_WIP','Facility WIP'],['PAGE_INVESTIGATION_HOLD','Investigation Hold'],['PAGE_LENS_INTAKE','Lens Intake'],['PAGE_TRUE_CURVE','True Curve'],
 ['PAGE_SURFACE_DASHBOARD','Surface Dashboard'],['PAGE_AR_DASHBOARD','AR Dashboard'],['PAGE_AR_RUN_TRACKER','AR Run Tracker'],['PAGE_PRODUCTION_FORECAST','Production Forecast'],
 ['PAGE_FINISH_DASHBOARD','Finish Dashboard'],['PAGE_COATING_BREAKAGE','Coating Breakage'],['PAGE_POWER_ANALYSIS','Power Analysis'],
 ['PAGE_INCOMING_JOBS','Incoming Jobs'],['PAGE_PICKING_DASHBOARD','Picking Dashboard'],['PAGE_PRODUCTIVITY_HUB','Productivity Hub'],
 ['PAGE_BREAKAGE_QUALITY_HUB','Breakage / Quality Hub'],['PAGE_FACILITY_BREAKAGE','Facility Breakage'],['PAGE_PPC_DOCUMENT_CONTROL','PPC Document Control'],['PAGE_SYSTEM','System']
 ];
 function button(label,color,handler){
  const b=document.createElement('button');b.type='button';b.textContent=label;
  b.style.cssText='cursor:pointer;border:1px solid #3e6682;border-radius:7px;padding:7px 10px;background:'+color+';color:white;font-weight:600;';
  b.onclick=handler;return b;
 }
 function invoke(name,args,buttons){
  buttons.forEach(b=>b.disabled=true);status.textContent='Saving decision…';
  google.script.run.withSuccessHandler(result=>{status.textContent=result.status+' · '+result.email;poll();})
   .withFailureHandler(err=>{buttons.forEach(b=>b.disabled=false);status.textContent=err?.message||'Could not save request decision';console.error(err);})[name](...args);
 }
 function render(data){
  const requests=Array.isArray(data?.requests)?data.requests:[];
  const count=Number(data?.count||0);
  bell.textContent='🔔 Access requests'+(count?' ('+count+')':'');
  panel.replaceChildren();
  const title=document.createElement('strong');title.textContent=count?'Pending registration / access requests':'No pending access requests';panel.append(title);
  if(count){const hint=document.createElement('p');hint.textContent='Approval grants only the explicitly selected pages. No pages are preselected.';hint.style.cssText='font-size:11px;color:#aec8db;margin:8px 0;';panel.append(hint);}
  requests.forEach(r=>{
   const row=document.createElement('div');row.style.cssText='border-top:1px solid #29435e;margin-top:9px;padding-top:9px;';
   const who=document.createElement('div');who.style.fontWeight='700';who.textContent=r.email;
   const info=document.createElement('div');info.textContent=r.reason+' · '+r.at+' · Attempts '+r.attempts;
   info.style.cssText='color:#9cb7cf;font-size:11px;margin:5px 0 9px;';
   const actions=document.createElement('div');actions.style.cssText='display:flex;flex-wrap:wrap;gap:6px;';
   const deny=button('Deny','#663146',()=>{
    if(confirm('Deny this access request for '+r.email+'? No permissions will be granted.'))invoke('qaV64DenyRequest',[r.id],[deny,review]);
   });
   const review=button('Mark reviewed','#2c5d6c',()=>invoke('qaV64ReviewRequest',[r.id],[deny,review]));
   actions.append(deny,review);
   row.append(who,info,actions);
   if(window.__ZENNI_V6_USER__.canApproveRequests===true){
    const approve=button('Approve / select pages','#246b51',()=>{details.hidden=!details.hidden;});actions.prepend(approve);
    const details=document.createElement('div');details.hidden=true;
    details.style.cssText='background:#0d2538;border:1px solid #345570;border-radius:8px;padding:10px;margin-top:9px;';
    const namePreview=document.createElement('div');
    namePreview.style.cssText='font-size:12px;line-height:1.65;color:#c6e4f7;margin-bottom:10px;';
    namePreview.textContent='Loading verified employee profile…';details.append(namePreview);
    const roleLabel=document.createElement('label');roleLabel.textContent='Role';
    roleLabel.style.cssText='display:block;font-size:12px;font-weight:700;margin-bottom:5px;';
    const roleSelect=document.createElement('select');
    roleSelect.style.cssText='width:100%;padding:9px;color:#eef8ff;background:#142a45;border:1px solid #51728d;border-radius:7px;margin-bottom:9px;';
    roleSelect.disabled=true;roleLabel.append(roleSelect);details.append(roleLabel);
    const roleNotice=document.createElement('p');
    roleNotice.style.cssText='font-size:11px;color:#edca7d;margin:0 0 11px;line-height:1.5;';
    roleNotice.textContent='Role defaults still apply to non-page permissions. Management roles may grant user administration access.';
    details.append(roleNotice);
    const label=document.createElement('div');label.textContent='Select permitted LMS pages:';
    label.style.cssText='font-size:12px;font-weight:700;margin-bottom:8px;';details.append(label);
    const checks=[];
    pageOptions.forEach(([key,title])=>{
     const line=document.createElement('label');line.style.cssText='display:flex;align-items:center;gap:8px;margin:5px 0;cursor:pointer;';
     const input=document.createElement('input');input.type='checkbox';input.value=key;input.checked=false;
     const text=document.createElement('span');text.textContent=title;line.append(input,text);details.append(line);checks.push(input);
    });
    let profileReady=false,profile=null;
    google.script.run.withSuccessHandler(data=>{
     profile=data;profileReady=true;
     namePreview.replaceChildren();
     [['Email',data.email],['Username',data.username],['First Name',data.firstName],['Last Name',data.lastName]].forEach(([key,val])=>{
      const entry=document.createElement('div');entry.textContent=key+': '+val;namePreview.append(entry);
     });
     (data.roles||[]).forEach(role=>{const opt=document.createElement('option');opt.value=role;opt.textContent=role;roleSelect.append(opt);});
     roleSelect.value='EMPLOYEE';roleSelect.disabled=false;
    }).withFailureHandler(error=>{
     namePreview.textContent='Profile unavailable: '+(error?.message||'Unknown error');
    }).qaV65ApprovalPreview(r.id);
    const submit=button('Confirm approval','#246b51',()=>{
     if(!profileReady || roleSelect.disabled){status.textContent='Wait for the verified profile to load.';return;}
     const selected=checks.filter(i=>i.checked).map(i=>i.value);
     if(!selected.length){status.textContent='Select at least one page to approve.';return;}
     const role=roleSelect.value;
     if(!confirm('Register '+profile.email+' as '+role+'\nUsername: '+profile.username+'\nName: '+profile.firstName+' '+profile.lastName+'\nPages: '+selected.length+'\n\nNon-page role permissions will also apply. Confirm?'))return;
     invoke('qaV65ApproveRequest',[r.id,role,selected],[submit,approve,deny,review]);
    });
    details.append(submit);row.append(details);
   }
   panel.append(row);
  });
  panel.append(status);
 }
 function poll(){
  if(document.visibilityState==='hidden')return;
  google.script.run.withSuccessHandler(render).withFailureHandler(err=>console.warn('[V6] Admin notification polling failed',err)).qaV6AdminNotifications();
 }
 poll();window.setInterval(poll,30000);
}
