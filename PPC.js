/* PPC Control Hub V12 — Workspace HTML Service read-only frontend.
 * Hosted as PPC / PPC_Styles / PPC_Scripts inside Apps Script.
 * If API cannot be read, display zero records and explicit error.
 * All write actions remain disabled pending authorized API integration.
 */
const CONFIG = {
  MODE: 'api', // live read-only; writes stay disabled
  API_URL: '', // Not used in HTML Service; google.script.run talks to Code.gs
  DUE_SOON_DAYS: 7
};
const normalizeStatus=(status,stage)=>{if(['Not Started','Rough Draft in Progress','Ready for Signature','In Routing','Published'].includes(status))return status;if(status==='Complete'||stage==='Published')return 'Published';if(stage==='Approval / Routing')return 'In Routing';if(stage==='WI Draft'||status==='In Progress')return 'Rough Draft in Progress';return 'Not Started';};
const seed = [
  ['2026-SUR-133-01','4 Station Inspection Process','Surface','Brian Honicker','Published','Published','WI-SUR-133'],
  ['2026-FIN-415-02','Disposal of Broken Production Trays','Finish','Kim Manack','Published','Published','WI-FIN-415'],
  ['2026-DST-416-03','Package Shipping Requests','Distribution','Scott Anderson','In Routing','Approval / Routing','WI-DST-416'],
  ['2026-CS-417-04','CS Data Entry','Customer Service','Jen Gray','In Routing','Approval / Routing','WI-CS-417'],
  ['2026-INV-091-05','LMS Putaway','Inventory','Crystal Bayliss','Rough Draft in Progress','WI Draft','WI-INV-091'],
  ['2026-SUR-132-06','Orbit 2 Fast/Mill Tool Thickness Correction','Surface','Nash Poston','Ready for Signature','Approval / Routing','WI-SUR-132'],
  ['2026-DST-286-07','Shipping Pairing Reject Line','Distribution','Kamal Chapagai','Rough Draft in Progress','WI Draft','WI-DST-286'],
  ['2026-FIN-030-08','Weco Handstone Operation BIGS','Finish','Brittany Heckman','Rough Draft in Progress','WI Draft','WI-FIN-030'],
  ['2026-FIN-009-09','FOCO SPF - Progressive Inspection','Finish','Brian Honicker','In Routing','Approval / Routing','WI-FIN-009'],
  ['2026-SUR-074-010','Auto Engraver Start Up','Surface','Bobby Karr','Rough Draft in Progress','WI Draft','WI-SUR-074'],
  ['2026-SUR-076-011','De-Blocker Start Up','Surface','Bobby Karr','Rough Draft in Progress','WI Draft','WI-SUR-076'],
  ['2026-SUR-246-012','WI-SUR-246 update (title needs verification)','Surface','Bobby Karr','Rough Draft in Progress','WI Draft','WI-SUR-246'],
  ['2026-SUR-108-013','Alloy Reclaim','Surface','Bobby Karr','Rough Draft in Progress','WI Draft','WI-SUR-108']
].map(([number,title,department,requestor,status,stage,wi])=>({number,title,department,requestor,status,stage,wi,dueDate:'',testEnd:'',publicNotes:''}));
let items = CONFIG.MODE === 'api' ? [] : seed.map(x=>({...x}));
const $ = s=>document.querySelector(s);
const $$ = s=>Array.from(document.querySelectorAll(s));
const escapeHtml = value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const daysUntil = value=>{if(!/^\d{4}-\d{2}-\d{2}$/.test(value||''))return null;const t=new Date();t.setHours(0,0,0,0);return Math.round((new Date(`${value}T00:00:00`)-t)/86400000)};
const health = r=>{if(r.status==='Published')return 'Published';const d=daysUntil(r.dueDate);return d===null?'Unscheduled':d<0?'Overdue':d<=CONFIG.DUE_SOON_DAYS?'Due Soon':'On Track';};
const displayDate = s=>s?new Date(`${s}T12:00:00`).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}):'Not set';
const tag = (s)=>`<span class="chip ${({'Published':'complete','In Routing':'routing','Ready for Signature':'signature','Rough Draft in Progress':'progress','Not Started':'started','Overdue':'overdue','Due Soon':'soon','On Track':'track','Unscheduled':'unscheduled'})[s]||'started'}">${escapeHtml(s)}</span>`;
function sortWork(rs){const rank={Overdue:0,'Due Soon':1,'On Track':2,Unscheduled:3,Published:4};return [...rs].sort((a,b)=>rank[health(a)]-rank[health(b)]||((a.dueDate||'9999')<(b.dueDate||'9999')?-1:1));}
function drawTable(rows,type='requests'){
 if(!rows.length)return '<div class="empty">No PPC requests match the current filters.</div>';
 const isTest=type==='testing'; const isApproval=type==='approval'; const isUpdate=type==='updates';
 return `<div class="tablewrap"><table><thead><tr><th>PPC number</th><th>Document / title</th><th>Requestor</th><th>${isTest?'Test end':isApproval?'Stage':isUpdate?'Public notes':'Department'}</th><th>Status</th><th>${isTest?'PPC due':isApproval?'PPC due':isUpdate?'Last update':'Target date'}</th><th>Action</th></tr></thead><tbody>${rows.map(r=>`<tr><td><b>${escapeHtml(r.number)}</b></td><td class="titlecell">${escapeHtml(r.title)}</td><td>${escapeHtml(r.requestor)}</td><td>${isTest?displayDate(r.testEnd):isApproval?escapeHtml(r.stage):isUpdate?escapeHtml(r.publicNotes||'No update yet'):escapeHtml(r.department)}</td><td>${tag(r.status)}</td><td>${displayDate(r.dueDate)}</td><td><button class="tableaction" data-view="${escapeHtml(r.number)}">View</button></td></tr>`).join('')}</tbody></table></div>`;
}
function draw(){const active=items.filter(r=>r.status!=='Published');const overdue=active.filter(r=>health(r)==='Overdue').length;const soon=active.filter(r=>health(r)==='Due Soon').length;const draft=items.filter(r=>r.status==='Rough Draft in Progress').length;const published=items.filter(r=>r.status==='Published').length;const unscheduled=active.filter(r=>health(r)==='Unscheduled').length;
 const stats=[['Total PPC Requests',items.length,'','All records in current view'],['Overdue',overdue,'red','Past target date'],['Due within 7 days',soon,'amber','Upcoming deadlines'],['Rough Draft',draft,'','Draft in progress'],['Published',published,'green','Published documents']];
 $('#metrics').innerHTML=stats.map(([label,n,cls,foot])=>`<article class="metric ${cls}"><span class="mt">${label}</span><strong>${n}</strong><span class="foot">${foot}</span></article>`).join('');
 const counts=[['Not Started',items.filter(r=>r.status==='Not Started').length,'gray'],['Rough Draft in Progress',draft,'amber'],['Ready for Signature',items.filter(r=>r.status==='Ready for Signature').length,'blue'],['In Routing',items.filter(r=>r.status==='In Routing').length,'blue'],['Published',published,'green']];
 $('#statusBreakdown').innerHTML=counts.map(([label,n,cls])=>`<div class="breakrow"><span>${label}</span><div class="track"><div class="fill ${cls}" style="width:${items.length?n/items.length*100:0}%"></div></div><b>${n}</b></div>`).join('');
 const deadlines=[['Overdue',overdue,'red'],['Due soon (7 days)',soon,'amber'],['On track',active.filter(r=>health(r)==='On Track').length,'green'],['Missing target date',unscheduled,'gray']];
 $('#deadlineBars').innerHTML=deadlines.map(([label,n,cls])=>`<div class="deadline-row"><div class="label"><b>${label}</b><b>${n}</b></div><div class="track"><div class="fill ${cls}" style="width:${active.length?n/active.length*100:0}%"></div></div></div>`).join('');
 $('#dashTable').innerHTML=drawTable(sortWork(items).slice(0,8));
 const q=$('#search').value.toLowerCase(),status=$('#statusFilter').value,due=$('#dueFilter').value;
 const filtered=sortWork(items).filter(r=>(!status||r.status===status)&&(!due||health(r)===due)&&[r.number,r.title,r.requestor,r.wi].some(s=>String(s).toLowerCase().includes(q)));
 $('#requestsTable').innerHTML=drawTable(filtered);$('#testingTable').innerHTML=drawTable(sortWork(items).filter(r=>r.testEnd),'testing');
 $('#approvalTable').innerHTML=drawTable(items.filter(r=>['Ready for Signature','In Routing','Published'].includes(r.status)),'approval');$('#updatesTable').innerHTML=drawTable(items,'updates');
}
function showPage(page){$$('.page').forEach(x=>x.classList.toggle('active',x.id===page));$$('.nav').forEach(x=>x.classList.toggle('active',x.dataset.page===page));$('#pageTitle').textContent=({dashboard:'Dashboard Overview',requests:'PPC Requests',testing:'Testing Schedule',approvals:'Approvals & Routing',updates:'Requestor Updates'})[page]||page;}
function toast(message){const node=$('#toast');node.textContent=message;node.classList.add('show');clearTimeout(toast.t);toast.t=setTimeout(()=>node.classList.remove('show'),3800)}
function openModal(){$('#overlay').classList.remove('hidden');$('#requestForm').reset();$('#requestForm').elements.requestDate.value=new Date().toISOString().slice(0,10);updatePpcSetup();$('#ppcDepartment').focus()}
function closeModal(){$('#overlay').classList.add('hidden')}
async function loadFromApi(){
 // google.script.run is same-origin to Apps Script HTML Service and carries
 // the deployed Google Workspace identity. No fetch/CORS, no public API key.
 const payload=await new Promise((resolve,reject)=>{
   if(!window.google || !google.script || !google.script.run){
     reject(new Error('Open this dashboard through the deployed Apps Script web-app URL.'));
     return;
   }
   google.script.run.withSuccessHandler(resolve)
     .withFailureHandler(err=>reject(new Error(err && err.message || String(err))))
     .getPpcDashboardData();
 });
 if(!payload || payload.ok!==true || !Array.isArray(payload.rows))
   throw new Error(payload && payload.error || 'Invalid response from PPC backend');
 items=payload.rows.map(r=>({
   number:String(r.number||''),title:String(r.title||''),requestor:String(r.requestor||''),
   department:String(r.department||''),wi:String(r.wi||''),
   status:normalizeStatus(r.status,r.stage),stage:String(r.stage||''),
   dueDate:String(r.dueDate||''),testEnd:String(r.testEnd||''),
   publicNotes:String(r.publicNotes||'')
 })).filter(r=>r.number);
 return items.length;
}

const DEPT_LABELS={SUR:'Surface',FIN:'Finish',AR:'AR',INV:'Inventory',DST:'Distribution',CS:'Customer Service',GEN:'Lab Wide',MTC:'Maintenance'};
const ALL_PPC_STATUSES=['Not Started','Rough Draft in Progress','Ready for Signature','In Routing','Published'];
function parseWIs(v){return [...new Set(String(v||'').toUpperCase().split(/[,;\n]+/).map(x=>x.trim()).filter(Boolean))];}
function validWIs(xs){return xs.length>0 && xs.every(x=>/^WI-[A-Z]{2,8}-\d{1,6}$/.test(x));}
function primaryDoc(wi){const m=wi.match(/^WI-[A-Z]{2,8}-(\d{1,6})$/);return m?m[1].padStart(3,'0'):'';}
function maxPpcSeq(year){return items.reduce((max,r)=>{let m=String(r.number).match(/^(\d{4})-[A-Z]{2,8}-\d+-(\d+)$/);return m && +m[1]===year?Math.max(max,+m[2]):max},0);}
function maxDepartmentDocument(dept){
 return items.reduce((max,r)=>{const m=String(r.number||'').match(/^\d{4}-([A-Z]{2,8})-(\d+)-\d+$/);return m&&m[1]===dept?Math.max(max,Number(m[2])):max;},0);
}
function updatePpcSetup(){
 const action=$('#ppcAction').value,method=$('#numberMethod').value,dept=$('#ppcDepartment').value,existing=$('#existingPpc');
 const selected=existing.value;
 $('#existingPpcWrap').hidden=action!=='EXISTING';
 $('#numberMethodWrap').hidden=action==='EXISTING';
 $('#manualPpcWrap').hidden=action==='EXISTING'||method!=='MANUAL';
 // Only offer PPCs from the requesting department; never carry an old selection across departments.
 const matches=items.filter(r=>r.department===DEPT_LABELS[dept] || String(r.number).split('-')[1]===dept);
 existing.innerHTML='<option value="">Select an existing PPC</option>'+matches.map(r=>`<option value="${escapeHtml(r.number)}">${escapeHtml(r.number)} — ${escapeHtml(r.title)}</option>`).join('');
 if(matches.some(r=>r.number===selected))existing.value=selected;
 const selectedRecord=items.find(r=>r.number===existing.value);
 const wiField=$('#wiNumbersEntry');
 wiField.value=action==='EXISTING'&&selectedRecord?selectedRecord.wi||'':'';
 wiField.readOnly=true;wiField.disabled=action==='NEW';wiField.required=false;
 wiField.placeholder=action==='NEW'?'Assigned after PPC creation':'Select existing PPC to view linked WI';
 $('#wiModeHelp').textContent=action==='NEW'?'Document / WI Number will be assigned later; the PPC uses the next department document number.':'WI numbers are loaded automatically from the PPC selected above.';
 $('#wiPrintNumber').value=wiField.value;
 const year=new Date().getFullYear(); const nextDoc=String(maxDepartmentDocument(dept)+1).padStart(3,'0');
 let number='Choose a department';
 if(action==='EXISTING') number=existing.value||'Select existing PPC';
 else if(method==='MANUAL')number=$('#manualPpc').value.trim().toUpperCase()||'Enter manual PPC number';
 else if(dept)number=`${year}-${dept}-${nextDoc}-${String(maxPpcSeq(year)+1).padStart(3,'0')} (preview only)`;
 $('#ppcNumberPreview').textContent=number;
 $('#ppcPrintNumber').value=number;
}
['ppcDepartment','wiNumbersEntry','ppcAction','numberMethod','manualPpc','existingPpc'].forEach(id=>$('#'+id).addEventListener('input',updatePpcSetup));
async function submit(e){e.preventDefault();const f=Object.fromEntries(new FormData(e.currentTarget));const action=$('#ppcAction').value,method=$('#numberMethod').value,department=$('#ppcDepartment').value,wis=parseWIs($('#wiNumbersEntry').value);
 if(!department){toast('Select a requesting department.');return;}
 if(action==='EXISTING'){
  const id=$('#existingPpc').value; const found=items.find(r=>r.number===id&&r.department===DEPT_LABELS[department]);
  if(!found){toast('Select an existing PPC for this department.');return;}
  if(CONFIG.MODE==='api'){toast('Existing PPC update requires a verified API adapter. No changes sent.');return;}
  found.publicNotes='Existing PPC selected in demo; no changes saved to Google Sheets.';draw();closeModal();showPage('requests');toast('Selected '+id+' in demo.');return;
 }
 const year=new Date().getFullYear();const doc=String(maxDepartmentDocument(department)+1).padStart(3,'0');let id='';
 if(method==='MANUAL'){
  id=$('#manualPpc').value.trim().toUpperCase();const match=id.match(/^(\d{4})-([A-Z]{2,8})-(\d{3,6})-(\d{1,6})$/);
  if(!match||+match[1]!==year||match[2]!==department){toast(`Manual number must start ${year}-${department}- and contain document and sequence numbers.`);return;}
 }else{id=`${year}-${department}-${doc}-${String(maxPpcSeq(year)+1).padStart(3,'0')}`;}
 if(items.some(x=>x.number===id)){toast('PPC number already exists. Select another manual number.');return;}
 if(CONFIG.MODE==='api'){toast('Live submission requires the verified API write adapter. No changes sent.');return;}
 items.unshift({number:id,title:f.title,requestor:f.requestor,department:DEPT_LABELS[department],wi:wis.join(', '),status:'Not Started',stage:'Request Submitted',dueDate:f.dueDate,testEnd:'',publicNotes:'Demo request only — not saved to Google Sheets.',description:f.description,justification:f.justification,manager:f.manager,impactQuality:f.impactQuality,impactCost:f.impactCost,impactSchedule:f.impactSchedule,impactResources:f.impactResources,impactOther:f.impactOther});draw();closeModal();showPage('requests');toast(`Demo PPC ${id} created (not saved).`);
}

let activeDetailPpc = null;
function viewPpc(id){
 const r=items.find(x=>x.number===id);if(!r)return;
 activeDetailPpc=id;
 $('#detailPpcNo').textContent=r.number;
 $('#detailPpcTitle').textContent=r.title;
 $('#detailStatus').innerHTML=ALL_PPC_STATUSES.map(s=>`<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
 $('#detailStatus').value=r.status;
 $('#detailMeta').replaceChildren();
 for(const [key,value] of [['Department',r.department],['WI number(s)',r.wi||'Not assigned'],['Requestor',r.requestor],['Target date',displayDate(r.dueDate)]]){
   const node=document.createElement('div');const b=document.createElement('b');b.textContent=key;const span=document.createElement('span');span.textContent=value;node.append(b,span);$('#detailMeta').append(node);
 }
 const fields=r.approval||{};
 $$('#approvalDocument input, #approvalDocument textarea').forEach(el=>{
   el.type==='checkbox'?el.checked=!!fields[el.name]:el.value=fields[el.name]||'';
 });
 updateApprovalGate();
 $('#detailOverlay').classList.remove('hidden');
}
function updateApprovalGate(){
 const allowed=$('#detailStatus').value==='Ready for Signature';
 $('#approvalVisible').hidden=!allowed;
 $('#approvalGateHelp').textContent=allowed?'Complete, save and print the management approval page below.':'The management approval form is available only while the PPC is Ready for Signature.';
 $('#detailPrint').hidden=!allowed;
}
function closeDetail(){ $('#detailOverlay').classList.add('hidden');activeDetailPpc=null; }
$('#closeDetail').addEventListener('click',closeDetail);
$('#detailOverlay').addEventListener('click',e=>{if(e.target.id==='detailOverlay')closeDetail()});
$('#detailStatus').addEventListener('change',e=>{
 const r=items.find(x=>x.number===activeDetailPpc);if(!r)return;
 if(CONFIG.MODE==='api'){e.target.value=r.status;toast('Live status changes require a verified, authorized API write endpoint.');return;}
 r.status=e.target.value;r.stage=r.status==='Ready for Signature'?'Awaiting Management Signatures':r.status;
 updateApprovalGate();draw();toast('Demo status updated for this session only.');
});
$('#saveApproval').addEventListener('click',()=>{
 const r=items.find(x=>x.number===activeDetailPpc);if(!r||r.status!=='Ready for Signature')return;
 if(CONFIG.MODE==='api'){toast('Live approval saving requires backend integration.');return;}
 r.approval={};$$('#approvalDocument input, #approvalDocument textarea').forEach(el=>r.approval[el.name]=el.type==='checkbox'?el.checked:el.value);
 toast('Approval form saved in this demo session only.');
});
$('#approvalAccepted').addEventListener('change',e=>{if(e.target.checked)$('#approvalRejected').checked=false});
$('#approvalRejected').addEventListener('change',e=>{if(e.target.checked)$('#approvalAccepted').checked=false});
function printApprovalOnly(){if($('#detailStatus').value!=='Ready for Signature')return;document.body.classList.add('printing-approval');window.print();}
$('#printApproval').addEventListener('click',printApprovalOnly);
$('#detailPrint').addEventListener('click',printApprovalOnly);
window.addEventListener('afterprint',()=>document.body.classList.remove('printing-approval'));

document.addEventListener('click',e=>{const nav=e.target.closest('[data-page]');if(nav)showPage(nav.dataset.page);const goto=e.target.closest('[data-goto]');if(goto)showPage(goto.dataset.goto);if(e.target.closest('[data-new]'))openModal();const view=e.target.closest('[data-view]');if(view)viewPpc(view.dataset.view)});
const demoTests=[];
function openTest(){
  $('#testForm').reset(); $('#testDepartment').value=''; $('#testAction').value='NEW'; $('#testExisting').value=''; $('#testExistingWrap').hidden=true; $('#testDurationHint').classList.remove('invalid');
  const linked=$('#testPpc');linked.innerHTML='<option value="">Select existing PPC</option>'+items.map(r=>`<option value="${escapeHtml(r.number)}">${escapeHtml(r.number)}</option>`).join('');
  $('#testSelectedPpc').innerHTML='<option value="">Choose a department first</option>';
  $('#testPpcDisplay').value='';$('#testLinkedDetails').textContent='Select a PPC to view its title, WI and current status.'; ['testSetupStart','testSetupEnd','testSetupOwner'].forEach(id=>$('#'+id).value=''); refreshExistingTests(); refreshTestMinDate();
  $('#testOverlay').classList.remove('hidden');setTestStep(1,true);$('#testDepartment').focus();
}
function closeTest(){$('#testOverlay').classList.add('hidden');}
function drawTests(){if(!demoTests.length)return;$('#testingTable').insertAdjacentHTML('afterbegin',`<div class="panel" style="margin-bottom:12px;padding:12px"><b>Test submissions (demo session only)</b><div class="tablewrap"><table><thead><tr><th>PPC No.</th><th>Test #</th><th>Test Reason</th><th>Owner</th><th>End</th><th>Result</th></tr></thead><tbody>${demoTests.map(t=>`<tr><td>${escapeHtml(t.ppc)}</td><td>${escapeHtml(t.testNumber)}</td><td>${escapeHtml(t.reason)}</td><td>${escapeHtml(t.owner)}</td><td>${escapeHtml(t.end)}</td><td>${escapeHtml(t.result)}</td></tr>`).join('')}</tbody></table></div></div>`);}
const originalDraw=draw;draw=function(){originalDraw();drawTests();};
$('#testForm').addEventListener('submit',e=>{e.preventDefault();const t=Object.fromEntries(new FormData(e.currentTarget));const duration=calendarDaySpan(t.start,t.end);if(duration===null||duration<14){toast('Test end date must be at least 14 calendar days after the start date.');return;}if(CONFIG.MODE==='api'){toast('API test writes not enabled: verify backend authorization and contract first.');return;}if($('#testAction').value==='EXISTING'){
  const i=demoTests.findIndex(x=>x.testKey===$('#testExisting').value);
  if(i<0){toast('Choose an existing test to update.');return;}
  t.testKey=demoTests[i].testKey;demoTests[i]=t;
 }else{t.testKey=`T-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;demoTests.unshift(t);}
 const r=items.find(x=>x.number===t.ppc);if(r)r.testEnd=t.end;draw();closeTest();showPage('testing');toast('Demo PPC test saved for this session (not Google Sheets).')});
document.addEventListener('click',e=>{if(e.target.closest('[data-test-new]'))openTest()});$('#closeTest').addEventListener('click',closeTest);$('#cancelTest').addEventListener('click',closeTest);$('#testOverlay').addEventListener('click',e=>{if(e.target.id==='testOverlay')closeTest()});$('#printForm').addEventListener('click',()=>window.print());$('#printTest').addEventListener('click',printTestDocument);
$('#addBtn').addEventListener('click',openModal);$('#closeModal').addEventListener('click',closeModal);$('#cancelModal').addEventListener('click',closeModal);$('#overlay').addEventListener('click',e=>{if(e.target.id==='overlay')closeModal()});$('#requestForm').addEventListener('submit',submit);['search','statusFilter','dueFilter'].forEach(id=>$('#'+id).addEventListener('input',draw));$('#refreshBtn').addEventListener('click',async()=>{if(CONFIG.MODE==='api'){try{await loadFromApi();toast('API data loaded')}catch(e){toast(e.message)}}else toast('Demo data refreshed — API calls disabled');draw()});
$('#today').textContent=new Date().toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric',year:'numeric'});
if(CONFIG.MODE==='api'){$('#connection').textContent='Workspace live';$('#banner').textContent='API mode is selected. Loading data; writes require a verified secure API.';loadFromApi().then(()=>{$('#banner').classList.add('live');$('#banner').textContent='Connected to Google Sheets through Workspace authentication. Editing is disabled.';draw()}).catch(e=>{items=[];$('#banner').classList.remove('live');$('#banner').textContent='LIVE CONNECTION FAILED: '+e.message+' — no demo records are displayed.';draw()})}else draw();

// V5 separate action rail and local-only draft controls.
function updateChecklist(){const f=document.getElementById('requestForm');const required=[['Department',document.getElementById('ppcDepartment').value],['Initiator',f.elements.requestor.value],['Change Title',f.elements.title.value],['Description',f.elements.description.value],['Justification',f.elements.justification.value]];document.getElementById('formChecklist').innerHTML=required.map(([label,v])=>`<div><b class="${String(v).trim()?'ok':'missing'}">${String(v).trim()?'✓':'○'}</b><span>${label}</span></div>`).join('');}
document.getElementById('sidePrint').addEventListener('click',()=>window.print());
document.getElementById('saveDraft').addEventListener('click',()=>{const form=document.getElementById('requestForm');const o={};for(const e of [...form.elements, ...document.querySelectorAll('#ppcSetup input, #ppcSetup select, #ppcSetup textarea')]){if(!e.name&&!e.id)continue;if(e.type==='radio'&&!e.checked)continue;if(e.type==='checkbox')o[e.name||e.id]=e.checked;else o[e.name||e.id]=e.value;}try{localStorage.setItem('ppcV5Draft',JSON.stringify(o));toast('Draft saved in this browser only — not Google Sheets.');}catch(err){toast('Browser storage unavailable. Draft not saved.')}});
document.getElementById('clearPpc').addEventListener('click',()=>{if(!confirm('Clear the current PPC form?'))return;document.getElementById('requestForm').reset();document.getElementById('ppcDepartment').value='';document.getElementById('wiNumbersEntry').value='';updatePpcSetup();updateChecklist();});
for(const event of ['input','change'])document.getElementById('overlay').addEventListener(event,updateChecklist);
const openModalV4=openModal;openModal=function(){openModalV4();updateChecklist()};


// V6 wizard: only one page is visible at once. Form state persists between steps.
let wizardStep=1;
function wizardMissingSetup(){
 const department=$('#ppcDepartment').value, wis=parseWIs($('#wiNumbersEntry').value);
 const errors=[];
 if(!department)errors.push('Select a requesting department.');
 
 if($('#ppcAction').value==='EXISTING'&&!$('#existingPpc').value)errors.push('Choose the existing PPC.');
 if($('#ppcAction').value==='NEW'&&$('#numberMethod').value==='MANUAL'&&!$('#manualPpc').value.trim())errors.push('Enter a manual PPC number.');
 return errors;
}
function wizardReview(){
 const f=$('#requestForm').elements;
 const vals=[
 ['Department',DEPT_LABELS[$('#ppcDepartment').value]||'Not selected'],
 ['WI Number(s)',$('#ppcAction').value==='NEW'?'Pending WI assignment':parseWIs($('#wiNumbersEntry').value).join(', ')||'Not linked'],
 ['PPC Mode',$('#ppcAction').value==='NEW'?'New PPC':'Existing PPC'],
 ['PPC Number',$('#ppcNumberPreview').textContent],
 ['Initiator',f.requestor.value||'Missing'],['Manager',f.manager.value||'Not provided'],
 ['Change Title',f.title.value||'Missing'],['Target Date',f.dueDate.value||'Not provided'],
 ['Description',f.description.value||'Missing'],['Justification',f.justification.value||'Missing'] ];
 const wrap=$('#wizardReviewContents'); wrap.replaceChildren();
 for(const [key,value] of vals){const d=document.createElement('dt');d.textContent=key;const e=document.createElement('dd');e.textContent=value;wrap.append(d,e);}
 const errors=wizardMissingSetup();
 for(const [name,label] of [['requestor','Initiator'],['title','Change Title'],['description','Description'],['justification','Justification']])if(!String(f[name].value).trim())errors.push(label+' is missing.');
 $('#wizardMissing').textContent=errors.length?'Please complete: '+errors.join(' '):'All core fields are entered. Verify accuracy and management approval requirements.';
 $('#wizardSubmit').disabled=errors.length>0;
}
function setWizardStep(next,{force=false}={}){
 if(next<1||next>3)return;
 if(!force&&next>wizardStep&&wizardStep===1){const missing=wizardMissingSetup();if(missing.length){toast(missing[0]);return;}}
 wizardStep=next;
 const modal=$('#overlay .document-modal');modal.dataset.wizardStep=String(wizardStep);
 $$('#overlay .step-tab').forEach(b=>{const n=+b.dataset.step;b.classList.toggle('is-active',n===wizardStep);b.classList.toggle('is-done',n<wizardStep);b.setAttribute('aria-current',n===wizardStep?'step':'false');});
 $('#wizardBack').hidden=wizardStep===1;
 $('#wizardNext').hidden=wizardStep===3;
 $('#wizardNext').textContent=wizardStep===1?'Next: Fill Request Form →':'Next: Review & Submit →';
 $('#wizardStepLabel').textContent=`Step ${wizardStep} of 3`;
 if(wizardStep===3)wizardReview();
 modal.scrollTop=0;
}
$('#wizardNext').addEventListener('click',()=>setWizardStep(wizardStep+1));
$('#wizardBack').addEventListener('click',()=>setWizardStep(wizardStep-1,{force:true}));
$$('#overlay .step-tab').forEach(b=>b.addEventListener('click',()=>{const next=+b.dataset.step;if(next<=wizardStep)setWizardStep(next,{force:true});else if(next===wizardStep+1)setWizardStep(next);}));
$('#wizardSubmit').addEventListener('click',()=>{wizardReview();if(!$('#wizardSubmit').disabled){const form=$('#requestForm');form.requestSubmit();}});
const openModalV5=openModal;
openModal=function(){openModalV5();setWizardStep(1,{force:true});};

// V9: independent three-step test wizard (SOP-06-02). Mirrors proposal wizard.
let testWizardStep=1;
function filterTestPpcs(){
 const department=$('#testDepartment').value;
 const options=items.filter(r=>r.department===department || (r.deptCode||'')===department || (r.number||'').split('-')[1]===department);
 $('#testSelectedPpc').innerHTML='<option value="">Select existing PPC</option>'+options.map(r=>`<option value="${escapeHtml(r.number)}">${escapeHtml(r.number)} — ${escapeHtml(r.title)}</option>`).join('');
 $('#testPpc').value=''; $('#testPpcDisplay').value='';$('#testLinkedDetails').textContent=department?(options.length?'Choose the PPC to test.':'No PPC records in this department.'):'Select a department first.'; refreshExistingTests();
}
function linkTestPpc(){
 const number=$('#testSelectedPpc').value;const record=items.find(r=>r.number===number);
 $('#testPpc').value=number;$('#testPpcDisplay').value=number;
 $('#testLinkedDetails').textContent=record?`${record.title} • ${record.status} • WI: ${record.wi||record.wiNumbers||'Not recorded'}`:'Select a PPC to view its details.'; refreshExistingTests();
}
for(const [id,name] of [['testSetupStart','start'],['testSetupEnd','end'],['testSetupOwner','owner']]){
 $('#'+id).addEventListener('input',e=>{$('#testForm').elements[name].value=e.target.value});
 $('#testForm').elements[name].addEventListener('input',e=>{$('#'+id).value=e.target.value});
}
// Compare calendar dates at UTC midnight so daylight saving transitions do not change the minimum.
function calendarDaySpan(start,end){if(!/^\d{4}-\d{2}-\d{2}$/.test(start||'')||!/^\d{4}-\d{2}-\d{2}$/.test(end||''))return null;const a=Date.parse(start+'T00:00:00Z'),b=Date.parse(end+'T00:00:00Z');return Number.isFinite(a)&&Number.isFinite(b)?Math.round((b-a)/86400000):null;}
function plusCalendarDays(date,days){if(!date)return '';const d=new Date(date+'T00:00:00Z');if(Number.isNaN(d.getTime()))return '';d.setUTCDate(d.getUTCDate()+days);return d.toISOString().slice(0,10);}
function refreshTestMinDate(){
 const min=plusCalendarDays($('#testSetupStart').value,14);$('#testSetupEnd').min=min;
 $('#testForm').elements.end.min=min;
 const span=calendarDaySpan($('#testSetupStart').value,$('#testSetupEnd').value);
 const valid=span===null||span>=14;
 $('#testDurationHint').classList.toggle('invalid',!valid);
 $('#testDurationHint').textContent=span===null?'Minimum testing period: 14 calendar days. Select a start date to see the earliest end date.':`Test duration: ${span} day(s). Earliest permitted end date: ${min}${valid?'':' — please extend the test'}.`;
}
function refreshExistingTests(){
 const mode=$('#testAction').value;$('#testExistingWrap').hidden=mode!=='EXISTING';
 const ppc=$('#testSelectedPpc').value, prior=$('#testExisting').value;
 const tests=demoTests.filter(t=>t.ppc===ppc);
 $('#testExisting').innerHTML='<option value="">Select an existing test</option>'+tests.map(t=>`<option value="${escapeHtml(t.testKey||'')}">${escapeHtml('Test '+(t.testNumber||'1')+' — '+(t.reason||'Untitled')+' ('+(t.start||'?')+' to '+(t.end||'?')+')')}</option>`).join('');
 if(tests.some(t=>t.testKey===prior))$('#testExisting').value=prior;
 if(mode==='NEW'){$('#testExisting').value='';}
}
function loadExistingTest(){
 const t=demoTests.find(x=>x.testKey===$('#testExisting').value);
 if(!t)return;
 const form=$('#testForm');for(const [key,value] of Object.entries(t)){const control=form.elements[key];if(!control)continue;if(control instanceof RadioNodeList){control.value=value;}else if(control.type!=='radio')control.value=value;}
 $('#testSetupStart').value=t.start||'';$('#testSetupEnd').value=t.end||'';$('#testSetupOwner').value=t.owner||'';
 refreshTestMinDate();
}
$('#testAction').addEventListener('change',()=>{refreshExistingTests();if($('#testAction').value==='NEW'){$('#testForm').reset();$('#testPpc').value=$('#testSelectedPpc').value;$('#testPpcDisplay').value=$('#testSelectedPpc').value;['testSetupStart','testSetupEnd','testSetupOwner'].forEach(x=>$('#'+x).value='');refreshTestMinDate();}});
$('#testExisting').addEventListener('change',loadExistingTest);
$('#testSetupStart').addEventListener('change',refreshTestMinDate);
$('#testSetupEnd').addEventListener('change',refreshTestMinDate);
$('#testDepartment').addEventListener('change',filterTestPpcs);
$('#testSelectedPpc').addEventListener('change',linkTestPpc);
function getTestErrors(includeForm=false){
 const errors=[];
 if(!$('#testDepartment').value)errors.push('Choose a department');
 if(!$('#testSelectedPpc').value)errors.push('Choose an existing PPC');
 if($('#testAction').value==='EXISTING'&&!$('#testExisting').value)errors.push('Choose an existing PPC Test');
 if(!$('#testSetupStart').value||!$('#testSetupEnd').value)errors.push('Select test start and end dates');
 if($('#testSetupStart').value&&$('#testSetupEnd').value&&calendarDaySpan($('#testSetupStart').value,$('#testSetupEnd').value)<14)errors.push('The test must span at least 14 calendar days');
 if(!$('#testSetupOwner').value.trim())errors.push('Enter main test POC');
 if(includeForm){
   const f=$('#testForm');
   for(const [name,label] of [['reason','Test reason'],['owner','Main POC'],['start','Start date'],['end','End date'],['defineTest','Test definition'],['expected','Expected measurable results'],['ownerEmail','Test owner email']]){
     if(!f.elements[name]?.value?.trim())errors.push(label);
   }
   if(f.elements.end.value&&f.elements.start.value&&calendarDaySpan(f.elements.start.value,f.elements.end.value)<14)errors.push('Test duration must be at least 14 calendar days');
 }
 return errors;
}
function testReviewSummary(){
 const form=$('#testForm'); const fields=[
 ['Test Action',$('#testAction').value==='NEW'?'Create NEW PPC Test':'Update EXISTING PPC Test'],['PPC Number',$('#testSelectedPpc').value],['Department',$('#testDepartment').selectedOptions[0]?.textContent],
 ['Test Reason',form.elements.reason.value],['Test Owner',form.elements.owner.value],['Owner Email',form.elements.ownerEmail.value],
 ['Test Dates',`${form.elements.start.value||'—'} to ${form.elements.end.value||'—'}`],
 ['Test Number',form.querySelector('[name="testNumber"]:checked')?.value],['Define Test',form.elements.defineTest.value],
 ['Expected Results',form.elements.expected.value],['Test Result',form.elements.result.value]
 ];
 $('#testReviewList').replaceChildren();fields.forEach(([k,v])=>{let dt=document.createElement('dt'),dd=document.createElement('dd');dt.textContent=k;dd.textContent=v||'Not entered';$('#testReviewList').append(dt,dd)});
 const missing=getTestErrors(true);$('#testReviewWarning').textContent=missing.length?'Complete before submitting: '+missing.join(', '):'Ready to submit. Review values for accuracy.';
 $('#testWizardSubmit').disabled=Boolean(missing.length);
}
function setTestStep(next,force=false){
 if(next<1||next>3)return;
 if(!force && next>testWizardStep){const errors=getTestErrors(testWizardStep===2);if(errors.length){toast(errors[0]);return;}}
 testWizardStep=next;
 const modal=$('#testOverlay .test-wizard-modal');modal.dataset.testStep=String(next);
 $$('#testOverlay .test-step-tab').forEach(tab=>{const n=Number(tab.dataset.testGo);tab.classList.toggle('is-active',n===next);tab.classList.toggle('is-done',n<next);tab.setAttribute('aria-selected',String(n===next))});
 $('#testWizardBack').hidden=next===1;$('#testWizardNext').hidden=next===3;
 $('#testWizardNext').textContent=next===1?'Next: Fill Test Form →':'Next: Review & Submit →';
 $('#testStepLabel').textContent=`Step ${next} of 3`;
 if(next===3)testReviewSummary();modal.scrollTop=0;
}
$('#testWizardNext').addEventListener('click',()=>setTestStep(testWizardStep+1));
$('#testWizardBack').addEventListener('click',()=>setTestStep(testWizardStep-1,true));
$$('#testOverlay .test-step-tab').forEach(tab=>tab.addEventListener('click',()=>{const n=+tab.dataset.testGo;if(n<=testWizardStep)setTestStep(n,true);else if(n===testWizardStep+1)setTestStep(n)}));
$('#testWizardSubmit').addEventListener('click',()=>{testReviewSummary();if(!$('#testWizardSubmit').disabled)$('#testForm').requestSubmit();});
function printTestDocument(){document.body.classList.add('printing-test');window.print();}
$('#testReviewPrint').addEventListener('click',printTestDocument);
window.addEventListener('afterprint',()=>document.body.classList.remove('printing-test'));
