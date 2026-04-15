/* ============================================================
   Surface Flow — Power Breakage v5
   power.js — Combo 4: Peach · Periwinkle · Teal

   COLOR MAP (mirrors CSS :root):
   CLR.peach  #fdba74  S-Power · Source→AR41 · event tag · KPI1
   CLR.peri   #818cf8  S-Axis  · AR41→Brk   · Rx(−)     · KPI2
   CLR.teal   #2dd4bf  machines· Brk→Proc   · Rx(+)     · KPI3
   CLR.series [...]    material/chart multi-series
   ============================================================ */

const API = 'https://script.google.com/macros/s/AKfycbzpPE9mQznnF6oaIYW897SgqRVpsiMqZsTYj9HyMw0JADU31-6jN7slD_wJQs6f2njn/exec';

const CLR = {
  peach:  '#d28c64',  /* Rose Gold  — S-Power · Source seg · event · KPI1 */
  peri:   '#e8a040',  /* Amber      — S-Axis  · AR41→Brk  · Rx(−) · KPI2 */
  teal:   '#d4c0a8',  /* Warm Cream — machines · Brk→Proc · Rx(+) · KPI3 */
  series: ['#d28c64','#e8a040','#d4c0a8','#c8967a','#f0b870','#e8d0b8','#b87850','#d4a060'],
  gc: 'rgba(255,255,255,0.05)',
  tc: 'rgba(255,255,255,0.65)',
};

const charts = {};
let _brk=[], _reason=[], _research=[], _anom=[];

/* ── helpers ── */
function fetchTab(tab) {
  return fetch(`${API}?tab=${tab}`)
    .then(r=>r.json())
    .then(d=>{ if(d.status!=='ok') throw new Error(d.message); return d.data; });
}

function fmtMin(m) {
  m = parseFloat(m);
  if (isNaN(m)||m<=0) return '--';
  if (m<60) return `${Math.round(m)}m`;
  const d=Math.floor(m/1440), h=Math.floor((m%1440)/60), mn=Math.round(m%60);
  return [d>0?`${d}d`:'',h>0?`${h}h`:'',mn>0?`${mn}m`:''].filter(Boolean).join(' ');
}

function fmtDate(s) {
  if (!s) return '--';
  const d=new Date(s); if(isNaN(d)) return s;
  return `${d.getMonth()+1}/${d.getDate()} ${d.getHours()}:${String(d.getMinutes()).padStart(2,'0')}`;
}

function diffMin(a,b) {
  const da=new Date(a),db=new Date(b);
  if(isNaN(da)||isNaN(db)) return 0;
  return Math.max(0,(db-da)/60000);
}

function avgArr(arr) {
  const n=arr.map(v=>parseFloat(v)).filter(v=>!isNaN(v)&&v!==null);
  return n.length ? n.reduce((a,b)=>a+b,0)/n.length : 0;
}

function reasonColor(r) {
  if (!r) return CLR.teal;
  const s=r.toLowerCase();
  if (s.includes('power')) return CLR.peach;  /* Rose Gold */
  if (s.includes('axis'))  return CLR.peri;   /* Amber */
  return CLR.teal;                            /* Warm Cream */
}

function reasonPillClass(r) {
  if (!r) return 'pill-other';
  const s=r.toLowerCase();
  if (s.includes('power')) return 'pill-peach';
  if (s.includes('axis'))  return 'pill-peri';
  return 'pill-other';
}

function mkChart(id,type,labels,datasets,extraOpts={}) {
  if (charts[id]) charts[id].destroy();
  function merge(t,s){const o=Object.assign({},t);for(const k in s)o[k]=(s[k]&&typeof s[k]==='object'&&!Array.isArray(s[k]))?merge(t[k]||{},s[k]):s[k];return o;}
  const base={
    responsive:true,maintainAspectRatio:false,
    plugins:{legend:{display:false}},
    scales:{
      x:{grid:{color:CLR.gc},ticks:{color:CLR.tc,font:{size:11,family:"'IBM Plex Mono'"}}},
      y:{grid:{color:CLR.gc},ticks:{color:CLR.tc,font:{size:11,family:"'IBM Plex Mono'"}}}
    }
  };
  charts[id]=new Chart(document.getElementById(id),{type,data:{labels,datasets},options:merge(base,extraOpts)});
}

/* ── navigation ── */
function sw(id,btn) {
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('t-'+id).classList.add('active');
  btn.classList.add('active');
}

function swInner(id,btn) {
  document.querySelectorAll('.inner-section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.inner-tab').forEach(b=>b.classList.remove('active'));
  document.getElementById('is-'+id).classList.add('active');
  btn.classList.add('active');
  setTimeout(()=>{ if(charts[id+'C']) charts[id+'C'].resize(); },50);
}

/* ── bar list builder ── */
function buildBarList(elId, entries) {
  const el=document.getElementById(elId); if(!el) return;
  const max=entries[0]?.[1]||1;
  el.innerHTML=entries.map(([label,val,display,color])=>`
    <div class="bar-row">
      <div class="bar-label">${label}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.round(val/max*100)}%;background:${color||CLR.teal}"></div></div>
      <div class="bar-val">${display||val}</div>
    </div>`).join('');
}

/* ============================================================
   OVERVIEW
   ============================================================ */
function buildOverview() {
  const totalJobs = _brk.length;
  const totalLens = _brk.reduce((s,r)=>s+(parseInt(r.LensesBroken)||0),0);
  const ar41Avg   = avgArr(_brk.map(r=>r.AR41_to_Breakage_Min));
  const brkAvg    = avgArr(_brk.map(r=>r.Breakage_to_Processed_Min));

  document.getElementById('m-jobs').textContent    = totalJobs;
  document.getElementById('m-jobs-sub').textContent = `${totalLens} lenses total`;
  document.getElementById('m-lens').textContent    = totalLens;
  document.getElementById('m-lens-sub').textContent = `across ${totalJobs} jobs`;
  document.getElementById('m-ar41brk').textContent = fmtMin(ar41Avg);
  document.getElementById('m-brkpro').textContent  = fmtMin(brkAvg);

  // Banner
  const worst = [..._brk].sort((a,b)=>(parseFloat(b.AR41_to_Breakage_Min)||0)-(parseFloat(a.AR41_to_Breakage_Min)||0))[0];
  if (worst) document.getElementById('bannerSub').textContent =
    `${worst.BrkSourceMachine||worst.AR41Machine} — avg AR41→Brk ${fmtMin(ar41Avg)} · ${totalJobs} jobs`;
  const now=new Date();
  document.getElementById('bannerTime').textContent =
    `${now.getMonth()+1}/${now.getDate()} ${now.getHours()}:${String(now.getMinutes()).padStart(2,'0')}`;

  // Reason timing (concept 3 stacked)
  const rValid = _reason.filter(r=>r.JobCount>0);
  const maxAR41= Math.max(...rValid.map(r=>parseFloat(r.AvgAR41_to_Breakage_Min)||0),1);
  const maxBrk = Math.max(...rValid.map(r=>parseFloat(r.AvgBreakage_to_Processed_Min)||0),1);
  const timingEl = document.getElementById('reasonTimingList');
  if (timingEl) {
    timingEl.innerHTML = `<div class="reason-timing">${rValid.map(r=>{
      const ar41v = parseFloat(r.AvgAR41_to_Breakage_Min)||0;
      const brkv  = parseFloat(r.AvgBreakage_to_Processed_Min)||0;
      const color = reasonColor(r.BrkReason);
      return `
        <div class="rt-block">
          <div class="rt-name" style="color:${color}">${r.BrkReason}
            <span class="rt-jobs">${r.JobCount} jobs · ${r.TotalLensesBroken||'--'} lenses</span>
          </div>
          <div class="rt-row">
            <div class="rt-row-label">AR41 → Brk</div>
            <div class="rt-track"><div class="rt-fill" style="width:${Math.round(ar41v/maxAR41*100)}%;background:${color}"></div></div>
            <div class="rt-val">${fmtMin(ar41v)}</div>
          </div>
          <div class="rt-row">
            <div class="rt-row-label">Brk → Proc</div>
            <div class="rt-track"><div class="rt-fill" style="width:${Math.round(brkv/maxBrk*100)}%;background:${color};opacity:0.6"></div></div>
            <div class="rt-val">${fmtMin(brkv)}</div>
          </div>
          ${rValid.indexOf(r)<rValid.length-1?'<div class="rt-divider"></div>':''}
        </div>`;
    }).join('')}</div>`;
  }

  // BrkSource machine bars — show all, full names, teal
  const srcAgg={}, srcRx={};
  _brk.forEach(r=>{
    const s=r.BrkSourceMachine||'Unknown';
    srcAgg[s]=(srcAgg[s]||0)+1;
    if(!srcRx[s]) srcRx[s]=[];
    srcRx[s].push(r);
  });
  const srcE=Object.entries(srcAgg).sort((a,b)=>b[1]-a[1]);
  const srcEl=document.getElementById('srcBarList');
  if (srcEl) {
    const max=srcE[0]?.[1]||1;
    srcEl.innerHTML=srcE.map(([label,val])=>{
      const jobs=srcRx[label]||[];
      const rc={}, tl=jobs.reduce((s,j)=>s+(parseInt(j.LensesBroken)||0),0);
      jobs.forEach(j=>{const r=j.BrkReason||'?'; rc[r]=(rc[r]||0)+1;});
      const tipLines=[...Object.entries(rc).map(([r,c])=>`${r}: ${c}`),`Lenses: ${tl}`].join(' · ');
      return `<div class="bar-row" title="${tipLines}">
        <div class="bar-label">${label}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${Math.round(val/max*100)}%;background:${CLR.teal}"></div></div>
        <div class="bar-val">${val}</div>
      </div>`;
    }).join('');
  }

  // Material bars — full names, alternating series colors
  const matAgg={};
  _brk.forEach(r=>{const m=r.Material||'Unknown'; matAgg[m]=(matAgg[m]||0)+(parseInt(r.LensesBroken)||0);});
  const matE=Object.entries(matAgg).sort((a,b)=>b[1]-a[1]);
  buildBarList('matList', matE.map(([label,val],i)=>[label,val,val,CLR.series[i%CLR.series.length]]));

  // Reason donut
  const rsnAgg={};
  _brk.forEach(r=>{const k=r.BrkReason||'Unknown'; rsnAgg[k]=(rsnAgg[k]||0)+1;});
  const rsnE=Object.entries(rsnAgg).sort((a,b)=>b[1]-a[1]);
  document.getElementById('reasonLeg').innerHTML=rsnE.map(([r,c])=>
    `<span class="leg-item"><span class="leg-sq" style="background:${reasonColor(r)}"></span>${r} — ${c}</span>`
  ).join('');
  mkChart('reasonPieC','doughnut',rsnE.map(e=>e[0]),[{
    data:rsnE.map(e=>e[1]),
    backgroundColor:rsnE.map(e=>reasonColor(e[0])),
    borderWidth:2,borderColor:'#0d1015'
  }],{plugins:{legend:{display:false}},scales:{x:{display:false},y:{display:false}}});
}

/* ============================================================
   FLOW FILTERS
   ============================================================ */
function populateFlowFilters() {
  [['filterReason',_brk.map(r=>r.BrkReason)],
   ['filterSource',_brk.map(r=>r.BrkSourceMachine)],
   ['filterAR41',  _brk.map(r=>r.AR41Machine)]
  ].forEach(([id,vals])=>{
    const el=document.getElementById(id);
    while(el.options.length>1) el.remove(1);
    [...new Set(vals.filter(Boolean))].sort().forEach(v=>{
      const o=document.createElement('option'); o.value=v; o.textContent=v; el.appendChild(o);
    });
  });
}

/* ── tooltip ── */
const tip=document.getElementById('ganttTip');
function showTip(e,html){tip.innerHTML=html;tip.style.display='block';moveTip(e);}
function moveTip(e){tip.style.left=(e.clientX+14)+'px';tip.style.top=(e.clientY+14)+'px';}
function hideTip(){tip.style.display='none';}
document.addEventListener('mousemove',e=>{if(tip.style.display==='block')moveTip(e);});

/* ============================================================
   RENDER FLOW
   ============================================================ */
function renderFlow() {
  const fR=document.getElementById('filterReason').value;
  const fS=document.getElementById('filterSource').value;
  const fA=document.getElementById('filterAR41').value;
  let data=[..._brk];
  if(fR) data=data.filter(r=>r.BrkReason===fR);
  if(fS) data=data.filter(r=>r.BrkSourceMachine===fS);
  if(fA) data=data.filter(r=>r.AR41Machine===fA);
  document.getElementById('flowCount').textContent=data.length;
  if(!data.length){document.getElementById('flowList').innerHTML='<div class="empty">No jobs match filters</div>';return;}

  const resMap={};
  _research.forEach(r=>{resMap[String(r.RxNumber)]=r;});

  const rows=data.map(r=>{
    const res=resMap[String(r.RxNumber)];
    const src=(r.BrkSourceMachine||'').toUpperCase();
    let srcTime=null;
    if(res) srcTime=src.includes('ORB')?res.ORBScanTime:res.OTBScanTime;
    const d1=srcTime?diffMin(srcTime,r.AR41ScanTime):0;
    const d2=parseFloat(r.AR41_to_Breakage_Min)||0;
    const d3=parseFloat(r.Breakage_to_Processed_Min)||0;
    const total=d1+d2+d3||1;
    const p1=Math.max(d1>0?1:0,Math.round(d1/total*100));
    const p2=Math.max(1,Math.round(d2/total*100));
    const p3=Math.max(1,100-p1-p2);
    const t1=`<b>Source → AR41</b><br>${r.BrkSourceMachine||'?'}: ${fmtDate(srcTime)}<br>AR41: ${fmtDate(r.AR41ScanTime)}<br>Duration: ${fmtMin(d1)}`;
    const t2=`<b>AR41 → Breakage</b><br>AR41: ${fmtDate(r.AR41ScanTime)}<br>Brk Table: ${fmtDate(r.BrkTableScanTime)}<br>Duration: ${fmtMin(d2)}`;
    const t3=`<b>Breakage → Processed</b><br>Brk Table: ${fmtDate(r.BrkTableScanTime)}<br>Processed: ${fmtDate(r.BreakageProcessedTime)}<br>Duration: ${fmtMin(d3)}`;
    return `
      <div class="gantt-row">
        <div class="gantt-rx" title="${r.RxNumber}">${r.RxNumber}</div>
        <div class="gantt-bar">
          ${d1>0?`<div class="gantt-seg seg-1" style="width:${p1}%;min-width:2px" data-tip="${t1}" onmouseenter="showTip(event,this.dataset.tip)" onmouseleave="hideTip()">${p1>10?fmtMin(d1):''}</div>`:''}
          <div class="gantt-seg seg-2" style="width:${p2}%;min-width:4px" data-tip="${t2}" onmouseenter="showTip(event,this.dataset.tip)" onmouseleave="hideTip()">${p2>10?fmtMin(d2):''}</div>
          <div class="gantt-seg seg-3" style="width:${p3}%;min-width:4px" data-tip="${t3}" onmouseenter="showTip(event,this.dataset.tip)" onmouseleave="hideTip()">${p3>10?fmtMin(d3):''}</div>
        </div>
        <span class="reason-pill ${reasonPillClass(r.BrkReason)}">${r.BrkReason||'--'}</span>
        <div class="gantt-total">${fmtMin(total)}</div>
      </div>`;
  }).join('');
  document.getElementById('flowList').innerHTML=rows;
}

/* ============================================================
   RESEARCH FILTERS
   ============================================================ */
function populateResearchFilters() {
  [['resORB',_research.map(r=>r.ORBMachine)],
   ['resReason',_research.map(r=>r.BrkReason)],
   ['resMaterial',_research.map(r=>r.Material)]
  ].forEach(([id,vals])=>{
    const el=document.getElementById(id);
    while(el.options.length>1) el.remove(1);
    [...new Set(vals.filter(Boolean))].sort().forEach(v=>{
      const o=document.createElement('option'); o.value=v; o.textContent=v; el.appendChild(o);
    });
  });
}

/* ── sph/cyl bucket helpers ── */
function sphBuckets(vals) {
  const b={'≤-6':0,'-6/-4':0,'-4/-2':0,'-2/0':0,'0/+2':0,'+2/+4':0,'≥+4':0};
  vals.forEach(v=>{
    if(isNaN(v)) return;
    if(v<=-6)b['≤-6']++;else if(v<=-4)b['-6/-4']++;else if(v<=-2)b['-4/-2']++;
    else if(v<0)b['-2/0']++;else if(v<2)b['0/+2']++;else if(v<4)b['+2/+4']++;else b['≥+4']++;
  });
  return b;
}
function cylBuckets(vals) {
  const b={'≤-3':0,'-3/-2':0,'-2/-1':0,'-1/0':0,'0':0,'0/+1':0,'≥+1':0};
  vals.forEach(v=>{
    if(isNaN(v)) return;
    if(v<=-3)b['≤-3']++;else if(v<=-2)b['-3/-2']++;else if(v<=-1)b['-2/-1']++;
    else if(v<0)b['-1/0']++;else if(v===0)b['0']++;else if(v<1)b['0/+1']++;else b['≥+1']++;
  });
  return b;
}
function sphColors(keys){
  return keys.map(k=>{
    if(k.startsWith('+')||k==='0/+2'||k==='≥+4') return '#d4c030';  /* golden yellow = positive */
    return '#4d94d4';  /* sky blue = negative */
  });
}
function cylColors(keys){
  const neg=new Set(['≤-3','-3/-2','-2/-1','-1/0']);
  const pos=new Set(['0/+1','≥+1']);
  return keys.map(k=>{
    if(k==='0')    return 'rgba(255,255,255,0.2)';
    if(pos.has(k)) return '#4d94d4';
    if(neg.has(k)) return '#d4c030';
    return 'rgba(255,255,255,0.2)';
  });
}

/* ============================================================
   RENDER RESEARCH
   ============================================================ */
function renderResearch() {
  const fO=document.getElementById('resORB').value;
  const fR=document.getElementById('resReason').value;
  const fM=document.getElementById('resMaterial').value;
  let data=[..._research];
  if(fO) data=data.filter(r=>r.ORBMachine===fO);
  if(fR) data=data.filter(r=>r.BrkReason===fR);
  if(fM) data=data.filter(r=>r.Material===fM);

  const totalLens=data.reduce((s,r)=>s+(parseInt(r.LensesBroken)||0),0);
  const avgRSph=avgArr(data.map(r=>parseFloat(r.R_Sph)));
  const avgLSph=avgArr(data.map(r=>parseFloat(r.L_Sph)));
  document.getElementById('res-jobs').textContent=data.length;
  document.getElementById('res-lens').textContent=totalLens;
  document.getElementById('res-rsph').textContent=avgRSph?avgRSph.toFixed(2):'--';
  document.getElementById('res-lsph').textContent=avgLSph?avgLSph.toFixed(2):'--';

  // R Sph chart
  const rsph=sphBuckets(data.map(r=>parseFloat(r.R_Sph)));
  mkChart('rsphC','bar',Object.keys(rsph),[{data:Object.values(rsph),backgroundColor:sphColors(Object.keys(rsph)),borderWidth:0,borderRadius:4}],{});

  // L Sph chart
  const lsph=sphBuckets(data.map(r=>parseFloat(r.L_Sph)));
  mkChart('lsphC','bar',Object.keys(lsph),[{data:Object.values(lsph),backgroundColor:sphColors(Object.keys(lsph)),borderWidth:0,borderRadius:4}],{});

  // R Cyl chart
  const rcyl=cylBuckets(data.map(r=>parseFloat(r.R_Cyl)));
  mkChart('rcylC','bar',Object.keys(rcyl),[{data:Object.values(rcyl),backgroundColor:cylColors(Object.keys(rcyl)),borderWidth:0,borderRadius:4}],{});

  // L Cyl chart
  const lcyl=cylBuckets(data.map(r=>parseFloat(r.L_Cyl)));
  mkChart('lcylC','bar',Object.keys(lcyl),[{data:Object.values(lcyl),backgroundColor:cylColors(Object.keys(lcyl)),borderWidth:0,borderRadius:4}],{});

  // Prescription table
  function pp(v){
    const n=parseFloat(v);
    if(isNaN(n)) return `<span class="px-zero">—</span>`;
    if(n===0)    return `<span class="px-zero">0</span>`;
    return n>0?`<span class="px-pos">+${n}</span>`:`<span class="px-neg">${n}</span>`;
  }
  document.getElementById('resTable').innerHTML=`
    <table class="res-table">
      <thead><tr>
        <th>RX</th><th>ORB</th><th>BrkSrc</th><th>Reason</th>
        <th>Material</th><th>Option</th><th>Curve</th>
        <th>R Sph</th><th>L Sph</th><th>R Cyl</th><th>L Cyl</th>
        <th>Add</th><th>Broken</th>
      </tr></thead>
      <tbody>${data.map(r=>`
        <tr>
          <td>${r.RxNumber}</td><td>${r.ORBMachine||'--'}</td><td>${r.BrkSourceMachine||'--'}</td>
          <td><span class="reason-pill ${reasonPillClass(r.BrkReason)}">${r.BrkReason||'--'}</span></td>
          <td>${r.Material||'--'}</td><td>${r.LensOption||'--'}</td><td>${r.BaseCurve||'--'}</td>
          <td>${pp(r.R_Sph)}</td><td>${pp(r.L_Sph)}</td>
          <td>${pp(r.R_Cyl)}</td><td>${pp(r.L_Cyl)}</td>
          <td>${r.R_AddPower||'—'}</td><td>${r.LensesBroken||'1'}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

/* ============================================================
   ALERTS
   ============================================================ */
function buildAlerts() {
  const total  = _anom.length;
  const sPower = _anom.filter(r=>(r.BrkReason||'').toLowerCase().includes('power')).length;
  const sAxis  = _anom.filter(r=>(r.BrkReason||'').toLowerCase().includes('axis')).length;
  document.getElementById('al-total').textContent  = total;
  document.getElementById('al-spower').textContent = sPower;
  document.getElementById('al-saxis').textContent  = sAxis;

  const sorted=[..._anom].map(r=>({...r,_v:parseFloat(r.AR41_to_Breakage_Min)||0})).sort((a,b)=>b._v-a._v);
  const maxV=sorted[0]?._v||1;

  document.getElementById('alertList').innerHTML=sorted.map(r=>{
    const color=reasonColor(r.BrkReason);
    const pct=Math.max(2,Math.round(r._v/maxV*100));
    return `
      <div class="alert-bar-row">
        <div>
          <div class="alert-bar-rx">${r.RxNumber}</div>
          <div class="alert-bar-desc">${r.AR41Machine||'--'} · ${r.BrkSourceMachine||'--'} · ${r.BrkReason||'--'}</div>
        </div>
        <div class="alert-bar-track"><div class="alert-bar-fill" style="width:${pct}%;background:${color}"></div></div>
        <div class="alert-bar-val" style="color:${color}">${fmtMin(r._v)}</div>
      </div>`;
  }).join('')||'<div class="empty">No alerts</div>';

  // Quick stats sidebar
  const worst=sorted[0];
  const avgAR41=avgArr(_anom.map(r=>r.AR41_to_Breakage_Min));
  const avgBrk =avgArr(_anom.map(r=>r.Breakage_to_Processed_Min));
  document.getElementById('quickStats').innerHTML=[
    ['Total delays',     total,                    ''],
    ['S-Power',          sPower,                   CLR.peach],
    ['S-Axis',           sAxis,                    CLR.peri],
    ['Avg AR41→Brk',     fmtMin(avgAR41),          CLR.teal],
    ['Avg Brk→Proc',     fmtMin(avgBrk),           ''],
    ['Worst job',        worst?fmtMin(worst._v):'--', CLR.peach],
  ].map(([label,val,color])=>`
    <div class="stat-row">
      <span class="stat-label">${label}</span>
      <span class="stat-val" style="color:${color||'inherit'}">${val}</span>
    </div>`).join('');
}

/* ============================================================
   LOAD ALL
   ============================================================ */
async function loadAll() {
  document.getElementById('liveStatus').textContent='Loading...';
  document.getElementById('liveDot').style.background=CLR.peri;
  try {
    const [brk,reason,research,anom]=await Promise.all([
      fetchTab('breakageSummary'),fetchTab('reasonSummary'),
      fetchTab('powerResearch'),fetchTab('anomalies')
    ]);
    _brk     =Array.isArray(brk)     ?brk     :[];
    _reason  =Array.isArray(reason)  ?reason  :[];
    _research=Array.isArray(research)?research:[];
    _anom    =Array.isArray(anom)    ?anom    :[];

    buildOverview();
    populateFlowFilters(); renderFlow();
    populateResearchFilters(); renderResearch();
    buildAlerts();

    document.getElementById('liveStatus').textContent='Live';
    document.getElementById('liveDot').style.background='#d4c0a8';
    document.getElementById('liveDot').style.animation='pulse 1.5s infinite';
  } catch(err) {
    console.error(err);
    document.getElementById('liveStatus').textContent='Error';
    document.getElementById('liveDot').style.background='#f87171';
    document.getElementById('liveDot').style.animation='none';
  }
}

setInterval(loadAll, 5*60*1000);
loadAll();