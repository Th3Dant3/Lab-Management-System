const dropzone   = document.getElementById("dropzone");
const fileInput  = document.getElementById("fileInput");
const sheetSelect= document.getElementById("sheetSelect");
const previewBtn = document.getElementById("previewBtn");
const downloadBtn= document.getElementById("downloadBtn");
const statusEl   = document.getElementById("status");
const previewEl  = document.getElementById("preview");
const columnBox  = document.getElementById("columnPicker");

let workbook=null, rows=[], headers=[];
let lastCsv="", lastName="output.csv";

function setStatus(msg,type="warn"){
  statusEl.className=`status ${type}`;
  statusEl.textContent=msg;
}

function escapeCsv(v){
  v=(v??"").toString();
  if(v.includes(",")||v.includes('"')||v.includes("\n"))
    return `"${v.replace(/"/g,'""')}"`;
  return v;
}

/* Drag & Drop */
dropzone.onclick=()=>fileInput.click();
["dragenter","dragover"].forEach(e=>{
  dropzone.addEventListener(e,ev=>{ev.preventDefault();dropzone.classList.add("drag")});
});
["dragleave","drop"].forEach(e=>{
  dropzone.addEventListener(e,ev=>{ev.preventDefault();dropzone.classList.remove("drag")});
});
dropzone.addEventListener("drop",e=>{
  fileInput.files=e.dataTransfer.files;
  loadFile();
});
fileInput.addEventListener("change",loadFile);

async function loadFile(){
  const file=fileInput.files[0];
  if(!file)return;
  try{
    const buf=await file.arrayBuffer();
    workbook=XLSX.read(buf,{type:"array"});
    sheetSelect.innerHTML="";
    workbook.SheetNames.forEach(s=>{
      const o=document.createElement("option");
      o.value=s;o.textContent=s;
      sheetSelect.appendChild(o);
    });
    sheetSelect.disabled=false;
    previewBtn.disabled=false;
    loadSheet();
    setStatus(`Loaded ${file.name}`,"ok");
  }catch(e){
    setStatus("Failed to load file","err");
  }
}

sheetSelect.onchange=loadSheet;

function loadSheet(){
  const ws=workbook.Sheets[sheetSelect.value];
  rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",blankrows:false});
  headers=rows[0]||[];
  columnBox.innerHTML="";
  headers.forEach((h,i)=>{
    const l=document.createElement("label");
    l.innerHTML=`<input type="checkbox" checked value="${i}"> ${h||"(blank)"}`;
    columnBox.appendChild(l);
  });
  previewEl.textContent=rows.slice(0,10).map(r=>r.join(" | ")).join("\n");
}

/* CSV */
previewBtn.onclick=()=>{
  const idxs=[...columnBox.querySelectorAll("input:checked")].map(i=>+i.value);
  const fmt=document.querySelector("input[name=fmt]:checked").value;
  const out=[];
  for(let i=1;i<rows.length;i++){
    const r=idxs.map(j=>escapeCsv(rows[i][j]));
    if(r.join("").trim()) out.push(r.join(","));
  }
  if(!out.length){setStatus("CSV empty","warn");return;}
  lastCsv=out.join("\n");
  lastName=`${fmt}_output.csv`;
  previewEl.textContent=lastCsv.split("\n").slice(0,12).join("\n");
  downloadBtn.disabled=false;
  setStatus("CSV ready","ok");
};

downloadBtn.onclick=()=>{
  const b=new Blob([lastCsv],{type:"text/csv"});
  const a=document.createElement("a");
  a.href=URL.createObjectURL(b);
  a.download=lastName;
  a.click();
};
