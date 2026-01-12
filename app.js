function setStateBar(el, state){
  if(!el) return;
  if(state==="green") el.style.background="linear-gradient(90deg,#16a34a,#22c55e)";
  if(state==="yellow") el.style.background="linear-gradient(90deg,#f59e0b,#92400e)";
  if(state==="red") el.style.background="linear-gradient(90deg,#ef4444,#7f1d1d)";
}

function flash(el){
  if(!el) return;
  el.classList.add("flash");
  setTimeout(()=>el.classList.remove("flash"),250);
}
