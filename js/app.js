const WS="e054d32b-402e-40d2-93ba-d4c84e859364";
const DS="eeefeaf2-4896-4450-afa1-3083f2ba7032";
const TOOL="mcp__Power_BI_Semantic_Models__powerbi_run_dax_query";
let L=DATA.locs, OBS=DATA.obs, FC=DATA.fc, META=DATA.meta;
const PAL=["#006747","#4C7422","#F29000","#7E9B53","#DB8143","#CCB879"];
const O={li:0,d:1,mn:2,mx:3,me:4,rn:5,su:6,wd:7,gu:8};
const F={li:0,d:1,mn:2,mx:3,me:4,rn:5,rp:6,su:7,wd:8,gu:9,pw:10,rh:11};
let locCountry,locName,locGrower,AGG;
const state={country:"",location:"",histDays:90,histStart:null,histEnd:null,active:"map",crops:[],cropGrowers:[],cropCountry:"",cropLocation:"",cropMonths:12,loaded:{forecast:false,historical:false,crop:false,cropplan:false,bycountry:false,map:false,accuracy:false,refs:false,actioned:false,about:false,alerts:false}};
const charts={};

const MON=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
function fmtDate(v){if(!v)return"";const p=String(v).slice(0,10).split("-");if(p.length<3)return v;return (+p[2])+" "+MON[(+p[1])-1]+" "+p[0];}
function num(v,dp=1){if(v==null||v==="")return"—";const n=Number(v);if(isNaN(n))return"—";return n.toLocaleString("en-GB",{minimumFractionDigits:dp,maximumFractionDigits:dp});}
function intf(v){if(v==null)return"—";return Number(v).toLocaleString("en-GB");}
function avg(a){const f=a.filter(x=>x!=null);return f.length?f.reduce((p,c)=>p+c,0)/f.length:null;}
function round(v,dp=1){if(v==null)return null;const m=Math.pow(10,dp);return Math.round(v*m)/m;}
function matchIds(){const s=new Set();for(let i=0;i<L.length;i++){if((!state.country||locCountry[i]===state.country)&&(!state.location||(locGrower[i]+"||"+locName[i]+"||"+locCountry[i])===state.location))s.add(i);}return s;}

/* charts / grid — resilient to library load timing */
function setChartDefaults(){if(window.Chart){Chart.defaults.font.family="'Calibri','Segoe UI',Arial,sans-serif";Chart.defaults.color="#594843";Chart.defaults.font.size=11;}}
function mk(id,cfg){if(!window.Chart){console.log("mk skipped, no Chart",id);return;}if(charts[id])charts[id].destroy();charts[id]=new Chart(document.getElementById(id),cfg);}
function grid(elId,columns,data,opts){opts=opts||{};const el=document.getElementById(elId);if(!el)return;el.innerHTML="";if(!window.gridjs){el.innerHTML='<div class="note">Table library unavailable.</div>';return;}const inner=document.createElement("div");el.appendChild(inner);columns=columns.map(c=>{if(c&&typeof c==="object"&&typeof c.width==="string"&&c.width.indexOf("%")>=0){c=Object.assign({},c);delete c.width;}return c;});columns=columns.map((col,ci)=>{const nm=(col&&typeof col==="object"?col.name:col);const low=nm?String(nm).toLowerCase():"";if(low==="year")return col;const someNum=data.some(r=>typeof r[ci]==="number");const allNum=data.length&&data.every(r=>r[ci]==null||r[ci]===""||typeof r[ci]==="number");let cls=null;if(someNum&&allNum)cls="numcell";else if(low==="grower"||low==="country"||low==="growers"||low==="growers / sites")cls="txtnowrap";if(!cls)return col;const b=(col&&typeof col==="object")?Object.assign({},col):{name:col};b.attributes=function(cell,row){return (cell==null&&row==null)?{}:{class:"gridjs-td "+cls};};return b;});const st=opts.wrap?{table:{width:"100%","table-layout":"fixed"},td:{"white-space":"normal","vertical-align":"top","word-break":"break-word"}}:{table:{"white-space":"nowrap"}};new gridjs.Grid({columns,data,resizable:opts.resizable!==false,search:opts.search!==false,sort:true,pagination:{limit:opts.limit||12},style:st,language:{search:{placeholder:"Filter…"}}}).render(inner);}
function whenReady(cb){let n=0;(function chk(){if(window.Chart&&window.gridjs){setChartDefaults();console.log("DASH libs ready after "+(n*50)+"ms");cb();}else if(n++<160){setTimeout(chk,50);}else{console.log("DASH libs TIMEOUT Chart="+!!window.Chart+" gridjs="+!!window.gridjs);setChartDefaults();cb();}})();}
function baseOpts(extra){return Object.assign({responsive:true,maintainAspectRatio:false,interaction:{mode:"index",intersect:false},plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}}},extra||{});}
function kpi(cls,val,label){return '<div class="kpi '+cls+'"><div class="v">'+val+'</div><div class="l">'+label+'</div></div>';}

function groupByDate(rows,spec){
  const m=new Map();
  for(const r of rows){const d=r[1];let o=m.get(d);
    if(!o){o={date:d};spec.forEach(s=>{o["_s_"+s.k]=0;o["_n_"+s.k]=0;if(s.f==="max")o["_mx_"+s.k]=-Infinity;});m.set(d,o);}
    spec.forEach(s=>{const v=r[s.i];if(v==null)return;o["_s_"+s.k]+=v;o["_n_"+s.k]++;if(s.f==="max"&&v>o["_mx_"+s.k])o["_mx_"+s.k]=v;});}
  const arr=[...m.values()].sort((a,b)=>a.date<b.date?-1:1);
  arr.forEach(o=>spec.forEach(s=>{o[s.k]=s.f==="sum"?o["_s_"+s.k]:s.f==="max"?(o["_mx_"+s.k]===-Infinity?null:o["_mx_"+s.k]):(o["_n_"+s.k]?o["_s_"+s.k]/o["_n_"+s.k]:null);}));
  return arr;}
function decimate(arr){if(arr.length<=400)return arr;const step=Math.ceil(arr.length/400);return arr.filter((_,i)=>i%step===0);}

/* FORECAST */
function fcAccCard(){
  if(typeof ACC==="undefined"||typeof accCompute!=="function")return "";
  const cur=accCompute(accScopeFilter(),ACC.leadMin,ACC.leadMax);
  const tot=accTotal(cur.vars,true);
  if(!cur.n||tot==null)return kpi("earth","&mdash;","Forecast accuracy");
  return '<div class="kpi kpi-acc" id="fc-acc-kpi" title="Open the Forecast accuracy tab for this selection"><div class="v">'+num(tot,1)+'%</div><div class="l">Forecast accuracy &#8594;</div></div>';
}
function loadForecast(){
  const ids=matchIds();let rows=FC.filter(r=>ids.has(r[F.li]));(function(){const b={};rows.forEach(r=>{(b[r[F.li]]=b[r[F.li]]||[]).push(r);});rows=[];Object.keys(b).forEach(k=>{b[k].sort((x,y)=>x[F.d]<y[F.d]?1:-1);rows=rows.concat(b[k].slice(0,14));});})();const kp=document.getElementById("fc-kpis");
  if(!rows.length){kp.innerHTML='<div class="err">No forecast data for this selection.</div>';["fcTemp","fcRain","fcWind"].forEach(id=>{if(charts[id])charts[id].destroy();});document.getElementById("fcTable").innerHTML="";state.loaded.forecast=true;return;}
  const locs=new Set(rows.map(r=>r[F.li])).size;
  const mxv=rows.map(r=>r[F.mx]).filter(v=>v!=null),mnv=rows.map(r=>r[F.mn]).filter(v=>v!=null),rnv=rows.map(r=>r[F.rn]).filter(v=>v!=null);
  const _hi=a=>a.length?Math.max(...a):null,_lo=a=>a.length?Math.min(...a):null;
  kp.innerHTML=kpi("",intf(locs),"Locations in view")
    +kpi("orange",num(_hi(mxv))+" °C","Highest max temp")
    +kpi("orange",num(_lo(mxv))+" °C","Lowest max temp")
    +kpi("",num(_hi(mnv))+" °C","Highest min temp")
    +kpi("",num(_lo(mnv))+" °C","Lowest min temp")
    +kpi("omni",num(_hi(rnv))+" mm","Highest daily rainfall")
    +kpi("omni",num(avg(rnv))+" mm","Avg daily rainfall")
    +kpi("earth",num(Math.max(...rows.map(r=>r[F.gu]||0)),0)+" km/h","Peak wind gust")
    +kpi("orange",num(avg(rows.map(r=>r[F.rp])),0)+" %","Avg rain probability")
    +fcAccCard();
  const accEl=document.getElementById("fc-acc-kpi");if(accEl)accEl.onclick=function(){state.loaded.accuracy=false;switchTab("accuracy");};
  const t=groupByDate(rows,[{k:"mx",i:F.mx,f:"avg"},{k:"mn",i:F.mn,f:"avg"},{k:"me",i:F.me,f:"avg"},{k:"rn",i:F.rn,f:"avg"},{k:"rp",i:F.rp,f:"avg"},{k:"wd",i:F.wd,f:"avg"},{k:"gu",i:F.gu,f:"max"}]);
  const labels=t.map(x=>fmtDate(x.date));
  mk("fcTemp",{type:"line",data:{labels,datasets:[
    {label:"Max",data:t.map(x=>x.mx),borderColor:PAL[2],backgroundColor:PAL[2],tension:.3,pointRadius:2},
    {label:"Mean",data:t.map(x=>x.me),borderColor:PAL[0],backgroundColor:PAL[0],tension:.3,pointRadius:2},
    {label:"Min",data:t.map(x=>x.mn),borderColor:PAL[1],backgroundColor:PAL[1],tension:.3,pointRadius:2}
  ]},options:baseOpts({scales:{y:{title:{display:true,text:"°C"}}}})});
  mk("fcRain",{data:{labels,datasets:[
    {type:"bar",label:"Rainfall (mm)",data:t.map(x=>x.rn),backgroundColor:PAL[1],yAxisID:"y"},
    {type:"line",label:"Rain probability (%)",data:t.map(x=>x.rp),borderColor:PAL[2],backgroundColor:PAL[2],tension:.3,pointRadius:2,yAxisID:"y1"}
  ]},options:baseOpts({scales:{y:{position:"left",title:{display:true,text:"mm"},beginAtZero:true},y1:{position:"right",title:{display:true,text:"%"},min:0,max:100,grid:{drawOnChartArea:false}}}})});
  mk("fcWind",{type:"line",data:{labels,datasets:[
    {label:"Wind speed (avg)",data:t.map(x=>x.wd),borderColor:PAL[0],backgroundColor:PAL[0],tension:.3,pointRadius:2},
    {label:"Peak gust",data:t.map(x=>x.gu),borderColor:PAL[2],backgroundColor:PAL[2],tension:.3,pointRadius:2,borderDash:[5,3]}
  ]},options:baseOpts({scales:{y:{title:{display:true,text:"km/h"},beginAtZero:true}}})});
  const tab=rows.slice().sort((a,b)=>a[F.d]<b[F.d]?-1:(a[F.d]>b[F.d]?1:locName[a[F.li]].localeCompare(locName[b[F.li]])));
  grid("fcTable",[
    {name:"Date",formatter:c=>fmtDate(c)},{name:"Grower"},{name:"Location"},{name:"Country"},
    {name:"Min °C",formatter:c=>num(c)},{name:"Max °C",formatter:c=>num(c)},{name:"Mean °C",formatter:c=>num(c)},
    {name:"Rain mm",formatter:c=>num(c)},{name:"Rain %",formatter:c=>num(c,0)},
    {name:"Wind km/h",formatter:c=>num(c,0)},{name:"Gust km/h",formatter:c=>num(c,0)},{name:"Wind dir"}
  ],tab.map(r=>[r[F.d],locGrower[r[F.li]],locName[r[F.li]],locCountry[r[F.li]],r[F.mn],r[F.mx],r[F.me],r[F.rn],r[F.rp],r[F.wd],r[F.gu],r[F.pw]]),{search:false,limit:14});
  console.log("DASH forecast rendered charts:"+(!!charts.fcTemp)+"/"+(!!charts.fcRain));
  state.loaded.forecast=true;
}

/* HISTORICAL */
function histRange(){
  const end=state.histEnd||META.obs_dates[1];let start;
  if(state.histStart)start=state.histStart;
  else if(state.histDays===0)start=META.obs_dates[0];
  else{const e=new Date(end);e.setDate(e.getDate()-state.histDays);start=e.toISOString().slice(0,10);}
  return {start,end};
}
function loadHistorical(){
  const ids=matchIds();const rng=histRange();
  const rows=OBS.filter(r=>ids.has(r[O.li])&&r[O.d]>=rng.start&&r[O.d]<=rng.end);const kp=document.getElementById("hs-kpis");
  if(!rows.length){kp.innerHTML='<div class="err">No observations for this selection/period.</div>';["hsTemp","hsRain","hsSun","hsWind"].forEach(id=>{if(charts[id])charts[id].destroy();});document.getElementById("hsTable").innerHTML="";state.loaded.historical=true;return;}
  kp.innerHTML=kpi("",num(avg(rows.map(r=>r[O.me])))+" °C","Avg mean temp")
    +kpi("orange",num(Math.max(...rows.map(r=>r[O.mx]??-99)))+" °C","Highest max temp")
    +kpi("",num(Math.min(...rows.map(r=>r[O.mn]??99)))+" °C","Lowest min temp")
    +kpi("omni",num(avg(rows.map(r=>r[O.rn])))+" mm","Avg daily rainfall")
    +kpi("earth",num(avg(rows.map(r=>r[O.su])))+" hrs","Avg sunshine / day")
    +kpi("",intf(rows.length),"Observations");
  const g=groupByDate(rows,[{k:"mx",i:O.mx,f:"avg"},{k:"mn",i:O.mn,f:"avg"},{k:"me",i:O.me,f:"avg"},{k:"rn",i:O.rn,f:"sum"},{k:"su",i:O.su,f:"avg"},{k:"wd",i:O.wd,f:"avg"},{k:"gu",i:O.gu,f:"avg"},{k:"gux",i:O.gu,f:"max"}]);
  const gd=decimate(g);const labels=gd.map(x=>fmtDate(x.date));
  mk("hsTemp",{type:"line",data:{labels,datasets:[
    {label:"Max",data:gd.map(x=>x.mx),borderColor:PAL[2],backgroundColor:PAL[2],pointRadius:0,tension:.25},
    {label:"Mean",data:gd.map(x=>x.me),borderColor:PAL[0],backgroundColor:PAL[0],pointRadius:0,tension:.25},
    {label:"Min",data:gd.map(x=>x.mn),borderColor:PAL[1],backgroundColor:PAL[1],pointRadius:0,tension:.25}
  ]},options:baseOpts({scales:{y:{title:{display:true,text:"°C"}},x:{ticks:{maxTicksLimit:12,autoSkip:true}}}})});
  const mm={};rows.forEach(r=>{const key=r[O.d].slice(0,7);if(!mm[key])mm[key]={rain:0,sites:new Set()};mm[key].rain+=(r[O.rn]||0);mm[key].sites.add(r[O.li]);});
  const mkeys=Object.keys(mm).sort();
  mk("hsRain",{type:"bar",data:{labels:mkeys.map(k=>{const p=k.split("-");return MON[+p[1]-1]+" "+p[0].slice(2);}),
    datasets:[{label:"Rainfall (mm)",data:mkeys.map(k=>mm[k].rain/mm[k].sites.size),backgroundColor:PAL[1]}]},
    options:baseOpts({plugins:{legend:{display:false}},scales:{y:{title:{display:true,text:"mm"},beginAtZero:true},x:{ticks:{maxTicksLimit:14,autoSkip:true}}}})});
  mk("hsSun",{type:"line",data:{labels,datasets:[{label:"Sunshine",data:gd.map(x=>x.su),borderColor:PAL[2],backgroundColor:"rgba(242,144,0,.15)",fill:true,pointRadius:0,tension:.25}]},
    options:baseOpts({plugins:{legend:{display:false}},scales:{y:{title:{display:true,text:"hours"},beginAtZero:true},x:{ticks:{maxTicksLimit:12,autoSkip:true}}}})});
  mk("hsWind",{type:"line",data:{labels,datasets:[{label:"Wind speed",data:gd.map(x=>x.wd),borderColor:PAL[3],backgroundColor:PAL[3],pointRadius:0,tension:.25},{label:"Gust (avg)",data:gd.map(x=>x.gu),borderColor:PAL[4],backgroundColor:PAL[4],pointRadius:0,tension:.25},{label:"Gust (peak)",data:gd.map(x=>x.gux),borderColor:PAL[2],backgroundColor:PAL[2],borderDash:[4,3],pointRadius:0,tension:.25}]},
    options:baseOpts({scales:{y:{title:{display:true,text:"km/h"},beginAtZero:true},x:{ticks:{maxTicksLimit:12,autoSkip:true}}}})});
  const tab=rows.slice().sort((a,b)=>a[O.d]>b[O.d]?-1:1).slice(0,300);
  grid("hsTable",[
    {name:"Date",formatter:c=>fmtDate(c)},{name:"Grower"},{name:"Location"},{name:"Country"},
    {name:"Min °C",formatter:c=>num(c)},{name:"Max °C",formatter:c=>num(c)},{name:"Mean °C",formatter:c=>num(c)},
    {name:"Rain mm",formatter:c=>num(c)},{name:"Sun hrs",formatter:c=>num(c)},{name:"Wind km/h",formatter:c=>num(c,0)},{name:"Gust km/h",formatter:c=>num(c,0)}
  ],tab.map(r=>[r[O.d],locGrower[r[O.li]],locName[r[O.li]],locCountry[r[O.li]],r[O.mn],r[O.mx],r[O.me],r[O.rn],r[O.su],r[O.wd],r[O.gu]]));
  state.loaded.historical=true;
}

/* ASK */
const EXAMPLES=["Average rainfall in Kenya over the last year","Which country is hottest in the current forecast?","Compare mean temperature: Spain vs Morocco (2025)","Total rainfall in the UK in 2024","Highest wind gust in the current forecast"];
function buildAgg(){
  const cy={};
  OBS.forEach(r=>{const c=locCountry[r[O.li]],y=r[O.d].slice(0,4),k=c+"|"+y;
    if(!cy[k])cy[k]={country:c,year:+y,_me:[],_mx:[],_mn:[],_su:[],rain:0,sites:new Set()};
    const o=cy[k];o._me.push(r[O.me]);o._mx.push(r[O.mx]);o._mn.push(r[O.mn]);o._su.push(r[O.su]);o.rain+=(r[O.rn]||0);o.sites.add(r[O.li]);});
  const obsByCountryYear=Object.values(cy).map(o=>({country:o.country,year:o.year,avg_mean_temp_C:round(avg(o._me)),avg_max_temp_C:round(avg(o._mx)),avg_min_temp_C:round(avg(o._mn)),avg_sunshine_hrs:round(avg(o._su)),avg_annual_rain_mm_per_site:round(o.rain/o.sites.size,0),sites:o.sites.size}));
  const fcm={};
  FC.forEach(r=>{const key=locName[r[F.li]];if(!fcm[key])fcm[key]={location:key,country:locCountry[r[F.li]],_mx:[],_mn:[],_rn:[],_rp:[],gust:0};
    const o=fcm[key];o._mx.push(r[F.mx]);o._mn.push(r[F.mn]);o._rn.push(r[F.rn]);o._rp.push(r[F.rp]);o.gust=Math.max(o.gust,r[F.gu]||0);});
  const forecastByLocation=Object.values(fcm).map(o=>({location:o.location,country:o.country,forecast_avg_max_C:round(avg(o._mx)),forecast_avg_min_C:round(avg(o._mn)),forecast_avg_daily_rain_mm:round(avg(o._rn)),forecast_avg_rain_prob_pc:round(avg(o._rp),0),forecast_peak_gust_kmh:round(o.gust,0)}));
  return {obsByCountryYear,forecastByLocation};
}
function rebuildLookups(){locCountry=L.map(x=>x[2]);locName=L.map(x=>x[1]);locGrower=L.map(x=>x[0]);AGG=buildAgg();}
function scopeLabel(){if(state.location){const _p=state.location.split("||");return _p[0]+" — "+_p[1]+(_p[2]?", "+_p[2]:"");}if(state.country)return state.country+" (all sites)";return "all sites, all countries";}
function scopedObsByYear(){
  const ids=matchIds();const cy={};
  OBS.forEach(r=>{if(!ids.has(r[O.li]))return;const y=r[O.d].slice(0,4);let o=cy[y];if(!o){o=cy[y]={year:+y,_me:[],_mx:[],_mn:[],_su:[],rain:0,sites:new Set()};}o._me.push(r[O.me]);o._mx.push(r[O.mx]);o._mn.push(r[O.mn]);o._su.push(r[O.su]);o.rain+=(r[O.rn]||0);o.sites.add(r[O.li]);});
  return Object.values(cy).sort((a,b)=>a.year-b.year).map(o=>({year:o.year,avg_mean_temp_C:round(avg(o._me)),avg_max_temp_C:round(avg(o._mx)),avg_min_temp_C:round(avg(o._mn)),avg_sunshine_hrs:round(avg(o._su)),avg_annual_rain_mm_per_site:round(o.rain/o.sites.size,0),sites:o.sites.size}));
}
function scopedForecast(){
  const ids=matchIds();const fm={};
  FC.forEach(r=>{if(!ids.has(r[F.li]))return;let o=fm[r[F.li]];if(!o){o=fm[r[F.li]]={location:locName[r[F.li]],grower:locGrower[r[F.li]],country:locCountry[r[F.li]],_mx:[],_mn:[],_rn:[],_rp:[],gust:0};}o._mx.push(r[F.mx]);o._mn.push(r[F.mn]);o._rn.push(r[F.rn]);o._rp.push(r[F.rp]);o.gust=Math.max(o.gust,r[F.gu]||0);});
  return Object.values(fm).sort((a,b)=>a.location.localeCompare(b.location)).map(o=>({location:o.location,grower:o.grower,country:o.country,forecast_avg_max_C:round(avg(o._mx)),forecast_avg_min_C:round(avg(o._mn)),forecast_avg_daily_rain_mm:round(avg(o._rn)),forecast_avg_rain_prob_pc:round(avg(o._rp),0),forecast_peak_gust_kmh:round(o.gust,0)}));
}
function commitTopScope(){state.country=(document.getElementById("fCountry")||{}).value||"";state.location=(document.getElementById("fLocation")||{}).value||"";}
function updateAskScope(){const el=document.getElementById("askScope");if(el)el.textContent="Answers use dashboard data for: "+scopeLabel()+". Change the country/location filter above to re-scope.";}

async function askClaudeSafe(prompt,data){
  if(!(window.cowork&&typeof window.cowork.askClaude==="function"))return null;
  try{const r=await window.cowork.askClaude(prompt,data||[]);
    if(typeof r==="string")return r;if(r&&typeof r==="object")return r.text||r.content||r.answer||r.result||null;return String(r);
  }catch(e){console.log("askClaude failed",e.message);return null;}
}
function fallbackAnswer(q){
  const ql=q.toLowerCase();
  const wantFc=/forecast|next|coming|upcoming|week|ahead/.test(ql);
  const wantRain=/rain|precip|wet|dry/.test(ql),wantTemp=/temp|hot|cold|warm|degree/.test(ql),wantWind=/wind|gust/.test(ql);
  let rows=[],cols=[],sentence="";
  if(wantFc){
    rows=scopedForecast();cols=rows.length?Object.keys(rows[0]):[];
    if(rows.length){
      if(wantWind){const t=[...rows].sort((a,b)=>b.forecast_peak_gust_kmh-a.forecast_peak_gust_kmh)[0];sentence="Highest forecast gust for "+scopeLabel()+": "+t.location+" ("+t.grower+") at "+t.forecast_peak_gust_kmh+" km/h.";}
      else if(wantTemp||/hot/.test(ql)){const t=[...rows].sort((a,b)=>b.forecast_avg_max_C-a.forecast_avg_max_C)[0];sentence="Warmest forecast for "+scopeLabel()+": "+t.location+" ("+t.grower+"), averaging "+t.forecast_avg_max_C+" °C max.";}
      else sentence="Current forecast for "+scopeLabel()+".";
    }
  }else{
    let ag=scopedObsByYear();const ym=ql.match(/20(2[0-6])/);if(ym)ag=ag.filter(r=>r.year===+("20"+ym[1]));
    rows=ag;cols=rows.length?Object.keys(rows[0]):[];
    if(rows.length)sentence=(wantRain?"Rainfall (average annual mm per site)":wantTemp?"Temperature summary":"Summary")+" for "+scopeLabel()+" by year.";
  }
  if(!rows.length)sentence="No data for the current selection ("+scopeLabel()+").";
  return {sentence,rows,cols};
}
async function runAsk(q){
  commitTopScope();updateAskScope();
  const out=document.getElementById("askOut");document.getElementById("askInput").value=q;
  out.innerHTML='<div class="loading"><span class="spin"></span>Working…</div>';document.getElementById("askBtn").disabled=true;
  const fb=fallbackAnswer(q);
  const payload=[{summary:"observations_by_year (scope: "+scopeLabel()+")",rows:scopedObsByYear()},{summary:"current_forecast (scope: "+scopeLabel()+")",rows:scopedForecast()}];
  const prompt="You are a horticulture weather analyst for Flamingo. The data provided is limited to the current dashboard selection: "+scopeLabel()+" — answer only for this selection and do not refer to other sites or countries. Answer in 2-4 sentences, UK English, metric units (°C, mm, km/h). Blend TWO sources: (1) the supplied Flamingo Weather figures (observations to "+META.obs_dates[1]+"; current forecast "+META.fc_dates[0]+" to "+META.fc_dates[1]+"), and (2) your own general knowledge of climate, geography, seasonality and horticulture for context and interpretation. Briefly signal which insights come from the data vs general knowledge. You do not have live web access in this view. Question: "+q;
  const ans=await askClaudeSafe(prompt,payload);
  let html='<div class="answer">'+((ans||fb.sentence||"Here is the most relevant data.").replace(/</g,"&lt;"))+'</div>';
  if(fb.rows&&fb.rows.length){html+='<div id="askGrid"></div>';out.innerHTML=html;
    grid("askGrid",fb.cols.map(c=>({name:c.replace(/_/g," "),formatter:cell=>c==="year"?String(cell):(typeof cell==="number"?num(cell,c.includes("prob")||c.includes("per_site")||c.includes("gust")?0:1):cell)})),fb.rows.map(r=>fb.cols.map(c=>r[c])),{search:false});
  }else out.innerHTML=html;
  document.getElementById("askBtn").disabled=false;
}

/* web-blended research via main Claude chat */
function getSendPrompt(){
  const cands=[];
  try{if(typeof window.sendPrompt==="function")cands.push(window.sendPrompt);}catch(e){}
  try{if(window.cowork&&typeof window.cowork.sendPrompt==="function")cands.push(window.cowork.sendPrompt.bind(window.cowork));}catch(e){}
  try{if(window.parent&&typeof window.parent.sendPrompt==="function")cands.push(window.parent.sendPrompt);}catch(e){}
  return cands[0]||null;
}
function relevantData(){
  return {selection:scopeLabel(), observations_by_year:scopedObsByYear(), current_forecast:scopedForecast()};
}
function composeWebPrompt(q){
  const ctx=relevantData();
  return "Using the Flamingo Weather dashboard data below (limited to the current dashboard selection: "+scopeLabel()+") TOGETHER WITH current information from the web, answer the question for this selection. Combine both sources, add relevant external context (e.g. seasonal norms, regional events, agronomic implications) and cite any web sources.\n\nQUESTION: "+q
    +"\n\nDASHBOARD DATA (snapshot "+fmtDate(META.generated)+", temperatures °C, rainfall mm, wind km/h, sunshine hours):\n"+JSON.stringify(ctx);
}
async function askWeb(q){
  commitTopScope();updateAskScope();
  const out=document.getElementById("askOut");const prompt=composeWebPrompt(q);
  const sp=getSendPrompt();
  if(sp){try{sp(prompt);out.innerHTML='<div class="answer">Sent to Claude in the chat &#8594; a web-researched answer that blends these figures with live external sources will appear in your conversation.</div>';return;}catch(e){}}
  out.innerHTML='<div class="answer">This dashboard view can\'t reach the web directly. Copy the prompt below into the Claude chat and Claude will blend this data with live web research:</div><textarea class="copybox" readonly onclick="this.select()">'+prompt.replace(/</g,"&lt;")+'</textarea><div class="srcnote">Tip: click the box to select all, then copy.</div>';
}
/* ---- CROP RISK ---- */
/* map sites come from CROP.sites, which is the verified Weather_locations mapping */
const CROP_RULES=CROP.rules;
/* wind thresholds now come from CROP.rules, set per crop from the crop climate risk workbook */
const GOLD={"Egypt||Extra fine beans":12,"Egypt||Fine beans":15,"Egypt||Green beans":20,"Egypt||Mangetout":7,"Egypt||Sugarsnap":7,"Egypt||Tenderstem broccoli":8,"Mexico||Asparagus":15,"Morocco||Chilli":22,"Morocco||Extra fine beans":14,"Morocco||Fine beans":18,"Morocco||Green beans":26,"Morocco||Helda (flat) beans":26,"Morocco||Tenderstem broccoli":8,"UK||Asparagus":8,"UK||Broad beans":10,"UK||Extra fine beans":10,"UK||Fine beans":11,"UK||Green beans":12,"UK||Helda (flat) beans":14,"UK||Runner beans":16,"UK||Sugarsnap":6,"UK||Tenderstem broccoli":7,"India||Baby aubergine":18,"India||Baby corn":2.5,"India||Okra":14,"India||Turmeric":30,"Peru||Asparagus":20,"Peru||Mangetout":8,"Peru||Sugarsnap":8,"Kenya||Baby aubergine":16,"Kenya||Baby corn":2.5,"Kenya||Extra fine beans":13,"Kenya||Fine beans":15,"Kenya||Mangetout":7,"Kenya||Okra":10,"Kenya||Runner beans":17,"Kenya||Sugarsnap":7,"Kenya||Tenderstem broccoli":8,"Spain||Tenderstem broccoli":8};
const CROPTIER={"Asparagus":"A","Baby aubergine":"B","Baby carrots":"C","Baby corn":"B","Baby leeks":"C","Boston beans":"B","Broad beans":"A","Broccoli (mixed heads)":"B","Carrots":"C","Chilli":"B","Extra fine beans":"A","Fine beans":"A","Garden peas":"A","Green beans":"B","Helda (flat) beans":"B","Mangetout":"A","Okra":"C","Peas (shelled)":"A","Petit pois":"A","Rhubarb":"C","Runner beans":"A","Samphire":"C","Savannah beans":"B","Serenade Red (chilli)":"B","Stringless beans":"B","Sugarsnap":"A","Tenderstem broccoli":"B","Turmeric":"C","Yellow beans":"B"};
const LOSSFR={A:[0.1,0.3,0.6],B:[0.07,0.2,0.45],C:[0.04,0.12,0.3]};
const SEVLBL=["None","Mild","Medium","Severe"];
function goldFor(c,cr){const v=GOLD[c+"||"+cr];return v===undefined?null:v;}
function cropTierOf(cr){return CROPTIER[cr]||"B";}
function advSevRank(t,val,thr,thr2){if(val==null||thr==null)return 1;if(t==="Heat"){if(thr2!=null){if(val<thr2)return 1;const y=val-thr2;return y>3?3:2;}const x=val-thr;return x>6?3:(x>3?2:1);}if(t==="Frost"){const x=thr-val;return x>4?3:(x>2?2:1);}if(t==="High wind"){const x=val-thr;return x>30?3:(x>15?2:1);}const ratio=thr>0?val/thr:1;return ratio>2.5?3:(ratio>1.5?2:1);}
function lossFrac(tier,rank){return rank>0?(LOSSFR[tier]||LOSSFR.B)[rank-1]:0;}
function advThresholds(cropName){const ru=(typeof CROP_RULES!=="undefined")?CROP_RULES[cropTypeOf(cropName)]:null;if(!ru)return "";return "Heat >"+ru.heat+"\u00b0C"+(ru.heat2!=null?" (damage >"+ru.heat2+"\u00b0C)":"")+" \u00b7 Cold <"+ru.frost+"\u00b0C \u00b7 Heavy rain >"+ru.rain+"mm"+(ru.wind!=null?" \u00b7 High wind >"+ru.wind+" km/h":"");}
function cropTypeOf(name){const r=CROP.rows.find(x=>x.crop===name);return r?r.t:"greenbean";}
function growerCountry(g){const x=CROP.growers.find(y=>y.g===g);return x?x.c:"";}
function activeCrops(){return (state.crops&&state.crops.length)?state.crops.slice():CROP.crops.slice();}
function activeGrowers(){return (state.cropGrowers&&state.cropGrowers.length)?state.cropGrowers.slice():null;}
function activeCountries(){const g=activeGrowers();if(g)return [...new Set(g.map(growerCountry).filter(Boolean))];return state.cropCountry?[state.cropCountry]:null;}
function planRowsFor(crop){const g=activeGrowers();const cs=activeCountries();return CROP.rows.filter(x=>x.crop===crop&&(!g||g.indexOf(x.g)>=0)&&(!cs||cs.indexOf(x.c)>=0));}
function seasonMonthsFrom(rows){const s=new Set();rows.forEach(x=>x.gmonths.forEach(m=>s.add(m)));return s;}
function matchIdsForCountries(countrySet,cropSet){const s=new Set();for(let i=0;i<L.length;i++){if(cropSet&&!cropSet.has(locCountry[i]))continue;if(countrySet&&!countrySet.has(locCountry[i]))continue;s.add(i);}return s;}
function cropIdsFor(crop){const cs=activeCountries();const cset=cs?new Set(cs):null;const cc=new Set(CROP.rows.filter(x=>x.crop===crop).map(x=>x.c));return matchIdsForCountries(cset,cc);}
function refreshCropGrowers(){const sel=document.getElementById("cropGrower");if(!sel)return;const c=(document.getElementById("cropCountry")||{}).value||"";const prev=new Set([...sel.selectedOptions].map(o=>o.value));sel.innerHTML="";CROP.growers.filter(g=>!c||g.c===c).sort((a,b)=>a.g.localeCompare(b.g)).forEach(g=>{const o=document.createElement("option");o.value=g.g;o.textContent=g.g+" ("+g.c+")";if(prev.has(g.g))o.selected=true;sel.appendChild(o);});}
function cropHistRange(){const end=META.obs_dates[1];if(state.cropMonths===0)return{start:META.obs_dates[0],end};const e=new Date(end);e.setMonth(e.getMonth()-state.cropMonths);return{start:e.toISOString().slice(0,10),end};}
function cropEvents(rows,isFc,r,months){
  const MX=isFc?F.mx:O.mx,MN=isFc?F.mn:O.mn,RN=isFc?F.rn:O.rn,LI=isFc?F.li:O.li,WG=isFc?F.gu:O.gu;const ev=[];const useM=months&&months.size&&months.size<12;
  rows.forEach(row=>{
    if(useM&&!months.has(+row[1].slice(5,7)))return;
    if(row[MX]!=null&&row[MX]>r.heat)ev.push({date:row[1],loc:locName[row[LI]],grower:locGrower[row[LI]],country:locCountry[row[LI]],type:"Heat",value:row[MX],unit:"°C",thr:r.heat,thr2:r.heat2,issue:r.heatIssue});
    if(row[MN]!=null&&row[MN]<r.frost)ev.push({date:row[1],loc:locName[row[LI]],grower:locGrower[row[LI]],country:locCountry[row[LI]],type:"Frost",value:row[MN],unit:"°C",thr:r.frost,issue:r.frostIssue});
    if(row[RN]!=null&&row[RN]>r.rain)ev.push({date:row[1],loc:locName[row[LI]],grower:locGrower[row[LI]],country:locCountry[row[LI]],type:"Heavy rain",value:row[RN],unit:"mm",thr:r.rain,issue:r.rainIssue});
    if(r.wind!=null&&row[WG]!=null&&row[WG]>r.wind)ev.push({date:row[1],loc:locName[row[LI]],grower:locGrower[row[LI]],country:locCountry[row[LI]],type:"High wind",value:row[WG],unit:"km/h",thr:r.wind,issue:r.windIssue});
  });
  return ev;
}
function cropScopeLabel(){const g=activeGrowers();if(g)return g.join(", ")+" ("+[...new Set(g.map(growerCountry))].join("/")+")";return state.cropCountry?state.cropCountry+" (all sites)":"all countries";}
function refreshCropLocations(){const lSel=document.getElementById("cropLocation");if(!lSel)return;const selC=(document.getElementById("cropCountry")||{}).value||"";lSel.innerHTML='<option value="">All locations</option>';const seen=new Set();L.map((x,i)=>i).filter(i=>!selC||locCountry[i]===selC).sort((a,b)=>locName[a].localeCompare(locName[b])).forEach(i=>{if(seen.has(locName[i]))return;seen.add(locName[i]);const o=document.createElement("option");o.value=locName[i];o.textContent=locName[i]+" — "+locGrower[i];lSel.appendChild(o);});}
function dayGap(a,b){return Math.round((new Date(b)-new Date(a))/86400000);}
function summarise(events){
  const groups={};
  events.forEach(e=>{const k=e.country+"|"+e.type;(groups[k]=groups[k]||[]).push(e);});
  const out=[];
  Object.values(groups).forEach(list=>{
    const byDate={};
    list.forEach(e=>{if(!byDate[e.date])byDate[e.date]={g:new Set(),l:new Set(),v:[]};byDate[e.date].g.add(e.grower);byDate[e.date].l.add(e.loc);byDate[e.date].v.push(e.value);});
    const dates=Object.keys(byDate).sort();const type=list[0].type,issue=list[0].issue,country=list[0].country,unit=list[0].unit,thr=list[0].thr,thr2=list[0].thr2;
    let i=0;
    while(i<dates.length){
      let j=i;while(j+1<dates.length&&dayGap(dates[j],dates[j+1])<5)j++;
      const g=new Set(),l=new Set();let v=[],nd=0;
      for(let k=i;k<=j;k++){const b=byDate[dates[k]];b.g.forEach(x=>g.add(x));b.l.forEach(x=>l.add(x));v=v.concat(b.v);nd++;}
      const peak=type==="Frost"?Math.min.apply(null,v):Math.max.apply(null,v);
      out.push({start:dates[i],end:dates[j],country,type,issue,growers:[...g].sort(),sites:l.size,days:nd,peak,unit,thr,thr2});
      i=j+1;
    }
  });
  out.sort((a,b)=>a.start<b.start?-1:(a.start>b.start?1:a.country.localeCompare(b.country)));
  return out;
}
function updateCropScope(){const el=document.getElementById("cropScope");if(!el)return;const rng=cropHistRange();const cropLbl=(state.crops&&state.crops.length)?(state.crops.length>3?state.crops.length+" crops":state.crops.join(", ")):"all crops";el.textContent="Crops: "+cropLbl+" · applies to: "+cropScopeLabel()+" · forecast "+fmtDate(META.fc_dates[0])+"–"+fmtDate(META.fc_dates[1])+" · history "+fmtDate(rng.start)+"–"+fmtDate(rng.end)+".";}
function condLabel(e){if(e.type==="Heat")return "Heat (>"+e.thr+"°C)";if(e.type==="Frost")return "Frost (<"+e.thr+"°C)";if(e.type==="Heavy rain")return "Heavy rain (>"+e.thr+" mm)";if(e.type==="High wind")return "High wind (>"+e.thr+" km/h)";return e.type;}
function forecastMonths(){const s=new Set();const a=new Date(META.fc_dates[0]),b=new Date(META.fc_dates[1]);for(let d=new Date(a);d<=b;d.setDate(d.getDate()+1))s.add(d.getMonth()+1);return s;}
function cropCountryBreaches(crop){
  const rules=CROP.rules[cropTypeOf(crop)];const cropSet=new Set(CROP.rows.filter(x=>x.crop===crop).map(x=>x.c));
  const byC={};CROP.rows.filter(x=>x.crop===crop).forEach(x=>{(byC[x.c]=byC[x.c]||new Set());x.gmonths.forEach(m=>byC[x.c].add(m));});
  const bad=new Set();
  Object.keys(byC).forEach(c=>{const ids=matchIdsForCountries(new Set([c]),cropSet);const fcRows=FC.filter(r=>ids.has(r[F.li]));if(cropEvents(fcRows,true,rules,byC[c]).some(e=>advSevRank(e.type,e.value,e.thr,e.thr2)>=2))bad.add(c);});
  return bad;
}
function substituteGrowers(crop,alertCountry){
  const bad=cropCountryBreaches(crop);const fm=forecastMonths();
  const cands=CROP.rows.filter(x=>x.crop===crop&&x.c!==alertCountry&&!bad.has(x.c)&&x.gmonths.some(m=>fm.has(m)));
  const seen={};cands.forEach(x=>{if(!seen[x.g]||x.y>seen[x.g].y)seen[x.g]={g:x.g,c:x.c,y:x.y};});
  return Object.values(seen).sort((a,b)=>b.y-a.y);
}
function subLabel(e){const subs=substituteGrowers(e.crop,e.country);if(!subs.length)return "\u2014";return subs.slice(0,3).map(s=>s.g+" ("+s.c+")").join("; ");}
const ACT_URL="/api/actioned";
let ACT={};
function actStore(){return ACT;}
function actLocalSave(){try{localStorage.setItem("flamingoActioned",JSON.stringify(ACT));}catch(e){}}
function actLocalLoad(){try{return JSON.parse(localStorage.getItem("flamingoActioned")||"{}");}catch(e){return {};}}
function actLoad(){return fetch(ACT_URL,{cache:"no-store"}).then(r=>{if(!r.ok)throw 0;return r.json();}).then(j=>{ACT=j||{};actLocalSave();window.__actShared=true;return "server";}).catch(()=>{ACT=actLocalLoad();window.__actShared=false;return "local";});}
function actServerPut(eid,rec){return fetch(ACT_URL+"/"+encodeURIComponent(eid),{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(rec)}).then(r=>r.ok?r.json():null).then(j=>{if(j&&j.record){ACT[eid]=j.record;actLocalSave();}}).catch(()=>{});}
function actServerDelete(eid){return fetch(ACT_URL+"/"+encodeURIComponent(eid),{method:"DELETE"}).catch(()=>{});}
function eventId(e){return [e.crop,e.country,e.type,e.start,e.end].join("|");}
function isActioned(e){return Object.prototype.hasOwnProperty.call(ACT,eventId(e));}
function escH(x){return String(x==null?"":x).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function eventGrowers(e){const ag=activeGrowers();return [...new Set(CROP.rows.filter(x=>x.crop===e.crop&&x.c===e.country&&(!ag||ag.indexOf(x.g)>=0)).map(x=>x.g))].join(", ");}
function eventSnap(e){return {crop:e.crop,country:e.country,cond:condLabel(e),period:fmtDate(e.start)+(e.start===e.end?"":" – "+fmtDate(e.end)),sev:SEVLBL[advSevRank(e.type,e.peak,e.thr,e.thr2)],peak:num(e.peak,1)+" "+e.unit,growers:eventGrowers(e),issue:e.issue||"",start:e.start,end:e.end};}
let curEventSnap={};
function actChkCell(eid){return window.gridjs?window.gridjs.html('<input type="checkbox" class="actchk" data-eid="'+encodeURIComponent(eid)+'" title="Mark as actioned (hide from Crop risk)">'):"";}
function onActChange(ev){const t=ev.target;if(!t||!t.classList)return;
  if(t.classList.contains("actchk")){const eid=decodeURIComponent(t.getAttribute("data-eid")||"");if(!eid)return;
    if(t.checked){const rec=Object.assign({},curEventSnap[eid]||ACT[eid]||{},{note:(ACT[eid]&&ACT[eid].note)||"",actionedAt:(ACT[eid]&&ACT[eid].actionedAt)||new Date().toISOString().slice(0,10)});ACT[eid]=rec;actLocalSave();actServerPut(eid,rec);}
    else{delete ACT[eid];actLocalSave();actServerDelete(eid);}
    state.loaded.crop=false;state.loaded.actioned=false;
    if(state.active==="actioned")renderActioned();else loadCrop();
  } else if(t.classList.contains("actnote")){const eid=decodeURIComponent(t.getAttribute("data-eid")||"");if(ACT[eid]){ACT[eid].note=t.value;actLocalSave();actServerPut(eid,ACT[eid]);}}
}
if(typeof document!=="undefined"){document.addEventListener("change",onActChange);actLoad().then(()=>{if(state.active==="crop"){state.loaded.crop=false;loadCrop();}else if(state.active==="actioned"){renderActioned();}});}
function loadActioned(){actLoad().then(renderActioned);}
function renderActioned(){
  const el=document.getElementById("actionedTable");if(!el){state.loaded.actioned=true;return;}
  const o=actStore();const keys=Object.keys(o);
  const src=window.__actShared?'Shared across all users of the hosted dashboard.':'Stored in this browser only (shared server store not reachable).';
  if(!keys.length){el.innerHTML='<div class="note">No actioned alerts yet. On the Crop risk tab, tick the “Actioned” box on an event to move it here and record what was done. '+src+'</div>';state.loaded.actioned=true;return;}
  keys.sort((a,b)=>String(o[b].updatedAt||o[b].actionedAt||"").localeCompare(String(o[a].updatedAt||o[a].actionedAt||"")));
  const ths='style="border:1px solid var(--line);padding:5px 7px;text-align:left;background:#f3efe9"';
  const tds='style="border:1px solid var(--line);padding:5px 7px;text-align:left;vertical-align:top"';
  const tdc='style="border:1px solid var(--line);padding:5px 7px;text-align:center;vertical-align:top"';
  let h='<div class="srcnote">'+src+'</div><table style="border-collapse:collapse;width:100%;font-size:12px"><thead><tr>'
    +['Done','Crop','Country','Condition','Period','Severity','Peak','Growers','Likely issue','Action taken','Actioned on','Actioned by'].map(x=>'<th '+ths+'>'+x+'</th>').join('')
    +'</tr></thead><tbody>';
  keys.forEach(k=>{const a=o[k];const eid=encodeURIComponent(k);
    h+='<tr><td '+tdc+'><input type="checkbox" class="actchk" data-eid="'+eid+'" checked title="Untick to restore to Crop risk"></td>'
      +'<td '+tds+'>'+escH(a.crop)+'</td><td '+tds+'>'+escH(a.country)+'</td><td '+tds+'>'+escH(a.cond)+'</td><td '+tds+'>'+escH(a.period)+'</td><td '+tds+'>'+escH(a.sev)+'</td><td '+tds+'>'+escH(a.peak)+'</td><td '+tds+'>'+escH(a.growers)+'</td><td '+tds+'>'+escH(a.issue)+'</td>'
      +'<td '+tds+'><textarea class="actnote" data-eid="'+eid+'" rows="2" placeholder="What was done…" style="width:100%;min-width:200px;font:inherit;border:1px solid var(--line);border-radius:4px;padding:3px;box-sizing:border-box">'+escH(a.note)+'</textarea></td>'
      +'<td '+tds+'>'+escH(a.actionedAt)+'</td><td '+tds+'>'+escH(a.actionedBy||a.updatedBy||"—")+'</td></tr>';
  });
  h+='</tbody></table>';el.innerHTML=h;state.loaded.actioned=true;
}
function loadAbout(){
  const M=(typeof META!=="undefined")?META:{};
  const set=(id,v)=>{const el=document.getElementById(id);if(el)el.textContent=v;};
  if(M.obs_dates)set("abObs",fmtDate(M.obs_dates[0])+" to "+fmtDate(M.obs_dates[1]));
  if(M.fc_dates)set("abFc",fmtDate(M.fc_dates[0])+" to "+fmtDate(M.fc_dates[1]));
  set("abGen",M.generated?fmtDate(M.generated):"\u2014");
  set("abLoc",(M.n_loc!=null?M.n_loc:"\u2014"));
  set("abCountries",(M.countries||[]).join(", "));
  state.loaded.about=true;
}
function loadCrop(){
  const crops=activeCrops();
  const multi=!(state.crops&&state.crops.length===1);
  const rng=cropHistRange();
  updateCropScope();
  let planAgg=[],fcPeriods=[],hsPeriods=[];
  let idsUnion=new Set(),fcAdverseSet=new Set();
  let fcHeat=0,fcRain=0,baseSum=0,baseN=0,adjSum=0,inAdv=0,inDays=0,goldSum=0,goldN=0,lostSum=0,worstSev=0;
  const histMonthly={};let anyGrown=false;
  crops.forEach(crop=>{
    const planRows=planRowsFor(crop);if(!planRows.length)return;anyGrown=true;
    const rules=CROP_RULES[cropTypeOf(crop)];
    const months=seasonMonthsFrom(planRows);
    const ids=cropIdsFor(crop);ids.forEach(x=>idsUnion.add(x));
    const fcRows=FC.filter(r=>ids.has(r[F.li]));
    const fcEv=cropEvents(fcRows,true,rules,months);
    fcEv.forEach(e=>fcAdverseSet.add(e.date+"|"+e.loc+"|"+e.grower+"|"+crop));
    fcHeat+=fcEv.filter(e=>e.type==="Heat").length;fcRain+=fcEv.filter(e=>e.type==="Heavy rain").length;
    const hsRows=OBS.filter(r=>ids.has(r[O.li])&&r[O.d]>=rng.start&&r[O.d]<=rng.end);
    const hsEv=cropEvents(hsRows,false,rules,months);
    const useM=months&&months.size&&months.size<12;
    const inSeasonDays=useM?hsRows.filter(r=>months.has(+r[O.d].slice(5,7))).length:hsRows.length;
    const hsAdverse=new Set(hsEv.map(e=>e.date+"|"+e.loc+"|"+e.grower)).size;
    inAdv+=hsAdverse;inDays+=inSeasonDays;
    const cropRate=inSeasonDays?hsAdverse/inSeasonDays:0;
    const allEv=hsEv.concat(fcEv);planRows.forEach(r=>{const evsC=allEv.filter(e=>e.country===r.c);const golden=goldFor(r.c,r.crop);const tier=cropTierOf(r.crop);const worstAll=evsC.length?Math.max.apply(null,evsC.map(e=>advSevRank(e.type,e.value,e.thr,e.thr2))):0;(function(){var cs=[["Heat","Heat >"+rules.heat+"\u00b0C"+(rules.heat2!=null?" (damage >"+rules.heat2+"\u00b0C)":""),rules.heatIssue],["Frost","Cold <"+rules.frost+"\u00b0C",rules.frostIssue],["Heavy rain","Heavy rain >"+rules.rain+"mm",rules.rainIssue]];if(rules.wind!=null)cs.push(["High wind","High wind >"+rules.wind+" km/h",rules.windIssue]);return cs;})().forEach(function(C){const ev=evsC.filter(e=>e.type===C[0]);const sr=ev.length?Math.max.apply(null,ev.map(e=>advSevRank(e.type,e.value,e.thr,e.thr2))):0;const fr=lossFrac(tier,sr);const lost=(golden!=null)?+(golden*fr).toFixed(2):null;planAgg.push({c:r.c,g:r.g,crop:r.crop,season:r.season,gest:r.gest,cond:C[1],iss:C[2],golden:golden,sev:sr,lost:lost,pct:Math.round(fr*100)});});baseSum+=r.y;baseN++;if(golden!=null){goldSum+=golden;goldN++;lostSum+=(golden*lossFrac(tier,worstAll));}if(worstAll>worstSev)worstSev=worstAll;});
    hsEv.forEach(e=>{const k=e.date.slice(0,7);if(!histMonthly[k])histMonthly[k]={Heat:0,Rain:0,Frost:0,Wind:0};if(e.type==="Heat")histMonthly[k].Heat++;else if(e.type==="Heavy rain")histMonthly[k].Rain++;else if(e.type==="Frost")histMonthly[k].Frost++;else if(e.type==="High wind")histMonthly[k].Wind++;});
    const lastObs=META.obs_dates[1];const curMonth=new Date().getUTCMonth()+1;const inSeasonNow=new Set(planRows.filter(r=>r.gmonths.indexOf(curMonth)>=0).map(r=>r.c));const _rc=new Date();_rc.setDate(_rc.getDate()-30);const recentCutoff=_rc.toISOString().slice(0,10);
    summarise(hsEv.concat(fcEv)).forEach(p=>{p.crop=crop;const current=(p.end>lastObs||p.end>=recentCutoff)&&inSeasonNow.has(p.country);if(current)fcPeriods.push(p);else hsPeriods.push(p);});
  });
  const baseline=baseN?Math.round(baseSum/baseN):null;
  const adjusted=baseN?Math.round(adjSum/baseN):null;
  const goldAvg=goldN?goldSum/goldN:null;const lostAvg=goldN?lostSum/goldN:null;
  const histRate=inDays?inAdv/inDays:0;
  const fcSiteDaysN=FC.filter(r=>idsUnion.has(r[F.li])).length;
  const countries=[...new Set([...idsUnion].map(i=>locCountry[i]))].sort();
  if(!anyGrown){
    document.getElementById("cropPlanTable").innerHTML='<div class="note">None of the selected crops are grown for '+cropScopeLabel()+'.</div>';
    document.getElementById("cropPlanNote").innerHTML="";document.getElementById("crop-kpis").innerHTML='';
    document.getElementById("cropFcTable").innerHTML='<div class="note">Select a crop grown in this selection to see weather cross-checks.</div>';
    document.getElementById("cropHsTable").innerHTML="";document.getElementById("cropFcNote").innerHTML="";document.getElementById("cropModelNote").innerHTML="";
    if(charts.cropHist)charts.cropHist.destroy();state.loaded.crop=true;return;
  }
  const showCrop=multi||new Set(planAgg.map(r=>r.crop)).size>1;
  let planCols,planData;
  if(showCrop){
    planCols=[{name:"Crop"},{name:"Grower"},{name:"Country"},{name:"Season"},{name:"Gest (wks)"},{name:"Adverse condition"},{name:"Likely issue"},{name:"Golden t/ha"},{name:"Severity"},{name:"Lost t/ha"}];
    planData=planAgg.map(r=>[r.crop,r.g,r.c,r.season,r.gest,r.cond,r.iss,(r.golden==null?"—":r.golden),SEVLBL[r.sev],(r.golden==null?"—":(r.sev===0?"—":num(r.lost,1)+" - "+r.pct+"%"))]);
  }else{
    planCols=[{name:"Grower"},{name:"Country"},{name:"Season"},{name:"Gest (wks)"},{name:"Adverse condition"},{name:"Likely issue"},{name:"Golden t/ha"},{name:"Severity"},{name:"Lost t/ha"}];
    planData=planAgg.map(r=>[r.g,r.c,r.season,r.gest,r.cond,r.iss,(r.golden==null?"—":r.golden),SEVLBL[r.sev],(r.golden==null?"—":(r.sev===0?"—":num(r.lost,1)+" - "+r.pct+"%"))]);
  }
  grid("cropPlanTable",planCols,planData,{search:showCrop,wrap:true});
  document.getElementById("cropPlanNote").innerHTML="Gest (wks) = weeks from seedling to harvest. Golden t/ha = best-practice yield under ideal conditions. Severity = worst adverse event, rated by how far it passes the crop’s heat/frost/heavy-rain threshold. Lost t/ha = golden × the loss at the worst adverse severity for the crop in that country, over the combined timeline (observations to "+fmtDate(META.obs_dates[1])+", forecast thereafter for "+cropScopeLabel()+").";
  document.getElementById("crop-kpis").innerHTML=kpi("",(goldAvg!=null?num(goldAvg,1)+" t/ha":"—"),"Golden yield (avg)")
    +kpi("earth",(lostAvg!=null?num(lostAvg,2)+" t/ha":"—"),"Potential lost yield (avg)")
    +kpi(worstSev>=2?"orange":"",SEVLBL[worstSev],"Worst adverse severity")
    +kpi("orange",num(histRate*100,1)+" %","Historical adverse-day rate")
    +kpi("omni",fcAdverseSet.size+" / "+fcSiteDaysN,"Forecast adverse site-days")
    +kpi("orange",fcHeat,"Forecast heat days")
    +kpi("",fcRain,"Forecast heavy-rain days");
  const sumCols=[{name:"Period"},{name:"Crop"},{name:"Country"},{name:"Growers"},{name:"Condition"},{name:"Peak"},{name:"Severity"},{name:"Est. loss t/ha"},{name:"Days"},{name:"Likely issue"}];
  function periodRow(e){const ag=activeGrowers();const cg=[...new Set(CROP.rows.filter(x=>x.crop===e.crop&&x.c===e.country&&(!ag||ag.indexOf(x.g)>=0)).map(x=>x.g))];const _g=goldFor(e.country,e.crop);const _sup=fcPeriods.concat(hsPeriods).some(q=>q!==e&&q.crop===e.crop&&q.country===e.country&&q.start>e.start&&dayGap(e.start,q.start)<=14);const _fr=lossFrac(cropTierOf(e.crop),advSevRank(e.type,e.peak,e.thr,e.thr2));const _loss=(_g==null)?"—":(_sup?"—":(num(_g*_fr,1)+" - "+Math.round(_fr*100)+"%"));return [fmtDate(e.start)+(e.start===e.end?"":" – "+fmtDate(e.end))+(e.end<=META.obs_dates[1]?" · ended":""),e.crop,e.country,cg.join(", "),condLabel(e),num(e.peak,1)+" "+e.unit,SEVLBL[advSevRank(e.type,e.peak,e.thr,e.thr2)],_loss,e.days,e.issue];}
  fcPeriods.sort((a,b)=>a.start>b.start?-1:(a.start<b.start?1:(a.country<b.country?-1:1)));
  const fcCols=[{name:"Actioned"}].concat(sumCols).concat([{name:"Substitute supply (in-season, no forecast breach)"}]);
  curEventSnap={};
  const fcVisible=fcPeriods.filter(e=>!isActioned(e));
  if(fcVisible.length){grid("cropFcTable",fcCols,fcVisible.map(e=>{const eid=eventId(e);curEventSnap[eid]=eventSnap(e);return [actChkCell(eid)].concat(periodRow(e)).concat([subLabel(e)]);}),{search:false,wrap:true,resizable:true});}
  else{document.getElementById("cropFcTable").innerHTML='<div class="note">No forecast periods breach the adverse thresholds for the selected crops in this selection.</div>';}
  document.getElementById("cropFcNote").innerHTML="Shows only crops in their current growing season for the country. Adverse events combine observed data (to "+fmtDate(META.obs_dates[1])+") with the latest forecast: an event in progress at the last observation that the forecast continues is shown as one event over its full length. The table also includes recent observed events that ended within the last 30 days (marked “ended”), alongside ongoing and forecast events. Severity is rated by exceedance beyond each crop’s heat/frost/heavy-rain/high-wind threshold. Est. loss t/ha = golden yield × the event’s severity loss, with only the latest of events within 14 days carrying the loss. <em>Substitute supply</em> = growers of the same crop in another country, in season with no significant (medium or severe) forecast breach — potential sources to bolster supply.";
  const keys=Object.keys(histMonthly).sort();
  mk("cropHist",{type:"bar",data:{labels:keys.map(k=>{const p=k.split("-");return MON[+p[1]-1]+" "+p[0].slice(2);}),datasets:[
    {label:"Heat",data:keys.map(k=>histMonthly[k].Heat),backgroundColor:PAL[2],stack:"s"},
    {label:"Heavy rain",data:keys.map(k=>histMonthly[k].Rain),backgroundColor:PAL[1],stack:"s"},
    {label:"Frost",data:keys.map(k=>histMonthly[k].Frost),backgroundColor:PAL[4],stack:"s"},
    {label:"High wind",data:keys.map(k=>histMonthly[k].Wind),backgroundColor:PAL[5],stack:"s"}
  ]},options:baseOpts({scales:{x:{stacked:true,ticks:{maxTicksLimit:14,autoSkip:true}},y:{stacked:true,beginAtZero:true,title:{display:true,text:"adverse days"}}}})});
  hsPeriods.sort((a,b)=>a.start<b.start?1:-1);hsPeriods=hsPeriods.slice(0,300);
  if(hsPeriods.length){grid("cropHsTable",sumCols,hsPeriods.map(periodRow));}
  else{document.getElementById("cropHsTable").innerHTML='<div class="note">No historical periods breach the adverse thresholds for the selected crops in this selection/window.</div>';}
  document.getElementById("cropModelNote").innerHTML="Weather cross-check restricted to each crop’s growing season. Each adverse event is rated Mild/Medium/Severe by how far it exceeds the crop’s threshold, and given an estimated loss = golden-case yield × the crop’s loss fraction at that severity — shown per event in the tables, with only the latest of any events within 14 days carrying the loss. The KPIs and plan table show the loss at the worst adverse severity for the crop in that country. Countries in scope: "+(countries.join(", ")||"—")+". Indicative only — excludes agronomy, variety, pests and market factors.";
  state.loaded.crop=true;
}

/* crop plan */
function loadCropPlan(){
  const cc=document.getElementById("cpCountry").value, ck=[...document.getElementById("cpCrop").selectedOptions].map(o=>o.value);
  const rows=CROP.rows.filter(r=>(!cc||r.c===cc)&&(!ck.length||ck.indexOf(r.crop)>=0));
  const golds=rows.map(r=>goldFor(r.c,r.crop)).filter(x=>x!=null);const avgGold=golds.length?golds.reduce((a,b)=>a+b,0)/golds.length:null;
  const sevs=rows.map(r=>{const gg=goldFor(r.c,r.crop);return gg!=null?gg*lossFrac(cropTierOf(r.crop),3):null;}).filter(x=>x!=null);const avgSev=sevs.length?sevs.reduce((a,b)=>a+b,0)/sevs.length:null;
  document.getElementById("cropplan-kpis").innerHTML=kpi("",rows.length,"Crop entries")
    +kpi("",new Set(rows.map(r=>r.g)).size,"Growers")
    +kpi("",new Set(rows.map(r=>r.crop)).size,"Distinct crops")
    +kpi("earth",(avgGold!=null?num(avgGold,1)+" t/ha":"—"),"Avg golden yield")
    +kpi("orange",(avgSev!=null?num(avgSev,2)+" t/ha":"—"),"Avg severe lost");
  grid("cropPlanFull",[{name:"Country",width:"7%"},{name:"Grower",width:"10%"},{name:"Crop",width:"11%"},{name:"Season",width:"7%"},{name:"Gest (wks)",width:"5%"},{name:"Adverse conditions",width:"20%"},{name:"Likely issues",width:"18%"},{name:"Golden t/ha",width:"6%"},{name:"Mild lost",width:"5%"},{name:"Medium lost",width:"5%"},{name:"Severe lost",width:"6%"}],
    rows.map(r=>{const gg=goldFor(r.c,r.crop),ti=cropTierOf(r.crop);const L=n=>gg!=null?+(gg*lossFrac(ti,n)).toFixed(2):null;const f=v=>v==null?"—":v;return [r.c,r.g,r.crop,r.season,r.gest,advThresholds(r.crop)+" \u2014 "+r.adv,r.iss,f(gg),f(L(1)),f(L(2)),f(L(3))];}),{wrap:true});
  state.loaded.cropplan=true;
}
/* crop-risk update button */
function setCropDirty(){const b=document.getElementById("cropUpdateBtn");if(b)b.classList.add("dirty");}
function clearCropDirty(){const b=document.getElementById("cropUpdateBtn");if(b)b.classList.remove("dirty");}
function applyCropFilters(){state.crops=[...document.getElementById("cropSel").selectedOptions].map(o=>o.value);state.cropGrowers=[...document.getElementById("cropGrower").selectedOptions].map(o=>o.value);state.cropCountry=document.getElementById("cropCountry").value;state.cropMonths=+document.getElementById("cropPeriod").value;state.loaded.crop=false;clearCropDirty();loadCrop();}
function syncCropControls(){const cs=document.getElementById("cropSel");[...cs.options].forEach(o=>o.selected=state.crops.indexOf(o.value)>=0);document.getElementById("cropCountry").value=state.cropCountry;refreshCropGrowers();const gs=document.getElementById("cropGrower");[...gs.options].forEach(o=>o.selected=state.cropGrowers.indexOf(o.value)>=0);document.getElementById("cropPeriod").value=String(state.cropMonths);refreshMSel("cropSel");refreshMSel("cropGrower");clearCropDirty();}
/* crops by country */
function loadByCountry(){
  const fCrop=[...((document.getElementById("bcCrop")||{selectedOptions:[]}).selectedOptions)].map(o=>o.value);
  const fCountry=(document.getElementById("bcCountry")||{}).value||"";
  const fGrower=(document.getElementById("bcGrower")||{}).value||"";
  const fSeason=+(((document.getElementById("bcSeason")||{}).value)||0);
  const rowsF=CROP.rows.filter(r=>(!fCrop.length||fCrop.indexOf(r.crop)>=0)&&(!fCountry||r.c===fCountry)&&(!fGrower||r.g===fGrower)&&(!fSeason||r.gmonths.indexOf(fSeason)>=0));
  const countries=[...new Set(rowsF.map(r=>r.c))].sort();
  const crops=[...new Set(rowsF.map(r=>r.crop))].sort();
  const cell={};rowsF.forEach(r=>{const k=r.crop+"|"+r.c;(cell[k]=cell[k]||new Set()).add(r.g);});
  const cols=[{name:"Crop",width:"22%"}].concat(countries.map(c=>({name:c})));
  const data=crops.map(cr=>[cr].concat(countries.map(c=>{const s=cell[cr+"|"+c];return s?[...s].join(", "):"";})));
  if(crops.length&&countries.length){grid("byCountryTable",cols,data,{search:false,wrap:true});}
  else{document.getElementById("byCountryTable").innerHTML='<div class="note">No crops match the current filters.</div>';}
  document.getElementById("bycountry-kpis").innerHTML=kpi("",crops.length,"Distinct crops")
    +kpi("",countries.length,"Countries")
    +kpi("",rowsF.length,"Crop-grower entries")
    +kpi("",new Set(rowsF.map(r=>r.g)).size,"Growers");
  state.loaded.bycountry=true;
}
/* filters / tabs */
function refreshLocationOptions(){
  const lSel=document.getElementById("fLocation");const cur=lSel.value;lSel.innerHTML='<option value="">All locations</option>';
  const country=document.getElementById("fCountry").value;const seen=new Set();
  L.map((x,i)=>i).filter(i=>!country||locCountry[i]===country).sort((a,b)=>locGrower[a].localeCompare(locGrower[b])||locName[a].localeCompare(locName[b]))
    .forEach(i=>{const o=document.createElement("option");o.value=locGrower[i]+"||"+locName[i]+"||"+locCountry[i];o.textContent=locGrower[i]+" — "+locName[i];lSel.appendChild(o);});if([...lSel.options].some(o=>o.value===cur))lSel.value=cur;
}
function setDirty(){const b=document.getElementById("updateBtn"),h=document.getElementById("hint");b.classList.add("dirty");h.classList.add("dirty");h.innerHTML="You have unapplied changes — press <strong>Update</strong> to refresh the view.";}
function clearDirty(){const b=document.getElementById("updateBtn"),h=document.getElementById("hint");b.classList.remove("dirty");h.classList.remove("dirty");h.innerHTML="Select your filters, then press <strong>Update</strong> to apply.";}
function applyFilters(){
  state.country=document.getElementById("fCountry").value;
  state.location=document.getElementById("fLocation").value;
  const ap=document.querySelector("#presets button.active");state.histDays=ap?+ap.dataset.days:90;
  state.histStart=document.getElementById("fStart").value||null;
  state.histEnd=document.getElementById("fEnd").value||null;
  state.loaded={forecast:false,historical:false,crop:false,cropplan:false,bycountry:false,map:false,accuracy:false};clearDirty();reloadActive();updateAskScope();
}
const SVGNS="http://www.w3.org/2000/svg";
let mapView={x:-180,y:-90,w:360,h:180};
let mapSel=new Set();function siteKey(s){return s.g+"|"+s.loc+"|"+s.c;}
function applyMapView(){
  const s=document.getElementById("worldsvg");if(!s)return;
  s.setAttribute("viewBox",mapView.x.toFixed(3)+" "+mapView.y.toFixed(3)+" "+mapView.w.toFixed(3)+" "+mapView.h.toFixed(3));
  (function(){const bx=document.getElementById("mapScale"),bar=document.getElementById("mapScaleBar"),lbl=document.getElementById("mapScaleLbl");if(!bx||!bar||!lbl)return;const pxW=(s.getBoundingClientRect&&s.getBoundingClientRect().width)||s.clientWidth||700;const centerLat=-(mapView.y+mapView.h/2);const kmPerDeg=111.32*Math.cos(centerLat*Math.PI/180);const totalKm=Math.max(1e-4,mapView.w*Math.max(0.05,kmPerDeg));const kmPerPx=totalKm/pxW;let target=kmPerPx*Math.min(150,pxW*0.28);const p=Math.pow(10,Math.floor(Math.log10(target)));const f=target/p;const nice=(f>=5?5:f>=2?2:1)*p;bar.style.width=Math.max(8,nice/kmPerPx).toFixed(1)+"px";lbl.textContent=nice>=1?(nice.toLocaleString("en-GB")+" km"):(Math.round(nice*1000)+" m");})();
  const r=mapView.w*0.0065;document.querySelectorAll("#markers circle:not([data-ring])").forEach(c=>c.setAttribute("r",r.toFixed(3)));document.querySelectorAll("#markers circle[data-ring]").forEach(c=>c.setAttribute("r",(r*1.9).toFixed(3)));
  const showLbl=mapView.w<180;const fs=mapView.w*0.011;const gap=r*1.4;
  const texts=[...document.querySelectorAll("#markers text")];
  texts.forEach(t=>{t.style.display=showLbl?"":"none";t.setAttribute("font-size",fs.toFixed(3));t.setAttribute("stroke-width",(mapView.w*0.004).toFixed(3));t.setAttribute("x",(parseFloat(t.getAttribute("data-lon"))+gap).toFixed(3));t.setAttribute("y",(-parseFloat(t.getAttribute("data-lat"))).toFixed(3));});
  if(showLbl){
    const inView=texts.filter(t=>{const x=+t.getAttribute("x"),y=+t.getAttribute("y");return x>=mapView.x&&x<=mapView.x+mapView.w&&y>=mapView.y&&y<=mapView.y+mapView.h;});
    inView.sort((a,b)=>(+a.getAttribute("y"))-(+b.getAttribute("y")));
    const placed=[];
    const hit=(b)=>placed.some(p=>!(b.x1<=p.x0||b.x0>=p.x1||b.y1<=p.y0||b.y0>=p.y1));
    inView.forEach(t=>{
      const px=parseFloat(t.getAttribute("data-lon")),py=-parseFloat(t.getAttribute("data-lat"));
      const w=t.textContent.length*fs*0.55,hh=fs*1.15;
      const boxFor=(cx,cy,an)=>{let x0,x1;if(an==="start"){x0=cx;x1=cx+w;}else if(an==="end"){x0=cx-w;x1=cx;}else{x0=cx-w/2;x1=cx+w/2;}return {x0:x0,x1:x1,y0:cy-hh/2,y1:cy+hh/2,cx:cx,cy:cy,an:an};};
      const cands=[
        boxFor(px+gap,py,"start"),
        boxFor(px-gap,py,"end"),
        boxFor(px,py-gap-hh*0.6,"middle"),
        boxFor(px,py+gap+hh*0.6,"middle"),
        boxFor(px+gap,py-hh,"start"),
        boxFor(px+gap,py+hh,"start"),
        boxFor(px-gap,py-hh,"end"),
        boxFor(px-gap,py+hh,"end"),
        boxFor(px,py-gap-hh*1.7,"middle"),
        boxFor(px,py+gap+hh*1.7,"middle"),
        boxFor(px-gap,py-hh*2,"end"),
        boxFor(px+gap,py-hh*2,"start")
      ];
      let chosen=null;
      for(let k=0;k<cands.length;k++){if(!hit(cands[k])){chosen=cands[k];break;}}
      if(!chosen){let cy=py,b=boxFor(px+gap,cy,"start"),tries=0;while(tries<40&&hit(b)){cy+=hh;b=boxFor(px+gap,cy,"start");tries++;}chosen=b;}
      t.setAttribute("x",chosen.cx.toFixed(3));
      t.setAttribute("y",chosen.cy.toFixed(3));
      t.setAttribute("text-anchor",chosen.an);
      t.setAttribute("dominant-baseline","central");
      placed.push(chosen);
    });
  }
}
function clampMapView(){mapView.w=Math.min(360,Math.max(0.02,mapView.w));mapView.h=mapView.w/2;if(mapView.x<-180)mapView.x=-180;if(mapView.y<-90)mapView.y=-90;if(mapView.x+mapView.w>180)mapView.x=180-mapView.w;if(mapView.y+mapView.h>90)mapView.y=90-mapView.h;}
function mapZoom(factor,cx,cy){const nw=Math.min(360,Math.max(0.02,mapView.w*factor));const nh=nw/2;if(cx==null){cx=mapView.x+mapView.w/2;cy=mapView.y+mapView.h/2;}mapView.x=cx-(cx-mapView.x)*(nw/mapView.w);mapView.y=cy-(cy-mapView.y)*(nh/mapView.h);mapView.w=nw;mapView.h=nh;clampMapView();applyMapView();}
let boxMode=false;
let mapMultiMode=false;
function setBoxMode(on){boxMode=on;const s=document.getElementById("worldsvg");if(s)s.style.cursor=on?"crosshair":"";const b=document.querySelector('.mapzoom button[data-z="box"]');if(b)b.classList.toggle("active",on);}
function setSelMode(on){mapMultiMode=on;const s=document.getElementById("worldsvg");if(s)s.style.cursor=on?"pointer":"";const b=document.querySelector('.mapzoom button[data-z="sel"]');if(b)b.classList.toggle("active",on);}
function mapInit(){
  const s=document.getElementById("worldsvg");if(!s)return;
  const rb=document.createElementNS(SVGNS,"rect");rb.setAttribute("id","zoomrect");rb.setAttribute("fill","rgba(0,103,71,0.12)");rb.setAttribute("stroke","#006747");rb.setAttribute("stroke-width","1.2");rb.setAttribute("stroke-dasharray","3 2");rb.setAttribute("vector-effect","non-scaling-stroke");rb.style.display="none";rb.style.pointerEvents="none";s.appendChild(rb);
  function toSvg(e){const r=s.getBoundingClientRect();return {x:mapView.x+(e.clientX-r.left)/r.width*mapView.w,y:mapView.y+(e.clientY-r.top)/r.height*mapView.h};}
  document.querySelectorAll(".mapzoom button").forEach(b=>b.addEventListener("click",()=>{const z=b.dataset.z;if(z==="in")mapZoom(0.7);else if(z==="out")mapZoom(1/0.7);else if(z==="box")setBoxMode(!boxMode);else if(z==="sel")setSelMode(!mapMultiMode);else{mapView={x:-180,y:-90,w:360,h:180};setBoxMode(false);applyMapView();}}));
  s.addEventListener("wheel",e=>{e.preventDefault();const r=s.getBoundingClientRect();const px=(e.clientX-r.left)/r.width,py=(e.clientY-r.top)/r.height;mapZoom(e.deltaY<0?0.85:1/0.85,mapView.x+px*mapView.w,mapView.y+py*mapView.h);},{passive:false});
  let drag=null,box=null;
  s.addEventListener("mousedown",e=>{const p=toSvg(e);if(boxMode||e.shiftKey){box={x0:p.x,y0:p.y};rb.style.display="block";rb.setAttribute("x",p.x);rb.setAttribute("y",p.y);rb.setAttribute("width",0);rb.setAttribute("height",0);e.preventDefault();}else{drag={x:e.clientX,y:e.clientY};s.style.cursor="grabbing";}});
  window.addEventListener("mousemove",e=>{
    if(box){const p=toSvg(e);const x=Math.min(box.x0,p.x),y=Math.min(box.y0,p.y),w=Math.abs(p.x-box.x0),hh=Math.abs(p.y-box.y0);rb.setAttribute("x",x);rb.setAttribute("y",y);rb.setAttribute("width",w);rb.setAttribute("height",hh);return;}
    if(drag){const r=s.getBoundingClientRect();mapView.x-=(e.clientX-drag.x)/r.width*mapView.w;mapView.y-=(e.clientY-drag.y)/r.height*mapView.h;drag={x:e.clientX,y:e.clientY};clampMapView();applyMapView();}
  });
  window.addEventListener("mouseup",e=>{
    if(box){const p=toSvg(e);const x0=Math.min(box.x0,p.x),y0=Math.min(box.y0,p.y),w=Math.abs(p.x-box.x0),hh=Math.abs(p.y-box.y0);box=null;rb.style.display="none";
      if(w>0.5&&hh>0.5){let cw=w,ch=hh;if(cw/ch>2)ch=cw/2;else cw=ch*2;const cx=x0+w/2,cy=y0+hh/2;mapView.w=Math.min(360,Math.max(0.02,cw));mapView.h=mapView.w/2;mapView.x=cx-mapView.w/2;mapView.y=cy-mapView.h/2;clampMapView();applyMapView();}
      setBoxMode(false);return;}
    drag=null;if(!boxMode)s.style.cursor="";
  });
  const mv=document.getElementById("mapSelView");if(mv)mv.addEventListener("click",()=>{const g=selectedMapGrowers();if(g.length)openGrowersRisk(g);});
  const mc=document.getElementById("mapSelClear");if(mc)mc.addEventListener("click",()=>{mapSel.clear();updateMapSelBar();loadMap();});
}
function showMapTip(ev,gr,st){const tip=document.getElementById("mapTip"),box=document.querySelector(".mapbox");if(!tip||!box)return;const r=box.getBoundingClientRect();tip.innerHTML="<strong>"+gr.g+"</strong>"+(gr.loc?" · "+gr.loc:"")+" — "+gr.c+"<br>"+(st.status==="none"?'<span style=\'color:#8a7f77\'>Not in the product plan</span>':st.adverse?'<span style=\'color:#c0392b;font-weight:600\'>Forecast adverse conditions</span>':'<span style=\'color:#2e7d32;font-weight:600\'>No adverse forecast</span>')+"<br><span style=\'color:#8a7f77\'>"+st.crops.join(", ")+"</span>";tip.style.display="block";let x=ev.clientX-r.left+12,y=ev.clientY-r.top+12;if(x>r.width-150)x=x-160;if(y>r.height-70)y=y-70;tip.style.left=x+"px";tip.style.top=y+"px";}
function hideMapTip(){const t=document.getElementById("mapTip");if(t)t.style.display="none";}
const siteLi={};function buildSiteLi(){for(let i=0;i<L.length;i++)siteLi[locGrower[i]+"|"+locName[i]+"|"+locCountry[i]]=i;}
function siteAdverse(site){
  const g=siteToGrower(site);
  const srcRows=g?CROP.rows.filter(x=>x.g===g):CROP.rows.filter(x=>x.c===site.c);
  const crops=[...new Set(srcRows.map(x=>x.crop))];
  if(!crops.length)return {status:"none",adverse:false,crops:[],grower:g,advCrops:[],advGrowers:[]};
  const li=siteLi[site.g+"|"+site.loc+"|"+site.c];
  const fcRows=(li==null)?[]:FC.filter(r=>r[F.li]===li);
  const advCrops=[];
  for(const crop of crops){
    const rules=CROP.rules[cropTypeOf(crop)];
    const gm=new Set();srcRows.filter(x=>x.crop===crop).forEach(x=>x.gmonths.forEach(m=>gm.add(m)));
    if(cropEvents(fcRows,true,rules,gm).length)advCrops.push(crop);
  }
  const adverse=advCrops.length>0;
  let advGrowers;
  if(g)advGrowers=adverse?[g]:[];
  else advGrowers=[...new Set(CROP.rows.filter(x=>x.c===site.c&&advCrops.indexOf(x.crop)>=0).map(x=>x.g))].sort();
  return {status:adverse?"red":"green",adverse,crops,grower:g,advCrops,advGrowers};
}
function siteToGrower(site){if(site&&site.cg)return site.cg;var norm=function(s){return (s||"").toLowerCase().replace(/[^a-z0-9]/g,"");};var gs=CROP.growers.filter(function(x){return x.c===site.c;});var sn=norm(site.g);var m=gs.find(function(x){var gn=norm(x.g);return sn.indexOf(gn)>=0||gn.indexOf(sn)>=0;});return m?m.g:"";}
function forecastBreachingCrops(growers,country){
  const rows=(growers&&growers.length)?CROP.rows.filter(x=>growers.indexOf(x.g)>=0):CROP.rows.filter(x=>x.c===country);
  const crops=[...new Set(rows.map(x=>x.crop))];
  const cset=(growers&&growers.length)?new Set(growers.map(growerCountry).filter(Boolean)):new Set([country]);
  const out=[];
  crops.forEach(crop=>{
    const rules=CROP.rules[cropTypeOf(crop)];
    const gm=new Set();rows.filter(x=>x.crop===crop).forEach(x=>x.gmonths.forEach(m=>gm.add(m)));
    const cc=new Set(CROP.rows.filter(x=>x.crop===crop).map(x=>x.c));
    const ids=matchIdsForCountries(cset,cc);
    const fcRows=FC.filter(r=>ids.has(r[F.li]));
    if(cropEvents(fcRows,true,rules,gm).length)out.push(crop);
  });
  return out;
}
function openSiteRisk(site,st){
  st=st||siteAdverse(site);
  const growers=(st.advGrowers&&st.advGrowers.length)?st.advGrowers.slice():(st.grower?[st.grower]:[]);
  state.cropGrowers=growers;
  state.cropCountry=site.c;
  state.crops=(st.advCrops&&st.advCrops.length)?st.advCrops.slice():[];
  state.loaded.crop=false;switchTab("crop");
}
function openGrowersRisk(growers){state.cropGrowers=(growers||[]).slice();state.cropCountry="";state.crops=[];state.loaded.crop=false;switchTab("crop");}
function toggleMapSel(site){const k=siteKey(site);if(mapSel.has(k))mapSel.delete(k);else mapSel.add(k);loadMap();}
function selectedMapGrowers(){const gs=new Set();(CROP.sites||[]).forEach(s=>{if(mapSel.has(siteKey(s))){const g=siteToGrower(s);if(g)gs.add(g);}});return [...gs].sort();}
function updateMapSelBar(){const bar=document.getElementById("mapSelBar");if(!bar)return;const n=mapSel.size;bar.style.display=n?"flex":"none";if(n){const gr=selectedMapGrowers();document.getElementById("mapSelInfo").textContent=n+" site"+(n>1?"s":"")+" selected"+(gr.length?" · growers: "+gr.join(", "):" · (no product grower matched)");const vb=document.getElementById("mapSelView");if(vb)vb.disabled=!gr.length;}}
function loadMap(){
  const mk=document.getElementById("markers");if(!mk){state.loaded.map=true;return;}
  buildSiteLi();
  while(mk.firstChild)mk.removeChild(mk.firstChild);
  let nRed=0,nGreen=0,nNone=0;
  (CROP.sites||[]).forEach(site=>{
    const st=siteAdverse(site);
    const col=st.status==="red"?"#c0392b":st.status==="green"?"#2e7d32":"#9aa0a6";
    if(st.status==="red")nRed++;else if(st.status==="green")nGreen++;else nNone++;
    const c=document.createElementNS(SVGNS,"circle");
    c.setAttribute("cx",site.lon);c.setAttribute("cy",(-site.lat));c.setAttribute("r",(mapView.w*0.0055).toFixed(2));
    c.setAttribute("fill",col);c.setAttribute("stroke","#ffffff");c.setAttribute("stroke-width","1");c.setAttribute("vector-effect","non-scaling-stroke");
    c.addEventListener("mousemove",ev=>showMapTip(ev,site,st));
    c.addEventListener("mouseleave",hideMapTip);
    c.addEventListener("click",ev=>{if(ev.ctrlKey||ev.metaKey||mapMultiMode){ev.stopPropagation();toggleMapSel(site);}else{openSiteRisk(site,st);}});
    if(mapSel.has(siteKey(site))){const rg=document.createElementNS(SVGNS,"circle");rg.setAttribute("data-ring","1");rg.setAttribute("cx",site.lon);rg.setAttribute("cy",(-site.lat));rg.setAttribute("r",(mapView.w*0.012).toFixed(2));rg.setAttribute("fill","none");rg.setAttribute("stroke","#006747");rg.setAttribute("stroke-width","2.5");rg.setAttribute("vector-effect","non-scaling-stroke");rg.style.pointerEvents="none";mk.appendChild(rg);}
    mk.appendChild(c);
    const tx=document.createElementNS(SVGNS,"text");
    tx.setAttribute("data-lon",site.lon);tx.setAttribute("data-lat",site.lat);
    tx.setAttribute("x",site.lon);tx.setAttribute("y",-site.lat);tx.setAttribute("fill","#3a332f");
    tx.setAttribute("paint-order","stroke");tx.setAttribute("stroke","#ffffff");tx.setAttribute("dominant-baseline","central");
    tx.style.pointerEvents="none";tx.style.fontFamily="Calibri,Arial,sans-serif";tx.textContent=site.g+" — "+site.loc+" — "+site.c;
    mk.appendChild(tx);
  });
  document.getElementById("map-kpis").innerHTML=kpi("earth",nRed,"Sites with forecast adverse conditions")+kpi("omni",nGreen,"Sites clear (no adverse forecast)")+kpi("",(CROP.sites||[]).length,"Grower sites mapped")+kpi("green",new Set((CROP.sites||[]).map(s=>s.c)).size,"Countries mapped");
  updateMapSelBar();
  applyMapView();
  state.loaded.map=true;
}
function reloadActive(){if(state.active==="forecast")loadForecast();else if(state.active==="historical")loadHistorical();else if(state.active==="crop")loadCrop();else if(state.active==="cropplan")loadCropPlan();else if(state.active==="bycountry")loadByCountry();else if(state.active==="map")loadMap();else if(state.active==="accuracy")loadAccuracy();}
/* FORECAST ACCURACY TAB */
function accCompute(inScope, leadLo, leadHi){
  const V=ACC.vars; const sa={},so={}; let n=0;
  V.forEach(v=>{sa[v]=0;so[v]=0;});
  ACC.rows.forEach(r=>{
    const li=r[0],lead=r[1];
    if(lead<leadLo||lead>leadHi)return;
    if(inScope&&!inScope(li))return;
    n+=r[2];
    for(let k=0;k<V.length;k++){so[V[k]]+=r[3+k*2];sa[V[k]]+=r[4+k*2];}
  });
  const out={n:n,vars:{}};
  V.forEach(v=>{
    let acc=null;const mae=n?sa[v]/n:null;
    if(so[v]>0)acc=Math.max(0,Math.min(100,100*(1-sa[v]/so[v])));
    out.vars[v]={acc:acc,mae:mae};
  });
  return out;
}
function accTotal(vars,includeRain){
  const list=ACC.vars.filter(v=>includeRain||v!=="rn");
  const a=list.map(v=>vars[v].acc).filter(x=>x!=null);
  return a.length?a.reduce((p,c)=>p+c,0)/a.length:null;
}
function accScopeFilter(){
  const c=state.country||"", l=state.location||"";
  if(!c&&!l)return null;
  return function(li){const L=ACC.locs[li];return (!c||L[2]===c)&&(!l||(L[0]+"||"+L[1]+"||"+L[2])===l);};
}
function loadAccuracy(){
  const VL={me:"Mean temp",mn:"Min temp",mx:"Max temp",rn:"Rainfall",wd:"Wind speed",gu:"Wind gust"};
  const lo=+((document.getElementById("accLeadLo")||{}).value||0);
  let hiv=(document.getElementById("accLeadHi")||{}).value; const hi=(hiv===undefined||hiv==="")?ACC.leadMax:+hiv;
  const includeRain=(document.getElementById("accRain")||{checked:true}).checked;
  const filt=accScopeFilter();
  const cur=accCompute(filt,lo,hi);
  const scope=scopeLabel();
  document.getElementById("accScope").innerHTML="Forecast accuracy over "+fmtDate(ACC.window[0])+" – "+fmtDate(ACC.window[1])+" · "+scope+" · lead times "+lo+"–"+hi+" days · "+intf(cur.n)+" matched forecast/observation days.";
  const tot=accTotal(cur.vars,includeRain);
  let h=kpi("",(tot==null?"—":num(tot,1)+" %"),"Overall accuracy"+(includeRain?"":" (excl. rain)"));
  ACC.vars.forEach(v=>{const o=cur.vars[v];const cls=(v==="rn")?"orange":"";h+=kpi(cls,(o.acc==null?"—":num(o.acc,1)+" %"),VL[v]+" accuracy");});
  h+=kpi("earth",intf(cur.n),"Matched days");
  document.getElementById("acc-kpis").innerHTML=h;
  const leads=[];for(let d=ACC.leadMin;d<=ACC.leadMax;d++)leads.push(d);
  const overall=[],tmp=[],wind=[],rain=[];
  leads.forEach(d=>{const cc=accCompute(filt,d,d);overall.push(accTotal(cc.vars,includeRain));tmp.push(cc.vars.mx.acc);wind.push(cc.vars.wd.acc);rain.push(cc.vars.rn.acc);});
  const GRID={color:"rgba(0,0,0,.06)"};
  mk("accLeadChart",{type:"line",data:{labels:leads,datasets:[
    {label:"Overall",data:overall,borderColor:"#006747",backgroundColor:"#006747",tension:.3,spanGaps:true},
    {label:"Max temp",data:tmp,borderColor:"#4C7422",tension:.3,spanGaps:true},
    {label:"Wind speed",data:wind,borderColor:"#F29000",tension:.3,spanGaps:true},
    {label:"Rainfall",data:rain,borderColor:"#DB8143",borderDash:[5,4],tension:.3,spanGaps:true}
  ]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{min:0,max:100,title:{display:true,text:"Accuracy %"},grid:GRID},x:{title:{display:true,text:"Forecast lead time (days ahead)"},grid:GRID}},plugins:{legend:{position:"bottom"}}}});
  const vlabels=ACC.vars.map(v=>VL[v]);const vdata=ACC.vars.map(v=>cur.vars[v].acc);
  mk("accVarChart",{type:"bar",data:{labels:vlabels,datasets:[{label:"Accuracy %",data:vdata,backgroundColor:["#006747","#4C7422","#7E9B53","#DB8143","#F29000","#CCB879"]}]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{min:0,max:100,title:{display:true,text:"Accuracy %"},grid:GRID}},plugins:{legend:{display:false}}}});
  function tableFor(dim){
    const groups={};
    ACC.locs.forEach((L,li)=>{
      if(filt&&!filt(li))return;
      const key=dim==="country"?L[2]:dim==="grower"?L[0]:(L[1]+" — "+L[0]);
      (groups[key]=groups[key]||[]).push(li);
    });
    let keys=Object.keys(groups);
    if(dim==="loc"){keys.sort((a,b)=>{const ca=ACC.locs[groups[a][0]][2],cb=ACC.locs[groups[b][0]][2];return ca===cb?a.localeCompare(b):ca.localeCompare(cb);});}else{keys.sort();}
    return keys.map(key=>{
      const set=new Set(groups[key]);
      const cc=accCompute(function(li){return set.has(li);},lo,hi);
      const t=accTotal(cc.vars,includeRain);
      const f=v=>cc.vars[v].acc==null?"—":num(cc.vars[v].acc,1);
      const row=[key,intf(cc.n),f("mn"),f("mx"),f("rn"),f("wd"),f("gu"),(t==null?"—":num(t,1))];
      if(dim==="loc")row.unshift(ACC.locs[groups[key][0]][2]);
      return row;
    });
  }
  const base=[{name:"Days"},{name:"Min °C %"},{name:"Max °C %"},{name:"Rain %"},{name:"Wind %"},{name:"Gust %"},{name:"Overall %"}];
  grid("accCountryTbl",[{name:"Country"}].concat(base),tableFor("country"),{search:false});
  grid("accGrowerTbl",[{name:"Grower"}].concat(base),tableFor("grower"),{search:false});
  grid("accLocTbl",[{name:"Country"},{name:"Grower site"}].concat(base),tableFor("loc"),{search:true});
  document.getElementById("accNote").innerHTML="Method: each forecast is matched to the observation for the same grower, location, country and date, restricted to dates present in the forecast table ("+fmtDate(ACC.window[0])+" – "+fmtDate(ACC.window[1])+"). Same-day (lead 0) forecasts and mean temperature are excluded from this report. For minimum and maximum temperature, wind speed and gust, accuracy = 100 × (1 − Σ|forecast − observed| ÷ Σ observed), clamped to 0–100%. Rainfall is scored per day: a day forecast dry that stays dry counts as 100% accurate, and a day with rain (forecast or observed) scores 100 × (1 − |forecast − observed| ÷ the larger of the two), so both misses and false alarms are penalised; the rainfall figure is the mean of those daily scores. Overall = mean of the available variable accuracies"+(includeRain?"":", excluding rainfall")+". Wind accuracy is inherently more volatile because its observed baseline is smaller. Figures are a static snapshot computed from historical forecasts as at "+fmtDate(ACC.generated)+" ("+intf(ACC.npairs)+" matched days) and are not recalculated by the live Update button, which fetches latest-flag forecasts only.";
  state.loaded.accuracy=true;
}

function alSevCol(s){return {Critical:"#DB8143",High:"#F29000",Medium:"#CCB879",Low:"#7E9B53"}[s]||"#8a7f77";}
function alSevRank(s){return {Critical:4,High:3,Medium:2,Low:1}[s]||0;}
function loadAlerts(){
  const A=(typeof ALERTS!=="undefined"&&ALERTS)?ALERTS:{rows:[],rules:[],dates:null,n:0};
  const rows=(A.rows||[]).slice();
  const kEl=document.getElementById("al-kpis");
  const total=rows.length;
  const open=rows.filter(r=>r.status==="Open").length;
  const high=rows.filter(r=>alSevRank(r.severity)>=3).length;
  const sites=new Set(rows.map(r=>r.farm+"|"+r.loc+"|"+r.country)).size;
  const dts=rows.map(r=>r.obsDate).sort();
  const latest=dts.length?dts[dts.length-1]:null;
  const ref=(typeof META!=="undefined"&&META.obs_dates)?META.obs_dates[1]:latest;
  let gap="—";
  if(latest&&ref)gap=intf(Math.round((Date.parse(ref)-Date.parse(latest))/86400000));
  if(kEl)kEl.innerHTML=kpi("",intf(total),"Alerts in register")
    +kpi("orange",intf(open),"Still open")
    +kpi("earth",intf(high),"High or critical")
    +kpi("omni",intf(sites),"Sites with alerts")
    +kpi("",latest?fmtDate(latest):"—","Most recent alert")
    +kpi("orange",gap,"Days from latest alert to latest observation");

  /* alerts per day, stacked by severity */
  const days=[...new Set(rows.map(r=>r.obsDate))].sort();
  const sevs=[...new Set(rows.map(r=>r.severity))].sort((a,b)=>alSevRank(b)-alSevRank(a));
  mk("alDay",{type:"bar",data:{labels:days.map(fmtDate),datasets:sevs.map(s=>({
      label:s,backgroundColor:alSevCol(s),
      data:days.map(d=>rows.filter(r=>r.obsDate===d&&r.severity===s).length)}))},
    options:baseOpts({scales:{x:{stacked:true,grid:{display:false}},
      y:{stacked:true,beginAtZero:true,ticks:{precision:0},title:{display:true,text:"Alerts"}}}})});

  /* alerts by rule */
  const rl=(A.rules||[]).map(r=>({lab:r.param+" "+r.threshold,n:rows.filter(x=>x.ruleId===r.id).length,sev:r.severity}));
  rl.sort((a,b)=>b.n-a.n);
  mk("alRule",{type:"bar",data:{labels:rl.map(r=>r.lab),datasets:[{label:"Alerts",
      backgroundColor:rl.map(r=>alSevCol(r.sev)),data:rl.map(r=>r.n)}]},
    options:baseOpts({indexAxis:"y",plugins:{legend:{display:false}},
      scales:{x:{beginAtZero:true,ticks:{precision:0},title:{display:true,text:"Alerts"}},y:{grid:{display:false}}}})});

  /* register */
  rows.sort((a,b)=>b.obsDate.localeCompare(a.obsDate)||alSevRank(b.severity)-alSevRank(a.severity)||String(a.id).localeCompare(String(b.id)));
  grid("alTable",[
    {name:"Alert date"},{name:"Grower site"},{name:"Country"},{name:"Parameter"},
    {name:"Measured"},{name:"Threshold"},{name:"Severity"},{name:"Status"},
    {name:"Assigned to"},{name:"Action taken"},{name:"Emailed"},{name:"Alert ID"}],
    rows.map(r=>[fmtDate(r.alertDate),r.farm+" \u2013 "+r.loc,r.country,r.param,
      (r.value==null?"\u2014":num(r.value,1)+" "+r.unit),r.threshold,r.severity,r.status,
      r.assigned||"\u2014",(r.action&&String(r.action).trim())?r.action:"\u2014",
      r.emailSent?"Yes":"No",r.id]),
    {wrap:true,limit:15});

  /* rules table */
  grid("alRules",[{name:"Rule"},{name:"Parameter"},{name:"Threshold"},{name:"Severity"},{name:"Alerts"},{name:"Likely impact"}],
    (A.rules||[]).map(r=>[r.id,r.param,r.threshold,r.severity,rows.filter(x=>x.ruleId===r.id).length,r.impact]),
    {wrap:true,search:false,limit:10});

  const nEl=document.getElementById("alNote");
  if(nEl){
    const cov=A.dates?(fmtDate(A.dates[0])+" – "+fmtDate(A.dates[1])):"—";
    const nSites=(typeof META!=="undefined"&&META.n_loc)?META.n_loc:"—";
    nEl.innerHTML="Source: the <em>Weather Alerts</em> list in the Power BI semantic model “Flamingo Weather”, captured "
      +fmtDate(A.generated)+". The register covers <strong>"+cov+"</strong> and <strong>"+intf(sites)
      +" of "+nSites+" sites</strong>, so it is a pilot rather than full coverage — no alerts have been written since "
      +(latest?fmtDate(latest):"—")+", while observations now run to "+(ref?fmtDate(ref):"—")
      +". Severity is fixed per rule rather than scaled by how far the threshold was exceeded, so every "
      +"temperature breach reads the same regardless of size. Recipient email addresses held on the list are "
      +"deliberately not embedded in this file. Like the Accuracy tab, this register is a static snapshot and is "
      +"not refreshed by the “Update data” button, which fetches observations and forecasts only.";
  }
  state.loaded.alerts=true;
}

function switchTab(t){
  state.active=t;
  document.querySelectorAll(".tab").forEach(b=>b.classList.toggle("active",b.dataset.tab===t));
  document.querySelectorAll(".panel").forEach(p=>p.classList.remove("active"));
  document.getElementById("panel-"+t).classList.add("active");
  const sd=(t==="historical");
  document.getElementById("dateCtl").style.display=sd?"flex":"none";
  document.getElementById("dateFrom").style.display=sd?"flex":"none";
  document.getElementById("dateTo").style.display=sd?"flex":"none";
  const ownCrop=(t==="crop"),ownPlan=(t==="cropplan"),noTop=(ownCrop||ownPlan||t==="bycountry"||t==="map"||t==="refs"||t==="actioned"||t==="about"||t==="alerts");
  document.getElementById("topControls").style.display=noTop?"none":"flex";
  document.getElementById("hint").style.display=noTop?"none":"block";
  document.getElementById("cropControls").style.display=ownCrop?"flex":"none";
  document.getElementById("cpControls").style.display=ownPlan?"flex":"none";
  document.getElementById("bcControls").style.display=(t==="bycountry")?"flex":"none";
  if(ownCrop)syncCropControls();
  if(t==="forecast"&&!state.loaded.forecast)loadForecast();
  if(t==="historical"&&!state.loaded.historical)loadHistorical();
  if(t==="crop"&&!state.loaded.crop)loadCrop();
  if(t==="cropplan"&&!state.loaded.cropplan)loadCropPlan();
  if(t==="bycountry"&&!state.loaded.bycountry)loadByCountry();
  if(t==="map"&&!state.loaded.map)loadMap();
  if(t==="actioned"&&!state.loaded.actioned)loadActioned();
  if(t==="about"&&!state.loaded.about)loadAbout();
  if(t==="refs"&&!state.loaded.refs)loadRefs();
  if(t==="accuracy"&&!state.loaded.accuracy)loadAccuracy();
  if(t==="alerts"&&!state.loaded.alerts)loadAlerts();
}

function updateHeaderMeta(){
  document.getElementById("asOf").textContent="Snapshot "+fmtDate(META.generated);
  document.getElementById("foot").innerHTML="Source: Power BI semantic model &ldquo;Flamingo Weather&rdquo; &middot; Snapshot as at "+fmtDate(META.generated)+" &middot; "+intf(META.n_obs)+" observations, "+intf(META.n_fc)+" forecast rows, "+META.n_loc+" sites &middot; Confidential &ndash; internal use only.";
}

/* REFRESH FROM MCP */
const OT="'Weather Weather_observations'", FT="'Weather Weather_Forecast'";
const Q_OBS="EVALUATE SELECTCOLUMNS("+OT+',"g",'+OT+'[Grower_name],"loc",'+OT+'[Location],"c",'+OT+'[Country],"d",'+OT+'[Observation_Date],"mn",'+OT+'[Min_temp],"mx",'+OT+'[Max_temp],"me",'+OT+'[Mean_temp],"rn",'+OT+'[Rain_volume_mm],"su",'+OT+'[Sunshine_hours],"dl",'+OT+'[Daylight_hours],"wd",'+OT+'[Wind_speed],"gu",'+OT+'[Wind_speed_Gust],"rh",'+OT+'[Rain_hours])';
const Q_FC="EVALUATE SELECTCOLUMNS(FILTER("+FT+","+FT+'[LatestFlag]=1),"g",'+FT+'[Grower_name],"loc",'+FT+'[Location],"c",'+FT+'[Country],"d",'+FT+'[Forecast_date],"mn",'+FT+'[Min_temp],"mx",'+FT+'[Max_temp],"me",'+FT+'[Mean_temp],"rn",'+FT+'[Rain_volume_mm],"rp",'+FT+'[Rain_probability_pc],"su",'+FT+'[Sun_hours],"wd",'+FT+'[Wind_speed],"gu",'+FT+'[Wind_gust],"pw",'+FT+'[Prevailing_Wind],"rh",'+FT+'[Rain_hours])';
async function daxPull(q){
  if(!(window.cowork&&typeof window.cowork.callMcpTool==="function"))throw new Error("connector bridge not available in this view");
  const r=await window.cowork.callMcpTool(TOOL,{workspace_id:WS,dataset_id:DS,dax:q,response_format:"json",max_rows:100000});
  if(r&&r.isError)throw new Error((r.content&&r.content[0]&&r.content[0].text)||"query failed");
  let p=r.structuredContent;if(!p&&r.content&&r.content[0]&&r.content[0].text)p=JSON.parse(r.content[0].text);
  if(!p||!p.rows)throw new Error("no rows returned");
  return p.rows.map(row=>{const o={};for(const k in row){const m=k.match(/\[([^\]]+)\]\s*$/);o[m?m[1]:k]=row[k];}return o;});
}
async function refreshData(){
  const btn=document.getElementById("refreshBtn");const msg=document.getElementById("msg");const orig=btn.innerHTML;
  msg.style.display="none";btn.disabled=true;btn.innerHTML="Refreshing…";
  try{
    const [obs,fc]=await Promise.all([daxPull(Q_OBS),daxPull(Q_FC)]);
    const idx=new Map(),locs=[];const gid=(g,l,c)=>{const k=g+"||"+l+"||"+c;if(!idx.has(k)){idx.set(k,locs.length);locs.push([g,l,c]);}return idx.get(k);};
    const dp=v=>String(v||"").slice(0,10);
    const EXCL=(gg,ll)=>/Bigot/i.test(gg||"")||/Baker Towers/i.test(ll||"");
    OBS=obs.filter(r=>!EXCL(r.g,r.loc)&&dp(r.d)>="2024-01-01").map(r=>[gid(r.g,r.loc,r.c),dp(r.d),round(r.mn),round(r.mx),round(r.me),round(r.rn),(r.su==null?null:round(r.su/3600)),round(r.wd),round(r.gu)]);
    FC=fc.filter(r=>!EXCL(r.g,r.loc)).map(r=>[gid(r.g,r.loc,r.c),dp(r.d),round(r.mn),round(r.mx),round(r.me),round(r.rn),round(r.rp,0),(r.su==null?null:round(r.su/3600)),round(r.wd),round(r.gu),r.pw,round(r.rh)]);
    L=locs;
    META.obs_dates=[OBS.reduce((a,b)=>b[1]<a?b[1]:a,"9999-99-99"),OBS.reduce((a,b)=>b[1]>a?b[1]:a,"0000")];
    META.fc_dates=[FC.reduce((a,b)=>b[1]<a?b[1]:a,"9999-99-99"),FC.reduce((a,b)=>b[1]>a?b[1]:a,"0000")];
    META.n_obs=OBS.length;META.n_fc=FC.length;META.n_loc=L.length;META.generated=new Date().toISOString().slice(0,10);
    META.countries=[...new Set(L.map(x=>x[2]))].sort();
    rebuildLookups();
    const cSel=document.getElementById("fCountry");cSel.innerHTML='<option value="">All countries</option>';
    META.countries.forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;cSel.appendChild(o);});
    refreshLocationOptions();
    const ccS=document.getElementById("cropCountry");ccS.innerHTML='<option value="">All countries</option>';[...new Set(CROP.rows.map(r=>r.c))].sort().forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;ccS.appendChild(o);});state.cropCountry="";
    state.country="";state.location="";state.loaded={forecast:false,historical:false,crop:false,cropplan:false,bycountry:false,map:false,accuracy:false};
    updateHeaderMeta();clearDirty();reloadActive();updateAskScope();
    msg.style.background="#e7f0ea";msg.style.borderColor="#9fc4b2";msg.textContent="Data refreshed from Power BI on "+fmtDate(META.generated)+" — "+intf(META.n_obs)+" observations, "+intf(META.n_fc)+" forecast rows.";msg.style.display="block";
    btn.innerHTML="&#10003; Refreshed";
  }catch(e){
    msg.style.background="";msg.style.borderColor="";
    msg.innerHTML="<strong>Live refresh unavailable in this view.</strong> This is the offline snapshot from "+fmtDate(META.generated)+". To pull fresh data from Power BI, open the live <strong>Flamingo Weather Intelligence</strong> dashboard in the Claude Cowork sidebar and press <strong>Update data</strong> there — the connector is only reachable from inside Cowork.";
    msg.style.display="block";btn.innerHTML=orig;
  }finally{btn.disabled=false;setTimeout(()=>{if(btn.innerHTML.indexOf("Refreshed")>=0)btn.innerHTML=orig;},4000);}
}

function enhanceMSelect(id,placeholder){
  const sel=document.getElementById(id);if(!sel||sel.dataset.msel)return;sel.dataset.msel="1";sel.style.display="none";
  const wrap=document.createElement("div");wrap.className="msel";sel.parentNode.insertBefore(wrap,sel.nextSibling);
  const btn=document.createElement("button");btn.type="button";btn.className="msel-btn";const lbl=document.createElement("span");lbl.className="msel-lbl";btn.appendChild(lbl);
  const panel=document.createElement("div");panel.className="msel-panel";
  wrap.appendChild(btn);wrap.appendChild(panel);
  function summary(){const s=[...sel.selectedOptions];lbl.textContent=(s.length===0)?placeholder:(s.length===1?s[0].textContent:s.length+" selected");}
  function build(){panel.innerHTML="";const selNow=[...sel.selectedOptions];
    if(selNow.length){const chips=document.createElement("div");chips.className="msel-chips";selNow.forEach(o=>{const ch=document.createElement("span");ch.className="msel-chip";const tx=document.createElement("b");tx.textContent=o.textContent;const x=document.createElement("span");x.className="x";x.textContent="\u00d7";x.title="Remove";x.onclick=(e)=>{e.stopPropagation();o.selected=false;sel.dispatchEvent(new Event("change"));build();};ch.appendChild(tx);ch.appendChild(x);chips.appendChild(ch);});panel.appendChild(chips);}
    const tools=document.createElement("div");tools.className="msel-tools";const a=document.createElement("a");a.textContent="Select all";const c=document.createElement("a");c.textContent="Clear";tools.appendChild(a);tools.appendChild(c);panel.appendChild(tools);
    a.onclick=()=>{[...sel.options].forEach(o=>o.selected=true);sel.dispatchEvent(new Event("change"));build();};
    c.onclick=()=>{[...sel.options].forEach(o=>o.selected=false);sel.dispatchEvent(new Event("change"));build();};
    [...sel.options].forEach(o=>{const row=document.createElement("label");row.className="msel-opt";const cb=document.createElement("input");cb.type="checkbox";cb.checked=o.selected;cb.onchange=()=>{o.selected=cb.checked;sel.dispatchEvent(new Event("change"));summary();};row.appendChild(cb);const t=document.createElement("span");t.textContent=o.textContent;row.appendChild(t);panel.appendChild(row);});
    summary();}
  panel.addEventListener("click",e=>e.stopPropagation());
  btn.addEventListener("click",e=>{e.stopPropagation();const open=panel.classList.contains("open");closeAllMSel();if(!open){build();panel.classList.add("open");btn.classList.add("open");}});
  sel._msel={refresh:summary};summary();
}
function closeAllMSel(){document.querySelectorAll(".msel-panel.open").forEach(p=>p.classList.remove("open"));document.querySelectorAll(".msel-btn.open").forEach(b=>b.classList.remove("open"));}
function refreshMSel(id){const s=document.getElementById(id);if(s&&s._msel)s._msel.refresh();}
document.addEventListener("click",closeAllMSel);
const REFS=[
 ["Asparagus",[["RHS — Asparagus","https://www.rhs.org.uk/vegetables/asparagus/grow-your-own"],["Old Farmer's Almanac — Asparagus","https://www.almanac.com/plant/asparagus"],["Oregon State University Extension","https://extension.oregonstate.edu/gardening/techniques/growing-asparagus-your-garden"]]],
 ["Green beans (dwarf, fine, extra fine)",[["RHS — French beans","https://www.rhs.org.uk/vegetables/french-beans/grow-your-own"],["Old Farmer's Almanac — Beans","https://www.almanac.com/plant/beans"],["University of Maryland Extension — Beans","https://extension.umd.edu/resource/snap-beans"]]],
 ["Flat beans & runner beans",[["RHS — Runner beans","https://www.rhs.org.uk/vegetables/runner-beans/grow-your-own"],["Marshalls Seeds — Runner beans","https://www.marshalls-seeds.co.uk"]]],
 ["Broad beans",[["RHS — Broad beans","https://www.rhs.org.uk/vegetables/broad-beans/grow-your-own"],["Old Farmer's Almanac — Fava/broad beans","https://www.almanac.com/plant/fava-beans"]]],
 ["Peas — sugarsnap & mangetout",[["RHS — Peas","https://www.rhs.org.uk/vegetables/peas/grow-your-own"],["Old Farmer's Almanac — Peas","https://www.almanac.com/plant/peas"]]],
 ["Tenderstem & organic tenderstem broccoli",[["RHS — Broccoli","https://www.rhs.org.uk/vegetables/broccoli/grow-your-own"],["Organic Growers Alliance","https://organicgrowersalliance.co.uk"],["GrowOrganic — Broccoli","https://www.groworganic.com"]]],
 ["Chillies (birdseye, finger, red, mixed, scotch bonnet)",[["PepperGeek — Growing peppers","https://peppergeek.com/how-to-grow-peppers/"],["Blooming Expert — Chilli guide","https://bloomingexpert.com"]]],
 ["Baby aubergine",[["RHS — Aubergines","https://www.rhs.org.uk/vegetables/aubergines/grow-your-own"],["University of Maryland Extension — Eggplant","https://extension.umd.edu/resource/eggplant"]]],
 ["Baby corn",[["University of Maryland Extension — Sweet corn","https://extension.umd.edu/resource/sweet-corn"],["Old Farmer's Almanac — Corn","https://www.almanac.com/plant/corn"]]],
 ["Okra",[["Sow Right Seeds — Okra guide","https://sowrightseeds.com"],["University of Maryland Extension — Okra","https://extension.umd.edu/resource/okra"]]],
 ["Dudhi (bottle gourd)",[["BigHaat — Bottle gourd","https://www.bighaat.com"],["PlotMyGarden — Bottle gourd","https://plotmygarden.com"]]],
 ["Turmeric",[["HOSS Tools — Growing turmeric","https://hosstools.com"],["Old Farmer's Almanac — Turmeric","https://www.almanac.com/plant/turmeric"]]],
 ["General climate ranges (all crops)",[["FAO ECOCROP — crop climate requirements","https://gaez.fao.org/pages/ecocrop"],["Wikifarmer — crop growing guides","https://wikifarmer.com/library/en"]]]
];
function loadRefs(){
  const rows=REFS.map(r=>'<tr><td class="rc gridjs-td">'+r[0]+'</td><td>'+r[1].map(s=>'<a href="'+s[1]+'" target="_blank" rel="noopener noreferrer">'+s[0]+'</a>').join('<br>')+'</td></tr>').join('');
  document.getElementById("refsBody").innerHTML='<table class="reftbl"><thead><tr><th>Crop / group</th><th>Optimum-conditions &amp; threshold sources</th></tr></thead><tbody>'+rows+'</tbody></table>';
  state.loaded.refs=true;
}
function init(){
  console.log("DASH init; Chart="+!!window.Chart+" gridjs="+!!window.gridjs+" cowork="+!!window.cowork);
  rebuildLookups();
  updateHeaderMeta();
  const cSel=document.getElementById("fCountry");
  META.countries.forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;cSel.appendChild(o);});
  refreshLocationOptions();
  document.querySelectorAll(".tab").forEach(b=>b.addEventListener("click",()=>switchTab(b.dataset.tab)));
  (function(){const lo=document.getElementById("accLeadLo"),hi=document.getElementById("accLeadHi");if(lo&&hi){for(let d=ACC.leadMin;d<=ACC.leadMax;d++){lo.insertAdjacentHTML("beforeend","<option value=\""+d+"\">"+d+"</option>");hi.insertAdjacentHTML("beforeend","<option value=\""+d+"\">"+d+"</option>");}lo.value=ACC.leadMin;hi.value=ACC.leadMax;lo.addEventListener("change",loadAccuracy);hi.addEventListener("change",loadAccuracy);const rc=document.getElementById("accRain");if(rc)rc.addEventListener("change",loadAccuracy);}})();
  cSel.addEventListener("change",()=>{refreshLocationOptions();setDirty();if(state.active==="ask"){commitTopScope();updateAskScope();}});
  document.getElementById("fLocation").addEventListener("change",()=>{setDirty();if(state.active==="ask"){commitTopScope();updateAskScope();}});
  document.querySelectorAll("#presets button").forEach(b=>b.addEventListener("click",()=>{document.querySelectorAll("#presets button").forEach(x=>x.classList.remove("active"));b.classList.add("active");document.getElementById("fStart").value="";document.getElementById("fEnd").value="";setDirty();}));
  document.getElementById("fStart").addEventListener("change",setDirty);
  document.getElementById("fEnd").addEventListener("change",setDirty);
  document.getElementById("updateBtn").addEventListener("click",applyFilters);
  document.getElementById("askBtn").addEventListener("click",()=>{const q=document.getElementById("askInput").value.trim();if(q)runAsk(q);});
  document.getElementById("askWebBtn").addEventListener("click",()=>{const q=document.getElementById("askInput").value.trim();if(q)askWeb(q);});
  document.getElementById("askInput").addEventListener("keydown",e=>{if(e.key==="Enter"){const q=e.target.value.trim();if(q)runAsk(q);}});
  const chips=document.getElementById("askChips");EXAMPLES.forEach(x=>{const c=document.createElement("span");c.className="chip";c.textContent=x;c.addEventListener("click",()=>runAsk(x));chips.appendChild(c);});
  const ccSel=document.getElementById("cropCountry");[...new Set(CROP.rows.map(r=>r.c))].sort().forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;ccSel.appendChild(o);});
  const cropSelEl=document.getElementById("cropSel");CROP.crops.forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;cropSelEl.appendChild(o);});state.crops=(state.crops||[]).filter(c=>CROP.crops.indexOf(c)>=0);if(!state.crops.length)state.crops=[CROP.crops[0]];
  const cpc=document.getElementById("cpCountry");META.countries.forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;cpc.appendChild(o);});
  const cpk=document.getElementById("cpCrop");CROP.crops.forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;cpk.appendChild(o);});
  cpc.addEventListener("change",()=>{state.loaded.cropplan=false;loadCropPlan();});
  cpk.addEventListener("change",()=>{state.loaded.cropplan=false;loadCropPlan();});
  const bcK=document.getElementById("bcCrop");CROP.crops.forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;bcK.appendChild(o);});
  const bcC=document.getElementById("bcCountry");[...new Set(CROP.rows.map(r=>r.c))].sort().forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;bcC.appendChild(o);});
  const bcG=document.getElementById("bcGrower");[...new Set(CROP.rows.map(r=>r.g))].sort().forEach(g=>{const o=document.createElement("option");o.value=g;o.textContent=g;bcG.appendChild(o);});
  const bcS=document.getElementById("bcSeason");MON.forEach((m,i)=>{const o=document.createElement("option");o.value=String(i+1);o.textContent=m;bcS.appendChild(o);});
  ["bcCrop","bcCountry","bcGrower","bcSeason"].forEach(id=>document.getElementById(id).addEventListener("change",()=>{state.loaded.bycountry=false;loadByCountry();}));
  document.getElementById("cropSel").addEventListener("change",setCropDirty);
  ccSel.addEventListener("change",()=>{const gs=document.getElementById("cropGrower");[...gs.options].forEach(o=>o.selected=false);refreshCropGrowers();refreshMSel("cropGrower");setCropDirty();});
  document.getElementById("cropGrower").addEventListener("change",setCropDirty);
  refreshCropGrowers();
  ["cropSel","cropGrower","cpCrop","bcCrop"].forEach(id=>enhanceMSelect(id,id==="cropGrower"?"All growers":"All crops"));
  document.getElementById("cropPeriod").addEventListener("change",setCropDirty);
  document.getElementById("cropUpdateBtn").addEventListener("click",applyCropFilters);
  const fs=document.getElementById("fStart"),fe=document.getElementById("fEnd");fs.min=META.obs_dates[0];fs.max=META.obs_dates[1];fe.min=META.obs_dates[0];fe.max=META.obs_dates[1];
  updateAskScope();
  mapInit();
  whenReady(()=>switchTab("map"));
}
init();
