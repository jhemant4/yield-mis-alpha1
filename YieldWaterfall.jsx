import { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

const INDIA1_ZONES = ["Central", "East", "North 1", "North 2"];
const INDIA2_ZONES = ["Maharashtra", "South 1", "South 2"];
const ALL_ZONES = [...INDIA1_ZONES, ...INDIA2_ZONES];
const ZONE_TO_INDIA = {};
INDIA1_ZONES.forEach(z => { ZONE_TO_INDIA[z] = "India 1"; });
INDIA2_ZONES.forEach(z => { ZONE_TO_INDIA[z] = "India 2"; });
const MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const C = {
  navy:"#1B2B4B", orange:"#E85D26", green:"#16a34a", red:"#dc2626",
  bg:"#F0F2F5", card:"#FFFFFF", border:"#E2E6EC",
  t1:"#0D1B2E", t2:"#475569", t3:"#94A3B8", blue:"#2563eb", lightBlue:"#EFF6FF",
};

// ── MONTH SORT ────────────────────────────────────────────────────────────────
function parseMonthScore(m) {
  const clean = String(m).replace(/[\u2018\u2019\u201A\u2032\u0060'`\/\-]/g," ").replace(/\s+/g," ").trim();
  const parts = clean.split(" ");
  const monIdx = MONTH_ORDER.findIndex(mo => mo.toLowerCase() === (parts[0]||"").toLowerCase().slice(0,3));
  if (monIdx >= 0 && parts[1]) {
    const yr = parseInt(parts[1]);
    return (yr < 100 ? 2000+yr : yr)*12 + monIdx;
  }
  return 0;
}

// ── YIELD LOGIC ───────────────────────────────────────────────────────────────
function computeYield(rows) {
  const d = rows.reduce((s,r)=>s+(r.disbCr||0),0);
  const w = rows.reduce((s,r)=>s+(r.wrr||0),0);
  return d>0 ? w/d : 0;
}

function computeWaterfall(p1,p2) {
  if (!p1?.length || !p2?.length) return null;
  const p1Y=computeYield(p1), p2Y=computeYield(p2);
  const p2Disb=p2.reduce((s,r)=>s+r.disbCr,0);
  const products=[...new Set([...p1,...p2].map(r=>r.product).filter(Boolean))];
  const p1PY={};
  products.forEach(pt=>{ p1PY[pt]=computeYield(p1.filter(r=>r.product===pt)); });
  const p2Mix={};
  products.forEach(pt=>{
    const d=p2.filter(r=>r.product===pt).reduce((s,r)=>s+r.disbCr,0);
    p2Mix[pt]=p2Disb>0?d/p2Disb:0;
  });
  const hypo=products.reduce((s,pt)=>s+p2Mix[pt]*(p1PY[pt]||0),0);
  return { p1Yield:p1Y, p2Yield:p2Y, mixEffect:hypo-p1Y, absChange:p2Y-hypo,
           p1Disb:p1.reduce((s,r)=>s+r.disbCr,0), p2Disb };
}

// ── PARSER ────────────────────────────────────────────────────────────────────
// Column map is hard-coded per sheet structure since the two sheets
// have completely different column names for the same fields.
//
// Sheet 1 (e.g. "Oct'25" / Mar'26 style — old format):
//   Disb in Cr  | DISB_ROI  | WRR (pre-computed) | Business Month
//   PRODUCT_TYPE | ZONE | GEO | BRANCH
//
// Sheet 2 (e.g. "Apr'26" — new format):
//   Disb Amt in Cr | Disb ROI | (no WRR — computed as disbCr*roi) | Business Month
//   Product Type | HL Zone | HL Geo | Branch | Country

function getColIndex(headers, ...names) {
  const norm = s => String(s||"").trim().toLowerCase().replace(/\s+/g," ");
  // exact match first
  for (const n of names) {
    const i = headers.findIndex(h => norm(h) === norm(n));
    if (i >= 0) return i;
  }
  // contains match fallback
  for (const n of names) {
    const i = headers.findIndex(h => norm(h).includes(norm(n)));
    if (i >= 0) return i;
  }
  return -1;
}

function detectSheetFormat(headers) {
  const norm = s => String(s||"").trim().toLowerCase();
  const has = name => headers.some(h => norm(h) === norm(name));

  // Sheet 1 (old format) has WRR and ZONE and DISB_ROI
  if (has("WRR") && has("ZONE")) return "sheet1";
  // Sheet 2 (new format) has HL Zone and Disb Amt in Cr
  if (has("HL Zone") && has("Disb Amt in Cr")) return "sheet2";
  // Generic fallback
  return "generic";
}

function parseWorkbook(wb) {
  let rows = [];
  const parseLog = { sheets: {} };

  wb.SheetNames.forEach(sheetName => {
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
    if (!json.length) return;

    const headers = Object.keys(json[0]);
    const fmt = detectSheetFormat(headers);

    let disbIdx, roiIdx, wrrIdx, monthIdx, productIdx, zoneIdx, geoIdx, branchIdx, countryIdx;

    if (fmt === "sheet1") {
      // Old format — exact column names from Sheet 1
      disbIdx    = getColIndex(headers, "Disb in Cr");
      roiIdx     = getColIndex(headers, "DISB_ROI");
      wrrIdx     = getColIndex(headers, "WRR");          // pre-computed
      monthIdx   = getColIndex(headers, "Business Month");
      productIdx = getColIndex(headers, "PRODUCT_TYPE");
      zoneIdx    = getColIndex(headers, "ZONE");
      geoIdx     = getColIndex(headers, "GEO");
      branchIdx  = getColIndex(headers, "BRANCH");
      countryIdx = -1; // absent — derive from zone
    } else if (fmt === "sheet2") {
      // New format — exact column names from Sheet 2
      disbIdx    = getColIndex(headers, "Disb Amt in Cr");
      roiIdx     = getColIndex(headers, "Disb ROI");
      wrrIdx     = -1; // absent — compute as disbCr * roi
      monthIdx   = getColIndex(headers, "Business Month");
      productIdx = getColIndex(headers, "Product Type");
      zoneIdx    = getColIndex(headers, "HL Zone");
      geoIdx     = getColIndex(headers, "HL Geo");
      branchIdx  = getColIndex(headers, "Branch");
      countryIdx = getColIndex(headers, "Country");
    } else {
      // Generic fallback — try common aliases
      disbIdx    = getColIndex(headers, "Disb Amt in Cr","Disb in Cr","Disb cr","Disb Amt","DISB_AMOUNT");
      roiIdx     = getColIndex(headers, "Disb ROI","DISB_ROI","ROI");
      wrrIdx     = getColIndex(headers, "WRR","wrr");
      monthIdx   = getColIndex(headers, "Business Month","BusinessMonth","Period","Month");
      productIdx = getColIndex(headers, "Product Type","PRODUCT_TYPE","Product");
      zoneIdx    = getColIndex(headers, "HL Zone","ZONE","Zone");
      geoIdx     = getColIndex(headers, "HL Geo","GEO","Geo");
      branchIdx  = getColIndex(headers, "Branch","BRANCH");
      countryIdx = getColIndex(headers, "Country","COUNTRY");
    }

    const sheetLog = { fmt, disbIdx, roiIdx, wrrIdx, monthIdx, productIdx, zoneIdx, geoIdx, total:0, kept:0, byMonth:{} };
    parseLog.sheets[sheetName] = sheetLog;

    if (disbIdx < 0) {
      console.warn(`[Parser] Sheet "${sheetName}": no disbursement column found`);
      return;
    }

    json.forEach(r => {
      const vals = Object.values(r);
      sheetLog.total++;

      const disbCr = parseFloat(vals[disbIdx]) || 0;
      if (disbCr <= 0) return;

      const roi = parseFloat(vals[roiIdx]) || 0;

      // WRR: use pre-computed column if present (Sheet 1), else compute
      let wrr;
      if (wrrIdx >= 0) {
        wrr = parseFloat(vals[wrrIdx]) || (disbCr * roi);
      } else {
        wrr = disbCr * roi;
      }

      // Month from Business Month column (always present in both sheets)
      let month = "";
      if (monthIdx >= 0) {
        const mv = vals[monthIdx];
        if (mv instanceof Date) {
          month = MONTH_ORDER[mv.getMonth()] + "'" + String(mv.getFullYear()).slice(2);
        } else {
          month = String(mv || "").trim();
        }
      }
      // Fallback to sheet name if column blank
      if (!month || month === "0" || month.toLowerCase() === "nan") {
        month = sheetName.trim();
      }
      if (!month) return;

      const zone    = zoneIdx    >= 0 ? String(vals[zoneIdx]    || "").trim() : "";
      const geo     = geoIdx     >= 0 ? String(vals[geoIdx]     || "").trim() : "";
      const branch  = branchIdx  >= 0 ? String(vals[branchIdx]  || "").trim() : "";
      const product = productIdx >= 0 ? String(vals[productIdx] || "").trim() : "";
      const rawCtry = countryIdx >= 0 ? String(vals[countryIdx] || "").trim() : "";
      const country = rawCtry || (INDIA1_ZONES.includes(zone) ? "India 1" : zone ? "India 2" : "Unknown");

      rows.push({ disbCr, wrr, roi, month, product, zone, geo, branch, country });
      sheetLog.kept++;
      sheetLog.byMonth[month] = (sheetLog.byMonth[month] || 0) + 1;
    });
  });

  console.log("[Parser] full log:", JSON.stringify(parseLog, null, 2));

  const uniqueMonths = [...new Set(rows.map(r=>r.month))].sort((a,b)=>parseMonthScore(a)-parseMonthScore(b));
  return { rows, uniqueMonths, sheets: wb.SheetNames, parseLog };
}

// ── WATERFALL CHART ───────────────────────────────────────────────────────────
function WFChart({data,p1Label,p2Label}) {
  const [tip,setTip]=useState(null);
  if (!data) return <div style={{height:120,display:"flex",alignItems:"center",justifyContent:"center",color:C.t3,fontSize:11}}>No data</div>;
  const {p1Yield,mixEffect,absChange,p2Yield}=data;
  const W=280,H=130,pL=8,pR=8,pT=28,pB=36;
  const cW=W-pL-pR,cH=H-pT-pB;
  const minY=Math.min(p1Yield,p2Yield)-0.4, maxY=Math.max(p1Yield,p2Yield)+0.4, rng=maxY-minY;
  const toY=v=>pT+cH-((v-minY)/rng)*cH;
  const toH=v=>(Math.abs(v)/rng)*cH;
  const bW=cW/5.5, gap=(cW-bW*4)/5;
  const xs=[pL+gap, pL+gap*2+bW, pL+gap*3+bW*2, pL+gap*4+bW*3];
  const b1T=toY(p1Yield), b1H=toY(minY)-b1T;
  const m2T=mixEffect>=0?toY(p1Yield+mixEffect):toY(p1Yield), m2H=toH(mixEffect);
  const a3B=p1Yield+mixEffect, a3T=absChange>=0?toY(a3B+absChange):toY(a3B), a3H=toH(absChange);
  const b4T=toY(p2Yield), b4H=toY(minY)-b4T;
  const fmt=v=>v.toFixed(2)+"%", fmtD=v=>(v>=0?"+":"")+v.toFixed(2);
  const labels=[p1Label,"Prod Mix","Absolute",p2Label];
  return (
    <div style={{position:"relative"}}>
      <svg width={W} height={H} style={{overflow:"visible"}}>
        <line x1={xs[0]+bW} y1={toY(p1Yield)} x2={xs[1]} y2={toY(p1Yield)} stroke={C.border} strokeWidth={1} strokeDasharray="3,2"/>
        <line x1={xs[1]+bW} y1={toY(p1Yield+mixEffect)} x2={xs[2]} y2={toY(p1Yield+mixEffect)} stroke={C.border} strokeWidth={1} strokeDasharray="3,2"/>
        <line x1={xs[2]+bW} y1={toY(p2Yield)} x2={xs[3]} y2={toY(p2Yield)} stroke={C.border} strokeWidth={1} strokeDasharray="3,2"/>
        <rect x={xs[0]} y={b1T} width={bW} height={b1H} fill={C.navy} rx={2} style={{cursor:"pointer"}} onMouseEnter={()=>setTip({x:xs[0]+bW/2,y:b1T-6,t:fmt(p1Yield)})} onMouseLeave={()=>setTip(null)}/>
        <text x={xs[0]+bW/2} y={b1T-5} textAnchor="middle" fontSize={9} fontWeight="700" fill={C.t1}>{fmt(p1Yield)}</text>
        <rect x={xs[1]} y={m2T} width={bW} height={Math.max(m2H,2)} fill={mixEffect>=0?C.green:C.red} rx={2} style={{cursor:"pointer"}} onMouseEnter={()=>setTip({x:xs[1]+bW/2,y:m2T-6,t:fmtD(mixEffect)})} onMouseLeave={()=>setTip(null)}/>
        <text x={xs[1]+bW/2} y={mixEffect>=0?m2T-5:m2T+Math.max(m2H,2)+12} textAnchor="middle" fontSize={8} fontWeight="700" fill={mixEffect>=0?C.green:C.red}>{fmtD(mixEffect)}</text>
        <rect x={xs[2]} y={a3T} width={bW} height={Math.max(a3H,2)} fill={absChange>=0?C.green:C.red} rx={2} style={{cursor:"pointer"}} onMouseEnter={()=>setTip({x:xs[2]+bW/2,y:a3T-6,t:fmtD(absChange)})} onMouseLeave={()=>setTip(null)}/>
        <text x={xs[2]+bW/2} y={absChange>=0?a3T-5:a3T+Math.max(a3H,2)+12} textAnchor="middle" fontSize={8} fontWeight="700" fill={absChange>=0?C.green:C.red}>{fmtD(absChange)}</text>
        <rect x={xs[3]} y={b4T} width={bW} height={b4H} fill={C.orange} rx={2} style={{cursor:"pointer"}} onMouseEnter={()=>setTip({x:xs[3]+bW/2,y:b4T-6,t:fmt(p2Yield)})} onMouseLeave={()=>setTip(null)}/>
        <text x={xs[3]+bW/2} y={b4T-5} textAnchor="middle" fontSize={9} fontWeight="700" fill={C.t1}>{fmt(p2Yield)}</text>
        {xs.map((x,i)=>(
          <text key={i} x={x+bW/2} y={H-pB+14} textAnchor="middle" fontSize={7.5} fill={C.t2} fontWeight={i===0||i===3?"600":"400"}>{labels[i]}</text>
        ))}
        {tip&&<g><rect x={tip.x-30} y={tip.y-18} width={60} height={20} fill={C.t1} rx={3} opacity={0.9}/><text x={tip.x} y={tip.y-5} textAnchor="middle" fontSize={9} fontWeight="700" fill="#fff">{tip.t}</text></g>}
      </svg>
    </div>
  );
}

function Delta({v}) {
  const pos=v>=0;
  return <span style={{display:"inline-flex",alignItems:"center",gap:3,background:pos?"#dcfce7":"#fee2e2",color:pos?"#166534":"#991b1b",borderRadius:5,fontSize:11,fontWeight:700,padding:"2px 7px"}}>{pos?"▲":"▼"} {Math.abs(v*100).toFixed(2)} bps</span>;
}

function WFCard({label,sublabel,data,p1,p2}) {
  if (!data) return null;
  const net=data.p2Yield-data.p1Yield;
  return (
    <div style={{background:C.card,borderRadius:10,border:`1px solid ${C.border}`,overflow:"hidden",transition:"box-shadow 0.15s"}}
      onMouseEnter={e=>e.currentTarget.style.boxShadow="0 4px 16px rgba(0,0,0,0.09)"}
      onMouseLeave={e=>e.currentTarget.style.boxShadow=""}>
      <div style={{padding:"10px 12px 4px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div>
          <div style={{fontSize:12,fontWeight:700,color:C.t1}}>{label}</div>
          {sublabel&&<div style={{fontSize:10,color:C.t3,marginTop:1}}>{sublabel}</div>}
        </div>
        <Delta v={net}/>
      </div>
      <div style={{display:"flex",justifyContent:"center",padding:"0 4px"}}>
        <WFChart data={data} p1Label={p1} p2Label={p2}/>
      </div>
      <div style={{padding:"6px 12px 10px",borderTop:`1px solid ${C.bg}`,display:"flex",justifyContent:"space-between"}}>
        <KV label={`Disb ${p1}`} val={data.p1Disb.toFixed(1)+" Crs"} color={C.navy}/>
        <KV label="Mix shift" val={(data.mixEffect>=0?"+":"")+data.mixEffect.toFixed(2)} color={data.mixEffect>=0?C.green:C.red}/>
        <KV label={`Disb ${p2}`} val={data.p2Disb.toFixed(1)+" Crs"} color={C.navy}/>
      </div>
    </div>
  );
}

function KV({label,val,color}) {
  return <div style={{textAlign:"center"}}><div style={{fontSize:8,color:C.t3,textTransform:"uppercase",letterSpacing:"0.04em"}}>{label}</div><div style={{fontSize:11,fontWeight:600,color}}>{val}</div></div>;
}

function SortBar({sortDir,onSort,label="yield change"}) {
  return (
    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
      <span style={{fontSize:11,color:C.t2,fontWeight:500}}>Sort by {label}:</span>
      <button onClick={()=>onSort("asc")} style={{padding:"4px 10px",fontSize:11,fontWeight:600,borderRadius:5,border:`1px solid ${C.border}`,background:sortDir==="asc"?C.navy:"transparent",color:sortDir==="asc"?"#fff":C.t2,cursor:"pointer"}}>▲ Worst first</button>
      <button onClick={()=>onSort("desc")} style={{padding:"4px 10px",fontSize:11,fontWeight:600,borderRadius:5,border:`1px solid ${C.border}`,background:sortDir==="desc"?C.navy:"transparent",color:sortDir==="desc"?"#fff":C.t2,cursor:"pointer"}}>▼ Best first</button>
      {sortDir&&<button onClick={()=>onSort(null)} style={{padding:"4px 8px",fontSize:11,color:C.red,background:"#fee2e2",border:"none",borderRadius:5,cursor:"pointer",fontWeight:600}}>✕ Clear</button>}
    </div>
  );
}

function HeaderKPI({label,value,color}) {
  return (
    <div style={{textAlign:"center"}}>
      <div style={{fontSize:9,color:"#8BA3C4",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:2}}>{label}</div>
      <div style={{fontSize:15,fontWeight:700,color:color||"#fff"}}>{value}</div>
    </div>
  );
}

// ── DATA PREVIEW ──────────────────────────────────────────────────────────────
function DataPreview({rows,period1,period2,parseLog,sheets}) {
  const p1rows=useMemo(()=>rows.filter(r=>r.month===period1),[rows,period1]);
  const p2rows=useMemo(()=>rows.filter(r=>r.month===period2),[rows,period2]);
  const stats=[
    {label:`${period1} rows`,val:p1rows.length.toLocaleString(),ok:p1rows.length>0},
    {label:`${period2} rows`,val:p2rows.length.toLocaleString(),ok:p2rows.length>0},
    {label:"Zones",val:new Set(rows.map(r=>r.zone).filter(Boolean)).size},
    {label:"Geos",val:new Set(rows.map(r=>r.geo).filter(Boolean)).size},
    {label:"Products",val:new Set(rows.map(r=>r.product).filter(Boolean)).size},
    {label:"Total rows",val:rows.length.toLocaleString()},
  ];
  const cols=["month","zone","geo","product","disbCr","roi","wrr","country"];
  const thSt={padding:"6px 10px",fontSize:10,fontWeight:600,color:C.t2,borderBottom:`1px solid ${C.border}`,textTransform:"uppercase",letterSpacing:"0.04em",background:C.bg,whiteSpace:"nowrap"};
  const tdSt={padding:"6px 10px",fontSize:11,color:C.t1,borderBottom:`1px solid ${C.border}`,fontFamily:"monospace"};

  // All months in data
  const allMonths={};
  rows.forEach(r=>{ allMonths[r.month]=(allMonths[r.month]||0)+1; });

  return (
    <div style={{padding:"16px 24px"}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:10,marginBottom:16}}>
        {stats.map((s,i)=>(
          <div key={i} style={{background:C.card,border:`1px solid ${s.ok===false?C.red:C.border}`,borderRadius:8,padding:"10px 14px"}}>
            <div style={{fontSize:10,color:C.t3,textTransform:"uppercase",letterSpacing:"0.04em"}}>{s.label}</div>
            <div style={{fontSize:18,fontWeight:700,color:s.ok===false?C.red:C.navy,marginTop:2}}>{s.val}</div>
          </div>
        ))}
      </div>

      {/* All months breakdown */}
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:8,padding:"12px 14px",marginBottom:12}}>
        <div style={{fontSize:11,fontWeight:600,color:C.t2,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.04em"}}>All months in data (from Business Month column)</div>
        <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
          {Object.entries(allMonths).sort((a,b)=>parseMonthScore(a[0])-parseMonthScore(b[0])).map(([m,cnt])=>(
            <div key={m} style={{background:m===period1||m===period2?"#EFF6FF":C.bg,border:`1px solid ${m===period1||m===period2?"#BFDBFE":C.border}`,borderRadius:6,padding:"6px 12px"}}>
              <div style={{fontSize:11,fontWeight:600,color:m===period1||m===period2?C.blue:C.t1}}>{m}</div>
              <div style={{fontSize:12,fontFamily:"monospace",color:C.t2}}>{cnt.toLocaleString()} rows</div>
            </div>
          ))}
        </div>
      </div>

      {/* Sheet detection */}
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:8,padding:"12px 14px",marginBottom:12}}>
        <div style={{fontSize:11,fontWeight:600,color:C.t2,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.04em"}}>Sheet detection</div>
        {Object.entries(parseLog?.sheets||{}).map(([sh,log])=>(
          <div key={sh} style={{display:"flex",gap:16,flexWrap:"wrap",padding:"6px 0",borderBottom:`1px solid ${C.bg}`}}>
            <span style={{fontSize:12,fontWeight:600,color:C.navy,minWidth:90,fontFamily:"monospace"}}>{sh}</span>
            <span style={{fontSize:11,color:C.t2}}>Format: <strong>{log.fmt}</strong></span>
            <span style={{fontSize:11,color:C.t2}}>Parsed: <strong>{log.kept?.toLocaleString()}</strong> / {log.total?.toLocaleString()}</span>
            <span style={{fontSize:11,color:C.t2}}>Months: {JSON.stringify(log.byMonth)}</span>
          </div>
        ))}
      </div>

      {/* Sample rows */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        {[{title:`${period1} — first 5 rows`,rows:p1rows.slice(0,5)},{title:`${period2} — first 5 rows`,rows:p2rows.slice(0,5)}].map((block,bi)=>(
          <div key={bi} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:8,overflow:"hidden"}}>
            <div style={{padding:"8px 12px",borderBottom:`1px solid ${C.border}`,fontSize:12,fontWeight:600,color:C.t1}}>{block.title}</div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr>{cols.map(c=><th key={c} style={thSt}>{c}</th>)}</tr></thead>
                <tbody>{block.rows.length>0?block.rows.map((r,i)=>(
                  <tr key={i}>{cols.map(c=><td key={c} style={tdSt}>{typeof r[c]==="number"?r[c].toFixed(3):r[c]||"—"}</td>)}</tr>
                )):<tr><td colSpan={cols.length} style={{...tdSt,color:C.t3,textAlign:"center",padding:16}}>No rows for this period</td></tr>}</tbody>
              </table>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
export default function YieldWaterfall() {
  const [rawData,setRawData]     = useState([]);
  const [months,setMonths]       = useState([]);
  const [period1,setPeriod1]     = useState("");
  const [period2,setPeriod2]     = useState("");
  const [product,setProduct]     = useState("All");
  const [activeTab,setActiveTab] = useState("preview");
  const [geoSort,setGeoSort]     = useState(null);
  const [zoneSort,setZoneSort]   = useState(null);
  const [dragOver,setDragOver]   = useState(false);
  const [loading,setLoading]     = useState(false);
  const [fileName,setFileName]   = useState("");
  const [sheets,setSheets]       = useState([]);
  const [parseLog,setParseLog]   = useState({});
  const fileRef=useRef();

  const processFile=useCallback((file)=>{
    if(!file) return;
    setLoading(true);
    setFileName(file.name);
    const reader=new FileReader();
    reader.onload=e=>{
      try {
        const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});
        const {rows,uniqueMonths,sheets:sh,parseLog:pl}=parseWorkbook(wb);
        setRawData(rows);
        setMonths(uniqueMonths);
        // Auto-select: period1 = earliest, period2 = latest
        setPeriod1(uniqueMonths[0]||"");
        setPeriod2(uniqueMonths[uniqueMonths.length-1]||"");
        setSheets(sh);
        setParseLog(pl);
        setActiveTab("preview");
      } catch(err){console.error("[Parser error]",err);}
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  },[]);

  const products=useMemo(()=>["All",...new Set(rawData.map(r=>r.product).filter(Boolean))]
    .sort((a,b)=>a==="All"?-1:a.localeCompare(b)),[rawData]);

  const filter=useCallback((rows,period)=>{
    return rows.filter(r=>{
      if(r.month!==period) return false;
      if(product!=="All"&&r.product!==product) return false;
      return true;
    });
  },[product]);

  const p1All=useMemo(()=>filter(rawData,period1),[rawData,period1,filter]);
  const p2All=useMemo(()=>filter(rawData,period2),[rawData,period2,filter]);

  const panData =useMemo(()=>computeWaterfall(p1All,p2All),[p1All,p2All]);
  const ind1Data=useMemo(()=>computeWaterfall(p1All.filter(r=>r.country==="India 1"),p2All.filter(r=>r.country==="India 1")),[p1All,p2All]);
  const ind2Data=useMemo(()=>computeWaterfall(p1All.filter(r=>r.country==="India 2"),p2All.filter(r=>r.country==="India 2")),[p1All,p2All]);

  const zoneItems=useMemo(()=>{
    const zones=[...new Set([...p1All,...p2All].map(r=>r.zone).filter(Boolean))];
    const items=zones.map(z=>({
      label:z, sublabel:ZONE_TO_INDIA[z]||"",
      data:computeWaterfall(p1All.filter(r=>r.zone===z),p2All.filter(r=>r.zone===z))
    })).filter(i=>i.data);
    if(zoneSort==="asc") return [...items].sort((a,b)=>(a.data.p2Yield-a.data.p1Yield)-(b.data.p2Yield-b.data.p1Yield));
    if(zoneSort==="desc") return [...items].sort((a,b)=>(b.data.p2Yield-b.data.p1Yield)-(a.data.p2Yield-a.data.p1Yield));
    return items;
  },[p1All,p2All,zoneSort]);

  const geoItems=useMemo(()=>{
    const geos=[...new Set([...p1All,...p2All].map(r=>r.geo).filter(Boolean))];
    const items=geos.map(g=>({
      label:g, sublabel:(p1All.find(r=>r.geo===g)||p2All.find(r=>r.geo===g))?.zone||"",
      data:computeWaterfall(p1All.filter(r=>r.geo===g),p2All.filter(r=>r.geo===g))
    })).filter(i=>i.data);
    if(geoSort==="asc") return [...items].sort((a,b)=>(a.data.p2Yield-a.data.p1Yield)-(b.data.p2Yield-b.data.p1Yield));
    if(geoSort==="desc") return [...items].sort((a,b)=>(b.data.p2Yield-b.data.p1Yield)-(a.data.p2Yield-a.data.p1Yield));
    return items;
  },[p1All,p2All,geoSort]);

  const hasData=rawData.length>0;
  const tabs=[{id:"preview",label:"Data Preview"},{id:"pan",label:"PAN India"},{id:"zones",label:"Zones"},{id:"geos",label:"Geos"}];

  return (
    <div style={{fontFamily:"'DM Sans','Segoe UI',sans-serif",minHeight:"100vh",background:C.bg,color:C.t1}}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>

      {/* HEADER */}
      <div style={{background:C.navy,padding:"0 24px",height:56,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:6,height:30,background:C.orange,borderRadius:2}}/>
          <div>
            <div style={{fontSize:14,fontWeight:700,color:"#fff",letterSpacing:"-0.01em"}}>Yield Waterfall Analysis</div>
            <div style={{fontSize:10,color:"#8BA3C4"}}>Month-on-Month Decomposition · Piramal Finance</div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {fileName&&<span style={{fontSize:10,color:"#8BA3C4",fontFamily:"'DM Mono',monospace"}}>{fileName}</span>}
          <button onClick={()=>fileRef.current?.click()} style={{display:"flex",alignItems:"center",gap:6,background:C.orange,color:"#fff",border:"none",borderRadius:7,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>
            ↑ {hasData?"Re-upload":"Upload Data"}
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>e.target.files[0]&&processFile(e.target.files[0])}/>
        </div>
      </div>

      {!hasData ? (
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"calc(100vh - 56px)"}}>
          <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
            onDrop={e=>{e.preventDefault();setDragOver(false);processFile(e.dataTransfer.files[0]);}}
            onClick={()=>fileRef.current?.click()}
            style={{border:`2px dashed ${dragOver?C.orange:C.border}`,borderRadius:16,padding:"56px 80px",textAlign:"center",cursor:"pointer",background:dragOver?"#fff5f0":C.card,maxWidth:480}}>
            <div style={{width:56,height:56,borderRadius:14,background:"#fff0e8",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px",fontSize:28}}>📊</div>
            <p style={{fontSize:16,fontWeight:700,margin:"0 0 8px"}}>Drop your Excel file</p>
            <p style={{fontSize:12,color:C.t2,margin:"0 0 6px",lineHeight:1.7}}>Supports both old and new Piramal sheet formats<br/>Multiple months in one sheet handled automatically</p>
            <p style={{fontSize:11,color:C.t3,margin:0}}>{loading?"Parsing…":"Click or drag & drop"}</p>
          </div>
        </div>
      ) : (
        <>
          {/* BANNER */}
          <div style={{background:"#EFF6FF",borderBottom:`1px solid #BFDBFE`,padding:"6px 24px",display:"flex",gap:20,alignItems:"center",flexWrap:"wrap"}}>
            <span style={{fontSize:11,color:"#1D4ED8",fontWeight:600}}>✓ {rawData.length.toLocaleString()} rows loaded</span>
            <span style={{fontSize:11,color:"#1D4ED8"}}><strong>{period1}</strong>: {rawData.filter(r=>r.month===period1).length.toLocaleString()} rows</span>
            <span style={{fontSize:11,color:C.orange}}><strong>{period2}</strong>: {rawData.filter(r=>r.month===period2).length.toLocaleString()} rows</span>
            <span style={{fontSize:11,color:C.t2}}>Sheets: {sheets.join(" · ")}</span>
          </div>

          {/* CONTROLS */}
          <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,padding:"10px 24px",display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
            <div style={{display:"flex",alignItems:"center",gap:8,background:C.bg,borderRadius:8,padding:"5px 10px"}}>
              <div style={{width:10,height:10,borderRadius:2,background:C.navy}}/>
              <span style={{fontSize:11,color:C.t2,fontWeight:500}}>Period 1</span>
              <select value={period1} onChange={e=>setPeriod1(e.target.value)} style={{fontSize:12,padding:"3px 6px",border:`1.5px solid ${C.navy}44`,borderRadius:5,color:C.navy,fontWeight:700,cursor:"pointer",background:"#fff"}}>
                {months.map(m=><option key={m} value={m}>{m}</option>)}
              </select>
            </div>
            <span style={{color:C.t3,fontSize:14}}>⇄</span>
            <div style={{display:"flex",alignItems:"center",gap:8,background:C.bg,borderRadius:8,padding:"5px 10px"}}>
              <div style={{width:10,height:10,borderRadius:2,background:C.orange}}/>
              <span style={{fontSize:11,color:C.t2,fontWeight:500}}>Period 2</span>
              <select value={period2} onChange={e=>setPeriod2(e.target.value)} style={{fontSize:12,padding:"3px 6px",border:`1.5px solid ${C.orange}44`,borderRadius:5,color:C.orange,fontWeight:700,cursor:"pointer",background:"#fff"}}>
                {months.map(m=><option key={m} value={m}>{m}</option>)}
              </select>
            </div>
            <div style={{width:1,height:24,background:C.border}}/>
            <div style={{display:"flex",alignItems:"center",gap:6}}>
              <span style={{fontSize:11,color:C.t3,fontWeight:500}}>Product</span>
              <select value={product} onChange={e=>setProduct(e.target.value)} style={{fontSize:11,padding:"4px 8px",border:`1px solid ${C.border}`,borderRadius:5,background:product!=="All"?C.lightBlue:C.bg,color:product!=="All"?C.blue:C.t2,fontWeight:product!=="All"?600:400,cursor:"pointer"}}>
                {products.map(p=><option key={p} value={p}>{p}</option>)}
              </select>
            </div>
            <div style={{marginLeft:"auto",display:"flex",gap:4}}>
              {tabs.map(t=>(
                <button key={t.id} onClick={()=>setActiveTab(t.id)} style={{padding:"5px 12px",borderRadius:6,fontSize:12,fontWeight:600,cursor:"pointer",border:"none",background:activeTab===t.id?C.navy:"transparent",color:activeTab===t.id?"#fff":C.t2,transition:"all 0.15s"}}>{t.label}</button>
              ))}
            </div>
          </div>

          {/* KPI STRIP */}
          {activeTab!=="preview"&&panData&&(
            <div style={{background:C.navy,padding:"10px 24px",display:"flex",gap:28,alignItems:"center",flexWrap:"wrap"}}>
              <HeaderKPI label={`${period1} — Overall All Dump Yield`} value={panData.p1Yield.toFixed(3)+"%"} color="#8BA3C4"/>
              <HeaderKPI label="Product Mix Effect" value={(panData.mixEffect>=0?"+":"")+panData.mixEffect.toFixed(3)+"%"} color={panData.mixEffect>=0?"#6EE7B7":"#FCA5A5"}/>
              <HeaderKPI label="Absolute Yield Chg" value={(panData.absChange>=0?"+":"")+panData.absChange.toFixed(3)+"%"} color={panData.absChange>=0?"#6EE7B7":"#FCA5A5"}/>
              <HeaderKPI label={`${period2} — Overall All Dump Yield`} value={panData.p2Yield.toFixed(3)+"%"} color={C.orange}/>
              <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
                <span style={{fontSize:10,color:"#8BA3C4"}}>Net Δ</span>
                <Delta v={panData.p2Yield-panData.p1Yield}/>
              </div>
            </div>
          )}

          {activeTab==="preview"&&<DataPreview rows={rawData} period1={period1} period2={period2} parseLog={parseLog} sheets={sheets}/>}

          {activeTab==="pan"&&(
            <div style={{padding:"20px 24px"}}>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
                <WFCard label="PAN India" data={panData} p1={period1} p2={period2}/>
                <WFCard label="India 1" sublabel="Central · East · North 1 · North 2" data={ind1Data} p1={period1} p2={period2}/>
                <WFCard label="India 2" sublabel="Maharashtra · South 1 · South 2" data={ind2Data} p1={period1} p2={period2}/>
              </div>
            </div>
          )}

          {activeTab==="zones"&&(
            <div style={{padding:"20px 24px"}}>
              <SortBar sortDir={zoneSort} onSort={setZoneSort}/>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:14}}>
                {zoneItems.map(z=><WFCard key={z.label} label={z.label} sublabel={z.sublabel} data={z.data} p1={period1} p2={period2}/>)}
                {zoneItems.length===0&&<div style={{color:C.t3,fontSize:13,padding:"32px 0"}}>No zone data found.</div>}
              </div>
            </div>
          )}

          {activeTab==="geos"&&(
            <div style={{padding:"20px 24px"}}>
              <SortBar sortDir={geoSort} onSort={setGeoSort}/>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(270px,1fr))",gap:12}}>
                {geoItems.map(g=><WFCard key={g.label} label={g.label} sublabel={g.sublabel} data={g.data} p1={period1} p2={period2}/>)}
                {geoItems.length===0&&<div style={{color:C.t3,fontSize:13,padding:"32px 0"}}>No geo data found.</div>}
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}
