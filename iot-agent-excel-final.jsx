// ShopFloor IQ — IoT Predictive Maintenance Agent (Excel-powered)
// Upload the demo Excel file to load machine telemetry & run AI analysis

import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ── Palette ──────────────────────────────────────────────────────────────────
const C = {
  bg:"#090E18", panel:"#0F1624", border:"#1E2D45",
  accent:"#00C2FF", safe:"#10B981", warn:"#F59E0B", danger:"#EF4444",
  muted:"#4A6080", text:"#C9D8EC", textDim:"#6B82A0",
};
const STATUS_COLOR = {NOMINAL:C.safe, WARNING:C.warn, CRITICAL:C.danger};

// ── Helpers ──────────────────────────────────────────────────────────────────
function riskScore(avg) {
  let s = 0;
  if (!avg) return 0;
  if (avg.temp > 85) s+=30; else if (avg.temp > 75) s+=15;
  if (avg.vib > 6.5) s+=35; else if (avg.vib > 5) s+=18;
  if (avg.oil < 30)  s+=25; else if (avg.oil < 50) s+=10;
  if (avg.curr > 17) s+=15;
  return Math.min(100, s);
}
function statusFromRisk(r) { return r>=65?"CRITICAL":r>=35?"WARNING":"NOMINAL"; }

// ── GaugeBar ─────────────────────────────────────────────────────────────────
function GaugeBar({ value, max, color }) {
  return (
    <div style={{background:C.border,borderRadius:4,height:5,width:"100%",overflow:"hidden"}}>
      <div style={{width:`${Math.min(100,(value/max)*100)}%`,height:"100%",background:color,
        borderRadius:4,transition:"width 0.4s",boxShadow:`0 0 6px ${color}88`}}/>
    </div>
  );
}

// ── Sparkline ─────────────────────────────────────────────────────────────────
function Sparkline({ data, color, w=200, h=36 }) {
  if (!data||data.length<2) return null;
  const min=Math.min(...data),max=Math.max(...data),range=max-min||1;
  const pts=data.map((v,i)=>`${(i/(data.length-1))*w},${h-((v-min)/range)*(h-6)-3}`).join(" ");
  return (
    <svg width={w} height={h} style={{display:"block"}}>
      <defs>
        <linearGradient id="sg" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={color} stopOpacity="0.3"/>
          <stop offset="100%" stopColor={color} stopOpacity="0"/>
        </linearGradient>
      </defs>
      <polyline points={pts} fill="none" stroke={color} strokeWidth="1.8" strokeLinejoin="round"/>
    </svg>
  );
}

// ── Machine Card ──────────────────────────────────────────────────────────────
function MachineCard({ m, selected, onClick }) {
  const risk = riskScore(m.avg);
  const status = statusFromRisk(risk);
  const sc = STATUS_COLOR[status];
  const isSelected = selected === m.id;

  return (
    <div onClick={onClick} style={{
      background:C.panel, border:`1px solid ${isSelected?sc:C.border}`,
      borderRadius:10, padding:"14px 16px", cursor:"pointer",
      boxShadow:isSelected?`0 0 16px ${sc}44`:"none", transition:"all 0.2s",
      position:"relative", overflow:"hidden"
    }}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:sc,
        opacity:status==="NOMINAL"?0.4:1}}/>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
        <div>
          <div style={{fontSize:10,color:C.textDim,fontFamily:"monospace",letterSpacing:1}}>{m.id} · {m.type?.toUpperCase()}</div>
          <div style={{fontSize:13,color:C.text,fontWeight:600,marginTop:2}}>{m.name}</div>
        </div>
        <div style={{background:`${sc}22`,border:`1px solid ${sc}66`,borderRadius:4,
          padding:"2px 7px",fontSize:10,color:sc,fontFamily:"monospace",letterSpacing:1,fontWeight:700}}>
          {status}
        </div>
      </div>

      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
        <span style={{fontSize:10,color:C.muted,fontFamily:"monospace"}}>RISK</span>
        <div style={{flex:1}}><GaugeBar value={risk} max={100} color={sc}/></div>
        <span style={{fontSize:12,color:sc,fontFamily:"monospace",fontWeight:700,minWidth:32}}>{risk}%</span>
      </div>

      {[
        {label:"TEMP", v:m.avg?.temp, unit:"°C", max:120, warn:75, crit:85},
        {label:"VIB",  v:m.avg?.vib,  unit:"g",  max:15,  warn:5,  crit:6.5},
        {label:"OIL",  v:m.avg?.oil,  unit:"%",  max:100, warn:50, crit:30, inv:true},
      ].map(({label,v,unit,max,warn,crit,inv})=>{
        if (v==null) return null;
        const col = inv
          ? (v<crit?C.danger:v<warn?C.warn:C.safe)
          : (v>crit?C.danger:v>warn?C.warn:C.safe);
        return (
          <div key={label} style={{marginBottom:6}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:C.textDim,marginBottom:3,fontFamily:"monospace"}}>
              <span>{label}</span>
              <span style={{color,fontWeight:700}}>{v}<span style={{color:C.muted}}> {unit}</span></span>
            </div>
            <GaugeBar value={v} max={max} color={col}/>
          </div>
        );
      })}

      <div style={{marginTop:8}}>
        <Sparkline data={m.history?.map(h=>h.temp)} color={sc} w={200} h={28}/>
      </div>
    </div>
  );
}

// ── Alert Row ─────────────────────────────────────────────────────────────────
function AlertBadge({ alert }) {
  const col = STATUS_COLOR[alert.severity] || C.muted;
  return (
    <div style={{display:"flex",gap:12,alignItems:"flex-start",padding:"8px 12px",
      background:`${col}12`,border:`1px solid ${col}33`,borderRadius:6,marginBottom:6}}>
      <span style={{fontSize:11,color:col,fontFamily:"monospace",whiteSpace:"nowrap"}}>
        {alert.severity}
      </span>
      <div>
        <div style={{fontSize:12,color:C.text,fontWeight:600}}>{alert.machineId} — {alert.type}</div>
        <div style={{fontSize:11,color:C.textDim,marginTop:2}}>{alert.param}: <strong style={{color:col}}>{alert.value}</strong> (threshold {alert.threshold})</div>
        <div style={{fontSize:11,color:C.muted,marginTop:3,fontStyle:"italic"}}>{alert.action}</div>
      </div>
    </div>
  );
}

// ── AI Panel ──────────────────────────────────────────────────────────────────
function AnalysisPanel({ machines, alerts, onAnalyze, analysis, analyzing, hasData }) {
  return (
    <div style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,padding:20,display:"flex",flexDirection:"column",gap:14}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontSize:10,color:C.accent,letterSpacing:2,fontFamily:"monospace"}}>AI INFERENCE ENGINE</div>
          <div style={{fontSize:16,color:C.text,fontWeight:700,marginTop:2}}>Predictive Downtime Agent</div>
        </div>
        <button onClick={onAnalyze} disabled={analyzing||!hasData} style={{
          background:analyzing||!hasData?C.border:`linear-gradient(135deg,${C.accent}33,${C.accent}11)`,
          border:`1px solid ${analyzing||!hasData?C.muted:C.accent}`,
          color:analyzing||!hasData?C.muted:C.accent,
          borderRadius:6,padding:"8px 18px",fontFamily:"monospace",fontSize:12,
          cursor:analyzing||!hasData?"not-allowed":"pointer",letterSpacing:1,
          boxShadow:analyzing||!hasData?"":`0 0 12px ${C.accent}33`,transition:"all 0.2s"
        }}>
          {analyzing?"⏳ ANALYZING...":"▶ RUN AI ANALYSIS"}
        </button>
      </div>

      {analyzing && (
        <div style={{display:"flex",gap:8,alignItems:"center",padding:"12px 14px",
          background:`${C.accent}11`,border:`1px solid ${C.accent}33`,borderRadius:6}}>
          <div style={{width:8,height:8,borderRadius:"50%",background:C.accent,
            animation:"pulse 1.2s ease-in-out infinite",boxShadow:`0 0 8px ${C.accent}`}}/>
          <span style={{fontSize:12,color:C.accent,fontFamily:"monospace"}}>Scanning telemetry & alerts from {machines.length} machines…</span>
        </div>
      )}

      {analysis && !analyzing && <MarkdownLike text={analysis}/>}

      {!analysis && !analyzing && (
        <div style={{textAlign:"center",padding:"24px 0",color:C.muted,fontSize:13}}>
          <div style={{fontSize:28,marginBottom:8}}>🔍</div>
          {hasData
            ? "Click Run AI Analysis to get downtime predictions, root causes, and production continuity recommendations."
            : "Upload the IoT Excel file first, then run the AI analysis."}
        </div>
      )}
    </div>
  );
}

function MarkdownLike({ text }) {
  const lines = text.split("\n");
  return (
    <>
      {lines.map((line,i)=>{
        if (line.startsWith("## ")) return (
          <div key={i} style={{fontSize:12,color:C.accent,fontFamily:"monospace",letterSpacing:1,
            fontWeight:700,marginTop:14,marginBottom:5,borderBottom:`1px solid ${C.border}`,paddingBottom:3}}>
            {line.slice(3).toUpperCase()}
          </div>
        );
        if (line.startsWith("### ")) return (
          <div key={i} style={{fontSize:12,color:C.warn,fontFamily:"monospace",fontWeight:700,marginTop:8,marginBottom:3}}>
            ▸ {line.slice(4)}
          </div>
        );
        if (line.startsWith("- ")) return (
          <div key={i} style={{display:"flex",gap:8,marginLeft:8,marginBottom:3}}>
            <span style={{color:C.accent,flexShrink:0}}>›</span>
            <span style={{color:C.text,fontSize:12}}>{line.slice(2)}</span>
          </div>
        );
        if (line.match(/^(🔴|🟠|🟢)/)) {
          const col=line.startsWith("🔴")?C.danger:line.startsWith("🟠")?C.warn:C.safe;
          return (
            <div key={i} style={{background:`${col}15`,border:`1px solid ${col}44`,
              borderRadius:6,padding:"8px 12px",marginBottom:5,fontSize:12,color:col}}>{line}</div>
          );
        }
        if (line.trim()) return <div key={i} style={{fontSize:12,color:C.textDim,marginBottom:3}}>{line}</div>;
        return <div key={i} style={{height:4}}/>;
      })}
    </>
  );
}

// ── Upload Screen ─────────────────────────────────────────────────────────────
function UploadScreen({ onFile }) {
  const ref = useRef();
  const [dragging, setDragging] = useState(false);

  const handle = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => onFile(e.target.result, file.name);
    reader.readAsArrayBuffer(file);
  };

  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",
      height:"100vh",background:C.bg,gap:20,padding:24}}>
      <style>{`@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(1.3)}}`}</style>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:28,fontWeight:800,color:C.text,letterSpacing:-0.5,marginBottom:4}}>
          🏭 ShopFloor <span style={{color:C.accent}}>IQ</span>
        </div>
        <div style={{fontSize:11,color:C.muted,fontFamily:"monospace",letterSpacing:2}}>IOT PREDICTIVE MAINTENANCE AGENT</div>
      </div>

      <div
        onDragOver={e=>{e.preventDefault();setDragging(true);}}
        onDragLeave={()=>setDragging(false)}
        onDrop={e=>{e.preventDefault();setDragging(false);handle(e.dataTransfer.files[0]);}}
        onClick={()=>ref.current.click()}
        style={{
          border:`2px dashed ${dragging?C.accent:C.border}`,borderRadius:16,
          padding:"48px 60px",textAlign:"center",cursor:"pointer",
          background:dragging?`${C.accent}08`:C.panel,transition:"all 0.2s",
          boxShadow:dragging?`0 0 24px ${C.accent}33`:"none",maxWidth:460,width:"100%"
        }}>
        <input ref={ref} type="file" accept=".xlsx,.xls" style={{display:"none"}}
          onChange={e=>handle(e.target.files[0])}/>
        <div style={{fontSize:40,marginBottom:12}}>📊</div>
        <div style={{fontSize:16,color:C.text,fontWeight:600,marginBottom:8}}>
          Drop IoT Excel file here
        </div>
        <div style={{fontSize:12,color:C.textDim,marginBottom:20}}>
          or click to browse — supports .xlsx format
        </div>
        <div style={{background:`${C.accent}22`,border:`1px solid ${C.accent}44`,
          borderRadius:6,padding:"8px 20px",display:"inline-block",
          color:C.accent,fontSize:12,fontFamily:"monospace",fontWeight:700}}>
          SELECT FILE
        </div>
      </div>

      <div style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,
        padding:16,maxWidth:460,width:"100%"}}>
        <div style={{fontSize:10,color:C.accent,fontFamily:"monospace",letterSpacing:1,marginBottom:10}}>EXPECTED SHEETS IN EXCEL</div>
        {[
          ["📋 Live Sensor Log","Timestamped readings: Temp, Vibration, Pressure, RPM, Current, Oil"],
          ["📊 Machine Summary","Aggregated avg/max per machine with pre-computed risk"],
          ["🚨 Alert Log","Threshold breach events with severity & recommended actions"],
        ].map(([title,desc])=>(
          <div key={title} style={{marginBottom:8}}>
            <div style={{fontSize:12,color:C.text,fontWeight:600}}>{title}</div>
            <div style={{fontSize:11,color:C.textDim}}>{desc}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [data, setData] = useState(null);
  const [selected, setSelected] = useState(null);
  const [analysis, setAnalysis] = useState("");
  const [analyzing, setAnalyzing] = useState(false);

  const onFile = useCallback((buffer, fileName) => {
    try {
      const wb = XLSX.read(buffer, { type:"array", cellDates:true });

      // Parse Machine Summary sheet
      const summarySheet = wb.Sheets["Machine Summary"];
      const summaryRows = summarySheet ? XLSX.utils.sheet_to_json(summarySheet) : [];

      // Parse Live Sensor Log
      const logSheet = wb.Sheets["Live Sensor Log"];
      const logRows = logSheet ? XLSX.utils.sheet_to_json(logSheet) : [];

      // Parse Alert Log
      const alertSheet = wb.Sheets["Alert Log"];
      const alertRows = alertSheet ? XLSX.utils.sheet_to_json(alertSheet) : [];

      // Build machine objects
      const machineMap = {};
      for (const row of summaryRows) {
        const id = row["Machine ID"];
        if (!id) continue;
        machineMap[id] = {
          id,
          name: row["Machine Name"] || id,
          type: row["Type"] || "",
          avg: {
            temp: row["Avg Temp"] || 0,
            maxTemp: row["Max Temp"] || 0,
            vib: row["Avg Vibration"] || 0,
            maxVib: row["Max Vibration"] || 0,
            pres: row["Avg Pressure"] || 0,
            rpm: row["Avg RPM"] || 0,
            curr: row["Avg Current"] || 0,
            oil: row["Avg Oil Level"] || 0,
          },
          risk: row["Risk Score"] || 0,
          status: row["Status"] || "NOMINAL",
          history: [],
        };
      }

      // Attach history (last 20 readings per machine)
      for (const row of logRows) {
        const id = row["Machine ID"];
        if (!machineMap[id]) continue;
        const hist = machineMap[id].history;
        if (hist.length < 20) {
          hist.push({
            temp: row["Temperature (°C)"] || 0,
            vib: row["Vibration (g)"] || 0,
            oil: row["Oil Level (%)"] || 0,
          });
        }
      }

      // Parse alerts
      const alerts = alertRows.map(r => ({
        ts: r["Timestamp"],
        machineId: r["Machine ID"],
        machineName: r["Machine Name"],
        type: r["Alert Type"],
        param: r["Parameter"],
        value: r["Value"],
        threshold: r["Threshold"],
        severity: r["Severity"],
        action: r["Recommended Action"],
      }));

      setData({ machines: Object.values(machineMap), alerts, fileName, totalReadings: logRows.length });
      setAnalysis("");
    } catch(err) {
      alert("Could not parse Excel file. Make sure it follows the expected format.");
    }
  }, []);

  const runAnalysis = useCallback(async () => {
    if (!data) return;
    setAnalyzing(true); setAnalysis("");

    const machineSnapshot = data.machines.map(m => ({
      ...m, risk: riskScore(m.avg), status: statusFromRisk(riskScore(m.avg)), history: undefined
    }));
    const criticalAlerts = data.alerts.filter(a => a.severity === "CRITICAL").slice(0,6);

    const prompt = `You are an industrial IoT Predictive Maintenance AI Agent on a manufacturing shopfloor.

MACHINE TELEMETRY SUMMARY (from uploaded Excel data — ${data.totalReadings} sensor readings):
${JSON.stringify(machineSnapshot, null, 2)}

RECENT CRITICAL ALERTS:
${JSON.stringify(criticalAlerts, null, 2)}

Produce a structured operational report with these sections (use ## for sections, ### for subsections, - for bullets):

## Machine Health Summary
For each machine: one line with 🔴/🟠/🟢 + machine name + risk% + top concern.

## Downtime Risk Forecast
Top 2–3 machines at highest unplanned downtime risk. Estimate probability and estimated time-to-failure window based on sensor trends.

## Root Cause Analysis
For at-risk machines, explain the sensor pattern driving the risk (e.g., rising temp + low oil = bearing overheating).

## Immediate Actions Required
Concrete, prioritized actions for NOW — specific to which machine, what parameter, what action, by whom.

## Production Continuity Plan
How to redistribute work or sequence production runs to maintain throughput if top-risk machines go offline.

## 24-Hour Maintenance Schedule
Proposed maintenance interventions: machine, time window, action, estimated duration.

Be specific and actionable. Plant operator language, not engineering jargon.`;

    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body: JSON.stringify({ model:"claude-sonnet-4-6", max_tokens:1000, messages:[{role:"user",content:prompt}] })
      });
      const d = await res.json();
      const text = d.content?.map(b=>b.type==="text"?b.text:"").join("") || "No response.";
      setAnalysis(text);
    } catch(e) {
      setAnalysis("## Error\nFailed to reach AI engine. Please try again.");
    } finally { setAnalyzing(false); }
  }, [data]);

  if (!data) return <UploadScreen onFile={onFile}/>;

  const { machines, alerts, fileName, totalReadings } = data;
  const critical = machines.filter(m=>statusFromRisk(riskScore(m.avg))==="CRITICAL").length;
  const warning  = machines.filter(m=>statusFromRisk(riskScore(m.avg))==="WARNING").length;
  const selMachine = machines.find(m=>m.id===selected);

  return (
    <div style={{background:C.bg,minHeight:"100vh",fontFamily:"system-ui,sans-serif",color:C.text}}>
      <style>{`@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(1.3)}} *{box-sizing:border-box;margin:0;padding:0} ::-webkit-scrollbar{width:4px;height:4px} ::-webkit-scrollbar-thumb{background:${C.border};border-radius:2px}`}</style>

      {/* Header */}
      <div style={{borderBottom:`1px solid ${C.border}`,padding:"11px 20px",
        display:"flex",alignItems:"center",justifyContent:"space-between",background:C.panel}}>
        <div>
          <div style={{fontSize:17,fontWeight:800,color:C.text,letterSpacing:-0.5}}>🏭 ShopFloor <span style={{color:C.accent}}>IQ</span></div>
          <div style={{fontSize:9,color:C.muted,fontFamily:"monospace",letterSpacing:1}}>IOT PREDICTIVE MAINTENANCE AGENT</div>
        </div>
        <div style={{display:"flex",gap:16,alignItems:"center"}}>
          {[["MACHINES",machines.length,C.accent],["CRITICAL",critical,C.danger],["WARNING",warning,C.warn],["NOMINAL",machines.length-critical-warning,C.safe]].map(([l,v,c])=>(
            <div key={l} style={{textAlign:"center"}}>
              <div style={{fontSize:18,fontWeight:800,color:c,lineHeight:1}}>{v}</div>
              <div style={{fontSize:9,color:C.muted,fontFamily:"monospace",letterSpacing:1}}>{l}</div>
            </div>
          ))}
          <div style={{borderLeft:`1px solid ${C.border}`,paddingLeft:16}}>
            <div style={{fontSize:10,color:C.textDim,fontFamily:"monospace"}}>{fileName}</div>
            <div style={{fontSize:9,color:C.muted,fontFamily:"monospace"}}>{totalReadings} sensor readings</div>
          </div>
          <button onClick={()=>{setData(null);setAnalysis("");}} style={{
            background:"transparent",border:`1px solid ${C.border}`,borderRadius:5,
            padding:"4px 10px",color:C.muted,fontSize:11,cursor:"pointer",fontFamily:"monospace"
          }}>↺ RELOAD</button>
        </div>
      </div>

      {/* Alert ticker */}
      {alerts.filter(a=>a.severity==="CRITICAL").length > 0 && (
        <div style={{background:`${C.danger}12`,borderBottom:`1px solid ${C.danger}33`,
          padding:"7px 20px",display:"flex",gap:12,overflowX:"auto",alignItems:"center"}}>
          <span style={{fontSize:10,color:C.danger,fontFamily:"monospace",fontWeight:700,flexShrink:0,letterSpacing:1}}>⚠ CRITICAL</span>
          {alerts.filter(a=>a.severity==="CRITICAL").map((a,i)=>(
            <span key={i} style={{fontSize:11,color:C.danger,fontFamily:"monospace",whiteSpace:"nowrap",
              background:`${C.danger}18`,padding:"2px 8px",borderRadius:4}}>
              {a.machineId}: {a.type} ({a.param} = {a.value})
            </span>
          ))}
        </div>
      )}

      {/* Main */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 340px",height:`calc(100vh - ${alerts.filter(a=>a.severity==="CRITICAL").length>0?"90":"60"}px)`}}>

        {/* Left */}
        <div style={{overflowY:"auto",padding:16,display:"flex",flexDirection:"column",gap:14}}>
          {/* Machine grid */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
            {machines.map(m=>(
              <MachineCard key={m.id} m={m} selected={selected}
                onClick={()=>setSelected(s=>s===m.id?null:m.id)}/>
            ))}
          </div>
          {/* AI Panel */}
          <AnalysisPanel machines={machines} alerts={alerts}
            onAnalyze={runAnalysis} analysis={analysis} analyzing={analyzing} hasData={!!data}/>
        </div>

        {/* Right sidebar */}
        <div style={{borderLeft:`1px solid ${C.border}`,overflowY:"auto",padding:14,display:"flex",flexDirection:"column",gap:12}}>
          {/* Selected machine detail */}
          {selMachine ? (
            <div style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,padding:14}}>
              <div style={{fontSize:10,color:C.accent,fontFamily:"monospace",letterSpacing:1,marginBottom:8}}>MACHINE DETAIL</div>
              <div style={{fontSize:14,fontWeight:700,color:C.text,marginBottom:12}}>{selMachine.name}</div>
              {[
                {label:"Avg Temp",v:selMachine.avg?.temp,unit:"°C",max:120},
                {label:"Max Temp",v:selMachine.avg?.maxTemp,unit:"°C",max:120},
                {label:"Avg Vibration",v:selMachine.avg?.vib,unit:"g",max:15},
                {label:"Avg Pressure",v:selMachine.avg?.pres,unit:"PSI",max:200},
                {label:"Avg RPM",v:selMachine.avg?.rpm,unit:"",max:4000},
                {label:"Avg Current",v:selMachine.avg?.curr,unit:"A",max:30},
                {label:"Oil Level",v:selMachine.avg?.oil,unit:"%",max:100},
              ].map(({label,v,unit,max})=>v!=null&&(
                <div key={label} style={{marginBottom:8}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:C.textDim,marginBottom:3,fontFamily:"monospace"}}>
                    <span>{label}</span><span style={{color:C.accent,fontWeight:700}}>{v} {unit}</span>
                  </div>
                  <GaugeBar value={v} max={max} color={C.accent}/>
                </div>
              ))}
              <div style={{marginTop:10,borderTop:`1px solid ${C.border}`,paddingTop:10}}>
                <div style={{fontSize:10,color:C.textDim,fontFamily:"monospace",marginBottom:6}}>TEMP TREND (LAST 20 READINGS)</div>
                <Sparkline data={selMachine.history?.map(h=>h.temp)} color={STATUS_COLOR[statusFromRisk(riskScore(selMachine.avg))]} w={280} h={40}/>
              </div>
            </div>
          ) : (
            <div style={{textAlign:"center",padding:"32px 12px",color:C.muted,fontSize:12}}>
              <div style={{fontSize:26,marginBottom:8}}>👆</div>
              Click a machine card to inspect its sensor detail
            </div>
          )}

          {/* Recent alerts */}
          <div style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,padding:14}}>
            <div style={{fontSize:10,color:C.accent,fontFamily:"monospace",letterSpacing:1,marginBottom:10}}>RECENT ALERTS ({alerts.length})</div>
            <div style={{maxHeight:280,overflowY:"auto"}}>
              {alerts.slice(0,8).map((a,i)=><AlertBadge key={i} alert={a}/>)}
            </div>
          </div>

          {/* Fleet overview */}
          <div style={{background:C.panel,border:`1px solid ${C.border}`,borderRadius:10,padding:14}}>
            <div style={{fontSize:10,color:C.accent,fontFamily:"monospace",letterSpacing:1,marginBottom:10}}>FLEET RISK OVERVIEW</div>
            {machines.map(m=>{
              const r=riskScore(m.avg);
              const sc=STATUS_COLOR[statusFromRisk(r)];
              return (
                <div key={m.id} style={{display:"flex",alignItems:"center",gap:8,marginBottom:7}}>
                  <span style={{fontSize:11,color:C.textDim,fontFamily:"monospace",minWidth:38}}>{m.id}</span>
                  <div style={{flex:1}}><GaugeBar value={r} max={100} color={sc}/></div>
                  <span style={{fontSize:11,color:sc,fontFamily:"monospace",minWidth:30}}>{r}%</span>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}
