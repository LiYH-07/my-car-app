import { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, ComposedChart, Area
} from "recharts";

// ─── Theme ────────────────────────────────────────────────────────────────────
const C = {
  bg: "#0d1117", surface: "#161b27", card: "#1e2435",
  border: "#2a3050", accent: "#f5a623", accent2: "#38bdf8",
  green: "#22d3a5", red: "#f87171", purple: "#a78bfa",
  text: "#dde3f5", muted: "#5a637a",
};

const GS = `
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
body{background:${C.bg};color:${C.text};font-family:'Syne',sans-serif;-webkit-font-smoothing:antialiased;}
::-webkit-scrollbar{width:4px;height:4px;}
::-webkit-scrollbar-track{background:${C.surface};}
::-webkit-scrollbar-thumb{background:${C.border};border-radius:2px;}
input,select,textarea{background:${C.bg};color:${C.text};border:1px solid ${C.border};border-radius:8px;
  padding:10px 13px;font-family:'Syne',sans-serif;font-size:14px;outline:none;width:100%;transition:border-color .2s;}
input:focus,select:focus,textarea:focus{border-color:${C.accent};}
input[type=range]{padding:2px 0;border:none;background:transparent;accent-color:${C.accent};cursor:pointer;}
select option{background:${C.card};}
label{font-size:11px;color:${C.muted};letter-spacing:.07em;text-transform:uppercase;display:block;margin-bottom:5px;font-weight:600;}
`;

// ─── Google Apps Script API ────────────────────────────────────────────────────
async function gasGet(url) {
  const r = await fetch(url + "?action=getAll");
  if (!r.ok) throw new Error("網路錯誤 " + r.status);
  return r.json();
}
async function gasPost(url, action, data) {
  const r = await fetch(url, {
    method: "POST", redirect: "follow",
    headers: { "Content-Type": "text/plain" },
    body: JSON.stringify({ action, data }),
  });
  if (!r.ok) throw new Error("網路錯誤 " + r.status);
  return r.json();
}

// ─── Local fallback storage ────────────────────────────────────────────────────
async function lsGet(k) {
  try { const r = await window.storage.get(k); return r ? JSON.parse(r.value) : null; } catch { return null; }
}
async function lsSave(k, v) {
  try { await window.storage.set(k, JSON.stringify(v)); } catch {}
}

// ─── Excel parser ──────────────────────────────────────────────────────────────
function parseDate(v) {
  const s = String(v || "");
  
  // 處理 Excel 序列號（數字）
  if (/^\d+(\.\d+)?$/.test(s)) {
    const excelDate = parseFloat(s);
    // Excel 序列號轉換：1900-01-01 是 1，1899-12-30 是 0
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    return date.toISOString().split('T')[0]; // 返回 YYYY-MM-DD
  }
  
  // 處理字符串日期格式
  const m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  }
  
  // 最後的備選方案
  return s.slice(0, 10);
}
function parseTitle(t) {
  const m = String(t||"").match(/([\d.]+)\s*公升[,，]?\s*(.*)/);
  return m ? { type:"fuel", liters:+m[1], fuelType:m[2].trim()||"98無鉛" }
           : { type:"maintenance", liters:0, fuelType:"" };
}
function importExcel(file, existing) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        // 關鍵：添加 raw:true 保留原始值
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
        const existKeys = new Set(existing.map(r => r.date+"_"+r.odometer));
        const out = [];
        for (let i = 1; i < rows.length; i++) {
          const [dateRaw, odo, amt, title, note] = rows[i];
          if (!dateRaw) continue;
          const date = parseDate(dateRaw);
          if (existKeys.has(date+"_"+odo)) continue;
          const { type, liters, fuelType } = parseTitle(title);
          const amount = parseFloat(String(amt||0).replace(/,/g,""))||0;
          out.push({
            id: Date.now() + i, date, time:"",
            odometer: +odo||0, amount, type, liters,
            fuelType, pricePerLiter: liters>0 ? +(amount/liters).toFixed(1) : 0,
            title: String(title||""), note: String(note||""), tankPct: 100,
          });
        }
        res(out);
      } catch(err) { rej(err); }
    };
    reader.readAsBinaryString(file);
  });
}

// ─── CPC Oil Price via Anthropic API ──────────────────────────────────────────
async function fetchCPCPrices() {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 500,
      tools: [{ type: "web_search_20250305", name: "web_search" }],
      messages: [{
        role: "user",
        content: `請搜尋台灣中油今日最新零售油價，然後只回傳這個JSON格式（不要加markdown）：
{"98":數字,"95":數字,"92":數字,"diesel":數字,"updateDate":"日期"}
數字是新台幣每公升價格，日期格式YYYY-MM-DD。`
      }]
    })
  });
  const data = await res.json();
  const text = data.content?.map(b => b.text||"").join("") || "";
  const match = text.match(/\{[^}]*"98"[^}]*\}/s);
  if (!match) throw new Error("無法解析油價");
  return JSON.parse(match[0]);
}

// ─── UI Primitives ─────────────────────────────────────────────────────────────
const Btn = ({ onClick, children, variant="primary", style={}, disabled, loading }) => {
  const bg = variant==="primary"?C.accent : variant==="danger"?C.red : variant==="ghost"?"transparent":C.card;
  const col = variant==="primary"?"#000" : variant==="ghost"?C.muted : C.text;
  return (
    <button onClick={onClick} disabled={disabled||loading} style={{
      background:bg, color:col, border: variant==="ghost"?`1px solid ${C.border}`:"none",
      borderRadius:10, padding:"11px 18px", fontFamily:"Syne", fontWeight:700,
      fontSize:13, cursor:(disabled||loading)?"not-allowed":"pointer",
      opacity:(disabled||loading)?.6:1, transition:"opacity .2s,transform .1s",
      display:"flex", alignItems:"center", justifyContent:"center", gap:6, ...style,
    }}
    onMouseDown={e=>{if(!disabled&&!loading)e.currentTarget.style.transform="scale(.97)"}}
    onMouseUp={e=>e.currentTarget.style.transform="scale(1)"}
    >{loading?"⏳":children}</button>
  );
};

const Card = ({ children, style={} }) => (
  <div style={{
    background:C.card, border:`1px solid ${C.border}`,
    borderRadius:16, padding:18, ...style
  }}>{children}</div>
);

const Field = ({ label, children, span }) => (
  <div style={{ marginBottom:14, gridColumn:span?`span ${span}`:undefined }}>
    <label>{label}</label>
    {children}
  </div>
);

const Badge = ({ color, children, small }) => (
  <span style={{
    background:color+"22", color, border:`1px solid ${color}44`,
    borderRadius:6, padding: small?"1px 7px":"3px 10px",
    fontSize: small?11:12, fontWeight:700,
  }}>{children}</span>
);

const Divider = ({ label }) => (
  <div style={{ display:"flex", alignItems:"center", gap:10, margin:"16px 0" }}>
    <div style={{ flex:1, height:1, background:C.border }} />
    {label && <span style={{ color:C.muted, fontSize:11, fontWeight:600, letterSpacing:".06em" }}>{label}</span>}
    <div style={{ flex:1, height:1, background:C.border }} />
  </div>
);

// ─── Setup Page ────────────────────────────────────────────────────────────────
function SetupPage({ onDone }) {
  const [url, setUrl] = useState("");
  const [testing, setTesting] = useState(false);
  const [status, setStatus] = useState(null);
  const [step, setStep] = useState(0);

  const steps = [
    { icon:"📊", title:"建立 Google Sheets", desc:"在 Google Drive 新建一個試算表，名稱自訂即可（例如：我的車輛記錄）" },
    { icon:"⚙️", title:"開啟 Apps Script", desc:"在試算表上方選單點「擴充功能」→「Apps Script」，會開啟新頁面" },
    { icon:"📋", title:"貼上程式碼", desc:"刪除預設內容，將下載的「apps-script.js」檔案內容全部貼入，按儲存（Ctrl+S）" },
    { icon:"🚀", title:"部署為網頁應用程式", desc:"點「部署」→「新增部署」→類型選「網頁應用程式」→執行身分：「我」→存取權限：「所有人」→按「部署」" },
    { icon:"🔗", title:"複製網址", desc:"複製「網頁應用程式網址」（格式為 https://script.google.com/macros/s/.../exec），貼到下方" },
  ];

  const test = async () => {
    if (!url.includes("script.google.com")) { setStatus({ ok:false, msg:"網址格式不正確" }); return; }
    setTesting(true); setStatus(null);
    try {
      const r = await gasGet(url);
      if (r.ok !== false) { setStatus({ ok:true, msg:"連線成功！" }); }
      else { setStatus({ ok:false, msg:"Apps Script 回傳錯誤: " + r.error }); }
    } catch(e) { setStatus({ ok:false, msg:"連線失敗：" + e.message }); }
    setTesting(false);
  };

  return (
    <div style={{ padding:"24px 16px", maxWidth:480, margin:"0 auto" }}>
      <div style={{ textAlign:"center", marginBottom:28 }}>
        <div style={{ fontSize:52, marginBottom:12 }}>🔧</div>
        <h1 style={{ fontSize:24, fontWeight:800, marginBottom:6 }}>初次設定</h1>
        <p style={{ color:C.muted, fontSize:13, lineHeight:1.7 }}>
          需要設定 Google Sheets 來儲存你的車輛資料<br/>資料存在你自己的 Google Drive，完全私密
        </p>
      </div>

      {/* Steps */}
      <div style={{ marginBottom:24 }}>
        {steps.map((s,i) => (
          <div key={i} onClick={()=>setStep(i)} style={{
            display:"flex", gap:14, padding:"14px 16px", borderRadius:12, marginBottom:8,
            background: step===i ? C.card : "transparent",
            border:`1px solid ${step===i ? C.border : "transparent"}`,
            cursor:"pointer", transition:"all .2s"
          }}>
            <div style={{
              width:36, height:36, borderRadius:10, flexShrink:0,
              background: i<step ? C.green+"22" : step===i ? C.accent+"22" : C.surface,
              border:`1px solid ${i<step ? C.green : step===i ? C.accent : C.border}`,
              display:"flex", alignItems:"center", justifyContent:"center",
              fontSize:16
            }}>{i<step ? "✓" : s.icon}</div>
            <div>
              <p style={{ fontWeight:700, fontSize:14, marginBottom:3 }}>步驟 {i+1}：{s.title}</p>
              {step===i && <p style={{ color:C.muted, fontSize:12, lineHeight:1.6 }}>{s.desc}</p>}
            </div>
          </div>
        ))}
      </div>

      <Card>
        <p style={{ color:C.muted, fontSize:12, marginBottom:10 }}>完成上述步驟後，貼上 Apps Script 網址：</p>
        <input
          value={url} onChange={e=>setUrl(e.target.value)}
          placeholder="https://script.google.com/macros/s/.../exec"
          style={{ fontFamily:"JetBrains Mono", fontSize:12, marginBottom:12 }}
        />
        {status && (
          <div style={{
            padding:"10px 14px", borderRadius:8, marginBottom:12, fontSize:13, fontWeight:600,
            background: status.ok ? C.green+"22" : C.red+"22",
            border:`1px solid ${status.ok ? C.green : C.red}`,
            color: status.ok ? C.green : C.red,
          }}>{status.ok ? "✅ " : "❌ "}{status.msg}</div>
        )}
        <div style={{ display:"flex", gap:10 }}>
          <Btn onClick={test} variant="ghost" loading={testing} style={{ flex:1 }}>測試連線</Btn>
          <Btn onClick={()=>onDone(url)} disabled={!status?.ok} style={{ flex:1 }}>
            開始使用 →
          </Btn>
        </div>
        <Divider label="或" />
        <Btn onClick={()=>onDone("")} variant="ghost" style={{ width:"100%", fontSize:12 }}>
          暫時略過，使用本機儲存（換裝置資料會消失）
        </Btn>
      </Card>
    </div>
  );
}

// ─── Oil Price Widget ──────────────────────────────────────────────────────────
function OilPriceBar() {
  const [prices, setPrices] = useState(null);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState(null);
  const [expanded, setExpanded] = useState(false);

  const fetch_ = useCallback(async () => {
    setLoading(true); setErr(null);
    try {
      const p = await fetchCPCPrices();
      setPrices(p);
    } catch(e) { setErr("無法取得"); }
    setLoading(false);
  }, []);

  useEffect(() => { fetch_(); }, [fetch_]);

  if (!expanded) return (
    <div onClick={()=>setExpanded(true)} style={{
      background: `linear-gradient(90deg, ${C.accent}15, ${C.surface})`,
      border:`1px solid ${C.accent}33`, borderRadius:10, padding:"10px 14px",
      display:"flex", alignItems:"center", gap:10, cursor:"pointer", marginBottom:16,
    }}>
      <span style={{ fontSize:18 }}>⛽</span>
      <div style={{ flex:1 }}>
        <p style={{ fontSize:11, color:C.muted, marginBottom:2 }}>台灣中油今日油價</p>
        {loading && <p style={{ fontSize:13, color:C.accent }}>載入中...</p>}
        {err && <p style={{ fontSize:13, color:C.red }}>{err}</p>}
        {prices && !loading && (
          <div style={{ display:"flex", gap:12 }}>
            <span style={{ fontSize:13, fontWeight:700, color:C.accent, fontFamily:"JetBrains Mono" }}>
              98無鉛 ${prices["98"]}
            </span>
            <span style={{ fontSize:13, fontWeight:700, color:C.text, fontFamily:"JetBrains Mono" }}>
              95無鉛 ${prices["95"]}
            </span>
          </div>
        )}
      </div>
      <span style={{ color:C.muted, fontSize:12 }}>▼</span>
    </div>
  );

  return (
    <Card style={{ marginBottom:16, background:`linear-gradient(135deg,${C.card},${C.surface})` }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          <span style={{ fontSize:20 }}>⛽</span>
          <span style={{ fontWeight:700 }}>台灣中油即時油價</span>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          <button onClick={fetch_} style={{
            background:"transparent",border:`1px solid ${C.border}`,borderRadius:6,
            padding:"4px 10px",color:C.muted,fontSize:11,cursor:"pointer",fontFamily:"Syne"
          }}>🔄 更新</button>
          <button onClick={()=>setExpanded(false)} style={{
            background:"transparent",border:"none",color:C.muted,cursor:"pointer",fontSize:16
          }}>✕</button>
        </div>
      </div>
      {loading && <p style={{ color:C.muted, fontSize:13, textAlign:"center", padding:"12px 0" }}>搜尋中油最新油價...</p>}
      {err && <p style={{ color:C.red, fontSize:13 }}>{err} <button onClick={fetch_} style={{ color:C.accent,background:"none",border:"none",cursor:"pointer",fontSize:13 }}>重試</button></p>}
      {prices && !loading && (
        <>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:8, marginBottom:10 }}>
            {[
              ["98無鉛", prices["98"], C.accent],
              ["95無鉛", prices["95"], C.accent2],
              ["92無鉛", prices["92"], C.green],
              ["柴油", prices.diesel, C.purple],
            ].map(([label, price, color]) => (
              <div key={label} style={{
                background:C.bg, borderRadius:10, padding:"10px 6px", textAlign:"center",
                border:`1px solid ${color}33`
              }}>
                <p style={{ fontSize:10, color:C.muted, marginBottom:4 }}>{label}</p>
                <p style={{ fontSize:18, fontWeight:800, color, fontFamily:"JetBrains Mono" }}>{price}</p>
                <p style={{ fontSize:9, color:C.muted }}>$/L</p>
              </div>
            ))}
          </div>
          {prices.updateDate && (
            <p style={{ fontSize:11, color:C.muted, textAlign:"center" }}>更新日期：{prices.updateDate}</p>
          )}
        </>
      )}
    </Card>
  );
}

// ─── Car Setup ─────────────────────────────────────────────────────────────────
function CarSetup({ car, gasUrl, onSave }) {
  const [form, setForm] = useState(car || {
    plate:"AUT-6752", nickname:"", brand:"Volkswagen", model:"2016 Passat 330 TSI",
    type:"轎車", fuelCategory:"汽油", cc:1798, tankSize:66,
    reserveSize:8.6, warningPct:10, hp:180, odometer:83892,
  });
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState("");
  const set = (k,v) => setForm(f=>({...f,[k]:v}));

  const save = async () => {
    setSaving(true);
    try {
      if (gasUrl) await gasPost(gasUrl, "saveCar", form);
      else await lsSave("car_info", form);
      onSave(form);
      setToast("✅ 已儲存");
    } catch(e) { setToast("❌ " + e.message); }
    setSaving(false);
    setTimeout(()=>setToast(""),2500);
  };

  return (
    <div>
      <h2 style={{ fontSize:22, fontWeight:800, marginBottom:20 }}>🚗 車輛設定</h2>
      <OilPriceBar />
      <Card style={{ marginBottom:14 }}>
        <p style={{ color:C.muted, fontSize:11, fontWeight:600, marginBottom:14, letterSpacing:".06em" }}>基本資料</p>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
          <Field label="車牌號碼"><input value={form.plate} onChange={e=>set("plate",e.target.value)} /></Field>
          <Field label="愛車暱稱"><input value={form.nickname} onChange={e=>set("nickname",e.target.value)} placeholder="選填" /></Field>
          <Field label="品牌"><input value={form.brand} onChange={e=>set("brand",e.target.value)} /></Field>
          <Field label="型號"><input value={form.model} onChange={e=>set("model",e.target.value)} /></Field>
          <Field label="類型">
            <select value={form.type} onChange={e=>set("type",e.target.value)}>
              {["轎車","SUV","MPV","貨車","休旅車","機車"].map(v=><option key={v}>{v}</option>)}
            </select>
          </Field>
          <Field label="燃料種類">
            <select value={form.fuelCategory} onChange={e=>set("fuelCategory",e.target.value)}>
              {["汽油","柴油","電動","混合動力"].map(v=><option key={v}>{v}</option>)}
            </select>
          </Field>
          <Field label="排氣量 (cc)"><input type="number" value={form.cc} onChange={e=>set("cc",+e.target.value)} /></Field>
          <Field label="馬力 (hp)"><input type="number" value={form.hp} onChange={e=>set("hp",+e.target.value)} /></Field>
        </div>
      </Card>
      <Card style={{ marginBottom:14 }}>
        <p style={{ color:C.muted, fontSize:11, fontWeight:600, marginBottom:14, letterSpacing:".06em" }}>油箱設定</p>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
          <Field label="目前里程 (km)"><input type="number" value={form.odometer} onChange={e=>set("odometer",+e.target.value)} /></Field>
          <Field label="油箱容量 (L)"><input type="number" value={form.tankSize} onChange={e=>set("tankSize",+e.target.value)} /></Field>
          <Field label="備用油量 (L)"><input type="number" value={form.reserveSize} onChange={e=>set("reserveSize",+e.target.value)} /></Field>
          <Field label={`警示燈百分比: ${form.warningPct}%`}>
            <input type="range" min={5} max={30} value={form.warningPct} onChange={e=>set("warningPct",+e.target.value)} />
          </Field>
        </div>
      </Card>
      {toast && <div style={{
        padding:"10px 14px", borderRadius:10, marginBottom:12, fontSize:13, fontWeight:600, textAlign:"center",
        background:toast.startsWith("✅")?C.green+"22":C.red+"22",
        border:`1px solid ${toast.startsWith("✅")?C.green:C.red}`,
        color:toast.startsWith("✅")?C.green:C.red
      }}>{toast}</div>}
      <Btn onClick={save} loading={saving} style={{ width:"100%" }}>💾 儲存車輛設定</Btn>
      {!gasUrl && <p style={{ textAlign:"center", color:C.muted, fontSize:11, marginTop:8 }}>⚠ 目前使用本機儲存，資料不會跨裝置同步</p>}
    </div>
  );
}

// ─── Add Record ────────────────────────────────────────────────────────────────
function AddRecord({ car, records, gasUrl, onAdd }) {
  const sorted = [...records].sort((a,b)=>b.odometer-a.odometer);
  const lastFuel = sorted.find(r=>r.type==="fuel");
  const today = new Date().toISOString().slice(0,10);
  const now = new Date().toTimeString().slice(0,5);
  const [form, setForm] = useState({
    date:today, time:now, odometer:car?.odometer||"", amount:"", type:"fuel",
    liters:"", pricePerLiter:"", fuelType:"98無鉛", title:"", note:"", tankPct:100,
  });
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState(null);
  const set = (k,v) => setForm(f=>({...f,[k]:v}));

  const driven = form.odometer && lastFuel?.odometer ? +form.odometer - lastFuel.odometer : null;
  const fuelEff = driven>0 && +form.liters>0 ? (driven/+form.liters).toFixed(2) : null;
  const estCost = driven && car?.odometer ? null : null;

  const submit = async () => {
    if (!form.date || !form.odometer || !form.amount) {
      setToast({ ok:false, msg:"請填寫日期、里程數及金額" }); return;
    }
    setSaving(true);
    const rec = {
      id: Date.now(), date:form.date, time:form.time,
      odometer:+form.odometer, amount:+form.amount,
      type:form.type, liters:+form.liters||0,
      pricePerLiter:+form.pricePerLiter||0, fuelType:form.fuelType,
      title:form.type==="fuel"?`${form.liters}公升,${form.fuelType}`:form.title,
      note:form.note, tankPct:form.tankPct,
    };
    try {
      if (gasUrl) await gasPost(gasUrl, "addRecords", [rec]);
      else await lsSave("records", [...records, rec]);
      onAdd(rec);
      setToast({ ok:true, msg:"✅ 記錄已新增至 Google Sheets" });
      setForm(f=>({...f, amount:"", liters:"", pricePerLiter:"", note:"", odometer:+form.odometer}));
    } catch(e) { setToast({ ok:false, msg:"❌ " + e.message }); }
    setSaving(false);
    setTimeout(()=>setToast(null), 3000);
  };

  const typeConf = { fuel:["⛽ 加油",C.accent], maintenance:["🔧 保養/維修",C.green], other:["📋 其他費用",C.purple] };

  return (
    <div>
      <h2 style={{ fontSize:22, fontWeight:800, marginBottom:20 }}>＋ 新增記錄</h2>
      <OilPriceBar />

      <div style={{ display:"flex", gap:8, marginBottom:16 }}>
        {Object.entries(typeConf).map(([v,[l,col]])=>(
          <button key={v} onClick={()=>set("type",v)} style={{
            flex:1, padding:"10px 0", borderRadius:10,
            background:form.type===v?col+"22":C.card,
            color:form.type===v?col:C.muted,
            border:`1px solid ${form.type===v?col:C.border}`,
            fontFamily:"Syne", fontWeight:700, fontSize:12, cursor:"pointer", transition:"all .2s"
          }}>{l}</button>
        ))}
      </div>

      <Card style={{ marginBottom:14 }}>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
          <Field label="日期"><input type="date" value={form.date} onChange={e=>set("date",e.target.value)} /></Field>
          <Field label="時間"><input type="time" value={form.time} onChange={e=>set("time",e.target.value)} /></Field>
          <Field label="里程數 (km)" span={2}>
            <input type="number" value={form.odometer} onChange={e=>set("odometer",e.target.value)} placeholder={lastFuel ? `上次：${lastFuel.odometer?.toLocaleString()}` : ""} />
          </Field>
          <Field label="總金額 ($)" span={2}>
            <input type="number" value={form.amount} onChange={e=>set("amount",e.target.value)} placeholder="輸入費用" />
          </Field>
        </div>

        {form.type==="fuel" && <>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:10, marginTop:4 }}>
            <Field label="加油量 (L)">
              <input type="number" value={form.liters} onChange={e=>{
                set("liters",e.target.value);
                if(form.amount&&+e.target.value>0) set("pricePerLiter",(+form.amount/+e.target.value).toFixed(1));
              }} />
            </Field>
            <Field label="油價 ($/L)">
              <input type="number" value={form.pricePerLiter} onChange={e=>set("pricePerLiter",e.target.value)} />
            </Field>
            <Field label="燃料種類">
              <select value={form.fuelType} onChange={e=>set("fuelType",e.target.value)}>
                {["98無鉛","95無鉛","92無鉛","柴油"].map(v=><option key={v}>{v}</option>)}
              </select>
            </Field>
          </div>

          <Field label={`加油後油表：${form.tankPct}%`}>
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:6 }}>
              <span style={{ fontSize:11, color:C.muted, width:14 }}>E</span>
              <input type="range" min={0} max={100} value={form.tankPct} onChange={e=>set("tankPct",+e.target.value)} style={{ flex:1 }} />
              <span style={{ fontSize:11, color:C.muted, width:14 }}>F</span>
            </div>
            <div style={{ height:8, borderRadius:4, background:C.bg, border:`1px solid ${C.border}`, overflow:"hidden" }}>
              <div style={{
                width:`${form.tankPct}%`, height:"100%", borderRadius:4, transition:"width .3s",
                background: form.tankPct<15?C.red : form.tankPct<30?C.accent : C.accent2
              }} />
            </div>
            {form.tankPct < (car?.warningPct||10) && (
              <p style={{ color:C.red, fontSize:11, marginTop:4 }}>⚠ 低於警示水位，請盡快加油</p>
            )}
          </Field>

          {driven!==null && (
            <div style={{
              background:C.bg, borderRadius:10, padding:"12px 14px",
              border:`1px solid ${C.border}`, marginTop:4,
              display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8,
            }}>
              <div>
                <p style={{ color:C.muted, fontSize:10, marginBottom:3 }}>距上次加油</p>
                <p style={{ fontWeight:700, color:C.accent2, fontFamily:"JetBrains Mono" }}>{driven.toLocaleString()} km</p>
              </div>
              {fuelEff && <div>
                <p style={{ color:C.muted, fontSize:10, marginBottom:3 }}>本次油耗</p>
                <p style={{ fontWeight:700, color:C.green, fontFamily:"JetBrains Mono" }}>{fuelEff} km/L</p>
              </div>}
              {lastFuel && <div>
                <p style={{ color:C.muted, fontSize:10, marginBottom:3 }}>上次加油</p>
                <p style={{ fontWeight:700, fontSize:12 }}>{lastFuel.date}</p>
              </div>}
            </div>
          )}
        </>}

        {form.type!=="fuel" && (
          <Field label="項目說明" span={2}>
            <input value={form.title} onChange={e=>set("title",e.target.value)} placeholder="例：更換機油、輪胎對調..." />
          </Field>
        )}

        <Field label="備註" span={2}>
          <textarea value={form.note} onChange={e=>set("note",e.target.value)} rows={2} placeholder="選填" style={{ resize:"vertical" }} />
        </Field>
      </Card>

      {toast && <div style={{
        padding:"11px 16px", borderRadius:10, marginBottom:12, fontSize:13, fontWeight:600, textAlign:"center",
        background:toast.ok?C.green+"22":C.red+"22",
        border:`1px solid ${toast.ok?C.green:C.red}`,
        color:toast.ok?C.green:C.red
      }}>{toast.msg}</div>}

      <Btn onClick={submit} loading={saving} style={{ width:"100%" }}>
        {gasUrl ? "💾 儲存至 Google Sheets" : "💾 新增記錄"}
      </Btn>
    </div>
  );
}

// ─── Single Record Card ────────────────────────────────────────────────────────
function RecordCard({ rec, effMap, onDelete }) {
  const [deleting, setDeleting] = useState(false);
  const [expanded, setExpanded] = useState(false);
  const typeColor = { fuel:C.accent2, maintenance:C.green, other:C.purple };
  const typeLabel = { fuel:"⛽ 加油", maintenance:"🔧 保養", other:"📋 其他" };

  const del = async () => {
    setDeleting(true);
    await onDelete(rec.id);
    setDeleting(false);
  };

  return (
    <div style={{
      background:C.card, border:`1px solid ${C.border}`, borderRadius:14,
      overflow:"hidden", transition:"border-color .2s",
      borderLeft:`3px solid ${typeColor[rec.type]||C.border}`,
    }}>
      {/* Main row — always visible */}
      <div
        onClick={()=>setExpanded(e=>!e)}
        style={{
          padding:"13px 14px", cursor:"pointer",
          display:"flex", justifyContent:"space-between", alignItems:"center", gap:10,
        }}
      >
        {/* Left: date + type */}
        <div style={{ flex:1, minWidth:0 }}>
          <div style={{ display:"flex", gap:7, alignItems:"center", marginBottom:5, flexWrap:"wrap" }}>
            <span style={{
              fontSize:11, fontWeight:700, letterSpacing:".04em",
              color:typeColor[rec.type]||C.muted,
            }}>{typeLabel[rec.type]||rec.type}</span>
            <span style={{ color:C.muted, fontSize:12, fontFamily:"JetBrains Mono" }}>
              {rec.date}{rec.time ? " "+rec.time : ""}
            </span>
          </div>
          <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"center" }}>
            <span style={{ fontFamily:"JetBrains Mono", fontWeight:800, color:C.accent, fontSize:17 }}>
              ${Number(rec.amount||0).toLocaleString()}
            </span>
            {rec.odometer>0 && (
              <span style={{ color:C.muted, fontSize:12 }}>📍 {Number(rec.odometer).toLocaleString()} km</span>
            )}
            {effMap[rec.id] && (
              <span style={{
                background:C.green+"22", color:C.green, border:`1px solid ${C.green}33`,
                borderRadius:5, padding:"1px 7px", fontSize:11, fontWeight:700,
              }}>🏁 {effMap[rec.id]} km/L</span>
            )}
          </div>
        </div>

        {/* Right: chevron */}
        <span style={{
          color:C.muted, fontSize:14, flexShrink:0,
          transform: expanded?"rotate(180deg)":"rotate(0deg)",
          transition:"transform .25s",
          display:"inline-block",
        }}>▼</span>
      </div>

      {/* Expanded detail */}
      {expanded && (
        <div style={{
          borderTop:`1px solid ${C.border}`, padding:"12px 14px",
          background:C.bg+"cc",
          animation:"fadeIn .2s ease"
        }}>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:10 }}>
            {rec.liters>0 && (
              <div>
                <p style={{ color:C.muted, fontSize:10, marginBottom:2 }}>加油量</p>
                <p style={{ fontWeight:700, color:C.accent2, fontFamily:"JetBrains Mono" }}>{rec.liters} L</p>
              </div>
            )}
            {rec.pricePerLiter>0 && (
              <div>
                <p style={{ color:C.muted, fontSize:10, marginBottom:2 }}>油價</p>
                <p style={{ fontWeight:700, fontFamily:"JetBrains Mono" }}>${rec.pricePerLiter}/L</p>
              </div>
            )}
            {rec.fuelType && rec.type==="fuel" && (
              <div>
                <p style={{ color:C.muted, fontSize:10, marginBottom:2 }}>燃料種類</p>
                <p style={{ fontWeight:700, fontSize:13 }}>{rec.fuelType}</p>
              </div>
            )}
            {rec.tankPct>0 && rec.type==="fuel" && (
              <div>
                <p style={{ color:C.muted, fontSize:10, marginBottom:4 }}>油表</p>
                <div style={{ height:6, borderRadius:3, background:C.surface, overflow:"hidden" }}>
                  <div style={{
                    width:`${rec.tankPct}%`, height:"100%", borderRadius:3,
                    background:rec.tankPct<20?C.red:rec.tankPct<40?C.accent:C.accent2
                  }}/>
                </div>
                <p style={{ color:C.muted, fontSize:10, marginTop:2 }}>{rec.tankPct}%</p>
              </div>
            )}
          </div>
          {rec.title && (
            <div style={{ marginBottom:8 }}>
              <p style={{ color:C.muted, fontSize:10, marginBottom:2 }}>項目說明</p>
              <p style={{ fontSize:13 }}>{rec.title}</p>
            </div>
          )}
          {rec.note && (
            <div style={{ marginBottom:10 }}>
              <p style={{ color:C.muted, fontSize:10, marginBottom:2 }}>備註</p>
              <p style={{ fontSize:12, color:C.muted, lineHeight:1.6 }}>{rec.note}</p>
            </div>
          )}
          <button onClick={del} disabled={deleting} style={{
            background:C.red+"18", border:`1px solid ${C.red}44`, color:C.red,
            borderRadius:8, padding:"7px 16px", fontSize:12, fontWeight:700,
            fontFamily:"Syne", cursor:deleting?"not-allowed":"pointer", opacity:deleting?.5:1
          }}>{deleting?"刪除中...":"🗑 刪除此筆記錄"}</button>
        </div>
      )}
    </div>
  );
}

// ─── Year Group Section ────────────────────────────────────────────────────────
function YearGroup({ year, recs, effMap, onDelete, defaultOpen }) {
  const [open, setOpen] = useState(defaultOpen);
  const fuelTotal = recs.filter(r=>r.type==="fuel").reduce((s,r)=>s+r.amount,0);
  const mainTotal = recs.filter(r=>r.type==="maintenance").reduce((s,r)=>s+r.amount,0);
  const litersTotal = recs.filter(r=>r.type==="fuel").reduce((s,r)=>s+(+r.liters||0),0);
  const fuelCount = recs.filter(r=>r.type==="fuel").length;
  const mainCount = recs.filter(r=>r.type==="maintenance").length;

  return (
    <div style={{ marginBottom:16 }}>
      {/* Year header */}
      <div
        onClick={()=>setOpen(o=>!o)}
        style={{
          display:"flex", justifyContent:"space-between", alignItems:"center",
          padding:"14px 16px", borderRadius:14, cursor:"pointer",
          background: open
            ? `linear-gradient(90deg,${C.accent}18,${C.surface})`
            : C.surface,
          border:`1px solid ${open?C.accent+"55":C.border}`,
          transition:"all .25s", marginBottom: open?10:0,
        }}
      >
        <div style={{ display:"flex", alignItems:"center", gap:14 }}>
          <div style={{
            fontSize:26, fontWeight:800, color: open?C.accent:C.muted,
            fontFamily:"JetBrains Mono", letterSpacing:"-.02em",
            transition:"color .25s",
          }}>{year}</div>
          <div>
            <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
              <span style={{ fontSize:11, color:C.accent, fontWeight:700 }}>
                ⛽ {fuelCount}次 · ${Math.round(fuelTotal).toLocaleString()}
              </span>
              {mainCount>0 && (
                <span style={{ fontSize:11, color:C.green, fontWeight:700 }}>
                  🔧 {mainCount}次 · ${Math.round(mainTotal).toLocaleString()}
                </span>
              )}
            </div>
            <p style={{ fontSize:10, color:C.muted, marginTop:2 }}>
              共 {recs.length} 筆 · 加油 {litersTotal.toFixed(0)} L
            </p>
          </div>
        </div>
        <span style={{
          color:open?C.accent:C.muted, fontSize:14, flexShrink:0,
          transform:open?"rotate(180deg)":"rotate(0deg)", transition:"transform .25s",
          display:"inline-block",
        }}>▼</span>
      </div>

      {/* Cards list */}
      {open && (
        <div style={{ display:"flex", flexDirection:"column", gap:8, paddingLeft:4 }}>
          {recs.map(rec=>(
            <RecordCard key={rec.id} rec={rec} effMap={effMap} onDelete={onDelete} />
          ))}
        </div>
      )}
    </div>
  );
}

// ─── Record List ───────────────────────────────────────────────────────────────
function RecordList({ records, gasUrl, onDelete, onImport, syncing }) {
  const [filter, setFilter] = useState("all");
  const [search, setSearch] = useState("");
  const fileRef = useRef();
  const yearNavRef = useRef();

  // Build eff map globally (across all years)
  const fuelRecs = [...records].filter(r=>r.type==="fuel").sort((a,b)=>a.odometer-b.odometer);
  const effMap = {};
  fuelRecs.forEach((r,i)=>{
    if(i===0) return;
    const prev=fuelRecs[i-1];
    const dist=r.odometer-prev.odometer;
    if(dist>0&&r.liters>0) effMap[r.id]=(dist/r.liters).toFixed(2);
  });

  // Filter + sort
  const filtered = records
    .filter(r=>filter==="all"||r.type===filter)
    .filter(r=>!search||(r.title||"").includes(search)||(r.date||"").includes(search)||(r.note||"").includes(search))
    .sort((a,b)=>(b.date||"").localeCompare(a.date||"")||b.odometer-a.odometer);

  // Group by year
  const yearGroups = useMemo(()=>{
    const map = {};
    filtered.forEach(r=>{
      const y = (r.date||"????").slice(0,4);
      if(!map[y]) map[y]=[];
      map[y].push(r);
    });
    return Object.entries(map).sort(([a],[b])=>b.localeCompare(a));
  },[filtered]);

  const years = yearGroups.map(([y])=>y);
  const latestYear = years[0];

  const handleDelete = async (id) => {
    try {
      if(gasUrl) await gasPost(gasUrl,"deleteRecord",{id});
      onDelete(id);
    } catch(e) { alert("刪除失敗："+e.message); }
  };

  const handleFile = async e => {
    const file=e.target.files[0]; if(!file)return;
    try {
      const newRecs = await importExcel(file, records);
      if(newRecs.length===0){alert("無新資料，全部已存在");return;}
      if(gasUrl) await gasPost(gasUrl,"addRecords",newRecs);
      onImport(newRecs);
      alert(`✅ 匯入 ${newRecs.length} 筆新記錄`);
    } catch(e){alert("匯入失敗: "+e.message);}
    e.target.value="";
  };

  const totalFuel = records.filter(r=>r.type==="fuel").reduce((s,r)=>s+r.amount,0);
  const totalMain = records.filter(r=>r.type==="maintenance").reduce((s,r)=>s+r.amount,0);

  return (
    <div>
      {/* Header */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
        <div>
          <h2 style={{ fontSize:22, fontWeight:800 }}>📋 記錄清單</h2>
          {syncing && <p style={{ fontSize:11, color:C.accent, marginTop:2 }}>⟳ 同步中...</p>}
        </div>
        <Btn onClick={()=>fileRef.current.click()} variant="ghost" style={{ fontSize:12, padding:"8px 12px" }}>
          ⬆ Excel
        </Btn>
      </div>
      <input type="file" ref={fileRef} accept=".xlsx,.xls" onChange={handleFile} style={{ display:"none" }} />

      {/* Global summary */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8, marginBottom:14 }}>
        {[
          ["總筆數", records.length+"筆", C.text],
          ["加油費", "$"+Math.round(totalFuel).toLocaleString(), C.accent],
          ["保養費", "$"+Math.round(totalMain).toLocaleString(), C.green],
        ].map(([l,v,c])=>(
          <div key={l} style={{
            background:C.card, border:`1px solid ${C.border}`, borderRadius:10,
            padding:"10px 8px", textAlign:"center"
          }}>
            <p style={{ fontSize:10, color:C.muted, marginBottom:3 }}>{l}</p>
            <p style={{ fontSize:14, fontWeight:800, color:c, fontFamily:"JetBrains Mono" }}>{v}</p>
          </div>
        ))}
      </div>

      {/* Year quick-nav horizontal scroll */}
      {years.length > 1 && (
        <div ref={yearNavRef} style={{
          display:"flex", gap:6, overflowX:"auto", marginBottom:12,
          paddingBottom:4, scrollbarWidth:"none",
        }}>
          {years.map(y=>(
            <button key={y} onClick={()=>{
              document.getElementById("year-"+y)?.scrollIntoView({behavior:"smooth",block:"start"});
            }} style={{
              flexShrink:0, padding:"5px 16px", borderRadius:20,
              background:y===latestYear?C.accent:C.card,
              color:y===latestYear?"#000":C.muted,
              border:`1px solid ${y===latestYear?C.accent:C.border}`,
              fontFamily:"JetBrains Mono", fontSize:13, fontWeight:700, cursor:"pointer"
            }}>{y}</button>
          ))}
        </div>
      )}

      {/* Type filter */}
      <div style={{ display:"flex", gap:6, marginBottom:10 }}>
        {[["all","全部"],["fuel","⛽加油"],["maintenance","🔧保養"],["other","📋其他"]].map(([v,l])=>(
          <button key={v} onClick={()=>setFilter(v)} style={{
            padding:"5px 12px", borderRadius:20, flexShrink:0,
            background:filter===v?C.accent:C.card,
            color:filter===v?"#000":C.muted,
            border:`1px solid ${filter===v?C.accent:C.border}`,
            fontFamily:"Syne", fontSize:11, fontWeight:700, cursor:"pointer"
          }}>{l}</button>
        ))}
      </div>

      {/* Search */}
      <input placeholder="🔍 搜尋日期、項目、備註..." value={search}
        onChange={e=>setSearch(e.target.value)} style={{ marginBottom:14 }} />

      {/* Empty state */}
      {filtered.length===0 && (
        <div style={{ textAlign:"center", padding:"52px 0", color:C.muted }}>
          <p style={{ fontSize:44, marginBottom:10 }}>📂</p>
          <p style={{ fontSize:14 }}>{search?"找不到符合的記錄":"暫無記錄"}</p>
          <p style={{ fontSize:12, marginTop:4 }}>請新增記錄或匯入 Excel 檔案</p>
        </div>
      )}

      {/* Year groups */}
      <style>{`@keyframes fadeIn{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:translateY(0)}}`}</style>
      {yearGroups.map(([year, recs],idx)=>(
        <div key={year} id={"year-"+year}>
          <YearGroup
            year={year}
            recs={recs}
            effMap={effMap}
            onDelete={handleDelete}
            defaultOpen={idx===0}
          />
        </div>
      ))}
    </div>
  );
}

// ─── Trends Page ───────────────────────────────────────────────────────────────
function TrendPage({ records }) {
  const [view, setView] = useState("month");
  const [year, setYear] = useState(new Date().getFullYear().toString());
  const years = [...new Set(records.map(r=>r.date?.slice(0,4)).filter(Boolean))].sort();
  if(years.length && !years.includes(year)) { /* use first */ }

  const fuelRecs = [...records].filter(r=>r.type==="fuel").sort((a,b)=>a.odometer-b.odometer);
  const effList = fuelRecs.map((r,i)=>{
    if(i===0) return null;
    const prev=fuelRecs[i-1]; const dist=r.odometer-prev.odometer;
    return dist>0&&r.liters>0?{...r,eff:dist/r.liters}:null;
  }).filter(Boolean);

  const monthData = useMemo(()=>{
    const mo = Array.from({length:12},(_,i)=>({
      label:`${i+1}月`, fuelCost:0, mainCost:0, liters:0, effSum:0, effCnt:0
    }));
    records.forEach(r=>{
      if(!r.date||r.date.slice(0,4)!==year)return;
      const m=parseInt(r.date.slice(5,7))-1;
      if(r.type==="fuel"){mo[m].fuelCost+=r.amount;mo[m].liters+=r.liters;}
      if(r.type==="maintenance")mo[m].mainCost+=r.amount;
    });
    effList.forEach(e=>{
      if(!e.date||e.date.slice(0,4)!==year)return;
      const m=parseInt(e.date.slice(5,7))-1;
      mo[m].effSum+=e.eff; mo[m].effCnt+=1;
    });
    return mo.map(m=>({...m, avgEff:m.effCnt>0?+(m.effSum/m.effCnt).toFixed(2):null}));
  },[records,year,effList]);

  const yearData = useMemo(()=>{
    const map={};
    records.forEach(r=>{
      if(!r.date)return;
      const y=r.date.slice(0,4);
      if(!map[y])map[y]={year:y,fuelCost:0,mainCost:0,liters:0};
      if(r.type==="fuel"){map[y].fuelCost+=r.amount;map[y].liters+=r.liters;}
      if(r.type==="maintenance")map[y].mainCost+=r.amount;
    });
    return Object.values(map).sort((a,b)=>a.year.localeCompare(b.year));
  },[records]);

  const data = view==="month"?monthData:yearData;
  const xKey = view==="month"?"label":"year";
  const tt = { contentStyle:{background:C.card,border:`1px solid ${C.border}`,borderRadius:8,fontFamily:"Syne",fontSize:12},labelStyle:{color:C.text} };

  const totalFuel=records.filter(r=>r.type==="fuel").reduce((s,r)=>s+r.amount,0);
  const totalMain=records.filter(r=>r.type==="maintenance").reduce((s,r)=>s+r.amount,0);
  const avgEff=effList.length?(effList.reduce((s,r)=>s+r.eff,0)/effList.length).toFixed(2):"—";
  const totalL=records.filter(r=>r.type==="fuel").reduce((s,r)=>s+r.liters,0);

  return (
    <div>
      <h2 style={{ fontSize:22, fontWeight:800, marginBottom:20 }}>📈 趨勢分析</h2>
      <OilPriceBar />

      <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:8, marginBottom:20 }}>
        {[
          ["💰 總加油費","$"+Math.round(totalFuel).toLocaleString(),C.accent],
          ["🔧 總保養費","$"+Math.round(totalMain).toLocaleString(),C.green],
          ["🏁 平均油耗",avgEff+" km/L",C.accent2],
          ["⛽ 總加油量",totalL.toFixed(0)+" L",C.purple],
        ].map(([l,v,c])=>(
          <Card key={l} style={{ textAlign:"center", padding:"14px 10px" }}>
            <p style={{ color:C.muted, fontSize:11, marginBottom:4 }}>{l}</p>
            <p style={{ fontSize:18, fontWeight:800, color:c, fontFamily:"JetBrains Mono" }}>{v}</p>
          </Card>
        ))}
      </div>

      <div style={{ display:"flex", gap:8, marginBottom:16, alignItems:"center", flexWrap:"wrap" }}>
        {[["month","月份"],["year","年份"]].map(([v,l])=>(
          <button key={v} onClick={()=>setView(v)} style={{
            padding:"7px 18px", borderRadius:20,
            background:view===v?C.accent:C.card, color:view===v?"#000":C.muted,
            border:`1px solid ${view===v?C.accent:C.border}`,
            fontFamily:"Syne", fontSize:12, fontWeight:700, cursor:"pointer"
          }}>{l}</button>
        ))}
        {view==="month" && years.length>0 && (
          <select value={year} onChange={e=>setYear(e.target.value)} style={{ width:"auto",padding:"7px 14px",fontSize:12 }}>
            {years.map(y=><option key={y} value={y}>{y}年</option>)}
          </select>
        )}
      </div>

      <Card style={{ marginBottom:12, padding:"18px 8px" }}>
        <p style={{ color:C.muted, fontSize:11, fontWeight:600, marginBottom:10, paddingLeft:10, letterSpacing:".06em" }}>費用分析 (新台幣)</p>
        <ResponsiveContainer width="100%" height={210}>
          <BarChart data={data} margin={{top:0,right:8,bottom:0,left:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey={xKey} tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} />
            <YAxis tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} width={52}
              tickFormatter={v=>v>=1000?(v/1000).toFixed(0)+"k":v} />
            <Tooltip {...tt} formatter={(v,n)=>["$"+Math.round(v).toLocaleString(),n]} />
            <Legend wrapperStyle={{fontSize:12,paddingTop:8}} />
            <Bar dataKey="fuelCost" name="加油費" fill={C.accent} radius={[4,4,0,0]} />
            <Bar dataKey="mainCost" name="保養費" fill={C.green} radius={[4,4,0,0]} />
          </BarChart>
        </ResponsiveContainer>
      </Card>

      <Card style={{ marginBottom:12, padding:"18px 8px" }}>
        <p style={{ color:C.muted, fontSize:11, fontWeight:600, marginBottom:10, paddingLeft:10, letterSpacing:".06em" }}>油耗趨勢 (km/L)</p>
        <ResponsiveContainer width="100%" height={180}>
          <ComposedChart data={monthData.filter(d=>d.avgEff)} margin={{top:0,right:8,bottom:0,left:0}}>
            <defs>
              <linearGradient id="eg" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor={C.accent2} stopOpacity={0.35}/>
                <stop offset="95%" stopColor={C.accent2} stopOpacity={0}/>
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="label" tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} />
            <YAxis tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} width={38} domain={["auto","auto"]} />
            <Tooltip {...tt} />
            <Area type="monotone" dataKey="avgEff" name="油耗 km/L" stroke={C.accent2}
              fill="url(#eg)" strokeWidth={2.5} dot={{fill:C.accent2,r:3,strokeWidth:0}} />
          </ComposedChart>
        </ResponsiveContainer>
      </Card>

      <Card style={{ padding:"18px 8px" }}>
        <p style={{ color:C.muted, fontSize:11, fontWeight:600, marginBottom:10, paddingLeft:10, letterSpacing:".06em" }}>加油量 (公升)</p>
        <ResponsiveContainer width="100%" height={160}>
          <BarChart data={data} margin={{top:0,right:8,bottom:0,left:0}}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey={xKey} tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} />
            <YAxis tick={{fill:C.muted,fontSize:11}} axisLine={false} tickLine={false} width={38} />
            <Tooltip {...tt} />
            <Bar dataKey="liters" name="公升" fill={C.accent2} radius={[4,4,0,0]} />
          </BarChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

// ─── Root App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("list");
  const [gasUrl, setGasUrl] = useState(null);   // null = not set yet
  const [car, setCar] = useState(null);
  const [records, setRecords] = useState([]);
  const [syncing, setSyncing] = useState(false);
  const [booted, setBooted] = useState(false);
  const [setupDone, setSetupDone] = useState(false);

  // Boot: load config from local storage
  useEffect(()=>{
    (async()=>{
      const savedUrl = await lsGet("gas_url");
      const savedSetup = await lsGet("setup_done");
      if (savedSetup) {
        setSetupDone(true);
        setGasUrl(savedUrl || "");
        await syncData(savedUrl || "");
      }
      setBooted(true);
    })();
  },[]);

  const syncData = async (url) => {
    setSyncing(true);
    try {
      if (url) {
        const res = await gasGet(url);
        if (res.ok !== false) {
          if (Object.keys(res.car||{}).length) setCar(res.car);
          if (res.records?.length) setRecords(res.records.map(r=>({...r,odometer:+r.odometer,amount:+r.amount,liters:+r.liters})));
        }
      } else {
        const c = await lsGet("car_info");
        const r = await lsGet("records");
        if(c) setCar(c);
        if(r) setRecords(r);
      }
    } catch(e) { console.warn("Sync failed:", e.message); }
    setSyncing(false);
  };

  const handleSetupDone = async (url) => {
    await lsSave("gas_url", url || "");
    await lsSave("setup_done", true);
    setGasUrl(url || "");
    setSetupDone(true);
    if (url) await syncData(url);
  };

  const addRecord = async (rec) => {
    const next = [...records, rec];
    setRecords(next);
    if (!gasUrl) await lsSave("records", next);
  };

  const deleteRecord = async (id) => {
    const next = records.filter(r=>r.id!=id);
    setRecords(next);
    if (!gasUrl) await lsSave("records", next);
  };

  const importRecords = async (recs) => {
    const next = [...records, ...recs];
    setRecords(next);
    if (!gasUrl) await lsSave("records", next);
  };

  const navTabs = [
    { id:"list", icon:"📋", label:"記錄" },
    { id:"add",  icon:"＋",  label:"新增" },
    { id:"trend",icon:"📈", label:"趨勢" },
    { id:"car",  icon:"🚗", label:"車輛" },
  ];

  if (!booted) return (
    <div style={{ height:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:C.bg }}>
      <div style={{ textAlign:"center" }}>
        <p style={{ fontSize:40, marginBottom:12 }}>🚗</p>
        <p style={{ color:C.muted, fontSize:14 }}>載入中...</p>
      </div>
    </div>
  );

  if (!setupDone) return (
    <>
      <style>{GS}</style>
      <div style={{ background:C.bg, minHeight:"100vh" }}>
        <SetupPage onDone={handleSetupDone} />
      </div>
    </>
  );

  return (
    <>
      <style>{GS}</style>
      <div style={{ maxWidth:480, margin:"0 auto", minHeight:"100vh", background:C.bg, display:"flex", flexDirection:"column" }}>
        {/* Header */}
        <div style={{
          background:C.surface, borderBottom:`1px solid ${C.border}`,
          padding:"14px 18px 12px", position:"sticky", top:0, zIndex:20
        }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
            <div>
              <h1 style={{ fontSize:19, fontWeight:800, letterSpacing:"-.01em", display:"flex", alignItems:"center", gap:8 }}>
                {car?.plate||"我的車"}
                <span style={{ width:7, height:7, borderRadius:"50%", background:gasUrl?C.green:C.accent, display:"inline-block" }} title={gasUrl?"已連線 Google Sheets":"本機儲存"} />
              </h1>
              <p style={{ fontSize:11, color:C.muted, marginTop:1 }}>
                {car?.model||"請先設定車輛"} {car?.odometer ? "· "+car.odometer.toLocaleString()+" km" : ""}
                {syncing && " · ⟳同步中"}
              </p>
            </div>
            <div style={{ display:"flex", gap:8, alignItems:"center" }}>
              <button onClick={()=>syncData(gasUrl)} style={{
                background:"transparent", border:`1px solid ${C.border}`, borderRadius:8,
                padding:"5px 10px", color:C.muted, cursor:"pointer", fontSize:13
              }} title="重新同步">🔄</button>
              <div style={{
                background:C.accent+"22", border:`1px solid ${C.accent}44`,
                borderRadius:8, padding:"5px 12px",
                fontSize:12, fontWeight:700, color:C.accent, fontFamily:"JetBrains Mono"
              }}>{records.length}筆</div>
            </div>
          </div>
        </div>

        {/* Content */}
        <div style={{ flex:1, overflowY:"auto", padding:"18px 14px 84px" }}>
          {tab==="car"  && <CarSetup car={car} gasUrl={gasUrl} onSave={c=>{setCar(c);setTab("list");}} />}
          {tab==="add"  && <AddRecord car={car} records={records} gasUrl={gasUrl} onAdd={addRecord} />}
          {tab==="list" && <RecordList records={records} gasUrl={gasUrl} syncing={syncing} onDelete={deleteRecord} onImport={importRecords} />}
          {tab==="trend"&& <TrendPage records={records} />}
        </div>

        {/* Bottom Nav */}
        <div style={{
          position:"fixed", bottom:0, left:"50%", transform:"translateX(-50%)",
          width:"100%", maxWidth:480, background:C.surface, borderTop:`1px solid ${C.border}`,
          display:"flex", zIndex:20, paddingBottom:2
        }}>
          {navTabs.map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)} style={{
              flex:1, padding:"10px 0 6px", background:"transparent", border:"none",
              color:tab===t.id?C.accent:C.muted, cursor:"pointer",
              fontFamily:"Syne", fontWeight:700, transition:"color .2s",
              borderTop:tab===t.id?`2px solid ${C.accent}`:"2px solid transparent"
            }}>
              <div style={{ fontSize:19 }}>{t.icon}</div>
              <div style={{ fontSize:10, marginTop:1 }}>{t.label}</div>
            </button>
          ))}
        </div>
      </div>
    </>
  );
}
