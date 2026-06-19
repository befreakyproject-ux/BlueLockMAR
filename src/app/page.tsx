"use client";

import React, { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ==========================================
// 🗂️ STRUCTURE BLUEPRINTS
// ==========================================
interface KPIMapping {
  id: string;
  displayName: string;
  excelColumnKeyword: string;
  category: "Quality" | "Productivity" | "Compliance";
  isHexFavorite: boolean; 
}

interface DynamicAgent {
  login: string;
  overallRating: number;
}

interface HistoricalWeekRecord {
  weekId: string;
  metrics: Record<string, number | string>;
}

const theme = {
  bg: "#050811",
  surface: "#0d1527",
  surfaceLight: "#16223f",
  accent: "#00f0ff",     
  accentGold: "#e2c044", 
  border: "#1e2f55",
  textMuted: "#64748b",
  textLight: "#f8fafc",
  success: "#34d399",
  danger: "#ef4444"
};

export default function BlueLockMA() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  // Navigation State Tab Engine
  const [activeTab, setActiveTab] = useState<"dashboard" | "parameters" | "actionPlan">("dashboard");
  
  // Core Selection Target States
  const [selectedAgentLogin, setSelectedAgentLogin] = useState<string>("");
  const [selectedCategoryFilter, setSelectedCategoryFilter] = useState<string>("All");

  // Local History Recalculation Repositories
  const [masterHistoryMap, setMasterHistoryMap] = useState<Record<string, HistoricalWeekRecord[]>>({});
  const [discoveredAgents, setDiscoveredAgents] = useState<DynamicAgent[]>([]);
  const [allDetectedHeaders, setAllDetectedHeaders] = useState<string[]>([]);
  const [selectedHeaderFromDropdown, setSelectedHeaderFromDropdown] = useState<string>("");

  // Base Parameter Mappings Core Configuration
  const [kpiRegistry, setKpiRegistry] = useState<KPIMapping[]>([
    { id: "ccx_t2b", displayName: "CCX Top 2 Box", excelColumnKeyword: "CCX-Overall Top-Two Box %", category: "Quality", isHexFavorite: true },
    { id: "missed_overall", displayName: "Overall Missed Rate", excelColumnKeyword: "Overall Missed Contact Rate", category: "Productivity", isHexFavorite: true },
    { id: "aht_min", displayName: "AHT (Min)", excelColumnKeyword: "AHT (Min)", category: "Productivity", isHexFavorite: true },
    { id: "adherence", displayName: "Adherence", excelColumnKeyword: "Adherence %", category: "Productivity", isHexFavorite: true },
    { id: "transfer_rate", displayName: "Transfer Rate", excelColumnKeyword: "Transfer Rate", category: "Compliance", isHexFavorite: true },
    { id: "actual_cph", displayName: "Actual CPH", excelColumnKeyword: "Actual CPH", category: "Productivity", isHexFavorite: true }
  ]);

  const [newKpiName, setNewKpiName] = useState("");
  const [newKpiCat, setNewKpiCat] = useState<"Quality" | "Productivity" | "Compliance">("Productivity");

  const [milestones, setMilestones] = useState([
    { week: "Weeks 1–2", title: "Deconstruction & Target Framing", completed: true, desc: "Isolate precise root-causes and baseline current metrics." },
    { week: "Weeks 3–4", title: "Analytical Blueprint Emulation", completed: false, desc: "Perform side-by-side active sessions and calibrate handling procedures." },
    { week: "Weeks 5–6", title: "Ego Weapon Awakening", completed: false, desc: "Stabilize threshold autonomy and secure the target metrics." }
  ]);

  // ==========================================
  // 💾 1. RECOVERY SYSTEM ENGINE (LOAD FROM DRIVE / STORAGE)
  // ==========================================
  useEffect(() => {
    try {
      const savedHistory = localStorage.getItem("bluelock_master_history");
      const savedAgents = localStorage.getItem("bluelock_discovered_agents");
      const savedHeaders = localStorage.getItem("bluelock_detected_headers");
      const savedKpis = localStorage.getItem("bluelock_kpi_registry");

      if (savedHistory) setMasterHistoryMap(JSON.parse(savedHistory));
      if (savedAgents) {
        const parsedAgents = JSON.parse(savedAgents);
        setDiscoveredAgents(parsedAgents);
        if (parsedAgents.length > 0) setSelectedAgentLogin(parsedAgents[0].login);
      }
      if (savedHeaders) {
        const parsedHeaders = JSON.parse(savedHeaders);
        setAllDetectedHeaders(parsedHeaders);
        if (parsedHeaders.length > 0) setSelectedHeaderFromDropdown(parsedHeaders[0]);
      }
      if (savedKpis) setKpiRegistry(JSON.parse(savedKpis));
    } catch (e) {
      console.error("Local recovery storage initialized empty.");
    }
  }, []);

  // Wipes history clean if you ever want a fresh generation reset
  const clearLocalDatabaseEngine = () => {
    if (confirm("Are you sure you want to wipe all local historical tracking memory?")) {
      localStorage.clear();
      setMasterHistoryMap({});
      setDiscoveredAgents([]);
      setAllDetectedHeaders([]);
      setSelectedAgentLogin("");
      alert("Local storage wiped.");
    }
  };

  // ==========================================
  // 📥 AUTO-CLEANING MATRIX INGESTION ENGINE
  // ==========================================
  const processUploadedExcelFiles = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const dataBinary = evt.target?.result;
        const workbook = XLSX.read(dataBinary, { type: "binary" });
        const targetSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json<any[]>(targetSheet, { header: 1 });

        if (!rawRows || rawRows.length === 0) return;

        let bestHeaderIdx = 0;
        let maxFilledColumns = 0;

        for (let r = 0; r < Math.min(rawRows.length, 15); r++) {
          const rowCells = rawRows[r] || [];
          const filledCount = rowCells.filter(c => c !== null && String(c).trim() !== "").length;
          
          if (filledCount > maxFilledColumns && r !== 0) {
            maxFilledColumns = filledCount;
            bestHeaderIdx = r;
          }
        }

        let detectedWeek = "Current Generation";
        for (let r = 0; r < Math.min(rawRows.length, bestHeaderIdx + 1); r++) {
          const cells = rawRows[r] || [];
          for (let c = 0; c < cells.length; c++) {
            const txt = String(cells[c]).trim();
            if (txt.toLowerCase().includes("week") || txt === "25" || txt === "24") {
              detectedWeek = txt.toLowerCase().includes("week") ? txt : `Week ${txt}`;
              break;
            }
          }
        }

        const headers = (rawRows[bestHeaderIdx] || []).map((h: any) => String(h || "").trim());
        const filteredKPIHeaders = headers.filter(h => h && h.toLowerCase() !== "agent" && !h.toLowerCase().includes("week") && h.length > 1);

        let finalHeadersList = [...allDetectedHeaders];
        if (filteredKPIHeaders.length > 0) {
          finalHeadersList = Array.from(new Set([...allDetectedHeaders, ...filteredKPIHeaders]));
          setAllDetectedHeaders(finalHeadersList);
          setSelectedHeaderFromDropdown(filteredKPIHeaders[0]);
          localStorage.setItem("bluelock_detected_headers", JSON.stringify(finalHeadersList));
        }

        const localHistoryUpdate = { ...masterHistoryMap };
        const updatedAgentsList = [...discoveredAgents];

        rawRows.slice(bestHeaderIdx + 1).forEach((row) => {
          if (!row || row[0] === undefined || row[0] === null) return;
          
          const loginStr = String(row[0]).trim();
          const lowerLogin = loginStr.toLowerCase();

          if (
            loginStr === "" || 
            lowerLogin.includes("total") || 
            lowerLogin.includes("agent") || 
            lowerLogin.includes("amazon") || 
            lowerLogin.includes("site group") || 
            lowerLogin.includes("marketplace") ||
            lowerLogin.includes("none") ||
            lowerLogin.includes("null") ||
            !isNaN(Number(loginStr)) 
          ) {
            return; 
          }

          if (!localHistoryUpdate[loginStr]) localHistoryUpdate[loginStr] = [];
          let existingWeekRecord = localHistoryUpdate[loginStr].find(record => record.weekId === detectedWeek);
          
          if (!existingWeekRecord) {
            existingWeekRecord = { weekId: detectedWeek, metrics: {} };
            localHistoryUpdate[loginStr].push(existingWeekRecord);
          }

          headers.forEach((h, index) => {
            if (!h) return;
            const targetKpi = kpiRegistry.find(k => k.excelColumnKeyword.toLowerCase() === h.toLowerCase());
            const parsedCellVal = row[index] !== undefined ? row[index] : "—";
            
            if (targetKpi) {
              existingWeekRecord!.metrics[targetKpi.id] = parsedCellVal;
            } else {
              existingWeekRecord!.metrics[h] = parsedCellVal;
            }
          });

          if (!updatedAgentsList.some(a => a.login === loginStr)) {
            updatedAgentsList.push({ login: loginStr, overallRating: Math.floor(Math.random() * (96 - 81 + 1)) + 81 });
          }
        });

        setMasterHistoryMap(localHistoryUpdate);
        setDiscoveredAgents(updatedAgentsList);
        
        // 💾 Save current state configuration objects natively to disk arrays
        localStorage.setItem("bluelock_master_history", JSON.stringify(localHistoryUpdate));
        localStorage.setItem("bluelock_discovered_agents", JSON.stringify(updatedAgentsList));

        if (updatedAgentsList.length > 0 && (!selectedAgentLogin || !updatedAgentsList.some(a => a.login === selectedAgentLogin))) {
          setSelectedAgentLogin(updatedAgentsList[0].login);
        }

        alert(`Ingested and locked tracking parameters into permanent local save for ${detectedWeek}!`);
      } catch (err) {
        console.error(err);
      }
    };
    reader.readAsBinaryString(file);
  };

  const toggleHexFavorite = (id: string) => {
    const activeFavoritesCount = kpiRegistry.filter(k => k.isHexFavorite).length;
    const targets = kpiRegistry.find(k => k.id === id);
    
    if (targets && !targets.isHexFavorite && activeFavoritesCount >= 6) {
      alert("The Ego Matrix Hexagon is locked to maximum 6 active favorite components. Unstar an old one first!");
      return;
    }
    const updatedKpis = kpiRegistry.map(k => k.id === id ? { ...k, isHexFavorite: !k.isHexFavorite } : k);
    setKpiRegistry(updatedKpis);
    localStorage.setItem("bluelock_kpi_registry", JSON.stringify(updatedKpis));
  };

  const currentAgentData = discoveredAgents.find(a => a.login === selectedAgentLogin);
  const activeAgentHistory = selectedAgentLogin ? (masterHistoryMap[selectedAgentLogin] || []) : [];
  const runningHexKPIs = kpiRegistry.filter(k => k.isHexFavorite);

  const renderDynamicHexPoints = () => {
    const center = 100;
    const maxRadius = 75;
    const points: string[] = [];
    
    for (let i = 0; i < 6; i++) {
      const angle = (i * 60 - 90) * (Math.PI / 180);
      const kpi = runningHexKPIs[i];
      let performancePct = 0.75; 
      
      if (kpi && currentAgentData) {
        const records = masterHistoryMap[currentAgentData.login] || [];
        const latest = records[records.length - 1];
        const rawScore = latest ? latest.metrics[kpi.id] || latest.metrics[kpi.excelColumnKeyword] : null;
        
        if (typeof rawScore === "number") {
          performancePct = rawScore < 1 ? rawScore : Math.min(rawScore / 100, 1);
        }
      }

      const r = maxRadius * performancePct;
      const x = center + r * Math.cos(angle);
      const y = center + r * Math.sin(angle);
      points.push(`${x},${y}`);
    }
    return points.join(" ");
  };

  return (
    <div style={{ backgroundColor: theme.bg, color: theme.textLight, minHeight: "100vh", padding: "32px", fontFamily: "sans-serif" }}>
      
      {/* HUB HEADER BAR COMPONENTS */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `2px solid ${theme.border}`, paddingBottom: "20px", marginBottom: "32px" }}>
        <div>
          <h1 style={{ fontSize: "26px", fontWeight: "900", color: theme.accent, letterSpacing: "1.5px", margin: 0 }}>PROJECT: BLUELOCKMA</h1>
          <div style={{ display: "flex", gap: "12px", marginTop: "12px" }}>
            <button onClick={() => setActiveTab("dashboard")} style={{ padding: "8px 16px", backgroundColor: activeTab === "dashboard" ? theme.accent : "transparent", color: activeTab === "dashboard" ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}>🏟️ Live Dashboard Matrix</button>
            <button onClick={() => setActiveTab("parameters")} style={{ padding: "8px 16px", backgroundColor: activeTab === "parameters" ? theme.accent : "transparent", color: activeTab === "parameters" ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}>⚙️ Settings Parameter Portfolio</button>
            <button onClick={() => setActiveTab("actionPlan")} style={{ padding: "8px 16px", backgroundColor: activeTab === "actionPlan" ? theme.accent : "transparent", color: activeTab === "actionPlan" ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}>⚔️ 6-Week Awakening Plan</button>
          </div>
        </div>
        
        <div style={{ display: "flex", gap: "12px" }}>
          <button onClick={clearLocalDatabaseEngine} style={{ backgroundColor: "transparent", color: theme.danger, border: `1px solid ${theme.danger}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold", cursor: "pointer" }}>🗑️ Wipe Local Save</button>
          <button onClick={() => fileInputRef.current?.click()} style={{ backgroundColor: theme.surfaceLight, color: theme.accent, border: `1px solid ${theme.border}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold", cursor: "pointer" }}>📥 Ingest Master File</button>
          <input type="file" ref={fileInputRef} onChange={processUploadedExcelFiles} accept=".xlsx, .xls" style={{ display: "none" }} />
          
          <select value={selectedAgentLogin} onChange={(e) => setSelectedAgentLogin(e.target.value)} style={{ backgroundColor: theme.surface, color: theme.accent, border: `2px solid ${theme.accent}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold" }}>
            {discoveredAgents.length === 0 ? <option value="">-- Save File to Memory --</option> : discoveredAgents.map(a => <option key={a.login} value={a.login}>@{a.login}</option>)}
          </select>
        </div>
      </div>

      {/* ==========================================
          🏟️ TAB 1: DASHBOARD
         ========================================== */}
      {activeTab === "dashboard" && (
        <div style={{ display: "grid", gridTemplateColumns: "380px 1fr", gap: "32px" }}>
          
          {/* PROFILE CARD & HEX */}
          <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
            {!currentAgentData ? (
              <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "40px", textAlign: "center" }}>
                <div style={{ fontSize: "50px" }}>💾</div>
                <h3 style={{ color: theme.accent }}>PERMANENT AUTOSAVE STORAGE ACTIVE</h3>
                <p style={{ fontSize: "12px", color: theme.textMuted }}>Your records stay on your hard drive. Drop spreadsheet files above to load up your last generation workspace view instantly.</p>
              </div>
            ) : (
              <>
                <div style={{ background: "linear-gradient(135deg, #16223f 0%, #0d1527 100%)", border: `3px solid ${theme.accentGold}`, borderRadius: "20px", padding: "24px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <div><span style={{ fontSize: "44px", fontWeight: "900", color: theme.accentGold, display: "block" }}>{currentAgentData.overallRating}</span><span style={{ fontSize: "11px", fontWeight: "bold", color: theme.accent }}>STRIKER</span></div>
                    <span style={{ fontSize: "13px" }}>🇲🇦 EGOIST</span>
                  </div>
                  <div style={{ textAlign: "center", margin: "16px 0" }}><span style={{ fontSize: "60px" }}>👤</span><h2 style={{ fontSize: "22px", color: theme.accentGold, margin: "6px 0" }}>{currentAgentData.login.toUpperCase()}</h2></div>
                  
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", borderTop: `1px solid ${theme.border}`, paddingTop: "12px", fontSize: "11px" }}>
                    {kpiRegistry.filter(k => k.isHexFavorite).slice(0, 6).map(kpi => {
                      const records = masterHistoryMap[currentAgentData.login] || [];
                      const latestWeek = records[records.length - 1];
                      const val = latestWeek ? latestWeek.metrics[kpi.id] || latestWeek.metrics[kpi.excelColumnKeyword] : "—";
                      const fmt = typeof val === "number" ? (val < 1 ? `${(val * 100).toFixed(0)}%` : val.toFixed(1)) : val;
                      return <div key={kpi.id}><strong>{fmt}</strong> <span style={{ color: theme.textMuted, display: "block" }}>{kpi.displayName}</span></div>;
                    })}
                  </div>
                </div>

                <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "20px", textAlign: "center" }}>
                  <h4 style={{ color: theme.accent, margin: "0 0 4px 0", fontSize: "13px", textAlign: "left" }}>⬡ EGO MATRIX FAVORITES (0-100)</h4>
                  
                  <div style={{ display: "flex", justifyContent: "center", alignItems: "center", position: "relative" }}>
                    <svg width="220" height="220" viewBox="0 0 200 200">
                      <polygon points="100,25 175,68 175,153 100,195 25,153 25,68" fill="none" stroke={theme.border} strokeWidth="1" />
                      <polygon points="100,50 150,79 150,136 100,165 50,136 50,79" fill="none" stroke={theme.border} strokeWidth="1" strokeDasharray="3" />
                      <polygon points={renderDynamicHexPoints()} fill="rgba(0, 240, 255, 0.25)" stroke={theme.accent} strokeWidth="2.5" />
                      {runningHexKPIs.map((kpi, idx) => {
                        const angle = (idx * 60 - 90) * (Math.PI / 180);
                        const labelRadius = 88;
                        const lx = 100 + labelRadius * Math.cos(angle);
                        const ly = 100 + labelRadius * Math.sin(angle);
                        return (
                          <text key={kpi.id} x={lx} y={ly} fill={theme.textLight} fontSize="9" fontWeight="bold" textAnchor="middle" alignmentBaseline="middle">
                            {kpi.displayName.slice(0, 8)}
                          </text>
                        );
                      })}
                    </svg>
                  </div>
                </div>
              </>
            )}
          </div>

          {/* MATRIX GRID UNIFIED TABLE */}
          <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
            <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "24px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "16px" }}>
                <div><h3 style={{ margin: 0, color: theme.accent }}>📊 Unified Operational Matrix Grid</h3></div>
                <div style={{ display: "flex", gap: "6px" }}>
                  {["All", "Productivity", "Quality", "Compliance"].map(cat => (
                    <button key={cat} onClick={() => setSelectedCategoryFilter(cat)} style={{ padding: "6px 12px", backgroundColor: selectedCategoryFilter === cat ? theme.accent : theme.bg, color: selectedCategoryFilter === cat ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontSize: "12px" }}>{cat}</button>
                  ))}
                </div>
              </div>

              {discoveredAgents.length === 0 ? <p style={{ color: theme.textMuted, textAlign: "center", padding: "40px" }}>Awaiting local memory cache stream loading index...</p> : (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                    <thead>
                      <tr style={{ backgroundColor: theme.surfaceLight, color: theme.accent, borderBottom: `2px solid ${theme.border}` }}>
                        <th style={{ padding: "12px", textAlign: "left" }}>Egoist Handle</th>
                        {kpiRegistry.filter(k => selectedCategoryFilter === "All" || k.category === selectedCategoryFilter).map(kpi => (
                          <th key={kpi.id} style={{ padding: "12px", textAlign: "center" }}>{kpi.displayName}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {discoveredAgents.map((ag, idx) => (
                        <tr key={ag.login} style={{ backgroundColor: idx % 2 === 0 ? theme.surface : theme.bg, borderBottom: `1px solid ${theme.border}` }}>
                          <td style={{ padding: "10px", color: theme.accent, fontWeight: "bold" }}>@{ag.login}</td>
                          {kpiRegistry.filter(k => selectedCategoryFilter === "All" || k.category === selectedCategoryFilter).map(kpi => {
                            const recs = masterHistoryMap[ag.login] || [];
                            const latestWeek = recs[recs.length - 1];
                            const raw = latestWeek ? latestWeek.metrics[kpi.id] || latestWeek.metrics[kpi.excelColumnKeyword] : "—";
                            const fmt = typeof raw === "number" ? (raw < 1 ? `${(raw * 100).toFixed(2)}%` : raw.toFixed(2)) : "—";
                            return <td key={kpi.id} style={{ padding: "10px", textAlign: "center" }}>{fmt}</td>;
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* PARAMETERS PANEL & ACTION PLANS TABS STAY IDENTICAL BELOW */}
      {activeTab === "parameters" && (
        <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "32px" }}>
          <h3 style={{ color: theme.accent, marginTop: 0 }}>⚙️ Parameter Control Base</h3>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px", marginBottom: "32px" }}>
            {kpiRegistry.map(k => (
              <div key={k.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, padding: "12px 16px", borderRadius: "8px" }}>
                <div><strong>{k.displayName}</strong><span style={{ color: theme.textMuted, fontSize: "11px", display: "block" }}>{k.excelColumnKeyword}</span></div>
                <div style={{ display: "flex", gap: "12px" }}>
                  <button onClick={() => toggleHexFavorite(k.id)} style={{ background: "none", border: "none", color: k.isHexFavorite ? theme.accentGold : theme.textMuted, fontSize: "18px", cursor: "pointer" }}>{k.isHexFavorite ? "★" : "☆"}</button>
                  <button onClick={() => setKpiRegistry(kpiRegistry.filter(item => item.id !== k.id))} style={{ background: "none", border: "none", color: theme.danger, fontWeight: "bold", cursor: "pointer" }}>✕</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {activeTab === "actionPlan" && (
        <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "32px" }}>
          <h2 style={{ color: theme.accentGold, margin: 0 }}>⚔️ 6-Week Action Roadmap</h2>
          <div style={{ display: "flex", flexDirection: "column", gap: "20px", marginTop: "24px" }}>
            {milestones.map((m, idx) => (
              <div key={idx} style={{ backgroundColor: theme.bg, border: `1px solid ${theme.border}`, padding: "24px", borderRadius: "12px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div><strong>{m.week}</strong>: {m.title} <p style={{ margin: "4px 0 0 0", fontSize: "13px", color: theme.textMuted }}>{m.desc}</p></div>
                <button onClick={() => setMilestones(milestones.map((item, i) => i === idx ? { ...item, completed: !item.completed } : item))} style={{ padding: "8px 16px", backgroundColor: m.completed ? theme.success : "transparent", color: m.completed ? theme.bg : theme.accentGold, border: `1px solid ${m.completed ? theme.success : theme.accentGold}`, borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}>{m.completed ? "✓ Stage Unlocked" : "⬡ Awaken Stage"}</button>
              </div>
            ))}
          </div>
        </div>
      )}

    </div>
  );
}
