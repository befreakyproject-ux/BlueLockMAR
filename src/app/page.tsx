"use client";

import React, { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ==========================================
// 🗂️ STRUCTURE INTERFACES
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
  weekNumber: number; // For Range Sliders
  monthName: string;  // e.g. "January"
  quarterId: string;  // e.g. "Q1", "Q2"
  yearId: string;     // e.g. "2026"
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

  // DYNAMIC ADVANCED TIME FRAME CALENDAR FILTERS STATE
  const [timeMacroView, setTimeMacroView] = useState<"weekly" | "monthly" | "quarterly" | "yearly">("weekly");
  const [startWeekSlider, setStartWeekSlider] = useState<number>(1);
  const [endWeekSlider, setEndWeekSlider] = useState<number>(53);
  const [selectedMonthFilter, setSelectedMonthFilter] = useState<string>("All");
  const [selectedQuarterFilter, setSelectedQuarterFilter] = useState<string>("All");

  // Mapped Core Configuration
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
  // 💾 HARD DRIVE AUTOSAVE PERSISTENCE
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
      console.error("Autosave index initialized clean.");
    }
  }, []);

  const clearLocalDatabaseEngine = () => {
    if (confirm("Wipe all locally tracking historical dataset logs permanently?")) {
      localStorage.clear();
      setMasterHistoryMap({});
      setDiscoveredAgents([]);
      setAllDetectedHeaders([]);
      setSelectedAgentLogin("");
      alert("Local registry reset.");
    }
  };

  // ==========================================
  // 📥 ROBUST SCANNING MULTI-FILE INGESTION
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

        // Parse explicit Week, Month, Quarter metadata flags
        let weekNum = 25;
        let weekIdStr = "Week 25";
        let monthNameStr = "June";
        let quarterIdStr = "Q2";

        for (let r = 0; r < Math.min(rawRows.length, bestHeaderIdx + 1); r++) {
          const cells = rawRows[r] || [];
          for (let c = 0; c < cells.length; c++) {
            const txt = String(cells[c]).trim();
            const numMatch = txt.match(/\d+/);
            if (txt.toLowerCase().includes("week") || txt === "25" || txt === "24") {
              weekNum = numMatch ? parseInt(numMatch[0]) : 25;
              weekIdStr = txt.toLowerCase().includes("week") ? txt : `Week ${txt}`;
              
              // Calibrate month and quarters natively from week sequences
              if (weekNum <= 13) { monthNameStr = "February"; quarterIdStr = "Q1"; }
              else if (weekNum <= 26) { monthNameStr = "June"; quarterIdStr = "Q2"; }
              else if (weekNum <= 39) { monthNameStr = "August"; quarterIdStr = "Q3"; }
              else { monthNameStr = "November"; quarterIdStr = "Q4"; }
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
            loginStr === "" || lowerLogin.includes("total") || lowerLogin.includes("agent") || 
            lowerLogin.includes("amazon") || lowerLogin.includes("site group") || 
            lowerLogin.includes("marketplace") || lowerLogin.includes("none") || 
            lowerLogin.includes("null") || !isNaN(Number(loginStr))
          ) {
            return; 
          }

          if (!localHistoryUpdate[loginStr]) localHistoryUpdate[loginStr] = [];
          
          let existingWeekRecord = localHistoryUpdate[loginStr].find(record => record.weekId === weekIdStr);
          if (!existingWeekRecord) {
            existingWeekRecord = { 
              weekId: weekIdStr, weekNumber: weekNum, monthName: monthNameStr, 
              quarterId: quarterIdStr, yearId: "2026", metrics: {} 
            };
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
        
        localStorage.setItem("bluelock_master_history", JSON.stringify(localHistoryUpdate));
        localStorage.setItem("bluelock_discovered_agents", JSON.stringify(updatedAgentsList));

        if (updatedAgentsList.length > 0 && (!selectedAgentLogin || !updatedAgentsList.some(a => a.login === selectedAgentLogin))) {
          setSelectedAgentLogin(updatedAgentsList[0].login);
        }

        alert(`Ingested dataset into local history archive mapping!`);
      } catch (err) {
        console.error(err);
      }
    };
    reader.readAsBinaryString(file);
  };

  // ==========================================
  // 📆 CALENDAR CALCULATION ENGINE
  // ==========================================
  const getFilteredAveragedMetrics = (agentLogin: string) => {
    const rawRecords = masterHistoryMap[agentLogin] || [];
    if (rawRecords.length === 0) return {};

    // Filter array records matching custom date ranges inputs
    const matchingTimeRecords = rawRecords.filter(rec => {
      if (timeMacroView === "weekly") {
        return rec.weekNumber >= startWeekSlider && rec.weekNumber <= endWeekSlider;
      }
      if (timeMacroView === "monthly") {
        return selectedMonthFilter === "All" || rec.monthName === selectedMonthFilter;
      }
      if (timeMacroView === "quarterly") {
        return selectedQuarterFilter === "All" || rec.quarterId === selectedQuarterFilter;
      }
      return true; // Yearly outputs everything
    });

    if (matchingTimeRecords.length === 0) return {};

    // Average numerical scores across selected chronological scopes
    const aggregatedSums: Record<string, number> = {};
    const countMap: Record<string, number> = {};

    matchingTimeRecords.forEach(rec => {
      Object.entries(rec.metrics).forEach(([kpiId, val]) => {
        if (typeof val === "number") {
          aggregatedSums[kpiId] = (aggregatedSums[kpiId] || 0) + val;
          countMap[kpiId] = (countMap[kpiId] || 0) + 1;
        }
      });
    });

    const finalAverages: Record<string, number | string> = {};
    Object.keys(aggregatedSums).forEach(id => {
      finalAverages[id] = aggregatedSums[id] / countMap[id];
    });

    return finalAverages;
  };

  const toggleHexFavorite = (id: string) => {
    const activeFavoritesCount = kpiRegistry.filter(k => k.isHexFavorite).length;
    const targets = kpiRegistry.find(k => k.id === id);
    if (targets && !targets.isHexFavorite && activeFavoritesCount >= 6) {
      alert("The Ego Matrix Hexagon is locked to maximum 6 active parameters.");
      return;
    }
    const updatedKpis = kpiRegistry.map(k => k.id === id ? { ...k, isHexFavorite: !k.isHexFavorite } : k);
    setKpiRegistry(updatedKpis);
    localStorage.setItem("bluelock_kpi_registry", JSON.stringify(updatedKpis));
  };

  const currentAgentData = discoveredAgents.find(a => a.login === selectedAgentLogin);
  const activeComputedMetrics = selectedAgentLogin ? getFilteredAveragedMetrics(selectedAgentLogin) : {};
  const runningHexKPIs = kpiRegistry.filter(k => k.isHexFavorite);

  const renderDynamicHexPoints = () => {
    const center = 100;
    const maxRadius = 75;
    const points: string[] = [];
    
    for (let i = 0; i < 6; i++) {
      const angle = (i * 60 - 90) * (Math.PI / 180);
      const kpi = runningHexKPIs[i];
      let performancePct = 0.75; 
      
      if (kpi && activeComputedMetrics) {
        const rawScore = activeComputedMetrics[kpi.id] || activeComputedMetrics[kpi.excelColumnKeyword];
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
      
      {/* 👑 HUB NAVIGATION Hub Header */}
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
            {discoveredAgents.length === 0 ? <option value="">-- Awaiting Ingest Files --</option> : discoveredAgents.map(a => <option key={a.login} value={a.login}>@{a.login}</option>)}
          </select>
        </div>
      </div>

      {/* ==========================================
          📅 CHRONOLOGICAL RANGE EXTRACTOR CALENDAR HUD
         ========================================== */}
      <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "12px", padding: "16px 24px", marginBottom: "32px", display: "flex", gap: "24px", alignItems: "center", flexWrap: "wrap" }}>
        <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
          <label style={{ fontSize: "11px", color: theme.accent, fontWeight: "bold" }}>CALENDAR TIME HORIZON MAP</label>
          <div style={{ display: "flex", backgroundColor: theme.bg, padding: "2px", borderRadius: "6px", border: `1px solid ${theme.border}` }}>
            {(["weekly", "monthly", "quarterly", "yearly"] as const).map(mode => (
              <button key={mode} onClick={() => setTimeMacroView(mode)} style={{ padding: "6px 12px", fontSize: "12px", textTransform: "capitalize", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: "bold", backgroundColor: timeMacroView === mode ? theme.accent : "transparent", color: timeMacroView === mode ? theme.bg : theme.textLight }}>{mode}</button>
            ))}
          </div>
        </div>

        {/* Dynamic Parameter Sliders for Variable Week Ranges */}
        {timeMacroView === "weekly" && (
          <div style={{ display: "flex", gap: "16px", flex: 1, alignItems: "center" }}>
            <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: "4px" }}>
              <span style={{ fontSize: "11px", color: theme.textMuted }}>START RUN TIME: <strong>Week {startWeekSlider}</strong></span>
              <input type="range" min="1" max="53" value={startWeekSlider} onChange={e => setStartWeekSlider(parseInt(e.target.value))} style={{ accentColor: theme.accent, width: "100%" }} />
            </div>
            <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: "4px" }}>
              <span style={{ fontSize: "11px", color: theme.textMuted }}>END RUN TIME: <strong>Week {endWeekSlider}</strong></span>
              <input type="range" min="1" max="53" value={endWeekSlider} onChange={e => setEndWeekSlider(Math.max(startWeekSlider, parseInt(e.target.value)))} style={{ accentColor: theme.accent, width: "100%" }} />
            </div>
          </div>
        )}

        {/* Month Dropdown Filters View Scope */}
        {timeMacroView === "monthly" && (
          <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
            <span style={{ fontSize: "11px", color: theme.textMuted }}>CHOOSE OPERATIONAL MONTH</span>
            <select value={selectedMonthFilter} onChange={e => setSelectedMonthFilter(e.target.value)} style={{ padding: "8px", backgroundColor: theme.bg, color: "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontSize: "12px" }}>
              <option value="All">All Active Calendar Months</option>
              {["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"].map(m => <option key={m} value={m}>{m}</option>)}
            </select>
          </div>
        )}

        {/* Corporate Quarter Dynamic Dropdowns */}
        {timeMacroView === "quarterly" && (
          <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
            <span style={{ fontSize: "11px", color: theme.textMuted }}>CHOOSE HORIZON QUARTER</span>
            <select value={selectedQuarterFilter} onChange={e => setSelectedQuarterFilter(e.target.value)} style={{ padding: "8px", backgroundColor: theme.bg, color: "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontSize: "12px" }}>
              <option value="All">Full Business Quarter Layouts</option>
              {["Q1", "Q2", "Q3", "Q4"].map(q => <option key={q} value={q}>{q}</option>)}
            </select>
          </div>
        )}
      </div>

      {/* ==========================================
          🏟️ TAB VIEW HUB CORE
         ========================================== */}
      {activeTab === "dashboard" && (
        <div style={{ display: "grid", gridTemplateColumns: "380px 1fr", gap: "32px" }}>
          
          {/* PROFILE LAYOUT PILLARS */}
          <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
            {!currentAgentData ? (
              <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "40px", textAlign: "center" }}>
                <div style={{ fontSize: "50px" }}>💾</div>
                <h3 style={{ color: theme.accent }}>HARD-DRIVE VAULT ONLINE</h3>
                <p style={{ fontSize: "12px", color: theme.textMuted }}>Ingest extraction spreadsheets to generate metrics maps automatically.</p>
              </div>
            ) : (
              <>
                {/* REVOLUTION FIFA ULTIMATE RATING SHOWCASE DESIGN */}
                <div style={{ background: "linear-gradient(135deg, #16223f 0%, #0d1527 100%)", border: `3px solid ${theme.accentGold}`, borderRadius: "20px", padding: "24px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <div><span style={{ fontSize: "44px", fontWeight: "900", color: theme.accentGold, display: "block" }}>{currentAgentData.overallRating}</span><span style={{ fontSize: "11px", fontWeight: "bold", color: theme.accent }}>STRIKER</span></div>
                    <span style={{ fontSize: "13px" }}>🇲🇦 EGOIST</span>
                  </div>
                  <div style={{ textAlign: "center", margin: "16px 0" }}><span style={{ fontSize: "60px" }}>👤</span><h2 style={{ fontSize: "22px", color: theme.accentGold, margin: "6px 0" }}>{currentAgentData.login.toUpperCase()}</h2></div>
                  
                  {/* Aggregated Output matching active calendar filters configurations */}
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", borderTop: `1px solid ${theme.border}`, paddingTop: "12px", fontSize: "11px" }}>
                    {kpiRegistry.filter(k => k.isHexFavorite).slice(0, 6).map(kpi => {
                      const val = activeComputedMetrics[kpi.id] || activeComputedMetrics[kpi.excelColumnKeyword] || "—";
                      const fmt = typeof val === "number" ? (val < 1 ? `${(val * 100).toFixed(0)}%` : val.toFixed(1)) : val;
                      return <div key={kpi.id}><strong>{fmt}</strong> <span style={{ color: theme.textMuted, display: "block" }}>{kpi.displayName}</span></div>;
                    })}
                  </div>
                </div>

                {/* ⬡ THE REAL-TIME GRADED MATRIX HEXAGON (0-100) */}
                <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "20px", textAlign: "center" }}>
                  <h4 style={{ color: theme.accent, margin: "0 0 4px 0", fontSize: "13px", textAlign: "left" }}>⬡ EGO RADAR PERFORMANCE CORES</h4>
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
                        return <text key={kpi.id} x={lx} y={ly} fill={theme.textLight} fontSize="9" fontWeight="bold" textAnchor="middle" alignmentBaseline="middle">{kpi.displayName.slice(0, 8)}</text>;
                      })}
                    </svg>
                  </div>
                </div>
              </>
            )}
          </div>

          {/* DYNAMIC INTEGRATED GRID MATRIX TABLE */}
          <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
            <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "24px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "16px" }}>
                <div><h3 style={{ margin: 0, color: theme.accent }}>📊 Unified Balanced Matrix Records</h3></div>
                <div style={{ display: "flex", gap: "6px" }}>
                  {["All", "Productivity", "Quality", "Compliance"].map(cat => (
                    <button key={cat} onClick={() => setSelectedCategoryFilter(cat)} style={{ padding: "6px 12px", backgroundColor: selectedCategoryFilter === cat ? theme.accent : theme.bg, color: selectedCategoryFilter === cat ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontSize: "12px" }}>{cat}</button>
                  ))}
                </div>
              </div>

              {discoveredAgents.length === 0 ? <p style={{ color: theme.textMuted, textAlign: "center", padding: "40px" }}>Awaiting localized file mapping integration streams...</p> : (
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
                      {discoveredAgents.map((ag, idx) => {
                        const calculatedRowMetrics = getFilteredAveragedMetrics(ag.login);
                        return (
                          <tr key={ag.login} style={{ backgroundColor: idx % 2 === 0 ? theme.surface : theme.bg, borderBottom: `1px solid ${theme.border}` }}>
                            <td style={{ padding: "10px", color: theme.accent, fontWeight: "bold" }}>@{ag.login}</td>
                            {kpiRegistry.filter(k => selectedCategoryFilter === "All" || k.category === selectedCategoryFilter).map(kpi => {
                              const raw = calculatedRowMetrics[kpi.id] || calculatedRowMetrics[kpi.excelColumnKeyword] || "—";
                              const fmt = typeof raw === "number" ? (raw < 1 ? `${(raw * 100).toFixed(2)}%` : raw.toFixed(2)) : "—";
                              return <td key={kpi.id} style={{ padding: "10px", textAlign: "center" }}>{fmt}</td>;
                            })}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ==========================================
          ⚙️ TAB 2: PARAMETERS MANAGEMENT PANEL WITH DRIVE HARDENING SAVES
         ========================================== */}
      {activeTab === "parameters" && (
        <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "32px" }}>
          <h3 style={{ color: theme.accent, marginTop: 0 }}>⚙️ Systems Parameter Settings Panel</h3>
          <p style={{ fontSize: "13px", color: theme.textMuted, marginBottom: "24px" }}>Click the **★ Star icon** to link choices to the Radar Chart. Changes automatically autosave onto your machine hard drive partition.</p>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px", marginBottom: "32px" }}>
            {kpiRegistry.map(k => (
              <div key={k.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, padding: "12px 16px", borderRadius: "8px" }}>
                <div><strong>{k.displayName}</strong><span style={{ color: theme.textMuted, fontSize: "11px", display: "block" }}>Header Keyword Reference: {k.excelColumnKeyword}</span></div>
                <div style={{ display: "flex", gap: "12px", alignItems: "center" }}>
                  <button onClick={() => toggleHexFavorite(k.id)} style={{ background: "none", border: "none", color: k.isHexFavorite ? theme.accentGold : theme.textMuted, fontSize: "18px", cursor: "pointer" }}>{k.isHexFavorite ? "★" : "☆"}</button>
                  <button onClick={() => {
                    const filteredKpis = kpiRegistry.filter(item => item.id !== k.id);
                    setKpiRegistry(filteredKpis);
                    localStorage.setItem("bluelock_kpi_registry", JSON.stringify(filteredKpis));
                  }} style={{ background: "none", border: "none", color: theme.danger, fontWeight: "bold", cursor: "pointer" }}>✕</button>
                </div>
              </div>
            ))}
          </div>

          <h4 style={{ color: theme.accentGold, marginBottom: "12px" }}>＋ Map and Ingest New Column Metric Option</h4>
          <form onSubmit={(e) => {
            e.preventDefault();
            if (!newKpiName || !selectedHeaderFromDropdown) return;
            const updatedRegistry = [...kpiRegistry, { id: `custom_${Date.now()}`, displayName: newKpiName, excelColumnKeyword: selectedHeaderFromDropdown, category: newKpiCat, isHexFavorite: false }];
            setKpiRegistry(updatedRegistry);
            localStorage.setItem("bluelock_kpi_registry", JSON.stringify(updatedRegistry));
            setNewKpiName("");
          }} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr auto", gap: "16px", alignItems: "end", backgroundColor: theme.bg, padding: "20px", borderRadius: "8px", border: `1px solid ${theme.border}` }}>
            <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
              <label style={{ fontSize: "12px", color: theme.accent }}>Custom Dashboard Label</label>
              <input type="text" placeholder="e.g. CSAT Score %" value={newKpiName} onChange={e => setNewKpiName(e.target.value)} style={{ padding: "10px", backgroundColor: theme.surface, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px" }} />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
              <label style={{ fontSize: "12px", color: theme.accent }}>Detected Column from File Streams</label>
              <select value={selectedHeaderFromDropdown} onChange={e => setSelectedHeaderFromDropdown(e.target.value)} style={{ padding: "10px", backgroundColor: theme.surface, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px" }}>
                {allDetectedHeaders.length === 0 ? <option>-- Upload Excel to populate variables --</option> : allDetectedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
              <label style={{ fontSize: "12px", color: theme.accent }}>Category Core Pillar</label>
              <select value={newKpiCat} onChange={e => setNewKpiCat(e.target.value as any)} style={{ padding: "10px", backgroundColor: theme.surface, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px" }}>
                <option value="Productivity">Productivity</option>
                <option value="Quality">Quality</option>
                <option value="Compliance">Compliance</option>
              </select>
            </div>
            <button type="submit" style={{ backgroundColor: theme.accent, color: theme.bg, padding: "12px 24px", border: "none", borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}>Inject Option</button>
          </form>
        </div>
      )}

      {/* ==========================================
          ⚔️ TAB 3: MILESTONES & HISTORY LEDGER PANEL
         ========================================== */}
      {activeTab === "actionPlan" && (
        <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "32px" }}>
          <h2 style={{ color: theme.accentGold, margin: 0 }}>⚔️ 6-Week Action Calibration Roadmap</h2>
          <div style={{ display: "flex", flexDirection: "column", gap: "20px", marginTop: "24px" }}>
            {milestones.map((m, idx) => (
              <div key={idx} style={{ backgroundColor: theme.bg, border: `1px solid ${theme.border}`, padding: "24px", borderRadius: "12px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div><strong>{m.week}</strong>: {m.title} <p style={{ margin: "4px 0 0 0", fontSize: "13px", color: theme.textMuted }}>{m.desc}</p></div>
                <button onClick={() => setMilestones(milestones.map((item, i) => i === idx ? { ...item, completed: !item.completed } : item))} style={{ padding: "8px 16px", backgroundColor: m.completed ? theme.success : "transparent", color: m.completed ? theme.bg : theme.accentGold, border: `1px solid ${m.completed ? theme.success : theme.accentGold}`, borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}>{m.completed ? "✓ Stage Unlocked" : "⬡ Awaken Stage"}</button>
              </div>
            ))}
          </div>

          <div style={{ marginTop: "40px", paddingTop: "24px", borderTop: `1px solid ${theme.border}` }}>
            <h3 style={{ color: theme.accent }}>📜 Local Extracted Time Segments History Ledger</h3>
            {!selectedAgentLogin ? <p style={{ color: theme.textMuted }}>Choose an active handle above to map timelines.</p> : (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "20px", marginTop: "16px" }}>
                {(masterHistoryMap[selectedAgentLogin] || []).map((run, idx) => (
                  <div key={idx} style={{ backgroundColor: theme.bg, padding: "20px", borderRadius: "8px", border: `1px solid ${theme.border}`, borderLeft: `4px solid ${theme.accent}` }}>
                    <h4 style={{ margin: "0 0 10px 0", color: theme.accentGold }}>{run.weekId} Run Upload Dataset Mappings</h4>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px", fontSize: "12px" }}>
                      {kpiRegistry.map(kpi => {
                        const val = run.metrics[kpi.id] || run.metrics[kpi.excelColumnKeyword] || "—";
                        const formatted = typeof val === "number" ? (val < 1 ? `${(val * 100).toFixed(1)}%` : val.toFixed(1)) : val;
                        return <div key={kpi.id} style={{ color: theme.textMuted }}>{kpi.displayName}: <span style={{ color: theme.textLight, fontWeight: "bold" }}>{formatted}</span></div>;
                      })}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

    </div>
  );
}
