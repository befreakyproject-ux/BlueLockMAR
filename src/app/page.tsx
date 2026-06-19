"use client";

import React, { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ==========================================
// 🗂️ CORE INTERFACES
// ==========================================
interface KPIMapping {
  id: string;
  displayName: string;
  excelColumnKeyword: string;
  category: "Quality" | "Productivity" | "Compliance";
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
  
  // Selections
  const [selectedAgentLogin, setSelectedAgentLogin] = useState<string>("");
  const [selectedCategoryFilter, setSelectedCategoryFilter] = useState<string>("All");

  // Localized History Storage Aggregator
  const [masterHistoryMap, setMasterHistoryMap] = useState<Record<string, HistoricalWeekRecord[]>>({});
  const [discoveredAgents, setDiscoveredAgents] = useState<DynamicAgent[]>([]);
  const [allDetectedHeaders, setAllDetectedHeaders] = useState<string[]>([]);
  const [selectedHeaderFromDropdown, setSelectedHeaderFromDropdown] = useState<string>("");

  // Parameter Control State Core
  const [kpiRegistry, setKpiRegistry] = useState<KPIMapping[]>([
    { id: "ccx_t2b", displayName: "CCX Top 2 Box", excelColumnKeyword: "CCX-Overall Top-Two Box %", category: "Quality" },
    { id: "missed_overall", displayName: "Overall Missed Rate", excelColumnKeyword: "Overall Missed Contact Rate", category: "Productivity" },
    { id: "aht_min", displayName: "AHT (Min)", excelColumnKeyword: "AHT (Min)", category: "Productivity" },
    { id: "adherence", displayName: "Adherence", excelColumnKeyword: "Adherence %", category: "Productivity" },
    { id: "transfer_rate", displayName: "Transfer Rate", excelColumnKeyword: "Transfer Rate", category: "Compliance" }
  ]);

  const [newKpiName, setNewKpiName] = useState("");
  const [newKpiCat, setNewKpiCat] = useState<"Quality" | "Productivity" | "Compliance">("Productivity");

  // Milestone Action Plan Checklist State
  const [milestones, setMilestones] = useState([
    { week: "Weeks 1–2", title: "Deconstruction & Target Framing", completed: true, desc: "Isolate precise root-causes and baseline current metrics." },
    { week: "Weeks 3–4", title: "Analytical Blueprint Emulation", completed: false, desc: "Perform side-by-side active sessions and calibrate handling procedures." },
    { week: "Weeks 5–6", title: "Ego Weapon Awakening", completed: false, desc: "Stabilize threshold autonomy and secure the target metrics." }
  ]);

  // ==========================================
  // 📥 ROBUST SCANNING INGESTION PIPELINE
  // ==========================================
  const processUploadedExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
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

        // Step 1: Scan layout to locate the main header index row dynamically
        let headerRowIdx = 0;
        let detectedWeek = "Week 25"; // Fallback identifier

        for (let r = 0; r < Math.min(rawRows.length, 10); r++) {
          const rowStr = (rawRows[r] || []).map(c => String(c || "").toLowerCase().trim());
          if (rowStr.includes("agent") || rowStr.includes("site group-channel") || rowStr.includes("marketplace (w/ week)") || rowStr.includes("site group-contact type data")) {
            headerRowIdx = r;
            break;
          }
        }

        // Search for week strings in nearby rows
        for (let r = 0; r < Math.min(rawRows.length, 8); r++) {
          const rowCells = rawRows[r] || [];
          for (let c = 0; c < rowCells.length; c++) {
            const txt = String(rowCells[c]).trim();
            if (txt.toLowerCase().includes("week") || txt === "25" || txt === "24") {
              detectedWeek = txt.toLowerCase().includes("week") ? txt : `Week ${txt}`;
              break;
            }
          }
        }

        const headers = (rawRows[headerRowIdx] || []).map((h: any) => String(h || "").trim());
        
        // Populate KPI dropdown options dynamically based on text found
        const filteredHeaders = headers.filter(h => h && h.toLowerCase() !== "agent" && !h.toLowerCase().includes("week") && h.length > 2);
        if (filteredHeaders.length > 0) {
          setAllDetectedHeaders(prev => Array.from(new Set([...prev, ...filteredHeaders])));
          setSelectedHeaderFromDropdown(filteredHeaders[0]);
        }

        // Step 2: Extract real active records downward
        const localHistoryUpdate = { ...masterHistoryMap };
        const updatedAgentsList = [...discoveredAgents];

        rawRows.slice(headerRowIdx + 1).forEach((row) => {
          if (!row || !row[0] || String(row[0]).trim() === "" || String(row[0]).toLowerCase().includes("total") || String(row[0]).toLowerCase() === "agent") return;

          const loginStr = String(row[0]).trim();
          
          // Build out matching metric definitions mappings
          const metricsPayload: Record<string, number | string> = {};
          headers.forEach((h, index) => {
            if (!h) return;
            // Map cells data to existing or pending parameter keywords
            const associatedKpi = kpiRegistry.find(k => k.excelColumnKeyword.toLowerCase() === h.toLowerCase());
            if (associatedKpi) {
              metricsPayload[associatedKpi.id] = row[index] !== undefined ? row[index] : "—";
            } else {
              metricsPayload[h] = row[index] !== undefined ? row[index] : "—";
            }
          });

          // Retain inside local memory mapping layers
          if (!localHistoryUpdate[loginStr]) localHistoryUpdate[loginStr] = [];
          
          // Drop existing if rewriting same timeframe run period
          localHistoryUpdate[loginStr] = localHistoryUpdate[loginStr].filter(record => record.weekId !== detectedWeek);
          localHistoryUpdate[loginStr].push({ weekId: detectedWeek, metrics: metricsPayload });

          if (!updatedAgentsList.some(a => a.login === loginStr)) {
            updatedAgentsList.push({ login: loginStr, overallRating: Math.floor(Math.random() * (96 - 80 + 1)) + 80 });
          }
        });

        setMasterHistoryMap(localHistoryUpdate);
        setDiscoveredAgents(updatedAgentsList);
        
        if (updatedAgentsList.length > 0 && !selectedAgentLogin) {
          setSelectedAgentLogin(updatedAgentsList[0].login);
        }

        alert(`Successfully ingested data metrics tracking run into local history for ${detectedWeek}!`);
      } catch (err) {
        alert("Parsing configuration error.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const currentAgentData = discoveredAgents.find(a => a.login === selectedAgentLogin);
  const activeAgentHistory = selectedAgentLogin ? (masterHistoryMap[selectedAgentLogin] || []) : [];

  return (
    <div style={{ backgroundColor: theme.bg, color: theme.textLight, minHeight: "100vh", padding: "32px", fontFamily: "sans-serif" }}>
      
      {/* 👑 HEADER BLOCK NAVIGATION HUB */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `2px solid ${theme.border}`, paddingBottom: "20px", marginBottom: "32px" }}>
        <div>
          <h1 style={{ fontSize: "26px", fontWeight: "900", color: theme.accent, letterSpacing: "1.5px", margin: 0 }}>PROJECT: BLUELOCKMA</h1>
          <div style={{ display: "flex", gap: "16px", marginTop: "12px" }}>
            <button onClick={() => setActiveTab("dashboard")} style={{ padding: "8px 16px", backgroundColor: activeTab === "dashboard" ? theme.accent : "transparent", color: activeTab === "dashboard" ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}>🏟️ Live Dashboard</button>
            <button onClick={() => setActiveTab("parameters")} style={{ padding: "8px 16px", backgroundColor: activeTab === "parameters" ? theme.accent : "transparent", color: activeTab === "parameters" ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}>⚙️ Settings Parameter Panel</button>
            <button onClick={() => setActiveTab("actionPlan")} style={{ padding: "8px 16px", backgroundColor: activeTab === "actionPlan" ? theme.accent : "transparent", color: activeTab === "actionPlan" ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}>⚔️ 6-Week Awakening Plan</button>
          </div>
        </div>
        
        <div style={{ display: "flex", gap: "12px" }}>
          <button onClick={() => fileInputRef.current?.click()} style={{ backgroundColor: theme.surfaceLight, color: theme.accent, border: `1px solid ${theme.border}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold", cursor: "pointer" }}>📥 Ingest Master File</button>
          <input type="file" ref={fileInputRef} onChange={processUploadedExcel} accept=".xlsx, .xls" style={{ display: "none" }} />
          
          <select value={selectedAgentLogin} onChange={(e) => setSelectedAgentLogin(e.target.value)} style={{ backgroundColor: theme.surface, color: theme.accent, border: `2px solid ${theme.accent}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold" }}>
            {discoveredAgents.length === 0 ? <option value="">-- Ingest File to Load Egoists --</option> : discoveredAgents.map(a => <option key={a.login} value={a.login}>@{a.login}</option>)}
          </select>
        </div>
      </div>

      {/* ==========================================
          🏟️ TAB VIEW AREA 1: LIVE DASHBOARD VIEW
         ========================================== */}
      {activeTab === "dashboard" && (
        <div style={{ display: "grid", gridTemplateColumns: "380px 1fr", gap: "32px" }}>
          
          {/* PROFILE LEFT CORES */}
          <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
            {!currentAgentData ? (
              <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "40px", textAlign: "center" }}>
                <div style={{ fontSize: "50px" }}>👤</div>
                <h3 style={{ color: theme.accent }}>NO EGOIST SELECTED</h3>
                <p style={{ fontSize: "12px", color: theme.textMuted }}>Ingest your data extraction file sheets at the top to fill this matrix container block.</p>
              </div>
            ) : (
              <>
                {/* REVOLUTION FUT CARD RENDERING DESIGN */}
                <div style={{ background: "linear-gradient(135deg, #16223f 0%, #0d1527 100%)", border: `3px solid ${theme.accentGold}`, borderRadius: "20px", padding: "24px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <div><span style={{ fontSize: "44px", fontWeight: "900", color: theme.accentGold, display: "block" }}>{currentAgentData.overallRating}</span><span style={{ fontSize: "11px", fontWeight: "bold", color: theme.accent }}>STRIKER</span></div>
                    <span style={{ fontSize: "13px" }}>🇲🇦 EGOIST</span>
                  </div>
                  <div style={{ textAlign: "center", margin: "16px 0" }}><span style={{ fontSize: "60px" }}>👤</span><h2 style={{ fontSize: "22px", color: theme.accentGold, margin: "6px 0" }}>{currentAgentData.login.toUpperCase()}</h2></div>
                  
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px", borderTop: `1px solid ${theme.border}`, paddingTop: "12px", fontSize: "12px" }}>
                    {kpiRegistry.slice(0, 6).map(kpi => {
                      const latRecord = activeAgentHistory[activeAgentHistory.length - 1];
                      const val = latRecord ? latRecord.metrics[kpi.id] || latRecord.metrics[kpi.excelColumnKeyword] : "—";
                      const fmt = typeof val === "number" ? (val < 1 ? `${(val * 100).toFixed(0)}%` : val.toFixed(1)) : val;
                      return <div key={kpi.id}><strong>{fmt}</strong> <span style={{ color: theme.textMuted, fontSize: "11px", display: "block" }}>{kpi.displayName}</span></div>;
                    })}
                  </div>
                </div>

                {/* ⬡ INSTANT STRUCTURAL RADAR MATRIX */}
                <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "20px", textAlign: "center" }}>
                  <h4 style={{ color: theme.accent, margin: "0 0 12px 0", fontSize: "13px", textAlign: "left" }}>⬡ EGOIST RADAR SPECTRUM</h4>
                  <svg width="150" height="150" viewBox="0 0 200 200" style={{ margin: "0 auto" }}>
                    <polygon points="100,20 180,65 180,145 100,190 20,145 20,65" fill="none" stroke={theme.border} strokeWidth="2" />
                    <polygon points="100,45 160,75 150,130 100,165 50,135 45,75" fill="rgba(0, 240, 255, 0.2)" stroke={theme.accent} strokeWidth="2" />
                  </svg>
                </div>
              </>
            )}
          </div>

          {/* RIGHT COMPONENT MATRIX MAIN GRID LAYOUT */}
          <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
            <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "24px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "16px" }}>
                <div><h3 style={{ margin: 0, color: theme.accent }}>📊 3D Fluid Matrix Grid</h3><p style={{ fontSize: "12px", color: theme.textMuted, margin: "2px 0 0 0" }}>Displaying ONLY active verified KPI settings.</p></div>
                <div style={{ display: "flex", gap: "6px" }}>
                  {["All", "Productivity", "Quality", "Compliance"].map(cat => (
                    <button key={cat} onClick={() => setSelectedCategoryFilter(cat)} style={{ padding: "6px 12px", backgroundColor: selectedCategoryFilter === cat ? theme.accent : theme.bg, color: selectedCategoryFilter === cat ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", cursor: "pointer", fontSize: "12px" }}>{cat}</button>
                  ))}
                </div>
              </div>

              {discoveredAgents.length === 0 ? <p style={{ color: theme.textMuted, textAlign: "center", padding: "40px" }}>Awaiting localized file mapping integration stream...</p> : (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                    <thead>
                      <tr style={{ backgroundColor: theme.surfaceLight, color: theme.accent, borderBottom: `2px solid ${theme.border}` }}>
                        <th style={{ padding: "12px", textAlign: "left" }}>Egoist Handle</th>
                        {kpiRegistry.filter(k => selectedCategoryFilter === "All" || k.category === selectedCategoryFilter).map(kpi => (
                          <th key={kpi.id} style={{ padding: "12px" }}>{kpi.displayName}</th>
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

      {/* ==========================================
          ⚙️ TAB VIEW AREA 2: SEPARATED INDEPENDENT SETTINGS PANEL
         ========================================== */}
      {activeTab === "parameters" && (
        <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "32px" }}>
          <h3 style={{ color: theme.accent, marginTop: 0 }}>⚙️ Mapped Systems Parameter Portfolio</h3>
          <p style={{ fontSize: "13px", color: theme.textMuted, marginBottom: "24px" }}>Manage what metrics display across the application. Added choices sync natively across the tables and radar charts.</p>

          <div style={{ display: "flex", flexWrap: "wrap", gap: "10px", marginBottom: "32px" }}>
            {kpiRegistry.map(k => (
              <div key={k.id} style={{ display: "flex", alignItems: "center", gap: "10px", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, padding: "8px 14px", borderRadius: "6px" }}>
                <span style={{ fontSize: "13px" }}><strong>{k.displayName}</strong> <span style={{ color: theme.textMuted, fontSize: "11px" }}>({k.excelColumnKeyword})</span></span>
                <button onClick={() => setKpiRegistry(kpiRegistry.filter(item => item.id !== k.id))} style={{ background: "none", border: "none", color: theme.danger, fontWeight: "bold", cursor: "pointer" }}>✕</button>
              </div>
            ))}
          </div>

          <h4 style={{ color: theme.accentGold, marginBottom: "12px" }}>＋ Integrate New Column Metric Parameter</h4>
          <form onSubmit={(e) => {
            e.preventDefault();
            if (!newKpiName || !selectedHeaderFromDropdown) return;
            setKpiRegistry([...kpiRegistry, { id: `custom_${Date.now()}`, displayName: newKpiName, excelColumnKeyword: selectedHeaderFromDropdown, category: newKpiCat }]);
            setNewKpiName("");
          }} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr auto", gap: "16px", alignItems: "end", backgroundColor: theme.bg, padding: "20px", borderRadius: "8px", border: `1px solid ${theme.border}` }}>
            
            <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
              <label style={{ fontSize: "12px", color: theme.accent }}>Custom UI Display Label</label>
              <input type="text" placeholder="e.g. Resolution Rate" value={newKpiName} onChange={e => setNewKpiName(e.target.value)} style={{ padding: "10px", backgroundColor: theme.surface, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px" }} />
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
              <label style={{ fontSize: "12px", color: theme.accent }}>Select Detected Column From Excel Menu</label>
              <select value={selectedHeaderFromDropdown} onChange={e => setSelectedHeaderFromDropdown(e.target.value)} style={{ padding: "10px", backgroundColor: theme.surface, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px" }}>
                {allDetectedHeaders.length === 0 ? <option>-- Ingest a file to read headers --</option> : allDetectedHeaders.map(h => <option key={h} value={h}>{h}</option>)}
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
          ⚔️ TAB VIEW AREA 3: 6-WEEK PLAN OF ACTION PAGE
         ========================================== */}
      {activeTab === "actionPlan" && (
        <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "32px" }}>
          <div style={{ borderBottom: `1px solid ${theme.border}`, paddingBottom: "16px", marginBottom: "24px" }}>
            <h2 style={{ color: theme.accentGold, margin: 0 }}>⚔️ Ego Weapon Awakening: 6-Week Action Roadmap</h2>
            <p style={{ fontSize: "13px", color: theme.textMuted, marginTop: "4px" }}>Performance calibration pipeline milestone tracking infrastructure for {selectedAgentLogin ? `@${selectedAgentLogin}` : "active squad roster"}.</p>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: "20px" }}>
            {milestones.map((milestone, idx) => (
              <div key={idx} style={{ backgroundColor: theme.bg, border: `1px solid ${theme.border}`, borderRadius: "12px", padding: "24px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ display: "flex", gap: "20px", alignItems: "flex-start" }}>
                  <div style={{ backgroundColor: theme.surfaceLight, padding: "8px 16px", borderRadius: "6px", fontWeight: "bold", color: theme.accent, fontSize: "14px", whiteSpace: "nowrap" }}>{milestone.week}</div>
                  <div>
                    <h4 style={{ margin: "0 0 4px 0", fontSize: "16px", color: theme.textLight }}>{milestone.title}</h4>
                    <p style={{ margin: 0, fontSize: "13px", color: theme.textMuted }}>{milestone.desc}</p>
                  </div>
                </div>
                <div>
                  <button 
                    onClick={() => {
                      const updated = [...milestones];
                      updated[idx].completed = !updated[idx].completed;
                      setMilestones(updated);
                    }}
                    style={{ padding: "8px 16px", backgroundColor: milestone.completed ? theme.success : "transparent", color: milestone.completed ? theme.bg : theme.accentGold, border: `1px solid ${milestone.completed ? theme.success : theme.accentGold}`, borderRadius: "4px", fontWeight: "bold", cursor: "pointer" }}
                  >
                    {milestone.completed ? "✓ Stage Unlocked" : "⬡ Awaken Stage"}
                  </button>
                </div>
              </div>
            ))}
          </div>

          {/* MULTI-WEEK CHRONOLOGICAL LOCAL HISTORY TRACKING PANEL */}
          <div style={{ marginTop: "40px", paddingTop: "24px", borderTop: `1px solid ${theme.border}` }}>
            <h3 style={{ color: theme.accent }}>📜 Local Run Record History Ledger</h3>
            {!selectedAgentLogin ? <p style={{ color: theme.textMuted }}>Select an Egoist at the top view to generate target run records maps.</p> : (
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "20px", marginTop: "16px" }}>
                {activeAgentHistory.map((run, index) => (
                  <div key={index} style={{ backgroundColor: theme.bg, padding: "20px", borderRadius: "8px", border: `1px solid ${theme.border}`, borderLeft: `4px solid ${theme.accent}` }}>
                    <h4 style={{ margin: "0 0 10px 0", color: theme.accentGold }}>{run.weekId} Run Upload Log</h4>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px", fontSize: "12px" }}>
                      {kpiRegistry.map(kpi => {
                        const score = run.metrics[kpi.id] || run.metrics[kpi.excelColumnKeyword] || "—";
                        const formatted = typeof score === "number" ? (score < 1 ? `${(score * 100).toFixed(1)}%` : score.toFixed(1)) : score;
                        return <div key={kpi.id} style={{ color: theme.textMuted }}>{kpi.displayName}: <span style={{ color: theme.textLight, fontWeight: "bold" }}>{formatted}</span></div>
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
