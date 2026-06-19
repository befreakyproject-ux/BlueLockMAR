"use client";

import React, { useState, useRef } from "react";
import * as XLSX from "xlsx";

// ==========================================
// 🗂️ STRUCTURE TYPE BLUEPRINTS
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

interface HistoricalData {
  weekId: string;
  metrics: Record<string, number | string>;
}

const theme = {
  bg: "#050811",
  surface: "#0d1527",
  surfaceLight: "#16223f",
  accent: "#00f0ff",     // Neon Blue Lock Cyan
  accentGold: "#e2c044", // FUT Card Premium Gold
  border: "#1e2f55",
  textMuted: "#64748b",
  textLight: "#f8fafc",
  success: "#34d399",
  danger: "#ef4444"
};

export default function BlueLockMA() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  // App Engines State
  const [selectedAgentLogin, setSelectedAgentLogin] = useState<string | null>(null);
  const [selectedCategoryFilter, setSelectedCategoryFilter] = useState<string>("All");
  
  // Extracted Data Holders
  const [discoveredAgents, setDiscoveredAgents] = useState<DynamicAgent[]>([]);
  const [parsedDataRows, setParsedDataRows] = useState<any[][]>([]);
  const [headerKeywordsRow, setHeaderKeywordsRow] = useState<string[]>([]);
  const [availableWeeks, setAvailableWeeks] = useState<string[]>([]);
  const [selectedWeekFilter, setSelectedWeekFilter] = useState<string>("All");

  // Parameter Control State Core
  const [kpiRegistry, setKpiRegistry] = useState<KPIMapping[]>([
    { id: "missed_overall", displayName: "Overall Missed Rate", excelColumnKeyword: "Overall Missed Contact Rate", category: "Productivity" },
    { id: "missed_phone", displayName: "Phone Missed Rate", excelColumnKeyword: "Phone Missed Contact Rate", category: "Productivity" },
    { id: "missed_chat", displayName: "Chat Missed Rate", excelColumnKeyword: "Chat Missed Contact Rate", category: "Productivity" },
    { id: "missed_email", displayName: "Email Missed Rate", excelColumnKeyword: "Email Missed Contact Rate", category: "Productivity" },
    { id: "ring_time", displayName: "Ring Time (Min)", excelColumnKeyword: "Ring Time (Min)", category: "Productivity" }
  ]);

  const [newKpiName, setNewKpiName] = useState("");
  const [newKpiKeyword, setNewKpiKeyword] = useState("");
  const [newKpiCat, setNewKpiCat] = useState<"Quality" | "Productivity" | "Compliance">("Productivity");

  // ==========================================
  // 📥 AUTO-ADAPTIVE INGESTION PIPELINE
  // ==========================================
  const handleExcelUploadStream = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const dataBinary = evt.target?.result;
        const workbook = XLSX.read(dataBinary, { type: "binary" });
        const targetSheet = workbook.Sheets[workbook.SheetNames[0]];
        const parsedRows = XLSX.utils.sheet_to_json<any[]>(targetSheet, { header: 1 });

        if (parsedRows.length < 5) return;

        // Auto-detect layout maps
        const subKpiHeaders = parsedRows[4].map((h: any) => String(h || "").trim());
        const weeksRow = parsedRows[3].map((w: any) => String(w || "").trim());

        setHeaderKeywordsRow(subKpiHeaders);
        setParsedDataRows(parsedRows);

        // Discovered Unique Weeks for History Pipeline
        const uniqueWeeks = Array.from(new Set(weeksRow.filter(w => w && w.toLowerCase().includes("week"))));
        setAvailableWeeks(uniqueWeeks);
        if (uniqueWeeks.length > 0) setSelectedWeekFilter(uniqueWeeks[0]);

        // Auto-extract ALL unique agents found in row 5 downward
        const agentList: DynamicAgent[] = [];
        parsedRows.slice(5).forEach((row) => {
          if (row && row[0] && String(row[0]).trim() !== "" && String(row[0]).toLowerCase() !== "agent" && String(row[0]).toLowerCase() !== "total") {
            const loginStr = String(row[0]).trim();
            if (!agentList.some(a => a.login === loginStr)) {
              agentList.push({ login: loginStr, overallRating: Math.floor(Math.random() * (95 - 78 + 1)) + 78 });
            }
          }
        });

        setDiscoveredAgents(agentList);
        if (agentList.length > 0) setSelectedAgentLogin(agentList[0].login);

      } catch (err) {
        alert("Parsing mismatch. Ensure header strings match configuration panels.");
      }
    };
    reader.readAsBinaryString(file);
  };

  // Safe Metric Extraction Handler
  const pullLiveCellData = (agentLogin: string, columnKeyword: string) => {
    if (parsedDataRows.length === 0 || headerKeywordsRow.length === 0) return null;
    
    const targetIdx = headerKeywordsRow.findIndex(h => h.toLowerCase() === columnKeyword.toLowerCase().trim());
    if (targetIdx === -1) return null;

    const matchedRow = parsedDataRows.find(r => r && String(r[0]).trim().toLowerCase() === agentLogin.toLowerCase());
    return matchedRow ? matchedRow[targetIdx] : null;
  };

  // Compile full structural timeline history records for an selected agent
  const compileAgentHistory = (agentLogin: string): HistoricalData[] => {
    if (!parsedDataRows || parsedDataRows.length === 0) return [];
    const weeksRow = parsedDataRows[3];
    const subKpiRow = parsedDataRows[4];
    const matchedRow = parsedDataRows.find(r => r && String(r[0]).trim().toLowerCase() === agentLogin.toLowerCase());

    if (!matchedRow) return [];

    return availableWeeks.map(week => {
      const metricsMap: Record<string, number | string> = {};
      kpiRegistry.forEach(kpi => {
        let cellsIndex = -1;
        // Search cell with identical week container bounds block
        for (let i = 0; i < matchedRow.length; i++) {
          if (String(weeksRow[i]).trim() === week && String(subKpiRow[i]).trim().toLowerCase() === kpi.excelColumnKeyword.toLowerCase()) {
            cellsIndex = i;
            break;
          }
        }
        metricsMap[kpi.id] = cellsIndex !== -1 ? matchedRow[cellsIndex] : "—";
      });
      return { weekId: week, metrics: metricsMap };
    });
  };

  const currentAgentData = discoveredAgents.find(a => a.login === selectedAgentLogin);
  const activeHistoryTimeline = selectedAgentLogin ? compileAgentHistory(selectedAgentLogin) : [];

  return (
    <div style={{ backgroundColor: theme.bg, color: theme.textLight, minHeight: "100vh", padding: "32px", fontFamily: "sans-serif" }}>
      
      {/* 👑 HEAD BAR SECTION */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `2px solid ${theme.border}`, paddingBottom: "20px", marginBottom: "32px" }}>
        <div>
          <h1 style={{ fontSize: "26px", fontWeight: "900", color: theme.accent, letterSpacing: "1.5px", margin: 0 }}>PROJECT: BLUELOCKMA</h1>
          <p style={{ fontSize: "12px", color: theme.textMuted, marginTop: "4px" }}>CORE DYNAMIC ADAPTIVE EXTRACTION MATRIX</p>
        </div>
        <div style={{ display: "flex", gap: "12px" }}>
          <button onClick={() => fileInputRef.current?.click()} style={{ backgroundColor: theme.surfaceLight, color: theme.accent, border: `1px solid ${theme.border}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold", cursor: "pointer" }}>📥 Ingest Master File</button>
          <input type="file" ref={fileInputRef} onChange={handleExcelUploadStream} accept=".xlsx, .xls" style={{ display: "none" }} />
          
          <select value={selectedAgentLogin || ""} onChange={(e) => setSelectedAgentLogin(e.target.value || null)} style={{ backgroundColor: theme.surface, color: theme.accent, border: `2px solid ${theme.accent}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold" }}>
            {discoveredAgents.length === 0 ? <option>-- Awaiting File Ingestion --</option> : discoveredAgents.map(a => <option key={a.login} value={a.login}>@{a.login} [EGOIST]</option>)}
          </select>
        </div>
      </div>

      {/* ==========================================
          🧱 LAYOUT CORE
         ========================================== */}
      <div style={{ display: "grid", gridTemplateColumns: "380px 1fr", gap: "32px" }}>
        
        {/* LEFT COMPONENT COLUMN (DYNAMIC FUT SHOWCASE) */}
        <div>
          {!currentAgentData ? (
            <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "40px", textAlign: "center" }}>
              <div style={{ fontSize: "60px", marginBottom: "16px" }}>⚽</div>
              <h3 style={{ fontSize: "16px", color: theme.accent, margin: 0 }}>ROSTER ENGINE STANDBY</h3>
              <p style={{ fontSize: "12px", color: theme.textMuted, marginTop: "8px" }}>Upload data parameters worksheet to map live profiles.</p>
            </div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
              {/* ADAPTIVE FIFA FUT STAT CARD */}
              <div style={{ background: "linear-gradient(135deg, #16223f 0%, #0d1527 100%)", border: `3px solid ${theme.accentGold}`, borderRadius: "20px", padding: "24px" }}>
                <div style={{ display: "flex", justifyContent: "space-between" }}>
                  <div><span style={{ fontSize: "46px", fontWeight: "900", color: theme.accentGold, display: "block" }}>{currentAgentData.overallRating}</span><span style={{ fontSize: "12px", fontWeight: "bold", color: theme.accent }}>STRIKER</span></div>
                  <span style={{ fontSize: "14px", fontWeight: "bold" }}>🇲🇦 MA</span>
                </div>
                <div style={{ textAlign: "center", margin: "20px 0" }}><span style={{ fontSize: "64px" }}>👤</span><h2 style={{ fontSize: "24px", margin: "8px 0", color: theme.accentGold }}>{currentAgentData.login.toUpperCase()}</h2></div>
                
                {/* Dynamically displays ONLY matching checked KPIs */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px", borderTop: `1px solid ${theme.border}`, paddingTop: "14px", fontSize: "13px" }}>
                  {kpiRegistry.slice(0, 6).map(kpi => {
                    const cellVal = pullLiveCellData(currentAgentData.login, kpi.excelColumnKeyword);
                    const formatted = typeof cellVal === "number" ? (cellVal < 1 ? `${(cellVal * 100).toFixed(0)}%` : cellVal.toFixed(1)) : "—";
                    return <div key={kpi.id} style={{ fontWeight: "bold" }}>{formatted} <span style={{ color: theme.textMuted, fontSize: "11px", display: "block" }}>{kpi.displayName.toUpperCase()}</span></div>;
                  })}
                </div>
              </div>

              {/* TIMELINE HISTORICAL TRACKING BLOCK */}
              <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "20px" }}>
                <h4 style={{ color: theme.accent, margin: "0 0 12px 0", fontSize: "14px" }}>📈 Chronological Timeline Metrics</h4>
                {activeHistoryTimeline.length === 0 ? <p style={{ fontSize: "12px", color: theme.textMuted }}>No metric runs tracked.</p> : (
                  <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
                    {activeHistoryTimeline.map(h => (
                      <div key={h.weekId} style={{ backgroundColor: theme.bg, padding: "10px", borderRadius: "6px", borderLeft: `3px solid ${theme.accentGold}` }}>
                        <div style={{ fontSize: "12px", fontWeight: "bold", color: theme.accent }}>{h.weekId} Run</div>
                        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6px", marginTop: "6px", fontSize: "11px" }}>
                          {Object.entries(h.metrics).map(([kId, val]) => {
                            const lbl = kpiRegistry.find(k => k.id === kId)?.displayName || kId;
                            const fmt = typeof val === "number" ? (val < 1 ? `${(val * 100).toFixed(1)}%` : val.toFixed(1)) : val;
                            return <div key={kId} style={{ color: theme.textLight }}>{lbl}: <span style={{ color: theme.accentGold }}>{fmt}</span></div>;
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

        {/* RIGHT COMPONENT COLUMN (SETTINGS CONTROLS PANEL & FILTERS) */}
        <div style={{ display: "flex", flexDirection: "column", gap: "32px" }}>
          
          {/* PARAMETER CONTROL SETTINGS ENGINE PANEL */}
          <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "24px" }}>
            <h3 style={{ color: theme.accent, margin: "0 0 8px 0", fontSize: "18px" }}>⚙️ bluelockMA Parameters Panel</h3>
            <p style={{ fontSize: "12px", color: theme.textMuted, marginBottom: "16px" }}>Add or delete target metrics. The data matrices will instantly filter and display ONLY the choices verified here.</p>

            {/* Active Param Badges Portfolio */}
            <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", marginBottom: "20px" }}>
              {kpiRegistry.map(kpi => (
                <div key={kpi.id} style={{ display: "flex", alignItems: "center", gap: "8px", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, padding: "6px 12px", borderRadius: "6px", fontSize: "12px" }}>
                  <span style={{ color: theme.accentGold }}>●</span>
                  <span><strong>{kpi.displayName}</strong></span>
                  <button onClick={() => setKpiRegistry(kpiRegistry.filter(k => k.id !== kpi.id))} style={{ background: "none", border: "none", color: theme.danger, cursor: "pointer", fontWeight: "bold" }}>✕</button>
                </div>
              ))}
            </div>

            {/* Injector Form */}
            <form onSubmit={(e) => {
              e.preventDefault();
              if (!newKpiName || !newKpiKeyword) return;
              setKpiRegistry([...kpiRegistry, { id: `custom_${Date.now()}`, displayName: newKpiName, excelColumnKeyword: newKpiKeyword, category: newKpiCat }]);
              setNewKpiName(""); setNewKpiKeyword("");
            }} style={{ display: "flex", gap: "10px" }}>
              <input type="text" placeholder="UI Display Label" value={newKpiName} onChange={e => setNewKpiName(e.target.value)} style={{ padding: "8px", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px", fontSize: "12px", flex: 1 }} />
              <input type="text" placeholder="Exact Excel Column Header Name" value={newKpiKeyword} onChange={e => setNewKpiKeyword(e.target.value)} style={{ padding: "8px", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px", fontSize: "12px", flex: 1 }} />
              <select value={newKpiCat} onChange={e => setNewKpiCat(e.target.value as any)} style={{ padding: "8px", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px", fontSize: "12px" }}>
                <option value="Productivity">Productivity</option>
                <option value="Quality">Quality</option>
                <option value="Compliance">Compliance</option>
              </select>
              <button type="submit" style={{ backgroundColor: theme.accent, color: theme.bg, padding: "8px 16px", border: "none", borderRadius: "4px", fontWeight: "bold", cursor: "pointer", fontSize: "12px" }}>＋ Inject Metric</button>
            </form>
          </div>

          {/* DYNAMIC FLUID DATA MATRIX GRID */}
          <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "24px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}>
              <div><h3 style={{ color: theme.accent, margin: 0 }}>📊 3D Fluid Matrix Grid</h3><p style={{ fontSize: "12px", color: theme.textMuted, margin: "2px 0 0 0" }}>Displaying ONLY parameter choices selected in settings above.</p></div>
              
              <div style={{ display: "flex", gap: "8px" }}>
                {/* Historical Week Filter Sync */}
                {availableWeeks.length > 0 && (
                  <select value={selectedWeekFilter} onChange={(e) => setSelectedWeekFilter(e.target.value)} style={{ backgroundColor: theme.bg, color: theme.accent, border: `1px solid ${theme.border}`, padding: "4px 8px", borderRadius: "4px", fontSize: "12px" }}>
                    <option value="All">All Historical Run Periods</option>
                    {availableWeeks.map(w => <option key={w} value={w}>{w}</option>)}
                  </select>
                )}
                {["All", "Productivity", "Quality", "Compliance"].map(cat => (
                  <button key={cat} onClick={() => setSelectedCategoryFilter(cat)} style={{ padding: "6px 12px", backgroundColor: selectedCategoryFilter === cat ? theme.accent : theme.bg, color: selectedCategoryFilter === cat ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", cursor: "pointer", fontSize: "12px", fontWeight: "bold" }}>{cat}</button>
                ))}
              </div>
            </div>

            {parsedDataRows.length === 0 ? (
              <div style={{ padding: "40px", border: `2px dashed ${theme.border}`, borderRadius: "8px", textAlign: "center", color: theme.textMuted }}>Ingest your spreadsheet matrix via the top button component to feed data models.</div>
            ) : (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px", textAlign: "left" }}>
                  <thead>
                    <tr style={{ backgroundColor: theme.surfaceLight, color: theme.accent, borderBottom: `2px solid ${theme.border}` }}>
                      <th style={{ padding: "12px", position: "sticky", left: 0, backgroundColor: theme.surfaceLight }}>Egoist Handle</th>
                      {kpiRegistry.filter(k => selectedCategoryFilter === "All" || k.category === selectedCategoryFilter).map(kpi => (
                        <th key={kpi.id} style={{ padding: "12px", borderRight: `1px solid ${theme.border}` }}>{kpi.displayName}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {parsedDataRows.slice(5).filter(row => row && row[0] && String(row[0]).toLowerCase() !== "agent").map((row, idx) => (
                      <tr key={idx} style={{ borderBottom: `1px solid ${theme.border}`, backgroundColor: idx % 2 === 0 ? theme.surface : theme.bg }}>
                        <td style={{ padding: "10px", fontWeight: "bold", position: "sticky", left: 0, backgroundColor: idx % 2 === 0 ? theme.surface : theme.bg, color: theme.accent }}>@{String(row[0]).trim()}</td>
                        {kpiRegistry.filter(k => selectedCategoryFilter === "All" || k.category === selectedCategoryFilter).map(kpi => {
                          const rawVal = pullLiveCellData(String(row[0]), kpi.excelColumnKeyword);
                          const formattedVal = typeof rawVal === "number" ? (rawVal < 1 ? `${(rawVal * 100).toFixed(2)}%` : rawVal.toFixed(2)) : "—";
                          return <td key={kpi.id} style={{ padding: "10px", borderRight: `1px solid ${theme.border}` }}>{formattedVal}</td>;
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
    </div>
  );
}
