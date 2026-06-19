"use client";

import React, { useState, useRef } from "react";
import * as XLSX from "xlsx";

// ==========================================
// 🗂️ 1. ARCHITECTURE TYPE BLUEPRINTS
// ==========================================
interface KPIMapping {
  id: string;
  displayName: string;
  excelColumnKeyword: string;
  category: "Quality" | "Productivity" | "Compliance";
  isActive: boolean;
}

interface AgentProfile {
  id: string;
  login: string;
  market: "FR" | "BE" | "UK" | "MA" | "DE" | "ES";
  primarySkill: string;
  backupSkill: string;
  teamPlacement: string;
  overallRating: number;
}

interface DayShift {
  startTime: string;
  endTime: string;
  lunchStart: string;
  lunchEnd: string;
  isShortHourDay: boolean;
}

interface WeeklySchedule {
  weekId: string;
  days: Record<string, DayShift>;
}

// ==========================================
// 🎛️ 2. HIGH-END MODERN NEON STYLES (Argon/Black UI Theme)
// ==========================================
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
  // ==========================================
  // 💾 3. APP ENGINE CENTRAL STATE MANAGEMENT
  // ==========================================
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [selectedAgentId, setSelectedAgentId] = useState<string | null>(null);
  const [activeWeek, setActiveWeek] = useState<string>("Week 25");
  const [selectedCategoryFilter, setSelectedCategoryFilter] = useState<string>("All");
  const [rawExcelMatrix, setRawExcelMatrix] = useState<any[][] | null>(null);
  
  // Parameter State to Control, Add, or Delete KPIs on the Fly
  const [kpiRegistry, setKpiRegistry] = useState<KPIMapping[]>([
    { id: "missed_overall", displayName: "Overall Missed Rate", excelColumnKeyword: "Overall Missed Contact Rate", category: "Productivity", isActive: true },
    { id: "missed_phone", displayName: "Phone Missed Rate", excelColumnKeyword: "Phone Missed Contact Rate", category: "Productivity", isActive: true },
    { id: "missed_chat", displayName: "Chat Missed Rate", excelColumnKeyword: "Chat Missed Contact Rate", category: "Productivity", isActive: true },
    { id: "missed_email", displayName: "Email Missed Rate", excelColumnKeyword: "Email Missed Contact Rate", category: "Productivity", isActive: true },
    { id: "ring_time", displayName: "Ring Time (Min)", excelColumnKeyword: "Ring Time (Min)", category: "Productivity", isActive: true }
  ]);

  // Dynamic States for Adding New Parameter KPIs
  const [newKpiName, setNewKpiName] = useState("");
  const [newKpiKeyword, setNewKpiKeyword] = useState("");
  const [newKpiCat, setNewKpiCat] = useState<"Quality" | "Productivity" | "Compliance">("Productivity");

  // Scheduling State Engine
  const [shortDay, setShortDay] = useState<string>("Friday");
  const [agentSchedules, setAgentSchedules] = useState<Record<string, Record<string, WeeklySchedule>>>({});

  // Seed Data: Core Egoist Roster
  const [agents] = useState<AgentProfile[]>([
    { id: "1", login: "aeeljond", market: "FR", primarySkill: "Tech Support", backupSkill: "Billing", teamPlacement: "Striker", overallRating: 92 },
    { id: "2", login: "meskani", market: "BE", primarySkill: "Retention", backupSkill: "Tech Support", teamPlacement: "Midfielder", overallRating: 87 },
    { id: "3", login: "anadamw", market: "UK", primarySkill: "Billing", backupSkill: "Retention", teamPlacement: "Winger", overallRating: 81 }
  ]);

  const currentAgent = agents.find(a => a.id === selectedAgentId);

  // ==========================================
  // 📥 4. PARSING RAW ADAPTIVE EXCEL DATA STREAMS
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
        setRawExcelMatrix(parsedRows);
      } catch (err) {
        alert("Error parsing operational data extraction worksheet template.");
      }
    };
    reader.readAsBinaryString(file);
  };

  // ==========================================
  // 🧮 5. 44-HOUR SHIFT ALLOCATION MATHEMATICS
  // ==========================================
  const fetchOrInitSchedule = (agentId: string, week: string): WeeklySchedule => {
    if (agentSchedules[agentId]?.[week]) return agentSchedules[agentId][week];
    return {
      weekId: week,
      days: {
        Monday: { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false },
        Tuesday: { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false },
        Wednesday: { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false },
        Thursday: { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false },
        Friday: { startTime: "08:00", endTime: "17:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: true }
      }
    };
  };

  const calculateHoursBalance = (schedule: WeeklySchedule): number => {
    let totalMins = 0;
    Object.values(schedule.days).forEach(d => {
      const [sh, sm] = d.startTime.split(":").map(Number);
      const [eh, em] = d.endTime.split(":").map(Number);
      const [lh, lm] = d.lunchStart.split(":").map(Number);
      const [leh, lem] = d.lunchEnd.split(":").map(Number);
      totalMins += ((eh * 60 + em) - (sh * 60 + sm)) - ((leh * 60 + lem) - (lh * 60 + lm));
    });
    return Math.round((totalMins / 60) * 10) / 10;
  };

  const handleShiftCellAdjustment = (day: string, field: keyof DayShift, val: any) => {
    if (!selectedAgentId) return;
    const currentSched = fetchOrInitSchedule(selectedAgentId, activeWeek);
    const updatedDays = { ...currentSched.days, [day]: { ...currentSched.days[day], [field]: val } };
    
    setAgentSchedules({
      ...agentSchedules,
      [selectedAgentId]: { ...(agentSchedules[selectedAgentId] || {}), [activeWeek]: { weekId: activeWeek, days: updatedDays } }
    });
  };

  const activeSchedule = selectedAgentId ? fetchOrInitSchedule(selectedAgentId, activeWeek) : null;
  const netHoursTotal = activeSchedule ? calculateHoursBalance(activeSchedule) : 0;

  // Dynamic cell parsing mapping logic from excel structure row index mappings
  const grabAgentExcelMetric = (login: string, columnKeyword: string) => {
    if (!rawExcelMatrix || rawExcelMatrix.length < 5) return null;
    const kpiRowLabels = rawExcelMatrix[4];
    const columnIdx = kpiRowLabels.findIndex((header: any) => String(header).trim() === columnKeyword);
    if (columnIdx === -1) return null;

    const matchedRow = rawExcelMatrix.find(r => r && String(r[0]).trim().toLowerCase() === login.trim().toLowerCase());
    return matchedRow ? matchedRow[columnIdx] : null;
  };

  return (
    <div style={{ backgroundColor: theme.bg, color: theme.textLight, minHeight: "100vh", padding: "32px", fontFamily: "sans-serif" }}>
      
      {/* 👑 TOP BANNER COMPONENT */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `2px solid ${theme.border}`, paddingBottom: "20px", marginBottom: "32px" }}>
        <div>
          <h1 style={{ fontSize: "26px", fontWeight: "900", color: theme.accent, letterSpacing: "1.5px", margin: 0 }}>PROJECT: BLUELOCKMA</h1>
          <p style={{ fontSize: "12px", color: theme.textMuted, marginTop: "4px" }}>CORE SYSTEM CONTROL & EXTRACTOR TABLE INTERFACES</p>
        </div>
        <div style={{ display: "flex", gap: "12px" }}>
          <button onClick={() => fileInputRef.current?.click()} style={{ backgroundColor: theme.surfaceLight, color: theme.accent, border: `1px solid ${theme.border}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold", cursor: "pointer" }}>📥 Upload Data Sheet</button>
          <input type="file" ref={fileInputRef} onChange={handleExcelUploadStream} accept=".xlsx, .xls" style={{ display: "none" }} />
          <select onChange={(e) => setSelectedAgentId(e.target.value || null)} style={{ backgroundColor: theme.surface, color: theme.accent, border: `2px solid ${theme.accent}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold" }}>
            <option value="">-- TEAM VIEW OVERVIEW --</option>
            {agents.map(a => <option key={a.id} value={a.id}>@{a.login} [{a.market}]</option>)}
          </select>
        </div>
      </div>

      {/* ==========================================
          🧱 MIDDLE CONTAINER: GAMIFICATION VS CONTROLS
         ========================================== */}
      <div style={{ display: "grid", gridTemplateColumns: "380px 1fr", gap: "32px" }}>
        
        {/* LEFT COLUMN PANEL (FIFA CARD / RADAR / ANIMATION ENGINE) */}
        <div>
          {!currentAgent ? (
            <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "40px", textAlign: "center" }}>
              <div style={{ fontSize: "80px", marginBottom: "16px" }}>🦁🦅</div>
              <h3 style={{ fontSize: "18px", color: theme.accent, margin: 0 }}>GRIFFON MATRICES SQUAD</h3>
              <p style={{ fontSize: "12px", color: theme.textMuted, marginTop: "8px" }}>Select an active Egoist parameter item row target.</p>
            </div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
              {/* FIFA ULTRA CARD INTERFACE RENDERING */}
              <div style={{ background: "linear-gradient(135deg, #16223f 0%, #0d1527 100%)", border: `3px solid ${theme.accentGold}`, borderRadius: "20px", padding: "20px", boxShadow: "0 10px 30px rgba(0,0,0,0.5)" }}>
                <div style={{ display: "flex", justifyContent: "space-between" }}>
                  <div>
                    <span style={{ fontSize: "42px", fontWeight: "900", color: theme.accentGold, display: "block" }}>{currentAgent.overallRating}</span>
                    <span style={{ fontSize: "12px", fontWeight: "bold", color: theme.accent }}>{currentAgent.teamPlacement.toUpperCase()}</span>
                  </div>
                  <span style={{ fontSize: "14px", fontWeight: "bold" }}>🇲🇦 {currentAgent.market}</span>
                </div>
                <div style={{ textAlign: "center", margin: "16px 0" }}><span style={{ fontSize: "52px" }}>👤</span><h2 style={{ fontSize: "22px", margin: "8px 0", color: theme.accentGold }}>{currentAgent.login.toUpperCase()}</h2></div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px", borderTop: `1px solid ${theme.border}`, paddingTop: "12px", fontSize: "13px" }}>
                  {kpiRegistry.slice(0, 4).map(kpi => {
                    const rawScore = grabAgentExcelMetric(currentAgent.login, kpi.excelColumnKeyword);
                    const parsedScore = typeof rawScore === "number" ? (rawScore < 1 ? `${(rawScore * 100).toFixed(0)}` : rawScore.toFixed(0)) : "85";
                    return <div key={kpi.id} style={{ fontWeight: "bold" }}>{parsedScore} <span style={{ color: theme.textMuted, fontSize: "11px" }}>{kpi.displayName.slice(0, 3).toUpperCase()}</span></div>;
                  })}
                </div>
              </div>

              {/* HEX MATRIX SVG BLUE LOCK BLOCKS */}
              <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "20px", textAlign: "center" }}>
                <h4 style={{ margin: "0 0 12px 0", fontSize: "13px", color: theme.accent, textAlign: "left" }}>⬡ EVALUATION RADAR COMPLEX</h4>
                <svg width="160" height="160" viewBox="0 0 200 200" style={{ margin: "0 auto" }}>
                  <polygon points="100,20 170,60 170,140 100,180 30,140 30,60" fill="none" stroke={theme.border} strokeWidth="2" />
                  <polygon points="100,45 150,75 145,120 100,150 50,125 45,70" fill="rgba(0, 240, 255, 0.2)" stroke={theme.accent} strokeWidth="2" />
                </svg>
              </div>
            </div>
          )}
        </div>

        {/* RIGHT COLUMN PANEL (SCHEDULER & DYNAMIC CONFIGURATION CONTROLS) */}
        <div style={{ display: "flex", flexDirection: "column", gap: "32px" }}>
          
          {/* CONTROL PARAMETERS MODULE */}
          <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "24px" }}>
            <h3 style={{ color: theme.accent, margin: "0 0 8px 0", fontSize: "18px" }}>⚙️ bluelockMA Parameters Panel</h3>
            <p style={{ fontSize: "12px", color: theme.textMuted, marginBottom: "16px" }}>Inject, monitor, or drop trackable KPI parameters in real-time without modifying code.</p>

            <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", marginBottom: "20px" }}>
              {kpiRegistry.map(kpi => (
                <div key={kpi.id} style={{ display: "flex", alignItems: "center", gap: "8px", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, padding: "6px 12px", borderRadius: "6px", fontSize: "12px" }}>
                  <span><strong>{kpi.displayName}</strong></span>
                  <button onClick={() => setKpiRegistry(kpiRegistry.filter(k => k.id !== kpi.id))} style={{ background: "none", border: "none", color: theme.danger, cursor: "pointer", fontWeight: "bold" }}>✕</button>
                </div>
              ))}
            </div>

            <form onSubmit={(e) => {
              e.preventDefault();
              if (!newKpiName || !newKpiKeyword) return;
              setKpiRegistry([...kpiRegistry, { id: `custom_${Date.now()}`, displayName: newKpiName, excelColumnKeyword: newKpiKeyword, category: newKpiCat, isActive: true }]);
              setNewKpiName(""); setNewKpiKeyword("");
            }} style={{ display: "flex", gap: "10px", alignItems: "center" }}>
              <input type="text" placeholder="UI Display Label" value={newKpiName} onChange={e => setNewKpiName(e.target.value)} style={{ padding: "8px", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px", fontSize: "12px" }} />
              <input type="text" placeholder="Exact Column Name" value={newKpiKeyword} onChange={e => setNewKpiKeyword(e.target.value)} style={{ padding: "8px", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px", fontSize: "12px" }} />
              <select value={newKpiCat} onChange={e => setNewKpiCat(e.target.value as any)} style={{ padding: "8px", backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", borderRadius: "4px", fontSize: "12px" }}>
                <option value="Productivity">Productivity</option>
                <option value="Quality">Quality</option>
                <option value="Compliance">Compliance</option>
              </select>
              <button type="submit" style={{ backgroundColor: theme.accent, color: theme.bg, padding: "8px 16px", border: "none", borderRadius: "4px", fontWeight: "bold", cursor: "pointer", fontSize: "12px" }}>＋ Add Parameter</button>
            </form>
          </div>

          {/* LABOR OPERATIONS GRID PLANNER */}
          <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "24px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}>
              <div><h3 style={{ fontSize: "18px", fontWeight: "900", margin: 0 }}>📆 Manual Labor Roster Grid</h3><p style={{ fontSize: "12px", color: theme.textMuted, margin: 0 }}>44-hour threshold assignment engine tracking matrix constraints</p></div>
              {activeSchedule && <div style={{ backgroundColor: theme.surfaceLight, border: `1px solid ${theme.border}`, padding: "6px 12px", borderRadius: "6px", textAlign: "right" }}><span style={{ fontSize: "11px", color: theme.textMuted, display: "block" }}>Net Duration</span><span style={{ fontSize: "16px", fontWeight: "900", color: netHoursTotal === 44 ? theme.success : theme.accentGold }}>{netHoursTotal} / 44 Hours</span></div>}
            </div>

            {!selectedAgentId ? <div style={{ padding: "24px", color: theme.textMuted, border: `1px dashed ${theme.border}`, borderRadius: "8px", textAlign: "center" }}>Select an egoist profile target value above to manipulate shift schedules manually.</div> : (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
                <thead><tr style={{ borderBottom: `2px solid ${theme.border}`, color: theme.textMuted, textAlign: "left" }}><th style={{ padding: "10px" }}>Weekday</th><th style={{ padding: "10px" }}>Shift Start</th><th style={{ padding: "10px" }}>Shift End</th><th style={{ padding: "10px" }}>Lunch Start</th><th style={{ padding: "10px" }}>Lunch End</th><th style={{ padding: "10px", textAlign: "center" }}>Short Hour Day (-1h)</th></tr></thead>
                <tbody>
                  {["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].map(day => {
                    const dShift = activeSchedule?.days[day] || { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false };
                    return (
                      <tr key={day} style={{ borderBottom: `1px solid ${theme.border}` }}><td style={{ padding: "10px", fontWeight: "bold", color: theme.accent }}>{day}</td>
                        <td style={{ padding: "10px" }}><input type="text" value={dShift.startTime} onChange={e => handleShiftCellAdjustment(day, "startTime", e.target.value)} style={{ backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", width: "60px", padding: "4px", textAlign: "center", borderRadius: "4px" }} /></td>
                        <td style={{ padding: "10px" }}><input type="text" value={dShift.endTime} onChange={e => handleShiftCellAdjustment(day, "endTime", e.target.value)} style={{ backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", width: "60px", padding: "4px", textAlign: "center", borderRadius: "4px" }} /></td>
                        <td style={{ padding: "10px" }}><input type="text" value={dShift.lunchStart} onChange={e => handleShiftCellAdjustment(day, "lunchStart", e.target.value)} style={{ backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", width: "60px", padding: "4px", textAlign: "center", borderRadius: "4px" }} /></td>
                        <td style={{ padding: "10px" }}><input type="text" value={dShift.lunchEnd} onChange={e => handleShiftCellAdjustment(day, "lunchEnd", e.target.value)} style={{ backgroundColor: theme.bg, border: `1px solid ${theme.border}`, color: "#fff", width: "60px", padding: "4px", textAlign: "center", borderRadius: "4px" }} /></td>
                        <td style={{ padding: "10px", textAlign: "center" }}><input type="checkbox" checked={dShift.isShortHourDay} onChange={e => handleShiftCellAdjustment(day, "isShortHourDay", e.target.checked)} style={{ cursor: "pointer" }} /></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>

        </div>
      </div>

      {/* ==========================================
          📊 6. THE FLUID MATRIX TABLES (Bottom Row Matrix)
         ========================================== */}
      <div style={{ backgroundColor: theme.surface, border: `1px solid ${theme.border}`, borderRadius: "16px", padding: "24px", marginTop: "32px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}>
          <div><h3 style={{ color: theme.accent, margin: 0 }}>📊 3D Fluid Matrix Grid</h3><p style={{ fontSize: "12px", color: theme.textMuted, margin: "2px 0 0 0" }}>Cross-referenced tracking logs synchronized dynamically with parameters configuration mappings.</p></div>
          <div style={{ display: "flex", gap: "6px" }}>
            {["All", "Productivity", "Quality", "Compliance"].map(cat => (
              <button key={cat} onClick={() => setSelectedCategoryFilter(cat)} style={{ padding: "6px 12px", backgroundColor: selectedCategoryFilter === cat ? theme.accent : theme.bg, color: selectedCategoryFilter === cat ? theme.bg : "#fff", border: `1px solid ${theme.border}`, borderRadius: "4px", cursor: "pointer", fontSize: "12px", fontWeight: "bold" }}>{cat}</button>
            ))}
          </div>
        </div>

        {!rawExcelMatrix ? <div style={{ padding: "40px", border: `2px dashed ${theme.border}`, borderRadius: "8px", textAlign: "center", color: theme.textMuted }}>Upload a compiled operational evaluation file spreadsheet via the top command button to populate grid structures.</div> : (
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
                {rawExcelMatrix.slice(5).filter(row => row && row[0]).map((row, idx) => (
                  <tr key={idx} style={{ borderBottom: `1px solid ${theme.border}`, backgroundColor: idx % 2 === 0 ? theme.surface : theme.bg }}>
                    <td style={{ padding: "10px", fontWeight: "bold", position: "sticky", left: 0, backgroundColor: idx % 2 === 0 ? theme.surface : theme.bg }}>@{row[0]}</td>
                    {kpiRegistry.filter(k => selectedCategoryFilter === "All" || k.category === selectedCategoryFilter).map(kpi => {
                      const excelLabels = rawExcelMatrix[4];
                      const targetIdx = excelLabels.findIndex((header: any) => String(header).trim() === kpi.excelColumnKeyword);
                      const rawVal = targetIdx !== -1 ? row[targetIdx] : null;
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
  );
}
