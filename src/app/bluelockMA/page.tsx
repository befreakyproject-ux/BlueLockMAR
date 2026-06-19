"use client";

import React, { useState, useEffect } from "react";

// ==========================================
// 🗂️ TYPE DEFINITIONS & BLUEPRINTS
// ==========================================
interface AgentProfile {
  id: string;
  login: string;
  market: "FR" | "BE" | "UK" | "MA" | "DE" | "ES";
  primarySkill: string;
  backupSkill: string;
  teamPlacement: string;
  // FUT/Blue Lock Stat Vectors (0-99)
  overallRating: number;
  quality: number;
  speed: number;
  compliance: number;
  adherence: number;
  resolution: number;
}

interface DayShift {
  startTime: string;
  endTime: string;
  lunchStart: string;
  lunchEnd: string;
  isShortHourDay: boolean; // Toggles the -1 hour rule for the 44h constraint
}

interface WeeklySchedule {
  weekId: string; // e.g., "2026-W26"
  days: Record<string, DayShift>;
}

interface CoachingSlot {
  time: string; // e.g., "09:00 - 10:00"
  status: "Available" | "Booked";
  assignedAgent?: string;
}

// ==========================================
// 🎛️ STYLES & THEME CONFIGURATION (NEON BLUE LOCK)
// ==========================================
const darkTheme = {
  bg: "#050811",
  surface: "#0d1527",
  surfaceLight: "#16223f",
  accent: "#00f0ff", // Neon Blue Lock Cyan
  accentGold: "#e2c044", // FIFA Gold
  border: "#1e2f55",
  textMuted: "#64748b",
  textLight: "#f8fafc"
};

export default function BlueLockMA() {
  // ==========================================
  // 💾 STATE HOOKS
  // ==========================================
  const [selectedAgentId, setSelectedAgentId] = useState<string | null>(null);
  const [activeWeek, setActiveWeek] = useState<string>("2026-W26");
  const [managerAvailabilities, setManagerAvailabilities] = useState<Record<string, string[]>>({
    Monday: ["09:00 - 10:00", "14:00 - 15:00"],
    Tuesday: ["10:00 - 11:00", "16:00 - 17:00"],
    Wednesday: ["11:00 - 12:00"],
    Thursday: ["09:00 - 10:00", "15:00 - 16:00"],
    Friday: ["14:00 - 15:00"]
  });
  
  // Track assigned 1-to-1 slots
  const [bookedSlots, setBookedSlots] = useState<Record<string, Record<string, string>>>({});

  // Seed Data representing the Griffon Squad
  const [agents, setAgents] = useState<AgentProfile[]>([
    { id: "1", login: "imanzaki", market: "FR", primarySkill: "Tech Support", backupSkill: "Billing", teamPlacement: "Striker", overallRating: 92, quality: 94, speed: 92, compliance: 84, adherence: 79, resolution: 88 },
    { id: "2", login: "zghouzgh", market: "BE", primarySkill: "Retention", backupSkill: "Tech Support", teamPlacement: "Midfielder", overallRating: 86, quality: 84, speed: 89, compliance: 80, adherence: 85, resolution: 83 },
    { id: "3", login: "hhoumain", market: "UK", primarySkill: "Billing", backupSkill: "Retention", teamPlacement: "Winger", overallRating: 79, quality: 72, speed: 75, compliance: 82, adherence: 69, resolution: 77 }
  ]);

  // Master Schedule Registry
  const [schedules, setSchedules] = useState<Record<string, Record<string, WeeklySchedule>>>({
    "1": {
      "2026-W26": {
        weekId: "2026-W26",
        days: {
          Monday: { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false },
          Tuesday: { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false },
          Wednesday: { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false },
          Thursday: { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false },
          Friday: { startTime: "08:00", endTime: "17:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: true } // 44h balance check
        }
      }
    }
  });

  const currentAgent = agents.find(a => a.id === selectedAgentId);

  // Initialize fresh shift templates for newly selected agents safely
  const getAgentSchedule = (agentId: string, week: string): WeeklySchedule => {
    if (schedules[agentId]?.[week]) return schedules[agentId][week];
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

  const handleShiftChange = (day: string, field: keyof DayShift, value: any) => {
    if (!selectedAgentId) return;
    const currentSched = getAgentSchedule(selectedAgentId, activeWeek);
    const updatedDays = {
      ...currentSched.days,
      [day]: { ...currentSched.days[day], [field]: value }
    };

    setSchedules({
      ...schedules,
      [selectedAgentId]: {
        ...(schedules[selectedAgentId] || {}),
        [activeWeek]: { weekId: activeWeek, days: updatedDays }
      }
    });
  };

  // ==========================================
  // 🧮 AUTOMATED WORKING HOURS MATHEMATICS ENGINE
  // ==========================================
  const calculateTotalWeeklyHours = (schedule: WeeklySchedule): number => {
    let totalMinutes = 0;
    Object.values(schedule.days).forEach(day => {
      const [sh, sm] = day.startTime.split(":").map(Number);
      const [eh, em] = day.endTime.split(":").map(Number);
      const [lh, lm] = day.lunchStart.split(":").map(Number);
      const [leh, lem] = day.lunchEnd.split(":").map(Number);

      const grossMinutes = (eh * 60 + em) - (sh * 60 + sm);
      const lunchMinutes = (leh * 60 + lem) - (lh * 60 + lm);
      totalMinutes += (grossMinutes - lunchMinutes);
    });
    return Math.round((totalMinutes / 60) * 10) / 10;
  };

  const activeSchedule = selectedAgentId ? getAgentSchedule(selectedAgentId, activeWeek) : null;
  const computedHours = activeSchedule ? calculateTotalWeeklyHours(activeSchedule) : 0;

  return (
    <div style={{ backgroundColor: darkTheme.bg, color: darkTheme.textLight, minHeight: "100vh", padding: "32px", fontFamily: "'Segoe UI', Roboto, sans-serif" }}>
      
      {/* 👑 GLOBAL HEADER APPLICATION BAR */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `2px solid ${darkTheme.border}`, paddingBottom: "20px", marginBottom: "32px" }}>
        <div>
          <h1 style={{ fontSize: "28px", fontWeight: "900", color: darkTheme.accent, letterSpacing: "1.5px", margin: 0 }}>PROJECT: BLUELOCKMA</h1>
          <p style={{ fontSize: "12px", color: darkTheme.textMuted, margin: "4px 0 0 0" }}>TMSW GAMIFIED DATA & PERFORMANCE HUBS</p>
        </div>

        {/* Global Workspace Controls */}
        <div style={{ display: "flex", gap: "16px" }}>
          <select 
            value={activeWeek} 
            onChange={(e) => setActiveWeek(e.target.value)}
            style={{ backgroundColor: darkTheme.surface, color: "#fff", border: `1px solid ${darkTheme.border}`, padding: "10px 16px", borderRadius: "6px", fontWeight: "bold" }}
          >
            <option value="2026-W26">Week 26 (Current)</option>
            <option value="2026-W27">Week 27</option>
            <option value="2026-W28">Week 28</option>
          </select>

          <select 
            onChange={(e) => setSelectedAgentId(e.target.value || null)}
            style={{ backgroundColor: darkTheme.surface, color: darkTheme.accent, border: `2px solid ${darkTheme.accent}`, padding: "10px 20px", borderRadius: "6px", fontWeight: "900" }}
          >
            <option value="">-- TEAM VISION (GRIFFON OVERVIEW) --</option>
            {agents.map(a => <option key={a.id} value={a.id}>@{a.login} [{a.market}]</option>)}
          </select>
        </div>
      </div>

      {/* ==========================================
          🧱 MASTER MAIN MULTI-PANEL DISPATCHER
         ========================================== */}
      <div style={{ display: "grid", gridTemplateColumns: "400px 1fr", gap: "32px" }}>
        
        {/* ⚽ LEFT ASPECT PANEL: FIFA RADAR & BLUELOCK PROFILE CORES */}
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
          {!currentAgent ? (
            /* 🦅 SYSTEM DEFAULT TEAM CREST: GRIFFON CREST INTERFACE */
            <div style={{ backgroundColor: darkTheme.surface, border: `1px solid ${darkTheme.border}`, borderRadius: "16px", width: "100%", padding: "40px", textAlign: "center", boxShadow: "0 20px 40px rgba(0,0,0,0.4)" }}>
              <div style={{ fontSize: "96px", marginBottom: "20px", filter: "drop-shadow(0 0 15px rgba(0,240,255,0.4))" }}>🦁🦅</div>
              <h2 style={{ fontSize: "22px", fontWeight: "900", color: darkTheme.accent, letterSpacing: "1px" }}>GRIFFON SQUAD</h2>
              <p style={{ fontSize: "13px", color: darkTheme.textMuted, maxWidth: "280px", margin: "12px auto 0 auto", lineHeight: "1.6" }}>
                Multi-market operations pipeline active. Select an agent to slide open their personalized FIFA Ultimate Skill Card.
              </p>
            </div>
          ) : (
            /* 🎴 ULTIMATE TEAM PLAYER CARD MODAL (Ref: image_3cdf28.jpg) */
            <div style={{ display: "flex", flexDirection: "column", gap: "24px", width: "100%" }}>
              <div style={{ 
                background: "linear-gradient(135deg, #1e293b 0%, #0f172a 100%)", 
                border: `3px solid ${darkTheme.accentGold}`,
                borderRadius: "24px", 
                width: "100%", 
                padding: "24px", 
                color: "#fff",
                boxShadow: `0 15px 35px rgba(226, 192, 68, 0.15)`,
                position: "relative"
              }}>
                {/* Upper Metrics Badging */}
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <span style={{ fontSize: "48px", fontWeight: "900", display: "block", color: darkTheme.accentGold, lineHeight: "1" }}>
                      {currentAgent.overallRating}
                    </span>
                    <span style={{ fontSize: "14px", fontWeight: "800", color: darkTheme.accent, tracking: "1px" }}>
                      {currentAgent.teamPlacement.toUpperCase()}
                    </span>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: "14px", fontWeight: "900", color: "#6ee7b7" }}>🇲🇦 {currentAgent.market}</div>
                    <div style={{ fontSize: "11px", color: darkTheme.textMuted, marginTop: "2px" }}>{currentAgent.primarySkill}</div>
                  </div>
                </div>

                {/* Main Identity Block */}
                <div style={{ textAlign: "center", margin: "20px 0" }}>
                  <div style={{ fontSize: "64px", margin: "10px 0" }}>👤</div>
                  <h2 style={{ fontSize: "24px", fontWeight: "900", letterSpacing: "1px", margin: 0, color: darkTheme.accentGold }}>
                    {currentAgent.login.toUpperCase()}
                  </h2>
                </div>

                {/* FIFA Stats Spread Grid */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px 24px", borderTop: `1px solid ${darkTheme.border}`, paddingTop: "16px", paddingX: "16px" }}>
                  <div style={{ fontSize: "16px", fontWeight: "bold" }}>{currentAgent.quality} <span style={{ color: darkTheme.textMuted, fontSize: "12px", fontWeight: "normal" }}>QLY</span></div>
                  <div style={{ fontSize: "16px", fontWeight: "bold" }}>{currentAgent.speed} <span style={{ color: darkTheme.textMuted, fontSize: "12px", fontWeight: "normal" }}>SPD</span></div>
                  <div style={{ fontSize: "16px", fontWeight: "bold" }}>{currentAgent.compliance} <span style={{ color: darkTheme.textMuted, fontSize: "12px", fontWeight: "normal" }}>CMP</span></div>
                  <div style={{ fontSize: "16px", fontWeight: "bold" }}>{currentAgent.adherence} <span style={{ color: darkTheme.textMuted, fontSize: "12px", fontWeight: "normal" }}>ADH</span></div>
                  <div style={{ fontSize: "16px", fontWeight: "bold" }}>{currentAgent.resolution} <span style={{ color: darkTheme.textMuted, fontSize: "12px", fontWeight: "normal" }}>RES</span></div>
                  <div style={{ fontSize: "16px", fontWeight: "bold" }}>90 <span style={{ color: darkTheme.textMuted, fontSize: "12px", fontWeight: "normal" }}>FDB</span></div>
                </div>
              </div>

              {/* ⬡ BLUE LOCK EGOIST RADAR CANVAS (Ref: image_3d4519.png) */}
              <div style={{ backgroundColor: darkTheme.surface, border: `1px solid ${darkTheme.border}`, borderRadius: "16px", padding: "20px", width: "100%" }}>
                <h3 style={{ fontSize: "14px", fontWeight: "900", color: darkTheme.accent, margin: "0 0 12px 0" }}>⬡ BLUE LOCK RADAR COMPLEX</h3>
                <div style={{ display: "flex", justifyContent: "center", margin: "16px 0" }}>
                  {/* Hexagon SVG Visualization */}
                  <svg width="200" height="200" viewBox="0 0 200 200">
                    <polygon points="100,20 170,60 170,140 100,180 30,140 30,60" fill="none" stroke={darkTheme.border} strokeWidth="2" />
                    <line x1="100" y1="100" x2="100" y2="20" stroke={darkTheme.border} />
                    <line x1="100" y1="100" x2="170" y2="60" stroke={darkTheme.border} />
                    <line x1="100" y1="100" x2="170" y2="140" stroke={darkTheme.border} />
                    <line x1="100" y1="100" x2="100" y2="180" stroke={darkTheme.border} />
                    <line x1="100" y1="100" x2="30" y2="140" stroke={darkTheme.border} />
                    <line x1="100" y1="100" x2="30" y2="60" stroke={darkTheme.border} />
                    {/* Simulated Filled Adaptive Profile */}
                    <polygon points="100,45 150,75 145,120 100,150 50,125 45,70" fill="rgba(0, 240, 255, 0.25)" stroke={darkTheme.accent} strokeWidth="2" />
                  </svg>
                </div>
                <div style={{ textAlign: "center", fontSize: "12px", color: darkTheme.textMuted }}>
                  Evaluated Class Status: <span style={{ color: "#fff", fontWeight: "bold" }}>RANK S EGOIST</span>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* 💻 RIGHT ASPECT PANEL: SHIFT SCHEDULER & AVAILABILITY SLOTS */}
        <div style={{ display: "flex", flexDirection: "column", gap: "32px" }}>
          
          {/* MANUAL HOURS WORK WEEK ENGINE */}
          <div style={{ backgroundColor: darkTheme.surface, border: `1px solid ${darkTheme.border}`, borderRadius: "16px", padding: "24px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "20px" }}>
              <div>
                <h3 style={{ fontSize: "18px", fontWeight: "900", margin: 0 }}>📆 Weekly Labor Allocator</h3>
                <p style={{ fontSize: "12px", color: darkTheme.textMuted, margin: "2px 0 0 0" }}>Enforcing custom 44-hour baseline working constraints</p>
              </div>
              {activeSchedule && (
                <div style={{ backgroundColor: computedHours === 44 ? "rgba(52, 211, 153, 0.1)" : "rgba(226, 192, 68, 0.1)", border: `1px solid ${computedHours === 44 ? "#34d399" : darkTheme.accentGold}`, padding: "8px 16px", borderRadius: "8px" }}>
                  <span style={{ fontSize: "11px", display: "block", color: darkTheme.textMuted }}>Calculated Total</span>
                  <span style={{ fontSize: "18px", fontWeight: "900", color: computedHours === 44 ? "#34d399" : darkTheme.accentGold }}>{computedHours} / 44 Hrs</span>
                </div>
              )}
            </div>

            {!selectedAgentId ? (
              <div style={{ padding: "30px", textAlign: "center", border: `2px dashed ${darkTheme.border}`, borderRadius: "8px", color: darkTheme.textMuted }}>
                Please select an active agent profile from the filter to manage individual weekly shift rosters.
              </div>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
                <thead>
                  <tr style={{ borderBottom: `2px solid ${darkTheme.border}`, color: darkTheme.textMuted, textAlign: "left" }}>
                    <th style={{ padding: "12px" }}>Weekday</th>
                    <th style={{ padding: "12px" }}>Shift Start</th>
                    <th style={{ padding: "12px" }}>Shift End</th>
                    <th style={{ padding: "12px" }}>Lunch Start</th>
                    <th style={{ padding: "12px" }}>Lunch End</th>
                    <th style={{ padding: "12px", textAlign: "center" }}>Short Hour Day (-1h)</th>
                  </tr>
                </thead>
                <tbody>
                  {["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].map(day => {
                    const shift = activeSchedule?.days[day] || { startTime: "08:00", endTime: "18:00", lunchStart: "12:00", lunchEnd: "13:00", isShortHourDay: false };
                    return (
                      <tr key={day} style={{ borderBottom: `1px solid ${darkTheme.border}` }}>
                        <td style={{ padding: "12px", fontWeight: "bold", color: darkTheme.accent }}>{day}</td>
                        <td style={{ padding: "12px" }}>
                          <input type="text" value={shift.startTime} onChange={(e) => handleShiftChange(day, "startTime", e.target.value)} style={{ backgroundColor: darkTheme.bg, border: `1px solid ${darkTheme.border}`, color: "#fff", padding: "4px 8px", width: "65px", borderRadius: "4px", textAlign: "center" }} />
                        </td>
                        <td style={{ padding: "12px" }}>
                          <input type="text" value={shift.endTime} onChange={(e) => handleShiftChange(day, "endTime", e.target.value)} style={{ backgroundColor: darkTheme.bg, border: `1px solid ${darkTheme.border}`, color: "#fff", padding: "4px 8px", width: "65px", borderRadius: "4px", textAlign: "center" }} />
                        </td>
                        <td style={{ padding: "12px" }}>
                          <input type="text" value={shift.lunchStart} onChange={(e) => handleShiftChange(day, "lunchStart", e.target.value)} style={{ backgroundColor: darkTheme.bg, border: `1px solid ${darkTheme.border}`, color: "#fff", padding: "4px 8px", width: "65px", borderRadius: "4px", textAlign: "center" }} />
                        </td>
                        <td style={{ padding: "12px" }}>
                          <input type="text" value={shift.lunchEnd} onChange={(e) => handleShiftChange(day, "lunchEnd", e.target.value)} style={{ backgroundColor: darkTheme.bg, border: `1px solid ${darkTheme.border}`, color: "#fff", padding: "4px 8px", width: "65px", borderRadius: "4px", textAlign: "center" }} />
                        </td>
                        <td style={{ padding: "12px", textAlign: "center" }}>
                          <input type="checkbox" checked={shift.isShortHourDay} onChange={(e) => handleShiftChange(day, "isShortHourDay", e.target.checked)} style={{ transform: "scale(1.2)", cursor: "pointer" }} />
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>

          {/* 📅 1-TO-1 COACHING CALENDAR BOOKER & SLOT MATCHMAKER */}
          <div style={{ backgroundColor: darkTheme.surface, border: `1px solid ${darkTheme.border}`, borderRadius: "16px", padding: "24px" }}>
            <h3 style={{ fontSize: "18px", fontWeight: "900", margin: "0 0 4px 0" }}>💬 Smart 1-on-1 Coaching Session Slots</h3>
            <p style={{ fontSize: "12px", color: darkTheme.textMuted, margin: "0 0 20px 0" }}>Cross-references team availability to map available time windows</p>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: "12px" }}>
              {Object.keys(managerAvailabilities).map(day => (
                <div key={day} style={{ backgroundColor: darkTheme.bg, border: `1px solid ${darkTheme.border}`, borderRadius: "8px", padding: "12px" }}>
                  <div style={{ fontSize: "13px", fontWeight: "bold", borderBottom: `1px solid ${darkTheme.border}`, paddingBottom: "6px", marginBottom: "8px", color: darkTheme.accent }}>{day}</div>
                  
                  {managerAvailabilities[day].map(slot => {
                    const isBooked = bookedSlots[activeWeek]?.[`${day}-${slot}`];
                    return (
                      <button
                        key={slot}
                        disabled={!selectedAgentId}
                        onClick={() => {
                          if (!selectedAgentId) return;
                          const currentWeekBooking = bookedSlots[activeWeek] || {};
                          if (currentWeekBooking[`${day}-${slot}`] === currentAgent?.login) {
                            delete currentWeekBooking[`${day}-${slot}`];
                          } else {
                            currentWeekBooking[`${day}-${slot}`] = currentAgent?.login || "";
                          }
                          setBookedSlots({ ...bookedSlots, [activeWeek]: currentWeekBooking });
                        }}
                        style={{
                          width: "100%",
                          padding: "6px",
                          margin: "4px 0",
                          backgroundColor: isBooked ? (isBooked === currentAgent?.login ? "rgba(0, 240, 255, 0.2)" : "rgba(226, 192, 68, 0.1)") : darkTheme.surface,
                          border: `1px solid ${isBooked ? darkTheme.accent : darkTheme.border}`,
                          color: isBooked ? "#fff" : darkTheme.textMuted,
                          fontSize: "11px",
                          borderRadius: "4px",
                          cursor: selectedAgentId ? "pointer" : "not-allowed",
                          textAlign: "center"
                        }}
                      >
                        <div>{slot}</div>
                        {isBooked && <div style={{ fontSize: "9px", color: darkTheme.accentGold, fontWeight: "bold", marginTop: "2px" }}>@{isBooked}</div>}
                      </button>
                    );
                  })}
                </div>
              ))}
            </div>
          </div>

        </div>
      </div>
    </div>
  );
}
