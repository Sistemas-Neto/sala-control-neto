import { useState, useEffect } from "react";
import { useMsal } from "@azure/msal-react";
import { getRoomStats } from "../services/graphService";
import { format, subDays, subMonths, startOfMonth, endOfMonth, startOfWeek, endOfWeek } from "date-fns";
import { es } from "date-fns/locale";

const PERIODS = [
  { label: "Esta semana",    id: "week" },
  { label: "Últimos 7 días", id: "7d" },
  { label: "Últimos 30 días",id: "30d" },
  { label: "Este mes",       id: "month" },
  { label: "Mes anterior",   id: "prevmonth" },
  { label: "Últimos 3 meses",id: "90d" },
  { label: "Personalizado",  id: "custom" },
];

function getRange(periodId) {
  const now = new Date();
  switch (periodId) {
    case "week":     return { start: startOfWeek(now, { weekStartsOn: 1 }), end: now };
    case "7d":       return { start: subDays(now, 7), end: now };
    case "30d":      return { start: subDays(now, 30), end: now };
    case "month":    return { start: startOfMonth(now), end: now };
    case "prevmonth":return { start: startOfMonth(subMonths(now, 1)), end: endOfMonth(subMonths(now, 1)) };
    case "90d":      return { start: subDays(now, 90), end: now };
    default:         return { start: subDays(now, 30), end: now };
  }
}

export default function StatsPanel({ rooms }) {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [period, setPeriod] = useState("30d");
  const [customStart, setCustomStart] = useState(format(subDays(new Date(), 30), "yyyy-MM-dd"));
  const [customEnd, setCustomEnd] = useState(format(new Date(), "yyyy-MM-dd"));
  const [stats, setStats] = useState({});
  const [loading, setLoading] = useState(true);

  const loadStats = (periodId, cStart, cEnd) => {
    if (rooms.length === 0) return;
    setLoading(true);

    const { start, end } = periodId === "custom"
      ? { start: new Date(cStart), end: new Date(cEnd + "T23:59:59") }
      : getRange(periodId);

    Promise.allSettled(
      rooms.map((r) =>
        getRoomStats(instance, account, r.emailAddress, start, end).then((s) => ({
          email: r.emailAddress,
          ...s,
        }))
      )
    ).then((results) => {
      const map = {};
      results.forEach((r) => {
        if (r.status === "fulfilled") map[r.value.email] = r.value;
      });
      setStats(map);
      setLoading(false);
    });
  };

  useEffect(() => {
    loadStats(period, customStart, customEnd);
  }, [rooms, instance, account]);

  const handlePeriod = (id) => {
    setPeriod(id);
    if (id !== "custom") loadStats(id, customStart, customEnd);
  };

  const handleCustomApply = () => loadStats("custom", customStart, customEnd);

  const totalBookings = Object.values(stats).reduce((s, r) => s + r.totalBookings, 0);
  const avgOccupancy = Object.values(stats).length > 0
    ? Math.round(Object.values(stats).reduce((s, r) => s + r.occupancyPct, 0) / Object.values(stats).length)
    : 0;

  const periodLabel = PERIODS.find(p => p.id === period)?.label || "";

  return (
    <div style={st.wrapper}>

      {/* Filtros de período */}
      <div style={st.filterBar}>
        {PERIODS.map(p => (
          <button
            key={p.id}
            style={{ ...st.filterBtn, ...(period === p.id ? st.filterBtnActive : {}) }}
            onClick={() => handlePeriod(p.id)}
          >
            {p.label}
          </button>
        ))}
      </div>

      {/* Rango personalizado */}
      {period === "custom" && (
        <div style={st.customRange}>
          <div style={st.customLabel}>Desde</div>
          <input style={st.dateInput} type="date" value={customStart} onChange={e => setCustomStart(e.target.value)} />
          <div style={st.customLabel}>Hasta</div>
          <input style={st.dateInput} type="date" value={customEnd} onChange={e => setCustomEnd(e.target.value)} />
          <button style={st.applyBtn} onClick={handleCustomApply}>Aplicar</button>
        </div>
      )}

      {loading ? (
        <div style={st.loading}>Cargando estadísticas…</div>
      ) : (
        <>
          <div style={st.periodTitle}>{periodLabel}</div>

          {/* Resumen global */}
          <div style={st.summaryGrid}>
            <StatCard label="Reservas totales" value={totalBookings} color="#1a5fa8" />
            <StatCard label="Ocupación promedio" value={`${avgOccupancy}%`} color="#27ae60" />
            <StatCard label="Salas activas" value={rooms.length} color="#8e44ad" />
          </div>

          {/* Por sala */}
          <div style={st.sectionTitle}>Por sala</div>
          <div style={st.roomGrid}>
            {rooms.map((room) => {
              const s = stats[room.emailAddress];
              if (!s) return null;
              return (
                <div key={room.id} style={st.roomCard}>
                  <div style={st.roomCardName}>{room.displayName}</div>
                  <div style={st.roomCardRow}>
                    <span style={st.metricLabel}>Reservas</span>
                    <span style={st.metricValue}>{s.totalBookings}</span>
                  </div>
                  <div style={st.roomCardRow}>
                    <span style={st.metricLabel}>Horas usadas</span>
                    <span style={st.metricValue}>{s.totalHours}h</span>
                  </div>
                  <div style={st.roomCardRow}>
                    <span style={st.metricLabel}>Ocupación</span>
                    <span style={st.metricValue}>{s.occupancyPct}%</span>
                  </div>
                  <div style={st.barTrack}>
                    <div style={{
                      ...st.barFill,
                      width: `${s.occupancyPct}%`,
                      background: s.occupancyPct > 75 ? "#e74c3c" : s.occupancyPct > 40 ? "#e67e22" : "#27ae60",
                    }} />
                  </div>
                </div>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
}

function StatCard({ label, value, color }) {
  return (
    <div style={st.statCard}>
      <div style={{ ...st.statValue, color }}>{value}</div>
      <div style={st.statLabel}>{label}</div>
    </div>
  );
}

const st = {
  wrapper: { padding: "4px 0" },
  filterBar: { display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 14 },
  filterBtn: { padding: "5px 11px", borderRadius: 20, border: "0.5px solid #ddd", fontSize: 12, background: "#f8f8f8", color: "#555", cursor: "pointer" },
  filterBtnActive: { background: "#042C53", color: "#fff", border: "0.5px solid #042C53" },
  customRange: { display: "flex", alignItems: "center", gap: 8, marginBottom: 14, padding: "10px 12px", background: "#f8f8f8", borderRadius: 8, flexWrap: "wrap" },
  customLabel: { fontSize: 12, color: "#666" },
  dateInput: { padding: "5px 8px", border: "0.5px solid #ddd", borderRadius: 6, fontSize: 12 },
  applyBtn: { padding: "5px 12px", background: "#042C53", color: "#fff", border: "none", borderRadius: 6, fontSize: 12, cursor: "pointer" },
  loading: { padding: 24, textAlign: "center", color: "#888", fontSize: 14 },
  periodTitle: { fontSize: 13, fontWeight: 600, color: "#555", marginBottom: 12 },
  summaryGrid: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginBottom: 20 },
  statCard: { background: "#f8f9fa", borderRadius: 10, padding: "14px 16px", border: "1px solid #eee" },
  statValue: { fontSize: 26, fontWeight: 700, lineHeight: 1 },
  statLabel: { fontSize: 12, color: "#888", marginTop: 4 },
  sectionTitle: { fontSize: 13, fontWeight: 600, color: "#555", marginBottom: 12 },
  roomGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: 10 },
  roomCard: { background: "#fff", border: "1px solid #eee", borderRadius: 10, padding: "12px 14px" },
  roomCardName: { fontSize: 13, fontWeight: 600, color: "#222", marginBottom: 10 },
  roomCardRow: { display: "flex", justifyContent: "space-between", marginBottom: 5 },
  metricLabel: { fontSize: 12, color: "#888" },
  metricValue: { fontSize: 12, fontWeight: 600, color: "#222" },
  barTrack: { marginTop: 8, height: 4, background: "#eee", borderRadius: 2, overflow: "hidden" },
  barFill: { height: "100%", borderRadius: 2, transition: "width 0.5s" },
};
