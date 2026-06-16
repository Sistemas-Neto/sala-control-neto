// ============================================================
//  COMPONENTE: StatsPanel
//  Métricas de ocupación de las salas (últimos 30 días)
// ============================================================

import { useState, useEffect } from "react";
import { useMsal } from "@azure/msal-react";
import { getRoomStats } from "../services/graphService";

export default function StatsPanel({ rooms }) {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [stats, setStats] = useState({});
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (rooms.length === 0) return;
    setLoading(true);

    Promise.allSettled(
      rooms.map((r) =>
        getRoomStats(instance, account, r.emailAddress).then((s) => ({
          email: r.emailAddress,
          ...s,
        }))
      )
    ).then((results) => {
      const map = {};
      results.forEach((r) => {
        if (r.status === "fulfilled") {
          map[r.value.email] = r.value;
        }
      });
      setStats(map);
      setLoading(false);
    });
  }, [rooms, instance, account]);

  if (loading) {
    return <div style={styles.loading}>Cargando estadísticas…</div>;
  }

  const totalBookings = Object.values(stats).reduce((s, r) => s + r.totalBookings, 0);
  const avgOccupancy =
    Object.values(stats).length > 0
      ? Math.round(
          Object.values(stats).reduce((s, r) => s + r.occupancyPct, 0) /
            Object.values(stats).length
        )
      : 0;

  return (
    <div style={styles.wrapper}>
      <h3 style={styles.sectionTitle}>Últimos 30 días</h3>

      {/* Resumen global */}
      <div style={styles.summaryGrid}>
        <StatCard label="Reservas totales" value={totalBookings} color="#1a5fa8" />
        <StatCard label="Ocupación promedio" value={`${avgOccupancy}%`} color="#27ae60" />
        <StatCard label="Salas activas" value={rooms.length} color="#8e44ad" />
      </div>

      {/* Por sala */}
      <h3 style={{ ...styles.sectionTitle, marginTop: 20 }}>Por sala</h3>
      <div style={styles.roomGrid}>
        {rooms.map((room) => {
          const s = stats[room.emailAddress];
          if (!s) return null;
          return (
            <div key={room.id} style={styles.roomCard}>
              <div style={styles.roomCardName}>{room.displayName}</div>
              <div style={styles.roomCardRow}>
                <span style={styles.metricLabel}>Reservas</span>
                <span style={styles.metricValue}>{s.totalBookings}</span>
              </div>
              <div style={styles.roomCardRow}>
                <span style={styles.metricLabel}>Horas usadas</span>
                <span style={styles.metricValue}>{s.totalHours}h</span>
              </div>
              <div style={styles.roomCardRow}>
                <span style={styles.metricLabel}>Ocupación</span>
                <span style={styles.metricValue}>{s.occupancyPct}%</span>
              </div>
              {/* Barra de ocupación */}
              <div style={styles.barTrack}>
                <div
                  style={{
                    ...styles.barFill,
                    width: `${s.occupancyPct}%`,
                    background:
                      s.occupancyPct > 75
                        ? "#e74c3c"
                        : s.occupancyPct > 40
                        ? "#e67e22"
                        : "#27ae60",
                  }}
                />
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function StatCard({ label, value, color }) {
  return (
    <div style={styles.statCard}>
      <div style={{ ...styles.statValue, color }}>{value}</div>
      <div style={styles.statLabel}>{label}</div>
    </div>
  );
}

const styles = {
  wrapper: { padding: "4px 0" },
  loading: { padding: 24, textAlign: "center", color: "#888", fontSize: 14 },
  sectionTitle: { fontSize: 13, fontWeight: 600, color: "#555", margin: "0 0 12px" },
  summaryGrid: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 },
  statCard: {
    background: "#f8f9fa", borderRadius: 10, padding: "14px 16px",
    border: "1px solid #eee",
  },
  statValue: { fontSize: 26, fontWeight: 700, lineHeight: 1 },
  statLabel: { fontSize: 12, color: "#888", marginTop: 4 },
  roomGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: 10 },
  roomCard: {
    background: "#fff", border: "1px solid #eee", borderRadius: 10,
    padding: "12px 14px",
  },
  roomCardName: { fontSize: 13, fontWeight: 600, color: "#222", marginBottom: 10 },
  roomCardRow: { display: "flex", justifyContent: "space-between", marginBottom: 5 },
  metricLabel: { fontSize: 12, color: "#888" },
  metricValue: { fontSize: 12, fontWeight: 600, color: "#222" },
  barTrack: {
    marginTop: 8, height: 4, background: "#eee", borderRadius: 2, overflow: "hidden",
  },
  barFill: { height: "100%", borderRadius: 2, transition: "width 0.5s" },
};
