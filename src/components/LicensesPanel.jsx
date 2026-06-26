import { useEffect, useState } from "react";
import { useMsal } from "@azure/msal-react";
import { getTeamsRoomsLicenses } from "../services/graphService";

const ROOMS = [
  { name: "Sala Tenacidad",   email: "tenacidad@salasneto.com" },
  { name: "Sala Practicidad", email: "practicidad@salasneto.com" },
  { name: "Sala Entusiasmo",  email: "entusiasmo@salasneto.com" },
];

// Simulación de fecha de expiración por sala (se reemplaza cuando lleguen licencias reales)
const MOCK_EXPIRY = [
  { date: null, label: "Sin licencia asignada" },
  { date: null, label: "Sin licencia asignada" },
  { date: null, label: "Sin licencia asignada" },
];

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const diff = new Date(dateStr) - new Date();
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}

function ExpiryBadge({ dateStr }) {
  if (!dateStr) {
    return <span style={s.badgePending}>Pendiente</span>;
  }
  const days = daysUntil(dateStr);
  if (days < 0)  return <span style={s.badgeExp}>Expirada</span>;
  if (days <= 30) return <span style={s.badgeWarn}>Vence en {days}d</span>;
  return <span style={s.badgeOk}>Activa</span>;
}

export default function LicensesPanel() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const [licenses, setLicenses] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    getTeamsRoomsLicenses(instance, account)
      .then(setLicenses)
      .finally(() => setLoading(false));
  }, []);

  const lic = licenses[0]; // Teams Rooms Pro

  return (
    <div style={s.card}>
      <div style={s.header}>
        <span style={s.title}>🎫 Licencias Teams Rooms Pro</span>
        {lic?.isMock && (
          <span style={s.mockBadge}>Sin licencias activas</span>
        )}
      </div>

      {loading ? (
        <div style={s.loading}>Consultando licencias…</div>
      ) : (
        <>
          {/* Resumen general */}
          <div style={s.summary}>
            <div style={s.stat}>
              <div style={{ ...s.statNum, color: "#185FA5" }}>{lic?.total ?? 0}</div>
              <div style={s.statLabel}>Total</div>
            </div>
            <div style={s.statDiv} />
            <div style={s.stat}>
              <div style={{ ...s.statNum, color: "#1D9E75" }}>{lic?.consumed ?? 0}</div>
              <div style={s.statLabel}>Asignadas</div>
            </div>
            <div style={s.statDiv} />
            <div style={s.stat}>
              <div style={{ ...s.statNum, color: lic?.available > 0 ? "#534AB7" : "#aaa" }}>
                {lic?.available ?? 0}
              </div>
              <div style={s.statLabel}>Disponibles</div>
            </div>
          </div>

          {/* Detalle por sala */}
          <div style={s.roomList}>
            {ROOMS.map((room, i) => {
              const expiry = lic?.isMock ? MOCK_EXPIRY[i] : { date: lic?.expiryDate, label: null };
              return (
                <div key={room.email} style={s.roomRow}>
                  <div style={s.roomIcon}>🏢</div>
                  <div style={s.roomInfo}>
                    <div style={s.roomName}>{room.name}</div>
                    <div style={s.roomSku}>Teams Rooms Pro</div>
                  </div>
                  <ExpiryBadge dateStr={expiry.date} />
                </div>
              );
            })}
          </div>

          {lic?.isMock && (
            <div style={s.notice}>
              💡 Cuando adquieras las licencias aparecerán aquí automáticamente con fecha de vigencia.
            </div>
          )}
        </>
      )}
    </div>
  );
}

const s = {
  card: { background: "#fff", border: "0.5px solid #eee", borderRadius: 10, padding: "12px 14px" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 },
  title: { fontSize: 13, fontWeight: 500, color: "#555" },
  mockBadge: { fontSize: 10, padding: "2px 7px", borderRadius: 10, background: "#FEF3E2", color: "#92570A", fontWeight: 500 },
  loading: { fontSize: 12, color: "#aaa", padding: "10px 0", textAlign: "center" },
  summary: { display: "flex", alignItems: "center", background: "#f8f8f8", borderRadius: 8, padding: "8px 12px", marginBottom: 10, gap: 4 },
  stat: { flex: 1, textAlign: "center" },
  statNum: { fontSize: 20, fontWeight: 600, lineHeight: 1 },
  statLabel: { fontSize: 10, color: "#aaa", marginTop: 3 },
  statDiv: { width: "0.5px", height: 28, background: "#eee" },
  roomList: { display: "flex", flexDirection: "column", gap: 6 },
  roomRow: { display: "flex", alignItems: "center", gap: 8, padding: "6px 8px", borderRadius: 7, border: "0.5px solid #eee", background: "#fafafa" },
  roomIcon: { fontSize: 14, flexShrink: 0 },
  roomInfo: { flex: 1 },
  roomName: { fontSize: 12, fontWeight: 500, color: "#222" },
  roomSku: { fontSize: 10, color: "#888", marginTop: 1 },
  badgeOk: { fontSize: 10, padding: "2px 7px", borderRadius: 10, background: "#E1F5EE", color: "#085041", fontWeight: 500, whiteSpace: "nowrap" },
  badgeWarn: { fontSize: 10, padding: "2px 7px", borderRadius: 10, background: "#FEF3E2", color: "#92570A", fontWeight: 500, whiteSpace: "nowrap" },
  badgeExp: { fontSize: 10, padding: "2px 7px", borderRadius: 10, background: "#FCEBEB", color: "#A32D2D", fontWeight: 500, whiteSpace: "nowrap" },
  badgePending: { fontSize: 10, padding: "2px 7px", borderRadius: 10, background: "#f0f0f0", color: "#888", fontWeight: 500, whiteSpace: "nowrap" },
  notice: { marginTop: 8, fontSize: 10, color: "#92570A", background: "#FEF3E2", borderRadius: 7, padding: "6px 9px", lineHeight: 1.5 },
};
