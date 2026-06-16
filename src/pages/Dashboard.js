// ============================================================
//  PÁGINA PRINCIPAL: Dashboard
//  Vista completa de control de salas
// ============================================================

import { useState } from "react";
import { useMsal } from "@azure/msal-react";
import { format, addDays, subDays } from "date-fns";
import { es } from "date-fns/locale";

import { useRooms } from "../hooks/useRooms";
import RoomCalendar from "../components/RoomCalendar";
import BookingModal from "../components/BookingModal";
import StatsPanel from "../components/StatsPanel";

export default function Dashboard() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [selectedDate, setSelectedDate] = useState(new Date());
  const [activeTab, setActiveTab] = useState("calendar"); // calendar | stats
  const [showModal, setShowModal] = useState(false);

  const { rooms, events, loading, error, refresh } = useRooms(selectedDate);

  const handleLogout = () => instance.logoutRedirect();

  return (
    <div style={styles.page}>
      {/* ── Top bar ── */}
      <header style={styles.topBar}>
        <div style={styles.brand}>
          <span style={styles.brandIcon}>🏢</span>
          <span style={styles.brandName}>Control de Salas</span>
        </div>

        <div style={styles.topCenter}>
          <button
            style={styles.navBtn}
            onClick={() => setSelectedDate((d) => subDays(d, 1))}
          >
            ‹
          </button>
          <span style={styles.dateLabel}>
            {format(selectedDate, "EEEE d 'de' MMMM, yyyy", { locale: es })}
          </span>
          <button
            style={styles.navBtn}
            onClick={() => setSelectedDate((d) => addDays(d, 1))}
          >
            ›
          </button>
          <button
            style={styles.todayBtn}
            onClick={() => setSelectedDate(new Date())}
          >
            Hoy
          </button>
        </div>

        <div style={styles.topRight}>
          <span style={styles.userLabel}>
            {account?.name || account?.username}
          </span>
          <button onClick={handleLogout} style={styles.logoutBtn}>
            Salir
          </button>
        </div>
      </header>

      {/* ── Contenido ── */}
      <main style={styles.main}>
        {/* Tabs + botón nueva reserva */}
        <div style={styles.toolbar}>
          <div style={styles.tabs}>
            <TabBtn
              active={activeTab === "calendar"}
              onClick={() => setActiveTab("calendar")}
            >
              📅 Calendario
            </TabBtn>
            <TabBtn
              active={activeTab === "stats"}
              onClick={() => setActiveTab("stats")}
            >
              📊 Estadísticas
            </TabBtn>
          </div>
          <button style={styles.newBtn} onClick={() => setShowModal(true)}>
            + Nueva reserva
          </button>
        </div>

        {/* Estados */}
        {loading && (
          <div style={styles.stateBox}>
            <div style={styles.spinner} />
            Cargando salas y eventos…
          </div>
        )}

        {error && !loading && (
          <div style={styles.errorBox}>
            <strong>Error al cargar datos:</strong> {error}
            <br />
            <small>Verifica que la app tenga los permisos correctos en Azure AD.</small>
          </div>
        )}

        {!loading && !error && activeTab === "calendar" && (
          <RoomCalendar rooms={rooms} events={events} onRefresh={refresh} />
        )}

        {!loading && !error && activeTab === "stats" && (
          <StatsPanel rooms={rooms} />
        )}
      </main>

      {/* Modal de nueva reserva */}
      {showModal && rooms.length > 0 && (
        <BookingModal
          rooms={rooms}
          selectedDate={selectedDate}
          onClose={() => setShowModal(false)}
          onSuccess={refresh}
        />
      )}
    </div>
  );
}

function TabBtn({ active, onClick, children }) {
  return (
    <button
      onClick={onClick}
      style={{
        ...styles.tab,
        ...(active ? styles.tabActive : {}),
      }}
    >
      {children}
    </button>
  );
}

const styles = {
  page: {
    minHeight: "100vh", background: "#f4f6f9",
    fontFamily: "'Segoe UI', system-ui, sans-serif",
  },
  topBar: {
    display: "flex", alignItems: "center", justifyContent: "space-between",
    padding: "12px 24px", background: "#fff",
    borderBottom: "1px solid #e8e8e8",
    position: "sticky", top: 0, zIndex: 100,
  },
  brand: { display: "flex", alignItems: "center", gap: 8 },
  brandIcon: { fontSize: 20 },
  brandName: { fontSize: 16, fontWeight: 700, color: "#1a5fa8" },
  topCenter: { display: "flex", alignItems: "center", gap: 8 },
  navBtn: {
    background: "none", border: "1px solid #ddd", borderRadius: 6,
    width: 28, height: 28, cursor: "pointer", fontSize: 16,
    display: "flex", alignItems: "center", justifyContent: "center",
    color: "#444",
  },
  dateLabel: { fontSize: 14, fontWeight: 500, color: "#222", minWidth: 240, textAlign: "center" },
  todayBtn: {
    padding: "4px 10px", background: "#f0f4fa", border: "1px solid #c5d5ea",
    borderRadius: 6, fontSize: 12, cursor: "pointer", color: "#1a5fa8",
  },
  topRight: { display: "flex", alignItems: "center", gap: 10 },
  userLabel: { fontSize: 13, color: "#666" },
  logoutBtn: {
    padding: "5px 12px", background: "none", border: "1px solid #ddd",
    borderRadius: 6, fontSize: 12, cursor: "pointer", color: "#888",
  },
  main: { maxWidth: 1100, margin: "0 auto", padding: "20px 24px" },
  toolbar: {
    display: "flex", justifyContent: "space-between", alignItems: "center",
    marginBottom: 16,
  },
  tabs: { display: "flex", gap: 4 },
  tab: {
    padding: "8px 16px", background: "none",
    border: "1px solid transparent", borderRadius: 8,
    fontSize: 13, cursor: "pointer", color: "#666",
  },
  tabActive: {
    background: "#fff", border: "1px solid #e0e0e0",
    color: "#1a5fa8", fontWeight: 600,
  },
  newBtn: {
    padding: "9px 18px", background: "#1a5fa8", color: "#fff",
    border: "none", borderRadius: 8, fontSize: 14,
    fontWeight: 500, cursor: "pointer",
  },
  stateBox: {
    display: "flex", alignItems: "center", gap: 12,
    padding: 32, color: "#888", fontSize: 14, justifyContent: "center",
  },
  spinner: {
    width: 18, height: 18, border: "2px solid #ddd",
    borderTopColor: "#1a5fa8", borderRadius: "50%",
    animation: "spin 0.7s linear infinite",
  },
  errorBox: {
    padding: 20, background: "#fff5f5", border: "1px solid #fcc",
    borderRadius: 10, color: "#c0392b", fontSize: 14, lineHeight: 1.6,
  },
};
