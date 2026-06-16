// ============================================================
//  COMPONENTE: RoomCalendar
//  Vista de calendario diario por sala con bloques de eventos
// ============================================================

import { useMsal } from "@azure/msal-react";
import { cancelBooking } from "../services/graphService";
import { format } from "date-fns";
import { es } from "date-fns/locale";

const HOURS = Array.from({ length: 12 }, (_, i) => i + 7); // 7:00 - 18:00
const SLOT_HEIGHT = 60; // px por hora

function eventToSlot(event) {
  const start = new Date(event.start.dateTime);
  const end = new Date(event.end.dateTime);
  const startH = start.getHours() + start.getMinutes() / 60;
  const endH = end.getHours() + end.getMinutes() / 60;
  const top = (startH - 7) * SLOT_HEIGHT;
  const height = Math.max((endH - startH) * SLOT_HEIGHT, 24);
  return { top, height, startH, endH };
}

function isNow(event) {
  const now = new Date();
  return new Date(event.start.dateTime) <= now && now <= new Date(event.end.dateTime);
}

function isUpcoming(event) {
  const now = new Date();
  const start = new Date(event.start.dateTime);
  return start > now && start - now < 2 * 60 * 60 * 1000; // próximas 2h
}

export default function RoomCalendar({ rooms, events, onRefresh }) {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const handleCancel = async (eventId, subject) => {
    if (!window.confirm(`¿Cancelar la reserva "${subject}"?`)) return;
    try {
      await cancelBooking(instance, account, eventId);
      onRefresh();
    } catch (err) {
      alert("Error al cancelar: " + err.message);
    }
  };

  if (rooms.length === 0) {
    return (
      <div style={styles.empty}>
        No se encontraron salas en el directorio. Verifica que tengas salas configuradas
        como recursos en Exchange Online.
      </div>
    );
  }

  const nowHour = new Date().getHours() + new Date().getMinutes() / 60;
  const nowTop = (nowHour - 7) * SLOT_HEIGHT;

  return (
    <div style={styles.wrapper}>
      {/* Header de salas */}
      <div style={styles.grid(rooms.length)}>
        <div style={styles.timeHeader} />
        {rooms.map((room) => {
          const roomEvents = events[room.emailAddress] || [];
          const busy = roomEvents.some(isNow);
          return (
            <div key={room.id} style={styles.roomHeader}>
              <span style={styles.roomName}>{room.displayName}</span>
              <span style={{ ...styles.statusDot, background: busy ? "#e74c3c" : "#27ae60" }} />
              <span style={styles.statusLabel}>{busy ? "En uso" : "Libre"}</span>
              {room.capacity && (
                <span style={styles.capacity}>cap. {room.capacity}</span>
              )}
            </div>
          );
        })}
      </div>

      {/* Grid de tiempo */}
      <div style={styles.scrollArea}>
        <div style={styles.grid(rooms.length)}>
          {/* Columna de horas */}
          <div style={styles.timeCol}>
            {HOURS.map((h) => (
              <div key={h} style={styles.timeCell}>
                {String(h).padStart(2, "0")}:00
              </div>
            ))}
          </div>

          {/* Columnas de salas */}
          {rooms.map((room) => {
            const roomEvents = events[room.emailAddress] || [];
            return (
              <div key={room.id} style={styles.roomCol}>
                {/* Líneas de hora */}
                {HOURS.map((h) => (
                  <div key={h} style={styles.hourLine} />
                ))}

                {/* Línea "ahora" */}
                {nowTop > 0 && nowTop < HOURS.length * SLOT_HEIGHT && (
                  <div style={{ ...styles.nowLine, top: nowTop }} />
                )}

                {/* Eventos */}
                {roomEvents.map((ev) => {
                  const { top, height } = eventToSlot(ev);
                  const now = isNow(ev);
                  const upcoming = isUpcoming(ev);
                  return (
                    <div
                      key={ev.id}
                      style={{
                        ...styles.event,
                        top,
                        height: height - 2,
                        ...(now ? styles.eventNow : upcoming ? styles.eventUpcoming : styles.eventFuture),
                      }}
                    >
                      <div style={styles.eventSubject}>{ev.subject || "(Sin asunto)"}</div>
                      <div style={styles.eventOrganizer}>
                        {ev.organizer?.emailAddress?.name || ev.organizer?.emailAddress?.address}
                      </div>
                      <div style={styles.eventTime}>
                        {format(new Date(ev.start.dateTime), "HH:mm")} –{" "}
                        {format(new Date(ev.end.dateTime), "HH:mm")}
                      </div>
                      {!now && (
                        <button
                          style={styles.cancelBtn}
                          onClick={() => handleCancel(ev.id, ev.subject)}
                          title="Cancelar reserva"
                        >
                          ✕
                        </button>
                      )}
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      </div>

      {/* Leyenda */}
      <div style={styles.legend}>
        <LegendItem color="#e74c3c" label="En uso ahora" />
        <LegendItem color="#e67e22" label="Próxima (≤2h)" />
        <LegendItem color="#2980b9" label="Reservada" />
        <LegendItem color="#27ae60" label="Libre" />
      </div>
    </div>
  );
}

function LegendItem({ color, label }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "#666" }}>
      <div style={{ width: 10, height: 10, borderRadius: 2, background: color }} />
      {label}
    </div>
  );
}

const SLOT_H = SLOT_HEIGHT;
const styles = {
  wrapper: { border: "1px solid #e8e8e8", borderRadius: 12, overflow: "hidden", background: "#fff" },
  grid: (cols) => ({
    display: "grid",
    gridTemplateColumns: `52px repeat(${cols}, 1fr)`,
  }),
  timeHeader: { background: "#fafafa", borderBottom: "1px solid #eee", borderRight: "1px solid #eee" },
  roomHeader: {
    padding: "10px 12px", background: "#fafafa",
    borderBottom: "1px solid #eee", borderRight: "1px solid #eee",
    display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap",
  },
  roomName: { fontSize: 13, fontWeight: 600, color: "#222", flex: 1 },
  statusDot: { width: 8, height: 8, borderRadius: "50%", flexShrink: 0 },
  statusLabel: { fontSize: 11, color: "#666" },
  capacity: {
    fontSize: 10, padding: "1px 6px", background: "#f0f0f0",
    borderRadius: 10, color: "#888",
  },
  scrollArea: { overflowY: "auto", maxHeight: 520 },
  timeCol: { borderRight: "1px solid #eee" },
  timeCell: {
    height: SLOT_H, display: "flex", alignItems: "flex-start",
    justifyContent: "flex-end", paddingRight: 8, paddingTop: 4,
    fontSize: 10, color: "#aaa", boxSizing: "border-box",
    borderBottom: "1px solid #f5f5f5",
  },
  roomCol: {
    position: "relative", borderRight: "1px solid #eee",
    height: SLOT_H * HOURS.length,
  },
  hourLine: {
    position: "absolute", left: 0, right: 0,
    height: SLOT_H, borderBottom: "1px solid #f5f5f5",
    boxSizing: "border-box",
  },
  nowLine: {
    position: "absolute", left: 0, right: 0,
    height: 2, background: "#e74c3c", zIndex: 5,
    boxSizing: "border-box",
  },
  event: {
    position: "absolute", left: 4, right: 4,
    borderRadius: 6, padding: "4px 8px",
    overflow: "hidden", cursor: "default", boxSizing: "border-box",
    zIndex: 2,
  },
  eventNow: { background: "#fde8e8", borderLeft: "3px solid #e74c3c" },
  eventUpcoming: { background: "#fef3e2", borderLeft: "3px solid #e67e22" },
  eventFuture: { background: "#e8f4fd", borderLeft: "3px solid #2980b9" },
  eventSubject: { fontSize: 11, fontWeight: 600, color: "#222", lineHeight: 1.3, paddingRight: 16 },
  eventOrganizer: { fontSize: 10, color: "#666", marginTop: 2 },
  eventTime: { fontSize: 10, color: "#888", marginTop: 1 },
  cancelBtn: {
    position: "absolute", top: 4, right: 4,
    background: "none", border: "none", cursor: "pointer",
    fontSize: 10, color: "#aaa", padding: "1px 3px",
    borderRadius: 3,
  },
  legend: {
    display: "flex", gap: 16, padding: "10px 14px",
    borderTop: "1px solid #eee", flexWrap: "wrap",
  },
  empty: {
    padding: 32, textAlign: "center", color: "#888", fontSize: 14,
    background: "#fafafa", borderRadius: 12, border: "1px dashed #ddd",
  },
};
