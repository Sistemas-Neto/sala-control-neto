import { useState } from "react";
import { useMsal } from "@azure/msal-react";
import { cancelBooking } from "../services/graphService";
import { format, addDays, startOfWeek, startOfMonth, getDaysInMonth } from "date-fns";
import { es } from "date-fns/locale";
import BookingModal from "./BookingModal";

const HOURS = Array.from({ length: 16 }, (_, i) => i + 7);
const SLOT_HEIGHT = 60;
const TZ = "America/Mexico_City";

function toLocal(dateStr) {
  return new Date(new Date(dateStr).toLocaleString("en-US", { timeZone: TZ }));
}

function eventToSlot(event) {
  const localStart = toLocal(event.start.dateTime);
  const localEnd = toLocal(event.end.dateTime);
  const startH = localStart.getHours() + localStart.getMinutes() / 60;
  const endH = localEnd.getHours() + localEnd.getMinutes() / 60;
  const top = (startH - 7) * SLOT_HEIGHT;
  const height = Math.max((endH - startH) * SLOT_HEIGHT, 24);
  return { top, height, startH, endH, localStart, localEnd };
}

function isNow(event) {
  const now = new Date();
  return new Date(event.start.dateTime) <= now && now <= new Date(event.end.dateTime);
}

function isUpcoming(event) {
  const now = new Date();
  const start = new Date(event.start.dateTime);
  return start > now && start - now < 2 * 60 * 60 * 1000;
}

function isSameDay(date1, date2) {
  return (
    date1.getFullYear() === date2.getFullYear() &&
    date1.getMonth() === date2.getMonth() &&
    date1.getDate() === date2.getDate()
  );
}

// ── VISTA DÍA ──────────────────────────────────────────────
function DayView({ rooms, events, selectedDate, onRefresh, onEditEvent }) {
  const nowLocal = toLocal(new Date().toISOString());
  const nowTop = (nowLocal.getHours() + nowLocal.getMinutes() / 60 - 7) * SLOT_HEIGHT;

  const handleCancel = async (instance, account, eventId, subject) => {
    if (!window.confirm(`¿Cancelar la reserva "${subject}"?`)) return;
    try {
      await cancelBooking(instance, account, eventId);
      onRefresh();
    } catch (err) {
      alert("Error al cancelar: " + err.message);
    }
  };

  const { instance, accounts } = useMsal ? { instance: null, accounts: [] } : { instance: null, accounts: [] };

  return (
    <div style={styles.scrollArea}>
      <div style={styles.grid(rooms.length)}>
        <div style={styles.timeCol}>
          {HOURS.map((h) => (
            <div key={h} style={styles.timeCell}>
              {String(h).padStart(2, "0")}:00
            </div>
          ))}
        </div>
        {rooms.map((room) => {
          const roomEvents = (events[room.emailAddress] || []).filter(ev =>
            isSameDay(toLocal(ev.start.dateTime), selectedDate)
          );
          return (
            <div key={room.id} style={styles.roomCol}>
              {HOURS.map((h) => <div key={h} style={styles.hourLine} />)}
              {nowTop > 0 && nowTop < HOURS.length * SLOT_HEIGHT && isSameDay(nowLocal, selectedDate) && (
                <div style={{ ...styles.nowLine, top: nowTop }} />
              )}
              {roomEvents.map((ev) => {
                const { top, height, localStart, localEnd } = eventToSlot(ev);
                const now = isNow(ev);
                const upcoming = isUpcoming(ev);
                return (
                  <div
                    key={ev.id}
                    style={{
                      ...styles.event, top, height: height - 2, cursor: "pointer",
                      ...(now ? styles.eventNow : upcoming ? styles.eventUpcoming : styles.eventFuture),
                    }}
                    onClick={() => onEditEvent(ev, room)}
                    title="Clic para editar"
                  >
                    <div style={styles.eventSubject}>{ev.subject || "(Sin asunto)"}</div>
                    <div style={styles.eventOrganizer}>{ev.organizer?.emailAddress?.name}</div>
                    <div style={styles.eventTime}>{format(localStart, "HH:mm")} – {format(localEnd, "HH:mm")}</div>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── VISTA SEMANA ───────────────────────────────────────────
function WeekView({ rooms, events, selectedDate, onRefresh, onEditEvent }) {
  const weekStart = startOfWeek(selectedDate, { weekStartsOn: 1 });
  const weekDays = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));
  const DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"];

  const allEvents = rooms.flatMap(room =>
    (events[room.emailAddress] || []).map(ev => ({ ...ev, roomName: room.displayName }))
  );

  return (
    <div style={{ overflowX: "auto" }}>
      <div style={{ display: "grid", gridTemplateColumns: `60px repeat(7, 1fr)`, minWidth: 700 }}>
        <div style={{ background: "#fafafa", borderBottom: "1px solid #eee", borderRight: "1px solid #eee", padding: 8 }} />
        {weekDays.map((day, i) => {
          const isToday = isSameDay(day, new Date());
          return (
            <div key={i} style={{
              padding: "8px 6px", background: "#fafafa",
              borderBottom: "1px solid #eee", borderRight: "1px solid #eee",
              textAlign: "center",
            }}>
              <div style={{ fontSize: 12, color: "#888" }}>{DIAS[i]}</div>
              <div style={{
                fontSize: 16, fontWeight: 600,
                color: isToday ? "#fff" : "#222",
                background: isToday ? "#042C53" : "transparent",
                borderRadius: "50%", width: 28, height: 28,
                display: "flex", alignItems: "center", justifyContent: "center",
                margin: "2px auto 0",
              }}>{day.getDate()}</div>
            </div>
          );
        })}

        <div style={{ borderRight: "1px solid #eee" }}>
          {HOURS.map(h => (
            <div key={h} style={{ height: SLOT_HEIGHT, borderBottom: "1px solid #f5f5f5", display: "flex", alignItems: "flex-start", justifyContent: "flex-end", paddingRight: 6, paddingTop: 4, fontSize: 11, color: "#aaa", boxSizing: "border-box" }}>
              {String(h).padStart(2, "0")}:00
            </div>
          ))}
        </div>

        {weekDays.map((day, i) => {
          const dayEvents = allEvents.filter(ev => isSameDay(toLocal(ev.start.dateTime), day));
          const nowLocal = toLocal(new Date().toISOString());
          const nowTop = (nowLocal.getHours() + nowLocal.getMinutes() / 60 - 7) * SLOT_HEIGHT;
          return (
            <div key={i} style={{ position: "relative", borderRight: "1px solid #eee", height: SLOT_HEIGHT * HOURS.length }}>
              {HOURS.map(h => <div key={h} style={{ position: "absolute", left: 0, right: 0, top: (h - 7) * SLOT_HEIGHT, borderBottom: "1px solid #f5f5f5", height: SLOT_HEIGHT, boxSizing: "border-box" }} />)}
              {isSameDay(day, new Date()) && nowTop > 0 && (
                <div style={{ position: "absolute", left: 0, right: 0, top: nowTop, height: 2, background: "#e74c3c", zIndex: 5 }} />
              )}
              {dayEvents.map((ev) => {
                const { top, height, localStart, localEnd } = eventToSlot(ev);
                const now = isNow(ev);
                const upcoming = isUpcoming(ev);
                return (
                  <div key={ev.id} style={{
                    position: "absolute", left: 2, right: 2, top, height: height - 2,
                    borderRadius: 5, padding: "3px 6px", overflow: "hidden",
                    boxSizing: "border-box", zIndex: 2, cursor: "pointer",
                    fontSize: 11,
                    ...(now ? styles.eventNow : upcoming ? styles.eventUpcoming : styles.eventFuture),
                  }}
                    onClick={() => onEditEvent(ev, rooms.find(r => r.displayName === ev.roomName))}
                  >
                    <div style={{ fontWeight: 600, lineHeight: 1.2, overflow: "hidden", whiteSpace: "nowrap", textOverflow: "ellipsis" }}>
                      {format(localStart, "HH:mm")} {ev.subject || "(Sin asunto)"}
                    </div>
                    <div style={{ color: "#555", fontSize: 10, overflow: "hidden", whiteSpace: "nowrap", textOverflow: "ellipsis" }}>📍 {ev.roomName}</div>
                    <div style={{ color: "#888", fontSize: 10 }}>{format(localStart, "HH:mm")}–{format(localEnd, "HH:mm")}</div>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── VISTA MES ──────────────────────────────────────────────
function MonthView({ rooms, events, selectedDate, onEditEvent }) {
  const year = selectedDate.getFullYear();
  const month = selectedDate.getMonth();
  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = getDaysInMonth(selectedDate);
  const DIAS = ["Dom", "Lun", "Mar", "Mié", "Jue", "Vie", "Sáb"];

  const allEvents = rooms.flatMap(room =>
    (events[room.emailAddress] || []).map(ev => ({ ...ev, roomName: room.displayName }))
  );

  const getEventsForDay = (day) => {
    const date = new Date(year, month, day);
    return allEvents.filter(ev => isSameDay(toLocal(ev.start.dateTime), date));
  };

  const cells = [];
  for (let i = 0; i < firstDay; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", borderBottom: "1px solid #eee" }}>
        {DIAS.map(d => (
          <div key={d} style={{ padding: "8px 0", textAlign: "center", fontSize: 13, fontWeight: 500, color: "#888" }}>{d}</div>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)" }}>
        {cells.map((day, i) => {
          const isToday = day && isSameDay(new Date(year, month, day), new Date());
          const dayEvents = day ? getEventsForDay(day) : [];
          return (
            <div key={i} style={{
              minHeight: 80, border: "0.5px solid #f0f0f0",
              padding: "5px 6px", background: isToday ? "#EBF4FD" : "#fff",
            }}>
              {day && (
                <>
                  <div style={{
                    fontSize: 13, fontWeight: isToday ? 700 : 400,
                    color: isToday ? "#042C53" : "#222",
                    marginBottom: 4,
                  }}>{day}</div>
                  {dayEvents.slice(0, 3).map((ev, j) => {
                    const localStart = toLocal(ev.start.dateTime);
                    return (
                      <div key={j}
                        onClick={() => onEditEvent(ev, rooms.find(r => r.displayName === ev.roomName))}
                        style={{
                          fontSize: 11, padding: "3px 6px", borderRadius: 4,
                          marginBottom: 3, cursor: "pointer",
                          ...(isNow(ev) ? styles.eventNow : styles.eventFuture),
                        }}>
                        <div style={{ fontWeight: 600, overflow: "hidden", whiteSpace: "nowrap", textOverflow: "ellipsis" }}>
                          {format(localStart, "HH:mm")} {ev.subject || "(Sin asunto)"}
                        </div>
                        <div style={{ fontSize: 10, color: "#555", overflow: "hidden", whiteSpace: "nowrap", textOverflow: "ellipsis" }}>
                          📍 {ev.roomName}
                        </div>
                      </div>
                    );
                  })}
                  {dayEvents.length > 3 && (
                    <div style={{ fontSize: 11, color: "#888" }}>+{dayEvents.length - 3} más</div>
                  )}
                </>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── COMPONENTE PRINCIPAL ───────────────────────────────────
export default function RoomCalendar({ rooms, events, onRefresh, selectedDate: propDate }) {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const [editEvent, setEditEvent] = useState(null);
  const [editRoom, setEditRoom] = useState(null);
  const [vista, setVista] = useState("dia");
  const [refreshing, setRefreshing] = useState(false);

  const selectedDate = propDate || new Date();

  const handleRefresh = async () => {
    setRefreshing(true);
    await onRefresh();
    setTimeout(() => setRefreshing(false), 800);
  };

  const handleCancel = async (eventId, subject) => {
    if (!window.confirm(`¿Cancelar la reserva "${subject}"?`)) return;
    try {
      await cancelBooking(instance, account, eventId);
      onRefresh();
    } catch (err) {
      alert("Error al cancelar: " + err.message);
    }
  };

  const handleEventClick = (ev, room) => {
    setEditEvent(ev);
    setEditRoom(room);
  };

  if (rooms.length === 0) {
    return <div style={styles.empty}>No se encontraron salas. Verifica la configuración en Azure AD.</div>;
  }

  const nowLocal = toLocal(new Date().toISOString());
  const nowTop = (nowLocal.getHours() + nowLocal.getMinutes() / 60 - 7) * SLOT_HEIGHT;

  return (
    <>
      <div style={styles.wrapper}>
        {/* Barra superior con filtros y botón actualizar */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 14px", borderBottom: "1px solid #eee", background: "#fafafa" }}>
          <div style={{ display: "flex", gap: 4 }}>
            {["dia", "semana", "mes"].map(v => (
              <button key={v} onClick={() => setVista(v)} style={{
                padding: "5px 14px", borderRadius: 6, fontSize: 13, cursor: "pointer",
                border: "0.5px solid " + (vista === v ? "#042C53" : "#ddd"),
                background: vista === v ? "#042C53" : "#fff",
                color: vista === v ? "#fff" : "#666",
                fontWeight: vista === v ? 500 : 400,
              }}>
                {v.charAt(0).toUpperCase() + v.slice(1)}
              </button>
            ))}
          </div>
          <button onClick={handleRefresh} style={{
            display: "flex", alignItems: "center", gap: 6,
            padding: "5px 14px", borderRadius: 6, fontSize: 13,
            border: "0.5px solid #ddd", background: "#fff", color: "#555", cursor: "pointer",
          }}>
            <span style={{ display: "inline-block", animation: refreshing ? "spin 0.8s linear infinite" : "none" }}>↻</span>
            Actualizar
          </button>
        </div>

        {/* Header de salas (solo en vista día) */}
        {vista === "dia" && (
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
                  {room.capacity && <span style={styles.capacity}>cap. {room.capacity}</span>}
                </div>
              );
            })}
          </div>
        )}

        {/* Vistas */}
        {vista === "dia" && (
          <div style={styles.scrollArea}>
            <div style={styles.grid(rooms.length)}>
              <div style={styles.timeCol}>
                {HOURS.map((h) => (
                  <div key={h} style={styles.timeCell}>{String(h).padStart(2, "0")}:00</div>
                ))}
              </div>
              {rooms.map((room) => {
                const roomEvents = (events[room.emailAddress] || []).filter(ev =>
                  isSameDay(toLocal(ev.start.dateTime), selectedDate)
                );
                return (
                  <div key={room.id} style={styles.roomCol}>
                    {HOURS.map((h) => <div key={h} style={styles.hourLine} />)}
                    {nowTop > 0 && nowTop < HOURS.length * SLOT_HEIGHT && isSameDay(nowLocal, selectedDate) && (
                      <div style={{ ...styles.nowLine, top: nowTop }} />
                    )}
                    {roomEvents.map((ev) => {
                      const { top, height, localStart, localEnd } = eventToSlot(ev);
                      const now = isNow(ev);
                      const upcoming = isUpcoming(ev);
                      return (
                        <div key={ev.id} style={{
                          ...styles.event, top, height: height - 2, cursor: "pointer",
                          ...(now ? styles.eventNow : upcoming ? styles.eventUpcoming : styles.eventFuture),
                        }}
                          onClick={() => handleEventClick(ev, room)}
                        >
                          <div style={styles.eventSubject}>{ev.subject || "(Sin asunto)"}</div>
                          <div style={styles.eventOrganizer}>{ev.organizer?.emailAddress?.name}</div>
                          <div style={styles.eventTime}>{format(localStart, "HH:mm")} – {format(localEnd, "HH:mm")}</div>
                          {!now && (
                            <button style={styles.cancelBtn}
                              onClick={(e) => { e.stopPropagation(); handleCancel(ev.id, ev.subject); }}
                            >✕</button>
                          )}
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {vista === "semana" && (
          <WeekView rooms={rooms} events={events} selectedDate={selectedDate} onRefresh={onRefresh} onEditEvent={handleEventClick} />
        )}

        {vista === "mes" && (
          <MonthView rooms={rooms} events={events} selectedDate={selectedDate} onEditEvent={handleEventClick} />
        )}

        {/* Leyenda */}
        <div style={styles.legend}>
          <LegendItem color="#e74c3c" label="En uso ahora" />
          <LegendItem color="#e67e22" label="Próxima (≤2h)" />
          <LegendItem color="#2980b9" label="Reservada" />
          <LegendItem color="#27ae60" label="Libre" />
        </div>
      </div>

      {/* Modal de edición */}
      {editEvent && editRoom && (
        <BookingModal
          rooms={rooms}
          selectedDate={new Date(editEvent.start.dateTime)}
          editEvent={editEvent}
          editRoom={editRoom}
          onClose={() => { setEditEvent(null); setEditRoom(null); }}
          onSuccess={() => { setEditEvent(null); setEditRoom(null); onRefresh(); }}
        />
      )}

      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </>
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

const styles = {
  wrapper: { border: "1px solid #e8e8e8", borderRadius: 10, overflow: "hidden", background: "#fff" },
  grid: (cols) => ({ display: "grid", gridTemplateColumns: `52px repeat(${cols}, 1fr)` }),
  timeHeader: { background: "#fafafa", borderBottom: "1px solid #eee", borderRight: "1px solid #eee" },
  roomHeader: { padding: "10px 12px", background: "#fafafa", borderBottom: "1px solid #eee", borderRight: "1px solid #eee", display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" },
  roomName: { fontSize: 13, fontWeight: 600, color: "#222", flex: 1 },
  statusDot: { width: 8, height: 8, borderRadius: "50%", flexShrink: 0 },
  statusLabel: { fontSize: 12, color: "#666" },
  capacity: { fontSize: 11, padding: "1px 6px", background: "#f0f0f0", borderRadius: 10, color: "#888" },
  scrollArea: { overflowY: "auto", maxHeight: 480 },
  timeCol: { borderRight: "1px solid #eee" },
  timeCell: { height: SLOT_HEIGHT, display: "flex", alignItems: "flex-start", justifyContent: "flex-end", paddingRight: 8, paddingTop: 4, fontSize: 11, color: "#aaa", boxSizing: "border-box", borderBottom: "1px solid #f5f5f5" },
  roomCol: { position: "relative", borderRight: "1px solid #eee", height: SLOT_HEIGHT * HOURS.length },
  hourLine: { position: "absolute", left: 0, right: 0, height: SLOT_HEIGHT, borderBottom: "1px solid #f5f5f5", boxSizing: "border-box" },
  nowLine: { position: "absolute", left: 0, right: 0, height: 2, background: "#e74c3c", zIndex: 5, boxSizing: "border-box" },
  event: { position: "absolute", left: 4, right: 4, borderRadius: 6, padding: "4px 8px", overflow: "hidden", boxSizing: "border-box", zIndex: 2 },
  eventNow: { background: "#fde8e8", borderLeft: "3px solid #e74c3c" },
  eventUpcoming: { background: "#fef3e2", borderLeft: "3px solid #e67e22" },
  eventFuture: { background: "#e8f4fd", borderLeft: "3px solid #2980b9" },
  eventSubject: { fontSize: 12, fontWeight: 600, color: "#222", lineHeight: 1.3, paddingRight: 16 },
  eventOrganizer: { fontSize: 11, color: "#666", marginTop: 2 },
  eventTime: { fontSize: 11, color: "#888", marginTop: 1 },
  cancelBtn: { position: "absolute", top: 4, right: 4, background: "none", border: "none", cursor: "pointer", fontSize: 11, color: "#aaa", padding: "1px 3px", borderRadius: 3 },
  legend: { display: "flex", gap: 16, padding: "10px 14px", borderTop: "1px solid #eee", flexWrap: "wrap" },
  empty: { padding: 32, textAlign: "center", color: "#888", fontSize: 14, background: "#fafafa", borderRadius: 10, border: "1px dashed #ddd" },
};
