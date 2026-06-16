// ============================================================
//  COMPONENTE: BookingModal
//  Formulario para crear una nueva reserva de sala
// ============================================================

import { useState } from "react";
import { useMsal } from "@azure/msal-react";
import { createBooking, checkAvailability } from "../services/graphService";
import { format, addHours, setHours, setMinutes } from "date-fns";
import { es } from "date-fns/locale";

const HOURS = Array.from({ length: 12 }, (_, i) => i + 7); // 7:00 - 18:00
const DURATIONS = [
  { label: "30 minutos", value: 0.5 },
  { label: "1 hora", value: 1 },
  { label: "1.5 horas", value: 1.5 },
  { label: "2 horas", value: 2 },
  { label: "3 horas", value: 3 },
];

export default function BookingModal({ rooms, selectedDate, onClose, onSuccess }) {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [form, setForm] = useState({
    roomEmail: rooms[0]?.emailAddress || "",
    subject: "",
    date: format(selectedDate, "yyyy-MM-dd"),
    hour: 9,
    minute: 0,
    duration: 1,
    attendees: "",
  });

  const [status, setStatus] = useState("idle"); // idle | checking | creating | error | success
  const [errorMsg, setErrorMsg] = useState("");

  const set = (field) => (e) =>
    setForm((f) => ({ ...f, [field]: e.target.value }));

  const getStartEnd = () => {
    const base = new Date(form.date + "T00:00:00");
    const start = setMinutes(setHours(base, Number(form.hour)), Number(form.minute));
    const end = addHours(start, Number(form.duration));
    return { start, end };
  };

  const handleSubmit = async () => {
    if (!form.subject.trim()) {
      setErrorMsg("El asunto es obligatorio.");
      setStatus("error");
      return;
    }

    const room = rooms.find((r) => r.emailAddress === form.roomEmail);
    if (!room) return;

    const { start, end } = getStartEnd();

    try {
      setStatus("checking");
      setErrorMsg("");

      const avail = await checkAvailability(instance, account, form.roomEmail, start, end);
      if (!avail.available) {
        setErrorMsg("La sala no está disponible en ese horario. Elige otro.");
        setStatus("error");
        return;
      }

      setStatus("creating");
      const attendeeList = form.attendees
        .split(",")
        .map((a) => a.trim())
        .filter(Boolean);

      await createBooking(instance, account, {
        subject: form.subject,
        roomEmail: form.roomEmail,
        roomName: room.displayName,
        start,
        end,
        attendees: attendeeList,
      });

      setStatus("success");
      setTimeout(() => {
        onSuccess();
        onClose();
      }, 1200);
    } catch (err) {
      setErrorMsg(err.message || "Error al crear la reserva.");
      setStatus("error");
    }
  };

  const isLoading = status === "checking" || status === "creating";

  return (
    <div style={styles.overlay}>
      <div style={styles.modal}>
        <div style={styles.header}>
          <h2 style={styles.title}>Nueva reserva</h2>
          <button onClick={onClose} style={styles.closeBtn} aria-label="Cerrar">✕</button>
        </div>

        <div style={styles.body}>
          {/* Sala */}
          <label style={styles.label}>Sala</label>
          <select style={styles.input} value={form.roomEmail} onChange={set("roomEmail")}>
            {rooms.map((r) => (
              <option key={r.id} value={r.emailAddress}>
                {r.displayName} {r.capacity ? `(cap. ${r.capacity})` : ""}
              </option>
            ))}
          </select>

          {/* Asunto */}
          <label style={styles.label}>Asunto *</label>
          <input
            style={styles.input}
            type="text"
            placeholder="Ej: Reunión de equipo"
            value={form.subject}
            onChange={set("subject")}
          />

          {/* Fecha y hora */}
          <div style={styles.row}>
            <div style={{ flex: 1 }}>
              <label style={styles.label}>Fecha</label>
              <input style={styles.input} type="date" value={form.date} onChange={set("date")} />
            </div>
            <div style={{ flex: 1 }}>
              <label style={styles.label}>Hora inicio</label>
              <select style={styles.input} value={form.hour} onChange={set("hour")}>
                {HOURS.map((h) => (
                  <option key={h} value={h}>
                    {String(h).padStart(2, "0")}:00
                  </option>
                ))}
              </select>
            </div>
            <div style={{ flex: 1 }}>
              <label style={styles.label}>Duración</label>
              <select style={styles.input} value={form.duration} onChange={set("duration")}>
                {DURATIONS.map((d) => (
                  <option key={d.value} value={d.value}>
                    {d.label}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Asistentes */}
          <label style={styles.label}>Asistentes (opcional, correos separados por coma)</label>
          <input
            style={styles.input}
            type="text"
            placeholder="ana@empresa.com, carlos@empresa.com"
            value={form.attendees}
            onChange={set("attendees")}
          />

          {/* Resumen */}
          {form.subject && (
            <div style={styles.summary}>
              <span>📅 </span>
              <strong>{form.subject}</strong> —{" "}
              {format(new Date(form.date + "T00:00:00"), "EEEE d MMM", { locale: es })},{" "}
              {String(form.hour).padStart(2, "0")}:00 por {form.duration}h
            </div>
          )}

          {/* Error */}
          {status === "error" && (
            <div style={styles.errorBox}>{errorMsg}</div>
          )}

          {/* Éxito */}
          {status === "success" && (
            <div style={styles.successBox}>✓ Reserva creada correctamente</div>
          )}
        </div>

        <div style={styles.footer}>
          <button onClick={onClose} style={styles.btnSec} disabled={isLoading}>
            Cancelar
          </button>
          <button onClick={handleSubmit} style={styles.btnPrimary} disabled={isLoading}>
            {status === "checking"
              ? "Verificando disponibilidad…"
              : status === "creating"
              ? "Creando reserva…"
              : "Confirmar reserva"}
          </button>
        </div>
      </div>
    </div>
  );
}

const styles = {
  overlay: {
    position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)",
    display: "flex", alignItems: "center", justifyContent: "center",
    zIndex: 1000, padding: "1rem",
  },
  modal: {
    background: "#fff", borderRadius: 16, width: "100%", maxWidth: 560,
    boxShadow: "0 8px 40px rgba(0,0,0,0.18)", overflow: "hidden",
  },
  header: {
    display: "flex", alignItems: "center", justifyContent: "space-between",
    padding: "18px 24px", borderBottom: "1px solid #eee",
  },
  title: { margin: 0, fontSize: 17, fontWeight: 600 },
  closeBtn: {
    background: "none", border: "none", fontSize: 18, cursor: "pointer",
    color: "#888", lineHeight: 1, padding: "2px 6px",
  },
  body: { padding: "20px 24px", display: "flex", flexDirection: "column", gap: 4 },
  label: { fontSize: 12, fontWeight: 500, color: "#666", marginTop: 10, marginBottom: 4 },
  input: {
    padding: "9px 12px", borderRadius: 8, border: "1px solid #ddd",
    fontSize: 14, width: "100%", boxSizing: "border-box", outline: "none",
    fontFamily: "inherit",
  },
  row: { display: "flex", gap: 12, marginTop: 4 },
  summary: {
    marginTop: 12, padding: "10px 14px", background: "#f0f7ff",
    borderRadius: 8, fontSize: 13, color: "#1a5fa8",
  },
  errorBox: {
    marginTop: 8, padding: "10px 14px", background: "#fff0f0",
    borderRadius: 8, fontSize: 13, color: "#c0392b",
  },
  successBox: {
    marginTop: 8, padding: "10px 14px", background: "#f0fff4",
    borderRadius: 8, fontSize: 13, color: "#1a7a4a",
  },
  footer: {
    display: "flex", gap: 10, justifyContent: "flex-end",
    padding: "16px 24px", borderTop: "1px solid #eee",
  },
  btnPrimary: {
    padding: "9px 20px", background: "#1a5fa8", color: "#fff",
    border: "none", borderRadius: 8, fontSize: 14, fontWeight: 500,
    cursor: "pointer",
  },
  btnSec: {
    padding: "9px 16px", background: "#f5f5f5", color: "#444",
    border: "1px solid #ddd", borderRadius: 8, fontSize: 14, cursor: "pointer",
  },
};
