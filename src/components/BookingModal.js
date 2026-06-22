import { useState } from "react";
import { useMsal } from "@azure/msal-react";
import { createBooking, checkAvailability } from "../services/graphService";
import { addHours, setHours, setMinutes } from "date-fns";

const HOURS = Array.from({ length: 12 }, (_, i) => i + 7);
const DURATIONS = [
  { label: "30 minutos", value: 0.5 },
  { label: "1 hora", value: 1 },
  { label: "1.5 horas", value: 1.5 },
  { label: "2 horas", value: 2 },
  { label: "3 horas", value: 3 },
];

const COMBO = {
  nombre: "Sala Magna (Ejecutiva + Practicidad)",
  salas: ["salapracticidad@soyneto.onmicrosoft.com", "salaentusiasmo@soyneto.onmicrosoft.com"],
  capacidad: 70,
};

export default function BookingModal({ rooms, selectedDate, onClose, onSuccess }) {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [selSala, setSelSala] = useState(null);
  const [isCombo, setIsCombo] = useState(false);
  const [form, setForm] = useState({
    subject: "",
    date: selectedDate.toISOString().split("T")[0],
    hour: 9,
    duration: 1,
    attendees: "",
    comments: "",
  });
  const [showTeams, setShowTeams] = useState(false);
  const [showComments, setShowComments] = useState(false);
  const [teamsLink, setTeamsLink] = useState("https://teams.microsoft.com/l/meetup-join/…");
  const [status, setStatus] = useState("idle");
  const [errorMsg, setErrorMsg] = useState("");

  const set = (field) => (e) => setForm((f) => ({ ...f, [field]: e.target.value }));

  const selectSala = (sala, combo = false) => {
    setSelSala(sala);
    setIsCombo(combo);
  };

  const getStartEnd = () => {
    const base = new Date(form.date + "T00:00:00");
    const start = setMinutes(setHours(base, Number(form.hour)), 0);
    const end = addHours(start, Number(form.duration));
    return { start, end };
  };

  const genTeamsLink = () => {
    const fake = "https://teams.microsoft.com/l/meetup-join/19%3Ameeting_" + Math.random().toString(36).substr(2, 8).toUpperCase() + "%40thread.v2/0";
    setTeamsLink(fake);
  };

  const handleSubmit = async () => {
    if (!selSala) { setErrorMsg("Selecciona una sala."); setStatus("error"); return; }
    if (!form.subject.trim()) { setErrorMsg("El asunto es obligatorio."); setStatus("error"); return; }

    const { start, end } = getStartEnd();

    try {
      setStatus("checking");
      setErrorMsg("");

      if (isCombo) {
        // Reservar las dos salas
        for (const email of COMBO.salas) {
          const room = rooms.find(r => r.emailAddress === email) || { emailAddress: email, displayName: email };
          await createBooking(instance, account, {
            subject: form.subject,
            roomEmail: email,
            roomName: room.displayName || email,
            start, end,
            attendees: form.attendees.split(",").map(a => a.trim()).filter(Boolean),
          });
        }
      } else {
        const room = rooms.find(r => r.emailAddress === selSala);
        await createBooking(instance, account, {
          subject: form.subject,
          roomEmail: selSala,
          roomName: room?.displayName || selSala,
          start, end,
          attendees: form.attendees.split(",").map(a => a.trim()).filter(Boolean),
        });
      }

      setStatus("success");
      setTimeout(() => { onSuccess(); onClose(); }, 1500);
    } catch (err) {
      setErrorMsg(err.message || "Error al crear la reserva.");
      setStatus("error");
    }
  };

  const isLoading = status === "checking" || status === "creating";

  return (
    <div style={ov} onClick={onClose}>
      <div style={modal} onClick={e => e.stopPropagation()}>
        <div style={mh}>
          <span style={{ fontSize: 13, fontWeight: 500 }}>📅 Nueva reserva</span>
          <button onClick={onClose} style={closeBtn}>✕</button>
        </div>
        <div style={mb}>
          {/* SELECTOR DE SALA */}
          <div>
            <div style={lbl}>Selecciona la sala</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {rooms.map(r => (
                <div key={r.id} onClick={() => selectSala(r.emailAddress, false)}
                  style={{ ...salaOpt, ...(selSala === r.emailAddress && !isCombo ? salaSelected : {}) }}>
                  <div style={salaIcon("🏢", "#E6F1FB")}></div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 12, fontWeight: 500, color: "#222" }}>{r.displayName}</div>
                    <div style={{ fontSize: 10, color: "#888" }}>Cap. {r.capacity || "—"}</div>
                  </div>
                  <span style={{ ...badge, background: "#E1F5EE", color: "#085041" }}>Libre</span>
                </div>
              ))}
              {/* SALA COMBINADA */}
              <div onClick={() => selectSala("magna", true)}
                style={{ ...salaOpt, border: isCombo ? "1.5px solid #7F77DD" : "0.5px dashed #AFA9EC", background: isCombo ? "#F0EAF7" : "#fff" }}>
                <div style={{ width: 30, height: 30, borderRadius: 6, background: "#F0EAF7", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, flexShrink: 0 }}>🔗</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 12, fontWeight: 500, color: "#3C3489" }}>Sala Magna (combinada)</div>
                  <div style={{ fontSize: 10, color: "#534AB7" }}>Tenacidad + Entusiasmo · Cap. {COMBO.capacidad}</div>
                </div>
                <span style={{ ...badge, background: "#EEEDFE", color: "#3C3489" }}>Disponible</span>
              </div>
            </div>
          </div>

          {/* AVISO SALA COMBINADA */}
          {isCombo && (
            <div style={{ padding: "8px 10px", background: "#FAEEDA", border: "0.5px solid #FAC775", borderRadius: 8, fontSize: 11, color: "#633806" }}>
              ⚠️ <strong>Reserva combinada</strong> — Ambas salas quedarán bloqueadas. Se abrirá la pared divisoria para cap. {COMBO.capacidad} personas.
            </div>
          )}

          <div><div style={lbl}>Asunto *</div><input style={inp} type="text" placeholder="Ej: Congreso anual" value={form.subject} onChange={set("subject")} /></div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            <div><div style={lbl}>Fecha</div><input style={inp} type="date" value={form.date} onChange={set("date")} /></div>
            <div><div style={lbl}>Hora inicio</div>
              <select style={inp} value={form.hour} onChange={set("hour")}>
                {HOURS.map(h => <option key={h} value={h}>{String(h).padStart(2,"0")}:00</option>)}
              </select>
            </div>
          </div>

          <div><div style={lbl}>Duración</div>
            <select style={inp} value={form.duration} onChange={set("duration")}>
              {DURATIONS.map(d => <option key={d.value} value={d.value}>{d.label}</option>)}
            </select>
          </div>

          <div><div style={lbl}>Asistentes (opcional)</div><input style={inp} type="text" placeholder="correos separados por coma" value={form.attendees} onChange={set("attendees")} /></div>

          <div style={{ height: "0.5px", background: "#eee" }} />

          {/* OPCIONALES */}
          <div style={{ display: "flex", gap: 7 }}>
            <button onClick={() => setShowTeams(!showTeams)} style={{ ...optBtn, ...(showTeams ? optBtnOn : {}) }}>
              🔵 Liga de Teams
            </button>
            <button onClick={() => setShowComments(!showComments)} style={{ ...optBtn, ...(showComments ? optBtnOn : {}) }}>
              💬 Comentarios
            </button>
          </div>

          {showTeams && (
            <div style={optSec}>
              <div style={{ fontSize: 11, fontWeight: 500, color: "#666", marginBottom: 5 }}>Liga de Microsoft Teams</div>
              <div style={{ display: "flex", gap: 5 }}>
                <div style={{ flex: 1, padding: "5px 8px", background: "#fff", border: "0.5px solid #ddd", borderRadius: 6, fontSize: 10, color: "#185FA5", fontFamily: "monospace", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{teamsLink}</div>
                <button onClick={genTeamsLink} style={{ padding: "4px 8px", background: "#5B5FC7", color: "#fff", border: "none", borderRadius: 6, fontSize: 10, cursor: "pointer" }}>Generar</button>
              </div>
            </div>
          )}

          {showComments && (
            <div style={optSec}>
              <div style={{ fontSize: 11, fontWeight: 500, color: "#666", marginBottom: 5 }}>Comentarios para el correo</div>
              <textarea style={{ ...inp, resize: "vertical", lineHeight: 1.5 }} rows={2} placeholder="Ej: Favor de llegar 5 min antes." value={form.comments} onChange={set("comments")} />
            </div>
          )}

          {status === "error" && <div style={{ padding: "7px 9px", background: "#FCEBEB", borderRadius: 8, fontSize: 11, color: "#791F1F" }}>{errorMsg}</div>}
          {status === "success" && <div style={{ padding: "7px 9px", background: "#E1F5EE", borderRadius: 8, fontSize: 11, color: "#085041" }}>✓ {isCombo ? "¡Sala Magna reservada! Ambas salas bloqueadas." : "¡Reserva creada!"}</div>}
        </div>

        <div style={mf}>
          <span style={{ fontSize: 10, color: "#aaa" }}>{isCombo ? "Ambas salas quedarán bloqueadas" : "Se enviará invitación por correo"}</span>
          <div style={{ display: "flex", gap: 6 }}>
            <button onClick={onClose} style={btnSec} disabled={isLoading}>Cancelar</button>
            <button onClick={handleSubmit} style={btnPri} disabled={isLoading}>
              {isLoading ? "Procesando…" : "✉ Reservar y enviar"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

const ov = { position: "fixed", inset: 0, background: "rgba(0,0,0,0.42)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: "1rem" };
const modal = { background: "#fff", borderRadius: 14, width: "100%", maxWidth: 420, boxShadow: "0 8px 40px rgba(0,0,0,0.18)", overflow: "hidden", maxHeight: "90vh", display: "flex", flexDirection: "column" };
const mh = { padding: "13px 16px", borderBottom: "0.5px solid #eee", display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 };
const closeBtn = { background: "none", border: "none", fontSize: 16, cursor: "pointer", color: "#888" };
const mb = { padding: "13px 16px", display: "flex", flexDirection: "column", gap: 9, overflowY: "auto" };
const mf = { padding: "11px 16px", borderTop: "0.5px solid #eee", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 6, flexShrink: 0 };
const lbl = { fontSize: 11, color: "#666", marginBottom: 3 };
const inp = { width: "100%", padding: "7px 10px", border: "0.5px solid #ddd", borderRadius: 8, fontSize: 12, boxSizing: "border-box", fontFamily: "inherit" };
const badge = { fontSize: 10, padding: "2px 7px", borderRadius: 10, fontWeight: 500 };
const salaOpt = { display: "flex", alignItems: "center", gap: 8, padding: "8px 10px", border: "0.5px solid #ddd", borderRadius: 8, cursor: "pointer", background: "#fff", transition: "all .12s" };
const salaSelected = { border: "1.5px solid #042C53", background: "#E6F1FB" };
const salaIcon = (_, bg) => ({ width: 30, height: 30, borderRadius: 6, background: bg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, flexShrink: 0 });
const optBtn = { display: "flex", alignItems: "center", gap: 5, padding: "5px 10px", borderRadius: 8, border: "0.5px dashed #ddd", background: "none", fontSize: 11, color: "#666", cursor: "pointer", flex: 1, justifyContent: "center" };
const optBtnOn = { borderColor: "#185FA5", borderStyle: "solid", background: "#E6F1FB", color: "#0C447C" };
const optSec = { padding: "9px 11px", background: "#f8f8f8", borderRadius: 8, border: "0.5px solid #eee" };
const btnSec = { padding: "6px 12px", border: "0.5px solid #ddd", borderRadius: 8, fontSize: 12, background: "none", color: "#666", cursor: "pointer" };
const btnPri = { padding: "6px 14px", border: "none", borderRadius: 8, fontSize: 12, background: "#042C53", color: "#E6F1FB", cursor: "pointer", fontWeight: 500 };
