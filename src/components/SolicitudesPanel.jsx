import { useState, useEffect } from "react";
import { useMsal } from "@azure/msal-react";
import { getSolicitudes, actualizarEstado } from "../services/solicitudesService";
import { createBooking, createComboBooking } from "../services/graphService";
import { GROUP_ADMINS } from "../authConfig";

const SALAS_EMAIL = {
  "Sala Tenacidad":   "tenacidad@salasneto.com",
  "Sala Practicidad": "practicidad@salasneto.com",
  "Sala Entusiasmo":  "entusiasmo@salasneto.com",
};

const ESTADO_STYLE = {
  Pendiente: { background: "#FFF8E1", color: "#B45309" },
  Aprobado:  { background: "#E1F5EE", color: "#085041" },
  Rechazado: { background: "#FCEBEB", color: "#A32D2D" },
};

function DetalleModal({ sol, onClose, onAprobar, onRechazar, procesando, isAdmin }) {
  const fecha = sol.HoraInicio
    ? new Date(sol.HoraInicio).toLocaleDateString("es-MX", { weekday: "long", day: "2-digit", month: "long", year: "numeric" })
    : "—";
  const hi = sol.HoraInicio
    ? new Date(sol.HoraInicio).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" })
    : "—";
  const hf = sol.HoraFin
    ? new Date(sol.HoraFin).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" })
    : "—";

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}
      onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 12, width: 520, maxWidth: "95vw", boxShadow: "0 8px 40px rgba(0,0,0,0.18)", overflow: "hidden" }}
        onClick={e => e.stopPropagation()}>
        <div style={{ background: "#042C53", padding: "18px 24px", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 11, color: "#85B7EB", textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 4 }}>Solicitud de reserva</div>
            <div style={{ fontSize: 16, fontWeight: 500, color: "#fff" }}>{sol.Asunto || "(Sin asunto)"}</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{
              ...(ESTADO_STYLE[sol.Estado] || { background: "#f0f0f0", color: "#888" }),
              padding: "3px 10px", borderRadius: 10, fontWeight: 500, fontSize: 11
            }}>{sol.Estado || "—"}</span>
            <button onClick={onClose} style={{ background: "none", border: "none", color: "#85B7EB", fontSize: 18, cursor: "pointer", lineHeight: 1 }}>✕</button>
          </div>
        </div>

        <div style={{ padding: "20px 24px", display: "flex", flexDirection: "column", gap: 16 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <Field label="Responsable" value={sol.ResponsableDeLaSesi_x00f3_n} />
            <Field label="Correo" value={sol.correo} />
            <Field label="Área" value={sol._x00c1_rea} />
            <Field label="Compañía" value={sol.Compa_x00f1_ia} />
          </div>
          <div style={{ height: "0.5px", background: "#eee" }} />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
            <Field label="Sala" value={sol.Sala} />
            <Field label="Asistentes" value={sol.Asistentes} />
            <Field label="Fecha" value={fecha} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <Field label="Hora inicio" value={hi} />
            <Field label="Hora fin" value={hf} />
          </div>
          <div style={{ height: "0.5px", background: "#eee" }} />
          <div>
            <div style={{ fontSize: 11, color: "#888", fontWeight: 500, textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 6 }}>
              Material o requerimientos adicionales
            </div>
            <div style={{
              padding: "10px 12px", background: "#f8f8f8", borderRadius: 8,
              fontSize: 13, color: sol.RequerimientosAdicionales ? "#222" : "#aaa",
              minHeight: 48, lineHeight: 1.6
            }}>
              {sol.RequerimientosAdicionales || "Sin requerimientos adicionales"}
            </div>
          </div>
        </div>

        {isAdmin && sol.Estado === "Pendiente" && (
          <div style={{ padding: "14px 24px", borderTop: "0.5px solid #eee", display: "flex", justifyContent: "flex-end", gap: 8, background: "#f8f8f8" }}>
            <button
              disabled={procesando}
              onClick={() => onRechazar(sol)}
              style={{ padding: "8px 16px", borderRadius: 7, fontSize: 13, cursor: "pointer", border: "0.5px solid #E24B4A", background: "#fff", color: "#A32D2D" }}
            >✕ Rechazar</button>
            <button
              disabled={procesando}
              onClick={() => onAprobar(sol)}
              style={{ padding: "8px 16px", borderRadius: 7, fontSize: 13, cursor: "pointer", border: "none", background: "#1D9E75", color: "#fff", fontWeight: 500 }}
            >{procesando ? "Procesando..." : "✓ Aprobar solicitud"}</button>
          </div>
        )}
      </div>
    </div>
  );
}

function Field({ label, value }) {
  return (
    <div>
      <div style={{ fontSize: 11, color: "#888", fontWeight: 500, textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 3 }}>{label}</div>
      <div style={{ fontSize: 13, color: value ? "#222" : "#aaa" }}>{value || "—"}</div>
    </div>
  );
}

export default function SolicitudesPanel() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const isAdmin = account?.idTokenClaims?.groups?.includes(GROUP_ADMINS);
  const [solicitudes, setSolicitudes] = useState([]);
  const [loading, setLoading] = useState(true);
  const [procesando, setProcesando] = useState(null);
  const [filtro, setFiltro] = useState("Pendiente");
  const [detalle, setDetalle] = useState(null);

  const cargar = async () => {
    setLoading(true);
    try {
      const data = await getSolicitudes(instance, account);
      setSolicitudes(data);
    } catch (e) {
      console.error(e);
    }
    setLoading(false);
  };

  useEffect(() => { cargar(); }, []);

  const filtradas = solicitudes.filter(s =>
    filtro === "Todas" ? true : s.Estado === filtro
  );

  const aprobar = async (sol) => {
    if (!isAdmin) return;
    setProcesando(sol.id);
    try {
      let eventoTeams;

      if (sol.Sala === "Sala Magna") {
        eventoTeams = await createComboBooking(instance, account, {
          subject:    sol.Asunto,
          roomNames:  ["Sala Tenacidad", "Sala Entusiasmo"],
          roomEmails: ["tenacidad@salasneto.com", "entusiasmo@salasneto.com"],
          start:      sol.HoraInicio,
          end:        sol.HoraFin,
          comments:   sol.RequerimientosAdicionales || "",
        });
      } else {
        eventoTeams = await createBooking(instance, account, {
          subject:   sol.Asunto,
          roomEmail: SALAS_EMAIL[sol.Sala] || "",
          roomName:  sol.Sala,
          start:     sol.HoraInicio,
          end:       sol.HoraFin,
          comments:  sol.RequerimientosAdicionales || "",
        });
      }

      console.log("Evento Teams creado:", eventoTeams);

      const teamsLink = eventoTeams?.onlineMeeting?.joinUrl || "";
      const teamsMeetingId = eventoTeams?.onlineMeeting?.conferenceId || "";
      const teamsPasscode = eventoTeams?.onlineMeeting?.tollFreeNumbers?.[0]?.tollFreePhoneNumbers?.[0] || "";

      await actualizarEstado(instance, account, sol.id, "Aprobado");

      // Enviar correo de confirmación con datos de Teams
      try {
        const payload = {
          correo:         sol.correo,
          responsable:    sol.ResponsableDeLaSesi_x00f3_n,
          asunto:         sol.Asunto,
          sala:           sol.Sala,
          fecha:          new Date(sol.HoraInicio).toLocaleDateString("es-MX", { weekday: "long", day: "2-digit", month: "long", year: "numeric" }),
          horaInicio:     new Date(sol.HoraInicio).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" }),
          horaFin:        new Date(sol.HoraFin).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" }),
          asistentes:     sol.Asistentes,
          requerimientos: sol.RequerimientosAdicionales || "",
          teamsLink,
          teamsMeetingId,
          teamsPasscode,
        };

        console.log("Enviando al webhook de correo:", payload);

        const r = await fetch("https://webhook.soyneto.com/webhook/sala-aprobar", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });

        const text = await r.text();
        console.log("Respuesta webhook correo:", r.status, text);
      } catch (errCorreo) {
        console.error("Error enviando correo de confirmación:", errCorreo);
      }

      setDetalle(null);
      await cargar();
    } catch (e) {
      alert("Error al aprobar: " + e.message);
    }
    setProcesando(null);
  };

  const rechazar = async (sol) => {
    if (!isAdmin) return;
    if (!window.confirm(`¿Rechazar la solicitud de ${sol.ResponsableDeLaSesi_x00f3_n}?`)) return;
    setProcesando(sol.id);
    try {
      await actualizarEstado(instance, account, sol.id, "Rechazado");
      setDetalle(null);
      await cargar();
    } catch (e) {
      alert("Error al rechazar: " + e.message);
    }
    setProcesando(null);
  };

  return (
    <>
      <div style={{ background: "#fff", border: "0.5px solid #eee", borderRadius: 10, padding: "15px 17px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <div style={{ fontSize: 14, fontWeight: 500, color: "#222" }}>📋 Solicitudes de reserva</div>
          <div style={{ display: "flex", gap: 6 }}>
            {["Pendiente", "Aprobado", "Rechazado", "Todas"].map(f => (
              <button key={f} onClick={() => setFiltro(f)} style={{
                padding: "4px 10px", borderRadius: 6, fontSize: 11, cursor: "pointer",
                border: "0.5px solid " + (filtro === f ? "#042C53" : "#ddd"),
                background: filtro === f ? "#042C53" : "#fff",
                color: filtro === f ? "#fff" : "#666",
              }}>{f}</button>
            ))}
            <button onClick={cargar} style={{
              padding: "4px 10px", borderRadius: 6, fontSize: 11, cursor: "pointer",
              border: "0.5px solid #ddd", background: "#f8f8f8", color: "#666"
            }}>↻ Actualizar</button>
          </div>
        </div>

        {loading ? (
          <div style={{ textAlign: "center", padding: 30, color: "#888", fontSize: 13 }}>Cargando solicitudes...</div>
        ) : filtradas.length === 0 ? (
          <div style={{ textAlign: "center", padding: 30, color: "#aaa", fontSize: 13 }}>
            No hay solicitudes {filtro !== "Todas" ? filtro.toLowerCase() + "s" : ""}
          </div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ borderBottom: "0.5px solid #eee" }}>
                  {["Responsable", "Área", "Compañía", "Asunto", "Sala", "Fecha", "Horario", "Asist.", "Estado", isAdmin ? "Acciones" : ""].map(h => (
                    <th key={h} style={{ textAlign: "left", padding: "7px 8px", fontSize: 11, fontWeight: 500, color: "#888", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtradas.map(sol => {
                  const fecha = sol.HoraInicio
                    ? new Date(sol.HoraInicio).toLocaleDateString("es-MX", { day: "2-digit", month: "short", year: "numeric" })
                    : "—";
                  const hi = sol.HoraInicio
                    ? new Date(sol.HoraInicio).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" })
                    : "—";
                  const hf = sol.HoraFin
                    ? new Date(sol.HoraFin).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" })
                    : "—";
                  return (
                    <tr key={sol.id}
                      onClick={() => setDetalle(sol)}
                      style={{ borderBottom: "0.5px solid #f0f0f0", cursor: "pointer" }}
                      onMouseEnter={e => e.currentTarget.style.background = "#f8fafc"}
                      onMouseLeave={e => e.currentTarget.style.background = ""}
                    >
                      <td style={{ padding: "9px 8px" }}>
                        <div style={{ fontWeight: 500, color: "#222" }}>{sol.ResponsableDeLaSesi_x00f3_n || "—"}</div>
                        <div style={{ color: "#888", fontSize: 11 }}>{sol.correo}</div>
                      </td>
                      <td style={{ padding: "9px 8px", color: "#555" }}>{sol._x00c1_rea || "—"}</td>
                      <td style={{ padding: "9px 8px", color: "#555" }}>{sol.Compa_x00f1_ia || "—"}</td>
                      <td style={{ padding: "9px 8px", color: "#222", maxWidth: 160 }}>{sol.Asunto}</td>
                      <td style={{ padding: "9px 8px", color: "#555", whiteSpace: "nowrap" }}>{sol.Sala}</td>
                      <td style={{ padding: "9px 8px", color: "#555", whiteSpace: "nowrap" }}>{fecha}</td>
                      <td style={{ padding: "9px 8px", color: "#555", whiteSpace: "nowrap" }}>{hi} – {hf}</td>
                      <td style={{ padding: "9px 8px", color: "#555", textAlign: "center" }}>{sol.Asistentes}</td>
                      <td style={{ padding: "9px 8px" }}>
                        <span style={{
                          ...(ESTADO_STYLE[sol.Estado] || { background: "#f0f0f0", color: "#888" }),
                          padding: "3px 8px", borderRadius: 10, fontWeight: 500, fontSize: 11
                        }}>{sol.Estado || "—"}</span>
                      </td>
                      {isAdmin && (
                        <td style={{ padding: "9px 8px", whiteSpace: "nowrap" }} onClick={e => e.stopPropagation()}>
                          {sol.Estado === "Pendiente" && (
                            <div style={{ display: "flex", gap: 5 }}>
                              <button
                                disabled={procesando === sol.id}
                                onClick={() => aprobar(sol)}
                                style={{ padding: "4px 9px", borderRadius: 6, fontSize: 11, cursor: "pointer", border: "none", background: "#1D9E75", color: "#fff", fontWeight: 500 }}
                              >{procesando === sol.id ? "..." : "✓ Aprobar"}</button>
                              <button
                                disabled={procesando === sol.id}
                                onClick={() => rechazar(sol)}
                                style={{ padding: "4px 9px", borderRadius: 6, fontSize: 11, cursor: "pointer", border: "0.5px solid #E24B4A", background: "#fff", color: "#A32D2D" }}
                              >✕ Rechazar</button>
                            </div>
                          )}
                        </td>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {detalle && (
        <DetalleModal
          sol={detalle}
          onClose={() => setDetalle(null)}
          onAprobar={aprobar}
          onRechazar={rechazar}
          procesando={procesando === detalle?.id}
          isAdmin={isAdmin}
        />
      )}
    </>
  );
}
