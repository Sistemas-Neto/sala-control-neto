import { useState, useEffect } from "react";
import { useMsal } from "@azure/msal-react";
import { getSolicitudes, actualizarEstado } from "../services/solicitudesService";
import { createBooking, createComboBooking } from "../services/graphService";
import { GROUP_ADMINS } from "../authConfig";

export default function SolicitudesPanel() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const isAdmin = account?.idTokenClaims?.groups?.includes(GROUP_ADMINS);
  const [solicitudes, setSolicitudes] = useState([]);
  const [loading, setLoading] = useState(true);
  const [procesando, setProcesando] = useState(null);
  const [filtro, setFiltro] = useState("Pendiente");

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
      // Crear el evento en el calendario de la sala
      const SALAS_EMAIL = {
        "Sala Tenacidad":   "tenacidad@salasneto.com",
        "Sala Practicidad": "practicidad@salasneto.com",
        "Sala Entusiasmo":  "entusiasmo@salasneto.com",
      };

      if (sol.Sala === "Sala Magna") {
        await createComboBooking(instance, account, {
          subject:    sol.Asunto,
          roomNames:  ["Sala Tenacidad", "Sala Entusiasmo"],
          roomEmails: ["tenacidad@salasneto.com", "entusiasmo@salasneto.com"],
          start:      sol.HoraInicio,
          end:        sol.HoraFin,
          attendees:  [sol.correo],
          comments:   sol.RequerimientosAdicionales || "",
        });
      } else {
        await createBooking(instance, account, {
          subject:   sol.Asunto,
          roomEmail: SALAS_EMAIL[sol.Sala] || "",
          roomName:  sol.Sala,
          start:     sol.HoraInicio,
          end:       sol.HoraFin,
          attendees: [sol.correo],
          comments:  sol.RequerimientosAdicionales || "",
        });
      }

      // Actualizar estado en SharePoint
      await actualizarEstado(instance, account, sol.id, "Aprobado");
      await cargar();
    } catch (e) {
      alert("Error al aprobar: " + e.message);
    }
    setProcesando(null);
  };

  const rechazar = async (sol) => {
    if (!isAdmin) return;
    if (!window.confirm(`¿Rechazar solicitud de ${sol.ResponsableDeLaSesi_x00f3_n}?`)) return;
    setProcesando(sol.id);
    try {
      await actualizarEstado(instance, account, sol.id, "Rechazado");
      await cargar();
    } catch (e) {
      alert("Error al rechazar: " + e.message);
    }
    setProcesando(null);
  };

  const ESTADO_STYLE = {
    Pendiente:  { background: "#FFF8E1", color: "#B45309" },
    Aprobado:   { background: "#E1F5EE", color: "#085041" },
    Rechazado:  { background: "#FCEBEB", color: "#A32D2D" },
  };

  return (
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
          <button onClick={cargar} style={{ padding: "4px 10px", borderRadius: 6, fontSize: 11, cursor: "pointer", border: "0.5px solid #ddd", background: "#f8f8f8", color: "#666" }}>↻ Actualizar</button>
        </div>
      </div>

      {loading ? (
        <div style={{ textAlign: "center", padding: 30, color: "#888", fontSize: 13 }}>Cargando solicitudes...</div>
      ) : filtradas.length === 0 ? (
        <div style={{ textAlign: "center", padding: 30, color: "#aaa", fontSize: 13 }}>No hay solicitudes {filtro !== "Todas" ? filtro.toLowerCase() + "s" : ""}</div>
      ) : (
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ borderBottom: "0.5px solid #eee" }}>
              {["Responsable", "Área", "Compañía", "Asunto", "Sala", "Fecha", "Horario", "Asistentes", "Estado", isAdmin ? "Acciones" : ""].map(h => (
                <th key={h} style={{ textAlign: "left", padding: "7px 8px", fontSize: 11, fontWeight: 500, color: "#888" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtradas.map(sol => {
              const fecha = sol.HoraInicio ? new Date(sol.HoraInicio).toLocaleDateString("es-MX", { day: "2-digit", month: "short", year: "numeric" }) : "—";
              const hi = sol.HoraInicio ? new Date(sol.HoraInicio).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" }) : "—";
              const hf = sol.HoraFin    ? new Date(sol.HoraFin).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit" }) : "—";
              return (
                <tr key={sol.id} style={{ borderBottom: "0.5px solid #f0f0f0" }}>
                  <td style={{ padding: "9px 8px" }}>
                    <div style={{ fontWeight: 500, color: "#222" }}>{sol.ResponsableDeLaSesi_x00f3_n}</div>
                    <div style={{ color: "#888", fontSize: 11 }}>{sol.correo}</div>
                  </td>
                  <td style={{ padding: "9px 8px", color: "#555" }}>{sol._x00c1_rea || "—"}</td>
                  <td style={{ padding: "9px 8px", color: "#555" }}>{sol.Compa_x00f1_ia || "—"}</td>
                  <td style={{ padding: "9px 8px", color: "#222", maxWidth: 160 }}>{sol.Asunto}</td>
                  <td style={{ padding: "9px 8px", color: "#555" }}>{sol.Sala}</td>
                  <td style={{ padding: "9px 8px", color: "#555", whiteSpace: "nowrap" }}>{fecha}</td>
                  <td style={{ padding: "9px 8px", color: "#555", whiteSpace: "nowrap" }}>{hi} – {hf}</td>
                  <td style={{ padding: "9px 8px", color: "#555", textAlign: "center" }}>{sol.Asistentes}</td>
                  <td style={{ padding: "9px 8px" }}>
                    <span style={{ ...ESTADO_STYLE[sol.Estado], padding: "3px 8px", borderRadius: 10, fontWeight: 500, fontSize: 11 }}>
                      {sol.Estado || "—"}
                    </span>
                  </td>
                  {isAdmin && (
                    <td style={{ padding: "9px 8px", whiteSpace: "nowrap" }}>
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
      )}
    </div>
  );
}
