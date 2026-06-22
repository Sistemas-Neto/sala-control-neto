import { useState } from "react";
import { useMsal } from "@azure/msal-react";
import { format, addDays, subDays } from "date-fns";
import { es } from "date-fns/locale";
import { useRooms } from "../hooks/useRooms";
import RoomCalendar from "../components/RoomCalendar";
import BookingModal from "../components/BookingModal";
import StatsPanel from "../components/StatsPanel";
import { GROUP_ADMINS } from "../authConfig";

export default function Dashboard() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const [selectedDate, setSelectedDate] = useState(new Date());
  const [activeView, setActiveView] = useState("dashboard");
  const [showModal, setShowModal] = useState(false);
  const { rooms, events, loading, error, refresh } = useRooms(selectedDate);

  const isAdmin = account?.idTokenClaims?.groups?.includes(GROUP_ADMINS);
  const handleLogout = () => instance.logoutRedirect();

  const VIEWS = {
    dashboard: "Dashboard de salas",
    calendario: "Calendario semanal",
    salas: "Gestión de salas",
    estadisticas: "Estadísticas de uso",
    exportar: "Exportar reportes",
    usuarios: "Usuarios con acceso",
    configuracion: "Configuración",
  };

  const initials = (name) => name ? name.split(" ").map(n => n[0]).join("").substring(0,2).toUpperCase() : "UN";

  // Calcular ocupación simulada basada en eventos
  const getOcupacion = (roomEmail) => {
    const evs = events[roomEmail] || [];
    return Math.min(Math.round((evs.length / 8) * 100), 100);
  };

  const getRoomStatus = (roomEmail) => {
    const evs = events[roomEmail] || [];
    const now = new Date();
    const busy = evs.some(ev => new Date(ev.start.dateTime) <= now && now <= new Date(ev.end.dateTime));
    return busy ? "En uso" : "Libre";
  };

  const getNextEvents = () => {
    const now = new Date();
    const all = [];
    rooms.forEach(r => {
      const evs = events[r.emailAddress] || [];
      evs.forEach(ev => {
        if (new Date(ev.start.dateTime) > now) {
          all.push({ ...ev, roomName: r.displayName });
        }
      });
    });
    return all.sort((a, b) => new Date(a.start.dateTime) - new Date(b.start.dateTime)).slice(0, 3);
  };

  return (
    <div style={s.app}>
      {/* SIDEBAR */}
      <div style={s.sb}>
        <div style={s.sbTop}>
          <div style={s.sbLabel}>Universidad Neto</div>
          <div style={s.sbTitle}>Control de salas</div>
          <div style={s.sbSub}>Campus principal</div>
        </div>
        <div style={s.sbDiv}/>
        <nav style={s.sbNav}>
          <div style={s.sbSec}>Principal</div>
          <SbItem label="Dashboard" id="dashboard" active={activeView} onClick={setActiveView}/>
          <SbItem label="Calendario" id="calendario" active={activeView} onClick={setActiveView}/>
          <SbItem label="Salas" id="salas" active={activeView} onClick={setActiveView}/>
          <div style={s.sbSec}>Análisis</div>
          <SbItem label="Estadísticas" id="estadisticas" active={activeView} onClick={setActiveView}/>
          <SbItem label="Exportar" id="exportar" active={activeView} onClick={setActiveView}/>
          {isAdmin && (<>
            <div style={s.sbSec}>Administración</div>
            <SbItem label="Usuarios" id="usuarios" active={activeView} onClick={setActiveView}/>
            <SbItem label="Configuración" id="configuracion" active={activeView} onClick={setActiveView}/>
          </>)}
        </nav>
        <div style={s.sbFoot}>
          <div style={s.sbAv}>{initials(account?.name)}</div>
          <div>
            <div style={s.sbName}>{account?.name || account?.username}</div>
            <div style={s.sbRole}>{isAdmin ? "Administrador" : "Usuario"}</div>
          </div>
        </div>
      </div>

      {/* MAIN */}
      <div style={s.main}>
        <div style={s.topbar}>
          <div style={s.tbInfo}>
            <div style={s.tbTitle}>{VIEWS[activeView]}</div>
            <div style={s.tbSub}>{format(selectedDate, "EEEE d 'de' MMMM, yyyy", { locale: es })}</div>
          </div>
          <button style={s.navBtn} onClick={() => setSelectedDate(d => subDays(d,1))}>‹</button>
          <span style={s.datePill}>{format(selectedDate, "EEE d MMM", { locale: es })}</span>
          <button style={s.navBtn} onClick={() => setSelectedDate(d => addDays(d,1))}>›</button>
          <button style={s.todayBtn} onClick={() => setSelectedDate(new Date())}>Hoy</button>
          <button style={s.newBtn} onClick={() => setShowModal(true)}>+ Nueva reserva</button>
          <button style={s.logoutBtn} onClick={handleLogout}>Salir</button>
        </div>

        <div style={s.content}>

          {/* ── DASHBOARD ── */}
          {activeView === "dashboard" && (
            <div>
              <div style={s.kpis}>
                <KPI n={rooms.length} label="Salas totales" color="#185FA5" bg="#E6F1FB" delta="Todas operativas" dc="#1D9E75"/>
                <KPI n={rooms.filter(r => getRoomStatus(r.emailAddress) === "Libre").length} label="Disponibles ahora" color="#1D9E75" bg="#E1F5EE" delta="Listas para reservar" dc="#1D9E75"/>
                <KPI n={rooms.filter(r => getRoomStatus(r.emailAddress) === "En uso").length} label="En uso ahora" color="#E24B4A" bg="#FCEBEB" delta={rooms.find(r => getRoomStatus(r.emailAddress) === "En uso")?.displayName || "—"} dc="#888"/>
                <KPI n={Object.values(events).reduce((s,evs) => s + evs.length, 0)} label="Reservas hoy" color="#534AB7" bg="#EEEDFE" delta="Total del día" dc="#534AB7"/>
              </div>

              {loading && <p style={s.msg}>Cargando salas...</p>}
              {error && <p style={{...s.msg, color:"#c0392b"}}>Error: {error}</p>}

              {!loading && !error && (
                <div style={s.bodyGrid}>
                  {/* Calendario */}
                  <div>
                    <RoomCalendar rooms={rooms} events={events} onRefresh={refresh}/>
                  </div>

                  {/* Panel derecho */}
                  <div style={s.rightCol}>
                    {/* Estado de salas */}
                    <div style={s.mini}>
                      <div style={s.miniTitle}>🏢 Estado de salas</div>
                      {rooms.map(r => (
                        <div key={r.id} style={s.roomRow}>
                          <span style={s.rrName}>{r.displayName}</span>
                          <span style={{...s.badge, ...(getRoomStatus(r.emailAddress) === "En uso" ? s.bRed : s.bGreen)}}>
                            {getRoomStatus(r.emailAddress)}
                          </span>
                        </div>
                      ))}
                      <div style={{...s.roomRow, borderTop:"0.5px solid #eee", marginTop:4, paddingTop:6}}>
                        <span style={{...s.rrName, fontSize:10, color:"#534AB7"}}>🔗 Sala Magna</span>
                        <span style={{...s.badge, ...s.bPurple}}>Disponible</span>
                      </div>
                    </div>

                    {/* Ocupación del mes */}
                    <div style={s.mini}>
                      <div style={s.miniTitle}>📊 Ocupación del mes</div>
                      {rooms.map(r => {
                        const pct = getOcupacion(r.emailAddress);
                        const color = pct > 75 ? "#E24B4A" : pct > 40 ? "#BA7517" : "#1D9E75";
                        return (
                          <div key={r.id} style={{marginBottom:7}}>
                            <div style={{display:"flex",justifyContent:"space-between",fontSize:10,marginBottom:2}}>
                              <span style={{color:"#888"}}>{r.displayName}</span>
                              <span style={{fontWeight:500,color:"#222"}}>{pct}%</span>
                            </div>
                            <div style={s.barT}><div style={{...s.barF, width:`${pct}%`, background:color}}/></div>
                          </div>
                        );
                      })}
                    </div>

                    {/* Próximas reservas */}
                    <div style={s.mini}>
                      <div style={s.miniTitle}>🕐 Próximas reservas</div>
                      {getNextEvents().length === 0 ? (
                        <div style={{fontSize:11,color:"#aaa",textAlign:"center",padding:"8px 0"}}>Sin reservas próximas</div>
                      ) : getNextEvents().map((ev, i) => (
                        <div key={i} style={s.nr}>
                          <div style={s.nrT}>{ev.subject || "(Sin asunto)"}</div>
                          <div style={s.nrM}>🕐 {format(new Date(ev.start.dateTime), "HH:mm")} · {ev.roomName}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ── ESTADÍSTICAS ── */}
          {activeView === "estadisticas" && <StatsPanel rooms={rooms}/>}

          {/* ── SALAS ── */}
          {activeView === "salas" && (
            <div style={s.card}>
              <div style={s.cardTitle}>🏢 Gestión de salas</div>
              <div style={s.roomGrid}>
                {rooms.map(r => (
                  <div key={r.id} style={s.roomCard}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                      <div style={s.roomName}>{r.displayName}</div>
                      <span style={{...s.badge,...(getRoomStatus(r.emailAddress)==="En uso"?s.bRed:s.bGreen)}}>{getRoomStatus(r.emailAddress)}</span>
                    </div>
                    <div style={s.roomMeta}>Cap. {r.capacity || "—"} · {r.building || "Campus principal"}</div>
                    <div style={{...s.barT,marginTop:7}}><div style={{...s.barF,width:`${getOcupacion(r.emailAddress)}%`,background:"#E24B4A"}}/></div>
                  </div>
                ))}
                <div style={{...s.roomCard,background:"#F0EAF7",border:"0.5px solid #AFA9EC"}}>
                  <div style={{fontSize:12,fontWeight:500,color:"#3C3489",marginBottom:4}}>🔗 Sala Magna (combinada)</div>
                  <div style={{fontSize:10,color:"#534AB7"}}>Tenacidad + Entusiasmo · Cap. 70 · Se activa al reservar como sala combinada</div>
                </div>
              </div>
            </div>
          )}

          {/* ── EXPORTAR ── */}
          {activeView === "exportar" && (
            <div style={s.card}>
              <div style={s.cardTitle}>📁 Exportar reportes</div>
              <div style={s.secTitle}>Ocupación y uso</div>
              {["Reporte mensual de ocupación","Horas pico por sala","Uso de sala combinada (Magna)"].map(r => (
                <div key={r} style={s.repRow}>
                  <div style={s.repName}>{r}</div>
                  <div style={s.dlBtns}>
                    <button style={s.dlXl}>⬇ Excel</button>
                    <button style={s.dlPdf}>⬇ PDF</button>
                  </div>
                </div>
              ))}
              <div style={{...s.secTitle,marginTop:14}}>Historial de reservas</div>
              {["Historial completo de reservas","Cancelaciones y no-shows","Reservas recurrentes"].map(r => (
                <div key={r} style={s.repRow}>
                  <div style={s.repName}>{r}</div>
                  <div style={s.dlBtns}>
                    <button style={s.dlXl}>⬇ Excel</button>
                    <button style={s.dlCsv}>⬇ CSV</button>
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* ── CALENDARIO ── */}
          {activeView === "calendario" && (
            <div style={s.card}>
              <div style={s.cardTitle}>📅 Vista semanal</div>
              {!loading && !error && <RoomCalendar rooms={rooms} events={events} onRefresh={refresh}/>}
            </div>
          )}

          {/* ── USUARIOS ── */}
          {activeView === "usuarios" && isAdmin && (
            <div style={s.card}>
              <div style={s.cardTitle}>👥 Usuarios con acceso</div>
              <table style={s.tbl}>
                <thead><tr><th style={s.th}>Nombre</th><th style={s.th}>Correo</th><th style={s.th}>Rol</th><th style={s.th}>Último acceso</th></tr></thead>
                <tbody>
                  <tr><td style={s.td}>Admin TI</td><td style={s.td}>sistemasneto@soyneto.onmicrosoft.com</td><td style={s.td}><span style={{...s.badge,...s.bRed}}>Admin</span></td><td style={s.td}>Hoy</td></tr>
                  <tr><td style={s.td}>Capacitación</td><td style={s.td}>capacitacion@soyneto.onmicrosoft.com</td><td style={s.td}><span style={{...s.badge,...s.bGreen}}>Usuario</span></td><td style={s.td}>—</td></tr>
                </tbody>
              </table>
            </div>
          )}

          {/* ── CONFIGURACIÓN ── */}
          {activeView === "configuracion" && isAdmin && (
            <div style={s.card}>
              <div style={s.cardTitle}>⚙️ Configuración general</div>
              <div style={{marginBottom:13}}>
                <div style={s.formSecTitle}>Organización</div>
                <div style={s.fgrid}>
                  <div><div style={s.fl}>Nombre</div><input style={s.fi} defaultValue="Universidad Neto"/></div>
                  <div><div style={s.fl}>Campus</div><input style={s.fi} defaultValue="Campus principal"/></div>
                  <div><div style={s.fl}>Zona horaria</div><select style={s.fi}><option>América/Mexico_City (UTC-6)</option></select></div>
                  <div><div style={s.fl}>Horario reservas</div><select style={s.fi}><option>07:00 – 21:00</option></select></div>
                </div>
              </div>
              <div style={{marginBottom:13}}>
                <div style={s.formSecTitle}>Integración Microsoft 365</div>
                <div style={s.fgrid}>
                  <div><div style={s.fl}>Tenant ID</div><input style={s.fi} defaultValue="e9379df0-6577-491c-ab83-65b8b438c942" readOnly/></div>
                  <div><div style={s.fl}>Client ID</div><input style={s.fi} defaultValue="c889b7fa-d0a4-4975-ae68-ed2eb9803445" readOnly/></div>
                </div>
                <div style={{marginTop:7,display:"flex",alignItems:"center",gap:6,fontSize:11,color:"#666"}}>
                  <div style={{width:7,height:7,borderRadius:"50%",background:"#1D9E75"}}/>
                  Conectado a Azure AD — sincronización activa
                </div>
              </div>
              <button style={s.newBtn}>✓ Guardar cambios</button>
            </div>
          )}
        </div>
      </div>

      {/* MODAL */}
      {showModal && rooms.length > 0 && (
        <BookingModal rooms={rooms} selectedDate={selectedDate} onClose={() => setShowModal(false)} onSuccess={refresh}/>
      )}
    </div>
  );
}

function SbItem({ label, id, active, onClick }) {
  return (
    <div onClick={() => onClick(id)} style={{
      display:"flex",alignItems:"center",gap:8,padding:"6px 9px",
      borderRadius:6,fontSize:12,
      color: active===id ? "#fff" : "#B5D4F4",
      cursor:"pointer",marginBottom:1,
      background: active===id ? "rgba(255,255,255,0.13)" : "none",
    }}>
      {label}
    </div>
  );
}

function KPI({ n, label, color, bg, delta, dc }) {
  return (
    <div style={{background:"#fff",border:"0.5px solid #eee",borderRadius:10,padding:"9px 11px"}}>
      <div style={{width:26,height:26,borderRadius:6,background:bg,marginBottom:6}}/>
      <div style={{fontSize:19,fontWeight:500,color,lineHeight:1}}>{n}</div>
      <div style={{fontSize:10,color:"#888",marginTop:2}}>{label}</div>
      <div style={{fontSize:10,color:dc,marginTop:2}}>{delta}</div>
    </div>
  );
}

const s = {
  app:{display:"flex",fontFamily:"'Segoe UI',system-ui,sans-serif",minHeight:"100vh",background:"#f0f3f8"},
  sb:{width:190,background:"#042C53",display:"flex",flexDirection:"column",minHeight:"100vh",flexShrink:0},
  sbTop:{padding:"14px 13px 11px"},
  sbLabel:{fontSize:10,color:"#85B7EB",letterSpacing:".08em",textTransform:"uppercase",fontWeight:500,marginBottom:3},
  sbTitle:{fontSize:14,fontWeight:500,color:"#fff",lineHeight:1.2},
  sbSub:{fontSize:10,color:"#85B7EB",marginTop:2},
  sbDiv:{height:"0.5px",background:"rgba(255,255,255,0.1)",margin:"0 13px"},
  sbNav:{padding:"9px 7px",flex:1},
  sbSec:{fontSize:10,color:"#85B7EB",letterSpacing:".06em",textTransform:"uppercase",padding:"7px 8px 3px",fontWeight:500},
  sbFoot:{padding:"9px 13px",borderTop:"0.5px solid rgba(255,255,255,0.1)",marginTop:"auto",display:"flex",alignItems:"center",gap:8},
  sbAv:{width:26,height:26,borderRadius:"50%",background:"#185FA5",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:500,color:"#E6F1FB",flexShrink:0},
  sbName:{fontSize:11,color:"#fff",fontWeight:500},
  sbRole:{fontSize:10,color:"#85B7EB"},
  main:{flex:1,display:"flex",flexDirection:"column",minWidth:0},
  topbar:{background:"#fff",borderBottom:"0.5px solid #eee",padding:"8px 13px",display:"flex",alignItems:"center",gap:7,flexShrink:0},
  tbInfo:{flex:1},
  tbTitle:{fontSize:13,fontWeight:500,color:"#222"},
  tbSub:{fontSize:10,color:"#888"},
  navBtn:{background:"none",border:"0.5px solid #ddd",borderRadius:6,width:24,height:24,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",color:"#666",flexShrink:0},
  datePill:{fontSize:11,fontWeight:500,color:"#222",padding:"3px 9px",background:"#f5f5f5",borderRadius:20,border:"0.5px solid #eee",whiteSpace:"nowrap"},
  todayBtn:{fontSize:11,padding:"3px 7px",borderRadius:6,border:"0.5px solid #B5D4F4",background:"#E6F1FB",color:"#0C447C",cursor:"pointer"},
  newBtn:{display:"flex",alignItems:"center",gap:4,padding:"5px 11px",background:"#042C53",color:"#E6F1FB",border:"none",borderRadius:6,fontSize:11,fontWeight:500,cursor:"pointer",flexShrink:0},
  logoutBtn:{padding:"5px 10px",border:"0.5px solid #ddd",borderRadius:6,fontSize:11,background:"none",color:"#888",cursor:"pointer"},
  content:{padding:11,flex:1},
  kpis:{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:7,marginBottom:10},
  bodyGrid:{display:"grid",gridTemplateColumns:"1fr 188px",gap:9},
  rightCol:{display:"flex",flexDirection:"column",gap:8},
  mini:{background:"#fff",border:"0.5px solid #eee",borderRadius:10,padding:"9px 11px"},
  miniTitle:{fontSize:11,fontWeight:500,color:"#888",marginBottom:7},
  roomRow:{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"4px 0",borderBottom:"0.5px solid #eee"},
  rrName:{fontSize:11,color:"#222"},
  badge:{fontSize:10,padding:"2px 6px",borderRadius:10,fontWeight:500},
  bRed:{background:"#FCEBEB",color:"#A32D2D"},
  bGreen:{background:"#E1F5EE",color:"#085041"},
  bBlue:{background:"#E6F1FB",color:"#0C447C"},
  bPurple:{background:"#EEEDFE",color:"#3C3489"},
  barT:{height:4,background:"#eee",borderRadius:2,overflow:"hidden",marginTop:3},
  barF:{height:"100%",borderRadius:2},
  nr:{padding:"6px 8px",background:"#f8f8f8",borderRadius:8,cursor:"pointer",marginBottom:4},
  nrT:{fontSize:11,fontWeight:500,color:"#222"},
  nrM:{fontSize:10,color:"#888",marginTop:1},
  msg:{padding:20,textAlign:"center",color:"#888",fontSize:13},
  card:{background:"#fff",border:"0.5px solid #eee",borderRadius:10,padding:"14px 16px"},
  cardTitle:{fontSize:13,fontWeight:500,color:"#222",marginBottom:12},
  roomGrid:{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8},
  roomCard:{background:"#f8f8f8",borderRadius:8,padding:"10px 12px"},
  roomName:{fontSize:12,fontWeight:500,color:"#222"},
  roomMeta:{fontSize:10,color:"#888",marginTop:3},
  secTitle:{fontSize:11,fontWeight:500,color:"#888",textTransform:"uppercase",letterSpacing:".06em",marginBottom:8,paddingBottom:6,borderBottom:"0.5px solid #eee"},
  repRow:{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 11px",background:"#f8f8f8",borderRadius:8,marginBottom:6},
  repName:{fontSize:11,fontWeight:500,color:"#222"},
  dlBtns:{display:"flex",gap:5},
  dlXl:{padding:"4px 8px",border:"0.5px solid #1D6F42",borderRadius:6,fontSize:10,color:"#1D6F42",background:"none",cursor:"pointer"},
  dlPdf:{padding:"4px 8px",border:"0.5px solid #A32D2D",borderRadius:6,fontSize:10,color:"#A32D2D",background:"none",cursor:"pointer"},
  dlCsv:{padding:"4px 8px",border:"0.5px solid #185FA5",borderRadius:6,fontSize:10,color:"#185FA5",background:"none",cursor:"pointer"},
  tbl:{width:"100%",borderCollapse:"collapse",fontSize:12},
  th:{textAlign:"left",padding:"7px 9px",fontSize:10,fontWeight:500,color:"#888",borderBottom:"0.5px solid #eee"},
  td:{padding:"7px 9px",borderBottom:"0.5px solid #eee",color:"#222"},
  fgrid:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8},
  formSecTitle:{fontSize:11,fontWeight:500,color:"#888",marginBottom:7,paddingBottom:5,borderBottom:"0.5px solid #eee"},
  fl:{fontSize:10,color:"#888",marginBottom:2},
  fi:{width:"100%",padding:"6px 9px",border:"0.5px solid #ddd",borderRadius:6,fontSize:12,boxSizing:"border-box"},
};
