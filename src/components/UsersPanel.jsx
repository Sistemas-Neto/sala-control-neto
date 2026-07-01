import { useEffect, useState } from "react";
import { useMsal } from "@azure/msal-react";
import {
  getUsers,
  getUserAuthMethods,
  deleteAuthMethod,
  sendPasswordResetLink,
  getUserGroupMembership,
  addUserToGroup,
  removeUserFromGroup,
} from "../services/graphService";
import { GROUP_ADMINS, GROUP_USUARIOS } from "../authConfig";

const METHOD_LABELS = {
  "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod": { label: "Microsoft Authenticator", icon: "📱" },
  "#microsoft.graph.phoneAuthenticationMethod": { label: "Teléfono / SMS", icon: "📞" },
  "#microsoft.graph.passwordAuthenticationMethod": { label: "Contraseña", icon: "🔑" },
  "#microsoft.graph.fido2AuthenticationMethod": { label: "Llave de seguridad", icon: "🔐" },
  "#microsoft.graph.softwareOathAuthenticationMethod": { label: "App OATH", icon: "🔢" },
  "#microsoft.graph.emailAuthenticationMethod": { label: "Correo alternativo", icon: "✉️" },
};

export default function UsersPanel() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const myObjectId = account?.idTokenClaims?.oid;

  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedUser, setSelectedUser] = useState(null);
  const [authMethods, setAuthMethods] = useState([]);
  const [loadingMethods, setLoadingMethods] = useState(false);
  const [actionStatus, setActionStatus] = useState({}); // { [key]: { type, msg } }

  // Roles reales, leídos de los grupos de seguridad: { [userId]: { admin: bool, usuario: bool } }
  const [roles, setRoles] = useState({});
  const [loadingRoles, setLoadingRoles] = useState(true);

  useEffect(() => {
    getUsers(instance, account)
      .then(setUsers)
      .finally(() => setLoading(false));
  }, []);

  // Una vez cargados los usuarios, revisar a qué grupos pertenece cada uno
  useEffect(() => {
    if (users.length === 0) {
      setLoadingRoles(false);
      return;
    }
    setLoadingRoles(true);
    Promise.all(
      users.map(async (u) => {
        const groupIds = await getUserGroupMembership(instance, account, u.id, [GROUP_ADMINS, GROUP_USUARIOS]);
        return [u.id, { admin: groupIds.includes(GROUP_ADMINS), usuario: groupIds.includes(GROUP_USUARIOS) }];
      })
    ).then((pairs) => {
      setRoles(Object.fromEntries(pairs));
      setLoadingRoles(false);
    });
  }, [users]);

  const handleSelectUser = async (user) => {
    if (selectedUser?.id === user.id) {
      setSelectedUser(null);
      setAuthMethods([]);
      return;
    }
    setSelectedUser(user);
    setLoadingMethods(true);
    try {
      const methods = await getUserAuthMethods(instance, account, user.id);
      setAuthMethods(methods);
    } catch (err) {
      setAuthMethods([]);
    } finally {
      setLoadingMethods(false);
    }
  };

  const handleResetPassword = async (user) => {
    if (!window.confirm(`¿Enviar link de reset de contraseña a ${user.displayName}?\n\nSe revocarán sus sesiones activas y recibirá un correo para restablecer su contraseña.`)) return;
    setActionStatus(s => ({ ...s, [user.id]: { type: "loading", msg: "Enviando..." } }));
    try {
      await sendPasswordResetLink(instance, account, user.id);
      setActionStatus(s => ({ ...s, [user.id]: { type: "success", msg: "✓ Sesiones revocadas. El usuario deberá restablecer su contraseña." } }));
    } catch (err) {
      setActionStatus(s => ({ ...s, [user.id]: { type: "error", msg: "Error: " + err.message } }));
    }
    setTimeout(() => setActionStatus(s => { const n = {...s}; delete n[user.id]; return n; }), 4000);
  };

  const handleDeleteMfa = async (method) => {
    if (method["@odata.type"] === "#microsoft.graph.passwordAuthenticationMethod") return;
    const label = METHOD_LABELS[method["@odata.type"]]?.label || "método";
    if (!window.confirm(`¿Eliminar "${label}" de ${selectedUser.displayName}?\n\nEl usuario deberá registrar nuevamente este método.`)) return;

    const key = `mfa_${method.id}`;
    setActionStatus(s => ({ ...s, [key]: { type: "loading", msg: "Eliminando..." } }));
    try {
      await deleteAuthMethod(instance, account, selectedUser.id, method.id, method["@odata.type"]);
      setAuthMethods(m => m.filter(x => x.id !== method.id));
      setActionStatus(s => ({ ...s, [key]: { type: "success", msg: "✓ Método eliminado" } }));
    } catch (err) {
      setActionStatus(s => ({ ...s, [key]: { type: "error", msg: "Error: " + err.message } }));
    }
    setTimeout(() => setActionStatus(s => { const n = {...s}; delete n[key]; return n; }), 3000);
  };

  // Asignar / quitar un rol (pertenencia a grupo). roleKey: "admin" | "usuario"
  const handleToggleRole = async (user, roleKey, groupId) => {
    const currentlyIn = !!roles[user.id]?.[roleKey];
    const roleLabel = roleKey === "admin" ? "Administrador" : "Usuario (aprobador)";

    // Protección: evitar que te quites el rol de admin a ti mismo por error
    if (roleKey === "admin" && currentlyIn && user.id === myObjectId) {
      if (!window.confirm("⚠️ Estás a punto de quitarte el rol de Administrador A TI MISMO. Podrías perder acceso a esta sección. ¿Continuar de todas formas?")) return;
    } else {
      const accion = currentlyIn ? "quitar" : "asignar";
      if (!window.confirm(`¿${accion === "quitar" ? "Quitar" : "Asignar"} el rol "${roleLabel}" ${accion === "quitar" ? "a" : "a"} ${user.displayName}?`)) return;
    }

    const key = `role_${user.id}_${roleKey}`;
    setActionStatus(s => ({ ...s, [key]: { type: "loading", msg: "..." } }));
    try {
      if (currentlyIn) {
        await removeUserFromGroup(instance, account, groupId, user.id);
      } else {
        await addUserToGroup(instance, account, groupId, user.id);
      }
      setRoles(r => ({ ...r, [user.id]: { ...r[user.id], [roleKey]: !currentlyIn } }));
      setActionStatus(s => ({ ...s, [key]: { type: "success", msg: "✓" } }));
    } catch (err) {
      setActionStatus(s => ({ ...s, [key]: { type: "error", msg: "Error: " + err.message } }));
    }
    setTimeout(() => setActionStatus(s => { const n = {...s}; delete n[key]; return n; }), 3000);
  };

  return (
    <div style={s.wrap}>
      <div style={s.header}>
        <div style={s.title}>👥 Usuarios con acceso</div>
        <div style={s.subtitle}>Gestiona roles, contraseñas y métodos de autenticación MFA</div>
      </div>

      {loading ? (
        <div style={s.loading}>Cargando usuarios...</div>
      ) : (
        <div style={s.layout}>
          {/* Lista de usuarios */}
          <div style={s.userList}>
            {users.map(user => {
              const status = actionStatus[user.id];
              const isSelected = selectedUser?.id === user.id;
              const roleInfo = roles[user.id] || {};
              const badgeLabel = roleInfo.admin ? "Admin" : roleInfo.usuario ? "Usuario" : "Sin rol";
              const badgeStyle = roleInfo.admin ? s.bRed : roleInfo.usuario ? s.bGreen : s.bGray;

              return (
                <div key={user.id} style={{ ...s.userCard, ...(isSelected ? s.userCardSelected : {}) }}>
                  <div style={s.userRow} onClick={() => handleSelectUser(user)}>
                    <div style={s.userAvatar}>{user.displayName?.[0]?.toUpperCase() || "U"}</div>
                    <div style={s.userInfo}>
                      <div style={s.userName}>{user.displayName}</div>
                      <div style={s.userEmail}>{user.userPrincipalName}</div>
                    </div>
                    <div style={s.userRight}>
                      <span style={{ ...s.badge, ...badgeStyle }}>
                        {loadingRoles ? "…" : badgeLabel}
                      </span>
                      <span style={{ ...s.dot, background: user.accountEnabled ? "#1D9E75" : "#aaa" }} title={user.accountEnabled ? "Activo" : "Deshabilitado"} />
                    </div>
                  </div>

                  {/* Control de roles */}
                  <div style={s.rolesRow} onClick={e => e.stopPropagation()}>
                    <span style={s.rolesLabel}>Rol:</span>
                    <button
                      style={{ ...s.rolePill, ...(roleInfo.admin ? s.rolePillAdminActive : {}) }}
                      disabled={loadingRoles || actionStatus[`role_${user.id}_admin`]?.type === "loading"}
                      onClick={() => handleToggleRole(user, "admin", GROUP_ADMINS)}
                    >
                      {roleInfo.admin ? "✓ Admin" : "+ Admin"}
                    </button>
                    <button
                      style={{ ...s.rolePill, ...(roleInfo.usuario ? s.rolePillUsuarioActive : {}) }}
                      disabled={loadingRoles || actionStatus[`role_${user.id}_usuario`]?.type === "loading"}
                      onClick={() => handleToggleRole(user, "usuario", GROUP_USUARIOS)}
                    >
                      {roleInfo.usuario ? "✓ Usuario" : "+ Usuario"}
                    </button>
                    {(actionStatus[`role_${user.id}_admin`] || actionStatus[`role_${user.id}_usuario`]) && (
                      <span style={{
                        fontSize: 10,
                        color: (actionStatus[`role_${user.id}_admin`]?.type === "error" || actionStatus[`role_${user.id}_usuario`]?.type === "error") ? "#A32D2D" : "#1D9E75",
                      }}>
                        {actionStatus[`role_${user.id}_admin`]?.msg || actionStatus[`role_${user.id}_usuario`]?.msg}
                      </span>
                    )}
                  </div>

                  {/* Acciones rápidas */}
                  <div style={s.actions}>
                    <button
                      style={s.btnReset}
                      onClick={() => handleResetPassword(user)}
                      disabled={actionStatus[user.id]?.type === "loading"}
                    >
                      🔑 Enviar link de reset
                    </button>
                    <button
                      style={{ ...s.btnMfa, ...(isSelected ? s.btnMfaActive : {}) }}
                      onClick={() => handleSelectUser(user)}
                    >
                      📱 {isSelected ? "Ocultar MFA" : "Ver MFA"}
                    </button>
                  </div>

                  {/* Feedback de acción de reset */}
                  {status && (
                    <div style={{ ...s.feedback, ...(status.type === "error" ? s.feedbackError : status.type === "success" ? s.feedbackSuccess : s.feedbackLoading) }}>
                      {status.msg}
                    </div>
                  )}

                  {/* Métodos de autenticación */}
                  {isSelected && (
                    <div style={s.mfaSection}>
                      <div style={s.mfaTitle}>Métodos de autenticación registrados</div>
                      {loadingMethods ? (
                        <div style={s.mfaLoading}>Cargando métodos...</div>
                      ) : authMethods.length === 0 ? (
                        <div style={s.mfaEmpty}>Sin métodos registrados</div>
                      ) : (
                        authMethods.map(method => {
                          const info = METHOD_LABELS[method["@odata.type"]] || { label: method["@odata.type"], icon: "🔒" };
                          const isPassword = method["@odata.type"] === "#microsoft.graph.passwordAuthenticationMethod";
                          const mfaKey = `mfa_${method.id}`;
                          const mfaStatus = actionStatus[mfaKey];
                          return (
                            <div key={method.id} style={s.methodRow}>
                              <span style={s.methodIcon}>{info.icon}</span>
                              <span style={s.methodLabel}>{info.label}</span>
                              {mfaStatus ? (
                                <span style={{ fontSize: 10, color: mfaStatus.type === "error" ? "#A32D2D" : "#1D9E75" }}>
                                  {mfaStatus.msg}
                                </span>
                              ) : !isPassword && (
                                <button
                                  style={s.btnDelete}
                                  onClick={() => handleDeleteMfa(method)}
                                >
                                  Eliminar
                                </button>
                              )}
                            </div>
                          );
                        })
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          {/* Panel de info */}
          <div style={s.infoPanel}>
            <div style={s.infoPanelTitle}>ℹ️ Guía de acciones</div>
            <div style={s.infoItem}>
              <div style={s.infoIcon}>🏷️</div>
              <div>
                <div style={s.infoLabel}>Asignar / quitar rol</div>
                <div style={s.infoDesc}><strong>Admin</strong> da acceso total (Usuarios, Configuración, aprobar solicitudes). <strong>Usuario</strong> permite aprobar/rechazar solicitudes de reserva sin el resto de permisos de administrador.</div>
              </div>
            </div>
            <div style={s.infoItem}>
              <div style={s.infoIcon}>🔑</div>
              <div>
                <div style={s.infoLabel}>Enviar link de reset</div>
                <div style={s.infoDesc}>Revoca las sesiones activas del usuario. Recibirá un correo de Azure AD para restablecer su contraseña (requiere SSPR habilitado en el tenant).</div>
              </div>
            </div>
            <div style={s.infoItem}>
              <div style={s.infoIcon}>📱</div>
              <div>
                <div style={s.infoLabel}>Ver / Eliminar MFA</div>
                <div style={s.infoDesc}>Muestra los métodos de autenticación registrados. Puedes eliminar el Microsoft Authenticator u otros métodos si el usuario perdió acceso a su dispositivo.</div>
              </div>
            </div>
            <div style={s.infoItem}>
              <div style={s.infoIcon}>⚠️</div>
              <div>
                <div style={s.infoLabel}>Permisos requeridos</div>
                <div style={s.infoDesc}>Estas acciones requieren <strong>UserAuthenticationMethod.ReadWrite.All</strong> y <strong>GroupMember.ReadWrite.All</strong> con consentimiento de administrador en Azure AD.</div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const s = {
  wrap: { display: "flex", flexDirection: "column", gap: 12 },
  header: { marginBottom: 4 },
  title: { fontSize: 14, fontWeight: 500, color: "#222", marginBottom: 3 },
  subtitle: { fontSize: 12, color: "#888" },
  loading: { padding: 32, textAlign: "center", color: "#888", fontSize: 13 },
  layout: { display: "grid", gridTemplateColumns: "1fr 280px", gap: 12, alignItems: "start" },
  userList: { display: "flex", flexDirection: "column", gap: 8 },
  userCard: { background: "#fff", border: "0.5px solid #eee", borderRadius: 10, padding: "12px 14px", transition: "all .12s" },
  userCardSelected: { border: "1px solid #185FA5", background: "#f8fbff" },
  userRow: { display: "flex", alignItems: "center", gap: 10, cursor: "pointer" },
  userAvatar: { width: 36, height: 36, borderRadius: "50%", background: "#042C53", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 500, flexShrink: 0 },
  userInfo: { flex: 1 },
  userName: { fontSize: 13, fontWeight: 500, color: "#222" },
  userEmail: { fontSize: 11, color: "#888", marginTop: 2 },
  userRight: { display: "flex", alignItems: "center", gap: 7 },
  badge: { fontSize: 10, padding: "2px 7px", borderRadius: 10, fontWeight: 500 },
  bRed: { background: "#FCEBEB", color: "#A32D2D" },
  bGreen: { background: "#E1F5EE", color: "#085041" },
  bGray: { background: "#f0f0f0", color: "#888" },
  dot: { width: 8, height: 8, borderRadius: "50%", flexShrink: 0 },
  rolesRow: { display: "flex", alignItems: "center", gap: 6, marginTop: 9, flexWrap: "wrap" },
  rolesLabel: { fontSize: 11, color: "#888", marginRight: 2 },
  rolePill: { padding: "3px 9px", borderRadius: 12, fontSize: 11, border: "0.5px solid #ddd", background: "#f8f8f8", color: "#666", cursor: "pointer" },
  rolePillAdminActive: { border: "0.5px solid #E24B4A", background: "#FCEBEB", color: "#A32D2D", fontWeight: 500 },
  rolePillUsuarioActive: { border: "0.5px solid #1D9E75", background: "#E1F5EE", color: "#085041", fontWeight: 500 },
  actions: { display: "flex", gap: 7, marginTop: 10, paddingTop: 10, borderTop: "0.5px solid #f0f0f0" },
  btnReset: { flex: 1, padding: "6px 10px", border: "0.5px solid #B5D4F4", borderRadius: 7, fontSize: 11, background: "#E6F1FB", color: "#0C447C", cursor: "pointer", fontWeight: 500 },
  btnMfa: { flex: 1, padding: "6px 10px", border: "0.5px solid #ddd", borderRadius: 7, fontSize: 11, background: "#f8f8f8", color: "#555", cursor: "pointer" },
  btnMfaActive: { borderColor: "#185FA5", background: "#E6F1FB", color: "#0C447C" },
  feedback: { marginTop: 8, padding: "6px 10px", borderRadius: 7, fontSize: 11 },
  feedbackLoading: { background: "#f8f8f8", color: "#888" },
  feedbackSuccess: { background: "#E1F5EE", color: "#085041" },
  feedbackError: { background: "#FCEBEB", color: "#A32D2D" },
  mfaSection: { marginTop: 10, paddingTop: 10, borderTop: "0.5px solid #f0f0f0" },
  mfaTitle: { fontSize: 11, fontWeight: 500, color: "#666", marginBottom: 8 },
  mfaLoading: { fontSize: 11, color: "#aaa", padding: "6px 0" },
  mfaEmpty: { fontSize: 11, color: "#aaa", padding: "6px 0" },
  methodRow: { display: "flex", alignItems: "center", gap: 8, padding: "6px 8px", background: "#f8f8f8", borderRadius: 7, marginBottom: 5 },
  methodIcon: { fontSize: 14, flexShrink: 0 },
  methodLabel: { flex: 1, fontSize: 12, color: "#333" },
  btnDelete: { padding: "3px 8px", border: "0.5px solid #E24B4A", borderRadius: 6, fontSize: 10, color: "#A32D2D", background: "none", cursor: "pointer" },
  infoPanel: { background: "#fff", border: "0.5px solid #eee", borderRadius: 10, padding: "14px 16px", position: "sticky", top: 12 },
  infoPanelTitle: { fontSize: 12, fontWeight: 500, color: "#555", marginBottom: 12 },
  infoItem: { display: "flex", gap: 10, marginBottom: 12 },
  infoIcon: { fontSize: 18, flexShrink: 0 },
  infoLabel: { fontSize: 12, fontWeight: 500, color: "#222", marginBottom: 3 },
  infoDesc: { fontSize: 11, color: "#888", lineHeight: 1.5 },
};
