// ============================================================
//  APP.JS — Punto de entrada con autenticación MSAL
// ============================================================

import { MsalProvider, AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./authConfig";
import Dashboard from "./pages/Dashboard";

// Instancia MSAL global (crear fuera del componente)
const msalInstance = new PublicClientApplication(msalConfig);

export default function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <AuthenticatedTemplate>
        <Dashboard />
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <LoginScreen />
      </UnauthenticatedTemplate>
    </MsalProvider>
  );
}

function LoginScreen() {
  const { instance } = useMsal();

  const handleLogin = () => {
    instance.loginRedirect(loginRequest).catch(console.error);
  };

  return (
    <div style={loginStyles.page}>
      <div style={loginStyles.card}>
        <div style={loginStyles.icon}>🏢</div>
        <h1 style={loginStyles.title}>Control de Salas</h1>
        <p style={loginStyles.subtitle}>
          Gestiona las salas de reunión de tu organización
        </p>
        <button onClick={handleLogin} style={loginStyles.btn}>
          <MicrosoftIcon />
          Iniciar sesión con Microsoft
        </button>
        <p style={loginStyles.note}>
          Se requiere cuenta corporativa con acceso a Microsoft 365
        </p>
      </div>
    </div>
  );
}

function MicrosoftIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 21 21" style={{ flexShrink: 0 }}>
      <rect x="1" y="1" width="9" height="9" fill="#f25022" />
      <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
      <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
      <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
    </svg>
  );
}

const loginStyles = {
  page: {
    minHeight: "100vh", display: "flex", alignItems: "center",
    justifyContent: "center", background: "#f0f4fa",
    fontFamily: "'Segoe UI', system-ui, sans-serif",
  },
  card: {
    background: "#fff", borderRadius: 16, padding: "48px 40px",
    textAlign: "center", boxShadow: "0 4px 32px rgba(0,0,0,0.1)",
    maxWidth: 360, width: "100%",
  },
  icon: { fontSize: 48, marginBottom: 12 },
  title: { fontSize: 22, fontWeight: 700, color: "#1a5fa8", margin: "0 0 8px" },
  subtitle: { fontSize: 14, color: "#888", margin: "0 0 28px", lineHeight: 1.5 },
  btn: {
    display: "flex", alignItems: "center", justifyContent: "center",
    gap: 10, width: "100%", padding: "12px 20px",
    background: "#fff", border: "1px solid #ddd", borderRadius: 8,
    fontSize: 14, fontWeight: 500, cursor: "pointer", color: "#222",
    transition: "background 0.15s",
  },
  note: { fontSize: 11, color: "#bbb", marginTop: 16 },
};
