// ============================================================
//  CONFIGURACIÓN AZURE AD — Universidad Neto
//  Archivo: src/authConfig.js
// ============================================================
export const msalConfig = {
  auth: {
    clientId: "c889b7fa-d0a4-4975-ae68-ed2eb9803445",
    authority: "https://login.microsoftonline.com/e9379df0-6577-491c-ab83-65b8b438c942",
    redirectUri: "https://sistemas-neto.github.io/sala-control-neto",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

// Permisos que solicita la app al usuario
export const loginRequest = {
  scopes: [
    "Calendars.ReadWrite",
    "Place.Read.All",
    "User.Read",
    "User.ReadWrite.All",
    "GroupMember.Read.All",
    "Directory.Read.All",
    "UserAuthenticationMethod.ReadWrite.All",
    "Sites.ReadWrite.All",
  ],
};

// Scopes para llamadas silenciosas a Graph API
export const graphScopes = {
  scopes: [
    "Calendars.ReadWrite",
    "Place.Read.All",
    "User.Read",
    "User.ReadWrite.All",
    "GroupMember.Read.All",
    "Directory.Read.All",
    "UserAuthenticationMethod.ReadWrite.All",
    "Sites.ReadWrite.All",
  ],
};

// ── Grupos de seguridad ──────────────────────────────────────
export const GROUP_ADMINS   = "c839bdc2-ff67-4411-a426-d1de3003acef";
export const GROUP_USUARIOS = "bdd1d693-dbca-45de-84db-3366774dbaa4";

// ── Room List ────────────────────────────────────────────────
export const ROOM_LIST_ID = "c65b9968-0347-4bc1-a4c2-07fa2a2c712a";

// ── Salas conocidas ──────────────────────────────────────────
export const SALAS = [
  {
    nombre: "Sala Entusiasmo",
    email: "entusiasmo@salasneto.com",
    capacidad: 35,
    edificio: "Campus principal",
  },
  {
    nombre: "Sala Practicidad",
    email: "practicidad@salasneto.com",
    capacidad: 20,
    edificio: "Campus principal",
  },
  {
    nombre: "Sala Tenacidad",
    email: "tenacidad@salasneto.com",
    capacidad: 35,
    edificio: "Campus principal",
  },
];
