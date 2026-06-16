# Control de Salas — Guía de instalación

App web React que se conecta a Microsoft 365 vía Graph API para gestionar reservas de salas de reunión integradas con Teams Rooms (Yealink y otros dispositivos MTR).

---

## Requisitos previos

- Node.js 18+
- Licencias Microsoft 365 con Exchange Online
- Salas configuradas como **buzones de recurso** en Exchange Online
- Acceso al portal de Azure Active Directory (como administrador)

---

## Paso 1 — Registrar la app en Azure AD

1. Ve a **portal.azure.com** → Azure Active Directory → Registros de aplicaciones → **Nueva registro**
2. Nombre: `Control de Salas` (o el que prefieras)
3. Tipos de cuenta: **Cuentas de este directorio organizativo únicamente**
4. URI de redirección: selecciona **SPA** y pon `http://localhost:3000` (en producción agrega tu dominio)
5. Clic en **Registrar**

### Copiar los IDs necesarios

En la página de la app registrada:
- **Id. de aplicación (cliente)** → cópialo como `CLIENT_ID`
- **Id. de directorio (inquilino)** → cópialo como `TENANT_ID`

### Configurar permisos de API

1. Ve a **Permisos de API** → Agregar un permiso → Microsoft Graph → **Permisos delegados**
2. Agrega estos permisos:
   - `Calendars.ReadWrite`
   - `Place.Read.All`
   - `User.Read`
   - `User.ReadBasic.All`
3. Haz clic en **Conceder consentimiento de administrador** (requiere rol de administrador global)

---

## Paso 2 — Configurar buzones de sala en Exchange Online

Si no tienes salas creadas, créalas desde el **Centro de administración de Microsoft 365**:

1. Ve a admin.microsoft.com → Recursos → Salas y equipos → **Agregar sala**
2. Asigna nombre, correo (ej: `sala-a@tuempresa.com`), capacidad y ubicación
3. Activa **Aceptación automática** de reservas

O desde PowerShell:
```powershell
New-Mailbox -Name "Sala A" -DisplayName "Sala A" -Alias sala-a -Room
Set-CalendarProcessing -Identity sala-a -AutomateProcessing AutoAccept
```

---

## Paso 3 — Instalar y configurar la app

```bash
# Clonar / descomprimir el proyecto
cd sala-control

# Instalar dependencias
npm install

# Editar configuración de Azure AD
# Abre src/authConfig.js y reemplaza:
#   TU_CLIENT_ID_AQUI  →  el Client ID del paso 1
#   TU_TENANT_ID_AQUI  →  el Tenant ID del paso 1
```

---

## Paso 4 — Ejecutar en desarrollo

```bash
npm start
# Abre http://localhost:3000
```

Inicia sesión con una cuenta corporativa de tu organización.

---

## Paso 5 — Despliegue en producción

### Opción A: Azure Static Web Apps (recomendado, integra con AAD)

```bash
npm run build
# Sube la carpeta /build a Azure Static Web Apps
```

Agrega el dominio de producción en Azure AD → URI de redirección de la app.

### Opción B: Servidor web cualquiera (nginx, IIS, Netlify, etc.)

```bash
npm run build
# Sirve la carpeta /build como sitio estático
# Configura el servidor para redirigir todas las rutas a index.html (SPA routing)
```

---

## Estructura del proyecto

```
src/
├── authConfig.js          # CLIENT_ID, TENANT_ID y permisos MSAL
├── App.js                 # Login screen + proveedor MSAL
├── pages/
│   └── Dashboard.js       # Página principal con navegación
├── components/
│   ├── RoomCalendar.js    # Vista de calendario con eventos por sala
│   ├── BookingModal.js    # Formulario de nueva reserva
│   └── StatsPanel.js      # Métricas de ocupación (30 días)
├── hooks/
│   └── useRooms.js        # Hook que carga salas y eventos
└── services/
    └── graphService.js    # Todas las llamadas a Graph API
```

---

## Funciones incluidas

| Función | Endpoint Graph API |
|---|---|
| Listar salas | `GET /places/microsoft.graph.room` |
| Ver eventos del día | `GET /users/{sala}/calendarView` |
| Verificar disponibilidad | `POST /me/calendar/getSchedule` |
| Crear reserva | `POST /me/events` |
| Cancelar reserva | `DELETE /me/events/{id}` |
| Estadísticas 30 días | `GET /users/{sala}/calendarView` (rango) |

---

## Preguntas frecuentes

**¿Por qué no veo salas?**
Verifica que existan buzones de recurso de tipo "Room" en Exchange Online y que hayas concedido consentimiento de administrador para `Place.Read.All`.

**¿La reserva aparece en la pantalla Yealink?**
Sí. Las reservas creadas desde esta app van al buzón de la sala en Exchange Online, que es la misma fuente que usa Teams Rooms. La pantalla Yealink se actualiza automáticamente.

**¿Puedo limitar quién puede reservar desde la app?**
Sí, desde el Centro de administración de M365 puedes configurar en cada buzón de sala quién tiene permisos para hacer reservas.
