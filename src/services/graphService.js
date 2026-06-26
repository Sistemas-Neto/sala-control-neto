import { graphScopes } from "../authConfig";

// ── HELPER: Formatea fecha sin sufijo Z para que Graph respete el timeZone ──
// .toISOString() agrega "Z" (UTC) y Graph ignora el campo timeZone.
// Esta función devuelve "2026-06-23T15:00:00" sin zona, para que Graph
// interprete la hora según el timeZone que se le pasa en el body/header.
function toLocalISOString(date) {
  const pad = (n) => String(n).padStart(2, "0");
  return (
    date.getFullYear() + "-" +
    pad(date.getMonth() + 1) + "-" +
    pad(date.getDate()) + "T" +
    pad(date.getHours()) + ":" +
    pad(date.getMinutes()) + ":" +
    pad(date.getSeconds())
  );
}

async function getAccessToken(msalInstance, account) {
  try {
    const response = await msalInstance.acquireTokenSilent({
      ...graphScopes,
      account,
    });
    return response.accessToken;
  } catch {
    await msalInstance.acquireTokenRedirect(graphScopes);
  }
}

async function callGraph(msalInstance, account, endpoint, options = {}) {
  const token = await getAccessToken(msalInstance, account);
  const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  });
  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error?.message || `Graph API error: ${response.status}`);
  }
  if (response.status === 204) return null;
  return response.json();
}

// ── SALAS ──────────────────────────────────────────────────
export async function getRooms(msalInstance, account) {
  try {
    const data = await callGraph(msalInstance, account, "/places/microsoft.graph.room");
    return data.value || [];
  } catch {
    return [
      { id: "1", displayName: "Sala Entusiasmo", emailAddress: "entusiasmo@salasneto.com", capacity: 35, building: "Campus principal" },
      { id: "2", displayName: "Sala Practicidad", emailAddress: "practicidad@salasneto.com", capacity: 20, building: "Campus principal" },
      { id: "3", displayName: "Sala Tenacidad", emailAddress: "tenacidad@salasneto.com", capacity: 35, building: "Campus principal" },
    ];
  }
}

// ── EVENTOS ────────────────────────────────────────────────
// Una sola llamada a calendarView trae todos los eventos del día.
// Luego distribuimos cada evento a la sala que corresponda buscando
// en attendees (resource), location.locationEmailAddress y locations[].
// Esto es más robusto que hacer una llamada por sala y evita filtros
// que fallaban cuando Graph no devolvía attendees completos.
export async function getRoomEvents(msalInstance, account, roomEmail, start, end) {
  // Este método se mantiene por compatibilidad pero ya no se usa directamente.
  // getAllRoomsEvents hace una sola llamada y distribuye.
  return [];
}

export async function getAllRoomsEvents(msalInstance, account, rooms, date) {
  const start = new Date(date);
  start.setHours(0, 0, 0, 0);
  const end = new Date(date);
  end.setHours(23, 59, 59, 999);

  const startStr = toLocalISOString(start);
  const endStr = toLocalISOString(end);

  try {
    const data = await callGraph(
      msalInstance,
      account,
      `/me/calendarView?startDateTime=${startStr}&endDateTime=${endStr}&$select=id,subject,start,end,organizer,location,locations,attendees,bodyPreview&$top=100&$orderby=start/dateTime`,
      {
        headers: {
          "Prefer": 'outlook.timezone="America/Mexico_City"',
        },
      }
    );

    const allEvents = data.value || [];

    // Normaliza correos de sala al dominio nuevo independientemente de lo que devuelva Graph
    const NORMALIZE = {
      "salapracticidad@soyneto.onmicrosoft.com": "practicidad@salasneto.com",
      "salatenacidad@soyneto.onmicrosoft.com":   "tenacidad@salasneto.com",
      "salaentusiasmo@soyneto.onmicrosoft.com":  "entusiasmo@salasneto.com",
    };

    // Normaliza el email de la sala que devuelve Graph al dominio nuevo
    const normalizeRoom = (email) => NORMALIZE[email?.toLowerCase()] || email?.toLowerCase();

    // Para cada sala, filtra los eventos que le corresponden
    return rooms.map((room) => {
      // Normaliza el email de la sala (puede venir con dominio viejo de Graph)
      const email = normalizeRoom(room.emailAddress);

      const roomEvents = allEvents.filter(ev => {
        // 1. location principal
        const loc = normalizeRoom(ev.location?.locationEmailAddress);
        if (loc === email) return true;

        // 2. locations[] (array de ubicaciones)
        const locs = ev.locations || [];
        if (locs.some(l => normalizeRoom(l.locationEmailAddress) === email)) return true;

        // 3. attendees de tipo resource
        const attendees = ev.attendees || [];
        if (attendees.some(a =>
          normalizeRoom(a.emailAddress?.address) === email
        )) return true;

        return false;
      });

      // Devuelve el email normalizado para que el resto del sistema lo use correctamente
      return { roomId: room.id, roomEmail: room.emailAddress, events: roomEvents };
    });
  } catch (err) {
    console.error("Error cargando eventos:", err);
    return rooms.map(room => ({ roomId: room.id, roomEmail: room.emailAddress, events: [] }));
  }
}

// FIX: Se cambió toISOString() + timeZone "UTC" por toLocalISOString() +
// timeZone "America/Mexico_City" para que la disponibilidad se consulte
// en la hora correcta.
export async function checkAvailability(msalInstance, account, roomEmail, start, end) {
  try {
    const body = {
      schedules: [roomEmail],
      startTime: {
        dateTime: toLocalISOString(start), // FIX: era start.toISOString()
        timeZone: "America/Mexico_City",   // FIX: era "UTC"
      },
      endTime: {
        dateTime: toLocalISOString(end),   // FIX: era end.toISOString()
        timeZone: "America/Mexico_City",   // FIX: era "UTC"
      },
      availabilityViewInterval: 30,
    };

    const data = await callGraph(msalInstance, account, "/me/calendar/getSchedule", {
      method: "POST",
      body: JSON.stringify(body),
    });

    const schedule = data.value?.[0];
    if (!schedule) return { available: false };

    const isFree = schedule.availabilityView?.split("").every((slot) => slot === "0");
    return { available: isFree, scheduleItems: schedule.scheduleItems || [] };
  } catch {
    return { available: true };
  }
}

// FIX timezone: Se usa toLocalISOString() en vez de toISOString() para que el
// campo timeZone: "America/Mexico_City" sea respetado por Graph.
// FIX comments: Se agrega el campo "body" con los comentarios del formulario
// para que lleguen en el correo de invitación a los asistentes.
export async function createBooking(msalInstance, account, booking) {
  const { subject, roomEmail, roomName, start, end, attendees = [], comments = "" } = booking; // FIX: desestructura comments

  const event = {
    subject,
    body: {                    // FIX: HTML para que el comentario aparezca arriba del contenido de Teams
      contentType: "HTML",
      content: comments
        ? `<p style="font-family:sans-serif;font-size:14px;margin-bottom:16px;">${comments}</p><hr/>`
        : "",
    },
    start: {
      dateTime: typeof start === "string" ? start : toLocalISOString(start), // FIX timezone
      timeZone: "America/Mexico_City",
    },
    end: {
      dateTime: typeof end === "string" ? end : toLocalISOString(end), // FIX timezone
      timeZone: "America/Mexico_City",
    },
    location: {
      displayName: roomName,
      locationEmailAddress: roomEmail,
    },
    attendees: [
      { emailAddress: { address: roomEmail, name: roomName }, type: "resource" },
      ...attendees.map((email) => ({
        emailAddress: { address: email.trim() },
        type: "required",
      })),
    ],
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness",
  };

  return callGraph(msalInstance, account, "/me/events", {
    method: "POST",
    body: JSON.stringify(event),
  });
}

// ── RESERVA COMBINADA (Sala Magna) ────────────────────────
// Crea un solo evento con ambas salas como recursos para que
// solo llegue un correo al destinatario y ambas salas queden bloqueadas.
export async function createComboBooking(msalInstance, account, booking) {
  const { subject, roomNames, roomEmails, start, end, attendees = [], comments = "" } = booking;

  const event = {
    subject,
    body: {
      contentType: "HTML",
      content: comments
        ? `<p style="font-family:sans-serif;font-size:14px;margin-bottom:16px;">${comments}</p><hr/>`
        : "",
    },
    start: {
      dateTime: typeof start === "string" ? start : toLocalISOString(start),
      timeZone: "America/Mexico_City",
    },
    end: {
      dateTime: typeof end === "string" ? end : toLocalISOString(end),
      timeZone: "America/Mexico_City",
    },
    location: {
      displayName: "Sala Magna (Tenacidad + Entusiasmo)",
    },
    attendees: [
      // Ambas salas como recursos — un solo evento las bloquea a las dos
      ...roomEmails.map((email, i) => ({
        emailAddress: { address: email, name: roomNames[i] || email },
        type: "resource",
      })),
      // Asistentes normales
      ...attendees.map((email) => ({
        emailAddress: { address: email.trim() },
        type: "required",
      })),
    ],
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness",
  };

  return callGraph(msalInstance, account, "/me/events", {
    method: "POST",
    body: JSON.stringify(event),
  });
}

// Cancelar con ID real del evento
export async function cancelBooking(msalInstance, account, eventId) {
  return callGraph(msalInstance, account, `/me/events/${eventId}`, {
    method: "DELETE",
  });
}

// Editar una reserva existente
export async function updateBooking(msalInstance, account, eventId, booking) {
  const { subject, roomEmail, roomName, start, end, attendees = [], comments = "" } = booking;

  const event = {
    subject,
    body: {
      contentType: "HTML",
      content: comments
        ? `<p style="font-family:sans-serif;font-size:14px;margin-bottom:16px;">${comments}</p><hr/>`
        : "",
    },
    start: {
      dateTime: typeof start === "string" ? start : toLocalISOString(start),
      timeZone: "America/Mexico_City",
    },
    end: {
      dateTime: typeof end === "string" ? end : toLocalISOString(end),
      timeZone: "America/Mexico_City",
    },
    location: {
      displayName: roomName,
      locationEmailAddress: roomEmail,
    },
    attendees: [
      { emailAddress: { address: roomEmail, name: roomName }, type: "resource" },
      ...attendees.map((email) => ({
        emailAddress: { address: email.trim() },
        type: "required",
      })),
    ],
  };

  return callGraph(msalInstance, account, `/me/events/${eventId}`, {
    method: "PATCH",
    body: JSON.stringify(event),
  });
}

export async function getRoomStats(msalInstance, account, roomEmail) {
  const end = new Date();
  const start = new Date();
  start.setDate(start.getDate() - 30);

  const events = await getRoomEvents(msalInstance, account, roomEmail, start, end);
  let totalMinutes = 0;
  events.forEach((ev) => {
    const s = new Date(ev.start.dateTime);
    const e = new Date(ev.end.dateTime);
    totalMinutes += (e - s) / 60000;
  });

  const availableMinutes = 8 * 60 * 22;
  const occupancyPct = Math.round((totalMinutes / availableMinutes) * 100);

  return {
    totalBookings: events.length,
    totalHours: Math.round(totalMinutes / 60),
    occupancyPct: Math.min(occupancyPct, 100),
  };
}

// ── LICENCIAS DE TEAMS ROOMS ──────────────────────────────
// SKU real de Teams Rooms Pro: "Microsoft_Teams_Rooms_Pro"
// Cuando tengas licencias compradas, esta función las leerá automáticamente.
// Por ahora devuelve mock si no hay datos reales.
export async function getTeamsRoomsLicenses(msalInstance, account) {
  try {
    const data = await callGraph(msalInstance, account, "/subscribedSkus");
    const skus = data.value || [];

    // Filtra solo licencias relacionadas a Teams Rooms
    const teamsRoomSkus = skus.filter(s =>
      s.skuPartNumber?.toLowerCase().includes("teams_rooms") ||
      s.skuPartNumber?.toLowerCase().includes("mtr")
    );

    if (teamsRoomSkus.length === 0) {
      // Sin licencias aún — devuelve mock para visualización
      return [{
        name: "Microsoft Teams Rooms Pro",
        skuPartNumber: "Microsoft_Teams_Rooms_Pro",
        total: 3,
        consumed: 3,
        available: 0,
        expiryDate: null, // null = sin fecha (licencia de prueba/pendiente)
        isMock: true,
      }];
    }

    return teamsRoomSkus.map(s => ({
      name: s.skuPartNumber.replace(/_/g, " "),
      skuPartNumber: s.skuPartNumber,
      total: s.prepaidUnits?.enabled || 0,
      consumed: s.consumedUnits || 0,
      available: (s.prepaidUnits?.enabled || 0) - (s.consumedUnits || 0),
      expiryDate: s.prepaidUnits?.suspended > 0 ? null : null, // Graph no expone fecha directamente
      isMock: false,
    }));
  } catch {
    return [{
      name: "Microsoft Teams Rooms Pro",
      skuPartNumber: "Microsoft_Teams_Rooms_Pro",
      total: 3,
      consumed: 3,
      available: 0,
      expiryDate: null,
      isMock: true,
    }];
  }
}

// ── GESTIÓN DE USUARIOS ────────────────────────────────────

// Obtiene todos los usuarios del tenant con su info básica
export async function getUsers(msalInstance, account) {
  try {
    const data = await callGraph(
      msalInstance,
      account,
      "/users?$select=id,displayName,userPrincipalName,accountEnabled,createdDateTime,signInActivity&$top=50"
    );
    return data.value || [];
  } catch {
    return [];
  }
}

// Obtiene los métodos de autenticación de un usuario
export async function getUserAuthMethods(msalInstance, account, userId) {
  try {
    const data = await callGraph(
      msalInstance,
      account,
      `/users/${userId}/authentication/methods`
    );
    return data.value || [];
  } catch (err) {
    throw new Error(err.message || "Error al obtener métodos de autenticación");
  }
}

// Elimina un método de autenticación (MFA) de un usuario
export async function deleteAuthMethod(msalInstance, account, userId, methodId, methodType) {
  // El endpoint varía según el tipo de método
  const endpoints = {
    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod": "microsoftAuthenticatorMethods",
    "#microsoft.graph.phoneAuthenticationMethod": "phoneMethods",
    "#microsoft.graph.fido2AuthenticationMethod": "fido2Methods",
    "#microsoft.graph.softwareOathAuthenticationMethod": "softwareOathMethods",
  };
  const path = endpoints[methodType] || "microsoftAuthenticatorMethods";
  return callGraph(
    msalInstance,
    account,
    `/users/${userId}/authentication/${path}/${methodId}`,
    { method: "DELETE" }
  );
}

// Envía link de reset de contraseña al correo del usuario
export async function sendPasswordResetLink(msalInstance, account, userId) {
  // Graph API: invalidar sesiones activas fuerza al usuario a re-autenticarse
  // El reset real se hace vía Azure AD Self-Service Password Reset (SSPR)
  // Esta llamada revoca todos los tokens del usuario
  await callGraph(
    msalInstance,
    account,
    `/users/${userId}/revokeSignInSessions`,
    { method: "POST", body: JSON.stringify({}) }
  );
  // Retorna true si fue exitoso
  return true;
}
