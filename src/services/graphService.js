import { graphScopes } from "../authConfig";

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

export async function getRoomEvents(msalInstance, account, roomEmail, start, end) {
  return [];
}

/**
 * Carga los eventos de todas las salas dentro de un RANGO de fechas,
 * consultando el calendario de CADA SALA directamente
 * (/users/{roomEmail}/calendarView) en vez del calendario del usuario logueado.
 *
 * Esto requiere que el usuario tenga permiso "Reviewer" (o superior) sobre
 * el calendario de cada sala en Exchange Online — ver grant-calendar-permissions.ps1.
 * Así, TODAS las reservas de la sala aparecen sin importar quién las creó.
 *
 * Acepta dos formas de uso (compatibilidad hacia atrás):
 *   getAllRoomsEvents(instance, account, rooms, singleDate)
 *   getAllRoomsEvents(instance, account, rooms, startDate, endDate)
 */
export async function getAllRoomsEvents(msalInstance, account, rooms, startDate, endDate) {
  const start = new Date(startDate);
  start.setHours(0, 0, 0, 0);

  const end = new Date(endDate || startDate);
  end.setHours(23, 59, 59, 999);

  const startStr = toLocalISOString(start);
  const endStr = toLocalISOString(end);

  const results = await Promise.all(
    rooms.map(async (room) => {
      try {
        const data = await callGraph(
          msalInstance,
          account,
          `/users/${encodeURIComponent(room.emailAddress)}/calendarView?startDateTime=${startStr}&endDateTime=${endStr}&$select=id,subject,start,end,organizer,location,locations,attendees,bodyPreview&$top=999&$orderby=start/dateTime`,
          { headers: { "Prefer": 'outlook.timezone="America/Mexico_City"' } }
        );
        return { roomId: room.id, roomEmail: room.emailAddress, events: data.value || [] };
      } catch (err) {
        console.error(`Error cargando eventos de ${room.displayName}:`, err);
        return { roomId: room.id, roomEmail: room.emailAddress, events: [] };
      }
    })
  );

  return results;
}

export async function checkAvailability(msalInstance, account, roomEmail, start, end) {
  try {
    const body = {
      schedules: [roomEmail],
      startTime: { dateTime: toLocalISOString(start), timeZone: "America/Mexico_City" },
      endTime:   { dateTime: toLocalISOString(end),   timeZone: "America/Mexico_City" },
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

export async function createBooking(msalInstance, account, booking) {
  const { subject, roomEmail, roomName, start, end, comments = "" } = booking;

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
    ],
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness",
  };

  // Crear el evento
  const created = await callGraph(msalInstance, account, "/me/events", {
    method: "POST",
    body: JSON.stringify(event),
  });

  // Obtener el evento completo con datos de Teams
  if (created?.id) {
    try {
      const full = await callGraph(
        msalInstance,
        account,
        `/me/events/${created.id}?$select=id,subject,onlineMeeting,onlineMeetingUrl`
      );
      return { ...created, onlineMeeting: full.onlineMeeting, onlineMeetingUrl: full.onlineMeetingUrl };
    } catch {
      return created;
    }
  }

  return created;
}

export async function createComboBooking(msalInstance, account, booking) {
  const { subject, roomNames, roomEmails, start, end, comments = "" } = booking;

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
    location: { displayName: "Sala Magna (Tenacidad + Entusiasmo)" },
    attendees: [
      ...roomEmails.map((email, i) => ({
        emailAddress: { address: email, name: roomNames[i] || email },
        type: "resource",
      })),
    ],
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness",
  };

  const created = await callGraph(msalInstance, account, "/me/events", {
    method: "POST",
    body: JSON.stringify(event),
  });

  if (created?.id) {
    try {
      const full = await callGraph(
        msalInstance,
        account,
        `/me/events/${created.id}?$select=id,subject,onlineMeeting,onlineMeetingUrl`
      );
      return { ...created, onlineMeeting: full.onlineMeeting, onlineMeetingUrl: full.onlineMeetingUrl };
    } catch {
      return created;
    }
  }

  return created;
}

export async function cancelBooking(msalInstance, account, eventId) {
  return callGraph(msalInstance, account, `/me/events/${eventId}`, { method: "DELETE" });
}

/**
 * Cancela una reserva borrando el evento directamente del calendario de LA SALA
 * (en vez del calendario personal de quien lo creó). Esto funciona sin importar
 * quién haya aprobado/creado la reserva originalmente, siempre que la cuenta que
 * cancela tenga FullAccess + Calendars.ReadWrite.Shared sobre esa sala.
 */
export async function cancelBookingFromRoom(msalInstance, account, roomEmail, eventId) {
  return callGraph(msalInstance, account, `/users/${encodeURIComponent(roomEmail)}/events/${eventId}`, { method: "DELETE" });
}

export async function updateBooking(msalInstance, account, eventId, booking) {
  const { subject, roomEmail, roomName, start, end, comments = "" } = booking;

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
    location: { displayName: roomName, locationEmailAddress: roomEmail },
    attendees: [
      { emailAddress: { address: roomEmail, name: roomName }, type: "resource" },
    ],
  };

  return callGraph(msalInstance, account, `/me/events/${eventId}`, {
    method: "PATCH",
    body: JSON.stringify(event),
  });
}

export async function getRoomStats(msalInstance, account, roomEmail, customStart, customEnd) {
  const end = customEnd || new Date();
  const start = customStart || new Date(new Date().setDate(new Date().getDate() - 30));

  const NORMALIZE = {
    "salapracticidad@soyneto.onmicrosoft.com": "practicidad@salasneto.com",
    "salatenacidad@soyneto.onmicrosoft.com":   "tenacidad@salasneto.com",
    "salaentusiasmo@soyneto.onmicrosoft.com":  "entusiasmo@salasneto.com",
  };
  const normalizeRoom = (email) => NORMALIZE[email?.toLowerCase()] || email?.toLowerCase();
  const normalizedEmail = normalizeRoom(roomEmail);

  try {
    const data = await callGraph(
      msalInstance, account,
      `/me/calendarView?startDateTime=${toLocalISOString(start)}&endDateTime=${toLocalISOString(end)}&$select=id,subject,start,end,attendees,location,locations&$top=100&$orderby=start/dateTime`,
      { headers: { "Prefer": 'outlook.timezone="America/Mexico_City"' } }
    );

    const events = (data.value || []).filter(ev => {
      const loc = normalizeRoom(ev.location?.locationEmailAddress);
      if (loc === normalizedEmail) return true;
      if ((ev.locations || []).some(l => normalizeRoom(l.locationEmailAddress) === normalizedEmail)) return true;
      if ((ev.attendees || []).some(a => normalizeRoom(a.emailAddress?.address) === normalizedEmail)) return true;
      return false;
    });

    let totalMinutes = 0;
    events.forEach(ev => { totalMinutes += (new Date(ev.end.dateTime) - new Date(ev.start.dateTime)) / 60000; });

    return {
      totalBookings: events.length,
      totalHours: Math.round(totalMinutes / 60),
      occupancyPct: Math.min(Math.round((totalMinutes / (8 * 60 * 22)) * 100), 100),
    };
  } catch {
    return { totalBookings: 0, totalHours: 0, occupancyPct: 0 };
  }
}

export async function getMagnaStats(msalInstance, account, customStart, customEnd) {
  const end = customEnd || new Date();
  const start = customStart || new Date(new Date().setDate(new Date().getDate() - 30));

  const MAGNA_ROOMS = ["tenacidad@salasneto.com", "entusiasmo@salasneto.com"];
  const NORMALIZE = {
    "salapracticidad@soyneto.onmicrosoft.com": "practicidad@salasneto.com",
    "salatenacidad@soyneto.onmicrosoft.com":   "tenacidad@salasneto.com",
    "salaentusiasmo@soyneto.onmicrosoft.com":  "entusiasmo@salasneto.com",
  };
  const normalizeRoom = (email) => NORMALIZE[email?.toLowerCase()] || email?.toLowerCase();

  try {
    const data = await callGraph(
      msalInstance, account,
      `/me/calendarView?startDateTime=${toLocalISOString(start)}&endDateTime=${toLocalISOString(end)}&$select=id,subject,start,end,attendees,location,locations&$top=100&$orderby=start/dateTime`,
      { headers: { "Prefer": 'outlook.timezone="America/Mexico_City"' } }
    );

    const magnaEvents = (data.value || []).filter(ev => {
      const attendeeEmails = (ev.attendees || []).map(a => normalizeRoom(a.emailAddress?.address));
      return MAGNA_ROOMS.every(r => attendeeEmails.includes(r));
    });

    let totalMinutes = 0;
    magnaEvents.forEach(ev => { totalMinutes += (new Date(ev.end.dateTime) - new Date(ev.start.dateTime)) / 60000; });

    return {
      totalBookings: magnaEvents.length,
      totalHours: Math.round(totalMinutes / 60),
      occupancyPct: Math.min(Math.round((totalMinutes / (8 * 60 * 22)) * 100), 100),
    };
  } catch {
    return { totalBookings: 0, totalHours: 0, occupancyPct: 0 };
  }
}

// Fecha de vencimiento real de la suscripción (compra: 25/06/2026, 1 año de vigencia)
const TEAMS_ROOMS_EXPIRY = "2027-06-25";

function calcularDiasRestantes(expiryDateStr) {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const vencimiento = new Date(expiryDateStr + "T00:00:00");
  const diffMs = vencimiento - hoy;
  const dias = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
  return dias;
}

export async function getTeamsRoomsLicenses(msalInstance, account) {
  const diasRestantes = calcularDiasRestantes(TEAMS_ROOMS_EXPIRY);
  const expiryLabel = new Date(TEAMS_ROOMS_EXPIRY + "T00:00:00").toLocaleDateString("es-MX", {
    day: "2-digit", month: "long", year: "numeric"
  });

  try {
    const data = await callGraph(msalInstance, account, "/subscribedSkus");
    const teamsRoomSkus = (data.value || []).filter(s =>
      s.skuPartNumber?.toLowerCase().includes("teams_rooms") ||
      s.skuPartNumber?.toLowerCase().includes("mtr")
    );

    if (teamsRoomSkus.length === 0) {
      return [{
        name: "Microsoft Teams Rooms Pro",
        skuPartNumber: "Microsoft_Teams_Rooms_Pro",
        total: 3,
        consumed: 3,
        available: 0,
        expiryDate: TEAMS_ROOMS_EXPIRY,
        expiryLabel,
        diasRestantes,
        isMock: true,
      }];
    }

    return teamsRoomSkus.map(s => ({
      name: s.skuPartNumber.replace(/_/g, " "),
      skuPartNumber: s.skuPartNumber,
      total: s.prepaidUnits?.enabled || 0,
      consumed: s.consumedUnits || 0,
      available: (s.prepaidUnits?.enabled || 0) - (s.consumedUnits || 0),
      expiryDate: TEAMS_ROOMS_EXPIRY,
      expiryLabel,
      diasRestantes,
      isMock: false,
    }));
  } catch {
    return [{
      name: "Microsoft Teams Rooms Pro",
      skuPartNumber: "Microsoft_Teams_Rooms_Pro",
      total: 3,
      consumed: 3,
      available: 0,
      expiryDate: TEAMS_ROOMS_EXPIRY,
      expiryLabel,
      diasRestantes,
      isMock: true,
    }];
  }
}

export async function getUsers(msalInstance, account) {
  try {
    const data = await callGraph(msalInstance, account, "/users?$select=id,displayName,userPrincipalName,accountEnabled,createdDateTime&$top=50");
    return data.value || [];
  } catch {
    return [];
  }
}

export async function getUserAuthMethods(msalInstance, account, userId) {
  try {
    const data = await callGraph(msalInstance, account, `/users/${userId}/authentication/methods`);
    return data.value || [];
  } catch (err) {
    throw new Error(err.message || "Error al obtener métodos de autenticación");
  }
}

export async function deleteAuthMethod(msalInstance, account, userId, methodId, methodType) {
  const endpoints = {
    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod": "microsoftAuthenticatorMethods",
    "#microsoft.graph.phoneAuthenticationMethod": "phoneMethods",
    "#microsoft.graph.fido2AuthenticationMethod": "fido2Methods",
    "#microsoft.graph.softwareOathAuthenticationMethod": "softwareOathMethods",
  };
  const path = endpoints[methodType] || "microsoftAuthenticatorMethods";
  return callGraph(msalInstance, account, `/users/${userId}/authentication/${path}/${methodId}`, { method: "DELETE" });
}

export async function sendPasswordResetLink(msalInstance, account, userId) {
  await callGraph(msalInstance, account, `/users/${userId}/revokeSignInSessions`, { method: "POST", body: JSON.stringify({}) });
  return true;
}

// ── Gestión de roles (pertenencia a grupos de seguridad) ────────────

/**
 * Revisa a cuáles de los groupIds dados pertenece un usuario.
 * Devuelve un array con los IDs de los grupos a los que SÍ pertenece.
 * Requiere GroupMember.Read.All / Directory.Read.All (ya delegados).
 */
export async function getUserGroupMembership(msalInstance, account, userId, groupIds) {
  try {
    const data = await callGraph(msalInstance, account, `/users/${userId}/checkMemberGroups`, {
      method: "POST",
      body: JSON.stringify({ groupIds }),
    });
    return data.value || [];
  } catch (err) {
    console.error("Error verificando grupos del usuario:", err);
    return [];
  }
}

/**
 * Agrega un usuario a un grupo de seguridad.
 * Requiere GroupMember.ReadWrite.All con consentimiento de administrador.
 */
export async function addUserToGroup(msalInstance, account, groupId, userId) {
  return callGraph(msalInstance, account, `/groups/${groupId}/members/$ref`, {
    method: "POST",
    body: JSON.stringify({
      "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`,
    }),
  });
}

/**
 * Quita a un usuario de un grupo de seguridad.
 * Requiere GroupMember.ReadWrite.All con consentimiento de administrador.
 */
export async function removeUserFromGroup(msalInstance, account, groupId, userId) {
  return callGraph(msalInstance, account, `/groups/${groupId}/members/${userId}/$ref`, {
    method: "DELETE",
  });
}
