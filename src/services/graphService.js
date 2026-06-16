// ============================================================
//  SERVICIO DE GRAPH API
//  Todas las llamadas a Microsoft Graph para salas y calendarios
// ============================================================

import { graphScopes } from "../authConfig";

/**
 * Obtiene un token de acceso de forma silenciosa (sin re-login)
 */
async function getAccessToken(msalInstance, account) {
  try {
    const response = await msalInstance.acquireTokenSilent({
      ...graphScopes,
      account,
    });
    return response.accessToken;
  } catch {
    // Si falla el token silencioso, redirige al login
    await msalInstance.acquireTokenRedirect(graphScopes);
  }
}

/**
 * Helper base para todas las llamadas a Graph API
 */
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

  // DELETE devuelve 204 sin body
  if (response.status === 204) return null;
  return response.json();
}

// ────────────────────────────────────────────────────────────
//  SALAS
// ────────────────────────────────────────────────────────────

/**
 * Lista todas las salas de reunión del directorio
 * Requiere: Place.Read.All
 */
export async function getRooms(msalInstance, account) {
  const data = await callGraph(msalInstance, account, "/places/microsoft.graph.room");
  return data.value || [];
}

// ────────────────────────────────────────────────────────────
//  CALENDARIO / EVENTOS
// ────────────────────────────────────────────────────────────

/**
 * Obtiene los eventos de una sala en un rango de fechas
 * @param {string} roomEmail - ej: sala-a@empresa.com
 * @param {Date} start
 * @param {Date} end
 */
export async function getRoomEvents(msalInstance, account, roomEmail, start, end) {
  const startISO = start.toISOString();
  const endISO = end.toISOString();

  const data = await callGraph(
    msalInstance,
    account,
    `/users/${encodeURIComponent(roomEmail)}/calendarView` +
      `?startDateTime=${startISO}&endDateTime=${endISO}` +
      `&$select=id,subject,start,end,organizer,attendees,bodyPreview` +
      `&$orderby=start/dateTime` +
      `&$top=50`
  );
  return data.value || [];
}

/**
 * Obtiene eventos de múltiples salas en paralelo para el día indicado
 */
export async function getAllRoomsEvents(msalInstance, account, rooms, date) {
  const start = new Date(date);
  start.setHours(0, 0, 0, 0);
  const end = new Date(date);
  end.setHours(23, 59, 59, 999);

  const results = await Promise.allSettled(
    rooms.map((room) =>
      getRoomEvents(msalInstance, account, room.emailAddress, start, end).then(
        (events) => ({ roomId: room.id, roomEmail: room.emailAddress, events })
      )
    )
  );

  return results
    .filter((r) => r.status === "fulfilled")
    .map((r) => r.value);
}

/**
 * Verifica disponibilidad de una sala en un horario específico
 */
export async function checkAvailability(msalInstance, account, roomEmail, start, end) {
  const body = {
    schedules: [roomEmail],
    startTime: {
      dateTime: start.toISOString(),
      timeZone: "America/Mexico_City",
    },
    endTime: {
      dateTime: end.toISOString(),
      timeZone: "America/Mexico_City",
    },
    availabilityViewInterval: 30,
  };

  const data = await callGraph(
    msalInstance,
    account,
    "/me/calendar/getSchedule",
    { method: "POST", body: JSON.stringify(body) }
  );

  const schedule = data.value?.[0];
  if (!schedule) return { available: false };

  // availabilityView: "0"=libre, "1"=ocupado, "2"=tentativo
  const isFree = schedule.availabilityView
    ?.split("")
    .every((slot) => slot === "0");

  return {
    available: isFree,
    scheduleItems: schedule.scheduleItems || [],
  };
}

/**
 * Crea una nueva reserva en el buzón de la sala
 * La sala queda como asistente requerido (resource)
 * @param {object} booking - { subject, roomEmail, roomName, start, end, attendees[], organizer }
 */
export async function createBooking(msalInstance, account, booking) {
  const { subject, roomEmail, roomName, start, end, attendees = [] } = booking;

  const event = {
    subject,
    start: {
      dateTime: start.toISOString(),
      timeZone: "America/Mexico_City",
    },
    end: {
      dateTime: end.toISOString(),
      timeZone: "America/Mexico_City",
    },
    location: {
      displayName: roomName,
      locationEmailAddress: roomEmail,
    },
    attendees: [
      // La sala como recurso
      {
        emailAddress: { address: roomEmail, name: roomName },
        type: "resource",
      },
      // Asistentes adicionales
      ...attendees.map((email) => ({
        emailAddress: { address: email.trim() },
        type: "required",
      })),
    ],
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness",
  };

  // Se crea en el calendario del usuario que hace la reserva
  return callGraph(msalInstance, account, "/me/events", {
    method: "POST",
    body: JSON.stringify(event),
  });
}

/**
 * Cancela (elimina) un evento del calendario del usuario
 * Para cancelar también en la sala, se debe eliminar el evento original
 */
export async function cancelBooking(msalInstance, account, eventId) {
  return callGraph(msalInstance, account, `/me/events/${eventId}`, {
    method: "DELETE",
  });
}

/**
 * Actualiza un evento existente (cambio de hora, asunto, etc.)
 */
export async function updateBooking(msalInstance, account, eventId, changes) {
  return callGraph(msalInstance, account, `/me/events/${eventId}`, {
    method: "PATCH",
    body: JSON.stringify(changes),
  });
}

// ────────────────────────────────────────────────────────────
//  ESTADÍSTICAS
// ────────────────────────────────────────────────────────────

/**
 * Obtiene estadísticas de uso de una sala en el último mes
 * Devuelve: total reservas, horas usadas, horas disponibles
 */
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

  // Horario laboral: 8h × 22 días hábiles aprox
  const availableMinutes = 8 * 60 * 22;
  const occupancyPct = Math.round((totalMinutes / availableMinutes) * 100);

  return {
    totalBookings: events.length,
    totalHours: Math.round(totalMinutes / 60),
    occupancyPct: Math.min(occupancyPct, 100),
  };
}
