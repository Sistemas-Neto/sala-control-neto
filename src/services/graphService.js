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
      { id: "1", displayName: "Sala Entusiasmo", emailAddress: "salaentusiasmo@soyneto.onmicrosoft.com", capacity: 35, building: "Campus principal" },
      { id: "2", displayName: "Sala Practicidad", emailAddress: "salapracticidad@soyneto.onmicrosoft.com", capacity: 20, building: "Campus principal" },
      { id: "3", displayName: "Sala Tenacidad", emailAddress: "salatenacidad@soyneto.onmicrosoft.com", capacity: 35, building: "Campus principal" },
    ];
  }
}

// ── EVENTOS con ID real usando /me/events ──────────────────
// FIX: Se usa toLocalISOString() en vez de toISOString() para evitar que
// Graph interprete las horas en UTC. El header "Prefer" hace que Graph
// devuelva los eventos ya convertidos a hora de Ciudad de México.
export async function getRoomEvents(msalInstance, account, roomEmail, start, end) {
  try {
    const startStr = toLocalISOString(start); // FIX: era start.toISOString()
    const endStr = toLocalISOString(end);     // FIX: era end.toISOString()

    const data = await callGraph(
      msalInstance,
      account,
      `/me/calendarView?startDateTime=${startStr}&endDateTime=${endStr}&$select=id,subject,start,end,organizer,location,attendees&$top=50&$orderby=start/dateTime`,
      {
        headers: {
          "Prefer": 'outlook.timezone="America/Mexico_City"', // FIX: Graph devuelve horas en CDMX
        },
      }
    );

    const events = data.value || [];

    return events.filter(ev => {
      const loc = ev.location?.locationEmailAddress?.toLowerCase() || "";
      const attendees = ev.attendees || [];
      const hasRoom = attendees.some(a =>
        a.emailAddress?.address?.toLowerCase() === roomEmail.toLowerCase()
      );
      return loc === roomEmail.toLowerCase() || hasRoom;
    });
  } catch {
    return [];
  }
}

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

// Cancelar con ID real del evento
export async function cancelBooking(msalInstance, account, eventId) {
  return callGraph(msalInstance, account, `/me/events/${eventId}`, {
    method: "DELETE",
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
