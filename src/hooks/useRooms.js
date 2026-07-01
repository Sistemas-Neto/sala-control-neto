// ============================================================
//  HOOK: useRooms
//  Carga las salas del directorio y sus eventos.
//  Carga siempre el rango de la semana-mes visible (padding incluido)
//  para que las vistas Día, Semana y Mes usen el mismo set de eventos.
// ============================================================
import { useState, useEffect, useCallback } from "react";
import { useMsal } from "@azure/msal-react";
import { startOfWeek, endOfWeek, startOfMonth, endOfMonth } from "date-fns";
import { getRooms, getAllRoomsEvents } from "../services/graphService";

// Calcula el rango [start, end] que cubre el mes completo de `date`,
// incluyendo los días de relleno de la semana anterior/siguiente
// (los mismos que se ven pintados en la vista Mes / Semana).
function getVisibleRange(date) {
  const start = startOfWeek(startOfMonth(date), { weekStartsOn: 1 });
  const end = endOfWeek(endOfMonth(date), { weekStartsOn: 1 });
  return { start, end };
}

export function useRooms(selectedDate) {
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const [rooms, setRooms] = useState([]);
  const [events, setEvents] = useState({}); // { roomEmail: [events] }
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  const loadRooms = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const roomList = await getRooms(instance, account);
      setRooms(roomList);
      return roomList;
    } catch (err) {
      setError(err.message);
      return [];
    }
  }, [instance, account]);

  const loadEvents = useCallback(
    async (roomList, date) => {
      try {
        const { start, end } = getVisibleRange(date);
        const allEvents = await getAllRoomsEvents(instance, account, roomList, start, end);
        const eventsMap = {};
        allEvents.forEach(({ roomEmail, events: evs }) => {
          eventsMap[roomEmail] = evs;
        });
        setEvents(eventsMap);
      } catch (err) {
        console.error("Error cargando eventos:", err);
      } finally {
        setLoading(false);
      }
    },
    [instance, account]
  );

  useEffect(() => {
    loadRooms().then((roomList) => {
      if (roomList.length > 0) loadEvents(roomList, selectedDate);
      else setLoading(false);
    });
  }, [loadRooms, loadEvents, selectedDate]);

  const refresh = useCallback(() => {
    loadRooms().then((roomList) => loadEvents(roomList, selectedDate));
  }, [loadRooms, loadEvents, selectedDate]);

  return { rooms, events, loading, error, refresh };
}
