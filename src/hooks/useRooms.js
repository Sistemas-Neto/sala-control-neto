// ============================================================
//  HOOK: useRooms
//  Carga las salas del directorio y sus eventos del día
// ============================================================

import { useState, useEffect, useCallback } from "react";
import { useMsal } from "@azure/msal-react";
import { getRooms, getAllRoomsEvents } from "../services/graphService";

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
        const allEvents = await getAllRoomsEvents(instance, account, roomList, date);
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
