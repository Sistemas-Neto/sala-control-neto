// ── SOLICITUDES DE RESERVA — SharePoint ──────────────────────
import { graphScopes } from "../authConfig";

const SITE_ID = "soyneto.sharepoint.com,5c0b849c-bdff-4d81-baae-63a1791481a8,5ceb8194-2a24-4b35-8f12-5b4a43c5b5e5";
const LIST_NAME = "SolicitudesReserva";

async function getToken(msalInstance, account) {
  try {
    const r = await msalInstance.acquireTokenSilent({ ...graphScopes, account });
    return r.accessToken;
  } catch {
    await msalInstance.acquireTokenRedirect(graphScopes);
  }
}

export async function getSolicitudes(msalInstance, account) {
  const token = await getToken(msalInstance, account);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_NAME}/items?expand=fields&$top=100&$orderby=createdDateTime desc`,
    { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
  );
  const data = await res.json();
  return (data.value || []).map(i => ({ id: i.id, ...i.fields }));
}

/**
 * Actualiza uno o más campos de una solicitud a la vez.
 * Ejemplo: actualizarCampos(instance, account, itemId, { Estado: "Aprobado", IdEvento: "AAMk...", OrganizadorEmail: "x@y.com" })
 */
export async function actualizarCampos(msalInstance, account, itemId, campos) {
  const token = await getToken(msalInstance, account);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_NAME}/items/${itemId}/fields`,
    {
      method: "PATCH",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(campos)
    }
  );
  if (!res.ok) throw new Error("Error al actualizar la solicitud");
  return true;
}

// Se mantiene por compatibilidad con código existente que solo actualiza el estado.
export async function actualizarEstado(msalInstance, account, itemId, estado) {
  return actualizarCampos(msalInstance, account, itemId, { Estado: estado });
}
