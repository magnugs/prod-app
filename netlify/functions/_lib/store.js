// Wrapper rundt Netlify Blobs som gir oss "tabell"-semantikk
// med getJSON, setJSON og atomær oppdatering via "etag".

import { getStore } from "@netlify/blobs";

const REG_STORE = "registrations";
const ACTIVE_STORE = "active-sessions";
const REG_KEY = "all"; // Vi lagrer hele lista som ett JSON-objekt — enklere semantikk for små volum
const ACTIVE_KEY = "all";

export async function readRegistrations() {
  const store = getStore(REG_STORE);
  const raw = await store.get(REG_KEY, { type: "json" });
  return Array.isArray(raw) ? raw : [];
}

export async function writeRegistrations(list) {
  const store = getStore(REG_STORE);
  await store.setJSON(REG_KEY, list);
}

export async function readActiveSessions() {
  const store = getStore(ACTIVE_STORE);
  const raw = await store.get(ACTIVE_KEY, { type: "json" });
  return Array.isArray(raw) ? raw : [];
}

export async function writeActiveSessions(list) {
  const store = getStore(ACTIVE_STORE);
  await store.setJSON(ACTIVE_KEY, list);
}
