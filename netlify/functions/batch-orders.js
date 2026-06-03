import { requireAuth, json } from "./_lib/auth.js";
import { getStore } from "@netlify/blobs";
import { timingSafeEqual } from "node:crypto";

const STORE = "batch-orders";
const KEY = "all";
const META_KEY = "meta";

// GET  /api/batch-orders          → liste batch orders + metadata
// POST /api/batch-orders          → erstatt hele lista (krever x-api-key)
//
// GET krever JWT (samme som resten av appen).
// POST krever INGEST_TOKEN i x-api-key-header — egnet for Power Automate.
export default async (request) => {
  const store = getStore(STORE);

  if (request.method === "GET") {
    const { error } = await requireAuth(request, ["user", "admin"]);
    if (error) return error;
    const list = await store.get(KEY, { type: "json" });
    const meta = await store.get(META_KEY, { type: "json" });
    return json({
      orders: Array.isArray(list) ? list : [],
      updatedAt: meta?.updatedAt || null,
      count: Array.isArray(list) ? list.length : 0,
    });
  }

  if (request.method === "POST") {
    // API-key-auth (ikke JWT) — Power Automate sender x-api-key
    const expected = process.env.INGEST_TOKEN;
    if (!expected || expected.length < 16) {
      return json({ error: "INGEST_TOKEN ikke konfigurert på server" }, 500);
    }
    const given = request.headers.get("x-api-key") || "";
    if (!safeEqualStr(given, expected)) {
      await new Promise((r) => setTimeout(r, 400));
      return json({ error: "Ugyldig API-key" }, 401);
    }

    let body;
    try { body = await request.json(); } catch { return json({ error: "Ugyldig JSON" }, 400); }

    // Aksepterer flere formater:
    //   - Array direkte
    //   - { value: [...] } (D365 OData-format)
    //   - { orders: [...] }
    let raw;
    if (Array.isArray(body)) raw = body;
    else if (Array.isArray(body?.value)) raw = body.value;
    else if (Array.isArray(body?.orders)) raw = body.orders;
    else return json({ error: "Forventet array, { value: [...] } eller { orders: [...] }" }, 400);

    if (raw.length === 0) {
      return json({ error: "Tom liste — avbryter for å unngå utilsiktet sletting" }, 400);
    }

    // Valider at det faktisk ligner batch orders (sjekker et par nøkkelfelter)
    const sample = raw[0];
    if (!sample || typeof sample !== "object" || !sample.BatchOrderNumber) {
      return json({ error: "Forventet objekter med BatchOrderNumber-felt" }, 400);
    }

    await store.setJSON(KEY, raw);
    const meta = {
      updatedAt: new Date().toISOString(),
      count: raw.length,
      source: request.headers.get("x-source") || "api",
    };
    await store.setJSON(META_KEY, meta);
    return json({ ok: true, ...meta });
  }

  return json({ error: "Method not allowed" }, 405);
};

function safeEqualStr(a, b) {
  if (typeof a !== "string" || typeof b !== "string") return false;
  if (a.length !== b.length) return false;
  try {
    return timingSafeEqual(Buffer.from(a, "utf8"), Buffer.from(b, "utf8"));
  } catch {
    return false;
  }
}

export const config = { path: "/api/batch-orders" };
