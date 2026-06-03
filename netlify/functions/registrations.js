import { requireAuth, json } from "./_lib/auth.js";
import { readRegistrations, writeRegistrations } from "./_lib/store.js";

// GET    /api/registrations          → hele lista
// POST   /api/registrations          → ny eller oppdater (basert på id)
// DELETE /api/registrations/:id      → slett (krever admin-rolle)
export default async (request) => {
  const url = new URL(request.url);
  // Path kan være /api/registrations eller /api/registrations/<id>
  const parts = url.pathname.split("/").filter(Boolean);
  const id = parts.length >= 3 ? parts[2] : null;

  if (request.method === "GET") {
    const { error, payload } = await requireAuth(request, ["user", "admin"]);
    if (error) return error;
    const list = await readRegistrations();
    return json(list);
  }

  if (request.method === "POST") {
    const { error } = await requireAuth(request, ["user", "admin"]);
    if (error) return error;
    let reg;
    try { reg = await request.json(); } catch { return json({ error: "Ugyldig JSON" }, 400); }
    if (!reg || !reg.id || !reg.batchOrderNumber) {
      return json({ error: "id og batchOrderNumber er påkrevd" }, 400);
    }
    const list = await readRegistrations();
    const idx = list.findIndex((r) => r.id === reg.id);
    const updated = { ...reg, updatedAt: new Date().toISOString() };
    let next;
    if (idx >= 0) {
      next = list.map((r, i) => (i === idx ? updated : r));
    } else {
      next = [...list, { ...updated, createdAt: updated.createdAt || new Date().toISOString() }];
    }
    await writeRegistrations(next);
    return json(updated);
  }

  if (request.method === "DELETE") {
    // Sletting krever admin-rolle
    const { error } = await requireAuth(request, ["admin"]);
    if (error) return error;
    if (!id) return json({ error: "Mangler id" }, 400);
    const list = await readRegistrations();
    const next = list.filter((r) => r.id !== id);
    if (next.length === list.length) return json({ error: "Ikke funnet" }, 404);
    await writeRegistrations(next);
    return json({ ok: true, id });
  }

  return json({ error: "Method not allowed" }, 405);
};

export const config = { path: ["/api/registrations", "/api/registrations/*"] };
