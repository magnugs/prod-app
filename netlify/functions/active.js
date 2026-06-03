import { requireAuth, json } from "./_lib/auth.js";
import { readActiveSessions, writeActiveSessions } from "./_lib/store.js";

// GET    /api/active                       → alle pågående sesjoner (alle brukere)
// POST   /api/active                       → start eller oppdater sesjon (overskriver (BO, linje))
// DELETE /api/active?bo=<bo>&line=<n>      → forkast sesjon
export default async (request) => {
  if (request.method === "GET") {
    const { error } = await requireAuth(request, ["user", "admin"]);
    if (error) return error;
    return json(await readActiveSessions());
  }

  if (request.method === "POST") {
    const { error } = await requireAuth(request, ["user", "admin"]);
    if (error) return error;
    let s;
    try { s = await request.json(); } catch { return json({ error: "Ugyldig JSON" }, 400); }
    if (!s || !s.batchOrderNumber || typeof s.line !== "number") {
      return json({ error: "batchOrderNumber og line er påkrevd" }, 400);
    }
    const list = await readActiveSessions();
    const others = list.filter(
      (x) => !(x.batchOrderNumber === s.batchOrderNumber && x.line === s.line)
    );
    const next = [...others, s];
    await writeActiveSessions(next);
    return json(s);
  }

  if (request.method === "DELETE") {
    // Brukere kan forkaste egne pågående sesjoner uten admin-passord
    const { error } = await requireAuth(request, ["user", "admin"]);
    if (error) return error;
    const url = new URL(request.url);
    const bo = url.searchParams.get("bo");
    const line = Number(url.searchParams.get("line"));
    if (!bo || isNaN(line)) {
      return json({ error: "bo og line query params er påkrevd" }, 400);
    }
    const list = await readActiveSessions();
    const next = list.filter(
      (x) => !(x.batchOrderNumber === bo && x.line === line)
    );
    await writeActiveSessions(next);
    return json({ ok: true });
  }

  return json({ error: "Method not allowed" }, 405);
};

export const config = { path: "/api/active" };
