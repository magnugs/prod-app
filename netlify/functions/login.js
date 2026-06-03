import { sha256Hex, safeEqualHex, signToken, json } from "./_lib/auth.js";

// POST /api/login   body: { password, asAdmin?: boolean }
// Returnerer { token, role, expiresIn }
export default async (request) => {
  if (request.method !== "POST") return json({ error: "Method not allowed" }, 405);

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ error: "Ugyldig JSON" }, 400);
  }
  const { password, asAdmin } = body || {};
  if (typeof password !== "string" || password.length === 0) {
    return json({ error: "Passord mangler" }, 400);
  }

  const expectedHash = asAdmin
    ? process.env.ADMIN_PASSWORD_HASH
    : process.env.APP_PASSWORD_HASH;
  if (!expectedHash) {
    return json({ error: "Server ikke konfigurert" }, 500);
  }

  const givenHash = sha256Hex(password);
  if (!safeEqualHex(givenHash, expectedHash)) {
    // Liten forsinkelse for å bremse brute-force
    await new Promise((r) => setTimeout(r, 400));
    return json({ error: "Feil passord" }, 401);
  }

  const role = asAdmin ? "admin" : "user";
  const token = await signToken({ role });
  return json({ token, role, expiresIn: "30d" });
};

export const config = { path: "/api/login" };
