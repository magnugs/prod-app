// Felles auth- og response-helpere for Netlify Functions.
//
// Sikkerhetsmodell:
//   1) Bruker logger seg inn med APP_PASSWORD og får en JWT (rolle "user")
//   2) Ved sletting kreves ADMIN_PASSWORD som gir en ny JWT med rolle "admin"
//   3) Alle skrive-/leseendepunkter krever gyldig JWT i Authorization-header
//   4) JWT signeres med JWT_SECRET (HS256), utløper etter 30 dager
//
// Miljøvariabler som MÅ være satt i Netlify-prosjektet:
//   APP_PASSWORD_HASH    sha256-hash av brukerpassordet (hex)
//   ADMIN_PASSWORD_HASH  sha256-hash av admin-passordet (hex)
//   JWT_SECRET           lang tilfeldig streng, minst 32 tegn

import { SignJWT, jwtVerify } from "jose";
import { createHash, timingSafeEqual } from "node:crypto";

const ALG = "HS256";
const TOKEN_TTL = "30d";

function getSecretKey() {
  const secret = process.env.JWT_SECRET;
  if (!secret || secret.length < 32) {
    throw new Error("JWT_SECRET miljøvariabel mangler eller for kort (minst 32 tegn)");
  }
  return new TextEncoder().encode(secret);
}

export function sha256Hex(text) {
  return createHash("sha256").update(text, "utf8").digest("hex");
}

// Sammenlign hash-strenger i konstant tid (motvirker timing-angrep)
export function safeEqualHex(aHex, bHex) {
  if (typeof aHex !== "string" || typeof bHex !== "string") return false;
  if (aHex.length !== bHex.length) return false;
  try {
    return timingSafeEqual(Buffer.from(aHex, "hex"), Buffer.from(bHex, "hex"));
  } catch {
    return false;
  }
}

export async function signToken(payload) {
  return await new SignJWT(payload)
    .setProtectedHeader({ alg: ALG })
    .setIssuedAt()
    .setExpirationTime(TOKEN_TTL)
    .sign(getSecretKey());
}

// Returnerer { role: "user"|"admin", ... } eller null hvis ugyldig/utløpt
export async function verifyToken(token) {
  if (!token) return null;
  try {
    const { payload } = await jwtVerify(token, getSecretKey(), { algorithms: [ALG] });
    return payload;
  } catch {
    return null;
  }
}

// Plukker JWT fra Authorization-header
export function extractToken(request) {
  const auth = request.headers.get("authorization") || request.headers.get("Authorization");
  if (!auth) return null;
  const match = auth.match(/^Bearer\s+(.+)$/i);
  return match ? match[1] : null;
}

// Krever gyldig JWT med en av de angitte rollene. Returnerer payload eller en Response (401/403).
export async function requireAuth(request, allowedRoles = ["user", "admin"]) {
  const token = extractToken(request);
  const payload = await verifyToken(token);
  if (!payload) {
    return { error: json({ error: "Mangler eller ugyldig token" }, 401) };
  }
  if (!allowedRoles.includes(payload.role)) {
    return { error: json({ error: "Mangler nødvendig rolle" }, 403) };
  }
  return { payload };
}

export function json(body, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store",
      ...extraHeaders,
    },
  });
}
