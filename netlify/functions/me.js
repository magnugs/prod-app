import { requireAuth, json } from "./_lib/auth.js";

// GET /api/me → { role } hvis gyldig token, ellers 401
export default async (request) => {
  const { error, payload } = await requireAuth(request, ["user", "admin"]);
  if (error) return error;
  return json({ role: payload.role });
};

export const config = { path: "/api/me" };
