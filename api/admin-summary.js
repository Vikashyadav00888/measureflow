import { loadAdminSummary, verifyAdminToken } from "../analytics.js";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "GET") {
    return res.status(405).json({
      ok: false,
      error: { message: "Method not allowed", type: "method_not_allowed" },
    });
  }

  try {
    const auth = String(req.headers.authorization || "");
    const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
    verifyAdminToken(token);
    const data = await loadAdminSummary();
    return res.status(200).json({ ok: true, data });
  } catch (err) {
    return res.status(err?.status || 500).json({
      ok: false,
      error: { message: err.message || "Server error", type: err.type || "server_error" },
    });
  }
}
