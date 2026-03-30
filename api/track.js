import { recordAnalyticsEvent } from "../analytics.js";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({
      ok: false,
      error: { message: "Method not allowed", type: "method_not_allowed" },
    });
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
    const data = await recordAnalyticsEvent({
      userId: body.userId,
      sessionId: body.sessionId,
      eventType: body.eventType,
      route: body.route,
      tab: body.tab,
      details: body.details,
      userAgent: req.headers["user-agent"] || "",
      referrer: req.headers.referer || "",
    });
    return res.status(200).json({ ok: true, data });
  } catch (err) {
    return res.status(err?.status || 500).json({
      ok: false,
      error: { message: err.message || "Server error", type: err.type || "server_error" },
    });
  }
}
