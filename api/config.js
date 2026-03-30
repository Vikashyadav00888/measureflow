import { isAnalyticsConfigured } from "../analytics.js";

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
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

  return res.status(200).json({
    ok: true,
    data: {
      adsenseClient: process.env.ADSENSE_CLIENT || process.env.VITE_ADSENSE_CLIENT || "",
      analyticsEnabled: isAnalyticsConfigured(),
    },
  });
}
