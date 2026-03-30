import http from "node:http";
import { isAnalyticsConfigured, loadAdminSummary, recordAnalyticsEvent, verifyAdminPassword, verifyAdminToken } from "./analytics.js";

const PORT = process.env.MEASUREFLOW_API_PORT || 8787;
const RETRYABLE = new Set([
  "overloaded_error",
  "rate_limit_error",
  "api_error",
  "server_error",
  "resource_exhausted",
  "quota_exceeded",
]);

function sendJson(res, status, body) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  });
  res.end(JSON.stringify(body));
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    let raw = "";
    req.on("data", (chunk) => {
      raw += chunk;
      if (raw.length > 20 * 1024 * 1024) {
        reject(new Error("Request too large"));
        req.destroy();
      }
    });
    req.on("end", () => {
      try {
        resolve(raw ? JSON.parse(raw) : {});
      } catch {
        reject(new Error("Invalid JSON body"));
      }
    });
    req.on("error", reject);
  });
}

function toDataUrl(mediaType, imageData) {
  return `data:${mediaType || "image/jpeg"};base64,${imageData}`;
}

function createApiError(message, type, status) {
  const err = new Error(message);
  err.type = type;
  err.status = status;
  return err;
}

function normalizeProviderError(err, fallbackType = "api_error") {
  const message = err?.message || "API request failed";
  const type = err?.type || fallbackType;
  const status = err?.status || 500;
  return createApiError(message, type, status);
}

async function postGemini({ apiKey, system, imageData, mediaType, promptText, model }) {
  const parts = [
    {
      text:
        promptText ||
        "Extract all measurements. GOLDEN RULE: if the second value has feet or inch markers it is sqft; if it is a plain number it is rnft. Return JSON array only.",
    },
  ];
  if (imageData) {
    parts.push({
      inline_data: {
        mime_type: mediaType || "image/jpeg",
        data: imageData,
      },
    });
  }
  const resp = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model || "gemini-2.5-flash"}:generateContent`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-goog-api-key": apiKey,
    },
    body: JSON.stringify({
      system_instruction: {
        parts: [{ text: system || "" }],
      },
      contents: [
        {
          role: "user",
          parts,
        },
      ],
      generationConfig: {
        responseMimeType: "application/json",
      },
    }),
  });

  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const mappedType = data?.error?.status === "RESOURCE_EXHAUSTED" ? "resource_exhausted" : (data?.error?.type || "").toLowerCase();
    throw createApiError(data?.error?.message || "Gemini request failed", mappedType || "api_error", resp.status);
  }
  const text =
    data?.candidates?.[0]?.content?.parts
      ?.map((part) => part?.text || "")
      .join("") || "";
  return {
    content: [{ text }],
    raw: data,
  };
}

async function postXai({ apiKey, system, imageData, mediaType, promptText, model }) {
  const content = [
    {
      type: "text",
      text:
        promptText ||
        "Extract all measurements. GOLDEN RULE: if the second value has feet or inch markers it is sqft; if it is a plain number it is rnft. Return JSON array only.",
    },
  ];
  if (imageData) {
    content.push({
      type: "image_url",
      image_url: {
        url: toDataUrl(mediaType, imageData),
        detail: "high",
      },
    });
  }
  const resp = await fetch("https://api.x.ai/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: model || "grok-4",
      temperature: 0,
      response_format: { type: "json_object" },
      messages: [
        { role: "system", content: system || "" },
        {
          role: "user",
          content,
        },
      ],
    }),
  });

  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const type = (data?.error?.type || "").toLowerCase();
    throw createApiError(data?.error?.message || "xAI request failed", type || "api_error", resp.status);
  }
  const text = data?.choices?.[0]?.message?.content || "";
  return {
    content: [{ text }],
    raw: data,
  };
}

function getProviderPlan(body) {
  const plan = [];
  const preferred = String(process.env.AI_PROVIDER_ORDER || "").split(",").map((x) => x.trim().toLowerCase()).filter(Boolean);
  const entries = {
    gemini1: {
      name: "Gemini Primary",
      kind: "gemini",
      apiKey: process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY || body.apiKey || "",
      model: process.env.GEMINI_MODEL || "gemini-2.5-flash",
    },
    gemini2: {
      name: "Gemini Backup",
      kind: "gemini",
      apiKey: process.env.GEMINI_API_KEY_2 || process.env.GOOGLE_API_KEY_2 || "",
      model: process.env.GEMINI_MODEL_2 || process.env.GEMINI_MODEL || "gemini-2.5-flash",
    },
    grok: {
      name: "Grok Backup",
      kind: "xai",
      apiKey: process.env.XAI_API_KEY || "",
      model: process.env.XAI_MODEL || "grok-4",
    },
  };

  const order = preferred.length ? preferred : ["gemini1", "gemini2", "grok"];
  order.forEach((key) => {
    const entry = entries[key];
    if (entry?.apiKey) plan.push(entry);
  });
  return plan;
}

async function callProvider(provider, body) {
  if (provider.kind === "xai") {
    return postXai({
      apiKey: provider.apiKey,
      system: body.system,
      imageData: body.imageData,
      mediaType: body.mediaType,
      promptText: body.promptText,
      model: provider.model,
    });
  }
  return postGemini({
    apiKey: provider.apiKey,
    system: body.system,
    imageData: body.imageData,
    mediaType: body.mediaType,
    promptText: body.promptText,
    model: provider.model,
  });
}

async function extractWithFallback(body) {
  const providers = getProviderPlan(body);
  if (!providers.length) {
    throw createApiError(
      "Missing AI provider keys. Set GEMINI_API_KEY and optionally GEMINI_API_KEY_2 or XAI_API_KEY.",
      "authentication_error",
      400
    );
  }
  if (!body.imageData) {
    throw createApiError("Missing image data.", "invalid_request_error", 400);
  }

  const failures = [];
  for (const provider of providers) {
    try {
      const data = await callProvider(provider, body);
      return { data, provider: provider.name, failures };
    } catch (err) {
      const normalized = normalizeProviderError(err);
      failures.push(`${provider.name}: ${normalized.message}`);
      if (!RETRYABLE.has(normalized.type)) {
        throw normalized;
      }
    }
  }

  const last = failures[failures.length - 1];
  throw createApiError(last || "All providers failed", "api_error", 503);
}

const server = http.createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    return sendJson(res, 204, {});
  }

  if (req.method === "GET" && req.url === "/api/config") {
    return sendJson(res, 200, {
      ok: true,
      data: {
        adsenseClient: process.env.ADSENSE_CLIENT || process.env.VITE_ADSENSE_CLIENT || "",
        analyticsEnabled: isAnalyticsConfigured(),
      },
    });
  }

  if (req.method === "POST" && req.url === "/api/track") {
    try {
      const body = await readJson(req);
      const result = await recordAnalyticsEvent({
        userId: body.userId,
        sessionId: body.sessionId,
        eventType: body.eventType,
        route: body.route,
        tab: body.tab,
        details: body.details,
        userAgent: req.headers["user-agent"] || "",
        referrer: req.headers.referer || "",
      });
      return sendJson(res, 200, { ok: true, data: result });
    } catch (err) {
      return sendJson(res, err?.status || 500, {
        ok: false,
        error: { message: err.message || "Server error", type: err.type || "server_error" },
      });
    }
  }

  if (req.method === "POST" && req.url === "/api/admin-login") {
    try {
      const body = await readJson(req);
      const token = verifyAdminPassword(body?.password);
      return sendJson(res, 200, { ok: true, data: { token } });
    } catch (err) {
      return sendJson(res, err?.status || 500, {
        ok: false,
        error: { message: err.message || "Server error", type: err.type || "server_error" },
      });
    }
  }

  if (req.method === "GET" && req.url === "/api/admin-summary") {
    try {
      const auth = String(req.headers.authorization || "");
      const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
      verifyAdminToken(token);
      const summary = await loadAdminSummary();
      return sendJson(res, 200, { ok: true, data: summary });
    } catch (err) {
      return sendJson(res, err?.status || 500, {
        ok: false,
        error: { message: err.message || "Server error", type: err.type || "server_error" },
      });
    }
  }

  if (req.method === "POST" && req.url === "/api/extract") {
    try {
      const body = await readJson(req);
      const result = await extractWithFallback(body);
      return sendJson(res, 200, {
        ok: true,
        data: result.data,
        meta: {
          provider: result.provider,
          fallbackCount: result.failures.length,
        },
      });
    } catch (err) {
      return sendJson(res, err?.status || 500, {
        ok: false,
        error: { message: err.message || "Server error", type: err.type || "server_error" },
      });
    }
  }

  sendJson(res, 404, { ok: false, error: { message: "Not found", type: "not_found" } });
});

server.listen(PORT, () => {
  console.log(`MeasureFlow API listening on http://localhost:${PORT}`);
});
