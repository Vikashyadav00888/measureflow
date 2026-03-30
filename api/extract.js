const RETRYABLE = new Set(["overloaded_error", "rate_limit_error", "api_error", "server_error", "resource_exhausted", "quota_exceeded"]);

function toDataUrl(mediaType, imageData) {
  return `data:${mediaType || "image/jpeg"};base64,${imageData}`;
}

function createApiError(message, type, status) {
  const err = new Error(message);
  err.type = type;
  err.status = status;
  return err;
}

async function postGemini({ apiKey, system, imageData, mediaType, promptText, model }) {
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
          parts: [
            {
              text:
                promptText ||
                "Extract all measurements. GOLDEN RULE: if the second value has feet or inch markers it is sqft; if it is a plain number it is rnft. Return JSON array only.",
            },
            {
              inline_data: {
                mime_type: mediaType || "image/jpeg",
                data: imageData,
              },
            },
          ],
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
          content: [
            {
              type: "text",
              text:
                promptText ||
                "Extract all measurements. GOLDEN RULE: if the second value has feet or inch markers it is sqft; if it is a plain number it is rnft. Return JSON array only.",
            },
            {
              type: "image_url",
              image_url: {
                url: toDataUrl(mediaType, imageData),
                detail: "high",
              },
            },
          ],
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
  return order.map((key) => entries[key]).filter((entry) => entry?.apiKey);
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
    const providers = getProviderPlan(body);
    if (!providers.length) {
      return res.status(400).json({
        ok: false,
        error: {
          message: "Missing AI provider keys. Set GEMINI_API_KEY and optionally GEMINI_API_KEY_2 or XAI_API_KEY.",
          type: "authentication_error",
        },
      });
    }

    if (!body.imageData) {
      return res.status(400).json({
        ok: false,
        error: { message: "Missing image data.", type: "invalid_request_error" },
      });
    }

    const failures = [];
    let lastErr = null;
    for (const provider of providers) {
      try {
        const data = await callProvider(provider, body);
        return res.status(200).json({
          ok: true,
          data,
          meta: { provider: provider.name, fallbackCount: failures.length },
        });
      } catch (err) {
        lastErr = err;
        failures.push(`${provider.name}: ${err.message}`);
        if (!RETRYABLE.has(err.type)) break;
      }
    }

    return res.status(lastErr?.status || 500).json({
      ok: false,
      error: { message: failures[failures.length - 1] || lastErr?.message || "All providers failed", type: lastErr?.type || "api_error" },
    });
  } catch (err) {
    return res.status(500).json({
      ok: false,
      error: { message: err.message || "Server error", type: "server_error" },
    });
  }
}
