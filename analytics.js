import crypto from "node:crypto";

const ADMIN_TTL_MS = 1000 * 60 * 60 * 24 * 7;

function env(name) {
  return process.env[name] || "";
}

export function getAnalyticsConfig() {
  return {
    supabaseUrl: env("SUPABASE_URL"),
    serviceRoleKey: env("SUPABASE_SERVICE_ROLE_KEY"),
    adminPassword: env("ADMIN_PASSWORD"),
    adminSecret: env("ADMIN_SECRET") || env("ADMIN_PASSWORD"),
  };
}

export function isAnalyticsConfigured() {
  const cfg = getAnalyticsConfig();
  return !!(cfg.supabaseUrl && cfg.serviceRoleKey);
}

function createError(message, status = 500, type = "server_error") {
  const err = new Error(message);
  err.status = status;
  err.type = type;
  return err;
}

async function supabaseFetch(path, init = {}) {
  const cfg = getAnalyticsConfig();
  if (!cfg.supabaseUrl || !cfg.serviceRoleKey) {
    throw createError("Analytics database is not configured.", 400, "analytics_not_configured");
  }
  const mergedHeaders = {
    "Content-Type": "application/json",
    apikey: cfg.serviceRoleKey,
    Authorization: `Bearer ${cfg.serviceRoleKey}`,
    Prefer: "return=representation",
    ...(init.headers || {}),
  };
  Object.keys(mergedHeaders).forEach((key) => {
    if (mergedHeaders[key] === undefined) delete mergedHeaders[key];
  });
  const resp = await fetch(`${cfg.supabaseUrl}/rest/v1/${path}`, {
    ...init,
    headers: mergedHeaders,
  });
  const payload = await resp.json().catch(() => null);
  if (!resp.ok) {
    throw createError(payload?.message || payload?.error || "Analytics request failed", resp.status, "analytics_error");
  }
  return { payload, headers: resp.headers };
}

export async function recordAnalyticsEvent(event) {
  if (!isAnalyticsConfigured()) {
    return { stored: false };
  }
  const body = {
    user_id: String(event.userId || "").slice(0, 120),
    session_id: String(event.sessionId || "").slice(0, 120),
    event_type: String(event.eventType || "unknown").slice(0, 80),
    route: String(event.route || "").slice(0, 120),
    tab: String(event.tab || "").slice(0, 80),
    details: event.details && typeof event.details === "object" ? event.details : {},
    user_agent: String(event.userAgent || "").slice(0, 500),
    referrer: String(event.referrer || "").slice(0, 500),
  };
  await supabaseFetch("mf_events", {
    method: "POST",
    body: JSON.stringify(body),
  });
  return { stored: true };
}

export function createAdminToken() {
  const cfg = getAnalyticsConfig();
  if (!cfg.adminSecret) throw createError("Admin password is not configured.", 500, "admin_not_configured");
  const payload = {
    iat: Date.now(),
    exp: Date.now() + ADMIN_TTL_MS,
    role: "admin",
  };
  const encoded = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const sig = crypto.createHmac("sha256", cfg.adminSecret).update(encoded).digest("base64url");
  return `${encoded}.${sig}`;
}

export function verifyAdminPassword(password) {
  const cfg = getAnalyticsConfig();
  if (!cfg.adminPassword) throw createError("Admin password is not configured.", 500, "admin_not_configured");
  if (String(password || "") !== cfg.adminPassword) {
    throw createError("Invalid admin password.", 401, "authentication_error");
  }
  return createAdminToken();
}

export function verifyAdminToken(token) {
  const cfg = getAnalyticsConfig();
  if (!cfg.adminSecret) throw createError("Admin password is not configured.", 500, "admin_not_configured");
  const raw = String(token || "");
  const [encoded, sig] = raw.split(".");
  if (!encoded || !sig) throw createError("Missing admin token.", 401, "authentication_error");
  const expected = crypto.createHmac("sha256", cfg.adminSecret).update(encoded).digest("base64url");
  if (sig !== expected) throw createError("Invalid admin token.", 401, "authentication_error");
  const payload = JSON.parse(Buffer.from(encoded, "base64url").toString("utf8"));
  if (!payload?.exp || payload.exp < Date.now()) throw createError("Admin session expired.", 401, "authentication_error");
  return payload;
}

async function fetchEvents(limit = 1000) {
  const { payload } = await supabaseFetch(`mf_events?select=id,user_id,session_id,event_type,route,tab,details,created_at,user_agent,referrer&order=created_at.desc&limit=${limit}`, {
    method: "GET",
    headers: {
      Prefer: undefined,
    },
  });
  return Array.isArray(payload) ? payload : [];
}

function uniqueCount(items) {
  return new Set(items.filter(Boolean)).size;
}

export async function loadAdminSummary() {
  if (!isAnalyticsConfigured()) {
    return {
      configured: false,
      totalUsers: 0,
      activeUsers7d: 0,
      inactiveUsers: 0,
      uploads: 0,
      downloads: 0,
      visits24h: 0,
      recent: [],
    };
  }

  const events = await fetchEvents(1500);
  const now = Date.now();
  const d1 = now - 24 * 60 * 60 * 1000;
  const d7 = now - 7 * 24 * 60 * 60 * 1000;
  const latestByUser = new Map();

  events.forEach((ev) => {
    const ts = new Date(ev.created_at).getTime();
    if (ev.user_id) {
      const prev = latestByUser.get(ev.user_id);
      if (!prev || ts > prev) latestByUser.set(ev.user_id, ts);
    }
  });

  const totalUsers = latestByUser.size;
  const activeUsers7d = Array.from(latestByUser.values()).filter((ts) => ts >= d7).length;
  const inactiveUsers = Math.max(0, totalUsers - activeUsers7d);
  const uploads = events.filter((ev) => ev.event_type === "upload_complete").length;
  const downloads = events.filter((ev) => ev.event_type === "download").length;
  const visits24h = uniqueCount(events.filter((ev) => ev.event_type === "app_open" && new Date(ev.created_at).getTime() >= d1).map((ev) => ev.user_id));
  const recent = events.slice(0, 30).map((ev) => ({
    id: ev.id,
    when: ev.created_at,
    type: ev.event_type,
    userId: ev.user_id,
    sessionId: ev.session_id,
    tab: ev.tab,
    route: ev.route,
    details: ev.details || {},
  }));

  return {
    configured: true,
    totalUsers,
    activeUsers7d,
    inactiveUsers,
    uploads,
    downloads,
    visits24h,
    recent,
  };
}
