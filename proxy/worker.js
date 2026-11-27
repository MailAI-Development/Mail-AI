/**
 * Mail AI extraction proxy (Cloudflare Worker)
 * --------------------------------------------
 * Holds the OpenAI key server-side so it never ships inside the desktop app.
 *
 *   app  ->  https://mailai.uk/v1/extract  (Netlify forwards /v1/*, see website/_redirects)
 *        ->  this Worker  ->  OpenAI chat completions
 *
 * v1 (desktop app v1.6 and earlier; frozen, do not change its behaviour):
 *   POST /v1/extract   header X-App-Token: <APP_TOKEN baked into the app>
 *   body {"messages": [{role: "system"|"user", content: "..."}, ...]}   (max 2)
 *   -> {"content": "<model output>"}
 *
 * v2 (web app): the prompt, schema and cleanup live server-side in extract.js.
 *   POST /v2/extract   header Authorization: Bearer <Microsoft ID token>   (web app)
 *                      or X-App-Token                                   (eval scripts)
 *   body {"subject": "...", "body": "<plain-text email body>"}
 *   -> {"vessels": [{"name", "dwt_k", "built", "open_location", "open_date"}]}
 *
 * Secrets (Cloudflare dashboard → Worker → Settings → Variables, type "Secret",
 * or `npx wrangler secret put NAME`):
 *   OPENAI_API_KEY      the OpenAI key
 *   MAILAI_APP_SECRET   must equal APP_TOKEN in app/_build_secrets.py
 * Plain vars (optional, defaults in wrangler.toml):
 *   OPENAI_MODEL        v1 model, default gpt-5.4-nano
 *   OPENAI_MODEL_V2     v2 model, default gpt-5.4-mini
 *   MS_CLIENT_ID        Microsoft Entra app (client) ID of the web app; v2 accepts its ID tokens
 * Secrets for the web app's accounts (see account.js):
 *   LICENSE_SECRET            same as the desktop app's; validates MAILAI-… keys entered on the web
 *   KOFI_VERIFICATION_TOKEN   from Ko-fi → Settings → API; authenticates POST /kofi/webhook
 *   RESEND_API_KEY            Resend key; /kofi/webhook emails each subscriber this month's licence key (kofi.js)
 * D1 binding DB (migrations/): users, subscriptions, usage.
 *   MAX_INPUT_CHARS     default 40000
 */

import { SYSTEM_PROMPT, RESPONSE_FORMAT, trimBody, cleanVessels } from "./extract.js";
import { verifyMicrosoftToken } from "./auth.js";
import { getOrCreateUser, getStatus, consume, refund, checkLicenseKey, saveLicenseKey } from "./account.js";
import { handleKofiPayment } from "./kofi.js";

const OPENAI_URL = "https://api.openai.com/v1/chat/completions";
// Netlify gives up on a forwarded request after ~30s; fail cleanly before that.
const UPSTREAM_TIMEOUT_MS = 25000;

function json(status, body) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

// Constant-time comparison: hash both sides to equal-length digests, then XOR-compare.
async function safeEqual(a, b) {
  const enc = new TextEncoder();
  const [da, db] = await Promise.all([
    crypto.subtle.digest("SHA-256", enc.encode(a)),
    crypto.subtle.digest("SHA-256", enc.encode(b)),
  ]);
  const x = new Uint8Array(da), y = new Uint8Array(db);
  let diff = 0;
  for (let i = 0; i < x.length; i++) diff |= x[i] ^ y[i];
  return diff === 0;
}

async function authorized(request, env) {
  const token = request.headers.get("X-App-Token") || "";
  return Boolean(env.MAILAI_APP_SECRET) && (await safeEqual(token, env.MAILAI_APP_SECRET));
}

async function handleExtract(request, env) {
  if (!(await authorized(request, env))) {
    return json(401, { detail: "Unauthorized" });
  }

  let body;
  try {
    body = await request.json();
  } catch {
    return json(400, { detail: "Invalid JSON" });
  }

  const messages = body && Array.isArray(body.messages) ? body.messages : [];
  if (messages.length === 0 || messages.length > 2) {
    return json(400, { detail: "Invalid request" });
  }
  for (const m of messages) {
    if (!m || typeof m !== "object" || !["system", "user"].includes(m.role) || typeof m.content !== "string") {
      return json(400, { detail: "Invalid request" });
    }
  }

  // Limit input size to control costs.
  const maxChars = parseInt(env.MAX_INPUT_CHARS || "40000", 10);
  const totalChars = messages.reduce((n, m) => n + m.content.length, 0);
  if (totalChars > maxChars) {
    return json(400, { detail: "Input too large" });
  }

  try {
    const resp = await fetch(OPENAI_URL, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${env.OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: env.OPENAI_MODEL || "gpt-5.4-nano",
        messages: messages.map(({ role, content }) => ({ role, content })),
      }),
      signal: AbortSignal.timeout(UPSTREAM_TIMEOUT_MS),
    });
    if (!resp.ok) {
      console.log(`OpenAI error ${resp.status}: ${(await resp.text()).slice(0, 500)}`);
      return json(502, { detail: "Upstream error" });
    }
    const data = await resp.json();
    return json(200, { content: data.choices[0].message.content });
  } catch (e) {
    console.log(`OpenAI request failed: ${e}`);
    return json(502, { detail: "Upstream error" });
  }
}

/** Signed-in web-app user from "Authorization: Bearer <Microsoft ID token>", or null. */
async function webUser(request, env) {
  const bearer = (request.headers.get("Authorization") || "").match(/^Bearer\s+(.+)$/i)?.[1];
  return bearer ? verifyMicrosoftToken(bearer, env.MS_CLIENT_ID) : null;
}

async function handleMe(request, env) {
  const user = await webUser(request, env);
  if (!user) return json(401, { detail: "Unauthorized" });
  return json(200, await getStatus(env, await getOrCreateUser(env, user)));
}

async function handleLicense(request, env) {
  const user = await webUser(request, env);
  if (!user) return json(401, { detail: "Unauthorized" });
  let body;
  try {
    body = await request.json();
  } catch {
    return json(400, { detail: "Invalid JSON" });
  }
  const expires = await checkLicenseKey(body?.key, env.LICENSE_SECRET);
  if (!expires) return json(400, { detail: "That key isn't valid or has expired." });
  const row = await getOrCreateUser(env, user);
  await saveLicenseKey(env, row.user_id, expires);
  return json(200, await getStatus(env, { ...row, key_expires: expires }));
}

// Ko-fi posts application/x-www-form-urlencoded with one field, `data`, holding the JSON payload.
async function handleKofi(request, env) {
  let d;
  try {
    d = JSON.parse((await request.formData()).get("data"));
  } catch {
    return json(400, { detail: "Invalid payload" });
  }
  if (!env.KOFI_VERIFICATION_TOKEN || !(await safeEqual(String(d?.verification_token || ""), env.KOFI_VERIFICATION_TOKEN))) {
    return json(401, { detail: "Unauthorized" });
  }
  // Ko-fi retries on non-200, so always 200 once the token checks out (failures are logged).
  return json(200, { ok: true, result: await handleKofiPayment(env, d) });
}

async function handleExtractV2(request, env) {
  const user = await webUser(request, env);
  if (!user && !(await authorized(request, env))) {
    return json(401, { detail: "Unauthorized" });
  }
  // Web-app users: plan + daily cap, checked BEFORE spending on OpenAI. (App token = eval scripts: no limits.)
  let account = null;
  if (user) {
    account = await getOrCreateUser(env, user);
    const gate = await consume(env, account);
    if (!gate.ok) {
      return json(gate.reason === "expired" ? 402 : 429, { detail: gate.reason, status: gate.status });
    }
  }
  let body;
  try {
    body = await request.json();
  } catch {
    return json(400, { detail: "Invalid JSON" });
  }
  const subject = typeof body?.subject === "string" ? body.subject : "";
  if (typeof body?.body !== "string" || !body.body.trim()) {
    return json(400, { detail: "Invalid request" });
  }
  const text = `Subject: ${subject}

${trimBody(body.body)}`;
  if (text.length > parseInt(env.MAX_INPUT_CHARS || "40000", 10)) {
    return json(400, { detail: "Input too large" });
  }

  try {
    const resp = await fetch(OPENAI_URL, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${env.OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: env.OPENAI_MODEL_V2 || "gpt-5.4-mini",
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: text },
        ],
        response_format: RESPONSE_FORMAT,
      }),
      signal: AbortSignal.timeout(UPSTREAM_TIMEOUT_MS),
    });
    if (!resp.ok) {
      console.log(`OpenAI error ${resp.status}: ${(await resp.text()).slice(0, 500)}`);
      if (account) await refund(env, account);
      return json(502, { detail: "Upstream error" });
    }
    const data = await resp.json();
    const parsed = JSON.parse(data.choices[0].message.content);
    return json(200, { vessels: cleanVessels(parsed.vessels) });
  } catch (e) {
    // Never log email content; only the error.
    console.log(`v2 extract failed: ${e}`);
    if (account) await refund(env, account);
    return json(502, { detail: "Upstream error" });
  }
}

export default {
  async fetch(request, env) {
    const { pathname } = new URL(request.url);

    if (request.method === "GET" && (pathname === "/health" || pathname === "/v1/health")) {
      return json(200, { ok: true });
    }
    if (pathname === "/v1/extract") {
      if (request.method !== "POST") return json(405, { detail: "Method not allowed" });
      return handleExtract(request, env);
    }
    if (pathname === "/v2/extract") {
      if (request.method !== "POST") return json(405, { detail: "Method not allowed" });
      return handleExtractV2(request, env);
    }
    if (pathname === "/v2/me") {
      if (request.method !== "GET") return json(405, { detail: "Method not allowed" });
      return handleMe(request, env);
    }
    if (pathname === "/v2/license") {
      if (request.method !== "POST") return json(405, { detail: "Method not allowed" });
      return handleLicense(request, env);
    }
    if (pathname === "/kofi/webhook") {
      if (request.method !== "POST") return json(405, { detail: "Method not allowed" });
      return handleKofi(request, env);
    }
    return json(404, { detail: "Not found" });
  },
};
