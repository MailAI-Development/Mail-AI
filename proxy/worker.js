/**
 * Mail AI API (Cloudflare Worker), reached at https://mailai.uk/v2/* (Netlify forwards it, see website/_redirects).
 * Holds the OpenAI key server-side; the web app at mailai.uk/app is its only client.
 *
 *   POST /v2/extract   Authorization: Bearer <Microsoft ID token>  (web app)  or X-App-Token (eval scripts)
 *                      body {"subject", "body"} -> {"vessels": [...]}; prompt, schema and cleanup in extract.js
 *   GET  /v2/me        plan, trial and usage for the signed-in user (account.js)
 *   POST /v2/checkout, /v2/checkout/confirm, /v2/billing-portal, /stripe/webhook   Stripe billing (stripe.js)
 *   POST /kofi/webhook legacy: records Ko-fi payments (kofi.js)
 *   /v1/extract, /v2/license   retired 2026-11-27 (desktop app, licence keys): 410 Gone
 *
 * Secrets: OPENAI_API_KEY; MAILAI_APP_SECRET (X-App-Token for eval scripts); KOFI_VERIFICATION_TOKEN;
 *          STRIPE_SECRET_KEY, STRIPE_WEBHOOK_SECRET, STRIPE_PORTAL_CONFIG, STRIPE_CHECKOUT_ALLOW.
 * Vars (wrangler.toml): OPENAI_MODEL_V2 (default gpt-5.4-mini), MAX_INPUT_CHARS, MS_CLIENT_ID.
 * D1 binding DB (migrations/): users (incl. Stripe plan), subscriptions (Ko-fi), usage, grants.
 */

import { SYSTEM_PROMPT, RESPONSE_FORMAT, trimBody, cleanVessels } from "./extract.js";
import { verifyMicrosoftToken } from "./auth.js";
import { getOrCreateUser, getStatus, consume, refund } from "./account.js";
import { handleKofiPayment } from "./kofi.js";
import { checkoutAllowed, createCheckout, confirmCheckout, createPortal, verifyWebhook, subscriptionIdFromEvent, syncSubscription } from "./stripe.js";

// Where Stripe sends people back to. The client says which; only these are accepted.
const APP_URLS = ["https://mailai.uk/app/", "http://localhost:5173/"];

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

async function webUser(request, env) {
  const bearer = (request.headers.get("Authorization") || "").match(/^Bearer\s+(.+)$/i)?.[1];
  return bearer ? verifyMicrosoftToken(bearer, env.MS_CLIENT_ID) : null;
}

async function handleMe(request, env) {
  const user = await webUser(request, env);
  if (!user) return json(401, { detail: "Unauthorized" });
  return json(200, await getStatus(env, await getOrCreateUser(env, user)));
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

async function readJson(request) {
  try {
    return await request.json();
  } catch {
    return null;
  }
}

async function handleCheckout(request, env) {
  const user = await webUser(request, env);
  if (!user) return json(401, { detail: "Unauthorized" });
  const body = await readJson(request);
  const row = await getOrCreateUser(env, user);
  if (!checkoutAllowed(env, row.email)) return json(403, { detail: "Checkout isn't open yet." });
  const appUrl = APP_URLS.includes(body?.app_url) ? body.app_url : APP_URLS[0];
  try {
    return json(200, { clientSecret: await createCheckout(env, row, body?.plan, body?.interval, appUrl) });
  } catch (e) {
    console.log(`checkout failed: ${e}`);
    return json(400, { detail: "Couldn't start checkout." });
  }
}

async function handleCheckoutConfirm(request, env) {
  const user = await webUser(request, env);
  if (!user) return json(401, { detail: "Unauthorized" });
  const body = await readJson(request);
  const row = await getOrCreateUser(env, user);
  try {
    await confirmCheckout(env, row, String(body?.session_id || ""));
  } catch (e) {
    console.log(`checkout confirm failed: ${e}`);
    return json(400, { detail: "Couldn't confirm that payment." });
  }
  return json(200, await getStatus(env, await getOrCreateUser(env, user)));
}

async function handlePortal(request, env) {
  const user = await webUser(request, env);
  if (!user) return json(401, { detail: "Unauthorized" });
  const body = await readJson(request);
  const row = await getOrCreateUser(env, user);
  try {
    return json(200, { url: await createPortal(env, row, APP_URLS.includes(body?.app_url) ? body.app_url : APP_URLS[0]) });
  } catch (e) {
    console.log(`portal failed: ${e}`);
    return json(400, { detail: "Couldn't open subscription settings." });
  }
}

// Stripe webhooks: verify the signature over the raw body, then re-read the subscription from Stripe.
async function handleStripeWebhook(request, env) {
  const raw = await request.text();
  if (!(await verifyWebhook(raw, request.headers.get("Stripe-Signature"), env.STRIPE_WEBHOOK_SECRET))) {
    return json(400, { detail: "Bad signature" });
  }
  const event = JSON.parse(raw);
  const subId = subscriptionIdFromEvent(event);
  if (subId) {
    try {
      await syncSubscription(env, subId, event.data?.object?.client_reference_id);
    } catch (e) {
      console.log(`stripe webhook ${event.type} sync failed: ${e}`);
      return json(500, { detail: "Sync failed" }); // Stripe retries
    }
  }
  return json(200, { received: true });
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
    // Retired 2026-11-27: the desktop app (v1) and licence-key redemption.
    if (pathname === "/v1/extract" || pathname === "/v2/license") return json(410, { detail: "The Mail AI desktop app and licence keys have been retired. Use Mail AI at https://mailai.uk/app" });
    if (pathname === "/v2/extract") {
      if (request.method !== "POST") return json(405, { detail: "Method not allowed" });
      return handleExtractV2(request, env);
    }
    if (pathname === "/v2/me") {
      if (request.method !== "GET") return json(405, { detail: "Method not allowed" });
      return handleMe(request, env);
    }
    if (request.method === "POST" && pathname === "/v2/checkout") return handleCheckout(request, env);
    if (request.method === "POST" && pathname === "/v2/checkout/confirm") return handleCheckoutConfirm(request, env);
    if (request.method === "POST" && pathname === "/v2/billing-portal") return handlePortal(request, env);
    if (request.method === "POST" && pathname === "/stripe/webhook") return handleStripeWebhook(request, env);
    if (pathname === "/kofi/webhook") {
      if (request.method !== "POST") return json(405, { detail: "Method not allowed" });
      return handleKofi(request, env);
    }
    return json(404, { detail: "Not found" });
  },
};
