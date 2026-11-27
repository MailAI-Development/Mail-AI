/**
 * Stripe subscriptions for the web app: embedded Checkout, customer portal, webhook, and syncing a
 * subscription's state onto the user's row. Plain fetch against Stripe's REST API (no SDK).
 *
 * Plans are Stripe Prices with lookup keys "<plan>_<interval>": starter|pro|max × monthly|annual
 * (created by scripts/stripe_setup.mjs). The price's lookup key decides the plan.
 *
 * Secrets: STRIPE_SECRET_KEY, STRIPE_WEBHOOK_SECRET.
 * Var: STRIPE_PORTAL_CONFIG — customer-portal configuration id printed by scripts/stripe_setup.mjs.
 * Var: STRIPE_CHECKOUT_ALLOW — comma-separated emails allowed to start checkout ("*" = everyone).
 *      Keep it to your own email while testing with test-mode keys.
 */

export const PLANS = ["starter", "pro", "max"];
export const INTERVALS = { monthly: "month", annual: "year" };

/** Flatten {a: {b: [x]}} into Stripe's form encoding: a[b][0]=x */
function form(params, prefix = "", out = new URLSearchParams()) {
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null) continue;
    const key = prefix ? `${prefix}[${k}]` : k;
    if (typeof v === "object") form(v, key, out);
    else out.append(key, String(v));
  }
  return out;
}

export async function stripe(env, method, path, params) {
  if (!env.STRIPE_SECRET_KEY) throw new Error("STRIPE_SECRET_KEY not set");
  let url = `https://api.stripe.com/v1/${path}`;
  const init = { method, headers: { Authorization: `Bearer ${env.STRIPE_SECRET_KEY}` } };
  if (params && method === "GET") url += `?${form(params)}`;
  else if (params) {
    init.headers["Content-Type"] = "application/x-www-form-urlencoded";
    init.body = form(params);
  }
  const res = await fetch(url, init);
  const body = await res.json();
  if (!res.ok) throw new Error(`Stripe ${method} ${path}: ${res.status} ${body?.error?.message ?? ""}`);
  return body;
}

export function checkoutAllowed(env, email) {
  const allow = (env.STRIPE_CHECKOUT_ALLOW || "").split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
  return allow.includes("*") || (email && allow.includes(email.toLowerCase()));
}

/** Embedded Checkout session for a plan; returns the client secret. */
export async function createCheckout(env, row, plan, interval, appUrl) {
  if (!PLANS.includes(plan) || !INTERVALS[interval]) throw new Error("Unknown plan");
  const prices = await stripe(env, "GET", "prices", { lookup_keys: [`${plan}_${interval}`], active: true });
  const price = prices.data?.[0];
  if (!price) throw new Error(`No Stripe price with lookup key ${plan}_${interval}`);
  const session = await stripe(env, "POST", "checkout/sessions", {
    ui_mode: "embedded_page",
    mode: "subscription",
    line_items: [{ price: price.id, quantity: 1 }],
    return_url: `${appUrl}?checkout={CHECKOUT_SESSION_ID}`,
    client_reference_id: row.user_id,
    ...(row.stripe_customer ? { customer: row.stripe_customer } : row.email ? { customer_email: row.email } : {}),
    subscription_data: { metadata: { user_id: row.user_id } },
    metadata: { user_id: row.user_id },
    allow_promotion_codes: true,
  });
  return session.client_secret;
}

/** After the embedded checkout completes: link the customer + subscription to the user right away. */
export async function confirmCheckout(env, row, sessionId) {
  const session = await stripe(env, "GET", `checkout/sessions/${encodeURIComponent(sessionId)}`);
  if (session.client_reference_id !== row.user_id) throw new Error("Session belongs to another account");
  if (session.subscription) await syncSubscription(env, session.subscription, row.user_id);
  return session.status;
}

/** Customer portal (change plan, card, invoices, cancel). */
export async function createPortal(env, row, appUrl) {
  if (!row.stripe_customer) throw new Error("No subscription yet");
  const portal = await stripe(env, "POST", "billing_portal/sessions", {
    customer: row.stripe_customer,
    return_url: appUrl,
    configuration: env.STRIPE_PORTAL_CONFIG || undefined, // from scripts/stripe_setup.mjs
  });
  return portal.url;
}

/**
 * Fetch a subscription from Stripe and store its state on the owning user. Always reads the live object,
 * so it works whatever shape the triggering webhook had and whatever order events arrive in.
 */
export async function syncSubscription(env, subscriptionId, userIdHint) {
  const sub = await stripe(env, "GET", `subscriptions/${encodeURIComponent(subscriptionId)}`, { expand: ["items.data.price"] });
  const item = sub.items?.data?.[0];
  const key = item?.price?.lookup_key || "";
  const [plan, interval] = key.split("_");
  // current_period_end moved from the subscription to its items in newer API versions; accept either.
  const periodEnd = sub.current_period_end ?? item?.current_period_end;
  const userId = sub.metadata?.user_id || userIdHint;
  const where = userId ? "user_id = ?8" : "stripe_customer = ?8";
  await env.DB.prepare(
    `UPDATE users SET stripe_customer = ?1, stripe_subscription = ?2, plan = ?3, plan_interval = ?4, plan_status = ?5,
       plan_period_end = ?6, plan_cancel_at_end = ?7 WHERE ${where}`,
  ).bind(
    typeof sub.customer === "string" ? sub.customer : sub.customer?.id,
    sub.id,
    PLANS.includes(plan) ? plan : null,
    INTERVALS[interval] ?? null,
    sub.status,
    periodEnd ? new Date(periodEnd * 1000).toISOString() : null,
    sub.cancel_at_period_end || sub.cancel_at ? 1 : 0,
    userId || (typeof sub.customer === "string" ? sub.customer : sub.customer?.id),
  ).run();
}

/** Verify Stripe-Signature (t=…, v1=…) over "<t>.<raw body>" with HMAC-SHA256; 5-minute tolerance. */
export async function verifyWebhook(rawBody, header, secret) {
  if (!header || !secret) return false;
  let t = null;
  const sigs = [];
  for (const p of header.split(",")) {
    const [k, v] = p.split("=", 2);
    if (k === "t") t = v;
    else if (k === "v1") sigs.push(v);
  }
  if (!t || !sigs.length || Math.abs(Date.now() / 1000 - Number(t)) > 300) return false;
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const mac = new Uint8Array(await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(`${t}.${rawBody}`)));
  const expected = [...mac].map((b) => b.toString(16).padStart(2, "0")).join("");
  return sigs.some((s) => {
    if (s.length !== expected.length) return false;
    let diff = 0;
    for (let i = 0; i < s.length; i++) diff |= s.charCodeAt(i) ^ expected.charCodeAt(i);
    return diff === 0;
  });
}

/** Subscription id carried by a webhook event, for the events that affect access. */
export function subscriptionIdFromEvent(event) {
  const o = event?.data?.object ?? {};
  switch (event?.type) {
    case "customer.subscription.created":
    case "customer.subscription.updated":
    case "customer.subscription.deleted":
    case "customer.subscription.paused":
    case "customer.subscription.resumed":
      return o.id;
    case "checkout.session.completed":
      return typeof o.subscription === "string" ? o.subscription : o.subscription?.id;
    case "invoice.paid":
    case "invoice.payment_failed":
      // invoice.subscription moved under parent.subscription_details in newer API versions; accept either.
      return o.subscription ?? o.parent?.subscription_details?.subscription ?? null;
    default:
      return null;
  }
}
