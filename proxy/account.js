/**
 * Accounts, plans and usage for the web app (Cloudflare D1, binding env.DB).
 *
 * Plan:  "starter" | "pro" | "max"   active Stripe subscription (stripe.js keeps users.plan* in sync)
 *        complimentary               `grants` table: exact email or company email domain (migrations/0003)
 *        "pro"                       legacy: Ko-fi subscription for the user's email, or a valid MAILAI-… key
 *        "trial"                     within TRIAL_DAYS of first use
 *        "expired"                   otherwise
 * Daily cap on extracted emails per plan (counted before the OpenAI call): see DAILY_CAP.
 */

import { checkoutAllowed } from "./stripe.js";

export const TRIAL_DAYS = 7;
export const DAILY_CAP = { trial: 100, starter: 100, pro: 500, max: 1500, expired: 0 };
// Stripe statuses that keep access: past_due gives the customer time to fix a failed card.
const PAID_STATUSES = new Set(["active", "trialing", "past_due"]);
const DAY_MS = 864e5;
// Domain grants never apply to shared consumer mail domains (anyone can sign up to these).
const PUBLIC_DOMAINS = new Set(["outlook.com", "hotmail.com", "live.com", "msn.com", "gmail.com", "googlemail.com",
  "yahoo.com", "icloud.com", "me.com", "aol.com", "proton.me", "protonmail.com", "hotmail.co.uk", "outlook.co.uk", "yahoo.co.uk"]);

/** Active complimentary grant for this email (exact address first, then its company domain), or null. */
async function findGrant(env, email) {
  if (!email) return null;
  const domain = email.split("@")[1] || "";
  const targets = [`email:${email}`, ...(domain && !PUBLIC_DOMAINS.has(domain) ? [`domain:${domain}`] : [])];
  const rows = await env.DB.prepare(
    `SELECT target, plan, expires FROM grants WHERE target IN (${targets.map((_, i) => `?${i + 1}`).join(",")})`,
  ).bind(...targets).all();
  const now = Date.now();
  const active = (rows.results ?? []).filter((g) => !g.expires || Date.parse(g.expires) > now);
  return active.find((g) => g.target.startsWith("email:")) ?? active[0] ?? null;
}
const KEY_VALID_DAYS = 35; // same as the desktop app: ~a month of access + a few days' grace

const today = () => new Date().toISOString().slice(0, 10);

/** Fetch the user's row, creating it (and starting the trial) on first sight. */
export async function getOrCreateUser(env, user) {
  const now = new Date().toISOString();
  const email = user.email ? user.email.toLowerCase() : null;
  await env.DB.prepare(
    `INSERT INTO users (user_id, email, name, trial_start, created_at) VALUES (?1, ?2, ?3, ?4, ?4)
     ON CONFLICT(user_id) DO UPDATE SET email = excluded.email, name = excluded.name`,
  ).bind(user.userId, email, user.name, now).run();
  return env.DB.prepare("SELECT * FROM users WHERE user_id = ?1").bind(user.userId).first();
}

/** Current plan and limits for a user row. */
export async function getStatus(env, row) {
  const now = Date.now();
  const sub = row.email
    ? await env.DB.prepare("SELECT expires, tier FROM subscriptions WHERE email = ?1").bind(row.email).first()
    : null;
  const subActive = sub && Date.parse(sub.expires) > now;
  const keyActive = row.key_expires && Date.parse(row.key_expires) > now;
  const trialEnds = Date.parse(row.trial_start) + TRIAL_DAYS * DAY_MS;
  const stripeActive = row.plan && PAID_STATUSES.has(row.plan_status) &&
    (!row.plan_period_end || Date.parse(row.plan_period_end) > now - DAY_MS); // 1 day grace for renewal webhooks
  const grant = stripeActive ? null : await findGrant(env, row.email);
  const plan = stripeActive ? row.plan : grant ? grant.plan : subActive || keyActive ? "pro" : now < trialEnds ? "trial" : "expired";
  const used = await env.DB.prepare("SELECT count FROM usage WHERE user_id = ?1 AND day = ?2")
    .bind(row.user_id, today()).first("count");
  return {
    plan,
    email: row.email,
    trial_days_left: plan === "trial" ? Math.max(1, Math.ceil((trialEnds - now) / DAY_MS)) : 0,
    pro_until: stripeActive ? row.plan_period_end : grant ? grant.expires : subActive ? sub.expires : keyActive ? row.key_expires : null,
    pro_source: stripeActive ? "stripe" : grant ? "grant" : subActive ? "kofi" : keyActive ? "key" : null,
    plan_interval: stripeActive ? row.plan_interval : null,
    plan_status: stripeActive ? row.plan_status : null,
    cancel_at_end: stripeActive ? Boolean(row.plan_cancel_at_end) : false,
    can_manage: Boolean(row.stripe_customer),
    can_checkout: checkoutAllowed(env, row.email),
    used_today: used ?? 0,
    daily_cap: DAILY_CAP[plan],
  };
}

/**
 * Count one extraction against today's cap. Returns {ok, status}; when not ok, nothing was counted
 * and the caller must not call OpenAI.
 */
export async function consume(env, row) {
  const status = await getStatus(env, row);
  if (status.plan === "expired") return { ok: false, reason: "expired", status };
  const count = await env.DB.prepare(
    `INSERT INTO usage (user_id, day, count) VALUES (?1, ?2, 1)
     ON CONFLICT(user_id, day) DO UPDATE SET count = count + 1
     RETURNING count`,
  ).bind(row.user_id, today()).first("count");
  if (count > status.daily_cap) {
    await env.DB.prepare("UPDATE usage SET count = count - 1 WHERE user_id = ?1 AND day = ?2")
      .bind(row.user_id, today()).run();
    return { ok: false, reason: "daily_cap", status: { ...status, used_today: count - 1 } };
  }
  return { ok: true, status: { ...status, used_today: count } };
}

/** Undo a counted extraction when the OpenAI call itself failed (the user shouldn't pay for our error). */
export async function refund(env, row) {
  await env.DB.prepare("UPDATE usage SET count = MAX(count - 1, 0) WHERE user_id = ?1 AND day = ?2")
    .bind(row.user_id, today()).run();
}

// --- Ko-fi webhook ---

/** Record a Ko-fi subscription payment. Payload is Ko-fi's documented webhook JSON. */
export async function recordKofiPayment(env, d) {
  if (!d?.email || !(d.type === "Subscription" || d.is_subscription_payment)) return false;
  const paid = Date.parse(d.timestamp) || Date.now();
  const expires = new Date(paid + KEY_VALID_DAYS * DAY_MS).toISOString();
  await env.DB.prepare(
    `INSERT INTO subscriptions (email, tier, expires, last_payment, last_tx) VALUES (?1, ?2, ?3, ?4, ?5)
     ON CONFLICT(email) DO UPDATE SET tier = excluded.tier, last_payment = excluded.last_payment,
       last_tx = excluded.last_tx, expires = MAX(subscriptions.expires, excluded.expires)`,
  ).bind(String(d.email).toLowerCase().trim(), d.tier_name ?? null, expires, new Date(paid).toISOString(), d.kofi_transaction_id ?? null).run();
  return true;
}
