/**
 * Accounts, plans and usage for the web app (Cloudflare D1, binding env.DB).
 *
 * Plan:  "pro"     Ko-fi subscription for the user's email, or a valid MAILAI-… key entered by the user
 *        "trial"   within TRIAL_DAYS of first use
 *        "expired" otherwise
 * Daily cap on extracted emails: trial 100, pro 1,500 (counted before the OpenAI call).
 */

export const TRIAL_DAYS = 7;
export const DAILY_CAP = { trial: 100, pro: 1500, expired: 0 };
const DAY_MS = 864e5;
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
  const plan = subActive || keyActive ? "pro" : now < trialEnds ? "trial" : "expired";
  const used = await env.DB.prepare("SELECT count FROM usage WHERE user_id = ?1 AND day = ?2")
    .bind(row.user_id, today()).first("count");
  return {
    plan,
    email: row.email,
    trial_days_left: plan === "trial" ? Math.max(1, Math.ceil((trialEnds - now) / DAY_MS)) : 0,
    pro_until: subActive ? sub.expires : keyActive ? row.key_expires : null,
    pro_source: subActive ? "kofi" : keyActive ? "key" : null,
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

// --- Licence keys (same format and check as the desktop app: MAILAI-<YYYYMMDD|YYYYMM>-<sig>) ---

function base32(bytes) {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  let bits = 0, value = 0, out = "";
  for (const b of bytes) {
    value = (value << 8) | b;
    bits += 8;
    while (bits >= 5) {
      out += alphabet[(value >>> (bits - 5)) & 31];
      bits -= 5;
    }
  }
  if (bits > 0) out += alphabet[(value << (5 - bits)) & 31];
  return out;
}

/** Returns the key's expiry (ISO) if valid and unexpired, else null. */
export async function checkLicenseKey(key, secret) {
  if (!secret || typeof key !== "string") return null;
  const parts = key.trim().toUpperCase().split("-");
  if (parts.length !== 3 || parts[0] !== "MAILAI") return null;
  const [, datestr, sig] = parts;
  if (!/^\d{6}(\d{2})?$/.test(datestr)) return null;
  const issued = Date.UTC(+datestr.slice(0, 4), +datestr.slice(4, 6) - 1, datestr.length === 8 ? +datestr.slice(6, 8) : 1);
  const expires = issued + KEY_VALID_DAYS * DAY_MS;
  if (!(expires > Date.now())) return null;
  const expected = await signDate(datestr, secret);
  // constant-time compare
  if (sig.length !== expected.length) return null;
  let diff = 0;
  for (let i = 0; i < sig.length; i++) diff |= sig.charCodeAt(i) ^ expected.charCodeAt(i);
  return diff === 0 ? new Date(expires).toISOString() : null;
}

async function signDate(datestr, secret) {
  const hmacKey = await crypto.subtle.importKey("raw", new TextEncoder().encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  return base32(new Uint8Array(await crypto.subtle.sign("HMAC", hmacKey, new TextEncoder().encode(datestr)))).slice(0, 10);
}

/** Today's (UTC) licence key, identical to what the desktop app validates and the old Ko-fi Worker issued. */
export async function issueLicenseKey(secret) {
  if (!secret) throw new Error("LICENSE_SECRET not set");
  const ymd = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  return `MAILAI-${ymd}-${await signDate(ymd, secret)}`;
}

export async function saveLicenseKey(env, userId, expiresIso) {
  await env.DB.prepare("UPDATE users SET key_expires = ?2 WHERE user_id = ?1").bind(userId, expiresIso).run();
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
