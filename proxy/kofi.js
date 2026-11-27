/**
 * Ko-fi subscription payments (POST /kofi/webhook). Replaces the separate key-delivery Worker.
 *
 * For each subscription payment:
 *   1. record the payer's email in D1 (unlocks the web app for that Microsoft account automatically), then
 *   2. email this month's licence key via Resend (for the desktop app, or a web user who paid with another email).
 *
 * Secrets: KOFI_VERIFICATION_TOKEN, LICENSE_SECRET, RESEND_API_KEY.  Var: FROM_EMAIL.
 * Ko-fi retries non-200 responses, so after the token check this always answers 200 and logs failures.
 */
import { recordKofiPayment, issueLicenseKey } from "./account.js";

const escapeHtml = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

async function emailKey(env, to, key) {
  const html = `
    <p>Thanks for subscribing to <strong>Mail AI Pro</strong>!</p>
    <p><strong>Web app:</strong> sign in at <a href="https://mailai.uk/app/">mailai.uk/app</a> with
       ${escapeHtml(to)} and Pro is already active. Nothing to enter.</p>
    <p><strong>Desktop app, or signing in with a different email:</strong> use this month's licence key:</p>
    <p style="font:600 18px monospace;letter-spacing:1px;">${escapeHtml(key)}</p>
    <p>Desktop: Settings → Pro License → paste → Activate. Web: paste it on the "trial ended" screen.</p>
    <p>You'll get a fresh key each month while your subscription is active.</p>
    <p>— Mail AI · mailai.uk</p>`;
  const resp = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: { Authorization: `Bearer ${env.RESEND_API_KEY}`, "Content-Type": "application/json" },
    body: JSON.stringify({ from: env.FROM_EMAIL, to: [to], subject: "Your Mail AI Pro licence key", html }),
  });
  if (!resp.ok) throw new Error(`Resend ${resp.status}: ${(await resp.text()).slice(0, 300)}`);
}

/** Handle one verified Ko-fi payload. Returns a short status string for the response body. */
export async function handleKofiPayment(env, d) {
  const isSub = d?.is_subscription_payment === true || d?.type === "Subscription";
  if (!isSub || !d.email) return "ignored";
  try {
    await recordKofiPayment(env, d);
  } catch (e) {
    console.error(`kofi: recording payment failed: ${e}`);
  }
  if (!env.RESEND_API_KEY || !env.FROM_EMAIL) {
    console.error("kofi: RESEND_API_KEY/FROM_EMAIL not set; licence email not sent");
    return "recorded (email not configured)";
  }
  try {
    await emailKey(env, d.email, await issueLicenseKey(env.LICENSE_SECRET));
    return "recorded, key sent";
  } catch (e) {
    console.error(`kofi: licence email failed: ${e}`);
    return "recorded (email failed)";
  }
}
