/**
 * Ko-fi subscription payments (POST /kofi/webhook). Legacy: Ko-fi customers have moved to Stripe and
 * licence keys were retired on 2026-11-27, so this only records the payer's email (which still unlocks
 * Pro for that Microsoft account until the payment period ends). No emails are sent.
 */
import { recordKofiPayment } from "./account.js";

export async function handleKofiPayment(env, d) {
  const isSub = d?.is_subscription_payment === true || d?.type === "Subscription";
  if (!isSub || !d.email) return "ignored";
  try {
    await recordKofiPayment(env, d);
    return "recorded";
  } catch (e) {
    console.error(`kofi: recording payment failed: ${e}`);
    return "not recorded";
  }
}
