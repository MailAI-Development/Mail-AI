import { idToken } from "./auth";

export const SUBSCRIBE_URL = "https://ko-fi.com/mailaiuk/tiers";

export type Plan = "starter" | "pro" | "max";
export type Interval = "monthly" | "annual";

export type Status = {
  plan: "trial" | Plan | "expired";
  email: string | null;
  trial_days_left: number;
  pro_until: string | null;
  pro_source: "stripe" | "grant" | "kofi" | "key" | null;
  plan_interval: "month" | "year" | null;
  plan_status: string | null;
  cancel_at_end: boolean;
  can_manage: boolean; // has a Stripe customer: show "Manage subscription"
  can_checkout: boolean; // in-app checkout open for this account (limited while testing)
  used_today: number;
  daily_cap: number;
};

export const PLAN_INFO: Record<Plan, { name: string; cap: number; monthly: number; annual: number; blurb: string }> = {
  starter: { name: "Starter", cap: 100, monthly: 19, annual: 190, blurb: "For a light inbox or trying it on one desk." },
  pro: { name: "Pro", cap: 500, monthly: 49, annual: 490, blurb: "For a working broker's daily position lists." },
  max: { name: "Max", cap: 1500, monthly: 99, annual: 990, blurb: "For heavy inboxes and busy desks." },
};

const appUrl = () => new URL(import.meta.env.BASE_URL, window.location.origin).href;

async function post<T>(path: string, body: unknown): Promise<T> {
  const res = await api(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  const out = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(out.detail || `Request failed (${res.status})`);
  return out as T;
}

/** Client secret for Stripe's embedded checkout. */
export async function startCheckout(plan: Plan, interval: Interval): Promise<string> {
  return (await post<{ clientSecret: string }>("/v2/checkout", { plan, interval, app_url: appUrl() })).clientSecret;
}

/** After payment: have the server sync the subscription now (webhooks may lag). Returns the new status. */
export function confirmCheckout(sessionId: string): Promise<Status> {
  return post<Status>("/v2/checkout/confirm", { session_id: sessionId });
}

export async function openBillingPortal(): Promise<void> {
  window.location.href = (await post<{ url: string }>("/v2/billing-portal", { app_url: appUrl() })).url;
}

/** The trial has ended, or today's cap is used up. Carries the latest status for the UI. */
export class LimitError extends Error {
  constructor(public reason: "expired" | "daily_cap", public status: Status) {
    super(reason === "expired" ? "Your free trial has ended" : "Daily limit reached");
  }
}

/** Call the Mail AI API with the user's Microsoft ID token; refreshes the token once on 401. */
export async function api(path: string, init: RequestInit = {}): Promise<Response> {
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(`${window.location.origin}${path}`, {
      ...init,
      headers: { ...init.headers, Authorization: `Bearer ${await idToken(attempt > 0)}` },
    });
    if (res.status !== 401 || attempt > 0) return res;
  }
}

export async function fetchStatus(): Promise<Status> {
  const res = await api("/v2/me");
  if (!res.ok) throw new Error(`Couldn't load your account (${res.status})`);
  return res.json();
}

export async function redeemKey(key: string): Promise<Status> {
  const res = await api("/v2/license", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ key }),
  });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(body.detail || `Couldn't check that key (${res.status})`);
  return body as Status;
}
