import { idToken } from "./auth";

export const SUBSCRIBE_URL = "https://ko-fi.com/mailaiuk/tiers";

export type Status = {
  plan: "trial" | "pro" | "expired";
  email: string | null;
  trial_days_left: number;
  pro_until: string | null;
  pro_source: "kofi" | "key" | null;
  used_today: number;
  daily_cap: number;
};

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
