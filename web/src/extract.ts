import { api, LimitError, type Status } from "./account";

export type Vessel = {
  name: string;
  dwt_k: number | null;
  built: number | null;
  open_location: string | null;
  open_country?: string | null; // ISO code; absent from older API versions
  open_date: string | null;
};

// Cheap local check before spending an API call: does this look like a position / tonnage email?
const SIGNALS = [
  /\bM\.?\s?V\b/,
  /\bDWT\b/i,
  /\bO\s*\/\s*A\b/i,
  /\bopen(ing)?\b/i,
  /\b(tonnage|positions?|open list)\b/i,
  /\b(bulk carrier|handy(size)?|supramax|ultramax|panamax|kamsarmax|capesize)\b/i,
  /\b\d{2,3}\s?K\b/,
];

export function looksRelevant(subject: string, body: string): boolean {
  const head = subject + "\n" + body.split(/\r?\n/, 100).join("\n");
  return SIGNALS.some((re) => re.test(head));
}

/** Extract open vessel positions from one email via the Mail AI API. Throws LimitError when blocked. */
export async function extractVessels(subject: string, body: string): Promise<Vessel[]> {
  for (let attempt = 0; attempt < 3; attempt++) {
    const res = await api("/v2/extract", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ subject, body }),
    });
    if (res.ok) return ((await res.json()) as { vessels: Vessel[] }).vessels;
    if (res.status === 402 || res.status === 429) {
      const b = (await res.json()) as { detail: "expired" | "daily_cap"; status: Status };
      throw new LimitError(b.detail, b.status);
    }
    if (res.status === 400) throw new Error("Email too large or unreadable");
    if (attempt < 2) await new Promise((r) => setTimeout(r, 1000 * 2 ** attempt));
    else throw new Error(`Extraction failed (${res.status})`);
  }
  return [];
}

/** Run fn over items with at most `limit` in flight. */
export async function pool<T>(items: T[], limit: number, fn: (item: T) => Promise<void>): Promise<void> {
  let i = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (i < items.length) await fn(items[i++]);
  });
  await Promise.all(workers);
}
