/**
 * Create (or check) Mail AI's Stripe products, prices and customer-portal settings. Safe to re-run.
 *
 *   node scripts/stripe_setup.mjs            # test mode: STRIPE_SECRET_KEY=sk_test_… in proxy/.dev.vars
 *   node scripts/stripe_setup.mjs --live     # live mode: key passed in STRIPE_SECRET_KEY env var
 *
 * Prices get lookup keys "<plan>_<monthly|annual>"; the Worker finds prices by these, so no IDs to copy.
 * Prints STRIPE_PORTAL_CONFIG to set as a Worker variable.
 */
import fs from "node:fs";

const PLANS = [
  { plan: "starter", name: "Mail AI Starter", cap: 100, monthly: 1900, annual: 19000 },
  { plan: "pro", name: "Mail AI Pro", cap: 500, monthly: 4900, annual: 49000 },
  { plan: "max", name: "Mail AI Max", cap: 1500, monthly: 9900, annual: 99000 },
];

const live = process.argv.includes("--live");
let key = process.env.STRIPE_SECRET_KEY;
if (!key && !live) {
  const vars = fs.readFileSync(new URL("../.dev.vars", import.meta.url), "utf8");
  key = vars.match(/^STRIPE_SECRET_KEY=(.+)$/m)?.[1]?.trim();
}
if (!key) throw new Error("No STRIPE_SECRET_KEY (proxy/.dev.vars for test mode, env var for --live)");
// Keys look like sk_test_…, rk_test_…, sk_live_…, rk_live_…
const isLiveKey = /^[a-z]+_live_/.test(key);
if (isLiveKey !== live) throw new Error(live ? "--live needs a live key" : "Refusing a live key without --live");

function form(params, prefix = "", out = new URLSearchParams()) {
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null) continue;
    const name = prefix ? `${prefix}[${k}]` : k;
    if (typeof v === "object") form(v, name, out);
    else out.append(name, String(v));
  }
  return out;
}
async function api(method, path, params) {
  const url = `https://api.stripe.com/v1/${path}${method === "GET" && params ? `?${form(params)}` : ""}`;
  const res = await fetch(url, {
    method,
    headers: { Authorization: `Bearer ${key}`, ...(method === "POST" ? { "Content-Type": "application/x-www-form-urlencoded" } : {}) },
    body: method === "POST" ? form(params ?? {}) : undefined,
  });
  const body = await res.json();
  if (!res.ok) throw new Error(`${method} ${path}: ${res.status} ${JSON.stringify(body.error)}`);
  return body;
}

const portalProducts = [];
for (const p of PLANS) {
  const keys = [`${p.plan}_monthly`, `${p.plan}_annual`];
  const existing = (await api("GET", "prices", { lookup_keys: keys, active: true, expand: ["data.product"] })).data;
  let product = existing[0]?.product;
  if (!product) {
    product = await api("POST", "products", {
      name: p.name,
      description: `Up to ${p.cap.toLocaleString("en-GB")} position emails a day`,
      metadata: { plan: p.plan },
    });
    console.log(`created product ${p.name}`);
  }
  const prices = [];
  for (const [interval, amount] of [["monthly", p.monthly], ["annual", p.annual]]) {
    const lookup = `${p.plan}_${interval}`;
    let price = existing.find((x) => x.lookup_key === lookup);
    if (price && price.unit_amount !== amount) {
      console.log(`  ${lookup}: price changed (${price.unit_amount} -> ${amount}), creating a new price and moving the lookup key`);
      price = null;
    }
    if (!price) {
      price = await api("POST", "prices", {
        product: product.id, currency: "gbp", unit_amount: amount, lookup_key: lookup, transfer_lookup_key: true,
        recurring: { interval: interval === "monthly" ? "month" : "year" },
        tax_behavior: "inclusive",
      });
      console.log(`  created ${lookup}: £${amount / 100}`);
    } else {
      console.log(`  ok ${lookup}: £${amount / 100} (${price.id})`);
    }
    prices.push(price.id);
  }
  portalProducts.push({ product: product.id, prices });
}

// Customer portal: switch between the three plans, update card, invoices, cancel at period end.
const portal = await api("POST", "billing_portal/configurations", {
  business_profile: { headline: "Mail AI subscription" },
  features: {
    invoice_history: { enabled: true },
    payment_method_update: { enabled: true },
    customer_update: { enabled: true, allowed_updates: ["email", "address", "tax_id"] },
    subscription_cancel: { enabled: true, mode: "at_period_end" },
    subscription_update: {
      enabled: true,
      default_allowed_updates: ["price"],
      proration_behavior: "create_prorations",
      products: portalProducts,
    },
  },
});
console.log(`\nportal configuration: ${portal.id}`);
console.log(`Set it on the Worker:  STRIPE_PORTAL_CONFIG = "${portal.id}"  (wrangler.toml [vars])`);
