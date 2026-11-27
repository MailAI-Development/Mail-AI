import { useCallback, useRef, useState } from "react";
import { loadStripe, type Stripe } from "@stripe/stripe-js";
import { EmbeddedCheckout, EmbeddedCheckoutProvider } from "@stripe/react-stripe-js";
import {
  confirmCheckout, openBillingPortal, PLAN_INFO, startCheckout, SUBSCRIBE_URL,
  type Interval, type Plan, type Status,
} from "./account";

const PUBLISHABLE_KEY = import.meta.env.VITE_STRIPE_PUBLISHABLE_KEY as string | undefined;
let stripePromise: Promise<Stripe | null> | null = null;
const getStripe = () => (stripePromise ??= PUBLISHABLE_KEY ? loadStripe(PUBLISHABLE_KEY) : Promise.resolve(null));

/** Pricing cards, then Stripe's embedded checkout for the chosen plan. */
export function Plans({ status, onPaid, onClose }: { status: Status; onPaid: (s: Status) => void; onClose?: () => void }) {
  const [interval, setInterval] = useState<Interval>("annual");
  const [chosen, setChosen] = useState<Plan | null>(null);
  const [err, setErr] = useState("");

  const sessionId = useRef("");
  const fetchClientSecret = useCallback(async () => {
    const secret = await startCheckout(chosen!, interval);
    sessionId.current = secret.split("_secret_")[0]; // cs_… id is the part before "_secret_"
    return secret;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chosen, interval]);

  const onComplete = useCallback(async () => {
    try {
      onPaid(await confirmCheckout(sessionId.current));
    } catch (e) {
      setErr((e as Error).message);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const current = status.pro_source === "stripe" ? status.plan : null;

  if (!status.can_checkout || !PUBLISHABLE_KEY) {
    // In-app checkout not open for this account yet: keep the existing Ko-fi route.
    return (
      <section className="plans">
        <h2>Upgrade</h2>
        <p className="muted">Subscribe on Ko-fi with the email you sign in with ({status.email}) and Mail AI unlocks automatically.</p>
        <p><a className="button primary" href={SUBSCRIBE_URL} target="_blank" rel="noreferrer">Subscribe on Ko-fi</a></p>
      </section>
    );
  }

  if (chosen) {
    return (
      <section className="plans">
        <div className="plans-head">
          <h2>{PLAN_INFO[chosen].name} · {interval === "annual" ? "annual" : "monthly"}</h2>
          <button className="link" onClick={() => setChosen(null)}>← Change plan</button>
        </div>
        {err && <p className="warn">{err}</p>}
        <div className="checkout">
          <EmbeddedCheckoutProvider stripe={getStripe()} options={{ fetchClientSecret, onComplete }}>
            <EmbeddedCheckout />
          </EmbeddedCheckoutProvider>
        </div>
      </section>
    );
  }

  return (
    <section className="plans">
      <div className="plans-head">
        <h2>Choose a plan</h2>
        <div className="toggle" role="group" aria-label="Billing period">
          <button className={interval === "monthly" ? "on" : ""} aria-pressed={interval === "monthly"} onClick={() => setInterval("monthly")}>Monthly</button>
          <button className={interval === "annual" ? "on" : ""} aria-pressed={interval === "annual"} onClick={() => setInterval("annual")}>
            Annual <span>2 months free</span>
          </button>
        </div>
        {onClose && <button className="link" onClick={onClose}>Close</button>}
      </div>
      <div className="plan-cards">
        {(Object.keys(PLAN_INFO) as Plan[]).map((p) => {
          const info = PLAN_INFO[p];
          const price = interval === "annual" ? info.annual : info.monthly;
          return (
            <div key={p} className={p === "pro" ? "plan-card featured" : "plan-card"}>
              <div className="plan-name">
                <h3>{info.name}</h3>
                {p === "pro" && <span className="badge">Most popular</span>}
              </div>
              <p className="price">£{price}<span>/{interval === "annual" ? "year" : "month"}</span></p>
              {interval === "annual" && <p className="muted small">£{(info.annual / 12).toFixed(2)} a month, billed yearly</p>}
              <p className="cap">{info.cap.toLocaleString()} position emails a day</p>
              <p className="muted small">{info.blurb}</p>
              {current === p ? (
                <button disabled>Current plan</button>
              ) : current ? (
                // Already subscribed: change plan in Stripe's portal (prorated), never a second subscription.
                <button onClick={() => openBillingPortal().catch((e) => setErr((e as Error).message))}>Switch to {info.name}</button>
              ) : (
                <button className={p === "pro" ? "primary" : ""} onClick={() => setChosen(p)}>Choose {info.name}</button>
              )}
            </div>
          );
        })}
      </div>
      {err && <p className="warn">{err}</p>}
      <p className="muted small">
        Cancel at any time; you keep access to Mail AI until the end of the period you've paid for. Don't worry – emails that aren't position lists
        are skipped and don't count towards the daily limit.
      </p>
    </section>
  );
}
