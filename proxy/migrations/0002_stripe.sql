-- Stripe subscriptions (in-app checkout). One subscription per Microsoft account.
ALTER TABLE users ADD COLUMN stripe_customer TEXT;      -- cus_…
ALTER TABLE users ADD COLUMN stripe_subscription TEXT;  -- sub_…
ALTER TABLE users ADD COLUMN plan TEXT;                 -- starter | pro | max (from the Stripe price)
ALTER TABLE users ADD COLUMN plan_interval TEXT;        -- month | year
ALTER TABLE users ADD COLUMN plan_status TEXT;          -- Stripe subscription status (active, past_due, canceled, …)
ALTER TABLE users ADD COLUMN plan_period_end TEXT;      -- ISO; access runs to here
ALTER TABLE users ADD COLUMN plan_cancel_at_end INTEGER NOT NULL DEFAULT 0;
CREATE INDEX users_stripe_customer ON users(stripe_customer);
