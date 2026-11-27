-- Mail AI web app: accounts, trial, licences, usage.

-- One row per Microsoft account (user_id = "<tenant id>:<object id>"). Created on first API call.
CREATE TABLE users (
  user_id      TEXT PRIMARY KEY,
  email        TEXT,              -- lowercased sign-in email / UPN, used to match Ko-fi payments
  name         TEXT,
  trial_start  TEXT NOT NULL,     -- ISO timestamp of first use
  key_expires  TEXT,              -- set when the user enters a valid MAILAI-… licence key
  created_at   TEXT NOT NULL
);
CREATE INDEX users_email ON users(email);

-- Ko-fi subscribers, keyed by the email they paid with. Extended by each monthly payment.
CREATE TABLE subscriptions (
  email         TEXT PRIMARY KEY,  -- lowercased
  tier          TEXT,
  expires       TEXT NOT NULL,     -- ISO; last payment + 35 days
  last_payment  TEXT NOT NULL,
  last_tx       TEXT               -- Ko-fi transaction id (idempotency)
);

-- Position emails extracted per user per UTC day (the daily cap).
CREATE TABLE usage (
  user_id  TEXT NOT NULL,
  day      TEXT NOT NULL,          -- YYYY-MM-DD (UTC)
  count    INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (user_id, day)
);
