-- Complimentary plans (friends & family, partners, testers). Managed by hand with `wrangler d1 execute`.
-- target: 'email:someone@firm.com' (one person) or 'domain:firm.com' (everyone signing in with @firm.com).
CREATE TABLE grants (
  target   TEXT PRIMARY KEY,
  plan     TEXT NOT NULL,        -- starter | pro | max
  expires  TEXT,                 -- ISO date; NULL = no end date
  note     TEXT,
  created  TEXT NOT NULL
);
