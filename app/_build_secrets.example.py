# Copy to app/_build_secrets.py (gitignored) and fill in before building.
# Each value can instead be supplied as an env var: MAILAI_PROXY_URL, MAILAI_APP_TOKEN,
# MAILAI_LICENSE_SECRET (env vars win).

PROXY_URL = "https://mailai.uk"   # Netlify forwards /v1/* to the proxy host (website/_redirects)
APP_TOKEN = "..."        # must equal MAILAI_APP_SECRET on the proxy
LICENSE_SECRET = "..."   # must match keygenerator.py and the Ko-fi Worker's LICENSE_SECRET
