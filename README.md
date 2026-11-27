# Mail AI

**Automated vessel position extraction from shipbroking emails to Excel.**

Mail AI connects to your Outlook inbox and uses GPT to extract vessel position data from shipbroking emails — MV name, DWT, build year, open port, open date and trade zone — straight into an Excel spreadsheet.

## Features

- **One-click extraction** — pick a date and time, extract every relevant email from that point on
- **Live listening mode** — watch the inbox in real time, with pause/resume
- **Smart deduplication** — duplicate vessels are filtered out across sessions
- **Port zone lookup** — ports are mapped to trade zones via the World Port Index, UN/LOCODE and your own custom zones
- **Bilingual interface** — English and Simplified Chinese

## Requirements

Windows 10+, with the Microsoft Outlook and Excel desktop apps installed.

## Install

Download the latest release [here](https://github.com/MailAI-Development/Mail-AI/releases/latest), run it, and follow the setup wizard.

---

## Development

```
app/          desktop app — main.py (PySide6 UI), logic.py (Outlook, extraction, zones, Excel, licensing, updates)
  assets/     bundled data (WPI + UN/LOCODE CSVs), fonts, icons, splash
packaging/    mailai.spec (PyInstaller), installer.iss (Inno Setup)
proxy/        Cloudflare Worker extraction proxy — holds the OpenAI key server-side
website/      mailai.uk static site (Netlify, see netlify.toml)
```

### How extraction works

The app never holds an OpenAI key. For each email it sends the trimmed text to the proxy (`POST {PROXY_URL}/v1/extract`, header `X-App-Token`), which calls OpenAI and returns the raw completion. The app parses the vessel fields and maps the port to a zone.

### Run locally

```powershell
python -m venv .venv; .venv\Scripts\activate
pip install -r app\requirements.txt
copy app\_build_secrets.example.py app\_build_secrets.py   # then fill in
python app\main.py
```

To run the proxy locally: from `proxy/`, put `OPENAI_API_KEY=...` and `MAILAI_APP_SECRET=...` in a `.dev.vars` file (gitignored), run `npx wrangler dev`, and point `PROXY_URL` at `http://localhost:8787`.

In production the app uses `https://mailai.uk`: Netlify serves the site and forwards `/v1/*` to the proxy host (rule in `website/_redirects`). Moving hosts = deploy elsewhere + edit that one line; no app rebuild.

### Deploy the proxy

From `proxy/`: `npx wrangler deploy`, then `npx wrangler secret put OPENAI_API_KEY` and `npx wrangler secret put MAILAI_APP_SECRET` (same value as `APP_TOKEN`). Put the resulting `*.workers.dev` host into `website/_redirects`.

### Secrets

Build-time values live in `app/_build_secrets.py` (gitignored) or the env vars `MAILAI_PROXY_URL`, `MAILAI_APP_TOKEN`, `MAILAI_LICENSE_SECRET`. The spec refuses to build without them. Never put an OpenAI key in the app.

### Release

1. Bump `APP_VERSION` in `app/logic.py`.
2. `pyinstaller packaging/mailai.spec` → `dist/Mail AI v<version>.exe`
3. Optional installer: `iscc /DAppVersion=<version> packaging\installer.iss`
4. Create a GitHub release tagged `v<version>` and attach the .exe — running copies pick it up through the built-in updater (it looks for any `.exe` asset on the latest release).

## Licence

MIT — see [LICENSE](LICENSE)
