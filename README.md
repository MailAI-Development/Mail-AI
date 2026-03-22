# Mail AI

**Automated vessel position extraction from shipbroking emails to Excel.**

Mail AI connects to your Outlook inbox, reads incoming shipbroking emails, and uses GPT to extract vessel position data — MV name, DWT, open port, open date, and trade zone — directly into a structured Excel spreadsheet. No copy-pasting, no manual data entry.

---

## Features

- **One-click extraction** — pick a date and time, extract all relevant emails from that point forward
- **Live listening mode** — monitor your inbox in real-time, processing new emails as they arrive with pause/resume control
- **Port zone lookup** — every extracted port is automatically mapped to its trade zone using the World Port Index
- **Bilingual interface** — full English and Simplified Chinese UI

---

## Requirements

- Windows 10 or later
- Microsoft Outlook Classic (desktop app, must be installed)
- An OpenAI API key — generate one at [platform.openai.com](https://platform.openai.com)

---

## Purchase

Purchase [here](https://mailai.gumroad.com/l/buy) via Gumroad.

---

## Setup

1. Run `Mail AI.exe`
2. Open from your Start Menu or desktop shortcut
3. Follow the setup instructions
4. Return to the home screen and click 'Extract' or 'Listen for emails'

---

## Outlook Trust Center

If Outlook shows a security warning when Mail AI tries to access your emails:

1. Close Outlook
2. Right-click the Outlook icon → **Run as Administrator**
3. Go to **File → Options → Trust Center → Trust Center Settings → Programmatic Access**
4. Select **Never warn me about suspicious activity**
5. Restart Outlook normally

If the option is greyed out, your machine may be managed by an IT administrator — ask them to whitelist Mail AI or adjust the Programmatic Access policy.

---

## OpenAI API costs

Mail AI uses GPT-5.4 Nano for extraction, one of the most cost-efficient models available. Typical costs are fractions of a cent per email — a full session of 50 emails usually costs less than $0.10. You are billed directly by OpenAI based on your usage.

---

## Contact

[hello@mailai.uk](mailto:hello@mailai.uk) · [mailai.uk](https://mailai.uk)

---

## Changelog

v1.0 - Initial release
