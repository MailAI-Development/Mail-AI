# Mail AI

**Automated vessel position extraction from shipbroking emails to Excel.**

Mail AI connects to your Outlook inbox and uses GPT to automatically extract vessel position data from shipbroking emails — MV name, DWT, open port, open date, and trade zone — directly into a structured Excel spreadsheet. No copy-pasting, no manual data entry.

---

## Features

- **One-click extraction** — pick a date and time, extract all relevant emails from that point forward
- **Live listening mode** — monitor your inbox in real-time as new emails arrive, with pause/resume control
- **Smart deduplication** — duplicate vessels are filtered out across sessions
- **Port zone lookup** — every extracted port is automatically mapped to its trade zone using the World Port Index
- **Bilingual interface** — full English and Simplified Chinese UI

---

## Requirements

- Windows 10 or later
- Microsoft Outlook (desktop app, must be installed)
- Microsoft Excel (desktop app, must be installed)

---

## Setup

1. Download the latest release [here](https://github.com/MailAI-Development/Mail-AI/releases/latest)
2. Open Mail AI from your Start Menu or desktop shortcut
3. Follow the setup process
4. Return to the home screen and click **Extract** or **Listen for emails**

---

## Licence

MIT — see [LICENSE](LICENSE)

---

## Contact

[hello@mailai.uk](mailto:hello@mailai.uk) · [mailai.uk](https://mailai.uk)

---

## Changelog

v1.0 - Initial release

v1.1 - What's new:
- **Vessel build year** is now extracted alongside MV, deadweight etc.
- Listening mode fixed
- Size reduced by **half**, from 150MB to 75MB
- Extraction table screen now stays **upon leaving main extraction screen**
- Custom zones fixed - are now **saved across sessions**

v1.2 - What's new:
- Greatly improved **zone detection**
- Greatly improved **email filtering**

v1.3 - What's new:
- **Freemium model introduced** - Mail AI is now free for up to **200 email extractions per month**
- Pro tier - **Unlimited extractions** for **£9/month** via Ko-fi (enter your monthly license key in Settings to activate)
- Privacy page at mailai.uk/privacy, which documents exactly what the app accesses, what gets sent to OpenAI, and what stays on your machine

v1.4 - What's new:
- Design overhaul of the app and website, with improved clarity of the interface overall
- Auto updating: all application versions from now on will have an auto-update from the app natively, with no need to reinstall
- Listening is auto-triggered after extraction ends
- Extraction table has been reordered to suitably prioritise important details
- Column ordering from A-Z is now integrated into the main extraction table; press column headings to order from A-Z or Z-A
- Newly listened emails now have a marker and are pinned to the top
- Important emails can now be starred and pinned to the top
- Each cell of the extraction table can be fully edited if data is not accurate
- Dropdowns for email address, folder, date and time for a better user experience
- Light/dark mode refined + fixed
- Fixed a bug where deduplication of vessels would persist over multiple days
- Fixed a bug where valid vessels would be dropped
