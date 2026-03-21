# Mail AI

**Automated vessel position extraction from shipbroking emails to Excel.**

Mail AI connects to your Outlook inbox, reads incoming shipbroking emails, and uses GPT to extract vessel position data — MV name, DWT, open port, open date, and trade zone — directly into a structured Excel spreadsheet. No copy-pasting, no manual data entry.

---

## Features

- **One-click extraction** — pick a date and time, extract all relevant emails from that point forward
- **Live listening mode** — monitor your inbox in real-time, processing new emails as they arrive with pause/resume control
- **Smart deduplication** — fuzzy name matching prevents duplicate vessel entries across sessions
- **Port zone lookup** — every extracted port is automatically mapped to its trade zone using a UNLOCODE reference table
- **Bilingual interface** — full English and Simplified Chinese UI with live switching
- **Dark and light mode**

---

## Requirements

- Windows 10 or later
- Microsoft Outlook (desktop app, must be installed)
- An OpenAI API key — generate one at [platform.openai.com](https://platform.openai.com)

---

## Download

Head to the [releases page](https://github.com/yourusername/mailai/releases/latest) to download the latest version.

---

## Setup

1. **Run `MailAI_Setup.exe`** and follow the installer
2. **Open Mail AI** from your Start Menu or desktop shortcut
3. **Go to Filtering settings** and enter:
   - Your Outlook email address
   - The folder name you want to monitor (e.g. `Inbox`)
   - The path to your Excel spreadsheet (e.g. `C:\Users\You\extraction.xlsx`)
   - Your OpenAI API key
4. **Return to the home screen** and click Extract or Listen

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

Mail AI uses GPT-4.1 Nano for extraction, one of the most cost-efficient models available. Typical costs are fractions of a cent per email — a full session of 50 emails usually costs less than $0.10. You are billed directly by OpenAI based on your usage.

---

## How it works

```
Outlook inbox → filter relevant emails → preprocess body → GPT extraction → parse & clean → deduplicate → write to Excel
```

1. Connects to Outlook via Windows COM API
2. Filters emails by folder, sender, and date
3. Strips irrelevant lines, keeps only vessel data lines
4. Sends preprocessed text to GPT with a structured extraction prompt
5. Parses the response with regex into structured fields
6. Checks for duplicates using fuzzy vessel name matching
7. Looks up the port zone from a UNLOCODE reference CSV
8. Appends results to your Excel file

---

## Tech stack

- **Python** with PySide6 (GUI)
- **win32com** for Outlook integration
- **OpenAI API** (gpt-4.1-nano) for extraction
- **openpyxl** for Excel output
- **rapidfuzz** for duplicate detection
- **Syne** and **DM Mono** fonts

---

## Licence

MIT — see [LICENSE](LICENSE)

---

## Contact

[hello@mailai.uk](mailto:hello@mailai.uk) · [mailai.uk](https://mailai.uk)
