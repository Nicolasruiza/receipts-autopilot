# Receipts Autopilot

Snap a receipt on your phone. Claude reads it, picks the expense category and files it in your Google Drive.

- **One tap:** a phone web app with two buttons, Take photo and Import.
- **AI filing:** every hour, Claude extracts date, merchant, total and category from new receipts.
- **Organized Drive:** `Receipts/2026/10-October/2026-10-25_meals_starbucks.jpg`, filed by the date printed on the receipt.
- **Expense log:** one Google Sheets row per receipt.
- **Review queue:** blurry or unclear receipts go to `_Review` instead of being guessed.
- **No servers:** runs on Google Apps Script in your own account. You only pay a few cents per receipt to the Claude API.

## Use it

**Easiest:** copy the template project and follow the 8-step guide (~15 min).

- Template: https://script.google.com/d/1hf1dI-1wOgyn62lJAdAAw_Tce9xCXjIVo4SPXG7Cabvhwy-LCJDxUlgM/edit?usp=sharing
- Setup guide: https://claude.ai/artifact/67Cz12983AWvSmMpnhtAYs
- Case study and live demo: https://nicolasruiz.net/projects/receipts-autopilot.html

**From this repo:** create a project at [script.google.com](https://script.google.com), paste `Code.gs`, add an HTML file named exactly `Index` with `Index.html`, then follow the guide from step 2.

## Requirements

- A Google account (Gmail or Workspace)
- A [Claude API](https://console.anthropic.com) key with a little credit

## How it works

```
iPhone (Take photo / Import)
  → Apps Script web app → Drive /_Inbox
  → hourly trigger → Claude API (vision) → JSON: date, merchant, category, total, confidence
  → Drive /Year/MM-Month/  (or /_Review if confidence < 0.6)
  → Google Sheets log
```

## Customize

Edit the `MY_CONFIG` block at the top of `Code.gs` (company name, location, folder name, app title). Edit `CATEGORIES` to change the expense categories.

## Privacy

Everything runs in your own Google account. Receipts are sent only to the Claude API for reading. The app link is protected by a random key and can only upload, never read.

## License

MIT. Built by Nicolas Ruiz with Claude.
