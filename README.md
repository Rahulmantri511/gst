# GST Filing Status Tracker

A simple, automated tool to check **GSTR-1 and GSTR-3B filing status** for a bulk list of GSTINs directly from the GST portal — and write the results back to your Google Sheet automatically.

No manual portal browsing. No copy-pasting. Just load your sheet, click Run, then Push.

---

## What This Does

- Reads GSTINs from your Google Sheet
- Checks GSTR-1 and GSTR-3B filing status on the official GST portal for any selected financial year
- Shows Filed / Not Filed status live in the app
- Pushes results back to your sheet with year-specific columns (e.g. `GSTR1 2025-26`, `GSTR3B 2025-26`)
- **Automatically skips rows** that already have data — safe to re-run anytime
- **Saves your URLs** in the browser so you never have to re-enter them

---

## How to Use

### Step 1 — Prepare Your Google Sheet

1. Make sure your sheet has a column with GSTIN numbers (the column name can be anything — the app will auto-detect it)
2. Share the sheet: click **Share → Change to "Anyone with the link" → Viewer**
3. Copy the URL from your browser

### Step 2 — Set Up Google Apps Script (one time only)

To write results back to the sheet, you need a small script deployed as a Web App:

1. In your Google Sheet, click **Extensions → Apps Script**
2. Delete any existing code
3. Copy the script from the **"How to set up Apps Script?"** button in the app and paste it
4. Click **Deploy → New deployment → Web app**
5. Set:
   - **Execute as**: Me
   - **Who has access**: Anyone
6. Click **Deploy**, authorize when prompted, and copy the **Web App URL**

You only need to do this once.

### Step 3 — Run the App

```bash
bun install
bun run dev
```

Open [http://localhost:3000](http://localhost:3000)

1. **Step 1 (Connect Sheet)**: Paste your Google Sheet URL → click Load. Select the GSTIN column and Financial Year.
2. **Step 2 (Fetch GST Status)**: Click **Run Batch**. The app will check each GSTIN and show Filed / Not Filed.
3. **Step 3 (Push to Sheet)**: Paste your Apps Script Web App URL → click **Push Records to Sheet**.

Your sheet will have these columns added automatically:
- `GSTR1 2025-26` — Filing status (Filed / Not Filed)
- `GSTR3B 2025-26` — Filing status
- `GSTR1 Date 2025-26` — Latest filing date
- `GSTR3B Date 2025-26` — Latest filing date

---

## Features

| Feature | Details |
|---|---|
| Auto-save URLs | Sheet URL and Apps Script URL saved in browser localStorage |
| Auto-skip | Rows that already have filing data are automatically skipped |
| Pause / Resume | Can pause and resume mid-batch |
| Adjustable delay | Set 200ms to 2s delay between requests to avoid rate limits |
| Mobile friendly | Works on phone browsers too |
| No API key needed | Reads sheets via public CSV, fetches GST data from portal directly |

---

## Tech Stack

- **Next.js** (App Router, TypeScript)
- **Bun** — package manager and runtime
- **Google Apps Script** — writes data back to Google Sheets
- **GST Portal API** — `taxpayerReturnDetails` endpoint for filing data

---

## Notes

- The Google Sheet must be set to **public view** (Anyone with link) for the app to read it
- The GST portal sometimes rate-limits requests — use 1 second delay if you see errors
- The Apps Script Web App URL is stored locally in your browser and never sent to any server
