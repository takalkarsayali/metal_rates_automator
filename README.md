# Metal Rates Automator

Automates the daily task described by the RS team: fetch Au/Ag/Pt/Pd (LBMA)
and Rh/Ir/Ru (Johnson Matthey), log them into the master Excel tracker, and
email the team a formatted summary — with no manual browsing or copy-paste.

## ⚠️ Important: LBMA data source changed (found while building this)

LBMA moved their **historical/tabulated** price table behind a paid,
licensed **MyLBMA Portal** (rolled out from Nov 2025; platinum/palladium
benchmark administration also moved from LME to IBA on 1 Jul 2026). The
public "Table" tab on the prices page — the one your existing manual
process uses — now just shows "Sign in to view the tables of data."

**This script does not scrape that table.** Instead it pulls straight from
LBMA's public JSON quote feed, which powers their chart and was left
public:
```
https://prices.lbma.org.uk/json/gold_pm.json
https://prices.lbma.org.uk/json/silver.json
https://prices.lbma.org.uk/json/platinum_pm.json
https://prices.lbma.org.uk/json/palladium_pm.json
```
This is simpler and more reliable than the original Selenium plan for
LBMA — no browser needed for those four metals at all. If your organisation
already holds (or is considering) an IBA licence for other reasons, note
that this feed is a courtesy/chart-support feed, not the licensed
benchmark distribution channel — worth a quick check with your compliance/
legal contact if this data feeds into any pricing or contractual use, per
LBMA's licensing notice.

Johnson Matthey's page still renders a live table for Rh/Ir/Ru, so that
part still uses a browser-based approach (Selenium), same as originally
planned.

## Setup

1. `pip install -r requirements.txt`
2. Install Chrome + a matching chromedriver (or `pip install
   webdriver-manager` and adjust `jm_fetcher.py` to use it) — needed only
   for the Johnson Matthey fallback path.
3. Edit `config.py`:
   - `EXCEL_PATH` — path to your master workbook
   - `SMTP_SERVER` / `SMTP_PORT` / `SMTP_USERNAME` — your corporate mail
     server details
   - `EMAIL_TO` / `EMAIL_CC` — the RS distribution list
   - Set `SMTP_PASSWORD` as an environment variable (never commit it to
     code): `setx SMTP_PASSWORD "..."` on Windows, or your mail server may
     support Windows-integrated/relay auth with no password at all — check
     with IT.
4. Test each piece independently before wiring up the full run:
   ```
   python lbma_fetcher.py     # should print 4 prices
   python jm_fetcher.py       # should print 3 prices
   python excel_updater.py    # writes a test row using sample prices
   python main.py             # full end-to-end run
   ```

## Scheduling

Run once each weekday, after both benchmarks are set for the day (LBMA PM
fix ~15:00 London / ~8:30pm IST; JM updates through the day) — e.g. 8:30pm
IST is a safe time.

**Windows Task Scheduler:**
- Trigger: Daily, weekdays, 20:30
- Action: `python.exe` with argument `C:\path\to\main.py`
- "Start in": the project folder (so relative paths resolve)

**Linux/cron:**
```
30 20 * * 1-5  cd /path/to/metal_rates_automator && /usr/bin/python3 main.py
```

## What it does each run

1. Skips automatically on weekends (no LBMA/JM fixings).
2. Fetches LBMA JSON feeds for Au/Ag/Pt/Pd, falling back to the most recent
   prior trading day if today's isn't published yet (bank holidays).
3. Fetches JM's page for Rh/Ir/Ru (fast HTTP attempt first, Selenium
   fallback).
4. Aborts the run (and logs why) rather than writing a partial/incorrect
   row if any of the 7 prices couldn't be fetched.
5. Appends a new date column to the current month's sheet in the Excel
   tracker (creating the sheet automatically on the 1st of a new month, by
   copying the previous month's structure).
6. Recalculates that month's Min/Date and Max/Date block.
7. Emails the RS team an HTML table of today's prices plus the month's
   min/max — formatted like the summary you currently circulate manually.

## Files

| File | Purpose |
|---|---|
| `config.py` | All settings — paths, layout, SMTP, recipients |
| `lbma_fetcher.py` | Au/Ag/Pt/Pd via LBMA's public JSON feed |
| `jm_fetcher.py` | Rh/Ir/Ru via Johnson Matthey (HTTP, Selenium fallback) |
| `excel_updater.py` | Appends today's column, updates Min/Max, creates new month sheets |
| `emailer.py` | Builds & sends the HTML summary email |
| `main.py` | Orchestrates the full daily run — **this is what you schedule** |

## Notes / things to sanity-check on your machine

- I couldn't execute this against your actual network from here (this
  environment can't reach `matthey.com` or `lbma.org.uk`), so please run
  `lbma_fetcher.py` and `jm_fetcher.py` standalone first and check the
  output against today's real prices before trusting the full pipeline.
- The exact JSON filenames for silver/platinum/palladium on LBMA's feed
  are inferred from the known `gold_pm.json` pattern and public
  documentation — `lbma_fetcher.py` tries sensible fallback filenames
  automatically, but if all fail for one metal, open
  `https://prices.lbma.org.uk/json/` variants in a browser dev tools
  Network tab (as you'd do for the JM portlet) to confirm the exact name.
- Johnson Matthey's table is behind client-side JS — if the fast HTTP path
  in `jm_fetcher.py` fails (likely, until tested), the Selenium fallback
  handles it, but you'll want to verify the row-label matching still
  finds "Rhodium/Iridium/Ruthenium" if JM ever restyles the page.
