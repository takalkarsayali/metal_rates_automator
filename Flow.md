# Execution Flow — Metal Rates Automator

> This document describes the current intended flow for `main.py` and the source, workbook, and email modules. Keep it synchronized with code changes. If implementation behavior differs, code is authoritative and this document must be corrected.

## 1. Overview

```text
Start: main.py / run(target_date)
        |
        v
Initialize logging and load config
        |
        v
Skip weekend run
        |
        v
Fetch LBMA fixings (PM required, AM best-effort)
        |
        v
Fetch Johnson Matthey PGM rates
        |
        v
Validate required seven metal prices and their source dates
        |
        v
Resolve common effective date for mandatory metals
        |
        +------------------------------+
        |                              |
        v                              v
Fetch RBI USD/INR, EUR/INR       Fetch IBJA Silver 999 AM/PM
as of metals' effective date    as of requested run date
        |                              |
        v                              v
Keep actual RBI publication      Keep actual IBJA session dates
dates from fx_dates              from ibja_dates
        |                              |
        +---------------+--------------+
                        |
                        v
Build dated session updates:
metal PM, available metal AM, IBJA AM, IBJA PM
                        |
                        v
Sort sessions chronologically; AM before PM on same date
                        |
                        v
Update Excel by each session's own date
                        |
                        v
Write RBI values using each currency's actual publication date
                        |
                        v
Read monthly Min/Max/Average data for metals' effective month
                        |
                        v
Prepare email inputs and render HTML report
                        |
                        v
Send email with updated workbook attachment
                        |
                        v
Print summary and log completion
```

## 2. Entry point and startup

- Entry point: `main.py`.
- The script calls `setup_logging()` and then `run()` when executed directly.
- `run(target_date=None)` defaults to the current date.
- Weekend runs are skipped because LBMA/JM do not publish the expected daily prices on weekends.
- Configuration is loaded through `config.py`; logging setup is handled by the project's logging module/configuration.

## 3. Fetching and validating source data

### LBMA — `lbma_fetcher.py`

- `fetch_lbma_fixings(target_date)` returns separate AM and PM structures containing `prices` and `dates`.
- PM prices (and the daily Silver fixing where applicable) are used for the mandatory metal set.
- AM values are best-effort; missing AM values are logged and do not automatically invalidate otherwise complete PM data.
- The LBMA fetcher uses browser-based retrieval where direct JSON endpoints are blocked or unsuitable.
- Validate that the returned metal values and dates correspond to the expected source records.

### Johnson Matthey — `jm_fetcher.py`

- `fetch_jm_prices(target_date)` supplies PGM prices and source dates for Rhodium, Iridium, and Ruthenium.
- The source's London 09:00 values are used according to the fetcher's parsing rules.
- A partially populated current row must not be treated as complete without checking dates and values.

### Mandatory metals and effective date

- The mandatory set comprises Gold, Silver, Platinum, Palladium, Rhodium, Iridium, and Ruthenium.
- The main workflow validates the required values and source dates before proceeding.
- `_effective_price_date(source_dates)` resolves the common effective date used to label the mandatory metals run. The current rule uses the earliest of the normalized dates so the common date is not later than any of the selected metal prices.
- AM fixing dates may differ from the PM/effective date; AM is written under its own effective AM date.

### IBJA — `ibja_fetcher.py`

- `fetch_ibja_silver999(target_date)` returns `ibja_prices` and `ibja_dates`.
- IBJA Silver 999 is a reference rate in INR/kg; it is not the same series as LBMA Silver in USD/troy ounce.
- AM and PM rates must be mapped only when the source structure identifies those sessions. Do not infer a session or fabricate a missing value.
- Each IBJA session is written to Excel using its own publication date.
- For email generation, available non-`None` IBJA rates are included even when their dates differ from the main metals effective date. The publication dates should be shown with the corresponding AM/PM values so the email does not imply they share the metals date.

### RBI — `rbi_fetcher.py`

- `fetch_rbi_exchange_rates(price_date)` retrieves USD/INR and EUR/INR rates as of the metals effective date.
- RBI may return the latest available publication date on or before that date, for example because of weekends or holidays.
- `fx_prices` and `fx_dates` must be treated as paired data. Each currency must be written under its own actual publication date, not automatically under the metals effective date.
- If an RBI date cannot be found in the appropriate workbook sheet/date column, log the issue and do not write the rate under a different date.

## 4. Excel update flow — `excel_updater.py` and `main.py`

1. Assemble update sessions for the mandatory metals PM values, available metal AM values, and available IBJA AM/PM values.
2. Associate each session with the date returned for that session.
3. Sort updates by date; on the same date, AM precedes PM.
4. Call the session-specific Excel update helper for each update.
5. Update the workbook using the configured layout, including AM/PM columns and the applicable date/month sheet.
6. Update the monthly summary for the metal rows using valid recorded daily rates.
7. Write RBI exchange rates separately using their actual `fx_dates`.
8. Save the workbook before attempting to send the email.

Workbook layout currently described in project code:
- First data column: C.
- Date row: 4.
- AM/PM columns are paired by day (C/D, E/F, G/H, and so on).
- Metal rows: 6–12.
- IBJA Silver 999 row: 15.
- RBI USD/INR and EUR/INR rows: configured in `config.LAYOUT`.
- Monthly summary rows: 22–28.

Check `config.py` and `excel_updater.py` before relying on these coordinates if the workbook template changes. Close the workbook in Excel during automated runs if the file would otherwise be locked.

### Monthly Min/Max/Average

- Monthly summaries use the rates recorded in the month, not `(min + max) / 2`.
- Average is the arithmetic mean of valid daily values recorded for each metal from the start of the month through the latest recorded date.
- Blank cells are skipped; merged duplicate cells are counted once.
- AM and PM are separate observations when both are valid and recorded.
- IBJA Silver 999 and RBI exchange rates are not mixed into the seven-metal USD/troy-ounce min/max/average series.

## 5. Email flow — `emailer.py`

1. `main.py` builds the email data through `_build_email_inputs(...)`.
2. The email report contains:
   - Seven metal rows with AM/PM values where available.
   - IBJA Silver 999 AM and PM rates in INR/kg, with their own publication dates.
   - RBI USD/INR and EUR/INR rates.
   - Monthly minimum and maximum values and their dates (and the configured monthly summary data).
3. `build_html(...)` formats values and dates. Missing values remain blank; they are not converted to zero.
4. `send_email(...)` creates the MIME message, attaches the saved Excel workbook if the path exists, and sends it using SMTP settings from `config.py`.
5. Email recipients and SMTP settings are configuration-controlled. Never log passwords or tokens.

The email's main heading date is the metals' common effective date. IBJA and RBI values can legitimately have different publication dates; show those dates separately rather than misrepresenting them as the metals date.

## 6. Error handling and run summary

- LBMA/JM failures or missing mandatory metal values can prevent the report from being sent.
- IBJA and RBI are supplementary sources; source failures are logged and the run may continue with those sections blank, according to current `main.py` behavior.
- Workbook session-update errors are logged; the current workflow stops before sending the email if a session update fails.
- RBI write failures are logged separately. Check the log and workbook before assuming the rate was recorded.
- The run summary reports selected metal dates, source values, IBJA dates, RBI dates, workbook path, sheet, and recipients.
- The email send should be considered successful only when SMTP sending completes without raising an exception.

## 7. Scheduled execution

For a Windows scheduled run:
- Run from the project directory and use the intended Python interpreter.
- Ensure the machine is available, network access and browser dependencies work, and the workbook is not locked.
- Keep `.env`, credentials, live workbooks, logs, backups, and internal BRD documents out of Git.
- Test the launcher manually before scheduling it. Monitor logs and verify the workbook and email after initial scheduled runs.

## 8. Troubleshooting

| Symptom | Inspect | Checks |
|---|---|---|
| IBJA values blank in email | `_build_email_inputs()` and `emailer.py` | Verify `ibja_prices` keys, `ibja_dates`, non-`None` values, and date labels passed to `build_html()` |
| IBJA AM/PM appears swapped or wrong | `ibja_fetcher.py` | Confirm source table headers, row/cell mapping, units, and actual publication dates |
| RBI rate under wrong date | `fx_dates` and RBI workbook writer in `main.py` | Confirm each currency uses its own date and the matching worksheet/date column |
| AM column blank | `lbma_fetcher.py`, `am_prices`, `am_dates` | Confirm source published AM and the AM date matches the expected row |
| `too many values to unpack` | Fetcher return contract and caller | Compare actual return shape with caller unpacking |
| LBMA HTTP 403 | `lbma_fetcher.py` | Check browser-based retrieval and payload validation |
| Email missing attachment | `emailer.py` | Check workbook path, file existence, lock, and attachment log |
| Monthly average looks incorrect | `monthly_average.py`, `excel_updater.py` | Check valid daily observations, AM/PM counting, blanks, and merged cells |

## 9. Documentation maintenance

After a behavior change:
1. Update this flow to reflect the real call order and data contracts.
2. Add a decision entry in `decisions.md` explaining the reason and trade-off.
3. Add or update tests for date selection, session mapping, workbook writes, and email inputs.
4. Record what was actually tested; do not claim unverified production behavior.
