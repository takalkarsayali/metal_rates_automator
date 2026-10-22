# Engineering Decisions — Metal Rates Automator

> Record material decisions, their rationale, trade-offs, and validation. The current implementation is authoritative; mark entries superseded rather than silently deleting history.

## Project context

The Python automation retrieves metal prices, updates a historical Excel tracker, calculates monthly summaries, and emails an HTML report with the workbook attached. Sources include LBMA, Johnson Matthey, IBJA Silver 999, and RBI exchange rates.

## Decision log

### D-001 — Use Python for the automation workflow
- **Decision:** Implement the workflow in Python.
- **Reason:** Python supports HTTP retrieval, HTML parsing, browser automation, Excel updates, and SMTP email.
- **Trade-off:** Dependencies and runtime configuration must be maintained on the target Windows environment.
- **Status:** Adopted.

### D-002 — Use `requests` for ordinary HTTP retrieval
- **Decision:** Use `requests` when a source can be retrieved using standard HTTP.
- **Reason:** It supports request headers, timeouts, status checks, and response handling.
- **Trade-off:** It cannot execute JavaScript or reproduce all browser sessions.
- **Implementation guidance:** Use timeouts, check HTTP status, and validate extracted data rather than assuming a successful response means valid rates.
- **Status:** Adopted where applicable.

### D-003 — Use Beautiful Soup for HTML table parsing
- **Decision:** Use `BeautifulSoup` to inspect table rows, cells, and headers.
- **Reason:** Preserving row/column relationships is necessary to map a rate to the correct metal and session.
- **Trade-off:** Source page layout changes can require parser maintenance.
- **Status:** Adopted for HTML parsing.

### D-004 — Use Playwright when browser behavior is required
- **Decision:** Use Playwright for sources where ordinary HTTP endpoints are blocked or do not expose the needed data.
- **Reason:** Browser execution can retrieve content unavailable through simple requests; this has been relevant to LBMA retrieval when direct JSON endpoints returned 403.
- **Trade-off:** Browser automation is heavier and requires browser/runtime setup.
- **Status:** Adopted where applicable.

### D-005 — Isolate source fetching and parsing
- **Decision:** Keep source-specific retrieval/parsing in separate modules: `lbma_fetcher.py`, `jm_fetcher.py`, `ibja_fetcher.py`, and `rbi_fetcher.py`.
- **Reason:** Providers differ in layout, date conventions, and failure modes.
- **Trade-off:** Some helper logic may be repeated; share helpers only when their contracts truly match.
- **Status:** Adopted.

### D-006 — Preserve AM and PM as separate observations
- **Decision:** Keep AM and PM values separate and do not infer one from the other.
- **Reason:** Missing values are not zero and a single rate may not identify a session.
- **Trade-off:** Missing sessions must remain representable and visible.
- **Implementation guidance:** Keep unavailable values as missing; never copy AM into PM or assign a session without source evidence.
- **Status:** Adopted.

### D-007 — Parse source tables structurally
- **Decision:** Match prices using table headers, row labels, and cell positions rather than flattened page text.
- **Reason:** Flattened text loses the relationship between labels, dates, sessions, and values.
- **Trade-off:** Structured parsing requires validation and can need changes when page markup changes.
- **Status:** Adopted as the preferred parsing method.

### D-008 — Validate data before writing it
- **Decision:** Validate required metals, numeric values, dates, units, and session mapping before workbook updates.
- **Reason:** A wrong rate distributed as valid is more harmful than a clearly reported missing rate.
- **Trade-off:** The workflow may stop or leave a supplementary section blank when data cannot be verified.
- **Status:** Required safeguard.

### D-009 — Use `openpyxl` for workbook updates
- **Decision:** Use `openpyxl` to read and update the `.xlsx` tracker.
- **Reason:** It enables direct worksheet/cell manipulation without automating the Excel UI for ordinary updates.
- **Trade-off:** Some Excel-specific features and formula recalculation behavior require testing with the actual template.
- **Status:** Adopted.

### D-010 — Use Python logging for diagnostics
- **Decision:** Use the standard-library `logging` module for lifecycle messages and errors.
- **Reason:** Logs support troubleshooting scheduled runs and source-specific failures.
- **Trade-off:** Logging configuration must avoid duplicate handlers and excessive console output.
- **Status:** Adopted.

### D-011 — Keep credentials out of source control
- **Decision:** Load SMTP credentials from environment/local configuration and exclude secrets from Git.
- **Reason:** Credentials must not be committed with the project.
- **Implementation guidance:** Keep `.env`, local config, passwords, tokens, and private endpoints out of the repository. Use placeholders in example configuration. Do not rely on a real password as a source-code fallback.
- **Trade-off:** Each runtime environment must be configured securely.
- **Status:** Required for repository publication and deployment.

### D-012 — Handle non-trading days and incomplete rows explicitly
- **Decision:** Respect provider-specific publication dates and incomplete data.
- **Reason:** Weekends, holidays, and partially populated current rows can yield missing or misleading rates.
- **Implementation guidance:** Validate dates and values, not just row position. The Johnson Matthey page has previously required care because its first/current row may be incomplete.
- **Status:** Adopted as a data-integrity requirement.

### D-013 — Keep the seven mandatory metal prices on a common effective date
- **Decision:** Resolve one common effective date for the mandatory set: Gold, Silver, Platinum, Palladium, Rhodium, Iridium, and Ruthenium.
- **Reason:** The main report and workbook date should not be later than any of the prices represented under that date.
- **Implementation detail:** `_effective_price_date(source_dates)` currently uses the earliest normalized source date.
- **Trade-off:** The effective date may be earlier than the run date or some individual source dates. AM fixings can have their own date and are written separately.
- **Status:** Adopted; retain this rule unless the business requirement changes.

### D-014 — Preserve actual IBJA publication dates
- **Decision:** Treat IBJA Silver 999 as an independent INR/kg reference series and write AM/PM sessions under their actual publication dates.
- **Reason:** IBJA is not the LBMA Silver USD/troy-ounce series, and its AM/PM publication dates may differ from the main metals date.
- **Email behavior:** Include available non-`None` IBJA rates in the email even if the dates differ from the main metals effective date; display the corresponding AM/PM dates to avoid implying the values share the metals date.
- **Data-integrity rule:** Do not invent a missing session, copy AM into PM, or assign a rate to a session without source evidence.
- **Trade-off:** The email can contain IBJA values from a different publication date, so separate date labels are necessary.
- **Status:** Adopted.

### D-015 — Store RBI exchange rates under their actual publication dates
- **Decision:** Write each RBI currency rate using the corresponding date in `fx_dates`, not automatically under the metals effective date.
- **Reason:** RBI may return the latest available date on or before the requested date, especially around weekends and holidays. The two currencies should use their own returned dates.
- **Implementation guidance:** Pair each `fx_prices` entry with its own `fx_dates` entry. Find the correct month sheet and date column; if the date cannot be found, log and skip rather than writing the rate under an incorrect date.
- **Trade-off:** RBI values can appear under a different date from the metals data in the workbook.
- **Status:** Adopted requirement; verify the current writer and caller implement this exact contract.

### D-016 — Calculate monthly average from valid observations
- **Decision:** Calculate the monthly average as the arithmetic mean of valid recorded daily rates for each metal, not as `(minimum + maximum) / 2`.
- **Reason:** The midpoint of the extremes does not represent the average of observations.
- **Implementation guidance:** Skip blank cells and count merged duplicate cells only once. Treat valid AM and PM observations as separate observations when both are recorded. Do not combine IBJA or RBI values with the seven metal USD/troy-ounce series.
- **Trade-off:** The average changes as new valid observations are added during the month.
- **Status:** Adopted requirement.

### D-017 — Keep the email presentation aligned with the data contract
- **Decision:** Preserve missing values as blank, center-align numeric values and dates in the monthly Min/Max table, and display IBJA AM/PM publication dates alongside those rates.
- **Reason:** The report should be readable without disguising missing values or suggesting that independently dated sources share the same date.
- **Trade-off:** The HTML builder needs to receive both rate values and their source dates.
- **Status:** Adopted presentation requirement; verify rendered email after edits.

### D-018 — Separate personal portfolio from organization deployment
- **Decision:** Use a sanitized personal repository for portfolio presentation and a company-approved repository/runtime for organizational operation.
- **Reason:** The workbook, BRD, logs, internal addresses, source configuration, and credentials may contain confidential business information.
- **Implementation guidance:** Exclude `.env`, real configuration secrets, live `.xlsx` files, backups, logs, internal BRD documents, and private data from public Git. Confirm company intellectual-property and publication approval before making the repository public.
- **Trade-off:** Some company-specific configuration must be documented with safe examples rather than published verbatim.
- **Status:** Required before public publication.

### D-019 — Use a scheduler only after a manual run is validated
- **Decision:** Prefer Windows Task Scheduler as the initial deployment option for the existing Windows-based project, subject to company approval.
- **Reason:** It can launch the existing Python entry point without converting the script into a web service.
- **Implementation guidance:** Run under an approved account, use the intended Python interpreter, ensure browser/network dependencies are available, keep the workbook unlocked, capture logs, and verify email and workbook results after scheduled runs.
- **Trade-off:** The host must be powered on and connected at run time; a dedicated company VM may be more reliable than a personal workstation.
- **Status:** Proposed deployment approach; not evidence that a scheduled task has already been installed.

## Change-entry template

### D-XXX — Short decision title
- **Date:** YYYY-MM-DD
- **Files changed:** `path/to/file.py`
- **Context/problem:** What prompted the change?
- **Decision:** What approach was selected?
- **Reason:** Why is it appropriate?
- **Alternatives considered:** What else was considered?
- **Libraries/APIs affected:** Relevant dependency or source.
- **Consequences/trade-offs:** What changes operationally?
- **Validation:** Tests or commands actually run.
- **Status:** Proposed / Adopted / Superseded / Reverted

## Documentation maintenance

1. Update `Flow.md` whenever call order, data contracts, date rules, workbook writes, or email inputs change.
2. Add a decision entry for material architectural or behavior changes.
3. Record only tests actually run; distinguish intended behavior from verified behavior.
4. Mark replaced decisions as **Superseded** and link to the replacement.
