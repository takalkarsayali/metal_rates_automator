# Engineering Decisions — Metal Rates Automator

> Purpose: record meaningful implementation decisions, the reason behind each choice, and the trade-offs.
>
> **Maintenance rule:** Add an entry whenever a code change materially affects architecture, data sources, parsing, validation, error handling, output files, configuration, security, or delivery. Do not log trivial formatting-only changes.

## Project context

Metal Rates Automator retrieves precious-metal prices from external sources, processes the rates, updates an Excel workbook, and distributes the workbook by email. The project has used LBMA data, Johnson Matthey data, and is being extended to fetch IBJA rates.

## Decision log

### D-001 — Use Python for the automation workflow

- **Decision:** Implement the workflow in Python.
- **Reason:** Python has suitable libraries for HTTP requests, HTML parsing, browser automation when needed, Excel workbook handling, and SMTP email delivery. It also keeps the workflow readable and maintainable as a scheduled or manually executed script.
- **Trade-off:** Runtime dependencies must be installed and managed in the target Windows environment.
- **Status:** Adopted.

### D-002 — Use `requests` for ordinary HTTP retrieval

- **Decision:** Use `requests` to retrieve pages or endpoints that can be accessed with a normal HTTP request.
- **Reason:** It provides a simple API for request headers, timeouts, status checks, and response content. This is sufficient when the source returns the required data directly.
- **Trade-off:** It cannot execute JavaScript or automatically reproduce a browser session. A site may block requests or return incomplete content.
- **Implementation guidance:** Set a timeout, use an appropriate User-Agent when needed, call `raise_for_status()`, and handle request exceptions. Never treat a successful HTTP status alone as proof that the extracted prices are correct.
- **Status:** Adopted where applicable.

### D-003 — Use Beautiful Soup for HTML table parsing

- **Decision:** Use `BeautifulSoup` (`bs4`) to parse HTML and inspect table rows and cells.
- **Reason:** It allows the scraper to preserve the table structure and extract headers and cell values, rather than flattening the page into one long string.
- **Trade-off:** Selectors and table layouts can change when the source website changes.
- **Implementation guidance:** Match values to explicit headers and row labels. Avoid assuming that a value next to a label is automatically an AM value or PM value. Validate the number and meaning of columns before returning a rate.
- **Status:** Adopted for IBJA HTML parsing.

### D-004 — Use Playwright when a source requires browser behavior

- **Decision:** Use Playwright for a source when ordinary HTTP retrieval is blocked or the required content is produced through browser-side behavior.
- **Reason:** A real browser context can execute JavaScript and access content that a simple HTTP request does not expose. This approach has been used for LBMA price retrieval where direct JSON endpoints returned HTTP 403.
- **Trade-off:** Browser automation is heavier, slower, and requires browser/runtime setup. It should not be the default for a simple static HTML page.
- **Implementation guidance:** Prefer a direct documented endpoint when it is accessible and permitted; otherwise keep browser automation isolated in the relevant fetcher and validate the captured payload.
- **Status:** Adopted for applicable sources.

### D-005 — Keep each rate source's fetching/parsing logic isolated

- **Decision:** Keep source-specific retrieval and parsing in separate modules (for example, LBMA, Johnson Matthey, and IBJA fetchers).
- **Reason:** Each provider has different page layouts, data formats, availability rules, and failure modes. Isolating those differences prevents one provider's parsing changes from unexpectedly breaking the others.
- **Trade-off:** Some small helper logic may be repeated. Shared helpers should only be introduced when they genuinely have the same contract and behavior.
- **Status:** Recommended architecture; confirm against the current repository.

### D-006 — Preserve AM and PM as separate values

- **Decision:** Represent AM and PM rates as separate fields and do not infer one from the other.
- **Reason:** A missing PM rate is not equivalent to a PM rate of zero, and a single displayed value may not identify which session it belongs to. Combining or guessing the sessions could put incorrect prices into the workbook or email.
- **Trade-off:** The pipeline must support a missing session and make its handling explicit.
- **Implementation guidance:** Use `None`/an explicit missing value when a session is unavailable; do not silently copy AM into PM or label an unverified value as either session. Confirm the source's actual headers and cell relationships before mapping prices.
- **Status:** Adopted as a data-integrity requirement.

### D-007 — Parse table rows and headers instead of relying on flattened text

- **Decision:** For tabular sources, inspect individual rows/cells and map values using headers and row labels.
- **Reason:** `table.get_text(" | ", strip=True)` is useful for debugging, but flattening loses structural relationships. A value such as `207098` appearing near “Silver 999” does not alone establish whether it is AM, PM, or a single current indicative value.
- **Trade-off:** Structured parsing needs more checks and may require updates if the website layout changes.
- **Status:** Adopted as the preferred parsing method.

### D-008 — Validate source data before updating the workbook

- **Decision:** Validate extracted rates, dates, required metals, and AM/PM mappings before writing output.
- **Reason:** Incorrect source mapping is more harmful than a failed run because a wrong price may be distributed as if it were valid.
- **Trade-off:** A run may stop or mark a value unavailable when the source layout changes.
- **Implementation guidance:** Check that required fields exist, values are parseable numeric rates where expected, dates are valid, and the selected session matches its header. Log a clear error when validation fails. Do not invent fallback values.
- **Status:** Required safeguard; verify current implementation.

### D-009 — Use `openpyxl` for Excel workbook updates

- **Decision:** Use `openpyxl` for direct edits to `.xlsx` workbooks when preserving and updating workbook cells is the main task.
- **Reason:** It supports reading and writing worksheets, cell values, and workbook structure without requiring Excel automation for common operations.
- **Trade-off:** Some advanced Excel features may not be recalculated or preserved exactly as Excel itself would handle them. The workbook should be tested with the actual template.
- **Implementation guidance:** Keep workbook paths/configuration explicit, validate target worksheets and cells, save to the intended output, and verify that the resulting file exists and is non-empty.
- **Status:** Adopted in the project context; confirm exact usage in the current code.

### D-010 — Use Python's built-in logging for operational diagnostics

- **Decision:** Use the standard-library `logging` module rather than scattered `print()` statements for operational messages.
- **Reason:** Logging supports severity levels, consistent formatting, and routing to the console or a file. This makes failures easier to diagnose during manual and scheduled runs.
- **Trade-off:** Logging configuration needs to be centralized to avoid duplicate or excessively verbose messages.
- **Implementation guidance:** Use `INFO` for important lifecycle events, `WARNING` for recoverable issues, and `ERROR`/`EXCEPTION` for failures. Avoid logging credentials or sensitive email content.
- **Status:** Adopted in the project context; tune verbosity to operational needs.

### D-011 — Read SMTP credentials from environment variables

- **Decision:** Read SMTP username/password from environment variables or a local environment configuration that is excluded from version control.
- **Reason:** Credentials should not be embedded in source code or committed to Git.
- **Trade-off:** The runtime environment must be configured correctly before email delivery can work.
- **Implementation guidance:** Do not use a real password as a source-code fallback. Validate required settings at startup, keep `.env` out of version control, and use a safe `.env.example` containing placeholders only.
- **Status:** Recommended security requirement.

### D-012 — Treat non-trading days and incomplete source data explicitly

- **Decision:** Resolve the appropriate available trading date according to each source's documented behavior rather than assuming that today's page always contains a complete rate.
- **Reason:** Weekends, holidays, delayed publication, and partially populated current-day tables can result in missing or incomplete prices.
- **Trade-off:** Date-selection logic is source-specific and needs tests against historical examples.
- **Implementation guidance:** Record the selected source date. For Johnson Matthey, the project's observed page behavior has required considering the second row when the first/current row may be incomplete; confirm the current page structure before relying on row position alone.
- **Status:** Source-specific; validate against the current implementation.

## Change-entry template

Copy this section for every meaningful decision:

### D-XXX — Short decision title

- **Date:** YYYY-MM-DD
- **Files changed:** `path/to/file.py`
- **Context/problem:** What prompted the change?
- **Decision:** What approach was selected?
- **Reason:** Why is this approach appropriate?
- **Alternatives considered:** What else could have been used, and why was it not chosen?
- **Libraries/APIs affected:** Name the dependency or external source and its role.
- **Consequences/trade-offs:** What becomes easier, harder, or newly constrained?
- **Validation:** What test, sample, or command verified the change?
- **Status:** Proposed / Adopted / Superseded / Reverted

## Documentation maintenance

1. Add a decision entry in the same change that introduces a meaningful architectural or behavioral change.
2. Explain the reason and trade-off, not just what changed.
3. Record the verification performed and any unverified assumptions.
4. If a decision is replaced, mark the old entry as **Superseded** and link to the new entry; do not silently erase useful history.
5. Keep `Flow.md` aligned with the actual code after changing function calls or execution order.
