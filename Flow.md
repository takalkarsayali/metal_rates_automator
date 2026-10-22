# Execution Flow — Metal Rates Automator

> This document explains how the project is intended to execute, how data moves between modules, and where to look when a step fails.
>
> **Accuracy note:** This guide is based on the project context currently available. Verify filenames, function names, and call order against the checked-out repository and update this document whenever the implementation differs. Do not treat illustrative function names below as confirmed names in every version.

## 1. High-level workflow

```text
Start script / scheduled run
          |
          v
Load configuration and initialize logging
          |
          v
Fetch rates from each configured source
  |             |                 |
  v             v                 v
LBMA fetcher   Johnson Matthey   IBJA fetcher
  |             |                 |
  +-------------+-----------------+
                |
                v
Normalize dates, units, and rate field names
                |
                v
Validate values and AM/PM session mapping
                |
                v
Resolve applicable trading date / missing values
                |
                v
Update the target Excel workbook
                |
                v
Save and verify output workbook
                |
                v
Send email with workbook attachment
                |
                v
Log success or report failure
```

The exact order may differ in the current entry-point script. Update the diagram to reflect the actual sequence in code.

## 2. Main stages

### Stage A — Startup and configuration

**Purpose:** Prepare dependencies, paths, source settings, logging, and email configuration.

Typical responsibilities:
- Read configuration such as workbook paths, output paths, log level, and source settings.
- Read SMTP credentials from environment variables.
- Initialize logging.
- Fail early with a clear message if required configuration is missing.

**Failure examples:** Missing environment variable, invalid workbook path, unavailable dependency.

### Stage B — Fetch data from each provider

Each source should have its own retrieval and parsing logic because the providers expose data differently.

#### LBMA

1. Request or load the required LBMA data.
2. If direct endpoints are blocked or do not expose the data, use the configured browser-based retrieval path.
3. Parse the metal series and select the expected price field and date.
4. Return a consistent result to the main workflow.
5. Raise or report a clear error if the source data cannot be obtained or parsed.

**Important:** Do not assume an HTTP response is valid price data merely because the request completed.

#### Johnson Matthey

1. Open or request the PGM price page.
2. Parse the relevant table and its date/session columns.
3. Extract the required metals, such as rhodium, iridium, and ruthenium, from the intended London 09:00 column.
4. Apply the project's trading-date rule only after verifying the row is complete.
5. Return the selected values and source date.

**Important:** The page may show an incomplete current row. Do not rely on “first row” or “second row” without validating the date, column labels, and values.

#### IBJA

1. Request `https://www.ibjarates.com/` using `requests`.
2. Check the HTTP response and parse its HTML with Beautiful Soup.
3. Iterate through table rows and cells rather than flattening an entire table into text.
4. Identify the relevant Silver 999 row and the actual AM/PM headers or the page's explicit current-rate structure.
5. Map each price to a session only when the HTML structure supports that mapping.
6. Convert numeric strings to the expected numeric type and represent unavailable sessions explicitly.
7. Return a normalized result or raise a descriptive parsing/validation error.

**Important IBJA finding:** A flattened result such as `Purity | AM | PM | ... | Silver 999 | 207098` does not prove that `207098` is the AM rate or the PM rate. The page may present a single current Silver 999 value separately from historical AM/PM tables. Confirm the actual row and column structure before assigning sessions. Never fabricate a missing PM value.

### Stage C — Normalize and validate data

Before updating the workbook, ensure that:
- Each value belongs to the expected metal and source.
- Dates are valid and associated with the correct record.
- AM and PM fields are not swapped or inferred.
- Numeric values are parseable and within reasonable, source-appropriate bounds.
- Missing values are handled according to an explicit policy.
- Units are consistent with the workbook's expectations.

Validation should stop the affected update or mark the value unavailable when the source layout is ambiguous. Do not silently substitute an unrelated or stale value.

### Stage D — Resolve the applicable trading date

The workflow may need to account for weekends, holidays, and delayed or incomplete publication.

Expected responsibilities:
1. Determine the date represented by each source record.
2. Apply the provider-specific rule for the latest complete available trading date.
3. Keep the chosen date alongside the prices.
4. Avoid mixing rates from different dates without an explicit business rule.
5. Log the source date selected.

### Stage E — Update the Excel workbook

If the implementation uses `openpyxl`, the workbook update generally follows this sequence:

1. Open the configured workbook/template.
2. Locate the target worksheet and cells/columns.
3. Write the validated date and metal rates.
4. Apply any required Min/Max or historical-row logic.
5. Save the workbook to the configured output path.
6. Confirm the output exists and is readable before attempting email delivery.

The actual worksheet names, cell coordinates, and Min/Max rules must be documented from the current code rather than guessed here.

### Stage F — Send the email

Typical sequence:
1. Load SMTP host, port, sender, recipient list, and credentials from configuration.
2. Build the email message.
3. Attach the saved workbook.
4. Connect and authenticate with the SMTP server.
5. Send the message.
6. Log delivery success or a useful error.

Do not log passwords or authentication tokens. An email step should not report success if workbook creation or attachment loading failed.

### Stage G — Error handling and logging

Use logs to identify the last completed stage:
- Startup/configuration succeeded.
- Source fetch started/completed.
- Source date and validation result.
- Workbook update/save completed.
- Email send completed or failed.

For failures, include the source or operation and a useful exception message. Avoid silently swallowing exceptions; if the overall run cannot produce a trustworthy workbook, exit with a failure status or use the project's explicit failure-notification mechanism.

## 3. Suggested function-level call map

Use this as a checklist while mapping the real code. Replace placeholder names with the exact functions in the repository.

```text
main() / entry point
  ├── load_config()
  ├── configure_logging()
  ├── fetch_lbma_prices()
  │     ├── request endpoint or launch browser
  │     ├── parse response
  │     └── validate LBMA result
  ├── fetch_johnson_matthey_prices()
  │     ├── request/load page
  │     ├── parse relevant table
  │     └── select complete dated row
  ├── fetch_ibja_silver999()
  │     ├── requests.get(...)
  │     ├── BeautifulSoup(response.text, "html.parser")
  │     ├── inspect table rows/cells and headers
  │     └── validate AM/PM mapping
  ├── normalize_and_validate_rates()
  ├── resolve_trading_date()
  ├── update_excel_workbook()
  ├── send_email_with_attachment()
  └── log final status
```

This is a conceptual map, not a claim that these exact functions already exist. If one function performs multiple responsibilities today, document that fact and consider separating responsibilities only when there is a clear maintenance benefit.

## 4. Library and tool responsibilities

| Library/tool | Intended responsibility | Why it fits |
|---|---|---|
| `requests` | HTTP requests to source pages/endpoints | Straightforward headers, timeouts, status checks, and response handling |
| `beautifulsoup4` (`bs4`) | Parse HTML tables and cells | Preserves table structure so labels and values can be matched |
| `playwright` | Browser-based source retrieval where necessary | Executes browser behavior for sources that cannot be retrieved reliably with ordinary HTTP |
| `openpyxl` | Read and update `.xlsx` workbooks | Direct cell/worksheet manipulation without requiring Excel UI automation for common operations |
| `smtplib` / `email` | Send email and build MIME messages | Python standard-library support for SMTP and attachments |
| `logging` | Runtime diagnostics | Consistent severity levels and console/file logging |
| `datetime` | Parse and compare dates | Standard-library date handling for trading-date logic |
| `os` / `pathlib` | Environment settings and paths | Standard-library configuration access and cross-path handling |

Only list a library as an active dependency in `requirements.txt` if the current code actually imports or uses it. `smtplib`, `email`, `datetime`, `os`, `pathlib`, and `logging` are part of Python's standard library and do not normally need to be installed separately.

## 5. Data contract recommendation

A normalized source result should make the date, metal, session, value, and source explicit. For example:

```python
{
    "source": "IBJA",
    "price_date": "YYYY-MM-DD",
    "metal": "Silver999",
    "unit": "INR/kg",
    "AM": None,   # Use a number only when the source explicitly identifies AM.
    "PM": None,   # Use a number only when the source explicitly identifies PM.
}
```

This is an illustrative contract, not an assertion that the current code returns this exact shape. Keep the project's existing interface unless you intentionally change it and update callers accordingly. For IBJA, a single current rate should be represented separately rather than forced into AM or PM without evidence.

## 6. Debugging guide

| Symptom | First place to inspect | What to verify |
|---|---|---|
| Wrong IBJA Silver 999 rate | IBJA fetcher | Table headers, cell positions, units, source date, and whether the value is a current rate or a historical AM/PM value |
| PM is `None` | IBJA parser/source page | Whether the page actually publishes PM yet; do not fill it with AM |
| `ValueError: too many values to unpack` | Function return value and caller | Compare the actual number/shape of returned values with the caller's unpacking pattern |
| LBMA returns 403 | LBMA retrieval function | Whether direct access is blocked and the browser-based path is configured correctly |
| Workbook is stale or empty | Workbook update stage | Selected date, target sheet/cells, save path, and save exceptions |
| Email has no attachment | Email stage | Output file exists, attachment path is correct, and MIME attachment was added |
| Run is too noisy | Logging configuration | Centralize log level/format and avoid duplicate handlers |

## 7. Keep this flow accurate

After a meaningful code change:
1. Update the affected stage and function names here.
2. Update the diagram if execution order or branching changed.
3. Document the architectural choice and reason in `decisions.md`.
4. Add or update a test for the changed behavior.
5. Run the relevant script or test and record what was actually verified.
