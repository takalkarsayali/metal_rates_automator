import inspect
import logging
import os
import sys
from datetime import date, datetime, timedelta
import config
from lbma_fetcher import fetch_lbma_fixings
from jm_fetcher import fetch_jm_prices
from rbi_fetcher import fetch_rbi_exchange_rates
from ibja_fetcher import fetch_ibja_silver999
from excel_updater import update_excel, _month_sheet_name
import openpyxl
from openpyxl.cell.cell import MergedCell
from emailer import build_html, send_email

LBMA_METALS = {"GOLD", "SILVER", "PLATINUM", "PALLADIUM"}
JM_METALS = {"RHODIUM", "IRIDIUM", "RUTHENIUM"}
AM_METALS = ["GOLD", "PLATINUM", "PALLADIUM"]      # silver has one daily fixing, no AM
METAL_DISPLAY_ORDER = ["GOLD", "SILVER", "PLATINUM", "PALLADIUM", "RHODIUM", "IRIDIUM", "RUTHENIUM"]

# rbi_fetcher key -> (key under config.LAYOUT["Exchange_Rate"], display label)
FX_ROW_MAP = {
    "USD_INR": ("INR_USD_row", "USD/INR"),
    "EUR_INR": ("INR_EUR_row", "EUR/INR"),
}
FX_NUMBER_FORMAT = "0.0000"   # RBI publishes 4 decimal places


class _QuietConsoleFormatter(logging.Formatter):
    """One-line console messages: never append the traceback (the log file
    keeps it). Needed because the file handler caches the traceback on the
    record, which would otherwise be appended here as well."""

    def format(self, record):
        saved = (record.exc_info, record.exc_text)
        record.exc_info = record.exc_text = None
        try:
            return super().format(record)
        finally:
            record.exc_info, record.exc_text = saved


def setup_logging():
    """Write run details to file; show errors only in the console.
    The final run summary is printed separately."""
    os.makedirs(config.LOG_DIR, exist_ok=True)
    log_file = os.path.join(config.LOG_DIR, f"run_{date.today():%Y-%m-%d}.log")

    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.ERROR)
    console_handler.setFormatter(_QuietConsoleFormatter("[%(levelname)s] %(message)s"))

    logging.basicConfig(level=logging.INFO, handlers=[file_handler, console_handler])


def read_minmax_for_email(target_date: date) -> dict:
    """Re-reads the just-saved sheet so the email reflects exactly what's
    now in the workbook (rather than recomputing separately)."""
    wb = openpyxl.load_workbook(config.EXCEL_PATH, data_only=True)
    sheet_name = _month_sheet_name(target_date)
    ws = wb[sheet_name]
    layout = config.LAYOUT

    result = {}
    for metal, row in layout["metal_minmax_rows"].items():
        min_val = ws.cell(row=row, column=layout["minmax_min_val_col"]).value
        min_date = ws.cell(row=row, column=layout["minmax_min_date_col"]).value
        max_val = ws.cell(row=row, column=layout["minmax_max_val_col"]).value
        max_date = ws.cell(row=row, column=layout["minmax_max_date_col"]).value
        if min_val is not None:
            result[metal] = {
                "min": min_val, "min_date": min_date,
                "max": max_val, "max_date": max_date,
            }
    return result


def _normalize_date(value) -> date:
    """Normalizes LBMA's string dates ('2026-10-05') and JM's date objects
    into a plain datetime.date, for comparison and display alike."""
    if isinstance(value, str):
        return datetime.strptime(value, "%Y-%m-%d").date()
    return value


def _fmt_source_date(value) -> str:
    try:
        return _normalize_date(value).strftime("%d-%m-%Y")
    except (ValueError, AttributeError):
        return str(value)


def _effective_price_date(source_dates: dict) -> date:
    """
    The Excel column (and email) should be labeled with the date the rates
    actually ARE for, not the calendar day the script happened to run on —
    e.g. if the script runs Tuesday morning before LBMA's Tuesday PM fix is
    published, the rates it gets are still Monday's, so the column should
    say Monday, not Tuesday.

    LBMA and JM can each independently fall back to their own most recent
    available date, and those two dates can differ (as they did here: LBMA
    fell back to Monday while JM already had Tuesday's). Using the OLDEST
    of all 7 metals' actual dates means the column label never overstates
    freshness — every value in that column is for that date or earlier,
    never later.
    """
    normalized = [_normalize_date(v) for v in source_dates.values()]
    return min(normalized)


def _find_column_for_date(ws, price_date: date):
    """Column (1-indexed) of the Date-row cell equal to price_date, or None.

    Each day now spans an AM/PM column pair and the header cells may be
    merged across it, so MergedCell placeholders (always empty) are stepped
    over instead of being mistaken for the end of the data. A day's date is
    found at its first (left-most) column. Days are filled left to right,
    so the first genuinely empty Date cell ends the search."""
    layout = config.LAYOUT
    col = layout["first_data_col"]
    last_possible = col + 90           # 31 days x 2 columns, with headroom
    while col <= last_possible:
        cell = ws.cell(row=layout["date_row"], column=col)
        if isinstance(cell, MergedCell):
            col += 1
            continue
        if cell.value is None:
            return None
        cell_date = cell.value.date() if hasattr(cell.value, "date") else cell.value
        if cell_date == price_date:
            return col
        col += 1
    return None


def _writable_cell(ws, row: int, col: int):
    """The cell to write to for (row, col). openpyxl won't accept a value
    in a merged range's non-anchor cells, so redirect to the range's
    top-left cell."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return ws.cell(row=rng.min_row, column=rng.min_col)
    return cell


def write_exchange_rates(fx_prices: dict, fx_dates: dict) -> list:
    """Write each RBI rate under its actual publication date."""
    logger = logging.getLogger(__name__)
    layout_rows = config.LAYOUT.get("Exchange_Rate", {})

    if not layout_rows:
        logger.warning("No Exchange_Rate layout configured.")
        return []

    wb = openpyxl.load_workbook(config.EXCEL_PATH)
    written = []

    for key, (row_key, _label) in FX_ROW_MAP.items():
        if key not in fx_prices:
            continue

        rate_date_value = fx_dates.get(key)
        if not rate_date_value:
            logger.warning("No publication date for RBI rate %s; skipping.", key)
            continue

        rate_date = _normalize_date(rate_date_value)
        sheet_name = _month_sheet_name(rate_date)

        if sheet_name not in wb.sheetnames:
            logger.warning(
                "RBI date %s belongs to missing sheet %s; skipping %s.",
                rate_date, sheet_name, key,
            )
            continue

        row = layout_rows.get(row_key)
        if row is None:
            logger.warning("No Excel row configured for %s.", key)
            continue

        ws = wb[sheet_name]
        col = _find_column_for_date(ws, rate_date)

        if col is None:
            logger.warning(
                "Could not find date %s in sheet %s for RBI rate %s.",
                rate_date, sheet_name, key,
            )
            continue

        cell = _writable_cell(ws, row, col)
        cell.value = fx_prices[key]
        cell.number_format = FX_NUMBER_FORMAT
        written.append(key)

        logger.info(
            "RBI %s = %s written under %s, sheet %s, cell %s",
            key, fx_prices[key], rate_date, sheet_name, cell.coordinate,
        )

    wb.save(config.EXCEL_PATH)
    return written


def _update_excel_for_session(session_prices: dict, for_date: date, session: str) -> bool:
    """
    Call update_excel() for one fixing session ("AM" or "PM") without
    assuming its exact signature: the date goes to whichever parameter
    follows `prices`, and `session=` is passed if (and only if) the
    function accepts it. An older update_excel with no session support can
    still take PM, but cannot place AM data, so AM is skipped (returns False).
    """
    params = list(inspect.signature(update_excel).parameters)
    kwargs = {}

    date_param = next((p for p in params[1:] if p != "session"), None)
    if date_param:
        kwargs[date_param] = for_date

    if "session" in params:
        kwargs["session"] = session
    elif session != "PM":
        logging.getLogger(__name__).info(
            "update_excel() has no 'session' parameter; %s data not written.", session
        )
        return False

    # IBJA's key is not in config.LAYOUT["metal_rows"]. Pass it through
    # update_excel's dedicated ibja_rate parameter instead of as a metal.
    ibja_rate = session_prices.get("SILVER999")
    regular_prices = {
        metal: value for metal, value in session_prices.items()
        if metal != "SILVER999"
    }
    if ibja_rate is not None:
        if "ibja_rate" not in params:
            raise RuntimeError(
                "IBJA Silver 999 was fetched, but update_excel() does not support ibja_rate."
            )
        kwargs["ibja_rate"] = ibja_rate

    update_excel(regular_prices, **kwargs)
    return True


def print_run_summary(target_date, prices, source_dates, excel_path, sheet_name, recipients,
                      fx_prices=None, fx_dates=None, am_prices=None, am_dates=None,
                      ibja_prices=None, ibja_dates=None):
    """A single, quiet, complete summary of the run — printed directly
    (not logged) so it always shows regardless of console log level."""
    rows = []
    for metal in METAL_DISPLAY_ORDER:
        if metal not in prices:
            continue
        source = "LBMA" if metal in LBMA_METALS else "Johnson Matthey"
        rows.append((
            metal.title(),
            f"{prices[metal]:,.2f}" if metal != "SILVER" else f"{prices[metal]:,.3f}",
            _fmt_source_date(source_dates.get(metal, "")),
            source,
        ))

    col_widths = [10, 16, 12, 16]
    for r in rows:
        for i, cell in enumerate(r):
            col_widths[i] = max(col_widths[i], len(cell))

    def fmt_row(cells):
        return "  ".join(c.ljust(w) for c, w in zip(cells, col_widths))

    line = "=" * (sum(col_widths) + 2 * (len(col_widths) - 1))

    print()
    print(line)
    print(f" METAL RATES AUTOMATOR — RUN SUMMARY ({target_date:%d-%m-%Y})")
    print(line)
    print(fmt_row(["Metal (PM)", "Price ($/Toz)", "Price Date", "Source"]))
    print("-" * len(line))
    for r in rows:
        print(fmt_row(list(r)))

    # LBMA AM fixing (Gold / Platinum / Palladium), if we have it
    if am_prices:
        print("-" * len(line))
        print(fmt_row(["Metal (AM)", "Price ($/Toz)", "Price Date", "Source"]))
        for metal in METAL_DISPLAY_ORDER:
            if metal in am_prices:
                print(fmt_row([
                    metal.title(),
                    f"{am_prices[metal]:,.2f}",
                    _fmt_source_date((am_dates or {}).get(metal, "")),
                    "LBMA",
                ]))

    # Exchange rates (INR per 1 unit of foreign currency), if we have them
    if fx_prices:
        print("-" * len(line))
        fx_rows = []
        for key, (_row_key, label) in FX_ROW_MAP.items():
            if key in fx_prices:
                fx_rows.append((
                    label,
                    f"{fx_prices[key]:,.4f}",
                    _fmt_source_date((fx_dates or {}).get(key, "")),
                    "RBI",
                ))
        print(fmt_row(["Exchange", "INR per unit", "Rate Date", "Source"]))
        for r in fx_rows:
            print(fmt_row(list(r)))

    # IBJA Silver 999 is a separate INR/kg benchmark, not the LBMA USD/toz silver rate.
    if ibja_prices:
        print("-" * len(line))
        print(fmt_row(["IBJA Silver 999", "INR per kg", "Rate Date", "Source"]))
        for session in ("AM", "PM"):
            key = f"SILVER999_{session}"
            if key in ibja_prices:
                print(fmt_row([
                    f"Silver 999 ({session})",
                    f"{ibja_prices[key]:,.0f}",
                    _fmt_source_date((ibja_dates or {}).get(key, "")),
                    "IBJA",
                ]))

    print(line)
    print(f" Excel updated : {os.path.basename(excel_path)}  (sheet: {sheet_name})")
    print(f" Email sent to : {', '.join(recipients)}")
    print(line)
    print()


def _build_email_inputs(prices, am_prices, am_dates, ibja_prices, ibja_dates, fx_prices, price_date):

    logger = logging.getLogger(__name__)

    email_prices = {}
    for metal, pm_value in prices.items():
        if metal in AM_METALS:
            am_value = None
            if metal in am_prices:
                if _normalize_date(am_dates[metal]) == price_date:
                    am_value = am_prices[metal]
                else:
                    logger.info("Email: %s AM (%s) is not for %s — left blank in the email.",
                                metal, am_dates[metal], price_date)
            email_prices[metal] = {"AM": am_value, "PM": pm_value}
        else:
            email_prices[metal] = pm_value

    email_ibja = {}
    for key, rate in ibja_prices.items():
        if rate is not None:
            email_ibja[key] = rate
            logger.info(
                "Email: including IBJA %s = %s (publication date: %s)",
                key,
                rate,
                ibja_dates.get(key),
            )
        else:
            logger.info("Email: IBJA %s (%s) is not for %s — left out of the email.",
                        key, ibja_dates[key], price_date)

    # "USD_INR" -> "USD" (the email prints it as "INR / 1 USD")
    email_fx = {key.split("_")[0]: rate for key, rate in fx_prices.items()}

    return email_prices, email_ibja, email_fx

def _num(value):
    return value if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def _read_slot(ws, row, col, slot):
    """slot 0 = AM column, 1 = PM column, None = single fixing (left cell,
    falling back to the right cell if the value sits there)."""
    if slot is None:
        v = _num(ws.cell(row=row, column=col).value)
        return v if v is not None else _num(ws.cell(row=row, column=col + 1).value)
    return _num(ws.cell(row=row, column=col + slot).value)


def _scan_back(ws, row, start_col, slot):
    first = config.LAYOUT["first_data_col"]
    col = start_col
    while col >= first:
        v = _read_slot(ws, row, col, slot)
        if v is not None:
            return v
        col -= 2                      # one day = AM/PM column pair
    return None


def _previous_value(wb, row, ref_date, slot):
    """Most recent recorded value BEFORE ref_date's column (skips weekends /
    holidays automatically). Falls back to last month's sheet."""
    ref_date = _normalize_date(ref_date)
    first = config.LAYOUT["first_data_col"]

    sheet = _month_sheet_name(ref_date)
    if sheet in wb.sheetnames:
        ws = wb[sheet]
        col = _find_column_for_date(ws, ref_date)
        if col is not None:
            v = _scan_back(ws, row, col - 2, slot)
            if v is not None:
                return v

    prev_month = ref_date.replace(day=1) - timedelta(days=1)
    prev_sheet = _month_sheet_name(prev_month)
    if prev_sheet in wb.sheetnames:
        ws = wb[prev_sheet]
        last = first + 2 * ((ws.max_column - first) // 2)
        return _scan_back(ws, row, last, slot)
    return None

IBJA_ROW_KEYS = ("SILVER999", "SILVER_999", "IBJA_SILVER999",
                 "ibja_row", "ibja_silver999_row", "silver999_row")


def _find_ibja_row(wb, ref_date):
    """Row of the IBJA Silver 999 data in the month sheet.
    1) known config keys, 2) fall back to the row label containing '999'."""
    layout = config.LAYOUT

    for container in (layout.get("metal_rows", {}), layout):
        for key in IBJA_ROW_KEYS:
            row = container.get(key)
            if isinstance(row, int):
                return row

    sheet = _month_sheet_name(_normalize_date(ref_date))
    if sheet in wb.sheetnames:
        ws = wb[sheet]
        for r in range(1, ws.max_row + 1):
            for c in range(1, layout["first_data_col"]):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and "999" in v:
                    return r
    return None


def _build_previous_rates(price_date, prices, ibja_prices, ibja_dates, fx_prices, fx_dates):
    """Previous recorded rates, read from Excel, used ONLY for email arrows."""
    logger = logging.getLogger(__name__)
    previous_prices, previous_ibja, previous_fx = {}, {}, {}
    try:
        wb = openpyxl.load_workbook(config.EXCEL_PATH, data_only=True)
        metal_rows = config.LAYOUT["metal_rows"]

        for metal in prices:
            row = metal_rows.get(metal)
            if row is None:
                continue
            if metal in AM_METALS:
                previous_prices[metal] = {
                    "AM": _previous_value(wb, row, price_date, 0),
                    "PM": _previous_value(wb, row, price_date, 1),
                }
            else:
                previous_prices[metal] = _previous_value(wb, row, price_date, None)

        ibja_dates_ref = [d for d in ibja_dates.values() if d]
        if ibja_dates_ref:
            ibja_row = _find_ibja_row(wb, ibja_dates_ref[0])
            if ibja_row is None:
                logger.warning("IBJA Silver 999 row not found; no email arrows for it.")
            else:
                logger.info("IBJA previous-rate lookup using Excel row %s", ibja_row)
                for session, slot in (("AM", 0), ("PM", 1)):
                    key = f"SILVER999_{session}"
                    if key in ibja_prices and ibja_dates.get(key):
                        previous_ibja[session] = _previous_value(
                            wb, ibja_row, ibja_dates[key], slot)

        fx_rows = config.LAYOUT.get("Exchange_Rate", {})
        for key, (row_key, _label) in FX_ROW_MAP.items():
            row = fx_rows.get(row_key)
            if key in fx_prices and row is not None and fx_dates.get(key):
                previous_fx[key.split("_")[0]] = _previous_value(
                    wb, row, fx_dates[key], None)
    except Exception:
        # warning level -> log file only, console stays untouched
        logger.warning("Could not read previous rates for email arrows.", exc_info=True)
    return previous_prices, previous_ibja, previous_fx

def run(target_date: date = None):
    target_date = target_date or date.today()
    logger = logging.getLogger(__name__)
    logger.info("=== Metal Rates Automator run: %s ===", target_date)

    if target_date.weekday() >= 5:
        print(f"Weekend ({target_date:%A}) — LBMA/JM don't publish prices, skipping.")
        return

    prices = {}
    source_dates = {}

    # One browser session returns both fixings, each price with its own date.
    # PM (+ silver's daily fixing) is required; AM is best-effort.
    try:
        fixings = fetch_lbma_fixings(target_date)
    except Exception:
        logger.exception("Failed to fetch LBMA prices — aborting this run.")
        print("ERROR: Failed to fetch LBMA prices — see log file for details. Run aborted.")
        return
    prices.update(fixings["PM"]["prices"])
    source_dates.update(fixings["PM"]["dates"])

    am_prices = dict(fixings["AM"]["prices"])
    am_dates = dict(fixings["AM"]["dates"])
    missing_am = [m for m in AM_METALS if m not in am_prices]
    if missing_am:
        logger.error("LBMA AM price missing for: %s", missing_am)
        # Keep the console quiet; missing AM fixings are recorded in the log file.

    try:
        jm_prices, jm_price_dates = fetch_jm_prices(target_date)
    except Exception:
        logger.exception("Failed to fetch JM prices — aborting this run.")
        print("ERROR: Failed to fetch JM prices — see log file for details. Run aborted.")
        return
    prices.update(jm_prices)
    source_dates.update(jm_price_dates)

    if len(prices) != 7:
        logger.error("Only got %d/7 metal prices (%s) — aborting to avoid a partial row.",
                      len(prices), list(prices.keys()))
        print(f"ERROR: Only got {len(prices)}/7 metal prices — aborting to avoid a partial row.")
        return

    # Label the column/email with the date the rates are actually FOR, not
    # necessarily today — see _effective_price_date's docstring.
    price_date = _effective_price_date(source_dates)
    if price_date != target_date:
        logger.info("Run date %s but effective price date is %s — recording under %s.",
                     target_date, price_date, price_date)

    # Exchange rates are looked up AS OF the metals' price date so both sit
    # under the same column date; RBI skips weekends/holidays on its own, so
    # if it has no rate for that exact day it returns the latest earlier one
    # (the rate date is shown in the summary). These are supplementary: if
    # RBI can't be reached the metals run still completes, with the exchange
    # rows left blank.
    fx_prices, fx_dates = {}, {}
    try:
        fx_prices, fx_dates = fetch_rbi_exchange_rates(price_date)
    except Exception:
        logger.exception("Failed to fetch RBI exchange rates — continuing without them.")
        # Optional source failure: details are retained in the log file.

    # Fetch IBJA Silver 999 separately: it is quoted in INR/kg and has its
    # own AM and PM tables/dates, so never mix it into LBMA SILVER (USD/toz).
    ibja_prices, ibja_dates = {}, {}
    try:
        ibja_prices, ibja_dates = fetch_ibja_silver999(target_date)
    except Exception:
        logger.exception("Failed to fetch IBJA Silver 999 rates — continuing without IBJA.")
        # Optional source failure: details are retained in the log file.

    # Write each fixing under ITS OWN date. AM is set earlier in the London
    # day than PM, so on an early run AM can be a trading day ahead of PM
    # (AM = today, PM = yesterday). Oldest date first keeps the columns in
    # chronological order; on the same date AM goes before PM.
    sessions = [("PM", price_date, prices)]
    if am_prices:
        am_date = _effective_price_date(am_dates)
        sessions.append(("AM", am_date, am_prices))
    # IBJA uses its own effective date for each session. The Excel layout
    # must define a SILVER999 row/key for update_excel() to write these values.
    for session in ("AM", "PM"):
        key = f"SILVER999_{session}"
        if key in ibja_prices:
            sessions.append((session, ibja_dates[key], {"SILVER999": ibja_prices[key]}))
    sessions.sort(key=lambda item: (item[1], 0 if item[0] == "AM" else 1))

    for session, session_date, session_prices in sessions:
        try:
            if _update_excel_for_session(session_prices, session_date, session):
                logger.info("Excel updated: %s fixing for %s", session, session_date)
        except Exception as exc:
            logger.exception("Excel update failed for the %s fixing.", session)
            print(f"ERROR: Could not update the Excel sheet ({session} fixing): {exc}")
            print("       The email was NOT sent. Fix the problem above and rerun.")
            return

    if fx_prices:
        try:
            written = write_exchange_rates(fx_prices, fx_dates)
            logger.info("Exchange rates written to Excel: %s", written)
        except Exception:
            logger.exception("Failed to write exchange rates to Excel.")
            # Optional source failure: details are retained in the log file.
            fx_prices, fx_dates = {}, {}

    sheet_name = _month_sheet_name(price_date)

    minmax = read_minmax_for_email(price_date)

    previous_prices, previous_ibja, previous_fx = _build_previous_rates(
        price_date, prices, ibja_prices, ibja_dates, fx_prices, fx_dates
    )
    # The email is dated with price_date (same as the Excel column), not the
    # run date, and gets its inputs in the shapes emailer.build_html() expects.
    email_prices, email_ibja, email_fx = _build_email_inputs(
        prices, am_prices, am_dates, ibja_prices, ibja_dates, fx_prices, price_date
    )

    html_body = build_html(
        prices=email_prices,
        minmax=minmax,
        target_date=price_date,
        ibja_rates=email_ibja,
        exchange_rates=email_fx,
        ibja_dates=ibja_dates,
        previous_prices=previous_prices,
        previous_ibja_rates=previous_ibja,
        previous_exchange_rates=previous_fx,
    )

    send_email(
        html_body=html_body,
        target_date=price_date,
        attachment_path=config.EXCEL_PATH,
    )

    recipients = list(config.EMAIL_TO) + list(config.EMAIL_CC)
    print_run_summary(price_date, prices, source_dates, config.EXCEL_PATH, sheet_name, recipients,
                      fx_prices=fx_prices, fx_dates=fx_dates,
                      am_prices=am_prices, am_dates=am_dates,
                      ibja_prices=ibja_prices, ibja_dates=ibja_dates)

    logger.info("=== Run complete ===")


if __name__ == "__main__":
    setup_logging()
    run()