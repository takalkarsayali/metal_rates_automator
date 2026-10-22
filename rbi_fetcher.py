"""
Fetches the RBI Reference Rates (USD/INR and EUR/INR) from
https://www.rbi.org.in/scripts/ReferenceRateArchive.aspx

The page is an old ASP.NET WebForms form (postback-driven, custom calendar
widget for the dates, no stable element IDs), so this drives it with
Playwright and finds every field by its visible label rather than by ID:
  - "From Date" / "To Date" inputs  -> value set directly (no calendar clicking)
  - USD / EUR checkboxes            -> matched by the text next to each box
  - GO button                       -> then the result table is polled for

RBI publishes the reference rate once per working day (around midday IST),
so today's rate usually doesn't exist yet in the morning, and weekends /
RBI holidays have no row at all. We therefore search a window of days
ending at `requested_date` and take the MOST RECENT date that actually has
a rate, per currency.

If the page layout ever changes and something can't be found, the rendered
page is saved to logs/rbi_page_debug.html (+ .png) so it can be inspected.
"""

import logging
import os
import re
import time
from datetime import date, datetime, timedelta

logger = logging.getLogger(__name__)

RBI_URL = "https://www.rbi.org.in/scripts/ReferenceRateArchive.aspx"
RBI_DATE_FORMAT = "%d/%m/%Y"          # what the page's inputs show, e.g. 13/10/2026
CURRENCIES = ("USD", "EUR")

# Fallback only, if a checkbox's label text can't be read: the row is
# All, USD, GBP, EUR, JPY, AED, IDR.
_CHECKBOX_FALLBACK_INDEX = {"USD": 1, "EUR": 3}

DEBUG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")


# --------------------------------------------------------------------------
# Pure parsing helpers (no browser needed — unit-testable)
# --------------------------------------------------------------------------

def _parse_date(value: str):
    value = (value or "").strip()
    for fmt in ("%d/%m/%Y", "%d-%b-%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def _parse_rate(value: str):
    value = (value or "").strip().replace(",", "")
    if not value or value in {"-", "NA", "N/A"}:
        return None
    try:
        rate = float(value)
        return rate if rate > 0 else None
    except ValueError:
        return None


def _parse_tables(tables: list) -> dict:
    """
    tables: every <table> on the page, as list[rows], rows = list[cell text].

    Finds the results table by its header row (a "Date" column plus
    "USD (INR / 1 USD)" / "EUR (INR / 1 EUR)" columns) and returns
        {"USD": {date: rate, ...}, "EUR": {date: rate, ...}}
    """
    found = {cur: {} for cur in CURRENCIES}

    for rows in tables:
        header_idx = None
        cols = {}

        for i, row in enumerate(rows):
            headers = [re.sub(r"\s+", " ", cell).strip().upper() for cell in row]
            date_col = next((j for j, h in enumerate(headers) if h.startswith("DATE")), None)
            currency_cols = {
                cur: next((j for j, h in enumerate(headers) if cur in h), None)
                for cur in CURRENCIES
            }
            currency_cols = {k: v for k, v in currency_cols.items() if v is not None}

            if date_col is not None and currency_cols:
                header_idx = i
                cols = {"DATE": date_col, **currency_cols}
                break

        if header_idx is None:
            continue

        needed = max(cols.values())
        for row in rows[header_idx + 1:]:
            if len(row) <= needed:
                continue
            row_date = _parse_date(row[cols["DATE"]])
            if row_date is None:
                continue
            for cur in CURRENCIES:
                if cur in cols:
                    rate = _parse_rate(row[cols[cur]])
                    if rate is not None:
                        found[cur][row_date] = rate

    return found


def _latest_per_currency(found: dict, start_date: date, end_date: date):
    """Most recent date in [start_date, end_date] that has a rate, per currency."""
    prices, rate_dates = {}, {}
    for cur in CURRENCIES:
        in_range = [d for d in found.get(cur, {}) if start_date <= d <= end_date]
        if not in_range:
            raise RuntimeError(
                f"No RBI {cur}/INR rate found between {start_date} and {end_date}."
            )
        latest = max(in_range)
        prices[f"{cur}_INR"] = found[cur][latest]
        rate_dates[f"{cur}_INR"] = latest
    return prices, rate_dates


# --------------------------------------------------------------------------
# Browser helpers
# --------------------------------------------------------------------------

def _save_debug(page):
    try:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        html_path = os.path.join(DEBUG_DIR, "rbi_page_debug.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(page.content())
        page.screenshot(path=os.path.join(DEBUG_DIR, "rbi_page_debug.png"), full_page=True)
        logger.error("Saved RBI page snapshot to %s (+ .png) for inspection", html_path)
    except Exception:  # noqa: BLE001
        logger.exception("Could not save RBI debug snapshot")


def _find_date_input(page, label: str):
    """The text input belonging to the 'From Date' / 'To Date' label."""
    by_label = page.locator(
        f"xpath=(//*[text()[contains(normalize-space(.),'{label}')]]"
        f"//input[not(@type) or @type='text'] | "
        f"//*[text()[contains(normalize-space(.),'{label}')]]"
        f"/following::input[not(@type) or @type='text'])[1]"
    )
    if by_label.count() > 0:
        return by_label.first

    # Fallback: the two date boxes are the LAST two visible text inputs on
    # the page (the header's site-search box comes earlier in the DOM).
    visible = [i for i in page.locator("input[type='text']").all() if i.is_visible()]
    if len(visible) >= 2:
        return visible[-2] if label == "From Date" else visible[-1]

    raise RuntimeError(f"Could not locate the '{label}' input on the RBI page.")


def _set_input_value(locator, value: str):
    """Set a (possibly read-only, calendar-driven) input directly and fire
    the events the page's scripts listen for."""
    locator.evaluate(
        """(el, v) => {
            el.removeAttribute('readonly');
            el.value = v;
            el.dispatchEvent(new Event('input',  {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
        }""",
        value,
    )


def _tick_currency(page, code: str):
    """Tick the checkbox whose adjacent label text is `code` (e.g. 'USD')."""
    boxes = page.locator("input[type='checkbox']").all()

    target = None
    for box in boxes:
        label = box.evaluate(
            """el => {
                const l = el.labels && el.labels[0];
                if (l) return l.innerText.trim();
                let n = el.nextSibling;
                while (n && n.nodeType === 3 && !n.textContent.trim()) n = n.nextSibling;
                return n ? (n.textContent || '').trim() : '';
            }"""
        )
        if label.upper() == code:
            target = box
            break

    if target is None:
        idx = _CHECKBOX_FALLBACK_INDEX[code]
        if len(boxes) <= idx:
            raise RuntimeError(f"Could not locate the {code} checkbox on the RBI page.")
        logger.warning("Checkbox label for %s not readable; using position %d", code, idx)
        target = boxes[idx]

    if not target.is_checked():
        try:
            target.check(timeout=5000)
        except Exception:  # noqa: BLE001
            target.evaluate("el => el.click()")  # custom-styled boxes


def _click_go(page):
    import re as _re
    button = page.get_by_role("button", name=_re.compile(r"^\s*GO\s*$", _re.IGNORECASE))
    if button.count() == 0:
        button = page.locator("input[value='GO'], input[value='Go'], button:has-text('GO')")
    if button.count() == 0:
        raise RuntimeError("Could not locate the RBI 'GO' button.")
    button.first.click()


def _read_all_tables(page) -> list:
    return page.evaluate(
        """() => Array.from(document.querySelectorAll('table')).map(t =>
              Array.from(t.rows).map(r =>
                Array.from(r.cells).map(c => (c.innerText || '').trim())))"""
    )


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------

def fetch_rbi_exchange_rates(requested_date: date = None, lookback_days: int = 15):
    """
    Returns:
        (
            {"USD_INR": 95.7485, "EUR_INR": 107.3022},
            {"USD_INR": date(2026, 10, 12), "EUR_INR": date(2026, 10, 12)},
        )
    The rate dates are the most recent dates on or before `requested_date`
    that RBI has actually published (weekends/holidays/not-yet-published
    days are skipped automatically).

    Raises RuntimeError if the rates can't be retrieved.
    """
    from playwright.sync_api import sync_playwright

    requested_date = requested_date or date.today()
    start_date = requested_date - timedelta(days=lookback_days)

    logger.info("Fetching RBI exchange rates between %s and %s", start_date, requested_date)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_default_timeout(20000)

        try:
            page.goto(RBI_URL, wait_until="domcontentloaded", timeout=45000)
            page.get_by_text("From Date").first.wait_for(state="visible")

            _set_input_value(_find_date_input(page, "From Date"), start_date.strftime(RBI_DATE_FORMAT))
            _set_input_value(_find_date_input(page, "To Date"), requested_date.strftime(RBI_DATE_FORMAT))

            for code in CURRENCIES:
                _tick_currency(page, code)

            _click_go(page)

            # The results arrive via a postback (full reload or partial
            # update — either way), so poll rather than wait on one event.
            found = {cur: {} for cur in CURRENCIES}
            deadline = time.time() + 25
            while time.time() < deadline:
                try:
                    found = _parse_tables(_read_all_tables(page))
                except Exception:  # noqa: BLE001  (page mid-navigation)
                    found = {cur: {} for cur in CURRENCIES}
                if all(found[cur] for cur in CURRENCIES):
                    break
                time.sleep(0.5)

            prices, rate_dates = _latest_per_currency(found, start_date, requested_date)

        except Exception:
            _save_debug(page)
            raise
        finally:
            browser.close()

    for key in prices:
        logger.info("RBI %s: %.4f (date: %s)", key, prices[key], rate_dates[key])

    return prices, rate_dates


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    prices, rate_dates = fetch_rbi_exchange_rates()

    print("\nRBI EXCHANGE RATES")
    print("-" * 40)
    for currency, rate in prices.items():
        print(f"{currency}: {rate:.4f} | Date: {rate_dates[currency]:%d-%m-%Y}")