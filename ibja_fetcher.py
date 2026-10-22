"""
Fetches the IBJA (India Bullion and Jewellers Association) Silver 999 rate,
AM (opening) and PM (closing), from https://www.ibjarates.com/

How the page is laid out (the "Previous Dates Rate" section):

    [AM] [PM]                      <- two tab links:  #tab-am  /  #tab-pm
    pane #tab-am:  a table   |  | 999 | 995 | 916 | 750 | 585 | Silver 999 | Platinum 999 |
                             | 15/10/2026 | 145564 | ... | 207098 | 57143 |
                             | 14/10/2026 | ...
    pane #tab-pm:  an identical table with the PM rates

Gold is quoted per 10 gm; SILVER IS PER 1 KG (rupees) — so a figure like
207098 is Rs 207,098 per kg.

The AM and PM tables are told apart by the tab pane they sit in, not by
any column heading (there is no "AM"/"PM" column). If the page ever drops
those pane ids, the first table is taken as AM and the second as PM and a
warning is logged so it can be double-checked.

IBJA does not publish on weekends or central-government holidays, so those
dates simply have no row; for each session the latest date on or before the
requested date is used.
"""

import logging
import os
import re
from datetime import date, datetime, timedelta

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

IBJA_URL = "https://www.ibjarates.com/"
SESSIONS = ("AM", "PM")
SILVER_HEADER = "silver 999"

DEBUG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")


# --------------------------------------------------------------------------
# Parsing (pure functions — no network)
# --------------------------------------------------------------------------

def _parse_date(value: str):
    try:
        return datetime.strptime((value or "").strip(), "%d/%m/%Y").date()
    except ValueError:
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


def _cells(row) -> list:
    return [re.sub(r"\s+", " ", c.get_text(" ", strip=True)) for c in row.find_all(["th", "td"])]


def _has_silver_header(table) -> bool:
    return any(
        SILVER_HEADER in [c.lower() for c in _cells(row)]
        for row in table.find_all("tr")
    )


def _parse_rate_table(table) -> dict:
    """{date: silver_999_rate} from one AM-or-PM table."""
    silver_idx = None
    rates = {}

    for row in table.find_all("tr"):
        cells = _cells(row)
        lowered = [c.lower() for c in cells]

        # Header row: locate the Silver 999 column. (Data rows have the
        # date in column 0 and the header row has an empty first cell, so
        # the indexes line up.)
        if SILVER_HEADER in lowered:
            silver_idx = lowered.index(SILVER_HEADER)
            continue

        if silver_idx is None or len(cells) <= silver_idx:
            continue

        row_date = _parse_date(cells[0])
        rate = _parse_rate(cells[silver_idx])

        if row_date is not None and rate is not None:
            rates[row_date] = rate

    return rates


def _locate_session_tables(soup):
    """
    ({"AM": table, "PM": table}, how) — `how` is "tab panes" or "position".
    Raises ValueError if the two tables can't be found.
    """
    by_pane = {}
    for session in SESSIONS:
        pane = soup.find(id=f"tab-{session.lower()}")
        table = pane.find("table") if pane else None
        if table is not None and _has_silver_header(table):
            by_pane[session] = table

    if len(by_pane) == len(SESSIONS):
        return by_pane, "tab panes"

    # Fallback: the innermost tables that carry a Silver 999 column, in
    # page order. (Outer layout tables that merely CONTAIN these are skipped
    # — matching them was what produced the one giant row in the first
    # version of this script.)
    candidates = [
        t for t in soup.find_all("table")
        if t.find("table") is None and _has_silver_header(t)
    ]

    if len(candidates) == 2:
        logger.warning(
            "IBJA tab panes #tab-am / #tab-pm not found; taking the first table "
            "as AM and the second as PM. Please verify against the website."
        )
        return {"AM": candidates[0], "PM": candidates[1]}, "position"

    raise ValueError(
        f"Expected the AM and PM Silver 999 tables; found {len(candidates)} candidate table(s) "
        "and no #tab-am / #tab-pm panes."
    )


def parse_ibja_silver999(html: str) -> dict:
    """
    {"AM": {date: rate, ...}, "PM": {date: rate, ...}, "_how": "tab panes"|"position"}
    """
    soup = BeautifulSoup(html, "html.parser")
    tables, how = _locate_session_tables(soup)
    parsed = {session: _parse_rate_table(tables[session]) for session in SESSIONS}
    parsed["_how"] = how
    return parsed


def _latest_on_or_before(rates: dict, end_date: date, start_date: date):
    in_range = [d for d in rates if start_date <= d <= end_date]
    if not in_range:
        return None, None
    latest = max(in_range)
    return rates[latest], latest


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------

def _save_debug(html: str):
    try:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        path = os.path.join(DEBUG_DIR, "ibja_page_debug.html")
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        logger.error("Saved IBJA page to %s for inspection", path)
    except Exception:  # noqa: BLE001
        logger.exception("Could not save IBJA debug page")


def fetch_ibja_silver999(requested_date: date = None, lookback_days: int = 10):
    """
    Returns:
        (
            {"SILVER999_AM": 207098.0, "SILVER999_PM": 208275.0},      # Rs per kg
            {"SILVER999_AM": date(2026, 10, 15), "SILVER999_PM": date(2026, 10, 15)},
        )

    Each session gets the latest rate on or before `requested_date`, with
    its own date (PM can lag AM early in the day, and weekends / holidays
    have no row). Raises RuntimeError if either session has nothing in the
    lookback window.
    """
    requested_date = requested_date or date.today()
    start_date = requested_date - timedelta(days=lookback_days)

    logger.info("Fetching IBJA Silver 999 rates for %s", requested_date)

    response = requests.get(IBJA_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    response.raise_for_status()
    html = response.text

    try:
        parsed = parse_ibja_silver999(html)

        prices, rate_dates = {}, {}
        for session in SESSIONS:
            rate, rate_date = _latest_on_or_before(parsed[session], requested_date, start_date)
            if rate is None:
                raise RuntimeError(
                    f"No IBJA Silver 999 {session} rate found between {start_date} and {requested_date}."
                )
            prices[f"SILVER999_{session}"] = rate
            rate_dates[f"SILVER999_{session}"] = rate_date
    except Exception:
        _save_debug(html)
        raise

    logger.info("IBJA tables identified via %s", parsed["_how"])
    for key in prices:
        logger.info("IBJA %s: %.0f (date: %s)", key, prices[key], rate_dates[key])

    return prices, rate_dates


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    prices, rate_dates = fetch_ibja_silver999()

    print("\nIBJA SILVER 999 (Rs per kg)")
    print("-" * 40)
    for key, rate in prices.items():
        print(f"{key}: {rate:,.0f} | Date: {rate_dates[key]:%d-%m-%Y}")