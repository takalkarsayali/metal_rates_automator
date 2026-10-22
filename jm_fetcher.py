import logging
import re
from datetime import date
import requests
import config

logger = logging.getLogger(__name__)

JM_ENDPOINT = (
    "https://matthey.com/products-and-markets/pgms-and-circularity/pgm-management"
    "?p_p_id=jm_metal_price_table_portlet_MetalPriceTablePortlet"
    "&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view&p_p_cacheability=cacheLevelPage"
)

FORM_RADIO_KEY = "_jm_metal_price_table_portlet_MetalPriceTablePortlet_priceRadio"
FORM_DATE_KEY = "_jm_metal_price_table_portlet_MetalPriceTablePortlet_pricetableDate"

# Metal -> the "priceRadio" code JM's form expects
JM_METAL_CODES = {
    "RHODIUM": "Rh",
    "IRIDIUM": "Ir",
    "RUTHENIUM": "Ru",
}

LONDON_PRICE_FIELD = "Price2"  # confirmed against the live site's display

_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _month_start_param(d: date) -> str:
    """JM's expected pricetableDate format: day fixed at 01, month with NO
    leading zero, e.g. "01-9-2026" for September 2026."""
    return f"01-{d.month}-{d.year}"


def _prev_month(d: date) -> date:
    if d.month == 1:
        return date(d.year - 1, 12, 1)
    return date(d.year, d.month - 1, 1)


def _parse_interval_date(s: str):
    """'Tue, 01 Sep 2026' -> date(2026, 9, 1). Returns None if unparseable."""
    m = re.search(r"(\d{1,2})\s+([A-Za-z]{3})[a-z]*\s+(\d{4})", s)
    if not m:
        return None
    day, mon_str, year = m.groups()
    month = _MONTHS.get(mon_str.lower())
    if not month:
        return None
    return date(int(year), month, int(day))


def _fetch_month_columns(metal_code: str, month_start: date) -> list:
    """Returns the raw `tableList.columns` list (oldest first) for the
    given metal and month."""
    resp = requests.post(
        JM_ENDPOINT,
        data={
            FORM_RADIO_KEY: metal_code,
            FORM_DATE_KEY: _month_start_param(month_start),
        },
        headers={
            "User-Agent": "Mozilla/5.0",
            "X-Requested-With": "XMLHttpRequest",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        timeout=20,
    )
    resp.raise_for_status()
    data = resp.json()
    return data.get("tableList", {}).get("columns", [])


# def _previous_trading_day_london_price(metal_code: str, run_date: date):
#     """Returns (price, interval_date_str) for the previous available
#     trading day relative to run_date, using JM's own data (so it already
#     accounts for weekends/holidays — whichever row JM considers the prior
#     trading day is used, no date arithmetic needed on our side)."""
#     columns = _fetch_month_columns(metal_code, run_date)

#     if len(columns) >= 2:
#         prev = columns[-2]
#         return float(prev[LONDON_PRICE_FIELD]), prev.get("IntervalDate")

#     if len(columns) == 1:
#         # Today is the first trading day recorded so far this month — the
#         # previous trading day falls in the prior month.
#         prev_month_columns = _fetch_month_columns(metal_code, _prev_month(run_date))
#         if prev_month_columns:
#             prev = prev_month_columns[-1]
#             return float(prev[LONDON_PRICE_FIELD]), prev.get("IntervalDate")

#     return None, None

def _latest_available_london_price(
    metal_code: str,
    requested_date: date
):
    """
    Returns the latest JM London 09:00 price whose actual
    IntervalDate is <= requested_date.

    This handles:
      - weekends
      - holidays
      - missing publication dates
      - month boundaries
    """

    month_to_check = requested_date.replace(day=1)

    # Check requested month, then progressively earlier months.
    for _ in range(12):

        columns = _fetch_month_columns(
            metal_code,
            month_to_check
        )

        candidates = []

        for row in columns:
            interval_date = _parse_interval_date(
                row.get("IntervalDate", "")
            )

            if interval_date is None:
                continue

            if interval_date <= requested_date:
                price = row.get(LONDON_PRICE_FIELD)

                if price in (None, "", "-"):
                    continue

                candidates.append(
                    (
                        interval_date,
                        float(price)
                    )
                )

        if candidates:
            actual_date, price = max(
                candidates,
                key=lambda x: x[0]
            )

            return price, actual_date

        month_to_check = _prev_month(
            month_to_check
        )

    return None, None

# def fetch_jm_prices(run_date: date = None) -> dict:
#     """Returns {"RHODIUM": 8800.0, "IRIDIUM": 7850.0, "RUTHENIUM": 1675.0} —
#     each the previous trading day's London 09:00 price."""
#     run_date = run_date or date.today()
#     prices = {}
#     for metal, code in JM_METAL_CODES.items():
#         price, interval_date = _previous_trading_day_london_price(code, run_date)
#         if price is None:
#             raise RuntimeError(f"Could not get a JM price for {metal} (code {code})")
#         prices[metal] = price
#         logger.info("JM %s: London 09:00 = %s (row: %s)", metal, price, interval_date)

#     if len(prices) != 3:
#         raise ValueError(f"Only found {len(prices)}/3 PGM prices: {prices}")

#     return prices

def fetch_jm_prices(run_date: date = None):
    """Returns:
        (
            {
                "RHODIUM": 8800.0,
                "IRIDIUM": 7850.0,
                "RUTHENIUM": 1675.0
            },
            {
                "RHODIUM": "30 Sep 2026",
                "IRIDIUM": "30 Sep 2026",
                "RUTHENIUM": "30 Sep 2026"
            }
        )

    Each price is the previous available trading day's
    London 09:00 price.
    """
    run_date = run_date or date.today()

    prices = {}
    price_dates = {}

    for metal, code in JM_METAL_CODES.items():

        price, interval_date = _latest_available_london_price(
            code,
            run_date
        )
        # price, interval_date = _previous_trading_day_london_price(
        #     code,
        #     run_date
        # )

        if price is None:
            raise RuntimeError(
                f"Could not get a JM price for {metal} (code {code})"
            )

        prices[metal] = price
        price_dates[metal] = interval_date

        logger.info(
            "JM %s: London 09:00 = %s (row: %s)",
            metal,
            price,
            interval_date
        )

    if len(prices) != 3:
        raise ValueError(
            f"Only found {len(prices)}/3 PGM prices: {prices}"
        )

    return prices, price_dates

def fetch_jm_prices(run_date: date = None):
    run_date = run_date or date.today()

    prices = {}
    price_dates = {}

    for metal, code in JM_METAL_CODES.items():

        price, interval_date = _latest_available_london_price(
            code,
            run_date
        )

        if price is None:
            raise RuntimeError(
                f"Could not get a JM price for "
                f"{metal} on or before {run_date}"
            )

        prices[metal] = price
        price_dates[metal] = interval_date

        logger.info(
            "JM %s: London 09:00 = %s (actual date: %s)",
            metal,
            price,
            interval_date
        )

    if len(prices) != 3:
        raise ValueError(
            f"Only found {len(prices)}/3 JM prices: {prices}"
        )

    return prices, price_dates
    
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    print(fetch_jm_prices())