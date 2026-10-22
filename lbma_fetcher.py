import logging

logger = logging.getLogger(__name__)

import logging
from datetime import date, timedelta
from typing import Optional

from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)

logger = logging.getLogger(__name__)


LBMA_PRICE_PAGE = (
    "https://www.lbma.org.uk/prices-and-data/"
    "lbma-precious-metal-prices"
)


METALS = [
    "GOLD",
    "PLATINUM",
    "PALLADIUM",
    "SILVER",
]


# ---------------------------------------------------------------------------
# Required LBMA series
# ---------------------------------------------------------------------------

# Gold / Platinum / Palladium have AM + PM.
# Silver has one daily fixing.
REQUIRED_SERIES = {
    ("GOLD", "AM"): "gold_am.json",
    ("GOLD", "PM"): "gold_pm.json",
    ("SILVER", "PM"): "silver.json",        # one daily fixing, filed under PM
    ("PLATINUM", "AM"): "platinum_am.json",
    ("PLATINUM", "PM"): "platinum_pm.json",
    ("PALLADIUM", "AM"): "palladium_am.json",
    ("PALLADIUM", "PM"): "palladium_pm.json",
}

FILENAME_TO_SERIES = {name: key for key, name in REQUIRED_SERIES.items()}

# Silver has no AM fixing.
AM_METALS = ["GOLD", "PLATINUM", "PALLADIUM"]


# ---------------------------------------------------------------------------
# Identify metal from URL
# ---------------------------------------------------------------------------

def _identify_metal(url: str) -> Optional[str]:
    """
    Identify metal from an LBMA JSON URL.
    """

    url_lower = url.lower()

    if "/gold" in url_lower:
        return "GOLD"

    if "/silver" in url_lower:
        return "SILVER"

    if "/platinum" in url_lower:
        return "PLATINUM"

    if "/palladium" in url_lower:
        return "PALLADIUM"

    return None


# ---------------------------------------------------------------------------
# Identify which required series a URL is
# ---------------------------------------------------------------------------

def _series_key(url: str):
    """
    ("GOLD", "AM"), ("PLATINUM", "PM"), ... if the URL is one of the
    series we need; otherwise None.

        .../json/limited/gold_am.json   -> ("GOLD", "AM")
        .../json/limited/gold_pm.json   -> ("GOLD", "PM")
        .../json/limited/silver.json    -> ("SILVER", "PM")
    """
    filename = url.lower().split("?")[0].split("/")[-1]
    return FILENAME_TO_SERIES.get(filename)


# ---------------------------------------------------------------------------
# Parse LBMA JSON
# ---------------------------------------------------------------------------

def _build_date_map(data: list) -> dict:
    """
    Convert LBMA JSON:

        [
            {
                "is_cms_locked": 0,
                "d": "2026-10-01",
                "v": [
                    4159.85,
                    3100.20,
                    3500.10
                ]
            }
        ]

    into:

        {
            "2026-10-01": 4159.85
        }

    v[0] = USD
    """

    result = {}

    if not isinstance(data, list):

        logger.warning(
            "Unexpected LBMA response type: %s",
            type(data).__name__,
        )

        return result

    for row in data:

        if not isinstance(row, dict):
            continue

        trading_date = row.get("d")
        values = row.get("v")

        if not trading_date:
            continue

        if not isinstance(values, list):
            continue

        if len(values) == 0:
            continue

        usd_price = values[0]

        if usd_price in (None, "", "-"):
            continue

        try:

            result[str(trading_date)] = float(
                usd_price
            )

        except (TypeError, ValueError):

            logger.debug(
                "Ignoring invalid LBMA price row: %s",
                row,
            )

    return result


# ---------------------------------------------------------------------------
# Find latest available trading day
# ---------------------------------------------------------------------------

def _find_latest_price(
    by_date: dict,
    target_date: date,
    max_days: int = 10,
):
    """
    Exact date first.

    If unavailable, walk backwards.

    This handles:

        Saturday
        Sunday
        UK bank holidays
        price not published yet
    """

    for delta in range(max_days + 1):

        candidate = (
            target_date -
            timedelta(days=delta)
        )

        candidate_string = candidate.isoformat()

        if candidate_string in by_date:

            return (
                by_date[candidate_string],
                candidate_string,
            )

    return None, None


# ---------------------------------------------------------------------------
# Browser capture
# ---------------------------------------------------------------------------

def _make_response_handler(captured: dict):
    """Build the Playwright 'response' callback that stores each required
    LBMA JSON series in `captured`, keyed by (METAL, "AM"|"PM")."""

    def handle_response(response):

        url = response.url

        if "prices.lbma.org.uk/json/" not in url:
            return

        key = _series_key(url)

        if key is None:
            logger.debug("Ignoring non-required series: %s", url)
            return

        logger.info("Required LBMA response detected: %s", url)

        try:
            data = response.json()
        except Exception as exc:
            logger.warning(
                "Could not parse %s %s response: %s", key[0], key[1], exc
            )
            return

        if isinstance(data, list):
            captured[key] = data
            logger.info("Captured %s %s: %d rows", key[0], key[1], len(data))
        else:
            logger.warning(
                "%s %s returned unexpected JSON structure", key[0], key[1]
            )

    return handle_response


def _await_capture(page, captured: dict, key, timeout_ms: int = 10000) -> bool:
    """Wait until the series `key` has been captured (or time runs out)."""
    waited = 0
    while captured.get(key) is None and waited < timeout_ms:
        page.wait_for_timeout(250)
        waited += 250
    return captured.get(key) is not None


def _capture_fixing(page, captured: dict, key, fixing: str) -> bool:
    """
    Click the AM / PM button and wait for that series.

    The widget remembers the last fixing when you switch metal, so the
    button is always clicked explicitly rather than assuming the default.
    One retry, because a click can land before the widget has re-rendered.
    """
    for attempt in (1, 2):
        _select_fixing(page, fixing)

        if _await_capture(page, captured, key):
            return True

        logger.warning(
            "%s %s not received after click (attempt %d)",
            key[0], key[1], attempt,
        )

    logger.error("LBMA %s %s fixing could not be captured", key[0], key[1])
    return False


def _capture_all_series(page, captured: dict):
    """
    For each metal: select it, then capture AM and PM (Gold / Platinum /
    Palladium) or just the daily series (Silver).
    """
    for label, metal in (
        ("Gold", "GOLD"),
        ("Silver", "SILVER"),
        ("Platinum", "PLATINUM"),
        ("Palladium", "PALLADIUM"),
    ):
        _select_metal(page, label)

        if metal in AM_METALS:
            _capture_fixing(page, captured, (metal, "AM"), "am")
            _capture_fixing(page, captured, (metal, "PM"), "pm")
        else:
            # Silver: the series loads as soon as the metal is selected.
            if not _await_capture(page, captured, (metal, "PM")):
                logger.error("LBMA %s daily fixing was not received", metal)

        page.wait_for_timeout(1000)


def _fetch_lbma_chart_data() -> dict:
    """
    Open the LBMA page, step through every metal / fixing we need, and
    return the captured JSON series:

        {
            ("GOLD", "AM"): [...],      ("GOLD", "PM"): [...],
            ("SILVER", "PM"): [...],
            ("PLATINUM", "AM"): [...],  ("PLATINUM", "PM"): [...],
            ("PALLADIUM", "AM"): [...], ("PALLADIUM", "PM"): [...],
        }

    A series that could not be captured is simply absent / None.
    """

    captured = {key: None for key in REQUIRED_SERIES}

    with sync_playwright() as p:

        logger.info("Launching Chromium for LBMA")

        browser = p.chromium.launch(headless=True)

        try:

            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/154.0.0.0 Safari/537.36"
                ),
                locale="en-US",
                timezone_id="Asia/Kolkata",
                viewport={"width": 1920, "height": 1080},
            )

            page = context.new_page()

            page.on("response", _make_response_handler(captured))

            logger.info("Opening LBMA price page: %s", LBMA_PRICE_PAGE)

            try:
                page.goto(
                    LBMA_PRICE_PAGE,
                    wait_until="domcontentloaded",
                    timeout=60000,
                )
            except PlaywrightTimeoutError:
                logger.warning(
                    "LBMA page load timed out; "
                    "continuing because browser may still be usable"
                )

            # Give the Vue application time to initialise.
            page.wait_for_timeout(5000)

            logger.info("Searching for LBMA Metal selector")

            _capture_all_series(page, captured)

            logger.info(
                "LBMA capture complete: %s",
                {
                    f"{metal} {fixing}": captured[(metal, fixing)] is not None
                    for (metal, fixing) in REQUIRED_SERIES
                },
            )

        finally:
            browser.close()

    return captured


# ---------------------------------------------------------------------------
# Metal selector
# ---------------------------------------------------------------------------

def _select_metal(
    page,
    metal_name: str,
):
    """
    Select a metal from the LBMA Metal control.

    The LBMA UI can change its internal Vue structure, so we try several
    safe strategies rather than depending on one CSS class.
    """

    logger.info(
        "Selecting LBMA metal: %s",
        metal_name,
    )

    # ---------------------------------------------------------------
    # Strategy 1: native <select>
    # ---------------------------------------------------------------

    selects = page.locator("select")

    count = selects.count()

    logger.debug(
        "Found %d select elements",
        count,
    )

    for i in range(count):

        select = selects.nth(i)

        try:

            options = select.locator("option")

            option_count = options.count()

            option_texts = []

            for j in range(option_count):

                try:
                    option_texts.append(
                        options.nth(j).inner_text()
                    )
                except Exception:
                    pass

            logger.debug(
                "Select %d options: %s",
                i,
                option_texts,
            )

            for j in range(option_count):

                option = options.nth(j)

                text = option.inner_text().strip()

                if text.lower() == metal_name.lower():

                    value = option.get_attribute(
                        "value"
                    )

                    if value is not None:

                        select.select_option(
                            value=value
                        )

                    else:

                        select.select_option(
                            label=metal_name
                        )

                    logger.info(
                        "Selected %s using native select",
                        metal_name,
                    )

                    return

        except Exception as exc:

            logger.debug(
                "Native select strategy failed: %s",
                exc,
            )

    # ---------------------------------------------------------------
    # Strategy 2: click visible Metal control
    # ---------------------------------------------------------------

    try:

        page.get_by_text(
            "Metal",
            exact=True,
        ).click(
            timeout=3000
        )

        page.wait_for_timeout(500)

        page.get_by_text(
            metal_name,
            exact=True,
        ).click(
            timeout=3000
        )

        logger.info(
            "Selected %s using Metal menu",
            metal_name,
        )

        return

    except Exception as exc:

        logger.debug(
            "Metal menu strategy failed: %s",
            exc,
        )

    # ---------------------------------------------------------------
    # Strategy 3: click metal text directly
    # ---------------------------------------------------------------

    try:

        page.get_by_text(
            metal_name,
            exact=True,
        ).click(
            timeout=3000
        )

        logger.info(
            "Selected %s directly",
            metal_name,
        )

        return

    except Exception as exc:

        logger.error(
            "Could not select LBMA metal %s: %s",
            metal_name,
            exc,
        )

        raise RuntimeError(
            f"Could not select LBMA metal: {metal_name}"
        ) from exc


# ---------------------------------------------------------------------------
# AM / PM fixing selector
# ---------------------------------------------------------------------------

def _dismiss_cookiebot(page):
    """Cookiebot can sit over the LBMA widget and swallow clicks."""
    try:
        cookie_dialog = page.locator("#CybotCookiebotDialog")

        if cookie_dialog.is_visible(timeout=1500):
            logger.info("Cookiebot dialog detected")

            cookie_buttons = [
                "#CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll",
                "#CybotCookiebotDialogBodyButtonAccept",
                "button:has-text('Allow all')",
                "button:has-text('Accept all')",
                "button:has-text('Accept')",
            ]

            clicked = False

            for selector in cookie_buttons:
                try:
                    button = page.locator(selector).first

                    if button.is_visible(timeout=500):
                        logger.info(
                            "Clicking Cookiebot consent button: %s",
                            selector,
                        )

                        button.click(timeout=2000, force=True)

                        clicked = True
                        break

                except Exception:
                    continue

            if clicked:
                page.wait_for_timeout(1000)
                logger.info("Cookiebot dialog handled")

    except Exception as exc:
        logger.debug("Cookiebot handling skipped: %s", exc)


def _select_fixing(page, fixing: str):
    """
    Select the AM or PM fixing on the LBMA price widget.

    fixing: "am" or "pm"
    """

    fixing = fixing.lower()
    label = fixing.upper()

    logger.info("Selecting %s fixing", label)

    # 1. Remove the Cookiebot blocker if present
    _dismiss_cookiebot(page)

    # 2. Locate the button: data attribute first, visible text as backup
    button = page.locator(f'button[data-fixing="{fixing}"]').first

    try:
        button.wait_for(state="visible", timeout=5000)
    except Exception:
        logger.warning(
            "No button[data-fixing=\"%s\"]; trying the visible %s text",
            fixing, label,
        )
        button = page.get_by_role("button", name=label, exact=True).first

        try:
            button.wait_for(state="visible", timeout=3000)
        except Exception as exc:
            raise RuntimeError(
                f"LBMA {label} button was not found"
            ) from exc

    # 3. Normal click
    try:
        button.click(timeout=3000)

        page.wait_for_timeout(1500)

        logger.info("%s fixing selected", label)
        return

    except Exception as exc:
        logger.warning("Normal %s click failed: %s", label, exc)

    # 4. JavaScript fallback
    try:
        logger.info("Using JavaScript fallback to select %s", label)

        page.evaluate(
            """
            (fixing) => {
                const button = document.querySelector(
                    `button[data-fixing="${fixing}"]`
                );

                if (!button) {
                    throw new Error("fixing button not found");
                }

                button.click();
            }
            """,
            fixing,
        )

        page.wait_for_timeout(2000)

        logger.info("%s fixing selected using JavaScript", label)

    except Exception as exc:
        raise RuntimeError(
            f"Could not select LBMA {label} fixing"
        ) from exc


def _select_pm(page):
    """Kept for backwards compatibility."""
    _select_fixing(page, "pm")


# ---------------------------------------------------------------------------
# Public functions
# ---------------------------------------------------------------------------

def _resolve_series(raw_data: dict, key, target_date: date):
    """
    (price, "YYYY-MM-DD") for the latest price on or before target_date
    in series `key`; (None, None) if the series is missing or has no
    usable price in range.
    """
    data = raw_data.get(key)

    if not data:
        return None, None

    by_date = _build_date_map(data)

    if not by_date:
        return None, None

    return _find_latest_price(by_date, target_date)


def fetch_lbma_fixings(
    target_date: Optional[date] = None,
) -> dict:
    """
    Fetch the LBMA USD fixings in ONE browser session.

    Returns:

        {
            "AM": {
                "prices": {"GOLD": 4141.65, "PLATINUM": 1731.65, "PALLADIUM": 1181.55},
                "dates":  {"GOLD": "2026-10-14", "PLATINUM": "2026-10-14", ...},
            },
            "PM": {
                "prices": {"GOLD": ..., "SILVER": ..., "PLATINUM": ..., "PALLADIUM": ...},
                "dates":  {"GOLD": "2026-10-13", ...},
            },
        }

    Each price carries ITS OWN actual date: the latest published fixing on
    or before target_date (weekends, UK bank holidays and not-yet-published
    fixings fall back to the previous available day). The AM fixing is set
    earlier in the London day than the PM one, so on an early run AM can
    legitimately be one trading day ahead of PM.

    Silver has a single daily fixing and is reported under "PM".

    PM is required: a missing PM series raises RuntimeError. AM is
    best-effort: an AM series that can't be captured is left out of
    result["AM"] (and logged) rather than failing the whole run.
    """

    target_date = target_date or date.today()

    logger.info("Fetching LBMA fixings for %s", target_date)

    raw_data = _fetch_lbma_chart_data()

    result = {
        "AM": {"prices": {}, "dates": {}},
        "PM": {"prices": {}, "dates": {}},
    }

    # ---- PM (and silver's daily fixing): required -----------------------
    for metal in METALS:

        key = (metal, "PM")

        if not raw_data.get(key):
            raise RuntimeError(
                f"LBMA did not provide required {metal} JSON series."
            )

        price, actual_date = _resolve_series(raw_data, key, target_date)

        if price is None:
            raise RuntimeError(
                f"No LBMA {metal} price found within 10 days "
                f"before {target_date}."
            )

        if actual_date != target_date.isoformat():
            logger.warning(
                "%s PM: no price for %s; using most recent available price: %s",
                metal, target_date, actual_date,
            )

        result["PM"]["prices"][metal] = price
        result["PM"]["dates"][metal] = actual_date

        logger.info("LBMA %s PM (%s): %s", metal, actual_date, price)

    # ---- AM: best effort -------------------------------------------------
    for metal in AM_METALS:

        key = (metal, "AM")

        price, actual_date = _resolve_series(raw_data, key, target_date)

        if price is None:
            logger.error(
                "LBMA %s AM price unavailable — leaving it out.", metal
            )
            continue

        if actual_date != target_date.isoformat():
            logger.warning(
                "%s AM: no price for %s; using most recent available price: %s",
                metal, target_date, actual_date,
            )

        result["AM"]["prices"][metal] = price
        result["AM"]["dates"][metal] = actual_date

        logger.info("LBMA %s AM (%s): %s", metal, actual_date, price)

    return result


def fetch_lbma_prices(
    target_date: Optional[date] = None,
) -> tuple[dict, str]:
    """
    Backwards-compatible PM-only view of fetch_lbma_fixings().

    Returns:

        (
            {"GOLD": 4159.85, "SILVER": 60.92, "PLATINUM": 1727.35, "PALLADIUM": 1193.95},
            "2026-10-01"
        )

    The date is the actual LBMA price date (the OLDEST one if the metals
    disagree, so it never implies fresher data than a metal really has).
    """

    pm = fetch_lbma_fixings(target_date)["PM"]

    unique_dates = set(pm["dates"].values())

    if len(unique_dates) > 1:
        logger.warning(
            "LBMA metals have different price dates: %s", pm["dates"]
        )

    return pm["prices"], min(unique_dates)


# ---------------------------------------------------------------------------
# Standalone test
# ---------------------------------------------------------------------------
if __name__ == "__main__":

    logging.basicConfig(
        level=logging.ERROR,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    try:

        fixings = fetch_lbma_fixings()

        print()
        print("=" * 60)
        print("LBMA FIXINGS - USD")
        print("=" * 60)

        for session in ("AM", "PM"):

            print(f"\n{session} fixing")
            print("-" * 60)

            if not fixings[session]["prices"]:
                print("  (none)")

            for metal, price in fixings[session]["prices"].items():
                print(
                    f"  {metal:<12}: {price:>10.3f}   "
                    f"date: {fixings[session]['dates'][metal]}"
                )

        print("=" * 60)

    except Exception:

        logger.exception("LBMA price fetch failed")