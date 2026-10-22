"""
Central configuration for the Metal Rates Automator.
Edit the values below for your environment before running.
"""

import os
from dotenv import load_dotenv

load_dotenv()

EXCEL_PATH = os.environ.get(
    "METAL_RATES_XLSX",
    r"D:\Users\S1412\Documents\Workspace\metal_rates_automator\2026 Live Metals Rates Daily.xlsx",
)

LAYOUT = {
    "weekday_row": 2,       # e.g. "Monday", "Tuesday" ...
    "source_row": 3,        # "LBMA & JM" header row
    "date_row": 4,          # actual date value for each column
    "time_row": 5,          # "PM"
    "first_data_col": 3,    # column D = first date column
    "label_col": 2,         # column C = metal name labels
    "monthly_average_col": 7,

    # Metal name (as scraped) -> row number in the daily price block
    "metal_rows": {
        "GOLD": 6,
        "PLATINUM": 7,
        "PALLADIUM": 8,
        "SILVER": 9,
        "RHODIUM": 10,
        "IRIDIUM": 11,
        "RUTHENIUM": 12,
    },

    "IBJA_Rate": {
        "Silver999_row": 15
    },

    "Exchange_Rate": {
        "INR_USD_row": 18,
        "INR_EUR_row": 19
    },

    # Min/Date, Max/Date summary block
    "minmax_header_row": 15,
    "minmax_label_col": 3,
    "minmax_min_val_col": 3,
    "minmax_min_date_col": 4,
    "minmax_max_val_col": 5,
    "minmax_max_date_col": 6,
    "metal_minmax_rows": {
        "GOLD": 22,
        "PLATINUM": 23,
        "PALLADIUM": 24,
        "SILVER": 25,
        "RHODIUM": 26,
        "IRIDIUM": 27,
        "RUTHENIUM": 28,
    },
}

DATE_NUMBER_FORMAT = "dd-mm-yyyy"
PRICE_NUMBER_FORMATS = {
    "SILVER": "0.000",
    # all other metals default to 2 decimals — see DEFAULT_PRICE_NUMBER_FORMAT
}
DEFAULT_PRICE_NUMBER_FORMAT = "0.00"

METAL_ORDER = ["GOLD", "PLATINUM", "PALLADIUM", "SILVER", "RHODIUM", "IRIDIUM", "RUTHENIUM"]

TROY_OUNCE_IN_GRAMS = 31.1035

# ---------------------------------------------------------------------------
# LBMA — public JSON quote feed (no login needed; confirmed public as of
# Aug 2026, unaffected by the Nov 2025 MyLBMA Portal migration of the
# tabulated/historical data). Fixing time is the PM fix, matching the
# existing tracker's "PM" row.
# ---------------------------------------------------------------------------
LBMA_JSON_FEEDS = {
    "GOLD": "https://prices.lbma.org.uk/json/gold_pm.json",
    "PLATINUM": "https://prices.lbma.org.uk/json/platinum_pm.json",
    "PALLADIUM": "https://prices.lbma.org.uk/json/palladium_pm.json",
    "SILVER": "https://prices.lbma.org.uk/json/silver.json",
}

# Johnson Matthey — PGM prices and trading page (Rh, Ir, Ru)
JM_PAGE_URL = "https://matthey.com/products-and-markets/pgms-and-circularity/pgm-management"
JM_METALS = ["RHODIUM", "IRIDIUM", "RUTHENIUM"]

# Email dispatch details
SMTP_SERVER = os.environ.get("SMTP_SERVER", "smtp-mail.outlook.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USE_TLS = True
SMTP_USERNAME = os.environ.get("SMTP_USERNAME", "workflow@hp.co.in")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")  


EMAIL_FROM = SMTP_USERNAME
EMAIL_TO = [
    "jayesh.kapadnis@hp.co.in",  # replace with the real RS distribution list
    "sayali.takalkar@hp.co.in",
]
EMAIL_CC = [
    "sayali.takalkar@hp.co.in",
]
EMAIL_SUBJECT_TEMPLATE = "Daily Metal Prices Update — {date}"

FAILURE_EMAIL_TO = [
    "sayali.takalkar@hp.co.in",
]

# Logging
LOG_DIR = os.path.join(os.path.dirname(__file__), "logs")