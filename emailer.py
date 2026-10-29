"""
emailer.py

Builds and sends the daily LIVE METALS PRICES HTML report.

Includes:
- LBMA & Johnson Matthey metal prices (USD / Toz)
- AM / PM prices where available
- IBJA Silver 999 rates (INR / Kg)
- RBI exchange rates
- Monthly Min / Date / Max / Date table
- Current Excel tracker attachment
- Pipeline failure notification
"""

import logging
import os
import smtplib
from datetime import date, datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import config

logger = logging.getLogger(__name__)

METAL_DISPLAY_ORDER = [
    "GOLD",
    "PLATINUM",
    "PALLADIUM",
    "SILVER",
    "RHODIUM",
    "IRIDIUM",
    "RUTHENIUM",
]

BORDER = "1px solid #000000"
HEADER_BG = "#9CD3F0"      # blue header rows
LABEL_BG = "#D6EEFA"       # light-blue label / date cells
GREEN = "#008000"
RED = "#FF0000"
NEUTRAL = "#7F7F7F"        # "same as previous" marker colour

UP_SYMBOL = "&#9650;"      # ▲ rate increased
DOWN_SYMBOL = "&#9660;"    # ▼ rate decreased
SAME_SYMBOL = "&#9679;"    # ● rate unchanged

# Column widths (px) shared by EVERY table so the edges line up:
#   main / IBJA / FX tables = LABEL_W + 2 * VALUE_W
#   min-max table           = LABEL_W + 4 * VALUE_W
LABEL_W = 150
VALUE_W = 110

FONT = "font-family:Calibri,Arial,sans-serif;"
CELL_STYLE = (
    f"border:{BORDER};padding:4px 7px;{FONT}"
    "font-size:12px;line-height:15px;vertical-align:middle;"
)
HEADER_STYLE = f"{CELL_STYLE}font-weight:bold;background-color:{HEADER_BG};"
HEADER_STYLE2 = f"{CELL_STYLE}font-weight:bold;"
LABEL_STYLE = f"{CELL_STYLE}background-color:{LABEL_BG};"
CENTER = "text-align:center;"


def _fmt(value):
    """Format numeric values without changing missing values into zero."""
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return f"{value:,.2f}"
    return str(value)


def _fmt_4(value):
    """Format numeric values to 4 decimal places."""
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return f"{value:,.4f}"
    return str(value)


def _date_fmt(value):
    """Format date objects or already-formatted date strings."""
    if value is None:
        return ""
    if isinstance(value, (date, datetime)):
        return value.strftime("%d/%m/%Y")
    return str(value)

def _get_ibja_rate(ibja_rates, period):
    """Support both short and existing fetcher-style IBJA key names."""
    for key in (period, f"SILVER999_{period}", f"Silver999_{period}"):
        if key in ibja_rates:
            return ibja_rates[key]
    return None


def _td(content, style="", colspan=1, width=None):
    span = f" colspan='{colspan}'" if colspan > 1 else ""
    w = f" width='{width}'" if width else ""
    return (
        f"<td{span}{w} valign='middle' style='{style}'>"
        f"{content or '&nbsp;'}</td>"
    )


def _label(content, style=CELL_STYLE):
    """First (label) column, always LABEL_W wide."""
    return _td(content, style, width=LABEL_W)


def _val(content, style=CELL_STYLE + CENTER, span=1):
    """Value cell covering `span` value columns, centred by default."""
    return _td(content, style, colspan=span, width=VALUE_W * span)


def _table(rows, value_cols):
    total = LABEL_W + VALUE_W * value_cols
    body = "".join(f"<tr>{row}</tr>" for row in rows)
    return (
        f"<table cellspacing='0' cellpadding='0' border='0' width='{total}' "
        f"style='border-collapse:collapse;table-layout:fixed;width:{total}px;'>"
        f"{body}</table>"
    )


def _title(text):
    return (
        f"<p style='{FONT}font-size:12px;font-weight:bold;"
        f"margin:0 0 4px 0;'>{text}</p>"
    )


def _arrow_html(current, previous):
    """
    Email-only comparison marker:
      higher  -> green ▲,  lower -> red ▼,  same -> grey ●
    Nothing is shown if either value is unavailable.
    """
    if current is None or previous is None:
        return ""
    try:
        cur = round(float(current), 4)     # rounding avoids float noise
        prev = round(float(previous), 4)
    except (TypeError, ValueError):
        return ""

    if cur > prev:
        symbol, color, tip = UP_SYMBOL, GREEN, "Higher than previous recorded rate"
    elif cur < prev:
        symbol, color, tip = DOWN_SYMBOL, RED, "Lower than previous recorded rate"
    else:
        symbol, color, tip = SAME_SYMBOL, NEUTRAL, "Same as previous recorded rate"

    return f" <span style='color:{color};font-size:10px;' title='{tip}'>{symbol}</span>"


def _rate_html(value, previous=None, exchange=False):
    """Format a value and append its comparison marker."""
    formatted = (_fmt_4 if exchange else _fmt)(value)
    if value is None or value == "":
        return formatted
    return formatted + _arrow_html(value, previous)


def _price_cells(value, previous=None):
    """AM/PM cells for a metal; single-fixing metals span both columns."""
    if isinstance(value, dict):
        prev = previous if isinstance(previous, dict) else {}
        return (
            _val(_rate_html(value.get("AM"), prev.get("AM")))
            + _val(_rate_html(value.get("PM"), prev.get("PM")))
        )
    prev = previous if isinstance(previous, (int, float)) else None
    return _val(_rate_html(value, prev), span=2)


def build_html(
    prices,
    minmax,
    target_date,
    ibja_rates=None,
    exchange_rates=None,
    ibja_dates=None,          # accepted for compatibility; not shown
    previous_prices=None,
    previous_ibja_rates=None,
    previous_exchange_rates=None,
):
    """
    Build the email body. Previous-rate dictionaries are used only to
    render the comparison markers; they are never written to Excel.
    """
    previous_prices = previous_prices or {}
    previous_ibja_rates = previous_ibja_rates or {}
    previous_exchange_rates = previous_exchange_rates or {}
    ibja_rates = ibja_rates or {}
    exchange_rates = exchange_rates or {}

    para = f"{FONT}font-size:13px;margin:0 0 16px 0;"

    parts = [
        f"<html><body style='{FONT}font-size:13px;'>",
        f"<p style='{para}'>Dear Team,</p>",
        f"<p style='{para}'>Please find below the metal prices in USD as on "
        f"<b>{target_date:%d %b %Y}</b> with reference to LBMA website for Au, Ag, "
        "LPPM website for Pt, Pd &amp; JM website for Ir, Rh, Ru.</p>",
    ]

    # ---- Live metals table (label | AM | PM) -------------------------
    rows = [
        _label("Day", HEADER_STYLE)
        + _val(f"{target_date:%A}", HEADER_STYLE + CENTER, 2),
        _label("LIVE METALS PRICES <br>in ($ / Toz):", HEADER_STYLE2)
        + _val("LBMA &amp; JM", HEADER_STYLE2 + CENTER, 2),
        _label("Date", LABEL_STYLE)
        + _val(f"{target_date:%d-%m-%Y}", LABEL_STYLE + CENTER, 2),
        _label("Time", HEADER_STYLE)
        + _val("AM", HEADER_STYLE + CENTER)
        + _val("PM", HEADER_STYLE + CENTER),
    ]
    for metal in METAL_DISPLAY_ORDER:
        value = prices.get(metal)
        if value is None:
            continue
        rows.append(_label(metal) + _price_cells(value, previous_prices.get(metal)))
    parts.append(_table(rows, 2))
    parts.append("<br>")

    # ---- IBJA Silver 999 (label | AM | PM) ---------------------------
    ibja_am = _get_ibja_rate(ibja_rates, "AM")
    ibja_pm = _get_ibja_rate(ibja_rates, "PM")
    if ibja_am is not None or ibja_pm is not None:
        prev_am = _get_ibja_rate(previous_ibja_rates, "AM")
        prev_pm = _get_ibja_rate(previous_ibja_rates, "PM")
        parts.append(_title("IBJA Silver Rate (INR/Kg)"))
        parts.append(_table([
            _label("Silver999", HEADER_STYLE)
            + _val("AM", HEADER_STYLE + CENTER)
            + _val("PM", HEADER_STYLE + CENTER),
            _label("Silver 999", LABEL_STYLE)
            + _val(_rate_html(ibja_am, prev_am))
            + _val(_rate_html(ibja_pm, prev_pm)),
        ], 2))
        parts.append("<br>")

    # ---- RBI exchange rates (label | value spanning AM+PM) -----------
    fx_rows = []
    for currency in ("USD", "EUR"):
        rate = exchange_rates.get(currency)
        if rate is None:
            continue
        fx_rows.append(
            _label(f"INR / 1 {currency}", LABEL_STYLE)
            + _val(_rate_html(rate, previous_exchange_rates.get(currency),
                              exchange=True), span=2)
        )
    if fx_rows:
        parts.append(_title("Exchange Rate from RBI Rate"))
        parts.append(_table(fx_rows, 2))
        parts.append("<br>")

    # ---- Monthly min / max (label | min | date | max | date) --------
    if minmax:
        parts.append(
            f"<p style='{para}'>Please find in below table, the metal rates "
            f"for <b>{target_date:%b %Y}</b> with Minimum &amp; Maximum rates:</p>"
        )
        mm_rows = [
            _label("LIVE METALS PRICES <br>in ($ / Toz):", HEADER_STYLE)
            + _val("Min / Date", HEADER_STYLE + CENTER, 2)
            + _val("Max / Date", HEADER_STYLE + CENTER, 2)
        ]
        for metal in METAL_DISPLAY_ORDER:
            item = minmax.get(metal)
            if not item:
                continue
            mm_rows.append(
                _label(metal)
                + _val(_fmt(item.get("min", item.get("minimum"))))
                + _val(_date_fmt(item.get("min_date")), LABEL_STYLE + CENTER)
                + _val(_fmt(item.get("max", item.get("maximum"))))
                + _val(_date_fmt(item.get("max_date")), LABEL_STYLE + CENTER)
            )
        parts.append(_table(mm_rows, 4))

    parts.append(
        "</body></html><br><br>"
        f"<p style='{FONT}font-size:11px;margin:0 0 10px 0;'>"
        "<b>Note:</b> "
        f"<span style='color:{GREEN};'>{UP_SYMBOL}</span> Higher than previous rate"
        "&nbsp;&nbsp;|&nbsp;&nbsp;"
        f"<span style='color:{RED};'>{DOWN_SYMBOL}</span> Lower than previous rate"
        "&nbsp;&nbsp;|&nbsp;&nbsp;"
        f"<span style='color:{NEUTRAL};'>{SAME_SYMBOL}</span> Same as previous rate"
        "</p>"
        f"<p style='margin-top:18px;color:#666666;font-size:11px;'>This email was automatically generated by the Metal Rates Automator.</p>"
    )
    return "".join(parts)


def send_email(html_body, attachment_path=None, subject=None,target_date=None,):
    """Send the HTML rates report and optionally attach the Excel workbook."""
    smtp_host = getattr(config, "SMTP_SERVER", None) or os.getenv("SMTP_SERVER")
    smtp_port = getattr(config, "SMTP_PORT", None) or os.getenv("SMTP_PORT", "587")
    smtp_username = (
        getattr(config, "SMTP_USERNAME", None)
        or os.getenv("SMTP_USERNAME", "workflow@hp.co.in")
    )
    smtp_password = (
        getattr(config, "SMTP_PASSWORD", None)
        or os.getenv("SMTP_PASSWORD")
    )

    sender = (
        getattr(config, "EMAIL_SENDER", None)
        or os.getenv("EMAIL_SENDER")
        or smtp_username
    )

    recipients = (
        getattr(config, "EMAIL_RECIPIENTS", None)
        or getattr(config, "EMAIL_TO", None)
        or os.getenv("EMAIL_RECIPIENTS")
        or os.getenv("EMAIL_TO", "")
    )

    if isinstance(recipients, str):
        recipients = [
            address.strip()
            for address in recipients.replace(";", ",").split(",")
            if address.strip()
        ]

    if not smtp_host:
        raise ValueError("SMTP_SERVER is not configured.")
    if not smtp_username or not smtp_password:
        raise ValueError("SMTP_USERNAME / SMTP_PASSWORD is not configured.")
    if not sender:
        raise ValueError("EMAIL_SENDER is not configured.")
    if not recipients:
        raise ValueError("No email recipients are configured.")

    message = MIMEMultipart("mixed")
    message["From"] = sender
    message["To"] = ", ".join(recipients)

    email_date = target_date or date.today()

    if isinstance(email_date, datetime):
        email_date = email_date.date()

    message["Subject"] = subject or (
        f"Daily Metal Rates - {email_date.strftime('%d/%m/%Y')}"
    )

    alternative = MIMEMultipart("alternative")
    alternative.attach(MIMEText("Please view this email in HTML format.", "plain"))
    alternative.attach(MIMEText(html_body, "html"))
    message.attach(alternative)

    if attachment_path and os.path.isfile(attachment_path):
        with open(attachment_path, "rb") as file:
            attachment = MIMEApplication(
                file.read(),
                Name=os.path.basename(attachment_path),
            )

        attachment["Content-Disposition"] = (
            f'attachment; filename="{os.path.basename(attachment_path)}"'
        )
        message.attach(attachment)

    with smtplib.SMTP(smtp_host, int(smtp_port), timeout=30) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(smtp_username, smtp_password)
        server.sendmail(sender, recipients, message.as_string())

    logging.info(
        "Rates email sent successfully to %s",
        ", ".join(recipients),
    )


def send_failure_email(error_message):
    """Send a plain failure notification if email configuration is available."""
    try:
        html = (
            "<html><body>"
            "<h3>Metal Rates Automator - Run Failed</h3>"
            f"<pre>{str(error_message)}</pre>"
            "</body></html>"
        )
        send_email(
            html_body=html,
            attachment_path=None,
            subject="Metal Rates Automator - Failure",
        )
    except Exception:
        logging.exception("Could not send failure notification email.")
