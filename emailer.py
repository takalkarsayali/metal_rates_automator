
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
PINK = "#E5A0D8"
LIGHT_PINK = "#EFD0EB"

CELL_STYLE = (
    f"border:{BORDER};"
    "padding:3px 7px;"
    "font-family:Calibri,Arial,sans-serif;"
    "font-size:12px;"
)


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


def _price_cells(value):
    """
    Return the AM/PM cells for a metal.

    Accepted formats:
        {"AM": 4158.20, "PM": 4152.25}
        {"SILVER_AM": 56.715, "SILVER_PM": 56.715}
        9150.00  # single fixing
        None
    """
    if isinstance(value, dict):
        am = value.get("AM")
        pm = value.get("PM")

        return (
            f"<td style='{CELL_STYLE}text-align:center;'>"
            f"{_fmt(am)}</td>"
            f"<td style='{CELL_STYLE}text-align:center;'>"
            f"{_fmt(pm)}</td>"
        )

    # Single fixing: show one value across the AM/PM columns.
    return (
        f"<td colspan='2' "
        f"style='{CELL_STYLE}text-align:center;'>"
        f"{_fmt(value)}</td>"
    )


def _get_ibja_rate(ibja_rates, period):
    """Support both short and existing fetcher-style IBJA key names."""
    candidates = (
        period,
        f"SILVER999_{period}",
        f"Silver999_{period}",
    )

    for key in candidates:
        if key in ibja_rates:
            return ibja_rates[key]

    return None


# def build_html(
#     prices: dict,
#     minmax: dict,
#     target_date: date,
#     ibja_rates: dict = None,
#     exchange_rates: dict = None,
#     ibja_dates: dict = None,
# ) -> str:
#     ibja_rates = ibja_rates or {}
#     exchange_rates = exchange_rates or {}
#     ibja_dates = ibja_dates or {}

#     # Normalize metal keys to uppercase.
#     prices = {
#         str(key).upper(): value
#         for key, value in prices.items()
#     }

#     minmax = {
#         str(key).upper(): value
#         for key, value in minmax.items()
#     }

#     # Main live metals price rows.
#     price_rows = "".join(
#         f"<tr>"
#         f"<td style='{CELL_STYLE}'>{metal}</td>"
#         f"{_price_cells(prices.get(metal))}"
#         f"</tr>"
#         for metal in METAL_DISPLAY_ORDER
#     )

#     # Monthly Min / Date / Max / Date rows.
#     minmax_rows = "".join(
#         f"<tr>"
#             f"<td style='{CELL_STYLE}'>{metal}</td>"
#             f"<td style='{CELL_STYLE}text-align:center;'>"
#             f"{_fmt(minmax[metal].get('min'))}</td>"
#             f"<td style='{CELL_STYLE}background:{LIGHT_PINK};text-align:center;'>"
#             f"{_date_fmt(minmax[metal].get('min_date'))}</td>"
#             f"<td style='{CELL_STYLE}text-align:center;'>"
#             f"{_fmt(minmax[metal].get('max'))}</td>"
#             f"<td style='{CELL_STYLE}background:{LIGHT_PINK};text-align:center;'>"
#             f"{_date_fmt(minmax[metal].get('max_date'))}</td>"
#         f"</tr>"
#         for metal in METAL_DISPLAY_ORDER
#         if metal in minmax
#     )

#     # IBJA Silver 999 AM/PM rates.
#     ibja_am = _get_ibja_rate(ibja_rates, "AM")
#     ibja_pm = _get_ibja_rate(ibja_rates, "PM")

#     # RBI exchange-rate rows.
#     exchange_rows = "".join(
#         f"<tr>"
#         f"<td style='{CELL_STYLE}background:{LIGHT_PINK};width:300px;'>"
#         f"INR / 1 {currency}</td>"
#         f"<td colspan='2' style='{CELL_STYLE}text-align:center;'>"
#         f"{_fmt(rate)}</td>"
#         f"</tr>"
#         for currency, rate in exchange_rates.items()
#     )

#     return f"""
#     <!DOCTYPE html>
#     <html>
#     <body style="
#         font-family:Calibri,Arial,sans-serif;
#         font-size:12px;
#         color:#111111;
#     ">

#         <p style="font-size: 15px"> Dear Team, </p>
#         <p style="font-size: 15px">
#             Please find below the metal prices in USD as on
#             <b>{target_date:%d %b %Y}</b> with reference to LBMA website
#             for Au, Ag, LPPM website for Pt, Pd &amp; JM website for Ir,
#             Rh, Ru.
#         </p>

#         <!-- LIVE METALS PRICE TABLE -->

#         <table style="
#             border-collapse:collapse;
#             width:300px;
#         ">
#             <tr>
#                 <th style="{CELL_STYLE}background:{LIGHT_PINK};text-align:left;">
#                     Day
#                 </th>
#                 <th colspan="2"
#                     style="{CELL_STYLE}background:{LIGHT_PINK};text-align:center;">
#                     {target_date:%A}
#                 </th>
#             </tr>

#             <tr>
#                 <th style="text-align:left;">
#                     LIVE METALS PRICES in ($ / Toz):
#                 </th>
#                 <th colspan="2"
#                     style="text-align:center;">
#                     LBMA &amp; JM
#                 </th>
#             </tr>

#             <tr>
#                 <th style="{CELL_STYLE}background:{PINK};text-align:left;">
#                     Date
#                 </th>
#                 <th colspan="2"
#                     style="{CELL_STYLE}background:{PINK};text-align:center;">
#                     {target_date:%d-%m-%Y}
#                 </th>
#             </tr>

#             <tr>
#                 <th style="{CELL_STYLE}text-align:left;">Time</th>
#                 <th style="{CELL_STYLE}text-align:center;">AM</th>
#                 <th style="{CELL_STYLE}text-align:center;">PM</th>
#             </tr>
#             {price_rows}
#         </table>

#         <!-- IBJA SILVER RATE -->

#         <p style="margin-top:18px;margin-bottom:4px;">
#             <b>IBJA Silver Rate (INR/Kg)</b>
#         </p>

#         <table style="border-collapse:collapse;width:300px;">
#             <tr>
#                 <th style="{CELL_STYLE}text-align:center;">Silver999</th>
#                 <th style="{CELL_STYLE}text-align:center;">AM</th>
#                 <th style="{CELL_STYLE}text-align:center;">PM</th>
#             </tr>
#             <tr>
#                 <td style="{CELL_STYLE}background:{LIGHT_PINK};width:300px;">
#                     Silver 999
#                 </td>
#                 <td style="{CELL_STYLE}text-align:center;">
#                     {_fmt(ibja_am)}
#                 </td>
#                 <td style="{CELL_STYLE}text-align:center;">
#                     {_fmt(ibja_pm)}
#                 </td>
#             </tr>
#         </table>

#         <!-- RBI EXCHANGE RATES -->

#         <p style="margin-top:28px;margin-bottom:4px;">
#             <b>Exchange Rate from RBI Rate</b>
#         </p>

#         <table style="border-collapse:collapse;width:300px;">
#             {exchange_rows}
#         </table>

#         <!-- MONTHLY MINIMUM / MAXIMUM TABLE -->

#         <p style="margin-top:20px; font-size: 15px">
#             Please find in below table, the metal rates for
#             <b>{target_date:%b %Y}</b> with Minimum &amp; Maximum rates:
#         </p>

#         <table style="border-collapse:collapse;width:300px;">
#             <tr style="background:{PINK};">
#                 <th rowspan="1"
#                     style="{CELL_STYLE}text-align:left;width:125px;">
#                     LIVE METALS PRICES<br/>
#                     in ($ / Toz):
#                 </th>
#                 <th colspan="2"
#                     style="{CELL_STYLE}text-align:center;">
#                     Min / Date
#                 </th>
#                 <th colspan="2"
#                     style="{CELL_STYLE}text-align:center;">
#                     Max / Date
#                 </th>
#             </tr>
#             {minmax_rows}
#         </table>

#         <p style="margin-top:18px;color:#666666;font-size:11px;">
#             This email was automatically generated by the Metal Rates
#             Automator.
#         </p>

#     </body>
#     </html>
#     """

def build_html(
    prices: dict,
    minmax: dict,
    target_date: date,
    ibja_rates: dict = None,
    exchange_rates: dict = None,
    ibja_dates: dict = None,
) -> str:
    ibja_rates = ibja_rates or {}
    exchange_rates = exchange_rates or {}
    ibja_dates = ibja_dates or {}

    # Normalize metal keys to uppercase.
    prices = {
        str(key).upper(): value
        for key, value in prices.items()
    }

    minmax = {
        str(key).upper(): value
        for key, value in minmax.items()
    }

    # Main live metals price rows.
    price_rows = "".join(
        f"<tr>"
        f"<td style='{CELL_STYLE}'>{metal}</td>"
        f"{_price_cells(prices.get(metal))}"
        f"</tr>"
        for metal in METAL_DISPLAY_ORDER
    )

    # Monthly Min / Date / Max / Date rows.
    minmax_rows = "".join(
        f"<tr>"
            f"<td style='{CELL_STYLE}'>{metal}</td>"
            f"<td style='{CELL_STYLE}text-align:center;'>"
            f"{_fmt(minmax[metal].get('min'))}</td>"
            f"<td style='{CELL_STYLE}background:{LIGHT_PINK};text-align:center;'>"
            f"{_date_fmt(minmax[metal].get('min_date'))}</td>"
            f"<td style='{CELL_STYLE}text-align:center;'>"
            f"{_fmt(minmax[metal].get('max'))}</td>"
            f"<td style='{CELL_STYLE}background:{LIGHT_PINK};text-align:center;'>"
            f"{_date_fmt(minmax[metal].get('max_date'))}</td>"
        f"</tr>"
        for metal in METAL_DISPLAY_ORDER
        if metal in minmax
    )

    # IBJA Silver 999 AM/PM rates.
    ibja_am = _get_ibja_rate(ibja_rates, "AM")
    ibja_pm = _get_ibja_rate(ibja_rates, "PM")

    # RBI exchange-rate rows.
    # FIX: Changed width to 150px to perfectly align with the top table.
    exchange_rows = "".join(
        f"<tr>"
        f"<td style='{CELL_STYLE}background:{LIGHT_PINK};width:150px;'>"
        f"INR / 1 {currency}</td>"
        f"<td colspan='2' style='{CELL_STYLE}text-align:center;'>"
        f"{_fmt_4(rate)}</td>"
        f"</tr>"
        for currency, rate in exchange_rates.items()
    )

    return f"""
    <!DOCTYPE html>
    <html>
    <body style="
        font-family:Calibri,Arial,sans-serif;
        font-size:12px;
        color:#111111;
    ">

        <p style="font-size: 15px"> Dear Team, </p>
        <p style="font-size: 15px">
            Please find below the metal prices in USD as on
            <b>{target_date:%d %b %Y}</b> with reference to LBMA website
            for Au, Ag, LPPM website for Pt, Pd &amp; JM website for Ir,
            Rh, Ru.
        </p>

        <!-- LIVE METALS PRICE TABLE -->

        <table style="
            border-collapse:collapse;
            width:350px;
        ">
            <tr>
                <th style="{CELL_STYLE}background:{LIGHT_PINK};text-align:left;width:150px;">
                    Day
                </th>
                <th colspan="2"
                    style="{CELL_STYLE}background:{LIGHT_PINK};text-align:center;">
                    {target_date:%A}
                </th>
            </tr>

            <tr>
                <th style="{CELL_STYLE}text-align:left;">
                    LIVE METALS PRICES in ($ / Toz):
                </th>
                <th colspan="2"
                    style="{CELL_STYLE}text-align:center;">
                    LBMA &amp; JM
                </th>
            </tr>

            <tr>
                <th style="{CELL_STYLE}background:{PINK};text-align:left;">
                    Date
                </th>
                <th colspan="2"
                    style="{CELL_STYLE}background:{PINK};text-align:center;">
                    {target_date:%d-%m-%Y}
                </th>
            </tr>

            <tr>
                <th style="{CELL_STYLE}text-align:left;width:150px;">Time</th>
                <th style="{CELL_STYLE}text-align:center;width:100px;">AM</th>
                <th style="{CELL_STYLE}text-align:center;width:100px;">PM</th>
            </tr>
            {price_rows}
        </table>

        <!-- IBJA SILVER RATE -->

        <p style="margin-top:18px;margin-bottom:4px;">
            <b>IBJA Silver Rate (INR/Kg)</b>
        </p>

        <table style="border-collapse:collapse;width:350px;">
            <tr>
                <th style="{CELL_STYLE}text-align:left;width:150px;">Silver999</th>
                <th style="{CELL_STYLE}text-align:center;width:100px;">AM</th>
                <th style="{CELL_STYLE}text-align:center;width:100px;">PM</th>
            </tr>
            <tr>
                <td style="{CELL_STYLE}background:{LIGHT_PINK};">
                    Silver 999
                </td>
                <td style="{CELL_STYLE}text-align:center;">
                    {_fmt(ibja_am)}
                </td>
                <td style="{CELL_STYLE}text-align:center;">
                    {_fmt(ibja_pm)}
                </td>
            </tr>
        </table>

        <!-- RBI EXCHANGE RATES -->

        <p style="margin-top:28px;margin-bottom:4px;">
            <b>Exchange Rate from RBI Rate</b>
        </p>

        <table style="border-collapse:collapse;width:350px;">
            {exchange_rows}
        </table>

        <!-- MONTHLY MINIMUM / MAXIMUM TABLE -->

        <p style="margin-top:20px; font-size: 15px">
            Please find in below table, the metal rates for
            <b>{target_date:%b %Y}</b> with Minimum &amp; Maximum rates:
        </p>
        
        <!-- Expanded width to 550px so dates don't wrap -->
        <table style="border-collapse:collapse;width:550px;">
            <tr style="background:{PINK};">
                <th rowspan="1"
                    style="{CELL_STYLE}text-align:left;width:150px;">
                    LIVE METALS PRICES<br/>
                    in ($ / Toz):
                </th>
                <th colspan="2"
                    style="{CELL_STYLE}text-align:center;width:200px;">
                    Min / Date
                </th>
                <th colspan="2"
                    style="{CELL_STYLE}text-align:center;width:200px;">
                    Max / Date
                </th>
            </tr>
            {minmax_rows}
        </table>

        <p style="margin-top:18px;color:#666666;font-size:11px;">
            This email was automatically generated by the Metal Rates
            Automator.
        </p>

    </body>
    </html>
    """

def send_email(
    html_body: str,
    target_date: date,
    attachment_path: str = None,
):
    """
    Send the report with the Excel tracker attached.

    Recipients and SMTP configuration are read from config.py.
    """
    msg = MIMEMultipart("mixed")

    msg["Subject"] = (
        f"LIVE METALS PRICES - {target_date:%d-%b-%Y} "
        f"| Min-Max Rates - {target_date:%b %Y}"
    )

    msg["From"] = config.EMAIL_FROM
    msg["To"] = ", ".join(config.EMAIL_TO)

    if getattr(config, "EMAIL_CC", None):
        msg["Cc"] = ", ".join(config.EMAIL_CC)

    # HTML message body.
    body_part = MIMEMultipart("alternative")
    body_part.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(body_part)

    # Attach the current Excel tracker, if available.
    if attachment_path:
        if os.path.isfile(attachment_path):
            with open(attachment_path, "rb") as file:
                attachment = MIMEApplication(
                    file.read(),
                    _subtype=(
                        "vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                )

            filename = os.path.basename(attachment_path)

            attachment.add_header(
                "Content-Disposition",
                "attachment",
                filename=filename,
            )

            msg.attach(attachment)
            logger.info("Attached Excel tracker: %s", filename)

        else:
            logger.warning(
                "Attachment path does not exist; skipping: %s",
                attachment_path,
            )

    recipients = (
        list(config.EMAIL_TO)
        + list(getattr(config, "EMAIL_CC", []) or [])
    )

    try:
        with smtplib.SMTP(
            config.SMTP_SERVER,
            config.SMTP_PORT,
            timeout=30,
        ) as server:

            if getattr(config, "SMTP_USE_TLS", False):
                server.starttls()

            if getattr(config, "SMTP_PASSWORD", None):
                server.login(
                    config.SMTP_USERNAME,
                    config.SMTP_PASSWORD,
                )

            server.sendmail(
                config.EMAIL_FROM,
                recipients,
                msg.as_string(),
            )

        logger.info("Metal rates email sent to %s", recipients)

    except Exception:
        logger.exception("Failed to send metal rates email")
        raise


def send_failure_email(
    stage: str,
    error: Exception,
    target_date: date,
):
    """
    Send a plain-text pipeline failure notification.
    """
    recipients = list(
        getattr(
            config,
            "FAILURE_EMAIL_TO",
            config.EMAIL_TO,
        )
    )

    subject = (
        f"FAILED - Metal Rates Automator - "
        f"{target_date:%d-%b-%Y}"
    )

    body = f"""
Metal Rates Automator pipeline failed.

Date:
{target_date:%d-%b-%Y}

Failed Stage:
{stage}

Error:
{type(error).__name__}: {error}

Please check the pipeline log for the complete traceback.

Log directory:
{config.LOG_DIR}
"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = config.EMAIL_FROM
    msg["To"] = ", ".join(recipients)

    msg.attach(MIMEText(body, "plain", "utf-8"))

    try:
        with smtplib.SMTP(
            config.SMTP_SERVER,
            config.SMTP_PORT,
            timeout=30,
        ) as server:

            if getattr(config, "SMTP_USE_TLS", False):
                server.starttls()

            if getattr(config, "SMTP_PASSWORD", None):
                server.login(
                    config.SMTP_USERNAME,
                    config.SMTP_PASSWORD,
                )

            server.sendmail(
                config.EMAIL_FROM,
                recipients,
                msg.as_string(),
            )

        logger.info(
            "Failure notification sent to %s",
            recipients,
        )

    except Exception:
        logger.exception("Failed to send failure notification")
        raise
