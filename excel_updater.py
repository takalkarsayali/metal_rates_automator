import logging
from datetime import date, datetime
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from monthly_average import monthly_average
import config

logger = logging.getLogger(__name__)


MONTH_SHEET_FMT = "%b %Y"
SESSION_OFFSETS = {"AM": 0, "PM": 1}


def _month_sheet_name(d: date) -> str:
    name = d.strftime(MONTH_SHEET_FMT)
    if name.startswith("Sep "):
        name = "Sept " + name[4:]
    return name


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def _merged_anchor(ws, row, col):
    """Return the writable anchor for a merged cell."""
    cell = ws.cell(row=row, column=col)

    if not isinstance(cell, MergedCell):
        return cell

    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(
                row=merged_range.min_row,
                column=merged_range.min_col,
            )

    raise ValueError(f"Cannot find merged-cell anchor for {cell.coordinate}")


def _set_value(ws, row, col, value, number_format=None):
    """Write only to a merged range's anchor cell."""
    cell = _merged_anchor(ws, row, col)
    cell.value = value

    if number_format:
        cell.number_format = number_format

    return cell


def _day_start_columns(ws, layout):
    """Daily columns start at C, E, G...; each day occupies AM + PM."""
    first = layout["first_data_col"]
    col = first

    while col <= ws.max_column:
        yield col
        col += 2


def _find_day_start(ws, target_date, layout):
    for col in _day_start_columns(ws, layout):
        existing = _as_date(
            _merged_anchor(ws, layout["date_row"], col).value
        )

        if existing == target_date:
            return col

    return None


def _find_next_empty_day_start(ws, layout):
    for col in _day_start_columns(ws, layout):
        value = _merged_anchor(
            ws, layout["date_row"], col
        ).value

        if value is None:
            return col

    return ((ws.max_column - layout["first_data_col"]) // 2 + 1) * 2 + layout["first_data_col"]


def _weekday_label(d):
    return d.strftime("%A")


def _price_format(metal):
    return config.PRICE_NUMBER_FORMATS.get(
        metal, config.DEFAULT_PRICE_NUMBER_FORMAT
    )


def _create_month_sheet(wb, sheet_name, layout):
    if not wb.sheetnames:
        raise RuntimeError("Workbook has no template sheet.")

    template_ws = wb[wb.sheetnames[-1]]
    ws = wb.copy_worksheet(template_ws)
    ws.title = sheet_name

    # Clear old daily values without writing into non-anchor merged cells.
    clear_rows = {
        layout["weekday_row"],
        layout["source_row"],
        layout["date_row"],
        layout["time_row"],
        *layout["metal_rows"].values(),
    }

    # Also clear the optional daily IBJA and RBI rows.
    if "IBJA_Rate" in layout:
        clear_rows.update(layout["IBJA_Rate"].values())

    if "Exchange_Rate" in layout:
        clear_rows.update(layout["Exchange_Rate"].values())

    for row in clear_rows:
        for col in range(layout["first_data_col"], template_ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if isinstance(cell, MergedCell):
                continue
            cell.value = None

    # Clear monthly summary values.
    for row in layout["metal_minmax_rows"].values():
        for key in (
            "minmax_min_val_col",
            "minmax_min_date_col",
            "minmax_max_val_col",
            "minmax_max_date_col",
        ):
            ws.cell(row=row, column=layout[key]).value = None

        avg_col = layout.get("monthly_average_col")
        if avg_col:
            ws.cell(row=row, column=avg_col).value = None

    logger.info("Created new month sheet: %s", sheet_name)
    return ws


def update_excel(
    prices: dict,
    target_date: date = None,
    session: str = "PM",
    ibja_rate=None,
    exchange_rates=None,
) -> str:
    """
    Update a day's AM or PM data.

    prices:
        {"GOLD": ..., "SILVER": ..., "PLATINUM": ..., ...}

    session:
        "AM" or "PM"

    ibja_rate:
        Optional IBJA Silver 999 rate in INR/kg.

    exchange_rates:
        Optional dict, e.g. {"USD_INR": 95.1234, "EUR_INR": 106.1234}

    Notes:
        Rhodium, Iridium and Ruthenium are merged across AM/PM.
        Their daily JM rate is written to the merged cell anchor.
    """
    target_date = target_date or date.today()
    session = session.upper()

    if session not in SESSION_OFFSETS:
        raise ValueError("session must be 'AM' or 'PM'")

    exchange_rates = exchange_rates or {}
    layout = config.LAYOUT

    wb = openpyxl.load_workbook(config.EXCEL_PATH)
    sheet_name = _month_sheet_name(target_date)

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = _create_month_sheet(wb, sheet_name, layout)

    day_start = _find_day_start(ws, target_date, layout)

    if day_start is None:
        day_start = _find_next_empty_day_start(ws, layout)

    session_col = day_start + SESSION_OFFSETS[session]

    # These header rows are merged across the day's AM/PM pair.
    _set_value(ws, layout["weekday_row"], day_start, _weekday_label(target_date))
    _set_value(ws, layout["source_row"], day_start, "LBMA & JM")
    _set_value(
        ws,
        layout["date_row"],
        day_start,
        target_date,
        config.DATE_NUMBER_FORMAT,
    )

    # AM/PM labels are separate cells.
    _set_value(ws, layout["time_row"], session_col, session)

    for metal, row in layout["metal_rows"].items():
        if metal not in prices:
            logger.warning("No price supplied for %s", metal)
            continue

        # For merged JM cells, this resolves to the shared daily cell.
        cell = _set_value(
            ws,
            row,
            session_col,
            prices[metal],
            _price_format(metal),
        )

        logger.info(
            "%s %s written to %s",
            metal, session, cell.coordinate,
        )

    # IBJA Silver row: write to the selected AM/PM column.
    ibja_row = layout.get("IBJA_Rate", {}).get("Silver999_row")

    if ibja_row and ibja_rate is not None:
        _set_value(
            ws,
            ibja_row,
            session_col,
            ibja_rate,
            "0.00",
        )

    # RBI rates: write to the selected AM/PM column.
    exchange_layout = layout.get("Exchange_Rate", {})

    exchange_row_map = {
        "USD_INR": exchange_layout.get("INR_USD_row"),
        "EUR_INR": exchange_layout.get("INR_EUR_row"),
    }

    for key, row in exchange_row_map.items():
        if row and key in exchange_rates:
            _set_value(
                ws,
                row,
                session_col,
                exchange_rates[key],
                "0.0000",
            )

    _update_minmax_block(ws, layout)

    try:
        wb.save(config.EXCEL_PATH)
    except PermissionError as exc:
        logger.error(
            "Cannot save workbook: %s. Close it in Excel and retry.",
            config.EXCEL_PATH,
        )
        raise RuntimeError(
            "Excel workbook is locked. Close the workbook in Excel "
            "and rerun the script."
        ) from exc

        logger.info(
            "Saved %s; sheet=%s; date=%s; session=%s; column=%s",
            config.EXCEL_PATH,
            sheet_name,
            target_date,
            session,
            get_column_letter(session_col),
        )

        return config.EXCEL_PATH

def _update_minmax_block(ws, layout):
    """Recalculate monthly minimum, maximum and average for each metal."""
    from monthly_average import monthly_average

    for metal, price_row in layout["metal_rows"].items():
        summary_row = layout["metal_minmax_rows"][metal]
        values_with_dates = []
        average_columns = []
        seen_cells = set()

        # Collect AM and PM columns for every day.
        for day_start in _day_start_columns(ws, layout):
            date_value = _as_date(
                _merged_anchor(
                    ws, layout["date_row"], day_start
                ).value
            )

            for col in (day_start, day_start + 1):
                cell = _merged_anchor(ws, price_row, col)

                # Merged JM cells must count only once per day.
                if cell.coordinate in seen_cells:
                    continue
                seen_cells.add(cell.coordinate)

                value = cell.value

                if (
                    date_value is not None
                    and isinstance(value, (int, float))
                    and not isinstance(value, bool)
                ):
                    values_with_dates.append((value, date_value))
                    average_columns.append(col)

        price_fmt = _price_format(metal)

        if values_with_dates:
            min_val, min_date = min(
                values_with_dates, key=lambda item: item[0]
            )
            max_val, max_date = max(
                values_with_dates, key=lambda item: item[0]
            )

            _set_value(
                ws,
                summary_row,
                layout["minmax_min_val_col"],
                min_val,
                price_fmt,
            )
            _set_value(
                ws,
                summary_row,
                layout["minmax_min_date_col"],
                min_date,
                config.DATE_NUMBER_FORMAT,
            )
            _set_value(
                ws,
                summary_row,
                layout["minmax_max_val_col"],
                max_val,
                price_fmt,
            )
            _set_value(
                ws,
                summary_row,
                layout["minmax_max_date_col"],
                max_date,
                config.DATE_NUMBER_FORMAT,
            )

        else:
            # Clear stale summary values when no rates are available.
            for key in (
                "minmax_min_val_col",
                "minmax_min_date_col",
                "minmax_max_val_col",
                "minmax_max_date_col",
            ):
                _set_value(
                    ws,
                    summary_row,
                    layout[key],
                    None,
                )

        # Average every numeric rate, not just the min and max.
        average = monthly_average(
            ws,
            price_row,
            average_columns,
        )

        average_col = layout.get("monthly_average_col", 7)

        _set_value(
            ws,
            summary_row,
            average_col,
            average,
            price_fmt,
        )

    logger.info("Monthly Min/Max/Average refreshed.")

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    sample_prices = {
        "GOLD": 1111.0,
        "PLATINUM": 1111.0,
        "PALLADIUM": 1111.0,
        "SILVER": 11.5,
        "RHODIUM": 1111.0,
        "IRIDIUM": 1111.0,
        "RUTHENIUM": 1111.0,
    }

    update_excel(sample_prices, session="PM")