from openpyxl.cell.cell import MergedCell


def monthly_average(ws, row: int, columns):
    """Average every numeric rate recorded in the supplied columns."""
    values = []

    for col in columns:
        cell = ws.cell(row=row, column=col)

        if isinstance(cell, MergedCell):
            continue

        value = cell.value

        if isinstance(value, bool) or not isinstance(value, (int, float)):
            continue

        values.append(float(value))

    return sum(values) / len(values) if values else None
