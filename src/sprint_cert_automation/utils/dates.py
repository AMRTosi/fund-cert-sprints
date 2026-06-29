from __future__ import annotations

from datetime import date


def previous_month(year: int, month: int) -> tuple[int, int]:
    if month < 1 or month > 12:
        raise ValueError("month must be in range 1..12")
    if month == 1:
        return year - 1, 12
    return year, month - 1


def month_name_es(month: int) -> str:
    names = {
        1: "Enero",
        2: "Febrero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre",
    }
    if month not in names:
        raise ValueError("month must be in range 1..12")
    return names[month]


def year_month_label(year: int, month: int) -> str:
    return f"{year:04d}-{month:02d}"


def today_year_month() -> tuple[int, int]:
    today = date.today()
    return today.year, today.month
