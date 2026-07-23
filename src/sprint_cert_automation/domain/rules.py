from __future__ import annotations

from datetime import date
from typing import Iterable

from .models import SprintWindow


def is_billable_in_month(sprint: SprintWindow, year: int, month: int) -> bool:
    """A sprint is billable if it ends in the target month and is not hatched."""
    return (
        sprint.end_date.year == year
        and sprint.end_date.month == month
        and not sprint.is_hatched
    )


def select_billable_sprints(
    sprints: list[SprintWindow],
    year: int,
    month: int,
) -> list[SprintWindow]:
    return [s for s in sprints if is_billable_in_month(s, year, month)]


def day_hours_for_date(day: date) -> float:
    """Return nominal hours for a calendar day according to summer schedule."""
    if (day.month == 6 and day.day >= 15) or day.month in {7, 8} or (day.month == 9 and day.day <= 14):
        return 7.5
    return 8.5


def sprint_hours_between_dates(start_date: date, end_date: date) -> float:
    """Sum nominal hours for weekdays in an inclusive sprint window."""
    if end_date < start_date:
        raise ValueError("end_date cannot be before start_date")

    total_hours = 0.0
    cursor = start_date
    while cursor <= end_date:
        if cursor.weekday() < 5:
            total_hours += day_hours_for_date(cursor)
        cursor = date.fromordinal(cursor.toordinal() + 1)
    return total_hours


def free_hours_from_non_working_dates(non_working_dates: Iterable[date]) -> float:
    """Sum free hours from weekday non-working dates using seasonal nominal hours."""
    total_hours = 0.0
    for non_working_date in non_working_dates:
        if non_working_date.weekday() < 5:
            total_hours += day_hours_for_date(non_working_date)
    return total_hours


def free_hours_from_non_working_days(non_working_days: int, day_hours: float = 8.5) -> float:
    if non_working_days < 0:
        raise ValueError("non_working_days cannot be negative")
    return non_working_days * day_hours


def sprint_hours(day_count: int, day_hours: float = 8.5) -> float:
    if day_count < 0:
        raise ValueError("day_count cannot be negative")
    return day_count * day_hours
