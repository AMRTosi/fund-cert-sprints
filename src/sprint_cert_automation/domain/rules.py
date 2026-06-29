from __future__ import annotations

from datetime import date

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


def free_hours_from_non_working_days(non_working_days: int, day_hours: float = 8.5) -> float:
    if non_working_days < 0:
        raise ValueError("non_working_days cannot be negative")
    return non_working_days * day_hours


def sprint_hours(day_count: int, day_hours: float = 8.5) -> float:
    if day_count < 0:
        raise ValueError("day_count cannot be negative")
    return day_count * day_hours
