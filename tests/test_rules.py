from datetime import date

from sprint_cert_automation.domain.models import SprintWindow
from sprint_cert_automation.domain.rules import (
    day_hours_for_date,
    free_hours_from_non_working_dates,
    is_billable_in_month,
    sprint_hours_between_dates,
)


def test_billable_when_end_date_matches_target_month() -> None:
    sprint = SprintWindow(
        team="Bonificaciones",
        sprint_id="SP277",
        start_date=date(2026, 5, 28),
        end_date=date(2026, 6, 10),
        source_sheet="FY26_jun",
        is_hatched=False,
    )
    assert is_billable_in_month(sprint, 2026, 6)


def test_not_billable_when_hatched() -> None:
    sprint = SprintWindow(
        team="Subvenciones",
        sprint_id="SP216",
        start_date=date(2026, 6, 1),
        end_date=date(2026, 6, 14),
        source_sheet="FY26_jun",
        is_hatched=True,
    )
    assert not is_billable_in_month(sprint, 2026, 6)


def test_day_hours_for_date_uses_summer_boundaries() -> None:
    assert day_hours_for_date(date(2026, 6, 14)) == 8.5
    assert day_hours_for_date(date(2026, 6, 15)) == 7.5
    assert day_hours_for_date(date(2026, 9, 14)) == 7.5
    assert day_hours_for_date(date(2026, 9, 15)) == 8.5


def test_sprint_hours_between_dates_sums_weekdays_with_mixed_rates() -> None:
    # 14-Jun is Sunday (excluded), 15-Jun (7.5) + 16-Jun (7.5)
    assert sprint_hours_between_dates(date(2026, 6, 14), date(2026, 6, 16)) == 15.0


def test_free_hours_from_non_working_dates_uses_per_day_rates() -> None:
    # 14-Jun is Sunday (excluded), only 16-Jun contributes 7.5
    non_working_dates = [date(2026, 6, 14), date(2026, 6, 16)]
    assert free_hours_from_non_working_dates(non_working_dates) == 7.5
