from datetime import date

from sprint_cert_automation.domain.models import SprintWindow
from sprint_cert_automation.domain.rules import is_billable_in_month


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
