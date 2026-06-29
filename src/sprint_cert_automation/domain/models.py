from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date


@dataclass(frozen=True)
class SprintWindow:
    team: str
    sprint_id: str
    start_date: date
    end_date: date
    source_sheet: str
    is_hatched: bool = False


@dataclass(frozen=True)
class TeamMember:
    name: str
    billing_line: str
    category: str
    team: str


@dataclass(frozen=True)
class TeamMemberWorkload:
    member: TeamMember
    sprint_hours: float
    free_hours: float


@dataclass(frozen=True)
class Holiday:
    holiday_date: date
    label: str


@dataclass
class CertificateDraft:
    team: str
    sprint_id: str
    start_date: date
    end_date: date
    file_name: str
    product_label: str
    holidays: list[Holiday] = field(default_factory=list)
    workloads: list[TeamMemberWorkload] = field(default_factory=list)
