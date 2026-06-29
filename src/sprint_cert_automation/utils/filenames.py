from __future__ import annotations


def normalize_team_label(team: str) -> str:
    return " ".join(team.split())


def certificate_filename(year: int, team: str, sprint_id: str) -> str:
    team_label = normalize_team_label(team)
    if team_label.lower() == "accenturetransversal":
        return f"{sprint_id} Inf_Certificacion {year} AccentureTranservsal.xlsm"
    return f"Inf_Certificacion {year} {team_label} {sprint_id}.xlsm"
