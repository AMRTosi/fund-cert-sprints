from __future__ import annotations


def normalize_json_output_name(output_file_name: str) -> str:
    name = output_file_name.strip()
    if name.lower().endswith(".json"):
        name = name[:-5]
    if not name:
        raise ValueError("output-file-name must not be empty")
    if any(char in name for char in "\\/:*?\"<>|"):
        raise ValueError("output-file-name contains invalid filesystem characters")
    return name


def normalize_team_label(team: str) -> str:
    return " ".join(team.split())


def certificate_filename(year: int, team: str, sprint_id: str) -> str:
    team_label = normalize_team_label(team)
    if team_label.lower() == "accenturetransversal":
        return f"{sprint_id} Inf_Certificacion {year} AccentureTranservsal.xlsm"
    return f"Inf_Certificacion {year} {team_label} {sprint_id}.xlsm"
