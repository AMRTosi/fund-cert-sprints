from sprint_cert_automation.utils.filenames import certificate_filename


def test_regular_team_filename() -> None:
    file_name = certificate_filename(2026, "Subvenciones", "SP216")
    assert file_name == "Inf_Certificacion 2026 Subvenciones SP216.xlsm"


def test_transversal_filename() -> None:
    file_name = certificate_filename(2026, "AccentureTransversal", "2026-06")
    assert file_name == "2026-06 Inf_Certificacion 2026 AccentureTranservsal.xlsm"
