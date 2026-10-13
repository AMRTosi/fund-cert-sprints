import pytest

from sprint_cert_automation.utils.filenames import certificate_filename, normalize_json_output_name


def test_regular_team_filename() -> None:
    file_name = certificate_filename(2026, "Subvenciones", "SP216")
    assert file_name == "Inf_Certificacion 2026 Subvenciones SP216.xlsm"


def test_transversal_filename() -> None:
    file_name = certificate_filename(2026, "AccentureTransversal", "2026-06")
    assert file_name == "2026-06 Inf_Certificacion 2026 AccentureTranservsal.xlsm"


@pytest.mark.parametrize("name", ["report", "report.json", " report.JSON "])
def test_json_output_name_normalization(name) -> None:
    assert normalize_json_output_name(name) == "report"


@pytest.mark.parametrize("name", ["", " ", ".json", *[f"a{char}b" for char in "\\/:*?\"<>|"]])
def test_json_output_name_invalid(name) -> None:
    with pytest.raises(ValueError, match="output-file-name"):
        normalize_json_output_name(name)
