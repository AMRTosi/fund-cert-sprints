from __future__ import annotations

from pathlib import Path

from sprint_cert_automation import cli


def test_cli_export_pdf_uses_month_folder_and_macro(monkeypatch, tmp_path: Path) -> None:
    captured = {}

    def fake_default_output_dir(year: int, month: int) -> Path:
        captured["year"] = year
        captured["month"] = month
        return tmp_path

    def fake_export(input_dir: Path, macro_name: str, dry_run: bool):
        captured["input_dir"] = input_dir
        captured["macro_name"] = macro_name
        captured["dry_run"] = dry_run

        class Result:
            exported_files = []
            errors = {}

        return Result()

    monkeypatch.setattr(cli, "default_output_dir", fake_default_output_dir)
    monkeypatch.setattr(cli, "today_year_month", lambda: (2026, 6))
    monkeypatch.setattr(cli, "export_certificates_to_pdf", fake_export)

    exit_code = cli.main(["export-pdf", "--dry-run"])

    assert exit_code == 0
    assert captured["year"] == 2026
    assert captured["month"] == 6
    assert captured["input_dir"] == tmp_path
    assert captured["macro_name"] == "SELECCIONAR_HOJAS_INFORME"
    assert captured["dry_run"] is True


def test_cli_legacy_generate_invocation_is_preserved(monkeypatch, tmp_path: Path) -> None:
    captured = {}

    def fake_generate(**kwargs):
        captured.update(kwargs)

        class Result:
            generated_files = []

        return Result()

    monkeypatch.setattr(cli, "generate_certificates", fake_generate)
    monkeypatch.setattr(cli, "default_output_dir", lambda year, month: tmp_path)
    monkeypatch.setattr(cli, "today_year_month", lambda: (2026, 6))

    exit_code = cli.main(
        [
            "--forecast",
            "forecast.xlsx",
            "--template",
            "template.xlsm",
            "--year",
            "2026",
            "--month",
            "6",
            "--dry-run",
        ]
    )

    assert exit_code == 0
    assert captured["forecast_path"] == Path("forecast.xlsx")
    assert captured["template_path"] == Path("template.xlsm")
    assert captured["output_dir"] == tmp_path
    assert captured["year"] == 2026
    assert captured["month"] == 6
    assert captured["dry_run"] is True


def test_cli_export_cost_hours_json_invocation(monkeypatch, tmp_path: Path) -> None:
    captured = {}

    def fake_export(**kwargs):
        captured.update(kwargs)

        class Result:
            rows_exported = 2
            table_name = kwargs["table_name"]
            output_path = tmp_path / "out.json"

        return Result()

    monkeypatch.setattr(cli, "export_current_period_cost_hours", fake_export)
    monkeypatch.setattr(cli, "default_cost_hours_output_dir", lambda: tmp_path)

    exit_code = cli.main(
        [
            "export-cost-hours-json",
            "--forecast",
            "forecast.xlsx",
            "--table-name",
            "TableHorasCosteQuincenas202609",
            "--output-file-name",
            "2026-09",
        ]
    )

    assert exit_code == 0
    assert captured["forecast_path"] == Path("forecast.xlsx")
    assert captured["table_name"] == "TableHorasCosteQuincenas202609"
    assert captured["output_file_name"] == "2026-09"
    assert captured["output_root"] == tmp_path


def test_cli_validate_beeline_timesheets_invocation(monkeypatch, tmp_path: Path) -> None:
    captured = {}

    def fake_validate(**kwargs):
        captured.update(kwargs)

        class Result:
            validated_rows = 3
            approved_rows = 1
            rejected_rows = 1
            not_found_rows = 1
            output_path = tmp_path / "validation.json"

        return Result()

    monkeypatch.setattr(cli, "validate_beeline_timesheets", fake_validate)
    monkeypatch.setattr(cli, "default_beeline_validation_output_dir", lambda: tmp_path)

    exit_code = cli.main(
        [
            "validate-beeline-timesheets",
            "--input-json",
            "input.json",
            "--control-workbook",
            "control.xlsx",
            "--output-file-name",
            "validation_2026_10",
        ]
    )

    assert exit_code == 0
    assert captured["input_json_path"] == Path("input.json")
    assert captured["workbook_path"] == Path("control.xlsx")
    assert captured["output_file_name"] == "validation_2026_10"
    assert captured["output_root"] == tmp_path