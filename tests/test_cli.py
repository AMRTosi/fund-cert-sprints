from __future__ import annotations

from pathlib import Path
import json

import pytest

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


def test_cli_myte_invocation_and_summary(monkeypatch, tmp_path, capsys) -> None:
    from sprint_cert_automation.services.myte_fortnight_validation_service import MyteValidationRunResult

    captured = {}

    def fake_validate(**kwargs):
        captured.update(kwargs)
        return MyteValidationRunResult(
            kwargs["excel_json_path"], kwargs["myte_json_path"], tmp_path / "result.json", 2, 1, 1,
        )

    monkeypatch.setattr(cli, "validate_myte_fortnight", fake_validate)
    monkeypatch.setattr(cli, "default_myte_validation_output_dir", lambda: tmp_path)
    assert cli.main([
        "validate-myte-fortnight", "--excel-json", "excel.json", "--myte-json", "myte.json",
        "--period-number", "2", "--output-file-name", "result", "--pretty-log",
        "--encoding", "utf-16", "--indent", "4",
    ]) == 0
    assert captured == {
        "excel_json_path": Path("excel.json"), "myte_json_path": Path("myte.json"),
        "period_number": 2, "output_file_name": "result", "output_root": tmp_path,
        "encoding": "utf-16", "indent": 4,
    }
    assert json.loads(capsys.readouterr().out) == {"total": 2, "passed": 1, "failed": 1}


@pytest.mark.parametrize("period", ["0", "3", "invalid"])
def test_cli_myte_invalid_period(period) -> None:
    with pytest.raises(SystemExit) as exc:
        cli.main([
            "validate-myte-fortnight", "--excel-json", "excel.json", "--myte-json", "myte.json",
            "--period-number", period, "--output-file-name", "result",
        ])
    assert exc.value.code == 2


@pytest.mark.parametrize("missing", ["--excel-json", "--myte-json", "--period-number", "--output-file-name"])
def test_cli_myte_required_arguments(missing) -> None:
    args = [
        "validate-myte-fortnight", "--excel-json", "excel.json", "--myte-json", "myte.json",
        "--period-number", "1", "--output-file-name", "result",
    ]
    index = args.index(missing)
    del args[index:index + 2]
    with pytest.raises(SystemExit) as exc:
        cli.main(args)
    assert exc.value.code == 2


@pytest.mark.parametrize("error", [
    FileNotFoundError("missing file"), ValueError("invalid contract"),
    PermissionError("locked output"), LookupError("unknown encoding"),
    UnicodeDecodeError("utf-8", b"\xff", 0, 1, "invalid encoding"),
])
def test_cli_myte_reports_known_errors(monkeypatch, capsys, error) -> None:
    def failed_validate(**kwargs):
        raise error

    monkeypatch.setattr(cli, "validate_myte_fortnight", failed_validate)
    assert cli.main([
        "validate-myte-fortnight", "--excel-json", "excel.json", "--myte-json", "myte.json",
        "--period-number", "1", "--output-file-name", "result",
    ]) == 1
    captured = capsys.readouterr()
    assert "Error durante la validacion myTE:" in captured.err
    assert captured.out == ""


def test_cli_myte_business_failure_is_successful_execution(monkeypatch, tmp_path, capsys) -> None:
    excel, myte = tmp_path / "excel.json", tmp_path / "myte.json"
    excel.write_text(json.dumps([{
        "reviewee": "alice", "email": "alice@example.com", "subco": "0",
        "horasQuincena1": "80", "horasQuincena2": "72",
    }]))
    myte.write_text("[]")
    monkeypatch.setattr(cli, "default_myte_validation_output_dir", lambda: tmp_path / "out")
    assert cli.main([
        "validate-myte-fortnight", "--excel-json", str(excel), "--myte-json", str(myte),
        "--period-number", "1", "--output-file-name", "result",
    ]) == 0
    assert "Failed: 1" in capsys.readouterr().out
    assert json.loads((tmp_path / "out" / "result.json").read_text()) == [{
        "reviewee": "alice", "email": "alice@example.com", "validationStatus": "Failed",
        "horasQuincena1": "80", "horasQuincena2": "72",
        "trStatus": "", "chargeCode": "", "totalHours": "", "workingHours": "",
        "percentageTotalHoursVsWorkingHours": "",
        "comments": "El reviewee registrado en el fichero excel no ha sido encontrado en myTE",
        "sendNotification": False, "notificationMessage": "",
    }]


def test_myte_default_output_is_project_relative() -> None:
    project = Path(cli.__file__).resolve().parents[2]
    assert cli.default_myte_validation_output_dir() == project / "out" / "myte-fortnight-validation"


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