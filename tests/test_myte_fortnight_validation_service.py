from __future__ import annotations

import json
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table

from sprint_cert_automation.app import validate_myte_fortnight
from sprint_cert_automation.services.cost_hours_export_service import CostHoursExportService
from sprint_cert_automation.services.myte_fortnight_validation_service import MyteFortnightValidationService


@pytest.fixture
def input_paths(tmp_path):
    excel = tmp_path / "excel.json"
    myte = tmp_path / "myte.json"
    excel.write_text(json.dumps([{
        "reviewee": "alice", "email": "alice@example.com", "subco": "0",
        "horasQuincena1": "80", "horasQuincena2": "72",
    }]), encoding="utf-8")
    myte.write_text(json.dumps([{
        "reviewee": "alice", "trStatus": "Processed", "chargeCode": "COMK3001",
        "workingHours": "80", "totalHours": "67.2",
        "percentageTotalHoursVsWorkingHours": "84%",
    }]), encoding="utf-8")
    return excel, myte


def test_app_runs_service_and_preserves_inputs(input_paths, tmp_path) -> None:
    excel, myte = input_paths
    before = [p.read_bytes() for p in input_paths]
    result = validate_myte_fortnight(excel, myte, 1, " report.JSON ", tmp_path / "out")
    assert result.output_path == tmp_path / "out" / "report.json"
    assert (result.validated_rows, result.passed_rows, result.failed_rows) == (1, 1, 0)
    assert json.loads(result.output_path.read_text(encoding="utf-8"))[0]["validationStatus"] == "Passed"
    assert result.output_path.read_bytes().endswith(b"\n")
    assert before == [p.read_bytes() for p in input_paths]
    assert list(result.output_path.parent.iterdir()) == [result.output_path]


@pytest.mark.parametrize("payload", ["invalid json", "{}", "[{}]", '[{"reviewee": 1}]'])
def test_invalid_input_does_not_replace_previous_report(input_paths, tmp_path, payload) -> None:
    excel, myte = input_paths
    excel.write_text(payload, encoding="utf-8")
    output = tmp_path / "report.json"
    output.write_text("previous", encoding="utf-8")
    with pytest.raises(ValueError):
        validate_myte_fortnight(excel, myte, 1, "report", tmp_path)
    assert output.read_text() == "previous"
    assert not list(tmp_path.glob("*.tmp"))


@pytest.mark.parametrize("name", ["", " ", ".json", "..\\report", "a/b", "a:b"])
def test_invalid_output_names(input_paths, tmp_path, name) -> None:
    with pytest.raises(ValueError, match="output-file-name"):
        validate_myte_fortnight(*input_paths, 1, name, tmp_path)


@pytest.mark.parametrize("period", [0, 3])
def test_service_rejects_invalid_period_before_io(tmp_path, period) -> None:
    with pytest.raises(ValueError, match="period-number"):
        MyteFortnightValidationService().run(tmp_path / "missing", tmp_path / "missing2", period, "report", tmp_path)


@pytest.mark.parametrize("name", ["excel", "myte"])
def test_cannot_overwrite_input(input_paths, tmp_path, name) -> None:
    before = [p.read_bytes() for p in input_paths]
    with pytest.raises(ValueError, match="overwrite"):
        validate_myte_fortnight(*input_paths, 1, name, tmp_path)
    assert before == [p.read_bytes() for p in input_paths]


def test_missing_file_and_output_directory_errors(input_paths, tmp_path) -> None:
    with pytest.raises(FileNotFoundError):
        validate_myte_fortnight(tmp_path / "missing.json", input_paths[1], 1, "report", tmp_path)
    blocked = tmp_path / "blocked"
    blocked.write_text("not a directory")
    with pytest.raises(OSError):
        validate_myte_fortnight(*input_paths, 1, "report", blocked)


def test_replace_failure_preserves_previous_report_and_removes_temporary(input_paths, tmp_path, monkeypatch) -> None:
    output = tmp_path / "report.json"
    output.write_text("previous")

    def denied_replace(self, target):
        raise PermissionError("locked output")

    monkeypatch.setattr(Path, "replace", denied_replace)
    with pytest.raises(PermissionError, match="locked"):
        validate_myte_fortnight(*input_paths, 1, "report", tmp_path)
    assert output.read_text() == "previous"
    assert not list(tmp_path.glob("*.tmp"))


def test_serialization_failure_cleans_partial_temporary_file(input_paths, tmp_path, monkeypatch) -> None:
    def failed_dump(payload, stream, **kwargs):
        stream.write("[")
        raise OSError("disk full")

    monkeypatch.setattr(json, "dump", failed_dump)
    with pytest.raises(OSError, match="disk full"):
        validate_myte_fortnight(*input_paths, 1, "report", tmp_path / "out")
    assert list((tmp_path / "out").iterdir()) == []


def test_later_numeric_error_does_not_publish_partial_results(input_paths, tmp_path) -> None:
    excel, myte = input_paths
    a = json.loads(excel.read_text())
    b = json.loads(myte.read_text())
    a.append(dict(a[0], reviewee="bob", horasQuincena1="invalid"))
    b.append(dict(b[0], reviewee="bob"))
    excel.write_text(json.dumps(a))
    myte.write_text(json.dumps(b))
    with pytest.raises(ValueError, match="Valor numerico invalido"):
        validate_myte_fortnight(excel, myte, 1, "report", tmp_path / "out")
    assert not (tmp_path / "out").exists()


def test_empty_report_published(input_paths, tmp_path) -> None:
    for path in input_paths:
        path.write_text("[]")
    result = validate_myte_fortnight(*input_paths, 2, "report", tmp_path / "out")
    assert (result.validated_rows, result.passed_rows, result.failed_rows) == (0, 0, 0)
    assert json.loads(result.output_path.read_text()) == []


def test_unknown_encoding_is_explicit_error(input_paths, tmp_path) -> None:
    with pytest.raises(LookupError, match="unknown encoding"):
        validate_myte_fortnight(*input_paths, 1, "report", tmp_path / "out", encoding="nonexistent-codec")
    assert not (tmp_path / "out").exists()


def test_original_encoding_and_indent_options(input_paths, tmp_path) -> None:
    excel, myte = input_paths
    payload = json.loads(excel.read_text(encoding="utf-8"))
    payload[0]["email"] = "jos\u00e9@example.com"
    excel.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-16")
    myte.write_text(myte.read_text(encoding="utf-8"), encoding="utf-16")
    result = validate_myte_fortnight(excel, myte, 1, "report", tmp_path / "out", encoding="utf-16", indent=4)
    text = result.output_path.read_text(encoding="utf-16")
    assert '\n    {' in text
    assert json.loads(text)[0]["email"] == "jos\u00e9@example.com"


@pytest.mark.parametrize("period,working", [(1, "80"), (2, "72")])
def test_export_cost_hours_to_myte_end_to_end(tmp_path, period, working) -> None:
    workbook_path = tmp_path / "forecast.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["reviewee", "email", "subco", "horasQuincena1", "horasQuincena2"])
    worksheet.append(["alice", "alice@example.com", 0, 80, 72])
    worksheet.append(["excluded", "excluded@example.com", 1, 80, 72])
    worksheet.add_table(Table(displayName="TableHorasCosteQuincenas202609", ref="A1:E3"))
    workbook.save(workbook_path)
    workbook.close()
    before = workbook_path.read_bytes()
    exported = CostHoursExportService().run(
        workbook_path, "TableHorasCosteQuincenas202609", "cost", tmp_path / "cost",
    )
    myte = tmp_path / "myte.json"
    myte.write_text(json.dumps([{
        "reviewee": " ALICE ", "trStatus": "Processed", "chargeCode": "COMK3002",
        "workingHours": working, "totalHours": "unused",
        "percentageTotalHoursVsWorkingHours": "92%",
    }]), encoding="utf-8")
    result = validate_myte_fortnight(exported.output_path, myte, period, "result", tmp_path / "validation")
    assert (result.validated_rows, result.passed_rows, result.failed_rows) == (1, 1, 0)
    assert workbook_path.read_bytes() == before
    assert json.loads(result.output_path.read_text())[0]["reviewee"] == "alice"
