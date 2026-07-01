from __future__ import annotations

from pathlib import Path

import pytest

from sprint_cert_automation.services import macro_export_service
from sprint_cert_automation.services.macro_export_service import MacroExportService


def test_macro_export_service_dry_run_lists_pdf_targets(tmp_path: Path) -> None:
    (tmp_path / "A.xlsm").write_bytes(b"")
    (tmp_path / "B.xlsm").write_bytes(b"")
    (tmp_path / "~$lock.xlsm").write_bytes(b"")
    (tmp_path / "ignore.xlsx").write_bytes(b"")

    result = MacroExportService().run(input_dir=tmp_path, dry_run=True)

    assert result.errors == {}
    assert result.exported_files == [
        tmp_path / "A.pdf",
        tmp_path / "B.pdf",
    ]


def test_macro_export_service_raises_for_missing_directory(tmp_path: Path) -> None:
    missing = tmp_path / "does-not-exist"
    with pytest.raises(FileNotFoundError):
        MacroExportService().run(input_dir=missing, dry_run=True)


def test_macro_export_service_runs_macro_and_exports_pdf(monkeypatch, tmp_path: Path) -> None:
    first = tmp_path / "A.xlsm"
    second = tmp_path / "B.xlsm"
    first.write_bytes(b"")
    second.write_bytes(b"")

    events: list[tuple[str, str, str]] = []

    class FakeWorkbook:
        def __init__(self, path: Path) -> None:
            self.path = path

        def Close(self, SaveChanges=True) -> None:  # noqa: N803
            events.append(("close", self.path.name, str(SaveChanges)))

    class FakeExcelComClient:
        def open(self) -> None:
            events.append(("open", "", ""))

        def close(self) -> None:
            events.append(("quit", "", ""))

        def open_workbook(self, workbook_path: Path):
            return FakeWorkbook(workbook_path)

        def export_workbook_to_pdf(self, workbook, output_path: Path) -> None:
            events.append(("pdf", workbook.path.name, output_path.name))
            if workbook.path.name == "B.xlsm":
                raise RuntimeError("pdf failure")

    monkeypatch.setattr(macro_export_service, "ExcelComClient", FakeExcelComClient)

    result = MacroExportService().run(input_dir=tmp_path)

    assert result.exported_files == [tmp_path / "A.pdf"]
    assert result.errors == {second: "pdf failure"}
    assert ("pdf", "A.xlsm", "A.pdf") in events