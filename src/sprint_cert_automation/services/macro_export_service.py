from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from sprint_cert_automation.infrastructure.excel_com import ExcelComClient


@dataclass
class MacroExportResult:
    exported_files: list[Path] = field(default_factory=list)
    skipped_files: dict[Path, str] = field(default_factory=dict)
    errors: dict[Path, str] = field(default_factory=dict)


class MacroExportService:
    def run(
        self,
        input_dir: Path,
        dry_run: bool = False,
    ) -> MacroExportResult:
        if not input_dir.exists():
            raise FileNotFoundError(f"Input directory does not exist: {input_dir}")
        if not input_dir.is_dir():
            raise NotADirectoryError(f"Input path is not a directory: {input_dir}")

        workbook_paths = sorted(
            path for path in input_dir.glob("*.xlsm") if not path.name.startswith("~$")
        )

        result = MacroExportResult()
        if dry_run:
            result.exported_files.extend(path.with_suffix(".pdf") for path in workbook_paths)
            return result

        client = ExcelComClient()
        client.open()
        try:
            for workbook_path in workbook_paths:
                workbook = None
                pdf_path = workbook_path.with_suffix(".pdf")
                try:
                    workbook = client.open_workbook(workbook_path)
                    client.export_workbook_to_pdf(workbook, pdf_path)
                    result.exported_files.append(pdf_path)
                except Exception as exc:
                    result.errors[workbook_path] = str(exc)
                finally:
                    if workbook is not None:
                        workbook.Close(SaveChanges=True)
        finally:
            client.close()

        return result