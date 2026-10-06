from __future__ import annotations

from pathlib import Path

from sprint_cert_automation.infrastructure.excel_com import DEFAULT_EXPORT_MACRO_NAME
from sprint_cert_automation.services.certificate_service import (
    CertificateGenerationService,
    GenerationResult,
)
from sprint_cert_automation.services.cost_hours_export_service import (
    CostHoursExportResult,
    CostHoursExportService,
)
from sprint_cert_automation.services.beeline_timesheet_validation_service import (
    BeelineTimesheetValidationService,
    BeelineValidationRunResult,
)
from sprint_cert_automation.services.macro_export_service import MacroExportResult, MacroExportService
from sprint_cert_automation.services.sheet_duplicator import DuplicateSheetResult, duplicate_sheet


def generate_certificates(
    forecast_path: Path,
    template_path: Path,
    output_dir: Path,
    year: int,
    month: int,
    dry_run: bool = False,
) -> GenerationResult:
    service = CertificateGenerationService(
        forecast_path=forecast_path,
        template_path=template_path,
    )
    return service.run(
        year=year,
        month=month,
        output_dir=output_dir,
        dry_run=dry_run,
    )


def export_certificates_to_pdf(
    input_dir: Path,
    macro_name: str = DEFAULT_EXPORT_MACRO_NAME,
    dry_run: bool = False,
) -> MacroExportResult:
    service = MacroExportService(macro_name=macro_name)
    return service.run(
        input_dir=input_dir,
        dry_run=dry_run,
    )


def duplicate_period_sheet(
    forecast_path: Path,
    source_sheet: str,
    new_sheet: str,
    year: int,
    month: int,
    previous_sheet: str | None = None,
    dry_run: bool = False,
) -> DuplicateSheetResult:
    return duplicate_sheet(
        workbook_path=forecast_path,
        source_sheet_name=source_sheet,
        new_sheet_name=new_sheet,
        year=year,
        month=month,
        previous_sheet_name=previous_sheet,
        dry_run=dry_run,
    )


def export_current_period_cost_hours(
    forecast_path: Path,
    table_name: str,
    output_file_name: str,
    output_root: Path,
) -> CostHoursExportResult:
    service = CostHoursExportService()
    return service.run(
        workbook_path=forecast_path,
        table_name=table_name,
        output_file_name=output_file_name,
        output_root=output_root,
    )


def validate_beeline_timesheets(
    workbook_path: Path,
    input_json_path: Path,
    output_file_name: str,
    output_root: Path,
) -> BeelineValidationRunResult:
    service = BeelineTimesheetValidationService()
    return service.run(
        workbook_path=workbook_path,
        input_json_path=input_json_path,
        output_file_name=output_file_name,
        output_root=output_root,
    )
