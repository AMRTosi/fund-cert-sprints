from __future__ import annotations

from pathlib import Path

from sprint_cert_automation.infrastructure.excel_com import DEFAULT_EXPORT_MACRO_NAME
from sprint_cert_automation.services.certificate_service import (
    CertificateGenerationService,
    GenerationResult,
)
from sprint_cert_automation.services.macro_export_service import MacroExportResult, MacroExportService


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
