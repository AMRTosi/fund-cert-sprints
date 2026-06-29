from __future__ import annotations

from pathlib import Path

from sprint_cert_automation.services.certificate_service import (
    CertificateGenerationService,
    GenerationResult,
)


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
    zip_name = f"certificaciones_{year:04d}_{month:02d}.zip"
    return service.run(
        year=year,
        month=month,
        output_dir=output_dir,
        zip_name=zip_name,
        dry_run=dry_run,
    )
