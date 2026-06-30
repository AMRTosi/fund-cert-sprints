from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from sprint_cert_automation.domain.models import CertificateDraft
from sprint_cert_automation.domain.rules import select_billable_sprints
from sprint_cert_automation.services.forecast_reader import ForecastReader
from sprint_cert_automation.services.template_writer import TemplateWriter
from sprint_cert_automation.utils.filenames import certificate_filename


@dataclass
class GenerationResult:
    generated_files: list[Path]


class CertificateGenerationService:
    def __init__(self, forecast_path: Path, template_path: Path) -> None:
        self._reader = ForecastReader(forecast_path)
        self._writer = TemplateWriter(template_path)

    def run(
        self,
        year: int,
        month: int,
        output_dir: Path,
        dry_run: bool = False,
    ) -> GenerationResult:
        sprints = self._reader.read_sprints_for_target_window(year, month)
        billable = select_billable_sprints(sprints, year, month)

        generated_files: list[Path] = []
        for sprint in billable:
            file_name = certificate_filename(year, sprint.team, sprint.sprint_id)
            holidays, workloads = self._reader.read_draft_data_for_sprint(sprint)
            draft = CertificateDraft(
                team=sprint.team,
                sprint_id=sprint.sprint_id,
                start_date=sprint.start_date,
                end_date=sprint.end_date,
                file_name=file_name,
                product_label=sprint.team,
                holidays=holidays,
                workloads=workloads,
            )
            output_path = output_dir / file_name
            if not dry_run:
                self._writer.write(draft, output_path)
            generated_files.append(output_path)

        return GenerationResult(generated_files=generated_files)
