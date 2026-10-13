from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
from tempfile import NamedTemporaryFile

from sprint_cert_automation.domain.myte_validation import validate_payloads, validate_period_number
from sprint_cert_automation.utils.filenames import normalize_json_output_name


@dataclass(frozen=True)
class MyteValidationRunResult:
    excel_json_path: Path
    myte_json_path: Path
    output_path: Path
    validated_rows: int
    passed_rows: int
    failed_rows: int


class MyteFortnightValidationService:
    def run(
        self,
        excel_json_path: Path,
        myte_json_path: Path,
        period_number: int,
        output_file_name: str,
        output_root: Path,
        encoding: str = "utf-8",
        indent: int = 2,
    ) -> MyteValidationRunResult:
        validate_period_number(period_number)
        name = normalize_json_output_name(output_file_name)
        output_path = output_root / f"{name}.json"
        if output_path.resolve() in {excel_json_path.resolve(), myte_json_path.resolve()}:
            raise ValueError("Output JSON must not overwrite an input file")

        with excel_json_path.open("r", encoding=encoding) as source:
            excel_payload = json.load(source)
        with myte_json_path.open("r", encoding=encoding) as source:
            myte_payload = json.load(source)
        report = validate_payloads(excel_payload, myte_payload, period_number)
        output_root.mkdir(parents=True, exist_ok=True)
        temporary_path: Path | None = None
        try:
            with NamedTemporaryFile(
                mode="w", encoding=encoding, dir=output_root, suffix=".tmp", delete=False,
            ) as destination:
                temporary_path = Path(destination.name)
                json.dump(report, destination, ensure_ascii=False, indent=indent)
                destination.write("\n")
            temporary_path.replace(output_path)
        finally:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)

        passed = sum(row["validationStatus"] == "Passed" for row in report)
        return MyteValidationRunResult(
            excel_json_path=excel_json_path,
            myte_json_path=myte_json_path,
            output_path=output_path,
            validated_rows=len(report),
            passed_rows=passed,
            failed_rows=len(report) - passed,
        )
