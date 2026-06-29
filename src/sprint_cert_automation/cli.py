from __future__ import annotations

import argparse
from pathlib import Path

from sprint_cert_automation.app import generate_certificates
from sprint_cert_automation.utils.dates import today_year_month


def default_output_dir(year: int, month: int) -> Path:
    project_root = Path(__file__).resolve().parents[2]
    return project_root / "certificaciones" / f"{year:04d}-{month:02d}"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate sprint certification files for target month",
    )
    parser.add_argument("--forecast", required=True, type=Path, help="Forecast workbook path")
    parser.add_argument("--template", required=True, type=Path, help="Template workbook path")
    parser.add_argument("--year", type=int, default=None, help="Target year")
    parser.add_argument("--month", type=int, default=None, help="Target month 1..12")
    parser.add_argument("--dry-run", action="store_true", help="Run without writing files")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    default_year, default_month = today_year_month()
    year = args.year if args.year is not None else default_year
    month = args.month if args.month is not None else default_month
    output_dir = default_output_dir(year, month)

    result = generate_certificates(
        forecast_path=args.forecast,
        template_path=args.template,
        output_dir=output_dir,
        year=year,
        month=month,
        dry_run=args.dry_run,
    )

    print(f"Generated files: {len(result.generated_files)}")
    for file_path in result.generated_files:
        print(f" - {file_path}")
    print(f"ZIP: {result.zip_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
