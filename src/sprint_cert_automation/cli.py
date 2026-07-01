from __future__ import annotations

import argparse
from pathlib import Path
import sys

from sprint_cert_automation.app import export_certificates_to_pdf, generate_certificates
from sprint_cert_automation.utils.dates import today_year_month


def default_output_dir(year: int, month: int) -> Path:
    project_root = Path(__file__).resolve().parents[2]
    return project_root / "certificaciones" / f"{year:04d}-{month:02d}"


def _add_generate_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--forecast", required=True, type=Path, help="Forecast workbook path")
    parser.add_argument("--template", required=True, type=Path, help="Template workbook path")
    parser.add_argument("--year", type=int, default=None, help="Target year")
    parser.add_argument("--month", type=int, default=None, help="Target month 1..12")
    parser.add_argument("--dry-run", action="store_true", help="Run without writing files")


def _add_export_pdf_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--input-dir", type=Path, default=None, help="Folder with validated .xlsm files")
    parser.add_argument("--year", type=int, default=None, help="Target year for default folder")
    parser.add_argument("--month", type=int, default=None, help="Target month 1..12 for default folder")
    parser.add_argument("--dry-run", action="store_true", help="List files without executing macro/export")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Sprint certification automation commands",
    )
    subparsers = parser.add_subparsers(dest="command")

    generate_parser = subparsers.add_parser(
        "generate",
        help="Generate monthly .xlsm certification files",
    )
    _add_generate_arguments(generate_parser)

    export_pdf_parser = subparsers.add_parser(
        "export-pdf",
        help="Execute macro and export validated .xlsm files to .pdf",
    )
    _add_export_pdf_arguments(export_pdf_parser)

    return parser


def _normalize_args(raw_args: list[str]) -> list[str]:
    if not raw_args:
        return ["generate"]
    if raw_args[0] in {"generate", "export-pdf"}:
        return raw_args
    return ["generate", *raw_args]


def _run_generate(args: argparse.Namespace) -> int:
    parser = build_parser()

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

    return 0


def _run_export_pdf(args: argparse.Namespace) -> int:
    default_year, default_month = today_year_month()
    year = args.year if args.year is not None else default_year
    month = args.month if args.month is not None else default_month

    input_dir = args.input_dir if args.input_dir is not None else default_output_dir(year, month)
    result = export_certificates_to_pdf(
        input_dir=input_dir,
        dry_run=args.dry_run,
    )

    print(f"PDF files {'planned' if args.dry_run else 'exported'}: {len(result.exported_files)}")
    for file_path in result.exported_files:
        print(f" - {file_path}")

    if result.errors:
        print(f"Files with errors: {len(result.errors)}")
        for file_path, error in result.errors.items():
            print(f" - {file_path}: {error}")
        return 1

    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    raw_args = sys.argv[1:] if argv is None else argv
    args = parser.parse_args(_normalize_args(raw_args))

    if args.command == "export-pdf":
        return _run_export_pdf(args)
    if args.command == "generate":
        return _run_generate(args)

    parser.error("Command must be one of: generate, export-pdf")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
