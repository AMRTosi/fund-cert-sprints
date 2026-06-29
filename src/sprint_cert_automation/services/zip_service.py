from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


def build_zip(zip_path: Path, files: list[Path]) -> None:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(zip_path, mode="w", compression=ZIP_DEFLATED) as bundle:
        for file_path in files:
            bundle.write(file_path, arcname=file_path.name)
