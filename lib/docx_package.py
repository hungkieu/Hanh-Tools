from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


def unzip_docx(docx_path: str | Path, output_dir: str | Path, *, verbose: bool = False) -> Path:
    source = Path(docx_path).expanduser().resolve()
    destination = Path(output_dir).expanduser().resolve()

    if not source.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {source}")

    if source.suffix.lower() != ".docx":
        raise ValueError("Chỉ hỗ trợ file .docx")

    destination.mkdir(parents=True, exist_ok=True)
    if verbose:
        print(f"Unzipping .docx: {source}", flush=True)
        print(f"Unzip output directory: {destination}", flush=True)

    with ZipFile(source, "r") as docx_zip:
        docx_zip.extractall(destination)

    return destination


def zip_docx(source_dir: str | Path, output_path: str | Path, *, verbose: bool = False) -> Path:
    source = Path(source_dir).expanduser().resolve()
    destination = Path(output_path).expanduser().resolve()

    if not source.exists():
        raise FileNotFoundError(f"Không tìm thấy folder: {source}")

    destination.parent.mkdir(parents=True, exist_ok=True)
    if verbose:
        print(f"Zipping translated .docx from: {source}", flush=True)
        print(f"Translated output path: {destination}", flush=True)

    with ZipFile(destination, "w", ZIP_DEFLATED) as docx_zip:
        for file_path in sorted(path for path in source.rglob("*") if path.is_file()):
            docx_zip.write(file_path, file_path.relative_to(source).as_posix())

    return destination
