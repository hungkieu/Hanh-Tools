from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


class DocConversionError(RuntimeError):
    """Raised when LibreOffice cannot convert a document."""


def convert_doc_to_docx(
    doc_path: str | Path,
    output_dir: str | Path | None = None,
    *,
    backend: str = "auto",
    soffice_path: str | Path | None = None,
    timeout: int = 120,
    verbose: bool = False,
) -> Path:
    """Convert a legacy .doc file to .docx using LibreOffice headless mode.

    LibreOffice is the most reliable option for binary .doc conversion from
    Python because python-docx only reads/writes .docx files.
    """
    source = Path(doc_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {source}")

    if source.suffix.lower() != ".doc":
        raise ValueError("Hàm này chỉ hỗ trợ input .doc")

    destination_dir = (
        Path(output_dir).expanduser().resolve() if output_dir else source.parent
    )
    destination_dir.mkdir(parents=True, exist_ok=True)

    selected_backend = backend.lower()
    if selected_backend not in {"auto", "libreoffice", "textutil"}:
        raise ValueError("backend phải là một trong: auto, libreoffice, textutil")

    converted_path = destination_dir / f"{source.stem}.docx"
    if selected_backend in {"auto", "libreoffice"}:
        soffice = _resolve_soffice(soffice_path, required=selected_backend == "libreoffice")
        if soffice:
            return _convert_with_libreoffice(
                source=source,
                destination_dir=destination_dir,
                converted_path=converted_path,
                soffice=soffice,
                timeout=timeout,
                verbose=verbose,
            )

    if selected_backend in {"auto", "textutil"}:
        textutil = _resolve_textutil(required=selected_backend == "textutil")
        if textutil:
            return _convert_with_textutil(
                source=source,
                converted_path=converted_path,
                textutil=textutil,
                timeout=timeout,
            )

    raise FileNotFoundError(
        "Không tìm thấy công cụ convert. Cài LibreOffice hoặc chạy trên macOS "
        "có sẵn `textutil`."
    )


def _convert_with_libreoffice(
    *,
    source: Path,
    destination_dir: Path,
    converted_path: Path,
    soffice: Path,
    timeout: int,
    verbose: bool,
) -> Path:
    with tempfile.TemporaryDirectory(prefix="hanhTool-lo-") as profile_dir:
        if verbose:
            print(f"Converting .doc with LibreOffice: {source}", flush=True)
            print(f"LibreOffice executable: {soffice}", flush=True)
            print(f"Output directory: {destination_dir}", flush=True)

        command = [
            str(soffice),
            f"-env:UserInstallation=file://{profile_dir}",
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            str(destination_dir),
            str(source),
        ]

        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    if completed.returncode != 0:
        details = (completed.stderr or completed.stdout).strip()
        raise DocConversionError(f"Convert .doc sang .docx thất bại: {details}")

    if not converted_path.exists():
        details = (completed.stderr or completed.stdout).strip()
        raise DocConversionError(
            f"LibreOffice đã chạy nhưng không tạo file: {converted_path}. {details}"
        )

    if verbose:
        print(f"Converted .docx: {converted_path}", flush=True)

    return converted_path


def _convert_with_textutil(
    *,
    source: Path,
    converted_path: Path,
    textutil: Path,
    timeout: int,
) -> Path:
    command = [
        str(textutil),
        "-convert",
        "docx",
        str(source),
        "-output",
        str(converted_path),
    ]

    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=timeout,
        check=False,
    )
    if completed.returncode != 0:
        details = (completed.stderr or completed.stdout).strip()
        raise DocConversionError(f"Convert bằng textutil thất bại: {details}")

    if not converted_path.exists():
        details = (completed.stderr or completed.stdout).strip()
        raise DocConversionError(
            f"textutil đã chạy nhưng không tạo file: {converted_path}. {details}"
        )

    return converted_path


def _resolve_soffice(
    explicit_path: str | Path | None = None,
    *,
    required: bool = True,
) -> Path | None:
    if explicit_path:
        path = Path(explicit_path).expanduser().resolve()
        if path.exists():
            return path
        raise FileNotFoundError(f"Không tìm thấy LibreOffice executable: {path}")

    macos_default = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if macos_default.exists():
        return macos_default

    windows_default = Path("C:/Program Files/LibreOffice/program/soffice.exe")
    if windows_default.exists():
        return windows_default

    for executable in ("soffice", "libreoffice"):
        found = shutil.which(executable)
        if found:
            return Path(found)

    if not required:
        return None

    raise FileNotFoundError(
        "Không tìm thấy LibreOffice. Cài LibreOffice rồi đảm bảo lệnh "
        "`soffice` hoặc `libreoffice` có trong PATH."
    )


def _resolve_textutil(*, required: bool = True) -> Path | None:
    found = shutil.which("textutil")
    if found:
        return Path(found)

    if not required:
        return None

    raise FileNotFoundError("Không tìm thấy `textutil` trên máy này.")
