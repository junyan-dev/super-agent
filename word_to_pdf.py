"""Convert Word documents to PDF files.

A Word document is rendered to PDF by a headless `LibreOffice
<https://www.libreoffice.org/>`_ installation, so the result looks like the
document opened in LibreOffice:

* pages, margins and section breaks are kept;
* fonts, sizes, colours, bold/italic emphasis, headings and list numbering
  are reproduced;
* tables, pictures and other embedded objects are rendered in place;
* headers, footers and page numbering are kept.

Word (``.docx``, macro-enabled ``.docm`` and the legacy ``.doc``) documents
are accepted.  Password-protected documents cannot be converted; they raise
:class:`WordToPdfError`.

Conversion needs LibreOffice, which is not a Python package.  Install it and
either put ``soffice`` on ``PATH`` or point ``WORD_TO_PDF_OFFICE`` at the
executable.  The Python package ``python-docx``, which the tests use to build
sample documents, is installed with ``pip install -r requirements.txt``.

Usage::

    from word_to_pdf import convert_word_to_pdf

    convert_word_to_pdf("report.docx")                # -> report.pdf
    convert_word_to_pdf("report.docx", "report.pdf")  # explicit target
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

__all__ = ["convert_word_to_pdf", "WordToPdfError"]

_OFFICE_HINT = (
    "Word to PDF conversion requires LibreOffice; install it and make sure "
    "`soffice` is on PATH, or set WORD_TO_PDF_OFFICE to its executable."
)

#: Word suffixes the converter accepts.
_WORD_SUFFIXES = {".docx", ".docm", ".doc"}
#: Suffixes that denote an Office Open XML package rather than a legacy file.
_OOXML_SUFFIXES = {".docx", ".docm"}
#: The main part every WordprocessingML document contains.
_OOXML_MAIN_PART = "word/document.xml"
#: Signature of the OLE compound file used by legacy ``.doc`` documents.
_COMPOUND_FILE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
#: Every PDF file starts with this signature.
_PDF_SIGNATURE = b"%PDF"

#: LibreOffice installation paths worth trying when it is not on ``PATH``.
_COMMON_OFFICE_PATHS = (
    Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    Path("/usr/lib/libreoffice/program/soffice"),
    Path("/usr/local/lib/libreoffice/program/soffice"),
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
)

#: Seconds a single conversion may take before it is abandoned.
_CONVERSION_TIMEOUT = 180.0


class WordToPdfError(Exception):
    """Raised when a Word document cannot be converted to a PDF file."""


def convert_word_to_pdf(
    source: str | Path,
    target: str | Path | None = None,
) -> Path:
    """Convert the Word document *source* into a PDF file.

    The document is rendered by LibreOffice, so its layout is preserved.

    Args:
        source: Path to an existing Word document (``.docx``, ``.docm`` or
            ``.doc``).
        target: Path of the ``.pdf`` file to write.  Defaults to *source*
            with its suffix replaced by ``.pdf``.  Parent directories are
            created when missing.

    Returns:
        The path of the written ``.pdf`` file.

    Raises:
        FileNotFoundError: If *source* does not exist.
        IsADirectoryError: If *source* is not a regular file.
        ValueError: If *source* is not a Word document or *target* does not
            use the ``.pdf`` suffix.
        WordToPdfError: If *source* is not a readable Word document, if
            LibreOffice is not installed, or if it fails to produce a PDF.
    """
    source_path = Path(source)
    target_path = (
        Path(target) if target is not None else source_path.with_suffix(".pdf")
    )
    _validate_source(source_path)
    _validate_target(target_path)
    _ensure_word_document(source_path)
    office = _find_office()

    # LibreOffice derives the output name from the source file, so convert
    # into a private directory and move the result to the requested target.
    with tempfile.TemporaryDirectory(prefix="word-to-pdf-") as workspace:
        output_directory = Path(workspace) / "output"
        output_directory.mkdir()
        profile_directory = Path(workspace) / "profile"
        profile_directory.mkdir()
        _run_office(office, source_path, output_directory, profile_directory)
        pdf_path = _produced_pdf(output_directory, source_path)
        _ensure_pdf(pdf_path, source_path)

        target_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            shutil.move(str(pdf_path), str(target_path))
        except OSError as exc:
            raise WordToPdfError(
                f"could not write PDF file {target_path}: {exc}"
            ) from exc
    return target_path


def _find_office() -> str:
    """Return the LibreOffice executable used for conversion."""
    configured = os.environ.get("WORD_TO_PDF_OFFICE", "").strip()
    if configured:
        executable = Path(configured).expanduser()
        if not executable.is_file() or not os.access(executable, os.X_OK):
            raise WordToPdfError(
                "WORD_TO_PDF_OFFICE does not point to an executable: "
                f"{configured}"
            )
        return str(executable)

    for name in ("soffice", "libreoffice", "soffice.exe"):
        found = shutil.which(name)
        if found:
            return found

    for candidate in _COMMON_OFFICE_PATHS:
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return str(candidate)

    raise WordToPdfError(_OFFICE_HINT)


def _validate_source(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Word file not found: {path}")
    if not path.is_file():
        raise IsADirectoryError(f"not a regular file: {path}")
    if path.suffix.lower() not in _WORD_SUFFIXES:
        raise ValueError(
            "source must be a Word document (.docx, .docm or .doc), "
            f"not {path.suffix or '<none>'!r}"
        )


def _validate_target(path: Path) -> None:
    if path.suffix.lower() != ".pdf":
        raise ValueError(
            f"target must use the .pdf suffix, not {path.suffix or '<none>'!r}"
        )
    if path.is_dir():
        raise IsADirectoryError(f"target is a directory: {path}")


def _ensure_word_document(path: Path) -> None:
    """Reject files that only pretend to be Word documents."""
    with path.open("rb") as handle:
        signature = handle.read(8)

    if path.suffix.lower() in _OOXML_SUFFIXES:
        if signature.startswith(_COMPOUND_FILE_SIGNATURE):
            raise WordToPdfError(
                f"Word file {path} is encrypted or not a valid "
                f"{path.suffix.lower()} document"
            )
        try:
            with zipfile.ZipFile(path) as package:
                parts = set(package.namelist())
        except (zipfile.BadZipFile, OSError) as exc:
            raise WordToPdfError(
                f"Word file {path} is not a valid document: {exc}"
            ) from exc
        if _OOXML_MAIN_PART not in parts:
            raise WordToPdfError(f"Word file {path} is not a Word document")
        return

    if signature != _COMPOUND_FILE_SIGNATURE:
        raise WordToPdfError(
            f"Word file {path} is not a valid legacy .doc document"
        )


def _run_office(
    office: str,
    source: Path,
    output_directory: Path,
    profile_directory: Path,
) -> None:
    """Convert *source* into *output_directory* with a private profile."""
    command = [
        office,
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nodefault",
        f"-env:UserInstallation={profile_directory.as_uri()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_directory),
        str(source),
    ]
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            errors="replace",
            timeout=_CONVERSION_TIMEOUT,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise WordToPdfError(
            f"converting {source} to PDF timed out after "
            f"{_CONVERSION_TIMEOUT:g} seconds"
        ) from exc
    except OSError as exc:
        raise WordToPdfError(
            f"could not run LibreOffice ({office}): {exc}"
        ) from exc

    if completed.returncode != 0:
        detail = (completed.stderr or "").strip() or (completed.stdout or "").strip()
        raise WordToPdfError(
            f"LibreOffice could not convert {source}: "
            f"{detail or f'exit status {completed.returncode}'}"
        )


def _produced_pdf(output_directory: Path, source: Path) -> Path:
    """Return the PDF LibreOffice wrote for *source*."""
    expected = output_directory / f"{source.stem}.pdf"
    if expected.is_file():
        return expected

    produced = sorted(output_directory.glob("*.pdf"))
    if produced:
        return produced[0]

    raise WordToPdfError(f"LibreOffice did not produce a PDF for {source}")


def _ensure_pdf(path: Path, source: Path) -> None:
    try:
        with path.open("rb") as handle:
            signature = handle.read(len(_PDF_SIGNATURE))
    except OSError as exc:
        raise WordToPdfError(
            f"could not read the PDF written for {source}: {exc}"
        ) from exc
    if signature != _PDF_SIGNATURE:
        raise WordToPdfError(
            f"LibreOffice wrote an invalid PDF file for {source}"
        )
