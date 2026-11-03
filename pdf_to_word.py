"""Convert PDF documents to Word (``.docx``) files.

The converter walks a PDF page by page and recreates its readable content in
a Word document:

* text becomes editable paragraphs, preserving font family, size, colour,
  bold/italic emphasis and superscripts, with clearly larger blocks promoted
  to Word heading styles;
* images are embedded in place;
* bordered tables are rebuilt as editable Word tables;
* each PDF page starts on a new Word page.

Scanned PDFs are carried over as page images because no OCR is performed.
Vector graphics and line art outside detected tables are not reproduced, and
page headers and footers are converted like ordinary page content.

Conversion needs `PyMuPDF <https://pymupdf.readthedocs.io/>`_ (``pymupdf``)
to read PDFs and `python-docx <https://python-docx.readthedocs.io/>`_ to write
Word files.  Install both with ``pip install -r requirements.txt``.

Usage::

    from pdf_to_word import convert_pdf_to_word

    convert_pdf_to_word("report.pdf")                 # -> report.docx
    convert_pdf_to_word("report.pdf", "report.docx")  # explicit target
"""

from __future__ import annotations

import io
import re
from collections import Counter
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Sequence

__all__ = ["convert_pdf_to_word", "PdfToWordError"]

_DEPENDENCY_HINT = (
    "PDF to Word conversion requires PyMuPDF and python-docx; "
    "install them with `pip install -r requirements.txt`."
)

#: A font name such as ``ABCDEF+Helvetica`` has a six-letter subset prefix.
_SUBSET_FONT_PREFIX = re.compile(r"^[A-Z]{6}\+")
#: Style suffixes that shrink ``Helvetica-Bold`` back to ``Helvetica``.
_FONT_STYLE_SUFFIX = re.compile(
    r"(?:-(?:Bold)?(?:Italic|Oblique)?(?:MT)?)+$", re.IGNORECASE
)
#: PDF font names that differ from their usual Word counterpart.
_FONT_ALIASES = {
    "ArialMT": "Arial",
    "TimesNewRomanPS": "Times New Roman",
    "TimesNewRomanPSMT": "Times New Roman",
    "TimesNewRoman": "Times New Roman",
    "CourierNewPS": "Courier New",
    "CourierNewPSMT": "Courier New",
    "CourierNew": "Courier New",
}

#: PyMuPDF span flag bits (<https://pymupdf.readthedocs.io/en/latest/textpage.html>).
_FLAG_SUPERSCRIPT = 1
_FLAG_ITALIC = 2
_FLAG_BOLD = 16

#: Blocks longer than this are never treated as headings.
_MAX_HEADING_LENGTH = 120


class PdfToWordError(Exception):
    """Raised when a PDF file cannot be converted to a Word document."""


def convert_pdf_to_word(
    source: str | Path,
    target: str | Path | None = None,
) -> Path:
    """Convert the PDF file *source* into a Word document.

    Text, images and bordered tables are recreated page by page.

    Args:
        source: Path to an existing PDF file.
        target: Path of the ``.docx`` file to write.  Defaults to *source*
            with its suffix replaced by ``.docx``.  Parent directories are
            created when missing.

    Returns:
        The path of the written ``.docx`` file.

    Raises:
        FileNotFoundError: If *source* does not exist.
        IsADirectoryError: If *source* is not a regular file.
        ValueError: If *target* does not use the ``.docx`` suffix.
        PdfToWordError: If *source* is not a readable, unencrypted PDF with
            at least one page.
        ImportError: If PyMuPDF or python-docx is not installed.
    """
    dependencies = _load_dependencies()
    source_path = Path(source)
    target_path = (
        Path(target) if target is not None else source_path.with_suffix(".docx")
    )
    _validate_source(source_path)
    _validate_target(target_path)

    try:
        pdf = dependencies.fitz.open(str(source_path))
    except RuntimeError as exc:
        raise PdfToWordError(
            f"could not read PDF file {source_path}: {exc}"
        ) from exc

    try:
        if pdf.needs_pass:
            raise PdfToWordError(
                f"PDF file {source_path} is password protected"
            )
        if pdf.page_count == 0:
            raise PdfToWordError(f"PDF file {source_path} has no pages")
        document = _build_document(dependencies, pdf)
    finally:
        pdf.close()

    target_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(target_path))
    return target_path


def _load_dependencies() -> SimpleNamespace:
    """Import optional third-party packages or explain how to install them."""
    try:
        import fitz
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.shared import Pt, RGBColor
    except ImportError as exc:
        raise ImportError(_DEPENDENCY_HINT) from exc

    return SimpleNamespace(
        fitz=fitz,
        Document=Document,
        WD_ALIGN_PARAGRAPH=WD_ALIGN_PARAGRAPH,
        qn=qn,
        Pt=Pt,
        RGBColor=RGBColor,
    )


def _validate_source(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"PDF file not found: {path}")
    if not path.is_file():
        raise IsADirectoryError(f"not a regular file: {path}")


def _validate_target(path: Path) -> None:
    if path.suffix.lower() != ".docx":
        raise ValueError(
            f"target must use the .docx suffix, not {path.suffix or '<none>'!r}"
        )


def _build_document(dependencies: SimpleNamespace, pdf: Any) -> Any:
    body_size, body_font, heading_levels = _document_text_styles(pdf)
    document = dependencies.Document()
    _configure_page(document, dependencies, pdf[0])
    _configure_normal_style(document, dependencies, body_size, body_font)

    for page_index, page in enumerate(pdf):
        if page_index:
            document.add_page_break()
        _add_page(dependencies, document, page, body_size, heading_levels)
    return document


def _document_text_styles(
    pdf: Any,
) -> tuple[float, str, dict[float, int]]:
    """Return body size, body family and a size-to-heading-level mapping.

    The body style is the most common size and family, weighted by text
    length.  Distinct sizes that are clearly larger become heading levels in
    descending order, so the largest text maps to ``Heading 1``.
    """
    sizes: Counter[float] = Counter()
    fonts: Counter[str] = Counter()
    for page in pdf:
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", ()):
                for span in line.get("spans", ()):
                    text = span.get("text", "").strip()
                    if not text:
                        continue
                    sizes[round(float(span.get("size", 0.0)) * 2) / 2] += len(
                        text
                    )
                    fonts[span.get("font", "")] += len(text)

    body_size = sizes.most_common(1)[0][0] if sizes else 11.0
    body_font = (
        _normalise_font_name(fonts.most_common(1)[0][0])
        if fonts
        else "Calibri"
    )
    heading_sizes = sorted(
        (size for size in sizes if size >= body_size * 1.15 and size > body_size),
        reverse=True,
    )
    heading_levels = {
        size: min(rank + 1, 3) for rank, size in enumerate(heading_sizes)
    }
    return body_size, body_font, heading_levels


def _configure_page(
    document: Any, dependencies: SimpleNamespace, first_page: Any
) -> None:
    section = document.sections[0]
    if first_page.rect.width > 0 and first_page.rect.height > 0:
        section.page_width = dependencies.Pt(first_page.rect.width)
        section.page_height = dependencies.Pt(first_page.rect.height)
    # Keep a one-inch margin, but never take more than a quarter of a page.
    margin_points = min(
        72.0,
        first_page.rect.width / 4,
        first_page.rect.height / 4,
    )
    margin = dependencies.Pt(margin_points)
    section.left_margin = margin
    section.right_margin = margin
    section.top_margin = margin
    section.bottom_margin = margin


def _configure_normal_style(
    document: Any,
    dependencies: SimpleNamespace,
    body_size: float,
    body_font: str,
) -> None:
    normal = document.styles["Normal"]
    normal.font.size = dependencies.Pt(body_size)
    if body_font:
        normal.font.name = body_font
        _set_east_asian_font(dependencies, normal.element, body_font)


def _add_page(
    dependencies: SimpleNamespace,
    document: Any,
    page: Any,
    body_size: float,
    heading_levels: dict[float, int],
) -> None:
    section = document.sections[0]
    usable_width = (
        section.page_width - section.left_margin - section.right_margin
    )
    page_width = page.rect.width or 1.0
    tables = _find_tables(dependencies.fitz, page)
    emitted: set[int] = set()

    for block in page.get_text("dict")["blocks"]:
        table_index = _containing_table(
            dependencies.fitz.Rect(block["bbox"]), tables
        )
        if table_index is not None:
            if table_index not in emitted:
                _add_table(document, tables[table_index][1])
                emitted.add(table_index)
            continue
        if block.get("type") == 1:
            _add_image_block(
                dependencies,
                document,
                block,
                page_width,
                usable_width,
            )
        else:
            _add_text_block(
                dependencies, document, block, body_size, heading_levels
            )

    for table_index, (_, rows) in enumerate(tables):
        if table_index not in emitted:
            _add_table(document, rows)


def _find_tables(fitz: Any, page: Any) -> list[tuple[Any, Sequence[Sequence[Any]]]]:
    """Return ``(bbox, rows)`` for the non-empty tables found on *page*."""
    finder = getattr(page, "find_tables", None)
    if finder is None:
        return []

    try:
        found = finder()
    except Exception:  # table detection is best effort
        return []

    tables = []
    for table in getattr(found, "tables", ()):
        rows = table.extract()
        if not rows or not any(
            any(cell not in (None, "") for cell in row) for row in rows
        ):
            continue
        tables.append((fitz.Rect(table.bbox), rows))
    return tables


def _containing_table(
    rect: Any,
    tables: Sequence[tuple[Any, Sequence[Sequence[Any]]]],
) -> int | None:
    """Return the index of the first table that fully contains *rect*."""
    for index, (bbox, _) in enumerate(tables):
        if bbox.contains(rect):
            return index
    return None


def _add_table(document: Any, rows: Sequence[Sequence[Any]]) -> None:
    columns = max((len(row) for row in rows), default=0)
    if columns == 0:
        return

    table = document.add_table(rows=len(rows), cols=columns)
    table.style = "Table Grid"
    for row_index, row in enumerate(rows):
        for column_index in range(columns):
            value = row[column_index] if column_index < len(row) else ""
            table.cell(row_index, column_index).text = (
                "" if value is None else str(value).strip()
            )


def _add_image_block(
    dependencies: SimpleNamespace,
    document: Any,
    block: Any,
    page_width: float,
    usable_width: Any,
) -> None:
    data = block.get("image")
    if not data:
        return

    x0, _, x1, _ = block["bbox"]
    width = min(int(usable_width * ((x1 - x0) / page_width)), int(usable_width))
    if width <= 0:
        return

    try:
        document.add_picture(io.BytesIO(data), width=width)
    except Exception:  # an image Word cannot decode must not stop conversion
        return
    document.paragraphs[-1].alignment = dependencies.WD_ALIGN_PARAGRAPH.CENTER


def _add_text_block(
    dependencies: SimpleNamespace,
    document: Any,
    block: Any,
    body_size: float,
    heading_levels: dict[float, int],
) -> None:
    fragments = _block_fragments(block)
    text = "".join(fragment_text for fragment_text, _ in fragments).strip()
    if not text:
        return

    largest_size = max(
        (
            float(span.get("size", 0.0))
            for fragment_text, span in fragments
            if fragment_text.strip()
        ),
        default=body_size,
    )
    heading_level = 0
    if len(text) <= _MAX_HEADING_LENGTH:
        heading_level = _heading_level(largest_size, heading_levels)

    paragraph = document.add_paragraph(
        style=f"Heading {heading_level}" if heading_level else None
    )
    if heading_level:
        paragraph.paragraph_format.space_before = dependencies.Pt(12)
    paragraph.paragraph_format.space_after = dependencies.Pt(6)
    _add_runs(dependencies, paragraph, fragments)


def _block_fragments(block: Any) -> list[tuple[str, Any]]:
    """Flatten a text block into ``(text, span)`` pieces with joined lines."""
    fragments: list[tuple[str, Any]] = []
    for line in block.get("lines", ()):
        spans = [
            (span.get("text", ""), span) for span in line.get("spans", ())
        ]
        spans = [(text, span) for text, span in spans if text]
        if not spans:
            continue

        if fragments:
            previous_text = fragments[-1][0]
            first_text = spans[0][0]
            if previous_text and first_text:
                if previous_text.endswith("-") and first_text[:1].islower():
                    fragments[-1] = (previous_text[:-1], fragments[-1][1])
                elif _needs_separator(previous_text[-1], first_text[0]):
                    fragments.append((" ", fragments[-1][1]))
        fragments.extend(spans)
    return fragments


def _needs_separator(left: str, right: str) -> bool:
    if left.isspace() or right.isspace():
        return False
    return not (_is_cjk(left) and _is_cjk(right))


def _is_cjk(character: str) -> bool:
    codepoint = ord(character)
    return (
        0x2E80 <= codepoint <= 0x9FFF
        or 0xAC00 <= codepoint <= 0xD7AF
        or 0xF900 <= codepoint <= 0xFAFF
        or 0xFF00 <= codepoint <= 0xFFEF
    )


def _heading_level(
    largest_size: float, heading_levels: dict[float, int]
) -> int:
    size = round(largest_size * 2) / 2
    return heading_levels.get(size, 0)


def _add_runs(
    dependencies: SimpleNamespace,
    paragraph: Any,
    fragments: Sequence[tuple[str, Any]],
) -> None:
    current_run = None
    current_key: tuple[Any, ...] | None = None
    for text, span in fragments:
        if not text:
            continue
        key = (
            span.get("font"),
            span.get("size"),
            span.get("flags"),
            span.get("color"),
        )
        if current_run is None or key != current_key:
            current_run = paragraph.add_run()
            _format_run(dependencies, current_run, span)
            current_key = key
        current_run.text = (current_run.text or "") + text


def _format_run(
    dependencies: SimpleNamespace, run: Any, span: Any
) -> None:
    flags = int(span.get("flags", 0))
    run.font.bold = bool(flags & _FLAG_BOLD)
    run.font.italic = bool(flags & _FLAG_ITALIC)
    if flags & _FLAG_SUPERSCRIPT:
        run.font.superscript = True

    size = span.get("size")
    if size:
        run.font.size = dependencies.Pt(float(size))

    font_name = _normalise_font_name(span.get("font", ""))
    if font_name:
        run.font.name = font_name
        _set_east_asian_font(dependencies, run._element, font_name)

    color = int(span.get("color", 0))
    run.font.color.rgb = dependencies.RGBColor(
        (color >> 16) & 0xFF,
        (color >> 8) & 0xFF,
        color & 0xFF,
    )


def _set_east_asian_font(
    dependencies: SimpleNamespace, element: Any, font_name: str
) -> None:
    """Apply *font_name* to East Asian text as well as Latin text."""
    rpr = element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(dependencies.qn("w:eastAsia"), font_name)


def _normalise_font_name(name: str) -> str:
    name = _SUBSET_FONT_PREFIX.sub("", name or "").strip()
    if not name:
        return ""
    name = _FONT_STYLE_SUFFIX.sub("", name).strip()
    return _FONT_ALIASES.get(name, name)
