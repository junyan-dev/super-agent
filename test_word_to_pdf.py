"""Tests for :mod:`word_to_pdf`."""

from __future__ import annotations

import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import word_to_pdf
from word_to_pdf import WordToPdfError, convert_word_to_pdf

try:
    import fitz
except ImportError:  # pragma: no cover - exercised only without dependencies
    fitz = None

try:
    from docx import Document
except ImportError:  # pragma: no cover - exercised only without dependencies
    Document = None


def _office_available() -> bool:
    try:
        word_to_pdf._find_office()
    except Exception:  # LibreOffice is simply unavailable
        return False
    return True


requires_office = unittest.skipUnless(
    fitz is not None and Document is not None and _office_available(),
    "python-docx, PyMuPDF and LibreOffice are required",
)


@requires_office
class ConvertWordToPdfTests(unittest.TestCase):
    def setUp(self) -> None:
        self._temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self._temporary_directory.cleanup)
        self.directory = Path(self._temporary_directory.name)

    def _path(self, name: str) -> Path:
        path = self.directory / name
        path.parent.mkdir(parents=True, exist_ok=True)
        return path

    def _document(self, name: str, build) -> Path:
        path = self._path(name)
        document = Document()
        build(document)
        document.save(str(path))
        return path

    def _text_document(self, name: str = "input.docx") -> Path:
        def build(document) -> None:
            document.add_heading("Quarterly Report", level=1)
            document.add_paragraph("Revenue grew in every region.")
            document.add_paragraph("Costs stayed flat.")

        return self._document(name, build)

    def _pdf_pages(self, path: Path) -> list[str]:
        pdf = fitz.open(str(path))
        try:
            return [page.get_text() for page in pdf]
        finally:
            pdf.close()

    def _fake_office(self, script: str) -> Path:
        path = self._path("fake-soffice")
        path.write_text(f"#!/bin/sh\n{script}\n", encoding="utf-8")
        path.chmod(0o755)
        return path

    def test_converts_document_to_pdf(self) -> None:
        source = self._text_document()
        original = source.read_bytes()

        output = convert_word_to_pdf(source)

        self.assertEqual(output, source.with_suffix(".pdf"))
        self.assertTrue(output.is_file())
        self.assertEqual(output.read_bytes()[:4], b"%PDF")
        self.assertEqual(source.read_bytes(), original)

        text = "\n".join(self._pdf_pages(output))
        self.assertIn("Quarterly Report", text)
        self.assertIn("Revenue grew in every region.", text)
        self.assertIn("Costs stayed flat.", text)

    def test_writes_to_explicit_target(self) -> None:
        source = self._text_document("meeting notes.docx")
        target = self._path("converted/report.pdf")

        output = convert_word_to_pdf(source, target)

        self.assertEqual(output, target)
        self.assertTrue(target.is_file())
        self.assertIn(
            "Quarterly Report", "".join(self._pdf_pages(target))
        )

    def test_preserves_tables_and_images(self) -> None:
        pixmap = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 40, 30), 0)
        pixmap.set_rect(pixmap.irect, (255, 0, 0))
        image_path = self._path("figure.png")
        pixmap.save(str(image_path))

        def build(document) -> None:
            table = document.add_table(rows=2, cols=2)
            table.style = "Table Grid"
            table.cell(0, 0).text = "Alpha"
            table.cell(0, 1).text = "Beta"
            table.cell(1, 0).text = "Gamma"
            table.cell(1, 1).text = "Delta"
            document.add_paragraph("Figure below")
            document.add_picture(str(image_path))

        source = self._document("table.docx", build)
        output = convert_word_to_pdf(source)

        pdf = fitz.open(str(output))
        try:
            text = "\n".join(page.get_text() for page in pdf)
            images = [
                image
                for page in pdf
                for image in page.get_images(full=True)
            ]
        finally:
            pdf.close()

        for value in ("Alpha", "Beta", "Gamma", "Delta"):
            self.assertIn(value, text)
        self.assertGreaterEqual(len(images), 1)

    def test_starts_each_page_break_on_a_new_pdf_page(self) -> None:
        def build(document) -> None:
            document.add_paragraph("First page body")
            document.add_page_break()
            document.add_paragraph("Second page body")

        source = self._document("pages.docx", build)
        pages = self._pdf_pages(convert_word_to_pdf(source))

        self.assertEqual(len(pages), 2)
        self.assertIn("First page body", pages[0])
        self.assertIn("Second page body", pages[1])

    def test_converts_blank_document(self) -> None:
        source = self._document("blank.docx", lambda document: None)

        output = convert_word_to_pdf(source)

        self.assertTrue(output.is_file())
        self.assertEqual(len(self._pdf_pages(output)), 1)

    def test_rejects_non_pdf_target(self) -> None:
        source = self._text_document()

        with self.assertRaises(ValueError):
            convert_word_to_pdf(source, self.directory / "output.docx")

    def test_rejects_directory_target(self) -> None:
        source = self._text_document()
        target = self._path("output.pdf")
        target.mkdir()

        with self.assertRaises(IsADirectoryError):
            convert_word_to_pdf(source, target)

    def test_rejects_missing_source(self) -> None:
        with self.assertRaises(FileNotFoundError):
            convert_word_to_pdf(self.directory / "missing.docx")

    def test_rejects_directory_source(self) -> None:
        with self.assertRaises(IsADirectoryError):
            convert_word_to_pdf(self.directory)

    def test_rejects_unsupported_source(self) -> None:
        source = self._path("notes.txt")
        source.write_text("plain text", encoding="utf-8")

        with self.assertRaises(ValueError):
            convert_word_to_pdf(source)

    def test_rejects_file_that_is_not_a_document(self) -> None:
        source = self._path("broken.docx")
        source.write_bytes(b"this is not a Word document")

        with self.assertRaises(WordToPdfError):
            convert_word_to_pdf(source)

    def test_rejects_package_without_word_content(self) -> None:
        source = self._path("spreadsheet.docx")
        with zipfile.ZipFile(source, "w") as package:
            package.writestr("xl/workbook.xml", "<workbook/>")

        with self.assertRaises(WordToPdfError):
            convert_word_to_pdf(source)

    def test_reports_missing_office_from_environment(self) -> None:
        source = self._text_document()
        missing = str(self.directory / "not-installed" / "soffice")

        with mock.patch.dict(os.environ, {"WORD_TO_PDF_OFFICE": missing}):
            with self.assertRaisesRegex(WordToPdfError, "WORD_TO_PDF_OFFICE"):
                convert_word_to_pdf(source)

    @unittest.skipUnless(os.name == "posix", "needs a POSIX shell script")
    def test_reports_office_failure(self) -> None:
        source = self._text_document()
        office = self._fake_office('echo "cannot render" >&2\nexit 3')

        with mock.patch.dict(os.environ, {"WORD_TO_PDF_OFFICE": str(office)}):
            with self.assertRaisesRegex(WordToPdfError, "cannot render"):
                convert_word_to_pdf(source)

    @unittest.skipUnless(os.name == "posix", "needs a POSIX shell script")
    def test_reports_conversion_timeout(self) -> None:
        source = self._text_document()
        office = self._fake_office("exec sleep 30")

        with mock.patch.dict(os.environ, {"WORD_TO_PDF_OFFICE": str(office)}):
            with mock.patch.object(word_to_pdf, "_CONVERSION_TIMEOUT", 0.2):
                with self.assertRaisesRegex(WordToPdfError, "timed out"):
                    convert_word_to_pdf(source)


if __name__ == "__main__":
    unittest.main()
