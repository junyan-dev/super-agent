"""Tests for :mod:`pdf_to_word`."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pdf_to_word import PdfToWordError, convert_pdf_to_word

try:
    import fitz
except ImportError:  # pragma: no cover - exercised only without dependencies
    fitz = None

try:
    from docx import Document
except ImportError:  # pragma: no cover - exercised only without dependencies
    Document = None

requires_dependencies = unittest.skipUnless(
    fitz is not None and Document is not None,
    "PyMuPDF and python-docx are required",
)


@requires_dependencies
class ConvertPdfToWordTests(unittest.TestCase):
    def setUp(self) -> None:
        self._temporary_directory = tempfile.TemporaryDirectory()
        self.addCleanup(self._temporary_directory.cleanup)
        self.directory = Path(self._temporary_directory.name)

    def _path(self, name: str) -> Path:
        path = self.directory / name
        path.parent.mkdir(parents=True, exist_ok=True)
        return path

    def _save_pdf(self, path: Path, build) -> None:
        pdf = fitz.open()
        try:
            build(pdf)
            pdf.save(str(path))
        finally:
            pdf.close()

    def _text_pdf(self, name: str = "input.pdf") -> Path:
        path = self._path(name)

        def build(pdf) -> None:
            page = pdf.new_page(width=595, height=842)
            page.insert_text(
                (72, 80), "Quarterly Report", fontsize=24, fontname="hebo"
            )
            page.insert_text(
                (72, 120),
                "Revenue grew in every region.\nCosts stayed flat.",
                fontsize=11,
            )

        self._save_pdf(path, build)
        return path

    def test_converts_text_and_headings(self) -> None:
        source = self._text_pdf()

        output = convert_pdf_to_word(source)

        self.assertEqual(output, source.with_suffix(".docx"))
        self.assertTrue(output.is_file())

        document = Document(str(output))
        paragraphs = [
            paragraph
            for paragraph in document.paragraphs
            if paragraph.text.strip()
        ]
        self.assertEqual(paragraphs[0].text, "Quarterly Report")
        self.assertEqual(paragraphs[0].style.name, "Heading 1")
        self.assertEqual(
            paragraphs[1].text,
            "Revenue grew in every region. Costs stayed flat.",
        )
        self.assertEqual(paragraphs[1].style.name, "Normal")
        self.assertAlmostEqual(
            paragraphs[1].runs[0].font.size.pt, 11.0, places=1
        )

    def test_maps_larger_text_to_heading_levels(self) -> None:
        source = self._path("outline.pdf")

        def build(pdf) -> None:
            page = pdf.new_page(width=595, height=842)
            page.insert_text(
                (72, 80), "Document title", fontsize=24, fontname="hebo"
            )
            page.insert_text(
                (72, 120), "Section heading", fontsize=16, fontname="hebo"
            )
            page.insert_text(
                (72, 160), "Body paragraph text.", fontsize=11
            )

        self._save_pdf(source, build)

        document = Document(str(convert_pdf_to_word(source)))
        paragraphs = [
            paragraph
            for paragraph in document.paragraphs
            if paragraph.text.strip()
        ]

        self.assertEqual(
            [paragraph.style.name for paragraph in paragraphs],
            ["Heading 1", "Heading 2", "Normal"],
        )

    def test_preserves_bold_and_italic_runs(self) -> None:
        source = self._path("emphasis.pdf")

        def build(pdf) -> None:
            page = pdf.new_page(width=595, height=842)
            page.insert_text(
                (72, 100), "Bold text", fontsize=11, fontname="hebo"
            )
            page.insert_text(
                (72, 130), "Italic text", fontsize=11, fontname="heit"
            )
            page.insert_text(
                (72, 160), "Plain text", fontsize=11, fontname="helv"
            )

        self._save_pdf(source, build)

        document = Document(str(convert_pdf_to_word(source)))

        runs = [
            run
            for paragraph in document.paragraphs
            for run in paragraph.runs
            if run.text.strip()
        ]
        bold_runs = [run for run in runs if run.font.bold]
        italic_runs = [run for run in runs if run.font.italic]
        self.assertEqual([run.text for run in bold_runs], ["Bold text"])
        self.assertEqual([run.text for run in italic_runs], ["Italic text"])

    def test_embeds_images(self) -> None:
        source = self._path("image.pdf")

        def build(pdf) -> None:
            page = pdf.new_page(width=595, height=842)
            page.insert_text((72, 80), "Figure below", fontsize=11)
            pixmap = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 40, 30), 0)
            pixmap.set_rect(pixmap.irect, (255, 0, 0))
            page.insert_image(fitz.Rect(72, 120, 152, 180), pixmap=pixmap)

        self._save_pdf(source, build)

        document = Document(str(convert_pdf_to_word(source)))

        self.assertEqual(len(document.inline_shapes), 1)
        self.assertGreater(document.inline_shapes[0].width, 0)

    def test_handles_image_only_pages(self) -> None:
        source = self._path("scan.pdf")

        def build(pdf) -> None:
            page = pdf.new_page(width=595, height=842)
            pixmap = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 40, 30), 0)
            pixmap.set_rect(pixmap.irect, (0, 0, 255))
            page.insert_image(fitz.Rect(72, 72, 523, 770), pixmap=pixmap)

        self._save_pdf(source, build)

        document = Document(str(convert_pdf_to_word(source)))

        self.assertEqual(len(document.inline_shapes), 1)

    def test_rebuilds_bordered_tables(self) -> None:
        source = self._path("table.pdf")

        def build(pdf) -> None:
            page = pdf.new_page(width=595, height=842)
            for y in (300, 330, 360):
                page.draw_line((72, y), (400, y))
            for x in (72, 250, 400):
                page.draw_line((x, 300), (x, 360))
            page.insert_text((80, 320), "A", fontsize=11)
            page.insert_text((260, 320), "B", fontsize=11)
            page.insert_text((80, 350), "1", fontsize=11)
            page.insert_text((260, 350), "2", fontsize=11)

        self._save_pdf(source, build)

        document = Document(str(convert_pdf_to_word(source)))

        self.assertEqual(len(document.tables), 1)
        table = document.tables[0]
        self.assertEqual(
            [[cell.text for cell in row.cells] for row in table.rows],
            [["A", "B"], ["1", "2"]],
        )
        self.assertEqual(
            [paragraph.text for paragraph in document.paragraphs if paragraph.text],
            [],
        )

    def test_starts_each_page_on_a_new_word_page(self) -> None:
        source = self._path("pages.pdf")

        def build(pdf) -> None:
            for page_number in (1, 2):
                page = pdf.new_page(width=595, height=842)
                page.insert_text(
                    (72, 80), f"Page {page_number} body", fontsize=11
                )

        self._save_pdf(source, build)

        document = Document(str(convert_pdf_to_word(source)))
        from docx.oxml.ns import qn

        page_breaks = [
            element
            for element in document.element.body.findall(f".//{qn('w:br')}")
            if element.get(qn("w:type")) == "page"
        ]
        self.assertEqual(len(page_breaks), 1)
        self.assertIn(
            "Page 1 body",
            [paragraph.text for paragraph in document.paragraphs],
        )
        self.assertIn(
            "Page 2 body",
            [paragraph.text for paragraph in document.paragraphs],
        )

    def test_joins_cjk_lines_without_spaces(self) -> None:
        source = self._path("cjk.pdf")

        def build(pdf) -> None:
            page = pdf.new_page(width=595, height=842)
            page.insert_text(
                (72, 100),
                "你好世界\n第二行中文",
                fontsize=11,
                fontname="china-s",
            )

        self._save_pdf(source, build)

        document = Document(str(convert_pdf_to_word(source)))

        self.assertEqual(
            [paragraph.text for paragraph in document.paragraphs],
            ["你好世界第二行中文"],
        )

    def test_writes_to_explicit_target(self) -> None:
        source = self._text_pdf()
        target = self._path("nested/output.docx")

        output = convert_pdf_to_word(source, target)

        self.assertEqual(output, target)
        self.assertTrue(target.is_file())

    def test_rejects_non_docx_target(self) -> None:
        source = self._text_pdf()

        with self.assertRaises(ValueError):
            convert_pdf_to_word(source, self.directory / "output.pdf")

    def test_rejects_missing_source(self) -> None:
        with self.assertRaises(FileNotFoundError):
            convert_pdf_to_word(self.directory / "missing.pdf")

    def test_rejects_directory_source(self) -> None:
        with self.assertRaises(IsADirectoryError):
            convert_pdf_to_word(self.directory)

    def test_rejects_invalid_pdf(self) -> None:
        source = self._path("broken.pdf")
        source.write_bytes(b"this is not a PDF")

        with self.assertRaises(PdfToWordError):
            convert_pdf_to_word(source)

    def test_rejects_encrypted_pdf(self) -> None:
        source = self._path("encrypted.pdf")
        pdf = fitz.open()
        try:
            pdf.new_page().insert_text((72, 72), "secret", fontsize=11)
            pdf.save(
                str(source),
                encryption=fitz.PDF_ENCRYPT_AES_256,
                owner_pw="owner",
                user_pw="secret",
            )
        finally:
            pdf.close()

        with self.assertRaisesRegex(PdfToWordError, "password"):
            convert_pdf_to_word(source)

    def test_converts_blank_page(self) -> None:
        source = self._path("blank.pdf")

        def build(pdf) -> None:
            pdf.new_page(width=595, height=842)

        self._save_pdf(source, build)

        output = convert_pdf_to_word(source)

        self.assertTrue(output.is_file())


if __name__ == "__main__":
    unittest.main()
