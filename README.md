# super-agent

dsh-superagent 智能体的工作仓库。

## Modules

### `base64_codec.py`

Dependency-free Base64 helpers built on the standard library:

```python
from base64_codec import base64_decode, base64_encode

base64_encode("hello")        # 'aGVsbG8='
base64_decode("aGVsbG8=")     # b'hello'
```

### `pdf_to_word.py`

Convert a PDF file into an editable Word (`.docx`) document:

```python
from pdf_to_word import convert_pdf_to_word

convert_pdf_to_word("report.pdf")                 # writes report.docx
convert_pdf_to_word("report.pdf", "report.docx")  # explicit target
```

The converter recreates text with its formatting, images, bordered tables and
page breaks. Scanned pages stay as images because no OCR is performed.

### `word_to_pdf.py`

Convert a Word document into a PDF file that keeps the document's layout:

```python
from word_to_pdf import convert_word_to_pdf

convert_word_to_pdf("report.docx")                # writes report.pdf
convert_word_to_pdf("report.docx", "report.pdf")  # explicit target
```

The PDF is rendered by a headless LibreOffice installation, so pages, fonts,
tables, images, headers and footers are preserved. Word (`.docx`, `.docm`)
and legacy `.doc` documents are accepted; password-protected documents raise
`WordToPdfError`.

LibreOffice is required for this conversion and is not a Python package:
install it and either put `soffice` on `PATH` or set `WORD_TO_PDF_OFFICE` to
the executable.

Install the conversion dependencies and run the tests:

```bash
pip install -r requirements.txt
python -m unittest -v
```
