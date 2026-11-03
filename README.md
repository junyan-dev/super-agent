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

Install the conversion dependencies and run the tests:

```bash
pip install -r requirements.txt
python -m unittest -v
```
