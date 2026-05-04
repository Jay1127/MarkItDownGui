# MarkItDown GUI

A small local desktop app for testing Microsoft MarkItDown conversion options.

## Prepare

Create and activate a virtual environment first.

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

## Install

Install only the packages needed for the test case.

### Basic Document Conversion

Use this for common document conversion tests such as PDF, DOCX, XLSX, PPTX, HTML, CSV, JSON, XML, ZIP, and EPUB.

```powershell
pip install 'markitdown[all]'
```

### LLM Image Descriptions

Use this when testing image descriptions with an LLM.

```powershell
pip install 'markitdown[all]' openai
```

Set `OPENAI_API_KEY` before running the app.

```powershell
$env:OPENAI_API_KEY = "your-api-key"
```

### OCR Plugin

Use this when testing the `markitdown-ocr` plugin.

```powershell
pip install 'markitdown[all]' openai markitdown-ocr
```

Use `requirements.txt` only when you want to install every test dependency at once.

```powershell
pip install -r requirements.txt
```

## Run

```powershell
python app.py
```

## Test Items

- Convert PDF to Markdown
- Convert DOCX headings and lists
- Convert XLSX tables
- Extract PPTX slide text
- Extract image metadata or text
- Enable plugins and select individual plugins
- Test the `keep_data_uris` option
- Test extension, MIME type, and charset hints
- Generate LLM image descriptions with `OPENAI_API_KEY`
- Override the exiftool path
- Test DOCX style map behavior

The app writes the converted Markdown to the selected output path and shows a preview in the window.
