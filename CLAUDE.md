# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Setup

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Running the app

```powershell
.venv\Scripts\python.exe "Juntador de PDFs.py"
```

## Building the executable

```powershell
pip install pyinstaller
.venv\Scripts\python.exe -m PyInstaller --onefile --noconsole --add-data "poppler_bin;poppler_bin" "Juntador de PDFs.py"
```

Output goes to `dist\Juntador de PDFs.exe`. The `poppler_bin` folder must be alongside the exe for PDF preview to work.

## Architecture

The entire application lives in a single file: `Juntador de PDFs.py`.

**`PDFMergerApp`** is the only class. It holds:
- `self.arquivos_pdf` — ordered list of `{'nome': str, 'caminho': str}` dicts representing the merge queue
- `self.pasta_origem` / `self.arquivo_destino` — `tk.StringVar` for the two path inputs

The UI is built with `tkinter`/`ttk` and split into sections created by dedicated methods (`criar_secao_selecao`, `criar_secao_arquivos`, `criar_secao_preview`, `criar_secao_progresso`). The left panel is the file list with reorder controls; the right panel is a canvas-based PDF preview.

PDF merging runs in a daemon thread (`juntar_pdfs_thread` → `juntar_pdfs`) to keep the UI responsive. It uses `PyPDF2.PdfMerger`.

PDF preview is optional: `pdf2image` and `poppler_bin` are required. The import is guarded at module level (`PDF2IMAGE_AVAILABLE` flag), and preview features degrade gracefully if the library is absent. `convert_from_path` is called with `poppler_path` implicitly resolved from the system PATH or bundled binary.

## Dependencies

| Package | Purpose |
|---------|---------|
| `PyPDF2` | PDF merge and OCR page assembly |
| `Pillow` | Image display in preview canvas |
| `pdf2image` | Converts PDF pages to images (preview + OCR pipeline) |
| `pytesseract` | OCR via Tesseract — wrapper for `image_to_pdf_or_hocr` |
| `poppler_bin/` | Native binaries required by pdf2image (bundled in repo) |

`pytesseract` also requires **Tesseract OCR** installed on the system (not a Python package).
Windows installer: github.com/tesseract-ocr/tesseract — install the `por` language pack for Portuguese OCR.

## Building the executable (updated)

The PyInstaller command must bundle both `poppler_bin` and the Tesseract data directory if you want OCR to work without a system installation:

```powershell
.venv\Scripts\python.exe -m PyInstaller --onefile --noconsole --add-data "poppler_bin;poppler_bin" "Juntador de PDFs.py"
```
