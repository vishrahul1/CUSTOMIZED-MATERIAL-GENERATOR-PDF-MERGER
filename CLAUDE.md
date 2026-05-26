# CUSTOMIZED MATERIAL GENERATOR

## Project Overview

A Python desktop application that automates creation of customized educational material PDFs. It reads an Excel configuration file, merges source PDFs in order, adds headers/footers with class names and page numbers, generates a Word-based index, and produces a single final combined PDF.

## Tech Stack

- **Python 3.12+**
- **GUI:** Tkinter (built-in)
- **PDF processing:** PyPDF2, ReportLab
- **Excel reading:** pandas + openpyxl
- **Word/index generation:** python-docx
- **Distribution:** PyInstaller (`main_gui.spec`)

## Project Structure

```
CUSTOMIZED MATERIAL GENERATOR/
├── main.py                  # CLI entry point — run_main_pipeline()
├── main_gui.py              # Tkinter GUI — MergeApp class
├── main_gui.spec            # PyInstaller build config
├── pdfs_extrator.py         # Utility: flatten nested PDF folder structures
├── requirements.txt         # Python dependencies
├── modules/
│   ├── input_handler.py     # Excel reading and validation
│   ├── file_collector.py    # PDF file discovery and ordering
│   ├── pdf_merger.py        # PDF merge + header/footer overlay
│   ├── template_updater.py  # Word index template → PDF
│   └── index_manager.py     # Combine index + content into final PDF
├── resources/
│   ├── excel.xlsx           # Template/sample Excel config
│   ├── template.docx        # Word index template
│   └── source_pdfs/         # Input PDF directory
├── output/                  # Generated PDFs land here
└── logs/                    # Process logs
```

## Virtual Environment

```powershell
# Activate
.venv\Scripts\Activate.ps1

# Install deps
pip install -r requirements.txt
```

## Running the App

```powershell
# GUI (recommended)
python main_gui.py

# CLI
python main.py

# PDF folder flattener utility
python pdfs_extrator.py
```

## Building a Standalone Executable

```powershell
pyinstaller main_gui.spec
```

## Pipeline Steps (main.py → run_main_pipeline)

1. **InputHandler** — reads Excel; required columns: `File_Name`, `Chapter_Name`, `Topic_Name`, `Class`
2. **FileCollector** — scans source folder, validates all PDFs referenced in Excel exist
3. **PDFMerger.merge** — merges PDFs in order, tracks each chapter/topic start page → `merged_content.pdf`
4. **PDFMerger.add_headers_footers** — overlays class name header + page number footer → `merged_with_header.pdf`
5. **TemplateUpdater** — fills Word index template with chapter/topic page numbers → `final_index.pdf`
6. **IndexManager** — concatenates index PDF + content PDF → `final_combined.pdf`

## Excel Input Format

| Column | Description |
|---|---|
| `File_Name` | PDF filename (without path) |
| `Chapter_Name` | Chapter label for the index |
| `Topic_Name` | Topic label for the index |
| `Class` | Class/grade name (used as page header) |

## Key Implementation Notes

- Headers/footers are added via **ReportLab Canvas overlay** using `BytesIO` (not in-place editing)
- PyInstaller resource paths use `sys._MEIPASS` fallback for bundled resources
- GUI runs the pipeline in a **daemon thread** to keep the UI responsive
- All intermediate PDFs (`merged_content.pdf`, `merged_with_header.pdf`, `final_index.pdf`) are preserved in `output/` for debugging
- `requirements.txt` is UTF-16 LE encoded — re-save as UTF-8 if pip has trouble reading it

## Common Issues

- **pip can't read requirements.txt** — file is UTF-16 encoded; use `pip install chardet et_xmlfile lxml numpy openpyxl pandas pillow PyPDF2 python-dateutil python-docx pytz reportlab six typing_extensions tzdata` directly, or re-save the file as UTF-8
- **Missing PDFs** — FileCollector will raise an error listing which files weren't found; verify filenames in Excel match actual filenames exactly (case-sensitive on Linux/macOS)
- **Word → PDF conversion** — requires Microsoft Word or LibreOffice installed; `template_updater.py` uses `docx2pdf` or a similar converter
