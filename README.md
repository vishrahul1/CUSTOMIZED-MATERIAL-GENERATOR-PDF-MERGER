# Customised Material Generator

A Python desktop application that automates the creation of customised educational PDF material. It reads an Excel sheet, finds the required PDF files, merges them in order, adds header/footer information, generates a chapter/topic index, and produces a final combined PDF.

## Features

- Reads Excel input with required columns:
  - `File_Name`
  - `Chapter_Name`
  - `Topic_Name`
  - `Class`
- Recursively scans the source PDF folder for matching files
- Merges PDFs while tracking chapter/topic page positions
- Adds custom header/footer text and page numbers
- Inserts chapter title divider pages
- Generates a final index PDF and combines it with the content PDF
- Includes a Tkinter GUI for easier use

## Tech Stack

- Python 3.12+
- Tkinter (GUI)
- PyPDF2 / pypdf
- ReportLab
- pandas + openpyxl
- python-docx
- PyInstaller (for building executables)

## Project Structure

- `main.py` — CLI pipeline entry point
- `main_gui.py` — GUI application
- `main_gui.spec` — PyInstaller build configuration
- `modules/` — core processing logic
- `resources/` — sample Excel, templates, and source PDFs
- `output/` — generated PDFs
- `logs/` — runtime logs

## Setup

1. Create and activate a virtual environment.
2. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

3. Make sure your Excel file and source PDF folder are ready.

## Running the Application

### GUI

```powershell
python main_gui.py
```

### CLI

```powershell
python main.py
```

## Output

The pipeline creates intermediate and final PDFs such as:

- `merged_content.pdf`
- `merged_with_header.pdf`
- `chapter_titles.pdf`
- `final_index.pdf`
- `final_combined.pdf`

## Notes

- The app expects the Excel file to contain exactly the required columns.
- PDF filenames must match the values in the Excel sheet.
- Word-to-PDF conversion is required when generating the index document.
- If the requirements file causes issues, reinstall dependencies manually using the package list in the project documentation.
