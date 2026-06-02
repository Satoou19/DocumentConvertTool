# 📄 Document Converter

A lightweight desktop app to convert documents between **Markdown, Excel, and Word** formats — with drag & drop support.

## Features

- **MD → Excel** — Each table in Markdown becomes a separate sheet, with styled headers and auto-fit columns
- **MD → Word** — Converts headings, tables, lists, and inline bold with Arial font styling
- **Excel → MD** — Each sheet becomes a Markdown section with a table
- **Word → MD** — Extracts content from `.docx` into clean Markdown
- Drag & drop file input
- Auto-fills output path based on input file
- Runs conversion on a background thread (UI stays responsive)

## Requirements

- Python 3.10+
- Windows / macOS / Linux

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python md_to_excel.py
```

Select a conversion mode, drag & drop your file (or use Browse), choose an output location, and click **Convert**.

## Build as Executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "Document Converter" --icon=favicon.ico md_to_excel.py
```

Output will be in the `dist/` folder.

## Project Structure

```
📁 document-converter
├── md_to_excel.py       # Main app
├── requirements.txt
├── README.md
├── .gitignore
└── favicon.ico          # Optional
```

## Dependencies

| Package | Purpose |
|---|---|
| customtkinter | Modern UI framework |
| tkinterdnd2 | Drag & drop support |
| pandas | Excel read/write |
| openpyxl | Excel styling |
| python-docx | Word generation |
| mammoth | Word → Markdown conversion |