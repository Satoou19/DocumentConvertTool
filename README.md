# Document Converter Workspace

![Python](https://img.shields.io/badge/Python-3.12%20--%203.13-blue)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

A modern desktop workspace for editing and converting documents between **Markdown**, **Excel**, and **Word** formats.

The application provides a unified Markdown-centric workflow, allowing users to extract content from Office documents, edit it in Markdown, and export it back into structured formats.

---

## Screenshot

![Application Screenshot](screenshot.png)

---

## Features

### Document Conversion

* **Markdown → Excel (.xlsx)**

  * Styled worksheet generation
  * Frozen header row
  * Auto-sized columns
  * Auto-filter support

* **Markdown → Word (.docx)**

  * Heading support
  * Lists support
  * Bold text support
  * Table rendering

* **Excel (.xlsx) → Markdown**

  * Multi-sheet extraction
  * Markdown table generation

* **Word (.docx) → Markdown**

  * Clean document extraction
  * Markdown-friendly formatting

### Workspace Features

* Unified Markdown editor
* Drag & drop file support
* Live content extraction on load
* One-click document opening
* Background-thread conversion pipeline
* Cross-platform support

---

## Supported Formats

| Input          | Output              | Status      |
| -------------- | ------------------- | ----------- |
| Markdown (.md) | Excel (.xlsx)       | ✅ Available |
| Markdown (.md) | Word (.docx)        | ✅ Available |
| Excel (.xlsx)  | Markdown (.md)      | ✅ Available |
| Word (.docx)   | Markdown (.md)      | ✅ Available |
| CSV (.csv)     | Markdown (.md)      | 🔄 Planning |
| Markdown (.md) | CSV (.csv)          | 🔄 Planning |
| PDF (.pdf)     | Markdown (.md)      | 🔄 Planning |
| Markdown (.md) | HTML (.html)        | 🔄 Planning |
| Markdown (.md) | PDF (.pdf)          | 🔄 Planning |

---

## Requirements

* Python 3.12 – 3.13
* Windows / macOS / Linux

---

## Quick Start

Clone the repository:

```bash
git clone https://github.com/duyphan1410/DocumentConvertTool
cd DocumentConvertTool
```

Create, activate virtual environment, install dependencies, and run:

### Windows (PowerShell)
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
python run.py
```

### macOS / Linux
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python run.py
```

---

## Drag & Drop

The application supports drag-and-drop input files.

### Supported Behaviors

* Drop a `.md` file to edit it directly.
* Drop a `.docx` file to extract its content into Markdown.
* Drop a `.xlsx` file to extract worksheets into Markdown tables.
* Paths containing spaces or Unicode characters are supported.

---

## Build Executable

### Install PyInstaller

```bash
pip install pyinstaller
```

### Windows

```cmd
pyinstaller --onefile --windowed --name "Document Converter" --icon=favicon.ico run.py
```

### macOS

```bash
pyinstaller --onefile --windowed --name "Document Converter" run.py
```

### Linux

```bash
pyinstaller --onefile --name "Document Converter" run.py
```

Build output:

```text
dist/
```

---

## Project Structure

```text
DocumentConvertTool/
│
├── src/
│   ├── main.py
│   │
│   ├── core/
│   │   ├── extractors.py
│   │   └── converters.py
│   │
│   ├── services/
│   │   ├── conversion_service.py
│   │   └── file_loader.py
│   │
│   ├── ui/
│   │   └── app.py
│   │
│   └── utils/
│       └── env.py
│
├── run.py
├── requirements.txt
└── README.md
```

### Directory Overview

| Path                     | Purpose                              |
| ------------------------ | ------------------------------------ |
| `src/main.py`            | Application entry point & initialization |
| `src/core/extractors.py` | Office → Markdown extraction         |
| `src/core/converters.py` | Markdown → Office conversion         |
| `src/services/conversion_service.py` | Conversion and validation service |
| `src/services/file_loader.py` | File ingestion and extraction service |
| `src/ui/app.py`          | Main GUI                             |
| `src/utils/env.py`       | UTF-8 encoding & Tcl/Tk path configuration |
| `run.py`                 | Launcher script                      |

---

## Dependencies

| Library       | Purpose                  |
| ------------- | ------------------------ |
| customtkinter | Modern UI framework      |
| tkinterdnd2   | Drag & drop support      |
| pandas        | Data processing          |
| openpyxl      | Excel export/import      |
| python-docx   | Word document generation |
| mammoth       | Word document extraction |
| markdown2     | Markdown → HTML conversion |

---

## Roadmap

### ✅ P0 — Stabilization (Completed)

* [x] Fix drag & drop path parser
* [x] File extension validation
* [x] Overwrite confirmation
* [x] Dependency fallback handling
* [x] File size warning
* [x] Progress indicator
* [x] Unsaved changes warning
* [x] Word → Markdown: images replaced with [image] placeholder

### Phase 1 — UX Improvements

* [ ] CSV ↔ Markdown support
* [ ] Smart table validator
* [ ] Search & replace panel
* [ ] Markdown syntax highlighting
* [ ] Autosave draft (restore when reopening app)

### Phase 2 — Format Expansion

* [ ] CSV → Markdown (extract like Excel)
* [ ] HTML export (with GitHub Markdown CSS styling)
* [ ] HTML preview in editor
* [ ] PDF → Markdown (requires Java 11+)
* [ ] Markdown → PDF (weasyprint)

### Phase 3 — Polish & Release

* [ ] Batch conversion
* [ ] Resizable window
* [ ] Multi-document tabs
* [ ] Conversion presets (save frequently-used conversion configs)
* [ ] v1.0 executable release (PyInstaller)

---

## Known Limitations

* **Large files:** File preview is truncated at 500KB to prevent UI lag. Full content will still be converted.
* **Word documents with images:** Images are replaced with `[image]` placeholder text since inline image handling in Markdown is limited.
* **Complex Word formatting:** Some advanced Word styles (columns, text boxes, etc.) may not be fully preserved in Markdown conversion.
* **PDF support:** Coming in Phase 2. PDF import requires Java 11+, PDF export uses weasyprint.
* **CSV support:** Coming in Phase 1.

---

## Development

Recommended branch strategy:

```text
main
├── fix/p0-stabilization
├── feature/progress-api
├── feature/html-export
└── feature/plugin-system
```

Commit example:

```bash
git commit -m "feat: add html export engine"
git commit -m "fix: improve drag drop parser"
```

---

## Contributing

Contributions are welcome.

Before opening a Pull Request:

1. Follow PEP 8 conventions.
2. Add type hints to new APIs.
3. Test conversion workflows.
4. Keep UI responsive during long-running operations.

---

## License

MIT License

You are free to use, modify, distribute, and build upon this project under the terms of the MIT License.
