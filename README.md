# pdf-crop-tool

Desktop GUI for cropping PDFs and building worksheet projects. Import PDFs or PNG folders, set crop margins (global or per-page), tag pages, then assemble worksheets from multiple sources with drag-and-drop reordering.

Built with Python, CustomTkinter, and PyMuPDF.

## Features

- **Source management** - import PDFs and PNG folders, organize in nested folders, search by name/tag
- **Visual crop editor** - interactive margin sliders with live preview and crop lines
- **Per-page overrides** - set custom crops on individual pages while keeping a global default
- **Page tagging** - tag pages manually or with AI auto-tagging (via OpenRouter vision models)
- **Project builder** - assemble worksheets from tagged source pages, filter/reorder, export to PDF or PNG
- **Presets** - save and reuse crop configurations

## Setup

```bash
pip install -r requirements.txt
python pdf_cropper.py
```

## Build (Windows exe)

```bash
build_windows.bat
# Output: dist/pdf-crop-tool.exe
```
