# Document Processor - Universal File Extractor

A modular system to extract text and tables from various file formats and convert them to Markdown.

## Project Structure

```
├── main.py                     # Main file - auto file type detection
├── requirements.txt            # Dependencies
├── extractors/                 # Package containing extractor modules
│   ├── __init__.py
│   ├── base_extractor.py       # Common base class
│   ├── excel_extractor.py      # Excel processing (.xlsx, .xls)
│   ├── docx_extractor.py       # Word document processing (.docx) 
│   └── pdf_extractor.py        # PDF processing (.pdf)
└── sample files/               # Sample files for testing
    ├── test.xlsx
    ├── sample.docx
    └── sample_with_table.pdf
```

## Usage

### 1. Command Line Mode

```bash
# Process file with auto-generated output name
python main.py input_file.xlsx

# Process file with custom output name  
python main.py input_file.docx output_file.md

# Show help
python main.py
```

### 2. Demo Mode

```bash
# Test with all sample files
python main.py --demo
```

## Supported Formats

| Format | Extensions | Features |
|--------|------------|----------|
| **Excel** | .xlsx, .xls | Multi-sheets, Tables, Headers |
| **Word** | .docx | Text + Tables extraction |
| **PDF** | .pdf | Text + GMFT table extraction + OCR fallback |

## Installation

1. **Install Python dependencies:**
```bash
pip install -r requirements.txt
```

> **Note**: This will automatically install GMFT for advanced table extraction using Microsoft's Table Transformers.

2. **OCR Support (for scanned PDFs - fallback only):**

**macOS:**
```bash
brew install tesseract tesseract-lang
```

**Ubuntu/Debian:**
```bash
sudo apt-get install tesseract-ocr tesseract-ocr-vie
```

**Windows:**
- Download Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki
- Add Tesseract to system PATH

> **Note**: The system prioritizes GMFT for table extraction (faster, more accurate). OCR is used as fallback for scanned PDFs only.

## Examples

```bash
# Excel file
python main.py test.xlsx

# Word document  
python main.py sample.docx

# PDF file
python main.py sample_with_table.pdf

# Test all formats
python main.py --demo
```

## Features

- **Auto file type detection**
- **Modular architecture** - Easy to extend
- **GMFT table extraction** - Advanced ML-based table detection using Microsoft's Table Transformers
- **Header recognition** - Intelligent header detection  
- **Markdown output** - Clean, formatted output
- **Error handling** - Robust exception handling
- **Command line interface** - Easy to use
- **OCR fallback** - Support for scanned PDFs when GMFT cannot extract tables

## Quick Test

```bash
# Test immediately with available sample files
python main.py --demo
```

---

**Ready to extract!** 