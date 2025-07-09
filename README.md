# 📊 Document Extractor: Docling + Enhanced Excel Implementation

**Implementation of Docling for PDF/Word processing and Enhanced Excel extraction for complex spreadsheets**

## 🎯 What This Repo Does

This repository implements **two specialized document processing approaches**:

### 🐍 **1. Docling Integration**
- **Purpose**: Process PDF and Word documents  
- **Technology**: IBM's Docling library
- **Output**: Standard markdown extraction

### 📊 **2. Enhanced Excel Processing** 
- **Purpose**: Handle complex Excel files with merged cells
- **Technology**: Custom openpyxl-based solution
- **Output**: LLM-optimized markdown with clean table structure

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run Extractor
```bash
# Excel files -> Enhanced processing
python main.py your_file.xlsx

# PDF/Word files -> Docling processing  
python main.py your_file.pdf
python main.py your_file.docx

# Demo mode (test both implementations)
python main.py --demo
```

## 🔧 Implementation Details

### 📊 Enhanced Excel Processor (`enhanced_excel_extractor.py`)
- **Problem Solved**: Standard Excel extractors create messy output with duplicate columns from merged cells
- **Solution**: Custom logic to handle merged cells properly and create clean tables
- **Key Features**:
  - ✅ Preserves merged cell structure
  - ✅ Eliminates duplicate columns  
  - ✅ Maintains business logic & formulas
  - ✅ LLM-optimized token count

### 🐍 Docling Integration (`docling_extractor.py`)  
- **Purpose**: Leverage IBM's Docling for PDF/Word processing
- **Implementation**: Wrapper around Docling with consistent output format
- **Supported**: PDF, DOCX files

## 📁 Processing Logic

The main script (`main.py`) routes files based on extension:

```python
if file_ext in ['.xlsx', '.xls']:
    # Use Enhanced Excel Processor
    process_excel_for_llm(input_file, output_file)
else:
    # Use Docling Integration  
    process_non_excel_file(input_file, output_file)
```

## 🎯 Example: Enhanced Excel Output

**Input**: Complex banking Excel with merged headers
**Output**: Clean LLM-ready markdown:
```markdown
Ngân Hàng TMCP Ngoại Thương Việt Nam
Chi nhánh: XXXXX (mã và tên)

| Ngày giao dịch | Ngày hiệu lực | Mã chi nhánh | CIF | Tên khách hàng |
|----------------|---------------|--------------|-----|----------------|
| 01/01/2024     | 01/01/2024    | HN001        | C01 | Company A      |

**Phạm vi báo cáo:**
Bản ghi lịch sử TKV (KEY: LHNOTE, LHRECN)...
```

## 🛠 Technical Stack

- **Core**: Python 3.8+
- **Excel Processing**: `openpyxl` with custom merged cell handling
- **PDF/Word Processing**: `docling` library  
- **Output**: Markdown format optimized for LLM consumption

## 📦 File Structure

```
document_extractor/
├── main.py                     # Main router
├── enhanced_excel_extractor.py # Custom Excel processor  
├── docling_extractor.py        # Docling wrapper
├── requirements.txt            # Dependencies
└── sample files               # Test documents
```

---
**Two specialized implementations, one unified interface** 🎉 