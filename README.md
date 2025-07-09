# 📊 LLM-Optimized Excel Document Extractor

Converts Excel files to LLM-friendly markdown format with maximum clarity and minimal noise.

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run Extractor
```bash
# Extract single file
python main.py your_file.xlsx

# Extract with custom output name
python main.py your_file.xlsx output.md

# Demo mode (process all sample files)
python main.py --demo
```

## 📁 What it does

✅ **Perfect for LLM processing**:
- Clean markdown tables (no empty columns)
- Preserves merged cell structure  
- Handles complex Excel layouts
- Maintains business logic & formulas
- Optimized token count

✅ **Supported formats**:
- `.xlsx`, `.xls` (Excel) - **LLM-optimized processing**
- `.pdf`, `.docx` - Standard extraction

## 🎯 Example Output

**Input**: Complex Excel with merged cells, business rules
**Output**: Clean markdown with:
```markdown
Ngân Hàng TMCP Ngoại Thương Việt Nam
Chi nhánh: XXXXX (mã và tên)

| Date | Branch | CIF | Customer | Amount |
|------|--------|-----|----------|--------|
| 2024 | HN001  | C01 | Company A| 1000   |
```

## 💡 Features

- 🧹 **No duplicate columns**
- 📋 **Structured content sections** 
- 🤖 **LLM token optimized**
- 🔄 **Preserves business rules**
- 📊 **Handles merged cells correctly**

## 🛠 Requirements

- Python 3.8+
- See `requirements.txt` for packages

---
Ready to extract! 🎉 