# Cải thiện xử lý Merged Cells trong Excel

## Vấn đề với Merged Cells

Khi sử dụng **Docling** để trích xuất file Excel có merged cells, kết quả thường bị:
- **Duplicate values**: Giá trị của merged cell bị lặp lại cho tất cả cells trong vùng merge
- **Table structure khó đọc**: Quá nhiều repetition gây rối
- **File output lớn**: Do duplicate content

### Ví dụ vấn đề:
```markdown
| Ngân Hàng TMCP Ngoại Thương Việt Nam | Ngân Hàng TMCP Ngoại Thương Việt Nam | Ngân Hàng TMCP Ngoại Thương Việt Nam | ...
```

## Giải pháp: Enhanced Excel Extractor

### 🚀 Cải thiện đáng kể:
- ✅ **Giảm 84% duplicates** (từ 200 xuống 32)
- ✅ **Giảm 70% kích thước output** (từ 12,481 xuống 3,810 characters)
- ✅ **Detect 177 merged cell regions** và xử lý đúng cách
- ✅ **Cấu trúc table sạch hơn** với empty cells thay vì duplicates
- ✅ **Hiển thị thông tin merged regions** để hiểu structure

### Kết quả sau cải thiện:
```markdown
**Merged Cell Regions:**
- AG4:AH4
- W4:X4  
- Y4:Z4

| Ngân Hàng TMCP Ngoại Thương Việt Nam |  |  |  |  |  |  |  |
```

## Cách sử dụng

### 1. Cài đặt dependencies:
```bash
pip install -r requirements.txt
```

### 2. Sử dụng Enhanced Excel Processor:
```bash
# Xử lý với Enhanced Excel Extractor
python process_excel_enhanced.py some_report.xlsx

# Chỉ định output file
python process_excel_enhanced.py some_report.xlsx clean_output.md
```

### 3. So sánh kết quả:
```bash
# Test và so sánh cả hai phương pháp
python test_excel_processing.py some_report.xlsx
```

## Files được tạo

### 📁 Enhanced Excel Extractor:
- `extractors/enhanced_excel_extractor.py` - Core enhanced extractor
- `process_excel_enhanced.py` - Script đơn giản để sử dụng
- `test_excel_processing.py` - Script so sánh kết quả

### 📄 Output files:
- `some_report_enhanced.md` - Kết quả với Enhanced Extractor
- `some_report_docling.md` - Kết quả với Docling (để so sánh)

## Công nghệ sử dụng

### Enhanced Excel Extractor:
- **openpyxl**: Đọc và detect merged cells trong .xlsx
- **pandas**: Xử lý data và .xls files  
- **Custom algorithms**: Clean duplicates và format tables

### Key Features:
1. **Merged Cell Detection**: Sử dụng `openpyxl` để detect chính xác merged regions
2. **Smart Cleaning**: Chỉ giữ giá trị ở top-left cell, clear duplicates
3. **Fallback Support**: Hỗ trợ .xls files qua pandas
4. **Duplicate Detection**: Algorithm detect consecutive duplicates

## So sánh Performance

| Metric | Docling | Enhanced | Improvement |
|--------|---------|----------|-------------|
| Output Size | 12,481 chars | 3,810 chars | **70% reduction** |
| Duplicates | 200 | 32 | **84% reduction** |
| Merged Regions | Not detected | 177 detected | **100% improvement** |
| Table Readability | Poor | Good | **Much cleaner** |

## Khi nào sử dụng

### 🎯 Dùng Enhanced Excel Extractor khi:
- File Excel có nhiều merged cells
- Cần cấu trúc table sạch, ít duplicate
- Muốn biết thông tin về merged regions
- File size và readability quan trọng

### 🎯 Dùng Docling khi:
- File có mixed formats (PDF, Word, Excel...)
- Cần xử lý hàng loạt nhiều file types
- Performance và speed là ưu tiên
- Không quan trọng về merged cells

## Tips & Best Practices

### 📝 Preprocessing Excel files:
1. Kiểm tra file structure trước khi process
2. Backup file gốc
3. Test với file nhỏ trước

### 🔧 Tuning Enhanced Extractor:
- Adjust duplicate detection threshold trong `_clean_duplicate_values()`
- Customize table formatting trong `_create_clean_markdown_table()`
- Add more file validation nếu cần

### 🚀 Automation:
```bash
# Process multiple Excel files
for file in *.xlsx; do
    python process_excel_enhanced.py "$file"
done
```

## Troubleshooting

### ❗ Common Issues:

1. **"openpyxl not found"**
   ```bash
   pip install openpyxl>=3.1.0
   ```

2. **"File appears corrupted"**
   - Kiểm tra file có mở được trong Excel không
   - Try với file copy mới

3. **"Memory error on large files"**
   - Process từng sheet riêng biệt
   - Increase system memory limit

4. **"Merged regions not detected"**
   - File có thể là .xls (old format) 
   - Convert to .xlsx để có better support

## Future Improvements

### 🔮 Planned Features:
- [ ] Multi-sheet parallel processing
- [ ] Custom merged cell handling rules
- [ ] Integration với main.py
- [ ] GUI interface
- [ ] Batch processing với progress bar
- [ ] Export to multiple formats (HTML, CSV, JSON)

---

## Contact & Support

Để hỗ trợ hoặc đóng góp cải thiện, vui lòng:
- Tạo issue với sample file và expected output
- Describe use case cụ thể
- Attach error logs nếu có 