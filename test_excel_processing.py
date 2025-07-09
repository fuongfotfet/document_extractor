#!/usr/bin/env python3
"""
Test script to compare Excel processing methods for merged cells
"""

import os
import sys
from pathlib import Path

# Add extractors to path
sys.path.append(str(Path(__file__).parent))

from docling_extractor import DoclingExtractor
from extractors.enhanced_excel_extractor import EnhancedExcelExtractor


def test_excel_processing(excel_file: str):
    """
    Test both extraction methods and compare results
    """
    if not os.path.exists(excel_file):
        print(f"Error: File '{excel_file}' not found.")
        return
    
    print("=" * 60)
    print(f"Testing Excel Processing for: {excel_file}")
    print("=" * 60)
    
    # Test 1: Docling Extractor (current method)
    print("\n1. DOCLING EXTRACTOR (Current Method)")
    print("-" * 40)
    try:
        docling_extractor = DoclingExtractor()
        docling_result = docling_extractor.extract(excel_file)
        docling_markdown = docling_extractor.to_markdown(docling_result)
        
        # Save Docling result
        with open(f"{os.path.splitext(excel_file)[0]}_docling.md", 'w', encoding='utf-8') as f:
            f.write(docling_markdown)
        
        print(f"✓ Processed successfully")
        print(f"✓ Output length: {len(docling_markdown)} characters")
        print(f"✓ Saved to: {os.path.splitext(excel_file)[0]}_docling.md")
        
        # Show preview
        print("\nPreview (first 200 chars):")
        print(docling_markdown[:200] + "..." if len(docling_markdown) > 200 else docling_markdown)
        
    except Exception as e:
        print(f"✗ Error: {e}")
    
    # Test 2: Enhanced Excel Extractor (new method)
    print("\n\n2. ENHANCED EXCEL EXTRACTOR (New Method)")
    print("-" * 40)
    try:
        enhanced_extractor = EnhancedExcelExtractor()
        enhanced_result = enhanced_extractor.extract(excel_file)
        enhanced_markdown = enhanced_extractor.to_markdown(enhanced_result)
        
        # Save Enhanced result
        with open(f"{os.path.splitext(excel_file)[0]}_enhanced.md", 'w', encoding='utf-8') as f:
            f.write(enhanced_markdown)
        
        print(f"✓ Processed successfully")
        print(f"✓ Output length: {len(enhanced_markdown)} characters")
        print(f"✓ Saved to: {os.path.splitext(excel_file)[0]}_enhanced.md")
        
        # Show merged cell info
        for sheet in enhanced_result['sheets']:
            if sheet.get('merged_regions'):
                print(f"✓ Detected {len(sheet['merged_regions'])} merged cell regions")
                print(f"  Sample regions: {sheet['merged_regions'][:3]}")
        
        # Show preview
        print("\nPreview (first 200 chars):")
        print(enhanced_markdown[:200] + "..." if len(enhanced_markdown) > 200 else enhanced_markdown)
        
    except Exception as e:
        print(f"✗ Error: {e}")
    
    # Comparison
    print("\n\n3. COMPARISON")
    print("-" * 40)
    try:
        # Count tables/duplicates
        docling_table_count = docling_markdown.count('|') if 'docling_markdown' in locals() else 0
        enhanced_table_count = enhanced_markdown.count('|') if 'enhanced_markdown' in locals() else 0
        
        # Count duplicate patterns (same text repeated)
        if 'docling_markdown' in locals() and 'enhanced_markdown' in locals():
            docling_duplicates = count_duplicate_patterns(docling_markdown)
            enhanced_duplicates = count_duplicate_patterns(enhanced_markdown)
            
            print(f"Docling - Table markers (|): {docling_table_count}")
            print(f"Enhanced - Table markers (|): {enhanced_table_count}")
            print(f"Docling - Potential duplicates: {docling_duplicates}")
            print(f"Enhanced - Potential duplicates: {enhanced_duplicates}")
            
            if enhanced_duplicates < docling_duplicates:
                print("✓ Enhanced extractor shows improvement in reducing duplicates!")
            elif enhanced_duplicates == docling_duplicates:
                print("→ Similar duplicate levels between methods")
            else:
                print("⚠ Enhanced extractor may need further tuning")
        
    except Exception as e:
        print(f"Comparison error: {e}")
    
    print("\n" + "=" * 60)
    print("Test completed! Check the generated .md files for detailed comparison.")


def count_duplicate_patterns(text: str) -> int:
    """
    Count potential duplicate patterns in text
    """
    lines = text.split('\n')
    duplicate_count = 0
    
    for line in lines:
        if '|' in line:  # Table line
            cells = [cell.strip() for cell in line.split('|')[1:-1]]  # Remove empty first/last
            if len(cells) > 1:
                # Count consecutive duplicates
                for i in range(1, len(cells)):
                    if cells[i] == cells[i-1] and cells[i].strip() and cells[i] != 'None':
                        duplicate_count += 1
    
    return duplicate_count


def main():
    """Main function"""
    print("Excel Processing Comparison Tool")
    print("Compares Docling vs Enhanced Excel Extractor")
    
    if len(sys.argv) < 2:
        print("\nUsage:")
        print("  python test_excel_processing.py <excel_file>")
        print("\nExample:")
        print("  python test_excel_processing.py some_report.xlsx")
        return
    
    excel_file = sys.argv[1]
    test_excel_processing(excel_file)


if __name__ == "__main__":
    main() 