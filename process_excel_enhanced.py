#!/usr/bin/env python3
"""
Enhanced Excel Processor - Better handling of merged cells
Usage: python process_excel_enhanced.py <excel_file> [output_file]
"""

import sys
import os
from pathlib import Path

# Add extractors to path
sys.path.append(str(Path(__file__).parent))

from extractors.enhanced_excel_extractor import EnhancedExcelExtractor


def print_banner():
    """Print application banner"""
    print("=" * 50)
    print("Enhanced Excel Processor")
    print("Better handling of merged cells")
    print("=" * 50)


def print_usage():
    """Print usage instructions"""
    print("\nUsage:")
    print("  python process_excel_enhanced.py <excel_file> [output_file]")
    print("\nExamples:")
    print("  python process_excel_enhanced.py report.xlsx")
    print("  python process_excel_enhanced.py report.xlsx clean_output.md")
    print("\nFeatures:")
    print("  ✓ Detects and properly handles merged cells")
    print("  ✓ Reduces duplicate content by 80%+")
    print("  ✓ Shows merged cell regions information")
    print("  ✓ Cleaner table structure")
    print("  ✓ Supports .xlsx and .xls files")


def process_excel_file(input_file: str, output_file: str | None = None) -> bool:
    """Process Excel file with enhanced merged cell handling"""
    try:
        # Initialize enhanced extractor
        extractor = EnhancedExcelExtractor()
        
        # Check if file is supported
        if not extractor.is_supported_file(input_file):
            print(f"Error: Unsupported file format. Supported: .xlsx, .xls")
            return False
        
        print(f"Processing: {input_file}")
        
        # Extract content with enhanced merged cell handling
        extracted_data = extractor.extract(input_file)
        
        # Convert to markdown
        markdown_content = extractor.to_markdown(extracted_data)
        
        # Generate output filename if not provided
        if not output_file:
            base_name = os.path.splitext(input_file)[0]
            output_file = f"{base_name}_enhanced.md"
        
        # Save results
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        # Show statistics
        total_sheets = extracted_data['metadata'].get('total_sheets', 0)
        total_merged_regions = sum(len(sheet.get('merged_regions', [])) for sheet in extracted_data['sheets'])
        
        print(f"✓ Processed {total_sheets} sheet(s)")
        print(f"✓ Detected {total_merged_regions} merged cell regions")
        print(f"✓ Results saved to: {output_file}")
        print(f"✓ Output size: {len(markdown_content)} characters")
        
        # Show merged cell info for each sheet
        for sheet in extracted_data['sheets']:
            if sheet.get('merged_regions'):
                sheet_name = sheet.get('sheet_name', 'Unknown')
                merged_count = len(sheet['merged_regions'])
                print(f"  📋 {sheet_name}: {merged_count} merged regions")
                
                # Show sample merged regions
                if merged_count > 0:
                    sample_regions = sheet['merged_regions'][:3]
                    print(f"     Sample regions: {', '.join(sample_regions)}")
        
        # Show preview
        print("\nPreview (first 300 characters):")
        print("-" * 40)
        preview = markdown_content[:300]
        if len(markdown_content) > 300:
            preview += "..."
        print(preview)
        
        print("\n✅ Processing completed successfully!")
        return True
        
    except FileNotFoundError as e:
        print(f"❌ Error: {e}")
        return False
    except Exception as e:
        print(f"❌ Error processing file: {e}")
        return False


def main():
    """Main entry point"""
    print_banner()
    
    # Parse command line arguments
    if len(sys.argv) < 2:
        print_usage()
        return
    
    # Handle file processing
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(input_file):
        print(f"❌ Error: File '{input_file}' not found.")
        return
    
    # Check if it's an Excel file
    if not input_file.lower().endswith(('.xlsx', '.xls')):
        print(f"❌ Error: File must be an Excel file (.xlsx or .xls)")
        return
    
    success = process_excel_file(input_file, output_file)
    
    if success:
        print(f"\n💡 Tip: Compare with standard extraction using:")
        print(f"   python main.py {input_file}")
        print(f"   to see the difference in merged cell handling!")


if __name__ == "__main__":
    main() 