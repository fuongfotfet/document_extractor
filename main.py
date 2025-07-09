#!/usr/bin/env python3
"""
LLM-Optimized Excel Document Extractor
Converts Excel files to LLM-friendly format with minimal noise and maximum clarity
"""

import sys
import os
from pathlib import Path
from enhanced_excel_extractor import EnhancedExcelExtractor
from docling_extractor import DoclingExtractor
from typing import Optional

def print_banner():
    """Print application banner"""
    print("🤖" + "=" * 60)
    print("📊 LLM-Optimized Excel Document Extractor")
    print("🎯 Maximum clarity, minimal noise for AI understanding")
    print("🤖" + "=" * 60)

def print_usage():
    """Print usage instructions"""
    print("\nUsage:")
    print("  python main.py <excel_file> [output_file]")
    print("  python main.py --demo")
    print("\nExamples:")
    print("  python main.py report.xlsx")
    print("  python main.py report.xlsx analysis.md")
    print("  python main.py --demo")
    print("\nOutput Format:")
    print("  📋 LLM-Optimized: Clean structure, no duplicates, clear context")
    print("  📄 Markdown format for easy LLM consumption")

def process_excel_for_llm(input_file: str, output_file: Optional[str] = None) -> bool:
    """
    Process Excel file specifically for LLM consumption
    """
    try:
        print(f"🔍 Analyzing Excel file: {input_file}")
        
        # Use Enhanced Excel Extractor for maximum control
        extractor = EnhancedExcelExtractor()
        
        # Extract content with full structure preservation
        extracted_data = extractor.extract(input_file)
        print(f"✅ Extracted {extracted_data['metadata']['sheets_count']} sheets")
        print(f"📊 Found {sum(len(sheet['merged_cells']) for sheet in extracted_data['sheets'].values())} merged cells")
        
        # Convert to LLM-optimized format
        print("🤖 Converting to LLM-optimized format...")
        llm_content = extractor.to_llm_optimized(extracted_data)
        
        # Generate output filename if not provided
        if not output_file:
            base_name = os.path.splitext(input_file)[0]
            output_file = f"{base_name}_llm_optimized.md"
        
        # Save results
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(llm_content)
        
        print(f"✅ LLM-optimized content saved to: {output_file}")
        
        # Show analysis preview
        print("\n📖 Preview (first 500 characters):")
        print("-" * 50)
        preview = llm_content[:500]
        if len(llm_content) > 500:
            preview += "..."
        print(preview)
        
        # Show format benefits
        print(f"\n🎯 LLM Benefits:")
        print(f"   📝 Content length: {len(llm_content):,} characters")
        print(f"   📊 Estimated tokens: ~{len(llm_content.split()):,}")
        print(f"   🧹 Pure Excel content: No section headers")
        print(f"   🚫 No duplicate columns or noise")
        
        return True
        
    except FileNotFoundError as e:
        print(f"❌ Error: {e}")
        return False
    except Exception as e:
        print(f"❌ Error processing file: {e}")
        return False

def process_non_excel_file(input_file: str, output_file: Optional[str] = None) -> bool:
    """
    Process non-Excel files using standard Docling
    """
    try:
        print(f"🔍 Processing non-Excel file: {input_file}")
        extractor = DoclingExtractor()
        
        if not extractor.is_supported_file(input_file):
            print(f"❌ Unsupported file format. Supported: {', '.join(extractor.get_supported_extensions())}")
            return False
        
        # Extract content
        extracted_data = extractor.extract(input_file)
        content = extractor.to_markdown(extracted_data)
        
        # Generate output filename
        if not output_file:
            base_name = os.path.splitext(input_file)[0]
            output_file = f"{base_name}_extracted.md"
        
        # Save results
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"✅ Content saved to: {output_file}")
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def run_demo():
    """Demo with available sample files"""
    print("🧪 Demo Mode - Processing sample files...")
    
    # Look for sample files
    sample_files = []
    for ext in ['.xlsx', '.xls', '.pdf', '.docx']:
        for file in Path('.').glob(f'*{ext}'):
            sample_files.append(str(file))
    
    if not sample_files:
        print("❌ No sample files found.")
        print("💡 Add some .xlsx, .pdf, or .docx files to test.")
        return False
    
    print(f"📁 Found {len(sample_files)} sample files:")
    
    excel_files = [f for f in sample_files if f.endswith(('.xlsx', '.xls'))]
    other_files = [f for f in sample_files if not f.endswith(('.xlsx', '.xls'))]
    
    # Process Excel files with LLM optimization
    if excel_files:
        print(f"\n🚀 Processing {len(excel_files)} Excel files with LLM optimization:")
        for file in excel_files[:3]:  # Limit to 3 files
            print(f"\n📊 Processing: {file}")
            success = process_excel_for_llm(file)
            print("✅ Success" if success else "❌ Failed")
    
    # Process other files with standard extraction
    if other_files:
        print(f"\n📄 Processing {len(other_files)} other files with standard extraction:")
        for file in other_files[:3]:  # Limit to 3 files
            print(f"\n📄 Processing: {file}")
            success = process_non_excel_file(file)
            print("✅ Success" if success else "❌ Failed")
    
    return True

def main():
    """Main entry point"""
    print_banner()
    
    # Parse command line arguments
    if len(sys.argv) < 2:
        print_usage()
        return
    
    # Handle demo mode
    if sys.argv[1] == '--demo':
        run_demo()
        return
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(input_file):
        print(f"❌ Error: File '{input_file}' not found.")
        return
    
    # Determine processing method based on file type
    file_ext = os.path.splitext(input_file)[1].lower()
    
    if file_ext in ['.xlsx', '.xls']:
        print("🎯 Mode: LLM-Optimized Excel Processing")
        process_excel_for_llm(input_file, output_file)
    else:
        print("📄 Mode: Standard Document Processing")
        process_non_excel_file(input_file, output_file)

if __name__ == "__main__":
    main() 