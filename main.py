#!/usr/bin/env python3
"""
Document Processor - Universal File Extractor using Docling
Supports: PDF, DOCX, PPTX, XLSX, HTML, WAV, MP3, images and more
"""

import sys
import os
from pathlib import Path
from docling_extractor import DoclingExtractor
from typing import Optional

def print_banner():
    """Print application banner"""
    print("=" * 50)
    print("Document Processor - Universal File Extractor")
    print("Powered by Docling")
    print("=" * 50)

def print_usage():
    """Print usage instructions"""
    print("\nUsage:")
    print("  python main.py <input_file> [output_file]")
    print("  python main.py --demo")
    print("\nExamples:")
    print("  python main.py document.pdf")
    print("  python main.py presentation.pptx output.md")
    print("  python main.py spreadsheet.xlsx")
    print("  python main.py --demo")
    print("\nSupported formats:")
    print("  PDF, DOCX, PPTX, XLSX, HTML, images (PNG, JPG), audio (WAV, MP3)")

def process_file(input_file: str, output_file: Optional[str] = None) -> bool:
    """Process a single file"""
    try:
        # Initialize extractor
        extractor = DoclingExtractor()
        
        # Check if file is supported
        if not extractor.is_supported_file(input_file):
            print(f"Error: Unsupported file format. Supported: {', '.join(extractor.get_supported_extensions())}")
            return False
        
        print(f"Processing: {input_file}")
        
        # Extract content
        extracted_data = extractor.extract(input_file)
        
        # Convert to markdown
        markdown_content = extractor.to_markdown(extracted_data)
        
        # Generate output filename if not provided
        if not output_file:
            base_name = os.path.splitext(input_file)[0]
            output_file = f"{base_name}_extracted.md"
        
        # Save results
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        print(f"Results saved to: {output_file}")
        print("Processing completed successfully!")
        
        # Show preview
        print("\nPreview (first 300 characters):")
        print("-" * 40)
        preview = markdown_content[:300]
        if len(markdown_content) > 300:
            preview += "..."
        print(preview)
        
        return True
        
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return False
    except Exception as e:
        print(f"Error processing file: {e}")
        return False

def run_demo():
    """Run demo with sample files"""
    print("Demo Mode - Testing available sample files...")
    
    # Look for sample files
    sample_files = []
    for ext in ['.pdf', '.docx', '.xlsx', '.pptx', '.html', '.png', '.jpg']:
        for file in Path('.').glob(f'*{ext}'):
            sample_files.append(str(file))
    
    if not sample_files:
        print("No sample files found in current directory.")
        print("Try adding some PDF, DOCX, XLSX, or image files to test.")
        return False
    
    print(f"Found {len(sample_files)} sample files:")
    for file in sample_files[:5]:  # Limit to first 5 files
        print(f"  - {file}")
        success = process_file(file)
        if success:
            print("✓ Success\n")
        else:
            print("✗ Failed\n")
    
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
    
    # Handle file processing
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return
    
    process_file(input_file, output_file)

if __name__ == "__main__":
    main() 