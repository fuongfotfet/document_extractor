"""
Unified Docling Extractor - Handles all file formats using Docling
Supports: PDF, DOCX, PPTX, XLSX, HTML, WAV, MP3, images and more
"""

import os
from typing import Dict, Any
from docling.document_converter import DocumentConverter

class DoclingExtractor:
    """
    Unified extractor using Docling for all supported file formats
    """
    
    def __init__(self):
        self.converter = DocumentConverter()
        # Docling automatically detects and handles these formats:
        # PDF, DOCX, PPTX, XLSX, HTML, WAV, MP3, images (PNG, JPG, etc.)
        
    def extract(self, file_path: str) -> Dict[str, Any]:
        """
        Extract content from any supported file format
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        result = {
            'filename': os.path.basename(file_path),
            'file_type': self._get_file_type(file_path),
            'content': '',
            'metadata': {}
        }
        
        try:
            # Convert document using Docling
            print(f"Processing with Docling...")
            doc_result = self.converter.convert(file_path)
            
            # Extract markdown content
            markdown_content = doc_result.document.export_to_markdown()
            result['content'] = markdown_content
            
            # Extract metadata
            result['metadata'] = {
                'pages': getattr(doc_result.document, 'page_count', 0),
                'tables': self._count_tables(doc_result.document),
                'extraction_method': 'Docling',
                'document_hash': getattr(doc_result.document, 'document_hash', ''),
                'processing_time': getattr(doc_result, 'processing_time_seconds', 0.0)
            }
            
            print(f"Docling processed successfully - {result['metadata']['pages']} pages, {result['metadata']['tables']} tables")
            
        except Exception as e:
            raise Exception(f"Error processing file with Docling: {str(e)}")
        
        return result
    
    def _count_tables(self, document) -> int:
        """Count tables in the document"""
        try:
            # Count table occurrences in markdown content
            markdown_content = document.export_to_markdown()
            # Simple table counting by looking for markdown table patterns
            table_count = markdown_content.count('|')
            if table_count > 0:
                # Rough estimate: tables typically have multiple | per row
                return max(1, table_count // 10)  # Conservative estimate
            return 0
        except:
            return 0
    
    def _get_file_type(self, file_path: str) -> str:
        """Get file type from extension"""
        ext = os.path.splitext(file_path)[1].lower()
        file_types = {
            '.pdf': 'PDF',
            '.docx': 'Word',
            '.doc': 'Word',
            '.pptx': 'PowerPoint', 
            '.ppt': 'PowerPoint',
            '.xlsx': 'Excel',
            '.xls': 'Excel',
            '.html': 'HTML',
            '.htm': 'HTML',
            '.png': 'Image',
            '.jpg': 'Image',
            '.jpeg': 'Image',
            '.tiff': 'Image',
            '.wav': 'Audio',
            '.mp3': 'Audio'
        }
        return file_types.get(ext, 'Unknown')
    
    def to_markdown(self, extracted_data: Dict[str, Any]) -> str:
        """
        Return the extracted markdown content - Docling already provides clean markdown
        """
        return extracted_data['content'].strip()
    
    def get_supported_extensions(self) -> list:
        """Get list of supported file extensions"""
        return [
            '.pdf', '.docx', '.doc', '.pptx', '.ppt', 
            '.xlsx', '.xls', '.html', '.htm',
            '.png', '.jpg', '.jpeg', '.tiff',
            '.wav', '.mp3'
        ]
    
    def is_supported_file(self, file_path: str) -> bool:
        """Check if file format is supported"""
        ext = os.path.splitext(file_path)[1].lower()
        return ext in self.get_supported_extensions() 