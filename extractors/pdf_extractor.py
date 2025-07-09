"""
PDF Extractor Module - Handles PDF files (.pdf)
"""

import os
import io
import pandas as pd
from typing import Dict, List, Any

from .base_extractor import BaseExtractor

# Import for PDF processing
try:
    import PyPDF2
    import fitz  # PyMuPDF
except ImportError:
    print("Warning: PDF libraries not installed. PDF support disabled.")
    PyPDF2 = None
    fitz = None

# Import for GMFT (advanced table extraction)
try:
    from gmft.core.auto_lazy import AutoTableDetector, AutoTableFormatter
    from gmft.pdf_bindings.pdfium import PyPDFium2Document
    print("GMFT loaded successfully")
except ImportError:
    print("Warning: GMFT not installed. Advanced table extraction disabled.")
    AutoTableDetector = None
    AutoTableFormatter = None
    PyPDFium2Document = None

# Import for OCR (fallback)
try:
    import pytesseract
    from PIL import Image
    import cv2
    import numpy as np
except ImportError:
    print("Warning: OCR libraries not installed. Scanned PDF support disabled.")
    pytesseract = None
    Image = None
    cv2 = None


class PdfExtractor(BaseExtractor):
    """
    Extractor for PDF files (.pdf) - supports both text-based and scanned PDFs
    Uses GMFT for advanced table extraction and OCR as fallback
    """
    
    def __init__(self):
        super().__init__()
        self.supported_extensions = ['.pdf']
    
    def extract(self, file_path: str) -> Dict[str, Any]:
        """
        Extract text and tables from PDF file
        """
        self.validate_file_exists(file_path)
        
        if not self.is_supported_file(file_path):
            raise ValueError(f"File extension not supported. Expected: {self.supported_extensions}")
        
        result = {
            'filename': os.path.basename(file_path),
            'file_type': 'pdf',
            'text_content': '',
            'tables': [],
            'metadata': {}
        }
        
        try:
            # Extract text using PyMuPDF
            if fitz is not None:
                result = self._extract_text_with_pymupdf(file_path, result)
            elif PyPDF2 is not None:
                result = self._extract_pdf_with_pypdf2(file_path, result)
            
            # Extract tables using GMFT (much better than OCR)
            if AutoTableDetector is not None:
                print("Extracting tables using GMFT (advanced ML model)...")
                gmft_tables = self._extract_tables_with_gmft(file_path)
                result['tables'].extend(gmft_tables)
                result['metadata']['total_tables'] = len(gmft_tables)
                result['metadata']['extraction_method'] = 'GMFT'
            else:
                # Fallback to OCR if GMFT not available
                if not result['text_content'].strip() and pytesseract is not None:
                    print("No text detected, trying OCR extraction...")
                    result = self._extract_pdf_with_ocr(file_path, result)
                else:
                    print("Warning: No advanced table extraction available. Install GMFT for better results.")
                
        except Exception as e:
            raise Exception(f"Error reading PDF file: {str(e)}")
        
        return result
    
    def _extract_text_with_pymupdf(self, file_path: str, result: Dict[str, Any]) -> Dict[str, Any]:
        """Extract text content using PyMuPDF"""
        if fitz is None:
            raise ImportError("PyMuPDF not available")
        
        doc = fitz.open(file_path)  # type: ignore
        all_text = []
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)  # type: ignore
            text = page.get_text()  # type: ignore
            if text.strip():
                all_text.append(f"=== Page {page_num + 1} ===\n{text}")
        
        result['text_content'] = '\n\n'.join(all_text)
        result['metadata']['total_pages'] = len(doc)
        
        doc.close()  # type: ignore
        return result
    
    def _extract_tables_with_gmft(self, file_path: str) -> List[Dict[str, Any]]:
        """Extract tables using GMFT (Give Me Formatted Tables)"""
        if AutoTableDetector is None or AutoTableFormatter is None or PyPDFium2Document is None:
            return []
        
        try:
            # Initialize GMFT components
            detector = AutoTableDetector()
            formatter = AutoTableFormatter()
            
            # Open PDF document
            doc = PyPDFium2Document(file_path)  # type: ignore
            
            tables = []
            table_count = 0
            
            # Process each page
            for page_num, page in enumerate(doc):
                try:
                    # Extract tables from page
                    page_tables = detector.extract(page)
                    
                    for table_idx, cropped_table in enumerate(page_tables):
                        table_count += 1
                        
                        try:
                            # Format table to DataFrame
                            formatted_table = formatter.extract(cropped_table)
                            df = formatted_table.df()
                            
                            # Convert to our standard format
                            table_data = df.values.tolist()
                            
                            # Detect header
                            has_header = self._detect_header_from_list(table_data) if table_data else False
                            
                            processed_table = {
                                'page': page_num + 1,
                                'table_index': table_count,
                                'shape': df.shape,
                                'has_header': has_header,
                                'data': df,
                                'raw_data': table_data,
                                'extraction_method': 'GMFT',
                                'confidence': cropped_table.confidence_score if hasattr(cropped_table, 'confidence_score') else 0.9
                            }
                            
                            tables.append(processed_table)
                            
                        except Exception as e:
                            print(f"Failed to process table {table_count} on page {page_num + 1}: {e}")
                            continue
                            
                except Exception as e:
                    print(f"Failed to extract tables from page {page_num + 1}: {e}")
                    continue
            
            # Close document
            doc.close()  # type: ignore
            
            print(f"GMFT extracted {len(tables)} tables successfully")
            return tables
            
        except Exception as e:
            print(f"GMFT extraction failed: {e}")
            return []
    
    def _extract_pdf_with_ocr(self, file_path: str, result: Dict[str, Any]) -> Dict[str, Any]:
        """Extract PDF using OCR (fallback for scanned PDFs)"""
        if pytesseract is None or fitz is None or cv2 is None:
            raise ImportError("OCR libraries not available. Install with: pip install pytesseract pillow opencv-python")
        
        doc = fitz.open(file_path)  # type: ignore
        all_text = []
        all_tables = []
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)  # type: ignore
            
            # Convert page to image
            mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for better OCR quality  # type: ignore
            pix = page.get_pixmap(matrix=mat)  # type: ignore
            img_data = pix.tobytes("png")  # type: ignore
            
            # Convert to PIL Image
            img = Image.open(io.BytesIO(img_data))  # type: ignore
            
            # OCR text extraction
            try:
                text = pytesseract.image_to_string(img, lang='vie+eng')  # Vietnamese + English  # type: ignore
                if text.strip():
                    all_text.append(f"=== Page {page_num + 1} (OCR) ===\n{text}")
            except Exception as e:
                print(f"OCR failed for page {page_num + 1}: {e}")
            
            # Try OCR table detection (basic)
            try:
                tables = self._extract_tables_with_ocr(img, page_num + 1)
                all_tables.extend(tables)
            except Exception as e:
                print(f"OCR table detection failed for page {page_num + 1}: {e}")
        
        result['text_content'] = '\n\n'.join(all_text)
        result['tables'] = all_tables
        result['metadata']['total_pages'] = len(doc)
        result['metadata']['total_tables'] = len(all_tables)
        result['metadata']['extraction_method'] = 'OCR'
        
        doc.close()  # type: ignore
        return result
    
    def _extract_tables_with_ocr(self, img: Any, page_num: int) -> List[Dict[str, Any]]:
        """Extract tables using OCR and image processing (fallback method)"""
        if cv2 is None or pytesseract is None:
            return []
        
        # Convert PIL to OpenCV
        img_array = np.array(img)  # type: ignore
        if len(img_array.shape) == 3:
            img_gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)  # type: ignore
        else:
            img_gray = img_array
        
        tables = []
        
        try:
            # Simple table detection using lines
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))  # type: ignore
            horizontal_lines = cv2.morphologyEx(img_gray, cv2.MORPH_OPEN, horizontal_kernel)  # type: ignore
            
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))  # type: ignore
            vertical_lines = cv2.morphologyEx(img_gray, cv2.MORPH_OPEN, vertical_kernel)  # type: ignore
            
            table_mask = cv2.addWeighted(horizontal_lines, 0.5, vertical_lines, 0.5, 0.0)  # type: ignore
            contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)  # type: ignore
            
            table_count = 0
            for contour in contours:
                area = cv2.contourArea(contour)  # type: ignore
                if area > 5000:  # Minimum table area
                    table_count += 1
                    x, y, w, h = cv2.boundingRect(contour)  # type: ignore
                    table_img = img.crop((x, y, x + w, y + h))  # type: ignore
                    
                    try:
                        table_text = pytesseract.image_to_string(table_img, lang='vie+eng')  # type: ignore
                        lines = [line.strip() for line in table_text.split('\n') if line.strip()]
                        
                        if len(lines) >= 2:
                            table_data = []
                            for line in lines:
                                cells = [cell.strip() for cell in line.split() if cell.strip()]
                                if cells:
                                    table_data.append(cells)
                            
                            if table_data:
                                processed_table = self._process_ocr_table_data(table_data, page_num, table_count)
                                tables.append(processed_table)
                                
                    except Exception as e:
                        print(f"OCR table parsing failed: {e}")
                        
        except Exception as e:
            print(f"Table detection failed: {e}")
        
        return tables
    
    def _process_ocr_table_data(self, table_data: List[List[str]], page_num: int, table_num: int) -> Dict[str, Any]:
        """Process OCR table data"""
        if not table_data:
            return {
                'page': page_num,
                'table_index': table_num,
                'shape': (0, 0),
                'has_header': False,
                'data': pd.DataFrame(),
                'raw_data': [],
                'extraction_method': 'OCR'
            }
        
        # Normalize row lengths
        max_cols = max(len(row) for row in table_data)
        normalized_data = []
        for row in table_data:
            normalized_row = row + [''] * (max_cols - len(row))
            normalized_data.append(normalized_row)
        
        df = pd.DataFrame(normalized_data)
        has_header = self._detect_header_from_list(normalized_data)
        
        return {
            'page': page_num,
            'table_index': table_num,
            'shape': (len(normalized_data), max_cols),
            'has_header': has_header,
            'data': df,
            'raw_data': normalized_data,
            'extraction_method': 'OCR'
        }
    
    def _extract_pdf_with_pypdf2(self, file_path: str, result: Dict[str, Any]) -> Dict[str, Any]:
        """Extract PDF using PyPDF2 (basic text extraction)"""
        if PyPDF2 is None:
            raise ImportError("PyPDF2 not available")
        
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)  # type: ignore
            
            all_text = []
            for page_num, page in enumerate(reader.pages):
                text = page.extract_text()  # type: ignore
                if text.strip():
                    all_text.append(f"=== Page {page_num + 1} ===\n{text}")
            
            result['text_content'] = '\n\n'.join(all_text)
            result['metadata']['total_pages'] = len(reader.pages)
            result['metadata']['total_tables'] = 0  # PyPDF2 does not support table extraction
        
        return result
    
    def _process_pdf_table(self, table_data: List[List], page_num: int, table_num: int) -> Dict[str, Any]:
        """Process table data extracted from PDF"""
        if not table_data:
            return {
                'page': page_num,
                'table_index': table_num,
                'shape': (0, 0),
                'has_header': False,
                'data': pd.DataFrame(),
                'raw_data': []
            }
        
        # Clean table data
        cleaned_data = []
        for row in table_data:
            cleaned_row = [str(cell).strip() if cell is not None else '' for cell in row]
            if any(cleaned_row):  # Only keep rows with at least 1 non-empty cell
                cleaned_data.append(cleaned_row)
        
        if cleaned_data:
            df = pd.DataFrame(cleaned_data)
            has_header = self._detect_header_from_list(cleaned_data)
            
            return {
                'page': page_num,
                'table_index': table_num,
                'shape': (len(cleaned_data), len(cleaned_data[0]) if cleaned_data else 0),
                'has_header': has_header,
                'data': df,
                'raw_data': cleaned_data
            }
        
        return {
            'page': page_num,
            'table_index': table_num,
            'shape': (0, 0),
            'has_header': False,
            'data': pd.DataFrame(),
            'raw_data': []
        }
    
    def to_markdown(self, extracted_data: Dict[str, Any]) -> str:
        """
        Convert extracted data to markdown format - simplified, preserving natural structure
        """
        markdown_content = []
        
        # Add text content if it exists, preserving original structure
        if extracted_data['text_content'].strip():
            markdown_content.append(extracted_data['text_content'])
            
            # Add spacing if there are tables after text
            if extracted_data['tables']:
                markdown_content.append("")
        
        # Add tables directly after text
        if extracted_data['tables']:
            for table in extracted_data['tables']:
                markdown_table = self.create_markdown_table_from_df(table['data'], table['has_header'])
                markdown_content.append(markdown_table)
                markdown_content.append("")
        
        return "\n".join(markdown_content).strip() 