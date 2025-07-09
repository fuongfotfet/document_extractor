"""
Excel Extractor Module - Handles Excel files (.xlsx, .xls)
"""

import os
import pandas as pd
import zipfile
from typing import Dict, List, Any

from .base_extractor import BaseExtractor


class ExcelExtractor(BaseExtractor):
    """
    Extractor for Excel files (.xlsx, .xls)
    Supports multi-sheet processing and smart table detection
    """
    
    def __init__(self):
        super().__init__()
        self.supported_extensions = ['.xlsx', '.xls']
    
    def _validate_excel_file(self, file_path: str) -> bool:
        """Validate if the file is a proper Excel file"""
        try:
            if file_path.endswith('.xlsx'):
                # Check if it's a valid ZIP file (xlsx is ZIP-based)
                with zipfile.ZipFile(file_path, 'r') as zip_file:
                    # Check for essential Excel components
                    required_files = ['[Content_Types].xml', 'xl/workbook.xml']
                    for required_file in required_files:
                        if required_file not in zip_file.namelist():
                            return False
                return True
            else:
                # For .xls files, try to open with xlrd
                import xlrd
                xlrd.open_workbook(file_path)
                return True
        except Exception:
            return False
    
    def extract(self, file_path: str) -> Dict[str, Any]:
        """
        Extract data from Excel file
        """
        self.validate_file_exists(file_path)
        
        if not self.is_supported_file(file_path):
            raise ValueError(f"File extension not supported. Expected: {self.supported_extensions}")
        
        # Validate file integrity
        if not self._validate_excel_file(file_path):
            raise ValueError(f"File appears to be corrupted or not a valid Excel file. Please check: {os.path.basename(file_path)}")
        
        result = {
            'filename': os.path.basename(file_path),
            'file_type': 'excel',
            'sheets': [],
            'metadata': {}
        }
        
        try:
            # Try method 1: Use ExcelFile with explicit engine
            engine = 'openpyxl' if file_path.endswith('.xlsx') else 'xlrd'
            excel_file = pd.ExcelFile(file_path, engine=engine)
            
            result['metadata']['total_sheets'] = len(excel_file.sheet_names)
            result['metadata']['sheet_names'] = excel_file.sheet_names
            
            # Process each sheet
            for sheet_name in excel_file.sheet_names:
                sheet_data = self._process_sheet(excel_file, sheet_name)
                result['sheets'].append(sheet_data)
                
        except Exception as first_error:
            try:
                # Fallback method: Direct read_excel approach
                print(f"Primary method failed, trying alternative approach...")
                
                # Get sheet names first
                if file_path.endswith('.xlsx'):
                    excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                else:
                    excel_file = pd.ExcelFile(file_path, engine='xlrd')
                
                result['metadata']['total_sheets'] = len(excel_file.sheet_names)
                result['metadata']['sheet_names'] = excel_file.sheet_names
                
                # Process each sheet with direct read_excel
                for sheet_name in excel_file.sheet_names:
                    sheet_data = self._process_sheet_direct(file_path, sheet_name, engine)
                    result['sheets'].append(sheet_data)
                    
            except Exception as second_error:
                # Final fallback: Try to read as CSV if it's really a CSV with .xlsx extension
                try:
                    print("Excel methods failed, attempting CSV fallback...")
                    import csv
                    with open(file_path, 'r', encoding='utf-8') as f:
                        csv.Sniffer().sniff(f.read(1024))  # Test if it's CSV
                        f.seek(0)
                        # If we get here, it might be a CSV file with wrong extension
                        raise ValueError(f"File appears to be a CSV file with .xlsx extension. Please rename to .csv and use appropriate processor.")
                except:
                    pass
                
                raise Exception(f"Unable to process Excel file. Possible issues:\n"
                              f"1. File is corrupted: {os.path.basename(file_path)}\n"
                              f"2. File is password protected\n"
                              f"3. File is not a valid Excel format\n"
                              f"4. File has unusual encoding\n"
                              f"Primary error: {str(first_error)}\n"
                              f"Fallback error: {str(second_error)}")
        
        return result
    
    def _process_sheet(self, excel_file: pd.ExcelFile, sheet_name: str) -> Dict[str, Any]:
        """Process individual sheet"""
        try:
            # Read sheet with explicit parameters
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            
            # Convert to list of lists for processing
            data_list = df.values.tolist()
            
            # Detect if there's a table structure
            tables = self._detect_tables_in_sheet(data_list)
            
            return {
                'sheet_name': sheet_name,
                'shape': df.shape,
                'tables': tables,
                'raw_data': data_list
            }
            
        except Exception as e:
            return {
                'sheet_name': sheet_name,
                'error': f"Error processing sheet: {str(e)}",
                'shape': (0, 0),
                'tables': [],
                'raw_data': []
            }
    
    def _process_sheet_direct(self, file_path: str, sheet_name: str, engine: str) -> Dict[str, Any]:
        """Process individual sheet using direct read_excel approach"""
        try:
            # Read sheet directly from file with proper engine handling
            if engine == 'openpyxl':
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
            
            # Convert to list of lists for processing
            data_list = df.values.tolist()
            
            # Detect if there's a table structure
            tables = self._detect_tables_in_sheet(data_list)
            
            return {
                'sheet_name': sheet_name,
                'shape': df.shape,
                'tables': tables,
                'raw_data': data_list
            }
            
        except Exception as e:
            return {
                'sheet_name': sheet_name,
                'error': f"Error processing sheet: {str(e)}",
                'shape': (0, 0),
                'tables': [],
                'raw_data': []
            }
    
    def _detect_tables_in_sheet(self, data_list: List[List]) -> List[Dict[str, Any]]:
        """Detect tables in sheet data"""
        if not data_list:
            return []
        
        # For now, treat entire sheet as one table if it has data
        non_empty_rows = [row for row in data_list if any(str(cell).strip() for cell in row if cell is not None)]
        
        if not non_empty_rows:
            return []
        
        # Create DataFrame from non-empty data
        df = pd.DataFrame(non_empty_rows)
        
        # Detect header
        has_header = self._detect_header_from_list(non_empty_rows)
        
        return [{
            'table_index': 1,
            'shape': df.shape,
            'has_header': has_header,
            'data': df,
            'raw_data': non_empty_rows
        }]
    
    def to_markdown(self, extracted_data: Dict[str, Any]) -> str:
        """
        Convert extracted data to markdown format - simplified, preserving natural structure
        """
        markdown_content = []
        
        # Process each sheet directly without file info headers
        for sheet in extracted_data['sheets']:
            if 'error' in sheet:
                markdown_content.append(f"**Error in {sheet['sheet_name']}**: {sheet['error']}\n")
                continue
            
            # Only add sheet name if there are multiple sheets
            if len(extracted_data['sheets']) > 1:
                markdown_content.append(f"## {sheet['sheet_name']}\n")
            
            # Add tables directly
            for table in sheet['tables']:
                markdown_table = self.create_markdown_table_from_df(table['data'], table['has_header'])
                markdown_content.append(markdown_table)
                markdown_content.append("")
        
        return "\n".join(markdown_content).strip() 