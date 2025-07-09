"""
Base Extractor Module - Abstract base class for all extractors
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Any
import os


class BaseExtractor(ABC):
    """
    Abstract base class for all file extractors
    Defines common interface and shared functionality
    """
    
    def __init__(self):
        self.supported_extensions = []
    
    @abstractmethod
    def extract(self, file_path: str) -> Dict[str, Any]:
        """
        Extract data from file
        
        Args:
            file_path: Path to the file to extract
            
        Returns:
            Dictionary containing extracted data
        """
        pass
    
    @abstractmethod
    def to_markdown(self, extracted_data: Dict[str, Any]) -> str:
        """
        Convert extracted data to markdown format
        
        Args:
            extracted_data: Data extracted from extract() method
            
        Returns:
            Markdown formatted string
        """
        pass
    
    def validate_file_exists(self, file_path: str):
        """Validate that file exists"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
    
    def is_supported_file(self, file_path: str) -> bool:
        """Check if file type is supported"""
        _, ext = os.path.splitext(file_path.lower())
        return ext in self.supported_extensions
    
    def create_markdown_table_from_df(self, df, has_header: bool = True) -> str:
        """
        Create markdown table from pandas DataFrame
        
        Args:
            df: pandas DataFrame
            has_header: Whether to treat first row as header
            
        Returns:
            Markdown formatted table string
        """
        if df.empty:
            return "*(Empty table)*\n"
        
        markdown_lines = []
        
        # Get data as list of lists
        data = df.values.tolist()
        
        if has_header and len(data) > 0:
            # Use first row as header
            headers = [str(cell).replace('|', '\\|') for cell in data[0]]
            data_rows = data[1:]
        else:
            # Generate column headers
            headers = [f"Column {i+1}" for i in range(len(data[0]) if data else 0)]
            data_rows = data
        
        if headers:
            # Header row
            markdown_lines.append("| " + " | ".join(headers) + " |")
            # Separator row
            markdown_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
            
            # Data rows
            for row in data_rows:
                escaped_row = [str(cell).replace('|', '\\|') for cell in row]
                # Ensure row has same length as headers
                while len(escaped_row) < len(headers):
                    escaped_row.append('')
                markdown_lines.append("| " + " | ".join(escaped_row[:len(headers)]) + " |")
        
        return "\n".join(markdown_lines) + "\n"
    
    def _detect_header_from_list(self, data_list: List[List]) -> bool:
        """
        Smart header detection algorithm
        Checks if first row contains mostly text (potential headers)
        """
        if not data_list or len(data_list) < 2:
            return False
        
        first_row = data_list[0]
        
        # Count text vs numeric cells in first row
        text_count = 0
        total_count = len(first_row)
        
        for cell in first_row:
            cell_str = str(cell).strip()
            if cell_str and not self._is_numeric(cell_str):
                text_count += 1
        
        # If more than 50% are text, likely header
        return (text_count / total_count) > 0.5 if total_count > 0 else False
    
    def _is_numeric(self, value: str) -> bool:
        """Check if string represents a number"""
        try:
            float(value.replace(',', ''))
            return True
        except (ValueError, AttributeError):
            return False 