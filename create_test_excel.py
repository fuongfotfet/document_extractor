"""
Create a test Excel file for testing
"""

import pandas as pd

def create_test_excel():
    """Create a simple test Excel file"""
    
    # Sample data
    data = {
        'Name': ['John Smith', 'Sarah Johnson', 'Mike Chen', 'Lisa Brown'],
        'Department': ['Engineering', 'Marketing', 'Engineering', 'HR'],
        'Salary': [75000, 65000, 80000, 60000],
        'Experience': [5, 3, 7, 4]
    }
    
    df = pd.DataFrame(data)
    
    # Save to Excel
    filename = 'test.xlsx'
    df.to_excel(filename, index=False)
    print(f"Created {filename} with sample data")

if __name__ == "__main__":
    create_test_excel() 