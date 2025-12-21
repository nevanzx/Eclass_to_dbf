import pandas as pd
from reports import XMLWordReport
import zipfile
import re

def debug_detailed():
    """
    Detailed debugging of the table insertion process
    """
    print("Detailed debugging of table insertion...")
    
    # Create test data
    test_df = pd.DataFrame({
        'Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
        'Grade': ['A', 'B+', 'A-'],
        'Remarks': ['Passed', 'Passed', 'Passed']
    })
    
    # Create report
    report = XMLWordReport("Report_template.docx")
    
    # Check original document
    original_xml = report.zip_content['word/document.xml']
    print(f"Original XML length: {len(original_xml)}")
    
    # Find tables in original
    all_tables = []
    start_pos = 0
    while start_pos < len(original_xml):
        tbl_start = original_xml.find('<w:tbl', start_pos)
        if tbl_start == -1:
            break
        
        content_start = original_xml.find('>', tbl_start) + 1
        nested_level = 1
        search_pos = content_start
        
        while nested_level > 0 and search_pos < len(original_xml):
            next_open = original_xml.find('<w:tbl', search_pos)
            next_close = original_xml.find('</w:tbl>', search_pos)
            
            if next_open != -1 and next_close != -1 and next_open < next_close:
                nested_level += 1
                search_pos = next_open + 1
            elif next_close != -1:
                nested_level -= 1
                if nested_level == 0:
                    table_content = original_xml[tbl_start:next_close+8]
                    all_tables.append(table_content)
                    start_pos = next_close + 1
                    break
                else:
                    search_pos = next_close + 1
            else:
                next_possible_close = original_xml.find('</w:tbl>', search_pos)
                if next_possible_close != -1:
                    table_content = original_xml[tbl_start:next_possible_close+8]
                    all_tables.append(table_content)
                    start_pos = next_possible_close + 1
                    break
                else:
                    start_pos = tbl_start + 1
                    break
    
    print(f"Found {len(all_tables)} tables in original document")
    
    for i, table in enumerate(all_tables):
        print(f"\nTable {i}:")
        rows = re.findall(r'(<w:tr[^>]*>.*?</w:tr>)', table, re.DOTALL)
        print(f"  Rows: {len(rows)}")
        for j, row in enumerate(rows):
            text_elements = re.findall(r'<w:t>(.*?)</w:t>', row)
            print(f"    Row {j}: {text_elements}")
    
    # Now run the populate method
    print("\nCalling populate_student_data_table...")
    report.populate_student_data_table(test_df)
    
    # Check modified document
    modified_xml = report.zip_content['word/document.xml']
    print(f"Modified XML length: {len(modified_xml)}")
    
    # Find tables in modified
    mod_tables = []
    start_pos = 0
    while start_pos < len(modified_xml):
        tbl_start = modified_xml.find('<w:tbl', start_pos)
        if tbl_start == -1:
            break
        
        content_start = modified_xml.find('>', tbl_start) + 1
        nested_level = 1
        search_pos = content_start
        
        while nested_level > 0 and search_pos < len(modified_xml):
            next_open = modified_xml.find('<w:tbl', search_pos)
            next_close = modified_xml.find('</w:tbl>', search_pos)
            
            if next_open != -1 and next_close != -1 and next_open < next_close:
                nested_level += 1
                search_pos = next_open + 1
            elif next_close != -1:
                nested_level -= 1
                if nested_level == 0:
                    table_content = modified_xml[tbl_start:next_close+8]
                    mod_tables.append(table_content)
                    start_pos = next_close + 1
                    break
                else:
                    search_pos = next_close + 1
            else:
                next_possible_close = modified_xml.find('</w:tbl>', search_pos)
                if next_possible_close != -1:
                    table_content = modified_xml[tbl_start:next_possible_close+8]
                    mod_tables.append(table_content)
                    start_pos = next_possible_close + 1
                    break
                else:
                    start_pos = tbl_start + 1
                    break
    
    print(f"Found {len(mod_tables)} tables in modified document")
    
    for i, table in enumerate(mod_tables):
        print(f"\nModified Table {i}:")
        rows = re.findall(r'(<w:tr[^>]*>.*?</w:tr>)', table, re.DOTALL)
        print(f"  Rows: {len(rows)}")
        for j, row in enumerate(rows):
            text_elements = re.findall(r'<w:t>(.*?)</w:t>', row)
            print(f"    Row {j}: {text_elements}")
    
    # Save and check final document
    doc_bytes = report.get_document_bytes()
    with open("detailed_debug.docx", "wb") as f:
        f.write(doc_bytes)
    
    print("\nDetailed debug document saved as detailed_debug.docx")

if __name__ == "__main__":
    debug_detailed()