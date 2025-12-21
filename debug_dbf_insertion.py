import pandas as pd
from reports import XMLWordReport
import os
import zipfile
import xml.etree.ElementTree as ET
import re

def debug_dbf_content_insertion():
    """
    Debug DBF content insertion functionality
    """
    print("Debugging DBF content insertion in DOCX...")
    
    # Check if template exists
    template_path = "Report_template.docx"
    if not os.path.exists(template_path):
        print(f"Template file {template_path} not found. Please ensure it exists.")
        return False
    
    try:
        # Create an XMLWordReport instance
        report = XMLWordReport(template_path)
        print("[OK] XMLWordReport instance created successfully")
        
        # Create a more realistic test DataFrame similar to DBF data
        test_df = pd.DataFrame({
            'Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
            'Grade': ['A', 'B+', 'A-'],
            'Remarks': ['Passed', 'Passed', 'Passed']
        })
        
        print("[OK] Test DataFrame created with student data")
        print(f"DataFrame columns: {list(test_df.columns)}")
        print(f"DataFrame content:\n{test_df}")
        
        # Before populating, check the original document content
        original_doc_xml = report.zip_content.get('word/document.xml', '')
        print(f"[INFO] Original document XML length: {len(original_doc_xml)}")
        
        # Find and analyze tables in the original document
        table_pattern = r'(<w:tbl[\s>][^<]*>.*?</w:tbl>)'
        all_tables = re.findall(table_pattern, original_doc_xml, re.DOTALL)
        print(f"[INFO] Found {len(all_tables)} tables in the document")
        
        for i, table in enumerate(all_tables):
            print(f"Table {i}:")
            row_count = len(re.findall(r'<w:tr', table))
            print(f"  - Row count: {row_count}")
            # Extract row content for analysis
            rows = re.findall(r'(<w:tr>.*?</w:tr>)', table, re.DOTALL)
            for j, row in enumerate(rows):
                text_elements = re.findall(r'<w:t>(.*?)</w:t>', row)
                print(f"    Row {j}: {text_elements}")
        
        # Check for student-related indicators in tables
        for i, table in enumerate(all_tables):
            table_lower = table.lower()
            student_indicators = [
                'name', 'student', 'grade', 'remark', 'no.', 'number', 'no ', 'id',
                'lastname', 'first name', 'score', 'result', 'status', 'comment'
            ]
            indicators_found = [ind for ind in student_indicators if ind in table_lower]
            print(f"Table {i} - Indicators found: {indicators_found}")
            
            # Check for placeholder patterns
            placeholders_found = re.findall(r'\[Insert.*?\]', table)
            if placeholders_found:
                print(f"Table {i} - Placeholders found: {placeholders_found}")
        
        # Call the populate method
        report.populate_student_data_table(test_df)
        print("[OK] Student data table populated (method completed)")
        
        # After populating, check the modified document content
        modified_doc_xml = report.zip_content.get('word/document.xml', '')
        print(f"[INFO] Modified document XML length: {len(modified_doc_xml)}")
        
        # Find and analyze tables in the modified document
        modified_tables = re.findall(table_pattern, modified_doc_xml, re.DOTALL)
        print(f"[INFO] Found {len(modified_tables)} tables in the modified document")
        
        for i, table in enumerate(modified_tables):
            print(f"Modified Table {i}:")
            row_count = len(re.findall(r'<w:tr', table))
            print(f"  - Row count: {row_count}")
            # Extract row content for analysis
            rows = re.findall(r'(<w:tr>.*?</w:tr>)', table, re.DOTALL)
            for j, row in enumerate(rows):
                text_elements = re.findall(r'<w:t>(.*?)</w:t>', row)
                print(f"    Row {j}: {text_elements}")
        
        # Check if student names are present in the modified document
        text_elements = re.findall(r'<w:t>(.*?)</w:t>', modified_doc_xml)
        student_names_present = any('John' in text or 'Jane' in text or 'Bob' in text for text in text_elements)
        
        if student_names_present:
            print("[SUCCESS] Student names found in the modified document!")
        else:
            print("[ISSUE] Student names NOT found in the modified document")
            print(f"Text elements found: {text_elements[:10]}...")  # Show first 10 elements
        
        # Test getting document bytes
        doc_bytes = report.get_document_bytes()
        print(f"[OK] Document bytes retrieved successfully ({len(doc_bytes)} bytes)")
        
        # Save the debug document
        output_path = "debug_dbf_insertion.docx"
        with open(output_path, 'wb') as f:
            f.write(doc_bytes)
        print(f"[OK] Debug document saved as {output_path}")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error during DBF content insertion debug: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = debug_dbf_content_insertion()
    if success:
        print("\n" + "="*60)
        print("Debugging completed. Check the results above.")
        print("="*60)
    else:
        print("\n" + "="*60)
        print("There were issues during debugging.")
        print("Please check the error messages above.")
        print("="*60)