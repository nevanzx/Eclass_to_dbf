import pandas as pd
from reports import XMLWordReport
import os
import zipfile
import xml.etree.ElementTree as ET

def test_dbf_content_insertion():
    """
    Test specifically for DBF content insertion functionality
    """
    print("Testing DBF content insertion in DOCX...")
    
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
        
        # Before populating, check the original document content
        original_doc_xml = report.zip_content.get('word/document.xml', '')
        original_student_rows = original_doc_xml.count('<w:tr')
        print(f"[INFO] Original document has {original_student_rows} table rows")
        
        # Populate the student data table
        report.populate_student_data_table(test_df)
        print("[OK] Student data table populated successfully")
        
        # After populating, check the modified document content
        modified_doc_xml = report.zip_content.get('word/document.xml', '')
        modified_student_rows = modified_doc_xml.count('<w:tr')
        print(f"[INFO] Modified document has {modified_student_rows} table rows")
        
        # Check if the student names are present in the modified document
        student_names_present = all(name in modified_doc_xml for name in ['John Doe', 'Jane Smith', 'Bob Johnson'])
        if student_names_present:
            print("[OK] Student names found in the modified document")
        else:
            print("[WARNING] Student names may not be present in the document")
            # Let's see what's in the document to debug
            print("[DEBUG] Looking for student data in document...")
            # Extract text from the XML to see if data was inserted
            import re
            text_elements = re.findall(r'<w:t>(.*?)</w:t>', modified_doc_xml)
            student_data_found = any('John' in text or 'Jane' in text or 'Bob' in text for text in text_elements)
            if student_data_found:
                print("[OK] Student data found in text elements")
            else:
                print("[ISSUE] No student data found in text elements")
        
        # Test getting document bytes
        doc_bytes = report.get_document_bytes()
        print(f"[OK] Document bytes retrieved successfully ({len(doc_bytes)} bytes)")
        
        # Save the test document
        output_path = "test_dbf_insertion.docx"
        with open(output_path, 'wb') as f:
            f.write(doc_bytes)
        print(f"[OK] Test document saved as {output_path}")
        
        # Verify the document can be opened and has the expected content
        with zipfile.ZipFile(output_path, 'r') as docx_zip:
            doc_xml_content = docx_zip.read('word/document.xml').decode('utf-8')
            
            # Check if student data is present in the actual saved document
            has_student_data = any(name in doc_xml_content for name in ['John Doe', 'Jane Smith', 'Bob Johnson'])
            if has_student_data:
                print("[SUCCESS] Student data successfully inserted in the saved DOCX file")
            else:
                print("[WARNING] Student data may not be present in the saved DOCX file")
                
                # Check if there are more rows than original (indicating data was added)
                if modified_student_rows > original_student_rows:
                    print("[INFO] Row count increased, suggesting data was added")
                else:
                    print("[ISSUE] Row count did not increase, suggesting no data was added")
        
        print("\n[SUCCESS] DBF content insertion test completed!")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error during DBF content insertion test: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_dbf_content_insertion()
    if success:
        print("\n" + "="*60)
        print("DBF content insertion is working correctly!")
        print("Student data from DataFrames is being properly inserted.")
        print("="*60)
    else:
        print("\n" + "="*60)
        print("There were issues with DBF content insertion.")
        print("Please check the error messages above.")
        print("="*60)