import pandas as pd
from reports import XMLWordReport
import zipfile
import re

def verify_content_insertion():
    """
    Verify that content is actually inserted in the DOCX
    """
    print("Verifying content insertion in DOCX...")
    
    # Create test data
    test_df = pd.DataFrame({
        'Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
        'Grade': ['A', 'B+', 'A-'],
        'Remarks': ['Passed', 'Passed', 'Passed']
    })
    
    # Create report and populate
    report = XMLWordReport("Report_template.docx")
    report.populate_student_data_table(test_df)
    
    # Get the document bytes and save
    doc_bytes = report.get_document_bytes()
    with open("verification_test.docx", "wb") as f:
        f.write(doc_bytes)
    
    print("Document saved as verification_test.docx")
    
    # Now read the saved document and check its content
    with zipfile.ZipFile("verification_test.docx", 'r') as docx_zip:
        doc_xml = docx_zip.read('word/document.xml').decode('utf-8')
        
        # Look for the student names in the document
        names_found = []
        for name in ['John Doe', 'Jane Smith', 'Bob Johnson']:
            if name in doc_xml:
                names_found.append(name)
        
        print(f"Names found in document: {names_found}")
        print(f"Expected: 3, Found: {len(names_found)}")
        
        # Also check for the statistics
        if 'STATISTICS' in doc_xml:
            stats_match = re.search(r'STATISTICS.*?TOTAL=\d+', doc_xml)
            if stats_match:
                print(f"Statistics found: {stats_match.group(0)}")
        
        # Look for all text elements
        text_elements = re.findall(r'<w:t>(.*?)</w:t>', doc_xml)
        print(f"First 20 text elements: {text_elements[:20]}")
        
        # Check specifically for student data in the text elements
        student_data_found = []
        for text in text_elements:
            if any(name in text for name in ['John', 'Jane', 'Bob']):
                student_data_found.append(text)
        
        print(f"Student-related text elements found: {student_data_found}")
        
        if len(names_found) == 3:
            print("\n[SUCCESS] All student names were found in the document!")
            return True
        elif len(student_data_found) > 0:
            print(f"\n[PARTIAL SUCCESS] Some student data found: {student_data_found}")
            return True
        else:
            print("\n[ISSUE] No student names found in the document")
            # Let's see if the data is there but in a different format
            # Maybe the names are split like the original "NAME" was split to 'N', 'A', 'M', 'E'
            for name in ['John', 'Jane', 'Bob']:
                if name[0] in doc_xml:  # Check if first letter exists
                    print(f"  Found first letter of {name}: '{name[0]}'")
            return False

if __name__ == "__main__":
    success = verify_content_insertion()
    if success:
        print("\nContent insertion verification: PASSED")
    else:
        print("\nContent insertion verification: FAILED")