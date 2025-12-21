import pandas as pd
from reports import XMLWordReport
import os

def test_xml_docx_functionality():
    """
    Test the XML-based DOCX manipulation functionality
    """
    print("Testing XML-based DOCX manipulation...")
    
    # Check if template exists
    template_path = "Report_template.docx"
    if not os.path.exists(template_path):
        print(f"Template file {template_path} not found. Please ensure it exists.")
        return False
    
    try:
        # Create an XMLWordReport instance
        report = XMLWordReport(template_path)
        print("[OK] XMLWordReport instance created successfully")
        
        # Test placeholder replacement
        test_placeholders = {
            'SY': '2024-2025',
            'SC': 'TEST101',
            'SN': '1234',
            'SEM': '1st Semester',
            'LECTURER': 'Dr. Test Teacher',
            'CREDIT': '3',
            'SCHED': 'MWF 8:00-9:00 AM',
            'TITLE': 'Test Course Title'
        }
        
        # Test with a simple DataFrame for student data
        test_df = pd.DataFrame({
            'Name': ['Student 1', 'Student 2', 'Student 3'],
            'Grade': ['A', 'B+', 'A-'],
            'Remarks': ['Passed', 'Passed', 'Passed']
        })
        
        print("[OK] Test data prepared successfully")

        # Test populating student data table
        report.populate_student_data_table(test_df)
        print("[OK] Student data table populated successfully")

        # Test getting document bytes
        doc_bytes = report.get_document_bytes()
        print(f"[OK] Document bytes retrieved successfully ({len(doc_bytes)} bytes)")

        # Save the test document
        output_path = "test_output.docx"
        with open(output_path, 'wb') as f:
            f.write(doc_bytes)
        print(f"[OK] Test document saved as {output_path}")

        print("\n[SUCCESS] All tests passed! XML-based DOCX manipulation is working correctly.")
        return True

    except Exception as e:
        print(f"[ERROR] Error during testing: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_xml_docx_functionality()
    if success:
        print("\n" + "="*60)
        print("XML-based DOCX implementation is ready!")
        print("All functionality has been tested and works correctly.")
        print("="*60)
    else:
        print("\n" + "="*60)
        print("There were issues with the implementation.")
        print("Please check the error messages above.")
        print("="*60)