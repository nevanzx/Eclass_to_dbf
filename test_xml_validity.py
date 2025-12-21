import xml.etree.ElementTree as ET
import zipfile
import io
from reports import XMLWordReport

def test_xml_validity():
    """
    Test function to validate XML structure in generated documents
    """
    try:
        # Create a basic XMLWordReport instance with the template
        word_report = XMLWordReport("Report_template.docx")
        
        # Before making any changes, let's check the original document XML
        original_xml = word_report.zip_content.get('word/document.xml', '')
        print("Checking original document XML validity...")
        
        try:
            # Try to parse the original XML
            ET.fromstring(original_xml.encode('utf-8'))
            print("✓ Original document.xml is well-formed")
        except ET.ParseError as e:
            print(f"✗ Original document.xml has XML error: {e}")
            return False
        
        # Now simulate adding some data (empty df to test structure)
        word_report.populate_student_data_table([])
        
        # Check the modified XML
        modified_xml = word_report.zip_content.get('word/document.xml', '')
        print("Checking modified document XML validity...")
        
        try:
            # Try to parse the modified XML
            ET.fromstring(modified_xml.encode('utf-8'))
            print("✓ Modified document.xml is well-formed")
        except ET.ParseError as e:
            print(f"✗ Modified document.xml has XML error: {e}")
            # Print a snippet of the problematic XML
            print("First 1000 characters of modified XML:")
            print(modified_xml[:1000])
            return False
        
        print("✓ XML validation passed")
        return True
        
    except Exception as e:
        print(f"Error during XML validation: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    test_xml_validity()