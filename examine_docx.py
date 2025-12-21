import zipfile
import re

def examine_docx_structure():
    """
    Examine the structure of the DOCX template to understand the table format
    """
    template_path = "Report_template.docx"
    
    if not template_path:
        print("No template found")
        return
        
    with zipfile.ZipFile(template_path, 'r') as docx_zip:
        # List all files in the DOCX
        file_list = docx_zip.namelist()
        print("Files in DOCX:")
        for file in file_list:
            print(f"  - {file}")
        
        # Read the main document
        if 'word/document.xml' in file_list:
            doc_xml = docx_zip.read('word/document.xml').decode('utf-8')
            print(f"\nDocument XML length: {len(doc_xml)}")
            
            # Look for any table-related content
            print("\nLooking for table-related tags:")
            table_tags = re.findall(r'<w:tbl.*?>|</w:tbl>|<w:tr.*?>|</w:tr>|<w:tc.*?>|</w:tc>', doc_xml)
            print(f"Found {len(table_tags)} table-related tags:")
            for i, tag in enumerate(table_tags[:20]):  # Show first 20
                print(f"  {i}: {tag}")
            
            # Look for content that might be in tables
            print("\nLooking for text content that might be in tables:")
            text_elements = re.findall(r'<w:t>(.*?)</w:t>', doc_xml)
            print(f"Found {len(text_elements)} text elements:")
            for i, text in enumerate(text_elements):
                print(f"  {i}: '{text}'")
            
            # Check for any existing tables with a broader pattern
            print("\nTrying broader table search...")
            # Look for content between table tags, including nested
            start_pos = 0
            table_count = 0
            while start_pos < len(doc_xml):
                tbl_start = doc_xml.find('<w:tbl', start_pos)
                if tbl_start == -1:
                    break
                
                # Look for the corresponding end tag
                content_start = doc_xml.find('>', tbl_start) + 1
                nested_level = 1
                search_pos = content_start
                
                while nested_level > 0 and search_pos < len(doc_xml):
                    next_open = doc_xml.find('<w:tbl', search_pos)
                    next_close = doc_xml.find('</w:tbl>', search_pos)
                    
                    if next_open != -1 and next_open < next_close:
                        nested_level += 1
                        search_pos = next_open + 1
                    elif next_close != -1:
                        nested_level -= 1
                        if nested_level == 0:
                            table_content = doc_xml[tbl_start:next_close+8]  # +8 for length of '</w:tbl>'
                            table_count += 1
                            print(f"\nTable {table_count}:")
                            print(f"  Length: {len(table_content)} chars")
                            # Extract text from this table
                            table_texts = re.findall(r'<w:t>(.*?)</w:t>', table_content)
                            print(f"  Texts: {table_texts[:10]}")  # First 10 texts
                        search_pos = next_close + 1
                    else:
                        break  # No more closing tags found
                
                start_pos = next_close + 1 if next_close != -1 else len(doc_xml)

if __name__ == "__main__":
    examine_docx_structure()