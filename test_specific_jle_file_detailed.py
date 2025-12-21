#!/usr/bin/env python3
"""
Test script to check what data is pulled from the specific JLE file: DSO_20243_565.JLE
"""
import os
import re
from config import parse_jle_with_filename, extract_jle_data
from io import BytesIO


def test_specific_jle_file():
    """Test the specific JLE file to see what data gets extracted"""
    jle_file_path = "testfiles/DSO_20243_565.JLE"
    
    print(f"Testing JLE file: {jle_file_path}")
    print("=" * 60)
    
    # Check if file exists
    if not os.path.exists(jle_file_path):
        print(f"ERROR: File {jle_file_path} does not exist!")
        return
    
    print(f"File size: {os.path.getsize(jle_file_path)} bytes")
    
    # Test 1: Using parse_jle_with_filename function
    print("\n1. Testing parse_jle_with_filename()...")
    try:
        df = parse_jle_with_filename(jle_file_path)
        
        print(f"   Records parsed: {len(df)}")
        if not df.empty:
            print(f"   Columns: {list(df.columns)}")
            
            print("\n   Sample of parsed data:")
            print(df.to_string(index=False))
            
            print("\n   Summary statistics:")
            print(f"   - Total records: {len(df)}")
            print(f"   - Unique subject codes: {df['Subject Code'].nunique() if 'Subject Code' in df.columns and not df.empty else 0}")
            print(f"   - Unique lecturers: {df['Lecturer'].nunique() if 'Lecturer' in df.columns and not df.empty else 0}")
            print(f"   - Academic Year: {df['Academic Year'].iloc[0] if 'Academic Year' in df.columns and not df.empty else 'N/A'}")
            print(f"   - Semester: {df['Semester'].iloc[0] if 'Semester' in df.columns and not df.empty else 'N/A'}")
            
            # Show unique subject codes if any
            if 'Subject Code' in df.columns:
                print(f"   - Subject codes found: {sorted(df['Subject Code'].unique())}")
                
        else:
            print("   No records were parsed from the file.")
            
    except Exception as e:
        print(f"   ERROR in parse_jle_with_filename: {e}")
        import traceback
        traceback.print_exc()
    
    # Test 2: Using extract_jle_data function (simulating file upload)
    print("\n2. Testing extract_jle_data()...")
    try:
        # Open the file as a binary stream
        with open(jle_file_path, 'rb') as f:
            file_content = f.read()
        
        # Create a BytesIO object to simulate an uploaded file
        jle_file_obj = BytesIO(file_content)
        jle_file_obj.name = os.path.basename(jle_file_path)
        
        jle_info = extract_jle_data(jle_file_obj)
        
        print(f"   File processed successfully")
        print(f"   - Total courses: {jle_info['total_courses']}")
        print(f"   - Academic Year: {jle_info.get('academic_year', 'Not found')}")
        print(f"   - Semester: {jle_info.get('semester', 'Not found')}")
        
        if jle_info['course_data'] is not None and not jle_info['course_data'].empty:
            print(f"   - Course data shape: {jle_info['course_data'].shape}")
            print(f"   - Sample of course data:")
            print(jle_info['course_data'].to_string(index=False))
        else:
            print("   - No course data was extracted.")
            
        if jle_info['course_codes']:
            print(f"   - Subject codes found: {sorted(set(jle_info['course_codes']))}")
        
    except Exception as e:
        print(f"   ERROR in extract_jle_data: {e}")
        import traceback
        traceback.print_exc()


def analyze_raw_file_structure():
    """Analyze the raw content of the JLE file to understand its structure better"""
    jle_file_path = "testfiles/DSO_20243_565.JLE"
    
    print("\n3. Analyzing raw file structure...")
    
    try:
        with open(jle_file_path, 'rb') as f:
            raw_bytes = f.read()
        
        print(f"   Raw file size: {len(raw_bytes)} bytes")
        
        # Decode with latin1 to see text content (similar to what the parser does)
        decoded_content = raw_bytes.decode('latin1', errors='ignore')
        
        print(f"   Decoded content length: {len(decoded_content)} characters")
        
        # Look for the actual text data part of the file
        # The beginning looks like DBF header, let's find where the actual records start
        # We can look for patterns that indicate subject records
        
        # Search for patterns that match the expected record format
        # Format: "4digits 3-4letters followed by numbers"
        record_start_pattern = re.compile(r"(\d{4})\s+([A-Z][A-Z0-9]{2,10})")
        matches = list(record_start_pattern.finditer(decoded_content))
        
        print(f"   Found {len(matches)} potential record start patterns:")
        for i, match in enumerate(matches):
            # Get context around the match
            start = max(0, match.start() - 20)
            end = min(len(decoded_content), match.end() + 100)
            context = decoded_content[start:end]
            print(f"     {i+1}. Context: {repr(context)}")
        
        # Print a larger portion of the text content focusing on the readable part
        # Skip the initial binary-like headers
        # Find where the actual course data starts (after the DBF header)
        text_start = 0
        for i, char in enumerate(decoded_content):
            if char.isalnum() and ord(char) >= 32:  # Printable character
                text_start = i
                break
        
        print(f"\n   Text portion starting at position {text_start}:")
        text_preview = decoded_content[text_start:text_start+1000]
        print(repr(text_preview))
        
        print(f"\n   More readable text (first 2000 chars after filtering):")
        readable_chars = []
        for char in decoded_content[text_start:text_start+2000]:
            if 32 <= ord(char) <= 126:  # Standard printable ASCII
                readable_chars.append(char)
            elif char in ['\n', '\r', '\t']:  # Common whitespace
                readable_chars.append(char)
            elif ord(char) > 126 and ord(char) < 256:  # Extended ASCII that might be readable
                readable_chars.append('?')  # Replace extended characters with ? for readability
            else:
                readable_chars.append('.')  # Replace control characters
        
        clean_text = ''.join(readable_chars)
        print(clean_text[:1000])  # Show first 1000 characters of cleaned text
        
    except Exception as e:
        print(f"   ERROR analyzing raw content: {e}")
        import traceback
        traceback.print_exc()


def debug_parsing_issue():
    """Debug why the lecturer names and credits aren't being parsed correctly"""
    jle_file_path = "testfiles/DSO_20243_565.JLE"
    
    print("\n4. Debugging parsing issues...")
    
    try:
        with open(jle_file_path, 'rb') as f:
            jle_bytes = f.read()
        
        raw_data = jle_bytes.decode('latin1', errors='ignore')
        
        # Find the record start patterns
        record_start_pattern = re.compile(r"(\d{4})\s+([A-Z][A-Z0-9]{2,10})")
        matches = list(record_start_pattern.finditer(raw_data))
        
        print(f"Found {len(matches)} records to debug:")
        
        for i in range(len(matches)):
            # Determine the start and end of the current record's text chunk
            start_idx = matches[i].start()

            # If there is a next match, end before it starts; otherwise go to EOF
            if i + 1 < len(matches):
                end_idx = matches[i+1].start()
            else:
                end_idx = len(raw_data)

            # Extract the raw chunk for this specific subject
            chunk = raw_data[start_idx:end_idx]

            # Extract Subject Num and Code from the match itself
            subj_num = matches[i].group(1)
            full_subj_code = matches[i].group(2)

            # The first letter of the subject code is actually part of the subject number
            subj_num_with_letter = subj_num + full_subj_code[0]  # 2506 + F = 2506F
            real_subj_code = full_subj_code[1:]  # BACC104 (everything after the first letter)

            # Remove the header (Num + Code) from the chunk to process the rest
            rest_of_text = chunk[len(matches[i].group(0)):].replace('\n', ' ').strip()

            print(f"\nRecord {i+1}:")
            print(f"  Original chunk: {repr(chunk[:200])}...")  # First 200 chars
            print(f"  Subject Num: {subj_num_with_letter}")
            print(f"  Subject Code: {real_subj_code}")
            print(f"  Rest of text: {repr(rest_of_text[:150])}...")  # First 150 chars
            
            # Try to find credit and lecturer
            lecturer_pattern = re.search(r"(\d)\s+([A-Z\.\s]+)$", rest_of_text)
            if lecturer_pattern:
                credit = lecturer_pattern.group(1)
                lecturer = lecturer_pattern.group(2).strip()
                print(f"  Found credit: '{credit}', lecturer: '{lecturer}'")
            else:
                print(f"  Could not find credit/lecturer pattern in: '{rest_of_text[-100:]}'")
                
                # Let's try a different pattern that might match
                # Sometimes there might be special characters or different spacing
                alt_pattern = re.search(r"(\d)\s+([A-Z][A-Z\s\.]+?)(?:\x00|\x01|\x02|\x03|\x04|\x05|\x06|\x07|\x08|\x09|\x0a|\x0b|\x0c|\x0d|\x0e|\x0f|\x10|\x11|\x12|\x13|\x14|\x15|\x16|\x17|\x18|\x19|\x1a|\x1b|\x1c|\x1d|\x1e|\x1f|\x7f|$)", rest_of_text)
                if alt_pattern:
                    credit = alt_pattern.group(1)
                    lecturer = alt_pattern.group(2).strip()
                    print(f"  Alternative pattern found credit: '{credit}', lecturer: '{lecturer}'")
                else:
                    print(f"  Also failed with alternative pattern")
        
    except Exception as e:
        print(f"   ERROR debugging parsing: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    print("Testing specific JLE file: DSO_20243_565.JLE")
    print("=" * 80)
    
    test_specific_jle_file()
    analyze_raw_file_structure()
    debug_parsing_issue()
    
    print("\n" + "=" * 80)
    print("Test completed!")