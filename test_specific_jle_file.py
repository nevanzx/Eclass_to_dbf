#!/usr/bin/env python3
"""
Test script to check what data is pulled from the specific JLE file: DSO_20243_565.JLE
"""
import os
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
            print(f"   - Unique subject codes: {df['Subject Code'].nunique() if 'Subject Code' in df.columns else 0}")
            print(f"   - Unique lecturers: {df['Lecturer'].nunique() if 'Lecturer' in df.columns else 0}")
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
            print(f"   - Subject codes found: {sorted(jle_info['course_codes'])}")
        
    except Exception as e:
        print(f"   ERROR in extract_jle_data: {e}")
        import traceback
        traceback.print_exc()


def analyze_raw_file_content():
    """Analyze the raw content of the JLE file to understand its structure"""
    jle_file_path = "testfiles/DSO_20243_565.JLE"
    
    print("\n3. Analyzing raw file content...")
    
    try:
        with open(jle_file_path, 'rb') as f:
            raw_bytes = f.read()
        
        print(f"   Raw file size: {len(raw_bytes)} bytes")
        
        # Decode with latin1 to see text content (similar to what the parser does)
        decoded_content = raw_bytes.decode('latin1', errors='ignore')
        
        print(f"   Decoded content length: {len(decoded_content)} characters")
        
        # Show first 1000 characters to get an idea of the structure
        print(f"\n   First 1000 characters of decoded content:")
        print(repr(decoded_content[:1000]))
        
        print(f"\n   First 500 characters (as readable text):")
        preview = decoded_content[:500].replace('\n', '\\n').replace('\r', '\\r')
        print(preview)
        
        # Show last 500 characters too
        print(f"\n   Last 500 characters (as readable text):")
        preview_end = decoded_content[-500:].replace('\n', '\\n').replace('\r', '\\r')
        print(preview_end)
        
    except Exception as e:
        print(f"   ERROR analyzing raw content: {e}")


if __name__ == "__main__":
    print("Testing specific JLE file: DSO_20243_565.JLE")
    print("=" * 80)
    
    test_specific_jle_file()
    analyze_raw_file_content()
    
    print("\n" + "=" * 80)
    print("Test completed!")