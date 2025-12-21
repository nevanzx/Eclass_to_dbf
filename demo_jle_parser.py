#!/usr/bin/env python3
"""
Demo script showing how to use the enhanced JLE parser with filename-based metadata extraction
"""
import os
from config import parse_jle_with_filename, extract_jle_data
from io import BytesIO


def demo_usage():
    """Demonstrate the usage of the new JLE parser functions"""
    print("Enhanced JLE Parser Demo")
    print("=" * 50)
    
    print("\n1. Using parse_jle_with_filename() function:")
    print("   This function takes a file path and extracts both content and metadata from the filename")
    
    # Example filename pattern: "DSO_20243_565.JLE" 
    # Where "2024" is the year and "3" indicates Summer semester
    print("   - Filename pattern recognized: YYYYX where XXXX is year and X is semester")
    print("   - Semester mapping: 1 = 1st Semester, 2 = 2nd Semester, 3 = Summer")
    
    print("\n2. Using extract_jle_data() function:")
    print("   This function works with uploaded file objects (like in Streamlit) and also extracts metadata")
    

def find_sample_jle_files():
    """Look for any existing JLE files in the project directory"""
    print("\n3. Looking for existing JLE files in the project...")
    
    jle_files = []
    for root, dirs, files in os.walk("."):
        for file in files:
            if file.lower().endswith('.jle'):
                full_path = os.path.join(root, file)
                jle_files.append(full_path)
                
    if jle_files:
        print(f"   Found {len(jle_files)} JLE file(s):")
        for jle_file in jle_files:
            print(f"     - {jle_file}")
    else:
        print("   No JLE files found in the project directory.")
        print("   You can use the parser with any JLE file by calling:")
        print("   df = parse_jle_with_filename('/path/to/your/file.JLE')")
    
    return jle_files


def demo_with_created_file():
    """Create a sample JLE file to demonstrate the functionality"""
    print("\n4. Creating a sample JLE file to demonstrate the parser...")
    
    import tempfile
    
    # Sample JLE content
    sample_content = (
        "2506   FBACC104 Introduction to Accounting 730AM- 900AM  MWF  C203 3      JOHN DOE\n"
        "2607   FCENG101 Advanced English 900AM-1030AM  TTH  A105 3      JANE SMITH\n"
        "2708   FMATH202 Calculus II 1030AM-1200PM  MWF  B301 4      ALICE JOHNSON\n"
    ).encode('latin1')
    
    # Create a temporary file with the expected naming pattern
    with tempfile.NamedTemporaryFile(suffix=".JLE", prefix="DSO_20242_", delete=False) as tmp_file:
        tmp_file.write(sample_content)
        temp_path = tmp_file.name
    
    try:
        print(f"   Created temporary file: {os.path.basename(temp_path)}")
        
        # Demonstrate the parse_jle_with_filename function
        df = parse_jle_with_filename(temp_path)
        
        print(f"   Successfully parsed {len(df)} records")
        print(f"   Academic Year: {df['Academic Year'].iloc[0] if not df.empty else 'N/A'}")
        print(f"   Semester: {df['Semester'].iloc[0] if not df.empty else 'N/A'}")
        
        print("\n   Sample of parsed data:")
        if not df.empty:
            print(df[['Subject Code', 'Subject Title', 'Lecturer', 'Academic Year', 'Semester']].to_string(index=False))
    
    finally:
        # Clean up
        os.unlink(temp_path)


if __name__ == "__main__":
    demo_usage()
    find_sample_jle_files()
    demo_with_created_file()
    
    print("\n" + "=" * 50)
    print("The enhanced JLE parser is now available in config.py")
    print("- Includes filename-based academic year and semester extraction")
    print("- Maintains all original parsing functionality")
    print("- Ready to use in your application!")