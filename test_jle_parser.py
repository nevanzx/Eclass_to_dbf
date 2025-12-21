#!/usr/bin/env python3
"""
Test script for the JLE parser with filename-based metadata extraction
"""
import os
import tempfile
from config import parse_jle_with_filename, extract_jle_data
from io import BytesIO


def create_sample_jle_content():
    """Create sample JLE content for testing"""
    # Sample JLE content with recognizable patterns
    content = (
        "2506   FBACC104 Introduction to Accounting 730AM- 900AM  MWF  C203 3      JOHN DOE\n"
        "2607   FCENG101 Advanced English 900AM-1030AM  TTH  A105 3      JANE SMITH\n"
        "2708   FMATH202 Calculus II 1030AM-1200PM  MWF  B301 4      ALICE JOHNSON\n"
    )
    return content.encode('latin1')


def test_parse_jle_with_filename():
    """Test the parse_jle_with_filename function"""
    print("Testing parse_jle_with_filename function...")
    
    # Create a temporary JLE file with a filename that follows the expected pattern
    sample_content = create_sample_jle_content()
    
    with tempfile.NamedTemporaryFile(suffix=".JLE", prefix="DSO_20243_", delete=False) as tmp_file:
        tmp_file.write(sample_content)
        temp_path = tmp_file.name
    
    try:
        # Test the function
        df = parse_jle_with_filename(temp_path)
        
        print(f"Successfully parsed JLE file: {os.path.basename(temp_path)}")
        print(f"Number of records parsed: {len(df)}")
        print(f"Columns in DataFrame: {list(df.columns)}")
        
        if not df.empty:
            print("\nSample record:")
            print(df.iloc[0])
            
            # Check if Academic Year and Semester columns were added
            if 'Academic Year' in df.columns and 'Semester' in df.columns:
                print(f"\nAcademic Year extracted: {df['Academic Year'].iloc[0]}")
                print(f"Semester extracted: {df['Semester'].iloc[0]}")
            else:
                print("\nERROR: Academic Year or Semester columns not found!")
        else:
            print("WARNING: No records were parsed from the JLE file.")
    
    finally:
        # Clean up the temporary file
        os.unlink(temp_path)


def test_extract_jle_data():
    """Test the extract_jle_data function"""
    print("\nTesting extract_jle_data function...")
    
    # Create a BytesIO object simulating an uploaded file
    sample_content = create_sample_jle_content()
    jle_file = BytesIO(sample_content)
    jle_file.name = "DSO_20243_565.JLE"  # Simulate a typical filename
    
    try:
        # Test the function
        jle_info = extract_jle_data(jle_file)
        
        print(f"Successfully processed JLE file: {jle_info['filename']}")
        print(f"File size: {jle_info['size']} bytes")
        print(f"Number of courses: {jle_info['total_courses']}")
        print(f"Academic Year: {jle_info.get('academic_year', 'Not found')}")
        print(f"Semester: {jle_info.get('semester', 'Not found')}")
        
        if jle_info['course_data'] is not None and not jle_info['course_data'].empty:
            print(f"Course data shape: {jle_info['course_data'].shape}")
            print("\nSample course data:")
            print(jle_info['course_data'].head())
        else:
            print("No course data was extracted.")
    
    except Exception as e:
        print(f"Error processing JLE data: {e}")
        import traceback
        traceback.print_exc()


def test_filename_parsing():
    """Test different filename patterns to ensure semester/year extraction works"""
    print("\nTesting filename pattern recognition...")
    
    test_filenames = [
        "DSO_20241_565.JLE",  # Should be 1st Semester
        "DSO_20242_565.JLE",  # Should be 2nd Semester  
        "DSO_20243_565.JLE",  # Should be Summer
        "REPORT_20231_123.JLE",  # Should be 1st Semester
        "TEST_20252_FILE.JLE",  # Should be 2nd Semester
    ]
    
    sample_content = create_sample_jle_content()
    
    for filename in test_filenames:
        with tempfile.NamedTemporaryFile(suffix=".JLE", prefix=filename[:-4]+"_", delete=False) as tmp_file:
            tmp_file.write(sample_content)
            temp_path = tmp_file.name
        
        try:
            df = parse_jle_with_filename(temp_path)
            if not df.empty:
                print(f"  {filename}: Academic Year = {df['Academic Year'].iloc[0]}, "
                      f"Semester = {df['Semester'].iloc[0]}")
            else:
                print(f"  {filename}: No data parsed")
        except Exception as e:
            print(f"  {filename}: Error - {e}")
        finally:
            os.unlink(temp_path)


if __name__ == "__main__":
    print("Running tests for JLE parser with filename-based metadata extraction...\n")
    
    test_parse_jle_with_filename()
    test_extract_jle_data()
    test_filename_parsing()
    
    print("\nAll tests completed!")