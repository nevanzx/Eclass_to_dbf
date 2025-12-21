#!/usr/bin/env python3
"""
Summary test to show what data is extracted from DSO_20243_565.JLE
This is the final, clean version showing exactly what the parser extracts
"""
import os
from config import parse_jle_with_filename


def test_jle_data_extraction():
    """Test and display what data is extracted from the JLE file"""
    jle_file_path = "testfiles/DSO_20243_565.JLE"
    
    print("DATA EXTRACTION REPORT: DSO_20243_565.JLE")
    print("=" * 50)
    
    # Verify file exists
    if not os.path.exists(jle_file_path):
        print(f"ERROR: File {jle_file_path} does not exist!")
        return
    
    print(f"File: {jle_file_path}")
    print(f"Size: {os.path.getsize(jle_file_path)} bytes")
    
    # Parse the file
    df = parse_jle_with_filename(jle_file_path)
    
    print(f"\nTotal records extracted: {len(df)}")
    
    if len(df) > 0:
        print(f"Data fields per record: {list(df.columns)}")
        
        print(f"\nEXTRACTED DATA:")
        print("-" * 50)
        
        for idx, record in df.iterrows():
            print(f"\nRecord {idx + 1}:")
            print(f"  Subject Number: {record['Subject Num']}")
            print(f"  Subject Code: {record['Subject Code']}")
            print(f"  Subject Title: {record['Subject Title']}")
            print(f"  Schedule: {record['Schedule']}")
            print(f"  Credit: {record['Credit']}")  # This shows the issue - original parser doesn't extract these
            print(f"  Lecturer: {record['Lecturer']}")  # This shows the issue
            print(f"  Academic Year: {record['Academic Year']}")
            print(f"  Semester: {record['Semester']}")
        
        print(f"\nSUMMARY:")
        print(f"  Total records: {len(df)}")
        print(f"  Unique subject codes: {df['Subject Code'].nunique()}")
        print(f"  Unique subject codes: {sorted(df['Subject Code'].unique())}")
        print(f"  Academic Year: {df['Academic Year'].iloc[0] if len(df) > 0 else 'N/A'}")
        print(f"  Semester: {df['Semester'].iloc[0] if len(df) > 0 else 'N/A'}")
        
        # Note about parsing limitations
        lecturers_extracted = df['Lecturer'].notna().sum()
        credits_extracted = df['Credit'].notna().sum()
        
        print(f"\nNOTE: The original parser has limitations:")
        print(f"  Credits extracted: {credits_extracted}/{len(df)}")
        print(f"  Lecturers extracted: {lecturers_extracted}/{len(df)}")
        print(f"  (See test_final_jle_extraction.py for a fixed version that properly extracts this data)")


if __name__ == "__main__":
    test_jle_data_extraction()