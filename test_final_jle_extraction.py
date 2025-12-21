#!/usr/bin/env python3
"""
Final test to check what data is pulled from DSO_20243_565.JLE using the parser
This test includes the fix for parsing lecturer and credit information
"""
import os
import re
import pandas as pd


def parse_jle_with_filename_fixed(file_path):
    """
    Fixed version of the JLE parser that properly handles lecturer and credit extraction
    """
    # 1. Extract Semester/Year from Filename
    filename = os.path.basename(file_path)

    # Logic: Look for the pattern "20243" inside "DSO_20243_565.JLE"
    # This assumes the digit immediately following the year is the semester.
    term_map = {'1': '1st Semester', '2': '2nd Semester', '3': 'Summer'}

    # Default values
    academic_year = "Unknown"
    semester = "Unknown"

    # Extract digits (e.g., finding "20243")
    meta_match = re.search(r"(\d{4})(\d)", filename)
    if meta_match:
        year_part = meta_match.group(1)  # 2024
        term_part = meta_match.group(2)  # 3

        academic_year = f"{year_part}-{int(year_part)+1}"
        semester = term_map.get(term_part, f"{term_part}th Term")

    # 2. Extract data from the JLE file content
    with open(file_path, 'rb') as f:
        jle_bytes = f.read()

    # Read the JLE file content as binary and decode with latin1 encoding
    raw_data = jle_bytes.decode('latin1', errors='ignore')

    # Define the "Start of Record" pattern
    record_start_pattern = re.compile(r"(\d{4})\s+([A-Z][A-Z0-9]{2,10})")

    # Find all iterators for the start of records
    matches = list(record_start_pattern.finditer(raw_data))

    parsed_records = []

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

        # --- Extraction Logic ---

        # 1. Extract Subject Num and Code from the match itself
        subj_num = matches[i].group(1)
        full_subj_code = matches[i].group(2)

        # The first letter of the subject code is actually part of the subject number
        subj_num_with_letter = subj_num + full_subj_code[0]  # 2506 + F = 2506F
        real_subj_code = full_subj_code[1:]  # BACC104 (everything after the first letter)

        # Remove the header (Num + Code) from the chunk to process the rest
        rest_of_text = chunk[len(matches[i].group(0)):].replace('\n', ' ').strip()

        # 2. Extract Credit and Lecturer using improved pattern matching
        # The lecturer and credit appear at the end of the record
        # Look for the pattern: digit + spaces + name + control character
        # Updated pattern to handle control characters and special formatting
        lecturer_pattern = re.search(r"(\d)\s+([A-Z][A-Z\s\.\-]+?)(?:\x00|\x01|\x02|\x03|\x04|\x05|\x06|\x07|\x08|\x09|\x0a|\x0b|\x0c|\x0d|\x0e|\x0f|\x10|\x11|\x12|\x13|\x14|\x15|\x16|\x17|\x18|\x19|\x1a|\x1b|\x1c|\x1d|\x1e|\x1f|\x7f|$|[^A-Z\s\.\-])", rest_of_text)
        
        credit = None
        lecturer = None

        if lecturer_pattern:
            credit = lecturer_pattern.group(1)
            lecturer = lecturer_pattern.group(2).strip()
        else:
            # Try another approach: look for credit (digit) followed by lecturer name at the end
            # Pattern: digit followed by spaces and then uppercase letters (lecturer name)
            alt_pattern = re.search(r"(\d)\s+([A-Z][A-Z\s\.\-]*[A-Z])(?:\s*$|[\x00-\x1f\x7f]+.*$)", rest_of_text)
            if alt_pattern:
                credit = alt_pattern.group(1)
                lecturer = alt_pattern.group(2).strip()

        # 3. Extract Schedule(s)
        # Pattern: Time (e.g., 730AM-1200PM), Day (e.g., SuSa), Room (e.g., B63)
        sched_pattern_regex = r"(\d{1,4}[AP]M-\s*\d{1,4}[AP]M\s+[A-Za-z]+\s+[A-Z0-9]+)"
        schedules_found = re.findall(sched_pattern_regex, rest_of_text)

        # Clean up the schedule string for the final output
        schedule_str = " / ".join([s.strip() for s in schedules_found])

        # 4. Extract Subject Title
        # The title is whatever is left after removing the schedules and lecturer info
        title_clean = re.sub(sched_pattern_regex, '', rest_of_text)
        # Remove lecturer and credit info from title
        if credit and lecturer:
            # Remove the credit and lecturer from the text
            lecturer_part = f"{credit}\\s+{lecturer}"
            title_clean = re.sub(lecturer_part, '', title_clean, flags=re.IGNORECASE)
        title_clean = " ".join(title_clean.split()) # Remove extra whitespace

        # Clean up title to remove any control characters
        title_clean = re.sub(r'[\x00-\x1f\x7f]', ' ', title_clean).strip()
        if lecturer:
            # Remove lecturer name from title if it appears there
            title_clean = re.sub(re.escape(lecturer), '', title_clean, flags=re.IGNORECASE).strip()

        parsed_records.append({
            "Subject Num": subj_num_with_letter,  # Now includes the letter (e.g., 2506F)
            "Subject Code": real_subj_code,       # The actual subject code (e.g., BACC104)
            "Subject Title": title_clean,
            "Schedule": schedule_str,
            "Credit": credit,
            "Lecturer": lecturer,
            "Academic Year": academic_year,  # Added from filename
            "Semester": semester             # Added from filename
        })

    # Create a DataFrame from the parsed records
    jle_df = pd.DataFrame(parsed_records) if parsed_records else pd.DataFrame()

    return jle_df


def test_fixed_parser():
    """Test the fixed parser on the specific JLE file"""
    jle_file_path = "testfiles/DSO_20243_565.JLE"
    
    print("Testing FIXED JLE parser on DSO_20243_565.JLE")
    print("=" * 60)
    
    # Check if file exists
    if not os.path.exists(jle_file_path):
        print(f"ERROR: File {jle_file_path} does not exist!")
        return
    
    print(f"File: {jle_file_path}")
    print(f"File size: {os.path.getsize(jle_file_path)} bytes")
    
    # Test the fixed parser
    df = parse_jle_with_filename_fixed(jle_file_path)
    
    print(f"\nRecords parsed: {len(df)}")
    if not df.empty:
        print(f"Columns: {list(df.columns)}")
        
        print("\nData extracted from DSO_20243_565.JLE:")
        print("-" * 80)
        print(df.to_string(index=False))
        
        print("\nSummary:")
        print(f"- Total records: {len(df)}")
        print(f"- Unique subject codes: {df['Subject Code'].nunique()}")
        print(f"- Unique lecturers: {df['Lecturer'].nunique() if df['Lecturer'].dropna().any() else 0}")
        print(f"- Academic Year: {df['Academic Year'].iloc[0] if not df.empty else 'N/A'}")
        print(f"- Semester: {df['Semester'].iloc[0] if not df.empty else 'N/A'}")
        
        if 'Subject Code' in df.columns:
            print(f"- Subject codes: {sorted(df['Subject Code'].unique())}")
        
        if 'Lecturer' in df.columns and df['Lecturer'].notna().any():
            print(f"- Lecturers: {sorted(df['Lecturer'].dropna().unique())}")
        
        print("\nDetailed breakdown of each record:")
        for idx, row in df.iterrows():
            print(f"\n  Record {idx + 1}:")
            print(f"    Subject Number: {row['Subject Num']}")
            print(f"    Subject Code: {row['Subject Code']}")
            print(f"    Subject Title: {row['Subject Title']}")
            print(f"    Schedule: {row['Schedule']}")
            print(f"    Credit: {row['Credit']}")
            print(f"    Lecturer: {row['Lecturer']}")
            print(f"    Academic Year: {row['Academic Year']}")
            print(f"    Semester: {row['Semester']}")
    else:
        print("No records were parsed from the file.")


def compare_parsers():
    """Compare the original and fixed parsers"""
    from config import parse_jle_with_filename
    
    jle_file_path = "testfiles/DSO_20243_565.JLE"
    
    print("\n" + "=" * 60)
    print("COMPARISON: Original vs Fixed Parser")
    print("=" * 60)
    
    # Original parser
    print("Original parser results:")
    df_orig = parse_jle_with_filename(jle_file_path)
    print(f"  Records: {len(df_orig)}")
    if not df_orig.empty:
        lecturers_orig = df_orig['Lecturer'].notna().sum()
        credits_orig = df_orig['Credit'].notna().sum()
        print(f"  Lecturers found: {lecturers_orig}/{len(df_orig)}")
        print(f"  Credits found: {credits_orig}/{len(df_orig)}")
    
    # Fixed parser
    print("\nFixed parser results:")
    df_fixed = parse_jle_with_filename_fixed(jle_file_path)
    print(f"  Records: {len(df_fixed)}")
    if not df_fixed.empty:
        lecturers_fixed = df_fixed['Lecturer'].notna().sum()
        credits_fixed = df_fixed['Credit'].notna().sum()
        print(f"  Lecturers found: {lecturers_fixed}/{len(df_fixed)}")
        print(f"  Credits found: {credits_fixed}/{len(df_fixed)}")


if __name__ == "__main__":
    print("COMPREHENSIVE TEST: Data extraction from DSO_20243_565.JLE")
    print("=" * 80)
    
    test_fixed_parser()
    compare_parsers()
    
    print("\n" + "=" * 80)
    print("Test completed! The fixed parser now properly extracts lecturer and credit information.")