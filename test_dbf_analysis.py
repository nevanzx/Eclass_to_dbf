#!/usr/bin/env python3
"""
Test script to analyze the DBF file and look for patterns that connect it to the JLE file
"""
import os
import pandas as pd
from dbf import Table


def analyze_dbf_file():
    """Analyze the DBF file to see its structure and content"""
    dbf_file_path = "testfiles/DSO_20243_2506B_BACC104_565.DBF"

    print("Analyzing DBF file:", dbf_file_path)
    print("=" * 60)

    # Check if file exists
    if not os.path.exists(dbf_file_path):
        print(f"ERROR: File {dbf_file_path} does not exist!")
        return

    print(f"File size: {os.path.getsize(dbf_file_path)} bytes")

    # Display filename components
    filename = os.path.basename(dbf_file_path)
    print(f"Filename: {filename}")

    # Break down the filename to understand the pattern
    parts = filename.replace('.DBF', '').split('_')
    print(f"Filename parts: {parts}")
    print("Pattern analysis:")
    print(f"  - Organization: {parts[0] if len(parts) > 0 else 'N/A'}")
    print(f"  - Academic Year/Semester: {parts[1] if len(parts) > 1 else 'N/A'} (likely 2024 year, 3 = Summer)")
    print(f"  - Subject Number: {parts[2] if len(parts) > 2 else 'N/A'}")
    print(f"  - Subject Code: {parts[3] if len(parts) > 3 else 'N/A'}")
    print(f"  - ID/Section: {parts[4] if len(parts) > 4 else 'N/A'}")

    try:
        # Load the DBF file using the dbf library
        table = Table(dbf_file_path)
        table.open()  # Open the table to access records

        print(f"\nDBF file info:")
        print(f"  - Number of records: {len(table)}")
        print(f"  - Number of fields: {table.field_count}")
        print(f"  - Field names: {table.field_names}")

        # Convert to pandas DataFrame
        records = []
        for record in table:
            record_dict = {}
            for field in table.field_names:
                record_dict[field] = record[field]
            records.append(record_dict)

        df = pd.DataFrame(records)

        print(f"\nDataFrame shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")

        print(f"\nFirst few records:")
        print(df.head().to_string(index=True))

        print(f"\nAll records:")
        print(df.to_string(index=True))

        # Check if any fields match what we'd expect from the JLE file
        print(f"\nAnalyzing potential connections to JLE file:")

        # Look for fields that might correspond to JLE data
        possible_matches = []

        for col in df.columns:
            try:
                unique_vals = df[col].unique()[:10].tolist()
                print(f"  - Field '{col}': unique values = {unique_vals}")  # Show first 10 unique values
            except TypeError:
                # Handle cases where unique() fails due to unhashable types
                unique_vals = list(set(df[col].dropna().astype(str).tolist()))[:10]
                print(f"  - Field '{col}': unique values (converted to str) = {unique_vals}")

            # Check if this field contains values similar to what we found in JLE
            if 'SUBJNUM' in col.upper():
                possible_matches.append(f"Subject Number field: {col}")
            elif 'SUBJCODE' in col.upper():
                possible_matches.append(f"Subject Code field: {col}")
            elif 'SUBJTITLE' in col.upper():
                possible_matches.append(f"Subject Title field: {col}")
            elif 'LECTURER' in col.upper():
                possible_matches.append(f"Lecturer field: {col}")
            elif 'CREDIT' in col.upper():
                possible_matches.append(f"Credit field: {col}")

        if possible_matches:
            print(f"\nPotential matches with JLE fields:")
            for match in possible_matches:
                print(f"  - {match}")

        # Compare with JLE data
        print(f"\nComparing with JLE file DSO_20243_565.JLE:")
        print(f"  - DBF Subject Number: {parts[2] if len(parts) > 2 else 'N/A'}")
        print(f"  - DBF Subject Code: {parts[3] if len(parts) > 3 else 'N/A'}")
        print(f"  - Expected from JLE: Subject Number 2506B, Subject Code BACC104")
        print(f"  - Match: {parts[2] == '2506B' and parts[3] == 'BACC104'}")

        table.close()  # Close the table when done

    except Exception as e:
        print(f"ERROR reading DBF file: {e}")
        import traceback
        traceback.print_exc()


def check_all_test_files():
    """Check all files in the testfiles directory for naming patterns"""
    testfiles_dir = "testfiles"
    
    print(f"\n" + "=" * 60)
    print("ANALYZING ALL FILES IN TESTFILES DIRECTORY")
    print("=" * 60)
    
    if not os.path.exists(testfiles_dir):
        print(f"ERROR: Directory {testfiles_dir} does not exist!")
        return
    
    files = os.listdir(testfiles_dir)
    
    print(f"Files found in {testfiles_dir}:")
    for file in files:
        print(f"  - {file}")
        if file.lower().endswith(('.dbf', '.jle')):
            parts = file.replace('.DBF', '').replace('.JLE', '').split('_')
            print(f"    Pattern: {parts}")
            if len(parts) >= 5:
                print(f"      Organization: {parts[0]}")
                print(f"      Year/Sem: {parts[1]}")
                print(f"      Subject Num: {parts[2]}")
                print(f"      Subject Code: {parts[3]}")
                print(f"      ID: {parts[4]}")


if __name__ == "__main__":
    print("DBF FILE ANALYSIS: DSO_20243_2506B_BACC104_565.DBF")
    print("=" * 80)
    
    analyze_dbf_file()
    check_all_test_files()
    
    print("\n" + "=" * 80)
    print("Analysis completed!")