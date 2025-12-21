import streamlit as st
import pandas as pd
import tempfile
import base64
from docxtpl import DocxTemplate
import io
import json
import tempfile
import os
import re
from dbf import Table


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
        # Updated pattern to handle control characters and special formatting
        lecturer_pattern = re.search(r"(\d)\s+([A-Z][A-Z\s\.\-]*[A-Z])(?:\s*$|[\x00-\x1f\x7f\s]*$)", rest_of_text)

        credit = None
        lecturer = None

        if lecturer_pattern:
            credit = lecturer_pattern.group(1)
            lecturer = lecturer_pattern.group(2).strip()
            # Remove this part from the text so we can find the Title/Schedule
            rest_of_text = rest_of_text[:lecturer_pattern.start()].strip()
        else:
            # Try another approach: look for credit (digit) followed by lecturer name at the end
            # Pattern: digit followed by spaces and then uppercase letters (lecturer name)
            alt_pattern = re.search(r"(\d)\s+([A-Z][A-Z\s\.\-]+?)(?:\s+(?:\d\s+)?[A-Z]|$)", rest_of_text)
            if alt_pattern:
                credit = alt_pattern.group(1)
                lecturer = alt_pattern.group(2).strip()
                rest_of_text = rest_of_text[:alt_pattern.start()].strip()

        # 3. Extract Schedule(s) - distinguish between LEC and LAB
        # Pattern: Time (e.g., 730AM-1200PM), Day (e.g., SuSa), Room (e.g., B63)
        sched_pattern_regex = r"(\d{1,4}[AP]M-\s*\d{1,4}[AP]M\s+[A-Za-z]+\s+[A-Z0-9]+)"
        schedules_found = re.findall(sched_pattern_regex, rest_of_text)

        # Initialize LEC and LAB schedules as empty
        lec_schedule_str = ""
        lab_schedule_str = ""

        # Try to identify LEC and LAB schedules if they're explicitly marked in the text
        # Look for LEC/LAB indicators in the context of schedules
        full_text_for_context = rest_of_text
        lec_pattern = r"LEC.*?(\d{1,4}[AP]M-\s*\d{1,4}[AP]M\s+[A-Za-z]+\s+[A-Z0-9]+)"
        lab_pattern = r"LAB.*?(\d{1,4}[AP]M-\s*\d{1,4}[AP]M\s+[A-Za-z]+\s+[A-Z0-9]+)"

        # Extract LEC schedules if explicitly marked
        lec_matches = re.findall(lec_pattern, rest_of_text, re.IGNORECASE)
        if lec_matches:
            lec_schedule_str = " / ".join([s.strip() for s in lec_matches])

        # Extract LAB schedules if explicitly marked
        lab_matches = re.findall(lab_pattern, rest_of_text, re.IGNORECASE)
        if lab_matches:
            lab_schedule_str = " / ".join([s.strip() for s in lab_matches])

        # If no explicit LEC/LAB markers found, use heuristic approach
        if not lec_schedule_str and not lab_schedule_str:
            if len(schedules_found) >= 2:
                # If there are 2 or more schedules, assume first is LEC, second is LAB (common pattern)
                lec_schedule_str = schedules_found[0] if schedules_found else ""
                lab_schedule_str = schedules_found[1] if len(schedules_found) > 1 else ""
            else:
                # If only one schedule, assign to LEC
                lec_schedule_str = schedules_found[0] if schedules_found else ""

        # Create combined schedule string
        all_schedule_str = " / ".join([s.strip() for s in schedules_found])

        # 4. Extract Subject Title
        # The title is whatever is left after removing the schedules
        # We use re.sub to remove the schedules we found
        title_clean = re.sub(sched_pattern_regex, '', rest_of_text)
        # Also remove the lecturer and credit from the title if they weren't removed by the main pattern
        if credit and lecturer:
            lecturer_part = f"{credit}\\s+{lecturer}"
            title_clean = re.sub(lecturer_part, '', title_clean, flags=re.IGNORECASE)
        title_clean = " ".join(title_clean.split()) # Remove extra whitespace

        # Clean up title to remove any control characters
        title_clean = re.sub(r'[\x00-\x1f\x7f]', ' ', title_clean).strip()

        parsed_records.append({
            "Subject Num": subj_num_with_letter,  # Now includes the letter (e.g., 2506F)
            "Subject Code": real_subj_code,       # The actual subject code (e.g., BACC104)
            "Subject Title": title_clean,
            "Schedule": all_schedule_str,          # All schedules combined
            "LEC_Schedule": lec_schedule_str,      # Lecture schedule
            "LAB_Schedule": lab_schedule_str,      # Laboratory schedule
            "Credit": credit,
            "Lecturer": lecturer,
            "Academic Year": academic_year,  # Added from filename
            "Semester": semester             # Added from filename
        })

    # Create a DataFrame from the parsed records
    jle_df = pd.DataFrame(parsed_records) if parsed_records else pd.DataFrame()

    return jle_df


def find_matching_dbf(jle_file_path, dbf_directory="testfiles"):
    """Find DBF file that matches the JLE file based on naming patterns"""
    # Extract metadata from JLE file
    jle_df = parse_jle_with_filename_fixed(jle_file_path)

    if jle_df.empty:
        return None, None

    # Get all DBF files in the directory
    dbf_files = [f for f in os.listdir(dbf_directory) if f.lower().endswith('.dbf')]

    for dbf_file in dbf_files:
        dbf_path = os.path.join(dbf_directory, dbf_file)

        # Extract parts from DBF filename to match against JLE
        # Pattern: ORG_YYYYX_SUBJNUM_SUBJCODE_ID.DBF
        parts = dbf_file.replace('.DBF', '').split('_')
        if len(parts) >= 4:
            year_semester = parts[1]  # YYYYX
            subj_num = parts[2]       # SUBJNUM
            subj_code = parts[3]      # SUBJCODE

            # Check if any of the JLE records match
            for _, jle_record in jle_df.iterrows():
                jle_academic_year = jle_record['Academic Year']  # Full format like "2024-2025"
                jle_year = jle_academic_year.split('-')[0]  # Get first part like "2024"
                jle_year_semester = f"{jle_year}{get_semester_digit(jle_record['Semester'])}"  # Full format like "20243"
                jle_subj_num = jle_record['Subject Num']
                jle_subj_code = jle_record['Subject Code']

                # Check if the components match
                if (year_semester == jle_year_semester and
                    subj_num == jle_subj_num and
                    subj_code == jle_subj_code):
                    return dbf_path, jle_record  # Return the matching DBF path and JLE record

    return None, None


def find_matching_dbf_from_jle_data(jle_data, dbf_directory="testfiles"):
    """Find DBF file that matches the JLE data based on naming patterns"""
    # Extract metadata from JLE data
    if jle_data is None or 'course_data' not in jle_data or jle_data['course_data'] is None or jle_data['course_data'].empty:
        return None, None

    jle_df = jle_data['course_data']

    # Get all DBF files in the directory
    dbf_files = [f for f in os.listdir(dbf_directory) if f.lower().endswith('.dbf')]

    for dbf_file in dbf_files:
        dbf_path = os.path.join(dbf_directory, dbf_file)

        # Extract parts from DBF filename to match against JLE
        # Pattern: ORG_YYYYX_SUBJNUM_SUBJCODE_ID.DBF
        parts = dbf_file.replace('.DBF', '').split('_')
        if len(parts) >= 4:
            year_semester = parts[1]  # YYYYX
            subj_num = parts[2]       # SUBJNUM
            subj_code = parts[3]      # SUBJCODE

            # Check if any of the JLE records match
            for _, jle_record in jle_df.iterrows():
                jle_academic_year = jle_record.get('Academic Year', '')  # Full format like "2024-2025"
                if jle_academic_year:
                    jle_year = jle_academic_year.split('-')[0]  # Get first part like "2024"
                    jle_year_semester = f"{jle_year}{get_semester_digit(jle_record.get('Semester', ''))}"  # Full format like "20243"
                    jle_subj_num = jle_record.get('Subject Num', '')
                    jle_subj_code = jle_record.get('Subject Code', '')

                    # Check if the components match
                    if (year_semester == jle_year_semester and
                        subj_num == jle_subj_num and
                        subj_code == jle_subj_code):
                        return dbf_path, jle_record  # Return the matching DBF path and JLE record

    return None, None


def get_semester_digit(semester_str):
    """Convert semester string to digit"""
    semester_map = {
        '1st Semester': '1',
        '2nd Semester': '2',
        'Summer': '3'
    }
    return semester_map.get(semester_str, '0')


def extract_dbf_data(dbf_path):
    """Extract data from DBF file"""
    table = Table(dbf_path)
    table.open()

    records = []
    for record in table:
        record_dict = {}
        for field in table.field_names:
            record_dict[field] = record[field]
        records.append(record_dict)

    df = pd.DataFrame(records)
    table.close()

    return df


def populate_template_with_jle_dbf_data(template_path, jle_file_path, dbf_directory="testfiles"):
    """
    Populate the template with data from matched JLE and DBF files using docxtpl
    """
    # Find matching DBF file for the JLE
    matching_dbf, jle_record = find_matching_dbf(jle_file_path, dbf_directory)

    if not matching_dbf:
        st.warning("No matching DBF file found for the JLE file.")
        # Still proceed with JLE-only data
        context = {
            'SY': jle_record['Academic Year'] if jle_record is not None else 'N/A',
            'SC': jle_record['Subject Code'] if jle_record is not None else 'N/A',
            'SN': jle_record['Subject Num'] if jle_record is not None else 'N/A',
            'SEM': jle_record['Semester'] if jle_record is not None else 'N/A',
            'LECTURER': jle_record['Lecturer'] if jle_record is not None else 'N/A',
            'CREDIT': jle_record['Credit'] if jle_record is not None else 'N/A',
            'SCHED': jle_record['Schedule'] if jle_record is not None else 'N/A',
            'TITLE': jle_record['Subject Title'] if jle_record is not None else 'N/A',
            'STUDENT_COUNT': 'N/A',  # No DBF data
            # Additional mappings as requested
            'ST': jle_record['Subject Title'] if jle_record is not None else 'N/A',  # Subject Title
            'Sem': jle_record['Semester'] if jle_record is not None else 'N/A',      # Semester
            'OC': jle_record['Subject Num'] if jle_record is not None else 'N/A',    # Subject Num
            'Time': pd.Timestamp.now().strftime('%Y-%m-%d'),  # Current day only
            'Faculty': jle_record['Lecturer'] if jle_record is not None else 'N/A',   # Lecturer
            'SCHED': jle_record['Schedule'] if jle_record is not None else 'N/A',    # All schedules
            'LeS': jle_record['LEC_Schedule'] if jle_record is not None else 'N/A',  # Lecture Schedule
            'LaS': jle_record['LAB_Schedule'] if jle_record is not None else 'N/A'   # Laboratory Schedule
        }
    else:
        # Extract data from DBF
        dbf_data = extract_dbf_data(matching_dbf)

        # Prepare context with both JLE and DBF data
        context = {
            'SY': jle_record['Academic Year'],  # School Year
            'SC': jle_record['Subject Code'],   # Subject Code
            'SN': jle_record['Subject Num'],    # Subject Number
            'SEM': jle_record['Semester'],      # Semester
            'LECTURER': jle_record['Lecturer'], # Lecturer
            'CREDIT': jle_record['Credit'],     # Credit
            'SCHED': jle_record['Schedule'],    # Schedule
            'TITLE': jle_record['Subject Title'], # Subject Title
            'STUDENT_COUNT': len(dbf_data),      # Number of students from DBF
            # Additional mappings as requested
            'ST': jle_record['Subject Title'],  # Subject Title
            'Sem': jle_record['Semester'],      # Semester
            'OC': jle_record['Subject Num'],    # Subject Num
            'Time': pd.Timestamp.now().strftime('%Y-%m-%d'),  # Current day only
            'Faculty': jle_record['Lecturer']   # Lecturer
        }

    # Handle the student data table specifically - look for the table with student information
    # Use the DBF data that was extracted if available
    student_data = []
    statistics = {}
    if matching_dbf:
        # Extract DBF data to populate the student table
        dbf_data = extract_dbf_data(matching_dbf)
        if dbf_data is not None and not dbf_data.empty:
            student_data, statistics = process_student_data(dbf_data)

    # Add student data to context
    context['students'] = student_data
    # Add statistics to context
    context.update(statistics)

    # Process the template with docxtpl
    doc = DocxTemplate(template_path)
    doc.render(context)

    # Save to BytesIO and return
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def populate_template_with_jle_dbf_data_from_jle_data(template_path, jle_data, dbf_directory="testfiles"):
    """
    Populate the template with data from matched JLE data and DBF files using docxtpl
    """
    # Find matching DBF file for the JLE data
    matching_dbf, jle_record = find_matching_dbf_from_jle_data(jle_data, dbf_directory)

    if not matching_dbf:
        # Still proceed with JLE-only data
        context = {
            'SY': jle_record.get('Academic Year', 'N/A') if jle_record is not None else 'N/A',  # School Year
            'SC': jle_record.get('Subject Code', 'N/A') if jle_record is not None else 'N/A',   # Subject Code
            'SN': jle_record.get('Subject Num', 'N/A') if jle_record is not None else 'N/A',    # Subject Number
            'SEM': jle_record.get('Semester', 'N/A') if jle_record is not None else 'N/A',      # Semester
            'LECTURER': jle_record.get('Lecturer', 'N/A') if jle_record is not None else 'N/A', # Lecturer
            'CREDIT': jle_record.get('Credit', 'N/A') if jle_record is not None else 'N/A',     # Credit
            'SCHED': jle_record.get('Schedule', 'N/A') if jle_record is not None else 'N/A',    # Schedule
            'TITLE': jle_record.get('Subject Title', 'N/A') if jle_record is not None else 'N/A', # Subject Title
            'STUDENT_COUNT': 'N/A',  # No DBF data
            # Additional mappings as requested
            'ST': jle_record.get('Subject Title', 'N/A') if jle_record is not None else 'N/A',  # Subject Title
            'Sem': jle_record.get('Semester', 'N/A') if jle_record is not None else 'N/A',      # Semester
            'OC': jle_record.get('Subject Num', 'N/A') if jle_record is not None else 'N/A',    # Subject Num
            'Time': pd.Timestamp.now().strftime('%Y-%m-%d'),  # Current day only
            'Faculty': jle_record.get('Lecturer', 'N/A') if jle_record is not None else 'N/A',   # Lecturer
            'SCHED': jle_record.get('Schedule', 'N/A') if jle_record is not None else 'N/A',    # All schedules
            'LeS': jle_record.get('LEC_Schedule', 'N/A') if jle_record is not None else 'N/A',  # Lecture Schedule
            'LaS': jle_record.get('LAB_Schedule', 'N/A') if jle_record is not None else 'N/A'   # Laboratory Schedule
        }
    else:
        # Extract data from DBF
        dbf_data = extract_dbf_data(matching_dbf)

        # Prepare context with both JLE and DBF data
        context = {
            'SY': jle_record.get('Academic Year', 'N/A'),  # School Year
            'SC': jle_record.get('Subject Code', 'N/A'),   # Subject Code
            'SN': jle_record.get('Subject Num', 'N/A'),    # Subject Number
            'SEM': jle_record.get('Semester', 'N/A'),      # Semester
            'LECTURER': jle_record.get('Lecturer', 'N/A'), # Lecturer
            'CREDIT': jle_record.get('Credit', 'N/A'),     # Credit
            'SCHED': jle_record.get('Schedule', 'N/A'),    # Schedule
            'TITLE': jle_record.get('Subject Title', 'N/A'), # Subject Title
            'STUDENT_COUNT': len(dbf_data),                # Number of students from DBF
            # Additional mappings as requested
            'ST': jle_record.get('Subject Title', 'N/A'),  # Subject Title
            'Sem': jle_record.get('Semester', 'N/A'),      # Semester
            'OC': jle_record.get('Subject Num', 'N/A'),    # Subject Num
            'Time': pd.Timestamp.now().strftime('%Y-%m-%d'),  # Current day only
            'Faculty': jle_record.get('Lecturer', 'N/A')   # Lecturer
        }

    # Handle the student data table specifically - look for the table with student information
    # Use the DBF data that was extracted if available
    student_data = []
    statistics = {}
    if matching_dbf:
        # Extract DBF data to populate the student table
        dbf_data = extract_dbf_data(matching_dbf)
        if dbf_data is not None and not dbf_data.empty:
            student_data, statistics = process_student_data(dbf_data)

    # Add student data to context
    context['students'] = student_data
    # Add statistics to context
    context.update(statistics)

    # Process the template with docxtpl
    doc = DocxTemplate(template_path)
    doc.render(context)

    # Save to BytesIO and return
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def populate_template_with_jle_and_uploaded_dbf_data(template_path, jle_data, uploaded_dbf_filename, dbf_data_df):
    """
    Populate the template with data from JLE and uploaded DBF file using docxtpl
    """
    matched_course = None
    matching_successful = False

    # Strategy 1: Try to match based on filename pattern
    # Pattern: ORG_YYYYX_SUBJNUM_SUBJCODE_ID.DBF
    if uploaded_dbf_filename:
        parts = uploaded_dbf_filename.replace('.DBF', '').replace('.dbf', '').split('_')
        if len(parts) >= 4:
            year_semester = parts[1]  # YYYYX
            subj_num = parts[2]       # SUBJNUM
            subj_code = parts[3]      # SUBJCODE

            # Find the matching course in JLE data
            if jle_data and 'course_data' in jle_data and jle_data['course_data'] is not None and not jle_data['course_data'].empty:
                jle_df = jle_data['course_data']

                # Find the matching record
                for _, jle_record in jle_df.iterrows():
                    jle_academic_year = jle_record.get('Academic Year', '')  # Full format like "2024-2025"
                    if jle_academic_year:
                        jle_year = jle_academic_year.split('-')[0]  # Get first part like "2024"
                        jle_year_semester = f"{jle_year}{get_semester_digit(jle_record.get('Semester', ''))}"  # Full format like "20243"
                        jle_subj_num = jle_record.get('Subject Num', '')
                        jle_subj_code = jle_record.get('Subject Code', '')

                        # Check if the components match
                        if (year_semester == jle_year_semester and
                            subj_num == jle_subj_num and
                            subj_code == jle_subj_code):
                            matched_course = jle_record
                            matching_successful = True
                            break  # Found the match, exit loop

    # If no match found by filename, try other strategies
    if matched_course is None and jle_data and 'course_data' in jle_data and jle_data['course_data'] is not None and not jle_data['course_data'].empty:
        jle_df = jle_data['course_data']

        # Strategy 2: If we have student data, we might be able to match on other criteria
        # For now, just take the first course if there's only one
        if len(jle_df) == 1:
            matched_course = jle_df.iloc[0]
            matching_successful = True

    # Prepare context based on whether we found a match
    if matched_course is not None:
        # Found matching course, prepare detailed context
        context = {
            'SY': matched_course.get('Academic Year', 'N/A'),  # School Year
            'SC': matched_course.get('Subject Code', 'N/A'),   # Subject Code
            'SN': matched_course.get('Subject Num', 'N/A'),    # Subject Number
            'SEM': matched_course.get('Semester', 'N/A'),      # Semester
            'LECTURER': matched_course.get('Lecturer', 'N/A'), # Lecturer
            'CREDIT': matched_course.get('Credit', 'N/A'),     # Credit
            'SCHED': matched_course.get('Schedule', 'N/A'),    # Schedule
            'TITLE': matched_course.get('Subject Title', 'N/A'), # Subject Title
            'STUDENT_COUNT': len(dbf_data_df) if dbf_data_df is not None else 'N/A',  # Number of students from DBF
            # Additional mappings as requested
            'ST': matched_course.get('Subject Title', 'N/A'),  # Subject Title
            'Sem': matched_course.get('Semester', 'N/A'),      # Semester
            'OC': matched_course.get('Subject Num', 'N/A'),    # Subject Num
            'Time': pd.Timestamp.now().strftime('%Y-%m-%d'),  # Current day only
            'Faculty': matched_course.get('Lecturer', 'N/A'),   # Lecturer
            'SCHED': matched_course.get('Schedule', 'N/A'),    # All schedules
            'LeS': matched_course.get('LEC_Schedule', 'N/A'),  # Lecture Schedule
            'LaS': matched_course.get('LAB_Schedule', 'N/A')   # Laboratory Schedule
        }
    else:
        # If no match found, we should still have basic placeholders to avoid leaving [Insert *] in the document
        # However, we'll handle this properly in the app by showing a warning
        context = {
            'SY': 'NO MATCH FOUND',
            'SC': 'NO MATCH FOUND',
            'SN': 'NO MATCH FOUND',
            'SEM': 'NO MATCH FOUND',
            'LECTURER': 'NO MATCH FOUND',
            'CREDIT': 'NO MATCH FOUND',
            'SCHED': 'NO MATCH FOUND',
            'TITLE': 'NO MATCH FOUND',
            'STUDENT_COUNT': len(dbf_data_df) if dbf_data_df is not None else 'N/A',
            # Additional mappings for no match case
            'ST': 'NO MATCH FOUND',
            'Sem': 'NO MATCH FOUND',
            'OC': 'NO MATCH FOUND',
            'Time': pd.Timestamp.now().strftime('%Y-%m-%d'),  # Current day only
            'Faculty': 'NO MATCH FOUND'
        }

    # Handle the student data table specifically - look for the table with student information
    # Use the DBF data that was passed as parameter
    student_data = []
    statistics = {}
    if dbf_data_df is not None and not dbf_data_df.empty:
        student_data, statistics = process_student_data(dbf_data_df)

    # Add student data to context
    context['students'] = student_data
    # Add statistics to context
    context.update(statistics)

    # Process the template with docxtpl
    doc = DocxTemplate(template_path)
    doc.render(context)

    # Save to BytesIO and return
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def process_student_data(df):
    """
    Process student data from DataFrame to prepare for docxtpl
    """
    # Debug: Print column names to see what's available
    print(f"DBF DataFrame columns: {list(df.columns)}")

    # Get the DataFrame columns that correspond to student data
    # Look for common column names in the DataFrame
    name_col = find_column_name(df, ['name', 'student_name', 'studname', 'lastname', 'student', 'name_field'])
    grade_col = find_column_name(df, ['grade', 'eg', 'score', 'mark', 'eg_field'])
    remarks_col = find_column_name(df, ['remarks', 'remark', 'comment', 'status', 'remark_field'])

    print(f"Detected name column: {name_col}")
    print(f"Detected grade column: {grade_col}")
    print(f"Detected remarks column: {remarks_col}")

    # Add row numbers column
    df_with_numbers = df.copy()
    df_with_numbers.insert(0, 'row_number', range(1, len(df_with_numbers) + 1))
    number_col = 'row_number'

    # Calculate statistics
    statistics = {}
    if remarks_col and remarks_col in df_with_numbers.columns:
        remarks_values = df_with_numbers[remarks_col].astype(str).str.upper()

        passed_count = sum(1 for value in remarks_values
                          if any(keyword in value for keyword in
                                ['PASSED', 'PASS', 'OK', 'COMPLETED']))
        no_grade_count = sum(1 for value in remarks_values
                            if any(keyword in value for keyword in
                                  ['NO GRADE', 'N/A', 'NO REMARK', 'NONE', 'INC', 'INCOMPLETE']))
        failed_count = sum(1 for value in remarks_values
                          if any(keyword in value for keyword in
                                ['FAILED', 'FAIL']))
        dropped_count = sum(1 for value in remarks_values
                           if any(keyword in value for keyword in
                                 ['DROPPED', 'DROP', 'WITHDRAWN', 'WITHDREW', 'DRP']))
    else:
        passed_count = failed_count = no_grade_count = dropped_count = 0

    total_count = len(df_with_numbers)

    statistics['STATISTICS_SUMMARY'] = f"STATISTICS   :   Passed={passed_count}  No Grade={no_grade_count}  Failed={failed_count}  Dropped={dropped_count}  TOTAL={total_count}"

    # Create a list of dictionaries for the template
    student_data = []
    for idx, row in df_with_numbers.iterrows():
        number_val = str(row[number_col]) if number_col in row else str(idx + 1)
        name_val = str(row[name_col]) if name_col and name_col in row else ""
        grade_val = str(row[grade_col]) if grade_col and grade_col in row else ""
        remarks_val = str(row[remarks_col]) if remarks_col and remarks_col in row else ""

        student_record = {
            'number': number_val,
            'name': name_val,
            'grade': grade_val,
            'remarks': remarks_val,
            # Also add a 'student' key that contains the same data for {{student.number}} format
            'student': {
                'number': number_val,
                'name': name_val,
                'grade': grade_val,
                'remarks': remarks_val
            }
        }
        print(f"Student record: {student_record}")  # Debug print
        student_data.append(student_record)

    print(f"Total student records created: {len(student_data)}")  # Debug print
    print(f"Statistics: {statistics}")  # Debug print

    return student_data, statistics


def find_column_name(df, possible_names):
    """
    Find a column in the DataFrame based on possible names (case-insensitive)
    """
    print(f"Looking for columns from possible names: {possible_names}")
    print(f"Available columns in DataFrame: {list(df.columns)}")

    df_columns = [col.lower() for col in df.columns]
    for name in possible_names:
        for col in df.columns:
            if name.lower() in col.lower():
                print(f"Found matching column for '{name}': '{col}'")
                return col
    # If no exact match, return the first column that seems appropriate
    for col in df.columns:
        for name in possible_names:
            if name.lower() in col.lower():
                print(f"Found matching column for '{name}': '{col}'")
                return col
    # If still no match, return the first column as a fallback
    print("No matching column found, using first column as fallback")
    return df.columns[0] if len(df.columns) > 0 else None


def generate_word_report_from_jle_dbf(jle_file_path, template_path="Report_template.docx", dbf_directory="testfiles"):
    """
    Generate a Word report from the matched JLE and DBF files using docxtpl
    """
    return populate_template_with_jle_dbf_data(template_path, jle_file_path, dbf_directory)


def generate_word_report_from_jle_data(jle_data, template_path="Report_template.docx", dbf_directory="testfiles"):
    """
    Generate a Word report from the matched JLE data and DBF files using docxtpl
    """
    return populate_template_with_jle_dbf_data_from_jle_data(template_path, jle_data, dbf_directory)


def generate_word_report_from_jle_and_uploaded_dbf(jle_data, uploaded_dbf_filename, dbf_data_df, template_path="Report_template.docx"):
    """
    Generate a Word report from JLE data and uploaded DBF file information using docxtpl
    """
    return populate_template_with_jle_and_uploaded_dbf_data(template_path, jle_data, uploaded_dbf_filename, dbf_data_df)


def generate_word_report(df, jle_data, template_path="Report_template.docx"):
    """
    Generate a Word report from the DataFrame and JLE data using docxtpl
    """
    # Prepare context from JLE data (old method)
    context = {}

    # Extract useful information from JLE data
    if jle_data:
        context.update({
            'JLE_FILENAME': jle_data.get('filename', 'N/A'),
            'JLE_SIZE': jle_data.get('size', 'N/A'),
            'DATE_GENERATED': pd.Timestamp.now().strftime('%Y-%m-%d'),
            'TIME_GENERATED': pd.Timestamp.now().strftime('%H:%M:%S'),  # Keep time for this specific field
            'TOTAL_COURSES': jle_data.get('total_courses', 'N/A')
        })

        # If there is course data from the JLE file, extract course information
        if 'course_data' in jle_data and jle_data['course_data'] is not None and not jle_data['course_data'].empty:
            course_df = jle_data['course_data']

            # Extract course information for reporting
            course_codes = ', '.join(course_df['Subject Code'].unique()) if 'Subject Code' in course_df.columns else 'N/A'
            lecturers = ', '.join(course_df['Lecturer'].dropna().unique()) if 'Lecturer' in course_df.columns else 'N/A'

            context.update({
                'COURSE_CODES': course_codes,
                'LECTURERS': lecturers,
                'COURSE_COUNT': len(course_df)
            })

            # Add first course details as examples
            if len(course_df) > 0:
                first_course = course_df.iloc[0]
                context.update({
                    'FIRST_COURSE_CODE': first_course.get('Subject Code', 'N/A'),
                    'FIRST_COURSE_TITLE': first_course.get('Subject Title', 'N/A'),
                    'FIRST_COURSE_SCHEDULE': first_course.get('Schedule', 'N/A'),
                    'FIRST_COURSE_CREDIT': first_course.get('Credit', 'N/A'),
                    'FIRST_COURSE_LECTURER': first_course.get('Lecturer', 'N/A')
                })

                # Add the new mappings based on the first course if available
                context.update({
                    'ST': first_course.get('Subject Title', 'N/A'),  # Subject Title
                    'Sem': first_course.get('Semester', 'N/A'),      # Semester
                    'OC': first_course.get('Subject Num', 'N/A'),    # Subject Num
                    'Time': pd.Timestamp.now().strftime('%Y-%m-%d'),  # Current day only
                    'Faculty': first_course.get('Lecturer', 'N/A'),   # Lecturer
                    'SCHED': first_course.get('Schedule', 'N/A'),    # All schedules
                    'LeS': first_course.get('LEC_Schedule', 'N/A'),  # Lecture Schedule
                    'LaS': first_course.get('LAB_Schedule', 'N/A')   # Laboratory Schedule
                })
            else:
                # Add default values for the new mappings if no course data
                context.update({
                    'ST': 'N/A',  # Subject Title
                    'Sem': 'N/A',      # Semester
                    'OC': 'N/A',    # Subject Num
                    'Time': pd.Timestamp.now().strftime('%Y-%m-%d'),  # Current day only
                    'Faculty': 'N/A',   # Lecturer
                    'SCHED': 'N/A',    # All schedules
                    'LeS': 'N/A',  # Lecture Schedule
                    'LaS': 'N/A'   # Laboratory Schedule
                })

        # Add dataframe statistics
        if df is not None and not df.empty:
            context.update({
                'TOTAL_RECORDS': len(df),
                'COLUMNS_COUNT': len(df.columns),
                'COLUMN_NAMES': ', '.join(df.columns.tolist())
            })

    # Handle the student data table specifically - look for the table with student information
    student_data = []
    statistics = {}
    if df is not None and not df.empty:
        student_data, statistics = process_student_data(df)

    # Add student data to context
    context['students'] = student_data
    # Add statistics to context
    context.update(statistics)

    # Process the template with docxtpl
    doc = DocxTemplate(template_path)
    doc.render(context)

    # Save to BytesIO and return
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def show_word_report_ui(df, jle_data):
    """
    Display the Word report UI
    """
    # Initialize session state variables if they don't exist
    if 'word_report_generated' not in st.session_state:
        st.session_state.word_report_generated = False

    st.write("Click the button below to generate a Word report using the template:")

    # Generate Word Report button
    if st.button("Generate Word Report", key="generate_word_report_btn"):
        with st.spinner('Generating Word report...'):
            try:
                # Generate Word report
                word_bytes = generate_word_report(df, jle_data)

                # Store the Word doc in session state
                st.session_state.word_bytes = word_bytes
                st.session_state.word_filename = f"E-Class_Report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.docx"
                st.session_state.word_report_generated = True

                st.success("Word report generated successfully!")

            except Exception as e:
                st.error(f"Error generating Word report: {str(e)}")

    # Show the download button if Word report is generated
    if st.session_state.word_report_generated and 'word_bytes' in st.session_state:
        st.download_button(
            label="Download Word Report",
            data=st.session_state.word_bytes,
            file_name=st.session_state.word_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="download_word_btn"
        )


def show_jle_dbf_report_ui(jle_file_path):
    """
    Display the Word report UI for JLE/DBF workflow
    """
    # Initialize session state variables if they don't exist
    if 'jle_dbf_word_report_generated' not in st.session_state:
        st.session_state.jle_dbf_word_report_generated = False

    st.write("Click the button below to generate a Word report using the JLE/DBF template:")

    # Generate Word Report button for JLE/DBF workflow
    if st.button("Generate Word Report (JLE/DBF)", key="generate_jle_dbf_report_btn"):
        with st.spinner('Generating Word report from JLE/DBF...'):
            try:
                # Generate Word report using JLE/DBF data
                word_bytes = generate_word_report_from_jle_dbf(jle_file_path)

                # Store the Word doc in session state
                st.session_state.jle_dbf_word_bytes = word_bytes
                st.session_state.jle_dbf_word_filename = f"E-Class_Report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.docx"
                st.session_state.jle_dbf_word_report_generated = True

                st.success("Word report generated successfully from JLE/DBF data!")

            except Exception as e:
                st.error(f"Error generating Word report from JLE/DBF: {str(e)}")

    # Show the download button if Word report is generated
    if st.session_state.jle_dbf_word_report_generated and 'jle_dbf_word_bytes' in st.session_state:
        st.download_button(
            label="Download Word Report (JLE/DBF)",
            data=st.session_state.jle_dbf_word_bytes,
            file_name=st.session_state.jle_dbf_word_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="download_jle_dbf_word_btn"
        )