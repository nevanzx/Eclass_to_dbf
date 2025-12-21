import os
import pandas as pd
import re


def parse_jle_with_filename(file_path):
    """
    Parse a JLE file extracting both the content data and metadata from the filename.
    
    Args:
        file_path: Path to the JLE file
    
    Returns:
        pd.DataFrame: DataFrame containing parsed course records with added metadata
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
    # Looks for 4 digits (Subject Num) + space + Alphanumeric (Subject Code)
    # The first letter of the subject code is part of the subject number
    # Example: "2506   FBACC104" should be interpreted as Subject Num: 2506F, Subject Code: BACC104
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
        # So FBACC104 becomes Subject Num: 2506F and Subject Code: BACC104
        subj_num_with_letter = subj_num + full_subj_code[0]  # 2506 + F = 2506F
        real_subj_code = full_subj_code[1:]  # BACC104 (everything after the first letter)

        # Remove the header (Num + Code) from the chunk to process the rest
        # We also replace newlines with spaces to make regex easier
        rest_of_text = chunk[len(matches[i].group(0)):].replace('\n', ' ').strip()

        # 2. Extract Credit and Lecturer - improved pattern to handle various formats
        # Look for digit followed by lecturer name at the end of the text
        # This pattern looks for a digit followed by whitespace and then the lecturer name
        lecturer_pattern = re.search(r"(\d)\s+([A-Z][A-Z\s\.\-]*[A-Z])(?:\s*$|[\x00-\x1f\x7f\s]*$)", rest_of_text)

        credit = None
        lecturer = None

        if lecturer_pattern:
            credit = lecturer_pattern.group(1)
            lecturer = lecturer_pattern.group(2).strip()
            # Remove this part from the text so we can find the Title/Schedule
            rest_of_text = rest_of_text[:lecturer_pattern.start()].strip()
        else:
            # Fallback: look for any digit followed by potential lecturer name
            # Pattern: single digit followed by name-like pattern
            alt_pattern = re.search(r"(\d)\s+([A-Z][A-Z\s\.\-]+?)(?:\s+(?:\d\s+)?[A-Z]|$)", rest_of_text)
            if alt_pattern:
                credit = alt_pattern.group(1)
                lecturer = alt_pattern.group(2).strip()
                rest_of_text = rest_of_text[:alt_pattern.start()].strip()

        # 3. Extract Schedule(s) - distinguish between LEC and LAB
        # Pattern: Time (e.g., 730AM-1200PM), Day (e.g., SuSa), Room (e.g., B63)
        # Enhanced to potentially identify LEC/LAB indicators if present in the format
        sched_pattern_regex = r"(\d{1,4}[AP]M-\s*\d{1,4}[AP]M\s+[A-Za-z]+\s+[A-Z0-9]+)"
        schedules_found = re.findall(sched_pattern_regex, rest_of_text)

        # For now, we'll treat all schedules as general schedules
        # In a more advanced implementation, we could identify LEC/LAB based on context
        all_schedule_str = " / ".join([s.strip() for s in schedules_found])

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

        # Ensure blank values when no schedules exist
        lec_schedule_str = lec_schedule_str if lec_schedule_str else ""
        lab_schedule_str = lab_schedule_str if lab_schedule_str else ""

        # 4. Extract Subject Title
        # The title is whatever is left after removing the schedules
        # We use re.sub to remove the schedules we found
        title_clean = re.sub(sched_pattern_regex, '', rest_of_text)
        # Also remove the lecturer and credit from the title if they weren't removed by the main pattern
        if credit and lecturer:
            lecturer_part = f"{credit}\\s+{lecturer}"
            title_clean = re.sub(lecturer_part, '', title_clean, flags=re.IGNORECASE)
        title_clean = " ".join(title_clean.split()) # Remove extra whitespace

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


def extract_jle_data(jle_file):
    """
    Extract data from JLE file. The JLE format contains academic course information
    with structured data including subject numbers, codes, titles, schedules, credits, and lecturers.
    This function extracts content from an uploaded file object and adds metadata from the filename.
    """
    import re
    import pandas as pd

    # Read the JLE file content as binary and decode with latin1 encoding
    jle_bytes = jle_file.getvalue()
    raw_data = jle_bytes.decode('latin1', errors='ignore')

    # Extract metadata from filename
    filename = jle_file.name
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

    # Define the "Start of Record" pattern
    # Looks for 4 digits (Subject Num) + space + Alphanumeric (Subject Code)
    # The first letter of the subject code is part of the subject number
    # Example: "2506   FBACC104" should be interpreted as Subject Num: 2506F, Subject Code: BACC104
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
        # So FBACC104 becomes Subject Num: 2506F and Subject Code: BACC104
        subj_num_with_letter = subj_num + full_subj_code[0]  # 2506 + F = 2506F
        real_subj_code = full_subj_code[1:]  # BACC104 (everything after the first letter)

        # Remove the header (Num + Code) from the chunk to process the rest
        # We also replace newlines with spaces to make regex easier
        rest_of_text = chunk[len(matches[i].group(0)):].replace('\n', ' ').strip()

        # 2. Extract Credit and Lecturer - improved pattern to handle various formats
        # Look for digit followed by lecturer name at the end of the text
        # This pattern looks for a digit followed by whitespace and then the lecturer name
        lecturer_pattern = re.search(r"(\d)\s+([A-Z][A-Z\s\.\-]*[A-Z])(?:\s*$|[\x00-\x1f\x7f\s]*$)", rest_of_text)

        credit = None
        lecturer = None

        if lecturer_pattern:
            credit = lecturer_pattern.group(1)
            lecturer = lecturer_pattern.group(2).strip()
            # Remove this part from the text so we can find the Title/Schedule
            rest_of_text = rest_of_text[:lecturer_pattern.start()].strip()
        else:
            # Fallback: look for any digit followed by potential lecturer name
            # Pattern: single digit followed by name-like pattern
            alt_pattern = re.search(r"(\d)\s+([A-Z][A-Z\s\.\-]+?)(?:\s+(?:\d\s+)?[A-Z]|$)", rest_of_text)
            if alt_pattern:
                credit = alt_pattern.group(1)
                lecturer = alt_pattern.group(2).strip()
                rest_of_text = rest_of_text[:alt_pattern.start()].strip()

        # 3. Extract Schedule(s) - distinguish between LEC and LAB
        # Pattern: Time (e.g., 730AM-1200PM), Day (e.g., SuSa), Room (e.g., B63)
        # Enhanced to potentially identify LEC/LAB indicators if present in the format
        sched_pattern_regex = r"(\d{1,4}[AP]M-\s*\d{1,4}[AP]M\s+[A-Za-z]+\s+[A-Z0-9]+)"
        schedules_found = re.findall(sched_pattern_regex, rest_of_text)

        # For now, we'll treat all schedules as general schedules
        # In a more advanced implementation, we could identify LEC/LAB based on context
        all_schedule_str = " / ".join([s.strip() for s in schedules_found])

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

        # Ensure blank values when no schedules exist
        lec_schedule_str = lec_schedule_str if lec_schedule_str else ""
        lab_schedule_str = lab_schedule_str if lab_schedule_str else ""

        # 4. Extract Subject Title
        # The title is whatever is left after removing the schedules
        # We use re.sub to remove the schedules we found
        title_clean = re.sub(sched_pattern_regex, '', rest_of_text)
        # Also remove the lecturer and credit from the title if they weren't removed by the main pattern
        if credit and lecturer:
            lecturer_part = f"{credit}\\s+{lecturer}"
            title_clean = re.sub(lecturer_part, '', title_clean, flags=re.IGNORECASE)
        title_clean = " ".join(title_clean.split()) # Remove extra whitespace

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

    # Prepare the info to return
    jle_info = {
        'raw_bytes': jle_bytes,
        'size': len(jle_bytes),
        'filename': jle_file.name,
        'course_data': jle_df,
        'total_courses': len(parsed_records),
        'course_codes': [record['Subject Code'] for record in parsed_records] if parsed_records else [],
        'academic_year': academic_year,
        'semester': semester
    }

    return jle_info