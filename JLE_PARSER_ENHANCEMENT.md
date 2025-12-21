# JLE Parser Enhancement - Documentation

## Overview
The JLE parser has been enhanced to extract academic year and semester information from the filename pattern. This enhancement was added to the `config.py` file and integrated into the existing application.

## Changes Made

### 1. Created `config.py` file
- Added `parse_jle_with_filename(file_path)` function that:
  - Extracts semester/year from the filename pattern (e.g., "DSO_20243_565.JLE")
  - Maintains all original content parsing functionality
  - Adds "Academic Year" and "Semester" columns to each record

- Added enhanced `extract_jle_data(jle_file)` function that:
  - Works with uploaded file objects (like in Streamlit)
  - Extracts filename-based metadata
  - Maintains all original functionality
  - Returns additional metadata in the info dictionary

### 2. Updated `app.py` file
- Added import for the new `extract_jle_data` function from config.py
- Removed duplicate function definition to avoid redundancy

## Filename Pattern Recognition
- Pattern: `YYYYX` where `YYYY` is the year and `X` is the semester indicator
- Semester mapping:
  - `1` → "1st Semester"
  - `2` → "2nd Semester" 
  - `3` → "Summer"
  - Other values → "{X}th Term"

## Example
For filename "DSO_20243_565.JLE":
- Academic Year: "2024-2025"
- Semester: "Summer"

## Usage

### For file paths:
```python
from config import parse_jle_with_filename
df = parse_jle_with_filename('/path/to/your/file.JLE')
```

### For uploaded file objects (Streamlit):
```python
from config import extract_jle_data
jle_info = extract_jle_data(uploaded_file)
```

## Benefits
- Automatically extracts academic term information from filenames
- Maintains backward compatibility with existing functionality
- Provides richer metadata for reports and analysis
- Centralized parsing logic in config.py for better maintainability