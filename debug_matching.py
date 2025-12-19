import tkinter as tk  
from tkinter import filedialog, messagebox  
from openpyxl import load_workbook  
from dbf import Table, READ_WRITE  
  
# --- Helper function to clean numeric values ---  
def clean_value(val):  
    if val is None:  
        return ''  
    if isinstance(val, (int, float)):  
        return f"{val:.1f}"  
    return val  
  
def debug_process_files():  
    excel_path = r"d:\Python\ECLASSRECORD2DBF WEB\testfiles\2506B.xlsm"  
    dbf_path = r"d:\Python\ECLASSRECORD2DBF WEB\testfiles\DSO_20243_2506B_BACC104_565.DBF" 
  
    if not excel_path or not dbf_path:  
        print("Please select both files.")  
        return  
  
    try:  
        # Load Excel  
        wb = load_workbook(excel_path, data_only=True)  
        print(f"Excel sheets: {wb.sheetnames}")  
  
        ws = wb["FFG"]  
        print(f"Using worksheet: FFG")  
  
        row = 11  
        excel_data = {}  
        print("Reading Excel data from column C ^(ID^) and columns H, I ^(values^):")  
  
        while True:  
            cell_val = ws[f'C{row}'].value  
            if cell_val is None:  
                print(f"Stopping at row {row} - cell C{row} is empty")  
                break  
  
            print(f"Row {row}, C{row} = {cell_val}, type: {type(cell_val)}")  
  
            try:  
                # This is the current approach in your code - converting to int  
                id_str = str(int(cell_val)).strip()  
                print(f"  Converted ID: {id_str}")  
            except Exception as e:  
                print(f"  Skipping row {row} - conversion failed: {e}")  
                row += 1  
                continue  
  
            col_h = ws[f'H{row}'].value  
            col_i = ws[f'I{row}'].value  
            print(f"  H{row} = {col_h}, I{row} = {col_i}")  
            excel_data[id_str] = (col_h, col_i)  
            row += 1  
  
            # Just process first 5 non-empty rows for debugging
            if len(excel_data) >= 5:
                break
  
        print(f"Excel data collected: {excel_data}")  
        wb.close()  
  
        # Open DBF for read/write  
        table = Table(dbf_path)  
        table.open(mode=READ_WRITE) 
  
        print(f"DBF field names: {table.field_names}")  
        id_index = 5  # 5th column, 0-based index  
        print(f"Checking field at index {id_index}: {table.field_names[id_index] if id_index < len(table.field_names) else 'FIELD INDEX OUT OF RANGE'}")  
  
        target_col3 = 2  # 3rd column (write H)  
        target_col4 = 3  # 4th column (write I)  
  
        print(f"Target columns: {table.field_names[target_col3]} (index {target_col3}) and {table.field_names[target_col4]} (index {target_col4})")  
  
        matched = 0  
        processed = 0  
  
        for i, record in enumerate(table):  
            processed += 1  
            dbf_id = str(record[id_index]).strip()  
            print(f"DBF record {i+1}, ID at field {id_index}: '{dbf_id}' (type: {type(record[id_index])})")  
  
            if dbf_id in excel_data:  
                h_val, i_val = excel_data[dbf_id]  
                print(f"  MATCH found! DBF ID '{dbf_id}' matches Excel data: H={h_val}, I={i_val}")  
                matched += 1  
                # Don't actually modify for this debug
                if matched >= 5:  # Just check first few matches
                    break
            else:  
                print(f"  No match for DBF ID '{dbf_id}' in Excel data")  
  
            if processed >= 10:  # Just check first 10 records
                break
  
        table.close()  
        print(f"Total DBF records processed: {processed}, matches found: {matched}")  
  
    except Exception as e:  
        print(f"Error: {str(e)}")  
        import traceback  
        traceback.print_exc()  
  
if __name__ == "__main__":  
    debug_process_files() 
