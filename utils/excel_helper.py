import os
from openpyxl import Workbook, load_workbook

DB_FILE = "students.xlsx"

COLUMNS = [
    "Roll_No", "Register_No", "Student_Name", "Father_Name", "Class", "Course",
    "Tuition_Fee_Year1", "Bus_Fee_Year1", "Total_Year1",
    "Tuition_Fee_Year2", "Bus_Fee_Year2", "Total_Year2",
    "Tuition_Fee_Year3", "Bus_Fee_Year3", "Total_Year3",
    "Tuition_Fee_Year4", "Bus_Fee_Year4", "Total_Year4",
    "Grand_Total"
]

def init_db():
    """Initializes the Excel database with standard columns if it does not exist."""
    if not os.path.exists(DB_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Students_Fees"
        ws.append(COLUMNS)
        wb.save(DB_FILE)

def calculate_row_totals(data: dict) -> dict:
    """Helper to convert inputs to floats, fill blanks with 0, and compute year/grand totals."""
    processed = {}
    
    # Text data extraction
    for field in ["Roll_No", "Register_No", "Student_Name", "Father_Name", "Class", "Course"]:
        processed[field] = str(data.get(field, "")).strip()

    # Numeric data handling & automated calculation
    grand_total = 0.0
    for i in range(1, 5):
        t_key = f"Tuition_Fee_Year{i}"
        b_key = f"Bus_Fee_Year{i}"
        
        # Safely convert inputs to float or throw readable clean errors
        try:
            t_val = float(data.get(t_key) or 0)
        except ValueError:
            raise ValueError(f"Invalid text value entered in Tuition Fee Year {i}. Please insert numbers only.")
            
        try:
            b_val = float(data.get(b_key) or 0)
        except ValueError:
            raise ValueError(f"Invalid text value entered in Bus Fee Year {i}. Please insert numbers only.")
        
        year_total = t_val + b_val
        grand_total += year_total
        
        processed[t_key] = t_val
        processed[f"Bus_Fee_Year{i}"] = b_val
        processed[f"Total_Year{i}"] = year_total

    processed["Grand_Total"] = grand_total
    return processed

def get_all_roll_numbers():
    """Returns a list of all existing roll numbers from the spreadsheet."""
    init_db()
    wb = load_workbook(DB_FILE)
    ws = wb.active
    roll_numbers = []
    
    # Iterate skipping the header
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0] is not None:
            roll_numbers.append(str(row[0]).strip())
    wb.close()
    return roll_numbers

def add_student_record(student_data: dict) -> bool:
    """Appends a new student record to the sheet."""
    init_db()
    processed = calculate_row_totals(student_data)
    
    if processed["Roll_No"] in get_all_roll_numbers():
        return False  # Already exists
        
    wb = load_workbook(DB_FILE)
    ws = wb.active
    
    # Construct row matching column indices
    row_to_add = [processed.get(col) for col in COLUMNS]
    ws.append(row_to_add)
    wb.save(DB_FILE)
    return True

def update_student_record(student_data: dict) -> bool:
    """Finds an existing roll number and replaces its cell data values."""
    init_db()
    processed = calculate_row_totals(student_data)
    target_roll = processed["Roll_No"]
    
    wb = load_workbook(DB_FILE)
    ws = wb.active
    
    row_found = False
    for row_idx in range(2, ws.max_row + 1):
        cell_val = str(ws.cell(row=row_idx, column=1).value).strip()
        if cell_val == target_roll:
            row_found = True
            # Update cells across columns
            for col_idx, col_name in enumerate(COLUMNS, start=1):
                ws.cell(row=row_idx, column=col_idx, value=processed.get(col_name))
            break
            
    if row_found:
        wb.save(DB_FILE)
    else:
        wb.close()
        
    return row_found

def delete_student_by_roll(roll_no: str) -> bool:
    """Searches for a student roll number row and deletes it from the spreadsheet."""
    init_db()
    target_roll = str(roll_no).strip()
    
    wb = load_workbook(DB_FILE)
    ws = wb.active
    
    row_deleted = False
    for row_idx in range(2, ws.max_row + 1):
        cell_val = str(ws.cell(row=row_idx, column=1).value).strip()
        if cell_val == target_roll:
            ws.delete_rows(row_idx, 1)
            row_deleted = True
            break
            
    if row_deleted:
        wb.save(DB_FILE)
    else:
        wb.close()
        
    return row_deleted

def find_student_by_roll(roll_no: str) -> dict:
    """Searches for a student by roll number and returns a clean dictionary."""
    init_db()
    target_roll = str(roll_no).strip()
    
    wb = load_workbook(DB_FILE)
    ws = wb.active
    
    student_dict = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None and str(row[0]).strip() == target_roll:
            student_dict = {}
            for idx, col_name in enumerate(COLUMNS):
                val = row[idx]
                # Format check: replace None/blanks with 0 if numeric column, otherwise text defaults
                if val is None or val == "":
                    student_dict[col_name] = 0 if ("Fee" in col_name or "Total" in col_name) else "0"
                else:
                    student_dict[col_name] = val
            break
            
    wb.close()
    return student_dict