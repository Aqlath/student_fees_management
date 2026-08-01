from fastapi import FastAPI, Request, Form, Query
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from urllib.parse import quote_plus
import re
from utils.excel_helper import (
    add_student_record, 
    update_student_record, 
    find_student_by_roll,
    delete_student_by_roll,
    ensure_excel_dir
)

app = FastAPI(title="Student Expenditure Management System")

ensure_excel_dir()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def home_page(request: Request):
    return templates.TemplateResponse(request=request, name="home.html")

@app.get("/api/student/{roll_no}")
async def get_student_api(roll_no: str):
    student_info = find_student_by_roll(roll_no)
    if student_info:
        return JSONResponse(content={"exists": True, "data": student_info})
    return JSONResponse(content={"exists": False})

@app.get("/add-student", response_class=HTMLResponse)
async def add_student_view(request: Request, status: str = None, message: str = None):
    return templates.TemplateResponse(
        request=request, 
        name="add_student.html", 
        context={"status": status, "message": message, "student": {}}
    )

@app.post("/add-student/submit")
async def handle_add_or_update(
    request: Request,
    action: str = Form(None),
    Roll_No: str = Form(...),
    Register_No: str = Form(...),
    Student_Name: str = Form(...),
    Father_Name: str = Form(""),
    Class: str = Form(...),
    Course: str = Form(...),
    Academic_Year: str = Form(""),
    Tuition_Fee_Year1: str = Form("0"), Bus_Fee_Year1: str = Form("0"),
    Tuition_Fee_Year2: str = Form("0"), Bus_Fee_Year2: str = Form("0"),
    Tuition_Fee_Year3: str = Form("0"), Bus_Fee_Year3: str = Form("0"),
    Tuition_Fee_Year4: str = Form("0"), Bus_Fee_Year4: str = Form("0"),
):
    submitted_student_data = {
        "Roll_No": Roll_No, "Register_No": Register_No, "Student_Name": Student_Name,
        "Father_Name": Father_Name, "Class": Class, "Course": Course, "Academic_Year": Academic_Year,
        "Tuition_Fee_Year1": Tuition_Fee_Year1, "Bus_Fee_Year1": Bus_Fee_Year1,
        "Tuition_Fee_Year2": Tuition_Fee_Year2, "Bus_Fee_Year2": Bus_Fee_Year2,
        "Tuition_Fee_Year3": Tuition_Fee_Year3, "Bus_Fee_Year3": Bus_Fee_Year3,
        "Tuition_Fee_Year4": Tuition_Fee_Year4, "Bus_Fee_Year4": Bus_Fee_Year4,
    }

    def reject_with_data(error_msg: str):
        return templates.TemplateResponse(
            request=request,
            name="add_student.html",
            context={
                "status": "error",
                "message": error_msg,
                "student": submitted_student_data
            }
        )

    try:
        if not action:
            return reject_with_data("Submission blocked. Please explicitly click 'Add Student' or 'Update Record'.")

        if not Roll_No.strip().isalnum():
            return reject_with_data("Roll Number must contain alphanumeric characters only.")
            
        if not Register_No.strip().isalnum():
            return reject_with_data("Register Number must contain alphanumeric characters only.")
        
        if not re.match(r"^[A-Za-z\s\.\-\']+$", Student_Name):
            return reject_with_data("Student Name must contain valid text characters only.")

        fee_inputs = [
            ("Tuition Fee Year 1", Tuition_Fee_Year1), ("Bus Fee Year 1", Bus_Fee_Year1),
            ("Tuition Fee Year 2", Tuition_Fee_Year2), ("Bus Fee Year 2", Bus_Fee_Year2),
            ("Tuition Fee Year 3", Tuition_Fee_Year3), ("Bus Fee Year 3", Bus_Fee_Year3),
            ("Tuition Fee Year 4", Tuition_Fee_Year4), ("Bus Fee Year 4", Bus_Fee_Year4),
        ]
        
        for fee_label, fee_value in fee_inputs:
            if not fee_value.strip().replace('.', '', 1).isdigit():
                return reject_with_data(f"{fee_label} must be a valid number.")

        form_data = {k: v.strip() for k, v in submitted_student_data.items()}

        if action == "add":
            success = add_student_record(form_data)
            if success:
                status, msg = "success", "Student added successfully."
                return RedirectResponse(url=f"/add-student?status={status}&message={quote_plus(msg)}", status_code=303)
            else:
                return reject_with_data("Student already exists with this Roll Number. Click Update Record.")
                
        elif action == "update":
            success = update_student_record(form_data)
            if success:
                status, msg = "success", "Student updated successfully."
                return RedirectResponse(url=f"/add-student?status={status}&message={quote_plus(msg)}", status_code=303)
            else:
                return reject_with_data("Student doesn't exist. Please add the student first.")

    except ValueError as val_err:
        return reject_with_data(f"Data format error: {str(val_err)}")
    except Exception as e:
        return reject_with_data(f"An unexpected error occurred: {str(e)}")

@app.post("/delete-student")
async def handle_delete_student(request: Request, del_roll_no: str = Form(...)):
    try:
        if not del_roll_no.strip():
            status, msg = "error", "Roll number cannot be empty."
        else:
            success = delete_student_by_roll(del_roll_no.strip())
            if success:
                status, msg = "success", f"Student with Roll No '{del_roll_no}' deleted successfully."
            else:
                status, msg = "error", "Student doesn't exist."
    except Exception as e:
        status, msg = "error", f"Error during deletion: {str(e)}"

    return RedirectResponse(url=f"/add-student?status={status}&message={quote_plus(msg)}", status_code=303)

@app.get("/show-student", response_class=HTMLResponse)
async def show_student_view(request: Request, roll_search: str = Query(None), status: str = None, message: str = None):
    student_info = None
    search_triggered = False
    
    if roll_search:
        search_triggered = True
        student_info = find_student_by_roll(roll_search.strip())
        
    return templates.TemplateResponse(
        request=request, 
        name="show_student.html", 
        context={
            "student": student_info, 
            "searched": search_triggered, 
            "query": roll_search,
            "status": status,
            "message": message
        }
    )