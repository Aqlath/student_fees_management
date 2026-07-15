# Student Fee Management System

A simple and efficient web application built with Python and Flask to manage student profiles and track academic fee breakdowns across up to 4 academic years. The application uses an Excel spreadsheet as its storage database, meaning it requires zero complex database configuration to set up and run.

---

## Features
- **Auto-Initializing Database:** Automatically checks for and builds the storage Excel spreadsheet with the correct structure at system boot.
- **Form Memory:** Fields stay "sticky" and retain user inputs if validation alerts trigger during data entry.
- **Dynamic Year Sections:** Add up to 4 years of financial breakdowns seamlessly using an interactive interface toggle.
- **Submission Protection:** Normal keyboard "Enter" submissions are blocked on the data entry form to prevent accidental half-filled student enrollments.

---

## Project Structure

```text
📁 student-fee-management/
│
├── 📁 static/
│   └── style.css            # Application styling and custom layouts
│
├── 📁 templates/
│   ├── index.html           # Main user landing dashboard 
│   └── add_student.html     # Data entry, update, and deletion panel
│
├── 📁 utils/
│   └── excel_helper.py      # Core logic handling data matrix interactions
│
├── main.py                  # Main Flask application initialization routing script
└── requirements.txt         # Required Python library configurations


<h2>Installation & Setup</h2>
Follow these steps to install the dependencies and run the application locally on your computer:

1. Prerequisites
Ensure you have Python 3.x installed on your machine.

2. Set Up a Virtual Environment (Recommended)
Open your terminal or command prompt inside the project directory and run:

On Windows:

Bash
python -m venv venv
venv\Scripts\activate
On Mac/Linux:

Bash
python3 -m venv venv
source venv/bin/activate
3. Install Dependencies
Install all the required Python packages specified in the requirements profile:

Bash
pip install -r requirements.txt
4. Run the Application
Execute the main application file to fire up the system server:

Bash
python main.py
Once running, open your web browser and navigate to:

Plaintext
[http://127.0.0.1:5000/](http://127.0.0.1:5000/)
