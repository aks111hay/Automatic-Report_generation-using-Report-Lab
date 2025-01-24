# README

## Project Overview
This project is designed to generate personalized report cards for students using data provided in an Excel file. The project utilizes the **ReportLab** library for creating PDF files and **Pandas** for data manipulation and processing. Each report card includes the following:
- Student ID
- Student Name
- Subject-Wise Marks
- Total Marks
- Average Marks
- A Pass/Fail Status based on the marks

The input Excel file should have the following columns:
- **Student ID**: Unique identifier for each student
- **Name**: Student's name
- **Subject**: The name of the subject
- **Score**: Marks scored by the student in that subject

The application reads the Excel file, processes the data, and generates individual PDF report cards for each student in the format `report_card_<ID>.pdf`.

## Prerequisites
Make sure you have the following installed:
- **Python 3.8 or later**
- **pip** (Python's package installer)

## Setup Instructions
Follow these steps to set up and run the project:

### Step 1: Download this Folder of  Project
Organized as:
```
invest4edu/
├── README.md
├── requirements.txt
├── student_scores.xlsx
├── report_card.py
├── report_card_1.pdf
├── .....
```

### Step 2: Create a Virtual Environment
It's recommended to use a virtual environment to avoid dependency conflicts. To create and activate a virtual environment:

On Windows:
```bash
python -m venv venv
venv\Scripts\activate
```

On macOS/Linux:
```bash
python3 -m venv venv
source venv/bin/activate
```

### Step 3: Install Dependencies
Install all necessary Python libraries listed in the `requirements.txt` file:
```bash
pip install -r requirements.txt
```

### Step 4: Provide the Input Excel File
Place the Excel file containing student data (e.g., `student_scores.xlsx`) in the root directory of the project. The file must have the following columns:
- **Student ID**: Unique identifier for each student
- **Name**: Student's name
- **Subject**: The name of the subject
- **Score**: Marks scored by the student in that subject

### Step 5: Run the Script
Execute the script to generate report cards:
```bash
python report_card.py
```

### Step 6: View the Generated PDFs
The generated report cards will be saved in the  directory. Each file will be named `report_card_<ID>.pdf`.

### Step 7: Sample Output
For convenience, i have attached some sample output and sample excel files.

## Additional Notes
1. Ensure the Excel file is properly formatted with valid data before running the script.
2. If the script encounters missing or invalid data, it will handle it gracefully and log warnings or in output.
3. The project structure is modular, allowing you to customize the template and logic for report card generation easily.

## Key Libraries Used
- **ReportLab**: Used to design and generate PDF report cards.
- **Pandas**: Used for reading and processing the student data from the Excel file.

## Example Input File Format (`student_scores.xlsx`)
| Studdent ID   | Name      | Subject      | Score |
|---------------|-----------|--------------|-------|
| 101           | Alice     | Math         | 85    |
| 101           | Alice     | English      | 78    |
| 102           | Bob       | Science      | 90    |
| 103           | Charlie   | History      | 65    |

