import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import random
import streamlit as st
import io
import base64

def generate_student_data(num_students=20):
    """Generate random student data"""
    classes = ["SS1", "SS2", "SS3"]
    programmes = ["Science", "Arts", "Commercial"]
    
    data = []
    for i in range(1, num_students + 1):
        index_number = f"STU{str(i).zfill(3)}"
        name = f"Student {i}"
        sex = random.choice(["M", "F"])
        class_name = random.choice(classes)
        programme = random.choice(programmes)
        
        # Generate random scores for core and elective subjects
        c_subjects = [random.randint(50, 100) for _ in range(4)]
        e_subjects = [random.randint(50, 100) for _ in range(4)]
        
        data.append([
            class_name, programme, index_number, name, sex,
            c_subjects[0], c_subjects[1], c_subjects[2], c_subjects[3],
            e_subjects[0], e_subjects[1], e_subjects[2], e_subjects[3]
        ])
    
    columns = [
        "Class", "PROGRAMMES", "INDEX NUMBER", "NAME", "Sex",
        "C-SUBJECT 1", "C-SUBJECT 2", "C-SUBJECT 3", "C-SUBJECT 4",
        "E-SUBJECT 1", "E-SUBJECT 2", "E-SUBJECT 3", "E-SUBJECT 4"
    ]
    
    return pd.DataFrame(data, columns=columns)

def create_subject_scores_sheet(min_score=0, max_score=100, num_students=20):
    """Create a sheet with subjects and random scores for three years"""
    # Create a workbook and select active sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Student Scores"
    
    # Define subjects
    core_subjects = ["Mathematics", "English", "Science", "Social Studies"]
    elective_subjects = ["Physics", "Chemistry", "Biology", "Economics"]
    all_subjects = core_subjects + elective_subjects
    
    # Add headers
    ws['A1'] = "Student Details"
    ws['B1'] = "S/N"
    ws['C1'] = "Subjects"
    ws['D1'] = "Year 1"
    ws['E1'] = "Year 2"
    ws['F1'] = "Year 3"
    
    # Style headers
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    # Add subjects with serial numbers
    for i, subject in enumerate(all_subjects, 1):
        row = i + 1
        ws[f'B{row}'] = i
        ws[f'C{row}'] = subject
        
        # Generate random scores for each year
        ws[f'D{row}'] = random.randint(min_score, max_score)
        ws[f'E{row}'] = random.randint(min_score, max_score)
        ws[f'F{row}'] = random.randint(min_score, max_score)
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15
    
    # Add student data to the same workbook
    student_ws = wb.create_sheet(title="Student Data")
    df = generate_student_data(num_students)
    
    # Write headers
    for col_idx, column_name in enumerate(df.columns, 1):
        student_ws.cell(row=1, column=col_idx).value = column_name
        student_ws.cell(row=1, column=col_idx).font = Font(bold=True)
    
    # Write data
    for row_idx, row in enumerate(df.values, 2):
        for col_idx, value in enumerate(row, 1):
            student_ws.cell(row=row_idx, column=col_idx).value = value
    
    return wb

def get_download_link(wb, filename):
    """Generate a download link for the Excel file"""
    virtual_file = io.BytesIO()
    wb.save(virtual_file)
    virtual_file.seek(0)
    b64 = base64.b64encode(virtual_file.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel file</a>'

def main():
    st.set_page_config(page_title="Student Score Generator", layout="wide")
    
    st.title("Student Score Generator")
    st.write("Generate random scores for students across different years")
    
    col1, col2 = st.columns(2)
    
    with col1:
        min_score = st.number_input("Minimum score", min_value=0, max_value=100, value=0)
    
    with col2:
        max_score = st.number_input("Maximum score", min_value=0, max_value=100, value=100)
    
    num_students = st.slider("Number of students", min_value=5, max_value=100, value=20)
    
    if st.button("Generate Excel File"):
        with st.spinner("Generating Excel file..."):
            wb = create_subject_scores_sheet(min_score, max_score, num_students)
            
            # Provide download link
            download_link = get_download_link(wb, "student_scores.xlsx")
            st.markdown(download_link, unsafe_allow_html=True)
            
            # Preview the data
            st.subheader("Preview of Student Scores")
            
            # Create a DataFrame to display a preview
            preview_data = []
            ws = wb["Student Scores"]
            
            # Get headers
            headers = [cell.value for cell in ws[1]]
            
            # Get a few rows of data
            for row in ws.iter_rows(min_row=2, max_row=6):
                preview_data.append([cell.value for cell in row])
            
            st.dataframe(pd.DataFrame(preview_data, columns=headers))
            
            # Preview student data
            st.subheader("Preview of Student Data")
            df = generate_student_data(num_students)
            st.dataframe(df.head())

if __name__ == "__main__":
    main()
