"""
Student Scores Formatter
Transforms Excel data from horizontal subject columns to vertical format
with Year 1 and Year 2 scores side by side.
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import os


def read_excel_sheets(file_path):
    """Read YEAR 1 and YEAR 2 sheets from the Excel workbook."""
    year1_df = pd.read_excel(file_path, sheet_name='YEAR 1')
    year2_df = pd.read_excel(file_path, sheet_name='YEAR 2')
    return year1_df, year2_df


def get_subject_columns(df):
    """Get the subject column names (everything after PROGRAMME)."""
    base_columns = ['S/N', 'ADMISSION NUMBER', 'STUDENT NAME', 'PROGRAMME']
    subject_columns = [col for col in df.columns if col not in base_columns]
    return subject_columns


def transform_student_data(year1_df, year2_df):
    """Transform the data into the required vertical format."""
    # Get all subject columns
    year1_subjects = get_subject_columns(year1_df)
    year2_subjects = get_subject_columns(year2_df)
    
    # Get unique students from both years
    all_students = {}
    
    # Process Year 1 data
    for _, row in year1_df.iterrows():
        admission_no = row['ADMISSION NUMBER']
        if admission_no not in all_students:
            all_students[admission_no] = {
                'name': row['STUDENT NAME'],
                'programme': row['PROGRAMME'],
                'admission_no': admission_no,
                'subjects': {}
            }
        
        for subject in year1_subjects:
            score = row[subject]
            if pd.notna(score) and score != '' and score != 0:
                if subject not in all_students[admission_no]['subjects']:
                    all_students[admission_no]['subjects'][subject] = {'year1': None, 'year2': None}
                all_students[admission_no]['subjects'][subject]['year1'] = score
    
    # Process Year 2 data
    for _, row in year2_df.iterrows():
        admission_no = row['ADMISSION NUMBER']
        if admission_no not in all_students:
            all_students[admission_no] = {
                'name': row['STUDENT NAME'],
                'programme': row['PROGRAMME'],
                'admission_no': admission_no,
                'subjects': {}
            }
        
        for subject in year2_subjects:
            score = row[subject]
            if pd.notna(score) and score != '' and score != 0:
                if subject not in all_students[admission_no]['subjects']:
                    all_students[admission_no]['subjects'][subject] = {'year1': None, 'year2': None}
                all_students[admission_no]['subjects'][subject]['year2'] = score
    
    return all_students


def create_formatted_excel(students_data, output_path):
    """Create the formatted Excel output file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Formatted Scores"
    
    # Define styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    cell_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Write headers
    headers = ['Student Details', 'S/N', 'Subjects', 'Year 1', 'Year 2', 'Year 3']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Set column widths
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 10
    
    current_row = 2
    
    # Alternate row colors for different students
    light_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    alt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    student_index = 0
    for admission_no, student in students_data.items():
        if not student['subjects']:
            continue
            
        subjects = list(student['subjects'].keys())
        num_subjects = len(subjects)
        
        # Determine row fill color (alternate between students)
        row_fill = light_fill if student_index % 2 == 0 else alt_fill
        
        # Create student details text
        student_details = f"{student['name']}\nClass: \nProgramme: {student['programme']}\nIndex No: {admission_no}"
        
        # Merge cells for student details column
        start_row = current_row
        end_row = current_row + num_subjects - 1
        
        if num_subjects > 1:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        
        # Write student details in the merged cell
        details_cell = ws.cell(row=start_row, column=1, value=student_details)
        details_cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        details_cell.border = thin_border
        details_cell.fill = row_fill
        
        # Write each subject row
        for i, subject in enumerate(subjects):
            row = current_row + i
            scores = student['subjects'][subject]
            
            # S/N
            sn_cell = ws.cell(row=row, column=2, value=i + 1)
            sn_cell.alignment = cell_alignment
            sn_cell.border = thin_border
            sn_cell.fill = row_fill
            
            # Subject name
            subj_cell = ws.cell(row=row, column=3, value=subject)
            subj_cell.alignment = left_alignment
            subj_cell.border = thin_border
            subj_cell.fill = row_fill
            
            # Year 1 score
            y1_cell = ws.cell(row=row, column=4, value=scores['year1'] if scores['year1'] else '')
            y1_cell.alignment = cell_alignment
            y1_cell.border = thin_border
            y1_cell.fill = row_fill
            
            # Year 2 score
            y2_cell = ws.cell(row=row, column=5, value=scores['year2'] if scores['year2'] else '')
            y2_cell.alignment = cell_alignment
            y2_cell.border = thin_border
            y2_cell.fill = row_fill
            
            # Year 3 score (empty for now)
            y3_cell = ws.cell(row=row, column=6, value='')
            y3_cell.alignment = cell_alignment
            y3_cell.border = thin_border
            y3_cell.fill = row_fill
            
            # Apply border to student details column for non-merged rows
            if i > 0:
                ws.cell(row=row, column=1).border = thin_border
                ws.cell(row=row, column=1).fill = row_fill
        
        current_row = end_row + 1
        student_index += 1
    
    # Freeze the header row
    ws.freeze_panes = 'A2'
    
    # Save the workbook
    wb.save(output_path)
    print(f"Formatted file saved to: {output_path}")


def main():
    # Get input file path from user
    input_file = input("Enter the path to your Excel file: ").strip()
    
    # Remove quotes if present
    if input_file.startswith('"') and input_file.endswith('"'):
        input_file = input_file[1:-1]
    if input_file.startswith("'") and input_file.endswith("'"):
        input_file = input_file[1:-1]
    
    if not os.path.exists(input_file):
        print(f"Error: File not found: {input_file}")
        return
    
    # Generate output file path
    base_name = os.path.splitext(input_file)[0]
    output_file = f"{base_name}_formatted.xlsx"
    
    print("Reading Excel sheets...")
    try:
        year1_df, year2_df = read_excel_sheets(input_file)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    print("Transforming data...")
    students_data = transform_student_data(year1_df, year2_df)
    
    print(f"Found {len(students_data)} students")
    
    print("Creating formatted output...")
    create_formatted_excel(students_data, output_file)
    
    print("Done!")


if __name__ == "__main__":
    main()
