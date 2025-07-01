import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
import io
import base64
from fpdf import FPDF

def create_class_list_pdf(arrangement_df, exam_date):
    """Generates a PDF class list for signing."""
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    
    rooms = sorted(arrangement_df['Room'].unique())
    
    for room in rooms:
        pdf.add_page()
        
        # Set title
        pdf.set_font('Arial', 'B', 16)
        title_text = f"Class List for {room}"
        pdf.cell(0, 10, title_text, 0, 1, 'C')
        
        exam_date_text = f"Exam Date: {exam_date.strftime('%A, %B %d, %Y')}"
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, exam_date_text, 0, 1, 'C')
        pdf.ln(5)
        
        # Set table headers
        pdf.set_font('Arial', 'B', 10)
        col_widths = [25, 30, 60, 30, 40] # Column widths
        headers = ['Seat Number', 'Index Number', 'Full Name', 'Class', 'Signature']
        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], 10, header, 1, 0, 'C')
        pdf.ln()
        
        # Add table rows
        pdf.set_font('Arial', '', 9)
        room_df = arrangement_df[arrangement_df['Room'] == room].sort_values('Seat Number')
        for _, row in room_df.iterrows():
            pdf.cell(col_widths[0], 10, str(row['Seat Number']), 1)
            pdf.cell(col_widths[1], 10, str(row['Index Number']), 1)
            # Use multi_cell for names to wrap if they are too long
            x_before_name = pdf.get_x()
            y_before_name = pdf.get_y()
            pdf.multi_cell(col_widths[2], 10, str(row['Full Name']), 1, 'L')
            # Reset position to the next cell
            pdf.set_xy(x_before_name + col_widths[2], y_before_name)
            pdf.cell(col_widths[3], 10, str(row['Class']), 1)
            pdf.cell(col_widths[4], 10, '', 1) # Empty signature cell
            pdf.ln()
            
    return pdf.output(dest='S')

def create_pdf(arrangement_df, exam_date):
    """Generates a PDF file from the arrangement dataframe."""
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    # Set title
    pdf.set_font('Arial', 'B', 16)
    title_text = f"Seating Arrangement for {exam_date.strftime('%A, %B %d, %Y')}"
    pdf.cell(0, 10, title_text, 0, 1, 'C')
    
    # Set table headers
    pdf.set_font('Arial', 'B', 10)
    col_widths = [35, 20, 35, 55, 35, 30, 30] # Column widths
    headers = ['Room', 'Seat Number', 'Index Number', 'Full Name', 'Class', 'Subject', 'Session']
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, 1, 0, 'C')
    pdf.ln()
    
    
    # Add table rows
    pdf.set_font('Arial', '', 8)
    for _, row in arrangement_df.iterrows():
        pdf.cell(col_widths[0], 10, str(row['Room']), 1)
        pdf.cell(col_widths[1], 10, str(row['Seat Number']), 1)
        pdf.cell(col_widths[2], 10, str(row['Index Number']), 1)
        pdf.cell(col_widths[3], 10, str(row['Full Name']), 1)
        pdf.cell(col_widths[4], 10, str(row['Class']), 1)
        pdf.cell(col_widths[5], 10, str(row['Subject']), 1)
        pdf.cell(col_widths[6], 10, str(row['Session']), 1)
        pdf.ln()
        
    return pdf.output(dest='S')

def generate_arrangement(df, room_capacities, subject_sessions, exam_date):
    """Generates an Excel workbook with seating arrangements for multiple subjects in rooms."""
    # Ensure required columns exist
    required_columns = ['IndexNumber', 'Full_Name', 'Core_Subjects', 'Elective_Subjects', 'Class']
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        st.error(f"The uploaded file is missing required columns: {', '.join(missing_cols)}")
        return None

    # Create a long-form dataframe of students and their subjects
    student_subjects = []
    for _, row in df.iterrows():
        core_subjects = str(row.get('Core_Subjects', '')).split(',') if pd.notna(row.get('Core_Subjects')) else []
        elective_subjects = str(row.get('Elective_Subjects', '')).split(',') if pd.notna(row.get('Elective_Subjects')) else []
        all_subjects_list = [s.strip() for s in core_subjects + elective_subjects if s.strip()]
        for subject in all_subjects_list:
            student_subjects.append({
                'IndexNumber': row['IndexNumber'],
                'Full_Name': row['Full_Name'],
                'Class': row['Class'],
                'Subject': subject
            })

    if not student_subjects:
        st.warning("No subjects could be processed from the uploaded file.")
        return None

    long_df = pd.DataFrame(student_subjects)
    
    # Filter for selected subjects
    selected_subjects = list(subject_sessions.keys())
    long_df = long_df[long_df['Subject'].isin(selected_subjects)]

    if long_df.empty:
        st.warning("No students found for the selected subjects.")
        return None

    # Add session information
    long_df['Session'] = long_df['Subject'].map(subject_sessions)

    # Handle "Both" session
    both_session_df = long_df[long_df['Session'] == 'Both'].copy()
    if not both_session_df.empty:
        morning_df = both_session_df.copy()
        morning_df['Session'] = 'Morning'
        afternoon_df = both_session_df.copy()
        afternoon_df['Session'] = 'Afternoon'
        
        long_df = pd.concat([
            long_df[long_df['Session'] != 'Both'],
            morning_df,
            afternoon_df
        ]).reset_index(drop=True)


    # Sort students by IndexNumber instead of shuffling
    grouped_by_subject = long_df.sort_values('IndexNumber').reset_index(drop=True)
    
    # Create a list of all available seats, including original class name in the room name
    all_seats = []
    room_names = {room: f"Room {i+1} ({room})" for i, room in enumerate(room_capacities.keys())}
    for room, capacity in room_capacities.items():
        for seat_num in range(1, capacity + 1):
            all_seats.append({'Room': room_names[room], 'Seat Number': seat_num})

    if len(grouped_by_subject) > len(all_seats):
        st.error(f"Not enough seats for all students ({len(grouped_by_subject)} required, {len(all_seats)} available). Increase room capacities or add more rooms.")
        return None

    # Assign students to seats
    seating_arrangement = []
    for i, student in grouped_by_subject.iterrows():
        seat = all_seats[i]
        seating_arrangement.append({
            'Room': seat['Room'],
            'Seat Number': seat['Seat Number'],
            'Index Number': student['IndexNumber'],
            'Full Name': student['Full_Name'],
            'Class': student['Class'],
            'Subject': student['Subject'],
            'Session': student['Session']
        })

    arrangement_df = pd.DataFrame(seating_arrangement)

    # Create an Excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Seating Arrangement"

    # Add title row with date
    title_text = f"Seating Arrangement for {exam_date.strftime('%A, %B %d, %Y')}"
    ws.merge_cells('A1:G1')
    title_cell = ws['A1']
    title_cell.value = title_text
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center')

    # Set headers in the second row
    headers = ['Room', 'Seat Number', 'Index Number', 'Full Name', 'Class', 'Subject', 'Session']
    ws.append(headers)
    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # Write data starting from the third row
    for r_idx, row in enumerate(arrangement_df.itertuples(index=False), 3):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Adjust column widths
    ws.column_dimensions['A'].width = 25 # Room name is longer now
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 15

    return wb, arrangement_df # Return dataframe for PDF generation

def get_excel_download_link(wb, filename):
    """Generates a download link for the given workbook."""
    virtual_file = io.BytesIO()
    wb.save(virtual_file)
    virtual_file.seek(0)
    b64 = base64.b64encode(virtual_file.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel file</a>'

def get_pdf_download_link(pdf_bytes, filename):
    """Generates a download link for the given PDF bytes."""
    b64 = base64.b64encode(pdf_bytes).decode()
    return f'<a href="data:application/pdf;base64,{b64}" download="{filename}">Download PDF file</a>'

def add_room_callback():
    """Callback to add a new room to the session state."""
    new_room = st.session_state.get("new_room_input", "").strip()
    if new_room and new_room not in st.session_state.all_classes:
        st.session_state.all_classes.append(new_room)
        st.session_state.all_classes.sort()
        # Clear the input box after adding by resetting its key in session_state
        st.session_state.new_room_input = ""

def run_app():
    """Main function to run the Streamlit app."""
    st.title("Exam Seating Arrangement Generator")
    st.write("Upload a file with student data to generate a seating arrangement for exams based on subjects.")

    uploaded_file = st.file_uploader(
        "Upload student data file (CSV or Excel)",
        type=["csv", "xlsx"],
        help="The file must contain columns: IndexNumber, Full_Name, Class, Core_Subjects, Elective_Subjects"
    )

    if uploaded_file:
        try:
            # Read file with string data type to preserve formats like leading zeros
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, dtype=str)
            else:
                df = pd.read_excel(uploaded_file, dtype=str)
            
            st.success("File uploaded successfully!")
            st.dataframe(df.head())

            # --- Configuration Section ---
            st.header("Exam Configuration")

            exam_date = st.date_input("Select exam date")

            # Get all unique subjects from the dataframe
            if 'Core_Subjects' in df.columns and 'Elective_Subjects' in df.columns:
                core_subjects = df['Core_Subjects'].str.split(',').explode().str.strip().unique()
                elective_subjects = df['Elective_Subjects'].str.split(',').explode().str.strip().unique()
                all_subjects = sorted(list(set(core_subjects) | set(elective_subjects)))
                
                selected_subjects = st.multiselect(
                    "Select subjects for this exam session",
                    options=all_subjects
                )

                subject_sessions = {}
                if selected_subjects:
                    st.subheader("Set Subject Sessions")
                    for subject in selected_subjects:
                        subject_sessions[subject] = st.selectbox(
                            f"Session for {subject}",
                            ["Morning", "Afternoon", "Both"],
                            key=f"session_{subject}"
                        )
            else:
                st.error("The file must contain 'Core_Subjects' and 'Elective_Subjects' columns.")
                return

            if 'Class' in df.columns:
                # Initialize session state on first run for a new file upload
                if 'current_file' not in st.session_state or st.session_state.current_file != uploaded_file.name:
                    st.session_state.current_file = uploaded_file.name
                    all_classes = sorted(df['Class'].unique())
                    st.session_state.all_classes = all_classes
                    st.session_state.selected_classes = all_classes.copy()

                # Allow users to add a custom room using a callback
                st.text_input("Add a new room (optional)", key="new_room_input")
                st.button("Add Room", on_click=add_room_callback)
                
                # The multiselect's state is now managed by Streamlit using its key
                selected_classes = st.multiselect(
                    "Select classes to be used as exam rooms",
                    options=st.session_state.all_classes,
                    default=st.session_state.selected_classes,
                    key="selected_classes_multiselect"
                )
                # Update session state with the current selection from the widget
                st.session_state.selected_classes = selected_classes
            else:
                st.error("The uploaded file must contain a 'Class' column.")
                return

            room_capacities = {}
            if selected_classes:
                st.subheader("Set Room Capacities")
                for room in selected_classes:
                    room_capacities[room] = st.number_input(
                        f"Capacity for {room}",
                        min_value=1,
                        value=30,
                        key=f"capacity_{room}"
                    )

            if st.button("Generate Seating Arrangement"):
                if not selected_classes:
                    st.warning("Please select at least one class to be used as a room.")
                elif not selected_subjects:
                    st.warning("Please select at least one subject for the session.")
                else:
                    with st.spinner("Generating arrangement..."):
                        # Pass the subject sessions to the arrangement function
                        result = generate_arrangement(df, room_capacities, subject_sessions, exam_date)
                        if result:
                            wb, arrangement_df = result
                            st.success("Seating arrangement generated successfully!")
                            
                            # Excel download link
                            excel_filename = f"seating_arrangement_{exam_date.strftime('%Y-%m-%d')}.xlsx"
                            excel_link = get_excel_download_link(wb, excel_filename)
                            
                            # PDF download link
                            pdf_bytes = create_pdf(arrangement_df, exam_date)
                            pdf_filename = f"seating_arrangement_{exam_date.strftime('%Y-%m-%d')}.pdf"
                            pdf_link = get_pdf_download_link(pdf_bytes, pdf_filename)

                            # Class list PDF download link
                            class_list_pdf_bytes = create_class_list_pdf(arrangement_df, exam_date)
                            class_list_pdf_filename = f"class_list_{exam_date.strftime('%Y-%m-%d')}.pdf"
                            class_list_pdf_link = get_pdf_download_link(class_list_pdf_bytes, class_list_pdf_filename)

                            st.markdown(excel_link, unsafe_allow_html=True)
                            st.markdown(pdf_link, unsafe_allow_html=True)
                            st.markdown(class_list_pdf_link.replace("Download PDF file", "Download Class List PDF"), unsafe_allow_html=True)

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")
