import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
import io
import base64
from fpdf import FPDF
from streamlit_sortables import sort_items
import datetime

def create_class_list_pdf(arrangement_df, exam_date):
    """Generates a PDF class list for signing for a specific date."""
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    
    date_df = arrangement_df[arrangement_df['Date'] == exam_date]
    rooms = sorted(date_df['Room'].unique())
    
    for room in rooms:
        pdf.add_page()
        
        pdf.set_font('Arial', 'B', 16)
        title_text = f"Exam List for {room}"
        pdf.cell(0, 10, title_text, 0, 1, 'C')
        
        exam_date_text = f"Exam Date: {exam_date.strftime('%A, %B %d, %Y')}"
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, exam_date_text, 0, 1, 'C')
        pdf.ln(5)
        
        pdf.set_font('Arial', 'B', 10)
        col_widths = [20, 25, 50, 25, 30, 35]
        headers = ['Seat Number', 'Index Number', 'Full Name', 'Class', 'Subject', 'Signature']
        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], 10, header, 1, 0, 'C')
        pdf.ln()
        
        pdf.set_font('Arial', '', 8)
        room_df = date_df[date_df['Room'] == room].sort_values('Seat Number')
        for _, row in room_df.iterrows():
            pdf.cell(col_widths[0], 10, str(row['Seat Number']), 1)
            pdf.cell(col_widths[1], 10, str(row['Index Number']), 1)
            pdf.cell(col_widths[2], 10, str(row['Full Name']), 1)
            pdf.cell(col_widths[3], 10, str(row['Class']), 1)
            pdf.cell(col_widths[4], 10, str(row['Subject']), 1)
            pdf.cell(col_widths[5], 10, '', 1)
            pdf.ln()
            
    return pdf.output(dest='S')

def create_pdf(arrangement_df, exam_date):
    """Generates a PDF file from the arrangement dataframe for a specific date."""
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    pdf.set_font('Arial', 'B', 16)
    title_text = f"Seating Arrangement for {exam_date.strftime('%A, %B %d, %Y')}"
    pdf.cell(0, 10, title_text, 0, 1, 'C')
    
    pdf.set_font('Arial', 'B', 10)
    col_widths = [35, 20, 35, 55, 35, 30, 30]
    headers = ['Room', 'Seat Number', 'Index Number', 'Full Name', 'Class', 'Subject', 'Session']
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, 1, 0, 'C')
    pdf.ln()
    
    pdf.set_font('Arial', '', 8)
    date_df = arrangement_df[arrangement_df['Date'] == exam_date]
    for _, row in date_df.iterrows():
        pdf.cell(col_widths[0], 10, str(row['Room']), 1)
        pdf.cell(col_widths[1], 10, str(row['Seat Number']), 1)
        pdf.cell(col_widths[2], 10, str(row['Index Number']), 1)
        pdf.cell(col_widths[3], 10, str(row['Full Name']), 1)
        pdf.cell(col_widths[4], 10, str(row['Class']), 1)
        pdf.cell(col_widths[5], 10, str(row['Subject']), 1)
        pdf.cell(col_widths[6], 10, str(row['Session']), 1)
        pdf.ln()
        
    return pdf.output(dest='S')

def generate_arrangement(df, room_capacities, subject_details):
    """Generates an Excel workbook with seating arrangements for multiple subjects in rooms."""
    required_columns = ['IndexNumber', 'Full_Name', 'Core_Subjects', 'Elective_Subjects', 'Class']
    if not all(col in df.columns for col in required_columns):
        missing_cols = [col for col in required_columns if col not in df.columns]
        st.error(f"The uploaded file is missing required columns: {', '.join(missing_cols)}")
        return None

    student_subjects = []
    for _, row in df.iterrows():
        core_subjects = str(row.get('Core_Subjects', '')).split(',')
        elective_subjects = str(row.get('Elective_Subjects', '')).split(',')
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
    
    selected_subjects = list(subject_details.keys())
    long_df = long_df[long_df['Subject'].isin(selected_subjects)]

    if long_df.empty:
        st.warning("No students found for the selected subjects.")
        return None

    long_df['Session'] = long_df['Subject'].map(lambda s: subject_details[s]['session'])
    long_df['Date'] = long_df['Subject'].map(lambda s: subject_details[s]['date'])

    both_session_df = long_df[long_df['Session'] == 'Both'].copy()
    if not both_session_df.empty:
        morning_df = both_session_df.copy()
        morning_df['Session'] = 'Morning'
        afternoon_df = both_session_df.copy()
        afternoon_df['Session'] = 'Afternoon'
        long_df = pd.concat([long_df[long_df['Session'] != 'Both'], morning_df, afternoon_df]).reset_index(drop=True)

    grouped_by_subject = long_df.sort_values('IndexNumber').reset_index(drop=True)
    
    all_seats = []
    room_names = {room: f"Room {i+1} ({room})" for i, room in enumerate(room_capacities.keys())}
    for room, capacity in room_capacities.items():
        for seat_num in range(1, capacity + 1):
            all_seats.append({'Room': room_names[room], 'Seat Number': seat_num})

    if len(grouped_by_subject) > len(all_seats):
        st.error(f"Not enough seats for all students ({len(grouped_by_subject)} required, {len(all_seats)} available).")
        return None

    seating_arrangement = []
    for i, student in grouped_by_subject.iterrows():
        seat = all_seats[i]
        seating_arrangement.append({
            'Date': student['Date'],
            'Room': seat['Room'],
            'Seat Number': seat['Seat Number'],
            'Index Number': student['IndexNumber'],
            'Full Name': student['Full_Name'],
            'Class': student['Class'],
            'Subject': student['Subject'],
            'Session': student['Session']
        })

    arrangement_df = pd.DataFrame(seating_arrangement)
    arrangement_df['Date'] = pd.to_datetime(arrangement_df['Date']).dt.date

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for exam_date in sorted(arrangement_df['Date'].unique()):
        ws = wb.create_sheet(title=exam_date.strftime('%Y-%m-%d'))
        
        title_text = f"Seating Arrangement for {exam_date.strftime('%A, %B %d, %Y')}"
        ws.merge_cells('A1:G1')
        title_cell = ws['A1']
        title_cell.value = title_text
        title_cell.font = Font(bold=True, size=14)
        title_cell.alignment = Alignment(horizontal='center')

        headers = ['Room', 'Seat Number', 'Index Number', 'Full Name', 'Class', 'Subject', 'Session']
        ws.append(headers)
        for cell in ws[2]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        date_df = arrangement_df[arrangement_df['Date'] == exam_date]
        for r_idx, row in enumerate(date_df.itertuples(index=False), 3):
            for c_idx, value in enumerate(row[1:], 1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 30
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 20
        ws.column_dimensions['G'].width = 15

    return wb, arrangement_df

def get_excel_download_link(wb, filename):
    virtual_file = io.BytesIO()
    wb.save(virtual_file)
    virtual_file.seek(0)
    b64 = base64.b64encode(virtual_file.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel file</a>'

def get_pdf_download_link(pdf_bytes, filename):
    b64 = base64.b64encode(pdf_bytes).decode()
    return f'<a href="data:application/pdf;base64,{b64}" download="{filename}">Download PDF file</a>'

def add_room_callback():
    new_room = st.session_state.get("new_room_input", "").strip()
    if new_room and new_room not in st.session_state.all_classes:
        st.session_state.all_classes.append(new_room)
        st.session_state.all_classes.sort()
        if 'ordered_rooms' in st.session_state and new_room not in st.session_state.ordered_rooms:
            st.session_state.ordered_rooms.append(new_room)
        st.session_state.new_room_input = ""

def run_app():
    st.title("Exam Seating Arrangement Generator")
    st.write("Upload a file with student data to generate a seating arrangement for exams based on subjects.")

    uploaded_file = st.file_uploader(
        "Upload student data file (CSV or Excel)",
        type=["csv", "xlsx"],
        help="The file must contain columns: IndexNumber, Full_Name, Class, Core_Subjects, Elective_Subjects"
    )

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, dtype=str) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file, dtype=str)
            st.success("File uploaded successfully!")
            st.dataframe(df.head())

            st.header("Exam Configuration")

            if 'Core_Subjects' in df.columns and 'Elective_Subjects' in df.columns:
                core_subjects = df['Core_Subjects'].str.split(',').explode().str.strip().unique()
                elective_subjects = df['Elective_Subjects'].str.split(',').explode().str.strip().unique()
                all_subjects = sorted(list(set(core_subjects) | set(elective_subjects)))
                
                selected_subjects = st.multiselect("Select subjects for this exam session", options=all_subjects)

                subject_details = {}
                if selected_subjects:
                    st.subheader("Set Subject Dates and Sessions")
                    today = datetime.date.today()
                    for subject in selected_subjects:
                        cols = st.columns([0.5, 0.5])
                        date = cols[0].date_input("Date", value=today, key=f"date_{subject}")
                        session = cols[1].selectbox(f"Session for {subject}", ["Morning", "Afternoon", "Both"], key=f"session_{subject}")
                        subject_details[subject] = {'date': date, 'session': session}
            else:
                st.error("The file must contain 'Core_Subjects' and 'Elective_Subjects' columns.")
                return

            if 'Class' in df.columns:
                if 'current_file' not in st.session_state or st.session_state.current_file != uploaded_file.name:
                    st.session_state.current_file = uploaded_file.name
                    all_classes = sorted(df['Class'].unique())
                    st.session_state.all_classes = all_classes
                    st.session_state.ordered_rooms = all_classes.copy()

                st.text_input("Add a new room (optional)", key="new_room_input")
                st.button("Add Room", on_click=add_room_callback)

                st.subheader("Manage and Order Rooms")
                selected_rooms_multiselect = st.multiselect("Select classes to use as rooms", options=st.session_state.all_classes, default=st.session_state.ordered_rooms)
                
                st.session_state.ordered_rooms = [room for room in st.session_state.ordered_rooms if room in selected_rooms_multiselect]
                for room in selected_rooms_multiselect:
                    if room not in st.session_state.ordered_rooms:
                        st.session_state.ordered_rooms.append(room)

                st.write("Order of Rooms (Drag to reorder, top has higher priority):")
                st.session_state.ordered_rooms = sort_items(st.session_state.ordered_rooms, direction='vertical')
            else:
                st.error("The uploaded file must contain a 'Class' column.")
                return

            room_capacities = {}
            if st.session_state.get('ordered_rooms'):
                st.subheader("Set Room Capacities")
                for room in st.session_state.ordered_rooms:
                    room_capacities[room] = st.number_input(f"Capacity for {room}", min_value=1, value=30, key=f"capacity_{room}")

            if st.button("Generate Seating Arrangement"):
                if not st.session_state.get('ordered_rooms'):
                    st.warning("Please select at least one class to be used as a room.")
                elif not selected_subjects:
                    st.warning("Please select at least one subject for the session.")
                else:
                    with st.spinner("Generating arrangement..."):
                        result = generate_arrangement(df, room_capacities, subject_details)
                        if result:
                            wb, arrangement_df = result
                            st.success("Seating arrangement generated successfully!")
                            
                            excel_filename = "seating_arrangement_all_dates.xlsx"
                            st.markdown(get_excel_download_link(wb, excel_filename), unsafe_allow_html=True)

                            st.subheader("Download PDFs by Date")
                            for exam_date in sorted(arrangement_df['Date'].unique()):
                                st.write(f"**Downloads for {exam_date.strftime('%A, %B %d, %Y')}:**")
                                
                                pdf_bytes = create_pdf(arrangement_df, exam_date)
                                pdf_filename = f"seating_arrangement_{exam_date.strftime('%Y-%m-%d')}.pdf"
                                st.markdown(get_pdf_download_link(pdf_bytes, pdf_filename), unsafe_allow_html=True)

                                class_list_pdf_bytes = create_class_list_pdf(arrangement_df, exam_date)
                                class_list_pdf_filename = f"exam_list_{exam_date.strftime('%Y-%m-%d')}.pdf"
                                st.markdown(get_pdf_download_link(class_list_pdf_bytes, class_list_pdf_filename).replace("Download PDF file", "Download Exam List PDF"), unsafe_allow_html=True)

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")
