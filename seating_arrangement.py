import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
import io
import base64
from fpdf import FPDF
from streamlit_sortables import sort_items
import datetime

class PDFWithPageNumber(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def create_class_list_pdf(arrangement_df, exam_date):
    """Generates a PDF class list for signing for a specific date."""
    pdf = PDFWithPageNumber(orientation='P', unit='mm', format='A4')
    
    date_df = arrangement_df[arrangement_df['Date'] == exam_date]
    if date_df.empty:
        return None
    
    for session in sorted(date_df['Session'].unique()):
        session_df = date_df[date_df['Session'] == session]
        rooms = sorted(session_df['Room'].unique())
        
        for room in rooms:
            pdf.add_page()
            
            pdf.set_font('Arial', 'B', 16)
            title_text = f"Exam List for {room} - {session} Session"
            pdf.cell(0, 10, title_text, 0, 1, 'C')
            
            exam_date_text = f"Exam Date: {exam_date.strftime('%A, %B %d, %Y')}"
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, exam_date_text, 0, 1, 'C')
            pdf.ln(5)
            
            pdf.set_font('Arial', 'B', 10)
            col_widths = [15, 25, 60, 25, 40, 30]
            headers = ['Seat Number', 'Index Number', 'Full Name', 'Class', 'Subject', 'Signature']
            for i, header in enumerate(headers):
                pdf.cell(col_widths[i], 10, header, 1, 0, 'C')
            pdf.ln()
            
            pdf.set_font('Arial', '', 8)
            room_df = session_df[session_df['Room'] == room].sort_values('Seat Number')
            for _, row in room_df.iterrows():
                pdf.cell(col_widths[0], 10, str(row['Seat Number']), 1)
                pdf.cell(col_widths[1], 10, str(row['Index Number']), 1)
                pdf.cell(col_widths[2], 10, str(row['Full Name']), 1)
                pdf.cell(col_widths[3], 10, str(row['Class']), 1)
                pdf.cell(col_widths[4], 10, str(row['Subject']), 1)
                pdf.cell(col_widths[5], 10, '', 1)
                pdf.ln()
                
    # Return the PDF output directly without encoding
    return pdf.output(dest='S')

def create_pdf(arrangement_df, exam_date):
    """Generates a PDF file from the arrangement dataframe for a specific date."""
    pdf = PDFWithPageNumber(orientation='L', unit='mm', format='A4')
    
    date_df = arrangement_df[arrangement_df['Date'] == exam_date]
    if date_df.empty:
        return None
    
    for session in sorted(date_df['Session'].unique()):
        pdf.add_page()
        
        pdf.set_font('Arial', 'B', 16)
        title_text = f"Seating Arrangement for {exam_date.strftime('%A, %B %d, %Y')} - {session} Session"
        pdf.cell(0, 10, title_text, 0, 1, 'C')
        
        pdf.set_font('Arial', 'B', 10)
        col_widths = [40, 20, 35, 65, 35, 40, 30]
        headers = ['Room', 'Seat Number', 'Index Number', 'Full Name', 'Class', 'Subject', 'Session']
        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], 10, header, 1, 0, 'C')
        pdf.ln()
        
        pdf.set_font('Arial', '', 8)
        session_df = date_df[date_df['Session'] == session]
        for _, row in session_df.iterrows():
            pdf.cell(col_widths[0], 10, str(row['Room']), 1)
            pdf.cell(col_widths[1], 10, str(row['Seat Number']), 1)
            pdf.cell(col_widths[2], 10, str(row['Index Number']), 1)
            pdf.cell(col_widths[3], 10, str(row['Full Name']), 1)
            pdf.cell(col_widths[4], 10, str(row['Class']), 1)
            pdf.cell(col_widths[5], 10, str(row['Subject']), 1)
            pdf.cell(col_widths[6], 10, str(row['Session']), 1)
            pdf.ln()
            
    # Return the PDF output directly without encoding
    return pdf.output(dest='S')

def generate_arrangement(df, room_capacities, subject_details):
    """Generates an Excel workbook with seating arrangements for multiple subjects in rooms."""
    required_columns = ['IndexNumber', 'Full_Name', 'Core_Subjects', 'Elective_Subjects', 'Class']
    if not all(col in df.columns for col in required_columns):
        missing_cols = [col for col in required_columns if col not in df.columns]
        st.error(f"The uploaded file is missing required columns: {', '.join(missing_cols)}")
        return None, None, None

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
        return None, None, None

    long_df = pd.DataFrame(student_subjects)
    
    selected_subjects = list(subject_details.keys())
    long_df = long_df[long_df['Subject'].isin(selected_subjects)]

    if long_df.empty:
        st.warning("No students found for the selected subjects.")
        return None, None, None

    long_df['Date'] = long_df['Subject'].map(lambda s: subject_details[s]['date'])
    long_df['Session'] = long_df['Subject'].map(lambda s: subject_details[s]['session'])

    both_session_df = long_df[long_df['Session'] == 'Both'].copy()
    if not both_session_df.empty:
        morning_df = both_session_df.copy()
        morning_df['Session'] = 'Morning'
        afternoon_df = both_session_df.copy()
        afternoon_df['Session'] = 'Afternoon'
        long_df = pd.concat([long_df[long_df['Session'] != 'Both'], morning_df, afternoon_df]).reset_index(drop=True)

    arrangement_df = pd.DataFrame()
    
    unique_dates = sorted(long_df['Date'].unique())

    for exam_date in unique_dates:
        date_df = long_df[long_df['Date'] == exam_date]

        for session in ['Morning', 'Afternoon']:
            session_df = date_df[date_df['Session'] == session]
            
            if session_df.empty:
                continue

            all_seats = []
            room_names = {room: f"Room {i+1} ({room})" for i, room in enumerate(room_capacities.keys())}
            for room_name_key in room_capacities.keys():
                capacity = room_capacities[room_name_key]
                for seat_num in range(1, capacity + 1):
                    all_seats.append({'Room': room_names[room_name_key], 'Seat Number': seat_num})

            if len(session_df) > len(all_seats):
                st.error(f"Not enough seats for the {session} session on {exam_date.strftime('%Y-%m-%d')} ({len(session_df)} required, {len(all_seats)} available).")
                continue

            seating_arrangement = []
            seat_index = 0
            
            subject_counts = session_df['Subject'].value_counts()
            subjects_in_session = subject_counts.index.tolist()
            
            for i, subject in enumerate(subjects_in_session):
                subject_df = session_df[session_df['Subject'] == subject].sort_values('IndexNumber')
                for _, student in subject_df.iterrows():
                    if seat_index < len(all_seats):
                        seat = all_seats[seat_index]
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
                        seat_index += 1
                    else:
                        st.error(f"Ran out of seats during arrangement for the {session} session on {exam_date.strftime('%Y-%m-%d')}.")
                        break
                
                if seat_index >= len(all_seats) and i < len(subjects_in_session) - 1:
                    st.error(f"Not enough seats for all subjects in the {session} session on {exam_date.strftime('%Y-%m-%d')}.")
                    break
            
            if seating_arrangement:
                arrangement_df = pd.concat([arrangement_df, pd.DataFrame(seating_arrangement)])

    if arrangement_df.empty:
        st.warning("No seating arrangement could be generated.")
        return None, None, None
        
    arrangement_df['Date'] = pd.to_datetime(arrangement_df['Date']).dt.date

    # Calculate statistics
    stats = {}
    for exam_date in unique_dates:
        date_df = arrangement_df[arrangement_df['Date'] == exam_date]
        for session in ['Morning', 'Afternoon']:
            session_df = date_df[date_df['Session'] == session]
            if not session_df.empty:
                total_students = len(session_df)
                total_seats = sum(room_capacities.values())
                rooms_used = session_df['Room'].nunique()
                stats[f"{exam_date.strftime('%Y-%m-%d')} {session}"] = {
                    "total_students": total_students,
                    "total_seats": total_seats,
                    "rooms_used": rooms_used
                }

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
            cell.alignment = Alignment(horizontal='center', wrap_text=True)

        date_df = arrangement_df[arrangement_df['Date'] == exam_date]
        for r_idx, row in enumerate(date_df.itertuples(index=False), 3):
            for c_idx, value in enumerate(row[1:], 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.alignment = Alignment(wrap_text=True, vertical='top')

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 40
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 25
        ws.column_dimensions['G'].width = 20

    return wb, arrangement_df, stats

def get_excel_download_link(wb, filename):
    virtual_file = io.BytesIO()
    wb.save(virtual_file)
    virtual_file.seek(0)
    b64 = base64.b64encode(virtual_file.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel file</a>'

def get_pdf_download_link(pdf_bytes, filename):
    b64 = base64.b64encode(pdf_bytes).decode()
    return f'<a href="data:application/pdf;base64,{b64}" download="{filename}">Download PDF file</a>'

def sync_ordered_rooms():
    if 'selected_rooms' not in st.session_state:
        return

    if 'ordered_rooms' not in st.session_state:
        st.session_state.ordered_rooms = st.session_state.selected_rooms.copy()
        return

    ordered = st.session_state.ordered_rooms
    selected = st.session_state.selected_rooms

    new_ordered = [room for room in ordered if room in selected]

    for room in selected:
        if room not in new_ordered:
            new_ordered.append(room)
    
    st.session_state.ordered_rooms = new_ordered

def add_room_callback():
    new_room = st.session_state.get("new_room_input", "").strip()
    if new_room and new_room not in st.session_state.all_classes:
        st.session_state.all_classes.append(new_room)
        st.session_state.all_classes.sort()
        if 'selected_rooms' in st.session_state:
            st.session_state.selected_rooms.append(new_room)
        st.session_state.new_room_input = ""
        sync_ordered_rooms()

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
                core_subjects = df['Core_Subjects'].str.split(',').explode().str.strip().astype(str).unique()
                elective_subjects = df['Elective_Subjects'].str.split(',').explode().str.strip().astype(str).unique()
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
                    st.session_state.selected_rooms = all_classes.copy()

                st.text_input("Add a new room (optional)", key="new_room_input")
                st.button("Add Room", on_click=add_room_callback)

                st.subheader("Manage and Order Rooms")
                st.multiselect(
                    "Select classes to use as rooms",
                    options=st.session_state.all_classes,
                    key="selected_rooms",
                    on_change=sync_ordered_rooms
                )

                st.write("Order of Rooms (Drag to reorder, top has higher priority):")
                if 'ordered_rooms' in st.session_state and st.session_state.ordered_rooms:
                    sorted_rooms = sort_items(st.session_state.ordered_rooms, direction='vertical')
                    if sorted_rooms != st.session_state.ordered_rooms:
                        st.session_state.ordered_rooms = sorted_rooms
                        st.rerun()
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
                        wb, arrangement_df, stats = generate_arrangement(df, room_capacities, subject_details)
                        if wb:
                            st.success("Seating arrangement generated successfully!")

                            st.subheader("Arrangement Statistics")
                            if stats:
                                for session, data in stats.items():
                                    st.write(f"**{session}:**")
                                    st.write(f"  - Total Students: {data['total_students']}")
                                    st.write(f"  - Total Seats Available: {data['total_seats']}")
                                    st.write(f"  - Rooms Used: {data['rooms_used']}")
                            
                            subjects_in_arrangement = arrangement_df['Subject'].unique()
                            subjects_str = "_".join(subjects_in_arrangement).replace(" ", "_")
                            excel_filename = f"seating_arrangement_{subjects_str}.xlsx"
                            st.markdown(get_excel_download_link(wb, excel_filename), unsafe_allow_html=True)

                            st.subheader("Download PDFs by Date")
                            if arrangement_df is not None and not arrangement_df.empty:
                                for exam_date in sorted(arrangement_df['Date'].unique()):
                                    st.write(f"**Downloads for {exam_date.strftime('%A, %B %d, %Y')}:**")

                                    subjects_on_date = arrangement_df[arrangement_df['Date'] == exam_date]['Subject'].unique()
                                    subjects_on_date_str = '_'.join(subjects_on_date).replace(' ', '_')
                                    
                                    pdf_bytes = create_pdf(arrangement_df, exam_date)
                                    if pdf_bytes:
                                        pdf_filename = f"seating_arrangement_{subjects_on_date_str}_{exam_date.strftime('%Y-%m-%d')}.pdf"
                                        st.markdown(get_pdf_download_link(pdf_bytes, pdf_filename), unsafe_allow_html=True)

                                    class_list_pdf_bytes = create_class_list_pdf(arrangement_df, exam_date)
                                    if class_list_pdf_bytes:
                                        class_list_pdf_filename = f"exam_list_{subjects_on_date_str}_{exam_date.strftime('%Y-%m-%d')}.pdf"
                                        st.markdown(get_pdf_download_link(class_list_pdf_bytes, class_list_pdf_filename).replace("Download PDF file", "Download Exam List PDF"), unsafe_allow_html=True)

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")

if __name__ == "__main__":
    run_app()
