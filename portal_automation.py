import streamlit as st
import pandas as pd
import time
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

def run_wassce_automation(df, username, password, status_container, log_container):
    """
    Runs the Playwright automation to submit data to the portal.
    """
    # Configuration
    PORTAL_LOGIN_URL = "https://cass.waecinternetsolution.org/"
    
    # DEFAULT SELECTORS (To be updated by the user)
    SELECTORS = {
        "login_username": "#username_id",   # placeholder
        "login_password": "#password_id",   # placeholder
        "login_btn": "#login_button_id",    # placeholder
        "dashboard_indicator": ".dashboard-welcome", # placeholder
        
        "surname": "#surname",             # placeholder
        "middle_name": "#middle_name",     # placeholder
        "first_name": "#first_name",       # placeholder
        "dob": "#dob",                     # placeholder
        
        "gender_male": "input[value='Male']",     # placeholder
        "gender_female": "input[value='Female']", # placeholder
        
        "school_index": "#index_number",   # placeholder
        "completion_year": "#year_completed", # placeholder
        "programme_dropdown": "#programme",   # placeholder
        
        "submit_btn": "#save_entry_btn",      # placeholder
    }

    log_logs = []

    def log(message):
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        log_logs.append(log_entry)
        # Update the UI with the latest 10 logs
        log_container.code("\n".join(log_logs[-10:]))
        print(log_entry)

    # Fix for Streamlit/Asyncio conflict
    import asyncio
    try:
        asyncio.get_event_loop()
    except RuntimeError:
        asyncio.set_event_loop(asyncio.new_event_loop())

    log("Starting automation...")

    try:
        with sync_playwright() as p:
            status_container.info("Launching browser...")
            browser = p.chromium.launch(headless=False)
            context = browser.new_context()
            page = context.new_page()

            # Login
            status_container.info("Navigating to login page...")
            log(f"Opening {PORTAL_LOGIN_URL}")
            page.goto(PORTAL_LOGIN_URL)

            try:
                # Fill credentials
                log("Entering credentials...")
                # Note: These selectors need to be valid for this to work
                if page.is_visible(SELECTORS["login_username"]):
                    page.fill(SELECTORS["login_username"], username)
                    page.fill(SELECTORS["login_password"], password)
                    page.click(SELECTORS["login_btn"])
                    
                    status_container.info("Waiting for dashboard...")
                    # wait for navigation or specific element
                    # page.wait_for_selector(SELECTORS["dashboard_indicator"], timeout=30000) 
                    time.sleep(5) # Temporary wait if selectors aren't real yet
                    log("Login successful (simulated wait).")
                else:
                    log("Warning: Login fields not found. Verify selectors.")
            except Exception as e:
                log(f"Login failed: {e}")
                browser.close()
                return

            # Process Rows
            total_rows = len(df)
            status_bar = status_container.progress(0)
            
            for index, row in df.iterrows():
                row_num = index + 1
                perc = int((index / total_rows) * 100)
                status_bar.progress(perc)
                
                surname = str(row.get('Surname', '')).strip()
                first_name = str(row.get('First Name', '')).strip()
                log(f"Processing Row {row_num}: {surname} {first_name}")

                try:
                    # Here we would interact with the form
                    # page.goto(FORM_URL) # if needed
                    
                    # Example of filling a field safely
                    # if page.is_visible(SELECTORS["surname"]):
                    #     page.fill(SELECTORS["surname"], surname)
                    
                    # Simulate processing time
                    time.sleep(0.5) 
                    
                except Exception as row_error:
                    log(f"Error processing row {row_num}: {row_error}")
                
            status_bar.progress(100)
            status_container.success("Automation completed!")
            browser.close()
            
    except Exception as e:
        status_container.error(f"Critical Error: {e}")
        log(f"Critical Error: {e}")

def run_app():
    st.header("WASSCE Portal Automation")
    st.warning("Ensure you have the 'playwright' library installed: `pip install playwright && playwright install`")
    
    with st.expander("Automation Instructions", expanded=True):
        st.write("""
        1. Upload an Excel file containing the student data.
        2. Enter your portal username and password.
        3. Click 'Start Automation' to launch the browser and submit data.
        **Required Excel Headers:** Surname, Middle Name, First Name, Date of Birth, Gender, Basic School IndexNumber, Basic School Completion Year, Programme.
        """)
        
    col_auth1, col_auth2 = st.columns(2)
    with col_auth1:
        username = st.text_input("Portal Username")
    with col_auth2:
        password = st.text_input("Portal Password", type="password")
        
    auto_file = st.file_uploader("Upload Automation Excel Data", type=["xlsx", "xls"], key="auto_upload")
    
    if auto_file and username and password:
        if st.button("Start Automation"):
            status_box = st.empty()
            log_box = st.empty()
            
            try:
                df_auto = pd.read_excel(auto_file, dtype=str)
                df_auto = df_auto.fillna("")
                run_wassce_automation(df_auto, username, password, status_box, log_box)
            except Exception as e:
                st.error(f"Failed to read file: {e}")

if __name__ == "__main__":
    st.set_page_config(page_title="WASSCE Automation")
    run_app()
