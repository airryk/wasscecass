import sys
import time
import logging
import pandas as pd
import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# Load environment variables
load_dotenv()

# ==========================================
# CONFIGURATION SECTION
# ==========================================

# Login Credentials
# Login Credentials
USERNAME = os.getenv("PORTAL_USERNAME")  # Ensure .env has PORTAL_USERNAME=your_username
PASSWORD = os.getenv("PORTAL_PASSWORD")  # Ensure .env has PORTAL_PASSWORD=your_password

# File Path
# You can use absolute path or relative path
EXCEL_FILE_PATH = "student_data.xlsx" 

# URLs
PORTAL_LOGIN_URL = "https://cass.waecinternetsolution.org/"
FORM_PAGE_URL = "https://cass.waecinternetsolution.org/Student/New"

# ------------------------------------------
# SELECTORS (Update these with real values)
# ------------------------------------------
SELECTORS = {
    # Login Page
    "login_username": "[name='UserName'], #username, input[name*='User']",   # Heuristic update
    "login_password": "[name='Password'], #password, input[name*='Pass']",   # Heuristic update
    "login_btn": "button:has-text('Login'), input[type='submit']",    # Heuristic update
    
    # Form Fields
    "surname": "#Surname",
    "middle_name": "#MiddleName",
    "first_name": "#FirstName",
    "dob": "#DateOfBirth",
    
    # Gender
    "gender_male": "#genderMale",
    "gender_female": "#genderFemale",
    
    # Other
    "school_index": "#BasicSchoolIndexNumber",
    "completion_year": "#BasicSchoolYear",
    "programme_dropdown": "#ProgrammeCode",
    
    "submit_btn": "input[value='Save Student']",
}

# ==========================================
# LOGGING SETUP
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("WASSCE_Automation")

# ==========================================
# MAIN SCRIPT
# ==========================================

def run_automation():
    logger.info("Starting automation script...")

    # 1. Load Excel Data
    try:
        current_file_path = EXCEL_FILE_PATH
        if not os.path.exists(current_file_path):
            logger.warning(f"Default file '{EXCEL_FILE_PATH}' not found.")
            current_file_path = input("Please enter the full path to your Excel file: ").replace('"', '').strip()
            
        logger.info(f"Loading Excel file from: {current_file_path}")
        # Check if file exists first or handle error
        df = pd.read_excel(current_file_path, dtype=str)
        df = df.fillna("")
        logger.info(f"Successfully loaded {len(df)} rows.")
    except Exception as e:
        logger.error(f"Failed to load Excel file: {e}")
        return

    with sync_playwright() as p:
        # Launch Browser
        logger.info("Launching browser...")
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        try:
            # --------------------------------------
            # 2. Login Process
            # --------------------------------------
            logger.info("Navigating to login page...")
            page.goto(PORTAL_LOGIN_URL)
            
            # ATTEMPT AUTOMATIC LOGIN
            try:
                page.wait_for_load_state("networkidle", timeout=5000)
                
                # --- DEBUG: Print Login Page Inputs ---
                logger.info("  --- Login Page Inputs ---")
                login_inputs = page.query_selector_all("input")
                for i, inp in enumerate(login_inputs):
                    id_attr = inp.get_attribute("id") or ""
                    name_attr = inp.get_attribute("name") or ""
                    ph_attr = inp.get_attribute("placeholder") or ""
                    type_attr = inp.get_attribute("type") or ""
                    logger.info(f"    Input {i}: ID='{id_attr}', Name='{name_attr}', Placeholder='{ph_attr}', Type='{type_attr}'")
                logger.info("  -------------------------")

                # Strategy 1: Exact Placeholders/Labels (User provided "User ID")
                if page.get_by_placeholder("User ID").count() > 0:
                    page.get_by_placeholder("User ID").fill(USERNAME)
                    logger.info("Filled Username via placeholder 'User ID'")
                elif page.get_by_label("User ID").count() > 0:
                    page.get_by_label("User ID").fill(USERNAME)
                    logger.info("Filled Username via label 'User ID'")
                # Strategy 2: Generic Names/IDs
                elif page.locator("input[name='UserName']").count() > 0:
                    page.locator("input[name='UserName']").fill(USERNAME)
                    logger.info("Filled Username via name='UserName'")
                elif page.locator("input[name='username']").count() > 0:
                    page.locator("input[name='username']").fill(USERNAME)
                    logger.info("Filled Username via name='username'")
                 # Strategy 3: Text match
                elif page.get_by_text("User ID", exact=False).count() > 0:
                     # This might target a label, we need the input
                     pass 

                # Password
                if page.get_by_placeholder("Password").count() > 0:
                    page.get_by_placeholder("Password").fill(PASSWORD)
                    logger.info("Filled Password via placeholder 'Password'")
                elif page.locator("input[name='Password']").count() > 0:
                    page.locator("input[name='Password']").fill(PASSWORD)
                    logger.info("Filled Password via name='Password'")
                elif page.locator("input[type='password']").count() > 0:
                     page.locator("input[type='password']").first.fill(PASSWORD)
                     logger.info("Filled Password via type='password'")

                # Login Button
                if page.get_by_role("button", name="Login").count() > 0:
                    page.get_by_role("button", name="Login").click()
                    logger.info("Clicked Login via Role")
                elif page.locator("input[type='submit']").count() > 0:
                    page.locator("input[type='submit']").click()
                    logger.info("Clicked Login via type='submit'")
                elif page.locator("button:has-text('Login')").count() > 0:
                    page.locator("button:has-text('Login')").click()
                    logger.info("Clicked Login via text")
                    
            except Exception as e:
                logger.warning(f"Auto-fill attempt had issues: {e}")

            # WAIT FOR LOGIN SUCCESS (Polling)
            logger.info("Waiting for login to complete...")
            try:
                # Poll for success indicators
                # We look for Logout form OR the 'Students' link in navbar
                for _ in range(30): # Wait 30 * 2 = 60 seconds max
                    if page.locator("#logoutForm").count() > 0 or \
                       page.locator("a[href='/Student/Index']").count() > 0:
                        logger.info("Login detected! Proceeding...")
                        break
                    time.sleep(2)
                else:
                    # If loop finishes without break
                    logger.warning("Could not auto-detect login success.")
                    input("If you are logged in, press Enter to continue script execution... ")
                    
            except Exception as e:
                logger.warning(f"Error during login wait: {e}")
                input("Press Enter if you are logged in...")
            
            # --------------------------------------
            # 3. Process Rows
            # --------------------------------------
            for index, row in df.iterrows():
                row_num = index + 1
                surname = str(row.get("Surname", "")).strip()
                first_name = str(row.get("First Name", "")).strip()
                
                logger.info(f"Processing Row {row_num}: {surname} {first_name}")

                try:
                    # Navigate to the Entry Page
                    logger.info(f"  Navigating to form: {FORM_PAGE_URL}")
                    page.goto(FORM_PAGE_URL)
                    page.wait_for_load_state("networkidle")
                    
                    # --- DEBUGGING: SAVE PAGE SOURCE & PRINT INPUTS ---
                    with open("form_page_source.html", "w", encoding="utf-8") as f:
                        f.write(page.content())
                    logger.info("  Saved page source to 'form_page_source.html'")
                    
                    # Print all inputs to help identify selectors
                    logger.info("  --- Form Fields Found ---")
                    inputs = page.query_selector_all("input, select, textarea")
                    for i, inp in enumerate(inputs):
                        id_attr = inp.get_attribute("id") or ""
                        name_attr = inp.get_attribute("name") or ""
                        placeholder = inp.get_attribute("placeholder") or ""
                        type_attr = inp.get_attribute("type") or ""
                        logger.info(f"    Field {i}: Tag={inp.evaluate('el => el.tagName')}, ID='{id_attr}', Name='{name_attr}', Placeholder='{placeholder}', Type='{type_attr}'")
                    logger.info("  -------------------------")

                    # -- FILL FORM --
                    logger.info("  Filling form fields...")
                    
                    if page.is_visible(SELECTORS["surname"]):
                         # Basic Info
                        page.fill(SELECTORS["surname"], surname)
                        page.fill(SELECTORS["first_name"], first_name)
                        
                        middle = str(row.get("Middle Name", "")).strip()
                        if middle:
                            page.fill(SELECTORS["middle_name"], middle)
                        
                        dob_val = str(row.get("Date of Birth", "")).strip()
                        # Ensure DOB format dd/mm/yyyy if possible
                        if dob_val:
                            try:
                                # Try to parse various date formats and convert to dd/mm/yyyy
                                from datetime import datetime
                                
                                # Try parsing common formats
                                for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"]:
                                    try:
                                        dt = datetime.strptime(dob_val, fmt)
                                        dob_val = dt.strftime("%d/%m/%Y")
                                        logger.info(f"  Converted DOB to: {dob_val}")
                                        break
                                    except ValueError:
                                        continue
                                
                                page.fill(SELECTORS["dob"], dob_val)
                            except Exception as e:
                                logger.warning(f"  Could not parse DOB '{dob_val}': {e}")
                                # Fill it as-is and hope for the best
                                page.fill(SELECTORS["dob"], dob_val)
                            
                        # Gender
                        gender = str(row.get("Gender", "")).strip().lower()
                        if "male" in gender and "female" not in gender:
                            page.click(SELECTORS["gender_male"])
                        elif "female" in gender:
                            page.click(SELECTORS["gender_female"])
                            
                        # Previous School Info
                        idx_num = str(row.get("Basic School IndexNumber", "")).strip()
                        if idx_num:
                            page.fill(SELECTORS["school_index"], idx_num)
                            
                        comp_year = str(row.get("Basic School Completion Year", "")).strip()
                        if comp_year:
                             page.fill(SELECTORS["completion_year"], comp_year)
                             
                        # Programme - THIS TRIGGERS RELOAD
                        prog = str(row.get("Programme", "")).strip()
                        if prog:
                            logger.info(f"  Selecting programme '{prog}' (Expect page reload)...")
                            # We expect a navigation event because of onchange="submit()"
                            try:
                                # We need to match the Excel text to the Option text.
                                # Example: "GENERAL ARTS" -> "General Arts"
                                # We can use label matching which is case-sensitive often, but let's try.
                                # Or we can find the option value first.
                                
                                # Heuristic: Match case-insensitive
                                options = page.query_selector_all(f"{SELECTORS['programme_dropdown']} option")
                                target_value = None
                                for opt in options:
                                    txt = opt.text_content().strip()
                                    if prog.lower() == txt.lower():
                                        target_value = opt.get_attribute("value")
                                        break
                                
                                if target_value:
                                    # Select option triggers submit
                                    with page.expect_navigation():
                                        page.select_option(SELECTORS["programme_dropdown"], value=target_value)
                                    logger.info("  Page reloaded after programme selection.")
                                    page.wait_for_load_state("networkidle")
                                    time.sleep(1)  # Extra wait for dynamic content
                                    
                                    # --- DEBUG: Save reloaded page and inspect checkboxes ---
                                    with open("form_page_after_programme.html", "w", encoding="utf-8") as f:
                                        f.write(page.content())
                                    logger.info("  Saved reloaded page to 'form_page_after_programme.html'")
                                    
                                    # Print ALL checkboxes with their associated labels
                                    logger.info("  --- ALL CHECKBOXES ON PAGE ---")
                                    checkboxes = page.locator("input[type='checkbox']").all()
                                    for idx, cb in enumerate(checkboxes):
                                        cb_id = cb.get_attribute("id") or ""
                                        cb_name = cb.get_attribute("name") or ""
                                        cb_value = cb.get_attribute("value") or ""
                                        cb_checked = cb.is_checked()
                                        
                                        # Try to find associated label
                                        label_text = ""
                                        if cb_id:
                                            label_elem = page.locator(f"label[for='{cb_id}']")
                                            if label_elem.count() > 0:
                                                label_text = label_elem.first.text_content().strip()
                                        
                                        logger.info(f"    Checkbox {idx}: ID='{cb_id}', Name='{cb_name}', Value='{cb_value}', Checked={cb_checked}, Label='{label_text}'")
                                    logger.info("  --------------------------------")
                                    
                                    # --- SELECT SUBJECTS ---
                                    # Check both 'Subject 1' and 'SUBJECT 1' (case variations)
                                    subject_cols = []
                                    for i in range(1, 5):
                                        if f"Subject {i}" in row.index:
                                            subject_cols.append(f"Subject {i}")
                                        elif f"SUBJECT {i}" in row.index:
                                            subject_cols.append(f"SUBJECT {i}")

                                    for sub_col in subject_cols:
                                        subject_name = str(row.get(sub_col, "")).strip()
                                        if subject_name:
                                            logger.info(f"  Selecting subject: {subject_name}")
                                            
                                            selected = False
                                            try:
                                                # Strategy 1: Exact Label Match (Case Insensitive)
                                                xpath = f"//label[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{subject_name.lower()}')]"
                                                
                                                if page.locator(xpath).count() > 0:
                                                    target_label = page.locator(xpath).first
                                                    target_label.scroll_into_view_if_needed()
                                                    
                                                    # Get 'for' attribute to find the actual checkbox input
                                                    for_id = target_label.get_attribute("for")
                                                    
                                                    if for_id:
                                                        checkbox = page.locator(f"#{for_id}")
                                                        
                                                        # Check if it's a hidden input (pre-selected core subject)
                                                        input_type = checkbox.get_attribute("type")
                                                        if input_type == "hidden":
                                                            logger.info(f"    Subject '{subject_name}' is a pre-selected core subject. Skipping.")
                                                            selected = True
                                                        elif checkbox.is_checked():
                                                            logger.info(f"    Subject '{subject_name}' is already selected. Skipping.")
                                                            selected = True
                                                        else:
                                                            checkbox.check()
                                                            logger.info(f"    Selected subject '{subject_name}'.")
                                                            selected = True
                                                    else:
                                                        # If no 'for' attribute, maybe input is nested inside label?
                                                        if target_label.locator("input[type='checkbox']").count() > 0:
                                                            target_label.locator("input[type='checkbox']").first.check()
                                                            selected = True
                                                            logger.info(f"    Selected via nested checkbox.")
                                                        else:
                                                            # Fallback: Click label, but warn it might toggle
                                                            logger.warning(f"    No 'for' ID found for '{subject_name}'. Clicking label (might toggle).")
                                                            target_label.click()
                                                            selected = True
                                                
                                                # Strategy 2: Input Value Match
                                                if not selected:
                                                    if page.locator(f"input[value='{subject_name}']").count() > 0:
                                                        page.locator(f"input[value='{subject_name}']").check()
                                                        selected = True
                                                        logger.info(f"    Selected via Input Value match.")
                                                
                                                if not selected:
                                                    logger.warning(f"    FAILED to find selector for subject '{subject_name}'. Verify spelling!")
                                                    
                                            except Exception as e_sub:
                                                logger.warning(f"    Error selecting subject '{subject_name}': {e_sub}")
                                                
                                    # --- VERIFY SELECTION COUNT ---
                                    # Heuristic: Count checked boxes in the subject area (assuming they are grouped)
                                    # This is hard without specific container selector, but we can try
                                    # page.locator("input[type='checkbox']:checked").count() 
                                                
                                else:
                                    logger.warning(f"  Programme '{prog}' not found in dropdown options.")
                                                

                                    
                            except Exception as e_prog:
                                logger.warning(f"  Issue selecting programme/waiting for reload: {e_prog}")

                        # Click Save Final
                        # After reload, we need to click "Save Student"
                        logger.info("  Clicking 'Save Student'...")
                        if page.is_visible(SELECTORS["submit_btn"]):
                            # Click and wait for navigation or response
                            try:
                                # Expect either navigation or page reload
                                with page.expect_navigation(timeout=10000):
                                    page.click(SELECTORS["submit_btn"])
                                
                                logger.info("  Page navigated after Save.")
                                page.wait_for_load_state("networkidle")
                                time.sleep(2)  # Extra wait for server processing
                                
                                # Check for error message
                                if page.locator("text=An error occurred").count() > 0:
                                    logger.error("  ERROR: Server returned an error message!")
                                    logger.error("  Please check the browser to see the error details.")
                                    input("  Press Enter after you've reviewed the error...")
                                else:
                                    logger.info(f"  Successfully saved row {row_num}.")
                                    
                            except Exception as nav_error:
                                # If no navigation happens, maybe it's an AJAX submit?
                                logger.warning(f"  No navigation detected after Save: {nav_error}")
                                time.sleep(3)
                                
                                # Check for success or error messages
                                if page.locator("text=An error occurred").count() > 0:
                                    logger.error("  ERROR: Server returned an error!")
                                    input("  Press Enter to continue...")
                                else:
                                    logger.info(f"  Saved row {row_num} (no navigation).")
                        else:
                            logger.warning("  Save button not found (maybe reload failed?).")
                            
                    else:
                         logger.warning("  Surname field not found. Navigation might have failed.")
                         
                    time.sleep(1) # Brief pause before next row

                except Exception as row_error:
                    logger.error(f"  Error processing row {row_num}: {row_error}")
            
            logger.info("All rows processed.")
            
            # Keep browser open until user is ready
            print("\n" + "="*50)
            input("Execution Completed. Press Enter to close the browser and exit...")
            print("="*50 + "\n")

        except Exception as e:
            logger.error(f"Critical script error: {e}")
        finally:
            browser.close()

if __name__ == "__main__":
    run_automation()
