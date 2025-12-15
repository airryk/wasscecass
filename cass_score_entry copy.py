
import sys
import time
import logging
import pandas as pd
import os
import re
from datetime import datetime
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ==========================================
# CONFIGURATION
# ==========================================

# LOGIN DETAILS (Load from .env or hardcode for testing)
load_dotenv()
USERNAME = os.getenv("PORTAL_USERNAME", "your_username")
PASSWORD = os.getenv("PORTAL_PASSWORD", "your_password")

# FILE SETTINGS
EXCEL_FILE_PATH = "processed_student_scores.xlsx"

# URLS
LOGIN_URL = "https://cass.waecinternetsolution.org/"
SCORE_ENTRY_URL = "https://cass.waecinternetsolution.org/Student/CassScoreEntry"

# DEFAULTS FOR MISSING DATA (Since your Excel only has Name/class/Scores)
# Update these if you want to hardcode values for missing fields
DEFAULT_DOB = "01/01/2005" 
DEFAULT_GENDER = "Female" # or "Male"
DEFAULT_COMPLETION_YEAR = "2024"

# TIMING
ACTION_DELAY = 0.5

# ==========================================
# LOGGING SETUP
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - [%(levelname)s] - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("CASS_Automation")

# ==========================================
# EXCEL PARSING LOGIC
# ==========================================
def parse_unstructured_excel(file_path):
    """
    Parses the custom Excel format where student data is in blocks.
    """
    logger.info(f"Parsing {file_path}...")
    try:
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        logger.error(f"Error reading Excel: {e}")
        return []

    students = []
    current_student = {}
    
    i = 0
    while i < len(df):
        row = df.iloc[i]
        col0 = str(row[0]).strip()
        
        # Skip headers or empty cells
        is_col0_empty = (col0 == 'nan' or col0 == '' or col0.lower() == 'student details')
        
        # Start of new student block: Non-empty Col 0 that isn't a label
        # (Heuristic: Labels are 'Class:', 'Programme:', 'Index No:')
        is_label = any(col0.lower().startswith(p) for p in ["class:", "programme:", "index no:"])
        
        if not current_student and not is_col0_empty and not is_label:
            # Init new student
            current_student = {
                "raw_name": col0,
                "subjects": []
            }
        
        if current_student:
            # Parse Metadata from Col 0
            if not is_col0_empty:
                val_lower = col0.lower()
                if val_lower.startswith("class:"):
                    current_student["class"] = col0.split(":", 1)[1].strip()
                elif val_lower.startswith("programme:"):
                    current_student["programme"] = col0.split(":", 1)[1].strip()
                elif val_lower.startswith("index no"):
                    current_student["index_no"] = col0.split(":", 1)[1].strip()
            
            # Parse Subjects from Columns 2 (Name), 3 (Y1), 4 (Y2), 5 (Y3)
            # Ensure we have enough columns
            if len(row) >= 6:
                subj_name = str(row[2]).strip()
                if subj_name and subj_name != 'nan' and subj_name.lower() != 'subjects':
                    # Extract scores, handling NaNs
                    def get_score(val):
                        if str(val) == 'nan' or str(val) == '': return ""
                        try:
                            return str(int(float(val)))
                        except:
                            return str(val)

                    s_obj = {
                        "name": subj_name,
                        "y1": get_score(row[3]),
                        "y2": get_score(row[4]),
                        "y3": get_score(row[5])
                    }
                    current_student["subjects"].append(s_obj)

            # Check if we should close this student block
            # Look ahead: If next row starts a new student (Name in Col 0, no label)
            if i + 1 < len(df):
                next_col0 = str(df.iloc[i+1][0]).strip()
                next_is_empty = (next_col0 == 'nan' or next_col0 == '')
                next_is_label = any(next_col0.lower().startswith(p) for p in ["class:", "programme:", "index no:"])
                
                # If next row has content in Col 0 and it's NOT a label, it's a new student name.
                if not next_is_empty and not next_is_label:
                    students.append(current_student)
                    current_student = {}
            else:
                # End of file
                students.append(current_student)
                
        i += 1
        
    logger.info(f"Successfully parsed {len(students)} students.")
    return students

# ==========================================
# PAGE INTERACTION FUNCTIONS
# ==========================================

def split_name(raw_name):
    # Heuristic: First word = Surname, Rest = First Name
    parts = raw_name.strip().split()
    if not parts:
        return "", "", ""
    surname = parts[0]
    first = ""
    middle = ""
    
    if len(parts) > 1:
        first = parts[1]
    if len(parts) > 2:
        middle = " ".join(parts[2:])
        
    return surname, first, middle

def safe_fill(page, selector, value):
    if not value: return
    try:
        if page.is_visible(selector):
            page.fill(selector, value)
        else:
            logger.warning(f"Selector {selector} not visible.")
    except Exception as e:
        logger.warning(f"Error filling {selector}: {e}")

# ==========================================
# MAIN EXECUTION
# ==========================================

def run_automation():
    if not os.path.exists(EXCEL_FILE_PATH):
        logger.error(f"Exel file {EXCEL_FILE_PATH} not found.")
        return

    students = parse_unstructured_excel(EXCEL_FILE_PATH)
    if not students:
        logger.error("No students found in Excel.")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(viewport={"width": 1366, "height": 768})
        page = context.new_page()

# ==========================================
# MAIN EXECUTION
# ==========================================

def run_automation():
    if not os.path.exists(EXCEL_FILE_PATH):
        logger.error(f"Exel file {EXCEL_FILE_PATH} not found.")
        return

    students = parse_unstructured_excel(EXCEL_FILE_PATH)
    if not students:
        logger.error("No students found in Excel.")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(viewport={"width": 1366, "height": 768})
        page = context.new_page()

        # 1. LOGIN
        logger.info("Logging in...")
        try:
            page.goto(LOGIN_URL)
            page.wait_for_load_state("networkidle")
            time.sleep(2)

            # Debug: List all inputs found to help identify selectors
            try:
                inputs = page.query_selector_all("input")
                logger.info(f"Debug: Found {len(inputs)} input fields on login page.")
                for inp in inputs:
                    name = inp.get_attribute("name")
                    id_ = inp.get_attribute("id")
                    ph = inp.get_attribute("placeholder")
                    logger.info(f" - Input: name='{name}', id='{id_}', placeholder='{ph}'")
            except:
                pass

            # ATTEMPT AUTO-LOGIN
            # Strategy: Use confirmed selectors
            try:
                if page.is_visible("#username"):
                    page.fill("#username", USERNAME)
                    logger.info("Filled username.")
                else:
                    logger.warning("Username field (#username) not found!")

                if page.is_visible("#password"):
                    page.fill("#password", PASSWORD)
                    logger.info("Filled password.")
                else:
                    logger.warning("Password field (#password) not found!")

                # Click Login
                # Try generic "Login" button or class 'btn' with text
                if page.locator("button:has-text('Login')").count() > 0:
                    page.locator("button:has-text('Login')").click()
                elif page.locator(".btn:has-text('Login')").count() > 0:
                    page.locator(".btn:has-text('Login')").click()
                elif page.locator("input[type='submit']").count() > 0:
                    page.locator("input[type='submit']").click()
                else:
                     logger.warning("Login button not found.")
                     
            except Exception as e:
                logger.warning(f"Auto-login attempt failed: {e}")

            # Wait for login success OR manual override
            logger.info("Waiting for login success (checking for 'Log off' or 'Student')...")
            
            # We poll for success
            max_retries = 30 # 60s
            logged_in = False
            for _ in range(max_retries):
                if page.locator("text=Log off").count() > 0 or \
                   page.locator("text=Log Out").count() > 0 or \
                   "Login" not in page.title():
                    logged_in = True
                    logger.info("Login detected!")
                    break
                time.sleep(2)
            
            if not logged_in:
                logger.warning("Login detection timed out.")
                input(">>> PLEASE LOGIN MANUALLY IN THE BROWSER, then press ENTER here to continue... <<<")

        except Exception as e:
            logger.error(f"Login sequence error: {e}")
            input("Error occurred. Press Enter to try proceeding anyway...")

        # 2. PROCESS STUDENTS
        for idx, student in enumerate(students):
            raw_name = student.get("raw_name", "Unknown")
            logger.info(f"\n--- Processing {idx+1}/{len(students)}: {raw_name} ---")
            
            surname, first_name, middle_name = split_name(raw_name)
            
            try:
                page.goto(SCORE_ENTRY_URL)
                page.wait_for_load_state("networkidle")
                
                # FILL BIO DATA
                logger.info("  Filling Bio Data...")
                # Note: Adjust selectors based on real page inspection
                safe_fill(page, "#Surname", surname)
                safe_fill(page, "#FirstName", first_name)
                safe_fill(page, "#MiddleName", middle_name)
                
                # Fill defaults for missing data
                safe_fill(page, "#DateOfBirth", DEFAULT_DOB)
                
                # Gender
                # Check based on default
                if "female" in DEFAULT_GENDER.lower():
                    if page.is_visible("#genderFemale"): page.click("#genderFemale")
                else:
                    if page.is_visible("#genderMale"): page.click("#genderMale")
                    
                # Index No
                idx_no = student.get("index_no", "")
                if idx_no:
                    safe_fill(page, "#BasicSchoolIndexNumber", idx_no) # Using WASSCE index as placeholder?
                    
                # Programme
                prog = student.get("programme", "")
                if prog:
                    logger.info(f"  Selecting Programme: {prog}")
                    # Try to select from dropdown
                    try:
                        # Heuristic to find dropdown
                        if page.locator("select[name*='Prog']").count() > 0:
                            page.select_option("select[name*='Prog']", label=prog)
                    except:
                        logger.warning(f"  Could not select programme '{prog}'")

                # FILL SCORES
                logger.info("  Filling Scores...")
                for subj in student["subjects"]:
                    s_name = subj["name"]
                    y1 = subj["y1"]
                    y2 = subj["y2"]
                    y3 = subj["y3"]
                    
                    if not s_name: continue
                    
                    logger.info(f"    Subject: {s_name} | Y1: {y1}, Y2: {y2}, Y3: {y3}")
                    
                    # FIND SUBJECT ROW/INPUTS
                    # This relies on finding the label for the subject and then the inputs
                    # Assuming a grid layout often found in these portals:
                    # <tr> <td>SubjectName</td> <td><input Y1></td> <td><input Y2></td> ... </tr>
                    
                    try:
                        # Find the element containing the subject name (label, or table cell)
                        subject_el = page.get_by_text(s_name, exact=True).first
                        if subject_el.count() == 0:
                            # Try case insensitive xpath
                            subject_el = page.locator(f"//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{s_name.lower()}')]")
                        
                        if subject_el.count() > 0:
                            # Assuming inputs are in the following siblings or same row
                            # Strategy: Find the ROW (tr) that contains this subject
                            row = subject_el.locator("xpath=./ancestor::tr")
                            
                            if row.count() > 0:
                                inputs = row.locator("input[type='text'], input[type='number']")
                                count = inputs.count()
                                
                                # Assuming order is Year 1, Year 2, Year 3
                                if count >= 1 and y1:
                                    inputs.nth(0).fill(y1)
                                if count >= 2 and y2:
                                    inputs.nth(1).fill(y2)
                                if count >= 3 and y3:
                                    # OPTIONAL YEAR 3 CHECK
                                    # The user requirement: "If Year 3 fields do not appear... skip"
                                    # Playwright will verify visibility automatically if we try to interact?
                                    # We check count. If input exists, fill it.
                                    inputs.nth(2).fill(y3)
                                    
                            else:
                                logger.warning(f"    Could not find row for {s_name}")
                        else:
                            logger.warning(f"    Subject label '{s_name}' not found on page.")
                            
                    except Exception as e:
                        logger.warning(f"    Error filling subject {s_name}: {e}")

                # SAVE
                logger.info("  Saving...")
                if page.is_visible("button:has-text('SAVE CASS SCORES')"):
                    page.click("button:has-text('SAVE CASS SCORES')")
                else:
                    # Generic save
                    if page.is_visible("input[value='Save']"):
                        page.click("input[value='Save']")
                        
                time.sleep(2)

            except Exception as e:
                logger.error(f"  Error processing {raw_name}: {e}")

        print("\n" + "="*40)
        logger.info("All iterations completed.")
        input("Press Enter to close the browser and exit script...")
        browser.close() 

if __name__ == "__main__":
    run_automation()
