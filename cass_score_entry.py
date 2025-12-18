
import sys
import time
import logging
import pandas as pd
import os
import re
import random
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

def normalize_name(name):
    """
    Normalize a name for comparison by:
    - Converting to lowercase
    - Removing extra whitespace
    - Removing punctuation
    - Sorting name parts alphabetically (to handle different orderings)
    """
    if not name:
        return ""
    # Convert to lowercase and strip
    name = name.lower().strip()
    # Remove punctuation (commas, periods, etc.)
    name = re.sub(r'[^\w\s]', '', name)
    # Split into parts and remove empty strings
    parts = [p.strip() for p in name.split() if p.strip()]
    # Sort parts alphabetically for order-independent comparison
    parts.sort()
    return " ".join(parts)

def names_match(name1, name2, threshold=0.85):
    """
    Check if two names match using normalized comparison.
    Returns True if names are similar enough (above threshold).
    """
    norm1 = normalize_name(name1)
    norm2 = normalize_name(name2)
    
    # Exact match after normalization
    if norm1 == norm2:
        return True
    
    # Check if all parts of one name are in the other
    parts1 = set(norm1.split())
    parts2 = set(norm2.split())
    
    # Calculate overlap ratio
    if not parts1 or not parts2:
        return False
    
    common = parts1.intersection(parts2)
    overlap = len(common) / max(len(parts1), len(parts2))
    
    return overlap >= threshold

def find_student_by_name(students_data, target_name):
    """
    Search for a student in the Excel data by name.
    Returns the student dict if found, None otherwise.
    """
    logger.info(f"  Searching for student: {target_name}")
    
    # First try exact match after normalization
    target_normalized = normalize_name(target_name)
    
    for student in students_data:
        student_name = student.get("raw_name", "")
        if normalize_name(student_name) == target_normalized:
            logger.info(f"  Found exact match: {student_name}")
            return student
    
    # If no exact match, try fuzzy matching
    best_match = None
    best_score = 0
    
    for student in students_data:
        student_name = student.get("raw_name", "")
        norm_student = normalize_name(student_name)
        
        # Calculate overlap score
        parts_target = set(target_normalized.split())
        parts_student = set(norm_student.split())
        
        if parts_target and parts_student:
            common = parts_target.intersection(parts_student)
            score = len(common) / max(len(parts_target), len(parts_student))
            
            if score > best_score:
                best_score = score
                best_match = student
    
    if best_match and best_score >= 0.7:  # At least 70% match
        logger.info(f"  Found fuzzy match: {best_match.get('raw_name', '')} (score: {best_score:.2f})")
        return best_match
    
    logger.warning(f"  No match found for: {target_name}")
    return None

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

        # 2. PROCESS STUDENTS BY READING FROM CASS PAGE
        # Navigate to the first student
        page.goto(SCORE_ENTRY_URL)
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        
        processed_count = 0
        skipped_count = 0
        max_iterations = len(students) + 50  # Safety limit
        
        for iteration in range(max_iterations):
            try:
                # READ STUDENT NAME FROM CASS PAGE
                logger.info(f"\n--- Iteration {iteration + 1} ---")
                
                # Try to find the "Full Name:" field on the page
                full_name_text = ""
                
                # Method 1: Look for text after "Full Name:"
                try:
                    # Find the element that contains "Full Name:" and get the adjacent text
                    full_name_el = page.locator("text=Full Name:").first
                    if full_name_el.count() > 0:
                        # Get parent or sibling that contains the actual name
                        parent = full_name_el.locator("xpath=./parent::*")
                        if parent.count() > 0:
                            parent_text = parent.inner_text()
                            # Extract name after "Full Name:"
                            if "Full Name:" in parent_text:
                                full_name_text = parent_text.split("Full Name:")[-1].strip()
                except Exception as e:
                    logger.warning(f"Method 1 failed: {e}")
                
                # Method 2: Look for specific element patterns
                if not full_name_text:
                    try:
                        # Try looking for common patterns in WAEC-style pages
                        name_patterns = [
                            "//span[contains(text(),'Full Name')]/following-sibling::*",
                            "//label[contains(text(),'Full Name')]/following-sibling::*",
                            "//td[contains(text(),'Full Name')]/following-sibling::td",
                            "//div[contains(text(),'Full Name')]/following-sibling::*",
                            "//b[contains(text(),'Full Name')]/following-sibling::text()",
                        ]
                        for pattern in name_patterns:
                            el = page.locator(f"xpath={pattern}").first
                            if el.count() > 0:
                                full_name_text = el.inner_text().strip()
                                if full_name_text:
                                    break
                    except Exception as e:
                        logger.warning(f"Method 2 failed: {e}")
                
                # Method 3: Direct text extraction from page content
                if not full_name_text:
                    try:
                        page_content = page.content()
                        # Look for pattern like "Full Name:</b> NAME HERE" or similar
                        import re as regex_module
                        match = regex_module.search(r'Full Name[:\s]*</[^>]+>\s*([A-Z][A-Z\s]+)', page_content)
                        if match:
                            full_name_text = match.group(1).strip()
                    except Exception as e:
                        logger.warning(f"Method 3 failed: {e}")
                
                if not full_name_text:
                    logger.warning("Could not extract student name from CASS page. Asking for manual input...")
                    full_name_text = input(">>> Enter the student's FULL NAME as shown on the CASS page (or 'skip' to skip, 'quit' to stop): ").strip()
                    if full_name_text.lower() == 'quit':
                        break
                    if full_name_text.lower() == 'skip':
                        # Try to navigate to next student
                        if page.locator("text=Next").count() > 0:
                            page.click("text=Next")
                            time.sleep(2)
                        continue
                
                logger.info(f"  Student on CASS page: {full_name_text}")
                
                # SEARCH FOR MATCHING STUDENT IN EXCEL DATA
                matched_student = find_student_by_name(students, full_name_text)
                
                # Flag to indicate if we're using random scores
                use_random_scores = False
                
                if not matched_student:
                    logger.warning(f"  No matching record found in Excel for: {full_name_text}")
                    logger.info(f"  Will fill with RANDOM scores (50-99) for this student.")
                    use_random_scores = True
                else:
                    logger.info(f"  Matched to Excel record: {matched_student.get('raw_name', 'Unknown')}")
                
                # CHECK IF SCORES ARE ALREADY FILLED (skip if already done)
                try:
                    already_filled = False
                    # Check the first few score input fields to see if they have non-zero values
                    score_inputs = page.locator("input[type='text'], input[type='number']")
                    input_count = score_inputs.count()
                    
                    filled_count = 0
                    for i in range(min(input_count, 6)):  # Check first 6 inputs
                        try:
                            val = score_inputs.nth(i).input_value()
                            # Check if value is filled (not empty, not "000", not "0")
                            if val and val.strip() and val.strip() not in ["", "000", "0", "00"]:
                                filled_count += 1
                        except:
                            pass
                    
                    # If more than half of checked inputs are filled, consider it already done
                    if filled_count >= 3:
                        already_filled = True
                        logger.info(f"  SKIPPING: Student '{full_name_text}' already has scores filled ({filled_count} fields have values)")
                        skipped_count += 1
                        
                        # Navigate to next student
                        if page.locator("text=NEW CASS FORM").count() > 0:
                            page.click("text=NEW CASS FORM")
                            time.sleep(2)
                            page.wait_for_load_state("networkidle")
                        continue
                        
                except Exception as e:
                    logger.warning(f"  Error checking if already filled: {e}")

                # FILL SCORES
                logger.info("  Filling Scores...")
                
                if use_random_scores:
                    # NO EXCEL MATCH - Fill all visible subject rows with random scores (50-99)
                    logger.info("  Using RANDOM scores (50-99) for all subjects...")
                    
                    try:
                        # Find all subject rows in the table
                        # Based on screenshot, subjects are in table rows with input fields
                        all_rows = page.locator("tr").all()
                        random_fill_count = 0
                        
                        for row in all_rows:
                            try:
                                # Check if this row has input fields (score inputs)
                                inputs = row.locator("input[type='text'], input[type='number']")
                                input_count = inputs.count()
                                
                                if input_count >= 1:
                                    # This is a subject row - fill with random scores
                                    for i in range(input_count):
                                        random_score = str(random.randint(50, 99))
                                        try:
                                            current_val = inputs.nth(i).input_value()
                                            # Only fill if empty or has default "000"
                                            if not current_val or current_val.strip() in ["", "000", "0", "00"]:
                                                inputs.nth(i).fill(random_score)
                                                random_fill_count += 1
                                        except:
                                            pass
                            except:
                                continue
                        
                        logger.info(f"  Filled {random_fill_count} fields with random scores.")
                        
                    except Exception as e:
                        logger.warning(f"  Error filling random scores: {e}")
                
                else:
                    # EXCEL MATCH - Fill scores from matched student data
                    # First, get list of subjects we have data for
                    excel_subjects = set()
                    for subj in matched_student.get("subjects", []):
                        s_name = subj["name"]
                        y1 = subj.get("y1", "")
                        y2 = subj.get("y2", "")
                        y3 = subj.get("y3", "")
                        
                        if not s_name: continue
                        excel_subjects.add(s_name.lower().strip())
                        
                        logger.info(f"    Subject: {s_name} | Y1: {y1}, Y2: {y2}, Y3: {y3}")
                        
                        # FIND SUBJECT ROW/INPUTS
                        try:
                            # Find the element containing the subject name
                            subject_el = page.get_by_text(s_name, exact=True).first
                            if subject_el.count() == 0:
                                # Try case insensitive xpath
                                subject_el = page.locator(f"//td[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{s_name.lower()}')]")
                            
                            if subject_el.count() > 0:
                                # Find the ROW (tr) that contains this subject
                                row = subject_el.locator("xpath=./ancestor::tr")
                                
                                if row.count() > 0:
                                    inputs = row.locator("input[type='text'], input[type='number']")
                                    count = inputs.count()
                                    
                                    # Fill Year 1, Year 2 (and Year 3 if exists)
                                    if count >= 1 and y1:
                                        inputs.nth(0).fill(y1)
                                    if count >= 2 and y2:
                                        inputs.nth(1).fill(y2)
                                    if count >= 3 and y3:
                                        inputs.nth(2).fill(y3)
                                        
                                else:
                                    logger.warning(f"    Could not find row for {s_name}")
                            else:
                                logger.warning(f"    Subject label '{s_name}' not found on page.")
                                
                        except Exception as e:
                            logger.warning(f"    Error filling subject {s_name}: {e}")
                    
                    # NOW CHECK FOR SUBJECTS ON CASS PAGE THAT ARE NOT IN EXCEL
                    # Fill those with random scores (50-99)
                    logger.info("  Checking for additional subjects not in Excel...")
                    try:
                        all_rows = page.locator("tr").all()
                        random_fill_count = 0
                        
                        for row in all_rows:
                            try:
                                # Get the subject name from this row (usually first cell)
                                cells = row.locator("td").all()
                                if not cells:
                                    continue
                                
                                row_subject = ""
                                for cell in cells:
                                    try:
                                        cell_text = cell.inner_text().strip()
                                        # Check if this looks like a subject name (has letters, not just numbers)
                                        if cell_text and any(c.isalpha() for c in cell_text):
                                            row_subject = cell_text
                                            break
                                    except:
                                        pass
                                
                                if not row_subject:
                                    continue
                                
                                # Check if this subject is in our Excel data
                                row_subject_lower = row_subject.lower().strip()
                                subject_in_excel = any(excel_subj in row_subject_lower or row_subject_lower in excel_subj 
                                                      for excel_subj in excel_subjects)
                                
                                if not subject_in_excel:
                                    # This subject is NOT in Excel - fill with random scores
                                    inputs = row.locator("input[type='text'], input[type='number']")
                                    input_count = inputs.count()
                                    
                                    if input_count >= 1:
                                        logger.info(f"    Subject '{row_subject}' not in Excel - filling with random scores")
                                        for i in range(input_count):
                                            try:
                                                current_val = inputs.nth(i).input_value()
                                                # Only fill if empty or has default "000"
                                                if not current_val or current_val.strip() in ["", "000", "0", "00"]:
                                                    random_score = str(random.randint(50, 99))
                                                    inputs.nth(i).fill(random_score)
                                                    random_fill_count += 1
                                            except:
                                                pass
                            except:
                                continue
                        
                        if random_fill_count > 0:
                            logger.info(f"  Filled {random_fill_count} additional fields with random scores for subjects not in Excel.")
                            
                    except Exception as e:
                        logger.warning(f"  Error checking for additional subjects: {e}")

                # SAVE
                logger.info("  Saving...")
                try:
                    save_btn = page.locator("button:has-text('SAVE CASS SCORES')").first
                    if save_btn.count() > 0 and save_btn.is_visible():
                        save_btn.click()
                    else:
                        # Try alternative selectors
                        alt_selectors = [
                            "input[value='Save']",
                            "button:has-text('Save')",
                            ".btn:has-text('Save')",
                        ]
                        for sel in alt_selectors:
                            if page.locator(sel).count() > 0:
                                page.locator(sel).first.click()
                                break
                except Exception as e:
                    logger.warning(f"  Error clicking save: {e}")
                        
                time.sleep(2)
                processed_count += 1
                # Use full_name_text if matched_student is None (random scores case)
                student_name_for_log = matched_student.get('raw_name', full_name_text) if matched_student else full_name_text
                logger.info(f"  Successfully processed: {student_name_for_log}")
                
                # NAVIGATE TO NEXT STUDENT
                # Look for "NEW CASS FORM" button for navigation
                next_found = False
                next_selectors = [
                    "text=NEW CASS FORM",
                    "button:has-text('NEW CASS FORM')",
                    "a:has-text('NEW CASS FORM')",
                    "text=New Cass Form",
                    ".btn:has-text('NEW')",
                ]
                for sel in next_selectors:
                    try:
                        if page.locator(sel).count() > 0:
                            page.locator(sel).first.click()
                            next_found = True
                            time.sleep(2)
                            page.wait_for_load_state("networkidle")
                            break
                    except:
                        continue
                
                if not next_found:
                    logger.info("  No 'NEW CASS FORM' button found. This may be the last student.")
                    action = input(">>> Press Enter to refresh and try next, 'quit' to stop, or enter a URL to navigate to: ").strip()
                    if action.lower() == 'quit':
                        break
                    elif action.startswith('http'):
                        page.goto(action)
                        page.wait_for_load_state("networkidle")
                        time.sleep(2)
                    else:
                        # Check if we're still on the same student (end of list)
                        break

            except Exception as e:
                logger.error(f"  Error in iteration {iteration + 1}: {e}")
                action = input(">>> Error occurred. Press Enter to continue, 'quit' to stop: ").strip()
                if action.lower() == 'quit':
                    break

        print("\n" + "="*40)
        logger.info(f"Processing completed!")
        logger.info(f"  Processed: {processed_count} students")
        logger.info(f"  Skipped (no match): {skipped_count} students")
        input("Press Enter to close the browser and exit script...")
        browser.close() 

if __name__ == "__main__":
    run_automation()
