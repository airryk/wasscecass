import sys
import time
import logging
import pandas as pd
import os
import re
import glob
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
EXCEL_FILE_PATH = "parent_contacts.xlsx"  # Excel file with parent contacts
PHOTOS_FOLDER = "All"  # Folder containing student photos
DEFAULT_PHOTO = "logo.jpg"  # Default photo when student photo not found

# URLS
LOGIN_URL = "https://cass.waecinternetsolution.org/"
STUDENT_DETAILS_URL = "https://cass.waecinternetsolution.org/Student/NewRegistration"

# DEFAULT VALUES
DEFAULT_CONTACT_START = 500000001  # Starting number for default contacts (0500000001)
DEFAULT_DISABILITY_STATUS = "None"

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
logger = logging.getLogger("StudentDetails_Automation")

# ==========================================
# EXCEL PARSING LOGIC
# ==========================================
def parse_parent_contacts_excel(file_path):
    """
    Parses the Excel file containing parent/guardian contact information.
    Expected columns: StudentIndex, StudentName, Contact
    Returns a dictionary keyed by StudentIndex for reliable matching.
    """
    logger.info(f"Parsing {file_path}...")
    
    if not os.path.exists(file_path):
        logger.warning(f"Excel file {file_path} not found. Will use default contacts.")
        return {}
    
    try:
        # Read Excel with all columns as strings to preserve leading zeros
        df = pd.read_excel(file_path, dtype=str)
        
        # Expected columns: StudentIndex, StudentName, Contact
        index_col = None
        name_col = None
        contact_col = None
        
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if col_lower == 'studentindex' or col_lower == 'student index' or 'index' in col_lower:
                index_col = col
            if col_lower == 'studentname' or col_lower == 'student name' or col_lower == 'name':
                name_col = col
            if col_lower == 'contact' or col_lower == 'phone' or col_lower == 'parent':
                contact_col = col
        
        logger.info(f"Identified columns - Index: {index_col}, Name: {name_col}, Contact: {contact_col}")
        
        # Dictionary keyed by StudentIndex
        students_data = {}
        
        for _, row in df.iterrows():
            student_index = str(row[index_col]).strip() if index_col and pd.notna(row[index_col]) else ""
            student_name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
            contact = str(row[contact_col]).strip() if contact_col and pd.notna(row[contact_col]) else ""
            
            # Clean up contact number
            if contact and contact != 'nan' and contact != 'None':
                # Remove any non-numeric characters except +
                contact = re.sub(r'[^\d+]', '', contact)
                
                # Remove + if present
                contact = contact.lstrip('+')
                
                # If contact starts with 233 (Ghana country code), replace with 0
                # e.g., 233542683576 → 0542683576
                if contact.startswith('233') and len(contact) == 12:
                    contact = '0' + contact[3:]
                    logger.debug(f"Converted 233 to 0: {contact}")
                
                # If contact is 9 digits and doesn't start with 0, add leading 0
                # (Ghana phone numbers are 10 digits starting with 0)
                elif contact.isdigit() and len(contact) == 9 and not contact.startswith('0'):
                    contact = '0' + contact
                    logger.debug(f"Added leading 0 to contact: {contact}")
            else:
                contact = ""
            
            # Clean up student index (remove .0 if it's a float)
            if student_index and student_index != 'nan':
                student_index = student_index.replace('.0', '')
                students_data[student_index] = {
                    'index': student_index,
                    'name': student_name,
                    'contact': contact
                }
        
        # Log a sample contact to verify formatting
        sample_contacts = [(k, v.get('contact', '')) for k, v in list(students_data.items())[:3] if v.get('contact')]
        logger.info(f"Sample contacts: {sample_contacts}")
        
        logger.info(f"Successfully parsed {len(students_data)} student records from Excel.")
        return students_data
        
    except Exception as e:
        logger.error(f"Error reading Excel: {e}")
        return {}

# ==========================================
# PHOTO MATCHING FUNCTIONS
# ==========================================
def get_photo_files(folder_path):
    """
    Get all photo files from the specified folder.
    Photos are named by index number (may be missing leading 0).
    Returns a dictionary mapping index numbers to file paths.
    """
    photos = {}
    
    if not os.path.exists(folder_path):
        logger.warning(f"Photos folder '{folder_path}' not found.")
        return photos
    
    # Common image extensions
    extensions = ['*.jpg', '*.jpeg', '*.png', '*.gif', '*.bmp', '*.webp']
    
    for ext in extensions:
        for file_path in glob.glob(os.path.join(folder_path, ext)):
            filename = os.path.basename(file_path)
            name_without_ext = os.path.splitext(filename)[0].strip()
            
            # Store the index as-is (might be missing leading 0)
            photos[name_without_ext] = file_path
            logger.debug(f"  Loaded photo: {name_without_ext} -> {file_path}")
            
            # Also store with leading 0 if it's a 9-digit number (should be 10)
            if name_without_ext.isdigit() and len(name_without_ext) == 9:
                photos['0' + name_without_ext] = file_path
            
    # Also check case-insensitive for uppercase extensions
    for ext in extensions:
        for file_path in glob.glob(os.path.join(folder_path, ext.upper())):
            filename = os.path.basename(file_path)
            name_without_ext = os.path.splitext(filename)[0].strip()
            
            if name_without_ext not in photos:
                photos[name_without_ext] = file_path
                
                # Also store with leading 0 if it's a 9-digit number
                if name_without_ext.isdigit() and len(name_without_ext) == 9:
                    with_zero = '0' + name_without_ext
                    if with_zero not in photos:
                        photos[with_zero] = file_path
    
    logger.info(f"Found {len(photos)} photo mappings in '{folder_path}'")
    # Log some sample filenames for debugging
    sample_keys = list(photos.keys())[:5]
    logger.info(f"Sample photo filenames: {sample_keys}")
    return photos

def find_photo_by_index(student_index, photos_dict, photos_folder):
    """
    Find the photo file for a given student index number.
    Handles cases where photo filename might be missing leading 0.
    Returns the file path if found, otherwise returns the default photo path.
    """
    if not student_index:
        # Return default photo
        default_path = os.path.join(photos_folder, DEFAULT_PHOTO)
        if os.path.exists(default_path):
            logger.info(f"  No index provided. Using default photo.")
            return default_path
        return None
    
    # Clean up the index
    clean_index = str(student_index).strip().replace('.0', '')
    logger.info(f"  Looking for photo with index: '{clean_index}'")
    
    # Try exact match first
    if clean_index in photos_dict:
        logger.info(f"  Found exact photo match for index '{clean_index}'")
        return photos_dict[clean_index]
    
    # Try without the FIRST leading 0 only (photo might have it removed)
    # Portal: 0310123001 -> Photo might be: 310123001
    if clean_index.startswith('0'):
        index_without_first_zero = clean_index[1:]  # Remove only first character
        logger.info(f"  Trying without first zero: '{index_without_first_zero}'")
        if index_without_first_zero in photos_dict:
            logger.info(f"  Found photo match for index '{clean_index}' (filename: {index_without_first_zero})")
            return photos_dict[index_without_first_zero]
    
    # Try adding leading 0 (in case index on portal doesn't have it but photo does)
    if not clean_index.startswith('0') and len(clean_index) == 9:
        index_with_zero = '0' + clean_index
        logger.info(f"  Trying with leading zero: '{index_with_zero}'")
        if index_with_zero in photos_dict:
            logger.info(f"  Found photo match for index '{clean_index}' (filename: {index_with_zero})")
            return photos_dict[index_with_zero]
    
    # Log available keys for debugging
    logger.warning(f"  No photo found for index '{clean_index}'.")
    logger.info(f"  Available photo keys (first 10): {list(photos_dict.keys())[:10]}")
    
    # Return default photo
    default_path = os.path.join(photos_folder, DEFAULT_PHOTO)
    if os.path.exists(default_path):
        logger.info(f"  Using default photo: {default_path}")
        return default_path
    
    logger.warning(f"  No photo found for index '{clean_index}' and default photo not available")
    return None

def find_photo_for_student(student_name, photos_dict, photos_folder):
    """
    Find the photo file for a given student name (legacy fallback).
    Returns the file path if found, otherwise returns the default photo path.
    """
    normalized_name = normalize_name(student_name)
    
    # Try exact match
    if normalized_name in photos_dict:
        logger.info(f"  Found exact photo match for '{student_name}'")
        return photos_dict[normalized_name]
    
    # Try partial matching
    name_parts = set(normalized_name.split())
    best_match = None
    best_score = 0
    
    for photo_name, photo_path in photos_dict.items():
        photo_parts = set(photo_name.split())
        
        if photo_parts and name_parts:
            common = name_parts.intersection(photo_parts)
            score = len(common) / max(len(name_parts), len(photo_parts))
            
            if score > best_score and score >= 0.6:  # At least 60% match
                best_score = score
                best_match = photo_path
    
    if best_match:
        logger.info(f"  Found fuzzy photo match for '{student_name}' (score: {best_score:.2f})")
        return best_match
    
    # Return default photo
    default_path = os.path.join(photos_folder, DEFAULT_PHOTO)
    if os.path.exists(default_path):
        logger.info(f"  Using default photo for '{student_name}'")
        return default_path
    
    logger.warning(f"  No photo found for '{student_name}' and default photo not available")
    return None

# ==========================================
# NAME NORMALIZATION FUNCTIONS
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

def find_student_by_index(student_index, students_data):
    """
    Find a student record by their StudentIndex.
    Returns the student data dict if found, otherwise returns None.
    """
    if not student_index:
        return None
    
    # Clean up the index for matching
    clean_index = str(student_index).strip().replace('.0', '')
    
    # Direct match
    if clean_index in students_data:
        return students_data[clean_index]
    
    # Try without leading zeros
    try:
        numeric_index = str(int(clean_index))
        if numeric_index in students_data:
            return students_data[numeric_index]
    except:
        pass
    
    return None

def find_contact_for_student(student_name, students_data):
    """
    Find the parent contact for a given student by name (fallback).
    Returns the contact if found, otherwise returns None.
    """
    normalized_name = normalize_name(student_name)
    
    # Search through students_data by name
    for index, data in students_data.items():
        if normalize_name(data.get('name', '')) == normalized_name:
            return data.get('contact', '')
    
    # Try partial matching
    name_parts = set(normalized_name.split())
    best_match = None
    best_score = 0
    
    for index, data in students_data.items():
        student_name_parts = set(normalize_name(data.get('name', '')).split())
        
        if student_name_parts and name_parts:
            common = name_parts.intersection(student_name_parts)
            score = len(common) / max(len(name_parts), len(student_name_parts))
            
            if score > best_score and score >= 0.7:  # At least 70% match
                best_score = score
                best_match = data.get('contact', '')
    
    return best_match

# ==========================================
# MAIN EXECUTION
# ==========================================
def run_automation():
    # Parse Excel for student data (StudentIndex, StudentName, Contact)
    students_data = parse_parent_contacts_excel(EXCEL_FILE_PATH)
    
    # Get photo files
    photos_dict = get_photo_files(PHOTOS_FOLDER)
    
    # Counter for default contact numbers
    default_contact_counter = DEFAULT_CONTACT_START
    
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

            # Debug: List all inputs found
            try:
                inputs = page.query_selector_all("input")
                logger.info(f"Debug: Found {len(inputs)} input fields on login page.")
            except:
                pass

            # ATTEMPT AUTO-LOGIN
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
            logger.info("Waiting for login success...")
            
            max_retries = 30
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

        # 2. NAVIGATE TO STUDENT REGISTRATION PAGE
        logger.info("Navigating to Student Registration page...")
        page.goto(STUDENT_DETAILS_URL)
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        logger.info(f"Navigated to: {STUDENT_DETAILS_URL}")
        
        processed_count = 0
        skipped_count = 0
        max_iterations = 500  # Safety limit
        
        for iteration in range(max_iterations):
            try:
                logger.info(f"\n--- Iteration {iteration + 1} ---")
                
                # FIRST: READ STUDENT INDEX FROM PORTAL (Basic School IndexNumber)
                student_index_on_portal = ""
                
                # Look for "Basic School IndexNumber" field
                try:
                    index_patterns = [
                        "Basic School IndexNumber",
                        "Basic School Index",
                        "IndexNumber",
                        "Index Number",
                        "Index No",
                    ]
                    for pattern in index_patterns:
                        try:
                            # Try to find the label and get the value from input or text
                            label_el = page.locator(f"text={pattern}").first
                            if label_el.count() > 0:
                                # Look for input field nearby
                                parent = label_el.locator("xpath=./parent::*")
                                if parent.count() > 0:
                                    # Check for input in parent or sibling
                                    input_el = parent.locator("input").first
                                    if input_el.count() > 0:
                                        student_index_on_portal = input_el.input_value().strip()
                                        if student_index_on_portal:
                                            break
                                    # Check for text value
                                    parent_text = parent.inner_text()
                                    if pattern in parent_text:
                                        student_index_on_portal = parent_text.split(pattern)[-1].strip()
                                        student_index_on_portal = student_index_on_portal.split('\n')[0].strip()
                                        # Remove any colons
                                        student_index_on_portal = student_index_on_portal.lstrip(':').strip()
                                        if student_index_on_portal:
                                            break
                        except:
                            pass
                    
                    # Also try finding input by name/id
                    if not student_index_on_portal:
                        index_input_selectors = [
                            "input[name*='Index' i]",
                            "input[name*='index' i]",
                            "input[id*='Index' i]",
                            "input[id*='BasicSchool' i]",
                        ]
                        for sel in index_input_selectors:
                            try:
                                if page.locator(sel).count() > 0:
                                    val = page.locator(sel).first.input_value()
                                    if val and val.strip():
                                        student_index_on_portal = val.strip()
                                        break
                            except:
                                pass
                                
                except Exception as e:
                    logger.warning(f"Error extracting student index: {e}")
                
                logger.info(f"  Student Index on portal: {student_index_on_portal}")
                
                # READ STUDENT NAME FROM PORTAL PAGE
                full_name_text = ""
                
                # Method 1: Look for input/text field with name-related labels (similar to index extraction)
                try:
                    name_label_patterns = [
                        "Full Name",
                        "Student Name",
                        "Fullname",
                        "StudentName",
                    ]
                    for pattern in name_label_patterns:
                        try:
                            # Try to find label and get value from nearby input
                            label_el = page.locator(f"text={pattern}").first
                            if label_el.count() > 0:
                                parent = label_el.locator("xpath=./parent::*")
                                if parent.count() > 0:
                                    # Check for input in parent
                                    input_el = parent.locator("input").first
                                    if input_el.count() > 0:
                                        full_name_text = input_el.input_value().strip()
                                        if full_name_text:
                                            break
                                    # Check for text content (might be displayed as text, not input)
                                    parent_text = parent.inner_text()
                                    if pattern in parent_text:
                                        full_name_text = parent_text.replace(pattern, '').strip()
                                        full_name_text = full_name_text.lstrip(':').strip()
                                        full_name_text = full_name_text.split('\n')[0].strip()
                                        if full_name_text:
                                            break
                        except:
                            pass
                except Exception as e:
                    logger.warning(f"Name extraction method 1 failed: {e}")
                
                # Method 2: Try finding input fields by name/id attributes
                if not full_name_text:
                    try:
                        name_input_selectors = [
                            "input[name*='FullName' i]",
                            "input[name*='fullname' i]",
                            "input[name*='StudentName' i]",
                            "input[name*='Name' i]",
                            "input[id*='FullName' i]",
                            "input[id*='fullname' i]",
                            "input[id*='StudentName' i]",
                            "input[id*='Name' i]",
                        ]
                        for sel in name_input_selectors:
                            try:
                                if page.locator(sel).count() > 0:
                                    val = page.locator(sel).first.input_value()
                                    if val and val.strip() and len(val.strip()) > 2:
                                        full_name_text = val.strip()
                                        logger.info(f"  Found name via input selector: {sel}")
                                        break
                            except:
                                pass
                    except Exception as e:
                        logger.warning(f"Name extraction method 2 failed: {e}")
                
                # Method 3: Look for specific element patterns via xpath
                if not full_name_text:
                    try:
                        xpath_patterns = [
                            "//label[contains(text(),'Full Name')]/following-sibling::input",
                            "//label[contains(text(),'Name')]/following-sibling::input",
                            "//span[contains(text(),'Full Name')]/following-sibling::*",
                            "//td[contains(text(),'Full Name')]/following-sibling::td",
                            "//b[contains(text(),'Full Name')]/parent::*",
                        ]
                        for pattern in xpath_patterns:
                            el = page.locator(f"xpath={pattern}").first
                            if el.count() > 0:
                                # Try to get input value first
                                try:
                                    val = el.input_value()
                                    if val and val.strip():
                                        full_name_text = val.strip()
                                        break
                                except:
                                    pass
                                # Try text content
                                text = el.inner_text().strip()
                                if text and not text.lower().startswith('full name'):
                                    full_name_text = text.split('\n')[0].strip()
                                    if full_name_text:
                                        break
                    except Exception as e:
                        logger.warning(f"Name extraction method 3 failed: {e}")
                
                # Method 4: Direct text extraction from page content
                if not full_name_text:
                    try:
                        page_content = page.content()
                        match = re.search(r'(?:Full Name|Student Name)[:\s]*</[^>]+>\s*([A-Z][A-Za-z\s]+)', page_content)
                        if match:
                            full_name_text = match.group(1).strip()
                    except Exception as e:
                        logger.warning(f"Name extraction method 4 failed: {e}")
                
                # If still no name, we can proceed with just the index since that's more important
                if not full_name_text:
                    if student_index_on_portal:
                        logger.warning("Could not extract student name, but have index. Proceeding with index only.")
                        full_name_text = f"Student_{student_index_on_portal}"
                    else:
                        logger.warning("Could not extract student name from portal page.")
                        full_name_text = input(">>> Enter the student's FULL NAME as shown on the portal (or 'skip' to skip, 'quit' to stop): ").strip()
                        if full_name_text.lower() == 'quit':
                            break
                    if full_name_text.lower() == 'skip':
                        # Try to navigate to next student
                        try:
                            if page.locator("text=Next").count() > 0:
                                page.click("text=Next")
                                time.sleep(2)
                        except:
                            pass
                        continue
                
                logger.info(f"  Student on portal: {full_name_text}")
                
                # CHECK IF DATA IS ALREADY FILLED (skip if already done)
                already_filled = False
                try:
                    # Check if contact field already has a value
                    contact_selectors = [
                        "input[name*='contact' i]",
                        "input[name*='phone' i]",
                        "input[name*='parent' i]",
                        "input[name*='guardian' i]",
                        "input[placeholder*='contact' i]",
                        "input[placeholder*='phone' i]",
                    ]
                    for sel in contact_selectors:
                        try:
                            if page.locator(sel).count() > 0:
                                val = page.locator(sel).first.input_value()
                                if val and val.strip() and len(val.strip()) >= 10:
                                    already_filled = True
                                    logger.info(f"  SKIPPING: Contact already filled for '{full_name_text}'")
                                    skipped_count += 1
                                    break
                        except:
                            pass
                except Exception as e:
                    logger.warning(f"  Error checking if already filled: {e}")
                
                if already_filled:
                    # Navigate to next student - clicking SAVE STUDENT loads the next one
                    try:
                        save_selectors = [
                            "button:has-text('SAVE STUDENT')",
                            "text=SAVE STUDENT",
                            ".btn:has-text('SAVE STUDENT')",
                        ]
                        for sel in save_selectors:
                            if page.locator(sel).count() > 0:
                                page.locator(sel).first.click()
                                time.sleep(2)
                                page.wait_for_load_state("networkidle")
                                break
                    except:
                        pass
                    continue
                
                # FIND STUDENT IN EXCEL BY INDEX (primary method)
                matched_student = None
                parent_contact = None
                
                if student_index_on_portal:
                    matched_student = find_student_by_index(student_index_on_portal, students_data)
                    if matched_student:
                        parent_contact = matched_student.get('contact', '')
                        logger.info(f"  Matched by Index: {student_index_on_portal} -> {matched_student.get('name', 'Unknown')}")
                        if parent_contact:
                            logger.info(f"  Found contact: {parent_contact}")
                
                # Fallback: Find by name if index match failed or no contact
                if not parent_contact and full_name_text:
                    parent_contact = find_contact_for_student(full_name_text, students_data)
                    if parent_contact:
                        logger.info(f"  Found contact by name match: {parent_contact}")
                
                if not parent_contact:
                    # Use default contact with incrementing number
                    parent_contact = f"0{default_contact_counter}"
                    logger.info(f"  No contact found in Excel. Using default: {parent_contact}")
                    default_contact_counter += 1
                
                # FIND PHOTO BY INDEX NUMBER (photos are named by index, may be missing leading 0)
                photo_path = find_photo_by_index(student_index_on_portal, photos_dict, PHOTOS_FOLDER)
                
                # FILL PARENT/GUARDIAN CONTACT
                logger.info("  Filling Parent/Guardian Contact...")
                contact_filled = False
                contact_selectors = [
                    "input[name*='Contact' i]",
                    "input[name*='contact' i]",
                    "input[name*='Phone' i]",
                    "input[name*='phone' i]",
                    "input[name*='Parent' i]",
                    "input[name*='parent' i]",
                    "input[name*='Guardian' i]",
                    "input[name*='guardian' i]",
                    "input[id*='Contact' i]",
                    "input[id*='contact' i]",
                    "input[id*='Phone' i]",
                    "input[id*='phone' i]",
                    "input[id*='Parent' i]",
                    "input[id*='Guardian' i]",
                    "input[placeholder*='Contact' i]",
                    "input[placeholder*='contact' i]",
                ]
                
                for sel in contact_selectors:
                    try:
                        locator = page.locator(sel)
                        if locator.count() > 0:
                            locator.first.fill(parent_contact)
                            contact_filled = True
                            logger.info(f"    Filled contact: {parent_contact} using selector: {sel}")
                            break
                    except Exception as e:
                        logger.debug(f"    Selector {sel} failed: {e}")
                        continue
                
                if not contact_filled:
                    # Try to find ALL text inputs and log them for debugging
                    logger.warning("    Could not find contact field with standard selectors.")
                    try:
                        all_inputs = page.locator("input[type='text'], input[type='tel'], input[type='number']").all()
                        logger.info(f"    Found {len(all_inputs)} text/tel/number inputs on page:")
                        for i, inp in enumerate(all_inputs[:10]):  # Log first 10
                            try:
                                name = inp.get_attribute("name") or "no-name"
                                id_attr = inp.get_attribute("id") or "no-id"
                                placeholder = inp.get_attribute("placeholder") or "no-placeholder"
                                logger.info(f"      Input {i}: name='{name}', id='{id_attr}', placeholder='{placeholder}'")
                            except:
                                pass
                    except Exception as e:
                        logger.warning(f"    Error listing inputs: {e}")
                    
                    # Try to find by label
                    try:
                        label_texts = ["Parent", "Guardian", "Contact", "Phone"]
                        for label_text in label_texts:
                            label = page.locator(f"label:has-text('{label_text}')").first
                            if label.count() > 0:
                                # Find associated input
                                for_attr = label.get_attribute("for")
                                if for_attr:
                                    page.locator(f"#{for_attr}").fill(parent_contact)
                                    contact_filled = True
                                    break
                                else:
                                    # Try sibling input
                                    sibling_input = label.locator("xpath=./following-sibling::input | ./parent::*/input")
                                    if sibling_input.count() > 0:
                                        sibling_input.first.fill(parent_contact)
                                        contact_filled = True
                                        break
                            if contact_filled:
                                break
                    except Exception as e:
                        logger.warning(f"    Error finding contact field by label: {e}")
                
                if not contact_filled:
                    logger.warning("    Could not find contact input field. Please fill manually.")
                    input(">>> Fill the contact field manually, then press Enter to continue...")
                
                # SELECT DISABILITY STATUS (dropdown - "None")
                logger.info("  Selecting Disability Status...")
                disability_filled = False
                disability_selectors = [
                    "select[name*='disability' i]",
                    "select[name*='Disability' i]",
                    "select[id*='disability' i]",
                    "select[id*='Disability' i]",
                ]
                
                for sel in disability_selectors:
                    try:
                        if page.locator(sel).count() > 0:
                            # Try to select "None" by various methods
                            select_el = page.locator(sel).first
                            
                            # Try selecting by label/text
                            try:
                                select_el.select_option(label=DEFAULT_DISABILITY_STATUS)
                                disability_filled = True
                                logger.info(f"    Selected disability status: {DEFAULT_DISABILITY_STATUS}")
                                break
                            except:
                                pass
                            
                            # Try selecting by value
                            try:
                                select_el.select_option(value="None")
                                disability_filled = True
                                break
                            except:
                                pass
                            
                            # Try selecting by index (often "None" is first option)
                            try:
                                select_el.select_option(index=0)
                                disability_filled = True
                                break
                            except:
                                pass
                    except Exception as e:
                        continue
                
                if not disability_filled:
                    # Try to find by label text
                    try:
                        label = page.locator("label:has-text('Disability')").first
                        if label.count() > 0:
                            for_attr = label.get_attribute("for")
                            if for_attr:
                                select_el = page.locator(f"#{for_attr}")
                                if select_el.count() > 0:
                                    select_el.select_option(label=DEFAULT_DISABILITY_STATUS)
                                    disability_filled = True
                    except Exception as e:
                        logger.warning(f"    Error finding disability dropdown: {e}")
                
                if not disability_filled:
                    # Try generic select elements
                    try:
                        all_selects = page.locator("select").all()
                        for select_el in all_selects:
                            # Check if any option contains "None" or disability-related text
                            options_text = select_el.inner_text().lower()
                            if 'none' in options_text or 'disability' in options_text:
                                try:
                                    select_el.select_option(label=DEFAULT_DISABILITY_STATUS)
                                    disability_filled = True
                                    logger.info("    Found and filled disability dropdown")
                                    break
                                except:
                                    try:
                                        select_el.select_option(index=0)
                                        disability_filled = True
                                        break
                                    except:
                                        pass
                    except:
                        pass
                
                if not disability_filled:
                    logger.warning("    Could not find disability dropdown. Please select manually.")
                    input(">>> Select 'None' for Disability Status manually, then press Enter to continue...")
                
                # UPLOAD PHOTO
                logger.info("  Uploading Photo...")
                photo_uploaded = False
                
                if photo_path and os.path.exists(photo_path):
                    # Convert to absolute path
                    photo_path = os.path.abspath(photo_path)
                    logger.info(f"    Photo to upload: {photo_path}")
                    
                    # Find file input - for "Choose File" button with input[type='file']
                    # The screenshot shows a standard HTML file input with "Choose File" button
                    file_input_selectors = [
                        "input[type='file']",
                        "input[accept*='image']",
                        "input[name*='photo' i]",
                        "input[name*='Photo' i]",
                        "input[name*='picture' i]",
                        "input[name*='image' i]",
                        "input[id*='photo' i]",
                        "input[id*='Photo' i]",
                        "input[id*='picture' i]",
                        "input[id*='student' i][type='file']",
                    ]
                    
                    for sel in file_input_selectors:
                        try:
                            file_input = page.locator(sel).first
                            if file_input.count() > 0:
                                # Set the file directly on the input element
                                file_input.set_input_files(photo_path)
                                photo_uploaded = True
                                logger.info(f"    Photo uploaded using selector: {sel}")
                                time.sleep(1)  # Wait for upload to process
                                break
                        except Exception as e:
                            logger.debug(f"    Selector {sel} failed: {e}")
                            continue
                    
                    if not photo_uploaded:
                        # Alternative: Try using file chooser dialog
                        try:
                            # Look for the "Choose File" button area and trigger file chooser
                            with page.expect_file_chooser(timeout=5000) as fc_info:
                                # Click on any file input or its label
                                file_input = page.locator("input[type='file']").first
                                if file_input.count() > 0:
                                    file_input.click()
                            file_chooser = fc_info.value
                            file_chooser.set_files(photo_path)
                            photo_uploaded = True
                            logger.info("    Photo uploaded via file chooser dialog")
                        except Exception as e:
                            logger.warning(f"    Error with file chooser: {e}")
                
                if not photo_uploaded:
                    logger.warning("    Could not upload photo automatically. Please upload manually.")
                    if photo_path:
                        logger.info(f"    Photo file: {photo_path}")
                    input(">>> Upload the photo manually, then press Enter to continue...")
                
                # SAVE
                logger.info("  Saving...")
                try:
                    save_selectors = [
                        "button:has-text('SAVE STUDENT')",
                        "text=SAVE STUDENT",
                        ".btn:has-text('SAVE STUDENT')",
                        "a:has-text('SAVE STUDENT')",
                        "button:has-text('Save Student')",
                        "button:has-text('Save')",
                        "button:has-text('SAVE')",
                        "input[type='submit']",
                        "input[value='Save']",
                        ".btn:has-text('Save')",
                        "button[type='submit']",
                    ]
                    
                    save_clicked = False
                    for sel in save_selectors:
                        try:
                            if page.locator(sel).count() > 0 and page.locator(sel).first.is_visible():
                                page.locator(sel).first.click()
                                save_clicked = True
                                logger.info(f"    Clicked save button")
                                break
                        except:
                            continue
                    
                    if not save_clicked:
                        logger.warning("    Could not find save button. Please save manually.")
                        input(">>> Click the Save button manually, then press Enter to continue...")
                        
                except Exception as e:
                    logger.warning(f"  Error clicking save: {e}")
                    input(">>> Please save manually, then press Enter to continue...")
                        
                # Wait for save to complete
                time.sleep(2)
                page.wait_for_load_state("networkidle")
                time.sleep(1)
                
                # Click "NEXT STUDENT" button to move to next student
                logger.info("  Clicking NEXT STUDENT button...")
                next_clicked = False
                next_selectors = [
                    "button:has-text('NEXT STUDENT')",
                    "text=NEXT STUDENT",
                    ".btn:has-text('NEXT STUDENT')",
                    "a:has-text('NEXT STUDENT')",
                    "button:has-text('Next Student')",
                    "text=Next Student",
                ]
                
                for sel in next_selectors:
                    try:
                        if page.locator(sel).count() > 0 and page.locator(sel).first.is_visible():
                            page.locator(sel).first.click()
                            next_clicked = True
                            logger.info(f"    Clicked NEXT STUDENT button")
                            break
                    except:
                        continue
                
                if not next_clicked:
                    logger.warning("    Could not find NEXT STUDENT button.")
                    input(">>> Click the NEXT STUDENT button manually, then press Enter to continue...")
                
                # Wait for next student page to load
                time.sleep(2)
                page.wait_for_load_state("networkidle")
                time.sleep(1)
                
                # Click "REGISTER STUDENT" button to open the form for the next student
                logger.info("  Clicking REGISTER STUDENT button...")
                register_clicked = False
                register_selectors = [
                    "button:has-text('REGISTER STUDENT')",
                    "text=REGISTER STUDENT",
                    ".btn:has-text('REGISTER STUDENT')",
                    "a:has-text('REGISTER STUDENT')",
                    "button:has-text('Register Student')",
                    "text=Register Student",
                ]
                
                for sel in register_selectors:
                    try:
                        if page.locator(sel).count() > 0 and page.locator(sel).first.is_visible():
                            page.locator(sel).first.click()
                            register_clicked = True
                            logger.info(f"    Clicked REGISTER STUDENT button")
                            break
                    except:
                        continue
                
                if not register_clicked:
                    logger.warning("    Could not find REGISTER STUDENT button.")
                    input(">>> Click the REGISTER STUDENT button manually, then press Enter to continue...")
                
                # Wait for registration form to load
                time.sleep(2)
                page.wait_for_load_state("networkidle")
                time.sleep(1)
                
                # Check if we moved to next student by reading the new index
                new_index = ""
                try:
                    index_input_selectors = [
                        "input[name*='Index' i]",
                        "input[id*='Index' i]",
                    ]
                    for sel in index_input_selectors:
                        if page.locator(sel).count() > 0:
                            new_index = page.locator(sel).first.input_value().strip()
                            if new_index:
                                break
                except:
                    pass
                
                # If same student, something went wrong
                if new_index == student_index_on_portal:
                    logger.warning(f"  Still on same student after navigation. Waiting longer...")
                    time.sleep(3)
                    page.wait_for_load_state("networkidle")
                    
                    # Check again
                    try:
                        for sel in index_input_selectors:
                            if page.locator(sel).count() > 0:
                                new_index = page.locator(sel).first.input_value().strip()
                                if new_index:
                                    break
                    except:
                        pass
                    
                    if new_index == student_index_on_portal:
                        logger.warning(f"  Page did not advance to next student!")
                        logger.info(f"  Current index still: {new_index}")
                        action = input(">>> The page didn't move to next student. Press Enter to retry, 'skip' to skip, 'quit' to stop: ").strip()
                        if action.lower() == 'quit':
                            break
                        elif action.lower() == 'skip':
                            # Try to find and click a Next button if exists
                            try:
                                next_btns = ["text=Next", "button:has-text('Next')", "a:has-text('Next')"]
                                for btn in next_btns:
                                    if page.locator(btn).count() > 0:
                                        page.locator(btn).first.click()
                                        time.sleep(2)
                                        break
                            except:
                                pass
                            continue
                        else:
                            continue
                
                processed_count += 1
                logger.info(f"  Successfully processed: {full_name_text}")
                
                # Check if we're at the end
                try:
                    if page.locator("text=No more students").count() > 0 or \
                       page.locator("text=All students processed").count() > 0 or \
                       page.locator("text=No records").count() > 0:
                        logger.info("  No more students to process.")
                        break
                except:
                    pass
                
                logger.info(f"  Moving to next student... (new index: {new_index})")

            except Exception as e:
                logger.error(f"  Error in iteration {iteration + 1}: {e}")
                action = input(">>> Error occurred. Press Enter to continue, 'quit' to stop: ").strip()
                if action.lower() == 'quit':
                    break

        print("\n" + "="*40)
        logger.info(f"Processing completed!")
        logger.info(f"  Processed: {processed_count} students")
        logger.info(f"  Skipped (already filled): {skipped_count} students")
        logger.info(f"  Last default contact used: 0{default_contact_counter - 1}")
        input("Press Enter to close the browser and exit script...")
        browser.close() 

if __name__ == "__main__":
    run_automation()
