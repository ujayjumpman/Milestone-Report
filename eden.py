import os
import logging
from io import BytesIO
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv
import ibm_boto3
from ibm_botocore.client import Config
import re

# ======================= CONFIG =======================
load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

COS_API_KEY    = os.getenv("COS_API_KEY")
COS_CRN        = os.getenv("COS_SERVICE_INSTANCE_CRN")
COS_ENDPOINT   = os.getenv("COS_ENDPOINT")
BUCKET         = os.getenv("COS_BUCKET_NAME")
KRA_KEY        = os.getenv("KRA_FILE_PATH")          # EDEN Targets Till August 2025.xlsx
TRACKER_KEY    = os.getenv("EDEN_TRACKER_PATH")      # Eden/Structure Work Tracker (01-07-2025).xlsx

# Fixed cell/row positions (THESE CAN BE HARDCODED as per requirement)
KRA_PARENT_ROW = {
    "Tower 4": [5, 6],    # B5, B6 (Upper basement, beam/slab etc.)
    "Tower 5": [8, 9],    # B8, B9 
    "Tower 6": [11, 12],  # B11, B12
    "Tower 7": [14, 15],  # B14, B15
    "NTA-01": [17, 18],   # B17, B18
    "NTA-02": [20, 21],   # B20, B21
}

KRA_ACTIVITY_ROW = {
    "Tower 4": 7,     # B7 - Child activity for Tower 4
    "Tower 5": 10,    # B10 - Child activity for Tower 5  
    "Tower 6": 13,    # B13 - Child activity for Tower 6
    "Tower 7": 16,    # B16 - Child activity for Tower 7
    "NTA-01": 19,     # B19 - Child activity for NTA-01
    "NTA-02": 22,     # B22 - Child activity for NTA-02
}

# Fixed column positions in tracker sheet (THESE CAN BE HARDCODED)
TASK_NAME_COL = 4  # D column (Task Name)
PCT_COL = 7        # G column (% Complete) - PRIMARY
PCT_COL_ALT = [6, 8, 9, 10, 5]  # Alternative percentage columns to check
RESPONSIBLE_COL = 6  # F column (Responsible Person)
DELAY_COL = 8        # H column (Delay Reasons)

# ============= DYNAMIC DISCOVERY FUNCTIONS ==================
def discover_months_and_columns(kra_ws):
    """Dynamically discover available months and their column positions from KRA sheet headers"""
    months_found = {}
    
    # Check first few rows for month headers (typically in row 1 or 2)
    for row in range(1, 5):  # Check first 4 rows
        for col in range(1, 25):  # Increased range to 25 columns
            cell_value = kra_ws.cell(row=row, column=col).value
            if cell_value:
                cell_str = str(cell_value).strip()
                # Look for month names (case insensitive) with year patterns
                month_patterns = {
                    'january': ['january', 'jan'], 'february': ['february', 'feb'], 
                    'march': ['march', 'mar'], 'april': ['april', 'apr'], 
                    'may': ['may'], 'june': ['june', 'jun'], 
                    'july': ['july', 'jul'], 'august': ['august', 'aug'], 
                    'september': ['september', 'sep', 'sept'], 'october': ['october', 'oct'], 
                    'november': ['november', 'nov'], 'december': ['december', 'dec']
                }
                
                cell_lower = cell_str.lower()
                for full_month, patterns in month_patterns.items():
                    for pattern in patterns:
                        if pattern in cell_lower:
                            # Use the full month name consistently
                            month_name = full_month.capitalize()
                            if month_name not in months_found:  # Avoid duplicates
                                months_found[month_name] = col
                                logger.info(f"Found month '{month_name}' in column {col} (original: '{cell_str}')")
                            break
                    if full_month.capitalize() in months_found:
                        break
    
    return months_found

def discover_current_month(tracker_filename):
    """Dynamically determine current month from tracker filename or latest data"""
    # Extract date from filename if present
    date_pattern = r'(\d{2}-\d{2}-\d{4})'
    match = re.search(date_pattern, tracker_filename)
    
    if match:
        date_str = match.group(1)
        try:
            file_date = datetime.strptime(date_str, "%d-%m-%Y")
            current_month = file_date.strftime("%B").capitalize()  # Full month name
            logger.info(f"Extracted current month '{current_month}' from tracker filename")
            return current_month
        except ValueError as e:
            logger.warning(f"Could not parse date from filename: {e}")
    
    # Fallback to current system date
    current_month = datetime.now().strftime("%B").capitalize()
    logger.info(f"Using system current month: {current_month}")
    return current_month

def discover_towers(kra_ws):
    """Dynamically discover available towers from KRA sheet - FIXED to be more precise"""
    towers_found = []
    
    # Look for tower names in the first few columns (typically column A or B)
    for col in range(1, 5):  # Check first few columns
        for row in range(1, 50):  # Check first 50 rows
            cell_value = kra_ws.cell(row=row, column=col).value
            if cell_value:
                cell_str = str(cell_value).strip()
                
                # Look for tower patterns - be more specific
                tower_match = re.match(r'^Tower\s*(\d+)$', cell_str, re.IGNORECASE)
                if tower_match:
                    tower_num = tower_match.group(1)
                    tower_name = f"Tower {tower_num}"
                    if tower_name not in towers_found:
                        towers_found.append(tower_name)
                        logger.info(f"Found tower: {tower_name} at row {row}, col {col}")
                
                # Look for NTA patterns - be more specific
                nta_match = re.match(r'^NTA[-\s]*(\d+)$', cell_str, re.IGNORECASE)
                if nta_match:
                    nta_num = nta_match.group(1)
                    nta_name = f"NTA-{nta_num.zfill(2)}"  # Ensure 2-digit format
                    if nta_name not in towers_found:
                        towers_found.append(nta_name)
                        logger.info(f"Found NTA: {nta_name} at row {row}, col {col}")
    
    return towers_found

def discover_tracker_sheets(tracker_wb):
    """Dynamically discover and map tracker sheets - FIXED"""
    sheet_mapping = {}
    
    for sheet_name in tracker_wb.sheetnames:
        sheet_name_clean = sheet_name.strip()
        
        # Map tower sheets - be more specific
        tower_match = re.search(r'Tower\s*(\d+)', sheet_name_clean, re.IGNORECASE)
        if tower_match:
            tower_number = tower_match.group(1)
            tower_key = f"Tower {tower_number}"
            sheet_mapping[tower_key] = sheet_name_clean
            logger.info(f"Mapped {tower_key} to sheet '{sheet_name_clean}'")
        
        # Map NTA sheets (usually named "Non Tower Area" or similar)
        elif re.search(r'non.*tower.*area', sheet_name_clean, re.IGNORECASE):
            # Both NTA-01 and NTA-02 typically map to the same "Non Tower Area" sheet
            sheet_mapping["NTA-01"] = sheet_name_clean
            sheet_mapping["NTA-02"] = sheet_name_clean
            logger.info(f"Mapped NTA areas to sheet '{sheet_name_clean}'")
    
    return sheet_mapping

def debug_tracker_sheet_structure(tracker_ws, tower_name):
    """Debug function to understand tracker sheet structure and find correct percentage columns"""
    logger.info(f"\n=== DEBUGGING TRACKER SHEET STRUCTURE FOR {tower_name} ===")
    
    # Check first 5 rows for headers
    logger.info("Header rows (1-5):")
    for row in range(1, 6):
        row_data = []
        for col in range(1, 15):  # Check first 15 columns
            cell_val = tracker_ws.cell(row=row, column=col).value
            if cell_val:
                row_data.append(f"Col{col}:{cell_val}")
        if row_data:
            logger.info(f"  Row {row}: {' | '.join(row_data)}")
    
    # Find some sample data rows to understand structure
    logger.info("\nSample data rows:")
    for row in range(10, min(35, tracker_ws.max_row + 1)):
        task_val = tracker_ws.cell(row=row, column=TASK_NAME_COL).value
        if task_val and str(task_val).strip():
            row_data = [f"Row{row}"]
            for col in range(4, 12):  # Columns D to K
                cell_val = tracker_ws.cell(row=row, column=col).value
                if cell_val is not None:
                    row_data.append(f"Col{col}:{cell_val}")
            logger.info(f"  {' | '.join(row_data)}")
            if len([r for r in range(10, row) if tracker_ws.cell(row=r, column=TASK_NAME_COL).value]) > 10:
                break  # Show only first 10 data rows
    
    logger.info(f"=== END DEBUG FOR {tower_name} ===\n")

def find_correct_percentage_column(tracker_ws, row, task_name):
    """Find the correct percentage column for a specific task by checking multiple columns"""
    percentage_candidates = []
    
    # Check multiple columns that might contain percentages
    check_columns = [PCT_COL] + PCT_COL_ALT
    
    for col in check_columns:
        try:
            cell_val = tracker_ws.cell(row=row, column=col).value
            if cell_val is not None:
                cell_str = str(cell_val).strip()
                
                # Check if this looks like a percentage
                if cell_str.endswith('%'):
                    pct_str = cell_str.replace('%', '').strip()
                    if pct_str.replace('.', '').replace('-', '').isdigit():
                        percentage_candidates.append((col, cell_val, 'percentage_symbol'))
                elif cell_str.replace('.', '').replace('-', '').isdigit():
                    num_val = float(cell_str)
                    if 0 <= num_val <= 100:
                        percentage_candidates.append((col, cell_val, 'number_0_100'))
                    elif 0 <= num_val <= 1:
                        percentage_candidates.append((col, cell_val, 'decimal_0_1'))
        except:
            continue
    
    if percentage_candidates:
        # Prefer columns with % symbol, then numbers 0-100, then decimals 0-1
        priority_order = ['percentage_symbol', 'number_0_100', 'decimal_0_1']
        for priority in priority_order:
            for col, val, type_found in percentage_candidates:
                if type_found == priority:
                    logger.info(f"Found percentage in column {col} for '{task_name}': {val} (type: {type_found})")
                    return col, val
    
    return None, None

def validate_expected_percentages(tower, extracted_pct):
    """Validate extracted percentages against expected values for debugging"""
    expected_values = {
        "Tower 4": 55,
        "Tower 5": 35, 
        "Tower 6": 60,
        "Tower 7": 0,
        "NTA-01": 0,
        "NTA-02": 0
    }
    
    if tower in expected_values:
        expected = expected_values[tower]
        if abs(extracted_pct - expected) > 5:  # Allow 5% tolerance
            logger.warning(f"⚠️  PERCENTAGE MISMATCH for {tower}:")
            logger.warning(f"   Expected: {expected}%")
            logger.warning(f"   Extracted: {extracted_pct}%")
            logger.warning(f"   Difference: {abs(extracted_pct - expected)}%")
            return False
        else:
            logger.info(f"✅ PERCENTAGE VALIDATED for {tower}: {extracted_pct}% (expected: {expected}%)")
            return True
    return True

def alternative_percentage_search(tracker_ws, child_name, tower):
    """Alternative method to search for percentages when hierarchy method fails - with basement filtering for NTA"""
    logger.info(f"\n=== ALTERNATIVE PERCENTAGE SEARCH for {child_name} in {tower} ===")
    
    child_name_clean = str(child_name).strip().lower()
    max_row = tracker_ws.max_row
    
    # Check if this is an NTA search that needs basement-level filtering
    is_nta_search = tower.startswith('NTA')
    
    # For NTA searches, we should NOT use alternative search as it bypasses basement filtering
    if is_nta_search:
        logger.info(f"Skipping alternative search for NTA area '{tower}' to maintain basement-level filtering")
        return 0.0
    
    # Simple row-by-row search for the activity (non-NTA only)
    for row in range(2, max_row + 1):
        task_val = tracker_ws.cell(row=row, column=TASK_NAME_COL).value
        if task_val:
            task_clean = str(task_val).strip().lower()
            
            # Calculate match score
            match_score = calculate_enhanced_match_score(task_clean, child_name_clean)
            if match_score >= 0.95:  # Very high threshold for exact matching
                logger.info(f"Alternative search found match at row {row}: '{task_val}'")
                
                # Find correct percentage column
                correct_col, pct_val = find_correct_percentage_column(tracker_ws, row, task_val)
                if pct_val is not None:
                    try:
                        if isinstance(pct_val, (int, float)):
                            if 0 <= pct_val <= 1:
                                result = float(pct_val * 100)
                            elif 0 <= pct_val <= 100:
                                result = float(pct_val)
                            else:
                                result = float(pct_val)
                        else:
                            pct_str = str(pct_val).replace("%", "").replace(" ", "").strip()
                            if pct_str:
                                result = float(pct_str)
                                if 0 <= result <= 1:
                                    result = result * 100
                            else:
                                result = 0.0
                        
                        logger.info(f"Alternative search extracted: {pct_val} -> {result}%")
                        return result
                    except Exception as e:
                        logger.warning(f"Error in alternative parsing: {e}")
    
    logger.warning(f"Alternative search found no matches for '{child_name}'")
    return 0.0

def calculate_dynamic_weightage(tower, kra_ws, month_columns):
    """Dynamically calculate weightage based on activity complexity or data in sheets"""
    # For now, use a simple heuristic based on tower type
    # This could be enhanced to read from a specific cell in KRA sheet
    
    if tower.startswith("NTA"):
        return 50  # NTA areas typically have lower weightage
    else:
        return 100  # Main towers have full weightage

# ============= COS HELPERS ==================
def init_cos():
    return ibm_boto3.client(
        "s3",
        ibm_api_key_id=COS_API_KEY,
        ibm_service_instance_id=COS_CRN,
        config=Config(signature_version="oauth"),
        endpoint_url=COS_ENDPOINT,
    )

def download_file_bytes(cos, key):
    obj = cos.get_object(Bucket=BUCKET, Key=key)
    return obj["Body"].read()

def get_activity_for_month(tower, month, month_col, kra_ws):
    """Get the activity name for a specific tower and month from KRA file"""
    if tower not in KRA_ACTIVITY_ROW:
        return ""
        
    child_row = KRA_ACTIVITY_ROW[tower]
    child_name = kra_ws.cell(row=child_row, column=month_col).value
    
    if child_name and str(child_name).strip():
        return str(child_name).strip()
    return ""

def get_parent_activities_for_month(tower, month, month_col, kra_ws):
    """Get the parent activity names for a specific tower and month from KRA file"""
    if tower not in KRA_PARENT_ROW:
        return ""
        
    parent_rows = KRA_PARENT_ROW[tower]
    parent_names = []
    
    for parent_row in parent_rows:
        parent_name = kra_ws.cell(row=parent_row, column=month_col).value
        if parent_name and str(parent_name).strip():
            parent_names.append(str(parent_name).strip())
    
    # Join multiple parent names with " & "
    return " & ".join(parent_names) if parent_names else ""

def get_all_activities_for_month(tower, month, month_col, kra_ws):
    """Get all activities (parent + child) for a specific tower and month from KRA file - EXACT text from sheet"""
    if tower not in KRA_PARENT_ROW or tower not in KRA_ACTIVITY_ROW:
        return ""
    
    all_activities = []
    
    # Get parent activities - EXACT text from cells
    parent_rows = KRA_PARENT_ROW[tower]
    for parent_row in parent_rows:
        parent_name = kra_ws.cell(row=parent_row, column=month_col).value
        if parent_name and str(parent_name).strip():
            # Add exact text as it appears in the sheet
            all_activities.append(str(parent_name).strip())
    
    # Get child activity - EXACT text from cell
    child_row = KRA_ACTIVITY_ROW[tower]
    child_name = kra_ws.cell(row=child_row, column=month_col).value
    if child_name and str(child_name).strip():
        # Add exact text as it appears in the sheet
        all_activities.append(str(child_name).strip())
    
    # Format activities exactly as they appear, with line breaks between them
    return format_activities_exactly_from_sheet(all_activities)

def format_activities_exactly_from_sheet(activities):
    """Format activities exactly as they appear in KRA sheet, just adding line breaks"""
    if not activities:
        return ""
    
    # Simply join all activities with line breaks - no parsing or reconstruction
    formatted_activities = []
    
    for activity in activities:
        activity_text = str(activity).strip()
        if activity_text:
            formatted_activities.append(activity_text)
    
    # Join with newlines for multi-line display in Excel
    return '\n'.join(formatted_activities) if formatted_activities else ""

def get_tower_name_from_kra(tower, kra_ws):
    """Get the actual tower name from the KRA sheet instead of using milestone names"""
    # Look for tower name in the first column around the tower's row area
    if tower not in KRA_PARENT_ROW:
        return tower  # fallback to tower key
    
    # Check rows around the parent rows for tower name
    parent_rows = KRA_PARENT_ROW[tower]
    start_row = min(parent_rows) - 2  # Check a couple rows above
    end_row = max(parent_rows) + 2    # Check a couple rows below
    
    for row in range(max(1, start_row), end_row + 1):
        for col in range(1, 5):  # Check first few columns
            cell_value = kra_ws.cell(row=row, column=col).value
            if cell_value:
                cell_str = str(cell_value).strip()
                
                # Look for tower patterns that match our tower key
                if tower.startswith("Tower") and re.search(r'Tower\s*\d+', cell_str, re.IGNORECASE):
                    tower_match = re.search(r'Tower\s*(\d+)', cell_str, re.IGNORECASE)
                    key_match = re.search(r'Tower\s*(\d+)', tower, re.IGNORECASE)
                    if tower_match and key_match and tower_match.group(1) == key_match.group(1):
                        return cell_str.strip()
                
                elif tower.startswith("NTA") and re.search(r'NTA[-\s]*\d+', cell_str, re.IGNORECASE):
                    nta_match = re.search(r'NTA[-\s]*(\d+)', cell_str, re.IGNORECASE)
                    key_match = re.search(r'NTA[-\s]*(\d+)', tower, re.IGNORECASE)
                    if nta_match and key_match and nta_match.group(1) == key_match.group(1):
                        return cell_str.strip()
    
    # Fallback: return a cleaned version of the tower key
    return tower.replace("-", " ").title()

# ============= ENHANCED NTA LOGIC ==============
def validate_nta_section_by_row_range(section_start, section_end, required_basement_type, nta_number):
    """
    Precise row-based validation for NTA sections based on actual Excel structure
    """
    logger.info(f"Validating NTA-{nta_number} section rows {section_start}-{section_end} for {required_basement_type}")
    
    if nta_number == "01" and required_basement_type == "upper basement":
        # NTA-01 Upper Basement: rows 6-35 (NTA-01 area in the sheet)
        if 6 <= section_start <= 35:
            logger.info(f"✅ NTA-01 section validation PASSED: section at row {section_start} is in NTA-01 area (6-35)")
            return True
        else:
            logger.info(f"❌ NTA-01 section validation FAILED: section at row {section_start} is outside NTA-01 area (6-35)")
            return False
    
    elif nta_number == "02" and required_basement_type == "lower basement":
        # NTA-02 Lower Basement: rows 36+ (NTA-02 area in the sheet)
        # Based on the Excel sheet, NTA-02 starts around row 36 and goes down
        if section_start >= 36:  # Only accept sections starting from NTA-02 area (row 36+)
            logger.info(f"✅ NTA-02 section validation PASSED: section at row {section_start} is in NTA-02 area (36+)")
            return True
        else:
            logger.info(f"❌ NTA-02 section validation FAILED: section at row {section_start} is in NTA-01 area, not NTA-02 (should be 36+, got {section_start})")
            return False
    
    return True  # Default: allow section for non-NTA cases

def verify_nta_section_identity(tracker_ws, section_start, required_nta_number):
    """
    Verify that we're in the correct NTA section (NTA-01 vs NTA-02) by looking for identifiers
    """
    # Look backwards from the section to find NTA identifier
    for check_row in range(max(1, section_start - 10), section_start + 5):
        for col in range(1, 8):  # Check first few columns
            cell_val = tracker_ws.cell(row=check_row, column=col).value
            if cell_val:
                cell_str = str(cell_val).strip().upper()
                if f"NTA-{required_nta_number}" in cell_str or f"NTA {required_nta_number}" in cell_str:
                    logger.info(f"✅ Found NTA-{required_nta_number} identifier at row {check_row}: '{cell_val}'")
                    return True
                # Also check for wrong NTA to reject it
                wrong_nta = "01" if required_nta_number == "02" else "02"
                if f"NTA-{wrong_nta}" in cell_str or f"NTA {wrong_nta}" in cell_str:
                    logger.info(f"❌ Found wrong NTA identifier NTA-{wrong_nta} at row {check_row}: '{cell_val}'")
                    return False
    
    logger.info(f"⚠️  No clear NTA-{required_nta_number} identifier found near row {section_start}")
    return True  # If no identifier found, allow it (fallback)

def verify_all_parents_in_section(tracker_ws, base_row, required_parents, max_row):
    """
    Verify that all required parent activities are found in the section vicinity
    """
    found_parents = []
    
    # Check the base row and nearby bold rows (within ±5 rows)
    for check_row in range(max(2, base_row - 5), min(base_row + 6, max_row + 1)):
        check_val = tracker_ws.cell(row=check_row, column=TASK_NAME_COL).value
        if check_val:
            check_val_clean = str(check_val).strip().lower()
            try:
                check_font = tracker_ws.cell(row=check_row, column=TASK_NAME_COL).font
                check_bold = check_font and check_font.bold
                if check_bold:
                    found_parents.append(check_val_clean)
            except:
                pass
    
    # Verify all required parents are present
    for required_parent in required_parents:
        parent_found_in_section = any(
            required_parent in found_parent or 
            found_parent in required_parent or
            enhanced_text_matching(required_parent, found_parent)
            for found_parent in found_parents
        )
        if not parent_found_in_section:
            return False
    
    return True

def find_exact_child_in_section(tracker_ws, start_row, end_row, child_name_clean):
    """Find exact child activity within a specific parent section with improved matching"""
    
    logger.info(f"Scanning rows {start_row} to {end_row} for exact child activity")
    
    best_match_row = None
    best_match_score = 0
    match_threshold = 0.95  # Very high threshold for exact matching
    
    for row in range(start_row, end_row + 1):
        task_val = tracker_ws.cell(row=row, column=TASK_NAME_COL).value
        
        if task_val is None or str(task_val).strip() == "":
            continue
        
        # Skip if this is a bold row (another parent)
        try:
            font = tracker_ws.cell(row=row, column=TASK_NAME_COL).font
            is_bold = font and font.bold
            if is_bold:
                continue
        except:
            pass
        
        task_val_clean = str(task_val).strip().lower()
        
        # Calculate exact match score
        match_score = calculate_enhanced_match_score(task_val_clean, child_name_clean)
        
        logger.info(f"Row {row}: '{task_val_clean}' vs '{child_name_clean}' -> Score = {match_score:.2f}")
        
        if match_score > best_match_score:
            best_match_score = match_score
            best_match_row = row
            logger.info(f"New best match at row {row} with score {match_score:.2f}")
    
    if best_match_row and best_match_score >= match_threshold:
        task_name = tracker_ws.cell(row=best_match_row, column=TASK_NAME_COL).value
        logger.info(f"✓ Exact match found at row {best_match_row}: '{task_name}' (score: {best_match_score:.2f})")
        return best_match_row
    
    logger.info(f"✗ No exact match found (best score: {best_match_score:.2f}, threshold: {match_threshold})")
    return None

def parse_percentage_value(pct_val):
    """Improved percentage parsing with better error handling"""
    if isinstance(pct_val, (int, float)):
        if 0 <= pct_val <= 1:
            return float(pct_val * 100)
        elif 0 <= pct_val <= 100:
            return float(pct_val)
        else:
            return float(pct_val)
    else:
        pct_str = str(pct_val).replace("%", "").replace(" ", "").strip()
        if pct_str:
            result = float(pct_str)
            if 0 <= result <= 1:
                result = result * 100
            return result
        else:
            return 0.0

def enhanced_text_matching(text1, text2):
    """Enhanced text matching for activity names"""
    common_words = ['work', 'activity', 'and', 'the', 'of', 'for', 'in', 'on', 'with', '&']
    overlap_threshold = 0.6
    
    def normalize_text(text):
        words = text.lower().replace('&', 'and').split()
        return [w for w in words if w not in common_words and len(w) > 2]
    
    words1 = normalize_text(text1)
    words2 = normalize_text(text2)
    
    if not words1 or not words2:
        return False
    
    overlap = len(set(words1) & set(words2))
    return overlap >= min(len(words1), len(words2)) * overlap_threshold

def calculate_enhanced_match_score(task_text, child_name_clean):
    """Enhanced match scoring with exact matching priority"""
    
    # Method 1: EXACT text matching (highest priority)
    if child_name_clean == task_text:
        logger.debug(f"EXACT text match found: '{child_name_clean}' == '{task_text}'")
        return 1.0
    
    # Method 2: Normalize and compare (handle spacing/formatting differences)
    def normalize_text(text):
        return re.sub(r'\s+', ' ', text.strip().lower())
    
    child_normalized = normalize_text(child_name_clean)
    task_normalized = normalize_text(task_text)
    
    if child_normalized == task_normalized:
        logger.debug(f"Normalized exact match: '{child_normalized}' == '{task_normalized}'")
        return 1.0
    
    # Method 3: Handle specific activity patterns with high precision
    if "checking" in child_name_clean and "casting" in child_name_clean:
        if ("checking & casting work" in task_text or 
            "checking and casting work" in task_text or
            "checking & casting" in task_text or  # Added this line to handle "Checking & Casting" without "Work"
            "checking and casting" in task_text or
            (all(word in task_text for word in ["checking", "casting"]))):
            logger.debug(f"Checking & casting activity match found")
            return 1.0
        else:
            logger.debug(f"Checking & casting activity mismatch - rejecting")
            return 0.0
    
    return 0.0  # Lower score for other cases

def find_child_activity_pct_with_hierarchy(tracker_ws, parent_names, child_name, tower=None):
    """
    Enhanced percentage extraction with precise parent-child matching and basement-level filtering
    """
    max_row = tracker_ws.max_row
    
    if isinstance(parent_names, str):
        parent_names = [parent_names]
    
    # Clean parent names
    parent_names = [str(p).strip().lower() for p in parent_names if p is not None and str(p).strip()]
    
    if not parent_names:
        logger.warning(f"No valid parent names provided for child: {child_name}")
        return 0.0
    
    child_name_clean = str(child_name).strip().lower() if child_name else ""
    if not child_name_clean:
        logger.warning("Child name is empty or None")
        return 0.0
    
    logger.info(f"=== HIERARCHY SEARCH for '{child_name}' ===")
    logger.info(f"Looking for parents: {parent_names}")
    
    # Check if this is an NTA search that needs basement-level filtering
    # Use tower parameter instead of parent names to detect NTA
    is_nta_search = tower and tower.startswith('NTA') if tower else False
    required_basement_type = None
    nta_number = None
    
    if is_nta_search:
        # Determine required basement type from parent names
        for parent in parent_names:
            parent_lower = parent.lower()
            if 'upper basement' in parent_lower:
                required_basement_type = 'upper basement'
                nta_number = "01"  # NTA-01 uses Upper Basement
                break
            elif 'lower basement' in parent_lower:
                required_basement_type = 'lower basement'
                nta_number = "02"  # NTA-02 uses Lower Basement
                break
        
        logger.info(f"NTA-{nta_number} search detected for tower '{tower}'. Required basement type: {required_basement_type}")
    
    # STEP 1: Find exact parent section matches with enhanced NTA validation
    matching_parent_sections = []
    
    for row in range(2, max_row + 1):
        cell_val = tracker_ws.cell(row=row, column=TASK_NAME_COL).value
        if cell_val:
            cell_val_clean = str(cell_val).strip().lower()
            
            # Check if this row is bold (parent activity)
            try:
                font = tracker_ws.cell(row=row, column=TASK_NAME_COL).font
                is_bold = font and font.bold
            except:
                is_bold = False
            
            if is_bold:
                # Find the section end
                section_start = row
                section_end = find_next_bold_parent(tracker_ws, row + 1, max_row)
                
                logger.debug(f"Processing bold row {row}: '{cell_val}' -> section {section_start} to {section_end}")
                
                # Check if this current bold row matches one of our required parents
                current_parent_match = False
                matched_parent = None
                for required_parent in parent_names:
                    if (required_parent in cell_val_clean or 
                        cell_val_clean in required_parent or
                        enhanced_text_matching(required_parent, cell_val_clean)):
                        current_parent_match = True
                        matched_parent = required_parent
                        logger.debug(f"Parent match found: '{cell_val_clean}' matches '{required_parent}'")
                        break
                
                if current_parent_match:
                    # For NTA areas, apply enhanced section validation BEFORE doing anything else
                    if is_nta_search and required_basement_type and nta_number:
                        # First check: Row range validation
                        if not validate_nta_section_by_row_range(section_start, section_end, required_basement_type, nta_number):
                            logger.info(f"❌ Skipping NTA-{nta_number} section at row {row}: failed row range validation")
                            continue
                        
                        # Second check: NTA section identity validation
                        if not verify_nta_section_identity(tracker_ws, section_start, nta_number):
                            logger.info(f"❌ Skipping section at row {row}: failed NTA-{nta_number} identity validation")
                            continue
                        
                        logger.info(f"✅ NTA-{nta_number} section at row {row}: passed all validations")
                    
                    # Now check if we can find the other required parents in nearby bold rows
                    all_parents_found = verify_all_parents_in_section(tracker_ws, row, parent_names, max_row)
                    
                    if all_parents_found:
                        section_desc = f"[{cell_val_clean}] rows {section_start}-{section_end}"
                        if is_nta_search and required_basement_type:
                            section_desc += f" [NTA-{nta_number} {required_basement_type.upper()}]"
                        logger.info(f"✅ Found valid parent section at row {row}: {section_desc}")
                        matching_parent_sections.append((section_start, section_end))
                        
                        # For NTA-02, take the FIRST valid section that passes our strict validation
                        # and stop looking for more to avoid confusion
                        if is_nta_search and nta_number == "02" and required_basement_type == "lower basement":
                            logger.info(f"✅ NTA-02: Found valid section at row {row} but continuing to search for Column/Shear Wall section")
                            # Don't break here for NTA-02 - we need to find the Column/Shear Wall section which likely contains the activity
                            pass
    
    if not matching_parent_sections:
        logger.warning(f"❌ No valid parent sections found for: {parent_names}")
        return 0.0
    
    # STEP 2: Search for the exact child activity in matching sections
    for section_start, section_end in matching_parent_sections:
        logger.info(f"\n--- Searching for EXACT child '{child_name}' in section {section_start} to {section_end} ---")
        
        found_row = find_exact_child_in_section(
            tracker_ws, section_start + 1, section_end, child_name_clean
        )
        
        if found_row:
            # Get percentage from the found row
            task_name = tracker_ws.cell(row=found_row, column=TASK_NAME_COL).value
            correct_col, pct_val = find_correct_percentage_column(tracker_ws, found_row, task_name)
            
            logger.info(f"✅ FOUND exact match at row {found_row}: '{task_name}' = {pct_val} (column {correct_col})")
            
            if pct_val is not None:
                try:
                    # Handle different percentage formats
                    result = parse_percentage_value(pct_val)
                    
                    logger.info(f"✅ Successfully extracted percentage: {pct_val} -> {result}%")
                    
                    # For NTA searches, return the FIRST valid match from the correct section
                    if is_nta_search and required_basement_type:
                        logger.info(f"✅ NTA-{nta_number} final result from row {found_row}: {result}%")
                        return result
                    
                    # For non-NTA, return the first match found
                    return result
                    
                except Exception as e:
                    logger.warning(f"❌ Error parsing percentage '{pct_val}': {e}")
                    continue
    
    logger.warning(f"❌ Child activity '{child_name}' not found in any matching parent section")
    return 0.0

def find_next_bold_parent(tracker_ws, start_row, max_row):
    """Find the next bold parent to determine section boundary - IMPROVED"""
    for row in range(start_row, max_row + 1):
        cell_val = tracker_ws.cell(row=row, column=TASK_NAME_COL).value
        if cell_val and str(cell_val).strip():
            try:
                font = tracker_ws.cell(row=row, column=TASK_NAME_COL).font
                is_bold = font and font.bold
                if is_bold:
                    logger.debug(f"Found next bold parent at row {row}: '{cell_val}'")
                    return row - 1  # Return row before the next bold parent
            except:
                pass
    
    # If no next bold parent found, extend the section significantly for NTA areas
    logger.debug(f"No next bold parent found after row {start_row}, using max_row {max_row}")
    return max_row  # If no next bold parent

def calculate_percentage_for_current_month(tower, month, month_col, kra_ws, tracker_wb, sheet_mapping):
    """Calculate percentage for the current tracker month using simplified hierarchy matching"""
    # Get parent activity names from multiple rows (using hardcoded KRA_PARENT_ROW)
    parent_names = []
    if tower in KRA_PARENT_ROW:
        parent_rows = KRA_PARENT_ROW[tower]
        for parent_row in parent_rows:
            parent_name = kra_ws.cell(row=parent_row, column=month_col).value
            if parent_name and str(parent_name).strip():
                parent_names.append(str(parent_name).strip())
    
    # Get child activity name (using hardcoded KRA_ACTIVITY_ROW)
    child_name = ""
    if tower in KRA_ACTIVITY_ROW:
        child_row = KRA_ACTIVITY_ROW[tower]
        child_name = kra_ws.cell(row=child_row, column=month_col).value
        if child_name:
            child_name = str(child_name).strip()
    
    if not parent_names or not child_name:
        logger.warning(f"Missing parent activities or child activity for {tower} in {month}")
        return 0.0
    
    logger.info(f"\n=== PROCESSING {tower} ({month}) ===")
    logger.info(f"Parent activities: {parent_names}")
    logger.info(f"Child activity: {child_name}")
    
    # Get corresponding tracker sheet
    tracker_sheetname = sheet_mapping.get(tower)
    if not tracker_sheetname or tracker_sheetname not in tracker_wb.sheetnames:
        logger.warning(f"Sheet for '{tower}' not found in tracker")
        return 0.0
    
    tracker_ws = tracker_wb[tracker_sheetname]
    logger.info(f"Using tracker sheet: {tracker_sheetname}")
    
    # Add debugging to understand sheet structure
    debug_tracker_sheet_structure(tracker_ws, tower)
    
    # Find the percentage completion using hierarchy
    pct = find_child_activity_pct_with_hierarchy(tracker_ws, parent_names, child_name, tower)
    
    # If hierarchy method didn't work well, try alternative method
    if pct == 0.0:
        logger.info(f"Hierarchy method returned 0%, trying alternative search...")
        pct = alternative_percentage_search(tracker_ws, child_name, tower)
    
    # Validate against expected values
    validate_expected_percentages(tower, pct)
    
    logger.info(f"✓ {tower} ({month}): '{child_name}' = {pct:.1f}% complete")
    logger.info(f"=== END {tower} ===\n")
    
    return pct

def format_progress_status(achieved_activities, planned_activities):
    """Format the progress status based on achieved vs planned activities"""
    if not achieved_activities and not planned_activities:
        return "No Progress"
    
    status_lines = []
    if achieved_activities:
        status_lines.append(f"Achieved-{achieved_activities}")
    else:
        status_lines.append("No Progress")
    
    if planned_activities:
        status_lines.append(f"Planned-{planned_activities}")
    
    return "\n".join(status_lines)

def main():
    logger.info("Starting Eden KRA Milestone Report generation...")
    
    try:
        # Initialize COS and download files
        cos = init_cos()
        logger.info("Downloading KRA file...")
        kra_xlsx = download_file_bytes(cos, KRA_KEY)
        logger.info("Downloading Tracker file...")
        tracker_xlsx = download_file_bytes(cos, TRACKER_KEY)
        
        # Load workbooks
        kra_wb = load_workbook(filename=BytesIO(kra_xlsx), data_only=True)
        tracker_wb = load_workbook(filename=BytesIO(tracker_xlsx), data_only=True)
        kra_ws = kra_wb.active
        
        # ============= DYNAMIC DISCOVERY =============
        logger.info("Discovering months and columns from KRA sheet...")
        month_columns = discover_months_and_columns(kra_ws)
        
        logger.info("Determining current month from tracker filename...")
        current_month = discover_current_month(TRACKER_KEY)
        
        logger.info("Discovering available towers...")
        available_towers = discover_towers(kra_ws)
        
        logger.info("Discovering tracker sheet mapping...")
        sheet_mapping = discover_tracker_sheets(tracker_wb)
        
        # Filter towers to only those we have row mappings for
        valid_towers = [tower for tower in available_towers if tower in KRA_ACTIVITY_ROW]
        
        if not valid_towers:
            logger.error("No valid towers found with row mappings!")
            return
        
        if current_month not in month_columns:
            logger.warning(f"Current month '{current_month}' not found in KRA sheet. Available months: {list(month_columns.keys())}")
            # Use the first available month as fallback
            current_month = list(month_columns.keys())[0] if month_columns else "June"
        
        logger.info(f"Processing {len(valid_towers)} towers for current month: {current_month}")
        
        # ============= PROCESS DATA =============
        results = []
        
        for tower in valid_towers:
            logger.info(f"Processing {tower}...")
            
            # Get June activities (all parent + child activities) - EXACT from KRA sheet
            june_month = "June"  # Fixed to June as requested
            if june_month not in month_columns:
                # Try to find June in available months (case insensitive)
                june_month = next((month for month in month_columns.keys() if 'june' in month.lower()), None)
                if not june_month:
                    logger.warning(f"June not found in available months: {list(month_columns.keys())}")
                    continue
            
            june_month_col = month_columns[june_month]
            
            # Get all June activities (parent + child) with exact text
            june_activities = get_all_activities_for_month(tower, june_month, june_month_col, kra_ws)
            
            # Get tower name from KRA sheet instead of milestone name
            tower_name = get_tower_name_from_kra(tower, kra_ws)
            
            # Calculate percentage for current month
            current_month_col = month_columns[current_month]
            current_month_pct = calculate_percentage_for_current_month(
                tower, current_month, current_month_col, kra_ws, tracker_wb, sheet_mapping
            )
            
            # Get dynamic weightage
            weightage = calculate_dynamic_weightage(tower, kra_ws, month_columns)
            
            # Calculate weighted work done
            weighted_work_done = round((current_month_pct * weightage) / 100, 1)
            
            # Get achieved and planned activities
            current_activity = get_activity_for_month(tower, current_month, current_month_col, kra_ws)
            achieved_activity = current_activity if current_month_pct > 0 else ""
            planned_activity = current_activity if current_month_pct == 0 else ""
            
            # Format progress status with separator line
            progress_status = format_progress_status(achieved_activity, planned_activity)
            
            # Create row data with Responsible Person and Delay Reasons at the end
            row_data = {
                "Milestone": tower_name,
                f"Activity- Target to be complete by June {datetime.now().year}": june_activities,
                f"% work done against Target- {current_month} Status": f"{current_month_pct:.0f}%" if current_month_pct > 0 else "0%",
                "Weightage": weightage,
                "Weighted Work done against Target": f"{weighted_work_done:.1f}%",
                f"Progress-{current_month}": progress_status,
                # Add July columns (blank for now)
                f"Activity- Target to be complete by July {datetime.now().year}": "",
                f"% work done against Target- July Status": "",
                "Weightage_July": "",
                "Weighted Work done against Target_July": "",
                "Progress-July": "",
                # Add August columns (blank for now)
                f"Activity- Target to be complete by August {datetime.now().year}": "",
                f"% work done against Target- August Status": "",
                "Weightage_August": "",
                "Weighted Work done against Target_August": "",
                "Progress-August": "",
                # Move Responsible Person and Delay Reasons to the end
                "Responsible Person": "",  # Keep empty as requested
                "Delay Reasons": ""        # Keep empty as requested
            }
            
            results.append(row_data)
        
        if not results:
            logger.error("No data found to generate report!")
            return
        
        # ============= GENERATE EXCEL REPORT =============
        df = pd.DataFrame(results)
        filename = f"Eden_Progress_Against_Milestones ({datetime.now():%Y-%m-%d}).xlsx"
        
        # Create formatted Excel file
        wb = Workbook()
        ws = wb.active
        ws.title = "Eden- Progress Against Milestones"
        
        # Add title row
        ws.append(["Eden- Progress Against Milestones"])
        
        # Add report generation date below the heading
        ws.append([f"Report Generated on: {datetime.now().strftime('%B %d, %Y')}"])
        ws.append([])  # Empty row for spacing
        
        # Add data
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # ============= FORMAT EXCEL =============
        header_font = Font(bold=True, size=10, color="000000")
        title_font = Font(bold=True, size=14, color="000000")
        date_font = Font(size=10, color="666666")
        data_font = Font(size=9)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        header_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
        
        # Format title row (row 1)
        ws.merge_cells(f'A1:{get_column_letter(len(df.columns))}1')
        ws['A1'].font = title_font
        ws['A1'].alignment = center_align
        
        # Format date row (row 2)
        ws.merge_cells(f'A2:{get_column_letter(len(df.columns))}2')
        ws['A2'].font = date_font
        ws['A2'].alignment = center_align
        
        # Format headers (row 4)
        for cell in ws[4]:
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
            cell.fill = header_fill
        
        # Format data rows
        for row_idx, row in enumerate(ws.iter_rows(min_row=5, max_row=ws.max_row), 5):
            for col_idx, cell in enumerate(row, 1):
                cell.border = border
                cell.font = data_font
                
                # Alignment based on column type
                # Updated column indices since Responsible Person and Delay Reasons moved to the end
                if col_idx in [1, 2, 6, 7, 12, 13, 18, 19, 20]:  # Text columns (Milestone, Activity columns, Progress columns, Responsible Person, Delay Reasons)
                    cell.alignment = left_align
                else:  # Percentage, Weightage columns
                    cell.alignment = center_align
        
        # Dynamic column widths based on content
        for col_idx in range(1, len(df.columns) + 1):
            col_letter = get_column_letter(col_idx)
            
            # Calculate optimal width based on column content
            max_length = 0
            for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
            
            # Set minimum and maximum width constraints
            calculated_width = min(max(max_length + 2, 10), 30)
            ws.column_dimensions[col_letter].width = calculated_width
        
        # Set row heights
        ws.row_dimensions[1].height = 25  # Title row
        ws.row_dimensions[2].height = 20  # Date row
        ws.row_dimensions[4].height = 40  # Header row
        
        # Set data row heights to accommodate wrapped text
        for row_idx in range(5, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 35
        
        # Save the file
        wb.save(filename)
        logger.info(f"Successfully saved Eden Progress Against Milestones report to {filename}")
        
        # Log summary
        logger.info("Report Summary:")
        logger.info(f"  Current Month: {current_month}")
        logger.info(f"  Available Months: {list(month_columns.keys())}")
        logger.info(f"  Processed Towers: {len(valid_towers)}")
        
        for result in results:
            milestone = result['Milestone']
            progress_key = f'% work done against Target- {current_month} Status'
            weighted_key = 'Weighted Work done against Target'
            progress = result[progress_key]
            weighted = result[weighted_key]
            logger.info(f"  {milestone}: Progress: {progress}, Weighted: {weighted}")
            
    except Exception as e:
        logger.error(f"Error generating report: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    main()
