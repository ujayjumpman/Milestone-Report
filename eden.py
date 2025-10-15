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
from typing import Optional, Tuple, List, Dict

# ======================= CONFIG =======================
load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

COS_API_KEY = os.getenv("COS_API_KEY")
COS_CRN = os.getenv("COS_SERVICE_INSTANCE_CRN")
COS_ENDPOINT = os.getenv("COS_ENDPOINT")
BUCKET = os.getenv("COS_BUCKET_NAME")
KRA_FOLDER = os.getenv("KRA_FOLDER", "")  # Folder where KRA files are stored
EDEN_TRACKER_FOLDER = os.getenv("EDEN_TRACKER_FOLDER", "Eden/")

TASK_NAME_COL = 4
PCT_COL = 7
PCT_COL_ALT = [6, 8, 9, 10, 5]
RESPONSIBLE_COL = 6
DELAY_COL = 8

# Quarterly month groups
QUARTERS = {
    "Q1": ["June", "July", "August"],
    "Q2": ["September", "October", "November"],
    "Q3": ["December", "January", "February"],
    "Q4": ["March", "April", "May"]
}

# Month to tracker month mapping (result month -> tracker month)
# e.g., June results come from July tracker (DD-07-YYYY)
MONTH_TO_TRACKER_MAPPING = {
    "June": 7,      # July tracker
    "July": 8,      # August tracker
    "August": 9,    # September tracker
    "September": 10,   # October tracker
    "October": 11,     # November tracker
    "November": 12,    # December tracker
    "December": 1,     # January tracker (next year)
    "January": 2,      # February tracker
    "February": 3,     # March tracker
    "March": 4,        # April tracker
    "April": 5,        # May tracker
    "May": 6           # June tracker
}

required_vars = {
    'COS_API_KEY': COS_API_KEY,
    'COS_SERVICE_INSTANCE_CRN': COS_CRN,
    'COS_ENDPOINT': COS_ENDPOINT,
    'COS_BUCKET_NAME': BUCKET
}

missing_vars = [var_name for var_name, var_value in required_vars.items() if not var_value]
if missing_vars:
    error_msg = f"Missing required environment variables: {', '.join(missing_vars)}"
    logger.error(error_msg)
    raise ValueError(error_msg)

# ================= COS HELPERS =================
def init_cos():
    return ibm_boto3.client("s3", ibm_api_key_id=COS_API_KEY, ibm_service_instance_id=COS_CRN,
                            config=Config(signature_version="oauth"), endpoint_url=COS_ENDPOINT)

def download_file_bytes(cos, key):
    return cos.get_object(Bucket=BUCKET, Key=key)["Body"].read()

# ================= DYNAMIC KRA DISCOVERY =================
def find_latest_kra_file(cos_client, bucket_name: str, folder_prefix: str = "") -> Optional[Tuple[str, List[str], int]]:
    """
    Find the latest KRA Milestones file and extract the quarter months from filename.
    Returns: (file_key, list_of_months, year)
    """
    logger.info(f"\n{'='*70}")
    logger.info(f"SEARCHING FOR LATEST KRA FILE")
    logger.info(f"{'='*70}")
    logger.info(f"Folder: {folder_prefix if folder_prefix else 'Root'}")
    
    try:
        response = cos_client.list_objects_v2(Bucket=bucket_name, Prefix=folder_prefix)
        
        if 'Contents' not in response:
            logger.error(f"No files found in folder '{folder_prefix}'")
            return None
        
        kra_files = []
        
        for obj in response['Contents']:
            file_key = obj['Key']
            filename = os.path.basename(file_key)
            filename_lower = filename.lower()
            
            if file_key.endswith('/'):
                continue
            
            # Look for KRA Milestones pattern
            is_kra = 'kra' in filename_lower and 'milestone' in filename_lower
            is_excel = filename_lower.endswith(('.xlsx', '.xls'))
            
            if is_kra and is_excel:
                # Extract months from filename
                # Pattern: "KRA Milestones for June July August 2025.xlsx"
                months_pattern = r'(January|February|March|April|May|June|July|August|September|October|November|December)'
                found_months = re.findall(months_pattern, filename, re.IGNORECASE)
                
                # Capitalize month names
                found_months = [m.capitalize() for m in found_months]
                
                # Extract year
                year_match = re.search(r'(\d{4})', filename)
                year = int(year_match.group(1)) if year_match else datetime.now().year
                
                kra_files.append({
                    'key': file_key,
                    'filename': filename,
                    'months': found_months,
                    'year': year,
                    'last_modified': obj['LastModified']
                })
                
                logger.info(f"Found: {filename}")
                logger.info(f"  Months: {', '.join(found_months)}")
                logger.info(f"  Year: {year}")
        
        if not kra_files:
            logger.error("No KRA Milestone files found")
            return None
        
        # Sort by last modified date to get the latest
        kra_files.sort(key=lambda f: f['last_modified'], reverse=True)
        
        latest = kra_files[0]
        
        logger.info(f"\n{'='*70}")
        logger.info(f"SELECTED LATEST KRA FILE")
        logger.info(f"{'='*70}")
        logger.info(f"File: {latest['filename']}")
        logger.info(f"Path: {latest['key']}")
        logger.info(f"Months: {', '.join(latest['months'])}")
        logger.info(f"Year: {latest['year']}")
        logger.info(f"Last Modified: {latest['last_modified']}")
        logger.info(f"{'='*70}\n")
        
        return latest['key'], latest['months'], latest['year']
        
    except Exception as e:
        logger.error(f"Error searching for KRA file: {str(e)}")
        raise

# ================= DYNAMIC TRACKER DISCOVERY BY MONTH =================
def find_tracker_for_month(cos_client, bucket_name: str, target_month: int, target_year: int, folder_prefix: str = "Eden/") -> Optional[str]:
    """
    Find tracker file for a specific month and year.
    target_month: 1-12 (e.g., 7 for July)
    target_year: e.g., 2025
    """
    logger.info(f"\nSearching for tracker: Month={target_month}, Year={target_year}")
    
    try:
        response = cos_client.list_objects_v2(Bucket=bucket_name, Prefix=folder_prefix)
        
        if 'Contents' not in response:
            logger.warning(f"No files found in folder '{folder_prefix}'")
            return None
        
        matching_trackers = []
        
        for obj in response['Contents']:
            file_key = obj['Key']
            filename = os.path.basename(file_key)
            filename_lower = filename.lower()
            
            if file_key.endswith('/'):
                continue
            
            # Check if it's a tracker file
            is_tracker = any(pattern in filename_lower for pattern in 
                           ['structure work tracker', 'tracker', 'structure tracker'])
            is_excel = filename_lower.endswith(('.xlsx', '.xls'))
            
            if is_tracker and is_excel:
                # Extract date from filename: (DD-MM-YYYY)
                date_pattern = r'\((\d{1,2})-(\d{1,2})-(\d{4})\)'
                date_match = re.search(date_pattern, filename)
                
                if date_match:
                    day, month, year = date_match.groups()
                    file_month = int(month)
                    file_year = int(year)
                    
                    # Check if this matches our target month and year
                    if file_month == target_month and file_year == target_year:
                        matching_trackers.append({
                            'key': file_key,
                            'filename': filename,
                            'day': int(day),
                            'month': file_month,
                            'year': file_year,
                            'last_modified': obj['LastModified']
                        })
                        logger.info(f"  Match found: {filename}")
        
        if not matching_trackers:
            logger.warning(f"No tracker found for {target_month}/{target_year}")
            return None
        
        # If multiple trackers for same month, take the latest modified
        matching_trackers.sort(key=lambda f: f['last_modified'], reverse=True)
        selected = matching_trackers[0]
        
        logger.info(f"✓ Selected: {selected['filename']}")
        return selected['key']
        
    except Exception as e:
        logger.warning(f"Error searching for tracker: {str(e)}")
        return None

# ================= KRA STRUCTURE DISCOVERY =================
def discover_months_in_kra(kra_ws):
    """Discover month columns in KRA sheet"""
    months = {}
    month_patterns = {
        'January': ['january', 'jan'], 'February': ['february', 'feb'],
        'March': ['march', 'mar'], 'April': ['april', 'apr'],
        'May': ['may'], 'June': ['june', 'jun'],
        'July': ['july', 'jul'], 'August': ['august', 'aug'],
        'September': ['september', 'sep', 'sept'], 'October': ['october', 'oct'],
        'November': ['november', 'nov'], 'December': ['december', 'dec']
    }
    
    for col in range(2, 15):
        cell_value = kra_ws.cell(row=4, column=col).value
        if cell_value:
            cell_str = str(cell_value).strip().lower()
            for full_month, patterns in month_patterns.items():
                if any(pattern in cell_str for pattern in patterns):
                    months[full_month] = col
                    break
    
    logger.info(f"Discovered months in KRA: {list(months.keys())}")
    return months

def discover_towers_in_kra(kra_ws):
    """Discover tower structure in KRA sheet"""
    towers = {}
    for row in range(5, 50):  # Increased range to find more towers
        cell_value = kra_ws.cell(row=row, column=1).value
        if cell_value:
            cell_str = str(cell_value).strip()
            tower_match = re.match(r'^Tower\s+(\d+)$', cell_str, re.IGNORECASE)
            nta_match = re.match(r'^NTA-(\d+)$', cell_str, re.IGNORECASE)
            
            if tower_match:
                key = f"Tower {tower_match.group(1)}"
            elif nta_match:
                key = f"NTA-{nta_match.group(1)}"
            else:
                continue
            
            # Find the next 3 rows for this tower
            towers[key] = {
                'parent_row_1': row,
                'parent_row_2': row + 1,
                'child_row': row + 2
            }
    
    logger.info(f"Discovered towers: {list(towers.keys())}")
    return towers

def discover_tracker_sheets(tracker_wb):
    """Map tower names to tracker sheets"""
    sheet_mapping = {}
    nta_sheet_name = None
    
    for sheet_name in tracker_wb.sheetnames:
        sheet_clean = sheet_name.strip()
        tower_match = re.search(r'Tower\s*(\d+)', sheet_clean, re.IGNORECASE)
        if tower_match:
            sheet_mapping[f"Tower {tower_match.group(1)}"] = sheet_clean
        elif re.search(r'non.*tower.*area', sheet_clean, re.IGNORECASE):
            nta_sheet_name = sheet_clean
    
    # Map all NTA sections to the same sheet
    if nta_sheet_name:
        sheet_mapping["NTA"] = nta_sheet_name
        for i in range(1, 20):
            sheet_mapping[f"NTA-{i}"] = nta_sheet_name
            sheet_mapping[f"NTA-{i:02d}"] = nta_sheet_name
    
    return sheet_mapping

# ================= ACTIVITY MATCHING FUNCTIONS =================
def get_activities_from_kra(tower, month_col, kra_ws, tower_structure):
    """Extract activities from KRA for a specific tower and month"""
    if tower not in tower_structure:
        return []
    
    rows = tower_structure[tower]
    activities = []
    
    for key in ['parent_row_1', 'parent_row_2', 'child_row']:
        val = kra_ws.cell(row=rows[key], column=month_col).value
        if val and str(val).strip():
            activities.append(str(val).strip())
    
    return activities

def calculate_match_score(text1, text2):
    """Calculate similarity score between two texts"""
    t1, t2 = text1.lower().strip(), text2.lower().strip()
    if t1 == t2:
        return 1.0
    if ' '.join(t1.split()) == ' '.join(t2.split()):
        return 1.0
    if t1 in t2 or t2 in t1:
        return 0.9
    
    w1, w2 = set(t1.split()), set(t2.split())
    if not w1 or not w2:
        return 0.0
    
    return len(w1 & w2) / len(w1 | w2)

def find_correct_percentage_column(tracker_ws, row, task_name):
    """Find the correct percentage column for a task"""
    check_columns = [PCT_COL] + PCT_COL_ALT
    
    for col in check_columns:
        try:
            val = tracker_ws.cell(row=row, column=col).value
            if val is not None:
                s = str(val).strip()
                if s.endswith('%'):
                    s = s.replace('%', '').strip()
                    if s.replace('.', '').isdigit():
                        return col, val
                elif s.replace('.', '').isdigit():
                    return col, val
        except:
            continue
    
    return None, None

def parse_percentage_value(pct_val):
    """Parse percentage value to float"""
    if pct_val is None:
        return 0.0
        
    if isinstance(pct_val, (int, float)):
        if 0 <= pct_val <= 1:
            return pct_val * 100
        return float(pct_val)
    
    s = str(pct_val).replace('%', '').strip()
    if s:
        try:
            f = float(s)
            if 0 <= f <= 1:
                f *= 100
            return f
        except ValueError:
            return 0.0
    return 0.0

def get_nta_section_bounds(tracker_ws, nta_key):
    """Get search bounds for NTA sections"""
    max_row = tracker_ws.max_row
    logger.info(f"[{nta_key}] Will search entire NTA sheet for matching activities")
    return 1, max_row

def find_activity_in_tracker(tracker_ws, parent_activities, child_activity, tower=None):
    """
    Find activity in tracker and extract percentage, responsible person, and delay.
    """
    max_row = tracker_ws.max_row
    
    if not parent_activities or not child_activity:
        logger.warning(f"[{tower}] Missing activities")
        return 0.0, "", ""
    
    parent_list = [str(p).strip().lower() for p in parent_activities if p]
    child_clean = str(child_activity).strip().lower()
    
    logger.info(f"\n[{tower}] Searching for activities...")
    
    is_nta = tower and tower.startswith('NTA')
    
    # Determine search range
    if is_nta:
        start_row, end_row = get_nta_section_bounds(tracker_ws, tower)
        if not start_row or not end_row:
            return 0.0, "", ""
        search_start, search_end = start_row, end_row
    else:
        search_start, search_end = 2, max_row
    
    # Find all bold rows (section headers)
    bold_rows = []
    for row in range(search_start, search_end + 1):
        val = tracker_ws.cell(row=row, column=TASK_NAME_COL).value
        if val:
            try:
                bold = tracker_ws.cell(row=row, column=TASK_NAME_COL).font.bold
            except:
                bold = False
            
            if bold:
                bold_rows.append({
                    'row': row,
                    'text': str(val).strip(),
                    'text_lower': str(val).strip().lower()
                })
    
    # Find matching parent sections
    matching_groups = []
    
    for i, bold_row in enumerate(bold_rows):
        first_parent = parent_list[0] if parent_list else None
        if not first_parent:
            continue
        
        first_match_score = calculate_match_score(bold_row['text_lower'], first_parent)
        
        if first_match_score >= 0.7:
            found_parents = {parent_list[0]: bold_row['row']}
            
            # Look for remaining parent activities
            for j in range(1, len(parent_list)):
                parent_to_find = parent_list[j]
                found = False
                
                for k in range(i + 1, min(i + 10, len(bold_rows))):
                    check_bold = bold_rows[k]
                    score = calculate_match_score(check_bold['text_lower'], parent_to_find)
                    
                    if score >= 0.7:
                        found_parents[parent_to_find] = check_bold['row']
                        found = True
                        break
                
                if not found:
                    break
            
            # If all parent activities found, add to matching groups
            if len(found_parents) == len(parent_list):
                group_start = min(found_parents.values())
                
                # Find end of group
                group_end = search_end
                for next_bold in bold_rows:
                    if next_bold['row'] > max(found_parents.values()):
                        if all(calculate_match_score(next_bold['text_lower'], p) < 0.5 for p in parent_list):
                            group_end = next_bold['row'] - 1
                            break
                
                matching_groups.append({
                    'start': group_start,
                    'end': group_end,
                    'parent_rows': sorted(found_parents.values())
                })
                
                if is_nta:
                    break  # For NTA, take first exact match
    
    # Search for child activity in matching groups
    for g in matching_groups:
        best_row = None
        best_score = 0
        
        for row in range(g['start'], g['end'] + 1):
            # Skip bold rows
            try:
                if tracker_ws.cell(row=row, column=TASK_NAME_COL).font.bold:
                    continue
            except:
                pass
            
            val = tracker_ws.cell(row=row, column=TASK_NAME_COL).value
            if not val:
                continue
            
            score = calculate_match_score(str(val).strip().lower(), child_clean)
            
            if score > best_score:
                best_score = score
                best_row = row
        
        if best_row and best_score >= 0.75:
            val = tracker_ws.cell(row=best_row, column=TASK_NAME_COL).value
            col, pct = find_correct_percentage_column(tracker_ws, best_row, val)
            
            if pct is not None:
                resp = tracker_ws.cell(row=best_row, column=RESPONSIBLE_COL).value or ""
                delay = tracker_ws.cell(row=best_row, column=DELAY_COL).value or ""
                parsed_pct = parse_percentage_value(pct)
                
                logger.info(f"  ✓ Match found: {parsed_pct:.1f}%")
                return parsed_pct, str(resp).strip(), str(delay).strip()
    
    logger.warning(f"  ✗ No match found")
    return 0.0, "", ""

# ================= MONTH DATA CALCULATION =================
def calculate_month_data(tower, month_name, month_col, kra_ws, tracker_cache, sheet_mapping, tower_structure, kra_year):
    """
    Calculate data for a specific month using the appropriate tracker.
    """
    activities = get_activities_from_kra(tower, month_col, kra_ws, tower_structure)
    
    if not activities:
        return {
            'activities_text': "",
            'percentage': 0.0,
            'progress_status': "No Progress",
            'responsible': "",
            'delay': ""
        }
    
    activities_text = '\n'.join(activities)
    
    # Get the tracker month for this result month
    tracker_month_num = MONTH_TO_TRACKER_MAPPING.get(month_name)
    
    if not tracker_month_num:
        logger.warning(f"No tracker mapping for month: {month_name}")
        return {
            'activities_text': activities_text,
            'percentage': 0.0,
            'progress_status': "No Progress",
            'responsible': "",
            'delay': ""
        }
    
    # Determine tracker year (handle year transition for December -> January)
    if month_name == "December" and tracker_month_num == 1:
        tracker_year = kra_year + 1
    elif month_name in ["January", "February"] and tracker_month_num in [2, 3]:
        tracker_year = kra_year
    else:
        tracker_year = kra_year
    
    # Check if tracker exists in cache
    tracker_key = f"{tracker_month_num}_{tracker_year}"
    
    if tracker_key not in tracker_cache:
        logger.info(f"[{month_name}] No tracker available for {tracker_month_num}/{tracker_year}")
        return {
            'activities_text': activities_text,
            'percentage': 0.0,
            'progress_status': "No Progress",
            'responsible': "",
            'delay': ""
        }
    
    tracker_wb = tracker_cache[tracker_key]
    tracker_sheet = sheet_mapping.get(tower)
    
    if not tracker_sheet or tracker_sheet not in tracker_wb.sheetnames:
        logger.warning(f"[{month_name}] Sheet for '{tower}' not found in tracker")
        return {
            'activities_text': activities_text,
            'percentage': 0.0,
            'progress_status': "No Progress",
            'responsible': "",
            'delay': ""
        }
    
    tracker_ws = tracker_wb[tracker_sheet]
    parent_activities = activities[:-1] if len(activities) > 1 else []
    child_activity = activities[-1]
    
    pct, responsible, delay = find_activity_in_tracker(tracker_ws, parent_activities, child_activity, tower)
    
    progress_status = f"Achieved-{child_activity}" if pct > 0 else "No Progress"
    
    return {
        'activities_text': activities_text,
        'percentage': pct,
        'progress_status': progress_status,
        'responsible': responsible,
        'delay': delay
    }

# ================= EXCEL FORMATTING =================
def format_excel_report(ws, df):
    """Format the Excel report with proper styling"""
    header_font = Font(bold=True, size=10)
    title_font = Font(bold=True, size=14)
    date_font = Font(size=10, color="666666")
    data_font = Font(size=9)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    
    # Format title and date rows
    ws.merge_cells(f'A1:{get_column_letter(len(df.columns))}1')
    ws['A1'].font = title_font
    ws['A1'].alignment = center_align
    
    ws.merge_cells(f'A2:{get_column_letter(len(df.columns))}2')
    ws['A2'].font = date_font
    ws['A2'].alignment = center_align
    
    # Format headers
    for cell in ws[4]:
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border
        cell.fill = header_fill
    
    # Format data rows
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
        for col_idx, cell in enumerate(row, 1):
            cell.border = border
            cell.font = data_font
            
            header_val = ws.cell(row=4, column=col_idx).value or ''
            if any(kw in str(header_val) for kw in ['Tower', 'Target', 'Activity', 'Progress', 'Responsible', 'Delay']):
                cell.alignment = left_align
            else:
                cell.alignment = center_align
    
    # Set column widths
    for col_idx in range(1, len(df.columns) + 1):
        max_length = 10
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_length + 2, 10), 35)
    
    # Set row heights
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[4].height = 40
    
    for row_idx in range(5, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 35

# ================= MAIN FUNCTION =================
def main():
    logger.info("Starting Quarterly Eden KRA Report Generator...")
    
    try:
        cos = init_cos()
        
        # Step 1: Find latest KRA file
        logger.info("\n" + "="*70)
        logger.info("STEP 1: Finding Latest KRA File")
        logger.info("="*70)
        
        kra_result = find_latest_kra_file(cos, BUCKET, KRA_FOLDER)
        if not kra_result:
            logger.error("Could not find KRA file. Exiting...")
            return
        
        kra_key, quarter_months, kra_year = kra_result
        
        logger.info(f"KRA File: {kra_key}")
        logger.info(f"Quarter Months: {', '.join(quarter_months)}")
        logger.info(f"Year: {kra_year}")
        
        # Step 2: Download and load KRA
        logger.info("\n" + "="*70)
        logger.info("STEP 2: Loading KRA File")
        logger.info("="*70)
        
        kra_bytes = download_file_bytes(cos, kra_key)
        kra_wb = load_workbook(filename=BytesIO(kra_bytes), data_only=True)
        kra_ws = kra_wb.active
        
        # Discover KRA structure
        months_in_kra = discover_months_in_kra(kra_ws)
        tower_structure = discover_towers_in_kra(kra_ws)
        
        # Step 3: Find and load all required trackers
        logger.info("\n" + "="*70)
        logger.info("STEP 3: Finding and Loading Trackers")
        logger.info("="*70)
        
        tracker_cache = {}  # {month_year: workbook}
        
        for month_name in quarter_months:
            tracker_month_num = MONTH_TO_TRACKER_MAPPING.get(month_name)
            
            if not tracker_month_num:
                continue
            
            # Handle year transition
            if month_name == "December" and tracker_month_num == 1:
                tracker_year = kra_year + 1
            elif month_name in ["January", "February"] and tracker_month_num in [2, 3]:
                tracker_year = kra_year
            else:
                tracker_year = kra_year
            
            logger.info(f"\nLooking for tracker for {month_name} results (tracker month: {tracker_month_num}/{tracker_year})...")
            
            tracker_key_found = find_tracker_for_month(cos, BUCKET, tracker_month_num, tracker_year, EDEN_TRACKER_FOLDER)
            
            if tracker_key_found:
                logger.info(f"  Downloading: {tracker_key_found}")
                tracker_bytes = download_file_bytes(cos, tracker_key_found)
                tracker_wb = load_workbook(filename=BytesIO(tracker_bytes), data_only=True)
                tracker_cache[f"{tracker_month_num}_{tracker_year}"] = tracker_wb
                logger.info(f"  ✓ Tracker loaded successfully")
            else:
                logger.warning(f"  ✗ Tracker not found for {month_name} (will leave columns blank)")
        
        logger.info(f"\nTotal trackers loaded: {len(tracker_cache)}")
        
        # Step 4: Get sheet mapping from first available tracker
        logger.info("\n" + "="*70)
        logger.info("STEP 4: Discovering Tracker Sheet Mapping")
        logger.info("="*70)
        
        sheet_mapping = {}
        if tracker_cache:
            first_tracker = list(tracker_cache.values())[0]
            sheet_mapping = discover_tracker_sheets(first_tracker)
            logger.info(f"Sheet mapping: {sheet_mapping}")
        else:
            logger.warning("No trackers available, will generate report with blank data")
        
        # Step 5: Process data for each tower and month
        logger.info("\n" + "="*70)
        logger.info("STEP 5: Processing Tower Data")
        logger.info("="*70)
        
        results = []
        target_month = quarter_months[-1]  # Last month in quarter is the target
        
        for tower in sorted(tower_structure.keys()):
            logger.info(f"\nProcessing: {tower}")
            
            row_data = {'Tower': tower}
            
            # Get target activities (from last month of quarter)
            target_col = months_in_kra.get(target_month)
            if target_col:
                target_activities = get_activities_from_kra(tower, target_col, kra_ws, tower_structure)
                row_data[f'Target till {target_month}'] = '\n'.join(target_activities) if target_activities else ""
            else:
                row_data[f'Target till {target_month}'] = ""
            
            # Process each month in the quarter
            total_weighted = 0
            weightage = 100 if not tower.startswith('NTA') else 50
            months_count = len(quarter_months)
            
            for month_name in quarter_months:
                logger.info(f"  Processing month: {month_name}")
                
                month_col = months_in_kra.get(month_name)
                
                if not month_col:
                    logger.warning(f"    Month '{month_name}' not found in KRA sheet")
                    # Add blank columns
                    row_data[f"Activity- Target to be complete by {month_name} {kra_year}"] = ""
                    row_data[f"% work done against Target- {month_name} Status"] = ""
                    row_data[f"Progress-{month_name}"] = ""
                    row_data[f"Responsible Person-{month_name}"] = ""
                    row_data[f"Delay Reasons-{month_name}"] = ""
                    continue
                
                # Calculate month data
                month_data = calculate_month_data(
                    tower, month_name, month_col, kra_ws, tracker_cache, 
                    sheet_mapping, tower_structure, kra_year
                )
                
                # Add to row data
                row_data[f"Activity- Target to be complete by {month_name} {kra_year}"] = month_data['activities_text']
                row_data[f"% work done against Target- {month_name} Status"] = f"{month_data['percentage']:.0f}%" if month_data['percentage'] > 0 else "0%"
                row_data[f"Progress-{month_name}"] = month_data['progress_status']
                row_data[f"Responsible Person-{month_name}"] = ""  # Leave blank as per requirement
                row_data[f"Delay Reasons-{month_name}"] = ""      # Leave blank as per requirement
                
                # Calculate weighted contribution
                total_weighted += round(month_data['percentage'] * weightage) / (100 * months_count)
                
                logger.info(f"    {month_name}: {month_data['percentage']:.1f}%")
            
            # Add final columns
            row_data['Weightage'] = weightage
            row_data['Weighted Work done against Target'] = f"{total_weighted:.1f}%"
            
            results.append(row_data)
            logger.info(f"  ✓ {tower} completed. Weighted total: {total_weighted:.1f}%")
        
        # Step 6: Generate Excel Report
        logger.info("\n" + "="*70)
        logger.info("STEP 6: Generating Excel Report")
        logger.info("="*70)
        
        if not results:
            logger.error("No data to generate report!")
            return
        
        df = pd.DataFrame(results)
        
        # Create filename
        quarter_name = f"{'_'.join(quarter_months)}"
        filename = f"Eden_Progress_Against_Milestones_{quarter_name}_{kra_year}.xlsx"
        
        # Create Excel file
        wb = Workbook()
        ws = wb.active
        ws.title = "Eden- Progress Against Milestones"
        
        # Add title and date
        ws.append(["Eden- Progress Against Milestones"])
        ws.append([f"Report Generated on: {datetime.now().strftime('%B %d, %Y')}"])
        ws.append([])
        
        # Add data
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Format the report
        format_excel_report(ws, df)
        
        # Save the file
        wb.save(filename)
        
        logger.info(f"\n{'='*70}")
        logger.info("REPORT GENERATION COMPLETE")
        logger.info(f"{'='*70}")
        logger.info(f"File saved: {filename}")
        logger.info(f"Quarter: {', '.join(quarter_months)} {kra_year}")
        logger.info(f"Total towers: {len(results)}")
        logger.info(f"Trackers used: {len(tracker_cache)}")
        
        # Summary
        logger.info(f"\nSummary by Tower:")
        for result in results:
            tower_name = result['Tower']
            weighted = result.get('Weighted Work done against Target', '0.0%')
            logger.info(f"  {tower_name}: {weighted}")
        
        logger.info(f"\n{'='*70}\n")
        
    except Exception as e:
        logger.error(f"Error generating report: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    main()
