# eden.py
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
from typing import Optional, Tuple, List

# ======================= CONFIG =======================
load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

COS_API_KEY = os.getenv("COS_API_KEY")
COS_CRN = os.getenv("COS_SERVICE_INSTANCE_CRN")
COS_ENDPOINT = os.getenv("COS_ENDPOINT")
BUCKET = os.getenv("COS_BUCKET_NAME")
KRA_KEY = os.getenv("KRA_FILE_PATH")
EDEN_TRACKER_FOLDER = os.getenv("EDEN_TRACKER_FOLDER", "Eden/")

EDEN_TRACKER_KEY = None

TASK_NAME_COL = 4
PCT_COL = 7
PCT_COL_ALT = [6, 8, 9, 10, 5]
RESPONSIBLE_COL = 6
DELAY_COL = 8

required_vars = {
    'COS_API_KEY': COS_API_KEY,
    'COS_SERVICE_INSTANCE_CRN': COS_CRN,
    'COS_ENDPOINT': COS_ENDPOINT,
    'COS_BUCKET_NAME': BUCKET,
    'KRA_FILE_PATH': KRA_KEY
}

missing_vars = [var_name for var_name, var_value in required_vars.items() if not var_value]
if missing_vars:
    error_msg = f"Missing required environment variables: {', '.join(missing_vars)}"
    logger.error(error_msg)
    raise ValueError(error_msg)

# ================= DYNAMIC TRACKER DISCOVERY =================
def find_latest_eden_tracker(cos_client, bucket_name: str, folder_prefix: str = "Eden/") -> Optional[str]:
    logger.info(f"\n{'='*70}")
    logger.info(f"SEARCHING FOR LATEST EDEN TRACKER")
    logger.info(f"{'='*70}")
    logger.info(f"Folder: {folder_prefix}")
    
    try:
        response = cos_client.list_objects_v2(Bucket=bucket_name, Prefix=folder_prefix)
        
        if 'Contents' not in response:
            logger.error(f"No files found in folder '{folder_prefix}'")
            return None
        
        tracker_files = []
        
        for obj in response['Contents']:
            file_key = obj['Key']
            filename = os.path.basename(file_key)
            filename_lower = filename.lower()
            
            if file_key.endswith('/'):
                continue
            
            is_tracker = any(pattern in filename_lower for pattern in 
                           ['structure work tracker', 'tracker', 'structure tracker'])
            is_excel = filename_lower.endswith(('.xlsx', '.xls'))
            
            if is_tracker and is_excel:
                date_pattern = r'\((\d{1,2})-(\d{1,2})-(\d{4})\)'
                date_match = re.search(date_pattern, filename)
                
                file_date = None
                if date_match:
                    try:
                        day, month, year = date_match.groups()
                        file_date = datetime.strptime(f"{day}-{month}-{year}", "%d-%m-%Y")
                    except ValueError as e:
                        logger.warning(f"Could not parse date from '{filename}': {e}")
                
                tracker_files.append({
                    'key': file_key,
                    'filename': filename,
                    'extracted_date': file_date,
                    'last_modified': obj['LastModified'],
                    'size': obj['Size']
                })
                
                logger.info(f"Found: {filename}")
                if file_date:
                    logger.info(f"  Date: {file_date.strftime('%d-%m-%Y')}")
        
        if not tracker_files:
            logger.error("No Eden tracker files found")
            return None
        
        def sort_key(f):
            return f['extracted_date'] or f['last_modified']
        
        tracker_files.sort(key=sort_key, reverse=True)
        
        latest = tracker_files[0]
        
        logger.info(f"\n{'='*70}")
        logger.info(f"SELECTED LATEST EDEN TRACKER")
        logger.info(f"{'='*70}")
        logger.info(f"File: {latest['filename']}")
        logger.info(f"Path: {latest['key']}")
        if latest['extracted_date']:
            logger.info(f"Date: {latest['extracted_date'].strftime('%B %d, %Y')}")
        logger.info(f"Last Modified: {latest['last_modified']}")
        logger.info(f"Size: {latest['size']:,} bytes")
        logger.info(f"{'='*70}\n")
        
        return latest['key']
        
    except Exception as e:
        logger.error(f"Error searching for Eden tracker: {str(e)}")
        raise

# ================= DYNAMIC MONTH CALCULATION =================
import re
from datetime import datetime, timedelta
from typing import List, Tuple

# ================= DYNAMIC MONTH CALCULATION =================
def calculate_eden_months_and_targets(tracker_key: str) -> Tuple[List[str], str, datetime]:
    """
    Dynamically calculates months to process and target month 
    based on tracker date and rules:
      1. If tracker is in September → include June, July, August.
      2. Otherwise → include previous, current, and next months.
    """
    # Extract date from tracker_key
    date_pattern = r'(\d{1,2})-(\d{1,2})-(\d{4})'
    match = re.search(date_pattern, tracker_key)
    
    if match:
        day, month, year = match.groups()
        try:
            tracker_date = datetime.strptime(f"{day}-{month}-{year}", "%d-%m-%Y")
        except ValueError:
            tracker_date = datetime.now()
    else:
        tracker_date = datetime.now()
    
    # Month mapping
    months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ]
    
    tracker_month = tracker_date.month
    tracker_year = tracker_date.year
    
    included_months = []

    # 🔹 Rule 1: September tracker → June, July, August
    if tracker_month == 9:
        included_months = ['June', 'July', 'August']
    
    # 🔹 Rule 2: Other months → previous, current, next
    else:
        prev_month_date = tracker_date.replace(day=1) - timedelta(days=1)
        next_month_year = tracker_year + (1 if tracker_month == 12 else 0)
        next_month_num = 1 if tracker_month == 12 else tracker_month + 1
        next_month_date = tracker_date.replace(year=next_month_year, month=next_month_num, day=1)
        
        included_months = [
            months[prev_month_date.month - 1],
            months[tracker_month - 1],
            months[next_month_date.month - 1],
        ]
    
    # ✅ Dynamic start month — first in the included list
    start_month = included_months[0]
    target_month = included_months[-1]

    return included_months, target_month, tracker_date


# ================= DYNAMIC KRA STRUCTURE =================
def discover_months_in_kra(kra_ws):
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
    return months

def discover_towers_in_kra(kra_ws):
    towers = {}
    for row in range(5, 25):
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
            towers[key] = {
                'parent_row_1': row,
                'parent_row_2': row + 1,
                'child_row': row + 2
            }
    return towers

def discover_tracker_sheets(tracker_wb):
    """Map all NTA sections to the same NTA sheet"""
    sheet_mapping = {}
    nta_sheet_name = None
    
    for sheet_name in tracker_wb.sheetnames:
        sheet_clean = sheet_name.strip()
        tower_match = re.search(r'Tower\s*(\d+)', sheet_clean, re.IGNORECASE)
        if tower_match:
            sheet_mapping[f"Tower {tower_match.group(1)}"] = sheet_clean
        elif re.search(r'non.*tower.*area', sheet_clean, re.IGNORECASE):
            nta_sheet_name = sheet_clean
    
    # Map ALL NTA-X entries to the same NTA sheet
    if nta_sheet_name:
        # Add base NTA mapping
        sheet_mapping["NTA"] = nta_sheet_name
        # Add mappings for both formats: NTA-1, NTA-01, NTA-001, etc.
        for i in range(1, 20):  # Support up to NTA-19
            sheet_mapping[f"NTA-{i}"] = nta_sheet_name           # NTA-1
            sheet_mapping[f"NTA-{i:02d}"] = nta_sheet_name       # NTA-01
            sheet_mapping[f"NTA-{i:03d}"] = nta_sheet_name       # NTA-001
    
    return sheet_mapping


def discover_tracker_sheets(tracker_wb):
    """Map all NTA sections to the same NTA sheet"""
    sheet_mapping = {}
    nta_sheet_name = None
    
    for sheet_name in tracker_wb.sheetnames:
        sheet_clean = sheet_name.strip()
        tower_match = re.search(r'Tower\s*(\d+)', sheet_clean, re.IGNORECASE)
        if tower_match:
            sheet_mapping[f"Tower {tower_match.group(1)}"] = sheet_clean
        elif re.search(r'non.*tower.*area', sheet_clean, re.IGNORECASE):
            nta_sheet_name = sheet_clean
    
    # Map ALL NTA-X entries to the same NTA sheet
    if nta_sheet_name:
        # Add base NTA mapping
        sheet_mapping["NTA"] = nta_sheet_name
        # Add mappings for both formats: NTA-1, NTA-01, NTA-001, etc.
        for i in range(1, 20):  # Support up to NTA-19
            sheet_mapping[f"NTA-{i}"] = nta_sheet_name           # NTA-1
            sheet_mapping[f"NTA-{i:02d}"] = nta_sheet_name       # NTA-01
            sheet_mapping[f"NTA-{i:03d}"] = nta_sheet_name       # NTA-001
    
    return sheet_mapping


def get_nta_section_bounds(tracker_ws, nta_key):
    """
    For NTA sections, we don't need strict bounds since we'll search
    by matching activities from KRA. Just return the general NTA area.
    """
    # Extract the NTA number
    nta_match = re.match(r'NTA-(\d+)', nta_key, re.IGNORECASE)
    if not nta_match:
        logger.warning(f"Could not parse NTA key: {nta_key}")
        return None, None
    
    max_row = tracker_ws.max_row
    
    # For NTA, we'll search the entire sheet since activities might be anywhere
    # The activity matching logic will handle finding the right ones
    logger.info(f"[{nta_key}] Will search entire NTA sheet for matching activities")
    
    # Return broad range - the find_activity_in_tracker will use KRA activities to match
    return 1, max_row


# ================== KRA TO TRACKER MAPPING ==================
def get_activities_from_kra(tower, month_col, kra_ws, tower_structure):
    if tower not in tower_structure:
        return []
    rows = tower_structure[tower]
    activities = []
    for key in ['parent_row_1','parent_row_2','child_row']:
        val = kra_ws.cell(row=rows[key], column=month_col).value
        if val and str(val).strip():
            activities.append(str(val).strip())
    return activities

def find_correct_percentage_column(tracker_ws, row, task_name):
    check_columns = [PCT_COL] + PCT_COL_ALT
    for col in check_columns:
        try:
            val = tracker_ws.cell(row=row, column=col).value
            if val is not None:
                s = str(val).strip()
                if s.endswith('%'):
                    s = s.replace('%','').strip()
                    if s.replace('.','').isdigit():
                        return col, val
                elif s.replace('.','').isdigit():
                    return col, val
        except:
            continue
    return None, None

def parse_percentage_value(pct_val):
    if isinstance(pct_val, (int,float)):
        if 0<=pct_val<=1:
            return pct_val*100
        return float(pct_val)
    s = str(pct_val).replace('%','').strip()
    if s:
        f = float(s)
        if 0<=f<=1:
            f*=100
        return f
    return 0.0

def calculate_match_score(text1, text2):
    t1,t2=text1.lower().strip(),text2.lower().strip()
    if t1==t2: return 1.0
    if ' '.join(t1.split())==' '.join(t2.split()): return 1.0
    if t1 in t2 or t2 in t1: return 0.9
    w1,w2=set(t1.split()),set(t2.split())
    if not w1 or not w2: return 0.0
    return len(w1&w2)/len(w1|w2)

def find_activity_in_tracker(tracker_ws, parent_activities, child_activity, tower=None):
    """
    Find activity and extract percentage, responsible person, and delay info.
    For NTA: ONLY match activities that are explicitly in the KRA.
    """
    max_row = tracker_ws.max_row
    
    if not parent_activities or not child_activity:
        logger.warning(f"[{tower}] Missing activities - parent: {len(parent_activities) if parent_activities else 0}, child: {'Yes' if child_activity else 'No'}")
        return 0.0, "", ""
    
    parent_list = [str(p).strip().lower() for p in parent_activities if p]
    child_clean = str(child_activity).strip().lower()
    
    logger.info(f"\n{'='*70}")
    logger.info(f"[{tower}] Searching for KRA activities in tracker:")
    logger.info(f"  Parent activities from KRA: {parent_list}")
    logger.info(f"  Child activity from KRA: {child_clean}")
    logger.info(f"{'='*70}")
    
    is_nta = tower and tower.startswith('NTA')
    
    # Determine search range
    if is_nta:
        start_row, end_row = get_nta_section_bounds(tracker_ws, tower)
        if not start_row or not end_row:
            logger.warning(f"[{tower}] Could not determine section bounds, skipping")
            return 0.0, "", ""
        search_start, search_end = start_row, end_row
    else:
        search_start, search_end = 2, max_row
    
    logger.info(f"  Search range: rows {search_start} to {search_end}")
    
    # Find all bold rows (headers) within the search range
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
    
    logger.info(f"  Found {len(bold_rows)} bold rows (potential headers)")
    
    # NEW LOGIC FOR NTA: Only match if ALL parent activities match bold rows
    # This ensures we're in the correct section that matches the KRA structure
    
    if is_nta:
        logger.info(f"\n  NTA Mode: Looking for exact match of KRA parent structure...")
        
        # Find groups where ALL parent activities appear as consecutive bold rows
        matching_groups = []
        
        for i, bold_row in enumerate(bold_rows):
            # Check if this bold row matches the FIRST parent activity
            first_parent = parent_list[0] if parent_list else None
            if not first_parent:
                continue
            
            first_match_score = calculate_match_score(bold_row['text_lower'], first_parent)
            
            if first_match_score >= 0.7:
                logger.info(f"    Potential match at row {bold_row['row']}: '{bold_row['text'][:50]}' matches '{first_parent}'")
                
                # Now check if subsequent parent activities also match in order
                found_parents = {parent_list[0]: bold_row['row']}
                
                # Look ahead for remaining parent activities
                for j in range(1, len(parent_list)):
                    parent_to_find = parent_list[j]
                    found = False
                    
                    # Check next few bold rows
                    for k in range(i + 1, min(i + 10, len(bold_rows))):
                        check_bold = bold_rows[k]
                        score = calculate_match_score(check_bold['text_lower'], parent_to_find)
                        
                        if score >= 0.7:
                            found_parents[parent_to_find] = check_bold['row']
                            logger.info(f"      Matched '{parent_to_find}' at row {check_bold['row']}")
                            found = True
                            break
                    
                    if not found:
                        logger.info(f"      Could NOT find parent '{parent_to_find}'")
                        break
                
                # If ALL parent activities were found, this is a valid group
                if len(found_parents) == len(parent_list):
                    logger.info(f"    ✓ ALL parent activities matched! Valid group found.")
                    
                    # Define search range for child activity
                    group_start = min(found_parents.values())
                    
                    # Find end of group (next unrelated bold row)
                    group_end = search_end
                    for next_bold in bold_rows:
                        if next_bold['row'] > max(found_parents.values()):
                            # Check if this is a new unrelated section
                            if all(calculate_match_score(next_bold['text_lower'], p) < 0.5 for p in parent_list):
                                group_end = next_bold['row'] - 1
                                break
                    
                    matching_groups.append({
                        'start': group_start,
                        'end': group_end,
                        'parent_rows': sorted(found_parents.values())
                    })
                    
                    logger.info(f"    Group range: rows {group_start} to {group_end}")
                    
                    # For NTA, take only the first exact match
                    break
        
        if not matching_groups:
            logger.warning(f"  ✗ No matching group found where ALL KRA parent activities appear in tracker")
            return 0.0, "", ""
    
    else:
        # Original logic for Towers (keep existing behavior)
        matching_groups = []
        
        for i, b in enumerate(bold_rows):
            if any(calculate_match_score(b['text_lower'], p) >= 0.7 for p in parent_list):
                search_range = bold_rows[i:min(i + 10, len(bold_rows))]
                found = {p: None for p in parent_list}
                
                for br in search_range:
                    for p in parent_list:
                        if found[p] is None and calculate_match_score(br['text_lower'], p) >= 0.7:
                            found[p] = br['row']
                
                if all(v is not None for v in found.values()):
                    s = min(found.values())
                    e = search_end
                    
                    for next_b in bold_rows:
                        if next_b['row'] > max(found.values()):
                            if all(calculate_match_score(next_b['text_lower'], p) < 0.5 for p in parent_list):
                                e = next_b['row'] - 1
                                break
                    
                    matching_groups.append({
                        'start': s,
                        'end': e,
                        'parent_rows': sorted(found.values())
                    })
    
    # Search for child activity in the matching groups
    for g_idx, g in enumerate(matching_groups):
        logger.info(f"\n  Searching for child activity in group {g_idx + 1} (rows {g['start']} to {g['end']})...")
        
        best_row = None
        best_score = 0
        
        for row in range(g['start'], g['end'] + 1):
            # Skip bold rows (they are headers)
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
                if score >= 0.6:  # Log good candidates
                    logger.info(f"    Candidate row {row} (score {score:.2f}): {str(val).strip()[:60]}")
        
        if best_row and best_score >= 0.75:
            logger.info(f"\n  ✓✓✓ MATCH FOUND ✓✓✓")
            logger.info(f"  Row {best_row} (score: {best_score:.2f})")
            
            val = tracker_ws.cell(row=best_row, column=TASK_NAME_COL).value
            logger.info(f"  Activity: {str(val).strip()}")
            
            col, pct = find_correct_percentage_column(tracker_ws, best_row, val)
            
            if pct is not None:
                resp = tracker_ws.cell(row=best_row, column=RESPONSIBLE_COL).value or ""
                delay = tracker_ws.cell(row=best_row, column=DELAY_COL).value or ""
                parsed_pct = parse_percentage_value(pct)
                logger.info(f"  Percentage: {parsed_pct:.1f}%")
                logger.info(f"  Responsible: {resp}")
                logger.info(f"  Delay: {delay}")
                logger.info(f"{'='*70}\n")
                return parsed_pct, str(resp).strip(), str(delay).strip()
        else:
            logger.warning(f"    Best match score {best_score:.2f} below threshold (0.75)")
    
    logger.warning(f"  ✗ No valid child activity match found")
    logger.info(f"{'='*70}\n")
    return 0.0, "", ""


# ================= HELPER DIAGNOSTIC FUNCTION =================
# Add this function to your code, right after the find_activity_in_tracker function

def diagnose_nta_sheet(tracker_wb, sheet_mapping, tower_structure):
    """
    Diagnostic function to understand the NTA sheet structure.
    Call this in main() to see what's happening.
    """
    logger.info("\n" + "="*70)
    logger.info("NTA SHEET DIAGNOSTIC")
    logger.info("="*70)
    
    # Get all NTA towers from structure
    nta_towers = sorted([t for t in tower_structure.keys() if t.startswith('NTA')])
    logger.info(f"NTA towers in KRA structure: {nta_towers}")
    
    # Check sheet mapping
    if nta_towers and nta_towers[0] in sheet_mapping:
        nta_sheet_name = sheet_mapping[nta_towers[0]]
        logger.info(f"NTA sheet name: '{nta_sheet_name}'")
        
        if nta_sheet_name in tracker_wb.sheetnames:
            ws = tracker_wb[nta_sheet_name]
            logger.info(f"Sheet has {ws.max_row} rows\n")
            
            # Show BOLD rows (these are section headers)
            logger.info("BOLD ROWS (Section Headers) in NTA sheet:")
            logger.info("-" * 70)
            bold_count = 0
            for row in range(1, min(ws.max_row + 1, 200)):
                val = ws.cell(row=row, column=TASK_NAME_COL).value
                if val:
                    try:
                        bold = ws.cell(row=row, column=TASK_NAME_COL).font.bold
                    except:
                        bold = False
                    
                    if bold:
                        bold_count += 1
                        val_str = str(val).strip()
                        logger.info(f"  Row {row:3d}: {val_str}")
            
            logger.info(f"\nTotal bold rows found: {bold_count}")
            logger.info("-" * 70)
            
            # Test section bounds for each NTA
            logger.info("\nTesting section detection for each NTA:")
            for nta_key in nta_towers:
                start, end = get_nta_section_bounds(ws, nta_key)
                if start and end:
                    logger.info(f"\n✓ {nta_key} SUCCESSFULLY MAPPED:")
                    logger.info(f"  Rows: {start} to {end}")
                    
                    # Show header
                    header_val = ws.cell(row=start, column=TASK_NAME_COL).value
                    logger.info(f"  Header: '{header_val}'")
                    
                    # Show first few activity rows
                    logger.info(f"  First few activities:")
                    count = 0
                    for row in range(start + 1, min(start + 8, end + 1)):
                        val = ws.cell(row=row, column=TASK_NAME_COL).value
                        if val:
                            try:
                                is_bold = ws.cell(row=row, column=TASK_NAME_COL).font.bold
                            except:
                                is_bold = False
                            
                            if not is_bold:  # Only show non-bold (actual activities)
                                count += 1
                                pct_val = ws.cell(row=row, column=PCT_COL).value
                                logger.info(f"    Row {row}: {str(val)[:50]} | %: {pct_val}")
                                if count >= 5:
                                    break
                else:
                    logger.error(f"\n✗ {nta_key}: Could not determine bounds!")
        else:
            logger.error(f"Sheet '{nta_sheet_name}' not found in workbook!")
    else:
        logger.error("No NTA sheet mapping found!")
    
    logger.info("="*70 + "\n")

def calculate_month_data(tower, month, month_col, kra_ws, tracker_wb, sheet_mapping, tower_structure, current_month):
    activities=get_activities_from_kra(tower,month_col,kra_ws,tower_structure)
    if not activities:
        return {'activities_text':"",'percentage':0.0,'progress_status':"No Progress",'responsible':"",'delay':""}
    activities_text='\n'.join(activities)
    mo=['January','February','March','April','May','June','July','August','September','October','November','December']
    is_current_or_past=month in mo and current_month in mo and mo.index(month)<=mo.index(current_month)
    if is_current_or_past:
        tracker_sheet=sheet_mapping.get(tower)
        if tracker_sheet and tracker_sheet in tracker_wb.sheetnames:
            tracker_ws=tracker_wb[tracker_sheet]
            parent_activities=activities[:-1] if len(activities)>1 else []
            child_activity=activities[-1]
            pct,responsible,delay=find_activity_in_tracker(tracker_ws,parent_activities,child_activity,tower)
        else: pct,responsible,delay=0.0,"",""
    else: pct,responsible,delay=0.0,"",""
    progress_status=f"Achieved-{activities[-1]}" if pct>0 else "No Progress"
    return {'activities_text':activities_text,'percentage':pct,'progress_status':progress_status,'responsible':responsible,'delay':delay}

# ================= COS =================
def init_cos():
    return ibm_boto3.client("s3",ibm_api_key_id=COS_API_KEY,ibm_service_instance_id=COS_CRN,
                            config=Config(signature_version="oauth"),endpoint_url=COS_ENDPOINT)

def download_file_bytes(cos,key):
    return cos.get_object(Bucket=BUCKET,Key=key)["Body"].read()

# ================= EXCEL FORMATTING =================
def format_excel_report(ws,df):
    header_font=Font(bold=True,size=10);title_font=Font(bold=True,size=14);date_font=Font(size=10,color="666666");data_font=Font(size=9)
    center_align=Alignment(horizontal="center",vertical="center",wrap_text=True)
    left_align=Alignment(horizontal="left",vertical="center",wrap_text=True)
    border=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    header_fill=PatternFill(start_color="D9E2F3",end_color="D9E2F3",fill_type="solid")
    
    ws.merge_cells(f'A1:{get_column_letter(len(df.columns))}1');ws['A1'].font=title_font;ws['A1'].alignment=center_align
    ws.merge_cells(f'A2:{get_column_letter(len(df.columns))}2');ws['A2'].font=date_font;ws['A2'].alignment=center_align
    for cell in ws[4]: cell.font=header_font;cell.alignment=center_align;cell.border=border;cell.fill=header_fill
    for row in ws.iter_rows(min_row=5,max_row=ws.max_row):
        for col_idx,cell in enumerate(row,1):
            cell.border=border;cell.font=data_font
            header_val=ws.cell(row=4,column=col_idx).value or ''
            cell.alignment=left_align if any(kw in str(header_val) for kw in ['Tower','Target','Activity','Progress','Responsible','Delay']) else center_align
    for col_idx in range(1,len(df.columns)+1):
        max_length=10
        for row in ws.iter_rows(min_row=4,max_row=ws.max_row,min_col=col_idx,max_col=col_idx):
            for cell in row:
                if cell.value: max_length=max(max_length,len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width=min(max(max_length+2,10),35)
    ws.row_dimensions[1].height=25;ws.row_dimensions[2].height=20;ws.row_dimensions[4].height=40
    for row_idx in range(5,ws.max_row+1): ws.row_dimensions[row_idx].height=35
    
    
 # ================= BOLD TEXT ANALYSIS =================
def calculate_bold_text_percentage(kra_ws):
    total_text_cells = 0
    bold_text_cells = 0
    
    for row in kra_ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.strip():
                total_text_cells += 1
                try:
                    if cell.font.bold:
                        bold_text_cells += 1
                except:
                    pass
    
    percentage = (bold_text_cells / total_text_cells * 100) if total_text_cells > 0 else 0
    logger.info(f"\n{'='*60}")
    logger.info("KRA BOLD TEXT ANALYSIS")
    logger.info(f"Total text cells: {total_text_cells}")
    logger.info(f"Bold text cells:  {bold_text_cells}")
    logger.info(f"Percentage bold:  {percentage:.2f}%")
    logger.info(f"{'='*60}\n")
    
    return percentage
   

# ================= MAIN =================
def main():
    logger.info("Starting Dynamic Eden KRA Report Generator...")
    try:
        cos=init_cos()
        logger.info("Downloading KRA file...")
        kra_wb=load_workbook(filename=BytesIO(download_file_bytes(cos,KRA_KEY)),data_only=True)
        kra_ws=kra_wb.active
        
        bold_pct = calculate_bold_text_percentage(kra_ws)
        logger.info(f"Bold text percentage in KRA sheet: {bold_pct:.2f}%")

        global EDEN_TRACKER_KEY
        EDEN_TRACKER_KEY=find_latest_eden_tracker(cos,BUCKET,EDEN_TRACKER_FOLDER)
        if not EDEN_TRACKER_KEY: logger.error("Could not find Eden tracker. Exiting..."); return
        months_to_process,target_month,tracker_date=calculate_eden_months_and_targets(EDEN_TRACKER_KEY)
        logger.info(f"Downloading tracker: {EDEN_TRACKER_KEY}")
        tracker_wb=load_workbook(filename=BytesIO(download_file_bytes(cos,EDEN_TRACKER_KEY)),data_only=True)
        months=discover_months_in_kra(kra_ws)
        tower_structure=discover_towers_in_kra(kra_ws)
        sheet_mapping=discover_tracker_sheets(tracker_wb)
         # Check if there are any NTA towers and diagnose them
        if any(t.startswith('NTA') for t in tower_structure.keys()):
            diagnose_nta_sheet(tracker_wb, sheet_mapping, tower_structure)
        
        current_month=tracker_date.strftime("%B")
        results=[]
        for tower in tower_structure.keys():
            row_data={'Tower':tower}
            target_col=months.get(target_month)
            row_data[f'Target till {target_month}']= '\n'.join(get_activities_from_kra(tower,target_col,kra_ws,tower_structure)) if target_col else ""
            year=tracker_date.year
            total_weighted=0; weightage=100 if not tower.startswith('NTA') else 50
            months_count = len(months_to_process)  # total months considered
            for month in months_to_process:
                month_col=months.get(month)
                if not month_col: continue
                month_data=calculate_month_data(tower,month,month_col,kra_ws,tracker_wb,sheet_mapping,tower_structure,current_month)
                row_data[f"Activity- Target to be complete by {month} {year}"]=month_data['activities_text']
                row_data[f"% work done against Target- {month} Status"]=f"{month_data['percentage']:.0f}%"
                row_data[f"Progress-{month}"]=month_data['progress_status']
                row_data[f"Responsible Person-{month}"]=""
                row_data[f"Delay Reasons-{month}"]=""
                total_weighted += round(month_data['percentage'] * weightage) / (100 * months_count)
            row_data['Weightage']=weightage
            row_data['Weighted Work done against Target']=f"{total_weighted:.1f}%"
            results.append(row_data)
        df=pd.DataFrame(results)
        filename=f"Eden_Progress_Against_Milestones_{tracker_date.strftime(('%Y-%m-%d'))}.xlsx"
        wb=Workbook(); ws=wb.active; ws.title="Eden- Progress Against Milestones"
        ws.append(["Eden- Progress Against Milestones"])
        ws.append([f"Report Generated on: {datetime.now().strftime('%B %d, %Y')}"])
        ws.append([])
        for r in dataframe_to_rows(df,index=False,header=True): ws.append(r)
        format_excel_report(ws,df)
        wb.save(filename)
        logger.info(f"\nReport saved: {filename}")
    except Exception as e:
        logger.error(f"Error: {str(e)}",exc_info=True)
        raise

if __name__=="__main__":
    main()
