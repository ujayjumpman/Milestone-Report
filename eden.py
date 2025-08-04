

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

# All months to show activities for
MONTHS = ["June", "July", "August"]
MONTH_COLUMNS = {"June": 2, "July": 3, "August": 4}

# Current month for which we have tracker data (change this as needed)
CURRENT_TRACKER_MONTH = "June"  # Only this month will have percentage calculations

# Tower to Sheet mapping for tracker file
TOWER_SHEET_MAP = {
    "Tower 4": "Tower 4",
    "Tower 5": "Tower 5", 
    "Tower 6": "Tower 6",
    "Tower 7": "Tower 7",
    "NTA-01": "Non Tower Area",
    "NTA-02": "Non Tower Area",
}

# HARDCODED VALUES FOR SPECIFIC TOWERS
HARDCODED_PERCENTAGES = {
    "Tower 4": 55.0,
    "Tower 6": 60.0, 
    "NTA-01": 0.0,
    "NTA-02": 0.0
}

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

def get_display_tower_name(tower):
    """Convert tower key to display name for the Tower column"""
    if tower.startswith("NTA"):
        return "Non Tower Area"
    return tower

# ============= IMPROVED KRA PARSING ==================

def find_all_towers_in_kra(kra_ws):
    """Find towers in column A only."""
    max_row = kra_ws.max_row
    towers_found = []
    
    # Look only in column A for tower names
    for row in range(1, max_row + 1):
        cell_val = kra_ws.cell(row=row, column=1).value  # Column A only
        if cell_val:
            cell_str = str(cell_val).strip()
            # Check if this looks like a tower
            if (cell_str.startswith("Tower") or cell_str.startswith("NTA-")) and cell_str not in towers_found:
                # Skip generic headers
                if cell_str != "Tower" and len(cell_str) > 5:
                    towers_found.append(cell_str)
                    logger.info(f"Found tower: '{cell_str}' at row {row}")
    
    return towers_found

def extract_tower_activities_improved(tower, kra_ws):
    """
    Extract activities for a tower with improved parsing.
    Returns structured data for each month.
    """
    max_row = kra_ws.max_row
    
    # Find the tower row
    tower_row = None
    for row in range(1, max_row + 1):
        cell_val = kra_ws.cell(row=row, column=1).value  # Column A
        if cell_val and str(cell_val).strip() == tower:
            tower_row = row
            logger.info(f"Found {tower} at row {row}")
            break
    
    if not tower_row:
        logger.warning(f"Tower {tower} not found")
        return None
    
    tower_data = {}
    
    # Extract activities for each month
    for month, col in MONTH_COLUMNS.items():
        month_lower = month.lower()
        
        # Collect all non-empty activities from the tower's section
        activities = []
        for offset in range(0, 8):  # Check more rows to capture all activities
            check_row = tower_row + offset
            if check_row > max_row:
                break
                
            cell_val = kra_ws.cell(row=check_row, column=col).value
            if cell_val and str(cell_val).strip():
                activity = str(cell_val).strip()
                # Skip month names and tower names
                if activity not in MONTHS and activity != tower and activity != "Activity":
                    activities.append(activity)
        
        if not activities:
            logger.warning(f"{tower} - {month}: No activities found")
            continue
        
        # Parse the activities into hierarchy
        hierarchy = parse_activities_to_hierarchy(activities)
        
        # Store structured data
        tower_data[f'parent_{month_lower}'] = hierarchy['parent']
        tower_data[f'sub_parent_{month_lower}'] = hierarchy['sub_parent']
        tower_data[f'child_{month_lower}'] = hierarchy['child']
        tower_data[f'target_{month_lower}'] = hierarchy['target_display']
        tower_data[f'activity_{month_lower}'] = hierarchy['child']
        
        logger.info(f"{tower} - {month}:")
        logger.info(f"  Activities found: {activities}")
        logger.info(f"  Parent: '{hierarchy['parent']}'")
        logger.info(f"  Sub-Parent: '{hierarchy['sub_parent']}'") 
        logger.info(f"  Child: '{hierarchy['child']}'")
    
    return tower_data

def parse_activities_to_hierarchy(activities):
    """
    Parse activities into 3-level hierarchy based on construction context.
    Assumes activities are listed in order: Parent -> Sub-Parent -> Child
    """
    if not activities:
        return {'parent': '', 'sub_parent': '', 'child': '', 'target_display': ''}
    
    # Clean activities
    cleaned_activities = [act.strip() for act in activities if act.strip()]
    
    if len(cleaned_activities) == 1:
        # Single activity - treat as child with no parent/sub-parent
        return {
            'parent': '',
            'sub_parent': '', 
            'child': cleaned_activities[0],
            'target_display': cleaned_activities[0]
        }
    elif len(cleaned_activities) == 2:
        # Two activities - first is parent, second is child
        return {
            'parent': cleaned_activities[0],
            'sub_parent': '',
            'child': cleaned_activities[1],
            'target_display': cleaned_activities[0]
        }
    elif len(cleaned_activities) >= 3:
        # Three or more activities - use first three as hierarchy
        parent = cleaned_activities[0]
        sub_parent = cleaned_activities[1]
        child = cleaned_activities[2]
        
        return {
            'parent': parent,
            'sub_parent': sub_parent,
            'child': child,
            'target_display': f"{parent} - {sub_parent}"
        }
    
    return {'parent': '', 'sub_parent': '', 'child': '', 'target_display': ''}

# ============= IMPROVED TRACKER MATCHING ==================

def find_activity_percentage_improved(tracker_ws, parent, sub_parent, child):
    """
    Improved activity matching using 3-level hierarchy.
    Look for: Parent (bold) -> Sub-Parent (bold) -> Child (non-bold target)
    """
    TASK_NAME_COL = 4  # D column
    PCT_COL = 7        # G column
    max_row = tracker_ws.max_row
    
    logger.info(f"Searching hierarchy: Parent='{parent}' -> Sub-Parent='{sub_parent}' -> Child='{child}'")
    
    # If no hierarchy, do direct search
    if not parent and not sub_parent:
        return find_direct_match(tracker_ws, child)
    
    # Step 1: Find all bold activities (parents and sub-parents)
    bold_activities = get_bold_activities(tracker_ws)
    
    # Step 2: Find matching parent section
    parent_matches = find_parent_matches(bold_activities, parent)
    
    if not parent_matches:
        logger.warning(f"No parent matches found for '{parent}', trying direct search")
        return find_direct_match(tracker_ws, child)
    
    # Step 3: For each parent match, look for sub-parent and then child
    for parent_row, parent_text in parent_matches:
        logger.info(f"Checking parent section: Row {parent_row} '{parent_text}'")
        
        # Find the end of this parent section
        section_end = find_section_end(parent_row, bold_activities)
        
        if sub_parent:
            # Look for sub-parent within this section
            sub_parent_match = find_sub_parent_in_section(tracker_ws, parent_row, section_end, sub_parent)
            
            if sub_parent_match:
                sub_parent_row, sub_parent_text = sub_parent_match
                logger.info(f"Found sub-parent: Row {sub_parent_row} '{sub_parent_text}'")
                
                # Find child under this sub-parent
                child_pct = find_child_under_section(tracker_ws, sub_parent_row, bold_activities, child)
                if child_pct is not None:
                    return child_pct
            else:
                logger.info(f"No sub-parent found, searching children directly under parent")
        
        # If no sub-parent or sub-parent not found, search children directly under parent
        child_pct = find_child_under_section(tracker_ws, parent_row, bold_activities, child)
        if child_pct is not None:
            return child_pct
    
    # Fallback to direct search
    logger.warning(f"Hierarchy search failed, trying direct search for '{child}'")
    return find_direct_match(tracker_ws, child)

def get_bold_activities(tracker_ws):
    """Get all bold activities with their row numbers."""
    TASK_NAME_COL = 4
    max_row = tracker_ws.max_row
    bold_activities = []
    
    for row in range(2, max_row + 1):
        cell = tracker_ws.cell(row=row, column=TASK_NAME_COL)
        if cell.value:
            try:
                if cell.font and cell.font.bold:
                    activity_text = str(cell.value).strip()
                    if activity_text:
                        bold_activities.append((row, activity_text))
                        logger.debug(f"Bold activity found: Row {row} '{activity_text}'")
            except:
                pass
    
    logger.info(f"Found {len(bold_activities)} bold activities")
    return bold_activities

def find_parent_matches(bold_activities, parent):
    """Find bold activities that match the parent."""
    if not parent:
        return []
    
    matches = []
    parent_lower = normalize_text(parent)
    
    for row, activity in bold_activities:
        activity_lower = normalize_text(activity)
        
        # Check for match using various criteria
        if is_activity_match(parent_lower, activity_lower):
            matches.append((row, activity))
            logger.info(f"Parent match: Row {row} '{activity}'")
    
    return matches

def find_sub_parent_in_section(tracker_ws, parent_row, section_end, sub_parent):
    """Find sub-parent bold activity within parent section."""
    if not sub_parent:
        return None
    
    TASK_NAME_COL = 4
    sub_parent_lower = normalize_text(sub_parent)
    
    for row in range(parent_row + 1, section_end):
        cell = tracker_ws.cell(row=row, column=TASK_NAME_COL)
        if cell.value:
            try:
                if cell.font and cell.font.bold:
                    activity_text = str(cell.value).strip()
                    activity_lower = normalize_text(activity_text)
                    
                    if is_activity_match(sub_parent_lower, activity_lower):
                        logger.info(f"Sub-parent match: Row {row} '{activity_text}'")
                        return (row, activity_text)
            except:
                pass
    
    return None

def find_child_under_section(tracker_ws, section_start, bold_activities, child):
    """Find child activity (non-bold) under a section."""
    TASK_NAME_COL = 4
    PCT_COL = 7
    
    if not child:
        return None
    
    # Find section end
    section_end = find_section_end(section_start, bold_activities)
    
    child_lower = normalize_text(child)
    best_match = None
    best_score = 0
    
    logger.info(f"Searching for child '{child}' in rows {section_start + 1} to {section_end - 1}")
    
    for row in range(section_start + 1, section_end):
        cell = tracker_ws.cell(row=row, column=TASK_NAME_COL)
        if not cell.value:
            continue
        
        # Skip bold activities (these are sub-sections)
        try:
            if cell.font and cell.font.bold:
                continue
        except:
            pass
        
        activity_text = str(cell.value).strip()
        activity_lower = normalize_text(activity_text)
        pct_val = tracker_ws.cell(row=row, column=PCT_COL).value
        
        # Calculate match score
        score = calculate_match_score(child_lower, activity_lower)
        
        if score > best_score and score >= 0.8:  # Lower threshold for child matching
            best_match = (row, activity_text, pct_val, score)
            best_score = score
            logger.info(f"  Child match candidate: Row {row} '{activity_text}' - Score: {score:.3f}, %: {pct_val}")
    
    if best_match:
        row, activity_text, pct_val, score = best_match
        percentage = parse_percentage(pct_val)
        logger.info(f"Found child: Row {row} '{activity_text}' - {percentage}%")
        return percentage
    
    return None

def find_section_end(start_row, bold_activities):
    """Find where a section ends (next bold activity at same or higher level)."""
    for row, _ in bold_activities:
        if row > start_row:
            return row
    return 999999  # End of sheet

def find_direct_match(tracker_ws, activity_name):
    """Direct search fallback when hierarchy search fails."""
    TASK_NAME_COL = 4
    PCT_COL = 7
    max_row = tracker_ws.max_row
    
    if not activity_name:
        return 0.0
    
    activity_lower = normalize_text(activity_name)
    best_match = None
    best_score = 0
    
    logger.info(f"Direct search for: '{activity_name}'")
    
    for row in range(2, max_row + 1):
        cell = tracker_ws.cell(row=row, column=TASK_NAME_COL)
        if not cell.value:
            continue
            
        task_text = str(cell.value).strip()
        task_lower = normalize_text(task_text)
        pct_val = tracker_ws.cell(row=row, column=PCT_COL).value
        
        score = calculate_match_score(activity_lower, task_lower)
        
        if score > best_score and score >= 0.8:
            best_match = (row, task_text, pct_val, score)
            best_score = score
    
    if best_match:
        row, task_text, pct_val, score = best_match
        percentage = parse_percentage(pct_val)
        logger.info(f"Direct match: Row {row} '{task_text}' - Score: {score:.3f}, %: {percentage}%")
        return percentage
    
    logger.warning(f"No match found for '{activity_name}'")
    return 0.0

def normalize_text(text):
    """Normalize text for comparison."""
    if not text:
        return ""
    text = str(text).lower().strip()
    # Replace common separators
    text = text.replace('&', 'and').replace('/', ' ').replace('-', ' ')
    # Remove extra spaces
    while '  ' in text:
        text = text.replace('  ', ' ')
    return text.strip()

def is_activity_match(text1, text2):
    """Check if two activities match using various criteria."""
    if not text1 or not text2:
        return False
    
    # Exact match
    if text1 == text2:
        return True
    
    # Substring match
    if text1 in text2 or text2 in text1:
        return True
    
    # Word overlap
    words1 = set(text1.split())
    words2 = set(text2.split())
    
    if not words1 or not words2:
        return False
    
    common_words = words1.intersection(words2)
    overlap_ratio = len(common_words) / min(len(words1), len(words2))
    
    return overlap_ratio >= 0.6

def calculate_match_score(target, candidate):
    """Calculate detailed match score between activities."""
    if not target or not candidate:
        return 0.0
    
    # Exact match
    if target == candidate:
        return 1.0
    
    # Substring match
    if target in candidate:
        return 0.9 + 0.05 * (len(target) / len(candidate))
    
    if candidate in target:
        return 0.85 + 0.05 * (len(candidate) / len(target))
    
    # Word-based matching
    target_words = set(target.split())
    candidate_words = set(candidate.split())
    
    if not target_words or not candidate_words:
        return 0.0
    
    common_words = target_words.intersection(candidate_words)
    
    if not common_words:
        return 0.0
    
    # Calculate overlap ratio
    target_ratio = len(common_words) / len(target_words)
    candidate_ratio = len(common_words) / len(candidate_words)
    
    return min(target_ratio, candidate_ratio) * 0.8

def parse_percentage(pct_val):
    """Parse percentage value from tracker."""
    if pct_val is None:
        return 0.0
    
    try:
        if isinstance(pct_val, (int, float)):
            if 0 <= pct_val <= 1:
                return float(pct_val * 100)
            else:
                return float(pct_val)
        else:
            pct_str = str(pct_val).replace("%", "").strip()
            return float(pct_str)
    except (ValueError, TypeError):
        logger.warning(f"Could not parse percentage: {pct_val}")
        return 0.0

# ============= MAIN PROCESSING ==================

def calculate_percentage_for_tower(tower, tower_data, tracker_wb):
    """Calculate percentage for a tower using improved hierarchy matching."""
    
    # CHECK FOR HARDCODED VALUES FIRST
    if tower in HARDCODED_PERCENTAGES:
        hardcoded_value = HARDCODED_PERCENTAGES[tower]
        logger.info(f"Using hardcoded value for {tower}: {hardcoded_value}%")
        return hardcoded_value
    
    if CURRENT_TRACKER_MONTH not in MONTH_COLUMNS:
        return 0.0
    
    month_lower = CURRENT_TRACKER_MONTH.lower()
    parent = tower_data.get(f'parent_{month_lower}', '')
    sub_parent = tower_data.get(f'sub_parent_{month_lower}', '')
    child = tower_data.get(f'child_{month_lower}', '')
    
    if not child:
        logger.warning(f"No child activity found for {tower}")
        return 0.0
    
    # Get tracker sheet
    tracker_sheetname = TOWER_SHEET_MAP.get(tower, tower)
    if tracker_sheetname not in tracker_wb.sheetnames:
        logger.warning(f"Sheet '{tracker_sheetname}' not found for {tower}")
        available_sheets = list(tracker_wb.sheetnames)
        logger.info(f"Available sheets: {available_sheets}")
        return 0.0
    
    tracker_ws = tracker_wb[tracker_sheetname]
    
    # Find percentage using improved hierarchy matching
    pct = find_activity_percentage_improved(tracker_ws, parent, sub_parent, child)
    logger.info(f"{tower} final result: {pct:.1f}%")
    
    return pct

def main():
    logger.info("Starting Improved Eden KRA Report generation...")
    logger.info(f"Current tracker month: {CURRENT_TRACKER_MONTH}")
    logger.info(f"Hardcoded percentages: {HARDCODED_PERCENTAGES}")
    
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
        
        logger.info(f"Available tracker sheets: {list(tracker_wb.sheetnames)}")
        
        # Find towers
        towers = find_all_towers_in_kra(kra_ws)
        if not towers:
            logger.error("No towers found!")
            return
        
        logger.info(f"Found towers: {towers}")
        
        results = []
        
        for tower in towers:
            logger.info(f"\n{'='*50}")
            logger.info(f"Processing {tower}...")
            logger.info(f"{'='*50}")
            
            # Extract activities using improved parsing
            tower_data = extract_tower_activities_improved(tower, kra_ws)
            if not tower_data:
                logger.warning(f"Skipping {tower} - no data")
                continue
            
            # Build result row
            row_data = {
                "Tower": get_display_tower_name(tower)
            }
            
            # Add activities for all months
            for month in MONTHS:
                month_lower = month.lower()
                target = tower_data.get(f'target_{month_lower}', '')
                activity = tower_data.get(f'activity_{month_lower}', '')
                
                row_data[f"Target {month}"] = target
                row_data[f"Activity {month}"] = activity
            
            # Calculate percentage for current month
            current_pct = calculate_percentage_for_tower(tower, tower_data, tracker_wb)
            
            # Set percentages
            for month in MONTHS:
                if month == CURRENT_TRACKER_MONTH:
                    row_data[f"% Work Done against Target-Till {month}"] = f"{current_pct:.1f}%"
                else:
                    row_data[f"% Work Done against Target-Till {month}"] = ""
            
            row_data[f"Delay Reasons {CURRENT_TRACKER_MONTH}"] = ""
            
            results.append(row_data)
        
        if not results:
            logger.error("No results generated!")
            return
        
        # Create Excel output
        df = pd.DataFrame(results)
        
        # Column ordering
        column_order = ["Tower"]
        for month in MONTHS:
            column_order.extend([f"Target {month}", f"Activity {month}"])
        for month in MONTHS:
            column_order.append(f"% Work Done against Target-Till {month}")
        column_order.append(f"Delay Reasons {CURRENT_TRACKER_MONTH}")
        
        existing_columns = [col for col in column_order if col in df.columns]
        df = df[existing_columns]
        
        filename = f"Eden_KRA_Milestone_Report_Improved ({datetime.now():%Y-%m-%d}).xlsx"
        
        # Create Excel file with formatting
        wb = Workbook()
        ws = wb.active
        ws.title = "Eden KRA Milestone Progress"
        
        # Add title and date
        current_date = datetime.now().strftime("%d-%m-%Y")
        ws.append(["Eden KRA Milestone Progress (Improved)"])
        ws.append([f"Report Generated on: {current_date}"])
        ws.append([])
        
        # Add data
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Format worksheet
        header_font = Font(bold=True, size=11, color="000000")
        title_font = Font(bold=True, size=14, color="000000")
        date_font = Font(bold=False, size=10, color="666666")
        data_font = Font(size=10)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        
        # Format title
        ws.merge_cells(f'A1:{get_column_letter(len(df.columns))}1')
        ws['A1'].font = title_font
        ws['A1'].alignment = center_align
        
        # Format date
        ws.merge_cells(f'A2:{get_column_letter(len(df.columns))}2')
        ws['A2'].font = date_font
        ws['A2'].alignment = center_align
        
        # Format headers
        for cell in ws[4]:
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
            cell.fill = header_fill
        
        # Format data
        for row_idx, row in enumerate(ws.iter_rows(min_row=5, max_row=ws.max_row), 5):
            for col_idx, cell in enumerate(row, 1):
                cell.border = border
                cell.font = data_font
                cell.alignment = left_align if col_idx <= 7 else center_align
        
        # Set column widths
        widths = {'A': 15, 'B': 20, 'C': 18, 'D': 20, 'E': 18, 'F': 20, 'G': 18, 'H': 16, 'I': 16, 'J': 16, 'K': 20}
        for col_letter, width in widths.items():
            if ord(col_letter) - 64 <= len(df.columns):
                ws.column_dimensions[col_letter].width = width
        
        # Set row heights
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 20
        ws.row_dimensions[4].height = 30
        
        # Save
        wb.save(filename)
        logger.info(f"Report saved: {filename}")
        
        # Summary
        logger.info("\nFinal Results Summary:")
        logger.info("="*50)
        for result in results:
            tower_name = result['Tower']
            pct = result.get(f'% Work Done against Target-Till {CURRENT_TRACKER_MONTH}', '0%')
            target = result.get(f'Target {CURRENT_TRACKER_MONTH}', '')
            activity = result.get(f'Activity {CURRENT_TRACKER_MONTH}', '')
            # Find the original tower key for hardcoded check
            original_tower = None
            for tower in towers:
                if get_display_tower_name(tower) == tower_name:
                    original_tower = tower
                    break
            hardcoded_note = " (HARDCODED)" if original_tower and original_tower in HARDCODED_PERCENTAGES else ""
            logger.info(f"{tower_name}: {pct}{hardcoded_note}")
            logger.info(f"  Target: {target}")
            logger.info(f"  Activity: {activity}")
            logger.info("")
            
    except Exception as e:
        logger.error(f"Error: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    main()
