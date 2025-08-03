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

# Parent activity rows (these contain the parent activity names)
KRA_PARENT_ROW = {
    "Tower 4": [5, 6],    # B5, B6 (Upper basement, beam/slab etc.)
    "Tower 5": [8, 9],    # B8, B9 
    "Tower 6": [11, 12],  # B11, B12
    "Tower 7": [14, 15],  # B14, B15
    "NTA-01": [17, 18],   # B17, B18
    "NTA-02": [20, 21],   # B20, B21
}

# Child activity rows (these contain the specific tasks to be tracked)
KRA_ACTIVITY_ROW = {
    "Tower 4": 7,     # B7 - Child activity for Tower 4
    "Tower 5": 10,    # B10 - Child activity for Tower 5  
    "Tower 6": 13,    # B13 - Child activity for Tower 6
    "Tower 7": 16,    # B16 - Child activity for Tower 7
    "NTA-01": 19,     # B19 - Child activity for NTA-01
    "NTA-02": 22,     # B22 - Child activity for NTA-02
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

def get_activity_for_month(tower, month, kra_ws):
    """Get the activity name for a specific tower and month from KRA file"""
    month_col = MONTH_COLUMNS[month]
    child_row = KRA_ACTIVITY_ROW[tower]
    child_name = kra_ws.cell(row=child_row, column=month_col).value
    
    if child_name and str(child_name).strip():
        return str(child_name).strip()
    return ""

# ============= MAIN LOGIC ==============
def find_child_activity_pct(tracker_ws, parent_names, child_name):
    """
    Find the parent row(s) (bold text), then find the child row under them (non-bold), and return % Complete from column G.
    parent_names can be a single string or list of strings to check for multiple parent activities.
    """
    TASK_NAME_COL = 4  # D column (Task Name)
    PCT_COL = 7        # G column (% Complete)
    max_row = tracker_ws.max_row
    
    if isinstance(parent_names, str):
        parent_names = [parent_names]
    
    # Remove None values and clean parent names
    parent_names = [str(p).strip().lower() for p in parent_names if p is not None and str(p).strip()]
    
    if not parent_names:
        logger.warning(f"No valid parent names provided for child: {child_name}")
        return 0.0
    
    logger.info(f"Looking for child '{child_name}' under parents: {parent_names}")
    logger.info(f"Searching in tracker sheet with max_row: {max_row}")
    
    child_name_clean = str(child_name).strip().lower() if child_name else ""
    if not child_name_clean:
        logger.warning("Child name is empty or None")
        return 0.0
    
    # Show what we're looking for
    logger.info(f"Child activity to find: '{child_name_clean}'")
    
    # Find parent rows (bold text)
    parent_rows = []
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
            
            if is_bold and any(parent.lower().strip() == cell_val_clean or 
                              cell_val_clean in parent.lower().strip() for parent in parent_names):
                parent_rows.append(row)
                logger.info(f"Found BOLD parent '{cell_val}' at row {row}")
    
    if not parent_rows:
        logger.warning(f"No BOLD parent activities found for: {parent_names}")
        # Try searching for non-bold parents as fallback
        for row in range(2, max_row + 1):
            cell_val = tracker_ws.cell(row=row, column=TASK_NAME_COL).value
            if cell_val:
                cell_val_clean = str(cell_val).strip().lower()
                if any(parent.lower() in cell_val_clean or cell_val_clean in parent.lower() for parent in parent_names):
                    parent_rows.append(row)
                    logger.info(f"Found non-bold parent '{cell_val}' at row {row} (fallback)")
    
    if not parent_rows:
        logger.warning(f"No parent activities found at all for: {parent_names}")
        return 0.0
    
    # Search for child under each parent (non-bold activities)
    for parent_row in parent_rows:
        logger.info(f"Searching for NON-BOLD child activities under parent at row {parent_row}")
        
        # Look for child activities under this parent
        for row in range(parent_row + 1, min(parent_row + 50, max_row + 1)):  # Search up to 50 rows below parent
            task_val = tracker_ws.cell(row=row, column=TASK_NAME_COL).value
            
            # Check if we've reached end of this parent's section
            if task_val is None or str(task_val).strip() == "":
                logger.debug(f"Blank row at {row}, continuing search")
                continue
                
            # Check if this row is bold (next parent) - if so, stop searching
            try:
                font = tracker_ws.cell(row=row, column=TASK_NAME_COL).font
                is_bold = font and font.bold
                if is_bold:
                    logger.debug(f"Found next BOLD parent at row {row}, stopping search under current parent")
                    break
            except:
                is_bold = False
            
            # This is a non-bold row (potential child activity)
            if not is_bold:
                task_val_clean = str(task_val).strip().lower()
                logger.debug(f"Row {row} (NON-BOLD): Checking '{task_val_clean}' against '{child_name_clean}'")
                
                # Check if this matches the child activity using keyword-based matching
                # Include both parent and child activity keywords for better matching
                match_found = False
                
                # Extract keywords from child activity name (more flexible matching)
                # Remove common words and focus on key terms
                common_words = ['work', 'activity', 'and', 'the', 'of', 'for', 'in', 'on', 'with', '100%', '-']
                
                # Clean and extract keywords from child name
                child_keywords = []
                child_parts = child_name_clean.replace('-', ' ').replace('%', '').split()
                for part in child_parts:
                    part_clean = part.strip().lower()
                    if len(part_clean) > 2 and part_clean not in common_words:
                        child_keywords.append(part_clean)
                
                # Also include parent activity keywords for comprehensive matching
                parent_keywords = []
                for parent_name in parent_names:
                    parent_parts = parent_name.lower().replace('-', ' ').replace('/', ' ').split()
                    for part in parent_parts:
                        part_clean = part.strip().lower()
                        if len(part_clean) > 2 and part_clean not in common_words and part_clean not in parent_keywords:
                            parent_keywords.append(part_clean)
                
                # Combine all keywords for matching
                all_search_keywords = child_keywords + parent_keywords
                
                # Clean task name
                task_keywords = []
                task_parts = task_val_clean.replace('/', ' ').replace('-', ' ').split()
                for part in task_parts:
                    part_clean = part.strip().lower()
                    if len(part_clean) > 2 and part_clean not in common_words:
                        task_keywords.append(part_clean)
                
                logger.debug(f"Child keywords: {child_keywords}")
                logger.debug(f"Parent keywords: {parent_keywords}")
                logger.debug(f"All search keywords: {all_search_keywords}")
                logger.debug(f"Task keywords: {task_keywords}")
                
                # Method 1: Keyword matching with parent + child terms
                if all_search_keywords and task_keywords:
                    matching_keywords = 0
                    matched_terms = []
                    
                    for search_keyword in all_search_keywords:
                        for task_keyword in task_keywords:
                            # Check for partial matches (substring)
                            if (search_keyword in task_keyword or 
                                task_keyword in search_keyword or
                                # Handle common variations
                                (search_keyword == 'reinforcement' and 'reinforc' in task_keyword) or
                                (search_keyword == 'casting' and 'cast' in task_keyword) or
                                (search_keyword == 'shuttering' and 'shutter' in task_keyword) or
                                (search_keyword == 'checking' and 'check' in task_keyword) or
                                (search_keyword == 'basement' and 'basement' in task_keyword) or
                                (search_keyword == 'binding' and 'bind' in task_keyword) or
                                (search_keyword == 'column' and 'col' in task_keyword) or
                                (search_keyword == 'shear' and 'shear' in task_keyword) or
                                (search_keyword == 'wall' and 'wall' in task_keyword) or
                                (search_keyword == 'upper' and 'upper' in task_keyword) or
                                (search_keyword == 'lower' and 'lower' in task_keyword)):
                                matching_keywords += 1
                                matched_terms.append(f"'{search_keyword}' ↔ '{task_keyword}'")
                                logger.debug(f"Keyword match: '{search_keyword}' matches '{task_keyword}'")
                                break
                    
                    # If at least 1 significant keyword matches, consider it a match
                    if matching_keywords >= 1:
                        match_found = True
                        logger.debug(f"Keyword-based match found ({matching_keywords} keywords matched): {', '.join(matched_terms)}")
                
                # Method 2: Fallback - simple substring match for safety
                if not match_found:
                    if child_name_clean in task_val_clean or task_val_clean in child_name_clean:
                        match_found = True
                        logger.debug(f"Fallback substring match found")
                
                if match_found:
                    # Get percentage from column G
                    pct_val = tracker_ws.cell(row=row, column=PCT_COL).value
                    logger.info(f"Found NON-BOLD child match at row {row}: '{task_val}' with percentage: {pct_val}")
                    
                    if pct_val is not None:
                        try:
                            # Handle different percentage formats
                            if isinstance(pct_val, (int, float)):
                                # If it's between 0 and 1, assume it's decimal format (0.8 = 80%)
                                if 0 <= pct_val <= 1:
                                    result = float(pct_val * 100)
                                else:
                                    result = float(pct_val)
                            else:
                                # Handle string percentages
                                pct_str = str(pct_val).replace("%", "").strip()
                                result = float(pct_str)
                            
                            logger.info(f"Successfully found and converted percentage: {pct_val} -> {result}%")
                            return result
                            
                        except Exception as e:
                            logger.warning(f"Error parsing percentage '{pct_val}': {e}")
                            continue
                    else:
                        logger.warning(f"No percentage value found in column G for matched row {row}")
    
    logger.warning(f"Child activity '{child_name}' not found under any parent")
    return 0.0

def get_parent_activities_for_month(tower, month, kra_ws):
    """Get the parent activity names for a specific tower and month from KRA file"""
    month_col = MONTH_COLUMNS[month]
    parent_rows = KRA_PARENT_ROW[tower]
    parent_names = []
    
    for parent_row in parent_rows:
        parent_name = kra_ws.cell(row=parent_row, column=month_col).value
        if parent_name and str(parent_name).strip():
            parent_names.append(str(parent_name).strip())
    
    # Join multiple parent names with " & "
    return " & ".join(parent_names) if parent_names else ""

def calculate_percentage_for_current_month(tower, kra_ws, tracker_wb):
    """Calculate percentage only for the current tracker month"""
    if CURRENT_TRACKER_MONTH not in MONTH_COLUMNS:
        return 0.0
        
    month_col = MONTH_COLUMNS[CURRENT_TRACKER_MONTH]
    
    # Get parent activity names from multiple rows
    parent_rows = KRA_PARENT_ROW[tower]
    parent_names = []
    for parent_row in parent_rows:
        parent_name = kra_ws.cell(row=parent_row, column=month_col).value
        if parent_name and str(parent_name).strip():
            parent_names.append(str(parent_name).strip())
    
    # Get child activity name
    child_row = KRA_ACTIVITY_ROW[tower]
    child_name = kra_ws.cell(row=child_row, column=month_col).value
    
    if not parent_names or not child_name or str(child_name).strip() == "":
        logger.warning(f"Missing parent activities or child activity for {tower} in {CURRENT_TRACKER_MONTH}")
        return 0.0
    
    child_name = str(child_name).strip()
    logger.info(f"{tower} ({CURRENT_TRACKER_MONTH}): Parent activities: {parent_names}, Child activity: {child_name}")
    
    # Get corresponding tracker sheet
    tracker_sheetname = TOWER_SHEET_MAP[tower]
    if tracker_sheetname not in tracker_wb.sheetnames:
        logger.warning(f"Sheet '{tracker_sheetname}' not found in tracker for {tower}")
        return 0.0
    
    tracker_ws = tracker_wb[tracker_sheetname]
    
    # Find the percentage completion
    pct = find_child_activity_pct(tracker_ws, parent_names, child_name)
    logger.info(f"{tower} ({CURRENT_TRACKER_MONTH}): {child_name} - {pct:.1f}% complete")
    
    return pct

def main():
    logger.info("Starting Eden KRA Milestone Report generation...")
    logger.info(f"Current tracker month: {CURRENT_TRACKER_MONTH}")
    
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
        
        logger.info("Processing data for all months...")
        
        results = []
        
        for tower in KRA_ACTIVITY_ROW:
            logger.info(f"Processing {tower}...")
            
            # Initialize row data with both Milestone and Tower columns
            row_data = {
                "Milestone": tower,
                "Tower": get_display_tower_name(tower)
            }
            
            # Get parent and child activity names for all months
            for month in MONTHS:
                # Get parent activities (called "Target")
                parent_activities = get_parent_activities_for_month(tower, month, kra_ws)
                row_data[f"Target {month}"] = parent_activities
                logger.info(f"{tower} - {month} Target: {parent_activities}")
                
                # Get child activities (called "Activity")
                activity_name = get_activity_for_month(tower, month, kra_ws)
                row_data[f"Activity {month}"] = activity_name
                logger.info(f"{tower} - {month} Activity: {activity_name}")
            
            # Calculate percentage only for current tracker month
            current_month_pct = calculate_percentage_for_current_month(tower, kra_ws, tracker_wb)
            
            # Set percentages - only current month gets actual value, others are blank or 0
            for month in MONTHS:
                if month == CURRENT_TRACKER_MONTH:
                    row_data[f"% Work Done against Target-Till {month}"] = f"{current_month_pct:.1f}%"
                else:
                    row_data[f"% Work Done against Target-Till {month}"] = ""  # Leave blank for future months
            
            # Add delay reasons column (empty for manual filling)
            row_data[f"Delay Reasons {CURRENT_TRACKER_MONTH}"] = ""
            
            results.append(row_data)
        
        if not results:
            logger.error("No data found to generate report!")
            return
        
        # Create Excel output
        df = pd.DataFrame(results)
        
        # Reorder columns to match the desired format
        column_order = ["Milestone", "Tower"]  # Added Tower column after Milestone
        
        # Add target and activity columns
        for month in MONTHS:
            column_order.append(f"Target {month}")
            column_order.append(f"Activity {month}")
        
        # Add percentage columns  
        for month in MONTHS:
            column_order.append(f"% Work Done against Target-Till {month}")
            
        # Add delay reasons
        column_order.append(f"Delay Reasons {CURRENT_TRACKER_MONTH}")
        
        # Filter columns that exist and reorder
        existing_columns = [col for col in column_order if col in df.columns]
        df = df[existing_columns]
        
        filename = f"Eden_KRA_Milestone_Report ({datetime.now():%Y-%m-%d}).xlsx"
        
        # Create formatted Excel file
        wb = Workbook()
        ws = wb.active
        ws.title = "Eden KRA Milestone Progress"
        
        # Add title and date at the top
        current_date = datetime.now().strftime("%d-%m-%Y")
        ws.append(["Eden KRA Milestone Progress"])
        ws.append([f"Report Generated on: {current_date}"])
        ws.append([])  # Empty row for spacing
        
        # Add data
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Format the worksheet
        header_font = Font(bold=True, size=11, color="000000")  # Changed to black text
        title_font = Font(bold=True, size=14, color="000000")
        date_font = Font(bold=False, size=10, color="666666")
        data_font = Font(size=10)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'),
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Changed to yellow
        
        # Format title row (row 1)
        ws.merge_cells(f'A1:{get_column_letter(len(df.columns))}1')
        ws['A1'].font = title_font
        ws['A1'].alignment = center_align
        
        # Format date row (row 2)
        ws.merge_cells(f'A2:{get_column_letter(len(df.columns))}2')
        ws['A2'].font = date_font
        ws['A2'].alignment = center_align
        
        # Format headers (row 4, since we added title, date, and empty row)
        for cell in ws[4]:
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
            cell.fill = header_fill
        
        # Format data rows (starting from row 5)
        for row_idx, row in enumerate(ws.iter_rows(min_row=5, max_row=ws.max_row), 5):
            for col_idx, cell in enumerate(row, 1):
                cell.border = border
                cell.font = data_font
                
                # Left align milestone, tower, target and activity columns
                if col_idx <= 8:  # Milestone, Tower, Target and Activity columns
                    cell.alignment = left_align
                else:  # Percentage and delay reason columns
                    cell.alignment = center_align
        
        # Set column widths based on number of columns (updated for additional Tower column)
        num_cols = ws.max_column
        column_widths = {}
        
        if num_cols >= 11:  # Full format with all columns including Tower and Target/Activity
            column_widths = {
                'A': 12,  # Milestone
                'B': 15,  # Tower
                'C': 20,  # Target June
                'D': 18,  # Activity June
                'E': 20,  # Target July  
                'F': 18,  # Activity July
                'G': 20,  # Target August
                'H': 18,  # Activity August
                'I': 16,  # % Work Done June
                'J': 16,  # % Work Done July
                'K': 16,  # % Work Done August
                'L': 20,  # Delay Reasons
            }
        else:  # Adjust for fewer columns
            base_width = 15
            for i in range(1, num_cols + 1):
                col_letter = chr(64 + i)  # A, B, C, etc.
                column_widths[col_letter] = base_width + (3 if i == 1 else 0)
        
        for col_letter, width in column_widths.items():
            if col_letter in [chr(65 + i) for i in range(num_cols)]:  # Only set widths for existing columns
                ws.column_dimensions[col_letter].width = width
        
        # Set row height for header and title rows
        ws.row_dimensions[1].height = 30  # Title row
        ws.row_dimensions[2].height = 20  # Date row
        ws.row_dimensions[4].height = 30  # Header row
        
        # Save the file
        wb.save(filename)
        logger.info(f"Successfully saved Eden KRA milestone report to {filename}")
        
        # Log summary
        logger.info("Report Summary:")
        for result in results:
            milestone = result['Milestone']
            tower = result['Tower']
            activities = []
            for month in MONTHS:
                target_key = f'Target {month}'
                activity_key = f'Activity {month}'
                if target_key in result and activity_key in result:
                    target_activity = result[target_key] or "N/A"
                    child_activity = result[activity_key] or "N/A"
                    activities.append(f"{month}: {target_activity} → {child_activity}")
            
            current_pct = result.get(f'% Work Done against Target-Till {CURRENT_TRACKER_MONTH}', '0.0%')
            logger.info(f"  {milestone} ({tower}): {' | '.join(activities)} - Current Progress ({CURRENT_TRACKER_MONTH}): {current_pct}")
            
    except Exception as e:
        logger.error(f"Error generating report: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    main()
