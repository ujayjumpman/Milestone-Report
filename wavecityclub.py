import os
import re
import logging
from io import BytesIO
from datetime import datetime

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from dotenv import load_dotenv
import ibm_boto3
from ibm_botocore.client import Config

# -----------------------------------------------------------------------------
# CONFIG / CONSTANTS
# -----------------------------------------------------------------------------
load_dotenv()
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)

# Validate required environment variables
required = {
    'COS_API_KEY': os.getenv('COS_API_KEY'),
    'COS_SERVICE_INSTANCE_CRN': os.getenv('COS_SERVICE_INSTANCE_CRN'),
    'COS_ENDPOINT': os.getenv('COS_ENDPOINT'),
    'COS_BUCKET_NAME': os.getenv('COS_BUCKET_NAME'),
    'KRA_FILE_PATH': os.getenv('KRA_FILE_PATH'),
    'WCC_TRACKER_PATH': os.getenv('WCC_TRACKER_PATH'),
}
missing = [k for k, v in required.items() if not v]
if missing:
    logger.error(f"Missing required environment variables: {', '.join(missing)}")
    raise SystemExit(1)

COS_API_KEY     = required['COS_API_KEY']
COS_CRN         = required['COS_SERVICE_INSTANCE_CRN']
COS_ENDPOINT    = required['COS_ENDPOINT']
BUCKET          = required['COS_BUCKET_NAME']
WCC_KRA_KEY     = required['KRA_FILE_PATH']
WCC_TRACKER_KEY = required['WCC_TRACKER_PATH']

# Block mapping from KRA to tracker sheets (exact mapping as specified)
BLOCK_MAPPING = {
    'Block 1 (B1) Banquet Hall': 'B1 Banket Hall & Finedine ',  # Note the trailing space
    'Fine Dine': 'B1 Banket Hall & Finedine ',  # Note the trailing space
    'Block 5 (B5) Admin + Member Lounge+Creche+Av Room + Surveillance Room +Toilets': 'B5',
    'Block 6 (B6) Toilets': 'B6',
    'Block 7(B7) Indoor Sports': 'B7',
    'Block 9 (B9) Spa & Saloon': 'B9',
    'Block 8 (B8) Squash Court': 'B8',
    'Block 2 & 3 (B2 & B3) Cafe & Bar': 'B2 & B3',
    'Block 4 (B4) Indoor Swimming Pool Changing Room & Toilets': 'B4',
    'Block 11 (B11) Guest House': 'B11',
    'Block 10 (B10) Gym': 'B10'
}

# Special handling for blocks that need enhanced search within specific sheets
SPECIAL_BLOCKS_ENHANCED_SEARCH = {
    'Block 1 (B1) Banquet Hall': 'B1 Banket Hall & Finedine ',  # Note the trailing space
    'Fine Dine': 'B1 Banket Hall & Finedine '  # Note the trailing space
}

# -----------------------------------------------------------------------------
# COS HELPERS
# -----------------------------------------------------------------------------

def init_cos():
    return ibm_boto3.client(
        's3',
        ibm_api_key_id=COS_API_KEY,
        ibm_service_instance_id=COS_CRN,
        config=Config(signature_version='oauth'),
        endpoint_url=COS_ENDPOINT,
    )

def download_file_bytes(cos, key):
    obj = cos.get_object(Bucket=BUCKET, Key=key)
    return obj['Body'].read()

def find_latest_wcc_tracker_key(cos):
    """List objects under the Wave City Club prefix and return the newest tracker file key."""
    prefix = 'Wave City Club/'
    resp = cos.list_objects_v2(Bucket=BUCKET, Prefix=prefix)
    contents = resp.get('Contents', [])
    candidates = [o for o in contents if 'Structure Work Tracker' in o['Key']]
    if not candidates:
        raise RuntimeError(f"No tracker files found under prefix {prefix!r}")
    latest = max(candidates, key=lambda o: o['LastModified'])
    key = latest['Key']
    logger.info(f"Auto-selected tracker key: {key}")
    return key

# -----------------------------------------------------------------------------
# UTILITIES
# -----------------------------------------------------------------------------

def extract_percentage(cell_value):
    """Extract percentage value from cell, handling different formats"""
    if not cell_value or cell_value == '-':
        return 0.0
    
    # Handle numeric values
    if isinstance(cell_value, (int, float)):
        if cell_value <= 1.0:
            return cell_value * 100  # Convert decimal to percentage
        return cell_value
    
    # Handle string values
    val_str = str(cell_value).replace('%', '').strip()
    try:
        val = float(val_str)
        if val <= 1.0:
            return val * 100  # Convert decimal to percentage
        return val
    except ValueError:
        # Try to extract numbers from strings
        numbers = re.findall(r'\d+\.?\d*', val_str)
        if numbers:
            val = float(numbers[0])
            return val if val > 1.0 else val * 100
        return 0.0

def normalize_activity_name(activity):
    """Normalize activity name for better matching"""
    if not activity:
        return ""
    return str(activity).strip().lower()

def activities_match(target_activity, tracker_activity):
    """Enhanced matching with case-insensitive comparison and better logging"""
    if not target_activity or not tracker_activity:
        return False
    
    # Clean both activities (remove extra spaces, convert to string)
    target = str(target_activity).strip()
    tracker = str(tracker_activity).strip()
    
    # Try exact match first
    if target == tracker:
        return True
    
    # Try case-insensitive match
    if target.lower() == tracker.lower():
        logger.info(f"CASE-INSENSITIVE MATCH: '{target}' matches '{tracker}'")
        return True
    
    # Log the mismatch for debugging
    logger.debug(f"NO MATCH: Target='{target}' (len={len(target)}) vs Tracker='{tracker}' (len={len(tracker)})")
    return False

# -----------------------------------------------------------------------------
# DATA EXTRACTION FUNCTIONS
# -----------------------------------------------------------------------------

def get_wcc_targets_from_kra(cos):
    """Extract targets from KRA file - B1=June, C1=July, D1=August with detailed logging"""
    raw = download_file_bytes(cos, WCC_KRA_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    sheet = wb['Wave City Club targets till Aug']
    
    targets = {}
    logger.info("=== DEBUG: Extracting targets from KRA file ===")
    
    # Read targets from the KRA file
    for row_num in range(2, sheet.max_row + 1):
        block_cell = sheet[f'A{row_num}']
        june_cell = sheet[f'B{row_num}']
        july_cell = sheet[f'C{row_num}']
        august_cell = sheet[f'D{row_num}']
        
        if block_cell.value:
            block_name = str(block_cell.value).strip()
            june_activity = str(june_cell.value or '').strip() if june_cell.value else ''
            july_activity = str(july_cell.value or '').strip() if july_cell.value else ''
            august_activity = str(august_cell.value or '').strip() if august_cell.value else ''
            
            targets[block_name] = {
                'June': june_activity,
                'July': july_activity,
                'August': august_activity
            }
            
            # Debug logging
            logger.info(f"Row {row_num}: Block='{block_name}', June='{june_activity}'")
    
    logger.info(f"Extracted targets for {len(targets)} blocks from KRA")
    logger.info("=== All extracted targets ===")
    for block, activities in targets.items():
        logger.info(f"Block: '{block}' -> June: '{activities['June']}'")
    
    return targets

def find_activity_progress_in_sheet(sheet, target_activity, sheet_name, block_name=None):
    """
    Enhanced function to handle special cases for Block 1 and Fine Dine
    For these blocks, perform enhanced search within the entire sheet
    Modified to return 100% when there are no target activities
    """
    logger.info(f"=== DEBUG: Looking for activity '{target_activity}' in sheet '{sheet_name}' for block '{block_name}' ===")
    
    # Check if there's no target activity - return 100% in these cases
    if not target_activity or target_activity.strip() == '' or target_activity.lower() in ['no target', 'no target for june', '-']:
        logger.info(f"No specific target activity found for {block_name}, returning 100% completion")
        return 100.0
    
    # Handle special cases for Block 1 and Fine Dine - enhanced search
    if block_name in SPECIAL_BLOCKS_ENHANCED_SEARCH:
        logger.info(f"=== SPECIAL CASE: {block_name} - performing enhanced search in entire sheet ===")
        logger.info(f"Target activity: '{target_activity}' (repr: {repr(target_activity)})")
        
        # Search through more rows for these special blocks
        max_rows_to_check = min(sheet.max_row, 60)  # Check more rows for special blocks
        found_activities = []
        
        for row_num in range(1, max_rows_to_check + 1):
            try:
                activity_cell = sheet[f'G{row_num}']
                if activity_cell.value:
                    tracker_activity = str(activity_cell.value).strip()
                    found_activities.append(f"G{row_num}: '{tracker_activity}'")
                    
                    # Check for match (now includes case-insensitive)
                    if activities_match(target_activity, tracker_activity):
                        # Found matching activity, get progress from AC column same row
                        progress_cell = sheet[f'AC{row_num}']
                        ac_value = progress_cell.value
                        logger.info(f"MATCH FOUND in G{row_num}: '{tracker_activity}'")
                        logger.info(f"Corresponding AC{row_num} value: {ac_value}")
                        
                        if ac_value is not None:
                            progress = extract_percentage(ac_value)
                            logger.info(f"Extracted progress for {block_name}: {progress}%")
                            return progress
                        else:
                            logger.warning(f"Found activity match in G{row_num} but AC{row_num} is empty")
                            return 0.0
                            
            except Exception as e:
                logger.debug(f"Error checking row {row_num}: {e}")
                continue
        
        # Log all found activities for debugging
        logger.warning(f"=== ALL ACTIVITIES FOUND in sheet '{sheet_name}' ===")
        for activity in found_activities:
            logger.warning(activity)
        
        logger.warning(f"NO MATCH found for {block_name} target: '{target_activity}' in enhanced search")
        
        # Additional debugging - show character codes
        logger.warning(f"Target activity character codes: {[ord(c) for c in target_activity]}")
        
        return 0.0
    
    # Original logic for other blocks
    # First, let's see what's actually in column G
    logger.info(f"=== Scanning column G in sheet '{sheet_name}' ===")
    activities_found = []
    max_rows_to_check = min(sheet.max_row, 20)  # Check first 20 rows for debugging
    
    for row_num in range(1, max_rows_to_check + 1):  # Start from row 1 to see headers too
        try:
            activity_cell = sheet[f'G{row_num}']
            if activity_cell.value:
                tracker_activity = str(activity_cell.value).strip()
                activities_found.append(f"G{row_num}: '{tracker_activity}'")
                logger.info(f"Found in G{row_num}: '{tracker_activity}'")
                
                # Check for EXACT match
                if activities_match(target_activity, tracker_activity):
                    # Found exact matching activity, get progress from AC column same row
                    progress_cell = sheet[f'AC{row_num}']
                    ac_value = progress_cell.value
                    logger.info(f"EXACT MATCH FOUND in G{row_num}: '{tracker_activity}'")
                    logger.info(f"Corresponding AC{row_num} value: {ac_value}")
                    
                    if ac_value is not None:
                        progress = extract_percentage(ac_value)
                        logger.info(f"Extracted progress: {progress}%")
                        return progress
                    else:
                        logger.warning(f"Found activity match in G{row_num} but AC{row_num} is empty")
                        
        except Exception as e:
            logger.debug(f"Error checking row {row_num}: {e}")
            continue
    
    # Log all activities found for debugging
    logger.info(f"=== All activities found in column G ===")
    for activity in activities_found[:10]:  # Show first 10
        logger.info(activity)
    
    logger.warning(f"NO EXACT MATCH found for target: '{target_activity}'")
    logger.warning(f"Target length: {len(target_activity)}, Target repr: {repr(target_activity)}")
    return 0.0

def get_wcc_progress_from_tracker(cos, targets, tracker_key):
    """Extract progress data from tracker file and match with targets - Display June data only"""
    raw = download_file_bytes(cos, tracker_key)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    logger.info(f"Available tracker sheets: {wb.sheetnames}")
    
    progress_data = []
    milestone_counter = 1
    
    for block_name, month_activities in targets.items():
        logger.info(f"Processing block: {block_name}")
        
        # Get the corresponding tracker sheet name
        sheet_name = BLOCK_MAPPING.get(block_name)
        
        if not sheet_name:
            logger.warning(f"No sheet mapping found for block: {block_name}")
            june_progress = 0.0
        elif sheet_name not in wb.sheetnames:
            logger.warning(f"Sheet '{sheet_name}' not found in tracker workbook")
            june_progress = 0.0
        else:
            # Get the sheet and find progress for June activity
            sheet = wb[sheet_name]
            june_activity = month_activities.get('June', '')
            # Pass block_name to the function to handle special cases
            june_progress = find_activity_progress_in_sheet(sheet, june_activity, sheet_name, block_name)
        
        # Create row data - Keep all columns but only populate June data
        row_data = {
            'Milestone': f"Milestone-{milestone_counter:02d}",
            'Block': block_name,
            'Activity June': month_activities.get('June', ''),
            'Activity July': month_activities.get('July', ''),
            'Activity August': month_activities.get('August', ''),
            '% Work Done against Target-Till June': f"{june_progress:.1f}%" if june_progress > 0 else "0.0%",
            '% Work Done against Target-Till July': "",  # Keep blank
            '% Work Done against Target-Till August': "",  # Keep blank
            'Delay Reasons June': "",
            'Delay Reasons July': "",
            'Delay Reasons August': ""
        }
        
        progress_data.append(row_data)
        milestone_counter += 1
        logger.info(f"Block {block_name} -> June Progress: {june_progress:.1f}%")
    
    # Create DataFrame with all columns (same format) - now includes Block column
    columns = [
        'Milestone',
        'Block',
        'Activity June', 'Activity July', 'Activity August',
        '% Work Done against Target-Till June',
        '% Work Done against Target-Till July', 
        '% Work Done against Target-Till August',
        'Delay Reasons June', 'Delay Reasons July', 'Delay Reasons August'
    ]
    
    df = pd.DataFrame(progress_data, columns=columns)
    logger.info(f"Created DataFrame with {len(df)} rows - June data populated, other months blank")
    return df

# -----------------------------------------------------------------------------
# EXCEL REPORT GENERATION
# -----------------------------------------------------------------------------

def write_wcc_excel_report(df, filename):
    """Generate formatted Excel report"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Wave City Club Progress'
    
    # Add title and date at the top
    current_date = datetime.now().strftime("%d-%m-%Y")
    ws.append(["Wave City Club Structure Work Progress"])
    ws.append([f"Report Generated on: {current_date}"])
    ws.append([])  # Empty row for spacing
    
    # Define styles
    yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    grey   = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
    bold   = Font(bold=True)
    norm   = Font(bold=False)
    title_font = Font(bold=True, size=14)
    date_font = Font(bold=False, size=10, color="666666")
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left   = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin   = Side(style='thin', color='000000')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    
    # Style title row
    ws.merge_cells(f'A1:{get_column_letter(len(df.columns))}1')
    ws['A1'].font = title_font
    ws['A1'].alignment = center
    ws['A1'].fill = grey
    
    # Style date row
    ws.merge_cells(f'A2:{get_column_letter(len(df.columns))}2')
    ws['A2'].font = date_font
    ws['A2'].alignment = center

    # Add section title
    ws.append(["Wave City Club Structure Work Progress Against Milestones"])
    title_row = ws.max_row
    ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=len(df.columns))
    
    for cell in ws[title_row]:
        cell.fill = grey
        cell.font = bold
        cell.alignment = center
        cell.border = border
    
    # Add DataFrame data
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    
    # Style header row
    header_row = title_row + 1
    for cell in ws[header_row]:
        cell.font = bold
        cell.alignment = center
        cell.border = border
    
    # Style data rows
    body_start = header_row + 1
    body_end = ws.max_row
    
    for r in range(body_start, body_end + 1):
        for cell in ws[r]:
            cell.font = norm
            # Left align text columns, center align percentage columns
            if cell.column in [1, 2, 3, 4, 5, 9, 10, 11]:  # Milestone, Block, and activity columns, delay reason columns
                cell.alignment = left
            else:  # Percentage columns
                cell.alignment = center
            cell.border = border
    
    # Adjust column widths - Updated for Block column
    column_widths = {
        1: 15,  # Milestone
        2: 35,  # Block (wider for long block names)
        3: 25,  # Activity June
        4: 25,  # Activity July
        5: 25,  # Activity August
        6: 22,  # % Work Done June
        7: 22,  # % Work Done July
        8: 22,  # % Work Done August
        9: 20,  # Delay Reasons June
        10: 20, # Delay Reasons July
        11: 20  # Delay Reasons August
    }
    
    for col_num, width in column_widths.items():
        ws.column_dimensions[get_column_letter(col_num)].width = width
    
    # Set row heights
    for i in range(1, ws.max_row + 1):
        ws.row_dimensions[i].height = 22
    
    wb.save(filename)
    logger.info(f'Report saved to {filename}')

# -----------------------------------------------------------------------------
# MAIN FUNCTION
# -----------------------------------------------------------------------------

def main():
    """Main execution function"""
    try:
        # Initialize COS client
        cos = init_cos()
        
        # Get targets from KRA file
        logger.info("Fetching Wave City Club targets from KRA file...")
        targets = get_wcc_targets_from_kra(cos)
        
        # Determine tracker file to use
        logger.info("Determining tracker file to use...")
        try:
            cos.head_object(Bucket=BUCKET, Key=WCC_TRACKER_KEY)
            tracker_key = WCC_TRACKER_KEY
            logger.info(f"Using configured tracker key: {tracker_key}")
        except Exception:
            tracker_key = find_latest_wcc_tracker_key(cos)
        
        # Extract progress data
        logger.info("Extracting progress data from tracker...")
        df = get_wcc_progress_from_tracker(cos, targets, tracker_key)
        
        # Generate report
        filename = f"Wave_City_Club_Milestone_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        logger.info(f"Generating Excel report with June data (other months blank)")
        write_wcc_excel_report(df, filename)
        
        logger.info("Report generation completed successfully! June data populated, July/August blank.")
        
    except Exception as e:
        logger.error(f"Error in main execution: {e}")
        raise

if __name__ == "__main__":
    main()
