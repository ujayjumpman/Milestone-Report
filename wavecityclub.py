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

MONTHS = ['June', 'July', 'August']

# Block mapping from KRA to tracker sheets
BLOCK_MAPPING = {
    'Block 1 (B1) Banquet Hall': 'B1 Banket Hall & Finedine',
    'Fine Dine': 'B1 Banket Hall & Finedine',
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

def extract_number(cell_value):
    if not cell_value or cell_value == '-':
        return 0.0
    match = re.search(r"(\d+)", str(cell_value))
    return float(match.group(1)) if match else 0.0

def extract_percentage(cell_value):
    if not cell_value:
        return 0.0
    
    # Handle different data types
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
        # Try to extract numbers from strings like "75% complete"
        import re
        numbers = re.findall(r'\d+\.?\d*', val_str)
        if numbers:
            val = float(numbers[0])
            return val if val > 1.0 else val * 100
        return 0.0

def get_previous_months():
    now = datetime.now()
    month_map = {'June': 6, 'July': 7, 'August': 8}
    return [m for m in MONTHS if month_map[m] < now.month]

# -----------------------------------------------------------------------------
# WAVE CITY CLUB DATA EXTRACTION
# -----------------------------------------------------------------------------

def get_wcc_targets_from_kra(cos):
    raw = download_file_bytes(cos, WCC_KRA_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    sheet = wb['Wave City Club targets till Aug']
    
    targets = {}
    # Read all blocks from the KRA file - don't skip any
    for row_num in range(2, sheet.max_row + 1):
        block_cell = sheet[f'A{row_num}']
        june_cell = sheet[f'B{row_num}']
        july_cell = sheet[f'C{row_num}']
        august_cell = sheet[f'D{row_num}']
        
        if block_cell.value:
            block_name = str(block_cell.value).strip()
            targets[block_name] = {
                'June': str(june_cell.value or '').strip(),
                'July': str(july_cell.value or '').strip(), 
                'August': str(august_cell.value or '').strip()
            }
    
    logger.info(f"Wave City Club targets extracted: {targets}")
    return targets

def get_progress_from_tracker_sheet(sheet, activity_name, block_name):
    """Extract progress percentage from AC column for a specific activity"""
    logger.info(f"Looking for progress in sheet for block: {block_name}, activity: {activity_name}")
    
    # Special handling for B1 Banket Hall & Finedine sheet
    if "B1" in block_name or "Banquet Hall" in block_name or "Fine Dine" in block_name:
        logger.info(f"Special handling for B1 block: {block_name}")
        
        # First check AC2 directly
        ac2_cell = sheet['AC2']
        if ac2_cell.value is not None:
            progress = extract_percentage(ac2_cell.value)
            logger.info(f"Found AC2 value for {block_name}: {progress}%")
            if progress > 0:
                return progress
        
        # Check multiple AC cells for B1 sheet
        for row_num in range(1, 20):  # Check first 20 rows
            try:
                progress_cell = sheet[f'AC{row_num}']
                if progress_cell.value is not None:
                    progress = extract_percentage(progress_cell.value)
                    if progress > 0:
                        logger.info(f"Found progress in AC{row_num} for {block_name}: {progress}%")
                        return progress
            except Exception:
                continue
    
    # For blocks with "No target for June", check AC2 directly first
    if "No target for June" in activity_name:
        ac2_cell = sheet['AC2']
        if ac2_cell.value is not None:
            progress = extract_percentage(ac2_cell.value)
            logger.info(f"Found AC2 value for {block_name}: {progress}%")
            return progress
    
    # Look through rows to find matching activity
    for row_num in range(1, min(sheet.max_row + 1, 100)):  # Check more rows
        try:
            # Check multiple columns for activity names (F, G, H, etc.)
            activity_cells = [sheet[f'{col}{row_num}'] for col in ['F', 'G', 'H', 'A', 'B', 'C', 'D']]
            
            for activity_cell in activity_cells:
                if activity_cell.value:
                    activity_val = str(activity_cell.value).strip().lower()
                    activity_name_lower = activity_name.lower()
                    
                    # More flexible matching
                    if (activity_name_lower in activity_val or 
                        activity_val in activity_name_lower or
                        any(word in activity_val for word in activity_name_lower.split() if len(word) > 2) or
                        any(keyword in activity_val for keyword in ['roof', 'slab', 'casting', 'foundation', 'brick', 'plaster', 'ff', 'gf'])):
                        
                        # Check AC column for this row
                        progress_cell = sheet[f'AC{row_num}']
                        if progress_cell.value is not None:
                            progress = extract_percentage(progress_cell.value)
                            logger.info(f"Found matching activity '{activity_val}' in row {row_num} with progress: {progress}%")
                            if progress > 0:
                                return progress
        except Exception as e:
            continue
    
    # If no specific activity found, try to get any progress from AC column
    for row_num in range(1, min(sheet.max_row + 1, 50)):
        try:
            progress_cell = sheet[f'AC{row_num}']
            if progress_cell.value is not None:
                progress = extract_percentage(progress_cell.value)
                if progress > 0:
                    logger.info(f"Found general progress in AC{row_num}: {progress}% for {block_name}")
                    return progress
        except Exception:
            continue
    
    logger.warning(f"No progress found for {block_name}")
    return 0.0

def get_wcc_progress_from_tracker(cos, targets, tracker_key):
    raw = download_file_bytes(cos, tracker_key)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    logger.info(f"Available tracker sheets: {wb.sheetnames}")
    
    progress_data = []
    prev_months = get_previous_months()
    milestone_counter = 1
    
    # Process ALL blocks from targets, don't skip any
    for block_name, month_activities in targets.items():
        logger.info(f"Processing block: {block_name}")
        
        # Try to find corresponding tracker sheet
        sheet_name = BLOCK_MAPPING.get(block_name)
        progress_pct = 0.0
        
        if sheet_name:
            try:
                sheet = wb[sheet_name]
                # Get the main activity (usually from June column)
                main_activity = month_activities.get('June', '')
                if main_activity:
                    progress_pct = get_progress_from_tracker_sheet(sheet, main_activity, block_name)
                else:
                    # If no June activity, try AC2 directly
                    progress_pct = extract_percentage(sheet['AC2'].value)
                
                logger.info(f"Block {block_name} -> Sheet {sheet_name} -> Progress: {progress_pct}%")
            except KeyError:
                logger.warning(f"Sheet '{sheet_name}' not found in tracker")
                progress_pct = 0.0
        else:
            logger.warning(f"No mapping found for block: {block_name}")
            progress_pct = 0.0
        
        # Create row data with correct format
        row = {
            'Milestone': f"Milestone-{milestone_counter:02d}",
            'Activity June': month_activities.get('June', 'No target'),
            'Activity July': month_activities.get('July', 'No target'),
            'Activity August': month_activities.get('August', 'No target'),
            'Delay Reasons June': '',
            'Delay Reasons July': '',
            'Delay Reasons August': '',
        }
        
        # Add progress columns for June only (since tracker is for June)
        if "June" in prev_months:
            row[f"% Work Done against Target-Till June"] = f"{progress_pct}%"
        else:
            row[f"% Work Done against Target-Till June"] = ""
        
        # Leave July and August empty since tracker is only for June
        row[f"% Work Done against Target-Till July"] = ""
        row[f"% Work Done against Target-Till August"] = ""
        
        progress_data.append(row)
        milestone_counter += 1
    
    # Create DataFrame with correct column order
    cols = [
        'Milestone',
        'Activity June', 'Activity July', 'Activity August',
        '% Work Done against Target-Till June',
        '% Work Done against Target-Till July', 
        '% Work Done against Target-Till August',
        'Delay Reasons June', 'Delay Reasons July', 'Delay Reasons August'
    ]
    
    return pd.DataFrame(progress_data, columns=cols)

# -----------------------------------------------------------------------------
# EXCEL WRITER / STYLING - UPDATED WITH DATE DISPLAY
# -----------------------------------------------------------------------------

def write_wcc_excel_report(df, filename):
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
    
    # Style title row (row 1)
    ws.merge_cells(f'A1:{get_column_letter(len(df.columns))}1')
    ws['A1'].font = title_font
    ws['A1'].alignment = center
    ws['A1'].fill = grey
    
    # Style date row (row 2)
    ws.merge_cells(f'A2:{get_column_letter(len(df.columns))}2')
    ws['A2'].font = date_font
    ws['A2'].alignment = center

    def append_block(title, df_block, total_label):
        start, end = 1, len(df_block.columns)
        
        # Title row for the section
        ws.append([title])
        title_row = ws.max_row
        ws.merge_cells(start_row=title_row, start_column=start, end_row=title_row, end_column=end)
        for cell in ws[title_row]:
            cell.fill = grey
            cell.font = bold
            cell.alignment = center
            cell.border = border
        
        # DataFrame rows
        for row in dataframe_to_rows(df_block, index=False, header=True):
            ws.append(row)
        
        header_row = title_row + 1
        body_start = header_row + 1
        body_end = ws.max_row
        
        # Header styling
        for cell in ws[header_row]:
            cell.font = bold
            cell.alignment = center
            cell.border = border
        
        # Body styling
        for r in range(body_start, body_end + 1):
            for cell in ws[r]:
                cell.font = norm
                cell.alignment = left if cell.col_idx in (1, 2, 3, 4) else center
                cell.border = border
        
        # No total delay row needed since we removed weighted delay columns

    # Write the report (starting after the title, date, and empty row)
    append_block('Wave City Club Structure Work Progress Against Milestones', df, '')
    
    # Adjust column widths
    for col in ws.columns:
        max_length = 0
        for cell in col:
            text = str(cell.value or '')
            max_length = max(max_length, len(text.split('\n')[0]))
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_length + 4, 60)
    
    # Set row heights
    for i in range(1, ws.max_row + 1):
        ws.row_dimensions[i].height = 22
    
    wb.save(filename)
    logger.info(f'Report saved to {filename}')

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------

def main():
    cos = init_cos()
    
    logger.info("Fetching Wave City Club targets from KRA...")
    targets = get_wcc_targets_from_kra(cos)
    
    logger.info("Extracting progress data from tracker...")
    try:
        cos.head_object(Bucket=BUCKET, Key=WCC_TRACKER_KEY)
        tracker_key = WCC_TRACKER_KEY
    except Exception:
        tracker_key = find_latest_wcc_tracker_key(cos)
    
    df = get_wcc_progress_from_tracker(cos, targets, tracker_key)
    
    filename = f"Wave_City_Club_Milestone_Report ({datetime.now():%Y-%m-%d}).xlsx"
    logger.info("Writing Excel report...")
    write_wcc_excel_report(df, filename)
    logger.info("Completed!")

if __name__ == "__main__":
    main()
