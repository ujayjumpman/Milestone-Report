import os
import re
import logging
from io import BytesIO
from datetime import datetime
from dateutil.relativedelta import relativedelta

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

# Dynamic tracker path - will be set by get_latest_tracker_paths()
WCC_TRACKER_KEY = None

# Dynamic months and years - will be set based on tracker date
MONTHS = []
MONTH_YEARS = {}  # Maps month name to year
TRACKER_DATE = None
TARGET_END_MONTH = None  # The last month in our 3-month range
TARGET_END_YEAR = None

# Block mapping from KRA to tracker sheets (exact mapping as specified)
BLOCK_MAPPING = {
    'Block 1 (B1) Banquet Hall': 'B1 Banket Hall & Finedine ',
    'Fine Dine': 'B1 Banket Hall & Finedine ',
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
    'Block 1 (B1) Banquet Hall': 'B1 Banket Hall & Finedine ',
    'Fine Dine': 'B1 Banket Hall & Finedine '
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
    if not key:
        raise ValueError("File key cannot be None or empty")
    obj = cos.get_object(Bucket=BUCKET, Key=key)
    return obj['Body'].read()

def list_files_in_folder(cos, folder_prefix):
    """List all files in a specific folder (prefix) in the COS bucket"""
    try:
        response = cos.list_objects_v2(Bucket=BUCKET, Prefix=folder_prefix)
        files = []
        if 'Contents' in response:
            for obj in response['Contents']:
                if not obj['Key'].endswith('/'):
                    files.append(obj['Key'])
        return files
    except Exception as e:
        logger.error(f"Error listing files in folder {folder_prefix}: {e}")
        return []

def extract_date_from_filename(filename):
    """Extract date from filename in format (dd-mm-yyyy)"""
    pattern = r'\((\d{2}-\d{2}-\d{4})\)'
    match = re.search(pattern, filename)
    if match:
        date_str = match.group(1)
        try:
            return datetime.strptime(date_str, '%d-%m-%Y')
        except ValueError:
            logger.warning(f"Could not parse date {date_str} from filename {filename}")
            return None
    return None

def get_month_name(month_num):
    """Convert month number to month name"""
    months = {
        1: "January", 2: "February", 3: "March", 4: "April",
        5: "May", 6: "June", 7: "July", 8: "August", 
        9: "September", 10: "October", 11: "November", 12: "December"
    }
    return months.get(month_num, "Unknown")

def get_month_number(month_name):
    """Convert month name to month number"""
    months = {
        "January": 1, "February": 2, "March": 3, "April": 4,
        "May": 5, "June": 6, "July": 7, "August": 8, 
        "September": 9, "October": 10, "November": 11, "December": 12
    }
    return months.get(month_name, 1)

from datetime import datetime, timedelta
from typing import List, Tuple

def setup_dynamic_months(tracker_date: datetime) -> Tuple[List[str], str, datetime, List[Tuple[int, int]]]:
    """
    Dynamic setup for included months, including year-aware transitions.
    Rules:
    1. If tracker is in September → include June, July, August.
    2. Otherwise → include previous, current, next months.
    """
    global MONTHS, MONTH_YEARS, TARGET_END_MONTH, TARGET_END_YEAR, TRACKER_DATE

    months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ]

    tracker_month = tracker_date.month
    tracker_year = tracker_date.year

    # ✅ Rule 1: September tracker → June–July–August (same year)
    if tracker_month == 9:
        MONTHS_DATA = [(6, tracker_year), (7, tracker_year), (8, tracker_year)]

    # ✅ Rule 2: Otherwise previous–current–next (year-aware)
    else:
        prev_month_date = tracker_date.replace(day=1) - timedelta(days=1)
        next_month_year = tracker_year + (1 if tracker_month == 12 else 0)
        next_month_num = 1 if tracker_month == 12 else tracker_month + 1
        next_month_date = tracker_date.replace(year=next_month_year, month=next_month_num, day=1)

        MONTHS_DATA = [
            (prev_month_date.month, prev_month_date.year),
            (tracker_month, tracker_year),
            (next_month_date.month, next_month_date.year)
        ]

    # Build labeled month list and set globals
    MONTHS = [f"{months[m - 1]} {y}" for m, y in MONTHS_DATA]
    
    # Build MONTH_YEARS mapping
    MONTH_YEARS = {f"{months[m - 1]} {y}": y for m, y in MONTHS_DATA}
    
    TARGET_END_MONTH = months[MONTHS_DATA[-1][0] - 1]
    TARGET_END_YEAR = MONTHS_DATA[-1][1]
    TRACKER_DATE = tracker_date
    
    TARGET_MONTH = MONTHS[-1]

    return MONTHS, TARGET_MONTH, tracker_date, MONTHS_DATA

def get_latest_tracker_paths(cos):
    """Get the latest tracker file path for Wave City Club"""
    global WCC_TRACKER_KEY
    
    logger.info("=== FINDING LATEST WAVE CITY CLUB TRACKER FILES ===")
    
    wcc_files = list_files_in_folder(cos, "Wave City Club/")
    logger.info(f"Found {len(wcc_files)} files in Wave City Club folder")
    
    tracker_pattern = r'Structure Work Tracker.*\.xlsx$'
    
    matching_files = []
    latest_date = None
    
    for file_path in wcc_files:
        filename = os.path.basename(file_path)
        if re.search(tracker_pattern, filename, re.IGNORECASE):
            file_date = extract_date_from_filename(filename)
            if file_date:
                matching_files.append((file_path, file_date))
                logger.info(f"Found: {filename} with date {file_date.strftime('%d-%m-%Y')}")
                
                if latest_date is None or file_date > latest_date:
                    latest_date = file_date
            else:
                logger.warning(f"Found matching file but no date: {filename}")
    
    if matching_files:
        latest_file = max(matching_files, key=lambda x: x[1])
        WCC_TRACKER_KEY = latest_file[0]
        logger.info(f"✅ Latest Wave City Club tracker: {WCC_TRACKER_KEY} ({latest_file[1].strftime('%d-%m-%Y')})")
    else:
        logger.error(f"❌ No Wave City Club tracker files found!")
        WCC_TRACKER_KEY = None
    
    logger.info(f"\n=== FINAL WAVE CITY CLUB TRACKER PATH ===")
    logger.info(f"WCC_TRACKER_KEY: {WCC_TRACKER_KEY}")
    logger.info(f"Latest tracker date found: {latest_date.strftime('%d-%m-%Y') if latest_date else 'None'}")
    
    if latest_date:
        setup_dynamic_months(latest_date)
    else:
        logger.error("No valid tracker date found - using default setup")
        fallback_date = datetime.now()
        setup_dynamic_months(fallback_date)
    
    if not WCC_TRACKER_KEY:
        raise Exception("Could not find latest Wave City Club tracker file")
    
    return WCC_TRACKER_KEY

# -----------------------------------------------------------------------------
# UTILITIES
# -----------------------------------------------------------------------------

def extract_percentage(cell_value):
    """Extract percentage value from cell, handling different formats"""
    if not cell_value or cell_value == '-':
        return 0.0
    
    if isinstance(cell_value, (int, float)):
        if cell_value <= 1.0:
            return cell_value * 100
        return cell_value
    
    val_str = str(cell_value).replace('%', '').strip()
    try:
        val = float(val_str)
        if val <= 1.0:
            return val * 100
        return val
    except ValueError:
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
    """Enhanced matching with case-insensitive comparison"""
    if not target_activity or not tracker_activity:
        return False
    
    target = str(target_activity).strip()
    tracker = str(tracker_activity).strip()
    
    if target == tracker:
        return True
    
    if target.lower() == tracker.lower():
        logger.info(f"CASE-INSENSITIVE MATCH: '{target}' matches '{tracker}'")
        return True
    
    logger.debug(f"NO MATCH: Target='{target}' vs Tracker='{tracker}'")
    return False

# -----------------------------------------------------------------------------
# DATA EXTRACTION FUNCTIONS
# -----------------------------------------------------------------------------

def detect_kra_column_mapping(sheet):
    """Detect month-to-column mapping from KRA file headers"""
    month_to_col = {}
    
    for col_idx in range(2, 14):  # Columns B to M (2 to 13)
        col_letter = get_column_letter(col_idx)
        header_cell = sheet[f'{col_letter}1']
        
        if header_cell.value:
            header_text = str(header_cell.value).strip()
            logger.info(f"KRA Header in column {col_letter}: '{header_text}'")
            
            # Try to extract month name from header
            for month in ["January", "February", "March", "April", "May", "June", 
                         "July", "August", "September", "October", "November", "December"]:
                if month.lower() in header_text.lower():
                    month_to_col[month] = col_letter
                    logger.info(f"Mapped {month} -> Column {col_letter}")
                    break
    
    return month_to_col

def get_wcc_targets_from_kra(cos):
    """Extract targets from KRA file dynamically based on current months"""
    raw = download_file_bytes(cos, WCC_KRA_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    sheet = wb['Wave City Club targets till Aug']
    
    targets = {}
    logger.info("=== EXTRACTING TARGETS FROM KRA FILE ===")
    logger.info(f"Looking for months: {MONTHS} with years: {MONTH_YEARS}")
    
    # Detect column mapping from file
    month_to_col = detect_kra_column_mapping(sheet)
    
    if not month_to_col:
        logger.warning("Could not detect column mapping, using default")
        month_to_col = {
            "June": "B", "July": "C", "August": "D",
            "September": "E", "October": "F", "November": "G", "December": "H",
            "January": "I", "February": "J", "March": "K", "April": "L", "May": "M"
        }
    
    # Read targets from the KRA file
    for row_num in range(2, sheet.max_row + 1):
        block_cell = sheet[f'A{row_num}']
        
        if block_cell.value:
            block_name = str(block_cell.value).strip()
            month_activities = {}
            
            for month in MONTHS:
                col = month_to_col.get(month, "B")
                cell = sheet[f'{col}{row_num}']
                activity = str(cell.value or '').strip() if cell.value else ''
                month_activities[month] = activity
                logger.info(f"Row {row_num}, {month}: Block='{block_name}', Activity='{activity}'")
            
            targets[block_name] = month_activities
    
    logger.info(f"Extracted targets for {len(targets)} blocks")
    return targets

def find_activity_progress_in_sheet(sheet, target_activity, sheet_name, block_name=None):
    """Find activity progress with enhanced search for special blocks"""
    logger.info(f"=== Looking for '{target_activity}' in '{sheet_name}' for '{block_name}' ===")
    
    # Return 100% when there's no target activity
    if not target_activity or target_activity.strip() == '' or target_activity.lower() in ['no target', '-']:
        logger.info(f"No target activity, returning 100%")
        return 100.0
    
    # Enhanced search for special blocks
    if block_name in SPECIAL_BLOCKS_ENHANCED_SEARCH:
        logger.info(f"SPECIAL CASE: {block_name} - enhanced search")
        max_rows = min(sheet.max_row, 60)
        
        for row_num in range(1, max_rows + 1):
            try:
                activity_cell = sheet[f'G{row_num}']
                if activity_cell.value:
                    tracker_activity = str(activity_cell.value).strip()
                    
                    if activities_match(target_activity, tracker_activity):
                        progress_cell = sheet[f'AC{row_num}']
                        ac_value = progress_cell.value
                        logger.info(f"MATCH in G{row_num}: '{tracker_activity}', AC{row_num}: {ac_value}")
                        
                        if ac_value is not None:
                            return extract_percentage(ac_value)
                        return 0.0
            except Exception as e:
                logger.debug(f"Error at row {row_num}: {e}")
                continue
        
        logger.warning(f"NO MATCH for '{target_activity}'")
        return 0.0
    
    # Standard search for other blocks
    max_rows = min(sheet.max_row, 20)
    for row_num in range(1, max_rows + 1):
        try:
            activity_cell = sheet[f'G{row_num}']
            if activity_cell.value:
                tracker_activity = str(activity_cell.value).strip()
                
                if activities_match(target_activity, tracker_activity):
                    progress_cell = sheet[f'AC{row_num}']
                    ac_value = progress_cell.value
                    
                    if ac_value is not None:
                        return extract_percentage(ac_value)
                    return 0.0
        except Exception as e:
            logger.debug(f"Error at row {row_num}: {e}")
            continue
    
    logger.warning(f"NO MATCH for '{target_activity}'")
    return 0.0

def get_wcc_progress_from_tracker_all_months(cos, targets, tracker_key):
    """Extract progress data from tracker file with ALL months displayed"""
    raw = download_file_bytes(cos, tracker_key)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    logger.info(f"Available tracker sheets: {wb.sheetnames}")
    
    progress_data = []
    milestone_counter = 1
    total_blocks = len(targets)
    site_weighted = round(100 / total_blocks, 2) if total_blocks > 0 else 0
    
    for block_name, month_activities in targets.items():
        logger.info(f"Processing block: {block_name}")
        
        sheet_name = BLOCK_MAPPING.get(block_name)
        month_progress = {month: 0.0 for month in MONTHS}
        
        if not sheet_name:
            logger.warning(f"No sheet mapping for block: {block_name}")
        elif sheet_name not in wb.sheetnames:
            logger.warning(f"Sheet '{sheet_name}' not found")
        else:
            sheet = wb[sheet_name]
            for month in MONTHS:
                if month in month_activities:
                    activity = month_activities[month]
                    month_progress[month] = find_activity_progress_in_sheet(
                        sheet, activity, sheet_name, block_name
                    )
        
        # Use the last month for weighted calculation
        last_month = MONTHS[-1] if MONTHS else ''
        main_weighted = round((site_weighted * month_progress[last_month]) / 100, 3)
        
        # Create row data with dynamic columns
        row_data = {
            'Milestone': f"Milestone-{milestone_counter:02d}",
            'Block': block_name,
            f'Target to be complete by {TARGET_END_MONTH}-{TARGET_END_YEAR}': month_activities.get(last_month, ''),
            'Site Weighted': site_weighted,
            'Weighted progress against target': main_weighted,
        }
        
        # Add month-specific columns
        for month in MONTHS:
            year = MONTH_YEARS[month]
            target_val = month_activities.get(month, '')
            progress_val = month_progress[month]
            
            row_data[f'Target - {month}-{year}'] = target_val
            row_data[f'% work done- {month} Status'] = f"{progress_val:.0f}%"
            
            if progress_val == 100:
                achieved = target_val if target_val else f'No target for {month}'
            elif progress_val == 0:
                achieved = 'No progress' if target_val else f'No target for {month}'
            else:
                achieved = f'{progress_val:.0f}% completed'
            
            row_data[f'Achieved- {month} {year}'] = achieved
            row_data[f'Responsible Person- {month}'] = ''
            row_data[f'Delay Reasons- {month}'] = ''
        
        progress_data.append(row_data)
        milestone_counter += 1
        
        month_info = ", ".join([f"{m}: {month_progress[m]:.1f}%" for m in MONTHS])
        logger.info(f"Block {block_name} -> {month_info}")
    
    # Create DataFrame with dynamic columns
    columns = [
        'Milestone', 
        'Block', 
        f'Target to be complete by {TARGET_END_MONTH}-{TARGET_END_YEAR}'
    ]
    
    for month in MONTHS:
        year = MONTH_YEARS[month]
        columns.extend([
            f'Target - {month}-{year}',
            f'% work done- {month} Status',
            f'Achieved- {month} {year}',
            f'Responsible Person- {month}',
            f'Delay Reasons- {month}'
        ])
    
    columns.extend(['Site Weighted', 'Weighted progress against target'])
    
    df = pd.DataFrame(progress_data, columns=columns)
    logger.info(f"Created DataFrame with {len(df)} rows for months: {MONTHS}")
    return df

# -----------------------------------------------------------------------------
# EXCEL REPORT GENERATION
# -----------------------------------------------------------------------------

def write_wcc_excel_report_consolidated(df, filename):
    """Generate formatted Excel report with dynamic month columns"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Wave City Club- Progress Against Milestones'
    
    # Add main title
    title_row = ["Wave City Club- Progress Against Milestones"]
    ws.append(title_row)
    ws.merge_cells(f'A1:{get_column_letter(len(df.columns))}1')
    
    # Add date row
    current_date = datetime.now().strftime("%d-%m-%Y")
    date_row = [f"Report Generated on: {current_date}"]
    ws.append(date_row)
    ws.merge_cells(f'A2:{get_column_letter(len(df.columns))}2')
    
    # Add month info row
    month_info = f"Months Covered: {', '.join(MONTHS)}"
    month_info_row = [month_info]
    ws.append(month_info_row)
    ws.merge_cells(f'A3:{get_column_letter(len(df.columns))}3')
    
    # Add empty row
    ws.append([])
    
    # Add DataFrame data with percentage formatting for weighted progress
    for row in dataframe_to_rows(df, index=False, header=True):
        # Format the weighted progress column (last column) to add % symbol
        if len(row) > 0 and isinstance(row[-1], (int, float)) and row[-1] != '':
            row[-1] = f"{row[-1]:.3f}%"
        ws.append(row)
    
    # Add Sum row - Only for the weighted progress column
    weighted_sum = df['Weighted progress against target'].sum()
    
    # Create sum row with blanks for all columns except the weighted progress column
    sum_row = [''] * len(df.columns)
    sum_row[-2] = 'Sum'  # Site Weighted column
    sum_row[-1] = f'{weighted_sum:.3f}%'  # Weighted progress column
    ws.append(sum_row)
    
    # Define styles
    title_font = Font(bold=True, size=12)
    header_font = Font(bold=True, size=8)
    normal_font = Font(bold=False, size=8)
    date_font = Font(bold=False, size=10, color="666666")
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin = Side(style='thin', color='000000')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    light_grey_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
    light_blue_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    
    # Style title (light grey background)
    ws['A1'].font = title_font
    ws['A1'].alignment = center
    ws['A1'].fill = light_grey_fill
    
    # Style date row
    ws['A2'].font = date_font
    ws['A2'].alignment = center
    
    # Style month info row
    ws['A3'].font = date_font
    ws['A3'].alignment = center
    
    # Style header row (row 5) with light grey background
    header_row = 5
    for cell in ws[header_row]:
        cell.font = header_font
        cell.alignment = center
        cell.border = border
        cell.fill = light_grey_fill
    
    # Style data rows
    data_start = 6
    data_end = ws.max_row - 1  # Exclude sum row for now
    
    for row_num in range(data_start, data_end + 1):
        for col_num in range(1, len(df.columns) + 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = normal_font
            cell.border = border
            
            # Alignment based on column type
            if col_num in [1, 2, 3] or 'Target' in str(ws.cell(row=header_row, column=col_num).value or ''):  # Text columns
                cell.alignment = left
            else:  # Numeric columns
                cell.alignment = center
    
    # Style sum row with light blue background
    sum_row_num = ws.max_row
    for col_num in range(1, len(df.columns) + 1):
        cell = ws.cell(row=sum_row_num, column=col_num)
        cell.font = header_font
        cell.border = border
        cell.fill = light_blue_fill
        cell.alignment = center
    
    # Dynamic column width adjustment
    for col_num in range(1, len(df.columns) + 1):
        col_letter = get_column_letter(col_num)
        
        # Calculate optimal width based on column content
        max_length = 0
        for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=col_num, max_col=col_num):
            for cell in row:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
        
        # Set minimum and maximum width constraints
        calculated_width = min(max(max_length + 2, 8), 15)
        ws.column_dimensions[col_letter].width = calculated_width
    
    # Set row heights
    ws.row_dimensions[1].height = 25  # Title row
    ws.row_dimensions[2].height = 20  # Date row
    ws.row_dimensions[3].height = 20  # Month info row
    for i in range(5, ws.max_row + 1):
        ws.row_dimensions[i].height = 25
    
    wb.save(filename)
    logger.info(f'Dynamic report saved to {filename}')


def get_unique_filename(base_name):
    """
    If file exists, append (1), (2), etc.
    """
    if not os.path.exists(base_name):
        return base_name

    name, ext = os.path.splitext(base_name)
    counter = 1
    new_name = f"{name}({counter}){ext}"
    while os.path.exists(new_name):
        counter += 1
        new_name = f"{name}({counter}){ext}"
    return new_name
# -----------------------------------------------------------------------------
# MAIN FUNCTION
# -----------------------------------------------------------------------------

def main():
    """Main execution function for dynamic report generation"""
    logger.info("=== STARTING WAVE CITY CLUB REPORT WITH DYNAMIC MONTHS AND TRACKER SELECTION ===")
    
    try:
        # Initialize COS client
        cos = init_cos()
        
        # Get latest tracker path and setup dynamic months
        get_latest_tracker_paths(cos)
        
        # Verify that we have the required tracker path
        if not WCC_TRACKER_KEY:
            logger.error("❌ Failed to find Wave City Club tracker file")
            return
        
        logger.info(f"Using months: {MONTHS}")
        logger.info(f"All months will be processed and displayed")
        
        # Get targets from KRA file with dynamic month support
        logger.info("Fetching Wave City Club targets from KRA file for dynamic months...")
        targets = get_wcc_targets_from_kra(cos)
        
        # Extract progress data for all dynamic months
        logger.info("Extracting progress data from tracker for ALL dynamic months...")
        df = get_wcc_progress_from_tracker_all_months(cos, targets, WCC_TRACKER_KEY)
        
        # Generate dynamic report
        current_date_for_filename = datetime.now().strftime('%d-%m-%Y')
        base_filename = f"Wave_City_Club Milestone Report ({current_date_for_filename}).xlsx"
        filename = get_unique_filename(base_filename)

        logger.info("Generating dynamic Excel report with ALL months")
        write_wcc_excel_report_consolidated(df, filename)
        
        logger.info("=== WAVE CITY CLUB REPORT GENERATION COMPLETE ===")
        logger.info(f"Report saved as: {filename}")
        
        # Log summary
        logger.info("Report Summary:")
        logger.info(f"  Generated Months: {MONTHS}")
        logger.info(f"  Processed Blocks: {len(targets)}")
        logger.info(f"  All months displayed in the report")
        
    except Exception as e:
        logger.error(f"Error in main execution: {e}")
        raise

if __name__ == "__main__":
    main()


