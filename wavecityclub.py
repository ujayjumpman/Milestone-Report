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

# Configuration for current reporting month
CURRENT_REPORTING_MONTH = "June"  # Change this to "July" or "August" as needed
CURRENT_YEAR = "2025"

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
            logger.info(f"Row {row_num}: Block='{block_name}', June='{june_activity}', July='{july_activity}', August='{august_activity}'")
    
    logger.info(f"Extracted targets for {len(targets)} blocks from KRA")
    return targets

def find_activity_progress_in_sheet(sheet, target_activity, sheet_name, block_name=None):
    """
    Enhanced function to handle special cases for Block 1 and Fine Dine
    For these blocks, perform enhanced search within the entire sheet
    Modified to return 100% when there are no target activities
    """
    logger.info(f"=== DEBUG: Looking for activity '{target_activity}' in sheet '{sheet_name}' for block '{block_name}' ===")
    
    # Check if there's no target activity - return 100% in these cases
    if not target_activity or target_activity.strip() == '' or target_activity.lower() in ['no target', 'no target for june', 'no target for july', 'no target for august', '-']:
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
        
        return 0.0
    
    # Original logic for other blocks
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
    
    logger.warning(f"NO EXACT MATCH found for target: '{target_activity}'")
    return 0.0

def get_wcc_progress_from_tracker_all_months(cos, targets, tracker_key):
    """Extract progress data from tracker file - Only June data populated, July and August columns blank"""
    raw = download_file_bytes(cos, tracker_key)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    logger.info(f"Available tracker sheets: {wb.sheetnames}")
    
    progress_data = []
    milestone_counter = 1
    total_blocks = len(targets)
    site_weighted = round(100 / total_blocks, 2) if total_blocks > 0 else 0
    
    for block_name, month_activities in targets.items():
        logger.info(f"Processing block: {block_name}")
        
        # Get the corresponding tracker sheet name
        sheet_name = BLOCK_MAPPING.get(block_name)
        
        # Initialize progress - only calculate June, July and August will be blank
        june_progress = 0.0
        
        if not sheet_name:
            logger.warning(f"No sheet mapping found for block: {block_name}")
        elif sheet_name not in wb.sheetnames:
            logger.warning(f"Sheet '{sheet_name}' not found in tracker workbook")
        else:
            # Get the sheet and find progress only for June
            sheet = wb[sheet_name]
            june_activity = month_activities.get('June', '')
            june_progress = find_activity_progress_in_sheet(sheet, june_activity, sheet_name, block_name)
        
        # Calculate weighted progress for June only (July and August will be blank)
        june_weighted = round((site_weighted * june_progress) / 100, 3)
        
        # Determine achieved status for June only
        june_achieved = month_activities.get('June', '') if june_progress == 100 else ('No progress' if june_progress == 0 else f'{june_progress:.0f}% completed')
        
        # Handle "No target" cases for June
        if not month_activities.get('June', '').strip():
            june_achieved = 'No target for June'
        
        # Create row data in the consolidated format - July and August columns left blank
        row_data = {
            'Milestone': f"Milestone-{milestone_counter:02d}",
            'Activity': block_name,
            'Target to be complete by August-2025': month_activities.get('August', ''),
            'Target - June-2025': month_activities.get('June', ''),
            '% work done- June Status': f"{june_progress:.0f}%",
            'Site Weighted (June)': site_weighted,
            'Weighted progress against target (June)': june_weighted,  # Keep as number for sum calculation
            'Achieved- June 2025': june_achieved,
            'Target - July-2025': '',  # Left blank
            '% work done- July Status': '',  # Left blank
            'Site Weighted (July)': '',  # Left blank
            'Weighted progress against target (July)': '',  # Left blank
            'Achieved- July 2025': '',  # Left blank
            'Target - August-2025': '',  # Left blank - removed .1
            '% work done- August Status': '',  # Left blank
            'Site Weighted (August)': '',  # Left blank
            'Weighted progress against target (August)': '',  # Left blank
            'Achieved- August 2025': '',  # Left blank
            'Responsible Person': '',
            'Delay Reasons': ''
        }
        
        progress_data.append(row_data)
        milestone_counter += 1
        logger.info(f"Block {block_name} -> June: {june_progress:.1f}% (July and August columns left blank)")
    
    # Create DataFrame with consolidated column structure
    columns = [
        'Milestone',
        'Activity', 
        'Target to be complete by August-2025',
        'Target - June-2025',
        '% work done- June Status',
        'Site Weighted (June)',
        'Weighted progress against target (June)',
        'Achieved- June 2025',
        'Target - July-2025',
        '% work done- July Status',
        'Site Weighted (July)',
        'Weighted progress against target (July)',
        'Achieved- July 2025',
        'Target - August-2025',
        '% work done- August Status',
        'Site Weighted (August)',
        'Weighted progress against target (August)',
        'Achieved- August 2025',
        'Responsible Person',
        'Delay Reasons'
    ]
    
    df = pd.DataFrame(progress_data, columns=columns)
    logger.info(f"Created consolidated DataFrame with {len(df)} rows for June only")
    return df

# -----------------------------------------------------------------------------
# EXCEL REPORT GENERATION
# -----------------------------------------------------------------------------

def write_wcc_excel_report_consolidated(df, filename):
    """Generate formatted Excel report with consolidated format for all months"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Wave City Club- Progress Against Milestones'
    
    # Add main title
    ws.append(["Wave City Club- Progress Against Milestones"])
    ws.merge_cells('A1:T1')
    
    # Add date row
    current_date = datetime.now().strftime("%d-%m-%Y")
    ws.append([f"Report Generated on: {current_date}"])
    ws.merge_cells('A2:T2')
    
    # Add empty row
    ws.append([])
    
    # Add DataFrame data with percentage formatting for weighted progress
    for row in dataframe_to_rows(df, index=False, header=True):
        # Format the weighted progress column (column 7) to add % symbol
        if len(row) >= 7 and isinstance(row[6], (int, float)) and row[6] != '':
            row[6] = f"{row[6]:.3f}%"
        ws.append(row)
    
    # Add Sum row - Only June has sum, July and August are blank
    june_sum = df['Weighted progress against target (June)'].sum()
    
    sum_row = ['', '', '', '', '', 'Sum', f'{june_sum:.3f}%', '', '', '', '', '', '', '', '', '', '', '', '', '']
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
    
    # Style header row (row 4) with light grey background
    header_row = 4
    for cell in ws[header_row]:
        cell.font = header_font
        cell.alignment = center
        cell.border = border
        cell.fill = light_grey_fill
    
    # Style data rows
    data_start = 5
    data_end = ws.max_row - 1  # Exclude sum row for now
    
    for row_num in range(data_start, data_end + 1):
        for col_num in range(1, 21):  # Columns A to T
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = normal_font
            cell.border = border
            
            # Alignment based on column type
            if col_num in [1, 2, 3, 4, 8, 9, 13, 14, 18, 19, 20]:  # Text columns
                cell.alignment = left
            else:  # Numeric columns
                cell.alignment = center
    
    # Style sum row with light blue background
    sum_row_num = ws.max_row
    for col_num in range(1, 21):  # Columns A to T
        cell = ws.cell(row=sum_row_num, column=col_num)
        cell.font = header_font
        cell.border = border
        cell.fill = light_blue_fill
        cell.alignment = center
    
    # Adjust column widths for consolidated format
    column_widths = {
        1: 8,   # Milestone
        2: 12,  # Activity
        3: 12,  # Target August
        4: 12,  # Target June
        5: 8,   # % work done June
        6: 6,   # Site Weighted
        7: 8,   # Weighted progress June
        8: 12,  # Achieved June
        9: 12,  # Target July
        10: 8,  # % work done July
        11: 6,  # Site Weighted July
        12: 8,  # Weighted progress July
        13: 12, # Achieved July
        14: 12, # Target August
        15: 8,  # % work done August
        16: 6,  # Site Weighted August
        17: 8,  # Weighted progress August
        18: 12, # Achieved August
        19: 12, # Responsible Person
        20: 10  # Delay Reasons
    }
    
    for col_num, width in column_widths.items():
        ws.column_dimensions[get_column_letter(col_num)].width = width
    
    # Set row heights
    ws.row_dimensions[1].height = 25  # Title row
    ws.row_dimensions[2].height = 20  # Date row
    for i in range(4, ws.max_row + 1):
        ws.row_dimensions[i].height = 25
    
    wb.save(filename)
    logger.info(f'Consolidated report saved to {filename}')

# -----------------------------------------------------------------------------
# MAIN FUNCTION
# -----------------------------------------------------------------------------

def main():
    """Main execution function for consolidated report"""
    try:
        # Initialize COS client
        cos = init_cos()
        
        # Get targets from KRA file
        logger.info("Fetching Wave City Club targets from KRA file for consolidated reporting...")
        targets = get_wcc_targets_from_kra(cos)
        
        # Determine tracker file to use
        logger.info("Determining tracker file to use...")
        try:
            cos.head_object(Bucket=BUCKET, Key=WCC_TRACKER_KEY)
            tracker_key = WCC_TRACKER_KEY
            logger.info(f"Using configured tracker key: {tracker_key}")
        except Exception:
            tracker_key = find_latest_wcc_tracker_key(cos)
        
        # Extract progress data for all months
        logger.info("Extracting progress data from tracker for June only (July/August blank)...")
        df = get_wcc_progress_from_tracker_all_months(cos, targets, tracker_key)
        
        # Generate consolidated report
        current_date_for_filename = datetime.now().strftime('%d-%m-%Y')
        filename = f"Wave_City_Club Milestone Report ({current_date_for_filename}).xlsx"
        logger.info("Generating consolidated Excel report")
        write_wcc_excel_report_consolidated(df, filename)
        
        logger.info("Consolidated report generation completed successfully!")
        
    except Exception as e:
        logger.error(f"Error in main execution: {e}")
        raise

if __name__ == "__main__":
    main()
