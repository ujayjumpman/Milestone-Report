# import os
# import logging
# from io import BytesIO
# from datetime import datetime
# import pandas as pd
# from openpyxl import Workbook, load_workbook
# from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# from openpyxl.utils.dataframe import dataframe_to_rows
# from dotenv import load_dotenv
# import ibm_boto3
# from ibm_botocore.client import Config

# # =============== CONFIG / CONSTANTS ===============
# load_dotenv()
# logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
# logger = logging.getLogger(__name__)

# COS_API_KEY = os.getenv("COS_API_KEY")
# COS_CRN = os.getenv("COS_SERVICE_INSTANCE_CRN")
# COS_ENDPOINT = os.getenv("COS_ENDPOINT")
# BUCKET = os.getenv("COS_BUCKET_NAME")
# EWS_LIG_STRUCTURE_KEY = os.getenv("EWS_LIG_STRUCTURE_TRACKER_PATH")
# EWS_LIG_KRA_KEY = os.getenv("KRA_FILE_PATH")

# MONTHS = ["June", "July", "August"]
# MONTH_TO_NUM = {"June": 6, "July": 7, "August": 8}

# KRA_SHEET = "EW-LI P4 Targets Till August "
# TOWER1_TARGETS_CELLS = {'June': 'B4', 'July': 'C4', 'August': 'D4'}
# TOWER3_TARGETS_CELLS = {'June': 'B12', 'July': 'C12', 'August': 'D12'}
# TOWER2_TARGETS_CELLS = {'June': 'B19', 'July': 'C19', 'August': 'D19'}

# TRACKER_SHEET = "Revised Baseline 45daysNGT+Rai"

# # Tower 1 rows/cols: rows 5–22, columns D, H, L, P
# TOWER1_POUR_COLS = ['D', 'H', 'L', 'P']
# TOWER1_ROW_START, TOWER1_ROW_END = 5, 22

# # Tower 3 rows/cols: rows 27–40, columns D, H, L, P (as per your screenshot)
# TOWER3_POUR_COLS = ['D', 'H', 'L', 'P']
# TOWER3_ROW_START, TOWER3_ROW_END = 27, 40

# # Tower 2 rows/cols: rows 5–22, columns U, Y, AC, AG
# TOWER2_POUR_COLS = ['U', 'Y', 'AC', 'AG']
# TOWER2_ROW_START, TOWER2_ROW_END = 5, 22

# YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
# GREY = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

# def get_previous_months():
#     # Modified to return only June for display purposes
#     return ["June"]

# def detect_tracker_year(sheet, pour_cols, row_start, row_end):
#     years_found = set()
#     for col in pour_cols:
#         for row in range(row_start, row_end+1):
#             cell_value = sheet[f"{col}{row}"].value
#             if cell_value is None: continue
#             parsed_date = None
#             if isinstance(cell_value, datetime):
#                 parsed_date = cell_value
#             elif isinstance(cell_value, str):
#                 parsed_date = pd.to_datetime(cell_value, errors='coerce', dayfirst=True)
#             if pd.notna(parsed_date):
#                 years_found.add(parsed_date.year)
#     return max(years_found) if years_found else datetime.now().year

# def init_cos():
#     return ibm_boto3.client(
#         "s3",
#         ibm_api_key_id=COS_API_KEY,
#         ibm_service_instance_id=COS_CRN,
#         config=Config(signature_version="oauth"),
#         endpoint_url=COS_ENDPOINT,
#     )

# def download_file_bytes(cos, key):
#     obj = cos.get_object(Bucket=BUCKET, Key=key)
#     return obj["Body"].read()

# def get_targets_from_kra(wb, sheet_name, cell_map):
#     sheet = wb[sheet_name]
#     targets = {}
#     for month, cell in cell_map.items():
#         value = sheet[cell].value
#         try:
#             targets[month] = int(str(value).strip().split()[0]) if value else 0
#         except Exception:
#             targets[month] = 0
#     return targets

# def count_pours(sheet, pour_cols, row_start, row_end, months, year):
#     month_counts = {m: 0 for m in months}
#     for month in months:
#         month_num = MONTH_TO_NUM[month]
#         count = 0
#         for col in pour_cols:
#             for row in range(row_start, row_end + 1):
#                 cell_value = sheet[f"{col}{row}"].value
#                 if cell_value is None:
#                     continue
#                 parsed_date = None
#                 if isinstance(cell_value, datetime):
#                     parsed_date = cell_value
#                 elif isinstance(cell_value, str) and cell_value.strip():
#                     parsed_date = pd.to_datetime(cell_value, dayfirst=True, errors='coerce')
#                     if pd.isna(parsed_date):
#                         for fmt in ['%d-%b-%y', '%d-%b-%Y', '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d']:
#                             try:
#                                 parsed_date = pd.to_datetime(cell_value, format=fmt, errors='coerce')
#                                 if pd.notna(parsed_date): break
#                             except: continue
#                 if pd.notna(parsed_date) and parsed_date.month == month_num and parsed_date.year == year:
#                     count += 1
#         month_counts[month] = count
#     return month_counts

# def build_structure_dataframe(tower_name, targets, completed):
#     # Only show results for June, but keep all targets for "Target Till August"
#     prev_months = get_previous_months()  # This will return only ["June"]
#     weightage = 100
    
#     # Calculate cumulative targets (still use all months for "Target Till August")
#     cum_targets = {}
#     cum_completed = {}
#     for i, m in enumerate(MONTHS):
#         months_to_count = MONTHS[:i+1]
#         cum_targets[m] = sum(targets[mm] for mm in months_to_count)
#         cum_completed[m] = sum(completed[mm] for mm in months_to_count if mm in prev_months)

#     def pct(m):
#         # Only show percentage for June
#         if m != "June":
#             return ""  # Leave July and August blank
#         t = cum_targets[m]
#         d = cum_completed[m]
#         if t == 0: return "0.0%"
#         val = min(round((d / t) * 100, 2), 100)
#         return f"{val}%"

#     row = {
#         "Milestone": f"{tower_name} Structure",
#         "Target Till August": f"{sum(targets.values())} Pours ({targets['June']} June, {targets['July']} July, {targets['August']} August)",
#         "% Work Done against Target-Till June": pct("June"),
#         "% Work Done against Target-Till July": "",  # Blank
#         "% Work Done against Target-Till August": "",  # Blank
#         "Weightage": weightage,
#         "Weighted Delay against Targets": "",
#         "Target achieved in June": f"{completed.get('June', 0)} out of {targets.get('June', 0)}",
#         "Target achieved in July": "",  # Blank
#         "Target achieved in August": "",  # Blank
#         "Total achieved": f"{completed.get('June', 0)} out of {sum(targets.values())}",  # Only June achieved vs total target
#         "Delay Reasons": "",
#     }
    
#     # Calculate weighted delay only for June
#     june_pct_str = pct("June")
#     if june_pct_str:
#         try:
#             june_pct = float(june_pct_str.replace("%", ""))
#             row["Weighted Delay against Targets"] = f"{round((june_pct * weightage) / 100, 2)}%"
#         except Exception:
#             row["Weighted Delay against Targets"] = "0.0%"
    
#     df = pd.DataFrame([row])
#     return df

# def write_excel_report(dfs, filename):
#     wb = Workbook()
#     ws = wb.active
#     ws.title = "EWS-LIG Milestones"

#     # Add title and date at the top
#     current_date = datetime.now().strftime("%d-%m-%Y")
#     ws.append(["EWS-LIG Milestones Report"])
#     ws.append([f"Report Generated on: {current_date}"])
#     ws.append([])  # Empty row for spacing

#     # Define styles
#     bold_font = Font(bold=True)
#     normal_font = Font(bold=False)
#     title_font = Font(bold=True, size=14)
#     date_font = Font(bold=False, size=10, color="666666")
#     center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
#     left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
#     thin = Side(style="thin", color="000000")
#     border = Border(top=thin, bottom=thin, left=thin, right=thin)
    
#     # Get max columns for merging (from first dataframe)
#     max_cols = len(dfs[0][1].columns) if dfs else 12  # fallback to 12 columns
    
#     # Style title row (row 1)
#     ws.merge_cells(f'A1:{get_column_letter(max_cols)}1')
#     ws['A1'].font = title_font
#     ws['A1'].alignment = center_align
#     ws['A1'].fill = GREY
    
#     # Style date row (row 2)
#     ws.merge_cells(f'A2:{get_column_letter(max_cols)}2')
#     ws['A2'].font = date_font
#     ws['A2'].alignment = center_align

#     for title, df, total_label in dfs:
#         # Section title row
#         ws.append([title])
#         title_row = ws.max_row
#         ws.merge_cells(start_row=title_row, start_column=1,
#                        end_row=title_row, end_column=len(df.columns))
#         for cell in ws[title_row]:
#             cell.fill = GREY
#             cell.font = bold_font
#             cell.alignment = center_align
#             cell.border = border

#         # DataFrame rows
#         for r in dataframe_to_rows(df, index=False, header=True):
#             ws.append(r)
#         header_row = title_row + 1
#         body_start = header_row + 1
#         body_end = ws.max_row
        
#         # Header styling
#         for cell in ws[header_row]:
#             cell.font = bold_font
#             cell.alignment = center_align
#             cell.border = border
            
#         # Body styling
#         for r in range(body_start, body_end + 1):
#             for cell in ws[r]:
#                 cell.font = normal_font
#                 cell.alignment = left_align if cell.col_idx in (1, 2) else center_align
#                 cell.border = border
                
#         # Total delay row
#         try:
#             total_delay = sum(float(str(v).strip('%')) for v in df["Weighted Delay against Targets"] if v and str(v).strip())
#         except Exception:
#             total_delay = 0
#         weighted_delay_col_idx = None
#         for idx, col_name in enumerate(df.columns, start=1):
#             if col_name == "Weighted Delay against Targets":
#                 weighted_delay_col_idx = idx
#                 break
#         total_row_data = [""] * len(df.columns)
#         if weighted_delay_col_idx:
#             total_row_data[weighted_delay_col_idx - 1] = f"{round(total_delay, 2)}%"
#             total_row_data[0] = total_label
#         ws.append(total_row_data)
#         delay_row = ws.max_row
#         for idx, cell in enumerate(ws[delay_row], start=1):
#             cell.font = bold_font
#             cell.fill = YELLOW
#             cell.alignment = left_align if idx == 1 else center_align
#             cell.border = border

#     # Column widths
#     for col in ws.columns:
#         max_len = max(len(str(cell.value or "")) for cell in col)
#         ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)
    
#     # Row heights
#     for r in range(1, ws.max_row + 1):
#         ws.row_dimensions[r].height = 22
    
#     wb.save(filename)
#     logger.info(f"EWS-LIG report saved to {filename}")

# def main():
#     cos = init_cos()
#     kra_raw = download_file_bytes(cos, EWS_LIG_KRA_KEY)
#     kra_wb = load_workbook(filename=BytesIO(kra_raw), data_only=True)
#     tracker_raw = download_file_bytes(cos, EWS_LIG_STRUCTURE_KEY)
#     tracker_wb = load_workbook(filename=BytesIO(tracker_raw), data_only=True)
#     sheet = tracker_wb[TRACKER_SHEET]

#     prev_months = get_previous_months()
#     tracker_year = detect_tracker_year(sheet, TOWER1_POUR_COLS, TOWER1_ROW_START, TOWER1_ROW_END)

#     # Tower 1
#     targets_t1 = get_targets_from_kra(kra_wb, KRA_SHEET, TOWER1_TARGETS_CELLS)
#     completed_t1 = count_pours(sheet, TOWER1_POUR_COLS, TOWER1_ROW_START, TOWER1_ROW_END, MONTHS, tracker_year)
#     df_t1 = build_structure_dataframe("Tower 1", targets_t1, completed_t1)

#     # Tower 3
#     targets_t3 = get_targets_from_kra(kra_wb, KRA_SHEET, TOWER3_TARGETS_CELLS)
#     completed_t3 = count_pours(sheet, TOWER3_POUR_COLS, TOWER3_ROW_START, TOWER3_ROW_END, MONTHS, tracker_year)
#     df_t3 = build_structure_dataframe("Tower 3", targets_t3, completed_t3)

#     # Tower 2
#     targets_t2 = get_targets_from_kra(kra_wb, KRA_SHEET, TOWER2_TARGETS_CELLS)
#     completed_t2 = count_pours(sheet, TOWER2_POUR_COLS, TOWER2_ROW_START, TOWER2_ROW_END, MONTHS, tracker_year)
#     df_t2 = build_structure_dataframe("Tower 2", targets_t2, completed_t2)

#     filename = f"EWS_LIG_Milestone_Report ({datetime.now():%Y-%m-%d}).xlsx"
#     dfs = [
#         ("Tower 1 Structure Progress Against Milestones", df_t1, "Total Delay Tower 1 Structure"),
#         ("Tower 3 Structure Progress Against Milestones", df_t3, "Total Delay Tower 3 Structure"),
#         ("Tower 2 Structure Progress Against Milestones", df_t2, "Total Delay Tower 2 Structure"),
#     ]
#     write_excel_report(dfs, filename)
#     logger.info("EWS-LIG milestone report generation completed successfully!")

# if __name__ == "__main__":
#     main()



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

# =============== CONFIG / CONSTANTS ===============
load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

COS_API_KEY = os.getenv("COS_API_KEY")
COS_CRN = os.getenv("COS_SERVICE_INSTANCE_CRN")
COS_ENDPOINT = os.getenv("COS_ENDPOINT")
BUCKET = os.getenv("COS_BUCKET_NAME")
EWS_LIG_KRA_KEY = os.getenv("KRA_FILE_PATH")

# Dynamic tracker path - will be set by get_latest_tracker_paths()
EWS_LIG_STRUCTURE_KEY = None

# Dynamic months - will be set based on tracker date
MONTHS = []
MONTH_TO_NUM = {}
TRACKER_DATE = None
PROCESSING_MONTHS = []  # Months to actually process (previous months)

KRA_SHEET = "EW-LI P4 Targets Till August "
TRACKER_SHEET = "Revised Baseline 45daysNGT+Rai"

# Dynamic target cells - will be updated based on months
TOWER1_TARGETS_CELLS = {}
TOWER3_TARGETS_CELLS = {}
TOWER2_TARGETS_CELLS = {}

# Tower 1 rows/cols: rows 5–22, columns D, H, L, P
TOWER1_POUR_COLS = ['D', 'H', 'L', 'P']
TOWER1_ROW_START, TOWER1_ROW_END = 5, 22

# Tower 3 rows/cols: rows 27–40, columns D, H, L, P (as per your screenshot)
TOWER3_POUR_COLS = ['D', 'H', 'L', 'P']
TOWER3_ROW_START, TOWER3_ROW_END = 27, 40

# Tower 2 rows/cols: rows 5–22, columns U, Y, AC, AG
TOWER2_POUR_COLS = ['U', 'Y', 'AC', 'AG']
TOWER2_ROW_START, TOWER2_ROW_END = 5, 22

YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
GREY = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

# =============== DYNAMIC SETUP FUNCTIONS ===============

def list_files_in_folder(cos, folder_prefix):
    """List all files in a specific folder (prefix) in the COS bucket"""
    try:
        response = cos.list_objects_v2(Bucket=BUCKET, Prefix=folder_prefix)
        files = []
        if 'Contents' in response:
            for obj in response['Contents']:
                # Only include actual files, not folder markers
                if not obj['Key'].endswith('/'):
                    files.append(obj['Key'])
        return files
    except Exception as e:
        logger.error(f"Error listing files in folder {folder_prefix}: {e}")
        return []

def extract_date_from_filename(filename):
    """Extract date from filename in format (dd-mm-yyyy)"""
    # Pattern to match date in parentheses like (01-07-2025)
    pattern = r'\((\d{2}-\d{2}-\d{4})\)'
    match = re.search(pattern, filename)
    if match:
        date_str = match.group(1)
        try:
            # Parse the date
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

def setup_dynamic_months_and_targets(tracker_date):
    """Setup months and target cells based on tracker date"""
    global MONTHS, MONTH_TO_NUM, TRACKER_DATE, PROCESSING_MONTHS
    global TOWER1_TARGETS_CELLS, TOWER3_TARGETS_CELLS, TOWER2_TARGETS_CELLS
    
    TRACKER_DATE = tracker_date
    tracker_month = tracker_date.month
    tracker_year = tracker_date.year
    
    logger.info(f"=== SETTING UP DYNAMIC MONTHS FOR EWS-LIG ===")
    logger.info(f"Tracker date: {tracker_date.strftime('%d-%m-%Y')} (Month: {tracker_month})")
    
    # Calculate the months to work with
    target_month = tracker_month - 1
    if target_month < 1:
        target_month = 12
        tracker_year -= 1
    
    # Generate 3 consecutive months ending with target_month
    months_numbers = []
    for i in range(2, -1, -1):  # 2, 1, 0 to get 3 months
        month_num = target_month - i
        if month_num < 1:
            month_num += 12
        months_numbers.append(month_num)
    
    MONTHS = [get_month_name(m) for m in months_numbers]
    
    # Create month to number mapping
    MONTH_TO_NUM = {month: num for month, num in zip(MONTHS, months_numbers)}
    
    # For processing, include months up to the target month
    current_date = datetime.now()
    PROCESSING_MONTHS = []
    
    for month_name in MONTHS:
        month_num = MONTH_TO_NUM[month_name]
        
        if month_num < tracker_month or (month_num > tracker_month and tracker_month <= 3):
            PROCESSING_MONTHS.append(month_name)
    
    if not PROCESSING_MONTHS and MONTHS:
        PROCESSING_MONTHS = [MONTHS[-1]]
    
    logger.info(f"Generated MONTHS: {MONTHS}")
    logger.info(f"MONTH_TO_NUM: {MONTH_TO_NUM}")
    logger.info(f"PROCESSING_MONTHS: {PROCESSING_MONTHS}")
    
    # Setup target cells dynamically based on months
    setup_target_cells()

def setup_target_cells():
    """Setup target cells for each tower based on dynamic months"""
    global TOWER1_TARGETS_CELLS, TOWER3_TARGETS_CELLS, TOWER2_TARGETS_CELLS
    
    # Map months to columns - this may need adjustment based on your KRA structure
    base_cols = ['B', 'C', 'D', 'E', 'F', 'G']  # B=col1, C=col2, D=col3, etc.
    
    # Tower 1 targets (row 4)
    TOWER1_TARGETS_CELLS = {}
    for i, month in enumerate(MONTHS):
        if i < len(base_cols):
            TOWER1_TARGETS_CELLS[month] = f'{base_cols[i]}4'
    
    # Tower 3 targets (row 12) 
    TOWER3_TARGETS_CELLS = {}
    for i, month in enumerate(MONTHS):
        if i < len(base_cols):
            TOWER3_TARGETS_CELLS[month] = f'{base_cols[i]}12'
    
    # Tower 2 targets (row 19)
    TOWER2_TARGETS_CELLS = {}
    for i, month in enumerate(MONTHS):
        if i < len(base_cols):
            TOWER2_TARGETS_CELLS[month] = f'{base_cols[i]}19'
    
    logger.info(f"Setup target cells:")
    logger.info(f"  Tower 1: {TOWER1_TARGETS_CELLS}")
    logger.info(f"  Tower 3: {TOWER3_TARGETS_CELLS}")
    logger.info(f"  Tower 2: {TOWER2_TARGETS_CELLS}")

def get_latest_tracker_paths(cos):
    """Get the latest tracker file path for EWS LIG P4"""
    global EWS_LIG_STRUCTURE_KEY
    
    logger.info("=== FINDING LATEST EWS LIG P4 TRACKER FILES ===")
    
    # List all files in EWS LIG P4 folder
    ews_files = list_files_in_folder(cos, "EWS LIG P4/")
    logger.info(f"Found {len(ews_files)} files in EWS LIG P4 folder")
    
    # Define tracker pattern
    tracker_pattern = r'Structure Work Tracker.*\.xlsx$'
    
    matching_files = []
    latest_date = None
    
    logger.info(f"--- Looking for EWS LIG P4 Structure tracker files ---")
    
    for file_path in ews_files:
        filename = os.path.basename(file_path)
        if re.search(tracker_pattern, filename, re.IGNORECASE):
            file_date = extract_date_from_filename(filename)
            if file_date:
                matching_files.append((file_path, file_date))
                logger.info(f"Found: {filename} with date {file_date.strftime('%d-%m-%Y')}")
                
                # Track the latest date
                if latest_date is None or file_date > latest_date:
                    latest_date = file_date
            else:
                logger.warning(f"Found matching file but no date: {filename}")
    
    if matching_files:
        # Sort by date and get the latest
        latest_file = max(matching_files, key=lambda x: x[1])
        EWS_LIG_STRUCTURE_KEY = latest_file[0]
        logger.info(f"✅ Latest EWS LIG P4 tracker: {EWS_LIG_STRUCTURE_KEY} ({latest_file[1].strftime('%d-%m-%Y')})")
    else:
        logger.error(f"❌ No EWS LIG P4 tracker files found!")
        EWS_LIG_STRUCTURE_KEY = None
    
    logger.info(f"\n=== FINAL EWS LIG P4 TRACKER PATH ===")
    logger.info(f"EWS_LIG_STRUCTURE_KEY: {EWS_LIG_STRUCTURE_KEY}")
    logger.info(f"Latest tracker date found: {latest_date.strftime('%d-%m-%Y') if latest_date else 'None'}")
    
    # Setup dynamic months and targets based on the latest tracker date
    if latest_date:
        setup_dynamic_months_and_targets(latest_date)
    else:
        logger.error("No valid tracker date found - using default setup")
        # Fallback to current date minus 1 month
        fallback_date = datetime.now()
        setup_dynamic_months_and_targets(fallback_date)
    
    # Verify tracker was found
    if not EWS_LIG_STRUCTURE_KEY:
        raise Exception("Could not find latest EWS LIG P4 tracker file")
    
    return EWS_LIG_STRUCTURE_KEY

def get_previous_months():
    """Return months that should be processed based on tracker date"""
    return PROCESSING_MONTHS

# =============== CORE FUNCTIONS ===============

def detect_tracker_year(sheet, pour_cols, row_start, row_end):
    years_found = set()
    for col in pour_cols:
        for row in range(row_start, row_end+1):
            cell_value = sheet[f"{col}{row}"].value
            if cell_value is None: continue
            parsed_date = None
            if isinstance(cell_value, datetime):
                parsed_date = cell_value
            elif isinstance(cell_value, str):
                parsed_date = pd.to_datetime(cell_value, errors='coerce', dayfirst=True)
            if pd.notna(parsed_date):
                years_found.add(parsed_date.year)
    return max(years_found) if years_found else datetime.now().year

def init_cos():
    return ibm_boto3.client(
        "s3",
        ibm_api_key_id=COS_API_KEY,
        ibm_service_instance_id=COS_CRN,
        config=Config(signature_version="oauth"),
        endpoint_url=COS_ENDPOINT,
    )

def download_file_bytes(cos, key):
    if not key:
        raise ValueError("File key cannot be None or empty")
    obj = cos.get_object(Bucket=BUCKET, Key=key)
    return obj["Body"].read()

def get_targets_from_kra(wb, sheet_name, cell_map):
    sheet = wb[sheet_name]
    targets = {}
    for month, cell in cell_map.items():
        value = sheet[cell].value
        try:
            targets[month] = int(str(value).strip().split()[0]) if value else 0
        except Exception:
            targets[month] = 0
    return targets

def count_pours(sheet, pour_cols, row_start, row_end, months, year):
    month_counts = {m: 0 for m in months}
    for month in months:
        month_num = MONTH_TO_NUM[month]
        count = 0
        for col in pour_cols:
            for row in range(row_start, row_end + 1):
                cell_value = sheet[f"{col}{row}"].value
                if cell_value is None:
                    continue
                parsed_date = None
                if isinstance(cell_value, datetime):
                    parsed_date = cell_value
                elif isinstance(cell_value, str) and cell_value.strip():
                    parsed_date = pd.to_datetime(cell_value, dayfirst=True, errors='coerce')
                    if pd.isna(parsed_date):
                        for fmt in ['%d-%b-%y', '%d-%b-%Y', '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d']:
                            try:
                                parsed_date = pd.to_datetime(cell_value, format=fmt, errors='coerce')
                                if pd.notna(parsed_date): break
                            except: continue
                if pd.notna(parsed_date) and parsed_date.month == month_num and parsed_date.year == year:
                    count += 1
        month_counts[month] = count
    return month_counts

def build_structure_dataframe(tower_name, targets, completed):
    """Build dataframe with dynamic month processing"""
    prev_months = get_previous_months()  
    weightage = 100
    
    # Calculate cumulative targets and completed for all months
    cum_targets = {}
    cum_completed = {}
    for i, m in enumerate(MONTHS):
        months_to_count = MONTHS[:i+1]
        cum_targets[m] = sum(targets[mm] for mm in months_to_count)
        cum_completed[m] = sum(completed[mm] for mm in months_to_count if mm in prev_months)

    def pct(m):
        # Only show percentage for processing months
        if m not in prev_months:
            return ""  # Leave future months blank
        t = cum_targets[m]
        d = cum_completed[m]
        if t == 0: return "0.0%"
        val = min(round((d / t) * 100, 2), 100)
        return f"{val}%"

    # Build target text dynamically
    target_parts = []
    total_target = 0
    for month in MONTHS:
        month_target = targets[month]
        total_target += month_target
        if month_target > 0:
            target_parts.append(f"{month_target} {month}")
    
    target_text = f"{total_target} Pours ({', '.join(target_parts)})"

    row = {
        "Milestone": f"{tower_name} Structure",
        "Target Till August": target_text,  # Column name kept for compatibility
        "Weightage": weightage,
        "Weighted Delay against Targets": "",
        "Total achieved": "",
        "Delay Reasons": "",
    }
    
    # Add percentage columns dynamically
    for month in MONTHS:
        row[f"% Work Done against Target-Till {month}"] = pct(month)
    
    # Add target achieved columns dynamically
    for month in MONTHS:
        if month in prev_months:
            month_target = targets[month]
            month_completed = completed[month]
            row[f"Target achieved in {month}"] = f"{month_completed} out of {month_target}"
        else:
            row[f"Target achieved in {month}"] = ""
    
    # Calculate total achieved vs total target (only for processing months)
    total_achieved = sum(completed[month] for month in prev_months)
    row["Total achieved"] = f"{total_achieved} out of {total_target}"
    
    # Calculate weighted delay for last processing month
    if prev_months:
        try:
            last_processed_month = prev_months[-1]
            last_pct_str = pct(last_processed_month)
            if last_pct_str:
                last_pct = float(last_pct_str.replace("%", ""))
                row["Weighted Delay against Targets"] = f"{round((last_pct * weightage) / 100, 2)}%"
        except Exception:
            row["Weighted Delay against Targets"] = "0.0%"
    
    # Create dynamic column list
    all_cols = ["Milestone", "Target Till August"]
    for month in MONTHS:
        all_cols.append(f"% Work Done against Target-Till {month}")
    all_cols.extend(["Weightage", "Weighted Delay against Targets"])
    for month in MONTHS:
        all_cols.append(f"Target achieved in {month}")
    all_cols.extend(["Total achieved", "Delay Reasons"])
    
    df = pd.DataFrame([row], columns=all_cols)
    return df

def write_excel_report(dfs, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "EWS-LIG Milestones"

    # Add title and date at the top
    current_date = datetime.now().strftime("%d-%m-%Y")
    ws.append(["EWS-LIG Milestones Report"])
    ws.append([f"Report Generated on: {current_date}"])
    ws.append([])  # Empty row for spacing

    # Define styles
    bold_font = Font(bold=True)
    normal_font = Font(bold=False)
    title_font = Font(bold=True, size=14)
    date_font = Font(bold=False, size=10, color="666666")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    
    # Get max columns for merging (from first dataframe)
    max_cols = len(dfs[0][1].columns) if dfs else 12  # fallback to 12 columns
    
    # Style title row (row 1)
    ws.merge_cells(f'A1:{get_column_letter(max_cols)}1')
    ws['A1'].font = title_font
    ws['A1'].alignment = center_align
    ws['A1'].fill = GREY
    
    # Style date row (row 2)
    ws.merge_cells(f'A2:{get_column_letter(max_cols)}2')
    ws['A2'].font = date_font
    ws['A2'].alignment = center_align

    for title, df, total_label in dfs:
        # Section title row
        ws.append([title])
        title_row = ws.max_row
        ws.merge_cells(start_row=title_row, start_column=1,
                       end_row=title_row, end_column=len(df.columns))
        for cell in ws[title_row]:
            cell.fill = GREY
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = border

        # DataFrame rows
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        header_row = title_row + 1
        body_start = header_row + 1
        body_end = ws.max_row
        
        # Header styling
        for cell in ws[header_row]:
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = border
            
        # Body styling
        for r in range(body_start, body_end + 1):
            for cell in ws[r]:
                cell.font = normal_font
                cell.alignment = left_align if cell.col_idx in (1, 2) else center_align
                cell.border = border
                
        # Total delay row
        try:
            total_delay = sum(float(str(v).strip('%')) for v in df["Weighted Delay against Targets"] if v and str(v).strip())
        except Exception:
            total_delay = 0
        weighted_delay_col_idx = None
        for idx, col_name in enumerate(df.columns, start=1):
            if col_name == "Weighted Delay against Targets":
                weighted_delay_col_idx = idx
                break
        total_row_data = [""] * len(df.columns)
        if weighted_delay_col_idx:
            total_row_data[weighted_delay_col_idx - 1] = f"{round(total_delay, 2)}%"
            total_row_data[0] = total_label
        ws.append(total_row_data)
        delay_row = ws.max_row
        for idx, cell in enumerate(ws[delay_row], start=1):
            cell.font = bold_font
            cell.fill = YELLOW
            cell.alignment = left_align if idx == 1 else center_align
            cell.border = border

    # Column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)
    
    # Row heights
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 22
    
    wb.save(filename)
    logger.info(f"EWS-LIG report saved to {filename}")

def main():
    logger.info("=== STARTING EWS-LIG REPORT WITH DYNAMIC MONTHS AND TRACKER SELECTION ===")
    
    try:
        cos = init_cos()
        
        # Get latest tracker path and setup dynamic months
        get_latest_tracker_paths(cos)
        
        # Verify that we have the required tracker path
        if not EWS_LIG_STRUCTURE_KEY:
            logger.error("❌ Failed to find EWS-LIG tracker file")
            return
        
        logger.info(f"Using months: {MONTHS}")
        logger.info(f"Processing months: {PROCESSING_MONTHS}")
        
        kra_raw = download_file_bytes(cos, EWS_LIG_KRA_KEY)
        kra_wb = load_workbook(filename=BytesIO(kra_raw), data_only=True)
        tracker_raw = download_file_bytes(cos, EWS_LIG_STRUCTURE_KEY)
        tracker_wb = load_workbook(filename=BytesIO(tracker_raw), data_only=True)
        sheet = tracker_wb[TRACKER_SHEET]

        prev_months = get_previous_months()
        tracker_year = detect_tracker_year(sheet, TOWER1_POUR_COLS, TOWER1_ROW_START, TOWER1_ROW_END)
        logger.info(f"Detected tracker year: {tracker_year}")

        # Tower 1
        targets_t1 = get_targets_from_kra(kra_wb, KRA_SHEET, TOWER1_TARGETS_CELLS)
        completed_t1 = count_pours(sheet, TOWER1_POUR_COLS, TOWER1_ROW_START, TOWER1_ROW_END, MONTHS, tracker_year)
        df_t1 = build_structure_dataframe("Tower 1", targets_t1, completed_t1)
        logger.info(f"Tower 1 - Targets: {targets_t1}, Completed: {completed_t1}")

        # Tower 3
        targets_t3 = get_targets_from_kra(kra_wb, KRA_SHEET, TOWER3_TARGETS_CELLS)
        completed_t3 = count_pours(sheet, TOWER3_POUR_COLS, TOWER3_ROW_START, TOWER3_ROW_END, MONTHS, tracker_year)
        df_t3 = build_structure_dataframe("Tower 3", targets_t3, completed_t3)
        logger.info(f"Tower 3 - Targets: {targets_t3}, Completed: {completed_t3}")

        # Tower 2
        targets_t2 = get_targets_from_kra(kra_wb, KRA_SHEET, TOWER2_TARGETS_CELLS)
        completed_t2 = count_pours(sheet, TOWER2_POUR_COLS, TOWER2_ROW_START, TOWER2_ROW_END, MONTHS, tracker_year)
        df_t2 = build_structure_dataframe("Tower 2", targets_t2, completed_t2)
        logger.info(f"Tower 2 - Targets: {targets_t2}, Completed: {completed_t2}")

        filename = f"EWS_LIG_Milestone_Report ({datetime.now():%Y-%m-%d}).xlsx"
        dfs = [
            ("Tower 1 Structure Progress Against Milestones", df_t1, "Total Delay Tower 1 Structure"),
            ("Tower 3 Structure Progress Against Milestones", df_t3, "Total Delay Tower 3 Structure"),
            ("Tower 2 Structure Progress Against Milestones", df_t2, "Total Delay Tower 2 Structure"),
        ]
        write_excel_report(dfs, filename)
        
        logger.info("=== EWS-LIG REPORT GENERATION COMPLETE ===")
        logger.info(f"Report saved as: {filename}")
        
        # Log summary
        logger.info("Report Summary:")
        logger.info(f"  Generated Months: {MONTHS}")
        logger.info(f"  Processing Months: {PROCESSING_MONTHS}")
        logger.info(f"  Tracker Year: {tracker_year}")
        
    except Exception as e:
        logger.error(f"Error in EWS-LIG report generation: {e}")
        raise

if __name__ == "__main__":
    main()
