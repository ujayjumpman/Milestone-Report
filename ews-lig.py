import os
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
EWS_LIG_STRUCTURE_KEY = os.getenv("EWS_LIG_STRUCTURE_TRACKER_PATH")
EWS_LIG_KRA_KEY = os.getenv("KRA_FILE_PATH")

MONTHS = ["June", "July", "August"]
MONTH_TO_NUM = {"June": 6, "July": 7, "August": 8}

KRA_SHEET = "EW-LI P4 Targets Till August "
TOWER1_TARGETS_CELLS = {'June': 'B4', 'July': 'C4', 'August': 'D4'}
TOWER3_TARGETS_CELLS = {'June': 'B12', 'July': 'C12', 'August': 'D12'}
TOWER2_TARGETS_CELLS = {'June': 'B19', 'July': 'C19', 'August': 'D19'}

TRACKER_SHEET = "Revised Baseline 45daysNGT+Rai"

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

def get_previous_months():
    # Modified to return only June for display purposes
    return ["June"]

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
    # Only show results for June, but keep all targets for "Target Till August"
    prev_months = get_previous_months()  # This will return only ["June"]
    weightage = 100
    
    # Calculate cumulative targets (still use all months for "Target Till August")
    cum_targets = {}
    cum_completed = {}
    for i, m in enumerate(MONTHS):
        months_to_count = MONTHS[:i+1]
        cum_targets[m] = sum(targets[mm] for mm in months_to_count)
        cum_completed[m] = sum(completed[mm] for mm in months_to_count if mm in prev_months)

    def pct(m):
        # Only show percentage for June
        if m != "June":
            return ""  # Leave July and August blank
        t = cum_targets[m]
        d = cum_completed[m]
        if t == 0: return "0.0%"
        val = min(round((d / t) * 100, 2), 100)
        return f"{val}%"

    row = {
        "Milestone": f"{tower_name} Structure",
        "Target Till August": f"{sum(targets.values())} Pours ({targets['June']} June, {targets['July']} July, {targets['August']} August)",
        "% Work Done against Target-Till June": pct("June"),
        "% Work Done against Target-Till July": "",  # Blank
        "% Work Done against Target-Till August": "",  # Blank
        "Weightage": weightage,
        "Weighted Delay against Targets": "",
        "Target achieved in June": f"{completed.get('June', 0)} out of {targets.get('June', 0)}",
        "Target achieved in July": "",  # Blank
        "Target achieved in August": "",  # Blank
        "Total achieved": f"{completed.get('June', 0)} out of {sum(targets.values())}",  # Only June achieved vs total target
        "Delay Reasons": "",
    }
    
    # Calculate weighted delay only for June
    june_pct_str = pct("June")
    if june_pct_str:
        try:
            june_pct = float(june_pct_str.replace("%", ""))
            row["Weighted Delay against Targets"] = f"{round((june_pct * weightage) / 100, 2)}%"
        except Exception:
            row["Weighted Delay against Targets"] = "0.0%"
    
    df = pd.DataFrame([row])
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
    cos = init_cos()
    kra_raw = download_file_bytes(cos, EWS_LIG_KRA_KEY)
    kra_wb = load_workbook(filename=BytesIO(kra_raw), data_only=True)
    tracker_raw = download_file_bytes(cos, EWS_LIG_STRUCTURE_KEY)
    tracker_wb = load_workbook(filename=BytesIO(tracker_raw), data_only=True)
    sheet = tracker_wb[TRACKER_SHEET]

    prev_months = get_previous_months()
    tracker_year = detect_tracker_year(sheet, TOWER1_POUR_COLS, TOWER1_ROW_START, TOWER1_ROW_END)

    # Tower 1
    targets_t1 = get_targets_from_kra(kra_wb, KRA_SHEET, TOWER1_TARGETS_CELLS)
    completed_t1 = count_pours(sheet, TOWER1_POUR_COLS, TOWER1_ROW_START, TOWER1_ROW_END, MONTHS, tracker_year)
    df_t1 = build_structure_dataframe("Tower 1", targets_t1, completed_t1)

    # Tower 3
    targets_t3 = get_targets_from_kra(kra_wb, KRA_SHEET, TOWER3_TARGETS_CELLS)
    completed_t3 = count_pours(sheet, TOWER3_POUR_COLS, TOWER3_ROW_START, TOWER3_ROW_END, MONTHS, tracker_year)
    df_t3 = build_structure_dataframe("Tower 3", targets_t3, completed_t3)

    # Tower 2
    targets_t2 = get_targets_from_kra(kra_wb, KRA_SHEET, TOWER2_TARGETS_CELLS)
    completed_t2 = count_pours(sheet, TOWER2_POUR_COLS, TOWER2_ROW_START, TOWER2_ROW_END, MONTHS, tracker_year)
    df_t2 = build_structure_dataframe("Tower 2", targets_t2, completed_t2)

    filename = f"EWS_LIG_Milestone_Report ({datetime.now():%Y-%m-%d}).xlsx"
    dfs = [
        ("Tower 1 Structure Progress Against Milestones", df_t1, "Total Delay Tower 1 Structure"),
        ("Tower 3 Structure Progress Against Milestones", df_t3, "Total Delay Tower 3 Structure"),
        ("Tower 2 Structure Progress Against Milestones", df_t2, "Total Delay Tower 2 Structure"),
    ]
    write_excel_report(dfs, filename)
    logger.info("EWS-LIG milestone report generation completed successfully!")

if __name__ == "__main__":
    main()
