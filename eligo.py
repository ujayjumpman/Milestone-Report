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

# ---------------------------------------------------------------------------
# CONFIG / CONSTANTS
# ---------------------------------------------------------------------------
load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

COS_API_KEY    = os.getenv("COS_API_KEY")
COS_CRN        = os.getenv("COS_SERVICE_INSTANCE_CRN")
COS_ENDPOINT   = os.getenv("COS_ENDPOINT")
BUCKET         = os.getenv("COS_BUCKET_NAME")
ELIGO_STRUCTURE_KEY = os.getenv("ELIGO_STRUCTURE_TRACKER_PATH")
ELIGO_TG_FINISHING_KEY = os.getenv("ELIGO_TG_TRACKER_PATH")
ELIGO_TH_FINISHING_KEY = os.getenv("ELIGO_TH_TRACKER_PATH")
ELIGO_KRA_KEY = os.getenv("KRA_PATH")

GREEN_HEX = "FF92D050"
MONTHS = ["June", "July", "August"]

ROWS_TO_BOLD = {1, 5, 12, 19}
TOWER_G_ANTICIPATED_COLS = ['N', 'R', 'V']
TOWER_H_ANTICIPATED_COLS = ['AB', 'AF', 'AJ', 'AN', 'AR', 'AV', 'AZ']

TOWER_G_ACTIVITIES = [
    "Water Proofing Works",
    "HVAC 2nd Fix",
    "Wall tiling (Toilet & Kitchen)",
    "Floor tiling"
]

TOWER_H_ACTIVITIES = [
    "HVAC first fix",
    "POP punning (Major area)",
    "Wall Tiling",
    "Floor Tiling"
]

# ---------------------------------------------------------------------------
# COS HELPERS
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# UTILITIES
# ---------------------------------------------------------------------------
def extract_number(cell_value):
    if not cell_value or cell_value == "-":
        return 0.0
    match = re.search(r"(\d+)", str(cell_value))
    return float(match.group(1)) if match else 0.0

def get_previous_months():
    now = datetime.now()
    current_month = now.month
    month_map = {"June": 6, "July": 7, "August": 8}
    return [m for m in MONTHS if month_map[m] < current_month]

def count_green_dates_in_month(wb, sheet_name, columns, year, month):
    """Count dates in green cells for a specific month and year from given columns, **all rows**"""
    if sheet_name not in wb.sheetnames:
        logger.warning(f"Sheet {sheet_name} not found in workbook")
        return 0
    sheet = wb[sheet_name]
    count = 0

    max_row = sheet.max_row
    for col_letter in columns:
        for row in range(4, max_row + 1):  # Excel data typically starts from row 4
            cell = sheet[f"{col_letter}{row}"]
            if cell.value:
                try:
                    cell_date = None
                    if isinstance(cell.value, datetime):
                        cell_date = cell.value
                    elif isinstance(cell.value, str):
                        cell_date = pd.to_datetime(cell.value, dayfirst=True, errors='coerce')
                    if pd.notna(cell_date) and cell_date.year == year and cell_date.month == month:
                        fill = cell.fill
                        color_code = getattr(fill, "start_color", None)
                        rgb = color_code.rgb if color_code else None
                        if fill.fill_type == "solid" and rgb == GREEN_HEX:
                            count += 1
                except Exception as e:
                    logger.debug(f"Error processing cell {col_letter}{row}: {e}")
                    continue
    return count

def count_completed_activities_by_month(wb, sheet_names, activity_name, year, month):
    count = 0
    for sheet_name in sheet_names:
        if sheet_name not in wb.sheetnames:
            continue
        try:
            sheet = wb[sheet_name]
            for row in sheet.iter_rows(min_row=2, max_row=100):
                if len(row) > 6:
                    activity_cell = row[6]
                    if activity_cell.value and str(activity_cell.value).strip().lower() == activity_name.lower():
                        if len(row) > 11:
                            finish_cell = row[11]
                            if finish_cell.value:
                                try:
                                    finish_date = None
                                    if isinstance(finish_cell.value, datetime):
                                        finish_date = finish_cell.value
                                    elif isinstance(finish_cell.value, str):
                                        finish_date = pd.to_datetime(finish_cell.value, dayfirst=True, errors='coerce')
                                    if pd.notna(finish_date) and finish_date.year == year and finish_date.month == month:
                                        count += 1
                                        logger.debug(f"Found completed {activity_name} in {sheet_name} on {finish_date}")
                                except Exception as e:
                                    logger.debug(f"Error processing finish date: {e}")
                                    continue
        except Exception as e:
            logger.warning(f"Error processing sheet {sheet_name}: {e}")
            continue
    return count

# ---------------------------------------------------------------------------
# TOWER G STRUCTURE
# ---------------------------------------------------------------------------
def get_tower_g_structure_targets():
    targets = {"June": 1, "July": 1, "August": 1}
    logger.info(f"Tower G Structure targets: {targets}")
    return targets

def count_tower_g_completed(cos):
    raw = download_file_bytes(cos, ELIGO_STRUCTURE_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    counts = {m: 0 for m in MONTHS}
    current_year = datetime.now().year
    month_map = {"June": 6, "July": 7, "August": 8}
    for month_name in MONTHS:
        month_num = month_map[month_name]
        count = count_green_dates_in_month(wb, "Revised Baselines- 25 days SC", TOWER_G_ANTICIPATED_COLS, current_year, month_num)
        counts[month_name] = count
    logger.info(f"Tower G completed pours by month: {counts}")
    return counts

def build_tower_g_structure_dataframe(targets, completed):
    total_milestones = 1
    weightage = round(100 / total_milestones, 2) if total_milestones else 0

    # Cumulative completed and targets
    cum_completed = {}
    cum_targets = {}
    running_done = 0
    running_target = 0
    for m in MONTHS:
        running_done += completed.get(m, 0)
        running_target += targets.get(m, 0)
        cum_completed[m] = running_done
        cum_targets[m] = running_target

    def pct(m):
        t = cum_targets[m]
        d = cum_completed[m]
        if t == 0:
            return "0.0%"
        val = min(round((d / t) * 100, 2), 100)
        return f"{val}%"

    target_text = f"{int(sum(targets.values()))} Pours ({int(targets['June'])} Pours-June, {int(targets['July'])} Pours-July & {int(targets['August'])} Pours-August)"

    row = {
        "Milestone": "Milestone-01",
        "Activity": "Pour Casting",
        "Target Till August": target_text,
        "% Work Done against Target-Till June": pct("June"),
        "% Work Done against Target-Till July": pct("July"),
        "% Work Done against Target-Till August": pct("August"),
        "Weightage": weightage,
        "Weighted Delay against Targets": "",  # Filled below
        "Target achieved in June": f"{completed.get('June', 0)} pour cast out of {int(targets['June'])} planned",
        "Target achieved in July": f"{completed.get('July', 0)} pour cast out of {int(targets['July'])} planned",
        "Target achieved in August": f"{completed.get('August', 0)} pour cast out of {int(targets['August'])} planned",
        "Total achieved": "",
        "Delay Reasons_June 2025": "",
    }

    # Weighted Delay: Use last month for which there's a target
    last_month = "August"
    for m in reversed(MONTHS):
        if cum_targets[m] > 0:
            last_month = m
            break
    try:
        last_pct = float(pct(last_month).replace("%", ""))
        row["Weighted Delay against Targets"] = f"{round((last_pct * weightage) / 100, 2)}%"
    except Exception:
        row["Weighted Delay against Targets"] = ""

    df = pd.DataFrame([row])
    return df

# ---------------------------------------------------------------------------
# TOWER H STRUCTURE
# ---------------------------------------------------------------------------
def get_tower_h_structure_targets():
    targets = {"June": 3, "July": 3, "August": 4}
    logger.info(f"Tower H Structure targets: {targets}")
    return targets

def count_tower_h_completed(cos):
    raw = download_file_bytes(cos, ELIGO_STRUCTURE_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    counts = {m: 0 for m in MONTHS}
    current_year = datetime.now().year
    month_map = {"June": 6, "July": 7, "August": 8}
    for month_name in MONTHS:
        month_num = month_map[month_name]
        count = count_green_dates_in_month(wb, "ELIGO SLAB CYCLE", TOWER_H_ANTICIPATED_COLS, current_year, month_num)
        counts[month_name] = count
    logger.info(f"Tower H completed pours by month: {counts}")
    return counts

def build_tower_h_structure_dataframe(targets, completed):
    total_milestones = 1
    weightage = round(100 / total_milestones, 2) if total_milestones else 0

    # Cumulative completed and targets
    cum_completed = {}
    cum_targets = {}
    running_done = 0
    running_target = 0
    for m in MONTHS:
        running_done += completed.get(m, 0)
        running_target += targets.get(m, 0)
        cum_completed[m] = running_done
        cum_targets[m] = running_target

    def pct(m):
        t = cum_targets[m]
        d = cum_completed[m]
        if t == 0:
            return "0.0%"
        val = min(round((d / t) * 100, 2), 100)
        return f"{val}%"

    target_text = f"{int(sum(targets.values()))} Pours ({int(targets['June'])} Pours-June, {int(targets['July'])} Pours-July & {int(targets['August'])} Pours-August)"

    row = {
        "Milestone": "Milestone-01",
        "Activity": "Pour Casting",
        "Target Till August": target_text,
        "% Work Done against Target-Till June": pct("June"),
        "% Work Done against Target-Till July": pct("July"),
        "% Work Done against Target-Till August": pct("August"),
        "Weightage": weightage,
        "Weighted Delay against Targets": "",
        "Target achieved in June": f"{completed.get('June', 0)} pour cast out of {int(targets['June'])} planned",
        "Target achieved in July": f"{completed.get('July', 0)} pour cast out of {int(targets['July'])} planned",
        "Target achieved in August": f"{completed.get('August', 0)} pour cast out of {int(targets['August'])} planned",
        "Total achieved": "",
        "Delay Reasons_June 2025": "",
    }

    # Weighted Delay: Use last month for which there's a target
    last_month = "August"
    for m in reversed(MONTHS):
        if cum_targets[m] > 0:
            last_month = m
            break
    try:
        last_pct = float(pct(last_month).replace("%", ""))
        row["Weighted Delay against Targets"] = f"{round((last_pct * weightage) / 100, 2)}%"
    except Exception:
        row["Weighted Delay against Targets"] = ""

    df = pd.DataFrame([row])
    return df

# ---------------------------------------------------------------------------
# TOWER G & H FINISHING (Unchanged from your previous code)
# ---------------------------------------------------------------------------
def get_tower_g_finishing_targets():
    targets = {
        "Water Proofing Works": {"June": 20, "July": 24, "August": 19},
        "HVAC 2nd Fix": {"June": 41, "July": 16, "August": 0},
        "Wall tiling (Toilet & Kitchen)": {"June": 0, "July": 1, "August": 43},
        "Floor tiling": {"June": 0, "July": 0, "August": 32}
    }
    logger.info(f"Tower G Finishing targets: {targets}")
    return targets

def count_tower_g_finishing_completed(cos):
    raw = download_file_bytes(cos, ELIGO_TG_FINISHING_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    target_sheets = ['Common Area', 'Pour G1', 'Pour G2', 'Pour G3']
    counts = {}
    current_year = datetime.now().year
    month_map = {"June": 6, "July": 7, "August": 8}
    for activity in TOWER_G_ACTIVITIES:
        counts[activity] = {}
        for month_name in MONTHS:
            month_num = month_map[month_name]
            count = count_completed_activities_by_month(wb, target_sheets, activity, current_year, month_num)
            counts[activity][month_name] = count
    logger.info(f"Tower G Finishing completed by month: {counts}")
    return counts

def build_tower_g_finishing_dataframe(targets, completed):
    prev_months = get_previous_months()
    month_indices = {m: i for i, m in enumerate(MONTHS)}
    progress_data = []
    total_milestones = len(TOWER_G_ACTIVITIES)
    weightage = round(100 / total_milestones, 2) if total_milestones else 0
    for i, activity in enumerate(TOWER_G_ACTIVITIES):
        row = {
            "Milestone": f"Milestone-{i+1:02d}",
            "Activity": activity,
            "Weightage": weightage,
            "Weighted Delay against Targets": "",
            "Total achieved": "",
            "Delay Reasons_June 2025": "",
        }
        for m in MONTHS:
            if m in prev_months:
                months_to_count = MONTHS[:month_indices[m] + 1]
                count_cumulative = sum(completed[activity][month] for month in months_to_count)
                target_cumulative = sum(targets[activity][month] for month in months_to_count)
                if target_cumulative == 0:
                    pct_done = 100.0
                else:
                    pct_done = min(round((count_cumulative / target_cumulative) * 100, 2), 100)
                row[f"% Work Done against Target-Till {m}"] = f"{pct_done}%"
                month_target = targets[activity][m]
                count_in_month = completed[activity][m]
                if month_target == 0:
                    future_months = []
                    for future_m in MONTHS[month_indices[m] + 1:]:
                        if targets[activity][future_m] > 0:
                            future_months.append(future_m)
                    if future_months:
                        if len(future_months) == 1:
                            row[f"Target achieved in {m}"] = f"Planned for {future_months[0]}"
                        else:
                            row[f"Target achieved in {m}"] = f"Planned for {' and '.join(future_months)}"
                    else:
                        row[f"Target achieved in {m}"] = f"{count_in_month} Flats out of {int(month_target)} planned"
                else:
                    row[f"Target achieved in {m}"] = f"{count_in_month} Flats out of {int(month_target)} planned"
            else:
                row[f"% Work Done against Target-Till {m}"] = ""
                row[f"Target achieved in {m}"] = ""
        if "June" in prev_months:
            pct_june = row.get("% Work Done against Target-Till June", "0%").replace("%", "")
            try:
                pct_june_val = float(pct_june)
                row["Weighted Delay against Targets"] = f"{round((pct_june_val * weightage) / 100, 2)}%"
            except ValueError:
                row["Weighted Delay against Targets"] = ""
        total_target = sum(targets[activity][month] for month in MONTHS)
        row["Target Till August"] = (
            f"{int(total_target)} Flats ({int(targets[activity]['June'])} Flats-June, "
            f"{int(targets[activity]['July'])} Flats-July & {int(targets[activity]['August'])} Flats-August)"
        )
        progress_data.append(row)
    all_cols = ["Milestone", "Activity", "Target Till August",
                "% Work Done against Target-Till June",
                "% Work Done against Target-Till July",
                "% Work Done against Target-Till August", 
                "Weightage", "Weighted Delay against Targets",
                "Target achieved in June", "Target achieved in July", "Target achieved in August",
                "Total achieved", "Delay Reasons_June 2025"]
    df_tg_finishing = pd.DataFrame(progress_data, columns=all_cols)
    return df_tg_finishing

def get_tower_h_finishing_targets():
    targets = {
        "HVAC first fix": {"June": 16, "July": 0, "August": 0},
        "POP punning (Major area)": {"June": 13, "July": 8, "August": 8},
        "Wall Tiling": {"June": 8, "July": 39, "August": 9},
        "Floor Tiling": {"June": 14, "July": 39, "August": 9}
    }
    logger.info(f"Tower H Finishing targets: {targets}")
    return targets

def count_tower_h_finishing_completed(cos):
    raw = download_file_bytes(cos, ELIGO_TH_FINISHING_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    target_sheets = ['Common Area', 'Pre-Construction Activities', 'Pour H1', 'Pour H2', 
                    'Pour H3', 'Pour H4', 'Pour H5', 'Pour H6', 'Pour H7']
    counts = {}
    current_year = datetime.now().year
    month_map = {"June": 6, "July": 7, "August": 8}
    for activity in TOWER_H_ACTIVITIES:
        counts[activity] = {}
        for month_name in MONTHS:
            month_num = month_map[month_name]
            count = count_completed_activities_by_month(wb, target_sheets, activity, current_year, month_num)
            counts[activity][month_name] = count
    logger.info(f"Tower H Finishing completed by month: {counts}")
    return counts

def build_tower_h_finishing_dataframe(targets, completed):
    prev_months = get_previous_months()
    month_indices = {m: i for i, m in enumerate(MONTHS)}
    progress_data = []
    total_milestones = len(TOWER_H_ACTIVITIES)
    weightage = round(100 / total_milestones, 2) if total_milestones else 0
    for i, activity in enumerate(TOWER_H_ACTIVITIES):
        row = {
            "Milestone": f"Milestone-{i+1:02d}",
            "Activity": activity,
            "Weightage": weightage,
            "Weighted Delay against Targets": "",
            "Total achieved": "",
            "Delay Reasons_June 2025": "",
        }
        for m in MONTHS:
            if m in prev_months:
                months_to_count = MONTHS[:month_indices[m] + 1]
                count_cumulative = sum(completed[activity][month] for month in months_to_count)
                target_cumulative = sum(targets[activity][month] for month in months_to_count)
                if target_cumulative == 0:
                    pct_done = 100.0
                else:
                    pct_done = min(round((count_cumulative / target_cumulative) * 100, 2), 100)
                row[f"% Work Done against Target-Till {m}"] = f"{pct_done}%"
                month_target = targets[activity][m]
                count_in_month = completed[activity][m]
                if month_target == 0:
                    future_months = []
                    for future_m in MONTHS[month_indices[m] + 1:]:
                        if targets[activity][future_m] > 0:
                            future_months.append(future_m)
                    if future_months:
                        if len(future_months) == 1:
                            row[f"Target achieved in {m}"] = f"Planned for {future_months[0]}"
                        else:
                            row[f"Target achieved in {m}"] = f"Planned for {' and '.join(future_months)}"
                    else:
                        row[f"Target achieved in {m}"] = f"{count_in_month} Flats out of {int(month_target)} planned"
                else:
                    row[f"Target achieved in {m}"] = f"{count_in_month} Flats out of {int(month_target)} planned"
            else:
                row[f"% Work Done against Target-Till {m}"] = ""
                row[f"Target achieved in {m}"] = ""
        if "June" in prev_months:
            pct_june = row.get("% Work Done against Target-Till June", "0%").replace("%", "")
            try:
                pct_june_val = float(pct_june)
                row["Weighted Delay against Targets"] = f"{round((pct_june_val * weightage) / 100, 2)}%"
            except ValueError:
                row["Weighted Delay against Targets"] = ""
        total_target = sum(targets[activity][month] for month in MONTHS)
        row["Target Till August"] = (
            f"{int(total_target)} Flats ({int(targets[activity]['June'])} Flats-June, "
            f"{int(targets[activity]['July'])} Flats-July & {int(targets[activity]['August'])} Flats-August)"
        )
        progress_data.append(row)
    all_cols = ["Milestone", "Activity", "Target Till August",
                "% Work Done against Target-Till June",
                "% Work Done against Target-Till July",
                "% Work Done against Target-Till August",
                "Weightage", "Weighted Delay against Targets",
                "Target achieved in June", "Target achieved in July", "Target achieved in August",
                "Total achieved", "Delay Reasons_June 2025"]
    df_th_finishing = pd.DataFrame(progress_data, columns=all_cols)
    return df_th_finishing

# ---------------------------------------------------------------------------
# WRITER / STYLING - UPDATED WITH DATE DISPLAY
# ---------------------------------------------------------------------------
def write_excel_report(df_tg_structure, df_th_structure, df_tg_finishing, df_th_finishing, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Eligo Time Delivery Milestones"
    
    # Add title and date at the top
    current_date = datetime.now().strftime("%d-%m-%Y")
    ws.append(["Eligo Time Delivery Milestones"])
    ws.append([f"Report Generated on: {current_date}"])
    ws.append([])  # Empty row for spacing
    
    # Define styles
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    grey = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    bold_font = Font(bold=True)
    normal_font = Font(bold=False)
    title_font = Font(bold=True, size=14)
    date_font = Font(bold=False, size=10, color="666666")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    
    # Get max columns for merging
    max_cols = max(len(df_tg_structure.columns), len(df_th_structure.columns), 
                   len(df_tg_finishing.columns), len(df_th_finishing.columns))
    
    # Style title row (row 1)
    ws.merge_cells(f'A1:{get_column_letter(max_cols)}1')
    ws['A1'].font = title_font
    ws['A1'].alignment = center_align
    ws['A1'].fill = grey
    
    # Style date row (row 2)
    ws.merge_cells(f'A2:{get_column_letter(max_cols)}2')
    ws['A2'].font = date_font
    ws['A2'].alignment = center_align
    
    def append_df_block(title, df, total_delay_label):
        start_col = 1
        end_col = len(df.columns)
        
        # Section title row
        ws.append([title])
        title_row = ws.max_row
        ws.merge_cells(start_row=title_row, start_column=start_col,
                       end_row=title_row, end_column=end_col)
        for cell in ws[title_row]:
            cell.fill = grey
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
                cell.font = bold_font if r in ROWS_TO_BOLD else normal_font
                cell.alignment = left_align if cell.col_idx in (1, 2) else center_align
                cell.border = border
                
        # Total delay row
        try:
            total_delay = sum(float(str(v).strip('%')) for v in df["Weighted Delay against Targets"] if v)
        except Exception:
            total_delay = 0
        weighted_delay_col_idx = None
        for idx, col_name in enumerate(df.columns, start=1):
            if col_name == "Weighted Delay against Targets":
                weighted_delay_col_idx = idx
                break
        total_row_data = [""] * end_col
        if weighted_delay_col_idx:
            total_row_data[weighted_delay_col_idx - 1] = f"{round(total_delay, 2)}%"
            total_row_data[0] = total_delay_label
        ws.append(total_row_data)
        delay_row = ws.max_row
        for idx, cell in enumerate(ws[delay_row], start=1):
            cell.font = bold_font
            cell.fill = yellow
            if idx == 1:
                cell.alignment = left_align
            elif idx == weighted_delay_col_idx:
                cell.alignment = center_align
            else:
                cell.alignment = center_align
            cell.border = border
        return title_row, delay_row
        
    # Write all sections (after title, date, and empty row)
    append_df_block("Tower G Structure Progress Against Milestones", df_tg_structure, "Total Delay Tower G Structure")
    append_df_block("Tower H Structure Progress Against Milestones", df_th_structure, "Total Delay Tower H Structure")
    append_df_block("Tower G Finishing Progress Against Milestones", df_tg_finishing, "Total Delay Tower G Finishing")
    append_df_block("Tower H Finishing Progress Against Milestones", df_th_finishing, "Total Delay Tower H Finishing")
    
    # Column widths
    for col in ws.columns:
        max_len = 0
        for cell in col:
            text = str(cell.value) if cell.value is not None else ""
            max_len = max(max_len, len(text.split("\n")[0]))
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)
    
    # Row heights
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 22
    
    wb.save(filename)
    logger.info(f"Eligo report saved to {filename}")

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    cos = init_cos()
    logger.info("Processing Tower G Structure milestones...")
    targets_tg_structure = get_tower_g_structure_targets()
    completed_tg_structure = count_tower_g_completed(cos)
    df_tg_structure = build_tower_g_structure_dataframe(targets_tg_structure, completed_tg_structure)
    logger.info("Processing Tower H Structure milestones...")
    targets_th_structure = get_tower_h_structure_targets()
    completed_th_structure = count_tower_h_completed(cos)
    df_th_structure = build_tower_h_structure_dataframe(targets_th_structure, completed_th_structure)
    logger.info("Processing Tower G Finishing milestones...")
    targets_tg_finishing = get_tower_g_finishing_targets()
    completed_tg_finishing = count_tower_g_finishing_completed(cos)
    df_tg_finishing = build_tower_g_finishing_dataframe(targets_tg_finishing, completed_tg_finishing)
    logger.info("Processing Tower H Finishing milestones...")
    targets_th_finishing = get_tower_h_finishing_targets()
    completed_th_finishing = count_tower_h_finishing_completed(cos)
    df_th_finishing = build_tower_h_finishing_dataframe(targets_th_finishing, completed_th_finishing)
    filename = f"Eligo_Time_Delivery_Milestone_Report ({datetime.now():%Y-%m-%d}).xlsx"
    logger.info("Writing Eligo Excel report...")
    write_excel_report(df_tg_structure, df_th_structure, df_tg_finishing, df_th_finishing, filename)
    logger.info("Eligo milestone report generation completed successfully!")
    return df_tg_structure, df_th_structure, df_tg_finishing, df_th_finishing

if __name__ == "__main__":
    main()
