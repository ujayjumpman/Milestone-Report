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
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

COS_API_KEY    = os.getenv("COS_API_KEY")
COS_CRN        = os.getenv("COS_SERVICE_INSTANCE_CRN")
COS_ENDPOINT   = os.getenv("COS_ENDPOINT")
BUCKET         = os.getenv("COS_BUCKET_NAME")
KRA_KEY        = os.getenv("KRA_FILE_PATH")
T6_TRACKER_KEY = os.getenv("T6_TRACKER_PATH")
T5_TRACKER_KEY = os.getenv("T5_TRACKER_PATH")
T7_TRACKER_KEY = os.getenv("T7_TRACKER_PATH")
GREEN3_TRACKER_KEY = os.getenv("G3_TRACKER_PATH")

# Green highlight RGB from tracker (Tower 6) used to count completed slabs
GREEN_HEX = "FF92D050"

MONTHS = ["June", "July", "August"]

# Only these absolute row numbers in the final sheet should be bold (section title rows only)
ROWS_TO_BOLD = {1, 5, 12, 19}

# Tower 6 slab parsing config
TOWER6_ROWS = [4, 5, 6, 7, 9, 10, 14, 15, 16, 17, 19, 20]
TOWER6_COLS = ['FK', 'FM', 'FO', 'FQ', 'FS', 'FU', 'FW', 'FY', 'GA', 'GC', 'GE', 'GG', 'GI', 'GK']

# Tower 5 & 7 fixed KRA cells
T5_TARGET_CELLS = {
    "Installation of Rear & Front balcony UPVC Windows": {"June": ("D23", "Flats"), "July": ("E23", "Flats"), "August": ("F23", "Flats")},
    "EL-Second Fix": {"June": ("D24", "Flats"), "July": ("E24", "Flats"), "August": ("F24", "Flats")},
    "Gypsum board false ceiling": {"June": ("D25", "Flats"), "July": ("E25", "Flats"), "August": ("F25", "Flats")},
    "Paint 1st Coat": {"June": ("D26", "Modules"), "July": ("E26", "Modules"), "August": ("F26", "Modules")},
}

T7_TARGET_CELLS = {
    "El- First Fix": {"June": ("D30", "Flats"), "July": ("E30", "Flats"), "August": ("F30", "Flats")},
    "Floor Tiling": {"June": ("D31", "Flats"), "July": ("E31", "Flats"), "August": ("F31", "Flats")},
    "False Ceiling Framing": {"June": ("D32", "Flats"), "July": ("E32", "Flats"), "August": ("F32", "Flats")},
    "C-Stone flooring": {"June": ("D33", "Modules"), "July": ("E33", "Modules"), "August": ("F33", "Modules")},
}

T5_ACTIVITIES = list(T5_TARGET_CELLS.keys())
T7_ACTIVITIES = list(T7_TARGET_CELLS.keys())

# -----------------------------------------------------------------------------
# COS HELPERS
# -----------------------------------------------------------------------------

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

# -----------------------------------------------------------------------------
# UTILITIES
# -----------------------------------------------------------------------------

def extract_number(cell_value):
    """Return first integer found; treat '-' or empty as 0."""
    if not cell_value or cell_value == "-":
        return 0.0
    match = re.search(r"(\d+)", str(cell_value))
    return float(match.group(1)) if match else 0.0

def get_previous_months():
    now = datetime.now()
    current_month = now.month
    month_map = {"June": 6, "July": 7, "August": 8}
    return [m for m in MONTHS if month_map[m] < current_month]

# -----------------------------------------------------------------------------
# TOWER 6
# -----------------------------------------------------------------------------

def get_slab_targets_fixed_cells(cos):
    raw = download_file_bytes(cos, KRA_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    sheet = wb["VeridiaTargets Till August 2025"]
    targets = {
        "June": extract_number(sheet["B18"].value),
        "July": extract_number(sheet["C18"].value),
        "August": extract_number(sheet["D18"].value),
    }
    logger.info(f"Slab targets (T6): {targets}")
    return targets

def count_tower6_completed(wb):
    counts = {m: 0 for m in MONTHS}
    sheet = wb["Revised baseline with 60d NGT"]
    for row in TOWER6_ROWS:
        for col in TOWER6_COLS:
            cell = sheet[f"{col}{row}"]
            val = cell.value
            cell_date = None
            if isinstance(val, datetime):
                cell_date = val
            elif isinstance(val, str):
                try:
                    cell_date = datetime.strptime(val, "%Y-%m-%d")
                except Exception:
                    continue
            if cell_date:
                month_name = cell_date.strftime("%B")
                if month_name in MONTHS:
                    fill = cell.fill
                    if fill.fill_type == "solid" and fill.start_color:
                        if fill.start_color.rgb == GREEN_HEX:
                            counts[month_name] += 1
    logger.info(f"Completed slabs by month (T6): {counts}")
    return counts

def build_t6_milestone_dataframe(targets, completed):
    """Build single-row DF for Tower 6 with correct weighted value.
    Weighted Delay against Targets = (% work done till last reported month * Weightage)/100
    """
    prev_months = get_previous_months()
    cum_done = 0
    cum_target = 0
    month_pct = {}

    total_milestones = 1
    weightage = round(100 / total_milestones, 2) if total_milestones else 0

    rows_tmp = []
    for m in MONTHS:
        done   = completed.get(m, 0) if m in prev_months else 0
        target = targets.get(m, 0)
        cum_done   += done
        cum_target += target
        pct_done = round((cum_done / cum_target) * 100, 2) if cum_target > 0 else 0
        pct_done = min(pct_done, 100)
        month_pct[m] = pct_done if m in prev_months else 0

        rows_tmp.append({
            "Milestone": "Milestone-01",
            "Activity": "Slab Casting",
            "Target Till August": f"{int(sum(targets.values()))} Slabs ({int(targets['June'])} Slabs-June, {int(targets['July'])} slabs-July & {int(targets['August'])} slabs-August)",
            f"% Work Done against Target-Till {m}": f"{pct_done}%" if m in prev_months else "",
            "Weightage": weightage,
            # placeholder; we'll fill it in the final DF
            "Weighted Delay against Targets": "",
            f"Target achieved in {m}": f"{done} slab cast out of {int(target)} planned" if target > 0 and m in prev_months else "",
            "Total achieved": "",
            "Delay Reasons_June 2025": "",
        })

    # decide which month to use for weighted delay (last month we have data for)
    if prev_months:
        last_m = prev_months[-1]
        weighted_delay_val = round((month_pct[last_m] * weightage) / 100, 2)
    else:
        weighted_delay_val = None

    all_cols = ["Milestone", "Activity", "Target Till August",
                "% Work Done against Target-Till June",
                "% Work Done against Target-Till July",
                "% Work Done against Target-Till August",
                "Weightage", "Weighted Delay against Targets",
                "Target achieved in June", "Target achieved in July", "Target achieved in August",
                "Total achieved", "Delay Reasons_June 2025"]

    for col in all_cols:
        for r in rows_tmp:
            r.setdefault(col, "")

    final_df = pd.DataFrame(columns=all_cols)
    final_df.loc[0] = {
        "Milestone": "Milestone-01",
        "Activity": "Slab Casting",
        "Target Till August": rows_tmp[0]["Target Till August"],
        "% Work Done against Target-Till June": rows_tmp[0]["% Work Done against Target-Till June"],
        "% Work Done against Target-Till July": rows_tmp[1]["% Work Done against Target-Till July"],
        "% Work Done against Target-Till August": rows_tmp[2]["% Work Done against Target-Till August"],
        "Weightage": weightage,
        "Weighted Delay against Targets": f"{weighted_delay_val}%" if weighted_delay_val is not None else "",
        "Target achieved in June": rows_tmp[0]["Target achieved in June"],
        "Target achieved in July": rows_tmp[1]["Target achieved in July"],
        "Target achieved in August": rows_tmp[2]["Target achieved in August"],
        "Total achieved": "",
        "Delay Reasons_June 2025": "",
    }
    return final_df

# -----------------------------------------------------------------------------
# TOWER 5 - UPDATED FOR MODULE-WISE COUNTING
# -----------------------------------------------------------------------------

def count_completed_activities_by_module_and_month(wb, sheet_name, activity_mapping):
    """
    Count completed activities module-wise from the tracker workbook.
    Returns a dictionary with structure: {activity: {month: count}}
    """
    sheet = wb[sheet_name]
    activity_counts = {}
    
    # Initialize counts for all activities and months
    for activity in activity_mapping.keys():
        activity_counts[activity] = {month: 0 for month in MONTHS}
    
    # Find the header row to locate the "Actual Finish" column
    actual_finish_col = None
    for row in sheet.iter_rows(min_row=1, max_row=10):  # Check first 10 rows for headers
        for cell in row:
            if cell.value and "Actual Finish" in str(cell.value):
                actual_finish_col = cell.column
                break
        if actual_finish_col:
            break
    
    if not actual_finish_col:
        logger.warning(f"Could not find 'Actual Finish' column in {sheet_name}")
        return activity_counts
    
    logger.info(f"Found 'Actual Finish' column at column {actual_finish_col} in {sheet_name}")
    
    # Iterate through all data rows
    for row in sheet.iter_rows(min_row=2):  # Start from row 2 to skip headers
        try:
            # Get the activity name (assuming it's in column F, index 5)
            activity_cell = row[5] if len(row) > 5 else None
            if not activity_cell or not activity_cell.value:
                continue
                
            activity_name = str(activity_cell.value).strip()
            
            # Map the activity name to our standard names
            mapped_activity = None
            for standard_name, variations in activity_mapping.items():
                if activity_name in variations or activity_name.lower() in [v.lower() for v in variations]:
                    mapped_activity = standard_name
                    break
            
            if not mapped_activity:
                continue
            
            # Get the actual finish date (column L is index 11, but we found the actual column)
            actual_finish_cell = row[actual_finish_col - 1] if len(row) >= actual_finish_col else None
            if not actual_finish_cell or not actual_finish_cell.value:
                continue
            
            # Parse the date
            actual_finish_date = None
            if isinstance(actual_finish_cell.value, datetime):
                actual_finish_date = actual_finish_cell.value
            elif isinstance(actual_finish_cell.value, str):
                try:
                    # Try different date formats
                    for date_format in ["%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y"]:
                        try:
                            actual_finish_date = datetime.strptime(actual_finish_cell.value, date_format)
                            break
                        except ValueError:
                            continue
                except Exception:
                    continue
            
            if actual_finish_date:
                month_name = actual_finish_date.strftime("%B")
                if month_name in MONTHS:
                    activity_counts[mapped_activity][month_name] += 1
                    logger.debug(f"Counted {mapped_activity} completed in {month_name}")
                    
        except Exception as e:
            logger.debug(f"Error processing row in {sheet_name}: {e}")
            continue
    
    logger.info(f"Activity counts for {sheet_name}: {activity_counts}")
    return activity_counts

def get_t5_targets_and_progress(cos):
    raw = download_file_bytes(cos, KRA_KEY)
    wb_kra = load_workbook(filename=BytesIO(raw), data_only=True)
    sheet_kra = wb_kra["VeridiaTargets Till August 2025"]

    t5_targets = {}
    for activity in T5_ACTIVITIES:
        t5_targets[activity] = {}
        for month in MONTHS:
            cell, unit = T5_TARGET_CELLS[activity][month]
            val = extract_number(sheet_kra[cell].value)
            t5_targets[activity][month] = (val, unit)
    logger.info(f"Tower 5 targets: {t5_targets}")

    raw_tracker = download_file_bytes(cos, T5_TRACKER_KEY)
    wb_tracker = load_workbook(filename=BytesIO(raw_tracker), data_only=True)

    # Define activity name mapping for Tower 5
    t5_activity_mapping = {
        "Installation of Rear & Front balcony UPVC Windows": [
            "Installation of Rear & Front balcony UPVC Windows",
            "UPVC Windows",
            "Balcony UPVC Windows"
        ],
        "EL-Second Fix": [
            "EL-Second Fix",
            "EL Second Fix",
            "Electrical Second Fix",
            "EL- Second Fix"
        ],
        "Gypsum board false ceiling": [
            "Gypsum board false ceiling",
            "Gypsum False Ceiling",
            "False Ceiling Gypsum"
        ],
        "Paint 1st Coat": [
            "Paint 1st Coat",
            "Painting First Coat",
            "Paint First Coat",
            "1st Coat Paint"
        ]
    }

    # Count completed activities module-wise
    activity_counts = count_completed_activities_by_module_and_month(
        wb_tracker, "TOWER 5 FINISHING.", t5_activity_mapping
    )

    prev_months = get_previous_months()
    month_indices = {m: i for i, m in enumerate(MONTHS)}

    progress_data = []
    total_milestones = len(T5_ACTIVITIES)
    weightage = round(100 / total_milestones, 2) if total_milestones else 0

    for i, activity in enumerate(T5_ACTIVITIES):
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
                months_to_count = MONTHS[: month_indices[m] + 1]
                count_cumulative = sum(activity_counts[activity][month] for month in months_to_count)

                target_cumulative, unit = 0, ""
                for month in months_to_count:
                    val, u = t5_targets[activity][month]
                    target_cumulative += val
                    unit = u

                # if target is '-' (0) treat as 100%
                if target_cumulative == 0:
                    pct_done = 100.0
                else:
                    pct_done = min(round((count_cumulative / target_cumulative) * 100, 2), 100)

                row[f"% Work Done against Target-Till {m}"] = f"{pct_done}%"
                
                # Get individual month target for display
                month_target, month_unit = t5_targets[activity][m]
                # For achieved count, we want to show only what was achieved in that specific month
                count_in_month = activity_counts[activity][m]
                
                if month_target == 0:
                    # Show which future months it's planned for
                    future_months = []
                    for future_m in MONTHS[month_indices[m] + 1:]:
                        future_target, _ = t5_targets[activity][future_m]
                        if future_target > 0:
                            future_months.append(future_m)
                    
                    if future_months:
                        if len(future_months) == 1:
                            row[f"Target achieved in {m}"] = f"Planned for {future_months[0]}"
                        else:
                            row[f"Target achieved in {m}"] = f"Planned for {' and '.join(future_months)}"
                    else:
                        row[f"Target achieved in {m}"] = f"{count_in_month} {month_unit} out of {int(month_target)} planned"
                else:
                    row[f"Target achieved in {m}"] = f"{count_in_month} {month_unit} out of {int(month_target)} planned"
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

        total_target = sum(t5_targets[activity][month][0] for month in MONTHS)
        unit = t5_targets[activity][MONTHS[0]][1] if total_target > 0 else ""
        row["Target Till August"] = (
            f"{int(total_target)} {unit} ({int(t5_targets[activity]['June'][0])} {unit}-June, "
            f"{int(t5_targets[activity]['July'][0])} {unit}-July & {int(t5_targets[activity]['August'][0])} {unit}-August)"
        )
        progress_data.append(row)

    all_cols = ["Milestone", "Activity", "Target Till August",
                "% Work Done against Target-Till June",
                "% Work Done against Target-Till July",
                "% Work Done against Target-Till August",
                "Weightage", "Weighted Delay against Targets",
                "Target achieved in June", "Target achieved in July", "Target achieved in August",
                "Total achieved", "Delay Reasons_June 2025"]
    df_t5 = pd.DataFrame(progress_data, columns=all_cols)
    return df_t5

# -----------------------------------------------------------------------------
# TOWER 7 - UPDATED FOR MODULE-WISE COUNTING
# -----------------------------------------------------------------------------

def get_t7_targets_and_progress(cos):
    raw = download_file_bytes(cos, KRA_KEY)
    wb_kra = load_workbook(filename=BytesIO(raw), data_only=True)
    sheet_kra = wb_kra["VeridiaTargets Till August 2025"]

    t7_targets = {}
    for activity in T7_ACTIVITIES:
        t7_targets[activity] = {}
        for month in MONTHS:
            cell, unit = T7_TARGET_CELLS[activity][month]
            val = extract_number(sheet_kra[cell].value)
            t7_targets[activity][month] = (val, unit)
    logger.info(f"Tower 7 targets: {t7_targets}")

    raw_tracker = download_file_bytes(cos, T7_TRACKER_KEY)
    wb_tracker = load_workbook(filename=BytesIO(raw_tracker), data_only=True)

    # Define activity name mapping for Tower 7
    t7_activity_mapping = {
        "El- First Fix": [
            "El- First Fix",
            "EL- First Fix",
            "EL First Fix",
            "Electrical First Fix",
            "El-First Fix"
        ],
        "Floor Tiling": [
            "Floor Tiling",
            "Flooring Tiling",
            "Tile Flooring"
        ],
        "False Ceiling Framing": [
            "False Ceiling Framing",
            "Ceiling Framing",
            "False Ceiling Frame"
        ],
        "C-Stone flooring": [
            "C-Stone flooring",
            "C Stone flooring",
            "C-Stone Flooring",
            "CStone flooring"
        ]
    }

    # Count completed activities module-wise
    activity_counts = count_completed_activities_by_module_and_month(
        wb_tracker, "TOWER 7 FINISHING.", t7_activity_mapping
    )

    prev_months = get_previous_months()
    month_indices = {m: i for i, m in enumerate(MONTHS)}

    progress_data = []
    total_milestones = len(T7_ACTIVITIES)
    weightage = round(100 / total_milestones, 2) if total_milestones else 0

    for i, activity in enumerate(T7_ACTIVITIES):
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
                months_to_count = MONTHS[: month_indices[m] + 1]
                count_cumulative = sum(activity_counts[activity][month] for month in months_to_count)

                target_cumulative, unit = 0, ""
                for month in months_to_count:
                    val, u = t7_targets[activity][month]
                    target_cumulative += val
                    unit = u

                if target_cumulative == 0:
                    pct_done = 100.0
                else:
                    pct_done = min(round((count_cumulative / target_cumulative) * 100, 2), 100)

                row[f"% Work Done against Target-Till {m}"] = f"{pct_done}%"
                
                # Get individual month target for display
                month_target, month_unit = t7_targets[activity][m]
                # For achieved count, we want to show only what was achieved in that specific month
                count_in_month = activity_counts[activity][m]
                
                if month_target == 0:
                    # Show which future months it's planned for
                    future_months = []
                    for future_m in MONTHS[month_indices[m] + 1:]:
                        future_target, _ = t7_targets[activity][future_m]
                        if future_target > 0:
                            future_months.append(future_m)
                    
                    if future_months:
                        if len(future_months) == 1:
                            row[f"Target achieved in {m}"] = f"Planned for {future_months[0]}"
                        else:
                            row[f"Target achieved in {m}"] = f"Planned for {' and '.join(future_months)}"
                    else:
                        row[f"Target achieved in {m}"] = f"{count_in_month} {month_unit} out of {int(month_target)} planned"
                else:
                    row[f"Target achieved in {m}"] = f"{count_in_month} {month_unit} out of {int(month_target)} planned"
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

        total_target = sum(t7_targets[activity][month][0] for month in MONTHS)
        unit = t7_targets[activity][MONTHS[0]][1] if total_target > 0 else ""
        row["Target Till August"] = (
            f"{int(total_target)} {unit} ({int(t7_targets[activity]['June'][0])} {unit}-June, "
            f"{int(t7_targets[activity]['July'][0])} {unit}-July & {int(t7_targets[activity]['August'][0])} {unit}-August)"
        )
        progress_data.append(row)

    all_cols = ["Milestone", "Activity", "Target Till August",
                "% Work Done against Target-Till June",
                "% Work Done against Target-Till July",
                "% Work Done against Target-Till August",
                "Weightage", "Weighted Delay against Targets",
                "Target achieved in June", "Target achieved in July", "Target achieved in August",
                "Total achieved", "Delay Reasons_June 2025"]
    df_t7 = pd.DataFrame(progress_data, columns=all_cols)
    return df_t7

# -----------------------------------------------------------------------------
# GREEN 3 (External Development)
# -----------------------------------------------------------------------------

def get_green3_targets_and_progress(cos):
    logger.info("Calculating Green 3 External Development Work progress...")
    raw = download_file_bytes(cos, GREEN3_TRACKER_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    sheet = wb.active

    # Only show June milestone as per your latest requirement
    green3_activities = [
        "Pathway Area & Planter",  # Milestone-01
        # (Removed milestones 02 & 03 because only June tracker is required)
    ]

    tracker_name_map = {
        "Pathway Area & Planter": "Pathway Area & Planter",
    }

    progress_data = []
    prev_months = get_previous_months()

    for i, act in enumerate(green3_activities):
        row = {
            "Milestone": f"Milestone-{i+1:02d}",
            "Activity": act,
            "Target": "100% in June" if act == "Pathway Area & Planter" else "Pending",
            "% Work Done against Target-Till June": "",
            "% Work Done against Target-Till July": "",
            "% Work Done against Target-Till August": "",
            "Weightage": 100,
            "Weighted Delay against Targets": "",
            "Target achieved in June": "",
            "Target achieved in July": "",
            "Target achieved in August": "",
            "Total achieved": "",
            "Delay Reasons_June 2025": "",
        }

        tracker_activity_name = tracker_name_map[act]
        found_percent = None
        for sheet_row in sheet.iter_rows(min_row=2, values_only=False):
            cell_val = sheet_row[2].value
            if cell_val and tracker_activity_name.lower() in str(cell_val).lower():
                percent_cell = sheet_row[11]
                if percent_cell.value is not None:
                    val = percent_cell.value
                    if isinstance(val, float) and val <= 1.0:
                        val = val * 100
                    found_percent = round(val, 2)
                break

        for m in MONTHS:
            if m in prev_months:
                if m == "June":
                    row[f"% Work Done against Target-Till {m}"] = f"{found_percent}%" if found_percent is not None else ""
                else:
                    row[f"% Work Done against Target-Till {m}"] = ""

        if "June" in prev_months:
            try:
                pct_june_val = found_percent if found_percent is not None else 0
                row["Weighted Delay against Targets"] = f"{round((pct_june_val * 100) / 100, 2)}%"
            except Exception:
                row["Weighted Delay against Targets"] = ""

        progress_data.append(row)

    all_cols = ["Milestone", "Activity", "Target",
                "% Work Done against Target-Till June",
                "% Work Done against Target-Till July",
                "% Work Done against Target-Till August",
                "Weightage", "Weighted Delay against Targets",
                "Target achieved in June", "Target achieved in July", "Target achieved in August",
                "Total achieved", "Delay Reasons_June 2025"]
    df_green3 = pd.DataFrame(progress_data, columns=all_cols)
    return df_green3

# -----------------------------------------------------------------------------
# WRITER / STYLING (NO UGLY OUTPUT!)
# -----------------------------------------------------------------------------

def write_excel_report(df_t6, df_t5, df_t7, df_green3, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Time Delivery Milestones"

    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    grey = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Grey fill for headers
    bold_font = Font(bold=True)
    normal_font = Font(bold=False)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    def append_df_block(title, df, total_delay_label):
        """Write section title, dataframe and total delay row. Return (start_row, end_row)."""
        start_col = 1
        end_col = len(df.columns)

        # Section title (merged row) - CHANGED TO GREY
        ws.append([title])
        title_row = ws.max_row
        ws.merge_cells(start_row=title_row, start_column=start_col,
                       end_row=title_row, end_column=end_col)
        for cell in ws[title_row]:
            cell.fill = grey  # Changed from yellow to grey
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = border

        # DataFrame rows
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        header_row = title_row + 1
        body_start = header_row + 1
        body_end = ws.max_row

        # Header styling - headers should ALWAYS be bold
        for cell in ws[header_row]:
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = border

        # Body styling - only bold if row number is in ROWS_TO_BOLD
        for r in range(body_start, body_end + 1):
            for cell in ws[r]:
                cell.font = bold_font if r in ROWS_TO_BOLD else normal_font
                cell.alignment = left_align if cell.col_idx in (1, 2) else center_align
                cell.border = border

        # Total delay row - FIXED TO APPEAR UNDER "Weighted Delay against Targets" COLUMN
        try:
            total_delay = sum(float(str(v).strip('%')) for v in df["Weighted Delay against Targets"] if v)
        except Exception:
            total_delay = 0

        # Find the column index for "Weighted Delay against Targets"
        weighted_delay_col_idx = None
        for idx, col_name in enumerate(df.columns, start=1):
            if col_name == "Weighted Delay against Targets":
                weighted_delay_col_idx = idx
                break

        # Create row with empty cells except for the total delay in the correct column
        total_row_data = [""] * end_col
        if weighted_delay_col_idx:
            total_row_data[weighted_delay_col_idx - 1] = f"{round(total_delay, 2)}%"  # -1 because list is 0-indexed
            total_row_data[0] = total_delay_label  # Put label in first column

        ws.append(total_row_data)
        delay_row = ws.max_row
        
        # Style the total delay row
        for idx, cell in enumerate(ws[delay_row], start=1):
            cell.font = bold_font
            cell.fill = yellow
            if idx == 1:  # First column (label)
                cell.alignment = left_align
            elif idx == weighted_delay_col_idx:  # Weighted delay column
                cell.alignment = center_align
            else:
                cell.alignment = center_align
            cell.border = border

        return title_row, delay_row

    # Write all sections without extra empty rows
    append_df_block("Tower 6 Progress Against Milestones", df_t6, "Total Delay Tower 6")
    append_df_block("Tower 5 Progress Against Milestones", df_t5, "Total Delay Tower 5")
    append_df_block("Tower 7 Progress Against Milestones", df_t7, "Total Delay Tower 7")
    append_df_block("External Development (Green 3) Progress Against Milestones (Structure Work)", df_green3, "Total Delay ED")

    # Column widths after wrapping
    for col in ws.columns:
        max_len = 0
        for cell in col:
            text = str(cell.value) if cell.value is not None else ""
            max_len = max(max_len, len(text.split("\n")[0]))
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

    # Consistent row height
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 22

    wb.save(filename)
    logger.info(f"Report saved to {filename}")

# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------

def main():
    cos = init_cos()

    logger.info("Fetching slab targets for Tower 6...")
    targets_t6 = get_slab_targets_fixed_cells(cos)

    logger.info("Downloading Tower 6 tracker workbook...")
    raw_tracker_t6 = download_file_bytes(cos, T6_TRACKER_KEY)
    wb_tracker_t6 = load_workbook(filename=BytesIO(raw_tracker_t6), data_only=True)

    logger.info("Counting completed slabs for Tower 6...")
    completed_t6 = count_tower6_completed(wb_tracker_t6)

    logger.info("Building Tower 6 milestone DataFrame...")
    df_t6 = build_t6_milestone_dataframe(targets_t6, completed_t6)

    logger.info("Calculating Tower 5 targets and progress...")
    df_t5 = get_t5_targets_and_progress(cos)

    logger.info("Calculating Tower 7 targets and progress...")
    df_t7 = get_t7_targets_and_progress(cos)

    logger.info("Calculating Green 3 External Development Work progress...")
    df_green3 = get_green3_targets_and_progress(cos)

    filename = f"Time_Delivery_Milestones_Report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    logger.info("Writing Excel report...")
    write_excel_report(df_t6, df_t5, df_t7, df_green3, filename)

if __name__ == "__main__":
    main()