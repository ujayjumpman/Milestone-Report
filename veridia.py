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

GREEN_HEX = "FF92D050"
MONTHS = ["June", "July", "August"]

TOWER6_ROWS = [4, 5, 6, 7, 9, 10, 14, 15, 16, 17, 19, 20]
TOWER6_COLS = ['FK', 'FM', 'FO', 'FQ', 'FS', 'FU', 'FW', 'FY', 'GA', 'GC', 'GE', 'GG', 'GI', 'GK']

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

def extract_number(cell_value):
    if not cell_value or cell_value == "-":
        return 0.0
    match = re.search(r"(\d+)", str(cell_value))
    return float(match.group(1)) if match else 0.0

def get_previous_months():
    now = datetime.now()
    current_month = now.month
    return ["June"] if 6 < current_month else []

def get_slab_targets_fixed_cells(cos):
    raw = download_file_bytes(cos, KRA_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    sheet = wb["VeridiaTargets Till August 2025"]
    targets = {
        "June": extract_number(sheet["B18"].value),
        "July": extract_number(sheet["C18"].value),
        "August": extract_number(sheet["D18"].value),
    }
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
                if month_name == "June":
                    fill = cell.fill
                    if fill.fill_type == "solid" and fill.start_color:
                        if fill.start_color.rgb == GREEN_HEX:
                            counts[month_name] += 1
    
    return counts

def build_t6_milestone_dataframe(targets, completed):
    prev_months = get_previous_months()
    total_milestones = 1
    weightage = round(100 / total_milestones, 2) if total_milestones else 0

    def pct(m):
        if m == "June" and m in prev_months:
            done = completed.get(m, 0)
            target = targets.get(m, 0)
            if target == 0:
                return "0.0%"
            pct_done = min(round((done / target) * 100, 2), 100)
            return f"{pct_done}%"
        else:
            return ""

    target_text = f"{int(sum(targets.values()))} Slabs ({int(targets['June'])} Slabs-June, {int(targets['July'])} slabs-July & {int(targets['August'])} slabs-August)"

    row = {
        "Milestone": "Milestone-01",
        "Activity": "Slab Casting",
        "Target Till August": target_text,
        "% Work Done against Target-Till June": pct("June"),
        "% Work Done against Target-Till July": pct("July"),
        "% Work Done against Target-Till August": pct("August"),
        "Weightage": weightage,
        "Weighted Delay against Targets": "",
        "Target achieved in June": f"{completed.get('June', 0)} slab cast out of {int(targets['June'])} planned" if "June" in prev_months else "",
        "Target achieved in July": "",
        "Target achieved in August": "",
        "Total achieved": "",
        "Delay Reasons_June 2025": "",
    }

    if "June" in prev_months:
        try:
            june_pct_str = pct("June").replace("%", "")
            if june_pct_str:
                june_pct_val = float(june_pct_str)
                row["Weighted Delay against Targets"] = f"{round((june_pct_val * weightage) / 100, 2)}%"
        except Exception:
            pass

    all_cols = ["Milestone", "Activity", "Target Till August",
                "% Work Done against Target-Till June",
                "% Work Done against Target-Till July",
                "% Work Done against Target-Till August",
                "Weightage", "Weighted Delay against Targets",
                "Target achieved in June", "Target achieved in July", "Target achieved in August",
                "Total achieved", "Delay Reasons_June 2025"]

    final_df = pd.DataFrame(columns=all_cols)
    final_df.loc[0] = row
    return final_df

def count_completed_activities_by_module_and_month(wb, sheet_name, activity_mapping):
    sheet = wb[sheet_name]
    activity_counts = {}
    
    for activity in activity_mapping.keys():
        activity_counts[activity] = {month: 0 for month in MONTHS}
    
    actual_finish_col = None
    activity_name_col = None
    
    # Find the columns for Actual Finish and Activity
    for row in sheet.iter_rows(min_row=1, max_row=10):
        for cell in row:
            if cell.value:
                if "Actual Finish" in str(cell.value):
                    actual_finish_col = cell.column
                if "Activity" in str(cell.value) or "Task" in str(cell.value):
                    activity_name_col = cell.column
        if actual_finish_col:
            break
    
    if not actual_finish_col:
        return activity_counts
    
    if not activity_name_col:
        activity_name_col = 6
    
    # Add debug logging to see what's being processed
    logger.info(f"Processing sheet: {sheet_name}")
    el_first_fix_found = 0  # Counter for debugging
    
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2), start=2):
        try:
            activity_cell = row[activity_name_col - 1] if len(row) >= activity_name_col else None
            if not activity_cell or not activity_cell.value:
                continue
                
            activity_name = str(activity_cell.value).strip()
            mapped_activity = None
            
            # IMPROVED: Check all activities in the mapping systematically
            for standard_name, variations in activity_mapping.items():
                if standard_name == "El- First Fix":
                    # Special comprehensive check for El- First Fix
                    activity_lower = activity_name.lower().strip()
                    if (activity_name == "EL-First Fix" or  # Most common in tracker
                        activity_name == "El- First Fix" or
                        activity_name == "EL- First Fix" or
                        activity_name == "EL First Fix" or
                        activity_name == "El-First Fix" or
                        activity_name == "Electrical First Fix" or
                        activity_lower == "el-first fix" or
                        activity_lower == "el- first fix" or
                        activity_lower == "el first fix"):
                        mapped_activity = standard_name
                        el_first_fix_found += 1
                        logger.info(f"  Found EL-First Fix variant: '{activity_name}' at row {row_idx}")
                        break
                        
                elif standard_name == "Installation of Rear & Front balcony UPVC Windows":
                    # Special handling for UPVC Windows
                    if (activity_name == standard_name or 
                        activity_name == "Installation of Rear &amp; Front balcony UPVC Windows" or
                        activity_name == "Installation of Rear and Front balcony UPVC Windows" or
                        activity_name == "Installation of Rear & Front Balcony UPVC Windows" or
                        activity_name == "Installation of rear & front balcony UPVC Windows"):
                        mapped_activity = standard_name
                        break
                        
                else:
                    # General mapping for other activities
                    if activity_name in variations or activity_name.lower() in [v.lower() for v in variations]:
                        mapped_activity = standard_name
                        break
            
            if not mapped_activity:
                continue
            
            # Check actual finish date
            actual_finish_cell = row[actual_finish_col - 1] if len(row) >= actual_finish_col else None
            if not actual_finish_cell or not actual_finish_cell.value:
                continue
            
            actual_finish_date = None
            if isinstance(actual_finish_cell.value, datetime):
                actual_finish_date = actual_finish_cell.value
            elif isinstance(actual_finish_cell.value, str):
                try:
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
                if month_name == "June":
                    activity_counts[mapped_activity][month_name] += 1
                    if mapped_activity == "El- First Fix":
                        logger.info(f"  Counted EL-First Fix for June: '{activity_name}' on {actual_finish_date.strftime('%Y-%m-%d')}")
                    
        except Exception as e:
            logger.warning(f"Error processing row {row_idx} in sheet {sheet_name}: {e}")
            continue
    
    # Debug logging for El- First Fix specifically
    if "El- First Fix" in activity_counts:
        logger.info(f"Sheet {sheet_name}: Found {el_first_fix_found} EL-First Fix entries, {activity_counts['El- First Fix']['June']} completed in June")
    
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

    raw_tracker = download_file_bytes(cos, T5_TRACKER_KEY)
    wb_tracker = load_workbook(filename=BytesIO(raw_tracker), data_only=True)

    t5_activity_mapping = {
        "Installation of Rear & Front balcony UPVC Windows": [
            "Installation of Rear & Front balcony UPVC Windows",
            "Installation of Rear &amp; Front balcony UPVC Windows",
            "Installation of Rear and Front balcony UPVC Windows"
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

    required_t5_sheets = ["M7 T5", "M6 T5", "M5 T5", "M4 T5", "M3 T5", "M2 T5"]
    t5_sheet_names = []
    available_sheets = wb_tracker.sheetnames
    
    for required_sheet in required_t5_sheets:
        if required_sheet in available_sheets:
            t5_sheet_names.append(required_sheet)
    
    if not t5_sheet_names:
        activity_counts = {}
        for activity in T5_ACTIVITIES:
            activity_counts[activity] = {month: 0 for month in MONTHS}
    else:
        activity_counts = {}
        for activity in T5_ACTIVITIES:
            activity_counts[activity] = {month: 0 for month in MONTHS}

        for sheet_name in t5_sheet_names:
            sheet_counts = count_completed_activities_by_module_and_month(
                wb_tracker, sheet_name, t5_activity_mapping
            )
            
            for activity in T5_ACTIVITIES:
                for month in MONTHS:
                    activity_counts[activity][month] += sheet_counts[activity][month]

    prev_months = get_previous_months()
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
            if m == "June" and m in prev_months:
                count_cumulative = activity_counts[activity]["June"]
                target_cumulative, unit = t5_targets[activity]["June"]

                if target_cumulative == 0:
                    pct_done = 100.0
                else:
                    pct_done = min(round((count_cumulative / target_cumulative) * 100, 2), 100)

                row[f"% Work Done against Target-Till {m}"] = f"{pct_done}%"
                
                month_target, month_unit = t5_targets[activity][m]
                count_in_month = activity_counts[activity][m]
                
                if month_target == 0:
                    future_months = []
                    for future_m in MONTHS[1:]:
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

    raw_tracker = download_file_bytes(cos, T7_TRACKER_KEY)
    wb_tracker = load_workbook(filename=BytesIO(raw_tracker), data_only=True)

    # DEBUGGING: Check what sheets are actually available
    available_sheets = wb_tracker.sheetnames
    logger.info(f"=== T7 TRACKER SHEET DEBUGGING ===")
    logger.info(f"All available sheets in T7 tracker: {available_sheets}")
    
    # Check for M1 specifically and any variations
    m1_variations = [sheet for sheet in available_sheets if 'M1' in sheet.upper()]
    logger.info(f"M1 sheet variations found: {m1_variations}")
    
    # Check for any other T7 sheets we might be missing
    t7_sheets_found = [sheet for sheet in available_sheets if 'T7' in sheet.upper()]
    logger.info(f"All T7 sheets found: {t7_sheets_found}")

    # UPDATED: More comprehensive activity mapping with exact tracker names
    t7_activity_mapping = {
        "El- First Fix": [
            "EL-First Fix",  # This is the actual name in tracker - MOST IMPORTANT
            "El- First Fix",
            "EL- First Fix", 
            "EL First Fix",
            "El-First Fix",
            "Electrical First Fix",
            "el-first fix",
            "el- first fix"
        ],
        "Floor Tiling": [
            "Floor Tiling",
            "Flooring Tiling",
            "Tile Flooring",
            "floor tiling"
        ],
        "False Ceiling Framing": [
            "False Ceiling Framing",
            "Ceiling Framing",
            "False Ceiling Frame",
            "false ceiling framing"
        ],
        "C-Stone flooring": [
            "C-Stone flooring",
            "C Stone flooring",
            "C-Stone Flooring",
            "CStone flooring",
            "c-stone flooring"
        ]
    }

    # UPDATED: Use the actual available T7 sheets instead of hardcoded list
    required_t7_sheets = ["M7 T7", "M6 T7", "M5 T7", "M4 T7", "M3 T7", "M2 T7", "M1 T7"]
    
    # Find all actual T7 sheets available (in case naming is different)
    actual_t7_sheets = []
    for sheet_name in available_sheets:
        # Check for any sheet that contains both a module identifier (M1-M7) and T7
        if any(module in sheet_name.upper() for module in ['M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7']):
            if 'T7' in sheet_name.upper():
                actual_t7_sheets.append(sheet_name)
    
    logger.info(f"Required T7 sheets: {required_t7_sheets}")
    logger.info(f"Actual T7 sheets found: {actual_t7_sheets}")
    
    # Use actual sheets instead of just the required ones
    t7_sheet_names = actual_t7_sheets if actual_t7_sheets else []
    
    # Also check the original method for comparison
    original_method_sheets = []
    for required_sheet in required_t7_sheets:
        if required_sheet in available_sheets:
            original_method_sheets.append(required_sheet)
    
    logger.info(f"Original method would find: {original_method_sheets}")
    logger.info(f"Using sheets: {t7_sheet_names}")
    
    if not t7_sheet_names:
        activity_counts = {}
        for activity in T7_ACTIVITIES:
            activity_counts[activity] = {month: 0 for month in MONTHS}
    else:
        activity_counts = {}
        for activity in T7_ACTIVITIES:
            activity_counts[activity] = {month: 0 for month in MONTHS}

        for sheet_name in t7_sheet_names:
            sheet_counts = count_completed_activities_by_module_and_month(
                wb_tracker, sheet_name, t7_activity_mapping
            )
            
            for activity in T7_ACTIVITIES:
                for month in MONTHS:
                    activity_counts[activity][month] += sheet_counts[activity][month]

    # Enhanced debug logging
    logger.info(f"=== FINAL T7 RESULTS ===")
    logger.info(f"Sheets processed: {t7_sheet_names}")
    logger.info(f"T7 Activity counts for June: {[(act, activity_counts[act]['June']) for act in T7_ACTIVITIES]}")
    total_el_first_fix = activity_counts.get('El- First Fix', {}).get('June', 0)
    logger.info(f"TOTAL EL-FIRST FIX COUNT: {total_el_first_fix}")
    
    # Check if we're missing M1 T7 specifically
    if "M1 T7" not in t7_sheet_names:
        logger.warning("⚠️  M1 T7 sheet is MISSING from processing!")
        logger.warning("This could explain the difference between expected (110) and actual (94) counts")
    
    prev_months = get_previous_months()
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
            if m == "June" and m in prev_months:
                count_cumulative = activity_counts[activity]["June"]
                target_cumulative, unit = t7_targets[activity]["June"]

                if target_cumulative == 0:
                    pct_done = 100.0
                else:
                    pct_done = min(round((count_cumulative / target_cumulative) * 100, 2), 100)

                row[f"% Work Done against Target-Till {m}"] = f"{pct_done}%"
                
                month_target, month_unit = t7_targets[activity][m]
                count_in_month = activity_counts[activity][m]
                
                if month_target == 0:
                    future_months = []
                    for future_m in MONTHS[1:]:
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
        june_target = int(t7_targets[activity]['June'][0])
        july_target = int(t7_targets[activity]['July'][0])
        august_target = int(t7_targets[activity]['August'][0])
        row["Target Till August"] = f"{int(total_target)} {unit} ({june_target} {unit}-June, {july_target} {unit}-July & {august_target} {unit}-August)"
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

def get_green3_targets_and_progress(cos):
    logger.info("Calculating Green 3 External Development Work progress...")
    raw = download_file_bytes(cos, GREEN3_TRACKER_KEY)
    wb = load_workbook(filename=BytesIO(raw), data_only=True)
    
    # Try to find the correct sheet - check available sheet names
    sheet_names = wb.sheetnames
    logger.info(f"Available sheets in Green 3 tracker: {sheet_names}")
    
    # Use the first sheet or try to find a specific one
    sheet = wb.active
    if len(sheet_names) > 1:
        # Look for sheets that might contain the progress data
        for name in sheet_names:
            if any(keyword in name.lower() for keyword in ['progress', 'track', 'work', 'green']):
                sheet = wb[name]
                logger.info(f"Using sheet: {name}")
                break

    # Define activities dynamically parsed from targets - this should come from your KRA or config
    # For now, keeping the structure but making it more flexible
    green3_activities = {
        "June": [
            {
                "parent": "Path Way Area", 
                "activity": "GSB", 
                "target": "100%"
            },
        ],
        "July": [
            {
                "parent": "Water Proofing - Water Body & Gazebo", 
                "activity": "Water Proofing", 
                "target": "100%"
            },
        ],
        "August": [
            {
                "parent": "Stone Work -Water Body & Gazebo", 
                "activity": "Stone Work", 
                "target": "100%"
            },
        ]
    }

    def find_parent_activity_row(sheet, parent_activity_name):
        """Find the row containing the bold parent activity with flexible matching"""
        logger.info(f"=== Looking for BOLD parent activity: '{parent_activity_name}' ===")
        
        # Define variations for parent activity names
        parent_variations = {
            "Path Way Area": ["pathway area", "path way area", "pathway area & planter", "path way area & planter"],
            "Water Proofing - Water Body & Gazebo": ["water proofing", "waterproofing", "water body", "gazebo", "water proofing - water body & gazebo"],
            "Stone Work -Water Body & Gazebo": ["stone work", "stonework", "water body", "gazebo", "stone work -water body & gazebo", "stone work - water body & gazebo"]
        }
        
        # Get variations for this parent activity
        search_terms = parent_variations.get(parent_activity_name, [parent_activity_name.lower()])
        search_terms.append(parent_activity_name.lower())  # Always include the original
        
        logger.info(f"Searching for variations: {search_terms}")
        
        for row_idx in range(1, sheet.max_row + 1):
            for col_idx in range(1, min(sheet.max_column + 1, 10)):  # Check first 10 columns
                cell = sheet.cell(row=row_idx, column=col_idx)
                
                if cell.value and cell.font and cell.font.bold:
                    cell_text = str(cell.value).strip().lower()
                    logger.info(f"Found BOLD text at row {row_idx}, col {col_idx}: '{cell.value}'")
                    
                    # Check if this bold cell matches any of our search terms
                    for search_term in search_terms:
                        if search_term in cell_text or cell_text in search_term:
                            logger.info(f"MATCH! Found parent activity '{parent_activity_name}' (matched with '{search_term}') at row {row_idx}")
                            return row_idx, col_idx
        
        logger.warning(f"Could not find BOLD parent activity: '{parent_activity_name}' with any variations")
        return None, None

    def find_sub_activity_percentage(sheet, parent_row, parent_col, sub_activity_name, max_search_rows=20):
        """Find the sub-activity below the parent and get its %Complete from column L"""
        logger.info(f"=== Looking for sub-activity '{sub_activity_name}' below row {parent_row} ===")
        
        # From the image, %Complete is in column L (column 12)
        percent_complete_col = 12  # Column L
        
        # Search in rows below the parent activity
        for search_row in range(parent_row + 1, min(parent_row + max_search_rows + 1, sheet.max_row + 1)):
            # Check column C (Activity column) which is column 3 based on the image
            activity_col = 3
            
            try:
                cell = sheet.cell(row=search_row, column=activity_col)
                if cell.value:
                    cell_text = str(cell.value).strip()
                    logger.info(f"Checking row {search_row}, activity: '{cell_text}'")
                    
                    # Check if this cell contains our sub-activity (exact or partial match)
                    if (sub_activity_name.lower() in cell_text.lower() or 
                        cell_text.lower() in sub_activity_name.lower() or
                        sub_activity_name.lower() == cell_text.lower()):
                        
                        logger.info(f"Found sub-activity '{sub_activity_name}' at row {search_row}: '{cell_text}'")
                        
                        # Get %Complete from column L (column 12)
                        pct_cell = sheet.cell(row=search_row, column=percent_complete_col)
                        
                        if pct_cell.value is not None:
                            try:
                                val = pct_cell.value
                                logger.info(f"Raw %Complete value for '{sub_activity_name}': {val} (type: {type(val)})")
                                
                                if isinstance(val, str):
                                    # Remove % sign if present and convert
                                    val = val.replace('%', '').strip()
                                    val = float(val)
                                elif isinstance(val, (int, float)):
                                    val = float(val)
                                
                                # Convert to percentage if it's a decimal (0-1 range)
                                if 0 <= val <= 1:
                                    val = val * 100
                                
                                # Validate percentage range
                                if 0 <= val <= 100:
                                    logger.info(f"SUCCESS! Found %Complete for '{sub_activity_name}': {val}% at row {search_row}")
                                    return round(val, 2)
                                else:
                                    logger.warning(f"Percentage value {val} is outside valid range (0-100)")
                                    
                            except (ValueError, TypeError) as e:
                                logger.warning(f"Could not parse percentage value '{pct_cell.value}': {e}")
                        else:
                            logger.warning(f"Found sub-activity '{sub_activity_name}' but %Complete cell is empty")
                        
                        # Found the activity but couldn't get percentage, try next occurrence
                        continue
                            
            except Exception as e:
                logger.debug(f"Error checking cell at row {search_row}: {e}")
                continue
        
        logger.warning(f"Could not find sub-activity '{sub_activity_name}' below parent row {parent_row}")
        return 0

    progress_data = []
    prev_months = get_previous_months()

    # Debug: Print out sheet structure to understand the layout
    logger.info("=== DEBUGGING Green 3 Sheet Structure ===")
    logger.info(f"Sheet max row: {sheet.max_row}, max column: {sheet.max_column}")
    
    # Print first few rows to understand structure and find headers
    for i in range(1, min(11, sheet.max_row + 1)):
        row_data = []
        for j in range(1, min(20, sheet.max_column + 1)):  # Check more columns for headers
            cell = sheet.cell(row=i, column=j)
            value = str(cell.value) if cell.value is not None else ""
            is_bold = cell.font and cell.font.bold
            row_data.append(f"{get_column_letter(j)}{i}:{value}{'(B)' if is_bold else ''}")
        logger.info(f"Row {i}: {row_data}")

    # Process each month's activities
    for month in MONTHS:
        activities_for_month = green3_activities.get(month, [])
        
        for i, act in enumerate(activities_for_month):
            row = {
                "Milestone": f"Milestone-{i+1:02d}",
                "Activity": f"{act['parent']}-{act['activity']}",
                # CHANGE: Use "Target" instead of "Target Till August" for Green 3
                "Target": f"{act['target']} in {month}",
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

            found_percent = 0
            
            # CHANGE: Only process June activities for now, leave July and August blank
            if month == "June" and month in prev_months:
                parent_activity = act['parent']
                sub_activity = act['activity']
                
                logger.info(f"=== Processing {month}: {parent_activity} - {sub_activity} ===")
                
                # Step 1: Find the bold parent activity
                parent_row, parent_col = find_parent_activity_row(sheet, parent_activity)
                
                if parent_row is not None:
                    # Step 2: Find the sub-activity below the parent and get its percentage
                    found_percent = find_sub_activity_percentage(sheet, parent_row, parent_col, sub_activity)
                else:
                    logger.warning(f"Parent activity '{parent_activity}' not found, defaulting to 0%")

                # Set the percentage for June only
                row["% Work Done against Target-Till June"] = f"{found_percent}%"
                row["Target achieved in June"] = f"{found_percent}% completed" if found_percent > 0 else "Not started"
                
                # Calculate weighted delay for June
                try:
                    row["Weighted Delay against Targets"] = f"{round((found_percent * 100) / 100, 2)}%"
                except Exception:
                    row["Weighted Delay against Targets"] = "0%"

            progress_data.append(row)

    # Create DataFrame with modified column name for Green 3
    # CHANGE: Replace "Target Till August" with "Target" for Green 3
    all_cols = ["Milestone", "Activity", "Target",  # Changed from "Target Till August"
                "% Work Done against Target-Till June",
                "% Work Done against Target-Till July",
                "% Work Done against Target-Till August",
                "Weightage", "Weighted Delay against Targets",
                "Target achieved in June", "Target achieved in July", "Target achieved in August",
                "Total achieved", "Delay Reasons_June 2025"]
    
    df_green3 = pd.DataFrame(progress_data, columns=all_cols)
    logger.info(f"Green 3 DataFrame created with {len(df_green3)} rows")
    return df_green3

def write_excel_report(df_t6, df_t5, df_t7, df_green3, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Time Delivery Milestones"

    current_date = datetime.now().strftime("%d-%m-%Y")
    ws.append(["Veridia Time Delivery Milestones Report"])
    ws.append([f"Report Generated on: {current_date}"])
    ws.append([])

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

    max_cols = max(len(df_t6.columns), len(df_t5.columns), len(df_t7.columns), len(df_green3.columns))
    
    ws.merge_cells(f'A1:{get_column_letter(max_cols)}1')
    ws['A1'].font = title_font
    ws['A1'].alignment = center_align
    ws['A1'].fill = grey
    
    ws.merge_cells(f'A2:{get_column_letter(max_cols)}2')
    ws['A2'].font = date_font
    ws['A2'].alignment = center_align

    section_title_rows = set()
    total_delay_rows = set()

    def append_df_block(title, df, total_delay_label):
        start_col = 1
        end_col = len(df.columns)

        ws.append([title])
        title_row = ws.max_row
        section_title_rows.add(title_row)
        ws.merge_cells(start_row=title_row, start_column=start_col,
                       end_row=title_row, end_column=end_col)
        for cell in ws[title_row]:
            cell.fill = grey
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = border

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        header_row = title_row + 1
        body_start = header_row + 1
        body_end = ws.max_row

        for cell in ws[header_row]:
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = border

        for r in range(body_start, body_end + 1):
            for cell in ws[r]:
                cell.font = normal_font
                cell.alignment = left_align if cell.col_idx in (1, 2) else center_align
                cell.border = border

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
        total_delay_rows.add(delay_row)
        
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

    append_df_block("Tower 6 Progress Against Milestones", df_t6, "Total Delay Tower 6")
    append_df_block("Tower 5 Progress Against Milestones", df_t5, "Total Delay Tower 5")
    append_df_block("Tower 7 Progress Against Milestones", df_t7, "Total Delay Tower 7")
    append_df_block("External Development (Green 3) Progress Against Milestones (Structure Work)", df_green3, "Total Delay ED")

    for col in ws.columns:
        max_len = 0
        for cell in col:
            text = str(cell.value) if cell.value is not None else ""
            max_len = max(max_len, len(text.split("\n")[0]))
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 22

    wb.save(filename)

def main():
    cos = init_cos()
    targets_t6 = get_slab_targets_fixed_cells(cos)
    raw_tracker_t6 = download_file_bytes(cos, T6_TRACKER_KEY)
    wb_tracker_t6 = load_workbook(filename=BytesIO(raw_tracker_t6), data_only=True)
    completed_t6 = count_tower6_completed(wb_tracker_t6)
    df_t6 = build_t6_milestone_dataframe(targets_t6, completed_t6)
    df_t5 = get_t5_targets_and_progress(cos)
    df_t7 = get_t7_targets_and_progress(cos)
    df_green3 = get_green3_targets_and_progress(cos)
    filename = f"Veridia_Time_Delivery_Milestone_Report ({datetime.now():%Y-%m-%d}).xlsx"
    write_excel_report(df_t6, df_t5, df_t7, df_green3, filename)

if __name__ == "__main__":
    main()
