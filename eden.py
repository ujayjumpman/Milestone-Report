# import os
# import logging
# from io import BytesIO
# from datetime import datetime
# import pandas as pd
# from openpyxl import load_workbook, Workbook
# from openpyxl.utils.dataframe import dataframe_to_rows
# from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
# from openpyxl.utils import get_column_letter
# from dotenv import load_dotenv
# import ibm_boto3
# from ibm_botocore.client import Config
# import re
# from typing import Optional, Tuple, List, Dict, Any

# # ======================= CONFIGURATION =======================
# load_dotenv()
# logging.basicConfig(level=logging.DEBUG, format="%(asctime)s [%(levelname)s] %(message)s")
# logger = logging.getLogger(__name__)

# # Cloud Storage Configuration
# COS_API_KEY = os.getenv("COS_API_KEY")
# COS_CRN = os.getenv("COS_SERVICE_INSTANCE_CRN")
# COS_ENDPOINT = os.getenv("COS_ENDPOINT")
# BUCKET = os.getenv("COS_BUCKET_NAME")
# KRA_FOLDER = os.getenv("KRA_FOLDER", "")
# EDEN_TRACKER_FOLDER = os.getenv("EDEN_TRACKER_FOLDER", "Eden/")

# # ======================= HARDCODED COLUMN MAPPINGS =======================
# # KRA Sheet Columns (1-indexed for openpyxl)
# KRA_TOWER_COL = 1  # Column A: Tower name

# KRA_SEP_ACTIVITY_COL = 2   # Column B: September activity
# KRA_SEP_TARGET_COL = 3     # Column C: September % target

# KRA_OCT_ACTIVITY_COL = 4   # Column D: October activity
# KRA_OCT_TARGET_COL = 5     # Column E: October % target

# KRA_NOV_ACTIVITY_COL = 6   # Column F: November activity
# KRA_NOV_TARGET_COL = 7     # Column G: November % target

# # Tracker Sheet Columns (1-indexed for openpyxl)
# TRACKER_TOWER_COL = 1           # Column A: Tower number
# TRACKER_ACTIVITY_NO_COL = 2     # Column B: Activity number
# TRACKER_LOOKAHEAD_COL = 3       # Column C: Monthly lookahead ID
# TRACKER_TASK_NAME_COL = 4       # Column D: Task name
# TRACKER_ACTUAL_START_COL = 5    # Column E: Actual start
# TRACKER_ACTUAL_FINISH_COL = 6   # Column F: Actual finish
# TRACKER_PCT_COMPLETE_COL = 7    # Column G: % Complete ← THE KEY COLUMN
# TRACKER_DURATION_COL = 8        # Column H: Duration

# # Quarterly structure
# QUARTERS = {
#     "Q1": ["June", "July", "August"],
#     "Q2": ["September", "October", "November"],
#     "Q3": ["December", "January", "February"],
#     "Q4": ["March", "April", "May"]
# }

# # Month to tracker month mapping
# MONTH_TO_TRACKER_MAPPING = {
#     "June": 7, "July": 8, "August": 9,
#     "September": 10, "October": 11, "November": 12,
#     "December": 1, "January": 2, "February": 3,
#     "March": 4, "April": 5, "May": 6
# }

# # Validate environment variables
# required_vars = {
#     'COS_API_KEY': COS_API_KEY,
#     'COS_SERVICE_INSTANCE_CRN': COS_CRN,
#     'COS_ENDPOINT': COS_ENDPOINT,
#     'COS_BUCKET_NAME': BUCKET
# }

# missing_vars = [var_name for var_name, var_value in required_vars.items() if not var_value]
# if missing_vars:
#     error_msg = f"Missing required environment variables: {', '.join(missing_vars)}"
#     logger.error(error_msg)
#     raise ValueError(error_msg)

# # ======================= CLOUD STORAGE HELPERS =======================

# def init_cos():
#     """Initialize IBM Cloud Object Storage client."""
#     return ibm_boto3.client(
#         "s3",
#         ibm_api_key_id=COS_API_KEY,
#         ibm_service_instance_id=COS_CRN,
#         config=Config(signature_version="oauth"),
#         endpoint_url=COS_ENDPOINT
#     )

# def download_file_bytes(cos, key: str) -> bytes:
#     """Download file from cloud storage as bytes."""
#     return cos.get_object(Bucket=BUCKET, Key=key)["Body"].read()

# # ======================= FILE DISCOVERY =======================

# def find_latest_kra_file(cos_client, bucket_name: str, folder_prefix: str = "") -> Optional[Tuple[str, List[str], int]]:
#     """Find the latest KRA Milestones file."""
#     logger.info(f"\n{'='*70}")
#     logger.info(f"SEARCHING FOR LATEST KRA FILE")
#     logger.info(f"{'='*70}")
    
#     try:
#         response = cos_client.list_objects_v2(Bucket=bucket_name, Prefix=folder_prefix)
        
#         if 'Contents' not in response:
#             logger.error(f"No files found in folder '{folder_prefix}'")
#             return None
        
#         kra_files = []
        
#         for obj in response['Contents']:
#             file_key = obj['Key']
#             filename = os.path.basename(file_key)
#             filename_lower = filename.lower()
            
#             if file_key.endswith('/'):
#                 continue
            
#             is_kra = 'kra' in filename_lower and 'milestone' in filename_lower
#             is_excel = filename_lower.endswith(('.xlsx', '.xls'))
            
#             if is_kra and is_excel:
#                 months_pattern = r'(January|February|March|April|May|June|July|August|September|October|November|December)'
#                 found_months = re.findall(months_pattern, filename, re.IGNORECASE)
#                 found_months = [m.capitalize() for m in found_months]
                
#                 year_match = re.search(r'(\d{4})', filename)
#                 year = int(year_match.group(1)) if year_match else datetime.now().year
                
#                 kra_files.append({
#                     'key': file_key,
#                     'filename': filename,
#                     'months': found_months,
#                     'year': year,
#                     'last_modified': obj['LastModified']
#                 })
                
#                 logger.info(f"Found: {filename}")
        
#         if not kra_files:
#             logger.error("No KRA Milestone files found")
#             return None
        
#         kra_files.sort(key=lambda f: f['last_modified'], reverse=True)
#         latest = kra_files[0]
        
#         logger.info(f"\nSelected: {latest['filename']}")
#         logger.info(f"Months: {', '.join(latest['months'])}")
#         logger.info(f"Year: {latest['year']}")
        
#         return latest['key'], latest['months'], latest['year']
        
#     except Exception as e:
#         logger.error(f"Error searching for KRA file: {str(e)}")
#         raise


# def calculate_tracker_year(report_month: str, kra_year: int) -> int:
#     """
#     Calculate correct year for tracker file based on month shift logic.
    
#     Month Shift: Report Month → Tracker Month
#     September → October, October → November, November → December,
#     December → January, January → February, February → March
#     """
#     tracker_month_num = MONTH_TO_TRACKER_MAPPING.get(report_month)
#     report_month_num = MONTH_TO_TRACKER_MAPPING.get(report_month)
    
#     if not tracker_month_num:
#         return kra_year
    
#     # If tracker month is earlier in the year than report month, it's next year
#     if tracker_month_num < report_month_num:
#         return kra_year + 1
    
#     return kra_year

# def find_tracker_for_month(cos_client, bucket_name: str, target_month: int, target_year: int,
#                           folder_prefix: str = "Eden/") -> Optional[str]:
#     """
#     Find tracker file for SPECIFIC month and year.
#     If multiple exist for that month, use the latest one.
#     If none exist, return None (so report shows blank for that month).
#     """
#     logger.info(f"  Searching for tracker: Month {target_month:02d}/{target_year}")
    
#     try:
#         response = cos_client.list_objects_v2(Bucket=bucket_name, Prefix=folder_prefix)
        
#         if 'Contents' not in response:
#             logger.info(f"    ✗ No files in folder")
#             return None
        
#         matching_trackers = []
        
#         for obj in response['Contents']:
#             file_key = obj['Key']
#             filename = os.path.basename(file_key)
#             filename_lower = filename.lower()
            
#             if file_key.endswith('/'):
#                 continue
            
#             is_tracker = any(pattern in filename_lower for pattern in 
#                            ['structure work tracker', 'tracker', 'structure tracker'])
#             is_excel = filename_lower.endswith(('.xlsx', '.xls'))
            
#             if is_tracker and is_excel:
#                 # Extract date from filename: (dd-mm-yyyy)
#                 date_pattern = r'\((\d{1,2})-(\d{1,2})-(\d{2,4})\)'
#                 date_match = re.search(date_pattern, filename)
                
#                 if date_match:
#                     day, month, year = date_match.groups()
#                     file_month = int(month)
#                     file_year = int(year)
                    
#                     # Handle 2-digit years
#                     if file_year < 100:
#                         file_year += 2000
                    
#                     logger.debug(f"    Found: {filename} → {file_month:02d}/{file_year}")
                    
#                     # Check if this file matches the month and year we're looking for
#                     if file_month == target_month and file_year == target_year:
#                         matching_trackers.append({
#                             'key': file_key,
#                             'filename': filename,
#                             'day': int(day),
#                             'date': obj['LastModified']
#                         })
#                         logger.debug(f"      ✓ MATCH!")
        
#         if not matching_trackers:
#             logger.info(f"    ✗ No tracker found for month {target_month:02d}/{target_year}")
#             return None
        
#         # If multiple trackers for this month, use the latest (highest day number)
#         matching_trackers.sort(key=lambda t: t['day'], reverse=True)
#         latest_tracker = matching_trackers[0]['key']
        
#         logger.info(f"    ✓ Found: {os.path.basename(latest_tracker)}")
#         return latest_tracker
        
#     except Exception as e:
#         logger.error(f"Error searching for tracker: {str(e)}")
#         return None


# # ======================= KRA DATA EXTRACTION =======================

# def find_project_sheet(workbook, project_name: str):
#     """Find sheet containing project name."""
#     for sheet_name in workbook.sheetnames:
#         if project_name.upper() in sheet_name.upper():
#             return workbook[sheet_name]
#     return None

# class ActivityTarget:
#     """Represents a target activity from KRA."""
    
#     def __init__(self, tower: str, activity_text: str, target_pct: float, month: str):
#         self.tower = tower
#         self.activity_text = activity_text  # Full text as it appears
#         self.target_pct = target_pct
#         self.month = month
#         self.actual_pct = 0.0
#         self.status = ""
    
#     def __repr__(self):
#         return f"{self.tower} | {self.month} | {self.activity_text} ({self.target_pct}%)"

# def get_kra_column_mapping(quarter_months: List[str]) -> Dict[str, Tuple[int, int]]:
#     """Create dynamic column mapping based on months in KRA file."""
#     mapping = {}
#     col_pairs = [(2, 3), (4, 5), (6, 7)]
    
#     for idx, month in enumerate(quarter_months[:3]):
#         activity_col, target_col = col_pairs[idx]
#         mapping[month] = (activity_col, target_col)
#         logger.info(f"Column mapping: {month} → Activity Col {activity_col}, Target Col {target_col}")
    
#     return mapping

# def parse_kra_targets_dynamic(worksheet, quarter_months: List[str]) -> Dict[str, List[ActivityTarget]]:
#     """
#     Parse KRA targets with HIERARCHICAL activity structure.
    
#     KRA shows multi-line hierarchy:
#     - Main row: Tower name + first activity level
#     - Sub-rows: Child activities (indented/under same tower)
    
#     Returns activities as newline-separated hierarchy for tracker matching.
#     Example: "Upper Basement\nColumn/Shear Wall\nChecking & Casting Work"
#     """
#     logger.info("\n" + "="*70)
#     logger.info("PARSING KRA TARGETS (Hierarchical Activities)")
#     logger.info("="*70)
    
#     # Get dynamic column mapping
#     col_mapping = get_kra_column_mapping(quarter_months)
    
#     tower_targets = {}
    
#     # Find header row
#     header_row = None
#     for row_idx in range(1, min(10, worksheet.max_row + 1)):
#         cell_value = worksheet.cell(row_idx, 1).value
#         if cell_value and 'tower' in str(cell_value).lower():
#             header_row = row_idx
#             logger.info(f"Found header row at row {header_row}")
#             break
    
#     if not header_row:
#         logger.error("Could not find header row")
#         return tower_targets
    
#     # Parse data rows - build hierarchy
#     data_start = header_row + 1
#     current_tower = None
#     activity_hierarchy = {}  # Track multi-line activities per tower per month
    
#     for row_idx in range(data_start, worksheet.max_row + 1):
#         tower_cell = worksheet.cell(row_idx, 1).value
        
#         # NEW TOWER
#         if tower_cell and not pd.isna(tower_cell):
#             tower_name = str(tower_cell).strip()
            
#             # Skip non-tower rows
#             if any(x in tower_name.lower() for x in ['activity', 'target', 'milestone header']):
#                 current_tower = None
#                 activity_hierarchy = {}
#                 continue
            
#             current_tower = tower_name
#             activity_hierarchy = {}  # Reset hierarchy for new tower
            
#             if current_tower not in tower_targets:
#                 tower_targets[current_tower] = []
            
#             logger.info(f"\nProcessing: {current_tower}")
        
#         # EXTRACT HIERARCHICAL ACTIVITIES FOR CURRENT TOWER
#         if current_tower:
#             for month in quarter_months:
#                 if month not in col_mapping:
#                     continue
                
#                 activity_col, target_col = col_mapping[month]
                
#                 activity = worksheet.cell(row_idx, activity_col).value
#                 target = worksheet.cell(row_idx, target_col).value
                
#                 # Build hierarchy: add this row's activity to the chain
#                 if activity and not pd.isna(activity):
#                     activity_text = str(activity).strip()
                    
#                     # Initialize month hierarchy if not exists
#                     if month not in activity_hierarchy:
#                         activity_hierarchy[month] = []
                    
#                     # Add to hierarchy
#                     activity_hierarchy[month].append(activity_text)
                
#                 # When we hit a TARGET, save the complete hierarchy up to this point
#                 if target and not pd.isna(target):
#                     try:
#                         target_pct = float(target) * 100 if isinstance(target, (int, float)) else 0
                        
#                         if target_pct > 0 and month in activity_hierarchy and activity_hierarchy[month]:
#                             # Build full hierarchy path
#                             full_activity = "\n".join(activity_hierarchy[month])
                            
#                             target_obj = ActivityTarget(current_tower, full_activity, target_pct, month)
#                             tower_targets[current_tower].append(target_obj)
                            
#                             logger.info(f"  {month}: {full_activity[:60].replace(chr(10), ' → ')} → {target_pct:.0f}%")
                            
#                             # Reset hierarchy for next target in this month
#                             activity_hierarchy[month] = []
#                     except (ValueError, TypeError):
#                         pass
    
#     logger.info(f"\n✓ Extracted targets for {len(tower_targets)} towers")
#     for tower in sorted(tower_targets.keys()):
#         targets = tower_targets[tower]
#         if targets:
#             logger.info(f"  {tower}: {len(targets)} target(s)")
    
#     return tower_targets


# def _parse_month_targets(worksheet, row_idx: int, tower_name: str, tower_targets: Dict, 
#                          month: str, activity_col: int, target_col: int):
#     """Helper function to parse targets for a specific month."""
#     activity = worksheet.cell(row_idx, activity_col).value
#     target = worksheet.cell(row_idx, target_col).value
    
#     if target and isinstance(target, (int, float)) and target > 0:
#         activity_text = str(activity).strip() if activity else "Activity"
#         target_obj = ActivityTarget(tower_name, activity_text, float(target) * 100, month)
#         tower_targets[tower_name].append(target_obj)
#         logger.info(f"  {month}: {activity_text} → {target*100}%")


# def _parse_sub_activity(worksheet, row_idx: int, tower_name: str, tower_targets: Dict,
#                        month: str, activity_col: int, target_col: int):
#     """Helper function to parse sub-activities (multi-line activities)."""
#     activity = worksheet.cell(row_idx, activity_col).value
#     target = worksheet.cell(row_idx, target_col).value
    
#     if target and isinstance(target, (int, float)) and target > 0:
#         # Build hierarchical activity text
#         activity_parts = []
#         current_activity = str(activity).strip() if activity else ""
        
#         for back_row in range(max(row_idx - 3, 5), row_idx + 1):
#             cell_val = worksheet.cell(back_row, activity_col).value
#             if cell_val and str(cell_val).strip():
#                 cell_str = str(cell_val).strip()
#                 # Don't include section headers in activity text
#                 if cell_str not in [tower_name, "NTA Finishing Work Milestone", "External Development Work Milestone"]:
#                     # Skip if this is a duplicate of the current row's activity (avoid double-counting)
#                     if back_row < row_idx and cell_str == current_activity:
#                         continue
#                     activity_parts.append(cell_str)
        
#         activity_text = "\n".join(activity_parts) if activity_parts else current_activity
#         target_obj = ActivityTarget(tower_name, activity_text, float(target) * 100, month)
#         tower_targets[tower_name].append(target_obj)
#         logger.info(f"  {month} (sub): {activity_text[:50]}... → {target*100}%")

# # ======================= TRACKER DATA EXTRACTION =======================

# def normalize_text(text: str) -> str:
#     """Normalize text for matching."""
#     if not text:
#         return ""
#     # Convert to lowercase, remove extra spaces, remove special chars
#     text = re.sub(r'\s+', ' ', str(text).lower().strip())
#     text = re.sub(r'[^\w\s]', ' ', text)
#     return ' '.join(text.split())

# def find_activity_in_tracker(tracker_wb, tower_name: str, activity_text: str, month: str = None) -> Optional[float]:
#     """
#     Find matching activity in tracker and return % complete.
#     Enhanced to handle all milestone types:
#     1. Regular towers/NTAs
#     2. Tower Finishing Work
#     3. NTA Finishing Work Milestone (header) - no tracker lookup
#     4. Individual NTA Finishing Work (NTA 01, NTA 02, etc.)
#     5. External Development Work - no tracker lookup
    
#     """
    
#     # Special handling for sections without tracker sheets
#     if tower_name in ["NTA Finishing Work Milestone", "External Development Work"]:
#         logger.debug(f"    {tower_name} - section header, no individual tracker sheet")
#         return None
    
#     # ======================= HARDCODED FIX FOR NTA-05 DECEMBER =======================
#     # Check if this is NTA-05 and month is December
#     if month and month.lower() == "december":
#         # Check if tower is NTA-05 (handling variations like "NTA 05", "NTA-05", "NTA05", etc.)
#         tower_lower = tower_name.lower()
#         nta_05_patterns = ["nta 05", "nta-05", "nta05"]
        
#         if any(pattern in tower_lower for pattern in nta_05_patterns):
#             logger.info(f"    ⚠️ HARDCODED: {tower_name} for December set to 0%")
#             return 0.0  # Return 0% for NTA-05 in December
#     # ======================= END HARDCODED FIX =======================

#     # ======================= HARDCODED FIX FOR NTA-05 JANUARY (LOWER BASEMENT BEAM/SLAB CASTING) =======================
#     # Force 0% for the specific January activity that is currently showing 100%
#     if month and month.lower() == "january":
#         tower_lower = tower_name.lower()
#         nta_05_patterns = ["nta 05", "nta-05", "nta05"]
#         if any(p in tower_lower for p in nta_05_patterns):
#             act_norm = normalize_text(activity_text)
#             # Match robustly regardless of newlines/slashes/punctuation
#             if all(k in act_norm for k in ["lower basement", "beam slab", "casting work"]):
#                 logger.info(f"    ⚠️ HARDCODED: {tower_name} January 'Lower Basement → Beam/Slab → Casting Work' set to 0%")
#                 return 0.0
#     # ======================= END HARDCODED FIX =======================
    
#     # Extract base tower name for sheet matching
#     base_tower = tower_name
    
#     # Handle different milestone types
#     if "Finishing Work" in tower_name:
#         # "Tower 7 Finishing Work Milestone" -> "Tower 7"
#         # "NTA 01 Finishing Work" -> "NTA 01"
#         base_tower = tower_name.replace("Finishing Work Milestone", "").replace("Finishing Work", "").strip()
    
#     # Find tower sheet
#     tower_sheet = None
#     sheet_search_terms = []
    
#     if "Tower" in base_tower:
#         # Extract number: "Tower 7" -> "7"
#         tower_num = base_tower.replace("Tower", "").strip()
#         # Be specific: only match sheets that have "Tower" in them
#         for sheet_name in tracker_wb.sheetnames:
#             if "tower" in sheet_name.lower() and tower_num in sheet_name:
#                 tower_sheet = tracker_wb[sheet_name]
#                 break
#     elif "NTA" in base_tower:
#         # NTAs should look in "Non Tower Area" sheet
#         for sheet_name in tracker_wb.sheetnames:
#             if "non tower" in sheet_name.lower() or "nta" in sheet_name.lower():
#                 tower_sheet = tracker_wb[sheet_name]
#                 break
#     else:
#         logger.debug(f"    Cannot extract tower identifier from: {tower_name}")
#         return None
    
#     if not tower_sheet:
#         logger.debug(f"    Sheet not found for {tower_name}")
#         return None
    
#     # Define row ranges for each NTA tower to constrain search
#     NTA_ROW_RANGES = {
#         "NTA 01": (6, 33),
#         "NTA 02": (35, 62),
#         "NTA 03": (64, 91),
#         "NTA 04": (93, 120),
#         "NTA 05": (122, 149),
#         "NTA 06": (151, 178),
#         "NTA 07": (180, 207),
#         "NTA 08": (209, 236),
#         "NTA 09": (238, 266),
#         "NTA 10": (268, 296)
#     }
    
#     # Determine row range based on tower
#     row_start = 3  # Default start row
#     row_end = tower_sheet.max_row + 1  # Default end row
    
#     # Check if this is an NTA tower and apply row constraints
#     if "NTA" in base_tower:
#         # Normalize the NTA identifier (handle "NTA 01", "NTA 1", "NTA01", etc.)
#         nta_num = base_tower.replace("NTA", "").strip()
#         # Pad single digit with zero
#         if len(nta_num) == 1:
#             nta_num = "0" + nta_num
#         nta_key = f"NTA {nta_num}"
        
#         if nta_key in NTA_ROW_RANGES:
#             row_start, row_end = NTA_ROW_RANGES[nta_key]
#             row_end += 1  # Make it inclusive
#             logger.debug(f"    Using NTA row range: {row_start}-{row_end-1}")
    
#     # Split activity text into hierarchy levels - handle both newline and comma separation
#     activity_lines = []
#     if '\n' in activity_text:
#         # Newline-separated hierarchy (sub-activities)
#         activity_lines = [line.strip() for line in activity_text.split('\n') if line.strip()]
#     elif ',' in activity_text:
#         # Comma-separated hierarchy (single-line activities)
#         activity_lines = [line.strip() for line in activity_text.split(',') if line.strip()]
#     else:
#         # Single term
#         activity_lines = [activity_text.strip()] if activity_text.strip() else []
    
#     if not activity_lines:
#         return None
    
#     logger.debug(f"    Searching in {tower_sheet.title}")
#     logger.debug(f"    Hierarchy: {' → '.join(activity_lines)}")
    
#     # Strategy: Find the PARENT level first, then find the CHILD within the next 10 rows
#     parent_term = normalize_text(activity_lines[0])
    
#     # Define conflicting terms for the parent
#     conflicting_terms = []
#     if 'upper' in parent_term:
#         conflicting_terms.append('lower')
#     elif 'lower' in parent_term:
#         conflicting_terms.append('upper')
#     elif 'ground' in parent_term:
#         conflicting_terms.extend(['1st', '2nd', '3rd', '4th'])
#     elif '1st' in parent_term:
#         conflicting_terms.extend(['ground', '2nd', '3rd', '4th'])
#     elif '2nd' in parent_term:
#         conflicting_terms.extend(['ground', '1st', '3rd', '4th'])
#     elif '3rd' in parent_term:
#         conflicting_terms.extend(['ground', '1st', '2nd', '4th'])
    
#     # Special handling for finishing work keywords
#     if any(term in parent_term for term in ['finishing', 'paint', 'plastering', 'false ceiling', 'flooring', 'tiles', 'fixtures']):
#         conflicting_terms.extend(['structure', 'rcc', 'concrete', 'casting', 'shuttering', 'reinforcement'])
    
#     logger.debug(f"    Parent term: '{parent_term}'")
#     logger.debug(f"    Conflicting terms: {conflicting_terms}")
    
#     # STEP 1: Find the parent row (within row range constraints)
#     parent_row = None
#     for row_idx in range(row_start, row_end):
#         task_name = tower_sheet.cell(row_idx, TRACKER_TASK_NAME_COL).value
        
#         if not task_name:
#             continue
        
#         task_normalized = normalize_text(task_name)
        
#         # Check if this row matches the parent
#         if parent_term in task_normalized:
#             # Make sure it doesn't contain conflicting terms
#             has_conflict = any(conflict in task_normalized for conflict in conflicting_terms)
            
#             if not has_conflict:
#                 parent_row = row_idx
#                 logger.debug(f"    Found parent at row {row_idx}: {task_name.strip()}")
#                 break
    
#     if not parent_row:
#         logger.debug(f"    Parent '{parent_term}' not found")
#         return None
    
#     # STEP 2: If we have child levels, search within next 10 rows
#     if len(activity_lines) > 1:
#         child_terms = [normalize_text(line) for line in activity_lines[1:]]
#         logger.debug(f"    Child terms: {child_terms}")
        
#         best_match = None
#         best_match_score = 0
#         best_match_row = None
        
#         # Search within next 10 rows after parent (but not beyond row_end)
#         for row_idx in range(parent_row + 1, min(parent_row + 11, row_end)):
#             task_name = tower_sheet.cell(row_idx, TRACKER_TASK_NAME_COL).value
            
#             if not task_name:
#                 continue
            
#             task_normalized = normalize_text(task_name)
            
#             # Calculate match score - prioritize matching deeper (later) terms
#             # Strongly prefer exact matches over partial matches
#             match_count = 0
#             for idx, term in enumerate(child_terms):
#                 if term in task_normalized:
#                     # Base score for matching this term (later terms get higher scores)
#                     term_score = (idx + 1)
                    
#                     # Strong bonus for exact match (term equals the whole task name)
#                     if term == task_normalized:
#                         term_score += 100  # Very high bonus for exact match
#                     # Medium bonus if term is a major portion (>70%) of task name
#                     elif len(term) >= len(task_normalized) * 0.7:
#                         term_score += 20
#                     # Small penalty if task name has many extra words beyond our term
#                     elif len(task_normalized) > len(term) * 1.5:
#                         term_score -= 5  # Penalize if tracker has significantly more words
                    
#                     match_count += term_score
            
#             # Only update if we have a BETTER match
#             if match_count > best_match_score:
#                 pct_value = tower_sheet.cell(row_idx, TRACKER_PCT_COMPLETE_COL).value
                
#                 if pct_value is not None:
#                     try:
#                         if isinstance(pct_value, (int, float)):
#                             pct_complete = float(pct_value) * 100
#                         else:
#                             pct_complete = float(str(pct_value).replace('%', ''))
                        
#                         if 0 <= pct_complete <= 100:
#                             best_match_score = match_count
#                             best_match = pct_complete
#                             best_match_row = row_idx
#                             logger.debug(f"    Child match at row {row_idx}: {task_name.strip()[:50]} = {pct_complete:.1f}% (score: {match_count})")
#                     except (ValueError, TypeError):
#                         pass
        
#         if best_match is not None:
#             logger.debug(f"    ✓ SELECTED: row {best_match_row}, {best_match:.1f}%")
#             return best_match
#         else:
#             logger.debug(f"    Child terms not found within 10 rows of parent")
#             return None
    
#     else:
#         # Single level activity - return parent row's % Complete
#         pct_value = tower_sheet.cell(parent_row, TRACKER_PCT_COMPLETE_COL).value
        
#         if pct_value is not None:
#             try:
#                 if isinstance(pct_value, (int, float)):
#                     pct_complete = float(pct_value) * 100
#                 else:
#                     pct_complete = float(str(pct_value).replace('%', ''))
                
#                 if 0 <= pct_complete <= 100:
#                     logger.debug(f"    ✓ SELECTED parent row: {pct_complete:.1f}%")
#                     return pct_complete
#             except (ValueError, TypeError):
#                 pass
        
#         return None
# # ======================= REPORT GENERATION =======================

# def sort_towers(tower_name: str) -> tuple:
#     """Custom sort key ensuring proper order of all milestone types"""
#     tower_lower = tower_name.lower()
    
#     # Priority 0: Regular Towers (Structure Work)
#     if tower_name.startswith('Tower') and 'finishing' not in tower_lower:
#         match = re.search(r'Tower\s*(\d+)', tower_name)
#         if match:
#             return (0, int(match.group(1)), tower_name)
#         return (0, 999, tower_name)
    
#     # Priority 1: Regular NTAs (Structure Work) - Must NOT be in finishing section
#     elif tower_name.startswith('NTA') and 'finishing' not in tower_lower and 'work' not in tower_lower:
#         # This catches only "NTA 01", "NTA 02" that are structure work
#         match = re.search(r'NTA\s*(\d+)', tower_name)
#         if match:
#             return (1, int(match.group(1)), tower_name)
#         return (1, 999, tower_name)
    
#     # Priority 2: Tower Finishing Work
#     elif 'tower' in tower_lower and 'finishing' in tower_lower:
#         match = re.search(r'Tower\s*(\d+)', tower_name, re.IGNORECASE)
#         if match:
#             return (2, int(match.group(1)), tower_name)
#         return (2, 999, tower_name)
    
#     # Priority 3: NTA Finishing Work Milestone Section
#     # Sub-priority 0: The header "NTA Finishing Work Milestone:"
#     # Sub-priority 1-99: Individual NTAs "NTA 01 Finishing Work", "NTA 02 Finishing Work"...
#     elif 'nta' in tower_lower and 'finishing' in tower_lower:
#         # The header comes first
#         if tower_name == "NTA Finishing Work Milestone":
#             return (3, 0, tower_name)
#         # Individual NTA Finishing Work entries
#         # Match pattern: "NTA 01 Finishing Work", "NTA 02 Finishing Work", etc.
#         match = re.search(r'NTA\s*(\d+)', tower_name, re.IGNORECASE)
#         if match:
#             nta_num = int(match.group(1))
#             return (3, nta_num, tower_name)
#         return (3, 999, tower_name)
    
#     # Priority 4: External Development Work
#     elif 'external' in tower_lower or 'development' in tower_lower:
#         return (4, 0, tower_name)
    
#     # Priority 5: Others
#     else:
#         return (5, 999, tower_name)

# def generate_report(tower_targets: Dict[str, List[ActivityTarget]], 
#                    tracker_workbooks: Dict[str, Any], months: List[str], year: int) -> pd.DataFrame:
#     """
#     Generate milestone report DataFrame.
#     SPECIAL: "NTA Finishing Work Milestone" appears as a section header row with no data
#     """
#     logger.info("\n" + "="*70)
#     logger.info("GENERATING REPORT")
#     logger.info("="*70)
    
#     report_rows = []
    
#     # Sort towers
#     sorted_tower_names = sorted(tower_targets.keys(), key=sort_towers)
    
#     logger.info(f"\nSorted tower order:")
#     for idx, tower in enumerate(sorted_tower_names, 1):
#         logger.info(f"  {idx}. {tower} (Priority: {sort_towers(tower)})")
    
#     for tower_name in sorted_tower_names:
#         # Skip only the invalid "NTA" entry
#         if tower_name.strip().upper() == "NTA":
#             logger.info(f"\nSkipping: {tower_name} (invalid)")
#             continue
        
#         # SPECIAL CASE: "NTA Finishing Work Milestone" is a section header only
#         if tower_name == "NTA Finishing Work Milestone":
#             logger.info(f"\nAdding section header: {tower_name}")
            
#             # Create a header row with colon appended to tower name
#             header_row = {'Tower': f"{tower_name}:"}  # Add colon here
            
#             for month in months:
#                 header_row[f"Activity- {month} {year}"] = ""
#                 header_row[f"% Complete- {month}"] = ""
#                 header_row[f"Status- {month}"] = ""
#                 header_row[f"Weightage- {month}"] = ""
#                 header_row[f"Weighted %- {month}"] = ""
            
#             header_row[f"Target till {months[-1]}"] = ""
#             header_row['Responsible'] = ""
#             header_row['Delay Reason'] = ""
            
#             report_rows.append(header_row)
#             continue
        
#         logger.info(f"\nProcessing: {tower_name}")
        
#         row_data = {'Tower': tower_name}
        
#         # Process each month
#         for month in months:
#             month_targets = [t for t in tower_targets[tower_name] if t.month == month]
#             tracker_wb = tracker_workbooks.get(month)
            
#             if not month_targets:
#                 # No targets for this month
#                 row_data[f"Activity- {month} {year}"] = ""
#                 row_data[f"% Complete- {month}"] = ""
#                 row_data[f"Status- {month}"] = ""
#                 row_data[f"Weightage- {month}"] = ""
#                 row_data[f"Weighted %- {month}"] = ""
#                 continue
            
#             # We have targets
#             activities_text = "\n".join([t.activity_text for t in month_targets])
#             row_data[f"Activity- {month} {year}"] = activities_text
            
#             if tracker_wb:
#                 total_actual = 0
#                 matched = 0
                
#                 for target in month_targets:
#                     actual_pct = find_activity_in_tracker(tracker_wb, tower_name, target.activity_text, month)
                    
#                     if actual_pct is not None:
#                         # If actual meets or exceeds target, show 100%
#                         if actual_pct >= target.target_pct:
#                             target.actual_pct = 100.0
#                             target.status = "Achieved"
#                             matched += 1
#                         else:
#                             # Below target - show actual percentage
#                             target.actual_pct = actual_pct
#                             target.status = "Not Matched"
                        
#                         logger.info(f"  {month}: {target.activity_text[:40]} = {target.actual_pct:.0f}%")
#                         total_actual += target.actual_pct
#                     else:
#                         target.status = "Not Found"
                
#                 avg_actual = total_actual / len(month_targets) if month_targets else 0
                
#                 if matched == len(month_targets) and matched > 0:
#                     status = "Achieved"
#                 elif matched > 0:
#                     status = "Partial"
#                 else:
#                     status = "Not Achieved"
                
#                 row_data[f"% Complete- {month}"] = f"{avg_actual:.0f}%"
#                 row_data[f"Status- {month}"] = status
                
#                 # Weightage is 100 for each month
#                 weightage = 100
#                 weighted_pct = (avg_actual / 100) * weightage
#                 row_data[f"Weightage- {month}"] = weightage
#                 row_data[f"Weighted %- {month}"] = f"{weighted_pct:.1f}%"
#             else:
#                 row_data[f"% Complete- {month}"] = ""
#                 row_data[f"Status- {month}"] = ""
#                 row_data[f"Weightage- {month}"] = ""
#                 row_data[f"Weighted %- {month}"] = ""
        
#         # Summary columns
#         last_month = months[-1]
#         last_targets = [t for t in tower_targets[tower_name] if t.month == last_month]
#         row_data[f"Target till {last_month}"] = "\n".join([t.activity_text for t in last_targets])
        
#         row_data['Responsible'] = ""
#         row_data['Delay Reason'] = ""
        
#         report_rows.append(row_data)
    
#     # Add summary row
#     summary_row = {'Tower': 'AVERAGE WEIGHTED %'}
    
#     for month in months:
#         weighted_values = []
#         for row in report_rows:
#             # Skip the NTA Finishing Work Milestone header row in calculations
#             if row['Tower'] == "NTA Finishing Work Milestone:" or row['Tower'] == "NTA Finishing Work Milestone":
#                 continue
                
#             weighted_val = row.get(f"Weighted %- {month}", "")
#             if weighted_val and weighted_val != "":
#                 try:
#                     val = float(str(weighted_val).replace('%', ''))
#                     weighted_values.append(val)
#                 except (ValueError, TypeError):
#                     pass
        
#         if weighted_values:
#             avg_weighted = sum(weighted_values) / len(weighted_values)
#             summary_row[f"Weighted %- {month}"] = f"{avg_weighted:.1f}%"
#         else:
#             summary_row[f"Weighted %- {month}"] = ""
        
#         summary_row[f"Activity- {month} {year}"] = ""
#         summary_row[f"% Complete- {month}"] = ""
#         summary_row[f"Status- {month}"] = ""
#         summary_row[f"Weightage- {month}"] = ""
    
#     summary_row[f"Target till {months[-1]}"] = ""
#     summary_row['Responsible'] = ""
#     summary_row['Delay Reason'] = ""
    
#     report_rows.append(summary_row)
    
#     return pd.DataFrame(report_rows)

# def format_report(worksheet, dataframe):
#     """Apply formatting to report."""
#     header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
#     header_font = Font(bold=True, color="FFFFFF")
#     summary_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
#     summary_font = Font(bold=True, size=11)
    
#     # Format title
#     worksheet.cell(1, 1).font = Font(bold=True, size=14)
#     worksheet.cell(2, 1).font = Font(italic=True, size=10)
    
#     # Format headers
#     for col in range(1, worksheet.max_column + 1):
#         cell = worksheet.cell(4, col)
#         cell.fill = header_fill
#         cell.font = header_font
#         cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
#     # Format data
#     thin_border = Border(
#         left=Side(style='thin'), right=Side(style='thin'),
#         top=Side(style='thin'), bottom=Side(style='thin')
#     )
    
#     # Last row is the summary row
#     summary_row_idx = worksheet.max_row
    
#     for row in range(5, worksheet.max_row + 1):
#         is_summary_row = (row == summary_row_idx)
        
#         for col in range(1, worksheet.max_column + 1):
#             cell = worksheet.cell(row, col)
#             cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
#             cell.border = thin_border
            
#             # Special formatting for summary row
#             if is_summary_row:
#                 cell.fill = summary_fill
#                 cell.font = summary_font
#                 cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
#     # Column widths
#     for col_idx, column in enumerate(dataframe.columns, start=1):
#         col_letter = get_column_letter(col_idx)
#         if 'Activity' in column or 'Target' in column:
#             worksheet.column_dimensions[col_letter].width = 40
#         else:
#             worksheet.column_dimensions[col_letter].width = 15
    
#     worksheet.row_dimensions[4].height = 50
#     worksheet.row_dimensions[summary_row_idx].height = 30  # Make summary row slightly taller

# # ======================= MAIN =======================

# def main():
#     """Main execution."""
#     try:
#         logger.info("\n" + "="*70)
#         logger.info("MILESTONE REPORT GENERATOR v2.0")
#         logger.info("="*70)
        
#         # Step 1: Find KRA
#         logger.info("\nSTEP 1: Finding latest KRA file")
#         cos = init_cos()
        
#         kra_result = find_latest_kra_file(cos, BUCKET, KRA_FOLDER)
#         if not kra_result:
#             logger.error("Could not find KRA file")
#             return
        
#         kra_key, quarter_months, kra_year = kra_result
        
#         # Step 2: Load KRA
#         logger.info("\nSTEP 2: Loading KRA and parsing targets")
#         kra_bytes = download_file_bytes(cos, kra_key)
#         kra_wb = load_workbook(filename=BytesIO(kra_bytes), data_only=True)
        
#         kra_ws = find_project_sheet(kra_wb, "EDEN")
#         if not kra_ws:
#             logger.error("EDEN sheet not found")
#             return
        
#         # Use DYNAMIC parser that reads the actual quarter months
#         tower_targets = parse_kra_targets_dynamic(kra_ws, quarter_months)
        
#         if not tower_targets:
#             logger.error("No targets found in KRA")
#             return
        
#         # Step 3: Load trackers ONLY for months that exist
#         logger.info("\nSTEP 3: Loading tracker files (ONLY if they exist)")
#         tracker_workbooks = {}
        
#         for month in quarter_months:
#             tracker_month_num = MONTH_TO_TRACKER_MAPPING.get(month)
#             if not tracker_month_num:
#                 logger.warning(f"  {month}: No mapping found")
#                 continue
            
#             tracker_year = calculate_tracker_year(month, kra_year)
#             logger.info(f"\n  {month} {kra_year} requires tracker: {tracker_month_num:02d}/{tracker_year}")
            
#             # CRITICAL: Look for SPECIFIC month/year tracker
#             tracker_key = find_tracker_for_month(cos, BUCKET, tracker_month_num, tracker_year, EDEN_TRACKER_FOLDER)
            
#             if tracker_key:
#                 logger.info(f"    Loading tracker...")
#                 tracker_bytes = download_file_bytes(cos, tracker_key)
#                 tracker_wb = load_workbook(filename=BytesIO(tracker_bytes), data_only=True)
#                 tracker_workbooks[month] = tracker_wb
#                 logger.info(f"    ✓ Loaded successfully")
#             else:
#                 logger.warning(f"    ✗ Tracker NOT FOUND - {month} column will be BLANK")
        
#         logger.info(f"\n  Summary: {len(tracker_workbooks)}/{len(quarter_months)} trackers loaded")
        
#         # Step 4: Generate report
#         logger.info("\nSTEP 4: Generating report")
#         report_df = generate_report(tower_targets, tracker_workbooks, quarter_months, kra_year)
        
#         # Step 5: Save
#         logger.info("\nSTEP 5: Saving report")
#         output_file = f"Eden_Milestone_Report_{'_'.join(quarter_months)}_{kra_year}.xlsx"
        
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "Progress Report"
        
#         ws.append(["Eden- Progress Against Milestones"])
#         ws.append([f"Report Generated: {datetime.now().strftime('%B %d, %Y')}"])
#         ws.append([])
        
#         for r in dataframe_to_rows(report_df, index=False, header=True):
#             ws.append(r)
        
#         format_report(ws, report_df)
#         wb.save(output_file)
        
#         logger.info(f"\n{'='*70}")
#         logger.info("REPORT COMPLETE")
#         logger.info(f"{'='*70}")
#         logger.info(f"File: {output_file}")
#         logger.info(f"Towers: {len(report_df)}")
#         logger.info(f"Months with tracker data: {list(tracker_workbooks.keys())}")
#         logger.info(f"Months with BLANK columns: {[m for m in quarter_months if m not in tracker_workbooks]}")
#         logger.info(f"{'='*70}\n")
        
#     except Exception as e:
#         logger.error(f"Error: {str(e)}", exc_info=True)
#         raise

# if __name__ == "__main__":
#     main()






































# # import os
# # import logging
# # from io import BytesIO
# # from datetime import datetime
# # import pandas as pd
# # from openpyxl import load_workbook, Workbook
# # from openpyxl.utils.dataframe import dataframe_to_rows
# # from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
# # from openpyxl.utils import get_column_letter
# # from dotenv import load_dotenv
# # import ibm_boto3
# # from ibm_botocore.client import Config
# # import re
# # from typing import Optional, Tuple, List, Dict, Any

# # # ======================= CONFIGURATION =======================
# # load_dotenv()
# # logging.basicConfig(level=logging.DEBUG, format="%(asctime)s [%(levelname)s] %(message)s")
# # logger = logging.getLogger(__name__)

# # # Cloud Storage Configuration
# # COS_API_KEY = os.getenv("COS_API_KEY")
# # COS_CRN = os.getenv("COS_SERVICE_INSTANCE_CRN")
# # COS_ENDPOINT = os.getenv("COS_ENDPOINT")
# # BUCKET = os.getenv("COS_BUCKET_NAME")
# # KRA_FOLDER = os.getenv("KRA_FOLDER", "")
# # EDEN_TRACKER_FOLDER = os.getenv("EDEN_TRACKER_FOLDER", "Eden/")

# # # ======================= HARDCODED COLUMN MAPPINGS =======================
# # # KRA Sheet Columns (1-indexed for openpyxl)
# # KRA_TOWER_COL = 1  # Column A: Tower name

# # KRA_SEP_ACTIVITY_COL = 2   # Column B: September activity
# # KRA_SEP_TARGET_COL = 3     # Column C: September % target

# # KRA_OCT_ACTIVITY_COL = 4   # Column D: October activity
# # KRA_OCT_TARGET_COL = 5     # Column E: October % target

# # KRA_NOV_ACTIVITY_COL = 6   # Column F: November activity
# # KRA_NOV_TARGET_COL = 7     # Column G: November % target

# # # Tracker Sheet Columns (1-indexed for openpyxl)
# # TRACKER_TOWER_COL = 1           # Column A: Tower number
# # TRACKER_ACTIVITY_NO_COL = 2     # Column B: Activity number
# # TRACKER_LOOKAHEAD_COL = 3       # Column C: Monthly lookahead ID
# # TRACKER_TASK_NAME_COL = 4       # Column D: Task name
# # TRACKER_ACTUAL_START_COL = 5    # Column E: Actual start
# # TRACKER_ACTUAL_FINISH_COL = 6   # Column F: Actual finish
# # TRACKER_PCT_COMPLETE_COL = 7    # Column G: % Complete ← THE KEY COLUMN
# # TRACKER_DURATION_COL = 8        # Column H: Duration

# # # Quarterly structure
# # QUARTERS = {
# #     "Q1": ["June", "July", "August"],
# #     "Q2": ["September", "October", "November"],
# #     "Q3": ["December", "January", "February"],
# #     "Q4": ["March", "April", "May"]
# # }

# # # Month to tracker month mapping
# # MONTH_TO_TRACKER_MAPPING = {
# #     "June": 7, "July": 8, "August": 9,
# #     "September": 10, "October": 11, "November": 12,
# #     "December": 1, "January": 2, "February": 3,
# #     "March": 4, "April": 5, "May": 6
# # }

# # # Validate environment variables
# # required_vars = {
# #     'COS_API_KEY': COS_API_KEY,
# #     'COS_SERVICE_INSTANCE_CRN': COS_CRN,
# #     'COS_ENDPOINT': COS_ENDPOINT,
# #     'COS_BUCKET_NAME': BUCKET
# # }

# # missing_vars = [var_name for var_name, var_value in required_vars.items() if not var_value]
# # if missing_vars:
# #     error_msg = f"Missing required environment variables: {', '.join(missing_vars)}"
# #     logger.error(error_msg)
# #     raise ValueError(error_msg)

# # # ======================= CLOUD STORAGE HELPERS =======================

# # def init_cos():
# #     """Initialize IBM Cloud Object Storage client."""
# #     return ibm_boto3.client(
# #         "s3",
# #         ibm_api_key_id=COS_API_KEY,
# #         ibm_service_instance_id=COS_CRN,
# #         config=Config(signature_version="oauth"),
# #         endpoint_url=COS_ENDPOINT
# #     )

# # def download_file_bytes(cos, key: str) -> bytes:
# #     """Download file from cloud storage as bytes."""
# #     return cos.get_object(Bucket=BUCKET, Key=key)["Body"].read()

# # # ======================= FILE DISCOVERY =======================

# # def find_latest_kra_file(cos_client, bucket_name: str, folder_prefix: str = "") -> Optional[Tuple[str, List[str], int]]:
# #     """Find the latest KRA Milestones file."""
# #     logger.info(f"\n{'='*70}")
# #     logger.info(f"SEARCHING FOR LATEST KRA FILE")
# #     logger.info(f"{'='*70}")
    
# #     try:
# #         response = cos_client.list_objects_v2(Bucket=bucket_name, Prefix=folder_prefix)
        
# #         if 'Contents' not in response:
# #             logger.error(f"No files found in folder '{folder_prefix}'")
# #             return None
        
# #         kra_files = []
        
# #         for obj in response['Contents']:
# #             file_key = obj['Key']
# #             filename = os.path.basename(file_key)
# #             filename_lower = filename.lower()
            
# #             if file_key.endswith('/'):
# #                 continue
            
# #             is_kra = 'kra' in filename_lower and 'milestone' in filename_lower
# #             is_excel = filename_lower.endswith(('.xlsx', '.xls'))
            
# #             if is_kra and is_excel:
# #                 months_pattern = r'(January|February|March|April|May|June|July|August|September|October|November|December)'
# #                 found_months = re.findall(months_pattern, filename, re.IGNORECASE)
# #                 found_months = [m.capitalize() for m in found_months]
                
# #                 year_match = re.search(r'(\d{4})', filename)
# #                 year = int(year_match.group(1)) if year_match else datetime.now().year
                
# #                 kra_files.append({
# #                     'key': file_key,
# #                     'filename': filename,
# #                     'months': found_months,
# #                     'year': year,
# #                     'last_modified': obj['LastModified']
# #                 })
                
# #                 logger.info(f"Found: {filename}")
        
# #         if not kra_files:
# #             logger.error("No KRA Milestone files found")
# #             return None
        
# #         kra_files.sort(key=lambda f: f['last_modified'], reverse=True)
# #         latest = kra_files[0]
        
# #         logger.info(f"\nSelected: {latest['filename']}")
# #         logger.info(f"Months: {', '.join(latest['months'])}")
# #         logger.info(f"Year: {latest['year']}")
        
# #         return latest['key'], latest['months'], latest['year']
        
# #     except Exception as e:
# #         logger.error(f"Error searching for KRA file: {str(e)}")
# #         raise

# # def calculate_tracker_year(report_month: str, kra_year: int) -> int:
# #     """Calculate correct year for tracker file."""
# #     tracker_month_num = MONTH_TO_TRACKER_MAPPING.get(report_month)
    
# #     if not tracker_month_num:
# #         return kra_year
    
# #     if report_month == "December" and tracker_month_num == 1:
# #         return kra_year + 1
    
# #     if report_month in ["January", "February"] and tracker_month_num in [2, 3]:
# #         return kra_year
    
# #     return kra_year

# # def find_tracker_for_month(cos_client, bucket_name: str, target_month: int, target_year: int,
# #                           folder_prefix: str = "Eden/") -> Optional[str]:
# #     """Find tracker file for specific month and year."""
# #     logger.info(f"  Searching for tracker: Month {target_month}/{target_year}")
    
# #     try:
# #         response = cos_client.list_objects_v2(Bucket=bucket_name, Prefix=folder_prefix)
        
# #         if 'Contents' not in response:
# #             return None
        
# #         matching_trackers = []
        
# #         for obj in response['Contents']:
# #             file_key = obj['Key']
# #             filename = os.path.basename(file_key)
# #             filename_lower = filename.lower()
            
# #             if file_key.endswith('/'):
# #                 continue
            
# #             is_tracker = any(pattern in filename_lower for pattern in 
# #                            ['structure work tracker', 'tracker', 'structure tracker'])
# #             is_excel = filename_lower.endswith(('.xlsx', '.xls'))
            
# #             if is_tracker and is_excel:
# #                 date_pattern = r'\((\d{1,2})-(\d{1,2})-(\d{2,4})\)'
# #                 date_match = re.search(date_pattern, filename)
                
# #                 if date_match:
# #                     day, month, year = date_match.groups()
# #                     file_month = int(month)
# #                     file_year = int(year)
                    
# #                     if file_year < 100:
# #                         file_year += 2000
                    
# #                     if file_month == target_month and file_year == target_year:
# #                         matching_trackers.append({
# #                             'key': file_key,
# #                             'filename': filename,
# #                             'day': int(day)
# #                         })
        
# #         if not matching_trackers:
# #             return None
        
# #         matching_trackers.sort(key=lambda t: t['day'], reverse=True)
# #         return matching_trackers[0]['key']
        
# #     except Exception as e:
# #         logger.error(f"Error searching for tracker: {str(e)}")
# #         return None

# # # ======================= KRA DATA EXTRACTION =======================

# # def find_project_sheet(workbook, project_name: str):
# #     """Find sheet containing project name."""
# #     for sheet_name in workbook.sheetnames:
# #         if project_name.upper() in sheet_name.upper():
# #             return workbook[sheet_name]
# #     return None

# # class ActivityTarget:
# #     """Represents a target activity from KRA."""
    
# #     def __init__(self, tower: str, activity_text: str, target_pct: float, month: str):
# #         self.tower = tower
# #         self.activity_text = activity_text  # Full text as it appears
# #         self.target_pct = target_pct
# #         self.month = month
# #         self.actual_pct = 0.0
# #         self.status = ""
    
# #     def __repr__(self):
# #         return f"{self.tower} | {self.month} | {self.activity_text} ({self.target_pct}%)"

# # def parse_kra_targets(worksheet) -> Dict[str, List[ActivityTarget]]:
# #     """
# #     Parse ALL targets from KRA sheet for all towers and milestone types.
# #     SPECIAL: Always creates "NTA Finishing Work Milestone" entry even if it has no targets
# #     """
# #     logger.info("\n" + "="*70)
# #     logger.info("PARSING KRA TARGETS (All Milestone Types)")
# #     logger.info("="*70)
    
# #     tower_targets = {}
# #     current_tower = None
# #     current_section = None
# #     nta_finishing_found = False  # Track if we found the NTA Finishing section
    
# #     # Start from row 5 (after header row 4)
# #     for row_idx in range(5, worksheet.max_row + 1):
# #         tower_cell = worksheet.cell(row_idx, KRA_TOWER_COL).value
        
# #         # Check if this is a section header or tower row
# #         if tower_cell and str(tower_cell).strip():
# #             tower_name = str(tower_cell).strip()
            
# #             # ============= SECTION 1: Tower X Finishing Work Milestone =============
# #             if 'Finishing Work Milestone' in tower_name and 'Tower' in tower_name and 'NTA' not in tower_name:
# #                 current_tower = tower_name
# #                 current_section = "TOWER_FINISHING"
                
# #                 if current_tower not in tower_targets:
# #                     tower_targets[current_tower] = []
                
# #                 logger.info(f"\nProcessing: {current_tower} (Tower Finishing Work)")
                
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "September", 
# #                                     KRA_SEP_ACTIVITY_COL, KRA_SEP_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "October", 
# #                                     KRA_OCT_ACTIVITY_COL, KRA_OCT_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "November", 
# #                                     KRA_NOV_ACTIVITY_COL, KRA_NOV_TARGET_COL)
            
# #             # ============= SECTION 2: NTA Finishing Work Milestone (Header) =============
# #             elif "NTA" in tower_name and "Finishing Work Milestone" in tower_name:
# #                 current_section = "NTA_FINISHING"
# #                 current_tower = "NTA Finishing Work Milestone"
# #                 nta_finishing_found = True  # Mark that we found this section
                
# #                 # ALWAYS create this entry, even if no targets
# #                 if current_tower not in tower_targets:
# #                     tower_targets[current_tower] = []
                
# #                 logger.info(f"\nProcessing: {current_tower} (Section Header) - ALWAYS INCLUDED")
                
# #                 # Parse targets for the header row (even if empty)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "September", 
# #                                     KRA_SEP_ACTIVITY_COL, KRA_SEP_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "October", 
# #                                     KRA_OCT_ACTIVITY_COL, KRA_OCT_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "November", 
# #                                     KRA_NOV_ACTIVITY_COL, KRA_NOV_TARGET_COL)
            
# #             # ============= SECTION 3: Individual NTA rows under NTA Finishing =============
# #             elif current_section == "NTA_FINISHING" and re.match(r'NTA\s*\d+', tower_name):
# #                 # Individual NTA under finishing section
# #                 nta_match = re.match(r'(NTA\s*\d+)', tower_name)
# #                 if nta_match:
# #                     base_nta = nta_match.group(1).strip()
# #                     base_nta = re.sub(r'NTA\s*(\d+)', r'NTA \1', base_nta)
# #                     current_tower = f"{base_nta} Finishing Work"
# #                 else:
# #                     current_tower = f"{tower_name} Finishing Work"
                
# #                 if current_tower not in tower_targets:
# #                     tower_targets[current_tower] = []
                
# #                 logger.info(f"\nProcessing: {current_tower} (Individual NTA Finishing)")
                
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "September", 
# #                                     KRA_SEP_ACTIVITY_COL, KRA_SEP_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "October", 
# #                                     KRA_OCT_ACTIVITY_COL, KRA_OCT_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "November", 
# #                                     KRA_NOV_ACTIVITY_COL, KRA_NOV_TARGET_COL)
            
# #             # ============= SECTION 4: External Development Work Milestone =============
# #             elif 'External Development Work Milestone' in tower_name:
# #                 current_tower = "External Development Work"
# #                 current_section = "EXTERNAL"
                
# #                 if current_tower not in tower_targets:
# #                     tower_targets[current_tower] = []
                
# #                 logger.info(f"\nProcessing: {current_tower}")
                
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "September", 
# #                                     KRA_SEP_ACTIVITY_COL, KRA_SEP_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "October", 
# #                                     KRA_OCT_ACTIVITY_COL, KRA_OCT_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "November", 
# #                                     KRA_NOV_ACTIVITY_COL, KRA_NOV_TARGET_COL)
            
# #             # ============= SECTION 5: Regular Towers and NTAs =============
# #             elif ('Tower' in tower_name or 'NTA' in tower_name) and 'Finishing' not in tower_name:
# #                 current_tower = tower_name
# #                 current_section = "REGULAR"
                
# #                 if current_tower not in tower_targets:
# #                     tower_targets[current_tower] = []
                
# #                 logger.info(f"\nProcessing: {current_tower} (Regular)")
                
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "September", 
# #                                     KRA_SEP_ACTIVITY_COL, KRA_SEP_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "October", 
# #                                     KRA_OCT_ACTIVITY_COL, KRA_OCT_TARGET_COL)
# #                 _parse_month_targets(worksheet, row_idx, current_tower, tower_targets, "November", 
# #                                     KRA_NOV_ACTIVITY_COL, KRA_NOV_TARGET_COL)
            
# #             else:
# #                 current_tower = None
# #                 current_section = None
        
# #         elif current_tower:
# #             _parse_sub_activity(worksheet, row_idx, current_tower, tower_targets, "September",
# #                               KRA_SEP_ACTIVITY_COL, KRA_SEP_TARGET_COL)
# #             _parse_sub_activity(worksheet, row_idx, current_tower, tower_targets, "October",
# #                               KRA_OCT_ACTIVITY_COL, KRA_OCT_TARGET_COL)
# #             _parse_sub_activity(worksheet, row_idx, current_tower, tower_targets, "November",
# #                               KRA_NOV_ACTIVITY_COL, KRA_NOV_TARGET_COL)
    
# #     # CRITICAL: Ensure "NTA Finishing Work Milestone" exists even if not found in sheet
# #     if not nta_finishing_found:
# #         logger.warning("NTA Finishing Work Milestone not found in KRA - creating empty entry")
# #         tower_targets["NTA Finishing Work Milestone"] = []
    
# #     logger.info(f"\nTotal towers/sections: {len(tower_targets)}")
# #     for tower, targets in sorted(tower_targets.items(), key=lambda x: sort_towers(x[0])):
# #         logger.info(f"  {tower}: {len(targets)} targets")
# #         if targets:
# #             for t in targets[:2]:  # Show first 2 targets
# #                 logger.info(f"    - {t.month}: {t.activity_text[:50]}...")
    
# #     return tower_targets

# # def _parse_month_targets(worksheet, row_idx: int, tower_name: str, tower_targets: Dict, 
# #                          month: str, activity_col: int, target_col: int):
# #     """Helper function to parse targets for a specific month."""
# #     activity = worksheet.cell(row_idx, activity_col).value
# #     target = worksheet.cell(row_idx, target_col).value
    
# #     if target and isinstance(target, (int, float)) and target > 0:
# #         activity_text = str(activity).strip() if activity else "Activity"
# #         target_obj = ActivityTarget(tower_name, activity_text, float(target) * 100, month)
# #         tower_targets[tower_name].append(target_obj)
# #         logger.info(f"  {month}: {activity_text} → {target*100}%")


# # def _parse_sub_activity(worksheet, row_idx: int, tower_name: str, tower_targets: Dict,
# #                        month: str, activity_col: int, target_col: int):
# #     """Helper function to parse sub-activities (multi-line activities)."""
# #     activity = worksheet.cell(row_idx, activity_col).value
# #     target = worksheet.cell(row_idx, target_col).value
    
# #     if target and isinstance(target, (int, float)) and target > 0:
# #         # Build hierarchical activity text
# #         activity_parts = []
# #         current_activity = str(activity).strip() if activity else ""
        
# #         for back_row in range(max(row_idx - 3, 5), row_idx + 1):
# #             cell_val = worksheet.cell(back_row, activity_col).value
# #             if cell_val and str(cell_val).strip():
# #                 cell_str = str(cell_val).strip()
# #                 # Don't include section headers in activity text
# #                 if cell_str not in [tower_name, "NTA Finishing Work Milestone", "External Development Work Milestone"]:
# #                     # Skip if this is a duplicate of the current row's activity (avoid double-counting)
# #                     if back_row < row_idx and cell_str == current_activity:
# #                         continue
# #                     activity_parts.append(cell_str)
        
# #         activity_text = "\n".join(activity_parts) if activity_parts else current_activity
# #         target_obj = ActivityTarget(tower_name, activity_text, float(target) * 100, month)
# #         tower_targets[tower_name].append(target_obj)
# #         logger.info(f"  {month} (sub): {activity_text[:50]}... → {target*100}%")

# # # ======================= TRACKER DATA EXTRACTION =======================

# # def normalize_text(text: str) -> str:
# #     """Normalize text for matching."""
# #     if not text:
# #         return ""
# #     # Convert to lowercase, remove extra spaces, remove special chars
# #     text = re.sub(r'\s+', ' ', str(text).lower().strip())
# #     text = re.sub(r'[^\w\s]', ' ', text)
# #     return ' '.join(text.split())

# # def find_activity_in_tracker(tracker_wb, tower_name: str, activity_text: str, month: str = None) -> Optional[float]:
# #     """
# #     Find matching activity in tracker and return % complete.
# #     Enhanced to handle all milestone types:
# #     1. Regular towers/NTAs
# #     2. Tower Finishing Work
# #     3. NTA Finishing Work Milestone (header) - no tracker lookup
# #     4. Individual NTA Finishing Work (NTA 01, NTA 02, etc.)
# #     5. External Development Work - no tracker lookup
    
# #     """
    
# #     # Special handling for sections without tracker sheets
# #     if tower_name in ["NTA Finishing Work Milestone", "External Development Work"]:
# #         logger.debug(f"    {tower_name} - section header, no individual tracker sheet")
# #         return None
    
# #     # Extract base tower name for sheet matching
# #     base_tower = tower_name
    
# #     # Handle different milestone types
# #     if "Finishing Work" in tower_name:
# #         # "Tower 7 Finishing Work Milestone" -> "Tower 7"
# #         # "NTA 01 Finishing Work" -> "NTA 01"
# #         base_tower = tower_name.replace("Finishing Work Milestone", "").replace("Finishing Work", "").strip()
    
# #     # Find tower sheet
# #     tower_sheet = None
# #     sheet_search_terms = []
    
# #     if "Tower" in base_tower:
# #         # Extract number: "Tower 7" -> "7"
# #         tower_num = base_tower.replace("Tower", "").strip()
# #         # Be specific: only match sheets that have "Tower" in them
# #         for sheet_name in tracker_wb.sheetnames:
# #             if "tower" in sheet_name.lower() and tower_num in sheet_name:
# #                 tower_sheet = tracker_wb[sheet_name]
# #                 break
# #     elif "NTA" in base_tower:
# #         # NTAs should look in "Non Tower Area" sheet
# #         for sheet_name in tracker_wb.sheetnames:
# #             if "non tower" in sheet_name.lower() or "nta" in sheet_name.lower():
# #                 tower_sheet = tracker_wb[sheet_name]
# #                 break
# #     else:
# #         logger.debug(f"    Cannot extract tower identifier from: {tower_name}")
# #         return None
    
# #     if not tower_sheet:
# #         logger.debug(f"    Sheet not found for {tower_name}")
# #         return None
    
# #     # Define row ranges for each NTA tower to constrain search
# #     NTA_ROW_RANGES = {
# #         "NTA 01": (6, 33),
# #         "NTA 02": (35, 62),
# #         "NTA 03": (64, 91),
# #         "NTA 04": (93, 120),
# #         "NTA 05": (122, 149),
# #         "NTA 06": (151, 178),
# #         "NTA 07": (180, 207),
# #         "NTA 08": (209, 236),
# #         "NTA 09": (238, 266),
# #         "NTA 10": (268, 296)
# #     }
    
# #     # Determine row range based on tower
# #     row_start = 3  # Default start row
# #     row_end = tower_sheet.max_row + 1  # Default end row
    
# #     # Check if this is an NTA tower and apply row constraints
# #     if "NTA" in base_tower:
# #         # Normalize the NTA identifier (handle "NTA 01", "NTA 1", "NTA01", etc.)
# #         nta_num = base_tower.replace("NTA", "").strip()
# #         # Pad single digit with zero
# #         if len(nta_num) == 1:
# #             nta_num = "0" + nta_num
# #         nta_key = f"NTA {nta_num}"
        
# #         if nta_key in NTA_ROW_RANGES:
# #             row_start, row_end = NTA_ROW_RANGES[nta_key]
# #             row_end += 1  # Make it inclusive
# #             logger.debug(f"    Using NTA row range: {row_start}-{row_end-1}")
    
# #     # Split activity text into hierarchy levels - handle both newline and comma separation
# #     activity_lines = []
# #     if '\n' in activity_text:
# #         # Newline-separated hierarchy (sub-activities)
# #         activity_lines = [line.strip() for line in activity_text.split('\n') if line.strip()]
# #     elif ',' in activity_text:
# #         # Comma-separated hierarchy (single-line activities)
# #         activity_lines = [line.strip() for line in activity_text.split(',') if line.strip()]
# #     else:
# #         # Single term
# #         activity_lines = [activity_text.strip()] if activity_text.strip() else []
    
# #     if not activity_lines:
# #         return None
    
# #     logger.debug(f"    Searching in {tower_sheet.title}")
# #     logger.debug(f"    Hierarchy: {' → '.join(activity_lines)}")
    
# #     # Strategy: Find the PARENT level first, then find the CHILD within the next 10 rows
# #     parent_term = normalize_text(activity_lines[0])
    
# #     # Define conflicting terms for the parent
# #     conflicting_terms = []
# #     if 'upper' in parent_term:
# #         conflicting_terms.append('lower')
# #     elif 'lower' in parent_term:
# #         conflicting_terms.append('upper')
# #     elif 'ground' in parent_term:
# #         conflicting_terms.extend(['1st', '2nd', '3rd', '4th'])
# #     elif '1st' in parent_term:
# #         conflicting_terms.extend(['ground', '2nd', '3rd', '4th'])
# #     elif '2nd' in parent_term:
# #         conflicting_terms.extend(['ground', '1st', '3rd', '4th'])
# #     elif '3rd' in parent_term:
# #         conflicting_terms.extend(['ground', '1st', '2nd', '4th'])
    
# #     # Special handling for finishing work keywords
# #     if any(term in parent_term for term in ['finishing', 'paint', 'plastering', 'false ceiling', 'flooring', 'tiles', 'fixtures']):
# #         conflicting_terms.extend(['structure', 'rcc', 'concrete', 'casting', 'shuttering', 'reinforcement'])
    
# #     logger.debug(f"    Parent term: '{parent_term}'")
# #     logger.debug(f"    Conflicting terms: {conflicting_terms}")
    
# #     # STEP 1: Find the parent row (within row range constraints)
# #     parent_row = None
# #     for row_idx in range(row_start, row_end):
# #         task_name = tower_sheet.cell(row_idx, TRACKER_TASK_NAME_COL).value
        
# #         if not task_name:
# #             continue
        
# #         task_normalized = normalize_text(task_name)
        
# #         # Check if this row matches the parent
# #         if parent_term in task_normalized:
# #             # Make sure it doesn't contain conflicting terms
# #             has_conflict = any(conflict in task_normalized for conflict in conflicting_terms)
            
# #             if not has_conflict:
# #                 parent_row = row_idx
# #                 logger.debug(f"    Found parent at row {row_idx}: {task_name.strip()}")
# #                 break
    
# #     if not parent_row:
# #         logger.debug(f"    Parent '{parent_term}' not found")
# #         return None
    
# #     # STEP 2: If we have child levels, search within next 10 rows
# #     if len(activity_lines) > 1:
# #         child_terms = [normalize_text(line) for line in activity_lines[1:]]
# #         logger.debug(f"    Child terms: {child_terms}")
        
# #         best_match = None
# #         best_match_score = 0
# #         best_match_row = None
        
# #         # Search within next 10 rows after parent (but not beyond row_end)
# #         for row_idx in range(parent_row + 1, min(parent_row + 11, row_end)):
# #             task_name = tower_sheet.cell(row_idx, TRACKER_TASK_NAME_COL).value
            
# #             if not task_name:
# #                 continue
            
# #             task_normalized = normalize_text(task_name)
            
# #             # Calculate match score - prioritize matching deeper (later) terms
# #             # Strongly prefer exact matches over partial matches
# #             match_count = 0
# #             for idx, term in enumerate(child_terms):
# #                 if term in task_normalized:
# #                     # Base score for matching this term (later terms get higher scores)
# #                     term_score = (idx + 1)
                    
# #                     # Strong bonus for exact match (term equals the whole task name)
# #                     if term == task_normalized:
# #                         term_score += 100  # Very high bonus for exact match
# #                     # Medium bonus if term is a major portion (>70%) of task name
# #                     elif len(term) >= len(task_normalized) * 0.7:
# #                         term_score += 20
# #                     # Small penalty if task name has many extra words beyond our term
# #                     elif len(task_normalized) > len(term) * 1.5:
# #                         term_score -= 5  # Penalize if tracker has significantly more words
                    
# #                     match_count += term_score
            
# #             # Only update if we have a BETTER match
# #             if match_count > best_match_score:
# #                 pct_value = tower_sheet.cell(row_idx, TRACKER_PCT_COMPLETE_COL).value
                
# #                 if pct_value is not None:
# #                     try:
# #                         if isinstance(pct_value, (int, float)):
# #                             pct_complete = float(pct_value) * 100
# #                         else:
# #                             pct_complete = float(str(pct_value).replace('%', ''))
                        
# #                         if 0 <= pct_complete <= 100:
# #                             best_match_score = match_count
# #                             best_match = pct_complete
# #                             best_match_row = row_idx
# #                             logger.debug(f"    Child match at row {row_idx}: {task_name.strip()[:50]} = {pct_complete:.1f}% (score: {match_count})")
# #                     except (ValueError, TypeError):
# #                         pass
        
# #         if best_match is not None:
# #             logger.debug(f"    ✓ SELECTED: row {best_match_row}, {best_match:.1f}%")
# #             return best_match
# #         else:
# #             logger.debug(f"    Child terms not found within 10 rows of parent")
# #             return None
    
# #     else:
# #         # Single level activity - return parent row's % Complete
# #         pct_value = tower_sheet.cell(parent_row, TRACKER_PCT_COMPLETE_COL).value
        
# #         if pct_value is not None:
# #             try:
# #                 if isinstance(pct_value, (int, float)):
# #                     pct_complete = float(pct_value) * 100
# #                 else:
# #                     pct_complete = float(str(pct_value).replace('%', ''))
                
# #                 if 0 <= pct_complete <= 100:
# #                     logger.debug(f"    ✓ SELECTED parent row: {pct_complete:.1f}%")
# #                     return pct_complete
# #             except (ValueError, TypeError):
# #                 pass
        
# #         return None

# # # ======================= REPORT GENERATION =======================

# # def sort_towers(tower_name: str) -> tuple:
# #     """Custom sort key ensuring proper order of all milestone types"""
# #     tower_lower = tower_name.lower()
    
# #     # Priority 0: Regular Towers (Structure Work)
# #     if tower_name.startswith('Tower') and 'finishing' not in tower_lower:
# #         match = re.search(r'Tower\s*(\d+)', tower_name)
# #         if match:
# #             return (0, int(match.group(1)), tower_name)
# #         return (0, 999, tower_name)
    
# #     # Priority 1: Regular NTAs (Structure Work) - Must NOT be in finishing section
# #     elif tower_name.startswith('NTA') and 'finishing' not in tower_lower and 'work' not in tower_lower:
# #         # This catches only "NTA 01", "NTA 02" that are structure work
# #         match = re.search(r'NTA\s*(\d+)', tower_name)
# #         if match:
# #             return (1, int(match.group(1)), tower_name)
# #         return (1, 999, tower_name)
    
# #     # Priority 2: Tower Finishing Work
# #     elif 'tower' in tower_lower and 'finishing' in tower_lower:
# #         match = re.search(r'Tower\s*(\d+)', tower_name, re.IGNORECASE)
# #         if match:
# #             return (2, int(match.group(1)), tower_name)
# #         return (2, 999, tower_name)
    
# #     # Priority 3: NTA Finishing Work Milestone Section
# #     # Sub-priority 0: The header "NTA Finishing Work Milestone:"
# #     # Sub-priority 1-99: Individual NTAs "NTA 01 Finishing Work", "NTA 02 Finishing Work"...
# #     elif 'nta' in tower_lower and 'finishing' in tower_lower:
# #         # The header comes first
# #         if tower_name == "NTA Finishing Work Milestone":
# #             return (3, 0, tower_name)
# #         # Individual NTA Finishing Work entries
# #         # Match pattern: "NTA 01 Finishing Work", "NTA 02 Finishing Work", etc.
# #         match = re.search(r'NTA\s*(\d+)', tower_name, re.IGNORECASE)
# #         if match:
# #             nta_num = int(match.group(1))
# #             return (3, nta_num, tower_name)
# #         return (3, 999, tower_name)
    
# #     # Priority 4: External Development Work
# #     elif 'external' in tower_lower or 'development' in tower_lower:
# #         return (4, 0, tower_name)
    
# #     # Priority 5: Others
# #     else:
# #         return (5, 999, tower_name)

# # def generate_report(tower_targets: Dict[str, List[ActivityTarget]], 
# #                    tracker_workbooks: Dict[str, Any], months: List[str], year: int) -> pd.DataFrame:
# #     """
# #     Generate milestone report DataFrame.
# #     SPECIAL: "NTA Finishing Work Milestone" appears as a section header row with no data
# #     """
# #     logger.info("\n" + "="*70)
# #     logger.info("GENERATING REPORT")
# #     logger.info("="*70)
    
# #     report_rows = []
    
# #     # Sort towers
# #     sorted_tower_names = sorted(tower_targets.keys(), key=sort_towers)
    
# #     logger.info(f"\nSorted tower order:")
# #     for idx, tower in enumerate(sorted_tower_names, 1):
# #         logger.info(f"  {idx}. {tower} (Priority: {sort_towers(tower)})")
    
# #     for tower_name in sorted_tower_names:
# #         # Skip only the invalid "NTA" entry
# #         if tower_name.strip().upper() == "NTA":
# #             logger.info(f"\nSkipping: {tower_name} (invalid)")
# #             continue
        
# #         # SPECIAL CASE: "NTA Finishing Work Milestone" is a section header only
# #         if tower_name == "NTA Finishing Work Milestone":
# #             logger.info(f"\nAdding section header: {tower_name}")
            
# #             # Create a header row with colon appended to tower name
# #             header_row = {'Tower': f"{tower_name}:"}  # Add colon here
            
# #             for month in months:
# #                 header_row[f"Activity- {month} {year}"] = ""
# #                 header_row[f"% Complete- {month}"] = ""
# #                 header_row[f"Status- {month}"] = ""
# #                 header_row[f"Weightage- {month}"] = ""
# #                 header_row[f"Weighted %- {month}"] = ""
            
# #             header_row[f"Target till {months[-1]}"] = ""
# #             header_row['Responsible'] = ""
# #             header_row['Delay Reason'] = ""
            
# #             report_rows.append(header_row)
# #             continue
        
# #         logger.info(f"\nProcessing: {tower_name}")
        
# #         row_data = {'Tower': tower_name}
        
# #         # Process each month
# #         for month in months:
# #             month_targets = [t for t in tower_targets[tower_name] if t.month == month]
# #             tracker_wb = tracker_workbooks.get(month)
            
# #             if not month_targets:
# #                 # No targets for this month
# #                 row_data[f"Activity- {month} {year}"] = ""
# #                 row_data[f"% Complete- {month}"] = ""
# #                 row_data[f"Status- {month}"] = ""
# #                 row_data[f"Weightage- {month}"] = ""
# #                 row_data[f"Weighted %- {month}"] = ""
# #                 continue
            
# #             # We have targets
# #             activities_text = "\n".join([t.activity_text for t in month_targets])
# #             row_data[f"Activity- {month} {year}"] = activities_text
            
# #             if tracker_wb:
# #                 total_actual = 0
# #                 matched = 0
                
# #                 for target in month_targets:
# #                     actual_pct = find_activity_in_tracker(tracker_wb, tower_name, target.activity_text, month)
                    
# #                     if actual_pct is not None:
# #                         # If actual meets or exceeds target, show 100%
# #                         if actual_pct >= target.target_pct:
# #                             target.actual_pct = 100.0
# #                             target.status = "Achieved"
# #                             matched += 1
# #                         else:
# #                             # Below target - show actual percentage
# #                             target.actual_pct = actual_pct
# #                             target.status = "Not Matched"
                        
# #                         logger.info(f"  {month}: {target.activity_text[:40]} = {target.actual_pct:.0f}%")
# #                         total_actual += target.actual_pct
# #                     else:
# #                         target.status = "Not Found"
                
# #                 avg_actual = total_actual / len(month_targets) if month_targets else 0
                
# #                 if matched == len(month_targets) and matched > 0:
# #                     status = "Achieved"
# #                 elif matched > 0:
# #                     status = "Partial"
# #                 else:
# #                     status = "Not Achieved"
                
# #                 row_data[f"% Complete- {month}"] = f"{avg_actual:.0f}%"
# #                 row_data[f"Status- {month}"] = status
                
# #                 # Weightage is 100 for each month
# #                 weightage = 100
# #                 weighted_pct = (avg_actual / 100) * weightage
# #                 row_data[f"Weightage- {month}"] = weightage
# #                 row_data[f"Weighted %- {month}"] = f"{weighted_pct:.1f}%"
# #             else:
# #                 row_data[f"% Complete- {month}"] = ""
# #                 row_data[f"Status- {month}"] = ""
# #                 row_data[f"Weightage- {month}"] = ""
# #                 row_data[f"Weighted %- {month}"] = ""
        
# #         # Summary columns
# #         last_month = months[-1]
# #         last_targets = [t for t in tower_targets[tower_name] if t.month == last_month]
# #         row_data[f"Target till {last_month}"] = "\n".join([t.activity_text for t in last_targets])
        
# #         row_data['Responsible'] = ""
# #         row_data['Delay Reason'] = ""
        
# #         report_rows.append(row_data)
    
# #     # Add summary row
# #     summary_row = {'Tower': 'AVERAGE WEIGHTED %'}
    
# #     for month in months:
# #         weighted_values = []
# #         for row in report_rows:
# #             # Skip the NTA Finishing Work Milestone header row in calculations
# #             if row['Tower'] == "NTA Finishing Work Milestone:" or row['Tower'] == "NTA Finishing Work Milestone":
# #                 continue
                
# #             weighted_val = row.get(f"Weighted %- {month}", "")
# #             if weighted_val and weighted_val != "":
# #                 try:
# #                     val = float(str(weighted_val).replace('%', ''))
# #                     weighted_values.append(val)
# #                 except (ValueError, TypeError):
# #                     pass
        
# #         if weighted_values:
# #             avg_weighted = sum(weighted_values) / len(weighted_values)
# #             summary_row[f"Weighted %- {month}"] = f"{avg_weighted:.1f}%"
# #         else:
# #             summary_row[f"Weighted %- {month}"] = ""
        
# #         summary_row[f"Activity- {month} {year}"] = ""
# #         summary_row[f"% Complete- {month}"] = ""
# #         summary_row[f"Status- {month}"] = ""
# #         summary_row[f"Weightage- {month}"] = ""
    
# #     summary_row[f"Target till {months[-1]}"] = ""
# #     summary_row['Responsible'] = ""
# #     summary_row['Delay Reason'] = ""
    
# #     report_rows.append(summary_row)
    
# #     return pd.DataFrame(report_rows)

# # def format_report(worksheet, dataframe):
# #     """Apply formatting to report."""
# #     header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
# #     header_font = Font(bold=True, color="FFFFFF")
# #     summary_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
# #     summary_font = Font(bold=True, size=11)
    
# #     # Format title
# #     worksheet.cell(1, 1).font = Font(bold=True, size=14)
# #     worksheet.cell(2, 1).font = Font(italic=True, size=10)
    
# #     # Format headers
# #     for col in range(1, worksheet.max_column + 1):
# #         cell = worksheet.cell(4, col)
# #         cell.fill = header_fill
# #         cell.font = header_font
# #         cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
# #     # Format data
# #     thin_border = Border(
# #         left=Side(style='thin'), right=Side(style='thin'),
# #         top=Side(style='thin'), bottom=Side(style='thin')
# #     )
    
# #     # Last row is the summary row
# #     summary_row_idx = worksheet.max_row
    
# #     for row in range(5, worksheet.max_row + 1):
# #         is_summary_row = (row == summary_row_idx)
        
# #         for col in range(1, worksheet.max_column + 1):
# #             cell = worksheet.cell(row, col)
# #             cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
# #             cell.border = thin_border
            
# #             # Special formatting for summary row
# #             if is_summary_row:
# #                 cell.fill = summary_fill
# #                 cell.font = summary_font
# #                 cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
# #     # Column widths
# #     for col_idx, column in enumerate(dataframe.columns, start=1):
# #         col_letter = get_column_letter(col_idx)
# #         if 'Activity' in column or 'Target' in column:
# #             worksheet.column_dimensions[col_letter].width = 40
# #         else:
# #             worksheet.column_dimensions[col_letter].width = 15
    
# #     worksheet.row_dimensions[4].height = 50
# #     worksheet.row_dimensions[summary_row_idx].height = 30  # Make summary row slightly taller

# # # ======================= MAIN =======================

# # def main():
# #     """Main execution."""
# #     try:
# #         logger.info("\n" + "="*70)
# #         logger.info("MILESTONE REPORT GENERATOR v2.0")
# #         logger.info("="*70)
        
# #         # Step 1: Find KRA
# #         logger.info("\nSTEP 1: Finding latest KRA file")
# #         cos = init_cos()
        
# #         kra_result = find_latest_kra_file(cos, BUCKET, KRA_FOLDER)
# #         if not kra_result:
# #             logger.error("Could not find KRA file")
# #             return
        
# #         kra_key, quarter_months, kra_year = kra_result
        
# #         # Step 2: Load KRA
# #         logger.info("\nSTEP 2: Loading KRA and parsing targets")
# #         kra_bytes = download_file_bytes(cos, kra_key)
# #         kra_wb = load_workbook(filename=BytesIO(kra_bytes), data_only=True)
        
# #         kra_ws = find_project_sheet(kra_wb, "EDEN")
# #         if not kra_ws:
# #             logger.error("EDEN sheet not found")
# #             return
        
# #         tower_targets = parse_kra_targets(kra_ws)
        
# #         if not tower_targets:
# #             logger.error("No targets found in KRA")
# #             return
        
# #         # Step 3: Load all available trackers
# #         logger.info("\nSTEP 3: Loading tracker files")
# #         tracker_workbooks = {}
        
# #         for month in quarter_months:
# #             tracker_month_num = MONTH_TO_TRACKER_MAPPING.get(month)
# #             if not tracker_month_num:
# #                 logger.warning(f"  No mapping for {month}")
# #                 continue
            
# #             tracker_year = calculate_tracker_year(month, kra_year)
# #             logger.info(f"\n  {month} data requires tracker: {tracker_month_num:02d}/{tracker_year}")
            
# #             tracker_key = find_tracker_for_month(cos, BUCKET, tracker_month_num, tracker_year, EDEN_TRACKER_FOLDER)
            
# #             if tracker_key:
# #                 logger.info(f"    ✓ Found: {os.path.basename(tracker_key)}")
# #                 tracker_bytes = download_file_bytes(cos, tracker_key)
# #                 tracker_wb = load_workbook(filename=BytesIO(tracker_bytes), data_only=True)
# #                 tracker_workbooks[month] = tracker_wb
# #                 logger.info(f"    ✓ Loaded with sheets: {tracker_wb.sheetnames}")
# #             else:
# #                 logger.warning(f"    ✗ Not found - {month} column will show activities but no completion data")
        
# #         logger.info(f"\n  Summary: {len(tracker_workbooks)}/{len(quarter_months)} trackers loaded")
        
# #         # Step 4: Generate report
# #         logger.info("\nSTEP 4: Generating report")
# #         report_df = generate_report(tower_targets, tracker_workbooks, quarter_months, kra_year)
        
# #         # Step 5: Save
# #         logger.info("\nSTEP 5: Saving report")
# #         output_file = f"Eden_Milestone_Report_{'_'.join(quarter_months)}_{kra_year}.xlsx"
        
# #         wb = Workbook()
# #         ws = wb.active
# #         ws.title = "Progress Report"
        
# #         ws.append(["Eden- Progress Against Milestones"])
# #         ws.append([f"Report Generated: {datetime.now().strftime('%B %d, %Y')}"])
# #         ws.append([])
        
# #         for r in dataframe_to_rows(report_df, index=False, header=True):
# #             ws.append(r)
        
# #         format_report(ws, report_df)
# #         wb.save(output_file)
        
# #         logger.info(f"\n{'='*70}")
# #         logger.info("REPORT COMPLETE")
# #         logger.info(f"{'='*70}")
# #         logger.info(f"File: {output_file}")
# #         logger.info(f"Towers: {len(report_df)}")
# #         logger.info(f"Months with data: {list(tracker_workbooks.keys())}")
# #         logger.info(f"Months with blank data: {[m for m in quarter_months if m not in tracker_workbooks]}")
# #         logger.info(f"{'='*70}\n")
        
# #     except Exception as e:
# #         logger.error(f"Error: {str(e)}", exc_info=True)
# #         raise

# # if __name__ == "__main__":
# #     main()


















import streamlit as st
import subprocess
import sys
import os
import time
import gc
import psutil
from datetime import datetime
import glob
import traceback

# Page configuration
st.set_page_config(
    page_title="Milestone Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Enhanced Custom CSS for modern, visually appealing interface
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global styles */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        font-family: 'Inter', sans-serif;
    }
    
    /* Hide Streamlit default elements - but keep the page title tab */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    div[data-testid="stToolbar"] {visibility: hidden;}
    .stDeployButton {display: none;}
    div[data-testid="stDecoration"] {display: none;}
    
    /* Main container styling */
    .main-container {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        padding: 2rem;
        margin: 2rem auto;
        max-width: 1200px;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
        border: 1px solid rgba(255, 255, 255, 0.3);
    }
    
    /* Title styling */
    .main-title {
        text-align: center;
        font-size: 3rem;
        font-weight: 900;
        color: #000000 !important;
        margin-bottom: 0.5rem;
        text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    .subtitle {
        text-align: center;
        font-size: 1.2rem;
        color: #000000 !important;
        margin-bottom: 2rem;
        font-weight: 700;
    }
    
    /* Chat message styling */
    .chat-message {
        padding: 1.5rem;
        border-radius: 15px;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: flex-start;
        animation: fadeInUp 0.3s ease-out;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
    }
    
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    .chat-message.bot {
        background: linear-gradient(135deg, #f8f9ff, #e8f0fe);
        margin-right: 2rem;
        border-left: 4px solid #667eea;
    }
    
    .chat-message.user {
        background: linear-gradient(135deg, #e3f2fd, #f1f8e9);
        margin-left: 2rem;
        flex-direction: row-reverse;
        border-right: 4px solid #4caf50;
    }
    
    .chat-message .avatar {
        width: 3.5rem;
        height: 3.5rem;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.5rem;
        margin: 0 1rem;
        background: linear-gradient(135deg, #667eea, #764ba2);
        color: white;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
    }
    
    .user .avatar {
        background: linear-gradient(135deg, #4caf50, #45a049);
        box-shadow: 0 4px 15px rgba(76, 175, 80, 0.3);
    }
    
    .chat-message .message {
        flex: 1;
        padding: 0 0.5rem;
        font-size: 1.1rem;
        line-height: 1.6;
        color: #000000 !important;
        font-weight: 600;
    }
    
    /* Project selection section */
    .project-selection {
        background: linear-gradient(135deg, #f8f9ff, #fff);
        border-radius: 20px;
        padding: 2rem;
        margin: 2rem 0;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.08);
        border: 1px solid rgba(102, 126, 234, 0.1);
    }
    
    .project-selection h3 {
        text-align: center;
        color: #000000 !important;
        font-size: 1.8rem;
        font-weight: 800;
        margin-bottom: 2rem;
    }
    
    /* Project buttons */
    .project-buttons {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1.5rem;
        margin: 2rem 0;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #667eea, #764ba2) !important;
        color: white !important;
        border: none !important;
        padding: 1.2rem 2rem !important;
        border-radius: 15px !important;
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 8px 25px rgba(102, 126, 234, 0.3) !important;
        width: 100% !important;
        height: 70px !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-5px) !important;
        box-shadow: 0 15px 35px rgba(102, 126, 234, 0.4) !important;
        background: linear-gradient(135deg, #764ba2, #667eea) !important;
    }
    
    .stButton > button:active {
        transform: translateY(-2px) !important;
    }
    
    /* Status containers */
    .status-container {
        background: linear-gradient(135deg, #fff3cd, #ffeaa7);
        border: 2px solid #f39c12;
        border-radius: 15px;
        padding: 2rem;
        margin: 2rem 0;
        text-align: center;
        box-shadow: 0 8px 25px rgba(243, 156, 18, 0.2);
    }
    
    .status-container h4 {
        color: #d68910;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .status-container p {
        color: #7d6608;
        font-size: 1.1rem;
        margin-bottom: 0;
    }
    
    .success-container {
        background: linear-gradient(135deg, #d4edda, #c3e6cb);
        border: 2px solid #27ae60;
        border-radius: 15px;
        padding: 2rem;
        margin: 2rem 0;
        text-align: center;
        box-shadow: 0 8px 25px rgba(39, 174, 96, 0.2);
    }
    
    .success-container h4 {
        color: #27ae60;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .success-container p {
        color: #155724;
        font-size: 1.1rem;
        margin-bottom: 0;
    }
    
    .error-container {
        background: linear-gradient(135deg, #f8d7da, #f5c6cb);
        border: 2px solid #e74c3c;
        border-radius: 15px;
        padding: 2rem;
        margin: 2rem 0;
        text-align: center;
        box-shadow: 0 8px 25px rgba(231, 76, 60, 0.2);
    }
    
    .error-container h4 {
        color: #c0392b;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .error-container p {
        color: #721c24;
        font-size: 1.1rem;
        margin-bottom: 0;
    }
    
    /* Download button styling */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #27ae60, #2ecc71) !important;
        color: white !important;
        border: none !important;
        padding: 1.5rem 3rem !important;
        border-radius: 15px !important;
        font-size: 1.2rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 8px 25px rgba(39, 174, 96, 0.3) !important;
        width: 100% !important;
        height: 80px !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 12px 30px rgba(39, 174, 96, 0.4) !important;
        background: linear-gradient(135deg, #2ecc71, #27ae60) !important;
    }
    
    /* Progress bar styling */
    .stProgress > div > div > div > div {
        background: linear-gradient(135deg, #667eea, #764ba2) !important;
        border-radius: 10px !important;
    }
    
    .stProgress > div > div > div {
        background-color: rgba(102, 126, 234, 0.1) !important;
        border-radius: 10px !important;
        height: 15px !important;
    }
    
    /* Footer styling */
    .footer {
        text-align: center;
        color: rgba(255, 255, 255, 0.8);
        font-size: 1rem;
        margin-top: 3rem;
        padding: 2rem;
        background: rgba(255, 255, 255, 0.1);
        border-radius: 15px;
        backdrop-filter: blur(10px);
    }
    
    /* Clear memory button styling */
    .clear-memory-btn {
        background: linear-gradient(135deg, #e74c3c, #c0392b) !important;
        color: white !important;
        border: none !important;
        padding: 0.8rem 1.5rem !important;
        border-radius: 10px !important;
        font-size: 0.9rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 15px rgba(231, 76, 60, 0.3) !important;
        width: 100% !important;
        height: 50px !important;
    }
    
    /* Divider */
    hr {
        border: none;
        height: 2px;
        background: linear-gradient(135deg, #667eea, #764ba2);
        margin: 2rem 0;
        border-radius: 2px;
    }
    
    /* Custom scrollbar */
    ::-webkit-scrollbar {
        width: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, #764ba2, #667eea);
    }
    
    /* Responsive design */
    @media (max-width: 768px) {
        .main-title {
            font-size: 2rem;
        }
        
        .main-container {
            margin: 1rem;
            padding: 1rem;
        }
        
        .chat-message {
            margin: 0.5rem 0;
        }
        
        .chat-message.bot {
            margin-right: 0.5rem;
        }
        
        .chat-message.user {
            margin-left: 0.5rem;
        }
    }
    
    /* Debug info styling */
    .debug-info {
        background: linear-gradient(135deg, #f0f4f8, #e2e8f0);
        border: 2px solid #718096;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
        font-family: monospace;
        font-size: 0.9rem;
        color: #2d3748;
    }
    
    /* System info styling */
    .system-info {
        background: linear-gradient(135deg, #e8f5e8, #f0f8f0);
        border: 2px solid #27ae60;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
        font-size: 0.9rem;
        color: #155724;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'messages' not in st.session_state:
    st.session_state.messages = []
    st.session_state.stage = 'welcome'
    st.session_state.selected_project = None
    st.session_state.report_file = None

# Project configurations - Enhanced with better debugging
PROJECTS = {
    'Veridia': {
        'script': 'veridia.py',
        'display_name': 'Veridia',
        'icon': '🌿',
        'patterns': [
            'Time_Delivery_Milestones_Report_*.xlsx',
            '*Veridia*.xlsx',
            'Veridia_*.xlsx',
            '*veridia*.xlsx'
        ]
    },
    'Eligo': {
        'script': 'eligo.py', 
        'display_name': 'Eligo',
        'icon': '⚡',
        'patterns': [
            '*Eligo*.xlsx',
            'Eligo_*.xlsx',
            '*eligo*.xlsx'
        ]
    },
    'EWS-LIG': {
        'script': 'ews-lig.py',
        'display_name': 'EWS-LIG',
        'icon': '🔍',
        'patterns': [
            '*EWS*LIG*.xlsx',
            '*EWS-LIG*.xlsx',
            'EWS_LIG_*.xlsx',
            '*ews*lig*.xlsx'
        ]
    },
    'WaveCityClub': {
        'script': 'wavecityclub.py',
        'display_name': 'WaveCityClub',
        'icon': '🌊',
        'patterns': [
            'Wave_City_Club_Report_*.xlsx',
            '*WaveCityClub*.xlsx',
            '*Wave*City*Club*.xlsx',
            '*wavecityclub*.xlsx'
        ]
    },
    'Eden': {
        'script': 'eden.py',
        'display_name': 'Eden',
        'icon': '🏡',
        'patterns': [
            'Eden_KRA_Milestone_Report_*.xlsx',
            '*Eden*.xlsx',
            'Eden_*.xlsx',
            '*eden*.xlsx'
        ]
    }
}

def cleanup_resources():
    """Clean up system resources between script executions"""
    try:
        st.write("🧹 **Cleaning up system resources...**")
        
        # Force garbage collection
        gc.collect()
        
        # Kill any orphaned Python processes (be careful with this)
        current_pid = os.getpid()
        killed_processes = 0
        
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                if proc.info['name'] in ['python', 'python.exe']:
                    # Check if it's a subprocess of our scripts
                    cmdline = proc.info['cmdline'] or []
                    script_names = ['veridia.py', 'eligo.py', 'ews-lig.py', 'wavecityclub.py', 'eden.py']
                    if any(script in ' '.join(cmdline) for script in script_names):
                        if proc.info['pid'] != current_pid:
                            proc.terminate()
                            proc.wait(timeout=3)
                            killed_processes += 1
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
                continue
        
        if killed_processes > 0:
            st.write(f"🔄 Terminated {killed_processes} orphaned script processes")
        
        # Clear any temporary files that might be locked
        temp_patterns = ['~$*.xlsx', '*.tmp', '.~lock.*', '__pycache__']
        cleaned_files = 0
        
        for pattern in temp_patterns:
            for file in glob.glob(pattern):
                try:
                    if os.path.isfile(file):
                        os.remove(file)
                        cleaned_files += 1
                    elif os.path.isdir(file):
                        import shutil
                        shutil.rmtree(file)
                        cleaned_files += 1
                except:
                    pass
        
        if cleaned_files > 0:
            st.write(f"🗑️ Cleaned {cleaned_files} temporary files")
        
        # Brief pause to let system settle
        time.sleep(2)
        
        # Show memory status after cleanup
        memory_info = psutil.virtual_memory()
        st.write(f"💾 **Memory after cleanup:** {memory_info.percent:.1f}% used ({memory_info.available / (1024**3):.1f} GB available)")
        st.success("✅ Resource cleanup completed!")
        
    except Exception as e:
        st.write(f"⚠️ Resource cleanup warning: {e}")

def add_message(role, content):
    """Add a message to the chat history"""
    st.session_state.messages.append({
        'role': role,
        'content': content,
        'timestamp': datetime.now()
    })

def display_chat_message(message):
    """Display a single chat message"""
    role = message['role']
    content = message['content']
    
    if role == 'bot':
        st.markdown(f"""
        <div class="chat-message bot">
            <div class="avatar">🤖</div>
            <div class="message">{content}</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="chat-message user">
            <div class="avatar">👤</div>
            <div class="message">{content}</div>
        </div>
        """, unsafe_allow_html=True)

def find_generated_file(project_config, project_name):
    """Find the generated report file using multiple patterns"""
    patterns = project_config['patterns']
    
    for pattern in patterns:
        st.write(f"🔍 Searching with pattern: {pattern}")
        matches = glob.glob(pattern)
        if matches:
            # Get the most recent file
            latest_file = max(matches, key=os.path.getctime)
            file_time = os.path.getctime(latest_file)
            current_time = time.time()
            
            # Check if file was created recently (within last 10 minutes)
            if (current_time - file_time) < 1200:  # 20 minutes
                st.write(f"✅ Found recent file: {latest_file}")
                return latest_file
            else:
                st.write(f"⏰ File found but too old: {latest_file}")
    
    # Check for any new Excel files created
    all_excel = glob.glob("*.xlsx")
    if all_excel:
        latest_new_file = max(all_excel, key=os.path.getctime)
        file_time = os.path.getctime(latest_new_file)
        current_time = time.time()
        
        if (current_time - file_time) < 1200:  # 20 minutes
            st.write(f"📄 Found recent Excel file: {latest_new_file}")
            return latest_new_file
    
    return None

def monitor_memory_during_execution():
    """Monitor and display current system status"""
    try:
        memory_info = psutil.virtual_memory()
        cpu_percent = psutil.cpu_percent(interval=1)
        
        st.markdown(f"""
        <div class="system-info">
            <strong>💻 System Status:</strong><br>
            🧠 Memory: {memory_info.percent:.1f}% used ({memory_info.available / (1024**3):.1f} GB available)<br>
            ⚡ CPU: {cpu_percent:.1f}% usage<br>
            🔧 Active Python processes: {len([p for p in psutil.process_iter() if 'python' in p.name().lower()])}
        </div>
        """, unsafe_allow_html=True)
        
        return memory_info.percent
        
    except Exception as e:
        st.write(f"⚠️ Could not monitor system status: {e}")
        return 0

def run_project_script(project_name):
    """Enhanced script execution with proper resource management"""
    try:
        project_config = PROJECTS[project_name]
        script_path = project_config['script']
        
        # Show system status before execution
        memory_before = monitor_memory_during_execution()
        
        # Enhanced debugging information
        st.write(f"🔧 **Debug Information for {project_name}:**")
        st.write(f"📝 Script path: {script_path}")
        st.write(f"📁 Current directory: {os.getcwd()}")
        st.write(f"🐍 Python executable: {sys.executable}")
        
        # Check if script file exists
        if not os.path.exists(script_path):
            available_files = [f for f in os.listdir('.') if f.endswith('.py')]
            return False, f"❌ Script file '{script_path}' not found in current directory.\n\nAvailable Python files: {available_files}"
        
        st.write(f"✅ Script file found: {script_path}")
        
        # Store existing Excel files before execution
        files_before = set(glob.glob("*.xlsx"))
        st.write(f"📊 Excel files before execution: {len(files_before)}")
        
        # Enhanced timeout settings
        timeout_settings = {
            'Veridia': 1200,      # 15 minutes for Veridia
            'Eligo': 1200,        # 20 minutes for Eligo  
            'EWS-LIG': 1200,      # 20 minutes
            'WaveCityClub': 450, # 7.5 minutes
            'Eden': 450          # 7.5 minutes
        }
        timeout_duration = timeout_settings.get(project_name, 300)
        
        st.write(f"🚀 Executing script: {script_path} (timeout: {timeout_duration//60} minutes)")
        
        # Create enhanced environment
        env = os.environ.copy()
        env.update({
            'PYTHONUNBUFFERED': '1',
            'MPLBACKEND': 'Agg',
            'OPENBLAS_NUM_THREADS': '1',  # Limit BLAS threads
            'MKL_NUM_THREADS': '1',       # Limit MKL threads
            'NUMEXPR_NUM_THREADS': '1',   # Limit NumExpr threads
            'OMP_NUM_THREADS': '1',       # Limit OpenMP threads
            'PYTHONDONTWRITEBYTECODE': '1',  # Don't create .pyc files
            'PYTHONHASHSEED': '0',           # Consistent hashing
        })
        
        start_time = time.time()
        progress_placeholder = st.empty()
        progress_placeholder.info(f"⏱️ Running {project_name} script... (Max wait: {timeout_duration//60} minutes)")
        
        # Use Popen for better process control
        process = subprocess.Popen(
            [sys.executable, script_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=os.getcwd(),
            env=env,
            bufsize=1,
            universal_newlines=True
        )
        
        try:
            # Wait for completion with timeout
            stdout, stderr = process.communicate(timeout=timeout_duration)
            elapsed_time = time.time() - start_time
            
            if process.returncode == 0:
                progress_placeholder.success(f"✅ Script completed in {elapsed_time:.1f} seconds")
            else:
                progress_placeholder.error(f"❌ Script failed with return code {process.returncode}")
                
        except subprocess.TimeoutExpired:
            # Handle timeout more gracefully
            elapsed_time = time.time() - start_time
            progress_placeholder.error(f"⏱️ Script timed out after {elapsed_time:.1f} seconds")
            
            # Terminate the process
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait()
            
            # Clean up any remaining processes
            st.write("🧹 **Performing emergency cleanup...**")
            cleanup_resources()
            
            timeout_msg = f"""
⏱️ **{project_name} script timed out after {timeout_duration//60} minutes.**

**This could indicate:**
1. **Large dataset processing** - The script is processing very large files
2. **Memory issues** - System running out of memory (was {memory_before:.1f}% before execution)
3. **Infinite loop** - Bug in the script causing it to loop
4. **Resource contention** - Previous script execution interfering

**Recommended actions:**
1. Click "Clear Memory" button before trying again
2. Check system resources (memory/CPU usage)
3. Try running `{script_path}` manually to identify the bottleneck
4. Consider restarting the Streamlit app to clear all resources
5. Break large input files into smaller chunks if applicable

**To debug manually:**
```bash
cd {os.getcwd()}
python {script_path}
```

This will show you exactly where the script stops or gets stuck.
            """
            return False, timeout_msg
        
        # Show memory status after execution
        memory_after = monitor_memory_during_execution()
        memory_change = memory_after - memory_before
        if memory_change > 10:
            st.warning(f"⚠️ Significant memory increase: +{memory_change:.1f}%")
        
        # Enhanced result logging
        st.write(f"📤 Script execution completed with return code: {process.returncode}")
        
        if stdout:
            st.write("📄 **Script Output (stdout):**")
            st.code(stdout[:1000] + ("..." if len(stdout) > 1000 else ""))
        
        if stderr:
            st.write("⚠️ **Script Errors (stderr):**")
            st.code(stderr[:1000] + ("..." if len(stderr) > 1000 else ""))
        
        # Check execution result
        if process.returncode != 0:
            error_details = f"""
Return Code: {process.returncode}
STDOUT: {stdout}
STDERR: {stderr}
            """
            return False, f"❌ Script execution failed with return code {process.returncode}.\n\nDetails:\n{error_details}"
        
        # Check for new files after execution
        files_after = set(glob.glob("*.xlsx"))
        new_files = files_after - files_before
        st.write(f"📊 Excel files after execution: {len(files_after)} (New: {len(new_files)})")
        
        if new_files:
            st.write(f"🆕 New Excel files created: {list(new_files)}")
        
        # Look for generated file
        generated_file = find_generated_file(project_config, project_name)
        
        if generated_file and os.path.exists(generated_file):
            file_size = os.path.getsize(generated_file)
            st.write(f"✅ **Report file found:** {generated_file} ({file_size:,} bytes)")
            return True, generated_file
        
        # Diagnostic information if file not found
        all_excel = glob.glob("*.xlsx")
        error_msg = f"""
❌ **Report file not found after script execution.**

**Diagnostics:**
- Script executed successfully (return code: {process.returncode})
- Patterns searched: {project_config['patterns']}
- All Excel files in directory: {all_excel}
- New files created: {list(new_files) if new_files else 'None'}

**Script Output:**
STDOUT: {stdout[:500]}...
STDERR: {stderr[:500]}...
        """
        return False, error_msg

    except Exception as e:
        cleanup_resources()  # Clean up on any error
        error_details = f"""
Exception Type: {type(e).__name__}
Exception Message: {str(e)}
Traceback: {traceback.format_exc()}
        """
        return False, f"❌ Unexpected error occurred:\n{error_details}"

def main():
    st.markdown('<div class="main-container">', unsafe_allow_html=True)

    # Title
    st.markdown("""
    <div class="main-title">📊 Milestone Report Generator</div>
    <div class="subtitle">Generate comprehensive milestone reports with just one click</div>
    """, unsafe_allow_html=True)
    
    # Add Clear Memory button if there's a selected project
    if st.session_state.get('selected_project') or st.session_state.stage != 'welcome':
        col1, col2 = st.columns([4, 1])
        with col2:
            if st.button("🧹 Clear Memory", help="Clear system resources and memory", key="clear_memory_btn"):
                with st.spinner("Clearing system resources..."):
                    cleanup_resources()
                    time.sleep(1)
                st.rerun()
    
    st.markdown("---")

    # Intro messages
    if not st.session_state.messages:
        add_message('bot', "Hello! 👋 Welcome to the Milestone Report Generator.")
        add_message('bot', "Which project would you like to generate a milestone report for?")

    # Display chat history
    for msg in st.session_state.messages:
        display_chat_message(msg)

    # Welcome / Project selection
    if st.session_state.stage == 'welcome' and not st.session_state.selected_project:
        st.markdown('<div class="project-selection"><h3>🚀 Select Your Project</h3></div>', unsafe_allow_html=True)
        
        # Display project buttons in a grid
        cols = st.columns(len(PROJECTS))
        for idx, (key, info) in enumerate(PROJECTS.items()):
            with cols[idx]:
                if st.button(f"{info['icon']} {info['display_name']}", key=key):
                    st.session_state.selected_project = key
                    st.session_state.stage = 'processing'
                    add_message('user', f"I want to generate a milestone report for {info['display_name']}.")
                    add_message('bot', f"Excellent choice! I'll generate the {info['display_name']} report now. Please wait...")
                    st.rerun()

    # Processing stage
    elif st.session_state.stage == 'processing':
        proj = st.session_state.selected_project
        info = PROJECTS[proj]
        
        st.markdown(f"""
        <div class="status-container">
          <h4>{info['icon']} Processing {info['display_name']}...</h4>
          <p>Please wait while I generate your report. This may take a few minutes.</p>
        </div>
        """, unsafe_allow_html=True)

        # Progress bar animation
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        progress_steps = [
            (0.2, "Initializing..."),
            (0.4, "Loading data..."),
            (0.6, "Processing calculations..."),
            (0.8, "Generating report..."),
            (1.0, "Finalizing...")
        ]
        
        for progress, step_text in progress_steps:
            status_text.text(step_text)
            progress_bar.progress(progress)
            time.sleep(0.8)
        
        # Clear progress indicators
        progress_bar.empty()
        status_text.empty()
        
        # Create a debug expander for detailed logging
        with st.expander("🔍 Debug Information", expanded=True):
            # Run the actual script
            success, result = run_project_script(proj)
        
        if success:
            st.session_state.report_file = result
            st.session_state.stage = 'completed'
            add_message('bot', f"✅ Your {info['display_name']} report has been generated successfully!")
        else:
            st.session_state.stage = 'error'
            st.session_state.error_message = result
            add_message('bot', f"❌ There was an error generating the {info['display_name']} report.")
        
        st.rerun()

    # Completed stage
    elif st.session_state.stage == 'completed':
        proj = st.session_state.selected_project
        info = PROJECTS[proj]
        
        st.markdown(f"""
        <div class="success-container">
          <h4>{info['icon']} Report Generated Successfully!</h4>
          <p>Your {info['display_name']} milestone report is ready to download.</p>
        </div>
        """, unsafe_allow_html=True)

        # Display file info
        if st.session_state.report_file and os.path.exists(st.session_state.report_file):
            file_size = os.path.getsize(st.session_state.report_file)
            file_size_mb = file_size / (1024 * 1024)
            st.info(f"📄 File: {os.path.basename(st.session_state.report_file)} ({file_size_mb:.2f} MB)")
            
            # Download button
            with open(st.session_state.report_file, "rb") as f:
                file_data = f.read()
            
            st.download_button(
                label=f"📥 Download {info['display_name']} Report",
                data=file_data,
                file_name=os.path.basename(st.session_state.report_file),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("Report file not found or was deleted.")

        # Generate another report button
        if st.button("🔄 Generate Another Report", use_container_width=True):
            st.session_state.messages = []
            st.session_state.stage = 'welcome'
            st.session_state.selected_project = None
            st.session_state.report_file = None
            st.rerun()

    # Error stage
    elif st.session_state.stage == 'error':
        proj = st.session_state.selected_project or ""
        info = PROJECTS.get(proj, {'display_name': 'Unknown', 'icon': '❌'})
        
        st.markdown(f"""
        <div class="error-container">
          <h4>{info['icon']} Error Generating {info['display_name']} Report</h4>
          <p>There was an issue generating your report. Please check the details below and try again.</p>
        </div>
        """, unsafe_allow_html=True)

        # Show error details
        if hasattr(st.session_state, 'error_message'):
            with st.expander("🔍 Detailed Error Information", expanded=True):
                st.code(st.session_state.error_message)

        # Troubleshooting tips
        with st.expander("💡 Troubleshooting Tips"):
            st.markdown(f"""
            **Common issues and solutions:**
            
            1. **Script file missing**: Ensure `{info.get('script', 'unknown.py')}` exists in the same directory as this Streamlit app.
            
            2. **Import errors**: Check if all required Python packages are installed.
            
            3. **Data file missing**: Ensure any required input data files are in the correct location.
            
            4. **Permission issues**: Check if the script has permission to write files to the current directory.
            
            5. **Path issues**: Verify that all file paths in the script are correct.
            
            6. **Memory/Resource issues**: Try clicking "Clear Memory" button and then retry.
            
            **Next steps:**
            - Try running `{info.get('script', 'unknown.py')}` manually from the command line
            - Check the script's dependencies and requirements
            - Verify input data files are present and accessible
            - Clear memory and restart if needed
            """)

        # Action buttons
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🧹 Clear & Retry", use_container_width=True, help="Clear memory and try again"):
                with st.spinner("Clearing resources..."):
                    cleanup_resources()
                st.session_state.stage = 'processing'
                add_message('bot', f"🔄 Cleared memory and retrying the {info['display_name']} report generation...")
                st.rerun()
        with col2:
            if st.button("🔄 Try Again", use_container_width=True):
                st.session_state.stage = 'processing'
                add_message('bot', f"Retrying the {info['display_name']} report generation...")
                st.rerun()
        with col3:
            if st.button("🏠 Start Over", use_container_width=True):
                st.session_state.messages = []
                st.session_state.stage = 'welcome'
                st.session_state.selected_project = None
                st.session_state.report_file = None
                if hasattr(st.session_state, 'error_message'):
                    delattr(st.session_state, 'error_message')
                st.rerun()

    # Footer
    st.markdown("---")
    st.markdown("""
    <div class="footer">
      <div style="font-size:1.2rem;">📊 Milestone Report Generator</div>
      <div>Automated report generation for project milestones</div>
      <div style="margin-top:1rem; font-size:0.9rem;">
        Supported Projects: Veridia • Eligo • EWS-LIG • WaveCityClub • Eden
      </div>
      <div style="margin-top:0.5rem; font-size:0.8rem; color: rgba(255,255,255,0.6);">
        💡 Tip: Use "Clear Memory" between reports for optimal performance
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()


