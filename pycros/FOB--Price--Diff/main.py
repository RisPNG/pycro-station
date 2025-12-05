import os
import csv
import re
import threading
import warnings
from datetime import datetime
from typing import List, Tuple, Any, Dict, Optional

# GUI Imports
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
    QTextEdit,
    QWidget,
    QSizePolicy
)
from qfluentwidgets import PrimaryPushButton, MessageBox

# Excel Imports
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Excel Automation for Formula Calculation
try:
    import xlwings as xw
    HAS_XLWINGS = True
except ImportError:
    HAS_XLWINGS = False

# --- Logic Implementation ---

# Extended Size Order for comparison logic
SIZE_ORDER = [
    "0", "2", "4", "6", "8", "10", "12", "14", "16", "18", "20", "22", "24", "26", "28", "30", "32", "34", "36", "38", "40", "42",
    "2XS", "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL",
    "2XSS", "XSS", "SS", "MS", "LS", "XLS", "2XLS", "3XLS", "4XLS", "5XLS",
    "2XSL", "XSL", "SL", "ML", "XLL", "2XLL", "3XLL", "4XLL", "5XLL",
    "2XST", "XST", "ST", "MT", "LT", "XLT", "2XLT", "3XLT", "4XLT", "5XLT",
    "2XSTT", "XSTT", "STT", "MTT", "LTT", "XLTT", "2XLTT", "3XLTT", "4XLTT", "5XLTT",
    "X", "0X", "1X", "2X", "3X", "4X", "5X",
    "XT", "0XT", "1XT", "2XT", "3XT", "4XT", "5XT",
    "XTT", "0XTT", "1XTT", "2XTT", "3XTT", "4XTT", "5XTT",
    "CUST2XS", "CUSTXS", "CUSTS", "CUSTM", "CUSTL", "CUSTXL", "CUST2XL", "CUST3XL", "CUST4XL", "CUST5XL",
    "CUST", "CUST0", "CUST1", "CUST2", "CUST3", "CUST4", "CUST5"
]

def normalize_header(header_text):
    """Normalize header text for comparison (remove newlines, extra spaces, uppercase)."""
    if not header_text:
        return ""
    return str(header_text).replace("\n", " ").replace("\r", "").strip().upper()

def get_col_index(headers, target_names):
    """Find index of a header that matches one of the target names."""
    if isinstance(target_names, str):
        target_names = [target_names]

    target_names = [normalize_header(t) for t in target_names]

    for idx, h in enumerate(headers):
        if normalize_header(h) in target_names:
            return idx
    return -1

def safe_float(value):
    """Safely convert value to float, handling currency strings."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    # String cleanup
    s_val = str(value).strip().replace(" ", "").replace("$", "").replace(",", "")
    if s_val == "-" or s_val == "":
        return 0.0
    try:
        return float(s_val)
    except ValueError:
        return 0.0

def normalize_date_str(date_val):
    """
    Convert various date formats (datetime obj, 'MM/DD/YYYY', 'YYYY-MM-DD')
    to standard 'MM/DD/YYYY' string for comparison.
    """
    if not date_val:
        return ""

    if isinstance(date_val, datetime):
        return date_val.strftime("%m/%d/%Y")

    s_val = str(date_val).strip()

    # Try parsing common formats
    formats = ["%m/%d/%Y", "%Y-%m-%d", "%d-%b-%y", "%m-%d-%Y"]
    for fmt in formats:
        try:
            dt = datetime.strptime(s_val, fmt)
            return dt.strftime("%m/%d/%Y")
        except ValueError:
            continue

    return s_val # Return as is if parsing fails (fallback)

def calculate_target_effective_date(buy_mth_str):
    """
    Converts OCCC 'BUY MTH' (e.g., '25-1E', '26-10M') to PPS Effective Date string.
    Logic:
    M=1,2,12 -> Dec 1st (Prev year for 1,2; Curr year for 12)
    M=3,4,5 -> Mar 1st
    M=6,7,8 -> Jun 1st
    M=9,10,11 -> Sep 1st
    """
    match = re.match(r"(\d{2})-(\d{1,2})", str(buy_mth_str).strip())
    if not match:
        return None

    yy = int(match.group(1))
    m = int(match.group(2))

    year = 2000 + yy
    target_month = 1
    target_year = year

    if m in [12, 1, 2]:
        target_month = 12
        if m in [1, 2]:
            target_year = year - 1
        else:
            target_year = year
    elif m in [3, 4, 5]:
        target_month = 3
    elif m in [6, 7, 8]:
        target_month = 6
    elif m in [9, 10, 11]:
        target_month = 9

    return f"{target_month:02d}/01/{target_year}"

def is_extended_size(ppm_size_str, occc_threshold_str):
    """Determine if a size is extended based on the OCCC threshold or TALL logic."""
    ppm_size = str(ppm_size_str).strip().upper()
    threshold = str(occc_threshold_str).strip().upper()

    if not ppm_size:
        return False

    if threshold in ["-", "", "NONE", "NA"]:
        return False

    try:
        idx_ppm = SIZE_ORDER.index(ppm_size)
    except ValueError:
        if "T" in ppm_size:
            return True
        return False

    try:
        idx_threshold = SIZE_ORDER.index(threshold)
    except ValueError:
        return False

    return idx_ppm >= idx_threshold

def refresh_excel_formulas(filepath, log_emit):
    """
    Uses xlwings to open, calculate, and save the file.
    This ensures openpyxl reads the calculated formula results instead of None/0.0.
    """
    if not HAS_XLWINGS:
        log_emit("Warning: xlwings not installed. Formulas might read as 0.0.")
        return False

    try:
        log_emit(f"Auto-calculating formulas for {os.path.basename(filepath)}... (This may take a moment)")
        app = xw.App(visible=False)
        app.display_alerts = False
        try:
            wb = app.books.open(filepath)
            wb.save()
            wb.close()
            log_emit("Formulas calculated and file saved.")
        except Exception as e:
            log_emit(f"Excel Automation Error: {e}")
        finally:
            try:
                app.quit()
            except:
                pass
    except Exception as e:
        log_emit(f"Could not launch Excel: {e}")

def refine_remarks(remarks_list):
    """Post-process remarks to consolidate messages."""
    if not remarks_list:
        return []

    s_pps_match_reg = "PPS OFOB match for regular sizes"
    s_pps_match_ext = "PPS OFOB match for extended sizes"
    s_pps_miss_reg = "PPS OFOB doesn't match for regular sizes"
    s_pps_miss_ext = "PPS OFOB doesn't match for extended sizes"
    s_final_miss_reg = "FINAL FOB (Regular sizes) doesn't match with PPM"
    s_final_miss_ext = "FINAL FOB (Extended sizes) doesn't match with PPM"

    r_set = set(remarks_list)
    targets = {s_pps_match_reg, s_pps_match_ext, s_pps_miss_reg, s_pps_miss_ext, s_final_miss_reg, s_final_miss_ext}

    final_list = []

    # 1. Keep unrelated remarks
    for r in remarks_list:
        if r not in targets:
            final_list.append(r)

    # 2. Regular Sizes Logic
    added_pps_issue_reg = False
    added_nike_issue_reg = False

    if s_pps_miss_reg in r_set and s_final_miss_reg in r_set:
        added_pps_issue_reg = True
    elif s_pps_match_reg in r_set and s_final_miss_reg in r_set:
        added_nike_issue_reg = True
    elif s_pps_miss_reg in r_set and s_final_miss_reg not in r_set:
        added_pps_issue_reg = True
    elif s_final_miss_reg in r_set:
        final_list.append(s_final_miss_reg)

    # 3. Extended Sizes Logic
    added_pps_issue_ext = False
    added_nike_issue_ext = False

    if s_pps_miss_ext in r_set and s_final_miss_ext in r_set:
        added_pps_issue_ext = True
    elif s_pps_match_ext in r_set and s_final_miss_ext in r_set:
        added_nike_issue_ext = True
    elif s_pps_miss_ext in r_set and s_final_miss_ext not in r_set:
        added_pps_issue_ext = True
    elif s_final_miss_ext in r_set:
        final_list.append(s_final_miss_ext)

    # 4. Consolidate
    if added_pps_issue_reg and added_pps_issue_ext:
        final_list.append("PPS OFOB issue for all sizes")
    else:
        if added_pps_issue_reg: final_list.append("PPS OFOB issue for regular sizes")
        if added_pps_issue_ext: final_list.append("PPS OFOB issue for extended sizes")

    if added_nike_issue_reg and added_nike_issue_ext:
        final_list.append("NIKE OFOB issue for all sizes")
    else:
        if added_nike_issue_reg: final_list.append("NIKE OFOB issue for regular sizes")
        if added_nike_issue_ext: final_list.append("NIKE OFOB issue for extended sizes")

    return final_list

def load_file_data(path, log_emit) -> Tuple[List[Any], List[List[Any]], Any]:
    """Load data from Excel or CSV."""
    ext = os.path.splitext(path)[1].lower()

    if ext == '.csv':
        try:
            with open(path, mode='r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows = list(reader)
                if not rows:
                    return [], [], None
                return rows, rows, None
        except Exception as e:
            log_emit(f"Error reading CSV {path}: {e}")
            raise e
    elif ext in ['.xlsx', '.xlsm']:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        return rows, rows, wb
    else:
        raise ValueError("Unsupported file format")

def process_logic(master_files, ppm_files, pps_files, log_emit, report_emit) -> Tuple[str, int, int]:
    success_count = 0
    fail_count = 0
    last_output = ""
    row_emit = report_emit if callable(report_emit) else log_emit

    # --- 1. Parse PPM Files ---
    ppm_lookup = {}
    log_emit("Parsing PPM Files...")
    for ppm_path in ppm_files:
        try:
            rows, raw_data, _ = load_file_data(ppm_path, log_emit)
            if not rows: continue

            header_row_idx = 0
            headers = [str(c) for c in rows[header_row_idx]]

            col_po = get_col_index(headers, ["Purchase Order Number", "TC PO (85/58)"])
            col_line = get_col_index(headers, ["PO Line Item Number", "PO LINE ITEM"])
            col_size = get_col_index(headers, ["Size Description"])

            # Costs
            col_ag = get_col_index(headers, ["Surcharge Min Mat Main Body"])
            col_ai = get_col_index(headers, ["Surcharge Min Material Trim"])
            col_ak = get_col_index(headers, ["Surcharge Min Productivity"])
            col_am = get_col_index(headers, ["Surcharge Misc"])
            col_ao = get_col_index(headers, ["Surcharge VAS"])
            col_aq = get_col_index(headers, ["Gross Price/FOB"])

            if col_po == -1 or col_line == -1:
                continue

            for r_idx in range(header_row_idx + 1, len(rows)):
                row = rows[r_idx]
                if not row: continue

                po_num = str(row[col_po]).strip()
                try: line_item = str(int(float(row[col_line])))
                except: line_item = str(row[col_line]).strip()
                key = (po_num, line_item)

                costs = {
                    'ag': safe_float(row[col_ag]) if col_ag != -1 else 0.0,
                    'ai': safe_float(row[col_ai]) if col_ai != -1 else 0.0,
                    'ak': safe_float(row[col_ak]) if col_ak != -1 else 0.0,
                    'am': safe_float(row[col_am]) if col_am != -1 else 0.0,
                    'ao': safe_float(row[col_ao]) if col_ao != -1 else 0.0,
                    'aq': safe_float(row[col_aq]) if col_aq != -1 else 0.0,
                    'size': str(row[col_size]).strip() if col_size != -1 else ""
                }

                if key not in ppm_lookup: ppm_lookup[key] = []
                ppm_lookup[key].append(costs)
        except Exception as e:
            log_emit(f"Error parsing PPM {os.path.basename(ppm_path)}: {e}")

    # --- 2. Parse PPS Files ---
    pps_lookup = {}
    log_emit("Parsing PPS Files...")
    for pps_path in pps_files:
        try:
            rows, raw_data, _ = load_file_data(pps_path, log_emit)
            if not rows: continue

            headers = [str(c) for c in rows[0]]

            col_style = get_col_index(headers, ["STYLE"])
            col_eff_date = get_col_index(headers, ["EFFECTIVE_DATE"])
            col_color = get_col_index(headers, ["COLOR"])
            col_size_data = get_col_index(headers, ["SIZE_DATA"])
            col_quote = get_col_index(headers, ["LOCAL_QUOTE_AMOUNT"])

            if col_style == -1 or col_eff_date == -1:
                continue

            for r_idx in range(1, len(rows)):
                row = rows[r_idx]
                if not row: continue

                style = str(row[col_style]).strip()
                eff_date = normalize_date_str(row[col_eff_date])

                if not style or not eff_date:
                    continue

                color = str(row[col_color]).strip() if col_color != -1 and row[col_color] is not None else ""
                size_data = str(row[col_size_data]).strip() if col_size_data != -1 and row[col_size_data] is not None else ""
                quote = safe_float(row[col_quote]) if col_quote != -1 else 0.0

                key = (style, eff_date)
                entry = {'color': color, 'size_data': size_data, 'quote': quote}

                if key not in pps_lookup: pps_lookup[key] = []
                pps_lookup[key].append(entry)

        except Exception as e:
            log_emit(f"Error parsing PPS {os.path.basename(pps_path)}: {e}")

    log_emit(f"PPS Data Loaded. Found {len(pps_lookup)} Style/Date keys.")

    # --- 3. Process OCCC Files ---
    for occc_path in master_files:
        try:
            # === xlwings Magic: Calculate Formulas ===
            if occc_path.lower().endswith(('.xlsx', '.xlsm')):
                refresh_excel_formulas(occc_path, log_emit)

            log_emit(f"Processing Master: {os.path.basename(occc_path)}")
            is_excel = occc_path.lower().endswith(('.xlsx', '.xlsm'))

            if is_excel:
                wb_read = load_workbook(occc_path, data_only=True)
                ws_read = wb_read.active
                rows_read = list(ws_read.values)
                wb_write = load_workbook(occc_path, data_only=False)
                ws_write = wb_write.active
            else:
                rows_read, _, _ = load_file_data(occc_path, log_emit)
                output_csv_data = [list(r) for r in rows_read]
                ws_write = None

            header_idx = 2
            if len(rows_read) <= header_idx:
                log_emit(f"Master file too short.")
                fail_count += 1
                continue

            headers = [str(x) for x in rows_read[header_idx]]

            # Map Columns
            idx_nk_po = get_col_index(headers, ["NK SAP PO (45/35)", "NK SAP PO"])
            idx_line = get_col_index(headers, ["PO LINE ITEM"])
            idx_ave_fob = get_col_index(headers, ["AVE FOB ON DPOM"])

            idx_sc_min_prod = get_col_index(headers, ["S/C Min Production (ZPMX)"])
            idx_sc_min_mat = get_col_index(headers, ["S/C Min Material (ZMMX)"])
            idx_sc_min_mat_comment = get_col_index(headers, ["S/C Min Material (ZMMX) Comment"])
            idx_sc_misc = get_col_index(headers, ["S/C Misc (ZMSX)"])
            idx_sc_misc_comment = get_col_index(headers, ["S/C Misc (ZMSX) Comment"])
            idx_sc_vas = get_col_index(headers, ["S/C VAS Manual (ZVAX)"])

            idx_style = get_col_index(headers, ["STYLE"])
            idx_buy_mth = get_col_index(headers, ["BUY MTH"])
            idx_cw = get_col_index(headers, ["CW"])

            idx_ofob_reg = get_col_index(headers, ["OFOB (Regular sizes)"])
            idx_ofob_ext = get_col_index(headers, ["OFOB (Extended sizes)"])
            idx_final_reg = get_col_index(headers, ["FINAL FOB (Regular sizes)"])
            idx_final_ext = get_col_index(headers, ["FINAL FOB (Extended sizes)", "FINAL FOB (Extended sizes) (2)"])
            idx_ext_sizes_def = get_col_index(headers, ["Extended Sizes"])

            idx_remarks = get_col_index(headers, ["PRICE DIFF REMARKS"])
            idx_dpom_fob = get_col_index(headers, ["DPOM - Incorrect FOB"])

            # Init Header for Remarks
            if idx_remarks == -1:
                ref_col = idx_ave_fob if idx_ave_fob != -1 else len(headers) - 1
                insert_pos = ref_col + 1
                if is_excel:
                    ws_write.insert_cols(insert_pos + 1)
                    ws_write.cell(row=header_idx+1, column=insert_pos+1).value = "PRICE DIFF REMARKS"
                    idx_remarks = insert_pos
                else:
                    output_csv_data[header_idx].append("PRICE DIFF REMARKS")
                    idx_remarks = len(headers)

            # Init Header for DPOM - Incorrect FOB (re-check in case it shifted due to insert)
            if is_excel:
                # Re-fetch headers if we inserted a column to ensure we don't mess up indexing
                headers = [str(cell.value) for cell in ws_write[header_idx+1]]
                idx_dpom_fob = get_col_index(headers, ["DPOM - Incorrect FOB"])

            if idx_dpom_fob == -1:
                # Place it after Remarks, or at the end
                insert_pos = idx_remarks + 1
                if is_excel:
                    ws_write.insert_cols(insert_pos + 1)
                    ws_write.cell(row=header_idx+1, column=insert_pos+1).value = "DPOM - Incorrect FOB"
                    idx_dpom_fob = insert_pos
                else:
                    output_csv_data[header_idx].append("DPOM - Incorrect FOB")
                    idx_dpom_fob = len(headers) - 1

            for r_i in range(header_idx + 1, len(rows_read)):
                row_vals = rows_read[r_i]
                if not row_vals: continue

                remarks = []
                dpom_errors = [] # Store "Size Price" mismatches

                # --- PPM Comparison ---
                po_val = str(row_vals[idx_nk_po]).strip() if idx_nk_po != -1 else ""
                line_val = ""
                if idx_line != -1:
                    try: line_val = str(int(float(row_vals[idx_line])))
                    except: line_val = str(row_vals[idx_line]).strip()

                if po_val and line_val:
                    ppm_entries = ppm_lookup.get((po_val, line_val))
                    if ppm_entries:
                        # Avg calc for surcharges
                        sum_ag = sum_ai = sum_am = sum_ao = 0.0
                        count = len(ppm_entries)
                        for entry in ppm_entries:
                            sum_ag += entry['ag']
                            sum_ai += entry['ai']
                            sum_am += entry['am']
                            sum_ao += entry['ao']

                        ave_ppm_ag = sum_ag / count if count else 0.0
                        ave_ppm_ai = sum_ai / count if count else 0.0
                        ave_ppm_am = sum_am / count if count else 0.0
                        ave_ppm_ao = sum_ao / count if count else 0.0

                        # Surcharge Checks - THRESHOLD 0.01
                        if abs(safe_float(row_vals[idx_sc_min_prod]) - ave_ppm_ag) > 0.01:
                            remarks.append("S/C MIN PRODUCTION (ZPMX) doesn't match")

                        # Min Mat
                        occc_zmmx = safe_float(row_vals[idx_sc_min_mat])
                        zmmx_cmt = str(row_vals[idx_sc_min_mat_comment]).strip().upper() if idx_sc_min_mat_comment != -1 and row_vals[idx_sc_min_mat_comment] else ""
                        if zmmx_cmt != "DN" and abs(occc_zmmx - ave_ppm_ai) > 0.01:
                            remarks.append("S/C Min Material (ZMMX) doesn't match")

                        # Misc
                        occc_zmsx = safe_float(row_vals[idx_sc_misc])
                        zmsx_cmt = str(row_vals[idx_sc_misc_comment]).strip().upper() if idx_sc_misc_comment != -1 and row_vals[idx_sc_misc_comment] else ""
                        if zmsx_cmt != "DN" and abs(occc_zmsx - ave_ppm_am) > 0.01:
                            remarks.append("S/C Misc (ZMSX) doesn't match")

                        if abs(safe_float(row_vals[idx_sc_vas]) - ave_ppm_ao) > 0.01:
                            remarks.append("S/C VAS Manual (ZVAX) doesn't match")

                        # Final FOB Checks
                        ext_threshold = str(row_vals[idx_ext_sizes_def]).strip() if idx_ext_sizes_def != -1 else ""
                        occc_final_reg = safe_float(row_vals[idx_final_reg]) if idx_final_reg != -1 else 0.0
                        occc_final_ext = safe_float(row_vals[idx_final_ext]) if idx_final_ext != -1 else 0.0

                        fob_mismatch_found = False # Flag to avoid spamming "Remarks" but keep collecting DPOM errors

                        for entry in ppm_entries:
                            ppm_total = round(entry['ag'] + entry['ai'] + entry['ak'] +
                                              entry['am'] + entry['ao'] + entry['aq'], 2)
                            is_ext = is_extended_size(entry['size'], ext_threshold)
                            target_fob = round(occc_final_ext if is_ext else occc_final_reg, 2)

                            # THRESHOLD 0.01
                            if ppm_total > 0 and abs(target_fob - ppm_total) > 0.01:
                                lbl = "Extended" if is_ext else "Regular"

                                # Add to DPOM Error List: "Size Price"
                                dpom_errors.append(f"{entry['size']} {ppm_total:.2f}")

                                # Add to Remarks (Only once per row to avoid clutter)
                                if not fob_mismatch_found:
                                    row_emit(f"Mismatch Row {r_i+1} PO {po_val}: {lbl} Size - OCCC {target_fob} vs PPM {ppm_total}")
                                    remarks.append(f"FINAL FOB ({lbl} sizes) doesn't match with PPM")
                                    fob_mismatch_found = True

                                # Do NOT break here. Continue checking other sizes for DPOM column.

                # --- PPS Comparison ---
                style_val = str(row_vals[idx_style]).strip() if idx_style != -1 else ""
                buy_mth_val = str(row_vals[idx_buy_mth]).strip() if idx_buy_mth != -1 else ""
                cw_val = str(row_vals[idx_cw]).strip() if idx_cw != -1 else ""

                if style_val and buy_mth_val:
                    target_date = calculate_target_effective_date(buy_mth_val)
                    if target_date:
                        pps_candidates = pps_lookup.get((style_val, target_date))
                        if pps_candidates:
                            matched_rows = [r for r in pps_candidates if r['color'] == cw_val]
                            if not matched_rows: matched_rows = [r for r in pps_candidates if not r['color']]

                            if not matched_rows:
                                remarks.append("No matching PPS (Color)")
                            else:
                                # Regular - THRESHOLD 0.01
                                reg_match = next((r for r in matched_rows if not r['size_data']), None)
                                occc_ofob_reg = safe_float(row_vals[idx_ofob_reg])
                                if reg_match:
                                    if abs(reg_match['quote'] - occc_ofob_reg) <= 0.01:
                                        remarks.append("PPS OFOB match for regular sizes")
                                    else:
                                        remarks.append("PPS OFOB doesn't match for regular sizes")
                                elif occc_ofob_reg > 0:
                                    remarks.append("PPS OFOB missing regular size entry")

                                # Extended - THRESHOLD 0.01
                                ext_threshold = str(row_vals[idx_ext_sizes_def]).strip() if idx_ext_sizes_def != -1 else ""
                                occc_ofob_ext = safe_float(row_vals[idx_ofob_ext])
                                if ext_threshold not in ["-", "", "NONE", "NA"] or occc_ofob_ext > 0:
                                    ext_match = next((r for r in matched_rows if is_extended_size(r['size_data'], ext_threshold)), None)
                                    if ext_match:
                                        if abs(ext_match['quote'] - occc_ofob_ext) <= 0.01:
                                            remarks.append("PPS OFOB match for extended sizes")
                                        else:
                                            remarks.append("PPS OFOB doesn't match for extended sizes")
                                    elif occc_ofob_ext > 0:
                                        remarks.append("PPS OFOB missing extended size entry")
                        else:
                            remarks.append("No matching PPS found")
                    else:
                        remarks.append("Invalid BUY MTH format")

                # --- Post-Processing ---
                if remarks:
                    remarks = refine_remarks(remarks)

                # --- Write Output ---

                # 1. PRICE DIFF REMARKS
                final_remark = "; ".join(remarks) if remarks else ""

                # 2. DPOM - Incorrect FOB
                final_dpom_val = " / ".join(dpom_errors) if dpom_errors else ""

                if is_excel:
                    # Write Remarks
                    target_col_idx = idx_remarks + 1
                    if ws_write.cell(row=header_idx+1, column=target_col_idx).value != "PRICE DIFF REMARKS":
                        ws_write.cell(row=header_idx+1, column=target_col_idx).value = "PRICE DIFF REMARKS"
                    ws_write.cell(row=r_i+1, column=target_col_idx).value = final_remark

                    # Write DPOM
                    target_dpom_idx = idx_dpom_fob + 1
                    if ws_write.cell(row=header_idx+1, column=target_dpom_idx).value != "DPOM - Incorrect FOB":
                        ws_write.cell(row=header_idx+1, column=target_dpom_idx).value = "DPOM - Incorrect FOB"
                    ws_write.cell(row=r_i+1, column=target_dpom_idx).value = final_dpom_val

                else:
                    # CSV Handling

                    # Ensure list is long enough for Remarks
                    while len(output_csv_data[r_i]) <= idx_remarks:
                         output_csv_data[r_i].append("")
                    output_csv_data[r_i][idx_remarks] = final_remark

                    # Ensure list is long enough for DPOM
                    while len(output_csv_data[r_i]) <= idx_dpom_fob:
                         output_csv_data[r_i].append("")
                    output_csv_data[r_i][idx_dpom_fob] = final_dpom_val

            if is_excel:
                wb_write.save(occc_path)
            else:
                with open(occc_path, mode='w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerows(output_csv_data)

            success_count += 1
            last_output = occc_path
            log_emit(f"Updated {os.path.basename(occc_path)}")

        except Exception as e:
            log_emit(f"Failed to process {os.path.basename(occc_path)}: {e}")
            fail_count += 1

    return last_output, success_count, fail_count

# --- UI Class ---

class MainWidget(QWidget):
    log_message = Signal(str)
    report_message = Signal(str)
    processing_done = Signal(int, int, str)

    def __init__(self):
        super().__init__()
        self.setObjectName("mmu_widget")
        self._build_ui()
        self._connect_signals()

    def _build_ui(self):
        self.desc_label = QLabel("", self)
        self.desc_label.setWordWrap(True)
        self.desc_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.desc_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.desc_label.setStyleSheet(
            "color: #dcdcdc; background: transparent; padding: 6px; "
            "border: 1px solid #3a3a3a; border-radius: 6px;"
        )
        self.set_long_description("")

        self.select_master_btn = PrimaryPushButton("Select Master (OCCC)", self)
        self.select_ppm_btn = PrimaryPushButton("Select PPM Reports", self)
        self.select_pps_btn = PrimaryPushButton("Select PPS Reports", self)
        self.run_btn = PrimaryPushButton("Run Validation", self)

        self.master_files_label = QLabel("Master file(s)", self)
        self.ppm_files_label = QLabel("PPM report file(s)", self)
        self.pps_files_label = QLabel("PPS report file(s)", self)
        self.logs_label = QLabel("Process logs", self)
        self.reports_label = QLabel("Report output", self)

        for lbl in [self.master_files_label, self.ppm_files_label, self.pps_files_label, self.logs_label, self.reports_label]:
            lbl.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        shared_style = (
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        self.master_files_box = QTextEdit(self)
        self.master_files_box.setReadOnly(True)
        self.master_files_box.setStyleSheet(shared_style)

        self.ppm_files_box = QTextEdit(self)
        self.ppm_files_box.setReadOnly(True)
        self.ppm_files_box.setStyleSheet(shared_style)

        self.pps_files_box = QTextEdit(self)
        self.pps_files_box.setReadOnly(True)
        self.pps_files_box.setStyleSheet(shared_style)

        self.reports_box = QTextEdit(self)
        self.reports_box.setReadOnly(True)
        self.reports_box.setStyleSheet(shared_style)

        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet(shared_style)

        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.desc_label)

        row1 = QHBoxLayout()
        row1.addWidget(self.select_master_btn)
        row1.addWidget(self.select_ppm_btn)
        row1.addWidget(self.select_pps_btn)
        main_layout.addLayout(row1)

        row_btn = QHBoxLayout()
        row_btn.addStretch()
        row_btn.addWidget(self.run_btn)
        row_btn.addStretch()
        main_layout.addLayout(row_btn)

        row_labels = QHBoxLayout()
        row_labels.addWidget(self.master_files_label)
        row_labels.addWidget(self.ppm_files_label)
        row_labels.addWidget(self.pps_files_label)
        main_layout.addLayout(row_labels)

        row_boxes = QHBoxLayout()
        row_boxes.addWidget(self.master_files_box)
        row_boxes.addWidget(self.ppm_files_box)
        row_boxes.addWidget(self.pps_files_box)
        main_layout.addLayout(row_boxes, 2)

        row_log_lbl = QHBoxLayout()
        row_log_lbl.addWidget(self.reports_label)
        row_log_lbl.addWidget(self.logs_label)
        main_layout.addLayout(row_log_lbl)

        row_logs = QHBoxLayout()
        row_logs.addWidget(self.reports_box)
        row_logs.addWidget(self.log_box)
        main_layout.addLayout(row_logs, 2)

    def set_long_description(self, text: str):
        clean = (text or "").strip()
        if clean:
            self.desc_label.setText(clean)
            self.desc_label.show()
        else:
            self.desc_label.hide()

    def _connect_signals(self):
        self.select_master_btn.clicked.connect(lambda: self.select_files(self.master_files_box))
        self.select_ppm_btn.clicked.connect(lambda: self.select_files(self.ppm_files_box))
        self.select_pps_btn.clicked.connect(lambda: self.select_files(self.pps_files_box))
        self.run_btn.clicked.connect(self.run_process)
        self.log_message.connect(self.append_log)
        self.report_message.connect(self.append_report)
        self.processing_done.connect(self.on_processing_done)

    def select_files(self, text_box):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", "Excel/CSV Files (*.xlsx *.xlsm *.csv)")
        if files:
            text_box.setPlainText("\n".join(files))
        else:
            text_box.clear()

    def get_files_from_box(self, text_box):
        text = text_box.toPlainText().strip()
        return [line.strip() for line in text.split("\n") if line.strip()]

    def run_process(self):
        master_files = self.get_files_from_box(self.master_files_box)
        ppm_files = self.get_files_from_box(self.ppm_files_box)
        pps_files = self.get_files_from_box(self.pps_files_box)

        if not master_files:
            MessageBox("Warning", "Please select OCCC Master file.", self).exec()
            return

        if not ppm_files and not pps_files:
             MessageBox("Warning", "Please select at least one report file (PPM or PPS).", self).exec()
             return

        self.log_box.clear()
        self.reports_box.clear()
        self.log_message.emit("Process Started...")

        self.run_btn.setEnabled(False)
        self.select_master_btn.setEnabled(False)
        self.select_ppm_btn.setEnabled(False)
        self.select_pps_btn.setEnabled(False)

        def worker():
            try:
                last_file, ok, fail = process_logic(master_files, ppm_files, pps_files, self.log_message.emit, self.report_message.emit)
                self.processing_done.emit(ok, fail, last_file)
            except Exception as e:
                self.log_message.emit(f"CRITICAL ERROR: {e}")
                self.processing_done.emit(0, 0, "")

        threading.Thread(target=worker, daemon=True).start()

    def append_log(self, text):
        self.log_box.append(text)

    def append_report(self, text):
        self.reports_box.append(text)

    def on_processing_done(self, ok, fail, last_file):
        self.log_message.emit(f"Done. Success: {ok}, Failed: {fail}")
        if last_file:
            self.log_message.emit(f"Last processed: {last_file}")
        self.run_btn.setEnabled(True)
        self.select_master_btn.setEnabled(True)
        self.select_ppm_btn.setEnabled(True)
        self.select_pps_btn.setEnabled(True)

        title = "Process complete" if fail == 0 else "Process finished with issues"
        lines = [f"Success: {ok}", f"Failed: {fail}"]
        if last_file:
            lines.append(f"Last processed: {last_file}")
        msg = MessageBox(title, "\n".join(lines), self)
        msg.yesButton.setText("OK")
        msg.cancelButton.hide()
        msg.exec()

def get_widget():
    return MainWidget()

# - If value contains "PPS OFOB match for extended sizes" but not accompanied with "FINAL FOB (Extended sizes) doesn't match", replace the "PPS OFOB match for extended sizes" with ""
# - If value contains "PPS OFOB match for regular sizes" but not accompanied with "FINAL FOB (Regular sizes) doesn't match", replace the "PPS OFOB match for regular sizes" with ""
# - If value contains "PPS OFOB match for extended sizes" but not accompanied with "FINAL FOB (Extended sizes) doesn't match", replace the "PPS OFOB match for extended sizes" with ""
# - If value contains "PPS OFOB doesn't match for regular sizes" and accompanied with "FINAL FOB (Regular sizes) doesn't match", replace them both with a single "PPS OFOB issue for regular sizes"
# - If value contains "PPS OFOB doesn't match for extended sizes" and accompanied with "FINAL FOB (Extended sizes) doesn't match", replace them both with a single "PPS OFOB issue for extended sizes"
# - If value contains "PPS OFOB match for regular sizes" and accompanied with "FINAL FOB (Regular sizes) doesn't match", replace them both with a single "NIKE OFOB issue for regular sizes"
# - If value contains "PPS OFOB match for extended sizes" and accompanied with "FINAL FOB (Extended sizes) doesn't match", replace them both with a single "NIKE OFOB issue for extended sizes"
# - If value contains "PPS OFOB doesn't match for regular sizes" and not accompanied with "FINAL FOB (Regular sizes) doesn't match", replace it with a single "PPS OFOB issue for regular sizes"
# - If value contains "PPS OFOB doesn't match for extended sizes" and not accompanied with "FINAL FOB (Extended sizes) doesn't match", replace it with a single "PPS OFOB issue for extended sizes"

# once all that is process, do another pass for the following:
# - If value contains "PPS OFOB issue for regular sizes" and accompanied with "PPS OFOB issue for extended sizes", replace them both with a single "PPS OFOB issue for all sizes"
# - If value contains "NIKE OFOB issue for regular sizes" and accompanied with "NIKE OFOB issue for extended sizes", replace them both with a single "NIKE OFOB issue for all sizes"
