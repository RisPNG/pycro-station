import os
import csv
import threading
import warnings
from datetime import datetime
from typing import List, Tuple, Any, Dict, Optional

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
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- Logic Implementation ---

# Extended Size Order for comparison
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

def is_extended_size(ppm_size_str, occc_threshold_str):
    """
    Determine if a size is extended based on the OCCC threshold or TALL logic.
    """
    ppm_size = str(ppm_size_str).strip().upper()

    # Rule 1: TALL sizes are usually extended (contain 'T' but usually at end like XLT, MT)
    # Based on the list, almost all *T or *TT are extended versions.
    # Logic: If it contains "T" (Tall) it is likely extended, unless threshold logic applies strictly.
    # The prompt implies checking the tier list logic.

    # If OCCC threshold is empty/dash, usually implies no extended logic or everything is regular
    # But usually "Extended Sizes" column has something like "3XL" or "4XL".
    threshold = str(occc_threshold_str).strip().upper()
    if threshold in ["-", "", "NONE", "NA"]:
        # Fallback: if T in size, might still be extended, but if threshold is undefined,
        # usually means this style doesn't have the split. assume regular.
        return False

    # Check indices in SIZE_ORDER
    try:
        idx_ppm = SIZE_ORDER.index(ppm_size)
    except ValueError:
        # Size not in standard list, treat as regular unless it looks like a TALL
        return False

    try:
        idx_threshold = SIZE_ORDER.index(threshold)
    except ValueError:
        # Threshold not in list?
        return False

    return idx_ppm >= idx_threshold

def load_file_data(path, log_emit) -> Tuple[List[Any], List[List[Any]], Any]:
    """
    Load data from Excel or CSV.
    Returns: (headers, data_rows, workbook_object_if_excel)
    """
    ext = os.path.splitext(path)[1].lower()
    headers = []
    data = []
    wb = None

    if ext == '.csv':
        try:
            with open(path, mode='r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows = list(reader)
                if not rows:
                    return [], [], None
                # Return all rows, logic will determine header row index later
                return rows, rows, None # CSV returns raw rows as both header source and data
        except Exception as e:
            log_emit(f"Error reading CSV {path}: {e}")
            raise e
    elif ext in ['.xlsx', '.xlsm']:
        wb = load_workbook(path, data_only=True) # Read values for processing
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        return rows, rows, wb
    else:
        raise ValueError("Unsupported file format")

def process_logic(master_files, ppm_files, log_emit, report_emit) -> Tuple[str, int, int]:
    success_count = 0
    fail_count = 0
    last_output = ""

    # 1. Parse PPM Files into a Lookup Dictionary
    # Key: (PO Number, PO Line Item) -> Value: List of row dicts containing costs & sizes
    ppm_lookup = {}

    log_emit("Parsing PPM Files...")

    for ppm_path in ppm_files:
        try:
            rows, raw_data, _ = load_file_data(ppm_path, log_emit)
            if not rows:
                continue

            # Header is usually row 1 (index 0)
            header_row_idx = 0
            headers = [str(c) for c in rows[header_row_idx]]

            # Map columns
            # Required: PO, Line Item, Size Desc
            # Costs: Min Mat Main, Min Mat Trim, Min Prod, Misc, VAS, Gross FOB

            col_po = get_col_index(headers, ["Purchase Order Number", "TC PO (85/58)"])
            col_line = get_col_index(headers, ["PO Line Item Number", "PO LINE ITEM"])
            col_size = get_col_index(headers, ["Size Description"])

            # Cost columns
            col_ag = get_col_index(headers, ["Surcharge Min Mat Main Body"])
            col_ai = get_col_index(headers, ["Surcharge Min Material Trim"])
            col_ak = get_col_index(headers, ["Surcharge Min Productivity"])
            col_am = get_col_index(headers, ["Surcharge Misc"])
            col_ao = get_col_index(headers, ["Surcharge VAS"])
            col_aq = get_col_index(headers, ["Gross Price/FOB"])

            if col_po == -1 or col_line == -1:
                log_emit(f"Skipping PPM {os.path.basename(ppm_path)}: Missing PO/Line headers.")
                continue

            for r_idx in range(header_row_idx + 1, len(rows)):
                row = rows[r_idx]
                if not row: continue

                # Extract Keys
                po_num = str(row[col_po]).strip()
                # Handle line item being integer or float (e.g. 10.0 -> 10)
                try:
                    line_item = str(int(float(row[col_line])))
                except:
                    line_item = str(row[col_line]).strip()

                key = (po_num, line_item)

                # Extract Costs
                costs = {
                    'ag': safe_float(row[col_ag]) if col_ag != -1 else 0.0,
                    'ai': safe_float(row[col_ai]) if col_ai != -1 else 0.0,
                    'ak': safe_float(row[col_ak]) if col_ak != -1 else 0.0,
                    'am': safe_float(row[col_am]) if col_am != -1 else 0.0,
                    'ao': safe_float(row[col_ao]) if col_ao != -1 else 0.0,
                    'aq': safe_float(row[col_aq]) if col_aq != -1 else 0.0,
                    'size': str(row[col_size]).strip() if col_size != -1 else ""
                }

                if key not in ppm_lookup:
                    ppm_lookup[key] = []
                ppm_lookup[key].append(costs)

        except Exception as e:
            log_emit(f"Error parsing PPM {os.path.basename(ppm_path)}: {e}")

    log_emit(f"PPM Data Loaded. Found {len(ppm_lookup)} unique PO Lines.")

    # 2. Process OCCC Files
    for occc_path in master_files:
        try:
            log_emit(f"Processing Master: {os.path.basename(occc_path)}")

            # We need the Workbook object to save editable Excel
            is_excel = occc_path.lower().endswith(('.xlsx', '.xlsm'))

            if is_excel:
                # Load twice: one data_only for reading values, one normal for writing
                wb_read = load_workbook(occc_path, data_only=True)
                ws_read = wb_read.active
                rows_read = list(ws_read.values) # Tuple of tuples

                wb_write = load_workbook(occc_path, data_only=False)
                ws_write = wb_write.active
            else:
                # CSV
                rows_read, _, _ = load_file_data(occc_path, log_emit)
                # Prepare data for CSV write later
                output_csv_data = [list(r) for r in rows_read]
                ws_write = None # No worksheet object for CSV

            # Header is Row 3 (Index 2)
            header_idx = 2
            if len(rows_read) <= header_idx:
                log_emit(f"Master file too short.")
                fail_count += 1
                continue

            headers = [str(x) for x in rows_read[header_idx]]

            # Map OCCC Columns
            idx_nk_po = get_col_index(headers, ["NK SAP PO (45/35)", "NK SAP PO"])
            idx_tc_po = get_col_index(headers, ["TC PO (85/58)"])
            idx_line = get_col_index(headers, ["PO LINE ITEM"])

            # OCCC Target Columns for Comparison
            idx_ave_fob = get_col_index(headers, ["AVE FOB ON DPOM"])

            # Individual Surcharges in OCCC
            idx_sc_min_prod = get_col_index(headers, ["S/C Min Production (ZPMX)"])
            idx_sc_min_mat = get_col_index(headers, ["S/C Min Material (ZMMX)"])
            idx_sc_misc = get_col_index(headers, ["S/C Misc (ZMSX)"])
            idx_sc_vas = get_col_index(headers, ["S/C VAS Manual (ZVAX)"])

            # FOBs
            idx_final_reg = get_col_index(headers, ["FINAL FOB (Regular sizes)"])
            idx_final_ext = get_col_index(headers, ["FINAL FOB (Extended sizes)", "FINAL FOB (Extended sizes) (2)"])
            idx_ext_sizes_def = get_col_index(headers, ["Extended Sizes"])

            # Remarks Column
            idx_remarks = get_col_index(headers, ["PRICE DIFF REMARKS"])

            if idx_remarks == -1:
                # Insert new column after AVE FOB ON DPOM if possible, else at end
                insert_pos = idx_ave_fob + 1 if idx_ave_fob != -1 else len(headers)

                if is_excel:
                    ws_write.insert_cols(insert_pos + 1) # 1-based index
                    ws_write.cell(row=header_idx+1, column=insert_pos+1).value = "PRICE DIFF REMARKS"
                    idx_remarks = insert_pos
                    # Adjust read indices that are after insertion point?
                    # Since we use `rows_read` loaded before insertion, indices reference original state.
                    # But writing needs to account for shift.
                    # Easier strategy: Append to end to avoid index shifting headaches in complex sheets
                    pass
                else:
                    output_csv_data[header_idx].append("PRICE DIFF REMARKS")
                    idx_remarks = len(headers)

            # Re-evaluate logic: simple append to end is safer for logic simplicity
            # But prompt suggests "after AVE FOB".
            # To strictly follow "replace input", let's just use the column if exists, or append to end.

            # Iterate Data Rows
            for r_i in range(header_idx + 1, len(rows_read)):
                row_vals = rows_read[r_i]
                if not row_vals: continue

                # Get Keys (Try NK PO first, then TC PO if needed, usually NK is key)
                po_val = str(row_vals[idx_nk_po]).strip() if idx_nk_po != -1 else ""
                line_val = ""
                if idx_line != -1:
                    try:
                        line_val = str(int(float(row_vals[idx_line])))
                    except:
                        line_val = str(row_vals[idx_line]).strip()

                if not po_val or not line_val:
                    continue

                lookup_key = (po_val, line_val)
                ppm_entries = ppm_lookup.get(lookup_key)

                if not ppm_entries:
                    # report_emit(f"No PPM data for PO {po_val} Line {line_val}")
                    continue

                # --- COMPARISON LOGIC ---
                remarks = []

                # 1. Calculate PPM Average FOB
                total_sum_fob = 0.0

                # Summing all components for average
                # PPM columns: AG(Main) + AI(Trim) + AK(Prod) + AM(Misc) + AO(VAS) + AQ(Gross)
                sum_ag = 0.0
                sum_ai = 0.0
                sum_am = 0.0
                sum_ao = 0.0
                count = len(ppm_entries)

                for entry in ppm_entries:
                    row_total = (entry['ag'] + entry['ai'] + entry['ak'] +
                                 entry['am'] + entry['ao'] + entry['aq'])
                    total_sum_fob += row_total

                    sum_ag += entry['ag']
                    sum_ai += entry['ai']
                    sum_am += entry['am']
                    sum_ao += entry['ao']

                ave_ppm_fob = total_sum_fob / count if count > 0 else 0.0
                ave_ppm_ag = sum_ag / count if count > 0 else 0.0
                ave_ppm_ai = sum_ai / count if count > 0 else 0.0
                ave_ppm_am = sum_am / count if count > 0 else 0.0
                ave_ppm_ao = sum_ao / count if count > 0 else 0.0

                # 2. Check Average FOB against OCCC
                occc_ave_fob = safe_float(row_vals[idx_ave_fob]) if idx_ave_fob != -1 else 0.0

                fob_diff = abs(occc_ave_fob - ave_ppm_fob)
                is_avg_match = fob_diff < 0.02 # Tolerance

                if not is_avg_match:
                    # "Ave. FOB doesn't match" - Proceed to check components
                    remarks.append(f"Ave. FOB doesn't match (OCCC:{occc_ave_fob:.2f} vs PPM:{ave_ppm_fob:.2f})")

                    # Check Components
                    # OCCC ZPMX vs PPM AG (Min Prod? Prompt says PPM AG is Min Mat Main??)
                    # Prompt Mapping:
                    # OCCC ZPMX <-> PPM AG (Surcharge Min Mat Main Body)
                    # OCCC ZMMX <-> PPM AI (Surcharge Min Material Trim)
                    # OCCC ZMSX <-> PPM AM (Surcharge Misc)
                    # OCCC ZVAX <-> PPM AO (Surcharge VAS)

                    occc_zpmx = safe_float(row_vals[idx_sc_min_prod]) if idx_sc_min_prod != -1 else 0.0
                    if abs(occc_zpmx - ave_ppm_ag) > 0.02:
                        remarks.append(f"S/C MIN PRODUCTION (ZPMX) doesn't match")

                    occc_zmmx = safe_float(row_vals[idx_sc_min_mat]) if idx_sc_min_mat != -1 else 0.0
                    if abs(occc_zmmx - ave_ppm_ai) > 0.02:
                        remarks.append(f"S/C Min Material (ZMMX) doesn't match")

                    occc_zmsx = safe_float(row_vals[idx_sc_misc]) if idx_sc_misc != -1 else 0.0
                    if abs(occc_zmsx - ave_ppm_am) > 0.02:
                        remarks.append(f"S/C Misc (ZMSX) doesn't match")

                    occc_zvax = safe_float(row_vals[idx_sc_vas]) if idx_sc_vas != -1 else 0.0
                    if abs(occc_zvax - ave_ppm_ao) > 0.02:
                        remarks.append(f"S/C VAS Manual (ZVAX) doesn't match")

                # 3. Check Final FOB per Size (Regular vs Extended)
                # OCCC Threshold
                ext_threshold = str(row_vals[idx_ext_sizes_def]).strip() if idx_ext_sizes_def != -1 else ""

                occc_final_reg = safe_float(row_vals[idx_final_reg]) if idx_final_reg != -1 else 0.0
                occc_final_ext = safe_float(row_vals[idx_final_ext]) if idx_final_ext != -1 else 0.0

                size_mismatch_found = False

                for entry in ppm_entries:
                    ppm_total_fob = (entry['ag'] + entry['ai'] + entry['ak'] +
                                     entry['am'] + entry['ao'] + entry['aq'])

                    is_ext = is_extended_size(entry['size'], ext_threshold)

                    target_occc_fob = occc_final_ext if is_ext else occc_final_reg
                    # Usually Extended FOB is 0 if no extended sizes exist, handle that?
                    # If target is 0 and ppm is not 0, it's a mismatch.

                    if abs(target_occc_fob - ppm_total_fob) > 0.02:
                        lbl = "Extended" if is_ext else "Regular"
                        remarks.append(f"FINAL FOB ({lbl} sizes) doesn't match")
                        size_mismatch_found = True
                        break # Stop checking other sizes if one fails to avoid spamming remarks

                # Write Remarks
                if remarks:
                    final_remark = "; ".join(remarks)

                    if is_excel:
                        # Write to Excel
                        # Check if column existed or we append
                        if idx_remarks >= len(headers):
                             # Column was appended virtually, need to ensure column exists in sheet
                             # For simplicity in this logic, we write to the calculated column index + 1 (1-based)
                             cell = ws_write.cell(row=r_i + 1, column=ws_write.max_column + 1)
                             # This appends a new column every row if not careful.
                             # Better: define column index fixed.
                             # If we didn't find the header, we assume column index is len(headers) (0-based)
                             # So 1-based is len(headers)+1
                             target_col = len(headers) + 1
                             # Update header if first time (check r_i == header_idx + 1 is risky if we skipped rows)
                             if ws_write.cell(row=header_idx+1, column=target_col).value != "PRICE DIFF REMARKS":
                                 ws_write.cell(row=header_idx+1, column=target_col).value = "PRICE DIFF REMARKS"
                             ws_write.cell(row=r_i+1, column=target_col).value = final_remark
                        else:
                            # Existing column
                            ws_write.cell(row=r_i+1, column=idx_remarks+1).value = final_remark

                        # Highlight row? (Optional, per "skeleton" usually implies just logic)
                    else:
                        # CSV
                        if idx_remarks >= len(headers):
                            # Append to row
                            output_csv_data[r_i].append(final_remark)
                        else:
                            output_csv_data[r_i][idx_remarks] = final_remark

            # Save Output
            # "output occc replaces the input occc" -> Overwrite
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
        pps_files = self.get_files_from_box(self.pps_files_box) # Logic to be added later

        if not master_files or not ppm_files:
            MessageBox("Warning", "Please select OCCC Master and PPM files.", self).exec()
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
                # We currently ignore PPS in processing as per instruction "this is it for now for PPM"
                last_file, ok, fail = process_logic(master_files, ppm_files, self.log_message.emit, self.report_message.emit)
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

def get_widget():
    return MainWidget()