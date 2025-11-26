import os
import threading
from datetime import datetime, date, timedelta
from typing import List, Tuple, Any, Optional, Dict, NamedTuple

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
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

try:
    import xlrd  # for legacy excel
except Exception:
    xlrd = None

# --- CONSTANTS & COLORS (Converted from VBA Long to ARGB Hex) ---
# VBA Color Longs are typically BGR. OpenPyXL expects RGB (or ARGB).
# cDMagenta = 8388736  -> 0x800080 (Purple)
# cCyan = 16776960     -> 0xFFFF00 (BGR) -> 0x00FFFF (RGB) -> FF00FFFF
# cPink = 12830955     -> 0xC3C4EB (BGR) -> 0xEBC4C3 (RGB) -> FFEBC4C3
# cGreen = 8252325     -> 0x7DE3A5 (BGR) -> 0xA5E37D (RGB) -> FFA5E37D
# cYellow = 12123902   -> 0xB8FEE6 (BGR) -> 0xE6FEB8 (RGB) -> FFE6FEB8
# cDGreen = 5287936    -> 0x50B000 (BGR) -> 0x00B050 (RGB) -> FF00B050
# cDarkBlue = 16737792 -> 0xFF6000 (BGR) -> 0x0060FF (RGB) -> FF0060FF
# cOrange = 3368703    -> 0x3366FF (BGR) -> 0xFF6633 (RGB) -> FFFF6633

COLOR_CYAN = PatternFill(start_color="FF00FFFF", end_color="FF00FFFF", fill_type="solid")
COLOR_PINK = PatternFill(start_color="FFEBC4C3", end_color="FFEBC4C3", fill_type="solid")
COLOR_GREEN = PatternFill(start_color="FFA5E37D", end_color="FFA5E37D", fill_type="solid")
COLOR_YELLOW = PatternFill(start_color="FFE6FEB8", end_color="FFE6FEB8", fill_type="solid")
COLOR_D_GREEN = PatternFill(start_color="FF00B050", end_color="FF00B050", fill_type="solid")
COLOR_BLACK = PatternFill(start_color="FF000000", end_color="FF000000", fill_type="solid")
COLOR_WHITE = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")

class BlindBuyItem(NamedTuple):
    sc: str
    bbj: str

class MainWidget(QWidget):
    log_message = Signal(str)
    processing_done = Signal(int, int, str)

    def __init__(self):
        super().__init__()
        self.setObjectName("dpom_sorter_widget")
        self._build_ui()
        self._connect_signals()

    # UI Construction
    def _build_ui(self):
        self.desc_label = QLabel("", self)
        self.desc_label.setWordWrap(True)
        self.desc_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.desc_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.desc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.desc_label.setStyleSheet(
            "color: #dcdcdc; background: transparent; padding: 6px; "
            "border: 1px solid #3a3a3a; border-radius: 6px;"
        )
        self.set_long_description("")

        self.select_btn2 = PrimaryPushButton("Select Size Excel", self)
        self.select_btn = PrimaryPushButton("Select Raw DPOM Excel Files", self)
        self.run_btn = PrimaryPushButton("Run", self)

        self.files_label = QLabel("Selected files", self)
        self.files_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")
        self.logs_label = QLabel("Process logs", self)
        self.logs_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.files_box = QTextEdit(self)
        self.files_box.setReadOnly(True)
        self.files_box.setPlaceholderText("Selected files will appear here")
        self.files_box.setStyleSheet(
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        self.files_box2 = QTextEdit(self)
        self.files_box2.setReadOnly(True)
        self.files_box2.setMaximumHeight(50)
        self.files_box2.setPlaceholderText("Selected size file will appear here")
        self.files_box2.setStyleSheet(
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Live process log will appear here")
        self.log_box.setStyleSheet(
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        main_layout.addWidget(self.desc_label, 1)

        row1_layout = QHBoxLayout()
        row1_layout.addStretch(1)
        row1_layout.addWidget(self.select_btn2, 1)
        row1_layout.addStretch(1)
        main_layout.addLayout(row1_layout, 0)

        row2_layout = QHBoxLayout()
        row2_layout.addWidget(self.files_box2, 1)
        main_layout.addLayout(row2_layout, 0)

        row3_layout = QHBoxLayout()
        row3_layout.addWidget(self.select_btn, 1)
        row3_layout.addWidget(self.run_btn, 1)
        main_layout.addLayout(row3_layout, 0)

        row4_layout = QHBoxLayout()
        row4_layout.addWidget(self.files_label, 1)
        row4_layout.addWidget(self.logs_label, 1)
        main_layout.addLayout(row4_layout, 0)

        row5_layout = QHBoxLayout()
        row5_layout.addWidget(self.files_box, 1)
        row5_layout.addWidget(self.log_box, 1)
        main_layout.addLayout(row5_layout, 4)

    def set_long_description(self, text: str):
        clean = (text or "").strip()
        if clean:
            self.desc_label.setText(clean)
            self.desc_label.show()
        else:
            self.desc_label.clear()
            self.desc_label.hide()

    def _connect_signals(self):
        self.select_btn2.clicked.connect(self.select_size_file)
        self.select_btn.clicked.connect(self.select_files)
        self.run_btn.clicked.connect(self.run_process)
        self.log_message.connect(self.append_log)
        self.processing_done.connect(self.on_processing_done)

    def select_size_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Size / Blind Buy workbook", filter="Excel Files (*.xlsx *.xlsm *.xls)")
        if file_path:
            self.files_box2.setPlainText(file_path)
        else:
            self.files_box2.clear()

    def _selected_size_file(self) -> str:
        return self.files_box2.toPlainText().strip()

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select", filter="Excel Files (*.xlsx *.xlsm *.xls)")
        if files:
            self.files_box.setPlainText("\n".join(files))
        else:
            self.files_box.clear()

    def _selected_files(self) -> List[str]:
        text = self.files_box.toPlainText().strip()
        if not text:
            return []
        return [line for line in text.split("\n") if line.strip()]

    def run_process(self):
        size_file = self._selected_size_file()
        if not size_file:
            MessageBox("Warning", "Size / Blind Buy workbook not selected.", self).exec()
            return

        files = self._selected_files()
        if not files:
            MessageBox("Warning", "Nothing to process.", self).exec()
            return

        self.log_box.clear()
        self.log_message.emit(f"Process starts")
        self.run_btn.setEnabled(False)
        self.select_btn.setEnabled(False)

        def worker():
            ok, fail, out_path = 0, 0, ""
            try:
                out_path, ok, fail = process_files(size_file, files, self.log_message.emit)
            except Exception as e:
                import traceback
                self.log_message.emit(f"CRITICAL ERROR: {e}\n{traceback.format_exc()}")
            self.processing_done.emit(ok, fail, out_path)

        threading.Thread(target=worker, daemon=True).start()

    def append_log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def on_processing_done(self, ok: int, fail: int, out_path: str):
        self.log_message.emit(f"Completed: {ok} success, {fail} failed.")
        self.run_btn.setEnabled(True)
        self.select_btn.setEnabled(True)

def get_widget():
    return MainWidget()

# --- HELPER FUNCTIONS ---

def normalize_size(size_val: Any) -> str:
    """Helper function to standardizes a size string for reliable comparison."""
    if size_val is None:
        return ""
    temp_size = str(size_val).strip().upper()
    temp_size = temp_size.replace("-", "")
    temp_size = temp_size.replace("XXXXXL", "6XL")
    temp_size = temp_size.replace("XXXXL", "5XL")
    temp_size = temp_size.replace("XXXXL", "4XL")
    temp_size = temp_size.replace("XXXL", "3XL")
    temp_size = temp_size.replace("XXL", "2XL")
    temp_size = temp_size.replace("XXXXXS", "6XS")
    temp_size = temp_size.replace("XXXXS", "5XS")
    temp_size = temp_size.replace("XXXS", "4XS")
    temp_size = temp_size.replace("XXXS", "3XS")
    temp_size = temp_size.replace("XXS", "2XS")
    return temp_size

def get_col_let(col_idx: int) -> str:
    return get_column_letter(col_idx)

def process_files(size_file_path: str, dpom_files: List[str], log_emit) -> Tuple[str, int, int]:
    processor = DPOMSorter(size_file_path, log_emit)
    success = 0
    fail = 0
    last_path = ""

    for f in dpom_files:
        try:
            processor.process_dpom(f)
            success += 1
            last_path = f
        except Exception as e:
            import traceback
            log_emit(f"Failed to process {os.path.basename(f)}: {e}")
            print(traceback.format_exc())
            fail += 1

    return last_path, success, fail

class DPOMSorter:
    def __init__(self, size_file_path, log_emit):
        self.log = log_emit
        self.size_file_path = size_file_path
        self.blind_buy_list: List[BlindBuyItem] = []
        self.master_sizes: List[str] = []

        # Default Columns (1-based, matching VBA)
        self.c_Vendor = 1
        self.c_SeasonCode = 2
        self.c_SeasonYear = 3
        self.c_Style = 4
        self.c_PO = 7
        self.c_TradingCoPO = 8
        self.c_POLine = 9
        self.c_OGAC = 11
        self.c_DocTypeCode = 15
        self.c_Transportation = 17
        self.c_ShipToCusNo = 18
        self.c_ShipToCusName = 19
        self.c_Country = 20
        self.c_InventorySegmentCode = 21
        self.c_SubCategoryDesc = 23
        self.c_SizeDesc = 24
        self.c_SizeQty = 25
        self.c_TotalSizeQty = 26
        self.c_FOB = 27

        self.max_col = 0
        self.size_col_start = 0

        self._load_reference_data()

    def _load_reference_data(self):
        self.log("Loading reference data (Size / Blind Buy)...")
        wb_ref = load_workbook(self.size_file_path, data_only=True)

        # Load Blind Buy
        if "Blind Buy" in wb_ref.sheetnames:
            ws_bb = wb_ref["Blind Buy"]
            for row in ws_bb.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    sc_val = str(row[0]).strip().upper()
                    bbj_val = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                    self.blind_buy_list.append(BlindBuyItem(sc_val, bbj_val))

        # Load Size Master
        if "Size" in wb_ref.sheetnames:
            ws_size = wb_ref["Size"]
            for row in ws_size.iter_rows(min_row=2, max_col=1, values_only=True):
                if row[0]:
                    self.master_sizes.append(str(row[0]))

        wb_ref.close()

    def _blind_buy_contains(self, style_code: str) -> bool:
        style_code = style_code.strip().upper()
        for item in self.blind_buy_list:
            if item.sc == style_code:
                # Note: VBA sets a global 'val' here used in AddNewColumn.
                # We will handle that retrieval when needed.
                return True
        return False

    def _get_blind_buy_val(self, style_code: str) -> str:
        style_code = style_code.strip().upper()
        for item in self.blind_buy_list:
            if item.sc == style_code:
                return item.bbj
        return ""

    def process_dpom(self, file_path):
        fname = os.path.basename(file_path)
        self.log(f"Opening {fname}...")

        wb = load_workbook(file_path) # Not data_only, we might want to preserve styles initially
        ws_dpom = wb.worksheets[0]

        # Validation
        if ws_dpom.cell(1, 27).value == "PROCESSED": # Check Column AA (27)
            self.log(f"{fname} already processed. Skipping.")
            return

        # Basic Header Check (Simplified from VBA)
        if str(ws_dpom.cell(1, self.c_Vendor).value).strip() != "Vendor Code":
            self.log(f"{fname} format incorrect. Skipping.")
            return

        # Create copy sheet
        ws_result = wb.copy_worksheet(ws_dpom)
        ws_result.title = ws_dpom.title + "-copy"
        wb.move_sheet(ws_result, offset=1)

        # Create Temp sheet
        ws_temp = wb.create_sheet("Temp")

        # Reset Column Pointers for this file
        self._reset_col_pointers()

        # Execute Steps
        self.log(f"Sorting Record {fname}...")
        self.sort_record(ws_result)

        self.log(f"Rearranging Size Columns {fname}...")
        self.rearrange_size_col(ws_result, ws_temp)

        # Cleanup temp sheet
        wb.remove(ws_temp)

        self.log(f"Running First Pass {fname}...")
        self.first_pass(ws_result)

        self.log(f"Running Second Pass {fname}...")
        self.second_pass(ws_result)

        self.log(f"Adding New Columns {fname}...")
        self.add_new_column(ws_result)

        self.log(f"Removing Columns {fname}...")
        self.remove_column(ws_result)

        self.log(f"Reverting to OBS Format {fname}...")
        self.revert_to_obs(ws_result)

        # Mark as Processed in Original Sheet
        ws_dpom.cell(1, 27).value = "PROCESSED"
        ws_dpom.cell(1, 27).font = Font(color="FFFFFFFF") # White text
        ws_result.cell(1, 1).value = "PROCESSED"

        # AutoFit (Approximate)
        for col in ws_result.columns:
             # openpyxl autofit is manual, usually skipped or approximated
             pass

        # MODULE 2 Logic
        self.log(f"Separating Data Groups (Module 2) {fname}...")
        self.separate_data_groups(ws_result)

        self.log(f"Saving {fname}...")
        wb.save(file_path)
        self.log(f"Finished {fname}.")

    def _reset_col_pointers(self):
        self.c_Vendor = 1
        self.c_SeasonCode = 2
        self.c_SeasonYear = 3
        self.c_Style = 4
        self.c_PO = 7
        self.c_TradingCoPO = 8
        self.c_POLine = 9
        self.c_OGAC = 11
        self.c_DocTypeCode = 15
        self.c_Transportation = 17
        self.c_ShipToCusNo = 18
        self.c_ShipToCusName = 19
        self.c_Country = 20
        self.c_InventorySegmentCode = 21
        self.c_SubCategoryDesc = 23
        self.c_SizeDesc = 24
        self.c_SizeQty = 25
        self.c_TotalSizeQty = 26
        self.c_FOB = 27

    def sort_record(self, ws: Worksheet):
        max_col = ws.max_column
        max_row = ws.max_row
        start_row = 2
        if max_row < start_row: return

        # Read data to list of dicts for sorting
        data_rows = []
        for r in range(start_row, max_row + 1):
            row_vals = [cell.value for cell in ws[r]]
            # Pad if row is short
            if len(row_vals) < max_col:
                row_vals.extend([None] * (max_col - len(row_vals)))

            # Extract sort keys
            style = str(row_vals[self.c_Style-1] or "").strip()
            season = str(row_vals[self.c_SeasonCode-1] or "").strip().upper()
            season_year = str(row_vals[self.c_SeasonYear-1] or "") # keep as is
            ogac = str(row_vals[self.c_OGAC-1] or "").strip()
            po = str(row_vals[self.c_PO-1] or "")
            po_line = str(row_vals[self.c_POLine-1] or "")
            country = str(row_vals[self.c_Country-1] or "")
            ship_to = str(row_vals[self.c_ShipToCusNo-1] or "")

            # Derived Keys
            style_head = style[:6] if len(style) >= 6 else style
            style_cw = style[-3:] if len(style) >= 3 else style

            # Season Rank
            s_rank = 5
            if season == "SP": s_rank = 1
            elif season == "SU": s_rank = 2
            elif season == "FA": s_rank = 3
            elif season == "HO": s_rank = 4

            # OGAC Sort Key
            og_key = ogac
            if "/" in ogac:
                parts = ogac.split("/")
                if len(parts) >= 3:
                    # mm/dd/yyyy -> yyyymmdd
                    yy = parts[2][-4:]
                    mm = parts[0].zfill(2)
                    dd = parts[1].zfill(2)
                    og_key = yy + mm + dd
            elif len(ogac) >= 4:
                 og_key = ogac[-4:] # Fallback

            # We store the full row object (cells) to move them, or just values.
            # Since we overwrite in place, values are safer.
            data_rows.append({
                "vals": row_vals,
                "k_style_head": style_head,
                "k_year": season_year,
                "k_s_rank": s_rank,
                "k_ogac": og_key,
                "k_po": po,
                "k_cw": style_cw,
                "k_poline": po_line,
                "k_country": country,
                "k_ship": ship_to
            })

        # Python Sort (Stable) - Sort in reverse order of precedence
        # Precedence: Ship, Country, POLine, CW, PO, OGAC, Rank, Year, StyleHead
        data_rows.sort(key=lambda x: (
            x["k_style_head"],
            x["k_year"],
            x["k_s_rank"],
            x["k_ogac"],
            x["k_po"],
            x["k_cw"],
            x["k_poline"],
            x["k_country"],
            x["k_ship"]
        ))

        # Write back
        for i, item in enumerate(data_rows):
            r = start_row + i
            for c, val in enumerate(item["vals"]):
                ws.cell(row=r, column=c+1).value = val

    def rearrange_size_col(self, ws: Worksheet, ws_temp: Worksheet):
        self.max_col = ws.max_column
        self.size_col_start = self.max_col + 1
        max_row = ws.max_row

        # Get unique sizes from colSizeDesc
        sizes = set()
        for r in range(2, max_row + 1):
            val = ws.cell(row=r, column=self.c_SizeDesc).value
            if val:
                sizes.add(str(val).strip())

        # Sort sizes based on Master List
        # 1. Put all unique sizes into a list
        unique_sizes = list(sizes)

        # 2. Sort Logic matching VBA (Normalize comparison)
        sorted_sizes = []

        # Match against Master List
        for master_size in self.master_sizes:
            norm_master = normalize_size(master_size)
            # Find matching in unique_sizes
            found_idx = -1
            for i, us in enumerate(unique_sizes):
                if normalize_size(us) == norm_master:
                    sorted_sizes.append(us)
                    found_idx = i
                    break
            if found_idx != -1:
                unique_sizes.pop(found_idx)

        # Add remaining
        for us in unique_sizes:
            if us: sorted_sizes.append(us)

        # Write Headers
        for i, s in enumerate(sorted_sizes):
            ws.cell(row=1, column=self.size_col_start + i).value = s

    def first_pass(self, ws: Worksheet):
        start = 2
        start_ptr = 2

        # Update max_col (headers added)
        self.max_col = ws.max_column

        ws.cell(row=1, column=self.max_col + 1).value = "Overall Result"

        # Build Header Map (Normalized Size -> Col Index)
        size_map = {}
        for c in range(self.size_col_start, self.max_col + 1):
            h_val = ws.cell(row=1, column=c).value
            if h_val:
                size_map[normalize_size(h_val)] = c

        row_idx = 2
        while row_idx <= ws.max_row:
            # Check if row matches next row
            matches_next = False
            if row_idx < ws.max_row:
                curr = self._get_key_vals(ws, row_idx)
                nxt = self._get_key_vals(ws, row_idx + 1)
                if curr == nxt:
                    matches_next = True

            if not matches_next:
                end_ptr = row_idx

                # Insert 2 new rows at end_ptr + 1
                ws.insert_rows(end_ptr + 1, amount=2)

                # Labels
                ws.cell(row=end_ptr, column=self.max_col + 2).value = "Total Item Qty"
                ws.cell(row=end_ptr + 1, column=self.max_col + 2).value = "Trading Co Net Unit Price"
                ws.cell(row=end_ptr + 2, column=self.max_col + 2).value = "Net Unit Price"

                # Aggregate Logic
                # Loop through the block [start_ptr, end_ptr]
                row_sum = 0
                for r in range(start_ptr, end_ptr + 1):
                    s_desc = ws.cell(row=r, column=self.c_SizeDesc).value
                    s_qty = ws.cell(row=r, column=self.c_SizeQty).value
                    fob = ws.cell(row=r, column=self.c_FOB).value

                    if s_qty is None: s_qty = 0
                    try: s_qty = float(s_qty)
                    except: s_qty = 0

                    norm_key = normalize_size(s_desc)
                    target_col = size_map.get(norm_key)

                    if target_col:
                        ws.cell(row=end_ptr, column=target_col).value = s_qty
                        ws.cell(row=end_ptr + 1, column=target_col).value = fob
                        ws.cell(row=end_ptr + 2, column=target_col).value = fob
                        row_sum += s_qty

                ws.cell(row=end_ptr, column=self.max_col + 1).value = row_sum

                # Next block
                start_ptr = end_ptr + 3
                row_idx = start_ptr - 1 # Adjusted because of inserts

            row_idx += 1

        # Clean up
        # Remove original rows (where helper col "Total Item Qty" is empty)
        # Iterate backwards to safely delete
        for r in range(ws.max_row, 1, -1):
            val = ws.cell(row=r, column=self.max_col + 2).value
            if val is None or str(val) == "":
                ws.delete_rows(r)

        # Calculate Totals at bottom
        last_row = ws.max_row
        ws.cell(row=last_row + 1, column=1).value = "Overall Result"
        ws.cell(row=last_row + 1, column=self.max_col + 2).value = "Total Item Qty"

        # Move helper column to ShipToCusNo position
        # In VBA: wsRESULT.Columns(maxCol + 2).Cut ... Insert
        # Py: Read column, insert at 18, delete old.
        # Helper Col Index = max_col + 2
        helper_col_idx = self.max_col + 2
        # Data is in c_ShipToCusNo = 18.
        # OpenPyXL insert_cols inserts BEFORE.
        ws.insert_cols(self.c_ShipToCusNo)
        # Copy data
        for r in range(1, ws.max_row + 2):
            val = ws.cell(row=r, column=helper_col_idx + 1).value # +1 because insert shifted it
            ws.cell(row=r, column=self.c_ShipToCusNo).value = val
        # Delete old column
        ws.delete_cols(helper_col_idx + 1)

        # Adjust indices
        # SizeDesc(24), SizeQty(25), TotalSizeQty(26), FOB(27) deleted in VBA
        # Indices shifted because of insert at 18.
        # Old 24 is now 25.

        # VBA:
        # colSizeDesc = 24+1; colSizeQty=25+1; colTotalSizeQty=26+1; colFOB=27+1
        # Delete Columns(25 to 28)

        del_start = self.c_SizeDesc + 1 # 25
        ws.delete_cols(del_start, 4) # Delete 4 columns

        # Update Class Pointers
        self.c_Ptr = 17 # Q (Old Transportation 17? No colPtr is ShipToCusNo logic)
        # VBA: colPtr = colShipToCusNo (before increment) = 17 (Q is 17) -> Actually colShipToCusNo is 18 in VBA init.
        # Let's trace carefully.
        # Init: ShipToCusNo=18.
        # Insert at 18. New ShipToCusNo data is at 18. Old ShipToCusNo shifted to 19.
        # VBA: colPtr = colShipToCusNo (17? No 18).
        # VBA code says: colPtr = colShipToCusNo '17 'Q. Wait, 18 is R. 17 is Q.
        # In VBA: colShipToCusNo = 18.
        # wsRESULT.Columns(maxCol + 2).Cut -> wsRESULT.Columns(colShipToCusNo).Insert
        # So "Total Item Qty" / "Overall Result" labels are now in Col 18.
        # colPtr tracks the Labels column.

        self.c_Ptr = 18
        self.c_ShipToCusNo = 19
        self.c_ShipToCusName = 20
        self.c_Country = 21
        self.c_InventorySegmentCode = 22
        self.c_SubCategoryDesc = 24

        # Current Layout:
        # 1..17 Data
        # 18: Labels (Total Item Qty)
        # 19..24 Data
        # 25..End Sizes

        self.size_col_start = self.size_col_start - 3 # -4 deleted + 1 inserted

    def _get_key_vals(self, ws, r):
        # Keys: SC, SY, Style, PO, POLine, ShipNo, ShipName, Qty(26)
        return (
            str(ws.cell(r, self.c_SeasonCode).value or "").strip(),
            str(ws.cell(r, self.c_SeasonYear).value or "").strip(),
            str(ws.cell(r, self.c_Style).value or "").strip(),
            str(ws.cell(r, self.c_PO).value or "").strip(),
            str(ws.cell(r, self.c_POLine).value or "").strip(),
            str(ws.cell(r, self.c_ShipToCusNo).value or "").strip(),
            str(ws.cell(r, self.c_ShipToCusName).value or "").strip(),
            str(ws.cell(r, self.c_TotalSizeQty).value or "").strip()
        )

    def second_pass(self, ws: Worksheet):
        self.max_col = ws.max_column

        # 1. STYLE GROUPING (Green)
        start = 2
        start_ptr = 2

        # Helper to get clean style
        def get_clean_style(r):
            v = str(ws.cell(r, self.c_Style).value or "").strip().upper()
            if "-" in v:
                v = v.split("-")[0]
            return v

        prev_st = get_clean_style(start)

        while start <= ws.max_row:
            curr_st = get_clean_style(start)

            # Check for change or end of data
            # Logic: If Style changes, process previous block [start_ptr, start-1]
            if curr_st != prev_st:
                end_ptr = start - 1

                # Find Point (Last "Total Item Qty" row in block)
                point = end_ptr # fallback
                for r in range(start_ptr, end_ptr + 1):
                    if ws.cell(r, self.c_Ptr).value == "Total Item Qty":
                        point = r

                # Insert 4 rows at start
                ws.insert_rows(start, 4)

                # Set Labels
                ws.cell(start, self.c_Ptr).value = "Total Style Qty"

                # Copy MetaData (Left side) from Point row
                # Copy row point+1 to start+1 (Trading Co Price)
                # Copy row point+2 to start+2 (Net Unit Price)
                # Apply Green formatting

                # Because we inserted 4 rows at 'start', 'point' is still valid relative to data above 'start'.

                # Copy logic (Cols 1 to max_col)
                self._copy_row_range(ws, point, start, self.c_Ptr) # Label row
                self._copy_row_range(ws, point + 1, start + 1, self.max_col) # Price row 1
                self._copy_row_range(ws, point + 2, start + 2, self.max_col) # Price row 2

                # Format Green
                self._format_range(ws, start, start, 1, self.max_col, COLOR_GREEN, True)
                self._format_range(ws, start+1, start+1, 1, self.max_col, COLOR_GREEN, True)
                self._format_range(ws, start+2, start+2, 1, self.max_col, COLOR_GREEN, True)
                self._format_range(ws, start+3, start+3, 1, self.max_col, COLOR_D_GREEN, False)
                ws.row_dimensions[start+3].height = 5

                # Calculate Sums
                total_sum1 = 0
                total_sum2 = 0

                for c in range(self.size_col_start, self.max_col): # max_col is inclusive in logic?
                    # Loop start_ptr to end_ptr (original block)
                    # Note: indices shifted if we inserted before? No, we inserted at 'start', which is AFTER end_ptr.
                    qty_sum = 0
                    m1 = 0
                    m2 = 0

                    for r in range(start_ptr, end_ptr + 1):
                        if ws.cell(r, self.c_Ptr).value == "Total Item Qty":
                            v = ws.cell(r, c).value
                            if v and isinstance(v, (int, float)): qty_sum += v

                            v1 = ws.cell(r+1, c).value
                            if v1: m1 = v1
                            v2 = ws.cell(r+2, c).value
                            if v2: m2 = v2

                    ws.cell(start, c).value = qty_sum if qty_sum != 0 else None
                    ws.cell(start+1, c).value = m1 if m1 != 0 else None
                    ws.cell(start+2, c).value = m2 if m2 != 0 else None

                    total_sum1 += (qty_sum * (float(m1) if m1 else 0))
                    total_sum2 += (qty_sum * (float(m2) if m2 else 0))

                ws.cell(start+1, self.max_col+1).value = total_sum1
                ws.cell(start+1, self.max_col+1).fill = COLOR_GREEN
                ws.cell(start+1, self.max_col+1).font = Font(bold=True)

                ws.cell(start+2, self.max_col+1).value = total_sum2
                ws.cell(start+2, self.max_col+1).fill = COLOR_GREEN
                ws.cell(start+2, self.max_col+1).font = Font(bold=True)

                # Total Qty (Col Max)
                block_qty = 0
                for r in range(start_ptr, end_ptr+1):
                    if ws.cell(r, self.c_Ptr).value == "Total Item Qty":
                        v = ws.cell(r, self.max_col).value
                        if v: block_qty += v
                ws.cell(start, self.max_col).value = block_qty

                # Advance
                start = start + 4
                start_ptr = start
                prev_st = curr_st

            start += 1

        # 2. OGAC GROUPING (Pink)
        start = 2
        start_ptr = 2
        while start <= ws.max_row:
            row_type = ws.cell(start, self.c_Ptr).value
            if row_type == "Total Style Qty":
                inner_start = start_ptr
                inner_end = start

                curr_ogac = str(ws.cell(inner_start, self.c_OGAC).value or "").strip().upper()
                prev_ogac = curr_ogac

                # Loop through the Style Block
                curr_r = inner_start
                while curr_r < inner_end:
                    curr_ogac = str(ws.cell(curr_r, self.c_OGAC).value or "").strip().upper()

                    # Logic: If OGAC changes, insert Total OGAC Qty row
                    # Or if we hit "Total Style Qty" (end of block)
                    is_change = (curr_ogac != prev_ogac)

                    if is_change:
                        # Insert Pink rows at curr_r
                        self._insert_summary_block(ws, curr_r, inner_start, "Total OGAC Qty", COLOR_PINK, True)

                        # Adjust indices
                        inner_end += 2
                        start += 2
                        inner_start = curr_r + 2 # Start of new group
                        curr_r += 2
                        prev_ogac = str(ws.cell(curr_r, self.c_OGAC).value or "").strip().upper()

                    curr_r += 1

                # Final group in block
                self._insert_summary_block(ws, start, inner_start, "Total OGAC Qty", COLOR_PINK, True)
                start += 2
                start_ptr = start + 1 # Next block starts after separator

            start += 1

        # 3. PO / Country / AFS / ShipTo (Yellow) + Blind Buy
        start = 2
        start_ptr = 2

        while start <= ws.max_row:
            row_type = ws.cell(start, self.c_Ptr).value

            if row_type == "Total OGAC Qty":
                inner_start = start_ptr
                block_end = start # The OGAC Total row index

                curr_r = inner_start
                # Init tracking vars
                prev_c = str(ws.cell(curr_r, self.c_Country).value or "").strip().upper()
                prev_p = str(ws.cell(curr_r, self.c_PO).value or "").strip().upper()
                prev_a = str(ws.cell(curr_r, self.c_InventorySegmentCode).value or "").strip().upper()
                prev_s = str(ws.cell(curr_r, self.c_ShipToCusNo).value or "").strip().upper()

                while curr_r < block_end:
                    curr_c = str(ws.cell(curr_r, self.c_Country).value or "").strip().upper()
                    curr_p = str(ws.cell(curr_r, self.c_PO).value or "").strip().upper()
                    curr_a = str(ws.cell(curr_r, self.c_InventorySegmentCode).value or "").strip().upper()
                    curr_s = str(ws.cell(curr_r, self.c_ShipToCusNo).value or "").strip().upper()
                    style = str(ws.cell(curr_r, self.c_Style).value or "").strip().upper()

                    is_bb = self._blind_buy_contains(style)

                    changed = (curr_c != prev_c or curr_p != prev_p or curr_a != prev_a or curr_s != prev_s)

                    if changed or is_bb:
                        if is_bb:
                            # Blind Buy Logic: Insert 3 yellow rows, 1 separator
                            # Must check if previous was same blind buy style to avoid duplicates?
                            # VBA checks: Not(Previous is BB) ...
                            prev_style = str(ws.cell(curr_r-1, self.c_Style).value or "").strip().upper()
                            is_prev_bb = self._blind_buy_contains(prev_style)

                            # Complex conditional from VBA:
                            # If BB(current) AND (Not BB(prev) AND not at start AND prev row has label)
                            # Actually, we just need to break group here.

                            # We treat BB rows as singleton groups often, or break whenever BB state changes.
                            # For simplicity/robustness: treat BB as a trigger to summarize the *previous* group.
                            pass

                        # Summarize Previous Group [inner_start to curr_r - 1]
                        # Insert Yellow Block at curr_r
                        added = self._insert_po_block(ws, curr_r, inner_start, COLOR_YELLOW, is_bb)

                        block_end += added
                        start += added
                        curr_r += added
                        inner_start = curr_r

                        # Reset prev pointers to current
                        prev_c = str(ws.cell(curr_r, self.c_Country).value or "").strip().upper()
                        prev_p = str(ws.cell(curr_r, self.c_PO).value or "").strip().upper()
                        prev_a = str(ws.cell(curr_r, self.c_InventorySegmentCode).value or "").strip().upper()
                        prev_s = str(ws.cell(curr_r, self.c_ShipToCusNo).value or "").strip().upper()

                    curr_r += 1

                # Summarize last group
                added = self._insert_po_block(ws, block_end, inner_start, COLOR_YELLOW, False) # Last one not necessarily BB special case?
                # Actually if the last group was BB, logic handles it inside the formatting of that block
                start += added
                start_ptr = start + 1

            start += 1
            # Skip gaps
            if start <= ws.max_row and ws.cell(start, self.c_Ptr).value is None:
                start += 1
                start_ptr = start

            # Skip Style Qty blocks (size 4)
            if start <= ws.max_row and ws.cell(start, self.c_Ptr).value == "Total Style Qty":
                start += 4
                start_ptr = start

        # 4. COLORWAY GROUPING (Cyan)
        # Scan inside Data Groups.
        start = 2
        start_ptr = 2
        cw_count = 0

        def get_cw(r):
            st = str(ws.cell(r, self.c_Style).value or "").strip().upper()
            if "-" in st: return st.split("-")[-1]
            return st

        prev_cw = get_cw(start)

        while start <= ws.max_row:
            row_lbl = ws.cell(start, self.c_Ptr).value

            if row_lbl == "Total Item Qty":
                curr_cw = get_cw(start)

                if curr_cw == prev_cw:
                    cw_count += 1
                else:
                    if cw_count >= 2:
                        # Insert Cyan row at start
                        ws.insert_rows(start, 1)
                        ws.cell(start, self.c_Ptr).value = "Total Colorway Qty"
                        self._copy_row_range(ws, start-1, start, self.c_Ptr) # copy label cols

                        # Sums
                        for c in range(self.size_col_start, self.max_col): # max_col inclusive? range is exclusive
                             s = 0
                             for r in range(start_ptr, start):
                                 if ws.cell(r, self.c_Ptr).value == "Total Item Qty":
                                     v = ws.cell(r, c).value
                                     if v and isinstance(v, (int, float)): s += v
                             ws.cell(start, c).value = s if s != 0 else None

                        # Total col
                        t_s = 0
                        for r in range(start_ptr, start):
                             if ws.cell(r, self.c_Ptr).value == "Total Item Qty":
                                 v = ws.cell(r, self.max_col).value
                                 if v: t_s += v
                        ws.cell(start, self.max_col).value = t_s

                        self._format_range(ws, start, start, 1, self.max_col, COLOR_CYAN, True)
                        start += 1 # skip inserted

                    start_ptr = start
                    cw_count = 1
                    prev_cw = curr_cw

            elif row_lbl == "Total PO Qty":
                 # Reset
                 start += 4 # Skip yellow block
                 if start <= ws.max_row:
                     prev_cw = get_cw(start)
                     start_ptr = start
                     cw_count = 0
                 continue

            start += 1

    def _insert_summary_block(self, ws, insert_at, data_start, label, fill_color, is_ogac_style):
        # Insert 2 rows (Total + Sep)
        ws.insert_rows(insert_at, 2)

        point = insert_at # The new row

        # Copy Label
        self._copy_row_range(ws, insert_at - 1, point, self.c_Ptr) # simplistic copy from row above
        ws.cell(point, self.c_Ptr).value = label

        # Sums
        for c in range(self.size_col_start, self.max_col):
            s = 0
            for r in range(data_start, insert_at): # Range is pre-insert indices? No, insert pushes down.
                # Data is above insert_at
                if ws.cell(r, self.c_Ptr).value == "Total Item Qty":
                    v = ws.cell(r, c).value
                    if v and isinstance(v, (int, float)): s += v
            ws.cell(point, c).value = s if s != 0 else None

        # Total Col
        t_s = 0
        for r in range(data_start, insert_at):
             if ws.cell(r, self.c_Ptr).value == "Total Item Qty":
                 v = ws.cell(r, self.max_col).value
                 if v: t_s += v
        ws.cell(point, self.max_col).value = t_s

        # Format
        self._format_range(ws, point, point, 1, self.max_col, fill_color, True)
        self._format_range(ws, point+1, point+1, 1, self.max_col, fill_color, False)
        ws.row_dimensions[point+1].height = 5

        return 2

    def _insert_po_block(self, ws, insert_at, data_start, fill_color, is_bb):
        # Insert 4 rows (Total, Price1, Price2, Sep)
        rows_to_add = 4
        ws.insert_rows(insert_at, rows_to_add)

        # Copy Labels
        self._copy_row_range(ws, insert_at - 1, insert_at, self.c_Ptr)
        ws.cell(insert_at, self.c_Ptr).value = "Total PO Qty"

        # Copy Prices Meta (Green style rows) logic...
        # VBA: Copies from last data row ("point")
        point = insert_at - 1
        while point >= data_start and ws.cell(point, self.c_Ptr).value != "Total Item Qty":
            point -= 1

        # Copy Price Rows
        self._copy_row_range(ws, point+1, insert_at+1, self.max_col)
        self._copy_row_range(ws, point+2, insert_at+2, self.max_col)

        # Calc Sums
        total_sum1 = 0
        total_sum2 = 0

        for c in range(self.size_col_start, self.max_col):
            qty_sum = 0
            m1 = 0
            m2 = 0

            # Find m1/m2 from last data row
            m1 = ws.cell(point+1, c).value
            m2 = ws.cell(point+2, c).value

            for r in range(data_start, insert_at):
                if ws.cell(r, self.c_Ptr).value == "Total Item Qty":
                     v = ws.cell(r, c).value
                     if v: qty_sum += v

            ws.cell(insert_at, c).value = qty_sum if qty_sum != 0 else None
            ws.cell(insert_at+1, c).value = m1
            ws.cell(insert_at+2, c).value = m2

            if qty_sum and m1: total_sum1 += (qty_sum * float(m1))
            if qty_sum and m2: total_sum2 += (qty_sum * float(m2))

        ws.cell(insert_at+1, self.max_col+1).value = total_sum1
        ws.cell(insert_at+2, self.max_col+1).value = total_sum2

        # Format
        self._format_range(ws, insert_at, insert_at+2, 1, self.max_col, fill_color, True)
        self._format_range(ws, insert_at+1, insert_at+2, self.max_col+1, self.max_col+1, fill_color, True)

        self._format_range(ws, insert_at+3, insert_at+3, 1, self.max_col, COLOR_BLACK, False)
        ws.row_dimensions[insert_at+3].height = 5

        return 4

    def add_new_column(self, ws: Worksheet):
        # Insert E:I (5 cols at col 5)
        ws.insert_cols(5, 5)
        ws.cell(1, 5).value = "Job Number"
        ws.cell(1, 6).value = "Product Type"
        ws.cell(1, 7).value = "VNFOB"
        ws.cell(1, 8).value = "Destination"
        ws.cell(1, 9).value = "Blind Buy Job #"

        # Insert P (1 col at 16) - Note prev columns shifted by 5
        # Old P (16) is now 21.
        # But wait, VBA indices are dynamic.
        # VBA: wsRESULT.Range(ColLett(16) & ":" & ColLett(16)).EntireColumn.Insert
        # This is AFTER the first insert.
        # Let's track shifts.
        shift1 = 5
        # Col 16 becomes 16+5 = 21.
        # So we insert at 21.
        ws.insert_cols(21, 1)
        ws.cell(1, 21).value = "Estimate BusWeekDate"

        self.max_col += 6

        # Update Pointers
        self.c_PO += 5
        self.c_TradingCoPO += 5
        self.c_POLine += 5
        self.c_OGAC += 6
        self.c_DocTypeCode += 6
        self.c_Transportation += 6
        self.c_Ptr += 6
        self.c_ShipToCusNo += 6
        self.c_ShipToCusName += 6
        self.c_Country += 6
        self.c_InventorySegmentCode += 6
        self.c_SubCategoryDesc += 6

        # Fill Data
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, self.c_Style).value:
                ws.cell(r, 7).value = "N" # VNFOB

            ogac_val = ws.cell(r, self.c_OGAC).value
            if ogac_val:
                # Calc Date
                try:
                    s_ogac = str(ogac_val).strip()
                    if len(s_ogac) == 8: # yyyymmdd from sort
                         dt = date(int(s_ogac[:4]), int(s_ogac[4:6]), int(s_ogac[6:8]))
                    elif "/" in s_ogac:
                         # try parse
                         pass # Assume formatted
                    else:
                         dt = None

                    if dt:
                         calc = dt - timedelta(days=10)
                         # Adjust to Monday? Logic: Calc - (Weekday - 1)
                         # Python weekday: Mon=0
                         bus_date = calc - timedelta(days=calc.weekday())
                         ws.cell(r, 21).value = bus_date.strftime("%m/%d/%Y")
                except:
                    pass

            # Blind Buy Job #
            if ws.cell(r, self.c_Ptr).value == "Total PO Qty":
                st = str(ws.cell(r, self.c_Style).value or "").strip().upper()
                if self._blind_buy_contains(st):
                    val = self._get_blind_buy_val(st)
                    ws.cell(r, 9).value = val
                    ws.cell(r+1, 9).value = val
                    ws.cell(r+2, 9).value = val

    def remove_column(self, ws: Worksheet):
        # Remove original rows (Total Item Qty) - Actually wait, VBA removes row label "Total Item Qty" rows
        # "If wsRESULT.Cells(start, colPtr).Value = "Total Item Qty" Then wsRESULT.Rows(start + 1 & ":" & start + 2).Delete"
        # It deletes the Price rows below "Total Item Qty".
        r = 2
        while r <= ws.max_row:
             if ws.cell(r, self.c_Ptr).value == "Total Item Qty":
                 ws.delete_rows(r+1, 2)
                 # Don't increment r, next row fell into r+1
             else:
                 r += 1

        # Formatting borders for Overall Result (max_col + 1)

        # Clean Labels (Remove -XX color code from Style for Totals)
        r = 2
        while r <= ws.max_row:
             lbl = ws.cell(r, self.c_Ptr).value
             if lbl in ["Total Style Qty", "Total PO Qty", "Total OGAC Qty"]:
                 st = str(ws.cell(r, self.c_Style).value or "")
                 if "-" in st:
                     clean = st.split("-")[0]
                     ws.cell(r, self.c_Style).value = clean
                     if lbl == "Total PO Qty":
                         ws.cell(r+1, self.c_Style).value = clean
                         ws.cell(r+2, self.c_Style).value = clean
             r += 1

    def revert_to_obs(self, ws: Worksheet):
        # Fill blanks with #
        for r in range(2, ws.max_row):
            if not ws.cell(r, self.c_ShipToCusNo).value and not ws.cell(r, self.c_ShipToCusName).value:
                lbl = ws.cell(r, self.c_Ptr).value
                if lbl not in ["Total PO Qty", "Total OGAC Qty", "Total Colorway Qty", "Total Style Qty"] and lbl:
                    ws.cell(r, self.c_ShipToCusNo).value = "#"
                    ws.cell(r, self.c_ShipToCusName).value = "#"

            if not ws.cell(r, self.c_TradingCoPO).value and ws.cell(r, self.c_Style).value:
                ws.cell(r, self.c_TradingCoPO).value = "-"

        # Rename Headers
        ws.cell(1, self.c_Style).value = "Material"
        ws.cell(1, self.c_ShipToCusName).value = "Customer Name"
        ws.cell(1, self.c_InventorySegmentCode).value = "AFS Category"
        ws.cell(1, self.c_SubCategoryDesc).value = "Sub Category Size Value"
        ws.cell(1, self.c_DocTypeCode).value = "Buy Group"
        ws.cell(1, self.c_Vendor).value = "Vendor"
        ws.cell(1, self.c_SeasonCode).value = "Planning Season"
        ws.cell(1, self.c_SeasonYear).value = "Year"
        ws.cell(1, self.c_PO).value = "PO Number"
        ws.cell(1, self.c_OGAC).value = "OGAC Date"
        ws.cell(1, self.c_Transportation).value = "Mode"

        # Insert Top Rows
        ws.insert_rows(1, 2)
        ws.cell(2, self.max_col).value = "Overall Result"
        ws.cell(3, self.max_col).value = "TOTAL"

        # Insert Col after Customer Name
        # c_ShipToCusName is index. Insert after means at index+1
        ws.insert_cols(self.c_ShipToCusName + 1)

    # --- MODULE 2 LOGIC ---
    def separate_data_groups(self, ws: Worksheet):
        # 1. First Pass: Identify Separators
        # Loop looking for empty rows (separators)
        # 2. Second Pass: Insert separators between Yellow and White rows

        # Since we modified the sheet structure in RevertToOBS, we need to locate columns again.
        # Data starts at Row 4.

        last_row = ws.max_row
        last_col = ws.max_column
        current_row = 4

        # Pass 2 Logic from VBA (Pass 1 in VBA essentially just looped, pass 2 did the work)
        # "Check if current row is yellow and next row is white" -> Insert Separator (Black)

        while current_row <= last_row:
            # Check for "Total PO Qty" cleanup
            # Column indices shifted. Use simple heuristic or find "Total PO Qty" in label col.
            # Label Col is around 25? Let's search row.

            # Identify Colors
            # OpenPyXL: `cell.fill.start_color.index`
            # Note: Colors might be None if not set.

            # We assume Column A (1) holds the color key as per VBA logic `Cells(currentRow, "A").Interior.Color`
            c_fill = ws.cell(current_row, 1).fill.start_color.index
            n_fill = ws.cell(current_row + 1, 1).fill.start_color.index if current_row < last_row else None

            # Map OpenPyXL Hex to VBA logic
            is_yellow = (c_fill == "FFE6FEB8")
            is_white = (n_fill == "00000000" or n_fill == "FFFFFFFF" or n_fill is None) # OpenPyXL default is often 00000000 (transparent/black) or FFFFFFFF

            if is_yellow and is_white:
                ws.insert_rows(current_row + 1)
                for c in range(1, last_col + 1):
                    ws.cell(current_row + 1, c).fill = COLOR_BLACK
                ws.row_dimensions[current_row + 1].height = 5
                last_row += 1
                current_row += 2
            else:
                current_row += 1

        # Process Data Groups logic
        # Iterate again. Find groups of White Rows between separators.
        # Apply the complex splitting logic.

        # Helper to find data groups
        # A data group is a block of White rows followed by Yellow rows, bounded by Black rows (separators).

        # Due to complexity of insert/delete in-place loop, we scan, identify ranges, and process.

        r = 4
        while r <= ws.max_row:
             # Check if White
             fill = ws.cell(r, 1).fill.start_color.index
             if fill == "00000000" or fill == "FFFFFFFF":
                 # Start of group
                 start_grp = r
                 # Find end
                 while r <= ws.max_row:
                     fill = ws.cell(r, 1).fill.start_color.index
                     if fill == "FF000000": # Black separator
                         break
                     r += 1
                 end_grp = r - 1

                 # Process this group [start_grp, end_grp]
                 # We need to calculate offset if rows are added
                 added = self._process_data_group_logic(ws, start_grp, end_grp)
                 r += added
             else:
                 r += 1

    def _process_data_group_logic(self, ws, start_row, end_row):
        # Identify White rows and Yellow rows in this block
        white_rows = []
        yellow_rows = []
        cyan_found = False

        # Determine the effective "end column" for the copy logic (VBA uses lastCol)
        # In VBA, loop is 31 To endCol - 1.
        # We assume standard width approx 38 columns based on VBA comments (AK=37).
        end_col_idx = ws.max_column

        for r in range(start_row, end_row + 1):
            fill = ws.cell(r, 1).fill.start_color.index
            # OpenPyXL hex colors are ARGB.
            if fill == "FF00FFFF": # Cyan
                cyan_found = True
            elif fill == "FFE6FEB8": # Yellow
                yellow_rows.append(r)
            else: # Assume White
                white_rows.append(r)

        # Check Split Conditions (Plant Code Diff OR Special Logic)
        c_plant = self._find_col_by_header(ws, "Customer Name")
        if c_plant == -1: c_plant = 27 # Fallback

        has_diff = False
        if len(white_rows) > 1:
            val1 = ws.cell(white_rows[0], c_plant).value
            for wr in white_rows[1:]:
                if ws.cell(wr, c_plant).value != val1:
                    has_diff = True
                    break

        if (has_diff and cyan_found) or self._check_special_split_condition(ws, white_rows):
            # Construction of the new block sequence
            new_block = []

            for wr in white_rows:
                # 1. Add White Row
                new_block.append((wr, None))

                # 2. Add Yellow Rows (with specific data copying from White Row)
                for yr in yellow_rows:
                    new_block.append((yr, wr)) # Pass wr as source for overrides

                # 3. Add Separator (Black)
                new_block.append(('SEP', None))

            # Remove trailing separator if it exists
            if new_block and new_block[-1][0] == 'SEP':
                new_block.pop()

            # --- EXECUTION: READ -> DELETE -> WRITE ---

            # 1. Read all data into memory first (because we will delete the source rows)
            temp_data = []

            for item in new_block:
                row_idx, source_white_idx = item

                if row_idx == 'SEP':
                    temp_data.append({'type': 'SEP'})
                else:
                    # Read the row data
                    row_dat = []
                    for c in range(1, ws.max_column + 1):
                        cell = ws.cell(row_idx, c)
                        row_dat.append({
                            'value': cell.value,
                            'fill': cell.fill.copy(),
                            'font': cell.font.copy(),
                            'border': cell.border.copy(),
                            'number_format': cell.number_format
                        })

                    # --- EXACT VBA LOGIC RECREATION ---
                    # If this is a Yellow row (source_white_idx is not None)
                    # VBA: If Not IsEmpty(Cells(newRow, endCol - 1)) Then ... Copy 31 to endCol-1
                    if source_white_idx:
                        # Check column AK (or end_col_idx - 1)
                        # OpenPyXL indices are 1-based.
                        check_col = end_col_idx - 1
                        val_check = ws.cell(row_idx, check_col).value

                        if val_check is not None and str(val_check) != "":
                            # Copy 14 (N) and 19 (S)
                            # Note: Column indices might have shifted in previous macros.
                            # We use direct index copying to match VBA strictness.
                            # VBA: Cells(newRow, 14) = white(1,14)
                            n_val = ws.cell(source_white_idx, 14).value
                            s_val = ws.cell(source_white_idx, 19).value

                            # Update the 'row_dat' list in memory (indices are col-1)
                            if 14 <= len(row_dat): row_dat[13]['value'] = n_val # Col 14
                            if 19 <= len(row_dat): row_dat[18]['value'] = s_val # Col 19

                            # Copy AE (31) to AK (check_col)
                            # VBA loop: For k = 31 To endCol - 1
                            for k in range(31, check_col + 1):
                                if k <= len(row_dat):
                                    w_val = ws.cell(source_white_idx, k).value
                                    row_dat[k-1]['value'] = w_val

                            # VBA: If j > 1 Then (For subsequent yellow rows)...
                            # This part of VBA clears cells if Empty.
                            # In Python reconstruction, we are copying the specific yellow row 'yr',
                            # so if 'yr' was empty in those cols, it stays empty.

                    temp_data.append({'type': 'ROW', 'data': row_dat})

            # 2. Delete the old block
            count = (end_row - start_row) + 1
            ws.delete_rows(start_row, count)

            # 3. Insert new rows
            ws.insert_rows(start_row, len(temp_data))

            # 4. Write data back
            for i, d in enumerate(temp_data):
                r = start_row + i
                if d['type'] == 'SEP':
                    # Black Separator
                    for c in range(1, ws.max_column + 1):
                        ws.cell(r, c).fill = COLOR_BLACK
                    ws.row_dimensions[r].height = 5
                else:
                    # Data Row
                    rd = d['data']
                    for c, cell_data in enumerate(rd):
                        target = ws.cell(r, c+1)
                        target.value = cell_data['value']
                        target.fill = cell_data['fill']
                        target.font = cell_data['font']
                        target.border = cell_data['border']
                        target.number_format = cell_data['number_format']

            return len(temp_data) - count # Return net change in rows

        return 0

    def _check_special_split_condition(self, ws, white_rows):
        # Condition: AB is Mex/Indo/Can AND Y not numeric AND W is VL/TR
        # Or: Y is numeric AND W is VL/TR

        # Map Cols:
        # W (23) -> +7 -> 30?
        # Y (25) -> +7 -> 32?
        # AB (28) -> +7 -> 35?

        # Let's verify shift.
        # Original: A(1)...W(23).
        # Inserted: E-I (5). P(1). Col after ShipName(1).
        # Shift = 7.

        c_W = 30
        c_Y = 32
        c_AB = 35

        for r in white_rows:
            ab = str(ws.cell(r, c_AB).value or "").strip()
            y_val = ws.cell(r, c_Y).value
            w_val = str(ws.cell(r, c_W).value or "").strip()

            is_y_num = isinstance(y_val, (int, float))

            cond1 = (ab in ["Mexico", "Indonesia", "Canada"]) and (not is_y_num) and (w_val in ["VL", "TR"])
            cond2 = is_y_num and (w_val in ["VL", "TR"])

            if cond1 or cond2:
                return True
        return False

    def _find_col_by_header(self, ws, name):
        for c in range(1, ws.max_column + 1):
            if str(ws.cell(3, c).value) == name: # Row 3 has headers
                return c
        return -1

    def _copy_row_range(self, ws, src_r, dest_r, max_c):
        for c in range(1, max_c + 1):
            s = ws.cell(src_r, c)
            d = ws.cell(dest_r, c)
            d.value = s.value
            # d.number_format = s.number_format # Optional, speeds up if skipped

    def _format_range(self, ws, r_start, r_end, c_start, c_end, fill, bold):
        font = Font(bold=True) if bold else Font(bold=False)
        for r in range(r_start, r_end + 1):
            for c in range(c_start, c_end + 1):
                cell = ws.cell(r, c)
                cell.fill = fill
                if bold: cell.font = font