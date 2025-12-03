import os
import re
import time
import threading
from typing import List, Tuple, Any

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
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font, NamedStyle
from openpyxl.utils import get_column_letter

try:
    import xlrd  # for legacy excel
except Exception:
    xlrd = None


class MainWidget(QWidget):
    log_message = Signal(str)
    report_message = Signal(str)
    processing_done = Signal(int, int, str)

    def __init__(self):
        super().__init__()
        self.setObjectName("mmu_widget")

        self._build_ui()
        self._connect_signals()

    # UI
    def _build_ui(self):
        # Description
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

        # Buttons
        self.select_master_btn = PrimaryPushButton("Select Master Excel Files", self)
        self.select_ppm_btn = PrimaryPushButton("Select PPM Report Excel Files", self)
        self.run_btn = PrimaryPushButton("Run", self)

        # Labels
        self.master_files_label = QLabel("Master file(s)", self)
        self.master_files_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.master_files_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")
        self.ppm_files_label = QLabel("PPM report file(s)", self)
        self.ppm_files_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.ppm_files_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")
        self.logs_label = QLabel("Process logs", self)
        self.logs_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.logs_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")
        self.reports_label = QLabel("Report output", self)
        self.reports_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.reports_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        # Text boxes
        shared_style = (
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        self.master_files_box = QTextEdit(self)
        self.master_files_box.setReadOnly(True)
        self.master_files_box.setPlaceholderText("Selected master file(s) will appear here")
        self.master_files_box.setStyleSheet(shared_style)

        self.ppm_files_box = QTextEdit(self)
        self.ppm_files_box.setReadOnly(True)
        self.ppm_files_box.setPlaceholderText("Selected PPM report file(s) will appear here")
        self.ppm_files_box.setStyleSheet(shared_style)

        self.reports_box = QTextEdit(self)
        self.reports_box.setReadOnly(True)
        self.reports_box.setPlaceholderText("Reports will appear here")
        self.reports_box.setStyleSheet(shared_style)

        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Live process log will appear here")
        self.log_box.setStyleSheet(shared_style)

        # Layouts
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        main_layout.addWidget(self.desc_label, 1)

        row1 = QHBoxLayout()
        row1.addWidget(self.select_master_btn, 1)
        row1.addWidget(self.select_ppm_btn, 1)
        # add another button for pps
        main_layout.addLayout(row1, 0)

        row2 = QHBoxLayout()
        row2.addStretch(1)
        row2.addWidget(self.run_btn, 1)
        row2.addStretch(1)
        main_layout.addLayout(row2, 0)

        row3 = QHBoxLayout()
        row3.addWidget(self.master_files_label, 1)
        row3.addWidget(self.ppm_files_label, 1)
        main_layout.addLayout(row3, 0)

        row4 = QHBoxLayout()
        row4.addWidget(self.master_files_box, 1)
        row4.addWidget(self.ppm_files_box, 1)
        main_layout.addLayout(row4, 3)

        row5 = QHBoxLayout()
        row5.addWidget(self.reports_label, 1)
        row5.addWidget(self.logs_label, 1)
        main_layout.addLayout(row5, 0)

        row6 = QHBoxLayout()
        row6.addWidget(self.reports_box, 1)
        row6.addWidget(self.log_box, 1)
        main_layout.addLayout(row6, 4)

    def set_long_description(self, text: str):
        clean = (text or "").strip()
        if clean:
            self.desc_label.setText(clean)
            self.desc_label.show()
        else:
            self.desc_label.clear()
            self.desc_label.hide()

    def _connect_signals(self):
        self.select_master_btn.clicked.connect(self.select_master_files)
        self.select_ppm_btn.clicked.connect(self.select_ppm_files)
        self.run_btn.clicked.connect(self.run_process)
        self.log_message.connect(self.append_log)
        self.report_message.connect(self.append_report)
        self.processing_done.connect(self.on_processing_done)

    # Functions
    def select_master_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select master Excel files")
        if files:
            self.master_files_box.setPlainText("\n".join(files))
        else:
            self.master_files_box.clear()

    def select_ppm_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select PPM report Excel files")
        if files:
            files = sorted(files, key=lambda x: os.path.basename(x))
            self.ppm_files_box.setPlainText("\n".join(files))
        else:
            self.ppm_files_box.clear()

    def _selected_master_files(self) -> List[str]:
        text = self.master_files_box.toPlainText().strip()
        if not text:
            return []
        return [line for line in text.split("\n") if line.strip()]

    def _selected_ppm_files(self) -> List[str]:
        text = self.ppm_files_box.toPlainText().strip()
        if not text:
            return []
        return [line for line in text.split("\n") if line.strip()]

    def run_process(self):
        master_files = self._selected_master_files()
        ppm_files = self._selected_ppm_files()

        if not master_files or not ppm_files:
            MessageBox("Warning", "Please select both master file(s) and PPM report file(s).", self).exec()
            return

        self.log_box.clear()
        self.reports_box.clear()
        self.log_message.emit("Process starts")
        self.run_btn.setEnabled(False)
        self.select_master_btn.setEnabled(False)
        self.select_ppm_btn.setEnabled(False)

        def worker():
            ok, fail, out_path = 0, 0, ""
            try:
                out_path, ok, fail = process_files(master_files, ppm_files, self.log_message.emit, self.report_message.emit)
            except Exception as e:
                self.log_message.emit(f"ERROR: {e}")
            self.processing_done.emit(ok, fail, out_path)

        threading.Thread(target=worker, daemon=True).start()

    def append_log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def append_report(self, text: str):
        self.reports_box.append(text)
        self.reports_box.ensureCursorVisible()

    def on_processing_done(self, ok: int, fail: int, out_path: str):
        if out_path:
            self.log_message.emit(f"Output workbook saved to: {out_path}")
        self.log_message.emit(f"Completed: {ok} success, {fail} failed.")
        self.run_btn.setEnabled(True)
        self.select_master_btn.setEnabled(True)
        self.select_ppm_btn.setEnabled(True)


def get_widget():
    return MainWidget()


# ========================
# Memo Match Utility core
# ========================


def process_files(master_files: List[str], report_files: List[str], log_cb=None, report_cb=None) -> Tuple[str, int, int]:
    """Core processing logic ported from the legacy MMU app."""
    log = log_cb or (lambda msg: None)
    report = report_cb or (lambda msg: None)

    start_time = time.time()
    master_columns_needed = [
        "NK SAP PO",
        "PO LINE ITEM",
        "CM FOB rec. date",
        "FINAL FOB (Regular sizes)",
        "FINAL FOB (Extended sizes)",
        "Extended Sizes",
        "DPOM - Incorrect FOB",
        "Price (Date)",
        "Price (Changes)",
        "Season",
        "Season Year",
        "Season (Date)",
        "Season (Changes)",
        "FG QTY",
        "FG QTY (Date)",
        "FG QTY (Changes)",
        "Doc Type",
        "Doc Type (Date)",
        "Doc Type (Changes)",
        "SHIP MODE",
        "SHIP MODE (Date)",
        "SHIP MODE (Changes)",
        "Plant Code",
        "Plant Code (Date)",
        "Plant Code (Changes)",
        "SHIP-TO",
        "SHIP-TO (Date)",
        "SHIP-TO (Changes)",
        "AFS Cat",
        "AFS Cat (Date)",
        "AFS Cat (Changes)",
        "VAS name",
        "VAS name (Date)",
        "VAS name (Changes)",
        "Hanger size",
        "Hanger size (Date)",
        "Hanger size (Changes)",
        "Ratio Qty",
        "Ratio Qty (Date)",
        "Ratio Qty (Changes)",
        "Customer PO (Deichmann Group only)",
        "Customer PO (Date)",
        "Customer PO (Changes)",
        "Latest CM Change Date",
        "JOB NO"
    ]
    ppm_columns_needed = [
        "Purchase Order Number",
        "PO Line Item Number",
        "Product Code",
        "Gross Price/FOB currency code",
        "Surcharge Min Mat Main Body currency code",
        "Surcharge Min Material Trim currency code",
        "Surcharge Misc currency code",
        "Surcharge VAS currency code",
        "Gross Price/FOB",
        "Surcharge Min Mat Main Body",
        "Surcharge Min Material Trim",
        "Surcharge Misc",
        "Surcharge VAS",
        "Size Description",
        "Planning Season Code",
        "Planning Season Year",
        "Total Item Quantity",
        "Doc Type",
        "Mode of Transportation Code",
        "Plant Code",
        "Ship To Customer Number",
        "Inventory Segment Code",
        "VAS name",
        "Hanger size",
        "Ratio quantity",
        "Customer PO",
        "Change Date",
        "GAC",
        "DPOM Line Item Status",
        "Document Date"
    ]
    output_texts = set()

    def find_columns_header(sheet, needed_columns):
        header_row_index = None
        for row in sheet.iter_rows(max_col=sheet.max_column):
            for cell in row:
                if cell.value:
                    for key in needed_columns:
                        if key in str(cell.value):
                            if header_row_index is None:
                                header_row_index = cell.row
                                return header_row_index

    def find_columns_master(sheet, needed_columns, header):
        column_positions = {column: None for column in needed_columns}
        column_positions["FINAL FOB (Extended sizes)"] = []
        column_positions["Season"] = []
        column_positions["FG QTY"] = []
        column_positions["Doc Type"] = []
        column_positions["SHIP MODE"] = []
        column_positions["Plant Code"] = []
        column_positions["SHIP-TO"] = []
        column_positions["AFS Cat"] = []
        column_positions["VAS name"] = []
        column_positions["Hanger size"] = []
        column_positions["Ratio Qty"] = []

        for row in sheet.iter_rows(min_row=header, max_row=header, max_col=sheet.max_column):
            for cell in row:
                if cell.value:
                    for key in needed_columns:
                        if str(key).strip().lower() in str(cell.value).strip().lower():
                            if key == "FINAL FOB (Extended sizes)":
                                column_positions[key].append(cell.column)
                            elif key == "Season":
                                column_positions[key].append(cell.column)
                            elif key == "FG QTY":
                                column_positions[key].append(cell.column)
                            elif key == "Doc Type":
                                column_positions[key].append(cell.column)
                            elif key == "SHIP MODE":
                                column_positions[key].append(cell.column)
                            elif key == "Plant Code":
                                column_positions[key].append(cell.column)
                            elif key == "SHIP-TO":
                                column_positions[key].append(cell.column)
                            elif key == "AFS Cat":
                                column_positions[key].append(cell.column)
                            elif key == "VAS name":
                                column_positions[key].append(cell.column)
                            elif key == "Hanger size":
                                column_positions[key].append(cell.column)
                            elif key == "Ratio Qty":
                                column_positions[key].append(cell.column)
                            else:
                                column_positions[key] = cell.column
        return column_positions

    def find_columns_report(sheet, needed_columns, header):
        column_positions = {column: None for column in needed_columns}
        for row in sheet.iter_rows(min_row=header, max_row=header, max_col=sheet.max_column):
            for cell in row:
                if cell.value:
                    for key in needed_columns:
                        if str(key).strip().lower() in str(cell.value).strip().lower():
                            column_positions[key] = cell.column
        return column_positions

    def safe_float(value):
        try:
            return float(value)
        except (TypeError, ValueError):
            return float(0)

    def letter(n):
        result = ""
        while n > 0:
            n -= 1
            result = chr(n % 26 + 65) + result
            n //= 26
        return result

    def append_date_changes(master_row, master_cols, report_cols, row, change_type, report_value, output_text=None):
        ex_date = str(master_row[master_cols[f"{change_type} (Date)"] - 1].value).strip() if not isEmptyCell(master_row[master_cols[f"{change_type} (Date)"] - 1].value) else ""
        ex_change = str(master_row[master_cols[f"{change_type} (Changes)"] - 1].value).strip() if not isEmptyCell(master_row[master_cols[f"{change_type} (Changes)"] - 1].value) else ""

        current_date = str(row[report_cols["Change Date"] - 1].value.strftime('%m/%d')).strip()
        current_date_mmddyy = str(row[report_cols["Change Date"] - 1].value.strftime('%m/%d/%y')).strip()
        current_date_mdyy = str(row[report_cols["Change Date"] - 1].value.strftime('%m/%d/%y')).strip().replace('/0', '/').lstrip('0')
        current_change = str(report_value).strip()

        if ex_date and ex_change:
            if "," in ex_date and "," in ex_change:
                ex_dates = [d.strip() for d in ex_date.split(",")]
                ex_changes = [c.strip() for c in ex_change.split(",")]
                existing_pairs = list(zip(ex_dates, ex_changes))
            else:
                existing_pairs = [(ex_date, ex_change)] if ex_date and ex_change else []

            current_pair = (current_date_mdyy, current_change)
            if current_pair not in existing_pairs:
                existing_pairs.append(current_pair)

            if existing_pairs:
                ex_dates, ex_changes = zip(*existing_pairs)
                ex_date = " , ".join(ex_dates)
                ex_change = " , ".join(ex_changes)
        else:
            ex_date = current_date_mdyy
            ex_change = current_change

        master_row[master_cols[f"{change_type} (Date)"] - 1].value = ex_date
        master_row[master_cols[f"{change_type} (Changes)"] - 1].value = ex_change
        master_row[master_cols["Latest CM Change Date"] - 1].value = current_date_mmddyy

        if output_text and output_text not in output_texts:
            output_texts.add(output_text)
            report(output_text)

        ex_fob_date = str(master_row[master_cols["CM FOB rec. date"] - 1].value).strip()
        if not isEmptyCell(ex_fob_date):
            ex_fob_dates = [d.strip() for d in str(ex_fob_date).split(",")]
            if current_date_mdyy not in ex_fob_dates:
                ex_fob_dates.append(current_date_mdyy)
                ex_fob_date = " , ".join(ex_fob_dates)
        else:
            ex_fob_date = current_date_mdyy

        master_row[master_cols["CM FOB rec. date"] - 1].value = ex_fob_date

    def keep_date_changes(master_row, master_cols, report_cols, row, change_type, report_value, key):
        ex_date = ori_date[change_type].get(f"{key}", "")
        ex_change = ori_change[change_type].get(f"{key}", "")

        current_date = str(row[report_cols["Change Date"] - 1].value.strftime('%m/%d')).strip()
        current_date_mmddyy = str(row[report_cols["Change Date"] - 1].value.strftime('%m/%d/%y')).strip()
        current_date_mdyy = str(row[report_cols["Change Date"] - 1].value.strftime('%m/%d/%y')).strip().replace('/0', '/').lstrip('0')
        current_change = str(report_value).strip()

        if ex_date and ex_change:
            if "," in ex_date and "," in ex_change:
                ex_dates = [d.strip() for d in ex_date.split(",")]
                ex_changes = [c.strip() for c in ex_change.split(",")]
                existing_pairs = list(zip(ex_dates, ex_changes))
            else:
                existing_pairs = [(ex_date, ex_change)] if ex_date and ex_change else []

            current_pair = (current_date_mdyy, current_change)
            if current_pair in existing_pairs:
                existing_pairs.remove(current_pair)

            if existing_pairs:
                ex_dates, ex_changes = zip(*existing_pairs)
                ex_date = " , ".join(ex_dates)
                ex_change = " , ".join(ex_changes)

        master_row[master_cols[f"{change_type} (Date)"] - 1].value = ex_date
        master_row[master_cols[f"{change_type} (Changes)"] - 1].value = ex_change

    def empty_date_changes(master_row, master_cols, report_cols, row, change_type, key):
        original_date = ori_date[change_type].get(f"{key}", "")
        original_change = ori_change[change_type].get(f"{key}", "")
        master_row[master_cols[f"{change_type} (Date)"] - 1].value = original_date
        master_row[master_cols[f"{change_type} (Changes)"] - 1].value = original_change

    def is_number(value):
        try:
            float(value)
            return True
        except ValueError:
            return False

    def isEmptyCell(cell):
        cell = str(cell).strip()
        if cell == "None" or not cell or cell == "-":
            return True
        else:
            return False

    def isNumberAfterDash(val):
        if '-' in val:
            parts = val.split('-')
            return all(is_number(part.strip()) for part in parts[1:])
        return False

    def hasComma(val):
        return ',' in str(val).strip()

    report_date_now = 0
    report_date_bfr = 0
    log("Starts processing...")
    total_data = 1
    master_dict = {}
    master_dict_data = {}
    newPO = {}
    existingPO = {}
    work_path = ""
    ok = 0
    fail = 0
    sizes = ["2XS", "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "2XSS", "XSS", "SS", "MS", "LS", "XLS", "2XLS", "3XLS", "4XLS", "5XLS", "2XSL", "XSL", "SL", "ML", "XLL", "2XLL", "3XLL", "4XLL", "5XLL", "2XST", "XST", "ST", "MT", "LT", "XLT", "2XLT", "3XLT", "4XLT", "5XLT", "2XSTT", "XSTT", "STT", "MTT", "LTT", "XLTT", "2XLTT", "3XLTT", "4XLTT", "5XLTT", "0X", "1X", "2X", "3X", "4X", "5X", "0XT", "1XT", "2XT", "3XT", "4XT", "5XT", "0XTT", "1XTT", "2XTT", "3XTT", "4XTT", "5XTT", "CUST0", "CUST1", "CUST2", "CUST3", "CUST4", "CUST5", "CUST6", "CUST7", "CUST2XS", "CUSTXS", "CUSTS", "CUSTM", "CUSTL", "CUSTXL", "CUST2XL", "CUST3XL", "CUST4XL", "CUST5XL"]

    for master_current, master_path in enumerate(master_files, start=1):
        try:
            work_path = os.path.dirname(master_path)
            log(f"Reading {os.path.basename(master_path)}...")
            master_wb = load_workbook(master_path)
            master_sheet = master_wb[master_wb.sheetnames[0]]
            master_wb_data = load_workbook(master_path, read_only=True, data_only=True)
            master_sheet_data = master_wb_data[master_wb_data.sheetnames[0]]
            master_header_indicator = ["NK SAP PO"]
            log("Finding header row...")
            master_header_row = find_columns_header(master_sheet, master_header_indicator)
            log(f"Header row found at {master_header_row}.")
            master_cols = find_columns_master(master_sheet, master_columns_needed, master_header_row)

            report_text = (
                f"Master Column Detection Report:\n"
                f"Header Row: {master_header_row}\n"
                f"NK SAP PO: {letter(master_cols['NK SAP PO'])}\n"
                f"PO LINE ITEM: {letter(master_cols['PO LINE ITEM'])}\n"
                f"CM FOB rec. date: {letter(master_cols['CM FOB rec. date'])}\n"
                f"FINAL FOB (Regular sizes): {letter(master_cols['FINAL FOB (Regular sizes)'])}\n"
                f"FINAL FOB (Extended sizes 1): {letter(master_cols['FINAL FOB (Extended sizes)'][0])}\n"
                f"FINAL FOB (Extended sizes 2): {letter(master_cols['FINAL FOB (Extended sizes)'][1])}\n"
                f"Extended Sizes 1: {letter(master_cols['Extended Sizes']-2)}\n"
                f"Extended Sizes 2: {letter(master_cols['Extended Sizes'])}\n"
                f"DPOM - Incorrect FOB: {letter(master_cols['DPOM - Incorrect FOB'])}\n"
                f"Season: {letter(master_cols['Season'][0])}\n"
                f"Season Year: {letter(master_cols['Season Year'])}\n"
                f"Season (Date): {letter(master_cols['Season (Date)'])}\n"
                f"Season (Changes): {letter(master_cols['Season (Changes)'])}\n"
                f"FG QTY: {letter(master_cols['FG QTY'][0])}\n"
                f"FG QTY (Date): {letter(master_cols['FG QTY (Date)'])}\n"
                f"FG QTY (Changes): {letter(master_cols['FG QTY (Changes)'])}\n"
                f"Doc Type: {letter(master_cols['Doc Type'][0])}\n"
                f"Doc Type (Date): {letter(master_cols['Doc Type (Date)'])}\n"
                f"Doc Type (Changes): {letter(master_cols['Doc Type (Changes)'])}\n"
                f"SHIP MODE: {letter(master_cols['SHIP MODE'][0])}\n"
                f"SHIP MODE (Date): {letter(master_cols['SHIP MODE (Date)'])}\n"
                f"SHIP MODE (Changes): {letter(master_cols['SHIP MODE (Changes)'])}\n"
                f"Plant Code: {letter(master_cols['Plant Code'][0])}\n"
                f"Plant Code (Date): {letter(master_cols['Plant Code (Date)'])}\n"
                f"Plant Code (Changes): {letter(master_cols['Plant Code (Changes)'])}\n"
                f"SHIP-TO: {letter(master_cols['SHIP-TO'][0])}\n"
                f"SHIP-TO (Date): {letter(master_cols['SHIP-TO (Date)'])}\n"
                f"SHIP-TO (Changes): {letter(master_cols['SHIP-TO (Changes)'])}\n"
                f"AFS Cat: {letter(master_cols['AFS Cat'][0])}\n"
                f"AFS Cat (Date): {letter(master_cols['AFS Cat (Date)'])}\n"
                f"AFS Cat (Changes): {letter(master_cols['AFS Cat (Changes)'])}\n"
                f"VAS name: {letter(master_cols['VAS name'][0])}\n"
                f"VAS name (Date): {letter(master_cols['VAS name (Date)'])}\n"
                f"VAS name (Changes): {letter(master_cols['VAS name (Changes)'])}\n"
                f"Hanger size: {letter(master_cols['Hanger size'][0])}\n"
                f"Hanger size (Date): {letter(master_cols['Hanger size (Date)'])}\n"
                f"Hanger size (Changes): {letter(master_cols['Hanger size (Changes)'])}\n"
                f"Ratio Qty: {letter(master_cols['Ratio Qty'][0])}\n"
                f"Ratio Qty (Date): {letter(master_cols['Ratio Qty (Date)'])}\n"
                f"Ratio Qty (Changes): {letter(master_cols['Ratio Qty (Changes)'])}\n"
                f"Customer PO (Deichmann Group only): {letter(master_cols['Customer PO (Deichmann Group only)'])}\n"
                f"Customer PO (Date): {letter(master_cols['Customer PO (Date)'])}\n"
                f"Customer PO (Changes): {letter(master_cols['Customer PO (Changes)'])}\n"
                f"Latest CM Change Date: {letter(master_cols['Latest CM Change Date'])}\n"
                f"JOB NO: {letter(master_cols['JOB NO'])}\n"
                f"\n"
            )
            log(report_text)

            master_dict[master_current] = {}
            for master_row in master_sheet.iter_rows(min_row=master_header_row + 1, max_row=master_sheet.max_row):
                master_po_num = str(master_row[master_cols["NK SAP PO"] - 1].value).strip()
                master_po_line = str(master_row[master_cols["PO LINE ITEM"] - 1].value).strip()
                master_dict[master_current].setdefault((master_po_num, master_po_line), []).append(master_row)

            ori_date = {
                "Season": {},
                "FG QTY": {},
                "Doc Type": {},
                "SHIP MODE": {},
                "Plant Code": {},
                "SHIP-TO": {},
                "AFS Cat": {},
                "VAS name": {},
                "Hanger size": {},
                "Ratio Qty": {},
                "Customer PO": {},
                "Price": {},
                "Currency": {},
            }
            ori_change = {
                "Season": {},
                "FG QTY": {},
                "Doc Type": {},
                "SHIP MODE": {},
                "Plant Code": {},
                "SHIP-TO": {},
                "AFS Cat": {},
                "VAS name": {},
                "Hanger size": {},
                "Ratio Qty": {},
                "Customer PO": {},
                "Price": {},
                "Currency": {},
            }
            ori_cm_date = {}

            keys = ["Season", "FG QTY", "Doc Type", "SHIP MODE", "Plant Code",
                    "SHIP-TO", "AFS Cat", "VAS name", "Hanger size", "Ratio Qty",
                    "Customer PO", "Price"]

            master_dict_data[master_current] = {}
            for master_row_data in master_sheet_data.iter_rows(min_row=master_header_row + 1, max_row=master_sheet.max_row):
                master_po_num = str(master_row_data[master_cols["NK SAP PO"] - 1].value).strip()
                master_po_line = str(master_row_data[master_cols["PO LINE ITEM"] - 1].value).strip()
                master_po_job = str(master_row_data[master_cols["JOB NO"] - 1].value).strip()

                if f"{master_po_num}{master_po_line}{master_po_job}" not in ori_date["Season"]:
                    for key in keys:
                        date_key = f"{key} (Date)"
                        change_key = f"{key} (Changes)"

                        ori_date[key][f"{master_po_num}{master_po_line}{master_po_job}"] = (
                            str(master_row_data[master_cols[date_key] - 1].value).strip()
                            if not isEmptyCell(master_row_data[master_cols[date_key] - 1].value)
                            else ""
                        )
                        ori_change[key][f"{master_po_num}{master_po_line}{master_po_job}"] = (
                            str(master_row_data[master_cols[change_key] - 1].value).strip()
                            if not isEmptyCell(master_row_data[master_cols[change_key] - 1].value)
                            else ""
                        )

                    ori_cm_date[f"{master_po_num}{master_po_line}{master_po_job}"] = (
                        str(master_row_data[master_cols["Latest CM Change Date"] - 1].value).strip()
                        if not isEmptyCell(master_row_data[master_cols["Latest CM Change Date"] - 1].value)
                        else ""
                    )

                master_dict_data[master_current].setdefault((master_po_num, master_po_line), []).append(master_row_data)

            for report_path in report_files:
                matching_rows = {
                    "Season": [],
                    "FG QTY": [],
                    "Doc Type": [],
                    "SHIP MODE": [],
                    "Plant Code": [],
                    "SHIP-TO": [],
                    "AFS Cat": [],
                    "VAS name": [],
                    "Hanger size": [],
                    "Ratio Qty": [],
                    "Customer PO": [],
                    "Price": [],
                    "Currency": [],
                }
                log(f"Reading {os.path.basename(report_path)}...")

                report_wb = load_workbook(report_path, read_only=True, data_only=True)
                report_sheet = report_wb.active

                report_header_indicator = ["Purchase Order Number"]
                log("Finding report header row...")
                report_header_row = find_columns_header(report_sheet, report_header_indicator)
                log(f"Report header row found at {report_header_row}.")
                report_cols = find_columns_report(report_sheet, ppm_columns_needed, report_header_row)

                report_text = (
                    f"PPM Report Column Detection Report:\n"
                    f"Header Row: {report_header_row}\n"
                    f"Purchase Order Number: {letter(report_cols['Purchase Order Number'])}\n"
                    f"PO Line Item Number: {letter(report_cols['PO Line Item Number'])}\n"
                    f"Product Code: {letter(report_cols['Product Code'])}\n"
                    f"Gross Price/FOB currency code: {letter(report_cols['Gross Price/FOB currency code'])}\n"
                    f"Surcharge Min Mat Main Body currency code: {letter(report_cols['Surcharge Min Mat Main Body currency code'])}\n"
                    f"Surcharge Min Material Trim currency code: {letter(report_cols['Surcharge Min Material Trim currency code'])}\n"
                    f"Surcharge Misc currency code: {letter(report_cols['Surcharge Misc currency code'])}\n"
                    f"Surcharge VAS currency code: {letter(report_cols['Surcharge VAS currency code'])}\n"
                    f"Gross Price/FOB: {letter(report_cols['Gross Price/FOB']-1)}\n"
                    f"Surcharge Min Mat Main Body: {letter(report_cols['Surcharge Min Mat Main Body']-1)}\n"
                    f"Surcharge Min Material Trim: {letter(report_cols['Surcharge Min Material Trim']-1)}\n"
                    f"Surcharge Misc: {letter(report_cols['Surcharge Misc']-1)}\n"
                    f"Surcharge VAS: {letter(report_cols['Surcharge VAS']-1)}\n"
                    f"Size Description: {letter(report_cols['Size Description'])}\n"
                    f"Planning Season Code: {letter(report_cols['Planning Season Code'])}\n"
                    f"Planning Season Year: {letter(report_cols['Planning Season Year'])}\n"
                    f"Total Item Quantity: {letter(report_cols['Total Item Quantity'])}\n"
                    f"Doc Type: {letter(report_cols['Doc Type']-1)}\n"
                    f"Mode of Transportation Code: {letter(report_cols['Mode of Transportation Code'])}\n"
                    f"Plant Code: {letter(report_cols['Plant Code'])}\n"
                    f"Ship To Customer Number: {letter(report_cols['Ship To Customer Number'])}\n"
                    f"Inventory Segment Code: {letter(report_cols['Inventory Segment Code'])}\n"
                    f"VAS name: {letter(report_cols['VAS name'])}\n"
                    f"Hanger size: {letter(report_cols['Hanger size'])}\n"
                    f"Ratio quantity: {letter(report_cols['Ratio quantity'])}\n"
                    f"Customer PO: {letter(report_cols['Customer PO'])}\n"
                    f"Change Date: {letter(report_cols['Change Date'])}\n"
                    f"GAC: {letter(report_cols['GAC']+1)}\n"
                    f"DPOM Line Item Status: {letter(report_cols['DPOM Line Item Status'])}\n"
                    f"Document Date: {letter(report_cols['Document Date'])}\n"
                    f"\n"
                )

                log(report_text)

                for row_index, row in enumerate(report_sheet.iter_rows(min_row=report_header_row + 1, max_row=report_sheet.max_row), start=report_header_row + 1):
                    if report_date_bfr == 0:
                        report_date_bfr = str(row[report_cols["Change Date"] - 1].value.strftime('%m/%d')).strip()
                    if row_index == (report_header_row + 1):
                        report_date_now = str(row[report_cols["Change Date"] - 1].value.strftime('%m/%d')).strip()

                    total_data += 1
                    report_po_num = str(row[report_cols["Purchase Order Number"] - 1].value).strip()
                    report_po_line = str(row[report_cols["PO Line Item Number"] - 1].value).strip()

                    if isEmptyCell(row[report_cols["Product Code"] - 1].value):
                        report_po_style = "INVALID"
                    else:
                        report_po_style = (row[report_cols["Product Code"] - 1].value).split('-')[0].strip()
                    key = (report_po_num, report_po_line)

                    if key in master_dict[master_current] and report_po_style != "INVALID":
                        existingPO[(report_po_style, report_po_num, report_po_line)] = row
                        for keyPO in list(newPO.keys()):
                            if keyPO[:3] == (report_po_style, report_po_num, report_po_line):
                                del newPO[keyPO]
                        master_rows = master_dict[master_current][key]
                        master_rows_data = master_dict_data[master_current][key]

                        for i, (master_row, master_row_data) in enumerate(zip(master_rows, master_rows_data)):
                            noDiscrepancy = True
                            master_po_fob = safe_float(master_row_data[master_cols["FINAL FOB (Regular sizes)"] - 1].value)
                            master_po_fob_ex_1 = safe_float(master_row_data[master_cols['FINAL FOB (Extended sizes)'][0] - 1].value)
                            master_po_fob_ex_2 = safe_float(master_row_data[master_cols['FINAL FOB (Extended sizes)'][1] - 1].value)

                            report_po_fob = safe_float(row[report_cols["Gross Price/FOB"] - 2].value) + safe_float(row[report_cols["Surcharge Min Mat Main Body"] - 2].value) + safe_float(row[report_cols["Surcharge Min Material Trim"] - 2].value) + safe_float(row[report_cols["Surcharge Misc"] - 2].value) + safe_float(row[report_cols["Surcharge VAS"] - 2].value)

                            if str(row[report_cols["Mode of Transportation Code"] - 1].value) == "VL":
                                report_ship_mode = "SEA"
                            elif str(row[report_cols["Mode of Transportation Code"] - 1].value) == "AF":
                                report_ship_mode = "NAF"
                            elif str(row[report_cols["Mode of Transportation Code"] - 1].value) in ("TR", "TRUCK"):
                                report_ship_mode = "TR"
                            else:
                                report_ship_mode = str(row[report_cols["Mode of Transportation Code"] - 1].value)

                            if str(row[report_cols["Inventory Segment Code"] - 1].value) == "1000":
                                report_afs_cat = "01000"
                            else:
                                report_afs_cat = str(row[report_cols["Inventory Segment Code"] - 1].value)

                            if isEmptyCell(master_row_data[master_cols["Ratio Qty"][0] - 1].value):
                                master_ratio_qty = "0"
                            else:
                                master_ratio_qty = str(master_row_data[master_cols["Ratio Qty"][0] - 1].value)

                            if isEmptyCell(row[report_cols["Ratio quantity"] - 1].value):
                                report_ratio_qty = "0"
                            else:
                                report_ratio_qty = str(row[report_cols["Ratio quantity"] - 1].value)

                            total_master_fg_qty = 0
                            for master_row_data_item in master_rows_data:
                                if isEmptyCell(master_row_data_item[master_cols["FG QTY"][0] - 1].value):
                                    row_fg_qty = 0
                                else:
                                    row_fg_qty = safe_float(master_row_data_item[master_cols["FG QTY"][0] - 1].value)
                                total_master_fg_qty += row_fg_qty
                            master_fg_qty = str(total_master_fg_qty).strip()

                            master_po_fob = f"{master_po_fob:.2f}"
                            master_po_fob_ex_1 = f"{master_po_fob_ex_1:.2f}"
                            master_po_fob_ex_2 = f"{master_po_fob_ex_2:.2f}"
                            report_po_fob = f"{report_po_fob:.2f}"

                            if (row[report_cols["Planning Season Code"] - 1].value != master_row_data[master_cols["Season"][0] - 1].value) or (row[report_cols["Planning Season Year"] - 1].value != master_row_data[master_cols["Season Year"] - 1].value):
                                noDiscrepancy = False
                                output_text = f"Season diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Season {row[report_cols['Planning Season Code'] - 1].value} {row[report_cols['Planning Season Year'] - 1].value} vs OCCC Season {master_row_data[master_cols['Season'][0] - 1].value} {master_row_data[master_cols['Season Year'] - 1].value}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "Season", f"{str(row[report_cols['Planning Season Code'] - 1].value)}{str(row[report_cols['Planning Season Year'] - 1].value)}", output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "Season", f"{str(row[report_cols['Planning Season Code'] - 1].value)}{str(row[report_cols['Planning Season Year'] - 1].value)}", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["Season"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["Season"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "Season", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if str(safe_float(row[report_cols["Total Item Quantity"] - 1].value)) != str(master_fg_qty):
                                noDiscrepancy = False
                                output_text = f"Total Item Quantity diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Total Item Quantity {str(row[report_cols['Total Item Quantity'] -1].value)} vs OCCC FG QTY {str(master_fg_qty)}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "FG QTY", row[report_cols["Total Item Quantity"] - 1].value, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "FG QTY", row[report_cols["Total Item Quantity"] - 1].value, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["FG QTY"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["FG QTY"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "FG QTY", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if str(row[report_cols["Doc Type"] - 2].value) != str(master_row_data[master_cols["Doc Type"][0] - 1].value):
                                noDiscrepancy = False
                                output_text = f"Document Type diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Doc Type {row[report_cols['Doc Type'] - 2].value} vs OCCC Doc Type {master_row_data[master_cols['Doc Type'][0] - 1].value}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "Doc Type", row[report_cols["Doc Type"] - 2].value, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "Doc Type", row[report_cols["Doc Type"] - 2].value, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["Doc Type"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["Doc Type"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "Doc Type", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if str(report_ship_mode) != str(master_row_data[master_cols["SHIP MODE"][0] - 1].value):
                                noDiscrepancy = False
                                output_text = f"SHIP MODE diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Mode of Transportation Code {report_ship_mode} vs OCCC SHIP MODE {master_row_data[master_cols['SHIP MODE'][0] - 1].value}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "SHIP MODE", report_ship_mode, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "SHIP MODE", report_ship_mode, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                            matching_rows["SHIP MODE"].append(i)
                            if report_path == report_files[-1] and i == matching_rows["SHIP MODE"][-1]:
                                empty_date_changes(master_row, master_cols, report_cols, row, "SHIP MODE", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if str(row[report_cols["Plant Code"] - 1].value) != str(master_row_data[master_cols["Plant Code"][0] - 1].value):
                                noDiscrepancy = False
                                output_text = f"Plant Code diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Plant Code {row[report_cols['Plant Code'] - 1].value} vs OCCC Plant Code {master_row_data[master_cols['Plant Code'][0] - 1].value}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "Plant Code", row[report_cols["Plant Code"] - 1].value, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "Plant Code", row[report_cols["Plant Code"] - 1].value, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["Plant Code"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["Plant Code"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "Plant Code", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if str(row[report_cols["Ship To Customer Number"] - 1].value) != str(master_row_data[master_cols["SHIP-TO"][0] - 1].value) and not isEmptyCell(row[report_cols["Ship To Customer Number"] - 1].value) and not isEmptyCell(master_row_data[master_cols["SHIP-TO"][0] - 1].value):
                                noDiscrepancy = False
                                output_text = f"SHIP-TO diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Ship To Customer Number {row[report_cols['Ship To Customer Number'] - 1].value} vs OCCC SHIP-TO {master_row_data[master_cols['SHIP-TO'][0] - 1].value}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "SHIP-TO", row[report_cols["Ship To Customer Number"] - 1].value, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "SHIP-TO", row[report_cols["Ship To Customer Number"] - 1].value, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["SHIP-TO"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["SHIP-TO"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "SHIP-TO", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if str(report_afs_cat) != str(master_row_data[master_cols["AFS Cat"][0] - 1].value):
                                noDiscrepancy = False
                                output_text = f"AFS Cat diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Inventory Segment Code {report_afs_cat} vs OCCC AFS Cat {master_row_data[master_cols['AFS Cat'][0] - 1].value}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "AFS Cat", report_afs_cat, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "AFS Cat", report_afs_cat, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["AFS Cat"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["AFS Cat"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "AFS Cat", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if row[report_cols["VAS name"] - 1].value != master_row_data[master_cols["VAS name"][0] - 1].value:
                                noDiscrepancy = False
                                output_text = f"VAS name diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM VAS name {row[report_cols['VAS name'] - 1].value} vs OCCC VAS name {master_row_data[master_cols['VAS name'][0] - 1].value}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "VAS name", row[report_cols["VAS name"] - 1].value, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "VAS name", row[report_cols["VAS name"] - 1].value, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["VAS name"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["VAS name"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "VAS name", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if row[report_cols["Hanger size"] - 1].value != master_row_data[master_cols["Hanger size"][0] - 1].value:
                                noDiscrepancy = False
                                output_text = f"Hanger size diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Hanger size {row[report_cols['Hanger size'] - 1].value} vs OCCC Hanger size {master_row_data[master_cols['Hanger size'][0] - 1].value}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "Hanger size", row[report_cols["Hanger size"] - 1].value, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "Hanger size", row[report_cols["Hanger size"] - 1].value, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["Hanger size"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["Hanger size"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "Hanger size", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if report_ratio_qty != master_ratio_qty:
                                noDiscrepancy = False
                                output_text = f"Ratio Qty diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Ratio quantity {report_ratio_qty} vs OCCC Ratio Qty {master_ratio_qty}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "Ratio Qty", report_ratio_qty, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "Ratio Qty", report_ratio_qty, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["Ratio Qty"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["Ratio Qty"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "Ratio Qty", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if row[report_cols["Customer PO"] - 1].value != master_row_data[master_cols["Customer PO (Deichmann Group only)"] - 1].value and any(group in str(master_row_data[master_cols["VAS name"][0] - 1].value).lower() for group in ("deichmann", "dechmann")):
                                noDiscrepancy = False
                                output_text = f"Customer PO (Deichmann Group only) diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM Customer PO {row[report_cols['Customer PO'] - 1].value} vs OCCC Customer PO (Deichmann Group only) {master_row_data[master_cols['Customer PO (Deichmann Group only)'] - 1].value}\n\n"
                                append_date_changes(master_row, master_cols, report_cols, row, "Customer PO", row[report_cols["Customer PO"] - 1].value, output_text)
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "Customer PO", row[report_cols["Customer PO"] - 1].value, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["Customer PO"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["Customer PO"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "Customer PO", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            columns_to_check = [
                                "Gross Price/FOB currency code",
                                "Surcharge Min Mat Main Body currency code",
                                "Surcharge Min Material Trim currency code",
                                "Surcharge Misc currency code",
                                "Surcharge VAS currency code"
                            ]

                            if any(str(row[report_cols[col] - 1].value).strip() not in ("USD", "None") for col in columns_to_check):
                                for col in columns_to_check:
                                    if str(row[report_cols[col] - 1].value).strip() not in ("USD", "None"):
                                        noDiscrepancy = False
                                        output_text = f"Currency diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nCurrency is {row[report_cols[col] - 1].value}\n\n"
                                        append_date_changes(master_row, master_cols, report_cols, row, "Price", row[report_cols[col] - 1].value, output_text)
                            else:
                                matching_rows["Currency"].append(i)
                                for col in columns_to_check:
                                    if str(row[report_cols[col] - 1].value).strip() not in ("USD", "None"):
                                        keep_date_changes(master_row, master_cols, report_cols, row, "Price", row[report_cols[col] - 1].value, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if (report_po_fob != master_po_fob or (report_po_fob != master_po_fob_ex_1 and master_po_fob_ex_1 != 0.00) or (report_po_fob != master_po_fob_ex_2 and master_po_fob_ex_2 != 0.00)):
                                sizesE = []
                                sizesE2 = []

                                extended_size_value_1 = str(master_row[master_cols["Extended Sizes"] - 3].value).strip().replace('-', '')
                                extended_size_value_1 = "T" if "tall" in extended_size_value_1.lower() else extended_size_value_1
                                extended_size_value_1 = re.sub(r'([A-Za-z0-9]+)[\.\(\&\+].*', r'\1', extended_size_value_1)
                                if not isEmptyCell(extended_size_value_1):
                                    matching_size_1 = next((size for size in sizes if extended_size_value_1 in size), None)
                                    if not isEmptyCell(matching_size_1):
                                        index_1 = sizes.index(matching_size_1)
                                        sizesE = sizes[index_1:]

                                extended_size_value_2 = str(master_row[master_cols["Extended Sizes"] - 1].value).strip().replace('-', '')
                                extended_size_value_2 = "T" if "tall" in extended_size_value_2.lower() else extended_size_value_2
                                extended_size_value_2 = re.sub(r'([A-Za-z0-9]+)[\.\(\&\+].*', r'\1', extended_size_value_2)
                                if not isEmptyCell(extended_size_value_2):
                                    matching_size_2 = next((size for size in sizes if extended_size_value_2 in size), None)
                                    if not isEmptyCell(matching_size_2):
                                        index_2 = sizes.index(matching_size_2)
                                        sizesE2 = sizes[index_2:]

                                if (not str(row[report_cols['Size Description'] - 1].value).strip().replace('-', '') in sizesE and not str(row[report_cols['Size Description'] - 1].value).strip().replace('-', '') in sizesE2 and report_po_fob != master_po_fob) or (extended_size_value_1 != 0.00 and str(row[report_cols['Size Description'] - 1].value).strip().replace('-', '') in sizesE and report_po_fob != master_po_fob_ex_1) or (extended_size_value_2 != 0.00 and str(row[report_cols['Size Description'] - 1].value).strip().replace('-', '') in sizesE2 and report_po_fob != master_po_fob_ex_2):
                                    noDiscrepancy = False
                                    if "/" in str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip():
                                        sizesInMaster = str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip().split('/')
                                        for sizeInMaster in sizesInMaster:
                                            if str(row[report_cols['Size Description'] - 1].value).strip().replace('-', '') == sizeInMaster.strip().split(" ")[0].replace('-', ''):
                                                sizesInMaster.remove(sizeInMaster)
                                                break
                                        master_row[master_cols["DPOM - Incorrect FOB"] - 1].value = '/'.join(sizesInMaster)
                                    elif str(row[report_cols['Size Description'] - 1].value).strip().replace('-', '') == str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip().split(" ")[0].replace('-', '') if 'CORRECT' not in str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip() and not is_number(str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip()) and not isNumberAfterDash(str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip()) and not hasComma(str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip()) else str(row[report_cols['Size Description'] - 1].value).strip().replace('-', ''):
                                        master_row[master_cols["DPOM - Incorrect FOB"] - 1].value = f""

                                    output_text = f"FOB diff found for PO {report_po_num} Line {report_po_line} Size {row[report_cols['Size Description'] - 1].value}:\nPPM PO FOB {report_po_fob} vs OCCC FOB {master_po_fob}"
                                    if safe_float(master_po_fob_ex_1) > 0:
                                        output_text += f" and FOB EXT. {master_po_fob_ex_1}"
                                    if safe_float(master_po_fob_ex_2) > 0:
                                        output_text += f" and FOB EXT. EXT. {master_po_fob_ex_2}"
                                    output_text += "\n\n"
                                    append_date_changes(master_row, master_cols, report_cols, row, "Price", report_po_fob, output_text)

                                    dpom_value = str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip() if not isEmptyCell(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value) else ""
                                    size_description_value = str(row[report_cols['Size Description'] - 1].value).strip() if not isEmptyCell(row[report_cols['Size Description'] - 1].value) else ""
                                    po_fob_value = str(safe_float(report_po_fob)).strip()
                                    new_dpom_value = (f"{size_description_value} {po_fob_value}").strip()

                                    if not isEmptyCell(dpom_value):
                                        dpom_values = str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip().split('/')
                                        dpom_values_cleaned = [
                                            value.strip() for value in dpom_values
                                            if 'CORRECT' not in value and not is_number(value) and not isNumberAfterDash(value) and not hasComma(value)
                                        ]

                                        cleaned_dpom_value = ' / '.join(dpom_values_cleaned)

                                        if new_dpom_value not in dpom_values_cleaned:
                                            cleaned_dpom_value = f"{cleaned_dpom_value} / {new_dpom_value}".strip(" / ")
                                    else:
                                        cleaned_dpom_value = new_dpom_value

                                    master_row[master_cols["DPOM - Incorrect FOB"] - 1].value = cleaned_dpom_value if cleaned_dpom_value else ""
                                else:
                                    if "/" in str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip():
                                        sizesInMaster = str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip().split('/')
                                        for sizeInMaster in sizesInMaster:
                                            if str(row[report_cols['Size Description'] - 1].value).strip().replace('-', '') == sizeInMaster.strip().split(" ")[0].replace('-', ''):
                                                sizesInMaster.remove(sizeInMaster)
                                                break
                                        master_row[master_cols["DPOM - Incorrect FOB"] - 1].value = '/'.join(sizesInMaster)
                                    elif str(row[report_cols['Size Description'] - 1].value).strip().replace('-', '') == str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip().split(" ")[0].replace('-', '') if 'CORRECT' not in str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip() and not is_number(str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip()) and not isNumberAfterDash(str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip()) and not hasComma(str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip()) else str(row[report_cols['Size Description'] - 1].value).strip().replace('-', ''):
                                        master_row[master_cols["DPOM - Incorrect FOB"] - 1].value = f""

                                    dpom_value = str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip() if not isEmptyCell(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value) else ""
                                    if not isEmptyCell(dpom_value):
                                        dpom_values = str(master_row[master_cols["DPOM - Incorrect FOB"] - 1].value).strip().split('/')
                                        dpom_values_cleaned = [
                                            value.strip() for value in dpom_values
                                            if 'CORRECT' not in value and not is_number(value) and not isNumberAfterDash(value) and not hasComma(value)
                                        ]

                                        cleaned_dpom_value = ' / '.join(dpom_values_cleaned)
                                    else:
                                        cleaned_dpom_value = f"CORRECT"
                                        matching_rows["Price"].append(i)
                                        if report_path == report_files[-1] and i == matching_rows["Price"][-1] and i == matching_rows["Currency"][-1]:
                                            empty_date_changes(master_row, master_cols, report_cols, row, "Price", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                                    master_row[master_cols["DPOM - Incorrect FOB"] - 1].value = cleaned_dpom_value
                            else:
                                keep_date_changes(master_row, master_cols, report_cols, row, "Price", report_po_fob, f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")
                                matching_rows["Price"].append(i)
                                if report_path == report_files[-1] and i == matching_rows["Price"][-1] and i == matching_rows["Currency"][-1]:
                                    empty_date_changes(master_row, master_cols, report_cols, row, "Price", f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}")

                            if noDiscrepancy:
                                master_row[master_cols["Latest CM Change Date"] - 1].value = ori_cm_date[f"{report_po_num}{report_po_line}{master_row_data[master_cols['JOB NO'] - 1].value}"]

                    elif key not in master_dict[master_current]:
                        new_key = (report_po_style, report_po_num, report_po_line, row[report_cols["GAC"] - 3].value, row[report_cols["DPOM Line Item Status"] - 1].value, row[report_cols["Doc Type"] - 2].value, row[report_cols["Document Date"] - 1].value, row[report_cols["Change Date"] - 1].value)
                        if new_key not in newPO:
                            newPO[new_key] = row

                    if row_index == report_sheet.max_row:
                        report_date_bfr = str(row[report_cols["Change Date"] - 1].value.strftime('%m/%d')).strip()

            log("Saving file... please wait")
            master_wb.save(f"{master_path}_UPDATED.xlsx")
            log(f"Saved new file at {master_path}_UPDATED.xlsx")
            ok += 1
        except Exception as master_exc:
            fail += 1
            log(f"ERROR processing {master_path}: {master_exc}")

    new_wb = Workbook()
    new_ws = new_wb.active
    new_ws.title = "New PO Data"

    headers = ["Report PO Style", "Report PO Num", "Report PO Line", "GAC", "DPOM Line Item Status", "Doc Type", "Document Date", "PPM Report Date"]
    new_ws.append(headers)

    short_date_style = NamedStyle(name="short_date_style", number_format="YYYY/MM/DD")
    if "short_date_style" not in new_wb.named_styles:
        new_wb.add_named_style(short_date_style)

    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    header_font = Font(bold=True)
    header_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for col_num, col_name in enumerate(headers, 1):
        cell = new_ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = header_border

    items_to_remove = []
    for new_key in newPO.keys():
        new_first_three_keys = new_key[:3]
        for existing_key in existingPO.keys():
            existing_first_three_keys = existing_key[:3]
            if new_first_three_keys == existing_first_three_keys:
                items_to_remove.append(new_key)
                break

    for key in items_to_remove:
        if key in newPO:
            del newPO[key]

    for row_idx, (key, row) in enumerate(newPO.items(), start=2):
        new_ws.append([
            key[0],
            key[1],
            key[2],
            key[3],
            key[4],
            key[5],
            key[6],
            key[7],
        ])

        new_ws.cell(row=row_idx, column=headers.index("GAC") + 1).style = short_date_style
        new_ws.cell(row=row_idx, column=headers.index("Document Date") + 1).style = short_date_style
        new_ws.cell(row=row_idx, column=headers.index("PPM Report Date") + 1).style = short_date_style

        for col_num in range(1, len(headers) + 1):
            cell = new_ws.cell(row=row_idx, column=col_num)
            cell.border = header_border

    for col in new_ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        adjusted_width = (max_length + 2)
        new_ws.column_dimensions[col_letter].width = adjusted_width

    new_po_path = os.path.join(work_path or os.getcwd(), "New_PO_List.xlsx")
    new_wb.save(new_po_path)
    log(f"Saved new PO list file at {new_po_path}")
    report(f"Compared a total of {total_data} data in a span of {(time.time() - start_time):.2f} seconds")
    log(f"Processed {total_data} row(s) in {(time.time() - start_time):.2f} seconds")
    log("Finished, you may exit now")

    return new_po_path, ok, fail
