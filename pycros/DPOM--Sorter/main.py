import os
import threading
import copy
from typing import List, Any, Optional, NamedTuple
from datetime import date, timedelta

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QTextEdit,
    QWidget,
    QSizePolicy
)
from qfluentwidgets import PrimaryPushButton, MessageBox
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Side, Border
from openpyxl.utils import get_column_letter

# --- CONSTANTS (VBA Colors -> ARGB Hex) ---
COLOR_CYAN = "FF00FFFF"
COLOR_PINK = "FFEBC4C3"
COLOR_GREEN = "FFA5E37D"
COLOR_YELLOW = "FFE6FEB8"
COLOR_D_GREEN = "FF00B050"
COLOR_BLACK = "FF000000"
COLOR_WHITE = "00000000" # Transparent

class BlindBuyItem(NamedTuple):
    sc: str
    bbj: str

# --- VIRTUAL EXCEL ENGINE (High Speed) ---
class VCell:
    __slots__ = ['value', 'fill', 'bold', 'number_format']
    def __init__(self, value=None):
        self.value = value
        self.fill = None
        self.bold = False
        self.number_format = "General"

class VRow:
    __slots__ = ['cells', 'height']
    def __init__(self, size, height=None):
        self.cells = [VCell() for _ in range(size)]
        self.height = height

class VirtualSheet:
    """Mimics an Excel Worksheet in RAM for O(1) operations."""
    def __init__(self, openpyxl_ws):
        self.rows: List[VRow] = []
        self.max_col = openpyxl_ws.max_column

        # Load all data into memory structure
        for row in openpyxl_ws.iter_rows():
            v_row = VRow(self.max_col)
            # Copy values. We assume default styles initially to be overwritten by logic.
            for i, cell in enumerate(row):
                if i < len(v_row.cells):
                    v_row.cells[i].value = cell.value
            self.rows.append(v_row)

    @property
    def max_row(self): return len(self.rows)

    def cell(self, r, c):
        # 1-based index (VBA style) -> 0-based list
        # Auto-expand rows
        while len(self.rows) < r:
            self.rows.append(VRow(self.max_col))

        row = self.rows[r-1]
        # Auto-expand columns
        while len(row.cells) < c:
            row.cells.append(VCell())
            self.max_col = max(self.max_col, c)

        return row.cells[c-1]

    def insert_rows(self, idx, amount=1):
        # Insert empty rows before 1-based index 'idx'
        new_block = [VRow(self.max_col) for _ in range(amount)]
        self.rows[idx-1:idx-1] = new_block

    def delete_rows(self, idx, amount=1):
        # Delete starting at 1-based index
        del self.rows[idx-1 : idx-1+amount]

    def insert_cols(self, idx, amount=1):
        # Insert empty cells at 1-based col index
        for row in self.rows:
            for _ in range(amount):
                row.cells.insert(idx-1, VCell())
        self.max_col += amount

    def delete_cols(self, idx, amount=1):
        for row in self.rows:
            del row.cells[idx-1 : idx-1+amount]
        self.max_col -= amount

    def row_height(self, r, h):
        if 1 <= r <= len(self.rows):
            self.rows[r-1].height = h

# --- UI CLASS ---
class MainWidget(QWidget):
    log_message = Signal(str)
    processing_done = Signal(int, int, str)

    def __init__(self):
        super().__init__()
        self.setObjectName("dpom_sorter_widget")
        self._build_ui()
        self._connect_signals()

    def _build_ui(self):
        self.desc_label = QLabel("", self)
        self.desc_label.setWordWrap(True)
        self.desc_label.setStyleSheet("color: #dcdcdc; background: transparent; padding: 6px;")

        self.select_btn2 = PrimaryPushButton("Select Size Excel", self)
        self.select_btn = PrimaryPushButton("Select Raw DPOM Excel Files", self)
        self.run_btn = PrimaryPushButton("Run", self)

        self.files_label = QLabel("Selected files", self)
        self.logs_label = QLabel("Process logs", self)

        self.files_box = QTextEdit(self)
        self.files_box.setPlaceholderText("Selected files...")
        self.files_box.setReadOnly(True)

        self.files_box2 = QTextEdit(self)
        self.files_box2.setPlaceholderText("Size File...")
        self.files_box2.setMaximumHeight(50)
        self.files_box2.setReadOnly(True)

        self.log_box = QTextEdit(self)
        self.log_box.setPlaceholderText("Logs...")
        self.log_box.setReadOnly(True)

        layout = QVBoxLayout(self)
        layout.addWidget(self.desc_label)

        r1 = QHBoxLayout()
        r1.addStretch(); r1.addWidget(self.select_btn2); r1.addStretch()
        layout.addLayout(r1)
        layout.addWidget(self.files_box2)

        r2 = QHBoxLayout()
        r2.addWidget(self.select_btn); r2.addWidget(self.run_btn)
        layout.addLayout(r2)

        r3 = QHBoxLayout()
        r3.addWidget(self.files_label); r3.addWidget(self.logs_label)
        layout.addLayout(r3)

        r4 = QHBoxLayout()
        r4.addWidget(self.files_box); r4.addWidget(self.log_box)
        layout.addLayout(r4)

    def _connect_signals(self):
        self.select_btn2.clicked.connect(self.sel_size)
        self.select_btn.clicked.connect(self.sel_files)
        self.run_btn.clicked.connect(self.run)
        self.log_message.connect(self.append_log)
        self.processing_done.connect(self.done)

    def sel_size(self):
        f, _ = QFileDialog.getOpenFileName(self, "Select Size File", filter="Excel (*.xlsx *.xlsm *.xls)")
        if f: self.files_box2.setText(f)

    def sel_files(self):
        fs, _ = QFileDialog.getOpenFileNames(self, "Select Files", filter="Excel (*.xlsx *.xlsm *.xls)")
        if fs: self.files_box.setText("\n".join(fs))

    def run(self):
        size = self.files_box2.toPlainText().strip()
        files = [x.strip() for x in self.files_box.toPlainText().split('\n') if x.strip()]
        if not size or not files: return
        self.run_btn.setEnabled(False)
        self.log_box.clear()
        threading.Thread(target=self.worker, args=(size, files), daemon=True).start()

    def worker(self, size_path, files):
        processor = DPOMProcessor(size_path, self.log_message.emit)
        ok, fail = 0, 0
        last = ""
        for f in files:
            try:
                processor.process(f)
                ok += 1
                last = f
            except Exception as e:
                import traceback
                self.log_message.emit(f"Error {os.path.basename(f)}: {e}")
                print(traceback.format_exc())
                fail += 1
        self.processing_done.emit(ok, fail, last)

    def append_log(self, t): self.log_box.append(t)
    def done(self, o, f, p):
        self.log_message.emit(f"Done. Success: {o}, Fail: {f}")
        self.run_btn.setEnabled(True)

def get_widget(): return MainWidget()

# --- UTILS ---
def normalize_size(v):
    if v is None: return ""
    s = str(v).strip().upper().replace("-", "")
    rep = {"XXXXXL":"6XL","XXXXL":"5XL","XXXL":"3XL","XXL":"2XL","XXXXXS":"6XS","XXXXS":"5XS","XXXS":"4XS","XXS":"2XS"}
    for k,v in rep.items(): s = s.replace(k,v)
    return s

# --- LOGIC ENGINE ---
class DPOMProcessor:
    def __init__(self, size_path, log):
        self.log = log
        self.blind_buys = []
        self.master_sizes = []
        self._load_refs(size_path)
        self._reset_ptrs()

    def _reset_ptrs(self):
        self.c_Vendor = 1; self.c_SeasonCode = 2; self.c_SeasonYear = 3; self.c_Style = 4
        self.c_PO = 7; self.c_TradingCoPO = 8; self.c_POLine = 9; self.c_OGAC = 11
        self.c_DocTypeCode = 15; self.c_Transportation = 17; self.c_ShipToCusNo = 18
        self.c_ShipToCusName = 19; self.c_Country = 20; self.c_InventorySegmentCode = 21
        self.c_SubCategoryDesc = 23; self.c_SizeDesc = 24; self.c_SizeQty = 25
        self.c_TotalSizeQty = 26; self.c_FOB = 27
        self.max_col = 0; self.size_col_start = 0

    def _load_refs(self, path):
        self.log(f"Loading Refs: {os.path.basename(path)}")
        wb = load_workbook(path, data_only=True)
        if "Blind Buy" in wb.sheetnames:
            for r in wb["Blind Buy"].iter_rows(min_row=2, values_only=True):
                if r[0]: self.blind_buys.append((str(r[0]).strip().upper(), str(r[1] or "").strip()))
        if "Size" in wb.sheetnames:
            for r in wb["Size"].iter_rows(min_row=2, max_col=1, values_only=True):
                if r[0]: self.master_sizes.append(str(r[0]))
        wb.close()

    def _is_bb(self, s):
        s = s.strip().upper()
        return any(x[0] == s for x in self.blind_buys)

    def _get_bb(self, s):
        s = s.strip().upper()
        for x in self.blind_buys:
            if x[0] == s: return x[1]
        return ""

    def process(self, path):
        fname = os.path.basename(path)
        self.log(f"Processing {fname}...")

        # 1. READ
        wb = load_workbook(path)
        ws_raw = wb.worksheets[0]
        if ws_raw.cell(1, 27).value == "PROCESSED":
            self.log("Skipping (Processed)")
            return

        # 2. VIRTUALIZE (Load to RAM)
        self.log("Loading to memory...")
        v_sheet = VirtualSheet(ws_raw)
        self._reset_ptrs()

        # 3. EXECUTE MODULE 1
        self.sort_record(v_sheet)
        self.rearrange_size_col(v_sheet)
        self.first_pass(v_sheet)
        self.second_pass(v_sheet)
        self.add_new_column(v_sheet)
        self.remove_column(v_sheet)
        self.revert_to_obs(v_sheet)

        # 4. EXECUTE MODULE 2
        self.separate_data_groups(v_sheet)

        # 5. WRITE BACK (AutoFit included here)
        self.log("Writing & AutoFitting...")
        ws_res = wb.create_sheet(ws_raw.title + "-copy")
        wb.move_sheet(ws_res, offset=-1)

        # Pre-define styles for speed
        fills = {
            COLOR_GREEN: PatternFill("solid", fgColor=COLOR_GREEN),
            COLOR_D_GREEN: PatternFill("solid", fgColor=COLOR_D_GREEN),
            COLOR_PINK: PatternFill("solid", fgColor=COLOR_PINK),
            COLOR_YELLOW: PatternFill("solid", fgColor=COLOR_YELLOW),
            COLOR_CYAN: PatternFill("solid", fgColor=COLOR_CYAN),
            COLOR_BLACK: PatternFill("solid", fgColor=COLOR_BLACK)
        }
        bold_font = Font(bold=True)

        # Track max width for AutoFit
        col_widths = {}

        for r_idx, v_row in enumerate(v_sheet.rows, 1):
            if v_row.height:
                # The "Minute" Thing: Explicit Row Height 5
                ws_res.row_dimensions[r_idx].height = v_row.height

            for c_idx, v_cell in enumerate(v_row.cells, 1):
                val = v_cell.value
                if val is not None:
                    c = ws_res.cell(r_idx, c_idx, val)

                    # Styles
                    if v_cell.fill and v_cell.fill in fills: c.fill = fills[v_cell.fill]
                    if v_cell.bold: c.font = bold_font
                    c.number_format = v_cell.number_format

                    # AutoFit Logic
                    s_len = len(str(val))
                    current_w = col_widths.get(c_idx, 0)
                    if s_len > current_w: col_widths[c_idx] = s_len

        # Apply AutoFit
        for col_idx, width in col_widths.items():
            let = get_column_letter(col_idx)
            # Cap width to reasonable max, min 10
            adj_width = min(max(width * 1.2, 8), 50)
            ws_res.column_dimensions[let].width = adj_width

        ws_raw.cell(1, 27).value = "PROCESSED"
        wb.save(path)
        self.log("Done.")

    def sort_record(self, ws: VirtualSheet):
        self.log("Sorting...")
        max_col = ws.max_col
        max_row = ws.max_row
        if max_row < 2: return

        data = []
        for r in range(2, max_row + 1):
            vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
            style = str(vals[self.c_Style-1] or "").strip()
            season = str(vals[self.c_SeasonCode-1] or "").strip().upper()
            yr = str(vals[self.c_SeasonYear-1] or "")
            ogac = str(vals[self.c_OGAC-1] or "").strip()
            po = str(vals[self.c_PO-1] or "")
            poline = str(vals[self.c_POLine-1] or "")
            country = str(vals[self.c_Country-1] or "")
            ship = str(vals[self.c_ShipToCusNo-1] or "")

            s_head = style[:6] if len(style)>=6 else style
            s_cw = style[-3:] if len(style)>=3 else style
            rank = {'SP':1,'SU':2,'FA':3,'HO':4}.get(season, 5)
            og = ogac
            if "/" in ogac:
                p = ogac.split('/')
                if len(p)>=3: og = p[2][-4:] + p[0].zfill(2) + p[1].zfill(2)
            elif len(og)>=4: og = og[-4:]

            data.append((s_head, yr, rank, og, po, s_cw, poline, country, ship, vals))

        data.sort(key=lambda x: (x[0], x[1], x[2], x[3], x[4], x[5], x[6], x[7], x[8]))

        for i, d in enumerate(data):
            r = i + 2
            for c, v in enumerate(d[9]): ws.cell(r, c+1).value = v

    def rearrange_size_col(self, ws: VirtualSheet):
        self.log("Arranging Columns...")
        self.max_col = ws.max_col
        self.size_col_start = self.max_col + 1

        sizes = set()
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, self.c_SizeDesc).value
            if v: sizes.add(str(v).strip())

        unique = list(sizes)
        sorted_sizes = []
        for m in self.master_sizes:
            nm = normalize_size(m)
            for u in unique[:]:
                if normalize_size(u) == nm:
                    sorted_sizes.append(u); unique.remove(u); break
        sorted_sizes.extend(unique)

        for i, s in enumerate(sorted_sizes):
            ws.cell(1, self.size_col_start + i).value = s

    def first_pass(self, ws: VirtualSheet):
        self.log("First Pass (Grouping)...")
        self.max_col = ws.max_col
        ws.cell(1, self.max_col + 1).value = "Overall Result"

        size_map = {}
        for c in range(self.size_col_start, self.max_col + 1):
            h = ws.cell(1, c).value
            if h: size_map[normalize_size(h)] = c

        start_ptr = 2; r = 2
        while r <= ws.max_row:
            match = False
            if r < ws.max_row:
                k1 = self._get_keys(ws, r); k2 = self._get_keys(ws, r+1)
                if k1 == k2: match = True

            if not match:
                end = r
                ws.insert_rows(end + 1, 2)
                ws.cell(end, self.max_col + 2).value = "Total Item Qty"
                ws.cell(end+1, self.max_col + 2).value = "Trading Co Net Unit Price"
                ws.cell(end+2, self.max_col + 2).value = "Net Unit Price"

                row_sum = 0
                for i in range(start_ptr, end + 1):
                    s_desc = normalize_size(ws.cell(i, self.c_SizeDesc).value)
                    qty = ws.cell(i, self.c_SizeQty).value
                    try: qty = float(qty)
                    except: qty = 0
                    fob = ws.cell(i, self.c_FOB).value
                    tc = size_map.get(s_desc)
                    if tc:
                        ws.cell(end, tc).value = qty
                        ws.cell(end+1, tc).value = fob; ws.cell(end+2, tc).value = fob
                        row_sum += qty
                ws.cell(end, self.max_col + 1).value = row_sum
                start_ptr = end + 3; r = start_ptr - 1
            r += 1

        # Cleanup
        to_del = []
        for r in range(ws.max_row, 1, -1):
            val = ws.cell(r, self.max_col + 2).value
            if val is None or str(val) == "": to_del.append(r)

        if to_del:
            to_del.sort(reverse=True)
            curr = to_del[0]; cnt = 1
            for i in range(1, len(to_del)):
                if to_del[i] == curr - 1: curr = to_del[i]; cnt += 1
                else: ws.delete_rows(curr, cnt); curr = to_del[i]; cnt = 1
            ws.delete_rows(curr, cnt)

        ws.cell(ws.max_row + 1, 1).value = "Overall Result"
        ws.cell(ws.max_row + 1, self.max_col + 2).value = "Total Item Qty"

        helper = self.max_col + 2
        ws.insert_cols(self.c_ShipToCusNo, 1)
        for r in range(1, ws.max_row + 2):
            ws.cell(r, self.c_ShipToCusNo).value = ws.cell(r, helper + 1).value
        ws.delete_cols(helper + 1, 1)
        ws.delete_cols(self.c_SizeDesc + 1, 4)

        self.c_Ptr = 18; self.c_ShipToCusNo = 19; self.c_ShipToCusName = 20
        self.c_Country = 21; self.c_InventorySegmentCode = 22; self.c_SubCategoryDesc = 24
        self.size_col_start -= 3

    def _get_keys(self, ws, r):
        return (str(ws.cell(r, self.c_SeasonCode).value or "").strip(), str(ws.cell(r, self.c_SeasonYear).value or "").strip(),
                str(ws.cell(r, self.c_Style).value or "").strip(), str(ws.cell(r, self.c_PO).value or "").strip(),
                str(ws.cell(r, self.c_POLine).value or "").strip(), str(ws.cell(r, self.c_ShipToCusNo).value or "").strip(),
                str(ws.cell(r, self.c_ShipToCusName).value or "").strip(), str(ws.cell(r, self.c_TotalSizeQty).value or "").strip())

    def second_pass(self, ws: VirtualSheet):
        self.log("Second Pass (Style/OGAC/PO)...")
        self.max_col = ws.max_col

        # Style
        start = 2; start_ptr = 2; prev_st = self._get_st(ws, start)
        while start <= ws.max_row:
            curr_st = self._get_st(ws, start)
            if curr_st != prev_st:
                self._ins_style(ws, start, start_ptr)
                start += 4; start_ptr = start; prev_st = curr_st
            start += 1

        # OGAC
        start = 2; start_ptr = 2
        while start <= ws.max_row:
            if ws.cell(start, self.c_Ptr).value == "Total Style Qty":
                curr = start_ptr
                prev_o = str(ws.cell(curr, self.c_OGAC).value or "").strip()
                while curr < start:
                    curr_o = str(ws.cell(curr, self.c_OGAC).value or "").strip()
                    if curr_o != prev_o:
                        self._ins_sum(ws, curr, start_ptr, "Total OGAC Qty", COLOR_PINK)
                        start += 2; curr += 2; start_ptr = curr
                        prev_o = str(ws.cell(curr, self.c_OGAC).value or "").strip()
                    curr += 1
                self._ins_sum(ws, start, start_ptr, "Total OGAC Qty", COLOR_PINK)
                start += 2; start_ptr = start + 1
            start += 1

        # PO
        start = 2; start_ptr = 2
        while start <= ws.max_row:
            if ws.cell(start, self.c_Ptr).value == "Total OGAC Qty":
                curr = start_ptr
                prev_k = self._get_po_k(ws, curr)
                while curr < start:
                    curr_k = self._get_po_k(ws, curr)
                    st = str(ws.cell(curr, self.c_Style).value or "").strip().upper()
                    if curr_k != prev_k or self._is_bb(st):
                        add = self._ins_po(ws, curr, start_ptr, COLOR_YELLOW)
                        start += add; curr += add; start_ptr = curr
                        prev_k = self._get_po_k(ws, curr)
                    curr += 1
                add = self._ins_po(ws, start, start_ptr, COLOR_YELLOW)
                start += add; start_ptr = start + 1

            if start <= ws.max_row and ws.cell(start, self.c_Ptr).value is None: start+=1; start_ptr=start
            if start <= ws.max_row and ws.cell(start, self.c_Ptr).value == "Total Style Qty": start+=4; start_ptr=start
            start += 1

        # Colorway
        start = 2; start_ptr = 2; cw_c = 0; prev_cw = self._get_cw(ws, start)
        while start <= ws.max_row:
            lbl = ws.cell(start, self.c_Ptr).value
            if lbl == "Total Item Qty":
                curr_cw = self._get_cw(ws, start)
                if curr_cw == prev_cw: cw_c += 1
                else:
                    if cw_c >= 2:
                        self._ins_cw(ws, start, start_ptr)
                        start += 1
                    start_ptr = start; cw_c = 1; prev_cw = curr_cw
            elif lbl == "Total PO Qty":
                start += 4
                if start <= ws.max_row: prev_cw = self._get_cw(ws, start); start_ptr = start; cw_c = 0
                continue
            start += 1

    def _get_st(self, ws, r): return str(ws.cell(r, self.c_Style).value or "").strip().upper().split("-")[0]
    def _get_cw(self, ws, r):
        v = str(ws.cell(r, self.c_Style).value or "").strip().upper()
        return v.split("-")[1] if "-" in v else v
    def _get_po_k(self, ws, r): return (str(ws.cell(r, self.c_Country).value), str(ws.cell(r, self.c_PO).value), str(ws.cell(r, self.c_InventorySegmentCode).value), str(ws.cell(r, self.c_ShipToCusNo).value))

    def _ins_style(self, ws, at, s_ptr):
        pt = at - 1
        for r in range(s_ptr, at):
            if ws.cell(r, self.c_Ptr).value == "Total Item Qty": pt = r
        ws.insert_rows(at, 4)
        ws.cell(at, self.c_Ptr).value = "Total Style Qty"
        self._cp(ws, pt, at, self.c_Ptr); self._cp(ws, pt+1, at+1, self.max_col); self._cp(ws, pt+2, at+2, self.max_col)
        self._fmt(ws, at, self.max_col, COLOR_GREEN, True); self._fmt(ws, at+1, self.max_col, COLOR_GREEN, True)
        self._fmt(ws, at+2, self.max_col, COLOR_GREEN, True); self._fmt(ws, at+3, self.max_col, COLOR_D_GREEN, False)
        ws.row_height(at+3, 5) # MINUTE ROW HEIGHT
        self._sum(ws, at, s_ptr, at, "Total Item Qty")

    def _ins_sum(self, ws, at, s_ptr, lbl, col):
        ws.insert_rows(at, 2)
        self._cp(ws, at-1, at, self.c_Ptr)
        ws.cell(at, self.c_Ptr).value = lbl
        self._sum(ws, at, s_ptr, at, "Total Item Qty")
        self._fmt(ws, at, self.max_col, col, True); self._fmt(ws, at+1, self.max_col, col, False)
        ws.row_height(at+1, 5) # MINUTE ROW HEIGHT

    def _ins_po(self, ws, at, s_ptr, col):
        ws.insert_rows(at, 4)
        self._cp(ws, at-1, at, self.c_Ptr)
        ws.cell(at, self.c_Ptr).value = "Total PO Qty"
        pt = at - 1
        while pt >= s_ptr and ws.cell(pt, self.c_Ptr).value != "Total Item Qty": pt -= 1
        self._cp(ws, pt+1, at+1, self.max_col); self._cp(ws, pt+2, at+2, self.max_col)
        self._sum(ws, at, s_ptr, at, "Total Item Qty")
        self._fmt(ws, at, self.max_col, col, True); self._fmt(ws, at+1, self.max_col, col, True)
        self._fmt(ws, at+2, self.max_col, col, True); self._fmt(ws, at+3, self.max_col, COLOR_BLACK, False)
        ws.row_height(at+3, 5) # MINUTE ROW HEIGHT
        return 4

    def _ins_cw(self, ws, at, s_ptr):
        ws.insert_rows(at, 1)
        ws.cell(at, self.c_Ptr).value = "Total Colorway Qty"
        self._cp(ws, at-1, at, self.c_Ptr)
        self._sum(ws, at, s_ptr, at, "Total Item Qty")
        self._fmt(ws, at, self.max_col, COLOR_CYAN, True)

    def _sum(self, ws, t_r, s, e, crit):
        t1=0; t2=0; tq=0
        for c in range(self.size_col_start, self.max_col):
            q=0
            for r in range(s, e):
                if ws.cell(r, self.c_Ptr).value == crit:
                    v = ws.cell(r, c).value
                    if isinstance(v, (int, float)): q += v
            ws.cell(t_r, c).value = q or None
            m1 = ws.cell(t_r+1, c).value; m2 = ws.cell(t_r+2, c).value
            if q and isinstance(m1, (int, float)): t1 += q*m1
            if q and isinstance(m2, (int, float)): t2 += q*m2
        for r in range(s, e):
            if ws.cell(r, self.c_Ptr).value == crit:
                v = ws.cell(r, self.max_col).value
                if v: tq += v
        ws.cell(t_r, self.max_col).value = tq
        if ws.cell(t_r, self.c_Ptr).value in ["Total Style Qty", "Total PO Qty"]:
            ws.cell(t_r+1, self.max_col+1).value = t1; ws.cell(t_r+2, self.max_col+1).value = t2
            ws.cell(t_r+1, self.max_col+1).bold = True; ws.cell(t_r+2, self.max_col+1).bold = True
            ws.cell(t_r+1, self.max_col+1).fill = ws.cell(t_r, 1).fill; ws.cell(t_r+2, self.max_col+1).fill = ws.cell(t_r, 1).fill

    def _cp(self, ws, s, d, mc):
        for c in range(1, mc + 1): ws.cell(d, c).value = ws.cell(s, c).value
    def _fmt(self, ws, r, mc, f, b):
        for c in range(1, mc + 1):
            ws.cell(r, c).fill = f; ws.cell(r, c).bold = b

    def add_new_column(self, ws):
        self.log("Adding Columns...")
        ws.insert_cols(5, 5); ws.insert_cols(21, 1)
        self.max_col+=6; self.c_PO+=5; self.c_TradingCoPO+=5; self.c_POLine+=5; self.c_OGAC+=6
        self.c_DocTypeCode+=6; self.c_Transportation+=6; self.c_Ptr+=6; self.c_ShipToCusNo+=6
        self.c_ShipToCusName+=6; self.c_Country+=6; self.c_InventorySegmentCode+=6; self.c_SubCategoryDesc+=6
        ws.cell(1, 5).value = "Job Number"; ws.cell(1, 6).value = "Product Type"; ws.cell(1, 7).value = "VNFOB"
        ws.cell(1, 8).value = "Destination"; ws.cell(1, 9).value = "Blind Buy Job #"; ws.cell(1, 21).value = "Estimate BusWeekDate"
        for r in range(2, ws.max_row+1):
            if ws.cell(r, self.c_Style).value: ws.cell(r, 7).value = "N"
            if ws.cell(r, self.c_Ptr).value == "Total PO Qty":
                st = str(ws.cell(r, self.c_Style).value or "").strip().upper()
                if self._is_bb(st):
                    v = self._get_bb(st)
                    ws.cell(r, 9).value = v; ws.cell(r+1, 9).value = v; ws.cell(r+2, 9).value = v

    def remove_column(self, ws):
        self.log("Removing Cols...")
        to_del = []
        for r in range(ws.max_row, 1, -1):
            if ws.cell(r, self.c_Ptr).value == "Total Item Qty": to_del.extend([r+1, r+2])
        to_del.sort(reverse=True)
        if to_del:
            curr = to_del[0]; cnt = 1
            for i in range(1, len(to_del)):
                if to_del[i] == curr - 1: curr = to_del[i]; cnt += 1
                else: ws.delete_rows(curr, cnt); curr = to_del[i]; cnt = 1
            ws.delete_rows(curr, cnt)
        for r in range(2, ws.max_row+1):
            lbl = ws.cell(r, self.c_Ptr).value
            if lbl in ["Total Style Qty", "Total PO Qty", "Total OGAC Qty"]:
                st = str(ws.cell(r, self.c_Style).value or "").split("-")[0]
                ws.cell(r, self.c_Style).value = st
                if lbl == "Total PO Qty": ws.cell(r+1, self.c_Style).value = st; ws.cell(r+2, self.c_Style).value = st

    def revert_to_obs(self, ws):
        self.log("Revert OBS...")
        for r in range(2, ws.max_row+1):
            if not ws.cell(r, self.c_ShipToCusNo).value:
                lbl = ws.cell(r, self.c_Ptr).value
                if lbl and "Total" not in lbl:
                    ws.cell(r, self.c_ShipToCusNo).value = "#"; ws.cell(r, self.c_ShipToCusName).value = "#"
            if not ws.cell(r, self.c_TradingCoPO).value and ws.cell(r, self.c_Style).value: ws.cell(r, self.c_TradingCoPO).value = "-"
        ws.cell(1, self.c_Style).value = "Material"; ws.cell(1, self.c_ShipToCusName).value = "Customer Name"
        ws.cell(1, self.c_InventorySegmentCode).value = "AFS Category"; ws.cell(1, self.c_SubCategoryDesc).value = "Sub Category Size Value"
        ws.cell(1, self.c_DocTypeCode).value = "Buy Group"; ws.cell(1, self.c_Vendor).value = "Vendor"
        ws.cell(1, self.c_SeasonCode).value = "Planning Season"; ws.cell(1, self.c_SeasonYear).value = "Year"
        ws.cell(1, self.c_PO).value = "PO Number"; ws.cell(1, self.c_OGAC).value = "OGAC Date"; ws.cell(1, self.c_Transportation).value = "Mode"
        ws.insert_rows(1, 2); ws.cell(2, self.max_col).value = "Overall Result"; ws.cell(3, self.max_col).value = "TOTAL"
        ws.insert_cols(self.c_ShipToCusName + 1)

    def separate_data_groups(self, ws):
        self.log("Module 2...")
        # Phase 1: Separators
        r = 4
        while r <= ws.max_row:
             # VBA Module 2: Specific check for Total PO Qty clearing AD/AE
             # "X" is 24. "AD" is 30. "AE" is 31.
             if ws.cell(r, 24).value == "Total PO Qty":
                 ws.cell(r, 30).value = None
                 ws.cell(r, 31).value = None

             c_f = ws.cell(r, 1).fill; n_f = ws.cell(r+1, 1).fill if r < ws.max_row else None
             if c_f == COLOR_YELLOW and (n_f == COLOR_WHITE or n_f is None):
                 ws.insert_rows(r+1)
                 for c in range(1, ws.max_col+1): ws.cell(r+1, c).fill = COLOR_BLACK
                 ws.row_height(r+1, 5) # THE MINUTE ROW HEIGHT
                 r += 2
             else: r += 1

        # Phase 2: Split
        r = 4
        while r <= ws.max_row:
            if ws.cell(r, 1).fill == COLOR_WHITE:
                s = r
                while r <= ws.max_row:
                    if ws.cell(r, 1).fill == COLOR_BLACK: break
                    r += 1
                e = r - 1

                whites = []; yellows = []; cyan = False
                for i in range(s, e+1):
                    cf = ws.cell(i, 1).fill
                    if cf == COLOR_WHITE: whites.append(i)
                    elif cf == COLOR_YELLOW: yellows.append(i)
                    elif cf == COLOR_CYAN: cyan = True

                c_P = 26; diff = False
                for c in range(1, ws.max_col+1):
                    if str(ws.cell(3, c).value) == "Customer Name": c_P = c
                if len(whites) > 1:
                    v = ws.cell(whites[0], c_P).value
                    for w in whites[1:]:
                        if ws.cell(w, c_P).value != v: diff = True

                spec = False; c_W=30; c_Y=32; c_AB=35
                for c in range(1, ws.max_col+1):
                    h = str(ws.cell(3, c).value)
                    if h == "Sub Category Size Value": c_W = c
                    if h == "Size Quantity": c_Y = c
                for w in whites:
                    y = ws.cell(w, c_Y).value; wv = str(ws.cell(w, c_W).value or "")
                    if isinstance(y, (int, float)) and wv in ["VL", "TR"]: spec = True

                if (diff and cyan) or spec:
                    block = []
                    for w in whites:
                        block.append(('W', self._cp_row(ws, w)))
                        for y in yellows:
                            dat = self._cp_row(ws, y)
                            dat['v'][20] = ws.cell(w, 21).value; dat['v'][25] = ws.cell(w, 26).value
                            if dat['v'][43]:
                                for k in range(37, 44): dat['v'][k] = ws.cell(w, k+1).value
                            block.append(('Y', dat))
                        block.append(('S', None))
                    if block[-1][0] == 'S': block.pop()
                    ws.delete_rows(s, e-s+1); ws.insert_rows(s, len(block))
                    for i, (t, d) in enumerate(block):
                        rr = s+i
                        if t == 'S':
                            for c in range(1, ws.max_col+1): ws.cell(rr, c).fill = COLOR_BLACK
                            ws.row_height(rr, 5) # THE MINUTE ROW HEIGHT
                        else:
                            for c, v in enumerate(d['v']):
                                cobj = ws.cell(rr, c+1)
                                cobj.value = v; cobj.fill = d['f'][c]; cobj.bold = d['b'][c]; cobj.number_format = d['n'][c]
                    r = s + len(block)
                else: r += 1
            else: r += 1

    def _cp_row(self, ws, r):
        v=[]; f=[]; b=[]; n=[]
        for c in range(1, ws.max_col+1):
            o = ws.cell(r, c)
            v.append(o.value); f.append(o.fill); b.append(o.bold); n.append(o.number_format)
        return {'v':v, 'f':f, 'b':b, 'n':n}