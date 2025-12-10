import os
import re
import threading
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import MessageBox, PrimaryPushButton
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

try:
    import xlrd  # for legacy .xls
except Exception:
    xlrd = None

# ---------- UI: Hook into Pycro Station ----------

EXCEL_FILTER = "Excel Files (*.xls *.xlsx *.xlsm *.xltx *.xltm)"


class MainWidget(QWidget):
    log_message = Signal(str)
    processing_done = Signal(str, int, int)  # out_path, ok, fail

    def __init__(self):
        super().__init__()
        self.setObjectName("finance_variance_report_widget")
        self._build_ui()
        self._connect_signals()

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
        self.desc_label.hide()

        self.select_btn = PrimaryPushButton("Select Management Accounts", self)
        self.run_btn = PrimaryPushButton("Generate Variance Report", self)

        self.files_box = QTextEdit(self)
        self.files_box.setReadOnly(True)
        self.files_box.setPlaceholderText("Choose 1 file (current month) or 2 files (current + comparison).")
        self.files_box.setStyleSheet(
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

        main_layout.addWidget(self.desc_label)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_row.addWidget(self.select_btn)
        btn_row.addSpacing(8)
        btn_row.addWidget(self.run_btn)
        btn_row.addStretch(1)
        main_layout.addLayout(btn_row, 0)

        label_row = QHBoxLayout()
        files_lbl = QLabel("Selected files", self)
        files_lbl.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")
        logs_lbl = QLabel("Process logs", self)
        logs_lbl.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")
        label_row.addWidget(files_lbl, 1)
        label_row.addWidget(logs_lbl, 1)
        main_layout.addLayout(label_row, 0)

        content_row = QHBoxLayout()
        content_row.addWidget(self.files_box, 1)
        content_row.addWidget(self.log_box, 1)
        main_layout.addLayout(content_row, 1)

    def _connect_signals(self):
        self.select_btn.clicked.connect(self.select_files)
        self.run_btn.clicked.connect(self.run_process)
        self.log_message.connect(self.append_log)
        self.processing_done.connect(self.on_processing_done)

    def set_long_description(self, text: str):
        clean = (text or "").strip()
        if clean:
            self.desc_label.setText(clean)
            self.desc_label.show()
        else:
            self.desc_label.hide()
            self.desc_label.clear()

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Management Account files", filter=EXCEL_FILTER)
        if not files:
            self.files_box.clear()
            return
        if len(files) > 2:
            self.log_message.emit("INFO: Using the first two files selected (newer, then older).")
        self.files_box.setPlainText("\n".join(files[:2]))

    def _selected_files(self) -> List[str]:
        text = self.files_box.toPlainText().strip()
        if not text:
            return []
        return [line for line in text.split("\n") if line.strip()][:2]

    def _set_busy(self, busy: bool):
        self.run_btn.setEnabled(not busy)
        self.select_btn.setEnabled(not busy)

    def run_process(self):
        files = self._selected_files()
        if len(files) < 1:
            MessageBox("Need at least one file", "Select the current month Management Account file (and optionally a comparison file).", self).exec()
            return


        self.log_box.clear()
        self.log_message.emit("Starting variance report generation...")
        self._set_busy(True)

        def worker(selected: List[str]):
            out_path, ok, fail = "", 0, 0
            try:
                out_path, ok, fail = process_files(selected, self.log_message.emit)
            except Exception as e:
                self.log_message.emit(f"ERROR: {e}")
            self.processing_done.emit(out_path, ok, fail)

        threading.Thread(target=worker, args=(list(files),), daemon=True).start()

    def append_log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def on_processing_done(self, out_path: str, ok: int, fail: int):
        if out_path:
            self.log_message.emit(f"Output workbook saved to: {out_path}")
        self.log_message.emit(f"Completed: {ok} success, {fail} failed.")
        self._set_busy(False)
        if out_path:
            MessageBox("Done", f"Saved report to:\n{out_path}", self).exec()


def get_widget():
    return MainWidget()

# ---------- Helper: Month tag parsing ----------

_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
}

_MONTH_TAG_RE = re.compile(r"(?i)\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[’']\s?(\d{2})\b")

def extract_month_tag(s: str) -> Optional[str]:
    """Return normalized like "May'25" if found in string."""
    m = _MONTH_TAG_RE.search(s)
    if not m:
        return None
    mon, yr2 = m.groups()
    return f"{mon.capitalize()}'{yr2}"

def parse_month_tag(tag: str) -> Optional[Tuple[int, int]]:
    """Return (YYYY, MM) from "May'25" -> (2025, 5)."""
    m = _MONTH_TAG_RE.match(tag)
    if not m:
        return None
    mon, yr2 = m.groups()
    mm = _MONTHS[mon.lower()]
    yy = 2000 + int(yr2)
    return (yy, mm)

def newer_first(a: str, b: str) -> Tuple[str, str]:
    """Return (newer, older) by comparing month tags."""
    pa, pb = parse_month_tag(a), parse_month_tag(b)
    if pa and pb:
        return (a, b) if pa > pb else (b, a)
    # fallback to filename order if parsing fails
    return (a, b)

# ---------- Helper: read Excel (.xls via xlrd, .xlsx via openpyxl) ----------

def _open_xls(path: str, log) -> List[Tuple[str, List[List[Any]]]]:
    """Return list of (sheet_name, rows) where rows is list of cell values."""
    if xlrd is None:
        raise RuntimeError("Reading .xls requires 'xlrd'. Please install it or convert to .xlsx.")
    book = xlrd.open_workbook(path)
    out = []
    for s in book.sheets():
        rows = []
        for r in range(s.nrows):
            rows.append([s.cell_value(r, c) for c in range(s.ncols)])
        out.append((s.name, rows))
    log(f"Read .xls with xlrd: {os.path.basename(path)} ({len(out)} sheets)")
    return out

def _open_xlsx(path: str, log) -> List[Tuple[str, List[List[Any]]]]:
    """Return list of (sheet_name, rows) using openpyxl (values only)."""
    wb = load_workbook(path, data_only=True, read_only=True)
    out = []
    for name in wb.sheetnames:
        ws = wb[name]
        block = []
        # pull a reasonable rectangle (up to current dimensions)
        max_r = ws.max_row or 0
        max_c = ws.max_column or 0
        for r in range(1, max_r + 1):
            row = []
            for c in range(1, max_c + 1):
                row.append(ws.cell(r, c).value)
            block.append(row)
        out.append((name, block))
    log(f"Read .xlsx with openpyxl: {os.path.basename(path)} ({len(out)} sheets)")
    return out

def read_excel_generic(path: str, log) -> List[Tuple[str, List[List[Any]]]]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        return _open_xls(path, log)
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return _open_xlsx(path, log)
    raise RuntimeError(f"Unsupported Excel type: {ext}")

# ---------- Config: labels & synonyms we want to pull ----------

CANONICAL_LABELS: Dict[str, List[str]] = {
    "Revenue - Local": [
        "revenue - local", "revenue local", "sales - local", "sales local", "local revenue"
    ],
    "Revenue - Vietnam": [
        "revenue - vietnam", "revenue vietnam", "sales - vietnam", "sales vietnam", "vietnam revenue"
    ],
    "(-) Material Costs": [
        "material cost", "materials", "raw material", "cost of material"
    ],
    "(-) Embroidery": ["embroidery"],
    "(-) H. Transfer": ["heat transfer", "h. transfer", "heat xfer", "heat x-fer"],
    "(-) Printing": ["printing"],
    "(-) Pad Print": ["pad print", "pad-print"],
    "(-) Laser Cut/ Sub Cutting": ["laser cut", "sub cutting", "laser cutting", "sub-cutting"],
    "(-) VAT on printing costs": ["vat on printing", "vat printing", "printing vat"],
    "(-) Direct  Wages - Local": ["direct wages local", "direct labour local", "direct labor local"],
    "(-) Direct  Wages - Vietnam": ["direct wages vietnam", "direct labour vietnam", "direct labor vietnam"],
    "(-) Additional / reversal of provision for sub-con fees (Vtec Invoices)": [
        "provision sub-con", "vtec invoices", "subcon provision", "sub con provision"
    ],
    "(-) Production Overhead (Jawi & Vietnam)": [
        "production overhead", "production o/h", "factory overhead"
    ],
    "(-) FOB Financial Charges": ["fob financial", "fob bank charges", "fob charges"],
    "(-) Production Overhead-Provision for COVID Fund": ["covid fund", "covid provision"],
    "(-) Production Overhead ": ["production overhead", "factory overhead", "overhead"],
    "(-) Purchase Related Cost": ["purchase related cost", "purchase costs"],
    "(-) Sales Related Cost - General": ["sales related cost", "selling expenses", "sales expense"],
    "(-) Sales Related Cost - Outward air freight": ["outward air freight", "air freight", "freight (outward)"],
    "(-) Finance Cost": ["finance cost", "interest expense", "bank charges"],
    "(-) Administrative Expenses - Bonus": ["bonus"],
    "(-) Administrative Expenses - Performance Incentive": ["performance incentive", "kpi incentive"],
    "(-) Administrative Expenses - Charity & donation": ["charity", "donation", "charity & donation"],
    "(-) Administrative Expenses - Fixed asset writen off": ["fixed asset written off", "fa written off", "impairment write off"],
    "(-) Administrative Expenses - Vietnam office rental": ["vietnam office rental", "office rental vietnam"],
    "Other Income / (Expenses)": ["other income", "other (income)/expenses", "other income / (expenses)"],
    "Gain / (Loss) on Forex - Realised": [
        "gain/(loss) on forex - realized",
        "gain/(loss) on forex - realised",
        "gain on forex realised",
        "loss on forex realised",
        "realised fx",
        "realized fx",
    ],
    "Gain / (Loss) on Forex - Unrealised": [
        "gain/(loss) on forex - unrealized",
        "gain/(loss) on forex - unrealised",
        "gain on forex unrealised",
        "loss on forex unrealised",
        "unrealised fx",
        "unrealized fx",
    ],

    "(-) Taxation": ["taxation", "income tax", "tax expense"],
    "Exchange Rate": ["exchange rate", "fx rate", "usd/rm", "rm/usd", "usdrm", "rmusd"]
}

# Which ones we expect as RM vs USD:
WANTS_RM = {
    # Revenues: we want both RM & USD if possible
    "Revenue - Local",
    "Revenue - Vietnam",
    # Most detail lines are RM inputs on the template
    "(-) Material Costs",
    "(-) Embroidery",
    "(-) H. Transfer",
    "(-) Printing",
    "(-) Pad Print",
    "(-) Laser Cut/ Sub Cutting",
    "(-) VAT on printing costs",
    "(-) Direct  Wages - Local",
    "(-) Direct  Wages - Vietnam",
    "(-) Additional / reversal of provision for sub-con fees (Vtec Invoices)",
    "(-) Production Overhead (Jawi & Vietnam)",
    "(-) FOB Financial Charges",
    "(-) Production Overhead-Provision for COVID Fund",
    "(-) Production Overhead ",
    "(-) Purchase Related Cost",
    "(-) Sales Related Cost - General",
    "(-) Sales Related Cost - Outward air freight",
    "(-) Finance Cost",
    "(-) Administrative Expenses - Bonus",
    "(-) Administrative Expenses - Performance Incentive",
    "(-) Administrative Expenses - Charity & donation",
    "(-) Administrative Expenses - Fixed asset writen off",
    "(-) Administrative Expenses - Vietnam office rental",
    "Other Income / (Expenses)",
    "Gain / (Loss) on Forex - Realised",
    "Gain / (Loss) on Forex - Unrealised",
    "(-) Taxation",
}

WANTS_USD_FOR_REVENUE = True  # We'll try to get USD revenue, else compute USD = RM / FX if FX found

def _row_rightmost_number(row: List[Any]) -> Optional[float]:
    """Pick the rightmost numeric in a row (typical 'Total' column)."""
    val = None
    for v in row:
        if isinstance(v, (int, float)) and v == v:
            val = float(v)
    return val

def _row_named_numbers(row: List[Any], header_map: Dict[str, int]) -> Dict[str, float]:
    """If we detected header columns (e.g., 'rm','usd'), pick those by name."""
    out = {}
    for name, idx in header_map.items():
        if 0 <= idx < len(row):
            v = row[idx]
            if isinstance(v, (int, float)) and v == v:
                out[name] = float(v)
    return out

def _find_template_path(candidates: List[str], log) -> Optional[str]:
    """Try to locate the template near us or alongside the data."""
    direct = "Report - Fin Result Var Analysis FY26 5.xlsx"
    # 1) CWD
    if os.path.exists(direct):
        return os.path.abspath(direct)
    # 2) alongside either data file
    for f in candidates:
        base_dir = os.path.dirname(os.path.abspath(f))
        p = os.path.join(base_dir, direct)
        if os.path.exists(p):
            return p
    log("ERROR: Could not locate 'Report - Fin Result Var Analysis FY26 5.xlsx'. Place it next to the app or your data.")
    return None

def _write_into_month_sheet(ws, metrics: Dict[str, Dict[str, float]], log) -> Tuple[int, int]:
    """
    Write values into column G of ws by matching labels in column B.
    Returns (ok_count, miss_count).
    """
    ok = 0
    miss = 0
    # Build map: label -> list of row indexes where column B == label
    label_rows: Dict[str, List[int]] = {}
    maxrow = ws.max_row or 200
    for r in range(1, maxrow + 1):
        b = ws.cell(r, 2).value  # col B
        if not isinstance(b, str):
            continue
        label_rows.setdefault(b.strip(), []).append(r)

    # Helper to find row for a label in a specific block (USD or RM)
    def find_row(label: str, want_usd: bool) -> Optional[int]:
        # Look for rows with exact label text first
        rows = label_rows.get(label)
        if not rows:
            # Try loose match if exact not found
            for k, idxs in label_rows.items():
                if fuzzy_match(label, k):
                    rows = idxs
                    break
        if not rows:
            return None
        if len(rows) == 1:
            return rows[0]
        # Disambiguate by checking the header above column G saying USD'000 or RM'000
        for r in rows:
            header = ws.cell(r-1, 7).value  # one row above, col G
            ht = _norm_text(header)
            if want_usd and "usd" in ht:
                return r
            if (not want_usd) and ("rm" in ht or "myr" in ht):
                return r
        # fallback to first
        return rows[0]

    # 1) Revenues (USD first if present)
    for label in ("Revenue - Local", "Revenue - Vietnam"):
        if "USD" in metrics and label in metrics["USD"]:
            r = find_row(label, want_usd=True)
            if r:
                ws.cell(r, 7).value = float(metrics["USD"][label])
                ok += 1
            else:
                log(f"WARNING: Could not locate USD row for '{label}' in template sheet {ws.title}")
                miss += 1
        if "RM" in metrics and label in metrics["RM"]:
            r = find_row(label, want_usd=False)
            if r:
                ws.cell(r, 7).value = float(metrics["RM"][label])
                ok += 1
            else:
                log(f"WARNING: Could not locate RM row for '{label}' in template sheet {ws.title}")
                miss += 1

    # 2) Other RM lines
    for label in WANTS_RM:
        if label in ("Revenue - Local", "Revenue - Vietnam"):
            continue
        if label in metrics.get("RM", {}):
            r = find_row(label, want_usd=False)
            if r:
                ws.cell(r, 7).value = float(metrics["RM"][label])
                ok += 1
            else:
                log(f"WARNING: Could not locate row for '{label}' in template sheet {ws.title}")
                miss += 1

    # 3) FX into the "Exchange Rate" RM block if present (optional)
    fx = metrics.get("FX", {}).get("Exchange Rate")
    if fx:
        r = find_row("Exchange Rate", want_usd=True) or find_row("Exchange Rate", want_usd=False)
        if r:
            ws.cell(r, 7).value = float(fx)
            ok += 1

    return ok, miss
def _strip_apostrophe(tag: str) -> str:
    return (tag or "").replace("'", "")

def _ensure_compare_sheet(wb, current_tag: str, older_tag: str, log) -> None:
    """
    Duplicate current_tag sheet as '<MonYY> vs <MonYY>' and replace
    '='<OlderDefault>' references with the explicit older_tag.
    """
    if current_tag not in wb.sheetnames:
        log(f"WARNING: current sheet '{current_tag}' not found; skip compare sheet.")
        return
    ws_src = wb[current_tag]
    new_name = f"{_strip_apostrophe(current_tag)} vs {_strip_apostrophe(older_tag)}"
    if new_name in wb.sheetnames:
        del wb[new_name]
    ws_new = wb.copy_worksheet(ws_src)
    ws_new.title = new_name

    # Find/Replace in all formulas: replace 'May''24' (or whatever YoY ref is)
    # with the explicit older_tag we want. We detect any reference like ='???''YY'!
    yoy_ref_re = re.compile(r"(')(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(''){0,1}(\d{2})(')!")

    replaced = 0
    for row in ws_new.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                # If it already references our desired older_tag, skip
                if older_tag in v:
                    continue
                # Replace first found month-ref with older_tag
                def _sub(m):
                    nonlocal replaced
                    replaced += 1
                    # Build like 'May''23'!
                    parts = older_tag.split("'")
                    if len(parts) == 2:
                        mon, yy = parts
                        return f"'{mon}''{yy}'!"
                    return f"'{older_tag}'!"
                new_v = yoy_ref_re.sub(_sub, v)
                cell.value = new_v

    log(f"Compare sheet created: '{new_name}' (updated {replaced} formula refs)")

def process_files(files: List[str], log) -> Tuple[str, int, int]:
    """
    Main entry point for your UI thread worker.
    Returns: (out_path, ok_count, fail_count)
    Supports 1-file (current only) or 2-file (current + comparison).
    """
    ok, fail = 0, 0
    if not files:
        log("ERROR: Select at least one Management Account file.")
        return "", ok, fail

    # --- Identify inputs & month tags (support 1 or 2 files) ---
    items: List[Tuple[str, str]] = []  # (path, month_tag)
    for f in files[:2]:
        tag = extract_month_tag(os.path.basename(f)) or ""
        if not tag:
            log(f"WARNING: Could not detect month tag from filename: {os.path.basename(f)} (expected like Jan'25, May'23)")
        items.append((f, tag))

    if len(items) == 2 and items[0][1] and items[1][1]:
        newer_tag, older_tag = newer_first(items[0][1], items[1][1])
        newer_path = items[0][0] if items[0][1] == newer_tag else items[1][0]
        older_path = items[1][0] if items[0][1] == newer_tag else items[0][0]
        single_mode = False
    else:
        # Single-file mode (or missing tags on one of them)
        newer_tag = items[0][1] or ""
        newer_path = items[0][0]
        older_tag = ""
        older_path = ""
        single_mode = True

    if single_mode:
        if not newer_tag:
            log("ERROR: Could not detect month tag from the selected file name. Please include a tag like May'25 in the file name.")
            return "", ok, fail
        log(f"Detected: {newer_tag} (single-file mode)")
    else:
        log(f"Detected months -> Newer: {newer_tag} | Older: {older_tag}")

    # --- Locate template ---
    cand = [newer_path] + ([older_path] if older_path else [])
    template_path = _find_template_path(cand, log)
    if not template_path:
        return "", ok, fail
    log(f"Using template: {os.path.basename(template_path)}")

    # --- Extract metrics ---
    log(f"Parsing management account: {os.path.basename(newer_path)}")
    metrics_newer = extract_metrics_from_file(newer_path, log)

    def _count_entries(m): return len(m.get("RM", {})) + len(m.get("USD", {})) + len(m.get("FX", {}))
    ok += _count_entries(metrics_newer)

    metrics_older = {}
    if older_path:
        log(f"Parsing management account: {os.path.basename(older_path)}")
        metrics_older = extract_metrics_from_file(older_path, log)
        ok += _count_entries(metrics_older)

    # --- Open template ---
    wb = load_workbook(template_path, data_only=False)

    # --- Find best sheet by tag ---
    def _best_sheet_name(tag: str) -> Optional[str]:
        if tag and tag in wb.sheetnames:
            return tag
        for n in wb.sheetnames:
            if tag and tag.replace(" ", "") in n.replace(" ", ""):
                return n
        return None

    # --- Write newer month ---
    newer_sheet = _best_sheet_name(newer_tag)
    if not newer_sheet:
        log(f"ERROR: Could not find sheet for '{newer_tag}' in template.")
        return "", ok, fail

    w_ok, w_miss = _write_into_month_sheet(wb[newer_sheet], metrics_newer, log)
    log(f"Wrote {w_ok} inputs to sheet '{newer_sheet}' ({w_miss} missing).")
    ok += w_ok; fail += w_miss

    # --- If older present, write it and build compare sheet ---
    if older_tag and older_path:
        older_sheet = _best_sheet_name(older_tag)
        if not older_sheet:
            log(f"ERROR: Could not find sheet for older month '{older_tag}' in template.")
            fail += 1
        else:
            w_ok, w_miss = _write_into_month_sheet(wb[older_sheet], metrics_older, log)
            log(f"Wrote {w_ok} inputs to sheet '{older_sheet}' ({w_miss} missing).")
            ok += w_ok; fail += w_miss

            _ensure_compare_sheet(wb, newer_tag, older_tag, log)

    # --- Save output (different names for 1-file vs 2-file) ---
    ts = datetime.now().strftime("%Y%m%d-%H%M")
    out_dir = os.path.dirname(template_path)

    if older_tag:
        clean_newer = _strip_apostrophe(newer_tag)
        clean_older = _strip_apostrophe(older_tag)
        out_name = f"Report - Fin Result Var Analysis ({clean_newer} vs {clean_older}) {ts}.xlsx"
    else:
        clean_newer = _strip_apostrophe(newer_tag)
        out_name = f"Report - Fin Result Var Analysis ({clean_newer}) {ts}.xlsx"

    out_path = os.path.join(out_dir, out_name)
    # Force a full recalculation when opened in Excel
    if getattr(wb, "calculation", None) is not None:
        wb.calculation.fullCalcOnLoad = True

    wb.save(out_path)
    log(f"Saved: {out_path}")
    return out_path, ok, fail

def _build_month_regex(month_tag: str):
    # month_tag like "May'25"
    m = re.match(r"(?i)^\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[’']\s?(\d{2})\s*$", month_tag or "")
    if not m:
        return None
    mon, yy2 = m.groups()
    yyyy = 2000 + int(yy2)
    # Match "May'25", "May 25", "May-25", "May 2025"
    pat = rf"(?i)\b{re.escape(mon)}\s*(?:['’\-]?\s*{re.escape(yy2)}|{yyyy})\b"
    return re.compile(pat)

def _find_month_col(rows: List[List[Any]], month_tag: Optional[str]) -> Optional[int]:
    """Scan top rows for a header cell that contains the month tag; return its column index."""
    if not month_tag:
        return None
    rx = _build_month_regex(month_tag)
    if not rx:
        return None
    scan_rows = min(25, len(rows))
    for r in range(scan_rows):
        row = rows[r]
        for c, v in enumerate(row or []):
            if isinstance(v, str) and rx.search(v):
                return c
    return None

def _build_month_regex(month_tag: str):
    # month_tag like "May'25"
    m = re.match(r"(?i)^\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[’']\s?(\d{2})\s*$", month_tag or "")
    if not m:
        return None
    mon, yy2 = m.groups()
    yyyy = 2000 + int(yy2)
    # Match "May'25", "May 25", "May-25", "May 2025", "May 2025 Actual" etc.
    pat = rf"(?i)\b{re.escape(mon)}\s*(?:['’\-]?\s*{re.escape(yy2)}|{yyyy})\b"
    return re.compile(pat)
    """
    Find the column index for the target MONTH **RM column**.
    Strategy:
      1) Find any header cell containing the month tag in the top ~25 rows.
      2) In a small window around that column, prefer a column whose header mentions 'RM' or 'MYR'.
      3) Otherwise choose the window column with the greatest count of numeric values below the header.
    """
    if not month_tag or not rows:
        return None
    rx = _build_month_regex(month_tag)
    if not rx:
        return None

    scan_rows = min(25, len(rows))
    candidates = []  # (row_idx, col_idx)
    for r in range(scan_rows):
        row = rows[r] or []
        for c, v in enumerate(row):
            if isinstance(v, str) and rx.search(v):
                candidates.append((r, c))
    if not candidates:
        log(f"WARNING: Could not find a header cell for month '{month_tag}'")
        return None

    # examine each candidate; return first good RM col
    for hr, hc in candidates:
        # window around month header col
        c_lo, c_hi = max(0, hc - 3), min((len(rows[0]) if rows and rows[0] else hc + 3), hc + 3)
        # 2a) prefer an explicit RM header near the header row (look +/- 3 rows)
        for c in range(c_lo, c_hi + 1):
            for r in range(max(0, hr - 3), min(scan_rows, hr + 4)):
                txt = rows[r][c] if c < len(rows[r]) else None
                t = _norm_text(txt)
                if t and ("rm" in t or "myr" in t):
                    log(f"Month '{month_tag}': picked RM column {c} (explicit header '{txt}') near ({r},{c})")
                    return c
        # 2b) no explicit RM header → choose the most numeric column in the window
        best_c, best_count = None, -1
        for c in range(c_lo, c_hi + 1):
            num_count = 0
            for r in range(hr + 1, min(len(rows), hr + 60)):
                v = rows[r][c] if c < len(rows[r]) else None
                if isinstance(v, (int, float)) and v == v:
                    num_count += 1
            if num_count > best_count:
                best_count, best_c = num_count, c
        if best_c is not None and best_count > 0:
            log(f"Month '{month_tag}': picked numeric-dense column {best_c} in window around {hc} (count={best_count})")
            return best_c

    log(f"WARNING: Month '{month_tag}' header found but no usable RM column nearby")
    return None

def _norm_text(x) -> str:
    """Normalize cell text for fuzzy matching."""
    s = str(x or "").lower()
    # Replace non-alphanumerics with space
    s = re.sub(r"[^a-z0-9]+", " ", s)
    # Collapse multiple spaces
    s = re.sub(r"\s+", " ", s).strip()
    return s


def fuzzy_match(label: str, target: str) -> bool:
    """Loose containment match after normalization."""
    l = _norm_text(label)
    t = _norm_text(target)
    if not l or not t:
        return False
    return l in t or t in l


def _detect_currency_headers(sheet_rows: List[List[Any]]) -> Dict[str, int]:
    """
    Detect columns named 'RM', 'RM'000, 'USD', 'USD'000, etc.
    Returns mapping like {'rm': col_index, 'usd': col_index}.
    We scan more rows because the RM / USD headers are often lower down.
    """
    header_map: Dict[str, int] = {}
    max_scan = min(50, len(sheet_rows))  # was 10
    for r in range(max_scan):
        row = sheet_rows[r] or []
        for c, v in enumerate(row):
            t = _norm_text(v)
            if not t:
                continue
            if "rm" in t or "myr" in t:
                header_map.setdefault("rm", c)
            # 'USD' or 'US$' -> 'usd' or 'us' after normalization
            if "usd" in t or t == "us":
                header_map.setdefault("usd", c)
    return header_map


def _find_month_rm_col(rows: List[List[Any]], month_tag: Optional[str], log) -> Optional[int]:
    """
    For this Management Account layout, the 'current month' RM column is simply
    the column where we see an 'RM' header. We ignore month_tag and rely on
    _detect_currency_headers instead.
    """
    if not rows:
        return None
    header_map = _detect_currency_headers(rows)
    rm_col = header_map.get("rm")
    if rm_col is not None:
        log(f"Using RM column index {rm_col} from currency headers.")
        return rm_col
    log("WARNING: Could not detect RM column from headers.")
    return None


def extract_metrics_from_file(path: str, log, month_tag: Optional[str] = None) -> Dict[str, Dict[str, float]]:
    """
    Extract current-month metrics from a Management Account file.

    Returns:
      {
        'RM':  { <canonical_label>: current-month value (full RM), ... },
        'USD': { 'Revenue - Local': x, 'Revenue - Vietnam': y },  # direct if present; fallback from RM/FX
        'FX':  { 'Exchange Rate': z }
      }
    """
    data: Dict[str, Dict[str, float]] = {"RM": {}, "USD": {}, "FX": {}}

    try:
        sheets = read_excel_generic(path, log)
    except Exception as e:
        log(f"ERROR reading {os.path.basename(path)}: {e}")
        return data

    for sname, rows in sheets:
        if not rows:
            continue

        header_map = _detect_currency_headers(rows)
        rm_col = _find_month_rm_col(rows, month_tag, log)

        for r_idx, row in enumerate(rows):
            row = row or []

            # Collect ALL text cells in the first few columns for label matching
            text_cells = [
                cell
                for cell in row[:6]
                if isinstance(cell, str) and _norm_text(cell)
            ]
            if not text_cells:
                continue

            for canon, alts in CANONICAL_LABELS.items():
                matched = False
                chosen_text: Optional[str] = None

                for txt in text_cells:
                    if any(fuzzy_match(alt, txt) for alt in (alts + [canon])):
                        matched = True
                        chosen_text = txt
                        break

                if not matched or not chosen_text:
                    continue

                nums_named = _row_named_numbers(row, header_map) if header_map else {}
                rightmost = _row_rightmost_number(row)

                # Prefer current-month value at rm_col
                mon_val: Optional[float] = None
                if rm_col is not None and rm_col < len(row):
                    mv = row[rm_col]
                    if isinstance(mv, (int, float)) and mv == mv:
                        mon_val = float(mv)

                nt = _norm_text(chosen_text)

                # ---------- FX: Exchange Rate ----------
                if canon == "Exchange Rate":
                    fx_val = mon_val

                    # Fallback: derive from RM / USD if both present in same row
                    if fx_val is None and "rm" in nums_named and "usd" in nums_named and nums_named["usd"]:
                        rm_v = nums_named["rm"]
                        usd_v = nums_named["usd"]
                        if usd_v:
                            ratio = rm_v / usd_v
                            fx_val = ratio if ratio > 1 else (usd_v / rm_v if rm_v else None)

                    # Fallback: rightmost numeric
                    if fx_val is None and rightmost is not None:
                        fx_val = rightmost

                    if fx_val:
                        data["FX"]["Exchange Rate"] = float(fx_val)
                        log(f"[{sname}] FX found: {fx_val} (row {r_idx+1})")
                    break  # done with this row

                # ---------- Revenue (Local / Vietnam) ----------
                if canon in ("Revenue - Local", "Revenue - Vietnam"):
                    # For your layout, current-month USD and RM are both in the
                    # same numeric column, but different rows. We decide which
                    # bucket (RM vs USD) based on the label text.
                    val = mon_val if mon_val is not None else nums_named.get("rm", rightmost)
                    if val is None:
                        break

                    is_usd_row = "usd" in nt
                    is_rm_row = "rm" in nt or not is_usd_row  # if no explicit 'usd', treat as RM

                    if is_usd_row:
                        data["USD"][canon] = float(val)
                    if is_rm_row:
                        data["RM"][canon] = float(val)
                    break

                # ---------- FX: Realised vs Unrealised split (Page 9) ----------
                if canon.startswith("Gain / (Loss) on Forex - Realised"):
                    # Only accept rows that actually mention 'real'
                    if "real" not in nt:
                        break
                    val = mon_val if mon_val is not None else nums_named.get("rm", rightmost)
                    if val is not None:
                        data["RM"][canon] = float(val)
                    break

                if canon.startswith("Gain / (Loss) on Forex - Unrealised"):
                    # Only accept rows that actually mention 'unreal'
                    if "unreal" not in nt:
                        break
                    val = mon_val if mon_val is not None else nums_named.get("rm", rightmost)
                    if val is not None:
                        data["RM"][canon] = float(val)
                    break

                # ---------- Taxation: store as positive cost line ----------
                if canon == "(-) Taxation":
                    val = mon_val if mon_val is not None else nums_named.get("rm", rightmost)
                    if val is not None:
                        data["RM"][canon] = abs(float(val))  # template expects positive '(-) Taxation'
                    break

                # ---------- All other RM lines ----------
                val = mon_val if mon_val is not None else nums_named.get("rm", rightmost)
                if val is not None:
                    data["RM"][canon] = float(val)
                break  # we handled this row for this canonical label

    # ---------- Post-pass: FX sanity & fallback USD from RM/FX ----------
    fx = data["FX"].get("Exchange Rate")
    if fx is not None:
        fx = float(fx)
        # sanity check – RM/USD should be roughly between 3 and 6
        if not (3.0 <= fx <= 6.0):
            data["FX"].pop("Exchange Rate", None)
            fx = None

    # If FX is valid, optionally compute USD from RM as a fallback/correction
    if fx:
        for k in ("Revenue - Local", "Revenue - Vietnam"):
            rm_val = data["RM"].get(k)
            if rm_val is None:
                continue
            usd_calc = rm_val / fx
            usd_existing = data["USD"].get(k)
            if (usd_existing is None) or (
                abs(usd_calc - usd_existing) / max(1.0, abs(usd_calc)) > 0.10
            ):
                data["USD"][k] = usd_calc

    return data
