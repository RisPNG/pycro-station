#!/usr/bin/env python3
"""
PDF to Excel converter using pdfplumber for direct text extraction.

Optimized for text-based PDFs (not scanned images).
Falls back to marker-pdf for complex layouts if needed.
"""
from __future__ import annotations

import re
import threading
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Tuple, Optional

import pandas as pd
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QSizePolicy,
    QComboBox,
)
from qfluentwidgets import PrimaryPushButton, MessageBox


def _emit(log_emit: Callable[[str], None] | None, text: str):
    """Log helper that tolerates missing/failed callbacks."""
    if callable(log_emit):
        try:
            log_emit(text)
            return
        except Exception:
            pass
    print(text)


def ensure_unique_path(path: Path) -> Path:
    """Return a non-existing path by appending (n) before the extension if needed."""
    if not path.exists():
        return path
    stem = path.stem
    ext = path.suffix
    parent = path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem} ({counter}){ext}"
        if not candidate.exists():
            return candidate
        counter += 1


def check_dependencies() -> Tuple[bool, str]:
    """Check if required dependencies are installed."""
    missing = []

    try:
        import pdfplumber
    except ImportError:
        missing.append("pdfplumber")

    try:
        import pandas
    except ImportError:
        missing.append("pandas")

    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")

    if missing:
        msg = "Missing dependencies:\n" + "\n".join(f"  - {m}" for m in missing)
        msg += "\n\nInstall with:\n  pip install pdfplumber pandas openpyxl"
        return False, msg

    return True, ""


def clean_cell(value: str | None) -> str:
    """Clean a cell value: normalize whitespace, strip edges."""
    if value is None:
        return ""
    text = str(value)
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def parse_amount(value: str) -> Optional[float]:
    """
    Parse an amount string like '24,172.90 USD' or '(1,806.74 USD)' into a float.
    Returns None if unparseable.
    """
    if not value:
        return None

    text = clean_cell(value)

    # Check for negative amounts in parentheses: (1,806.74 USD)
    is_negative = text.startswith('(') and text.endswith(')')
    if is_negative:
        text = text[1:-1].strip()

    # Remove currency codes (USD, EUR, MYR, etc.)
    text = re.sub(r'\s*(USD|EUR|GBP|MYR|SGD|JPY|CNY)\s*', '', text, flags=re.IGNORECASE)

    # Remove thousand separators (commas) but keep decimal point
    text = text.replace(',', '')

    try:
        amount = float(text)
        return -amount if is_negative else amount
    except ValueError:
        return None


def is_reference_number(value: str) -> bool:
    """Check if a value looks like a reference number (e.g., 25V35375)."""
    pattern = r'^\d{2}[A-Z]\d{5}$'
    return bool(re.match(pattern, clean_cell(value)))


def is_contract_number(value: str) -> bool:
    """Check if a value looks like a contract number (e.g., 2500669921)."""
    text = clean_cell(value)
    return bool(re.match(r'^\d{10}$', text))


def is_amount_cell(value: str) -> bool:
    """Check if a cell contains an amount value."""
    return parse_amount(value) is not None


def classify_table_row(row: List[str]) -> str:
    """
    Classify a table row to determine what type of data it contains.
    Returns: 'reference', 'transaction', 'fee', 'header', or 'unknown'
    """
    if not row or all(not clean_cell(c) for c in row):
        return 'empty'

    row_text = ' '.join(clean_cell(c) for c in row).lower()

    # Header detection
    header_keywords = ['reference number', 'amount type', 'contract', 'invoice number',
                       'fee type', 'settlement amount', 'principal amount', 'transaction fees']
    if any(kw in row_text for kw in header_keywords):
        return 'header'

    # Check first cell for pattern
    first_cell = clean_cell(row[0]) if row else ''

    if is_reference_number(first_cell):
        return 'reference'

    if is_contract_number(first_cell):
        return 'transaction'

    if 'fee' in row_text.lower():
        return 'fee'

    return 'unknown'


def extract_tables_pdfplumber(pdf_path: Path, log_emit=None) -> List[List[List[str]]]:
    """
    Extract all tables from a PDF using pdfplumber.
    Returns a list of tables, where each table is a list of rows.
    """
    import pdfplumber

    all_tables = []

    with pdfplumber.open(pdf_path) as pdf:
        _emit(log_emit, f"PDF has {len(pdf.pages)} page(s)")

        for page_num, page in enumerate(pdf.pages, 1):
            _emit(log_emit, f"Extracting tables from page {page_num}...")

            # Try to extract tables with explicit settings
            tables = page.extract_tables({
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "snap_tolerance": 5,
                "join_tolerance": 5,
            })

            if tables:
                for table in tables:
                    # Clean each cell in the table
                    cleaned_table = []
                    for row in table:
                        if row:
                            cleaned_row = [clean_cell(c) for c in row]
                            if any(cleaned_row):  # Skip completely empty rows
                                cleaned_table.append(cleaned_row)

                    if cleaned_table:
                        all_tables.append(cleaned_table)
                        _emit(log_emit, f"  Found table with {len(cleaned_table)} rows")
            else:
                _emit(log_emit, f"  No tables found on page {page_num}")

    return all_tables


def extract_structured_data_from_text(pdf_path: Path, log_emit=None) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Extract structured data directly from PDF text using regex patterns.
    This is more reliable than table extraction for certain PDF layouts.

    Returns (reference_df, transaction_df, fees_df)
    """
    import pdfplumber

    all_text = []
    with pdfplumber.open(pdf_path) as pdf:
        _emit(log_emit, f"PDF has {len(pdf.pages)} page(s)")
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text.append(text)

    full_text = "\n".join(all_text)
    lines = full_text.split('\n')

    reference_rows = []
    transaction_rows = []
    fee_rows = []

    # Pattern for reference number rows: 25V35375 24,172.90 USD 24,172.90 USD
    ref_pattern = re.compile(r'^(\d{2}[A-Z]\d{5})\s+([\d,]+\.\d{2})\s+USD\s+([\d,]+\.\d{2})\s+USD')

    # Pattern for transaction rows: 2500669921 25V31613 Nike 2500669921 2,706.68 USD Fee Schedule
    trans_pattern = re.compile(r'^(\d{10})\s+(\d{2}[A-Z]\d{5})\s+(\w+)\s+(\d{10})\s+([\d,]+\.\d{2})\s+USD')

    # Pattern for fee rows: Invoice Fee (1,806.74 USD)
    fee_pattern = re.compile(r'^([\w\s]+Fee)\s+\(([\d,]+\.\d{2})\s+USD\)')

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Try reference pattern
        ref_match = ref_pattern.match(line)
        if ref_match:
            ref_num = ref_match.group(1)
            amount = parse_amount(ref_match.group(2))
            reference_rows.append({
                'Reference Number': ref_num,
                'Amount': amount
            })
            continue

        # Try transaction pattern
        trans_match = trans_pattern.match(line)
        if trans_match:
            transaction_rows.append({
                'Contract': trans_match.group(1),
                'Invoice Number': trans_match.group(2),
                'Payer': trans_match.group(3),
                'Purchase Order(s)': trans_match.group(4),
                'Principal Amount': parse_amount(trans_match.group(5))
            })
            continue

        # Try fee pattern
        fee_match = fee_pattern.match(line)
        if fee_match:
            fee_rows.append({
                'Fee Type': fee_match.group(1).strip(),
                'Amount': -parse_amount(fee_match.group(2))  # Fees are negative
            })

    _emit(log_emit, f"Text parsing found: {len(reference_rows)} references, "
                    f"{len(transaction_rows)} transactions, {len(fee_rows)} fees")

    ref_df = pd.DataFrame(reference_rows) if reference_rows else pd.DataFrame()
    trans_df = pd.DataFrame(transaction_rows) if transaction_rows else pd.DataFrame()
    fees_df = pd.DataFrame(fee_rows) if fee_rows else pd.DataFrame()

    return ref_df, trans_df, fees_df


def extract_text_pdfplumber(pdf_path: Path, log_emit=None) -> str:
    """Extract all text from a PDF using pdfplumber."""
    import pdfplumber

    all_text = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if text:
                all_text.append(text)

    return "\n\n".join(all_text)


def parse_text_to_tables(text: str, log_emit=None) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Parse extracted text into structured tables using regex.
    Returns (reference_df, transaction_df)
    """
    lines = text.strip().split('\n')

    reference_rows = []
    transaction_rows = []

    # Patterns for different row types
    ref_pattern = re.compile(r'(\d{2}[A-Z]\d{5})\s+([\d,]+\.\d{2})\s+USD\s+([\d,]+\.\d{2})\s+USD')
    trans_pattern = re.compile(r'(\d{10})\s+(\d{2}[A-Z]\d{5})\s+(\w+)\s+(\d{10})\s+([\d,]+\.\d{2})\s+USD')

    for line in lines:
        line = line.strip()

        # Try reference pattern
        ref_match = ref_pattern.search(line)
        if ref_match:
            ref_num = ref_match.group(1)
            amount = parse_amount(ref_match.group(2))
            reference_rows.append({
                'Reference Number': ref_num,
                'Amount': amount
            })
            continue

        # Try transaction pattern
        trans_match = trans_pattern.search(line)
        if trans_match:
            transaction_rows.append({
                'Contract': trans_match.group(1),
                'Invoice Number': trans_match.group(2),
                'Payer': trans_match.group(3),
                'Purchase Order(s)': trans_match.group(4),
                'Principal Amount': parse_amount(trans_match.group(5))
            })

    ref_df = pd.DataFrame(reference_rows) if reference_rows else pd.DataFrame()
    trans_df = pd.DataFrame(transaction_rows) if transaction_rows else pd.DataFrame()

    return ref_df, trans_df


def process_extracted_tables(
    tables: List[List[List[str]]],
    log_emit=None
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Process extracted tables into structured DataFrames.
    Returns (reference_df, transaction_df, fees_df)
    """
    reference_rows = []
    transaction_rows = []
    fee_rows = []

    for table_idx, table in enumerate(tables):
        _emit(log_emit, f"Processing table {table_idx + 1} ({len(table)} rows)...")

        current_type = None
        header_row = None

        for row in table:
            row_type = classify_table_row(row)

            if row_type == 'header':
                # Detect what kind of header this is
                row_text = ' '.join(clean_cell(c) for c in row).lower()
                if 'contract' in row_text or 'invoice number' in row_text:
                    current_type = 'transaction'
                    header_row = [clean_cell(c) for c in row]
                elif 'reference number' in row_text:
                    current_type = 'reference'
                    header_row = [clean_cell(c) for c in row]
                elif 'fee type' in row_text:
                    current_type = 'fee'
                    header_row = [clean_cell(c) for c in row]
                continue

            if row_type == 'empty':
                continue

            # Process data rows based on detected patterns
            if row_type == 'reference' or (current_type == 'reference' and is_amount_cell(row[1] if len(row) > 1 else '')):
                # Reference number row: RefNum | Amount | Amount(Inv) | Rate | Settlement
                if len(row) >= 2:
                    ref_num = clean_cell(row[0])
                    amount = parse_amount(row[1]) if len(row) > 1 else None

                    if ref_num and amount is not None:
                        reference_rows.append({
                            'Reference Number': ref_num,
                            'Amount': amount
                        })

            elif row_type == 'transaction' or current_type == 'transaction':
                # Transaction row: Contract | Invoice | Payer | PO | Amount | Fees
                if len(row) >= 5:
                    contract = clean_cell(row[0])
                    invoice = clean_cell(row[1]) if len(row) > 1 else ''
                    payer = clean_cell(row[2]) if len(row) > 2 else ''
                    po = clean_cell(row[3]) if len(row) > 3 else ''
                    amount = parse_amount(row[4]) if len(row) > 4 else None

                    if contract and (is_contract_number(contract) or invoice):
                        transaction_rows.append({
                            'Contract': contract,
                            'Invoice Number': invoice,
                            'Payer': payer,
                            'Purchase Order(s)': po,
                            'Principal Amount': amount
                        })

            elif row_type == 'fee' or current_type == 'fee':
                # Fee row
                if len(row) >= 2:
                    fee_type = clean_cell(row[0])
                    amount = parse_amount(row[1]) if len(row) > 1 else None

                    if fee_type and 'fee' in fee_type.lower():
                        fee_rows.append({
                            'Fee Type': fee_type,
                            'Amount': amount
                        })

    ref_df = pd.DataFrame(reference_rows) if reference_rows else pd.DataFrame()
    trans_df = pd.DataFrame(transaction_rows) if transaction_rows else pd.DataFrame()
    fees_df = pd.DataFrame(fee_rows) if fee_rows else pd.DataFrame()

    _emit(log_emit, f"Extracted: {len(reference_rows)} reference rows, "
                    f"{len(transaction_rows)} transaction rows, {len(fee_rows)} fee rows")

    return ref_df, trans_df, fees_df


def fallback_marker_extraction(pdf_path: Path, log_emit=None) -> List[List[List[str]]]:
    """
    Fallback to marker-pdf for complex PDFs.
    Uses CPU mode which is suitable for systems without discrete GPUs.
    """
    import os
    os.environ["TORCH_DEVICE"] = "cpu"

    _emit(log_emit, "Falling back to marker-pdf (CPU mode)...")
    _emit(log_emit, "This may take several minutes...")

    try:
        from marker.converters.pdf import PdfConverter
        from marker.models import create_model_dict
        from marker.output import text_from_rendered

        _emit(log_emit, "Loading marker-pdf models...")
        converter = PdfConverter(artifact_dict=create_model_dict())

        _emit(log_emit, "Converting PDF...")
        rendered = converter(str(pdf_path))

        # Get markdown text
        text, _, _ = text_from_rendered(rendered)

        # Parse markdown tables
        tables = parse_markdown_tables(text)

        if tables:
            _emit(log_emit, f"Marker extracted {len(tables)} table(s)")
        else:
            _emit(log_emit, "Marker found no tables, extracting as text lines")
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            if lines:
                tables = [[[line] for line in lines]]

        return tables

    except ImportError as e:
        _emit(log_emit, f"marker-pdf not available: {e}")
        return []
    except Exception as e:
        _emit(log_emit, f"marker-pdf error: {e}")
        return []


def parse_markdown_tables(text: str) -> List[List[List[str]]]:
    """Parse markdown tables from text."""
    lines = text.strip().split("\n")
    tables = []
    current_table = []
    in_table = False

    for line in lines:
        line_stripped = line.strip()

        if line_stripped.startswith("|") and line_stripped.endswith("|"):
            cells = [clean_cell(c) for c in line_stripped.split("|")[1:-1]]

            # Skip separator rows
            if all(set(c) <= {"-", ":", " "} for c in cells):
                continue

            current_table.append(cells)
            in_table = True
        elif in_table:
            if current_table:
                tables.append(current_table)
            current_table = []
            in_table = False

    if current_table:
        tables.append(current_table)

    return tables


def convert_pdf_to_excel(
    pdf_path: Path,
    use_marker_fallback: bool = True,
    log_emit: Callable[[str], None] | None = None,
) -> Path:
    """
    Convert a PDF file to Excel.

    Primary method: Text-based extraction with regex (fastest, most reliable)
    Secondary: pdfplumber table extraction
    Fallback: marker-pdf on CPU (for complex layouts)
    """
    _emit(log_emit, f"Processing: {pdf_path.name}")

    ref_df = pd.DataFrame()
    trans_df = pd.DataFrame()
    fees_df = pd.DataFrame()

    # Primary method: Direct text extraction with regex patterns
    _emit(log_emit, "Attempting text-based extraction...")
    try:
        ref_df, trans_df, fees_df = extract_structured_data_from_text(pdf_path, log_emit)
    except Exception as e:
        _emit(log_emit, f"Text extraction error: {e}")

    # If text extraction didn't get good results, try table extraction
    if ref_df.empty and trans_df.empty:
        _emit(log_emit, "Text extraction incomplete, trying table extraction...")
        tables = extract_tables_pdfplumber(pdf_path, log_emit)
        if tables:
            ref_df, trans_df, fees_df = process_extracted_tables(tables, log_emit)

    # If still no results and marker fallback enabled, try marker-pdf
    if ref_df.empty and trans_df.empty and use_marker_fallback:
        _emit(log_emit, "pdfplumber extraction failed, trying marker-pdf fallback...")
        tables = fallback_marker_extraction(pdf_path, log_emit)
        if tables:
            ref_df, trans_df, fees_df = process_extracted_tables(tables, log_emit)

    # Create output
    if ref_df.empty and trans_df.empty:
        raise ValueError("No tables could be extracted from the PDF")

    # Create output path
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"{pdf_path.stem}__converted__{ts}.xlsx"
    out_path = ensure_unique_path(pdf_path.with_name(out_name))

    _emit(log_emit, f"Writing to {out_path.name}...")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        sheet_count = 0

        if not ref_df.empty:
            ref_df.to_excel(writer, sheet_name="Reference Numbers", index=False)
            sheet_count += 1
            _emit(log_emit, f"  Sheet 'Reference Numbers': {len(ref_df)} rows")

        if not trans_df.empty:
            trans_df.to_excel(writer, sheet_name="Transaction Details", index=False)
            sheet_count += 1
            _emit(log_emit, f"  Sheet 'Transaction Details': {len(trans_df)} rows")

        if not fees_df.empty:
            fees_df.to_excel(writer, sheet_name="Fees", index=False)
            sheet_count += 1
            _emit(log_emit, f"  Sheet 'Fees': {len(fees_df)} rows")

    _emit(log_emit, f"Created Excel with {sheet_count} sheet(s)")
    return out_path


def process_files(
    files: List[str],
    use_marker_fallback: bool = True,
    log_emit: Callable[[str], None] | None = None,
) -> Tuple[int, int, List[Path]]:
    """Process multiple PDF files."""
    ok = 0
    fail = 0
    outputs: List[Path] = []

    for raw in files:
        path = Path(raw)
        _emit(log_emit, f"\n[START] {path.name}")
        try:
            out = convert_pdf_to_excel(path, use_marker_fallback, log_emit)
            outputs.append(out)
            ok += 1
            _emit(log_emit, f"[DONE] {path.name}")
        except Exception as e:
            fail += 1
            _emit(log_emit, f"[FAIL] {path.name}: {e}")

    return ok, fail, outputs


class MainWidget(QWidget):
    log_message = Signal(str)
    processing_done = Signal(int, int, list)

    def __init__(self):
        super().__init__()
        self.setObjectName("pdf_to_excel_plumber_widget")
        self._deps_ok = True
        self._deps_msg = ""
        self._check_deps()
        self._build_ui()
        self._connect_signals()

    def _check_deps(self):
        self._deps_ok, self._deps_msg = check_dependencies()

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
        self._load_long_description()

        # Controls
        self.select_btn = PrimaryPushButton("Select PDF File(s)", self)
        self.run_btn = PrimaryPushButton("Convert to Excel", self)

        # Marker fallback checkbox
        self.marker_check = QCheckBox("Enable marker-pdf fallback (slower, for complex PDFs)", self)
        self.marker_check.setChecked(False)
        self.marker_check.setStyleSheet("color: #dcdcdc;")

        # Status label for dependencies
        self.status_label = QLabel("", self)
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("color: #ffcc00; padding: 4px;")

        if not self._deps_ok:
            self.status_label.setText(f"⚠️ {self._deps_msg}")
            self.run_btn.setEnabled(False)

        # Selected files / log
        self.files_label = QLabel("Selected PDF files", self)
        self.files_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.files_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.logs_label = QLabel("Process logs", self)
        self.logs_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.logs_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.files_box = QTextEdit(self)
        self.files_box.setReadOnly(True)
        self.files_box.setPlaceholderText("Selected PDF files will appear here")
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

        # Layouts
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        main_layout.addWidget(self.desc_label)
        main_layout.addWidget(self.status_label)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_row.addWidget(self.select_btn)
        btn_row.addStretch(1)
        main_layout.addLayout(btn_row)

        main_layout.addWidget(self.marker_check)

        run_row = QHBoxLayout()
        run_row.addStretch(1)
        run_row.addWidget(self.run_btn)
        run_row.addStretch(1)
        main_layout.addLayout(run_row)

        labels_row = QHBoxLayout()
        labels_row.addWidget(self.files_label, 1)
        labels_row.addWidget(self.logs_label, 1)
        main_layout.addLayout(labels_row)

        boxes_row = QHBoxLayout()
        boxes_row.addWidget(self.files_box, 1)
        boxes_row.addWidget(self.log_box, 1)
        main_layout.addLayout(boxes_row)

    def _connect_signals(self):
        self.select_btn.clicked.connect(self.select_files)
        self.run_btn.clicked.connect(self.run_process)
        self.log_message.connect(self.append_log)
        self.processing_done.connect(self.on_processing_done)

    def _load_long_description(self):
        md_path = Path(__file__).with_name("description.md")
        if not md_path.exists():
            return
        try:
            lines = md_path.read_text(encoding="utf-8").splitlines()
        except Exception:
            return
        body: List[str] = []
        for ln in lines:
            if ln.startswith(">"):
                continue
            body.append(ln)
        text = "\n".join(body).strip()
        if text:
            self.desc_label.setText(text)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PDF files",
            "",
            "PDF Files (*.pdf);;All Files (*)",
        )
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
        if not self._deps_ok:
            MessageBox("Error", "Missing dependencies. Please install required packages.", self).exec()
            return

        files = self._selected_files()
        if not files:
            MessageBox("Warning", "No files selected.", self).exec()
            return

        use_marker = self.marker_check.isChecked()

        self.log_box.clear()
        self.log_message.emit(f"Starting conversion for {len(files)} file(s)...")
        if use_marker:
            self.log_message.emit("marker-pdf fallback enabled (may be slow on CPU)")

        self.run_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.marker_check.setEnabled(False)

        def worker():
            try:
                ok, fail, outputs = process_files(files, use_marker, log_emit=self.log_message.emit)
            except Exception as e:
                ok, fail, outputs = 0, len(files), []
                self.log_message.emit(f"Fatal error: {e}")
            self.processing_done.emit(ok, fail, outputs)

        threading.Thread(target=worker, daemon=True).start()

    def append_log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def on_processing_done(self, ok: int, fail: int, outputs: list):
        if outputs:
            self.log_message.emit("\nGenerated files:")
            for p in outputs:
                self.log_message.emit(f" - {p}")
        self.log_message.emit(f"\nCompleted: {ok} succeeded, {fail} failed.")
        self.run_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.marker_check.setEnabled(True)


def get_widget():
    return MainWidget()
