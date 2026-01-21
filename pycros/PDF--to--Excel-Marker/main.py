#!/usr/bin/env python3
from __future__ import annotations

import os
import re
import threading
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Tuple

import pandas as pd
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
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


def detect_gpu_info() -> dict:
    """
    Detect GPU availability and capabilities.

    Returns a dict with:
        - has_cuda: bool
        - device_count: int
        - devices: list of dicts with name, vram_gb, is_discrete
        - recommended_device: str ('cuda', 'cpu')
        - recommended_vram: int (GB)
    """
    info = {
        "has_cuda": False,
        "device_count": 0,
        "devices": [],
        "recommended_device": "cpu",
        "recommended_vram": 4,
    }

    try:
        import torch

        if not torch.cuda.is_available():
            return info

        info["has_cuda"] = True
        info["device_count"] = torch.cuda.device_count()

        for i in range(info["device_count"]):
            props = torch.cuda.get_device_properties(i)
            vram_gb = props.total_memory / (1024 ** 3)

            # Heuristic: discrete GPUs typically have >= 4GB VRAM
            # Integrated GPUs (Intel UHD, AMD APU) usually have < 4GB dedicated
            is_discrete = vram_gb >= 4.0

            device_info = {
                "index": i,
                "name": props.name,
                "vram_gb": round(vram_gb, 1),
                "is_discrete": is_discrete,
            }
            info["devices"].append(device_info)

        # Find best GPU (prefer discrete with most VRAM)
        discrete_gpus = [d for d in info["devices"] if d["is_discrete"]]

        if discrete_gpus:
            best = max(discrete_gpus, key=lambda x: x["vram_gb"])
            info["recommended_device"] = "cuda"
            info["recommended_vram"] = int(best["vram_gb"])
        else:
            # Only integrated GPU(s) available - use CPU to avoid OOM
            # marker-pdf needs ~4GB VRAM minimum for comfortable operation
            info["recommended_device"] = "cpu"
            info["recommended_vram"] = 4

    except ImportError:
        pass
    except Exception:
        pass

    return info


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


def clean_cell_value(value: str) -> str:
    """Clean cell values by removing HTML tags and normalizing whitespace."""
    if not value:
        return ""
    # Replace <br> and <br/> with newlines, then collapse to spaces
    value = re.sub(r'<br\s*/?>', ' ', value, flags=re.IGNORECASE)
    # Remove any other HTML tags
    value = re.sub(r'<[^>]+>', '', value)
    # Normalize whitespace
    value = ' '.join(value.split())
    return value.strip()


def normalize_text(value: str) -> str:
    """Normalize text for fuzzy matching (case/whitespace insensitive)."""
    if not value:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def row_contains_keywords(row: List[str], keywords: List[str]) -> bool:
    """Return True if any keyword appears in any cell of the row."""
    if not row or not keywords:
        return False
    for cell in row:
        cell_norm = normalize_text(cell)
        if not cell_norm:
            continue
        for kw in keywords:
            if kw and kw in cell_norm:
                return True
    return False


def table_contains_keywords(table: List[List[str]], keywords: List[str], max_rows: int | None = 25) -> bool:
    """Return True if any keyword appears in the first `max_rows` (or all) rows."""
    if not table or not keywords:
        return False
    rows = table if max_rows is None else table[:max_rows]
    return any(row_contains_keywords(row, keywords) for row in rows)


def is_empty_row(row: List[str]) -> bool:
    return all(not normalize_text(c) for c in row)


def trim_empty_rows(table: List[List[str]]) -> List[List[str]]:
    """Trim leading/trailing empty rows."""
    if not table:
        return table
    start = 0
    end = len(table)
    while start < end and is_empty_row(table[start]):
        start += 1
    while end > start and is_empty_row(table[end - 1]):
        end -= 1
    return table[start:end]


def split_table_before_keywords(
    table: List[List[str]],
    start_keywords: List[str],
) -> Tuple[List[List[str]], List[List[str]]]:
    """
    Split a table into (before, after), where `after` begins at the first row
    containing any of `start_keywords`.
    """
    if not table:
        return [], []
    for idx, row in enumerate(table):
        if row_contains_keywords(row, start_keywords):
            return trim_empty_rows(table[:idx]), trim_empty_rows(table[idx:])
    return trim_empty_rows(table), []


def classify_table(table: List[List[str]]) -> str:
    """
    Classify tables into coarse buckets to drive sheet organization.

    This is intentionally heuristic; if it can't confidently classify, returns "other".
    """
    if table_contains_keywords(table, ["contract", "invoice number", "transaction fees", "transaction detail"]):
        return "transaction"
    if table_contains_keywords(table, ["amount type", "fee type", "settlement amount"]):
        return "amount"
    return "other"


def extract_tables_from_document(rendered, log_emit=None) -> List[List[List[str]]]:
    """
    Extract tables directly from marker-pdf's rendered document structure.

    Returns a list of tables, where each table is a list of rows,
    and each row is a list of cell values.
    """
    tables = []

    # marker-pdf's rendered document contains blocks with table data
    if hasattr(rendered, 'children'):
        for block in rendered.children:
            if hasattr(block, 'block_type') and block.block_type == 'Table':
                table_data = extract_table_block(block)
                if table_data:
                    tables.append(table_data)
            # Recursively check nested blocks
            elif hasattr(block, 'children'):
                for child in block.children:
                    if hasattr(child, 'block_type') and child.block_type == 'Table':
                        table_data = extract_table_block(child)
                        if table_data:
                            tables.append(table_data)

    return tables


def extract_table_block(block) -> List[List[str]]:
    """Extract table data from a marker-pdf table block."""
    rows = []

    if hasattr(block, 'cells') and block.cells:
        # Direct cell access if available
        max_row = max(c.row_id for c in block.cells) + 1 if block.cells else 0
        max_col = max(c.col_id for c in block.cells) + 1 if block.cells else 0

        # Initialize grid
        grid = [['' for _ in range(max_col)] for _ in range(max_row)]

        for cell in block.cells:
            text = clean_cell_value(cell.text if hasattr(cell, 'text') else str(cell))
            grid[cell.row_id][cell.col_id] = text

        rows = grid
    elif hasattr(block, 'children'):
        # Parse from children (TableRow, TableCell structure)
        for row_block in block.children:
            if hasattr(row_block, 'children'):
                row = []
                for cell_block in row_block.children:
                    text = ''
                    if hasattr(cell_block, 'text'):
                        text = clean_cell_value(cell_block.text)
                    elif hasattr(cell_block, 'children'):
                        # Concatenate text from nested elements
                        texts = []
                        for elem in cell_block.children:
                            if hasattr(elem, 'text'):
                                texts.append(elem.text)
                        text = clean_cell_value(' '.join(texts))
                    row.append(text)
                if row:
                    rows.append(row)

    return rows


def parse_markdown_tables(text: str) -> List[List[List[str]]]:
    """
    Parse markdown text and extract all tables.

    Returns a list of tables, where each table is a list of rows,
    and each row is a list of cell values.
    """
    lines = text.strip().split("\n")
    tables = []
    current_table = []
    in_table = False

    for line in lines:
        line_stripped = line.strip()

        # Detect markdown table row
        if line_stripped.startswith("|") and line_stripped.endswith("|"):
            cells = [clean_cell_value(c) for c in line_stripped.split("|")[1:-1]]

            # Skip separator rows (|---|---|)
            if all(set(c) <= {"-", ":", " "} for c in cells):
                continue

            current_table.append(cells)
            in_table = True
        elif in_table:
            # End of current table
            if current_table:
                tables.append(current_table)
            current_table = []
            in_table = False

    # Add any remaining table
    if current_table:
        tables.append(current_table)

    return tables


def looks_like_header(row: List[str]) -> bool:
    """
    Determine if a row looks like a header row.

    Header rows typically contain descriptive text without:
    - Currency amounts (USD, EUR, etc.)
    - Reference numbers (patterns like 25V35xxx)
    - Mostly numeric values
    """
    if not row:
        return False

    numeric_count = 0
    currency_pattern = re.compile(r'\d+[,.]?\d*\s*(USD|EUR|GBP|MYR|SGD)', re.IGNORECASE)
    reference_pattern = re.compile(r'^\d{2}[A-Z]\d{5}$')  # Pattern like 25V35719

    for cell in row:
        cell_clean = cell.strip()
        if not cell_clean:
            continue

        # Check for currency amounts
        if currency_pattern.search(cell_clean):
            return False

        # Check for reference number patterns
        if reference_pattern.match(cell_clean):
            return False

        # Check if mostly numeric (allowing commas and periods)
        cell_digits = re.sub(r'[,.\s]', '', cell_clean)
        if cell_digits.isdigit() and len(cell_digits) > 3:
            numeric_count += 1

    # If more than half the cells are numeric, probably not a header
    non_empty_cells = len([c for c in row if c.strip()])
    if non_empty_cells > 0 and numeric_count / non_empty_cells > 0.5:
        return False

    return True


def get_effective_column_count(table: List[List[str]]) -> int:
    """
    Get the maximum column count actually used in the table.
    """
    if not table:
        return 0
    return max(len(row) for row in table)


def normalize_table_columns(table: List[List[str]], target_cols: int) -> List[List[str]]:
    """
    Normalize all rows in a table to have the same number of columns.
    """
    normalized = []
    for row in table:
        if len(row) < target_cols:
            # Pad with empty strings
            normalized.append(row + [''] * (target_cols - len(row)))
        else:
            normalized.append(row[:target_cols])
    return normalized


def tables_are_similar_structure(table1: List[List[str]], table2: List[List[str]]) -> bool:
    """
    Check if two tables have similar enough structure to be merged.

    Considers:
    - Column count (exact match or table2 is subset)
    - Data patterns in first few rows
    """
    if not table1 or not table2:
        return False

    cols1 = get_effective_column_count(table1)
    cols2 = get_effective_column_count(table2)

    # Table2 can have same or fewer columns (it's a continuation without header columns)
    if cols2 > cols1:
        return False

    # Check if table2's first row looks like data (not headers)
    if looks_like_header(table2[0]):
        return False

    # Additional check: if column counts are very different, be more careful
    # Allow merge if cols2 is at least 50% of cols1
    if cols2 < cols1 * 0.5:
        return False

    return True


def merge_continuation_tables(tables: List[List[List[str]]]) -> List[List[List[str]]]:
    """
    Merge tables that appear to be continuations of each other.

    Tables are merged if:
    - They have similar column structure (same or fewer columns)
    - The second table's first row doesn't look like headers
    """
    if len(tables) <= 1:
        return tables

    merged = []
    i = 0

    while i < len(tables):
        current = tables[i]
        if not current:
            i += 1
            continue

        current_cols = get_effective_column_count(current)

        # Try to merge following tables
        j = i + 1
        while j < len(tables):
            next_table = tables[j]
            if not next_table:
                j += 1
                continue

            # Check if tables can be merged
            if tables_are_similar_structure(current, next_table):
                # Normalize next table to match current's column count
                next_normalized = normalize_table_columns(next_table, current_cols)
                # Also normalize current table rows for consistency
                current = normalize_table_columns(current, current_cols)
                # Merge: append all rows from next table
                current = current + next_normalized
                j += 1
            else:
                # Different structure, stop merging
                break

        merged.append(current)
        i = j

    return merged


def create_dataframe_smart(table: List[List[str]]) -> pd.DataFrame:
    """
    Create a DataFrame with smart header detection.

    Only uses the first row as headers if it actually looks like headers.
    """
    if not table:
        return pd.DataFrame()

    if len(table) == 1:
        # Single row - just return as data
        return pd.DataFrame(table)

    first_row = table[0]

    if looks_like_header(first_row):
        # First row looks like headers
        return pd.DataFrame(table[1:], columns=first_row)
    else:
        # First row is data, generate generic column names
        num_cols = len(first_row)
        columns = [f"Column_{i+1}" for i in range(num_cols)]
        return pd.DataFrame(table, columns=columns)


def convert_pdf_to_excel(
    pdf_path: Path,
    device: str = "cpu",
    vram_gb: int = 4,
    log_emit: Callable[[str], None] | None = None,
) -> Path:
    """
    Convert a PDF file to Excel using marker-pdf.

    Args:
        pdf_path: Path to the PDF file
        device: 'cuda' or 'cpu'
        vram_gb: Available VRAM in GB (used for INFERENCE_RAM)
        log_emit: Optional logging callback

    Returns:
        Path to the generated Excel file
    """
    # Set environment variables before importing marker
    os.environ["TORCH_DEVICE"] = device
    if device == "cuda":
        os.environ["INFERENCE_RAM"] = str(vram_gb)

    _emit(log_emit, f"Using device: {device}" + (f" with {vram_gb}GB VRAM" if device == "cuda" else ""))

    # Import marker after setting environment
    from marker.converters.pdf import PdfConverter
    from marker.models import create_model_dict
    from marker.output import text_from_rendered

    _emit(log_emit, "Loading marker-pdf models...")
    converter = PdfConverter(artifact_dict=create_model_dict())

    _emit(log_emit, f"Converting {pdf_path.name}...")
    rendered = converter(str(pdf_path))

    # Try to extract tables from document structure first
    _emit(log_emit, "Extracting tables from document...")
    tables = extract_tables_from_document(rendered, log_emit)

    # If no tables found in structure, fall back to markdown parsing
    if not tables:
        _emit(log_emit, "No structured tables found, parsing markdown output...")
        text, _, _ = text_from_rendered(rendered)
        tables = parse_markdown_tables(text)

    # If still no tables, extract as line-by-line content
    if not tables:
        _emit(log_emit, "No tables found, extracting text content...")
        text, _, _ = text_from_rendered(rendered)
        lines = [clean_cell_value(line) for line in text.strip().split("\n") if line.strip()]
        if lines:
            tables = [[[line] for line in lines]]

    if not tables:
        raise ValueError("No content could be extracted from the PDF")

    _emit(log_emit, f"Found {len(tables)} table(s) before merging")

    # Merge continuation tables (tables split across pages)
    tables = merge_continuation_tables(tables)
    _emit(log_emit, f"After merging continuations: {len(tables)} table(s)")

    # Create output path
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"{pdf_path.stem}__converted__{ts}.xlsx"
    out_path = ensure_unique_path(pdf_path.with_name(out_name))

    # Write tables to Excel (each table as a separate sheet or merged)
    _emit(log_emit, f"Writing to {out_path.name}...")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # --- Post-process for "main data + tables" layouts ---
        #
        # Some PDFs (like payment confirmations) contain a small key/value "main data"
        # section above one or more tables. marker-pdf may extract those together as
        # a single table. When we detect that pattern, we split the main section into
        # its own sheet and keep the remaining tables on subsequent sheets.
        main_data: List[List[str]] = []
        amount_tables: List[List[List[str]]] = []
        transaction_tables: List[List[List[str]]] = []
        other_tables: List[List[List[str]]] = []

        remaining_tables: List[List[List[str]]] = []
        for idx, table in enumerate(tables):
            if not table:
                continue

            table = trim_empty_rows(table)
            if not table:
                continue

            # Split the first extracted table if it contains a recognizable table header
            # after some key/value rows.
            if idx == 0:
                before, after = split_table_before_keywords(
                    table,
                    start_keywords=[
                        "amount type",
                        "fee type",
                        "contract",
                        "transaction detail",
                    ],
                )
                if before and after:
                    main_data = before
                    table = after

            remaining_tables.append(table)

        for table in remaining_tables:
            kind = classify_table(table)
            if kind == "amount":
                amount_tables.append(table)
            elif kind == "transaction":
                transaction_tables.append(table)
            else:
                other_tables.append(table)

        wrote_any = False

        # 1) Main data sheet (no header row)
        if main_data:
            main_cols = get_effective_column_count(main_data)
            main_norm = normalize_table_columns(main_data, main_cols)
            df_main = pd.DataFrame(main_norm)
            df_main.to_excel(writer, sheet_name="Main_Data", index=False, header=False)
            wrote_any = True

        # 2) Amount-related tables (append into one sheet to match "first table page")
        if amount_tables:
            startrow = 0
            for table in amount_tables:
                df = create_dataframe_smart(table)
                df.to_excel(writer, sheet_name="Table_1", index=False, startrow=startrow)
                startrow += df.shape[0] + 2  # header + data + blank row
            wrote_any = True

        # 3) Transaction detail table(s)
        if transaction_tables:
            startrow = 0
            for table in transaction_tables:
                df = create_dataframe_smart(table)
                df.to_excel(writer, sheet_name="Table_2", index=False, startrow=startrow)
                startrow += df.shape[0] + 2
            wrote_any = True

        # 4) Fallback / extra tables
        if other_tables or not wrote_any:
            # If we didn't detect a "main data" split, keep the previous behavior:
            # write each extracted table to its own sheet in order.
            if not wrote_any:
                for i, table in enumerate(tables):
                    if not table:
                        continue
                    sheet_name = "Sheet1" if len(tables) == 1 else f"Table_{i + 1}"
                    df = create_dataframe_smart(table)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # Keep any unclassified tables so data isn't silently dropped.
                for i, table in enumerate(other_tables, start=1):
                    sheet_name = f"Other_{i}"
                    df = create_dataframe_smart(table)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

    return out_path


def process_files(
    files: List[str],
    device: str,
    vram_gb: int,
    log_emit: Callable[[str], None] | None = None,
) -> Tuple[int, int, List[Path]]:
    """Process all PDF files, returning (ok_count, fail_count, output_paths)."""
    ok = 0
    fail = 0
    outputs: List[Path] = []

    for raw in files:
        path = Path(raw)
        _emit(log_emit, f"[START] {path.name}")
        try:
            out = convert_pdf_to_excel(path, device, vram_gb, log_emit)
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
        self.setObjectName("pdf_to_excel_widget")
        self.gpu_info = detect_gpu_info()
        self._build_ui()
        self._connect_signals()

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
        self.select_btn = PrimaryPushButton("Select PDF Files", self)
        self.run_btn = PrimaryPushButton("Convert to Excel", self)

        # Device selection
        self.device_label = QLabel("Processing Device", self)
        self.device_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.device_combo = QComboBox(self)
        self.device_combo.addItem("CPU (Safe, slower)", "cpu")

        if self.gpu_info["has_cuda"]:
            for dev in self.gpu_info["devices"]:
                gpu_type = "Discrete" if dev["is_discrete"] else "Integrated"
                label = f"GPU {dev['index']}: {dev['name']} ({dev['vram_gb']}GB - {gpu_type})"
                self.device_combo.addItem(label, f"cuda:{dev['index']}")

            # Auto-select recommended device
            if self.gpu_info["recommended_device"] == "cuda":
                # Find first discrete GPU
                for i, dev in enumerate(self.gpu_info["devices"]):
                    if dev["is_discrete"]:
                        self.device_combo.setCurrentIndex(i + 1)  # +1 because CPU is first
                        break

        # GPU info label
        self.gpu_info_label = QLabel("", self)
        self.gpu_info_label.setWordWrap(True)
        self.gpu_info_label.setStyleSheet("color: #888888; background: transparent; padding-left: 2px; font-size: 11px;")
        self._update_gpu_info_label()

        # Selected files / log
        self.files_label = QLabel("Selected files", self)
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

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_row.addWidget(self.select_btn)
        btn_row.addStretch(1)
        main_layout.addLayout(btn_row)

        grid = QGridLayout()
        grid.setColumnStretch(1, 1)
        grid.addWidget(self.device_label, 0, 0, Qt.AlignLeft)
        grid.addWidget(self.device_combo, 0, 1, Qt.AlignLeft)
        grid.addWidget(self.gpu_info_label, 1, 0, 1, 2, Qt.AlignLeft)
        main_layout.addLayout(grid)

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
        self.device_combo.currentIndexChanged.connect(self._update_gpu_info_label)
        self.log_message.connect(self.append_log)
        self.processing_done.connect(self.on_processing_done)

    def _update_gpu_info_label(self):
        device_data = self.device_combo.currentData()

        if device_data == "cpu":
            self.gpu_info_label.setText(
                "CPU mode: Slower but works on all systems. "
                "Recommended if you have an integrated GPU or limited VRAM."
            )
        elif device_data and device_data.startswith("cuda:"):
            idx = int(device_data.split(":")[1])
            dev = self.gpu_info["devices"][idx]

            if dev["is_discrete"]:
                self.gpu_info_label.setText(
                    f"Discrete GPU detected with {dev['vram_gb']}GB VRAM. "
                    "This GPU is recommended for faster processing."
                )
            else:
                self.gpu_info_label.setText(
                    f"Integrated GPU detected with {dev['vram_gb']}GB VRAM. "
                    "Consider using CPU mode to avoid out-of-memory errors. "
                    "marker-pdf typically needs 4GB+ VRAM."
                )

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

    def _get_device_config(self) -> Tuple[str, int]:
        """Get the device string and VRAM allocation."""
        device_data = self.device_combo.currentData()

        if device_data == "cpu":
            return "cpu", 4

        if device_data and device_data.startswith("cuda:"):
            idx = int(device_data.split(":")[1])
            dev = self.gpu_info["devices"][idx]
            # Set CUDA_VISIBLE_DEVICES to use specific GPU
            os.environ["CUDA_VISIBLE_DEVICES"] = str(idx)
            return "cuda", int(dev["vram_gb"])

        return "cpu", 4

    def run_process(self):
        files = self._selected_files()
        if not files:
            MessageBox("Warning", "No PDF files selected.", self).exec()
            return

        device, vram_gb = self._get_device_config()

        self.log_box.clear()
        self.log_message.emit(f"Starting conversion for {len(files)} file(s)...")
        self.run_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.device_combo.setEnabled(False)

        def worker():
            try:
                ok, fail, outputs = process_files(files, device, vram_gb, log_emit=self.log_message.emit)
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
            self.log_message.emit("Generated files:")
            for p in outputs:
                self.log_message.emit(f" - {p}")
        self.log_message.emit(f"Completed: {ok} succeeded, {fail} failed.")
        self.run_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.device_combo.setEnabled(True)


def get_widget():
    return MainWidget()
