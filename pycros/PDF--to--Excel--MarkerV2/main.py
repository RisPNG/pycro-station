#!/usr/bin/env python3
"""
PDF to Excel converter with LLM-powered document understanding.

Pipeline:
1. marker-pdf: Extract tables/text from ANY PDF (text or scanned)
2. LLM (Ministral-3B or similar): Understand document context and structure
3. Output: Intelligently structured Excel

Optimized for systems without discrete GPUs (Intel Iris Xe, AMD APU).
"""
from __future__ import annotations

import json
import os
import re
import threading
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Tuple, Optional, Dict, Any

import pandas as pd
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QSizePolicy,
    QCheckBox,
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


def check_dependencies() -> Tuple[bool, str, bool]:
    """
    Check if required dependencies are installed.
    Returns (all_ok, error_message, llm_available)
    """
    missing = []
    llm_available = False

    try:
        import marker
    except ImportError:
        missing.append("marker-pdf")

    try:
        import pandas
    except ImportError:
        missing.append("pandas")

    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")

    try:
        import torch
    except ImportError:
        missing.append("torch")

    # Check for LLM support (optional but recommended)
    try:
        from llama_cpp import Llama
        llm_available = True
    except ImportError:
        pass  # LLM is optional

    if missing:
        msg = "Missing dependencies:\n" + "\n".join(f"  - {m}" for m in missing)
        msg += "\n\nInstall with:\n  pip install marker-pdf pandas openpyxl torch"
        return False, msg, llm_available

    return True, "", llm_available


def clean_cell(value: str | None) -> str:
    """Clean a cell value: remove HTML, normalize whitespace."""
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r'<br\s*/?>', ' ', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def parse_amount(value: str) -> Optional[float]:
    """Parse an amount string into a float."""
    if not value:
        return None

    text = clean_cell(value)
    is_negative = text.startswith('(') and text.endswith(')')
    if is_negative:
        text = text[1:-1].strip()

    text = re.sub(r'\s*(USD|EUR|GBP|MYR|SGD|JPY|CNY|IDR|THB|VND|PHP)\s*', '', text, flags=re.IGNORECASE)
    text = text.replace(',', '')
    text = re.sub(r'[^\d.\-]', '', text)

    try:
        amount = float(text)
        return -amount if is_negative else amount
    except ValueError:
        return None


# =============================================================================
# LLM Translation Layer
# =============================================================================

class LLMTranslator:
    """
    Uses a local LLM (via llama-cpp-python) to understand document context
    and decide how to structure the Excel output.
    """

    def __init__(self, model_path: str, log_emit: Callable[[str], None] | None = None):
        self.model_path = model_path
        self.log_emit = log_emit
        self.llm = None

    def load_model(self):
        """Load the LLM model."""
        from llama_cpp import Llama

        _emit(self.log_emit, f"Loading LLM model: {Path(self.model_path).name}")
        _emit(self.log_emit, "This may take a moment on first load...")

        self.llm = Llama(
            model_path=self.model_path,
            n_ctx=8192,  # Context window
            n_threads=4,  # CPU threads (adjust for your CPU)
            n_gpu_layers=0,  # CPU only
            verbose=False,
        )
        _emit(self.log_emit, "LLM model loaded successfully")

    def analyze_document(self, extracted_text: str, tables_summary: str) -> Dict[str, Any]:
        """
        Ask the LLM to analyze the document and provide structuring instructions.

        Returns a dict with:
        - document_type: str (e.g., "payment_settlement", "invoice", "report")
        - sheets: list of sheet definitions
        - column_mappings: how to map extracted data to columns
        """
        if not self.llm:
            self.load_model()

        # Truncate if too long (leave room for response)
        max_content_len = 6000
        if len(extracted_text) > max_content_len:
            extracted_text = extracted_text[:max_content_len] + "\n... [truncated]"

        prompt = f"""<s>[INST] You are a document analysis assistant. Analyze this extracted PDF content and determine how to structure it into an Excel file.

EXTRACTED CONTENT:
{extracted_text}

TABLES FOUND:
{tables_summary}

Respond with a JSON object containing:
1. "document_type": What type of document is this (e.g., "payment_settlement", "invoice", "financial_report", "data_table")
2. "sheets": Array of sheet definitions, each with:
   - "name": Sheet name (max 31 chars)
   - "description": What data this sheet contains
   - "columns": Array of column names to extract
   - "data_pattern": Regex or description to identify rows for this sheet
3. "notes": Any important observations about the document structure

Focus on creating a clean, usable Excel structure. Identify the main data tables and their columns.

Respond ONLY with valid JSON, no other text. [/INST]"""

        _emit(self.log_emit, "Asking LLM to analyze document structure...")

        try:
            response = self.llm(
                prompt,
                max_tokens=2048,
                temperature=0.1,  # Low temperature for consistent output
                stop=["</s>", "[INST]"],
            )

            response_text = response["choices"][0]["text"].strip()

            # Try to extract JSON from response
            json_match = re.search(r'\{[\s\S]*\}', response_text)
            if json_match:
                result = json.loads(json_match.group())
                _emit(self.log_emit, f"LLM identified document type: {result.get('document_type', 'unknown')}")
                return result
            else:
                _emit(self.log_emit, "LLM response was not valid JSON, using fallback")
                return self._fallback_analysis(extracted_text)

        except json.JSONDecodeError as e:
            _emit(self.log_emit, f"Failed to parse LLM response: {e}")
            return self._fallback_analysis(extracted_text)
        except Exception as e:
            _emit(self.log_emit, f"LLM error: {e}")
            return self._fallback_analysis(extracted_text)

    def _fallback_analysis(self, text: str) -> Dict[str, Any]:
        """Fallback analysis when LLM fails."""
        return {
            "document_type": "unknown",
            "sheets": [
                {
                    "name": "Data",
                    "description": "Extracted data",
                    "columns": [],
                    "data_pattern": ".*"
                }
            ],
            "notes": "Fallback mode - LLM analysis failed"
        }

    def structure_data(
        self,
        tables: List[List[List[str]]],
        analysis: Dict[str, Any],
        log_emit: Callable[[str], None] | None = None
    ) -> Dict[str, pd.DataFrame]:
        """
        Use LLM analysis to structure the extracted tables into DataFrames.
        """
        if not self.llm:
            self.load_model()

        dataframes = {}

        # Convert tables to a format we can work with
        all_rows = []
        for table in tables:
            for row in table:
                cleaned_row = [clean_cell(c) for c in row]
                if any(cleaned_row):
                    all_rows.append(cleaned_row)

        if not all_rows:
            return dataframes

        # For each sheet defined by LLM, try to extract relevant data
        for sheet_def in analysis.get("sheets", []):
            sheet_name = sheet_def.get("name", "Data")[:31]
            columns = sheet_def.get("columns", [])
            pattern = sheet_def.get("data_pattern", "")

            _emit(log_emit, f"Processing sheet: {sheet_name}")

            # Try to match rows to this sheet
            sheet_rows = []
            header_row = None

            for row in all_rows:
                row_text = ' '.join(row).lower()

                # Check if this looks like a header row for our columns
                if columns and not header_row:
                    matches = sum(1 for col in columns if col.lower() in row_text)
                    if matches >= len(columns) * 0.5:  # At least 50% column match
                        header_row = row
                        continue

                # Check if row matches pattern
                if pattern:
                    try:
                        if re.search(pattern, row_text, re.IGNORECASE):
                            sheet_rows.append(row)
                    except re.error:
                        # Invalid regex, just add all data rows
                        if header_row and row != header_row:
                            sheet_rows.append(row)
                else:
                    sheet_rows.append(row)

            if sheet_rows:
                # Create DataFrame
                if header_row:
                    # Normalize column count
                    max_cols = max(len(header_row), max(len(r) for r in sheet_rows))
                    header_row = header_row + [''] * (max_cols - len(header_row))
                    sheet_rows = [r + [''] * (max_cols - len(r)) for r in sheet_rows]
                    df = pd.DataFrame(sheet_rows, columns=header_row[:max_cols])
                else:
                    df = pd.DataFrame(sheet_rows)

                # Convert amount columns
                for col in df.columns:
                    if any(kw in str(col).lower() for kw in ['amount', 'fee', 'total', 'price', 'principal']):
                        df[col] = df[col].apply(lambda x: parse_amount(str(x)) if pd.notna(x) else None)

                dataframes[sheet_name] = df
                _emit(log_emit, f"  Created sheet '{sheet_name}': {len(df)} rows")

        # If no sheets created, create a default one
        if not dataframes and all_rows:
            _emit(log_emit, "No specific sheets matched, creating default Data sheet")
            max_cols = max(len(r) for r in all_rows)
            normalized = [r + [''] * (max_cols - len(r)) for r in all_rows]
            dataframes["Data"] = pd.DataFrame(normalized)

        return dataframes


# =============================================================================
# Marker PDF Extraction
# =============================================================================

def extract_table_from_block(block) -> List[List[str]]:
    """Extract table data from a marker-pdf block."""
    rows = []

    if hasattr(block, 'cells') and block.cells:
        max_row = max((c.row_id for c in block.cells), default=0) + 1
        max_col = max((c.col_id for c in block.cells), default=0) + 1
        grid = [['' for _ in range(max_col)] for _ in range(max_row)]

        for cell in block.cells:
            text = clean_cell(cell.text if hasattr(cell, 'text') else str(cell))
            if cell.row_id < max_row and cell.col_id < max_col:
                grid[cell.row_id][cell.col_id] = text
        rows = grid

    elif hasattr(block, 'children'):
        for row_block in block.children:
            if hasattr(row_block, 'children'):
                row = []
                for cell_block in row_block.children:
                    text = ''
                    if hasattr(cell_block, 'text'):
                        text = clean_cell(cell_block.text)
                    elif hasattr(cell_block, 'children'):
                        texts = [elem.text for elem in cell_block.children if hasattr(elem, 'text')]
                        text = clean_cell(' '.join(texts))
                    row.append(text)
                if row:
                    rows.append(row)
    return rows


def extract_tables_from_marker(rendered, log_emit=None) -> List[List[List[str]]]:
    """Extract all tables from marker-pdf's rendered document."""
    tables = []

    def traverse(block):
        if hasattr(block, 'block_type') and block.block_type == 'Table':
            table_data = extract_table_from_block(block)
            if table_data:
                tables.append(table_data)
                _emit(log_emit, f"  Found table: {len(table_data)} rows")

        if hasattr(block, 'children'):
            for child in block.children:
                traverse(child)

    traverse(rendered)
    return tables


def parse_markdown_tables(text: str, log_emit=None) -> List[List[List[str]]]:
    """Parse markdown tables from text."""
    lines = text.strip().split("\n")
    tables = []
    current_table = []
    in_table = False

    for line in lines:
        line_stripped = line.strip()

        if line_stripped.startswith("|") and line_stripped.endswith("|"):
            cells = [clean_cell(c) for c in line_stripped.split("|")[1:-1]]
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


def merge_continuation_tables(tables: List[List[List[str]]], log_emit=None) -> List[List[List[str]]]:
    """Merge tables split across pages."""
    if len(tables) <= 1:
        return tables

    merged = []
    i = 0

    while i < len(tables):
        current = tables[i]
        if not current:
            i += 1
            continue

        max_cols = max(len(row) for row in current) if current else 0
        j = i + 1
        merge_count = 0

        while j < len(tables):
            next_table = tables[j]
            if not next_table:
                j += 1
                continue

            # Check if can merge (similar column count, no header in second table)
            next_cols = len(next_table[0]) if next_table else 0
            if abs(max_cols - next_cols) <= 1:
                # Check if first row of next table looks like header
                first_row_text = ' '.join(clean_cell(c) for c in next_table[0]).lower()
                header_words = ['number', 'amount', 'type', 'date', 'name', 'contract', 'invoice']
                is_header = sum(1 for w in header_words if w in first_row_text) >= 2

                if not is_header:
                    for row in next_table:
                        if len(row) < max_cols:
                            row = row + [''] * (max_cols - len(row))
                        current.append(row)
                    merge_count += 1
                    j += 1
                    continue
            break

        if merge_count > 0:
            _emit(log_emit, f"  Merged {merge_count + 1} table continuations")

        merged.append(current)
        i = j

    return merged


# =============================================================================
# Main Conversion
# =============================================================================

def convert_pdf_to_excel(
    pdf_path: Path,
    model_path: str | None = None,
    log_emit: Callable[[str], None] | None = None,
) -> Path:
    """
    Convert a PDF file to Excel using marker-pdf + LLM understanding.
    """
    os.environ["TORCH_DEVICE"] = "cpu"

    _emit(log_emit, f"Processing: {pdf_path.name}")
    _emit(log_emit, "="*50)

    # Step 1: Extract with marker-pdf
    _emit(log_emit, "\n[Step 1/3] Extracting content with marker-pdf...")
    _emit(log_emit, "This handles both text-based and scanned PDFs.")

    from marker.converters.pdf import PdfConverter
    from marker.models import create_model_dict
    from marker.output import text_from_rendered

    _emit(log_emit, "Loading document understanding models...")
    converter = PdfConverter(artifact_dict=create_model_dict())

    _emit(log_emit, "Analyzing document layout...")
    rendered = converter(str(pdf_path))

    # Get markdown text
    markdown_text, _, _ = text_from_rendered(rendered)

    # Extract tables
    _emit(log_emit, "Extracting tables...")
    tables = extract_tables_from_marker(rendered, log_emit)

    if not tables:
        _emit(log_emit, "No structured tables found, parsing markdown...")
        tables = parse_markdown_tables(markdown_text, log_emit)

    if not tables:
        _emit(log_emit, "Extracting as text lines...")
        lines = [clean_cell(line) for line in markdown_text.split("\n") if line.strip()]
        if lines:
            tables = [[[line] for line in lines]]

    if not tables:
        raise ValueError("No content could be extracted from the PDF")

    # Merge continuation tables
    tables = merge_continuation_tables(tables, log_emit)
    _emit(log_emit, f"Extracted {len(tables)} table(s)")

    # Create tables summary for LLM
    tables_summary = []
    for i, table in enumerate(tables):
        if table:
            cols = len(table[0]) if table else 0
            sample = table[0][:3] if table else []
            tables_summary.append(f"Table {i+1}: {len(table)} rows x {cols} cols. Sample: {sample}")

    # Step 2: LLM Analysis (if model provided)
    dataframes = {}

    if model_path and Path(model_path).exists():
        _emit(log_emit, "\n[Step 2/3] LLM analyzing document context...")

        try:
            translator = LLMTranslator(model_path, log_emit)
            analysis = translator.analyze_document(
                markdown_text[:6000],
                "\n".join(tables_summary)
            )

            _emit(log_emit, f"Document type: {analysis.get('document_type', 'unknown')}")
            if analysis.get('notes'):
                _emit(log_emit, f"Notes: {analysis.get('notes')}")

            # Step 3: Structure data based on LLM analysis
            _emit(log_emit, "\n[Step 3/3] Structuring data based on LLM analysis...")
            dataframes = translator.structure_data(tables, analysis, log_emit)

        except Exception as e:
            _emit(log_emit, f"LLM processing failed: {e}")
            _emit(log_emit, "Falling back to basic structuring...")

    # Fallback: Basic structuring if LLM not used or failed
    if not dataframes:
        _emit(log_emit, "\n[Step 2-3/3] Basic table structuring (no LLM)...")

        for idx, table in enumerate(tables):
            if not table:
                continue

            # Clean and normalize
            table = [[clean_cell(c) for c in row] for row in table]
            table = [row for row in table if any(row)]

            if not table:
                continue

            max_cols = max(len(row) for row in table)
            table = [row + [''] * (max_cols - len(row)) for row in table]

            # Detect header
            first_row = table[0]
            first_text = ' '.join(first_row).lower()
            header_words = ['number', 'amount', 'type', 'date', 'contract', 'invoice', 'fee']

            if sum(1 for w in header_words if w in first_text) >= 2:
                df = pd.DataFrame(table[1:], columns=first_row)
            else:
                df = pd.DataFrame(table)

            # Convert amounts
            for col in df.columns:
                if any(kw in str(col).lower() for kw in ['amount', 'fee', 'total', 'price']):
                    df[col] = df[col].apply(lambda x: parse_amount(str(x)) if pd.notna(x) else None)

            sheet_name = f"Table_{idx+1}" if idx > 0 else "Data"
            dataframes[sheet_name] = df
            _emit(log_emit, f"  Sheet '{sheet_name}': {len(df)} rows")

    if not dataframes:
        raise ValueError("No data could be structured from the PDF")

    # Write Excel
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"{pdf_path.stem}__converted__{ts}.xlsx"
    out_path = ensure_unique_path(pdf_path.with_name(out_name))

    _emit(log_emit, f"\nWriting to {out_path.name}...")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name, df in dataframes.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    _emit(log_emit, f"\nCreated Excel with {len(dataframes)} sheet(s)")
    return out_path


def process_files(
    files: List[str],
    model_path: str | None = None,
    log_emit: Callable[[str], None] | None = None,
) -> Tuple[int, int, List[Path]]:
    """Process multiple PDF files."""
    ok = 0
    fail = 0
    outputs: List[Path] = []

    for raw in files:
        path = Path(raw)
        _emit(log_emit, f"\n{'='*60}")
        _emit(log_emit, f"[START] {path.name}")
        try:
            out = convert_pdf_to_excel(path, model_path, log_emit)
            outputs.append(out)
            ok += 1
            _emit(log_emit, f"[DONE] {path.name}")
        except Exception as e:
            fail += 1
            _emit(log_emit, f"[FAIL] {path.name}: {e}")
            import traceback
            _emit(log_emit, traceback.format_exc())

    return ok, fail, outputs


# =============================================================================
# UI
# =============================================================================

class MainWidget(QWidget):
    log_message = Signal(str)
    processing_done = Signal(int, int, list)

    def __init__(self):
        super().__init__()
        self.setObjectName("pdf_to_excel_plumber_widget")
        self._deps_ok = True
        self._deps_msg = ""
        self._llm_available = False
        self._check_deps()
        self._build_ui()
        self._connect_signals()

    def _check_deps(self):
        self._deps_ok, self._deps_msg, self._llm_available = check_dependencies()

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

        # Status
        self.status_label = QLabel("", self)
        self.status_label.setWordWrap(True)

        if not self._deps_ok:
            self.status_label.setText(f"⚠️ {self._deps_msg}")
            self.status_label.setStyleSheet("color: #ffcc00; padding: 4px;")
        elif self._llm_available:
            self.status_label.setText("✓ marker-pdf + LLM available (full context understanding)")
            self.status_label.setStyleSheet("color: #90EE90; padding: 4px;")
        else:
            self.status_label.setText("✓ marker-pdf ready | ⚠️ LLM not installed (pip install llama-cpp-python)")
            self.status_label.setStyleSheet("color: #ffcc00; padding: 4px;")

        # Controls
        self.select_btn = PrimaryPushButton("Select PDF File(s)", self)
        self.run_btn = PrimaryPushButton("Convert to Excel", self)
        if not self._deps_ok:
            self.run_btn.setEnabled(False)

        # LLM Model Path
        self.model_label = QLabel("LLM Model Path (GGUF file):", self)
        self.model_label.setStyleSheet("color: #dcdcdc; padding-left: 2px;")

        self.model_input = QLineEdit(self)
        self.model_input.setPlaceholderText("e.g., /path/to/Ministral-3B-Instruct-Q4_K_M.gguf (optional)")
        self.model_input.setStyleSheet(
            "QLineEdit{background: #1f1f1f; color: #d0d0d0; padding: 6px; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        self.model_browse_btn = PrimaryPushButton("Browse", self)
        self.model_browse_btn.setMaximumWidth(80)

        self.use_llm_check = QCheckBox("Enable LLM context understanding", self)
        self.use_llm_check.setChecked(True)
        self.use_llm_check.setStyleSheet("color: #dcdcdc;")
        self.use_llm_check.setEnabled(self._llm_available)

        # Files / Log
        self.files_label = QLabel("Selected PDF files", self)
        self.files_label.setStyleSheet("color: #dcdcdc; padding-left: 2px;")

        self.logs_label = QLabel("Process logs", self)
        self.logs_label.setStyleSheet("color: #dcdcdc; padding-left: 2px;")

        self.files_box = QTextEdit(self)
        self.files_box.setReadOnly(True)
        self.files_box.setPlaceholderText("Selected PDF files will appear here")
        self.files_box.setStyleSheet(
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText(
            "Processing pipeline:\n"
            "1. marker-pdf: Extract from any PDF (text/scanned)\n"
            "2. LLM: Understand context & structure (optional)\n"
            "3. Output: Intelligently structured Excel\n\n"
            "Estimated time: 2-5 min/page (CPU mode)"
        )
        self.log_box.setStyleSheet(
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        # Layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        main_layout.addWidget(self.desc_label)
        main_layout.addWidget(self.status_label)

        # Select button
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_row.addWidget(self.select_btn)
        btn_row.addStretch(1)
        main_layout.addLayout(btn_row)

        # Model path
        main_layout.addWidget(self.model_label)
        model_row = QHBoxLayout()
        model_row.addWidget(self.model_input, 1)
        model_row.addWidget(self.model_browse_btn)
        main_layout.addLayout(model_row)
        main_layout.addWidget(self.use_llm_check)

        # Run button
        run_row = QHBoxLayout()
        run_row.addStretch(1)
        run_row.addWidget(self.run_btn)
        run_row.addStretch(1)
        main_layout.addLayout(run_row)

        # Files and logs
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
        self.model_browse_btn.clicked.connect(self.browse_model)
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
        body = [ln for ln in lines if not ln.startswith(">")]
        text = "\n".join(body).strip()
        if text:
            self.desc_label.setText(text)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF files", "", "PDF Files (*.pdf);;All Files (*)"
        )
        if files:
            self.files_box.setPlainText("\n".join(files))

    def browse_model(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select GGUF Model", "", "GGUF Models (*.gguf);;All Files (*)"
        )
        if file:
            self.model_input.setText(file)

    def _selected_files(self) -> List[str]:
        text = self.files_box.toPlainText().strip()
        return [line for line in text.split("\n") if line.strip()] if text else []

    def run_process(self):
        if not self._deps_ok:
            MessageBox("Error", "Missing dependencies.", self).exec()
            return

        files = self._selected_files()
        if not files:
            MessageBox("Warning", "No files selected.", self).exec()
            return

        model_path = None
        if self.use_llm_check.isChecked():
            model_path = self.model_input.text().strip()
            if model_path and not Path(model_path).exists():
                MessageBox("Warning", f"Model file not found: {model_path}\nProceeding without LLM.", self).exec()
                model_path = None

        self.log_box.clear()
        self.log_message.emit(f"Starting conversion for {len(files)} file(s)...")
        if model_path:
            self.log_message.emit(f"LLM: {Path(model_path).name}")
        else:
            self.log_message.emit("LLM: Not configured (basic structuring only)")
        self.log_message.emit("")

        self.run_btn.setEnabled(False)
        self.select_btn.setEnabled(False)

        def worker():
            try:
                ok, fail, outputs = process_files(files, model_path, self.log_message.emit)
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
            self.log_message.emit("\n" + "="*50)
            self.log_message.emit("Generated files:")
            for p in outputs:
                self.log_message.emit(f"  {p}")
        self.log_message.emit(f"\nCompleted: {ok} succeeded, {fail} failed.")
        self.run_btn.setEnabled(True)
        self.select_btn.setEnabled(True)


def get_widget():
    return MainWidget()
