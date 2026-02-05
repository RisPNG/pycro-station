#!/usr/bin/env python3
from __future__ import annotations

import random
import threading
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, List, Tuple

import pandas as pd
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QSizePolicy,
)
from qfluentwidgets import PrimaryPushButton, MessageBox

SUPPORTED_EXTS = {".csv", ".xlsx", ".xls"}


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


def read_table(path: Path) -> pd.DataFrame:
    """Load CSV/XLS/XLSX without assuming headers."""
    ext = path.suffix.lower()
    if ext not in SUPPORTED_EXTS:
        raise ValueError(f"Unsupported file type: {ext}")

    if ext == ".csv":
        return pd.read_csv(path, header=None, dtype=object)

    engine = "xlrd" if ext == ".xls" else "openpyxl"
    return pd.read_excel(path, header=None, dtype=object, engine=engine)


def trim_table(df: pd.DataFrame, header_row: int, keep_count: int, log_emit=None) -> Tuple[pd.DataFrame, int, int]:
    """
    Keep the header row + a random subset of data rows.

    Returns (new_df, total_data_rows, kept_data_rows).
    """
    total_rows = df.shape[0]
    if total_rows == 0:
        raise ValueError("File is empty.")

    header_idx = header_row - 1
    if header_idx < 0 or header_idx >= total_rows:
        raise ValueError(f"Header row {header_row} is outside the file (rows=1..{total_rows}).")

    data_start = header_idx + 1
    data_count = max(total_rows - data_start, 0)
    target_keep = max(0, keep_count)

    if data_count == 0:
        _emit(log_emit, "No data rows found below the header; keeping the header only.")
        keep_indices = list(range(0, data_start))
        return df.iloc[keep_indices].reset_index(drop=True), 0, 0

    if target_keep >= data_count:
        _emit(log_emit, f"Requested to keep {target_keep} rows but only {data_count} available; keeping all data rows.")
        keep_indices = list(range(total_rows))
        return df.iloc[keep_indices].reset_index(drop=True), data_count, data_count

    selected = sorted(random.sample(range(data_count), target_keep))
    keep_indices = list(range(0, data_start)) + [data_start + idx for idx in selected]
    new_df = df.iloc[keep_indices].reset_index(drop=True)
    return new_df, data_count, target_keep


def write_table(df: pd.DataFrame, dest: Path, ext: str):
    """Persist the dataframe in the same format (no header/index written)."""
    if ext == ".csv":
        df.to_csv(dest, header=False, index=False)
        return

    engine = "xlwt" if ext == ".xls" else "openpyxl"
    df.to_excel(dest, header=False, index=False, engine=engine)


def process_file(path: Path, header_row: int, keep_rows: int, log_emit=None) -> Path:
    """Process a single file and return the output path."""
    if not path.exists():
        raise FileNotFoundError(path)

    ext = path.suffix.lower()
    if ext not in SUPPORTED_EXTS:
        raise ValueError(f"Unsupported file type: {ext}")

    _emit(log_emit, f"Reading {path.name} ...")
    df = read_table(path)

    new_df, total_data_rows, kept = trim_table(df, header_row, keep_rows, log_emit=log_emit)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"{path.stem}__kept_{kept}__{ts}{ext}"
    out_path = ensure_unique_path(path.with_name(out_name))
    write_table(new_df, out_path, ext)

    _emit(
        log_emit,
        f"Saved {out_path.name}: kept {kept} of {total_data_rows} data rows (header row preserved)."
    )
    return out_path


def process_files(files: Iterable[str], header_row: int, keep_rows: int, log_emit=None) -> Tuple[int, int, List[Path]]:
    """Process all files, returning (ok_count, fail_count, output_paths)."""
    ok = 0
    fail = 0
    outputs: List[Path] = []
    for raw in files:
        path = Path(raw)
        _emit(log_emit, f"[START] {path.name}")
        try:
            out = process_file(path, header_row, keep_rows, log_emit=log_emit)
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
        self.setObjectName("random_row_reducer_widget")
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
        self.select_btn = PrimaryPushButton("Select CSV / XLSX / XLS", self)
        self.run_btn = PrimaryPushButton("Run", self)

        self.header_spin = QSpinBox(self)
        self.header_spin.setMinimum(1)
        self.header_spin.setMaximum(1_000_000)
        self.header_spin.setValue(1)
        self.header_label = QLabel("Header row (1-based)", self)
        self.header_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.keep_spin = QSpinBox(self)
        self.keep_spin.setMinimum(0)
        self.keep_spin.setMaximum(1_000_000)
        self.keep_spin.setValue(100)
        self.keep_label = QLabel("Data rows to keep", self)
        self.keep_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        # Selected files / log
        self.files_label = QLabel("Selected files", self)
        self.files_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.files_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")
        self.logs_label = QLabel("Process logs", self)
        self.logs_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.logs_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.files_box = QTextEdit(self)
        self.files_box.setReadOnly(True)
        self.files_box.setPlaceholderText("Selected files will appear here")
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
        grid.addWidget(self.header_label, 0, 0, Qt.AlignLeft)
        grid.addWidget(self.header_spin, 0, 1, Qt.AlignLeft)
        grid.addWidget(self.keep_label, 1, 0, Qt.AlignLeft)
        grid.addWidget(self.keep_spin, 1, 1, Qt.AlignLeft)
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
            "Select CSV/XLSX/XLS files",
            "",
            "Spreadsheets (*.csv *.xlsx *.xls);;All Files (*)",
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
        files = self._selected_files()
        if not files:
            MessageBox("Warning", "No files selected.", self).exec()
            return

        header_row = int(self.header_spin.value())
        keep_rows = int(self.keep_spin.value())

        self.log_box.clear()
        self.log_message.emit(f"Starting run for {len(files)} file(s)...")
        self.run_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.header_spin.setEnabled(False)
        self.keep_spin.setEnabled(False)

        def worker():
            try:
                ok, fail, outputs = process_files(files, header_row, keep_rows, log_emit=self.log_message.emit)
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
        self.header_spin.setEnabled(True)
        self.keep_spin.setEnabled(True)
        title = "Processing complete" if fail == 0 else "Processing finished with issues"
        lines = [f"Success: {ok}", f"Failed: {fail}"]
        if outputs:
            try:
                sample = outputs[:5]
                lines.append("")
                lines.append("Outputs:")
                for p in sample:
                    name = getattr(p, "name", None) or str(p)
                    lines.append(f"- {name}")
                remaining = len(outputs) - len(sample)
                if remaining > 0:
                    lines.append(f"... and {remaining} more")
            except Exception:
                pass
        msg = MessageBox(title, "\n".join(lines), self)
        msg.yesButton.setText("OK")
        msg.cancelButton.hide()
        msg.exec()


def get_widget():
    return MainWidget()
