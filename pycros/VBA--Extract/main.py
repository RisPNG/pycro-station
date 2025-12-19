#!/usr/bin/env python3
from __future__ import annotations

import argparse
import threading
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

from oletools.olevba import VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML

# GUI Imports (for Pycro Station)
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
    QTextEdit,
    QWidget,
    QSizePolicy,
)
from qfluentwidgets import PrimaryPushButton, MessageBox


# ===== Core logic =====
def _emit(log_emit, text: str):
    if callable(log_emit):
        try:
            log_emit(text)
            return
        except Exception:
            pass
    print(text)


def get_file_type_name(vba_parser: VBA_Parser) -> str:
    """Get human-readable file type name."""
    type_map = {
        TYPE_OLE: "OLE (Binary)",
        TYPE_OpenXML: "OpenXML",
        TYPE_Word2003_XML: "Word 2003 XML",
        TYPE_MHTML: "MHTML",
    }
    return type_map.get(vba_parser.type, "Unknown")


def extract_vba_from_file(file_path: Path, log_emit=None) -> Tuple[Optional[str], int]:
    """
    Extract VBA code from a macro-enabled file.

    Returns:
        Tuple of (formatted_output, module_count)
    """
    _emit(log_emit, f"Processing: {file_path.name}")

    try:
        vba_parser = VBA_Parser(str(file_path))
    except Exception as e:
        _emit(log_emit, f"  [ERROR] Failed to parse file: {e}")
        return None, 0

    if not vba_parser.detect_vba_macros():
        _emit(log_emit, f"  [SKIP] No VBA macros found in: {file_path.name}")
        vba_parser.close()
        return None, 0

    lines: List[str] = []
    module_count = 0

    # Header
    lines.append("=" * 80)
    lines.append(f"VBA CODE EXTRACTION REPORT")
    lines.append("=" * 80)
    lines.append(f"Source File: {file_path.name}")
    lines.append(f"Full Path:   {file_path}")
    lines.append(f"File Type:   {get_file_type_name(vba_parser)}")
    lines.append(f"Extracted:   {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 80)
    lines.append("")

    # Extract all VBA modules
    for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
        if vba_code and vba_code.strip():
            module_count += 1

            lines.append("-" * 80)
            lines.append(f"MODULE #{module_count}")
            lines.append("-" * 80)
            lines.append(f"  Stream Path: {stream_path}")
            lines.append(f"  VBA Filename: {vba_filename}")
            lines.append("-" * 80)
            lines.append("")
            lines.append(vba_code)
            lines.append("")
            lines.append("")

    vba_parser.close()

    if module_count == 0:
        _emit(log_emit, f"  [SKIP] No VBA code content found in: {file_path.name}")
        return None, 0

    # Footer
    lines.append("=" * 80)
    lines.append(f"END OF EXTRACTION - {module_count} module(s) extracted")
    lines.append("=" * 80)

    _emit(log_emit, f"  [OK] Extracted {module_count} module(s)")

    return "\n".join(lines), module_count


def proposed_output_path(input_file: Path) -> Path:
    """Generate output path with timestamp suffix."""
    base_dir = input_file.parent
    stem = input_file.stem
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"{stem}_VBA_{timestamp}.txt"
    candidate = base_dir / output_name
    if not candidate.exists():
        return candidate
    n = 1
    while True:
        candidate = base_dir / f"{stem}_VBA_{timestamp} ({n}).txt"
        if not candidate.exists():
            return candidate
        n += 1


def process_file(
    input_path: str,
    output_path: Optional[str] = None,
    log_emit=None,
) -> Tuple[str, int]:
    """
    Process a single xlsm file and extract VBA code.

    Returns:
        Tuple of (output_file_path, module_count)
    """
    file_path = Path(input_path)

    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    if not file_path.is_file():
        raise ValueError(f"Not a file: {file_path}")

    valid_extensions = {".xlsm", ".xlsb", ".xls", ".xlam", ".xla", ".docm", ".dotm", ".pptm"}
    if file_path.suffix.lower() not in valid_extensions:
        raise ValueError(f"Unsupported file type: {file_path.suffix}. Supported: {', '.join(valid_extensions)}")

    _emit(log_emit, f"Input file: {file_path}")
    _emit(log_emit, "")

    content, module_count = extract_vba_from_file(file_path, log_emit=log_emit)

    if content is None or module_count == 0:
        raise ValueError("No VBA code found in the file.")

    out_path = Path(output_path) if output_path else proposed_output_path(file_path)

    _emit(log_emit, "")
    _emit(log_emit, f"Writing output to: {out_path}")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(content)

    _emit(log_emit, f"[DONE] Successfully extracted {module_count} VBA module(s)")

    return str(out_path), module_count


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Extract VBA code from Excel macro-enabled files (.xlsm, .xlsb, .xls, etc.)"
    )
    ap.add_argument(
        "input",
        help="Input macro-enabled file (.xlsm, .xlsb, .xls, .xlam, .docm, .pptm, etc.)",
    )
    ap.add_argument(
        "--output",
        "-o",
        help="Output .txt file path. If omitted, auto-generated beside input file.",
    )
    args = ap.parse_args()

    try:
        out_path, module_count = process_file(args.input, args.output)
        print(f"\nDone. Extracted {module_count} module(s).")
        print(f"Output file: {out_path}")
    except Exception as e:
        print(f"Error: {e}")
        exit(1)


if __name__ == "__main__":
    main()


# --- UI Class (Pycro Station) ---

class MainWidget(QWidget):
    log_message = Signal(str)
    processing_done = Signal(int, str)

    def __init__(self):
        super().__init__()
        self.setObjectName("vba_extract_widget")
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
        self.set_long_description("")

        self.select_btn = PrimaryPushButton("Select Macro-Enabled File", self)
        self.run_btn = PrimaryPushButton("Extract VBA Code", self)

        self.file_label = QLabel("Selected file", self)
        self.file_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.file_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.logs_label = QLabel("Process logs", self)
        self.logs_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.logs_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        shared_style = (
            "QTextEdit{background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px;}"
        )

        self.file_box = QTextEdit(self)
        self.file_box.setReadOnly(True)
        self.file_box.setPlaceholderText("Selected macro-enabled file will appear here")
        self.file_box.setStyleSheet(shared_style)

        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Live process log will appear here")
        self.log_box.setStyleSheet(shared_style)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        main_layout.addWidget(self.desc_label, 1)

        row1 = QHBoxLayout()
        row1.addStretch(1)
        row1.addWidget(self.select_btn, 1)
        row1.addStretch(1)
        main_layout.addLayout(row1, 0)

        row2 = QHBoxLayout()
        row2.addStretch(1)
        row2.addWidget(self.run_btn, 1)
        row2.addStretch(1)
        main_layout.addLayout(row2, 0)

        row3 = QHBoxLayout()
        row3.addWidget(self.file_label, 1)
        row3.addWidget(self.logs_label, 1)
        main_layout.addLayout(row3, 0)

        row4 = QHBoxLayout()
        row4.addWidget(self.file_box, 1)
        row4.addWidget(self.log_box, 1)
        main_layout.addLayout(row4, 4)

    def set_long_description(self, text: str):
        clean = (text or "").strip()
        if clean:
            self.desc_label.setText(clean)
            self.desc_label.show()
        else:
            self.desc_label.clear()
            self.desc_label.hide()

    def _connect_signals(self):
        self.select_btn.clicked.connect(self.select_file)
        self.run_btn.clicked.connect(self.run_process)
        self.log_message.connect(self.append_log)
        self.processing_done.connect(self.on_done)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Macro-Enabled File",
            "",
            "Macro-Enabled Files (*.xlsm *.xlsb *.xls *.xlam *.xla *.docm *.dotm *.pptm);;All Files (*)",
        )
        if file_path:
            self.file_box.setPlainText(file_path)
        else:
            self.file_box.clear()

    def _selected_file(self) -> Optional[str]:
        text = self.file_box.toPlainText().strip()
        return text if text else None

    def run_process(self):
        input_file = self._selected_file()
        if not input_file:
            MessageBox("Warning", "Please select a macro-enabled file to extract VBA from.", self).exec()
            return

        self.log_box.clear()
        self.log_message.emit("VBA extraction started...")
        self.log_message.emit("")

        self.run_btn.setEnabled(False)
        self.select_btn.setEnabled(False)

        def worker():
            try:
                out_path, module_count = process_file(input_file, log_emit=self.log_message.emit)
                self.processing_done.emit(module_count, out_path)
            except Exception as e:
                self.log_message.emit(f"ERROR: {e}")
                self.processing_done.emit(0, "")

        threading.Thread(target=worker, daemon=True).start()

    def append_log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def on_done(self, module_count: int, out_path: str):
        self.log_message.emit("")
        if module_count > 0:
            self.log_message.emit(f"Extraction complete: {module_count} module(s) extracted")
            self.log_message.emit(f"Output: {out_path}")
        else:
            self.log_message.emit("Extraction failed or no VBA code found.")

        self.run_btn.setEnabled(True)
        self.select_btn.setEnabled(True)

        if module_count > 0:
            title = "Extraction Complete"
            lines = [f"Extracted {module_count} VBA module(s)", f"Output: {Path(out_path).name}"]
        else:
            title = "Extraction Failed"
            lines = ["No VBA code was extracted.", "Please check the log for details."]

        msg = MessageBox(title, "\n".join(lines), self)
        msg.yesButton.setText("OK")
        msg.cancelButton.hide()
        msg.exec()


def get_widget():
    return MainWidget()
