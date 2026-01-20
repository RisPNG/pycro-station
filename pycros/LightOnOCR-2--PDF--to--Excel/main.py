#!/usr/bin/env python3
from __future__ import annotations

import re
import threading
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Tuple, TYPE_CHECKING

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QSizePolicy,
)
from qfluentwidgets import PrimaryPushButton, MessageBox

# Lazy imports for heavy ML dependencies
if TYPE_CHECKING:
    import pandas as pd
    import pypdfium2 as pdfium
    import torch
    from PIL import Image
    from transformers import LightOnOcrForConditionalGeneration, LightOnOcrProcessor


def check_dependencies() -> Tuple[bool, str]:
    """
    Check if required dependencies are installed.

    Returns:
        Tuple of (all_ok, error_message)
    """
    missing = []

    try:
        import torch
    except ImportError:
        missing.append("torch")

    try:
        import pandas
    except ImportError:
        missing.append("pandas")

    try:
        import pypdfium2
    except ImportError:
        missing.append("pypdfium2")

    try:
        from PIL import Image
    except ImportError:
        missing.append("pillow")

    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")

    try:
        from transformers import LightOnOcrForConditionalGeneration, LightOnOcrProcessor
    except ImportError as e:
        if "No module named 'transformers'" in str(e):
            missing.append("transformers (from git)")
        else:
            missing.append(f"transformers with LightOnOCR support: {e}")

    if missing:
        msg = "Missing dependencies:\n" + "\n".join(f"  - {m}" for m in missing)
        msg += "\n\nInstall with:\n  pip install torch pandas pypdfium2 pillow openpyxl"
        msg += "\n  pip install git+https://github.com/huggingface/transformers"
        return False, msg

    return True, ""


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


def detect_gpu_info() -> dict:
    """
    Detect available GPU hardware and return info dict.

    Returns dict with:
        - devices: list of (device_str, display_name, dtype_str, is_discrete)
        - recommended: index of recommended device
    """
    devices = []
    recommended_idx = 0

    try:
        import torch
    except ImportError:
        # torch not installed, return CPU-only option
        devices.append(("cpu", "CPU - (torch not installed)", "float32", False))
        return {"devices": devices, "recommended": 0}

    # Check CUDA (NVIDIA discrete GPUs)
    if torch.cuda.is_available():
        for i in range(torch.cuda.device_count()):
            props = torch.cuda.get_device_properties(i)
            vram_gb = props.total_memory / (1024**3)
            name = f"CUDA:{i} - {props.name} ({vram_gb:.1f}GB)"
            # Use bfloat16 for CUDA (optimal for transformer models)
            devices.append((f"cuda:{i}", name, "bfloat16", True))

    # Check MPS (Apple Silicon)
    if torch.backends.mps.is_available():
        name = "MPS - Apple Silicon GPU"
        # Use float32 for MPS (bfloat16 support is limited)
        devices.append(("mps", name, "float32", True))

    # Always add CPU option
    cpu_name = "CPU - System Processor"
    # Check if this is likely integrated graphics only
    if not devices:
        cpu_name = "CPU - No discrete GPU detected (slower)"
    devices.append(("cpu", cpu_name, "float32", False))

    # Recommend first discrete GPU if available, otherwise CPU
    for i, (_, _, _, is_discrete) in enumerate(devices):
        if is_discrete:
            recommended_idx = i
            break

    return {"devices": devices, "recommended": recommended_idx}


def render_pdf_page(pdf_path: Path, page_num: int, target_size: int = 1540):
    """
    Render a PDF page to PIL Image at target longest dimension.

    Args:
        pdf_path: Path to PDF file
        page_num: 0-indexed page number
        target_size: Target size for longest dimension (default 1540px per model recommendation)

    Returns:
        PIL Image of the rendered page
    """
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(pdf_path)
    page = pdf[page_num]

    # Get page dimensions at 72 DPI (PDF standard)
    width, height = page.get_size()

    # Calculate scale to achieve target longest dimension
    longest = max(width, height)
    scale = target_size / longest

    # Render at calculated scale
    bitmap = page.render(scale=scale)
    pil_image = bitmap.to_pil()

    return pil_image


def parse_ocr_to_dataframe(ocr_text: str):
    """
    Parse OCR output text into a DataFrame.

    Attempts to detect table structures in the OCR output.
    Falls back to line-by-line text if no table structure found.
    """
    import pandas as pd

    lines = ocr_text.strip().split("\n")

    if not lines:
        return pd.DataFrame({"Text": [""]})

    # Try to detect markdown table format
    table_lines = []
    in_table = False

    for line in lines:
        # Check for markdown table separator (e.g., |---|---|)
        if re.match(r"^\s*\|[-:\s|]+\|\s*$", line):
            in_table = True
            continue

        # Check for table row (starts and ends with |)
        if line.strip().startswith("|") and line.strip().endswith("|"):
            cells = [cell.strip() for cell in line.strip()[1:-1].split("|")]
            table_lines.append(cells)
            in_table = True
        elif in_table and not line.strip():
            # Empty line ends table
            break

    if table_lines and len(table_lines) > 1:
        # Use first row as headers
        headers = table_lines[0]
        data = table_lines[1:]

        # Ensure all rows have same number of columns
        max_cols = max(len(row) for row in table_lines)
        headers = headers + [""] * (max_cols - len(headers))
        data = [row + [""] * (max_cols - len(row)) for row in data]

        return pd.DataFrame(data, columns=headers)

    # Try to detect tab-separated or multi-space separated values
    potential_rows = []
    for line in lines:
        if "\t" in line:
            cells = line.split("\t")
        elif "  " in line:  # Two or more spaces as delimiter
            cells = re.split(r"\s{2,}", line.strip())
        else:
            cells = [line]
        potential_rows.append(cells)

    if potential_rows:
        # Check if most rows have consistent column count
        col_counts = [len(row) for row in potential_rows]
        most_common_count = max(set(col_counts), key=col_counts.count)

        if most_common_count > 1:
            # Normalize rows to most common column count
            normalized = []
            for row in potential_rows:
                if len(row) < most_common_count:
                    row = row + [""] * (most_common_count - len(row))
                elif len(row) > most_common_count:
                    row = row[:most_common_count]
                normalized.append(row)

            return pd.DataFrame(normalized[1:], columns=normalized[0]) if len(normalized) > 1 else pd.DataFrame(normalized)

    # Fallback: single column with all text
    return pd.DataFrame({"Text": lines})


class OCRModel:
    """Singleton-ish class to manage model loading and inference."""

    _instance = None
    _model = None
    _processor = None
    _device = None
    _dtype = None

    @classmethod
    def load(cls, device: str, dtype_str: str, log_emit=None) -> "OCRModel":
        """Load or return cached model."""
        import torch
        from transformers import LightOnOcrForConditionalGeneration, LightOnOcrProcessor

        # Convert dtype string to torch dtype
        dtype_map = {
            "bfloat16": torch.bfloat16,
            "float16": torch.float16,
            "float32": torch.float32,
        }
        dtype = dtype_map.get(dtype_str, torch.float32)

        if cls._model is None or cls._device != device:
            _emit(log_emit, f"Loading LightOnOCR-2-1B on {device}...")

            cls._processor = LightOnOcrProcessor.from_pretrained("lightonai/LightOnOCR-2-1B")
            cls._model = LightOnOcrForConditionalGeneration.from_pretrained(
                "lightonai/LightOnOCR-2-1B",
                torch_dtype=dtype
            ).to(device)
            cls._device = device
            cls._dtype = dtype

            _emit(log_emit, "Model loaded successfully.")

        return cls

    @classmethod
    def run_ocr(cls, image, max_tokens: int = 2048) -> str:
        """Run OCR on a single image."""
        if cls._model is None:
            raise RuntimeError("Model not loaded. Call load() first.")

        # Prepare conversation format for the model
        conversation = [
            {
                "role": "user",
                "content": [{"type": "image", "image": image}]
            }
        ]

        inputs = cls._processor.apply_chat_template(
            conversation,
            add_generation_prompt=True,
            tokenize=True,
            return_dict=True,
            return_tensors="pt",
        )

        # Move inputs to device with appropriate dtype
        inputs = {
            k: v.to(device=cls._device, dtype=cls._dtype) if v.is_floating_point() else v.to(cls._device)
            for k, v in inputs.items()
        }

        # Generate output
        output_ids = cls._model.generate(**inputs, max_new_tokens=max_tokens)
        generated_ids = output_ids[0, inputs["input_ids"].shape[1]:]
        output_text = cls._processor.decode(generated_ids, skip_special_tokens=True)

        return output_text

    @classmethod
    def unload(cls):
        """Unload model to free memory."""
        if cls._model is not None:
            del cls._model
            del cls._processor
            cls._model = None
            cls._processor = None
            cls._device = None
            cls._dtype = None

            # Clear CUDA cache if available
            try:
                import torch
                if torch.cuda.is_available():
                    torch.cuda.empty_cache()
            except ImportError:
                pass


def process_pdf(
    pdf_path: Path,
    device: str,
    dtype_str: str,
    max_tokens: int,
    log_emit=None,
    progress_emit=None,
) -> Path:
    """
    Process a single PDF and return the output Excel path.

    Args:
        pdf_path: Path to input PDF
        device: Torch device string
        dtype_str: Dtype string for model (e.g. "bfloat16", "float32")
        max_tokens: Max tokens for OCR generation
        log_emit: Callback for log messages
        progress_emit: Callback for progress updates (current, total)

    Returns:
        Path to output Excel file
    """
    import pandas as pd
    import pypdfium2 as pdfium

    _emit(log_emit, f"Processing: {pdf_path.name}")

    # Load model
    OCRModel.load(device, dtype_str, log_emit)

    # Open PDF and get page count
    pdf = pdfium.PdfDocument(pdf_path)
    total_pages = len(pdf)
    _emit(log_emit, f"PDF has {total_pages} page(s)")

    all_dataframes = []

    for page_num in range(total_pages):
        _emit(log_emit, f"Processing page {page_num + 1}/{total_pages}...")

        if progress_emit:
            progress_emit(page_num, total_pages)

        # Render page to image
        image = render_pdf_page(pdf_path, page_num)

        # Run OCR
        ocr_text = OCRModel.run_ocr(image, max_tokens)
        _emit(log_emit, f"Page {page_num + 1} OCR complete ({len(ocr_text)} chars)")

        # Parse OCR output to DataFrame
        df = parse_ocr_to_dataframe(ocr_text)
        df.insert(0, "_Page", page_num + 1)
        all_dataframes.append(df)

    if progress_emit:
        progress_emit(total_pages, total_pages)

    # Combine all pages
    if all_dataframes:
        # Try to merge DataFrames with same columns
        if len(all_dataframes) > 1:
            try:
                combined_df = pd.concat(all_dataframes, ignore_index=True)
            except Exception:
                # If concat fails, just use first DataFrame
                combined_df = all_dataframes[0]
        else:
            combined_df = all_dataframes[0]
    else:
        combined_df = pd.DataFrame({"Text": ["No content extracted"]})

    # Generate output path
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"{pdf_path.stem}__ocr__{ts}.xlsx"
    out_path = ensure_unique_path(pdf_path.with_name(out_name))

    # Write to Excel
    combined_df.to_excel(out_path, index=False, engine="openpyxl")
    _emit(log_emit, f"Saved: {out_path.name}")

    return out_path


def process_files(
    files: List[str],
    device: str,
    dtype_str: str,
    max_tokens: int,
    log_emit=None,
    progress_emit=None,
) -> Tuple[int, int, List[Path]]:
    """Process all files, returning (ok_count, fail_count, output_paths)."""
    ok = 0
    fail = 0
    outputs: List[Path] = []

    for raw in files:
        path = Path(raw)
        _emit(log_emit, f"[START] {path.name}")
        try:
            out = process_pdf(path, device, dtype_str, max_tokens, log_emit, progress_emit)
            outputs.append(out)
            ok += 1
            _emit(log_emit, f"[DONE] {path.name}")
        except Exception as e:
            fail += 1
            _emit(log_emit, f"[FAIL] {path.name}: {e}")

    return ok, fail, outputs


class MainWidget(QWidget):
    log_message = Signal(str)
    progress_update = Signal(int, int)
    processing_done = Signal(int, int, list)

    def __init__(self):
        super().__init__()
        self.setObjectName("lightonocr2_pdf_to_excel_widget")
        self._gpu_info = detect_gpu_info()
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
        self.run_btn = PrimaryPushButton("Run OCR", self)
        self.unload_btn = PrimaryPushButton("Unload Model", self)
        self.unload_btn.setToolTip("Free GPU memory by unloading the model")

        # Device selection
        self.device_combo = QComboBox(self)
        for device_str, display_name, dtype, is_discrete in self._gpu_info["devices"]:
            self.device_combo.addItem(display_name, userData=(device_str, dtype))
        self.device_combo.setCurrentIndex(self._gpu_info["recommended"])
        self.device_label = QLabel("Compute Device", self)
        self.device_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        # Max tokens
        self.tokens_spin = QSpinBox(self)
        self.tokens_spin.setMinimum(256)
        self.tokens_spin.setMaximum(8192)
        self.tokens_spin.setValue(2048)
        self.tokens_spin.setSingleStep(256)
        self.tokens_label = QLabel("Max tokens per page", self)
        self.tokens_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        # Progress bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("Ready")

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
        grid.addWidget(self.tokens_label, 1, 0, Qt.AlignLeft)
        grid.addWidget(self.tokens_spin, 1, 1, Qt.AlignLeft)
        main_layout.addLayout(grid)

        run_row = QHBoxLayout()
        run_row.addStretch(1)
        run_row.addWidget(self.run_btn)
        run_row.addWidget(self.unload_btn)
        run_row.addStretch(1)
        main_layout.addLayout(run_row)

        main_layout.addWidget(self.progress_bar)

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
        self.unload_btn.clicked.connect(self.unload_model)
        self.log_message.connect(self.append_log)
        self.progress_update.connect(self.update_progress)
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
        files = self._selected_files()
        if not files:
            MessageBox("Warning", "No PDF files selected.", self).exec()
            return

        # Check dependencies before running
        deps_ok, deps_msg = check_dependencies()
        if not deps_ok:
            MessageBox("Missing Dependencies", deps_msg, self).exec()
            return

        # Get device configuration
        device_str, dtype_str = self.device_combo.currentData()
        max_tokens = int(self.tokens_spin.value())

        self.log_box.clear()
        self.log_message.emit(f"Starting OCR for {len(files)} file(s)...")
        self.log_message.emit(f"Using device: {self.device_combo.currentText()}")

        self._set_controls_enabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Starting...")

        def worker():
            try:
                ok, fail, outputs = process_files(
                    files,
                    device_str,
                    dtype_str,
                    max_tokens,
                    log_emit=self.log_message.emit,
                    progress_emit=self.progress_update.emit,
                )
            except Exception as e:
                ok, fail, outputs = 0, len(files), []
                self.log_message.emit(f"Fatal error: {e}")
            self.processing_done.emit(ok, fail, outputs)

        threading.Thread(target=worker, daemon=True).start()

    def unload_model(self):
        """Unload model to free GPU memory."""
        OCRModel.unload()
        self.log_message.emit("Model unloaded. GPU memory freed.")

    def _set_controls_enabled(self, enabled: bool):
        self.run_btn.setEnabled(enabled)
        self.select_btn.setEnabled(enabled)
        self.device_combo.setEnabled(enabled)
        self.tokens_spin.setEnabled(enabled)
        self.unload_btn.setEnabled(enabled)

    def append_log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def update_progress(self, current: int, total: int):
        if total > 0:
            percent = int((current / total) * 100)
            self.progress_bar.setValue(percent)
            self.progress_bar.setFormat(f"Page {current}/{total}")

    def on_processing_done(self, ok: int, fail: int, outputs: list):
        if outputs:
            self.log_message.emit("Generated files:")
            for p in outputs:
                self.log_message.emit(f" - {p}")
        self.log_message.emit(f"Completed: {ok} succeeded, {fail} failed.")

        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("Done")
        self._set_controls_enabled(True)


def get_widget():
    return MainWidget()
