import os
import json
import threading
from pathlib import Path
from typing import List, Tuple, Callable
from datetime import datetime

# CRITICAL: Set before any imports
os.environ["CUDA_VISIBLE_DEVICES"] = ""
os.environ["CUDA_LAUNCH_BLOCKING"] = "1"

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog, QHBoxLayout, QVBoxLayout, QLabel,
    QTextEdit, QWidget, QSizePolicy, QComboBox
)
from qfluentwidgets import PrimaryPushButton, MessageBox

try:
    from chandra.model.schema import BatchInputItem
    from PIL import Image
    CHANDRA_AVAILABLE = True
except ImportError:
    CHANDRA_AVAILABLE = False


class MainWidget(QWidget):
    log_message = Signal(str)
    processing_done = Signal(int, int, str)

    def __init__(self):
        super().__init__()
        self.setObjectName("ocr_processor_widget")
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
        self.set_long_description(
            "Chandra OCR Processor - Extract text and structure from images and PDFs.\n"
            "Output will be saved as JSON files in the same directory as input files.\n\n"
            "⚠ Running in CPU-only mode (no GPU)."
        )

        self.mode_label = QLabel("Processing Mode:", self)
        self.mode_label.setStyleSheet("color: #dcdcdc; background: transparent; padding-left: 2px;")

        self.mode_selector = QComboBox(self)
        self.mode_selector.addItems([
            "Local (CPU Only) - Free",
            "API (Hosted) - Requires API key"
        ])
        self.mode_selector.setStyleSheet(
            "QComboBox{background: #2a2a2a; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px; padding: 5px;}"
        )

        self.select_btn = PrimaryPushButton("Select Files (Images/PDFs)", self)
        self.run_btn = PrimaryPushButton("Run OCR", self)

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

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        main_layout.addWidget(self.desc_label, 1)

        mode_layout = QHBoxLayout()
        mode_layout.addWidget(self.mode_label)
        mode_layout.addWidget(self.mode_selector, 1)
        main_layout.addLayout(mode_layout, 0)

        row1_layout = QHBoxLayout()
        row1_layout.addStretch(1)
        row1_layout.addWidget(self.select_btn, 1)
        row1_layout.addStretch(1)
        main_layout.addLayout(row1_layout, 0)

        row2_layout = QHBoxLayout()
        row2_layout.addStretch(1)
        row2_layout.addWidget(self.run_btn, 1)
        row2_layout.addStretch(1)
        main_layout.addLayout(row2_layout, 0)

        row3_layout = QHBoxLayout()
        row3_layout.addWidget(self.files_label, 1)
        row3_layout.addWidget(self.logs_label, 1)
        main_layout.addLayout(row3_layout, 0)

        row4_layout = QHBoxLayout()
        row4_layout.addWidget(self.files_box, 1)
        row4_layout.addWidget(self.log_box, 1)
        main_layout.addLayout(row4_layout, 4)

    def set_long_description(self, text: str):
        clean = (text or "").strip()
        if clean:
            self.desc_label.setText(clean)
            self.desc_label.show()
        else:
            self.desc_label.clear()
            self.desc_label.hide()

    def _connect_signals(self):
        self.select_btn.clicked.connect(self.select_files)
        self.run_btn.clicked.connect(self.run_process)
        self.log_message.connect(self.append_log)
        self.processing_done.connect(self.on_processing_done)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Images or PDFs",
            "",
            "Images and PDFs (*.png *.jpg *.jpeg *.pdf *.tiff *.bmp *.webp);;All Files (*)"
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
        if not CHANDRA_AVAILABLE:
            MessageBox(
                "Error",
                "Chandra OCR not installed.\n\nInstall with:\npip install chandra-ocr",
                self
            ).exec()
            return

        files = self._selected_files()
        if not files:
            MessageBox("Warning", "No files selected.", self).exec()
            return

        use_api = self.mode_selector.currentIndex() == 1

        self.log_box.clear()
        self.log_message.emit(f"Process starts - OCR processing {len(files)} file(s)")
        self.log_message.emit(f"Mode: {'API (Hosted)' if use_api else 'Local (CPU Only)'}")
        self.run_btn.setEnabled(False)
        self.select_btn.setEnabled(False)

        def worker():
            ok, fail = 0, 0
            try:
                ok, fail = process_files(files, self.log_message.emit, use_api)
            except Exception as e:
                self.log_message.emit(f"CRITICAL ERROR: {e}")
                import traceback
                self.log_message.emit(traceback.format_exc())
            self.processing_done.emit(ok, fail, "")

        threading.Thread(target=worker, daemon=True).start()

    def append_log(self, text: str):
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def on_processing_done(self, ok: int, fail: int, out_path: str):
        self.log_message.emit(f"\n{'='*50}")
        self.log_message.emit(f"Processing Complete!")
        self.log_message.emit(f"  Success: {ok}")
        self.log_message.emit(f"  Failed:  {fail}")
        self.log_message.emit(f"{'='*50}")
        self.run_btn.setEnabled(True)
        self.select_btn.setEnabled(True)


def process_files(file_paths: List[str], log_fn: Callable[[str], None], use_api: bool = False) -> Tuple[int, int]:
    if not CHANDRA_AVAILABLE:
        raise ImportError("Chandra OCR not available. Install with: pip install chandra-ocr")

    if use_api:
        log_fn("⚠ API mode selected but not yet implemented")
        log_fn("Please use Local mode or implement API integration")
        log_fn("See: https://www.datalab.to for API documentation")
        return 0, len(file_paths)

    log_fn("Initializing Chandra OCR model...")
    log_fn("Loading model with HuggingFace method (CPU-only mode)...")
    log_fn("\n⚠ If this is your first run, model download may take 5-10 minutes")
    log_fn("⚠ Model size: ~3 GB")
    log_fn("")

    try:
        import torch
        import torch.nn as nn

        log_fn("⚙ Forcing CPU-only mode...")

        # Patch torch.cuda
        class FakeCUDA:
            @staticmethod
            def is_available():
                return False
            @staticmethod
            def device_count():
                return 0

        torch.cuda = FakeCUDA()

        # Patch .cuda() methods
        torch.Tensor.cuda = lambda self, *args, **kwargs: self
        nn.Module.cuda = lambda self, *args, **kwargs: self

        log_fn("✓ CPU-only mode active")
        log_fn("")

        from chandra.model import InferenceManager

        log_fn("→ Loading model...")
        manager = InferenceManager(method="hf")
        log_fn("✓ Model loaded\n")

    except Exception as e:
        log_fn(f"✗ Failed to load model: {e}")
        raise

    success_count = 0
    failure_count = 0

    for idx, file_path in enumerate(file_paths, 1):
        try:
            log_fn(f"[{idx}/{len(file_paths)}] Processing: {os.path.basename(file_path)}")

            input_path = Path(file_path)

            if not input_path.exists():
                log_fn(f"  ✗ File not found: {file_path}")
                failure_count += 1
                continue

            output_path = input_path.parent / f"{input_path.stem}.json"

            if input_path.suffix.lower() == '.pdf':
                log_fn(f"  → Loading PDF...")
                try:
                    from pdf2image import convert_from_path
                    images = convert_from_path(str(input_path), dpi=300)
                    log_fn(f"  → Found {len(images)} page(s)")
                except ImportError:
                    log_fn("  ✗ pdf2image not installed.")
                    log_fn("     Install with: pip install pdf2image")
                    failure_count += 1
                    continue
                except Exception as e:
                    log_fn(f"  ✗ Failed to load PDF: {e}")
                    failure_count += 1
                    continue
            else:
                log_fn(f"  → Loading image...")
                try:
                    images = [Image.open(str(input_path))]
                except Exception as e:
                    log_fn(f"  ✗ Failed to load image: {e}")
                    failure_count += 1
                    continue

            results = []
            for page_num, image in enumerate(images, 1):
                log_fn(f"  → Processing page {page_num}/{len(images)}...")

                batch = [BatchInputItem(image=image, prompt_type="ocr_layout")]
                result = manager.generate(batch)[0]

                page_data = {
                    "page": page_num,
                    "markdown": result.markdown,
                }

                if hasattr(result, 'json') and result.json:
                    try:
                        page_data["structured_data"] = json.loads(result.json) if isinstance(result.json, str) else result.json
                    except:
                        page_data["raw_json"] = result.json

                results.append(page_data)

            output_data = {
                "source_file": str(input_path.absolute()),
                "filename": input_path.name,
                "processed_at": datetime.now().isoformat(),
                "total_pages": len(images),
                "ocr_results": results
            }

            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(output_data, f, indent=2, ensure_ascii=False)

            log_fn(f"  ✓ Saved to: {output_path.name}")
            success_count += 1

        except Exception as e:
            log_fn(f"  ✗ Failed: {str(e)}")
            failure_count += 1

        log_fn("")

    return success_count, failure_count


def get_widget():
    return MainWidget()