from __future__ import annotations

import importlib.util
import sys
import threading
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import MessageBox, PrimaryPushButton


processor_name = "sales_forecast_reconcile_processor"
processor_spec = importlib.util.spec_from_file_location(
    processor_name, Path(__file__).with_name("processor.py")
)
if processor_spec is None or processor_spec.loader is None:
    raise ImportError("Could not load the Sales Forecast reconciliation processor")
processor_module = importlib.util.module_from_spec(processor_spec)
sys.modules[processor_name] = processor_module
processor_spec.loader.exec_module(processor_module)
process_reconciliation = processor_module.process_reconciliation


SOURCE_LABELS = {
    "sales": "Sales Report",
    "fotp": "FOTP by Season",
    "shipment": "Shipment Forecast (Ann)",
    "bds": "Previous-Week BDS",
}


class MainWidget(QWidget):
    log_message = Signal(str)
    processing_done = Signal(bool, str, str)

    def __init__(self):
        super().__init__()
        self.setObjectName("sales_forecast_reconcile_widget")
        self._paths: dict[str, QLineEdit] = {}
        self._browse_buttons: list[PrimaryPushButton] = []
        self._build_ui()
        self.run_btn.clicked.connect(self.run_process)
        self.log_message.connect(self._append_log)
        self.processing_done.connect(self._on_processing_done)

    def _build_ui(self):
        self.desc_label = QLabel("", self)
        self.desc_label.setWordWrap(True)
        self.desc_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.desc_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.desc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.set_long_description("")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)
        main_layout.addWidget(self.desc_label)

        main_layout.addWidget(QLabel("Source workbooks", self))

        source_grid = QGridLayout()
        source_grid.setHorizontalSpacing(10)
        source_grid.setVerticalSpacing(8)
        for row, (role, label) in enumerate(SOURCE_LABELS.items()):
            field = QLineEdit(self)
            field.setReadOnly(True)
            field.setPlaceholderText("Select an .xlsx workbook")
            button = PrimaryPushButton("Browse", self)
            button.clicked.connect(lambda _checked=False, r=role: self._browse_source(r))
            self._paths[role] = field
            self._browse_buttons.append(button)
            source_grid.addWidget(QLabel(label, self), row, 0)
            source_grid.addWidget(field, row, 1)
            source_grid.addWidget(button, row, 2)
        source_grid.setColumnStretch(1, 1)
        main_layout.addLayout(source_grid)

        main_layout.addWidget(QLabel("Report settings", self))

        output_grid = QGridLayout()
        output_grid.setHorizontalSpacing(10)
        output_grid.setVerticalSpacing(8)
        self.output_field = QLineEdit(self)
        self.output_field.setReadOnly(True)
        self.output_field.setPlaceholderText("Defaults to the Sales Report folder")
        self.output_button = PrimaryPushButton("Browse", self)
        self.output_button.clicked.connect(self._browse_output_folder)
        output_grid.addWidget(QLabel("Output folder", self), 0, 0)
        output_grid.addWidget(self.output_field, 0, 1)
        output_grid.addWidget(self.output_button, 0, 2)
        output_grid.setColumnStretch(1, 1)
        main_layout.addLayout(output_grid)

        action_row = QHBoxLayout()
        action_row.addStretch(1)
        self.run_btn = PrimaryPushButton("Build Reconciliation", self)
        self.run_btn.setMinimumWidth(240)
        action_row.addWidget(self.run_btn)
        action_row.addStretch(1)
        main_layout.addLayout(action_row)

        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Process log")
        main_layout.addWidget(self.log_box, 1)

    def set_long_description(self, text: str):
        clean = (text or "").strip()
        if clean:
            self.desc_label.setText(clean)
            self.desc_label.show()
        else:
            self.desc_label.clear()
            self.desc_label.hide()

    def _browse_source(self, role: str):
        path, _ = QFileDialog.getOpenFileName(
            self,
            f"Select {SOURCE_LABELS[role]}",
            "",
            "Excel workbooks (*.xlsx);;All files (*)",
        )
        if path:
            self._paths[role].setText(path)
            if role == "sales" and not self.output_field.text().strip():
                self.output_field.setText(str(Path(path).parent))

    def _browse_output_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Select output folder")
        if path:
            self.output_field.setText(path)

    def run_process(self):
        paths = {role: field.text().strip() for role, field in self._paths.items()}
        missing = [SOURCE_LABELS[role] for role, path in paths.items() if not path]
        if missing:
            MessageBox("Missing files", "Select: " + ", ".join(missing), self).exec()
            return

        output_dir = self.output_field.text().strip() or str(Path(paths["sales"]).parent)
        self.log_box.clear()
        self.run_btn.setEnabled(False)
        self.output_button.setEnabled(False)
        for button in self._browse_buttons:
            button.setEnabled(False)
        self.log_message.emit("Process starts")

        def worker():
            try:
                result = process_reconciliation(
                    paths["sales"],
                    paths["fotp"],
                    paths["shipment"],
                    paths["bds"],
                    output_dir,
                    self.log_message.emit,
                )
                self.processing_done.emit(True, str(result), "")
            except Exception as exc:
                self.log_message.emit(f"ERROR: {exc}")
                self.processing_done.emit(False, "", str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _append_log(self, message: str):
        self.log_box.append(message)
        self.log_box.ensureCursorVisible()

    def _on_processing_done(self, ok: bool, output_path: str, error: str):
        self.run_btn.setEnabled(True)
        self.output_button.setEnabled(True)
        for button in self._browse_buttons:
            button.setEnabled(True)
        if ok:
            self.log_message.emit(f"Output workbook saved to: {output_path}")
            box = MessageBox("Processing complete", f"Created:\n{output_path}", self)
        else:
            box = MessageBox("Processing failed", error or "See the process log.", self)
        box.yesButton.setText("OK")
        box.cancelButton.hide()
        box.exec()


def get_widget():
    return MainWidget()
