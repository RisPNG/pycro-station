from PySide6.QtWidgets import QWidget, QVBoxLayout, QTextEdit, QFileDialog
from qfluentwidgets import PrimaryPushButton, PushButton


class MainWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName('test_macro_one_widget')

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        self.select_btn = PrimaryPushButton('Select Files...', self)
        self.select_btn.clicked.connect(self.select_files)

        self.files_box = QTextEdit(self)
        self.files_box.setReadOnly(True)
        self.files_box.setPlaceholderText('Selected files will appear here')
        self.files_box.setStyleSheet('QTextEdit{background: #2a2a2a; color: white; border: 1px solid #3a3a3a; border-radius: 6px;}')

        self.run_btn = PrimaryPushButton('Run', self)
        # Currently does nothing; placeholder
        self.run_btn.clicked.connect(lambda: None)

        layout.addWidget(self.select_btn)
        layout.addWidget(self.files_box, 1)
        layout.addWidget(self.run_btn)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, 'Select files')
        if files:
            self.files_box.setPlainText('\n'.join(files))


def get_widget():
    return MainWidget()
