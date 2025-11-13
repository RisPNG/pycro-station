from PySide6.QtCore import Qt
from PySide6.QtWidgets import *


class PycroGrid(QScrollArea):
    """Scrollable grid container for listing Pycros.

    This replaces the text editor area with a frame suitable
    for displaying macro entries in a grid-like layout.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        # Make the grid area visually match Settings by being transparent
        self.setStyleSheet("QScrollArea{background: transparent; border: none;}")
        self.viewport().setStyleSheet("background: transparent;")

        # Content widget with a grid layout
        self._content = QWidget(self)
        self._content.setStyleSheet("background: transparent;")
        self._grid = QGridLayout(self._content)
        self._grid.setContentsMargins(0, 0, 0, 0)
        self._grid.setSpacing(0)
        self.setWidget(self._content)

        # Placeholder empty-state label
        placeholder = QLabel("Pycros will appear here", self._content)
        placeholder.setAlignment(Qt.AlignCenter)
        placeholder.setStyleSheet("color: #aaa; font-size: 14px;")

        # Use a simple container to center the placeholder within the grid
        container = QWidget(self._content)
        v = QVBoxLayout(container)
        v.addStretch(1)
        v.addWidget(placeholder, alignment=Qt.AlignCenter)
        v.addStretch(1)

        self._grid.addWidget(container, 0, 0)
        # Ensure the single cell expands to fill available space
        self._grid.setRowStretch(0, 1)
        self._grid.setColumnStretch(0, 1)

    def gridLayout(self) -> QGridLayout:
        """Expose the underlying grid layout for adding Pycro widgets later."""
        return self._grid
