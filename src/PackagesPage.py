import os
import re
import sys
from importlib import metadata

from app_paths import REQUIREMENTS_TXT
from PySide6.QtCore import *
from PySide6.QtGui import *
from PySide6.QtWidgets import *
from qfluentwidgets import TransparentToolButton, PrimaryPushButton, isDarkTheme
from pytablericons import TablerIcons, OutlineIcon, FilledIcon
from PIL.ImageQt import ImageQt


def ti_icon(icon_enum, size=24, color="#FFFFFF", stroke_width=2.0) -> QIcon:
    img = TablerIcons.load(icon_enum, size=size, color=color, stroke_width=stroke_width)
    return QIcon(QPixmap.fromImage(ImageQt(img)))


# Packages to hide from the list (exact names provided)
HIDDEN_PACKAGES_EXACT: set[str] = {
    'pywin32',
    'autocommand',
    'backports.tarfile',
    'darkdetect',
    'importlib_metadata',
    'inflect',
    'jaraco.collections',
    'jaraco.context',
    'jaraco.functools',
    'jaraco.text',
    'more-itertools',
    'packaging',
    'pillow',
    'pip',
    'platformdirs',
    'pygame',
    'PySide6_Addons',
    'PySide6_Essentials',
    'PySideSix-Frameless-Window',
    'setuptools',
    'shiboken6',
    'tomli',
    'typeguard',
    'typing_extensions',
    'wheel',
    'zipp',
}
# Also compare in lower-case to avoid case-only mismatches while keeping exact-name intent
_HIDDEN_PACKAGES_LOWER = {s.lower() for s in HIDDEN_PACKAGES_EXACT}


class CheckIconButton(TransparentToolButton):
    toggledManually = Signal(bool)

    def __init__(self, parent=None, initially_checked=False):
        super().__init__(parent)
        self._checked = initially_checked
        self._partial = False
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedSize(32, 32)
        self.setIconSize(QSize(24, 24))
        self._sync_icon()
        self.clicked.connect(self._on_clicked)

    def setChecked(self, checked: bool):
        self._checked = bool(checked)
        self._partial = False
        self._sync_icon()

    def setPartial(self, partial: bool):
        self._partial = bool(partial)
        self._sync_icon()

    def isChecked(self):
        return self._checked

    def _on_clicked(self):
        if self._partial:
            # from partial, select all
            self._partial = False
            self._checked = True
        else:
            self._checked = not self._checked
        self._sync_icon()
        self.toggledManually.emit(self._checked)

    def _sync_icon(self):
        if self._partial:
            self.setIcon(ti_icon(OutlineIcon.SQUARE_OFF))
        else:
            self.setIcon(ti_icon(FilledIcon.SQUARE) if self._checked else ti_icon(OutlineIcon.SQUARE))


class PackageRow(QWidget):
    toggled = Signal(str, bool)
    removeClicked = Signal(str)

    def __init__(self, name: str, parent=None):
        super().__init__(parent)
        self.name = name
        h = QHBoxLayout(self)
        h.setContentsMargins(8, 8, 8, 8)
        h.setSpacing(10)

        self.check = CheckIconButton(self, initially_checked=False)
        self.label = QLabel(name, self)
        self.label.setStyleSheet('color: #ddd;')
        self.remove_btn = TransparentToolButton(self)
        self.remove_btn.setIcon(ti_icon(OutlineIcon.TRASH))
        self.remove_btn.setFixedSize(32, 32)
        self.remove_btn.setIconSize(QSize(24, 24))
        self.remove_btn.setCursor(Qt.PointingHandCursor)
        self.remove_btn.setToolTip('Uninstall')

        self.check.toggledManually.connect(lambda c: self.toggled.emit(self.name, c))
        self.remove_btn.clicked.connect(lambda: self.removeClicked.emit(self.name))

        h.addWidget(self.check, 0, Qt.AlignVCenter)
        h.addWidget(self.label)
        h.addStretch(1)
        h.addWidget(self.remove_btn, 0, Qt.AlignVCenter)

        self.setLayout(h)
        # sanitize object name for stylesheet id selector
        safe = re.sub(r'[^A-Za-z0-9_]', '_', name)
        self.setObjectName(f'pkg_row__{safe}')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setAutoFillBackground(True)
        # Background will be applied by page (alternating)
        self.apply_background('#000')

    def setSelected(self, selected: bool):
        self.check.setChecked(selected)

    def apply_background(self, color: str):
        border = "rgba(255,255,255,0.06)" if isDarkTheme() else "rgba(0,0,0,0.06)"
        self.setStyleSheet(
            f"""
            QWidget#{self.objectName()} {{
                background-color: {color};
                border: 1px solid {border};
                border-radius: 6px;
            }}
            """
        )


class PackagesPage(QWidget):
    packagesChanged = Signal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('packagesInterface')

        self._baseline = self._read_baseline_requirements()
        self._selected: set[str] = set()
        self._rows: dict[str, PackageRow] = {}
        self._proc: QProcess | None = None

        v = QVBoxLayout(self)
        v.setContentsMargins(12, 12, 12, 12)
        v.setSpacing(12)

        # Toolbar
        toolbar = QHBoxLayout()
        toolbar.setSpacing(10)

        self.select_all_btn = CheckIconButton(self, initially_checked=False)
        self.select_all_btn.setToolTip('Select All')
        self.select_all_btn.toggledManually.connect(self._on_select_all_clicked)

        self.trash_selected_btn = TransparentToolButton(self)
        self.trash_selected_btn.setIcon(ti_icon(OutlineIcon.TRASH))
        self.trash_selected_btn.setFixedSize(32, 32)
        self.trash_selected_btn.setIconSize(QSize(24, 24))
        self.trash_selected_btn.setToolTip('Uninstall selected')
        self.trash_selected_btn.clicked.connect(self._remove_selected)

        self.invert_btn = PrimaryPushButton('Invert Selections', self)
        self.invert_btn.clicked.connect(self._invert_selections)

        toolbar.addWidget(self.select_all_btn, 0, Qt.AlignVCenter)
        toolbar.addWidget(self.trash_selected_btn, 0, Qt.AlignVCenter)
        toolbar.addSpacing(8)
        toolbar.addWidget(self.invert_btn)
        toolbar.addStretch(1)

        v.addLayout(toolbar)

        # Lock notice when packages adjustments are disabled
        self.lock_label = QLabel('Close all tabs to manage packages.', self)
        self.lock_label.setStyleSheet('color:#d9534f;')
        self.lock_label.setVisible(False)
        v.addWidget(self.lock_label)

        # List area
        self.scroll = QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setStyleSheet('QScrollArea{background:transparent;border:none;}')
        self.viewport = QWidget(self.scroll)
        self.viewport.setStyleSheet('background:transparent;')
        self.listLayout = QVBoxLayout(self.viewport)
        self.listLayout.setContentsMargins(0, 0, 0, 0)
        self.listLayout.setSpacing(8)
        self.scroll.setWidget(self.viewport)
        v.addWidget(self.scroll, 1)

        self.setLayout(v)
        QTimer.singleShot(0, self.refresh)

    def _clear_list(self):
        # Remove all items (rows, empty label, spacers) from the list layout
        layout = self.listLayout
        while layout.count():
            item = layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

    def _read_baseline_requirements(self) -> set[str]:
        path = os.fspath(REQUIREMENTS_TXT)
        names: set[str] = set()
        if not os.path.isfile(path):
            return names
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    s = line.strip()
                    if not s or s.startswith('#'):
                        continue
                    m = re.match(r"^[A-Za-z0-9_.\-]+", s)
                    if m:
                        names.add(m.group(0).lower())
        except Exception:
            pass
        return names

    def _installed_packages(self) -> list[str]:
        names = []
        try:
            for d in metadata.distributions():
                name = d.metadata.get('Name') or d.metadata.get('Summary')
                if not name:
                    continue
                names.append(str(name))
        except Exception:
            pass
        # Normalize and unique
        norm = []
        seen = set()
        for n in names:
            l = n.lower()
            if l not in seen:
                seen.add(l)
                norm.append(n)
        return sorted(norm, key=lambda s: s.lower())

    def _extra_packages(self) -> list[str]:
        extras = []
        baseline = self._baseline
        for n in self._installed_packages():
            # hide by exact or case-insensitive match
            if n in HIDDEN_PACKAGES_EXACT or n.lower() in _HIDDEN_PACKAGES_LOWER:
                continue
            if n.lower() not in baseline:
                extras.append(n)
        return extras

    def refresh(self):
        # clear any previous content including empty labels/spacers
        self._clear_list()
        self._rows.clear()
        self._selected.clear()

        pkgs = self._extra_packages()
        if not pkgs:
            empty = QLabel('No extra packages found. 🎉', self.viewport)
            empty.setStyleSheet('color:#aaa;')
            self.listLayout.addWidget(empty)
        else:
            for idx, n in enumerate(pkgs):
                row = PackageRow(n, self.viewport)
                row.toggled.connect(self._on_row_toggled)
                row.removeClicked.connect(self._on_row_remove)
                self._rows[n] = row
                self.listLayout.addWidget(row)
                # alternating background similar to sidebar
                base = "#242424" if isDarkTheme() else "#F2F2F2"
                alt = "#2A2A2A" if isDarkTheme() else "#EDEDED"
                row.apply_background(base if idx % 2 == 0 else alt)
            self.listLayout.addStretch(1)
        self._update_select_all_state()

    def _on_row_toggled(self, name: str, checked: bool):
        if checked:
            self._selected.add(name)
        else:
            self._selected.discard(name)
        self._update_select_all_state()

    def _on_row_remove(self, name: str):
        self._uninstall_packages([name])

    def _on_select_all_clicked(self, checked: bool):
        # If tri-state partial -> toggled to full selected in button handler
        if checked:
            for row in self._rows.values():
                row.setSelected(True)
                self._selected.add(row.name)
        else:
            for row in self._rows.values():
                row.setSelected(False)
            self._selected.clear()
        self._update_select_all_state()

    def _invert_selections(self):
        self._selected = {name for name, row in self._rows.items() if not row.check.isChecked()}
        for name, row in self._rows.items():
            row.setSelected(name in self._selected)
        self._update_select_all_state()

    def _update_select_all_state(self):
        total = len(self._rows)
        sel = len(self._selected)
        if sel == 0:
            self.select_all_btn.setPartial(False)
            self.select_all_btn.setChecked(False)
        elif sel == total:
            self.select_all_btn.setPartial(False)
            self.select_all_btn.setChecked(True)
        else:
            # partial
            self.select_all_btn.setPartial(True)

    def _remove_selected(self):
        if not self._selected:
            return
        self._uninstall_packages(sorted(self._selected))

    def _uninstall_packages(self, names: list[str]):
        if not names:
            return
        if self._proc is not None:
            return
        self._set_enabled(False)
        self._proc = QProcess(self)
        self._proc.setProgram(sys.executable)
        self._proc.setArguments(['-m', 'pip', 'uninstall', '-y', *names])
        self._proc.finished.connect(self._on_uninstall_finished)
        self._proc.start()

    def _on_uninstall_finished(self):
        self._proc = None
        self._set_enabled(True)
        self.refresh()
        # notify listeners (Hub) to revalidate requirements state
        try:
            self.packagesChanged.emit()
        except Exception:
            pass

    def _set_enabled(self, enabled: bool):
        self.select_all_btn.setEnabled(enabled)
        self.trash_selected_btn.setEnabled(enabled)
        self.invert_btn.setEnabled(enabled)
        for row in self._rows.values():
            row.setEnabled(enabled)

    def setLocked(self, locked: bool):
        """Disable all adjustments when locked is True."""
        self.lock_label.setVisible(locked)
        self._set_enabled(not locked)
