import os
import sys
import re
import importlib.util
from importlib import metadata

from PySide6.QtCore import Qt, QFileSystemWatcher, QTimer, Signal, QProcess
from PySide6.QtWidgets import *
from PySide6.QtGui import QIcon
from qfluentwidgets import PrimaryPushButton, isDarkTheme


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
        self._grid.setContentsMargins(12, 12, 12, 12)
        self._grid.setSpacing(12)
        self.setWidget(self._content)

        # State
        self._cards: list[PycroCard] = []
        self._watcher = QFileSystemWatcher(self)
        self._watcher.directoryChanged.connect(self._on_dir_changed)
        self._watcher.fileChanged.connect(self._on_file_changed)
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.setInterval(250)
        self._debounce_timer.timeout.connect(self.refresh)

        # Initial content
        self._empty_label = QLabel("Pycros will appear here", self._content)
        self._empty_label.setAlignment(Qt.AlignCenter)
        self._empty_label.setStyleSheet("color: #aaa; font-size: 14px;")
        self._grid.addWidget(self._empty_label, 0, 0)
        self._grid.setRowStretch(0, 1)
        self._grid.setColumnStretch(0, 1)

        # Root folder for pycros
        self._root = os.path.join(os.getcwd(), 'pycros')
        if os.path.isdir(self._root):
            self._watcher.addPath(self._root)
        # Build initially
        QTimer.singleShot(0, self.refresh)

    def gridLayout(self) -> QGridLayout:
        """Expose the underlying grid layout for adding Pycro widgets later."""
        return self._grid

    # --- Public API ---
    def refresh(self):
        """Rescan pycros folder and rebuild cards."""
        infos = self._scan_pycros()
        self._rebuild(infos)

    # --- Internal helpers ---
    def _scan_pycros(self):
        root = self._root
        infos = []
        if not os.path.isdir(root):
            return infos
        try:
            subdirs = sorted([d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))])
        except Exception:
            subdirs = []

        # Reset watcher to follow files within subdirs
        try:
            self._watcher.removePaths(self._watcher.directories())
            self._watcher.removePaths(self._watcher.files())
        except Exception:
            pass
        self._watcher.addPath(root)

        for d in subdirs:
            folder = os.path.join(root, d)
            main_py = os.path.join(folder, 'main.py')
            req_txt = os.path.join(folder, 'requirements.txt')
            desc_md = os.path.join(folder, 'description.md')
            # determine if folder contains any .py file
            has_python = False
            try:
                for fn in os.listdir(folder):
                    if fn.lower().endswith('.py'):
                        has_python = True
                        break
            except Exception:
                has_python = False

            short_desc, long_desc = self._parse_description(desc_md)
            display_name = d.replace('--', ' ')

            info = PycroInfo(
                name=d,
                display_name=display_name,
                folder=folder,
                main_py=main_py,
                requirements=req_txt if os.path.isfile(req_txt) else None,
                description=desc_md if os.path.isfile(desc_md) else None,
                short_desc=short_desc,
                long_desc=long_desc,
                has_python=has_python
            )

            infos.append(info)

            # Watch description and requirements for changes
            if os.path.isfile(desc_md):
                self._watcher.addPath(desc_md)
            if os.path.isfile(req_txt):
                self._watcher.addPath(req_txt)

        return infos

    def _parse_description(self, path):
        short_desc = ''
        long_desc_lines = []
        if not os.path.isfile(path):
            return short_desc, ''
        try:
            with open(path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            for line in lines:
                s = line.strip('\n')
                if s.startswith('>') and not short_desc:
                    short_desc = s.lstrip('> ').strip()
                elif not s.startswith('>'):
                    long_desc_lines.append(s)
        except Exception:
            pass
        long_desc = '\n'.join(long_desc_lines).strip()
        return short_desc, long_desc

    def _rebuild(self, infos):
        # clear existing
        for card in self._cards:
            self._grid.removeWidget(card)
            card.deleteLater()
        self._cards.clear()

        if not infos:
            self._empty_label.show()
            self._grid.setRowStretch(0, 1)
            self._grid.setColumnStretch(0, 1)
            return
        else:
            self._empty_label.hide()

        # Fixed column count for simplicity
        columns = 3
        row = 0
        col = 0
        for info in infos:
            card = PycroCard(info, parent=self)
            self._cards.append(card)
            self._grid.addWidget(card, row, col)
            col += 1
            if col >= columns:
                col = 0
                row += 1

        # Stretch last row/col
        for r in range(row + 1):
            self._grid.setRowStretch(r, 0)
        self._grid.setRowStretch(row + 1, 1)
        for c in range(columns):
            self._grid.setColumnStretch(c, 0)
        self._grid.setColumnStretch(columns, 1)

    def _on_dir_changed(self, _):
        self._debounce_timer.start()

    def _on_file_changed(self, _):
        self._debounce_timer.start()


class PycroInfo:
    def __init__(self, name, display_name, folder, main_py, requirements, description, short_desc, long_desc, has_python: bool):
        self.name = name
        self.display_name = display_name
        self.folder = folder
        self.main_py = main_py
        self.requirements = requirements
        self.description = description
        self.short_desc = short_desc
        self.long_desc = long_desc
        self.has_python = has_python


class PycroCard(QWidget):
    def __init__(self, info: 'PycroInfo', parent=None):
        super().__init__(parent)
        self.info = info
        self._grid: PycroGrid = parent

        # sanitize object name for stylesheet id selector
        safe_name = re.sub(r'[^A-Za-z0-9_]', '_', info.name)
        self.setObjectName(f"card__{safe_name}")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setAutoFillBackground(True)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.setMinimumSize(280, 210)
        self.setMaximumSize(320, 250)
        # Background color similar to sidebar; target only this widget to avoid affecting children
        bg = "#242424" if isDarkTheme() else "#F2F2F2"
        border = "rgba(255,255,255,0.08)" if isDarkTheme() else "rgba(0,0,0,0.08)"
        self.setStyleSheet(
            f"QWidget#{self.objectName()}{{background-color:{bg}; border:1px solid {border}; border-radius:8px;}}"
        )

        v = QVBoxLayout(self)
        v.setContentsMargins(12, 12, 12, 12)
        v.setSpacing(8)

        title = QLabel(info.display_name, self)
        title.setWordWrap(True)
        title.setStyleSheet(
            f"background: transparent; border: none; color: {'#fff' if isDarkTheme() else '#111'}; font-size:16px; font-weight:600;"
        )

        desc = QLabel(info.short_desc or '(no description)', self)
        desc.setWordWrap(True)
        desc.setStyleSheet(
            f"background: transparent; border: none; color: {'#bbb' if isDarkTheme() else '#444'}; font-size:12px;"
        )

        v.addWidget(title)
        v.addWidget(desc)
        v.addStretch(1)

        # Buttons row
        h = QHBoxLayout()
        self.launch_btn = PrimaryPushButton('Launch', self)
        # Keep Launch button a constant size across states/cards
        self.launch_btn.setFixedHeight(28)
        self.launch_btn.setFixedWidth(90)
        # Make install button same style as Launch (primary) when clickable
        self.install_btn = PrimaryPushButton('Install Requirements', self)
        self.install_btn.setFixedHeight(28)
        self.install_btn.setCursor(Qt.PointingHandCursor)
        self.launch_btn.setCursor(Qt.PointingHandCursor)
        self.launch_btn.clicked.connect(self._on_launch)
        self.install_btn.clicked.connect(self._on_install)
        h.addWidget(self.launch_btn)
        h.addWidget(self.install_btn)
        v.addLayout(h)

        if not self.info.has_python:
            self.launch_btn.hide()
            self.install_btn.hide()
        else:
            self._update_requirements_state()

    def _req_packages(self):
        path = self.info.requirements
        if not path or not os.path.isfile(path):
            return []
        names = []
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    s = line.strip()
                    if not s or s.startswith('#'):
                        continue
                    # very simple name parse (before version specifiers / extras)
                    m = re.match(r"^[A-Za-z0-9_.\-]+", s)
                    if m:
                        names.append(m.group(0))
        except Exception:
            pass
        return names

    def _are_requirements_satisfied(self) -> bool:
        pkgs = self._req_packages()
        if not pkgs:
            return True
        for name in pkgs:
            try:
                metadata.version(name)
            except metadata.PackageNotFoundError:
                return False
            except Exception:
                # if metadata lookup fails, attempt import as fallback
                try:
                    __import__(name.replace('-', '_'))
                except Exception:
                    return False
        return True

    def _update_requirements_state(self, installing: bool = False):
        ok = self._are_requirements_satisfied()
        if installing:
            # Keep Primary style; just disable and change text
            self.install_btn.setEnabled(False)
            self.install_btn.setText('Installing...')
            self.launch_btn.setEnabled(False)
            return

        if ok:
            # Green, non-clickable
            self.install_btn.setEnabled(False)
            self.install_btn.setText('Requirements OK')
            self.install_btn.setStyleSheet(
                "QPushButton:disabled { background-color: #21a366; color: white; border: none; padding: 6px 12px; border-radius:6px;}"
            )
            self.launch_btn.setEnabled(True)
        else:
            # Blue, clickable (same Primary style as Launch)
            self.install_btn.setEnabled(True)
            self.install_btn.setText('Install Requirements')
            self.launch_btn.setEnabled(False)

    def _on_install(self):
        if not self.info.requirements or not os.path.isfile(self.info.requirements):
            return
        self._update_requirements_state(installing=True)

        # Run pip install in background
        self._proc = QProcess(self)
        self._proc.setProgram(sys.executable)
        self._proc.setArguments(['-m', 'pip', 'install', '-r', self.info.requirements])
        self._proc.finished.connect(lambda *_: self._on_install_finished())
        self._proc.start()

    def _on_install_finished(self):
        self._proc = None
        # Re-evaluate requirements and update buttons
        self._update_requirements_state(installing=False)

    def _on_launch(self):
        # Call parent window to add a macro tab
        window = self.window()
        widget = self._load_widget()
        if widget is None:
            QMessageBox.warning(self, 'Launch failed', 'Could not load macro widget (missing MainWidget/get_widget)')
            return
        # Add a wrapper page with long description + widget
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(12)

        if self.info.long_desc:
            desc = QTextBrowser()
            desc.setOpenExternalLinks(True)
            desc.setPlainText(self.info.long_desc)
            desc.setStyleSheet('QTextBrowser{background:transparent; color:#ddd; border:none;}')
            desc.setFixedHeight(120)
            v.addWidget(desc)

        v.addWidget(widget, 1)

        try:
            window.addMacroTab(self.info.name, self.info.display_name, QIcon(), page)
        except Exception:
            # Fallback: show as separate window
            widget.setWindowTitle(self.info.display_name)
            widget.resize(800, 600)
            widget.show()

    def _load_widget(self) -> QWidget | None:
        main_path = self.info.main_py
        if not os.path.isfile(main_path):
            return None
        try:
            spec = importlib.util.spec_from_file_location(f"pycro_{self.info.name}", main_path)
            if spec is None or spec.loader is None:
                return None
            module = importlib.util.module_from_spec(spec)
            sys.modules[spec.name] = module
            spec.loader.exec_module(module)  # type: ignore
            # Prefer get_widget()
            if hasattr(module, 'get_widget') and callable(module.get_widget):
                w = module.get_widget()
                if isinstance(w, QWidget):
                    return w
            # Else try MainWidget class
            if hasattr(module, 'MainWidget'):
                cls = getattr(module, 'MainWidget')
                try:
                    inst = cls()
                    if isinstance(inst, QWidget):
                        return inst
                except Exception:
                    return None
        except Exception:
            return None
        return None
