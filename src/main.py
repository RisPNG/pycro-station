"""
The main python file. Run this file to use the app.
"""
APP_VERSION = "1.6.0.0"
SHOW_REPO_FIELDS_IN_SETTINGS = False
import datetime
import json
import os
import re
import shutil
import sys
import tempfile
import time
import urllib.error
import urllib.request
import zipfile
from tkinter import filedialog

from PySide6.QtCore import *
from PySide6.QtGui import *
from PySide6.QtWidgets import *
from qfluentwidgets import *
# https://fluenticons.co
from qfluentwidgets import FluentIcon as FIF
# https://tabler.io/icons
from pytablericons import TablerIcons, OutlineIcon
from PIL.ImageQt import ImageQt

from PycroGrid import PycroGrid
from PackagesPage import PackagesPage, CheckIconButton
from TitleBar import CustomTitleBar


class AnimatedStackedWidget(QStackedWidget):
    """QStackedWidget with fly-up (load) transitions."""

    def __init__(self, parent=None):
        super().__init__(parent)
        # Match the MSFluentWindow default "fly up" feel for consistency
        self._duration = 300  # milliseconds
        self._offset = 76
        self._slideAnimation = None

    def _stop_animations(self):
        anim = getattr(self, "_slideAnimation", None)
        if anim is None:
            return
        try:
            anim.stop()
        except Exception:
            pass
        try:
            anim.deleteLater()
        except Exception:
            pass
        self._slideAnimation = None

    def setCurrentWidgetNoAnimation(self, widget):
        """Switch immediately (no animation)."""
        if widget is None:
            return
        if self.indexOf(widget) == -1:
            self.addWidget(widget)
        self._stop_animations()
        super().setCurrentWidget(widget)
        try:
            widget.move(0, 0)
        except Exception:
            pass

    def setCurrentWidget(self, widget):
        """Switch to widget with fly-up (load) animation."""
        if widget is None:
            return
        current = self.currentWidget()
        if current is widget:
            return

        # Ensure widget is in the stack
        if self.indexOf(widget) == -1:
            self.addWidget(widget)

        # Skip animation if not visible (e.g. selecting a tab while sidebar is active)
        if not self.isVisible() or self.width() <= 0 or self.height() <= 0:
            self.setCurrentWidgetNoAnimation(widget)
            return

        self._stop_animations()

        super().setCurrentWidget(widget)

        # Fly-up the new widget (load)
        try:
            widget.move(0, self._offset)
        except Exception:
            pass

        self._slideAnimation = QPropertyAnimation(widget, b"pos", self)
        self._slideAnimation.setDuration(self._duration)
        self._slideAnimation.setStartValue(QPoint(0, self._offset))
        self._slideAnimation.setEndValue(QPoint(0, 0))
        self._slideAnimation.setEasingCurve(QEasingCurve.OutQuad)
        slide_anim = self._slideAnimation

        def _cleanup_slide():
            try:
                slide_anim.deleteLater()
            except Exception:
                pass
            if self._slideAnimation is slide_anim:
                self._slideAnimation = None

        self._slideAnimation.finished.connect(_cleanup_slide)
        self._slideAnimation.start()

    def setCurrentIndex(self, index):
        """Switch to index with fly-up (load) animation."""
        widget = self.widget(index)
        if widget:
            self.setCurrentWidget(widget)

class Settings(QWidget):
    updateFinished = Signal(bool, str)
    appUpdateFinished = Signal(bool, str)
    appUpdateAvailable = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings_file = os.path.join(os.path.dirname(__file__), "settings.json")
        self._show_update_dialog = True
        self._launch_update_started = False
        self._show_app_update_dialog = True
        self._launch_app_update_started = False

        # Track editing state for each field
        self.editing_states = {
            "repo_url": False,
            "repo_branch": False,
            "repo_directory": False,
            "app_repo_url": False,
            "app_repo_branch": False,
            "app_repo_directory": False,
        }

        # Match the BoM--to--MSL log/description styling
        self.field_disabled_style = (
            "QLineEdit {background: #1f1f1f; color: #d0d0d0; "
            "border: 1px solid #3a3a3a; border-radius: 6px; padding: 6px;}"
            "QLineEdit:disabled {background: #1f1f1f; color: #d0d0d0;}"
        )
        self.field_enabled_style = (
            "QLineEdit {background: transparent; color: #dcdcdc; "
            "border: 1px solid #3a3a3a; border-radius: 6px; padding: 6px;}"
        )

        self._build_ui()
        self._load_settings()
        self.updateFinished.connect(self._finish_update)
        self.appUpdateFinished.connect(self._finish_app_update)
        self.appUpdateAvailable.connect(self._prompt_app_update)

    def _build_ui(self):
        # Main vertical layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(200, 100, 200, 100)
        main_layout.setSpacing(12)
        main_layout.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)

        # Title
        self.remote_settings_title = QLabel("Remote Pycros", self)
        self.remote_settings_title.setStyleSheet(
            "color: #dcdcdc; background: transparent; font-size: 18px; font-weight: 600;"
        )
        main_layout.addWidget(self.remote_settings_title, 0, Qt.AlignLeft)

        # Row 1: Repo URL
        row1_layout = QHBoxLayout()
        self.repo_url_label = QLabel("Repo URL", self)
        self.repo_url_label.setFixedWidth(100)
        self.repo_url_label.setStyleSheet("color: #dcdcdc; background: transparent;")
        self.repo_url_field = LineEdit(self)
        self.repo_url_field.setEnabled(False)
        self.repo_url_field.setStyleSheet(self.field_disabled_style)
        self.repo_url_btn = PrimaryPushButton("Edit", self)
        self.repo_url_btn.setFixedWidth(80)
        self.repo_url_btn.clicked.connect(lambda: self._toggle_edit("repo_url"))
        row1_layout.addWidget(self.repo_url_label)
        row1_layout.addWidget(self.repo_url_field)
        row1_layout.addWidget(self.repo_url_btn)
        main_layout.addLayout(row1_layout)

        # Row 2: Branch
        row2_layout = QHBoxLayout()
        self.branch_label = QLabel("Branch", self)
        self.branch_label.setFixedWidth(100)
        self.branch_label.setStyleSheet("color: #dcdcdc; background: transparent;")
        self.branch_field = LineEdit(self)
        self.branch_field.setEnabled(False)
        self.branch_field.setStyleSheet(self.field_disabled_style)
        self.branch_btn = PrimaryPushButton("Edit", self)
        self.branch_btn.setFixedWidth(80)
        self.branch_btn.clicked.connect(lambda: self._toggle_edit("repo_branch"))
        row2_layout.addWidget(self.branch_label)
        row2_layout.addWidget(self.branch_field)
        row2_layout.addWidget(self.branch_btn)
        main_layout.addLayout(row2_layout)

        # Row 3: Directory
        row3_layout = QHBoxLayout()
        self.directory_label = QLabel("Directory", self)
        self.directory_label.setFixedWidth(100)
        self.directory_label.setStyleSheet("color: #dcdcdc; background: transparent;")
        self.directory_field = LineEdit(self)
        self.directory_field.setEnabled(False)
        self.directory_field.setStyleSheet(self.field_disabled_style)
        self.directory_btn = PrimaryPushButton("Edit", self)
        self.directory_btn.setFixedWidth(80)
        self.directory_btn.clicked.connect(lambda: self._toggle_edit("repo_directory"))
        row3_layout.addWidget(self.directory_label)
        row3_layout.addWidget(self.directory_field)
        row3_layout.addWidget(self.directory_btn)
        main_layout.addLayout(row3_layout)

        # Row 4: Update remote on launch (icon toggle matches packages screen)
        row4_layout = QHBoxLayout()
        self.update_remote_toggle = CheckIconButton(self, initially_checked=False)
        self.update_remote_label = QLabel("Update remote pycros on launch", self)
        self.update_remote_label.setStyleSheet("color: #dcdcdc; background: transparent;")
        row4_layout.addSpacing(110)
        row4_layout.addWidget(self.update_remote_toggle, 0, Qt.AlignVCenter)
        row4_layout.addWidget(self.update_remote_label, 0, Qt.AlignVCenter)
        row4_layout.addStretch(1)
        main_layout.addLayout(row4_layout)
        self.update_remote_toggle.toggledManually.connect(lambda _: self._save_settings())

        # Row 5: Update button
        row5_layout = QHBoxLayout()
        self.update_btn = PrimaryPushButton("Update", self)
        self.update_btn.setFixedWidth(150)
        self.update_btn.clicked.connect(self._on_update_clicked)
        row5_layout.addStretch(1)
        row5_layout.addWidget(self.update_btn)
        row5_layout.addStretch(1)
        main_layout.addLayout(row5_layout)

        # Divider
        main_layout.addSpacing(24)
        divider = QFrame(self)
        divider.setFrameShape(QFrame.HLine)
        divider.setFrameShadow(QFrame.Sunken)
        divider.setStyleSheet("color: #3a3a3a; background: transparent;")
        main_layout.addWidget(divider)
        main_layout.addSpacing(12)

        # App Source title
        self.app_source_title = QLabel("App Source", self)
        self.app_source_title.setStyleSheet(
            "color: #dcdcdc; background: transparent; font-size: 18px; font-weight: 600;"
        )
        main_layout.addWidget(self.app_source_title, 0, Qt.AlignLeft)

        # App Row 1: Repo URL
        app_row1_layout = QHBoxLayout()
        self.app_repo_url_label = QLabel("Repo URL", self)
        self.app_repo_url_label.setFixedWidth(100)
        self.app_repo_url_label.setStyleSheet("color: #dcdcdc; background: transparent;")
        self.app_repo_url_field = LineEdit(self)
        self.app_repo_url_field.setEnabled(False)
        self.app_repo_url_field.setStyleSheet(self.field_disabled_style)
        self.app_repo_url_btn = PrimaryPushButton("Edit", self)
        self.app_repo_url_btn.setFixedWidth(80)
        self.app_repo_url_btn.clicked.connect(lambda: self._toggle_edit("app_repo_url"))
        app_row1_layout.addWidget(self.app_repo_url_label)
        app_row1_layout.addWidget(self.app_repo_url_field)
        app_row1_layout.addWidget(self.app_repo_url_btn)
        main_layout.addLayout(app_row1_layout)

        # App Row 2: Branch
        app_row2_layout = QHBoxLayout()
        self.app_branch_label = QLabel("Branch", self)
        self.app_branch_label.setFixedWidth(100)
        self.app_branch_label.setStyleSheet("color: #dcdcdc; background: transparent;")
        self.app_branch_field = LineEdit(self)
        self.app_branch_field.setEnabled(False)
        self.app_branch_field.setStyleSheet(self.field_disabled_style)
        self.app_branch_btn = PrimaryPushButton("Edit", self)
        self.app_branch_btn.setFixedWidth(80)
        self.app_branch_btn.clicked.connect(lambda: self._toggle_edit("app_repo_branch"))
        app_row2_layout.addWidget(self.app_branch_label)
        app_row2_layout.addWidget(self.app_branch_field)
        app_row2_layout.addWidget(self.app_branch_btn)
        main_layout.addLayout(app_row2_layout)

        # App Row 3: Directory
        app_row3_layout = QHBoxLayout()
        self.app_directory_label = QLabel("Directory", self)
        self.app_directory_label.setFixedWidth(100)
        self.app_directory_label.setStyleSheet("color: #dcdcdc; background: transparent;")
        self.app_directory_field = LineEdit(self)
        self.app_directory_field.setEnabled(False)
        self.app_directory_field.setStyleSheet(self.field_disabled_style)
        self.app_directory_btn = PrimaryPushButton("Edit", self)
        self.app_directory_btn.setFixedWidth(80)
        self.app_directory_btn.clicked.connect(lambda: self._toggle_edit("app_repo_directory"))
        app_row3_layout.addWidget(self.app_directory_label)
        app_row3_layout.addWidget(self.app_directory_field)
        app_row3_layout.addWidget(self.app_directory_btn)
        main_layout.addLayout(app_row3_layout)

        # App Row 4: Check updates on launch
        app_row4_layout = QHBoxLayout()
        self.app_update_toggle = CheckIconButton(self, initially_checked=False)
        self.app_update_label = QLabel("Check for updates on launch", self)
        self.app_update_label.setStyleSheet("color: #dcdcdc; background: transparent;")
        app_row4_layout.addSpacing(110)
        app_row4_layout.addWidget(self.app_update_toggle, 0, Qt.AlignVCenter)
        app_row4_layout.addWidget(self.app_update_label, 0, Qt.AlignVCenter)
        app_row4_layout.addStretch(1)
        main_layout.addLayout(app_row4_layout)
        self.app_update_toggle.toggledManually.connect(lambda _: self._save_settings())

        # App Row 5: Version (info)
        app_row5_layout = QHBoxLayout()
        self.app_version_label = QLabel("Version", self)
        self.app_version_label.setFixedWidth(100)
        self.app_version_label.setStyleSheet("color: #dcdcdc; background: transparent;")
        self.app_version_field = LineEdit(self)
        self.app_version_field.setEnabled(False)
        self.app_version_field.setStyleSheet(self.field_disabled_style)
        self.app_version_field.setText(APP_VERSION)
        app_row5_layout.addWidget(self.app_version_label)
        app_row5_layout.addWidget(self.app_version_field)
        app_row5_layout.addSpacing(80)
        main_layout.addLayout(app_row5_layout)

        # App Row 6: Force update button
        app_row6_layout = QHBoxLayout()
        self.force_update_btn = PrimaryPushButton("Force Update", self)
        self.force_update_btn.setFixedWidth(150)
        self.force_update_btn.clicked.connect(self._on_force_update_clicked)
        app_row6_layout.addStretch(1)
        app_row6_layout.addWidget(self.force_update_btn)
        app_row6_layout.addStretch(1)
        main_layout.addLayout(app_row6_layout)

        if not SHOW_REPO_FIELDS_IN_SETTINGS:
            for w in (
                self.app_repo_url_label,
                self.app_repo_url_field,
                self.app_repo_url_btn,
                self.app_branch_label,
                self.app_branch_field,
                self.app_branch_btn,
                self.app_directory_label,
                self.app_directory_field,
                self.app_directory_btn,
            ):
                w.hide()

    def _load_settings(self):
        """Load settings from settings.json"""
        try:
            settings = {}
            if os.path.exists(self.settings_file):
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    if isinstance(loaded, dict):
                        settings = loaded

            self.repo_url_field.setText(settings.get("repo_url", ""))
            self.branch_field.setText(settings.get("repo_branch", ""))
            self.directory_field.setText(settings.get("repo_directory", ""))
            update_on_launch = settings.get("update_remote_on_launch", False)
            self.update_remote_toggle.setChecked(bool(update_on_launch))

            self.app_repo_url_field.setText(settings.get("app_repo_url", "https://github.com/RisPNG/pycro-station.git"))
            self.app_branch_field.setText(settings.get("app_repo_branch", "main"))
            self.app_directory_field.setText(settings.get("app_repo_directory", "src"))
            app_update_on_launch = settings.get("app_update_on_launch", False)
            self.app_update_toggle.setChecked(bool(app_update_on_launch))
            self.app_version_field.setText(APP_VERSION)
        except Exception as e:
            print(f"Error loading settings: {e}")

    def _save_settings(self, *_args):
        """Save settings to settings.json"""
        try:
            settings = {}
            try:
                if os.path.exists(self.settings_file):
                    with open(self.settings_file, "r", encoding="utf-8") as f:
                        loaded = json.load(f)
                        if isinstance(loaded, dict):
                            settings = loaded
            except Exception:
                settings = {}

            settings.update({
                "repo_url": self.repo_url_field.text(),
                "repo_branch": self.branch_field.text(),
                "repo_directory": self.directory_field.text(),
                "update_remote_on_launch": self.update_remote_toggle.isChecked(),
                "app_repo_url": self.app_repo_url_field.text(),
                "app_repo_branch": self.app_branch_field.text(),
                "app_repo_directory": self.app_directory_field.text(),
                "app_update_on_launch": self.app_update_toggle.isChecked(),
            })

            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def _toggle_edit(self, field_name):
        """Toggle edit/save mode for a field"""
        field_map = {
            "repo_url": (self.repo_url_field, self.repo_url_btn),
            "repo_branch": (self.branch_field, self.branch_btn),
            "repo_directory": (self.directory_field, self.directory_btn),
            "app_repo_url": (self.app_repo_url_field, self.app_repo_url_btn),
            "app_repo_branch": (self.app_branch_field, self.app_branch_btn),
            "app_repo_directory": (self.app_directory_field, self.app_directory_btn),
        }

        pair = field_map.get(field_name)
        if not pair:
            return
        field, btn = pair

        if self.editing_states[field_name]:
            # Currently editing, save the changes
            field.setEnabled(False)
            field.setStyleSheet(self.field_disabled_style)
            btn.setText("Edit")
            self.editing_states[field_name] = False
            self._save_settings()
        else:
            # Not editing, enable the field
            field.setEnabled(True)
            field.setStyleSheet(self.field_enabled_style)
            field.setFocus()
            btn.setText("Save")
            self.editing_states[field_name] = True

    def _on_update_clicked(self):
        """Handle update button click"""
        self._start_update(show_dialog=True)

    def _on_force_update_clicked(self):
        """Handle force app update button click"""
        self._start_app_update(show_dialog=True)

    def _start_update(self, show_dialog: bool):
        """Start update process with optional popup."""
        self._show_update_dialog = show_dialog

        # Prefer values from settings.json (falls back to current field values)
        repo_url = (self.repo_url_field.text() or "").strip()
        branch = (self.branch_field.text() or "").strip()
        repo_dir = (self.directory_field.text() or "").strip()
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    repo_url = (saved.get("repo_url") or repo_url).strip()
                    branch = (saved.get("repo_branch") or branch).strip()
                    repo_dir = (saved.get("repo_directory") or repo_dir).strip()
        except Exception:
            pass
        branch = branch or "main"

        if not repo_url:
            if show_dialog:
                MessageBox("Missing repo URL", "Please provide a repository URL.", self).exec()
            else:
                print("Skipping remote update: repo URL missing.")
            return

        # Disable the button and show status
        self.update_btn.setEnabled(False)
        self.update_btn.setText("Updating...")

        dest_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        dest_path = os.path.join(dest_root, "remote_pycros")

        def worker():
            temp_dir = tempfile.mkdtemp(prefix="pycro_update_")
            clone_dir = os.path.join(temp_dir, "repo")
            dest_tmp = dest_path + ".tmp"
            try:
                # Remove any previous synced content before pulling new copy
                for path in (dest_tmp, dest_path):
                    if os.path.exists(path):
                        shutil.rmtree(path)

                archive_url = self._build_archive_url(repo_url, branch)
                archive_file = os.path.join(temp_dir, "repo.zip")

                # Download archive (no git dependency)
                self._download_url_to_file(archive_url, archive_file)

                # Extract safely
                os.makedirs(clone_dir, exist_ok=True)
                with zipfile.ZipFile(archive_file, 'r') as zf:
                    self._safe_extract(zf, clone_dir)

                clone_root = self._find_extract_root(clone_dir)
                source_path = clone_root if not repo_dir else os.path.abspath(os.path.join(clone_root, repo_dir))
                if os.path.commonpath([clone_root, source_path]) != clone_root:
                    raise ValueError("Invalid directory path specified.")
                if not os.path.isdir(source_path):
                    raise FileNotFoundError(f"Directory '{repo_dir}' not found in branch '{branch}'.")

                # clean existing targets
                # copy to temp then atomically replace
                shutil.copytree(source_path, dest_tmp)
                os.replace(dest_tmp, dest_path)

                self.updateFinished.emit(True, f"Fetched '{repo_dir or 'entire repo'}' from {branch}.")
            except Exception as e:
                self.updateFinished.emit(False, str(e))
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)
                shutil.rmtree(dest_tmp, ignore_errors=True)

        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _finish_update(self, success: bool, message: str):
        # Restore button state
        self.update_btn.setEnabled(True)
        self.update_btn.setIcon(QIcon())
        self.update_btn.setText("Update")

        if not self._show_update_dialog:
            # Reset for future manual updates
            self._show_update_dialog = True
            if not success:
                print(f"Silent remote update failed: {message}")
            return

        title = "Success" if success else "Update failed"
        msg = MessageBox(title, message or "", self)
        msg.yesButton.setText("OK")
        msg.cancelButton.hide()
        msg.exec()

    def run_update_on_launch_if_enabled(self):
        """Kick off remote update at launch when setting is enabled."""
        if self._launch_update_started:
            return
        if not self.update_remote_toggle.isChecked():
            return
        self._launch_update_started = True
        QTimer.singleShot(0, lambda: self._start_update(show_dialog=False))

    def run_app_update_check_on_launch_if_enabled(self):
        """Check for app updates at launch when enabled."""
        if self._launch_app_update_started:
            return
        if not self.app_update_toggle.isChecked():
            return
        self._launch_app_update_started = True
        QTimer.singleShot(0, self._start_app_update_check)

    def _start_app_update_check(self):
        """Check remote main.py version and prompt if newer."""
        repo_url, branch, repo_dir = self._get_app_source_settings()
        if not repo_url:
            return

        def worker():
            temp_dir = tempfile.mkdtemp(prefix="pycro_app_check_")
            clone_dir = os.path.join(temp_dir, "repo")
            try:
                archive_url = self._build_archive_url(repo_url, branch)
                archive_file = os.path.join(temp_dir, "repo.zip")

                self._download_url_to_file(archive_url, archive_file)

                os.makedirs(clone_dir, exist_ok=True)
                with zipfile.ZipFile(archive_file, "r") as zf:
                    self._safe_extract(zf, clone_dir)

                clone_root = self._find_extract_root(clone_dir)
                source_path = clone_root if not repo_dir else os.path.abspath(os.path.join(clone_root, repo_dir))
                if os.path.commonpath([clone_root, source_path]) != clone_root:
                    raise ValueError("Invalid directory path specified.")

                main_py = os.path.join(source_path, "main.py")
                if not os.path.isfile(main_py):
                    return

                with open(main_py, "r", encoding="utf-8", errors="ignore") as f:
                    text = f.read()

                remote_version = self._extract_version_from_text(text)
                if not remote_version:
                    return
                if self._is_version_newer(remote_version, APP_VERSION):
                    self.appUpdateAvailable.emit(remote_version)
            except Exception as e:
                print(f"App update check failed: {e}")
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _prompt_app_update(self, remote_version: str):
        """Show update prompt with later/disable/update-now options."""
        parent = self.window() or self
        text = f"An update is available for version {remote_version}.\nCurrent version: {APP_VERSION}"

        msg = MessageBox("Update Available", text, parent)
        msg.yesButton.setText("Update now")
        msg.cancelButton.setText("Later")

        # Make the dialog a bit wider for readability
        try:
            min_width = 520
            msg.widget.setFixedWidth(max(msg.widget.width(), min_width))
        except Exception:
            pass

        disable_btn = QPushButton("Disable update check", msg.buttonGroup)
        disable_btn.setObjectName("cancelButton")
        try:
            disable_btn.setAttribute(Qt.WA_LayoutUsesWidgetRect)
        except Exception:
            pass

        # Prefer button order: Later / Disable / Update now
        try:
            msg.buttonLayout.removeWidget(msg.yesButton)
            msg.buttonLayout.removeWidget(msg.cancelButton)
            msg.buttonLayout.addWidget(msg.cancelButton, 1, Qt.AlignVCenter)
            msg.buttonLayout.addWidget(disable_btn, 1, Qt.AlignVCenter)
            msg.buttonLayout.addWidget(msg.yesButton, 1, Qt.AlignVCenter)
        except Exception:
            try:
                msg.buttonLayout.insertWidget(1, disable_btn, 1, Qt.AlignVCenter)
            except Exception:
                pass

        choice = {"value": "later"}

        def choose(value: str):
            choice["value"] = value

        def on_disable_clicked():
            choose("disable")
            try:
                msg.reject()
            except Exception:
                try:
                    msg.close()
                except Exception:
                    pass

        msg.cancelButton.clicked.connect(lambda: choose("later"))
        msg.yesButton.clicked.connect(lambda: choose("update"))
        disable_btn.clicked.connect(on_disable_clicked)

        try:
            msg.exec()
        except Exception:
            return

        if choice["value"] == "disable":
            self.app_update_toggle.setChecked(False)
            self._save_settings()
        elif choice["value"] == "update":
            # Let the dialog fully close before kicking off the update.
            QTimer.singleShot(0, lambda: self._start_app_update(show_dialog=True))
    @staticmethod
    def _build_archive_url(repo_url: str, branch: str) -> str:
        base = (repo_url or "").strip()
        if base.endswith(".git"):
            base = base[:-4]
        base = base.rstrip("/")
        return f"{base}/archive/refs/heads/{branch}.zip"

    @staticmethod
    def _download_url_to_file(url: str, dst_path: str, *, timeout: int = 60, attempts: int = 3):
        """Download a URL to a local file with retry/backoff for transient network failures."""
        headers = {"User-Agent": "pycro-station"}
        attempts = max(1, int(attempts))
        last_err: Exception | None = None

        for attempt in range(1, attempts + 1):
            try:
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=timeout) as resp, open(dst_path, "wb") as out:
                    shutil.copyfileobj(resp, out)
                return
            except urllib.error.HTTPError as e:
                last_err = e
                code = int(getattr(e, "code", 0) or 0)
                # Retry 5xx errors; surface others immediately.
                if not (500 <= code < 600 and attempt < attempts):
                    raise RuntimeError(f"Failed to download archive (HTTP {e.code}).") from e
            except (urllib.error.URLError, ConnectionResetError, TimeoutError, OSError) as e:
                last_err = e
            except Exception as e:
                last_err = e

            if attempt < attempts:
                time.sleep(min(2**attempt, 8))

        detail = ""
        if isinstance(last_err, urllib.error.URLError) and getattr(last_err, "reason", None):
            detail = str(last_err.reason)
        elif last_err is not None:
            detail = str(last_err)

        msg = "Failed to download update archive."
        if detail:
            msg = f"{msg} ({detail})"
        msg = f"{msg} Please check your internet/VPN and try again."
        raise RuntimeError(msg) from last_err

    @staticmethod
    def _safe_extract(zip_file: zipfile.ZipFile, target_dir: str):
        target_dir_abs = os.path.abspath(target_dir)
        for member in zip_file.infolist():
            member_path = os.path.abspath(os.path.join(target_dir, member.filename))
            if not member_path.startswith(target_dir_abs + os.sep) and member_path != target_dir_abs:
                raise ValueError("Archive contains unsafe paths.")
        zip_file.extractall(target_dir)

    @staticmethod
    def _find_extract_root(tmp_dir: str) -> str:
        dirs = [d for d in os.listdir(tmp_dir) if os.path.isdir(os.path.join(tmp_dir, d))]
        preferred = [d for d in dirs if d != "__MACOSX"]
        target_list = preferred if preferred else dirs
        if len(target_list) >= 1:
            return os.path.join(tmp_dir, sorted(target_list)[0])
        return tmp_dir

    def _get_app_source_settings(self) -> tuple[str, str, str]:
        repo_url = (self.app_repo_url_field.text() or "").strip()
        branch = (self.app_branch_field.text() or "").strip()
        repo_dir = (self.app_directory_field.text() or "").strip()
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    if isinstance(saved, dict):
                        repo_url = (saved.get("app_repo_url") or repo_url).strip()
                        branch = (saved.get("app_repo_branch") or branch).strip()
                        repo_dir = (saved.get("app_repo_directory") or repo_dir).strip()
        except Exception:
            pass
        branch = branch or "main"
        repo_dir = repo_dir or "src"
        return repo_url, branch, repo_dir

    @staticmethod
    def _extract_version_from_text(text: str) -> str | None:
        patterns = [
            r'^\s*APP_VERSION\s*=\s*[\'"]([^\'"]+)[\'"]\s*$',
            r'^\s*__version__\s*=\s*[\'"]([^\'"]+)[\'"]\s*$',
            r'^\s*VERSION\s*=\s*[\'"]([^\'"]+)[\'"]\s*$',
            r'@version\s+([0-9A-Za-z.+-]+)',
        ]
        for pat in patterns:
            m = re.search(pat, text, flags=re.MULTILINE)
            if m:
                return (m.group(1) or "").strip()
        return None

    @staticmethod
    def _version_key(version: str) -> tuple[int, ...]:
        nums = [int(x) for x in re.findall(r"\d+", version or "")]
        return tuple(nums) if nums else (0,)

    @classmethod
    def _is_version_newer(cls, remote_version: str, current_version: str) -> bool:
        r = cls._version_key(remote_version)
        c = cls._version_key(current_version)
        max_len = max(len(r), len(c))
        r = r + (0,) * (max_len - len(r))
        c = c + (0,) * (max_len - len(c))
        return r > c

    def _start_app_update(self, show_dialog: bool):
        """Update this app by replacing src/*.py from the configured repo."""
        self._show_app_update_dialog = show_dialog

        repo_url, branch, repo_dir = self._get_app_source_settings()
        if not repo_url:
            if show_dialog:
                MessageBox("Missing repo URL", "Please provide a repository URL.", self).exec()
            else:
                print("Skipping app update: repo URL missing.")
            return

        self.force_update_btn.setEnabled(False)
        self.force_update_btn.setText("Updating...")

        src_dir = os.path.abspath(os.path.dirname(__file__))
        targets = ["TitleBar.py", "PycroGrid.py", "PackagesPage.py", "main.py"]

        def worker():
            temp_dir = tempfile.mkdtemp(prefix="pycro_app_update_")
            clone_dir = os.path.join(temp_dir, "repo")
            try:
                archive_url = self._build_archive_url(repo_url, branch)
                archive_file = os.path.join(temp_dir, "repo.zip")

                self._download_url_to_file(archive_url, archive_file)

                os.makedirs(clone_dir, exist_ok=True)
                with zipfile.ZipFile(archive_file, "r") as zf:
                    self._safe_extract(zf, clone_dir)

                clone_root = self._find_extract_root(clone_dir)
                source_path = clone_root if not repo_dir else os.path.abspath(os.path.join(clone_root, repo_dir))
                if os.path.commonpath([clone_root, source_path]) != clone_root:
                    raise ValueError("Invalid directory path specified.")
                if not os.path.isdir(source_path):
                    raise FileNotFoundError(f"Directory '{repo_dir}' not found in branch '{branch}'.")

                missing = [n for n in targets if not os.path.isfile(os.path.join(source_path, n))]
                if missing:
                    missing_list = ", ".join(missing)
                    raise FileNotFoundError(f"Missing {missing_list} in '{repo_dir or 'repo root'}'.")

                for name in targets:
                    src_file = os.path.join(source_path, name)
                    dst_file = os.path.join(src_dir, name)
                    tmp_file = dst_file + ".tmp"
                    shutil.copy2(src_file, tmp_file)
                    os.replace(tmp_file, dst_file)

                msg = "Updated app source files."
                try:
                    main_py = os.path.join(source_path, "main.py")
                    with open(main_py, "r", encoding="utf-8", errors="ignore") as f:
                        remote_version = self._extract_version_from_text(f.read())
                    if remote_version:
                        msg = f"Updated to {remote_version}."
                except Exception:
                    pass

                self.appUpdateFinished.emit(True, msg)
            except Exception as e:
                self.appUpdateFinished.emit(False, str(e))
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)
                for name in targets:
                    tmp_path = os.path.join(src_dir, name + ".tmp")
                    try:
                        if os.path.exists(tmp_path):
                            os.remove(tmp_path)
                    except Exception:
                        pass

        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _relaunch_app(self) -> bool:
        """Launch a new instance and quit the current one."""
        program = sys.executable
        main_py = os.path.realpath(__file__)
        args = [main_py] + sys.argv[1:]
        workdir = os.path.abspath(os.path.join(os.path.dirname(main_py), ".."))

        ok = False
        try:
            result = QProcess.startDetached(program, args, workdir)
            ok = result[0] if isinstance(result, tuple) else bool(result)
        except Exception:
            ok = False

        if not ok:
            import subprocess

            try:
                kwargs = {"cwd": workdir, "close_fds": True}
                if os.name == "nt":
                    kwargs["creationflags"] = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
                subprocess.Popen([program] + args, **kwargs)
                ok = True
            except Exception as e:
                MessageBox("Relaunch failed", str(e), self.window() or self).exec()
                return False

        try:
            app = QApplication.instance()
            if app is not None:
                app.closeAllWindows()
        except Exception:
            pass

        QTimer.singleShot(0, QCoreApplication.quit)
        return True

    def _finish_app_update(self, success: bool, message: str):
        self.force_update_btn.setEnabled(True)
        self.force_update_btn.setText("Force Update")

        if not self._show_app_update_dialog:
            self._show_app_update_dialog = True
            if not success:
                print(f"Silent app update failed: {message}")
            return

        if success:
            details = (message or "").strip()
            text = f"{details}\n\nRelaunch Pycro Station now?" if details else "Relaunch Pycro Station now?"
            parent = self.window() or self
            prompt = MessageBox(
                "Update complete",
                text,
                parent,
            )
            prompt.yesButton.setText("Relaunch")
            prompt.cancelButton.setText("Later")
            if prompt.exec():
                self._relaunch_app()
            return

        msg = MessageBox("Update failed", message or "", self.window() or self)
        msg.yesButton.setText("OK")
        msg.cancelButton.hide()
        msg.exec()



class Window(MSFluentWindow):
    def __init__(self):
        # self.isMicaEnabled = False
        super().__init__()
        self.setTitleBar(CustomTitleBar(self))
        self.tabBar = self.titleBar.tabBar  # type: TabBar

        setTheme(Theme.DARK)

        # Create shortcuts for Save and Open
        self.save_shortcut = QShortcut(QKeySequence.StandardKey.Save, self)
        self.open_shortcut = QShortcut(QKeySequence.StandardKey.Open, self)

        # Connect the shortcuts to functions
        self.save_shortcut.activated.connect(self.save_document)
        self.open_shortcut.activated.connect(self.open_document)

        # Holds active macro pages mapped by routeKey
        self.macro_pages: dict[str, QWidget] = {}
        self.macro_labels: dict[str, str] = {}
        # Remember last active sidebar interface (used when closing tabs)
        self._last_sidebar_widget: QWidget | None = None
        self._last_sidebar_route_key: str | None = None


        # create sub interface
        self.homeInterface = QStackedWidget(self, objectName='homeInterface')
        # remove frame and let it use full space
        self.homeInterface.setFrameShape(QFrame.NoFrame)
        self.homeInterface.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.homeInterface.setStyleSheet("background: transparent;")

        # Separate container for macro tab content (disconnected from Hub navigation)
        self.tabsInterface = AnimatedStackedWidget(self)
        self.tabsInterface.setObjectName('tabsInterface')
        self.tabsInterface.setFrameShape(QFrame.NoFrame)
        self.tabsInterface.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tabsInterface.setStyleSheet("background: transparent;")

        self.settingsInterface = Settings(self)
        self.settingsInterface.setObjectName("settingsInterface")
        # No initial tabs; Hub shows the grid

        self.initNavigation()
        self.initWindow()
        try:
            self._select_start_interface()
        except Exception:
            pass
        try:
            self.settingsInterface.updateFinished.connect(self._on_settings_update)
        except Exception:
            pass
        try:
            self.settingsInterface.run_update_on_launch_if_enabled()
        except Exception:
            pass
        try:
            self.settingsInterface.run_app_update_check_on_launch_if_enabled()
        except Exception:
            pass
        # Disable tab highlight when switching to navigation items
        self.stackedWidget.currentChanged.connect(self.onContentChanged)
        # initialize packages lock state
        try:
            self.packagesPage.setLocked(False)
        except Exception:
            pass

    def initNavigation(self):
        hub = QIcon(QPixmap.fromImage(ImageQt(TablerIcons.load(
            OutlineIcon.CATEGORY,
            size=24,
            color="#FFFFFF",
            stroke_width=2.0,

        ))))
        self.addSubInterface(self.homeInterface, hub, 'Hub', hub, NavigationItemPosition.TOP)

        # Stars page under Hub
        ti_stars = QIcon(QPixmap.fromImage(ImageQt(TablerIcons.load(
            OutlineIcon.STARS,
            size=24,
            color="#FFFFFF",
            stroke_width=2.0,
        ))))
        self.starsGrid = PycroGrid(self, stars_only=True)
        self.starsGrid.setObjectName("starsInterface")
        self.addSubInterface(self.starsGrid, ti_stars, 'Stars', ti_stars, NavigationItemPosition.TOP)

        # Packages page button under Hub
        ti_pkg = QIcon(QPixmap.fromImage(ImageQt(TablerIcons.load(
            OutlineIcon.PLAYLIST_X,
            size=24,
            color="#FFFFFF",
            stroke_width=2.0,
        ))))
        self.packagesPage = PackagesPage(self)
        self.addSubInterface(self.packagesPage, ti_pkg, 'Packages', ti_pkg, NavigationItemPosition.TOP)
        # revalidate hub grid after uninstall/install changes
        try:
            self.packagesPage.packagesChanged.connect(lambda: (self.hubGrid.refresh(), getattr(self, "starsGrid", None) and self.starsGrid.refresh()))
        except Exception:
            pass
        self.addSubInterface(self.settingsInterface, FIF.SETTING, 'Settings', FIF.SETTING, NavigationItemPosition.BOTTOM)
        # self.addSubInterface(self.settingInterface, FIF.SETTING, 'Settings', FIF.SETTING,  NavigationItemPosition.BOTTOM)
        self.navigationInterface.addItem(
            routeKey='Help',
            icon=FIF.INFO,
            text='About',
            onClick=self.showMessageBox,
            selectable=False,
            position=NavigationItemPosition.BOTTOM)

        # Build Hub grid page
        self.hubGrid = PycroGrid(self)
        self.homeInterface.addWidget(self.hubGrid)
        self.homeInterface.setCurrentWidget(self.hubGrid)

        # Add tabsInterface to stackedWidget (not to navigation - keeps it disconnected)
        self.stackedWidget.addWidget(self.tabsInterface)

        # Select Hub in navigation
        self.navigationInterface.setCurrentItem(self.homeInterface.objectName())

        # Tab bar hooks: tabs represent launched Pycros only
        self.tabBar.currentChanged.connect(self.onTabChanged)

    def initWindow(self):
        # Load icon from file (fallback to generated icon if file doesn't exist)
        icon_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "pycro-station-icon.png")
        if os.path.exists(icon_path):
            app_icon = QIcon(icon_path)
        else:
            # Fallback to generated icon
            ti_planet = ImageQt(TablerIcons.load(
                OutlineIcon.PLANET,
                size=24,
                color="#FFFFFF",
                stroke_width=2.0,
            ))
            app_icon = QIcon(QPixmap.fromImage(ti_planet))

        self.resize(975, 780)
        self.setWindowIcon(app_icon)
        self.setWindowTitle('Pycro Station')

        screen = QGuiApplication.primaryScreen()
        avail = screen.availableGeometry() if screen else QRect(0, 0, 1920, 1080)
        target_w = min(self.width(), avail.width())
        target_h = min(self.height(), avail.height())
        self.resize(target_w, target_h)
        self.move(
            avail.x() + (avail.width() - target_w) // 2,
            avail.y() + (avail.height() - target_h) // 2,
        )

    def _select_start_interface(self):
        """Select the initial navigation interface on launch."""
        default_widget: QWidget = self.homeInterface

        try:
            settings_path = os.path.join(os.path.dirname(__file__), "settings.json")
            if os.path.exists(settings_path):
                with open(settings_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    if isinstance(loaded, dict):
                        starred = loaded.get("starred_pycros", [])
                        if isinstance(starred, list) and len(starred) > 0:
                            default_widget = getattr(self, "starsGrid", None) or self.homeInterface
        except Exception:
            default_widget = self.homeInterface

        try:
            if default_widget is self.homeInterface:
                self.homeInterface.setCurrentWidget(self.hubGrid)
        except Exception:
            pass

        try:
            self.switchTo(default_widget)
        except Exception:
            pass

        try:
            self.navigationInterface.setCurrentItem(default_widget.objectName())
        except Exception:
            pass

        try:
            self._remember_sidebar_interface(default_widget)
        except Exception:
            pass

    def showHub(self):
        """Switch to the Hub grid view."""
        self.switchTo(self.homeInterface)
        self.homeInterface.setCurrentWidget(self.hubGrid)
        try:
            self.navigationInterface.setCurrentItem(self.homeInterface.objectName())
        except Exception:
            pass

    def _remember_sidebar_interface(self, widget: QWidget | None):
        if widget is None or widget is self.tabsInterface:
            return
        try:
            self._last_sidebar_widget = widget
            self._last_sidebar_route_key = widget.objectName()
        except Exception:
            self._last_sidebar_widget = widget

    def _restore_last_sidebar_interface(self):
        """Return to the last navigation interface used before entering tabs."""
        target = self._last_sidebar_widget
        if target is None or target is self.tabsInterface:
            target = self.homeInterface

        try:
            if target is self.homeInterface:
                self.homeInterface.setCurrentWidget(self.hubGrid)
        except Exception:
            pass

        try:
            self.switchTo(target)
        except Exception:
            try:
                self.switchTo(self.homeInterface)
            except Exception:
                pass

        try:
            self.navigationInterface.setCurrentItem(target.objectName())
        except Exception:
            try:
                if self._last_sidebar_route_key:
                    self.navigationInterface.setCurrentItem(self._last_sidebar_route_key)
            except Exception:
                pass

    def showMessageBox(self):
        w = MessageBox(
            'Pycro Station',
            (
                    f"Version : {APP_VERSION}"
                    + "\n" + "\n" + "\n" + "Welcome aboard Pycronauts!" + "\n" + "This is the hub to store and launch Pycros" + "\n" + "\n" + "\n" +
                    "Made with 💚 By Ris Peng"
            ),
            self
        )
        w.yesButton.setText('GitHub')
        w.cancelButton.setText('Return')

        if w.exec():
            QDesktopServices.openUrl(QUrl("https://github.com/rispng/"))

    def onTabChanged(self, index: int):
        # Switch content area to the selected macro page
        try:
            routeKey = self.tabBar.currentTab().routeKey()
        except Exception:
            return
        self._show_macro_page(routeKey)

    def onTabClicked(self, index: int):
        """Handle clicks on the current tab to re-activate its content."""
        # Try to derive routeKey from the tab label
        routeKey = None
        try:
            label = self.tabBar.tabText(index)
            for rk, txt in self.macro_labels.items():
                if txt == label:
                    routeKey = rk
                    break
        except Exception:
            pass
        # Fallback to current tab's routeKey if available
        if routeKey is None:
            try:
                routeKey = self.tabBar.currentTab().routeKey()
            except Exception:
                return
        self._show_macro_page(routeKey)

    def _show_macro_page(self, routeKey: str):
        """Ensure the macro page is in the stack and visible."""
        page = self.macro_pages.get(routeKey)
        if page is None:
            return
        # Ensure page is in tabsInterface and visible
        try:
            if self.tabsInterface.indexOf(page) == -1:
                self.tabsInterface.addWidget(page)
        except Exception:
            pass
        # If coming from a different interface, preselect the page without animating
        # (the main switch animation will handle the transition).
        try:
            from_other_interface = self.stackedWidget.currentWidget() is not self.tabsInterface
        except Exception:
            from_other_interface = True

        if from_other_interface:
            try:
                self.tabsInterface.setCurrentWidgetNoAnimation(page)
            except Exception:
                try:
                    self.tabsInterface.setCurrentWidget(page)
                except Exception:
                    pass
            self.switchTo(self.tabsInterface)
        else:
            # tabsInterface is AnimatedStackedWidget - will animate tab-to-tab
            self.tabsInterface.setCurrentWidget(page)
        # Deselect sidebar since tabs are active
        self._deselect_navigation()

    # No external tab add; tabs are created by launching a Pycro from the Hub
    def onTabAddRequested(self):
        pass

    def open_document(self):
        # Guard: open document only applies to text editor mode
        if not self._has_text_editor():
            MessageBox(
                'Not Available',
                'Open is not available in the Pycro grid view.',
                self
            ).exec()
            return

        file_dir = filedialog.askopenfilename(
            title="Select file",
        )
        filename = os.path.basename(file_dir).split('/')[-1]

        if file_dir:
            try:
                with open(file_dir, "r") as f:
                    filedata = f.read()
                    self.current_editor.setPlainText(filedata)
                    try:
                        self.setWindowTitle(f"{os.path.basename(filename)} ~ ZenNotes")
                    except Exception:
                        pass

                    # Check the first line of the text
                    first_line = filedata.split('\n')[0].strip()
                    if first_line == ".LOG":
                        self.current_editor.append(str(datetime.datetime.now()))

            except UnicodeDecodeError:
                MessageBox(
                    'Wrong Filetype! 📝',
                    (
                        "Make sure you've selected a valid file type. Also note that PDF, DOCX, Image Files, are NOT supported in ZenNotes as of now."
                    ),
                    self
                )

    def closeEvent(self, event):
        # If there's no text editor active, just accept the close
        if not self._has_text_editor():
            event.accept()
            return

        a = self.current_editor.toPlainText()

        if a != "":

            w = MessageBox(
                'Confirm Exit',
                (
                        "Do you want to save your 'magnum opus' before exiting? " +
                        "Or would you like to bid adieu to your unsaved masterpiece?"
                ),
                self
            )
            w.yesButton.setText('Yeah')
            w.cancelButton.setText('Nah')

            if w.exec():
                self.save_document()
        else:
            event.accept()  # Close the application

    def save_document(self):
        try:
            if not self._has_text_editor():
                print("No active TWidget found.")
                return  # Check if there is an active TWidget

            text_to_save = self.current_editor.toPlainText()
            print("Text to save:", text_to_save)  # Debug print

            name = filedialog.asksaveasfilename(
                title="Save Your Document"
            )

            print("File path to save:", name)  # Debug print

            if name:
                with open(name, 'w') as file:
                    file.write(text_to_save)
                    title = os.path.basename(name) + " ~ ZenNotes"
                    active_tab_index = self.tabBar.currentIndex()
                    self.tabBar.setTabText(active_tab_index, os.path.basename(name))
                    self.setWindowTitle(title)
                    print("File saved successfully.")  # Debug print
        except Exception as e:
            print(f"An error occurred while saving the document: {e}")

    def addMacroTab(self, routeKey: str, text: str, icon: QIcon, content_widget: QWidget, replace_existing: bool = False):
        """Add or activate a macro tab and show its content.

        replace_existing=True will rebuild the page even if it's already open,
        so live pycro code/description changes are reflected immediately.
        """
        if routeKey in self.macro_pages:
            if not replace_existing:
                # Just focus existing tab/page
                try:
                    self.tabBar.setCurrentTab(routeKey)
                except Exception:
                    pass
                page = self.macro_pages[routeKey]
                try:
                    if self.tabsInterface.indexOf(page) == -1:
                        self.tabsInterface.addWidget(page)
                except Exception:
                    pass
                try:
                    from_other_interface = self.stackedWidget.currentWidget() is not self.tabsInterface
                except Exception:
                    from_other_interface = True

                if from_other_interface:
                    try:
                        self.tabsInterface.setCurrentWidgetNoAnimation(page)
                    except Exception:
                        try:
                            self.tabsInterface.setCurrentWidget(page)
                        except Exception:
                            pass
                    self.switchTo(self.tabsInterface)
                else:
                    self.tabsInterface.setCurrentWidget(page)
                try:
                    self.packagesPage.setLocked(True)
                except Exception:
                    pass
                return

            # Replace existing page with fresh content (hot reload)
            old_page = self.macro_pages.get(routeKey)
            if old_page is not None:
                try:
                    self.tabsInterface.removeWidget(old_page)
                except Exception:
                    pass
                old_page.deleteLater()
            self.macro_pages[routeKey] = content_widget
            self.macro_labels[routeKey] = text

            content_widget.setObjectName(routeKey)
            content_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.tabsInterface.addWidget(content_widget)

            try:
                self.tabBar.setCurrentTab(routeKey)
            except Exception:
                pass
            try:
                from_other_interface = self.stackedWidget.currentWidget() is not self.tabsInterface
            except Exception:
                from_other_interface = True

            if from_other_interface:
                try:
                    self.tabsInterface.setCurrentWidgetNoAnimation(content_widget)
                except Exception:
                    try:
                        self.tabsInterface.setCurrentWidget(content_widget)
                    except Exception:
                        pass
                self.switchTo(self.tabsInterface)
            else:
                self.tabsInterface.setCurrentWidget(content_widget)
            try:
                self.packagesPage.setLocked(True)
            except Exception:
                pass
            return

        # New macro page
        content_widget.setObjectName(routeKey)
        content_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tabsInterface.addWidget(content_widget)
        self.macro_pages[routeKey] = content_widget
        self.macro_labels[routeKey] = text

        self.tabBar.addTab(routeKey, text, icon)
        try:
            self.tabBar.setCurrentTab(routeKey)
        except Exception:
            pass
        try:
            from_other_interface = self.stackedWidget.currentWidget() is not self.tabsInterface
        except Exception:
            from_other_interface = True

        if from_other_interface:
            try:
                self.tabsInterface.setCurrentWidgetNoAnimation(content_widget)
            except Exception:
                try:
                    self.tabsInterface.setCurrentWidget(content_widget)
                except Exception:
                    pass
            self.switchTo(self.tabsInterface)
        else:
            self.tabsInterface.setCurrentWidget(content_widget)
        try:
            self.packagesPage.setLocked(True)
        except Exception:
            pass

    def onTabCloseRequested(self, index: int):
        # Remember current context to decide whether to return to sidebar after closing
        was_current = False
        try:
            was_current = index == self.tabBar.currentIndex()
        except Exception:
            was_current = False
        tabs_were_active = False
        try:
            tabs_were_active = self.stackedWidget.currentWidget() is self.tabsInterface
        except Exception:
            tabs_were_active = False

        # Find routeKey by temporarily selecting the tab or matching text
        routeKey = None
        label = None
        try:
            label = self.tabBar.tabText(index)
        except Exception:
            pass

        # Match by label first
        if label is not None:
            for rk, txt in list(self.macro_labels.items()):
                if txt == label:
                    routeKey = rk
                    break

        # Fallback: select and read routeKey
        if routeKey is None:
            try:
                prev = self.tabBar.currentIndex()
                self.tabBar.setCurrentIndex(index)
                routeKey = self.tabBar.currentTab().routeKey()
                self.tabBar.setCurrentIndex(prev)
            except Exception:
                pass

        if routeKey is None:
            # as a last resort, remove the tab and return to last sidebar view
            try:
                self.tabBar.removeTab(index)
            except Exception:
                pass
            if tabs_were_active and was_current:
                self._restore_last_sidebar_interface()
            return

        # Dispose content widget
        page = self.macro_pages.pop(routeKey, None)
        self.macro_labels.pop(routeKey, None)
        if page is not None:
            try:
                self.tabsInterface.removeWidget(page)
            except Exception:
                pass
            page.deleteLater()

        # Remove the tab
        try:
            self.tabBar.removeTab(index)
        except Exception:
            pass

        remaining_tabs = 0
        try:
            remaining_tabs = len(self.macro_pages)
        except Exception:
            try:
                remaining_tabs = self.tabBar.count()
            except Exception:
                remaining_tabs = 0

        # Return to last sidebar item only when closing the active tab (or when no tabs remain)
        if tabs_were_active and (was_current or remaining_tabs == 0):
            self._restore_last_sidebar_interface()
        # unlock packages if no tabs remain
        try:
            self.packagesPage.setLocked(len(self.macro_pages) > 0)
        except Exception:
            pass

    def closeAllTabs(self):
        """Close all open macro tabs."""
        count = 0
        try:
            count = self.tabBar.count()
        except Exception:
            count = 0
        for i in reversed(range(count)):
            try:
                self.onTabCloseRequested(i)
            except Exception:
                pass

    def _on_settings_update(self, success: bool, _msg: str):
        """Refresh hub when settings update pulls new remote_pycros."""
        if success:
            try:
                self.hubGrid._last_changed_path = os.path.join(os.getcwd(), "remote_pycros")
            except Exception:
                pass
        try:
            self.hubGrid.refresh()
        except Exception:
            pass
        try:
            if hasattr(self, "starsGrid"):
                self.starsGrid.refresh()
        except Exception:
            pass

    def _deselect_navigation(self):
        """Clear selection on navigation sidebar when tabs are active."""
        try:
            for widget in self.navigationInterface.items.values():
                widget.setSelected(False)
        except Exception:
            pass

    def onContentChanged(self, index: int):
        """Toggle tab highlight based on whether tabs or navigation is active."""
        w = self.stackedWidget.widget(index)
        if w is self.tabsInterface:
            # Tabs are active - enable highlight, deselect navigation
            self.titleBar.setTabsSelectionHighlightEnabled(True)
            self._deselect_navigation()
        else:
            # A navigation item is active - disable tab highlight
            self.titleBar.setTabsSelectionHighlightEnabled(False)
            try:
                self._remember_sidebar_interface(w)
            except Exception:
                pass

    def _has_text_editor(self) -> bool:
        """Return True if the current view is a text editor-like widget."""
        editor = getattr(self, 'current_editor', None)
        if editor is None:
            return False
        return callable(getattr(editor, 'toPlainText', None)) and callable(getattr(editor, 'textCursor', None))


if __name__ == '__main__':
    app = QApplication()

    # Set application metadata for proper Linux integration
    app.setApplicationName("Pycro Station")
    app.setApplicationDisplayName("Pycro Station")

    # Set organization for settings persistence
    app.setOrganizationName("PycroStation")
    app.setOrganizationDomain("pycrostation")

    # Set application icon for taskbar/alt-tab (important for GNOME/Linux)
    # Get absolute path to icon file
    icon_path = os.path.abspath(os.path.join(os.path.dirname(os.path.dirname(__file__)), "pycro-station-icon.png"))

    if os.path.exists(icon_path):
        icon = QIcon(icon_path)
        # Set at application level for taskbar/dock
        app.setWindowIcon(icon)
        print(f"Icon loaded from: {icon_path}")
    else:
        print(f"Warning: Icon not found at {icon_path}")

    w = Window()
    w.show()
    app.exec()
