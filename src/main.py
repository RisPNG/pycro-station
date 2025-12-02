"""
The main python file. Run this file to use the app.
"""
import datetime
import json
import os
import shutil
import tempfile
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
from PackagesPage import PackagesPage
from TitleBar import CustomTitleBar

class Settings(QWidget):
    updateFinished = Signal(bool, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings_file = os.path.join(os.path.dirname(__file__), "settings.json")

        # Track editing state for each field
        self.editing_states = {
            "repo_url": False,
            "repo_branch": False,
            "repo_directory": False
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

    def _build_ui(self):
        # Main vertical layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(200, 100, 200, 100)
        main_layout.setSpacing(12)
        main_layout.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)

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

        # Row 4: Update button
        row4_layout = QHBoxLayout()
        self.update_btn = PrimaryPushButton("Update", self)
        self.update_btn.setFixedWidth(150)
        self.update_btn.clicked.connect(self._on_update_clicked)
        row4_layout.addWidget(self.update_btn)
        row4_layout.addStretch(2)
        main_layout.addLayout(row4_layout)

    def _load_settings(self):
        """Load settings from settings.json"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                    self.repo_url_field.setText(settings.get("repo_url", ""))
                    self.branch_field.setText(settings.get("repo_branch", ""))
                    self.directory_field.setText(settings.get("repo_directory", ""))
        except Exception as e:
            print(f"Error loading settings: {e}")

    def _save_settings(self):
        """Save settings to settings.json"""
        try:
            settings = {
                "repo_url": self.repo_url_field.text(),
                "repo_branch": self.branch_field.text(),
                "repo_directory": self.directory_field.text()
            }
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def _toggle_edit(self, field_name):
        """Toggle edit/save mode for a field"""
        if field_name == "repo_url":
            field = self.repo_url_field
            btn = self.repo_url_btn
        elif field_name == "repo_branch":
            field = self.branch_field
            btn = self.branch_btn
        else:  # repo_directory
            field = self.directory_field
            btn = self.directory_btn

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
            MessageBox("Missing repo URL", "Please provide a repository URL.", self).exec()
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
                try:
                    req = urllib.request.Request(archive_url, headers={"User-Agent": "pycro-station"})
                    with urllib.request.urlopen(req, timeout=60) as resp, open(archive_file, "wb") as out:
                        shutil.copyfileobj(resp, out)
                except urllib.error.HTTPError as e:
                    raise RuntimeError(f"Failed to download archive (HTTP {e.code}).")
                except urllib.error.URLError as e:
                    raise RuntimeError(f"Failed to download archive: {e.reason}")

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

        title = "Success" if success else "Update failed"
        msg = MessageBox(title, message or "", self)
        msg.yesButton.setText("OK")
        msg.cancelButton.hide()
        msg.exec()

    @staticmethod
    def _build_archive_url(repo_url: str, branch: str) -> str:
        base = (repo_url or "").strip()
        if base.endswith(".git"):
            base = base[:-4]
        base = base.rstrip("/")
        return f"{base}/archive/refs/heads/{branch}.zip"

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


        # create sub interface
        self.homeInterface = QStackedWidget(self, objectName='homeInterface')
        # remove frame and let it use full space
        self.homeInterface.setFrameShape(QFrame.NoFrame)
        self.homeInterface.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.homeInterface.setStyleSheet("background: transparent;")
        self.settingsInterface = Settings(self)
        self.settingsInterface.setObjectName("settingsInterface")
        # No initial tabs; Hub shows the grid

        self.initNavigation()
        self.initWindow()
        try:
            self.settingsInterface.updateFinished.connect(self._on_settings_update)
        except Exception:
            pass
        # Keep Hub/tab selection mutually exclusive
        self.stackedWidget.currentChanged.connect(self.onContentChanged)
        # Hub active: keep last tab selection internally; hide its highlight
        # Also disable tab highlight when Hub is active
        try:
            self.titleBar.setTabsSelectionHighlightEnabled(False)
        except Exception:
            pass
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
        # Ensure clicking Hub also switches inner stack back to the grid view
        try:
            self.navigationInterface.widget(self.homeInterface.objectName()).clicked.connect(self.showHub)
        except Exception:
            pass
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
            self.packagesPage.packagesChanged.connect(lambda: self.hubGrid.refresh())
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

        self.resize(980, 727)
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

    def showHub(self):
        """Switch to the Hub (inner grid) regardless of current tab state."""
        try:
            self.switchTo(self.homeInterface)
        except Exception:
            self.stackedWidget.setCurrentWidget(self.homeInterface)
        try:
            self.homeInterface.setCurrentWidget(self.hubGrid)
        except Exception:
            pass
        try:
            self.titleBar.setTabsSelectionHighlightEnabled(False)
        except Exception:
            pass

    def showMessageBox(self):
        w = MessageBox(
            'Pycro Station',
            (
                    "Version : 0.0.1"
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
        # Switch content area to the selected macro page (robust)
        try:
            routeKey = self.tabBar.currentTab().routeKey()
        except Exception:
            return
        self._show_macro_page(routeKey)
        # When a tab is active, deselect sidebar items so the tab highlight is clear
        try:
            self._deselect_navigation()
            self.titleBar.setTabsSelectionHighlightEnabled(True)
        except Exception:
            pass

    def onTabClicked(self, index: int):
        """Handle clicks on the current tab to re-activate its content when Hub is active."""
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
        try:
            self._deselect_navigation()
            self.titleBar.setTabsSelectionHighlightEnabled(True)
        except Exception:
            pass

    def _show_macro_page(self, routeKey: str):
        """Ensure the macro page is in the stack and visible."""
        page = self.macro_pages.get(routeKey)
        if page is None:
            return
        # Ensure page is in homeInterface and visible
        try:
            if self.homeInterface.indexOf(page) == -1:
                self.homeInterface.addWidget(page)
        except Exception:
            pass
        try:
            self.stackedWidget.setCurrentWidget(self.homeInterface)
        except Exception:
            pass
        self.homeInterface.setCurrentWidget(page)
        # ensure highlight is visible and hub deselected
        try:
            self.titleBar.setTabsSelectionHighlightEnabled(True)
        except Exception:
            pass
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
                    if self.homeInterface.indexOf(page) == -1:
                        self.homeInterface.addWidget(page)
                except Exception:
                    pass
                try:
                    self.stackedWidget.setCurrentWidget(self.homeInterface)
                except Exception:
                    pass
                self.homeInterface.setCurrentWidget(page)
                self._deselect_navigation()
                try:
                    self.titleBar.setTabsSelectionHighlightEnabled(True)
                except Exception:
                    pass
                try:
                    self.packagesPage.setLocked(True)
                except Exception:
                    pass
                return

            # Replace existing page with fresh content (hot reload)
            old_page = self.macro_pages.get(routeKey)
            if old_page is not None:
                try:
                    self.homeInterface.removeWidget(old_page)
                except Exception:
                    pass
                old_page.deleteLater()
            self.macro_pages[routeKey] = content_widget
            self.macro_labels[routeKey] = text

            content_widget.setObjectName(routeKey)
            content_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.homeInterface.addWidget(content_widget)

            try:
                self.tabBar.setCurrentTab(routeKey)
            except Exception:
                pass
            try:
                self.stackedWidget.setCurrentWidget(self.homeInterface)
            except Exception:
                pass
            self.homeInterface.setCurrentWidget(content_widget)
            self._deselect_navigation()
            try:
                self.titleBar.setTabsSelectionHighlightEnabled(True)
            except Exception:
                pass
            try:
                self.packagesPage.setLocked(True)
            except Exception:
                pass
            return

        # New macro page
        content_widget.setObjectName(routeKey)
        content_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.homeInterface.addWidget(content_widget)
        self.macro_pages[routeKey] = content_widget
        self.macro_labels[routeKey] = text

        self.tabBar.addTab(routeKey, text, icon)
        try:
            self.tabBar.setCurrentTab(routeKey)
        except Exception:
            pass
        try:
            self.stackedWidget.setCurrentWidget(self.homeInterface)
        except Exception:
            pass
        self.homeInterface.setCurrentWidget(content_widget)
        self._deselect_navigation()
        try:
            self.titleBar.setTabsSelectionHighlightEnabled(True)
        except Exception:
            pass
        try:
            self.packagesPage.setLocked(True)
        except Exception:
            pass

    def onTabCloseRequested(self, index: int):
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
            # as a last resort, remove the tab and return to Hub
            try:
                self.tabBar.removeTab(index)
            except Exception:
                pass
            self.navigationInterface.setCurrentItem(self.homeInterface.objectName())
            self.stackedWidget.setCurrentWidget(self.homeInterface)
            self.homeInterface.setCurrentWidget(self.hubGrid)
            return

        # Dispose content widget
        page = self.macro_pages.pop(routeKey, None)
        self.macro_labels.pop(routeKey, None)
        if page is not None:
            try:
                self.homeInterface.removeWidget(page)
            except Exception:
                pass
            page.deleteLater()

        # Remove the tab
        try:
            self.tabBar.removeTab(index)
        except Exception:
            pass

        # Show Hub
        self.navigationInterface.setCurrentItem(self.homeInterface.objectName())
        self.stackedWidget.setCurrentWidget(self.homeInterface)
        self.homeInterface.setCurrentWidget(self.hubGrid)
        try:
            self.titleBar.setTabsSelectionHighlightEnabled(False)
        except Exception:
            pass
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

    def onContentChanged(self, index: int):
        """Ensure Hub and tabs aren't selected simultaneously."""
        w = self.stackedWidget.widget(index)
        # If homeInterface is showing the hub grid, treat it as Hub; otherwise treat as macro content
        is_hub = (w is self.homeInterface and getattr(self.homeInterface, 'currentWidget', lambda: None)() is self.hubGrid)
        if is_hub or w is self.settingsInterface or w is getattr(self, 'packagesPage', object()):
            try:
                self.titleBar.setTabsSelectionHighlightEnabled(False)
            except Exception:
                pass
        else:
            self._deselect_navigation()
            try:
                self.titleBar.setTabsSelectionHighlightEnabled(True)
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

    # removed unused text-editor helpers and tab stubs

    def _deselect_navigation(self):
        # Explicitly clear selection on NavigationBar
        try:
            for widget in getattr(self.navigationInterface, 'items', {}).values():
                widget.setSelected(False)
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
