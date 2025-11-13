import sys
from PySide6.QtWidgets import QHBoxLayout
from PySide6.QtCore import *
#from PyQt6.QtGui import QIcon
from qfluentwidgets import FluentIcon as FIF
from qfluentwidgets import *
# from TextWidget import TWidget  # Removed: no editor actions in burger menu


class CustomTitleBar(MSFluentTitleBar):

    """ Title bar with icon and title """

    def __init__(self, parent):
        super().__init__(parent)

        # add buttons
        self.toolButtonLayout = QHBoxLayout()
        color = QColor(206, 206, 206) if isDarkTheme() else QColor(96, 96, 96)
        self.menuButton = TransparentToolButton(FIF.MENU, self)
        #self.forwardButton = TransparentToolButton(FIF.RIGHT_ARROW.icon(color=color), self)
        #self.backButton = TransparentToolButton(FIF.LEFT_ARROW.icon(color=color), self)

        #self.forwardButton.setDisabled(True)
        self.toolButtonLayout.setContentsMargins(20, 0, 20, 0)
        self.toolButtonLayout.setSpacing(15)
        self.toolButtonLayout.addWidget(self.menuButton)
        #self.toolButtonLayout.addWidget(self.backButton)
        #self.toolButtonLayout.addWidget(self.forwardButton)

        self.hBoxLayout.insertLayout(4, self.toolButtonLayout)

        self.tabBar = TabBar(self)

        self.tabBar.setMovable(True)
        self.tabBar.setTabMaximumWidth(220)
        self.tabBar.setTabShadowEnabled(False)
        # remember default selected colors to allow toggling highlight visibility
        self._tabSelectedLight = QColor(255, 255, 255, 125)
        self._tabSelectedDark = QColor(255, 255, 255, 50)
        self.tabBar.setTabSelectedBackgroundColor(self._tabSelectedLight, self._tabSelectedDark)
        self.tabBar.setScrollable(True)
        self.tabBar.setCloseButtonDisplayMode(TabCloseButtonDisplayMode.ON_HOVER)

        # Delegate close handling to parent so it can dispose content widgets
        self.tabBar.tabCloseRequested.connect(parent.onTabCloseRequested)

        # Hide/disable add-tab button if present
        try:
            self.tabBar.setAddButtonVisible(False)  # newer qfluentwidgets
        except Exception:
            try:
                self.tabBar.setAddButtonEnabled(False)  # fallback
            except Exception:
                pass

        # Also react to clicking the already-selected tab to re-show content
        try:
            self.tabBar.tabBarClicked.connect(parent.onTabClicked)
        except Exception:
            pass
        # self.tabBar.currentChanged.connect(lambda i: print(self.tabBar.tabText(i)))

        self.hBoxLayout.insertWidget(5, self.tabBar, 1)
        self.hBoxLayout.setStretch(6, 0)

        # self.hBoxLayout.insertWidget(7, self.saveButton, 0, Qt.AlignmentFlag.AlignLeft)
        # self.hBoxLayout.insertWidget(7, self.openButton, 0, Qt.AlignmentFlag.AlignLeft)
        # self.hBoxLayout.insertWidget(7, self.newButton, 0, Qt.AlignmentFlag.AlignLeft)
        # self.hBoxLayout.insertSpacing(8, 20)

        # Remove all dropdown options for now; keep button but show nothing

        # Create the menuButton
        # self.menuButton = TransparentToolButton(FIF.MENU, self)
        self.menuButton.clicked.connect(self.showMenu)

    def showMenu(self):
        # Intentionally empty: burger dropdown has no options for now
        return

    # --- Helpers for Window to control tab selection highlight ---
    def setTabsSelectionHighlightEnabled(self, enabled: bool):
        try:
            if enabled:
                self.tabBar.setTabSelectedBackgroundColor(self._tabSelectedLight, self._tabSelectedDark)
            else:
                self.tabBar.setTabSelectedBackgroundColor(QColor(0, 0, 0, 0), QColor(0, 0, 0, 0))
        except Exception:
            pass

    def canDrag(self, pos: QPoint):
        if not super().canDrag(pos):
            return False

        pos.setX(pos.x() - self.tabBar.x())
        return not self.tabBar.tabRegion().contains(pos)

    def test(self):
        print("hello")
