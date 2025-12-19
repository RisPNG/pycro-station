from PySide6.QtWidgets import QHBoxLayout, QApplication
from PySide6.QtCore import *
from qfluentwidgets import FluentIcon as FIF
from qfluentwidgets import *


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

        # Enable middle-click to close tabs via event filter
        try:
            self.tabBar.installEventFilter(self)
        except Exception:
            pass

        # Global event filter as fallback (some events hit child widgets)
        try:
            QApplication.instance().installEventFilter(self)
        except Exception:
            pass

        # guard to avoid double handling
        self._lastMiddleCloseTs = 0
        self._leftPressTabIndex = -1
        self._leftPressWasCurrent = False
        # self.tabBar.currentChanged.connect(lambda i: print(self.tabBar.tabText(i)))

        self.hBoxLayout.insertWidget(5, self.tabBar, 1)
        self.hBoxLayout.setStretch(6, 0)

        # self.hBoxLayout.insertWidget(7, self.saveButton, 0, Qt.AlignmentFlag.AlignLeft)
        # self.hBoxLayout.insertWidget(7, self.openButton, 0, Qt.AlignmentFlag.AlignLeft)
        # self.hBoxLayout.insertWidget(7, self.newButton, 0, Qt.AlignmentFlag.AlignLeft)
        # Build a simple dropdown with tab operations
        self.menu = RoundMenu("Menu", self)
        self.menu.setStyleSheet("QMenu{color : red;}")

        close_current_action = Action(FIF.CLOSE, 'Close selected tab')
        close_all_action = Action(FIF.DELETE, 'Close all tabs')

        def _close_current_tab():
            try:
                idx = self.tabBar.currentIndex()
                if idx >= 0:
                    parent.onTabCloseRequested(idx)
            except Exception:
                pass

        def _close_all_tabs():
            try:
                parent.closeAllTabs()
            except Exception:
                # fallback: remove directly
                try:
                    cnt = self.tabBar.count()
                except Exception:
                    cnt = 0
                for i in reversed(range(cnt)):
                    try:
                        parent.onTabCloseRequested(i)
                    except Exception:
                        pass

        close_current_action.triggered.connect(_close_current_tab)
        close_all_action.triggered.connect(_close_all_tabs)

        self.menu.addAction(close_current_action)
        self.menu.addAction(close_all_action)

        # Create the menuButton
        self.menuButton.clicked.connect(self.showMenu)

    def showMenu(self):
        # Show the prebuilt menu at the bottom-left of the button
        self.menu.exec(self.menuButton.mapToGlobal(self.menuButton.rect().bottomLeft()))

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

    def _tabIndexAtGlobalPos(self, gpos: QPoint) -> int:
        tb = self.tabBar
        try:
            local = tb.mapFromGlobal(gpos)
            if not tb.rect().contains(local):
                return -1
        except Exception:
            return -1

        # Try built-in hit test methods first
        for name in (
            'tabAt', 'indexAt', 'tabIndexAt', 'tabIndexFromPosition',
            'tabIndexAtPos', 'tabIndexAtPosition', 'tabAtPos'
        ):
            meth = getattr(tb, name, None)
            if callable(meth):
                try:
                    idx = meth(local)
                    if isinstance(idx, int) and idx >= 0:
                        return idx
                except Exception:
                    continue

        # Fallback: check per-tab rects, then region approximation
        count = 0
        try:
            count = tb.count()
        except Exception:
            count = 0

        if count > 0:
            # Try tabRect geometry per index
            for name in ('tabRect', 'tabGeometry'):
                geom = getattr(tb, name, None)
                if not callable(geom):
                    continue
                for i in range(count):
                    try:
                        r = geom(i)
                        if r.contains(local):
                            return i
                    except Exception:
                        continue

            # Approximate by dividing tabRegion if still not found
            try:
                region = tb.tabRegion()
                if region.contains(local):
                    relx = local.x() - region.x()
                    w = region.width() if region.width() > 0 else 1
                    width_per = w / count
                    approx = int(relx / width_per)
                    if 0 <= approx < count:
                        return approx
            except Exception:
                pass

        return -1

    def eventFilter(self, obj, e):
        try:
            # Track left click on the already-selected tab to re-show content (without duplicating tab switch logic)
            if e.type() == QEvent.MouseButtonPress and hasattr(e, 'button') and e.button() == Qt.LeftButton:
                gpos = e.globalPosition().toPoint() if hasattr(e, 'globalPosition') else e.globalPos()
                idx = self._tabIndexAtGlobalPos(gpos)
                if isinstance(idx, int) and idx >= 0:
                    self._leftPressTabIndex = idx
                    try:
                        self._leftPressWasCurrent = idx == self.tabBar.currentIndex()
                    except Exception:
                        self._leftPressWasCurrent = False
                return super().eventFilter(obj, e)

            if e.type() == QEvent.MouseButtonRelease and hasattr(e, 'button') and e.button() == Qt.LeftButton:
                gpos = e.globalPosition().toPoint() if hasattr(e, 'globalPosition') else e.globalPos()
                idx = self._tabIndexAtGlobalPos(gpos)
                if self._leftPressWasCurrent and idx == self._leftPressTabIndex and isinstance(idx, int) and idx >= 0:
                    try:
                        self.parent().onTabClicked(idx)
                    except Exception:
                        pass
                self._leftPressTabIndex = -1
                self._leftPressWasCurrent = False
                return super().eventFilter(obj, e)

            # Handle middle-click close on the tab bar: close only the tab under cursor
            if e.type() == QEvent.MouseButtonRelease and hasattr(e, 'button') and e.button() == Qt.MiddleButton:
                # throttle to avoid duplicate from multiple filters
                now = QDateTime.currentMSecsSinceEpoch()
                if now - self._lastMiddleCloseTs < 150:
                    return True
                gpos = e.globalPosition().toPoint() if hasattr(e, 'globalPosition') else e.globalPos()
                index = self._tabIndexAtGlobalPos(gpos)
                if isinstance(index, int) and index >= 0:
                    self._lastMiddleCloseTs = now
                    try:
                        self.parent().onTabCloseRequested(index)
                    except Exception:
                        pass
                    return True
        except Exception:
            pass
        return super().eventFilter(obj, e)
