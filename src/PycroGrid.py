import os
import sys
import re
import html
import json
import importlib.util
from difflib import SequenceMatcher
from importlib import metadata

from PySide6.QtCore import Qt, QFileSystemWatcher, QTimer, QProcess, QEvent, QSize, QRect, QPoint, Signal, QRectF
from PySide6.QtWidgets import *
from PySide6.QtGui import QIcon, QCursor, QAction, QPainter
from qfluentwidgets import (
    PrimaryPushButton,
    TransparentToolButton,
    isDarkTheme,
    FluentIcon as FIF,
    LineEdit,
    RoundMenu,
)
from pytablericons import OutlineIcon, FilledIcon
from PackagesPage import CheckIconButton, ti_icon


class ClickableMenuRow(QWidget):
    clicked = Signal()

    def mousePressEvent(self, event):
        try:
            if event.button() == Qt.LeftButton:
                self.clicked.emit()
        except Exception:
            pass
        return super().mousePressEvent(event)


class IconOffsetToolButton(TransparentToolButton):
    def __init__(self, *args, icon_offset_y: int = 0, **kwargs):
        super().__init__(*args, **kwargs)
        self._icon_offset_y = int(icon_offset_y)

    def setIconOffsetY(self, offset_y: int):
        self._icon_offset_y = int(offset_y)
        self.update()

    def paintEvent(self, event):
        QToolButton.paintEvent(self, event)
        icon = getattr(self, "_icon", None)
        if icon is None:
            return

        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)

        if not self.isEnabled():
            painter.setOpacity(0.43)
        elif getattr(self, "isPressed", False):
            painter.setOpacity(0.63)

        w, h = self.iconSize().width(), self.iconSize().height()
        x = (self.width() - w) / 2
        y = (self.height() - h) / 2 + self._icon_offset_y
        self._drawIcon(icon, painter, QRectF(x, y, w, h))


class PycroGrid(QScrollArea):
    """Scrollable grid container for listing Pycros."""

    def __init__(self, parent=None, stars_only: bool = False):
        super().__init__(parent)

        self._stars_only = bool(stars_only)
        self._settings_file = os.path.join(os.path.dirname(__file__), "settings.json")
        self._sort_mode = "recently_used"
        self._show_remote_pycros = True
        self._recently_launched: list[str] = []
        self._starred_pycros: set[str] = set()
        self._settings_watcher = QFileSystemWatcher(self)
        self._settings_watcher.fileChanged.connect(self._on_settings_file_changed)
        try:
            if os.path.isfile(self._settings_file):
                self._settings_watcher.addPath(self._settings_file)
        except Exception:
            pass
        self._reload_preferences_from_disk(apply_filter=False)
        icon_color = "#FFFFFF" if isDarkTheme() else "#111111"
        self._star_outline_icon = ti_icon(OutlineIcon.STAR, size=20, color=icon_color, stroke_width=2.0)
        self._star_filled_icon = ti_icon(FilledIcon.STAR, size=20, color=icon_color, stroke_width=2.0)

        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setStyleSheet("QScrollArea{background: transparent; border: none;}")
        self.viewport().setStyleSheet("background: transparent;")

        self._content = QWidget(self)
        self._content.setStyleSheet("background: transparent;")
        self._content_layout = QVBoxLayout(self._content)
        self._content_layout.setContentsMargins(0, 0, 0, 0)
        self._content_layout.setSpacing(8)

        # Search bar
        search_row = QHBoxLayout()
        self._search_row = search_row
        search_row.setContentsMargins(12, 12, 12, 0)
        search_row.setSpacing(8)
        self._search_field = LineEdit(self._content)
        self._search_field.setPlaceholderText("Search pycros...")
        search_row.addWidget(self._search_field)

        self._filter_menu = RoundMenu("", self)
        self._filter_menu.setMinimumWidth(220)
        self._sort_cycle_action = QAction("", self)
        self._sort_cycle_action.triggered.connect(self._cycle_sort_mode)
        self._filter_menu.addAction(self._sort_cycle_action)
        self._filter_menu.addSeparator()

        self._remote_toggle = None
        self._remote_row = None
        self._build_remote_toggle_row()

        self._filter_action = None
        self._filter_btn = None
        self._install_filter_icon()
        self._content_layout.addLayout(search_row)

        # Grid host
        self._grid_host = QWidget(self._content)
        self._grid = QGridLayout(self._grid_host)
        self._grid.setContentsMargins(12, 12, 12, 12)
        self._grid.setSpacing(12)
        self._content_layout.addWidget(self._grid_host, 1)
        self.setWidget(self._content)

        self._cards: list[PycroCard] = []
        self._watcher = QFileSystemWatcher(self)
        self._watcher.directoryChanged.connect(self._on_dir_changed)
        self._watcher.fileChanged.connect(self._on_file_changed)
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.setInterval(250)
        self._debounce_timer.timeout.connect(self.refresh)
        self._last_changed_path = None
        self._info_popup: QLabel | None = None
        self._info_hover_timer = QTimer(self)
        self._info_hover_timer.setSingleShot(True)
        self._info_hover_timer.timeout.connect(self._show_pending_hover_popup)
        self._pending_callout: tuple[QWidget, str] | None = None
        self._last_infos: list['PycroInfo'] = []
        self._all_infos: list['PycroInfo'] = []
        self._last_column_count: int | None = None
        self._resize_relayout_timer = QTimer(self)
        self._resize_relayout_timer.setSingleShot(True)
        self._resize_relayout_timer.setInterval(120)
        self._resize_relayout_timer.timeout.connect(self._relayout_on_resize)

        self._empty_label = QLabel("Pycros will appear here", self._content)
        if self._stars_only:
            self._empty_label.setText("No starred pycros yet")
        self._empty_label.setAlignment(Qt.AlignCenter)
        self._empty_label.setStyleSheet("color: #aaa; font-size: 14px;")
        self._grid.addWidget(self._empty_label, 0, 0)
        self._grid.setRowStretch(0, 1)
        self._grid.setColumnStretch(0, 1)

        self._roots = [
            os.path.join(os.getcwd(), 'pycros'),
            os.path.join(os.getcwd(), 'remote_pycros'),
        ]
        for root in self._roots:
            if os.path.isdir(root):
                self._watcher.addPath(root)
        self._invalid_pycros: set[str] = set()

        try:
            QToolTip.setStyleSheet(
                "QToolTip {"
                " color: #f5f5f5;"
                " background-color: rgba(44,44,44,0.95);"
                " border: 1px solid rgba(85,85,85,0.9);"
                " border-radius: 8px;"
                " padding: 10px;"
                " }"
            )
        except Exception:
            pass
        self._search_field.textChanged.connect(lambda _: self._apply_filter())
        self._sync_filter_menu_state()

        QTimer.singleShot(0, self.refresh)

    def _read_settings(self) -> dict:
        try:
            if os.path.isfile(self._settings_file):
                with open(self._settings_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        return data
        except Exception:
            pass
        return {}

    def _sort_mode_settings_key(self) -> str:
        return "stars_sort_mode" if self._stars_only else "hub_sort_mode"

    def _show_remote_settings_key(self) -> str:
        return "stars_show_remote_pycros" if self._stars_only else "hub_show_remote_pycros"

    def _write_settings_updates(self, updates: dict) -> dict:
        data = self._read_settings()
        data.update(updates)
        try:
            with open(self._settings_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception:
            pass
        return data

    def _reload_preferences_from_disk(self, apply_filter: bool = True):
        data = self._read_settings()

        sort_mode = data.get(self._sort_mode_settings_key(), "recently_used")
        if sort_mode not in ("recently_used", "alphabetical"):
            sort_mode = "recently_used"
        self._sort_mode = sort_mode

        show_remote = data.get(self._show_remote_settings_key(), None)
        if show_remote is None:
            # backward compatibility with older single setting
            show_remote = data.get("show_remote_pycros", True)
        self._show_remote_pycros = bool(show_remote)

        recently = data.get("recently_launched", [])
        self._recently_launched = recently if isinstance(recently, list) else []

        starred = data.get("starred_pycros", [])
        self._starred_pycros = set(starred) if isinstance(starred, list) else set()

        self._sync_filter_menu_state()
        if apply_filter:
            self._apply_filter()

    def _on_settings_file_changed(self, path: str):
        # On some platforms, a file watch is removed after change; re-add it.
        try:
            if os.path.isfile(self._settings_file) and self._settings_file not in self._settings_watcher.files():
                self._settings_watcher.addPath(self._settings_file)
        except Exception:
            pass
        self._reload_preferences_from_disk(apply_filter=True)

    def _sync_filter_menu_state(self):
        try:
            mode_label = "Alphabetical" if self._sort_mode == "alphabetical" else "Recently used"
            self._sort_cycle_action.setText(f"Sort: {mode_label}")
        except Exception:
            pass
        try:
            if self._remote_toggle is not None:
                self._remote_toggle.setChecked(bool(self._show_remote_pycros))
        except Exception:
            pass

    def _cycle_sort_mode(self):
        next_mode = "alphabetical" if self._sort_mode == "recently_used" else "recently_used"
        self._set_sort_mode(next_mode)

    def _install_filter_icon(self):
        # Prefer embedding the filter icon inside the LineEdit trailing position.
        icon = QIcon()
        try:
            icon = FIF.FILTER.icon()
        except Exception:
            try:
                icon = QIcon(FIF.FILTER)
            except Exception:
                icon = QIcon()

        try:
            pos = getattr(QLineEdit, "TrailingPosition", None)
            if pos is None:
                pos = QLineEdit.ActionPosition.TrailingPosition
            self._filter_action = self._search_field.addAction(icon, pos)
            self._filter_action.triggered.connect(self._show_filter_menu)
            return
        except Exception:
            self._filter_action = None

        # Fallback: manually place a tool button inside the LineEdit
        try:
            self._filter_btn = TransparentToolButton(FIF.FILTER, self._search_field)
            self._filter_btn.setCursor(Qt.PointingHandCursor)
            self._filter_btn.setStyleSheet("QToolButton{border:none;}")
            self._filter_btn.clicked.connect(self._show_filter_menu)
            try:
                self._search_field.installEventFilter(self)
            except Exception:
                pass
            self._position_filter_btn()
        except Exception:
            self._filter_btn = None

    def _position_filter_btn(self):
        if self._filter_btn is None or self._filter_btn.parent() is not self._search_field:
            return
        try:
            h = max(20, int(self._search_field.height()))
            margin = 4
            btn_size = max(20, min(28, h - (margin * 2)))
            self._filter_btn.setFixedSize(btn_size, btn_size)
            self._filter_btn.setIconSize(QSize(max(12, btn_size - 10), max(12, btn_size - 10)))
            x = max(0, self._search_field.width() - btn_size - margin)
            y = max(0, (h - btn_size) // 2)
            self._filter_btn.move(x, y)
            try:
                self._filter_btn.raise_()
            except Exception:
                pass
            try:
                self._search_field.setTextMargins(0, 0, btn_size + (margin * 2), 0)
            except Exception:
                pass
        except Exception:
            pass

    def _show_filter_menu(self):
        try:
            anchor = self._filter_btn if self._filter_btn is not None else self._search_field
        except Exception:
            anchor = self._search_field

        try:
            self._filter_menu.ensurePolished()
        except Exception:
            pass

        try:
            menu_size = self._filter_menu.sizeHint()
            if menu_size.width() <= 0 or menu_size.height() <= 0:
                self._filter_menu.adjustSize()
                menu_size = self._filter_menu.sizeHint()
        except Exception:
            menu_size = None

        try:
            anchor_br = anchor.mapToGlobal(anchor.rect().bottomRight())
            anchor_tr = anchor.mapToGlobal(anchor.rect().topRight())
        except Exception:
            anchor_br = QCursor.pos()
            anchor_tr = anchor_br

        x = anchor_br.x()
        y = anchor_br.y()
        if menu_size is not None:
            x = x - menu_size.width()
            y = y + 2

        # Clamp to the app window so the menu stays "inside" the window bounds.
        try:
            w = self.window()
            if w is not None:
                tl = w.mapToGlobal(w.rect().topLeft())
                br = w.mapToGlobal(w.rect().bottomRight())
                win = QRect(tl, br)
            else:
                win = None
        except Exception:
            win = None

        if win is not None and menu_size is not None and win.isValid():
            margin = 8
            max_x = win.right() - menu_size.width() - margin
            min_x = win.left() + margin
            x = max(min_x, min(x, max_x))

            # Prefer below; if it doesn't fit, place above.
            if y + menu_size.height() > win.bottom() - margin:
                y = anchor_tr.y() - menu_size.height() - 2
            max_y = win.bottom() - menu_size.height() - margin
            min_y = win.top() + margin
            y = max(min_y, min(y, max_y))

        pos = QPoint(int(x), int(y))
        try:
            self._filter_menu.exec(pos)
        except Exception:
            try:
                self._filter_menu.exec_(pos)
            except Exception:
                pass

    def _build_remote_toggle_row(self):
        if self._remote_toggle is not None or self._remote_row is not None:
            return

        row = ClickableMenuRow(self._filter_menu)
        try:
            row.setCursor(Qt.PointingHandCursor)
        except Exception:
            pass
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 4, 10, 4)
        h.setSpacing(0)

        self._remote_toggle = CheckIconButton(row, initially_checked=bool(self._show_remote_pycros))
        try:
            self._remote_toggle.setFixedSize(24, 24)
            self._remote_toggle.setIconSize(QSize(15, 15))
        except Exception:
            pass
        self._remote_toggle.setToolTip("Show/hide remote pycros")
        label = QLabel("Show remote pycros", row)
        label.setStyleSheet(
            "color:#dcdcdc; background:transparent; font-size:13px; font-weight:500;"
        )
        try:
            label.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        except Exception:
            pass

        h.addWidget(self._remote_toggle, 0, Qt.AlignVCenter)
        h.addWidget(label, 0, Qt.AlignVCenter)
        h.addStretch(1)

        try:
            row.setFixedWidth(max(220, int(self._filter_menu.minimumWidth())))
        except Exception:
            row.setFixedWidth(220)
        try:
            row.setFixedHeight(max(32, int(row.sizeHint().height())))
        except Exception:
            row.setFixedHeight(32)

        try:
            self._filter_menu.addWidget(row, selectable=False)
        except Exception:
            # fallback: at least try to add as a regular action (won't embed widget)
            try:
                action = QAction("Show remote pycros", self)
                action.setCheckable(True)
                action.setChecked(bool(self._show_remote_pycros))
                action.triggered.connect(lambda c: self._on_show_remote_toggled(bool(c)))
                self._filter_menu.addAction(action)
            except Exception:
                pass
        self._remote_row = row

        try:
            self._remote_toggle.toggledManually.connect(self._on_show_remote_toggled)
        except Exception:
            pass
        try:
            row.clicked.connect(lambda: self._remote_toggle.click())
        except Exception:
            pass

    def _on_show_remote_toggled(self, checked: bool):
        self._show_remote_pycros = bool(checked)
        self._write_settings_updates({self._show_remote_settings_key(): self._show_remote_pycros})
        self._refresh_all_grids()

    def _set_sort_mode(self, mode: str):
        if mode not in ("recently_used", "alphabetical"):
            mode = "recently_used"
        if self._sort_mode == mode:
            return
        self._sort_mode = mode
        self._write_settings_updates({self._sort_mode_settings_key(): self._sort_mode})
        self._refresh_all_grids()

    def _refresh_all_grids(self):
        """Refresh Hub + Stars grids (in this process) after settings changes."""
        window = self.window()
        if window is None:
            self._reload_preferences_from_disk(apply_filter=True)
            return

        for attr in ("hubGrid", "starsGrid"):
            grid = getattr(window, attr, None)
            if grid is None:
                continue
            reload_fn = getattr(grid, "_reload_preferences_from_disk", None)
            if callable(reload_fn):
                try:
                    reload_fn(apply_filter=True)
                except Exception:
                    continue

    def record_launch(self, info: 'PycroInfo'):
        """Record a successful launch for 'recently used' sorting."""
        pid = self._pycro_id(info)
        recent = [x for x in self._recently_launched if x not in (pid, info.name)]
        recent.insert(0, pid)
        # keep the list bounded
        if len(recent) > 200:
            recent = recent[:200]
        self._recently_launched = recent
        self._write_settings_updates({"recently_launched": list(self._recently_launched)})
        self._refresh_all_grids()

    def toggle_star(self, info: 'PycroInfo'):
        pid = self._pycro_id(info)
        currently_starred = pid in self._starred_pycros or info.name in self._starred_pycros
        self.set_starred(info, not currently_starred)

    def set_starred(self, info: 'PycroInfo', starred: bool):
        pid = self._pycro_id(info)
        # Remove any legacy/bare-name entries first
        self._starred_pycros.discard(info.name)
        self._starred_pycros.discard(pid)
        if starred:
            self._starred_pycros.add(pid)
        self._write_settings_updates({"starred_pycros": sorted(self._starred_pycros)})
        self._refresh_all_grids()

    def refresh(self):
        infos = self._scan_pycros()
        self._all_infos = infos
        self._apply_filter()
        self._reload_open_tabs(infos)

    def _scan_pycros(self):
        infos = []
        roots = [r for r in self._roots if os.path.isdir(r)]
        if not roots:
            return infos

        try:
            dirs = self._watcher.directories()
            files = self._watcher.files()
            if dirs:
                self._watcher.removePaths(dirs)
            if files:
                self._watcher.removePaths(files)
        except Exception:
            pass

        for root in roots:
            try:
                self._watcher.addPath(root)
            except Exception:
                pass
            try:
                subdirs = sorted([d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))])
            except Exception:
                subdirs = []

            is_remote = os.path.basename(root) == 'remote_pycros'

            for d in subdirs:
                folder = os.path.join(root, d)
                main_py = os.path.join(folder, 'main.py')
                req_txt = os.path.join(folder, 'requirements.txt')
                desc_md = os.path.join(folder, 'description.md')
                has_python = False
                try:
                    for fn in os.listdir(folder):
                        if fn.lower().endswith('.py'):
                            has_python = True
                            break
                except Exception:
                    has_python = False

                short_desc, long_desc, info_lines = self._parse_description(desc_md)
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
                    info_lines=info_lines,
                    has_python=has_python,
                    is_remote=is_remote,
                )

                infos.append(info)

                if os.path.isfile(desc_md):
                    self._watcher.addPath(desc_md)
                if os.path.isfile(req_txt):
                    self._watcher.addPath(req_txt)
                if os.path.isfile(main_py):
                    self._watcher.addPath(main_py)

        return infos

    def _apply_filter(self):
        query = (self._search_field.text() or "").strip()
        filtered = list(self._all_infos)
        if not self._show_remote_pycros:
            filtered = [info for info in filtered if not info.is_remote]
        if self._stars_only:
            filtered = [info for info in filtered if self._is_starred(info)]
        if query:
            filtered = [info for info in filtered if self._matches_query(info, query)]
        filtered = self._sort_infos(filtered)
        self._rebuild(filtered)

    @staticmethod
    def _pycro_id(info: 'PycroInfo') -> str:
        prefix = "remote:" if getattr(info, "is_remote", False) else "local:"
        return prefix + (getattr(info, "name", "") or "")

    def _is_starred(self, info: 'PycroInfo') -> bool:
        pid = self._pycro_id(info)
        return pid in self._starred_pycros or info.name in self._starred_pycros

    def _sort_infos(self, infos: list['PycroInfo']) -> list['PycroInfo']:
        if self._sort_mode == "alphabetical":
            return sorted(infos, key=lambda i: (i.display_name or i.name).lower())

        order: dict[str, int] = {}
        for idx, pid in enumerate(self._recently_launched):
            if isinstance(pid, str) and pid not in order:
                order[pid] = idx

        def sort_key(i: 'PycroInfo'):
            pid = self._pycro_id(i)
            rank = order.get(pid)
            if rank is None:
                rank = order.get(i.name)
            if rank is None:
                rank = 10**9
            return (rank, (i.display_name or i.name).lower())

        return sorted(infos, key=sort_key)

    def _matches_query(self, info: 'PycroInfo', query: str) -> bool:
        q = query.lower()
        name = info.display_name.lower()
        desc = (info.short_desc or "").lower()
        # quick substring match
        if q in name or q in desc:
            return True
        def ratio(a: str, b: str) -> float:
            return SequenceMatcher(None, a, b).ratio() if a and b else 0.0
        return max(ratio(q, name), ratio(q, desc)) >= 0.8

    def _parse_description(self, path):
        short_lines: list[str] = []
        info_lines: list[str] = []
        long_desc_lines: list[str] = []
        if not os.path.isfile(path):
            return '', '', []
        try:
            with open(path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            info_mode = False
            for raw in lines:
                s = raw.rstrip('\n')
                if s.strip().lower().startswith('> [!info]'):
                    info_mode = True
                    continue
                if info_mode and s.strip() == "":
                    info_lines.append("")
                    continue
                if s.startswith('>'):
                    content = s.lstrip('> ').strip()
                    if info_mode:
                        info_lines.append(content)
                    else:
                        short_lines.append(content)
                else:
                    long_desc_lines.append(s)
        except Exception:
            pass
        short_desc = '\n'.join([ln for ln in short_lines if ln]).strip()
        long_desc = '\n'.join(long_desc_lines).strip()
        return short_desc, long_desc, info_lines

    def _rebuild(self, infos, rebuild_cards: bool = True):
        self._last_infos = infos
        if rebuild_cards:
            for card in self._cards:
                self._grid.removeWidget(card)
                card.deleteLater()
            self._cards.clear()

            for info in infos:
                try:
                    card = PycroCard(info, parent=self)
                except Exception as e:
                    try:
                        print(f"Failed to build card for '{getattr(info, 'name', '?')}': {e}")
                    except Exception:
                        pass
                    continue
                if info.name in self._invalid_pycros:
                    try:
                        card.set_invalid(True)
                    except Exception:
                        pass
                self._cards.append(card)

        self._relayout_cards(force=True)

    def _on_dir_changed(self, path):
        self._last_changed_path = path
        self._debounce_timer.start()

    def _on_file_changed(self, path):
        self._last_changed_path = path
        self._debounce_timer.start()

    def _reload_open_tabs(self, infos):
        if not self._last_changed_path:
            return
        window = self.window()
        if window is None or not hasattr(window, 'macro_pages'):
            self._last_changed_path = None
            return

        changed = os.path.normcase(self._last_changed_path)
        targets = [
            info for info in infos
            if changed.startswith(os.path.normcase(info.folder))
        ]
        if not targets:
            self._last_changed_path = None
            return

        for info in targets:
            try:
                open_tabs = getattr(window, 'macro_pages', {})
                if info.name not in open_tabs:
                    continue
                page = self._build_page(info)
                if page is None:
                    self._invalid_pycros.add(info.name)
                    try:
                        for card in self._cards:
                            if card.info.name == info.name:
                                card.set_invalid(True)
                                break
                    except Exception:
                        pass
                    try:
                        tab_bar = getattr(window, 'tabBar', None)
                        macro_labels = getattr(window, 'macro_labels', {})
                        label = macro_labels.get(info.name)
                        if tab_bar is not None and label is not None:
                            count = tab_bar.count()
                            for i in range(count):
                                try:
                                    if tab_bar.tabText(i) == label:
                                        window.onTabCloseRequested(i)
                                        break
                                except Exception:
                                    continue
                    except Exception:
                        pass
                    continue
                self._invalid_pycros.discard(info.name)
                try:
                    for card in self._cards:
                        if card.info.name == info.name:
                            card.set_invalid(False)
                            break
                except Exception:
                    pass
                window.addMacroTab(info.name, info.display_name, QIcon(), page, replace_existing=True)
            except Exception:
                continue

        self._last_changed_path = None

    def _build_page(self, info: 'PycroInfo') -> QWidget | None:
        widget = _load_pycro_widget(info)
        if widget is None:
            return None

        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(12)

        try:
            if hasattr(widget, "set_long_description"):
                widget.set_long_description(info.long_desc)
            elif info.long_desc:
                widget.setProperty("pycro_long_desc", info.long_desc)
        except Exception:
            pass

        v.addWidget(widget, 1)
        return page

    @staticmethod
    def _format_info_tooltip(lines: list[str]) -> str:
        if not lines:
            lines = ["No additional info"]
        html_lines = []
        for ln in lines:
            if ln:
                html_lines.append(f"<div style='margin:0; padding:0 0 6px 0'>{PycroGrid._render_colored_text(ln)}</div>")
            else:
                html_lines.append("<div style='margin:0; padding:8px 0'>&nbsp;</div>")
        body = "".join(html_lines)
        return (
            "<div style='color:#f5f5f5; padding:6px 6px 2px 6px;'>"
            f"{body}"
            "</div>"
        )

    @staticmethod
    def _render_colored_text(text: str) -> str:
        """Render [text](#RRGGBB[AA]) into colored spans, escaping other content."""
        if not text:
            return ""
        pattern = re.compile(r"\[([^\]]+)\]\(#([0-9A-Fa-f]{6})([0-9A-Fa-f]{2})?\)")
        parts = []
        last = 0
        for m in pattern.finditer(text):
            parts.append(html.escape(text[last:m.start()]))
            label = html.escape(m.group(1))
            hex_part = m.group(2)
            alpha_part = m.group(3)
            r = int(hex_part[0:2], 16)
            g = int(hex_part[2:4], 16)
            b = int(hex_part[4:6], 16)
            if alpha_part:
                a = int(alpha_part, 16) / 255.0
                color_css = f"rgba({r},{g},{b},{a:.2f})"
            else:
                color_css = f"rgb({r},{g},{b})"
            parts.append(f"<span style='color:{color_css};'>{label}</span>")
            last = m.end()
        parts.append(html.escape(text[last:]))
        return "".join(parts)

    def _format_desc_html(self, text: str) -> str:
        if not text:
            return "(no description)"
        lines = text.splitlines() or [text]
        rendered = []
        for ln in lines:
            rendered.append(self._render_colored_text(ln) or "&nbsp;")
        return "<br>".join(rendered)

    def _show_info_popup(self, html_text: str, anchor: QWidget | None, duration: int = 3500):
        """Show a custom popup for info text to avoid native tooltip styling issues."""
        try:
            if self._info_popup is not None:
                self._info_popup.close()
        except Exception:
            pass
        self._info_popup = QLabel()
        self._info_popup.setWindowFlags(Qt.ToolTip)
        self._info_popup.setAttribute(Qt.WA_ShowWithoutActivating)
        self._info_popup.setStyleSheet(
            "QLabel{"
            "color:#f5f5f5;"
            "background-color: rgba(44,44,44,0.95);"
            "border:1px solid rgba(85,85,85,0.9);"
            "border-radius:8px;"
            "padding:10px;"
            "}"
        )
        self._info_popup.setText(html_text)
        self._info_popup.setTextFormat(Qt.RichText)
        self._info_popup.adjustSize()

        pos = QCursor.pos()
        try:
            if anchor is not None:
                pos = anchor.mapToGlobal(anchor.rect().bottomRight())
        except Exception:
            pass

        self._info_popup.move(pos)
        self._info_popup.show()
        if duration > 0:
            QTimer.singleShot(duration, lambda: self._hide_info_popup())

    def _hide_info_popup(self):
        try:
            if self._info_popup is not None:
                self._info_popup.close()
        except Exception:
            pass
        self._info_popup = None

    def _show_pending_hover_popup(self):
        if self._pending_callout is None:
            return
        anchor, txt = self._pending_callout
        self._show_info_popup(txt, anchor, duration=0)

    def eventFilter(self, obj, event):
        if obj is getattr(self, "_search_field", None) and event.type() in (QEvent.Resize, QEvent.Show):
            self._position_filter_btn()
        # Handle hover over info buttons to show popup quickly
        if isinstance(obj, QWidget) and obj.property("callout_text") is not None:
            if event.type() == QEvent.Enter:
                self._pending_callout = (obj, obj.property("callout_text"))
                self._info_hover_timer.start(200)  # shorter delay
            elif event.type() in (QEvent.Leave, QEvent.HoverLeave):
                self._info_hover_timer.stop()
                self._pending_callout = None
                self._hide_info_popup()
        return super().eventFilter(obj, event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._resize_relayout_timer.start()

    def _relayout_on_resize(self):
        if not self._cards and not self._last_infos:
            return
        self._relayout_cards()

    def _clear_grid_items(self):
        while self._grid.count():
            self._grid.takeAt(0)

    def _card_width_hint(self) -> int:
        if not self._cards:
            return 300
        card = self._cards[0]
        width = max(card.sizeHint().width(), card.minimumWidth())
        max_width = card.maximumWidth()
        if max_width > 0:
            width = min(width, max_width)
        return width

    def _compute_columns(self) -> int:
        spacing = self._grid.horizontalSpacing()
        if spacing is None or spacing < 0:
            spacing = self._grid.spacing()
        spacing = spacing if spacing is not None and spacing >= 0 else 0

        margins = self._grid.contentsMargins()
        available_width = self.viewport().width() - margins.left() - margins.right()
        card_width = max(1, self._card_width_hint())
        raw_columns = (available_width + spacing) // (card_width + spacing)
        columns = max(1, int(raw_columns))
        max_columns = len(self._cards) if self._cards else 1
        return min(columns, max_columns)

    def _relayout_cards(self, force: bool = False):
        if not self._cards:
            self._clear_grid_items()
            self._empty_label.show()
            self._grid.addWidget(self._empty_label, 0, 0)
            self._grid.setRowStretch(0, 1)
            self._grid.setColumnStretch(0, 1)
            self._last_column_count = None
            return

        columns = self._compute_columns()
        if not force and self._last_column_count == columns:
            return

        self._clear_grid_items()
        self._empty_label.hide()

        row = 0
        col = 0
        for card in self._cards:
            self._grid.addWidget(card, row, col)
            col += 1
            if col >= columns:
                col = 0
                row += 1

        for r in range(row + 1):
            self._grid.setRowStretch(r, 0)
        self._grid.setRowStretch(row + 1, 1)
        for c in range(columns):
            self._grid.setColumnStretch(c, 0)
        self._grid.setColumnStretch(columns, 1)
        self._last_column_count = columns


class PycroInfo:
    def __init__(self, name, display_name, folder, main_py, requirements, description, short_desc, long_desc, info_lines, has_python: bool, is_remote: bool = False):
        self.name = name
        self.display_name = display_name
        self.folder = folder
        self.main_py = main_py
        self.requirements = requirements
        self.description = description
        self.short_desc = short_desc
        self.long_desc = long_desc
        self.info_lines = info_lines
        self.has_python = has_python
        self.is_remote = is_remote


def _load_pycro_widget(info: 'PycroInfo') -> QWidget | None:
    main_path = info.main_py
    if not os.path.isfile(main_path):
        return None
    try:
        module_name = f"pycro_{info.name}"
        if module_name in sys.modules:
            del sys.modules[module_name]

        spec = importlib.util.spec_from_file_location(module_name, main_path)
        if spec is None or spec.loader is None:
            return None
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)  # type: ignore
        if hasattr(module, 'get_widget') and callable(module.get_widget):
            w = module.get_widget()
            if isinstance(w, QWidget):
                return w
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


class PycroCard(QWidget):
    def __init__(self, info: 'PycroInfo', parent=None):
        super().__init__(parent)
        self.info = info
        self._grid: PycroGrid = parent
        self._invalid = False

        safe_name = re.sub(r'[^A-Za-z0-9_]', '_', info.name)
        self.setObjectName(f"card__{safe_name}")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setAutoFillBackground(True)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.setMinimumSize(280, 210)
        self.setMaximumSize(320, 250)
        bg = "#242424" if isDarkTheme() else "#F2F2F2"
        border = "#3a7bd5" if info.is_remote else ("rgba(255,255,255,0.08)" if isDarkTheme() else "rgba(0,0,0,0.08)")
        self.setStyleSheet(
            f"QWidget#{self.objectName()}{{background-color:{bg}; border:1px solid {border}; border-radius:8px;}}"
        )

        v = QVBoxLayout(self)
        v.setContentsMargins(12, 12, 12, 12)
        v.setSpacing(8)

        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.setSpacing(6)

        title = QLabel(info.display_name, self)
        title.setWordWrap(True)
        title.setStyleSheet(
            f"background: transparent; border: none; color: {'#fff' if isDarkTheme() else '#111'}; font-size:16px; font-weight:600;"
        )
        title_row.addWidget(title, 1)

        self.star_btn = IconOffsetToolButton(self, icon_offset_y=1)
        self.star_btn.setFixedSize(24, 24)
        self.star_btn.setIconSize(QSize(20, 20))
        self.star_btn.setCursor(Qt.PointingHandCursor)
        self.star_btn.setStyleSheet("QToolButton{border:none;}")
        self.star_btn.clicked.connect(self._on_star_clicked)
        self._sync_star_icon()
        title_row.addWidget(self.star_btn, 0, Qt.AlignVCenter | Qt.AlignRight)

        if info.info_lines:
            info_btn = TransparentToolButton(FIF.INFO, self)
            info_btn.setFixedSize(24, 24)
            info_btn.setCursor(Qt.PointingHandCursor)
            info_btn.setStyleSheet("QToolButton{border:none;}")
            tooltip_html = self._grid._format_info_tooltip(info.info_lines)
            info_btn.setProperty("callout_text", tooltip_html)
            info_btn.installEventFilter(self._grid)
            info_btn.clicked.connect(lambda _, txt=tooltip_html: self._grid._show_info_popup(txt, info_btn, duration=3500))
            title_row.addWidget(info_btn, 0, Qt.AlignVCenter | Qt.AlignRight)

        title_row.addStretch(0)

        desc = QTextBrowser(self)
        desc.setReadOnly(True)
        desc.setOpenExternalLinks(False)
        desc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        desc.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        desc.setHtml(self._grid._format_desc_html(info.short_desc))
        desc.setMaximumHeight(90)
        desc.setFrameStyle(QFrame.NoFrame)
        desc.setStyleSheet(
            f"QTextBrowser{{background: transparent; border: none; color: {'#bbb' if isDarkTheme() else '#444'}; font-size:12px;}}"
        )

        v.addLayout(title_row)
        v.addWidget(desc)
        v.addStretch(1)

        h = QHBoxLayout()
        self.launch_btn = PrimaryPushButton('Launch', self)
        self.launch_btn.setFixedHeight(28)
        self.launch_btn.setFixedWidth(90)
        self._launch_btn_default_style = self.launch_btn.styleSheet()
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
                try:
                    __import__(name.replace('-', '_'))
                except Exception:
                    return False
        return True

    def _update_requirements_state(self, installing: bool = False):
        ok = self._are_requirements_satisfied()
        if installing:
            self.install_btn.setEnabled(False)
            self.install_btn.setText('Installing...')
            self.launch_btn.setEnabled(False)
            return

        if ok:
            self.install_btn.setEnabled(False)
            self.install_btn.setText('Requirements OK')
            self.install_btn.setStyleSheet(
                "QPushButton:disabled { background-color: #21a366; color: white; border: none; padding: 6px 12px; border-radius:6px;}"
            )
            self.launch_btn.setEnabled(True)
        else:
            self.install_btn.setEnabled(True)
            self.install_btn.setText('Install Requirements')
            self.launch_btn.setEnabled(False)

    def _on_install(self):
        if not self.info.requirements or not os.path.isfile(self.info.requirements):
            return
        self._update_requirements_state(installing=True)

        self._proc = QProcess(self)
        self._proc.setProgram(sys.executable)
        self._proc.setArguments(['-m', 'pip', 'install', '-r', self.info.requirements])
        self._proc.finished.connect(lambda *_: self._on_install_finished())
        self._proc.start()

    def _on_install_finished(self):
        self._proc = None
        self._update_requirements_state(installing=False)
        try:
            window = self.window()
            packages_page = getattr(window, 'packagesPage', None)
            if packages_page is not None:
                try:
                    packages_page.refresh()
                except Exception:
                    pass
                try:
                    packages_page.packagesChanged.emit()
                except Exception:
                    pass
        except Exception:
            pass

    def set_invalid(self, invalid: bool):
        self._invalid = invalid
        if invalid:
            self.launch_btn.setStyleSheet(
                "QPushButton{background-color:#D13438; color:white; border:none; padding:6px 12px; border-radius:6px;}"
                "QPushButton:hover{background-color:#B91C1C;}"
            )
            self.launch_btn.setToolTip('Last launch failed – click to retry after fixing main.py')
        else:
            self.launch_btn.setStyleSheet(self._launch_btn_default_style)
            self.launch_btn.setToolTip('')

    def _sync_star_icon(self):
        try:
            starred = self._grid._is_starred(self.info)
        except Exception:
            starred = False
        try:
            self.star_btn.setIcon(self._grid._star_filled_icon if starred else self._grid._star_outline_icon)
            self.star_btn.setToolTip("Unstar" if starred else "Star")
        except Exception:
            pass

    def _on_star_clicked(self):
        try:
            self._grid.toggle_star(self.info)
        except Exception:
            self._sync_star_icon()

    def _on_launch(self):
        window = self.window()
        page = self._grid._build_page(self.info)
        if page is None:
            try:
                self.set_invalid(True)
            except Exception:
                pass
            QMessageBox.warning(self, 'Launch failed', 'Could not load macro widget (missing MainWidget/get_widget)')
            return
        try:
            self.set_invalid(False)
        except Exception:
            pass
        try:
            window.addMacroTab(self.info.name, self.info.display_name, QIcon(), page, replace_existing=True)
            try:
                self._grid.record_launch(self.info)
            except Exception:
                pass
        except Exception:
            widget = _load_pycro_widget(self.info)
            if widget is not None:
                widget.setWindowTitle(self.info.display_name)
                widget.resize(800, 600)
                widget.show()
                try:
                    self._grid.record_launch(self.info)
                except Exception:
                    pass
