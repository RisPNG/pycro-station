import os
import sys
import re
import html
import importlib.util
from importlib import metadata

from PySide6.QtCore import Qt, QFileSystemWatcher, QTimer, Signal, QProcess, QEvent
from PySide6.QtWidgets import *
from PySide6.QtGui import QIcon, QCursor
from qfluentwidgets import PrimaryPushButton, TransparentToolButton, isDarkTheme, FluentIcon as FIF


class PycroGrid(QScrollArea):
    """Scrollable grid container for listing Pycros."""

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setStyleSheet("QScrollArea{background: transparent; border: none;}")
        self.viewport().setStyleSheet("background: transparent;")

        self._content = QWidget(self)
        self._content.setStyleSheet("background: transparent;")
        self._grid = QGridLayout(self._content)
        self._grid.setContentsMargins(12, 12, 12, 12)
        self._grid.setSpacing(12)
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

        self._empty_label = QLabel("Pycros will appear here", self._content)
        self._empty_label.setAlignment(Qt.AlignCenter)
        self._empty_label.setStyleSheet("color: #aaa; font-size: 14px;")
        self._grid.addWidget(self._empty_label, 0, 0)
        self._grid.setRowStretch(0, 1)
        self._grid.setColumnStretch(0, 1)

        self._root = os.path.join(os.getcwd(), 'pycros')
        if os.path.isdir(self._root):
            self._watcher.addPath(self._root)

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

        QTimer.singleShot(0, self.refresh)

    def refresh(self):
        infos = self._scan_pycros()
        self._rebuild(infos)
        self._reload_open_tabs(infos)

    def _scan_pycros(self):
        root = self._root
        infos = []
        if not os.path.isdir(root):
            return infos
        try:
            subdirs = sorted([d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))])
        except Exception:
            subdirs = []

        try:
            dirs = self._watcher.directories()
            files = self._watcher.files()
            if dirs:
                self._watcher.removePaths(dirs)
            if files:
                self._watcher.removePaths(files)
        except Exception:
            pass
        self._watcher.addPath(root)

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
                has_python=has_python
            )

            infos.append(info)

            if os.path.isfile(desc_md):
                self._watcher.addPath(desc_md)
            if os.path.isfile(req_txt):
                self._watcher.addPath(req_txt)
            if os.path.isfile(main_py):
                self._watcher.addPath(main_py)
            if os.path.isfile(main_py):
                self._watcher.addPath(main_py)

        return infos

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

    def _rebuild(self, infos):
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

        for r in range(row + 1):
            self._grid.setRowStretch(r, 0)
        self._grid.setRowStretch(row + 1, 1)
        for c in range(columns):
            self._grid.setColumnStretch(c, 0)
        self._grid.setColumnStretch(columns, 1)

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
                    continue
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

        if info.long_desc:
            desc = QTextBrowser()
            desc.setOpenExternalLinks(True)
            desc.setHtml(self._format_desc_html(info.long_desc))
            desc.setStyleSheet('QTextBrowser{background:transparent; color:#ddd; border:none;}')
            desc.setFixedHeight(120)
            v.addWidget(desc)

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


class PycroInfo:
    def __init__(self, name, display_name, folder, main_py, requirements, description, short_desc, long_desc, info_lines, has_python: bool):
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

        safe_name = re.sub(r'[^A-Za-z0-9_]', '_', info.name)
        self.setObjectName(f"card__{safe_name}")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setAutoFillBackground(True)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.setMinimumSize(280, 210)
        self.setMaximumSize(320, 250)
        bg = "#242424" if isDarkTheme() else "#F2F2F2"
        border = "rgba(255,255,255,0.08)" if isDarkTheme() else "rgba(0,0,0,0.08)"
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

    def _on_launch(self):
        window = self.window()
        page = self._grid._build_page(self.info)
        if page is None:
            QMessageBox.warning(self, 'Launch failed', 'Could not load macro widget (missing MainWidget/get_widget)')
            return
        try:
            window.addMacroTab(self.info.name, self.info.display_name, QIcon(), page, replace_existing=True)
        except Exception:
            widget = _load_pycro_widget(self.info)
            if widget is not None:
                widget.setWindowTitle(self.info.display_name)
                widget.resize(800, 600)
                widget.show()

    def _load_widget(self) -> QWidget | None:
        return _load_pycro_widget(self.info)
