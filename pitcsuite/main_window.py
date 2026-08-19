from PySide6.QtWidgets import QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QStackedWidget, QFrame, QLineEdit, QSpinBox, QSizePolicy, QScrollArea, QBoxLayout
from PySide6.QtCore import Qt

from pitcsuite.icons import app_icon
from pitcsuite.ui_helpers import NAV_STYLE
from pitcsuite.tools.pitc_downloader import PITCPanel
from pitcsuite.tools.extract_page import ExtractPagePanel
from pitcsuite.tools.bill_downloader import BillSoftPanel
from pitcsuite.tools.payment_extract_88l import PaymentExtract88LPanel
from pitcsuite.tools.multi_merge import MultiMergePanel
from pitcsuite.tools.pdf_merge import PDFMergePanel
from pitcsuite.tools.file_checker import SrNoCheckerPanel
from pitcsuite.tools.file_renamer import FileRenamerPanel
from pitcsuite.tools.double_cr import DoubleCRPanel
from pitcsuite.tools.settings import SettingsPanel
from pitcsuite.tools.integrated_tools import CP22TCheckerPanel, BillScraperPanel, FileMoverPanel, CP52PostingCheckerPanel


TOOLS = [
    ("📥  PITC Downloader",               PITCPanel),
    ("🔍  Extract PITC Pages",            ExtractPagePanel),
    ("🧾  E/Bills Downloader",            BillSoftPanel),
    ("💰  Payment Extract (88L)",         PaymentExtract88LPanel),
    ("📚  Multi-Merge",                   MultiMergePanel),
    ("🔗  PDF Merge",                     PDFMergePanel),
    ("📂  File Checker",                  SrNoCheckerPanel),
    ("🔁  Double CR Checker",             DoubleCRPanel),
    ("🧾  CP22T Posting Checker",          CP22TCheckerPanel),
    ("🧮  Bill Scrapper",                  BillScraperPanel),
    ("📦  File Mover",                     FileMoverPanel),
    ("🏷️  File Renamer",                    FileRenamerPanel),
    ("📋  CP-52 Posting Checker",           CP52PostingCheckerPanel),
    ("⚙️   Settings",                     SettingsPanel),
]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PITC Suite  ·  Developed by Tahir Naqash")
        self.setWindowIcon(app_icon())
        self.resize(720, 680)
        self.setMinimumSize(600, 480)
        self.nav_collapsed = True

        root = QWidget()
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.nav = QWidget()
        self.nav.setFixedWidth(245)
        self.nav.setStyleSheet("background:#111318; border-right:1px solid #1e2230;")
        nav_layout = QVBoxLayout(self.nav)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        nav_layout.setSpacing(0)

        nav_head = QHBoxLayout()
        nav_head.setContentsMargins(6, 10, 6, 8)
        self.nav_menu_btn = QPushButton("☰")
        self.nav_menu_btn.setFixedSize(34, 28)
        self.nav_menu_btn.setToolTip("Show/hide tool labels")
        self.nav_menu_btn.clicked.connect(self.toggle_nav)
        nav_head.addWidget(self.nav_menu_btn)
        self.title_lbl = QLabel("PITC\nSuite")
        self.title_lbl.setAlignment(Qt.AlignCenter)
        self.title_lbl.setStyleSheet("color:#4a90d9; font-size:18px; font-weight:bold; letter-spacing:2px;")
        nav_head.addWidget(self.title_lbl, 1)
        nav_layout.addLayout(nav_head)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color:#1e2230;")
        nav_layout.addWidget(sep)

        self.nav_btns = []
        self.nav_names = []
        self.nav_icons = []
        self.stack = QStackedWidget()
        self.stack.setMinimumSize(0, 0)
        # Keep the tool form readable on wide/maximized windows instead of
        # allowing every group box and row to stretch across the screen.
        self.stack.setMaximumWidth(720)
        self.stack.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)

        for i, (name, panel_class) in enumerate(TOOLS):
            btn = QPushButton(name)
            btn.setStyleSheet(NAV_STYLE)
            btn.setCheckable(False)
            btn.clicked.connect(lambda _, idx=i: self.switch(idx))
            nav_layout.addWidget(btn)
            self.nav_btns.append(btn)
            self.nav_names.append(name)
            self.nav_icons.append(self._nav_icon(name))
            panel = panel_class()
            panel.setMinimumSize(0, 0)
            panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            self.stack.addWidget(panel)

        nav_layout.addStretch()
        self.ver_lbl = QLabel("v1.8")
        self.ver_lbl.setAlignment(Qt.AlignCenter)
        self.ver_lbl.setStyleSheet("color:#3a4a6a; font-size:10px; padding:8px;")
        nav_layout.addWidget(self.ver_lbl)
        self._apply_nav_state()

        outer.addWidget(self.nav)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(16, 12, 16, 12)
        content_layout.setSpacing(6)
        self.content_scroll = QScrollArea()
        self.content_scroll.setWidgetResizable(True)
        self.content_scroll.setFrameShape(QFrame.NoFrame)
        # Keep wheel/trackpad scrolling, but hide the scrollbar controls from
        # the compact layout so they do not consume space or add visual noise.
        self.content_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.content_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.content_scroll.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.content_scroll.setWidget(self.stack)
        content_layout.addWidget(self.content_scroll)
        outer.addWidget(content, 1)

        self._make_inputs_responsive()
        self._update_responsive_sizes()

        self.switch(0)

    def toggle_nav(self):
        self.nav_collapsed = not self.nav_collapsed
        self._apply_nav_state()
        self._update_responsive_sizes()

    def _nav_icon(self, text):
        parts = text.split()
        return parts[0] if parts else text[:1]

    def _apply_nav_state(self):
        self.nav.setFixedWidth(44 if self.nav_collapsed else 245)
        self.title_lbl.setVisible(not self.nav_collapsed)
        self.ver_lbl.setVisible(not self.nav_collapsed)
        for btn, full, icon in zip(self.nav_btns, self.nav_names, self.nav_icons):
            btn.setText(icon if self.nav_collapsed else full)
            btn.setStyleSheet(
                NAV_STYLE + "\nQPushButton { padding: 10px 0; text-align: center; }"
                if self.nav_collapsed else NAV_STYLE
            )

    def _make_inputs_responsive(self):
        for edit in self.stack.findChildren(QLineEdit):
            if isinstance(edit.parentWidget(), QSpinBox):
                continue
            edit.setMinimumWidth(70)
            edit.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            edit.setCursorPosition(0)

        for btn in self.stack.findChildren(QPushButton):
            text = btn.text().lower()
            if "download template" in text:
                btn.setText("Template")
                text = btn.text().lower()
            btn.setMinimumWidth(36)
            preferred = btn.property("preferred_width")
            if preferred:
                btn.setMaximumWidth(int(preferred))
            elif any(word in text for word in ("browse", "change", "open", "edit")):
                btn.setMaximumWidth(64)
            elif "template" in text:
                btn.setMaximumWidth(82)
            elif any(word in text for word in ("start", "stop", "pause", "resume", "save")):
                btn.setMaximumWidth(92)
            btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        for i in range(self.stack.count()):
            self._compact_line_edit_rows(self.stack.widget(i).layout())

    def _compact_line_edit_rows(self, layout):
        if layout is None:
            return
        has_line_edit = any(
            layout.itemAt(i).widget() is not None
            and isinstance(layout.itemAt(i).widget(), QLineEdit)
            and not isinstance(layout.itemAt(i).widget().parentWidget(), QSpinBox)
            for i in range(layout.count())
        )
        for i in range(layout.count()):
            item = layout.itemAt(i)
            widget = item.widget()
            child = item.layout()
            if isinstance(layout, QBoxLayout) and has_line_edit:
                # Inputs use compact fixed widths; keep them immediately
                # beside their labels and leave spare space at row end.
                layout.setStretch(i, 0)
                layout.setAlignment(Qt.AlignLeft)
            if widget is not None and widget.layout() is not None:
                self._compact_line_edit_rows(widget.layout())
            if child is not None:
                self._compact_line_edit_rows(child)
        if isinstance(layout, QBoxLayout) and has_line_edit:
            layout.addStretch(1)

    def _update_responsive_sizes(self):
        nav_width = 44 if self.nav_collapsed else 245
        content_width = max(self.width() - nav_width - 48, 240)
        font_size = 10 if self.isMaximized() else 8
        for widget in self.stack.findChildren(QWidget):
            font = widget.font()
            font.setPointSize(font_size)
            widget.setFont(font)

        for edit in self.stack.findChildren(QLineEdit):
            if isinstance(edit.parentWidget(), QSpinBox):
                continue
            preferred = edit.property("preferred_width")
            # Keep fields close to their labels instead of letting them fill
            # the entire row on wide/maximized windows.  They still shrink
            # proportionally when the available content area is narrow.
            base_width = min(int(preferred) if preferred else 220, 220)
            target_width = min(base_width, max(70, int(content_width * 0.34)))
            edit.setFixedWidth(target_width)
        for btn in self.stack.findChildren(QPushButton):
            preferred = btn.property("preferred_width")
            if preferred:
                btn.setMaximumWidth(min(int(preferred), max(80, int(content_width * 0.5))))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_responsive_sizes()

    def switch(self, idx):
        self.stack.setCurrentIndex(idx)
        for i, btn in enumerate(self.nav_btns):
            btn.setProperty("active", "true" if i == idx else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)
