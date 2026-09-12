from PySide6.QtWidgets import QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QStackedWidget, QFrame, QLineEdit, QSpinBox, QSizePolicy, QScrollArea, QAbstractScrollArea, QBoxLayout
from PySide6.QtCore import Qt, QSize

from pitcsuite.icons import app_icon, pdf_merger_icon
from pitcsuite.ui_helpers import NAV_STYLE
from pitcsuite.tools.pitc_downloader import PITCPanel
from pitcsuite.tools.extract_page import ExtractPagePanel
from pitcsuite.tools.bill_downloader import BillSoftPanel
from pitcsuite.tools.payment_extract_88l import PaymentExtract88LPanel
from pitcsuite.tools.pdf_merger import PDFMerger
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
    ("PDF Merger",                    PDFMerger),
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
        self.setWindowTitle("PITC Suite")
        self.setWindowIcon(app_icon())
        self.resize(780, 680)
        self.setMinimumSize(540, 500)
        root = QWidget()
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        self.nav_collapsed = True
        self.nav = QWidget()
        nav_layout = QVBoxLayout(self.nav)
        nav_layout.setContentsMargins(0, 6, 0, 6)
        nav_layout.setSpacing(0)
        self.nav_menu_btn = QPushButton("☰")
        self.nav_menu_btn.setToolTip("Show/hide tool labels")
        self.nav_menu_btn.clicked.connect(self.toggle_nav)
        nav_layout.addWidget(self.nav_menu_btn)
        self.title_lbl = QLabel("PITC Suite")
        self.title_lbl.setAlignment(Qt.AlignCenter)
        nav_layout.addWidget(self.title_lbl)
        self.nav_btns = []
        self.nav_names = []
        for i, (name, panel_class) in enumerate(TOOLS):
            button = QPushButton(name)
            if panel_class is PDFMerger:
                button.setIcon(pdf_merger_icon())
                button.setIconSize(QSize(20, 20))
            button.setToolTip(name)
            button.clicked.connect(lambda checked=False, idx=i: self.switch(idx))
            nav_layout.addWidget(button)
            self.nav_btns.append(button)
            self.nav_names.append(name)
        nav_layout.addStretch()
        self.ver_lbl = QLabel("v1.8")
        self.ver_lbl.setAlignment(Qt.AlignCenter)
        nav_layout.addWidget(self.ver_lbl)
        self._apply_nav_state()
        outer.addWidget(self.nav)
        content = QWidget()
        outer.addWidget(content, 1)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)
        heading = QLabel("PITC Suite")
        heading.setStyleSheet("font-size:13pt; font-weight:600;")
        layout.addWidget(heading)
        self.stack = QStackedWidget()
        self.panels = []
        for name, panel_class in TOOLS:
            panel = panel_class()
            self.panels.append(panel)
            if panel.layout():
                panel.layout().setContentsMargins(0, 6, 0, 0)
                panel.layout().setAlignment(Qt.AlignTop)
            for button in panel.findChildren(QPushButton):
                if button.property("role") is None and not button.styleSheet():
                    button.setProperty("role", "secondary")
                button.setMinimumWidth(0)
                button.setMaximumWidth(16777215)
            for edit in panel.findChildren(QLineEdit):
                if isinstance(edit.parentWidget(), QSpinBox):
                    continue
                edit.setMinimumWidth(70)
                edit.setMaximumWidth(16777215)
                edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setAlignment(Qt.AlignLeft | Qt.AlignTop)
            scroll.setWidget(panel)
            # Hide the controls without disabling wheel, touchpad or keyboard scrolling.
            for area in [scroll, *panel.findChildren(QAbstractScrollArea)]:
                area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.stack.addWidget(scroll)
        layout.addWidget(self.stack, 1)
        credit = QLabel("Developed by Tahir Naqash")
        credit.setAlignment(Qt.AlignRight)
        credit.setStyleSheet("color:#b8c0cc; font-size:7pt;")
        layout.addWidget(credit)
        self.switch(0)

    def switch(self, idx):
        self.stack.setCurrentIndex(idx)
        for i, button in enumerate(self.nav_btns):
            button.setProperty("active", "true" if i == idx else "false")
            button.style().unpolish(button)
            button.style().polish(button)

    def toggle_nav(self):
        self.nav_collapsed = not self.nav_collapsed
        self._apply_nav_state()

    def _apply_nav_state(self):
        self.nav.setFixedWidth(44 if self.nav_collapsed else 200)
        self.title_lbl.setVisible(not self.nav_collapsed)
        self.ver_lbl.setVisible(not self.nav_collapsed)
        for button, name in zip(self.nav_btns, self.nav_names):
            if not button.icon().isNull():
                button.setText("" if self.nav_collapsed else name)
            else:
                button.setText(name.split()[0] if self.nav_collapsed else name)
            button.setStyleSheet(NAV_STYLE + (
                "\nQPushButton { padding:10px 0; text-align:center; }"
                if self.nav_collapsed else ""))

    def closeEvent(self, event):
        # Embedded widgets do not receive the main window's close event.
        for panel in self.panels:
            if isinstance(panel, PDFMerger) and panel.worker and panel.worker.isRunning():
                panel.cancel()
                panel.log.appendPlainText("Stopping merge. Close PITC Suite after it finishes.")
                event.ignore()
                return
        super().closeEvent(event)
