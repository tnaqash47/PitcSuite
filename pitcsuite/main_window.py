from PySide6.QtWidgets import QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QStackedWidget, QFrame
from PySide6.QtCore import Qt

from pitcsuite.icons import app_icon
from pitcsuite.ui_helpers import NAV_STYLE
from pitcsuite.tools.pitc_downloader import PITCPanel
from pitcsuite.tools.extract_page import ExtractPagePanel
from pitcsuite.tools.bill_downloader import BillSoftPanel
from pitcsuite.tools.sdiv_extractor import ExtractSdPanel
from pitcsuite.tools.payment_extractor import PaymentExtractPanel
from pitcsuite.tools.multi_merge import MultiMergePanel
from pitcsuite.tools.pdf_merge import PDFMergePanel
from pitcsuite.tools.file_checker import SrNoCheckerPanel
from pitcsuite.tools.double_cr import DoubleCRPanel
from pitcsuite.tools.settings import SettingsPanel


TOOLS = [
    ("📥  PITC Downloader",               PITCPanel),
    ("🔍  Extract Specific PITC Page",    ExtractPagePanel),
    ("🧾  E/Bills Downloader",            BillSoftPanel),
    ("✂️   SDiv Extractor from 88L",      ExtractSdPanel),
    ("💰  Payment Extractor from 88L",    PaymentExtractPanel),
    ("📚  Multi-Merge",                   MultiMergePanel),
    ("🔗  PDF Merge",                     PDFMergePanel),
    ("📂  File Existence Checker",        SrNoCheckerPanel),
    ("🔁  Double CR Checker",             DoubleCRPanel),
    ("⚙️   Settings",                     SettingsPanel),
]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PITC Suite  ·  Developed by Tahir Naqash")
        self.setWindowIcon(app_icon())
        self.resize(820, 620)

        root = QWidget()
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        nav = QWidget()
        nav.setFixedWidth(300)
        nav.setStyleSheet("background:#111318; border-right:1px solid #1e2230;")
        nav_layout = QVBoxLayout(nav)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        nav_layout.setSpacing(0)

        title_lbl = QLabel("PITC\nSuite")
        title_lbl.setAlignment(Qt.AlignCenter)
        title_lbl.setStyleSheet("color:#4a90d9; font-size:18px; font-weight:bold; padding:20px 10px 14px; letter-spacing:2px;")
        nav_layout.addWidget(title_lbl)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color:#1e2230;")
        nav_layout.addWidget(sep)

        self.nav_btns = []
        self.stack = QStackedWidget()

        for i, (name, panel_class) in enumerate(TOOLS):
            btn = QPushButton(name)
            btn.setStyleSheet(NAV_STYLE)
            btn.setCheckable(False)
            btn.clicked.connect(lambda _, idx=i: self.switch(idx))
            nav_layout.addWidget(btn)
            self.nav_btns.append(btn)
            self.stack.addWidget(panel_class())

        nav_layout.addStretch()
        ver = QLabel("v1.5")
        ver.setAlignment(Qt.AlignCenter)
        ver.setStyleSheet("color:#3a4a6a; font-size:10px; padding:8px;")
        nav_layout.addWidget(ver)

        outer.addWidget(nav)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(16, 16, 16, 12)
        content_layout.setSpacing(6)
        content_layout.addWidget(self.stack)
        outer.addWidget(content, 1)

        self.switch(0)

    def switch(self, idx):
        self.stack.setCurrentIndex(idx)
        for i, btn in enumerate(self.nav_btns):
            btn.setProperty("active", "true" if i == idx else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)
