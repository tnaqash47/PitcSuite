from PySide6.QtWidgets import QLabel, QFrame

DARK_STYLE = """
QMainWindow, QWidget {
    background-color: #1a1d23;
    color: #e0e6f0;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 11px;
}
QGroupBox {
    border: 1px solid #2e3340;
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 8px;
    font-weight: bold;
    color: #7eb8f7;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
}
QLineEdit {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    padding: 3px 7px;
    color: #e0e6f0;
    font-size: 11px;
}
QLineEdit:focus { border-color: #4a90d9; }
QPushButton {
    background: #252933;
    border: 1px solid #3a3f50;
    border-radius: 4px;
    padding: 5px 10px;
    color: #c8d6f0;
}
QPushButton:hover { background: #2e3340; border-color: #4a90d9; color: #fff; }
QPushButton:pressed { background: #1e2430; }
QPushButton:disabled { color: #555; border-color: #2a2a2a; }
QTextEdit {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    color: #a8c0e0;
    font-family: 'Consolas', monospace;
    font-size: 11px;
}
QProgressBar {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    height: 14px;
    text-align: center;
    color: #e0e6f0;
}
QProgressBar::chunk { background: #4a90d9; border-radius: 3px; }
QComboBox {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    padding: 4px 8px;
    color: #e0e6f0;
}
QComboBox QAbstractItemView {
    background: #1a1d23;
    border: 1px solid #4a90d9;
    selection-background-color: #2e3340;
}
QSpinBox {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    padding: 3px 6px;
    color: #e0e6f0;
}
QCheckBox { color: #c8d6f0; }
QCheckBox::indicator { width: 14px; height: 14px; border: 1px solid #3a3f50; background: #12151c; border-radius: 3px; }
QCheckBox::indicator:checked { background: #4a90d9; }
QLabel { color: #c8d6f0; }
QScrollBar:vertical { background: #12151c; width: 8px; border-radius: 4px; }
QScrollBar::handle:vertical { background: #2e3340; border-radius: 4px; }
"""

NAV_STYLE = """
QPushButton {
    background: transparent;
    border: none;
    border-left: 3px solid transparent;
    border-radius: 0;
    padding: 10px 12px;
    text-align: left;
    color: #8090a8;
    font-size: 11px;
    font-family: 'Consolas', monospace;
}
QPushButton:hover { background: #1e2230; color: #c8d6f0; border-left-color: #3a4a6a; }
QPushButton[active="true"] {
    background: #1e2844;
    color: #7eb8f7;
    border-left-color: #4a90d9;
    font-weight: bold;
}
"""

START_BTN = "background:#2a6; border:none; border-radius:4px; padding:6px 18px; color:#fff; font-weight:bold;"
STOP_BTN  = "background:#922; border:none; border-radius:4px; padding:6px 18px; color:#fff; font-weight:bold;"


def lbl(text, bold=False, color=None):
    w = QLabel(text)
    w.setWordWrap(True)
    if bold:
        f = w.font(); f.setBold(True); w.setFont(f)
    if color:
        w.setStyleSheet(f"color:{color};")
    return w


def hline():
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setStyleSheet("color:#2e3340;")
    return f


# ──────────────────────────────────────────────────────────────
#  TOOL 1 — PITC Downloader (CDS10)
#  All report types · OTP login · Month+Amount filter · Highlight · Para Sr. Stamp
# ──────────────────────────────────────────────────────────────
