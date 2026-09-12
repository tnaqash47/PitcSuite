from PySide6.QtGui import QPalette, QColor
from PySide6.QtWidgets import QLabel, QFrame, QMessageBox
from PySide6.QtCore import Qt

def configure_app(app):
    app.setStyle("Fusion")
    font = app.font()
    font.setPointSizeF(font.pointSizeF())
    app.setFont(font)
    palette = QPalette()
    colors = {
        "Window": "#20242c", "WindowText": "#edf0f5",
        "Base": "#161a21", "AlternateBase": "#292f3a",
        "ToolTipBase": "#edf0f5", "ToolTipText": "#161a21",
        "Text": "#edf0f5", "Button": "#303846",
        "ButtonText": "#edf0f5", "BrightText": "#ffffff",
        "Highlight": "#347bd1", "HighlightedText": "#ffffff",
        "PlaceholderText": "#9ba8ba", "Link": "#79b7ff",
    }
    for role, color in colors.items():
        palette.setColor(getattr(QPalette.ColorRole, role), QColor(color))
    for role in (QPalette.ColorRole.Text, QPalette.ColorRole.WindowText, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor("#818b9b"))
    app.setPalette(palette)

DARK_STYLE = """
        QPushButton[role=primary] { background: #3c9b5f; color: white; border: 1px solid #55b876; border-radius: 5px; padding: 6px 14px; font-weight: 600; }
        QPushButton[role=primary]:hover { background: #49ad6d; }
        QPushButton[role=secondary] { background: #b8753c; color: white; border: 1px solid #d39152; border-radius: 5px; padding: 5px 11px; }
        QPushButton[role=secondary]:hover { background: #ca874a; }
        QPushButton:disabled { color: #818b9b; background: #363d49; border-color: #4a5361; }
    """

START_BTN = "QPushButton { background:#3c9b5f; color:white; border:1px solid #55b876; border-radius:5px; padding:6px 14px; font-weight:600; } QPushButton:hover { background:#49ad6d; } QPushButton:disabled { color:#818b9b; background:#363d49; border-color:#4a5361; }"
STOP_BTN = "QPushButton { background:#b8753c; color:white; border:1px solid #d39152; border-radius:5px; padding:5px 11px; } QPushButton:hover { background:#ca874a; } QPushButton:disabled { color:#818b9b; background:#363d49; border-color:#4a5361; }"


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


def notify(parent, title, message, critical=False):
    """Show a result without blocking the Qt event loop with exec()."""
    box = QMessageBox(QMessageBox.Critical if critical else QMessageBox.Information,
                      title, message, parent=parent)
    box.setAttribute(Qt.WA_DeleteOnClose)
    active = getattr(parent, "_active_notifications", None)
    if active is None:
        active = parent._active_notifications = []
    active.append(box)
    box.finished.connect(lambda: active.remove(box) if box in active else None)
    box.open()
    return box


# ──────────────────────────────────────────────────────────────
#  TOOL 1 — PITC Downloader (CDS10)
#  All report types · OTP login · Month+Amount filter · Highlight · Para Sr. Stamp
# ──────────────────────────────────────────────────────────────

NAV_STYLE = """
QPushButton {
    background: transparent; color: #b8c0cc;
    border: none; border-left: 3px solid transparent;
    border-radius: 0; padding: 10px 12px; text-align: left;
}
QPushButton:hover { background: #292f3a; color: #edf0f5; }
QPushButton[active="true"] {
    background: #303846; color: #79b7ff;
    border-left-color: #347bd1; font-weight: 600;
}
"""
