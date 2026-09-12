from pathlib import Path

from PySide6.QtGui import QIcon


def app_icon():
    return QIcon(str(Path(__file__).with_name("assets") / "pitcsuite_icon.ico"))


def pdf_merger_icon():
    return QIcon(str(Path(__file__).with_name("assets") / "pdfmerger.ico"))
