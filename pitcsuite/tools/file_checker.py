import os
import re
import time
import base64
import threading
import configparser

from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QLineEdit, QFileDialog, QCheckBox, QTextEdit, QProgressBar,
    QVBoxLayout, QHBoxLayout, QGroupBox, QMessageBox, QComboBox,
    QSpinBox
)
from PySide6.QtCore import Signal, QObject, QThread

from pitcsuite.ui_helpers import lbl, hline, START_BTN, STOP_BTN
from pitcsuite.templates import download_template
from pitcsuite.config import config_path as _config_path

class SrNoCheckerPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._stop_flag = False
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("File Existence Checker", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Checks if Sr. No. from Excel column A exists as a file in the selected folder.", color="#556"))
        layout.addWidget(hline())

        g = QGroupBox("Inputs")
        gv = QVBoxLayout(g)

        label_w = 70
        r1 = QHBoxLayout(); lbl_excel = QLabel("Excel File:"); lbl_excel.setFixedWidth(label_w); r1.addWidget(lbl_excel)
        self.excel_ed = QLineEdit(); self.excel_ed.setReadOnly(True); self.excel_ed.setProperty("preferred_width", 245); r1.addWidget(self.excel_ed, 1)
        b1 = QPushButton("Browse"); b1.clicked.connect(self._browse_excel); r1.addWidget(b1)
        gv.addLayout(r1)

        r2 = QHBoxLayout(); lbl_folder = QLabel("Folder:"); lbl_folder.setFixedWidth(label_w); r2.addWidget(lbl_folder)
        self.folder_ed = QLineEdit(); self.folder_ed.setReadOnly(True); self.folder_ed.setProperty("preferred_width", 245); r2.addWidget(self.folder_ed, 1)
        b2 = QPushButton("Browse"); b2.clicked.connect(self._browse_folder); r2.addWidget(b2)
        gv.addLayout(r2)
        layout.addWidget(g)

        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_stop  = QPushButton("■  STOP");  self.btn_stop.setStyleSheet(STOP_BTN)
        self.btn_start.clicked.connect(self._start); self.btn_stop.clicked.connect(self._stop)
        bh.addWidget(self.btn_start); bh.addWidget(self.btn_stop); bh.addStretch()
        layout.addLayout(bh)

        self.status_lbl = QLabel("Status: Idle"); self.status_lbl.setStyleSheet("color:#4a90d9;")
        layout.addWidget(self.status_lbl)

    def _browse_excel(self):
        f, _ = QFileDialog.getOpenFileName(self, "Excel File", "", "Excel Files (*.xlsx)")
        if f: self.excel_ed.setText(f)

    def _browse_folder(self):
        p = QFileDialog.getExistingDirectory(self, "Select Folder")
        if p: self.folder_ed.setText(p)

    def _start(self):
        self._stop_flag = False
        threading.Thread(target=self._process, daemon=True).start()

    def _stop(self):
        self._stop_flag = True
        self.status_lbl.setText("Status: Stopped")

    def _process(self):
        from openpyxl import load_workbook
        excel_path  = self.excel_ed.text()
        folder_path = self.folder_ed.text()

        if not excel_path or not folder_path:
            QMessageBox.critical(self, "Error", "Please select Excel file and Folder"); return

        try:
            wb = load_workbook(excel_path); ws = wb.active
            files_in_folder = os.listdir(folder_path)
            row = 2
            while True:
                if self._stop_flag: break
                sr_no = ws[f"A{row}"].value
                if sr_no is None: break
                found = any(os.path.splitext(f)[0] == str(sr_no) for f in files_in_folder)
                ws[f"B{row}"] = "Found" if found else "Not Found"
                self.status_lbl.setText(f"Checking Sr. No.: {sr_no}")
                row += 1
            wb.save(excel_path)
            self.status_lbl.setText("Status: Completed")
            QMessageBox.information(self, "Done", "Process Completed Successfully")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


# ──────────────────────────────────────────────────────────────
#  TOOL 11 — Double CR Checker (DoubleCRchecker.py)
# ──────────────────────────────────────────────────────────────
