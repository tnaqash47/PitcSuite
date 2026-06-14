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

class PDFMergeWorkerSignals(QObject):
    update = Signal(str, int, int, bool)


class PDFMergePanel(QWidget):
    def __init__(self):
        super().__init__()
        self.stop_flag = {"stop": False}
        self.signals = PDFMergeWorkerSignals()
        self.signals.update.connect(self.on_update)

        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("PDF Merge", bold=True, color="#7eb8f7"))
        layout.addWidget(hline())

        g = QGroupBox("Folder")
        gv = QVBoxLayout(g)
        r = QHBoxLayout(); r.addWidget(QLabel("PDF Folder:"))
        self.folder_ed = QLineEdit(); self.folder_ed.setReadOnly(True); r.addWidget(self.folder_ed, 1)
        b = QPushButton("Browse"); b.clicked.connect(lambda: (p := QFileDialog.getExistingDirectory(self,"Folder")) and self.folder_ed.setText(p)); r.addWidget(b)
        gv.addLayout(r)
        layout.addWidget(g)

        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_start.clicked.connect(self.start_stop); bh.addWidget(self.btn_start); bh.addStretch()
        layout.addLayout(bh)

        ph = QHBoxLayout(); self.lbl_prog = QLabel("0 / 0"); ph.addWidget(self.lbl_prog); ph.addStretch()
        layout.addLayout(ph)
        self.progress = QProgressBar(); layout.addWidget(self.progress)
        self.log_lbl = QLabel("Ready."); self.log_lbl.setWordWrap(True); layout.addWidget(self.log_lbl)

    def natural_key(self, filename):
        name = os.path.splitext(filename)[0]
        m = re.match(r"^(\d+)(?:\s*\((\d+)\))?$", name)
        if m: return (0, int(m.group(1)), int(m.group(2)) if m.group(2) else 0)
        return (1, name.lower(), 0)

    def start_stop(self):
        if self.btn_start.text().endswith("START"):
            if not self.folder_ed.text(): QMessageBox.warning(self, "Error", "Select a folder first."); return
            self.stop_flag["stop"] = False
            self.btn_start.setText("■  STOP"); self.btn_start.setStyleSheet(STOP_BTN)
            threading.Thread(target=self.merge_worker, daemon=True).start()
        else:
            self.stop_flag["stop"] = True

    def on_update(self, msg, cur, total, done):
        self.log_lbl.setText(msg)
        if total: self.progress.setMaximum(total); self.progress.setValue(cur); self.lbl_prog.setText(f"{cur} / {total}")
        if done: self.btn_start.setText("▶  START"); self.btn_start.setStyleSheet(START_BTN)

    def merge_worker(self):
        from pypdf import PdfReader, PdfWriter
        folder = self.folder_ed.text()
        pdfs = sorted([f for f in os.listdir(folder) if f.lower().endswith(".pdf")], key=self.natural_key)
        if not pdfs: self.signals.update.emit("No PDFs found.", 0, 0, True); return
        writer = PdfWriter(); total = len(pdfs)
        for i, pdf in enumerate(pdfs, 1):
            if self.stop_flag["stop"]: self.signals.update.emit("Stopped.", i, total, True); return
            self.signals.update.emit(f"Merging: {pdf}", i, total, False)
            try:
                reader = PdfReader(os.path.join(folder, pdf))
                for page in reader.pages: writer.add_page(page)
            except Exception: pass
        out_name = f"Merged_{os.path.basename(folder)}.pdf"
        out_path = os.path.join(folder, out_name)
        with open(out_path, "wb") as f: writer.write(f)
        self.signals.update.emit(f"Saved: {out_name} ✔", total, total, True)


# ──────────────────────────────────────────────────────────────
#  TOOL 10 — File Existence Checker (SrNoCheckinfolder.py)
# ──────────────────────────────────────────────────────────────
