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

class ExtractSdPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.stop_flag = False
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("Sub-Division Extractor", bold=True, color="#7eb8f7"))
        layout.addWidget(hline())

        g = QGroupBox("Settings")
        gv = QVBoxLayout(g)

        r0 = QHBoxLayout(); r0.addWidget(QLabel("Folder Path:"))
        self.folder_ed = QLineEdit(); self.folder_ed.setReadOnly(True); r0.addWidget(self.folder_ed, 1)
        bb = QPushButton("Browse"); bb.clicked.connect(lambda: (p := QFileDialog.getExistingDirectory(self,"Folder")) and self.folder_ed.setText(p)); r0.addWidget(bb)
        gv.addLayout(r0)

        r1 = QHBoxLayout(); r1.addWidget(QLabel("From Sub-Div Code:")); self.from_ed = QLineEdit(); r1.addWidget(self.from_ed)
        r1.addWidget(QLabel("To Sub-Div Code:")); self.to_ed = QLineEdit(); r1.addWidget(self.to_ed)
        gv.addLayout(r1)

        r2 = QHBoxLayout(); r2.addWidget(QLabel("Output Folder:"))
        self.out_ed = QLineEdit(r"D:\ExtractedFiles"); self.out_ed.setReadOnly(True); r2.addWidget(self.out_ed, 1)
        ch = QPushButton("Change"); ch.clicked.connect(lambda: (p := QFileDialog.getExistingDirectory(self,"Output")) and self.out_ed.setText(p)); r2.addWidget(ch)
        gv.addLayout(r2)
        layout.addWidget(g)

        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_stop  = QPushButton("■  STOP");  self.btn_stop.setStyleSheet(STOP_BTN); self.btn_stop.setEnabled(False)
        self.btn_start.clicked.connect(self.start); self.btn_stop.clicked.connect(self.stop)
        bh.addWidget(self.btn_start); bh.addWidget(self.btn_stop); bh.addStretch()
        layout.addLayout(bh)

        self.progress = QProgressBar(); layout.addWidget(self.progress)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(self.log_box)

    def log(self, msg): self.log_box.append(msg)

    def stop(self):
        self.stop_flag = True; self.log("⛔ Stopping…")

    def start(self):
        if not self.folder_ed.text() or not self.from_ed.text() or not self.to_ed.text():
            QMessageBox.warning(self, "Error", "Please fill in all fields."); return
        self.stop_flag = False
        self.btn_start.setEnabled(False); self.btn_stop.setEnabled(True)
        self.log("▶ Process started…")
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        folder = self.folder_ed.text()
        from_code = self.from_ed.text(); to_code = self.to_ed.text()
        out_dir = self.out_ed.text()

        from_regex = re.compile(rf"(S/DIV:\s*{from_code})|(S/Div:\s{{2}}{from_code})|(Sub\s+Division\s+{from_code})", re.IGNORECASE)
        to_regex   = re.compile(rf"(S/DIV:\s*{to_code})|(S/Div:\s{{2}}{to_code})|(Sub\s+Division\s+{to_code})", re.IGNORECASE)
        os.makedirs(out_dir, exist_ok=True)
        files = [f for f in os.listdir(folder) if f.lower().endswith(".txt")]
        total = len(files)

        for idx, filename in enumerate(files, 1):
            if self.stop_flag: break
            try:
                with open(os.path.join(folder, filename), "r", encoding="utf-8", errors="ignore") as f:
                    lines = f.readlines()
                from_index = next((max(i-2,0) for i,l in enumerate(lines) if from_regex.search(l)), None)
                if from_index is None: self.log(f"❌ FROM not found: {filename}"); continue
                to_index = next((min(i+40,len(lines)) for i in range(len(lines)-1,-1,-1) if to_regex.search(lines[i])), None)
                if to_index is None: self.log(f"❌ TO not found: {filename}"); continue
                final_lines = lines[from_index:to_index]
                base, ext = os.path.splitext(filename); out_name = filename; c2 = 1
                while os.path.exists(os.path.join(out_dir, out_name)):
                    out_name = f"{base} ({c2}){ext}"; c2 += 1
                with open(os.path.join(out_dir, out_name), "w", encoding="utf-8") as f:
                    f.writelines(final_lines)
                self.log(f"✅ Processed: {out_name}")
            except Exception as e:
                self.log(f"⚠ Error in {filename}: {e}")
            self.progress.setValue(int(idx / total * 100))

        self.btn_start.setEnabled(True); self.btn_stop.setEnabled(False)
        self.log("✔ Completed.")


# ──────────────────────────────────────────────────────────────
#  TOOL 4 — Payment Extractor (PaymentExtractSoft.py)
# ──────────────────────────────────────────────────────────────
