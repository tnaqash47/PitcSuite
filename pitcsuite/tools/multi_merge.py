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

class MultiMergeWorkerSignals(QObject):
    update = Signal(str, int, int, bool)


class MultiMergePanel(QWidget):
    def __init__(self):
        super().__init__()
        self.stop_flag = {"stop": False}
        self.signals = MultiMergeWorkerSignals()
        self.signals.update.connect(self.on_update)

        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("Multi-Folder PDF Merge", bold=True, color="#7eb8f7"))
        layout.addWidget(hline())

        g = QGroupBox("Folders (up to 5)")
        gv = QVBoxLayout(g)
        self.folder_eds = []
        for i in range(5):
            r = QHBoxLayout(); r.addWidget(QLabel(f"Folder {i+1}:"))
            ed = QLineEdit(); ed.setReadOnly(True); r.addWidget(ed, 1)
            self.folder_eds.append(ed)
            idx = i
            b = QPushButton("Browse"); b.clicked.connect(lambda _, e=ed: (p := QFileDialog.getExistingDirectory(self,"Folder")) and e.setText(p)); r.addWidget(b)
            gv.addLayout(r)

        r_out = QHBoxLayout(); r_out.addWidget(QLabel("Output Folder:"))
        self.out_ed = QLineEdit(r"D:\MergedPDFs"); self.out_ed.setReadOnly(True); r_out.addWidget(self.out_ed, 1)
        c_out = QPushButton("Change"); c_out.clicked.connect(lambda: (p := QFileDialog.getExistingDirectory(self,"Output")) and self.out_ed.setText(p)); r_out.addWidget(c_out)
        gv.addLayout(r_out)
        layout.addWidget(g)

        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_start.clicked.connect(self.start_stop); bh.addWidget(self.btn_start); bh.addStretch()
        layout.addLayout(bh)

        ph = QHBoxLayout(); self.lbl_prog = QLabel("0 / 0"); ph.addWidget(self.lbl_prog); ph.addStretch()
        layout.addLayout(ph)
        self.progress = QProgressBar(); layout.addWidget(self.progress)
        self.log_lbl = QLabel("Ready."); self.log_lbl.setWordWrap(True); layout.addWidget(self.log_lbl)

    def natural_key(self, name):
        try: return (0, int(name))
        except ValueError: return (1, name.lower())

    def para_no(self, filename):
        name = os.path.splitext(filename)[0]
        name = re.sub(r"\s*\(\d+\)$", "", name)
        return name.strip()

    def start_stop(self):
        if self.btn_start.text().endswith("START"):
            folders = [e.text() for e in self.folder_eds if e.text()]
            if not folders: QMessageBox.warning(self, "Error", "Select at least one folder."); return
            self.stop_flag["stop"] = False
            self.btn_start.setText("■  STOP"); self.btn_start.setStyleSheet(STOP_BTN)
            threading.Thread(target=self.merge_worker, args=(folders,), daemon=True).start()
        else:
            self.stop_flag["stop"] = True

    def on_update(self, msg, cur, total, done):
        self.log_lbl.setText(msg)
        if total:
            self.progress.setMaximum(total); self.progress.setValue(cur)
            self.lbl_prog.setText(f"{cur} / {total}")
        if done:
            self.btn_start.setText("▶  START"); self.btn_start.setStyleSheet(START_BTN)

    def merge_worker(self, folders):
        from pypdf import PdfReader, PdfWriter
        out_dir = self.out_ed.text(); os.makedirs(out_dir, exist_ok=True)
        folder_data = []
        for folder in folders:
            lst = [(self.para_no(f), f) for f in os.listdir(folder) if f.lower().endswith(".pdf")]
            lst.sort(key=lambda x: self.natural_key(x[0])); folder_data.append(lst)
        if not any(folder_data):
            self.signals.update.emit("No PDFs found.", 0, 0, True); return

        writer = PdfWriter(); step = 0; total = sum(len(x) for x in folder_data)
        while any(folder_data):
            if self.stop_flag["stop"]: self.signals.update.emit("Stopped.", step, total, True); return
            # Choose the lowest remaining Para number across all folders.
            # Folder 1 may finish before the others; using it as the sole
            # source of `current` would otherwise leave this loop spinning.
            current_key, current = min(
                (
                    (self.natural_key(files[0][0]), files[0][0])
                    for files in folder_data if files
                ),
                key=lambda item: item[0],
            )
            for idx2, files in enumerate(folder_data, start=1):
                if not files: continue
                para, fname = files[0]
                if self.natural_key(para) != current_key: continue
                files.pop(0)
                step += 1; self.signals.update.emit(f"Folder {idx2} → Para {para}", step, total, False)
                try:
                    reader = PdfReader(os.path.join(folders[idx2-1], fname))
                    for page in reader.pages: writer.add_page(page)
                except Exception: pass

        out = os.path.join(out_dir, "Merged_Final_Order.pdf")
        with open(out, "wb") as f: writer.write(f)
        self.signals.update.emit("Merged_Final_Order.pdf created ✔", step, step, True)


# ──────────────────────────────────────────────────────────────
#  TOOL 6 — PDF Merge (pdfmerge.py)
# ──────────────────────────────────────────────────────────────
