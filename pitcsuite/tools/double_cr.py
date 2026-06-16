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

DOUBLE_CR_EXPORT_DIR = r"D:\ExportedXls"
DOUBLE_CR_MAX_OCC    = 5   # 1st to 5th occurrence


class DoubleCRPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._cr_file  = ""
        self._ac_file  = ""
        self._stop_req = False

        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("Double CR Checker", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Matches AC Ref. numbers against CR file and lists up to 5 occurrences.", color="#556"))
        layout.addWidget(hline())

        g = QGroupBox("Files")
        gv = QVBoxLayout(g)

        r1 = QHBoxLayout(); r1.addWidget(QLabel("CR File:"))
        self.cr_lbl = QLabel("(not selected)"); self.cr_lbl.setStyleSheet("color:#556;"); r1.addWidget(self.cr_lbl, 1)
        b1 = QPushButton("Browse"); b1.clicked.connect(self._browse_cr); r1.addWidget(b1)
        gv.addLayout(r1)

        r2 = QHBoxLayout(); r2.addWidget(QLabel("AC Nos File:"))
        self.ac_lbl = QLabel("(not selected)"); self.ac_lbl.setStyleSheet("color:#556;"); r2.addWidget(self.ac_lbl, 1)
        b2 = QPushButton("Browse"); b2.clicked.connect(self._browse_ac); r2.addWidget(b2)
        gv.addLayout(r2)

        tpl_h = QHBoxLayout()
        btn_cr_tpl = QPushButton("⬇  CR Template");  btn_cr_tpl.clicked.connect(self._save_cr_template)
        btn_ac_tpl = QPushButton("⬇  AC Template");  btn_ac_tpl.clicked.connect(self._save_ac_template)
        btn_cr_tpl.setProperty("preferred_width", 128)
        btn_ac_tpl.setProperty("preferred_width", 128)
        tpl_h.addWidget(btn_cr_tpl); tpl_h.addWidget(btn_ac_tpl); tpl_h.addStretch()
        gv.addLayout(tpl_h)
        layout.addWidget(g)

        self.progress = QProgressBar(); layout.addWidget(self.progress)

        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_stop  = QPushButton("■  STOP");  self.btn_stop.setStyleSheet(STOP_BTN)
        self.btn_start.clicked.connect(self._start); self.btn_stop.clicked.connect(self._stop)
        bh.addWidget(self.btn_start); bh.addWidget(self.btn_stop); bh.addStretch()
        layout.addLayout(bh)

        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(self.log_box)

    def _log(self, msg):
        self.log_box.append(msg)
        QApplication.processEvents()

    def _browse_cr(self):
        f, _ = QFileDialog.getOpenFileName(self, "Select CR File", "", "Excel Files (*.xlsx)")
        if f:
            self._cr_file = f
            self.cr_lbl.setText(os.path.basename(f)); self.cr_lbl.setStyleSheet("color:#4a90d9;")
            self._log("✔ CR File selected")

    def _browse_ac(self):
        f, _ = QFileDialog.getOpenFileName(self, "Select AC Nos File", "", "Excel Files (*.xlsx)")
        if f:
            self._ac_file = f
            self.ac_lbl.setText(os.path.basename(f)); self.ac_lbl.setStyleSheet("color:#4a90d9;")
            self._log("✔ AC File selected")

    def _save_cr_template(self):
        import pandas as pd
        f, _ = QFileDialog.getSaveFileName(self, "Save CR Template", "CR_Template.xlsx", "Excel Files (*.xlsx)")
        if f:
            pd.DataFrame(columns=["Sr#", "Ref.#", "SST", "Date", "Amount"]).to_excel(f, index=False)
            self._log("✔ CR template saved")

    def _save_ac_template(self):
        import pandas as pd
        f, _ = QFileDialog.getSaveFileName(self, "Save AC Template", "AC_Template.xlsx", "Excel Files (*.xlsx)")
        if f:
            pd.DataFrame(columns=["Sr#", "Ref.#"]).to_excel(f, index=False)
            self._log("✔ AC template saved")

    def _stop(self):
        self._stop_req = True
        self._log("🛑 Stop requested")

    def _start(self):
        if not self._cr_file or not self._ac_file:
            self._log("❌ Please select both files"); return
        self._stop_req = False
        self.progress.setValue(0)
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        import pandas as pd
        os.makedirs(DOUBLE_CR_EXPORT_DIR, exist_ok=True)

        self._log("📖 Reading CR file...")
        cr_df = pd.read_excel(self._cr_file, header=0, dtype={1: str})
        self._log("📖 Reading AC file...")
        ac_df = pd.read_excel(self._ac_file, header=0, dtype={1: str})

        cr_ref_col = cr_df.columns[1]
        cr_sst_col = cr_df.columns[2]
        cr_dt_col  = cr_df.columns[3]
        cr_amt_col = cr_df.columns[4]
        ac_sr_col  = ac_df.columns[0]
        ac_ref_col = ac_df.columns[1]

        grouped = cr_df.groupby(cr_ref_col, dropna=False)
        total   = len(ac_df)
        output  = []

        self._log("⚙ Processing records...")

        for i, row in ac_df.iterrows():
            if self._stop_req:
                self._log("⛔ Process stopped"); return

            ref = str(row[ac_ref_col]).strip()
            matches = (grouped.get_group(ref) if ref in grouped.groups else pd.DataFrame())

            row_data = [row[ac_sr_col], ref]
            for occ in range(DOUBLE_CR_MAX_OCC):
                if occ < len(matches):
                    row_data.extend([
                        matches.iloc[occ][cr_sst_col],
                        matches.iloc[occ][cr_dt_col],
                        matches.iloc[occ][cr_amt_col]
                    ])
                else:
                    row_data.extend(["", "", ""])

            output.append(row_data)
            self.progress.setValue(int((len(output) / total) * 100))

        cols = ["Sr#", "Ref.#"]
        for i in range(1, DOUBLE_CR_MAX_OCC + 1):
            suffix = "st" if i == 1 else "nd" if i == 2 else "rd" if i == 3 else "th"
            cols += [f"{i}{suffix} SST", f"{i}{suffix} Date", f"{i}{suffix} Amount"]

        out_df   = pd.DataFrame(output, columns=cols)
        out_path = os.path.join(DOUBLE_CR_EXPORT_DIR, "Output.xlsx")
        out_df.to_excel(out_path, index=False, engine="openpyxl")

        self.progress.setValue(100)
        self._log("✅ Completed successfully")
        self._log(f"📁 Output saved at: {out_path}")


# ──────────────────────────────────────────────────────────────
#  TEMPLATE HELPER — generate .xlsx template and offer Save dialog
# ──────────────────────────────────────────────────────────────
