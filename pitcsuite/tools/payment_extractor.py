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

class PaymentWorker(QThread):
    log      = Signal(str)
    progress = Signal(int)
    finished = Signal()

    def __init__(self, excel_path, text_folder):
        super().__init__()
        self.excel_path  = excel_path
        self.text_folder = text_folder
        self.stop_flag   = False

    def stop(self): self.stop_flag = True

    def run(self):
        try:
            import pandas as pd
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.pdfgen import canvas as rl_canvas
            from reportlab.lib import colors
            from PyPDF2 import PdfMerger
        except ImportError as e:
            self.log.emit(f"ERROR: Missing library — {e}"); self.finished.emit(); return

        OUTPUT_FOLDER = r"D:\Payment_PDFs"; os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        PAGE_SIZE = landscape(A4); LEFT_MARGIN = 18; TOP_MARGIN = 30; BOTTOM_MARGIN = 22; LINE_GAP = 9

        def get_ledger_header(text):
            for line in text.splitlines():
                if line.lstrip().startswith("BATCH:"): return line
            return ""

        def extract_matching_pages(text, ref_raw, ledger_header):
            lines = text.splitlines(); pages = []
            ref_pattern = re.compile(rf"(?<!\d){re.escape(ref_raw)}(?!\d)")
            start_idx = None
            for i, line in enumerate(lines):
                if re.search(r"(S/DIV:|S/Div:\s{2}|Sub Division\s{2,})", line, re.IGNORECASE):
                    if start_idx is not None:
                        body = lines[start_idx:i-1]
                        page = "\n".join(([ledger_header] if ledger_header else []) + body)
                        if ref_pattern.search(page): pages.append(page)
                    start_idx = i
                if start_idx is not None and line.lstrip().startswith("BATCH:"):
                    body = lines[start_idx:i-1]
                    page = "\n".join(([ledger_header] if ledger_header else []) + body)
                    if ref_pattern.search(page): pages.append(page)
                    start_idx = None
            return pages

        def extract_payment_from_page(page_text, ref_no):
            lines = page_text.splitlines()
            for i, line in enumerate(lines):
                if re.search(rf"\b{re.escape(ref_no)}\b", line):
                    block = " ".join(lines[i:i+3])
                    m = re.search(r"\b(\d+)\s+(\d{2}/\d{2})\b", block)
                    if m: return m.group(1), m.group(2)
            return None, None

        def get_next_pay_cols(df, idx):
            col_idx = 7
            while True:
                pay_col = f"Payment_{(col_idx-5)//2}"; date_col = f"Date_{(col_idx-5)//2}"
                if pay_col not in df.columns: df[pay_col] = ""; df[date_col] = ""
                if not df.at[idx, pay_col]: return pay_col, date_col
                col_idx += 2

        def create_pdf_page(text, output_path, para_label, ref_no, amount, date_dd_mm):
            c = rl_canvas.Canvas(output_path, pagesize=PAGE_SIZE)
            w, h = PAGE_SIZE; lines = text.splitlines()
            if not lines: c.save(); return
            first_y = h - TOP_MARGIN; c.setFont("Courier", 9)
            first_width = c.stringWidth(lines[0], "Courier", 9); right_edge = LEFT_MARGIN + first_width
            c.setFont("Helvetica-Bold", 10); label_width = c.stringWidth(para_label, "Helvetica-Bold", 10)
            label_x = right_edge - label_width; label_y = first_y + LINE_GAP
            c.drawString(label_x, label_y, para_label); c.line(label_x, label_y - 1.5, label_x + label_width, label_y - 1.5)
            y = first_y - LINE_GAP; c.setFont("Courier", 9)
            for line in lines:
                if y < BOTTOM_MARGIN: break
                if ref_no in line and amount in line and date_dd_mm in line:
                    c.saveState(); c.setFillAlpha(0.25); c.setFillColor(colors.yellow)
                    wtxt = c.stringWidth(line[:260], "Courier", 9)
                    c.rect(LEFT_MARGIN - 2, y - 2, wtxt + 4, LINE_GAP + 2, 0, 1); c.restoreState()
                c.drawString(LEFT_MARGIN, y, line[:260]); y -= LINE_GAP
            c.save()

        df = pd.read_excel(self.excel_path, dtype=str)
        for col in ["Status", "Found in Files"]:
            if col not in df.columns: df[col] = ""
        total = len(df)

        for idx, row in df.iterrows():
            if self.stop_flag: self.log.emit("⛔ Stopped."); break
            batch = str(row["BN"]).zfill(2); sub_div = str(row["Sdiv"])
            ref_no = str(row["AC No."]); para_sr = str(row["Para Sr. No."])
            if not para_sr or para_sr.lower() == "nan": continue
            self.log.emit(f"\n🔍 BN {batch} | SDiv {sub_div} | AC No. {ref_no}")
            matched_files = [f for f in os.listdir(self.text_folder) if re.search(fr"-B{batch}-", f, re.IGNORECASE)]
            temp_pdfs, found_files, page_no = [], [], 1
            for fi, file in enumerate(matched_files, 1):
                self.log.emit(f"📂 {fi}/{len(matched_files)}: {file}")
                year_match = re.search(r"(\d{4})\.txt$", file)
                year = year_match.group(1) if year_match else ""
                with open(os.path.join(self.text_folder, file), "r", encoding="utf-8", errors="ignore") as f:
                    text = f.read()
                header = get_ledger_header(text)
                pages  = extract_matching_pages(text, ref_no, header)
                for page in pages:
                    if not re.search(fr"{sub_div}", page): continue
                    amount, date_dd_mm = extract_payment_from_page(page, ref_no)
                    if not amount: continue
                    self.log.emit(f"✅ Payment {page_no} | Amount: {amount} | Date: {date_dd_mm}/{year}")
                    pay_col, date_col = get_next_pay_cols(df, idx)
                    df.at[idx, pay_col] = amount; df.at[idx, date_col] = f"{date_dd_mm}/{year}"
                    label = f"Para Sr. No. {para_sr}" if page_no == 1 else f"Para Sr. No. {para_sr} ({page_no})"
                    tmp = os.path.join(OUTPUT_FOLDER, f"temp_{para_sr}_{page_no}.pdf")
                    create_pdf_page(page, tmp, label, ref_no, amount, date_dd_mm)
                    temp_pdfs.append(tmp); found_files.append(file); page_no += 1
            if temp_pdfs:
                final_pdf = os.path.join(OUTPUT_FOLDER, f"{para_sr}.pdf")
                merger = PdfMerger()
                for pp in temp_pdfs: merger.append(pp)
                merger.write(final_pdf); merger.close()
                for pp in temp_pdfs: os.remove(pp)
                df.at[idx, "Status"] = "Payment Found & Extracted"
                df.at[idx, "Found in Files"] = "; ".join(set(found_files))
                self.log.emit(f"📄 PDF created: {os.path.basename(final_pdf)}")
            else:
                df.at[idx, "Status"] = "No Payment"; self.log.emit("❌ No payment found")
            self.progress.emit(int((idx + 1) / total * 100))

        df.to_excel(self.excel_path, index=False)
        self.log.emit("🎯 PROCESS COMPLETED"); self.finished.emit()


class PaymentExtractPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("Payment Extractor", bold=True, color="#7eb8f7"))
        layout.addWidget(hline())

        g = QGroupBox("Inputs")
        gv = QVBoxLayout(g)
        label_w = 82
        r1 = QHBoxLayout(); lbl_excel = QLabel("Excel File:"); lbl_excel.setFixedWidth(label_w); r1.addWidget(lbl_excel)
        self.excel_ed = QLineEdit(); self.excel_ed.setReadOnly(True); self.excel_ed.setProperty("preferred_width", 245); r1.addWidget(self.excel_ed, 1)
        b1 = QPushButton("Browse"); b1.clicked.connect(lambda: (p := QFileDialog.getOpenFileName(self,"Excel","","Excel (*.xlsx)")[0]) and self.excel_ed.setText(p)); r1.addWidget(b1)
        btn_tpl = QPushButton("⬇  Download Template")
        btn_tpl.setStyleSheet("background:#194; border:none; border-radius:4px; padding:5px 10px; color:#fff;")
        btn_tpl.clicked.connect(self.get_template); r1.addWidget(btn_tpl)
        gv.addLayout(r1)
        r2 = QHBoxLayout(); lbl_text = QLabel("Text Folder:"); lbl_text.setFixedWidth(label_w); r2.addWidget(lbl_text)
        self.text_ed = QLineEdit(); self.text_ed.setReadOnly(True); self.text_ed.setProperty("preferred_width", 245); r2.addWidget(self.text_ed, 1)
        b2 = QPushButton("Browse"); b2.clicked.connect(lambda: (p := QFileDialog.getExistingDirectory(self,"Text Folder")) and self.text_ed.setText(p)); r2.addWidget(b2)
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

    def get_template(self):
        download_template(
            self,
            headers      = ["BN", "Sdiv", "AC No.", "Para Sr. No."],
            default_name = "PaymentExtract_Template.xlsx"
        )

    def start(self):
        if not self.excel_ed.text() or not self.text_ed.text():
            QMessageBox.warning(self, "Missing", "Select Excel and Text Folder first."); return
        self.worker = PaymentWorker(self.excel_ed.text(), self.text_ed.text())
        self.worker.log.connect(self.log_box.append)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(lambda: (self.btn_start.setEnabled(True), self.btn_stop.setEnabled(False)))
        self.worker.start()
        self.btn_start.setEnabled(False); self.btn_stop.setEnabled(True)

    def stop(self):
        if self.worker: self.worker.stop()


# ──────────────────────────────────────────────────────────────
#  TOOL 5 — Multi-Folder PDF Merge (multimerge.py)
# ──────────────────────────────────────────────────────────────
