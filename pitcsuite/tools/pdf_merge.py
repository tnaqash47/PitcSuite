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
        self._updating_print_mode = False
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

        print_group = QGroupBox("Print layout")
        print_layout = QVBoxLayout(print_group)

        self.add_blank_pages_cb = QCheckBox(
            "Back-to-back only: add 1 white page after an odd-page PDF"
        )
        self.add_blank_pages_cb.setToolTip(
            "Use this for ordinary duplex printing: each merged PDF is padded to an even page count."
        )
        self.add_blank_pages_cb.toggled.connect(
            lambda checked: self._on_print_mode_toggled(self.add_blank_pages_cb, checked)
        )
        print_layout.addWidget(self.add_blank_pages_cb)

        nup_row = QHBoxLayout()
        self.multiple_pages_cb = QCheckBox(
            "Multiple pages per side (N-up duplex; add required white pages)"
        )
        self.multiple_pages_cb.setToolTip(
            "Use this when the printer will place multiple PDF pages on each side. It pads each merged PDF to a complete front-and-back sheet."
        )
        self.multiple_pages_cb.toggled.connect(
            lambda checked: self._on_print_mode_toggled(self.multiple_pages_cb, checked)
        )
        nup_row.addWidget(self.multiple_pages_cb)
        nup_row.addWidget(QLabel("Pages per side:"))
        self.pages_per_side = QComboBox()
        for value, label in (
            (1, "1 (1x2 duplex)"),
            (2, "2 (2x2 duplex)"),
            (4, "4 (4x2 duplex)"),
            (6, "6 (6x2 duplex)"),
            (9, "9 (9x2 duplex)"),
            (16, "16 (16x2 duplex)"),
        ):
            self.pages_per_side.addItem(label, value)
        # Default N-up mode to 2 pages per side (2x2 duplex). The 1x2
        # back-to-back layout remains available as an explicit choice.
        self.pages_per_side.setCurrentIndex(self.pages_per_side.findData(2))
        self.pages_per_side.setEnabled(False)
        nup_row.addWidget(self.pages_per_side)
        nup_row.addStretch()
        print_layout.addLayout(nup_row)

        self.print_hint = QLabel(
            "Select one mode. Back-to-back adds one blank page only when a PDF has an odd number of pages. "
            "N-up adds however many blank pages are needed for the selected layout."
        )
        self.print_hint.setWordWrap(True)
        self.print_hint.setStyleSheet("color:#667085;")
        print_layout.addWidget(self.print_hint)
        layout.addWidget(print_group)

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
            if self.multiple_pages_cb.isChecked():
                add_blank_pages = True
                pages_per_side = self.pages_per_side.currentData()
            else:
                add_blank_pages = self.add_blank_pages_cb.isChecked()
                pages_per_side = 1
            threading.Thread(
                target=self.merge_worker,
                args=(add_blank_pages, pages_per_side),
                daemon=True,
            ).start()
        else:
            self.stop_flag["stop"] = True

    def _on_print_mode_toggled(self, source, checked):
        if self._updating_print_mode:
            return
        self._updating_print_mode = True
        try:
            if checked:
                other = (
                    self.multiple_pages_cb
                    if source is self.add_blank_pages_cb
                    else self.add_blank_pages_cb
                )
                other.setChecked(False)
            self.pages_per_side.setEnabled(self.multiple_pages_cb.isChecked())
        finally:
            self._updating_print_mode = False

    def on_update(self, msg, cur, total, done):
        self.log_lbl.setText(msg)
        if total: self.progress.setMaximum(total); self.progress.setValue(cur); self.lbl_prog.setText(f"{cur} / {total}")
        if done: self.btn_start.setText("▶  START"); self.btn_start.setStyleSheet(START_BTN)

    @staticmethod
    def blank_pages_needed(page_count, pages_per_side=1, pages_before=0):
        """Return padding needed after a PDF to align the next PDF to a new sheet."""
        pages_per_sheet = max(1, int(pages_per_side)) * 2
        merged_page_count = int(pages_before) + int(page_count)
        return (-merged_page_count) % pages_per_sheet

    def merge_worker(self, add_blank_pages=False, pages_per_side=1):
        from pypdf import PdfReader, PdfWriter
        folder = self.folder_ed.text()
        out_name = f"Merged_{os.path.basename(folder)}.pdf"
        out_path = os.path.join(folder, out_name)
        pdfs = sorted(
            [
                f for f in os.listdir(folder)
                if f.lower().endswith(".pdf")
                and os.path.abspath(os.path.join(folder, f)) != os.path.abspath(out_path)
            ],
            key=self.natural_key,
        )
        if not pdfs: self.signals.update.emit("No PDFs found.", 0, 0, True); return
        writer = PdfWriter(); total = len(pdfs); blank_total = 0; merged_page_count = 0; failed = []
        for i, pdf in enumerate(pdfs, 1):
            if self.stop_flag["stop"]: self.signals.update.emit("Stopped.", i, total, True); return
            self.signals.update.emit(f"Merging: {pdf}", i, total, False)
            try:
                reader = PdfReader(os.path.join(folder, pdf))
                for page in reader.pages: writer.add_page(page)
                page_count = len(reader.pages)
                merged_page_count += page_count
                if add_blank_pages:
                    blanks = self.blank_pages_needed(
                        page_count,
                        pages_per_side,
                        pages_before=merged_page_count - page_count,
                    )
                    for _ in range(blanks):
                        writer.add_blank_page(
                            width=float(reader.pages[-1].mediabox.width),
                            height=float(reader.pages[-1].mediabox.height),
                        )
                    blank_total += blanks
                    merged_page_count += blanks
                    self.signals.update.emit(
                        f"Merging: {pdf} ({page_count} page(s), added {blanks} white page(s))",
                        i,
                        total,
                        False,
                    )
            except Exception as exc:
                failed.append(f"{pdf}: {exc}")
        with open(out_path, "wb") as f: writer.write(f)
        summary = f"Saved: {out_name} ✔"
        if add_blank_pages:
            summary += f" Added {blank_total} white page(s) for {pages_per_side} page(s) per side."
        if failed:
            summary += " Failed: " + "; ".join(failed)
        self.signals.update.emit(summary, total, total, True)


# ──────────────────────────────────────────────────────────────
#  TOOL 10 — File Existence Checker (SrNoCheckinfolder.py)
# ──────────────────────────────────────────────────────────────
