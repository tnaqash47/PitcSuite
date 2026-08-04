import os
import shutil
import threading
from pathlib import Path

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QFileDialog, QDoubleSpinBox, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
    QListWidget, QMessageBox, QPushButton, QTextEdit, QVBoxLayout, QWidget,
)

from pitcsuite.ui_helpers import START_BTN, STOP_BTN, hline, lbl, notify


class CP22TWorker(QThread):
    log = Signal(str)
    done = Signal(str)
    failed = Signal(str)

    def __init__(self, xlsx_path, pdf_paths):
        super().__init__()
        self.xlsx_path = xlsx_path or None
        self.pdf_paths = pdf_paths

    def run(self):
        try:
            from pitcsuite.tools.cp22t_posting_checker import run_check
            output = run_check(self.xlsx_path, self.pdf_paths, self.log.emit)
            self.done.emit(str(output))
        except Exception as exc:
            self.failed.emit(str(exc))


class CP22TCheckerPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.pdf_paths = []
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("CP22T Posting Checker", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Match CP22T PDF transaction records against an optional Excel list.", color="#556"))
        layout.addWidget(hline())

        group = QGroupBox("Inputs")
        form = QVBoxLayout(group)
        row = QHBoxLayout(); row.addWidget(QLabel("Excel (optional):"))
        self.xlsx_ed = QLineEdit(); self.xlsx_ed.setReadOnly(True); self.xlsx_ed.setProperty("preferred_width", 245)
        row.addWidget(self.xlsx_ed, 1)
        browse = QPushButton("Browse"); browse.clicked.connect(self._browse_excel); row.addWidget(browse)
        form.addLayout(row)
        form.addWidget(QLabel("CP22T PDFs:"))
        self.pdf_list = QListWidget(); self.pdf_list.setMinimumHeight(110); form.addWidget(self.pdf_list)
        buttons = QHBoxLayout()
        add = QPushButton("Add PDFs"); add.clicked.connect(self._add_pdfs); buttons.addWidget(add)
        clear = QPushButton("Clear"); clear.clicked.connect(self._clear_pdfs); buttons.addWidget(clear); buttons.addStretch()
        form.addLayout(buttons)
        layout.addWidget(group)

        actions = QHBoxLayout()
        self.start_btn = QPushButton("▶  START"); self.start_btn.setStyleSheet(START_BTN); self.start_btn.clicked.connect(self._start)
        self.stop_btn = QPushButton("■  STOP"); self.stop_btn.setStyleSheet(STOP_BTN); self.stop_btn.setEnabled(False); self.stop_btn.clicked.connect(self._stop)
        actions.addWidget(self.start_btn); actions.addWidget(self.stop_btn); actions.addStretch(); layout.addLayout(actions)
        self.status = QLabel("Status: Idle"); self.status.setStyleSheet("color:#4a90d9;"); layout.addWidget(self.status)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(self.log_box)

    def _browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel input", "", "Excel files (*.xlsx *.xlsm);;All files (*.*)")
        if path: self.xlsx_ed.setText(path)

    def _add_pdfs(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "CP22T PDFs", "", "PDF files (*.pdf);;All files (*.*)")
        for path in paths:
            if path not in self.pdf_paths:
                self.pdf_paths.append(path); self.pdf_list.addItem(path)

    def _clear_pdfs(self):
        self.pdf_paths.clear(); self.pdf_list.clear()

    def _start(self):
        if not self.pdf_paths:
            QMessageBox.warning(self, "Missing PDFs", "Please select at least one CP22T PDF."); return
        self.log_box.clear(); self.status.setText("Status: Working...")
        self.start_btn.setEnabled(False); self.stop_btn.setEnabled(True)
        self.worker = CP22TWorker(self.xlsx_ed.text().strip(), list(self.pdf_paths))
        self.worker.log.connect(self._log); self.worker.done.connect(self._done); self.worker.failed.connect(self._failed)
        self.worker.finished.connect(lambda: (self.start_btn.setEnabled(True), self.stop_btn.setEnabled(False)))
        self.worker.start()

    def _stop(self):
        if self.worker and self.worker.isRunning():
            self.worker.requestInterruption(); self.status.setText("Status: Finishing current operation...")

    def _log(self, message): self.log_box.append(message)
    def _done(self, output):
        self.status.setText("Status: Completed"); notify(self, "Completed", f"Output saved to:\n{output}")
    def _failed(self, message):
        self.status.setText("Status: Failed"); notify(self, "Error", message, critical=True)


class BillScraperWorker(QThread):
    progress = Signal(int, int, str, str)
    done = Signal(str, bool)
    failed = Signal(str)

    def __init__(self, path, delay, stop_event):
        super().__init__(); self.path = Path(path); self.delay = delay; self.stop_event = stop_event

    def run(self):
        try:
            from pitcsuite.tools.bill_scraper_logic import PITCBillScraper
            def progress(done, total, row):
                self.progress.emit(done, total, row.get("_account", ""), row.get("scrape_status", ""))
            PITCBillScraper(delay=self.delay).update_workbook(self.path, progress, self.stop_event)
            self.done.emit(str(self.path), self.stop_event.is_set())
        except Exception as exc:
            self.failed.emit(str(exc))


class BillScraperPanel(QWidget):
    def __init__(self):
        super().__init__(); self.worker = None; self.stop_event = threading.Event()
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("PITC Bill Scrapper", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Update an Excel workbook with live PITC bill details.", color="#556")); layout.addWidget(hline())
        group = QGroupBox("Workbook")
        form = QVBoxLayout(group); row = QHBoxLayout(); row.addWidget(QLabel("Excel file:"))
        self.xlsx_ed = QLineEdit(); self.xlsx_ed.setReadOnly(True); self.xlsx_ed.setProperty("preferred_width", 245); row.addWidget(self.xlsx_ed, 1)
        browse = QPushButton("Browse"); browse.clicked.connect(self._browse); row.addWidget(browse); form.addLayout(row)
        options = QHBoxLayout(); options.addWidget(QLabel("Delay between bills (seconds):"))
        self.delay = QDoubleSpinBox(); self.delay.setRange(0, 30); self.delay.setSingleStep(.1); self.delay.setValue(5); options.addWidget(self.delay); options.addStretch(); form.addLayout(options)
        layout.addWidget(group)
        actions = QHBoxLayout(); self.start_btn = QPushButton("▶  START"); self.start_btn.setStyleSheet(START_BTN); self.start_btn.clicked.connect(self._start)
        self.stop_btn = QPushButton("■  STOP"); self.stop_btn.setStyleSheet(STOP_BTN); self.stop_btn.setEnabled(False); self.stop_btn.clicked.connect(self._stop)
        actions.addWidget(self.start_btn); actions.addWidget(self.stop_btn); actions.addStretch(); layout.addLayout(actions)
        self.progress = QLabel("Processed: 0"); layout.addWidget(self.progress)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(self.log_box)

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose workbook", "", "Excel workbook (*.xlsx);;All files (*.*)")
        if path: self.xlsx_ed.setText(path)

    def _start(self):
        path = Path(self.xlsx_ed.text().strip())
        if not path.is_file(): QMessageBox.warning(self, "Input required", "Choose an existing .xlsx workbook first."); return
        if not QMessageBox.question(self, "Update workbook", "The workbook will be updated in place. Continue?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes: return
        self.stop_event.clear(); self.log_box.clear(); self.start_btn.setEnabled(False); self.stop_btn.setEnabled(True)
        self.worker = BillScraperWorker(path, self.delay.value(), self.stop_event)
        self.worker.progress.connect(self._progress); self.worker.done.connect(self._done); self.worker.failed.connect(self._failed); self.worker.start()

    def _stop(self): self.stop_event.set(); self.progress.setText("Stopping after the current bill...")
    def _progress(self, done, total, account, status):
        self.progress.setText(f"Processed {done} of {total} | {account}: {status}"); self.log_box.append(f"{account}: {status}")
    def _done(self, path, cancelled):
        self.start_btn.setEnabled(True); self.stop_btn.setEnabled(False); self.progress.setText("Stopped." if cancelled else "Completed")
        if not cancelled: notify(self, "Completed", f"Workbook updated:\n{path}")
    def _failed(self, message):
        self.start_btn.setEnabled(True); self.stop_btn.setEnabled(False); notify(self, "Error", message, critical=True)


class FileMoverWorker(QThread):
    done = Signal(int, list)
    failed = Signal(str)

    def __init__(self, excel, source, destination):
        super().__init__(); self.excel = excel; self.source = source; self.destination = destination

    def run(self):
        try:
            import pandas as pd
            df = pd.read_excel(self.excel, header=None)
            files = {}
            for name in os.listdir(self.source):
                path = os.path.join(self.source, name)
                if os.path.isfile(path): files[os.path.splitext(name)[0].strip()] = name
            moved = 0; missing = []
            for value in df.iloc[:, 0].dropna():
                sr = str(int(value)) if isinstance(value, (int, float)) and float(value).is_integer() else str(value).strip()
                if sr.endswith(".0"): sr = sr[:-2]
                if sr in files:
                    shutil.move(os.path.join(self.source, files[sr]), os.path.join(self.destination, files[sr])); moved += 1
                else: missing.append(sr)
            self.done.emit(moved, missing)
        except Exception as exc: self.failed.emit(str(exc))


class FileMoverPanel(QWidget):
    def __init__(self):
        super().__init__(); self.worker = None
        layout = QVBoxLayout(self); layout.setSpacing(10); layout.addWidget(lbl("PDF / File Mover", bold=True, color="#7eb8f7")); layout.addWidget(hline())
        group = QGroupBox("Inputs"); form = QVBoxLayout(group)
        self.excel_ed, self.source_ed, self.destination_ed = QLineEdit(), QLineEdit(), QLineEdit()
        for label, edit, kind in (("Excel file:", self.excel_ed, "excel"), ("Source folder:", self.source_ed, "source"), ("Destination folder:", self.destination_ed, "destination")):
            row = QHBoxLayout(); row.addWidget(QLabel(label)); edit.setReadOnly(True); edit.setProperty("preferred_width", 245); row.addWidget(edit, 1); button = QPushButton("Browse"); button.clicked.connect(lambda _, e=edit, k=kind: self._browse(e, k)); row.addWidget(button); form.addLayout(row)
        layout.addWidget(group)
        self.start_btn = QPushButton("▶  MOVE FILES"); self.start_btn.setStyleSheet(START_BTN); self.start_btn.clicked.connect(self._start); layout.addWidget(self.start_btn)
        self.status = QLabel("Status: Idle"); layout.addWidget(self.status); layout.addStretch()

    def _browse(self, edit, kind):
        path = QFileDialog.getOpenFileName(self, "Excel file", "", "Excel files (*.xlsx *.xls)")[0] if kind == "excel" else QFileDialog.getExistingDirectory(self, "Select folder")
        if path: edit.setText(path)

    def _start(self):
        if not all((self.excel_ed.text(), self.source_ed.text(), self.destination_ed.text())):
            QMessageBox.warning(self, "Missing input", "Select the Excel file, source folder, and destination folder."); return
        self.start_btn.setEnabled(False); self.status.setText("Status: Moving files...")
        self.worker = FileMoverWorker(self.excel_ed.text(), self.source_ed.text(), self.destination_ed.text())
        self.worker.done.connect(self._done); self.worker.failed.connect(self._failed); self.worker.finished.connect(lambda: self.start_btn.setEnabled(True)); self.worker.start()

    def _done(self, moved, missing):
        self.status.setText(f"Status: Completed — moved {moved} file(s)")
        message = f"Moved {moved} file(s).\n\n" + ("Missing files:\n\n" + "\n".join(missing) if missing else "All files moved successfully.")
        notify(self, "Completed", message)

    def _failed(self, message): self.status.setText("Status: Failed"); notify(self, "Error", message, critical=True)


class CP52Worker(QThread):
    log = Signal(str)
    done = Signal(str)
    failed = Signal(str)

    def __init__(self, source_dir, workbook, output_dir):
        super().__init__(); self.source_dir = Path(source_dir); self.workbook = Path(workbook); self.output_dir = Path(output_dir)

    def run(self):
        try:
            # Reload during development so a long-running VS Code session
            # does not keep using an older CP-52 renderer after the source
            # file has been edited.
            import importlib
            from pitcsuite.tools import cp52_posting_checker_logic
            logic = importlib.reload(cp52_posting_checker_logic)
            result = logic.process_workbook(self.source_dir, self.workbook, self.output_dir, self.log.emit)
            self.done.emit(str(result))
        except Exception as exc:
            self.failed.emit(str(exc))


class CP52PostingCheckerPanel(QWidget):
    def __init__(self):
        super().__init__(); self.worker = None
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("CP-52 Posting Checker", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Match CP-52 posting records from TXT/PDF files and create highlighted PDFs.", color="#556")); layout.addWidget(hline())
        group = QGroupBox("Inputs")
        form = QVBoxLayout(group)
        self.source_ed, self.workbook_ed, self.output_ed = QLineEdit(), QLineEdit(), QLineEdit()
        for label, edit, kind in (("Posting folder:", self.source_ed, "folder"), ("Excel input:", self.workbook_ed, "excel"), ("Output folder:", self.output_ed, "folder")):
            row = QHBoxLayout(); row.addWidget(QLabel(label)); edit.setProperty("preferred_width", 245); row.addWidget(edit, 1)
            button = QPushButton("Browse"); button.clicked.connect(lambda _, e=edit, k=kind: self._browse(e, k)); row.addWidget(button); form.addLayout(row)
        layout.addWidget(group)
        actions = QHBoxLayout(); self.start_btn = QPushButton("▶  CHECK POSTING"); self.start_btn.setStyleSheet(START_BTN); self.start_btn.clicked.connect(self._start); actions.addWidget(self.start_btn)
        open_btn = QPushButton("Open Output"); open_btn.clicked.connect(self._open_output); actions.addWidget(open_btn); actions.addStretch(); layout.addLayout(actions)
        self.status = QLabel("Status: Ready"); self.status.setStyleSheet("color:#4a90d9;"); layout.addWidget(self.status)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(self.log_box)

    def _browse(self, edit, kind):
        if kind == "excel":
            path = QFileDialog.getOpenFileName(self, "Excel input", "", "Excel files (*.xlsx);;All files (*.*)")[0]
        else:
            path = QFileDialog.getExistingDirectory(self, "Select folder")
        if path: edit.setText(path)

    def _start(self):
        source, workbook = Path(self.source_ed.text().strip()), Path(self.workbook_ed.text().strip())
        if not source.is_dir() or not workbook.is_file():
            QMessageBox.warning(self, "Invalid input", "Select an existing posting folder and Excel workbook first."); return
        output = Path(self.output_ed.text().strip()) if self.output_ed.text().strip() else source / "output"
        self.output_ed.setText(str(output)); self.log_box.clear(); self.status.setText("Status: Processing..."); self.start_btn.setEnabled(False)
        self.worker = CP52Worker(source, workbook, output)
        self.worker.log.connect(self._log); self.worker.done.connect(self._done); self.worker.failed.connect(self._failed); self.worker.finished.connect(lambda: self.start_btn.setEnabled(True)); self.worker.start()

    def _open_output(self):
        path = Path(self.output_ed.text().strip()) if self.output_ed.text().strip() else Path.cwd() / "output"
        path.mkdir(parents=True, exist_ok=True); os.startfile(str(path))

    def _log(self, message): self.log_box.append(message)
    def _done(self, result):
        self.status.setText(f"Status: Completed — {result}"); notify(self, "Completed", f"Excel updated:\n{result}")
    def _failed(self, message):
        self.status.setText("Status: Failed"); notify(self, "Processing error", message, critical=True)
