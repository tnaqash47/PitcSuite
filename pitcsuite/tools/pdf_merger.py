import sys
import os
import re
import tempfile
import threading
from io import BytesIO
from pathlib import Path
from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtGui import QIcon, QPalette, QColor
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QFileDialog, QComboBox, QGroupBox,
    QProgressBar, QPlainTextEdit, QMessageBox, QCheckBox)
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

class Cancelled(Exception):
    pass

def natural_key(name):
    stem = Path(name).stem
    match = re.fullmatch(r"(\d+)(?:\s*\((\d+)\))?", stem)
    if match:
        return (0, int(match[1]), int(match[2] or 0))
    return (1, stem.lower(), 0)

def para_key(path):
    name = re.sub(r"\s*\(\d+\)$", "", path.stem).strip()
    try:
        return (0, int(name))
    except ValueError:
        return (1, name.lower())

def ordered_pdfs(folders, output, interleave=False):
    output = Path(output).resolve()
    groups = []
    seen = set()
    for folder in folders:
        folder = Path(folder).resolve()
        if folder in seen:
            raise ValueError("The same input folder was selected more than once.")
        seen.add(folder)
        if not folder.is_dir():
            raise ValueError(f"Input folder does not exist: {folder}")
        groups.append(sorted((p for p in folder.iterdir()
                              if p.is_file() and p.suffix.lower() == ".pdf"
                              and p.resolve() != output),
                             key=lambda p: (natural_key(p.name), p.name)))
    if not interleave:
        return [p for group in groups for p in group]
    numbered = []
    for folder_index, group in enumerate(groups):
        counts = {}
        for path in group:
            key = para_key(path)
            occurrence = counts.get(key, 0)
            counts[key] = occurrence + 1
            numbered.append((key, occurrence, folder_index, path))
    return [item[3] for item in sorted(numbered, key=lambda item: item[:3])]

def stamp_page(page, filename):
    width = float(page.mediabox.width)
    height = float(page.mediabox.height)
    label = f"Sr. No. {Path(filename).stem}"
    stream = BytesIO()
    overlay = canvas.Canvas(stream, pagesize=(width, height))
    overlay.setFont("Helvetica-Bold", 18)
    text_width = overlay.stringWidth(label, "Helvetica-Bold", 18)
    x = max(18, width - 18 - text_width)
    baseline = height - 30
    overlay.drawString(x, baseline, label)
    overlay.setLineWidth(1.2)
    overlay.line(x, baseline - 3, x + text_width, baseline - 3)
    overlay.save()
    stream.seek(0)
    page.merge_page(PdfReader(stream).pages[0])

def merge(folders, output, interleave=False, pages_per_side=0, stamp=False,
          cancelled=lambda: False, progress=lambda message, current, total: None):
    output = Path(output).resolve()
    paths = ordered_pdfs(folders, output, interleave)
    if not paths:
        raise ValueError("No input PDFs found.")
    writer = PdfWriter()
    temporary = None
    blanks = 0
    try:
        for index, path in enumerate(paths, 1):
            if cancelled():
                raise Cancelled()
            progress(f"Merging {path.parent.name} / {path.name}", index - 1, len(paths))
            try:
                with path.open("rb") as source:
                    reader = PdfReader(source)
                    if reader.is_encrypted and not reader.decrypt(""):
                        raise ValueError("Password-protected PDF")
                    if not len(reader.pages):
                        raise ValueError("PDF has no pages")
                    for page in reader.pages:
                        if cancelled():
                            raise Cancelled()
                        if stamp:
                            stamp_page(page, path.name)
                        writer.add_page(page)
                    padding = (-len(writer.pages)) % (2 * pages_per_side) if pages_per_side else 0
                    for _ in range(padding):
                        last = reader.pages[-1]
                        writer.add_blank_page(width=float(last.mediabox.width), height=float(last.mediabox.height))
                    blanks += padding
            except Cancelled:
                raise
            except Exception as exc:
                raise ValueError(f"Cannot merge {path}: {exc}") from exc
            progress(f"Added {path.name}", index, len(paths))
        if cancelled():
            raise Cancelled()
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".tmp", delete=False) as target:
            temporary = target.name
            writer.write(target)
        if cancelled():
            raise Cancelled()
        os.replace(temporary, output)
        temporary = None
        return f"Saved {len(paths)} PDFs ({len(writer.pages)} pages, {blanks} blank pages) to {output}"
    finally:
        writer.close()
        if temporary:
            Path(temporary).unlink(missing_ok=True)

class Worker(QThread):
    progress = Signal(str, int, int)
    result = Signal(str)
    def __init__(self, options, parent):
        super().__init__(parent)
        self.options = options
        self.stop = threading.Event()
    def run(self):
        try:
            message = merge(**self.options, cancelled=self.stop.is_set, progress=self.progress.emit)
        except Cancelled:
            message = "Stopped. No output was replaced."
        except Exception as exc:
            message = f"Merge failed: {exc}"
        self.result.emit(message)

class PDFMerger(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.setWindowTitle("PDF Merger")
        self.resize(540, 500)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)
        heading = QLabel("PDF Merger")
        heading.setStyleSheet("font-size: 13pt; font-weight: 600")
        layout.addWidget(heading)
        self.settings = QGroupBox("Merge settings")
        form = QVBoxLayout(self.settings)
        form.setSpacing(5)
        self.mode = QComboBox()
        self.mode.addItems(["Single Folder", "Multiple folders"])
        # Open on the original single-folder tool by default.
        self.mode.setCurrentIndex(0)
        form.addWidget(self.mode)
        self.hint = QLabel()
        self.hint.setWordWrap(True)
        form.addWidget(self.hint)
        self.folder_rows = []
        self.folders = []
        self.folder_labels = []
        for i in range(5):
            row_widget = QWidget()
            row = QHBoxLayout(row_widget)
            row.setContentsMargins(0, 0, 0, 0)
            folder_label = QLabel(f"Folder {i + 1}")
            row.addWidget(folder_label)
            self.folder_labels.append(folder_label)
            edit = QLineEdit()
            edit.setPlaceholderText("Select or paste an input folder")
            row.addWidget(edit, 1)
            button = QPushButton("Browse")
            button.setProperty("role", "secondary")
            button.clicked.connect(lambda checked=False, field=edit: self.browse_folder(field))
            row.addWidget(button)
            form.addWidget(row_widget)
            self.folder_rows.append(row_widget)
            self.folders.append(edit)
        row = QHBoxLayout()
        row.addWidget(QLabel("Output PDF"))
        self.output = QLineEdit()
        self.output.setPlaceholderText("Choose where to save")
        row.addWidget(self.output, 1)
        button = QPushButton("Save as")
        button.setProperty("role", "secondary")
        button.clicked.connect(self.browse_output)
        row.addWidget(button)
        form.addLayout(row)
        self.padding = QComboBox()
        self.padding.addItem("No blank pages", 0)
        self.padding.addItem("Back-to-back / duplex (1 page per side)", 1)
        for count in (2, 4, 6, 9, 16):
            self.padding.addItem(f"N-up duplex ({count} pages per side)", count)
        form.addWidget(self.padding)
        self.stamp = QCheckBox("Stamp: Add filename at top right")
        form.addWidget(self.stamp)
        note = QLabel("Duplex adds blanks between PDFs. For N-up, also set pages per side when printing.")
        note.setWordWrap(True)
        form.addWidget(note)
        layout.addWidget(self.settings)
        row = QHBoxLayout()
        self.start = QPushButton("Merge PDFs")
        self.start.setProperty("role", "primary")
        self.start.clicked.connect(self.begin)
        self.stop = QPushButton("Stop")
        self.stop.setProperty("role", "secondary")
        self.stop.setEnabled(False)
        self.stop.clicked.connect(self.cancel)
        row.addWidget(self.start)
        row.addWidget(self.stop)
        layout.addLayout(row)
        self.progress = QProgressBar()
        layout.addWidget(self.progress)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(1000)
        self.log.setMinimumHeight(60)
        self.log.setMaximumHeight(90)
        layout.addWidget(self.log, 1)
        self.mode.currentIndexChanged.connect(self.update_mode)
        self.update_mode()
    def update_mode(self):
        multi = bool(self.mode.currentIndex())
        self.folder_labels[0].setText("Folder 1" if multi else "Folder")
        for i, row in enumerate(self.folder_rows):
            row.setVisible(multi or i == 0)
        self.hint.setText("Choose up to 5 folders." if multi else "Choose folder:")
    def browse_folder(self, field):
        folder = QFileDialog.getExistingDirectory(self, "Input folder", field.text())
        if folder:
            field.setText(folder)
    def browse_output(self):
        name, _ = QFileDialog.getSaveFileName(self, "Output PDF", self.output.text() or self.default_output(), "PDF files (*.pdf)")
        if name:
            self.output.setText(name if name.lower().endswith(".pdf") else name + ".pdf")
    def default_output(self):
        if self.mode.currentIndex():
            return str(Path("D:/MergedPDFs") / "Merged_Final_Order.pdf")
        folder = Path(self.folders[0].text().strip())
        return str(folder / f"Merged_{folder.name}.pdf")
    def begin(self):
        fields = self.folders if self.mode.currentIndex() else self.folders[:1]
        folders = [field.text().strip() for field in fields if field.text().strip()]
        if not folders or any(not Path(folder).is_dir() for folder in folders):
            QMessageBox.warning(self, "Input folders", "Select valid input folders first.")
            return
        output = Path(self.output.text().strip() or self.default_output()).absolute()
        if output.suffix.lower() != ".pdf":
            QMessageBox.warning(self, "Output PDF", "The output filename must end in .pdf.")
            return
        if output.exists() and QMessageBox.question(self, "Replace output?", f"Replace this existing file?\n{output}") != QMessageBox.StandardButton.Yes:
            return
        self.output.setText(str(output))
        self.log.clear()
        self.progress.setValue(0)
        self.settings.setEnabled(False)
        self.start.setEnabled(False)
        self.stop.setEnabled(True)
        options = dict(folders=folders, output=str(output), interleave=bool(self.mode.currentIndex()), pages_per_side=self.padding.currentData(), stamp=self.stamp.isChecked())
        self.worker = Worker(options, self)
        self.worker.progress.connect(self.on_progress)
        self.worker.result.connect(self.log.appendPlainText)
        self.worker.finished.connect(self.finished)
        self.worker.start()
    def on_progress(self, message, current, total):
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.log.appendPlainText(message)
    def cancel(self):
        if self.worker:
            self.worker.stop.set()
            self.stop.setEnabled(False)
    def finished(self):
        self.settings.setEnabled(True)
        self.start.setEnabled(True)
        self.stop.setEnabled(False)
        self.worker.deleteLater()
        self.worker = None
    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.cancel()
            self.log.appendPlainText("Stopping merge. Close the window after it finishes.")
            event.ignore()
        else:
            event.accept()
