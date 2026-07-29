from __future__ import annotations

import os
import re
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import END, BOTH, LEFT, RIGHT, X, Y, StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import Color, black, yellow
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
import pdfplumber


ACCOUNT_DIGITS = re.compile(r"(?<![0-9])[0-9]{13,15}(?![0-9])")
ACCOUNT_LINE = re.compile(r"(?m)^\s*[0-9]{14}\s*$")
AMOUNT_TOKEN = re.compile(r"(?<![0-9])-?[0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{1,2})?(?![0-9])|-?(?<![0-9])[0-9]+(?:\.[0-9]{1,2})?(?![0-9])")
HIGHLIGHT = Color(1.0, 0.88, 0.15, alpha=0.40)
SMART_AMOUNT_TOLERANCE = 1.00


@dataclass
class SourcePage:
    number: int
    text: str
    source_index: int
    source_path: Path


@dataclass
class MatchResult:
    page: SourcePage | None
    account_hits: list[tuple[int, int]]
    amount_hits: list[tuple[int, int]]
    amount_found: bool
    amount_exact: bool


def compact_digits(value: object) -> str:
    digits = re.sub(r"[^0-9]", "", "" if value is None else str(value))
    # Excel often drops a leading zero when a fixed-width AC number is entered as a number.
    return digits.zfill(14) if len(digits) == 13 else digits


def dashed_account(value: object) -> str:
    digits = compact_digits(value)
    if len(digits) < 9:
        return ""
    return f"{digits[:2]}-{digits[2:-7]}-{digits[-7:]}"


def numeric_value(value: object) -> float | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        return round(float(str(value).replace(",", "").strip()), 2)
    except ValueError:
        return None


def amount_matches(token: str, wanted: float) -> bool:
    try:
        return round(float(token.replace(",", "")), 2) == wanted
    except ValueError:
        return False


def amount_spans(text: str, wanted: float) -> tuple[list[tuple[int, int]], bool]:
    candidates = []
    for m in AMOUNT_TOKEN.finditer(text):
        try:
            difference = abs(float(m.group().replace(",", "")) - wanted)
        except ValueError:
            continue
        if difference <= SMART_AMOUNT_TOLERANCE:
            candidates.append((difference, m.start(), m.end()))
    if not candidates:
        return [], False
    closest = min(item[0] for item in candidates)
    selected = [(start, end) for difference, start, end in candidates if difference == closest]
    return selected, closest == 0


def split_text_pages(path: Path) -> list[SourcePage]:
    text = path.read_text(encoding="utf-8", errors="replace")
    chunks = text.split("\f")
    if len(chunks) > 1:
        return [SourcePage(i + 1, chunk, i, path) for i, chunk in enumerate(chunks) if chunk.strip()]
    starts = [m.start() for m in re.finditer(r"WAPDA\s+ELECTRICITY", text, re.IGNORECASE)]
    if not starts:
        return [SourcePage(1, text, 0, path)]
    pages = []
    for i, start in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(text)
        pages.append(SourcePage(i + 1, text[start:end], i, path))
    return pages


def pdf_pages(path: Path) -> list[SourcePage]:
    reader = PdfReader(str(path))
    result = []
    for i, page in enumerate(reader.pages):
        result.append(SourcePage(i + 1, page.extract_text() or "", i, path))
    return result


def find_match(pages: list[SourcePage], raw_account: object, amount: object) -> MatchResult:
    dashed = dashed_account(raw_account)
    digits = compact_digits(raw_account)
    if not dashed:
        return MatchResult(None, [], [], False, False)
    wanted_amount = numeric_value(amount)
    account_only_result = None
    near_result = None
    for page_index, page in enumerate(pages):
        account_hits = []
        for pattern in (re.escape(dashed), re.escape(digits)):
            account_hits.extend((m.start(), m.end()) for m in re.finditer(pattern, page.text, re.IGNORECASE))
        if account_hits:
            amount_hits, amount_exact = amount_spans(page.text, wanted_amount) if wanted_amount is not None else ([], False)
            result = MatchResult(page, account_hits, amount_hits, wanted_amount is None or bool(amount_hits), amount_exact)
            if wanted_amount is None:
                return result
            if amount_exact:
                return result
            if amount_hits and near_result is None:
                near_result = result
            if account_only_result is None:
                account_only_result = MatchResult(page, account_hits, [], False, False)
            # A fixed-width TXT export can split the last account line from its
            # details. Include the beginning of the next TXT page when testing
            # the amount, so the output page can show the complete record.
            if wanted_amount is not None and not amount_hits and page.source_path.suffix.lower() == ".txt":
                if page_index + 1 < len(pages):
                    next_page = pages[page_index + 1]
                    if next_page.source_path == page.source_path and next_page.source_index == page.source_index + 1:
                        continuation = ACCOUNT_LINE.search(next_page.text)
                        if continuation:
                            combined_text = page.text.rstrip() + "\n\n" + next_page.text[:continuation.start()]
                            combined_page = SourcePage(page.number, combined_text, page.source_index, page.source_path)
                            combined_account_hits = []
                            for pattern in (re.escape(dashed), re.escape(digits)):
                                combined_account_hits.extend((m.start(), m.end()) for m in re.finditer(pattern, combined_text, re.IGNORECASE))
                            combined_amount_hits, combined_amount_exact = amount_spans(combined_text, wanted_amount)
                            if combined_amount_hits:
                                return MatchResult(combined_page, combined_account_hits, combined_amount_hits, True, combined_amount_exact)
    if near_result is not None:
        return near_result
    if account_only_result is not None:
        return account_only_result
    return MatchResult(None, [], [], False, False)


def add_sr_overlay(page, sr: object, output: Path) -> None:
    if sr is None or str(sr).strip() == "":
        return
    writer = PdfWriter()
    with open(output, "rb") as f:
        source = PdfReader(f)
        page_obj = source.pages[0]
        width = float(page_obj.mediabox.width)
        height = float(page_obj.mediabox.height)
        overlay_path = output.with_suffix(".overlay.pdf")
        c = canvas.Canvas(str(overlay_path), pagesize=(width, height))
        label = str(sr).strip()
        c.setFont("Helvetica-Bold", 16)
        text_width = c.stringWidth(label, "Helvetica-Bold", 16)
        x = width - 24 - text_width
        y = height - 28
        c.setFillColor(black)
        c.drawString(x, y, label)
        c.setLineWidth(1.2)
        c.line(x, y - 2, x + text_width, y - 2)
        c.save()
        overlay = PdfReader(str(overlay_path)).pages[0]
        page_obj.merge_page(overlay)
        writer.add_page(page_obj)
        with open(output, "wb") as out:
            writer.write(out)
    overlay_path.unlink(missing_ok=True)


def render_text_page(page: SourcePage, result: MatchResult, sr: object, output: Path) -> None:
    # Posting-list TXT output is intended for landscape A4 printing.
    page_width, page_height = landscape(A4)
    margin_x, margin_y = 12, 18
    lines = page.text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    max_len = max((len(line) for line in lines), default=1)
    font_size = min(8.5, max(4.5, (page_width - 2 * margin_x) / (max_len * 0.60)))
    leading = min(8.2, max(5.1, (page_height - 2 * margin_y) / max(len(lines), 1)))
    c = canvas.Canvas(str(output), pagesize=(page_width, page_height))
    c.setFont("Courier", font_size)
    for line_no, line in enumerate(lines):
        y = page_height - margin_y - (line_no + 1) * leading
        if y < 10:
            break
        spans = []
        for start, end in result.account_hits + result.amount_hits:
            # absolute spans are mapped to the current line below
            pass
        line_start = sum(len(x) + 1 for x in lines[:line_no])
        line_end = line_start + len(line)
        for start, end in result.account_hits + result.amount_hits:
            if start < line_end and end > line_start:
                local_start, local_end = max(start, line_start) - line_start, min(end, line_end) - line_start
                spans.append((local_start, local_end))
        for start, end in spans:
            x = margin_x + start * font_size * 0.60
            w = max((end - start) * font_size * 0.60, 3)
            c.setFillColor(HIGHLIGHT)
            c.rect(x - 1, y - 1, w + 2, leading, stroke=0, fill=1)
        c.setFillColor(black)
        c.drawString(margin_x, y, line)
    if sr is not None and str(sr).strip():
        label = str(sr).strip()
        c.setFont("Helvetica-Bold", 16)
        width = c.stringWidth(label, "Helvetica-Bold", 16)
        x, y = page_width - 24 - width, page_height - 28
        c.drawString(x, y, label)
        c.setLineWidth(1.2)
        c.line(x, y - 2, x + width, y - 2)
    c.save()


def render_pdf_page(source_path: Path, page_number: int, result: MatchResult, account: object, sr: object, output: Path) -> None:
    reader = PdfReader(str(source_path))
    original = reader.pages[page_number - 1]
    width, height = float(original.mediabox.width), float(original.mediabox.height)
    with pdfplumber.open(str(source_path)) as pdf:
        words = pdf.pages[page_number - 1].extract_words()
    targets = [dashed_account(account), compact_digits(account)]
    wanted = []
    raw_account = result.page.text
    # Match the requested account directly against extracted words; this avoids highlighting unrelated numbers.
    for word in words:
        normalized = word["text"].replace(" ", "")
        if normalized in targets:
            wanted.append((word["x0"] - 1, height - word["bottom"] - 1, word["x1"] - word["x0"] + 2, word["bottom"] - word["top"] + 2))
    wanted_amount = None
    # The amount result spans are sufficient to identify that an amount exists; highlight matching numeric words.
    amount_value = None
    for m in AMOUNT_TOKEN.finditer(raw_account):
        if m.start() in {s for s, _ in result.amount_hits}:
            amount_value = numeric_value(m.group())
            break
    if amount_value is not None:
        for word in words:
            if amount_matches(word["text"], amount_value):
                wanted.append((word["x0"] - 1, height - word["bottom"] - 1, word["x1"] - word["x0"] + 2, word["bottom"] - word["top"] + 2))
    overlay_path = output.with_suffix(".overlay.pdf")
    c = canvas.Canvas(str(overlay_path), pagesize=(width, height))
    c.setFillColor(HIGHLIGHT)
    for x, y, w, h in wanted:
        c.rect(x, y, w, h, stroke=0, fill=1)
    if sr is not None and str(sr).strip():
        label = str(sr).strip()
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", 16)
        tw = c.stringWidth(label, "Helvetica-Bold", 16)
        x, y = width - 24 - tw, height - 28
        c.drawString(x, y, label)
        c.setLineWidth(1.2)
        c.line(x, y - 2, x + tw, y - 2)
    c.save()
    overlay = PdfReader(str(overlay_path)).pages[0]
    original.merge_page(overlay)
    writer = PdfWriter()
    writer.add_page(original)
    with open(output, "wb") as f:
        writer.write(f)
    overlay_path.unlink(missing_ok=True)


def process_workbook(source_dir: Path, workbook_path: Path, output_dir: Path, log) -> Path:
    source_files = sorted(
        [p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() in {".txt", ".pdf"}],
        key=lambda p: p.name.lower(),
    )
    if not source_files:
        raise ValueError("The selected folder does not contain any .txt or .pdf files.")
    pages = []
    for source_path in source_files:
        pages.extend(pdf_pages(source_path) if source_path.suffix.lower() == ".pdf" else split_text_pages(source_path))
    wb = load_workbook(workbook_path)
    ws = wb.active
    output_dir.mkdir(parents=True, exist_ok=True)
    for row in range(2, ws.max_row + 1):
        account = ws.cell(row, 1).value
        if account is None or str(account).strip() == "":
            continue
        amount = ws.cell(row, 2).value
        sr = ws.cell(row, 3).value
        result = find_match(pages, account, amount)
        cell = ws.cell(row, 4)
        label = dashed_account(account) or str(account)
        if result.page is None:
            cell.value = f"Not found: {label}"
            log(f"Row {row}: not found {label}")
            continue
        requested_name = str(sr).strip() if sr not in (None, "") else compact_digits(account)
        safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", requested_name)
        out_pdf = output_dir / f"{safe_name}.pdf"
        if out_pdf.exists():
            duplicate = 2
            while (output_dir / f"{safe_name}_{duplicate}.pdf").exists():
                duplicate += 1
            out_pdf = output_dir / f"{safe_name}_{duplicate}.pdf"
        if result.page.source_path.suffix.lower() == ".pdf":
            render_pdf_page(result.page.source_path, result.page.number, result, account, sr, out_pdf)
        else:
            render_text_page(result.page, result, sr, out_pdf)
        if numeric_value(amount) is None:
            cell.value = f"Found and PDF made (page {result.page.number})"
        elif result.amount_found and result.amount_exact:
            cell.value = f"Found AC and amount; PDF made (page {result.page.number})"
        elif result.amount_found:
            cell.value = f"Found AC and near amount (within +/- {SMART_AMOUNT_TOLERANCE:.2f}); PDF made (page {result.page.number})"
        else:
            cell.value = f"AC found, amount not found; PDF made (page {result.page.number})"
        log(f"Row {row}: {cell.value}")
    # Update the workbook selected by the user; only PDFs are written to the output folder.
    wb.save(workbook_path)
    return workbook_path


class App:
    def __init__(self, root: Tk):
        self.root = root
        root.title("Posting Checker")
        root.geometry("680x500")
        root.minsize(680, 420)
        self.source_dir = StringVar()
        self.workbook = StringVar()
        self.output = StringVar(value=str(Path.cwd() / "output"))
        self.status = StringVar(value="Ready")
        self._build()

    def _build(self):
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill=BOTH, expand=True)
        ttk.Label(frame, text="Posting Checker", font=("Segoe UI", 20, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Developed by Tahir Naqash", wraplength=700).pack(anchor="w", pady=(4, 18))
        self._file_row(frame, "Posting folder (.txt/.pdf)", self.source_dir, self.pick_source_dir, None)
        self._file_row(frame, "Excel input (.xlsx)", self.workbook, self.pick_workbook, ("Excel files", "*.xlsx"))
        self._file_row(frame, "Output folder", self.output, self.pick_output, None)
        buttons = ttk.Frame(frame); buttons.pack(fill=X, pady=(12, 8))
        self.run_button = ttk.Button(buttons, text="Check posting", command=self.start)
        self.run_button.pack(side=LEFT)
        ttk.Button(buttons, text="Open output folder", command=self.open_output).pack(side=LEFT, padx=8)
        ttk.Label(frame, textvariable=self.status).pack(anchor="w", pady=(0, 5))
        self.log = ttk.Treeview(frame, columns=("message",), show="headings", height=14)
        self.log.heading("message", text="Activity")
        self.log.column("message", width=690)
        self.log.pack(fill=BOTH, expand=True)

    def _file_row(self, parent, label, var, command, filetypes):
        row = ttk.Frame(parent); row.pack(fill=X, pady=5)
        ttk.Label(row, text=label, width=25).pack(side=LEFT)
        ttk.Entry(row, textvariable=var).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(row, text="Browse", command=lambda: command(filetypes)).pack(side=LEFT, padx=(8, 0))

    def pick_source_dir(self, _):
        value = filedialog.askdirectory()
        if value: self.source_dir.set(value)

    def pick_workbook(self, filetypes):
        value = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if value: self.workbook.set(value)

    def pick_output(self, _):
        value = filedialog.askdirectory()
        if value: self.output.set(value)

    def open_output(self):
        path = Path(self.output.get())
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)

    def start(self):
        if not self.source_dir.get() or not self.workbook.get():
            messagebox.showerror("Missing input", "Select both a posting folder and an Excel workbook.")
            return
        source_dir, workbook, output = Path(self.source_dir.get()), Path(self.workbook.get()), Path(self.output.get())
        if not source_dir.exists() or not source_dir.is_dir() or not workbook.exists():
            messagebox.showerror("Invalid input", "One or more selected files do not exist.")
            return
        self.run_button.configure(state="disabled")
        self.status.set("Processing...")
        threading.Thread(target=self._worker, args=(source_dir, workbook, output), daemon=True).start()

    def _worker(self, source_dir, workbook, output):
        try:
            result = process_workbook(source_dir, workbook, output, lambda msg: self.root.after(0, lambda: self._log(msg)))
            self.root.after(0, lambda: self._done(result))
        except Exception as exc:
            self.root.after(0, lambda: self._error(exc))

    def _log(self, message):
        self.log.insert("", END, values=(message,))
        self.log.yview_moveto(1)

    def _done(self, result):
        self.run_button.configure(state="normal")
        self.status.set(f"Completed: {result}")
        messagebox.showinfo("Completed", f"Excel updated:\n{result}\n\nPDFs saved in:\n{result.parent}")

    def _error(self, exc):
        self.run_button.configure(state="normal")
        self.status.set("Error")
        messagebox.showerror("Processing error", str(exc))


if __name__ == "__main__":
    root = Tk()
    App(root)
    root.mainloop()

