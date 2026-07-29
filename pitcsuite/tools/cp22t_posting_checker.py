import re
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pdfplumber
from openpyxl import Workbook, load_workbook


OUTPUT_FIELDS = [
    "TRANS_REF_NO", "METER_SERIAL_NO", "TRANS_ID", "LOCATION_ID", "METER_ID",
    "METER_SEAL_NO", "METER_INSTALL_DATE", "STATUS", "AMR_SIM_ID", "AMR_SIM_NO",
    "METER_MFG_ID", "METER_MF", "SOURCE_FILE", "MATCH_STATUS", "MATCH_METHOD"
]

DATE_RE = re.compile(r"\d{1,2}-[A-Za-z]{3}-\d{4}")
TID_RE = re.compile(r"^T\d+$", re.I)
LREF_RE = re.compile(r"^L\d+$", re.I)
LONG_NUM_RE = re.compile(r"^\d{14,20}$")
TEN_NUM_RE = re.compile(r"^\d{8,12}$")


def norm(value):
    if value is None:
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(value)).upper()


def words_for_record(words, start, end):
    return [w for w in words[start:end] if w["text"].strip()]


def values_in_record(record):
    vals = [w["text"].strip() for w in record]
    upper = [v.upper() for v in vals]
    base_top = next((w["top"] for w in record if TID_RE.match(w["text"].strip())), None)
    tids = [v for v in vals if TID_RE.match(v)]
    lrefs = [v for v in vals if LREF_RE.match(v)]
    dates = [v for v in vals if DATE_RE.fullmatch(v)]
    long_nums = [v for v in vals if LONG_NUM_RE.fullmatch(v)]
    ten_nums = [v for v in vals if TEN_NUM_RE.fullmatch(v)]
    sim_ids = [v for v in vals if re.fullmatch(r"\d{16,20}", v)]
    sim_nos = [v for v in vals if re.fullmatch(r"\d{9,12}", v)]

    # The two report layouts place fields differently. These type-based rules
    # preserve the identifiers even when PDF text extraction changes order.
    status = "NOT POSTED" if "NOT POSTED" in " ".join(upper) else ("POSTED" if "POSTED" in upper else "")
    trans_ref = long_nums[0] if long_nums else (lrefs[0] if lrefs else "")
    msn = ""
    serial_candidates = [v for v in vals if re.fullmatch(r"\d{10}", v) and not v.startswith("0")]
    if serial_candidates:
        msn = serial_candidates[-1]
    elif long_nums:
        msn = long_nums[-1]

    location = lrefs[0] if lrefs else ""
    for v in ten_nums:
        if v.startswith("26"):
            location = location or v
            break
    meter_id = next((v for v in ten_nums if v.startswith("26")), "")
    meter_seal = next((v for v in serial_candidates if v == msn), "")
    meter_mfg = next((v for v in vals if v in {"36", "48"}), "")

    # For CP22T addition records, the first 10-digit value after the L ref is
    # the meter/location identifier; for change records this is also the
    # location identifier. Keep the visible manufacturer code when present.
    # METER_MF is printed in different positions in the two CP22T layouts:
    # CH: rightmost column on the transaction's first data line.
    # Addition: middle column on the transaction's second data line.
    meter_mf = ""
    if base_top is not None:
        mf_words = []
        for w in record:
            text_value = w["text"].strip()
            if not re.fullmatch(r"\d+(?:\.\d+)?", text_value):
                continue
            x = w["x0"]
            relative_top = w["top"] - base_top
            if 720 <= x <= 790 and abs(relative_top) <= 4:
                mf_words.append(w)
            elif 255 <= x <= 315 and abs(relative_top - 18) <= 4:
                mf_words.append(w)
        if mf_words:
            meter_mf = mf_words[0]["text"].strip()
    install_date = dates[0] if dates else ""
    sim_id = sim_ids[-1] if sim_ids else ""
    sim_no = ""
    for v in reversed(vals):
        if re.fullmatch(r"0\d{9,10}", v):
            sim_no = v
            break

    return {
        "TRANS_REF_NO": trans_ref,
        "METER_SERIAL_NO": msn,
        "TRANS_ID": tids[0] if tids else "",
        "LOCATION_ID": location,
        "METER_ID": meter_id,
        "METER_SEAL_NO": meter_seal,
        "METER_INSTALL_DATE": install_date,
        "STATUS": status,
        "AMR_SIM_ID": sim_id,
        "AMR_SIM_NO": sim_no,
        "METER_MFG_ID": meter_mfg,
        "METER_MF": meter_mf,
        "_all_values": {norm(v) for v in vals},
    }


def extract_pdf(pdf_path):
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            words = page.extract_words(x_tolerance=1, y_tolerance=2, keep_blank_chars=False)
            starts = [i for i, w in enumerate(words) if TID_RE.match(w["text"].strip())]
            for pos, start in enumerate(starts):
                end = starts[pos + 1] if pos + 1 < len(starts) else len(words)
                item = values_in_record(words_for_record(words, start, end))
                if item["TRANS_ID"]:
                    item["SOURCE_FILE"] = Path(pdf_path).name
                    item["_page"] = page_no
                    records.append(item)
    return records


def run_check(xlsx_path, pdf_paths, log):
    all_records = []
    for pdf in pdf_paths:
        log(f"Reading {Path(pdf).name}...")
        all_records.extend(extract_pdf(pdf))
    if not all_records:
        raise ValueError("No CP22T transaction records were found in the selected PDFs.")

    by_ac = {}
    by_msn = {}
    for r in all_records:
        ac = norm(r["TRANS_REF_NO"])
        msn = norm(r["METER_SERIAL_NO"])
        if ac:
            by_ac.setdefault(ac, r)
        if msn:
            by_msn.setdefault(msn, r)

    if xlsx_path:
        wb = load_workbook(xlsx_path)
        ws = wb.active
        output = Path(xlsx_path).with_name(Path(xlsx_path).stem + "_checked.xlsx")
        input_start_col = 3
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "CP22T PDF Records"
        output = Path(pdf_paths[0]).with_name("CP22T_PDF_Records.xlsx")
        input_start_col = 1

    for col, field in enumerate(OUTPUT_FIELDS, input_start_col):
        ws.cell(1, col).value = field
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{ws.cell(1, input_start_col - 1 + len(OUTPUT_FIELDS)).coordinate}"

    counts = {"POSTED": 0, "NOT POSTED": 0, "NOT FOUND": 0}
    if not xlsx_path:
        for row, found in enumerate(all_records, 2):
            for col, field in enumerate(OUTPUT_FIELDS, input_start_col):
                if field == "MATCH_STATUS":
                    value = found["STATUS"] or "FOUND"
                elif field == "MATCH_METHOD":
                    value = "PDF RECORD"
                else:
                    value = found.get(field, "")
                ws.cell(row, col).value = value
            counts[found["STATUS"] if found["STATUS"] in counts else "NOT FOUND"] += 1
    else:
        for row in range(2, ws.max_row + 1):
            ac = norm(ws.cell(row, 1).value)
            msn = norm(ws.cell(row, 2).value)
            # Priority is intentional: AC number first, MSN only as fallback.
            found = by_ac.get(ac) if ac else None
            match_method = "FOUND BY TRANS_REF_NO" if found else ""
            if not found and msn:
                found = by_msn.get(msn)
                if found:
                    match_method = "FOUND BY METER_SERIAL_NO"
            if found:
                for col, field in enumerate(OUTPUT_FIELDS, input_start_col):
                    if field == "MATCH_STATUS":
                        value = found["STATUS"] or "FOUND"
                    elif field == "MATCH_METHOD":
                        value = match_method
                    else:
                        value = found.get(field, "")
                    ws.cell(row, col).value = value
                counts[found["STATUS"] if found["STATUS"] in counts else "NOT FOUND"] += 1
            else:
                for col in range(input_start_col, input_start_col + len(OUTPUT_FIELDS)):
                    ws.cell(row, col).value = "NOT FOUND" if col == input_start_col + OUTPUT_FIELDS.index("MATCH_STATUS") else ""
                counts["NOT FOUND"] += 1

    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)
    for col in range(1, (3 if xlsx_path else 1) + len(OUTPUT_FIELDS)):
        letter = ws.cell(1, col).column_letter
        if ws.column_dimensions[letter].width is None or ws.column_dimensions[letter].width < 14:
            ws.column_dimensions[letter].width = 18
    wb.save(output)
    log(f"Done: {sum(counts.values())} input rows; POSTED={counts['POSTED']}, NOT POSTED={counts['NOT POSTED']}, NOT FOUND={counts['NOT FOUND']}")
    return output


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CP22T Posting Checker")
        self.geometry("720x430")
        self.xlsx = tk.StringVar()
        self.pdfs = []
        self.status = tk.StringVar(value="Excel is optional. Select PDFs; add Excel only for matching existing rows.")
        self.build()

    def build(self):
        pad = {"padx": 12, "pady": 8}
        ttk.Label(self, text="CP22T Posting Checker", font=("Segoe UI", 16, "bold")).pack(anchor="w", **pad)
        frame = ttk.Frame(self); frame.pack(fill="x", **pad)
        ttk.Label(frame, text="Excel input (optional; AC No. in column A, MSN in column B):").pack(anchor="w")
        row = ttk.Frame(frame); row.pack(fill="x", pady=4)
        ttk.Entry(row, textvariable=self.xlsx).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Browse Excel", command=self.pick_excel).pack(side="left", padx=6)
        ttk.Label(frame, text="Selected CP22T PDFs:").pack(anchor="w", pady=(12, 0))
        self.listbox = tk.Listbox(frame, height=8)
        self.listbox.pack(fill="x", pady=4)
        buttons = ttk.Frame(frame); buttons.pack(fill="x")
        ttk.Button(buttons, text="Add PDFs", command=self.pick_pdfs).pack(side="left")
        ttk.Button(buttons, text="Clear PDFs", command=self.clear_pdfs).pack(side="left", padx=6)
        ttk.Button(buttons, text="Check & Save", command=self.start).pack(side="right")
        ttk.Label(self, textvariable=self.status, foreground="#234").pack(anchor="w", **pad)

    def pick_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if p: self.xlsx.set(p)

    def pick_pdfs(self):
        ps = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        for p in ps:
            if p not in self.pdfs:
                self.pdfs.append(p); self.listbox.insert("end", p)

    def clear_pdfs(self):
        self.pdfs.clear(); self.listbox.delete(0, "end")

    def start(self):
        if not self.pdfs:
            messagebox.showwarning("Missing PDFs", "Please select at least one CP22T PDF.")
            return
        self.status.set("Working...")
        threading.Thread(target=self.worker, daemon=True).start()

    def worker(self):
        try:
            output = run_check(self.xlsx.get(), self.pdfs, lambda msg: self.after(0, self.status.set, msg))
            self.after(0, lambda: messagebox.showinfo("Completed", f"Output saved to:\n{output}"))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))


if __name__ == "__main__":
    App().mainloop()

