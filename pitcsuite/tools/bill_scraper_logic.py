"""Desktop PITC Bill scraper.

The selected .xlsx workbook is updated in place using its existing headings.
Excluded fields are removed before requests are made, and returned values are
written only below their matching headings.
"""

from __future__ import annotations

import queue
import re
import threading
import time
import tkinter as tk
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook


URL = "https://bill.pitc.com.pk/pescobill"


# Fields that are intentionally never requested/exported.  These are applied
# before scraping as well as before writing, so skipped fields cannot reappear
# as dynamically-created columns.
DROP_LETTERS = {"B", "C", "E", "H", "K", "L", "M", "N", "P", "Q", "S"}
DROP_RANGES = ((24, 50), (28, 55), (36, 64), (67, 67), (98, 214), (217, 244))  # X:AX, AB:BC, AJ:BL, BO, CT:HF and HI:IJ
DROP_EXACT = {
    "reference_no", "consumer_id", "sub_division", "tariff_cat", "transformer",
    "energy_charges", "taxes", "current_bill", "arrears", "installment",
    "tracking_id",
    "total_electricity_charges", "net_electricity_charges", "total_fpa", "subsidies", "category", "26",
    "kwh_units_40", "kwh_units_41", "kwh_units_44", "kwh_units_45", "kwh_units_46",
    "table_1_kwh_type_1",
    "table_1_kvarh_meter_1", "table_1_kvarh_meter_2", "table_1_kvarh_meter_3", "table_1_kvarh_meter_4",
    "table_1_kvarh_type_1", "table_1_mdi_meter_1", "table_1_mdi_meter_2", "table_1_mdi_meter_3", "table_1_mdi_meter_4",
    "table_1_mdi_type_1", "table_1_mdi_previous_1",
    "table_3_787_1", "table_3_787_off_peak_1", "table_3_787_peak_1",
    "table_3_2402_1", "table_3_2402_off_peak_1", "table_3_2402_peak_1",
    "table_3_1615_1", "table_3_1615_off_peak_1", "table_3_1615_peak_1",
}


def excel_col(n):
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def normalized_key(value):
    key = field_key(value)
    renamed = {
        "kwh_units_30": "kvarh_units",
        "kwh_units_31": "mdi_read_1", "kwh_units_32": "mdi_read_2",
        "kwh_units_33": "mdi_read_3", "kwh_units_34": "mdi_read_4",
        "kwh_units_35": "mdi_mf_1", "kwh_units_36": "mdi_mf_2",
        "kwh_units_37": "mdi_mf_3", "kwh_units_38": "mdi_mf_4",
        "kwh_units_39": "mdi_units",
        "kwh_units_42": "cum_mdi_1", "kwh_units_43": "cum_mdi_2",
        "mdi_present_1": "mdi_read_1", "mdi_present_2": "mdi_read_2",
        "mdi_present_3": "mdi_read_3", "mdi_present_4": "mdi_read_4",
    }
    if key in renamed:
        return renamed[key]
    if key == "table_3_787_mdi_1":
        return "export_mdi"
    if key == "table_3_2402_mdi_1":
        return "import_mdi"
    key = re.sub(r"^table_\d+_", "", key)
    return re.sub(r"^\d+(?:_\d+)*_", "", key)


def is_preserved_meter_key(key):
    return (
        key.startswith("kvarh_present_")
        or key == "kvarh_units"
        or key.startswith("kvarh_units_")
        or key.startswith("mdi_read_")
        or key.startswith("mdi_mf_")
        or key == "mdi_units"
        or key.startswith("mdi_units_")
        or key.startswith("cum_mdi_")
    )


def is_dropped(name, source_col=None, legacy_columns=False):
    key = field_key(name)
    if source_col is not None:
        normalized = normalized_key(name)
        in_dropped_range = any(start <= source_col <= end for start, end in DROP_RANGES)
        if (legacy_columns and excel_col(source_col) in DROP_LETTERS) or (in_dropped_range and not is_preserved_meter_key(normalized)):
            return True
    match = re.fullmatch(r"kwh_units_(\d+)", key)
    if match and int(match.group(1)) > 4:
        return True
    return (
        key in DROP_EXACT
        or "previous" in key
        or "bill_history" in key
        or key.startswith(("kvarh_meter_", "kvarh_mf_", "mdi_meter_"))
        or key.startswith("bill_history")
        or key.startswith(("table_4_", "table_5_", "table_6_"))
        or key in {"account_no", "raw_bill_text"}
    )


def clean(value):
    return re.sub(r"\s+", " ", value or "").strip()


def field_key(value):
    return re.sub(r"[^a-zA-Z0-9]+", "_", clean(value).lower()).strip("_") or "field"


GENERAL_LABELS = {
    "san load": "san_load",
    "san load kw": "san_load",
    "tariff": "tariff",
    "feeder": "feeder",
    "name address": "name_address",
    "name and address": "name_address",
    "energy charges": "energy_charges",
    "taxes": "taxes",
    "current bill": "current_bill",
    "arrears": "arrears",
    "installment": "installment",
    "adjustment": "adjustments",
    "adjustments": "adjustments",
    "grand total": "grand_total",
    "ntn no": "ntn_no",
    "bill month": "bill_month",
    "due date": "due_date",
    "transformer": "transformer",
}


def canonical_label(value):
    label = re.sub(r"[^a-z0-9]+", " ", clean(value).lower()).strip()
    return GENERAL_LABELS.get(label, normalized_key(value))


def order_output_headers(headers):
    """Keep tariff/san-load stable and keep meter readings under fixed names."""
    first = [
        "AC No.", "tariff", "san_load", "adjustments", "grand_total",
        "ntn_no", "bill_month", "name_address", "due_date",
    ]
    ordered = [name for name in first if name in headers]
    meter_order = [
        *(f"kwh_meter_{i}" for i in range(1, 5)),
        *(f"kwh_present_{i}" for i in range(1, 5)),
        *(f"kwh_mf_{i}" for i in range(1, 5)),
        *(f"kwh_units_{i}" for i in range(1, 5)),
        *(f"kvarh_present_{i}" for i in range(1, 5)),
        *(f"kvarh_units_{i}" for i in range(1, 5)),
        *(f"mdi_read_{i}" for i in range(1, 5)),
        *(f"mdi_mf_{i}" for i in range(1, 5)),
        *(f"mdi_units_{i}" for i in range(1, 5)),
        *(f"cum_mdi_{i}" for i in range(1, 5)),
    ]
    ordered += [name for name in meter_order if name in headers and name not in ordered]
    ordered += [name for name in headers if name not in ordered and name not in {"feeder", "scrape_status"}]
    if "feeder" in headers:
        ordered.append("feeder")
    if "scrape_status" in headers:
        ordered.append("scrape_status")
    return ordered


def text_lines(element):
    for br in element.find_all("br"):
        br.replace_with("\n")
    return [clean(x) for x in element.get_text("\n").splitlines() if clean(x)]


class PITCBillScraper:
    def __init__(self, delay=0.5, timeout=90):
        self.delay = delay
        self.timeout = timeout

    @staticmethod
    def put(output, name, value):
        value = clean(value)
        if not value or is_dropped(name):
            return
        key = canonical_label(name)
        # The general bill sometimes exposes the Urdu label "ایڈجسٹمنٹ"
        # where the amount should be. Never export that label by itself;
        # retain an adjustment only when the cell contains a numeric amount.
        if key == "adjustments":
            numeric_adjustment = re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", value)
            if not numeric_adjustment:
                return
            value = numeric_adjustment.group(0)
        # Meter readings sometimes arrive as "I 123" / "E 123".  Keep the
        # numeric reading itself, regardless of whether the source is MDI,
        # KWH, KVARH, or a general-bill table.
        if re.search(r"(?:meter|present|mdi|kwh|kvarh|reading)", key, re.I):
            value = re.sub(r"^\s*[IE]\s+", "", value, flags=re.I).strip()
        if "units" in key:
            value = re.sub(r"\(\s*[OP]\s*\)\s*", "", value, flags=re.I).strip()
        if value:
            existing = output.get(key)
            # Some general-bill layouts expose a placeholder 0 before the
            # actual PRESENT READING. Prefer the later nonzero reading.
            if existing is None or (str(existing).strip() in {"0", "0.0", "0.00"} and str(value).strip() not in {"0", "0.0", "0.00"}):
                output[key] = value

    @staticmethod
    def put_general_pairs(output, soup):
        """Read label/value pairs from both MDI and ordinary bill layouts."""
        known = set(GENERAL_LABELS)
        for row in soup.select("tr"):
            cells = [clean(x.get_text(" ")) for x in row.select("th, td") if clean(x.get_text(" "))]
            if len(cells) >= 2:
                label = cells[0].lower().rstrip(":")
                if label in known:
                    PITCBillScraper.put(output, cells[0], cells[-1])
        for node in soup.select(".grid-col-cell, .charges-bd-row, .right-section-cell, .ibn-bill-messages__cell"):
            bits = [clean(x) for x in node.stripped_strings if clean(x)]
            if len(bits) >= 2 and bits[0].lower().rstrip(":") in known:
                PITCBillScraper.put(output, bits[0], bits[-1])

    @staticmethod
    def put_general_text_pairs(output, soup):
        """Fallback for general batches whose labels are not table cells."""
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        if not main:
            return
        labels = set(GENERAL_LABELS)
        lines = [clean(x) for x in main.stripped_strings if clean(x)]
        for index, line in enumerate(lines):
            label = re.sub(r"[^a-z0-9]+", " ", line.lower()).strip()
            if label not in labels:
                continue
            key = GENERAL_LABELS[label]
            if is_dropped(key):
                continue
            if key in output:
                existing = str(output[key])
                if key == "san_load" and re.search(r"[A-Za-z]-?\d", existing):
                    del output[key]
                elif key == "tariff" and not re.search(r"[A-Za-z]-?\d", existing):
                    del output[key]
                else:
                    continue
            for candidate in lines[index + 1:index + 4]:
                candidate_key = re.sub(r"[^a-z0-9]+", " ", candidate.lower()).strip()
                if candidate_key in labels or not candidate:
                    continue
                if any(ord(ch) > 127 for ch in candidate) and not re.search(r"\d", candidate):
                    continue
                if key == "san_load" and not re.search(r"\d", candidate):
                    continue
                if key == "tariff" and not re.search(r"[A-Za-z]-?\d", candidate):
                    continue
                PITCBillScraper.put(output, key, candidate)
                break

    @staticmethod
    def put_meter_tables(output, soup):
        """Map both MDI and ordinary meter tables to stable reading headings."""
        header_aliases = {
            "meter": "meter", "meter no": "meter", "meter no.": "meter", "meter number": "meter",
            "present": "present", "present reading": "present", "present readings": "present",
            "mf": "mf", "units": "units", "unit": "units", "type": None,
            "previous": None, "previous reading": None, "previous readings": None,
        }
        for table in soup.select("table"):
            headers = [clean(x.get_text(" ")).lower() for x in table.select("thead th")]
            if not headers:
                headers = [clean(x.get_text(" ")).lower() for x in table.select("tr:first-child th, tr:first-child td")]
            mapped = [header_aliases.get(re.sub(r"\s+", " ", h).strip()) for h in headers]
            if not any(x in {"meter", "present", "mf", "units"} for x in mapped):
                continue
            rows = table.select("tbody tr") or table.select("tr")[1:]
            carry = None
            for row in rows:
                cells = row.select("th, td")
                if not cells:
                    continue
                if headers and len(cells) == len(headers) - 1 and carry is not None:
                    cells = [carry] + cells
                if cells and headers and headers[0] in {"meter", "meter no", "meter no."}:
                    carry = cells[0]
                row_text = " ".join(clean(x.get_text(" ")) for x in cells[:2]).lower()
                prefix = "kvarh" if "kvarh" in row_text or "kvar h" in row_text else "mdi" if re.search(r"\bmdi\b", row_text) else "kwh"
                for col_no, cell in enumerate(cells):
                    kind = mapped[col_no] if col_no < len(mapped) else None
                    if kind is None:
                        continue
                    for value_no, value in enumerate(text_lines(cell), 1):
                        self_value = re.sub(r"\b[IE]\s+(?=\d)", "", value, flags=re.I).strip()
                        PITCBillScraper.put(output, f"{prefix}_{kind}_{value_no}", self_value)

    @staticmethod
    def put_meter_text_sections(output, soup):
        """Fallback for general bills that render meter data as div/text blocks."""
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        lines = [clean(x) for x in main.stripped_strings if clean(x)]
        lowered = [re.sub(r"[^a-z0-9]+", " ", x.lower()).strip() for x in lines]
        try:
            start = next(i for i, value in enumerate(lowered) if value in {"meter info", "meter information"})
        except StopIteration:
            return
        end = next((i for i in range(start + 1, len(lines)) if lowered[i] in {"bill charges breakdown", "bill history"}), len(lines))
        section_lines = lines[start:end]
        section_keys = [re.sub(r"[^a-z0-9]+", " ", x.lower()).strip() for x in section_lines]
        # MDI bills use a real METER/TYPE/PREVIOUS/PRESENT/MF/UNITS table.
        # Only ordinary/general bills contain the text label METER NO.
        if not any(value in {"meter no", "meter number"} for value in section_keys):
            return
        labels = {
            "meter no": "meter", "meter number": "meter", "meter": "meter",
            "present reading": "present", "present readings": "present",
            "present": "present", "mf": "mf", "units": "units", "unit": "units",
        }
        for index, label in enumerate(section_keys):
            kind = labels.get(label)
            if not kind:
                continue
            values = []
            for candidate in section_lines[index + 1:]:
                candidate_key = re.sub(r"[^a-z0-9]+", " ", candidate.lower()).strip()
                if candidate_key in labels or candidate_key in {"previous", "previous reading", "type", "meter info"}:
                    break
                if any(ord(ch) > 127 for ch in candidate) and not re.search(r"\d", candidate):
                    continue
                if kind == "meter" and not re.search(r"\d", candidate):
                    continue
                if kind != "meter" and not re.search(r"[-+]?\d", candidate):
                    continue
                values.append(re.sub(r"\b[IE]\s+(?=\d)", "", candidate, flags=re.I).strip())
            for value_no, value in enumerate(values, 1):
                PITCBillScraper.put(output, f"kwh_{kind}_{value_no}", value)

    @staticmethod
    def put_cum_mdi(output, soup):
        """Read present cumulative MDI values without importing previous values."""
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        lines = [clean(x) for x in main.stripped_strings if clean(x)]
        keys = [re.sub(r"[^a-z0-9]+", " ", x.lower()).strip() for x in lines]
        start = next((i for i, value in enumerate(keys) if value in {"prs cumm mdi", "present cumm mdi", "present cumulative mdi"}), None)
        if start is None:
            return
        values = []
        for candidate in lines[start + 1:start + 7]:
            if re.fullmatch(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", candidate):
                values.append(candidate)
            elif values:
                break
        for value_no, value in enumerate(values[:4], 1):
            PITCBillScraper.put(output, f"cum_mdi_{value_no}", value)

    def search(self, account):
        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; PITC-Bill-Exporter/1.0)"})
        page = session.get(URL, timeout=self.timeout)
        page.raise_for_status()
        soup = BeautifulSoup(page.text, "lxml")
        form = soup.select_one("form#SubmitForm") or soup.find("form")
        if not form:
            raise RuntimeError("PITC search form was not found")
        payload = {i["name"]: i.get("value", "") for i in form.select("input[type=hidden][name]")}
        payload.update({"rbSearchByList": "refno", "searchTextBox": account, "ruCodeTextBox": "", "btnSearch": "Search"})
        result = session.post(URL, data=payload, timeout=self.timeout)
        result.raise_for_status()
        if "CONSUMER DETAIL" not in result.text.upper() and "CONSUMER DETAILS" not in result.text.upper():
            raise RuntimeError(clean(BeautifulSoup(result.text, "lxml").get_text(" "))[:300] or "No bill returned")
        return BeautifulSoup(result.text, "lxml")

    def scrape_account(self, raw_account):
        raw = str(raw_account).strip()
        if re.fullmatch(r"[0-9.]+e[+\-]?[0-9]+", raw, re.I):
            raw = format(Decimal(raw), "f")
        account = re.sub(r"\D", "", raw)
        if not re.fullmatch(r"\d{10,14}", account):
            return {"_account": account, "scrape_status": "invalid account number"}
        try:
            soup = self.search(account)
            out = {"_account": account, "scrape_status": "ok"}
            # General-bill labels are collected first so later table variants
            # cannot overwrite tariff/name/san-load with shifted values.
            self.put_general_pairs(out, soup)
            self.put_general_text_pairs(out, soup)
            for cell in soup.select(".grid-col-cell"):
                label, value = cell.select_one(".en-lbl"), cell.select_one(".val-space")
                if label and value:
                    self.put(out, label.get_text(" "), " | ".join(text_lines(value)))

            for node in soup.select(".charges-bd-row, .right-section-cell, .ibn-bill-messages__cell"):
                bits = [clean(x) for x in node.stripped_strings if clean(x)]
                if len(bits) >= 2:
                    self.put(out, bits[0], bits[-1])

            self.put_meter_tables(out, soup)
            self.put_meter_text_sections(out, soup)
            self.put_cum_mdi(out, soup)

            return out
        except Exception as exc:
            return {"_account": account, "scrape_status": f"error: {exc}"}

    @staticmethod
    def prepare_workbook(sheet):
        """Compact the source schema once, before any network request."""
        first = str(sheet.cell(1, 1).value or "").strip().lower().rstrip(".")
        has_heading = first in {"ac", "ac no", "account", "account no"}
        if not has_heading:
            sheet.insert_rows(1)
            sheet.cell(1, 1).value = "AC No."

        old_max_col = sheet.max_column
        old_max_row = sheet.max_row
        layout_with_explicit_ranges = old_max_col >= 44
        legacy_layout = old_max_col >= 200
        old_headers = [sheet.cell(1, col).value for col in range(1, old_max_col + 1)]
        kept = []
        seen = set()
        for col, raw_name in enumerate(old_headers, 1):
            if col == 1:
                name = "AC No."
            else:
                if not raw_name or is_dropped(raw_name, col if layout_with_explicit_ranges else None, legacy_layout):
                    continue
                name = "scrape_status" if normalized_key(raw_name) == "scrape_status" else normalized_key(raw_name)
            if name not in seen:
                kept.append((col, name))
                seen.add(name)

        # Status is always the last output column.
        kept = [item for item in kept if item[1] != "scrape_status"]
        # Feeder is useful on both general and MDI bills. Add its heading if
        # the source workbook did not contain it, without shifting the meter
        # columns already present in an existing workbook.
        if "feeder" not in {name for _, name in kept}:
            kept.append((None, "feeder"))
        # Keep the adjustment column in every exported workbook. Individual
        # accounts may have no adjustment, but their cells should remain blank
        # rather than changing the shared schema between rows.
        if "adjustments" not in {name for _, name in kept}:
            kept.append((None, "adjustments"))
        kept.append((next((col for col, name in enumerate(old_headers, 1) if normalized_key(name) == "scrape_status"), None), "scrape_status"))
        if kept[-1][0] is None:
            kept[-1] = (None, "scrape_status")

        source_by_name = {name: col for col, name in kept}
        kept = [(source_by_name.get(name), name) for name in order_output_headers([name for _, name in kept])]

        source_rows = []
        for row_no in range(2, old_max_row + 1):
            if sheet.cell(row_no, 1).value is None or not str(sheet.cell(row_no, 1).value).strip():
                continue
            source_rows.append([sheet.cell(row_no, old_col).value if old_col else None for old_col, _ in kept])

        # Remove the old wide/dynamic layout completely, then write only the
        # selected schema. This prevents B/C/E and other excluded columns from
        # surviving as blank columns.
        sheet.delete_cols(1, old_max_col)
        for col, (_, name) in enumerate(kept, 1):
            sheet.cell(1, col).value = name
        for row_no, values in enumerate(source_rows, 2):
            for col, value in enumerate(values, 1):
                if value is not None:
                    sheet.cell(row_no, col).value = value
        return [name for _, name in kept], len(source_rows)

    def update_workbook(self, path, progress, stop_event):
        if path.suffix.lower() != ".xlsx":
            raise RuntimeError("Please select an .xlsx workbook. The selected workbook is updated in place.")
        workbook = load_workbook(path)
        sheet = workbook.active
        headers, input_count = self.prepare_workbook(sheet)
        rows = [(row_no, sheet.cell(row_no, 1).value) for row_no in range(2, sheet.max_row + 1) if sheet.cell(row_no, 1).value]
        total = len(rows)
        all_results = []
        for done, (row_no, account) in enumerate(rows, 1):
            if stop_event.is_set():
                break
            result = self.scrape_account(account)
            all_results.append((row_no, result))
            progress(done, total, result)
            time.sleep(self.delay)

        # Only headerless inputs may gain headings from returned bill fields.
        # Existing headings are authoritative: no fields are appended or
        # shifted for a particular bill type.
        for _, result in all_results:
            for name in result:
                if name.startswith("_") or name == "scrape_status" or is_dropped(name):
                    continue
                if name not in headers:
                    headers.insert(len(headers) - 1, name)
        headers = order_output_headers(headers)
        for col, name in enumerate(headers, 1):
            sheet.cell(1, col).value = name

        for row_no, result in all_results:
            for col, name in enumerate(headers, 1):
                value = result.get(name)
                # Missing fields in one bill must not erase existing/general
                # bill cells, and every value stays below its matching header.
                if value not in (None, ""):
                    sheet.cell(row_no, col).value = value
        workbook.save(path)


class BillScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PITC Bill Scrapper")
        self.geometry("480x480")
        self.minsize(480, 480)
        self.events = queue.Queue()
        self.stop_event = threading.Event()
        self.input_var = tk.StringVar()
        self.delay_var = tk.DoubleVar(value=5.0)
        self.status_var = tk.StringVar(value="Choose an Excel file containing AC Nos. in column A.")
        self.build()
        self.after(100, self.drain_events)

    def build(self):
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="PITC Bill Scrapper", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(root, text="Developed by Tahir Naqash").pack(anchor="w", pady=(3, 18))
        files = ttk.LabelFrame(root, text="Workbook", padding=12)
        files.pack(fill="x")
        ttk.Label(files, text="Excel file").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(files, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", padx=8, pady=5)
        ttk.Button(files, text="Browse...", command=self.choose_input).grid(row=0, column=2, pady=5)
        files.columnconfigure(1, weight=1)
        options = ttk.Frame(root)
        options.pack(fill="x", pady=14)
        ttk.Label(options, text="Delay between bills (seconds)").pack(side="left")
        ttk.Spinbox(options, from_=0, to=30, increment=0.1, textvariable=self.delay_var, width=8).pack(side="left", padx=8)
        ttk.Label(options, text="Runs in the background.").pack(side="left")
        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", pady=4)
        ttk.Label(root, textvariable=self.status_var).pack(anchor="w", pady=(0, 8))
        self.log = tk.Text(root, height=7, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True)
        buttons = ttk.Frame(root)
        buttons.pack(fill="x", pady=(12, 0))
        self.start_button = tk.Button(buttons, text="Start Scrapping", command=self.start, bg="#2ecc71", fg="white", activebackground="#27ae60", activeforeground="white", relief="raised", padx=12)
        self.start_button.pack(side="left")
        self.cancel_button = tk.Button(buttons, text="Stop", command=self.cancel, bg="#e74c3c", fg="white", activebackground="#c0392b", activeforeground="white", relief="raised", padx=12, state="disabled")
        self.cancel_button.pack(side="left", padx=8)

    def choose_input(self):
        path = filedialog.askopenfilename(title="Choose account workbook", filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.input_var.set(path)

    def write_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self):
        path = Path(self.input_var.get().strip())
        if not path.is_file():
            messagebox.showerror("Input required", "Choose an existing .xlsx workbook first.")
            return
        if not messagebox.askyesno("Update workbook", "The workbook will be compacted to the selected Bill headings and updated in place. Continue?"):
            return
        try:
            delay = max(0.0, float(self.delay_var.get()))
        except (TypeError, ValueError):
            messagebox.showerror("Invalid delay", "Delay must be a number.")
            return
        self.stop_event.clear()
        self.progress.configure(value=0)
        self.start_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.write_log("Starting in-place workbook update...")
        threading.Thread(target=self.run, args=(path, delay), daemon=True).start()

    def run(self, path, delay):
        try:
            def progress(done, total, row):
                self.events.put(("progress", done, total, row.get("_account", ""), row.get("scrape_status", "")))
            PITCBillScraper(delay=delay).update_workbook(path, progress, self.stop_event)
            self.events.put(("done", str(path), self.stop_event.is_set()))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def cancel(self):
        self.stop_event.set()
        self.status_var.set("Finishing the current bill, then stopping...")
        self.cancel_button.configure(state="disabled")

    def drain_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, done, total, account, status = event
                    self.progress.configure(maximum=max(total, 1), value=done)
                    self.status_var.set(f"Processed {done} of {total}")
                    self.write_log(f"{account}: {status}")
                elif event[0] == "done":
                    _, path, cancelled = event
                    self.start_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.status_var.set("Stopped." if cancelled else "Completed successfully.")
                    self.write_log(("Stopped. " if cancelled else "Completed. ") + path)
                    if not cancelled:
                        messagebox.showinfo("Finished", f"Workbook updated:\n{path}")
                elif event[0] == "error":
                    self.start_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.status_var.set("Failed.")
                    self.write_log("ERROR: " + event[1])
                    messagebox.showerror("Export failed", event[1])
        except queue.Empty:
            pass
        self.after(100, self.drain_events)


if __name__ == "__main__":
    BillScraperApp().mainloop()
