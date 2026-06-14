"""
PITC Suite — Unified Launcher
All tools in one dark-themed PySide6 application.
Developed by Tahir Naqash
"""

import sys
import os
import re
import time
import base64
import threading
import configparser

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QLineEdit, QFileDialog, QCheckBox, QTextEdit, QProgressBar,
    QVBoxLayout, QHBoxLayout, QGroupBox, QMessageBox, QComboBox,
    QSpinBox, QStackedWidget, QFrame, QScrollArea, QSizePolicy
)
from PySide6.QtCore import Qt, Signal, QObject, QThread
from PySide6.QtGui import QFont, QColor, QPalette, QIcon

# ──────────────────────────────────────────────────────────────
#  DARK THEME PALETTE
# ──────────────────────────────────────────────────────────────
DARK_STYLE = """
QMainWindow, QWidget {
    background-color: #1a1d23;
    color: #e0e6f0;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
    fgadfasdf
    adfasf
}
QGroupBox {
    border: 1px solid #2e3340;
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 8px;
    font-weight: bold;
    color: #7eb8f7;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
}
QLineEdit {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    padding: 4px 8px;
    color: #e0e6f0;
}
QLineEdit:focus { border-color: #4a90d9; }
QPushButton {
    background: #252933;
    border: 1px solid #3a3f50;
    border-radius: 4px;
    padding: 5px 14px;
    color: #c8d6f0;
}
QPushButton:hover { background: #2e3340; border-color: #4a90d9; color: #fff; }
QPushButton:pressed { background: #1e2430; }
QPushButton:disabled { color: #555; border-color: #2a2a2a; }
QTextEdit {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    color: #a8c0e0;
    font-family: 'Consolas', monospace;
    font-size: 11px;
}
QProgressBar {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    height: 14px;
    text-align: center;
    color: #e0e6f0;
}
QProgressBar::chunk { background: #4a90d9; border-radius: 3px; }
QComboBox {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    padding: 4px 8px;
    color: #e0e6f0;
}
QComboBox QAbstractItemView {
    background: #1a1d23;
    border: 1px solid #4a90d9;
    selection-background-color: #2e3340;
}
QSpinBox {
    background: #12151c;
    border: 1px solid #2e3340;
    border-radius: 4px;
    padding: 3px 6px;
    color: #e0e6f0;
}
QCheckBox { color: #c8d6f0; }
QCheckBox::indicator { width: 14px; height: 14px; border: 1px solid #3a3f50; background: #12151c; border-radius: 3px; }
QCheckBox::indicator:checked { background: #4a90d9; }
QLabel { color: #c8d6f0; }
QScrollBar:vertical { background: #12151c; width: 8px; border-radius: 4px; }
QScrollBar::handle:vertical { background: #2e3340; border-radius: 4px; }
"""

NAV_STYLE = """
QPushButton {
    background: transparent;
    border: none;
    border-left: 3px solid transparent;
    border-radius: 0;
    padding: 12px 20px;
    text-align: left;
    color: #8090a8;
    font-size: 12px;
    font-family: 'Consolas', monospace;
}
QPushButton:hover { background: #1e2230; color: #c8d6f0; border-left-color: #3a4a6a; }
QPushButton[active="true"] {
    background: #1e2844;
    color: #7eb8f7;
    border-left-color: #4a90d9;
    font-weight: bold;
}
"""

START_BTN = "background:#2a6; border:none; border-radius:4px; padding:6px 18px; color:#fff; font-weight:bold;"
STOP_BTN  = "background:#922; border:none; border-radius:4px; padding:6px 18px; color:#fff; font-weight:bold;"


def lbl(text, bold=False, color=None):
    w = QLabel(text)
    if bold:
        f = w.font(); f.setBold(True); w.setFont(f)
    if color:
        w.setStyleSheet(f"color:{color};")
    return w


def hline():
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setStyleSheet("color:#2e3340;")
    return f


# ──────────────────────────────────────────────────────────────
#  TOOL 1 — PITC Downloader (CDS10)
#  All report types · OTP login · Month+Amount filter · Highlight · Para Sr. Stamp
# ──────────────────────────────────────────────────────────────
REPORT_HREF = {
    "meter":      "MeterReadingAndConsumption",
    "billing":    "BillingDetails",
    "adjustment": "BillingAdjustmentDetails",
    "payment":    "PaymentDetails",
    "summary":    "CustomerDataSummary",
}


class PITCWorker(QThread):
    progress = Signal(int)
    log      = Signal(str)
    finished = Signal()

    def __init__(self, excel, report_key, download_dir, profile_dir,
                 restart_after, warmup, print_wait, print_x, print_y):
        super().__init__()
        self.excel         = excel
        self.report_key    = report_key
        self.download_dir  = download_dir
        self.profile_dir   = profile_dir
        self.restart_after = restart_after
        self.warmup        = warmup
        self.print_wait    = print_wait
        self.print_x       = print_x
        self.print_y       = print_y
        self.running       = True

    def stop(self): self.running = False

    def run(self):
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.service import Service
            from selenium.webdriver.chrome.options import Options
            from selenium.webdriver.common.by import By
            from selenium.webdriver.common.keys import Keys
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
            from webdriver_manager.chrome import ChromeDriverManager
            from openpyxl import load_workbook
            from reportlab.pdfgen import canvas
            from pypdf import PdfReader, PdfWriter
            import pyautogui, fitz, shutil, traceback
            import datetime as _dt
            import pyotp
        except ImportError as e:
            self.log.emit(f"ERROR: Missing library — {e}")
            self.finished.emit()
            return

        cfg = configparser.ConfigParser()
        cfg.read(_config_path())
        try:
            USERNAME   = cfg["credentials"]["username"]
            PASSWORD   = cfg["credentials"]["password"]
            OTP_SECRET = cfg["credentials"].get("otp_secret", "")
        except Exception:
            self.log.emit("ERROR: config.ini missing or invalid. Set credentials in ⚙️ Settings first.")
            self.finished.emit()
            return

        dl = self.download_dir
        os.makedirs(dl, exist_ok=True)
        os.makedirs(self.profile_dir, exist_ok=True)
        pyautogui.FAILSAFE = False
        pyautogui.PAUSE = 0.3

        # ── OTP ──
        def generate_otp():
            totp = pyotp.TOTP(OTP_SECRET)
            while True:
                remaining = totp.interval - (time.time() % totp.interval)
                if remaining < 10:
                    self.log.emit("Waiting for fresh OTP window…")
                    time.sleep(remaining + 1)
                else:
                    break
            return totp.now()

        # ── PDF helpers ──
        def wait_for_new_pdf(existing, timeout=60):
            start = time.time()
            while time.time() - start < timeout:
                pdfs = {f for f in os.listdir(dl)
                        if f.lower().endswith(".pdf") and not f.lower().endswith(".crdownload")}
                diff = pdfs - existing
                if diff:
                    path = os.path.join(dl, diff.pop())
                    if os.path.getsize(path) < 5000:
                        raise Exception("Invalid PDF (login/session issue)")
                    return path
                time.sleep(0.5)
            raise Exception("PDF not downloaded within timeout")

        def get_unique_name(base_name):
            base_name = str(base_name).strip()
            name = f"{base_name}.pdf"
            if not os.path.exists(os.path.join(dl, name)): return name
            i = 1
            while True:
                name = f"{base_name}_{i}.pdf"
                if not os.path.exists(os.path.join(dl, name)): return name
                i += 1

        def format_month(val):
            if isinstance(val, _dt.datetime):
                return val.strftime("%b - %Y")
            s = str(val).strip()
            for fmt in ("%b-%y", "%b-%Y", "%m-%y", "%m-%Y"):
                try:
                    return _dt.datetime.strptime(s, fmt).strftime("%b - %Y")
                except:
                    pass
            return s

        def filter_pdf(pdf_path, month=None, amount=None):
            """Keep only the matching page and highlight month/amount."""
            temp_filtered = pdf_path.replace(".pdf", "_filtered.pdf")
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            month_str = format_month(month).lower() if month else None
            found = False
            for page in reader.pages:
                text  = (page.extract_text() or "").lower()
                clean = text.replace(",", "")
                m = month_str and month_str in text
                a = amount and str(amount).replace(",", "") in clean
                if m or a:
                    writer.add_page(page)
                    found = True
                    break
            if not found:
                os.remove(pdf_path)
                raise Exception("No matching page found (month/amount)")
            with open(temp_filtered, "wb") as f:
                writer.write(f)
            # highlight
            doc = fitz.open(temp_filtered)
            for page in doc:
                if month:
                    for rect in page.search_for(format_month(month)):
                        page.add_highlight_annot(rect).update()
                if amount:
                    areas = page.search_for(str(amount).replace(",", ""))
                    if not areas:
                        areas = page.search_for(str(amount))
                    for rect in areas:
                        page.add_highlight_annot(rect).update()
            highlighted = pdf_path.replace(".pdf", "_highlighted.pdf")
            doc.save(highlighted, garbage=4, deflate=True, clean=True)
            doc.close()
            os.remove(pdf_path)
            os.remove(temp_filtered)
            shutil.move(highlighted, pdf_path)
            self.log.emit("PDF filtered and highlighted")

        def stamp_sr(pdf_path, sr):
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            total  = len(reader.pages)
            for i, page in enumerate(reader.pages, 1):
                w = float(page.mediabox.width); h = float(page.mediabox.height)
                overlay = pdf_path.replace(".pdf", "_overlay.pdf")
                c = canvas.Canvas(overlay, pagesize=(w, h))
                c.setFont("Helvetica-Bold", 14)
                c.drawString(w - 260, h - 30, f"Para Sr. No. {sr} ({i}/{total})")
                c.save()
                page.merge_page(PdfReader(overlay).pages[0])
                writer.add_page(page)
                os.remove(overlay)
            with open(pdf_path, "wb") as f:
                writer.write(f)

        def start_driver():
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--no-sandbox")
            options.add_argument(f"--user-data-dir={self.profile_dir}")
            options.add_experimental_option("prefs", {
                "download.default_directory": dl,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": False,
            })
            drv = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            return drv, WebDriverWait(drv, 40)

        def do_login(drv, wt):
            drv.get("https://cp.pitc.com.pk")
            wt.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-id='ctl00_ContentPlaceHolder1_cmbSelectDisco']"))).click()
            wt.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='PESCO']"))).click()
            wt.until(EC.presence_of_element_located((By.NAME, "ctl00$ContentPlaceHolder1$txtUsername"))).send_keys(USERNAME)
            wt.until(EC.presence_of_element_located((By.NAME, "ctl00$ContentPlaceHolder1$txtPassword"))).send_keys(PASSWORD)
            wt.until(EC.element_to_be_clickable((By.ID, "btnLogin"))).click()
            if OTP_SECRET:
                self.log.emit("Waiting for OTP…")
                for attempt in range(5):
                    try:
                        otp_box = wt.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_txtOtp")))
                        time.sleep(0.5)
                        otp_code = generate_otp()
                        self.log.emit(f"OTP attempt {attempt+1}: {otp_code}")
                        otp_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
                        otp_box.send_keys(otp_code)
                        wt.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnVerify"))).click()
                        time.sleep(2)
                        if drv.find_elements(By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"):
                            self.log.emit("OTP verified ✔")
                            return
                        time.sleep(3)
                    except Exception as e:
                        self.log.emit(f"OTP retry: {e}")
                        time.sleep(3)
                raise Exception("OTP verification failed")

        def ensure_logged_in(drv, wt):
            if not drv.find_elements(By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"):
                self.log.emit("Session expired → re-login")
                do_login(drv, wt)

        def open_report_link(drv, wt, href_kw):
            """Click the report link with stale-element retry."""
            for attempt in range(5):
                try:
                    link = wt.until(EC.element_to_be_clickable(
                        (By.XPATH, f"//a[contains(@href,'{href_kw}')]")
                    ))
                    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
                    time.sleep(0.5)
                    drv.execute_script("arguments[0].click();", link)
                    return
                except StaleElementReferenceException:
                    self.log.emit(f"Stale element retry {attempt+1}")
                    time.sleep(1)
                except TimeoutException:
                    self.log.emit(f"Timeout retry {attempt+1}")
                    time.sleep(1)
            raise Exception(f"Could not open report link: {href_kw}")

        driver, wait = start_driver()
        do_login(driver, wait)

        wb         = load_workbook(self.excel)
        sheet      = wb.active
        total_rows = sheet.max_row - 1
        processed  = 0
        row        = 2
        href_kw    = REPORT_HREF[self.report_key]

        time.sleep(self.warmup)

        while True:
            if not self.running: break
            ac = sheet[f"A{row}"].value
            if not ac: break
            sr     = sheet[f"B{row}"].value
            month  = sheet[f"C{row}"].value
            amount = sheet[f"D{row}"].value
            status = sheet[f"E{row}"]

            self.log.emit(f"Row {row} | AC: {ac}")

            try:
                ensure_logged_in(driver, wait)

                box = wait.until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber")))
                box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
                time.sleep(0.5)
                box.send_keys(str(ac))

                wait.until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$btnGo"))).click()
                time.sleep(2)

                open_report_link(driver, wait, href_kw)

                wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnPrint"))).click()

                existing = {f for f in os.listdir(dl) if f.lower().endswith(".pdf")}
                time.sleep(self.print_wait)

                pyautogui.click(self.print_x, self.print_y)
                pyautogui.press(["down"])
                pyautogui.press(["tab", "tab", "tab", "enter"])

                new_pdf    = wait_for_new_pdf(existing)
                base_name  = sr if sr else ac
                final_name = get_unique_name(base_name)
                final_path = os.path.join(dl, final_name)
                os.rename(new_pdf, final_path)

                if month or amount:
                    filter_pdf(final_path, month, amount)

                if sr:
                    stamp_sr(final_path, sr)

                status.value = "Done"
                self.log.emit(f"✅ PDF created: {final_name}")

                pyautogui.hotkey("ctrl", "w")
                time.sleep(2)

            except Exception as e:
                traceback.print_exc()
                status.value = "Error"
                self.log.emit(f"⚠ ERROR: {e}")
                try: pyautogui.hotkey("ctrl", "w")
                except: pass
                time.sleep(3)

            row += 1
            processed += 1
            self.progress.emit(int((processed / max(total_rows, 1)) * 100))

            if processed % self.restart_after == 0:
                self.log.emit("Restarting browser session…")
                driver.quit(); time.sleep(6)
                driver, wait = start_driver()
                do_login(driver, wait)
                time.sleep(self.warmup)

        wb.save(self.excel)
        driver.quit()
        self.finished.emit()


class PITCPanel(QWidget):
    REPORT_LABELS = [
        "Meter Reading & Consumption",
        "Billing Detail",
        "Billing Adjustment",
        "Payment Detail",
        "Customer Data Summary",
    ]
    REPORT_KEYS = ["meter", "billing", "adjustment", "payment", "summary"]

    def __init__(self):
        super().__init__()
        self.worker = None
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        layout.addWidget(lbl("PITC Downloader", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("OTP login  ·  All report types  ·  Month & Amount filter  ·  Highlight  ·  Para Sr. Stamp", color="#556"))
        layout.addWidget(hline())

        # ── Excel + Template ──
        g1 = QGroupBox("Excel File  (Columns: A = AC No  ·  B = Sr No  ·  C = Month  ·  D = Amount  ·  E = Status)")
        g1v = QVBoxLayout(g1)
        r_ex = QHBoxLayout()
        self.excel_ed = QLineEdit("D:\\PITC.xlsx"); self.excel_ed.setReadOnly(True)
        b1 = QPushButton("Browse"); b1.clicked.connect(self.pick_excel)
        btn_tpl = QPushButton("⬇  Download Template")
        btn_tpl.setStyleSheet("background:#194; border:none; border-radius:4px; padding:5px 10px; color:#fff;")
        btn_tpl.clicked.connect(self.get_template)
        r_ex.addWidget(self.excel_ed); r_ex.addWidget(b1); r_ex.addWidget(btn_tpl)
        g1v.addLayout(r_ex)
        layout.addWidget(g1)

        # ── Report Type ──
        g_rep = QGroupBox("Report Type")
        g_rep_h = QHBoxLayout(g_rep)
        self.report_cb = QComboBox()
        self.report_cb.addItems(self.REPORT_LABELS)
        self.report_cb.setCurrentIndex(4)
        g_rep_h.addWidget(self.report_cb); g_rep_h.addStretch()
        layout.addWidget(g_rep)

        # ── Paths ──
        g2 = QGroupBox("Paths")
        g2v = QVBoxLayout(g2)
        r_dl = QHBoxLayout(); r_dl.addWidget(QLabel("Download Dir:"))
        self.dl_ed = QLineEdit(r"D:\PITC_PDFs"); r_dl.addWidget(self.dl_ed, 1)
        g2v.addLayout(r_dl)
        r_pr = QHBoxLayout(); r_pr.addWidget(QLabel("Chrome Profile:"))
        self.pr_ed = QLineEdit(r"C:\ChromeProfiles\PITC"); r_pr.addWidget(self.pr_ed, 1)
        g2v.addLayout(r_pr)
        layout.addWidget(g2)

        # ── Timing & Print ──
        g3 = QGroupBox("Timing & Print Dialog")
        g3h = QHBoxLayout(g3)

        def spin(val, mn=1, mx=9999):
            s = QSpinBox(); s.setRange(mn, mx); s.setValue(val); return s

        g3h.addWidget(QLabel("Restart After"));  self.sp_restart = spin(35, 1, 200); g3h.addWidget(self.sp_restart)
        g3h.addWidget(QLabel("Warmup (s)"));     self.sp_warmup  = spin(10, 1, 200); g3h.addWidget(self.sp_warmup)
        g3h.addWidget(QLabel("Print Wait (s)")); self.sp_pwait   = spin(6,  1, 200); g3h.addWidget(self.sp_pwait)
        g3h.addWidget(QLabel("Print X"));        self.sp_px      = spin(28);         g3h.addWidget(self.sp_px)
        g3h.addWidget(QLabel("Print Y"));        self.sp_py      = spin(152);        g3h.addWidget(self.sp_py)
        layout.addWidget(g3)

        # ── Buttons ──
        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_stop  = QPushButton("■  STOP");  self.btn_stop.setStyleSheet(STOP_BTN); self.btn_stop.setEnabled(False)
        self.btn_start.clicked.connect(self.run); self.btn_stop.clicked.connect(self.halt)
        bh.addWidget(self.btn_start); bh.addWidget(self.btn_stop); bh.addStretch()
        layout.addLayout(bh)

        self.progress = QProgressBar(); layout.addWidget(self.progress)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box)

    def pick_excel(self):
        f, _ = QFileDialog.getOpenFileName(self, "Excel File", "", "Excel (*.xlsx)")
        if f: self.excel_ed.setText(f)

    def get_template(self):
        download_template(
            self,
            headers      = ["AC No", "Sr No", "Month", "Amount", "Status"],
            default_name = "PITC_Downloader_Template.xlsx",
            sample_rows  = [
                ["07269330425000", 1, "Jan-26", 1500, ""],
                ["07269330426000", 2, "Feb-26", 2300, ""],
            ]
        )

    def run(self):
        self.btn_start.setEnabled(False); self.btn_stop.setEnabled(True)
        self.log_box.clear()
        self.worker = PITCWorker(
            excel         = self.excel_ed.text(),
            report_key    = self.REPORT_KEYS[self.report_cb.currentIndex()],
            download_dir  = self.dl_ed.text(),
            profile_dir   = self.pr_ed.text(),
            restart_after = self.sp_restart.value(),
            warmup        = self.sp_warmup.value(),
            print_wait    = self.sp_pwait.value(),
            print_x       = self.sp_px.value(),
            print_y       = self.sp_py.value(),
        )
        self.worker.log.connect(self.log_box.append)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.done)
        self.worker.start()

    def halt(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop(); self.worker.wait()

    def done(self):
        self.btn_start.setEnabled(True); self.btn_stop.setEnabled(False)
        QMessageBox.information(self, "Done", "PITC Downloader completed.")


# ──────────────────────────────────────────────────────────────
#  TOOL 2 — E/Bills Downloader (BillSoft.py)
# ──────────────────────────────────────────────────────────────
class BillWorkerSignals(QObject):
    log      = Signal(str)
    progress = Signal(int, int)
    finished = Signal()


class BillSoftPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.stop_flag = False; self.pause_flag = False
        self.signals = BillWorkerSignals()
        self.signals.log.connect(lambda m: self.log_box.append(m))
        self.signals.progress.connect(self.update_progress)
        self.signals.finished.connect(self.on_finished)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.addWidget(lbl("E/Bills Downloader", bold=True, color="#7eb8f7"))
        layout.addWidget(hline())

        g = QGroupBox("Inputs")
        iv = QVBoxLayout(g)
        iv.addLayout(self._excel_row())
        iv.addLayout(self._url_row())
        iv.addLayout(self._out_row())
        layout.addWidget(g)

        og = QGroupBox("Options")
        oh = QHBoxLayout(og)
        self.chk_sr       = QCheckBox("Add Sr. No. on Bill")
        self.chk_sep      = QCheckBox("Separate PDFs")
        self.chk_headless = QCheckBox("Headless (background)"); self.chk_headless.setChecked(True)
        oh.addWidget(self.chk_sr); oh.addWidget(self.chk_sep); oh.addWidget(self.chk_headless); oh.addStretch()
        layout.addWidget(og)

        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_pause = QPushButton("⏸  PAUSE"); self.btn_pause.setStyleSheet("background:#a80; border:none; border-radius:4px; padding:6px 14px; color:#fff;")
        self.btn_start.clicked.connect(self.start_stop); self.btn_pause.clicked.connect(self.pause_resume)
        bh.addWidget(self.btn_start); bh.addWidget(self.btn_pause); bh.addStretch()
        layout.addLayout(bh)

        ph = QHBoxLayout()
        self.lbl_prog = QLabel("0 / 0 bills"); ph.addWidget(self.lbl_prog); ph.addStretch()
        layout.addLayout(ph)
        self.bar = QProgressBar(); layout.addWidget(self.bar)

        lg = QGroupBox("Logs")
        lv = QVBoxLayout(lg)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); self.log_box.setFixedHeight(100)
        lv.addWidget(self.log_box)
        layout.addWidget(lg)

    def _excel_row(self):
        r = QHBoxLayout(); r.addWidget(QLabel("Excel File:"))
        self.ed_excel = QLineEdit(); self.ed_excel.setReadOnly(True); r.addWidget(self.ed_excel, 1)
        b = QPushButton("Browse"); b.clicked.connect(lambda: (p := QFileDialog.getOpenFileName(self,"Excel","","Excel (*.xlsx)")[0]) and self.ed_excel.setText(p)); r.addWidget(b)
        return r

    def _url_row(self):
        r = QHBoxLayout(); r.addWidget(QLabel("Bills URL:"))
        self.ed_url = QLineEdit("https://bill.pitc.com.pk/pescobill/"); self.ed_url.setReadOnly(True); r.addWidget(self.ed_url, 1)
        self.btn_url = QPushButton("Edit"); self.btn_url.clicked.connect(self.toggle_url); r.addWidget(self.btn_url)
        return r

    def _out_row(self):
        r = QHBoxLayout(); r.addWidget(QLabel("Output Folder:"))
        self.ed_out = QLineEdit("D:\\BillsPDF"); self.ed_out.setReadOnly(True); r.addWidget(self.ed_out, 1)
        o = QPushButton("Open");  o.clicked.connect(lambda: os.startfile(self.ed_out.text())); r.addWidget(o)
        c = QPushButton("Change"); c.clicked.connect(lambda: (p := QFileDialog.getExistingDirectory(self,"Output")) and self.ed_out.setText(p)); r.addWidget(c)
        return r

    def toggle_url(self):
        if self.btn_url.text() == "Edit":
            self.ed_url.setReadOnly(False); self.btn_url.setText("Save")
        else:
            self.ed_url.setReadOnly(True); self.btn_url.setText("Edit")

    def start_stop(self):
        if self.btn_start.text().endswith("START"):
            if not self.ed_excel.text():
                QMessageBox.warning(self, "Error", "Select Excel file first"); return
            self.stop_flag = False; self.pause_flag = False
            self.btn_start.setText("■  STOP"); self.btn_start.setStyleSheet(STOP_BTN)
            threading.Thread(target=self.run_worker, daemon=True).start()
        else:
            self.stop_flag = True; self.btn_start.setText("▶  START"); self.btn_start.setStyleSheet(START_BTN)

    def pause_resume(self):
        self.pause_flag = not self.pause_flag
        self.btn_pause.setText("▶  RESUME" if self.pause_flag else "⏸  PAUSE")

    def update_progress(self, cur, total):
        self.bar.setMaximum(total); self.bar.setValue(cur)
        self.lbl_prog.setText(f"{cur} / {total} bills")

    def on_finished(self):
        self.btn_start.setText("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_pause.setText("⏸  PAUSE")

    def _unique_pdf(self, folder, base):
        p = os.path.join(folder, f"{base}.pdf")
        if not os.path.exists(p): return p
        i = 1
        while True:
            p2 = os.path.join(folder, f"{base} ({i}).pdf"); 
            if not os.path.exists(p2): return p2
            i += 1

    def _inject_sr(self, driver, sr):
        driver.execute_script(f"""
        let d=document.getElementById('sr-overlay'); if(d) d.remove();
        d=document.createElement('div'); d.id='sr-overlay'; d.innerText='Sr. No. {sr}';
        d.style.cssText='position:fixed;top:10px;right:20px;font-size:20px;font-weight:bold;text-decoration:underline;z-index:99999;';
        document.body.appendChild(d);""")

    def run_worker(self):
        try:
            from selenium import webdriver
            from selenium.webdriver.common.by import By
            from selenium.webdriver.chrome.service import Service
            from webdriver_manager.chrome import ChromeDriverManager
            from openpyxl import load_workbook
            from pypdf import PdfReader, PdfWriter
        except ImportError as e:
            self.signals.log.emit(f"ERROR: Missing library — {e}"); self.signals.finished.emit(); return

        excel = self.ed_excel.text(); out = self.ed_out.text(); url = self.ed_url.text()
        os.makedirs(out, exist_ok=True)
        wb = load_workbook(excel); sh = wb.active

        opts = webdriver.ChromeOptions()
        if self.chk_headless.isChecked(): opts.add_argument("--headless=new")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

        total = sh.max_row - 1; done = 0; collected = []
        WAIT = 8; MAX_ATTEMPTS = 3

        for r in range(2, sh.max_row + 1):
            if self.stop_flag:
                sh.cell(r, 3).value = "Stopped"; wb.save(excel); break
            while self.pause_flag: time.sleep(0.3)
            ac = sh.cell(r, 1).value; sr = sh.cell(r, 2).value
            if not ac: break
            ac = str(ac).strip(); self.signals.log.emit(f"Processing AC: {ac}")
            status = None; bill_loaded = False

            for attempt in range(1, MAX_ATTEMPTS + 1):
                driver.get(url); time.sleep(2)
                box = driver.find_element(By.ID, "searchTextBox"); box.clear(); box.send_keys(ac)
                driver.find_element(By.ID, "btnSearch").click(); time.sleep(WAIT)
                if (not driver.find_elements(By.ID, "searchTextBox") and
                    driver.find_elements(By.XPATH, "//div[contains(@class,'tab-content') and contains(@class,'active')]//h2[contains(normalize-space(),'Bill Not Found')]")):
                    status = "Bill not Found"; break
                if driver.find_elements(By.XPATH, "//*[contains(text(),'does not belongs')]"):
                    status = "Wrong AC No."; break
                if (driver.find_elements(By.CSS_SELECTOR, ".auto-center.btn.btn-secondary") and
                    not driver.find_elements(By.ID, "searchTextBox")):
                    bill_loaded = True; status = "Saved"; break

            if status is None: status = "Bill not loaded (PITC Bug)"

            if not bill_loaded:
                sh.cell(r, 3).value = status; wb.save(excel); done += 1
                self.signals.progress.emit(done, total); self.signals.log.emit(f"→ {status}"); continue

            if self.chk_sr.isChecked() and sr:
                self._inject_sr(driver, sr); time.sleep(1)

            pdf = driver.execute_cdp_cmd("Page.printToPDF", {
                "printBackground": True, "paperWidth": 8.27, "paperHeight": 11.69,
                "marginTop": 0.15, "marginBottom": 0.15, "marginLeft": 0.15, "marginRight": 0.15, "scale": 0.60
            })
            pdf_name = str(sr).strip() if sr else ac
            p = self._unique_pdf(out, pdf_name)
            with open(p, "wb") as f: f.write(base64.b64decode(pdf["data"]))
            collected.append((int(sr) if sr and str(sr).isdigit() else 999999, p))
            sh.cell(r, 3).value = "Saved"; self.signals.log.emit("→ PDF created"); wb.save(excel)
            done += 1; self.signals.progress.emit(done, total)

        driver.quit()
        if not self.chk_sep.isChecked() and collected:
            if self.chk_sr.isChecked(): collected.sort(key=lambda x: x[0])
            merged = self._unique_pdf(out, os.path.splitext(os.path.basename(excel))[0])
            from pypdf import PdfWriter, PdfReader
            writer = PdfWriter()
            for _, pp in collected:
                for pg in PdfReader(pp).pages: writer.add_page(pg)
            with open(merged, "wb") as f: writer.write(f)
            for _, pp in collected: os.remove(pp)

        self.signals.finished.emit()


# ──────────────────────────────────────────────────────────────
#  TOOL 3 — Sub-Division Extractor (extractSd.py)
# ──────────────────────────────────────────────────────────────
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
        r1 = QHBoxLayout(); r1.addWidget(QLabel("Excel File:"))
        self.excel_ed = QLineEdit(); self.excel_ed.setReadOnly(True); r1.addWidget(self.excel_ed, 1)
        b1 = QPushButton("Browse"); b1.clicked.connect(lambda: (p := QFileDialog.getOpenFileName(self,"Excel","","Excel (*.xlsx)")[0]) and self.excel_ed.setText(p)); r1.addWidget(b1)
        btn_tpl = QPushButton("⬇  Download Template")
        btn_tpl.setStyleSheet("background:#194; border:none; border-radius:4px; padding:5px 10px; color:#fff;")
        btn_tpl.clicked.connect(self.get_template); r1.addWidget(btn_tpl)
        gv.addLayout(r1)
        r2 = QHBoxLayout(); r2.addWidget(QLabel("Text Folder:"))
        self.text_ed = QLineEdit(); self.text_ed.setReadOnly(True); r2.addWidget(self.text_ed, 1)
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
            default_name = "PaymentExtract_Template.xlsx",
            sample_rows  = [
                ["01", "101", "07269330425000", "1"],
                ["02", "102", "07269330426000", "2"],
            ]
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
            current = folder_data[0][0][0] if folder_data[0] else None
            for idx2, files in enumerate(folder_data, start=1):
                if not files: continue
                para, fname = files[0]
                if current and para == current: files.pop(0)
                elif current and self.natural_key(para) < self.natural_key(current): files.pop(0)
                else: continue
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
class PDFMergeWorkerSignals(QObject):
    update = Signal(str, int, int, bool)


class PDFMergePanel(QWidget):
    def __init__(self):
        super().__init__()
        self.stop_flag = {"stop": False}
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
            threading.Thread(target=self.merge_worker, daemon=True).start()
        else:
            self.stop_flag["stop"] = True

    def on_update(self, msg, cur, total, done):
        self.log_lbl.setText(msg)
        if total: self.progress.setMaximum(total); self.progress.setValue(cur); self.lbl_prog.setText(f"{cur} / {total}")
        if done: self.btn_start.setText("▶  START"); self.btn_start.setStyleSheet(START_BTN)

    def merge_worker(self):
        from pypdf import PdfReader, PdfWriter
        folder = self.folder_ed.text()
        pdfs = sorted([f for f in os.listdir(folder) if f.lower().endswith(".pdf")], key=self.natural_key)
        if not pdfs: self.signals.update.emit("No PDFs found.", 0, 0, True); return
        writer = PdfWriter(); total = len(pdfs)
        for i, pdf in enumerate(pdfs, 1):
            if self.stop_flag["stop"]: self.signals.update.emit("Stopped.", i, total, True); return
            self.signals.update.emit(f"Merging: {pdf}", i, total, False)
            try:
                reader = PdfReader(os.path.join(folder, pdf))
                for page in reader.pages: writer.add_page(page)
            except Exception: pass
        out_name = f"Merged_{os.path.basename(folder)}.pdf"
        out_path = os.path.join(folder, out_name)
        with open(out_path, "wb") as f: writer.write(f)
        self.signals.update.emit(f"Saved: {out_name} ✔", total, total, True)


# ──────────────────────────────────────────────────────────────
#  TOOL 10 — File Existence Checker (SrNoCheckinfolder.py)
# ──────────────────────────────────────────────────────────────
class SrNoCheckerPanel(QWidget):
    def __init__(self):
        super().__init__()
        self._stop_flag = False
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("File Existence Checker", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Checks if Sr. No. from Excel column A exists as a file in the selected folder.", color="#556"))
        layout.addWidget(hline())

        g = QGroupBox("Inputs")
        gv = QVBoxLayout(g)

        r1 = QHBoxLayout(); r1.addWidget(QLabel("Excel File:"))
        self.excel_ed = QLineEdit(); self.excel_ed.setReadOnly(True); r1.addWidget(self.excel_ed, 1)
        b1 = QPushButton("Browse"); b1.clicked.connect(self._browse_excel); r1.addWidget(b1)
        gv.addLayout(r1)

        r2 = QHBoxLayout(); r2.addWidget(QLabel("Folder:"))
        self.folder_ed = QLineEdit(); self.folder_ed.setReadOnly(True); r2.addWidget(self.folder_ed, 1)
        b2 = QPushButton("Browse"); b2.clicked.connect(self._browse_folder); r2.addWidget(b2)
        gv.addLayout(r2)
        layout.addWidget(g)

        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_stop  = QPushButton("■  STOP");  self.btn_stop.setStyleSheet(STOP_BTN)
        self.btn_start.clicked.connect(self._start); self.btn_stop.clicked.connect(self._stop)
        bh.addWidget(self.btn_start); bh.addWidget(self.btn_stop); bh.addStretch()
        layout.addLayout(bh)

        self.status_lbl = QLabel("Status: Idle"); self.status_lbl.setStyleSheet("color:#4a90d9;")
        layout.addWidget(self.status_lbl)

    def _browse_excel(self):
        f, _ = QFileDialog.getOpenFileName(self, "Excel File", "", "Excel Files (*.xlsx)")
        if f: self.excel_ed.setText(f)

    def _browse_folder(self):
        p = QFileDialog.getExistingDirectory(self, "Select Folder")
        if p: self.folder_ed.setText(p)

    def _start(self):
        self._stop_flag = False
        threading.Thread(target=self._process, daemon=True).start()

    def _stop(self):
        self._stop_flag = True
        self.status_lbl.setText("Status: Stopped")

    def _process(self):
        from openpyxl import load_workbook
        excel_path  = self.excel_ed.text()
        folder_path = self.folder_ed.text()

        if not excel_path or not folder_path:
            QMessageBox.critical(self, "Error", "Please select Excel file and Folder"); return

        try:
            wb = load_workbook(excel_path); ws = wb.active
            files_in_folder = os.listdir(folder_path)
            row = 2
            while True:
                if self._stop_flag: break
                sr_no = ws[f"A{row}"].value
                if sr_no is None: break
                found = any(os.path.splitext(f)[0] == str(sr_no) for f in files_in_folder)
                ws[f"B{row}"] = "Found" if found else "Not Found"
                self.status_lbl.setText(f"Checking Sr. No.: {sr_no}")
                row += 1
            wb.save(excel_path)
            self.status_lbl.setText("Status: Completed")
            QMessageBox.information(self, "Done", "Process Completed Successfully")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


# ──────────────────────────────────────────────────────────────
#  TOOL 11 — Double CR Checker (DoubleCRchecker.py)
# ──────────────────────────────────────────────────────────────
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
def download_template(parent, headers, default_name, sample_rows=None):
    from openpyxl import Workbook as _WB
    from openpyxl.styles import Font, PatternFill, Alignment
    path, _ = QFileDialog.getSaveFileName(
        parent, "Save Template As", default_name, "Excel (*.xlsx)"
    )
    if not path:
        return
    wb = _WB(); ws = wb.active; ws.title = "Data"
    hdr_font  = Font(bold=True, color="FFFFFF")
    hdr_fill  = PatternFill("solid", fgColor="1F4E79")
    hdr_align = Alignment(horizontal="center")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = hdr_align
        ws.column_dimensions[cell.column_letter].width = max(len(h) + 6, 16)
    if sample_rows:
        for r, row_data in enumerate(sample_rows, 2):
            for c, val in enumerate(row_data, 1):
                ws.cell(row=r, column=c, value=val)
    wb.save(path)
    QMessageBox.information(parent, "Template Saved", f"Template saved to:\n{path}")


# ──────────────────────────────────────────────────────────────
#  SETTINGS PANEL — edit & save config.ini from inside the app
# ──────────────────────────────────────────────────────────────
def _config_path():
    """Always resolves config.ini next to the running EXE or script."""
    return os.path.join(os.path.dirname(sys.argv[0]), "config.ini")


class SettingsPanel(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self); layout.setSpacing(14)

        layout.addWidget(lbl("Settings  ·  PITC Credentials", bold=True, color="#7eb8f7"))
        layout.addWidget(hline())

        g = QGroupBox("PITC Login (cp.pitc.com.pk)")
        gv = QVBoxLayout(g); gv.setSpacing(10)

        r1 = QHBoxLayout()
        r1.addWidget(QLabel("Username:")); self.user_ed = QLineEdit(); r1.addWidget(self.user_ed, 1)
        gv.addLayout(r1)

        r2 = QHBoxLayout()
        r2.addWidget(QLabel("Password:")); self.pass_ed = QLineEdit()
        self.pass_ed.setEchoMode(QLineEdit.Password); r2.addWidget(self.pass_ed, 1)
        self.chk_show = QCheckBox("Show password")
        self.chk_show.toggled.connect(lambda on: self.pass_ed.setEchoMode(QLineEdit.Normal if on else QLineEdit.Password))
        r2.addWidget(self.chk_show)
        gv.addLayout(r2)

        r3 = QHBoxLayout()
        r3.addWidget(QLabel("OTP Secret:")); self.otp_ed = QLineEdit()
        self.otp_ed.setEchoMode(QLineEdit.Password)
        self.otp_ed.setPlaceholderText("TOTP base32 secret (leave blank if no OTP required)")
        r3.addWidget(self.otp_ed, 1)
        self.chk_otp_show = QCheckBox("Show")
        self.chk_otp_show.toggled.connect(lambda on: self.otp_ed.setEchoMode(QLineEdit.Normal if on else QLineEdit.Password))
        r3.addWidget(self.chk_otp_show)
        gv.addLayout(r3)

        layout.addWidget(g)

        note = QLabel(f"Saved to:  {_config_path()}")
        note.setStyleSheet("color:#556; font-size:10px;")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.path_lbl = note  # update after save

        bh = QHBoxLayout()
        btn_save = QPushButton("💾  Save Credentials")
        btn_save.setStyleSheet(START_BTN + " padding:8px 24px;")
        btn_save.clicked.connect(self.save)
        bh.addWidget(btn_save); bh.addStretch()
        layout.addLayout(bh)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color:#4c4; font-size:11px;")
        layout.addWidget(self.status_lbl)

        layout.addStretch()
        layout.addWidget(lbl("Tip: credentials are stored in plain text in config.ini.\nKeep the file in a safe location.", color="#556"))

        # load existing values on open
        self.load()

    def load(self):
        path = _config_path()
        cfg = configparser.ConfigParser()
        if os.path.exists(path):
            cfg.read(path)
            self.user_ed.setText(cfg.get("credentials", "username", fallback=""))
            self.pass_ed.setText(cfg.get("credentials", "password", fallback=""))
            self.otp_ed.setText(cfg.get("credentials", "otp_secret", fallback=""))
            self.status_lbl.setText("")
        else:
            self.status_lbl.setText("⚠  config.ini not found — will be created on Save.")
            self.status_lbl.setStyleSheet("color:#ca4; font-size:11px;")

    def save(self):
        username   = self.user_ed.text().strip()
        password   = self.pass_ed.text().strip()
        otp_secret = self.otp_ed.text().strip()
        if not username or not password:
            QMessageBox.warning(self, "Missing", "Username and password cannot be empty."); return

        path = _config_path()
        cfg = configparser.ConfigParser()
        cfg["credentials"] = {"username": username, "password": password, "otp_secret": otp_secret}
        with open(path, "w") as f:
            cfg.write(f)

        self.status_lbl.setText("✔  Saved successfully.")
        self.status_lbl.setStyleSheet("color:#4c4; font-size:11px;")
        self.path_lbl.setText(f"Saved to:  {path}")


# ──────────────────────────────────────────────────────────────
#  TOOL 10 — Extract Specific PITC Page  (ExtPITC.py)
#  Search Criteria: month (Feb-26 → Feb - 2026) or any number
#  Extracts matching page · Highlights matches · Stamps Para Sr. No.
# ──────────────────────────────────────────────────────────────
class ExtractPageWorkerSignals(QObject):
    log      = Signal(str)
    progress = Signal(int, int)
    finished = Signal()


class ExtractPagePanel(QWidget):
    def __init__(self):
        super().__init__()
        self.signals = ExtractPageWorkerSignals()
        self.signals.log.connect(lambda m: self.log_box.append(m))
        self.signals.progress.connect(self.update_progress)
        self.signals.finished.connect(self.on_finished)

        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("Extract Specific PITC Page", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Search Criteria: month e.g. Feb-26  or  any number  ·  Extracts matching page  ·  Highlights  ·  Stamps Para Sr. No.", color="#556"))
        layout.addWidget(hline())

        g = QGroupBox("Inputs")
        gv = QVBoxLayout(g)

        r1 = QHBoxLayout(); r1.addWidget(QLabel("PDF Folder:"))
        self.pdf_ed = QLineEdit(); self.pdf_ed.setReadOnly(True); r1.addWidget(self.pdf_ed, 1)
        b1 = QPushButton("Browse"); b1.clicked.connect(lambda: (p := QFileDialog.getExistingDirectory(self,"PDF Folder")) and self.pdf_ed.setText(p)); r1.addWidget(b1)
        gv.addLayout(r1)

        r2 = QHBoxLayout(); r2.addWidget(QLabel("Excel File:"))
        self.excel_ed = QLineEdit(); self.excel_ed.setReadOnly(True); r2.addWidget(self.excel_ed, 1)
        b2 = QPushButton("Browse"); b2.clicked.connect(lambda: (p := QFileDialog.getOpenFileName(self,"Excel","","Excel (*.xlsx)")[0]) and self.excel_ed.setText(p)); r2.addWidget(b2)
        btn_tpl = QPushButton("⬇  Download Template")
        btn_tpl.setStyleSheet("background:#194; border:none; border-radius:4px; padding:5px 10px; color:#fff;")
        btn_tpl.clicked.connect(self.get_template); r2.addWidget(btn_tpl)
        gv.addLayout(r2)

        r3 = QHBoxLayout(); r3.addWidget(QLabel("Output Folder:"))
        self.out_ed = QLineEdit(r"D:\ExtractedPages"); r3.addWidget(self.out_ed, 1)
        b3 = QPushButton("Change"); b3.clicked.connect(lambda: (p := QFileDialog.getExistingDirectory(self,"Output")) and self.out_ed.setText(p)); r3.addWidget(b3)
        gv.addLayout(r3)

        layout.addWidget(g)

        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START"); self.btn_start.setStyleSheet(START_BTN)
        self.btn_stop  = QPushButton("■  STOP");  self.btn_stop.setStyleSheet(STOP_BTN); self.btn_stop.setEnabled(False)
        self.btn_start.clicked.connect(self.start); self.btn_stop.clicked.connect(self.stop)
        bh.addWidget(self.btn_start); bh.addWidget(self.btn_stop); bh.addStretch()
        layout.addLayout(bh)

        ph = QHBoxLayout()
        self.lbl_prog = QLabel("0 / 0"); ph.addWidget(self.lbl_prog); ph.addStretch()
        layout.addLayout(ph)
        self.progress = QProgressBar(); layout.addWidget(self.progress)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(self.log_box)

        self._stop_flag = False

    def get_template(self):
        download_template(
            self,
            headers      = ["AC No", "Search Criteria", "Para Sr. No."],
            default_name = "ExtractPage_Template.xlsx",
            sample_rows  = [
                ["07269330425000", "Jan-26", "101"],
                ["07269330426000", 5250,     "102"],
            ]
        )

    def update_progress(self, cur, total):
        self.progress.setMaximum(total); self.progress.setValue(cur)
        self.lbl_prog.setText(f"{cur} / {total}")

    def on_finished(self):
        self.btn_start.setEnabled(True); self.btn_stop.setEnabled(False)
        QMessageBox.information(self, "Done", "Extraction completed.")

    def start(self):
        if not self.pdf_ed.text() or not self.excel_ed.text():
            QMessageBox.warning(self, "Missing", "Select PDF folder and Excel file."); return
        self._stop_flag = False
        self.btn_start.setEnabled(False); self.btn_stop.setEnabled(True)
        self.log_box.clear()
        threading.Thread(target=self.process, daemon=True).start()

    def stop(self):
        self._stop_flag = True; self.log_box.append("⛔ Stopping…")

    def process(self):
        try:
            import fitz
            from openpyxl import load_workbook
            import datetime as _dt
        except ImportError as e:
            self.signals.log.emit(f"ERROR: Missing library — {e}")
            self.signals.finished.emit(); return

        def parse_criteria(val):
            """Feb-26 → 'Feb - 2026'.  Pure number or anything else → str as-is."""
            s = str(val).strip()
            for fmt in ("%b-%y", "%b-%Y", "%m-%y", "%m-%Y"):
                try:
                    return _dt.datetime.strptime(s, fmt).strftime("%b - %Y")
                except:
                    pass
            return s

        pdf_folder = self.pdf_ed.text()
        out_folder = self.out_ed.text()
        os.makedirs(out_folder, exist_ok=True)

        wb = load_workbook(self.excel_ed.text()); ws = wb.active
        last_row = ws.max_row
        self.signals.progress.emit(0, last_row - 1)

        for row in range(2, last_row + 1):
            if self._stop_flag: break

            ac_no     = str(ws[f"A{row}"].value or "").strip()
            criteria  = str(ws[f"B{row}"].value or "").strip()
            para_text = str(ws[f"C{row}"].value or "").strip()

            if not ac_no: continue

            search_text = parse_criteria(criteria)
            self.signals.log.emit(f"AC: {ac_no}  |  Searching: '{search_text}'")

            filename = ac_no if ac_no.lower().endswith(".pdf") else ac_no + ".pdf"
            pdf_path = os.path.join(pdf_folder, filename)

            if not os.path.exists(pdf_path):
                ws[f"D{row}"] = "File Not Found"
                self.signals.log.emit("❌ File Not Found")
                self.signals.progress.emit(row - 1, last_row - 1)
                continue

            self.signals.log.emit("✔ File Found")

            try:
                doc   = fitz.open(pdf_path)
                found = False

                for pg_idx, page in enumerate(doc):
                    areas = page.search_for(search_text)
                    # fallback: try compact form e.g. "Feb-2026" if spaced form not found
                    if not areas and " - " in search_text:
                        areas = page.search_for(search_text.replace(" - ", "-"))

                    if areas:
                        found = True
                        self.signals.log.emit(f"✔ Match on page {pg_idx + 1}")

                        # highlight all matches
                        for rect in areas:
                            page.add_highlight_annot(rect).update()

                        # stamp Para Sr. No. top-right
                        if para_text and para_text.lower() != "none":
                            stamp = f"Para Sr. No. {para_text}"
                            pw    = page.rect.width
                            tw    = fitz.get_text_length(stamp, fontname="helv", fontsize=14)
                            page.insert_text(
                                fitz.Point(pw - tw - 30, 26),
                                stamp,
                                fontsize=14,
                                fontname="helv",
                                color=(0, 0, 0),
                            )

                        # save single extracted page
                        out_doc = fitz.open()
                        out_doc.insert_pdf(doc, from_page=pg_idx, to_page=pg_idx)

                        out_path = os.path.join(out_folder, filename)
                        if os.path.exists(out_path):
                            base, ext = os.path.splitext(filename)
                            i = 1
                            while os.path.exists(os.path.join(out_folder, f"{base}_{i}{ext}")):
                                i += 1
                            out_path = os.path.join(out_folder, f"{base}_{i}{ext}")

                        out_doc.save(out_path, garbage=4, deflate=True)
                        out_doc.close()

                        ws[f"D{row}"] = f"Extracted page {pg_idx + 1}"
                        self.signals.log.emit("✅ Extracted & highlighted")
                        break

                doc.close()

                if not found:
                    ws[f"D{row}"] = "Not Found"
                    self.signals.log.emit(f"❌ '{search_text}' not found in PDF")

            except Exception as e:
                ws[f"D{row}"] = "Error"
                self.signals.log.emit(f"⚠ {e}")

            self.signals.progress.emit(row - 1, last_row - 1)

        wb.save(self.excel_ed.text())
        self.signals.finished.emit()

# ──────────────────────────────────────────────────────────────
#  MAIN WINDOW
# ──────────────────────────────────────────────────────────────
TOOLS = [
    ("📥  PITC Downloader",               PITCPanel),
    ("🔍  Extract Specific PITC Page",    ExtractPagePanel),
    ("🧾  E/Bills Downloader",            BillSoftPanel),
    ("✂️   SDiv Extractor from 88L",      ExtractSdPanel),
    ("💰  Payment Extractor from 88L",    PaymentExtractPanel),
    ("📚  Multi-Merge",                   MultiMergePanel),
    ("🔗  PDF Merge",                     PDFMergePanel),
    ("📂  File Existence Checker",        SrNoCheckerPanel),
    ("🔁  Double CR Checker",             DoubleCRPanel),
    ("⚙️   Settings",                     SettingsPanel),
]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PITC Suite  ·  Developed by Tahir Naqash")
        self.resize(820, 620)

        root = QWidget(); self.setCentralWidget(root)
        outer = QHBoxLayout(root); outer.setContentsMargins(0,0,0,0); outer.setSpacing(0)

        # ── NAV SIDEBAR ──
        nav = QWidget(); nav.setFixedWidth(300)
        nav.setStyleSheet("background:#111318; border-right:1px solid #1e2230;")
        nav_layout = QVBoxLayout(nav); nav_layout.setContentsMargins(0,0,0,0); nav_layout.setSpacing(0)

        title_lbl = QLabel("PITC\nSuite")
        title_lbl.setAlignment(Qt.AlignCenter)
        title_lbl.setStyleSheet("color:#4a90d9; font-size:18px; font-weight:bold; padding:20px 10px 14px; letter-spacing:2px;")
        nav_layout.addWidget(title_lbl)
        sep = QFrame(); sep.setFrameShape(QFrame.HLine); sep.setStyleSheet("color:#1e2230;"); nav_layout.addWidget(sep)

        self.nav_btns = []
        self.stack = QStackedWidget()

        for i, (name, PanelClass) in enumerate(TOOLS):
            btn = QPushButton(name)
            btn.setStyleSheet(NAV_STYLE)
            btn.setCheckable(False)
            btn.clicked.connect(lambda _, idx=i: self.switch(idx))
            nav_layout.addWidget(btn)
            self.nav_btns.append(btn)
            self.stack.addWidget(PanelClass())

        nav_layout.addStretch()
        ver = QLabel("v1.5"); ver.setAlignment(Qt.AlignCenter)
        ver.setStyleSheet("color:#3a4a6a; font-size:10px; padding:8px;")
        nav_layout.addWidget(ver)

        outer.addWidget(nav)

        # ── CONTENT AREA ──
        content = QWidget()
        cl = QVBoxLayout(content); cl.setContentsMargins(16,16,16,12); cl.setSpacing(6)
        cl.addWidget(self.stack)
        outer.addWidget(content, 1)

        self.switch(0)

    def switch(self, idx):
        self.stack.setCurrentIndex(idx)
        for i, btn in enumerate(self.nav_btns):
            btn.setProperty("active", "true" if i == idx else "false")
            btn.style().unpolish(btn); btn.style().polish(btn)


# ──────────────────────────────────────────────────────────────
#  ENTRY POINT
# ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(DARK_STYLE)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
