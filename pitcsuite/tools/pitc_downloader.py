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
            from selenium.common.exceptions import InvalidSessionIdException, StaleElementReferenceException, TimeoutException
            from webdriver_manager.chrome import ChromeDriverManager
            from openpyxl import load_workbook
            from reportlab.pdfgen import canvas
            from pypdf import PdfReader, PdfWriter
            import pyautogui, fitz, shutil, traceback
            import datetime as _dt
            import pyotp
        except ImportError as e:
            self.log.emit(f"ERROR: Missing library — {e}")
            return

        cfg = configparser.ConfigParser()
        cfg.read(_config_path())
        try:
            USERNAME   = cfg["credentials"]["username"]
            PASSWORD   = cfg["credentials"]["password"]
            OTP_SECRET = cfg["credentials"].get("otp_secret", "")
        except Exception:
            self.log.emit("ERROR: config.ini missing or invalid. Set credentials in ⚙️ Settings first.")
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
                        otp_code = generate_otp()
                        self.log.emit(f"OTP attempt {attempt+1}: {otp_code}")
                        otp_box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
                        otp_box.send_keys(otp_code)
                        wt.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnVerify"))).click()
                        WebDriverWait(drv, 10).until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber")))
                        if drv.find_elements(By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"):
                            self.log.emit("OTP verified ✔")
                            return
                        time.sleep(0.5)
                    except TimeoutException:
                        self.log.emit("OTP not accepted; retrying")
                    except InvalidSessionIdException:
                        raise
                    except Exception as e:
                        self.log.emit(f"OTP retry: {e}")
                    time.sleep(1)
                raise Exception("OTP verification failed")

        def ensure_logged_in(drv, wt):
            if not drv.find_elements(By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"):
                self.log.emit("Session expired → re-login")
                do_login(drv, wt)

        def login_with_session_recovery(drv, wt):
            try:
                do_login(drv, wt)
                return drv, wt
            except InvalidSessionIdException:
                self.log.emit("Browser session lost during OTP; restarting Chrome")
                try:
                    drv.quit()
                except Exception:
                    pass
                drv, wt = start_driver()
                do_login(drv, wt)
                return drv, wt

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
        driver, wait = login_with_session_recovery(driver, wait)

        wb         = load_workbook(self.excel)
        sheet      = wb.active
        total_rows = sheet.max_row - 1
        processed  = 0
        row        = 2
        href_kw    = REPORT_HREF[self.report_key]

        wait.until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber")))
        if self.warmup > 0:
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
                driver, wait = login_with_session_recovery(driver, wait)
                wait.until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber")))
                if self.warmup > 0:
                    time.sleep(self.warmup)

        wb.save(self.excel)
        driver.quit()


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
        g3h.addWidget(QLabel("Warmup (s)"));     self.sp_warmup  = spin(0, 0, 200); g3h.addWidget(self.sp_warmup)
        g3h.addWidget(QLabel("Print Wait (s)")); self.sp_pwait   = spin(6,  1, 200); g3h.addWidget(self.sp_pwait)
        g3h.addWidget(QLabel("Print X"));        self.sp_px      = spin(30);         g3h.addWidget(self.sp_px)
        g3h.addWidget(QLabel("Print Y"));        self.sp_py      = spin(165);        g3h.addWidget(self.sp_py)
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
            default_name = "PITC_Downloader_Template.xlsx"
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
