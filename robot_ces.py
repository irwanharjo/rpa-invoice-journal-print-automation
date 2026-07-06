"""
Robot CES - Automated receivables-journal lookup + invoice print workflow for a corporate
finance web portal (genericized for portfolio use; original built against a real internal
portal referred to here simply as "the portal").

Flow:
  1. Read every "Debet Nota Receipt Summary" PDF from the Siap_Input folder
  2. Extract the DN (Debit Note) number (3rd column of the table) -> also saved to Excel
     (Hasil_Konversi) as a safety net so it can be checked manually, nothing gets missed
  3. Log into the portal
  4. Navigate Finance -> Receivables -> Receivables Journal
  5. For each DN number: fill the DN field, click the search (magnifying glass) icon
  6. If status is "Not Verified": click the verify (thumbs-up) icon, confirm -> journal posted
  7. Print invoice: icon A ("no signature required") if not a priority client & amount >= a
     configurable threshold; icon B ("signature required") if priority client OR amount below
     threshold
  8. Fill in the invoice date (today), click Print -> invoice opens in a new tab
  9. Download the invoice PDF from that tab, then print it physically to a user-chosen
     printer via SumatraPDF's silent command-line printing (see cetak_pdf_langsung() for why
     browser kiosk-printing was abandoned)

Safe to resume: every DN number that completes successfully is recorded in log_progress.json.
If the process is interrupted mid-run (e.g. printer jam / power cut), restarting automatically
skips DN numbers already completed and resumes from the rest -- no invoice ever gets printed
twice. The trigger PDF is only moved to Selesai_Input once every DN number inside it is done.

Status: validated end-to-end manually against the real portal for a sample DN number.
Extra dependencies: pywin32 (win32print for queue checks), SumatraPDF portable (silent print).
"""

import base64
import json
import os
import re
import shutil
import subprocess
import threading
import time
import tkinter as tk
from datetime import date, datetime
from tkinter import messagebox

import pdfplumber
import openpyxl
from openpyxl.styles import Font
import win32print
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException

load_dotenv()

# ============================================================
# CONFIGURATION
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_SIAP = os.path.join(BASE_DIR, "Siap_Input")
FOLDER_SELESAI = os.path.join(BASE_DIR, "Selesai_Input")
FOLDER_HASIL = os.path.join(BASE_DIR, "Hasil_Konversi")
FOLDER_CETAK_SEMENTARA = os.path.join(BASE_DIR, "Cetak_Sementara")
FILE_LOG_PROGRESS = os.path.join(BASE_DIR, "log_progress.json")

PORTAL_URL = os.environ.get("PORTAL_BASE_URL", "https://your-finance-portal.example.com")
PORTAL_USERNAME = os.environ.get("PORTAL_USERNAME", "")
PORTAL_PASSWORD = os.environ.get("PORTAL_PASSWORD", "")

URL_LOGIN = f"{PORTAL_URL}/login"
URL_JURNAL_PIUTANG = f"{PORTAL_URL}/jurnalPiutang"
URL_DEBET_NOTA = f"{PORTAL_URL}/debetNota"

# SumatraPDF portable, used specifically for silent printing to a chosen printer -- see
# cetak_pdf_langsung() for why this approach was chosen over browser kiosk-printing.
SUMATRA_EXE = os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe")

FOLDER_PRIORITY_STAGING = os.path.join(BASE_DIR, "PriorityClient_Staging")
FOLDER_PRIORITY_ARSIP = os.path.join(BASE_DIR, "PriorityClient_Arsip")
FOLDER_SIAP_APS = os.environ.get("DOWNSTREAM_CLAIMS_INPUT_FOLDER", os.path.join(BASE_DIR, "..", "rpa-claims-portal-automation", "Siap_Input"))

# DN number pattern, e.g.: N017DN00000002
POLA_NO_DN = re.compile(r"N\d{3}DN\d{8}")

# Selectors confirmed via direct inspection of the portal's DOM
SELECTOR = {
    "input_username": (By.NAME, "username"),
    "input_password": (By.NAME, "password"),
    "tombol_login": (By.ID, "kt_login_signin_submit"),
    "input_no_dn": (By.CSS_SELECTOR, "input[name='srcNoDebetNota']"),
    "tombol_cari_kaca_pembesar": (By.ID, "cari"),
    "baris_hasil": (By.CSS_SELECTOR, "#kt_datatable tbody tr"),
    "tombol_verifikasi": (By.XPATH, "//*[@id='kt_datatable']/tbody/tr/td[11]/center/a/i"),
    "tombol_cetak_ungu": (By.CSS_SELECTOR, "a[href^=\"javascript:cetakInvoice('\"]"),
    "tombol_cetak_merah": (By.CSS_SELECTOR, "a[href^=\"javascript:cetakInvoiceTtd('\"]"),
    "input_tgl_invoice": (By.NAME, "modalInvoiceDate"),
    "tombol_print_invoice": (By.ID, "printInvoice"),
    # Debet Nota page (for downloading the Debet Nota recap, separate from the journal page)
    "debetnota_input_periode": (By.ID, "srcPeriode"),
    "debetnota_input_no_dn": (By.ID, "srcNoDn"),
    "debetnota_tombol_cari": (By.ID, "cari"),
    "debetnota_baris_hasil": (By.CSS_SELECTOR, "#kt_datatable tbody tr"),
    # First icon in the Actions column -- onclick="cetakDebetNota(...)", NOT cetakBillingAll/exportDebetNota
    "debetnota_tombol_cetak": (By.CSS_SELECTOR, "a[href^=\"javascript:cetakDebetNota('\"]"),
}

# Illustrative entity codes for a "priority client" whose invoices always require a
# signature regardless of amount. Replace with your own portal's entity codes.
KODE_PRIORITY_CLIENT = {
    "E01AAA", "E01DDD", "E01CCC", "E01EEE", "E01FFF", "E0CCCC", "E0CAAA",
    "E01GGG", "E01BBB", "E01HHH", "E01III", "E01JJJ", "E00KKK", "E01LLL",
}

BATAS_NOMINAL_TTD = int(os.environ.get("SIGNATURE_THRESHOLD_AMOUNT", "5000000"))  # below this, signature icon is mandatory
PRINTER_SARAN_AWAL = os.environ.get("DEFAULT_PRINTER_NAME", "Your Default Printer")  # just a pre-select suggestion; printer is chosen manually every print


def _parse_nominal(teks):
    """Parse an amount string from the portal's table into a float. The portal turned out to
    use the international number format (comma = thousands separator, dot = decimal, e.g.
    '73,408,168'), NOT the Indonesian format initially assumed -- confirmed via direct
    inspection of the Receivables Journal table (this earlier wrong assumption was the root
    cause of repeated "could not convert string to float" errors in log_progress.json). Made
    defensive with a whitelist regex (strip everything except digits/comma/dot) so it's
    resilient to unexpected characters (non-breaking spaces, currency symbols, etc.)."""
    hanya_dipakai = re.sub(r"[^0-9,.]", "", teks)
    if not hanya_dipakai:
        return 0.0
    return float(hanya_dipakai.replace(",", ""))


def turunkan_periode_dari_no_dn(no_dn):
    """N017DN00000003 -> '2026-06'. Digits 1-2 after 'DN' = year (26->2026), digits 3-4 =
    month (06). Used to fill the Period field on the Debet Nota page."""
    angka = re.search(r"DN(\d{8})", no_dn).group(1)
    tahun = "20" + angka[0:2]
    bulan = angka[2:4]
    return f"{tahun}-{bulan}"


def eselon_termasuk_priority_client(teks_eselon):
    """Check whether this journal row belongs to the "priority client" (which always
    requires the signature/Ttd icon)."""
    teks_upper = teks_eselon.upper()
    if "PRIORITY" in teks_upper:
        return True
    for kode in KODE_PRIORITY_CLIENT:
        if kode in teks_upper:
            return True
    return False


# ============================================================
# STAGE 1 - EXTRACT DN NUMBERS FROM PDF
# ============================================================
def ekstrak_no_dn_dari_pdf(path_pdf):
    """Read the "Debet Nota Receipt Summary" table, grab the DN number in the 3rd column.
    Falls back to a regex scan of the full text if table extraction fails/columns are
    inconsistent, so no DN number is ever missed."""
    hasil = []
    with pdfplumber.open(path_pdf) as pdf:
        for halaman in pdf.pages:
            tabel = halaman.extract_table()
            if tabel:
                for baris in tabel:
                    if baris and len(baris) > 2:
                        sel = (baris[2] or "").strip()
                        if POLA_NO_DN.fullmatch(sel):
                            hasil.append(sel)

    if not hasil:
        with pdfplumber.open(path_pdf) as pdf:
            teks_penuh = "\n".join(h.extract_text() or "" for h in pdf.pages)
        hasil = POLA_NO_DN.findall(teks_penuh)

    # Second safety net: cross-check against a full-text regex scan, adding any DN numbers
    # that might have slipped through table extraction
    with pdfplumber.open(path_pdf) as pdf:
        teks_penuh = "\n".join(h.extract_text() or "" for h in pdf.pages)
    dari_regex = POLA_NO_DN.findall(teks_penuh)
    for no_dn in dari_regex:
        if no_dn not in hasil:
            hasil.append(no_dn)

    return hasil


def simpan_ke_excel(no_dn_list, nama_file_asal):
    """Save the list of DN numbers to Excel in Hasil_Konversi as a manual-verification safety net."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "No DN"
    ws["A1"] = "No"
    ws["B1"] = "No DN"
    ws["A1"].font = ws["B1"].font = Font(bold=True)

    for i, no_dn in enumerate(no_dn_list, start=1):
        ws.cell(row=i + 1, column=1, value=i)
        ws.cell(row=i + 1, column=2, value=no_dn)

    nama_dasar = os.path.splitext(os.path.basename(nama_file_asal))[0]
    path_keluar = os.path.join(FOLDER_HASIL, f"{nama_dasar}.xlsx")
    wb.save(path_keluar)
    print(f"  -> Verification Excel saved: {path_keluar}")
    return path_keluar


# ============================================================
# PER-DN PROGRESS LOG (safe resume after interruption/restart)
# ============================================================
def muat_log_progress():
    """Read log_progress.json. If it doesn't exist / is corrupted, start empty."""
    if not os.path.exists(FILE_LOG_PROGRESS):
        return {}
    try:
        with open(FILE_LOG_PROGRESS, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def simpan_log_progress(log):
    """Rewrite the log to disk. Called after EVERY DN number completes/fails, not just at
    the end of the batch, so progress isn't lost if the process dies unexpectedly."""
    with open(FILE_LOG_PROGRESS, "w", encoding="utf-8") as f:
        json.dump(log, f, indent=2, ensure_ascii=False)


def tandai_selesai(log, no_dn, nama_file):
    log[no_dn] = {"status": "selesai", "file": nama_file, "waktu": datetime.now().isoformat(timespec="seconds")}
    simpan_log_progress(log)


def tandai_gagal(log, no_dn, nama_file, pesan_error):
    log[no_dn] = {"status": "gagal", "file": nama_file, "waktu": datetime.now().isoformat(timespec="seconds"), "error": str(pesan_error)}
    simpan_log_progress(log)


def sudah_selesai(log, no_dn):
    return log.get(no_dn, {}).get("status") == "selesai"


def _noop(_teks):
    pass


# ============================================================
# STAGE 2 - LOGIN & PORTAL NAVIGATION
# ============================================================
def buka_browser():
    opsi = Options()
    opsi.add_experimental_option("excludeSwitches", ["enable-automation"])
    # Manual testing in a regular browser is fast (<5s), but via Selenium it hangs/lags on
    # interaction ("timed out receiving message from renderer") -- a classic symptom of a
    # site detecting navigator.webdriver=true and running extra verification scripts that
    # weigh the tab down. This flag stops Blink from marking the tab as automation-controlled.
    opsi.add_argument("--disable-blink-features=AutomationControlled")
    # eager: Selenium proceeds as soon as the DOM is ready (interactive), without waiting for
    # every network resource to finish -- the portal has something (analytics/websocket/
    # polling) that keeps the "load" event from ever firing, which made driver.get() hang
    # until the 120s read timeout even though the page had rendered normally.
    opsi.page_load_strategy = "eager"
    driver = webdriver.Chrome(options=opsi)
    driver.maximize_window()
    # Backup in case "eager" alone isn't enough: cap driver.get() at 20 seconds. If the
    # portal never fires a full load event (polling/websocket keeps running), driver.get()
    # raises TimeoutException after 20s (rather than hanging the full 120s via
    # _get_dengan_timeout()) -- the DOM is usually ready even though "load" hasn't fired.
    driver.set_page_load_timeout(20)
    return driver


def _get_dengan_timeout(driver, url):
    """A driver.get() that's resilient to sites that never fire a full load event. The
    timeout is set via driver.set_page_load_timeout() in buka_browser(). If it times out,
    catch it here and continue -- in practice the DOM is already ready (confirmed in earlier
    sessions: the page title was already correct even though driver.get() timed out). Callers
    MUST still WebDriverWait for the element they need after calling this -- never assume
    readiness."""
    try:
        driver.get(url)
    except TimeoutException:
        print(f"  -> (info) Page load timeout opening {url}, continuing (DOM is usually ready).")


def cetak_pdf_langsung(path_pdf, nama_printer):
    """Print a PDF file to a SPECIFIC printer explicitly via SumatraPDF (-print-to -silent),
    bypassing Windows' "default printer" mechanism entirely. Two other approaches were tried
    and failed: (1) SetDefaultPrinter() + Chrome kiosk-printing -- Chrome kiosk-printing was
    confirmed to cache the default-printer resolution at the browser-process level, so
    changing the OS default printer mid-session wasn't always picked up (invoices silently
    printed to the OLD printer even after the user picked a different one); (2) the shell
    'printto' verb via the default PDF handler -- that handler itself hung/failed silent
    print on this machine (likely a hidden dialog requiring manual interaction, e.g. an
    activation/trial screen). SumatraPDF portable is used because it's purpose-built for
    silent command-line printing and was confirmed to work cleanly (returncode 0, no dialogs)
    in direct testing."""
    hasil = subprocess.run(
        [SUMATRA_EXE, "-print-to", nama_printer, "-silent", path_pdf],
        capture_output=True, text=True, timeout=30,
    )
    if hasil.returncode != 0:
        raise RuntimeError(
            f"SumatraPDF failed to print '{path_pdf}' to printer '{nama_printer}' "
            f"(returncode {hasil.returncode}): {hasil.stderr or hasil.stdout}"
        )


def jumlah_job_di_antrian(nama_printer):
    """Count how many jobs are still in the target printer's spooler queue (not yet printed)."""
    handle = win32print.OpenPrinter(nama_printer)
    try:
        jobs = win32print.EnumJobs(handle, 0, 999, 1)
        return len(jobs)
    finally:
        win32print.ClosePrinter(handle)


def tunggu_print_benar_benar_selesai(nama_printer, timeout=30, lapor=_noop):
    """Wait until the target printer's spooler queue is EMPTY, meaning the print job has
    truly been printed/processed -- not just assumed done via a fixed time.sleep. Raises
    TimeoutError if the job is still stuck in the queue after `timeout` seconds (e.g. printer
    offline/jammed/out of paper), so this DN is marked FAILED rather than assumed successful
    when the paper never actually came out."""
    mulai = time.time()
    time.sleep(1.5)  # give the job a moment to actually enter the queue before checking
    while time.time() - mulai < timeout:
        sisa = jumlah_job_di_antrian(nama_printer)
        if sisa == 0:
            return True
        lapor(f"Waiting for the printer to finish... ({sisa} job(s) in queue)")
        time.sleep(2)
    raise TimeoutError(
        f"Print job still stuck in queue for printer '{nama_printer}' after {timeout}s "
        f"(check printer: offline / out of paper / jammed)."
    )


def _dengan_retry_renderer(func, percobaan=3, jeda=2, lapor=_noop):
    """Call func() with retries on TimeoutException 'Timed out receiving message from
    renderer' -- a transient error specific to this portal that was confirmed (through
    isolated testing) to appear on ANY WebDriver command (not just driver.get(), also
    WebDriverWait/find_element/click), even though the DOM/page is actually already ready
    when checked again a moment later (title & elements are found immediately with no
    problem). _get_dengan_timeout() only catches this for driver.get() specifically; this
    helper wraps a wider range of commands so transient errors anywhere are retried instead
    of failing outright."""
    for percobaan_ke in range(1, percobaan + 1):
        try:
            return func()
        except TimeoutException as e:
            if "receiving message from renderer" not in str(e) or percobaan_ke == percobaan:
                raise
            pesan = f"Hit a transient renderer timeout (attempt {percobaan_ke}/{percobaan}), retrying..."
            print(f"  -> (info) {pesan}")
            lapor(pesan)
            time.sleep(jeda)


def login_ces(driver, username, password, lapor=_noop):
    def _coba():
        _get_dengan_timeout(driver, URL_LOGIN)
        tunggu = WebDriverWait(driver, 20)

        kolom_user = tunggu.until(EC.presence_of_element_located(SELECTOR["input_username"]))
        kolom_user.send_keys(username)

        kolom_pass = driver.find_element(*SELECTOR["input_password"])
        kolom_pass.send_keys(password)

        driver.find_element(*SELECTOR["tombol_login"]).click()

    _dengan_retry_renderer(_coba, lapor=lapor)
    print("Login submitted, waiting for the home page...")
    time.sleep(3)


def buka_jurnal_piutang(driver, lapor=_noop):
    def _coba():
        _get_dengan_timeout(driver, URL_JURNAL_PIUTANG)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(SELECTOR["input_no_dn"])
        )

    _dengan_retry_renderer(_coba, lapor=lapor)
    print("Receivables Journal page opened.")
    time.sleep(1)


def cari_no_dn(driver, no_dn):
    tunggu = WebDriverWait(driver, 15)
    kolom_no_dn = tunggu.until(EC.presence_of_element_located(SELECTOR["input_no_dn"]))
    kolom_no_dn.clear()
    kolom_no_dn.send_keys(no_dn)

    tombol_cari = driver.find_element(*SELECTOR["tombol_cari_kaca_pembesar"])
    tombol_cari.click()
    print(f"  -> Searching: {no_dn}")

    # Wait until the table is ACTUALLY filtered to this DN number -- previously this was
    # just a fixed time.sleep(2), which did NOT verify the table's contents. On a slow
    # request, the DataTables AJAX filter might not have finished by the time
    # baca_baris_hasil() runs, so the robot would read the OLD (unfiltered) table -- confirmed
    # in testing: data for a completely different DN number (status/entity/amount) was read
    # incorrectly due to this race condition, not because the wrong DN was searched.
    from selenium.common.exceptions import StaleElementReferenceException

    def _tabel_sudah_terfilter(_):
        try:
            baris = driver.find_elements(*SELECTOR["baris_hasil"])
            if not baris:
                return False
            tds = baris[0].find_elements(By.TAG_NAME, "td")
            return len(tds) > 6 and no_dn in tds[6].text
        except StaleElementReferenceException:
            return False

    try:
        WebDriverWait(driver, 15).until(_tabel_sudah_terfilter)
    except TimeoutException:
        # Diagnostics: record what's actually in the first row when this times out, so it's
        # clear whether this is a genuinely slow AJAX request (a different DN in the first
        # row, table probably not filtered yet) vs. a different bug (e.g. an empty table /
        # "No matching records found").
        try:
            baris = driver.find_elements(*SELECTOR["baris_hasil"])
            tds = baris[0].find_elements(By.TAG_NAME, "td") if baris else []
            info_baris = [td.text for td in tds] if tds else "(empty table / no rows)"
        except Exception as e_diag:
            info_baris = f"(failed to read row for diagnostics: {e_diag})"
        print(f"  -> (diagnostic) Table not yet filtered to {no_dn} after 15s. Current first row: {info_baris}")
        raise


def klik_konfirmasi(driver, teks_tombol, tunggu_detik=10):
    """Click a button on a confirmation/success dialog by its text (e.g. 'Yes' / 'Ok!')."""
    tunggu = WebDriverWait(driver, tunggu_detik)
    tombol = tunggu.until(EC.element_to_be_clickable(
        (By.XPATH, f"//button[normalize-space()='{teks_tombol}']")
    ))
    tombol.click()
    time.sleep(1)


def baca_baris_hasil(driver, percobaan=3):
    """Read the status, entity text, and amount from the first row of search results.
    Retries a few times on stale elements (table just redrew via AJAX)."""
    from selenium.common.exceptions import StaleElementReferenceException

    tunggu = WebDriverWait(driver, 15)
    for percobaan_ke in range(percobaan):
        try:
            baris = tunggu.until(EC.presence_of_element_located(SELECTOR["baris_hasil"]))
            tds = baris.find_elements(By.TAG_NAME, "td")
            status = tds[9].text.strip()
            eselon = tds[7].text.strip()
            jumlah_teks = tds[8].text.strip()
            jumlah = _parse_nominal(jumlah_teks)
            return status, eselon, jumlah
        except StaleElementReferenceException:
            time.sleep(0.5)
    raise StaleElementReferenceException(
        "Result table row kept changing (stale) after several read attempts."
    )


def verifikasi_jurnal_jika_perlu(driver, lapor=_noop):
    """If status is still 'Not Verified', click the verify (thumbs-up) icon and confirm Yes."""
    status, _, _ = baca_baris_hasil(driver)
    if status != "Belum Ver":
        print("  -> Already verified, skipping.")
        return

    lapor("Verifying receivables journal entry...")
    tombol_verifikasi = driver.find_element(*SELECTOR["tombol_verifikasi"])
    tombol_verifikasi.click()
    klik_konfirmasi(driver, "Ya")
    klik_konfirmasi(driver, "Ok!")

    # Wait until the status ACTUALLY becomes "Sudah Ver" in the table before moving on to the
    # printing stage, so we don't proceed with a row whose data hasn't updated post-AJAX-redraw
    from selenium.common.exceptions import StaleElementReferenceException

    def _sudah_ter_verifikasi(_):
        try:
            status_terkini, _, _ = baca_baris_hasil(driver)
            return status_terkini == "Sudah Ver"
        except StaleElementReferenceException:
            return False

    WebDriverWait(driver, 15).until(_sudah_ter_verifikasi)
    print("  -> Receivables journal entry verified (journal posting succeeded).")


def unduh_pdf_tab_aktif(driver, handle_awal, path_tujuan):
    """Fetch the PDF currently open in the active tab (Chrome's native PDF viewer) via a
    JavaScript fetch() -- automatically carries the logged-in session cookie -- then save it
    to disk. Avoids the fragile, version-dependent shadow-DOM download button in the viewer.
    Closes the active tab and returns to the original tab afterward."""
    url_pdf = driver.current_url
    time.sleep(1)  # give the PDF a moment to finish loading before fetching

    base64_data = driver.execute_async_script(
        """
        const url = arguments[0];
        const callback = arguments[arguments.length - 1];
        fetch(url).then(r => r.blob()).then(blob => {
            const reader = new FileReader();
            reader.onloadend = () => callback(reader.result.split(',')[1]);
            reader.onerror = () => callback('ERROR:' + reader.error);
            reader.readAsDataURL(blob);
        }).catch(err => callback('ERROR:' + err));
        """,
        url_pdf,
    )

    if base64_data is None or str(base64_data).startswith("ERROR:"):
        driver.close()
        driver.switch_to.window(handle_awal)
        raise RuntimeError(f"Failed to fetch PDF from {url_pdf}: {base64_data}")

    os.makedirs(os.path.dirname(path_tujuan), exist_ok=True)
    with open(path_tujuan, "wb") as f:
        f.write(base64.b64decode(base64_data))

    driver.close()
    driver.switch_to.window(handle_awal)
    print(f"  -> PDF saved: {path_tujuan}")


def unduh_invoice_digital(driver, no_dn, path_tujuan):
    """Priority-client & amount below the signature threshold only: click the signature-
    required icon, BUT download the PDF (instead of physical printing) -- used for fully
    digital invoices that later get merged with the Debet Nota recap."""
    driver.find_element(*SELECTOR["tombol_cetak_merah"]).click()
    klik_konfirmasi(driver, "Ya")

    tunggu = WebDriverWait(driver, 10)
    input_tanggal = tunggu.until(EC.presence_of_element_located(SELECTOR["input_tgl_invoice"]))
    tanggal_hari_ini = date.today().strftime("%Y-%m-%d")
    driver.execute_script("arguments[0].value = arguments[1];", input_tanggal, tanggal_hari_ini)

    handle_awal = driver.current_window_handle
    driver.find_element(*SELECTOR["tombol_print_invoice"]).click()
    time.sleep(2)
    tunggu.until(EC.number_of_windows_to_be(2))
    handle_invoice = [h for h in driver.window_handles if h != handle_awal][0]
    driver.switch_to.window(handle_invoice)
    time.sleep(2)

    unduh_pdf_tab_aktif(driver, handle_awal, path_tujuan)
    print(f"  -> Digital invoice {no_dn} downloaded (priority client, below signature threshold).")


def unduh_rekap_debet_nota(driver, no_dn, path_tujuan):
    """Open the Debet Nota page (separate from the Receivables Journal), search the same DN
    number, download its Debet Nota recap. Navigates back to the Receivables Journal at the
    end so the main loop in jalankan_otomasi() isn't broken (that loop assumes the driver
    stays on the Receivables Journal page)."""
    _get_dengan_timeout(driver, URL_DEBET_NOTA)
    tunggu = WebDriverWait(driver, 20)
    tunggu.until(EC.presence_of_element_located(SELECTOR["debetnota_input_periode"]))

    periode = turunkan_periode_dari_no_dn(no_dn)
    kolom_periode = driver.find_element(*SELECTOR["debetnota_input_periode"])
    driver.execute_script("arguments[0].value = arguments[1];", kolom_periode, periode)

    kolom_no_dn = driver.find_element(*SELECTOR["debetnota_input_no_dn"])
    kolom_no_dn.clear()
    kolom_no_dn.send_keys(no_dn)

    driver.find_element(*SELECTOR["debetnota_tombol_cari"]).click()
    time.sleep(2)

    tunggu.until(EC.presence_of_element_located(SELECTOR["debetnota_baris_hasil"]))
    driver.find_element(*SELECTOR["debetnota_tombol_cetak"]).click()
    klik_konfirmasi(driver, "Ya")

    handle_awal = driver.current_window_handle
    tunggu.until(EC.number_of_windows_to_be(2))
    handle_rekap = [h for h in driver.window_handles if h != handle_awal][0]
    driver.switch_to.window(handle_rekap)
    time.sleep(2)

    unduh_pdf_tab_aktif(driver, handle_awal, path_tujuan)
    print(f"  -> Debet Nota recap {no_dn} downloaded.")

    buka_jurnal_piutang(driver)


def _dari_thread_lain(root, fungsi_dialog):
    """Run fungsi_dialog() (which shows a Tkinter widget) on the MAIN thread via root.after()
    -- Tkinter is not thread-safe, every dialog must go through the main thread that runs
    root.mainloop(), not be called directly from a Selenium/background thread. This function
    waits (blocking, from the calling thread's side) until the dialog completes and returns
    its result -- same pattern as self.root.after(0, ...) already used in RobotCESGui, with a
    threading.Event added so the result value can be awaited."""
    if root is None:
        return fungsi_dialog()  # fallback if called without a GUI (e.g. manual console testing)

    hasil = {}
    selesai = threading.Event()

    def _jalankan():
        hasil["nilai"] = fungsi_dialog()
        selesai.set()

    root.after(0, _jalankan)
    selesai.wait()
    return hasil["nilai"]


def pilih_printer_manual(pilihan_terakhir=None, root=None):
    """Show a Tkinter dialog to manually choose a printer (rather than hardcoding one),
    since the operator doesn't always use the same printer -- if one has issues, they switch
    to another. The last choice (if any) is pre-selected."""
    daftar_printer = [p[2] for p in win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )]
    if not daftar_printer:
        raise RuntimeError("No printers detected on this system.")

    def _tampilkan():
        hasil = {"printer": None}
        dialog = tk.Toplevel(root)
        dialog.title("Choose Printer")
        dialog.geometry("360x140")
        dialog.attributes("-topmost", True)

        tk.Label(dialog, text="Choose a printer for this invoice:").pack(pady=(15, 5))

        default_awal = pilihan_terakhir if pilihan_terakhir in daftar_printer else daftar_printer[0]
        var_printer = tk.StringVar(value=default_awal)
        dropdown = tk.OptionMenu(dialog, var_printer, *daftar_printer)
        dropdown.pack(pady=5, padx=20, fill="x")

        def _konfirmasi():
            hasil["printer"] = var_printer.get()
            dialog.destroy()

        tk.Button(dialog, text="Print to this printer", command=_konfirmasi).pack(pady=10)
        dialog.grab_set()
        dialog.wait_window()
        return hasil["printer"]

    return _dari_thread_lain(root, _tampilkan)


def konfirmasi_print_selesai(no_dn, root=None):
    """Human-in-the-loop gate: ask for manual confirmation that the invoice REALLY printed
    (an empty spooler queue doesn't always mean the physical result is correct -- paper can
    jam/come out blank/run out of ink even after the job has left the queue). The robot must
    not move on to the next DN number before this is confirmed by a human."""
    return _dari_thread_lain(root, lambda: messagebox.askyesno(
        "Confirm Print",
        f"Did invoice {no_dn} print successfully?",
    ))


def cetak_invoice(driver, no_dn, lapor=_noop, pilihan_printer_sebelumnya=None, root=None):
    """Decide which print icon to use (no-signature / signature) based on the priority-client
    and amount rules.
    - Priority client & amount below threshold: fully digital (download invoice+recap, merge,
      convert to Excel, drop straight into Siap_Input) -- see pipeline_priority_client.proses_priority_kecil().
    - Priority client & amount at/above threshold: prepare the Debet Nota recap (download+
      convert to Excel) first, then CONTINUE to physical printing as usual (needs a stamp +
      signature + manual scan by the operator).
    - Non-priority client: physical printing as usual, unchanged.
    Returns: the printer name used (to pre-select for the next DN), or None if no physical
    print happened (the priority-client-below-threshold case, fully digital)."""
    import pipeline_priority_client

    _, eselon, jumlah = baca_baris_hasil(driver)
    is_priority_client = eselon_termasuk_priority_client(eselon)
    gunakan_ttd = is_priority_client or jumlah < BATAS_NOMINAL_TTD

    if is_priority_client and jumlah < BATAS_NOMINAL_TTD:
        lapor(f"Priority client below Rp{BATAS_NOMINAL_TTD:,.0f}: fully digital processing (download+merge+convert)...")
        pipeline_priority_client.proses_priority_kecil(
            driver, no_dn, unduh_invoice_digital, unduh_rekap_debet_nota, lapor=lapor
        )
        return None  # no physical print for this case

    if is_priority_client:
        lapor(f"Priority client >= Rp{BATAS_NOMINAL_TTD:,.0f}: preparing Debet Nota recap (Excel) first...")
        pipeline_priority_client.proses_priority_besar(driver, no_dn, unduh_rekap_debet_nota, lapor=lapor)
        # unduh_rekap_debet_nota() visits the Debet Nota page then returns to the Receivables
        # Journal page with the table NOT re-filtered (default ~100k+ row table) -- this DN
        # must be searched again, otherwise find_element below would click the FIRST row of
        # the default table (wrong invoice printed, with no error).
        cari_no_dn(driver, no_dn)
        # NOT returning here -- continues to physical printing below as usual

    selector_tombol = SELECTOR["tombol_cetak_merah"] if gunakan_ttd else SELECTOR["tombol_cetak_ungu"]
    alasan = "priority client" if is_priority_client else f"amount Rp{jumlah:,.0f}"
    print(f"  -> {no_dn}: using {'signature-required' if gunakan_ttd else 'no-signature'} icon ({alasan})")
    lapor(f"Printing invoice ({'with signature' if gunakan_ttd else 'no signature'})...")

    driver.find_element(*selector_tombol).click()
    klik_konfirmasi(driver, "Ya")

    tunggu = WebDriverWait(driver, 10)
    input_tanggal = tunggu.until(EC.presence_of_element_located(SELECTOR["input_tgl_invoice"]))
    tanggal_hari_ini = date.today().strftime("%Y-%m-%d")
    driver.execute_script("arguments[0].value = arguments[1];", input_tanggal, tanggal_hari_ini)

    handle_awal = driver.current_window_handle
    driver.find_element(*SELECTOR["tombol_print_invoice"]).click()
    time.sleep(2)

    tunggu.until(EC.number_of_windows_to_be(2))
    handle_invoice = [h for h in driver.window_handles if h != handle_awal][0]
    driver.switch_to.window(handle_invoice)
    time.sleep(2)  # wait for the PDF to finish rendering before downloading

    # Download first (rather than window.print() to the default printer) -- see
    # cetak_pdf_langsung() for why Chrome kiosk-printing was abandoned.
    # unduh_pdf_tab_aktif() already handles driver.close() + switching back to handle_awal.
    os.makedirs(FOLDER_CETAK_SEMENTARA, exist_ok=True)
    path_pdf_cetak = os.path.join(FOLDER_CETAK_SEMENTARA, f"{no_dn}_invoice.pdf")
    unduh_pdf_tab_aktif(driver, handle_awal, path_pdf_cetak)

    printer_dipilih = pilih_printer_manual(pilihan_terakhir=pilihan_printer_sebelumnya, root=root)
    lapor(f"Printing to {printer_dipilih}...")
    cetak_pdf_langsung(path_pdf_cetak, printer_dipilih)
    tunggu_print_benar_benar_selesai(printer_dipilih, timeout=30, lapor=lapor)

    # An empty queue is not a guarantee the physical print is correct -- ask for a human
    # confirmation as the final gate before this DN is considered truly done.
    while not konfirmasi_print_selesai(no_dn, root=root):
        cetak_ulang = _dari_thread_lain(root, lambda: messagebox.askyesno(
            "Print Not Confirmed",
            f"Retry printing invoice {no_dn} to a different printer?\n\n"
            "Yes = choose a printer again and reprint\nNo = mark as failed, handle manually later",
        ))
        if not cetak_ulang:
            raise RuntimeError(f"Printing invoice {no_dn} was confirmed as failed by the user.")

        printer_dipilih = pilih_printer_manual(pilihan_terakhir=printer_dipilih, root=root)
        lapor(f"Reprinting to {printer_dipilih}...")
        cetak_pdf_langsung(path_pdf_cetak, printer_dipilih)
        tunggu_print_benar_benar_selesai(printer_dipilih, timeout=30, lapor=lapor)

    os.remove(path_pdf_cetak)
    print(f"  -> Invoice {no_dn} CONFIRMED printed to '{printer_dipilih}'.")
    lapor(f"Invoice {no_dn} finished printing (confirmed).")
    return printer_dipilih


# ============================================================
# MAIN (called from a background thread by the GUI)
# ============================================================
def jalankan_otomasi(lapor=_noop, lapor_progress=_noop, lapor_file=_noop, stop_event=None, root=None):
    """lapor: status text for the current stage.
    lapor_progress: "which DN number out of the list" text (e.g. '3 / 12').
    lapor_file: name of the trigger PDF currently being processed.
    stop_event: threading.Event() - if set, the robot stops AFTER the currently running DN
    number finishes (never cuts off mid verify/print, which would be risky).
    root: the tk.Tk() GUI instance -- passed to the printer-choice/print-confirmation dialogs
    so they're shown via the main thread (Tkinter is not thread-safe)."""
    file_pdf_list = [
        f for f in os.listdir(FOLDER_SIAP) if f.lower().endswith(".pdf")
    ]
    if not file_pdf_list:
        lapor("No PDF files in Siap_Input.")
        return

    log = muat_log_progress()
    printer_terakhir = PRINTER_SARAN_AWAL

    lapor("Opening browser & logging into the portal...")
    driver = buka_browser()
    login_ces(driver, PORTAL_USERNAME, PORTAL_PASSWORD, lapor=lapor)
    buka_jurnal_piutang(driver, lapor=lapor)

    for nama_file in file_pdf_list:
        path_pdf = os.path.join(FOLDER_SIAP, nama_file)
        print(f"\nProcessing: {nama_file}")
        lapor_file(nama_file)

        daftar_no_dn = ekstrak_no_dn_dari_pdf(path_pdf)
        total = len(daftar_no_dn)
        print(f"  Found {total} DN number(s).")
        simpan_ke_excel(daftar_no_dn, nama_file)

        for i, no_dn in enumerate(daftar_no_dn, start=1):
            if stop_event is not None and stop_event.is_set():
                lapor(f"STOPPED by user before processing {no_dn}.")
                print(f"  -> Stopped by user before processing {no_dn}.")
                break

            lapor_progress(f"DN {i} / {total}  -  {no_dn}")

            if sudah_selesai(log, no_dn):
                print(f"  -> {no_dn} already succeeded previously, skipping.")
                lapor(f"{no_dn} already done previously, skipping.")
                continue

            lapor(f"Searching DN {no_dn}...")
            try:
                cari_no_dn(driver, no_dn)
                verifikasi_jurnal_jika_perlu(driver, lapor=lapor)
                printer_dipakai = cetak_invoice(
                    driver, no_dn, lapor=lapor, pilihan_printer_sebelumnya=printer_terakhir, root=root
                )
                if printer_dipakai:
                    printer_terakhir = printer_dipakai
                tandai_selesai(log, no_dn, nama_file)
            except Exception as e:
                print(f"  -> FAILED processing {no_dn}: {e}")
                lapor(f"FAILED processing {no_dn}: {e}")
                tandai_gagal(log, no_dn, nama_file, e)

        if stop_event is not None and stop_event.is_set():
            lapor(f"{nama_file}: stopped at user request, remainder not processed.")
            break

        semua_tuntas = all(sudah_selesai(log, dn) for dn in daftar_no_dn)
        if semua_tuntas:
            shutil.move(path_pdf, os.path.join(FOLDER_SELESAI, nama_file))
            print("  -> All DN numbers complete, file moved to Selesai_Input.")
        else:
            sisa = [dn for dn in daftar_no_dn if not sudah_selesai(log, dn)]
            print(f"  -> {len(sisa)} DN number(s) still incomplete, file STAYS in Siap_Input for retry: {sisa}")
            lapor(f"{nama_file}: {len(sisa)} DN number(s) incomplete, file stays in Siap_Input.")

    lapor("Done. Browser left open for manual review.")
    lapor_progress("All DN numbers processed.")
    print("\nDone. Browser left open for manual review.")


# ============================================================
# GUI - J.A.R.V.I.S.-style HUD (consistent with the claims-portal automation)
# ============================================================
class RobotCESGui:
    def __init__(self, root):
        self.root = root
        self.root.title("Robot CES - Receivables Journal [V1]")

        window_width, window_height = 420, 260
        screen_width = self.root.winfo_screenwidth()
        x_coordinate = screen_width - window_width - 20
        y_coordinate = 20
        self.root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
        self.root.attributes("-topmost", True)

        self.bg_color = "#0B0C10"
        self.fg_neon = "#66FCF1"
        self.fg_accent = "#45A29E"
        self.fg_text = "#C5C6C7"
        self.fg_warn = "#F2A900"
        self.fg_alert = "#FF4C4C"

        self.root.configure(bg=self.bg_color)
        self.stop_event = threading.Event()

        font_header = ("Consolas", 11, "bold")
        font_title = ("Consolas", 10, "bold")
        font_text = ("Consolas", 9)

        tk.Label(root, text="ROBOT CES - RECEIVABLES JOURNAL\nWELCOME BACK, OPERATOR.",
                 font=font_header, bg=self.bg_color, fg=self.fg_neon).pack(pady=(15, 5))

        self.lbl_file = tk.Label(root, text="FILE TRIGGER: -",
                                  font=font_text, bg=self.bg_color, fg=self.fg_accent)
        self.lbl_file.pack(pady=(5, 2))

        self.lbl_progress = tk.Label(root, text="PROGRESS: -",
                                      font=font_title, bg=self.bg_color, fg=self.fg_warn)
        self.lbl_progress.pack(pady=2)

        self.lbl_status = tk.Label(root, text="SYSTEM STANDBY: Awaiting command.",
                                    font=font_text, bg=self.bg_color, fg=self.fg_text, wraplength=380)
        self.lbl_status.pack(pady=10)

        self.btn_mulai = tk.Button(root, text="[ INITIALIZE AUTOMATION ]", font=font_title,
                                    bg="#1F2833", fg=self.fg_neon, activebackground=self.fg_neon,
                                    activeforeground="#1F2833", relief="ridge", bd=2, cursor="hand2",
                                    command=self.mulai_proses)
        self.btn_mulai.pack(pady=8, fill="x", padx=40, ipady=3)

        self.btn_stop = tk.Button(root, text="[ STOP AFTER THIS DN ]", font=font_title,
                                   bg="#1F2833", fg=self.fg_alert, activebackground=self.fg_alert,
                                   activeforeground="#1F2833", relief="ridge", bd=2, cursor="hand2",
                                   state="disabled", command=self.minta_stop)
        self.btn_stop.pack(pady=5, fill="x", padx=40, ipady=3)

    def _update_label(self, label, teks):
        self.root.after(0, lambda: label.config(text=teks))

    def lapor_status(self, teks):
        self._update_label(self.lbl_status, f"STATUS: {teks}")

    def lapor_progress(self, teks):
        self._update_label(self.lbl_progress, f"PROGRESS: {teks}")

    def lapor_file(self, teks):
        self._update_label(self.lbl_file, f"FILE TRIGGER: {teks}")

    def mulai_proses(self):
        self.btn_mulai.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.stop_event.clear()
        self.lapor_status("Starting automation...")

        def target():
            try:
                jalankan_otomasi(
                    lapor=self.lapor_status,
                    lapor_progress=self.lapor_progress,
                    lapor_file=self.lapor_file,
                    stop_event=self.stop_event,
                    root=self.root,
                )
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("SYSTEM ERROR", str(e)))
                self.lapor_status(f"ERROR: {e}")
            finally:
                self.root.after(0, lambda: self.btn_mulai.config(state="normal"))
                self.root.after(0, lambda: self.btn_stop.config(state="disabled"))

        threading.Thread(target=target, daemon=True).start()

    def minta_stop(self):
        self.stop_event.set()
        self.btn_stop.config(state="disabled")
        self.lapor_status("STOP requested, finishing the current DN first...")


if __name__ == "__main__":
    root = tk.Tk()
    app = RobotCESGui(root)
    root.mainloop()
