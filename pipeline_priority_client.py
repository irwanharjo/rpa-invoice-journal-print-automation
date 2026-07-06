"""
pipeline_priority_client.py - Orchestrator that connects robot_ces.py (download invoice & Debet Nota
recap from the portal) -> a PDF-to-Excel converter module -> a PDF-merge module -> the
downstream claims-automation input folder.

Special handling for the "priority client" entity:
- Priority client & amount below the signature threshold (proses_priority_kecil): fully digital,
  the robot merges invoice+recap itself.
- Priority client & amount at/above the threshold (proses_priority_besar): the robot only
  prepares the recap Excel. The scanned invoice PDF (already stamped+signed, physically
  merged with the recap) is dropped manually by the operator into the downstream input
  folder, named with the same DN number.

The downstream claims-automation script is NOT called from here -- that's still run manually
by the operator whenever they're ready. That's the final gate before data actually gets
submitted to the counterpart via the claims portal -- these scripts remain independent
modules; this file just wires the flow between them.
"""

import os
import shutil
import sys

import pdfplumber

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FOLDER_PRIORITY_STAGING = os.path.join(BASE_DIR, "PriorityClient_Staging")
FOLDER_PRIORITY_ARSIP = os.path.join(BASE_DIR, "PriorityClient_Arsip")
FOLDER_SIAP_APS = os.environ.get("DOWNSTREAM_CLAIMS_INPUT_FOLDER", os.path.join(BASE_DIR, "..", "rpa-claims-portal-automation", "Siap_Input"))

# Paths to sibling automation projects that provide the PDF->Excel and PDF-merge building
# blocks. Set these env vars (or edit the defaults) to point at your own copies.
KONVERSI_MODULE_PATH = os.environ.get("KONVERSI_MODULE_PATH", os.path.join(BASE_DIR, "..", "pdf-to-excel-converter"))
MERGE_MODULE_PATH = os.environ.get("MERGE_MODULE_PATH", os.path.join(BASE_DIR, "..", "auto-merge-pdf"))
sys.path.insert(0, KONVERSI_MODULE_PATH)
sys.path.insert(0, MERGE_MODULE_PATH)
import konversi        # noqa: E402  (from the pdf-to-excel-converter project -- not modified here)
import merge_invoice    # noqa: E402  (from the auto-merge-pdf project -- not modified here)


def _noop(_teks):
    pass


def konversi_rekap_ke_excel(path_pdf_rekap, path_excel_keluar):
    """Convert one Debet Nota recap PDF to Excel, reusing konversi.py's low-level functions
    directly (instead of jalankan_konversi(), which scans the Input_PDF/Hasil_Excel folders)
    -- here a custom path is needed because the output must go straight to the downstream
    input folder."""
    with pdfplumber.open(path_pdf_rekap) as pdf:
        halaman_1 = pdf.pages[0]
        teks_hal_1 = halaman_1.extract_text() or ""
        header = konversi.ekstrak_header(halaman_1, teks_hal_1)
        tabel_mentah = konversi.ekstrak_tabel_mentah(pdf)
        data_final = konversi.jahit_baris_terpotong(tabel_mentah)

    if not data_final:
        raise RuntimeError(f"No table structure found in {path_pdf_rekap}")

    os.makedirs(os.path.dirname(path_excel_keluar), exist_ok=True)
    konversi.simpan_ke_excel(header, data_final, path_excel_keluar)
    return path_excel_keluar


def _pindah_ke_arsip(path_file):
    if os.path.exists(path_file):
        os.makedirs(FOLDER_PRIORITY_ARSIP, exist_ok=True)
        tujuan = os.path.join(FOLDER_PRIORITY_ARSIP, os.path.basename(path_file))
        if os.path.exists(tujuan):
            os.remove(tujuan)
        shutil.move(path_file, tujuan)


def proses_priority_kecil(driver, no_dn, unduh_invoice_fn, unduh_rekap_fn, lapor=_noop):
    """Priority client & amount below the signature threshold: fully digital -- download the
    invoice (signature icon) + download the Debet Nota recap, convert the recap to Excel,
    merge invoice+recap into a single PDF, and place {no_dn}.pdf + {no_dn}.xlsx directly into
    the downstream input folder.

    unduh_invoice_fn / unduh_rekap_fn: functions from robot_ces.py (unduh_invoice_digital /
    unduh_rekap_debet_nota), passed in as parameters so this module doesn't need to import
    robot_ces.py back (avoiding a circular/duplicate import when robot_ces.py is run directly
    as the main script)."""
    os.makedirs(FOLDER_PRIORITY_STAGING, exist_ok=True)
    os.makedirs(FOLDER_SIAP_APS, exist_ok=True)

    path_invoice = os.path.join(FOLDER_PRIORITY_STAGING, f"{no_dn}.pdf")
    path_rekap = os.path.join(FOLDER_PRIORITY_STAGING, f"RekapDebetNota_{no_dn}.pdf")
    path_excel_akhir = os.path.join(FOLDER_SIAP_APS, f"{no_dn}.xlsx")
    path_pdf_akhir = os.path.join(FOLDER_SIAP_APS, f"{no_dn}.pdf")

    if os.path.exists(path_excel_akhir) and os.path.exists(path_pdf_akhir):
        print(f"  -> {no_dn} already complete in the downstream input folder, skipping.")
        lapor(f"{no_dn} already processed previously, skipping.")
        return

    if not os.path.exists(path_invoice):
        lapor(f"Downloading digital invoice {no_dn}...")
        unduh_invoice_fn(driver, no_dn, path_invoice)

    if not os.path.exists(path_rekap):
        lapor(f"Downloading Debet Nota recap {no_dn}...")
        unduh_rekap_fn(driver, no_dn, path_rekap)

    if not os.path.exists(path_excel_akhir):
        lapor(f"Converting Debet Nota recap {no_dn} to Excel...")
        konversi_rekap_ke_excel(path_rekap, path_excel_akhir)
        print(f"  -> Excel saved: {path_excel_akhir}")

    if not os.path.exists(path_pdf_akhir):
        lapor(f"Merging invoice + recap {no_dn}...")
        merge_invoice.merge_invoice_and_rekap(FOLDER_PRIORITY_STAGING, FOLDER_SIAP_APS)
        if not os.path.exists(path_pdf_akhir):
            raise RuntimeError(f"Failed to merge invoice+recap for {no_dn} (check the file pair in {FOLDER_PRIORITY_STAGING}).")
        print(f"  -> Merged PDF saved: {path_pdf_akhir}")

    _pindah_ke_arsip(path_invoice)
    _pindah_ke_arsip(path_rekap)

    print(f"  -> {no_dn} (priority client, below threshold) READY in the downstream input folder.")
    lapor(f"{no_dn} ready in the downstream folder (priority client, fully digital).")


def proses_priority_besar(driver, no_dn, unduh_rekap_fn, lapor=_noop):
    """Priority client & amount at/above the signature threshold: the robot only prepares the
    Debet Nota recap (download+convert to Excel), placing {no_dn}.xlsx directly into the
    downstream input folder. The invoice PDF is NOT touched by the robot -- the operator will
    apply a stamp+signature+manual scan, then drop {no_dn}.pdf (already physically merged
    with the recap) into the downstream folder themselves. Once the pair is complete, this DN
    is ready for the downstream claims-automation step."""
    os.makedirs(FOLDER_PRIORITY_STAGING, exist_ok=True)
    os.makedirs(FOLDER_SIAP_APS, exist_ok=True)

    path_rekap = os.path.join(FOLDER_PRIORITY_STAGING, f"RekapDebetNota_{no_dn}.pdf")
    path_excel_akhir = os.path.join(FOLDER_SIAP_APS, f"{no_dn}.xlsx")

    if os.path.exists(path_excel_akhir):
        print(f"  -> Excel for {no_dn} already exists in the downstream folder, skipping.")
        lapor(f"Excel for {no_dn} already processed previously, skipping.")
        return

    if not os.path.exists(path_rekap):
        lapor(f"Downloading Debet Nota recap {no_dn}...")
        unduh_rekap_fn(driver, no_dn, path_rekap)

    lapor(f"Converting Debet Nota recap {no_dn} to Excel...")
    konversi_rekap_ke_excel(path_rekap, path_excel_akhir)
    print(f"  -> Excel saved: {path_excel_akhir}")

    _pindah_ke_arsip(path_rekap)

    print(f"  -> Excel for {no_dn} (priority client, at/above threshold) ready in the downstream folder. "
          f"Waiting for the scanned (stamped+signed) invoice named {no_dn}.pdf.")
    lapor(f"Excel for {no_dn} ready, waiting for the operator's scanned invoice.")
