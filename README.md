# Receivables Journal & Invoice Print Automation

A resilient RPA tool that automates a receivables-journal verification and invoice-printing workflow on a corporate finance web portal, including physical printing to a specific printer and a human-in-the-loop confirmation gate. Genericized for portfolio purposes (portal URL, entity codes, and printer name have been replaced with placeholders/environment variables).

This is the piece of the portfolio I'm most proud of engineering-wise — the code comments document several real production bugs and their root causes (see below), not just what the code does.

## What it does

1. Extracts Debit Note (DN) numbers from incoming "receipt summary" PDFs (table extraction with a full-text regex fallback so nothing is missed), and writes them to an Excel safety-net file.
2. Logs into the portal and, for each DN number: searches for it, verifies the receivables journal entry if needed, and decides which print variant to use based on client type and invoice amount.
3. For a "priority client" flow, delegates to `pipeline_priority_client.py`, which chains together this project's downstream automations (PDF→Excel conversion, PDF merging) so a fully-digital invoice ends up ready for the next stage in the pipeline with no manual file handling.
4. Downloads the rendered invoice PDF (via a `fetch()` against the authenticated session, rather than clicking the viewer's download button) and prints it to a specific printer chosen by the operator, using SumatraPDF's silent CLI printing.
5. Polls the Windows print spooler until the job is confirmed gone from the queue, then asks the operator to manually confirm the physical print actually came out correctly — the last line of defense against silent no-ops (jammed paper, empty toner, wrong printer).
6. Tracks per-DN progress in `log_progress.json` so an interrupted run can resume safely without ever double-printing an invoice.

## Notable engineering details (see inline comments for the full story)

- **Renderer timeout retries**: this portal intermittently throws "Timed out receiving message from renderer" on arbitrary WebDriver calls even though the page is genuinely ready — handled with a targeted retry wrapper instead of blanket retries or longer sleeps.
- **AJAX race condition**: a naive fixed `sleep()` after searching would sometimes read the *previous* (unfiltered) table state, silently returning wrong data for a different DN. Fixed by polling until the table content actually matches the searched DN.
- **Printer selection reliability**: browser kiosk-printing was abandoned after discovering Chrome caches the OS default printer at the process level, so switching printers mid-session silently kept printing to the old one. Direct CLI printing via SumatraPDF, followed by spooler-queue polling, replaced it.

## Setup

```bash
pip install -r requirements.txt
copy .env.example .env   # then fill in your own values
```

Update `KODE_PRIORITY_CLIENT` in `robot_ces.py` and the folder paths in `pipeline_priority_client.py` to match your own portal/pipeline. This project also expects two sibling projects for the downstream digital-invoice pipeline (see `.env.example`): a PDF-to-Excel converter and a PDF merger, each with their own repo in this portfolio.

## Run

```bash
python robot_ces.py
```

## Stack

Python · Selenium · pdfplumber · openpyxl · pywin32 · Tkinter · SumatraPDF (CLI)

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for what changed after the first public version (production hardening notes).
