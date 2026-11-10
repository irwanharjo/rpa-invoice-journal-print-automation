# Changelog

## 2026-11 — Browser hardening

- **New `browser_hardening.py`**, shared by every robot that opens Chrome against the portal:
  - `suppress_chrome_popups()` — prevents Chrome's Password Manager "Change your password"
    (data-breach) dialog. That dialog lives outside the DOM, so Selenium cannot close it; it
    stole keyboard focus in the middle of `send_keys` and typed values silently went nowhere.
    It is now prevented through profile prefs at launch, merged safely with the robot's own
    options (idempotent, never overwrites an existing `--disable-features=` list).
  - `clear_page_blockers()` — closes native alerts, SweetAlert2 popups, Bootstrap modals and
    generic dialogs **once, at the entry point only**. The robot's own working modals are
    allow-listed, and everything closed is returned and logged — nothing is closed silently.
  - `detect_forced_password_change()` — a narrow, fail-open guard for the day the portal starts
    forcing password changes (HTML comments are stripped first to avoid false positives).
- `robot_ces.py` wired to both hooks (launch + post-login).

## 2026-07 — Initial public version

- Journal verification, invoice print with spooler polling and human confirmation, resumable
  progress log, priority-client digital pipeline.
