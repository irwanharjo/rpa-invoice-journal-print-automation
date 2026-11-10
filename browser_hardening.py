# -*- coding: utf-8 -*-
"""
Browser hardening for Selenium robots: popup suppression + entry-point blocker cleanup.

One shared module for every robot that opens Chrome against the portal, so the fix
lives in one place instead of being copy-pasted in a different style into each robot.

THE ORIGINAL PROBLEM
--------------------
Every new session, Chrome showed a dialog:

    "Change your password - The password you just used was found in a data breach.
     Google Password Manager recommends changing your password now."

That dialog belongs to the BROWSER, not the page. It lives OUTSIDE the DOM, so no
`driver.find_element`, `Swal.close()` or `alert.accept()` can touch it -- yet it steals
keyboard focus in the middle of `send_keys`, so typed values silently land nowhere.
The only correct fix is to PREVENT it through Chrome profile preferences when the
browser is launched -> `suppress_chrome_popups()`.

What was verified before writing this (so nobody re-investigates without new data):
* The portal pages and their custom JS contain zero password modals/SweetAlerts.
* A real Selenium login monitored for 25 s after landing: no modal, no swal, no
  native alert -- i.e. the popup was 100% browser-side.
* A "change password" modal block DOES exist in the portal layout, but commented
  out. If the portal team ever un-comments it, it becomes a real DOM blocker --
  `clear_page_blockers()` and `detect_forced_password_change()` are ready for that.

USAGE
-----
    from browser_hardening import suppress_chrome_popups, clear_page_blockers

    opts = Options()
    ...                                  # robot-specific flags
    suppress_chrome_popups(opts)         # BEFORE webdriver.Chrome(options=opts)
    driver = webdriver.Chrome(options=opts)
    ...                                  # login
    clear_page_blockers(driver)          # ONLY at the entry point

GUARDRAIL: `clear_page_blockers()` is deliberately NOT called mid-workflow. The robot's
own working modals are listed in WORKFLOW_MODAL_ALLOWLIST and are never closed by it.
Lesson learned the hard way: an over-eager "defensive" check layered everywhere ends up
blocking work that was already correct.
"""
import logging
import re

__all__ = ["suppress_chrome_popups", "clear_page_blockers",
           "detect_forced_password_change", "WORKFLOW_MODAL_ALLOWLIST"]


# ============================================================
# 1. PREVENTION - the Chrome dialog never appears
# ============================================================
# `profile.password_manager_leak_detection` is THE key one: it is the equivalent of the
# chrome://settings/security toggle "Warn you if passwords are exposed in a data breach".
# The rest close the other password-popup paths ("Save password?", autofill) so nothing
# can steal keyboard focus mid `send_keys`.
SUPPRESS_PREFS = {
    "profile.password_manager_leak_detection": False,
    "profile.password_manager_enabled": False,
    "credentials_enable_service": False,
    "autofill.profile_enabled": False,
    "autofill.credit_card_enabled": False,
    # 2 = block: never show notification permission prompts
    "profile.default_content_setting_values.notifications": 2,
}

# Second line of defence. The prefs above do the real work; feature names a given
# Chrome version doesn't know are silently ignored (no error).
DISABLED_FEATURES = ("PasswordLeakDetection", "PasswordLeakToggleMove")

SUPPRESS_ARGS = (
    "--disable-notifications",
    "--no-first-run",
    "--no-default-browser-check",
    # don't touch the OS credential store
    "--password-store=basic",
)


def _merge_disable_features(opts, features):
    """Add `--disable-features=` WITHOUT overwriting the caller's own list.

    Chrome does not merge the switch itself when it is passed twice (the last one
    wins), so merging is done here.
    """
    existing, rest = set(), []
    for a in list(opts.arguments):
        if a.startswith("--disable-features="):
            existing.update(x for x in a.split("=", 1)[1].split(",") if x)
        else:
            rest.append(a)
    existing.update(features)
    merged = "--disable-features=" + ",".join(sorted(existing))
    try:
        # Selenium 4: `arguments` returns the underlying list.
        opts.arguments.clear()
        for a in rest:
            opts.arguments.append(a)
        opts.arguments.append(merged)
    except Exception:
        # If Selenium's internals change, never let a backup flag kill the robot.
        # The prefs are still applied.
        opts.add_argument(merged)


def suppress_chrome_popups(opts, extra_prefs=None):
    """Attach popup suppression to the caller's ChromeOptions.

    Safe on Options that ALREADY have prefs or a `--disable-features=` switch: both are
    MERGED, not overwritten. Idempotent (safe to call twice). Returns the same `opts`
    object so it can be chained.
    """
    prefs = dict(opts.experimental_options.get("prefs") or {})
    prefs.update(SUPPRESS_PREFS)
    if extra_prefs:
        prefs.update(extra_prefs)
    opts.add_experimental_option("prefs", prefs)

    present = set(opts.arguments)
    for a in SUPPRESS_ARGS:
        if a not in present:
            opts.add_argument(a)
    _merge_disable_features(opts, DISABLED_FEATURES)
    return opts


# ============================================================
# 2. CLEANUP - blockers that REALLY are in the DOM
# ============================================================
# The robot's own workflow modals. NEVER closed by the cleaner.
# Replace with the ids of your own portal's working modals.
WORKFLOW_MODAL_ALLOWLIST = (
    "addDetailModal",
    "createHeaderModal",
    "viewDetailModal",
)

_JS_CLEAN = r"""
const allow = new Set(arguments[0]);
const out = [];
const cut = (t) => (t || '').trim().replace(/\s+/g, ' ').slice(0, 200);

// --- SweetAlert2 -------------------------------------------------------
document.querySelectorAll('.swal2-container').forEach(sw => {
  const text = cut(sw.innerText);
  const btn = sw.querySelector('.swal2-confirm')
           || sw.querySelector('.swal2-close')
           || sw.querySelector('.swal2-cancel');
  if (btn) { btn.click(); } else { sw.remove(); }
  out.push({kind: 'swal2', text: text});
});

// --- Bootstrap modals currently SHOWN ----------------------------------
document.querySelectorAll('.modal').forEach(m => {
  const st = getComputedStyle(m);
  const shown = m.classList.contains('show') || st.display === 'block';
  if (!shown) return;
  if (m.id && allow.has(m.id)) { out.push({kind: 'skipped', id: m.id}); return; }
  const text = cut(m.innerText);
  try { if (window.jQuery) window.jQuery(m).modal('hide'); } catch (e) {}
  m.classList.remove('show');
  m.style.display = 'none';
  m.setAttribute('aria-hidden', 'true');
  out.push({kind: 'modal', id: m.id || '(no id)', text: text});
});

// Leftover backdrops are only removed when NO modal is shown any more --
// if an allow-listed working modal is still open, the backdrop belongs to it.
const stillShown = Array.from(document.querySelectorAll('.modal'))
  .some(m => m.classList.contains('show') || getComputedStyle(m).display === 'block');
if (!stillShown) {
  document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
  document.body.classList.remove('modal-open');
  document.body.style.removeProperty('padding-right');
  document.body.style.removeProperty('overflow');
}

// --- generic dialogs outside .modal ------------------------------------
document.querySelectorAll('[role=dialog], dialog[open]').forEach(d => {
  if (d.closest('.modal') || d.classList.contains('modal')) return;
  if (d.id && allow.has(d.id)) { out.push({kind: 'skipped', id: d.id}); return; }
  const st = getComputedStyle(d);
  if (st.display === 'none' || st.visibility === 'hidden') return;
  const text = cut(d.innerText);
  const close = d.querySelector('[aria-label=Close], .close, [data-dismiss=modal]');
  if (close) { close.click(); }
  else if (typeof d.close === 'function') { try { d.close(); } catch (e) { d.remove(); } }
  else { d.remove(); }
  out.push({kind: 'dialog', id: d.id || '(no id)', text: text});
});

return out;
"""


def clear_page_blockers(driver, allow=(), report=None):
    """Close blockers that ARE in the DOM, at the page entry point.

    Handles, in order: native alert/confirm -> SweetAlert2 -> shown Bootstrap modals ->
    generic `[role=dialog]` / `<dialog open>`.

    It CANNOT touch Chrome's own dialogs -- that is `suppress_chrome_popups()`'s job at
    browser launch.

    allow  : extra ids that must not be closed (merged with WORKFLOW_MODAL_ALLOWLIST)
    report : optional callable(str) to forward messages to the robot's UI

    Returns a LIST of everything that was closed -- deliberately returned and logged,
    so nothing is ever closed silently.
    """
    allowed = list(WORKFLOW_MODAL_ALLOWLIST) + list(allow)
    result = []

    # 1) native alert/confirm/prompt -- this FREEZES every following command if left
    #    open, so it goes first.
    try:
        a = driver.switch_to.alert
        text = a.text
        a.accept()
        result.append({"kind": "native-alert", "text": text})
    except Exception:
        pass

    # 2) everything else via the DOM
    try:
        result += (driver.execute_script(_JS_CLEAN, allowed) or [])
    except Exception as e:
        logging.warning("clear_page_blockers: JS failed (%s) - skipped, "
                        "robot keeps running.", type(e).__name__)

    closed = [r for r in result if r.get("kind") != "skipped"]
    if closed:
        msg = "Page popups closed: " + "; ".join(
            "{}[{}] {}".format(r["kind"], r.get("id", ""), r.get("text", "")).strip()
            for r in closed)
        logging.info(msg)
        print("  -> " + msg)
        if report:
            try:
                report(msg)
            except Exception:
                pass
    return result


# ============================================================
# 3. GUARD - if the portal one day FORCES a password change
# ============================================================
# Deliberately NARROW to avoid false positives: the password modal block exists on
# EVERY page but is commented out, so HTML comments are stripped before searching.
_RE_COMMENT = re.compile(r"<!--.*?-->", re.S)
_RE_PW_MODAL = re.compile(r"changePasswordModal|modalPassword", re.I)
_RE_PW_URL = re.compile(r"/change[-_]?password|/password/change", re.I)


def detect_forced_password_change(html="", url=""):
    """-> reason (str) if the portal is redirecting to a password change, else None.

    FAIL-OPEN: when in doubt, return None. The caller decides whether to stop (forced
    URL) or just warn (active modal).
    """
    if url and _RE_PW_URL.search(url):
        return "Portal redirected to a password-change page: " + url
    if html:
        clean = _RE_COMMENT.sub("", html)
        if _RE_PW_MODAL.search(clean):
            return "Portal password-change modal is now ACTIVE (no longer commented out)"
    return None


if __name__ == "__main__":
    # quick self-test of the option merging, without opening a browser
    from selenium.webdriver.chrome.options import Options
    o = Options()
    o.add_argument("--disable-blink-features=AutomationControlled")
    o.add_argument("--disable-features=AlreadyThere")
    o.add_experimental_option("prefs", {"download.default_directory": "/tmp/x"})
    suppress_chrome_popups(o)
    suppress_chrome_popups(o)                   # idempotent
    assert o.experimental_options["prefs"]["download.default_directory"] == "/tmp/x"
    assert o.experimental_options["prefs"]["profile.password_manager_leak_detection"] is False
    df = [a for a in o.arguments if a.startswith("--disable-features=")]
    assert len(df) == 1, df
    assert "AlreadyThere" in df[0] and "PasswordLeakDetection" in df[0]
    assert sum(1 for a in o.arguments if a == "--password-store=basic") == 1
    assert "--disable-blink-features=AutomationControlled" in o.arguments
    print("option merge OK")
    assert detect_forced_password_change("<html><!-- modalPassword --></html>", "") is None
    assert detect_forced_password_change("<div id='changePasswordModal'></div>", "")
    print("guard OK")
