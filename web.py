"""
web.py - Web-surfing scenario for the Synthetic Human Agent.
Drives Chrome via engine.py, keyboard-first, built for the CLEAN PATH.
Cold-launch discipline: close ALL Chrome windows before surf_session runs.

BATCH 1 changes (vs the version in the repo):
  1. surf_session() now RETURNS a dict with a per-site log; sequencer logs it.
  2. Every site gets its own deadline carved out of the remaining block budget.
     When a site runs late it skips its remaining scrolls / link click / find
     instead of overrunning, but every planned site is still visited.
  3. All dwells (landing, reading, after a link click) come from one bounded,
     STOP-aware helper, _dwell(). No more unbounded human_delay("read") stacking
     and no more fixed time.sleep(3, 9).
  4. thisisunsafe is typed only when the window title really says "Privacy error".
  5. _save_page() waits for the "Save As" window before typing the file name.

UNVERIFIED engine.py assumptions (P2: please check on the VM, see notes at the
bottom of the file):  wait_for_window(title_substring, timeout) returns truthy
when a matching window appears and falsy on timeout (or raises);  press("escape")
is a valid key name;  engine.SPEED may be missing.
"""

import math
import os
import random
import time

import engine

DEV_BYPASS_CERT = True   # DEV ONLY: pass expired-cert pages. Set False for the clean run.

# CLEAN LIST ONLY (load without a cert warning).
CLEAN_SITES = [
    "https://www.apple.com/",
    "https://www.amazon.com/",
    "https://www.dell.com/",
    "https://www.samsung.com/",
    "https://www.coursera.org/",
    "https://www.coinbase.com/",
    "https://www.coingecko.com/",
    "https://apnews.com/",
    "https://aws.amazon.com/",
]

# Plausible filenames for the occasional "Save page as" action.
_SAVE_NAMES = ["article", "pricing-page", "product-specs", "notes",
               "reference", "saved-page", "overview"]

CHROME_WINDOW_HINT = "Chrome"           # TODO(VM): confirm wait_for_window sees this
CERT_WINDOW_HINT = "Privacy error"      # title of Chrome's certificate warning page
SAVE_WINDOW_HINT = "Save As"            # title of the Windows save dialog

# Dwell settings: (median s, sigma, min s, max s). One place to tune everything.
LANDING = (4.0, 0.5, 2.0, 8.0)          # first look at a freshly loaded page
READING = (11.0, 0.6, 5.0, 25.0)        # pause between scroll bursts
LONG_READ = (25.0, 0.4, 15.0, 45.0)     # now and then something catches the eye
LONG_READ_P = 0.10
AFTER_LINK = (3.5, 0.4, 2.0, 6.0)       # glance at the page a link opened
FINAL_LOOK = (4.0, 0.5, 2.0, 10.0)      # last look before leaving a site
CLOSE_RESERVE_S = 20                    # kept back for Ctrl+W x N and the exit wait
SAVE_RESERVE_S = 12                     # extra kept back when a Save page as is planned


# --------------------------------------------------------------------------- helpers

def _stop_check():
    if os.path.exists("STOP"):
        raise SystemExit("STOP file found")


def _speed():
    return float(getattr(engine, "SPEED", 1.0) or 1.0)


def _left(deadline):
    return float("inf") if deadline is None else deadline - time.monotonic()


def _dwell(params, deadline=None):
    """Pause like a person reading: log-normal, clipped to [lo, hi], scaled by the
    persona speed, never past `deadline`, checks the STOP file every second."""
    median, sigma, lo, hi = params
    d = random.lognormvariate(math.log(median), sigma) * _speed()
    d = max(lo, min(hi, d))
    d = min(d, max(0.0, _left(deadline)))
    end = time.monotonic() + d
    while True:
        rest = end - time.monotonic()
        if rest <= 0:
            break
        _stop_check()
        time.sleep(min(1.0, rest))
    return d


def _window_appeared(hint, timeout):
    """True if a window whose title contains `hint` shows up within `timeout` s.
    Any exception from engine counts as 'did not appear'."""
    try:
        return bool(engine.wait_for_window(hint, timeout=timeout))
    except Exception:
        return False


# --------------------------------------------------------------------------- actions

def _dismiss_cert_warning():
    """Chrome's hidden bypass: with the error page focused, type the magic phrase.
    Only done when the warning page really is on screen. Returns True if typed."""
    if not _window_appeared(CERT_WINDOW_HINT, timeout=4):
        return False
    engine.human_delay("read")        # look at the red warning for a moment
    engine.type_text("thisisunsafe")
    engine.human_delay("action")
    return True


def _navigate(url, new_tab=False):
    """Returns True if a certificate warning was bypassed."""
    if new_tab:
        engine.press("ctrl", "t")
        engine.human_delay("action")
    engine.press("ctrl", "l")          # focus address bar (reliable + human)
    engine.human_delay("action")
    engine.type_text(url)
    engine.human_delay("action")
    engine.press("enter")
    if DEV_BYPASS_CERT:
        return _dismiss_cert_warning()
    return False


def _scroll_burst():
    # a short flick of the wheel: a few small scrolls with tiny gaps
    for _ in range(random.randint(2, 4)):
        engine.scroll(random.randint(1, 3))   # positive = down
        engine.human_delay("scroll")


def _click_a_link(deadline=None):
    for _ in range(random.randint(3, 10)):
        engine.press("tab")
        engine.human_delay("scroll")
    engine.press("enter")
    _dwell(AFTER_LINK, deadline)       # glance at the (usually dead) link, then bounce back
    engine.press("alt", "left")
    engine.human_delay("action")


_FIND_WORDS = ["the", "more", "about", "home", "contact", "support", "price", "learn"]

def _find_on_page():
    """Ctrl+F, type a word, jump to it, close. In-page find bar - no OS dialog."""
    engine.press("ctrl", "f")
    engine.human_delay("action")
    engine.type_text(random.choice(_FIND_WORDS))
    engine.human_delay("action")
    engine.press("enter")                  # jump to first match
    engine.human_delay("notice")           # glance at it (was "read": too long for a glance)
    engine.press("escape")                 # close the find bar


def _bookmark_page():
    """Ctrl+D saves the bookmark immediately; Escape closes the bubble and returns focus to the page."""
    engine.press("ctrl", "d")
    engine.human_delay("notice")       # bubble appears (bookmark already saved)
    engine.press("escape")             # dismiss it, focus back on the page
    engine.human_delay("action")       # 'Done'


def _switch_tab():
    """Flip to the next open tab, like a person checking another one."""
    engine.press("ctrl", "tab")
    engine.human_delay("action")


def _save_page():
    """Save the current page: Ctrl+S -> (wait for dialog) -> filename -> Enter.
    Returns True if the file name was typed and confirmed."""
    engine.press("ctrl", "s")
    if not _window_appeared(SAVE_WINDOW_HINT, timeout=8):
        engine.press("escape")          # nothing (or something odd) opened: back out
        engine.human_delay("action")
        return False
    engine.human_delay("read")          # dialog is open; look at it
    engine.type_text(random.choice(_SAVE_NAMES) + "-" + str(random.randint(1000, 9999)))
    engine.human_delay("action")
    engine.press("enter")               # confirm -> Chrome writes the page to disk
    engine.human_delay("notice")        # brief wait while it writes
    return True


# --------------------------------------------------------------------------- one site

def surf_site(url, new_tab=False, deadline=None):
    """Visit one site. `deadline` is a time.monotonic() value; when it is close,
    optional actions are skipped. Returns {"url", "secs", "actions", "cert"}."""
    t0 = time.monotonic()
    actions = []
    cert = _navigate(url, new_tab=new_tab)
    if cert:
        actions.append("cert-bypass")
    _dwell(LANDING, deadline)          # page loads, eyes land on it

    for _ in range(random.randint(2, 5)):
        if _left(deadline) < 12:
            actions.append("skipped-rest")
            break
        _scroll_burst()
        actions.append("scroll")
        _dwell(LONG_READ if random.random() < LONG_READ_P else READING, deadline)

    if random.random() < 0.7:          # most visits include a link click
        if _left(deadline) > 15:
            _click_a_link(deadline)
            actions.append("link")
        else:
            actions.append("skip-link")
    if random.random() < 0.30 and _left(deadline) > 8:
        _find_on_page()
        actions.append("find")
    if random.random() < 0.12 and _left(deadline) > 6:
        _bookmark_page()
        actions.append("bookmark")
    _dwell(FINAL_LOOK, deadline)
    return {"url": url, "secs": round(time.monotonic() - t0), "actions": actions, "cert": cert}


# --------------------------------------------------------------------------- session

def surf_session(n_sites=None, max_seconds=None):
    """Run one Chrome session. Returns a dict for the sequencer's log:
    {"ok", "secs", "planned", "visited", "saved", "log": [per-site dicts]}"""
    start = time.monotonic()
    end = None if max_seconds is None else start + max_seconds
    if n_sites is None:
        n_sites = random.randint(2, 4)
    sites = random.sample(CLEAN_SITES, k=min(n_sites, len(CLEAN_SITES)))

    result = {"ok": True, "secs": 0, "planned": len(sites), "visited": 0,
              "saved": None, "log": []}

    engine.launch("chrome")                 # Win+R -> explorer parent (COLD launch!)
    engine.wait_for_window(CHROME_WINDOW_HINT, timeout=20)
    engine.human_delay("action")

    open_tabs = 1
    save_idx = random.randrange(len(sites)) if random.random() < 0.6 else -1
    for i, url in enumerate(sites):
        deadline = None
        if end is not None:
            reserve = CLOSE_RESERVE_S + (SAVE_RESERVE_S if save_idx >= i else 0)
            time_left = end - time.monotonic() - reserve
            if time_left < 20 and i > 0:
                result["log"].append({"url": url, "secs": 0, "actions": ["not-started"]})
                continue                    # truly out of budget
            sites_left = len(sites) - i
            deadline = time.monotonic() + max(25.0, time_left / sites_left)
        new_tab = (i > 0 and open_tabs < 3 and random.random() < 0.5)
        if new_tab:
            open_tabs += 1
        entry = surf_site(url, new_tab=new_tab, deadline=deadline)
        result["visited"] += 1
        if open_tabs > 1 and random.random() < 0.3:
            _switch_tab()
            entry["actions"].append("tab-switch")
        if i == save_idx:
            if _save_page():
                result["saved"] = url
                entry["actions"].append("save-page")
            else:
                entry["actions"].append("save-failed")
        result["log"].append(entry)
        engine.human_delay("action")

    for _ in range(open_tabs):
        engine.press("ctrl", "w")           # close tabs; last one closes the window
        engine.human_delay("action")
    engine.human_delay("read")              # let chrome.exe fully exit before returning
    result["secs"] = round(time.monotonic() - start)
    result["tabs"] = open_tabs
    return result


if __name__ == "__main__":
    # Close ALL Chrome windows before running.
    print(surf_session(n_sites=2, max_seconds=330))