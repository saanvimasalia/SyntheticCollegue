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
URL_TYPO_RATE = 0.015                   # slips while typing a URL (fixed with Backspace)

# Look-up -> Notepad research step (see research_to_notepad below)
RESEARCH = True                         # do one search and take notes in Notepad per session
RESEARCH_EST_S = 200                    # kept back from the site visits (VM test, no
                                        # deadline pressure: 242 s; with a deadline the
                                        # reading dwells and extra notes shrink)
RESEARCH_MIN_S = 130                    # skip the step if less than this is left
NOTEPAD_PART_S = 120                    # kept back from the reading for Notepad: open, paste,
                                        # trim, heading, link, notes, save, close
SPLIT_SCREEN = True                     # Chrome left half, Notepad right half while taking notes
NOTEPAD_RETURN = None                   # back to Notepad: None = random, "click" / "alttab" (tests)
NOTEPAD_HINT = "Notepad"
SAVE_DIALOG_HINT = "Save as"            # Notepad's dialog (wait_for_window ignores case)
# Only CLEAN_SITES are reachable on the VM, so "research" = go to the site that
# fits the topic, look for a keyword on the page (Ctrl+F), copy what's there.
# url -> [(topic for the notes heading, word to look for on the page)]
RESEARCH_TOPICS = {
    "https://www.apple.com/":     [("macbook for university", "mac"), ("iphone upgrade", "iphone")],
    "https://www.amazon.com/":    [("laptop stand options", "deals"), ("gift ideas", "gift")],
    "https://www.dell.com/":      [("dell laptop specs", "laptop"), ("monitor for home office", "monitor")],
    "https://www.samsung.com/":   [("galaxy tablet battery life", "galaxy"), ("new tv prices", "tv")],
    "https://www.coursera.org/":  [("data science course", "data"), ("free online courses", "free")],
    "https://www.coinbase.com/":  [("how crypto buying works", "buy"), ("bitcoin basics", "bitcoin")],
    "https://www.coingecko.com/": [("bitcoin price today", "bitcoin"), ("top coins by market cap", "market")],
    "https://apnews.com/":        [("today's headlines", "news"), ("business news", "business")],
    "https://aws.amazon.com/":    [("aws free tier limits", "free"), ("cloud storage pricing", "pricing")],
}
NOTE_LINES = [
    "- compare prices again before deciding",
    "- most reviews mention the battery life",
    "- check if the student discount still applies",
    "- worth asking the team what they use",
    "- bookmark the pricing page for later",
    "- the second result looked more reliable",
    "- need to double check the dates",
]


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
    engine.type_human(url, typo_rate=URL_TYPO_RATE)
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


def _wander_mouse():
    """Small aimless move, like a hand resting on the mouse while reading."""
    try:
        x, y = engine.cursor_pos()
        w, h = engine.user32.GetSystemMetrics(0), engine.user32.GetSystemMetrics(1)
        engine.move_to(max(5, min(w - 5, x + random.randint(-200, 200))),
                       max(5, min(h - 5, y + random.randint(-120, 120))))
    except Exception:
        pass                              # cosmetic: never let it break a visit


def _scroll_back_up(deadline=None):
    """Scroll back up to something already passed and read it again."""
    engine.scroll(-random.randint(3, 8))   # negative = up
    _dwell(READING, deadline)
    engine.scroll(random.randint(2, 6))
    engine.human_delay("scroll")


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
        if random.random() < 0.3:
            _wander_mouse()            # hand drifts while reading
        _dwell(LONG_READ if random.random() < LONG_READ_P else READING, deadline)

    # optional actions in a different order every visit
    extras = []
    if random.random() < 0.7:          # most visits include a link click
        extras.append(("link", 15, lambda: _click_a_link(deadline)))
    if random.random() < 0.30:
        extras.append(("find", 8, _find_on_page))
    if random.random() < 0.12:
        extras.append(("bookmark", 6, _bookmark_page))
    if random.random() < 0.25:
        extras.append(("reread", 14, lambda: _scroll_back_up(deadline)))
    random.shuffle(extras)
    for name, need, fn in extras:
        if _left(deadline) <= need:
            actions.append("skip-" + name)
            continue
        fn()
        actions.append(name)
        if engine.maybe_pause(0.15):   # a stop between steps
            actions.append("pause")
    _dwell(FINAL_LOOK, deadline)
    return {"url": url, "secs": round(time.monotonic() - t0), "actions": actions, "cert": cert}


# --------------------------------------------------------------------------- research -> Notepad

def _alt_tab(hint, timeout=6):
    """Alt+Tab back to the previous window; True if `hint` is now in front."""
    engine.press("alt", "tab")
    ok = _window_appeared(hint, timeout)
    engine.human_delay("action")
    return ok


def _down(n):
    for _ in range(n):
        engine.press("down")
        engine.human_delay("keystroke")


def _snap(side, hint):
    """Win+Left / Win+Right: snap the front window to half the screen. On Windows
    11 Snap Assist may then offer the other windows; Escape dismisses it.
    True if `hint` is still the front window afterwards."""
    engine.press("win", side)
    engine.human_delay("action")
    if not _window_appeared(hint, 1.5):
        engine.press("escape")             # close Snap Assist
        engine.human_delay("notice")
    return _window_appeared(hint, 3)


def _to_notepad(split=False):
    """Back to Notepad: when split, half the time click into its text on the right
    half, otherwise Alt+Tab. The click moves the caret, so callers put it back
    with the keyboard. (Not the title bar: in Win11 Notepad that focuses the tab
    strip and typed text is lost, VM test 21 Oct.)"""
    use_click = random.random() < 0.5 if NOTEPAD_RETURN is None else NOTEPAD_RETURN == "click"
    if split and use_click:
        try:
            w = engine.user32.GetSystemMetrics(0)
            h = engine.user32.GetSystemMetrics(1)
            engine.move_to(int(w * random.uniform(0.62, 0.9)), int(h * random.uniform(0.4, 0.7)))
            engine.human_delay("notice")
            engine.click()
            if _window_appeared(NOTEPAD_HINT, 3):
                engine.human_delay("action")
                return True
        except Exception:
            pass
    return _alt_tab(NOTEPAD_HINT)


def _glance_at_chrome(deadline=None):
    """While writing notes, look back at the page on the left for a moment."""
    if not _alt_tab(CHROME_WINDOW_HINT):
        return _window_appeared(NOTEPAD_HINT, 1)
    engine.scroll(random.choice([-1, 1]) * random.randint(2, 5))
    _dwell(READING, deadline)
    return _to_notepad(split=True)


def research_to_notepad(deadline=None, avoid=()):
    """Look something up on one of the CLEAN_SITES in a new Chrome tab (Ctrl+F for
    a keyword, read around it), copy the page into Notepad, then tidy it up there:
    cut it down to the useful top part, delete a junk line, add a heading, the
    link and a few notes of our own, save. Sites in `avoid` (already visited this
    session) are skipped when possible. Notepad is closed again afterwards;
    Chrome stays open for the caller.
    Returns {"query", "url", "file", "steps", "typos", "secs"}."""
    t0 = time.monotonic()
    urls = [u for u in RESEARCH_TOPICS if u not in avoid] or list(RESEARCH_TOPICS)
    url = random.choice(urls)
    query, word = random.choice(RESEARCH_TOPICS[url])
    rec = {"query": query, "url": url, "file": None, "steps": [], "typos": 0, "secs": 0}

    def in_notepad():
        """Only type while Notepad is really in front - anything typed into Chrome
        by mistake ends up as a Google search (offline dino page on the VM)."""
        if _window_appeared(NOTEPAD_HINT, 2):
            return True
        rec["steps"].append("lost-notepad")
        rec["secs"] = round(time.monotonic() - t0)
        return False

    # Reading dwells may only use time the Notepad part doesn't need, otherwise
    # they eat the deadline and the link / notes get skipped (VM test 21 Oct).
    read_by = None if deadline is None else deadline - NOTEPAD_PART_S
    paste_by = None if deadline is None else deadline - (NOTEPAD_PART_S - 15)

    # 1. go to the site and look for the topic on the page
    _navigate(url, new_tab=True)
    _dwell(LANDING, read_by)
    engine.press("ctrl", "f")
    engine.human_delay("action")
    rec["typos"] += engine.type_human(word)
    engine.human_delay("action")
    for _ in range(random.randint(1, 3)):  # step through a few matches
        engine.press("enter")
        _dwell(READING, read_by)
    engine.press("escape")
    engine.human_delay("action")
    if random.random() < 0.6 and _left(read_by) > 10:
        _scroll_burst()
        _dwell(READING, read_by)
    engine.press("home")                   # back to the top before selecting
    engine.human_delay("scroll")
    rec["steps"].append("search")

    # 2. copy the page text
    engine.press("ctrl", "a")
    engine.human_delay("notice")
    engine.press("ctrl", "c")
    engine.human_delay("action")
    rec["steps"].append("copy")
    if _left(deadline) > 120:
        engine.maybe_pause(0.3)

    # 3. Chrome to the left half, Notepad (from the Start menu) to the right, paste
    split = SPLIT_SCREEN and _snap("left", CHROME_WINDOW_HINT)
    engine.open_app("notepad", expect=NOTEPAD_HINT, fallback="notepad", timeout=15)
    if not _window_appeared(NOTEPAD_HINT, 10):
        rec["steps"].append("notepad-missing")
        return rec
    engine.human_delay("action")
    engine.press("ctrl", "n")              # fresh page (Win11 Notepad may restore old tabs)
    engine.human_delay("action")
    if split and _snap("right", NOTEPAD_HINT):
        rec["steps"].append("split")
    if not in_notepad():
        return rec
    engine.press("ctrl", "v")
    _dwell(READING, paste_by)              # look over what got pasted
    rec["steps"].append("paste")

    # 4. keep only the top of it: down a few lines, select to the end, delete
    if not in_notepad():
        return rec
    engine.press("ctrl", "home")
    engine.human_delay("action")
    _down(random.randint(8, 14))
    engine.press("ctrl", "shift", "end")
    engine.human_delay("notice")
    engine.press("delete")
    engine.human_delay("action")
    rec["steps"].append("trim")

    # 5. delete one junk line near the top (menu text, "Sign in" ...)
    engine.press("ctrl", "home")
    _down(random.randint(1, 4))
    engine.press("home")
    engine.press("shift", "end")
    engine.human_delay("notice")
    engine.press("delete")
    engine.human_delay("action")
    rec["steps"].append("delete-line")
    if _left(deadline) > 90:
        engine.maybe_pause(0.25)

    # 6. heading at the top
    if not in_notepad():
        return rec
    engine.press("ctrl", "home")
    engine.human_delay("action")
    rec["typos"] += engine.type_human("Notes: %s\n" % query)

    # 7. fetch the link from Chrome and paste it under the heading
    if _left(deadline) > 45 and _alt_tab("Chrome"):
        engine.press("ctrl", "l")
        engine.human_delay("notice")
        engine.press("ctrl", "c")
        engine.human_delay("action")
        engine.press("escape")
        if _to_notepad(split) and in_notepad():
            engine.press("ctrl", "home")   # caret back to the line under the heading
            engine.press("down")
            engine.press("home")
            engine.human_delay("notice")
            engine.type_text("Source: ")
            engine.press("ctrl", "v")
            engine.press("enter")
            engine.human_delay("action")
            rec["steps"].append("link")
        else:
            if "lost-notepad" not in rec["steps"]:
                rec["steps"].append("lost-notepad")
            return rec
    if not in_notepad():
        return rec
    engine.press("enter")

    # 8. our own notes at the bottom, with one word rethought
    engine.press("ctrl", "end")
    engine.human_delay("action")
    rec["typos"] += engine.type_human("\n\nMy takeaways:\n")
    glance_at = random.randint(0, 1) if split else -1
    for n, line in enumerate(random.sample(NOTE_LINES, random.randint(2, 3))):
        if _left(deadline) < 35:
            break
        if n == glance_at and _left(deadline) > 70:
            if not _glance_at_chrome(deadline):     # check the page again, come back
                rec["steps"].append("lost-notepad")
                return rec
            engine.press("ctrl", "end")
            rec["steps"].append("glance")
        engine.human_delay("think")
        if not in_notepad():
            return rec
        rec["typos"] += engine.type_human(line + "\n")
    if _left(deadline) > 30 and in_notepad():
        engine.type_text("- overall ")
        engine.retype_word("good", "useful")
        engine.type_text(" for now\n")
        rec["steps"].append("retype")
    rec["steps"].append("notes")
    _dwell(FINAL_LOOK, deadline)            # read it over

    # 9. save and close Notepad
    first = "".join(c for c in query.split()[0] if c.isalnum())
    fname = "research_%s_%d.txt" % (first, random.randint(10, 99))
    if not in_notepad():
        return rec
    engine.press("ctrl", "s")
    if _window_appeared(SAVE_DIALOG_HINT, 8):
        engine.human_delay("action")
        engine.type_human(fname, typo_rate=0.01)
        engine.human_delay("action")
        engine.press("enter")
        if _window_appeared(fname.rsplit(".", 1)[0], 8):
            rec["file"] = fname
            rec["steps"].append("save")
        else:
            rec["steps"].append("save-unconfirmed")
    else:
        rec["steps"].append("save-failed")
    engine.human_delay("action")
    engine.press("alt", "f4")
    engine.human_delay("action")
    if _window_appeared(NOTEPAD_HINT, 2) and rec["file"] is None:
        engine.press("tab")                # "Save changes?" -> Don't save
        engine.human_delay("notice")
        engine.press("enter")
    if _window_appeared(CHROME_WINDOW_HINT, 6) and split:   # focus falls back to Chrome
        # full screen again (Chrome remembers its size): window menu -> Maximize.
        # Not Win+Up: on a snapped window that only moves it to a quarter.
        engine.press("alt", "space")
        engine.human_delay("notice")
        engine.press("x")
        engine.human_delay("action")
    rec["secs"] = round(time.monotonic() - t0)
    return rec


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
              "saved": None, "research": None, "log": []}
    research = RESEARCH and (end is None or max_seconds >= RESEARCH_EST_S + 100)
    site_end = None if end is None else end - (RESEARCH_EST_S if research else 0)

    # Start menu -> explorer parent (COLD launch!); Win+R only if that fails
    result["launch"] = engine.open_app("chrome", expect=CHROME_WINDOW_HINT,
                                       fallback="chrome", timeout=20)
    engine.wait_for_window(CHROME_WINDOW_HINT, timeout=20)
    engine.human_delay("action")

    open_tabs = 1
    save_idx = random.randrange(len(sites)) if random.random() < 0.6 else -1
    for i, url in enumerate(sites):
        deadline = None
        if end is not None:
            reserve = CLOSE_RESERVE_S + (SAVE_RESERVE_S if save_idx >= i else 0)
            time_left = site_end - time.monotonic() - reserve
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
        engine.maybe_pause(0.25)            # a stop before the next site

    if research:
        r_deadline = None if end is None else end - CLOSE_RESERVE_S
        if r_deadline is None or r_deadline - time.monotonic() >= RESEARCH_MIN_S:
            open_tabs += 1                  # the search opens its own tab
            try:
                result["research"] = research_to_notepad(r_deadline, avoid=sites)
            except SystemExit:
                raise
            except Exception as e:
                result["research"] = {"error": repr(e)}
                _window_appeared(CHROME_WINDOW_HINT, 4)
        else:
            result["research"] = "skipped: out of time"

    for _ in range(open_tabs):
        engine.press("ctrl", "w")           # close tabs; last one closes the window
        engine.human_delay("action")
    if _window_appeared(CHROME_WINDOW_HINT, 2):
        # a tab we didn't count (e.g. a link that opened one): close the whole window.
        # Only while Chrome is in front - Ctrl+Shift+W closes a terminal tab too.
        engine.press("ctrl", "shift", "w")
        engine.human_delay("action")
        result["closed_window"] = True
    engine.human_delay("read")              # let chrome.exe fully exit before returning
    result["secs"] = round(time.monotonic() - start)
    result["tabs"] = open_tabs
    return result


if __name__ == "__main__":
    # Close ALL Chrome windows before running.
    print(surf_session(n_sites=2, max_seconds=330))