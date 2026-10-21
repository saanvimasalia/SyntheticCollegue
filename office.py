"""
office.py -- P3: Office (Writer) scenario for the Synthetic Human Agent (team05).

Hooks for P1's sequencer
------------------------
    office.write_document(max_seconds=None, doc_type=None) -> dict
        Full end-to-end task: cold-launch Writer, type a short document with
        human-timed typing + typo/backspace model, Save As with a realistic
        filename, keep typing, save again, quit Writer.

    office.edit_document(path=None, max_seconds=None) -> dict      (optional)
        Re-open the last saved document (cold launch), add a paragraph,
        save, quit.  Only use after write_document() succeeded.

Both return a result dict and never raise (unless raise_on_error=True).
The dict carries a "log" list (this call's action-log lines) so the sequencer
can write it to agent_log.txt; office never imports the sequencer.
Both leave Writer CLOSED on exit (cold-launch discipline for later scenarios).

write_document(max_seconds) is budget-aware: before each paragraph it checks
that typing it plus the wrap-up (final save + close) still fits, so it should
finish inside max_seconds (launch time included).  edit_document ignores
max_seconds (it is short, ~1 min).

Rules followed (see MASTER_BUILD_PLAN section 11)
-------------------------------------------------
* Launch ONLY through engine.open_app() / engine.launch() (Start-menu search,
  Win+R as fallback, both via SendInput -> explorer parent).
* No subprocess / os.startfile / Start-Process anywhere in this file.
  (Window checks use ctypes user32 only, so no extra python->child process
  events pollute the Sysmon log.)
* No fixed sleeps for human-facing delays: every wait is sampled
  (engine.human_delay or a log-normal pause).
* Stdlib only, no busy loops (polling sleeps are 0.5 s).

Engine keys this file needs (all mapped by P1; backspace/home/end NOT yet run):
    ctrl, s, q, b, backspace, enter        (+ ctrl+end for edit_document,
                                            home/end in --keytest)
Run `python office.py --keytest` FIRST on the VM: it types a line with
( ) : _ , digits and capitals, uses backspace/home/end, saves, and then reads
the typed text back out of the saved .odt to compare it with what was meant.

launch() convention (P1): path already quoted, args as a string:
    engine.launch('"C:\\\\...\\\\soffice.exe"', "-writer")
This works both before and after P1's launch() update.

    office.write_spreadsheet(max_seconds=None, sheet_type=None) -> dict   (C1)
        Cold-launch Calc, type a small table (header row, 5-8 data rows, a
        formula per row, a SUM footer) with Tab/Enter and the same human timing
        and typo model, Save As .ods, keep going, save again, quit.  The saved
        .ods is read back and every cell/formula is checked (rec["verified"]).

Known dialogs (S1 safety net)
-----------------------------
If Writer was killed earlier, the next start can open the "Document Recovery"
window (buttons Start Recovery / Cancel) instead of a blank document, and the
run would stall.  Every launch wait now watches for such windows (title match
+ owned by soffice) and dismisses them with Escape (= Cancel), falling back to
Alt+F4 only when the dialog is verifiably the foreground window.  Handled
dialogs are logged and returned in the result dict under "dialogs".

CLI
---
    python office.py              one write_document() run (budget 4 min)
    python office.py --keytest    verify backspace/home/end + ( ) : _ typing
    python office.py --quick      short run with lots of typos (key/typo check)
    python office.py --edit       write_document() then edit_document()
    python office.py --calc         one write_spreadsheet() run (budget 4 min)
    python office.py --calc-quick   short Calc run with lots of typos (90 s)
    python office.py --calc-keytest verify Tab/Enter entry + formulas in Calc
    python office.py --calc-navtest verify arrow-key navigation + bold header in Calc
    python office.py --dialog-keys  check Escape and Alt+F4 are mapped by the engine
    python office.py --probe        start Writer, log every Office window, dismiss
                                    a recovery dialog if one shows up
"""

import html
import math
import os
import random
import re
import sys
import time
import zipfile

import engine


# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------
SOFFICE_PATH = None          # set to override auto-detect
SOFFICE_CANDIDATES = [
    r"C:\Program Files (x86)\OpenOffice 4\program\soffice.exe",
    r"C:\Program Files\OpenOffice 4\program\soffice.exe",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]
WRITER_ARGS = ["-writer"]            # opens straight into a blank Writer doc
WINDOW_PART = "Writer"               # title substring of the Writer window
OFFICE_TITLE_RE = re.compile(r"(OpenOffice|LibreOffice)", re.I)

TYPO_RATE = 0.05          # chance per eligible word
TYPOS_ENABLED = True      # auto-disabled if backspace is not supported
BOLD_TITLE = True         # Ctrl+B around the title line
SAVE_EXT = ".odt"         # default ODF format: no "Keep format?" dialog
WRAP_UP_RESERVE = 25.0    # s kept free for final save + read glance + close

ACTION_LOG = []
_LAST = {"path": None, "name": None, "kind": None}
_LEFT_OPEN = {"open": False}   # True while write_document(keep_open=True) has left Writer running


class OfficeError(RuntimeError):
    pass


def _log(msg):
    line = "[office %s] %s" % (time.strftime("%H:%M:%S"), msg)
    ACTION_LOG.append(line)
    print(line, flush=True)


# --------------------------------------------------------------------------
# Sampled pauses (log-normal, heavy tail) -- never a fixed sleep
# --------------------------------------------------------------------------
def _pause(median, sigma=0.5, lo=0.05, hi=30.0):
    t = random.lognormvariate(math.log(median), sigma)
    time.sleep(min(max(t, lo), hi))


def _delay(kind, median, sigma=0.5, lo=0.05, hi=30.0):
    """Use the engine's central delay kind (so it is tuned in one place);
    fall back to our own sampled pause while the engine lacks that kind
    ("think"/"notice" raise KeyError until P1's update is in)."""
    try:
        engine.human_delay(kind)
    except KeyError:
        _pause(median, sigma, lo, hi)


# --------------------------------------------------------------------------
# Window helpers (ctypes only: creates no processes)
# --------------------------------------------------------------------------
def _enum_windows():
    """[(hwnd, title)] of visible top-level windows. Empty off Windows."""
    if sys.platform != "win32":
        return []
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    found = []
    enum_proc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

    def cb(hwnd, _lparam):
        if user32.IsWindowVisible(hwnd):
            n = user32.GetWindowTextLengthW(hwnd)
            if n:
                buf = ctypes.create_unicode_buffer(n + 1)
                user32.GetWindowTextW(hwnd, buf, n + 1)
                found.append((hwnd, buf.value))
        return True

    user32.EnumWindows(enum_proc(cb), 0)
    return found


def _office_windows():
    return [(h, t) for h, t in _enum_windows() if OFFICE_TITLE_RE.search(t)]


def _title_present(part):
    p = part.lower()
    return any(p in t.lower() for _h, t in _enum_windows())


def _focus(hwnd):
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.SetForegroundWindow(hwnd)
        except Exception:
            pass


def _find_window(part):
    p = part.lower()
    for h, t in _enum_windows():
        if p in t.lower():
            return h
    return None


def _wait(title_part, timeout=15, dialogs=False):
    """
    Wait for a window whose title contains `title_part`.
    On Windows we poll the visible-window list ourselves (any visible window,
    not only the foreground one - engine.wait_for_window only sees the
    foreground window and can give up early while Writer is still starting),
    then bring that window to the front so the keystrokes that follow reach it.
    dialogs=True: known dialogs (Document Recovery ...) are dismissed while
    waiting, and each one adds 15 s to the wait.
    Off Windows (tests) we defer to engine.wait_for_window.
    """
    if sys.platform == "win32":
        end = time.monotonic() + timeout
        extended = 0
        while True:
            if dialogs:
                n = len(_handle_dialogs())
                if n and extended < 3:                  # never extend forever
                    end += 15
                    extended += 1
            hwnd = _find_window(title_part)
            if hwnd:
                _focus(hwnd)
                time.sleep(0.3)
                return True
            if time.monotonic() >= end:
                break
            time.sleep(0.5)
        titles = [t for _h, t in _enum_windows()][:10]
        _log("window %r not seen after %ss; visible windows: %r" % (title_part, timeout, titles))
        return False
    if dialogs:
        _handle_dialogs()
    try:
        r = engine.wait_for_window(title_part, timeout=timeout)
    except Exception as exc:
        _log("wait_for_window(%r) raised %r" % (title_part, exc))
        return False
    return True if r is None else bool(r)


def _wait_closed(timeout=20):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        if not _office_windows():
            return True
        time.sleep(0.5)
    return not _office_windows()


# --------------------------------------------------------------------------
# Known dialogs (e.g. Document Recovery after Writer was killed)
# --------------------------------------------------------------------------
OFFICE_EXES = ("soffice.bin", "soffice.exe")
# plans = key sequences to try, in order.  Escape is the dialog's Cancel button
# (for Document Recovery: "skip recovery").  Alt+F4 is only sent when the
# dialog is verifiably the foreground window (it would close Writer otherwise).
# Observed on VM 1051 (12 Oct, OpenOffice 4.1.16): the Document Recovery window
# is titled only "OpenOffice 4.1.16" (exe soffice.bin); "Document Recovery" is a
# heading INSIDE the window, so it cannot be matched by title.  A window owned
# by soffice whose title is just the product name (+ version) is never a
# document window ("Untitled 1 - OpenOffice Writer") - it is a startup/message
# dialog - so Escape (Cancel) is the safe response.
KNOWN_DIALOGS = [
    {"name": "recovery",
     "title": re.compile(r"recover", re.I),
     "plans": [("escape",), ("alt", "f4")]},
    {"name": "startup-dialog (Document Recovery on OpenOffice 4.x)",
     "title": re.compile(r"^\s*(OpenOffice|LibreOffice)(\s+[\d.]+)?\s*$", re.I),
     "plans": [("escape",), ("alt", "f4")],
     "need_exe": True},       # only trust it when the owner is verifiably soffice
]
# Escape on the recovery window opens "Are you sure you want to cancel the
# OpenOffice document recovery?  [Yes] [No]" - ALSO titled just "OpenOffice
# 4.1.16".  Yes is the default button, so Enter confirms.  Only used for a
# bare-title window that is owned by / smaller than another bare-title window.
CONFIRM_SPEC = {"name": "recovery-confirm (cancel recovery? Yes)",
                "title": None,
                "plans": [("enter",), ("alt", "y")]}
_DIALOG_TRIES = {}            # hwnd -> times we acted on it (flap guard)
_ESCAPED = set()              # windows we pressed Escape on = the recovery window
                              # itself, never the Yes/No box: no Enter goes there
MAX_DIALOG_TRIES = 3
# Titles of normal windows that must never be treated as a dialog
_MAIN_WIN_RE = re.compile(
    r"(\.od[tsp]\b|\.docx?\b| - (OpenOffice|LibreOffice)\b|Untitled \d"
    r"| - Google Chrome| - Visual Studio Code|PowerShell)", re.I)
DIALOGS_HANDLED = []


def _window_exe(hwnd):
    """Lower-case exe name that owns `hwnd` (ctypes only), or None."""
    if sys.platform != "win32":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        user32, k32 = ctypes.windll.user32, ctypes.windll.kernel32
        k32.OpenProcess.restype = wintypes.HANDLE
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        h = k32.OpenProcess(0x1000, False, pid.value)   # QUERY_LIMITED_INFORMATION
        if not h:
            return None
        try:
            buf = ctypes.create_unicode_buffer(520)
            size = wintypes.DWORD(520)
            if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                return os.path.basename(buf.value).lower()
        finally:
            k32.CloseHandle(h)
    except Exception:
        pass
    return None


def _window_owner(hwnd):
    """Owner window (modal parent) of `hwnd`, 0 when none, None if unknown."""
    if sys.platform != "win32":
        return None
    try:
        import ctypes
        return ctypes.windll.user32.GetWindow(hwnd, 4) or 0     # GW_OWNER
    except Exception:
        return None


def _window_area(hwnd):
    if sys.platform != "win32":
        return None
    try:
        import ctypes
        from ctypes import wintypes
        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r)):
            return max(0, r.right - r.left) * max(0, r.bottom - r.top)
    except Exception:
        pass
    return None


def _foreground_hwnd():
    if sys.platform != "win32":
        return None
    try:
        import ctypes
        return ctypes.windll.user32.GetForegroundWindow()
    except Exception:
        return None


def _dialog_windows():
    """[(hwnd, title, spec)] for visible windows that look like a known dialog
    and belong to Office (or whose owner could not be determined)."""
    out = []
    for h, t in _enum_windows():
        if _MAIN_WIN_RE.search(t):
            continue
        for spec in KNOWN_DIALOGS:
            if spec["title"].search(t):
                exe = _window_exe(h)
                if exe in OFFICE_EXES or (exe is None and not spec.get("need_exe")):
                    out.append((h, t, spec))
                break
    # A second bare-title window on top of the first is the Yes/No confirmation
    bare = [c for c in out if c[2].get("need_exe")]
    if len(bare) >= 2:
        hs = {c[0] for c in bare}
        owners = {c[0]: _window_owner(c[0]) for c in bare}
        confirm = {h for h, o in owners.items() if o in hs}
        if not confirm:                                   # owner info unavailable
            areas = {c[0]: _window_area(c[0]) for c in bare}
            if all(a is not None for a in areas.values()):
                confirm = {min(areas, key=areas.get)}
        out = ([(h, t, CONFIRM_SPEC) for h, t, _s in bare if h in confirm]
               + [c for c in out if c[0] not in confirm])
    return out


def _dismiss_dialog(hwnd, title, spec):
    """Try each key plan.  True when the window is gone afterwards, or when a
    NEW dialog appeared on top (a chain step: Escape -> "cancel recovery?" box;
    _handle_dialogs deals with the new one next round)."""
    known = {c[0] for c in _dialog_windows()}
    for plan in spec["plans"]:
        _focus(hwnd)
        time.sleep(0.4)
        fg = _foreground_hwnd()
        if fg is not None and fg != hwnd:
            _focus(hwnd)                              # one more try
            time.sleep(0.4)
            fg = _foreground_hwnd()
            if fg is not None and fg != hwnd:
                _log("dialog %r is not in front; not sending %s" % (title, "+".join(plan)))
                continue
        try:
            engine.press(*plan)
        except Exception as exc:
            _log("key %s failed on dialog %r (%r)" % ("+".join(plan), title, exc))
            continue
        end = time.monotonic() + 3.0
        while time.monotonic() < end:
            if hwnd not in [h for h, _t in _enum_windows()]:
                return True
            fresh = [c for c in _dialog_windows() if c[0] not in known]
            if fresh:
                _log("new dialog %r appeared after %s" % (fresh[0][1], "+".join(plan)))
                return True
            time.sleep(0.3)
        _log("dialog %r still open after %s" % (title, "+".join(plan)))
    return False


def _handle_dialogs(max_rounds=6):
    """Dismiss any known dialogs that are open.  Returns a list of
    {name, title, dismissed}.  Cheap (one window enumeration) when none are."""
    handled = []
    for _ in range(max_rounds):
        found = [c for c in _dialog_windows()
                 if not (c[2] is CONFIRM_SPEC and c[0] in _ESCAPED)]
        if not found:
            break
        h, t, spec = found[0]
        if _DIALOG_TRIES.get(h, 0) >= MAX_DIALOG_TRIES:
            _log("giving up on dialog %r (%s): acted %d times and it keeps coming back"
                 % (t, spec["name"], MAX_DIALOG_TRIES))
            handled.append({"name": spec["name"], "title": t, "dismissed": False,
                            "gave_up": True})
            break
        _DIALOG_TRIES[h] = _DIALOG_TRIES.get(h, 0) + 1
        if spec is not CONFIRM_SPEC:
            _ESCAPED.add(h)
        _log("dialog %r (%s) is open [hwnd=%s owner=%s area=%s]; dismissing it"
             % (t, spec["name"], h, _window_owner(h), _window_area(h)))
        _pause(1.6, 0.4, 0.8, 5)                      # a person reads it first
        ok = _dismiss_dialog(h, t, spec)
        item = {"name": spec["name"], "title": t, "dismissed": ok}
        handled.append(item)
        DIALOGS_HANDLED.append(item)
        _log("dialog %r %s" % (t, "dismissed" if ok else "could NOT be dismissed"))
        if not ok:
            break
        _pause(0.8, 0.3, 0.4, 2)
    return handled


# --------------------------------------------------------------------------
# Locating things
# --------------------------------------------------------------------------
def _soffice():
    if SOFFICE_PATH:
        return SOFFICE_PATH
    for c in SOFFICE_CANDIDATES:
        if os.path.exists(c):
            return c
    if sys.platform != "win32":
        return SOFFICE_CANDIDATES[0]
    raise FileNotFoundError("soffice.exe not found; set office.SOFFICE_PATH")


def _save_dirs():
    home = os.path.expanduser("~")
    return [
        os.path.join(home, "Documents"),
        os.path.join(home, "Desktop"),
        os.path.join(home, "OneDrive", "Documents"),
        home,
    ]


def _find_saved(filename):
    for d in _save_dirs():
        p = os.path.join(d, filename)
        if os.path.exists(p):
            return p
    return None


# --------------------------------------------------------------------------
# Typo model
# --------------------------------------------------------------------------
_ROWS = ["qwertyuiop", "asdfghjkl", "zxcvbnm"]


def _neighbors(ch):
    c = ch.lower()
    for r, row in enumerate(_ROWS):
        if c in row:
            i = row.index(c)
            out = []
            for d in (-1, 1):
                if 0 <= i + d < len(row):
                    out.append(row[i + d])
            for rr in (r - 1, r + 1):
                if 0 <= rr < len(_ROWS) and i < len(_ROWS[rr]):
                    out.append(_ROWS[rr][i])
            return out
    return []


def _plan_typo(word):
    """
    Plan a mistake in `word` (alphabetic, len >= 3).
    Returns (typed, n_back, fix):  type `typed`, press Backspace n_back
    times, type `fix`  ==>  the document ends up containing `word`.
    The person notices the slip 0-3 characters late, like a real typist.
    """
    n = len(word)
    i = random.randrange(1, n)                 # never the very first letter
    k = random.choice([0, 0, 1, 2, 3])         # chars typed before noticing
    kind = random.choices(
        ["sub", "swap", "double", "omit"], weights=[45, 25, 15, 15])[0]

    def sub():
        nb = _neighbors(word[i])
        if not nb:
            return None
        x = random.choice(nb)
        if word[i].isupper():
            x = x.upper()
        j = min(n, i + 1 + k)
        return x + word[i + 1:j]

    def swap():
        if i + 1 >= n or word[i] == word[i + 1]:
            return None
        j = min(n, i + 2 + k)
        return word[i + 1] + word[i] + word[i + 2:j]

    def double():
        j = min(n, i + 1 + k)
        return word[i] + word[i] + word[i + 1:j]

    def omit():
        j = min(n, i + 1 + k)
        tail = word[i + 1:j]
        return tail if tail else None

    tail = {"sub": sub, "swap": swap, "double": double, "omit": omit}[kind]()
    if tail is None:
        tail = double()
    return word[:i] + tail, len(tail), word[i:]


def _type_human(text, typo_rate=None):
    """Type `text` word by word with hesitations and typo/backspace fixes."""
    global TYPOS_ENABLED
    rate = TYPO_RATE if typo_rate is None else typo_rate
    typos = 0
    words = text.split(" ")
    for idx, w in enumerate(words):
        trail = " " if idx < len(words) - 1 else ""
        if random.random() < 0.04:
            _pause(0.9, 0.5, 0.2, 6)                   # brief hesitation
        core = w.rstrip(".,;:!?)")
        suffix = w[len(core):]
        if (TYPOS_ENABLED and rate > 0 and len(core) >= 3 and core.isalpha()
                and random.random() < rate):
            typed, n_back, fix = _plan_typo(core)
            engine.type_text(typed)
            _delay("notice", 0.5, 0.35, 0.3, 1.5)      # "wait, that's wrong"
            for _ in range(n_back):
                try:
                    engine.press("backspace")
                except Exception as exc:
                    _log("backspace failed (%r); typo model disabled" % (exc,))
                    TYPOS_ENABLED = False
                    break
                engine.human_delay("keystroke")
            engine.type_text(fix + suffix + trail)
            typos += 1
        else:
            engine.type_text(w + trail)
        if suffix.endswith(".") and random.random() < 0.5:
            _pause(1.4, 0.6, 0.3, 8)                   # end-of-sentence beat
    return typos


# --------------------------------------------------------------------------
# Content (no LLM: sentence banks + random slots)
# --------------------------------------------------------------------------
PEOPLE = ["Wei Jie", "Priya", "Daniel", "Mei Ling", "Arjun", "Sarah",
          "Kenneth", "Aisyah", "Marcus", "Hui Min"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
TOPICS = ["the demo", "the budget", "the vendor review", "the onboarding flow",
          "the test results", "the project timeline", "the release plan"]
STUDY_TOPICS = ["operating systems", "computer networks", "data structures",
                "database design", "software testing"]

# Every sentence ends with a full stop: keeps Writer's word-completion from
# swallowing the Enter key at the end of a paragraph.
DOC_TYPES = {
    "meeting": {
        "title": "Meeting notes ({date})",
        "bases": ["Meeting_Notes", "Team_Meeting", "Notes"],
        "intro": [
            "Attendees were {p}, {q} and {r}. We met to go through {topic} and agree on the next steps.",
            "This is a short summary of the meeting held on {day}. The main item on the agenda was {topic}.",
            "Quick notes from today with {p} and {q}. The goal was to align on {topic} before the end of the month.",
        ],
        "body": [
            "We reviewed the progress made since the last meeting and everyone shared a short update.",
            "The main concern was the schedule, since two of the tasks are still waiting for approval.",
            "{p} will follow up with the vendor about the revised quote and report back by {day}.",
            "We also discussed how to split the remaining work more evenly across the team.",
            "A few questions were raised about the testing plan and these will be answered by email.",
            "{q} suggested keeping the next review shorter and sending the slides one day earlier.",
            "There were no major blockers, although the budget still needs to be confirmed.",
            "It was agreed that the draft should be shared with the wider group before it is finalised.",
            "We counted {n} open items on the tracker and agreed to close at least half of them this week.",
        ],
        "close": [
            "The next meeting is planned for {day} afternoon.",
            "Please let me know if I missed anything important.",
            "Action items will be added to the tracker after this document is reviewed.",
        ],
    },
    "project": {
        "title": "Project update ({date})",
        "bases": ["Project_Update", "Status_Update", "Progress"],
        "intro": [
            "This update covers the progress on {topic} over the past week.",
            "Here is a short status update for {topic} ahead of the review on {day}.",
        ],
        "body": [
            "The team completed {n} of the planned tasks and the rest are on track for next week.",
            "{p} finished the first version of the draft and shared it for comments.",
            "One risk is that the test environment is not always available, which slows down the checks.",
            "We plan to start the next phase once the feedback from {q} has been collected.",
            "Overall the schedule looks reasonable, although we may need extra time for the final review.",
            "Documentation is about {n}0 percent complete and should be ready by {day}.",
            "The main change this week was a simpler layout, which makes the steps easier to follow.",
            "Costs are still within the limit we set at the start of the project.",
        ],
        "close": [
            "I will send another update next week.",
            "Feel free to add comments directly in this document.",
        ],
    },
    "report": {
        "title": "Weekly report ({date})",
        "bases": ["Weekly_Report", "Report", "Week_Summary"],
        "intro": [
            "This report summarises what I worked on this week and what is planned next.",
            "Below is a summary of my work this week, mainly around {topic}.",
        ],
        "body": [
            "Most of my time went into {topic}, which took longer than I first expected.",
            "I met with {p} on {day} to clear up a few open questions.",
            "I also spent some time tidying the notes so that they are easier to share.",
            "Two smaller tasks were finished early and one was moved to next week.",
            "The feedback from {q} was helpful and I have already started on the changes.",
            "I wrote down {n} ideas for improving the process and picked the simplest ones to try first.",
            "There were a few delays because some files were not available when I needed them.",
            "Next week I will focus on finishing the draft and preparing for the review.",
        ],
        "close": [
            "Thank you for reading, and please tell me if anything needs more detail.",
            "I will update this report again on Friday.",
        ],
    },
    "study": {
        "title": "Study notes ({date})",
        "bases": ["Study_Notes", "Lecture_Notes", "Revision"],
        "intro": [
            "These are my notes for this week on {topic2}.",
            "Revision notes for {topic2}, written after the lecture on {day}.",
        ],
        "body": [
            "The key idea is to break a big problem into smaller parts and solve each one separately.",
            "I still need to revisit the section on {topic2} before the quiz.",
            "The lecturer gave {n} examples, and the second one was the easiest to follow.",
            "{p} shared a summary sheet that covers most of the main points.",
            "A good way to remember this is to write a short example for every new term.",
            "I found the practice questions harder than the lecture slides suggested.",
            "It helps to read the notes once and then try to explain them without looking.",
            "My plan is to do one past paper on {day} and check the answers with {q}.",
        ],
        "close": [
            "I will add more notes after the tutorial.",
            "Still to do: finish the practice set and clean up these notes.",
        ],
    },
}


def _today():
    return time.strftime("%d %b").lstrip("0")


def _compose(doc_type=None):
    kind = doc_type if doc_type in DOC_TYPES else random.choice(list(DOC_TYPES))
    spec = DOC_TYPES[kind]
    p, q, r = random.sample(PEOPLE, 3)
    fmt = dict(p=p, q=q, r=r, topic=random.choice(TOPICS),
               topic2=random.choice(STUDY_TOPICS), day=random.choice(DAYS),
               n=random.randint(3, 9), date=_today())

    def f(s):
        return s.format(**fmt)

    title = f(spec["title"])
    paras = [f(random.choice(spec["intro"]))]
    body = random.sample(spec["body"], k=random.randint(4, min(7, len(spec["body"]))))
    i = 0
    while i < len(body):
        take = random.choice([2, 2, 3])
        paras.append(" ".join(f(s) for s in body[i:i + take]))
        i += take
    paras.append(f(random.choice(spec["close"])))
    return kind, title, paras, random.choice(spec["bases"])


def _filename(base, ext=None):
    ext = ext or SAVE_EXT
    opts = [time.strftime("%d%b").lstrip("0"), "v2", "final", "draft",
            time.strftime("%H%M"), "", "wk%d" % random.randint(1, 14)]
    suf = random.choice(opts)
    name = base + ("_" + suf if suf else "")
    if _find_saved(name + ext):
        name += "_%d" % random.randint(2, 99)
    return name + ext


def _additions():
    return [
        "I added a few more points after looking at this again.",
        "One more thing to check is the wording in the second paragraph.",
        "After a second read I think the order of the points could be better.",
        "Updated with the latest comments from the team.",
        "Please see the extra notes below, which came up after the discussion.",
        "I will tidy up the formatting once the content is final.",
    ]


# --------------------------------------------------------------------------
# Writer actions
# --------------------------------------------------------------------------
def _ensure_closed():
    """Cold-launch discipline: Writer must not already be running."""
    _DIALOG_TRIES.clear()
    _ESCAPED.clear()
    handled = _handle_dialogs()            # a stale recovery window hides a live soffice
    if handled:
        end = time.monotonic() + 6.0       # Writer may open a blank doc after Cancel
        while not _office_windows() and time.monotonic() < end:
            time.sleep(0.5)
    wins = _office_windows()
    if not wins:
        _LEFT_OPEN["open"] = False
        return
    _log("Writer/Office already open (%r); closing before cold launch" % wins[0][1])
    _focus(wins[0][0])
    time.sleep(0.5)
    engine.press("ctrl", "q")
    if not _wait_closed(15):
        raise OfficeError("Office already running and would not close; "
                          "a Win+R launch would not be a cold launch")


def _close_writer(timeout=20):
    """Ctrl+Q and wait until Writer is really gone (windows AND a short settle
    so soffice.bin has exited) - the next launch must be a cold launch."""
    engine.press("ctrl", "q")
    if not _wait_closed(timeout):
        _log("Writer still open after Ctrl+Q (save prompt?); Enter = Save")
        engine.press("enter")
        if not _wait_closed(timeout):
            raise OfficeError("Writer did not close")
    _LEFT_OPEN["open"] = False
    _pause(1.2, 0.4, 0.6, 4)                      # let soffice.bin finish exiting
    return True


def close_writer(timeout=20):
    """Public: close the Office window that write_document(keep_open=True) left
    open (Ctrl+Q, waits until it is gone).  Safe to call when nothing is open."""
    if not _LEFT_OPEN["open"] and not _office_windows():
        return True
    wins = _office_windows()
    if wins:
        _focus(wins[0][0])
        time.sleep(0.3)
    return _close_writer(timeout)


def _force_close(appear_wait=6.0):
    """Best-effort cleanup after a failure.  Writer may still be starting up
    (its window can appear late), so give it a few seconds to show before
    concluding there is nothing to close."""
    try:
        _handle_dialogs()
        end = time.monotonic() + appear_wait
        while not _office_windows() and time.monotonic() < end:
            time.sleep(0.5)
        if _office_windows():
            _focus(_office_windows()[0][0])
            engine.press("ctrl", "q")
            if not _wait_closed(10):
                if _dialog_windows():
                    _log("a dialog is still open; not pressing Enter blindly")
                else:
                    engine.press("enter")          # "Save changes?" -> Save
                    _wait_closed(10)
        _LEFT_OPEN["open"] = False
    except Exception as exc:
        _log("force close failed: %r" % (exc,))


def _launch_writer(extra_arg=None):
    """P1 convention: path already quoted, args as ONE string.
    Blank Writer/Calc open from the Start menu ("libreoffice writer", Enter), the
    way a person does it; Win+R with the full path is only the fallback. Opening a
    specific file (extra_arg is a path) still goes through Win+R."""
    quoted = '"%s"' % _soffice()
    args = extra_arg if extra_arg is not None else " ".join(WRITER_ARGS)
    app = {" ".join(WRITER_ARGS): "writer", " ".join(CALC_ARGS): "calc"}.get(args)
    if app is None:
        engine.launch(quoted, args)
        return
    brand = "openoffice" if "openoffice" in _soffice().lower() else "libreoffice"
    try:
        engine.open_app("%s %s" % (brand, app))
    except RuntimeError as exc:                 # Start menu never came up: nothing was typed
        _log("Start menu not available (%s); using Win+R" % exc)
        engine.launch(quoted, args)
        return
    # our own _wait sees any window (not only the foreground one) and handles the
    # recovery/tip dialogs, so it decides whether the Start-menu launch worked
    if _wait(WINDOW_PART if app == "writer" else CALC_PART, 40, dialogs=True):
        _log("opened %s from the Start menu" % app)
        return
    _log("Start-menu launch of %s not seen; falling back to Win+R" % app)
    engine.launch(quoted, args)


def _save_as(filename):
    """Ctrl+S on an untitled doc -> Save As dialog -> type name -> Enter."""
    engine.press("ctrl", "s")
    if not _wait("Save", 10):
        _log("Save As dialog title not seen; continuing blind")
    engine.human_delay("action")
    engine.type_text(filename)
    engine.human_delay("action")
    engine.press("enter")
    if not _wait(filename, 15):
        raise OfficeError("window title never showed %s after save" % filename)
    _log("saved as %s" % filename)


def _quick_save():
    engine.press("ctrl", "s")
    _pause(0.8, 0.4, 0.3, 3)


def _type_title(title):
    bold_ok = BOLD_TITLE
    if bold_ok:
        try:
            engine.press("ctrl", "b")
        except Exception as exc:
            _log("Ctrl+B unsupported (%r); skipping bold" % (exc,))
            bold_ok = False
    n = _type_human(title)
    if bold_ok:
        engine.press("ctrl", "b")
    engine.press("enter")
    engine.press("enter")
    return n


def write_document(max_seconds=None, doc_type=None, raise_on_error=False, keep_open=False):
    t0 = time.monotonic()
    deadline = (t0 + max_seconds) if max_seconds else None
    log_start = len(ACTION_LOG)
    dlg_start = len(DIALOGS_HANDLED)
    rec = {"scenario": "office.write_document", "ok": False, "error": None, "left_open": False,
           "doc_type": None, "filename": None, "path": None, "bytes": 0,
           "typos": 0, "saves": 0, "paragraphs": 0, "duration_s": 0.0,
           "dialogs": [], "log": []}
    spc = [0.25]          # seconds per typed char, refined as we type

    def _measure(elapsed, nchars):
        spc[0] = min(1.0, max(0.1, elapsed / max(nchars, 1)))

    def fits(para):
        """Typing this paragraph + wrap-up still inside the budget?"""
        if deadline is None:
            return True
        # 1.3x + 6 s allowance: hesitation/end-of-sentence pauses have a heavy tail
        need = 6.0 + 1.3 * len(para) * spc[0] + WRAP_UP_RESERVE
        return time.monotonic() + need <= deadline

    try:
        _ensure_closed()
        _log("launching Writer (Start menu, Win+R fallback)")
        _launch_writer()
        if not _wait(WINDOW_PART, 45, dialogs=True):
            raise OfficeError("Writer window did not appear")
        _pause(2.0, 0.4, 0.8, 6)                      # blank doc settles

        kind, title, paras, base = _compose(doc_type)
        fname = _filename(base)
        rec["doc_type"], rec["filename"] = kind, fname
        _log("doc=%s file=%s paragraphs=%d" % (kind, fname, len(paras)))

        t_ty = time.monotonic()
        rec["typos"] += _type_title(title)
        _measure(time.monotonic() - t_ty, len(title) + 2)

        save_after = random.choice([0, 1])            # first save after para 1/2
        saved = False
        for i, para in enumerate(paras):
            if i > 0 and not fits(para):
                _log("time budget: not enough left for paragraph %d; wrapping up" % (i + 1))
                break
            if i > 0 and random.random() < 0.25:
                _pause(5.0, 0.5, 2.0, 12.0)           # rereads what's there / looks away
            _delay("think", 2.2, 0.35, 1.2, 6.0)      # thinking what to write
            t_ty = time.monotonic()
            rec["typos"] += _type_human(para)
            _measure(time.monotonic() - t_ty, len(para))
            rec["paragraphs"] += 1
            engine.press("enter")
            if random.random() < 0.6:
                engine.press("enter")
            if not saved and i >= save_after:
                _save_as(fname)
                saved, rec["saves"] = True, rec["saves"] + 1
            elif saved and random.random() < 0.35:
                _quick_save()
                rec["saves"] += 1

        if not saved:
            _save_as(fname)
            saved, rec["saves"] = True, rec["saves"] + 1

        _pause(1.5, 0.5, 0.5, 6)
        engine.press("ctrl", "s")                     # final save
        rec["saves"] += 1
        _pause(1.2, 0.4, 0.5, 4)

        path = _find_saved(fname)
        if not path or os.path.getsize(path) == 0:
            raise OfficeError("saved file not found/empty: %s" % fname)
        rec["path"], rec["bytes"] = path, os.path.getsize(path)
        _LAST.update(path=path, name=fname, kind=kind)
        _log("verified on disk: %s (%d bytes)" % (path, rec["bytes"]))

        _pause(2.5, 0.5, 1.0, 8.0)                    # glance over the result (bounded: inside WRAP_UP_RESERVE)
        if keep_open:
            _LEFT_OPEN["open"] = True                 # caller must call close_writer() later
            rec["left_open"] = True
            _log("keep_open: Writer left open; call office.close_writer() when done")
        else:
            _close_writer()
        rec["ok"] = True
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _close_explorer()
        _force_close()
        if raise_on_error:
            raise
    finally:
        rec["duration_s"] = round(time.monotonic() - t0, 1)
        rec["dialogs"] = list(DIALOGS_HANDLED[dlg_start:])
        rec["log"] = ACTION_LOG[log_start:]
    return rec


# --------------------------------------------------------------------------
# Re-opening a saved document through File Explorer (what a person does)
# --------------------------------------------------------------------------
EDIT_VIA_EXPLORER = True      # staff 15 Oct: either way is fine. Chain is explorer -> swriter.exe -> soffice.exe
_EXPLORER_OPENED = []         # File Explorer windows this module opened (closed again after)
_EXPLORER_DIALOG_RE = re.compile(
    r"^(File Explorer|Open with|How do you want|Select an app|Windows cannot|Error)", re.I)


def _explorer_windows():
    """[(hwnd, title)] of File Explorer folder windows (window class CabinetWClass)."""
    if sys.platform != "win32":
        return []
    import ctypes
    user32 = ctypes.windll.user32
    out = []
    for h, t in _enum_windows():
        buf = ctypes.create_unicode_buffer(64)
        user32.GetClassNameW(h, buf, 64)
        if buf.value == "CabinetWClass":
            out.append((h, t))
    return out


def _post_close(hwnd):
    """WM_CLOSE to a window (no keystrokes, no new processes)."""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.PostMessageW(hwnd, 0x0010, 0, 0)
        except Exception:
            pass


def _close_explorer():
    for h in list(_EXPLORER_OPENED):
        _post_close(h)
    _EXPLORER_OPENED.clear()


def _dismiss_explorer_dialog():
    """Escape on an Explorer error / 'Open with' dialog, if one is showing.
    Returns its title or None."""
    mains = {h for h, _t in _explorer_windows()}
    for h, t in _enum_windows():
        if h not in mains and _EXPLORER_DIALOG_RE.search(t):
            _log("explorer dialog %r -> Escape" % (t,))
            _focus(h)
            time.sleep(0.3)
            engine.press("escape")
            time.sleep(0.5)
            return t
    return None


def _office_doc_open(name):
    """Is `name` open in an OpenOffice/LibreOffice window (not some other app)?"""
    if sys.platform != "win32":
        return True
    n = name.lower()
    return any(n in t.lower() and OFFICE_TITLE_RE.search(t) for _h, t in _enum_windows())


def _newest_document(exts=(".odt", ".ods")):
    best, best_t = None, -1
    for d in _save_dirs():
        try:
            for f in os.listdir(d):
                fp = os.path.join(d, f)
                if f.lower().endswith(exts) and os.path.isfile(fp) and os.path.getmtime(fp) > best_t:
                    best, best_t = fp, os.path.getmtime(fp)
        except OSError:
            continue
    return best


EXPLORER_PICK = "path"        # "path": address bar opens the exact file (reliable); "list": type-ahead in the folder list
_DOC_WIN_RE = re.compile(r" - (OpenOffice|LibreOffice) (Writer|Calc|Impress|Draw|Base|Math)", re.I)


def _wrong_documents(name):
    """Office document windows that are NOT `name` (e.g. Explorer opened the wrong file)."""
    n = name.lower()
    return [(h, t) for h, t in _enum_windows() if _DOC_WIN_RE.search(t) and n not in t.lower()]


def _wait_open(name, timeout):
    """Wait for `name` to open in Office.  Returns "ok", "wrong" (some other
    document opened; it has been closed again) or "none" (nothing opened)."""
    if sys.platform != "win32":
        return "ok" if (_wait(name, timeout, dialogs=True) and _office_doc_open(name)) else "none"
    end = time.monotonic() + timeout
    extended = 0
    while True:
        if len(_handle_dialogs()) and extended < 3:
            end += 15
            extended += 1
        if _office_doc_open(name):
            hwnd = _find_window(name)
            if hwnd:
                _focus(hwnd)
                time.sleep(0.3)
            return "ok"
        wrong = _wrong_documents(name)
        if wrong:
            _log("Explorer opened the wrong document (%r); closing it" % (wrong[0][1],))
            for h, _t in wrong:
                _focus(h)
                time.sleep(0.4)
                engine.press("ctrl", "q")
                time.sleep(0.5)
            if not _wait_closed(15):
                raise OfficeError("a wrong document opened from Explorer and would not close")
            _pause(1.5, 0.4, 0.8, 4)                  # soffice.bin finishes exiting
            return "wrong"
        if time.monotonic() >= end:
            return "none"
        time.sleep(0.5)


def _open_via_explorer(path):
    """Win+E, go to the document's folder, open the file.  EXPLORER_PICK="path"
    types the full path in the address bar (always the right file); "list" picks
    it by typing its name in the folder list, which can select the wrong file,
    so a wrong document is closed again and the path route is used.  Returns
    "explorer-path" or "explorer-list"; raises OfficeError if Explorer or the
    document never appears (the caller then falls back to Win+R)."""
    folder, name = os.path.split(path)
    stem = os.path.splitext(name)[0]
    before = {h for h, _t in _explorer_windows()}
    _log("opening File Explorer (Win+E) to reach %s" % name)
    engine.press("win", "e")
    new, end = None, time.monotonic() + 12
    while time.monotonic() < end:
        cur = [h for h, _t in _explorer_windows() if h not in before]
        if cur:
            new = cur[0]
            break
        time.sleep(0.5)
    if new is None:
        raise OfficeError("File Explorer window did not appear")
    _EXPLORER_OPENED.append(new)
    _focus(new)
    _pause(1.6, 0.4, 0.7, 5)                          # looks at the window
    engine.press("ctrl", "l")                         # address bar
    _pause(0.7, 0.4, 0.3, 2.5)
    engine.type_text(folder)
    _pause(0.6, 0.4, 0.3, 2.5)
    engine.press("enter")
    leaf = os.path.basename(folder).lower()
    end = time.monotonic() + 10
    while time.monotonic() < end:
        if any(leaf in t.lower() for _h, t in _explorer_windows()):
            break
        time.sleep(0.5)
    else:
        _log("Explorer title never showed %r; continuing" % leaf)
    _pause(1.4, 0.5, 0.5, 5)                          # looks at the folder contents

    if EXPLORER_PICK == "list":
        _focus(new)
        engine.type_text(stem)                        # type-ahead selects the file
        _pause(0.9, 0.4, 0.4, 3)
        engine.press("enter")
        t_open = time.monotonic()
        # Calc takes about 30 s to cold-start, so wait as long as a normal launch
        res = _wait_open(name, 45)
        if res == "ok":
            return "explorer-list"
        _log("file did not open from the list (%s, %.0f s); using the address bar"
             % (res, time.monotonic() - t_open))
        _dismiss_explorer_dialog()

    _focus(new)
    engine.press("ctrl", "l")
    _pause(0.7, 0.4, 0.3, 2.5)
    engine.type_text(path)
    _pause(0.6, 0.4, 0.3, 2.5)
    engine.press("enter")
    res = _wait_open(name, 45)
    if res == "ok":
        return "explorer-path"
    _dismiss_explorer_dialog()
    raise OfficeError("the document did not open from File Explorer (%s)" % res)


def explorer_test(path=None):
    """Open a saved document through File Explorer, check it opened in Office,
    close it without editing."""
    log_start = len(ACTION_LOG)
    rec = {"scenario": "office.explorer_test", "ok": False, "error": None,
           "path": None, "method": None, "log": []}
    try:
        path = path or _LAST["path"] or _newest_document()
        if not path or not os.path.exists(path):
            raise OfficeError("no saved .odt/.ods found; run --quick or --calc first")
        rec["path"] = path
        _ensure_closed()
        rec["method"] = _open_via_explorer(path)
        _pause(2.5, 0.4, 1.0, 6)                      # glance at the opened document
        _close_explorer()
        _close_writer()
        rec["ok"] = True
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _close_explorer()
        _force_close()
    finally:
        rec["log"] = ACTION_LOG[log_start:]
    return rec


def edit_document(path=None, max_seconds=None, raise_on_error=False, via_explorer=None):
    """Re-open the last saved document, add a paragraph, save, quit."""
    t0 = time.monotonic()
    log_start = len(ACTION_LOG)
    dlg_start = len(DIALOGS_HANDLED)
    rec = {"scenario": "office.edit_document", "ok": False, "error": None,
           "path": None, "typos": 0, "duration_s": 0.0, "dialogs": [], "log": [],
           "open_method": None}
    try:
        path = path or _LAST["path"]
        if not path or not os.path.exists(path):
            raise OfficeError("no saved document to edit")
        rec["path"] = path
        name = os.path.basename(path)
        before = os.path.getsize(path)
        _ensure_closed()
        use_explorer = EDIT_VIA_EXPLORER if via_explorer is None else via_explorer
        method = None
        if use_explorer:
            try:
                method = _open_via_explorer(path)
            except Exception as exc:
                _log("File Explorer route failed (%r); falling back to Win+R" % (exc,))
                _close_explorer()
                _ensure_closed()
        if method is None:
            _log("re-opening %s via engine.launch (Win+R)" % name)
            _launch_writer('"%s"' % path)
            if not _wait(name, 45, dialogs=True):
                raise OfficeError("document window did not appear")
            method = "win+r"
        rec["open_method"] = method
        _close_explorer()                             # tidy: the folder window is done with
        engine.human_delay("read")                    # re-read what is there

        try:
            engine.press("ctrl", "end")               # jump to end of document
        except Exception as exc:
            _log("Ctrl+End unsupported (%r); typing at top instead" % (exc,))
        _pause(1.0, 0.4, 0.3, 4)
        engine.press("enter")
        for sent in random.sample(_additions(), k=random.randint(1, 2)):
            _delay("think", 2.2, 0.35, 1.2, 6.0)
            rec["typos"] += _type_human(sent)
            engine.type_text(" ")
        _pause(1.2, 0.5, 0.4, 5)
        engine.press("ctrl", "s")
        _pause(1.5, 0.4, 0.5, 4)
        after = os.path.getsize(path)
        _log("size %d -> %d bytes" % (before, after))
        engine.human_delay("read")
        _close_writer()
        rec["ok"] = True
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _force_close()
        if raise_on_error:
            raise
    finally:
        rec["duration_s"] = round(time.monotonic() - t0, 1)
        rec["dialogs"] = list(DIALOGS_HANDLED[dlg_start:])
        rec["log"] = ACTION_LOG[log_start:]
    return rec


# --------------------------------------------------------------------------
# Key test: proves backspace / home / end and ( ) : _ , digits on the real VM
# --------------------------------------------------------------------------
KEYTEST_EXPECTED = "[A] Key test (10 Oct): file_name_v2, Ready 42. [B]"


def _odt_text(path):
    """Plain text of a saved .odt, read straight from content.xml (stdlib)."""
    with zipfile.ZipFile(path) as z:
        xml = z.read("content.xml").decode("utf-8", "replace")
    xml = re.sub(r"<text:tab\s*/>", "\t", xml)
    xml = re.sub(r"<text:s\b[^>]*/>", " ", xml)
    xml = xml.replace("</text:p>", "\n").replace("</text:h>", "\n")
    return html.unescape(re.sub(r"<[^>]+>", "", xml))


def keytest():
    """
    Types "Key tezt", Backspaces 3, completes the line with ( ) : _ , . digits,
    uses Home/End to add [A]/[B], saves, closes, then reads the text back out
    of the .odt and compares.  rec["got"] shows what actually landed.
    """
    log_start = len(ACTION_LOG)
    rec = {"scenario": "office.keytest", "ok": False, "error": None,
           "expected": KEYTEST_EXPECTED, "got": None, "file": None, "log": []}
    try:
        _ensure_closed()
        _log("keytest: launching Writer")
        _launch_writer()
        if not _wait(WINDOW_PART, 45, dialogs=True):
            raise OfficeError("Writer window did not appear")
        _pause(2.0, 0.4, 0.8, 6)
        engine.type_text("Key tezt")
        for _ in range(3):
            engine.press("backspace")
            engine.human_delay("keystroke")
        engine.type_text("est (10 Oct): file_name_v2, Ready 42.")
        engine.press("home")
        engine.type_text("[A] ")
        engine.press("end")
        engine.type_text(" [B]")
        fname = "keytest_%s.odt" % time.strftime("%H%M%S")
        _save_as(fname)
        _pause(1.0, 0.4, 0.5, 3)
        engine.press("ctrl", "s")
        _pause(1.2, 0.4, 0.5, 4)
        path = _find_saved(fname)
        if not path:
            raise OfficeError("keytest file not found: %s" % fname)
        rec["file"] = path
        _close_writer()
        rec["got"] = _odt_text(path).strip()
        rec["ok"] = (rec["got"] == KEYTEST_EXPECTED)
        if not rec["ok"]:
            rec["error"] = "typed text differs from expected (see 'got')"
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _force_close()
    finally:
        rec["log"] = ACTION_LOG[log_start:]
    return rec


# --------------------------------------------------------------------------
# Calc (spreadsheet) scenario  -- C1
# --------------------------------------------------------------------------
CALC_ARGS = ["-calc"]
CALC_PART = "Office Calc"      # "OpenOffice Calc" / "LibreOffice Calc"; NOT "Calculator"
SHEET_EXT = ".ods"
_TAB_MODE = {"mode": "press"}
FOOTER_LABEL = "Total"

# (name, (qty lo, hi), (price lo, hi)).  Row labels are picked with DISTINCT
# first letters (and not the header's or the footer's), so Calc's AutoInput
# can never complete a typed label from an earlier entry in the column.
SHEET_TYPES = {
    "budget": {
        "headers": ["Item", "Qty", "Unit price", "Total"],
        "bases": ["Budget", "Office_Budget", "Costs"],
        "ints": True,
        "items": [("Paper reams", (2, 12), (3, 9)), ("Stationery", (1, 6), (8, 30)),
                  ("Cleaning supplies", (1, 4), (6, 25)), ("Parking", (2, 10), (4, 12)),
                  ("Equipment hire", (1, 3), (40, 180)), ("Water dispenser", (1, 2), (20, 60)),
                  ("Books", (2, 8), (9, 35)), ("Laptop stands", (2, 6), (15, 45)),
                  ("Hand sanitiser", (3, 12), (2, 7)), ("Name tags", (10, 50), (1, 3)),
                  ("Decorations", (1, 5), (6, 28)), ("Folders", (5, 30), (1, 4)),
                  ("Markers", (4, 20), (1, 5)), ("Gift cards", (2, 10), (10, 50)),
                  ("Refreshments", (2, 8), (12, 45)), ("Adapters", (1, 5), (8, 25)),
                  ("Keyboards", (1, 4), (18, 60)), ("Office chairs", (1, 3), (60, 190)),
                  ("Visitor passes", (5, 25), (1, 4)), ("Postage", (3, 15), (1, 6))],
    },
    "inventory": {
        "headers": ["Product", "Units", "Unit cost", "Value"],
        "bases": ["Inventory", "Stock_List", "Supplies"],
        "ints": True,
        "items": [("Notebooks", (10, 80), (1, 6)), ("Envelopes", (20, 200), (1, 3)),
                  ("Staplers", (2, 15), (4, 15)), ("Folders", (10, 90), (1, 4)),
                  ("Markers", (10, 60), (1, 5)), ("Highlighters", (10, 50), (1, 4)),
                  ("Batteries", (8, 60), (2, 10)), ("Labels", (20, 150), (1, 5)),
                  ("Desk lamps", (2, 10), (12, 40)), ("Cables", (5, 40), (3, 14)),
                  ("Whiteboards", (1, 6), (25, 90)), ("Sticky notes", (20, 120), (1, 4)),
                  ("Binders", (10, 60), (2, 8)), ("Clipboards", (5, 30), (2, 9)),
                  ("Erasers", (10, 100), (1, 2)), ("Rulers", (10, 70), (1, 3)),
                  ("Scissors", (4, 25), (2, 8)), ("Glue sticks", (10, 80), (1, 3)),
                  ("Pencils", (20, 200), (1, 2)), ("Calculators", (2, 12), (6, 22))],
    },
    "timesheet": {
        "headers": ["Task", "Hours", "Rate", "Pay"],
        "bases": ["Timesheet", "Hours_Log", "Work_Log"],
        "ints": False,
        "items": [("Planning", (1, 6), (20, 40)), ("Documentation", (1, 8), (20, 40)),
                  ("Review", (1, 5), (20, 40)), ("Design", (2, 8), (22, 45)),
                  ("Meetings", (1, 4), (20, 40)), ("Reporting", (1, 5), (20, 40)),
                  ("Research", (2, 8), (22, 45)), ("Support", (1, 6), (18, 35)),
                  ("Admin", (1, 4), (18, 30)), ("Coding", (2, 8), (25, 50)),
                  ("Deployment", (1, 4), (25, 50)), ("Analysis", (2, 7), (24, 48)),
                  ("Briefing", (1, 3), (20, 40)), ("Filing", (1, 3), (16, 28)),
                  ("Scheduling", (1, 3), (18, 32)), ("Onboarding", (1, 5), (20, 36)),
                  ("Proofreading", (1, 4), (18, 34)), ("Budgeting", (1, 5), (22, 42)),
                  ("Workshops", (2, 6), (22, 42)), ("Interviews", (1, 5), (22, 42))],
    },
}


def _sheet_plan(sheet_type=None):
    kind = sheet_type if sheet_type in SHEET_TYPES else random.choice(list(SHEET_TYPES))
    spec = SHEET_TYPES[kind]
    used = {spec["headers"][0][0].lower(), FOOTER_LABEL[0].lower()}
    pool = list(spec["items"])
    random.shuffle(pool)
    rows = []
    for name, (qlo, qhi), (plo, phi) in pool:
        ch = name[0].lower()
        if ch in used:
            continue
        used.add(ch)
        if spec["ints"]:
            qty = random.randint(qlo, qhi)
        else:
            qty = random.choice([x / 2 for x in range(qlo * 2, qhi * 2 + 1)])
        price = round(random.randint(int(plo * 20), int(phi * 20)) / 20, 2)
        rows.append((name, qty, price))
        if len(rows) >= random.randint(5, 8):
            break
    return {"kind": kind, "headers": list(spec["headers"]), "base": random.choice(spec["bases"]),
            "rows": rows}


def _fmt_num(x):
    return ("%d" % x) if float(x).is_integer() else ("%g" % x if abs(x) < 1e6 else repr(x))


def _tab(human=False):
    """Tab between cells.  Falls back to typing a tab character if the engine
    has no 'tab' key.  (The read-back check catches it if neither works.)"""
    if human:
        _pause(0.3, 0.5, 0.1, 2.0)                  # stepping over an empty cell
    if _TAB_MODE["mode"] == "press":
        try:
            engine.press("tab")
            return
        except Exception as exc:
            _log("engine has no 'tab' key (%r); typing a tab character instead" % (exc,))
            _TAB_MODE["mode"] = "type"
    engine.type_text("\t")


NUM_TYPO_RATE = 0.06      # chance a typed number gets a wrong digit that is fixed


def _type_number(txt, human=True):
    """Type a number; now and then a wrong digit slips in and is backspaced
    (never touches AutoInput: numbers are not auto-completed).  Returns typos."""
    global TYPOS_ENABLED
    digits = [i for i, ch in enumerate(txt) if ch.isdigit()]
    if (human and TYPOS_ENABLED and len(digits) >= 2 and NUM_TYPO_RATE > 0
            and random.random() < NUM_TYPO_RATE * (TYPO_RATE / 0.05)):
        i = random.choice(digits[1:])                 # never the first digit
        wrong = str((int(txt[i]) + random.choice([-1, 1, 2, -2])) % 10)
        k = random.choice([0, 0, 1, 2])               # chars typed before noticing
        j = min(len(txt), i + 1 + k)
        typed = txt[:i] + wrong + txt[i + 1:j]
        engine.type_text(typed)
        _delay("notice", 0.5, 0.35, 0.3, 1.5)
        try:
            for _ in range(len(typed) - i):
                engine.press("backspace")
                engine.human_delay("keystroke")
        except Exception as exc:
            _log("backspace failed (%r); typo model disabled" % (exc,))
            TYPOS_ENABLED = False
            raise
        engine.type_text(txt[i:])
        return 1
    engine.type_text(txt)
    return 0


def _type_cells(cells, human=True):
    """cells: [(text, kind)] with kind text|num|formula.  Tab between, Enter at
    the end (Calc returns to the first column of the next row).  Returns typos.
    human=True adds the pauses a person has between cells (moving from field to
    field, glancing at a number, remembering the formula); human=False is the
    fast functional mode used by --calc-keytest."""
    typos = 0
    for i, (txt, kind) in enumerate(cells):
        if i > 0:
            if human:
                if kind == "formula":
                    _pause(0.9, 0.5, 0.3, 4.0)       # "what was the formula again"
                else:
                    _pause(0.45, 0.55, 0.15, 3.0)    # hand moves to the next field
            _tab()
            if human:
                engine.human_delay("keystroke")
        if kind == "text":
            typos += _type_human(txt) if human else (engine.type_text(txt) or 0)
        elif kind == "num":
            typos += _type_number(txt, human)
        else:
            engine.type_text(txt)
        if human and kind != "text" and random.random() < 0.25:
            _pause(0.8, 0.5, 0.3, 4.0)               # glance at what was typed
    if human:
        _pause(0.5, 0.5, 0.15, 3.0)                  # look before committing the row
    engine.press("enter")
    return typos


def _ods_grid(path):
    """First table of a saved .ods as rows of (value, formula); stdlib only."""
    import xml.etree.ElementTree as ET
    T = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
    O = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("content.xml"))
    rows = []
    table = next(root.iter(T + "table"), None)
    if table is None:
        return rows
    for tr in table.iter(T + "table-row"):
        cells = []
        for tc in tr:
            if tc.tag != T + "table-cell":
                continue
            reps = min(int(tc.get(T + "number-columns-repeated", "1")), 40)
            vt = tc.get(O + "value-type")
            if vt in ("float", "percentage", "currency"):
                val = float(tc.get(O + "value"))
            elif vt == "string":
                val = "".join("".join(p.itertext()) for p in tc if p.tag.endswith("}p"))
            else:
                val = None
            cells += [(val, tc.get(T + "formula"))] * reps
        while cells and cells[-1] == (None, None):
            cells.pop()
        reps_r = min(int(tr.get(T + "number-rows-repeated", "1")), 3)
        rows += [cells] * (reps_r if cells else 1)
    while rows and not rows[-1]:
        rows.pop()
    return rows


def _close(a, b, tol=0.006):
    return a is not None and b is not None and abs(a - b) <= tol


def _check_grid(grid, plan):
    """(ok, note): does the saved sheet hold exactly what was meant to be typed?"""
    def cell(r, c):
        try:
            return grid[r][c]
        except IndexError:
            return (None, None)

    for c, h in enumerate(plan["headers"]):
        if cell(0, c)[0] != h:
            return False, "header %d is %r, expected %r" % (c, cell(0, c)[0], h)
    total = 0.0
    for i, (name, qty, price) in enumerate(plan["rows"]):
        r = i + 1
        if cell(r, 0)[0] != name:
            return False, "row %d label is %r, expected %r" % (r + 1, cell(r, 0)[0], name)
        if not _close(cell(r, 1)[0], qty) or not _close(cell(r, 2)[0], price):
            return False, "row %d numbers are %r/%r, expected %r/%r" % (
                r + 1, cell(r, 1)[0], cell(r, 2)[0], qty, price)
        v, f = cell(r, 3)
        if not f or not _close(v, qty * price):
            return False, "row %d formula/value is %r/%r, expected %r" % (r + 1, f, v, qty * price)
        total += qty * price
    fr = len(plan["rows"]) + 1
    if plan.get("footer", True):
        if cell(fr, 0)[0] != FOOTER_LABEL:
            return False, "footer label is %r" % (cell(fr, 0)[0],)
        v, f = cell(fr, 3)
        if not f or "SUM" not in f.upper() or not _close(v, total):
            return False, "footer formula/value is %r/%r, expected SUM = %r" % (f, v, total)
    return True, "all %d rows, formulas and the SUM match" % len(plan["rows"])


CALC_BOLD_HEADER = True       # after typing: select the header row, Ctrl+B
CALC_REVISE_VALUE = True      # after typing: go back and change one quantity
CALC_REALISM_P = 0.75         # chance each of the two actions happens in a run
_NAV = {"ok": True}           # cleared if the engine lacks an arrow key


def _nav(key, n=1):
    """Press an arrow key n times with key-repeat style pauses.  Raises if the
    engine has no such key (callers catch it and skip the action)."""
    for _ in range(n):
        engine.press(*key) if isinstance(key, tuple) else engine.press(key)
        _pause(0.2, 0.4, 0.07, 0.9)


def _ods_header_bold(path, ncols):
    """True/False: are the first `ncols` cells of row 1 bold?  None if unreadable."""
    import xml.etree.ElementTree as ET
    T = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
    S = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
    FO = "{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}"
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("content.xml"))
        bold = {}
        for st in root.iter(S + "style"):
            tp = st.find(S + "text-properties")
            bold[st.get(S + "name")] = bool(tp is not None and tp.get(FO + "font-weight") == "bold")
        table = next(root.iter(T + "table"))
        row = next(table.iter(T + "table-row"))
        flags = []
        for tc in row:
            if tc.tag != T + "table-cell":
                continue
            reps = min(int(tc.get(T + "number-columns-repeated", "1")), 40)
            flags += [bold.get(tc.get(T + "style-name"), False)] * reps
        flags = flags[:ncols]
        return len(flags) == ncols and all(flags)
    except Exception:
        return None


def _realism_pass(plan, typed, rec, spare_ok):
    """What a person does after the data is in: tweak one quantity (the row's
    total and the SUM update), bold the header.  Every step is optional and
    never fails the run; the plan is updated so the read-back still matches."""
    if not _NAV["ok"] or not typed:
        return
    try:
        if CALC_REVISE_VALUE and random.random() < CALC_REALISM_P and spare_ok(25):
            k = random.randrange(len(typed))               # index into data rows
            name, qty, price = typed[k]
            if float(qty).is_integer():
                new = max(1, int(qty) + random.choice([-3, -2, -1, 1, 2, 3, 5]))
                if new == int(qty):
                    new += 1
            else:
                new = max(0.5, qty + random.choice([-1.5, -1, -0.5, 0.5, 1, 1.5]))
                if new == qty:
                    new += 0.5
            _delay("think", 2.4, 0.4, 1.0, 7.0)
            _log("revising %s quantity %s -> %s" % (name, _fmt_num(qty), _fmt_num(new)))
            engine.press("ctrl", "home")
            _pause(0.5, 0.4, 0.2, 2.0)
            _nav("down", k + 1)
            _nav("right", 1)
            _pause(0.6, 0.5, 0.2, 2.5)                    # check it is the right cell
            engine.type_text(_fmt_num(new))
            _pause(0.5, 0.5, 0.15, 2.5)
            engine.press("enter")
            typed[k] = (name, new, price)
            plan["rows"] = typed
            rec["revised"] = True
        if CALC_BOLD_HEADER and random.random() < CALC_REALISM_P and spare_ok(15):
            _delay("think", 2.0, 0.4, 0.8, 6.0)
            engine.press("ctrl", "home")
            _pause(0.5, 0.4, 0.2, 2.0)
            for _ in range(len(plan["headers"]) - 1):
                engine.press("shift", "right")
                _pause(0.18, 0.4, 0.06, 0.8)
            engine.press("ctrl", "b")
            _pause(0.6, 0.5, 0.2, 2.5)
            engine.press("ctrl", "home")                   # clears the selection
            rec["bold_header_done"] = True
    except Exception as exc:
        _NAV["ok"] = False
        _log("realism step skipped (%r); navigation keys unavailable" % (exc,))


def write_spreadsheet(max_seconds=None, sheet_type=None, raise_on_error=False):
    t0 = time.monotonic()
    deadline = (t0 + max_seconds) if max_seconds else None
    log_start = len(ACTION_LOG)
    dlg_start = len(DIALOGS_HANDLED)
    rec = {"scenario": "office.write_spreadsheet", "ok": False, "error": None,
           "sheet_type": None, "filename": None, "path": None, "bytes": 0,
           "rows": 0, "typos": 0, "saves": 0, "verified": None, "verify_note": None,
           "revised": False, "bold_header_done": False, "bold_header_verified": None,
           "duration_s": 0.0, "dialogs": [], "log": []}
    spc = [0.2]

    def fits(chars):
        if deadline is None:
            return True
        need = 6.0 + 1.3 * chars * spc[0] + 3.0 + WRAP_UP_RESERVE   # row + think pause + wrap-up
        return time.monotonic() + need <= deadline

    def spare_ok(seconds):
        """Enough budget left for an optional action of about `seconds`?"""
        if deadline is None:
            return True
        return time.monotonic() + seconds + 8.0 + WRAP_UP_RESERVE <= deadline

    try:
        _ensure_closed()
        _log("launching Calc (Start menu, Win+R fallback)")
        _launch_writer(" ".join(CALC_ARGS))
        if not _wait(CALC_PART, 45, dialogs=True):
            raise OfficeError("Calc window did not appear")
        _pause(2.0, 0.4, 0.8, 6)                          # blank sheet settles

        plan = _sheet_plan(sheet_type)
        fname = _filename(plan["base"], SHEET_EXT)
        rec["sheet_type"], rec["filename"] = plan["kind"], fname
        _log("sheet=%s file=%s planned rows=%d" % (plan["kind"], fname, len(plan["rows"])))

        t_ty = time.monotonic()
        n_chars = sum(len(h) for h in plan["headers"])
        rec["typos"] += _type_cells([(h, "text") for h in plan["headers"]])
        spc[0] = min(1.0, max(0.1, (time.monotonic() - t_ty) / max(n_chars, 1)))

        save_after = random.choice([1, 2])
        saved, typed = False, []
        for i, (name, qty, price) in enumerate(plan["rows"]):
            r = i + 2                                      # sheet row number (header = 1)
            chars = len(name) + len(_fmt_num(qty)) + len("%.2f" % price) + 8
            if i > 0 and not fits(chars):
                _log("time budget: not enough left for row %d; wrapping up" % r)
                break
            _delay("think", 2.6, 0.5, 1.0, 9.0)
            t_ty = time.monotonic()
            rec["typos"] += _type_cells([(name, "text"), (_fmt_num(qty), "num"),
                                         ("%.2f" % price, "num"),
                                         ("=B%d*C%d" % (r, r), "formula")])
            spc[0] = min(1.0, max(0.1, (time.monotonic() - t_ty) / max(chars, 1)))
            typed.append((name, qty, price))
            rec["rows"] += 1
            if not saved and len(typed) >= save_after:
                _save_as(fname)
                saved, rec["saves"] = True, rec["saves"] + 1
            elif saved and random.random() < 0.35:
                _quick_save()
                rec["saves"] += 1
        plan["rows"] = typed

        # footer: label, three Tabs to column D, SUM formula
        _delay("think", 1.8, 0.4, 0.8, 5.0)
        engine.type_text(FOOTER_LABEL)                     # no typos: AutoInput safety
        for _ in range(3):
            _tab(human=True)
        _pause(0.9, 0.5, 0.3, 4.0)
        engine.type_text("=SUM(D2:D%d)" % (len(typed) + 1))
        _pause(0.6, 0.5, 0.2, 3.0)
        engine.press("enter")

        if not saved:
            _save_as(fname)
            saved, rec["saves"] = True, rec["saves"] + 1
        _pause(1.5, 0.5, 0.5, 6)
        _realism_pass(plan, typed, rec, spare_ok)          # tweak a value, bold the header
        _pause(1.0, 0.5, 0.3, 4)
        engine.press("ctrl", "s")                          # final save
        rec["saves"] += 1
        _pause(1.2, 0.4, 0.5, 4)

        path = _find_saved(fname)
        if not path or os.path.getsize(path) == 0:
            raise OfficeError("saved file not found/empty: %s" % fname)
        rec["path"], rec["bytes"] = path, os.path.getsize(path)
        _LAST.update(path=path, name=fname, kind="sheet:" + plan["kind"])
        if rec["bold_header_done"]:
            rec["bold_header_verified"] = _ods_header_bold(path, len(plan["headers"]))
        try:
            ok, note = _check_grid(_ods_grid(path), plan)
        except Exception as exc:
            ok, note = False, "could not read the saved sheet back (%r)" % (exc,)
        rec["verified"], rec["verify_note"] = ok, note
        _log("verified on disk: %s (%d bytes); read-back %s: %s"
             % (path, rec["bytes"], "OK" if ok else "MISMATCH", note))

        _pause(2.5, 0.5, 1.0, 8.0)                         # glance over the sheet
        _close_writer()
        rec["ok"] = True
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _force_close()
        if raise_on_error:
            raise
    finally:
        rec["duration_s"] = round(time.monotonic() - t0, 1)
        rec["dialogs"] = list(DIALOGS_HANDLED[dlg_start:])
        rec["log"] = ACTION_LOG[log_start:]
    return rec


def calc_keytest():
    """Types a tiny sheet with Tab/Enter and two formulas, saves, closes, reads
    the .ods back and compares.  Proves Tab (or the typed-tab fallback), Enter
    auto-return and formula entry on the real VM."""
    log_start = len(ACTION_LOG)
    plan = {"kind": "keytest", "headers": ["Item", "Qty", "Price", "Total"],
            "rows": [("Pens", 3, 1.5), ("Tape", 2, 4.25)], "footer": True}
    rec = {"scenario": "office.calc_keytest", "ok": False, "error": None, "note": None,
           "file": None, "tab_mode": None, "log": []}
    try:
        _ensure_closed()
        _log("calc-keytest: launching Calc")
        _launch_writer(" ".join(CALC_ARGS))
        if not _wait(CALC_PART, 45, dialogs=True):
            raise OfficeError("Calc window did not appear")
        _pause(2.0, 0.4, 0.8, 6)
        _type_cells([(h, "text") for h in plan["headers"]], human=False)   # functional check: fast on purpose
        for i, (name, qty, price) in enumerate(plan["rows"]):
            r = i + 2
            _type_cells([(name, "text"), (_fmt_num(qty), "num"), ("%.2f" % price, "num"),
                         ("=B%d*C%d" % (r, r), "formula")], human=False)
        engine.type_text(FOOTER_LABEL)
        for _ in range(3):
            _tab()
        engine.type_text("=SUM(D2:D3)")
        engine.press("enter")
        fname = "calckeytest_%s.ods" % time.strftime("%H%M%S")
        _save_as(fname)
        _pause(1.0, 0.4, 0.5, 3)
        engine.press("ctrl", "s")
        _pause(1.2, 0.4, 0.5, 4)
        path = _find_saved(fname)
        if not path:
            raise OfficeError("calc keytest file not found: %s" % fname)
        rec["file"], rec["tab_mode"] = path, _TAB_MODE["mode"]
        _close_writer()
        grid = _ods_grid(path)
        rec["ok"], rec["note"] = _check_grid(grid, plan)
        if not rec["ok"]:
            rec["error"] = rec["note"]
            rec["grid"] = [[c[0] for c in row] for row in grid]
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _force_close()
    finally:
        rec["log"] = ACTION_LOG[log_start:]
    return rec


def calc_navtest():
    """Arrow-key navigation + selection + bold, checked by reading the .ods back:
    types a 3x2 grid, selects the header with Shift+Right and bolds it, goes back
    with Ctrl+Home / Down / Right and overwrites one cell."""
    log_start = len(ACTION_LOG)
    rec = {"scenario": "office.calc_navtest", "ok": False, "error": None, "note": None,
           "missing_keys": [], "nav_ok": None, "bold_ok": None, "file": None, "log": []}
    try:
        _ensure_closed()
        _log("calc-navtest: launching Calc")
        _launch_writer(" ".join(CALC_ARGS))
        if not _wait(CALC_PART, 45, dialogs=True):
            raise OfficeError("Calc window did not appear")
        _pause(2.0, 0.4, 0.8, 6)
        _type_cells([("Aa", "text"), ("Bb", "text")], human=False)
        _type_cells([("1", "num"), ("2", "num")], human=False)
        _type_cells([("3", "num"), ("4", "num")], human=False)

        def tryp(label, *keys):
            try:
                engine.press(*keys)
                _pause(0.25, 0.3, 0.1, 1.0)
                return True
            except Exception as exc:
                rec["missing_keys"].append(label)
                _log("key %s failed: %r" % (label, exc))
                return False

        if tryp("ctrl+home", "ctrl", "home") and tryp("shift+right", "shift", "right"):
            tryp("ctrl+b", "ctrl", "b")
        tryp("ctrl+home", "ctrl", "home")
        if tryp("down", "down") and tryp("right", "right"):
            engine.type_text("9")
            engine.press("enter")
        fname = "calcnav_%s.ods" % time.strftime("%H%M%S")
        _save_as(fname)
        _pause(1.0, 0.4, 0.5, 3)
        engine.press("ctrl", "s")
        _pause(1.2, 0.4, 0.5, 4)
        path = _find_saved(fname)
        if not path:
            raise OfficeError("calc navtest file not found: %s" % fname)
        rec["file"] = path
        _close_writer()
        grid = _ods_grid(path)
        vals = [[c[0] for c in row] for row in grid]
        rec["grid"] = vals
        rec["nav_ok"] = vals == [["Aa", "Bb"], [1.0, 9.0], [3.0, 4.0]]
        rec["bold_ok"] = _ods_header_bold(path, 2)
        rec["ok"] = bool(rec["nav_ok"] and rec["bold_ok"])
        rec["note"] = "navigation %s, bold header %s" % (
            "OK" if rec["nav_ok"] else "WRONG (grid %r)" % (vals,),
            "OK" if rec["bold_ok"] else ("not bold" if rec["bold_ok"] is False else "unreadable"))
        if not rec["ok"]:
            rec["error"] = rec["note"]
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _force_close()
    finally:
        rec["log"] = ACTION_LOG[log_start:]
    return rec


# --------------------------------------------------------------------------
# Dialog tools for the VM:  --dialog-keys  and  --probe
# --------------------------------------------------------------------------
def dialog_keys():
    """Does the engine map Escape and Alt+F4?  Escape on a blank Writer does
    nothing; Alt+F4 on an untitled empty document closes Writer without a prompt."""
    log_start = len(ACTION_LOG)
    rec = {"scenario": "office.dialog_keys", "ok": False, "error": None,
           "escape": None, "alt_f4": None, "log": []}
    try:
        _ensure_closed()
        _log("dialog-keys: launching Writer")
        _launch_writer()
        if not _wait(WINDOW_PART, 45, dialogs=True):
            raise OfficeError("Writer window did not appear")
        _pause(2.0, 0.4, 0.8, 6)
        try:
            engine.press("escape")
            rec["escape"] = True
        except Exception as exc:
            rec["escape"] = False
            _log("escape failed: %r" % (exc,))
        _pause(1.0, 0.3, 0.4, 3)
        try:
            engine.press("alt", "f4")
            rec["alt_f4"] = bool(_wait_closed(10))
            if not rec["alt_f4"]:
                _log("Alt+F4 sent but Writer is still open")
        except Exception as exc:
            rec["alt_f4"] = False
            _log("alt+f4 failed: %r" % (exc,))
        rec["ok"] = bool(rec["escape"] and rec["alt_f4"])
        if not rec["ok"]:
            rec["error"] = "a key is not mapped or had no effect (see log)"
            _force_close()
        else:
            _pause(1.2, 0.4, 0.6, 4)
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _force_close()
    finally:
        rec["log"] = ACTION_LOG[log_start:]
    return rec


def probe(seconds=40):
    """Cold-start Writer and log every Office-owned window (title + exe) as it
    appears; dismiss a known dialog if one shows up.  Use it to capture the
    exact title of the recovery window and to see what Escape does."""
    log_start = len(ACTION_LOG)
    dlg_start = len(DIALOGS_HANDLED)
    rec = {"scenario": "office.probe", "ok": False, "error": None,
           "writer_seen": False, "windows": [], "dialogs": [], "log": []}
    try:
        _ensure_closed()
        _log("probe: launching Writer")
        _launch_writer()
        seen, stable_since = set(), None
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            for h, t in _enum_windows():
                exe = _window_exe(h)
                if (exe in OFFICE_EXES or OFFICE_TITLE_RE.search(t)
                        or any(s["title"].search(t) for s in KNOWN_DIALOGS)):
                    key = (t, exe)
                    if key not in seen:
                        seen.add(key)
                        rec["windows"].append({"title": t, "exe": exe})
                        _log("window: %r  (exe=%s)" % (t, exe))
            if _handle_dialogs():
                stable_since = None
                continue
            if _find_window(WINDOW_PART):
                rec["writer_seen"] = True
                stable_since = stable_since or time.monotonic()
                if time.monotonic() - stable_since >= 3.0:
                    break
            time.sleep(0.5)
        rec["dialogs"] = list(DIALOGS_HANDLED[dlg_start:])
        if not rec["dialogs"]:
            _log("probe: no known dialog appeared (no recovery data present)")
        if not rec["writer_seen"]:
            raise OfficeError("Writer window never appeared during the probe")
        rec["ok"] = True
        _pause(1.0, 0.3, 0.4, 3)
        _close_writer()
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _force_close()
    finally:
        rec["log"] = ACTION_LOG[log_start:]
    return rec


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def keepopen_test(max_seconds=90):
    """write_document(keep_open=True) must leave Writer running; close_writer()
    must then close it.  Checks the real window list on Windows."""
    rec = {"scenario": "office.keepopen_test", "ok": False, "error": None,
           "open_after_write": None, "open_after_close": None, "log": []}
    log_start = len(ACTION_LOG)
    try:
        r = write_document(max_seconds=max_seconds, keep_open=True)
        if not r["ok"]:
            raise OfficeError("write_document failed: %s" % r["error"])
        win = sys.platform == "win32"
        rec["open_after_write"] = bool(_office_windows()) if win else bool(_LEFT_OPEN["open"])
        _log("keepopen-test: Writer open after write_document: %s" % rec["open_after_write"])
        _pause(4.0, 0.3, 2.0, 8.0)                    # Writer stays up for a few seconds
        close_writer()
        rec["open_after_close"] = bool(_office_windows()) if win else bool(_LEFT_OPEN["open"])
        _log("keepopen-test: Writer open after close_writer: %s" % rec["open_after_close"])
        rec["ok"] = bool(rec["open_after_write"]) and not rec["open_after_close"] and r["left_open"] is True
        if not rec["ok"]:
            rec["error"] = "open_after_write=%s open_after_close=%s left_open=%s" % (
                rec["open_after_write"], rec["open_after_close"], r["left_open"])
    except Exception as exc:
        rec["error"] = "%s: %s" % (type(exc).__name__, exc)
        _log("FAILED: " + rec["error"])
        _force_close()
    finally:
        rec["log"] = ACTION_LOG[log_start:]
    return rec


def _countdown(n=3):
    print("Click away from this terminal now. Starting in...")
    for i in range(n, 0, -1):
        print("  %d" % i, flush=True)
        time.sleep(1)


if __name__ == "__main__":
    argv = sys.argv[1:]
    _countdown(3)
    if "--calc-keytest" in argv:
        result = calc_keytest()
        print("\nCALC-KEYTEST", "PASS" if result["ok"] else "FAIL")
        print("tab mode :", result["tab_mode"], "(press = engine 'tab' key, type = typed tab char)")
        print("check    :", result["note"])
        if result["error"]:
            print("error    :", result["error"])
        if result.get("grid"):
            print("saved grid:", result["grid"])
        sys.exit(0 if result["ok"] else 1)
    if "--calc-navtest" in argv:
        result = calc_navtest()
        print("\nCALC-NAVTEST", "PASS" if result["ok"] else "FAIL")
        print("check        :", result["note"])
        print("missing keys :", result["missing_keys"] or "none")
        if result.get("grid") is not None:
            print("saved grid   :", result["grid"])
        if result["error"]:
            print("error        :", result["error"])
        sys.exit(0 if result["ok"] else 1)
    if "--calc" in argv or "--calc-quick" in argv:
        if "--calc-quick" in argv:
            TYPO_RATE = 0.4
            result = write_spreadsheet(max_seconds=90)
        else:
            result = write_spreadsheet(max_seconds=240)
        print({k: v for k, v in result.items() if k != "log"})
        sys.exit(0 if result["ok"] and result["verified"] is not False else 1)
    if "--keepopen-test" in argv:
        TYPO_RATE = 0.4
        result = keepopen_test()
        print("\nKEEPOPEN-TEST", "PASS" if result["ok"] else "FAIL")
        print("Writer open after write_document(keep_open=True):", result["open_after_write"], "(want True)")
        print("Writer open after close_writer()                :", result["open_after_close"], "(want False)")
        if result["error"]:
            print("error:", result["error"])
        sys.exit(0 if result["ok"] else 1)
    if "--explorer-test" in argv:
        result = explorer_test()
        print("\nEXPLORER-TEST", "PASS" if result["ok"] else "FAIL")
        print("file   :", result["path"])
        print("method :", result["method"], "(explorer-list = picked from the folder list, explorer-path = typed path)")
        if result["error"]:
            print("error  :", result["error"])
        sys.exit(0 if result["ok"] else 1)
    if "--reopen" in argv:
        newest = _newest_document((".odt",))
        result = edit_document(path=newest, via_explorer=(True if "--explorer" in argv else None))
        print({k: v for k, v in result.items() if k != "log"})
        sys.exit(0 if result["ok"] else 1)
    if "--dialog-keys" in argv:
        result = dialog_keys()
        print("\nDIALOG-KEYS", "PASS" if result["ok"] else "FAIL")
        print("escape mapped :", result["escape"])
        print("alt+f4 closes :", result["alt_f4"])
        if result["error"]:
            print("error         :", result["error"])
        sys.exit(0 if result["ok"] else 1)
    if "--probe" in argv:
        result = probe()
        print("\nPROBE", "OK" if result["ok"] else "FAIL")
        for w in result["windows"]:
            print("  window:", w)
        print("  dialogs handled:", result["dialogs"] or "none")
        if result["error"]:
            print("  error:", result["error"])
        sys.exit(0 if result["ok"] else 1)
    if "--keytest" in argv:
        result = keytest()
        print("\nKEYTEST", "PASS" if result["ok"] else "FAIL")
        print("expected:", result["expected"])
        print("got     :", result["got"])
        if result["error"]:
            print("error   :", result["error"])
        sys.exit(0 if result["ok"] else 1)
    if "--quick" in argv:
        TYPO_RATE = 0.4                    # make typos visible for a key check
        result = write_document(max_seconds=90)
    else:
        result = write_document(max_seconds=240)
    if "--edit" in argv and result["ok"]:
        _pause(3.0, 0.5, 1, 10)
        print({k: v for k, v in edit_document().items() if k != "log"})
    print({k: v for k, v in result.items() if k != "log"})