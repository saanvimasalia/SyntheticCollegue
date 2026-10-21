"""
test_new.py - try the new pieces one at a time instead of a full 15-min run.

    python test_new.py pauses              # timing numbers only, sends no input
    python test_new.py startmenu           # is the Start menu recognised (needed before typing)
    python test_new.py launch chrome       # Start-menu launch (also: notepad, writer, calc)
    python test_new.py typos               # Notepad: typo + fix typing, saved and read back
    python test_new.py site                # one Chrome site visit (shuffled actions, pauses)
    python test_new.py switch              # split screen, back to Notepad by click AND Alt+Tab
    python test_new.py research [secs]     # look-up on a listed site -> Notepad -> tidy -> save
                                           # (secs = time allowed, default RESEARCH_EST_S)
    python test_new.py all                 # launch x4, typos, research

Close Chrome / Notepad / Office first, run from the project folder, then click
away from the terminal during the countdown. A STOP file still stops it.
For launch/research, check the parent of the new process afterwards
(validate.ps1 or Sysmon ID 1): it should be explorer.exe.
"""
import os, random, sys, time

import engine
import web

PASS, FAIL = [], []


def check(name, ok, detail=""):
    (PASS if ok else FAIL).append(name)
    print("  [%s] %s %s" % ("PASS" if ok else "FAIL", name, detail))


def countdown(n=5):
    for i in range(n, 0, -1):
        print("starting in", i)
        time.sleep(1)


def _find_file(name):
    home = os.path.expanduser("~")
    for d in ("Documents", "Desktop", os.path.join("OneDrive", "Documents"), ""):
        p = os.path.join(home, d, name)
        if os.path.exists(p):
            return p
    return None


def _read(path):
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            return raw.decode(enc).replace("\r\n", "\n")
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", "replace")


def _notepad_save(fname):
    engine.press("ctrl", "s")
    if not web._window_appeared(web.SAVE_DIALOG_HINT, 8):
        return False
    engine.human_delay("action")
    engine.type_text(fname)
    engine.human_delay("action")
    engine.press("enter")
    return web._window_appeared(fname.rsplit(".", 1)[0], 8)


# --------------------------------------------------------------------------- tests

def test_pauses():
    """Sample the delay kinds without sleeping: are the new ones sane?"""
    print("\n== pauses (no input sent) ==")
    for kind in ("action", "pause", "think", "notice"):
        xs = sorted(engine._sample(kind) for _ in range(2000))
        print("  %-7s min %.2f  med %.2f  p90 %.2f  max %.2f" %
              (kind, xs[0], xs[1000], xs[1800], xs[-1]))
    real_sleep = engine.time.sleep
    engine.time.sleep = lambda s: None          # don't actually wait
    try:
        hits = sum(1 for _ in range(2000) if engine.maybe_pause(0.2))
    finally:
        engine.time.sleep = real_sleep
    check("maybe_pause(0.2) fires ~20%", 0.15 < hits / 2000 < 0.25, "(%.1f%%)" % (hits / 20))
    lo, hi = engine._PARAMS["pause"][2:]
    check("pause bounded", lo <= min(engine._sample("pause") for _ in range(500)) and
          max(engine._sample("pause") for _ in range(500)) <= hi * engine.SPEED)


def test_startmenu():
    """Is the Start menu recognised? open_app types nothing unless it is."""
    print("\n== Start menu detection ==")
    print("  front window before: %r (%s)" % (engine._foreground_title(), engine._foreground_exe()))
    engine.press("win")
    ok = engine._start_menu_open()
    print("  front window after Win: %r (%s)" % (engine._foreground_title(), engine._foreground_exe()))
    check("Start menu recognised after Win", ok,
          "" if ok else "-> add this exe name to engine._START_EXES")
    engine.type_text("note")                    # search box takes focus once you type
    engine.human_delay("think")
    ok2 = engine._foreground_exe() in engine._START_EXES
    print("  front window after typing: %r (%s)" % (engine._foreground_title(), engine._foreground_exe()))
    check("still recognised after typing", ok2,
          "" if ok2 else "-> add this exe name to engine._START_EXES")
    engine.press("escape")
    engine.human_delay("action")
    check("Start menu closed again", not engine._start_menu_open(1.0))


def test_launch(app):
    print("\n== launch %s from the Start menu ==" % app)
    if app in ("writer", "calc"):
        import office
        office._ensure_closed()
        t0 = time.monotonic()
        office._launch_writer(None if app == "writer" else " ".join(office.CALC_ARGS))
        part = office.WINDOW_PART if app == "writer" else office.CALC_PART
        ok = office._wait(part, 45, dialogs=True)
        used_start = not any("falling back to Win+R" in str(l) for l in office.ACTION_LOG[-5:])
        check("%s window appeared" % app, ok, "(%.0fs)" % (time.monotonic() - t0))
        check("%s opened via Start menu (no Win+R fallback)" % app, ok and used_start)
        engine.human_delay("read")
        office._force_close()
        return
    hint = {"chrome": web.CHROME_WINDOW_HINT, "notepad": web.NOTEPAD_HINT}[app]
    t0 = time.monotonic()
    how = engine.open_app(app, expect=hint, fallback=app, timeout=20)
    ok = web._window_appeared(hint, 10)
    check("%s window appeared" % app, ok, "(%.0fs)" % (time.monotonic() - t0))
    check("%s opened via Start menu (no Win+R fallback)" % app, how == "start", "(%s)" % how)
    engine.human_delay("read")
    engine.press("alt", "f4")
    engine.human_delay("action")


def test_typos():
    print("\n== typing with typos in Notepad ==")
    text = ("The quick brown fox jumps over the lazy dog while the team "
            "reviews the budget for next quarter.")
    engine.open_app("notepad", expect=web.NOTEPAD_HINT, fallback="notepad", timeout=15)
    if not web._window_appeared(web.NOTEPAD_HINT, 10):
        check("notepad opened", False)
        return
    engine.press("ctrl", "n")
    engine.human_delay("action")
    n = engine.type_human(text, typo_rate=0.12)      # high rate so we see several
    engine.type_text("\n")
    engine.type_text("result is ")
    engine.retype_word("good", "useful")
    print("  typos made: %d" % n)
    check("some typos happened", n > 0, "(%d)" % n)
    fname = "typo_test_%d.txt" % random.randint(100, 999)
    saved = _notepad_save(fname)
    check("saved %s" % fname, saved)
    engine.human_delay("action")
    engine.press("alt", "f4")
    path = _find_file(fname)
    if not path:
        check("file found on disk", False, "(looked in Documents/Desktop)")
        return
    got = _read(path).strip()
    want = text + "\nresult is useful"
    check("text on disk is exactly right after fixes", got == want,
          "" if got == want else "\n    want: %r\n    got:  %r" % (want, got))


def test_site():
    print("\n== one site visit ==")
    engine.open_app("chrome", expect=web.CHROME_WINDOW_HINT, fallback="chrome", timeout=20)
    web._window_appeared(web.CHROME_WINDOW_HINT, 10)
    engine.human_delay("action")
    t0 = time.monotonic()
    entry = web.surf_site(random.choice(web.CLEAN_SITES), deadline=time.monotonic() + 150)
    print("  ", entry)
    check("visit finished inside its deadline", time.monotonic() - t0 < 165,
          "(%.0fs)" % (time.monotonic() - t0))
    engine.press("ctrl", "w")
    engine.human_delay("read")


def test_switch():
    """Split screen, then go Chrome -> Notepad once by click and once by Alt+Tab,
    putting the caret back and typing a line each time. Both lines must be in the
    saved file, in the right place."""
    print("\n== split screen: back to Notepad by click and by Alt+Tab ==")
    engine.open_app("chrome", expect=web.CHROME_WINDOW_HINT, fallback="chrome", timeout=20)
    web._window_appeared(web.CHROME_WINDOW_HINT, 10)
    engine.human_delay("action")
    web._navigate(random.choice(web.CLEAN_SITES))
    web._dwell(web.LANDING)
    split = web._snap("left", web.CHROME_WINDOW_HINT)
    engine.open_app("notepad", expect=web.NOTEPAD_HINT, fallback="notepad", timeout=15)
    web._window_appeared(web.NOTEPAD_HINT, 10)
    engine.human_delay("action")
    engine.press("ctrl", "n")
    engine.human_delay("action")
    split = split and web._snap("right", web.NOTEPAD_HINT)
    check("split screen", split)
    engine.type_text("Heading\nlast line")
    modes = ("click", "alttab")
    for mode in modes:
        web.NOTEPAD_RETURN = mode
        try:
            ok = web._alt_tab(web.CHROME_WINDOW_HINT)
            check("%s: Chrome in front" % mode, ok)
            engine.scroll(3)
            engine.human_delay("action")
            ok = web._to_notepad(split)
            check("%s: Notepad in front" % mode, ok)
        finally:
            web.NOTEPAD_RETURN = None
        engine.press("ctrl", "home")         # same caret fix-up as research step 7
        engine.press("down")
        engine.press("home")
        engine.human_delay("notice")
        engine.type_text("via %s\n" % mode)
    fname = "switch_test_%d.txt" % random.randint(100, 999)
    check("saved %s" % fname, _notepad_save(fname))
    engine.human_delay("action")
    engine.press("alt", "f4")
    engine.human_delay("action")
    path = _find_file(fname)
    if path:
        got = _read(path).strip()
        want = "Heading\nvia alttab\nvia click\nlast line"
        check("both lines typed where expected", got == want,
              "" if got == want else "\n    want: %r\n    got:  %r" % (want, got))
    else:
        check("file found on disk", False)
    if web._window_appeared(web.CHROME_WINDOW_HINT, 4) and split:
        engine.press("alt", "space")
        engine.human_delay("notice")
        engine.press("x")
        engine.human_delay("action")
    engine.press("ctrl", "w")
    engine.human_delay("read")


def test_research(budget=None):
    """budget: seconds the step is allowed (default RESEARCH_EST_S, what it gets in a
    real run). It may overrun a little: typing and launching aren't cut short."""
    budget = budget or web.RESEARCH_EST_S
    print("\n== research -> Notepad (budget %ds) ==" % budget)
    engine.open_app("chrome", expect=web.CHROME_WINDOW_HINT, fallback="chrome", timeout=20)
    web._window_appeared(web.CHROME_WINDOW_HINT, 10)
    engine.human_delay("action")
    t0 = time.monotonic()
    rec = web.research_to_notepad(deadline=time.monotonic() + budget)
    secs = time.monotonic() - t0
    print("  ", rec)
    check("used a listed site only", rec["url"] in web.CLEAN_SITES, "(%s)" % rec["url"])
    check("stayed within budget (+20s grace)", secs <= budget + 20,
          "(%.0fs of %ds)" % (secs, budget))
    for step in ("search", "copy", "paste", "trim", "notes", "save"):
        check("step " + step, step in rec["steps"])
    check("Chrome back in front after Notepad closed",
          web._window_appeared(web.CHROME_WINDOW_HINT, 3))
    if web.SPLIT_SCREEN:
        check("Chrome + Notepad were split screen", "split" in rec["steps"])
        print("   (glanced back at Chrome while writing: %s)" % ("glance" in rec["steps"]))
        check("Chrome maximized again afterwards",
              bool(engine.user32.IsZoomed(engine.user32.GetForegroundWindow())))
    path = rec.get("file") and _find_file(rec["file"])
    if path:
        body = _read(path)
        print("   --- %s ---\n%s\n   ---" % (path, body[:800]))
        check("heading on top", body.startswith("Notes: " + rec["query"]))
        check("source link pasted", "Source: http" in body,
              "" if "link" in rec["steps"] else "(step skipped: out of time or lost Notepad)")
        check("own notes added", "My takeaways:" in body)
        check("rethought word fixed", "- overall useful for now" in body,
              "" if "retype" in rec["steps"] else "(step skipped: out of time or lost Notepad)")
        check("page was trimmed (< 40 lines)", body.count("\n") < 40,
              "(%d lines)" % body.count("\n"))
    else:
        check("research file found on disk", False)
    for _ in range(2):                         # research tab + the first tab -> Chrome closes
        engine.press("ctrl", "w")
        engine.human_delay("action")
    engine.human_delay("read")
    check("Chrome closed at the end", not web._window_appeared(web.CHROME_WINDOW_HINT, 2))


if __name__ == "__main__":
    args = sys.argv[1:] or ["all"]
    what = args[0]
    if what == "pauses":
        test_pauses()
    else:
        countdown()
        if what == "startmenu":
            test_startmenu()
        elif what == "launch":
            for app in (args[1:] or ["chrome", "notepad", "writer", "calc"]):
                test_launch(app)
                engine.human_delay("read")
        elif what == "typos":
            test_typos()
        elif what == "site":
            test_site()
        elif what == "switch":
            test_switch()
        elif what == "research":
            test_research(int(args[1]) if len(args) > 1 else None)
        elif what == "all":
            test_pauses()
            test_startmenu()
            for app in ("chrome", "notepad", "writer", "calc"):
                test_launch(app)
                engine.human_delay("read")
            test_typos()
            test_switch()
            test_research()
        else:
            sys.exit(__doc__)
    print("\n%d passed, %d failed%s" % (len(PASS), len(FAIL),
                                         (": " + ", ".join(FAIL)) if FAIL else ""))
