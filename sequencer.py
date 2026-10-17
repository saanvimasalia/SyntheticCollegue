# sequencer.py (P1): assembles a different ~15-minute routine each run
import datetime, math, os, random, time
import engine
import web
try:
    import office
except ImportError:
    office = None

LOG_FILE = "agent_log.txt"
SEC_PER_SITE = 150      # from the real agent_log: 2 sites took 257 s and 338 s, 1 site 128 s
OFFICE_EST = 240        # used only by the old random picker (GUARANTEE = False)
SHEET_EST = 260         # same: spreadsheet needs ~200 s+, so only offer it with this much left
MIN_BLOCK_S = 130       # never start a block with less than this left

# --- coverage guarantee -----------------------------------------------------
# GUARANTEE = True : every run does one web session, one document and one
#                    spreadsheet (shuffled order) and exactly one AFK gap.
#                    Everything a judge reads as "human" stays random: order,
#                    timing, sites, typos, filenames, persona.
# GUARANTEE = False: the old behaviour (random pick, random AFK chance).
GUARANTEE = True
MIN_NEED = {"web": 130, "office": 150, "spreadsheet": 200}    # least time a block can run in
TARGET = {"web": 330, "office": 230, "spreadsheet": 200}      # time a block is aimed at
                             # (web: the block gets TARGET+20 s and runs int((that-30)/SEC_PER_SITE)
                             #  sites, so 330 is the least that allows 2 sites at 150 s each)
WEB_SITES = (2, 2)           # guarantee mode: sites per web session (min, max).
                             # 2 sites ~ 260-340 s with the current web.py dwells. Go to (3, 3)
                             # (and TARGET["web"] ~ 420) only after web.py's dwells are retuned.
AFK_RANGE = (40.0, 90.0)     # the one forced AFK gap (seconds)
IDLE_GAP_S = 25              # allowance reserved per idle gap between blocks
SPREAD_SLACK = False         # True = lengthen idle gaps to absorb leftover time. Off by default:
                             # in 400 dry runs it cut the mean tail by only ~5 s but made
                             # ordinary gaps much longer (mean 21 s -> 38 s, max 75 s).
TAIL_KEEP = 30               # leftover we are happy to end the run with (seconds)
IDLE_MAX_S = 75              # longest ordinary idle gap, so the AFK stays the odd one out

KEEP_P = 0.0            # chance of leaving Writer open during the next web block.
                        # 0 = never (recommended for the live run until a rehearsal
                        # proves the overlap); raise only for tests.
_office_open = False    # True while Writer is left open on purpose
_virtual = 0.0          # used only by dry runs

_todo = []              # block kinds still owed this run (guarantee mode)
_afk_done = True        # has the forced AFK happened yet
_afk_after = 1          # the AFK comes after this many finished blocks
_afk_len = 60.0
_blocks_done = 0


def _now():
    return time.monotonic() + _virtual


def _advance(s):
    global _virtual
    _virtual += s


def log(msg):
    line = "%s  %s" % (datetime.datetime.now().strftime("%H:%M:%S"), msg)
    print(line)
    with open(LOG_FILE, "a") as f:
        f.write(line + "\n")


PERSONAS = {                      # speed < 1 = faster, > 1 = slower
    "focused":    {"idle_med": 12, "afk_p": 0.03, "speed": 0.9},
    "casual":     {"idle_med": 20, "afk_p": 0.08, "speed": 1.0},
    "distracted": {"idle_med": 30, "afk_p": 0.15, "speed": 1.15},
}
_persona = PERSONAS["casual"]


def _sleep(d):
    end = time.monotonic() + d
    while True:
        left = end - time.monotonic()
        if left <= 0:
            return
        if os.path.exists("STOP"):
            raise SystemExit("STOP file found")
        time.sleep(min(1.0, left))


def _reserve():
    """Time to keep back for blocks still owed, their idle gaps and the AFK."""
    r = sum(MIN_NEED[k] + IDLE_GAP_S for k in _todo)
    if not _afk_done:
        r += _afk_len
    return r


def _slack_share(remaining):
    """Spread time that would otherwise be dead air at the end of the run across
    the idle gaps that come before the last owed block.

    Slack = what is left once every block still owed has run at its target length
    (plus its gap and the pending AFK) and TAIL_KEEP has been set aside. It is
    recomputed at every gap from the real clock, so an early or late block is
    corrected at the next gap instead of piling up at the end."""
    need = sum(TARGET[k] + IDLE_GAP_S for k in _todo)
    if not _afk_done:
        need += _afk_len
    slack = remaining - need - TAIL_KEEP
    if slack <= 0 or not _todo:
        return 0.0
    # share across the gaps still ahead of the last owed block, never all at once
    return (slack / len(_todo)) * random.uniform(0.6, 1.0)


def _idle(dry, remaining):
    global _afk_done
    if GUARANTEE:
        afk = (not _afk_done) and _blocks_done >= _afk_after
    else:
        afk = random.random() < _persona["afk_p"]
    if afk:
        d = _afk_len if GUARANTEE else random.uniform(90.0, 200.0)
        _afk_done = True
    else:
        d = min(60.0, max(5.0, random.lognormvariate(math.log(_persona["idle_med"]), 0.6)))
        if GUARANTEE and SPREAD_SLACK:
            d = min(IDLE_MAX_S, d + _slack_share(remaining))
    # never let a gap eat the time that blocks still owed need
    # (+1 s margin: the clock keeps ticking, and landing on 129.99 s would skip the next block)
    # (floor is 2 s, not 5 s: with 133 s left a 5 s gap leaves 128 s and the next block is skipped)
    d = min(d, max(2.0, remaining - max(MIN_BLOCK_S + 1, _reserve())))
    log("%s %.0fs" % ("AFK" if afk else "idle", d))
    if dry:
        _advance(d)
        return
    if not afk and random.random() < 0.5:     # drift the mouse a little
        try:
            w = engine.user32.GetSystemMetrics(0)
            h = engine.user32.GetSystemMetrics(1)
            engine.move_to(random.randint(int(w * .2), int(w * .8)),
                           random.randint(int(h * .2), int(h * .8)))
        except Exception as e:
            log("idle mouse-drift skipped: %r" % (e,))   # don't let an input error kill the run
    _sleep(d)


def _run_web(remaining, dry):
    lo, hi = WEB_SITES if GUARANTEE else (1, 3)
    n = max(1, min(random.randint(lo, hi), int((remaining - 30) / SEC_PER_SITE)))
    log("BLOCK web: %d site(s)" % n)
    if dry:
        _advance(min(n * SEC_PER_SITE * random.uniform(0.8, 1.3), remaining))
    else:
        try:
            result = web.surf_session(n, max_seconds=max(30, remaining - 20))
            log("web result: %s" % (result,))
        except Exception as e:
            log("web error: %r" % (e,))


def _close_office_if_open(dry):
    global _office_open
    if not _office_open:
        return
    _office_open = False
    if dry or os.path.exists("STOP"):
        return
    try:
        log("closing Writer left open")
        office._force_close()
    except Exception as e:
        log("close error: %r" % (e,))


def _run_office(remaining, dry):
    global _office_open
    _close_office_if_open(dry)
    if GUARANTEE:
        # leaving Writer open is only allowed while a web block is still owed,
        # because _pick forces a web block right after it
        can_keep = "web" in _todo
    else:
        can_keep = remaining > OFFICE_EST + MIN_BLOCK_S + 120
    keep = can_keep and random.random() < KEEP_P
    log("BLOCK office: write document" + (" (keep open)" if keep else ""))
    if dry:
        _advance(min(TARGET["office"] * random.uniform(0.8, 1.3), remaining))
        _office_open = keep
        return
    try:
        result = office.write_document(max_seconds=max(60, remaining - 20), keep_open=keep)
        log("office result: %s" % (result,))
        _office_open = keep and bool(result.get("ok"))
    except Exception as e:
        log("office error: %r" % (e,))


def _run_spreadsheet(remaining, dry):
    log("BLOCK office: spreadsheet")
    if dry:
        _advance(min(TARGET["spreadsheet"] * random.uniform(0.75, 1.2), remaining))
        return
    try:
        result = office.write_spreadsheet(max_seconds=max(60, remaining - 20))
        log("spreadsheet result: %s" % (result,))
    except Exception as e:
        log("spreadsheet error: %r" % (e,))


def _pick(last, remaining, dry):
    if _office_open:
        return "web"
    for k in _todo:                            # owed kinds first, in shuffled order
        if remaining >= MIN_NEED[k]:
            return k
    kinds = ["web"]
    have_office = dry or (office is not None and hasattr(office, "write_document"))
    if have_office and remaining > OFFICE_EST:
        kinds.append("office")
    have_sheet = dry or (office is not None and hasattr(office, "write_spreadsheet"))
    if have_sheet and remaining > SHEET_EST:
        kinds.append("spreadsheet")
    if len(kinds) > 1 and last in kinds:
        kinds.remove(last)                     # never the same block twice in a row
    return random.choice(kinds)


def run(budget_s=900, dry=False):
    global _virtual, _persona, _afk_done, _afk_after, _afk_len, _blocks_done
    _virtual = 0.0
    log("=== run start, budget %ds, dry=%s ===" % (budget_s, dry))
    name = random.choice(list(PERSONAS))
    _persona = PERSONAS[name]
    engine.SPEED = _persona["speed"]
    log("persona: %s (speed %.2f)" % (name, _persona["speed"]))

    # the to-do list for this run: only kinds the office module can actually do
    _todo[:] = []
    if GUARANTEE:
        _todo.append("web")
        if dry or (office is not None and hasattr(office, "write_document")):
            _todo.append("office")
        if dry or (office is not None and hasattr(office, "write_spreadsheet")):
            _todo.append("spreadsheet")
        random.shuffle(_todo)
    _afk_done = not GUARANTEE
    _blocks_done = 0
    _afk_after = random.choice([1, 2])
    _afk_len = random.uniform(*AFK_RANGE)
    if GUARANTEE:
        log("plan: %s, AFK after block %d (%.0fs)" % (" > ".join(_todo), _afk_after, _afk_len))

    if not dry:
        for i in (3, 2, 1):                        # time to click away from the terminal
            print("starting in", i)
            time.sleep(1)
    start, last = _now(), None
    try:
        while True:
            remaining = budget_s - (_now() - start)
            if remaining < MIN_BLOCK_S:
                break
            kind = _pick(last, remaining, dry)
            if kind in _todo:
                _todo.remove(kind)
            if GUARANTEE:
                # cap this block so the ones still owed keep their minimum time
                avail = min(remaining, max(MIN_NEED[kind],
                                           min(remaining - _reserve(), TARGET[kind] + 20)))
            else:
                avail = remaining
            t0 = _now()
            if kind == "web":
                _run_web(avail, dry)
                _close_office_if_open(dry)
            elif kind == "spreadsheet":
                _run_spreadsheet(avail, dry)
            else:
                _run_office(avail, dry)
            _blocks_done += 1
            log("block %s took %.0fs" % (kind, _now() - t0))
            last = kind
            if budget_s - (_now() - start) < MIN_BLOCK_S:
                break
            _idle(dry, budget_s - (_now() - start))

        left = budget_s - (_now() - start)
        if left > 1:
            log("tail idle %.0fs" % left)
            if dry:
                _advance(left)
            else:
                _sleep(left)
    finally:
        _close_office_if_open(dry)
        log("=== run end, elapsed %.0fs ===" % (_now() - start))