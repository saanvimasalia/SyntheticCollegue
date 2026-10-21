# engine.py (P1). Real: human_delay, press, type_text. Still stubs: launch, wait_for_window, mouse.
import ctypes, math, os, random, time
from ctypes import wintypes

user32 = ctypes.WinDLL("user32", use_last_error=True)
user32.VkKeyScanW.restype = ctypes.c_short
user32.VkKeyScanW.argtypes = [wintypes.WCHAR]
ULONG_PTR = wintypes.WPARAM

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx", wintypes.LONG), ("dy", wintypes.LONG),
                ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                ("dwExtraInfo", ULONG_PTR))

class _U(ctypes.Union):
    _fields_ = (("mi", MOUSEINPUT), ("ki", KEYBDINPUT))

class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = (("type", wintypes.DWORD), ("u", _U))

KEYEVENTF_EXTENDEDKEY, KEYEVENTF_KEYUP, KEYEVENTF_UNICODE = 0x1, 0x2, 0x4

_VK = {"ctrl": 0x11, "control": 0x11, "shift": 0x10, "alt": 0x12, "win": 0x5B,
       "enter": 0x0D, "return": 0x0D, "tab": 0x09, "esc": 0x1B, "escape": 0x1B,
       "backspace": 0x08, "delete": 0x2E, "space": 0x20,
       "left": 0x25, "up": 0x26, "right": 0x27, "down": 0x28,
       "home": 0x24, "end": 0x23, "pageup": 0x21, "pagedown": 0x22}
for _i in range(1, 13):
    _VK["f%d" % _i] = 0x6F + _i
_EXTENDED = {0x5B, 0x2E, 0x25, 0x26, 0x27, 0x28, 0x24, 0x23, 0x21, 0x22}

# ---------- timing ----------
# kind: (median s, sigma, min s, max s)
_PARAMS = {"keystroke": (0.13, 0.40, 0.04, 1.0),
           "action":    (1.6, 0.60, 0.4, 12.0),
           "read":      (9.0, 0.60, 2.0, 60.0),
           "scroll":    (0.8, 0.50, 0.2, 6.0),
           "think":     (2.2, 0.35, 1.2, 6.0),
           "notice":    (0.5, 0.35, 0.3, 1.5),
           "pause":     (6.0, 0.60, 2.5, 25.0)}   # stops between steps (look away, reread)
SPEED = 1.0                      # set per run by the sequencer (persona)
def _sample(kind):
    med, sig, lo, hi = _PARAMS[kind]
    return min(hi, max(lo, random.lognormvariate(math.log(med), sig))) * SPEED

def human_delay(kind):
    if os.path.exists("STOP"):
        raise SystemExit("STOP file found")
    d = _sample(kind)
    time.sleep(d)
    return d

def maybe_pause(p=0.2):
    """Now and then a person stops between steps for a few seconds."""
    if random.random() < p:
        return human_delay("pause")
    return 0.0

def _hold():   # how long a key stays pressed
    return min(0.2, max(0.025, random.lognormvariate(math.log(0.07), 0.3)))

def _gap():    # short gap between a modifier and its key
    return random.uniform(0.03, 0.09)

# ---------- low-level input ----------
def _send(vk=0, scan=0, flags=0):
    inp = INPUT(type=1, ki=KEYBDINPUT(vk, scan, flags, 0, 0))
    if user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)) != 1:
        raise ctypes.WinError(ctypes.get_last_error())

def _key(vk, up=False):
    flags = (KEYEVENTF_EXTENDEDKEY if vk in _EXTENDED else 0) | (KEYEVENTF_KEYUP if up else 0)
    _send(vk, user32.MapVirtualKeyW(vk, 0), flags)

def _unicode(ch, up=False):
    _send(0, ord(ch), KEYEVENTF_UNICODE | (KEYEVENTF_KEYUP if up else 0))

def _vk_for(key):
    k = key.lower()
    if k in _VK:
        return _VK[k]
    if len(k) == 1:
        r = user32.VkKeyScanW(k)
        if r != -1:
            return r & 0xFF
    raise ValueError("unknown key: %r" % key)

def press(*keys):
    """press("ctrl","l") holds ctrl, taps l, releases. press("enter") taps enter."""
    vks = [_vk_for(k) for k in keys]
    mods, last = vks[:-1], vks[-1]
    for m in mods:
        _key(m); time.sleep(_gap())
    _key(last); time.sleep(_hold()); _key(last, up=True)
    for m in reversed(mods):
        time.sleep(_gap()); _key(m, up=True)

def _type_char(ch):
    r = user32.VkKeyScanW(ch)
    if r == -1 or ((r >> 8) & 0x06):      # not on keyboard, or needs ctrl/alt
        _unicode(ch); time.sleep(_hold()); _unicode(ch, up=True)
        return
    vk, shift = r & 0xFF, bool((r >> 8) & 1)
    if shift:
        _key(0x10); time.sleep(_gap())
    _key(vk); time.sleep(_hold()); _key(vk, up=True)
    if shift:
        time.sleep(_gap()); _key(0x10, up=True)

def type_text(text):
    for ch in text:
        if ch == "\n":
            press("enter")
        elif ch == "\t":
            press("tab")
        else:
            _type_char(ch)
        human_delay("keystroke")
        if ch == " ":
            time.sleep(random.uniform(0.02, 0.12))
            if random.random() < 0.10:     # words come out in bursts
                time.sleep(random.uniform(0.3, 1.1))
        elif ch in ".,;:!?\n":
            time.sleep(min(2.0, random.lognormvariate(math.log(0.35), 0.5)))
        if random.random() < 1 / 40:       # occasional thinking pause
            time.sleep(min(6.0, max(1.2, random.lognormvariate(math.log(2.2), 0.35))))

# ---------- typing with typos ----------
_ROWS = ["qwertyuiop", "asdfghjkl", "zxcvbnm"]
TYPO_RATE = 0.03                    # chance per letter of hitting a neighbouring key

def _neighbors(ch):
    c = ch.lower()
    for r, row in enumerate(_ROWS):
        if c in row:
            i = row.index(c)
            out = [row[j] for j in (i - 1, i + 1) if 0 <= j < len(row)]
            out += [_ROWS[rr][i] for rr in (r - 1, r + 1)
                    if 0 <= rr < len(_ROWS) and i < len(_ROWS[rr])]
            return out
    return []

def type_human(text, typo_rate=None):
    """type_text, but now and then a neighbouring key gets hit; the slip is noticed
    0-2 letters later and backspaced away. The final text is always `text`.
    Returns the number of typos made."""
    rate = TYPO_RATE if typo_rate is None else typo_rate
    typos = 0
    for i, ch in enumerate(text):
        nb = _neighbors(ch) if ch.isalpha() else []
        if nb and i > 0 and random.random() < rate:
            wrong = random.choice(nb)
            wrong = wrong.upper() if ch.isupper() else wrong
            ahead = text[i + 1:i + 1 + random.choice([0, 0, 1, 2])]
            ahead = ahead.split("\n")[0].split("\t")[0]
            type_text(wrong + ahead)
            human_delay("notice")              # "wait, that's wrong"
            for _ in range(1 + len(ahead)):
                press("backspace")
                human_delay("keystroke")
            typos += 1
        type_text(ch)
    return typos

def retype_word(word, alternative):
    """Type a word, think better of it, Ctrl+Backspace it, type another one."""
    type_human(word)
    human_delay("think")
    press("ctrl", "backspace")
    human_delay("keystroke")
    type_human(alternative)

# ---------- still stubs ----------
# ---------- window waiting + launch ----------
user32.GetForegroundWindow.restype = wintypes.HWND
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

def _foreground_title():
    hwnd = user32.GetForegroundWindow()
    n = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, buf, n + 1)
    return buf.value

def wait_for_window(title_part, timeout=15, exact=False):
    """Wait until the foreground window title matches. Returns True/False."""
    end = time.monotonic() + timeout
    want = title_part.lower()
    while time.monotonic() < end:
        t = _foreground_title().lower()
        if (t == want) if exact else (want in t):
            return True
        time.sleep(random.uniform(0.1, 0.35))
    return False

def launch(path, args=None):
    """Open an app the way a person does: Win+R, type, Enter. Parent = explorer.exe."""
    if " " in path and not path.startswith('"'):
        path = '"%s"' % path
    if isinstance(args, (list, tuple)):
        args = " ".join(args)
    cmd = path if not args else path + " " + args
    press("win", "r")
    if not wait_for_window("Run", 5, exact=True):
        # never type a command into whatever else is in front
        raise RuntimeError("Run dialog did not open (front: %r)" % _foreground_title())
    human_delay("action")
    type_text(cmd)
    human_delay("action")
    press("enter")

LAUNCH_VIA_START = True     # False = always Win+R with the full path (old behaviour)

# processes that own the Start menu / its search box (Win10 and Win11)
_START_EXES = {"startmenuexperiencehost.exe", "searchapp.exe", "searchhost.exe",
               "searchui.exe", "shellexperiencehost.exe"}
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_kernel32.OpenProcess.restype = wintypes.HANDLE
_kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
_kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                                 wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
_kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

def _foreground_exe():
    """Exe name of the process that owns the front window ('' if unknown).
    ctypes only: no child process."""
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), ctypes.byref(pid))
    h = _kernel32.OpenProcess(0x1000, False, pid.value)    # QUERY_LIMITED_INFORMATION
    if not h:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(520)
        n = wintypes.DWORD(520)
        if _kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(n)):
            return os.path.basename(buf.value).lower()
        return ""
    finally:
        _kernel32.CloseHandle(h)

def _start_menu_open(timeout=3.0):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        if _foreground_exe() in _START_EXES:
            return True
        time.sleep(random.uniform(0.1, 0.25))
    return False

def open_app(search, expect=None, fallback=None, fallback_args=None, timeout=25):
    """Open an app the way most people do: tap Win, type part of its name, glance
    at the results, Enter. Start-menu launches are parented by explorer.exe, like
    Win+R, but nobody has to type an exe path. If the `expect` window never shows
    up, fall back to launch(fallback) via Win+R. Returns "start" or "run".
    Nothing is typed unless the Start menu is really in front: otherwise the app
    name lands in whatever had focus (VM 21 Oct: Chrome's address bar -> Google
    search -> offline dino page)."""
    if LAUNCH_VIA_START:
        opened = False
        for _ in range(2):                 # a second tap if the first didn't take
            press("win")
            if _start_menu_open():
                opened = True
                break
            human_delay("action")
            if _start_menu_open(0.5):      # it may have toggled shut: tap again
                opened = True
                break
        if opened:
            human_delay("action")          # Start menu slides open
            type_human(search, typo_rate=0.02)
            human_delay("think")           # results list fills in, find the right one
            if _foreground_exe() in _START_EXES:
                press("enter")
                if expect is None or wait_for_window(expect, timeout):
                    return "start"
            if _foreground_exe() in _START_EXES:
                press("escape")            # Start menu still up: close it
                human_delay("action")
    if fallback is None:
        raise RuntimeError("could not open %r from the Start menu" % search)
    launch(fallback, fallback_args)
    return "run"

# ---------- mouse ----------
MOUSEEVENTF_MOVE, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP = 0x1, 0x2, 0x4
MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP = 0x8, 0x10
MOUSEEVENTF_WHEEL, MOUSEEVENTF_ABSOLUTE = 0x800, 0x8000

class _POINT(ctypes.Structure):
    _fields_ = (("x", wintypes.LONG), ("y", wintypes.LONG))

def cursor_pos():
    p = _POINT()
    user32.GetCursorPos(ctypes.byref(p))
    return p.x, p.y

def _mouse(flags, dx=0, dy=0, data=0):
    inp = INPUT(type=0, mi=MOUSEINPUT(dx, dy, data & 0xFFFFFFFF, flags, 0, 0))
    if user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)) != 1:
        raise ctypes.WinError(ctypes.get_last_error())

def _abs_move(x, y):
    w, h = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
    _mouse(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE,
           int(x * 65535 / (w - 1)), int(y * 65535 / (h - 1)))

def move_to(x, y):
    """Glide to (x, y) along a gentle curve: slow start, fast middle, slow end."""
    x0, y0 = cursor_pos()
    dx, dy = x - x0, y - y0
    dist = math.hypot(dx, dy)
    if dist < 2:
        _abs_move(x, y)
        return
    off = random.uniform(-0.15, 0.15) * dist          # sideways bend
    cx = (x0 + x) / 2 - dy / dist * off
    cy = (y0 + y) / 2 + dx / dist * off
    total = min(2.0, (0.2 + dist / 1500) * random.lognormvariate(0, 0.25))
    steps = max(12, int(dist / 10))
    for i in range(1, steps + 1):
        t = i / steps
        t = t * t * (3 - 2 * t)                       # ease in/out
        px = (1 - t) ** 2 * x0 + 2 * (1 - t) * t * cx + t * t * x
        py = (1 - t) ** 2 * y0 + 2 * (1 - t) * t * cy + t * t * y
        _abs_move(round(px), round(py))
        time.sleep(total / steps * random.uniform(0.7, 1.3))

def click(button="left"):
    down, up = ((MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP) if button == "right"
                else (MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP))
    _mouse(down)
    time.sleep(_hold())
    _mouse(up)

def scroll(clicks):
    """Positive = down, negative = up. Scrolls in bursts of 2-5 notches with pauses."""
    direction = -1 if clicks > 0 else 1               # wheel: negative = down
    remaining = abs(clicks)
    while remaining > 0:
        burst = min(remaining, random.randint(2, 5))
        for _ in range(burst):
            _mouse(MOUSEEVENTF_WHEEL, data=direction * 120)
            time.sleep(random.uniform(0.02, 0.09))
        remaining -= burst
        if remaining:
            human_delay("scroll")