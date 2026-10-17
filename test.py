import time, statistics as st
from collections import Counter
import engine, web

EVENTS = []
_t0 = time.monotonic()
def _stamp(): return time.monotonic() - _t0

# capture every labeled delay (includes per-keystroke, via type_text internals)
_orig_delay = engine.human_delay
def _timed_delay(kind):
    a = time.monotonic()
    _orig_delay(kind)
    EVENTS.append((_stamp(), "delay", kind, time.monotonic() - a))
engine.human_delay = _timed_delay

# timestamp every discrete action
def _wrap(name):
    orig = getattr(engine, name)
    def w(*a, **k):
        EVENTS.append((_stamp(), "action", name, 0.0))
        return orig(*a, **k)
    return w
for _n in ["press","type_text","scroll","click","move_to","launch","wait_for_window"]:
    if hasattr(engine, _n): setattr(engine, _n, _wrap(_n))

print("Running surf_session(n_sites=2)... close all Chrome first, then click away.")
web.surf_session(n_sites=2)
total = _stamp()

print("\n==== RUN METRICS ====")
print("total wall time: %.1f s" % total)

delays = {}
for t, typ, label, dt in EVENTS:
    if typ == "delay": delays.setdefault(label, []).append(dt)

print("\n-- delays by kind (s) --")
print("%-10s %4s %7s %7s %7s %7s %7s" % ("kind","n","min","med","mean","max","stdev"))
for k, xs in delays.items():
    sd = st.pstdev(xs) if len(xs) > 1 else 0
    print("%-10s %4d %7.2f %7.2f %7.2f %7.2f %7.2f" %
          (k, len(xs), min(xs), st.median(xs), st.mean(xs), max(xs), sd))

print("\n-- repetition check (rounded 0.1s; high repeat = robotic) --")
for k, xs in delays.items():
    val, cnt = Counter(round(x,1) for x in xs).most_common(1)[0]
    print("  %-10s most common %.1fs x%d of %d" % (k, val, cnt, len(xs)))

acts = Counter(l for t,ty,l,d in EVENTS if ty=="action")
print("\n-- action counts --")
for n, c in acts.most_common(): print("  %-16s %d" % (n, c))

at = [t for t,ty,l,d in EVENTS if ty=="action"]
gaps = [b-a for a,b in zip(at, at[1:])]
if gaps:
    print("\n-- gaps between actions (s) --")
    print("  n=%d min=%.2f med=%.2f mean=%.2f max=%.2f" %
          (len(gaps), min(gaps), st.median(gaps), st.mean(gaps), max(gaps)))

with open("metrics_log.txt","w") as f:
    for t,ty,l,d in EVENTS: f.write("%.3f\t%s\t%s\t%.3f\n" % (t,ty,l,d))
print("\nraw timeline -> metrics_log.txt\n==== END ====")

# ---------------------------------------------------------------------------
# CV summary - paste at the END of test.py (after metrics_log.txt is written).
# Reuses the same delay-kind filter as make_histogram.py so the numbers match
# the report. Action-call markers (press/type_text/wait_for_window/launch) have
# no duration and are excluded, exactly like the HTML report.
# ---------------------------------------------------------------------------
def _print_cv_summary(path="metrics_log.txt"):
    import math, os
    if not os.path.exists(path):
        print("[cv] no", path, "to summarise")
        return
    series, all_d = {}, []
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            p = line.rstrip("\n").split("\t")
            if len(p) < 4:
                continue
            typ, label = p[1].strip().lower(), p[2].strip()
            try:
                dt = float(p[3])
            except ValueError:
                continue
            # Only "delay" rows are human_delay samples. "action" rows are engine
            # calls (and engine.scroll() shares the label "scroll"), so filter on
            # the type column, not the label.
            if dt < 0 or typ != "delay":
                continue
            series.setdefault(label, []).append(dt)
            all_d.append(dt)

    def stats(xs):
        n = len(xs)
        if n == 0:
            return (0, 0.0, 0.0, 0.0)
        m = sum(xs) / n
        sd = math.sqrt(sum((x - m) ** 2 for x in xs) / (n - 1)) if n > 1 else 0.0
        cv = sd / m if m else 0.0
        return (n, m, cv, max(xs))

    print("\n=== timing CV (real delays only) ===")
    n, m, cv, mx = stats(all_d)
    print(f"ALL delays   n={n:<4} mean={m:5.2f}s  CV={cv:4.2f}  max={mx:6.2f}s")
    for k in ("keystroke", "action", "read", "scroll", "think", "notice"):
        if k in series:
            n, m, cv, mx = stats(series[k])
            print(f"  {k:<10} n={n:<4} mean={m:5.2f}s  CV={cv:4.2f}  max={mx:6.2f}s")
    print("CV~0 = metronomic/robotic; high CV + big max = human thinking pauses.")


_print_cv_summary()