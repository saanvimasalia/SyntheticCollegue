"""
make_histogram.py - Timing-evidence report for the Synthetic Human Agent (P2).

Reads the two timing sources we already produce on the VM and writes ONE
self-contained HTML page with inline-SVG histograms. No matplotlib, no numpy,
no internet - pure stdlib, so it runs on the VM as-is and opens in Chrome.

Inputs (either or both; it uses whatever it finds):
    metrics_log.txt            tab-separated: t <TAB> type <TAB> label <TAB> dt
                               (produced by test.py - the per-action delay timeline)
    validation\\gaps_*.csv      header "GapSeconds", one gap per line
                               (produced by validate.ps1 - Sysmon process-event gaps)

Usage (from the project folder on the VM):
    python make_histogram.py
    # then open timing_report.html in Chrome

Optional explicit paths:
    python make_histogram.py metrics_log.txt validation\\gaps_1234.csv

The point for judges: human timing is a heavy-tailed distribution, not fixed
intervals. The page shows the distribution shape + the coefficient of variation
(CV = sd/mean) for each series. CV near 0 = metronomic (robotic); CV well above
0 with a right tail = human-like.
"""

import sys
import os
import glob
import math
import html

# The six kinds sampled by engine.human_delay. Everything else in the log
# (press / type_text / wait_for_window ...) is an action-call MARKER - a bare
# timestamp with no duration - and must NOT be charted as a delay, or it shows
# up as a misleading CV=0.00 card.
DELAY_KINDS = {"keystroke", "action", "read", "scroll", "think", "notice"}


# ----------------------------- stats helpers -----------------------------

def mean(xs):
    return sum(xs) / len(xs) if xs else 0.0


def stdev(xs):
    if len(xs) < 2:
        return 0.0
    m = mean(xs)
    return math.sqrt(sum((x - m) ** 2 for x in xs) / (len(xs) - 1))


def cv(xs):
    m = mean(xs)
    return (stdev(xs) / m) if m else 0.0


def median(xs):
    if not xs:
        return 0.0
    s = sorted(xs)
    n = len(s)
    mid = n // 2
    return s[mid] if n % 2 else (s[mid - 1] + s[mid]) / 2


def percentile(xs, p):
    if not xs:
        return 0.0
    s = sorted(xs)
    k = (len(s) - 1) * (p / 100.0)
    lo = math.floor(k)
    hi = math.ceil(k)
    if lo == hi:
        return s[int(k)]
    return s[lo] * (hi - k) + s[hi] * (k - lo)


def histogram(xs, bins):
    """Return (edges, counts) with `bins` equal-width bins over [min, max]."""
    if not xs:
        return [], []
    lo, hi = min(xs), max(xs)
    if hi <= lo:
        hi = lo + 1.0
    width = (hi - lo) / bins
    edges = [lo + i * width for i in range(bins + 1)]
    counts = [0] * bins
    for x in xs:
        idx = int((x - lo) / width)
        if idx == bins:            # right edge falls in the last bin
            idx = bins - 1
        counts[idx] += 1
    return edges, counts


# ----------------------------- input parsing -----------------------------

def load_metrics_log(path):
    """metrics_log.txt -> (series, markers).

    series  : kind -> [dt, ...] for real human_delay samples, plus 'ALL delays'
    markers : label -> count for action-call rows (NOT delays; excluded from charts)
    """
    series = {}
    all_delays = []
    markers = {}
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line.strip():
                continue
            parts = line.split("\t")
            if len(parts) < 4:
                continue
            _t, typ, label, dt = parts[0], parts[1], parts[2], parts[3]
            try:
                dt = float(dt)
            except ValueError:
                continue              # header or malformed row
            if dt < 0:
                continue
            label = (label or "").strip()
            typ_l = (typ or "").strip().lower()
            # The `type` column is the authoritative separator: "delay" rows are
            # human_delay samples, "action" rows are engine-function calls. Keying
            # on type (not label) avoids the name collision where engine.scroll()
            # logs label "scroll" just like the human_delay("scroll") kind does.
            if typ_l == "delay":
                series.setdefault(label or "delay", []).append(dt)
                all_delays.append(dt)
            else:
                key = label or typ_l or "?"
                markers[key] = markers.get(key, 0) + 1
    if all_delays:
        series["ALL delays"] = all_delays
    return series, markers


def load_gaps_csv(path):
    """gaps_*.csv with header GapSeconds -> list of floats."""
    gaps = []
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            s = line.strip().strip('"')
            if not s:
                continue
            if s.lower().startswith("gap"):   # header
                continue
            # tolerate a stray comma-wrapped cell
            s = s.split(",")[0].strip().strip('"')
            try:
                v = float(s)
            except ValueError:
                continue
            if v >= 0:
                gaps.append(v)
    return gaps


def find_default_inputs():
    metrics = "metrics_log.txt" if os.path.exists("metrics_log.txt") else None
    gaps = None
    candidates = sorted(
        glob.glob(os.path.join("validation", "gaps_*.csv")),
        key=lambda p: os.path.getmtime(p),
        reverse=True,
    )
    if candidates:
        gaps = candidates[0]        # most recent run
    return metrics, gaps


# ----------------------------- SVG rendering -----------------------------

def svg_histogram(xs, title, unit="s", bins=24, color="#4f7cff"):
    """One labelled histogram as an inline <svg> string."""
    if not xs:
        return f'<div class="chart"><h3>{html.escape(title)}</h3><p class="empty">no data</p></div>'

    W, H = 560, 300
    pad_l, pad_r, pad_t, pad_b = 46, 14, 16, 42
    plot_w = W - pad_l - pad_r
    plot_h = H - pad_t - pad_b

    edges, counts = histogram(xs, bins)
    cmax = max(counts) if counts else 1
    bw = plot_w / len(counts)

    bars = []
    for i, c in enumerate(counts):
        bh = (c / cmax) * plot_h if cmax else 0
        x = pad_l + i * bw
        y = pad_t + (plot_h - bh)
        bars.append(
            f'<rect x="{x:.1f}" y="{y:.1f}" width="{max(bw-1,0.5):.1f}" '
            f'height="{bh:.1f}" fill="{color}"><title>{counts[i]} '
            f'in [{edges[i]:.2f}, {edges[i+1]:.2f}] {unit}</title></rect>'
        )

    # x-axis ticks: min, mid, max
    xlo, xhi = edges[0], edges[-1]
    xticks = []
    for frac in (0.0, 0.5, 1.0):
        val = xlo + frac * (xhi - xlo)
        px = pad_l + frac * plot_w
        xticks.append(
            f'<text x="{px:.1f}" y="{H-pad_b+16:.0f}" text-anchor="middle" '
            f'class="tick">{val:.1f}{unit}</text>'
        )

    # y-axis ticks: 0, cmax
    yticks = [
        f'<text x="{pad_l-6}" y="{pad_t+plot_h:.0f}" text-anchor="end" class="tick">0</text>',
        f'<text x="{pad_l-6}" y="{pad_t+8:.0f}" text-anchor="end" class="tick">{cmax}</text>',
    ]

    # overlay mean + median lines
    def vline(val, col, dash):
        if xhi <= xlo:
            return ""
        px = pad_l + (val - xlo) / (xhi - xlo) * plot_w
        px = max(pad_l, min(pad_l + plot_w, px))
        return (f'<line x1="{px:.1f}" y1="{pad_t}" x2="{px:.1f}" '
                f'y2="{pad_t+plot_h}" stroke="{col}" stroke-width="1.5" '
                f'stroke-dasharray="{dash}"/>')

    m = mean(xs)
    md = median(xs)
    overlay = vline(m, "#e23b3b", "4,3") + vline(md, "#1d9d55", "2,2")

    axis = (f'<line x1="{pad_l}" y1="{pad_t+plot_h}" x2="{pad_l+plot_w}" '
            f'y2="{pad_t+plot_h}" stroke="#999" stroke-width="1"/>'
            f'<line x1="{pad_l}" y1="{pad_t}" x2="{pad_l}" '
            f'y2="{pad_t+plot_h}" stroke="#999" stroke-width="1"/>')

    svg = (
        f'<svg viewBox="0 0 {W} {H}" width="100%" height="auto" '
        f'preserveAspectRatio="xMidYMid meet" role="img" '
        f'aria-label="{html.escape(title)} histogram">'
        f'{"".join(bars)}{overlay}{axis}'
        f'{"".join(xticks)}{"".join(yticks)}'
        f'</svg>'
    )

    stats = (
        f'<div class="stats">'
        f'<span>n={len(xs)}</span>'
        f'<span>mean={m:.2f}{unit}</span>'
        f'<span>median={md:.2f}{unit}</span>'
        f'<span>sd={stdev(xs):.2f}{unit}</span>'
        f'<span class="cv">CV={cv(xs):.2f}</span>'
        f'<span>p90={percentile(xs,90):.2f}{unit}</span>'
        f'<span>max={max(xs):.2f}{unit}</span>'
        f'</div>'
    )

    legend = ('<div class="legend">'
              '<span><i class="sw mean"></i>mean</span>'
              '<span><i class="sw med"></i>median</span>'
              '</div>')

    return (f'<div class="chart"><h3>{html.escape(title)}</h3>'
            f'{svg}{legend}{stats}</div>')


# ----------------------------- page assembly -----------------------------

PAGE_CSS = """
:root{--bg:#f6f7fb;--card:#fff;--ink:#1a1d24;--muted:#667085;--line:#e5e8ef;
      --accent:#4f7cff;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
     font:15px/1.5 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;}
.wrap{max-width:1180px;margin:0 auto;padding:28px 18px 60px;}
h1{font-size:24px;margin:0 0 4px}
.sub{color:var(--muted);margin:0 0 22px}
.verdict{background:var(--card);border:1px solid var(--line);border-radius:12px;
         padding:16px 18px;margin:0 0 24px}
.verdict h2{font-size:15px;margin:0 0 8px;text-transform:uppercase;
            letter-spacing:.04em;color:var(--muted)}
.verdict p{margin:6px 0}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));
      gap:18px}
/* aspect-ratio reserves the correct box height, so the SVG can't overflow its
   card and spill onto the next heading (the earlier height:auto overlap bug). */
.chart svg{display:block;width:100%;height:auto;aspect-ratio:560/300}
.chart{background:var(--card);border:1px solid var(--line);border-radius:12px;
       padding:14px 16px 16px}
.chart h3{font-size:14px;margin:0 0 6px}
.empty{color:var(--muted);font-style:italic}
.stats{display:flex;flex-wrap:wrap;gap:4px 14px;margin-top:8px;
       font-size:12px;color:var(--muted)}
.stats .cv{color:var(--accent);font-weight:600}
.legend{display:flex;gap:16px;font-size:12px;color:var(--muted);margin-top:4px}
.legend .sw{display:inline-block;width:14px;height:0;border-top-width:2px;
            border-top-style:dashed;margin-right:5px;vertical-align:middle}
.legend .sw.mean{border-top-color:#e23b3b}
.legend .sw.med{border-top-color:#1d9d55}
.tick{font-size:11px;fill:#667085}
.section-title{font-size:13px;text-transform:uppercase;letter-spacing:.04em;
               color:var(--muted);margin:26px 0 10px}
code{background:#eef1f8;padding:1px 5px;border-radius:5px;font-size:13px}
@media (prefers-color-scheme:dark){
  :root{--bg:#12151c;--card:#1b1f29;--ink:#e7eaf0;--muted:#9aa4b8;--line:#2a2f3c;}
  code{background:#2a2f3c}
}
"""


def build_page(action_series, gap_series, meta, markers=None):
    charts = []
    markers = markers or {}

    # --- process-event gaps (Sysmon) : the telemetry-side evidence ---
    if gap_series:
        charts.append(('section', 'Sysmon process-event gaps (telemetry side)'))
        charts.append(('chart', svg_histogram(
            gap_series, "Inter-event gaps (Event ID 1/5/11/22)",
            unit="s", bins=18, color="#8a5cf6")))

    # --- per-action delays (metrics harness) : the richer evidence ---
    if action_series:
        charts.append(('section', 'Per-action delays (metrics harness)'))
        # stable, meaningful order first, then anything else
        order = ["ALL delays", "keystroke", "action", "read",
                 "scroll", "think", "notice"]
        keys = [k for k in order if k in action_series]
        keys += [k for k in action_series if k not in keys]
        for k in keys:
            xs = action_series[k]
            if len(xs) < 2:
                continue
            charts.append(('chart', svg_histogram(
                xs, f"{k}  (n={len(xs)})", unit="s", bins=20)))

    # --- verdict block ---
    lines = []
    if action_series.get("ALL delays"):
        a = action_series["ALL delays"]
        lines.append(
            f"Across <b>{len(a)}</b> sampled action delays the coefficient of "
            f"variation is <b>CV={cv(a):.2f}</b> "
            f"(mean {mean(a):.2f}s, median {median(a):.2f}s, "
            f"p90 {percentile(a,90):.2f}s, max {max(a):.2f}s). "
            "A metronomic/scripted agent sits near CV=0; the heavy right tail "
            "here is the human 'thinking pause' signature.")
    if gap_series:
        lines.append(
            f"On the telemetry side, <b>{len(gap_series)}</b> Sysmon "
            f"process-event gaps give <b>CV={cv(gap_series):.2f}</b> "
            f"(mean {mean(gap_series):.2f}s, max {max(gap_series):.2f}s). "
            "Clean runs log only a handful of process events, so this series is "
            "sparse by design — the per-action histogram above is the fuller "
            "picture of the same timing model.")
    if markers:
        tot = sum(markers.values())
        detail = ", ".join(f"{k}&times;{v}" for k, v in sorted(markers.items()))
        lines.append(
            f"<span style='color:var(--muted);font-size:13px'>Excluded "
            f"<b>{tot}</b> action-call markers ({detail}) — these are event "
            f"timestamps with no duration, not sampled delays, so they're left "
            f"out of the histograms above.</span>")
    if not lines:
        lines.append("No timing data was found. Run <code>python test.py</code> "
                     "to produce <code>metrics_log.txt</code>, and/or run "
                     "<code>validate.ps1</code> to produce "
                     "<code>validation\\gaps_*.csv</code>, then re-run this.")

    body = []
    body.append('<div class="wrap">')
    body.append('<h1>Synthetic Human Agent — timing evidence</h1>')
    body.append(f'<p class="sub">{html.escape(meta)}</p>')
    body.append('<div class="verdict"><h2>Why this looks human</h2>'
                + "".join(f"<p>{ln}</p>" for ln in lines) + '</div>')

    # render sections + their chart grids
    i = 0
    while i < len(charts):
        kind, payload = charts[i]
        if kind == 'section':
            body.append(f'<div class="section-title">{html.escape(payload)}</div>')
            body.append('<div class="grid">')
            i += 1
            while i < len(charts) and charts[i][0] == 'chart':
                body.append(charts[i][1])
                i += 1
            body.append('</div>')
        else:
            i += 1
    body.append('</div>')

    return (
        "<!doctype html><html lang='en'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        "<title>Timing evidence</title>"
        f"<style>{PAGE_CSS}</style></head><body>"
        + "".join(body) +
        "</body></html>"
    )


# ----------------------------- main -----------------------------

def main(argv):
    metrics_path = None
    gaps_path = None

    args = argv[1:]
    if args:
        for a in args:
            low = a.lower()
            if low.endswith(".csv"):
                gaps_path = a
            elif low.endswith(".txt"):
                metrics_path = a
    if metrics_path is None and gaps_path is None:
        metrics_path, gaps_path = find_default_inputs()

    action_series = {}
    markers = {}
    gap_series = []
    found = []

    if metrics_path and os.path.exists(metrics_path):
        action_series, markers = load_metrics_log(metrics_path)
        found.append(os.path.basename(metrics_path))
    if gaps_path and os.path.exists(gaps_path):
        gap_series = load_gaps_csv(gaps_path)
        found.append(os.path.basename(gaps_path))

    meta = ("Sources: " + ", ".join(found)) if found else \
           "No input files found (looked for metrics_log.txt and validation\\gaps_*.csv)."

    out = "timing_report.html"
    page = build_page(action_series, gap_series, meta, markers)
    with open(out, "w", encoding="utf-8") as f:
        f.write(page)

    print("Wrote", out)
    print("Sources used:", ", ".join(found) if found else "(none)")
    if action_series.get("ALL delays"):
        a = action_series["ALL delays"]
        print(f"  action delays: n={len(a)}  CV={cv(a):.2f}  "
              f"mean={mean(a):.2f}s  max={max(a):.2f}s")
    if markers:
        print("  excluded action markers:",
              ", ".join(f"{k}x{v}" for k, v in sorted(markers.items())))
    if gap_series:
        print(f"  Sysmon gaps:   n={len(gap_series)}  CV={cv(gap_series):.2f}  "
              f"mean={mean(gap_series):.2f}s  max={max(gap_series):.2f}s")
    print("Open", out, "in Chrome on the VM.")


if __name__ == "__main__":
    main(sys.argv)