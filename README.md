# Synthetic Human Agent

A pure-Python tool that performs a realistic ~15-minute routine of office and
web-browsing activity on a Windows VM, so that the resulting **Sysmon / Windows
Event Log telemetry is indistinguishable from a real human user** — correct
process parentage, human-like timing, and no scripting artifacts.

Built for the NTU × ST Engineering "Building a Synthetic Human Agent" challenge
(team05).

---

## What it does

Each run assembles a structurally different but plausible "day at the desk" from
a pool of complete, end-to-end tasks:

- **Web browsing** (Chrome): visit a few sites, scroll in bursts with reading
  dwells, follow a link and come back, find-on-page, bookmark, switch tabs, and
  occasionally save a page.
- **Document writing** (OpenOffice Writer): type a multi-paragraph document with
  a realistic typo/backspace model, save it under a plausible filename.
- **Spreadsheet** (OpenOffice Calc): build a small table with per-row formulas
  and a SUM footer, revise a cell, save as `.ods`. The saved file is read back
  and every cell/formula is verified.

A stochastic sequencer shuffles the order, picks a random "persona" (which sets
the pace), inserts idle gaps and one longer away-from-keyboard pause, and keeps
the whole run inside a time budget.

## Design principles (these are the scored criteria)

- **Authentic process parentage.** Apps are launched via the Windows **Run
  dialog (Win+R) driven by `SendInput`**, so the real parent is `explorer.exe` —
  never `subprocess`, `Start-Process`, or `os.startfile`. The agent itself runs
  via `pythonw.exe` from a shortcut, so it too is parented to `explorer.exe`.
- **Human timing.** Every human-facing delay is sampled from a **log-normal
  distribution** with a heavy right tail (thinking pauses) — never a fixed
  `sleep(n)`.
- **Lean footprint.** Pure Python standard library — no third-party packages and
  **no LLM in the loop**. CPU/RAM stay minimal (the agent's own overhead is
  near-zero; the only load on the graph is the real applications doing real
  work).
- **Complete tasks.** Each scenario is a full end-to-end task (open → act →
  save → close), not a fragment.
- **No Task Scheduler.** The agent is launched manually so the scheduler service
  never becomes the parent of the launch.
- **Cold-launch discipline.** Every app is fully closed before it is launched
  again, so each launch produces a fresh `explorer → app` process-create event.

---

## Requirements

- Windows 10/11
- Python 3.12 (on `PATH`; `pythonw.exe` available)
- Google Chrome
- Apache OpenOffice 4 (Writer + Calc)
- Sysmon installed and logging (`Sysmon64.exe -accepteula -i`) — for validation
- The range's root CA installed in Trusted Root (for the mock browsing sites)

No `pip install` is needed — the agent uses the standard library only.

---

## How to run

Launch the agent from the **desktop shortcut** (a `pythonw` launcher), **not**
from a terminal. Launching from a terminal would parent the agent to the shell;
the shortcut gives the clean `explorer → pythonw` parent.

Equivalent command line (for reference):

```
pythonw agent.py [budget_seconds]
```

`budget_seconds` defaults to 900 (~15 minutes).

Before a run, for clean telemetry:
- Close Chrome, Writer, and Calc (cold-launch discipline).
- Close File Explorer, Settings, and any elevated shell (an elevated window in
  the foreground can block input injection).
- Don't touch the keyboard or mouse while it runs.

**Stopping a run:** create an empty file named `STOP` (no extension) in the
project folder. The agent checks for it during every delay and exits cleanly
within ~1 second. Delete it before the next run.

---

## How to check a run

Run the validator in an **elevated** PowerShell (reading the Sysmon log requires
elevation), from the project folder:

```
Set-ExecutionPolicy -Scope Process Bypass
.\validate.ps1
```

It reads the Sysmon log for the most recent run and checks process parentage,
cold-launch discipline, no script/shell parents, run length, timing spread,
document save, and the agent's own process tree. Results are printed and written
to `validation\` as CSVs.

**Timing report** (human-readable histograms of the delay distribution, pure
stdlib → self-contained HTML):

```
python make_histogram.py
```

Open the generated `timing_report.html` in Chrome.

**Resource usage** (CPU/RAM of the agent process):

```
.\monitor.ps1
```

---

## Files

| File | Role |
|------|------|
| `agent.py` | Entry point — parses the time budget, calls the sequencer. |
| `agent.pyw` | `pythonw` launcher target (no console; explorer-parented). |
| `engine.py` | Core: `SendInput` keyboard/mouse, Win+R `launch`, `wait_for_window`, log-normal `human_delay`, persona `SPEED`. |
| `sequencer.py` | Assembles a different ~15-min routine each run: persona, shuffled block order, guaranteed coverage, one AFK pause, time budget. |
| `web.py` | Web-surfing scenario (Chrome). |
| `office.py` | Office scenarios (Writer document + Calc spreadsheet) with typo model and read-back verification. |
| `validate.ps1` | Telemetry validator (run elevated). |
| `make_histogram.py` | Builds `timing_report.html` from the run's timing logs. |
| `monitor.ps1` | Samples the agent's CPU/RAM to a CSV. |

---

## How the "arbitrary routine" works

The sequencer runs with `GUARANTEE = True`: every run performs one web session,
one document and one spreadsheet, in **shuffled** order, with exactly one
away-from-keyboard pause. Everything a judge reads as human — order, timing,
which sites, the typos, the filenames, the persona — stays random, so no two
runs look alike, while coverage of all three task types stays reliable.
