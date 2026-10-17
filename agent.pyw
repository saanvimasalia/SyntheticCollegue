import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
if os.path.exists("STOP"):
    os.remove("STOP")
import ctypes, sys
_m = ctypes.windll.kernel32.CreateMutexW(None, False, "Local\\SyntheticAgent")
if ctypes.windll.kernel32.GetLastError() == 183:
    sys.exit(0)
import sequencer
try:
    sequencer.run(budget_s=900)
except Exception as e:
    sequencer.log("agent crashed: %r" % (e,))