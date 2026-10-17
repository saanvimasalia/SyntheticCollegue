import sys
import sequencer

budget = int(sys.argv[1]) if len(sys.argv) > 1 else 880
sequencer.run(budget_s=budget)
