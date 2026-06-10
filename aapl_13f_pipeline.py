#!/usr/bin/env python3
# Backward-compatible wrapper -> general engine build_13f.py with Apple defaults.
import os,sys,subprocess
d=os.path.dirname(os.path.abspath(__file__))
folder=sys.argv[1] if (len(sys.argv)>1 and not sys.argv[1].startswith('-')) else d
sys.exit(subprocess.run([sys.executable,os.path.join(d,'build_13f.py'),'--folder',folder]).returncode)
