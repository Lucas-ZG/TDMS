"""One-click test runner: python tests/run_all.py   (from anywhere; exit code 0 = all tests passed).

Runs the whole unittest suite (read -> DataFrame, xlsx cells, PNG, CLI parsing, background processing, thread lifecycle,
worker/thread destruction). Optional large-capture tests run when the 4.99 MB capture is found
(TDMS_LARGE_SAMPLE or '<project>/test file/'); otherwise they are skipped with a message.
"""
import os
import sys
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("MPLBACKEND", "Agg")
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))

if __name__ == "__main__":
    suite = unittest.defaultTestLoader.discover(str(HERE), top_level_dir=str(HERE))
    result = unittest.TextTestRunner(verbosity=1).run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)
