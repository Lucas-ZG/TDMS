"""Regression tests: after a processing run the worker (ProcessWorker) and the QThread are REALLY destroyed
(not just "deleteLater() was called"), runs do not accumulate objects, and the close-timeout branch keeps both alive.

Background: Qt never processes a deferred delete that is posted to the event queue of a QThread that has already
finished, so deleting the worker from the GUI thread after thread.wait() leaked it. The worker is now deleted when the
thread itself announces `finished` (see TDMSFilterApp.process_files).
"""
import contextlib
import gc
import io
import os
import sys
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("MPLBACKEND", "Agg")
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import shiboken6
from PySide6.QtCore import QCoreApplication, QEvent, QThread
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import QApplication

import tdms
from test_background_processing import (BackgroundTestCase, Gate, finish, pump_until,  # noqa: F401 (module fixtures)
                                         setUpModule, tearDownModule)
from test_regression_export_cli import SMALL_08


def settle(ms=300):
    """Process events including deferred deletes, like a running application would."""
    end = time.monotonic() + ms / 1000
    while time.monotonic() < end:
        QApplication.instance().processEvents()
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        time.sleep(0.01)


def live_workers():
    gc.collect()
    return sum(1 for o in gc.get_objects() if isinstance(o, tdms.ProcessWorker) and shiboken6.isValid(o))


class DestroyTests(BackgroundTestCase):
    def watch(self, window):
        """Call right after the run was started: returns (worker, thread, destroyed-dict)."""
        worker, thread = window._worker, window._thread
        destroyed = {}
        worker.destroyed.connect(lambda *_: destroyed.__setitem__("worker", True))
        thread.destroyed.connect(lambda *_: destroyed.__setitem__("thread", True))
        return worker, thread, destroyed

    def assert_destroyed(self, worker, thread, destroyed, scenario):
        settle()
        self.assertTrue(destroyed.get("worker"), f"{scenario}: worker 'destroyed' signal not received")
        self.assertTrue(destroyed.get("thread"), f"{scenario}: thread 'destroyed' signal not received")
        self.assertFalse(shiboken6.isValid(worker), f"{scenario}: worker is still a valid C++ object")
        self.assertFalse(shiboken6.isValid(thread), f"{scenario}: thread is still a valid C++ object")
        self.assertIsNone(self.windows[-1]._thread)
        self.assertIsNone(self.windows[-1]._worker)

    def test_worker_and_thread_are_destroyed_after_a_normal_run(self):
        window = self.new_window(self.copy(SMALL_08))
        with contextlib.redirect_stdout(io.StringIO()):
            window.btn_run.click()
            worker, thread, destroyed = self.watch(window)
            finish(window)
        self.assert_destroyed(worker, thread, destroyed, "normal")

    def test_worker_and_thread_are_destroyed_after_a_cancelled_run(self):
        window = self.new_window(self.copy(SMALL_08, "a.tdms"), self.copy(SMALL_08, "b.tdms"))
        gate = Gate(tdms._read_frame, hold_index=0)
        self.addCleanup(gate.release.set)
        with patch.object(tdms, "_read_frame", gate), contextlib.redirect_stdout(io.StringIO()):
            window.btn_run.click()
            worker, thread, destroyed = self.watch(window)
            self.assertTrue(pump_until(gate.entered.is_set, 10000))
            window.btn_cancel.click()
            gate.release.set()
            summary = finish(window)
        self.assertTrue(summary.cancelled)
        self.assert_destroyed(worker, thread, destroyed, "cancel")

    def test_worker_and_thread_are_destroyed_after_a_failed_run(self):
        window = self.new_window(self.copy(SMALL_08))
        with patch.object(tdms.ProcessWorker, "_run", side_effect=RuntimeError("boom")), contextlib.redirect_stdout(io.StringIO()):
            window.btn_run.click()
            worker, thread, destroyed = self.watch(window)
            self.assertIsNone(finish(window))
        self.assert_destroyed(worker, thread, destroyed, "failure")

    def test_consecutive_runs_do_not_accumulate_workers(self):
        window = self.new_window(self.copy(SMALL_08))
        baseline = live_workers()                              # other tests may legitimately keep objects alive
        for run in range(8):
            with contextlib.redirect_stdout(io.StringIO()):
                window.btn_run.click()
                finish(window)
            settle(150)
            self.assertEqual(live_workers(), baseline, f"live ProcessWorker objects after run {run + 1}")
            time.sleep(1.02)                                  # output names carry a per-second timestamp

    def test_close_timeout_keeps_worker_and_thread_alive_then_destroys_them(self):
        window = self.new_window(self.copy(SMALL_08))
        gate = Gate(tdms._read_frame, hold_index=0)
        self.addCleanup(gate.release.set)
        with patch.object(tdms, "_read_frame", gate), contextlib.redirect_stdout(io.StringIO()):
            window.btn_run.click()
            worker, thread, destroyed = self.watch(window)
            self.assertTrue(pump_until(gate.entered.is_set, 10000))
            event = QCloseEvent()
            with patch.object(tdms, "THREAD_WAIT_TIMEOUT_MS", 200):
                window.closeEvent(event)
            settle(200)
            self.assertFalse(event.isAccepted())
            self.assertTrue(shiboken6.isValid(worker) and shiboken6.isValid(thread), "objects destroyed while the thread still runs")
            self.assertEqual(destroyed, {})
            self.assertTrue(thread.isRunning())
            gate.release.set()
            finish(window)
        self.assert_destroyed(worker, thread, destroyed, "close-timeout then natural end")


if __name__ == "__main__":
    unittest.main()
