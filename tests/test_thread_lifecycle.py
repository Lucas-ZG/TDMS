"""Phase 2-B round 2: QThread lifecycle on close timeout and on a slow thread shutdown.

Rules under test (see tdms.TDMSFilterApp._release_thread / closeEvent / _complete):
  * the QThread/worker references are cleared and deleteLater() is called ONLY after thread.wait() succeeded;
  * on timeout the window keeps the thread, the worker and `_processing`, refuses to close and shows
    "取消中，等待目前階段結束…"; when the worker ends by itself the normal completion path restores the UI;
  * no "QThread: Destroyed while thread is still running" message is ever emitted.
"""
import contextlib
import io
import os
import sys
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("MPLBACKEND", "Agg")
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtCore import QThread, qInstallMessageHandler
from PySide6.QtGui import QCloseEvent

import tdms
from test_background_processing import (BackgroundTestCase, Gate, finish, pump_until,  # noqa: F401 (module fixtures)
                                         setUpModule, tearDownModule)
from test_regression_export_cli import SMALL_08

DESTROYED = "Destroyed while thread is still running"


class CloseTimeoutTests(BackgroundTestCase):
    def setUp(self):
        super().setUp()
        self.qt_messages = []
        qInstallMessageHandler(lambda mode, context, message: self.qt_messages.append(message))
        self.addCleanup(qInstallMessageHandler, None)
        self.deleted = {"thread": 0, "worker": 0}
        real_thread_delete, real_worker_delete = tdms.QThread.deleteLater, tdms.ProcessWorker.deleteLater

        def thread_delete(obj):
            self.deleted["thread"] += 1
            return real_thread_delete(obj)

        def worker_delete(obj):
            self.deleted["worker"] += 1
            return real_worker_delete(obj)

        for target, name, fake in ((tdms.QThread, "deleteLater", thread_delete), (tdms.ProcessWorker, "deleteLater", worker_delete)):
            patcher = patch.object(target, name, fake)
            patcher.start()
            self.addCleanup(patcher.stop)

    def test_close_timeout_keeps_thread_refuses_close_and_recovers_when_worker_ends(self):
        source = self.copy(SMALL_08)
        window = self.new_window(source)
        gate = Gate(tdms._read_frame, hold_index=0)
        self.addCleanup(gate.release.set)                     # never leave the worker blocked after a failure
        with patch.object(tdms, "_read_frame", gate):
            with contextlib.redirect_stdout(io.StringIO()):
                window.btn_run.click()
            self.assertTrue(pump_until(gate.entered.is_set, 10000))
            thread, worker = window._thread, window._worker

            # --- close while the current stage is blocked, with the wait timeout shortened to 200 ms
            event = QCloseEvent()
            with patch.object(tdms, "THREAD_WAIT_TIMEOUT_MS", 200, create=True):
                window.closeEvent(event)
            self.assertFalse(event.isAccepted(), "close must be refused while the worker is still running")
            self.assertIs(window._thread, thread)
            self.assertIs(window._worker, worker)
            self.assertTrue(thread.isRunning())
            self.assertTrue(window.is_processing())
            self.assertEqual(self.deleted, {"thread": 0, "worker": 0}, "deleteLater() before the thread ended")
            self.assertTrue(window.label_status.text().startswith("取消中"), window.label_status.text())
            self.assertFalse(window.btn_run.isEnabled())
            self.assertFalse(window.btn_cancel.isEnabled())
            self.assertTrue(all(not g.isEnabled() for g in window._input_groups))
            self.assertEqual(self.dialogs, [])

            # --- the user may press close again; still refused while the worker is blocked
            again = QCloseEvent()
            with patch.object(tdms, "THREAD_WAIT_TIMEOUT_MS", 100, create=True):
                window.closeEvent(again)
            self.assertFalse(again.isAccepted())
            self.assertIs(window._thread, thread)

            # --- release the worker: the normal completion path cleans up and restores the UI by itself
            box = []
            window.processing_finished.connect(lambda summary: box.append(summary))
            gate.release.set()
            self.assertTrue(pump_until(lambda: bool(box), 30000))
        summary = box[0]
        self.assertTrue(summary.cancelled)                    # the close requested a cancel
        self.assertIsNone(window._thread)
        self.assertIsNone(window._worker)
        self.assertFalse(window.is_processing())
        self.assertTrue(window.btn_run.isEnabled())
        self.assertFalse(window.btn_cancel.isEnabled())
        self.assertTrue(all(g.isEnabled() for g in window._input_groups))
        self.assertEqual(self.deleted, {"thread": 1, "worker": 1})
        self.assertEqual(self.open_files_mock.call_count, 0)
        final = QCloseEvent()
        window.closeEvent(final)
        self.assertTrue(final.isAccepted(), "close is accepted once the worker has ended")

        # --- the same window can run again afterwards (cancel state is reset: it is a brand-new worker)
        self.dialogs.clear()
        summary2 = self.click_run_and_finish(window)
        self.assertFalse(summary2.cancelled)
        self.assertEqual((summary2.succeeded, summary2.skipped), (1, 0))
        self.assertEqual(len(window.generated_excel_paths), 1)
        self.assertEqual(len(window.generated_image_paths), 1)
        self.assertEqual([m for m in self.qt_messages if DESTROYED in m], [])

    def test_normal_close_still_closes_cleanly_when_the_worker_ends_quickly(self):
        window = self.new_window(self.copy(SMALL_08))
        gate = Gate(tdms._read_frame, hold_index=0)
        self.addCleanup(gate.release.set)
        with patch.object(tdms, "_read_frame", gate):
            with contextlib.redirect_stdout(io.StringIO()):
                window.btn_run.click()
            self.assertTrue(pump_until(gate.entered.is_set, 10000))
            threading.Timer(0.2, gate.release.set).start()
            event = QCloseEvent()
            window.closeEvent(event)                          # default 10 s timeout, worker ends after ~0.2 s
        self.assertTrue(event.isAccepted())
        self.assertIsNone(window._thread)
        self.assertIsNone(window._worker)
        self.assertEqual(self.deleted, {"thread": 1, "worker": 1})
        self.assertEqual([m for m in self.qt_messages if DESTROYED in m], [])


class SlowThreadShutdownTests(BackgroundTestCase):
    """The completion path (finished / failed signal) must not clean up before thread.wait() succeeded either."""

    def run_with_one_failed_wait(self, make_failure):
        window = self.new_window(self.copy(SMALL_08))
        observed = []
        real_wait = QThread.wait
        deleted = []
        real_thread_delete = QThread.deleteLater

        def fake_wait(thread, *args, **kwargs):
            if not observed:
                observed.append({"processing": window.is_processing(), "run_enabled": window.btn_run.isEnabled(),
                                 "has_thread": window._thread is not None, "deleted": len(deleted)})
                return False                                  # pretend the thread has not finished yet
            return real_wait(thread, *args, **kwargs)

        def counting_delete(thread):
            deleted.append(1)
            return real_thread_delete(thread)

        with patch.object(QThread, "wait", fake_wait), patch.object(QThread, "deleteLater", counting_delete), make_failure():
            box = []
            window.processing_finished.connect(lambda summary: box.append(summary))
            with contextlib.redirect_stdout(io.StringIO()):
                window.btn_run.click()
                self.assertTrue(pump_until(lambda: bool(box), 30000))
        self.assertEqual(len(box), 1, "processing_finished must be emitted exactly once")
        self.assertEqual(len(observed), 1)
        self.assertEqual(observed[0], {"processing": True, "run_enabled": False, "has_thread": True, "deleted": 0})
        self.assertIsNone(window._thread)
        self.assertFalse(window.is_processing())
        self.assertTrue(window.btn_run.isEnabled())
        self.assertTrue(all(g.isEnabled() for g in window._input_groups))
        self.assertEqual(len(deleted), 1)
        return window, box[0]

    def test_finished_path_keeps_references_until_the_thread_really_ended(self):
        window, summary = self.run_with_one_failed_wait(contextlib.nullcontext)
        self.assertEqual((summary.succeeded, summary.skipped), (1, 0))
        self.assertEqual(len(window.generated_excel_paths), 1)

    def test_failed_path_keeps_references_until_the_thread_really_ended(self):
        window, summary = self.run_with_one_failed_wait(
            lambda: patch.object(tdms.ProcessWorker, "_run", side_effect=RuntimeError("boom")))
        self.assertIsNone(summary)
        self.assertEqual(len(self.dialogs), 1)
        self.assertIn("boom", self.dialogs[0][2])


class CancelPrimitiveTests(unittest.TestCase):
    def test_cancel_flag_is_a_threading_event(self):
        settings = tdms.RunSettings(("TIME",), True, True, True, True, False, False, False, True, False, False)
        worker = tdms.ProcessWorker([], settings)
        self.assertIsInstance(worker._cancel_event, threading.Event)
        self.assertFalse(worker._cancel_event.is_set())
        worker.request_cancel()
        self.assertTrue(worker._cancel_event.is_set())


if __name__ == "__main__":
    unittest.main()
