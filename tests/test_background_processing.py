"""Tests for Phase 2-B: processing runs on a worker QThread (progress, cancel, one summary dialog), plotting uses
matplotlib's Figure (no pyplot), and the GUI thread stays responsive.

Run from the project root:  python -m unittest discover -s tests -v
Portable: uses the committed fixtures; the large-capture tests are skipped unless the 4.99 MB capture is found
(TDMS_LARGE_SAMPLE or '<project>/test file/', see test_regression_export_cli.py).
Set TDMS_TEST_METRICS_FILE=<path> to append the measured heartbeat gaps of the responsiveness test.
"""
import contextlib
import dataclasses
import gc
import io
import os
import shutil
import struct
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("MPLBACKEND", "Agg")
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.figure import Figure
from PIL import Image
from PySide6.QtCore import QEventLoop, QThread, QTimer, qInstallMessageHandler
from PySide6.QtWidgets import QApplication, QCheckBox, QPushButton

import tdms
from test_regression_export_cli import (FIXTURES, SMALL_08, SMALL_1459, cells_of, find_large_sample,
                                         LARGE_SKIP_MESSAGE, load_frame, make_window, quiet)

_app = None
HEARTBEAT_THRESHOLD_S = 0.25   # see test_large_all_channels_gui_thread_stays_responsive for the evidence behind it


def setUpModule():
    global _app
    _app = QApplication.instance() or QApplication([])


def tearDownModule():
    global _app
    _app = None


def pump_until(predicate, timeout_ms=30000):
    loop = QEventLoop()
    timer = QTimer()
    timer.setInterval(5)
    deadline = time.monotonic() + timeout_ms / 1000

    def check():
        if predicate() or time.monotonic() > deadline:
            loop.quit()

    timer.timeout.connect(check)
    timer.start()
    loop.exec()
    timer.stop()
    return predicate()


def finish(window, timeout_ms=60000):
    """Spin the Qt event loop until the window reports `processing_finished`; returns the RunSummary (or None)."""
    box = []
    window.processing_finished.connect(lambda summary: box.append(summary))
    ok = pump_until(lambda: bool(box), timeout_ms)
    if not ok:
        raise AssertionError("processing did not finish within %d ms" % timeout_ms)
    return box[0]


class Gate:
    """Wraps tdms._read_frame; the call with index `hold_index` blocks until `release` is set."""

    def __init__(self, real, hold_index=None, sleep=0.0):
        self.real, self.hold_index, self.sleep = real, hold_index, sleep
        self.calls = 0
        self.entered = threading.Event()
        self.release = threading.Event()

    def __call__(self, path, wanted):
        index = self.calls
        self.calls += 1
        if index == self.hold_index:
            self.entered.set()
            self.release.wait(30)
        if self.sleep:
            time.sleep(self.sleep)
        return self.real(path, wanted)


def legacy_pyplot_plot(df, filename, folder, flags):
    """Reference: the pre-Phase-2-B pyplot implementation of plot_and_save (kept verbatim in spirit) for pixel comparison."""
    fig, ax = plt.subplots(figsize=(10, 6))
    t = df["TIME"]
    targets = {"M1 T-Case": flags["case"], "M1 T-Evap": flags["evap"], "M1 TSD": flags["tsd"], "M1 Set Temp": flags["set"]}
    for col, show in targets.items():
        if show and col in df.columns:
            ax.plot(t, df[col], label=col)
            if flags["mm"] and col == "M1 TSD":
                y = df[col].dropna()
                if not y.empty:
                    imax, imin = y.idxmax(), y.idxmin()
                    ax.annotate(f"Max: {y[imax]:.1f}", xy=(t[imax], y[imax]), xytext=(5, 5), textcoords='offset points', fontsize=9, color='red')
                    ax.annotate(f"Min: {y[imin]:.1f}", xy=(t[imin], y[imin]), xytext=(5, -15), textcoords='offset points', fontsize=9, color='blue')
    if flags["gtsd"] and "M1 T-Evap" in df.columns and "M1 TSD" in df.columns:
        gap = (df["M1 T-Evap"] - df["M1 TSD"]).abs().dropna()
        if not gap.empty:
            idx = gap.idxmax()
            ax.annotate(f"Max Evap-TSD Gap: {gap[idx]:.1f} @ Time: {t[idx]:.1f}", xy=(t[idx], df["M1 T-Evap"][idx]), xytext=(10, 10),
                        textcoords='offset points', fontsize=9, color='purple', arrowprops=dict(arrowstyle='->', color='purple'))
    if flags["gcase"] and "M1 T-Evap" in df.columns and "M1 T-Case" in df.columns:
        gap = (df["M1 T-Evap"] - df["M1 T-Case"]).abs().dropna()
        if not gap.empty:
            idx = gap.idxmax()
            ax.annotate(f"Max Evap-Case Gap: {gap[idx]:.1f} @ Time: {t[idx]:.1f}", xy=(t[idx], df["M1 T-Evap"][idx]), xytext=(10, -25),
                        textcoords='offset points', fontsize=9, color='green', arrowprops=dict(arrowstyle='->', color='green'))
    ax.set_xlabel("Time"); ax.set_ylabel("Value")
    ax.set_title(f"TDMS Thermal Data - {filename}")
    ax.legend(); ax.grid(True); fig.tight_layout()
    path = os.path.join(folder, "legacy_reference.png")
    plt.savefig(path)
    plt.close(fig)
    return path


def png_array(path):
    return np.asarray(Image.open(path).convert("RGBA"))


FLAG_SETS = {
    "default": dict(case=True, evap=True, tsd=True, set=True, mm=False, gtsd=False, gcase=False),
    "all_curves_all_annotations": dict(case=True, evap=True, tsd=True, set=True, mm=True, gtsd=True, gcase=True),
    "tsd_only_minmax": dict(case=False, evap=False, tsd=True, set=False, mm=True, gtsd=False, gcase=False),
}


def apply_flags(window, flags):
    window.show_t_case.setChecked(flags["case"]); window.show_t_evap.setChecked(flags["evap"])
    window.show_t_tsd.setChecked(flags["tsd"]); window.show_t_set.setChecked(flags["set"])
    window.annotate_min_max.setChecked(flags["mm"]); window.annotate_gap_tsd.setChecked(flags["gtsd"])
    window.annotate_gap_case.setChecked(flags["gcase"])


def interleaved_cobra_file(rows=6, declared_rows=None):
    """In-memory interleaved TDMS with the tool's 'COBRA THERMAL DATA' group; declared_rows < rows => recoverable."""
    channels = ("TIME", "M1 T-Case", "M1 T-Evap", "M1 TSD", "M1 Set Temp")

    def text(value):
        raw = value.encode("utf-8")
        return struct.pack("<I", len(raw)) + raw

    meta = struct.pack("<I", len(channels))
    for name in channels:
        meta += text("/'COBRA THERMAL DATA'/'%s'" % name)
        meta += struct.pack("<I", 20) + struct.pack("<II", 9, 1) + struct.pack("<Q", 1) + struct.pack("<I", 0)
    matrix = np.arange(rows * len(channels), dtype="<f4").reshape(rows, len(channels)) + 0.5
    declared = rows if declared_rows is None else declared_rows
    lead = (b"TDSm" + struct.pack("<I", 0x2E) + struct.pack("<I", 4713)
            + struct.pack("<Q", len(meta) + declared * 4 * len(channels)) + struct.pack("<Q", len(meta)))
    return lead + meta + matrix.tobytes()


class BackgroundTestCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)
        self.main_thread = QApplication.instance().thread()
        self.dialogs = []
        self.windows = []

        def recorder(kind):
            def record(parent, title, text, *args, **kwargs):
                self.dialogs.append((kind, title, text, QThread.currentThread() is self.main_thread))
            return record

        for kind in ("critical", "information", "warning"):
            patcher = patch.object(tdms.QMessageBox, kind, side_effect=recorder(kind))
            patcher.start()
            self.addCleanup(patcher.stop)
        self.open_files = patch.object(tdms.TDMSFilterApp, "open_files")
        self.open_files_mock = self.open_files.start()
        self.addCleanup(self.open_files.stop)
        self.addCleanup(self.close_windows)

    def close_windows(self):
        for window in self.windows:
            if window.is_processing():
                window.cancel_processing()
                with contextlib.suppress(AssertionError):
                    finish(window, 20000)
            window.close()
            window.deleteLater()

    def new_window(self, *files, all_channels=False):
        window = make_window()
        self.windows.append(window)
        window.file_paths = [str(f) for f in files]
        window.enable_excel.setChecked(True)
        window.auto_open_excel.setChecked(False)
        window.auto_open_plot.setChecked(False)
        if files and all(Path(f).is_file() for f in files):
            quiet(window._rebuild_checklist_from_all_files)
            (window._select_all_channels if all_channels else window._deselect_all_channels)()
        return window

    def copy(self, source, name=None, subdir=""):
        folder = self.tmp / subdir if subdir else self.tmp
        folder.mkdir(parents=True, exist_ok=True)
        target = folder / (name or source.name)
        shutil.copyfile(source, target)
        return target

    def run_button(self, window):
        return window.btn_run

    def click_run_and_finish(self, window, timeout_ms=60000):
        box = []
        window.processing_finished.connect(lambda summary: box.append(summary))
        with contextlib.redirect_stdout(io.StringIO()):
            window.btn_run.click()
            if not pump_until(lambda: bool(box), timeout_ms):
                raise AssertionError("processing did not finish")
        return box[0]


class EquivalenceTests(BackgroundTestCase):
    """Asynchronous Figure-based output must equal the legacy pyplot output and the pandas reference."""

    def check_fixture(self, fixture, flags_name):
        source = self.copy(fixture, subdir=f"{fixture.name[-12:-5]}_{flags_name}")
        window = self.new_window(source)
        apply_flags(window, FLAG_SETS[flags_name])
        df, _ = load_frame(window, source, all_channels=False)
        summary = self.click_run_and_finish(window)
        self.assertFalse(summary.cancelled)
        self.assertEqual(self.dialogs, [])
        self.assertEqual(len(window.generated_image_paths), 1)
        self.assertEqual(len(window.generated_excel_paths), 1)
        # PNG: pixel-identical to the legacy pyplot rendering of the same data/settings
        reference = legacy_pyplot_plot(df, source.name, str(self.tmp), FLAG_SETS[flags_name])
        new_pixels, old_pixels = png_array(window.generated_image_paths[0]), png_array(reference)
        self.assertEqual(new_pixels.shape, old_pixels.shape)
        self.assertTrue(np.array_equal(new_pixels, old_pixels), f"{fixture.name[-12:]} / {flags_name}: PNG pixels differ")
        # xlsx: identical to what DataFrame.to_excel writes, cell by cell
        reference_xlsx = self.tmp / f"reference_{fixture.name[-12:-5]}_{flags_name}.xlsx"
        df.to_excel(reference_xlsx, index=False)
        self.assertEqual(cells_of(window.generated_excel_paths[0]), cells_of(reference_xlsx))

    def test_png_and_xlsx_equivalence_on_small_fixtures(self):
        for fixture in (SMALL_08, SMALL_1459):
            for flags_name in FLAG_SETS:
                with self.subTest(fixture=fixture.name[-12:], flags=flags_name):
                    self.check_fixture(fixture, flags_name)

    def test_png_and_xlsx_equivalence_on_large_capture(self):
        large = find_large_sample()
        if large is None:
            self.skipTest(LARGE_SKIP_MESSAGE)
        for flags_name in ("default", "all_curves_all_annotations"):
            with self.subTest(flags=flags_name):
                self.check_fixture(large, flags_name)

    def test_output_names_locations_and_order_for_a_batch(self):
        first = self.copy(SMALL_08, "a_first.tdms")
        second = self.copy(SMALL_1459, "b_second.tdms")
        window = self.new_window(first, second)
        summary = self.click_run_and_finish(window)
        self.assertEqual([Path(p).stem.rsplit("_", 2)[0] for p in window.generated_excel_paths], ["a_first", "b_second"])
        self.assertEqual([Path(p).stem.rsplit("_", 2)[0] for p in window.generated_image_paths], ["a_first", "b_second"])
        for path in window.generated_excel_paths + window.generated_image_paths:
            self.assertEqual(Path(path).parent, self.tmp)
        self.assertEqual((summary.succeeded, summary.skipped, summary.recovered), (2, 0, 0))
        self.assertEqual(self.dialogs, [])

    def test_legacy_methods_still_work_synchronously(self):
        source = self.copy(SMALL_1459)
        window = self.new_window(source)
        df, status = load_frame(window, source, all_channels=False)
        self.assertEqual((status, df.shape), ("normal", (16, 5)))
        window.plot_and_save(df, source.name, str(self.tmp))
        window.export_to_excel(df, str(source))
        self.assertEqual(len(window.generated_image_paths), 1)
        self.assertEqual(len(window.generated_excel_paths), 1)
        self.assertEqual(png_array(window.generated_image_paths[0])[:8].shape[-1], 4)


class UiBehaviourTests(BackgroundTestCase):
    def test_widgets_exist_and_idle_state(self):
        window = self.new_window()
        self.assertEqual(window.windowTitle(), f"TDMS小工具 {tdms.VERSION}")
        self.assertEqual(window.btn_run.text(), "▶ 執行處理 TDMS 檔案")
        self.assertEqual(window.btn_cancel.text(), "取消")
        self.assertTrue(window.btn_run.isEnabled())
        self.assertFalse(window.btn_cancel.isEnabled())
        self.assertTrue(all(g.isEnabled() for g in window._input_groups))
        self.assertEqual(window.label_status.text(), "")
        self.assertGreaterEqual(window.height(), window.minimumSizeHint().height() - 1)

    def test_running_state_then_restored_and_reentrancy_ignored(self):
        source = self.copy(SMALL_08)
        window = self.new_window(source)
        gate = Gate(tdms._read_frame, hold_index=0)
        with patch.object(tdms, "_read_frame", gate):
            with contextlib.redirect_stdout(io.StringIO()):
                window.btn_run.click()
            self.assertTrue(gate.entered.wait(10) or pump_until(gate.entered.is_set, 10000))
            pump_until(lambda: window.progress_bar.maximum() == 3 and bool(window.label_status.text()) and "讀取 TDMS" in window.label_status.text(), 5000)
            # while running
            self.assertTrue(window.is_processing())
            self.assertFalse(window.btn_run.isEnabled())
            self.assertTrue(window.btn_cancel.isEnabled())
            self.assertTrue(all(not g.isEnabled() for g in window._input_groups))
            self.assertEqual(window.progress_bar.maximum(), 3)           # read + Excel + plot
            self.assertIn("TDMS", window.label_status.text())
            self.assertIn(source.name, window.label_status.text())
            # re-entrancy: a second trigger while running is ignored
            thread_before = window._thread
            window.process_files()
            window.btn_run.click()
            self.assertIs(window._thread, thread_before)
            self.assertEqual(self.dialogs, [])
            box = []
            window.processing_finished.connect(lambda summary: box.append(summary))
            gate.release.set()
            self.assertTrue(pump_until(lambda: bool(box), 30000))
        summary = box[0]
        self.assertEqual(gate.calls, 1)                                  # only ONE run was started
        self.assertEqual((summary.succeeded, summary.skipped), (1, 0))
        self.assertEqual(len(window.generated_excel_paths), 1)
        self.assertFalse(window.is_processing())
        self.assertTrue(window.btn_run.isEnabled())
        self.assertFalse(window.btn_cancel.isEnabled())
        self.assertTrue(all(g.isEnabled() for g in window._input_groups))
        self.assertEqual(window.progress_bar.value(), window.progress_bar.maximum())
        self.assertRegex(window.label_status.text(), r"^完成：成功 1 檔、略過 0 檔、耗時 \d+\.\d s$")
        self.assertIsNone(window._thread)

    def test_three_consecutive_runs_on_one_window(self):
        source = self.copy(SMALL_1459)
        window = self.new_window(source)
        for run in range(3):
            summary = self.click_run_and_finish(window)
            self.assertEqual((summary.succeeded, summary.skipped), (1, 0), f"run {run + 1}")
            self.assertIsNone(window._thread)
            self.assertIsNone(window._worker)
            self.assertFalse(window.is_processing())
            self.assertEqual(len(window.generated_excel_paths), 1)      # lists are reset on every run
            self.assertEqual(len(window.generated_image_paths), 1)
            time.sleep(1.05)                                              # output names carry a per-second timestamp
        self.assertEqual(self.dialogs, [])

    def test_empty_file_list_shows_immediate_error_dialog(self):
        window = self.new_window()
        window.process_files()
        self.assertEqual([d[:3] for d in self.dialogs], [("critical", "錯誤", "請選擇至少一個 TDMS 檔案")])
        self.assertFalse(window.is_processing())
        self.assertIsNone(window._thread)

    def test_command_line_loaded_file_then_run_button(self):
        source = self.copy(SMALL_08)
        window = make_window()
        self.windows.append(window)
        window.auto_open_excel.setChecked(False)
        quiet(window.load_initial_tdms_file, str(source))
        self.assertEqual(window.file_paths, [str(source.resolve())])
        summary = self.click_run_and_finish(window)
        self.assertEqual((summary.succeeded, summary.skipped), (1, 0))
        self.assertEqual(len(window.generated_excel_paths), 1)
        self.assertEqual(len(window.generated_image_paths), 1)

    def test_auto_open_after_normal_completion_only_once_per_kind(self):
        source = self.copy(SMALL_08)
        window = self.new_window(source)
        window.auto_open_excel.setChecked(True)
        window.auto_open_plot.setChecked(True)
        with patch.object(tdms.plt, "show", side_effect=AssertionError("interactive plt.show() must not be used")):
            self.click_run_and_finish(window)
        calls = [c.args[0] for c in self.open_files_mock.call_args_list]
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[0], window.generated_excel_paths)
        self.assertEqual(calls[1], window.generated_image_paths)


class CancelTests(BackgroundTestCase):
    def test_cancel_after_first_file_keeps_only_first_outputs(self):
        files = [self.copy(SMALL_08, "one.tdms"), self.copy(SMALL_1459, "two.tdms"), self.copy(SMALL_08, "three.tdms")]
        window = self.new_window(*files)
        window.auto_open_excel.setChecked(True)
        window.auto_open_plot.setChecked(True)
        gate = Gate(tdms._read_frame, hold_index=1)            # hold the SECOND file right after it starts
        with patch.object(tdms, "_read_frame", gate):
            with contextlib.redirect_stdout(io.StringIO()):
                window.btn_run.click()
            self.assertTrue(pump_until(gate.entered.is_set, 20000))
            pump_until(lambda: len(window.generated_excel_paths) == 1, 5000)   # first file's signals delivered
            window.btn_cancel.click()                                            # the real cancel button
            self.assertFalse(window.btn_cancel.isEnabled())
            self.assertIn("取消中", window.label_status.text())
            box = []
            window.processing_finished.connect(lambda summary: box.append(summary))
            gate.release.set()
            self.assertTrue(pump_until(lambda: bool(box), 30000))
        summary = box[0]
        self.assertTrue(summary.cancelled)
        self.assertEqual((summary.processed, summary.total_files), (1, 3))
        self.assertEqual([Path(p).name.split("_")[0] for p in window.generated_excel_paths], ["one"])
        self.assertEqual([Path(p).name.split("_")[0] for p in window.generated_image_paths], ["one"])
        self.assertEqual(sorted(p.name.split("_")[0] for p in self.tmp.glob("*.xlsx")), ["one"])
        self.assertEqual(sorted(p.name.split("_")[0] for p in self.tmp.glob("*.png")), ["one"])
        self.assertEqual(self.open_files_mock.call_count, 0)    # cancelled => nothing is auto-opened
        self.assertEqual(self.dialogs, [])
        self.assertEqual(window.label_status.text(), "已取消：完成 1／3 檔")
        self.assertTrue(window.btn_run.isEnabled())
        self.assertFalse(window.btn_cancel.isEnabled())
        self.assertTrue(all(g.isEnabled() for g in window._input_groups))
        self.assertFalse(window.is_processing())
        self.assertEqual(gate.calls, 2)                          # the third file was never started


class ErrorAndSummaryTests(BackgroundTestCase):
    def test_bad_file_between_good_files_gives_one_summary_dialog(self):
        good1 = self.copy(SMALL_08, "good1.tdms")
        bad = self.tmp / "bad.tdms"
        bad.write_bytes(bytes(range(256)) * 8)
        good2 = self.copy(SMALL_1459, "good2.tdms")
        window = self.new_window(good1, good2)
        window.file_paths = [str(good1), str(bad), str(good2)]
        summary = self.click_run_and_finish(window)
        self.assertEqual((summary.succeeded, summary.skipped, summary.processed), (2, 1, 3))
        self.assertEqual(len(window.generated_excel_paths), 2)
        self.assertEqual(len(window.generated_image_paths), 2)
        self.assertEqual(len(self.dialogs), 1)                    # exactly ONE dialog, shown after the batch
        kind, title, text, on_main = self.dialogs[0]
        self.assertEqual((kind, title, on_main), ("critical", "錯誤", True))
        self.assertTrue(text.startswith("bad.tdms：無法讀取 TDMS 檔案："), text)
        self.assertEqual(len(text.splitlines()), 1)
        self.assertRegex(window.label_status.text(), r"^完成：成功 2 檔、略過 1 檔、耗時 ")
        self.assertTrue(window.btn_run.isEnabled())
        self.assertTrue(all(g.isEnabled() for g in window._input_groups))
        self.assertEqual(window.progress_bar.value(), window.progress_bar.maximum())   # skipped steps are accounted

    def test_clean_run_shows_no_dialog(self):
        window = self.new_window(self.copy(SMALL_08))
        self.click_run_and_finish(window)
        self.assertEqual(self.dialogs, [])

    def test_recovered_file_is_reported_in_the_summary_with_information_dialog(self):
        recovered = self.tmp / "recovered.tdms"
        recovered.write_bytes(interleaved_cobra_file(rows=6, declared_rows=4))
        window = self.new_window(recovered)
        window.file_paths = [str(recovered)]
        self.dialogs.clear()      # (building the checklist for a recovered file already shows its own notice: T06, not tested here)
        summary = self.click_run_and_finish(window)
        self.assertEqual((summary.succeeded, summary.recovered), (1, 1))
        self.assertEqual([d[:3] for d in self.dialogs],
                         [("information", "TDMS 復原", f"recovered.tdms：{tdms.RECOVERED_SEGMENT_MESSAGE}")])
        self.assertEqual(len(window.generated_excel_paths), 1)
        df = pd.read_excel(window.generated_excel_paths[0])
        self.assertEqual(len(df), 6)                              # all rows recovered, none lost

    def test_recovered_and_error_are_combined_into_one_critical_dialog(self):
        recovered = self.tmp / "recovered.tdms"
        recovered.write_bytes(interleaved_cobra_file(rows=6, declared_rows=4))
        missing = self.tmp / "missing.tdms"
        window = self.new_window(recovered)
        window.file_paths = [str(recovered), str(missing)]
        self.dialogs.clear()
        self.click_run_and_finish(window)
        self.assertEqual(len(self.dialogs), 1)
        kind, title, text, _ = self.dialogs[0]
        self.assertEqual(kind, "critical")
        self.assertIn(f"recovered.tdms：{tdms.RECOVERED_SEGMENT_MESSAGE}", text)
        self.assertIn("missing.tdms：找不到檔案：", text)

    def test_summary_is_truncated_after_20_lines(self):
        window = self.new_window()
        window.file_paths = [str(self.tmp / f"missing_{i:02d}.tdms") for i in range(25)]
        self.click_run_and_finish(window)
        self.assertEqual(len(self.dialogs), 1)
        lines = self.dialogs[0][2].splitlines()
        self.assertEqual(len(lines), 21)
        self.assertEqual(lines[-1], "另有 5 筆")
        self.assertTrue(lines[0].startswith("missing_00.tdms：找不到檔案："))

    def test_every_message_is_still_printed(self):
        window = self.new_window()
        window.file_paths = [str(self.tmp / "missing.tdms")]
        out = io.StringIO()
        box = []
        window.processing_finished.connect(lambda s: box.append(s))
        with contextlib.redirect_stdout(out):
            window.process_files()
            pump_until(lambda: bool(box), 20000)
        self.assertIn("[TDMS 略過] missing.tdms：找不到檔案：", out.getvalue())

    def test_unexpected_failure_outside_the_file_loop_restores_the_ui(self):
        window = self.new_window(self.copy(SMALL_08))
        with patch.object(tdms, "process_one_file", side_effect=None), \
                patch.object(tdms.ProcessWorker, "_run", side_effect=RuntimeError("boom")):
            box = []
            window.processing_finished.connect(lambda s: box.append(("done", s)))
            window.btn_run.click()
            self.assertTrue(pump_until(lambda: bool(box), 20000))
        self.assertEqual(box[0], ("done", None))
        self.assertEqual(len(self.dialogs), 1)
        self.assertIn("boom", self.dialogs[0][2])
        self.assertTrue(window.btn_run.isEnabled())
        self.assertTrue(all(g.isEnabled() for g in window._input_groups))
        self.assertFalse(window.is_processing())


class ThreadSafetyTests(BackgroundTestCase):
    def test_worker_never_touches_widgets_dialogs_or_pyplot(self):
        source = self.copy(SMALL_08)
        window = self.new_window(source)
        window.show_t_tsd.setChecked(True)
        checked_threads = []
        real_is_checked = QCheckBox.isChecked

        def spy(self_):
            checked_threads.append(QThread.currentThread() is self.main_thread)
            return real_is_checked(self_)

        pyplot_calls = Mock(side_effect=AssertionError("pyplot must not be used"))
        with patch.object(QCheckBox, "isChecked", spy), patch.object(tdms.plt, "subplots", pyplot_calls), \
                patch.object(tdms.plt, "figure", pyplot_calls), patch.object(tdms.plt, "savefig", pyplot_calls), \
                patch.object(tdms.plt, "show", pyplot_calls):
            summary = self.click_run_and_finish(window)
        self.assertEqual(summary.succeeded, 1)
        self.assertTrue(checked_threads and all(checked_threads), "a checkbox was read from a non-GUI thread")
        self.assertEqual(pyplot_calls.call_count, 0)
        self.assertEqual(plt.get_fignums(), [])
        self.assertTrue(all(d[3] for d in self.dialogs))

    def test_worker_runs_in_another_thread_and_gets_only_a_frozen_snapshot(self):
        source = self.copy(SMALL_08)
        window = self.new_window(source)
        window.show_t_evap.setChecked(False)
        window.annotate_gap_tsd.setChecked(True)
        settings = window._capture_settings()
        self.assertTrue(dataclasses.is_dataclass(settings))
        with self.assertRaises(dataclasses.FrozenInstanceError):
            settings.enable_excel = False
        for field in dataclasses.fields(settings):
            value = getattr(settings, field.name)
            self.assertIsInstance(value, (bool, str, tuple), field.name)   # no widget / QObject inside
        self.assertEqual(settings.wanted_channels, tuple(window.base_channels))
        self.assertFalse(settings.show_t_evap)
        self.assertTrue(settings.annotate_gap_tsd)
        threads = []
        real = tdms.process_one_file

        def spy(path, snapshot, *args, **kwargs):
            threads.append((QThread.currentThread() is self.main_thread, type(snapshot).__name__))
            return real(path, snapshot, *args, **kwargs)

        with patch.object(tdms, "process_one_file", spy):
            self.click_run_and_finish(window)
        self.assertEqual(threads, [(False, "RunSettings")])             # worker thread, receives the snapshot only

    def test_snapshot_is_taken_when_run_is_pressed_not_later(self):
        source = self.copy(SMALL_1459)
        window = self.new_window(source)
        gate = Gate(tdms._read_frame, hold_index=0)
        with patch.object(tdms, "_read_frame", gate):
            with contextlib.redirect_stdout(io.StringIO()):
                window.btn_run.click()
            self.assertTrue(pump_until(gate.entered.is_set, 10000))
            window.enable_excel.setChecked(False)                          # (widget is disabled, but force the value)
            window.show_t_case.setChecked(False)
            box = []
            window.processing_finished.connect(lambda s: box.append(s))
            gate.release.set()
            self.assertTrue(pump_until(lambda: bool(box), 20000))
        self.assertEqual(len(window.generated_excel_paths), 1)            # the Excel setting from the click was used

    def test_no_figure_leak_over_20_renders(self):
        source = self.copy(SMALL_08)
        window = self.new_window(source)
        df, _ = load_frame(window, source, all_channels=False)
        settings = window._capture_settings()

        def live_figures():
            gc.collect()
            return sum(1 for obj in gc.get_objects() if isinstance(obj, Figure))

        render = lambda: tdms.render_plot(df, source.name, str(self.tmp), settings)
        render()                                                            # warm-up (font cache etc.)
        before = live_figures()
        for _ in range(20):
            render()
        after = live_figures()
        self.assertLessEqual(after, before, f"live Figure objects grew from {before} to {after}")
        self.assertEqual(plt.get_fignums(), [])


class LifecycleTests(BackgroundTestCase):
    def test_closing_the_window_while_processing_cancels_and_does_not_crash(self):
        files = [self.copy(SMALL_08, f"f{i}.tdms") for i in range(3)]
        window = self.new_window(*files)
        gate = Gate(tdms._read_frame, hold_index=0)
        messages = []
        qInstallMessageHandler(lambda mode, context, message: messages.append(message))
        try:
            with patch.object(tdms, "_read_frame", gate):
                with contextlib.redirect_stdout(io.StringIO()):
                    window.btn_run.click()
                self.assertTrue(pump_until(gate.entered.is_set, 10000))
                threading.Timer(0.3, gate.release.set).start()
                started = time.monotonic()
                self.assertTrue(window.close())
                elapsed = time.monotonic() - started
                pump_until(lambda: False, 200)
        finally:
            qInstallMessageHandler(None)
        self.assertLess(elapsed, 10)
        self.assertIsNone(window._thread)
        self.assertFalse(window.is_processing() and window._thread is not None)
        self.assertEqual([m for m in messages if "Destroyed while thread is still running" in m], [])
        self.assertEqual(self.dialogs, [])                               # no dialog while closing
        self.assertEqual(self.open_files_mock.call_count, 0)
        self.assertEqual(len(list(self.tmp.glob("*.png"))), 0)           # cancelled before any output was written


class ResponsivenessTests(BackgroundTestCase):
    def heartbeat(self):
        beats = []
        state = {"last": time.perf_counter()}

        def beat():
            now = time.perf_counter()
            beats.append(now - state["last"])
            state["last"] = now

        timer = QTimer()
        timer.setInterval(10)
        timer.timeout.connect(beat)
        timer.start()
        pump_until(lambda: False, 150)           # settle
        beats.clear()
        state["last"] = time.perf_counter()
        return timer, beats

    def test_large_all_channels_gui_thread_stays_responsive(self):
        """Evidence behind HEARTBEAT_THRESHOLD_S (0.25 s): measured on the large capture with all 41 columns,
        the GUI thread's worst 10 ms-heartbeat gap was ~13-40 ms while the worker ran (GIL contention only), whereas
        running the same work synchronously on the GUI thread blocks it for ~1.5-4 s (see evidence/ of the task)."""
        large = find_large_sample()
        if large is None:
            self.skipTest(LARGE_SKIP_MESSAGE)
        source = self.copy(large, subdir="async")
        window = self.new_window(source, all_channels=True)
        timer, beats = self.heartbeat()
        box = []
        window.processing_finished.connect(lambda s: box.append(s))
        with contextlib.redirect_stdout(io.StringIO()):
            window.process_files()
            self.assertTrue(pump_until(lambda: bool(box), 120000))
        timer.stop()
        async_gap = max(beats)
        self.assertEqual(box[0].succeeded, 1)

        # control: the same work executed synchronously on the GUI thread
        source2 = self.copy(large, subdir="sync")
        window2 = self.new_window(source2, all_channels=True)
        settings = window2._capture_settings()
        timer, beats = self.heartbeat()
        with contextlib.redirect_stdout(io.StringIO()):
            result = tdms.process_one_file(str(source2), settings)
        pump_until(lambda: len(beats) > 0, 2000)
        timer.stop()
        sync_gap = max(beats)
        self.assertIsNone(result.error)
        metrics = f"large all-channels: async max heartbeat gap {async_gap * 1000:.1f} ms; synchronous (control) {sync_gap * 1000:.1f} ms"
        print(metrics)
        path = os.environ.get("TDMS_TEST_METRICS_FILE")
        if path:
            with open(path, "a", encoding="utf-8") as fh:
                fh.write(metrics + "\n")
        self.assertLess(async_gap, HEARTBEAT_THRESHOLD_S, metrics)
        self.assertGreater(sync_gap, 4 * async_gap, metrics)


if __name__ == "__main__":
    unittest.main()
