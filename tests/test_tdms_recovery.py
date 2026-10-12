"""Read-only regression checks using the two supplied TDMS captures."""
import io
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import numpy as np
from nptdms import TdmsFile

import tdms


DOWNLOADS = Path(r"C:\Users\15ZD90R\Downloads")
GOOD = DOWNLOADS / "J3100002101469 - NI-sbRIO-9651-021c4630 - 26 09 13 16 00.tdms"
BAD = DOWNLOADS / "J3100002101469 - NI-sbRIO-9651-021c4630 - 26 09 13 13 16.tdms"
DECLARED_END = 500273


@unittest.skipUnless(GOOD.is_file() and BAD.is_file(), "Real-capture recovery tests skipped: the two original TDMS captures (J3100002101469 ... 26 09 13 16 00.tdms and ... 26 09 13 13 16.tdms, not shipped in the repository) are not in the Downloads folder configured above. Edit DOWNLOADS/GOOD/BAD to run them; tests/test_recovery_synthetic.py covers the recovery path with a synthetic file.")
class TdmsRecoveryTests(unittest.TestCase):
    def test_normal_file_uses_original_reader(self):
        with patch.object(tdms, "_recover_final_segment", side_effect=AssertionError("unexpected recovery")):
            parsed, status = tdms.read_tdms_with_status(str(GOOD))
        self.assertEqual(status, "normal")
        self.assertEqual(len(parsed["COBRA THERMAL DATA"]["TIME"]), 29900)

    def test_final_segment_recovers_all_rows_and_preserves_prefix(self):
        original = BAD.read_bytes()
        parsed, status = tdms.read_tdms_with_status(str(BAD))
        self.assertEqual(status, "recovered")
        cobra = parsed["COBRA THERMAL DATA"]
        self.assertEqual(len(cobra.channels()), 41)
        self.assertTrue(all(len(channel) == 3487 for channel in cobra.channels()))
        prefix = TdmsFile.read(io.BytesIO(original[:DECLARED_END]))
        for channel in prefix["COBRA THERMAL DATA"].channels():
            old = channel[:]
            new = cobra[channel.name][:len(old)]
            self.assertEqual(len(old), 2961)
            np.testing.assert_array_equal(old, new)
        self.assertEqual(BAD.read_bytes(), original)

    def test_misaligned_tail_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            candidate = Path(directory) / "misaligned.tdms"
            candidate.write_bytes(BAD.read_bytes() + b"\x00")
            self.assertIsNone(tdms._recover_final_segment(str(candidate)))
            with self.assertRaisesRegex(ValueError, "Segment does not start"):
                tdms.read_tdms_with_status(str(candidate))

    def test_processing_continues_after_a_failed_file(self):
        plotted = Mock()
        fake = SimpleNamespace(
            file_paths=["bad.tdms", "good.tdms", "recovered.tdms"],
            generated_excel_paths=[1],
            generated_image_paths=[1],
            enable_excel=SimpleNamespace(isChecked=lambda: False),
            auto_open_plot=SimpleNamespace(isChecked=lambda: False),
            plot_and_save=plotted,
        )
        def read_file(path):
            if path == "bad.tdms":
                raise ValueError("invalid segment")
            return object(), "recovered" if path == "recovered.tdms" else "normal"
        fake.read_tdms_file = read_file
        with patch.object(tdms.QMessageBox, "critical") as critical, patch.object(
            tdms.QMessageBox, "information") as information:
            tdms.TDMSFilterApp.process_files(fake)
        self.assertEqual(plotted.call_count, 2)
        self.assertEqual(critical.call_count, 1)
        self.assertIn("bad.tdms", critical.call_args.args[2])
        self.assertEqual(information.call_count, 1)
        self.assertIn("recovered.tdms", information.call_args.args[2])

    def test_channel_scan_warns_about_skipped_file(self):
        fake = SimpleNamespace(
            file_paths=["bad.tdms", "good.tdms"],
            base_channels=[],
            _iter_cobra_groups=lambda td: [],
            _all_groups=lambda td: [],
            _rebuild_channel_checklist=Mock(),
        )
        def read_file(path):
            if path == "bad.tdms":
                raise ValueError("invalid segment")
            return object(), "normal"
        with patch.object(tdms, "read_tdms_with_status", side_effect=read_file), patch.object(
            tdms.QMessageBox, "warning") as warning:
            tdms.TDMSFilterApp._rebuild_checklist_from_all_files(fake)
        self.assertEqual(warning.call_count, 1)
        self.assertIn("bad.tdms", warning.call_args.args[2])
        fake._rebuild_channel_checklist.assert_called_once_with([])

    def test_extra_segment_marker_in_tail_is_rejected(self):
        data = bytearray(BAD.read_bytes())
        data[DECLARED_END + 164:DECLARED_END + 168] = b"TDSm"
        with tempfile.TemporaryDirectory() as directory:
            candidate = Path(directory) / "extra_header.tdms"
            candidate.write_bytes(data)
            self.assertIsNone(tdms._recover_final_segment(str(candidate)))
            with self.assertRaisesRegex(ValueError, "Segment does not start"):
                tdms.read_tdms_with_status(str(candidate))


if __name__ == "__main__":
    unittest.main()
