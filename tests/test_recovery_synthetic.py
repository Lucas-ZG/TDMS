"""Synthetic, self-contained coverage of the final-segment recovery path (no external capture needed).

A tiny interleaved TDMS file (toc 0x2E, one float32 value per channel per row) is built in memory. A copy whose
final segment declares fewer rows than the file really contains mimics an interrupted write; the tool must
read it back as "recovered" with every row, and must refuse to "recover" anything it cannot prove.
"""
import os
import struct
import sys
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("MPLBACKEND", "Agg")
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import numpy as np

import tdms

CHANNELS = ("TIME", "A", "B")


def _tdms_string(text):
    raw = text.encode("utf-8")
    return struct.pack("<I", len(raw)) + raw


def build_interleaved_tdms(rows, declared_rows=None):
    """Return (file_bytes, expected_matrix). declared_rows < rows => final segment length is understated."""
    meta = struct.pack("<I", len(CHANNELS))
    for name in CHANNELS:
        meta += _tdms_string("/'G'/'%s'" % name)
        meta += struct.pack("<I", 20) + struct.pack("<II", 9, 1) + struct.pack("<Q", 1) + struct.pack("<I", 0)
    matrix = np.arange(rows * len(CHANNELS), dtype="<f4").reshape(rows, len(CHANNELS)) + 0.5
    declared = rows if declared_rows is None else declared_rows
    next_offset = len(meta) + declared * 4 * len(CHANNELS)
    lead_in = (b"TDSm" + struct.pack("<I", 0x2E) + struct.pack("<I", 4713)
               + struct.pack("<Q", next_offset) + struct.pack("<Q", len(meta)))
    return lead_in + meta + matrix.tobytes(), matrix


class SyntheticRecoveryTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.dir = Path(self._tmp.name)

    def write(self, name, data):
        path = self.dir / name
        path.write_bytes(data)
        return str(path)

    def test_complete_file_is_read_normally_without_recovery(self):
        data, matrix = build_interleaved_tdms(10)
        path = self.write("complete.tdms", data)
        self.assertIsNone(tdms._recover_final_segment(path))  # a complete chain needs no repair
        parsed, status = tdms.read_tdms_with_status(path)
        self.assertEqual(status, "normal")
        np.testing.assert_array_equal(parsed["G"]["B"][:], matrix[:, 2])

    def test_understated_final_segment_is_recovered_with_all_rows(self):
        data, matrix = build_interleaved_tdms(10, declared_rows=6)
        path = self.write("understated.tdms", data)
        with self.assertRaisesRegex(ValueError, "Segment does not start"):
            tdms.TdmsFile.read(path)  # nptdms alone cannot read it
        parsed, status = tdms.read_tdms_with_status(path)
        self.assertEqual(status, "recovered")
        for index, name in enumerate(CHANNELS):
            channel = parsed["G"][name]
            self.assertEqual(len(channel), 10, name)
            np.testing.assert_array_equal(channel[:], matrix[:, index])
        self.assertEqual(Path(path).read_bytes(), data)  # recovery happens in memory only

    def test_misaligned_tail_is_rejected(self):
        data, _ = build_interleaved_tdms(10, declared_rows=6)
        path = self.write("misaligned.tdms", data + b"\x00")
        self.assertIsNone(tdms._recover_final_segment(path))
        with self.assertRaisesRegex(ValueError, "Segment does not start"):
            tdms.read_tdms_with_status(path)

    def test_non_interleaved_toc_is_not_recovered(self):
        data, _ = build_interleaved_tdms(10, declared_rows=6)
        patched = bytearray(data)
        struct.pack_into("<I", patched, 4, 0x0E)  # contiguous layout: recovery is only defined for interleaved
        path = self.write("contiguous.tdms", bytes(patched))
        self.assertIsNone(tdms._recover_final_segment(path))
        with self.assertRaisesRegex(ValueError, "Segment does not start"):
            tdms.read_tdms_with_status(path)


if __name__ == "__main__":
    unittest.main()
