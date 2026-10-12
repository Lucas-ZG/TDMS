# Author: Lucas.Zhuge
# VERSION NOTE:
# 3.8:
#   - 穩定版本整合：全面修正勾選欄位不同步問題
#   - 新增「即時讀取 UI 勾選狀態」機制，確保使用者勾選的所有 channel 皆正確輸出
#   - 勾選的 channel 會完整複製整列資料進 Excel（若找不到則自動建立空欄）
#   - 強化 channel 名稱比對邏輯（群組內精確→寬鬆→跨群組搜尋）
#   - 介面維持一致，但內部邏輯更簡潔與穩定
#   - 圖表仍僅繪製四條主要溫度曲線（M1 T-Case / M1 T-Evap / M1 TSD / M1 Set Temp）
#
# 3.7.3:
#   - 勾選的 channel 會自動加入 Excel（找不到則產生空欄）
#   - 強化多群組匹配機制，支援跨群組尋找 channel
#   - 額外欄位改為多欄勾選清單，增加全選/全不選功能
#
# 3.7.2:
#   - 新增同時支援「選擇檔案」與「選擇資料夾」兩種來源
#   - 匯入資料夾時會自動遞迴搜尋所有 .tdms 檔案並合併至清單
#   - 增加額外欄位下拉選單（後續演進為勾選清單）
#
# 3.7.1:
#   - 改進資料夾搜尋邏輯，支援多層資料夾自動掃描
#   - 新增版本號與更新紀錄系統
#
# 3.7:
#   - 新增「選擇資料夾」功能（可遞迴掃描所有子資料夾內的 TDMS）
#   - 新增「是否匯出 Excel」與「自動開啟 Excel / 圖檔」選項
#   - 新增圖上標註功能：可顯示 TSD 的最高/最低點與 Evap/Case 差值 (Gap)
#
VERSION = "v3.8"

from typing import List, Dict, Set, Optional, Tuple
from dataclasses import dataclass
from pathlib import Path
import io
import struct
import sys
import threading
import os
import re
import platform
import subprocess
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt

# 常數定義
WINDOW_WIDTH = 600
WINDOW_HEIGHT = 760
FONT_SIZE = 10
FONT_NAME = "Microsoft JhengHei"
COLUMNS_PER_ROW = 4
TARGET_GROUP_NAME = "COBRA THERMAL DATA"

# 檔案類型
TDMS_EXTENSION = ".tdms"
EXCEL_EXTENSION = ".xlsx"
PNG_EXTENSION = ".png"
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QCheckBox,
    QFileDialog, QMessageBox, QGroupBox, QScrollArea, QHBoxLayout, QFrame, QGridLayout
)
from PySide6.QtGui import QFont
from PySide6.QtCore import Qt, QTimer

try:
    from nptdms import TdmsFile
except ImportError:
    TdmsFile = None

RECOVERED_SEGMENT_MESSAGE = "Recovered final TDMS segment using EOF"


def _recover_final_segment(path: str):
    """Repair only an understated final, fixed-width interleaved segment in memory."""
    data = Path(path).read_bytes()
    size = len(data)
    position = 0
    positions = []
    file_version = None
    while position + 28 <= size:
        if data[position:position + 4] != b"TDSm":
            return None
        toc, version, next_offset, raw_offset = struct.unpack_from("<IIQQ", data, position + 4)
        if toc not in (0x0E, 0x2E) or version not in (4712, 4713):
            return None  # Only the simple little-endian metadata/raw-data layout is supported.
        if file_version is not None and version != file_version:
            return None
        file_version = version
        end = position + 28 + next_offset
        raw_start = position + 28 + raw_offset
        if not (position + 28 <= raw_start <= end <= size) or end <= position:
            return None
        positions.append(position)
        if end == size:
            return None  # A complete chain needs no repair.
        if data[end:end + 4] == b"TDSm":
            position = end
            continue

        # A missing header is recoverable only after the last declared segment.
        if data.find(b"TDSm", end) != -1 or toc != 0x2E:
            return None
        try:
            prefix = TdmsFile.read(io.BytesIO(data[:end]))
            segments = prefix._reader._segments
            if [segment.position for segment in segments] != positions:
                return None
            segment = segments[-1]
            if (segment.data_position != raw_start or segment.next_segment_pos != end
                    or segment.segment_incomplete):
                return None
            objects = segment.ordered_objects
            if not objects or any(
                type(obj).__name__ != "TdmsSegmentObject"
                or not obj.has_data
                or obj.number_values != 1
                or obj.data_type.size is None
                or obj.data_size != obj.data_type.size
                for obj in objects
            ):
                return None
            row_size = sum(obj.data_size for obj in objects)
            if (row_size <= 0 or segment.final_chunk_lengths_override is not None
                    or (end - raw_start) % row_size != 0
                    or (size - end) % row_size != 0):
                return None
            repaired = bytearray(data)
            struct.pack_into("<Q", repaired, position + 12, size - (position + 28))
            return TdmsFile.read(io.BytesIO(repaired))
        except Exception:
            return None
    return None


def read_tdms_with_status(path: str):
    """Return (TdmsFile, normal|recovered); let callers report failures."""
    try:
        return TdmsFile.read(path), "normal"
    except ValueError as exc:
        if not str(exc).startswith("Segment does not start with b'TDSm'"):
            raise
        recovered = _recover_final_segment(path)
        if recovered is None:
            raise
        return recovered, "recovered"

# ---- 名稱正規化 ----
def _norm_group(s: str) -> str:
    return "".join(c.lower() for c in s if c.isalnum())

_space_re = re.compile(r"\s+")
_non_alnum_re = re.compile(r"[^0-9a-z]")

def _norm_space(s: str) -> str:
    s = s.strip()
    s = _space_re.sub(" ", s)
    return s.lower()

def _norm_alnum(s: str) -> str:
    return _non_alnum_re.sub("", s.lower())

TARGET_GROUP_NAME = "COBRA THERMAL DATA"

# ---- Phase 2-B：背景處理（設定快照／處理核心／worker）-------------------------------------------
# 以下新增的 import 與其使用處相鄰（既有頂層 import 區塊保持原樣，T04 另案處理）。
import time as _time
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtWidgets import QProgressBar

# 等待工作執行緒結束的最長時間（毫秒）；模組層級常數，測試可暫時縮短。
THREAD_WAIT_TIMEOUT_MS = 10000

CURVE_COLUMNS = ("M1 T-Case", "M1 T-Evap", "M1 TSD", "M1 Set Temp")


@dataclass(frozen=True)
class RunSettings:
    """按下「執行」時在主執行緒一次擷取的不可變設定；工作執行緒只使用這個物件，不讀任何 widget。"""
    wanted_channels: Tuple[str, ...]      # base + 額外勾選（順序規則與舊版 read_tdms_file 相同）
    show_t_case: bool
    show_t_evap: bool
    show_t_tsd: bool
    show_t_set: bool
    annotate_min_max: bool
    annotate_gap_tsd: bool
    annotate_gap_case: bool
    enable_excel: bool
    auto_open_excel: bool
    auto_open_plot: bool

    def curve_flags(self):
        return {
            "M1 T-Case": self.show_t_case,
            "M1 T-Evap": self.show_t_evap,
            "M1 TSD": self.show_t_tsd,
            "M1 Set Temp": self.show_t_set,
        }

    def steps_per_file(self) -> int:
        return 3 if self.enable_excel else 2   # 讀檔 + (Excel) + 繪圖


@dataclass(frozen=True)
class FileResult:
    """單一檔案處理結果（不可變，可安全跨執行緒傳遞）。"""
    path: str
    status: Optional[str] = None            # "normal" / "recovered" / None（讀檔失敗）
    excel_path: Optional[str] = None
    image_path: Optional[str] = None
    error: Optional[str] = None             # f"{檔名}：{例外}"，與舊版訊息文字相同
    completed: bool = True                  # False = 在階段之間被取消
    timings: Tuple[Tuple[str, float], ...] = ()

    @property
    def recovered_message(self) -> Optional[str]:
        if self.status == "recovered":
            return f"{os.path.basename(self.path)}：{RECOVERED_SEGMENT_MESSAGE}"
        return None


@dataclass(frozen=True)
class RunSummary:
    results: Tuple[FileResult, ...]
    total_files: int
    cancelled: bool
    elapsed: float
    messages: Tuple[Tuple[str, str], ...]   # ("recovered"|"error", 訊息文字)，依處理順序
    excel_paths: Tuple[str, ...]
    image_paths: Tuple[str, ...]

    @property
    def processed(self) -> int:             # 已完整處理的檔案（含失敗略過者）
        return sum(1 for r in self.results if r.completed)

    @property
    def skipped(self) -> int:
        return sum(1 for r in self.results if r.completed and r.error)

    @property
    def succeeded(self) -> int:
        return sum(1 for r in self.results if r.completed and not r.error)

    @property
    def recovered(self) -> int:
        return sum(1 for r in self.results if r.status == "recovered")


# 以下四個函式與舊版 TDMSFilterApp 的同名方法內容相同（方法改為轉呼叫），讓處理核心不需要 self。
def _find_group_cobra(td):
    target = _norm_group(TARGET_GROUP_NAME)
    try:
        for g in td.groups():
            if target in _norm_group(getattr(g, "name", str(g))): return g
    except Exception: pass
    try: return td[TARGET_GROUP_NAME]
    except Exception: return None


def _all_groups(td):
    try:
        return list(td.groups())
    except Exception:
        return []


def _resolve_in_group(grp, wanted: str):
    # 精確
    try:
        _ = grp[wanted]; return wanted
    except KeyError:
        pass
    # 寬鬆
    try:
        names = [ch.name for ch in grp.channels()]
    except Exception:
        names = []
    map_space = {_norm_space(n): n for n in names}
    map_alnum = {_norm_alnum(n): n for n in names}
    ws, wa = _norm_space(wanted), _norm_alnum(wanted)
    if ws in map_space:   return map_space[ws]
    if wa in map_alnum:   return map_alnum[wa]
    for n in names:
        if wa and (wa in _norm_alnum(n) or _norm_alnum(n) in wa): return n
    return None


def _resolve_anywhere(td, wanted: str):
    for grp in _all_groups(td):
        real = _resolve_in_group(grp, wanted)
        if real is not None: return grp, real
    return None, None


def _read_frame(path: str, wanted: List[str]) -> Tuple[pd.DataFrame, str]:
    """讀取 TDMS 並轉成 DataFrame（舊版 read_tdms_file 的核心，不讀任何 UI 狀態）。"""
    if TdmsFile is None:
        raise ImportError("尚未安裝 nptdms 套件")

    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到檔案：{path}")

    try:
        td, read_status = read_tdms_with_status(path)
    except Exception as e:
        raise ValueError(f"無法讀取 TDMS 檔案：{e}") from e

    cobra = _find_group_cobra(td)

    # 優先取 TIME 以確定列數
    time_data = None
    if cobra is not None:
        r = _resolve_in_group(cobra, "TIME")
        if r is None:
            g2, r2 = _resolve_anywhere(td, "TIME")
            if r2 is not None:
                cobra = g2; r = r2
        if r is not None:
            try: time_data = cobra[r][:]
            except Exception: time_data = None

    col_data = {}
    found, missed = [], []

    if time_data is not None:
        col_data["TIME"] = time_data

    for name in wanted:
        if name == "TIME":
            continue
        grp = cobra
        real = None
        if grp is not None:
            real = _resolve_in_group(grp, name)
        if real is None:
            g2, r2 = _resolve_anywhere(td, name)
            if r2 is not None:
                grp = g2; real = r2
        if real is None:
            missed.append(name); col_data[name] = None; continue
        try:
            col_data[name] = grp[real][:]
            found.append(name)
        except Exception:
            col_data[name] = None; missed.append(name)

    lengths = [len(v) for v in col_data.values() if v is not None]
    n = max(lengths) if lengths else 1

    df_parts = []
    for name in wanted:
        arr = col_data.get(name, None)
        if arr is None:
            df_parts.append(pd.DataFrame({name: [None] * n}))
        else:
            s = pd.Series(arr)
            if len(s) < n: s = s.reindex(range(n))
            df_parts.append(pd.DataFrame({name: s}))
    df = pd.concat(df_parts, axis=1)

    if missed:
        print(f"[{os.path.basename(path)}] 找到: {sorted(found)}；找不到: {sorted(missed)}")
    else:
        print(f"[{os.path.basename(path)}] 找到: {sorted(found)}")
    return df, read_status


def render_plot(df, filename, folder, settings) -> Optional[str]:
    """以 Figure + FigureCanvasAgg 繪圖並存檔（不使用 pyplot 全域狀態，可在背景執行緒執行）。
    繪圖內容與舊版 plot_and_save 相同；回傳 PNG 路徑，沒有 TIME 欄時回傳 None。"""
    # 只畫四條
    if "TIME" not in df.columns: return None
    fig = Figure(figsize=(10, 6))
    FigureCanvasAgg(fig)
    ax = fig.subplots()
    t = df["TIME"]

    targets = settings.curve_flags()
    for col, show in targets.items():
        if show and col in df.columns:
            ax.plot(t, df[col], label=col)
            if settings.annotate_min_max and col == "M1 TSD":
                y = df[col].dropna()
                if not y.empty:
                    imax, imin = y.idxmax(), y.idxmin()
                    ax.annotate(f"Max: {y[imax]:.1f}", xy=(t[imax], y[imax]),
                                xytext=(5,5), textcoords='offset points', fontsize=9, color='red')
                    ax.annotate(f"Min: {y[imin]:.1f}", xy=(t[imin], y[imin]),
                                xytext=(5,-15), textcoords='offset points', fontsize=9, color='blue')

    if settings.annotate_gap_tsd and "M1 T-Evap" in df.columns and "M1 TSD" in df.columns:
        gap = (df["M1 T-Evap"] - df["M1 TSD"]).abs().dropna()
        if not gap.empty:
            idx = gap.idxmax()
            ax.annotate(f"Max Evap-TSD Gap: {gap[idx]:.1f} @ Time: {t[idx]:.1f}",
                        xy=(t[idx], df["M1 T-Evap"][idx]), xytext=(10,10),
                        textcoords='offset points', fontsize=9, color='purple',
                        arrowprops=dict(arrowstyle='->', color='purple'))
    if settings.annotate_gap_case and "M1 T-Evap" in df.columns and "M1 T-Case" in df.columns:
        gap = (df["M1 T-Evap"] - df["M1 T-Case"]).abs().dropna()
        if not gap.empty:
            idx = gap.idxmax()
            ax.annotate(f"Max Evap-Case Gap: {gap[idx]:.1f} @ Time: {t[idx]:.1f}",
                        xy=(t[idx], df["M1 T-Evap"][idx]), xytext=(10,-25),
                        textcoords='offset points', fontsize=9, color='green',
                        arrowprops=dict(arrowstyle='->', color='green'))

    ax.set_xlabel("Time"); ax.set_ylabel("Value")
    ax.set_title(f"TDMS Thermal Data - {filename}")
    ax.legend(); ax.grid(True); fig.tight_layout()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    img_path = os.path.join(folder, f"{os.path.splitext(filename)[0]}_{ts}.png")
    fig.savefig(img_path)
    return img_path


class _ExcelSink:
    """讓 TDMSFilterApp.export_to_excel（T02，本體不改）能在沒有 widget 的情況下被處理核心呼叫。"""
    def __init__(self):
        self.generated_excel_paths = []


def process_one_file(path: str, settings: RunSettings, before_stage=None, after_stage=None) -> FileResult:
    """單一檔案的完整處理（讀檔 → Excel → 繪圖）。純 Python、不碰 Qt。

    before_stage(stage) -> bool：每個階段開始前呼叫，回傳 False 表示要取消（階段之間才檢查，不中斷進行中的階段）。
    after_stage(stage)：每個階段完成後呼叫。stage 為 "read" / "excel" / "plot"。
    例外處理與舊版 process_files 相同：任何階段失敗 → error = f"{檔名}：{例外}"，已完成階段的產出保留。
    """
    name = os.path.basename(path)
    timings = []
    status = excel_path = image_path = error = None
    completed = True

    def gate(stage):
        return before_stage is None or before_stage(stage) is not False

    def done(stage, started):
        timings.append((stage, _time.perf_counter() - started))
        if after_stage is not None:
            after_stage(stage)

    try:
        if not gate("read"):
            return FileResult(path, completed=False, timings=tuple(timings))
        started = _time.perf_counter()
        df, status = _read_frame(path, list(settings.wanted_channels))
        done("read", started)
        if settings.enable_excel:
            if not gate("excel"):
                return FileResult(path, status, completed=False, timings=tuple(timings))
            started = _time.perf_counter()
            sink = _ExcelSink()
            TDMSFilterApp.export_to_excel(sink, df, path)
            excel_path = sink.generated_excel_paths[-1]
            done("excel", started)
        if not gate("plot"):
            return FileResult(path, status, excel_path, completed=False, timings=tuple(timings))
        started = _time.perf_counter()
        image_path = render_plot(df, name, os.path.dirname(path), settings)
        done("plot", started)
    except Exception as e:
        error = f"{name}：{e}"
    return FileResult(path, status, excel_path, image_path, error, completed, tuple(timings))


class ProcessWorker(QObject):
    """在 QThread 內執行整批處理。只使用傳入的 RunSettings 與路徑，不存取任何 widget／對話框。"""
    progress = Signal(int, int, str)        # 已完成步數, 總步數, 說明文字
    file_done = Signal(object)              # FileResult
    finished_all = Signal(object)           # RunSummary
    failed = Signal(str)                    # 迴圈外的非預期例外

    STAGE_LABELS = {"read": "讀取 TDMS", "excel": "匯出 Excel", "plot": "繪圖"}

    def __init__(self, paths, settings: RunSettings):
        super().__init__()
        self._paths = tuple(paths)
        self._settings = settings
        # 取消旗標用 threading.Event：set() 與 is_set() 本身是執行緒安全的（內含鎖），
        # GUI 執行緒呼叫 request_cancel()、工作執行緒在階段之間以 is_set() 檢查，不依賴 GIL 對普通屬性的寫入細節。
        self._cancel_event = threading.Event()

    def request_cancel(self):
        """可由任何執行緒呼叫：設定 threading.Event（工作執行緒在階段與檔案之間以 is_set() 檢查）。"""
        self._cancel_event.set()

    def cancel_requested(self) -> bool:
        return self._cancel_event.is_set()

    @Slot()
    def delete_when_thread_finishes(self):
        """連到 QThread.finished：在工作執行緒結束前，於該執行緒內排程刪除 worker 自己。"""
        self.deleteLater()

    @Slot()
    def run(self):
        try:
            self._run()
        except Exception as e:  # 迴圈外的非預期例外：回報並讓 UI 恢復
            self.failed.emit(f"{type(e).__name__}: {e}")

    def _run(self):
        settings = self._settings
        total = len(self._paths) * settings.steps_per_file()
        started = _time.perf_counter()
        done_steps = 0
        results, messages, excel_paths, image_paths = [], [], [], []
        cancelled = False

        for path in self._paths:
            if self._cancel_event.is_set():
                cancelled = True
                break
            name = os.path.basename(path)

            def before(stage, name=name):
                if self._cancel_event.is_set():
                    return False
                self.progress.emit(done_steps, total, f"{name}：{self.STAGE_LABELS[stage]}")
                return True

            def after(stage):
                nonlocal done_steps
                done_steps += 1

            step_mark = done_steps
            result = process_one_file(path, settings, before, after)
            results.append(result)
            if result.recovered_message:
                print(result.recovered_message)
                messages.append(("recovered", result.recovered_message))
            if result.error:
                print(f"[TDMS 略過] {result.error}")
                messages.append(("error", result.error))
            if result.excel_path: excel_paths.append(result.excel_path)
            if result.image_path: image_paths.append(result.image_path)
            if not result.completed:
                cancelled = True
            elif result.error:   # 失敗的檔案：把略過的步數補上，進度條才能走到底
                done_steps = step_mark + settings.steps_per_file()
            self.file_done.emit(result)
            if cancelled:
                break

        self.progress.emit(done_steps, total, "已取消" if cancelled else "完成")
        self.finished_all.emit(RunSummary(
            tuple(results), len(self._paths), cancelled, _time.perf_counter() - started,
            tuple(messages), tuple(excel_paths), tuple(image_paths)))


class TDMSFilterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"TDMS小工具 {VERSION}")
        self.resize(600, 830)   # Phase 2-B: 進度條／狀態文字／取消按鈕使最小高度從 737 增加到 812

        # 狀態
        self.file_paths = []
        self.generated_excel_paths, self.generated_image_paths = [], []
        # 背景處理狀態（Phase 2-B）
        self._thread = None
        self._worker = None
        self._processing = False
        self._closing = False
        self._run_settings = None
        self._outcome = None            # 等待收尾的完成／失敗結果（見 _try_complete）

        # 固定輸出的五欄（TIME+四條）
        self.base_channels = ["TIME", "M1 T-Case", "M1 T-Evap", "M1 TSD", "M1 Set Temp"]

        # UI 勾選元件（不再用內部 set 記錄）
        self.channel_checkboxes = {}      # name -> QCheckBox
        self.columns_per_row = 4          # 勾選清單每列幾個

        self.setFont(QFont("Microsoft JhengHei", 10))
        self._init_ui()

    # ---------------- UI ----------------
    def _init_ui(self):
        main = QVBoxLayout()

        # 檔案來源
        g_file = QGroupBox("檔案來源")
        v = QVBoxLayout()
        b1 = QPushButton("選擇 TDMS 檔案（可多選）"); b1.clicked.connect(self.select_tdms_files)
        b2 = QPushButton("選擇資料夾（遞迴掃描 .tdms）"); b2.clicked.connect(self.select_folder)
        self.label_file = QLabel("尚未選擇檔案/資料夾")
        v.addWidget(b1); v.addWidget(b2); v.addWidget(self.label_file)
        g_file.setLayout(v)

        # 輸出設定
        g_out = QGroupBox("輸出設定")
        v2 = QVBoxLayout()
        self.enable_excel = QCheckBox("是否匯出 Excel"); self.enable_excel.setChecked(True)
        self.auto_open_excel = QCheckBox("匯出後自動開啟 Excel"); self.auto_open_excel.setChecked(True)
        self.auto_open_plot  = QCheckBox("儲存後自動開啟圖檔"); self.auto_open_plot.setChecked(False)
        v2.addWidget(self.enable_excel); v2.addWidget(self.auto_open_excel); v2.addWidget(self.auto_open_plot)
        g_out.setLayout(v2)

        # 顯示曲線（固定四條）
        g_curve = QGroupBox("顯示曲線")
        v3 = QVBoxLayout()
        self.show_t_case = QCheckBox("M1 T-Case")
        self.show_t_evap = QCheckBox("M1 T-Evap")
        self.show_t_tsd  = QCheckBox("M1 TSD")
        self.show_t_set  = QCheckBox("M1 Set Temp")
        for c in [self.show_t_case, self.show_t_evap, self.show_t_tsd, self.show_t_set]:
            c.setChecked(True); v3.addWidget(c)
        g_curve.setLayout(v3)

        # 標註設定
        g_anno = QGroupBox("標註設定")
        v4 = QVBoxLayout()
        self.annotate_min_max = QCheckBox("標註 TSD 的最高/最低溫")
        self.annotate_gap_tsd = QCheckBox("標註 Evap 與 TSD 間最大差值 (Gap)")
        self.annotate_gap_case = QCheckBox("標註 Evap 與 Case 間最大差值 (Gap)")
        for c in [self.annotate_min_max, self.annotate_gap_tsd, self.annotate_gap_case]:
            v4.addWidget(c)
        g_anno.setLayout(v4)

        # 額外欄位
        g_extra = QGroupBox("額外欄位")
        v5 = QVBoxLayout()

        op = QHBoxLayout()
        btn_all  = QPushButton("全選");  btn_all.setFixedWidth(72)
        btn_none = QPushButton("全不選"); btn_none.setFixedWidth(72)
        btn_all.clicked.connect(self._select_all_channels)
        btn_none.clicked.connect(self._deselect_all_channels)
        op.addWidget(btn_all); op.addWidget(btn_none); op.addStretch()
        v5.addLayout(op)

        # 設定捲動區域
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.StyledPanel)
        self.scroll.setMinimumHeight(120)  # 調整最小高度
        self.scroll.setMaximumHeight(120)  # 限制最大高度
        
        # 設定內部容器
        self.inner = QWidget()
        self.grid = QGridLayout(self.inner)
        self.grid.setSpacing(4)  # 減少項目間距
        self.grid.setContentsMargins(5, 5, 5, 5)  # 減少邊距
        self.grid.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        
        self.scroll.setWidget(self.inner)
        v5.addWidget(self.scroll)

        v5.addWidget(QLabel("※ 額外欄位勾選後輸出至Excel中"))
        g_extra.setLayout(v5)

        # 執行
        btn_run = QPushButton("▶ 執行處理 TDMS 檔案")
        btn_run.setStyleSheet("font-weight: bold; font-size: 12pt; padding: 6px;")
        btn_run.clicked.connect(self.process_files)
        self.btn_run = btn_run

        # 佈局
        main.addWidget(g_file); main.addWidget(g_out); main.addWidget(g_curve); main.addWidget(g_anno)
        main.addWidget(g_extra); main.addWidget(btn_run)
        # 進度條／狀態文字／取消按鈕（處理中才可取消）
        self.progress_bar = QProgressBar(); self.progress_bar.setRange(0, 1); self.progress_bar.setValue(0)
        self.label_status = QLabel("")
        self.btn_cancel = QPushButton("取消"); self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.cancel_processing)
        main.addWidget(self.progress_bar); main.addWidget(self.label_status); main.addWidget(self.btn_cancel)
        self._input_groups = [g_file, g_out, g_curve, g_anno, g_extra]
        main.addStretch()
        self.setLayout(main)

    # ---------------- 檔案選擇 ----------------
    def select_tdms_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "選擇 TDMS 檔案", "", "TDMS files (*.tdms)")
        if not files: self._rebuild_channel_checklist([]); return
        self._add_files(files); self._rebuild_checklist_from_all_files()

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "選擇資料夾")
        if not folder: self._rebuild_channel_checklist([]); return
        tdms_files = []
        for root_dir, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".tdms"): tdms_files.append(os.path.join(root_dir, f))
        if not tdms_files:
            QMessageBox.warning(self, "提示", "資料夾內沒有 TDMS 檔案")
            self._rebuild_channel_checklist([]); return
        self._add_files(tdms_files); self._rebuild_checklist_from_all_files()

    def _add_files(self, paths: List[str]) -> None:
        exist = {Path(p).resolve().as_posix().lower() for p in self.file_paths}
        add = 0
        for p in paths:
            path = Path(p).resolve()
            if path.suffix.lower() == TDMS_EXTENSION and path.as_posix().lower() not in exist:
                self.file_paths.append(str(path))
                add += 1
        self.label_file.setText(f"已選擇 {len(self.file_paths)} 筆資料（本次新增 {add}）")

    def load_initial_tdms_file(self, path: str) -> None:
        """Load one command-line TDMS file through the existing UI flow."""
        self._add_files([path])
        self._rebuild_checklist_from_all_files()

    # ---------------- 勾選清單（聯集 & 群組模糊） ----------------
    def _iter_cobra_groups(self, td: TdmsFile):
        target = _norm_group(TARGET_GROUP_NAME)
        try:
            for grp in td.groups():
                name = getattr(grp, "name", str(grp))
                if target in _norm_group(name): yield grp
        except Exception: return

    def _all_groups(self, td: TdmsFile):
        return _all_groups(td)

    def _rebuild_checklist_from_all_files(self):
        if not self.file_paths: self._rebuild_channel_checklist([]); return
        if TdmsFile is None:
            QMessageBox.critical(self, "錯誤", "尚未安裝 nptdms 套件"); self._rebuild_channel_checklist([]); return

        union_names = set()
        failed_files = []
        recovered_files = []
        for p in self.file_paths:
            try:
                td, status = read_tdms_with_status(p)
                if status == "recovered":
                    recovered_files.append(os.path.basename(p))
            except Exception as exc:
                failed_files.append(f"{os.path.basename(p)}：{exc}")
                continue

            matched = False
            for grp in self._iter_cobra_groups(td):
                matched = True
                try: union_names.update(ch.name for ch in grp.channels())
                except Exception: pass

            if not matched:
                # COBRA 找不到時，把所有群組的通道也列進清單
                for grp in self._all_groups(td):
                    try: union_names.update(ch.name for ch in grp.channels())
                    except Exception: pass

        for name in recovered_files:
            print(f"[{name}] {RECOVERED_SEGMENT_MESSAGE}")
        if recovered_files:
            QMessageBox.information(self, "TDMS 復原", "\n".join(
                f"{name}：{RECOVERED_SEGMENT_MESSAGE}" for name in recovered_files))
        if failed_files:
            for failure in failed_files:
                print(f"[TDMS 略過] {failure}")
            QMessageBox.warning(self, "略過無法讀取的 TDMS", "\n".join(failed_files))

        # 排除 TIME 與四條
        exclude = set(self.base_channels)
        candidates = sorted([n for n in union_names if n not in exclude])
        self._rebuild_channel_checklist(candidates)

    def _rebuild_channel_checklist(self, names):
        self.channel_checkboxes.clear()
        while self.grid.count():
            it = self.grid.takeAt(0)
            w = it.widget()
            if w is not None: w.deleteLater()

        if not names:
            lbl = QLabel("（找不到可選欄位；請確認檔案內是否存在相關 channel）")
            lbl.setStyleSheet("color:#aaa;"); self.grid.addWidget(lbl, 0, 0); return

        colN = max(1, self.columns_per_row)
        r = c = 0
        for name in names:
            cb = QCheckBox(name)
            cb.setStyleSheet("""
                QCheckBox {
                    padding: 2px;
                    margin: 1px;
                }
                QCheckBox:hover {
                    background-color: #f0f0f0;
                    border-radius: 2px;
                }
            """)
            self.grid.addWidget(cb, r, c)
            self.channel_checkboxes[name] = cb
            c += 1
            if c >= colN: c = 0; r += 1

    def _select_all_channels(self):
        for cb in self.channel_checkboxes.values(): cb.setChecked(True)

    def _deselect_all_channels(self):
        for cb in self.channel_checkboxes.values(): cb.setChecked(False)

    # ---- 取得「目前 UI 勾選的額外欄位」 ----
    def _get_checked_extra_channels(self) -> List[str]:
        """獲取目前在 UI 上被勾選的額外通道清單。

        Returns:
            已勾選的通道名稱列表
        """
        return [name for name, cb in self.channel_checkboxes.items() if cb.isChecked()]

    # ---------------- 名稱解析 ----------------
    def _find_group_cobra(self, td: TdmsFile):
        return _find_group_cobra(td)

    def _resolve_in_group(self, grp, wanted: str):
        return _resolve_in_group(grp, wanted)

    def _resolve_anywhere(self, td: TdmsFile, wanted: str):
        return _resolve_anywhere(td, wanted)

    # ---------------- 讀檔 / 匯出 / 畫圖 ----------------
    def read_tdms_file(self, path: str) -> Tuple[pd.DataFrame, str]:
        """讀取 TDMS 檔案並轉換為 DataFrame，回傳 normal/recovered 狀態。

        Args:
            path: TDMS 檔案路徑

        Returns:
            (包含所需通道資料的 DataFrame, 讀取狀態)

        Raises:
            ImportError: 未安裝 nptdms
            FileNotFoundError: 檔案不存在
            ValueError: 檔案格式錯誤
        """
        # ⚠️ 核心：當下讀 UI 勾選；實際讀取由不碰 UI 的處理核心 _read_frame 完成
        extras_now = self._get_checked_extra_channels()
        wanted = list(dict.fromkeys(self.base_channels + sorted(extras_now)))
        return _read_frame(path, wanted)

    def export_to_excel(self, df, tdms_path):
        # 全欄輸出（包含空欄）
        base = os.path.splitext(os.path.basename(tdms_path))[0]
        folder = os.path.dirname(tdms_path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(folder, f"{base}_{ts}.xlsx")
        # 取代 df.to_excel(out_path, index=False)：以 openpyxl write_only 串流寫入，輸出與原做法在儲存格層級
        # 一致（工作表 Sheet1、標題列含 pandas 的粗體／細框線／置中樣式、無索引欄、NaN/None 以空字串
        # 儲存格保留，全空欄仍保留）。
        # 效能：openpyxl 逐格以純 Python 序列化 XML 是主要耗時，因此「全部欄位皆為浮點數」（本工具的
        # 實際資料）時，標題列與活頁簿骨架仍由 openpyxl write_only 產生，資料列則直接輸出 <row>/<c> XML；
        # 任何其他型別組合走一般的 write_only 路徑。兩條路徑與 df.to_excel 的比對結果見測試與 evidence。
        import datetime as _dt
        import io as _io
        import zipfile as _zipfile
        import numpy as np
        from openpyxl import Workbook
        from openpyxl.cell import Cell, WriteOnlyCell
        from openpyxl.styles import Alignment, Border, Font, Side
        from openpyxl.utils import get_column_letter
        from pandas.api.types import is_bool, is_float, is_integer, is_scalar

        inf = float("inf")
        thin = Side(style="thin")

        def convert(val, sheet):
            # 與 pandas ExcelFormatter._format_value + ExcelWriter._value_with_fmt 相同的轉換規則
            kind = type(val)
            if kind is float:  # 最常見的情況（float32 欄經 tolist() 為 Python float）
                if val != val:
                    return ""
                if val == inf:
                    return "inf"
                if val == -inf:
                    return "-inf"
                return val
            if kind is int or kind is str or kind is bool:
                return val
            if is_scalar(val) and pd.isna(val):
                return ""
            if is_float(val):
                if val == inf:
                    return "inf"
                if val == -inf:
                    return "-inf"
            if getattr(val, "tzinfo", None) is not None:
                raise ValueError(
                    "Excel does not support datetimes with "
                    "timezones. Please ensure that datetimes "
                    "are timezone unaware before writing to Excel."
                )
            if is_integer(val):
                return int(val)
            if is_float(val):
                return float(val)
            if is_bool(val):
                return bool(val)
            if isinstance(val, _dt.datetime):
                number_format = "YYYY-MM-DD HH:MM:SS"
            elif isinstance(val, _dt.date):
                number_format = "YYYY-MM-DD"
            elif isinstance(val, _dt.timedelta):
                val = val.total_seconds() / 86400
                number_format = "0"
            else:
                return str(val)
            cell = WriteOnlyCell(sheet, value=val)
            cell.number_format = number_format
            return cell

        def header_cell(label, sheet):
            value = convert(label, sheet)
            cell = value if isinstance(value, Cell) else WriteOnlyCell(sheet, value=value)
            cell.font = Font(bold=True)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(horizontal="center", vertical="top")
            return cell

        def write_general():
            workbook = Workbook(write_only=True)
            try:
                sheet = workbook.create_sheet("Sheet1")
                try:
                    sheet.append([header_cell(label, sheet) for label in df.columns])
                    columns = [df.iloc[:, i].tolist() for i in range(df.shape[1])]
                    for row in zip(*columns):
                        sheet.append([convert(val, sheet) for val in row])
                except Exception:
                    # 與 df.to_excel 相同：寫到一半失敗時已寫入的部分仍會存檔，例外照常向外拋出
                    try:
                        workbook.save(out_path)
                    except Exception:
                        pass
                    raise
                workbook.save(out_path)
            finally:
                workbook.close()

        def write_all_float():
            # 骨架（含標題列與樣式）由 openpyxl 產生，資料列 XML 直接串流寫入同一個 zip；
            # 版面不符預期時回傳 False（尚未動到 out_path），由呼叫端改走一般路徑。
            workbook = Workbook(write_only=True)
            skeleton = _io.BytesIO()
            try:
                sheet = workbook.create_sheet("Sheet1")
                sheet.append([header_cell(label, sheet) for label in df.columns])
                workbook.save(skeleton)
            finally:
                workbook.close()
            tail_marker = b"</sheetData>"
            with _zipfile.ZipFile(skeleton) as source:
                sheet_xml = source.read("xl/worksheets/sheet1.xml")
                if sheet_xml.count(tail_marker) != 1 or b'<row r="1"' not in sheet_xml:
                    return False
                head, tail = sheet_xml.split(tail_marker)
                letters = [get_column_letter(i + 1) for i in range(df.shape[1])]
                columns = [df.iloc[:, i].tolist() for i in range(df.shape[1])]
                with _zipfile.ZipFile(out_path, "w", _zipfile.ZIP_DEFLATED) as target:
                    for item in source.infolist():
                        if item.filename != "xl/worksheets/sheet1.xml":
                            target.writestr(item, source.read(item.filename))
                            continue
                        entry = _zipfile.ZipInfo(item.filename, date_time=item.date_time)
                        entry.compress_type = _zipfile.ZIP_DEFLATED
                        with target.open(entry, "w") as stream:
                            stream.write(head)
                            pieces = []
                            for number, row in enumerate(zip(*columns), start=2):
                                cells = []
                                for letter, val in zip(letters, row):
                                    if val != val:  # NaN -> 空字串儲存格（與 pandas 相同）
                                        cells.append('<c r="%s%d" t="inlineStr" />' % (letter, number))
                                    elif val == inf or val == -inf:
                                        cells.append('<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>'
                                                     % (letter, number, "inf" if val > 0 else "-inf"))
                                    else:
                                        cells.append('<c r="%s%d" t="n"><v>%.16g</v></c>' % (letter, number, val))
                                pieces.append('<row r="%d">%s</row>' % (number, "".join(cells)))
                                if len(pieces) >= 2000:
                                    stream.write("".join(pieces).encode("utf-8"))
                                    pieces.clear()
                            stream.write("".join(pieces).encode("utf-8"))
                            stream.write(tail_marker + tail)
            return True

        all_float = df.shape[1] > 0 and all(
            isinstance(dtype, np.dtype) and dtype.kind == "f" and dtype.itemsize <= 8 for dtype in df.dtypes)
        if not (all_float and write_all_float()):
            write_general()
        self.generated_excel_paths.append(out_path)

    def plot_and_save(self, df, filename, folder):
        # 只畫四條：繪圖核心在 render_plot（Figure，不用 pyplot）；此方法以目前 UI 狀態建快照後委派
        img_path = render_plot(df, filename, folder, self._capture_settings())
        if img_path is not None:
            self.generated_image_paths.append(img_path)

    def open_files(self, paths: List[str]) -> None:
        """使用系統預設程式開啟檔案。

        Args:
            paths: 要開啟的檔案路徑列表
        """
        open_commands = {
            "Windows": lambda p: os.startfile(p),
            "Darwin": lambda p: subprocess.run(["open", p], check=True),
            "Linux": lambda p: subprocess.run(["xdg-open", p], check=True)
        }
        
        system = platform.system()
        open_func = open_commands.get(system)
        
        if not open_func:
            print(f"不支援的作業系統：{system}")
            return
            
        for p in paths:
            try:
                open_func(p)
            except Exception as e:
                print(f"無法開啟檔案 {p}: {e}")

    # ---------------- 背景處理（QThread）----------------
    processing_finished = Signal(object)   # RunSummary（失敗時為 None）；UI 已恢復、彙總對話框已處理之後發出

    def _capture_settings(self) -> RunSettings:
        """在主執行緒擷取目前 UI 設定的不可變快照（工作執行緒只使用這份快照）。"""
        extras_now = self._get_checked_extra_channels()
        return RunSettings(
            wanted_channels=tuple(dict.fromkeys(self.base_channels + sorted(extras_now))),
            show_t_case=self.show_t_case.isChecked(),
            show_t_evap=self.show_t_evap.isChecked(),
            show_t_tsd=self.show_t_tsd.isChecked(),
            show_t_set=self.show_t_set.isChecked(),
            annotate_min_max=self.annotate_min_max.isChecked(),
            annotate_gap_tsd=self.annotate_gap_tsd.isChecked(),
            annotate_gap_case=self.annotate_gap_case.isChecked(),
            enable_excel=self.enable_excel.isChecked(),
            auto_open_excel=self.auto_open_excel.isChecked(),
            auto_open_plot=self.auto_open_plot.isChecked(),
        )

    def is_processing(self) -> bool:
        return self._processing

    def _set_running_ui(self, running: bool) -> None:
        self.btn_run.setEnabled(not running)
        self.btn_cancel.setEnabled(running)
        for group in self._input_groups:
            group.setEnabled(not running)

    def process_files(self):
        if self._processing:
            return                      # 處理中再次觸發：忽略
        if not self.file_paths:
            QMessageBox.critical(self, "錯誤", "請選擇至少一個 TDMS 檔案"); return

        self.generated_excel_paths.clear(); self.generated_image_paths.clear()
        settings = self._capture_settings()
        self._run_settings = settings
        self._processing = True
        self._set_running_ui(True)
        total = len(self.file_paths) * settings.steps_per_file()
        self.progress_bar.setRange(0, total); self.progress_bar.setValue(0)
        self.label_status.setText("準備中…")

        thread = QThread(self)
        worker = ProcessWorker(list(self.file_paths), settings)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        thread.finished.connect(self._on_thread_finished)
        # worker 屬於工作執行緒：必須在該執行緒「結束前」排程刪除（Qt 不會處理已結束執行緒佇列中的延遲刪除事件，
        # 在 thread.wait() 之後才從 GUI 執行緒呼叫 deleteLater 會讓 worker 永遠不被銷毀）。receiver 是 worker 本身
        # （與 finished 的發出者同屬工作執行緒 → 直接呼叫）；逾時分支執行緒仍在執行，finished 尚未發出，不會提前銷毀。
        thread.finished.connect(worker.delete_when_thread_finishes)
        worker.progress.connect(self._on_progress)
        worker.file_done.connect(self._on_file_done)
        worker.finished_all.connect(self._on_finished_all)
        worker.failed.connect(self._on_failed)
        self._thread, self._worker = thread, worker
        thread.start()

    def cancel_processing(self) -> None:
        """要求取消（合作式：工作執行緒在階段之間檢查，已寫出的檔案保留）。"""
        if self._processing and self._worker is not None:
            self._worker.request_cancel()
            self.btn_cancel.setEnabled(False)
            self.label_status.setText("取消中…（目前階段結束後停止）")

    @Slot(int, int, str)
    def _on_progress(self, done: int, total: int, text: str) -> None:
        self.progress_bar.setRange(0, max(total, 1)); self.progress_bar.setValue(done)
        self.label_status.setText(f"{text}（{done}/{total}）")

    @Slot(object)
    def _on_file_done(self, result) -> None:
        if result.excel_path: self.generated_excel_paths.append(result.excel_path)
        if result.image_path: self.generated_image_paths.append(result.image_path)

    def _release_thread(self) -> bool:
        """要求工作執行緒結束並等待。只有 thread.wait() 成功（執行緒確實已結束）之後才清除 _thread／_worker
        並對 thread deleteLater()；逾時回傳 False，且 _thread／_worker 原封不動（不清除、不排程刪除）。"""
        thread, worker = self._thread, self._worker
        if thread is None:
            return True
        thread.quit()
        if not thread.wait(THREAD_WAIT_TIMEOUT_MS):
            return False
        self._thread = self._worker = None
        thread.deleteLater()                # QThread 物件屬於 GUI 執行緒；worker 由 thread.finished 時刪除（見 process_files）
        return True

    @Slot(object)
    def _on_finished_all(self, summary) -> None:
        self._outcome = ("done", summary)
        self._try_complete()

    @Slot(str)
    def _on_failed(self, message: str) -> None:
        self._outcome = ("failed", message)
        self._try_complete()

    @Slot()
    def _on_thread_finished(self) -> None:
        self._try_complete()                # 先前 wait 逾時的收尾會在執行緒真的結束後於此重試

    def _try_complete(self) -> None:
        """收尾（完成／失敗共用）：確認執行緒已結束才恢復 UI；否則保持受保護的「處理中」狀態，
        等 QThread.finished 再重試。"""
        if self._outcome is None:
            return
        if not self._release_thread():
            self.label_status.setText("取消中，等待目前階段結束…" if self._cancel_pending() else "等待處理執行緒結束…")
            return
        kind, payload = self._outcome
        self._outcome = None
        self._processing = False
        self._set_running_ui(False)
        if kind == "failed":
            self.label_status.setText(f"處理失敗：{payload}")
            if not self._closing:
                QMessageBox.critical(self, "錯誤", f"處理過程發生非預期錯誤：{payload}")
            self.processing_finished.emit(None)
            return
        summary = payload
        self.progress_bar.setRange(0, max(summary.total_files * self._run_settings.steps_per_file(), 1))
        if summary.cancelled:
            self.label_status.setText(f"已取消：完成 {summary.processed}／{summary.total_files} 檔")
        else:
            self.label_status.setText(
                f"完成：成功 {summary.succeeded} 檔、略過 {summary.skipped} 檔、耗時 {summary.elapsed:.1f} s")
            self.progress_bar.setValue(self.progress_bar.maximum())
        if not self._closing:
            self._show_summary_dialog(summary)
            if not summary.cancelled:
                settings = self._run_settings
                if settings.enable_excel and settings.auto_open_excel:
                    self.open_files(list(self.generated_excel_paths))
                if settings.auto_open_plot:
                    self.open_files(list(self.generated_image_paths))
        self.processing_finished.emit(summary)

    def _cancel_pending(self) -> bool:
        return self._worker is not None and self._worker.cancel_requested()


    def _show_summary_dialog(self, summary) -> None:
        """全部結束後至多一個彙總對話框；沒有可報告的內容就不彈。"""
        if not summary.messages:
            return
        lines = [text for _kind, text in summary.messages]
        body = "\n".join(lines[:20])
        if len(lines) > 20:
            body += f"\n另有 {len(lines) - 20} 筆"
        if any(kind == "error" for kind, _text in summary.messages):
            QMessageBox.critical(self, "錯誤", body)
        else:
            QMessageBox.information(self, "TDMS 復原", body)

    def closeEvent(self, event):
        if self._thread is not None:
            self._closing = True
            if self._worker is not None:
                self._worker.request_cancel()
            self.btn_cancel.setEnabled(False)
            if not self._release_thread():
                # 逾時：執行緒仍在執行。保留 _thread／_worker／_processing，拒絕關閉；worker 自然結束時
                # 由 _on_finished_all／_on_failed → _try_complete 完成清理與 UI 恢復，使用者可再次關閉。
                self._closing = False
                self.label_status.setText("取消中，等待目前階段結束…")
                event.ignore()
                return
            self._processing = False
        super().closeEvent(event)

def command_line_tdms_path(argv: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """Validate the optional single-file launch argument without changing no-arg behavior."""
    if len(argv) <= 1:
        return None, None

    candidate = Path(argv[1]).expanduser()
    if candidate.suffix.lower() != TDMS_EXTENSION:
        return None, f"啟動參數不是 .tdms 檔案：\n{candidate}"
    if not candidate.is_file():
        return None, f"找不到啟動參數指定的 TDMS 檔案：\n{candidate}"
    return str(candidate.resolve()), None


if __name__ == "__main__":
    initial_tdms_path, launch_error = command_line_tdms_path(sys.argv)
    app = QApplication(sys.argv)
    w = TDMSFilterApp()
    w.show()
    if initial_tdms_path is not None:
        QTimer.singleShot(0, lambda: w.load_initial_tdms_file(initial_tdms_path))
    elif launch_error is not None:
        QTimer.singleShot(0, lambda: QMessageBox.warning(w, "無法載入 TDMS", launch_error))
    sys.exit(app.exec())
