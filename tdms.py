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
import sys
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
from PySide6.QtCore import Qt

try:
    from nptdms import TdmsFile
except ImportError:
    TdmsFile = None

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

class TDMSFilterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"TDMS小工具 {VERSION}")
        self.resize(600, 760)

        # 狀態
        self.file_paths = []
        self.generated_excel_paths, self.generated_image_paths = [], []

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

        # 佈局
        main.addWidget(g_file); main.addWidget(g_out); main.addWidget(g_curve); main.addWidget(g_anno)
        main.addWidget(g_extra); main.addWidget(btn_run); main.addStretch()
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

    # ---------------- 勾選清單（聯集 & 群組模糊） ----------------
    def _iter_cobra_groups(self, td: TdmsFile):
        target = _norm_group(TARGET_GROUP_NAME)
        try:
            for grp in td.groups():
                name = getattr(grp, "name", str(grp))
                if target in _norm_group(name): yield grp
        except Exception: return

    def _all_groups(self, td: TdmsFile):
        try:
            return list(td.groups())
        except Exception:
            return []

    def _rebuild_checklist_from_all_files(self):
        if not self.file_paths: self._rebuild_channel_checklist([]); return
        if TdmsFile is None:
            QMessageBox.critical(self, "錯誤", "尚未安裝 nptdms 套件"); self._rebuild_channel_checklist([]); return

        union_names = set()
        for p in self.file_paths:
            try:
                td = TdmsFile.read(p)
            except Exception:
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
        target = _norm_group(TARGET_GROUP_NAME)
        try:
            for g in td.groups():
                if target in _norm_group(getattr(g, "name", str(g))): return g
        except Exception: pass
        try: return td[TARGET_GROUP_NAME]
        except Exception: return None

    def _resolve_in_group(self, grp, wanted: str):
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

    def _resolve_anywhere(self, td: TdmsFile, wanted: str):
        for grp in self._all_groups(td):
            real = self._resolve_in_group(grp, wanted)
            if real is not None: return grp, real
        return None, None

    # ---------------- 讀檔 / 匯出 / 畫圖 ----------------
    def read_tdms_file(self, path: str) -> pd.DataFrame:
        """讀取 TDMS 檔案並轉換為 DataFrame。

        Args:
            path: TDMS 檔案路徑

        Returns:
            包含所需通道資料的 DataFrame

        Raises:
            ImportError: 未安裝 nptdms
            FileNotFoundError: 檔案不存在
            ValueError: 檔案格式錯誤
        """
        if TdmsFile is None:
            raise ImportError("尚未安裝 nptdms 套件")
        
        if not os.path.exists(path):
            raise FileNotFoundError(f"找不到檔案：{path}")

        try:
            td = TdmsFile.read(path)
        except Exception as e:
            raise ValueError(f"無法讀取 TDMS 檔案：{e}")

        # ⚠️ 核心：當下讀 UI 勾選
        extras_now = self._get_checked_extra_channels()
        wanted = list(dict.fromkeys(self.base_channels + sorted(extras_now)))

        cobra = self._find_group_cobra(td)

        # 優先取 TIME 以確定列數
        time_data = None
        if cobra is not None:
            r = self._resolve_in_group(cobra, "TIME")
            if r is None:
                g2, r2 = self._resolve_anywhere(td, "TIME")
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
                real = self._resolve_in_group(grp, name)
            if real is None:
                g2, r2 = self._resolve_anywhere(td, name)
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
        return df

    def export_to_excel(self, df, tdms_path):
        # 全欄輸出（包含空欄）
        base = os.path.splitext(os.path.basename(tdms_path))[0]
        folder = os.path.dirname(tdms_path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(folder, f"{base}_{ts}.xlsx")
        df.to_excel(out_path, index=False)
        self.generated_excel_paths.append(out_path)

    def plot_and_save(self, df, filename, folder):
        # 只畫四條
        if "TIME" not in df.columns: return
        fig, ax = plt.subplots(figsize=(10, 6))
        t = df["TIME"]

        targets = {
            "M1 T-Case": self.show_t_case.isChecked(),
            "M1 T-Evap": self.show_t_evap.isChecked(),
            "M1 TSD": self.show_t_tsd.isChecked(),
            "M1 Set Temp": self.show_t_set.isChecked()
        }
        for col, show in targets.items():
            if show and col in df.columns:
                ax.plot(t, df[col], label=col)
                if self.annotate_min_max.isChecked() and col == "M1 TSD":
                    y = df[col].dropna()
                    if not y.empty:
                        imax, imin = y.idxmax(), y.idxmin()
                        ax.annotate(f"Max: {y[imax]:.1f}", xy=(t[imax], y[imax]),
                                    xytext=(5,5), textcoords='offset points', fontsize=9, color='red')
                        ax.annotate(f"Min: {y[imin]:.1f}", xy=(t[imin], y[imin]),
                                    xytext=(5,-15), textcoords='offset points', fontsize=9, color='blue')

        if self.annotate_gap_tsd.isChecked() and "M1 T-Evap" in df.columns and "M1 TSD" in df.columns:
            gap = (df["M1 T-Evap"] - df["M1 TSD"]).abs().dropna()
            if not gap.empty:
                idx = gap.idxmax()
                ax.annotate(f"Max Evap-TSD Gap: {gap[idx]:.1f} @ Time: {t[idx]:.1f}",
                            xy=(t[idx], df["M1 T-Evap"][idx]), xytext=(10,10),
                            textcoords='offset points', fontsize=9, color='purple',
                            arrowprops=dict(arrowstyle='->', color='purple'))
        if self.annotate_gap_case.isChecked() and "M1 T-Evap" in df.columns and "M1 T-Case" in df.columns:
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
        plt.savefig(img_path); self.generated_image_paths.append(img_path)
        if self.auto_open_plot.isChecked(): plt.show()
        else: plt.close(fig)

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

    def process_files(self):
        if not self.file_paths:
            QMessageBox.critical(self, "錯誤", "請選擇至少一個 TDMS 檔案"); return

        self.generated_excel_paths.clear(); self.generated_image_paths.clear()

        for path in self.file_paths:
            try:
                df = self.read_tdms_file(path)
                if self.enable_excel.isChecked(): self.export_to_excel(df, path)
                self.plot_and_save(df, os.path.basename(path), os.path.dirname(path))
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"處理檔案 {path} 時發生錯誤：\n{e}")

        if self.enable_excel.isChecked() and self.auto_open_excel.isChecked():
            self.open_files(self.generated_excel_paths)
        if self.auto_open_plot.isChecked():
            self.open_files(self.generated_image_paths)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = TDMSFilterApp()
    w.show()
    sys.exit(app.exec())
