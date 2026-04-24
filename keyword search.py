# Author: Lucas.Zhuge
# VERSION NOTE:
# 1.4-PyQt:
#   - 修正 UX：拆成「關鍵字輸入」與「貼上內容」兩個區塊，避免把信件內容誤當 query
#   - 新增搜尋目標切換：搜尋檔案 / 搜尋貼上文字
#   - 保留：AND/OR、$$ OR、%% 萬用字元、~~N 模糊比對、大小寫/全半形正規化
VERSION = "v1.4-PyQt"

import os
import re
import unicodedata
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QFileDialog,
    QComboBox, QRadioButton, QButtonGroup
)
from nptdms import TdmsFile


def _try_fuzzy_ratio(a: str, b: str) -> float:
    try:
        from rapidfuzz.fuzz import ratio  # type: ignore
        return float(ratio(a, b))
    except Exception:
        import difflib
        return difflib.SequenceMatcher(None, a, b).ratio() * 100.0


class TDMSKeywordSearcher(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"TDMS 關鍵字搜尋工具  {VERSION}")
        self.resize(1200, 820)
        self.file_paths = []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout()

        # 檔案選擇列（僅檔案模式會用到）
        file_row = QHBoxLayout()
        self.file_label = QLabel("未選擇檔案")
        self.btn_files = QPushButton("選擇檔案", clicked=self.select_files)
        self.btn_folder = QPushButton("選擇資料夾", clicked=self.select_folder)
        file_row.addWidget(self.btn_files)
        file_row.addWidget(self.btn_folder)
        file_row.addWidget(self.file_label)
        layout.addLayout(file_row)

        # 搜尋目標切換：檔案 / 貼上文字
        target_row = QHBoxLayout()
        target_row.addWidget(QLabel("搜尋目標："))
        self.rb_files = QRadioButton("檔案（.tdms/.txt）")
        self.rb_paste = QRadioButton("貼上文字")
        self.rb_files.setChecked(True)

        self.target_group = QButtonGroup(self)
        self.target_group.addButton(self.rb_files)
        self.target_group.addButton(self.rb_paste)
        self.rb_files.toggled.connect(self._apply_target_mode)

        target_row.addWidget(self.rb_files)
        target_row.addWidget(self.rb_paste)
        target_row.addStretch(1)
        layout.addLayout(target_row)

        # 關鍵字 + 模式 + 搜尋按鈕
        kw_row = QHBoxLayout()
        kw_row.addWidget(QLabel("關鍵字："))
        self.input_keyword = QLineEdit()
        self.input_keyword.setPlaceholderText("keyword，例如：indium、ind%%、socket$$ddr、temprature~~75")
        kw_row.addWidget(self.input_keyword)

        kw_row.addWidget(QLabel("條件："))
        self.combo_mode = QComboBox()
        self.combo_mode.addItems(["OR（任一條件）", "AND（全部條件）"])
        self.combo_mode.setCurrentIndex(0)
        kw_row.addWidget(self.combo_mode)

        self.btn_search = QPushButton("開始搜尋", clicked=self.search)
        kw_row.addWidget(self.btn_search)
        layout.addLayout(kw_row)

        # 語法說明
        help_box = QTextEdit()
        help_box.setReadOnly(True)
        help_box.setFixedHeight(85)
        help_box.setPlainText(
            "語法：\n"
            "- $$ 表示 OR：socket$$ddr$$contact\n"
            "- %% 表示萬用字元：ind%% -> indium / indixxx / indi123\n"
            "- ~~N 表示模糊比對：temprature~~75\n"
        
        )
        layout.addWidget(help_box)

        # 貼上內容區（貼上文字模式使用）
        layout.addWidget(QLabel("貼上內容（僅『貼上文字』模式使用）："))
        self.paste_box = QTextEdit()
        self.paste_box.setPlaceholderText("把信件/LOG/會議紀錄貼在這裡。然後在上方『關鍵字』輸入要找的字。")
        self.paste_box.setFixedHeight(180)
        layout.addWidget(self.paste_box)

        # 結果區
        layout.addWidget(QLabel("搜尋結果："))
        self.result_box = QTextEdit()
        self.result_box.setReadOnly(True)
        layout.addWidget(self.result_box)

        self.setLayout(layout)
        self._apply_target_mode()

    def _apply_target_mode(self):
        is_files = self.rb_files.isChecked()
        self.btn_files.setEnabled(is_files)
        self.btn_folder.setEnabled(is_files)
        self.file_label.setEnabled(is_files)

        # 貼上文字模式下，提示更明顯
        self.paste_box.setEnabled(not is_files)

    def normalize(self, text):
        return unicodedata.normalize("NFKC", str(text)).lower()

    # -----------------------------
    # Query 解析（關鍵字欄位使用）
    # -----------------------------
    def parse_query_terms(self, raw: str):
        raw = (raw or "").strip()
        if not raw:
            return []
        if "$$" in raw:
            return [p.strip() for p in raw.split("$$") if p.strip()]
        # 無 $$：允許用空白拆多條件（方便）
        return [t.strip() for t in re.split(r"\s+", raw) if t.strip()]

    def build_matchers(self, raw_query: str):
        terms = self.parse_query_terms(raw_query)
        if not terms:
            return []

        matchers = []
        for term in terms:
            term = (term or "").strip()
            if not term:
                continue

            fuzzy_threshold = None
            if "~~" in term:
                left, right = term.rsplit("~~", 1)
                left = left.strip()
                right = right.strip()
                if left:
                    try:
                        fuzzy_threshold = float(right)
                    except Exception:
                        fuzzy_threshold = 80.0
                    term = left

            term_norm = self.normalize(term)

            if "%%" in term_norm:
                pattern = re.escape(term_norm).replace(r"\%\%", ".*")
                rx = re.compile(pattern, re.IGNORECASE)

                def _m(s, _rx=rx):
                    return _rx.search(s) is not None

                matchers.append((_m, f"WILDCARD:{term}"))
                continue

            if fuzzy_threshold is not None:
                th = max(0.0, min(100.0, fuzzy_threshold))

                def _m(s, _t=term_norm, _th=th):
                    r1 = _try_fuzzy_ratio(_t, s)
                    r2 = _try_fuzzy_ratio(s, _t)
                    return max(r1, r2) >= _th

                matchers.append((_m, f"FUZZY({int(th)}):{term}"))
                continue

            def _m(s, _t=term_norm):
                return _t in s

            matchers.append((_m, f"CONTAINS:{term}"))

        return matchers

    def match_eval(self, s_norm: str, matchers, mode_and: bool):
        tags = []
        for fn, tag in matchers:
            ok = fn(s_norm)
            if ok:
                tags.append(tag)
            elif mode_and:
                return False, []
        return (len(tags) > 0) if not mode_and else True, tags

    # -----------------------------
    # 檔案選擇
    # -----------------------------
    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "選擇檔案", "", "TDMS/TXT Files (*.tdms *.txt)")
        if files:
            self.file_paths = files
            self.file_label.setText(f"已選擇 {len(files)} 個檔案")

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "選擇資料夾")
        if folder:
            matched = []
            for dirpath, _, filenames in os.walk(folder):
                for name in filenames:
                    if name.lower().endswith((".tdms", ".txt")):
                        matched.append(os.path.join(dirpath, name))
            self.file_paths = matched
            self.file_label.setText(f"已選擇資料夾，共 {len(matched)} 個檔案")

    # -----------------------------
    # 搜尋：檔案模式 / 貼上文字模式
    # -----------------------------
    def search_in_pasted_text(self, matchers, mode_and: bool):
        text = self.paste_box.toPlainText()
        if not text.strip():
            self.result_box.setPlainText("[提醒] 目前是『貼上文字』模式，請先在『貼上內容』區貼文字。")
            return

        self.result_box.clear()
        found = False

        lines = text.splitlines()
        for line_no, line in enumerate(lines, 1):
            s_norm = self.normalize(line)
            hit, tags = self.match_eval(s_norm, matchers, mode_and)
            if hit:
                found = True
                tag_str = " | ".join(tags) if tags else ""
                self.result_box.append(f"[PASTE]  {tag_str}\n  Line {line_no}: {line}")

        if not found:
            self.result_box.append("未找到任何符合的關鍵字。")

    def search_in_files(self, matchers, mode_and: bool):
        if not self.file_paths:
            self.result_box.setPlainText("[提醒] 目前是『檔案』模式，請先選擇檔案或資料夾。")
            return

        self.result_box.clear()
        found = False

        for path in self.file_paths:
            ext = os.path.splitext(path)[1].lower()
            try:
                if ext == ".tdms":
                    tdms_file = TdmsFile.read(path)
                    for group in tdms_file.groups():
                        for channel in group.channels():
                            data = channel[:]
                            for idx, v in enumerate(data):
                                s_norm = self.normalize(v)
                                hit, tags = self.match_eval(s_norm, matchers, mode_and)
                                if hit:
                                    found = True
                                    tag_str = " | ".join(tags) if tags else ""
                                    self.result_box.append(
                                        f"[{os.path.basename(path)}]  {tag_str}\n"
                                        f"  Group: {group.name}  Channel: {channel.name}  Index: {idx}  Value: {v}"
                                    )

                elif ext == ".txt":
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
                        for line_no, line in enumerate(f, 1):
                            s_norm = self.normalize(line)
                            hit, tags = self.match_eval(s_norm, matchers, mode_and)
                            if hit:
                                found = True
                                tag_str = " | ".join(tags) if tags else ""
                                self.result_box.append(
                                    f"[{os.path.basename(path)}]  {tag_str}\n"
                                    f"  Line {line_no}: {line.strip()}"
                                )
            except Exception as e:
                self.result_box.append(f"[讀取失敗] {path}  Error: {e}")

        if not found:
            self.result_box.append("未找到任何符合的關鍵字。")

    def search(self):
        raw_kw = self.input_keyword.text().strip()
        if not raw_kw:
            self.result_box.setPlainText("[提醒] 請在『關鍵字』欄位輸入要找的內容。")
            return

        matchers = self.build_matchers(raw_kw)
        if not matchers:
            self.result_box.setPlainText("[提醒] 關鍵字解析失敗，請確認輸入格式。")
            return

        mode_and = (self.combo_mode.currentIndex() == 1)

        if self.rb_paste.isChecked():
            self.search_in_pasted_text(matchers, mode_and)
        else:
            self.search_in_files(matchers, mode_and)


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    win = TDMSKeywordSearcher()
    win.show()
    sys.exit(app.exec_())
