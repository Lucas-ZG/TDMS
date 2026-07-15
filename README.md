# TDMS Tool Suite

[English](#english) | [繁體中文](#繁體中文)

## English

### Overview

TDMS Tool Suite is a collection of Python desktop utilities for inspecting TDMS test data, exporting thermal channels, generating plots, searching recorded values, and checking machine-to-system time offsets.

### Features

The repository contains three standalone tools:

#### Thermal Data Tool (`tdms.py`)

- Opens one or more `.tdms` files or recursively scans a folder.
- Prioritizes the `COBRA THERMAL DATA` group and falls back to normalized, cross-group channel matching.
- Extracts `TIME`, `M1 T-Case`, `M1 T-Evap`, `M1 TSD`, and `M1 Set Temp`.
- Lets users include additional channels in the Excel output.
- Exports `.xlsx` files and saves `.png` thermal plots beside each source file.
- Optionally annotates TSD minimum/maximum values, the largest Evap–TSD gap, and the largest Evap–Case gap.
- Can open generated files with the operating system's default application.

#### Keyword Search Tool (`keyword search.py`)

- Searches `.tdms` channel values, UTF-8 `.txt` files, or pasted text.
- Supports individual files and recursive folder scanning.
- Provides AND/OR matching, `$$` OR expressions, `%%` wildcards, and `~~N` fuzzy matching.
- Normalizes case and full-width/half-width characters before matching.
- Reports the matching file, group, channel, index, value, or text line.

| Syntax | Meaning | Example |
| --- | --- | --- |
| Space-separated terms | Uses the selected AND/OR mode | `socket contact` |
| `$$` | OR expression | `socket$$contact` |
| `%%` | Wildcard | `temp%%` |
| `~~N` | Fuzzy match at N% similarity | `temprature~~75` |

#### Time Offset Tool (`tdms translater.py`)

- Selects a TDMS file as a reference.
- Accepts a machine timestamp in `YYYY-MM-DD HH:MM` format.
- Calculates the difference between the entered machine time and the current system time.

This utility does not read or modify TDMS channel data.

### Tech Stack

- Python
- PySide6 and PyQt5
- Tkinter
- npTDMS
- pandas
- Matplotlib
- openpyxl
- RapidFuzz, with `difflib` as the fallback

### Requirements

- Python 3.10 or later
- Tkinter, normally included with standard Python distributions

Install the third-party dependencies:

```bash
pip install PySide6 PyQt5 nptdms pandas matplotlib openpyxl rapidfuzz
```

### Getting Started

Run the required utility directly:

```bash
python tdms.py
python "keyword search.py"
python "tdms translater.py"
```

On Windows, `build_exe.bat` can package one or more selected Python entry points into standalone executables. PyInstaller must be installed first:

```bash
pip install pyinstaller
```

### Usage

For thermal export and plotting:

1. Select TDMS files or a folder.
2. Choose the plotted curves, annotations, and optional Excel channels.
3. Start processing.
4. Review the timestamped `.xlsx` and `.png` files beside the original TDMS files.

For keyword search:

1. Choose file mode or pasted-text mode.
2. Enter a query and select AND or OR behavior.
3. Start the search and review the reported match locations.

### Notes

- Thermal extraction is designed primarily for files containing a `COBRA THERMAL DATA` group.
- Missing requested channels are preserved as empty Excel columns.
- Additional selected channels are exported to Excel but are not added to the plot.
- PNG plots are generated even when Excel export is disabled.
- Processing time depends on TDMS file size and channel count.

---

## 繁體中文

### 專案簡介

TDMS Tool Suite 是一組以 Python 開發的桌面工具，用於檢視 TDMS 測試資料、匯出溫度 Channel、產生圖表、搜尋記錄值，以及確認機台時間與系統時間的差異。

### 主要功能

此 Repository 包含三個可獨立執行的工具：

#### 溫度資料工具（`tdms.py`）

- 支援單檔、多檔及資料夾遞迴讀取 `.tdms` 檔案。
- 優先讀取 `COBRA THERMAL DATA` 群組，找不到時會以正規化名稱跨群組比對 Channel。
- 擷取 `TIME`、`M1 T-Case`、`M1 T-Evap`、`M1 TSD` 與 `M1 Set Temp`。
- 可選擇額外 Channel 一併匯出至 Excel。
- 在來源檔案旁輸出 `.xlsx` 資料與 `.png` 溫度曲線圖。
- 可標註 TSD 最高值、最低值、Evap–TSD 最大溫差及 Evap–Case 最大溫差。
- 可使用作業系統預設程式開啟輸出檔案。

#### 關鍵字搜尋工具（`keyword search.py`）

- 搜尋 `.tdms` Channel 值、UTF-8 `.txt` 檔案或直接貼上的文字。
- 支援單一檔案及資料夾遞迴搜尋。
- 支援 AND／OR、`$$` OR 表達式、`%%` 萬用字元及 `~~N` 模糊比對。
- 搜尋前會統一大小寫及全形／半形字元。
- 顯示符合結果的檔案、Group、Channel、Index、Value 或文字行號。

| 語法 | 用途 | 範例 |
| --- | --- | --- |
| 空白分隔 | 依選擇的 AND／OR 模式搜尋 | `socket contact` |
| `$$` | OR 表達式 | `socket$$contact` |
| `%%` | 萬用字元 | `temp%%` |
| `~~N` | 以 N% 相似度進行模糊比對 | `temprature~~75` |

#### 時間差工具（`tdms translater.py`）

- 選擇一個 TDMS 檔案作為參考。
- 接受 `YYYY-MM-DD HH:MM` 格式的機台時間。
- 計算輸入的機台時間與目前系統時間之間的差異。

此工具不會讀取或修改 TDMS Channel 資料。

### 技術架構

- Python
- PySide6 與 PyQt5
- Tkinter
- npTDMS
- pandas
- Matplotlib
- openpyxl
- RapidFuzz，未安裝時改用 `difflib`

### 環境需求

- Python 3.10 以上
- Tkinter，一般已包含於標準 Python 發行版本

安裝第三方套件：

```bash
pip install PySide6 PyQt5 nptdms pandas matplotlib openpyxl rapidfuzz
```

### 開始使用

直接執行需要的工具：

```bash
python tdms.py
python "keyword search.py"
python "tdms translater.py"
```

Windows 使用者可透過 `build_exe.bat` 選擇一個或多個 Python 主程式並封裝為獨立執行檔。使用前需先安裝 PyInstaller：

```bash
pip install pyinstaller
```

### 使用方式

溫度資料匯出與繪圖：

1. 選擇 TDMS 檔案或資料夾。
2. 選擇曲線、標註及需要額外匯出的 Channel。
3. 執行處理。
4. 在原始 TDMS 檔案旁查看帶有時間戳記的 `.xlsx` 與 `.png`。

關鍵字搜尋：

1. 選擇檔案模式或貼上文字模式。
2. 輸入搜尋條件並選擇 AND 或 OR。
3. 執行搜尋並查看符合位置。

### 注意事項

- 溫度資料擷取主要針對包含 `COBRA THERMAL DATA` 群組的檔案設計。
- 找不到的指定 Channel 仍會在 Excel 中保留空欄。
- 額外選取的 Channel 只會匯出至 Excel，不會加入圖表。
- 即使關閉 Excel 匯出，程式仍會產生 PNG 圖表。
- 處理時間會受到 TDMS 檔案大小與 Channel 數量影響。
