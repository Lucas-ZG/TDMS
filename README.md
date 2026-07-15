# TDMS Tool Suite

一組以 Python 開發的桌面工具，用於整理、分析與搜尋 TDMS 測試資料。專案包含溫度資料匯出與圖表產生工具，以及支援 TDMS、TXT 和貼上文字的關鍵字搜尋工具。

## 核心功能

### TDMS Data Tool

- 支援單檔、多檔及資料夾遞迴讀取
- 優先解析 `COBRA THERMAL DATA`，並支援跨群組 Channel 搜尋
- 匯出主要溫度資料與自選 Channel 至 Excel
- 產生 T-Case、T-Evap、TSD 與 Set Temperature 曲線圖
- 標註 TSD 最高值、最低值及指定溫差的最大 Gap
- 可選擇是否自動開啟輸出的 Excel 與 PNG

主要輸出欄位：

- `TIME`
- `M1 T-Case`
- `M1 T-Evap`
- `M1 TSD`
- `M1 Set Temp`

### Keyword Search Tool

- 搜尋 TDMS、TXT 或直接貼上的文字
- 支援單檔、多檔及資料夾遞迴搜尋
- 支援 AND、OR、萬用字元及模糊比對
- 自動正規化大小寫及全形／半形字元
- 顯示檔名、Group、Channel、Index 或文字行號

搜尋語法：

| 語法 | 用途 | 範例 |
|---|---|---|
| 空白分隔 | 依目前模式執行 AND 或 OR | `socket contact` |
| `$$` | OR 搜尋 | `socket$$contact` |
| `%%` | 萬用字元 | `temp%%` |
| `~~N` | 以 N% 相似度模糊比對 | `temprature~~75` |

## 使用流程

### 匯出與繪圖

1. 選擇 TDMS 檔案或資料夾。
2. 選擇要匯出的 Channel、顯示曲線及圖表標註。
3. 執行處理。
4. Excel 與 PNG 會輸出至來源檔案所在位置。

### 關鍵字搜尋

1. 選擇檔案、資料夾或貼上文字。
2. 輸入關鍵字並選擇 AND 或 OR 模式。
3. 執行搜尋並查看符合位置與內容。

## 技術

- Python 3.10+
- PySide6 / PyQt5
- npTDMS
- pandas
- Matplotlib
- openpyxl
- RapidFuzz（未安裝時使用 Python `difflib`）

安裝相依套件：

```bash
pip install PySide6 PyQt5 nptdms pandas matplotlib openpyxl rapidfuzz
```

## 注意事項

- 溫度資料工具以 `COBRA THERMAL DATA` 的資料結構為主要設計對象。
- Channel 名稱不完全一致時，程式會依正規化名稱及跨群組搜尋結果進行比對。
- 找不到的指定 Channel 會在 Excel 中保留空欄。
- 額外選取的 Channel 僅匯出至 Excel，不會加入溫度曲線圖。
- 大型 TDMS 檔案的讀取、搜尋與匯出時間會依資料量增加。

## Author

Lucas Zhuge
