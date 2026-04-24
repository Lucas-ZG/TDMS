# TDMS Tool Suite

Author: Lucas.Zhuge

TDMS Tool Suite 是一組針對 TDMS 檔案處理與搜尋需求開發的 GUI 工具，主要用於半導體設備測試資料整理、COBRA thermal data 匯出、溫度曲線繪圖，以及 TDMS / TXT / 貼上文字的關鍵字搜尋。

本專案目前包含兩個工具：

1. **TDMS 小工具 v3.8**  
   用於讀取 TDMS 檔案，匯出 Excel，並產生溫度曲線圖。

2. **TDMS 關鍵字搜尋工具 v1.4-PyQt**  
   用於搜尋 TDMS / TXT 檔案或貼上文字中的指定關鍵字。

---

## 目錄

- [專案功能總覽](#專案功能總覽)
- [工具一：TDMS 小工具 v3.8](#工具一tdms-小工具-v38)
- [工具二：TDMS 關鍵字搜尋工具 v14-pyqt](#工具二tdms-關鍵字搜尋工具-v14-pyqt)
- [安裝環境](#安裝環境)
- [專案結構](#專案結構)
- [執行方式](#執行方式)
- [輸出檔案說明](#輸出檔案說明)
- [打包成執行檔](#打包成執行檔)
- [建議的 .gitignore](#建議的-gitignore)
- [版本紀錄](#版本紀錄)
- [注意事項](#注意事項)

---

## 專案功能總覽

### TDMS 小工具

- 選擇單一或多個 TDMS 檔案
- 選擇資料夾並遞迴掃描所有子資料夾內的 `.tdms`
- 自動讀取 TDMS channel
- 針對 `COBRA THERMAL DATA` 群組優先讀取資料
- 若找不到指定群組，會自動嘗試跨群組搜尋 channel
- 固定輸出主要溫度資料欄位
- 支援額外 channel 勾選輸出
- 勾選的 channel 會完整輸出至 Excel
- 找不到的 channel 會自動建立空欄
- 支援 Excel 匯出
- 支援 PNG 溫度曲線圖輸出
- 支援自動開啟 Excel
- 支援自動開啟圖檔
- 支援 TSD 最高 / 最低溫標註
- 支援 Evap 與 TSD 最大差值標註
- 支援 Evap 與 Case 最大差值標註

### TDMS 關鍵字搜尋工具

- 支援搜尋 `.tdms` 檔案
- 支援搜尋 `.txt` 檔案
- 支援選擇資料夾並遞迴搜尋檔案
- 支援貼上文字搜尋
- 支援 OR 搜尋
- 支援 AND 搜尋
- 支援 `$$` OR 語法
- 支援 `%%` 萬用字元搜尋
- 支援 `~~N` 模糊比對
- 支援大小寫正規化
- 支援全形 / 半形正規化
- 顯示搜尋結果所在檔名、群組、channel、index 或行號

---

## 工具一：TDMS 小工具 v3.8

### 版本

```text
VERSION = "v3.8"
```

### 主要用途

TDMS 小工具主要用於處理測試設備產生的 TDMS 檔案，讀取其中的 thermal data，並輸出為 Excel 與 PNG 圖表。

此工具特別針對以下群組設計：

```text
COBRA THERMAL DATA
```

程式會優先搜尋此群組中的 channel。若找不到完全符合的群組名稱，會透過名稱正規化與模糊邏輯進行比對。

### 固定輸出欄位

Excel 固定輸出以下 5 個欄位：

```text
TIME
M1 T-Case
M1 T-Evap
M1 TSD
M1 Set Temp
```

### 圖表曲線

圖表可繪製以下四條主要曲線：

```text
M1 T-Case
M1 T-Evap
M1 TSD
M1 Set Temp
```

使用者可在 GUI 中勾選是否顯示各曲線。

### 額外欄位輸出

程式會掃描 TDMS 檔案中的 channel，並建立額外欄位清單。

使用者可在 GUI 中勾選需要額外輸出的 channel。

處理邏輯如下：

| 狀況 | 處理方式 |
|---|---|
| 找得到 channel | 將完整資料輸出至 Excel |
| 找不到 channel | 在 Excel 中建立空欄 |
| 勾選額外 channel | 只影響 Excel，不影響圖表 |
| 未勾選額外 channel | 只輸出固定欄位 |

### Channel 名稱比對邏輯

程式會依照以下順序尋找 channel：

1. 群組內精確比對
2. 群組內空白正規化比對
3. 群組內英數字正規化比對
4. 群組內包含關係比對
5. 跨群組搜尋

此設計可降低 TDMS channel 名稱因空白、大小寫、符號差異導致讀取失敗的機率。

### 標註功能

GUI 提供以下標註選項：

```text
標註 TSD 的最高/最低溫
標註 Evap 與 TSD 間最大差值 (Gap)
標註 Evap 與 Case 間最大差值 (Gap)
```

#### TSD 最高 / 最低溫

若勾選此功能，程式會在圖表中標示：

- M1 TSD 最大值
- M1 TSD 最小值

#### Evap 與 TSD 最大差值

計算公式：

```text
Gap = abs(M1 T-Evap - M1 TSD)
```

程式會找出最大 gap，並在圖上標註：

- 最大 gap 數值
- 對應時間點

#### Evap 與 Case 最大差值

計算公式：

```text
Gap = abs(M1 T-Evap - M1 T-Case)
```

程式會找出最大 gap，並在圖上標註：

- 最大 gap 數值
- 對應時間點

### GUI 區塊說明

#### 檔案來源

| 按鈕 | 功能 |
|---|---|
| 選擇 TDMS 檔案（可多選） | 選擇一個或多個 `.tdms` 檔案 |
| 選擇資料夾（遞迴掃描 .tdms） | 選擇資料夾並自動搜尋所有子資料夾中的 TDMS |

#### 輸出設定

| 選項 | 功能 |
|---|---|
| 是否匯出 Excel | 勾選後產生 `.xlsx` |
| 匯出後自動開啟 Excel | 匯出後使用系統預設程式開啟 Excel |
| 儲存後自動開啟圖檔 | 儲存後開啟 PNG 圖表 |

#### 顯示曲線

| 選項 | 功能 |
|---|---|
| M1 T-Case | 顯示 Case 溫度曲線 |
| M1 T-Evap | 顯示 Evap 溫度曲線 |
| M1 TSD | 顯示 TSD 溫度曲線 |
| M1 Set Temp | 顯示設定溫度曲線 |

#### 額外欄位

| 按鈕 | 功能 |
|---|---|
| 全選 | 勾選所有額外 channel |
| 全不選 | 取消勾選所有額外 channel |

---

## 工具二：TDMS 關鍵字搜尋工具 v1.4-PyQt

### 版本

```text
VERSION = "v1.4-PyQt"
```

### 主要用途

TDMS 關鍵字搜尋工具用於快速搜尋 TDMS、TXT 檔案或貼上文字中的指定內容。

適用情境：

- 搜尋 TDMS 檔案中是否出現特定錯誤字串
- 搜尋 channel data 是否包含特定文字
- 搜尋 log / txt 檔案內容
- 搜尋信件、會議紀錄、debug note 中的關鍵字
- 透過模糊比對找出可能拼錯的字

### 搜尋目標

GUI 提供兩種搜尋目標：

```text
檔案（.tdms/.txt）
貼上文字
```

### 支援檔案類型

```text
.tdms
.txt
```

### 搜尋條件

| 條件 | 說明 |
|---|---|
| OR（任一條件） | 任一關鍵字符合即顯示結果 |
| AND（全部條件） | 所有關鍵字都符合才顯示結果 |

### 關鍵字語法

#### 一般搜尋

```text
indium
```

代表搜尋包含 `indium` 的內容。

#### 使用空白分隔多個條件

```text
socket ddr contact
```

若選擇 OR：任一條件符合即可。  
若選擇 AND：全部條件都必須符合。

#### 使用 `$$` 表示 OR

```text
socket$$ddr$$contact
```

代表搜尋：

```text
socket 或 ddr 或 contact
```

#### 使用 `%%` 表示萬用字元

```text
ind%%
```

可匹配：

```text
indium
ind123
ind_xxx
```

#### 使用 `~~N` 表示模糊比對

```text
temprature~~75
```

代表使用 75% 相似度進行模糊搜尋。

此功能適合用於搜尋可能拼錯的字，例如：

```text
temprature
```

可能接近：

```text
temperature
```

### 模糊比對套件

若有安裝 `rapidfuzz`，程式會優先使用 `rapidfuzz`。

若沒有安裝，程式會自動改用 Python 內建的 `difflib`。

建議安裝：

```bash
pip install rapidfuzz
```

### 搜尋結果顯示

#### TDMS 搜尋結果

TDMS 搜尋結果會顯示：

```text
檔案名稱
符合條件
Group
Channel
Index
Value
```

#### TXT 搜尋結果

TXT 搜尋結果會顯示：

```text
檔案名稱
符合條件
Line number
Line content
```

#### 貼上文字搜尋結果

貼上文字搜尋結果會顯示：

```text
PASTE
符合條件
Line number
Line content
```

---

## 安裝環境

建議使用：

```text
Python 3.10+
```

### TDMS 小工具 v3.8 需要套件

```bash
pip install PySide6 nptdms pandas matplotlib openpyxl
```

### TDMS 關鍵字搜尋工具 v1.4-PyQt 需要套件

```bash
pip install PyQt5 nptdms rapidfuzz
```

### 完整 requirements.txt

可建立 `requirements.txt`：

```text
PySide6
PyQt5
nptdms
pandas
matplotlib
openpyxl
rapidfuzz
```

安裝方式：

```bash
pip install -r requirements.txt
```

---

## 專案結構

建議專案結構如下：

```text
tdms-tool-suite/
│
├─ tdms_filter_tool.py
├─ tdms_keyword_searcher.py
├─ requirements.txt
├─ README.md
└─ output/
```

### 檔案說明

| 檔案 | 說明 |
|---|---|
| tdms_filter_tool.py | TDMS 小工具主程式 |
| tdms_keyword_searcher.py | TDMS 關鍵字搜尋工具主程式 |
| requirements.txt | Python 套件需求清單 |
| README.md | 專案說明文件 |
| output/ | 可自行建立，用於整理輸出檔案 |

---

## 執行方式

### 執行 TDMS 小工具

```bash
python tdms_filter_tool.py
```

### 執行 TDMS 關鍵字搜尋工具

```bash
python tdms_keyword_searcher.py
```

---

## 使用流程

### TDMS 小工具使用流程

1. 開啟程式。
2. 選擇 TDMS 檔案，或選擇資料夾。
3. 若選擇資料夾，程式會自動遞迴搜尋所有 `.tdms`。
4. 確認是否匯出 Excel。
5. 確認是否自動開啟 Excel。
6. 確認是否自動開啟圖檔。
7. 勾選需要顯示的曲線。
8. 勾選需要的標註功能。
9. 勾選需要額外輸出的 channel。
10. 點選「執行處理 TDMS 檔案」。
11. 程式產生 Excel 與 PNG 圖檔。

### TDMS 關鍵字搜尋工具使用流程

1. 開啟程式。
2. 選擇搜尋目標：檔案或貼上文字。
3. 若選擇檔案，選擇檔案或資料夾。
4. 若選擇貼上文字，將內容貼到貼上內容區。
5. 在關鍵字欄位輸入搜尋條件。
6. 選擇 OR 或 AND。
7. 點選「開始搜尋」。
8. 在搜尋結果區查看符合內容。

---

## 輸出檔案說明

### TDMS 小工具輸出

輸出位置：原 TDMS 檔案所在資料夾。

#### Excel 檔名格式

```text
原檔名_YYYYMMDD_HHMMSS.xlsx
```

#### 圖片檔名格式

```text
原檔名_YYYYMMDD_HHMMSS.png
```

### Excel 內容

Excel 會包含：

1. 固定欄位
2. 使用者勾選的額外 channel
3. 找不到的 channel 空欄

### PNG 圖表內容

PNG 圖表會包含：

1. 使用者勾選的主要溫度曲線
2. TSD 最高 / 最低溫標註
3. Evap-TSD 最大 gap 標註
4. Evap-Case 最大 gap 標註

---

## 打包成執行檔

可使用 PyInstaller 打包成 Windows `.exe`。

### 安裝 PyInstaller

```bash
pip install pyinstaller
```

### 打包 TDMS 小工具

```bash
pyinstaller --onefile --windowed tdms_filter_tool.py
```

### 打包 TDMS 關鍵字搜尋工具

```bash
pyinstaller --onefile --windowed tdms_keyword_searcher.py
```

### 打包完成位置

```text
dist/
```

---

## 建議的 .gitignore

```gitignore
__pycache__/
*.pyc
*.pyo
*.spec
build/
dist/
*.xlsx
*.png
*.csv
.env
.venv/
venv/
```

若不想把 TDMS 原始檔上傳到 GitHub，也可以加入：

```gitignore
*.tdms
```

---

## 版本紀錄

### TDMS 小工具

#### v3.8

- 穩定版本整合
- 全面修正勾選欄位不同步問題
- 新增即時讀取 UI 勾選狀態機制
- 確保使用者勾選的所有 channel 皆正確輸出
- 勾選的 channel 會完整複製整列資料進 Excel
- 若找不到 channel，會自動建立空欄
- 強化 channel 名稱比對邏輯
- 比對順序：群組內精確 → 寬鬆 → 跨群組搜尋
- 介面維持一致
- 內部邏輯更簡潔與穩定
- 圖表仍僅繪製四條主要溫度曲線

#### v3.7.3

- 勾選的 channel 會自動加入 Excel
- 找不到 channel 時產生空欄
- 強化多群組匹配機制
- 支援跨群組尋找 channel
- 額外欄位改為多欄勾選清單
- 增加全選 / 全不選功能

#### v3.7.2

- 新增同時支援選擇檔案與選擇資料夾兩種來源
- 匯入資料夾時自動遞迴搜尋所有 `.tdms`
- 合併搜尋結果至清單
- 增加額外欄位下拉選單

#### v3.7.1

- 改進資料夾搜尋邏輯
- 支援多層資料夾自動掃描
- 新增版本號與更新紀錄系統

#### v3.7

- 新增選擇資料夾功能
- 可遞迴掃描所有子資料夾內的 TDMS
- 新增是否匯出 Excel 選項
- 新增自動開啟 Excel / 圖檔選項
- 新增圖上標註功能
- 可顯示 TSD 最高 / 最低點
- 可顯示 Evap / Case 差值 Gap

### TDMS 關鍵字搜尋工具

#### v1.4-PyQt

- 修正 UX
- 拆成「關鍵字輸入」與「貼上內容」兩個區塊
- 避免把信件內容誤當 query
- 新增搜尋目標切換
- 支援搜尋檔案
- 支援搜尋貼上文字
- 保留 AND / OR 搜尋
- 保留 `$$ OR` 語法
- 保留 `%%` 萬用字元語法
- 保留 `~~N` 模糊比對語法
- 支援大小寫正規化
- 支援全半形正規化

---

## 注意事項

1. TDMS 小工具主要針對 `COBRA THERMAL DATA` 群組設計。
2. 若 TDMS 檔案內群組名稱不同，程式會嘗試使用正規化方式搜尋。
3. 若找不到指定 channel，Excel 仍會建立該欄位，但內容為空。
4. 額外 channel 只會輸出至 Excel，不會加入圖表。
5. 圖表固定只處理主要四條溫度曲線。
6. 若 TDMS 檔案很大，Excel 匯出與搜尋時間會增加。
7. 自動開啟檔案功能依賴作業系統預設程式。
8. Windows 使用 `os.startfile` 開啟檔案。
9. macOS 使用 `open` 指令開啟檔案。
10. Linux 使用 `xdg-open` 指令開啟檔案。
11. 關鍵字搜尋工具中的模糊比對建議安裝 `rapidfuzz`，效能較佳。
12. 若未安裝 `rapidfuzz`，程式會自動使用 `difflib`。

---

## License

此工具為內部使用工具。若需公開或商業使用，請先確認授權範圍。

---

## Author

Lucas.Zhuge
