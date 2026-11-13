# tmdb-importer

由影片網站抓取劇集、電影同演員資料，轉成**香港繁體中文**，再寫入 [TMDB](https://www.themoviedb.org/) 嘅 **zh-HK** 翻譯。

- 支援網站：Netflix、Disney+、Prime Video、Apple TV、viu、viu.tv
- 寫入內容：劇集 / 電影標題同簡介、季同分集標題同簡介、演員名同簡介、劇照

> ⚠️ 個人自用工具。程式會用**你嘅 TMDB 帳戶直接修改 TMDB**。第一次用請先揀少量項目測試，確認結果啱先大量運行。

---

## 目錄

1. [需要](#需要)
2. [快速開始](#快速開始)
3. [運作流程](#運作流程)
4. [啟動選項](#啟動選項)
5. [bulk 任務表](#bulk-任務表databulk_videosxlsx)
6. [搵 TMDB 網址嘅小工具](#搵-tmdb-網址嘅小工具sub-program)
7. [中文轉換同香港譯名](#中文轉換同香港譯名)
8. [設定檔同資料檔](#設定檔同資料檔)
9. [常見問題](#常見問題)

---

## 需要

| 項目 | 說明 |
|---|---|
| Windows 10 / 11 | |
| Python 3 | 喺 Microsoft Store 版 Python 3.13 測試過。用 python.org 安裝程式時要剔「Add python.exe to PATH」 |
| Microsoft Edge | 程式用 Edge 登入同修改 TMDB |
| TMDB 帳戶 | 建議另外申請 **API Read Access Token**（v4，`eyJ` 開頭），讀資料會快好多 |

其他 Python 套件同 OpenCC 字典，`start.py` 會自動下載安裝，之後每日最多檢查一次更新。
想即時檢查：`python start.py --update-deps`。

## 快速開始

1. 下載或者 clone 呢個 repo。
2. 將 `configs.example.json` 複製做 `configs.json`，填 `tmdb_api`（API Read Access Token）。唔填都用得，不過會慢啲。
3. 雙擊 **`start.bat`**，開啟監控面板。
4. 揀好選項（見[啟動選項](#啟動選項)），撳「▶ 啟動」。
5. 第一次用到 TMDB、Netflix、Disney+ 時，程式會問帳號密碼，可以揀儲存喺 `configs.json`。

唔想用面板的話，可以喺命令視窗執行 `python start.py`，跟住回答問題。

## 運作流程

```
搵網址                    任務表                     處理                         結果
───────────               ──────────────             ──────────────────           ──────────────
自己貼 / 小工具搵  ──→   bulk_videos.xlsx   ──→    抓取影片網站資料       ──→   寫入 TMDB zh-HK
(sub-program)             （Video URL、           轉香港中文、比較 TMDB          Excel Done 欄記錄結果
                           TMDB URL）              （同時進行）
```

- **single 模式**：逐個輸入「影片網址 + TMDB 網址」。
- **bulk 模式**：處理 `data/bulk_videos.xlsx`。可以一次揀幾張工作表，按次序處理；抓取同上傳同時進行。
- **監控面板**（`start.bat`）：顯示進度、目前處理緊邊行、log；可以啟動、停止、暫停，或者只暫停上傳。

## 啟動選項

| 選項 | 意思 | 默認 |
|---|---|---|
| 覆蓋劇集標題 / 描述 | TMDB 已經有香港中文時，係咪用新資料取代（y / n / ask） | |
| 覆蓋分集標題 / 描述 | 同上，用喺分集；留空就跟劇集設定 | 同劇集 |
| 來源無資料時清空 TMDB | 影片網站冇呢項資料時，係咪清空 TMDB 現有內容 | n |
| 維基人名取代現有 | 維基百科嘅香港譯名係咪取代 TMDB 現有嘅香港人名 | n |
| bulk 工作表 | 要處理 `bulk_videos.xlsx` 邊幾張工作表，按次序處理 | Bulk |

以上選項都可以喺 bulk 任務表逐行設定（見下面）。

## bulk 任務表（`data/bulk_videos.xlsx`）

每張工作表一行一個任務：

| 欄位 | 內容 |
|---|---|
| Video URL | 影片網站網址，或者 TMDB 網址 |
| TMDB URL | 對應嘅 TMDB 網址 |
| Split | Netflix 分集標題有「/」時，係咪拆成兩集（y / n）；平時留空 |
| Done | 程式處理完自動填寫 |
| 覆蓋劇集標題、覆蓋劇集描述、覆蓋分集標題、覆蓋分集描述 | **只影響呢一行**：y / n / ask；留空就用啟動選項 |
| 來源無資料時清空TMDB、維基人名取代現有 | **只影響呢一行**：y / n；留空就用啟動選項 |

- 檔案冇嘅話會自動建立；舊檔案缺少欄位，開始時會自動加上。
- 開始處理一張工作表時，會自動：TMDB 網址拎走 slug（`/tv/334779-xxx` → `/tv/334779`）、刪除重複網址、刪除空行。
- 程式運行期間可以喺工作表**底部**加新行，或者改未處理嘅行；唔好喺中間插入或者重新排序。

## 搵 TMDB 網址嘅小工具（`sub-program/`）

雙擊 **`sub-program/tools.bat`** 開小工具面板。每個工具喺新嘅命令視窗運行。

| 工具 | 做咩 |
|---|---|
| TMDB 關鍵字搜尋 | 用關鍵字、TMDB 網址或 ID 搵有中文翻譯嘅劇集 / 電影 |
| 檢查影片網址 | 貼上 TMDB 影片網址，保留有中文內容嘅 |
| 影片 / 演員擴展 | 由影片或演員搵相關演員，再搵佢哋其他有中文嘅作品 |
| 演員名搜尋 | 貼上演員名列表，搵演員同佢哋有中文嘅作品 |
| 維基百科列表 | 由維基百科頁面（例如劇集列表）抽出標題，再搜尋 TMDB |
| TMDB 貢獻 / JAVHD / MissAV | 由 TMDB 首頁、排行榜、用家貢獻或者名單搵演員，再搵有中文嘅作品 |
| viu.tv 節目建立 / 批量更新 / 劇照上載 | 喺 TMDB 建立 viu.tv 節目、加集數、上載劇照（**會直接修改 TMDB**） |

- 搵網址嘅工具**唔會改** `bulk_videos.xlsx`，結果寫入 `sub-program/<工具>_urls.txt`；已經喺 Excel 任何工作表嘅網址會自動略過。
- 喺面板揀個 txt，撳 **「匯入 Excel…」** 就可以加到某張工作表或者新工作表。主程式運行緊都可以匯入。
- MissAV 會擋自動瀏覽器，所以名單要用 `sub-program/missav_bookmarklet.html` 嘅書籤小程式，喺你平時用嘅瀏覽器下載。

## 中文轉換同香港譯名

### 簡轉繁

次序：OpenCC → 修正字典 → 台灣譯名轉香港譯名（`tw_hk_name.xlsx`）→ AI 校對多義字（可選）→ `customize.txt`。

| 字典 | 位置 | 說明 |
|---|---|---|
| 官方字典 | `opencc_data/dictionary/` | OpenCC 最新正式版，換版前備份到 `opencc_data/dictionary_backup/` |
| 修正字典 | `opencc_overrides/*.txt` | 格式「簡體<Tab>繁體」，同官方同名字典合併，同一個原文用修正版本。想某個詞唔轉換，就寫「原文<Tab>原文」 |
| 日文新字體 | `opencc_overrides/JPShinjitaiCharacters.txt` | 完整字典，唔用官方版本（官方版本好多字同簡體重疊） |
| 最後微調 | `opencc_data/dictionary/customize.txt` | AI 校對之後全文替換，只適合加「喺繁體入面永遠都係錯」嘅詞語 |

### 演員香港譯名

TMDB 冇香港中文名時，程式會按以下次序搵：

1. **Wikidata ID → 中文維基百科**嘅香港譯名（例如 Tom Cruise → 湯·告魯斯；日本、韓國人都適用）
2. 用名字搜尋維基百科（主要用喺日本、韓國人；冇中文維基頁嘅西方人會略過）
3. TMDB 嘅台灣 / 大陸譯名、中文別名，轉成香港用字
4. 日本人全漢字名（TMDB 或者 Wikidata 日文名，新字體轉繁體，例如 平田広明 → 平田廣明）

名字有假名嘅人（例如 綾瀬はるか）冇固定中文寫法，只有維基百科有寫先搵到。

## 設定檔同資料檔

以下檔案**唔會上傳去 git**，冇嘅話程式會自動建立：

| 檔案 | 用途 | 冇嘅時候 |
|---|---|---|
| `configs.json` | 帳號密碼、TMDB API token | 用到時會問；token 要自己加，參考 `configs.example.json` |
| `ai_config.json` | AI 校對設定（provider、model、API key） | 自動建立預設版本；AI 用唔到就只用 OpenCC |
| `data/bulk_videos.xlsx` | bulk 任務表 | 自動建立空白表 |
| `data/tw_hk_name.xlsx` | 台灣譯名 → 香港譯名、受保護人名、AI 校對字詞 | 自動建立範例；改完下次運行先生效 |
| `opencc_data/dictionary/customize.txt` | 最優先嘅自訂字詞替換（左右用 Tab 分隔） | 自動建立範例 |
| `log/` | 運行記錄 | 每次啟動會清空 |

## 常見問題

**撳停止之後顯示「執行緒卡住，強制結束」？**
有工作執行緒冇喺時限內停低，程式會自己關閉佢開嘅瀏覽器再結束。已處理嘅 Done 值會先寫入 Excel，資料唔會冇。log 會記錄卡住嘅位置（`[卡住位置]`）。

**匯入 Excel 失敗，話檔案開住？**
用 Excel 開住 `bulk_videos.xlsx` 時唔可以寫入，先關閉 Excel 再試。

**改咗程式或者設定，幾時生效？**
要重新啟動 `start.py`（喺面板撳停止再啟動）。`tw_hk_name.xlsx` 下次運行先生效。

**想改程式？**
歡迎 fork 自己改。改完請先揀少量項目測試，因為程式會直接修改 TMDB。
