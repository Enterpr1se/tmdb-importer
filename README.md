# tmdb-importer

從影片網站抓取劇集、電影、演員資料，轉成**香港繁體中文**，再寫入 [TMDB](https://www.themoviedb.org/) 嘅 zh-HK 翻譯。

支援網站：Netflix、Disney+、Prime Video、Apple TV、viu、viu.tv。

> 呢個係個人自用工具，會用你嘅 TMDB 帳戶直接修改 TMDB 上面嘅資料。正式運行前，請先用少量項目測試。

## 功能

- **寫入內容**：劇集或電影嘅標題同簡介、季同分集標題同簡介、演員名同簡介、劇照。
- **簡轉繁**：用 OpenCC 轉成香港用字，加上自訂修正字典同台灣譯名轉香港譯名表。可以再用 AI 校對多義字。
- **香港人名**：TMDB 冇香港譯名時，會用 Wikidata / 維基百科嘅香港譯名。
- **兩種模式**：
  - single：逐個輸入影片網址同 TMDB 網址。
  - bulk：處理 Excel 任務表。可以一次揀幾張工作表，按次序處理，抓取同上傳同時進行。
- **監控面板**：睇進度，啟動、停止、暫停程式，或者只暫停上傳。

## 需要

- Windows 10 / 11
- Python 3（喺 Microsoft Store 版 Python 3.13 測試過）。用安裝程式安裝時要勾選「Add python.exe to PATH」。
- Microsoft Edge
- TMDB 帳戶。建議另外申請 TMDB API Read Access Token。

其他 Python 套件同 OpenCC 官方字典，`start.py` 會自動下載安裝，之後每日最多檢查一次更新。
想即時檢查就執行 `python start.py --update-deps`。

## 開始使用

1. 雙擊 **`start.bat`** 開監控面板，揀好選項後按「▶ 啟動」。
   你亦都可以喺命令視窗執行 `python start.py`。
2. 揀覆蓋設定，同揀 single 或 bulk 模式。
3. 第一次用到 TMDB、Netflix、Disney+ 時，程式會問帳號密碼，可以揀儲存。

### 啟動選項

| 選項 | 意思 |
|---|---|
| 覆蓋劇集標題 / 描述 | TMDB 已經有香港中文時，係咪用新資料取代（y / n / ask） |
| 覆蓋分集標題 / 描述 | 同上，用喺分集。留空就跟劇集設定 |
| 來源無資料時清空 TMDB | 影片網站冇呢項資料時，係咪清空 TMDB 現有內容（默認 n） |
| 維基人名取代現有 | 維基香港譯名係咪取代 TMDB 現有嘅香港人名（默認 n） |
| bulk 工作表 | bulk 模式要處理 `bulk_videos.xlsx` 邊幾張工作表（默認 Bulk） |

## bulk 任務表（`data/bulk_videos.xlsx`）

每張工作表一行一個任務：

| 欄位 | 內容 |
|---|---|
| Video URL | 影片網站網址 |
| TMDB URL | 對應嘅 TMDB 網址 |
| Split | Netflix 分集標題有「/」時，係咪拆成兩集（y / n）。平時留空 |
| Done | 程式處理完會自動填寫 |
| 覆蓋劇集標題、覆蓋劇集描述、覆蓋分集標題、覆蓋分集描述 | 只影響呢一行：y / n / ask。留空就用啟動選項 |
| 來源無資料時清空TMDB、維基人名取代現有 | 只影響呢一行：y / n。留空就用啟動選項 |

- 檔案冇嘅話會自動建立。舊檔案缺少欄位，開始 bulk 時會自動加上。
- 重複網址同空行會喺開始時自動刪除。
- 程式運行期間都可以加新行，或者改未處理嘅行。

## 簡轉繁字典

- **官方字典**：OpenCC 最新正式版，放喺 `opencc_data/dictionary/`。換版前會備份到 `opencc_data/dictionary_backup/`。
- **修正字典**：`opencc_overrides/*.txt`，格式係「簡體<Tab>繁體」，會同官方同名字典合併，同一個原文用修正版本。
  想某個詞唔轉換，就寫「原文<Tab>原文」，例如 `一出	一出`。
- **日文新字體字典**：`opencc_overrides/JPShinjitaiCharacters.txt` 係完整字典，唔用官方版本，因為官方版本好多字同簡體中文重疊。
- **最後微調**：`opencc_data/dictionary/customize.txt`，喺 AI 校對之後全文替換，只適合加「喺繁體入面永遠都係錯」嘅詞語。

## 設定檔同資料檔

以下檔案唔會上傳去 git，冇嘅話程式會自動建立：

| 檔案 | 用途 | 冇嘅時候 |
|---|---|---|
| `configs.json` | 帳號密碼、TMDB API token | 用到時會問；API token 要自己加，參考 `configs.example.json` |
| `ai_config.json` | AI 校對設定（provider、model、API key） | 自動建立預設版本；AI 用唔到就只用 OpenCC |
| `data/bulk_videos.xlsx` | bulk 任務表 | 自動建立空白表 |
| `data/tw_hk_name.xlsx` | 台灣譯名 → 香港譯名、受保護人名、AI 校對字詞 | 自動建立範例；改完下次運行先生效 |
| `opencc_data/dictionary/customize.txt` | 最優先嘅自訂字詞替換（左右用 Tab 分隔） | 自動建立範例 |

## 其他

- `sub-program/` 入面係獨立小工具，同主程式分開運行。
- 開發同修改須知見 `AI_HANDOFF.md`。
