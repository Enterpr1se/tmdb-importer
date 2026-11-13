# TMDB-parallel：AI 交接文件

> 給接手的 AI：先讀完「工作守則」同「已知陷阱」再改任何嘢。
> 用家以廣東話溝通，回覆請用廣東話（繁體）。最後更新：2026-11-11。

## 1. 呢個程式做咩

從影片網站（Netflix、Disney+、Prime Video、Apple TV、viu、viu.tv）抓取劇集 / 電影 / 演員資料，
轉成**香港繁體中文**，再寫入 TMDB（themoviedb.org）嘅 **zh-HK** 翻譯（標題、簡介、分集、人名、劇照）。

- 讀 TMDB：用 API（`configs.json` 嘅 `tmdb_api` 係 **v4 bearer token**，要放 `Authorization` header，唔係 `api_key` 參數）。
- 寫 TMDB：用 Selenium 控制瀏覽器（Edge / msedgedriver）填網頁表格，唔係 API。
- 環境：Windows 11、Microsoft Store 版 Python 3.13（套件裝喺 user site）。用家部機嘅系統 PATH **冇** `C:\Windows\System32`，批次檔要用完整路徑或內建指令。

## 2. 運行流程

```
start.bat → monitor_gui.py（tkinter 監控面板，可代為啟動 start.py）
start.py
  ├─ 檢查/安裝依賴（每日最多升級一次，記錄於 data/.dependency_check.json；--update-deps 強制）
  ├─ 下載 OpenCC 字典、動態建立 opencc_data/config/*.json
  ├─ 問 7 條問題：覆蓋劇集標題/描述、覆蓋分集標題/描述（留空＝同劇集）、
  │   來源無資料時是否清空 TMDB（默認 n）、維基香港譯名是否取代 TMDB 現有香港人名（默認 n）、模式 single/bulk
  └─ tmdb_importer.main()
       ├─ single：循環輸入「影片網址 + TMDB 網址」→ 提取 → 上傳
       └─ bulk：讀 data/bulk_videos.xlsx 的 Bulk 頁
            主執行緒放任務 → extractor_worker（抓取+決策）→ uploader_worker（寫 TMDB）
            → 回寫 Done 值（批次寫入執行緒）
```

### 主要檔案

| 檔案 | 作用 |
|---|---|
| `start.py` | 入口：依賴、OpenCC 設定、啟動問題 |
| `tmdb_importer.py` | 主流程、日誌設定、single/bulk、worker 執行緒、共享快取 |
| `extractors/*_extractor.py` | 各網站抓取；`tmdb_extractor_api.py` 讀 TMDB API 並決定要寫邊啲欄位 |
| `importers/tmdb_uploader.py` | 用瀏覽器填 TMDB 表格（系列、季、分集、電影、人物、劇照） |
| `others/utils.py` | 文字正規化 `normalize_text`、候選值決策、暫停 / 中斷 / 強制退出處理 |
| `others/ai_s2t.py` | AI 簡轉繁校對（見第 3 節） |
| `others/excel_handler.py` | 所有 Excel 讀寫、`tw_hk_name.xlsx` 譯名替換與快取、Done 值批次寫入 |
| `others/owned_procs.py` | 記錄程式自己開嘅瀏覽器 / driver PID，結束時只殺呢啲 |
| `others/chrome_process.py`、`credentials.py` | 瀏覽器啟動、各網站登入 |
| `monitor_gui.py` / `start.bat` | 監控面板：讀 log 同 Excel 顯示進度、啟動/停止/暫停 start.py、AI 被拒用詞清單 |
| `sub-program/` | 獨立小工具（viu.tv 批量更新等），同主程式分開。搵 TMDB 網址嘅工具**只輸出到 `sub-program/<工具>_urls.txt`**（`bulk_txt.py`），唔可以直接寫 `bulk_videos.xlsx`。MissAV 名單由 `missav_bookmarklet.html`（用家瀏覽器書籤小程式）下載，`tmdb_merged_runner.py` 選項 4 讀取。總面板：`tools.bat` → `tools_gui.py`（每個工具開新命令視窗；加新工具要加入 `TOOLS` 清單）。唯一例外：面板「匯入 Excel」（`bulk_txt.import_urls_to_bulk`）由用家手動觸發，用同主程式一樣嘅 `bulk_videos.xlsx.lock`，匯入期間每 3 秒更新 lock 時間（主程式當 15 秒冇更新係過期），openpyxl 加喺工作表底部、暫存檔驗證後取代，改前備份 `data/backup/bulk_videos.before_import.xlsx` |
| `_archive/` | 已停用嘅舊檔，唔好再用 |

### 資料與設定（全部 gitignored）

- `configs.json`：TMDB 帳號、API token 等。`ai_config.json`：AI provider、model、API key。**唔好讀出或外傳內容。**
- `data/bulk_videos.xlsx`：bulk 任務表（Excel 行號 = index + 2）。內有 `old_Bulk` 頁約 24 萬行，令完整載入要約 6.5 秒、儲存約 5 秒。
- `data/tw_hk_name.xlsx`：用家手動維護嘅譯名表，用家會**喺程式運行期間用 Excel 編輯**。分頁：
  - `Default`、`Movie name`：一般譯名（TW Name → HK Name）
  - 關鍵字專用頁（例如 `Gundam`、`One Piece`）：標題/描述含 Keyword 時改用該頁
  - `Protected Name`：人名保護字
  - `AI Convert`：觸發 AI 校對嘅多義字；`AI No Change`：AI 不准改嘅字詞
- `opencc_data/dictionary/customize.txt`：最高優先級自訂替換。
- `data/ai_s2t_cache.db`（SQLite）：AI 回覆快取、`rejected`（被拒字詞）、`accepted`（用家選用 AI 版本）。
- `data/.pause`：暫停旗標，monitor 建立 / 刪除，worker 喺檢查點等待。
- `data/.pause_upload`：「暫停上傳」旗標，只有 bulk uploader 等待，extractor 繼續提取（上傳佇列冇上限）。啟動時兩個旗標都會清除。
- `log/`：運行日誌，**每次 `tmdb_importer` 被 import 時清空**。monitor 靠 log 字眼解析進度。

## 3. 文字轉換次序（`utils.normalize_text`）

OpenCC 字典由 `others/opencc_dict.py` 管理（`start.py` 啟動時呼叫）：
- 官方字典（STPhrases、STCharacters、TSPhrases、TSCharacters、HKVariants）下載自 OpenCC 最新 release，每日最多查一次；刪註解行（`opencc-python-reimplemented` 讀到 `#` 行或空行會出錯）。換版前備份到 `opencc_data/dictionary_backup/`。
- 修正字典 `opencc_overrides/*.txt` 同官方字典**合併**成 `Merged_*.txt`（同 key 取代），設定檔用 Merged 版本。唔好改成「排喺前面嘅獨立字典」：此套件按字典先後配對，短修正會搶走長詞語。
- `JPShinjitaiCharacters.txt` 用 `opencc_overrides/` 嘅完整版本，**唔可以用官方版**：官方版好多字同簡體重疊（来→來），jp2t 喺簡轉繁之前執行，會令 `_needs_simp_to_trad` 誤判為繁體而跳過轉換。
- 此套件唔支援官方 mmseg 分詞，係「最長詞語優先」配對，所以官方短詞條（例如「一出→一齣」）會誤中「一出現」，要用修正字典處理。
- 比較表匯入：`python -m others.opencc_dict import-choices`；即時更新：`python -m others.opencc_dict update`。


1. 人名保護：受保護字詞換成 `@@PROTECTED_n@@`
2. OpenCC jp2t（只用 JPShinjitaiCharacters，套件會忽略 json chain）→ JP 手動字典 → 彎引號轉「」『』（`convert_cjk_quotes`：引號內有中文，或者引號外緊貼中文字 / 中文標點先轉；緊貼英文、數字、空格就保留，例如 `Ebsen “Buddy” Jr.`）
3. 只有偵測為簡體先做：OpenCC **s2t 然後 s2hk**（呢個 opencc 套件唔支援設定檔 chain，一定要分兩步）
4. **AI 校對**（`ai_s2t.refine`）：只送含「AI Convert」觸發字嘅句子
5. TW 字典 → HK 字典 → `customize.txt`
6. 還原保護字詞、NFC、人名去中文間空格
   - 人名保護：name 欄保護「Protected Name」所有字（含單字）；bio 欄只保護 ≥2 字詞，加上呢個人自己嘅名
     （`protect_names`，由 `tmdb_extractor_api._bio_protect_names` 傳入，各種寫法統一寫成最終人名）。
     bio 嘅單字（于、范、里…）只喺呢啲名入面受保護，所以「位于」「范围」會正常轉換。
7. 之後喺 `utils` 候選值決策入面：加中英數空格 → **`tw_hk_name.xlsx` 譯名替換（最後一關）**

AI 校對（relaxed 模式，預設）規則：
- AI 回覆同 OpenCC 草稿逐段比較，接受一般改字 / 換詞。
- 不准改：《》〈〉作品名、`@@PROTECTED_n@@`、受保護人名、`AI No Change`、`Default` / `Movie name` 譯名（≥2 字）。
- 「合法候選字一定接受」（例如《》內 后→後）**唔適用於** `AI No Change` 字詞同受保護人名：連候選字（例如 台→臺）都會還原。
- 單段改動 > 4 字還原；整句改動 > 20% → 整句用 OpenCC。
- 用家喺 monitor 選「以後用 AI 版本」嘅字詞優先，但仍然唔可以改作品名、標記同譯名。
- 快取存 AI 原始回覆，每次重新驗證，所以改規則或清單即時生效，唔使重問 AI。
- AI 失敗會按 `provider_order` 換下一個，全部失敗就用 OpenCC 結果。

## 4. 覆蓋規則（用家定嘅，唔好擅自改）

- 來源網站冇資料而 TMDB 有：**保留 TMDB**，只係簡體先轉繁。只有「來源無資料時清空」= y **同埋**覆蓋 = y 先清空。
- 覆蓋 = n 而 TMDB 有內容但只需正規化（例如引號）：寫入正規化後嘅 TMDB 文字，**唔可以**用來源文字取代。
- 分集覆蓋設定獨立（`utils.EPISODE_OVERWRITE_*`，`None` = 跟劇集設定）。
- 人物（人名/簡介）嘅 force 邏輯刻意未改，用家話唔需要。
- 網站媒體類型同 TMDB 網址類型唔一致（/tv/ vs /movie/）→ 該項當失敗，不上傳。

## 5. 工作守則

- **改變行為前先問用家**，提出方案等批准。用家重視「保留所有原有功能同條件」。純 bug 修正都要講清楚影響。
- **唔好未問就改** `customize.txt`、`tw_hk_name.xlsx`、`configs.json`、`ai_config.json`。程式本身對 `tw_hk_name.xlsx` 只讀。
- **只關閉程式自己開嘅瀏覽器**（`owned_procs`），唔可以殺用家其他 Edge/Chrome。
- Commit：只 commit 自己改嘅檔案，用家要求先 commit / push。`sub-program/tmdb_contributed_links.txt` 係用家自己嘅檔案，唔好一齊 commit。
- 檔案係 **CRLF**，用編輯工具改，唔好用 shell 重寫整個檔案。
- 修改前可另存備份（`_backup_YYYYMMDD_*/`，已 gitignore）。

## 6. 已知陷阱

- **import `tmdb_importer` 會清空 `./log`**（`setup_loggers` 喺 import 時執行）。測試要喺暫存資料夾做 cwd，唔好喺專案目錄。
- `setup_loggers` 入面，`handlers.clear()` 同 `addHandler` 之間**唔可以 log**，否則會寫去已關閉嘅 stderr。
- `importer.log` 有 `ThreadNameFilter`，只收指定執行緒名稱。**新增具名執行緒要加入該清單**，否則只出 console。
- `ai_s2t._load_cache()` 持有 `_state.lock` 時會叫 `_load_config()`（普通 Lock）→ 一定要**先**呼叫 `_load_config()`，否則自鎖。
- monitor 用正則解析 log 字眼（例如「[Extractor] 開始處理任務 (索引: N)」、「已將 Done 值更新 (row_index=N」）同自動回答啟動問題（按問題文字比對）。**改 log 字眼或 start 問題文字，要同步改 `monitor_gui.py`**（`RunState` 正則、`AUTO_ANSWERS`、`RUN_OPTIONS`）。
- monitor 用隱藏 console 啟動 start.py（stdin 用 pipe）。Ctrl+C 靠 Windows console ctrl handler，唔係 Python signal。
- Edge 157 起：由新分頁第一次導航後，`window_handles` 會多咗一個隱藏嘅 `edge://newtab` handle（排最後）。切去佢再導航會出
  `no such window: target window already closed / web view not found`。`chrome_process.ensure_browser_window` 一定要**先用目前視窗**，唔好揀最後一個 handle。
- 人物香港譯名（`tmdb_extractor_api.get_api_data_for_item`）：西方名（`_is_western_person_name`：中文譯名有「·」/「-」或多過 4 字，或冇中文名）
  先用 TMDB `external_ids.wikidata_id` → `wiki_search.lookup_hk_name_by_wikidata`（Wikidata 中文維基頁名 → `action=parse&variant=zh-hk` 顯示標題，
  會套用 NoteTA 香港譯名）。經 `wiki_queue` 同一條隊。快取 `wiki_cache.db` media_type=`person_qid`，NOT_FOUND 90 日後重查。
  中日韓名照舊（有中文名就唔搜維基）。`?variant=zh-hk` 嘅 `varianttitles` API **唔會**套用 NoteTA，唔好用。
  所有維基請求用 `wiki_search.WIKI_USER_AGENT`（含項目網址），Wikimedia 要求 User-Agent 有聯絡方式。
  啟動選項「維基香港譯名取代 TMDB 現有香港人名」（`utils.WIKI_NAME_REPLACES_EXISTING`，默認 n）。
  舊嘅後備名稱搜尋（`search_wiki_for_hk_name`）命中率極低（2026-11 實測 318 次中 1 次，每次排隊 ~1 分鐘），所以以下情況跳過：
  選項 = n 而 TMDB 已有中文香港人名；或者 Wikidata 快取已確認冇中文維基頁而又唔係日韓人（`_is_jk_person`：假名 / 韓文 / 漢字原名 / 出生地日韓）。
- 簡介換人名（`replace_name_in_bio`）：完整舊名（TW / CN / TMDB 現有 HK / 原名）→ 新名；之後 `_replace_name_parts` 再換部份名
  （姓或名，≥2 字，部份數唔同就只換姓），避開《》作品名同已經係新全名嘅位置。
- bulk 工作表：啟動時揀 `bulk_videos.xlsx` 邊幾張工作表，逗號分隔、按次序處理（名稱或編號，唔分大細楷；
  默認 / 全部無效 = `Bulk`，無效嘅會略過）。`utils.BULK_SHEETS` = 次序，`utils.BULK_SHEET` = 而家處理緊嗰張；
  `excel_handler._bulk_sheet()` 同 `tmdb_importer` 所有讀寫（共享快取、重新掃描、去重、Done 批次寫入）都用 `BULK_SHEET`。
  一張做完（本輪冇新任務、全部閒置）先 `_switch_bulk_sheet`：等 Done 全部寫入 → 轉 → 去重 → 清 `details_cache` → 重讀共享快取。
  因為 Done 同 `details_{行號}.xlsx` 都用行號，絕對唔可以有任務做緊時轉。共享快取讀取期間轉咗工作表會棄用舊結果。
  log `[工作表] 開始處理工作表 'X'（第 i/n 張）` 畀監控面板跟住切換 Excel 進度（`RE_SHEET`）。
  監控面板「bulk 工作表」用「選擇…」對話框揀同排次序。
- 每行設定欄位（`excel_handler.BULK_OPTION_COLUMNS`，Done 後面第 5–10 欄）：覆蓋劇集標題 / 覆蓋劇集描述 / 覆蓋分集標題 /
  覆蓋分集描述（y/n/ask）、來源無資料時清空TMDB / 維基人名取代現有（y/n）。空白 = 程式選項；有值只影響嗰行。
  `check_create_bulk_excel` 會為所有 bulk 格式工作表（表頭頭兩欄 Video URL / TMDB URL）補欄（openpyxl，保留格式）。
  實現：`utils.ROW_OPTIONS`（ContextVar），extractor 讀行時 `parse_row_options` 並 set，經 upload_queue `row_options` 傳畀 uploader；
  深層用 `utils.wiki_name_replaces_existing()` / `clear_tmdb_when_source_empty()` / `episode_overwrite()`，**唔好直接讀**
  `WIKI_NAME_REPLACES_EXISTING` / `CLEAR_TMDB_WHEN_SOURCE_EMPTY` 全域值。分集次序：該行分集欄 → 程式分集選項 → 劇集設定（含該行劇集欄）。
- 空行（`_is_blank_bulk_row`：Video URL 同 TMDB URL 都冇內容）：整理 Excel（`normalize_and_deduplicate_urls`，開始時同轉工作表時）會刪除，
  下面嘅行上移；運行中途出現嘅空行唔會入隊（以前會寫 Done=0）。中途唔可以刪行，因為行號會變。
  `sub-program/*` 加網址嘅工具唔寫 Excel，只輸出 `sub-program/<工具>_urls.txt`。
- `EXCEL_LOCK` 爭用訊息係程式內部鎖，唔係 Excel 本身鎖檔。
- bulk 中途可以喺 Excel **底部**加行（以 index 對應），唔好喺中間插入或重排。
- Netflix OTP 登入後會自動跳去 /browse，要確認頁面真係該作品詳情頁先抓取。
- 簡體偵測（`_needs_simp_to_trad`）只認「繁體冇用嘅簡體字」（`_trad_only_chars` 由 STCharacters 計），含「台」「于」「干」嘅純繁體唔會再誤判（2026-11-12 驗證）。
  只有偵測為簡體先會做 s2t / s2hk / AI 校對，所以影片網站官方繁體字幕 / 簡介唔會經 AI。

## 7. 測試方法

- 冇正式測試套件。做法：喺暫存資料夾寫小腳本，`sys.path.insert(0, r'C:\TMDB-parallel')`，cwd 設為暫存資料夾（需要時複製 `data/tw_hk_name.xlsx` 副本過去），直接呼叫函數驗證。
- AI 單句測試（喺專案根目錄）：`python -m others.ai_s2t test 几天后她回来了`；`scan` / `qualify` 用嚟揀 OpenRouter 免費 model。
- 真實上傳 TMDB 係對外寫入，未得用家同意唔好跑。
