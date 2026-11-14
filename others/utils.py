import contextvars
import langid
import logging
import opencc
import unicodedata
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from others.analysis_excel import is_generic_title, clean_generic_title
from others.excel_handler import clear_tw_hk_name_cache
import time
import random
import re
import os
import openpyxl # 【新增】導入 openpyxl 以便安全地檢查 sheet
import pandas as pd # 【新增】導入 pandas 以便讀取 sheet
import traceback # 【新增】導入 traceback 以便在日誌中顯示更詳細的錯誤

"""
graph TD
    subgraph "D. 文字正規化管線 (utils.normalize_text)"
        direction LR
        D_IN[輸入文字] --> D_CHECK{是人名字段嗎?};
        D_CHECK -- 是 --> D_PROTECT[保護特殊人名<br>(從 Protected Name sheet 載入)];
        D_CHECK -- 否 --> D1;
        D_PROTECT --> D1[1. OpenCC: 日文新字體 -> 繁體<br>(jp2t.json)];
        D1 --> D2[2. 手動字典: JP*.txt];
        D2 --> D3[3. OpenCC: 簡體 -> 繁體<br>(s2t.json)];
        D3 --> D4[4. 手動字典: TW*.txt];
        D4 --> D5[5. OpenCC: 簡體 -> 香港繁體<br>(s2hk.json)];
        D5 --> D6[6. 手動字典: HK*.txt];
        D6 --> D7[7. <b>最高優先級</b>: 自定義字典<br>(customize.txt)];
        D7 --> D_RESTORE_CHECK{是人名字段嗎?};
        D_RESTORE_CHECK -- 是 --> D_RESTORE[還原被保護的特殊人名];
        D_RESTORE_CHECK -- 否 --> D8;
        D_RESTORE --> D8;
        D8 --> D9[8. Unicode 正規化 (NFC)];
        D9 --> D10[9. (人名專用) 移除中文間空格];
        D10 --> D_OUT[輸出標準化文字];
    end
"""

# 【新增】【全局狀態】程序中斷標誌，從 tmdb_importer.py 移至此處以解決循環導入問題。
_PROGRAM_INTERRUPTED = False

# 【2026-11】監控面板「暫停」：旗標檔存在時，做完手上項目後不開始新項目（未處理的留在佇列）。
PAUSE_FLAG = os.path.join('data', '.pause')
# 「暫停上傳」：只有 bulk uploader 停（做完手上項目後不取新任務），extractor 照樣提取，結果留在上傳佇列。
UPLOAD_PAUSE_FLAG = os.path.join('data', '.pause_upload')


def clear_pause_flag():
    """程式啟動時呼叫：上次暫停中結束留下的旗標不應影響今次運行。"""
    for flag in (PAUSE_FLAG, UPLOAD_PAUSE_FLAG):
        try:
            os.remove(flag)
        except OSError:
            pass


def wait_if_paused(logger, label, flags=(PAUSE_FLAG,)):
    """任何一個旗標存在就在此等待，直至面板按「繼續」或程式被中斷（Ctrl+C 不受暫停影響）。"""
    def paused():
        return any(os.path.exists(f) for f in flags)
    if not paused():
        return
    logger.info(f"[{label}] 已暫停，等待繼續...")
    print(f"[{label}] 已暫停，等待繼續...")
    while paused() and not _PROGRAM_INTERRUPTED:
        time.sleep(0.5)
    if not _PROGRAM_INTERRUPTED:
        logger.info(f"[{label}] 已繼續。")
        print(f"[{label}] 已繼續。")


# 【2026-11 N1】來源（影片網站 / Excel）沒有標題或描述、而 TMDB 已有內容時：
#   False（預設）= 保留 TMDB 內容；只在 TMDB 內容是簡體時寫入轉換後的繁體。
#   True        = 舊行為：覆蓋設為 y 時清空 TMDB 欄位（讓 TMDB 顯示預設語言）。
# 由 start.py 的啟動問題「來源無資料時是否清空 TMDB」設定（tmdb_importer.main 參數 clear_when_source_empty）。
CLEAR_TMDB_WHEN_SOURCE_EMPTY = False

# 【2026-11】維基香港譯名是否取代 TMDB 現有（中文）香港人名。
#   False（預設）= TMDB 已有中文 zh-HK 人名就保留；維基譯名只用於 TMDB 冇香港人名時。
#   True        = 維基譯名優先（例如 讓·雷諾 → 尚·連奴）。
# 由 start.py 啟動問題「是否用維基香港譯名取代 TMDB 現有香港人名」設定（tmdb_importer.main 參數 wiki_name_replace）。
WIKI_NAME_REPLACES_EXISTING = False

# 【2026-11】bulk 模式處理 data/bulk_videos.xlsx 入面邊幾張工作表、按咩次序（默認只有 'Bulk'）。
# 由 start.py 啟動問題「bulk 模式要處理 bulk_videos.xlsx 邊一張工作表」設定（tmdb_importer.main 參數 bulk_sheet，
# 逗號分隔，例如 "Bulk, Eng" 或 "1, 13"）。BULK_SHEETS = 次序；BULK_SHEET = 而家處理緊嗰張。
# 一張做完（所有任務完成、Done 全部寫入）先會轉下一張，因為 Done 值同 details_{行號}.xlsx 都係用行號對應。
DEFAULT_BULK_SHEET = 'Bulk'
BULK_SHEETS = [DEFAULT_BULK_SHEET]
BULK_SHEET = DEFAULT_BULK_SHEET

# 【2026-11 P2】分集 (episode) 的覆蓋選項，可與劇集（系列/季/電影）分開設定。
# None = 沿用劇集設定。由 tmdb_importer.main 參數 episode_overwrite_titles / episode_overwrite_descriptions 設定。
EPISODE_OVERWRITE_TITLES = None
EPISODE_OVERWRITE_DESCRIPTIONS = None


# 【2026-11】bulk_videos.xlsx 每行設定（欄位見 excel_handler.BULK_OPTION_COLUMNS）。
# 用 ContextVar：extractor / uploader 每條 thread 處理每行之前 set 一次，互不影響；
# asyncio.run / asyncio.to_thread 會自動帶埋。冇設定 = {} = 用程式選項。
ROW_OPTIONS = contextvars.ContextVar('bulk_row_options', default={})
_ROW_OPTION_VALUES = {
    'overwrite_titles': ('y', 'n', 'ask'), 'overwrite_descriptions': ('y', 'n', 'ask'),
    'episode_overwrite_titles': ('y', 'n', 'ask'), 'episode_overwrite_descriptions': ('y', 'n', 'ask'),
    'clear_when_source_empty': ('y', 'n'), 'wiki_name_replace': ('y', 'n'),
}


def parse_row_options(row):
    """由 bulk 表一行讀取設定欄位 → {key: 'y'/'n'/'ask'}；空白或無效值唔計（用程式選項）。"""
    from others.excel_handler import BULK_OPTION_COLUMNS
    opts = {}
    for col, key in BULK_OPTION_COLUMNS.items():
        try:
            v = row.get(col)
        except Exception:
            v = None
        if v is None or (isinstance(v, float) and v != v):
            continue
        v = str(v).strip().lower()
        v = {'yes': 'y', 'no': 'n', '是': 'y', '否': 'n'}.get(v, v)
        if v in _ROW_OPTION_VALUES[key]:
            opts[key] = v
        elif v not in ('', 'nan', 'none'):
            logging.warning(f"bulk 欄位「{col}」值 '{v}' 無效（可用 {'/'.join(_ROW_OPTION_VALUES[key])}），用返程式選項。")
    return opts


def row_option(key):
    return ROW_OPTIONS.get().get(key)


def wiki_name_replaces_existing():
    """維基香港譯名係咪取代 TMDB 現有香港人名：該行欄位優先，否則程式選項。"""
    v = row_option('wiki_name_replace')
    return v == 'y' if v else WIKI_NAME_REPLACES_EXISTING


def clear_tmdb_when_source_empty():
    """來源無資料時係咪清空 TMDB：該行欄位優先，否則程式選項。"""
    v = row_option('clear_when_source_empty')
    return v == 'y' if v else CLEAR_TMDB_WHEN_SOURCE_EMPTY


def episode_overwrite(kind, fallback):
    """回傳分集的覆蓋選項 (kind = 'title' / 'description')。次序：該行分集欄位 → 程式分集選項 → fallback（劇集設定）。"""
    value = row_option('episode_overwrite_titles' if kind == 'title' else 'episode_overwrite_descriptions')
    if not value:
        value = EPISODE_OVERWRITE_TITLES if kind == 'title' else EPISODE_OVERWRITE_DESCRIPTIONS
    value = str(value or '').strip().lower()
    return value if value in ('y', 'n', 'ask') else fallback

def signal_handler(sig, frame):
    """
    【新增】處理 CTRL+C 信號，設定全局中斷標誌。
    此函數從 tmdb_importer.py 移至此處，以避免循環導入。
    """
    global _PROGRAM_INTERRUPTED
    if not _PROGRAM_INTERRUPTED: # 防止重複日誌/操作
        logging.info("檢測到 Ctrl+C，設置全局中斷標誌。程序將在適當時機終止。")
        _PROGRAM_INTERRUPTED = True
        if not _win_ctrl_handler_installed:  # Windows handler 已自行顯示提示及處理
            _console_write(_FIRST_PRESS_HINT)
            _on_first_interrupt()
    elif not _win_ctrl_handler_installed:
        # 非 Windows（或未能安裝 Windows handler）：第二次 Ctrl+C 在此強制結束
        _force_exit()


# ---------------------------------------------------------------------------
# 【2026-11】第二次 Ctrl+C → 立即強制結束
# 第一次 Ctrl+C 維持原有「安全停止」（設定 _PROGRAM_INTERRUPTED，各執行緒在檢查點自行結束）。
# 但如果主執行緒正卡在 Selenium / 網絡 / lock 等待，Python 的 signal handler 根本不會執行，
# 所以在 Windows 另外用 SetConsoleCtrlHandler 註冊系統層 handler：它由 Windows 在獨立執行緒呼叫，
# 即使主執行緒卡住，第二次 Ctrl+C 都能結束程式。
# ---------------------------------------------------------------------------
_win_ctrl_handler_installed = False
_win_ctrl_handler_ref = None   # 必須保留引用，否則 ctypes callback 會被回收
_ctrl_c_presses = 0
_FIRST_PRESS_HINT = (
    "\n[Ctrl+C] 正在安全停止（完成目前步驟並儲存，之後自動關閉本程式開啟的瀏覽器）..."
    "再按一次 Ctrl+C 立即強制結束。\n"
)

# ---------------------------------------------------------------------------
# 【2026-11 L1】按一次 Ctrl+C 就自動安全停止並完整關閉（包括瀏覽器）
# - 主執行緒正在等使用者輸入（interruptible_input）：沒有工作進行中，立即執行收尾並結束。
#   （監控面板以 pipe 提供 stdin，input() 會一直阻塞，Python 的 signal handler 根本沒機會執行。）
# - 其他情況：維持原有安全停止；另開 watchdog，SAFE_STOP_TIMEOUT 秒後仍未結束就收尾並結束。
#   需要等使用者處理的步驟（例如「請關閉 Excel 後按 Enter」）用 hold_watchdog() 暫停倒數。
# 收尾 = 執行 register_shutdown_hook 登記的函數（例如寫入 Done 值）+ 關閉本程式開啟的瀏覽器。
# ---------------------------------------------------------------------------
SAFE_STOP_TIMEOUT = 60
_input_waiting = 0
_watchdog_hold = 0
_watchdog_started = False
_graceful_exit_started = False
_shutdown_hooks = []
_state_lock = __import__('threading').Lock()


def register_shutdown_hook(fn, timeout=30):
    """登記收尾函數（Ctrl+C 自動結束前執行，每個最多 timeout 秒）。"""
    _shutdown_hooks.append((fn, timeout))


def interruptible_input(prompt=''):
    """取代 input()：等待期間按 Ctrl+C 會立即安全結束程式（無工作進行中）。"""
    global _input_waiting
    if _PROGRAM_INTERRUPTED:
        _graceful_exit_now("已收到停止指令")
    with _state_lock:
        _input_waiting += 1
    try:
        return input(prompt)
    finally:
        with _state_lock:
            _input_waiting -= 1


class hold_watchdog:
    """with hold_watchdog(): 期間暫停自動結束倒數（例如等使用者關閉 Excel）。"""
    def __enter__(self):
        global _watchdog_hold
        with _state_lock:
            _watchdog_hold += 1
        return self

    def __exit__(self, *exc):
        global _watchdog_hold
        with _state_lock:
            _watchdog_hold -= 1
        return False


def _on_first_interrupt():
    """第一次 Ctrl+C（可在任何執行緒呼叫）。"""
    import threading as _th
    global _watchdog_started
    with _state_lock:
        waiting = _input_waiting > 0
        start_watchdog = not waiting and not _watchdog_started
        if start_watchdog:
            _watchdog_started = True
    if waiting:
        _th.Thread(target=_graceful_exit_now, args=("等待輸入時收到 Ctrl+C",),
                   name="GracefulExit", daemon=True).start()
    elif start_watchdog:
        _th.Thread(target=_safe_stop_watchdog, name="SafeStopWatchdog", daemon=True).start()


def _safe_stop_watchdog():
    import time as _t
    remaining = SAFE_STOP_TIMEOUT
    while remaining > 0:
        _t.sleep(1)
        with _state_lock:
            held = _watchdog_hold > 0
            waiting = _input_waiting > 0
        if waiting:  # 安全停止途中又回到等待輸入 → 不必再等
            break
        if held:
            remaining = SAFE_STOP_TIMEOUT
        else:
            remaining -= 1
    reason = "停止途中回到等待輸入" if remaining > 0 else f"安全停止 {SAFE_STOP_TIMEOUT} 秒後仍未結束"
    _graceful_exit_now(reason)


def _graceful_exit_now(reason):
    """執行收尾函數、關閉本程式開啟的瀏覽器，然後結束程式（只執行一次）。"""
    global _graceful_exit_started
    import threading as _th
    with _state_lock:
        if _graceful_exit_started:
            return
        _graceful_exit_started = True
    _console_write(f"\n[結束] {reason}：正在收尾並關閉本程式開啟的瀏覽器...\n")
    for fn, timeout in list(_shutdown_hooks):
        t = _th.Thread(target=_run_quietly, args=(fn,), daemon=True)
        t.start()
        t.join(timeout=timeout)
    # 先正常關閉瀏覽器（driver.quit，有逾時），_force_exit 再強制結束仍殘留的行程
    try:
        from others import owned_procs
        owned_procs.kill_owned(quit_drivers=True)
    except Exception:
        pass
    try:
        logging.info(f"[結束] {reason}，已完成收尾，程式結束。")
        for h in logging.getLogger().handlers:
            try:
                h.flush()
            except Exception:
                pass
    except Exception:
        pass
    _force_exit(130, message="[結束] 已關閉瀏覽器，程式結束。\n")


def _run_quietly(fn):
    try:
        fn()
    except Exception as e:
        _console_write(f"[結束] 收尾步驟出錯（略過）: {e}\n")


def _console_write(msg):
    """不經 logging（避免卡住的執行緒持有 logging lock 時死鎖），直接寫到 stderr。"""
    try:
        os.write(2, msg.encode('utf-8', errors='replace'))
    except Exception:
        pass


def _force_exit(exit_code=130, message="\n[Ctrl+C] 第二次 Ctrl+C：正在強制結束程式及其瀏覽器...\n"):
    """結束本程式啟動的所有子行程（chromedriver / 瀏覽器等），然後立即結束。
    只處理本程式自己的子孫行程，不會影響使用者自行開啟的瀏覽器。"""
    import threading as _th
    _console_write(message)

    def _kill_children():
        try:
            import psutil
            for child in psutil.Process().children(recursive=True):
                try:
                    child.kill()
                except Exception:
                    pass
        except Exception:
            pass

    def _kill_all():
        # 【2026-11 L2】先關閉啟動時記錄的瀏覽器（即使 driver 已先被 Ctrl+C 結束、瀏覽器失去父行程）
        try:
            from others import owned_procs
            owned_procs.kill_owned(quit_drivers=False)
        except Exception:
            pass
        _kill_children()

    t = _th.Thread(target=_kill_all, daemon=True)
    t.start()
    t.join(timeout=8)  # 最多等 8 秒，之後無論如何都結束
    os._exit(exit_code)


def _log_stuck_thread_stacks(stuck_threads):
    """【2026-11】強制結束前，將卡住嘅執行緒停喺邊一行寫入 log，方便之後搵原因。"""
    try:
        import sys as _sys
        import traceback as _tb
        frames = _sys._current_frames()
        for t in stuck_threads:
            frame = frames.get(t.ident)
            if frame is None:
                continue
            stack = ''.join(_tb.format_stack(frame)[-15:])   # 最尾 15 層已足夠
            logging.warning(f"[卡住位置] 執行緒 {t.name} 目前停喺:\n{stack}")
    except Exception as e:
        try:
            logging.warning(f"[卡住位置] 無法取得執行緒堆疊: {e}")
        except Exception:
            pass


def exit_if_threads_stuck(threads):
    """【2026-11 F1】工作執行緒 join 逾時後仍未結束時使用。
    若 threads 中仍有非 daemon 執行緒存活，Python 結束時會無限等待它們（程式及其瀏覽器殘留），
    所以在此記錄卡住的執行緒，關閉本程式自己開啟的 driver / 瀏覽器（子孫行程），然後結束程式。
    所有執行緒都已正常結束時不做任何事。"""
    stuck_threads = [t for t in threads if t is not None and t.is_alive() and not t.daemon]
    stuck = [t.name for t in stuck_threads]
    if not stuck:
        return
    try:
        logging.warning(f"以下工作執行緒逾時仍未結束: {', '.join(stuck)}；"
                        f"將關閉本程式開啟的瀏覽器並結束程式。")
    except Exception:
        pass
    _log_stuck_thread_stacks(stuck_threads)
    _force_exit(130 if _PROGRAM_INTERRUPTED else 1,
                message=f"\n[結束] 執行緒卡住未能結束 ({', '.join(stuck)})，"
                        f"正在關閉本程式開啟的瀏覽器並結束程式...\n")


def install_force_exit_handler():
    """在 signal.signal(SIGINT, signal_handler) 之後呼叫。"""
    global _win_ctrl_handler_installed, _win_ctrl_handler_ref
    if os.name != 'nt' or _win_ctrl_handler_installed:
        return
    try:
        import ctypes
        from ctypes import wintypes

        HandlerRoutine = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.DWORD)

        def _handler(ctrl_type):
            global _ctrl_c_presses, _PROGRAM_INTERRUPTED
            if ctrl_type in (0, 1):  # CTRL_C_EVENT, CTRL_BREAK_EVENT
                _ctrl_c_presses += 1
                if _ctrl_c_presses >= 2:
                    _force_exit()
                _console_write(_FIRST_PRESS_HINT)
                # 【2026-11 L1】Windows 在獨立執行緒呼叫此 handler：主執行緒卡在 input() 時亦能處理
                _PROGRAM_INTERRUPTED = True
                _on_first_interrupt()
            # 回傳 False：交給 Python 原本的處理（→ signal_handler 設定安全停止標誌）
            return False

        _win_ctrl_handler_ref = HandlerRoutine(_handler)
        if ctypes.windll.kernel32.SetConsoleCtrlHandler(_win_ctrl_handler_ref, True):
            _win_ctrl_handler_installed = True
    except Exception as e:
        logging.debug(f"安裝 Windows Ctrl+C handler 失敗，改用 Python signal 處理: {e}")

# --- 【修改】實現新的轉換流程與字典載入機制 ---

# 全局變數，用於儲存所有轉換器和手動載入的字典
manual_dictionary_maps = {
    'JP': {},
    'TW': {},
    'HK': {}
}
customize_dictionary_map = {}
opencc_converters = {}
# 【使用者要求】新增全局變數，用於儲存受保護的人名
protected_person_names = set()

# 緩存語言檢測結果，減少重複日誌
_language_cache = {}
_program_interrupted_callback_for_utils = lambda: False # Default callback

# 不需要的描述文本
NO_DESCRIPTION_TEXT = [
    "目前沒有相關簡介，您可協助新增簡介資料。",
    "There are no episodes added to this season.",
    "暂无英文的剧情简介，請幫助我们添加到数据库。",
    "沒有翻譯為正體中文的概要，請協助我們新增資料。",
    "剧情简介缺失，請幫助我们更新。",
    "暫無簡體中文的劇情簡介，請幫助添加到數據庫。",
    "暂无简体中文的剧情简介，请帮助添加到数据库。",
    "暫無",
    "暂无"
]

def initialize_opencc_converters():
    """
    【修改】初始化所有需要的 OpenCC 轉換器。
    此函數現在會直接載入由 start.py 動態建立的 custom_*.json 設定檔。

    【核心修正】為了解決 .json.json 的 FileNotFoundError，並確保程式在任何路徑下都能正確運行，
    我們採用了以下兩個關鍵步驟：
    1.  取得設定檔的絕對路徑 (os.path.abspath)：
        這可以防止 opencc 套件在其自己的安裝目錄（site-packages/opencc/config）中尋找設定檔，
        而是強制它使用我們在專案資料夾中動態建立的設定檔（例如 opencc_data/config/custom_s2hk.json）。
        這是實現路徑獨立性（Portability）的關鍵。
    2.  移除 .json 副檔名 (os.path.splitext)：
        opencc-python-reimplemented 套件的建構函式 opencc.OpenCC() 會自動在其接收到的路徑後面
        加上 ".json"。如果我們提供一個已經包含 ".json" 的路徑，就會變成 ".json.json"，
        導致檔案找不到的錯誤。因此，在傳遞路徑之前，我們必須先將副檔名移除。

    這個方法是目前最穩健、最可靠的作法。
    """
    global opencc_converters, _TRAD_ONLY_CHARS
    import opencc
    import os
    import traceback # 導入 traceback 以便顯示更詳細的錯誤
    _TRAD_ONLY_CHARS = None   # 字典可能已更新

    # 指向由 start.py 建立的設定檔目錄
    config_dir = os.path.join('opencc_data', 'config')

    # 【關鍵】定義要載入的自定義設定檔，名稱與 start.py 中建立的檔案一致
    configs = {
        'jp2t': 'custom_jp2t.json',
        's2t': 'custom_s2t.json',
        's2hk': 'custom_s2hk.json',
        's2t_check': 'custom_s2t.json',  # 用於簡體檢查
        # 【2026-11】已停用繁轉簡字典（t2s_check），繁體判斷改用 _trad_only_chars()
    }
    
    for name, config_filename in configs.items():
        try:
            # 步驟 1: 建立設定檔的完整路徑
            config_path = os.path.join(config_dir, config_filename)
            
            # 檢查設定檔是否存在，如果不存在，則可能是 start.py 的設定步驟未成功
            if not os.path.exists(config_path):
                logging.error(f"預期的 OpenCC 設定檔不存在: {config_path}。請確保 start.py 已成功執行且 'opencc_data' 資料夾位於正確位置。")
                opencc_converters[name] = None
                continue

            # 【核心修正步驟 1】將路徑轉換為絕對路徑，確保路徑的唯一性和準確性
            absolute_config_path = os.path.abspath(config_path)

            # 【核心修正步驟 2】使用 os.path.splitext() 安全地移除 .json 副檔名
            path_without_extension, _ = os.path.splitext(absolute_config_path)

            # 將處理過的、不含副檔名的絕對路徑傳遞給 OpenCC
            opencc_converters[name] = opencc.OpenCC(path_without_extension)
            logging.info(f"成功初始化 OpenCC 轉換器 '{name}'，使用自定義設定檔: '{config_filename}'。")

        except Exception as e:
            opencc_converters[name] = None
            # 記錄更詳細的錯誤訊息，方便除錯
            logging.error(f"初始化 OpenCC 轉換器 '{name}' 失敗 (設定檔: {config_filename})。錯誤: {e}\n{traceback.format_exc()}")

def load_manual_dictionaries(dict_dir='opencc_data/dictionary'):
    """
    【修改】載入指定資料夾中的所有 .txt 字典檔，並按前綴分類。
    【修正】跳過 OpenCC 的核心系統字典檔，避免因「一對多」的空格格式導致「吃吃 吃吃」等重複錯誤。
    """
    if not os.path.exists(dict_dir):
        logging.warning(f"字典資料夾 '{dict_dir}' 不存在，正在創建...")
        try:
            os.makedirs(dict_dir)
        except Exception as e:
            logging.error(f"創建字典資料夾 '{dict_dir}' 失敗: {e}")
            return

    global manual_dictionary_maps, customize_dictionary_map
    
    # 檢查並創建 customize.txt (如果不存在)
    customize_file_path = os.path.join(dict_dir, 'customize.txt')
    if not os.path.exists(customize_file_path):
        logging.info(f"'customize.txt' 不存在，將在 '{dict_dir}' 中創建一個帶有範例的新檔案。")
        example_content = (
            "# 這是一個自定義修正字典檔案 (customize.txt)。\n"
            "# 此檔案中的規則會在所有 OpenCC 和其他字典轉換之後最後執行，擁有最高優先級。\n"
            "# 用於修正 OpenCC 的錯誤轉換或添加您個人的特定翻譯。\n"
            "#\n"
            "# 格式：\n"
            "# 原始文字<tab>目標文字\n"
            "#\n"
            "# 範例 (請移除 '#' 來啟用):\n"
            "#慄\t栗\n"
            "#權力遊戲\t冰與火之歌\n"
        )
        try:
            with open(customize_file_path, 'w', encoding='utf-8') as f:
                f.write(example_content)
            logging.info(f"已成功創建範例檔案 '{customize_file_path}'。")
        except Exception as e:
            logging.error(f"創建範例檔案 '{customize_file_path}' 失敗: {e}")

    # 讀取所有字典檔案
    prefixes = ['JP', 'TW', 'HK']
    try:
        all_files_in_dir = sorted(os.listdir(dict_dir))
    except FileNotFoundError:
        logging.warning(f"字典資料夾 '{dict_dir}' 不存在或無法訪問。")
        return

    # 【新增】定義 OpenCC 系統內建字典檔案清單，這些不應被當作簡單字串替換載入
    opencc_core_files = {
        'HKVariants.txt', 'HKVariantsPhrases.txt', 'HKVariantsRevPhrases.txt',
        'TWPhrases.txt', 'TWPhrasesIT.txt', 'TWPhrasesName.txt', 'TWPhrasesOther.txt', 'TWPhrasesRev.txt',
        'TWVariants.txt', 'TWVariantsRevPhrases.txt',
        'JPVariants.txt', 'JPShinjitaiCharacters.txt', 'JPShinjitaiPhrases.txt',
        'STCharacters.txt', 'STPhrases.txt', 'TSCharacters.txt', 'TSPhrases.txt'
    }

    # 處理 JP, TW, HK 前綴的字典
    for prefix in prefixes:
        combined_map = {}
        files_for_prefix = [f for f in all_files_in_dir if f.startswith(prefix) and f.endswith('.txt')]
        
        for filename in files_for_prefix:
            # 【修正】跳過 OpenCC 的系統字典檔，避免載入 1對多 (帶空格) 的錯誤替換格式
            if filename in opencc_core_files:
                continue

            file_path = os.path.join(dict_dir, filename)
            file_specific_map = {}
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.startswith('#') or not line.strip(): continue
                        parts = line.strip().split('\t')
                        if len(parts) >= 2 and parts[0] and parts[1]:
                            key = parts[0]
                            # OpenCC 格式允許一對多（多個候選以空格分隔，例如「衰退 衰頽」），
                            # 只取第一個候選，與 OpenCC 原生行為一致，防止空格污染輸出文字。
                            value = parts[1].split(' ')[0]
                            if not value:
                                continue

                            # 萬一使用者自己建立了同名檔案且不在 opencc_core_files 裡
                            if filename == 'JPVariants.txt':
                                file_specific_map[value] = key
                            else:
                                file_specific_map[key] = value
                
                # 顯示從每個檔案載入的規則數
                logging.info(f"   > 從 '{filename}' 載入 {len(file_specific_map)} 條規則。")
                combined_map.update(file_specific_map)

            except Exception as e:
                logging.error(f"讀取字典檔 {file_path} 時發生錯誤: {e}")
        
        if prefix == 'JP' and 'JPVariants.txt' in files_for_prefix and 'JPVariants.txt' not in opencc_core_files:
            logging.info("已為 'JPVariants.txt' 的規則反轉轉換方向。")

        sorted_keys = sorted(combined_map.keys(), key=len, reverse=True)
        manual_dictionary_maps[prefix] = {k: combined_map[k] for k in sorted_keys}
        logging.info(f"成功為前綴 '{prefix}' 載入共 {len(manual_dictionary_maps[prefix])} 條詞彙映射規則。")

    # 獨立處理 customize.txt
    customize_file = 'customize.txt'
    if customize_file in all_files_in_dir:
        temp_customize_map = {}
        try:
            with open(customize_file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.startswith('#') or not line.strip(): continue
                    parts = line.strip().split('\t')
                    if len(parts) >= 2 and parts[0] and parts[1]:
                        # 同樣只取第一個候選，與 OpenCC 原生行為一致
                        first_value = parts[1].split(' ')[0]
                        if first_value:
                            temp_customize_map[parts[0]] = first_value
            
            sorted_keys = sorted(temp_customize_map.keys(), key=len, reverse=True)
            customize_dictionary_map = {k: temp_customize_map[k] for k in sorted_keys}
            logging.info(f"成功載入 {len(customize_dictionary_map)} 條來自 'customize.txt' 的自定義規則。")
        except Exception as e:
            logging.error(f"讀取 'customize.txt' 時發生錯誤: {e}")

def _apply_manual_dict(text, dict_map):
    """輔助函數，應用手動字典進行替換"""
    if not dict_map:
        return text
    for key, value in dict_map.items():
        if key in text:
            text = text.replace(key, value)
    return text

def load_protected_names(mapping_file='data/tw_hk_name.xlsx'):
    """
    【使用者要求】從 tw_hk_name.xlsx 載入 'Protected Name' sheet 中的保護名單。
    【修改】強化了 sheet 名稱的比對邏輯，使其更能應對命名上的細微差異。
    【修正】將 clear() 移至 try 塊內的 update() 之前，確保載入失敗時不會清空舊有的保護名單。
    """
    global protected_person_names

    if not os.path.exists(mapping_file):
        # 檔案不存在時，excel_handler 會在需要時創建，這裡只需記錄並返回即可
        logging.info(f"受保護人名列表 '{mapping_file}' 不存在，將在需要時創建。本次跳過載入。")
        return

    try:
        # 使用 openpyxl 檢查 sheet 是否存在，避免 pd.read_excel 因 sheet 不存在而拋出錯誤
        wb = openpyxl.load_workbook(mapping_file, read_only=True)
        
        # --- 【錯誤修正 & 穩健性提升 2025-08-20】 ---
        # 為了處理 'Protected Name', 'ProtectedName', 'protected name' 等各種可能的命名方式，
        # 在比對時，我們會將工作表名稱轉為小寫並移除所有空格，然後與目標 'protectedname' 進行比對。
        # 這樣可以大幅提升找到正確工作表的成功率。
        # 增加註解以符合使用者要求，說明此修改是為了應對找不到 'Protected Name' sheet 的問題。
        target_sheet_name_normalized = 'protectedname'
        found_sheet_name = None
        
        # 增加日誌記錄，方便未來除錯時查看程式實際讀取到的所有 sheet 名稱
        logging.debug(f"在 '{mapping_file}' 中找到的 sheet 名稱: {wb.sheetnames}")

        for sheet in wb.sheetnames:
            # 將讀取到的 sheet 名稱正規化：轉為小寫並移除所有空白字符（包括前後和中間的空格）
            normalized_sheet_name = re.sub(r'\s+', '', sheet.lower())
            
            if normalized_sheet_name == target_sheet_name_normalized:
                found_sheet_name = sheet
                logging.info(f"成功匹配到目標 sheet: '{sheet}' (正規化後為: '{normalized_sheet_name}')")
                break

        if not found_sheet_name:
            # 日誌訊息更清晰，告知用戶正在尋找的 sheet 名稱（不分大小寫、忽略空格）
            logging.info(f"在 '{mapping_file}' 中未找到名為 'protected name' (不分大小寫，忽略空格) 的 sheet，跳過載入受保護人名。")
            return
        
        # 確認 Sheet 存在後，再使用 pandas 讀取
        df = pd.read_excel(mapping_file, sheet_name=found_sheet_name, engine='openpyxl')
        
        # 假設保護名單在第一欄
        if not df.empty and df.columns[0] is not None:
            # 移除 NaN 和空字串
            names = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
            # 過濾掉空字串，然後更新到全局 set 中
            valid_names = {name for name in names if name}
            protected_person_names.clear()  # 僅在成功取得資料後才清空，避免載入失敗時遺失舊有保護名單
            protected_person_names.update(valid_names)
            logging.info(f"成功從 '{found_sheet_name}' sheet 載入 {len(valid_names)} 個受保護人名。")
        else:
            logging.info(f"'{found_sheet_name}' sheet 為空或格式不正確，未載入任何受保護人名。")

    except Exception as e:
        logging.error(f"讀取 '{mapping_file}' 的 'protected name' sheet 時發生錯誤: {e}")

# ---------------------------------------------------------------------------
# Simplified-Chinese detection helper
# ---------------------------------------------------------------------------
# Characters that exist ONLY in Simplified Chinese (their Traditional
# equivalents are written differently).  A text with a meaningful ratio of
# these characters is genuinely Simplified and should go through s2t/s2hk.
# Traditional Chinese text should contain virtually none of them, so we can
# safely skip the conversion and avoid false positives like 後 → 后.
_SIMP_ONLY_CHARS: frozenset = frozenset(
    # Common unambiguous simplified-only characters
    # 【2026-11】移除繁體亦使用的 台阿缺参别 及重複字，避免繁體文字被誤判為簡體
    "爱贝备变边补层产场长车尘处传错达带单"
    "担当导灯递点电东动对夺发丰复个过国归"
    "汇会际价见将节结来励炉乱马买满门梦灭"
    "难农盘钱亲请穷书说岁态听图万网问习显"
    "险选学压样叶医义优语远运杂战这种务读"
    "写话应从时间开关记纪线细级练总统继绕"
    "组纸织终纯绿红绑经给络绝综绩缘维续缓"
    "积队陕阵陆随阶联"
)

_SIMP_DETECTION_THRESHOLD = 0.04  # ≥4% simplified-only chars → apply s2t/s2hk


_TRAD_ONLY_CHARS = None


def _trad_only_chars():
    """繁體獨有字：簡轉繁單字字典嘅轉換結果，減去任何喺簡體寫法（單字 key 同詞語 key）出現過嘅字。
    例如「後、髮、們」係繁體獨有；「像」雖然係「象」嘅候選字，但簡體都用，所以唔計。
    字典更新後由 initialize_opencc_converters() 清除快取。"""
    global _TRAD_ONLY_CHARS
    if _TRAD_ONLY_CHARS is None:
        d = os.path.join('opencc_data', 'dictionary')

        def rows(*names):
            for f in names:
                p = os.path.join(d, f)
                if os.path.exists(p):
                    with open(p, 'r', encoding='utf-8') as fin:
                        for line in fin:
                            parts = line.rstrip('\n').split('\t')
                            if len(parts) >= 2 and not line.startswith('#'):
                                yield parts[0], parts[1]
                    return

        simp_chars, values = set(), set()
        for k, v in rows('Merged_STCharacters.txt', 'STCharacters.txt'):
            if len(k) == 1:
                simp_chars.add(k)
                values.update(x for x in v.split() if len(x) == 1 and x != k)
        for k, _v in rows('Merged_STPhrases.txt', 'STPhrases.txt'):
            simp_chars.update(k)
        _TRAD_ONLY_CHARS = values - simp_chars
    return _TRAD_ONLY_CHARS


def _needs_simp_to_trad(text: str) -> bool:
    """Return True if `text` contains Simplified Chinese that needs converting.

    Three-stage strategy:
      1. If text contains chars from _SIMP_ONLY_CHARS → definitely Simplified.
         These chars are unambiguously Simplified-only and never appear in
         Traditional text, so this is a safe fast path.
      2. If s2t changes nothing → already Traditional (or no Chinese). Done.
      3. s2t changes something but no _SIMP_ONLY_CHARS found → ambiguous chars
         like 余 (surname) that s2t converts to 餘. If the text already has
         Traditional-only chars (see _trad_only_chars), treat it as Traditional.

    Falls back to the character-ratio heuristic only when converters are not yet
    initialised (startup edge-case).
    """
    s2t = opencc_converters.get('s2t')
    if s2t is not None:
        import unicodedata as _ud
        nfc = _ud.normalize('NFC', text)

        # Stage 1: fast definitive check — unambiguously Simplified chars present
        if any(c in _SIMP_ONLY_CHARS for c in nfc):
            return True

        # Stage 2: s2t changes nothing → already Traditional
        if s2t.convert(nfc) == nfc:
            return False

        # Stage 3: s2t changes something (ambiguous chars only, e.g. 余→餘).
        # Check if the text already has Traditional-specific chars.
        # 【2026-11】唔再用繁轉簡字典（用家要求停用，曾經令「惡」變「噁」），
        # 改用簡轉繁字典反推出嚟嘅「繁體獨有字」。
        if any(c in _trad_only_chars() for c in nfc):
            return False  # Traditional chars present → already Traditional
        return True

    # Fallback: character-ratio heuristic (used only if s2t not initialised)
    cjk_count = sum(1 for c in text if '一' <= c <= '鿿')
    if cjk_count == 0:
        return False
    simp_count = sum(1 for c in text if c in _SIMP_ONLY_CHARS)
    ratio = simp_count / cjk_count
    if ratio > 0 and ratio < _SIMP_DETECTION_THRESHOLD:
        logging.debug(f"簡體字比率 {ratio:.1%} (共 {simp_count}/{cjk_count})，低於閾值，跳過 s2t/s2hk。")
    return ratio >= _SIMP_DETECTION_THRESHOLD


_CJK_CHAR = r'㐀-䶿一-鿿豈-﫿'
_DQUOTE_RE = re.compile('“([^“”]*)”')   # “…”
_SQUOTE_RE = re.compile('‘([^‘’]*)’')   # ‘…’
_HAS_CJK_RE = re.compile(f'[{_CJK_CHAR}]')
# 引號外緊貼的字：中文字、中文標點（、。「」《》等 U+3000–303F）、全形標點（，：（）等 U+FF00–FFEF）
_CJK_NEIGHBOR_RE = re.compile(f'[{_CJK_CHAR}　-〿＀-￯]')


def convert_cjk_quotes(text):
    """把成對的中文彎引號改為直角引號：“…”→「…」、‘…’→『…』。
    以下任何一個成立就轉換：
      1. 引號內含中文字（例如 “我會回來”）；
      2.【2026-11】引號內沒有中文，但引號外緊貼中文字或中文 / 全形標點（例如 主題曲“Let It Go”大受歡迎）。
    引號外緊貼英文、數字或空格（例如 Christian Ludolf “Buddy” Ebsen Jr.、Grey’s）則保持不變。"""
    if not text or ('“' not in text and '‘' not in text):
        return text

    def _rep(open_q, close_q):
        def f(m):
            inner = m.group(1)
            s, e = m.start(), m.end()
            before = m.string[s - 1] if s > 0 else ''
            after = m.string[e] if e < len(m.string) else ''
            if (_HAS_CJK_RE.search(inner) or _CJK_NEIGHBOR_RE.match(before)
                    or _CJK_NEIGHBOR_RE.match(after)):
                return f'{open_q}{inner}{close_q}'
            return m.group(0)
        return f

    text = _DQUOTE_RE.sub(_rep('「', '」'), text)
    return _SQUOTE_RE.sub(_rep('『', '』'), text)


def normalize_text(text, field_type=None, protect_names=None):
    """
    protect_names（只用於 bio）：呢個人自己嘅名，{原文寫法: 輸出寫法} 或寫法清單（輸出 = 原文）。
      例如 {'于成龙': '于成龍'}：簡介入面成個名唔經 OpenCC 轉換，直接寫成最終人名，
      所以「于成龙出现在位于湾仔」→「于成龍出現在位於灣仔」（名以外嘅「于」照常轉換）。
    【修改】實現全新的文字正規化流程，以解決轉換錯誤並遵循指定順序。
    流程: OpenCC(jp2t) -> JP dict -> [s2t → s2hk，僅當偵測為簡體時] -> TW dict -> HK dict -> customize.txt -> NFC -> (人名專用: 移除空格)
    【使用者要求 2025-09-15】修改轉換鏈，恢復使用 s2t.json 進行核心的簡轉繁轉換。
    【使用者要求 2025-09-15 新增】在 s2t 之後，再執行一次 s2hk 以應用香港地區用詞。
    【使用者要求】【修改 2025-09-19】將 is_person_field 參數替換為 field_type ('name', 'bio', 'title', etc.)。
    【使用者要求 2025-12-10】新增：僅針對人名欄位 (field_type='name')，移除中文字之間的空格 (例如 '澀谷 早苗' -> '澀谷早苗')。
    【2026-05】新增簡體偵測：s2t/s2hk 僅在文字被偵測為簡體時執行，避免繁體「後」被誤轉為「后」等問題。
    """
    if not text or not isinstance(text, str):
        return text or ""

    original_text = text
    processed_text = text
    
    # 保護機制：防止 protected_person_names 中的字/詞被 OpenCC 誤轉
    # name 欄位：保護所有長度的名稱（含單字，如「范」）
    # bio 欄位：只保護多字名稱（≥2字），避免單字在非人名語境被過度保護；另加呢個人自己嘅名（protect_names）
    placeholders = {}
    if field_type in ('name', 'bio'):
        min_len = 1 if field_type == 'name' else 2
        protect_targets = {n: n for n in (protected_person_names or ()) if len(n) >= min_len}
        if field_type == 'bio' and protect_names:
            pairs = protect_names.items() if isinstance(protect_names, dict) else ((n, n) for n in protect_names)
            for src, out in pairs:
                if isinstance(src, str) and len(src) >= 2:
                    protect_targets[src] = out or src
        for i, name in enumerate(sorted(protect_targets, key=len, reverse=True)):
            if name in processed_text:
                placeholder = f"@@PROTECTED_{i}@@"
                processed_text = processed_text.replace(name, placeholder)
                placeholders[placeholder] = protect_targets[name]

    # 步驟 1: OpenCC (jp2t.json) - From Japanese new characters to Traditional
    if opencc_converters.get('jp2t'):
        processed_text = opencc_converters['jp2t'].convert(processed_text)
    
    # 步驟 2: JP 手動字典
    processed_text = _apply_manual_dict(processed_text, manual_dictionary_maps['JP'])
    
    # 步驟 2b【2026-11 H2】中文彎引號改為香港慣用的直角引號：“…” → 「…」、‘…’ → 『…』
    # 放在 AI 校對之前：AI 不需再改標點（以前會被當成不合法字還原）。
    processed_text = convert_cjk_quotes(processed_text)

    # 步驟 3 & 4: 僅當文字被偵測為簡體中文時才執行 s2t / s2hk。
    # 若文字已是繁體（如 TMDB zh-HK 內容），跳過這兩步，避免「後」→「后」等誤轉。
    if _needs_simp_to_trad(processed_text):
        pre_s2t_text = processed_text
        if opencc_converters.get('s2t'):
            processed_text = opencc_converters['s2t'].convert(processed_text)
        if opencc_converters.get('s2hk'):
            processed_text = opencc_converters['s2hk'].convert(processed_text)
        # 步驟 4b: 包含多義字 (data/tw_hk_name.xlsx「AI Convert」頁) 的句子交 AI 校對 (設定見 ai_config.json)
        try:
            from others import ai_s2t
            processed_text = ai_s2t.refine(pre_s2t_text, processed_text)
        except Exception as e:
            logging.warning(f"[AI簡轉繁] 校對失敗，保留 OpenCC 結果: {e}")
    else:
        logging.debug("文字已為繁體，跳過 s2t/s2hk 轉換。")

    # 步驟 5: TW 手動字典 (修正台灣用語)
    processed_text = _apply_manual_dict(processed_text, manual_dictionary_maps['TW'])
        
    # 步驟 6: HK 手動字典 (修正香港用語) - 在已轉為繁體的基礎上，應用香港地區用詞。
    processed_text = _apply_manual_dict(processed_text, manual_dictionary_maps['HK'])

    # 步驟 7: 自定義字典 (customize.txt) - 最高優先級修正
    processed_text = _apply_manual_dict(processed_text, customize_dictionary_map)

    # 還原被保護的人名/字詞
    if placeholders:
        for placeholder, original_name in placeholders.items():
            processed_text = processed_text.replace(placeholder, original_name)

    # 步驟 8: Unicode NFC 正規化
    processed_text = unicodedata.normalize('NFC', processed_text)

    # 【使用者要求 2025-12-10】步驟 9: 人名專用處理 - 移除中文字之間的空格
    # 例如：'澀谷 早苗' -> '澀谷早苗'，但 'Tom Cruise' 保持不變。
    if field_type == 'name':
        # 正則表達式：匹配 [漢字] [空格] [漢字] 的模式
        # \u4e00-\u9fff: 常用漢字範圍
        # \s+: 一個或多個空白字符
        pattern = r'([\u4e00-\u9fff])\s+([\u4e00-\u9fff])'
        
        # 循環替換直到沒有匹配項，以處理多個連續空格的情況 (如 '中 文 字')
        while re.search(pattern, processed_text):
            processed_text = re.sub(pattern, r'\1\2', processed_text)

    # Fallback: 若 placeholder 機制未能保護受保護字元（例如 protected_person_names 在呼叫時為空），
    # 則在最後逐字元比對原文，把被誤轉的字元強制還原。
    # 僅在文字長度不變時適用（s2hk 的 HK variant 替換均為 1:1，不改變長度）。
    # 注意：此 fallback 不依賴 protected_person_names 是否有值——即使 set 為空亦可運作，
    # 但必須有 protected_person_names 才知道哪些字元需要保護。
    # 【2026-11】bio 只還原受保護詞語（≥2 字）範圍入面嘅字：以前會還原全文所有受保護單字，
    # 令「位于」呢類非人名嘅「于」都保持簡體。
    if field_type in ('name', 'bio') and protected_person_names and len(processed_text) == len(original_text):
        allowed = None
        if field_type == 'bio':
            allowed = set()
            for w in (n for n in protected_person_names if len(n) >= 2):
                start = original_text.find(w)
                while start != -1:
                    allowed.update(range(start, start + len(w)))
                    start = original_text.find(w, start + 1)
        corrected = list(processed_text)
        reverted = []
        for idx, (orig_c, cur_c) in enumerate(zip(original_text, processed_text)):
            if orig_c != cur_c and orig_c in protected_person_names and (allowed is None or idx in allowed):
                corrected[idx] = orig_c
                reverted.append(f"pos {idx}: '{cur_c}'(U+{ord(cur_c):04X}) → '{orig_c}'(U+{ord(orig_c):04X})")
        if reverted:
            processed_text = ''.join(corrected)
            logging.info(f"normalize_text [保護回退]: 已還原受保護字元 {reverted}")

    final_text = processed_text

    if final_text != original_text:
        log_msg = f"正規化轉換: 原文='{original_text[:50]}...' -> 轉換後='{final_text[:50]}...'"
        if field_type == 'name' and placeholders:
            log_msg += f" (人名字段模式，保護了: {list(placeholders.values())})"
        logging.info(log_msg)
    
    return final_text

def normalize_to_hk(text, field_type=None, is_person_field=None, protect_names=None):
    """
    為保持舊代碼兼容性，此函數現在調用新的、統一的正規化函數 normalize_text。
    【修改 2025-09-19】傳遞 field_type 參數以區分人名和其他欄位。
    is_person_field 為舊參數別名：True 時等同 field_type='name'。
    protect_names：見 normalize_text（只用於 bio）。
    """
    if field_type is None and is_person_field:
        field_type = 'name'
    return normalize_text(text, field_type=field_type, protect_names=protect_names)

def is_all_hanzi(text):
    """
    檢查字串是否完全由漢字/Kanji組成。
    """
    if not text or not isinstance(text, str):
        return False
    # 此正則表達式涵蓋了最常見的 CJK 統一漢字範圍
    # U+4E00..U+9FFF: CJK 統一漢字
    # U+3400..U+4DBF: CJK 統一漢字擴展A
    # U+F900..U+FAFF: CJK 兼容漢字
    return bool(re.match(r'^[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]+$', text))

# --- 在模組載入時，執行一次初始化 ---
initialize_opencc_converters()
load_manual_dictionaries()

def is_simplified_chinese(text):
    """
    【修改 2025-09-15】使用雙向轉換檢查文本是否為簡體中文，以避免地區用字差異造成的誤判。
    判定為簡體的條件：
    1. 文本經過「簡轉繁」後必須有所改變。
    2. 將第1步的結果再經過「繁轉簡」後，必須能還原成原始文本。
    這樣可以過濾掉 '裏' -> '裡' 這類僅地區用字不同的情況，並準確識別出像 '随着戴夫的幻觉' 這類真正的簡體內容。
    【核心修正 2025-09-15】簡化判定邏輯。由於雙向轉換在某些情況下可能因異體字而失敗，
    現在改為更直接的單向比較：只要文本經過「簡轉繁」處理後與原文不同，就將其判定為簡體中文。
    這提高了對簡體字的識別率，是解決日誌中 is_simplified=False 錯誤的關鍵。
    """
    if not text or not isinstance(text, str): # is_chinese check removed to avoid circular dependency if it uses this
        return False

    s2t_checker = opencc_converters.get('s2t_check')

    if not s2t_checker:
        logging.warning("is_simplified_chinese: 簡轉繁檢查器 's2t_check' 未初始化，無法檢查。假設為非簡體。")
        return False

    # 為了比較，先對原始文本進行 NFC 正規化
    text_nfc = unicodedata.normalize('NFC', text)

    # 步驟 1: 簡轉繁
    converted_to_traditional = s2t_checker.convert(text_nfc)
    converted_to_traditional_nfc = unicodedata.normalize('NFC', converted_to_traditional)

    # 【核心修正】只要簡轉繁後文本有變化，就判定為簡體。
    is_simplified = (converted_to_traditional_nfc != text_nfc)
    
    # 增加詳細日誌以供除錯
    if is_simplified:
        logging.debug(f"文本 '{text[:30]}...' 被判定為簡體。 原文 -> 簡轉繁: '{converted_to_traditional_nfc[:30]}...'")
    else:
        logging.debug(f"文本 '{text[:30]}...' 被判定為非簡體 (簡轉繁後無變化)。")

    return is_simplified

def compare_with_zh_hk(new_title, new_description, new_runtime, new_name, new_bio, zh_hk_title, zh_hk_description, zh_hk_runtime, zh_hk_name, zh_hk_bio, media_type, overwrite_titles="ask", overwrite_descriptions="ask", tv_show_title=None, context="Unknown"):
    from others.analysis_excel import is_generic_title # Keep local import if preferred
    
    # 統一對所有輸入文本進行normalize和spacing處理，以便公平比較和決定最終輸出
    # "new_" 系列來自API的多語言來源，已經過初步選擇
    # "zh_hk_" 系列是TMDB上已有的zh-HK數據

    # 處理 "new" 數據 (來自 API 的候選翻譯)
    processed_new_title = add_spaces_between_en_zh_num(normalize_to_hk(new_title), is_long_text=False) if new_title else ""
    processed_new_description = add_spaces_between_en_zh_num(normalize_to_hk(new_description), is_long_text=True) if new_description else ""
    processed_new_name = add_spaces_between_en_zh_num(normalize_to_hk(new_name), is_long_text=False) if new_name else ""
    processed_new_bio = add_spaces_between_en_zh_num(normalize_to_hk(new_bio), is_long_text=True) if new_bio else ""

    # 處理 TMDB 上已有的 "zh-HK" 數據
    # 假設 zh_hk_title 等也可能不是完美的香港繁體 (例如可能是簡體或未NFC)，所以也用 normalize_to_hk
    processed_zh_hk_title = add_spaces_between_en_zh_num(normalize_to_hk(zh_hk_title), is_long_text=False) if zh_hk_title else ""
    processed_zh_hk_description = add_spaces_between_en_zh_num(normalize_to_hk(zh_hk_description), is_long_text=True) if zh_hk_description else ""
    processed_zh_hk_name = add_spaces_between_en_zh_num(normalize_to_hk(zh_hk_name), is_long_text=False) if zh_hk_name else ""
    processed_zh_hk_bio = add_spaces_between_en_zh_num(normalize_to_hk(zh_hk_bio), is_long_text=True) if zh_hk_bio else ""

    logging.info(f"{context} 比較 (已處理空格和繁體化): "
                 f"新標題='{processed_new_title}', 新描述='{processed_new_description[:50] if processed_new_description else ''}...', 新時長={new_runtime}, "
                 f"新名稱='{processed_new_name}', 新簡介='{processed_new_bio[:50] if processed_new_bio else ''}...', "
                 f"zh-HK 標題='{processed_zh_hk_title}', zh-HK 描述='{processed_zh_hk_description[:50] if processed_zh_hk_description else ''}...', zh-HK 時長={zh_hk_runtime}, "
                 f"zh-HK 名稱='{processed_zh_hk_name}', zh-HK 簡介='{processed_zh_hk_bio[:50] if processed_zh_hk_bio else ''}...'")

    # Initialize return values
    should_write_to_excel = False
    is_same_as_zh_hk = False 
    final_title_for_excel = ""
    final_description_for_excel = ""
    final_runtime_for_excel = None 
    final_name_for_excel = ""
    final_bio_for_excel = ""

    if media_type in ["movie", "tv", "season", "episode"]:
        # TITLE processing
        candidate_title_for_excel = processed_new_title
        if is_generic_title(candidate_title_for_excel): # 檢查處理後的 "new" 標題是否泛用
            logging.info(f"{context} Title: 處理後的新標題 '{candidate_title_for_excel}' 為泛用型，Excel中此欄位視為空。")
            candidate_title_for_excel = "" 
        
        # 檢查現有 zh-HK 標題是否為“理想的”繁體中文 (非簡體，且是中文)
        # is_simplified_chinese 會處理 NFC 並與 s2hk 轉換結果比較
        zh_hk_title_is_ideal_trad = zh_hk_title and is_chinese(zh_hk_title) and not is_simplified_chinese(zh_hk_title)

        if zh_hk_title_is_ideal_trad:
            final_title_for_excel = "" # Excel stays empty if zh-HK is already ideal
            logging.info(f"{context} Title (No Overwrite): zh-HK '{zh_hk_title}' (處理後: '{processed_zh_hk_title}') 已是理想繁體。Excel中此欄位設為空。")
        # 如果 zh-HK 不是理想繁體 (例如是簡體，非中文，或空)，且我們有候選標題
        elif candidate_title_for_excel:
            # 如果候選標題是英文但無描述，則不寫入
            if is_english(candidate_title_for_excel) and not is_chinese(candidate_title_for_excel) and not processed_new_description:
                final_title_for_excel = ""
                logging.info(f"{context} Title: 候選英文標題 '{candidate_title_for_excel}' 但無描述，不寫入。")
            # 否則，使用候選標題 (此標題已處理過空格和繁體化)
            else:
                final_title_for_excel = candidate_title_for_excel
                should_write_to_excel = True
                logging.info(f"{context} Title: 使用標題 '{final_title_for_excel}'。")
        else: # 無有效候選標題可寫入
            final_title_for_excel = ""
            logging.info(f"{context} Title: 無有效標題可供寫入。")

        # DESCRIPTION processing
        candidate_desc_for_excel = processed_new_description
        if candidate_desc_for_excel in NO_DESCRIPTION_TEXT: candidate_desc_for_excel = ""
        
        zh_hk_desc_is_ideal_trad = zh_hk_description and is_chinese(zh_hk_description) and not is_simplified_chinese(zh_hk_description)

        if zh_hk_desc_is_ideal_trad:
            final_description_for_excel = ""
            logging.info(f"{context} Description (No Overwrite): zh-HK 描述 (處理後: '{processed_zh_hk_description[:30]}...') 已是理想繁體。Excel中此欄位設為空。")
        elif candidate_desc_for_excel:
            final_description_for_excel = candidate_desc_for_excel
            should_write_to_excel = True
            logging.info(f"{context} Description: 使用描述 (處理後長度: {len(final_description_for_excel)}).")
        else:
            final_description_for_excel = ""
            logging.info(f"{context} Description: 無有效描述可供寫入。")
        
        # RUNTIME processing
        if new_runtime is not None and new_runtime != zh_hk_runtime:
            final_runtime_for_excel = new_runtime
            should_write_to_excel = True
            logging.info(f"{context} Runtime: 時長不同，使用新時長 {new_runtime} (zh-HK: {zh_hk_runtime})。")
        else:
             final_runtime_for_excel = None 
             if new_runtime == zh_hk_runtime and new_runtime is not None:
                  logging.info(f"{context} Runtime: 時長與 zh-HK ({zh_hk_runtime}) 相同，不寫入。")
             elif new_runtime is None and zh_hk_runtime is not None:
                  logging.info(f"{context} Runtime: 新時長為空，zh-HK 時長 ({zh_hk_runtime}) 已存在，不改動。")

        # Re-evaluate should_write_to_excel if all final values are empty
        if not final_title_for_excel and \
           not final_description_for_excel and \
           (final_runtime_for_excel is None): # Simplified: if all are empty, nothing to write
            # Check if the original intent was to write something that got cleared
            # This part is tricky. The should_write_to_excel flag is set if *any* final_X is non-empty.
            # If, after all logic, final_title, final_desc, final_runtime are all empty/None, then nothing should be written.
            if not (final_title_for_excel or final_description_for_excel or final_runtime_for_excel is not None):
                should_write_to_excel = False 
                logging.info(f"{context}: 最終標題、描述、時長均無需寫入Excel (皆為空/None)。")


        # is_same_as_zh_hk determination:
        # Compare processed new data with processed existing zh-HK data.
        # An item is "same" if the processed new value is identical to the processed zh-HK value.
        # Or if new value is empty and zh-HK is already ideal (implies no change needed from new data).
        
        is_title_effectively_same = (processed_new_title == processed_zh_hk_title)
        is_desc_effectively_same = (processed_new_description == processed_zh_hk_description)
        is_runtime_effectively_same = (new_runtime == zh_hk_runtime) if new_runtime is not None else (zh_hk_runtime is None)

        # If zh-HK is already ideal, and the new processed data matches it, then it's "same".
        # If zh-HK is NOT ideal, then new data (even if identical to raw new data) cannot be "same" in terms of "no change needed from zh-HK's current state".
        # However, the question is "is the *new data* effectively the same as what's already *ideally* on zh-HK".
        # So, if `processed_new_title` (ideal form of new data) == `processed_zh_hk_title` (ideal form of current zh-HK data), they are the same.
        is_same_as_zh_hk = is_title_effectively_same and is_desc_effectively_same and is_runtime_effectively_same
        
        if is_same_as_zh_hk:
             logging.info(f"{context}: 判斷為所有提供的 Video 新資料經處理後均與已處理的 zh-HK 狀態相同。")
        else:
             logging.info(f"{context}: Video 新資料經處理後與已處理的 zh-HK 狀態存在差異。Title same: {is_title_effectively_same}, Desc same: {is_desc_effectively_same}, Runtime same: {is_runtime_effectively_same}")


    elif media_type == "person":
        candidate_name_for_excel = processed_new_name
        # Rule: 人名只接受中文 (after processing)
        if candidate_name_for_excel and not is_chinese(candidate_name_for_excel):
            logging.info(f"{context} Name: 處理後的新人名 '{candidate_name_for_excel}' 非中文，視為空。")
            candidate_name_for_excel = ""

        zh_hk_name_is_ideal_trad = zh_hk_name and is_chinese(zh_hk_name) and not is_simplified_chinese(zh_hk_name)

        if zh_hk_name_is_ideal_trad:
            final_name_for_excel = ""
            logging.info(f"{context} Name (No Overwrite): zh-HK人名 '{zh_hk_name}' (處理後: '{processed_zh_hk_name}') 已是理想繁體。Excel中此欄位設為空。")
        elif candidate_name_for_excel: # new_name is valid processed Chinese
            final_name_for_excel = candidate_name_for_excel
            should_write_to_excel = True
            logging.info(f"{context} Name: 使用處理後的人名 '{final_name_for_excel}'。")
        else: 
            final_name_for_excel = ""
            logging.info(f"{context} Name: 無有效中文人名可供寫入。")

        # BIO processing
        candidate_bio_for_excel = processed_new_bio
        zh_hk_bio_is_ideal_trad = zh_hk_bio and is_chinese(zh_hk_bio) and not is_simplified_chinese(zh_hk_bio) # Bio can be non-Chinese, but if Chinese, must be trad.

        if zh_hk_bio_is_ideal_trad or (zh_hk_bio and not is_chinese(zh_hk_bio) and not is_simplified_chinese(zh_hk_bio)): # If zh-HK bio is ideal (Trad Chinese or valid non-Chinese)
            final_bio_for_excel = ""
            logging.info(f"{context} Bio (No Overwrite): zh-HK簡介 (處理後: '{processed_zh_hk_bio[:30]}...') 已是理想狀態。Excel中此欄位設為空。")
        elif candidate_bio_for_excel: # We have a new bio to consider
            final_bio_for_excel = candidate_bio_for_excel
            should_write_to_excel = True
            logging.info(f"{context} Bio: 使用處理後的簡介。")
        elif not candidate_bio_for_excel and zh_hk_bio and is_simplified_chinese(zh_hk_bio): # New bio is empty, but existing zh-HK bio is simplified
            final_bio_for_excel = processed_zh_hk_bio # This will be the HK-Traditional version of the simplified bio
            should_write_to_excel = True
            logging.info(f"{context} Bio: 新簡介為空，但 zh-HK 簡介為簡體，轉換並寫入處理後的版本。")
        else: # New bio is empty, and zh-HK is not simplified (or also empty)
            final_bio_for_excel = ""
            logging.info(f"{context} Bio: 無有效新簡介可供寫入，且zh-HK非簡體或已處理。")


        if not final_name_for_excel and not final_bio_for_excel:
            if not (final_name_for_excel or final_bio_for_excel): # if both are truly empty strings for excel
                should_write_to_excel = False
                logging.info(f"{context}: 最終人名和簡介均無需寫入Excel。")

        is_name_effectively_same = (processed_new_name == processed_zh_hk_name)
        is_bio_effectively_same = (processed_new_bio == processed_zh_hk_bio)
            
        is_same_as_zh_hk = is_name_effectively_same and is_bio_effectively_same
        if is_same_as_zh_hk:
            logging.info(f"{context}: 判斷為所有提供的 Person 新資料經處理後均與已處理的 zh-HK 狀態相同。")
        else:
            logging.info(f"{context}: Person 新資料經處理後與已處理的 zh-HK 狀態存在差異。Name same: {is_name_effectively_same}, Bio same: {is_bio_effectively_same}")


    logging.info(f"{context} 最終決策: 寫入 Excel={should_write_to_excel}, 與 zh-HK 理想狀態(處理後比較)相同={is_same_as_zh_hk}. "
                 f"Excel輸出 Title='{final_title_for_excel}', Desc='{final_description_for_excel[:30]}...', Runtime={final_runtime_for_excel}, "
                 f"Name='{final_name_for_excel}', Bio='{final_bio_for_excel[:30]}...'")
    
    return should_write_to_excel, is_same_as_zh_hk, \
           final_title_for_excel, final_description_for_excel, final_runtime_for_excel, \
           final_name_for_excel, final_bio_for_excel

def process_translatable_item(
    item_type,  # "title", "description", "name", "bio", "runtime"
    media_category, # "video" (movie/tv/season/episode), "person"
    zh_hk_value,
    zh_tw_value,
    zh_cn_value,
    en_value=None,
    ja_value=None,
    ko_value=None,
    original_value=None, # 【新增】接收 original_title / original_name
    overwrite_mode=False,
    sibling_desc_is_empty_for_eng_title=None, 
    current_runtime_for_video=None, 
    zh_hk_runtime_for_video=None,
    context_for_logging="Item",
    protect_names=None  # 【2026-11】只用於 bio：呢個人自己嘅名 {原文寫法: 最終人名}，見 normalize_text
):
    """
    【修改】處理單個可翻譯字段，決定最終寫入 Excel 的值和是否需要寫入。
    此版本修正了以下問題：
    1. 【使用者要求 2025-08-25】修正了人員姓名的選擇邏輯，將維基百科找到的譯名 (`en_value`) 提升至最高優先級，
       以便用它來修正 TMDB 上可能不正確的 `zh-HK` 譯名。此修改僅對人物姓名生效。
    2. 【修改 2025-09-19】移除了 is_person_field 判斷，改為直接傳遞更精確的 item_type 給正規化函數，
       以實現對人名和簡介不同級別的文字保護。
    3. 修正了判斷 zh-HK 值是否「完美」的邏輯，改為直接比較原始值與其理想形態。
    4. 【使用者要求】新增 tw_hk_name.xlsx 作為最終的「守門人」規則。在所有 OpenCC 和字典轉換完成後，
       會使用 apply_tw_hk_name_replacements 函數對「候選值」和「zh-HK 理想形態」進行最終的詞彙修正，
       然後才進行比較。這確保了如「汽車人」->「博派」的轉換能被正確評估和觸發。
    5. 【使用者要求】新增返回 `was_imperfect` 標誌，用於判斷原始 zh-HK 內容是否需要修正才能達到理想狀態。
    6. 【使用者要求 2025-09-05 修改】如果 original_value (原始標題/名稱) 本身是中文，則將其作為高優先級的候選來源 (僅次於 zh-HK)。
    7. 【使用者要求 2025-08-18】調整語言選擇優先級，以處理 en-US 欄位包含中文內容的情況。
       當 zh-HK, zh-TW, zh-CN 均無有效中文內容時，會同時檢查 en_value (英文區標題) 和 original_value (原始標題)，並優先使用 original_value 中的中文內容。
    8. 【使用者要求 2025-09-02, 2025-09-05 修改】新增回退邏輯：如果遍歷所有來源後都找不到一個有效的中文標題/名稱，
       則僅回退檢查 zh-TW, zh-CN, zh-HK 欄位是否包含可用的英文標題。
    返回: (processed_candidate_value, should_write_this_item_to_excel, is_identical_to_ideal_zh_hk, was_imperfect)
    """
    from others.excel_handler import replace_tw_hk_names as apply_tw_hk_name_replacements
    
    _clean = lambda v: "" if v in NO_DESCRIPTION_TEXT else (v or "")
    zh_hk_value_raw, zh_tw_value_raw, zh_cn_value_raw, en_value_raw, ja_value_raw, ko_value_raw, original_value_raw = map(_clean, [zh_hk_value, zh_tw_value, zh_cn_value, en_value, ja_value, ko_value, original_value])

    is_long_text_field = item_type in ["description", "bio"]

    if item_type == "title":
        zh_hk_value_raw = clean_generic_title(zh_hk_value_raw)
        zh_tw_value_raw = clean_generic_title(zh_tw_value_raw)
        zh_cn_value_raw = clean_generic_title(zh_cn_value_raw)
        en_value_raw = clean_generic_title(en_value_raw)
        ja_value_raw = clean_generic_title(ja_value_raw)
        ko_value_raw = clean_generic_title(ko_value_raw)
        original_value_raw = clean_generic_title(original_value_raw)

    raw_prospective_value = ""
    source_language_tag = ""
    
    if item_type != "runtime":
        # --- 語言選擇優先級 ---
        # 步驟 1: 優先尋找任何來源的中文譯名。
        # 步驟 2: 如果找不到任何中文譯名，則回退尋找英文譯名。
        
        # 階段 1: 搜尋中文候選值
        if (media_category == "person" and item_type == "name" and is_chinese(en_value_raw)
                and (wiki_name_replaces_existing() or not is_chinese(zh_hk_value_raw))):
            raw_prospective_value, source_language_tag = en_value_raw, "wiki"
        elif is_chinese(zh_hk_value_raw): 
            raw_prospective_value, source_language_tag = zh_hk_value_raw, "zh-HK"
        # 【使用者要求 2025-09-05】如果 original_value (原始標題/名稱) 是中文，賦予其高優先級 (僅次於 zh-HK)。
        elif original_value_raw and is_chinese(original_value_raw):
            raw_prospective_value, source_language_tag = original_value_raw, "original_value (as Chinese)"
        elif is_chinese(zh_tw_value_raw): 
            raw_prospective_value, source_language_tag = zh_tw_value_raw, "zh-TW"
        elif is_chinese(zh_cn_value_raw): 
            raw_prospective_value, source_language_tag = zh_cn_value_raw, "zh-CN"
        elif en_value_raw and is_chinese(en_value_raw):
            raw_prospective_value, source_language_tag = en_value_raw, "en_value (as Chinese)"
        elif ja_value_raw and is_all_hanzi(ja_value_raw):
            raw_prospective_value, source_language_tag = ja_value_raw, "ja-JP (Hanzi)"
        elif ko_value_raw and is_all_hanzi(ko_value_raw):
            raw_prospective_value, source_language_tag = ko_value_raw, "ko-KR (Hanzi)"
        else:
            # 階段 2: 若無中文候選值，則回退至英文候選值
            logging.info(f"{context_for_logging} ({item_type}): 未找到任何中文候選值，將嘗試尋找英文候選值。")
            if media_category == "video": # 此回退邏輯主要針對影片標題
                if is_english(zh_tw_value_raw):
                    raw_prospective_value, source_language_tag = zh_tw_value_raw, "zh-TW (as English)"
                elif is_english(zh_cn_value_raw):
                    raw_prospective_value, source_language_tag = zh_cn_value_raw, "zh-CN (as English)"
                elif is_english(zh_hk_value_raw): 
                    raw_prospective_value, source_language_tag = zh_hk_value_raw, "zh-HK (as English)"

        if media_category == "person" and item_type == "name" and (not raw_prospective_value or not is_chinese(raw_prospective_value)):
            logging.info(f"{context_for_logging} ({item_type}): 人名候選值 '{raw_prospective_value}' 非中文或為空，清空。")
            raw_prospective_value = ""
        if item_type == "description" and (not raw_prospective_value or not is_chinese(raw_prospective_value)):
            logging.info(f"{context_for_logging} ({item_type}): 描述候選值 '{raw_prospective_value[:30]}...' 非中文或為空，清空。")
            raw_prospective_value = ""

        if raw_prospective_value:
            # 【修改 2025-09-19】將 is_person_field 替換為更精確的 field_type=item_type
            value_after_s2hk = normalize_to_hk(raw_prospective_value, field_type=item_type, protect_names=protect_names)
            prospective_value_processed = add_spaces_between_en_zh_num(value_after_s2hk, item_type_for_log=f"{context_for_logging} {item_type} (來自 {source_language_tag})", is_long_text=is_long_text_field)
            logging.info(f"{context_for_logging} ({item_type}): 候選值 (來自 {source_language_tag}) '{raw_prospective_value[:30]}...' -> 初步處理後 (簡繁+空格) '{prospective_value_processed[:30]}...'")
        else:
            logging.info(f"{context_for_logging} ({item_type}): 無有效原始候選值。")
            prospective_value_processed = ""
    else: # Runtime
        prospective_value_processed = ""

    zh_hk_value_ideal_form = ""
    if item_type != "runtime" and zh_hk_value_raw:
        # 【修改 2025-09-19】將 is_person_field 替換為更精確的 field_type=item_type
        zh_hk_value_ideal_form = add_spaces_between_en_zh_num(normalize_to_hk(zh_hk_value_raw, field_type=item_type, protect_names=protect_names), item_type_for_log=f"{context_for_logging} {item_type} (API zh-HK基準)", is_long_text=is_long_text_field)
    
    if item_type in ["title", "name", "description", "bio"]:
        logging.debug(f"{context_for_logging} ({item_type}): 正在應用 tw_hk_name.xlsx 作為最終守門人規則...")
        
        context_title_for_sheet_selection = prospective_value_processed if item_type in ["title", "name"] else (zh_hk_value_raw or "")
        context_desc_for_sheet_selection = prospective_value_processed if item_type in ["description", "bio"] else (zh_hk_value_raw or "")

        if item_type in ["title", "name"]:
            prospective_value_processed, _ = apply_tw_hk_name_replacements(title=prospective_value_processed, description=context_desc_for_sheet_selection, return_strings=True)
        else: # description, bio
            _, prospective_value_processed = apply_tw_hk_name_replacements(title=context_title_for_sheet_selection, description=prospective_value_processed, return_strings=True)

        if item_type in ["title", "name"]:
            zh_hk_value_ideal_form, _ = apply_tw_hk_name_replacements(title=zh_hk_value_ideal_form, description=context_desc_for_sheet_selection, return_strings=True)
        else: # description, bio
            _, zh_hk_value_ideal_form = apply_tw_hk_name_replacements(title=context_title_for_sheet_selection, description=zh_hk_value_ideal_form, return_strings=True)
            
        logging.debug(f"{context_for_logging} ({item_type}): 應用守門人規則後: 候選值='{prospective_value_processed[:30]}...', zh_hk_理想態='{zh_hk_value_ideal_form[:30]}...'")

    # 【使用者要求 2025-09-11】錯誤修正：在做出最終決定前，檢查處理後的候選值是否為泛用型標題 (例如純數字 "1")。
    # 這是為了確保在所有語言轉換和選擇後，不會意外地將一個無意義的數字作為最終標題。
    if item_type in ["title", "name"] and is_generic_title(prospective_value_processed):
        logging.info(f"{context_for_logging} ({item_type}): 最終候選值 '{prospective_value_processed}' 被識別為泛用型，將其視為空值。")
        prospective_value_processed = ""

    should_write_this_item_to_excel = False
    is_identical_to_ideal_zh_hk = True
    was_imperfect_and_needed_correction = False

    if item_type != "runtime":
        is_identical_to_ideal_zh_hk = (prospective_value_processed == zh_hk_value_ideal_form)
        
        if zh_hk_value_raw and zh_hk_value_raw != zh_hk_value_ideal_form:
            # 剔除 TMDB API 固有的 \r\n / \r 及換行後空格，這類差異寫入後也不會改變，視為無需修正
            _raw_for_cmp = re.sub(r'\n[ \t]+', '\n', zh_hk_value_raw.replace('\r\n', '\n').replace('\r', '\n'))
            if _raw_for_cmp != zh_hk_value_ideal_form:
                was_imperfect_and_needed_correction = True
            else:
                logging.info(f"{context_for_logging} ({item_type}): 原始 zh-HK 值與理想態的差異僅為換行符/空格，視為無需修正。")

        if prospective_value_processed:
            if not is_identical_to_ideal_zh_hk:
                should_write_this_item_to_excel = True
                logging.info(f"{context_for_logging} ({item_type}): 候選值處理後與 API zh-HK 理想態不同，應寫入。差異: {_diff_hint(prospective_value_processed, zh_hk_value_ideal_form)}")
            else:
                if was_imperfect_and_needed_correction:
                    should_write_this_item_to_excel = True
                    logging.info(f"{context_for_logging} ({item_type}): 候選值與 API zh-HK 理想態相同，但原始 zh-HK 值與其理想態不同，應寫入以修正。差異: {_diff_hint(zh_hk_value_raw, zh_hk_value_ideal_form)}")
                else:
                    logging.info(f"{context_for_logging} ({item_type}): 候選值處理後與 API zh-HK 理想態相同，無需寫入。")
        
        if not overwrite_mode:
            if not was_imperfect_and_needed_correction and zh_hk_value_raw:
                should_write_this_item_to_excel = False
                logging.info(f"{context_for_logging} ({item_type}) (No Overwrite): API zh-HK 值 ('{zh_hk_value_raw[:30]}...') 已是最終理想形態，無需寫入。")

    elif item_type == "runtime" and media_category == "video":
        is_identical_to_ideal_zh_hk = (current_runtime_for_video == zh_hk_runtime_for_video)
        if current_runtime_for_video is not None and not is_identical_to_ideal_zh_hk:
            should_write_this_item_to_excel = True
            prospective_value_processed = str(current_runtime_for_video)
            logging.info(f"{context_for_logging} ({item_type}): 時長 {current_runtime_for_video} 與 API zh-HK 時長 {zh_hk_runtime_for_video} 不同，應寫入。")
        else:
            if is_identical_to_ideal_zh_hk:
                logging.info(f"{context_for_logging} ({item_type}): 時長與 API zh-HK 時長相同 ({current_runtime_for_video})。")
            else: 
                logging.info(f"{context_for_logging} ({item_type}): 無有效目標時長，API zh-HK 時長為 {zh_hk_runtime_for_video}。")
    
    final_excel_value = prospective_value_processed if should_write_this_item_to_excel else ""
    logging.info(f"{context_for_logging} ({item_type}) 最終決策 -> 應寫入: {should_write_this_item_to_excel}, "
                 f"與 zh-HK 理想態一致: {is_identical_to_ideal_zh_hk}, "
                 f"將寫入 Excel 的值: '{final_excel_value[:50]}...'")

    return prospective_value_processed, should_write_this_item_to_excel, is_identical_to_ideal_zh_hk, was_imperfect_and_needed_correction

def _diff_hint(a: str, b: str, ctx: int = 15) -> str:
    """返回兩個字串首個差異位置的摘要，包含 Unicode 碼點，方便 debug。"""
    for i, (ca, cb) in enumerate(zip(a, b)):
        if ca != cb:
            start = max(0, i - ctx)
            snippet_a = a[start:i + ctx]
            snippet_b = b[start:i + ctx]
            return (f"位置 {i}: ...'{snippet_a}'... vs ...'{snippet_b}'... "
                    f"(U+{ord(ca):04X} '{ca}' → U+{ord(cb):04X} '{cb}')")
    if len(a) != len(b):
        shorter, longer = (a, b) if len(a) < len(b) else (b, a)
        extra = longer[len(shorter):len(shorter) + ctx]
        return f"長度不同 ({len(a)} vs {len(b)})，多出部份: '{extra}...'"
    return "完全相同"


def is_chinese(text):
    """檢查文本是否為中文，排除日文假名、韓文、特定日文符號等非中文字符，使用緩存避免重複日誌"""
    if not text or not isinstance(text, str):
        logging.debug(f"文本為空或非字符串: {text}")
        return False
    
    # 檢查緩存
    cache_key = f"is_chinese_{text[:50]}"
    if cache_key in _language_cache:
        return _language_cache[cache_key]
    
    # 【使用者要求】修改日文假名檢查邏輯，以排除常用的間隔號（・, U+30FB）。
    has_japanese_kana = any(
        ('\u3040' <= char <= '\u309F' or  # 平假名
         '\u30A0' <= char <= '\u30FA' or  # 片假名 (U+30A0 到 U+30FA)
         '\u30FC' <= char <= '\u30FF')   # 片假名 (U+30FC 到 U+30FF)
        for char in text if char != '\u30FB' # 顯式排除間隔號 (U+30FB)
    )

    if has_japanese_kana:
        logging.info(f"檢測到日文假名 (已排除間隔號): {text[:50]}...，視為非中文")
        _language_cache[cache_key] = False
        return False

    # 【新增】檢查韓文 (Hangul Syllables)
    has_korean = any('\uac00' <= char <= '\ud7af' for char in text)
    if has_korean:
        logging.info(f"檢測到韓文: {text[:50]}...，視為非中文")
        _language_cache[cache_key] = False
        return False
        
    # 【新增】檢查日文疊字符號 々 (U+3005)
    if '々' in text:
        logging.info(f"檢測到日文疊字符號 '々': {text[:50]}...，視為非中文")
        _language_cache[cache_key] = False
        return False

    # 檢查是否有中文字符（漢字、標點等）
    # 此處的 \u3000-\u303f 範圍包含 U+3005 (々)，但上面已經先行排除
    has_chinese_char = any(
        ('\u4e00' <= char <= '\u9fff') or
        ('\u3000' <= char <= '\u303f') or
        char in '《》·'
        for char in text
    )
    
    # 使用 langid 檢測語言
    lang, _ = langid.classify(text)
    
    # 計算中文字符比例
    chinese_char_count = sum(1 for char in text if '\u4e00' <= char <= '\u9fff' or char in '《》·')
    total_char_count = len(text.strip())
    is_mostly_chinese = chinese_char_count / total_char_count > 0.3 if total_char_count > 0 else False

    # 僅當包含中文字符且語言為 'zh' 或中文字符比例高時視為中文
    # 注意：has_chinese_char 確保了必須有中文字符，所以純英文會被排除
    result = has_chinese_char and (lang == 'zh' or is_mostly_chinese or chinese_char_count > 0)
    
    # 特殊情況：如果全是英文和數字，langid 可能判為 en，has_chinese_char 為 False -> Result False (正確)
    # 如果是 "Andy Lau 劉德華"，langid 可能判為 en 或 zh，has_chinese_char True -> Result True (正確)
    
    logging.info(f"文本 '{text[:50]}...' {'檢測為中文' if result else f'非中文，lang={lang}'}")
    _language_cache[cache_key] = result
    return result

def is_english(text):
    """檢查文本是否為英文，包含備用 ASCII 檢查"""
    if not text or not isinstance(text, str):
        logging.debug(f"文本為空或非字符串: {text}")
        return False
    # 首先檢查 langid
    lang, _ = langid.classify(text)
    if lang == 'en':
        logging.info(f"文本 '{text[:50]}...' 檢測為英文，lang={lang}")
        return True
    # 備用檢查：如果文本僅包含 ASCII 字符（排除純數字），視為英文
    if text.strip().isascii() and not text.strip().isdigit():
        logging.info(f"文本 '{text[:50]}...' 通過 ASCII 檢查，視為英文")
        return True
    logging.debug(f"文本 '{text[:50]}...' 非英文，lang={lang}")
    return False

def is_valid_language(text):
    """檢查文本是否為中文或英文，確保排除日文等其他語言"""
    if not text or not isinstance(text, str):
        logging.debug(f"文本為空或非字符串: {text}")
        return False
    lang, _ = langid.classify(text)
    if lang == 'zh':
        # 額外檢查是否有中文字符，避免日文誤判
        has_chinese = any('\u4e00' <= char <= '\u9fff' for char in text)
        logging.info(f"文本 '{text[:50]}...' {'檢測為中文' if has_chinese else '無中文字符，視為無效語言'}，lang={lang}")
        return has_chinese
    is_english_result = lang == 'en' or (text.strip().isascii() and not text.strip().isdigit())
    logging.info(f"文本 '{text[:50]}...' {'檢測為英文' if is_english_result else '非英文，視為無效語言'}，lang={lang}")
    return is_english_result

def is_valid_zh_hk_title(text, field_type=None):
    """
    【修改】檢查 zh-HK 標題/名稱是否為有效的、完美的香港繁體。
    一個完美的標題/名稱必須是：
    1. 有效的繁體中文 (非簡體、非日文假名)。
    2. (如果不是人名) 不包含可被移除的泛用型前綴 (如 "第 X 集")。
    【修改 2025-09-19】將 is_person_field 替換為 field_type，以進行更精確的判斷。
    """
    if not text or not isinstance(text, str):
        logging.debug(f"zh-HK 標題/名稱為空或非字符串: {text}")
        return False

    # 【使用者要求】【修改 2025-09-19】將 is_person_field 替換為 field_type，以進行更精確的判斷。
    # 只有在欄位類型不是 'name' 時，才檢查泛用型標題。
    if field_type != 'name':
        if clean_generic_title(text) != text:
            logging.info(f"zh-HK 標題 '{text[:50]}...' 包含泛用前綴，視為不完美。")
            return False

    # 檢查 2: 語言和字體
    cache_key = f"is_valid_zh_hk_title_lang_{text[:50]}"
    if cache_key in _language_cache:
        return _language_cache[cache_key]

    has_japanese_kana = any('\u3040' <= char <= '\u309F' or '\u30A0' <= char <= '\u30FF' for char in text)
    if has_japanese_kana:
        logging.info(f"zh-HK 標題/名稱包含日文假名: {text[:50]}...，視為無效")
        _language_cache[cache_key] = False
        return False
    
    is_chinese_result = is_chinese(text)
    is_simplified_result = is_simplified_chinese(text)
    
    result = is_chinese_result and not is_simplified_result
    logging.info(f"zh-HK 標題/名稱 '{text[:50]}...' 語言有效性檢查: is_chinese={is_chinese_result}, is_simplified={is_simplified_result}, 結果={result}")
    
    _language_cache[cache_key] = result
    return result

def add_spaces_between_en_zh_num(text, item_type_for_log="文本", is_long_text=False):
    """
    在英文、漢字和數字之間添加空格，移除補充平面字符，處理連字號前後的空格，並保留換行符。
    【修改】增加對全形空格的處理，將其轉換為換行符，並移除 wiki 格式的備註 [數字]。
    【修改】在指定的半形標點符號 (; , . ?) 後面添加空格（如果不存在）。
    【新增】is_long_text 參數，僅在為 True 時才將標點 (。？！) 後的空格轉為換行。
    【使用者要求 2025-08-18】修正半形括號前後未正確加上空格的問題。例如 '字(字)' -> '字 (字)'。
    【使用者要求 2025-09-04】新增特殊處理，以避免在色情電影的英文數字編號 (如 'IPZZ-196') 周圍錯誤地添加或保留空格。
    【使用者要求 2025-09-19】修正換行符 (\r) 被移除的問題，統一將 \r 和 \r\n 轉換為 \n 以保留段落。
    【新增 2025-10-22】增加詳細日誌，記錄添加空格前的文本狀態，以便排查空格添加問題。
    """
    if not text or not isinstance(text, str):
        # 【新增日誌】記錄空或非字符串輸入
        logging.debug(f"add_spaces_between_en_zh_num ({item_type_for_log}): 輸入為空或非字符串: {text!r}")
        return text or ""

    original_text_for_log = str(text) # 保留原始輸入以供日誌記錄
    
    # 【新增】移除類似 [1], [15] 等 wiki 格式的備註
    text_processed = re.sub(r'\[\d+\]', '', original_text_for_log)
    
    # 【新增】僅在 is_long_text 為 True 時，將全形句號、問號、感嘆號後的任意空格轉換為換行符
    if is_long_text:
        # [ \t　] 匹配半形空格、Tab、或全形空格
        text_processed = re.sub(r'([。？！])([ \t　]+)', r'\1\n', text_processed)

    # 【修改】將剩餘的獨立全形空格轉換為標準換行符，以處理簡介中的格式問題。
    text_processed = text_processed.replace('　', '\n')

    # 移除 unicode 補充平面中的字符
    text_bmp_only = ''.join(c for c in text_processed if ord(c) <= 0xFFFF)
    
    # 【使用者要求 2025-09-19】為了保留所有類型的換行，先將所有換行符（\r\n 和 \r）統一標準化為 \n
    text_to_process = text_bmp_only.replace('\r\n', '\n').replace('\r', '\n')

    lines = text_to_process.split('\n')
    processed_lines = []
    for line_index, line in enumerate(lines):
        # 【新增日誌】記錄處理前的單行文本
        logging.debug(f"add_spaces_between_en_zh_num ({item_type_for_log}) - Line {line_index} Before Space Regex: '{line}'")

        # 【使用者要求 2025-09-04】新增保護機制，以正確處理色情電影編號 (如 IPZZ-196)。
        # 步驟 1: 預處理，將已存在的錯誤間距 (如 "IPZZ - 196") 修正為 "IPZZ-196"。
        #         這樣即使 API 源數據格式不佳，也能被正確識別和保護。
        line = re.sub(r'\b([A-Z]{2,})\s*-\s*(\d+)\b', r'\1-\2', line)

        # 步驟 2: 使用佔位符保護修正後的編號，防止後續的正則表達式再次錯誤地添加空格。
        porn_code_pattern = re.compile(r'\b[A-Z]{2,}-\d+\b')
        placeholders = {}
        
        def placeholder_replacer(match):
            code = match.group(0)
            # 使用一個不會與文本衝突的唯一佔位符
            placeholder = f"__SPECIAL_CODE_PLACEHOLDER_{len(placeholders)}__"
            placeholders[placeholder] = code
            return placeholder

        line_with_placeholders = porn_code_pattern.sub(placeholder_replacer, line)
        
        # 【使用者要求 2025-09-19】修改此處的正則表達式，移除 \r，因為換行已在 split('\n') 之前處理。
        # 將多個空白符合併為一個半形空格
        line_processed = re.sub(r'[ \t\f\v]+', ' ', line_with_placeholders)
        
        # --- 開始添加空格的核心邏輯 ---
        # 在中英文、中數字之間加上空格
        line_processed = re.sub(r'([a-zA-Z0-9])([\u4e00-\u9fff])', r'\1 \2', line_processed)
        line_processed = re.sub(r'([\u4e00-\u9fff])([a-zA-Z0-9])', r'\1 \2', line_processed)

        # --- 【使用者要求 2025-08-18】處理半形括號的間距 ---
        # 目標：修正 '蘇格蘭：德爺再戰自然之子(上)' 為 '蘇格蘭：德爺再戰自然之子 (上)'
        # 規則：在一個漢字/英文/數字字元和一個半形左括號 '(' 之間加上空格。
        line_processed = re.sub(r'([\u4e00-\u9fffA-Za-z0-9])(\()', r'\1 \2', line_processed)
        # 規則：在一個半形右括號 ')' 和一個漢字/英文/數字字元之間加上空格。
        line_processed = re.sub(r'(\))([\u4e00-\u9fffA-Za-z0-9])', r'\1 \2', line_processed)

        # --- 【新增】處理標點符號間距的邏輯 ---
        # 步驟 1: 在指定的半形標點符號 (; , . ?) 後方，如果沒有跟著空白符，就加上一個空格。
        line_processed = re.sub(r'([;,.?!])(?!\s)', r'\1 ', line_processed)
        
        # 步驟 2: 修正上一步可能錯誤影響的小數 (例如 "1. 5" -> "1.5")。
        line_processed = re.sub(r'(\d)\. (\d)', r'\1.\2', line_processed)

        # 步驟 3: 修正上一步可能錯誤影響的檔案名或網域名稱 (例如 "example. com" -> "example.com")。
        line_processed = re.sub(r'([a-zA-Z])\. ([a-zA-Z])', r'\1.\2', line_processed)
        # --- 結束新增邏輯 ---

        # 在連字號前後加上空格
        line_processed = re.sub(r'([\u4e00-\u9fffA-Za-z0-9])([-\–—])([\u4e00-\u9fffA-Za-z0-9])', r'\1 \2 \3', line_processed)
        # 修正：數字之間的連字號或斜線不加空格 (如日期 2026-05-13, 2026/05/13)
        line_processed = re.sub(r'(\d) ([-/]) (\d)', r'\1\2\3', line_processed)
        
        # 【使用者要求 2025-09-04】步驟 3: 在所有空格處理完成後，將編號還原。
        for placeholder, original_code in placeholders.items():
            line_processed = line_processed.replace(placeholder, original_code)
            
        # 【新增日誌】記錄處理後的單行文本
        logging.debug(f"add_spaces_between_en_zh_num ({item_type_for_log}) - Line {line_index} After Space Regex: '{line_processed}'")
        processed_lines.append(line_processed.strip())

    final_text_joined = "\n".join(processed_lines)
    final_text = final_text_joined.strip() # 移除最終結果前後的空白，但保留行間的換行

    # 為了日誌簡潔，只比較最終結果與原始輸入（去除首尾空白）
    original_text_stripped = original_text_for_log.strip()

    if final_text != original_text_stripped: # 比較去除首尾空白後的結果
        original_snippet = original_text_stripped[:70].replace('\n', '[NL]').replace('\r', '[CR]')
        final_snippet = final_text[:70].replace('\n', '[NL]')
        log_msg = f"格式化文本 ({item_type_for_log}): 原文片段='{original_snippet}', 處理後='{final_snippet}'"
        if is_long_text:
            log_msg += " (長文本模式)"
        # 根據是否真的添加了空格來調整日誌級別
        if re.search(r'[a-zA-Z0-9][\u4e00-\u9fff]', original_text_stripped) or re.search(r'[\u4e00-\u9fff][a-zA-Z0-9]', original_text_stripped):
             logging.info(log_msg) # 如果原文可能需要加空格，用 INFO
        else:
             logging.debug(log_msg) # 其他情況（如僅是標點空格或換行變化）用 DEBUG
    else:
        # 如果最終結果與去除首尾空白的原始輸入相同，用 DEBUG 級別記錄
        final_snippet = final_text[:70].replace('\n', '[NL]')
        logging.debug(f"格式化文本 ({item_type_for_log}): 文本經處理後無實質變化: '{final_snippet}'")

    return final_text

def convert_runtime_to_minutes(runtime):
    """將運行時間字符串轉換為分鐘並返回整數，支持多種格式，秒數 > 30 秒進位為 1 分鐘"""
    if not runtime:
        logging.debug("運行時長為空")
        return None
    if isinstance(runtime, (int, float)):
        logging.debug(f"運行時長為數字: {runtime}")
        return int(runtime)
    
    runtime = str(runtime).strip().lower()
    if not runtime:
        logging.debug("運行時長為空字符串")
        return None

    # 移除附加文本（如 '• 紀錄片'）
    runtime = re.sub(r'•.*$', '', runtime).strip()

    # 支持的格式：1hr10min, 1h 10m, 1小時10分鐘, 1 hour 10 minutes, 90 min, 1.5h, 90, 1 小時, 19 分鐘, 40 秒
    try:
        # 純數字（假設為分鐘）
        if runtime.replace('.', '', 1).isdigit():
            result = int(float(runtime))
            logging.debug(f"運行時長解析為純數字: {runtime} -> {result} 分鐘")
            return result if result > 0 else None

        # 小時、分鐘和秒格式（1 小時,19 分鐘,40 秒 或 1hr10min40sec）
        pattern = r'^(?:(\d+)\s*(?:hr|h|hour|小時))?(?:\s*,?\s*)?(?:(\d+)\s*(?:min|m|minute|分鐘))?(?:\s*,?\s*)?(?:(\d+)\s*(?:sec|s|second|秒))?$'
        match = re.match(pattern, runtime, re.IGNORECASE)
        if match:
            hours = int(match.group(1) or 0)
            minutes = int(match.group(2) or 0)
            seconds = int(match.group(3) or 0)
            # 秒數 > 30 秒進位為 1 分鐘
            minutes += 1 if seconds > 30 else 0
            result = hours * 60 + minutes
            logging.debug(f"運行時長解析為小時、分鐘和秒: {runtime} -> {result} 分鐘")
            return result if result > 0 else None

        # 小時和分鐘格式（1hr10min, 1h 10m, 1 hour 10 minutes）
        hours_match = re.search(r'(\d+)\s*(?:hr|h|hour|小時)', runtime, re.IGNORECASE)
        minutes_match = re.search(r'(\d+)\s*(?:min|m|minute|分鐘)', runtime, re.IGNORECASE)
        
        hours = int(hours_match.group(1)) if hours_match else 0
        minutes = int(minutes_match.group(1)) if minutes_match else 0
        
        if hours or minutes:
            result = hours * 60 + minutes
            logging.debug(f"運行時長解析為小時和分鐘: {runtime} -> {result} 分鐘")
            return result if result > 0 else None

        # 小時的小數格式（1.5h）
        hours_decimal_match = re.search(r'(\d+\.\d+)\s*(?:hr|h|hour|小時)', runtime, re.IGNORECASE)
        if hours_decimal_match:
            hours_decimal = float(hours_decimal_match.group(1))
            result = int(hours_decimal * 60)
            logging.debug(f"運行時長解析為小時小數: {runtime} -> {result} 分鐘")
            return result if result > 0 else None

        # 僅分鐘（90 min）
        minutes_only_match = re.search(r'(\d+)\s*(?:min|m|minute|分鐘)', runtime, re.IGNORECASE)
        if minutes_only_match:
            result = int(minutes_only_match.group(1))
            logging.debug(f"運行時長解析為僅分鐘: {runtime} -> {result} 分鐘")
            return result if result > 0 else None

        logging.warning(f"無法解析運行時長格式: {runtime}")
        return None
    except (ValueError, AttributeError) as e:
        logging.error(f"轉換運行時長失敗: {runtime}, 錯誤: {e}")
        return None

def set_utils_interrupt_callback(callback):
    """設置一個回調函數，用於檢查程序是否已被用戶中斷。"""
    global _program_interrupted_callback_for_utils
    _program_interrupted_callback_for_utils = callback

def reload_dictionaries_and_converters(dict_dir='opencc_data/dictionary'):
    """
    【修改】重新載入所有 OpenCC 轉換器和手動字典檔。
    【核心修正 2025-09-16】更新預設路徑 `dict_dir` 為 `opencc_data/dictionary`，以匹配新的統一化字典管理策略。
    """
    logging.info("重新載入 OpenCC 轉換器、本地字典和自定義規則...")
    # 再次調用初始化，以確保在 start.py 的補丁生效後能正確加載
    initialize_opencc_converters()
    # 從統一的路徑載入所有手動字典
    load_manual_dictionaries(dict_dir)
    load_protected_names()
    clear_tw_hk_name_cache()
    try:
        from others import ai_s2t
        ai_s2t.load_triggers(force=True)  # 重新讀取 tw_hk_name.xlsx「AI Convert」頁
        ai_s2t.load_no_change_words(force=True)  # 【2026-11 R1】「AI No Change」不准改清單
    except Exception as e:
        logging.warning(f"[AI簡轉繁] 重新載入觸發字失敗: {e}")

def _uploader_pause_requested():
    """【2026-11-14】上傳執行緒專用：「暫停上傳」或「暫停」旗標存在。extractor 等其他執行緒一律 False。"""
    import threading
    if threading.current_thread().name != 'UploaderThread':
        return False
    return any(os.path.exists(f) for f in (PAUSE_FLAG, UPLOAD_PAUSE_FLAG))


def _wait_if_uploader_paused():
    """【2026-11-14】上傳執行緒喺頁面步驟之間即時暫停；繼續後回傳 True（曾暫停）。暫停期間按停止照樣中斷。"""
    if not _uploader_pause_requested():
        return False
    wait_if_paused(logging.getLogger('importer'), 'Uploader', (PAUSE_FLAG, UPLOAD_PAUSE_FLAG))
    if _program_interrupted_callback_for_utils():
        raise KeyboardInterrupt("Program interrupted by user signal while upload paused.")
    return True


def _limit_wait(seconds):
    """限流等待：每秒檢查停止同暫停。返回 True = 因暫停而提早結束（已等到繼續）。"""
    start = time.time()
    while time.time() - start < seconds:
        if _program_interrupted_callback_for_utils():
            logging.warning("check_wait_and_try_again: 在等待期間檢測到程序中斷信號。")
            raise KeyboardInterrupt("Program interrupted by user signal during wait period.")
        if _uploader_pause_requested():
            logging.info("check_wait_and_try_again: 限流等待期間按咗暫停，停止倒數。")
            _wait_if_uploader_paused()
            return True
        time.sleep(max(0.001, min(1, seconds - (time.time() - start))))
    return False


def check_wait_and_try_again(driver, current_url=None, video_url=None):
    """
    【修改】檢查是否有 'Wait and try again' 訊息。
    【使用者要求】新增強大的重試機制，在發生網路錯誤時無限重試。
    如果找到，將使用遞增的等待時間後重新載入頁面，並無限重試直到訊息消失。
    
    返回:
        True: 如果發現並處理了 'Wait and try again' 訊息 (表示曾經被限流)。
        False: 如果頁面正常，沒有發現該訊息。
    """
    retry_count = 0
    encountered_limit = False # 標記是否曾經遇到限流

    while True: # 實現無限重試
        # 在每次循環開始時，立即檢查是否需要中斷
        if _program_interrupted_callback_for_utils():
            logging.warning("check_wait_and_try_again: 檢測到程序中斷信號，將引發 KeyboardInterrupt。")
            raise KeyboardInterrupt("Program interrupted by user signal during check_wait_and_try_again.")
        # 【2026-11-14】上傳執行緒：按咗暫停就喺呢個頁面步驟停低，繼續後由呢度接住做
        _wait_if_uploader_paused()

        retry_count += 1
        # 只有第一次檢查時記錄 debug，避免刷屏，除非真的遇到問題
        if retry_count == 1:
            logging.debug("檢查是否有 'Wait and try again' 訊息")

        try:
            # 檢查 "Wait and try again" 訊息是否存在
            # 使用極短的 timeout 來快速檢查，避免正常情況下拖慢速度
            # 注意：這裡假設 caller 在調用前已經確保 body 載入
            WebDriverWait(driver, 0.5).until(
                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Wait and try again')]"))
            )

            # --- 找到訊息後的處理 ---
            encountered_limit = True # 標記遇到了限流

            if _program_interrupted_callback_for_utils():
                logging.warning("check_wait_and_try_again: 在檢測到 'Wait' 訊息後，檢測到程序中斷信號。")
                raise KeyboardInterrupt("Program interrupted by user signal after detecting 'Wait' message.")

            # 計算等待時間：從1分鐘開始，每次遞增1分鐘，上限為10分鐘
            wait_minutes = min(retry_count, 10)
            wait_time_seconds = wait_minutes * 60

            logging.warning(f"發現 'Wait and try again' 訊息。將等待 {wait_minutes} 分鐘後重新載入頁面 (重試第 {retry_count} 次)...")

            # 在等待期間，定期檢查中斷信號同暫停【2026-11-14】暫停 → 停止倒數；繼續後即刻重新載入，等待由 1 分鐘重新計
            if _limit_wait(wait_time_seconds):
                retry_count = 0

            # 重新載入頁面
            page_to_reload = current_url if current_url else driver.current_url
            logging.info(f"限流等待結束，重新載入頁面: {page_to_reload}")
            if page_to_reload:
                driver.get(page_to_reload)
            else:
                logging.warning("無法獲取 current_url 進行重載，將嘗試刷新當前頁面。")
                driver.refresh()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

        except TimeoutException:
            # --- 未找到訊息 (正常情況) ---
            if retry_count > 1:
                logging.info(f"已恢復正常 ('Wait and try again' 消失)。")
            return encountered_limit # 回傳是否曾經遇到限流

        except (Exception) as e: # 捕獲 WebDriverException 等
             # --- 處理檢查過程中發生的其他異常 ---
            if _program_interrupted_callback_for_utils():
                raise KeyboardInterrupt(f"Program interrupted: {e}") from e

            from others.chrome_process import is_dead_browser_error, ensure_browser_window
            if is_dead_browser_error(e):
                recovered = False
                try:
                    recovered = ensure_browser_window(driver)
                except Exception:
                    recovered = False
                if not recovered:
                    logging.error(f"檢查限流時發現瀏覽器視窗已關閉，停止重試並交由上層重啟瀏覽器: {e}")
                    raise
                logging.warning("檢查限流時視窗失效，已切換/重建視窗，繼續檢查...")
                continue

            # 發生異常時也採用遞增等待，視同被阻擋
            encountered_limit = True
            wait_minutes_on_exc = min(retry_count, 10)
            logging.error(f"檢查限流時發生錯誤: {e}。視為潛在阻擋，等待 {wait_minutes_on_exc} 分鐘後重試...")
            if _limit_wait(wait_minutes_on_exc * 60):   # 【2026-11-14】以前係 time.sleep，停止 / 暫停都要等完
                retry_count = 0

            # 嘗試重新載入
            try:
                page_to_reload = current_url if current_url else driver.current_url
                if page_to_reload: driver.get(page_to_reload)
                else: driver.refresh()
            except Exception:
                pass