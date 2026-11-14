import json
import os
import re
import logging
import sys
import contextlib
import importlib.util
import subprocess
import traceback
import pandas as pd
import requests
import signal
import openpyxl
import numpy as np
import gc  # 【新增】導入 gc 模組，用於強制垃圾回收以釋放被佔用的文件句柄
import gzip # 【新增】用於壓縮日誌
import shutil # 【新增】用於檔案操作
from logging.handlers import RotatingFileHandler # 【新增】導入 RotatingFileHandler 用於日誌輪替
from inputimeout import inputimeout, TimeoutOccurred
from selenium.common.exceptions import WebDriverException # 【新增】導入 WebDriverException 以處理瀏覽器崩潰
from selenium.webdriver.common.by import By
from others.credentials import load_configs, get_tmdb_credentials, is_tmdb_logged_in, get_disneyplus_credentials
from others.excel_handler import (
    check_create_excel, clear_sheet, check_and_release_excel_control, release_excel_control,
    save_unavailable_video_info_to_excel, save_skipped_identical_video_info_to_excel, clean_url,
    clean_sheet_data, replace_tw_hk_names, check_create_bulk_excel, update_done_value,
    save_no_chinese_content_info_to_excel,
    save_english_site_info_to_excel,
    save_skipped_identical_to_bulk_excel,
    update_done_value_with_lock_check,
    update_done_value_fast,
    get_pending_done_value,
    pending_done_count,
    start_batch_update_worker,
    stop_batch_update_worker,
    flush_pending_done_updates,
    write_excel_atomic,
)
from others.chrome_process import check_and_update_dependencies, get_browser_driver, cleanup_browser, ensure_browser_window, is_dead_browser_error
from others.clear_titles import clear_title, process_titles
from others.analysis_excel import is_generic_title
from others import utils # 導入整個 utils 模塊
from importers.tmdb_uploader import upload_to_tmdb, upload_person_to_tmdb, extract_tv_show_id, extract_movie_id, extract_person_id
from extractors.tmdb_extractor import main as extract_tmdb_data_selenium
from extractors.tmdb_extractor_api import main as extract_tmdb_data_api, extract_person_details
from extractors.netflix_extractor import extract_netflix_episodes
from extractors.primevideo_extractor import extract_primevideo_data
from extractors.appletv_extractor import extract_appletv_data
from extractors.disneyplus_extractor import login_to_disneyplus, extract_disneyplus_data, is_disneyplus_logged_in
from extractors.viu_extractor import extract_viu_data
from extractors.viutv_extractor import extract_viutv_data
import threading
import queue
import time # 【新增】導入 time 模組以使用 time.sleep()

# 【新增】全局變數，用於執行緒間的同步和資料交換
EXCEL_LOCK = threading.Lock()
DETAILS_CACHE_DIR = os.path.join('data', 'details_cache')

# 【共享緩存】所有 extractor threads 共用同一份 in-memory DataFrame
# 由 background refresher 定時更新，threads 讀取時完全不需要 EXCEL_LOCK
_shared_bulk_df = None
_shared_bulk_df_lock = threading.Lock()
_shared_cache_refresher_running = False
_shared_cache_refresher_thread = None

def _refresh_shared_cache(excel_lock, file_path, retry_interval=5, max_retries=3):
    """
    從 Excel 讀取最新資料並更新共享緩存。
    使用檔案鎖（.lock 檔）保護跨 process 的讀取，避免與其他程式（如 tmdb_merged_runner）衝突。
    失敗後最多重試 max_retries 次，每次間隔 retry_interval 秒。
    """
    global _shared_bulk_df
    lock_file = file_path + '.lock'

    for attempt in range(max_retries):
        # 等待跨 process 的檔案鎖釋放（其他程式寫入時會建立此鎖）
        waited = 0
        while os.path.exists(lock_file) and waited < 10:
            time.sleep(0.2)
            waited += 0.2

        try:
            with excel_lock:  # 同 process 內的 thread 鎖
                # 確保檔案與 Bulk sheet 存在，否則自動建立（防禦性）
                if not os.path.exists(file_path):
                    check_create_bulk_excel(file_path)
                else:
                    try:
                        # 【2026-11 I1】read_only：只需 sheet 名稱；完整載入大檔（如 24 萬行 old_Bulk）要 6-7 秒並長時間佔用 EXCEL_LOCK
                        wb = openpyxl.load_workbook(file_path, read_only=True)
                        if utils.BULK_SHEET not in wb.sheetnames:
                            wb.close()
                            check_create_bulk_excel(file_path)
                        else:
                            wb.close()
                    except Exception:
                        check_create_bulk_excel(file_path)
                sheet = utils.BULK_SHEET
                df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
            with _shared_bulk_df_lock:
                if sheet != utils.BULK_SHEET:   # 【2026-11】讀取期間已轉工作表 → 棄用舊結果
                    logging.info(f"[SharedCache] 已轉去工作表 '{utils.BULK_SHEET}'，棄用 '{sheet}' 嘅讀取結果。")
                    return True
                _shared_bulk_df = df
            logging.info(f"[SharedCache] 已更新共享緩存，共 {len(df)} 行。")
            return True
        except Exception as e:
            if attempt < max_retries - 1:
                logging.warning(f"[SharedCache] 更新緩存失敗（第 {attempt+1}/{max_retries} 次），{retry_interval} 秒後重試: {e}")
                time.sleep(retry_interval)
            else:
                logging.error(f"[SharedCache] 更新緩存失敗，已重試 {max_retries} 次，放棄本次更新: {e}")
    return False

def start_shared_cache_refresher(excel_lock, file_path='data/bulk_videos.xlsx', interval=30):
    """
    啟動背景執行緒，定時將 Excel 讀入共享記憶體緩存。
    extractor threads 直接從緩存讀取，無需持有 EXCEL_LOCK，消除讀寫競爭。
    interval: 更新間隔（秒），用家修改 Excel 後最多等待此時間生效。
    """
    global _shared_cache_refresher_running, _shared_cache_refresher_thread

    def _worker():
        _refresh_shared_cache(excel_lock, file_path)
        while _shared_cache_refresher_running and not utils._PROGRAM_INTERRUPTED:
            for _ in range(interval * 10):
                if not _shared_cache_refresher_running or utils._PROGRAM_INTERRUPTED:
                    break
                time.sleep(0.1)
            if _shared_cache_refresher_running and not utils._PROGRAM_INTERRUPTED:
                _refresh_shared_cache(excel_lock, file_path)

    _shared_cache_refresher_running = True
    _shared_cache_refresher_thread = threading.Thread(
        target=_worker, name="SharedCacheRefresher", daemon=True
    )
    _shared_cache_refresher_thread.start()

    # 等待初始讀取完成，最多 60 秒
    waited = 0
    while _shared_bulk_df is None and waited < 60:
        time.sleep(0.2)
        waited += 0.2
    if _shared_bulk_df is None:
        raise RuntimeError("[SharedCache] 初始讀取 bulk_videos.xlsx 超時，程式無法啟動。")
    logging.info(f"[SharedCache] 初始化完成，每 {interval} 秒自動更新。")

def stop_shared_cache_refresher():
    """停止共享緩存更新執行緒。"""
    global _shared_cache_refresher_running
    _shared_cache_refresher_running = False

def _set_shared_cache(df):
    """【2026-11 F2】以剛讀取的 Excel 內容更新共享緩存，令「發現新項目」與「extractor 取資料」使用同一份資料。"""
    global _shared_bulk_df
    if df is None:
        return
    with _shared_bulk_df_lock:
        _shared_bulk_df = df

def get_row_from_shared_cache(row_index):
    """從共享緩存取得指定行，返回 Series 或 None。"""
    with _shared_bulk_df_lock:
        if _shared_bulk_df is not None and row_index < len(_shared_bulk_df):
            return _shared_bulk_df.iloc[row_index].copy()
    return None

# 設置日誌配置
# 【修改】繼承 RotatingFileHandler 以支援檔案大小限制和輪替，flush 即時寫入
class ImmediateRotatingFileHandler(RotatingFileHandler):
    # 【使用者要求】修改日Log Handler，確保日誌即時寫入檔案並支援 1GB 上限輪替
    # 1. 繼承 RotatingFileHandler 實現 maxBytes 輪替。
    # 2. emit 中 flush 實現即時寫入（2026-11 已移除 fsync）。
    def __init__(self, filename, mode='a', maxBytes=0, backupCount=0, encoding='utf-8', delay=False):
        super().__init__(filename, mode, maxBytes, backupCount, encoding, delay)

    def emit(self, record):
        try:
            # 調用父類 (RotatingFileHandler) 的 emit，它會處理寫入和輪替邏輯
            super().emit(record)
            
            # 執行即時寫入邏輯：flush 後檔案內容即時可見。
            # 【2026-11】移除每行 os.fsync（強制寫入實體磁碟），避免多執行緒大量日誌時拖慢程式；
            # 只有在突然斷電時才可能遺失最後幾行。
            self.flush()
        except Exception:
            self.handleError(record)

class UTF8StreamHandler(logging.StreamHandler):
# ... (此 class 保持不變) ...
# ... (您提供的程式碼中，此 class 內容是完整的，故保留) ...
    def __init__(self, stream=None):
        super().__init__(stream)
        if self.stream is None:
            self.stream = sys.stdout
        if hasattr(self.stream, 'reconfigure'):
            self.stream.reconfigure(encoding='utf-8')
    def emit(self, record):
        try:
            super().emit(record)
        except UnicodeEncodeError:
            msg = record.getMessage()
            msg = msg.encode('utf-8', errors='replace').decode('utf-8')
            self.stream.write(msg + self.terminator)
            self.flush()

class ThreadNameFilter(logging.Filter):
# ... (此 class 保持不變) ...
# ... (您提供的程式碼中，此 class 內容是完整的，故保留) ...
    def __init__(self, thread_names, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.thread_names = thread_names

    def filter(self, record):
        # 檢查當前日誌記錄是否來自指定的執行緒之一
        return record.threadName in self.thread_names

# 確保 log 目錄存在
os.makedirs('log', exist_ok=True)

def setup_loggers():
    """
    【修改】統一設定日誌記錄器，使用執行緒名稱過濾器將日誌正確路由。
    # 觸發條件：日誌上限 1GB，自動輪替並壓縮為 .gz。
    """
    import gzip
    import shutil
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    log_dir = 'log'
    # 清除舊 log 前：把「AI 被拒字詞」補錄到 data/ai_s2t_cache.db（重複匯入不會重複計數）。
    # 【2026-11 J1】必須在清空 handlers 之前執行：補錄途中的 logging.info 若遇上 root 無 handler，
    # Python 會自動 basicConfig 一個寫去 sys.stderr 的 handler；start.py 匯入時 stderr 是暫時的 devnull，
    # 之後被 close → 每句 log 都報「I/O operation on closed file」。放在這裡，該 handler 會被下面清走。
    if os.path.exists(log_dir):
        try:
            from others import ai_s2t
            n_rej = ai_s2t.import_rejected_from_logs(log_dir)
            if n_rej:
                print(f"[Log Setup] 已由舊 log 補錄 {n_rej} 句 AI 被拒字詞（監控面板「AI 被拒用詞」分頁）。")
        except Exception as e:
            print(f"[Log Setup] 補錄 AI 被拒字詞失敗（不影響運作）: {e}")

    # --- 清理現有的 handlers ---
    root_logger = logging.getLogger()
    if root_logger.hasHandlers():
        print("[Log Setup] 已清除根日誌記錄器的現有 handlers。")
        root_logger.handlers.clear()

    # --- 程式啟動時清除所有日誌 ---
    if os.path.exists(log_dir):
        print(f"[Log Setup] 正在清除所有日誌檔案...")
        for filename in os.listdir(log_dir):
            file_path = os.path.join(log_dir, filename)
            try:
                os.remove(file_path)
            except OSError as e:
                print(f"[Log Setup] 無法刪除舊日誌 {filename}: {e}")

    # --- 定義壓縮函式與命名規則 ---
    def compress_namer(name):
        """命名規則：確保輪替後的檔名正確附加 .gz"""
        return name + ".gz"

    def compress_rotator(source, dest):
        """
        執行壓縮：將 source 壓縮到 dest。
        # 修正：針對 Windows 檔案鎖定進行最佳化，處理 .log.1 -> .log.1.gz 及序列位移。
        """
        try:
            # 如果目標是 .gz 檔案且來源檔案存在
            if dest.endswith('.gz') and os.path.exists(source):
                # 如果來源也是 .gz (這是序列位移，例如 .1.gz -> .2.gz)，直接重新命名
                if source.endswith('.gz'):
                    if os.path.exists(dest):
                        os.remove(dest)
                    os.rename(source, dest)
                else:
                    # 這是新產生的 .1 日誌，執行壓縮
                    with open(source, 'rb') as f_in:
                        with gzip.open(dest, 'wb') as f_out:
                            shutil.copyfileobj(f_in, f_out)
                    try:
                        # 壓縮後嘗試刪除原始檔案
                        os.chmod(source, 0o777)
                        os.remove(source)
                    except Exception:
                        pass
        except Exception as e:
            print(f"[Log Rotate] 處理日誌輪替失敗 (source={source}, dest={dest}): {e}")

    # --- 建立 Handlers ---
    console_handler = UTF8StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)

    MAX_BYTES = 1 * 1024 * 1024 * 1024  # 1GB
    BACKUP_COUNT = 5 

    # Importer 檔案 Handler
    importer_fh = ImmediateRotatingFileHandler(
        'log/importer.log', mode='a', encoding='utf-8', 
        maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT
    )
    importer_fh.setFormatter(formatter)
    # 背景執行緒（Done 寫入、共享緩存刷新）亦寫入 importer.log；新增具名執行緒時記得加入此名單
    importer_fh.addFilter(ThreadNameFilter([
        'MainThread', 'UploaderThread',
        'BatchUpdateThread', 'DoneUpdateThread', 'SharedCacheRefresher',
    ]))
    importer_fh.namer = compress_namer
    importer_fh.rotator = compress_rotator

    # Extractor 檔案 Handlers (5 條執行緒，各自寫入獨立的 log 檔)
    extractor_fhs = []
    for _i in range(5):
        _fh = ImmediateRotatingFileHandler(
            f'log/extractor_{_i}.log', mode='a', encoding='utf-8',
            maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT
        )
        _fh.setFormatter(formatter)
        _fh.addFilter(ThreadNameFilter([f'ExtractorThread_{_i}']))
        _fh.namer = compress_namer
        _fh.rotator = compress_rotator
        extractor_fhs.append(_fh)

    # --- 設定 Root Logger ---
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(console_handler)
    root_logger.addHandler(importer_fh)
    for _fh in extractor_fhs:
        root_logger.addHandler(_fh)

    logging.getLogger('importer').setLevel(logging.INFO)
    logging.getLogger('extractor').setLevel(logging.INFO)

    # 靜音第三方 logger
    logging.getLogger('selenium').setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)
    logging.getLogger('webdriver_manager').setLevel(logging.WARNING)
    logging.getLogger('selenium.webdriver').setLevel(logging.WARNING)

    root_logger.info("日誌記錄器已成功設定 (1GB 上限, 自動輪替並壓縮 .gz)。")

# --- 程式開始時執行日誌設定 ---
setup_loggers()

# 取得 logger 以便在模組中使用
importer_logger = logging.getLogger('importer')
extractor_logger = logging.getLogger('extractor')

# logged_in = False 
tmdb_login_attempted_this_session = False 

# 全局標誌和信號處理已移至 utils.py，這裡僅作為初始化時的引用
# _PROGRAM_INTERRUPTED = False  <-- Removed

def ensure_tmdb_accessible(driver):
    """
    【新增】確保 TMDB 可訪問的輔助函數。
    執行流程：
    1. 確保瀏覽器位於 TMDB 網域（若否則前往首頁）。
    2. 調用 utils.check_wait_and_try_again 檢查並處理限流（無限重試）。
    3. 嘗試點擊 "Accept Cookies" 按鈕（如果存在），以免遮擋後續操作。
    視窗已關閉時必須往上拋，讓呼叫端重啟瀏覽器，不可當成非致命錯誤吞掉。
    """
    log = logging.getLogger('importer')
    try:
        if not ensure_browser_window(driver):
            raise WebDriverException("Browser has no usable window")

        # 1. 確保在 TMDB 網域
        if "themoviedb.org" not in (driver.current_url or ""):
            log.info("[Access Check] 正在前往 TMDB 首頁以檢查連線狀態...")
            driver.get("https://www.themoviedb.org/")
        
        # 2. 檢查限流 (這是核心，會無限重試直到成功)
        # 注意：check_wait_and_try_again 會處理 "Wait and try again" 畫面並自動重新載入
        utils.check_wait_and_try_again(driver)
        
        # 3. 嘗試點擊 Cookie 同意 (如果有)
        try:
            # TMDB 常見的 Cookie 同意按鈕 ID
            accept_btns = driver.find_elements(By.ID, "onetrust-accept-btn-handler")
            if accept_btns and accept_btns[0].is_displayed():
                accept_btns[0].click()
                log.info("[Access Check] 已自動點擊 TMDB Cookie 同意按鈕。")
                time.sleep(1) # 稍作等待讓遮罩消失
        except Exception:
            # 點擊 Cookie 失敗不應阻擋主流程
            pass
            
    except Exception as e:
        if is_dead_browser_error(e) or isinstance(e, WebDriverException):
            log.error(f"[Access Check] 瀏覽器視窗已失效，交由上層重啟: {e}")
            raise
        log.warning(f"[Access Check] 檢查 TMDB 可訪問性時發生非致命錯誤: {e}")

def is_valid_video_url(url):
    if not isinstance(url, str) or pd.isna(url):
        return False
    valid_domains = ["netflix.com", "primevideo.com", "amazon.com", "appletv.com", "tv.apple.com", "disneyplus.com", "viu.com", "viu.tv"]
    return any(domain in url for domain in valid_domains)

def get_user_input_with_timeout(prompt, timeout=120):
    """獲取用戶輸入，支援短輸入（excel(e)/eof(f)），已移除超時功能"""
    importer_logger.debug(f"獲取用戶輸入，提示: {prompt}")

    # 移除超時功能，直接等待用戶輸入
    # 【2026-11 L1】等待輸入期間按 Ctrl+C → 立即安全結束（含關閉瀏覽器）
    user_input = utils.interruptible_input(prompt).lower().strip()
    importer_logger.debug(f"使用標準 input 獲取輸入: {user_input}")

    # 處理 bulk_mode 中的短輸入
    if "Video site 的 URL" in prompt:
        if user_input == 'e':
            importer_logger.info("用戶輸入短命令 'e'，解析為 'excel'")
            return 'excel'
        elif user_input == 'f':
            importer_logger.info("用戶輸入短命令 'f'，解析為 'eof'")
            return 'eof'
        elif user_input in ['excel', 'eof']:
            importer_logger.info(f"用戶輸入完整命令: {user_input}")
            return user_input
        importer_logger.info(f"用戶輸入其他值: {user_input}")
        return user_input

    # 處理 TMDB URL 輸入
    if "TMDB 的 URL" in prompt:
        if user_input == 'f':
            importer_logger.info("用戶輸入短命令 'f'，解析為 'eof'")
            return 'eof'
        importer_logger.info(f"用戶輸入 TMDB URL: {user_input}")
        return user_input

    importer_logger.debug(f"返回默認輸入: {user_input or 'ask'}")
    return user_input or 'ask'

def ensure_all_sheets_exist(file_name):
    required_sheets = {
        'Title': ['TV Show Title', 'TV Show Description'],
        'Seasons': ['Season Name', 'Season Number', 'Season Description'],
        'Episodes': ['Season Number', 'Episode Number', 'Episode Runtime', 'Episode Title', 'Episode Description'],
        'Movies': ['Movie Title', 'Movie Description', 'Runtime'],
        'People': ['Person URL', 'Name', 'Role', 'Biography']
    }
    max_retries = 3
    for attempt in range(max_retries):
        try:
            if not os.path.exists(file_name):
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    for sheet_name_iter, columns in required_sheets.items():
                        pd.DataFrame(columns=columns).to_excel(writer, sheet_name=sheet_name_iter, index=False)
                logging.info(f"已創建新的 {file_name}，包含工作表: {list(required_sheets.keys())}")
                return
            workbook = openpyxl.load_workbook(file_name)
            existing_sheets = workbook.sheetnames
            sheets_updated = False
            for sheet_name_iter, columns in required_sheets.items():
                if sheet_name_iter not in existing_sheets:
                    ws = workbook.create_sheet(sheet_name_iter)
                    ws.append(columns)
                    logging.info(f"工作表 {sheet_name_iter} 不存在於 {file_name}，已創建。")
                    sheets_updated = True
                elif sheet_name_iter == 'People':
                    ws = workbook[sheet_name_iter]
                    header = [cell.value for cell in ws[1]]
                    if 'Biography' not in header:
                        ws.cell(row=1, column=len(header) + 1, value='Biography')
                        logging.info(f"已為 {file_name} 的 People sheet 添加 Biography 欄位")
                        sheets_updated = True
            if sheets_updated:
                workbook.save(file_name)
            logging.info(f"成功確保 {file_name} 中所有工作表存在且結構正確")
            return
        except PermissionError:
            if attempt < max_retries - 1:
                logging.warning(f"無法訪問 {file_name}，可能正被其他程序占用，重試 {attempt + 1}/{max_retries}")
                input(f"請關閉 {file_name}，然後按 Enter 鍵繼續...")
            else:
                logging.error(f"超過最大重試次數 {max_retries}，無法訪問 {file_name}")
                raise
        except Exception as e:
            logging.error(f"確保工作表存在時發生錯誤: {e}\n{traceback.format_exc()}")
            raise

def clean_titles(file_name, is_movie):
    from others.excel_handler import clean_sheet_data
    try:
        logging.info("嘗試清理 video_detail.xlsx 中的 Title 或 Movies sheet")
        if is_movie:
            movie_df = clean_sheet_data(file_name, 'Movies')
            if not movie_df.empty and 'Movie Title' in movie_df.columns:
                movie_df['Movie Title'] = movie_df['Movie Title'].apply(lambda x: clear_title(x, 'show_movie'))
                with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    movie_df.to_excel(writer, sheet_name='Movies', index=False)
                logging.info("成功清理並保存 Movies sheet")
            else:
                logging.warning("Movies sheet 中無有效 'Movie Title' 數據，跳過清理")
        else:
            title_df = clean_sheet_data(file_name, 'Title')
            if not title_df.empty and 'TV Show Title' in title_df.columns:
                title_df['TV Show Title'] = title_df['TV Show Title'].apply(lambda x: clear_title(x, 'show_movie'))
                with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    title_df.to_excel(writer, sheet_name='Title', index=False)
                logging.info("成功清理並保存 Title sheet")
            else:
                logging.warning("Title sheet 中無有效 'TV Show Title' 數據，跳過清理")
    except Exception as e:
        logging.error(f"清理 Title/Movies sheet 時發生錯誤: {e}\n{traceback.format_exc()}")
        raise

def is_valid_tmdb_api_key(api_key):
    """【修改】驗證 TMDB API 金鑰，並新增無限重試機制以處理網路不穩。"""
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry

    if not api_key:
        logging.error("未提供 TMDB API 密鑰")
        return False
    
    session = requests.Session()
    retry_strategy = Retry(
        total=50,  # 總重試次數，設為一個較大的值以模擬無限重試
        status_forcelist=[429, 500, 502, 503, 504],  # 針對這些伺服器或速率限制錯誤進行重試
        allowed_methods=["HEAD", "GET", "OPTIONS"],
        backoff_factor=1  # 重試間的等待時間因子
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    
    headers = {"Authorization": f"Bearer {api_key}"}
    try:
        # 使用配置好重試機制的 session 發送請求
        response = session.get("https://api.themoviedb.org/3/configuration", headers=headers, timeout=10)
        if response.status_code == 200:
            logging.info("TMDB API 密鑰驗證成功")
            return True
        else:
            logging.error(f"TMDB API 密鑰無效，HTTP 狀態碼: {response.status_code}, 回應: {response.text}")
            return False
    except requests.exceptions.RequestException as e:
        # 只有在所有重試都失敗後，才會觸發此異常
        logging.error(f"TMDB API 請求失敗，已達最大重試次數或發生無法重試的錯誤: {str(e)}\n{traceback.format_exc()}")
        return False

def get_tmdb_extractor(configs):
    api_key = configs.get("tmdb_api")
    logging.info(f"從 configs.json 載入 TMDB API 密鑰: {'****' + api_key[-4:] if api_key else '未提供'}")
    if api_key and is_valid_tmdb_api_key(api_key):
        logging.info("檢測到有效的 TMDB API 密鑰，提取時將使用 tmdb_extractor_api")
        return extract_tmdb_data_api, api_key, True # True indicates API for EXTRACTION
    else:
        logging.info("無有效 TMDB API 密鑰，提取時將使用 tmdb_extractor_selenium")
        return extract_tmdb_data_selenium, None, False # False indicates Selenium for EXTRACTION

def extract_person_id(url):
    import re
    try:
        match = re.search(r'/person/([^/?]+)', url)
        if match:
            return match.group(1).split('-')[0]
        logging.error(f"無效的人員 URL: {url}")
        return None
    except Exception as e:
        logging.error(f"提取人員 ID 失敗: {url}, 錯誤: {e}")
        return None

def extract_and_save_details(driver_param, video_url, tmdb_url, split_value,
                                 tmdb_extractor, api_key, uses_api_for_extraction, configs,
                                 overwrite_titles, overwrite_descriptions,
                                 row_index, output_excel_path):
    """
    【修正】此函數專注於執行資料提取並儲存。
    新增了回退機制以兼容不支援 'file_path' 參數的舊版 extractor 函數。
    【修改】新增回傳值 is_overall_identical_or_skipped_body，用於告知呼叫者是否因內容相同而無需上傳。
    【使用者要求 2025-09-18】增加多個中斷檢查點，以便在收到 CTRL+C 信號後能更快地中止任務。
    【修改】將 _PROGRAM_INTERRUPTED 的檢查改為 utils._PROGRAM_INTERRUPTED。
    【使用者要求】在開始提取前，先執行 ensure_tmdb_accessible 以處理限流和 Cookie。
    """
    from others.credentials import login
    global tmdb_login_attempted_this_session

    log = logging.getLogger('extractor') # 使用 extractor logger

    # 【使用者要求】在函數開始時立即檢查中斷信號
    if utils._PROGRAM_INTERRUPTED:
        log.warning(f"[Extractor] 程序已中斷，跳過提取項目: {video_url or tmdb_url}")
        return False, None, False, False, True # 返回 is_identical=True 以跳過後續步驟

    overall_process_status = False
    media_type_for_extractor = "unknown"
    no_chinese_content_body_was_skipped = False
    is_overall_identical_or_skipped_body = True
    
    target_tmdb_url_for_finally = tmdb_url
    if video_url and not tmdb_url and "themoviedb.org" in video_url:
        target_tmdb_url_for_finally = video_url
    elif not tmdb_url and not video_url:
        target_tmdb_url_for_finally = "N/A"

    temporary_driver = None
    try:
        if utils._PROGRAM_INTERRUPTED: return False, None, False, False, True

        os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)
        check_create_excel(output_excel_path)
        
        target_tmdb_url = tmdb_url
        
        site_extraction_succeeded = False
        content_is_tv_show_from_site_extractor = None
        is_person = "/person/" in (target_tmdb_url or video_url or "")
        if is_valid_video_url(video_url) and not is_person:
            # 【使用者要求】在開始耗時的網站提取前檢查中斷信號
            if utils._PROGRAM_INTERRUPTED: return False, None, False, False, True
            log.info(f"[Extractor] 處理影片網站 URL: {video_url}")
            if driver_param:
                # 【使用者要求】保留 netflix, primevideo, appletv, disneyplus 等不同來源的處理邏輯，即使部分程式碼未提供。
                if "netflix.com" in video_url:
                    netflix_result = extract_netflix_episodes(driver_param, video_url, output_excel_path, target_tmdb_url, split_value=split_value)
                    found_tmdb_url = None
                    if isinstance(netflix_result, tuple) and len(netflix_result) == 3:
                        site_extraction_succeeded, content_is_tv_show_from_site_extractor, found_tmdb_url = netflix_result
                    elif isinstance(netflix_result, tuple) and len(netflix_result) == 2:
                        site_extraction_succeeded, content_is_tv_show_from_site_extractor = netflix_result
                    else:
                        site_extraction_succeeded = False
                    
                    if site_extraction_succeeded and found_tmdb_url and not target_tmdb_url:
                        target_tmdb_url = found_tmdb_url
                        target_tmdb_url_for_finally = found_tmdb_url
                elif "primevideo.com" in video_url or "amazon.com" in video_url:
                    is_movie_pv = extract_primevideo_data(driver_param, video_url, output_excel_path)
                    content_is_tv_show_from_site_extractor = not is_movie_pv
                    site_extraction_succeeded = True
                elif "appletv.com" in video_url or "tv.apple.com" in video_url:
                    extract_appletv_data(driver_param, video_url, output_excel_path)
                    site_extraction_succeeded = True
                elif "disneyplus.com" in video_url:
                    if not is_disneyplus_logged_in(driver_param):
                        disneyplus_email, disneyplus_password = get_disneyplus_credentials(configs)
                        if disneyplus_email and disneyplus_password:
                            login_to_disneyplus(driver_param, disneyplus_email, disneyplus_password)
                    result_type = extract_disneyplus_data(driver_param, video_url, output_excel_path)
                    if result_type is not None:
                        content_is_tv_show_from_site_extractor = result_type
                        site_extraction_succeeded = True
                elif "viu.com" in video_url:
                    viu_result = extract_viu_data(driver_param, video_url, output_excel_path)
                    if viu_result:
                        content_is_tv_show_from_site_extractor = True
                        site_extraction_succeeded = True
                elif "viu.tv" in video_url:
                    _viutv_tmdb_id = None
                    _tmdb_m = re.search(r'themoviedb\.org/tv/(\d+)', target_tmdb_url or '')
                    if _tmdb_m:
                        _viutv_tmdb_id = _tmdb_m.group(1)
                    viutv_result = extract_viutv_data(driver_param, video_url, output_excel_path, tmdb_id=_viutv_tmdb_id)
                    if viutv_result:
                        content_is_tv_show_from_site_extractor = True
                        site_extraction_succeeded = True
            else:
                log.error(f"[Extractor] 網站提取 ({video_url}) 需要 driver 但未提供。")
                site_extraction_succeeded = False
        
        # 【使用者要求】在網站提取之後，開始 TMDB 提取之前再次檢查
        if utils._PROGRAM_INTERRUPTED: return False, None, False, False, True

        extraction_succeeded_flag = False
        if site_extraction_succeeded:
            log.info(f"[Extractor] 已從 {video_url} 成功提取數據。")
            extraction_succeeded_flag = True
            is_overall_identical_or_skipped_body = False
            media_type_for_extractor = "tv" if content_is_tv_show_from_site_extractor else "movie"
        else:
            log.info("[Extractor] 未執行網站提取或提取失敗，將嘗試從 TMDB 提取數據。")
            current_target_url = video_url if ("themoviedb.org" in (video_url or "")) else tmdb_url
            if not current_target_url:
                log.warning(f"[Extractor] 無有效的 TMDB URL 可供提取 (video_url: {video_url}, tmdb_url: {tmdb_url})")
                return False, None, False, False, True

            target_tmdb_url_for_finally = current_target_url
            if "/tv/" in current_target_url: media_type_for_extractor = "tv"
            elif "/movie/" in current_target_url: media_type_for_extractor = "movie"
            elif "/person/" in current_target_url: media_type_for_extractor = "person"
            else: return False, None, False, False, True

            try:
                # 【新增】如果是使用瀏覽器進行 TMDB 提取 (非 API 模式)，先確保沒有被限流
                if driver_param and not uses_api_for_extraction:
                    ensure_tmdb_accessible(driver_param)

                # 理想情況：呼叫支援 file_path 的新版 extractor
                extractor_result = tmdb_extractor(current_target_url, api_key, driver_param, overwrite_titles, overwrite_descriptions, bulk_mode=(row_index is not None and row_index >= 0), row_index=row_index, file_path=output_excel_path)
            except TypeError as e:
                # 回退機制：如果 extractor 不支援 file_path 參數
                if "unexpected keyword argument 'file_path'" in str(e):
                    log.warning("偵測到您的 extractor 函數不支援 'file_path' 參數。將回退到預設檔案路徑 'data/video_detail.xlsx' 並移動檔案。")
                    log.warning(">>> 長期解決方案：請更新 'extractors/tmdb_extractor.py' 和 'extractors/tmdb_extractor_api.py' 中的 main 函數以接受 'file_path' 參數，例如： def main(..., file_path='data/video_detail.xlsx'):")
                    
                    default_path = 'data/video_detail.xlsx'
                    if os.path.exists(default_path):
                        try:
                            # 使用 safe_remove_file 替代 os.remove
                            safe_remove_file(default_path)
                        except OSError as oe:
                            log.error(f"回退機制：無法刪除舊的預設檔案 {default_path}: {oe}")

                    # 不帶 file_path 參數重新呼叫
                    extractor_result = tmdb_extractor(current_target_url, api_key, driver_param, overwrite_titles, overwrite_descriptions, bulk_mode=(row_index is not None and row_index >= 0), row_index=row_index)
                    
                    # 將 extractor 生成的預設檔案移動到目標路徑
                    if os.path.exists(default_path):
                        shutil.move(default_path, output_excel_path)
                        log.info(f"已將 {default_path} 移動至 {output_excel_path}")
                    else:
                        log.error(f"Extractor 在回退模式下未能創建預期的預設檔案: {default_path}")
                else:
                    raise e # 如果是其他 TypeError，則重新拋出

            if isinstance(extractor_result, tuple) and len(extractor_result) == 2:
                is_overall_identical_or_skipped_body, no_chinese_content_body_was_skipped = extractor_result
                extraction_succeeded_flag = True
            else:
                log.error(f"[Extractor] TMDB Extractor 返回了意外的格式: {extractor_result}。 URL: {current_target_url}")
                extraction_succeeded_flag = False

        if utils._PROGRAM_INTERRUPTED: return False, None, False, False, True
        
        if not extraction_succeeded_flag:
            log.error(f"[Extractor] 數據提取步驟失敗 for {target_tmdb_url_for_finally}")
            overall_process_status = False
        else:
            log.info("[Extractor] 數據提取成功，進行後處理。")
            overall_process_status = True
            # 【錯誤修正 2025-08-17】
            # 只有在 extractor 實際寫入了內容 (is_overall_identical_or_skipped_body 為 False) 且檔案確實存在時，才進行標題處理。
            # 這樣可以避免在 extractor 因無內容更新而刪除檔案後，這裡仍然嘗試讀取該已刪除檔案而導致 "File is not a zip file" 錯誤。
            if not is_overall_identical_or_skipped_body and os.path.exists(output_excel_path):
                process_titles(output_excel_path, output_excel_path)
            elif is_overall_identical_or_skipped_body:
                # 這是正常情況：extractor 判斷無需寫入，所以刪除了檔案，我們也無需處理。
                log.info(f"Extractor 判斷無需寫入 Excel 或內容相同，因此跳過 process_titles 步驟。")
            else: # not is_overall_identical_or_skipped_body 但檔案不存在，這是不預期情況
                log.warning(f"Extractor 應已寫入內容，但預期的 excel 檔案 {output_excel_path} 不存在，跳過 process_titles。")
        
        is_movie_final = (media_type_for_extractor == "movie")

        # 【2026-11 K2】網站判斷的類型必須與 TMDB URL 一致，否則上傳會用錯 ID 格式（例如 /tv/ 被當電影 → IndexError），
        # 甚至把電影資料寫進劇集。不一致時當作提取失敗，single / bulk 都不會上傳。
        _tmdb_for_check = target_tmdb_url_for_finally or ''
        if overall_process_status and media_type_for_extractor in ("tv", "movie"):
            _tmdb_type = "tv" if "/tv/" in _tmdb_for_check else ("movie" if "/movie/" in _tmdb_for_check else None)
            if _tmdb_type and _tmdb_type != media_type_for_extractor:
                _label = {"tv": "劇集", "movie": "電影"}
                log.error(
                    f"[Extractor] 類型不符，已跳過上傳：網站提取判斷為{_label[media_type_for_extractor]}，"
                    f"但 TMDB URL 是{_label[_tmdb_type]} ({_tmdb_for_check})。影片 URL: {video_url}"
                )
                print(f"⚠  類型不符（網站={_label[media_type_for_extractor]}，TMDB={_label[_tmdb_type]}），已跳過上傳: {_tmdb_for_check}")
                overall_process_status = False

        return overall_process_status, target_tmdb_url_for_finally, is_movie_final, no_chinese_content_body_was_skipped, is_overall_identical_or_skipped_body

    except (KeyboardInterrupt, SystemExit):
        # 【使用者要求】捕獲 KeyboardInterrupt 以更快地停止
        log.warning(f"在 extract_and_save_details 中檢測到中斷信號，正在終止此任務。")
        utils._PROGRAM_INTERRUPTED = True # 確保標誌被設定
        return False, None, False, False, True
    except Exception as e_global:
        if not utils._PROGRAM_INTERRUPTED:
            log.error(f"[Extractor] 處理項目 ({video_url or tmdb_url}) 時發生頂層錯誤: {e_global}\n{traceback.format_exc()}")
        save_unavailable_video_info_to_excel(video_url or tmdb_url, target_tmdb_url_for_finally or "N/A", file_name='data/unavailable_video.xlsx')
        return False, None, False, False, True
    finally:
        if temporary_driver:
            cleanup_browser(temporary_driver)

def is_details_file_empty(file_path):
    """
    【新增】檢查 details Excel 檔案是否完全沒有可上傳的資料。
    【修改】使用 with 語句確保文件句柄被正確關閉。
    """
    log = logging.getLogger('importer')
    if not os.path.exists(file_path):
        log.warning(f"is_details_file_empty: 檔案不存在 {file_path}")
        return True
    try:
        with pd.ExcelFile(file_path, engine='openpyxl') as xls:
            sheets_to_check = {
                'Title': ['TV Show Title', 'TV Show Description'],
                'Seasons': ['Season Name', 'Season Description'],
                'Episodes': ['Episode Title', 'Episode Description', 'Episode Runtime'],
                'Movies': ['Movie Title', 'Movie Description', 'Runtime'],
                'People': ['Name', 'Biography']
            }
            for sheet_name, columns_to_check in sheets_to_check.items():
                if sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    if not df.empty:
                        for col in columns_to_check:
                            if col in df.columns and df[col].notna().any() and df[col].astype(str).str.strip().any():
                                log.info(f"is_details_file_empty: 在 '{sheet_name}' sheet 的 '{col}' 欄位中找到有效資料。檔案非空。")
                                return False
        log.info(f"is_details_file_empty: 檢查完所有工作表後，未找到任何有效資料。檔案視為空。")
        return True
    except Exception as e:
        log.error(f"is_details_file_empty: 檢查檔案 {file_path} 時發生錯誤: {e}")
        return True # 出錯時，保守地認為檔案是空的以避免不必要的上傳

def safe_remove_file(file_path, max_retries=50, delay=1.0, interactive_prompt=True):
    """
    安全刪除檔案的輔助函數（支援無限重試）。

    如果刪除失敗（最常見原因是 Microsoft Excel 鎖定檔案），會持續重試並搭配 gc.collect()。

    使用方式：
    - 預設 max_retries=50（相容舊行為）
    - 設為 None 或 <=0 代表「無限重試」（推薦 Single 模式使用）
    - interactive_prompt=True 時會在適當時機提示使用者關閉 Excel，並進入低頻等待模式。
    """
    logger = logging.getLogger('importer')

    if not os.path.exists(file_path):
        return True

    infinite_retry = (max_retries is None or max_retries <= 0)
    attempt = 0
    prompted = False
    slow_mode = False
    current_delay = delay

    while True:
        attempt += 1
        try:
            os.remove(file_path)
            logger.info(f"成功刪除檔案: {file_path}（共重試 {attempt} 次）")
            return True
        except OSError as e:
            # 有限重試時超過次數就放棄
            if not infinite_retry and attempt >= max_retries:
                logger.error(f"刪除檔案 {file_path} 最終失敗（已重試 {attempt} 次）: {e}")
                return False

            # 控制 log 噪音（無限重試時特別重要）
            log_every = 8 if slow_mode else 3
            if attempt <= 5 or attempt % log_every == 0:
                remaining = "無限" if infinite_retry else str(max_retries - attempt)
                logger.warning(
                    f"刪除檔案 {file_path} 失敗（第 {attempt} 次，剩餘 {remaining}）: {e}。"
                    f" 等待 {current_delay:.1f} 秒後重試..."
                )

            gc.collect()
            time.sleep(current_delay)

            # 第一次達到提示時機 → 明確請使用者關閉 Excel
            if interactive_prompt and not prompted and attempt >= 5:
                print("\n" + "=" * 74)
                print(f"【檔案被鎖定】無法刪除：{os.path.basename(file_path)}")
                print("這幾乎一定是因為您在 Microsoft Excel 中開啟了這個檔案。")
                print("請在 Excel 中完整關閉該檔案，然後按 Enter 鍵。")
                print("程式會進入低頻等待模式（每 5~8 秒檢查一次），直到成功刪除為止。")
                print("您隨時可以按 Ctrl+C 中斷。")
                print("=" * 74)
                try:
                    input()
                except Exception:
                    pass
                prompted = True
                slow_mode = True
                current_delay = max(current_delay, 5.0)
                continue

            # 已經提示過後，進入更安靜的慢速模式
            if prompted and not slow_mode:
                slow_mode = True
                current_delay = max(current_delay, 5.0)

def extractor_worker(extract_queue, upload_queue, excel_lock, tmdb_extractor, api_key, uses_api, configs, overwrite_titles, overwrite_descriptions, in_progress_indices=None, in_progress_lock=None, shared_driver_container=None, shared_driver_lock=None):
    """
    Extractor 執行緒的目標函數。
    【修改】增加瀏覽器自動重啟與登入機制，以應對不穩定的瀏覽器關閉情況。
    shared_driver_container: [driver] 列表，所有 extractor 共用同一個瀏覽器實例。
    shared_driver_lock: threading.Lock，確保同一時間只有一個 extractor 使用瀏覽器。
    """
    log = logging.getLogger('extractor')
    driver = None  # 僅在非共享模式下使用
    use_shared = shared_driver_container is not None and shared_driver_lock is not None

    def get_current_driver():
        return shared_driver_container[0] if use_shared else driver

    def set_current_driver(d):
        nonlocal driver
        if use_shared:
            shared_driver_container[0] = d
        else:
            driver = d

    # 【新增】內部函數：初始化或重啟瀏覽器
    def ensure_driver_valid(force=False):
        # 如果使用 API 模式且未強制要求，通常不需要瀏覽器
        # force=True 用於影片網站 URL (如 viu.com) 即使在 API 模式下也需要瀏覽器的情況
        if uses_api and not force:
            return True

        current = get_current_driver()
        is_valid = False
        try:
            if current:
                # 嘗試執行一個輕量級指令來檢查瀏覽器是否活著
                current.title
                is_valid = True
        except Exception:
            is_valid = False

        if not is_valid:
            if current:
                log.warning("[Extractor] 檢測到瀏覽器失效，正在嘗試清理並重啟...")
                try:
                    cleanup_browser(current)
                except Exception:
                    pass
                set_current_driver(None)

            log.info("[Extractor] 正在初始化/重啟瀏覽器...")
            try:
                driver_tuple = get_browser_driver(profile_suffix='_extractor', remote_port=9222)
                set_current_driver(driver_tuple[0])
                log.info("[Extractor] 瀏覽器初始化完畢。")
                return True
            except Exception as e:
                log.error(f"[Extractor] 瀏覽器初始化失敗: {e}")
                return False
        return True

    try:
        log.info("[Extractor] 初始化提取執行緒...")
        
        # 初始啟動
        if not uses_api:
            ensure_driver_valid()
        else:
            log.info("[Extractor] 使用 API 模式，無需初始化瀏覽器。")

        log.info("[Extractor] 初始化完成，將從共享緩存讀取任務資料。")

        while not utils._PROGRAM_INTERRUPTED:
            utils.wait_if_paused(log, 'Extractor')
            try:
                task = extract_queue.get(timeout=1)
            except queue.Empty:
                continue

            if utils._PROGRAM_INTERRUPTED:
                log.warning("[Extractor] 接收到任務後檢測到中斷信號，將跳過並準備退出。")
                extract_queue.task_done()
                break
            if task is None:
                log.info("[Extractor] 收到結束信號，正在將結束信號傳遞給上傳執行緒。")
                upload_queue.put(None)
                break

            # 立即追蹤此任務，防止主執行緒在兩個 queue 都空時重複入隊同一 row
            row_index = task['row_index']
            if in_progress_indices is not None:
                with in_progress_lock:
                    in_progress_indices.add(row_index)

            # 【新增】在處理任務前確保瀏覽器有效 (僅限非 API 模式)
            if not uses_api:
                if not ensure_driver_valid():
                    log.error("[Extractor] 無法建立有效的瀏覽器連線，將此任務放回佇列稍後重試。")
                    extract_queue.put({'row_index': row_index})
                    extract_queue.task_done()
                    time.sleep(5)
                    continue

            utils.reload_dictionaries_and_converters()
            video_url, tmdb_url, split_value = None, None, None
            row = None
            try:
                row = get_row_from_shared_cache(row_index)
                if row is None:
                    # 【2026-11 F2】緩存可能比剛加入的新項目舊，先重新讀取 Excel 一次再判斷
                    log.info(f"[Extractor] 共享緩存中找不到任務 (索引: {row_index})，重新讀取 Excel 再試...")
                    _refresh_shared_cache(excel_lock, 'data/bulk_videos.xlsx', retry_interval=2, max_retries=2)
                    row = get_row_from_shared_cache(row_index)
                if row is None:
                    log.error(f"[Extractor] 共享緩存中找不到任務 (索引: {row_index})，跳過。")
                    if in_progress_indices is not None:
                        with in_progress_lock:
                            in_progress_indices.discard(row_index)
                    extract_queue.task_done()
                    continue

                video_url = str(row.get('Video URL', '')).strip()
                tmdb_url = str(row.get('TMDB URL', '')).strip()
                done_value = row.get('Done', '')
                done_str = str(done_value).strip()
                if not (pd.isna(done_value) or done_str == '' or done_str.lower() in ['nan', '0', '0.0']):
                    log.info(f"[Extractor] 跳過任務 (索引: {row_index})，在處理前發現其狀態已為 'Done'。 [Video URL: {video_url}, TMDB URL: {tmdb_url}]")
                    if in_progress_indices is not None:
                        with in_progress_lock:
                            in_progress_indices.discard(row_index)
                    extract_queue.task_done()
                    continue

                # 檢查批次緩存（比共享緩存更新），若 Done=1 則跳過，防止重複處理
                _pending = get_pending_done_value(row_index)
                if _pending is not None and _pending != 0:
                    log.info(f"[Extractor] 跳過任務 (索引: {row_index})，批次緩存顯示已完成 (Done={_pending})，避免重複處理。")
                    if in_progress_indices is not None:
                        with in_progress_lock:
                            in_progress_indices.discard(row_index)
                    extract_queue.task_done()
                    continue

                split_value = str(row.get('Split', '')).strip().lower()
                # 【2026-11】該行設定欄位（空白 = 程式選項）
                row_opts = utils.parse_row_options(row)
                utils.ROW_OPTIONS.set(row_opts)
                task_ot = row_opts.get('overwrite_titles', overwrite_titles)
                task_od = row_opts.get('overwrite_descriptions', overwrite_descriptions)
                if row_opts:
                    log.info(f"[Extractor] 任務 (索引: {row_index}) 使用 Excel 欄位設定: {row_opts}")
            except Exception as e:
                log.error(f"[Extractor] 從 bulk_videos.xlsx 讀取任務 (索引: {row_index}) 資料時失敗: {e}")
                if in_progress_indices is not None:
                    with in_progress_lock:
                        in_progress_indices.discard(row_index)
                extract_queue.task_done()
                continue

            # 即使在 API 模式下，影片網站 URL (如 viu.com) 也需要瀏覽器
            # 在共享模式下，用 lock 確保同一時間只有一個 thread 使用瀏覽器
            needs_browser = is_valid_video_url(video_url)
            lock_ctx = shared_driver_lock if (use_shared and needs_browser) else contextlib.nullcontext()

            with lock_ctx:
                if needs_browser and get_current_driver() is None:
                    log.info(f"[Extractor] 偵測到影片網站 URL，強制初始化瀏覽器: {video_url}")
                    if not ensure_driver_valid(force=True):
                        log.error("[Extractor] 無法為影片網站建立瀏覽器連線，將跳過此任務。")
                        if in_progress_indices is not None:
                            with in_progress_lock:
                                in_progress_indices.discard(row_index)
                        extract_queue.task_done()
                        continue

                active_driver = get_current_driver()
                log.info(f"[Extractor] 開始處理任務 (索引: {row_index}), Video URL: {video_url}, TMDB URL: {tmdb_url}")

                output_excel_path = os.path.join(DETAILS_CACHE_DIR, f"details_{row_index}.xlsx")

                try:
                    success, final_tmdb_url, is_movie, no_chinese, is_identical = extract_and_save_details(
                        active_driver, video_url, tmdb_url, split_value,
                        tmdb_extractor, api_key, uses_api, configs,
                        task_ot, task_od,
                        row_index, output_excel_path
                    )
                except (WebDriverException, Exception) as e:
                    # 【新增】捕獲提取過程中的瀏覽器錯誤
                    if "chrome not reachable" in str(e) or "no such window" in str(e) or "target window already closed" in str(e) or isinstance(e, WebDriverException):
                        log.error(f"[Extractor] 瀏覽器在提取過程中發生致命錯誤: {e}。將標記為失敗並重啟瀏覽器。")
                        current_drv = get_current_driver()
                        try:
                            cleanup_browser(current_drv)
                        except Exception:
                            pass
                        set_current_driver(None)
                        success = False  # 視為失敗
                        is_identical = False
                    else:
                        raise e  # 其他錯誤重新拋出或讓外層處理

            if utils._PROGRAM_INTERRUPTED:
                log.info(f"[Extractor] 在處理任務 (索引: {row_index}) 後偵測到中斷，終止。")
                break
            if success:
                if is_identical:
                    log.info(f"[Extractor] 提取成功 (索引: {row_index})，但無任何內容需要寫入 Excel。直接標記為完成 (Done=1)。")
                    update_done_value_fast(row_index, 1)
                    safe_remove_file(output_excel_path)
                    if in_progress_indices is not None:
                        with in_progress_lock:
                            in_progress_indices.discard(row_index)
                else:
                    log.info(f"[Extractor] 提取成功 (索引: {row_index})，將任務放入上傳隊列。")
                    upload_queue.put({
                        'row_index': row_index,
                        'details_file': output_excel_path,
                        'tmdb_url': final_tmdb_url,
                        'is_movie': is_movie,
                        'video_url': video_url,
                        'no_chinese_skip': no_chinese,
                        'row_options': row_opts,
                    })
                    # in_progress will be removed by uploader_worker after upload completes
            else:
                log.error(f"[Extractor] 提取失敗 (索引: {row_index})，此項目將不會被上傳。將 Done 標記為 0。")
                update_done_value_fast(row_index, 0)
                safe_remove_file(output_excel_path)
                if in_progress_indices is not None:
                    with in_progress_lock:
                        in_progress_indices.discard(row_index)
            extract_queue.task_done()
    except KeyboardInterrupt:
        log.info("[Extractor] 檢測到 Ctrl+C，正在終止提取執行緒...")
    except Exception as e:
        if not utils._PROGRAM_INTERRUPTED:
            log.error(f"[Extractor] 執行緒發生未預期錯誤: {e}\n{traceback.format_exc()}")
    finally:
        # 共享模式下，瀏覽器由 bulk mode 統一清理；非共享模式下才在此關閉
        if not use_shared and driver:
            cleanup_browser(driver)
            log.info("[Extractor] 瀏覽器已關閉。")
        log.info("[Extractor] 提取執行緒已結束。")

def compare_and_filter_excel_with_tmdb_api(details_file, tmdb_url, api_key, overwrite_titles='ask', overwrite_descriptions='ask'):
    """
    Fetch current TMDB data and compare with details_file content.
    Removes rows that are already identical to what's on TMDB.
    Returns True if all content is identical (upload can be skipped), False otherwise.

    When overwrite_titles/descriptions == 'y' and the Excel value is empty,
    we treat it as "intent to clear", so it will NOT be considered identical
    (the clear will be allowed to proceed).
    """
    log = logging.getLogger('importer')
    from others.utils import normalize_text, convert_runtime_to_minutes

    TMDB_BASE = "https://api.themoviedb.org/3"
    headers = {"Authorization": f"Bearer {api_key}"}

    def _fetch(endpoint, lang='zh-TW'):
        try:
            r = requests.get(f"{TMDB_BASE}{endpoint}", headers=headers,
                             params={'language': lang}, timeout=15)
            if r.status_code == 200:
                return r.json()
        except Exception:
            pass
        return None

    def _fetch_fb(endpoint):
        data = _fetch(endpoint, 'zh-HK')
        if data is None:
            data = _fetch(endpoint, 'zh-TW')
        return data if data is not None else _fetch(endpoint, 'zh-CN')

    def _s(val):
        if val is None:
            return ''
        try:
            if pd.isna(val):
                return ''
        except (TypeError, ValueError):
            pass
        return str(val).strip()

    def _nt(raw):
        s = _s(raw)
        return normalize_text(s, field_type='title') if s else ''

    def _nn(raw):
        s = _s(raw)
        return normalize_text(s, field_type='name') if s else ''

    def _nb(raw):
        s = _s(raw)
        return normalize_text(s, field_type='bio') if s else ''

    def _match(excel_raw, tmdb_raw, norm_fn, field_overwrite='ask'):
        """
        Returns True if content matches (or excel is empty *and we are not forcing a clear*).
        Special case for clearing intent:
        - If Excel is empty + overwrite=y, normally we want to force clear.
        - Exception: if the current TMDB value is a generic title (泛標題),
          and Excel is empty, we consider it "the same" (no need to clear a generic title).
        """
        ev = norm_fn(excel_raw)
        tv = norm_fn(tmdb_raw)

        if not ev:
            # Empty in Excel
            if field_overwrite.lower() == 'y':
                # Check if TMDB already has a generic title — if so, no need to "clear" it
                if is_generic_title(tv):
                    return True  # treat as same, skip unnecessary clear of generic title
                return False  # force clear intent
            return True  # normal case: empty means nothing to do
        return ev == tv

    def _rt_match(excel_raw, tmdb_minutes):
        ev = convert_runtime_to_minutes(_s(excel_raw))
        if not _s(excel_raw):
            return True
        if ev is None and (tmdb_minutes is None or tmdb_minutes == 0):
            return True
        if ev is None or not tmdb_minutes:
            return False
        return ev == int(tmdb_minutes)

    try:
        xls = pd.ExcelFile(details_file, engine='openpyxl')
        sheets = {name: pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names}
        xls.close()

        modified = False
        any_different = False

        def filter_sheet(sname, row_fn):
            nonlocal modified, any_different
            if sname not in sheets:
                return
            df = sheets[sname]
            if df.empty:
                return
            keep_flags = []
            for _, row in df.iterrows():
                is_same = row_fn(row)
                keep_flags.append(not is_same)
                if not is_same:
                    any_different = True
            new_df = df[keep_flags].reset_index(drop=True)
            if len(new_df) < len(df):
                sheets[sname] = new_df
                modified = True

        def _strip_cn_fallback(hk_val, cn_val, norm_fn):
            """Return '' if zh-HK value is just TMDB's implicit zh-CN language fallback.
            TMDB endpoints return zh-CN text when zh-HK translation is missing. After
            normalization (simplified→traditional) this matches Excel content, causing
            false 'identical' detection."""
            if hk_val and cn_val and norm_fn(hk_val) == norm_fn(cn_val):
                return ''
            return hk_val

        person_cache = {}
        person_cache_cn = {}

        def _get_person(pid):
            if pid not in person_cache:
                person_cache[pid] = _fetch_fb(f"/person/{pid}")
            return person_cache[pid]

        def _get_person_cn(pid):
            if pid not in person_cache_cn:
                person_cache_cn[pid] = _fetch(f"/person/{pid}", 'zh-CN')
            return person_cache_cn[pid]

        def _person_row_same(row):
            person_url = _s(row.get('Person URL'))
            if not person_url or '/person/' not in person_url:
                return False
            try:
                pid = extract_person_id(person_url)
                p_data = _get_person(pid)
                if p_data is None:
                    return False
                p_cn = _get_person_cn(pid) or {}
                actual_name = _strip_cn_fallback(p_data.get('name', ''), p_cn.get('name', ''), _nn)
                actual_bio = _strip_cn_fallback(p_data.get('biography', ''), p_cn.get('biography', ''), _nb)
                return (_match(row.get('Name'), actual_name, _nn, overwrite_titles) and
                        _match(row.get('Biography'), actual_bio, _nb, overwrite_descriptions))
            except Exception:
                return False

        if '/person/' in tmdb_url:
            pid = extract_person_id(tmdb_url)
            p_data = _get_person(pid)
            if p_data is None:
                return False
            p_cn = _get_person_cn(pid) or {}
            actual_name = _strip_cn_fallback(p_data.get('name', ''), p_cn.get('name', ''), _nn)
            actual_bio = _strip_cn_fallback(p_data.get('biography', ''), p_cn.get('biography', ''), _nb)

            def person_row_same(row):
                return (_match(row.get('Name'), actual_name, _nn, overwrite_titles) and
                        _match(row.get('Biography'), actual_bio, _nb, overwrite_descriptions))

            filter_sheet('People', person_row_same)

        elif '/movie/' in tmdb_url:
            mid = extract_movie_id(tmdb_url)
            data = _fetch_fb(f"/movie/{mid}")
            if data is None:
                return False
            data_cn = _fetch(f"/movie/{mid}", 'zh-CN') or {}
            tmdb_title = _strip_cn_fallback(data.get('title', ''), data_cn.get('title', ''), _nt)
            tmdb_desc = _strip_cn_fallback(data.get('overview', ''), data_cn.get('overview', ''), _nb)
            tmdb_rt = data.get('runtime', None)

            def movie_row_same(row):
                return (_match(row.get('Movie Title'), tmdb_title, _nt, overwrite_titles) and
                        _match(row.get('Movie Description'), tmdb_desc, _nb, overwrite_descriptions) and
                        _rt_match(row.get('Runtime'), tmdb_rt))

            filter_sheet('Movies', movie_row_same)
            filter_sheet('People', _person_row_same)

        elif '/tv/' in tmdb_url:
            tid = extract_tv_show_id(tmdb_url)
            tv_data = _fetch_fb(f"/tv/{tid}")
            if tv_data is None:
                return False
            tv_data_cn = _fetch(f"/tv/{tid}", 'zh-CN') or {}

            def title_row_same(row):
                actual_name = _strip_cn_fallback(tv_data.get('name', ''), tv_data_cn.get('name', ''), _nt)
                actual_overview = _strip_cn_fallback(tv_data.get('overview', ''), tv_data_cn.get('overview', ''), _nb)
                return (_match(row.get('TV Show Title'), actual_name, _nt, overwrite_titles) and
                        _match(row.get('TV Show Description'), actual_overview, _nb, overwrite_descriptions))

            filter_sheet('Title', title_row_same)

            season_cache = {}
            season_cache_cn = {}

            def _get_season(snum):
                if snum not in season_cache:
                    season_cache[snum] = _fetch_fb(f"/tv/{tid}/season/{snum}")
                return season_cache[snum]

            def _get_season_cn(snum):
                if snum not in season_cache_cn:
                    season_cache_cn[snum] = _fetch(f"/tv/{tid}/season/{snum}", 'zh-CN')
                return season_cache_cn[snum]

            def season_row_same(row):
                snum_s = _s(row.get('Season Number'))
                if not snum_s:
                    return False
                snum = int(float(snum_s))
                s_data = _get_season(snum)
                if s_data is None:
                    return False
                s_cn = _get_season_cn(snum) or {}
                actual_name = _strip_cn_fallback(s_data.get('name', ''), s_cn.get('name', ''), _nt)
                actual_overview = _strip_cn_fallback(s_data.get('overview', ''), s_cn.get('overview', ''), _nb)
                return (_match(row.get('Season Name'), actual_name, _nt, overwrite_titles) and
                        _match(row.get('Season Description'), actual_overview, _nb, overwrite_descriptions))

            filter_sheet('Seasons', season_row_same)

            def episode_row_same(row):
                snum_s = _s(row.get('Season Number'))
                enum_s = _s(row.get('Episode Number'))
                if not snum_s or not enum_s:
                    return False
                snum = int(float(snum_s))
                enum = int(float(enum_s))
                s_data = _get_season(snum)
                if s_data is None:
                    return False
                ep_data = next((e for e in s_data.get('episodes', [])
                                if e.get('episode_number') == enum), None)
                if ep_data is None:
                    return False
                s_cn = _get_season_cn(snum) or {}
                ep_cn = next((e for e in s_cn.get('episodes', [])
                              if e.get('episode_number') == enum), {}) or {}
                actual_name = _strip_cn_fallback(ep_data.get('name', ''), ep_cn.get('name', ''), _nt)
                actual_overview = _strip_cn_fallback(ep_data.get('overview', ''), ep_cn.get('overview', ''), _nb)
                return (_match(row.get('Episode Title'), actual_name, _nt, utils.episode_overwrite('title', overwrite_titles)) and
                        _match(row.get('Episode Description'), actual_overview, _nb, utils.episode_overwrite('description', overwrite_descriptions)) and
                        _rt_match(row.get('Episode Runtime'), ep_data.get('runtime')))

            filter_sheet('Episodes', episode_row_same)

            filter_sheet('People', _person_row_same)

        if modified:
            with pd.ExcelWriter(details_file, engine='openpyxl') as writer:
                for sname, df in sheets.items():
                    df.to_excel(writer, sheet_name=sname, index=False)

        if not any_different:
            log.info(f"[Uploader] API 比較: 所有內容已與 TMDB 相同，跳過上傳。URL: {tmdb_url}")
        else:
            log.info(f"[Uploader] API 比較: 有差異，已過濾相同行，僅上傳不同部份。URL: {tmdb_url}")
        return not any_different

    except Exception as e:
        log.error(f"[Uploader] API 比較時出錯: {e}\n{traceback.format_exc()}")
        return False


def apply_force_clear_from_master(master_row, details_file, overwrite_titles, overwrite_descriptions):
    """
    【新功能】支援「清空 TMDB 資料」需求。
    如果使用者在 bulk_videos.xlsx 把 title / description 相關欄位清空，
    且 overwrite 允許，則在準備上傳前強制把 details_file 內的對應欄位清空，
    讓 uploader 的清空邏輯 (check_and_fill_form 的 empty 分支) 可以真正生效。
    """
    if not os.path.exists(details_file) or master_row is None:
        return

    try:
        # 偵測 master row 是否有清空意圖
        force_clear_title = False
        force_clear_desc = False

        title_cols = ['TV Show Title', 'Movie Title', 'Name', 'Season Name', 'Episode Title']
        desc_cols = ['TV Show Description', 'Movie Description', 'Biography', 'Season Description', 'Episode Description']

        for col in title_cols:
            val = master_row.get(col)
            if pd.notna(val) and str(val).strip() == '':
                force_clear_title = True
                break

        for col in desc_cols:
            val = master_row.get(col)
            if pd.notna(val) and str(val).strip() == '':
                force_clear_desc = True
                break

        if not force_clear_title and not force_clear_desc:
            return

        if (force_clear_title and str(overwrite_titles).lower() == 'n') or \
           (force_clear_desc and str(overwrite_descriptions).lower() == 'n'):
            logging.info("[ForceClear] Master 要求清空，但 overwrite 為 n，跳過。")
            return

        logging.info(f"[ForceClear] 偵測到 master 清空要求，強制清空 details_file 中的 title/desc: {os.path.basename(details_file)}")

        with pd.ExcelFile(details_file, engine='openpyxl') as xls:
            sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

        modified = False

        # Title sheet
        if 'Title' in sheets and force_clear_title:
            df = sheets['Title']
            for c in ['TV Show Title', 'Movie Title']:
                if c in df.columns:
                    df[c] = ''
                    modified = True

        # Movies
        if 'Movies' in sheets and (force_clear_title or force_clear_desc):
            df = sheets['Movies']
            if force_clear_title and 'Movie Title' in df.columns:
                df['Movie Title'] = ''
                modified = True
            if force_clear_desc and 'Movie Description' in df.columns:
                df['Movie Description'] = ''
                modified = True

        # Seasons
        if 'Seasons' in sheets and (force_clear_title or force_clear_desc):
            df = sheets['Seasons']
            if force_clear_title and 'Season Name' in df.columns:
                df['Season Name'] = ''
                modified = True
            if force_clear_desc and 'Season Description' in df.columns:
                df['Season Description'] = ''
                modified = True

        # Episodes
        if 'Episodes' in sheets and (force_clear_title or force_clear_desc):
            df = sheets['Episodes']
            if force_clear_title and 'Episode Title' in df.columns:
                df['Episode Title'] = ''
                modified = True
            if force_clear_desc and 'Episode Description' in df.columns:
                df['Episode Description'] = ''
                modified = True

        # People
        if 'People' in sheets and (force_clear_title or force_clear_desc):
            df = sheets['People']
            if force_clear_title and 'Name' in df.columns:
                df['Name'] = ''
                modified = True
            if force_clear_desc and 'Biography' in df.columns:
                df['Biography'] = ''
                modified = True

        if modified:
            with pd.ExcelWriter(details_file, engine='openpyxl') as writer:
                for sname, df in sheets.items():
                    df.to_excel(writer, sheet_name=sname, index=False)
            logging.info("[ForceClear] 已強制清空 detail Excel 內的 title/desc 欄位。")

    except Exception as e:
        logging.warning(f"[ForceClear] 套用強制清空失敗: {e}")


def uploader_worker(upload_queue, excel_lock, configs, overwrite_titles, overwrite_descriptions, extractor_count=1, in_progress_indices=None, in_progress_lock=None, api_key=None):
    """
    Uploader 執行緒的目標函數。
    【修改】增加瀏覽器自動重啟與登入機制，確保在瀏覽器崩潰或關閉後能自動恢復並登入 TMDB。
    【修改】在登入前執行 ensure_tmdb_accessible，以確保能處理限流頁面。
    """
    log = logging.getLogger('importer')
    driver = None
    # 比較用的 API key：不論是否使用 API 模式提取，只要 configs 有設定就可用於比對
    _cmp_api_key = api_key or (configs.get("tmdb_api") if isinstance(configs, dict) else None)

    # 【新增】內部函數：初始化瀏覽器並登入
    def initialize_driver_and_login():
        nonlocal driver
        from others.credentials import login
        max_retries = 3
        for i in range(max_retries):
            if utils._PROGRAM_INTERRUPTED:   # 【2026-11-14】停止中唔再開瀏覽器
                return None
            try:
                log.info(f"[Uploader] 正在啟動瀏覽器並登入 (嘗試 {i+1}/{max_retries})...")
                driver_tuple = get_browser_driver(profile_suffix='_uploader', remote_port=9223)
                new_driver = driver_tuple[0]
                
                # 【新增】在檢查登入狀態前，先確保 TMDB 可訪問 (處理限流/Cookie)
                ensure_tmdb_accessible(new_driver)

                # 直接對 uploader 瀏覽器做實際 URL 檢查，
                # 避免被其他執行緒設定的 _tmdb_confirmed_login_status_session 快取誤導
                new_driver.get("https://www.themoviedb.org/login")
                time.sleep(2)
                actually_logged_in = '/u/' in new_driver.current_url
                if not actually_logged_in:
                    tmdb_username, tmdb_password = get_tmdb_credentials(configs)
                    if tmdb_username and tmdb_password:
                        login(new_driver, tmdb_username, tmdb_password)
                        log.info("[Uploader] 已完成 TMDB 登入。")
                    else:
                        log.warning("[Uploader] 未找到 TMDB 憑證，無法自動登入。")
                else:
                    log.info("[Uploader] Uploader 瀏覽器已確認登入狀態 (URL 檢查)。")
                
                log.info("[Uploader] 瀏覽器啟動並登入成功。")
                return new_driver
            except Exception as e:
                log.error(f"[Uploader] 瀏覽器啟動或登入失敗: {e}")
                if 'new_driver' in locals() and new_driver:
                    cleanup_browser(new_driver)
                time.sleep(5)
        log.error("[Uploader] 無法啟動瀏覽器，上傳執行緒將無法工作。")
        return None

    def ensure_uploader_driver_ready():
        """真正需要寫入 TMDB 時才啟動/重啟瀏覽器。"""
        nonlocal driver
        is_driver_alive = False
        try:
            if driver:
                _ = driver.title
                is_driver_alive = ensure_browser_window(driver)
        except Exception:
            is_driver_alive = False

        if is_driver_alive:
            return True
        # 【2026-11-14】已收到停止信號：唔重開瀏覽器（以前會重試 3 次再開 Edge、下載 driver，拖慢停止約 1 分鐘）
        if utils._PROGRAM_INTERRUPTED:
            return False

        log.info("[Uploader] " + ("首次初始化瀏覽器..." if driver is None else "檢測到瀏覽器已斷開，正在嘗試重新啟動並登入..."))
        if driver:
            try:
                cleanup_browser(driver)
            except Exception:
                pass
            driver = None
        driver = initialize_driver_and_login()
        return driver is not None

    try:
        log.info("[Uploader] 上傳執行緒已啟動，等待任務後才初始化瀏覽器...")
        none_received = 0

        while not utils._PROGRAM_INTERRUPTED:
            utils.wait_if_paused(log, 'Uploader', (utils.PAUSE_FLAG, utils.UPLOAD_PAUSE_FLAG))
            try:
                task = upload_queue.get(timeout=1)
            except queue.Empty:
                continue
            if utils._PROGRAM_INTERRUPTED:
                log.warning("[Uploader] 接收到任務後檢測到中斷信號，將跳過並準備退出。")
                upload_queue.task_done()
                break
            if task is None:
                none_received += 1
                log.info(f"[Uploader] 收到結束信號 ({none_received}/{extractor_count})。")
                if none_received >= extractor_count:
                    log.info("[Uploader] 所有提取執行緒均已結束，上傳執行緒準備退出。")
                    break
                continue

            row_index = task['row_index']
            details_file = task['details_file']
            tmdb_url = task['tmdb_url']
            is_movie = task['is_movie']
            video_url = task['video_url']
            no_chinese_skip = task['no_chinese_skip']
            # 【2026-11】該行設定欄位（由 extractor 讀取時傳過嚟）
            row_opts = task.get('row_options') or {}
            utils.ROW_OPTIONS.set(row_opts)
            task_ot = row_opts.get('overwrite_titles', overwrite_titles)
            task_od = row_opts.get('overwrite_descriptions', overwrite_descriptions)
            try:
                hardcoded_extractor_path = 'data/video_detail.xlsx'
                if os.path.exists(hardcoded_extractor_path):
                    log.warning(f"[Uploader] 檢測到 extractor 寫入的預設檔案 '{hardcoded_extractor_path}'。將其移動至正確的快取位置 '{details_file}'。")
                    try:
                        # 【修改】使用 safe_remove_file
                        safe_remove_file(details_file) 
                        shutil.move(hardcoded_extractor_path, details_file)
                    except Exception as e:
                        log.error(f"[Uploader] 移動 extractor 的檔案時出錯: {e}")

                if is_details_file_empty(details_file):
                    log.info(f"[Uploader] 項目 (索引: {row_index}) 的 details excel 檔案為空，無需上傳。直接標記為完成 (Done=1)。")
                    update_done_value_fast(row_index, 1)
                else:
                    log.info(f"[Uploader] 開始處理上傳任務 (索引: {row_index}), TMDB URL: {tmdb_url}")

                    # 驗證 TMDB URL 是否為有效的內容 URL
                    valid_content_url = any(seg in (tmdb_url or '') for seg in ('/tv/', '/movie/', '/person/'))
                    if not valid_content_url:
                        log.error(f"[Uploader] TMDB URL 格式無效，需包含 /tv/、/movie/ 或 /person/: {tmdb_url}，跳過上傳。")
                        update_done_value_fast(row_index, 0)
                        continue

                    upload_skipped = False
                    if _cmp_api_key:
                        try:
                            all_identical = compare_and_filter_excel_with_tmdb_api(
                                details_file, tmdb_url, _cmp_api_key,
                                overwrite_titles=task_ot,
                                overwrite_descriptions=task_od
                            )
                            if all_identical or is_details_file_empty(details_file):
                                log.info(f"[Uploader] API 比較: 內容已與 TMDB 相同，跳過上傳 (索引: {row_index})")
                                update_done_value_fast(row_index, 1)
                                upload_skipped = True
                        except Exception as _cmp_err:
                            log.warning(f"[Uploader] API 比較失敗，繼續正常上傳: {_cmp_err}")

                    if not upload_skipped:
                        # 【新功能】支援使用者在 bulk_videos.xlsx 清空 title/description 欄位 + overwrite=Y 來強制清空 TMDB 資料
                        try:
                            master_row = get_row_from_shared_cache(row_index)
                            if master_row is not None:
                                apply_force_clear_from_master(master_row, details_file, task_ot, task_od)
                        except Exception as clear_err:
                            log.warning(f"[Uploader] 套用 master 強制清空時發生錯誤: {clear_err}")

                        upload_success = False
                        driver_ready = ensure_uploader_driver_ready()
                        if not driver_ready and utils._PROGRAM_INTERRUPTED:
                            # 【2026-11-14】停止中：唔上傳、唔寫 Done，下次啟動會重新處理呢行
                            log.warning(f"[Uploader] 程式停止中，未上傳任務 (索引: {row_index})，下次啟動會重新處理。")
                        elif not driver_ready:
                            log.error(f"[Uploader] 無法啟動瀏覽器，跳過此任務 (索引: {row_index})。")
                            update_done_value_fast(row_index, 0)
                        else:
                            try:
                                if "/person/" in tmdb_url:
                                    upload_success = upload_person_to_tmdb(driver, tmdb_url, details_file, task_ot, task_od)
                                else:
                                    upload_success = upload_to_tmdb(driver, tmdb_url, details_file, is_movie, task_ot, task_od, upload_only_people=no_chinese_skip)
                            except (KeyboardInterrupt, SystemExit):
                                log.warning(f"[Uploader] 在上傳任務 (索引: {row_index}) 期間捕獲到中斷信號。")
                                utils._PROGRAM_INTERRUPTED = True
                                upload_success = False
                            except Exception as e:
                                if is_dead_browser_error(e) or isinstance(e, WebDriverException):
                                    log.error(f"[Uploader] 瀏覽器在上傳過程中崩潰: {e}。標記為失敗，下次循環將重啟瀏覽器。")
                                    try:
                                        cleanup_browser(driver)
                                    except Exception:
                                        pass
                                    driver = None
                                else:
                                    log.error(f"[Uploader] 上傳任務 (索引: {row_index}) 失敗: {e}\n{traceback.format_exc()}")
                                upload_success = False

                            if upload_success:
                                log.info(f"[Uploader] 上傳成功 (索引: {row_index})，更新 Done 狀態。")
                                update_done_value_fast(row_index, 1)
                            else:
                                log.error(f"[Uploader] 上傳失敗 (索引: {row_index})，標記 Done=0（本次 session 不再重試，重啟程式可重試）。")
                                update_done_value_fast(row_index, 0)
            finally:
                if in_progress_indices is not None:
                    with in_progress_lock:
                        in_progress_indices.discard(row_index)
                safe_remove_file(details_file)
                upload_queue.task_done()

    except KeyboardInterrupt:
        log.info("[Uploader] 檢測到 Ctrl+C，正在終止上傳執行緒...")
    except Exception as e:
        if not utils._PROGRAM_INTERRUPTED:
            log.error(f"[Uploader] 執行緒發生未預期錯誤: {e}\n{traceback.format_exc()}")
    finally:
        if driver:
            cleanup_browser(driver)
            log.info("[Uploader] 瀏覽器已關閉。")
        log.info("[Uploader] 上傳執行緒已結束。")

def normalize_tmdb_url(url):
    """將 TMDB URL 正規化為 'https://www.themoviedb.org/[type]/[id]' 格式，移除多餘部分。"""
    if not isinstance(url, str) or 'themoviedb.org' not in url:
        return url
    
    # 正則表達式，用於捕獲類型（tv, movie, person）和 ID
    match = re.search(r'themoviedb\.org/(tv|movie|person)/(\d+)', url)
    if match:
        media_type = match.group(1)
        tmdb_id = match.group(2)
        return f"https://www.themoviedb.org/{media_type}/{tmdb_id}"
    
    # 如果不匹配，返回原始 URL
    return url

def _is_blank_bulk_row(row):
    """【2026-11】空行：Video URL 同 TMDB URL 都冇內容（Split / Done 有冇值都當空行，冇網址根本處理唔到）。"""
    for col in ('Video URL', 'TMDB URL'):
        v = row.get(col, '')
        if not pd.isna(v) and str(v).strip().lower() not in ('', 'nan', 'none'):
            return False
    return True


def normalize_and_deduplicate_urls(file_path='data/bulk_videos.xlsx'):
    """
    【修改】讀取 bulk Excel 檔案，正規化 TMDB URL 並移除重複項，然後儲存回原檔案，同時保留其他工作表。
    - 新增了重新建立索引 (re-indexing) 的步驟，以確保 Excel 中的行號與程式內部的 DataFrame 索引保持一致。
    - 這樣做是為了避免在刪除重複行後，出現索引不連續的問題，從而解決用戶對於快取檔案名稱與 Excel 行號不匹配的困惑。
    """
    logging.info(f"開始對 {file_path} 進行 URL 正規化與去重...")
    
    try:
        with EXCEL_LOCK:
            # 讀取所有工作表以保留它們
            try:
                all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            except FileNotFoundError:
                logging.info(f"檔案 {file_path} 不存在，跳過 URL 正規化與去重。")
                return

            _sheet = utils.BULK_SHEET
            if _sheet not in all_sheets:
                logging.warning(f"在 {file_path} 中找不到 '{_sheet}' 工作表，跳過處理。")
                return

            bulk_df = all_sheets[_sheet]
            if bulk_df.empty:
                logging.info(f"'{_sheet}' 工作表為空，無需處理。")
                return

            original_row_count = len(bulk_df)
            logging.info(f"原始 '{_sheet}' 工作表包含 {original_row_count} 行。")

            # --- 正規化步驟 ---
            # 正規化 TMDB URL 和 Video URL 欄中的 TMDB 連結，去除 slug（例如 /tv/301944-marriage-toxin → /tv/301944）。
            # 【2026-11】記錄有幾多格真係改咗：以前只係有重複 / 空行先儲存，淨係正規化咗嘅網址唔會寫返入 Excel
            urls_normalized = 0
            for _col in ('Video URL', 'TMDB URL'):
                _before = bulk_df.get(_col, pd.Series(dtype=str)).fillna('').astype(str)
                _after = _before.apply(normalize_tmdb_url)
                urls_normalized += int((_before.str.strip() != _after.str.strip()).sum())
                bulk_df[_col] = _after

            # --- 去重步驟 ---
            # 優先使用 'TMDB URL' 作為去重鍵。如果為空，則使用 'Video URL' (如果它是TMDB URL)。
            # 如果兩者都不是 TMDB URL，則給予一個唯一的鍵以避免被錯誤地刪除。
            bulk_df.reset_index(inplace=True) # 保留原始索引以便生成唯一鍵
            def get_dedup_key(row):
                tmdb_url = row.get('TMDB URL', '')
                video_url = row.get('Video URL', '')
                if tmdb_url and 'themoviedb.org' in tmdb_url:
                    return tmdb_url
                if video_url and 'themoviedb.org' in video_url:
                    return video_url
                # 使用原始行索引創建唯一鍵，確保非TMDB行不會被錯誤刪除
                return f"unique_row_{row['index']}"

            bulk_df['dedup_key'] = bulk_df.apply(get_dedup_key, axis=1)

            bulk_df.drop_duplicates(subset=['dedup_key'], keep='first', inplace=True)
            duplicates_removed = original_row_count - len(bulk_df)

            # 【2026-11】刪除空行（兩個網址都冇），下面嘅行上移，唔會再被處理同寫 Done=0
            _blank_mask = bulk_df.apply(_is_blank_bulk_row, axis=1)
            blank_removed = int(_blank_mask.sum())
            if blank_removed:
                bulk_df = bulk_df[~_blank_mask]
            
            # 【關鍵修改】重新建立索引，使其從 0 開始連續
            bulk_df.reset_index(drop=True, inplace=True)
            logging.info("正在重新索引 DataFrame 以確保行號與索引一致...")

            # 刪除臨時的輔助欄位
            bulk_df.drop(columns=['dedup_key', 'index'], inplace=True, errors='ignore')

            final_row_count = len(bulk_df)
            rows_removed = original_row_count - final_row_count

            if rows_removed > 0 or urls_normalized > 0:
                logging.info(f"URL 正規化與去重完成。正規化了 {urls_normalized} 個網址，移除了 {duplicates_removed} 個重複的 TMDB 項目、{blank_removed} 個空行。")
                all_sheets[_sheet] = bulk_df

                write_excel_atomic(file_path, all_sheets)
                logging.info(f"已將正規化、去重並重新索引後的 '{_sheet}' 工作表 ({final_row_count} 行) 保存回 {file_path}。")
            else:
                logging.info("網址已經係正規格式，亦未發現重複項目或空行，無需修改檔案。")

    except Exception as e:
        logging.error(f"在 URL 正規化與去重過程中發生錯誤: {e}\n{traceback.format_exc()}")
        logging.warning("程式將繼續使用原始（未經處理）的 bulk_videos.xlsx 檔案。")

def _read_monitor_run_options():
    """【2026-11】由監控面板啟動時（環境變數 TMDB_MONITOR_SETTINGS 指向 monitor_settings.json），
    讀取面板目前的覆蓋選項。回傳 {key: value}（分集「同劇集」→ None）；非面板啟動或讀取失敗回傳 {}。"""
    path = os.environ.get('TMDB_MONITOR_SETTINGS')
    if not path or not os.path.isfile(path):
        return {}
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        logging.getLogger('importer').warning(f"[Single] 讀取監控面板選項失敗，沿用目前設定: {e}")
        return {}
    out = {}
    for key in ('overwrite_titles', 'overwrite_descriptions'):
        if data.get(key) in ('y', 'n', 'ask'):
            out[key] = data[key]
    for key in ('episode_overwrite_titles', 'episode_overwrite_descriptions'):
        val = data.get(key)
        if val in ('y', 'n', 'ask'):
            out[key] = val
        elif val is not None:          # 「同劇集」或其他 → 沿用劇集設定
            out[key] = None
    for key in ('clear_when_source_empty', 'wiki_name_replace'):
        if data.get(key) in ('y', 'n'):
            out[key] = data[key]
    return out


def _match_bulk_sheet(answer, names):
    """名稱（先完全相同，再唔分大細楷）或 1 起計嘅編號 → 工作表名稱；搵唔到回傳 None。"""
    if answer.isdigit() and 1 <= int(answer) <= len(names):
        return names[int(answer) - 1]
    if answer in names:
        return answer
    return next((n for n in names if n.strip().lower() == answer.lower()), None)


def resolve_bulk_sheets(answer, file_path='data/bulk_videos.xlsx'):
    """【2026-11】多張工作表：答案用逗號分隔（例如 "Bulk, Eng" / "1, 13"），按輸入次序處理；重複嘅只計一次。
    全部無效 / 空白 → ['Bulk']。"""
    parts = [p.strip() for p in re.split(r'[,，;；]', str(answer or '')) if p.strip()]
    if not parts:
        return [utils.DEFAULT_BULK_SHEET]
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
    except Exception as e:
        logging.warning(f"讀取 {file_path} 工作表清單失敗（{e}），使用默認工作表 '{utils.DEFAULT_BULK_SHEET}'。")
        return [utils.DEFAULT_BULK_SHEET]
    sheets = []
    for p in parts:
        s = _match_bulk_sheet(p, names)
        if s is None:
            logging.warning(f"bulk_videos.xlsx 冇工作表 '{p}'，略過。")
            print(f"警告：bulk_videos.xlsx 冇工作表 '{p}'，略過。")
        elif s not in sheets:
            sheets.append(s)
    if not sheets:
        print(f"警告：冇有效工作表，使用默認工作表 '{utils.DEFAULT_BULK_SHEET}'。")
    return sheets or [utils.DEFAULT_BULK_SHEET]


def _switch_bulk_sheet(next_sheet, excel_lock, log, file_path='data/bulk_videos.xlsx'):
    """【2026-11】轉去下一張工作表。呼叫時所有任務已完成；先確保 Done 全部寫入（Done 用行號對應工作表），
    再清走 details_{行號}.xlsx 快取，最後重新讀取共享緩存。回傳 False = 中斷。"""
    while pending_done_count() > 0:
        if utils._PROGRAM_INTERRUPTED:
            return False
        flush_pending_done_updates(force=True)
        if pending_done_count() > 0:
            log.warning(f"[工作表] 仍有 {pending_done_count()} 筆 Done 值未寫入 '{utils.BULK_SHEET}'"
                        f"（Excel 可能開住），10 秒後重試，寫入後先轉去 '{next_sheet}'。")
            for _ in range(10):
                if utils._PROGRAM_INTERRUPTED:
                    return False
                time.sleep(1)
    with excel_lock:
        utils.BULK_SHEET = next_sheet
    check_create_bulk_excel(file_path)
    normalize_and_deduplicate_urls(file_path)
    try:
        if os.path.exists(DETAILS_CACHE_DIR):
            shutil.rmtree(DETAILS_CACHE_DIR)
    except Exception as e:
        log.warning(f"[工作表] 清理快取目錄 '{DETAILS_CACHE_DIR}' 失敗：{e}")
    os.makedirs(DETAILS_CACHE_DIR, exist_ok=True)
    _refresh_shared_cache(excel_lock, file_path, retry_interval=2, max_retries=3)
    return True


def main(mode='single', overwrite_titles='ask', overwrite_descriptions='ask', clear_when_source_empty='n',
         episode_overwrite_titles=None, episode_overwrite_descriptions=None, wiki_name_replace='n',
         bulk_sheet=None):
    """
    【重構】程式主入口點，負責設置和管理執行緒。
    【修改】更新 single 模式下的邏輯，以處理 extract_and_save_details 新增的回傳值，
    並在內容相同時顯示提示訊息，而不是嘗試上傳。
    【修正】為 shutil.rmtree 新增錯誤處理，以避免因權限問題導致程式崩潰。
    【修正】在 bulk 模式開始時，增加對 normalize_and_deduplicate_urls 的呼叫，以實現 URL 正規化與去重功能。
【新功能】在 bulk 模式下，完成一輪任務後，會自動重新掃描 Excel 檔案以處理新增的資料行，不再詢問用戶。
    【使用者要求 2025-09-18】在 single 模式的主循環中增加中斷檢查。
    【使用者要求 2025-09-18】修改 bulk 模式的主循環，使其對 CTRL+C 更敏感。
    【修改】將 _PROGRAM_INTERRUPTED 和 signal_handler 相關邏輯移至 utils.py。
    【使用者要求】在 Single 模式初始化瀏覽器時，加入 ensure_tmdb_accessible 檢查。
    【使用者要求】將 Done 值的批次寫入邏輯移至 excel_handler.py，並在此處管理其生命週期。
    """
    from others.credentials import login # Explicit import
    # 【修改】使用具名 logger
    log = logging.getLogger('importer') 
    global tmdb_login_attempted_this_session

    # 【修改】將 signal handler 移至 utils
    signal.signal(signal.SIGINT, utils.signal_handler)
    # 【2026-11】第一次 Ctrl+C = 安全停止；第二次 Ctrl+C = 立即強制結束（即使主執行緒卡住）
    utils.install_force_exit_handler()
    # 【修改】傳遞 utils 中的中斷狀態
    utils.set_utils_interrupt_callback(lambda: utils._PROGRAM_INTERRUPTED)
    utils.clear_pause_flag()
    # 【2026-11-13】Bangumi token 到期前 30 日起每次運行都提醒（監控面板問題分頁會見到 WARNING）
    try:
        from others.bangumi_token import token_expiry_message
        _bgm_msg = token_expiry_message()
        if _bgm_msg:
            log.warning(f"[Bangumi] {_bgm_msg}")
    except Exception as e:
        log.debug(f"檢查 Bangumi token 到期日失敗: {e}")

    # 【2026-11 N1】來源無資料時：n（預設）= 保留 TMDB 內容（簡體轉繁體）；y = 覆蓋設為 y 時清空 TMDB 欄位
    utils.CLEAR_TMDB_WHEN_SOURCE_EMPTY = str(clear_when_source_empty or 'n').strip().lower().startswith('y')
    # 【2026-11】維基香港譯名是否取代 TMDB 現有香港人名：n（預設）= 保留 TMDB
    utils.WIKI_NAME_REPLACES_EXISTING = str(wiki_name_replace or 'n').strip().lower().startswith('y')
    # 【2026-11】bulk 模式處理邊一張工作表（默認 Bulk；工作表唔存在就用返 Bulk）
    if mode == 'bulk':
        utils.BULK_SHEETS = resolve_bulk_sheets(bulk_sheet)
        utils.BULK_SHEET = utils.BULK_SHEETS[0]
    # 【2026-11 P2】分集覆蓋選項（None/空 = 沿用劇集設定）
    utils.EPISODE_OVERWRITE_TITLES = episode_overwrite_titles
    utils.EPISODE_OVERWRITE_DESCRIPTIONS = episode_overwrite_descriptions
    log.info(f"程式運行參數: mode={mode}, 劇集 overwrite_titles={overwrite_titles}, overwrite_descriptions={overwrite_descriptions}, "
             f"分集 overwrite_titles={utils.episode_overwrite('title', overwrite_titles)}, "
             f"overwrite_descriptions={utils.episode_overwrite('description', overwrite_descriptions)}, "
             f"clear_when_source_empty={'y' if utils.CLEAR_TMDB_WHEN_SOURCE_EMPTY else 'n'}, "
             f"wiki_name_replace={'y' if utils.WIKI_NAME_REPLACES_EXISTING else 'n'}, bulk_sheet={', '.join(utils.BULK_SHEETS)}")

    configs = load_configs()
    tmdb_extractor, api_key, uses_api = get_tmdb_extractor(configs)

    if mode == 'bulk':
        log.info("進入並行 Bulk 模式")
        print("進入並行 Bulk 模式")
        
        # 【新增】啟動 Done 值批次寫入的背景執行緒
        start_batch_update_worker(excel_lock=EXCEL_LOCK)
        # 【2026-11 L1】Ctrl+C 安全停止逾時、自動結束前，先把記憶體中的 Done 值寫入 Excel
        utils.register_shutdown_hook(lambda: flush_pending_done_updates(force=True), timeout=40)

        # 確保 bulk_videos.xlsx 存在（含 Bulk sheet），否則共享緩存初始讀取會失敗
        check_create_bulk_excel('data/bulk_videos.xlsx')

        # 【新增】在啟動共享緩存前，先對 bulk_videos.xlsx 進行 URL 正規化與去重
        # 這樣初始緩存讀取才不會因檔案不存在或缺 sheet 而失敗
        normalize_and_deduplicate_urls('data/bulk_videos.xlsx')

        # 啟動共享緩存更新執行緒（初始讀取 + 每30秒自動 refresh）
        # extractor threads 從此緩存讀取，完全不持有 EXCEL_LOCK，消除讀寫競爭
        start_shared_cache_refresher(excel_lock=EXCEL_LOCK)
        
        # 清理舊的詳細資料快取
        log.info(f"正在清理並重建快取目錄: {DETAILS_CACHE_DIR}")
        if os.path.exists(DETAILS_CACHE_DIR):
            try:
                shutil.rmtree(DETAILS_CACHE_DIR)
            except PermissionError as e:
                log.warning(f"無法自動刪除快取目錄 '{DETAILS_CACHE_DIR}'：{e}")
                log.warning("請手動檢查是否有檔案總管或終端機正在使用此目錄。程式將繼續執行...")
            except Exception as e:
                log.error(f"清理快取目錄 '{DETAILS_CACHE_DIR}' 時發生未預期的錯誤：{e}")
        
        # 確保目錄存在
        os.makedirs(DETAILS_CACHE_DIR, exist_ok=True)

        NUM_EXTRACTORS = 5 if uses_api else 1
        extract_queue = queue.Queue()  # 所有 threads 共用同一個 queue，誰空閒誰取下一個 job
        upload_queue = queue.Queue()

        # 追蹤已被 worker 取走但尚未完成的 row，防止主執行緒在 queue 暫時清空時重複入隊
        in_progress_indices = set()
        in_progress_lock = threading.Lock()

        # API 模式：所有 extractor 共用同一個瀏覽器實例（避免開啟多個 browser 視窗）
        # 非 API 模式：只有 1 個 extractor，不需要共享（保持原有行為）
        if uses_api:
            shared_ext_driver = [None]
            shared_ext_lock = threading.Lock()
        else:
            shared_ext_driver = None
            shared_ext_lock = None

        # 所有 extractor threads 共用同一個 extract_queue（next available 模式）
        extractors = [
            threading.Thread(
                target=extractor_worker,
                name=f"ExtractorThread_{i}",
                args=(extract_queue, upload_queue, EXCEL_LOCK, tmdb_extractor, api_key, uses_api, configs, overwrite_titles, overwrite_descriptions, in_progress_indices, in_progress_lock, shared_ext_driver, shared_ext_lock),
                daemon=False
            )
            for i in range(NUM_EXTRACTORS)
        ]
        uploader = threading.Thread(target=uploader_worker, name="UploaderThread", args=(upload_queue, EXCEL_LOCK, configs, overwrite_titles, overwrite_descriptions), kwargs={'extractor_count': NUM_EXTRACTORS, 'in_progress_indices': in_progress_indices, 'in_progress_lock': in_progress_lock, 'api_key': api_key}, daemon=False)

        for extractor in extractors:
            extractor.start()
        uploader.start()
        
        try:
            # processed_indices：Excel 已確認 Done 有值的 row（永久跳過）
            # claimed_this_session：本輪執行已入隊／處理中的 row，避免上傳尚未寫 Done 時被重新掃描重複提取
            processed_indices = set()
            claimed_this_session = set()
            _sheet_no = 0
            log.info(f"[工作表] 開始處理工作表 '{utils.BULK_SHEET}'（第 1/{len(utils.BULK_SHEETS)} 張）")

            while not utils._PROGRAM_INTERRUPTED:
                # 重新載入並掃描 bulk_videos.xlsx，將未完成任務加入提取佇列
                log.info("新一輪任務開始：重新載入並掃描 bulk_videos.xlsx...")
                file_path = 'data/bulk_videos.xlsx'

                with EXCEL_LOCK:
                    bulk_df = pd.read_excel(file_path, sheet_name=utils.BULK_SHEET, engine='openpyxl')
                _set_shared_cache(bulk_df)  # 【2026-11 F2】同步共享緩存
                log.info(f"成功重新載入 bulk_videos.xlsx，發現 {len(bulk_df)} 行。")

                tasks_added_this_round = 0
                queued_this_round = set()  # 本輪已加入佇列的索引，防止同一輪重複入隊
                for index, row in bulk_df.iterrows():
                    if utils._PROGRAM_INTERRUPTED:
                        log.warning("在添加任務到隊列時檢測到中斷信號。")
                        break

                    if index in processed_indices or index in claimed_this_session:
                        continue

                    if _is_blank_bulk_row(row):   # 【2026-11】空行唔處理，下次整理 Excel 時刪除
                        continue

                    done_value = row.get('Done', '')
                    done_str = str(done_value).strip()
                    if not(pd.isna(done_value) or done_str == '' or done_str.lower() in ['nan', '0', '0.0']):
                        processed_indices.add(index)
                        continue

                    pending = get_pending_done_value(index)
                    if pending is not None:  # 本 session 已處理（Done=1 成功或 Done=0 失敗），均不重試
                        processed_indices.add(index)
                        continue

                    if index in queued_this_round:
                        continue

                    with in_progress_lock:
                        if index in in_progress_indices:
                            continue

                    extract_queue.put({'row_index': index})
                    queued_this_round.add(index)
                    claimed_this_session.add(index)
                    tasks_added_this_round += 1

                if utils._PROGRAM_INTERRUPTED:
                    break

                if tasks_added_this_round > 0:
                    log.info(f"已將 {tasks_added_this_round} 個新任務放入提取隊列。")
                else:
                    log.info("未發現新的待處理任務。")

                # 等待所有任務完成。
                # 同時每 30 秒掃描一次 Excel：若有新增項目，立即加入提取佇列，
                # 無需等待上載完成——提取與上載可並行執行。
                log.info("等待當前批次任務處理中，同時定期偵測 Excel 新增項目...")
                _queues_empty_since = None
                _last_excel_rescan = time.monotonic()
                while True:
                    if utils._PROGRAM_INTERRUPTED:
                        log.warning("在等待佇列清空時檢測到中斷信號。")
                        break
                    with in_progress_lock:
                        _q_empty = extract_queue.empty() and upload_queue.empty()
                        _in_progress_snapshot = set(in_progress_indices)
                        all_idle = _q_empty and not _in_progress_snapshot
                    if all_idle:
                        break
                    _now = time.monotonic()
                    if _q_empty:
                        if _queues_empty_since is None:
                            _queues_empty_since = _now
                        elif _now - _queues_empty_since > 60:
                            log.info(
                                f"佇列已清空，仍在等待 {len(_in_progress_snapshot)} 個進行中任務完成: "
                                f"{sorted(_in_progress_snapshot)[:20]}"
                                f"{'...' if len(_in_progress_snapshot) > 20 else ''}"
                            )
                            _queues_empty_since = _now
                    else:
                        _queues_empty_since = None

                    # 每 30 秒掃描 Excel，若有新增項目立即加入提取佇列（不等待上載完成）
                    if _now - _last_excel_rescan >= 30:
                        _last_excel_rescan = _now
                        try:
                            with EXCEL_LOCK:
                                _rescan_df = pd.read_excel(file_path, sheet_name=utils.BULK_SHEET, engine='openpyxl')
                            _set_shared_cache(_rescan_df)  # 【2026-11 F2】同步共享緩存，避免新項目被跳過
                            _new_found = 0
                            for _idx, _row in _rescan_df.iterrows():
                                if utils._PROGRAM_INTERRUPTED:
                                    break
                                if _idx in processed_indices or _idx in claimed_this_session or _idx in queued_this_round:
                                    continue
                                if _is_blank_bulk_row(_row):   # 【2026-11】空行唔處理
                                    continue
                                _dv = _row.get('Done', '')
                                _ds = str(_dv).strip()
                                if not (pd.isna(_dv) or _ds == '' or _ds.lower() in ['nan', '0', '0.0']):
                                    processed_indices.add(_idx)
                                    continue
                                _pend = get_pending_done_value(_idx)
                                if _pend is not None:  # 本 session 已處理，均不重試
                                    processed_indices.add(_idx)
                                    continue
                                with in_progress_lock:
                                    if _idx in in_progress_indices:
                                        continue
                                extract_queue.put({'row_index': _idx})
                                queued_this_round.add(_idx)
                                claimed_this_session.add(_idx)
                                tasks_added_this_round += 1
                                _new_found += 1
                            if _new_found > 0:
                                log.info(f"[定期掃描] 發現 {_new_found} 個新增項目，已加入提取佇列（上載仍在進行中）。")
                        except Exception as _scan_err:
                            log.warning(f"[定期掃描] 重新掃描 Excel 時發生錯誤: {_scan_err}")

                    time.sleep(1)

                if utils._PROGRAM_INTERRUPTED:
                    break

                log.info("當前批次任務已全部完成。")

                # 若本輪（含定期掃描）均無新任務，代表呢張工作表已處理完畢 → 轉下一張；冇下一張就自動結束
                if tasks_added_this_round == 0 and _sheet_no + 1 < len(utils.BULK_SHEETS):
                    _done_sheet = utils.BULK_SHEET
                    _sheet_no += 1
                    _next_sheet = utils.BULK_SHEETS[_sheet_no]
                    log.info(f"[工作表] 工作表 '{_done_sheet}' 已全部完成，轉去下一張 '{_next_sheet}'。")
                    print(f"工作表 '{_done_sheet}' 已全部完成，轉去下一張 '{_next_sheet}'。")
                    if not _switch_bulk_sheet(_next_sheet, EXCEL_LOCK, log):
                        break
                    processed_indices = set()
                    claimed_this_session = set()
                    log.info(f"[工作表] 開始處理工作表 '{utils.BULK_SHEET}'（第 {_sheet_no + 1}/{len(utils.BULK_SHEETS)} 張）")
                    continue
                if tasks_added_this_round == 0:
                    log.info("所有任務已完成，程式將自動結束。")
                    print("所有任務已完成，程式將自動結束。")
                    break

                # 有任務完成後，等待 10 秒再掃描，確保不遺漏最後一批新增項目
                log.info("所有當前任務已處理完畢。將在 10 秒後重新掃描檔案以尋找新增項目...")
                for _ in range(10):
                    if utils._PROGRAM_INTERRUPTED:
                        break
                    time.sleep(1)

                if utils._PROGRAM_INTERRUPTED:
                    log.info("在等待重新掃描期間偵測到中斷信號，程序即將結束。")
                    break

            # 向每條提取執行緒各發送一個結束信號（共 NUM_EXTRACTORS 個 None）
            for _ in range(NUM_EXTRACTORS):
                extract_queue.put(None)

        except KeyboardInterrupt:
            utils._PROGRAM_INTERRUPTED = True
            log.info("主執行緒捕獲到 Ctrl+C，等待工作執行緒結束...")
        except Exception as e:
            utils._PROGRAM_INTERRUPTED = True
            log.error(f"主執行緒發生錯誤: {e}\n{traceback.format_exc()}")
        finally:
            # 先把已完成的 Done 值寫入 Excel，避免 join 期間再按 Ctrl+C 導致記憶體快取全部遺失
            try:
                flush_pending_done_updates(force=True)
            except Exception as flush_err:
                log.error(f"中斷時立即寫入 Done 值失敗: {flush_err}")

            try:
                for extractor in extractors:
                    if extractor.is_alive():
                        extract_queue.put(None)
                        extractor.join(timeout=15)
                if uploader.is_alive():
                    uploader.join(timeout=30)
            except KeyboardInterrupt:
                log.warning("等待工作執行緒期間再次收到中斷，繼續清理並寫入 Done 值。")

            if shared_ext_driver is not None and shared_ext_driver[0]:
                try:
                    cleanup_browser(shared_ext_driver[0])
                    shared_ext_driver[0] = None
                except Exception:
                    pass

            try:
                stop_shared_cache_refresher()
                stop_batch_update_worker()
            except KeyboardInterrupt:
                log.warning("停止 worker 期間再次收到中斷，最後再嘗試寫入 Done 值。")
                try:
                    flush_pending_done_updates(force=True)
                except Exception:
                    pass

            # 【2026-11 F1】join 逾時後仍有執行緒卡住 → 關閉本程式開啟的瀏覽器並結束，
            # 否則 Python 會在結束時無限等待這些非 daemon 執行緒（殘留 python + 瀏覽器）。
            utils.exit_if_threads_stuck(list(extractors) + [uploader])
            log.info("所有工作執行緒已結束。")

    else: # single mode (循序執行)
        log.info("進入 Single 模式")
        print("進入 Single 模式")
        driver_session = None

        def _is_tmdb_url(url):
            return isinstance(url, str) and 'themoviedb.org' in url

        def ensure_single_driver_ready():
            """啟動（或重啟）單一模式瀏覽器並登入 TMDB。上傳前必須成功。"""
            nonlocal driver_session
            if driver_session is not None:
                try:
                    _ = driver_session.title
                    return True
                except Exception:
                    log.warning("[Single] 偵測到瀏覽器已斷開，正在重新啟動...")
                    try:
                        cleanup_browser(driver_session)
                    except Exception:
                        pass
                    driver_session = None

            try:
                log.info("[Single] 正在啟動瀏覽器並登入...")
                driver_tuple = get_browser_driver(profile_suffix='_single', remote_port=9224)
                driver_session = driver_tuple[0] if driver_tuple else None
            except Exception as e:
                log.error(f"[Single] 啟動瀏覽器失敗: {e}")
                driver_session = None

            if not driver_session:
                log.error("[Single] 無法取得瀏覽器 driver。")
                return False

            ensure_tmdb_accessible(driver_session)
            if not is_tmdb_logged_in(driver_session):
                creds = get_tmdb_credentials(configs)
                if creds[0] and creds[1]:
                    login(driver_session, *creds)
                    log.info("[Single] 已完成 TMDB 登入。")
                else:
                    log.warning("[Single] 未找到 TMDB 憑證，無法自動登入。")
            return True

        def do_single_upload(final_tmdb_url, output_excel, is_movie):
            """確保瀏覽器已就緒後再上傳。回傳 upload_* 的結果。"""
            if not ensure_single_driver_ready():
                log.error(f"[Single] driver 為 None，無法上傳。URL: {final_tmdb_url}")
                print("無法啟動瀏覽器，上傳已取消。")
                return False
            if "/person/" in final_tmdb_url:
                return upload_person_to_tmdb(driver_session, final_tmdb_url, output_excel, overwrite_titles, overwrite_descriptions)
            return upload_to_tmdb(driver_session, final_tmdb_url, output_excel, is_movie, overwrite_titles, overwrite_descriptions)

        try:
            while not utils._PROGRAM_INTERRUPTED:
                # 【修改】在每次循環開始時重新載入字典
                utils.reload_dictionaries_and_converters()
                
                video_url = get_user_input_with_timeout("請輸入 Video site 的 URL、'excel' 直接從 Excel 上傳，或 'eof' 結束: ")
                if video_url.lower() == 'eof': break

                tmdb_url_input = get_user_input_with_timeout("請輸入 TMDB 的 URL (可留空): ")

                utils.wait_if_paused(log, 'Single')
                if utils._PROGRAM_INTERRUPTED:
                    break

                # 【2026-11】由監控面板啟動時：以輸入 TMDB URL 當刻面板上的覆蓋選項為準
                _latest = _read_monitor_run_options()
                if _latest:
                    _changes = []
                    def _pick(key, cur):
                        if key not in _latest:
                            return cur
                        new = _latest[key]
                        if new != cur:
                            _changes.append(f"{key}: {cur or '同劇集'} → {new or '同劇集'}")
                            return new
                        return cur
                    overwrite_titles = _pick('overwrite_titles', overwrite_titles)
                    overwrite_descriptions = _pick('overwrite_descriptions', overwrite_descriptions)
                    utils.EPISODE_OVERWRITE_TITLES = _pick('episode_overwrite_titles', utils.EPISODE_OVERWRITE_TITLES)
                    utils.EPISODE_OVERWRITE_DESCRIPTIONS = _pick('episode_overwrite_descriptions', utils.EPISODE_OVERWRITE_DESCRIPTIONS)
                    _clear = _pick('clear_when_source_empty', 'y' if utils.CLEAR_TMDB_WHEN_SOURCE_EMPTY else 'n')
                    utils.CLEAR_TMDB_WHEN_SOURCE_EMPTY = _clear == 'y'
                    _wiki = _pick('wiki_name_replace', 'y' if utils.WIKI_NAME_REPLACES_EXISTING else 'n')
                    utils.WIKI_NAME_REPLACES_EXISTING = _wiki == 'y'
                    if _changes:
                        log.info("[Single] 已套用監控面板最新選項: " + "; ".join(_changes))
                        print("已套用監控面板最新選項: " + "; ".join(_changes))

                output_excel = os.path.join(DETAILS_CACHE_DIR, "details_single_mode.xlsx")

                # 若 video_url 為空或 'excel'，且已有現成 Excel 且 TMDB URL 已提供，直接上傳，不重新抓取
                video_url_clean = clean_url(video_url)
                tmdb_url_clean = clean_url(tmdb_url_input)
                use_existing_excel = (
                    not video_url_clean or video_url_clean.lower() == 'excel'
                ) and os.path.exists(output_excel) and not is_details_file_empty(output_excel)

                if use_existing_excel and tmdb_url_clean:
                    final_tmdb_url = tmdb_url_clean
                    is_movie = "/movie/" in final_tmdb_url
                    log.info(f"[Single] 偵測到現有 Excel，跳過提取直接上傳: {output_excel} → {final_tmdb_url}")
                    print(f"使用現有 Excel 直接上傳至 TMDB: {final_tmdb_url}")
                    do_single_upload(final_tmdb_url, output_excel, is_movie)
                elif use_existing_excel and not tmdb_url_clean:
                    log.warning("[Single] 現有 Excel 已存在，但未提供 TMDB URL，無法上傳。")
                    print("請輸入 TMDB URL 以上傳現有 Excel 的內容。")
                else:
                    # 提取階段：只有影片站、或非 API 的 TMDB 頁面，才需要瀏覽器。
                    # 僅貼 TMDB URL + API 模式時，提取不開瀏覽器；上傳前再啟動。
                    needs_browser_for_extract = is_valid_video_url(video_url) or (
                        not uses_api and (_is_tmdb_url(video_url) or bool(tmdb_url_input))
                    )
                    if needs_browser_for_extract and not ensure_single_driver_ready():
                        log.error("[Single] 提取需要瀏覽器，但啟動失敗。")
                        print("無法啟動瀏覽器，提取已取消。")
                        continue

                    # Single 模式使用無限重試，比較不會因為 Excel 還開著就一直失敗
                    safe_remove_file(output_excel, max_retries=None, interactive_prompt=True)
                    success, final_tmdb_url, is_movie, no_chinese, is_identical = extract_and_save_details(
                        driver_session, video_url_clean, tmdb_url_clean, '', tmdb_extractor,
                        api_key, uses_api, configs, overwrite_titles, overwrite_descriptions, -1, output_excel
                    )

                    if success:
                        if is_identical and not no_chinese:
                            log.info(f"提取成功，但內容與 TMDB 上已有的 zh-HK 理想版本相同，無需上傳。")
                            print("提取成功，但內容與 TMDB 上已有的 zh-HK 理想版本相同，無需上傳。")
                        elif not no_chinese and final_tmdb_url:
                            log.info(f"提取成功，準備上傳至 TMDB: {final_tmdb_url}")
                            do_single_upload(final_tmdb_url, output_excel, is_movie)
                        elif no_chinese:
                            log.info(f"提取成功，但因無中文內容而被跳過，不上傳。")
                            print("提取成功，但因無中文內容而被跳過，不上傳。")
                        else:
                            log.warning(f"提取成功，但未找到有效的 TMDB URL，無法上傳。")
                            print("提取成功，但未找到有效的 TMDB URL，無法上傳。")
                    else:
                        log.error(f"提取失敗。")
                        print("提取失敗。")
                
                # 【使用者要求】在 single 模式的循環中增加中斷檢查
                if utils._PROGRAM_INTERRUPTED:
                    log.warning("檢測到中斷信號，即將退出 single 模式循環。")
                    break

                if (get_user_input_with_timeout("是否繼續? (y/n): ") or 'y').lower() != 'y':
                    break
        finally:
            if driver_session:
                cleanup_browser(driver_session)

    log.info("程序已安全結束。")

if __name__ == "__main__":
    param_overwrite_titles = get_user_input_with_timeout("是否覆蓋劇集標題 (系列/季/電影)？(y/n/ask, 默認 ask): ", timeout=600) or 'ask'
    param_overwrite_descriptions = get_user_input_with_timeout("是否覆蓋劇集描述 (系列/季/電影)？(y/n/ask, 默認 ask): ", timeout=600) or 'ask'
    param_ep_overwrite_titles = get_user_input_with_timeout("是否覆蓋分集標題？(y/n/ask, 默認 = 同劇集): ", timeout=600) or None
    param_ep_overwrite_descriptions = get_user_input_with_timeout("是否覆蓋分集描述？(y/n/ask, 默認 = 同劇集): ", timeout=600) or None
    param_clear_when_source_empty = get_user_input_with_timeout("來源無資料時是否清空 TMDB 現有標題/描述？(y/n, 默認 n = 保留 TMDB，簡體會轉繁體): ", timeout=600) or 'n'
    param_wiki_name_replace = get_user_input_with_timeout("人名：是否用維基香港譯名取代 TMDB 現有香港人名？(y/n, 默認 n = 保留 TMDB): ", timeout=600) or 'n'
    param_mode_input = get_user_input_with_timeout("請選擇運行模式 (single(s)/bulk(b), 默認 single): ", timeout=600) or 's'
    
    current_mode = 'bulk' if param_mode_input.startswith('b') else 'single'
    param_bulk_sheet = None
    if current_mode == 'bulk':
        param_bulk_sheet = get_user_input_with_timeout("bulk 模式要處理 bulk_videos.xlsx 邊一張工作表？(名稱或編號；多張用逗號分隔，按次序處理，例如 1, 13；默認 Bulk): ", timeout=600) or None
    
    main(mode=current_mode, overwrite_titles=param_overwrite_titles, overwrite_descriptions=param_overwrite_descriptions,
         clear_when_source_empty=param_clear_when_source_empty,
         wiki_name_replace=param_wiki_name_replace,
         bulk_sheet=param_bulk_sheet,
         episode_overwrite_titles=param_ep_overwrite_titles,
         episode_overwrite_descriptions=param_ep_overwrite_descriptions)