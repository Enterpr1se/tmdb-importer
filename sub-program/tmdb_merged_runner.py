import asyncio
import logging
import os
import re
import sys
import time
import json
import random
import queue
import threading
import subprocess
import sqlite3
import pandas as pd
import openpyxl
import httpx
import traceback
import itertools
import glob
import shutil
import ctypes  # 用於檢查管理員權限
from datetime import datetime
from urllib.parse import urljoin
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bulk_txt import write_urls_txt, output_path  # sub-program 共用：輸出到 txt，唔直接改 bulk_videos.xlsx

# ==========================================
# 全域設定與常數
# ==========================================

sys.stdout.reconfigure(encoding='utf-8')

script_dir = os.path.dirname(os.path.abspath(__file__))

# --- [設定] OpenVPN 自動切換設定 ---
OPENVPN_CONFIG_DIR = "" 
OPENVPN_BIN_PATH = "" # 程式會自動偵測並填入此變數

# --- 智慧路徑偵測 ---
possible_roots = [
    os.path.dirname(script_dir),
    script_dir
]

project_root = script_dir
config_file_path = None

for path in possible_roots:
    check_path = os.path.join(path, 'configs.json')
    if os.path.exists(check_path):
        project_root = path
        config_file_path = check_path
        break

if config_file_path is None:
    project_root = os.path.dirname(script_dir)
    config_file_path = os.path.join(project_root, 'configs.json')

print(f"[Init] 專案根目錄: {project_root}")
print(f"[Init] 設定檔路徑: {config_file_path}")

DB_FILE_PATH = os.path.join(project_root, 'data', 'scan_history.db')
# 搵到嘅網址輸出到 sub-program/tmdb_merged_runner_urls.txt（見 bulk_txt.py），唔直接改 bulk_videos.xlsx
MISSAV_NAMES_FILE = None   # main() 設定：MissAV 書籤小程式下載嘅名單
MISSAV_SEARCHED_FILE = os.path.join(project_root, 'data', 'missav_searched_names.txt')

VIDEO_TO_PERSON_TABLE = 'videos_to_people_history'
VIDEO_TO_PERSON_RETENTION = 30
PERSON_TO_VIDEO_TABLE = 'people_to_videos_history'
PERSON_TO_VIDEO_RETENTION = 60
CHINESE_CHECK_TABLE = 'chinese_check_history'
CHINESE_CHECK_RETENTION = 365  # 365 天（同 bulk_video_person.py 一致；兩個共用 scan_history.db，清理時用短嘅會刪走另一個嘅紀錄）

MAX_CONCURRENT_REQUESTS = 20
TMDB_API_BASE_URL = "https://api.themoviedb.org/3"

EXCLUDED_ACCOUNTS = {'enterpr1se'}

url_queue = queue.Queue()
stop_event = threading.Event()
excel_lock = threading.Lock()
session_seen_urls = set()
session_seen_lock = threading.Lock()

# 全域統計與結果暫存
session_stats = {'added': 0, 'updated': 0}
stats_lock = threading.Lock()
all_valid_urls = [] 
results_lock = threading.Lock()

# VPN 鎖與狀態旗標
vpn_lock = threading.Lock()
vpn_feature_working = True  # 預設為 True，會在 main 中根據使用者選擇調整

# 加入 \r 確保每條 Log 都從行首開始印，避免接在進度條後面
logging.basicConfig(level=logging.INFO, format='\r%(asctime)s - %(message)s', datefmt='%H:%M:%S')

# 關閉 httpx 和 httpcore 的 INFO 日誌
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

# ==========================================
# 共用工具函式 (Config, DB, Excel, VPN)
# ==========================================

def load_configs():
    if os.path.exists(config_file_path):
        try:
            # 使用 utf-8-sig 以兼容帶有 BOM 的檔案
            with open(config_file_path, 'r', encoding='utf-8-sig') as f:
                data = json.load(f)
                return data
        except Exception as e:
            logging.error(f"讀取設定檔失敗 ({config_file_path}): {e}")
            return {}
    else:
        logging.warning(f"找不到設定檔: {config_file_path}")
    return {}

def save_configs(configs):
    try:
        os.makedirs(os.path.dirname(config_file_path), exist_ok=True)
        with open(config_file_path, 'w', encoding='utf-8') as f:
            json.dump(configs, f, indent=4, ensure_ascii=False)
            logging.info(f"設定已儲存至: {config_file_path}")
    except Exception as e:
        logging.error(f"儲存設定檔失敗: {e}")

def normalize_tmdb_url(url):
    if not isinstance(url, str): return None
    match = re.search(r'(?:https?://)?(?:www\.)?themoviedb\.org/(movie|tv|person)/(\d+)', url)
    if match:
        return f"https://www.themoviedb.org/{match.group(1)}/{match.group(2)}"
    return url

def parse_tmdb_url(url):
    match = re.search(r'themoviedb\.org/(movie|tv|person)/(\d+)', url)
    if match:
        return match.group(1), match.group(2)
    return None, None

def init_db():
    try:
        os.makedirs(os.path.dirname(DB_FILE_PATH), exist_ok=True)
        with sqlite3.connect(DB_FILE_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA synchronous=NORMAL;")
            cursor = conn.cursor()
            def create_table(table_name):
                cursor.execute(f'''
                    CREATE TABLE IF NOT EXISTS {table_name} (
                        tmdb_id_str TEXT PRIMARY KEY,
                        timestamp REAL
                    )
                ''')
                cursor.execute(f'CREATE INDEX IF NOT EXISTS idx_{table_name}_ts ON {table_name}(timestamp)')
            create_table(VIDEO_TO_PERSON_TABLE)
            create_table(PERSON_TO_VIDEO_TABLE)
            create_table(CHINESE_CHECK_TABLE)
            conn.commit()
    except Exception as e:
        logging.error(f"[DB] 初始化失敗: {e}")

def check_db_history(key, table_name, retention_days):
    cutoff_time = time.time() - (retention_days * 24 * 60 * 60)
    try:
        with sqlite3.connect(DB_FILE_PATH, timeout=10) as conn:
            cursor = conn.cursor()
            cursor.execute(f'SELECT 1 FROM {table_name} WHERE tmdb_id_str = ? AND timestamp > ?', (key, cutoff_time))
            return cursor.fetchone() is not None
    except Exception:
        return False

def save_db_record(table_name, key):
    try:
        with sqlite3.connect(DB_FILE_PATH, timeout=10) as conn:
            cursor = conn.cursor()
            cursor.execute(f'INSERT OR REPLACE INTO {table_name} (tmdb_id_str, timestamp) VALUES (?, ?)', (key, time.time()))
            conn.commit()
    except Exception as e:
        logging.error(f"[DB] 寫入失敗 ({table_name}): {e}")

def cleanup_and_optimize_db(vacuum=False):
    tables_to_clean = {
        VIDEO_TO_PERSON_TABLE: VIDEO_TO_PERSON_RETENTION,
        PERSON_TO_VIDEO_TABLE: PERSON_TO_VIDEO_RETENTION,
        CHINESE_CHECK_TABLE: CHINESE_CHECK_RETENTION
    }
    try:
        logging.info("\n[DB] 開始檢查過期資料...")
        with sqlite3.connect(DB_FILE_PATH, timeout=30) as conn:
            cursor = conn.cursor()
            total_deleted = 0
            for table_name, retention_days in tables_to_clean.items():
                cutoff_time = time.time() - (retention_days * 24 * 60 * 60)
                cursor.execute(f'DELETE FROM {table_name} WHERE timestamp < ?', (cutoff_time,))
                deleted_count = cursor.rowcount
                logging.info(f"[DB] 表格 {table_name}: 清除 {deleted_count} 筆過期紀錄 (保留期限: {retention_days}天)")
                total_deleted += deleted_count
            conn.commit()
            if vacuum:
                logging.info("[DB] 執行資料庫壓縮 (VACUUM)...")
                conn.execute('VACUUM')
                logging.info("[DB] 資料庫壓縮完成。")
    except Exception as e:
        logging.error(f"[DB] 清理資料庫錯誤: {e}")

def display_progress(step_name, current, total, start_time):
    elapsed = time.time() - start_time
    rate = current / elapsed if elapsed > 0 else 0
    remaining_items = total - current
    eta_seconds = remaining_items / rate if rate > 0 else 0
    
    eta_str = f"{int(eta_seconds//60)}m{int(eta_seconds%60)}s"
    msg = f"\r   -> {step_name} 進度: {current}/{total} - ETA: {eta_str}"
    sys.stdout.write(msg.ljust(80))
    sys.stdout.flush()

def write_to_excel(valid_urls):
    """搵到嘅網址寫入 txt（名稱保留係為咗唔使改 call site）。"""
    if not valid_urls:
        return
    _, added = write_urls_txt(valid_urls, 'tmdb_merged_runner')
    with stats_lock:
        session_stats['added'] += added

# --- [修改] 自動切換 OpenVPN 函式 ---

def is_admin():
    """檢查是否具有管理員權限 (Windows)"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def request_admin():
    """請求提升管理員權限"""
    if sys.platform == 'win32':
        try:
            cwd = os.getcwd()
            # 使用 ShellExecuteW 啟動新行程，verb="runas"
            ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{os.path.abspath(__file__)}"', cwd, 1)
            # 如果回傳值 > 32 代表成功執行
            return ret > 32
        except:
            return False
    return False

def detect_openvpn_info():
    """
    偵測 OpenVPN 執行檔路徑和類型 (GUI/CLI)
    Returns: (path, type) where type is 'gui' or 'cli' or None
    """
    if OPENVPN_BIN_PATH:
        if "gui" in os.path.basename(OPENVPN_BIN_PATH).lower():
            return OPENVPN_BIN_PATH, 'gui'
        return OPENVPN_BIN_PATH, 'cli'
        
    candidates_gui = [
        r"C:\Program Files\OpenVPN\bin\openvpn-gui.exe",
        r"C:\Program Files (x86)\OpenVPN\bin\openvpn-gui.exe"
    ]
    for c in candidates_gui:
        if os.path.exists(c):
            return c, 'gui'
    
    if shutil.which("openvpn-gui"):
        return "openvpn-gui", 'gui'

    candidates_cli = [
        r"C:\Program Files\OpenVPN\bin\openvpn.exe",
        r"C:\Program Files (x86)\OpenVPN\bin\openvpn.exe"
    ]
    for c in candidates_cli:
        if os.path.exists(c):
            return c, 'cli'
            
    if shutil.which("openvpn"):
        return "openvpn", 'cli'
        
    return None, None

def rotate_vpn():
    """
    自動切換 OpenVPN 節點
    Returns: True (成功/嘗試切換), False (失敗/設定錯誤/無權限)
    """
    global vpn_feature_working, OPENVPN_BIN_PATH
    
    # 1. 決定設定檔目錄
    config_dir = OPENVPN_CONFIG_DIR
    if not config_dir:
        # 嘗試預設路徑 C:\Users\<User>\OpenVPN\config
        user_home = os.path.expanduser("~")
        default_config = os.path.join(user_home, "OpenVPN", "config")
        if os.path.exists(default_config):
            config_dir = default_config
    
    if not config_dir or not os.path.exists(config_dir):
        logging.warning("[VPN] 未設定且找不到 OpenVPN 設定檔目錄，略過切換。")
        vpn_feature_working = False
        return False

    # 2. 決定執行檔路徑 (優先尋找 GUI)
    bin_path = OPENVPN_BIN_PATH
    if not bin_path:
        # 如果全域變數空了，再偵測一次
        b_path, b_type = detect_openvpn_info()
        if b_path:
            bin_path = b_path
            OPENVPN_BIN_PATH = b_path

    if not bin_path:
        logging.warning("[VPN] 找不到 OpenVPN 執行檔 (openvpn-gui.exe 或 openvpn.exe)！")
        vpn_feature_working = False
        return False

    # 判斷是否為 GUI 模式
    is_gui_mode = "openvpn-gui" in os.path.basename(bin_path).lower()

    # [關鍵修改] 權限檢查：只有 CLI 模式必須要 Admin
    if sys.platform == "win32" and not is_admin() and not is_gui_mode:
        logging.error("[VPN] ⚠️ 錯誤：OpenVPN CLI 需要「以系統管理員身分執行」才能修改路由！")
        vpn_feature_working = False
        return False
    
    if not vpn_lock.acquire(blocking=False):
        logging.info("[VPN] 另一個線程正在切換 VPN，等待中...")
        return True 

    try:
        current_configs = load_configs()
        vpn_user = current_configs.get('vpn_username')
        vpn_pass = current_configs.get('vpn_password')
        
        # 準備帳密檔 (僅 CLI 模式有用)
        auth_file_path = None
        if vpn_user and vpn_pass:
            auth_file_path = os.path.abspath(os.path.join(project_root, '_temp_vpn_auth.txt'))
            try:
                with open(auth_file_path, 'w', encoding='utf-8') as f:
                    f.write(f"{vpn_user}\n{vpn_pass}")
            except Exception as e:
                logging.error(f"[VPN] 無法建立驗證檔案: {e}")
                if not is_gui_mode: 
                    vpn_feature_working = False
                    return False
        elif not is_gui_mode:
            logging.warning("[VPN] configs.json 中缺少 vpn_username 或 vpn_password！")
            vpn_feature_working = False
            return False

        logging.info(f"\n[VPN] 正在啟動 OpenVPN 節點切換程序 (模式: {'GUI' if is_gui_mode else 'CLI'})...")
        
        ovpn_files = glob.glob(os.path.join(config_dir, "**", "*.ovpn"), recursive=True)
        if not ovpn_files:
            logging.error("[VPN] 找不到任何 .ovpn 設定檔！")
            vpn_feature_working = False
            return False

        new_config = random.choice(ovpn_files)
        config_filename = os.path.basename(new_config)
        config_cwd = os.path.dirname(new_config)
        
        logging.info(f"[VPN] 選定新節點: {config_filename}")

        # --- 執行切換 ---
        
        if is_gui_mode:
            # GUI 模式：發送指令到背景程序，並隱藏 stdout/stderr
            logging.info("[VPN] 發送 GUI 斷線指令...")
            subprocess.run([bin_path, "--command", "disconnect_all"], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            time.sleep(3) 

            logging.info(f"[VPN] 發送 GUI 連線指令...")
            subprocess.Popen([bin_path, "--connect", config_filename, "--config_dir", config_cwd], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
        else:
            # CLI 模式
            logging.info("[VPN] 正在終止舊的 OpenVPN CLI 連線...")
            if sys.platform == "win32":
                subprocess.call("taskkill /F /IM openvpn.exe", shell=True, stderr=subprocess.DEVNULL)
            else:
                subprocess.call("killall openvpn", shell=True, stderr=subprocess.DEVNULL)
            
            time.sleep(2)

            logging.info(f"[VPN] 正在啟動 CLI 連線...")
            cmd = [bin_path, "--config", new_config]
            
            if auth_file_path and os.path.exists(auth_file_path):
                cmd.extend(["--auth-user-pass", auth_file_path])
            
            # [修改] 靜默啟動
            if sys.platform == "win32":
                subprocess.Popen(cmd, cwd=config_cwd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            else:
                subprocess.Popen(["sudo"] + cmd, cwd=config_cwd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        logging.info("[VPN] 等待 15 秒建立連線...")
        time.sleep(15)
        logging.info("[VPN] 節點切換指令已發送。")
        return True

    except Exception as e:
        logging.error(f"[VPN] 切換失敗: {e}")
        return False
    finally:
        vpn_lock.release()

def play_sound_alert():
    try:
        print("\n\033[91m [!!!] 注意：已連續 5 次被限流，請考慮更換 IP 或檢查網路！ [!!!] \033[0m")
        if sys.platform == "win32":
            try:
                import winsound
                for _ in range(5):
                    winsound.Beep(1000, 400)
                    time.sleep(0.1)
            except ImportError:
                print('\a' * 5)
        elif sys.platform == "darwin":
            try:
                subprocess.run(["afplay", "/System/Library/Sounds/Ping.aiff"], check=False)
            except:
                print('\a' * 5)
        else:
            print('\a' * 5)
    except:
        print('\a')

# ==========================================
# PART 1: Scraper 邏輯 (Thread 1 - Playwright)
# ==========================================

def parse_complex_pages(input_str):
    if not input_str or input_str.lower() == "all":
        return None, True 
    pages = set()
    parts = re.split(r'[,\s，]+', input_str)
    try:
        for part in parts:
            part = part.strip()
            if not part: continue
            if "-" in part:
                start_s, end_s = part.split("-", 1)
                start, end = int(start_s), int(end_s)
                if start > end: start, end = end, start
                for p in range(start, end + 1):
                    pages.add(p)
            else:
                pages.add(int(part))
        if not pages: return None, True
        return sorted(list(pages)), False
    except:
        return None, True

def is_tmdb_throttled(html):
    txt = html.lower()
    return "temporarily throttled" in txt or "wait and try again" in txt or "retry-after" in txt or "are being rate limited" in txt

def check_404_title(title):
    if re.search(r"\b404\b", title): return True
    if "page not found" in title.lower(): return True
    return False

def has_next_page(soup):
    nav = soup.find("nav", class_="pagy") or soup.find("div", class_="pagination_wrapper")
    if not nav: return False
    next_btn = nav.find("a", {"aria-label": re.compile(r"next|下一頁", re.I)}) or \
               nav.find("a", string=re.compile(r">|下一頁|next", re.I))
    if not next_btn: return False
    if next_btn.get("aria-disabled") == "true": return False
    return bool(next_btn.get("href"))

def wait_countdown(seconds, msg):
    print()
    for i in range(seconds, 0, -1):
        if stop_event.is_set():
            print(f"\n[System] 偵測到中斷信號，停止 {msg}。")
            return
        sys.stdout.write(f"\r [Scraper] {msg}: {i} 秒...   ")
        sys.stdout.flush()
        time.sleep(1)
    sys.stdout.write(f"\r [Scraper] {msg} 結束，繼續執行...               \n")

def handle_tmdb_popups(page):
    try:
        for selector in ['#onetrust-accept-btn-handler', 'button:has-text("Accept All Cookies")', '.js-accept-cookies']:
            if page.locator(selector).is_visible(timeout=500):
                page.locator(selector).click()
                time.sleep(0.5)
                break
        for btn in ['#cf_exit_close', '.close-button', '#cf_exit_modal .close']:
            if page.locator(btn).is_visible(timeout=500):
                try: page.locator(btn).click(timeout=1000)
                except: page.evaluate(f'document.querySelector("{btn}").click()')
                time.sleep(0.5)
    except: pass

def scraper_safe_navigate(page, url, desc, on_throttle_callback=None):
    throttle_count = 0 
    
    while not stop_event.is_set():
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            handle_tmdb_popups(page)
            try:
                page.wait_for_load_state("networkidle", timeout=3000)
            except: pass

            content = page.content()
            if is_tmdb_throttled(content):
                throttle_count += 1
                logging.warning(f"[Scraper] ⚠️ 載入 {desc} 被限流 (連續 {throttle_count} 次)。")
                if throttle_count > 0 and throttle_count % 5 == 0:
                    if vpn_feature_working:
                        if not rotate_vpn():
                            play_sound_alert()
                    else:
                        play_sound_alert()
                if on_throttle_callback:
                    on_throttle_callback()
                wait_countdown(60, "冷卻中")
                continue 
            
            throttle_count = 0
            title = page.title()
            if "Just a moment" in title or "Cloudflare" in title:
                logging.warning(f"[Scraper] ⚠️ 遇到 Cloudflare 驗證...")
                wait_countdown(20, "等待驗證")
                if "Just a moment" in page.title():
                    continue
                return True
            return True
        except PlaywrightTimeoutError:
            logging.warning(f"[Scraper] ⚠️ 載入 {desc} 超時。")
            wait_countdown(60, "網路超時重試")
            continue
        except Exception as e:
            if "target closed" in str(e).lower(): raise e 
            logging.error(f"[Scraper] 導航錯誤: {e}")
            wait_countdown(60, "連線錯誤重試")
            continue
    return False

def tmdb_login_robust(page, username, password):
    if not username or not password: return False
    logging.info(f"[Scraper] >>> 正在啟動登入流程 (帳號: {username})...")
    throttle_count = 0 

    while not stop_event.is_set():
        try:
            logging.info("[Scraper] 前往登入頁面...")
            page.goto("https://www.themoviedb.org/login", wait_until="domcontentloaded", timeout=60000)
            handle_tmdb_popups(page)
            if "/u/" in page.url:
                logging.info("[Scraper] 🎉 偵測到已登入狀態！")
                return True

            content = page.content()
            if is_tmdb_throttled(content):
                throttle_count += 1
                logging.warning(f"[Scraper] 登入頁面被限流 (連續 {throttle_count} 次)。")
                if throttle_count > 0 and throttle_count % 5 == 0: 
                    if vpn_feature_working:
                        if not rotate_vpn():
                            play_sound_alert()
                    else:
                        play_sound_alert()
                wait_countdown(60, "登入頁面被限流")
                continue

            if "Just a moment" in page.title() or "Cloudflare" in page.title():
                wait_countdown(20, "等待驗證")
                continue

            logging.info("[Scraper] 填寫帳密...")
            try:
                page.wait_for_selector('#username', state='visible', timeout=10000)
                page.fill('#username', username)
                page.fill('#password', password)
            except:
                logging.warning("[Scraper] 找不到帳號密碼欄位，重試...")
                continue
            
            logging.info("[Scraper] 點擊登入按鈕...")
            if page.locator('#login_button').is_visible():
                page.click('#login_button')
            else:
                page.keyboard.press("Enter")
            
            logging.info("[Scraper] 等待跳轉至個人頁面...")
            try:
                page.wait_for_url("**/u/**", timeout=20000)
                logging.info("[Scraper] 🎉 TMDB 登入成功 (URL 驗證)！")
                return True
            except PlaywrightTimeoutError:
                logging.warning(f"[Scraper] 等待跳轉超時，當前網址: {page.url}")
                try:
                    if page.locator('header .avatar').count() > 0 or page.locator(f'header a[href^="/u/"]').count() > 0:
                        logging.info("[Scraper] 🎉 TMDB 登入成功 (Header 驗證)！")
                        return True
                except: pass

            content = page.content()
            if is_tmdb_throttled(content):
                throttle_count += 1
                logging.warning(f"[Scraper] 提交後被限流 (連續 {throttle_count} 次)。")
                if throttle_count > 0 and throttle_count % 5 == 0: 
                    if vpn_feature_working:
                        if not rotate_vpn():
                            play_sound_alert()
                    else:
                        play_sound_alert()
                wait_countdown(60, "登入提交冷卻中")
                continue

            throttle_count = 0
            if "Invalid username and/or password" in content or "無效的使用者名稱" in content:
                logging.error("[Scraper] 登入失敗：帳號或密碼錯誤 (請檢查 configs.json)。")
                return False

            logging.warning("[Scraper] 登入未成功，準備重試...")
            wait_countdown(10, "重新嘗試登入")
            continue
        except Exception as e:
            if "target closed" in str(e).lower(): raise e
            logging.error(f"[Scraper] 登入過程錯誤: {e}")
            wait_countdown(60, "錯誤重試")
            continue
    return False

def scraper_extract_links(soup):
    links = set()
    exclude = ["season", "episode", "company", "network", "collection", "discuss", "edit?", "images", "videos", "rating", "watchlist", "changes", "remote", "export", "list"]
    for a in soup.find_all("a", href=re.compile(r"^/(movie|tv|person)/\d+")):
        href = a["href"]
        if any(ex in href for ex in exclude): continue
        links.add(normalize_tmdb_url("https://www.themoviedb.org" + href))
    return links

def process_found_links(links, source_desc=""):
    total = len(links)
    session_dupes = 0
    added = 0
    
    for link in links:
        norm_link = normalize_tmdb_url(link)
        if not norm_link: continue
        
        with session_seen_lock:
            if norm_link in session_seen_urls:
                session_dupes += 1
                continue
            session_seen_urls.add(norm_link)
        
        m_type, m_id = parse_tmdb_url(norm_link)
        if m_type and m_id:
            url_queue.put(norm_link)
            added += 1
    
    if added > 0:
        logging.info(f"[Scraper] 來自 {source_desc} 發現 {total} 連結 (新增任務: {added}, Session過濾: {session_dupes}) -> 已送往 Processor")
    elif total > 0:
        logging.info(f"[Scraper] 來自 {source_desc} 發現 {total} 連結 (全部已在本次Session處理過)")

# === 特殊任務類別與函式 ===

class JavHdScraper:
    def __init__(self, page):
        self.page = page
        self.base_url = "https://javhd.com/ja/model"
        self.names = set()

    def scrape(self):
        cats = ["popular", "justadded", "top"]
        logging.info("[Scraper] 開始 JAVHD 抓取...")
        try:
            self.page.goto(self.base_url, timeout=60000)
            self._handle_popups()
        except: pass

        for cat in cats:
            if stop_event.is_set(): break 
            url = f"{self.base_url}/{cat}"
            self._scrape_cat(url)
        return list(self.names)

    def _handle_popups(self):
        try:
            for sel in ['.js-warningPopupAgreeBtn', '#onesignal-slidedown-cancel-button', '.js-jcookie-allow']:
                if self.page.locator(sel).is_visible(timeout=1000):
                    self.page.locator(sel).click()
        except: pass

    def _scrape_cat(self, url):
        try:
            self.page.goto(url, timeout=60000)
            self._handle_popups()
        except: return
        
        page_num = 1
        while not stop_event.is_set(): 
            self._handle_popups()
            try:
                self.page.wait_for_selector('.section__thumbs-list', timeout=5000)
            except: break

            elements = self.page.locator('.thumb--model .thumb__title').all_inner_texts()
            if not elements: break
            
            new_names = [n.strip() for n in elements if n.strip()]
            for n in new_names: self.names.add(n)
            
            next_btn = self.page.locator('.list-pagination__button--next')
            if not next_btn.is_visible(timeout=2000) or "disable" in str(next_btn.get_attribute("class")):
                break
            next_btn.click()
            page_num += 1
            time.sleep(random.uniform(1.0, 2.0))

def load_missav_names(path):
    """讀 MissAV 書籤小程式（missav_bookmarklet.html）下載嘅 missav_names.txt；
    略過以前已經搜尋過嘅名（記錄喺 data/missav_searched_names.txt）。"""
    try:
        with open(path, encoding='utf-8-sig') as f:
            names = list(dict.fromkeys(line.strip() for line in f if line.strip()))
    except OSError as e:
        logging.error(f"[Scraper] 讀取 MissAV 名單失敗: {e}")
        return []
    searched = set()
    if os.path.exists(MISSAV_SEARCHED_FILE):
        with open(MISSAV_SEARCHED_FILE, encoding='utf-8') as f:
            searched = {line.strip() for line in f if line.strip()}
    todo = [n for n in names if n not in searched]
    logging.info(f"[Scraper] MissAV 名單 {len(names)} 個名，其中 {len(names) - len(todo)} 個以前搜尋過，今次搜尋 {len(todo)} 個。")
    return todo


def mark_missav_searched(name):
    try:
        os.makedirs(os.path.dirname(MISSAV_SEARCHED_FILE), exist_ok=True)
        with open(MISSAV_SEARCHED_FILE, 'a', encoding='utf-8') as f:
            f.write(name + '\n')
    except OSError:
        pass


def find_latest_missav_file():
    downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    files = glob.glob(os.path.join(downloads, 'missav_names*.txt'))
    return max(files, key=os.path.getmtime) if files else None


def search_people_tmdb(page, names, on_searched=None):
    logging.info(f"[Scraper] 正在 TMDB 搜尋 {len(names)} 位人物...")
    for name in names:
        if stop_event.is_set(): break 
        url = f"https://www.themoviedb.org/search/person?query={name}"
        if scraper_safe_navigate(page, url, f"搜尋 {name}"):
            soup = BeautifulSoup(page.content(), "html.parser")
            res = soup.select('.search_results.person .card .header a, .item.profile .name a')
            links = set()
            for a in res:
                href = a.get('href')
                if href and '/person/' in href: 
                    links.add("https://www.themoviedb.org" + href)
            process_found_links(links, f"搜尋:{name}")
            if on_searched: on_searched(name)
        time.sleep(random.uniform(1.5, 3.0))

def fetch_homepage_content_full(page):
    url = "https://www.themoviedb.org/"
    logging.info(f"[Scraper] >>> 正在前往首頁 ({url})...")
    
    if not scraper_safe_navigate(page, url, "首頁"):
        return
    
    logging.info("[Scraper] 首頁載入成功，開始執行自動操作 (捲動/點擊)...")
    
    for _ in range(3):
        if stop_event.is_set(): return 
        page.mouse.wheel(0, 1000)
        time.sleep(0.3)
    page.mouse.wheel(0, -5000) 
    time.sleep(0.5)

    soup = BeautifulSoup(page.content(), "html.parser")
    process_found_links(scraper_extract_links(soup), "首頁初始")

    anchors = page.locator('.selector .anchor h3 a')
    count = anchors.count()
    logging.info(f"[Scraper] 偵測到 {count} 個分類標籤，開始遍歷...")

    for i in range(count):
        if stop_event.is_set(): break 
        try:
            anchor = anchors.nth(i)
            label = anchor.inner_text().strip()
            handle_tmdb_popups(page)
            try:
                anchor.click(force=True, timeout=2000)
            except:
                anchor.evaluate("node => node.click()")
            
            time.sleep(random.uniform(1.2, 2.0))
            if is_tmdb_throttled(page.content()):
                wait_countdown(60, "操作被限流")
                continue

            page.evaluate("""() => {
                document.querySelectorAll('.scroller').forEach(div => {
                    div.scrollLeft = div.scrollWidth;
                });
            }""")
            time.sleep(1.0) 
            
            soup = BeautifulSoup(page.content(), "html.parser")
            process_found_links(scraper_extract_links(soup), f"首頁-{label}")
        except Exception:
            pass

# === Scraper 主執行緒 ===

def run_scraper_thread(configs, input_lines, target_pages_list, fetch_all_pages):
    logging.info("[Scraper] 執行緒啟動")
    
    tasks = set()
    usernames = set()
    
    for line in input_lines:
        parts = re.split(r'[,\s]+', line)
        for p in parts:
            p = p.strip()
            if not p: continue
            if p in ['1', 'homepage']: tasks.add('homepage')
            elif p in ['2', 'leaderboard']: tasks.add('leaderboard')
            elif p in ['3', 'javhd']: tasks.add('javhd')
            elif p.lower() in ['4', 'missav']: tasks.add('missav')
            elif '/u/' in p: usernames.add(p.split('/u/')[1].split('/')[0].split('?')[0])
            else: usernames.add(p)

    tmdb_user = configs.get('tmdb_username')
    tmdb_pass = configs.get('tmdb_password')

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # 強制設定 User-Agent，解決 headless 模式下登入被擋的問題
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080},
            locale="zh-HK"
        )
        page = context.new_page()

        if tmdb_user and tmdb_pass:
            if not tmdb_login_robust(page, tmdb_user, tmdb_pass):
                logging.warning("[Scraper] 登入失敗，將以訪客身份繼續。")
        
        if 'javhd' in tasks and not stop_event.is_set():
            jav = JavHdScraper(page)
            names = jav.scrape()
            if names: search_people_tmdb(page, names)

        if 'missav' in tasks and not stop_event.is_set() and MISSAV_NAMES_FILE:
            names = load_missav_names(MISSAV_NAMES_FILE)
            if names: search_people_tmdb(page, names, on_searched=mark_missav_searched)

        if 'homepage' in tasks and not stop_event.is_set():
            fetch_homepage_content_full(page)

        if 'leaderboard' in tasks and not stop_event.is_set():
            if scraper_safe_navigate(page, "https://www.themoviedb.org/leaderboard", "排行榜"):
                soup = BeautifulSoup(page.content(), "html.parser")
                for a in soup.select('a[href^="/u/"]'):
                    href = a.get("href", "")
                    if any(x in href for x in ["/u/new", "/u/login", "/u/signup"]): continue
                    u = href.split('/u/')[1].split('/')[0]
                    if u in EXCLUDED_ACCOUNTS:
                        logging.info(f"[Scraper] 已略過排除帳號: {u}")
                        continue
                    usernames.add(u)

        total_users = len(usernames)
        for idx, user in enumerate(usernames, 1):
            if stop_event.is_set(): break 
            logging.info(f"\n[Scraper] ({idx}/{total_users}) 準備掃描使用者: {user}")
            
            if fetch_all_pages:
                iter_pages = itertools.count(1)
            else:
                iter_pages = target_pages_list

            empty_consecutive = 0
            session_count = 0
            rest_threshold = random.randint(10, 12)

            def reset_session_count():
                nonlocal session_count
                session_count = 0

            for cur in iter_pages:
                if stop_event.is_set(): break 
                logging.info(f"[Scraper] [{idx}/{total_users}] User: [{user}] - 正在處理第 {cur} 頁...")
                
                if session_count > 0 and session_count >= rest_threshold:
                    wait_countdown(random.randint(5, 10), "休息中")
                    session_count = 0
                    rest_threshold = random.randint(10, 12)

                url = f"https://www.themoviedb.org/u/{user}/activity?page={cur}"
                
                if not scraper_safe_navigate(page, url, f"{user} P{cur}", on_throttle_callback=reset_session_count):
                    break
                
                if check_404_title(page.title()):
                    logging.info(f"[Scraper] User: [{user}] - 第 {cur} 頁 404 (結束)")
                    break

                soup = BeautifulSoup(page.content(), "html.parser")
                links = scraper_extract_links(soup)
                
                if links:
                    process_found_links(links, f"{user} P{cur}")
                    empty_consecutive = 0
                else:
                    logging.info(f"[Scraper] User: [{user}] - 第 {cur} 頁掃描到 0 個有效連結 (可能是空頁或全過濾)")
                    empty_consecutive += 1
                    if fetch_all_pages and empty_consecutive >= 20:
                        logging.info(f"[Scraper] User: [{user}] - 連續 20 頁無新資料，停止。")
                        break
                
                session_count += 1
                
                if not has_next_page(soup):
                    logging.info(f"[Scraper] User: [{user}] - 已達最後一頁。")
                    break
                
                time.sleep(random.uniform(1.5, 3.0))

    logging.info("[Scraper] 所有抓取任務完成，等待 Processor 清空佇列...")
    stop_event.set()

# ==========================================
# PART 2: Processor 邏輯 (Thread 2 - Asyncio)
# ==========================================

async def tmdb_api_get(session, url, params, context=""):
    try:
        resp = await session.get(url, params=params, timeout=20)
        if resp.status_code == 429:
            retry = int(resp.headers.get("Retry-After", 10))
            logging.warning(f"[API] ⚠️ 429 限流 ({context})，等待 {retry}s")
            
            # API 端也嘗試切換 VPN（如果連續限流）
            # 這裡簡單處理：如果 retry > 30s 可能已經很嚴重
            if retry > 30:
                if vpn_feature_working:
                    rotate_vpn()
                
            await asyncio.sleep(retry)
            return await tmdb_api_get(session, url, params, context)
        if resp.status_code == 404: return None
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        logging.warning(f"[API] 請求失敗 ({context}): {e}")
        return None

# --- Bulk Logic Steps ---

async def step_get_person_credits(session, semaphore, person_id, api_key, check_cache=True):
    key = str(person_id)
    if check_cache and check_db_history(key, PERSON_TO_VIDEO_TABLE, PERSON_TO_VIDEO_RETENTION):
        return set()
    
    url = f"{TMDB_API_BASE_URL}/person/{person_id}/combined_credits"
    async with semaphore:
        data = await tmdb_api_get(session, url, {'api_key': api_key}, f"Person {person_id}")
        save_db_record(PERSON_TO_VIDEO_TABLE, key)
        videos = set()
        if data:
            for item in data.get('cast', []) + data.get('crew', []):
                if item.get('media_type') in ['movie', 'tv']:
                    videos.add((item['media_type'], item['id']))
        return videos

async def step_get_video_credits(session, semaphore, v_type, v_id, api_key, check_cache=True):
    key = f"{v_type}_{v_id}"
    if check_cache and check_db_history(key, VIDEO_TO_PERSON_TABLE, VIDEO_TO_PERSON_RETENTION):
        return set()
    
    url = f"{TMDB_API_BASE_URL}/{v_type}/{v_id}/credits" if v_type == 'movie' else f"{TMDB_API_BASE_URL}/tv/{v_id}/aggregate_credits"
    async with semaphore:
        data = await tmdb_api_get(session, url, {'api_key': api_key}, f"{v_type} {v_id}")
        save_db_record(VIDEO_TO_PERSON_TABLE, key)
        people = set()
        if data:
            for item in data.get('cast', []) + data.get('crew', []):
                people.add(item['id'])
        return people

async def step_check_chinese(session, semaphore, v_type, v_id, api_key, check_cache=True):
    key = f"{v_type}_{v_id}"
    full_url = f"https://www.themoviedb.org/{v_type}/{v_id}"
    
    if check_cache and check_db_history(key, CHINESE_CHECK_TABLE, CHINESE_CHECK_RETENTION):
        return None 

    async with semaphore:
        trans = await tmdb_api_get(session, f"{TMDB_API_BASE_URL}/{v_type}/{v_id}/translations", {'api_key': api_key}, "Trans")
        has_zh = False
        if trans:
            for t in trans.get('translations', []):
                if t.get("iso_639_1") == "zh" and t.get("iso_3166_1") in ['HK', 'TW', 'CN']:
                    has_zh = True
                    break
        
        if not has_zh:
            details = await tmdb_api_get(session, f"{TMDB_API_BASE_URL}/{v_type}/{v_id}", {'api_key': api_key, 'language': 'en-US'}, "Details")
            if details:
                t_title = details.get('title') or details.get('name') or ""
                o_title = details.get('original_title') or details.get('original_name') or ""
                if (re.search(r'[\u4e00-\u9fff]', t_title)) or (re.search(r'[\u4e00-\u9fff]', o_title)):
                    has_zh = True

        save_db_record(CHINESE_CHECK_TABLE, key)
        return full_url if has_zh else None

async def process_batch_pipeline(session, semaphore, api_key, batch_urls):
    initial_videos = set()
    initial_people = set()

    for url in batch_urls:
        t, i = parse_tmdb_url(url)
        if not t or not i: continue
        if t == 'person': initial_people.add(int(i))
        else: initial_videos.add((t, i))

    logging.info(f"[Batch] 啟動 4 步驟處理: 人物={len(initial_people)}, 影片={len(initial_videos)}")

    # Step 1: Initial Person -> Videos (強制檢查初始人物)
    videos_from_step1 = set()
    if initial_people:
        tasks = [step_get_person_credits(session, semaphore, pid, api_key, check_cache=False) for pid in initial_people]
        total = len(tasks)
        done = 0
        start_t = time.time()
        for f in asyncio.as_completed(tasks):
            res = await f
            videos_from_step1.update(res)
            done += 1
            display_progress("[Batch] Step 1", done, total, start_t)
        print()

    # Step 2: Videos -> People
    videos_for_step2 = initial_videos.union(videos_from_step1)
    people_from_step2 = set()
    if videos_for_step2:
        tasks = []
        skipped_count = 0
        scan_count = 0
        
        for t, i in videos_for_step2:
            is_initial = (t, i) in initial_videos
            if not is_initial and check_db_history(f"{t}_{i}", VIDEO_TO_PERSON_TABLE, VIDEO_TO_PERSON_RETENTION):
                skipped_count += 1
            else:
                scan_count += 1
                tasks.append(step_get_video_credits(session, semaphore, t, i, api_key, check_cache=False))
        
        logging.info(f"   -> [Step 2] 影片擴展人物: 總數 {len(videos_for_step2)}, DB忽略 {skipped_count}, 需掃描 {scan_count}")
        
        if tasks:
            total = len(tasks)
            done = 0
            start_t = time.time()
            for f in asyncio.as_completed(tasks):
                res = await f
                people_from_step2.update(res)
                done += 1
                display_progress("[Batch] Step 2", done, total, start_t)
            print()

    # Step 3: People -> Videos
    all_persons_for_step3 = people_from_step2.union(initial_people)
    videos_from_step3 = set()
    if all_persons_for_step3:
        tasks = []
        skipped_count = 0
        scan_count = 0
        
        for pid in all_persons_for_step3:
            if check_db_history(str(pid), PERSON_TO_VIDEO_TABLE, PERSON_TO_VIDEO_RETENTION):
                skipped_count += 1
            else:
                scan_count += 1
                tasks.append(step_get_person_credits(session, semaphore, pid, api_key, check_cache=False))
        
        logging.info(f"   -> [Step 3] 人物擴展影片: 總數 {len(all_persons_for_step3)}, DB忽略 {skipped_count}, 需掃描 {scan_count}")

        if tasks:
            total = len(tasks)
            done = 0
            start_t = time.time()
            for f in asyncio.as_completed(tasks):
                res = await f
                videos_from_step3.update(res)
                done += 1
                display_progress("[Batch] Step 3", done, total, start_t)
            print()

    # Step 4: Check Chinese
    all_videos_to_check = initial_videos.union(videos_from_step1).union(videos_from_step3)
    if all_videos_to_check:
        tasks = []
        skipped_count = 0
        scan_count = 0
        
        for t, i in all_videos_to_check:
            is_initial = (t, i) in initial_videos
            if not is_initial and check_db_history(f"{t}_{i}", CHINESE_CHECK_TABLE, CHINESE_CHECK_RETENTION):
                skipped_count += 1
            else:
                scan_count += 1
                tasks.append(step_check_chinese(session, semaphore, t, i, api_key, check_cache=False))
        
        logging.info(f"   -> [Step 4] 中文檢查: 總數 {len(all_videos_to_check)}, DB忽略 {skipped_count}, 需掃描 {scan_count}")

        valid_urls = []
        if tasks:
            total = len(tasks)
            done = 0
            start_t = time.time()
            for f in asyncio.as_completed(tasks):
                res = await f
                if res: valid_urls.append(res)
                done += 1
                display_progress("[Batch] Step 4", done, total, start_t)
            print()
        
        # 寫入暫存而不直接寫檔
        if valid_urls:
            with results_lock:
                all_valid_urls.extend(valid_urls)
            logging.info(f"[Processor] 🎉 發現 {len(valid_urls)} 部有效中文影片，已暫存 (目前總數: {len(all_valid_urls)})...")
        else:
            logging.info("[Processor] 本批次無符合條件影片。")

async def processor_loop(api_key):
    logging.info("[Processor] 啟動 Asyncio 迴圈")
    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)
    limits = httpx.Limits(max_keepalive_connections=20, max_connections=20)
    
    async with httpx.AsyncClient(limits=limits, timeout=20.0) as session:
        while True:
            if stop_event.is_set() and url_queue.empty():
                logging.info("[Processor] 工作結束。")
                break
                
            batch = []
            try:
                if batch == [] and not stop_event.is_set():
                    item = url_queue.get(timeout=2)
                    batch.append(item)
                    url_queue.task_done()
                
                # 無上限提取 Queue 中的所有連結
                while True:
                    try:
                        item = url_queue.get_nowait()
                        batch.append(item)
                        url_queue.task_done()
                    except queue.Empty: break
            except queue.Empty:
                if stop_event.is_set():
                    logging.info("[Processor] 工作結束。")
                    break
                continue

            if batch:
                await process_batch_pipeline(session, semaphore, api_key, batch)

def run_processor_thread(api_key):
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    asyncio.run(processor_loop(api_key))

# ==========================================
# Main
# ==========================================

def main():
    print("\n=== TMDB 並行抓取與處理系統 (Merged) ===")
    
    configs = load_configs()
    init_db()
    
    api_key = configs.get('tmdb_api_v3')
    if not api_key:
        api_key = input("請手動輸入 TMDB API Key (v3): ").strip()
        if api_key:
            configs['tmdb_api_v3'] = api_key
            save_configs(configs)
        else: return
    else:
        print(f"✅ 已載入 API Key")

    if not configs.get('tmdb_username'):
        u = input("請輸入 TMDB Username (可選, 按 Enter 跳過): ").strip()
        if u:
            p = input("請輸入 TMDB Password: ").strip()
            configs['tmdb_username'] = u
            configs['tmdb_password'] = p
            save_configs(configs)

    # [修改] 詢問是否啟用自動 VPN (需要管理員權限)
    global vpn_feature_working, OPENVPN_BIN_PATH
    use_vpn_input = input("\n是否啟用自動 VPN 切換功能? (y/N) [預設: N]: ").strip().lower()
    
    if use_vpn_input == 'y':
        # 先偵測一次是 GUI 還是 CLI
        OPENVPN_BIN_PATH = "" # Reset to allow re-detection
        _, vpn_type = detect_openvpn_info()
        
        if vpn_type == 'gui':
            print("✅ 偵測到 OpenVPN GUI，將使用服務模式切換 (通常無需管理員權限)。")
            vpn_feature_working = True
        elif vpn_type == 'cli':
            # CLI 模式必須要 Admin
            if is_admin():
                print("✅ 已具備管理員權限，VPN 自動切換功能已啟用 (CLI 模式)。")
                vpn_feature_working = True
            else:
                print("⚠️ 偵測到 OpenVPN CLI 但無管理員權限，正在請求提升權限...")
                if request_admin():
                    print("✅ 已啟動新視窗執行，本視窗將自動關閉。")
                    sys.exit(0)
                else:
                    print("❌ 無法獲取權限。VPN 功能已停用 (聲音提示模式)。")
                    vpn_feature_working = False
        else:
            print("❌ 找不到 OpenVPN 執行檔。VPN 功能已停用 (聲音提示模式)。")
            vpn_feature_working = False
    else:
        print("ℹ️ 使用者選擇不啟用 VPN，功能已停用 (聲音提示模式)。")
        vpn_feature_working = False

    print("\n[輸入 Scraper 指令]")
    print(" - '1' 或 'homepage': 首頁 (Homepage)")
    print(" - '2' 或 'leaderboard': 排行榜 (Leaderboard)")
    print(" - '3' 或 'javhd': JAVHD")
    print(" - '4' 或 'missav': MissAV（讀 missav_bookmarklet.html 書籤小程式下載嘅 missav_names.txt）")
    print(" - 使用者名稱")
    print(" - 直接按 Enter 開始執行")
    
    input_lines = []
    while True:
        line = input(">>> ").strip()
        if not line: break
        input_lines.append(line)
    
    if not input_lines: return

    global MISSAV_NAMES_FILE
    if any(t.strip().lower() in ('4', 'missav') for line in input_lines for t in re.split(r'[,\s]+', line)):
        latest = find_latest_missav_file()
        hint = f"Enter = {latest}" if latest else "下載資料夾冇 missav_names*.txt"
        path = input(f"\nMissAV 名單 txt 路徑（{hint}）: ").strip().strip('"')
        MISSAV_NAMES_FILE = path or latest
        if not MISSAV_NAMES_FILE or not os.path.exists(MISSAV_NAMES_FILE):
            print("❌ 搵唔到 MissAV 名單，略過 MissAV。請先用 sub-program/missav_bookmarklet.html 下載名單。")
            MISSAV_NAMES_FILE = None

    needs_user_scan = False
    all_tokens = []
    for line in input_lines:
        parts = re.split(r'[,\s]+', line)
        for p in parts:
            p = p.strip()
            if p: all_tokens.append(p)

    for token in all_tokens:
        t_low = token.lower()
        if t_low in ['1', 'homepage', '3', 'javhd', '4', 'missav']:
            continue
        if t_low in ['2', 'leaderboard']:
            needs_user_scan = True
            break
        needs_user_scan = True
        break
            
    target_pages_list = None
    fetch_all_pages = False
    
    if needs_user_scan:
        print("\n設定使用者貢獻抓取頁數 (Enter=全部, '5'=第5頁, '1-5'=範圍):")
        pg_input = input(">>> ").strip()
        target_pages_list, fetch_all_pages = parse_complex_pages(pg_input)
    else:
        fetch_all_pages = True

    proc_thread = threading.Thread(target=run_processor_thread, args=(api_key,), daemon=False)
    proc_thread.start()

    scrap_thread = threading.Thread(
        target=run_scraper_thread, 
        args=(configs, input_lines, target_pages_list, fetch_all_pages), 
        daemon=False
    )
    scrap_thread.start()

    try:
        scrap_thread.join()
        url_queue.join()
        proc_thread.join()
    except KeyboardInterrupt:
        print("\n\n[System] 收到強制中斷訊號 (Ctrl+C)...")
        stop_event.set()
    
    # 一次性寫入
    print("\n=== 正在寫入結果 txt (一次性) ===")
    if all_valid_urls:
        write_to_excel(list(set(all_valid_urls))) # 去除重複後寫入
    else:
        print("本次執行沒有發現需寫入的資料。")

    print("\n=== Session Summary ===")
    print(f"Total Added: {session_stats['added']}")
    print(f"結果檔案位置: {output_path('tmdb_merged_runner')}")
    
    # 最後執行 DB 清理與壓縮
    cleanup_and_optimize_db(vacuum=True)

if __name__ == "__main__":
    main()