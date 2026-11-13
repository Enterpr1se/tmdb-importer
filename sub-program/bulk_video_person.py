import asyncio
import logging
import os
import re
import sys
import time
import httpx
import pandas as pd
import json
import openpyxl
import traceback
import sqlite3
from datetime import datetime, timedelta
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bulk_txt import write_urls_txt  # sub-program 共用：輸出到 txt，唔直接改 bulk_videos.xlsx

# --- 路徑設定 ---
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

# --- 設定日誌 ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- API 和並行設定 ---
MAX_CONCURRENT_REQUESTS = 20
TMDB_API_BASE_URL = "https://api.themoviedb.org/3"

# --- 資料庫設定與常量 ---
DB_FILE_PATH = os.path.join(project_root, 'data', 'scan_history.db')

# 擴展步驟 2: 影片抓取演員 (Video -> Person)
VIDEO_TO_PERSON_TABLE = 'videos_to_people_history'
VIDEO_TO_PERSON_RETENTION = 30  # 30 天

# 擴展步驟 3: 演員抓取影片 (Person -> Video)
PERSON_TO_VIDEO_TABLE = 'people_to_videos_history'
PERSON_TO_VIDEO_RETENTION = 60  # 60 天

# 步驟 4: 中文翻譯檢查
CHINESE_CHECK_TABLE = 'chinese_check_history'
CHINESE_CHECK_RETENTION =  365 # 365 天

# --- 設定與 Excel 處理函數 ---

def load_configs():
    """從位於專案根目錄的 configs.json 載入設定。"""
    config_path = os.path.join(project_root, 'configs.json')
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as config_file:
                return json.load(config_file)
        except (json.JSONDecodeError, IOError) as e:
            logging.error(f"讀取設定檔 {config_path} 時發生錯誤: {e}")
            return {}
    return {}

def save_configs(configs):
    """將設定儲存到位於專案根目錄的 configs.json。"""
    config_path = os.path.join(project_root, 'configs.json')
    try:
        with open(config_path, 'w', encoding='utf-8') as config_file:
            json.dump(configs, config_file, indent=4, ensure_ascii=False)
    except IOError as e:
        logging.error(f"儲存設定檔至 {config_path} 時發生錯誤: {e}")

def get_tmdb_api_key(configs):
    """從設定檔或使用者輸入獲取 TMDB API 金鑰。"""
    if 'tmdb_api_v3' not in configs or not configs['tmdb_api_v3']:
        api_key = input("請輸入您的 TMDB API Key (v3): ")
        save_config = input("是否希望儲存此 API Key 留待下次使用? (y/n): ")
        if save_config.lower() == 'y':
            configs['tmdb_api_v3'] = api_key
            save_configs(configs)
    else:
        api_key = configs['tmdb_api_v3']
    return api_key

def parse_tmdb_url(url):
    """從 TMDB URL 中解析出類型（movie, tv, person）和 ID。"""
    match = re.search(r'themoviedb\.org/(movie|tv|person)/(\d+)', url)
    if match:
        return match.group(1), match.group(2)
    return None, None

def contains_chinese(text):
    """檢查字串是否包含中文（排除日文假名）。"""
    if not text or not isinstance(text, str):
        return False
    if re.search(r'[\u3040-\u30ff]', text):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

# --- 通用進度顯示函數 ---
def display_progress(step_name, current, total, start_time, extra_info=""):
    """顯示進度條與 ETA 的通用函數。"""
    elapsed = time.time() - start_time
    rate = current / elapsed if elapsed > 0 else 0
    remaining_items = total - current
    eta_seconds = remaining_items / rate if rate > 0 else 0
    
    eta_str = f"{int(eta_seconds//60)}m{int(eta_seconds%60)}s"
    
    msg = f"\r{step_name} 進度: {current}/{total} - ETA: {eta_str}"
    if extra_info:
        msg += f" - {extra_info}"
        
    sys.stdout.write(msg.ljust(80))
    sys.stdout.flush()

# --- 資料庫管理函數群 (優化版：惰性過期檢查) ---

def init_db():
    """初始化 SQLite 資料庫，啟用 WAL 模式以提升併發性能。"""
    try:
        os.makedirs(os.path.dirname(DB_FILE_PATH), exist_ok=True)
        with sqlite3.connect(DB_FILE_PATH) as conn:
            # 啟用 Write-Ahead Logging (WAL) 模式
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
        logging.error(f"初始化資料庫失敗: {e}")

def get_scanned_keys_batch(table_name, keys_list, retention_days):
    """
    (修改) 批量檢查快取。
    新增 retention_days 參數，在查詢時直接過濾掉過期紀錄。
    """
    if not keys_list:
        return set()
    
    existing_keys = set()
    chunk_size = 900
    
    # 計算有效時間的截止點
    cutoff_time = time.time() - (retention_days * 24 * 60 * 60)
    
    try:
        with sqlite3.connect(DB_FILE_PATH, timeout=10) as conn:
            cursor = conn.cursor()
            for i in range(0, len(keys_list), chunk_size):
                chunk = keys_list[i:i + chunk_size]
                if not chunk: continue
                
                placeholders = ",".join(["?"] * len(chunk))
                # 關鍵修改：SQL 查詢加入 timestamp 判斷
                sql = f"SELECT tmdb_id_str FROM {table_name} WHERE tmdb_id_str IN ({placeholders}) AND timestamp > ?"
                
                # 將 cutoff_time 加入參數列表
                params = chunk + [cutoff_time]
                
                cursor.execute(sql, params)
                rows = cursor.fetchall()
                for row in rows:
                    existing_keys.add(row[0])
                    
    except Exception as e:
        logging.error(f"批量讀取資料庫失敗 ({table_name}): {e}")
        
    return existing_keys

def is_scanned(table_name, key, retention_days):
    """
    (修改) 單筆檢查，加入過期時間判斷。
    """
    cutoff_time = time.time() - (retention_days * 24 * 60 * 60)
    try:
        with sqlite3.connect(DB_FILE_PATH, timeout=5) as conn:
            cursor = conn.cursor()
            # 關鍵修改：SQL 查詢加入 timestamp 判斷
            cursor.execute(f'SELECT 1 FROM {table_name} WHERE tmdb_id_str = ? AND timestamp > ?', (key, cutoff_time))
            return cursor.fetchone() is not None
    except Exception as e:
        logging.error(f"讀取 DB 失敗: {e}")
    return False

def save_record(table_name, key):
    """寫入單筆紀錄 (含重試)。"""
    max_retries = 10
    for attempt in range(max_retries):
        try:
            with sqlite3.connect(DB_FILE_PATH, timeout=10) as conn:
                cursor = conn.cursor()
                timestamp = time.time()
                cursor.execute(f'''
                    INSERT OR REPLACE INTO {table_name} (tmdb_id_str, timestamp)
                    VALUES (?, ?)
                ''', (key, timestamp))
                conn.commit()
                return
        except sqlite3.OperationalError as e:
            if "database is locked" in str(e) and attempt < max_retries - 1:
                time.sleep(0.1 * (attempt + 1))
                continue
            logging.error(f"寫入 DB 鎖定失敗 ({table_name}): {e}")
            break
        except Exception as e:
            logging.error(f"寫入 DB 錯誤 ({table_name}): {e}")
            break

def cleanup_and_optimize_db(vacuum=False):
    """
    清理過期資料。
    """
    tables_to_clean = {
        VIDEO_TO_PERSON_TABLE: VIDEO_TO_PERSON_RETENTION,
        PERSON_TO_VIDEO_TABLE: PERSON_TO_VIDEO_RETENTION,
        CHINESE_CHECK_TABLE: CHINESE_CHECK_RETENTION
    }

    try:
        with sqlite3.connect(DB_FILE_PATH, timeout=30) as conn:
            cursor = conn.cursor()
            total_deleted = 0
            
            for table_name, retention_days in tables_to_clean.items():
                cutoff_time = time.time() - (retention_days * 24 * 60 * 60)
                cursor.execute(f'DELETE FROM {table_name} WHERE timestamp < ?', (cutoff_time,))
                deleted_count = cursor.rowcount
                if deleted_count > 0:
                    logging.info(f"已從 '{table_name}' 刪除 {deleted_count} 筆過期紀錄。")
                    total_deleted += deleted_count
            
            conn.commit()

            # 只要 vacuum=True 請求壓縮，就執行 VACUUM
            if vacuum:
                logging.info("執行資料庫壓縮 (VACUUM)...")
                conn.execute('VACUUM')
                logging.info("資料庫壓縮完成。")
            
    except Exception as e:
        logging.error(f"清理資料庫錯誤: {e}")

# --- API 請求 Wrapper ---

async def make_tmdb_request(session, url, params, context=""):
    """
    統一的 API 請求處理函數。
    """
    attempt = 0
    while True:
        attempt += 1
        try:
            response = await session.get(url, params=params, timeout=20)
            
            if response.status_code == 404:
                return None # Not Found
            
            if response.status_code == 429:
                retry_after = response.headers.get("Retry-After")
                wait_time = int(retry_after) if retry_after and retry_after.isdigit() else min(2 ** attempt, 60)
                logging.warning(f"⚠️ Rate Limit (429): {context} - 等待 {wait_time} 秒...")
                await asyncio.sleep(wait_time)
                continue

            response.raise_for_status()
            return response.json()

        except (httpx.ConnectError, httpx.ReadTimeout, httpx.ConnectTimeout, httpx.NetworkError, httpx.RemoteProtocolError) as e:
            wait_time = min(2 ** attempt, 60)
            logging.warning(f"網路連線錯誤 ({context}): {e}。等待 {wait_time} 秒...")
            await asyncio.sleep(wait_time)
            
        except httpx.HTTPStatusError as e:
            status = e.response.status_code
            if 500 <= status < 600: # Server Error, retry
                wait_time = min(2 ** attempt, 60)
                logging.warning(f"伺服器錯誤 {status} ({context})。等待 {wait_time} 秒...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"API 請求失敗 {status} ({context})")
                return None # 4xx (非 404/429) 不重試

        except Exception as e:
            logging.error(f"未預期錯誤 ({context}): {e}")
            await asyncio.sleep(5)

# --- 核心邏輯函數 ---

async def get_movie_person_ids_api(session, semaphore, movie_id, api_key):
    person_ids = set()
    key = f"movie_{movie_id}"
    
    # 傳入 VIDEO_TO_PERSON_RETENTION (3天)
    if is_scanned(VIDEO_TO_PERSON_TABLE, key, VIDEO_TO_PERSON_RETENTION):
        return set()

    url = f"{TMDB_API_BASE_URL}/movie/{movie_id}/credits"
    params = {'api_key': api_key}
    
    async with semaphore:
        data = await make_tmdb_request(session, url, params, context=f"movie/{movie_id}/credits")
        
        save_record(VIDEO_TO_PERSON_TABLE, key)
        
        if data:
            for member in data.get('cast', []) + data.get('crew', []):
                if 'id' in member:
                    person_ids.add(member['id'])
            
    return person_ids

async def get_tv_show_person_ids_aggregate_api(session, semaphore, tv_id, api_key):
    person_ids = set()
    key = f"tv_{tv_id}"
    
    # 傳入 VIDEO_TO_PERSON_RETENTION (3天)
    if is_scanned(VIDEO_TO_PERSON_TABLE, key, VIDEO_TO_PERSON_RETENTION):
        return set()

    url = f"{TMDB_API_BASE_URL}/tv/{tv_id}/aggregate_credits"
    params = {'api_key': api_key}
    
    async with semaphore:
        data = await make_tmdb_request(session, url, params, context=f"tv/{tv_id}/agg_credits")
        save_record(VIDEO_TO_PERSON_TABLE, key)
        
        if data:
            for member in data.get('cast', []) + data.get('crew', []):
                if 'id' in member:
                    person_ids.add(member['id'])

    return person_ids

async def get_filmography_ids_api(session, semaphore, person_id, api_key, check_cache=True):
    """
    使用 API 獲取一位工作人員的所有作品 ID。
    """
    video_ids = set()
    key = str(person_id)
    
    # 傳入 PERSON_TO_VIDEO_RETENTION (7天)
    if check_cache and is_scanned(PERSON_TO_VIDEO_TABLE, key, PERSON_TO_VIDEO_RETENTION):
        return set()

    url = f"{TMDB_API_BASE_URL}/person/{person_id}/combined_credits"
    params = {'api_key': api_key}
    
    async with semaphore:
        data = await make_tmdb_request(session, url, params, context=f"person/{person_id}/credits")
        
        save_record(PERSON_TO_VIDEO_TABLE, key)
        
        if data:
            for item in data.get('cast', []) + data.get('crew', []):
                media_type = item.get('media_type')
                if media_type in ['movie', 'tv'] and 'id' in item:
                    video_ids.add((media_type, item['id']))
            
    return video_ids

# --- Step 4 兩段式優化 ---
async def check_chinese_content_optimized(session, semaphore, video_type, video_id, api_key, valid_set, valid_lock):
    """Step 4: 兩段式檢查"""
    key = f"{video_type}_{video_id}"
    context = f"{video_type}/{video_id}"
    
    async with semaphore:
        # 1. 檢查 Translations
        trans_url = f"{TMDB_API_BASE_URL}/{video_type}/{video_id}/translations"
        trans_data = await make_tmdb_request(session, trans_url, {'api_key': api_key}, context + "/trans")
        
        if trans_data is None: # 404
            save_record(CHINESE_CHECK_TABLE, key)
            return False

        for t in trans_data.get('translations', []):
            if t.get("iso_639_1") == "zh" and t.get("iso_3166_1") in ['HK', 'TW', 'CN']:
                async with valid_lock:
                    valid_set.add(f"https://www.themoviedb.org/{video_type}/{video_id}")
                save_record(CHINESE_CHECK_TABLE, key)
                return True # 命中官方翻譯，結束

        # 2. 檢查 Details (en-US title)
        details_url = f"{TMDB_API_BASE_URL}/{video_type}/{video_id}"
        details_data = await make_tmdb_request(session, details_url, {'api_key': api_key, 'language': 'en-US'}, context + "/details")
        
        if details_data:
            title = details_data.get('title') if video_type == 'movie' else details_data.get('name')
            orig_title = details_data.get('original_title') if video_type == 'movie' else details_data.get('original_name')
            
            if (title and contains_chinese(title)) or (orig_title and contains_chinese(orig_title)):
                async with valid_lock:
                    valid_set.add(f"https://www.themoviedb.org/{video_type}/{video_id}")
                save_record(CHINESE_CHECK_TABLE, key)
                return True

        save_record(CHINESE_CHECK_TABLE, key)
        return False

# --- 主程式 ---

async def main():
    init_db()

    print("請貼上或逐行輸入影片/演員 URL (可用逗號、空格或換行分隔)。")
    print("輸入完成後，請在新的一行直接按 Enter 鍵開始處理：")
    
    lines = []
    while True:
        try:
            line = input()
            if line == "": break
            lines.append(line)
        except EOFError: break
    
    urls_input = " ".join(lines)
    initial_urls = [url.strip() for url in re.split(r'[,，\s]+', urls_input) if url.strip()]

    if not initial_urls:
        logging.info("未輸入任何 URL。")
        cleanup_and_optimize_db(vacuum=True) # 結束時壓縮
        return

    configs = load_configs()
    api_key = get_tmdb_api_key(configs)
    if not api_key:
        return

    # [修改] 移除了開頭的 cleanup_and_optimize_db(vacuum=False)
    # 現在採用「惰性過期」檢查，讀取時才過濾，無需預先刪除。

    # 解析初始 URL
    initial_video_ids = set()
    initial_person_ids = set()

    for url in initial_urls:
        item_type, item_id = parse_tmdb_url(url)
        if item_type and item_id:
            if item_type in ['movie', 'tv']:
                initial_video_ids.add((item_type, item_id))
            elif item_type == 'person':
                initial_person_ids.add(int(item_id))
    
    # 建立初始輸入影片的標準化 URL 集合
    initial_input_urls_set = set()
    for t, i in initial_video_ids:
        initial_input_urls_set.add(f"https://www.themoviedb.org/{t}/{i}")

    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)
    
    valid_urls_with_chinese = set()
    valid_lock = asyncio.Lock()
    
    videos_from_persons = set()
    related_video_ids = set()
    all_person_ids = set()

    limits = httpx.Limits(max_keepalive_connections=20, max_connections=20)

    async with httpx.AsyncClient(limits=limits, timeout=20.0) as session:
        videos_for_expansion = set(initial_video_ids)

        # Step 1: 初始人員 -> 作品
        if initial_person_ids:
            logging.info(f"--- Step 1: 掃描 {len(initial_person_ids)} 位初始人員 ---")
            tasks = [get_filmography_ids_api(session, semaphore, p_id, api_key, check_cache=False) for p_id in initial_person_ids]
            
            total = len(tasks)
            done = 0
            start_t = time.time()
            
            for f in asyncio.as_completed(tasks):
                res = await f
                videos_from_persons.update(res)
                done += 1
                display_progress("Step 1", done, total, start_t)
            print()
            
            videos_for_expansion.update(videos_from_persons)
            logging.info(f"從人員找到 {len(videos_from_persons)} 部作品。")

        # Step 2: 影片 -> 更多人員
        if videos_for_expansion:
            logging.info(f"--- Step 2: 從 {len(videos_for_expansion)} 部影片掃描工作人員 ---")
            
            video_keys = [f"{t}_{i}" for t, i in videos_for_expansion]
            # [修改] 傳入 VIDEO_TO_PERSON_RETENTION (3天)
            scanned_videos = get_scanned_keys_batch(VIDEO_TO_PERSON_TABLE, video_keys, VIDEO_TO_PERSON_RETENTION)
            videos_to_scan = []
            for t, i in videos_for_expansion:
                if f"{t}_{i}" not in scanned_videos:
                    videos_to_scan.append((t, i))
            
            logging.info(f"快取命中：跳過 {len(scanned_videos)} 部，實際需掃描 {len(videos_to_scan)} 部。")

            tasks = []
            for v_type, v_id in videos_to_scan:
                if v_type == 'movie':
                    tasks.append(get_movie_person_ids_api(session, semaphore, v_id, api_key))
                else:
                    tasks.append(get_tv_show_person_ids_aggregate_api(session, semaphore, v_id, api_key))
            
            if tasks:
                total = len(tasks)
                done = 0
                start_t = time.time()
                for f in asyncio.as_completed(tasks):
                    res = await f
                    all_person_ids.update(res)
                    done += 1
                    display_progress("Step 2", done, total, start_t)
                print()
            else:
                logging.info("所有影片的工作人員資料皆已快取。")

            all_person_ids.update(initial_person_ids)
            logging.info(f"總共收集到 {len(all_person_ids)} 位工作人員 ID。")

        # Step 3: 所有人員 -> 所有作品
        if all_person_ids:
            logging.info(f"--- Step 3: 掃描 {len(all_person_ids)} 位人員的作品 ---")
            
            person_keys = [str(pid) for pid in all_person_ids]
            # [修改] 傳入 PERSON_TO_VIDEO_RETENTION (7天)
            scanned_persons = get_scanned_keys_batch(PERSON_TO_VIDEO_TABLE, person_keys, PERSON_TO_VIDEO_RETENTION)
            persons_to_scan = [pid for pid in all_person_ids if str(pid) not in scanned_persons]
            
            logging.info(f"快取命中：跳過 {len(scanned_persons)} 位，實際需掃描 {len(persons_to_scan)} 位。")
            
            tasks = [get_filmography_ids_api(session, semaphore, p_id, api_key) for p_id in persons_to_scan]
            if tasks:
                total = len(tasks)
                done = 0
                start_t = time.time()
                for f in asyncio.as_completed(tasks):
                    res = await f
                    related_video_ids.update(res)
                    done += 1
                    display_progress("Step 3", done, total, start_t)
                print()
            else:
                logging.info("所有人員皆已快取。")

        # Step 4: 最終檢查
        all_videos = initial_video_ids.union(videos_from_persons).union(related_video_ids)
        logging.info(f"--- Step 4: 檢查 {len(all_videos)} 部影片的中文內容 ---")

        video_keys = [f"{t}_{i}" for t, i in all_videos]
        # [修改] 傳入 CHINESE_CHECK_RETENTION (365天)
        scanned_videos = get_scanned_keys_batch(CHINESE_CHECK_TABLE, video_keys, CHINESE_CHECK_RETENTION)
        
        videos_to_check = []
        for t, i in all_videos:
            key = f"{t}_{i}"
            if (t, i) in initial_video_ids:
                videos_to_check.append((t, i))
                if key in scanned_videos:
                    save_record(CHINESE_CHECK_TABLE, key)
            elif key not in scanned_videos:
                videos_to_check.append((t, i))
        
        logging.info(f"快取命中：跳過 {len(all_videos) - len(videos_to_check)} 部，實際需 API 檢查 {len(videos_to_check)} 部。")

        if videos_to_check:
            tasks = [
                check_chinese_content_optimized(session, semaphore, t, i, api_key, valid_urls_with_chinese, valid_lock)
                for t, i in videos_to_check
            ]
            
            total = len(tasks)
            done = 0
            start_t = time.time()
            
            for f in asyncio.as_completed(tasks):
                await f
                done += 1
                display_progress("Step 4", done, total, start_t, f"找到: {len(valid_urls_with_chinese)}")
            print()
        else:
            logging.info("所有影片皆已快取，無需 API 檢查。")

    # 寫入 txt（唔直接改 bulk_videos.xlsx）
    if valid_urls_with_chinese:
        write_urls_txt(valid_urls_with_chinese, 'bulk_video_person')

    # 程式結束前，執行清理並強制壓縮 (vacuum=True)
    # 這一步現在負責刪除所有過期資料
    cleanup_and_optimize_db(vacuum=True)
    logging.info("執行完畢。")

def normalize_tmdb_url(url):
    if not isinstance(url, str): return None
    match = re.search(r'(?:https?://)?(?:www\.)?themoviedb\.org/(movie|tv)/(\d+)', url)
    if match:
        return f"https://www.themoviedb.org/{match.group(1)}/{match.group(2)}"
    return None

if __name__ == '__main__':
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())