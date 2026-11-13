import os
import re
import sys
import time
import logging
import pandas as pd
import openpyxl
import traceback
import json
import asyncio
import httpx
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bulk_txt import write_urls_txt  # sub-program 共用：輸出到 txt，唔直接改 bulk_videos.xlsx

# --- 設定日誌 ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- 路徑設定 ---
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)
TMDB_API_BASE_URL = "https://api.themoviedb.org/3"

# --- 設定檔與 API Key 管理 ---
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

def get_tmdb_api_key():
    """從設定檔或使用者輸入獲取 TMDB API 金鑰。"""
    configs = load_configs()
    if 'tmdb_api_v3' not in configs or not configs['tmdb_api_v3']:
        print("未檢測到 API Key。")
        api_key = input("請輸入您的 TMDB API Key (v3): ").strip()
        if not api_key:
            return None
        save_config = input("是否希望儲存此 API Key 留待下次使用? (y/n): ")
        if save_config.lower() == 'y':
            configs['tmdb_api_v3'] = api_key
            save_configs(configs)
    else:
        api_key = configs['tmdb_api_v3']
    return api_key

# --- 輔助功能 ---
def contains_chinese(text):
    """檢查字串是否包含中文（排除日文假名）。"""
    if not text or not isinstance(text, str):
        return False
    # 檢查是否有日文假名，若有則視為非中文
    if re.search(r'[\u3040-\u30ff]', text):
        return False
    # 檢查是否包含 CJK 統一表意文字
    return bool(re.search(r'[\u4e00-\u9fff]', text))

def extract_tmdb_info(url):
    """從 URL 解析出類型和 ID，用於 API 呼叫。"""
    # 這裡的正則表達式限制了只能抓取 movie 或 tv
    match = re.search(r'themoviedb\.org/(movie|tv)/(\d+)', url)
    if match:
        return match.group(1), match.group(2)
    return None, None

def normalize_tmdb_url(url):
    """
    標準化 TMDB URL。
    輸入: https://www.themoviedb.org/tv/9297-brigadoon
    輸出: https://www.themoviedb.org/tv/9297
    注意：此函數透過正則表達式 (movie|tv) 僅接受電影和影集，
    如果傳入 person 類型的 URL 會回傳 None。
    """
    if not isinstance(url, str):
        return None
    match = re.search(r'(?:https?://)?(?:www\.)?themoviedb\.org/(movie|tv)/(\d+)', url)
    if match:
        return f"https://www.themoviedb.org/{match.group(1)}/{match.group(2)}"
    return None

# --- API 檢查邏輯 ---
async def check_chinese_content(client, semaphore, url, api_key):
    """
    檢查單一 URL 是否有中文內容。
    回傳: (URL, Boolean) - True 表示有中文
    具有無限重試機制，直到成功或遇到 404/程式錯誤。
    """
    video_type, video_id = extract_tmdb_info(url)
    if not video_type or not video_id:
        return url, False

    api_url = f"{TMDB_API_BASE_URL}/{video_type}/{video_id}"
    params = {'api_key': api_key, 'language': 'en-US', 'append_to_response': 'translations'}
    
    attempt = 0
    
    async with semaphore:
        while True: # 無限迴圈，直到 return 或 break
            attempt += 1
            try:
                response = await client.get(api_url, params=params, timeout=20)
                
                # 404 是致命錯誤，代表 ID 不存在，不需要重試
                if response.status_code == 404:
                    logging.warning(f"❌ 項目不存在 (404): {url}")
                    return url, False
                
                response.raise_for_status() # 檢查其他 HTTP 錯誤 (如 500, 502, 429 等)
                data = response.json()
                
                # 1. 檢查官方翻譯 (zh-CN, zh-HK, zh-TW)
                translations = data.get('translations', {}).get('translations', [])
                for t in translations:
                    lang = t.get("iso_639_1")
                    country = t.get("iso_3166_1")
                    if lang == 'zh' and country in ['HK', 'TW', 'CN']:
                        return url, True

                # 2. 檢查 en-US 標題或原標題是否包含中文字
                title = data.get('title') if video_type == 'movie' else data.get('name')
                original_title = data.get('original_title') if video_type == 'movie' else data.get('original_name')
                
                if (title and contains_chinese(title)) or (original_title and contains_chinese(original_title)):
                    return url, True

                # 若都不符合，正常返回 False
                return url, False

            except (httpx.RequestError, httpx.HTTPStatusError) as e:
                # 遇到網路錯誤或 HTTP 狀態錯誤 (如 429 Too Many Requests, 503 Service Unavailable)
                # 使用指數退避策略計算等待時間，但設定上限為 60 秒
                wait_time = min(1.5 ** attempt, 60)
                logging.warning(f"檢查 {url} 時發生錯誤 (第 {attempt} 次): {e}。將在 {wait_time:.1f} 秒後重試...")
                await asyncio.sleep(wait_time)
                
            except Exception as e:
                # 未預期的程式錯誤 (如解析 JSON 失敗、變數型別錯誤等)，不重試直接跳出
                logging.error(f"檢查 {url} 時發生未預期錯誤: {e}")
                logging.error(traceback.format_exc())
                return url, False
    
    return url, False

async def main():
    # 1. 獲取 API Key
    api_key = get_tmdb_api_key()
    if not api_key:
        logging.error("未提供 API Key，無法進行中文檢查，程式結束。")
        return

    # 2. 獲取輸入
    print("請貼上或逐行輸入影片 URL (可用逗號、空格或換行分隔)。")
    print("程式僅處理 /movie/ 或 /tv/ 網址，並檢查是否有中文內容。")
    print("輸入完成後，請在新的一行直接按 Enter 鍵開始處理：")
    
    lines = []
    while True:
        try:
            line = input()
            if line == "":
                break
            lines.append(line)
        except EOFError:
            break
    
    urls_input = " ".join(lines)
    input_urls = [url.strip() for url in re.split(r'[,，\s]+', urls_input) if url.strip()]

    if not input_urls:
        logging.info("未輸入任何 URL，程式結束。")
        return

    # 3. 標準化 URL 並準備檢查
    unique_check_urls = set()
    for url in input_urls:
        norm_url = normalize_tmdb_url(url)
        if norm_url:
            unique_check_urls.add(norm_url)
        else:
            if 'themoviedb.org' in url:
                if '/person/' in url:
                    logging.info(f"忽略 Person URL (僅支援 TV/Movie): {url}")
                elif '/collection/' in url:
                     logging.info(f"忽略 Collection URL (僅支援 TV/Movie): {url}")
                elif '/network/' in url:
                     logging.info(f"忽略 Network URL (僅支援 TV/Movie): {url}")
                else:
                    logging.warning(f"忽略不支援的類型或格式錯誤 (僅支援 TV/Movie): {url}")

    if not unique_check_urls:
        logging.info("沒有有效的 TMDB Movie 或 TV URL 需要處理。")
        return

    logging.info(f"正在檢查 {len(unique_check_urls)} 個 URL 是否包含中文內容...")

    # 4. 非同步檢查中文內容
    semaphore = asyncio.Semaphore(20) # 限制同時連線數
    valid_urls_with_chinese = []
    
    async with httpx.AsyncClient() as client:
        tasks = [check_chinese_content(client, semaphore, url, api_key) for url in unique_check_urls]
        
        # 使用進度顯示
        total = len(tasks)
        completed = 0
        start_time = time.time() # 初始化開始時間

        for future in asyncio.as_completed(tasks):
            url, has_chinese = await future
            completed += 1
            
            # --- 計算剩餘時間 (ETC) ---
            elapsed_time = time.time() - start_time
            items_per_second = completed / elapsed_time if elapsed_time > 0 else 0
            remaining_items = total - completed
            
            if items_per_second > 0:
                etc_seconds = remaining_items / items_per_second
                etc_str = f"預計剩餘 {int(etc_seconds // 60)} 分 {int(etc_seconds % 60)} 秒"
            else:
                etc_str = "計算中..."

            if has_chinese:
                valid_urls_with_chinese.append(url)
                msg = f"進度: {completed}/{total} - ✔️  {url} (含中文) - {etc_str}"
            else:
                msg = f"進度: {completed}/{total} - ❌ {url} (無中文) - {etc_str}"
            
            # 使用 ljust 清除行尾殘留的文字
            sys.stdout.write(f"\r{msg}".ljust(110))
            sys.stdout.flush()
            
        print() # 換行

    if not valid_urls_with_chinese:
        logging.info("檢查完畢，所有輸入的影片均無中文內容，冇嘢要輸出。")
        return

    logging.info(f"檢查完畢，共 {len(valid_urls_with_chinese)} 個影片包含中文。")
    write_urls_txt(valid_urls_with_chinese, 'manage_video_urls')

if __name__ == '__main__':
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())