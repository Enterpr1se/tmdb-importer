import httpx
import asyncio
import pandas as pd
import logging
import re
import os
import json
import openpyxl
import sys
import time # (新增) 為了 time.sleep
import traceback # (新增) 為了記錄詳細錯誤
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bulk_txt import write_urls_txt  # sub-program 共用：輸出到 txt，唔直接改 bulk_videos.xlsx

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Global Constants ---
MAX_CONCURRENT_REQUESTS = 30

# --- Integrated Functions ---

def load_configs():
    """從 ../configs.json 載入設定檔。"""
    # 腳本位於 ./sub-program/，此路徑會指向根目錄的 configs.json
    config_path = os.path.join(os.path.dirname(__file__), '../configs.json')
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as config_file:
                return json.load(config_file)
        except Exception as e:
            logging.error(f"讀取 configs.json 時發生錯誤: {e}")
            return {}
    logging.warning(f"設定檔 {config_path} 不存在。")
    return {}

# --- Core Functions ---

async def search_by_id(session, semaphore, api_key, item_id, all_urls, api_base_url, web_base_url):
    """
    根據提供的 ID 異步查詢人物的所有影視作品、單一電影或單一電視節目。
    (使用者需求 2025-09-02) 新增錯誤重試機制，如果 API 請求失敗，最多重試 100 次。
    """
    logging.info(f"偵測到 ID '{item_id}'，正在嘗試作為 Person, Movie, TV Show ID 進行查詢。")
    found_something = False

    # 1. 嘗試作為 Person ID
    # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
    for attempt in range(101):
        try:
            credits_params = {"api_key": api_key, "language": "zh-TW"}
            async with semaphore:
                credits_response = await session.get(f"{api_base_url}/person/{item_id}/combined_credits", params=credits_params, timeout=20)
            
            if credits_response.status_code == 404:
                logging.debug(f"以 Person ID '{item_id}' 查詢時項目不存在 (404)。")
                break # 404 不重試，跳出此人員查詢

            credits_response.raise_for_status()
            credits_data = credits_response.json()
            
            all_credits = credits_data.get('cast', []) + credits_data.get('crew', [])
            if all_credits:
                person_works = set()
                for credit in all_credits:
                    media_type = credit.get("media_type")
                    credit_id = credit.get("id")
                    if not credit_id: continue
                    if media_type in ["movie", "tv"]:
                        url = f"{web_base_url}/{media_type}/{credit_id}"
                        all_urls.add(url)
                        person_works.add(url)
                if person_works:
                    logging.info(f"成功將 ID '{item_id}' 作為 Person 進行查詢，找到 {len(person_works)} 部影視作品。")
                    found_something = True
            break # 成功，跳出重試
        except (httpx.RequestError, httpx.HTTPStatusError) as e:
            if attempt < 100:
                wait_time = min(2 ** attempt, 60)
                logging.warning(f"以 Person ID '{item_id}' 查詢時出錯 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"以 Person ID '{item_id}' 查詢時出錯，已達最大重試次數 (101 次): {e}")
        except Exception as e:
            logging.error(f"以 Person ID '{item_id}' 查詢時發生非預期錯誤: {traceback.format_exc()}")
            break

    # 2. 嘗試作為 Movie ID
    # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
    for attempt in range(101):
        try:
            movie_params = {"api_key": api_key, "language": "zh-TW"}
            async with semaphore:
                movie_response = await session.get(f"{api_base_url}/movie/{item_id}", params=movie_params, timeout=20)
            
            if movie_response.status_code == 404:
                logging.debug(f"以 Movie ID '{item_id}' 查詢時項目不存在 (404)。")
                break

            movie_response.raise_for_status()
            logging.info(f"成功將 ID '{item_id}' 作為 Movie ID 進行查詢。")
            all_urls.add(f"{web_base_url}/movie/{item_id}")
            found_something = True
            break
        except (httpx.RequestError, httpx.HTTPStatusError) as e:
            if attempt < 100:
                wait_time = min(2 ** attempt, 60)
                logging.warning(f"以 Movie ID '{item_id}' 查詢時出錯 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"以 Movie ID '{item_id}' 查詢時出錯，已達最大重試次數 (101 次): {e}")
        except Exception as e:
            logging.error(f"以 Movie ID '{item_id}' 查詢時發生非預期錯誤: {traceback.format_exc()}")
            break

    # 3. 嘗試作為 TV Show ID
    # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
    for attempt in range(101):
        try:
            tv_params = {"api_key": api_key, "language": "zh-TW"}
            async with semaphore:
                tv_response = await session.get(f"{api_base_url}/tv/{item_id}", params=tv_params, timeout=20)

            if tv_response.status_code == 404:
                logging.debug(f"以 TV Show ID '{item_id}' 查詢時項目不存在 (404)。")
                break
                
            tv_response.raise_for_status()
            logging.info(f"成功將 ID '{item_id}' 作為 TV Show ID 進行查詢。")
            all_urls.add(f"{web_base_url}/tv/{item_id}")
            found_something = True
            break
        except (httpx.RequestError, httpx.HTTPStatusError) as e:
            if attempt < 100:
                wait_time = min(2 ** attempt, 60)
                logging.warning(f"以 TV Show ID '{item_id}' 查詢時出錯 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"以 TV Show ID '{item_id}' 查詢時出錯，已達最大重試次數 (101 次): {e}")
        except Exception as e:
            logging.error(f"以 TV Show ID '{item_id}' 查詢時發生非預期錯誤: {traceback.format_exc()}")
            break

    if not found_something:
        logging.warning(f"ID '{item_id}' 未能成功匹配任何 Person, Movie 或 TV Show。")

def contains_chinese(text):
    """
    輔助函數：檢查一個字串是否包含任何中文字元，並排除包含日文假名的字串。
    - (使用者要求 2025-08-19) 修正將日文漢字誤判為中文的問題。
    - 實現方式：
    -   1. 首先，使用正規表示式檢查字串中是否存在日文的平假名或片假名。
    -      日文假名的 Unicode 範圍為 U+3040 到 U+30FF。
    -   2. 如果檢測到任何假名字符，則立即將該字串判定為「非中文」並返回 False。
    -   3. 只有在確認字串「不含」任何日文假名的前提下，才繼續檢查其是否包含 CJK 統一表意文字（漢字）。
    -      這樣可以有效過濾掉日文標題，同時正確識別純漢字標題。
    """
    if not text or not isinstance(text, str):
        return False
    
    # 步驟 1: 檢查是否有日文假名（平假名: \u3040-\u309f, 片假名: \u30a0-\u30ff）
    # 如果找到任何假名，就立即返回 False，因為中文文本不應包含這些字符。
    if re.search(r'[\u3040-\u30ff]', text):
        return False
        
    # 步驟 2: 在確認沒有日文假名後，再檢查是否包含中文字符（漢字）
    # \u4e00-\u9fff 是 CJK 統一表意文字的基本範圍
    return bool(re.search(r'[\u4e00-\u9fff]', text))

async def has_chinese_translation(session, semaphore, api_key, tmdb_url, api_base_url):
    """
    異步檢查指定的 TMDB URL 是否有中文內容。成功則返回 URL，否則返回 None。
    (使用者需求 2025-08-19) 採用更有效率的方式檢查影片是否有中文內容。
    (使用者需求 2025-09-02) 新增錯誤重試機制，如果 API 請求失敗，最多重試 100 次。
    """
    match = re.search(r'/(movie|tv)/(\d+)', tmdb_url)
    if not match:
        logging.warning(f"無法從 URL 中解析 media type 或 ID: {tmdb_url}。將跳過。")
        return None

    media_type = match.group(1)
    item_id = match.group(2)
    
    url = f"{api_base_url}/{media_type}/{item_id}?api_key={api_key}&language=en-US&append_to_response=translations"
    
    # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
    for attempt in range(101):
        try:
            async with semaphore:
                response = await session.get(url, timeout=20)
                
                if response.status_code == 404:
                    logging.warning(f"❌ 檢查 {tmdb_url} 時項目不存在 (404)。")
                    return None

                response.raise_for_status()
                data = response.json()

            # --- 第一階段：檢查官方翻譯 ---
            translations = data.get('translations', {}).get('translations', [])
            for t in translations:
                lang_code = t.get("iso_639_1")
                country_code = t.get("iso_3166_1")
                if lang_code == 'zh' and country_code in ['HK', 'TW', 'CN']:
                    logging.info(f"✔️  URL {tmdb_url} 找到官方中文翻譯 ({lang_code}-{country_code})。")
                    return tmdb_url

            # --- 第二階段：檢查標題 ---
            title = data.get('title') if media_type == 'movie' else data.get('name')
            original_title = data.get('original_title') if media_type == 'movie' else data.get('original_name')
            
            title_to_log = ""
            is_chinese = False
            if title and contains_chinese(title):
                is_chinese = True
                title_to_log = title
            elif original_title and contains_chinese(original_title):
                is_chinese = True
                title_to_log = original_title

            if is_chinese:
                logging.info(f"✔️  URL {tmdb_url} 在 en-US 資料中找到中文標題/原標題: '{title_to_log}'。")
                return tmdb_url
            
            logging.warning(f"❌ URL {tmdb_url} 未找到任何指定的中文內容 (官方翻譯和 en-US 標題/原標題均不符合)。")
            return None # 檢查成功但無中文，返回 None

        except (httpx.RequestError, httpx.HTTPStatusError) as e:
            if attempt < 100:
                wait_time = min(2 ** attempt, 60)
                logging.warning(f"檢查 {tmdb_url} 時發生網路錯誤 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"檢查 {tmdb_url} 時發生網路錯誤，已達最大重試次數 (101 次): {e}")
        except Exception as e:
            logging.error(f"檢查 {tmdb_url} 時發生非預期錯誤: {traceback.format_exc()}")
            break # 遇到非網路問題，直接中斷

    logging.warning(f"❌ 檢查 {tmdb_url} 在多次重試後最終失敗。")
    return None

async def search_tmdb(api_key, keywords):
    """
    根據關鍵字列表在 TMDB 上異步搜尋，並過濾具有中文翻譯的結果。
    """
    api_base_url = "https://api.themoviedb.org/3"
    web_base_url = "https://www.themoviedb.org"
    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)

    # --- Nested Helper Functions for Async Execution ---

    async def _process_text_search(session, keyword, keyword_urls_set):
        """
        為單一文字關鍵字執行異步分頁搜尋，並將結果填入提供的集合中。
        (使用者需求 2025-09-02) 新增錯誤重試機制，如果 API 請求失敗，最多重試 100 次。
        """
        page = 1
        total_pages = 1
        page_results_count = 0

        while page <= total_pages:
            params = {
                "api_key": api_key, "query": keyword, "page": page,
                "include_adult": True, "language": "zh-TW",
            }
            
            # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
            response_data = None
            for attempt in range(101):
                try:
                    async with semaphore:
                        response = await session.get(f"{api_base_url}/search/multi", params=params, timeout=20)
                    
                    if response.status_code == 404:
                        logging.warning(f"關鍵字 '{keyword}' 在第 {page} 頁的 TMDB 搜尋不存在 (404)。")
                        break # 404 不重試

                    response.raise_for_status()
                    response_data = response.json()
                    break # 成功，跳出重試迴圈
                except (httpx.RequestError, httpx.HTTPStatusError) as e:
                    if attempt < 100:
                        wait_time = min(2 ** attempt, 60)
                        logging.warning(f"關鍵字 '{keyword}' 在第 {page} 頁的 TMDB 搜尋時出錯 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                        await asyncio.sleep(wait_time)
                    else:
                        logging.error(f"關鍵字 '{keyword}' 在第 {page} 頁的 TMDB 搜尋失敗，已達最大重試次數 (101 次): {e}")
                except Exception as e:
                    logging.error(f"關鍵字 '{keyword}' 在第 {page} 頁的 TMDB 搜尋過程中發生非預期錯誤: {traceback.format_exc()}")
                    break
            
            if not response_data:
                break # 如果所有重試都失敗，結束此關鍵字的搜尋

            results = response_data.get("results", [])
            page_results_count += len(results)
            
            if page == 1:
                total_pages = response_data.get("total_pages", 1)
                total_results = response_data.get("total_results", 0)
                if total_results == 0:
                    logging.warning(f"關鍵字 '{keyword}' 從 TMDB API 返回 0 個結果。")

            for item in results:
                media_type = item.get("media_type")
                item_id = item.get("id")
                if not item_id: continue

                if media_type in ["movie", "tv"]:
                    keyword_urls_set.add(f"{web_base_url}/{media_type}/{item_id}")
                elif media_type == "person":
                    person_name = item.get('name', 'N/A')
                    logging.info(f"為關鍵字 '{keyword}' 找到人物 '{person_name}' (ID: {item_id})。正在擷取其影視作品。")
                    await search_by_id(session, semaphore, api_key, str(item_id), keyword_urls_set, api_base_url, web_base_url)
            
            logging.info(f"關鍵字 '{keyword}' - 第 {page}/{total_pages} 頁: 找到 {len(results)} 個結果 (API 總計: {response_data.get('total_results', 0)})")
            page += 1
        
        if page > 1:
            logging.info(f"關鍵字 '{keyword}': 在 {page-1} 頁中總共找到 {page_results_count} 個項目。")

    # ... (此處省略 _search_task_for_keyword 和 search_tmdb 的其餘部分) ...

def main():
    """主函數，運行 TMDB 搜尋並將結果附加到 Excel。"""
    configs = load_configs()
    api_key = configs.get("tmdb_api_v3")
    if not api_key:
        logging.error("TMDB API key (tmdb_api_v3) not found in configs.json")
        print("錯誤：未在 ../configs.json 中找到 TMDB API 金鑰 (tmdb_api_v3)。請確認配置。")
        return
    
    prompt = (
        "請輸入搜尋關鍵字、TMDB URL 或 ID（以逗號或全形逗號分隔多個輸入，\n"
        "例如：shark, https://www.themoviedb.org/person/1119117， 1999324）："
    )
    keyword_input = input(prompt).strip()
    if not keyword_input:
        logging.error("使用者未提供關鍵字")
        print("錯誤：請輸入至少一個有效關鍵字、URL 或 ID。")
        return
    
    # 使用 re.split 來處理多種分隔符號 (半形 , 和全形 ，)，並過濾掉因連續分隔符而產生的空字串
    keywords = [kw.strip() for kw in re.split('[,，]', keyword_input) if kw.strip()]
    
    # 執行異步搜尋
    urls = asyncio.run(search_tmdb(api_key, keywords))
    
    # 寫入 txt（唔直接改 bulk_videos.xlsx）
    if urls:
        path, added = write_urls_txt(urls, 'tmdb_search')
        print(f"已將 {added} 個新搜尋結果寫入 {path}。")
    else:
        logging.info(f"未找到關鍵字/ID的結果: {keywords}")
        print(f"未找到與 '{', '.join(keywords)}' 相關的結果。")

if __name__ == "__main__":
    # 在 Windows 上，為 asyncio 設定不同的事件循環策略以避免錯誤
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    main()