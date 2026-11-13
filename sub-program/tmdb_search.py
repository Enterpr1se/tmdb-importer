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

def check_create_bulk_excel(file_name='../data/bulk_videos.xlsx'):
    """檢查並創建 bulk_videos.xlsx，確保包含 Bulk 工作表及必要欄位。"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    required_columns = ['Video URL', 'TMDB URL', 'Split', 'Done', '覆蓋劇集標題', '覆蓋劇集描述', '覆蓋分集標題', '覆蓋分集描述', '來源無資料時清空TMDB', '維基人名取代現有']
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            if not os.path.exists(file_name):
                # 如果檔案不存在，創建一個新的
                pd.DataFrame(columns=required_columns).to_excel(file_name, sheet_name='Bulk', index=False, engine='openpyxl')
                logging.info(f"已創建新的 {file_name}，包含 Bulk 工作表。")
            else:
                # 如果檔案存在，檢查工作表和欄位
                wb = openpyxl.load_workbook(file_name)
                if 'Bulk' not in wb.sheetnames:
                    # 如果工作表不存在，創建它
                    wb.create_sheet('Bulk')
                    ws = wb['Bulk']
                    ws.append(required_columns)
                    logging.info(f"工作表 Bulk 不存在於 {file_name}，已創建。")
                    wb.save(file_name)
                else:
                    # 如果工作表存在，檢查欄位
                    ws = wb['Bulk']
                    header = [cell.value for cell in ws[1]]
                    missing_cols = [col for col in required_columns if col not in header]
                    if missing_cols:
                        for col in missing_cols:
                            ws.cell(row=1, column=len(header) + 1 + missing_cols.index(col), value=col)
                        logging.info(f"已為 {file_name} 的 Bulk 工作表添加缺失的欄位: {missing_cols}")
                        wb.save(file_name)
            return # 成功後退出
        except PermissionError:
            if attempt < max_retries - 1:
                logging.warning(f"無法訪問 {file_name}，檔案可能正被其他程式使用。重試 {attempt + 1}/{max_retries}")
                input(f"請關閉 {file_name}，然後按 Enter 鍵繼續...")
            else:
                logging.error(f"超過最大重試次數，無法訪問 {file_name}。")
                raise
        except Exception as e:
            logging.error(f"創建或檢查 {file_name} 時發生錯誤: {e}")
            raise

def append_to_bulk_excel(urls, file_name='../data/bulk_videos.xlsx', sheet_name='Bulk'):
    """
    將 TMDB URL 添加到 bulk_videos.xlsx 的 Bulk 工作表中。
    此函數會保留現有數據的順序，只將不存在於檔案中的新 URL 追加到最末端。
    (使用者需求 2025-08-19) 新增檔案鎖定機制，避免與 bulk_video_person.py 等其他程式同時寫入檔案造成損毀。
    (使用者需求 2025-08-19) 升級鎖定機制，使其與 tmdb_importer.py 同步，
    增加一個內部重試迴圈來處理來自外部程式（如 Excel）的檔案鎖定（PermissionError）。
    """
    check_create_bulk_excel(file_name)

    # --- 檔案鎖定機制 ---
    # 為了避免此程式 (tmdb_search.py) 和其他程式 (如 bulk_video_person.py) 同時寫入 Excel 檔案，
    # 引入一個 .lock 檔案作為跨行程的鎖。
    # 只有成功創建 .lock 檔案的程式實例才能繼續執行寫入操作。
    lock_file = file_name + ".lock"
    output_filename = os.path.basename(file_name)
    
    logging.info(f"準備寫入 '{output_filename}'，開始檢查檔案鎖...")

    # --- 階段一：獲取腳本鎖 (.lock 檔案)，防止其他 Python 腳本同時寫入 ---
    lock_acquired = False
    script_wait_message_logged = False
    while True:
        try:
            # 嘗試以獨佔模式創建鎖定檔
            lock_fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(lock_fd)
            lock_acquired = True
            if script_wait_message_logged:
                 logging.info(f"檢測到 '{output_filename}' 已可用，現在開始寫入資料")
            logging.info(f"成功獲取 '{output_filename}' 的腳本寫入權限。")
            break # 成功獲取鎖，跳出迴圈
        except FileExistsError:
            if not script_wait_message_logged:
                logging.info(f"'{output_filename}' 正被另一個程式實例處理中，現在等待其完成...")
                script_wait_message_logged = True
            time.sleep(1) # 等待 1 秒後重試
        except Exception as e:
            logging.error(f"獲取檔案鎖時發生未知錯誤: {e}")
            return # 發生其他錯誤時，放棄

    # --- 階段二：在持有腳本鎖的情況下，進行寫入並處理外部程式鎖（如 Excel） ---
    if lock_acquired:
        # 建立一個外層的 try...finally 結構，
        # 目的是確保無論內部重試邏輯成功或失敗，最終都一定會執行 finally 區塊來釋放檔案鎖。
        try:
            # 新增 Excel 寫入重試機制。
            # 當寫入失敗時 (例如檔案暫時被防毒軟體鎖定或發生 I/O 錯誤)，會自動重試。
            write_successful = False
            permission_error_logged = False
            while not write_successful:
                try:
                    # 只有在第一次嘗試，且之前等待過鎖時，才顯示此訊息
                    if script_wait_message_logged and not permission_error_logged:
                        logging.info(f"檢測到 '{output_filename}' 已可用，現在開始寫入資料")

                    # --- 原有的寫入邏輯 ---
                    # 步驟 1: 讀取現有數據
                    try:
                        existing_df = pd.read_excel(file_name, sheet_name=sheet_name, engine='openpyxl')
                    except (FileNotFoundError, ValueError): # 處理檔案或工作表不存在的情況
                        existing_df = pd.DataFrame(columns=['Video URL', 'TMDB URL', 'Split', 'Done', '覆蓋劇集標題', '覆蓋劇集描述', '覆蓋分集標題', '覆蓋分集描述', '來源無資料時清空TMDB', '維基人名取代現有'])

                    # 步驟 2: 準備並篩選出真正新的 URL
                    new_urls_df = pd.DataFrame(urls, columns=['Video URL'])
                    
                    if not existing_df.empty and 'Video URL' in existing_df.columns:
                        existing_urls = set(existing_df['Video URL'].dropna())
                        new_unique_df = new_urls_df[~new_urls_df['Video URL'].isin(existing_urls)]
                    else:
                        new_unique_df = new_urls_df

                    # 步驟 3: 如果有新數據，則追加並保存
                    if not new_unique_df.empty:
                        final_df = pd.concat([existing_df, new_unique_df], ignore_index=True)
                        
                        # 確保新加入的行在所有欄位都有預設值
                        for col in ['Done', 'TMDB URL', 'Split']:
                            if col not in final_df.columns:
                                final_df[col] = '' if col != 'Done' else 0
                            else:
                                fill_value = '' if col != 'Done' else 0
                                final_df[col] = final_df[col].fillna(fill_value)
                        
                        logging.info(f"發現 {len(new_unique_df)} 個新的 URL，將其追加到 {file_name} 中。")
                        
                        # 將最終的 DataFrame 寫回 Excel
                        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            final_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        logging.info(f"成功將新數據寫入 Excel。")
                    else:
                        logging.info("本次運行未發現新的 URL，Excel 檔案內容未變動。")
                    
                    # 成功後退出重試迴圈
                    write_successful = True
                
                except PermissionError:
                    # 如果是權限錯誤，表示檔案被 Excel 等外部程式鎖定
                    if not permission_error_logged:
                        logging.warning(f"'{output_filename}' 正被外部程式(如 Excel)使用中，請關閉後程式將自動重試...")
                        permission_error_logged = True
                    time.sleep(1) # 等待 1 秒後重試
                except Exception as e:
                    # 如果是其他錯誤，記錄詳細訊息並終止，因為可能不是暫時性問題
                    logging.error(f"在持有鎖的情況下寫入 Excel 時發生非預期的錯誤:")
                    logging.error(traceback.format_exc())
                    break # 跳出重試迴圈，放棄寫入
        finally:
            # (極其重要) 無論 try 區塊是否成功，都必須確保鎖定檔案被刪除
            try:
                os.remove(lock_file)
                logging.info(f"已釋放 '{output_filename}' 的寫入權限。")
            except OSError as e:
                logging.error(f"釋放檔案鎖 '{lock_file}' 時失敗: {e}")

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
    
    # 定義輸出檔案的路徑（相對於父目錄）
    output_file = '../data/bulk_videos.xlsx'
    if urls:
        append_to_bulk_excel(urls, file_name=output_file)
        # 顯示絕對路徑，讓使用者清楚知道檔案位置
        absolute_path = os.path.abspath(output_file)
        print(f"已將 {len(urls)} 個搜尋結果加入到 {absolute_path} 的 Bulk 工作表中。")
    else:
        logging.info(f"未找到關鍵字/ID的結果: {keywords}")
        print(f"未找到與 '{', '.join(keywords)}' 相關的結果。")

if __name__ == "__main__":
    # 在 Windows 上，為 asyncio 設定不同的事件循環策略以避免錯誤
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    main()