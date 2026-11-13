import logging
import os
import pandas as pd
import json
import openpyxl
import time
import threading
import queue
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import requests # 確保 requests 已導入
from others.excel_handler import check_create_excel
from others.utils import is_chinese, normalize_to_hk
from extractors.tmdb_extractor_api import get_api_data_for_item
from others.credentials import load_configs

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 進度檔案路徑
PROGRESS_FILE = 'tmdb_scan_progress.json'

# 共享隊列作為暫存器
excel_queue = queue.Queue()
progress_queue = queue.Queue()

# 進度檔案鎖
progress_lock = threading.Lock()

RETRY_COUNT = 50
RETRY_BACKOFF = 0.5 # 原為 1，縮短退避因子

def load_progress():
    """載入上次掃描進度。如果檔案因寫入操作而暫時不存在，會等待其重新出現。"""
    max_wait_attempts = 10  # 最多等待 10 * 0.5 = 5 秒
    wait_interval = 0.5

    for attempt in range(max_wait_attempts):
        if os.path.exists(PROGRESS_FILE):
            try:
                # 避免讀取在寫入過程中產生的空檔案
                if os.path.getsize(PROGRESS_FILE) == 0:
                    if attempt == 0: # 首次遇到空檔案時，記錄一次
                        logging.info(f"進度檔案 {PROGRESS_FILE} 為空，可能正在寫入中。等待重試...")
                    time.sleep(wait_interval)
                    continue

                with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
                    content = f.read()

                # 再次確認讀取到的內容
                if not content.strip():
                    if attempt == 0: # 首次遇到空內容時，記錄一次
                        logging.warning(f"進度檔案 {PROGRESS_FILE} 內容為空，可能正在寫入中。等待重試...")
                    time.sleep(wait_interval)
                    continue

                data = json.loads(content)
                odd_id = data.get('odd_last_tmdb_id', 0)
                even_id = data.get('even_last_tmdb_id', 0)

                if not isinstance(odd_id, int):
                    logging.warning(f"進度檔案中 odd_last_tmdb_id 值 '{odd_id}' 非整數，將重置為 0。")
                    odd_id = 0
                if not isinstance(even_id, int):
                    logging.warning(f"進度檔案中 even_last_tmdb_id 值 '{even_id}' 非整數，將重置為 0。")
                    even_id = 0
                
                # 成功讀取，即使是在重試後，也用 INFO 級別記錄
                if attempt > 0:
                    logging.info(f"在重試 {attempt} 次後，成功從 {PROGRESS_FILE} 載入進度: 單數 ID {odd_id}, 雙數 ID {even_id}")
                else:
                    logging.info(f"成功從 {PROGRESS_FILE} 載入進度: 單數 ID {odd_id}, 雙數 ID {even_id}")
                
                return odd_id, even_id

            except json.JSONDecodeError as jde:
                # 檔案可能已損壞或正在寫入。等待重試。
                if attempt == 0:
                    logging.warning(f"解析進度檔案失敗 (JSONDecodeError): {jde}。檔案可能正在更新。等待重試...")
                time.sleep(wait_interval)
                continue
            
            except FileNotFoundError:
                # 競爭條件：檔案在 os.path.exists 和 open() 之間被移除。等待重試。
                if attempt == 0:
                    logging.warning(f"進度檔案 {PROGRESS_FILE} 在開啟時消失，可能正在更新。等待重試...")
                time.sleep(wait_interval)
                continue

            except Exception as e:
                # 捕獲其他預期外的錯誤，這時應該直接失敗。
                logging.error(f"載入進度檔案 {PROGRESS_FILE} 時發生預期外的錯誤: {e}。將返回初始進度 (0, 0)。")
                return 0, 0
        else:
            # 檔案不存在。這可能是首次運行，也可能是暫時狀態。
            if attempt == 0:
                # 這是第一次嘗試，所以只記錄一次它不存在，並開始等待。
                logging.info(f"進度檔案 {PROGRESS_FILE} 不存在，等待其出現...")
            
            time.sleep(wait_interval)

    # 如果循環完成後仍未返回，表示檔案從未出現或持續無效。
    logging.info(f"等待超時。進度檔案 {PROGRESS_FILE} 未找到或持續無效，將返回初始進度 (0, 0)。")
    return 0, 0

def save_progress(odd_id, even_id):
    """保存當前掃描進度"""
    with progress_lock:
        try:
            with open(PROGRESS_FILE, 'w') as f:
                json.dump({'odd_last_tmdb_id': odd_id, 'even_last_tmdb_id': even_id}, f)
            logging.debug(f"已保存進度: 單數 ID {odd_id}, 雙數 ID {even_id}")
        except Exception as e:
            logging.error(f"保存進度檔案失敗: {e}")

def interruptible_sleep(duration, stop_event, check_interval=0.2):
    """
    可中斷的睡眠函數。會分段睡眠，並在每段睡眠後檢查停止事件。
    返回 True 如果睡眠被中斷，否則返回 False。
    """
    slept_time = 0
    while slept_time < duration and not stop_event.is_set():
        # 計算本次睡眠時長，確保不會超過總時長，且大於0
        sleep_chunk = max(0, min(check_interval, duration - slept_time))
        if sleep_chunk == 0 and duration - slept_time > 0: # 如果計算出的 chunk 為0但仍有剩餘時間，設一個極小值
            sleep_chunk = 0.001 
        
        if sleep_chunk > 0:
            time.sleep(sleep_chunk)
            slept_time += sleep_chunk
        elif duration - slept_time <= 0: # 如果剩餘時間為0或負，則跳出
            break
        # 如果 sleep_chunk 為0且 duration - slept_time > 0 (理論上不太會進入此分支，除非 check_interval 極小)
        # 為了安全起見，避免無限循環，可以短暫休眠或直接跳出
        else: 
            time.sleep(0.001) # 短暫休眠避免CPU空轉
            slept_time += 0.001


    return stop_event.is_set()

def ensure_tw_hk_excel(file_name='tw_hk_name.xlsx'):
    """確保 tw_hk_name.xlsx 存在並包含 Movie name sheet"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            if not os.path.exists(file_name):
                df = pd.DataFrame(columns=['TW Name', 'HK Name'])
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Movie name', index=False)
                logging.info(f"已創建新的 {file_name}，包含 Movie name sheet")
            else:
                workbook = openpyxl.load_workbook(file_name)
                if 'Movie name' not in workbook.sheetnames:
                    df = pd.DataFrame(columns=['TW Name', 'HK Name'])
                    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
                        df.to_excel(writer, sheet_name='Movie name', index=False)
                    logging.info(f"已在 {file_name} 中創建 Movie name sheet")
            return
        except PermissionError:
            if attempt < max_retries - 1:
                logging.warning(f"無法訪問 {file_name}，可能正被其他程序占用，重試 {attempt + 1}/{max_retries}")
                time.sleep(1)
            else:
                logging.error(f"超過最大重試次數 {max_retries}，無法創建或訪問 {file_name}")
                raise
        except Exception as e:
            logging.error(f"創建或檢查 {file_name} 時發生錯誤: {e}")
            raise

def save_to_tw_hk_excel(zh_hk_title, zh_tw_title, zh_cn_title, file_name='tw_hk_name.xlsx'):
    """將標題差異保存到 tw_hk_name.xlsx 的 Movie name sheet，僅當 zh-HK 標題非空時記錄"""
    max_retries = 3
    rows_to_add = []
    
    # 清理和格式化標題
    zh_hk_clean = normalize_to_hk(zh_hk_title.strip()) if zh_hk_title else ""
    zh_tw_clean = normalize_to_hk(zh_tw_title.strip()) if zh_tw_title else ""
    zh_cn_clean = normalize_to_hk(zh_cn_title.strip()) if zh_cn_title else ""

    # 檢查 zh-HK 標題是否為空
    if not zh_hk_clean:
        logging.debug(f"zh-HK 標題為空，跳過差異記錄: zh-TW='{zh_tw_clean}', zh-CN='{zh_cn_clean}'")
        return

    # 檢查 zh-TW 和 zh-CN 是否相同（經香港繁體化後）
    if zh_tw_clean and zh_cn_clean and zh_tw_clean == zh_cn_clean and is_chinese(zh_tw_clean):
        if zh_tw_clean != zh_hk_clean:
            rows_to_add.append({'TW Name': f'《{zh_tw_clean}》', 'HK Name': f'《{zh_hk_clean}》'})
            logging.info(f"檢測到 zh-TW/zh-CN 與 zh-HK 差異: TW/CN='{zh_tw_clean}', HK='{zh_hk_clean}'")
    else:
        # 獨立檢查 zh-TW
        if zh_tw_clean and zh_tw_clean != zh_hk_clean and is_chinese(zh_tw_clean):
            rows_to_add.append({'TW Name': f'《{zh_tw_clean}》', 'HK Name': f'《{zh_hk_clean}》'})
            logging.info(f"檢測到 zh-TW 與 zh-HK 差異: TW='{zh_tw_clean}', HK='{zh_hk_clean}'")
        # 獨立檢查 zh-CN
        if zh_cn_clean and zh_cn_clean != zh_hk_clean and is_chinese(zh_cn_clean):
            rows_to_add.append({'TW Name': f'《{zh_cn_clean}》', 'HK Name': f'《{zh_hk_clean}》'})
            logging.info(f"檢測到 zh-CN 與 zh-HK 差異: CN='{zh_cn_clean}', HK='{zh_hk_clean}'")

    if not rows_to_add:
        logging.debug(f"無標題差異需要記錄: zh-HK='{zh_hk_clean}', zh-TW='{zh_tw_clean}', zh-CN='{zh_cn_clean}'")
        return

    for attempt in range(max_retries):
        try:
            # 讀取現有 Movie name sheet
            # Note: openpyxl.load_workbook() is not strictly needed here if only reading one sheet with pandas
            # and then replacing it. However, if 'Movie name' doesn't exist, pd.read_excel will raise error.
            # ensure_tw_hk_excel should guarantee 'Movie name' exists.
            try:
                df = pd.read_excel(file_name, sheet_name='Movie name', engine='openpyxl')
            except FileNotFoundError: # If ensure_tw_hk_excel failed or was skipped.
                logging.warning(f"Excel 檔案 {file_name} 不存在，將創建新的。")
                df = pd.DataFrame(columns=['TW Name', 'HK Name'])
            except ValueError: # If 'Movie name' sheet does not exist for some reason.
                logging.warning(f"Excel 檔案 {file_name} 中 'Movie name' 工作表不存在，將創建。")
                df = pd.DataFrame(columns=['TW Name', 'HK Name'])

            # 添加新行
            new_df = pd.DataFrame(rows_to_add)
            df = pd.concat([df, new_df], ignore_index=True)
            
            # 去重
            df.drop_duplicates(subset=['TW Name', 'HK Name'], keep='last', inplace=True)
            
            # 保存到 Movie name sheet，僅替換 Movie name sheet，保留其他 sheet
            # pd.ExcelWriter with mode='a' and if_sheet_exists='replace' should handle this correctly
            # for the openpyxl engine.
            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='Movie name', index=False)
            
            logging.info(f"已保存 {len(rows_to_add)} 筆標題差異到 {file_name} 的 Movie name sheet")
            return
        except PermissionError:
            if attempt < max_retries - 1:
                logging.warning(f"無法訪問 {file_name}，可能正被其他程序占用，重試 {attempt + 1}/{max_retries}")
                time.sleep(1) # Consider interruptible_sleep here if this sleep can be long
            else:
                logging.error(f"超過最大重試次數 {max_retries}，無法保存到 {file_name}")
                raise # Re-raise the exception to be caught by the writing thread
        except Exception as e:
            logging.error(f"保存到 {file_name} 時發生錯誤: {e}")
            raise # Re-raise the exception

def process_tmdb_titles(api_key, base_url_tmdb, start_id, is_odd, stop_event, excel_queue, progress_queue):
    """處理單數或雙數 TMDB ID，並將結果放入隊列"""
    current_tmdb_id = start_id
    parity = "odd" if is_odd else "even"
    
    # max_id 相關的邏輯是為了平衡兩個線程的進度，保持其大致同步。
    # 這裡的 max_id 計算和使用方式可能需要根據實際運行情況調整，但不是當前修復的重點。
    # 移除 other_last_id 和 max_id 的初始加載，因為 current_max_id_for_check 會在循環中更新。

    while not stop_event.is_set():
        # 動態獲取另一個線程的進度來決定是否需要等待
        # 注意: load_progress() 可能會頻繁讀取檔案，如果有效能問題可以考慮優化，
        # 例如讓 write_progress_thread 也將最新進度放入一個共享變數中（需加鎖保護）。
        # 但目前保持原樣，因為主要問題在於 Ctrl+C 和寫入。
        other_thread_last_id = load_progress()[1 if is_odd else 0]
        if current_tmdb_id > other_thread_last_id + 200 and current_tmdb_id > start_id + 200 : # 與另一個線程ID差距過大時等待
            logging.debug(f"{parity} 線程 ID {current_tmdb_id} 暫停，等待另一線程 (另一線程最後 ID: {other_thread_last_id})")
            if interruptible_sleep(5, stop_event): # 使用可中斷的睡眠
                logging.info(f"{parity} 線程在等待另一線程時被停止。")
                break 
            continue

        api_call_start_time = None # 初始化
        try:
            logging.info(f"{parity} 線程準備處理電影 TMDB ID: {current_tmdb_id}")
            api_call_start_time = time.time()
            api_data = get_api_data_for_item(
                api_key=api_key,
                base_url_tmdb=base_url_tmdb,
                endpoint_suffix=f"/movie/{current_tmdb_id}",
                item_is_person=False,
                preferred_language="zh-HK",
                fallback_language="zh-TW", # zh-TW 作為一個常見的後備
                context_for_logging=f"Movie {current_tmdb_id} ({parity})"
            )
            api_call_duration = time.time() - api_call_start_time
            logging.info(f"{parity} 線程完成 API 調用 (TMDB ID: {current_tmdb_id}), 耗時: {api_call_duration:.2f} 秒")

            zh_hk_title = api_data.get("zh_hk_title", "")
            zh_tw_title = api_data.get("zh_tw_title", "")
            zh_cn_title = api_data.get("zh_cn_title", "")

            if zh_hk_title or zh_tw_title or zh_cn_title: # 只要任一標題存在
                excel_queue.put((zh_hk_title, zh_tw_title, zh_cn_title))
                logging.debug(f"{parity} 線程將 TMDB ID {current_tmdb_id} 的標題加入 Excel 隊列")
            else:
                logging.info(f"{parity} 線程 TMDB ID {current_tmdb_id}: API返回數據中無有效中文標題可供處理。")
            
            progress_queue.put((is_odd, current_tmdb_id))
            logging.debug(f"{parity} 線程將 TMDB ID {current_tmdb_id} 的進度加入進度隊列")
            
            current_tmdb_id += 2

        except requests.exceptions.HTTPError as e:
            if api_call_start_time:
                api_call_duration = time.time() - api_call_start_time
                logging.error(f"{parity} 線程處理 TMDB ID {current_tmdb_id} 時發生 HTTP 錯誤 (耗時: {api_call_duration:.2f} 秒): {e}", exc_info=False) # 通常 exc_info=True 更好
            else:
                logging.error(f"{parity} 線程處理 TMDB ID {current_tmdb_id} 時發生 HTTP 錯誤: {e}", exc_info=False)

            if e.response is not None and e.response.status_code == 429:
                retry_after_seconds = int(e.response.headers.get("Retry-After", 60))
                logging.warning(f"{parity} 線程遇到 429 錯誤於 TMDB ID {current_tmdb_id}，等待 {retry_after_seconds} 秒後重試")
                if interruptible_sleep(retry_after_seconds, stop_event): # 使用可中斷的睡眠
                    logging.info(f"{parity} 線程在處理 429 錯誤等待時被停止。")
                    break # 跳出主 while 循環
                # continue # 重試當前 ID (隱含在外層循環)
            else: # 其他 HTTP 錯誤，記錄進度並嘗試下一個ID
                progress_queue.put((is_odd, current_tmdb_id)) 
                current_tmdb_id += 2
            continue # 對於 HTTPError (包括429後的正常結束等待)，繼續下一次循環嘗試（可能重試同一ID或下一個）
        
        except Exception as e:
            if api_call_start_time:
                api_call_duration = time.time() - api_call_start_time
                logging.error(f"{parity} 線程處理 TMDB ID {current_tmdb_id} 時發生一般錯誤 (耗時: {api_call_duration:.2f} 秒): {e}", exc_info=True)
            else:
                logging.error(f"{parity} 線程處理 TMDB ID {current_tmdb_id} 時發生一般錯誤: {e}", exc_info=True)
            
            progress_queue.put((is_odd, current_tmdb_id)) # 記錄當前失敗的ID
            current_tmdb_id += 2
            if interruptible_sleep(1, stop_event): # 短暫等待，可中斷
                logging.info(f"{parity} 線程在一般錯誤後等待時被停止。")
                break # 跳出主 while 循環
            continue
    logging.info(f"{parity} 線程正常結束或因停止信號而終止。")

def write_to_excel_thread(data_q, stop_event):
    """從隊列中取出數據並寫入 Excel"""
    while not stop_event.is_set() or not data_q.empty():
        item = None  # 初始化 item 變數
        try:
            item = data_q.get(timeout=1)
            zh_hk_title, zh_tw_title, zh_cn_title = item
            save_to_tw_hk_excel(zh_hk_title, zh_tw_title, zh_cn_title)
        except queue.Empty:
            # 超時，隊列為空，繼續檢查 stop_event 和隊列狀態
            continue
        except Exception as e:
            # 處理從 save_to_tw_hk_excel 拋出的或其他意外錯誤
            logging.error(f"寫入 Excel 時發生錯誤 (項目: {item}): {e}", exc_info=True)
            # 根據需要，可以考慮將 item 重新放回隊列或其他錯誤處理策略
        finally:
            if item is not None: # 確保只有在成功取出項目後才調用 task_done
                try:
                    data_q.task_done()
                except ValueError: # task_done() called too many times
                    logging.warning(f"對項目 {item} 嘗試調用 task_done 時出錯 (可能已被處理或多次調用)")
    logging.info("Excel 寫入線程正常結束。")

def write_progress_thread(progress_q, stop_event): # 將參數 queue 重命名為 progress_q
    """從進度隊列中取出更新並即時寫入進度檔案，使用更安全的寫入方式"""
    import os # 確保 os 已導入
    # 嘗試載入初始進度，如果 load_progress 內部處理了錯誤並返回0,0，則從0,0開始
    odd_last_id, even_last_id = load_progress()
    logging.info(f"進度寫入線程已啟動。初始載入進度: 單數 ID {odd_last_id}, 雙數 ID {even_last_id}")

    max_retries = 3
    retry_interval = 1 # 秒

    while not stop_event.is_set() or not progress_q.empty():
        item_data = None # 初始化 item_data
        try:
            item_data = progress_q.get(timeout=1) # 等待1秒獲取項目，以便能周期性檢查 stop_event
            is_odd, tmdb_id = item_data
            
            changed = False
            # 使用臨時變量來確定是否真的需要更新文件中的值
            # 這確保了即使有多個相同的進度更新請求，也只在ID增加時寫入
            if is_odd:
                if tmdb_id > odd_last_id:
                    odd_last_id = tmdb_id # 更新線程內部的最新ID
                    changed = True
            else:
                if tmdb_id > even_last_id:
                    even_last_id = tmdb_id # 更新線程內部的最新ID
                    changed = True

            if changed: # 只有在進度實際更新時才寫入檔案
                for attempt in range(max_retries):
                    if stop_event.is_set(): # 在嘗試寫入前再次檢查停止事件
                        logging.info(f"進度寫入線程檢測到停止信號，放棄保存進度 (單: {odd_last_id}, 雙: {even_last_id})。")
                        break 
                    try:
                        progress_path_abs = os.path.abspath(PROGRESS_FILE)
                        temp_progress_file = progress_path_abs + ".tmp" # 臨時文件名
                        
                        # 確保臨時文件所在的目錄存在 (通常不需要，除非 PROGRESS_FILE 在配置的深層子目錄)
                        # temp_dir = os.path.dirname(temp_progress_file)
                        # if temp_dir and not os.path.exists(temp_dir):
                        #     os.makedirs(temp_dir, exist_ok=True)

                        # 寫入到臨時文件
                        with open(temp_progress_file, 'w', encoding='utf-8') as f_temp:
                            json.dump({'odd_last_tmdb_id': odd_last_id, 'even_last_tmdb_id': even_last_id}, f_temp, ensure_ascii=False, indent=4)
                        
                        # 原子性替換原文件 (在同一文件系統下通常是原子操作)
                        os.replace(temp_progress_file, progress_path_abs)
                        logging.debug(f"即時安全保存進度: 單數 ID {odd_last_id}, 雙數 ID {even_last_id} 到 {progress_path_abs}")
                        break # 寫入成功，跳出重試循環
                    except (OSError, IOError) as e_io: # 包括 PermissionError 等
                        if attempt < max_retries - 1:
                            logging.warning(f"第 {attempt + 1}/{max_retries} 次嘗試安全保存進度檔案 '{progress_path_abs}' 失敗 (IOError/OSError): {e_io}。將在 {retry_interval} 秒後重試。")
                            if interruptible_sleep(retry_interval, stop_event, check_interval=0.1): # 使用可中斷的睡眠
                                logging.info("進度寫入線程在重試保存期間被停止。")
                                break # 從重試循環中跳出
                        else:
                            logging.error(f"安全保存進度檔案 '{progress_path_abs}' 失敗，已達最大重試次數 {max_retries} (IOError/OSError): {e_io}")
                            # 即使保存失敗，也應將任務標記為完成，以避免隊列阻塞
                    except Exception as e_exc: # 捕獲其他可能的錯誤，如 json.dump 內部問題 (較少見)
                        if attempt < max_retries - 1:
                            logging.warning(f"第 {attempt + 1}/{max_retries} 次嘗試安全保存進度檔案 '{progress_path_abs}' 時發生未預期錯誤: {e_exc}。將在 {retry_interval} 秒後重試。", exc_info=True)
                            if interruptible_sleep(retry_interval, stop_event, check_interval=0.1):
                                logging.info("進度寫入線程在重試保存期間因未預期錯誤被停止。")
                                break
                        else:
                            logging.error(f"安全保存進度檔案 '{progress_path_abs}' 因未預期錯誤失敗，已達最大重試次數 {max_retries}: {e_exc}", exc_info=True)
                if stop_event.is_set() and changed:
                     logging.warning(f"進度寫入線程被停止，最後更新的進度 (單: {odd_last_id}, 雙: {even_last_id}) 可能未完全寫入或寫入已終止。")
            else:
                logging.debug(f"收到的進度 TMDB ID {tmdb_id} ({'單數' if is_odd else '雙數'}) 未更新線程內最高記錄 ({odd_last_id if is_odd else even_last_id})，不寫入檔案。")

        except queue.Empty:
            # 隊列為空時，如果 stop_event 未設置，則繼續循環等待
            # 如果 stop_event 已設置，則循環條件 (not stop_event.is_set() or not progress_q.empty()) 將使其退出
            if stop_event.is_set():
                logging.debug("進度隊列為空且停止事件已設置，進度寫入線程準備退出主循環。")
            continue # 繼續下一次循環檢查
        except Exception as e: # 捕獲 progress_q.get() 或 item_data 解包時的錯誤
            logging.error(f"處理進度隊列項目時發生錯誤 (項目數據: {item_data}): {e}", exc_info=True)
            # 發生這種錯誤時，可能需要決定是否將 item_data 放回隊列或丟棄
        finally:
            if item_data is not None: # 確保只有在成功取出項目後才調用 task_done
                try:
                    progress_q.task_done()
                except ValueError: # task_done() called too many times
                     logging.warning(f"對進度項目 {item_data} 嘗試調用 task_done 時出錯 (可能已被處理或多次調用)")

    # 線程即將結束（stop_event 已設置且隊列已空）
    # 嘗試最後一次保存當前線程持有的最新進度，以防萬一主循環中最後的更改未被寫入
    logging.info(f"進度寫入線程主循環結束。準備進行最後的進度保存。當前線程持有進度: 單數 ID {odd_last_id}, 雙數 ID {even_last_id}。")
    try:
        # 檢查是否有必要進行最後一次保存 (例如，如果這些值比文件中最新的還要新)
        # 為簡化，這裡直接嘗試保存當前線程持有的最新值。
        # 也可以先 load_progress() 一次比較，但要注意 load_progress 可能因文件問題返回0,0。
        # 假設 odd_last_id 和 even_last_id 是最準確的。
        final_progress_path_abs = os.path.abspath(PROGRESS_FILE)
        final_temp_progress_file = final_progress_path_abs + ".final_save.tmp" # 使用不同的臨時文件名以防衝突
        
        # 檢查目錄是否存在，以防萬一
        # final_temp_dir = os.path.dirname(final_temp_progress_file)
        # if final_temp_dir and not os.path.exists(final_temp_dir):
        #     os.makedirs(final_temp_dir, exist_ok=True)

        with open(final_temp_progress_file, 'w', encoding='utf-8') as f_final_temp:
            json.dump({'odd_last_tmdb_id': odd_last_id, 'even_last_tmdb_id': even_last_id}, f_final_temp, ensure_ascii=False, indent=4)
        os.replace(final_temp_progress_file, final_progress_path_abs)
        logging.info(f"進度寫入線程結束前，已成功安全保存最終進度: 單數 ID {odd_last_id}, 雙數 ID {even_last_id} 到 {final_progress_path_abs}")
    except Exception as e_final_save:
        logging.error(f"進度寫入線程結束前，嘗試最終安全保存進度失敗: {e_final_save}", exc_info=True)
    
    logging.info("進度寫入線程已結束。")

def create_retry_session():
    """創建帶重試機制的 requests Session"""
    session = requests.Session()
    retries = Retry(
        total=RETRY_COUNT,
        backoff_factor=RETRY_BACKOFF,
        status_forcelist=[429, 500, 502, 503, 504, 408], # 包括 408 (Request Timeout)
        allowed_methods=["GET"] # 只對 GET 請求重試
    )
    session.mount('https://', HTTPAdapter(max_retries=retries))
    return session

def scan_tmdb_titles():
    """掃描 TMDB 電影標題並記錄差異，使用多線程"""
    configs = load_configs()
    api_key = configs.get('tmdb_api', '')
    if not api_key:
        logging.error("未找到 TMDB API 密鑰")
        return

    base_url_tmdb = "https://api.themoviedb.org/3"
    odd_last_id, even_last_id = load_progress()
    odd_start_id = odd_last_id + 2 if odd_last_id % 2 == 1 else odd_last_id + 1
    even_start_id = even_last_id + 2 if even_last_id % 2 == 0 else even_last_id + 1

    # 確保 Excel 檔案存在
    ensure_tw_hk_excel()
    check_create_excel()

    # 檢查 Movie name sheet 中的重複項目
    file_name = 'tw_hk_name.xlsx'
    try:
        df = pd.read_excel(file_name, sheet_name='Movie name', engine='openpyxl')
        if not df.empty:
            duplicates = df[df.duplicated(subset=['TW Name', 'HK Name'], keep=False)]
            if not duplicates.empty:
                logging.warning(f"檢測到 Movie name sheet 中有 {len(duplicates)} 筆重複項目：")
                for _, row in duplicates.iterrows():
                    logging.warning(f"重複項目: TW Name='{row['TW Name']}', HK Name='{row['HK Name']}'")
            else:
                logging.info("Movie name sheet 中無重複項目")
        else:
            logging.info("Movie name sheet 為空，無需檢查重複項目")
    except Exception as e:
        logging.error(f"檢查 Movie name sheet 重複項目時發生錯誤: {e}")

    logging.info(f"開始掃描 TMDB 電影，單數從 ID {odd_start_id}，雙數從 ID {even_start_id}")

    # 創建停止事件
    stop_event = threading.Event()

    # 啟動處理線程
    odd_thread = threading.Thread(target=process_tmdb_titles, args=(api_key, base_url_tmdb, odd_start_id, True, stop_event, excel_queue, progress_queue))
    even_thread = threading.Thread(target=process_tmdb_titles, args=(api_key, base_url_tmdb, even_start_id, False, stop_event, excel_queue, progress_queue))
    excel_writer_thread = threading.Thread(target=write_to_excel_thread, args=(excel_queue, stop_event)) # 將 write_excel_thread 重命名以更清晰
    progress_writer_thread = threading.Thread(target=write_progress_thread, args=(progress_queue, stop_event)) # 重命名局部變數，避免與函數同名

    odd_thread.daemon = True
    even_thread.daemon = True
    excel_writer_thread.daemon = True # 使用新的變數名
    progress_writer_thread.daemon = True # 使用新的變數名

    odd_thread.start()
    even_thread.start()
    excel_writer_thread.start() # 使用新的變數名
    progress_writer_thread.start() # 使用新的變數名

    try:
        # 主線程等待中斷
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("收到中斷信號，準備停止線程")
        stop_event.set()
        # 等待隊列清空
        excel_queue.join()
        progress_queue.join()
        logging.info("所有數據和進度已保存，程式結束")

def _set_working_directory():
    """將工作目錄設定為本腳本所在的目錄，以確保所有相對路徑的檔案都能被正確存取。"""
    try:
        # 獲取腳本檔案所在的目錄
        script_dir = os.path.dirname(os.path.abspath(__file__))
        # 將當前工作目錄更改為腳本所在的目錄
        os.chdir(script_dir)
        logging.info(f"工作目錄已成功設定為: {os.getcwd()}")
    except Exception as e:
        logging.error(f"設定工作目錄時發生錯誤: {e}")
        # 如果設定失敗，可能需要考慮是否要終止程式
        raise

def main():
    """主函數"""
    try:
        # 在執行任何操作前，先設定好工作目錄
        _set_working_directory()
        scan_tmdb_titles()
    except Exception as e:
        logging.error(f"程式發生未預期的錯誤: {e}")

if __name__ == "__main__":
    main()