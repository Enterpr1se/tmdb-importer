import asyncio
import logging
import os
import re
import sys
import time
import json
import traceback
import httpx
import pandas as pd
import openpyxl
import psutil
import shutil
import platform
import subprocess

# --- Selenium and WebDriver Manager Imports ---
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# --- 基本設定 ---
# 設定日誌記錄，方便追蹤程式執行狀況
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- 全域常數 ---
# API 和並行請求設定
MAX_CONCURRENT_REQUESTS = 30
TMDB_API_BASE_URL = "https://api.themoviedb.org/3"
TMDB_WEB_BASE_URL = "https://www.themoviedb.org"

# =====================================================================================
#  整合自 chrome_process.py 的瀏覽器管理函數
#  為了讓此腳本能獨立運作，將瀏覽器啟動、關閉及 WebDriver 管理的相關函數整合進來
# =====================================================================================

def get_browser_driver(remote_port=9223):
    """
    (使用者需求 2025-08-22) 實現雙模式啟動，使其更具彈性。
    1.  優先嘗試附加到一個在指定埠號 (預設 9223) 上運行的現有 Chrome 實例（與主程式共用模式）。
        -   附加成功後，會開啟一個全新的獨立視窗進行所有操作，避免干擾。
    2.  如果附加失敗（例如主程式未運行），則自動切換到獨立模式。
        -   在獨立模式下，程式會自己啟動一個全新的、使用獨立設定檔的 Chrome 實例來運作。
    """
    logger = logging.getLogger(__name__)
    
    # --- 模式一：嘗試附加到現有瀏覽器 ---
    try:
        logger.info(f"--- 模式一：嘗試附加到在埠 {remote_port} 上運行的現有 Chrome 實例 ---")
        chrome_options = ChromeOptions()
        chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{remote_port}")
        driver = webdriver.Chrome(options=chrome_options)
        logger.info("成功附加到現有的 Chrome 瀏覽器！")
        
        original_windows = driver.window_handles
        driver.execute_script("window.open('about:blank', '_blank');")
        WebDriverWait(driver, 10).until(EC.new_window_is_opened(original_windows))
        
        new_window_handle = [handle for handle in driver.window_handles if handle not in original_windows][0]
        driver.switch_to.window(new_window_handle)
        logger.info(f"已開啟並切換到一個全新的獨立視窗 (Handle: {new_window_handle}) 進行操作。")
        
        # 返回 driver, profile_path=None, 以及一個布林值表示是否為獨立啟動
        return driver, None, False # False = 不是獨立啟動 (是附加模式)
    except Exception as e:
        logger.warning(f"附加到現有 Chrome 失敗: {e}")
        logger.info("--- 將切換到模式二：獨立啟動一個新的 Chrome 實例 ---")

    # --- 模式二：獨立啟動一個新的瀏覽器 ---
    try:
        profile_dir_name = f"browser_profile_person_search_standalone"
        profile_path = os.path.join(os.getcwd(), profile_dir_name)
        os.makedirs(profile_path, exist_ok=True)
        logger.info(f"獨立模式：瀏覽器設定檔將儲存於: {profile_path}")

        chrome_options = ChromeOptions()
        # 使用一個不同的埠號，避免與主程式預設的 9223 衝突
        standalone_port = remote_port + 1 
        common_options = [
            '--log-level=3',
            '--ignore-certificate-errors',
            '--no-sandbox',
            f'--user-data-dir={profile_path}',
            f'--remote-debugging-port={standalone_port}'
        ]
        for option in common_options:
            chrome_options.add_argument(option)

        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        logger.info("獨立模式：成功啟動一個全新的 Chrome 實例。")
        
        # 返回 driver, profile_path, 以及布林值
        return driver, profile_path, True # True = 是獨立啟動
    except Exception as e_standalone:
        logger.error(f"獨立啟動新 Chrome 實例也失敗: {e_standalone}")
        raise Exception("無法附加也無法獨立啟動 Chrome，請檢查您的 Chrome 安裝與網路連線。")

def cleanup_browser(driver, profile_path, is_standalone_instance):
    """
    安全地關閉瀏覽器。
    - 如果是獨立啟動的實例，則完全關閉瀏覽器 (driver.quit())。
    - 如果是附加到現有瀏覽器的模式，則只關閉由腳本開啟的新視窗 (driver.close())。
    """
    logger = logging.getLogger(__name__)
    if not driver:
        logger.info("無有效的 WebDriver 實例，跳過清理。")
        return
    try:
        if is_standalone_instance:
            logger.info("獨立模式：正在完全關閉由本程式開啟的瀏覽器...")
            driver.quit()
            logger.info("瀏覽器已完全關閉。")
        else:
            logger.info("附加模式：正在關閉由本程式開啟的視窗...")
            driver.close()
            logger.info("程式視窗已關閉。您手動開啟的原始瀏覽器視窗不受影響。")
    except Exception as e:
        logger.error(f"關閉瀏覽器/視窗時發生錯誤: {e}")

# =====================================================================================
#  整合自 tmdb_search.py 和 bulk_video_person.py 的輔助函數
# =====================================================================================

def load_configs():
    """從 ../configs.json 載入設定檔。"""
    config_path = os.path.join(os.path.dirname(__file__), '../configs.json')
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"讀取 configs.json 時發生錯誤: {e}")
    logging.warning(f"設定檔 {config_path} 不存在。")
    return {}

def check_create_bulk_excel(file_name='../data/bulk_videos.xlsx'):
    """檢查並創建 bulk_videos.xlsx，確保包含 Bulk 工作表及必要欄位。"""
    output_dir = os.path.dirname(file_name)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    required_columns = ['Video URL', 'TMDB URL', 'Split', 'Done', '覆蓋劇集標題', '覆蓋劇集描述', '覆蓋分集標題', '覆蓋分集描述', '來源無資料時清空TMDB', '維基人名取代現有']
    
    if not os.path.exists(file_name):
        pd.DataFrame(columns=required_columns).to_excel(file_name, sheet_name='Bulk', index=False)
        logging.info(f"已創建新的 {file_name}。")
    else:
        # 檢查工作表和欄位
        try:
            wb = openpyxl.load_workbook(file_name)
            if 'Bulk' not in wb.sheetnames:
                with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
                    pd.DataFrame(columns=required_columns).to_excel(writer, sheet_name='Bulk', index=False)
                logging.info(f"已為 {file_name} 添加 Bulk 工作表。")
            else:
                ws = wb['Bulk']
                header = [cell.value for cell in ws[1]]
                if not all(col in header for col in required_columns):
                    logging.warning(f"{file_name} 的 Bulk 工作表欄位不完整，建議手動修正。")
        except Exception as e:
            logging.error(f"檢查 {file_name} 時發生錯誤: {e}")

def append_to_bulk_excel(urls, file_name='../data/bulk_videos.xlsx'):
    """
    將 URL 添加到 bulk_videos.xlsx 的 Bulk 工作表中，並包含完整的檔案鎖定和重試機制。
    """
    check_create_bulk_excel(file_name)
    lock_file = file_name + ".lock"
    output_filename = os.path.basename(file_name)
    
    logging.info(f"準備寫入 '{output_filename}'，開始檢查檔案鎖...")

    # 階段一：獲取腳本鎖
    lock_acquired = False
    for _ in range(60): # 最多等待 60 秒
        try:
            lock_fd = os.open(lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(lock_fd)
            lock_acquired = True
            logging.info(f"成功獲取 '{output_filename}' 的腳本寫入權限。")
            break
        except FileExistsError:
            time.sleep(1)
        except Exception as e:
            logging.error(f"獲取檔案鎖時發生未知錯誤: {e}")
            return
    
    if not lock_acquired:
        logging.error(f"等待超時，無法獲取 '{output_filename}' 的寫入鎖。")
        return

    # 階段二：持有鎖進行寫入
    try:
        write_successful = False
        for _ in range(60): # 最多等待 60 秒
            try:
                # 讀取現有數據
                try:
                    existing_df = pd.read_excel(file_name, sheet_name='Bulk')
                except (FileNotFoundError, ValueError):
                    existing_df = pd.DataFrame(columns=['Video URL'])

                # 篩選出新的 URL
                new_urls_df = pd.DataFrame(urls, columns=['Video URL'])
                existing_urls = set(existing_df['Video URL'].dropna())
                new_unique_df = new_urls_df[~new_urls_df['Video URL'].isin(existing_urls)]

                if not new_unique_df.empty:
                    logging.info(f"發現 {len(new_unique_df)} 個新的 URL，將其追加到 {output_filename} 中。")
                    final_df = pd.concat([existing_df, new_unique_df], ignore_index=True)
                    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        final_df.to_excel(writer, sheet_name='Bulk', index=False)
                else:
                    logging.info("本次運行未發現新的 URL，Excel 檔案內容未變動。")
                
                write_successful = True
                break
            except PermissionError:
                logging.warning(f"'{output_filename}' 正被外部程式(如 Excel)使用中，等待 1 秒後重試...")
                time.sleep(1)
            except Exception as e:
                logging.error(f"寫入 Excel 時發生非預期的錯誤: {traceback.format_exc()}")
                break
        
        if not write_successful:
            logging.error(f"重試多次後仍無法寫入 '{output_filename}'。")

    finally:
        # 確保釋放鎖
        try:
            os.remove(lock_file)
            logging.info(f"已釋放 '{output_filename}' 的寫入權限。")
        except OSError as e:
            logging.error(f"釋放檔案鎖 '{lock_file}' 時失敗: {e}")

# =====================================================================================
#  全新 Web 搜尋及登入相關函數
# =====================================================================================

def handle_network_and_rate_limit(driver, url_to_load=None):
    """
    (使用者需求 2025-08-26) 整合了網路錯誤和速率限制的處理。
    此函數會在每次頁面載入後被呼叫，以確保頁面處於可用狀態。
    - 如果偵測到網路中斷 (如 chrome-error)，會等待並重試載入。
    - 如果偵測到 TMDB 的速率限制頁面 ("Wait and try again")，會等待更長時間並重試。
    """
    # (使用者需求 2025-08-26) 修正 NameError: 'logger' is not defined
    logger = logging.getLogger(__name__)
    retry_count = 0
    while True:
        try:
            # 檢查1：網路錯誤頁面
            if "chrome-error://" in driver.current_url:
                raise Exception("Browser is on a network error page.")
            
            # 檢查2：TMDB 速率限制頁面
            # 使用較短的超時快速檢查，避免在正常頁面浪費時間
            WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'Wait and try again')]"))
            )
            
            # 如果找到速率限制訊息
            retry_count += 1
            wait_minutes = min(retry_count, 10)
            wait_seconds = wait_minutes * 60
            logger.warning(f"發現 'Wait and try again' 訊息。將等待 {wait_minutes} 分鐘後重試 (第 {retry_count} 次)...")
            time.sleep(wait_seconds)
            
            # 重新載入頁面
            reload_target = url_to_load if url_to_load else driver.current_url
            logger.info(f"正在嘗試重新載入頁面: {reload_target}")
            driver.get(reload_target)
            continue # 繼續迴圈，再次檢查

        except Exception as e:
            error_msg = str(e).lower()
            # 判斷是否為網路錯誤
            if "net::" in error_msg or "dns" in error_msg or "chrome-error" in error_msg:
                retry_count += 1
                wait_seconds = 15
                logger.warning(f"偵測到網路中斷或瀏覽器錯誤 (第 {retry_count} 次)。將在 {wait_seconds} 秒後重試...")
                time.sleep(wait_seconds)
                
                reload_target = url_to_load if url_to_load else driver.current_url
                logger.info(f"正在嘗試重新載入頁面: {reload_target}")
                try:
                    driver.get(reload_target)
                except Exception as e_reload:
                    logger.error(f"重新載入時再次失敗: {e_reload}")
                continue # 繼續迴圈，再次檢查
            
            # 如果不是網路錯誤，也不是速率限制訊息 (TimeoutException)，則代表頁面正常
            if retry_count > 0:
                logger.info("網路連線或頁面狀態已恢復。")
            return True

def accept_all_cookies(driver):
    """
    (使用者需求 2025-09-02) 新增：處理網站可能出現的 Cookie 同意橫幅。
    此函數會嘗試點擊多種常見的 "接受所有" 按鈕。
    它設計為非阻塞性，如果找不到按鈕，只會記錄訊息而不會引發錯誤。
    """
    logger = logging.getLogger(__name__)
    # (使用者需求 2025-09-02) 參考主程式的 cookie 處理邏輯
    # 定義可能的 Cookie 按鈕 ID、class 或 xpath
    # 這裡以 TMDB 的 onetrust-accept-btn-handler 為主要目標
    possible_selectors = [
        (By.ID, "onetrust-accept-btn-handler"),
        # 可以擴展以處理其他網站
        # (By.XPATH, "//button[contains(text(), 'Accept All')]"),
        # (By.XPATH, "//button[contains(text(), '同意所有')]"),
    ]
    
    for by, selector in possible_selectors:
        try:
            # 使用短暫等待，避免在沒有 Cookie 橫幅的頁面浪費太多時間
            accept_button = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((by, selector))
            )
            # 使用 JavaScript 點擊，更為可靠
            driver.execute_script("arguments[0].click();", accept_button)
            logger.info(f"成功點擊 Cookie 同意按鈕 (Selector: {by}='{selector}')。")
            time.sleep(1) # 等待橫幅消失
            return True # 找到並點擊後就返回
        except:
            # TimeoutException 或 NoSuchElementException
            logger.debug(f"未找到 Cookie 同意按鈕 (Selector: {by}='{selector}')。")
            pass # 繼續嘗試下一個選擇器
            
    logger.info("檢查完成，無需處理 Cookie 或未找到已知的同意按鈕。")
    return False

def login_to_tmdb(driver, username, password):
    """
    (使用者需求 2025-08-21) 採用更可靠的登入檢測邏輯。
    1. 直接導航至登入頁面。
    2. 如果網站自動重定向到使用者個人頁面 (URL 包含 /u/)，則表示已登入。
    3. 如果頁面停留在登入頁，則執行自動登入流程。
    這個方法比檢測頁面元素更穩定，不受 cookie 同步延遲影響。
    (使用者需求 2025-09-02) 新增自動接受 Cookie 功能。
    """
    logger = logging.getLogger(__name__)
    try:
        logger.info("正在導航至 TMDB 登入頁面以檢查狀態...")
        driver.get(f"{TMDB_WEB_BASE_URL}/login")
        handle_network_and_rate_limit(driver, url_to_load=f"{TMDB_WEB_BASE_URL}/login")

        # (使用者需求 2025-09-02) 載入頁面後，嘗試接受 Cookie
        accept_all_cookies(driver)

        # 等待頁面穩定，給予重定向足夠的時間
        # WebDriverWait 會比 time.sleep 更可靠
        try:
            # 我們給 3 秒的時間，看 URL 是否會變成包含 /u/ 的個人頁面
            WebDriverWait(driver, 3).until(EC.url_contains("/u/"))
            logger.info(f"偵測到已登入，當前 URL: {driver.current_url}")
            return True
        except Exception: # 主要捕捉 TimeoutException
            # 如果 3 秒後 URL 沒有改變，代表未登入，繼續執行登入流程
            logger.info("未偵測到自動重定向，頁面停留在登入頁，將執行自動登入...")

        # 等待登入表單元素加載完成
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username")))
        
        # (使用者需求 2025-09-02) 再次嘗試接受 Cookie，以防登入表單加載後才出現
        accept_all_cookies(driver)
        
        logger.info("輸入使用者名稱和密碼...")
        driver.find_element(By.ID, "username").send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)
        
        logger.info("點擊登入按鈕...")
        driver.find_element(By.ID, "login_button").click()
        
        # 登入後，等待 URL 變為使用者個人頁面來確認成功
        WebDriverWait(driver, 10).until(
            EC.url_contains("/u/")
        )
        logger.info(f"TMDB 登入成功！已跳轉至: {driver.current_url}")
        return True
    except Exception as e:
        logger.error(f"TMDB 登入或狀態檢查失敗: {e}")
        logger.error("請確認您的帳號密碼是否正確，或在手動開啟的 Chrome 中先登入 TMDB。")
        return False

def search_person_url_web(driver, person_name):
    """
    使用網頁介面搜尋演員，並返回所有搜尋結果的個人頁面 URL 列表。
    - 使用者需求：必須使用網頁搜尋，因為 API 不會返回 NSFW 相關演員。
    - 速率限制：此函數在每次搜尋後會暫停 1 秒，以避免觸發網站的速率限制。
    - (使用者需求 2025-08-21) 修改為抓取所有搜尋結果，而不僅僅是第一個。
    - (使用者需求 2025-08-21) 新增分頁處理，抓取所有頁面的結果。
    - (使用者需求 2025-08-21) 新增對 "Wait and try again" 速率限制頁面的處理。
    """
    logger = logging.getLogger(__name__)
    found_urls = []
    page_num = 1
    
    try:
        search_url = f"{TMDB_WEB_BASE_URL}/search/person?query={person_name}"
        logger.info(f"正在搜尋演員 '{person_name}'，URL: {search_url}")
        driver.get(search_url)

        while True:
            # (使用者需求 2025-08-26) 在每次循環（包括初次和翻頁後）開始時，檢查頁面狀態
            handle_network_and_rate_limit(driver, url_to_load=driver.current_url)

            logger.info(f"正在抓取第 {page_num} 頁的結果...")
            
            try:
                # 等待當前頁面的結果容器加載
                wait = WebDriverWait(driver, 10)
                results_container = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.search_results.person"))
                )
                
                # 抓取當前頁面的所有人物連結
                result_links = results_container.find_elements(By.CSS_SELECTOR, "div.item.profile .content .name a")
                
                page_found_count = 0
                if result_links:
                    for link in result_links:
                        person_url = link.get_attribute('href')
                        if person_url and person_url not in found_urls:
                            found_urls.append(person_url)
                            page_found_count += 1
                
                logger.info(f"第 {page_num} 頁找到 {page_found_count} 個新結果。")

            except Exception as e: # 主要捕捉 TimeoutException
                logger.warning(f"在第 {page_num} 頁等待搜尋結果時發生錯誤: {e}")
                logger.info("可能是網路中斷或頁面載入不完整，將觸發恢復機制並重試本頁。")
                # 不用做任何事，下一次迴圈開始時 handle_network_and_rate_limit 會處理
                continue

            # 尋找「下一頁」按鈕並點擊
            try:
                # (使用者需求 2025-08-21) 修正了下一頁按鈕的選擇器，使其更精確。
                # 舊的選擇器 'nav.pagy a[aria-label="下一頁"]' 可能會錯誤地點擊到隱藏的「電影」分頁按鈕。
                # 新的選擇器 'div.search_results.person nav.pagy a[aria-label="下一頁"]' 確保只點擊人物搜尋結果區塊內的下一頁按鈕。
                next_page_button = driver.find_element(By.CSS_SELECTOR, 'div.search_results.person nav.pagy a[aria-label="下一頁"]')
                
                # (使用者需求 2025-08-21) 新增檢查：確認「下一頁」按鈕是否被禁用。
                # TMDB 在最後一頁時，會為「下一頁」按鈕添加 aria-disabled="true" 屬性。
                if next_page_button.get_attribute('aria-disabled') == 'true':
                    logger.info("「下一頁」按鈕已禁用，已到達最後一頁。")
                    break

                logger.info("找到「下一頁」按鈕，準備前往下一頁...")
                # 使用 JS 點擊更穩定，可以避免元素被遮擋等問題
                driver.execute_script("arguments[0].click();", next_page_button)
                page_num += 1
                # 等待一小段時間讓頁面跳轉
                time.sleep(2) 
            except Exception: # 主要捕捉 NoSuchElementException
                logger.info("未找到「下一頁」按鈕，已到達最後一頁。")
                break # 找不到按鈕，跳出迴圈

    except Exception as e:
        logger.error(f"搜尋 '{person_name}' 過程中發生無法恢復的錯誤: {e}")
    
    logger.info(f"對 '{person_name}' 的搜尋完成，共在 {page_num} 頁中找到 {len(found_urls)} 個不重複的結果。")
    return found_urls

# =====================================================================================
#  整合自 bulk_video_person.py 的 API 深度搜尋函數
#  這些函數將在透過網頁獲取到演員 URL 後，使用 API 進行高效的資料擴展
# =====================================================================================

def parse_tmdb_url(url):
    """從 TMDB URL 中解析出類型（movie, tv, person）和 ID。"""
    match = re.search(r'themoviedb\.org/(movie|tv|person)/(\d+)', url)
    if match:
        return match.group(1), match.group(2)
    return None, None

async def get_filmography_ids_api(session, semaphore, person_id, api_key):
    """
    【API】獲取一位工作人員的所有作品 ID。
    (使用者需求 2025-09-02) 新增錯誤重試機制，最多重試 100 次。
    """
    video_ids = set()
    url = f"{TMDB_API_BASE_URL}/person/{person_id}/combined_credits?api_key={api_key}"
    
    # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
    for attempt in range(101): # 嘗試 101 次 (初次 + 100 次重試)
        try:
            async with semaphore:
                response = await session.get(url, timeout=20)
                response.raise_for_status() # 如果是 4xx 或 5xx 錯誤，會拋出異常
                data = response.json()
                for item in data.get('cast', []) + data.get('crew', []):
                    if item.get('media_type') in ['movie', 'tv'] and 'id' in item:
                        video_ids.add((item['media_type'], item['id']))
                return video_ids # 成功後立即返回
        except (httpx.RequestError, httpx.HTTPStatusError) as e:
            if attempt < 100:
                wait_time = min(2 ** attempt, 60) # 指數退避，最多等 60 秒
                logging.warning(f"API 獲取人員 {person_id} 的作品時出錯 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"API 獲取人員 {person_id} 的作品時出錯，已達最大重試次數 (101 次): {e}")
        except Exception as e:
            # 處理非網路請求相關的其他意外錯誤
            logging.error(f"處理人員 {person_id} 的作品時發生非預期錯誤: {traceback.format_exc()}")
            break # 遇到非網路問題，可能重試也沒用，直接中斷

    return video_ids # 如果所有重試都失敗，返回空集合

async def get_person_ids_from_video_api(session, semaphore, video_type, video_id, api_key):
    """【API】從單一影片(電影或影集)中獲取所有相關人員的 ID。
    (使用者需求 2025-09-02) 新增錯誤重試機制，如果 API 請求失敗，最多重試 100 次。
    """
    person_ids = set()
    endpoint = "aggregate_credits" if video_type == 'tv' else "credits"
    url = f"{TMDB_API_BASE_URL}/{video_type}/{video_id}/{endpoint}?api_key={api_key}"
    context = f"{video_type}/{video_id}/{endpoint}"

    # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
    for attempt in range(101): # 嘗試 101 次 (初次 + 100 次重試)
        try:
            async with semaphore:
                response = await session.get(url, timeout=20)

                # 404 (Not Found) 是一種有效的失敗情況，不應重試
                if response.status_code == 404:
                    logging.warning(f"獲取 {context} 時項目不存在 (404)。")
                    return person_ids # 返回空集合

                response.raise_for_status()
                data = response.json()
                
                for member in data.get('cast', []) + data.get('crew', []):
                    if 'id' in member:
                        person_ids.add(member['id'])
                return person_ids # 成功，返回 ID 集合

        except (httpx.RequestError, httpx.HTTPStatusError) as e:
            if attempt < 100:
                wait_time = min(2 ** attempt, 60) # 指數退避，最多等 60 秒
                logging.warning(f"API 獲取 {context} 的工作人員時出錯 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"API 獲取 {context} 的工作人員時出錯，已達最大重試次數 (101 次): {e}")
        except Exception as e:
            # 處理非網路請求相關的其他意外錯誤
            logging.error(f"處理 {context} 時發生非預期錯誤: {traceback.format_exc()}")
            break # 遇到非網路問題，可能重試也沒用，直接中斷
    
    logging.error(f"API 獲取 {video_type}/{video_id} 的工作人員時最終失敗。")
    return person_ids

def contains_chinese(text):
    """輔助函數：檢查一個字串是否包含任何中文字元，並排除包含日文假名的字串。"""
    if not isinstance(text, str): return False
    if re.search(r'[\u3040-\u30ff]', text): return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

async def check_chinese_translation_api(session, semaphore, video_type, video_id, api_key, valid_urls_list):
    """
    【API】檢查指定的 TMDB 項目是否有中文內容。成功則將 URL 加入列表並返回 True，否則返回 False。
    (使用者需求 2025-08-28) 統一日誌格式，使用 ✔️ 和 ❌。
    (使用者需求 2025-09-02) 新增錯誤重試機制，最多重試 100 次。
    """
    url = f"{TMDB_API_BASE_URL}/{video_type}/{video_id}?api_key={api_key}&language=en-US&append_to_response=translations"

    # (使用者需求 2025-09-02) 如果 API 請求失敗，重試最多 100 次
    for attempt in range(101): # 嘗試 101 次 (初次 + 100 次重試)
        try:
            async with semaphore:
                response = await session.get(url, timeout=20)
                
                # 404 (Not Found) 是一種有效的失敗情況，不應重試
                if response.status_code == 404:
                    logging.warning(f"❌ 檢查 {video_type}/{video_id} 時項目不存在 (404)。")
                    return False

                response.raise_for_status()
                data = response.json()

                # 檢查官方翻譯
                for t in data.get('translations', {}).get('translations', []):
                    lang_code = t.get("iso_639_1")
                    country_code = t.get("iso_3166_1")
                    if lang_code == 'zh' and country_code in ['HK', 'TW', 'CN']:
                        logging.info(f"✔️  找到官方中文翻譯: {video_type}/{video_id} (代碼: {lang_code}-{country_code})")
                        valid_urls_list.append(f"{TMDB_WEB_BASE_URL}/{video_type}/{video_id}")
                        return True
                
                # 檢查標題
                title = data.get('title') or data.get('name')
                original_title = data.get('original_title') or data.get('original_name')
                title_to_log = ""
                is_chinese_title = False
                if title and contains_chinese(title):
                    is_chinese_title = True
                    title_to_log = title
                elif original_title and contains_chinese(original_title):
                    is_chinese_title = True
                    title_to_log = original_title
                    
                if is_chinese_title:
                    logging.info(f"✔️  在 en-US 資料中找到中文標題/原標題: {video_type}/{video_id} (標題: {title_to_log})")
                    valid_urls_list.append(f"{TMDB_WEB_BASE_URL}/{video_type}/{video_id}")
                    return True

                logging.warning(f"❌ 未找到任何中文內容: {video_type}/{video_id}。官方翻譯和 en-US 標題/原標題均不符合。")
                return False # 檢查成功，但無中文內容，返回 False 結束

        except (httpx.RequestError, httpx.HTTPStatusError) as e:
            if attempt < 100:
                wait_time = min(2 ** attempt, 60)
                logging.warning(f"檢查 {video_type}/{video_id} 時發生網路錯誤 (第 {attempt + 1}/101 次): {e}。將在 {wait_time} 秒後重試...")
                await asyncio.sleep(wait_time)
            else:
                logging.error(f"檢查 {video_type}/{video_id} 時發生網路錯誤，已達最大重試次數 (101 次): {e}")
        except Exception as e:
            logging.error(f"處理 {video_type}/{video_id} 時發生非預期錯誤: {traceback.format_exc()}")
            break # 遇到非網路問題，可能重試也沒用，直接中斷

    logging.warning(f"❌ 檢查 {video_type}/{video_id} 在多次重試後最終失敗。")
    return False

# =====================================================================================
#  主執行函數 (Main)
# =====================================================================================

async def async_main():
    """
    主執行函數，整合了網頁搜尋和 API 深度搜尋的流程。
    
    (使用者需求 2025-09-02) 在所有擴展階段新增進度條與預計完成時間顯示。
    (使用者需求 2025-09-02) 新增網頁搜尋快取機制。
    (使用者需求 2025-09-02) 恢復多行輸入功能，以利貼上演員名單，並以 'done' 或 'search' 關鍵字結束輸入。
    決策流程圖 (Mermaid):
    graph TD
        A[開始] --> B{載入設定檔};
        B --> C{獲取使用者輸入};
        C --> D{使用者是否輸入了演員名單?};
        D -- 是 --> D1[接收多行輸入直到 'done'];
        D1 --> E[啟動/附加瀏覽器];
        E --> F{登入 TMDB};
        F -- 成功 --> G[迴圈: 網頁搜尋每個演員];
        G --> H{收集所有演員的個人頁面 URL};
        H --> I[將找到的 URL 寫入快取檔案 actors_cache.txt];
        I --> J[關閉瀏覽器視窗/實例];
        J --> K{開始 API 深度搜尋};
        D -- 否 (直接按 Enter) --> L[讀取快取檔案 actors_cache.txt];
        L --> M{是否成功讀取 URL?};
        M -- 是 --> K;
        M -- 否 --> Z[結束];
        K --> N[擴展階段1: 透過初始人員 URL 獲取其所有作品];
        N --> O[擴展階段2: 透過作品獲取所有相關人員];
        O --> P[擴展階段3: 透過所有相關人員獲取其所有作品];
        P --> Q[最終階段: 檢查所有作品是否有中文內容];
        Q --> R{顯示進度與預計完成時間};
        R --> S[收集所有含中文內容的 URL];
        S --> T{是否有找到結果?};
        T -- 是 --> U[將結果寫入 Excel];
        T -- 否 --> V[提示未找到結果];
        U --> Z;
        V --> Z;
        F -- 失敗 --> Z;
    """
    
    # --- 步驟 1: 獲取使用者輸入和設定 ---
    configs = load_configs()
    tmdb_user = configs.get("tmdb_username")
    tmdb_pass = configs.get("tmdb_password")
    api_key = configs.get("tmdb_api_v3")

    if not all([tmdb_user, tmdb_pass, api_key]):
        logging.error("configs.json 中缺少 tmdb_username, tmdb_password 或 tmdb_api_v3。")
        print("錯誤：請確保 configs.json 檔案中包含 TMDB 的使用者名稱、密碼和 API v3 金鑰。")
        return

    ACTOR_URL_CACHE_FILE = os.path.join(os.path.dirname(__file__), '../data/actors_cache.txt')

    # (使用者需求 2025-09-02) 恢復多行輸入機制
    print("請貼上或逐行輸入演員名稱列表，完成後輸入 'done' 或 'search' 開始搜尋。")
    print("或者，直接按 Enter 使用上次搜尋的快取:")
    
    lines = []
    actor_names = []
    
    try:
        # 讀取第一行輸入以判斷使用者意圖
        first_line = input()
        if first_line.strip() == "":
            # 如果第一行為空，代表使用者希望使用快取，actor_names 保持為空
            pass
        else:
            # 如果第一行有內容，則開始接收多行輸入
            lines.append(first_line)
            while True:
                line = input()
                if line.strip().lower() in ['search', 'done']:
                    break
                lines.append(line)
    except EOFError:
        pass # 處理管線輸入或檔案重定向結束的情況

    # 如果 lines 列表不為空，表示使用者輸入了演員名稱，進行處理
    if lines:
        full_input = "\n".join(lines)
        # 使用正規表示式分割，處理逗號、空格、換行等多種分隔符
        actor_names = [name.strip() for name in re.split(r'[,，\s\n]+', full_input) if name.strip()]

    initial_person_urls = []

    if actor_names:
        # --- 路徑 A: 使用者輸入了演員名稱，執行網頁搜尋 ---
        logging.info("偵測到演員名稱輸入，將執行網頁搜尋。")
        driver, profile_path, is_standalone = None, None, False
        try:
            driver, profile_path, is_standalone = get_browser_driver(remote_port=9223)
            if not login_to_tmdb(driver, tmdb_user, tmdb_pass):
                return 

            logging.info(f"--- 開始網頁搜尋 {len(actor_names)} 位演員 ---")
            for name in actor_names:
                urls = search_person_url_web(driver, name)
                if urls:
                    initial_person_urls.extend(urls)
            
            logging.info(f"網頁搜尋完成，共找到 {len(initial_person_urls)} 個有效的演員頁面 URL。")

        finally:
            if driver:
                cleanup_browser(driver, profile_path, is_standalone)

        # 將網頁搜尋結果寫入快取檔案
        if initial_person_urls:
            logging.info(f"將 {len(initial_person_urls)} 個演員 URL 寫入快取檔案...")
            try:
                output_dir = os.path.dirname(ACTOR_URL_CACHE_FILE)
                os.makedirs(output_dir, exist_ok=True)
                with open(ACTOR_URL_CACHE_FILE, 'w', encoding='utf-8') as f:
                    for url in set(initial_person_urls): # 寫入時去重
                        f.write(f"{url}\n")
                logging.info(f"快取已儲存至 {ACTOR_URL_CACHE_FILE}")
            except Exception as e:
                logging.error(f"寫入快取檔案時發生錯誤: {e}")
        else:
             logging.warning("網頁搜尋未找到任何演員 URL，快取未更新。")
    
    else:
        # --- 路徑 B: 使用者直接按 Enter，從快取讀取 ---
        logging.info("使用者未輸入演員名稱，嘗試從快取檔案讀取...")
        try:
            if os.path.exists(ACTOR_URL_CACHE_FILE):
                with open(ACTOR_URL_CACHE_FILE, 'r', encoding='utf-8') as f:
                    initial_person_urls = [line.strip() for line in f if line.strip()]
                if initial_person_urls:
                    logging.info(f"從快取檔案成功讀取 {len(initial_person_urls)} 個演員 URL。")
                else:
                    logging.warning("快取檔案為空，無事可做。")
                    return
            else:
                logging.warning("找不到快取檔案，請先輸入演員名稱執行一次網頁搜尋以生成快取。")
                return
        except Exception as e:
            logging.error(f"讀取快取檔案時發生錯誤: {e}")
            return

    if not initial_person_urls:
        logging.warning("無有效的演員 URL 可供處理，程式結束。")
        return

    # --- API 擴展階段 (共通邏輯) ---
    logging.info("--- 開始 API 深度搜尋 ---")
    initial_person_ids = set()
    for url in set(initial_person_urls): # 使用 set 去除重複 URL
        item_type, item_id = parse_tmdb_url(url)
        if item_type == 'person' and item_id:
            initial_person_ids.add(item_id)

    semaphore = asyncio.Semaphore(MAX_CONCURRENT_REQUESTS)
    async with httpx.AsyncClient(timeout=30.0) as session:
        # 階段 A: 從初始人員獲取他們的所有作品，作為「種子影片」
        filmography_tasks = [get_filmography_ids_api(session, semaphore, p_id, api_key) for p_id in initial_person_ids]
        
        total_tasks = len(filmography_tasks)
        if total_tasks > 0:
            processed_tasks = 0
            start_time = time.time()
            filmography_results = []
            logging.info(f"擴展階段 1: 從 {total_tasks} 位初始人員抓取其所有作品...")

            for future in asyncio.as_completed(filmography_tasks):
                result = await future
                if result:
                    filmography_results.append(result)
                processed_tasks += 1

                elapsed_time = time.time() - start_time
                items_per_second = processed_tasks / elapsed_time if elapsed_time > 0 else 0
                remaining_items = total_tasks - processed_tasks
                
                if items_per_second > 0:
                    etc_seconds = remaining_items / items_per_second
                    etc_str = f"預計剩餘 {int(etc_seconds // 60)} 分 {int(etc_seconds % 60)} 秒"
                else:
                    etc_str = "預計剩餘時間計算中..."

                progress_bar = f"進度: {processed_tasks}/{total_tasks}"
                sys.stdout.write(f"\r{progress_bar} - {etc_str}".ljust(80))
                sys.stdout.flush()

            print()
            seed_video_ids = set.union(*filmography_results) if filmography_results else set()
            logging.info(f"找到 {len(seed_video_ids)} 部種子影片。")
        else:
            seed_video_ids = set()
            logging.info("擴展階段 1: 無初始人員，跳過。")


        # 階段 B: 從「種子影片」中，抓取所有相關的演職員
        personnel_tasks = [get_person_ids_from_video_api(session, semaphore, v_type, v_id, api_key) for v_type, v_id in seed_video_ids]
        
        total_tasks = len(personnel_tasks)
        if total_tasks > 0:
            processed_tasks = 0
            start_time = time.time()
            personnel_results = []
            logging.info(f"擴展階段 2: 從 {total_tasks} 部種子影片中抓取所有相關人員...")

            for future in asyncio.as_completed(personnel_tasks):
                result = await future
                if result:
                    personnel_results.append(result)
                processed_tasks += 1

                elapsed_time = time.time() - start_time
                items_per_second = processed_tasks / elapsed_time if elapsed_time > 0 else 0
                remaining_items = total_tasks - processed_tasks
                
                if items_per_second > 0:
                    etc_seconds = remaining_items / items_per_second
                    etc_str = f"預計剩餘 {int(etc_seconds // 60)} 分 {int(etc_seconds % 60)} 秒"
                else:
                    etc_str = "預計剩餘時間計算中..."

                progress_bar = f"進度: {processed_tasks}/{total_tasks}"
                sys.stdout.write(f"\r{progress_bar} - {etc_str}".ljust(80))
                sys.stdout.flush()

            print()
            all_person_ids = set.union(*personnel_results) if personnel_results else set()
        else:
            all_person_ids = set()
            logging.info("擴展階段 2: 無種子影片，跳過。")
            
        all_person_ids.update(initial_person_ids) # 確保初始人員也被包含
        logging.info(f"總共找到 {len(all_person_ids)} 位獨特的工作人員。")

        # 階段 C: 從所有找到的演職員中，再次抓取他們的所有作品
        final_filmography_tasks = [get_filmography_ids_api(session, semaphore, p_id, api_key) for p_id in all_person_ids]
        
        total_tasks = len(final_filmography_tasks)
        if total_tasks > 0:
            processed_tasks = 0
            start_time = time.time()
            final_filmography_results = []
            logging.info(f"擴展階段 3: 從 {total_tasks} 位工作人員抓取所有相關影片...")

            for future in asyncio.as_completed(final_filmography_tasks):
                result = await future
                if result:
                    final_filmography_results.append(result)
                processed_tasks += 1

                elapsed_time = time.time() - start_time
                items_per_second = processed_tasks / elapsed_time if elapsed_time > 0 else 0
                remaining_items = total_tasks - processed_tasks
                
                if items_per_second > 0:
                    etc_seconds = remaining_items / items_per_second
                    etc_str = f"預計剩餘 {int(etc_seconds // 60)} 分 {int(etc_seconds % 60)} 秒"
                else:
                    etc_str = "預計剩餘時間計算中..."

                progress_bar = f"進度: {processed_tasks}/{total_tasks}"
                sys.stdout.write(f"\r{progress_bar} - {etc_str}".ljust(80))
                sys.stdout.flush()
                
            print()
            all_video_ids_to_check = set.union(*final_filmography_results) if final_filmography_results else set()
        else:
            all_video_ids_to_check = set()
            logging.info("擴展階段 3: 無工作人員可供搜尋，跳過。")

        all_video_ids_to_check.update(seed_video_ids) # 確保種子影片也被包含
        logging.info(f"總共找到 {len(all_video_ids_to_check)} 個獨特的影片，準備進行翻譯檢查。")

        # 階段 D: 檢查所有找到的影片是否有中文內容
        logging.info("--- 最終階段：並行檢查所有影片的中文翻譯 ---")
        final_urls = []
        translation_tasks = [check_chinese_translation_api(session, semaphore, v_type, v_id, api_key, final_urls) for v_type, v_id in all_video_ids_to_check]
        
        total_tasks = len(translation_tasks)
        if total_tasks > 0:
            processed_tasks = 0
            start_time = time.time() 
            for future in asyncio.as_completed(translation_tasks):
                await future 
                processed_tasks += 1
                
                elapsed_time = time.time() - start_time
                items_per_second = processed_tasks / elapsed_time if elapsed_time > 0 else 0
                remaining_items = total_tasks - processed_tasks
                
                if items_per_second > 0:
                    estimated_remaining_seconds = remaining_items / items_per_second
                    remaining_minutes = int(estimated_remaining_seconds // 60)
                    remaining_seconds_val = int(estimated_remaining_seconds % 60)
                    etc_str = f"預計剩餘 {remaining_minutes} 分 {remaining_seconds_val} 秒"
                else:
                    etc_str = "預計剩餘時間計算中..."
                    
                progress_bar = f"檢查進度: {processed_tasks}/{total_tasks}"
                found_count = f"已找到 {len(final_urls)} 個"
                
                sys.stdout.write(f"\r{progress_bar} ({found_count}) - {etc_str}".ljust(80))
                sys.stdout.flush()

            print() # 完成後換行
        
        logging.info(f"檢查完成。共有 {len(final_urls)} 個影片包含中文內容。")

    # --- 步驟 5: 將結果寫入 Excel ---
    if final_urls:
        output_file = os.path.join(os.path.dirname(__file__), '../data/bulk_videos.xlsx')
        append_to_bulk_excel(sorted(list(set(final_urls))), file_name=output_file)
        absolute_path = os.path.abspath(output_file)
        print(f"\n已將 {len(final_urls)} 個結果加入到 {absolute_path} 的 Bulk 工作表中。")
    else:
        print("\n未找到任何包含中文內容的關聯影片。")

    logging.info("程式執行完畢。")

if __name__ == "__main__":
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(async_main())
