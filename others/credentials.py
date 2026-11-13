import os
import json
import time  # 確保已導入 time 模組
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException

_tmdb_confirmed_login_status_session = False

def load_configs():
    config_path = os.path.join(os.path.dirname(__file__), '../configs.json')
    if os.path.exists(config_path):
        with open(config_path, 'r') as config_file:
            return json.load(config_file)
    return {}

def save_configs(configs):
    config_path = os.path.join(os.path.dirname(__file__), '../configs.json')
    with open(config_path, 'w') as config_file:
        json.dump(configs, config_file)

def get_tmdb_credentials(configs):
    if 'tmdb_username' not in configs or 'tmdb_password' not in configs:
        tmdb_username = input("請輸入 TMDB 的 username: ")
        tmdb_password = input("請輸入 TMDB 的 password: ")
        save_config = input("是否希望儲存這些資料留待下次使用? (y/n): ")
        if save_config.lower() == 'y':
            configs['tmdb_username'] = tmdb_username
            configs['tmdb_password'] = tmdb_password
            save_configs(configs)
    else:
        tmdb_username = configs['tmdb_username']
        tmdb_password = configs['tmdb_password']
    return tmdb_username, tmdb_password

def get_disneyplus_credentials(configs):
    if 'disneyplus_email' not in configs or 'disneyplus_password' not in configs:
        disneyplus_email = input("請輸入 Disney+ 的 email: ")
        disneyplus_password = input("請輸入 Disney+ 的 password: ")
        save_config = input("是否希望儲存這些資料留待下次使用? (y/n): ")
        if save_config.lower() == 'y':
            configs['disneyplus_email'] = disneyplus_email
            configs['disneyplus_password'] = disneyplus_password
            save_configs(configs)
    else:
        disneyplus_email = configs['disneyplus_email']
        disneyplus_password = configs['disneyplus_password']
    return disneyplus_email, disneyplus_password

def is_tmdb_logged_in(driver):
    """
    檢查 TMDB 是否已登入。
    如果已透過快取確認登入，則直接返回 True。
    否則，訪問登入頁面並檢查是否重定向到用戶頁面。
    """
    global _tmdb_confirmed_login_status_session
    if _tmdb_confirmed_login_status_session:
        logging.info("已從快取確認登入狀態。")
        return True

    try:
        # 為了避免在已知登入狀態下依然載入 profile 頁，
        # 這裡的 get 操作是必要的，用以確認當前實際的登入狀態（如果快取未命中）
        # 但如果_tmdb_confirmed_login_status_session 為 True，則不會執行到這裡。
        logging.info("正在透過網頁檢查 TMDB 登入狀態...")
        driver.get("https://www.themoviedb.org/login") # 導航到登入頁
        # 在實際應用中，根據頁面載入速度和元素出現時機，可能需要更長的 time.sleep 或 WebDriverWait
        time.sleep(2) # 等待頁面可能的重定向
        
        current_url = driver.current_url
        if '/u/' in current_url:  # 如果 URL 包含 /u/，表示已登入並跳轉到使用者頁面
            logging.info("已檢測到登入狀態 (透過網頁驗證)。")
            _tmdb_confirmed_login_status_session = True # 更新快取
            return True
        else:
            # 檢查是否仍在登入頁面或發生其他情況
            # 例如，可以檢查頁面上是否有登入表單元素來進一步確認
            logging.info("未檢測到登入狀態 (當前 URL: %s)，需要登入。", current_url)
            return False
    except Exception as e:
        logging.error(f"檢查登入狀態時發生錯誤: {e}")
        return False

def login(driver, username, password):
    """
    登入到 TMDB。
    如果登入成功，會更新快取的登入狀態。
    """
    global _tmdb_confirmed_login_status_session
    logging.info("嘗試登入 TMDB。")
    login_url = "https://www.themoviedb.org/login" # 修正變數名稱
    try:
        driver.get(login_url) # 使用修正後的變數
        logging.info("導航到登入頁面。")
        time.sleep(2) # 等待頁面載入
        
        # 接受 cookies (如果存在)
        try:
            accept_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
            )
            accept_button.click()
            time.sleep(1)
            logging.info("接受 Cookie。")
        except Exception:
            logging.info("無需接受 Cookie 或 Cookie 按鈕未找到/已接受。")
        
        # 填寫登入信息
        username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "username")))
        password_field = driver.find_element(By.NAME, "password")
        
        username_field.send_keys(username)
        password_field.send_keys(password)
        logging.info("已填寫用戶名和密碼。")
        password_field.send_keys(Keys.RETURN) # 提交表單
        logging.info("提交登入信息。")
        
        # 等待登入完成後的跳轉或頁面變化
        # is_tmdb_logged_in 會被調用來確認，並且它會更新 _tmdb_confirmed_login_status_session
        # 增加一個短暫等待確保頁面有足夠時間反應
        time.sleep(5) 

        if is_tmdb_logged_in(driver): # 這裡會使用到快取機制（如果是第二次調用）或更新快取
            logging.info("成功登入到 TMDB。")
            # _tmdb_confirmed_login_status_session 已在 is_tmdb_logged_in 中設定
        else:
            logging.error("登入失敗，請檢查用戶名和密碼或頁面狀態。")
            # 如果登入失敗，之前的 _tmdb_confirmed_login_status_session 應該是 False 或保持 False
            _tmdb_confirmed_login_status_session = False # 明確設置為 False

    except Exception as e:
        logging.error(f"登入過程中出錯: {e}")
        _tmdb_confirmed_login_status_session = False # 登入異常，狀態視為未登入
        # 原有的 login_failsafe 邏輯可以保留，但它也應該考慮更新全域狀態
        # 為簡化，此處未包含 login_failsafe 的修改，但應注意其對狀態的影響
        current_page_source = driver.page_source # 獲取當前頁面源碼用於調試
        if "Wait and try again" in current_page_source: # 檢查特定錯誤訊息
            logging.warning("TMDB 登入頁面提示 'Wait and try again'。")

def login_failsafe(driver, current_url, username, password):
    try:
        logging.info(f"進入 login failsafe 模式的 URL: {current_url}")
        if "Wait and try again" in driver.page_source:
            logging.warning("檢測到 'Wait and try again' 訊息，等待 1 分鐘後重試...")
            time.sleep(60)
            logging.info("重新加載頁面...")
            driver.get(current_url)
            login(driver, username, password)
        else:
            logging.error("未知錯誤，請檢查。")
    except Exception as e:
        logging.error(f"login failsafe 過程中發生錯誤: {e}")

# ─── Netflix ──────────────────────────────────────────────────────────────────

_netflix_confirmed_login_status_session = False

def get_netflix_credentials(configs):
    if 'netflix_username' not in configs or 'netflix_password' not in configs:
        netflix_username = input("請輸入 Netflix 的 email 或手機號碼: ")
        netflix_password = input("請輸入 Netflix 的 password (如使用 OTP 登入可留空): ")
        save_config = input("是否希望儲存這些資料留待下次使用? (y/n): ")
        if save_config.lower() == 'y':
            configs['netflix_username'] = netflix_username
            configs['netflix_password'] = netflix_password
            save_configs(configs)
    else:
        netflix_username = configs['netflix_username']
        netflix_password = configs['netflix_password']
    return netflix_username, netflix_password

def _select_first_netflix_profile(driver):
    """若當前頁面出現帳戶選擇畫面，點選第一個可用帳戶。"""
    try:
        profile_container = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.list-profiles'))
        )
        profile_links = profile_container.find_elements(By.CSS_SELECTOR, 'a.profile-link[href]')
        if profile_links:
            profile_links[0].click()
            logging.info("已選擇第一個可用的 Netflix 帳戶。")
            time.sleep(3)
            return True
    except (TimeoutException, NoSuchElementException):
        pass
    except Exception as e:
        logging.info("帳戶選擇略過: %s", e)
    return False

def is_netflix_logged_in(driver):
    global _netflix_confirmed_login_status_session
    if _netflix_confirmed_login_status_session:
        logging.info("已從快取確認 Netflix 登入狀態。")
        return True
    try:
        logging.info("正在透過網頁檢查 Netflix 登入狀態...")
        driver.get("https://www.netflix.com/browse")
        time.sleep(3)
        # 帳戶選擇畫面可能阻擋瀏覽器停留在 /browse，先行處理
        _select_first_netflix_profile(driver)
        time.sleep(2)
        current_url = driver.current_url
        if 'netflix.com/browse' in current_url:
            logging.info("Netflix 已登入。")
            _netflix_confirmed_login_status_session = True
            return True
        logging.info("Netflix 未登入，當前 URL: %s", current_url)
        return False
    except Exception as e:
        logging.error(f"檢查 Netflix 登入狀態時發生錯誤: {e}")
        return False

def login_netflix(driver, username, password):
    global _netflix_confirmed_login_status_session
    logging.info("嘗試登入 Netflix。")
    try:
        driver.get("https://www.netflix.com/login")
        wait = WebDriverWait(driver, 20)

        # 填寫帳號 (email 或手機號碼)
        username_field = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'input[name="userLoginId"]')
        ))
        username_field.clear()
        username_field.send_keys(username)
        logging.info("已填寫 Netflix 帳號。")

        # 點擊 Continue 按鈕
        continue_button = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'button[data-uia="continue-button"]')
        ))
        continue_button.click()
        logging.info("已點擊 Continue，等待下一步...")
        time.sleep(3)

        # 若出現密碼欄位則填寫密碼
        password_fields = driver.find_elements(By.CSS_SELECTOR, 'input[type="password"]')
        if password_fields and password:
            password_fields[0].send_keys(password)
            logging.info("已填寫 Netflix 密碼。")
            sign_in_buttons = driver.find_elements(By.CSS_SELECTOR,
                'button[data-uia="sign-in-button"], button[type="submit"]')
            if sign_in_buttons:
                sign_in_buttons[0].click()
                logging.info("已點擊登入按鈕。")
                time.sleep(3)

        # 檢查是否出現 OTP 驗證碼頁面
        try:
            otp_title_elements = driver.find_elements(By.CSS_SELECTOR, '[data-uia="title"]')
            is_otp_page = any(
                "code" in el.text.lower() or "驗證" in el.text or "verify" in el.text.lower()
                for el in otp_title_elements if el.text
            )
            if is_otp_page:
                logging.info("Netflix 要求輸入驗證碼。")
                print("\nNetflix 已發送 4 位驗證碼到您的 email，請查收。")
                otp_code = input("請輸入 4 位驗證碼: ").strip()

                # 優先嘗試 4 個獨立的單字元輸入框
                digit_inputs = driver.find_elements(By.CSS_SELECTOR, 'input[maxlength="1"]')
                if len(digit_inputs) >= 4:
                    for i, digit in enumerate(otp_code[:4]):
                        digit_inputs[i].clear()
                        digit_inputs[i].send_keys(digit)
                        time.sleep(0.3)
                    logging.info("已分別輸入 4 位驗證碼。")
                else:
                    # 嘗試單一輸入框
                    single_inputs = driver.find_elements(By.CSS_SELECTOR,
                        'input[name="passcode"], input[data-uia*="code"], input[type="tel"]')
                    if not single_inputs:
                        single_inputs = driver.find_elements(By.CSS_SELECTOR, 'input[type="text"]')
                    if single_inputs:
                        single_inputs[0].clear()
                        single_inputs[0].send_keys(otp_code)
                        single_inputs[0].send_keys(Keys.RETURN)
                        logging.info("已輸入驗證碼（單一欄位）。")
                    else:
                        logging.error("找不到驗證碼輸入欄位。")
                        return
                time.sleep(5)
        except Exception as e:
            logging.info("OTP 頁面檢查略過: %s", e)

        # 若出現帳戶選擇頁面，點選第一個有連結的帳戶
        time.sleep(2)
        if not _select_first_netflix_profile(driver):
            logging.info("未出現帳戶選擇頁面，繼續。")

        # 確認登入結果
        current_url = driver.current_url
        if 'netflix.com' in current_url and 'login' not in current_url:
            logging.info("Netflix 登入成功，當前 URL: %s", current_url)
            _netflix_confirmed_login_status_session = True
        else:
            logging.warning("Netflix 登入後頁面狀態未知，當前 URL: %s", current_url)
            _netflix_confirmed_login_status_session = True

    except Exception as e:
        logging.error(f"Netflix 登入過程中出錯: {e}")
        _netflix_confirmed_login_status_session = False

def ensure_netflix_logged_in(driver):
    """確保 Netflix 已登入，否則自動執行登入流程。"""
    if is_netflix_logged_in(driver):
        return
    configs = load_configs()
    username, password = get_netflix_credentials(configs)
    login_netflix(driver, username, password)
