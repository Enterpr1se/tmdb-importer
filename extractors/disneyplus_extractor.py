import time
import pandas as pd
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from others.credentials import load_configs, get_disneyplus_credentials
from others.utils import add_spaces_between_en_zh_num, convert_runtime_to_minutes

is_logged_in = False  # 全域變數，用於追踪是否已經登入過
profile_selected = False  # 全域變數，用於追踪是否已經選擇過 profile

def is_disneyplus_logged_in(driver):
    global is_logged_in, profile_selected
    try:
        logging.info("檢查是否已登入 Disney+")
        driver.get("https://www.disneyplus.com")
        # 等待頁面完全載入
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        # 優先檢查 Profile 清單
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'ul[data-testid="account-dropdown-list"]'))
            )
            logging.info("檢測到 Profile 清單，確認已登入 Disney+")
            is_logged_in = True
            profile_selected = True  # 已登入，無需選擇 Profile
            return True
        except TimeoutException:
            # 次優先檢查 Profile 頭像
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="button"][data-testid^="profile-avatar-"]'))
                )
                logging.info("檢測到 Profile 頭像，確認已登入 Disney+")
                is_logged_in = True
                profile_selected = True  # 已登入，無需選擇 Profile
                return True
            except TimeoutException:
                # 檢查當前 URL 是否為 Profile 選擇頁面
                if "disneyplus.com/profiles" in driver.current_url:
                    logging.info("當前 URL 為 Profile 選擇頁面，確認已登入 Disney+")
                    is_logged_in = True
                    profile_selected = False  # 需要選擇 Profile
                    return True
                # 最後檢查登入按鈕
                login_buttons = driver.find_elements(By.CSS_SELECTOR, "a[data-testid='login-header'], a[data-testid='log_in']")
                if not login_buttons:
                    login_buttons = driver.find_elements(By.CSS_SELECTOR, "a[href='/login'], a[href='/identity/login']")
                if login_buttons:
                    logging.info("檢測到登入按鈕，確認未登入 Disney+")
                    is_logged_in = False
                    profile_selected = False
                    return False
                logging.warning("未找到 Profile 清單、頭像或登入按鈕，假設未登入")
                is_logged_in = False
                profile_selected = False
                return False
    except Exception as e:
        logging.error(f"檢查 Disney+ 登錄狀態過程中出錯: {e}")
        is_logged_in = False
        profile_selected = False
        return False

def login_to_disneyplus(driver, email, password):
    global is_logged_in, profile_selected
    try:
        logging.info("開始登入 Disney+")
        driver.get("https://www.disneyplus.com")
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        # 嘗試點擊登入按鈕
        login_button = None
        try:
            login_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[data-testid='login-header'], a[data-testid='log_in']"))
            )
        except TimeoutException:
            logging.warning("未找到 data-testid 為 login-header 或 log_in 的按鈕，嘗試 href 定位")
            login_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/login'], a[href='/identity/login']"))
            )
        login_button.click()
        logging.info("已點擊登入按鈕")
        
        # 輸入電子郵件
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.NAME, "email"))
        ).send_keys(email)
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
        logging.info("已輸入電子郵件，並點擊提交按鈕")

        # 處理密碼或驗證碼
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@name='password'] | //div[@class='passcode-key']"))
            )
            password_field = driver.find_elements(By.XPATH, "//input[@name='password']")
            passcode_field = driver.find_elements(By.CLASS_NAME, "passcode-key")

            if password_field:
                logging.info("檢測到需要輸入密碼")
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.NAME, "password"))
                ).send_keys(password)
                driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
                logging.info("已輸入密碼，並點擊提交按鈕")
            elif passcode_field:
                logging.info("檢測到需要輸入 6 位數字碼")
                while True:
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), \"Who's watching?\")]"))
                        )
                        break
                    except TimeoutException:
                        logging.info("等待 6 位數字碼驗證...")
                        time.sleep(5)
        except TimeoutException:
            logging.warning("未檢測到密碼或驗證碼輸入框，可能已自動登入")

        # 等待 Profile 選擇頁面
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="button"][data-testid^="profile-avatar-"]'))
        )
        profiles = driver.find_elements(By.CSS_SELECTOR, 'div[role="button"][data-testid^="profile-avatar-"]')
        profiles[0].click()
        logging.info("已選擇第一個 Profile")
        time.sleep(3)

        is_logged_in = True
        profile_selected = True
    except Exception as e:
        logging.error(f"Disney+ 登入失敗: {e}")
        failsafe(driver, driver.current_url, login_to_disneyplus, driver, email, password)

def click_details_button(driver, retries=5):
    for attempt in range(retries):
        try:
            logging.info("開始尋找「簡介」按鈕")
            details_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'li[data-testid="details-page-tab"][aria-controls="details"]'))
            )
            logging.info("已找到「簡介」按鈕並現在點擊")
            details_button.click()
            logging.info("已點擊「簡介」按鈕")
            time.sleep(2)  # 等待 2 秒以確保頁面完全加載
            return
        except TimeoutException:
            logging.warning(f"找不到「簡介」按鈕，重試第 {attempt + 1} 次...")
            time.sleep(2)
    
    logging.error("多次重試後仍無法找到「簡介」按鈕，交由 failsafe 處理。")
    failsafe(driver, driver.current_url, click_details_button, driver, retries)

def click_episodes_button(driver):
    try:
        logging.info("開始尋找「集數」按鈕")
        episodes_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'li[data-testid="details-page-tab"][aria-controls="episodes"]'))
        )
        episodes_button.click()
        logging.info("已點擊「集數」按鈕")
        time.sleep(2)  # 等待 2 秒以確保頁面完全加載
    except Exception as e:
        logging.error(f"Error clicking episodes button: {e}")
        failsafe(driver, driver.current_url, click_episodes_button, driver)

def scroll_to_bottom(driver):
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

def scroll_to_top(driver):
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)

def click_using_js(driver, element):
    driver.execute_script("arguments[0].click();", element)

def extract_runtime(runtime_text):
    """將運行時間文本轉換為分鐘，使用 utils.convert_runtime_to_minutes"""
    try:
        result = convert_runtime_to_minutes(runtime_text)
        if result is not None:
            logging.info(f"解析運行時間: {runtime_text} -> {result} 分鐘")
        else:
            logging.warning(f"無法解析運行時間格式: {runtime_text}")
        return result
    except Exception as e:
        logging.error(f"解析運行時間時發生錯誤: {e}, 原始文本: {runtime_text}")
        return None

def extract_season_info(driver):
    logging.info("開始抓取季數資訊")
    seasons = []
    all_episodes = []
    season_names = set()

    try:
        # 嘗試尋找並點擊多季的下拉選單
        dropdown_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="dropdown-button"]'))
        )
        click_using_js(driver, dropdown_button)
        logging.info("已點擊下拉菜單按鈕")
        time.sleep(2)

        season_dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="dropdown-list"]'))
        )
        season_elements = season_dropdown.find_elements(By.TAG_NAME, 'li')
        logging.info(f"找到多季，共有 {len(season_elements)} 季")
        for i in range(len(season_elements)):
            season_dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="dropdown-list"]'))
            )
            season_elements = season_dropdown.find_elements(By.TAG_NAME, 'li')
            season_name_text = season_elements[i].text
            if not season_name_text: continue # 跳過空的列表項
            
            season_name = add_spaces_between_en_zh_num(season_name_text)
            if season_name in season_names:
                continue
            season_names.add(season_name)
            
            season_num_match = re.search(r'\d+', season_name)
            current_season_number = int(season_num_match.group()) if season_num_match else (i + 1)
            
            try:
                season_description_element = driver.find_element(By.CSS_SELECTOR, '[data-testid="details-featured-description"]')
                season_description = add_spaces_between_en_zh_num(season_description_element.text)
            except NoSuchElementException:
                season_description = ''
            seasons.append([season_name, current_season_number, season_description])
            season_elements[i].click()
            logging.info(f"已選擇第 {current_season_number} 季")
            time.sleep(2) 
            grab_episodes(driver, current_season_number, all_episodes)
            scroll_to_top(driver) 
            dropdown_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="dropdown-button"]'))
            )
            click_using_js(driver, dropdown_button)
            time.sleep(1)
            
    except (NoSuchElementException, TimeoutException) as e:
        logging.info(f"找不到多季資訊，檢查是否為單一季: {type(e).__name__}")
        # 如果找不到多季選單，則嘗試處理單一季的情況
        season_description = ''
        try:
            season_description_element = driver.find_element(By.CSS_SELECTOR, '[data-testid="details-featured-description"]')
            season_description = add_spaces_between_en_zh_num(season_description_element.text)
        except NoSuchElementException:
            pass # 描述是可選的

        try:
            # 情況 1: 找到一個「已停用」的下拉選單，代表只有一季
            disabled_dropdown = driver.find_element(By.CSS_SELECTOR, '[data-testid="dropdown-button"][disabled]')
            season_name_text = disabled_dropdown.find_element(By.TAG_NAME, 'span').text
            season_name = add_spaces_between_en_zh_num(season_name_text) if season_name_text else "第 1 季"
            season_num_match = re.search(r'\d+', season_name)
            current_season_number = int(season_num_match.group()) if season_num_match else 1
            logging.info(f"找到單一季數資訊 (來自禁用下拉選單): 季數 = {current_season_number}")
            seasons.append([season_name, current_season_number, season_description])
            grab_episodes(driver, current_season_number, all_episodes)
        except NoSuchElementException:
            # 情況 2: 連「已停用」的下拉選單都找不到，代表是迷你劇等沒有季數選單的內容
            logging.info("也找不到禁用的下拉選單。假設為單一季，直接抓取頁面上的集數。")
            season_name = "第 1 季"
            current_season_number = 1
            seasons.append([season_name, current_season_number, season_description])
            grab_episodes(driver, current_season_number, all_episodes)

    return seasons, all_episodes

def grab_episodes(driver, season_number, all_episodes):
    try:
        logging.info(f"開始抓取第 {season_number} 季的集數")
        scroll_to_bottom(driver)
        episodes = driver.find_elements(By.CSS_SELECTOR, '[data-testid="set-item"]')
        episode_numbers = set()
        for episode in episodes:
            try:
                # 嘗試在每個項目中尋找標題和描述
                episode_title_element = episode.find_element(By.CSS_SELECTOR, '[data-testid="standard-regular-list-item-title"]')
                episode_title = add_spaces_between_en_zh_num(episode_title_element.text)

                # 如果標題為空，可能不是一個有效的集數項目，跳過
                if not episode_title.strip():
                    continue

                episode_description = add_spaces_between_en_zh_num(episode.find_element(By.CSS_SELECTOR, '[data-testid="standard-regular-list-item-description"]').text)
                
                try:
                    # 運行時間是可選的
                    episode_runtime_element = episode.find_element(By.CSS_SELECTOR, 'span._1jf071b0[aria-hidden="true"]').text
                    episode_runtime = extract_runtime(episode_runtime_element)
                except NoSuchElementException:
                    episode_runtime = None

                try:
                    episode_number = int(episode_title.split('.')[0])
                except (ValueError, IndexError):
                    logging.warning(f"無法從標題解析集數編號: {episode_title}")
                    episode_number = None

                if episode_number and episode_number in episode_numbers:
                    continue
                
                if episode_number:
                    episode_numbers.add(episode_number)
                    episode_title_parts = episode_title.split('. ', 1)
                    if len(episode_title_parts) > 1:
                        episode_title = add_spaces_between_en_zh_num(episode_title_parts[1].strip())

                    unwanted_texts = [
                        '部分閃光片段或圖案可能會影響對光敏感的觀眾。',
                        '一些閃爍的燈光片段或圖案可能會影響有光敏的觀眾'
                    ]
                    for unwanted_text in unwanted_texts:
                        if unwanted_text in episode_description:
                            episode_description = episode_description.replace(unwanted_text, '').strip()

                    if not episode_title and not episode_description and episode_runtime is None:
                        logging.info("標題、描述和時長都為空，跳過此集數。")
                        continue

                    logging.info(f"抓取到集數: 季數={season_number}, 集數={episode_number}, 標題={episode_title}, 描述={episode_description}, 時間={episode_runtime} 分鐘")
                    all_episodes.append([season_number, episode_number, episode_runtime, episode_title, episode_description])
            
            except NoSuchElementException:
                # 如果項目內部沒有找到標題等元素，則記錄警告並安全跳過
                logging.warning("找到一個 'set-item' 但其內部結構不符合預期，已跳過。")
                continue

        logging.info(f"完成抓取第 {season_number} 季的集數")
    except Exception as e:
        logging.error(f"Error grabbing episodes for season {season_number}: {e}")
        failsafe(driver, driver.current_url, grab_episodes, driver, season_number, all_episodes)

def extract_disneyplus_data(driver, url, output_path, retry_count=5):
    global is_logged_in, profile_selected
    logging.info(f"開始抓取 Disney+ 資料，目標 URL: {url}")
    attempt = 0
    content_is_tv_show = None # 初始化返回值，None 表示提取失敗或類型未知

    while attempt < retry_count:
        try:
            if not is_disneyplus_logged_in(driver): # 確保 is_disneyplus_logged_in 是正確的
                logging.info("尚未登入 Disney+，開始登入")
                configs = load_configs()
                disneyplus_email = configs.get('disneyplus_email')
                disneyplus_password = configs.get('disneyplus_password')
                if not disneyplus_email or not disneyplus_password:
                    logging.error("Disney+ 帳號或密碼未在 configs.json 中設定。")
                    return None # 無法登入，提取失敗
                login_to_disneyplus(driver, disneyplus_email, disneyplus_password) # 確保 login_to_disneyplus 能處理各種情況

            logging.info(f"導航到 URL: {url}")
            driver.get(url)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            for i in range(5):
                if driver.execute_script("return document.readyState") == "complete":
                    logging.info("網頁已完全載入")
                    break
                else:
                    logging.info(f"等待網頁載入...第 {i+1} 次")
                    time.sleep(3)
            else: # 如果循環結束仍未 complete
                logging.warning("網頁在多次嘗試後仍未完全載入，可能影響提取。")


            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'li[data-testid="details-page-tab"][aria-controls="details"]'))
            )
            time.sleep(2) # 給予頁面穩定時間

            click_details_button(driver) # 確保此函數健壯

            logging.info("開始抓取標題和簡介")
            # 使用 WebDriverWait 增加穩定性
            title_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="details-tab-title"]'))
            )
            description_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="details-tab-description"]'))
            )

            title = add_spaces_between_en_zh_num(title_element.text)
            description = add_spaces_between_en_zh_num(description_element.text)

            unwanted_texts = [
                '部分閃光片段或圖案可能會影響對光敏感的觀眾。',
                '一些閃爍的燈光片段或圖案可能會影響有光敏的觀眾'
            ]
            for unwanted_text in unwanted_texts:
                if unwanted_text in description:
                    description = description.replace(unwanted_text, '').strip()

            logging.info(f"名稱: {title}")
            logging.info(f"簡介: {description}")

            # 判斷是劇集還是電影
            if driver.find_elements(By.CSS_SELECTOR, '[data-testid="details-page-tab"][aria-controls="episodes"]'):
                logging.info("檢測到此為劇集，開始抓取季數和集數資料")
                save_series_info_to_excel(title, description, output_path) # 假設此函數來自 excel_handler 或在此定義
                click_episodes_button(driver) # 確保此函數健壯
                seasons, episodes = extract_season_info(driver) # episodes 是 all_episodes
                
                if not seasons and not episodes:
                    logging.warning(f"警告：未能抓取到劇集 '{title}' 的任何季數或集數資訊。Excel 中的 Title sheet 可能已寫入，但 Seasons 和 Episodes sheet 將為空。")
                
                save_to_excel(seasons, episodes, output_path) # 假設此函數來自 excel_handler 或在此定義
                content_is_tv_show = True
            else:
                logging.info("檢測到此為電影，開始保存電影資料")
                runtime = extract_movie_runtime(driver) # 確保此函數健壯
                save_movie_info_to_excel(title, description, runtime, output_path) # 假設此函數來自 excel_handler 或在此定義
                content_is_tv_show = False

            logging.info(f"成功完成 Disney+ 資料抓取 ({'劇集' if content_is_tv_show else '電影'})")
            return content_is_tv_show # 成功提取並返回類型

        except TimeoutException:
            logging.error(f"抓取 Disney+ 資料時發生超時錯誤 (嘗試 {attempt + 1}/{retry_count})，URL: {url}")
            attempt += 1
            if attempt < retry_count:
                logging.info(f"等待 60 秒後重試第 {attempt + 1} 次...")
                time.sleep(60)
            # 若循環結束，則會在下面處理
        except Exception as e:
            logging.error(f"抓取 Disney+ 資料時發生未預期錯誤 (嘗試 {attempt + 1}/{retry_count})，URL: {url}: {e}\n{traceback.format_exc()}")
            attempt += 1
            if attempt < retry_count:
                logging.info(f"等待 60 秒後重試第 {attempt + 1} 次...")
                time.sleep(60)
            # 若循環結束，則會在下面處理

    # 如果所有重試都失敗 (attempt == retry_count)
    logging.error(f"extract_disneyplus_data 在所有 ({retry_count}) 次嘗試後失敗，URL: {url}。返回 None。")
    return None # 表示提取失敗或無法確定類型

def extract_movie_runtime(driver):
    try:
        runtime_element = driver.find_element(By.CSS_SELECTOR, 'span._1jf071b0')
        runtime_text = runtime_element.text
        return extract_runtime(runtime_text)
    except NoSuchElementException:
        return None

def save_movie_info_to_excel(title, description, runtime, output_path):
    try:
        df = pd.DataFrame([[title, description, runtime]], columns=['Movie Title', 'Movie Description', 'Movie Runtime'])
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name='Movies', index=False, header=False, startrow=writer.sheets['Movies'].max_row)
        release_excel_control(output_path)
    except Exception as e:
        logging.error(f"Error saving movie info to excel: {e}")
        failsafe(None, None, save_movie_info_to_excel, title, description, runtime, output_path)

def save_series_info_to_excel(title, description, output_path):
    try:
        if re.match(r"第\s*\d+\s*季[。\s]*$", title):
            title = ""
        df = pd.DataFrame([[title, description]], columns=['TV Show Title', 'TV Show Description'])
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name='Title', index=False, header=False, startrow=writer.sheets['Title'].max_row)
        release_excel_control(output_path)
    except Exception as e:
        logging.error(f"Error saving series info to excel: {e}")
        failsafe(None, None, save_series_info_to_excel, title, description, output_path)

def save_to_excel(seasons, episodes, output_path):
    try:
        seasons_df = pd.DataFrame(seasons, columns=['Season Name', 'Season Number', 'Season Description'])
        episodes_df = pd.DataFrame(episodes, columns=['Season Number', 'Episode Number', 'Episode Runtime', 'Episode Title', 'Episode Description'])
        
        logging.info(f"寫入季數數據到 Excel: {seasons}")
        logging.info(f"寫入集數數據到 Excel: {episodes}")

        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            seasons_df.to_excel(writer, sheet_name='Seasons', index=False, header=False, startrow=writer.sheets['Seasons'].max_row)
            episodes_df.to_excel(writer, sheet_name='Episodes', index=False, header=False, startrow=writer.sheets['Episodes'].max_row)
        
        release_excel_control(output_path)
    except Exception as e:
        logging.error(f"Error saving to excel: {e}")
        failsafe(None, None, save_to_excel, seasons, episodes, output_path)

def release_excel_control(output_path):
    """釋放對 Excel 文件的控制"""
    pass  # 在這裡可以添加任何需要的操作來釋放對 Excel 文件的控制

def check_login_error_and_relogin(driver, url):
    global is_logged_in, profile_selected
    try:
        login_button = driver.find_element(By.CSS_SELECTOR, "a[data-testid='login-header']")
        if login_button:
            logging.info("檢測到 LOG IN 按鈕，嘗試重新登錄")
            login_button.click()
            is_logged_in = False
            profile_selected = False
            configs = load_configs()
            email = configs.get('disneyplus_email')
            password = configs.get('disneyplus_password')
            login_to_disneyplus(driver, email, password)
            driver.get(url)
            time.sleep(2)  # 等待頁面加載
            extract_disneyplus_data(driver, url, 'video_detail.xlsx')
    except NoSuchElementException:
        logging.info("未檢測到 LOG IN 按鈕")

def failsafe(driver, url=None, func=None, *args, **kwargs):
    try:
        if driver:
            # 修正 XPath 語法
            if driver.find_elements(By.XPATH, "//*[contains(text(), \"Oops! We can't find the page you're looking for\")]"):
                logging.error("頁面找不到，跳過該項目。")
                return
            if driver.find_elements(By.XPATH, "//*[contains(text(), 'Wait and try again')]"):
                logging.warning("檢測到 'Wait and try again'，重試中...")
                time.sleep(60)
                driver.refresh()
                return
            # 檢查登入按鈕
            login_buttons = driver.find_elements(By.CSS_SELECTOR, "a[data-testid='login-header'], a[data-testid='log_in']")
            if not login_buttons:
                login_buttons = driver.find_elements(By.CSS_SELECTOR, "a[href='/login'], a[href='/identity/login']")
            if login_buttons:
                logging.info("檢測到登入按鈕，嘗試重新登錄")
                login_button = login_buttons[0]
                login_button.click()
                time.sleep(2)
                global is_logged_in, profile_selected
                is_logged_in = False
                profile_selected = False
                configs = load_configs()
                email = configs.get('disneyplus_email')
                password = configs.get('disneyplus_password')
                if email and password:
                    login_to_disneyplus(driver, email, password)
                    driver.get(url)
                    if func:
                        func(*args, **kwargs)
                else:
                    logging.error("無法從配置中獲取到 Disney+ 登錄憑據")
                return
        if func:
            logging.info("進入 failsafe 模式，嘗試重試該函數。")
            func(*args, **kwargs)
    except Exception as e:
        logging.error(f"Failsafe 過程中發生錯誤: {e}")