import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
from selenium.common.exceptions import NoSuchElementException

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def login_to_disneyplus(driver, email, password):
    driver.get("https://www.disneyplus.com/zh-hk/identity/login")
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(email)
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(password)
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="button"][data-testid^="profile-avatar-"]')))
    profiles = driver.find_elements(By.CSS_SELECTOR, 'div[role="button"][data-testid^="profile-avatar-"]')
    profiles[0].click()  # 隨機選擇第一個 Profile
    time.sleep(3)  # 等待 3 秒

def click_details_button(driver):
    logging.info("開始尋找「簡介」按鈕")
    details_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'li[data-testid="details-page-tab"][aria-controls="details"]'))
    )
    details_button.click()
    logging.info("已點擊「簡介」按鈕")
    time.sleep(2)  # 等待 5 秒以確保頁面完全加載

def click_episodes_button(driver):
    logging.info("開始尋找「集數」按鈕")
    episodes_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'li[data-testid="details-page-tab"][aria-controls="episodes"]'))
    )
    episodes_button.click()
    logging.info("已點擊「集數」按鈕")
    time.sleep(2)  # 等待 5 秒以確保頁面完全加載

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

def extract_season_info(driver):
    logging.info("開始抓取季數資訊")
    seasons = []
    all_episodes = []
    season_names = set()

    # 檢查是否有多季
    try:
        dropdown_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="dropdown-button"]'))
        )
        click_using_js(driver, dropdown_button)
        logging.info("已點擊下拉菜單按鈕")
        time.sleep(2)  # 等待 3 秒

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
            season_name = season_elements[i].text
            if season_name in season_names:
                continue
            season_names.add(season_name)
            current_season_number = int(season_name.split(' ')[1])
            seasons.append([season_name, current_season_number, ''])
            season_elements[i].click()
            time.sleep(2)  # 等待 5 秒以確保頁面完全加載
            grab_episodes(driver, current_season_number, all_episodes)
            scroll_to_top(driver)  # 滾動到頂部以避免點擊被阻擋
            dropdown_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="dropdown-button"]'))
            )
            click_using_js(driver, dropdown_button)
            time.sleep(1)  # 再次打開下拉菜單
    except NoSuchElementException as e:
        logging.info("找不到多季資訊: %s", e)
        # 如果只有一季，抓取單一季數資訊
        try:
            disabled_dropdown = driver.find_element(By.CSS_SELECTOR, '[data-testid="dropdown-button"][disabled]')
            season_name = disabled_dropdown.find_element(By.TAG_NAME, 'span').text
            current_season_number = int(season_name.split(' ')[1])
            seasons.append([season_name, current_season_number, ''])
            grab_episodes(driver, current_season_number, all_episodes)
        except NoSuchElementException as ex:
            logging.info("找不到單一季數資訊: %s", ex)

    return seasons, all_episodes

def grab_episodes(driver, season_number, all_episodes):
    logging.info(f"開始抓取第 {season_number} 季的集數")
    scroll_to_bottom(driver)
    episodes = driver.find_elements(By.CSS_SELECTOR, '[data-testid="set-item"]')
    episode_numbers = set()
    for episode in episodes:
        episode_title = episode.find_element(By.CSS_SELECTOR, '[data-testid="standard-regular-list-item-title"]').text
        episode_description = episode.find_element(By.CSS_SELECTOR, '[data-testid="standard-regular-list-item-description"]').text
        episode_number = int(episode_title.split('.')[0])
        if episode_number in episode_numbers:
            continue
        episode_numbers.add(episode_number)
        episode_title = episode_title.split('.')[1].strip()

        # 去除不需要的段落
        unwanted_text = '部分閃光片段或圖案可能會影響對光敏感的觀眾。'
        if unwanted_text in episode_description:
            episode_description = episode_description.replace(unwanted_text, '').strip()

        all_episodes.append([season_number, episode_number, episode_title, episode_description])
    logging.info(f"完成抓取第 {season_number} 季的集數")

def extract_disneyplus_data(driver, url, output_path):
    driver.get(url)
    time.sleep(2)  # 等待頁面加載

    try:
        click_details_button(driver)
        # 抓取標題和簡介
        title_element = driver.find_element(By.CSS_SELECTOR, '[data-testid="details-tab-title"]')
        description_element = driver.find_element(By.CSS_SELECTOR, '[data-testid="details-tab-description"]')

        title = title_element.text
        description = description_element.text

        # 去除不需要的段落
        unwanted_text = '部分閃光片段或圖案可能會影響對光敏感的觀眾。'
        if unwanted_text in description:
            description = description.replace(unwanted_text, '').strip()

        logging.info(f"名稱: {title}")
        logging.info(f"簡介: {description}")

        # 檢查是否為電影
        is_movie = not driver.find_elements(By.CSS_SELECTOR, '[aria-controls="episodes"]')

        if is_movie:
            save_movie_info_to_excel(title, description, output_path)
        else:
            save_series_info_to_excel(title, description, output_path)
            click_episodes_button(driver)
            seasons, episodes = extract_season_info(driver)
            save_to_excel(seasons, episodes, output_path)

    except Exception as e:
        logging.error(f"抓取 Disney+ 資料時發生錯誤: {e}")

def save_movie_info_to_excel(movie_title, movie_description, output_path):
    movie_data = {'Movie Title': [movie_title], 'Movie Description': [movie_description]}
    movie_df = pd.DataFrame(movie_data)
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        movie_df.to_excel(writer, sheet_name='Movies', index=False)

def save_series_info_to_excel(tv_show_title, tv_show_description, output_path):
    title_data = {'TV Show Title': [tv_show_title], 'TV Show Description': [tv_show_description]}
    title_df = pd.DataFrame(title_data)
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        title_df.to_excel(writer, sheet_name='Title', index=False)

def save_to_excel(seasons, episodes, output_path):
    logging.info("保存資料到 Excel")
    # 為每季信息建立 DataFrame
    seasons_df = pd.DataFrame(seasons, columns=['Season Name', 'Season Number', 'Season Description'])
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        seasons_df.to_excel(writer, sheet_name='Seasons', index=False)

    # 保存集數信息到 Episodes
    episodes_df = pd.DataFrame(episodes, columns=['Season Number', 'Episode Number', 'Episode Title', 'Episode Description'])
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        episodes_df.to_excel(writer, sheet_name='Episodes', index=False)

    logging.info(f"資料已保存到 {output_path}")

if __name__ == "__main__":
    from selenium import webdriver
    driver = webdriver.Chrome()
    email = 'YOUR_EMAIL'
    password = 'YOUR_PASSWORD'
    url = 'YOUR_URL'
    output_path = './output.xlsx'

    login_to_disneyplus(driver, email, password)
    extract_disneyplus_data(driver, url, output_path)
