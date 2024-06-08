import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
import re

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def extract_primevideo_data(driver, url, output_path):
    driver.get(url)
    time.sleep(5)  # 等待頁面加載

    # 確認語言是否為繁體中文
    try:
        current_language = driver.find_element(By.CLASS_NAME, 'QDmWMz').text
        if current_language != 'ZH':
            # 點擊語言選擇器
            language_selector = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, 'bBPMYR'))
            )
            language_selector.click()
            # 點擊繁體中文選項
            traditional_chinese_option = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'form[action*="zh_TW"] input[type="submit"]'))
            )
            traditional_chinese_option.click()
            time.sleep(5)  # 等待頁面重新加載
    except Exception as e:
        logging.error(f"設置語言為繁體中文時發生錯誤: {e}")

    try:
        # 獲取劇集或電影名稱
        title_element = driver.find_element(By.CSS_SELECTOR, 'h1[data-automation-id="title"]')
        title = title_element.text

        # 獲取劇集或電影簡介
        description_element = driver.find_element(By.CSS_SELECTOR, 'span._1H6ABQ')
        description = description_element.text.strip()

        logging.info(f"名稱: {title}")
        logging.info(f"簡介: {description}")

        # 檢查是否為劇集
        if driver.find_elements(By.ID, 'tab-content-episodes'):
            click_episodes_button(driver)
            save_series_info_to_excel(title, description, output_path)
            extract_primevideo_seasons_and_episodes(driver, output_path, title, description)
        else:
            save_movie_info_to_excel(title, description, output_path)

    except Exception as e:
        logging.error(f"抓取 Prime Video 資料時發生錯誤: {e}")

def click_episodes_button(driver):
    logging.info("開始尋找「劇集」按鈕")
    try:
        episodes_tab_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="btf-episodes-tab"]'))
        )
        logging.info("找到「劇集」按鈕")
        episodes_tab_button.click()
        logging.info("成功點擊了「劇集」按鈕")
        time.sleep(3)  # 等待頁面加載
    except Exception as e:
        logging.error(f"尋找或點擊「劇集」按鈕時發生錯誤: {e}")

def extract_season_info(driver, season_number, title, description):
    # 獲取季數名稱
    season_name_element = driver.find_elements(By.CSS_SELECTOR, 'span._36qUej')
    if season_name_element:
        season_name = next((elem.text for elem in season_name_element if '第' in elem.text and '季' in elem.text), f'第 {season_number} 季')
    else:
        season_name = f'第 {season_number} 季'

    season_description_element = driver.find_elements(By.CSS_SELECTOR, 'span._1H6ABQ')
    season_description = season_description_element[0].text.strip() if season_description_element else description

    logging.info(f"找到季數: {season_name}")
    logging.info(f"季數描述: {season_description}")

    return season_number, season_name, season_description

def extract_primevideo_seasons_and_episodes(driver, output_path, title, description):
    episodes = []
    seasons = []

    try:
        # 檢查是否有下拉式季數選單
        season_elements = driver.find_elements(By.CSS_SELECTOR, 'div._3R4jka ul li a')
        if season_elements:
            season_links = [elem.get_attribute('href') for elem in season_elements if elem.get_attribute('href')]
        else:
            # 處理只有一季的情況
            season_links = [driver.current_url]

        if not season_links:
            logging.error("未找到任何季數連結")
            # 直接抓取當前頁面的劇集資料
            season_number = 1
            season_number, season_name, season_description = extract_season_info(driver, season_number, title, description)
            extract_single_season_episodes(driver, output_path, season_number, season_name, season_description)
        else:
            for idx, season_link in enumerate(season_links):
                logging.info(f"正在處理季數連結: {season_link}")
                season_number = idx + 1
                driver.get(season_link)
                time.sleep(3)  # 等待頁面加載
                season_number, season_name, season_description = extract_season_info(driver, season_number, title, description)
                seasons.append([season_number, season_name, season_description])

                # 獲取每一集的資料
                episode_elements = driver.find_elements(By.CSS_SELECTOR, 'li[id^="av-ep-episodes-"]')
                logging.info(f"季數 {season_number} 找到 {len(episode_elements)} 集")
                for episode_element in episode_elements:
                    try:
                        episode_info = episode_element.find_element(By.CSS_SELECTOR, 'span._36qUej').text
                        episode_number_match = re.search(r'季第 \d+ 集(\d+)', episode_info)
                        episode_number = int(episode_number_match.group(1)) if episode_number_match else None
                        episode_title = episode_element.find_element(By.CSS_SELECTOR, 'span.P1uAb6').text
                        episode_description = episode_element.find_element(By.CSS_SELECTOR, 'div._3qsVvm.e8yjMf > div[dir="auto"]').text

                        episodes.append([season_number, episode_number, episode_title, episode_description])
                        logging.info(f"季數 {season_number} 第 {episode_number} 集 - 標題: {episode_title}")
                    except Exception as e:
                        logging.error(f"抓取劇集資料時發生錯誤: {e}")

        # 無論是否找到季數連結，都嘗試保存資料
        save_seasons_info_to_excel(seasons, output_path)
        save_episodes_info_to_excel(episodes, output_path)

    except Exception as e:
        logging.error(f"抓取季數資料時發生錯誤: {e}")

def extract_single_season_episodes(driver, output_path, season_number, season_name, season_description):
    episodes = []
    try:
        # 獲取每一集的資料
        episode_elements = driver.find_elements(By.CSS_SELECTOR, 'li[id^="av-ep-episodes-"]')
        logging.info(f"找到 {len(episode_elements)} 集")
        for episode_element in episode_elements:
            try:
                episode_info = episode_element.find_element(By.CSS_SELECTOR, 'span._36qUej').text
                episode_number_match = re.search(r'季第 \d+ 集(\d+)', episode_info)
                episode_number = int(episode_number_match.group(1)) if episode_number_match else None
                episode_title = episode_element.find_element(By.CSS_SELECTOR, 'span.P1uAb6').text
                episode_description = episode_element.find_element(By.CSS_SELECTOR, 'div._3qsVvm.e8yjMf > div[dir="auto"]').text

                episodes.append([season_number, episode_number, episode_title, episode_description])
                logging.info(f"季數 {season_number} 第 {episode_number} 集 - 標題: {episode_title}")
            except Exception as e:
                logging.error(f"抓取劇集資料時發生錯誤: {e}")

        save_episodes_info_to_excel(episodes, output_path)
        save_seasons_info_to_excel([[season_number, season_name, season_description]], output_path)

    except Exception as e:
        logging.error(f"抓取劇集資料時發生錯誤: {e}")

def save_movie_info_to_excel(movie_title, movie_description, output_path):
    movie_data = {'Movie Title': [movie_title], 'Movie Description': [movie_description]}
    movie_df = pd.DataFrame(movie_data)
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        movie_df.to_excel(writer, sheet_name='Movies', index=False)

def save_series_info_to_excel(tv_show_title, tv_show_description, output_path):
    title_data = {'Series Title': [tv_show_title], 'Series Description': [tv_show_description]}
    title_df = pd.DataFrame(title_data)
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        title_df.to_excel(writer, sheet_name='Title', index=False)

def save_seasons_info_to_excel(seasons, output_path):
    seasons_df = pd.DataFrame(seasons, columns=['Season Number', 'Season Name', 'Season Description'])
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        seasons_df.to_excel(writer, sheet_name='Seasons', index=False)

def save_episodes_info_to_excel(episodes, output_path):
    episodes_df = pd.DataFrame(episodes, columns=['Season Number', 'Episode Number', 'Episode Title', 'Episode Description'])
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        episodes_df.to_excel(writer, sheet_name='Episodes', index=False)
