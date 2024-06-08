import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import re
import logging

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def extract_appletv_data(driver, url, output_path):
    driver.get(url)
    time.sleep(1)  # 等待頁面加載

    try:
        # 提取劇名
        title_tag = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'script#schema\\:breadcrumb-list')))
        title_json = json.loads(title_tag.get_attribute('innerHTML'))
        tv_show_title = title_json['itemListElement'][-1]['item']['name']

        # 提取簡介
        description_tag = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.product-header__content__details__synopsis')))
        tv_show_description = description_tag.text.strip()

        logging.info(f"劇集名稱: {tv_show_title}")
        logging.info(f"劇集簡介: {tv_show_description}")

        # 檢查是否為電影
        if not driver.find_elements(By.CSS_SELECTOR, '.episode-lockup__content'):
            save_movie_info_to_excel(tv_show_title, tv_show_description, output_path)
            logging.info("識別為電影，已保存到 Movie 頁面")
            return

        # 提取季數信息
        page_source = driver.page_source
        season_data_match = re.search(r'"seasonSummaries":\[(.*?)\],"selectedEpisodeIndex"', page_source)
        if season_data_match:
            season_data_json = '[' + season_data_match.group(1).replace('},{', '}|{') + ']'
            season_data_json = season_data_json.replace('|', ',')
            season_data = json.loads(season_data_json)
        else:
            season_data = []

        seasons = []
        all_episodes = []
        seen_titles = set()  # 存儲已抓取的集數標題，避免重複
        current_episode_number = 0

        for season in season_data:
            season_name = season['title']
            season_number = season['seasonNumber']
            seasons.append((season_name, season_number, ""))  # 目前沒有季數簡介

            logging.info(f"季數: {season_name} (第 {season_number} 季)")

            # 抓取每季的集數信息
            while True:
                try:
                    episodes = driver.find_elements(By.CSS_SELECTOR, '.episode-lockup__content')
                    for episode in episodes:
                        try:
                            title = episode.find_element(By.CSS_SELECTOR, '.typ-subhead.text-truncate.episode-lockup__content__title').text
                            if title not in seen_titles:
                                episode_number_text = episode.find_element(By.CSS_SELECTOR, '.episode-lockup__content__episode-number span').text.strip()
                                episode_number = int(episode_number_text.split(" ")[1].strip("集"))

                                if episode_number < current_episode_number:
                                    # 假定這是新的季數
                                    season_number += 1

                                current_episode_number = episode_number

                                description = episode.find_element(By.CSS_SELECTOR, '.episode-lockup__description.clr-secondary-text').text
                                all_episodes.append((season_number, episode_number, title, description))
                                seen_titles.add(title)

                                logging.info(f"  集數: {episode_number_text} - 標題: {title}")
                                logging.info(f"    簡介: {description}")
                        except Exception as e:
                            logging.error(f"抓取集數資料時發生錯誤: {e}")

                    # 查找“下一頁”按鈕並點擊
                    next_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.shelf-grid-nav__arrow.shelf-grid-nav__arrow--next'))
                    )
                    next_button.click()
                    time.sleep(1)  # 等待新內容加載

                except Exception as e:
                    logging.info(f"沒有更多頁面或發生錯誤: {e}")
                    break

        logging.info(f"所有集數信息: {all_episodes}")

        # 保存到 Excel
        save_to_excel(tv_show_title, tv_show_description, seasons, all_episodes, output_path)

    except Exception as e:
        logging.error(f"抓取 Apple TV 資料時發生錯誤: {e}")

def save_movie_info_to_excel(movie_title, movie_description, output_path):
    movie_data = {'Movie Title': [movie_title], 'Movie Description': [movie_description]}
    movie_df = pd.DataFrame(movie_data)
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        movie_df.to_excel(writer, sheet_name='Movie', index=False)

def save_to_excel(tv_show_title, tv_show_description, seasons, episodes, output_path):
    # 保存劇集信息到 Title
    title_data = {'TV Show Title': [tv_show_title], 'TV Show Description': [tv_show_description]}
    title_df = pd.DataFrame(title_data)
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        title_df.to_excel(writer, sheet_name='Title', index=False)

    # 保存季數信息到 Seasons
    seasons_df = pd.DataFrame(seasons, columns=['Season Name', 'Season Number', 'Season Description'])
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        seasons_df.to_excel(writer, sheet_name='Seasons', index=False)

    # 保存集數信息到 Episodes
    episodes_df = pd.DataFrame(episodes, columns=['Season Number', 'Episode Number', 'Episode Title', 'Episode Description'])
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        episodes_df.to_excel(writer, sheet_name='Episodes', index=False)

    logging.info(f"資料已保存到 {output_path}")

if __name__ == "__main__":
    import logging
    logging.getLogger('page_load_metrics_update_dispatcher').setLevel(logging.ERROR)

    apple_tv_url = input("請輸入 Apple TV 劇集 URL: ")
    output_file_path = 'video_detail.xlsx'
    driver = webdriver.Chrome()
    extract_appletv_data(driver, apple_tv_url, output_file_path)
    driver.quit()
