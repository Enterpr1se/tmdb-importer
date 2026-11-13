import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import json
import re
import logging
from others.utils import add_spaces_between_en_zh_num

def extract_appletv_data(driver, url, output_path):
    driver.get(url)
    time.sleep(2)  # 等待頁面初步加載，應對 JS 渲染延遲

    tv_show_title = ""
    tv_show_description = ""

    try:
        # --- 策略 1: 新版網頁結構 (主要策略) ---
        logging.info("嘗試使用新版網頁結構抓取資料 (about-product-footer)...")
        about_section = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.about-product-footer"))
        )
        title_element = about_section.find_element(By.CSS_SELECTOR, 'div.card:first-of-type h3.title')
        tv_show_title = add_spaces_between_en_zh_num(title_element.text.strip())

        # --- 精煉後的簡介抓取邏輯 ---
        try:
            # 檢查 "更多" 按鈕是否存在
            more_button = about_section.find_element(By.CSS_SELECTOR, 'button[data-testid="truncate-more-button"]')
            logging.info("找到 '更多' 按鈕，將點擊以獲取完整簡介。")
            driver.execute_script("arguments[0].click();", more_button)
            time.sleep(0.5) # 等待對話框動畫
            
            # 等待對話框內的簡介元素出現
            dialog_description_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'dialog[data-testid="dialog"][open] p[data-testid="content-modal-text"]'))
            )
            tv_show_description = add_spaces_between_en_zh_num(dialog_description_element.text.strip())
            logging.info("成功從彈出式對話框抓取完整簡介。")

        except (NoSuchElementException, TimeoutException):
            # 如果沒有 "更多" 按鈕或處理失敗，直接抓取頁面上的簡介
            logging.info("未找到 '更多' 按鈕或處理對話框時發生錯誤，直接抓取頁面上的簡介。")
            description_element = about_section.find_element(By.CSS_SELECTOR, 'p[data-testid="truncate-text"]')
            tv_show_description = add_spaces_between_en_zh_num(description_element.text.strip())

        logging.info("新版網頁結構資料抓取完成。")

    except (TimeoutException, NoSuchElementException) as e1:
        logging.warning(f"使用新版網頁結構抓取失敗: {e1}。嘗試退回舊版結構 (schema/breadcrumb)...")
        try:
            # --- 策略 2: 舊版網頁結構 (備用方案) ---
            title_tag = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'script#schema\\:breadcrumb-list'))
            )
            title_json = json.loads(title_tag.get_attribute('innerHTML'))
            tv_show_title = add_spaces_between_en_zh_num(title_json['itemListElement'][-1]['item']['name'])

            description_tag = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.product-header__content__details__synopsis'))
            )
            tv_show_description = add_spaces_between_en_zh_num(description_tag.text.strip())
            logging.info("成功使用舊版網頁結構抓取到標題與簡介。")

        except Exception as e2:
            logging.error(f"致命錯誤: 所有抓取策略均失敗。新結構錯誤: {e1}。舊結構錯誤: {e2}")
            raise Exception("使用所有可用方法從 Apple TV 頁面抓取資料均失敗。") from e2

    if not tv_show_title:
        raise Exception("致命錯誤: 嘗試所有策略後仍無法抓取到標題。")

    # 驗證簡介內容，防止抓到無效的 "..."
    if tv_show_description.strip() == "...":
        logging.warning("抓取到的簡介內容為 '...'，表示簡介抓取不完整，將其視為空值。")
        tv_show_description = ""

    logging.info(f"最終抓取到的標題: {tv_show_title}")
    logging.info(f"最終抓取到的簡介: {tv_show_description[:100]}...")

    # 檢查是電影還是劇集 (此邏輯沿用舊版)
    if not driver.find_elements(By.CSS_SELECTOR, '.episode-lockup__content'):
        save_movie_info_to_excel(tv_show_title, tv_show_description, output_path)
        logging.info("識別為電影，資料已保存到 Movie 工作表。")
        return

    # 如果是劇集，繼續抓取季數與集數
    logging.info("識別為劇集，繼續抓取季數與集數資訊...")
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
    seen_titles = set()

    for season_info in season_data:
        season_name = add_spaces_between_en_zh_num(season_info['title'])
        season_number = season_info['seasonNumber']
        seasons.append((season_name, season_number, ""))

        logging.info(f"處理季: {season_name} (第 {season_number} 季)")
        
        while True:
            try:
                episodes_on_page = driver.find_elements(By.CSS_SELECTOR, '.episode-lockup__content')
                new_episodes_found_on_page = False
                for episode in episodes_on_page:
                    try:
                        title = add_spaces_between_en_zh_num(episode.find_element(By.CSS_SELECTOR, '.typ-subhead.text-truncate.episode-lockup__content__title').text)
                        if title not in seen_titles:
                            new_episodes_found_on_page = True
                            seen_titles.add(title)
                            
                            episode_number_text = episode.find_element(By.CSS_SELECTOR, '.episode-lockup__content__episode-number span').text.strip()
                            episode_number_match = re.search(r'\d+', episode_number_text)
                            if episode_number_match:
                                episode_number = int(episode_number_match.group(0))
                            else:
                                logging.warning(f"無法從 '{episode_number_text}' 解析集數，跳過此集。")
                                continue

                            description = add_spaces_between_en_zh_num(episode.find_element(By.CSS_SELECTOR, '.episode-lockup__description.clr-secondary-text').text)
                            all_episodes.append((season_number, episode_number, title, description))
                            logging.info(f"  集: S{season_number}E{episode_number} - 標題: {title}")
                    except Exception as e_ep:
                        logging.error(f"抓取單一集數資料時發生錯誤: {e_ep}")

                if not new_episodes_found_on_page and len(episodes_on_page) > 0:
                     logging.info("此頁面未發現新集數，視為本季結束。")
                     break

                next_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.shelf-grid-nav__arrow.shelf-grid-nav__arrow--next'))
                )
                next_button.click()
                time.sleep(1.5)

            except Exception as e_page:
                logging.info(f"沒有更多頁面或在季內分頁時發生錯誤: {e_page}")
                break

    logging.info(f"所有集數資訊已抓取完成，共 {len(all_episodes)} 集。")
    save_to_excel(tv_show_title, tv_show_description, seasons, all_episodes, output_path)

def save_movie_info_to_excel(movie_title, movie_description, output_path):
    movie_data = {'Movie Title': [movie_title], 'Movie Description': [movie_description]}
    movie_df = pd.DataFrame(movie_data)
    # 使用 'replace' 確保在清空 sheet 後能正確寫入
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        movie_df.to_excel(writer, sheet_name='Movies', index=False)

def save_to_excel(tv_show_title, tv_show_description, seasons, episodes, output_path):
    # 保存劇集信息到 Title
    title_data = {'TV Show Title': [tv_show_title], 'TV Show Description': [tv_show_description]}
    title_df = pd.DataFrame(title_data)
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        title_df.to_excel(writer, sheet_name='Title', index=False)

    # 保存季數信息到 Seasons
    seasons_df = pd.DataFrame(seasons, columns=['Season Name', 'Season Number', 'Season Description'])
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        seasons_df.to_excel(writer, sheet_name='Seasons', index=False)

    # 保存集數信息到 Episodes
    episodes_df = pd.DataFrame(episodes, columns=['Season Number', 'Episode Number', 'Episode Title', 'Episode Description'])
    with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        episodes_df.to_excel(writer, sheet_name='Episodes', index=False)

    logging.info(f"資料已保存到 {output_path}")

if __name__ == "__main__":
    from selenium import webdriver
    logging.getLogger('page_load_metrics_update_dispatcher').setLevel(logging.ERROR)

    apple_tv_url = input("請輸入 Apple TV 劇集或電影 URL: ")
    output_file_path = 'video_detail.xlsx'
    
    # 請確保已配置好您的 WebDriver
    # 例如：driver = webdriver.Chrome()
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless') # 可選：在背景執行
    driver = webdriver.Chrome(options=options)
    
    try:
        extract_appletv_data(driver, apple_tv_url, output_file_path)
    finally:
        driver.quit()