import json
import os
import pandas as pd
from selenium import webdriver
from extractors.netflix_extractor import extract_netflix_episodes
from others.analysis_excel import process_excel
from importors.tmdb_uploader import TMDBUploader
from extractors.appletv_extractor import extract_appletv_data
from extractors.disneyplus_extractor import login_to_disneyplus, extract_disneyplus_data
from extractors.primevideo_extractor import extract_primevideo_data
import logging

# 設置環境變量來抑制 TensorFlow Lite 的訊息
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'

# 設置日誌級別，忽略不需要的訊息
logging.basicConfig(level=logging.WARNING, format='%(asctime)s - %(levelname)s - %(message)s')

# 檢查和創建 Excel 文件
def check_create_excel():
    if not os.path.exists('video_detail.xlsx'):
        with pd.ExcelWriter('video_detail.xlsx') as writer:
            pd.DataFrame(columns=['TV Show Title', 'TV Show Description']).to_excel(writer, sheet_name='Title', index=False)
            pd.DataFrame(columns=['Season Name', 'Season Number', 'Season Description']).to_excel(writer, sheet_name='Seasons', index=False)
            pd.DataFrame(columns=['Season Number', 'Episode Number', 'Episode Title', 'Episode Description']).to_excel(writer, sheet_name='Episodes', index=False)
            pd.DataFrame(columns=['Movie Title', 'Movie Description']).to_excel(writer, sheet_name='Movies', index=False)
    return pd.ExcelFile('video_detail.xlsx')

# 清空 Excel 文件中的數據，但保留欄位
def clear_excel():
    with pd.ExcelWriter('video_detail.xlsx') as writer:
        pd.DataFrame(columns=['TV Show Title', 'TV Show Description']).to_excel(writer, sheet_name='Title', index=False)
        pd.DataFrame(columns=['Season Name', 'Season Number', 'Season Description']).to_excel(writer, sheet_name='Seasons', index=False)
        pd.DataFrame(columns=['Season Number', 'Episode Number', 'Episode Title', 'Episode Description']).to_excel(writer, sheet_name='Episodes', index=False)
        pd.DataFrame(columns=['Movie Title', 'Movie Description']).to_excel(writer, sheet_name='Movies', index=False)

def load_configs():
    if os.path.exists('configs.json'):
        with open('configs.json', 'r') as config_file:
            return json.load(config_file)
    return {}

def save_configs(configs):
    with open('configs.json', 'w') as config_file:
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

def main():
    driver = None
    tmdb_uploader = None
    disneyplus_logged_in = False
    configs = load_configs()

    while True:
        # 問詢用戶輸入
        video_url = input("請輸入 Video site 的 URL 或 'excel': ")
        
        while True:
            tmdb_url = input("請輸入 TMDB 的 URL (可留空): ")
            if tmdb_url == "" or "themoviedb.org" in tmdb_url:
                break
            else:
                print("只可以輸入 TMDB 的網址或直接按 enter。請重新輸入。")

        if driver is None:
            options = webdriver.ChromeOptions()
            options.add_argument('--log-level=3')  # 隱藏所有的 INFO 以下日誌信息
            driver = webdriver.Chrome(options=options)

        while True:
            try:
                # 檢查和創建 Excel 文件
                check_create_excel()

                # 如果用戶輸入 'excel'，直接上傳現有 Excel 資料到 TMDB
                if video_url.lower() == 'excel':
                    if tmdb_url:
                        tmdb_username, tmdb_password = get_tmdb_credentials(configs)
                        if tmdb_uploader is None:
                            tmdb_uploader = TMDBUploader(driver, tmdb_username, tmdb_password)
                        is_movie = "/movie/" in tmdb_url
                        tmdb_uploader.upload_to_tmdb(tmdb_url, 'video_detail.xlsx', is_movie)
                else:
                    # 清空 Excel 文件中的數據
                    clear_excel()

                    # 處理 Video URL
                    if "netflix.com" in video_url:
                        extract_netflix_episodes(driver, video_url, 'video_detail.xlsx')
                    elif "tv.apple.com" in video_url:
                        extract_appletv_data(driver, video_url, 'video_detail.xlsx')
                    elif "disneyplus.com" in video_url:
                        disneyplus_email, disneyplus_password = get_disneyplus_credentials(configs)
                        if not disneyplus_logged_in:
                            login_to_disneyplus(driver, disneyplus_email, disneyplus_password)
                            disneyplus_logged_in = True
                        extract_disneyplus_data(driver, video_url, 'video_detail.xlsx')
                    elif "primevideo.com" in video_url:
                        extract_primevideo_data(driver, video_url, 'video_detail.xlsx')
                    
                    # 確保文件存在並有數據
                    if os.path.exists('video_detail.xlsx'):
                        process_excel('video_detail.xlsx', 'video_detail.xlsx')

                    # 處理 TMDB 上傳
                    if tmdb_url:
                        tmdb_username, tmdb_password = get_tmdb_credentials(configs)
                        if tmdb_uploader is None:
                            tmdb_uploader = TMDBUploader(driver, tmdb_username, tmdb_password)
                        is_movie = "/movie/" in tmdb_url
                        tmdb_uploader.upload_to_tmdb(tmdb_url, 'video_detail.xlsx', is_movie)
                break
            except PermissionError as e:
                logging.error(f"發生錯誤: {e}")
                input("請關閉 video_detail.xlsx，然後按任何鍵繼續...")
                continue
            except Exception as e:
                logging.error(f"發生錯誤: {e}")
                break

        continue_use = input("是否繼續? (y/n): ")
        if continue_use.lower() == 'n':
            if driver:
                driver.quit()
            break

if __name__ == "__main__":
    main()
