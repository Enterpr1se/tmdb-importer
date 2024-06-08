import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import logging

class TMDBUploader:
    def __init__(self, driver, username, password):
        self.driver = driver
        self.username = username
        self.password = password
        self.login()

    def login(self):
        self.driver.get("https://www.themoviedb.org/login")
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(self.username)
        self.driver.find_element(By.NAME, "password").send_keys(self.password, Keys.RETURN)
        time.sleep(1)
        logging.info("Logged in to TMDB.")
        self.accept_cookies()

    def accept_cookies(self):
        try:
            accept_button = self.driver.find_element(By.ID, "onetrust-accept-btn-handler")
            accept_button.click()
            time.sleep(1)
            logging.info("Accepted cookies.")
        except:
            logging.info("No cookies to accept or already accepted.")

    def check_and_add_translation(self):
        try:
            add_translation_button = self.driver.find_element(By.CSS_SELECTOR, "button.k-button.k-primary.pad_top.background_color.light_blue.translate")
            if add_translation_button:
                add_translation_button.click()
                time.sleep(3)  # 按下新增翻譯後等待 3 秒
                logging.info("Added translation for zh-HK.")
                return True
        except:
            logging.info("No translation needed or already present.")
        return False

    def extract_tv_show_id(self, url):
        return url.split('/tv/')[1].split('?')[0].split('/')[0]

    def extract_movie_id(self, url):
        return url.split('/movie/')[1].split('?')[0].split('/')[0]

    def update_series_info(self, url, excel_path):
        tv_show_id = self.extract_tv_show_id(url)
        self.driver.get(f"https://www.themoviedb.org/tv/{tv_show_id}/edit?language=zh-HK")
        logging.info(f"Updating series info for URL: https://www.themoviedb.org/tv/{tv_show_id}/edit?language=zh-HK")

        if not self.check_and_fill_form('zh_HK_name', 'zh_HK_overview', excel_path, 'Title'):
            if self.check_and_add_translation():
                self.driver.get(f"https://www.themoviedb.org/tv/{tv_show_id}/edit?language=zh-HK")
                self.check_and_fill_form('zh_HK_name', 'zh_HK_overview', excel_path, 'Title')

    def update_seasons(self, url, excel_path):
        tv_show_id = self.extract_tv_show_id(url)

        df = pd.read_excel(excel_path, sheet_name='Seasons')

        for index, row in df.iterrows():
            season_name = row['Season Name']
            season_number = row['Season Number']
            season_description = row['Season Description']
            
            self.driver.get(f"https://www.themoviedb.org/tv/{tv_show_id}/season/{season_number}/edit?language=zh-HK")
            logging.info(f"Updating season info for URL: https://www.themoviedb.org/tv/{tv_show_id}/season/{season_number}/edit?language=zh-HK")

            self.check_and_fill_form('zh_HK_name', 'zh_HK_overview', excel_path, 'Seasons', season_name, season_description, skip_empty_overview=True)

    def update_episodes(self, url, excel_path):
        tv_show_id = self.extract_tv_show_id(url)

        df = pd.read_excel(excel_path, sheet_name='Episodes')

        for index, row in df.iterrows():
            season = row['Season Number']
            episode_number = row['Episode Number']
            episode_title = row['Episode Title']
            episode_description = row['Episode Description']
            
            self.driver.get(f"https://www.themoviedb.org/tv/{tv_show_id}/season/{season}/episode/{episode_number}/edit?language=zh-HK")
            logging.info(f"Updating episode info for URL: https://www.themoviedb.org/tv/{tv_show_id}/season/{season}/episode/{episode_number}/edit?language=zh-HK")

            self.check_and_fill_form('zh_HK_name', 'zh_HK_overview', excel_path, 'Episodes', episode_title, episode_description, skip_empty_overview=True)

    def update_movie_info(self, url, excel_path):
        movie_id = self.extract_movie_id(url)
        self.driver.get(f"https://www.themoviedb.org/movie/{movie_id}/edit?language=zh-HK")
        logging.info(f"Updating movie info for URL: https://www.themoviedb.org/movie/{movie_id}/edit?language=zh-HK")

        if not self.check_and_fill_form('zh_HK_translated_title', 'zh_HK_overview', excel_path, 'Movies'):
            if self.check_and_add_translation():
                self.driver.get(f"https://www.themoviedb.org/movie/{movie_id}/edit?language=zh-HK")
                self.check_and_fill_form('zh_HK_translated_title', 'zh_HK_overview', excel_path, 'Movies')

    def check_and_fill_form(self, title_field_id, overview_field_id, excel_path, sheet_name, title=None, description=None, skip_empty_overview=False):
        try:
            name_field = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, title_field_id)))
            overview_field = self.driver.find_element(By.ID, overview_field_id)

            if title is None or description is None:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                title = df.iloc[0, 0]
                description = df.iloc[0, 1]

            name_field.clear()
            name_field.send_keys(title)
            logging.info(f"Updated title: {title}")

            if not skip_empty_overview or (description and str(description).strip() != 'nan'):
                logging.info("Updating description...")
                overview_field.clear()
                overview_field.send_keys(description)

            submit_button = self.driver.find_element(By.ID, "submit")
            self.driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
            submit_button.click()
            time.sleep(1)
            logging.info("Successfully updated the information.")
            return True
        except Exception as e:
            logging.warning(f"No translation found or other issue: {e}")
            return False

    def upload_to_tmdb(self, url, excel_path, is_movie=False):
        if is_movie:
            self.update_movie_info(url, excel_path)
        else:
            self.update_series_info(url, excel_path)
            self.update_seasons(url, excel_path)
            self.update_episodes(url, excel_path)
