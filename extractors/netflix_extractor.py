import time
import pandas as pd
from selenium import webdriver
from bs4 import BeautifulSoup
import re
import urllib.parse
import logging
import os

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def clean_title(title):
    # 清理標題，去除 "第 x 集" 或 "第 x 話"
    title = re.sub(r'^第 \d+ 集。', '', title).strip()
    title = re.sub(r'^第\d+話', '', title).strip()
    return title

def extract_video_id(url):
    parsed_url = urllib.parse.urlparse(url)
    query_params = urllib.parse.parse_qs(parsed_url.query)
    path_segments = parsed_url.path.split('/')
    if 'jbv' in query_params:
        return query_params['jbv'][0]
    for segment in path_segments:
        if segment.isdigit():
            return segment
    return None

def get_series_info(driver, url):
    driver.get(url)
    time.sleep(1)
    html_content = driver.page_source

    soup = BeautifulSoup(html_content, 'html.parser')

    title = soup.find('h1', {'class': 'title-title'})
    if title:
        title = title.get_text(strip=True)
    else:
        title = ""

    description = soup.find('div', {'class': 'title-info-synopsis'})
    if description:
        description = description.get_text(strip=True)
    else:
        description = ""

    logging.info(f"劇集名稱: {title}, 劇集簡介: {description}")

    return title, description

def get_season_info(driver, url, series_description):
    driver.get(url)
    time.sleep(1)
    html_content = driver.page_source

    soup = BeautifulSoup(html_content, 'html.parser')
    season_data = []

    season_select = soup.find('select', {'data-uia': 'season-selector'})
    season_names = season_select.find_all('option') if season_select else []
    seasons = soup.find_all('div', class_='season')
    logging.info(f"找到 {len(seasons)} 個季數")

    if not seasons:
        # 如果沒有季數資料，用劇集簡介作為第一季的簡介
        season_data.append(('第 1 季', 1, series_description))
    else:
        for i, season in enumerate(seasons):
            if i < len(season_names):
                season_name = season_names[i].get_text(strip=True)
            else:
                season_name = f'第 {i+1} 季'
            season_number = i + 1
            season_description_element = season.find('p', class_='season-synopsis')
            if season_description_element:
                season_description = season_description_element.get_text(strip=True)
            else:
                season_description = series_description
            season_data.append((season_name, season_number, season_description))
            logging.info(f"季度 {season_name} - 描述: {season_description}")

    return season_data

def get_episode_info(driver, url):
    driver.get(url)
    time.sleep(1)
    html_content = driver.page_source

    soup = BeautifulSoup(html_content, 'html.parser')
    all_episode_data = []
    seasons = soup.find_all('div', class_='season')
    logging.info(f"找到 {len(seasons)} 個季數")

    season_number = 1
    for season in seasons:
        episode_elements = season.find_all('div', class_='episode-item') or \
                           season.find_all('div', class_='episode') or \
                           season.find_all('li', class_='episode')

        logging.info(f"找到 {len(episode_elements)} 個集數在第 {season_number} 季")

        episode_number = 1
        for episode_element in episode_elements:
            try:
                episode_title = clean_title(episode_element.find('h3').get_text(strip=True))
                episode_description = episode_element.find('p').get_text(strip=True)
                all_episode_data.append((season_number, episode_number, episode_title, episode_description))
                logging.info(f"第 {episode_number} 集 - 標題: {episode_title}, 描述: {episode_description}")
                episode_number += 1
            except Exception as e:
                logging.error(f"抓取集數資訊時發生錯誤: {e}")
                continue
        season_number += 1
    return all_episode_data

def save_to_excel(data, filename='video_detail.xlsx', sheet_name='Episodes'):
    with pd.ExcelWriter(filename, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df = pd.DataFrame(data, columns=['Season Number', 'Episode Number', 'Episode Title', 'Episode Description'])
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    logging.info(f"資料已保存到 {filename}")

def save_season_info_to_excel(data, filename='video_detail.xlsx', sheet_name='Seasons'):
    with pd.ExcelWriter(filename, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df = pd.DataFrame(data, columns=['Season Name', 'Season Number', 'Season Description'])
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    logging.info(f"季度信息已保存到 {filename}")

def save_series_info_to_excel(title, description, filename='video_detail.xlsx', sheet_name='Title'):
    with pd.ExcelWriter(filename, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df = pd.DataFrame([(title, description)], columns=['Title', 'Description'])
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    logging.info(f"劇集信息已保存到 {filename}")

def save_movie_info_to_excel(title, description, filename='video_detail.xlsx', sheet_name='Movies'):
    with pd.ExcelWriter(filename, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df = pd.DataFrame([(title, description)], columns=['Movie Title', 'Movie Description'])
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    logging.info(f"電影信息已保存到 {filename}")

def extract_netflix_episodes(driver, url, output_path):
    video_id = extract_video_id(url)
    if video_id:
        standardized_url = f"https://www.netflix.com/hk/title/{video_id}"
        logging.info(f"標準化URL: {standardized_url}")
        title, description = get_series_info(driver, standardized_url)

        driver.get(standardized_url)
        time.sleep(1)
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, 'html.parser')

        if soup.find('div', {'class': 'episode-metadata'}):
            # 是劇集
            save_series_info_to_excel(title, description, output_path)
            season_data = get_season_info(driver, standardized_url, description)
            save_season_info_to_excel(season_data, output_path)
            episode_data = get_episode_info(driver, standardized_url)
            if not episode_data:
                logging.error("無法從該網址取得資料")
            else:
                save_to_excel(episode_data, output_path)
        else:
            # 是電影
            save_movie_info_to_excel(title, description, output_path)
    else:
        logging.error("無法提取影片ID，請檢查輸入的URL")

